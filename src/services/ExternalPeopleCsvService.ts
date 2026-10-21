// Reads admin-maintained "external people" (contractors, partners, vacancies) from a
// CSV file saved in a SharePoint document library on the current site collection.
// Uses the page's SPHttpClient, so it runs with the viewer's own permissions and
// needs no extra API access.
import { SPHttpClient } from '@microsoft/sp-http';
import { IExternalPeopleStatus, IExternalPersonConfig } from './GraphService';

// Prefix for ids of people that came from the CSV file (inline people use "ext-")
export const CSV_PERSON_ID_PREFIX = 'csv-';

// A CSV this large is far beyond any org chart; the cap keeps a wrong file from freezing the page
export const MAX_CSV_ROWS = 5000;

export type ExternalCsvErrorCode = 'invalidUrl' | 'notFound' | 'accessDenied' | 'failed';

export interface IExternalCsvResult {
  people: IExternalPersonConfig[];
  missingColumns: string[];                         // required columns the file does not have
  error?: ExternalCsvErrorCode;
  detail?: string;                                  // technical reason (HTTP status, ...) shown in the pane to help diagnose
}

// Everything the property pane shows about external people after the chart has resolved them
export interface IExternalPeopleReport {
  csvError?: ExternalCsvErrorCode;
  csvDetail?: string;
  csvCount: number;                                 // people read from the CSV file
  missingColumns: string[];
  unresolved: IExternalPeopleStatus['unresolved'];
}

// Header names (lower-cased, with spaces, dashes and underscores removed) → our fields
const HEADER_ALIASES: { [field: string]: string[] } = {
  name: ['name', 'title', 'displayname', 'fullname'],
  jobTitle: ['jobtitle', 'position', 'role'],
  email: ['email', 'mail', 'emailaddress'],
  managerEmail: ['manageremail', 'manager', 'reportsto', 'reportstoemail', 'manageremailaddress'],
};
const REQUIRED_COLUMNS: Array<{ field: string; label: string }> = [
  { field: 'name', label: 'Name' },
  { field: 'managerEmail', label: 'ManagerEmail' },
];

const normalizeHeader = (h: string): string => h.toLowerCase().replace(/[\s_-]+/g, '');

// The delimiter is whichever of , ; or tab appears most in the header line
// (outside quotes) — Excel in many locales writes semicolons.
function detectDelimiter(text: string): string {
  const counts: { [d: string]: number } = { ',': 0, ';': 0, '\t': 0 };
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') inQuotes = !inQuotes;
    else if (!inQuotes && (ch === '\n' || ch === '\r')) break;
    else if (!inQuotes && counts[ch] !== undefined) counts[ch]++;
  }
  return counts[';'] > counts[','] && counts[';'] >= counts['\t'] ? ';'
    : counts['\t'] > counts[','] ? '\t' : ',';
}

// RFC 4180: quoted fields, "" escapes a quote, newlines allowed inside quotes.
// Blank lines are skipped.
export function parseCsv(input: string): string[][] {
  const text = input.charCodeAt(0) === 0xFEFF ? input.substring(1) : input;
  const delimiter = detectDelimiter(text);
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  let wasQuoted = false;

  const endField = (): void => { row.push(field); field = ''; wasQuoted = false; };
  const endRow = (): void => {
    const quoted = wasQuoted;
    endField();
    if (row.length > 1 || row[0].trim() !== '' || quoted) rows.push(row);
    row = [];
  };

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; } else inQuotes = false;
      } else {
        field += ch;
      }
    } else if (ch === '"' && field === '') {
      inQuotes = true;
      wasQuoted = true;
    } else if (ch === delimiter) {
      endField();
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      endRow();
    } else {
      field += ch;
    }
  }
  if (field !== '' || row.length > 0 || wasQuoted) endRow();
  return rows;
}

// Turns CSV text into people. Missing optional columns are fine; if the Name
// column is missing nobody can be read.
export function parseExternalPeopleCsv(text: string): Pick<IExternalCsvResult, 'people' | 'missingColumns'> {
  const rows = parseCsv(text);
  const result: Pick<IExternalCsvResult, 'people' | 'missingColumns'> = { people: [], missingColumns: [] };
  if (rows.length === 0) {
    result.missingColumns = REQUIRED_COLUMNS.map(c => c.label);
    return result;
  }
  const headers = rows[0].map(normalizeHeader);
  const index: { [field: string]: number } = {};
  Object.keys(HEADER_ALIASES).forEach(field => {
    index[field] = headers.findIndex(h => HEADER_ALIASES[field].indexOf(h) >= 0);
  });
  result.missingColumns = REQUIRED_COLUMNS.filter(c => index[c.field] < 0).map(c => c.label);
  if (index.name < 0) return result;

  const cell = (row: string[], field: string): string => (index[field] >= 0 ? (row[index[field]] || '').trim() : '');
  for (let r = 1; r < rows.length && r <= MAX_CSV_ROWS; r++) {
    const name = cell(rows[r], 'name');
    if (!name) continue;                            // a row without a name is not a person
    result.people.push({
      id: CSV_PERSON_ID_PREFIX + r,
      name,
      jobTitle: cell(rows[r], 'jobTitle'),
      email: cell(rows[r], 'email'),
      managerEmail: cell(rows[r], 'managerEmail'),
    });
  }
  return result;
}

// UTF-8 first (what "CSV UTF-8" writes); legacy Excel "CSV" is Windows-1252, and
// "Unicode Text" is UTF-16 — both are handled so accented names don't turn to garbage.
export function decodeCsvBytes(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  if (bytes.length >= 2 && bytes[0] === 0xFF && bytes[1] === 0xFE) return new TextDecoder('utf-16le').decode(bytes);
  if (bytes.length >= 2 && bytes[0] === 0xFE && bytes[1] === 0xFF) return new TextDecoder('utf-16be').decode(bytes);
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    return new TextDecoder('windows-1252').decode(bytes);
  }
}

// Accepts a full URL on this tenant or a server-relative path; returns the
// decoded server-relative path of a .csv file, or undefined if it isn't one.
export function toServerRelativePath(fileUrl: string, webUrl: string): string | undefined {
  const input = (fileUrl || '').trim();
  if (!input) return undefined;
  let path: string;
  if (/^https?:\/\//i.test(input)) {
    let url: URL;
    let web: URL;
    try { url = new URL(input); web = new URL(webUrl); } catch { return undefined; }
    if (url.origin.toLowerCase() !== web.origin.toLowerCase()) return undefined;   // other tenants/hosts are never read
    path = url.pathname;
  } else if (input.charAt(0) === '/') {
    path = input.split(/[?#]/)[0];
  } else {
    return undefined;
  }
  try { path = decodeURIComponent(path); } catch { /* keep as typed */ }
  if (!/\.csv$/i.test(path) || path.split('/').indexOf('..') >= 0) return undefined;
  return path;
}

function errorCode(status: number): ExternalCsvErrorCode {
  if (status === 404) return 'notFound';
  if (status === 401 || status === 403) return 'accessDenied';
  return 'failed';
}

// Reads the file through one web's REST API. `status` is the HTTP status when it failed.
async function readFromWeb(
  client: SPHttpClient,
  webUrl: string,
  path: string
): Promise<IExternalCsvResult & { status?: number }> {
  const quoted = encodeURIComponent(path.replace(/'/g, "''")).replace(/%2F/gi, '/');
  const url = `${webUrl.replace(/\/$/, '')}/_api/web/GetFileByServerRelativePath(decodedurl='${quoted}')/$value`;
  const resp = await client.get(url, SPHttpClient.configurations.v1, { headers: { 'Accept': '*/*' } });
  if (!resp.ok) return { people: [], missingColumns: [], error: errorCode(resp.status), status: resp.status, detail: `HTTP ${resp.status}` };
  return parseExternalPeopleCsv(decodeCsvBytes(await resp.arrayBuffer()));
}

// A web can only serve files in its own site collection, so for a file somewhere
// else on the tenant ask SharePoint which web the file's URL belongs to.
async function findWebUrlOfFile(client: SPHttpClient, webUrl: string, path: string): Promise<string | undefined> {
  try {
    const origin = new URL(webUrl).origin;
    const fullUrl = origin + path.split('/').map(encodeURIComponent).join('/');
    const url = `${webUrl.replace(/\/$/, '')}/_api/SP.Web.GetWebUrlFromPageUrl(@v)?@v='${encodeURIComponent(fullUrl.replace(/'/g, "''"))}'`;
    const resp = await client.post(url, SPHttpClient.configurations.v1, { headers: { 'Accept': 'application/json;odata=nometadata' } });
    if (!resp.ok) return undefined;
    const json = await resp.json();
    const found = json && (json.value || (json.d && json.d.GetWebUrlFromPageUrl));
    return typeof found === 'string' && found ? found : undefined;
  } catch {
    return undefined;
  }
}

// Loads and parses the file. Never throws — problems come back as `error` so
// the chart keeps working with whatever else is configured.
export async function loadExternalPeopleFromCsv(
  client: SPHttpClient,
  webUrl: string,
  fileUrl: string
): Promise<IExternalCsvResult> {
  const empty: IExternalCsvResult = { people: [], missingColumns: [] };
  const path = toServerRelativePath(fileUrl, webUrl);
  if (!path) {
    console.warn('[Smart Org Chart] People CSV: the link is not a .csv file on this tenant:', fileUrl);
    return { ...empty, error: 'invalidUrl', detail: 'not a .csv link on this tenant' };
  }
  try {
    let result = await readFromWeb(client, webUrl, path);
    if (result.error === 'notFound') {
      // Probably in another site collection: read it through the web that owns it
      const ownerWeb = await findWebUrlOfFile(client, webUrl, path);
      if (ownerWeb && ownerWeb.replace(/\/$/, '').toLowerCase() !== webUrl.replace(/\/$/, '').toLowerCase()) {
        result = await readFromWeb(client, ownerWeb, path);
      }
    }
    if (result.error) console.warn(`[Smart Org Chart] People CSV: could not read ${path} (HTTP ${result.status}, ${result.error})`);
    return { people: result.people, missingColumns: result.missingColumns, error: result.error, detail: result.detail };
  } catch (err) {
    console.warn('[Smart Org Chart] People CSV: the request failed', err);
    return { ...empty, error: 'failed', detail: err instanceof Error ? err.message : 'request failed' };
  }
}
