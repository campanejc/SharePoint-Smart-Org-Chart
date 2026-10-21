/* Unit tests for src/services/ExternalPeopleCsvService.ts against a fake SPHttpClient.
   Run: node tests/external-people-csv.cjs (no build or browser needed). */
const assert = require('assert/strict');
const fs = require('fs');
const path = require('path');
const ts = require('typescript');

const source = fs.readFileSync(path.resolve(__dirname, '../src/services/ExternalPeopleCsvService.ts'), 'utf8');
const js = ts.transpileModule(source, { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2019 } }).outputText;
const exported = {};
const stubRequire = id => {
  if (id === '@microsoft/sp-http') return { SPHttpClient: { configurations: { v1: 'v1' } } };
  throw new Error('unexpected import ' + id);
};
new Function('require', 'exports', js)(stubRequire, exported);
const { parseCsv, parseExternalPeopleCsv, decodeCsvBytes, toServerRelativePath, loadExternalPeopleFromCsv, MAX_CSV_ROWS } = exported;

const WEB = 'https://contoso.sharepoint.com/sites/hr/';
const FILE = 'https://contoso.sharepoint.com/sites/hr/Shared%20Documents/External%20People.csv';

const bytes = (text, encoding = 'utf8') => {
  const b = Buffer.from(text, encoding);
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
};
const response = (status, body) => ({ ok: status >= 200 && status < 300, status, arrayBuffer: async () => body });
const fakeClient = handler => { const calls = []; return { calls, get: async url => { calls.push(url); return handler(url); } }; };

(async () => {
  // ── Parsing ──
  assert.deepEqual(parseCsv('a,b\r\n1,2\n\n3,4\n'), [['a', 'b'], ['1', '2'], ['3', '4']]);
  assert.deepEqual(parseCsv('a,b\n"x, y","he said ""hi"""\n"multi\nline",z'),
    [['a', 'b'], ['x, y', 'he said "hi"'], ['multi\nline', 'z']]);
  assert.deepEqual(parseCsv('﻿a;b\n1;2'), [['a', 'b'], ['1', '2']], 'BOM stripped, semicolon detected');
  assert.deepEqual(parseCsv('a\tb\n1\t2'), [['a', 'b'], ['1', '2']], 'tab detected');
  assert.deepEqual(parseCsv('a,"b;c"\n1,2'), [['a', 'b;c'], ['1', '2']], 'delimiter inside quotes ignored when detecting');
  assert.deepEqual(parseCsv(''), []);
  assert.deepEqual(parseCsv('a,b'), [['a', 'b']], 'no trailing newline');
  console.log('PASS parse: quotes, newlines in quotes, BOM, delimiters, blank lines');

  // ── Mapping rows to people ──
  {
    const r = parseExternalPeopleCsv(
      'Name,Job Title,E-mail,Manager Email\n' +      // "E-mail" normalizes to "email"
      '  Ana ,Contractor, ana@p.com ,boss@x.com\n' +
      ',No name,,boss@x.com\n' +
      'Cara,,,\n');
    assert.deepEqual(r.missingColumns, []);
    assert.deepEqual(r.people, [
      { id: 'csv-1', name: 'Ana', jobTitle: 'Contractor', email: 'ana@p.com', managerEmail: 'boss@x.com' },
      { id: 'csv-3', name: 'Cara', jobTitle: '', email: '', managerEmail: '' },
    ]);
  }
  {
    const r = parseExternalPeopleCsv('title;reports to\nAna;boss@x.com');
    assert.deepEqual(r.people.map(p => [p.name, p.managerEmail]), [['Ana', 'boss@x.com']], 'aliases + semicolons');
    assert.deepEqual(r.missingColumns, []);
  }
  {
    const r = parseExternalPeopleCsv('Name,JobTitle\nAna,Contractor');
    assert.deepEqual(r.missingColumns, ['ManagerEmail']);
    assert.equal(r.people.length, 1, 'people still load; the pane reports the missing column');
  }
  {
    const r = parseExternalPeopleCsv('JobTitle,ManagerEmail\nx,y');
    assert.deepEqual(r.missingColumns, ['Name']);
    assert.deepEqual(r.people, []);
    assert.deepEqual(parseExternalPeopleCsv('').missingColumns, ['Name', 'ManagerEmail']);
  }
  {
    let big = 'Name,ManagerEmail\n';
    for (let i = 0; i < MAX_CSV_ROWS + 50; i++) big += `P${i},boss@x.com\n`;
    assert.equal(parseExternalPeopleCsv(big).people.length, MAX_CSV_ROWS, 'row cap');
  }
  console.log('PASS mapping: trimming, skipped rows, aliases, missing columns, row cap');

  // ── Decoding ──
  assert.equal(decodeCsvBytes(bytes('﻿Name\nJosé')), 'Name\nJosé', 'UTF-8 with BOM');
  assert.equal(decodeCsvBytes(bytes('Name\nJos\xe9', 'latin1')), 'Name\nJosé', 'legacy Windows-1252 fallback');
  assert.equal(decodeCsvBytes(bytes('﻿Name\nJosé', 'utf16le')), 'Name\nJosé', 'UTF-16 (BOM stripped)');
  console.log('PASS decoding: UTF-8, Windows-1252, UTF-16');

  // ── URLs ──
  assert.equal(toServerRelativePath(FILE, WEB), '/sites/hr/Shared Documents/External People.csv');
  assert.equal(toServerRelativePath('/sites/hr/Shared Documents/a.csv?web=1#x', WEB), '/sites/hr/Shared Documents/a.csv');
  assert.equal(toServerRelativePath('/sites/hr/Docs/A.CSV', WEB), '/sites/hr/Docs/A.CSV');
  for (const bad of ['', '   ', 'https://evil.example/sites/hr/a.csv', 'https://contoso.sharepoint.com/sites/hr/a.xlsx',
                     'relative/a.csv', '/sites/hr/../other/a.csv', 'javascript:alert(1)', 'ftp://contoso.sharepoint.com/a.csv']) {
    assert.equal(toServerRelativePath(bad, WEB), undefined, bad);
  }
  console.log('PASS urls: full and relative forms accepted; other hosts, non-csv, traversal rejected');

  // ── Loading ──
  {
    const client = fakeClient(() => response(200, bytes('Name,ManagerEmail\nAna,boss@x.com')));
    const r = await loadExternalPeopleFromCsv(client, WEB, FILE);
    assert.equal(r.error, undefined);
    assert.deepEqual(r.people.map(p => p.name), ['Ana']);
    assert.equal(client.calls.length, 1);
    assert.equal(client.calls[0],
      "https://contoso.sharepoint.com/sites/hr/_api/web/GetFileByServerRelativePath(decodedurl='/sites/hr/Shared%20Documents/External%20People.csv')/$value");
  }
  {
    const client = fakeClient(() => response(200, bytes('Name,ManagerEmail\nAna,b@x.com')));
    await loadExternalPeopleFromCsv(client, WEB, "/sites/hr/Docs/Bob's%20%231%20list.csv");
    assert(client.calls[0].includes("Bob''s%20%231%20list.csv"), 'quote doubled, # encoded: ' + client.calls[0]);
  }
  for (const [status, code] of [[404, 'notFound'], [403, 'accessDenied'], [401, 'accessDenied'], [500, 'failed']]) {
    const r = await loadExternalPeopleFromCsv(fakeClient(() => response(status, new ArrayBuffer(0))), WEB, FILE);
    assert.equal(r.error, code);
    assert.deepEqual(r.people, []);
  }
  assert.equal((await loadExternalPeopleFromCsv({ get: async () => { throw new Error('network'); } }, WEB, FILE)).error, 'failed');
  {
    const client = fakeClient(() => response(200, new ArrayBuffer(0)));
    const r = await loadExternalPeopleFromCsv(client, WEB, 'https://other.example/a.csv');
    assert.equal(r.error, 'invalidUrl');
    assert.equal(client.calls.length, 0, 'an invalid URL never reaches a request');
  }
  // A file in another site collection: the current web says 404, so the web that owns the file is looked up and used
  {
    const OTHER = 'https://contoso.sharepoint.com/sites/TestSite';
    const OTHER_FILE = 'https://contoso.sharepoint.com/sites/TestSite/Documents/externalusers.csv';
    const calls = [];
    const client = {
      get: async url => {
        calls.push('GET ' + url);
        return url.startsWith(OTHER) ? response(200, bytes('Name,ManagerEmail\nAna,boss@x.com')) : response(404, new ArrayBuffer(0));
      },
      post: async url => {
        calls.push('POST ' + url);
        return { ok: true, status: 200, json: async () => ({ value: OTHER }) };
      },
    };
    const r = await loadExternalPeopleFromCsv(client, WEB, OTHER_FILE);
    assert.equal(r.error, undefined);
    assert.deepEqual(r.people.map(p => p.name), ['Ana']);
    assert(calls[1].startsWith('POST https://contoso.sharepoint.com/sites/hr/_api/SP.Web.GetWebUrlFromPageUrl(@v)?@v='), calls[1]);
    assert(calls[1].includes(encodeURIComponent(OTHER_FILE)), 'the full file URL is sent');
    assert(calls[2].startsWith('GET ' + OTHER + '/_api/web/GetFileByServerRelativePath'), calls[2]);

    // the owner lookup failing (or naming the same web) just leaves the original not-found error
    const failing = { get: async () => response(404, new ArrayBuffer(0)), post: async () => ({ ok: false, status: 500, json: async () => ({}) }) };
    assert.equal((await loadExternalPeopleFromCsv(failing, WEB, OTHER_FILE)).error, 'notFound');
    const same = { get: async () => response(404, new ArrayBuffer(0)), post: async () => ({ ok: true, status: 200, json: async () => ({ value: WEB.replace(/\/$/, '') }) }) };
    assert.equal((await loadExternalPeopleFromCsv(same, WEB, FILE)).error, 'notFound');
  }
  console.log('PASS load: request URL, encoding, error codes, invalid URL never requested');
})().catch(e => { console.error(e); process.exitCode = 1; });
