/* Run after building demo/webpack.demo.config.js: node tests/external-people.cjs
   Drives the demo harness (?pane=1&csv=1): inline people come from the web part's real
   property pane editor, CSV people from the demo stand-in for the people CSV file. */
const assert = require('assert/strict');
const fs = require('fs');
const path = require('path');
const http = require('http');
const puppeteer = require('puppeteer-core');

const root = path.resolve(__dirname, '../demo/dist');
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

const level = (page, id) => page.$eval(`[data-soc-id="${id}"]`, el => Number(el.getAttribute('aria-level')));
const exists = (page, id) => page.$(`[data-soc-id="${id}"]`).then(el => !!el);
const paneText = page => page.$eval('aside', el => el.innerText);

async function waitForPane(page, text) {
  await page.waitForFunction(
    t => !!document.querySelector('aside') && document.querySelector('aside').innerText.includes(t), { timeout: 10000 }, text);
}

// Adds an inline person through the pane editor and returns its row id
async function addPerson(page, { name, email, manager }) {
  const before = await page.$$eval('aside input[id^="soc-external-name-"]', els => els.length);
  await page.click('aside button::-p-text(+ Add person)');
  await page.waitForFunction(n => document.querySelectorAll('aside input[id^="soc-external-name-"]').length === n + 1, {}, before);
  const rowId = await page.$$eval('aside input[id^="soc-external-name-"]',
    els => els[els.length - 1].id.replace(/^soc-external-name-/, '').replace(/-demo$/, ''));
  const fill = async (key, value) => {
    if (!value) return;
    await page.click(`#soc-external-${key}-${rowId}-demo`);
    await page.keyboard.type(value);
    await page.keyboard.press('Tab');   // blur commits
  };
  await fill('name', name);
  await fill('email', email);
  await fill('managerEmail', manager);
  return rowId;
}

(async () => {
  const server = http.createServer((req, res) => {
    const file = path.join(root, decodeURIComponent(req.url.split('?')[0] === '/' ? '/index.html' : req.url.split('?')[0]));
    if (!file.startsWith(root + path.sep) || !fs.existsSync(file)) { res.writeHead(404).end(); return; }
    res.setHeader('Content-Type', file.endsWith('.js') ? 'application/javascript' : file.endsWith('.html') ? 'text/html' : 'application/octet-stream');
    fs.createReadStream(file).pipe(res);
  });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const browser = await puppeteer.launch({ headless: true,
    executablePath: process.env.CHROME_PATH || 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe' });
  const errors = [];
  try {
    const page = await browser.newPage();
    page.on('pageerror', e => errors.push(e.message));
    await page.setViewport({ width: 1400, height: 950 });
    const base = `http://127.0.0.1:${server.address().port}/?pane=1&csv=1&mockSize=150`;

    // ── Org chart: inline people, list people, external manager chain ──
    await page.goto(`${base}&view=orgchart&layout=vertical&levels=5`, { waitUntil: 'networkidle0' });
    await page.waitForSelector('[data-soc-id="external:csv-2"]');
    assert.equal(await level(page, 'external:demo-ext-1'), 2, 'inline person reports to the CEO');
    assert.equal(await level(page, 'external:csv-1'), 2, 'CSV person reports to the CEO');
    assert.equal(await level(page, 'external:csv-2'), 3, 'CSV person can report to another external (by email)');
    assert.equal(await exists(page, 'external:csv-3'), false, 'person with an unknown manager is not placed in the chart');
    const badges = await page.$$eval('[data-soc-id^="external:"]', els => els.filter(e => e.innerText.includes('External')).length);
    assert(badges >= 3, 'external cards show the External badge');
    console.log('PASS chart: inline and CSV people placed under the right managers, external chains, unknown manager');

    // ── Property pane: list status and unresolved warning ──
    await waitForPane(page, '3 people loaded from the file.');
    assert((await paneText(page)).includes('Open Position'), 'pane names the CSV person who could not be placed');
    console.log('PASS pane: CSV count and unresolved CSV people reported');

    // ── Property pane: saved people are collapsed to their name ──
    const rowState = () => page.$$eval('aside button[aria-controls^="soc-external-body-"]', els => els.map(b => ({
      text: b.innerText.trim(), expanded: b.getAttribute('aria-expanded'),
      bodyShown: getComputedStyle(document.getElementById(b.getAttribute('aria-controls'))).display !== 'none' })));
    const collapsed = await rowState();
    assert.deepEqual(collapsed.map(r => r.text.replace(/^[^\p{L}]+/u, '')).sort(), ['Marco Bellini', 'Open Position', 'Priya Natarajan']);
    assert(collapsed.every(r => r.expanded === 'false' && !r.bodyShown), 'saved people start collapsed');
    assert((await paneText(page)).includes('People added: 3'));
    await page.click('aside button[aria-controls^="soc-external-body-"]');
    const afterToggle = await rowState();
    assert(afterToggle.some(r => r.expanded === 'true' && r.bodyShown), 'toggling opens the fields');
    await page.click('aside button[aria-expanded="true"]');
    assert((await rowState()).every(r => !r.bodyShown), 'and closes them again');
    console.log('PASS pane: saved people collapse to their name; toggle opens and closes the fields');

    // ── Directory: everyone is listed, including the person not placed in the chart ──
    await page.goto(`${base}&view=directory`, { waitUntil: 'networkidle0' });
    await page.waitForSelector('input[aria-label="Search people"]');
    await page.type('input[aria-label="Search people"]', 'Open Position');
    await page.waitForFunction(() => document.body.innerText.includes('Data Analyst (vacancy)'), { timeout: 10000 });
    assert((await page.evaluate(() => document.body.innerText)).includes('External'));
    console.log('PASS directory: unplaced external person is still listed with the badge');

    // ── Pane editor: email as reference key, validation and warnings ──
    await page.goto(`${base}&view=orgchart&layout=vertical&levels=6`, { waitUntil: 'networkidle0' });
    await page.waitForSelector('aside input[id^="soc-external-name-"]');
    await addPerson(page, { name: 'Pat Lead', email: 'pat.lead@partner.example', manager: 'a.chen@contoso.com' });
    await addPerson(page, { name: 'Report One', manager: 'pat.lead@partner.example' });
    const orphan = await addPerson(page, { name: 'Lost Person', manager: 'nobody@nowhere.example' });
    await page.waitForFunction(() => !![...document.querySelectorAll('[data-soc-id^="external:ext-"]')].length, { timeout: 10000 });
    await delay(400);
    const ids = await page.$$eval('[data-soc-id^="external:ext-"]', els => els.map(e => ({ id: e.getAttribute('data-soc-id'), text: e.innerText, level: Number(e.getAttribute('aria-level')) })));
    const pat = ids.find(x => x.text.includes('Pat Lead'));
    const one = ids.find(x => x.text.includes('Report One'));
    assert(pat && one, 'new people appear in the chart');
    assert.equal(pat.level, 2);
    assert.equal(one.level, 3, 'report to an inline external by email');
    assert(!ids.some(x => x.text.includes('Lost Person')), 'unknown manager: not placed in the chart');
    await page.waitForFunction(id => {
      const w = document.querySelector(`#soc-external-managerWarning-${id}-demo`);
      return !!w && w.style.display !== 'none' && w.textContent.length > 0;
    }, { timeout: 10000 }, orphan);
    // collapsing the row leaves a ⚠ in its header
    await page.click(`aside button[aria-controls="soc-external-body-${orphan}-demo"]`);
    const flag = await page.$eval(`aside button[aria-controls="soc-external-body-${orphan}-demo"]`, b => {
      const f = b.querySelector('span:last-child');
      return { shown: getComputedStyle(f).display !== 'none', title: f.title };
    });
    assert(flag.shown && flag.title.length > 0, 'collapsed row flags a manager that matches no one');
    await page.click(`aside button[aria-controls="soc-external-body-${orphan}-demo"]`);   // reopen to keep editing
    console.log('PASS editor: external manager by email, unknown-manager warning, flag when collapsed');

    // invalid email is flagged in the pane and does not place anyone
    await page.click(`#soc-external-email-${orphan}-demo`);
    await page.keyboard.type('not-an-email');
    await page.keyboard.press('Tab');
    assert.equal(await page.$eval(`#soc-external-email-${orphan}-demo`, el => el.getAttribute('aria-invalid')), 'true');
    console.log('PASS editor: invalid email flagged');

    // Page edit mode: with a profile card open, the property pane (which sits above the web part)
    // must stay usable — focus may move into it, and clicking it dismisses the card
    await page.goto(`${base}&view=orgchart&layout=vertical&levels=5`, { waitUntil: 'networkidle0' });
    await page.waitForSelector('[data-soc-id="external:demo-ext-1"]');
    await page.$eval('aside', el => { el.style.zIndex = '100000'; });
    await page.click('[data-soc-id="external:demo-ext-1"]');
    await page.waitForSelector('[role="dialog"]');
    await page.focus('#soc-external-csv-demo');
    await delay(200);
    assert.equal(await page.evaluate(() => document.activeElement.id), 'soc-external-csv-demo', 'the card does not steal focus back');
    assert(await page.$('[role="dialog"]'), 'moving focus alone leaves the card open');
    const before = await page.$eval('#soc-external-csv-demo', el => el.value);
    await page.click('#soc-external-csv-demo');
    await page.keyboard.type('x');
    assert.equal((await page.$eval('#soc-external-csv-demo', el => el.value)).length, before.length + 1, 'the pane field takes input');
    await page.waitForSelector('[role="dialog"]', { hidden: true });
    console.log('PASS edit mode: property pane stays usable while a profile card is open, and clicking it closes the card');

    assert.deepEqual(errors, []);
    console.log('PASS no runtime errors');
  } finally {
    await browser.close();
    server.close();
  }
})().catch(e => { console.error(e); process.exitCode = 1; });
