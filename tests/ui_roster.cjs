// DOM integration against the real API, using synthetic rows only.
// Usage: NODE_PATH=/path/to/jsdom/node_modules node tests/ui_roster.cjs
const assert = require('node:assert/strict');
const fs = require('node:fs');
const { JSDOM, VirtualConsole } = require('jsdom');
const base = process.env.FORMAFLOW_TEST_URL || 'http://127.0.0.1:8765';
const failures = [];
const virtualConsole = new VirtualConsole();
virtualConsole.on('jsdomError', error => failures.push(error));
const dom = new JSDOM(fs.readFileSync('static/index.html', 'utf8'), { url: base, runScripts: 'outside-only', virtualConsole });
const w = dom.window;
w.FormData = FormData;
w.confirm = () => true;
let sessionId;
let latestPreview;
let failNextPreview = false;
let delayedPreview = null;
w.fetch = async (url, options) => {
  if (url === '/api/preview' && failNextPreview) { failNextPreview = false; throw Error('Simulated connection failure'); }
  const response = await fetch(new URL(url, base), options);
  if (url === '/api/upload') sessionId = (await response.clone().json()).session_id;
  if (url === '/api/preview') {
    const data = await response.clone().json();
    if (delayedPreview) { const pause = delayedPreview; delayedPreview = null; await pause; }
    latestPreview = data;
  }
  return response;
};
require('node:vm').runInContext(fs.readFileSync('static/app.js', 'utf8'), dom.getInternalVMContext());
const $ = id => w.document.getElementById(id);
const fire = (element, event) => element.dispatchEvent(new w.Event(event, { bubbles: true }));
const waitFor = async predicate => {
  const deadline = Date.now() + 10000;
  while (!predicate()) {
    if (Date.now() > deadline) throw Error('UI timed out: ' + $('previewNote').textContent + ' / ' + $('uploadMessage').textContent);
    await new Promise(resolve => setTimeout(resolve, 20));
  }
};
const previewReady = () => w.eval('state.previewReady');
const input = (element, value) => { element.value = value; fire(element, 'input'); };
const names = () => [...$('previewBody').rows].map(row => row.cells[2]?.textContent);

(async () => {
 try {
  $('languageSelect').value = 'uz'; fire($('languageSelect'), 'change');
  const rows = ['№,Student name,Class'];
  for (let i = 0; i < 135; i++) rows.push(`${i + 1},Example Learner ${String(i + 1).padStart(3, '0')},4-A`);
  Object.defineProperty($('fileInput'), 'files', { configurable: true, value: [new File([rows.join('\n')], 'synthetic.csv', { type: 'text/csv' })] });
  fire($('uploadForm'), 'submit');
  await waitFor(() => previewReady() && $('previewBody').rows.length === 100);
  assert.match($('previewHead').textContent, /O‘quvchi F.I.Sh./);
  assert.equal($('addStudent').textContent, '+ O‘quvchi qo‘shish');
  assert.equal($('manualHelp').textContent, 'Yangi qatorlar jadvalda va yuklangan faylda ko‘rinadi.');
  assert(!w.document.getElementById('manualEmpty'));
  $('addStudent').click();
  assert.equal($('manualBody').querySelector('[data-key="row_number"]'), null);
  input($('manualBody').querySelector('[data-key="student_name"]'), 'Added Learner');
  await waitFor(() => previewReady() && names().includes('Added Learner'));
  assert.equal($('previewPage').value, '2');
  assert.equal($('previewBody').rows.length, 36);
  assert($('previewBody').lastElementChild.classList.contains('added-row'));
  input($('manualBody').querySelector('[data-key="class"]'), '7-A');
  await waitFor(() => previewReady());
  assert.match($('previewBody').lastElementChild.textContent, /8-A/);

  // Position field moves a row across pages and keeps it visible.
  const position = $('previewBody').lastElementChild.querySelector('.row-position');
  position.value = '1'; fire(position, 'change');
  await waitFor(() => previewReady() && names()[0] === 'Added Learner');
  assert.equal($('previewPage').value, '1');
  assert.equal($('previewBody').rows[0].cells[1].textContent, '1');

  // Exercise the actual drag/drop listeners with a browser-shaped DataTransfer.
  const dataTransfer = { effectAllowed: '', dropEffect: '', setData() {} };
  for (const [type, element] of [
    ['dragstart', $('previewBody').rows[0].querySelector('.row-handle')],
    ['dragover', $('previewBody').rows[2].cells[2]],
    ['drop', $('previewBody').rows[2].cells[2]],
  ]) {
    const event = new w.Event(type, { bubbles: true, cancelable: true });
    Object.defineProperty(event, 'dataTransfer', { value: dataTransfer });
    element.dispatchEvent(event);
  }
  await waitFor(() => previewReady() && names()[2] === 'Added Learner');
  assert.deepEqual([...$('previewBody').rows].slice(0, 3).map(row => row.cells[1].textContent), ['1', '2', '3']);
  $('previewBody').rows[2].querySelector('[data-direction="-1"]').click();
  await waitFor(() => previewReady() && names()[1] === 'Added Learner');

  // Export current page configuration: full result, same order and numbering.
  const exported = await fetch(base + '/api/export', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ ...w.eval('buildPayload()'), format: 'csv' }) });
  const csv = await exported.text();
  assert.equal(exported.status, 200);
  assert.match(csv.split('\n')[2], /^2,Added Learner,/);
  assert.equal(csv.trim().split('\n').length, 137);

  // Explicit sort replaces manual order; editing still focuses the added student.
  $('sortKey').value = 'student_name'; fire($('sortKey'), 'change');
  await waitFor(() => previewReady() && names()[0] === 'Added Learner');
  assert.equal(w.eval('buildPayload().row_order.length'), 0);
  $('lastPage').click();
  await waitFor(() => previewReady() && $('previewPage').value === '2');
  input($('manualBody').querySelector('[data-key="student_name"]'), 'Changed Learner');
  await waitFor(() => previewReady() && names()[0] === 'Changed Learner');

  // Late responses must not replace a newer edit.
  let release;
  delayedPreview = new Promise(resolve => { release = resolve; });
  input($('manualBody').querySelector('[data-key="student_name"]'), 'Old Learner');
  await waitFor(() => delayedPreview === null);
  input($('manualBody').querySelector('[data-key="student_name"]'), 'Newest Learner');
  await waitFor(() => previewReady() && names().includes('Newest Learner'));
  release();
  await new Promise(resolve => setTimeout(resolve, 60));
  assert(names().includes('Newest Learner'));
  assert(!names().includes('Old Learner'));

  // A failed preview cannot leave an apparently current stale table.
  failNextPreview = true;
  input($('manualBody').querySelector('[data-key="class"]'), '6-A');
  await waitFor(() => $('previewNote').textContent.includes('Simulated'));
  assert.equal($('previewBody').rows.length, 0);
  input($('manualBody').querySelector('[data-key="class"]'), '7-A');
  await waitFor(previewReady);

  // Rename labels update the manual editor too, without erasing draft values.
  const card = [...w.document.querySelectorAll('.column-card')].find(c => c.querySelector('.column-check').dataset.key === 'student_name');
  input(card.querySelector('.column-name'), 'Ism / familiya');
  await waitFor(previewReady);
  assert.match($('manualHead').textContent, /Ism \/ familiya/);
  assert.equal($('manualBody').querySelector('[data-key="student_name"]').value, 'Newest Learner');

  $('languageSelect').value = 'ru'; fire($('languageSelect'), 'change');
  await waitFor(previewReady);
  assert.equal($('addStudent').textContent, '+ Добавить ученика');
  assert.equal($('sortKey').value, 'student_name');
  assert.match($('manualHead').textContent, /Ism \/ familiya/);
  assert.equal($('manualBody').querySelector('[data-key="student_name"]').value, 'Newest Learner');

  $('manualBody').querySelector('.remove').click();
  await waitFor(() => previewReady() && latestPreview.row_count === 135);
  assert(!latestPreview.row_order.some(id => id.startsWith('manual:')));
  assert.equal(failures.length, 0, failures.map(e => e.stack).join('\n'));
  console.log('PASS: DOM upload, Uzbek labels, additions past 100, drag/drop, arrows, cross-page movement, renumbering, export parity, sorting, edit focus, stale/error handling, label editing, removal.');
 } finally {
  if (sessionId) await fetch(base + '/api/session/' + sessionId, { method: 'DELETE' });
  w.close();
 }
})().catch(error => { console.error(error); process.exitCode = 1; });
