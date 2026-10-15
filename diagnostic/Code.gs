const SHEET_NAME = 'Results';

// Optional read protection. Leave '' to disable. If you set a key, the dashboard must request
// ".../exec?action=getAll&key=YOURKEY" — edit fetchData() in dashboard.html to append it. Submissions
// (POST) never need the key.
const READ_KEY = '';

// Columns that are always present, in this order. Any other field a page sends (unit_*, bi_*, sp_*,
// diff_*, ...) is added automatically as a new column the first time it appears.
const CORE_HEADERS = [
  'submitTime', 'studentName', 'studentId', 'period', 'diagnosticType', 'startTime', 'timeTaken',
  'totalScore', 'totalPossible', 'pctScore', 'wrongCount', 'skippedCount',
  'wrongQs', 'skippedQs', 'rawAnswers', 'topicScores', 'spScores'
];

// Columns stored as plain text so Sheets never reformats them (dates, "1,5,9" lists, JSON, IDs).
const TEXT_COLUMNS = ['submitTime', 'startTime', 'studentId', 'wrongQs', 'skippedQs', 'rawAnswers', 'topicScores', 'spScores'];

// Columns the dashboard expects as parsed objects rather than strings.
const JSON_COLUMNS = ['topicScores', 'spScores'];

/* ───────────────────────── Web app entry points ───────────────────────── */

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);                       // several students can submit at once
    const raw = e && e.postData && e.postData.contents;
    if (!raw) return json_({ ok: false, error: 'Empty request body' });
    const payload = JSON.parse(raw);
    const rec = prepareRecord_(payload);
    appendRecord_(rec);
    return json_({ ok: true, row: rec.studentName || '' });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = p.action || 'ping';
  if (action === 'ping') {
    return json_({ ok: true, sheet: SHEET_NAME, rows: Math.max(0, getSheet_().getLastRow() - 1) });
  }
  if (action === 'getAll') {
    if (READ_KEY && p.key !== READ_KEY) return json_({ ok: false, error: 'Unauthorized' });
    return json_({ ok: true, data: readAll_() });
  }
  return json_({ ok: false, error: 'Unknown action: ' + action });
}

/* ───────────────────────── Record preparation ───────────────────────── */

function prepareRecord_(p) {
  const rec = {};
  Object.keys(p).forEach(function (k) {
    const v = p[k];
    rec[k] = (v !== null && typeof v === 'object') ? JSON.stringify(v) : v;   // objects → JSON text
  });

  // The full diagnostic does not send a type; unit diagnostics send "Unit N".
  rec.diagnosticType = rec.diagnosticType || 'Full';

  // Counts derived from the comma-separated question-number lists.
  const countList = function (s) { return s ? String(s).split(',').filter(function (x) { return x.trim() !== ''; }).length : 0; };
  rec.wrongCount = countList(rec.wrongQs);
  rec.skippedCount = countList(rec.skippedQs);

  // For every *_correct / *_total pair (unit_*, bi_*, sp_*, diff_*), add a *_pct column.
  Object.keys(rec).forEach(function (k) {
    const m = k.match(/^(.*)_correct$/);
    if (!m) return;
    const total = Number(rec[m[1] + '_total']);
    const correct = Number(rec[k]) || 0;
    if (total > 0) rec[m[1] + '_pct'] = Math.round(correct / total * 100);
  });

  // The dashboard builds the AP projection from pctScore; make sure it exists.
  if (rec.pctScore === undefined && Number(rec.totalPossible) > 0) {
    rec.pctScore = Math.round(Number(rec.totalScore) / Number(rec.totalPossible) * 100);
  }
  if (!rec.submitTime) rec.submitTime = new Date().toISOString();
  return rec;
}

/* ───────────────────────── Sheet read / write ───────────────────────── */

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, CORE_HEADERS.length).setValues([CORE_HEADERS]);
    styleHeader_(sh);
    sh.setFrozenRows(1);
  }
  return sh;
}

function styleHeader_(sh) {
  const lastCol = Math.max(1, sh.getLastColumn());
  sh.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground('#C41230').setFontColor('#FFFFFF');
}

function getHeaders_(sh) {
  const lastCol = sh.getLastColumn();
  return lastCol ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
}

function appendRecord_(rec) {
  const sh = getSheet_();
  let headers = getHeaders_(sh);

  // Add any new fields as new columns on the right.
  const missing = Object.keys(rec).filter(function (k) { return headers.indexOf(k) === -1; });
  if (missing.length) {
    sh.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    headers = headers.concat(missing);
    styleHeader_(sh);
  }

  // Force text format on fragile columns before writing so Sheets does not convert them.
  const nextRow = sh.getLastRow() + 1;
  TEXT_COLUMNS.forEach(function (name) {
    const idx = headers.indexOf(name);
    if (idx !== -1) sh.getRange(nextRow, idx + 1).setNumberFormat('@');
  });

  const row = headers.map(function (h) { return rec[h] === undefined ? '' : rec[h]; });
  sh.getRange(nextRow, 1, 1, headers.length).setValues([row]);
}

function readAll_() {
  const sh = getSheet_();
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return [];
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map(String);
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const obj = {};
    let hasData = false;
    for (let c = 0; c < headers.length; c++) {
      let v = values[r][c];
      if (v === '' || v === null) continue;
      hasData = true;
      if (v instanceof Date) v = v.toISOString();
      if (JSON_COLUMNS.indexOf(headers[c]) !== -1 && typeof v === 'string') {
        try { v = JSON.parse(v); } catch (x) { v = {}; }
      }
      obj[headers[c]] = v;
    }
    if (hasData) out.push(obj);
  }
  return out;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ───────────────────────── Utilities (run from the editor) ───────────────────────── */

/** Run once: creates the Results tab with headers and checks permissions. */
function setup() {
  const sh = getSheet_();
  SpreadsheetApp.getActiveSpreadsheet().toast('Results tab is ready: ' + sh.getName(), 'AP Stats Diagnostic', 5);
}

/** Adds a menu to the sheet: Diagnostic → Add test row / Clear all results. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Diagnostic')
    .addItem('Add a test row', 'addTestRow')
    .addItem('Clear all results (keeps headers)', 'clearResults')
    .addToUi();
}

/** Writes one fake submission so you can confirm the dashboard reads it. */
function addTestRow() {
  const answers = {};
  for (let i = 0; i < 42; i++) answers[i] = 'ABCD'.charAt(i % 4);
  const res = doPost({ postData: { contents: JSON.stringify({
    studentName: 'Test Student', studentId: 'test@example.com', period: 'October',
    startTime: new Date().toISOString(), submitTime: new Date().toISOString(), timeTaken: 55,
    totalScore: 20, totalPossible: 42, pctScore: 48,
    unit_Unit_1_correct: 6, unit_Unit_1_total: 12,
    rawAnswers: JSON.stringify(answers), wrongQs: '1,2,3', skippedQs: ''
  }) } });
  SpreadsheetApp.getActiveSpreadsheet().toast(res.getContent(), 'Test row', 5);
}

function clearResults() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Delete ALL results?', 'This removes every student row and cannot be undone.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const sh = getSheet_();
  if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
}
