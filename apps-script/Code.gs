/**
 * RICHMAX Thesis Desk — Review Log (Google Apps Script web app)
 * Bound to the Google Sheet "RICHMAX Thesis Desk - Review Log".
 * Self-contained: paste this whole file into Extensions > Apps Script of that sheet.
 *
 * doPost : appends one row per review/note turn from ดร.วิชิต or พี่บิ๊ก (LockService + shared token, formula-escaped text).
 * doGet  : health check (open the Web App URL in a browser -> {"ok":true,...}).
 *
 * Append-only: ทุกครั้งเพิ่มแถวใหม่ ไม่ลบของเก่า
 * Restore link: คอลัมน์ลิงก์เปิดโต๊ะตรวจกลับไปสถานะเวอร์ชันนั้น (#restore=... ฝัง state ใน URL เพราะ Pages อ่านชีตส่วนตัวไม่ได้)
 *
 * NOTE: SHARED_TOKEN is also in the PUBLIC web page, so it is a demo-level filter
 *       against random traffic, not a real secret. Review notes live only in this private sheet.
 */

// ===== Settings (edit here) =====
var SHARED_TOKEN = 'richmax-td-2026';     // must match SHEETS_TOKEN in index.html
var SHEET_NAME = '';                      // '' = first tab of this spreadsheet
var TZ = 'Asia/Bangkok';
var DESK_URL = 'https://phiistphisit-glitch.github.io/richmax-thesis-desk/';
var MAX_STATE_JSON = 4000;
var MAX_RESTORE_URL = 2000;

var EMAIL_ENABLED = false;                // no email notifications for Thesis Desk

// Column layout: [payload key, header, type, max length for text]
// type: 't' = text (formula-escaped), 'n' = number (else blank)
var FIELDS = [
  ['eventType', 'Event type', 't', 20],
  ['reviewer', 'Reviewer', 't', 60],
  ['role', 'Role', 't', 40],
  ['itemId', 'Item ID', 't', 80],
  ['itemTitle', 'Item title/question', 't', 300],
  ['status', 'Status (ผ่าน/ต้องแก้ไข/ล้าง)', 't', 40],
  ['marksSummary', 'Marks summary', 't', 200],
  ['noteText', 'Note text', 't', 2000],
  ['appVersion', 'App version', 't', 40]
];
// Append-only: ทุกครั้งเพิ่มแถวใหม่ ไม่ลบ/ไม่เขียนทับของเก่า — แถวเก่าคือเวอร์ชันเก่า กด Restore link เพื่อเปิดใช้ใหม่ได้
var HEADERS = ['Timestamp (Asia/Bangkok)', 'Version']
  .concat(FIELDS.map(function (f) { return f[1]; }))
  .concat(['State JSON', 'Restore link']);

// ===== Web app entry points =====
function doGet(e) {
  return json_({ ok: true, service: 'RICHMAX Thesis Desk - Review Log', time: nowText_() });
}

function doPost(e) {
  var data;
  try {
    data = JSON.parse((e && e.postData && e.postData.contents) || '');
  } catch (err) {
    return json_({ ok: false, error: 'bad_json' });
  }
  if (!data || typeof data !== 'object' || data.token !== SHARED_TOKEN) return json_({ ok: false, error: 'bad_token' });

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return json_({ ok: false, error: 'busy' });

  var sheet, ts, rowNumber, ver;
  try {
    sheet = getSheet_();
    ensureHeaders_(sheet);
    ts = nowText_();
    ver = nextVersion_(sheet);   // เลขเวอร์ชันรันต่อเนื่อง ไม่ทับของเก่า
    sheet.appendRow(buildRow_(data, ts, ver));
    rowNumber = sheet.getLastRow();
  } finally {
    lock.releaseLock();
  }

  return json_({ ok: true, row: rowNumber, version: ver });
}

// ===== Helpers =====
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return (SHEET_NAME && ss.getSheetByName(SHEET_NAME)) || ss.getSheets()[0];
}

function nowText_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
}

// Text cells: trim, cap length, and stop spreadsheet formulas (=, +, -, @) being injected.
function txt_(v, max) {
  var s = String(v == null ? '' : v).trim().slice(0, max || 200);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
// Number cells: numbers only, otherwise blank.
function num_(v) {
  if (v === '' || v == null || typeof v === 'boolean') return '';
  var n = Number(v);
  return isFinite(n) ? n : '';
}

function nextVersion_(sheet) {
  // Version = จำนวนแถวข้อมูล (ไม่นับหัวตาราง) + 1 — ไม่ลบแถวเก่า จึงเรียง 1,2,3... ตลอด
  var last = sheet.getLastRow();
  return last <= 1 ? 1 : last; // after header at row 1, lastRow before append is previous version count+header
}

/** Ensure header row exists; append new columns (State JSON / Restore link) if sheet already had older headers. */
function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    return;
  }
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var have = {};
  for (var i = 0; i < existing.length; i++) {
    if (existing[i]) have[String(existing[i])] = true;
  }
  var missing = [];
  for (var j = 0; j < HEADERS.length; j++) {
    if (!have[HEADERS[j]]) missing.push(HEADERS[j]);
  }
  if (missing.length) {
    sheet.getRange(1, lastCol + 1, 1, lastCol + missing.length).setValues([missing]);
  }
}

function statusCode_(label) {
  var s = String(label == null ? '' : label);
  if (s === 'ผ่าน' || s === 'ok') return 'ok';
  if (s === 'ต้องแก้ไข' || s === 'fix') return 'fix';
  if (s === 'ล้าง' || s === '' || s === 'null') return null;
  return s;
}

function parseMarks_(d) {
  if (d && d.marksJson) {
    try {
      var m = JSON.parse(String(d.marksJson));
      if (m && typeof m === 'object') return m;
    } catch (e) {}
  }
  if (d && d.stateJson) {
    try {
      var full = JSON.parse(String(d.stateJson));
      if (full && full.marks && typeof full.marks === 'object') return full.marks;
    } catch (e2) {}
  }
  return {};
}

/** Compact state for URL: {id, st, mk, nt, rv, et} — self-contained restore. */
function buildCompact_(d) {
  var marks = parseMarks_(d);
  var note = String(d.noteText == null ? '' : d.noteText);
  var st = statusCode_(d.status);
  // Prefer codes from stateJson if present
  if (d.stateJson) {
    try {
      var full = JSON.parse(String(d.stateJson));
      if (full) {
        if (full.itemId) d = Object.assign({}, d, { itemId: full.itemId });
        if (full.status !== undefined) st = statusCode_(full.status);
        if (full.marks && typeof full.marks === 'object') marks = full.marks;
        if (full.noteText != null && !note) note = String(full.noteText);
        if (full.reviewer && !d.reviewer) d = Object.assign({}, d, { reviewer: full.reviewer });
      }
    } catch (e) {}
  }
  return {
    id: String(d.itemId || ''),
    st: st,
    mk: marks,
    nt: note,
    rv: String(d.reviewer || ''),
    et: String(d.eventType || '')
  };
}

function restoreUrlFromCompact_(c) {
  var compact = {
    id: c.id || '',
    st: c.st,
    mk: c.mk || {},
    nt: String(c.nt || ''),
    rv: String(c.rv || ''),
    et: String(c.et || '')
  };
  function make(c2) {
    return DESK_URL + '#restore=' + encodeURIComponent(JSON.stringify(c2));
  }
  var url = make(compact);
  if (url.length <= MAX_RESTORE_URL) return url;

  // Truncate note first
  compact.nt = String(compact.nt || '').slice(0, 120);
  compact.tr = true;
  url = make(compact);
  if (url.length <= MAX_RESTORE_URL) return url;

  // Drop mark details if still too long (keep status)
  compact.mk = {};
  compact.tr = true;
  url = make(compact);
  if (url.length <= MAX_RESTORE_URL) return url;

  // Last resort: drop note entirely
  compact.nt = '';
  return make(compact);
}

function stateJsonCell_(d) {
  if (d.stateJson) return txt_(d.stateJson, MAX_STATE_JSON);
  var full = {
    itemId: d.itemId || '',
    status: statusCode_(d.status),
    marks: parseMarks_(d),
    noteText: d.noteText || '',
    reviewer: d.reviewer || '',
    at: d.at || ''
  };
  return txt_(JSON.stringify(full), MAX_STATE_JSON);
}

function buildRow_(d, ts, ver) {
  var row = [ts, ver];
  for (var i = 0; i < FIELDS.length; i++) {
    var f = FIELDS[i];
    row.push(f[2] === 'n' ? num_(d[f[0]]) : txt_(d[f[0]], f[3]));
  }
  var stateCell = stateJsonCell_(d);
  row.push(stateCell);
  var compact = buildCompact_(d);
  row.push(restoreUrlFromCompact_(compact));
  return row;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
