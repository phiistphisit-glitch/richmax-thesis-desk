/**
 * RICHMAX Thesis Desk — Review Log (Google Apps Script web app)
 * Bound to the Google Sheet "RICHMAX Thesis Desk - Review Log".
 * Self-contained: paste this whole file into Extensions > Apps Script of that sheet.
 *
 * doPost : appends one row per review/note turn from ดร.วิชิต or พี่บิ๊ก (LockService + shared token, formula-escaped text).
 * doGet  : health check (open the Web App URL in a browser -> {"ok":true,...}).
 *
 * NOTE: SHARED_TOKEN is also in the PUBLIC web page, so it is a demo-level filter
 *       against random traffic, not a real secret. Review notes live only in this private sheet.
 */

// ===== Settings (edit here) =====
var SHARED_TOKEN = 'richmax-td-2026';     // must match SHEETS_TOKEN in index.html
var SHEET_NAME = '';                      // '' = first tab of this spreadsheet
var TZ = 'Asia/Bangkok';

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
// Append-only: ทุกครั้งเพิ่มแถวใหม่ ไม่ลบ/ไม่เขียนทับของเก่า — แถวเก่าคือเวอร์ชันเก่า เปิดดูได้ตาม Timestamp / Version
var HEADERS = ['Timestamp (Asia/Bangkok)', 'Version'].concat(FIELDS.map(function (f) { return f[1]; }));

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

  var sheet, ts, rowNumber;
  try {
    sheet = getSheet_();
    if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);   // creates the header row on an empty sheet
    ts = nowText_();
    var ver = nextVersion_(sheet);   // เลขเวอร์ชันรันต่อเนื่อง ไม่ทับของเก่า
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
function buildRow_(d, ts, ver) {
  var row = [ts, ver];
  for (var i = 0; i < FIELDS.length; i++) {
    var f = FIELDS[i];
    row.push(f[2] === 'n' ? num_(d[f[0]]) : txt_(d[f[0]], f[3]));
  }
  return row;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
