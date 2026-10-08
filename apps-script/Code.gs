/**
 * RICHMAX Thesis Desk — Review Log (Google Apps Script web app)
 * Bound to the Google Sheet "RICHMAX Thesis Desk - Review Log".
 * Self-contained: paste this whole file into Extensions > Apps Script of that sheet.
 *
 * doPost : appends one row (one Version) per review/note turn from ดร.วิชิต or พี่บิ๊ก (LockService + shared token, formula-escaped text).
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

  var rowNumber, ver;
  try {
    var sheet = getSheet_();
    var header = ensureHeaders_(sheet);      // หัวตารางจริงในแถว 1 (ชีตเก่าอาจเรียงคอลัมน์ต่างจาก HEADERS)
    ver = nextVersion_(sheet, header);       // เลขเวอร์ชันรันต่อเนื่อง ไม่ทับของเก่า
    var rec = buildRecord_(data, nowText_(), ver);
    sheet.appendRow(rowForHeader_(header, rec)); // append-only: เพิ่มแถวใหม่ท้ายชีตเสมอ
    rowNumber = sheet.getLastRow();
  } catch (err2) {
    return json_({ ok: false, error: 'server_error', message: String(err2 && err2.message || err2) });
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

/**
 * Version ถัดไป = max(จำนวนแถวข้อมูล, Version สูงสุดที่มีอยู่) + 1
 * - ชีตว่าง / มีแค่หัวตาราง -> 1
 * - หัวตารางแถว 1 + ข้อมูล n แถว -> n + 1 (แถวเก่าก่อนมีคอลัมน์ Version ก็นับด้วย)
 * - ถ้ามีคนลบแถวกลางชีตด้วยมือ ก็ยังไม่ออกเลขซ้ำ เพราะดูเลขสูงสุดด้วย
 */
function nextVersion_(sheet, header) {
  var dataRows = Math.max(sheet.getLastRow() - 1, 0);
  var maxVer = 0;
  var hdr = header || [];
  var vCol = -1;
  for (var i = 0; i < hdr.length; i++) if (String(hdr[i]) === 'Version') { vCol = i + 1; break; }
  if (vCol > 0 && dataRows > 0) {
    var vals = sheet.getRange(2, vCol, dataRows, 1).getValues();
    for (var r = 0; r < vals.length; r++) {
      var n = Number(vals[r][0]);
      if (isFinite(n) && n > maxVer) maxVer = Math.floor(n);
    }
  }
  return Math.max(dataRows, maxVer) + 1;
}

/**
 * Ensure header row exists; append missing columns (e.g. Version / State JSON / Restore link)
 * to the right of an older header row. Never moves or overwrites existing columns or data.
 * Returns the header row as it is now in row 1.
 */
function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    return HEADERS.slice();
  }
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  // ตัดช่องว่างท้ายหัวตาราง เพื่อให้คอลัมน์ใหม่ต่อท้ายหัวเดิมพอดี
  var used = existing.length;
  while (used > 0 && (existing[used - 1] === '' || existing[used - 1] == null)) used--;
  var have = {};
  for (var i = 0; i < used; i++) {
    if (existing[i] !== '' && existing[i] != null) have[String(existing[i])] = true;
  }
  var missing = [];
  for (var j = 0; j < HEADERS.length; j++) {
    if (!have[HEADERS[j]]) missing.push(HEADERS[j]);
  }
  var header = existing.slice(0, used).map(function (h) { return h == null ? '' : String(h); });
  if (missing.length) {
    // getRange(row, column, numRows, numColumns) — อาร์กิวเมนต์ที่ 4 คือ "จำนวนคอลัมน์"
    sheet.getRange(1, used + 1, 1, missing.length).setValues([missing]);
    header = header.concat(missing);
  }
  return header;
}

function statusCode_(label) {
  var s = String(label == null ? '' : label);
  if (s === 'ผ่าน' || s === 'ok') return 'ok';
  if (s === 'ต้องแก้ไข' || s === 'fix') return 'fix';
  if (s === 'ล้าง' || s === '' || s === 'null') return null;
  return s;
}

function parseState_(v) {
  return (v && typeof v === 'object') ? v : JSON.parse(String(v));
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
      var full = parseState_(d.stateJson);
      if (full && full.marks && typeof full.marks === 'object') return full.marks;
    } catch (e2) {}
  }
  return {};
}

/** Keep only {paragraphIndex: 'ok'|'fix'} so the restore link stays short and safe. */
function cleanMarks_(m) {
  var out = {};
  if (!m || typeof m !== 'object') return out;
  for (var k in m) {
    if (Object.prototype.hasOwnProperty.call(m, k) && /^\d{1,4}$/.test(k) && (m[k] === 'ok' || m[k] === 'fix')) out[k] = m[k];
  }
  return out;
}

/** Compact state for URL: {id, st, mk, nt, rv, et} — self-contained restore. */
function buildCompact_(d) {
  var marks = parseMarks_(d);
  var note = String(d.noteText == null ? '' : d.noteText);
  var st = statusCode_(d.status);
  // Prefer codes from stateJson if present
  if (d.stateJson) {
    try {
      var full = parseState_(d.stateJson);
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
    st: (st === 'ok' || st === 'fix') ? st : null,
    mk: cleanMarks_(marks),
    nt: note,
    rv: String(d.reviewer || ''),
    et: String(d.eventType || '')
  };
}

// ตัดข้อความตามจำนวนตัวอักษรจริง (ไม่ผ่ากลางอีโมจิ ซึ่งจะทำให้ encodeURIComponent พัง)
function cut_(s, n) {
  var a = Array.from(String(s == null ? '' : s));
  return a.length <= n ? a.join('') : a.slice(0, Math.max(n, 0)).join('');
}

function restoreUrlFromCompact_(c) {
  var compact = {
    id: String(c.id || ''),
    st: c.st === undefined ? null : c.st,
    mk: (c.mk && typeof c.mk === 'object') ? c.mk : {},
    nt: String(c.nt || ''),
    rv: String(c.rv || ''),
    et: String(c.et || '')
  };
  function make(c2) {
    return DESK_URL + '#restore=' + encodeURIComponent(JSON.stringify(c2));
  }
  function fits(c2) { return make(c2).length <= MAX_RESTORE_URL; }
  // ตัดโน้ตให้ยาวที่สุดเท่าที่ลิงก์ยังไม่เกิน MAX_RESTORE_URL (binary search)
  function fitNote(c2, full) {
    var arr = Array.from(full), lo = 0, hi = arr.length;
    while (lo < hi) {
      var mid = Math.ceil((lo + hi) / 2);
      c2.nt = arr.slice(0, mid).join('');
      if (fits(c2)) lo = mid; else hi = mid - 1;
    }
    c2.nt = arr.slice(0, lo).join('');
    return c2;
  }

  if (fits(compact)) return make(compact);

  var fullNote = compact.nt;
  compact.tr = true;              // ธงบอกว่าลิงก์นี้ถูกตัดบางส่วน (ฉบับเต็มอยู่ในคอลัมน์ State JSON)
  compact.nt = '';
  if (fits(compact)) return make(fitNote(compact, fullNote));

  // ติ๊กย่อหน้ายาวเกิน: ตัดรายละเอียดติ๊กออก (เก็บสถานะหัวข้อไว้)
  compact.mk = {};
  if (fits(compact)) return make(fitNote(compact, fullNote));

  compact.nt = '';
  compact.rv = cut_(compact.rv, 20);
  compact.id = cut_(compact.id, 80);
  return make(compact);
}

function stateJsonCell_(d) {
  if (d.stateJson) return txt_(typeof d.stateJson === 'string' ? d.stateJson : JSON.stringify(d.stateJson), MAX_STATE_JSON);
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

/** One log record keyed by header name. */
function buildRecord_(d, ts, ver) {
  var rec = {};
  rec[HEADERS[0]] = ts;
  rec['Version'] = ver;
  for (var i = 0; i < FIELDS.length; i++) {
    var f = FIELDS[i];
    rec[f[1]] = f[2] === 'n' ? num_(d[f[0]]) : txt_(d[f[0]], f[3]);
  }
  rec['State JSON'] = stateJsonCell_(d);
  rec['Restore link'] = restoreUrlFromCompact_(buildCompact_(d));
  return rec;
}

/** Row in the order of the sheet's real header row (unknown header -> blank cell). */
function rowForHeader_(header, rec) {
  var row = [];
  for (var i = 0; i < header.length; i++) {
    var k = String(header[i]);
    row.push(Object.prototype.hasOwnProperty.call(rec, k) ? rec[k] : '');
  }
  return row;
}

/** Row in standard HEADERS order. */
function buildRow_(d, ts, ver) {
  return rowForHeader_(HEADERS, buildRecord_(d, ts, ver));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
