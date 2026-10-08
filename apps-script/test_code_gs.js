// Offline tests for apps-script/Code.gs (Thesis Desk review log).
// Mocks SpreadsheetApp, LockService, ContentService, Utilities and a fixed clock.
// Run: node apps-script/test_code_gs.js
const fs = require("fs"), vm = require("vm");
const src = fs.readFileSync(__dirname + "/Code.gs", "utf8");
let pass = 0, fail = 0;
function ok(name, cond, extra) { cond ? pass++ : fail++; console.log(`${cond ? "PASS" : "FAIL"}  ${name}${extra !== undefined && !cond ? "  -> " + extra : ""}`); }

// A tiny in-memory sheet that behaves like Apps Script for the calls Code.gs uses.
// getRange(row, col, numRows, numCols) is strict: setValues must match numRows x numCols,
// exactly like the real Sheets API (this is what caught the old ensureHeaders_ bug).
function makeSheet(preRows) {
  const rows = (preRows || []).map(r => r.slice());
  const writes = []; // every in-place write (setValues) for audit
  const lastCol = () => rows.reduce((m, r) => { let c = r.length; while (c > 0 && (r[c - 1] === "" || r[c - 1] == null)) c--; return Math.max(m, c); }, 0);
  const sheet = {
    rows, writes,
    appendRow: r => { rows.push(r.slice()); },
    getLastRow: () => rows.length,
    getLastColumn: () => lastCol(),
    getRange: (row, col, nr, nc) => {
      if (![row, col, nr, nc].every(n => Number.isInteger(n) && n >= 1)) throw new Error(`bad getRange(${row},${col},${nr},${nc})`);
      return {
        getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => { const r = rows[row - 1 + i]; const v = r ? r[col - 1 + j] : undefined; return v == null ? "" : v; })),
        setValues: vals => {
          if (vals.length !== nr || vals.some(v => v.length !== nc)) throw new Error(`The number of columns in the data does not match the number of columns in the range. The data has ${vals[0].length} but the range has ${nc}.`);
          writes.push({ row, col, nr, nc, vals });
          vals.forEach((v, i) => { const r = rows[row - 1 + i] || (rows[row - 1 + i] = []); v.forEach((x, j) => { r[col - 1 + j] = x; }); });
        }
      };
    }
  };
  return sheet;
}

function makeEnv(opts = {}) {
  const nowMs = Date.parse("2026-10-08T08:30:00Z"); // 15:30 Bangkok
  const RealDate = Date;
  class FakeDate extends RealDate { constructor(...a) { a.length ? super(...a) : super(nowMs); } static now() { return nowMs; } }
  const sheet = makeSheet(opts.preRows);
  const lockLog = [];
  const ctx = {
    Date: FakeDate, JSON, Math, String, Number, Array, Object, isFinite, Error,
    encodeURIComponent, decodeURIComponent,
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheets: () => [sheet], getSheetByName: () => null }) },
    LockService: { getScriptLock: () => ({ tryLock: () => { lockLog.push("lock"); return !opts.lockFails; }, releaseLock: () => lockLog.push("release") }) },
    Utilities: { formatDate: (d, tz, f) => {
      const p = Object.fromEntries(new Intl.DateTimeFormat("en-GB", { timeZone: tz, year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false }).formatToParts(d).map(x => [x.type, x.value]));
      return f.replace("yyyy", p.year).replace("MM", p.month).replace("dd", p.day).replace("HH", p.hour).replace("mm", p.minute).replace("ss", p.second); } },
    ContentService: { MimeType: { JSON: "json" }, createTextOutput: t => ({ text: t, setMimeType() { return this; } }) }
  };
  vm.createContext(ctx); vm.runInContext(src, ctx);
  return { ctx, sheet, rows: sheet.rows, lockLog,
    post: body => JSON.parse(ctx.doPost({ postData: { contents: typeof body === "string" ? body : JSON.stringify(body) } }).text) };
}

// Payload shaped like the page sends (saveReview / addNote in teacher mode).
function reviewPayload(extra = {}) {
  const marks = { 0: "ok", 2: "fix" };
  return Object.assign({
    token: "richmax-td-2026", appVersion: "test", reviewer: "พี่บิ๊ก", role: "อาจารย์",
    eventType: "review", itemId: "c1-1", itemTitle: "ความเป็นมาและความสำคัญของปัญหา",
    status: "ต้องแก้ไข", marksSummary: "ผ่าน 1 · ต้องแก้ 1", noteText: "",
    marksJson: JSON.stringify(marks),
    stateJson: JSON.stringify({ itemId: "c1-1", status: "fix", marks, noteText: "", reviewer: "พี่บิ๊ก", at: 1 }), at: 1
  }, extra);
}
const restoreJson = url => { const i = url.indexOf("#restore="); return JSON.parse(decodeURIComponent(url.slice(i + 9))); };

// ---------- 1. health ----------
let E = makeEnv();
let g = JSON.parse(E.ctx.doGet({}).text);
ok("doGet health check", g.ok === true && g.time === "2026-10-08 15:30:00", JSON.stringify(g));

// ---------- 2. wrong token / bad json ----------
ok("wrong token -> bad_token, no row", E.post(reviewPayload({ token: "nope" })).error === "bad_token" && E.rows.length === 0);
ok("missing token -> bad_token", E.post({ itemId: "x" }).error === "bad_token" && E.rows.length === 0);
ok("bad JSON -> bad_json", E.post("{oops").error === "bad_json" && E.rows.length === 0);
ok("no lock taken for rejected posts", E.lockLog.length === 0);

// ---------- 3. empty sheet -> header + Version 1 ----------
let res = E.post(reviewPayload());
const H = E.ctx.HEADERS;
ok("empty sheet: header row written", JSON.stringify(E.rows[0]) === JSON.stringify(H), JSON.stringify(E.rows[0]));
ok("header has Version, State JSON, Restore link", H[1] === "Version" && H.includes("State JSON") && H[H.length - 1] === "Restore link");
ok("first post -> ok, row 2, Version 1", res.ok === true && res.row === 2 && res.version === 1 && E.rows[1][1] === 1, JSON.stringify(res));
ok("row length = header length", E.rows[1].length === H.length);
ok("lock acquired and released", E.lockLog.join(",") === "lock,release");
const col = (row, name) => row[H.indexOf(name)];
ok("text columns landed", col(E.rows[1], "Reviewer") === "พี่บิ๊ก" && col(E.rows[1], "Status (ผ่าน/ต้องแก้ไข/ล้าง)") === "ต้องแก้ไข" && col(E.rows[1], "Item ID") === "c1-1");

// ---------- 4. three posts -> Versions 1,2,3, nothing overwritten ----------
E = makeEnv();
const snaps = [];
const vers = [];
for (let i = 1; i <= 3; i++) {
  vers.push(E.post(reviewPayload({ noteText: "รอบ " + i, eventType: i === 2 ? "note" : "review" })).version);
  snaps.push(JSON.stringify(E.rows.slice(0, i + 1)));
}
ok("3 posts -> versions 1,2,3", vers.join(",") === "1,2,3", vers.join(","));
ok("Version column is 1,2,3", E.rows.slice(1).map(r => r[1]).join(",") === "1,2,3");
ok("4 rows total (header + 3)", E.rows.length === 4);
ok("no row overwritten (earlier rows unchanged after later posts)", snaps.every((s, i) => JSON.stringify(E.rows.slice(0, i + 2)) === s));
ok("no in-place writes on data rows", E.sheet.writes.every(w => w.row === 1));
ok("header not duplicated", E.rows.filter(r => r[0] === H[0]).length === 1);

// ---------- 5. nextVersion_ direct checks ----------
E = makeEnv();
ok("nextVersion_ on totally empty sheet = 1", E.ctx.nextVersion_(E.sheet, []) === 1);
E = makeEnv({ preRows: [H.slice()] });
ok("nextVersion_ header only = 1", E.ctx.nextVersion_(E.sheet, H) === 1);
E = makeEnv({ preRows: [H.slice(), ["t", 1], ["t", 2], ["t", 3], ["t", 4], ["t", 5]] });
ok("nextVersion_ header + 5 rows = 6", E.ctx.nextVersion_(E.sheet, H) === 6);
E = makeEnv({ preRows: [H.slice(), ["t", 1], ["t", 7]] });
ok("nextVersion_ never reuses a number (row deleted by hand)", E.ctx.nextVersion_(E.sheet, H) === 8);
E = makeEnv({ preRows: [H.slice(), ["t", 1], ["t", 2]] });
res = E.post(reviewPayload());
ok("sheet with data (header in row 1): next post gets Version 3 at row 4", res.version === 3 && res.row === 4 && E.rows[3][1] === 3, JSON.stringify(res));

// ---------- 6. formula injection ----------
E = makeEnv();
E.post(reviewPayload({ itemTitle: "=IMPORTXML(\"http://x\")", noteText: "+1+1", reviewer: "@me", marksSummary: "-cmd", itemId: "  =A1" }));
let r = E.rows[1];
ok("'=' escaped", col(r, "Item title/question") === "'=IMPORTXML(\"http://x\")", col(r, "Item title/question"));
ok("'+' escaped", col(r, "Note text") === "'+1+1");
ok("'@' escaped", col(r, "Reviewer") === "'@me");
ok("'-' escaped", col(r, "Marks summary") === "'-cmd");
ok("leading spaces trimmed then escaped", col(r, "Item ID") === "'=A1", col(r, "Item ID"));
ok("normal Thai text not escaped", E.ctx.txt_("ผ่าน", 10) === "ผ่าน");

// ---------- 7. Restore link ----------
E = makeEnv();
E.post(reviewPayload({ eventType: "note", noteText: "ขาดการอ้างอิง ย่อหน้า 3" }));
let link = col(E.rows[1], "Restore link");
let c = restoreJson(link);
ok("Restore link starts with DESK_URL#restore=", link.startsWith("https://phiistphisit-glitch.github.io/richmax-thesis-desk/#restore="));
ok("Restore link <= 2000 chars", link.length <= 2000, link.length);
ok("#restore= parses back to {id, st, mk, nt, rv, et}", c.id === "c1-1" && c.st === "fix" && c.mk["0"] === "ok" && c.mk["2"] === "fix" && c.nt === "ขาดการอ้างอิง ย่อหน้า 3" && c.rv === "พี่บิ๊ก" && c.et === "note", JSON.stringify(c));
ok("State JSON column is valid JSON", JSON.parse(col(E.rows[1], "State JSON")).itemId === "c1-1");
// long Thai note + many marks -> still <= 2000 and parseable
const bigMarks = {}; for (let i = 0; i < 120; i++) bigMarks[i] = i % 3 ? "ok" : "fix";
const longNote = "ข้อความยาวมาก 😀 ".repeat(400);
E = makeEnv();
E.post(reviewPayload({ eventType: "note", noteText: longNote, marksJson: JSON.stringify(bigMarks),
  stateJson: JSON.stringify({ itemId: "c2-3", status: "fix", marks: bigMarks, noteText: longNote, reviewer: "ดร.วิชิต" }) }));
link = col(E.rows[1], "Restore link");
c = restoreJson(link);
ok("very long note: Restore link still <= 2000", link.length <= 2000, link.length);
ok("very long note: #restore= still parses, truncated flag set", c.tr === true && c.id === "c2-3" && c.st === "fix" && longNote.startsWith(c.nt), JSON.stringify(c).slice(0, 120));
ok("very long note: keeps as much note as fits (> 50 chars)", c.nt.length > 50, c.nt.length);
ok("State JSON capped at 4000", col(E.rows[1], "State JSON").length <= 4000);
// status "ล้าง" -> st null
E = makeEnv();
E.post(reviewPayload({ status: "ล้าง", marksJson: "{}", stateJson: JSON.stringify({ itemId: "c1-2", status: null, marks: {} }) }));
c = restoreJson(col(E.rows[1], "Restore link"));
ok("cleared status -> st null in restore", c.st === null && c.id === "c1-2" && Object.keys(c.mk).length === 0, JSON.stringify(c));
// payload without stateJson still builds a restore link
E = makeEnv();
E.post({ token: "richmax-td-2026", eventType: "review", itemId: "c3-1", status: "ผ่าน", reviewer: "ดร.วิชิต", marksJson: "{\"1\":\"ok\"}" });
c = restoreJson(col(E.rows[1], "Restore link"));
ok("no stateJson: restore built from status/marksJson", c.id === "c3-1" && c.st === "ok" && c.mk["1"] === "ok" && c.rv === "ดร.วิชิต", JSON.stringify(c));

// ---------- 8. ensureHeaders_ on an older header row ----------
const OLD = ["Timestamp (Asia/Bangkok)", "Event type", "Reviewer", "Role", "Item ID", "Item title/question",
  "Status (ผ่าน/ต้องแก้ไข/ล้าง)", "Marks summary", "Note text", "App version"];
const oldRow = ["2026-10-06 10:00:00", "review", "ดร.วิชิต", "อาจารย์", "c1-1", "หัวข้อเดิม", "ผ่าน", "", "", "v1"];
E = makeEnv({ preRows: [OLD.slice(), oldRow.slice()] });
let hdr = E.ctx.ensureHeaders_(E.sheet);
const expectAdded = ["Version", "State JSON", "Restore link"];
ok("ensureHeaders_ appends exactly the missing columns after the old ones", JSON.stringify(E.rows[0]) === JSON.stringify(OLD.concat(expectAdded)), JSON.stringify(E.rows[0]));
ok("ensureHeaders_ uses getRange(1, 11, 1, 3) (4th arg = number of columns)", E.sheet.writes.length === 1 && E.sheet.writes[0].col === 11 && E.sheet.writes[0].nc === 3, JSON.stringify(E.sheet.writes.map(w => [w.row, w.col, w.nr, w.nc])));
ok("ensureHeaders_ returns the real header", JSON.stringify(hdr) === JSON.stringify(OLD.concat(expectAdded)));
ok("old data row untouched", JSON.stringify(E.rows[1]) === JSON.stringify(oldRow));
ok("ensureHeaders_ second call adds nothing", (E.ctx.ensureHeaders_(E.sheet), E.sheet.writes.length === 1 && E.rows[0].length === 13));
// post onto the old-layout sheet: values must line up with that sheet's own header order
E = makeEnv({ preRows: [OLD.slice(), oldRow.slice()] });
res = E.post(reviewPayload({ noteText: "แถวใหม่" }));
const hdrNow = E.rows[0], newRow = E.rows[2];
const at = name => newRow[hdrNow.indexOf(name)];
ok("old sheet: post ok, Version 2 (old row counts as v1)", res.ok && res.version === 2 && at("Version") === 2, JSON.stringify(res));
ok("old sheet: new row aligned with header (Reviewer/Note/Restore in right columns)", at("Reviewer") === "พี่บิ๊ก" && at("Note text") === "แถวใหม่" && String(at("Restore link")).includes("#restore=") && String(at("State JSON")).startsWith("{"), JSON.stringify(newRow));
ok("old sheet: new row length = header length", newRow.length === hdrNow.length);
ok("old sheet: old row still untouched", JSON.stringify(E.rows[1]) === JSON.stringify(oldRow));
// header row with a gap at the end of a wider sheet
E = makeEnv({ preRows: [H.slice(0, 5)] });
E.ctx.ensureHeaders_(E.sheet);
ok("partial header gets the rest appended in order", JSON.stringify(E.rows[0]) === JSON.stringify(H), JSON.stringify(E.rows[0]));

// ---------- 9. busy lock / server error ----------
E = makeEnv({ lockFails: true });
ok("lock busy -> busy, no row", E.post(reviewPayload()).error === "busy" && E.rows.length === 0);
E = makeEnv();
E.sheet.appendRow = () => { throw new Error("quota"); };
res = E.post(reviewPayload());
ok("sheet error -> JSON error (not a crash) and lock released", res.ok === false && res.error === "server_error" && E.lockLog.join(",") === "lock,release", JSON.stringify(res));

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
