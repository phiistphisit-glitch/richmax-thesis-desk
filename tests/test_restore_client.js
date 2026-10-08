// Node test for the page's restore-link parser (<script id="restore-lib"> in index.html),
// round-tripped against the real Restore links built by apps-script/Code.gs.
// Run: node tests/test_restore_client.js
const fs = require("fs"), vm = require("vm"), path = require("path");
const root = path.join(__dirname, "..");
const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
let pass = 0, fail = 0;
function ok(name, cond, extra) { cond ? pass++ : fail++; console.log(`${cond ? "PASS" : "FAIL"}  ${name}${extra !== undefined && !cond ? "  -> " + extra : ""}`); }

const m = /<script id="restore-lib">([\s\S]*?)<\/script>/.exec(html);
ok("index.html has <script id=\"restore-lib\">", !!m);
const win = {};
vm.runInNewContext(m[1], { window: win, JSON, String, Object, Array, decodeURIComponent });
const P = win.ThesisDeskRestore;
ok("restore-lib exposes parseRestore / rawFromLocation", P && typeof P.parseRestore === "function" && typeof P.rawFromLocation === "function");

// Load Code.gs to build links exactly like the Sheet does.
const gs = { JSON, Math, String, Number, Array, Object, isFinite, encodeURIComponent };
vm.createContext(gs); vm.runInContext(fs.readFileSync(path.join(root, "apps-script/Code.gs"), "utf8"), gs);
const hashOf = url => url.slice(url.indexOf("#"));

// 1. round trip: status fix + marks + Thai note
let url = gs.restoreUrlFromCompact_({ id: "c1-1", st: "fix", mk: { 0: "ok", 2: "fix" }, nt: "ขาดการอ้างอิง 100% & \"คำพูด\"", rv: "พี่บิ๊ก", et: "note" });
let raw = P.rawFromLocation(hashOf(url), "");
let st = P.parseRestore(raw);
ok("round trip: id/status", st && st.id === "c1-1" && st.status === "fix", JSON.stringify(st));
ok("round trip: per-paragraph marks", st.marks["0"] === "ok" && st.marks["2"] === "fix" && Object.keys(st.marks).length === 2);
ok("round trip: note text with % & quotes", st.noteText === "ขาดการอ้างอิง 100% & \"คำพูด\"", st.noteText);
ok("round trip: reviewer + event", st.reviewer === "พี่บิ๊ก" && st.eventType === "note");

// 2. status ok / cleared
st = P.parseRestore(P.rawFromLocation(hashOf(gs.restoreUrlFromCompact_({ id: "s-3", st: "ok", mk: {} })), ""));
ok("status ok", st.status === "ok" && st.id === "s-3");
st = P.parseRestore(P.rawFromLocation(hashOf(gs.restoreUrlFromCompact_({ id: "c1-2", st: null, mk: {} })), ""));
ok("status cleared -> null", st.status === null);

// 3. browser that already decoded the hash (old Firefox style) still works
st = P.parseRestore(decodeURIComponent(raw));
ok("already-decoded hash still parses", st && st.id === "c1-1" && st.marks["2"] === "fix");

// 4. ?restore= query form
st = P.parseRestore(P.rawFromLocation("", "?x=1&restore=" + encodeURIComponent(JSON.stringify({ id: "c4-1", st: "fix", mk: { 1: "fix" } }))));
ok("?restore= query also works", st && st.id === "c4-1" && st.marks["1"] === "fix");

// 5. truncated long link
url = gs.restoreUrlFromCompact_({ id: "c2-3", st: "fix", mk: { 1: "ok" }, nt: "ยาว ".repeat(3000), rv: "ดร.วิชิต", et: "note" });
st = P.parseRestore(P.rawFromLocation(hashOf(url), ""));
ok("long note link <= 2000 and parses with truncated flag", url.length <= 2000 && st.truncated === true && st.status === "fix" && st.marks["1"] === "ok", url.length);

// 6. junk is rejected / sanitised
ok("no #restore -> null", P.rawFromLocation("#other", "") === null);
ok("broken JSON -> null", P.parseRestore("%7Bbroken") === null);
ok("missing id -> null", P.parseRestore(encodeURIComponent('{"st":"ok"}')) === null);
ok("array -> null", P.parseRestore(encodeURIComponent('[1,2]')) === null);
st = P.parseRestore(encodeURIComponent(JSON.stringify({ id: "c1-1", st: "hack", mk: { 0: "ok", x: "fix", 1: "<b>", 2: "fix" } })));
ok("unknown status -> null, bad marks dropped", st.status === null && JSON.stringify(st.marks) === '{"0":"ok","2":"fix"}', JSON.stringify(st));
st = P.parseRestore(encodeURIComponent(JSON.stringify({ itemId: "c1-1", status: "ต้องแก้ไข", marks: { 3: "ok" }, noteText: "เก่า", reviewer: "ดร.วิชิต" })));
ok("long-name keys (State JSON style) also accepted", st.id === "c1-1" && st.status === "fix" && st.marks["3"] === "ok" && st.noteText === "เก่า" && st.reviewer === "ดร.วิชิต");

// 7. guard rails in the page source
ok("SHEETS_WEBAPP_URL is still empty (no fake URL)", /const SHEETS_WEBAPP_URL = '';/.test(html));
ok("toast text 'โหลดเวอร์ชันเก่าแล้ว' present", html.includes("toast('โหลดเวอร์ชันเก่าแล้ว')"));

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
