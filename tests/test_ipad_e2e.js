// End-to-end test: drives index.html like a reviewer on an iPad (touch, portrait + landscape).
// Serves the repo on localhost so the service worker can register, and blocks every outside
// request (Google Fonts, Sheets web app) so the run is offline and never writes to the real Sheet.
// Run: node tests/test_ipad_e2e.js        (needs Playwright + Chromium)
// Screenshots: SHOTS=/some/dir node tests/test_ipad_e2e.js
const fs = require("fs"), http = require("http"), path = require("path");
const { chromium, devices } = require("playwright");
const root = path.join(__dirname, "..");
let pass = 0, fail = 0;
function ok(name, cond, extra) { cond ? pass++ : fail++; console.log(`${cond ? "PASS" : "FAIL"}  ${name}${extra !== undefined && !cond ? "  -> " + extra : ""}`); }

const TYPES = { ".html": "text/html; charset=utf-8", ".js": "text/javascript", ".webmanifest": "application/manifest+json", ".png": "image/png" };
const server = http.createServer((req, res) => {
  const p = path.join(root, decodeURIComponent(req.url.split("?")[0]).replace(/\/$/, "/index.html"));
  if (!p.startsWith(root) || !fs.existsSync(p) || fs.statSync(p).isDirectory()) { res.writeHead(404); return res.end(); }
  res.writeHead(200, { "Content-Type": TYPES[path.extname(p)] || "application/octet-stream" });
  fs.createReadStream(p).pipe(res);
});

const TOUCH_TARGETS = ".role button, .tabs button, .vbtn, .mbar button";

async function run(browser, base, label, device) {
  const ctx = await browser.newContext({ ...device });
  const page = await ctx.newPage();
  const errors = [];
  page.on("pageerror", e => errors.push(e.message));
  await page.route(url => !url.href.startsWith(base), r => r.abort());
  await page.goto(base + "/index.html");
  await page.waitForSelector("#qlist .qi", { state: "attached" });
  const t = s => `[${label}] ${s}`;

  ok(t("55 review items listed"), await page.locator("#qlist .qi").count() === 55);
  ok(t("progress starts at 0/55"), /0\/55/.test(await page.textContent("#progTxt")), await page.textContent("#progTxt"));
  ok(t("no horizontal page scroll"), await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth + 1));

  const small = await page.$$eval(TOUCH_TARGETS, els => els
    .filter(e => e.offsetParent !== null)
    .map(e => ({ t: e.textContent.trim().slice(0, 20) || e.getAttribute("aria-label"), h: Math.round(e.getBoundingClientRect().height) }))
    .filter(b => b.h < 44));
  ok(t("visible touch buttons are >= 44px tall"), small.length === 0, JSON.stringify(small));

  // Verdict: tap "ผ่าน" on the current item; progress moves and survives a reload.
  await page.tap(".vbtn.ok");
  ok(t("tap ผ่าน marks the item"), await page.getAttribute(".vbtn.ok", "aria-pressed") === "true");
  ok(t("progress moves to 1/55"), /1\/55/.test(await page.textContent("#progTxt")), await page.textContent("#progTxt"));
  await page.reload();
  await page.waitForSelector("#qlist .qi", { state: "attached" });
  ok(t("verdict kept after reload"), /1\/55/.test(await page.textContent("#progTxt")) && await page.getAttribute(".vbtn.ok", "aria-pressed") === "true");

  // Narrow layouts swap the question column for the mobile bar (‹ รายการคำถาม ›).
  if (await page.isVisible(".mbar")) {
    const before = await page.textContent("#mCur");
    await page.tap("#mNext");
    ok(t("› moves to the next item"), await page.textContent("#mCur") !== before, before);
    await page.tap("#mList");
    ok(t("รายการคำถาม opens the question list"), await page.isVisible("#qcol"));
  } else {
    ok(t("question column visible side by side"), await page.isVisible("#qcol"));
  }

  await page.tap('.tabs [data-tab="cal"]');
  ok(t("calendar tab shows calendar"), await page.isVisible("#calView") && !(await page.isVisible("#deskView")));
  await page.tap('.tabs [data-tab="sum"]');
  ok(t("summary tab shows summary"), await page.isVisible("#sumView"));

  const sw = await page.evaluate(() => Promise.race([
    navigator.serviceWorker.ready.then(() => true), new Promise(r => setTimeout(() => r(false), 5000))]));
  ok(t("service worker registers (offline / home-screen use)"), sw);
  ok(t("no uncaught page errors"), errors.length === 0, JSON.stringify(errors));

  if (process.env.SHOTS) {
    fs.mkdirSync(process.env.SHOTS, { recursive: true });
    await page.tap('.tabs [data-tab="outline"]');
    await page.screenshot({ path: path.join(process.env.SHOTS, `ipad-${label}.png`) });
  }
  await ctx.close();
}

(async () => {
  await new Promise(r => server.listen(0, "127.0.0.1", r));
  const base = `http://127.0.0.1:${server.address().port}`;
  const browser = await chromium.launch();
  try {
    await run(browser, base, "portrait", devices["iPad (gen 7)"]);
    await run(browser, base, "landscape", devices["iPad (gen 7) landscape"]);
  } finally {
    await browser.close();
    server.close();
  }
  console.log(`\n${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
