/**
 * Loan Floor · team sync (Google Apps Script web app)
 *
 * Keeps the Loan Floor program live for the whole team without anyone logging in.
 * Every laptop on the program link reads and saves here; changes show up for
 * everyone within about 2 seconds. Data is also written to a "Board" sheet so you
 * can see it.
 *
 * One-time setup (about 3 minutes, on a laptop):
 * 1. Open your TEST & PRACTICE Google Sheet (or any sheet you own) → Extensions → Apps Script.
 * 2. Click "+" next to Files → Script → name it LoanFloorSync. Paste this whole file. Save.
 *    (Keep the other script file there; they do not clash.)
 * 3. Deploy → New deployment → gear icon → Web app.
 *      Execute as: Me
 *      Who has access: Anyone
 *    → Deploy → Authorize access → your account → Advanced → Go to project → Allow.
 * 4. Copy the "Web app URL" (ends in /exec) and send it to Claude, or add ?sync=<URL> to the program link.
 *
 * Reset password: 1234
 */
const RESET_PASSWORD = "1234";
const BOARD_SHEET = "Board";

function doGet(e) {
  return out_(state_());
}

function doPost(e) {
  let req = {};
  try { req = JSON.parse((e && e.postData && e.postData.contents) || "{}"); } catch (err) { return out_({ ok: false, error: "bad request" }); }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return out_({ ok: false, error: "busy, try again" });
  try {
    const P = PropertiesService.getScriptProperties();
    const key = "loan:" + req.id;
    switch (req.action) {
      case "patch": {
        const cur = JSON.parse(P.getProperty(key) || "{}");
        P.setProperty(key, JSON.stringify(merge_(cur, req.patch || {})));
        mirror_(req.id, P.getProperty(key));
        return out_({ ok: true });
      }
      case "set":
        P.setProperty(key, JSON.stringify(req.data || {}));
        mirror_(req.id, P.getProperty(key));
        return out_({ ok: true });
      case "delete":
        P.deleteProperty(key);
        return out_({ ok: true });
      case "run":
        P.setProperty("run", JSON.stringify(req.data || {}));
        return out_({ ok: true });
      case "lock": {
        const lk = "lock:" + req.id, now = Date.now();
        const held = JSON.parse(P.getProperty(lk) || "null");
        if (held && held.until > now && held.holder !== req.holder) return out_({ ok: true, acquired: false });
        P.setProperty(lk, JSON.stringify({ holder: req.holder, until: now + (req.ttlMs || 15000) }));
        return out_({ ok: true, acquired: true });
      }
      case "reset": {
        if (String(req.password) !== RESET_PASSWORD) return out_({ ok: false, error: "wrong password" });
        const all = P.getProperties();
        Object.keys(all).forEach(k => { if (k.indexOf("loan:") === 0 || k.indexOf("lock:") === 0) P.deleteProperty(k); });
        const run = JSON.parse(all.run || "{}");
        P.setProperty("run", JSON.stringify({ allTasks: !!run.allTasks, useRounded: !!run.useRounded }));
        clearBoard_();
        return out_({ ok: true });
      }
      default:
        return out_({ ok: false, error: "unknown action" });
    }
  } finally {
    lock.releaseLock();
  }
}

function state_() {
  const all = PropertiesService.getScriptProperties().getProperties();
  const loans = {};
  Object.keys(all).forEach(k => { if (k.indexOf("loan:") === 0) loans[k.slice(5)] = JSON.parse(all[k]); });
  return { ok: true, loans: loans, run: JSON.parse(all.run || "{}"), now: Date.now() };
}

// Nested objects (like the per-step log) merge; anything else replaces.
function merge_(a, b) {
  const o = Object.assign({}, a);
  Object.keys(b).forEach(k => {
    const v = b[k];
    if (v && typeof v === "object" && !Array.isArray(v) && o[k] && typeof o[k] === "object" && !Array.isArray(o[k])) o[k] = merge_(o[k], v);
    else o[k] = v;
  });
  return o;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---- readable copy in a "Board" tab (only when the script lives inside a spreadsheet) ---- */
function boardSheet_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActive(); } catch (e) { return null; }
  if (!ss) return null;
  let sh = ss.getSheetByName(BOARD_SHEET);
  if (!sh) {
    sh = ss.insertSheet(BOARD_SHEET);
    sh.getRange(1, 1, 1, 10).setValues([["App", "LTV %", "Front %", "Back %", "Words", "Entered", "Who (last step)", "L", "H", "V"]]).setFontWeight("bold");
    for (let n = 1; n <= 49; n++) sh.getRange(n + 1, 1).setValue(n);
  }
  return sh;
}
function mirror_(id, json) {
  try {
    const sh = boardSheet_(); if (!sh) return;
    const d = JSON.parse(json || "{}"); const n = Number(String(id).replace(/\D/g, "")); if (!n) return;
    const pct = x => (x == null ? "" : Math.round(x * 100));
    const logs = d.log ? Object.keys(d.log).map(k => d.log[k]).sort((a, b) => (a.at || 0) - (b.at || 0)) : [];
    const who = logs.length ? logs[logs.length - 1].by : (d.claim && d.claim.who) || "";
    sh.getRange(n + 1, 2, 1, 9).setValues([[pct(d.ltv), pct(d.front), pct(d.back), Array.isArray(d.words) ? d.words.length : "",
      d.submitted ? "Yes" : "", who, d.L || "", d.H || "", d.V || ""]]);
  } catch (e) { /* the board keeps working even if the copy fails */ }
}
function clearBoard_() {
  try { const sh = boardSheet_(); if (sh) sh.getRange(2, 2, 49, 9).clearContent(); } catch (e) {}
}
