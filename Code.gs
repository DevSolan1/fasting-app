// Google Apps Script backend.
// - Students register through the public form (doPost).
// - A repeated Student ID is REJECTED with "duplicate".
// - Personal data is only returned to the admin page, and only with ADMIN_PASSWORD.
//
// ► HOW TO SET / RESET THE ADMIN PASSWORD (run once from the editor):
//   1. Open this file in script.google.com
//   2. Change the value of NEW_PASSWORD below to whatever you want
//   3. Select the function "setupPassword" in the Run dropdown at the top
//   4. Click ▶ Run — you will see "Done." in the execution log
//   5. After that, delete or blank out NEW_PASSWORD here for security
//   6. Redeploy the web app (Deploy → Manage deployments → new version)

function setupPassword() {
  const NEW_PASSWORD = ""; // ← SET A STRONG PASSWORD (min 12 characters), then run this function once
  if (!NEW_PASSWORD || NEW_PASSWORD.length < 12) throw new Error("Password must be at least 12 characters!");
  PropertiesService.getScriptProperties().setProperty("ADMIN_PASSWORD", NEW_PASSWORD);
  Logger.log("Done. Admin password has been set successfully."); // password intentionally not logged
}

const SHEET = "Registrations";
const HEAD = ["Time", "Student ID", "Full name", "Department", "Year", "Phone",
              "Group", "Note", "Language"];

function norm(s) { return String(s).replace(/\s+/g, "").toLowerCase(); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const d = JSON.parse(e.postData.contents);
    if (d.action === "list") return adminList(d);
    if (d.action === "delete") return adminDelete(d);
    if (d.action === "setPassword") return adminSetPassword(d);
    return register(d);
  } catch (err) {
    return out({ ok: false, error: "server" });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  if (sh.getLastRow() === 0) sh.appendRow(HEAD);
  return sh;
}

function register(d) {
  const need = ["name", "studentId", "department", "year", "phone", "group"];
  for (const k of need) if (!d[k] || !String(d[k]).trim()) return out({ ok: false, error: "missing" });
  if (d.group !== "A" && d.group !== "B") return out({ ok: false, error: "missing" });
  d.studentId = String(d.studentId).replace(/\s+/g, "").toUpperCase();
  if (!/^(NSR|SSR)\/\d+\/\d+$/.test(d.studentId)) return out({ ok: false, error: "idformat" });
  // validate phone: digits, spaces, dashes, +, parens only; 7–15 chars
  const phone = String(d.phone).replace(/\s/g, "");
  if (!/^\+?[\d\-()]{7,15}$/.test(phone)) return out({ ok: false, error: "phoneformat" });

  const sh = sheet();
  if (sh.getLastRow() > 1) {
    const ids = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().flat();
    if (ids.some(v => norm(v) === norm(d.studentId))) return out({ ok: false, error: "duplicate" });
  }
  // only save known fields — drop consent and anything unexpected
  const row = [new Date(), d.studentId, d.name, d.department, d.year, d.phone,
               d.group, d.note || "", d.lang || ""]
    .map(v => (typeof v === "string" ? clean(v) : v));
  sh.appendRow(row);
  return out({ ok: true });
}

function adminList(d) {
  if (!isAdmin(d)) {
    Utilities.sleep(2000); // slows down password guessing
    return out({ ok: false, error: "auth" });
  }
  const sh = sheet();
  const rows = sh.getLastRow() > 1
    ? sh.getRange(2, 1, sh.getLastRow() - 1, HEAD.length).getValues()
        .map(r => r.map(v => (v instanceof Date ? v.toISOString() : v)))
    : [];
  return out({ ok: true, head: HEAD, rows: rows });
}

// Public link shows nothing. Counts and data are only available on admin.html (password).
function doGet() {
  return out({ ok: true });
}

// stop spreadsheet formula injection (cells starting with = + - @)
function clean(s) {
  s = s.trim().slice(0, 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function isAdmin(d) {
  const real = PropertiesService.getScriptProperties().getProperty("ADMIN_PASSWORD");
  return !!real && d.password === real;
}

// Admin only: change the admin password.
function adminSetPassword(d) {
  if (!isAdmin(d)) { Utilities.sleep(2000); return out({ ok: false, error: "auth" }); }
  const np = String(d.newPassword || "").trim();
  if (np.length < 12) return out({ ok: false, error: "tooshort" });
  PropertiesService.getScriptProperties().setProperty("ADMIN_PASSWORD", np);
  return out({ ok: true });
}

// Admin only: remove one registration by Student ID (test entries, mistakes).
function adminDelete(d) {
  if (!isAdmin(d)) { Utilities.sleep(2000); return out({ ok: false, error: "auth" }); }
  const sh = sheet();
  if (sh.getLastRow() > 1) {
    const ids = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().flat();
    const i = ids.findIndex(v => norm(v) === norm(d.studentId));
    if (i >= 0) { sh.deleteRow(i + 2); return out({ ok: true }); }
  }
  return out({ ok: false, error: "notfound" });
}