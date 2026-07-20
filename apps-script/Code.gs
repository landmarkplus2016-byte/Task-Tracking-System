/**
 * AppList backend — Google Apps Script Web App
 * ─────────────────────────────────────────────────────────────────────
 * Backs the "Settings" admin tab of the Telecom tracking web app.
 * Reads/writes the list data that used to live in list.xlsx.
 *
 * SHEET STRUCTURE — create a Google Sheet with these THREE tabs.
 * Row 1 of each tab is the header row (exact names below); data starts row 2.
 *
 *   Tab "Google Sheets URLs"
 *     A: Sheet Name   B: URL
 *
 *   Tab "Team Salaries"
 *     A: Name   B: Salary   C: Account Number      (B = FULL MONTHLY salary)
 *
 *   Tab "Driver Salaries"
 *     A: Name   B: Salary   C: Account Number      (B = FULL MONTHLY salary)
 *
 * DEPLOY
 *   1. Extensions → Apps Script. Paste this file in. Save.
 *   2. Set PASSWORD below to your admin password.
 *   3. Deploy → New deployment → type "Web app".
 *        Execute as: Me
 *        Who has access: Anyone with the link
 *   4. Copy the "/exec" Web app URL. Paste it into the app config
 *      (APPLIST_ENDPOINT in appData.js) and give the password to the admin.
 *
 * NOTE ON CORS: the browser client sends POST with Content-Type text/plain
 * (a CORS "simple request") so no preflight is needed. The body is a JSON
 * string. Do NOT switch the client to application/json — it triggers a
 * preflight OPTIONS that Apps Script web apps cannot answer.
 */

// ── CONFIG ──────────────────────────────────────────────────────────────
var PASSWORD = 'CHANGE_ME';   // ← set your admin password here

var URLS_TAB    = 'Google Sheets URLs';
var TEAM_TAB    = 'Team Salaries';
var DRIVERS_TAB = 'Driver Salaries';

// ── READ ────────────────────────────────────────────────────────────────
function doGet() {
  var data = {
    sheetUrls: readUrls_(),
    salaries: {
      team:    readSalaries_(TEAM_TAB),
      drivers: readSalaries_(DRIVERS_TAB),
    },
  };
  return json_(data);
}

// ── WRITE ───────────────────────────────────────────────────────────────
function doPost(e) {
  var payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'Bad JSON payload.' });
  }

  if (!payload || payload.password !== PASSWORD) {
    return json_({ ok: false, error: 'Wrong password.' });
  }

  var data = payload.data || {};
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);   // serialise concurrent saves
  try {
    writeUrls_(data.sheetUrls || []);
    writeSalaries_(TEAM_TAB,    (data.salaries && data.salaries.team)    || []);
    writeSalaries_(DRIVERS_TAB, (data.salaries && data.salaries.drivers) || []);
  } finally {
    lock.releaseLock();
  }

  // Return the freshly-saved state so the client can confirm.
  return json_({
    ok: true,
    data: {
      sheetUrls: readUrls_(),
      salaries: { team: readSalaries_(TEAM_TAB), drivers: readSalaries_(DRIVERS_TAB) },
    },
  });
}

// ── Sheet helpers ───────────────────────────────────────────────────────
function sheet_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Tab "' + name + '" not found.');
  return sh;
}

function rows_(name) {
  var sh = sheet_(name);
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2) return [];                    // header only / empty
  return sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

function readUrls_() {
  return rows_(URLS_TAB)
    .map(function (r) {
      return { name: String(r[0] || '').trim(), url: String(r[1] || '').trim() };
    })
    .filter(function (o) { return o.name || o.url; });
}

function readSalaries_(tab) {
  return rows_(tab)
    .map(function (r) {
      return {
        name:        String(r[0] || '').trim(),
        salary:      String(r[1] == null ? '' : r[1]).trim(),   // full MONTHLY salary
        bankAccount: String(r[2] == null ? '' : r[2]).trim(),
      };
    })
    .filter(function (o) { return o.name; });
}

// Overwrite a tab's data area (keeps row-1 header) with the given matrix.
function replaceData_(name, header, matrix) {
  var sh = sheet_(name);
  // Reset header row so column names stay canonical.
  sh.getRange(1, 1, 1, header.length).setValues([header]);
  // Clear any previous data below the header.
  var lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).clearContent();
  }
  if (matrix.length) {
    sh.getRange(2, 1, matrix.length, header.length).setValues(matrix);
  }
}

function writeUrls_(list) {
  var matrix = list
    .filter(function (o) { return (o.name || o.url); })
    .map(function (o) { return [String(o.name || '').trim(), String(o.url || '').trim()]; });
  replaceData_(URLS_TAB, ['Sheet Name', 'URL'], matrix);
}

function writeSalaries_(tab, list) {
  var matrix = list
    .filter(function (o) { return o.name; })
    .map(function (o) {
      return [
        String(o.name || '').trim(),
        String(o.salary == null ? '' : o.salary).trim(),
        String(o.bankAccount == null ? '' : o.bankAccount).trim(),
      ];
    });
  replaceData_(tab, ['Name', 'Salary', 'Account Number'], matrix);
}

// ── JSON response ───────────────────────────────────────────────────────
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
