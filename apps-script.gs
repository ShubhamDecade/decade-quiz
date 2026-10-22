// How should Decade sound? Saves each quiz submission as one row in the "Answers" tab.
// Paste this into Extensions > Apps Script in the answers sheet, then deploy it as a web app.

var SHEET_NAME = 'Answers';
var FIXED = ['Respondent ID', 'Submitted at', 'Name'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getSheet_();
    var values = flatten_(data);
    var headers = ensureHeaders_(sheet, Object.keys(values));
    var row = headers.map(function (h) { return values.hasOwnProperty(h) ? values[h] : ''; });

    // One row per person: sending again replaces their earlier row.
    var target = 0, last = sheet.getLastRow();
    if (last > 1 && data.id) {
      var ids = sheet.getRange(2, 1, last - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(data.id)) { target = i + 2; break; }
      }
    }
    if (target) sheet.getRange(target, 1, 1, row.length).setValues([row]);
    else sheet.appendRow(row);
    return json_({ ok: true, updated: !!target });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Open the web app URL in a browser to check it's live.
function doGet() {
  return json_({ ok: true, message: 'The quiz sheet is listening.' });
}

function getSheet_() {
  var book = SpreadsheetApp.getActiveSpreadsheet();
  return book.getSheetByName(SHEET_NAME) || book.insertSheet(SHEET_NAME);
}

// Column names come from the quiz itself, so new moments add new columns.
function ensureHeaders_(sheet, keys) {
  var lastCol = sheet.getLastColumn();
  var headers = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].filter(String) : [];
  if (!headers.length) headers = FIXED.slice();
  keys.forEach(function (k) { if (headers.indexOf(k) < 0) headers.push(k); });
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  return headers;
}

function flatten_(data) {
  var out = {};
  out['Respondent ID'] = clean_(data.id);
  out['Submitted at'] = new Date();
  out['Name'] = clean_(data.name);
  (data.answers || []).forEach(function (a) {
    out[a.label + ' | pick'] = clean_(a.pick);
    out[a.label + ' | why'] = clean_(a.why);
    out[a.label + ' | own line'] = clean_(a.own);
  });
  out['Anything else'] = clean_(data.final);
  return out;
}

// Keeps cells short and stops anything that looks like a formula from running.
function clean_(v) {
  var s = String(v == null ? '' : v).slice(0, 2000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
