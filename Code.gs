/**
 * Levi's Road to Black — Sheets sync backend.
 * Stores the whole app state as one JSON blob in a "State" sheet.
 *   GET  ?action=pull            -> { ok, data, updatedAt }
 *   POST { action:'push', data, updatedAt } -> { ok, updatedAt }
 *   GET  ?action=ping            -> { ok:true, pong:true }   (for testing in a browser)
 *
 * Deploy as a Web app: Execute as = Me, Who has access = Anyone.
 * updatedAt is epoch milliseconds (a number) to avoid Sheets date-coercion.
 */

var SHEET_NAME = 'State';
var ROW_KEY = 'levi';

function doGet(e)  { return handle_(e, 'GET'); }
function doPost(e) { return handle_(e, 'POST'); }

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['key', 'data', 'updatedAt']);
  }
  return sh;
}

function findRow_(sh) {
  var vals = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][0] === ROW_KEY) return i + 1; // 1-based row number
  }
  return -1;
}

function handle_(e, method) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    var action = (e && e.parameter && e.parameter.action) || '';
    var payload = null;
    if (method === 'POST' && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
      action = payload.action || action || 'push';
    }

    if (action === 'ping') {
      return json_({ ok: true, pong: true, time: Date.now() });
    }

    var sh = getSheet_();

    if (action === 'push') {
      var data = JSON.stringify((payload && payload.data) || {});
      var ts = (payload && Number(payload.updatedAt)) || Date.now();
      var row = findRow_(sh);
      if (row === -1) {
        sh.appendRow([ROW_KEY, data, ts]);
      } else {
        sh.getRange(row, 2).setValue(data);
        sh.getRange(row, 3).setValue(ts);
      }
      return json_({ ok: true, updatedAt: ts });
    }

    // default: pull
    var r = findRow_(sh);
    if (r === -1) return json_({ ok: true, data: null, updatedAt: 0 });
    var cell = sh.getRange(r, 2).getValue();
    var when = Number(sh.getRange(r, 3).getValue()) || 0;
    var parsed = null;
    try { parsed = JSON.parse(cell || 'null'); } catch (err) { parsed = null; }
    return json_({ ok: true, data: parsed, updatedAt: when });

  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}
