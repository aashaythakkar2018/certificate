/**
 * Rhytara certificates -> Google Sheet.
 *
 * Paste this whole file into the Sheet's Extensions -> Apps Script editor,
 * set SECRET below to the same value as GOOGLE_SHEETS_SECRET in the
 * dashboard's .env, then Deploy -> New deployment -> Web app
 * (Execute as: Me, Who has access: Anyone). Put the web app URL in .env as
 * GOOGLE_SHEETS_WEBHOOK_URL. Full steps: README section 13.
 *
 * Each certificate is one row, matched by Certificate ID: a new certificate
 * adds a row, a name correction or regeneration updates that same row.
 */

var SECRET = 'paste-the-GOOGLE_SHEETS_SECRET-value-here';
var SHEET_NAME = 'Certificates';

var COLUMNS = [
  ['certificateId', 'Certificate ID'],
  ['issuedAt', 'Date issued'],
  ['orderNumber', 'Order number'],
  ['customerName', 'Customer name'],
  ['customerEmail', 'Customer email'],
  ['design', 'Design'],
  ['designCode', 'Design code'],
  ['sku', 'SKU'],
  ['edition', 'Edition'],
  ['status', 'Status'],
  ['updatedAt', 'Last updated'],
];
var DATE_FIELDS = { issuedAt: true, updatedAt: true };

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.secret !== SECRET) return reply({ ok: false, error: 'Wrong secret' });

    lock.waitLock(30000);
    var sheet = getSheet();
    var ids = sheet.getLastRow() > 1
      ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]); })
      : [];

    var rows = body.rows || [];
    rows.forEach(function (row) {
      var values = COLUMNS.map(function (c) {
        var v = row[c[0]];
        if (DATE_FIELDS[c[0]] && v) return new Date(v);
        // Leading apostrophe keeps order numbers/SKUs as text (no "1042" -> 1,042).
        if (c[0] === 'orderNumber' || c[0] === 'sku') return v ? "'" + v : '';
        return v == null ? '' : v;
      });
      var index = ids.indexOf(String(row.certificateId));
      if (index >= 0) {
        sheet.getRange(index + 2, 1, 1, values.length).setValues([values]);
      } else {
        sheet.appendRow(values);
        ids.push(String(row.certificateId));
      }
    });
    return reply({ ok: true, written: rows.length });
  } catch (err) {
    return reply({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** Lets you open the web app URL in a browser to check it's deployed. */
function doGet() {
  return reply({ ok: true, message: 'Rhytara certificate sync is running.' });
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS.map(function (c) { return c[1]; }));
    sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange('B:B').setNumberFormat('d mmm yyyy');
    sheet.getRange('K:K').setNumberFormat('d mmm yyyy, h:mm am/pm');
  }
  return sheet;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
