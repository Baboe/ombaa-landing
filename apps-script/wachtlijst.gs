// Ombaa wachtlijst: slaat aanmeldingen op in de Google Sheet waar dit script aan hangt.
// Zie README-stappen in de chat: Extensies > Apps Script > plakken > Implementeren als web-app.

function doPost(e) {
  var p = (e && e.parameter) || {};
  var email = String(p.email || '').trim();
  var name = String(p.first_name || '').trim();

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return ContentService.createTextOutput('invalid');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Datum', 'Naam', 'E-mail']);
    }
    // dubbele aanmelding overslaan
    var emails = sheet.getLastRow() > 1
      ? sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues().flat()
      : [];
    var exists = emails.some(function (v) {
      return String(v).toLowerCase() === email.toLowerCase();
    });
    if (!exists) {
      sheet.appendRow([new Date(), name, email]);
    }
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput('ok');
}
