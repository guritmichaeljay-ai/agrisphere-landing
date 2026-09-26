const SHEET_NAME = 'Access Records';

function setupSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME)
    || SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Name', 'Email', 'Access type', 'Farm / organization', 'Purpose', 'Date', 'Status']);
  }
}

function doPost(event) {
  setupSheet();
  const record = JSON.parse(event.postData.contents);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME).appendRow([
    record.name || '',
    record.email || '',
    record.accessType || '',
    record.organization || '',
    record.purpose || '',
    record.date || '',
    'Recorded'
  ]);
  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}