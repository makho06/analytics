const SPREADSHEET_ID = '1cX994z94XDP5j9ZNsjaN4KdgMnmLXTOCac8fdF4AHEQ';
const SHEET_NAME = ''; // Leave blank to use the first sheet, or enter its tab name.

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: 'A/B feedback endpoint is running.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const book = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = SHEET_NAME ? book.getSheetByName(SHEET_NAME) : book.getSheets()[0];
    if (!sheet) throw new Error('The requested sheet tab was not found.');

    const record = JSON.parse((e.postData && e.postData.contents) || '{}');
    const headers = ['Timestamp', 'Layout', 'Response', 'Session ID'];
    if (sheet.getLastRow() === 0) sheet.appendRow(headers);

    // The browser already prevents repeat submissions in normal use.
    // This server-side check also prevents the same session ID being recorded twice.
    const lastRow = sheet.getLastRow();
    const sessionIds = lastRow > 1
      ? sheet.getRange(2, 4, lastRow - 1, 1).getValues().flat().map(String)
      : [];
    if (record.sessionId && sessionIds.includes(String(record.sessionId))) {
      return json({ ok: true, duplicate: true });
    }

    sheet.appendRow([
      new Date(),
      String(record.variant || ''),
      Number(record.response || 0),
      String(record.sessionId || '')
    ]);
    return json({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function json(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
