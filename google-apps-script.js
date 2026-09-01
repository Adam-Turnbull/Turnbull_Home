/**
 * Ridge House booking backend.
 * Paste this whole file into script.google.com (see setup steps in chat),
 * bound to the Google Sheet you want bookings saved to.
 */

const SHEET_NAME = 'Bookings';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Room', 'Check-in', 'Check-out', 'Name', 'Email', 'Submitted At']);
  }
  return sheet;
}

// Returns all bookings as JSON so the website can grey out taken dates.
function doGet(e) {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();
  const bookings = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    bookings.push({
      room: row[0],
      checkin: formatDate_(row[1]),
      checkout: formatDate_(row[2]),
      name: row[3],
      email: row[4]
    });
  }
  return ContentService
    .createTextOutput(JSON.stringify(bookings))
    .setMimeType(ContentService.MimeType.JSON);
}

// Appends a new booking row.
function doPost(e) {
  const sheet = getSheet_();
  const body = JSON.parse(e.postData.contents);

  if (!body.room || !body.checkin || !body.checkout) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: 'Missing room or dates' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  sheet.appendRow([
    body.room,
    body.checkin,
    body.checkout,
    body.name || '',
    body.email || '',
    new Date()
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatDate_(value) {
  if (!value) return '';
  if (typeof value === 'string') return value;
  const y = value.getFullYear();
  const m = String(value.getMonth() + 1).padStart(2, '0');
  const d = String(value.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}
