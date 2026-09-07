const SHEET_NAME = 'Registrations';

function doGet() {
  return ContentService
    .createTextOutput('ACRESERVE 2027 registration endpoint is active.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    const sheet = getRegistrationSheet_();
    const p = e && e.parameter ? e.parameter : {};

    sheet.appendRow([
      new Date(),
      p.fullName || '',
      p.email || '',
      p.mobile || '',
      p.organization || '',
      p.designation || '',
      p.city || '',
      p.interest || '',
      p.message || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: String(error) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getRegistrationSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Submitted At',
      'Full Name',
      'Email',
      'Mobile',
      'Organization / College',
      'Designation / Role',
      'City',
      'Area of Interest',
      'Message / Special Requirement'
    ]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}
