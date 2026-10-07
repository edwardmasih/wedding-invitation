function doPost(e) {
  const payload = JSON.parse(e.postData.contents || '{}');
  const spreadsheet = SpreadsheetApp.openById('PASTE_YOUR_SHEET_ID_HERE');
  const sheet = spreadsheet.getSheetByName('RSVP') || spreadsheet.insertSheet('RSVP');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Name',
      'Attending',
      'Note for the Couple',
      'Question 1',
      'Answer 1',
      'Question 2',
      'Answer 2',
      'Question 3',
      'Answer 3'
    ]);
  }

  const answers = Array.isArray(payload.answers) ? payload.answers : [];
  sheet.appendRow([
    payload.name || '',
    payload.attending || '',
    payload.note || '',
    answers[0] ? answers[0].question : '',
    answers[0] ? answers[0].answer : '',
    answers[1] ? answers[1].question : '',
    answers[1] ? answers[1].answer : '',
    answers[2] ? answers[2].question : '',
    answers[2] ? answers[2].answer : ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}