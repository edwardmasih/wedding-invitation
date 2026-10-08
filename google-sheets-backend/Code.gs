const ADMIN_PASSWORD = '2806';

function doPost(e) {
  const payload = JSON.parse(e.postData.contents || '{}');
  const spreadsheet = SpreadsheetApp.openById('1a3EBD4zQT4S6uZwZpSH7DaY6vhq3SmlQ6aKIptYM8M4');
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

function doGet(e) {
  const callback = e && e.parameter && e.parameter.callback ? String(e.parameter.callback) : '';
  if (!e || !e.parameter || e.parameter.password !== ADMIN_PASSWORD) {
    return respond({ ok: false, error: 'Invalid password.' }, callback);
  }

  const spreadsheet = SpreadsheetApp.openById('1a3EBD4zQT4S6uZwZpSH7DaY6vhq3SmlQ6aKIptYM8M4');
  const sheet = spreadsheet.getSheetByName('RSVP');
  if (!sheet) {
    return respond({ ok: true, headers: [], rows: [] }, callback);
  }

  const values = sheet.getDataRange().getDisplayValues();
  const headers = values.length ? values[0] : [];
  const rows = values.length > 1 ? values.slice(1) : [];
  return respond({ ok: true, headers: headers, rows: rows }, callback);
}

function respond(payload, callback) {
  const json = JSON.stringify(payload);
  if (callback) {
    return ContentService
      .createTextOutput(`${callback}(${json});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}