const ADMIN_PASSWORD = '2806';
const SHEET_ID = '1a3EBD4zQT4S6uZwZpSH7DaY6vhq3SmlQ6aKIptYM8M4';
const SHEET_NAME = 'RSVP';
const HEADERS = [
  'Entry ID',
  'Source',
  'Name',
  'Attending',
  'Note for the Couple',
  'Question 1',
  'Answer 1',
  'Question 2',
  'Answer 2',
  'Question 3',
  'Answer 3'
];

function doPost(e) {
  const payload = JSON.parse(e.postData.contents || '{}');
  const sheet = getSheet();

  const answers = Array.isArray(payload.answers) ? payload.answers : [];
  sheet.appendRow([
    Utilities.getUuid(),
    payload.source || '',
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

  const action = e.parameter.action || 'admin';
  if (action === 'delete') {
    return handleDelete(e.parameter, callback);
  }

  const sheet = getSheet();
  if (!sheet) {
    return respond({ ok: true, headers: [], rows: [] }, callback);
  }

  const values = sheet.getDataRange().getDisplayValues();
  const headers = values.length ? values[0] : [];
  const rows = values.length > 1 ? values.slice(1) : [];
  const sourceIndex = headers.indexOf('Source');
  const source = e.parameter.source || '';
  const filteredRows = source && sourceIndex !== -1
    ? rows.filter(function(row) { return row[sourceIndex] === source; })
    : rows;
  return respond({ ok: true, headers: headers, rows: filteredRows }, callback);
}

function handleDelete(params, callback) {
  const entryId = params.entryId || '';
  const source = params.source || '';
  if (!entryId) {
    return respond({ ok: false, error: 'Entry ID is required.' }, callback);
  }

  const sheet = getSheet();
  const values = sheet.getDataRange().getDisplayValues();
  if (!values.length) {
    return respond({ ok: false, error: 'No RSVP data found.' }, callback);
  }

  const headers = values[0];
  const idIndex = headers.indexOf('Entry ID');
  const sourceIndex = headers.indexOf('Source');
  if (idIndex === -1) {
    return respond({ ok: false, error: 'Entry ID column is missing.' }, callback);
  }

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row[idIndex] === entryId && (!source || sourceIndex === -1 || row[sourceIndex] === source)) {
      sheet.deleteRow(i + 1);
      return respond({ ok: true }, callback);
    }
  }

  return respond({ ok: false, error: 'RSVP entry not found.' }, callback);
}

function getSheet() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
  ensureHeaders(sheet);
  return sheet;
}

function ensureHeaders(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), HEADERS.length);
  const currentHeaders = lastColumn ? sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0] : [];
  const normalizedHeaders = HEADERS.slice();
  var changed = false;

  HEADERS.forEach(function(header, index) {
    if (currentHeaders[index] !== header) {
      changed = true;
    }
  });

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    return;
  }

  if (changed) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([normalizedHeaders]);
  }
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