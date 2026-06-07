const HEADERS = [
  'Company', 'Position / Job Title', 'Location', 'Job URL', 'Category',
  'Job Type', 'Stage', 'Date Applied', 'Job Source', 'Salary Range',
  'Lebenslauf', 'Anschreiben', 'Dokumente', 'Bewerbungsportal (Ja/Nein)',
  'Notes', 'Interview Date', 'Next Step', 'Files & media'
];
const MULTISELECT = ['Category', 'Job Source', 'Dokumente'];
const DROPDOWN    = ['Job Type', 'Stage', 'Lebenslauf', 'Anschreiben'];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Bewerbungen')
    .addItem('➕ New (leere Zeile)', 'insertEmptyRow')
    .addItem('📝 New (Formular ausfüllen)', 'showForm')
    .addToUi();
}

function insertEmptyRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.insertRowBefore(2);
  sheet.getRange(2, 1).activate();
}

function showForm() {
  const tmpl = HtmlService.createTemplateFromFile('Form');
  tmpl.configJson = JSON.stringify(getFormConfig()).replace(/<\//g, '<\\/');
  const html = tmpl.evaluate().setWidth(420).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Neue Bewerbung');
}

function addRow(data) {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.insertRowBefore(2);
  const row = HEADERS.map(h =>
    h === 'Bewerbungsportal (Ja/Nein)' ? data[h] === 'TRUE' : (data[h] || '')
  );
  sheet.getRange(2, 1, 1, row.length).setValues([row]);
  return true;
}

function getFormConfig() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const config = {};

  HEADERS.forEach((header, i) => {
    const isMulti = MULTISELECT.includes(header);
    if (!isMulti && !DROPDOWN.includes(header)) return;

    const seen = new Set();
    data.forEach(row => {
      const cell = String(row[i] || '').trim();
      if (!cell) return;
      if (isMulti) cell.split(',').forEach(v => seen.add(v.trim()));
      else seen.add(cell);
    });

    config[header] = { type: isMulti ? 'multiselect' : 'dropdown', values: [...seen].sort() };
  });

  return config;
}
