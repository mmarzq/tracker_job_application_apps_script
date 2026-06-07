const HEADERS = [
  'Company', 'Position / Job Title', 'Location', 'Job URL', 'Folder', 'Category',
  'Job Type', 'Stage', 'Date Applied', 'Job Source', 'Salary Range',
  'Lebenslauf', 'Anschreiben', 'Dokumente', 'Bewerbungsportal (Ja/Nein)',
  'Notes', 'Interview Date', 'Next Step', 'Files & media'
];
const BEWERBUNG_FOLDER_ID    = '1b_Olkq1J-5r4Rq_ynooEzt_UCavNWC3k';
const ANSCHREIBEN_VORLAGE_ID = '1fnmRhmMF_BUY7aIRbHfVqa6Vc6wy8-ve';
const LEBENSLAUF_VORLAGE_ID  = '1CxnqhxYDDEZX6p4snwW9g5yxLILoErtl';

const MULTISELECT = ['Category', 'Job Source', 'Dokumente'];
const DROPDOWN    = ['Job Type', 'Stage', 'Lebenslauf', 'Anschreiben'];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Bewerbungen')
    .addItem('➕ New (leere Zeile)', 'insertEmptyRow')
    .addItem('📝 New (Formular ausfüllen)', 'showForm')
    .addItem('📊 Dashboard', 'showDashboard')
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

function addRow(data, fileData) {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.insertRowBefore(2);

  const folder = DriveApp.getFolderById(BEWERBUNG_FOLDER_ID)
    .createFolder(`${data['Company']} - ${data['Location']}`);
  DriveApp.getFileById(ANSCHREIBEN_VORLAGE_ID).makeCopy(folder);
  DriveApp.getFileById(LEBENSLAUF_VORLAGE_ID).makeCopy(folder);

  let fileUrl = '';
  if (fileData) {
    const blob = Utilities.newBlob(Utilities.base64Decode(fileData.base64), fileData.mimeType, fileData.name);
    fileUrl = folder.createFile(blob).getUrl();
  }

  const row = HEADERS.map(h => {
    if (h === 'Bewerbungsportal (Ja/Nein)') return data[h] === 'TRUE';
    if (h === 'Folder')        return folder.getUrl();
    if (h === 'Files & media') return fileUrl;
    return data[h] || '';
  });
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
