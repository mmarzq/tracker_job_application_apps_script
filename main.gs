// Spalten-Header in genau dieser Reihenfolge
const HEADERS = [
  'Company', 'Position / Job Title', 'Location', 'Job URL', 'Category',
  'Job Type', 'Stage', 'Date Applied', 'Job Source', 'Salary Range',
  'Lebenslauf', 'Anschreiben', 'Dokumente', 'Bewerbungsportal (Ja/Nein)',
  'Notes', 'Interview Date', 'Next Step', 'Files & media'
];

// Menü beim Öffnen erstellen
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Bewerbungen')
    .addItem('➕ New (leere Zeile)', 'insertEmptyRow')
    .addItem('📝 New (Formular ausfüllen)', 'showForm')
    .addToUi();
}

// Variante 1: Leere Zeile in Zeile 2 einfügen
function insertEmptyRow() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.insertRowBefore(2);
  sheet.getRange(2, 1).activate();
}

// Variante 2: Formular-Dialog anzeigen
function showForm() {
  const html = HtmlService.createHtmlOutputFromFile('Form')
    .setWidth(420)
    .setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Neue Bewerbung');
}

// Vom Formular aufgerufen: Daten in Zeile 2 schreiben
function addRow(data) {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.insertRowBefore(2);
  const row = HEADERS.map(h => data[h] || '');
  sheet.getRange(2, 1, 1, row.length).setValues([row]);
  return true;
}