function doGet() {
  return HtmlService.createHtmlOutputFromFile('dashboard')
    .setTitle('Dashboard – Bewerbungen')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

const DASHBOARD_URL = 'https://script.google.com/macros/s/AKfycbzk2XI193Qa44aTZOlCe2mE7LjKjjjOd9l-Xa4zIkA/dev';

function showDashboard() {
  const html = HtmlService.createHtmlOutput(
    `<p style="font-family:Arial;font-size:14px">
       <a href="${DASHBOARD_URL}" target="_blank" onclick="google.script.host.close()">
         📊 Dashboard öffnen
       </a>
     </p>`
  );
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard');
}

function getDashboardData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { stats: { total: 0, byStage: {}, thisWeek: 0, thisMonth: 0 }, upcoming: [], rows: [] };

  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  const tz     = Session.getScriptTimeZone();
  const fmtDate = d => (d instanceof Date && !isNaN(d))
    ? Utilities.formatDate(d, tz, 'dd.MM.yyyy') : '';

  const now        = new Date();
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  const weekAgo    = new Date(now - 7 * 24 * 60 * 60 * 1000);

  const byStage = {};
  let thisWeek = 0, thisMonth = 0;
  const upcoming = [];
  const rows = [];

  values.forEach(row => {
    if (!row[0]) return;
    const obj = {};
    HEADERS.forEach((h, i) => {
      obj[h] = (row[i] instanceof Date) ? fmtDate(row[i]) : String(row[i] || '');
    });
    rows.push(obj);

    const stage = obj['Stage'] || '–';
    byStage[stage] = (byStage[stage] || 0) + 1;

    const dateApplied = row[HEADERS.indexOf('Date Applied')];
    if (dateApplied instanceof Date && !isNaN(dateApplied)) {
      if (dateApplied >= weekAgo)    thisWeek++;
      if (dateApplied >= monthStart) thisMonth++;
    }

    const intDate = row[HEADERS.indexOf('Interview Date')];
    if (intDate instanceof Date && !isNaN(intDate) && intDate > now)
      upcoming.push({ rawDate: intDate, company: obj['Company'], position: obj['Position / Job Title'] });
  });

  upcoming.sort((a, b) => a.rawDate - b.rawDate);
  const upcomingFmt = upcoming.map(u => ({ date: fmtDate(u.rawDate), company: u.company, position: u.position }));

  return { stats: { total: rows.length, byStage, thisWeek, thisMonth }, upcoming: upcomingFmt, rows };
}
