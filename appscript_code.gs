// ━━━ AIECO KALKULATOR — Google Apps Script ━━━
// Wklej ten kod do script.google.com i opublikuj jako Web App:
// Wykonaj jako: Ja
// Kto ma dostęp: Wszyscy (anonimowi)

const SHEET_NAME = 'Leady'; // Zmień jeśli arkusz ma inną nazwę

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'AiEco API działa' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    
    // Jeśli arkusz nie istnieje - utwórz go
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      // Dodaj nagłówki
      sheet.appendRow([
        'Data i godzina',
        'Email',
        'Telefon',
        'Typ instalacji',
        'Moc PV (kWp)',
        'Zużycie (kWh)',
        'Backup',
        'Źródło'
      ]);
      // Styl nagłówków
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#10B981').setFontColor('#ffffff');
    }

    // Parsuj dane z POST
    let data = {};
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    }

    // Dodaj wiersz z danymi
    sheet.appendRow([
      data.timestamp || new Date().toLocaleString('pl-PL'),
      data.email || '',
      data.phone || '',
      data.type || '',
      data.pvPower || '',
      data.consumption || '',
      data.backup || '',
      data.source || 'kalkulator-aieco'
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', message: 'Lead zapisany' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
