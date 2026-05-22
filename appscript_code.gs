// ━━━ AIECO KALKULATOR — Google Apps Script ━━━
// Instrukcja wdrożenia:
// 1. Wejdź na script.google.com
// 2. Otwórz projekt powiązany z arkuszem
// 3. Zastąp cały kod tym plikiem
// 4. Kliknij: Wdróż → Nowe wdrożenie
//    - Typ: Aplikacja internetowa
//    - Wykonaj jako: Ja
//    - Kto ma dostęp: Wszyscy (anonimowi)
// 5. Skopiuj URL wdrożenia → wklej do index.html jako GOOGLE_SHEETS_API_URL

const SHEET_NAME = 'Leady';

const HEADERS = [
  'Data i godzina',
  'Email',
  'Telefon',
  'Typ klienta',
  'Obecna moc PV (kWp)',
  'Obecny falownik (kW)',
  'Typ falownika',
  'Docelowa moc PV (kWp)',
  'Nowa moc PV (kWp)',
  'Kierunek PV',
  'Moc przyłącza (kW)',
  'Operator',
  'Taryfa obecna',
  'Zasady rozliczeń',
  'Tryb wprowadzania',
  'Zużycie roczne (kWh)',
  'Rachunek mies. (zł)',
  'Zużycie pompa (kWh)',
  'Zużycie EV (kWh)',
  'Profil zużycia',
  'Strategia zimowa',
  'Backup awaryjny',
  'Taryfa docelowa',
  'Dofinansowanie (zł)',
  'Stawka podatkowa',
  'Źródło'
];

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'AiEco API działa ✅' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    // Utwórz arkusz z nagłówkami jeśli nie istnieje
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADERS);
      const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#10B981');
      headerRange.setFontColor('#ffffff');
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, HEADERS.length);
    }

    // Parsuj dane JSON z body
    let data = {};
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch(parseErr) {
        // fallback: spróbuj z params URL
        data = e.parameter || {};
      }
    } else if (e.parameter) {
      data = e.parameter;
    }

    // Zapisz wiersz w kolejności zgodnej z HEADERS
    sheet.appendRow([
      data.timestamp           || new Date().toLocaleString('pl-PL'),
      data.email               || '',
      data.phone               || '',
      data.typ_klienta         || '',
      data.obecna_moc_pv_kwp   || '',
      data.obecny_falownik_kw  || '',
      data.typ_falownika       || '',
      data.docelowa_moc_pv_kwp || '',
      data.nowa_moc_pv_kwp     || '',
      data.kierunek_pv         || '',
      data.moc_przylacza_kw    || '',
      data.operator            || '',
      data.taryfa_obecna       || '',
      data.zasady_rozliczen    || '',
      data.tryb_wprowadzania   || '',
      data.zuzycie_roczne_kwh  || '',
      data.rachunek_miesieczny_zl || '',
      data.zuzycie_pompa_kwh   || '',
      data.zuzycie_ev_kwh      || '',
      data.profil_zuzycia      || '',
      data.strategia_zimowa    || '',
      data.backup_awaryjny     || '',
      data.taryfa_docelowa     || '',
      data.dofinansowanie_zl   || '',
      data.stawka_podatkowa    || '',
      data.source              || 'kalkulator-aieco'
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', message: 'Lead zapisany ✅' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
