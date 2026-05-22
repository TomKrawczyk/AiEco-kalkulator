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
  // Kontakt
  'Data i godzina',
  'Email',
  'Telefon',

  // Krok 1
  'Typ klienta',

  // Krok 2 - instalacja
  'Obecna moc PV (kWp)',
  'Obecny falownik (kW)',
  'Typ falownika',
  'Chce rozbudować PV',
  'Docelowa moc PV (kWp)',
  'Nowa moc PV (kWp)',
  'Kierunek PV',
  'Moc przyłącza (kW)',
  'Operator',
  'Taryfa obecna',
  'Zasady rozliczeń',

  // Krok 3 - zużycie
  'Tryb wprowadzania',
  'Zużycie roczne (kWh)',
  'Rachunek mies. (zł)',
  'Ma pompę ciepła',
  'Zużycie pompy (kWh)',
  'Ma auto EV',
  'Zużycie EV (kWh)',
  'Profil zużycia',
  'Strategia zimowa',
  'Chce backup awaryjny',

  // Krok 4 - finanse
  'Taryfa docelowa',
  'Chce dofinansowanie',
  'Dofinansowanie (zł)',
  'Jest podatnikiem',
  'Stawka podatkowa',

  // Meta
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

    let data = {};
    if (e.postData && e.postData.contents) {
      try { data = JSON.parse(e.postData.contents); }
      catch(err) { data = e.parameter || {}; }
    } else if (e.parameter) {
      data = e.parameter;
    }

    sheet.appendRow([
      data.timestamp              || new Date().toLocaleString('pl-PL'),
      data.email                  || '',
      data.phone                  || '',

      data.typ_klienta            || '',

      data.obecna_moc_pv_kwp      || '',
      data.obecny_falownik_kw     || '',
      data.typ_falownika          || '',
      data.rozbudowa_pv           || '',
      data.docelowa_moc_pv_kwp    || '',
      data.nowa_moc_pv_kwp        || '',
      data.kierunek_pv            || '',
      data.moc_przylacza_kw       || '',
      data.operator               || '',
      data.taryfa_obecna          || '',
      data.zasady_rozliczen       || '',

      data.tryb_wprowadzania      || '',
      data.zuzycie_roczne_kwh     || '',
      data.rachunek_miesieczny_zl || '',
      data.ma_pompe_ciepla        || '',
      data.zuzycie_pompa_kwh      || '',
      data.ma_auto_ev             || '',
      data.zuzycie_ev_kwh         || '',
      data.profil_zuzycia         || '',
      data.strategia_zimowa       || '',
      data.backup_awaryjny        || '',

      data.taryfa_docelowa        || '',
      data.chce_dofinansowanie    || '',
      data.dofinansowanie_zl      || '',
      data.jest_podatnikiem       || '',
      data.stawka_podatkowa       || '',

      data.source                 || 'kalkulator-aieco'
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
