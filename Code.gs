const SS = SpreadsheetApp.getActiveSpreadsheet();

const SHEET_PASSATIONS = 'Passations';
const SHEET_ACCIDENTS = 'Accidents';
const SHEET_RAPPORT = 'Rapport';

const FLEET_COLUMNS = ['Available', 'Charging', 'Discharged', 'Need Investigation', 'Maintenance', 'Stolen', 'Not Ready', 'Rebalancing'];

const PASSATION_HEADERS = [
  'ID', 'Date', 'Créneau', 'Ville', 'Astreinte', 'Créé par',
  ...FLEET_COLUMNS,
  'ID Verification', 'ID Verification Comment',
  'Damage', 'Damage Comment', 'Damages Open',
  'TripsOk', 'Trips Detail', 'Message'
];

const ACCIDENT_HEADERS = ['ID_Passation', 'Date', 'Nom', 'ID utilisateur', 'Lieu', 'Description', 'Ramener au garage'];

function doGet(e) {
  const action = e.parameter.action;
  let result;
  if (action === 'getPassations') {
    result = { passations: readSheet_(SHEET_PASSATIONS) };
  } else if (action === 'getAccidents') {
    result = { accidents: readSheet_(SHEET_ACCIDENTS) };
  } else {
    result = { error: 'Unknown action: ' + action };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const action = body.action;
  if (action === 'addPassation') {
    addPassation_(body);
  } else if (action === 'exportRapport') {
    exportRapport_(body.rows || []);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}

// ─── SHEET HELPERS ───

function getSheet_(name, headers) {
  let sheet = SS.getSheetByName(name);
  if (!sheet) {
    sheet = SS.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readSheet_(name) {
  const sheet = SS.getSheetByName(name);
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1)
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        const val = row[i];
        obj[h] = (val instanceof Date) ? val.toISOString() : val;
      });
      return obj;
    });
}

// ─── WRITE ACTIONS ───

function addPassation_(p) {
  const sheet = getSheet_(SHEET_PASSATIONS, PASSATION_HEADERS);
  const id = Utilities.getUuid();

  const row = [
    id,
    p.date || new Date().toISOString(),
    p.creneau || '',
    p.ville || '',
    p.astreinte || '',
    p.creePar || '',
    p.available ?? 0,
    p.charging ?? 0,
    p.discharged ?? 0,
    p.needInvestigation ?? 0,
    p.maintenance ?? 0,
    p.stolen ?? 0,
    p.notReady ?? 0,
    p.rebalancing ?? 0,
    p.idVerification || '',
    p.idVerificationComment || '',
    p.damage || '',
    p.damageComment || '',
    p.damagesOpen ?? '',
    p.tripsOk || '',
    p.tripsDetail || '',
    p.message || ''
  ];
  sheet.appendRow(row);

  if (p.accidents && p.accidents.length) {
    const accSheet = getSheet_(SHEET_ACCIDENTS, ACCIDENT_HEADERS);
    p.accidents.forEach(a => {
      if (!a.nom) return;
      accSheet.appendRow([id, p.date || new Date().toISOString(), a.nom, a.userid || '', a.lieu || '', a.desc || '', a.garage ? 'Oui' : 'Non']);
    });
  }
}

function exportRapport_(rows) {
  if (!rows.length) return;
  const headers = Object.keys(rows[0]);
  const sheet = getSheet_(SHEET_RAPPORT, headers);
  rows.forEach(r => {
    sheet.appendRow(headers.map(h => r[h] ?? ''));
  });
}
