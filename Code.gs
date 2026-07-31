const SS = SpreadsheetApp.getActiveSpreadsheet();

const SHEET_PASSATIONS = 'Passations';
const SHEET_ACCIDENTS = 'Accidents';
const SHEET_RAPPORT = 'Rapport';

const FLEET_COLUMNS = [
  'Available', 'Charging', 'Discharged', 'Need Investigation', 'Maintenance',
  'Stolen', 'Not Ready', 'Rebalancing', 'In Use', 'Transportation', 'Storage'
];

const PASSATION_HEADERS = [
  'ID', 'Date', 'Créneau', 'Ville', 'Astreinte', 'Créé par',
  ...FLEET_COLUMNS,
  'ID Verification', 'ID Verification Comment',
  'Damage', 'Damage Comment', 'Damages Open',
  'Damages Reported Today', 'Damages Reported Yesterday',
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

// Crée l'onglet avec les en-têtes s'il n'existe pas. S'il existe déjà,
// ajoute en fin de ligne les en-têtes qui manqueraient (ex: nouveaux champs
// ajoutés côté front après la création initiale de l'onglet), sans jamais
// réordonner ou supprimer les colonnes existantes — pour ne pas décaler les
// données déjà écrites.
function getSheet_(name, headers) {
  let sheet = SS.getSheetByName(name);
  if (!sheet) {
    sheet = SS.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  } else {
    ensureHeaders_(sheet, headers);
  }
  return sheet;
}

function ensureHeaders_(sheet, headers) {
  const lastCol = sheet.getLastColumn();
  const existing = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  const missing = headers.filter(h => existing.indexOf(h) === -1);
  if (missing.length) {
    sheet.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
  }
}

// Écrit `values` (objet { "Nom de colonne": valeur }) dans une nouvelle ligne,
// en alignant chaque valeur sur la colonne correspondante d'après l'en-tête
// réel de la feuille (peu importe l'ordre) plutôt que sur une position fixe.
function appendRowByHeader_(sheet, values) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const row = headers.map(h => (h in values) ? values[h] : '');
  sheet.appendRow(row);
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

  const values = {
    'ID': id,
    'Date': p.date || new Date().toISOString(),
    'Créneau': p.creneau || '',
    'Ville': p.ville || '',
    'Astreinte': p.astreinte || '',
    'Créé par': p.creePar || '',
    'Available': p.available ?? 0,
    'Charging': p.charging ?? 0,
    'Discharged': p.discharged ?? 0,
    'Need Investigation': p.needInvestigation ?? 0,
    'Maintenance': p.maintenance ?? 0,
    'Stolen': p.stolen ?? 0,
    'Not Ready': p.notReady ?? 0,
    'Rebalancing': p.rebalancing ?? 0,
    'In Use': p.inUse ?? 0,
    'Transportation': p.transportation ?? 0,
    'Storage': p.storage ?? 0,
    'ID Verification': p.idVerification || '',
    'ID Verification Comment': p.idVerificationComment || '',
    'Damage': p.damage || '',
    'Damage Comment': p.damageComment || '',
    'Damages Open': p.damagesOpen ?? '',
    'Damages Reported Today': p.damagesReportedToday ?? '',
    'Damages Reported Yesterday': p.damagesReportedYesterday ?? '',
    'TripsOk': p.tripsOk || '',
    'Trips Detail': p.tripsDetail || '',
    'Message': p.message || ''
  };
  appendRowByHeader_(sheet, values);

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
