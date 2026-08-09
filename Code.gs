const SS = SpreadsheetApp.getActiveSpreadsheet();

const SHEET_PASSATIONS = 'Passations';
const SHEET_ACCIDENTS = 'Accidents';
const SHEET_RAPPORT = 'Rapport';
const SHEET_IOT = 'IoT Signal';
const SHEET_IOT_JOURNAL = 'IoT Journal';

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

const IOT_HEADERS = [
  'Vehicle Number', 'Group', 'Vehicle Status', 'Vehicle Battery', 'IoT Last Update',
  'Hours Since Signal', 'Latitude', 'Longitude', 'Last Ride',
  'Position Check', 'Contact Joint', 'Lieu Indique', 'Resolved',
  'First Seen', 'Last Seen', 'Created By'
];
const IOT_JOURNAL_HEADERS = ['Timestamp', 'Vehicle Number', 'Action', 'Detail', 'By'];

function doGet(e) {
  const action = e.parameter.action;
  let result;
  if (action === 'getPassations') {
    result = { passations: readSheet_(SHEET_PASSATIONS) };
  } else if (action === 'getAccidents') {
    result = { accidents: readSheet_(SHEET_ACCIDENTS) };
  } else if (action === 'getIoT') {
    result = { entries: readSheet_(SHEET_IOT), journal: readSheet_(SHEET_IOT_JOURNAL) };
  } else {
    result = { error: 'Unknown action: ' + action };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const action = body.action;
  let result = { ok: true };
  if (action === 'addPassation') {
    addPassation_(body);
  } else if (action === 'exportRapport') {
    exportRapport_(body.rows || []);
  } else if (action === 'importIoT') {
    result = Object.assign({ ok: true }, importIoT_(body));
  } else if (action === 'updateIoT') {
    result = updateIoT_(body);
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
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

// ─── IoT SIGNAL TRACKING ───
//
// Chaque opérateur importe son propre export Atom (Vehicles > List view,
// filtré "Last IoT signal ≥ 60min") au moment de sa passation. L'export est
// un instantané complet de "tout ce qui n'a pas de signal en ce moment" —
// donc à chaque import :
//   1. un véhicule ouvert (non résolu) qui n'apparaît PLUS dans l'export
//      a retrouvé du signal → résolu automatiquement
//   2. un véhicule déjà ouvert qui réapparaît → ses champs Atom sont
//      rafraîchis (statut, batterie...) mais son suivi opérateur (position
//      vérifiée, contact, résolu) est préservé, pas écrasé
//   3. un véhicule totalement nouveau (ou déjà résolu précédemment) →
//      nouvelle entrée ouverte
// Ça permet à plusieurs opérateurs d'importer le même export à quelques
// heures d'intervalle sans dupliquer une entrée : la clé de dédoublonnage
// est Vehicle Number, et seules les entrées non résolues sont candidates au
// rapprochement.

function readIotSheetIndexed_(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const col = {};
  headers.forEach((h, i) => { col[h] = i + 1; }); // 1-based colonne
  const lastRow = sheet.getLastRow();
  const data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  return { col, data };
}

function importIoT_(body) {
  const sheet = getSheet_(SHEET_IOT, IOT_HEADERS);
  const journalSheet = getSheet_(SHEET_IOT_JOURNAL, IOT_JOURNAL_HEADERS);
  const now = new Date().toISOString();
  const importedBy = body.importedBy || '';
  const incoming = (body.rows || []).filter(r => r.vehicleNumber);
  const incomingKeys = new Set(incoming.map(r => r.vehicleNumber));

  const { col, data } = readIotSheetIndexed_(sheet);

  // Index des entrées actuellement ouvertes (non résolues) par Vehicle Number
  const openRowByVn = {};
  data.forEach((row, i) => {
    const vn = row[col['Vehicle Number'] - 1];
    const resolved = row[col['Resolved'] - 1];
    if (vn && resolved !== 'oui') openRowByVn[vn] = i + 2; // numéro de ligne réel (1-based + en-tête)
  });

  // 1) Signal revenu : entrées ouvertes absentes de ce nouvel export
  let autoResolved = 0;
  Object.keys(openRowByVn).forEach(vn => {
    if (!incomingKeys.has(vn)) {
      const rowNum = openRowByVn[vn];
      sheet.getRange(rowNum, col['Resolved']).setValue('oui');
      sheet.getRange(rowNum, col['Last Seen']).setValue(now);
      journalSheet.appendRow([now, vn, 'resolved', "Signal IoT revenu (résolu automatiquement à l'import)", 'Système']);
      autoResolved++;
    }
  });

  // 2) Mise à jour des entrées existantes / création des nouvelles
  let updated = 0, added = 0;
  incoming.forEach(r => {
    const vn = r.vehicleNumber;
    if (openRowByVn[vn]) {
      const rowNum = openRowByVn[vn];
      sheet.getRange(rowNum, col['Group']).setValue(r.group);
      sheet.getRange(rowNum, col['Vehicle Status']).setValue(r.status);
      sheet.getRange(rowNum, col['Vehicle Battery']).setValue(r.battery);
      sheet.getRange(rowNum, col['IoT Last Update']).setValue(r.iotLastUpdate);
      sheet.getRange(rowNum, col['Hours Since Signal']).setValue(r.hoursSinceSignal);
      sheet.getRange(rowNum, col['Latitude']).setValue(r.lat);
      sheet.getRange(rowNum, col['Longitude']).setValue(r.lng);
      sheet.getRange(rowNum, col['Last Ride']).setValue(r.lastRide);
      sheet.getRange(rowNum, col['Last Seen']).setValue(now);
      updated++;
    } else {
      const values = {
        'Vehicle Number': vn, 'Group': r.group, 'Vehicle Status': r.status,
        'Vehicle Battery': r.battery, 'IoT Last Update': r.iotLastUpdate,
        'Hours Since Signal': r.hoursSinceSignal, 'Latitude': r.lat, 'Longitude': r.lng,
        'Last Ride': r.lastRide, 'Position Check': 'non_verifie', 'Contact Joint': '',
        'Lieu Indique': '', 'Resolved': 'non', 'First Seen': now, 'Last Seen': now,
        'Created By': importedBy
      };
      appendRowByHeader_(sheet, values);
      journalSheet.appendRow([now, vn, 'import', 'Signalé (' + r.group + ', ' + r.status + ')', importedBy]);
      added++;
    }
  });

  return { added: added, updated: updated, autoResolved: autoResolved };
}

function updateIoT_(body) {
  const sheet = getSheet_(SHEET_IOT, IOT_HEADERS);
  const journalSheet = getSheet_(SHEET_IOT_JOURNAL, IOT_JOURNAL_HEADERS);
  const now = new Date().toISOString();
  const vn = body.vehicleNumber;
  const by = body.by || '';
  const type = body.type; // 'position' | 'contact' | 'resolved'

  const { col, data } = readIotSheetIndexed_(sheet);
  let rowNum = null;
  for (let i = 0; i < data.length; i++) {
    if (data[i][col['Vehicle Number'] - 1] === vn && data[i][col['Resolved'] - 1] !== 'oui') {
      rowNum = i + 2;
      break;
    }
  }
  if (!rowNum) return { ok: false, error: 'Véhicule introuvable ou déjà résolu' };

  if (type === 'position') {
    sheet.getRange(rowNum, col['Position Check']).setValue(body.value);
    const label = body.value === 'correspond' ? 'Correspond' : body.value === 'ne_correspond_pas' ? 'Ne correspond pas' : 'Non vérifié';
    journalSheet.appendRow([now, vn, 'position_check', label, by]);
  } else if (type === 'contact') {
    sheet.getRange(rowNum, col['Contact Joint']).setValue(body.contactJoint);
    sheet.getRange(rowNum, col['Lieu Indique']).setValue(body.lieuIndique || '');
    const detail = body.contactJoint === 'oui'
      ? 'Contact joint — lieu indiqué : ' + (body.lieuIndique || '—')
      : 'Contact non joint';
    journalSheet.appendRow([now, vn, 'contact', detail, by]);
  } else if (type === 'resolved') {
    sheet.getRange(rowNum, col['Resolved']).setValue(body.resolved ? 'oui' : 'non');
    journalSheet.appendRow([now, vn, body.resolved ? 'resolved' : 'reopened', body.resolved ? 'Marqué résolu manuellement' : 'Ré-ouvert', by]);
  } else {
    return { ok: false, error: 'Type inconnu: ' + type };
  }
  sheet.getRange(rowNum, col['Last Seen']).setValue(now);
  return { ok: true };
}
