const SS = SpreadsheetApp.getActiveSpreadsheet();

const SHEET_PASSATIONS = 'Passations';
const SHEET_IOT = 'IoT Signal';
const SHEET_IOT_JOURNAL = 'IoT Journal';
const SHEET_IOT_ARCHIVE = 'IoT Archive';

// Les anciennes colonnes (statuts flotte, damage...) restent dans la feuille
// pour l'historique : ensureHeaders_ ne supprime jamais de colonne, il ajoute
// seulement celles qui manquent (ici les champs "Offline ...").
const PASSATION_HEADERS = [
  'ID', 'Date', 'Créneau', 'Ville', 'Astreinte', 'Créé par',
  'ID Verification', 'ID Verification Comment',
  'TripsOk', 'Trips Detail', 'Message',
  'Offline Available', 'Offline Total', 'Offline Liste'
];

// 'Last Seen' vs 'Last Import' : 'Last Seen' est touché par n'importe quelle
// interaction sur l'entrée (import, mais aussi vérification de position,
// contact joint, résolution manuelle depuis l'onglet Offline — voir
// updateIoTLocked_). 'Last Import' n'est écrit QUE par importIoTLocked_,
// jamais par une action manuelle. Le front-end s'en sert pour ne retenir
// que les véhicules confirmés par le tout dernier export importé (voir
// currentOpenIotEntries_ côté index.html).
const IOT_HEADERS = [
  'Vehicle Number', 'Group', 'Vehicle Status', 'Vehicle Battery', 'IoT Last Update',
  'Hours Since Signal', 'Latitude', 'Longitude', 'Address', 'Last Ride',
  'Position Check', 'Contact Joint', 'Lieu Indique', 'Resolved',
  'First Seen', 'Last Seen', 'Last Import', 'Created By',
  // Adresse saisie à la main pendant la passation (prime sur l'adresse GPS
  // dans le message). Jamais écrasée par un import.
  'Adresse Reelle'
];
const IOT_JOURNAL_HEADERS = ['Timestamp', 'Vehicle Number', 'Action', 'Detail', 'By'];

// Le journal n'est jamais purgé dans la feuille, mais l'onglet Offline n'a
// besoin que de l'historique récent de chaque véhicule.
const IOT_JOURNAL_DAYS = 30;

function doGet(e) {
  let result;
  try {
    const action = e.parameter.action;
    if (action === 'getAll') {
      // Un seul aller-retour pour tout ce dont l'app a besoin au démarrage :
      // chaque appel Apps Script coûte ~2-3s fixes, quelle que soit la taille
      // de la réponse.
      result = { passations: readSheet_(SHEET_PASSATIONS), entries: readSheet_(SHEET_IOT) };
      if (e.parameter.withJournal === '1') result.journal = readRecentJournal_();
    } else if (action === 'getPassations') {
      result = { passations: readSheet_(SHEET_PASSATIONS) };
    } else if (action === 'getIoT') {
      result = { entries: readSheet_(SHEET_IOT) };
      if (e.parameter.withJournal === '1') result.journal = readRecentJournal_();
    } else {
      result = { ok: false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result;
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    if (action === 'addPassation') {
      result = addPassation_(body);
    } else if (action === 'importIoT') {
      result = importIoT_(body);
    } else if (action === 'updateIoT') {
      result = updateIoT_(body);
    } else {
      result = { ok: false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    // Renvoyé en JSON plutôt que la page d'erreur HTML d'Apps Script : le
    // front affiche alors la vraie cause au lieu d'un échec opaque.
    result = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ─── SHEET HELPERS ───

// Crée l'onglet avec les en-têtes s'il n'existe pas. S'il existe déjà,
// ajoute en fin de ligne les en-têtes qui manqueraient, sans jamais
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
    .map(row => rowToObject_(headers, row));
}

function rowToObject_(headers, row) {
  const obj = {};
  headers.forEach((h, i) => {
    const val = row[i];
    obj[h] = (val instanceof Date) ? val.toISOString() : val;
  });
  return obj;
}

function readRecentJournal_() {
  const cutoff = Date.now() - IOT_JOURNAL_DAYS * 86400000;
  return readSheet_(SHEET_IOT_JOURNAL).filter(j => {
    const t = new Date(j.Timestamp).getTime();
    return isNaN(t) || t >= cutoff;
  });
}

function appendJournal_(rows) {
  if (!rows.length) return;
  const sheet = getSheet_(SHEET_IOT_JOURNAL, IOT_JOURNAL_HEADERS);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, IOT_JOURNAL_HEADERS.length).setValues(rows);
}

// ─── PASSATION ───

function addPassation_(p) {
  const sheet = getSheet_(SHEET_PASSATIONS, PASSATION_HEADERS);
  const id = Utilities.getUuid();
  appendRowByHeader_(sheet, {
    'ID': id,
    'Date': p.date || new Date().toISOString(),
    'Créneau': p.creneau || '',
    'Ville': p.ville || '',
    'Astreinte': p.astreinte || '',
    'Créé par': p.creePar || '',
    'ID Verification': p.idVerification || '',
    'ID Verification Comment': p.idVerificationComment || '',
    'TripsOk': p.tripsOk || '',
    'Trips Detail': p.tripsDetail || '',
    'Message': p.message || '',
    'Offline Available': p.offlineAvailable ?? '',
    'Offline Total': p.offlineTotal ?? '',
    'Offline Liste': p.offlineListe || ''
  });
  return { ok: true, id: id };
}

// ─── IoT / SCOOTERS OFFLINE ───
//
// Chaque opérateur importe son export Atom (Vehicles > List view, filtré
// "Last IoT signal ≥ 60min") au moment de sa passation. L'export est un
// instantané complet de "tout ce qui n'a pas de signal en ce moment" — donc
// à chaque import :
//   1. un véhicule ouvert qui n'apparaît PLUS dans l'export a retrouvé du
//      signal → résolu automatiquement
//   2. un véhicule déjà ouvert qui réapparaît → ses champs Atom sont
//      rafraîchis, son suivi opérateur (position, contact) est préservé
//   3. un véhicule nouveau → nouvelle entrée ouverte
// Puis la feuille est compactée : les lignes résolues partent dans
// "IoT Archive" et les doublons éventuels sont fusionnés. "IoT Signal" ne
// contient donc jamais que les scooters offline en ce moment, ce qui garde
// toutes les lectures rapides sans maintenance manuelle.

function readIotSheetIndexed_(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const col = {};
  headers.forEach((h, i) => { col[h] = i + 1; }); // 1-based colonne
  const lastRow = sheet.getLastRow();
  const data = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  return { col, data };
}

// Pas de géocodage automatique : la position GPS part telle quelle (lien
// carte) et l'adresse n'est saisie à la main que si elle diffère du GPS
// ('Adresse Reelle', voir updateIoTLocked_).
function importIoT_(body) {
  const incoming = (body.rows || []).filter(r => r.vehicleNumber);

  // Deux opérateurs peuvent importer à quelques minutes d'intervalle : le
  // verrou rend le cycle lecture-fusion-écriture atomique.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return importIoTLocked_(body, incoming);
  } finally {
    lock.releaseLock();
  }
}

// Toute la fusion se fait en mémoire puis s'écrit en une fois (setValues) :
// avant, chaque véhicule mis à jour coûtait une dizaine d'appels setValue,
// soit des centaines d'écritures par import.
function importIoTLocked_(body, incoming) {
  const sheet = getSheet_(SHEET_IOT, IOT_HEADERS);
  const now = new Date().toISOString();
  const importedBy = body.importedBy || '';
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const c = {};
  headers.forEach((h, i) => { c[h] = i; });
  const lastRow = sheet.getLastRow();
  const rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  const journal = [];

  // Index des entrées ouvertes par véhicule. S'il reste des doublons, on
  // garde la plus récente et on clôt les autres (elles partent à l'archive).
  const time = v => { const d = new Date(v); return isNaN(d) ? 0 : d.getTime(); };
  const rank = row => time(row[c['Last Import']]) || time(row[c['Last Seen']]);
  const openIdx = {};
  rows.forEach((row, i) => {
    const vn = row[c['Vehicle Number']];
    if (!vn || row[c['Resolved']] === 'oui') return;
    const key = String(vn);
    const prev = openIdx[key];
    if (prev === undefined) { openIdx[key] = i; return; }
    const keep = rank(row) >= rank(rows[prev]) ? i : prev;
    rows[keep === i ? prev : i][c['Resolved']] = 'oui';
    openIdx[key] = keep;
  });

  // 1) Signal revenu : entrées ouvertes absentes de ce nouvel export
  const incomingKeys = new Set(incoming.map(r => String(r.vehicleNumber)));
  let autoResolved = 0;
  Object.keys(openIdx).forEach(vn => {
    if (incomingKeys.has(vn)) return;
    const row = rows[openIdx[vn]];
    row[c['Resolved']] = 'oui';
    row[c['Last Seen']] = now;
    journal.push([now, vn, 'resolved', "Signal IoT revenu (résolu automatiquement à l'import)", 'Système']);
    autoResolved++;
  });

  // 2) Mise à jour des entrées existantes / création des nouvelles
  let updated = 0, added = 0;
  incoming.forEach(r => {
    const vn = String(r.vehicleNumber);
    const fields = {
      'Group': r.group, 'Vehicle Status': r.status, 'Vehicle Battery': r.battery,
      'IoT Last Update': r.iotLastUpdate, 'Hours Since Signal': r.hoursSinceSignal,
      'Latitude': r.lat, 'Longitude': r.lng,
      'Last Ride': r.lastRide, 'Last Seen': now, 'Last Import': now
    };
    const idx = openIdx[vn];
    if (idx !== undefined) {
      Object.keys(fields).forEach(h => { if (h in c) rows[idx][c[h]] = fields[h]; });
      updated++;
    } else {
      const values = Object.assign({
        'Vehicle Number': vn, 'Position Check': 'non_verifie', 'Contact Joint': '',
        'Lieu Indique': '', 'Resolved': 'non', 'First Seen': now, 'Created By': importedBy
      }, fields);
      rows.push(headers.map(h => (h in values) ? values[h] : ''));
      openIdx[vn] = rows.length - 1; // un vn répété plus loin dans l'import met à jour cette ligne
      journal.push([now, vn, 'import', 'Signalé (' + r.group + ', ' + r.status + ')', importedBy]);
      added++;
    }
  });

  // 3) Compactage : les résolues partent à l'archive, la feuille ne garde
  // que les scooters offline en ce moment.
  const toArchive = rows.filter(row => row[c['Resolved']] === 'oui');
  const toKeep = rows.filter(row => row[c['Resolved']] !== 'oui');
  if (toArchive.length) {
    const archive = getSheet_(SHEET_IOT_ARCHIVE, headers);
    archive.getRange(archive.getLastRow() + 1, 1, toArchive.length, lastCol).setValues(toArchive);
  }
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  if (toKeep.length) sheet.getRange(2, 1, toKeep.length, lastCol).setValues(toKeep);
  appendJournal_(journal);

  // La liste à jour est renvoyée directement : le front n'a pas à refaire
  // un appel de lecture (~3s) juste après l'import.
  return {
    ok: true, added: added, updated: updated, autoResolved: autoResolved, archived: toArchive.length,
    entries: toKeep.map(row => rowToObject_(headers, row))
  };
}

function updateIoT_(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return updateIoTLocked_(body);
  } finally {
    lock.releaseLock();
  }
}

function updateIoTLocked_(body) {
  const sheet = getSheet_(SHEET_IOT, IOT_HEADERS);
  const now = new Date().toISOString();
  const vn = body.vehicleNumber;
  const by = body.by || '';
  const type = body.type; // 'position' | 'contact' | 'adresse' | 'resolved'

  const { col, data } = readIotSheetIndexed_(sheet);
  let rowNum = null;
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][col['Vehicle Number'] - 1]) === String(vn) && data[i][col['Resolved'] - 1] !== 'oui') {
      rowNum = i + 2;
      break;
    }
  }
  if (!rowNum) return { ok: false, error: 'Véhicule introuvable ou déjà résolu' };

  let journalRow;
  if (type === 'position') {
    sheet.getRange(rowNum, col['Position Check']).setValue(body.value);
    const label = body.value === 'correspond' ? 'Correspond' : body.value === 'ne_correspond_pas' ? 'Ne correspond pas' : 'Non vérifié';
    journalRow = [now, vn, 'position_check', label, by];
  } else if (type === 'contact') {
    sheet.getRange(rowNum, col['Contact Joint']).setValue(body.contactJoint);
    sheet.getRange(rowNum, col['Lieu Indique']).setValue(body.lieuIndique || '');
    const detail = body.contactJoint === 'oui'
      ? 'Contact joint — lieu indiqué : ' + (body.lieuIndique || '—')
      : 'Contact non joint';
    journalRow = [now, vn, 'contact', detail, by];
  } else if (type === 'adresse') {
    sheet.getRange(rowNum, col['Adresse Reelle']).setValue(body.adresse || '');
    journalRow = [now, vn, 'adresse', 'Adresse réelle : ' + (body.adresse || '(effacée)'), by];
  } else if (type === 'resolved') {
    sheet.getRange(rowNum, col['Resolved']).setValue(body.resolved ? 'oui' : 'non');
    journalRow = [now, vn, body.resolved ? 'resolved' : 'reopened', body.resolved ? 'Marqué résolu manuellement' : 'Ré-ouvert', by];
  } else {
    return { ok: false, error: 'Type inconnu: ' + type };
  }
  sheet.getRange(rowNum, col['Last Seen']).setValue(now);
  appendJournal_([journalRow]);
  return { ok: true };
}
