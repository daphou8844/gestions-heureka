/**
 * ============================================================
 *  GESTIONS HEUREKA — Google Apps Script
 *  Connecte admin.html + punch.html à Google Sheets
 * ============================================================
 *
 *  INSTALLATION (5 minutes):
 *  1. Ouvrir votre Google Sheets (nouveau ou existant)
 *  2. Extensions → Apps Script
 *  3. Tout sélectionner (Ctrl+A) et coller ce code
 *  4. Sauvegarder (Ctrl+S) → nommer "Heureka API"
 *  5. Exécuter → initSheets() une première fois
 *     (accepter les permissions demandées)
 *  6. Déployer → Nouvelle déployée
 *     · Type            : Application Web
 *     · Exécuter en tant: Vous-même
 *     · Accès           : Tout le monde (anonyme)
 *  7. Copier l'URL /exec → coller dans Paramètres de admin.html
 *     et dans punch.html (⚙ Paramètres)
 *
 *  ONGLETS CRÉÉS AUTOMATIQUEMENT:
 *  · Équipe    — liste des employés
 *  · Jobs      — chantiers de construction
 *  · Punchs    — pointages (lecture/écriture depuis les 2 apps)
 *  · Vue Live  — statut temps réel des employés
 *  · Horaires  — horaires envoyés aux gars
 * ============================================================
 */

// ============================================================
//  CONFIGURATION
// ============================================================

const SHEETS = {
  EQUIPE:   'Équipe',
  JOBS:     'Jobs',
  PUNCHS:   'Punchs',
  LIVE:     'Vue Live',
  HORAIRES: 'Horaires',
};

const HEADERS = {
  EQUIPE:   ['ID', 'Prénom', 'Nom', 'Rôle', 'Taux/h ($)', 'Téléphone', 'Email', 'NAS (4 der.)', 'Actif', 'Date ajout'],
  JOBS:     ['ID', 'Nom chantier', 'Client', 'Adresse', 'Statut', 'Date début', 'Date fin', 'Budget ($)', 'Contremaître', 'Notes', 'Dernière MAJ'],
  PUNCHS:   ['ID', 'Employé ID', 'Nom employé', 'Job ID', 'Nom chantier', 'Date', 'Punch In', 'Punch Out', 'Heures nettes', 'Pauses (min)', 'Notes', 'Photos #', 'Source', 'Enregistré le'],
  LIVE:     ['Employé ID', 'Nom employé', 'Job ID', 'Nom chantier', 'Statut', 'Punch In', 'Date', 'Dernière MAJ'],
  HORAIRES: ['ID', 'Job ID', 'Nom chantier', 'Date', 'Début', 'Fin', 'Employés (IDs)', 'Employés (Noms)', 'Note', 'Envoyé le'],
};

// Couleurs d'en-tête par onglet
const COLORS = {
  EQUIPE:   { bg: '#1a1a1a', fg: '#D4AF37' },
  JOBS:     { bg: '#1a1a1a', fg: '#D4AF37' },
  PUNCHS:   { bg: '#1a1a1a', fg: '#D4AF37' },
  LIVE:     { bg: '#0d3322', fg: '#4CAF50' },
  HORAIRES: { bg: '#1a1a2e', fg: '#D4AF37' },
};

// ============================================================
//  POINT D'ENTRÉE GET — lecture
// ============================================================

function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || 'ping';
    let result;

    switch (action) {
      case 'ping':
        result = { status: 'ok', message: 'Heureka Sheets API active', time: new Date().toISOString() };
        break;
      case 'getEmployees':
        result = { status: 'ok', employees: readSheet(SHEETS.EQUIPE, employeeFromRow) };
        break;
      case 'getJobs':
        result = { status: 'ok', jobs: readSheet(SHEETS.JOBS, jobFromRow) };
        break;
      case 'getPunchs':
        result = { status: 'ok', punchs: getPunchs(e.parameter) };
        break;
      case 'getLive':
        result = { status: 'ok', live: readSheet(SHEETS.LIVE, liveFromRow) };
        break;
      case 'getHoraires':
        result = { status: 'ok', horaires: getHoraires(e.parameter) };
        break;
      case 'getAll':
        result = {
          status: 'ok',
          employees: readSheet(SHEETS.EQUIPE, employeeFromRow),
          jobs:      readSheet(SHEETS.JOBS, jobFromRow),
          punchs:    readSheet(SHEETS.PUNCHS, punchFromRow),
          live:      readSheet(SHEETS.LIVE, liveFromRow),
          horaires:  readSheet(SHEETS.HORAIRES, horaireFromRow),
        };
        break;
      default:
        result = { status: 'error', message: 'Action inconnue: ' + action };
    }

    return jsonOk(result);
  } catch (err) {
    return jsonError(err.toString());
  }
}

// ============================================================
//  POINT D'ENTRÉE POST — écriture
// ============================================================

function doPost(e) {
  try {
    if (!e.postData || !e.postData.contents) {
      return jsonError('Corps de requête vide');
    }

    const data   = JSON.parse(e.postData.contents);
    const action = data.action || '';
    let result;

    switch (action) {

      // --- Sync bulk depuis admin.html ---
      case 'sync':
        result = syncAll(data);
        break;

      // --- Punch individuel depuis punch.html ---
      case 'punch':
        result = savePunchFromApp(data.punch);
        break;

      // --- Mise à jour vue live depuis punch.html ---
      case 'live':
        result = updateLive(data);
        break;

      // --- CRUD employé ---
      case 'saveEmployee':
        result = saveEmployee(data.employee);
        break;
      case 'deleteEmployee':
        result = deleteById(SHEETS.EQUIPE, data.id);
        break;

      // --- CRUD job ---
      case 'saveJob':
        result = saveJob(data.job);
        break;
      case 'deleteJob':
        result = deleteById(SHEETS.JOBS, data.id);
        break;

      // --- CRUD punch ---
      case 'savePunch':
        result = savePunchRecord(data.punch);
        break;
      case 'deletePunch':
        result = deleteById(SHEETS.PUNCHS, data.id);
        break;

      // --- Horaire ---
      case 'saveHoraire':
        result = saveHoraire(data.horaire);
        break;
      case 'deleteHoraire':
        result = deleteById(SHEETS.HORAIRES, data.id);
        break;

      default:
        result = { status: 'error', message: 'Action inconnue: ' + action };
    }

    return jsonOk(result);
  } catch (err) {
    return jsonError(err.toString());
  }
}

// ============================================================
//  SYNC BULK (admin.html → Sheets → admin.html)
// ============================================================

function syncAll(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Écrire les employés
  if (data.employees && data.employees.length) {
    const sheet = getOrCreateSheet(SHEETS.EQUIPE, HEADERS.EQUIPE, 'EQUIPE');
    bulkUpsert(sheet, data.employees, employeeToRow, 0); // col 0 = ID
  }

  // 2. Écrire les jobs
  if (data.jobs && data.jobs.length) {
    const sheet = getOrCreateSheet(SHEETS.JOBS, HEADERS.JOBS, 'JOBS');
    bulkUpsert(sheet, data.jobs, jobToRow, 0);
  }

  // 3. Écrire les horaires
  if (data.horaires && data.horaires.length) {
    const sheet = getOrCreateSheet(SHEETS.HORAIRES, HEADERS.HORAIRES, 'HORAIRES');
    bulkUpsert(sheet, data.horaires, horaireToRow, 0);
  }

  // 4. Merge punchs: écrire les punchs locaux, retourner TOUS les punchs du Sheets
  //    (inclut ceux ajoutés par punch.html depuis d'autres appareils)
  const punchSheet = getOrCreateSheet(SHEETS.PUNCHS, HEADERS.PUNCHS, 'PUNCHS');
  if (data.punchs && data.punchs.length) {
    bulkUpsert(punchSheet, data.punchs, punchToRow, 0);
  }

  // Relire punchs, jobs ET employés pour retourner à admin
  // (inclut les données ajoutées directement dans le Sheet ou via le pipeline)
  const allPunchs    = readSheet(SHEETS.PUNCHS, punchFromRow);
  const allJobs      = readSheet(SHEETS.JOBS,   jobFromRow);
  const allEmployees = readSheet(SHEETS.EQUIPE, employeeFromRow);

  return {
    status:    'ok',
    message:   'Sync complète',
    synced_at: new Date().toISOString(),
    counts: {
      employees: allEmployees.length,
      jobs:      allJobs.length,
      punchs:    allPunchs.length,
      horaires:  (data.horaires || []).length,
    },
    punchs:    allPunchs,     // admin met à jour son localStorage
    jobs:      allJobs,       // admin importe les nouveaux chantiers du pipeline
    employees: allEmployees,  // admin importe les employés ajoutés dans le Sheet
  };
}

// ============================================================
//  PUNCH — depuis punch.html (appareil du gars)
// ============================================================

function savePunchFromApp(p) {
  if (!p || !p.id) return { status: 'error', message: 'Punch invalide' };

  const sheet = getOrCreateSheet(SHEETS.PUNCHS, HEADERS.PUNCHS, 'PUNCHS');

  // Chercher si la ligne existe déjà
  const existing = findRowById(sheet, p.id);

  const now = Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd HH:mm:ss');

  // Calcul heures nettes
  const netHours = calcNetHours(p.punchIn, p.punchOut, p.totalPauseMin || 0);

  const row = [
    p.id,
    p.empId       || '',
    p.empName     || '',
    p.jobId       || '',
    p.jobName     || '',
    p.date        || '',
    p.punchIn     || '',
    p.punchOut    || '',
    netHours      >= 0 ? netHours.toFixed(2) : '',
    p.totalPauseMin || 0,
    p.notes       || '',
    p.photosCount || 0,
    'Punch App',
    now,
  ];

  if (existing.rowIndex > 0) {
    sheet.getRange(existing.rowIndex, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  // Colorier la ligne si punch out rempli
  const lastRow = existing.rowIndex > 0 ? existing.rowIndex : sheet.getLastRow();
  stylePunchRow(sheet, lastRow, !!p.punchOut);

  // Mettre à jour la vue Live
  updateLive({
    empId:    p.empId,
    empName:  p.empName,
    jobId:    p.jobId,
    jobName:  p.jobName,
    status:   p.punchOut ? 'out' : 'in',
    punchIn:  p.punchIn,
    date:     p.date,
  });

  return { status: 'ok', message: 'Punch enregistré', punchId: p.id };
}

function savePunchRecord(p) {
  return savePunchFromApp({ ...p, empName: p.empName || '', jobName: p.jobName || '', totalPauseMin: 0, photosCount: 0 });
}

// ============================================================
//  VUE LIVE — statut temps réel
// ============================================================

function updateLive(data) {
  if (!data.empId) return { status: 'error', message: 'empId manquant' };

  const sheet = getOrCreateSheet(SHEETS.LIVE, HEADERS.LIVE, 'LIVE');
  const existing = findRowById(sheet, data.empId, 0); // col 0 = empId

  const now = Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd HH:mm:ss');
  const row = [
    data.empId   || '',
    data.empName || '',
    data.jobId   || '',
    data.jobName || '',
    data.status  || 'in',
    data.punchIn || '',
    data.date    || '',
    now,
  ];

  if (existing.rowIndex > 0) {
    sheet.getRange(existing.rowIndex, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }

  // Mettre en évidence les lignes actives
  const rowIdx = existing.rowIndex > 0 ? existing.rowIndex : sheet.getLastRow();
  const range = sheet.getRange(rowIdx, 1, 1, row.length);
  if (data.status === 'in') {
    range.setBackground('#0d3322');
  } else if (data.status === 'paused') {
    range.setBackground('#2a1f00');
  } else {
    range.setBackground('#1a1a1a');
  }

  return { status: 'ok', message: 'Live mis à jour' };
}

function clearLive() {
  // Réinitialiser la vue live chaque matin (peut être appelé par un trigger)
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LIVE);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).setBackground('#1a1a1a');
  }
}

// ============================================================
//  EMPLOYÉS
// ============================================================

function getEmployees() {
  return readSheet(SHEETS.EQUIPE, employeeFromRow);
}

function saveEmployee(emp) {
  if (!emp || !emp.id) return { status: 'error', message: 'Employé invalide' };
  const sheet = getOrCreateSheet(SHEETS.EQUIPE, HEADERS.EQUIPE, 'EQUIPE');
  const now = Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd');
  const row = [
    emp.id, emp.fname || '', emp.lname || '', emp.role || '',
    emp.rate || 0, emp.phone || '', emp.email || '', emp.nas || '',
    emp.active === '1' || emp.active === 1 ? 'Oui' : 'Non',
    now,
  ];
  upsertRow(sheet, emp.id, row);
  return { status: 'ok', message: 'Employé enregistré' };
}

// ============================================================
//  JOBS
// ============================================================

function getJobsList() {
  return readSheet(SHEETS.JOBS, jobFromRow);
}

function saveJob(job) {
  if (!job || !job.id) return { status: 'error', message: 'Job invalide' };
  const sheet = getOrCreateSheet(SHEETS.JOBS, HEADERS.JOBS, 'JOBS');
  const now = Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd HH:mm:ss');
  const row = [
    job.id, job.name || '', job.client || '', job.address || '',
    job.status || 'actif', job.start || '', job.end || '',
    job.budget || 0, job.foreman || '', job.notes || '', now,
  ];
  upsertRow(sheet, job.id, row);
  return { status: 'ok', message: 'Job enregistré' };
}

// ============================================================
//  PUNCHS — lecture avec filtres
// ============================================================

function getPunchs(params) {
  const all = readSheet(SHEETS.PUNCHS, punchFromRow);
  if (!params) return all;

  return all.filter(p => {
    if (params.empId && p.empId !== params.empId) return false;
    if (params.date  && p.date  !== params.date)  return false;
    if (params.from  && p.date  <  params.from)   return false;
    if (params.to    && p.date  >  params.to)      return false;
    return true;
  });
}

// ============================================================
//  HORAIRES — lecture avec filtres
// ============================================================

function getHoraires(params) {
  const all = readSheet(SHEETS.HORAIRES, horaireFromRow);
  if (!params) return all;

  return all.filter(h => {
    if (params.date  && h.date  !== params.date)  return false;
    if (params.empId && !h.empIds.includes(params.empId)) return false;
    return true;
  });
}

function saveHoraire(h) {
  if (!h || !h.id) return { status: 'error', message: 'Horaire invalide' };
  const sheet = getOrCreateSheet(SHEETS.HORAIRES, HEADERS.HORAIRES, 'HORAIRES');
  const sentAt = Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd HH:mm:ss');
  const row = [
    h.id, h.jobId || '', h.jobName || '', h.date || '',
    h.start || '', h.end || '',
    (h.empIds || []).join(','),
    (h.empNames || []).join(', '),
    h.note || '', sentAt,
  ];
  upsertRow(sheet, h.id, row);
  return { status: 'ok', message: 'Horaire enregistré' };
}

// ============================================================
//  ROW CONVERTERS — objet → ligne Sheets
// ============================================================

function employeeToRow(e) {
  const now = Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd');
  return [
    e.id, e.fname || '', e.lname || '', e.role || '',
    e.rate || 0, e.phone || '', e.email || '', e.nas || '',
    e.active === '1' || e.active === 1 ? 'Oui' : 'Non', now,
  ];
}

function jobToRow(j) {
  const now = Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd HH:mm:ss');
  return [
    j.id, j.name || '', j.client || '', j.address || '',
    j.status || 'actif', j.start || '', j.end || '',
    j.budget || 0, j.foreman || '', j.notes || '', now,
  ];
}

function punchToRow(p) {
  const now = Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd HH:mm:ss');
  const netHours = calcNetHours(p.punchIn, p.punchOut, 0);
  return [
    p.id, p.empId || '', p.empName || '', p.jobId || '', p.jobName || '',
    p.date || '', p.punchIn || '', p.punchOut || '',
    netHours >= 0 ? netHours.toFixed(2) : '',
    p.totalPauseMin || 0, p.note || p.notes || '', 0, 'Admin', now,
  ];
}

function horaireToRow(h) {
  const now = Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd HH:mm:ss');
  return [
    h.id, h.jobId || '', '', h.date || '',
    h.start || '', h.end || '',
    (h.empIds || []).join(','),
    '',
    h.note || '', now,
  ];
}

// ============================================================
//  ROW CONVERTERS — ligne Sheets → objet
// ============================================================

function employeeFromRow(row, headers) {
  return {
    id:     cell(row, headers, 'ID'),
    fname:  cell(row, headers, 'Prénom'),
    lname:  cell(row, headers, 'Nom'),
    role:   cell(row, headers, 'Rôle'),
    rate:   parseFloat(cell(row, headers, 'Taux/h ($)')) || 0,
    phone:  cell(row, headers, 'Téléphone'),
    email:  cell(row, headers, 'Email'),
    nas:    cell(row, headers, 'NAS (4 der.)'),
    active: cell(row, headers, 'Actif') === 'Oui' ? '1' : '0',
  };
}

function jobFromRow(row, headers) {
  return {
    id:      cell(row, headers, 'ID'),
    name:    cell(row, headers, 'Nom chantier'),
    client:  cell(row, headers, 'Client'),
    address: cell(row, headers, 'Adresse'),
    status:  cell(row, headers, 'Statut'),
    start:   cell(row, headers, 'Date début'),
    end:     cell(row, headers, 'Date fin'),
    budget:  parseFloat(cell(row, headers, 'Budget ($)')) || 0,
    foreman: cell(row, headers, 'Contremaître'),
    notes:   cell(row, headers, 'Notes'),
  };
}

function punchFromRow(row, headers) {
  return {
    id:           cell(row, headers, 'ID'),
    empId:        cell(row, headers, 'Employé ID'),
    empName:      cell(row, headers, 'Nom employé'),
    jobId:        cell(row, headers, 'Job ID'),
    jobName:      cell(row, headers, 'Nom chantier'),
    date:         cell(row, headers, 'Date'),
    punchIn:      cell(row, headers, 'Punch In'),
    punchOut:     cell(row, headers, 'Punch Out'),
    netHours:     parseFloat(cell(row, headers, 'Heures nettes')) || 0,
    totalPauseMin:parseInt(cell(row, headers, 'Pauses (min)')) || 0,
    note:         cell(row, headers, 'Notes'),
  };
}

function liveFromRow(row, headers) {
  return {
    empId:   cell(row, headers, 'Employé ID'),
    empName: cell(row, headers, 'Nom employé'),
    jobId:   cell(row, headers, 'Job ID'),
    jobName: cell(row, headers, 'Nom chantier'),
    status:  cell(row, headers, 'Statut'),
    punchIn: cell(row, headers, 'Punch In'),
    date:    cell(row, headers, 'Date'),
    lastMaj: cell(row, headers, 'Dernière MAJ'),
  };
}

function horaireFromRow(row, headers) {
  const empIdsRaw = cell(row, headers, 'Employés (IDs)');
  return {
    id:       cell(row, headers, 'ID'),
    jobId:    cell(row, headers, 'Job ID'),
    jobName:  cell(row, headers, 'Nom chantier'),
    date:     cell(row, headers, 'Date'),
    start:    cell(row, headers, 'Début'),
    end:      cell(row, headers, 'Fin'),
    empIds:   empIdsRaw ? empIdsRaw.split(',').map(s => s.trim()).filter(Boolean) : [],
    note:     cell(row, headers, 'Note'),
    sentAt:   cell(row, headers, 'Envoyé le'),
  };
}

// ============================================================
//  UTILITAIRES — Sheets
// ============================================================

/**
 * Lire un onglet complet et retourner un tableau d'objets.
 */
function readSheet(sheetName, fromRowFn) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(h => h.toString().trim());
  const results = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] || row[0].toString().trim() === '') continue; // skip empty rows
    try {
      results.push(fromRowFn(row, headers));
    } catch (e) {
      // skip malformed row
    }
  }

  return results;
}

/**
 * Insérer ou mettre à jour une ligne par ID (colonne 0).
 */
function upsertRow(sheet, id, row) {
  const existing = findRowById(sheet, id, 0);
  if (existing.rowIndex > 0) {
    sheet.getRange(existing.rowIndex, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}

/**
 * Insérer/mettre à jour plusieurs lignes en batch (par ID, col idCol).
 */
function bulkUpsert(sheet, items, toRowFn, idCol) {
  if (!items || !items.length) return;

  const data = sheet.getDataRange().getValues();
  const existingIds = {};

  // Indexer les IDs existants
  for (let i = 1; i < data.length; i++) {
    const id = data[i][idCol] ? data[i][idCol].toString().trim() : '';
    if (id) existingIds[id] = i + 1; // 1-based row
  }

  const newRows = [];

  items.forEach(item => {
    const row = toRowFn(item);
    const id  = row[idCol] ? row[idCol].toString() : '';

    if (id && existingIds[id]) {
      // Update existing
      sheet.getRange(existingIds[id], 1, 1, row.length).setValues([row]);
    } else {
      // Queue for append
      newRows.push(row);
    }
  });

  // Append new rows in one batch
  if (newRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
  }
}

/**
 * Supprimer une ligne par ID (cherche dans colonne 0).
 */
function deleteById(sheetName, id) {
  if (!id) return { status: 'error', message: 'ID manquant' };
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { status: 'error', message: 'Onglet introuvable: ' + sheetName };

  const result = findRowById(sheet, id, 0);
  if (result.rowIndex > 0) {
    sheet.deleteRow(result.rowIndex);
    return { status: 'ok', message: 'Ligne supprimée' };
  }
  return { status: 'error', message: 'ID introuvable: ' + id };
}

/**
 * Trouver l'index de ligne (1-based) d'un ID dans une colonne donnée.
 */
function findRowById(sheet, id, colIndex) {
  colIndex = colIndex || 0;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][colIndex] && data[i][colIndex].toString().trim() === id.toString().trim()) {
      return { rowIndex: i + 1, rowData: data[i] };
    }
  }
  return { rowIndex: -1, rowData: null };
}

/**
 * Récupérer ou créer un onglet avec ses en-têtes formatés.
 */
function getOrCreateSheet(name, headers, colorKey) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
    formatSheet(sheet, headers, colorKey);
  } else {
    // Vérifier que les en-têtes existent
    const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    if (!firstRow[0] || firstRow[0].toString().trim() !== headers[0]) {
      sheet.insertRowBefore(1);
      formatSheet(sheet, headers, colorKey);
    }
  }

  return sheet;
}

/**
 * Formater l'en-tête d'un onglet (couleurs, police, freeze).
 */
function formatSheet(sheet, headers, colorKey) {
  const color = COLORS[colorKey] || { bg: '#1a1a1a', fg: '#D4AF37' };
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const hdrRange = sheet.getRange(1, 1, 1, headers.length);
  hdrRange.setBackground(color.bg);
  hdrRange.setFontColor(color.fg);
  hdrRange.setFontWeight('bold');
  hdrRange.setFontSize(10);

  sheet.setFrozenRows(1);

  // Largeurs de colonnes approximatives
  try {
    sheet.setColumnWidth(1, 120); // ID
    if (headers.length > 1) sheet.setColumnWidth(2, 100);
    if (headers.length > 2) sheet.setColumnWidth(3, 120);
  } catch (e) { /* ignore */ }

  // Couleur de fond du corps
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), headers.length)
         .setBackground('#0d0d0d')
         .setFontColor('#ffffff');
  }
}

/**
 * Colorier une ligne de punch selon son statut.
 */
function stylePunchRow(sheet, rowIdx, isComplete) {
  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(rowIdx, 1, 1, lastCol);
  range.setBackground(isComplete ? '#0d1a0d' : '#0d0d1a');
  range.setFontColor('#ffffff');
  // Colonne Heures nettes en or si complète
  if (isComplete) {
    sheet.getRange(rowIdx, 9).setFontColor('#D4AF37').setFontWeight('bold');
  }
}

// ============================================================
//  UTILITAIRES — Calculs
// ============================================================

/**
 * Calculer les heures nettes travaillées.
 */
function calcNetHours(punchIn, punchOut, totalPauseMin) {
  if (!punchIn || !punchOut) return -1;

  const parseTime = t => {
    const parts = t.toString().trim().split(':');
    return parseInt(parts[0]) * 60 + parseInt(parts[1]);
  };

  try {
    const inMins  = parseTime(punchIn);
    const outMins = parseTime(punchOut);
    const pause   = parseFloat(totalPauseMin) || 0;
    const net     = (outMins - inMins - pause) / 60;
    return Math.max(0, net);
  } catch (e) {
    return -1;
  }
}

/**
 * Lire la valeur d'une cellule par nom de colonne.
 */
function cell(row, headers, colName) {
  const idx = headers.indexOf(colName);
  if (idx < 0 || idx >= row.length) return '';
  const val = row[idx];
  return val !== null && val !== undefined ? val.toString().trim() : '';
}

// ============================================================
//  UTILITAIRES — Réponse HTTP
// ============================================================

function jsonOk(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function jsonError(message) {
  const output = ContentService.createTextOutput(JSON.stringify({ status: 'error', message }));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ============================================================
//  INITIALISATION — créer tous les onglets
// ============================================================

/**
 * À exécuter UNE SEULE FOIS après avoir collé ce script.
 * Crée tous les onglets avec les bons en-têtes et la mise en forme.
 */
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.entries(SHEETS).forEach(([key, name]) => {
    getOrCreateSheet(name, HEADERS[key], key);
  });

  // Réorganiser l'ordre des onglets
  const order = [SHEETS.EQUIPE, SHEETS.JOBS, SHEETS.PUNCHS, SHEETS.LIVE, SHEETS.HORAIRES];
  order.forEach((name, i) => {
    const sheet = ss.getSheetByName(name);
    if (sheet) ss.setActiveSheet(sheet), ss.moveActiveSheet(i + 1);
  });

  // Mettre l'onglet Punchs actif
  const punchSheet = ss.getSheetByName(SHEETS.PUNCHS);
  if (punchSheet) ss.setActiveSheet(punchSheet);

  SpreadsheetApp.getUi().alert(
    '✅ Gestions Heureka — Prêt!\n\n' +
    '5 onglets créés:\n' +
    '  · Équipe\n  · Jobs\n  · Punchs\n  · Vue Live\n  · Horaires\n\n' +
    'Prochaine étape:\n' +
    'Déployer → Nouvelle déployée → Application Web\n' +
    'Copier l\'URL /exec dans admin.html (Paramètres)'
  );
}

// ============================================================
//  TRIGGERS AUTOMATIQUES
// ============================================================

/**
 * Installer les triggers automatiques.
 * À exécuter une fois après initSheets().
 */
function installTriggers() {
  // Supprimer les triggers existants pour éviter les doublons
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // Réinitialiser Vue Live chaque matin à 5h
  ScriptApp.newTrigger('clearLive')
    .timeBased()
    .everyDays(1)
    .atHour(5)
    .create();

  // Rapport hebdomadaire le vendredi à 17h
  ScriptApp.newTrigger('weeklyPayrollReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)
    .atHour(17)
    .create();

  SpreadsheetApp.getUi().alert('✅ Triggers installés:\n· Réinitialisation Live: chaque jour à 5h\n· Rapport paie: vendredi 17h');
}

// ============================================================
//  RAPPORT PAIE HEBDOMADAIRE (trigger vendredi)
// ============================================================

function weeklyPayrollReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Calculer la semaine dernière (lun-dim)
  const now    = new Date();
  const dayOfW = now.getDay() || 7;
  const lastMon = new Date(now); lastMon.setDate(now.getDate() - dayOfW - 6);
  const lastSun = new Date(now); lastSun.setDate(now.getDate() - dayOfW);

  const fmtD  = d => Utilities.formatDate(d, 'America/Toronto', 'yyyy-MM-dd');
  const from  = fmtD(lastMon);
  const to    = fmtD(lastSun);

  const employees = readSheet(SHEETS.EQUIPE, employeeFromRow);
  const punchs    = getPunchs({ from, to });

  // Construire l'onglet rapport
  const reportName = 'Paie ' + from + ' → ' + to;
  let reportSheet  = ss.getSheetByName(reportName);
  if (reportSheet) ss.deleteSheet(reportSheet);
  reportSheet = ss.insertSheet(reportName);

  const headers = ['Employé', 'Rôle', 'Heures totales', 'Taux/h ($)', 'Total brut ($)', 'Heures sup.', 'Total sup. ($)', 'GRAND TOTAL ($)'];
  reportSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const hdrRange = reportSheet.getRange(1, 1, 1, headers.length);
  hdrRange.setBackground('#1a1a1a').setFontColor('#D4AF37').setFontWeight('bold');
  reportSheet.setFrozenRows(1);

  let grandTotal = 0;
  const OT_HOURS = 40, OT_RATE = 1.5;
  const rows = [];

  employees.forEach(emp => {
    const empPunchs = punchs.filter(p => p.empId === emp.id && p.punchOut);
    const totalHrs  = empPunchs.reduce((s, p) => s + p.netHours, 0);
    if (!totalHrs) return;

    const reg   = Math.min(totalHrs, OT_HOURS);
    const ot    = Math.max(0, totalHrs - OT_HOURS);
    const rate  = emp.rate || 0;
    const total = reg * rate + ot * rate * OT_RATE;
    grandTotal += total;

    rows.push([
      emp.fname + ' ' + emp.lname, emp.role,
      totalHrs.toFixed(2), rate,
      (reg * rate).toFixed(2),
      ot.toFixed(2),
      (ot * rate * OT_RATE).toFixed(2),
      total.toFixed(2),
    ]);
  });

  if (rows.length) {
    reportSheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    reportSheet.getRange(2, 1, rows.length, headers.length)
               .setBackground('#0d0d0d').setFontColor('#ffffff');
    // Colonne TOTAL en or
    reportSheet.getRange(2, 8, rows.length, 1).setFontColor('#D4AF37').setFontWeight('bold');
  }

  // Ligne total
  const totalRow = reportSheet.getLastRow() + 1;
  reportSheet.getRange(totalRow, 1, 1, headers.length).setValues([
    ['TOTAL', '', '', '', '', '', '', grandTotal.toFixed(2)]
  ]);
  reportSheet.getRange(totalRow, 1, 1, headers.length)
             .setBackground('#D4AF37').setFontColor('#000000').setFontWeight('bold');

  SpreadsheetApp.getActive().setActiveSheet(reportSheet);
  Logger.log('Rapport paie généré: ' + reportName + ' — Total: $' + grandTotal.toFixed(2));
}

// ============================================================
//  MENU PERSONNALISÉ dans Google Sheets
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏗 Heureka')
    .addItem('⚙ Initialiser les onglets', 'initSheets')
    .addItem('⏰ Installer les triggers', 'installTriggers')
    .addSeparator()
    .addItem('🔄 Réinitialiser Vue Live', 'clearLive')
    .addItem('💰 Rapport paie (semaine courante)', 'weeklyPayrollReport')
    .addSeparator()
    .addItem('🔗 Afficher l\'URL du déploiement', 'showDeployUrl')
    .addToUi();
}

function showDeployUrl() {
  SpreadsheetApp.getUi().alert(
    'URL de votre déploiement:\n\n' +
    'Pour trouver l\'URL:\n' +
    '1. Déployer → Gérer les déployées\n' +
    '2. Cliquer sur l\'icône de lien (🔗)\n' +
    '3. Copier l\'URL /exec\n\n' +
    'Coller dans:\n' +
    '· admin.html → ⚙ Paramètres → URL Apps Script\n' +
    '· punch.html → ⚙ → URL Google Sheets'
  );
}
