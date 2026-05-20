// ============================================================
//  HEURÉKA — CRUD CENTRALISÉ
//  Extension de sheets-script.gs — à déployer dans le même
//  projet Apps Script (nouveau fichier .gs dans l'éditeur).
//  Gère les 12 onglets du Sheet unifié « Gestions Heuréka ».
// ============================================================

// ── Définitions des nouveaux onglets ──────────────────────

const NEW_SHEETS = {
  PROJETS:     'Projets',
  ACTIVITES:   'Activités',
  COURRIELS:   'Courriels',
  CLIENTS:     'Clients',
  EMPLOYES:    'Employés',
  CHANTIERS:   'Chantiers',
  PLANNING:    'Planning',
  PUNCH:       'Punch',
  MKT_CONTENU: 'Marketing_Contenu',
  MKT_TIKTOK:  'Marketing_TikTok',
  CONFIG:      'Config',
};

const NEW_HEADERS = {
  PROJETS:     ['ID_Projet',    'Client_ID',    'Nom_Projet',       'Statut',          'Priorité',         'Type_Projet',       'Source_Référence',  'Prix_Estimé',      'Coût_Matériaux',  'Coût_MO',     'Lien_Soumission', 'Notes', 'Date_Créé', 'Date_RDV'],
  ACTIVITES:   ['ID_Activité',  'Projet_ID',    'Client_ID',        'Type',            'Date',             'Description',       'Complétée',         'Date_Créé'],
  COURRIELS:   ['ID_Courriel',  'Projet_ID',    'Client_ID',        'Destinataire',    'Sujet',            'Message',           'Template',          'Date_Envoi'],
  CLIENTS:     ['ID_Client',    'Nom',          'Téléphone',        'Courriel',        'Adresse',          'Ville',             'Référé_Par',        'Notes',            'Date_Créé'],
  EMPLOYES:    ['ID_Employé',   'Prénom',       'Nom',              'Rôle',            'Taux_Horaire',     'Téléphone',         'Email',             'NAS_4_Derniers',   'Statut',          'Type_Équipe'],
  CHANTIERS:   ['ID_Chantier',  'Nom_Chantier', 'Client_ID',        'Adresse',         'Statut',           'Date_Début',        'Date_Fin_Prévue',   'Budget',           'Contremaître_ID', 'Notes',       'Lieu_Fixe',       'Type_Équipe_Associée'],
  PLANNING:    ['ID_Planning',  'Horaire_ID',   'Employé_ID',       'Chantier_ID',     'Date',             'Heure_Début',       'Heure_Fin',         'Note'],
  PUNCH:       ['ID_Punch',     'Employé_ID',   'Chantier_ID',      'Date',            'Punch_In',         'Punch_Out',         'Heures',            'Statut',           'Note',            'Rapport_Fin_Journée'],
  MKT_CONTENU: ['ID_Contenu',   'Projet_ID',    'Client',           'Type_Travaux',    'Ville',            'Date_Fin_Chantier', 'Statut_Contenu',    'Type_Post',        'Plateforme',      'Texte_Contenu','Date_Publication'],
  MKT_TIKTOK:  ['ID_Tendance',  'Hashtag',      'Vues',             'Score_Pertinence','Date_Capturée',    'Utilisé'],
  CONFIG:      ['Clé',          'Valeur'],
};

const NEW_COLORS = {
  PROJETS:     { bg: '#0a1628', fg: '#D4AF37' },
  ACTIVITES:   { bg: '#0a2010', fg: '#4CAF50' },
  COURRIELS:   { bg: '#1a1a2e', fg: '#7B68EE' },
  CLIENTS:     { bg: '#1a1a1a', fg: '#D4AF37' },
  EMPLOYES:    { bg: '#1a1a1a', fg: '#D4AF37' },
  CHANTIERS:   { bg: '#1a1000', fg: '#FFA500' },
  PLANNING:    { bg: '#0a1628', fg: '#D4AF37' },
  PUNCH:       { bg: '#1a0000', fg: '#FF6B6B' },
  MKT_CONTENU: { bg: '#0d1428', fg: '#FF69B4' },
  MKT_TIKTOK:  { bg: '#0d0d1a', fg: '#00FFFF' },
  CONFIG:      { bg: '#1a1a1a', fg: '#888888' },
};

const ID_PREFIXES = {
  'Projets':           'PRJ',
  'Activités':         'ACT',
  'Courriels':         'COU',
  'Clients':           'CLI',
  'Employés':          'EMP',
  'Chantiers':         'CHA',
  'Planning':          'PLN',
  'Punch':             'PUN',
  'Marketing_Contenu': 'MCT',
  'Marketing_TikTok':  'TIK',
  'Config':            'CFG',
};

// ── CRUD générique ─────────────────────────────────────────

/**
 * Retourne toutes les lignes d'un onglet en JSON (tableau d'objets).
 * Le nom des clés = noms de colonnes (ligne 1 du Sheet).
 */
function getData(sheetName) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const raw = sheet.getDataRange().getValues();
  if (raw.length < 2) return [];

  const headers = raw[0].map(h => h.toString().trim());
  const rows    = [];

  for (let i = 1; i < raw.length; i++) {
    if (!raw[i][0] || raw[i][0].toString().trim() === '') continue;
    const obj = {};
    headers.forEach((h, j) => {
      const v = raw[i][j];
      obj[h]  = v !== null && v !== undefined ? v.toString().trim() : '';
    });
    rows.push(obj);
  }

  return rows;
}

/**
 * Ajoute une ligne. Génère l'ID automatiquement (format PRJ-0001).
 * data = { NomColonne: valeur, ... }
 */
function addRow(sheetName, data) {
  var headerKey = _findHeaderKey(sheetName);
  if (!headerKey) return { status: 'error', message: 'Onglet inconnu: ' + sheetName };

  var sheet   = _getOrInitNewSheet(sheetName);
  var headers = NEW_HEADERS[headerKey];
  var prefix  = ID_PREFIXES[sheetName] || 'ROW';
  var id      = _generateId(sheet, prefix);

  // Injecter l'ID et les valeurs par défaut
  data[headers[0]] = id;

  if (headers.indexOf('Date_Créé')        >= 0 && !data['Date_Créé'])        data['Date_Créé']        = _today();
  if (headers.indexOf('Date_Envoi')       >= 0 && !data['Date_Envoi'])       data['Date_Envoi']       = _today();
  if (headers.indexOf('Date_Capturée')    >= 0 && !data['Date_Capturée'])    data['Date_Capturée']    = _today();

  if (sheetName === 'Projets'           && !data['Statut'])         data['Statut']         = 'À contacter';
  if (sheetName === 'Marketing_Contenu' && !data['Statut_Contenu']) data['Statut_Contenu'] = 'En attente de contenu';
  if (sheetName === 'Employés'          && !data['Statut'])         data['Statut']         = 'Actif';
  if (sheetName === 'Punch'             && !data['Statut'])         data['Statut']         = 'En cours';

  var row = headers.map(function(h) { return data[h] !== undefined ? data[h] : ''; });
  sheet.appendRow(row);

  var lastRow = sheet.getLastRow();
  var color   = NEW_COLORS[headerKey] || { bg: '#0d0d0d', fg: '#ffffff' };
  sheet.getRange(lastRow, 1, 1, headers.length)
    .setBackground(color.bg)
    .setFontColor(color.fg);

  return { status: 'ok', id: id, message: 'Ligne ajoutée' };
}

/**
 * Met à jour une ligne par son ID (colonne 0).
 * Seules les clés présentes dans data sont modifiées.
 */
function updateRow(sheetName, id, data) {
  var headerKey = _findHeaderKey(sheetName);
  if (!headerKey) return { status: 'error', message: 'Onglet inconnu: ' + sheetName };

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { status: 'error', message: 'Onglet non trouvé: ' + sheetName };

  var headers = NEW_HEADERS[headerKey];
  var result  = findRowById(sheet, id, 0);
  if (result.rowIndex < 0) return { status: 'error', message: 'ID introuvable: ' + id };

  var existing = result.rowData;
  var row = headers.map(function(h, i) {
    if (i === 0) return id;
    return data[h] !== undefined ? data[h] : (existing[i] !== undefined ? existing[i].toString() : '');
  });

  sheet.getRange(result.rowIndex, 1, 1, headers.length).setValues([row]);
  return { status: 'ok', id: id, message: 'Ligne mise à jour' };
}

/**
 * Supprime une ligne immédiatement par son ID. Elle ne réapparaît jamais.
 */
function deleteRow(sheetName, id) {
  if (!id) return { status: 'error', message: 'ID manquant' };
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { status: 'error', message: 'Onglet introuvable: ' + sheetName };

  var result = findRowById(sheet, id, 0);
  if (result.rowIndex > 0) {
    sheet.deleteRow(result.rowIndex);
    return { status: 'ok', id: id, message: 'Ligne supprimée' };
  }
  return { status: 'error', message: 'ID introuvable: ' + id };
}

/**
 * Lit une valeur dans l'onglet Config (colonne Clé → Valeur).
 */
function getConfigValue(key) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(NEW_SHEETS.CONFIG);
  if (!sheet) return null;

  var raw = sheet.getDataRange().getValues();
  for (var i = 1; i < raw.length; i++) {
    if (raw[i][0] && raw[i][0].toString().trim() === key) {
      return raw[i][1] !== undefined ? raw[i][1].toString() : '';
    }
  }
  return null;
}

// ── Initialisation complète ───────────────────────────────

/**
 * Crée tous les onglets NEW_SHEETS qui n'existent pas encore.
 * Ne touche pas aux onglets existants.
 * À exécuter une fois, ou via le menu Heureka.
 */
function initAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var created = [];
  var skipped = [];

  Object.keys(NEW_SHEETS).forEach(function(key) {
    var name    = NEW_SHEETS[key];
    var headers = NEW_HEADERS[key];
    var color   = NEW_COLORS[key] || { bg: '#1a1a1a', fg: '#D4AF37' };

    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      _formatNewSheet(sheet, headers, color);
      created.push(name);
    } else {
      skipped.push(name);
    }
  });

  // Pré-remplir Config si vide
  var configSheet = ss.getSheetByName(NEW_SHEETS.CONFIG);
  if (configSheet && configSheet.getLastRow() <= 1) {
    var defaults = [
      ['CODE_ACCES',      'Heureka456654!?'],
      ['NOM_ENTREPRISE',  'Heuréka Construction'],
      ['AVIS_GOOGLE_URL', ''],
      ['SESSION_DUREE_H', '8'],
    ];
    configSheet.getRange(2, 1, defaults.length, 2).setValues(defaults);
    configSheet.getRange(2, 1, defaults.length, 2)
      .setBackground('#0d0d0d').setFontColor('#cccccc');
  }

  Logger.log('initAllSheets: créés=' + created.join(', ') + ' | ignorés=' + skipped.join(', '));

  try {
    SpreadsheetApp.getUi().alert(
      '✅ initAllSheets() terminé!\n\n' +
      'Créés (' + created.length + '): ' + (created.join(', ') || 'aucun') + '\n\n' +
      'Déjà existants (' + skipped.length + '): ' + (skipped.join(', ') || 'aucun') + '\n\n' +
      'Prochaine étape:\n' +
      'Déployer → Gérer les déployées → Nouvelle version'
    );
  } catch (e) {
    // Sans UI (trigger) — OK
  }
}

// ── Helpers internes ──────────────────────────────────────

function _findHeaderKey(sheetName) {
  return Object.keys(NEW_SHEETS).find(function(k) { return NEW_SHEETS[k] === sheetName; }) || null;
}

function _getOrInitNewSheet(sheetName) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    var headerKey = _findHeaderKey(sheetName);
    sheet = ss.insertSheet(sheetName);
    _formatNewSheet(sheet, NEW_HEADERS[headerKey], NEW_COLORS[headerKey]);
  }
  return sheet;
}

function _formatNewSheet(sheet, headers, color) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var hdrRange = sheet.getRange(1, 1, 1, headers.length);
  hdrRange.setBackground(color.bg)
          .setFontColor(color.fg)
          .setFontWeight('bold')
          .setFontSize(10);

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);

  for (var c = 1; c <= headers.length; c++) {
    sheet.setColumnWidth(c, c === 1 ? 130 : 140);
  }
}

function _generateId(sheet, prefix) {
  var data   = sheet.getDataRange().getValues();
  var maxNum = 0;
  for (var i = 1; i < data.length; i++) {
    var id = data[i][0] ? data[i][0].toString() : '';
    if (id.indexOf(prefix + '-') === 0) {
      var num = parseInt(id.split('-')[1]) || 0;
      if (num > maxNum) maxNum = num;
    }
  }
  var next = String(maxNum + 1);
  while (next.length < 4) next = '0' + next;
  return prefix + '-' + next;
}

function _today() {
  return Utilities.formatDate(new Date(), 'America/Toronto', 'dd/MM/yyyy');
}
