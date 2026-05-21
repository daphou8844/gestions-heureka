// ============================================================
//  HEURÉKA — DRIVE SCRIPT
//  Gestion du lecteur partagé "Gestions Heureka Construction"
//
//  INSTALLATION :
//  1. Dans Apps Script, crée un nouveau fichier .gs "drive-script"
//  2. Colle ce contenu
//  3. Exécute showSharedDriveId() une fois pour trouver l'ID
//  4. Colle l'ID dans SHARED_DRIVE_FOLDER_ID ci-dessous
//  5. Redéploie le Apps Script
// ============================================================

// ── CONFIG — à remplir après avoir exécuté showSharedDriveId() ──
var SHARED_DRIVE_FOLDER_ID = '0ABkq_Qy0NQ6IUk9PVA';

var DRIVE_FOLDER_NAMES = {
  CHANTIERS:   '01 - Chantiers',
  SOUMISSIONS: '02 - Soumissions',
  RH_PAIE:     '03 - RH & Paie',
  ADMIN:       '04 - Administration',
  MODELES:     '05 - Modèles',
  APPS:        '06 - Apps Heureka',
};

var CHANTIER_SUBFOLDERS = [
  'Plans & Devis',
  'Photos',
  'Contrats & Permis',
  'Factures & Paiements',
  'Heures',
];

// ── SETUP — exécute cette fonction pour trouver l'ID ────────

function showSharedDriveId() {
  // Méthode 1 : via Drive Advanced Service (si activé)
  try {
    var page = Drive.Drives.list({ maxResults: 50 });
    var drives = (page.drives || page.items || []);
    if (drives.length > 0) {
      Logger.log('=== LECTEURS PARTAGÉS TROUVÉS ===');
      drives.forEach(function(d) {
        Logger.log(d.name + '  →  ID: ' + d.id);
      });
      Logger.log('\nCopie l\'ID du lecteur "Gestions Heureka Construction"');
      Logger.log('et colle-le dans SHARED_DRIVE_FOLDER_ID dans drive-script.gs');
      return;
    }
  } catch(e) {
    Logger.log('Drive Advanced Service non activé — utilise la méthode manuelle:');
  }

  // Méthode 2 : manuelle
  Logger.log('=== COMMENT TROUVER L\'ID MANUELLEMENT ===');
  Logger.log('1. Ouvre drive.google.com dans ton navigateur');
  Logger.log('2. Clique sur "Lecteurs partagés" dans la colonne gauche');
  Logger.log('3. Clique sur "Gestions Heureka Construction"');
  Logger.log('4. Regarde l\'URL — elle ressemble à:');
  Logger.log('   https://drive.google.com/drive/folders/XXXXXXXXXXXXXXXX');
  Logger.log('5. Copie la partie XXXXXXXXXXXXXXXX');
  Logger.log('6. Colle-la dans SHARED_DRIVE_FOLDER_ID dans drive-script.gs');
}

// ── HELPERS DRIVE ────────────────────────────────────────────

function _driveRoot() {
  if (!SHARED_DRIVE_FOLDER_ID) {
    throw new Error(
      'SHARED_DRIVE_FOLDER_ID non configuré. ' +
      'Exécute showSharedDriveId() pour trouver l\'ID.'
    );
  }
  return DriveApp.getFolderById(SHARED_DRIVE_FOLDER_ID);
}

// Obtenir ou créer un sous-dossier
function _getOrCreateFolder(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

// Obtenir le dossier 01 - Chantiers / [année]
function _getYearFolder() {
  var root    = _driveRoot();
  var chanDir = _getOrCreateFolder(root, DRIVE_FOLDER_NAMES.CHANTIERS);
  var year    = new Date().getFullYear().toString();
  return _getOrCreateFolder(chanDir, year);
}

// Trouver le prochain numéro CHAN-XXX
function _nextChanNumber(yearFolder) {
  var max = 0;
  var it  = yearFolder.getFolders();
  while (it.hasNext()) {
    var name = it.next().getName();
    var match = name.match(/^CHAN-(\d+)/);
    if (match) {
      var n = parseInt(match[1], 10);
      if (n > max) max = n;
    }
  }
  var next = String(max + 1);
  while (next.length < 3) next = '0' + next;
  return 'CHAN-' + next;
}

// ── CRÉATION DOSSIER CHANTIER ────────────────────────────────

/**
 * Crée la structure complète dans Drive.
 * Retourne { chanId, driveId, driveUrl }
 *
 * @param {string} clientNom - Nom du client
 * @param {string} ville     - Ville du chantier
 * @param {string} projetId  - ID du projet (pour mise à jour Sheet)
 */
function createChantierDriveFolder(clientNom, ville, projetId) {
  var yearFolder = _getYearFolder();
  var chanId     = _nextChanNumber(yearFolder);
  var folderName = chanId + ' - ' + (clientNom || 'Client') + (ville ? ' - ' + ville : '');

  var chanFolder = _getOrCreateFolder(yearFolder, folderName);

  CHANTIER_SUBFOLDERS.forEach(function(sub) {
    _getOrCreateFolder(chanFolder, sub);
  });

  var driveId  = chanFolder.getId();
  var driveUrl = 'https://drive.google.com/drive/folders/' + driveId;

  // Mettre à jour l'onglet Chantiers dans le Sheet
  _saveDriveIdToSheet(projetId, chanId, driveId, driveUrl);

  return { chanId: chanId, driveId: driveId, driveUrl: driveUrl };
}

// Sauvegarde Drive_ID et Drive_URL dans l'onglet Chantiers
function _saveDriveIdToSheet(projetId, chanId, driveId, driveUrl) {
  try {
    var sheet = _ss().getSheetByName('Chantiers');
    if (!sheet) return;

    var data    = sheet.getDataRange().getValues();
    var headers = data[0];
    var projetCol = headers.indexOf('Projet_ID');
    var driveIdCol = headers.indexOf('Drive_ID');
    var driveUrlCol = headers.indexOf('Drive_URL');

    // Créer les colonnes si elles n'existent pas
    if (driveIdCol === -1) {
      driveIdCol = headers.length;
      sheet.getRange(1, driveIdCol + 1).setValue('Drive_ID');
      headers.push('Drive_ID');
    }
    if (driveUrlCol === -1) {
      driveUrlCol = headers.length;
      sheet.getRange(1, driveUrlCol + 1).setValue('Drive_URL');
      headers.push('Drive_URL');
    }

    // Trouver la ligne par Projet_ID ou ID_Chantier
    for (var i = 1; i < data.length; i++) {
      var rowProjetId = data[i][projetCol] || '';
      var rowChanId   = data[i][0] || '';
      if (rowProjetId === projetId || rowChanId === chanId) {
        sheet.getRange(i + 1, driveIdCol + 1).setValue(driveId);
        sheet.getRange(i + 1, driveUrlCol + 1).setValue(driveUrl);
        break;
      }
    }
  } catch(e) {
    Logger.log('_saveDriveIdToSheet error: ' + e);
  }
}

// ── HANDLER doPost ────────────────────────────────────────────
// Appelé depuis sheets-script.gs doPost avec action:'createChantierFolder'

function handleCreateChantierFolder(data) {
  try {
    var result = createChantierDriveFolder(
      data.clientNom || '',
      data.ville     || '',
      data.projetId  || ''
    );

    // Envoyer email de confirmation si email fourni
    if (data.email) {
      _sendFolderCreatedEmail(data.email, data.clientNom, result.driveUrl);
    }

    return {
      status:   'ok',
      chanId:   result.chanId,
      driveId:  result.driveId,
      driveUrl: result.driveUrl,
    };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

// Email de confirmation
function _sendFolderCreatedEmail(to, clientNom, driveUrl) {
  try {
    GmailApp.sendEmail(
      to,
      'Dossier chantier créé — ' + clientNom,
      '',
      {
        htmlBody:
          '<p>Bonjour,</p>' +
          '<p>Le dossier Drive pour le chantier <strong>' + clientNom + '</strong> a été créé automatiquement.</p>' +
          '<p><a href="' + driveUrl + '" style="background:#D4AF37;color:#000;padding:10px 20px;text-decoration:none;border-radius:6px;font-weight:bold">📁 Ouvrir le dossier Drive</a></p>' +
          '<p>Structure créée :<br>' +
          '&nbsp;📁 Plans & Devis<br>' +
          '&nbsp;📁 Photos<br>' +
          '&nbsp;📁 Contrats & Permis<br>' +
          '&nbsp;📁 Factures & Paiements<br>' +
          '&nbsp;📁 Heures</p>' +
          '<p>— Heuréka Construction</p>',
      }
    );
  } catch(e) {
    Logger.log('Email error: ' + e);
  }
}

// ── UTILITAIRE — obtenir un sous-dossier d'un chantier ───────

function getChantierSubfolder(driveId, subfolderName) {
  try {
    var chanFolder = DriveApp.getFolderById(driveId);
    var it = chanFolder.getFoldersByName(subfolderName);
    return it.hasNext() ? it.next() : chanFolder.createFolder(subfolderName);
  } catch(e) {
    Logger.log('getChantierSubfolder error: ' + e);
    return null;
  }
}
