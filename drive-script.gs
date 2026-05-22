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

// Met à jour Drive_ID dans l'onglet Chantiers par ID_Chantier direct
function _updateChantierDriveId(jobId, driveId, driveUrl) {
  try {
    var sheet = _ss().getSheetByName('Chantiers');
    if (!sheet) return;
    var data = sheet.getDataRange().getValues();
    var h = data[0].map(function(x){ return x.toString().trim(); });
    var idCol     = h.indexOf('ID_Chantier');
    var driveIdCol  = h.indexOf('Drive_ID');
    var driveUrlCol = h.indexOf('Drive_URL');
    if (idCol < 0 || driveIdCol < 0) return;
    for (var i = 1; i < data.length; i++) {
      if ((data[i][idCol]||'').toString().trim() === jobId) {
        sheet.getRange(i+1, driveIdCol+1).setValue(driveId);
        if (driveUrlCol >= 0) sheet.getRange(i+1, driveUrlCol+1).setValue(driveUrl||'');
        Logger.log('Drive_ID mis à jour pour ' + jobId);
        return;
      }
    }
  } catch(e) { Logger.log('_updateChantierDriveId: ' + e); }
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

// ── RAPPORT DE TEMPS — depuis punch.html ─────────────────────

/**
 * Cherche le Drive_ID d'un chantier dans l'onglet Chantiers.
 * Essaie d'abord par ID exact, puis par nom partiel.
 */
function _findChantierDriveId(jobId, jobName) {
  try {
    var sheet = _ss().getSheetByName('Chantiers');
    if (!sheet) return null;

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return null;

    var headers = data[0].map(function(h){ return h.toString().trim(); });
    var idCol      = headers.indexOf('ID_Chantier');
    var nameCol    = headers.indexOf('Nom_Chantier');
    var driveIdCol = headers.indexOf('Drive_ID');
    if (driveIdCol === -1) return null;

    for (var i = 1; i < data.length; i++) {
      var rowDriveId = data[i][driveIdCol] ? data[i][driveIdCol].toString().trim() : '';
      if (!rowDriveId) continue;

      var rowId   = idCol   >= 0 ? (data[i][idCol]   || '').toString().trim() : '';
      var rowName = nameCol >= 0 ? (data[i][nameCol]  || '').toString().trim().toLowerCase() : '';

      if ((jobId && rowId === jobId) ||
          (jobName && rowName && rowName.indexOf(jobName.toLowerCase()) >= 0)) {
        return rowDriveId;
      }
    }
  } catch(e) {
    Logger.log('_findChantierDriveId error: ' + e);
  }
  return null;
}

/**
 * Crée un Google Doc de rapport dans 📁 Heures du chantier Drive.
 * Fallback: 📁 03 - RH & Paie / Heures Ponctuels si le chantier n'est pas lié.
 *
 * data: { empId, empName, empPrenom, jobId, jobName,
 *         date, punchIn, punchOut, heures, pauseMin,
 *         notes, rapport:{travaux,problemes,materiaux} }
 */
function createPunchReport(data) {
  // 1. Trouver le dossier Heures (driveId direct en priorité)
  var driveId = data.driveId || _findChantierDriveId(data.jobId, data.jobName);
  var heuresFolder;

  if (driveId) {
    heuresFolder = getChantierSubfolder(driveId, 'Heures');
  } else {
    var root     = _driveRoot();
    var rhFolder = _getOrCreateFolder(root, DRIVE_FOLDER_NAMES.RH_PAIE);
    heuresFolder = _getOrCreateFolder(rhFolder, 'Heures Ponctuels');
  }

  if (!heuresFolder) throw new Error('Impossible de trouver le dossier Heures');

  // 2. Nom du document
  var prenom  = (data.empPrenom || (data.empName || 'Employe').split(' ')[0]);
  var dateStr = data.date || Utilities.formatDate(new Date(), 'America/Toronto', 'yyyy-MM-dd');
  var docName = 'Heures_' + prenom + '_' + dateStr + '_' + (data.jobId || 'CHANTIER');

  // 3. Créer le Google Doc
  var doc    = DocumentApp.create(docName);
  var docId  = doc.getId();
  var docFile = DriveApp.getFileById(docId);
  heuresFolder.addFile(docFile);
  DriveApp.getRootFolder().removeFile(docFile);

  // 4. Rédiger le contenu
  var body = doc.getBody();
  body.clear();

  var title = body.appendParagraph('Rapport de Temps — Heuréka Construction');
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('');

  body.appendParagraph('Employé  : ' + (data.empName  || ''));
  body.appendParagraph('Date     : ' + dateStr);
  body.appendParagraph('Chantier : ' + (data.jobName  || '') + (data.jobId ? '  (' + data.jobId + ')' : ''));
  body.appendParagraph('');

  body.appendParagraph('Punch In     : ' + (data.punchIn  || ''));
  body.appendParagraph('Punch Out    : ' + (data.punchOut || ''));
  body.appendParagraph('Heures nettes: ' + (data.heures   || '') + ' h');
  if (data.pauseMin) {
    body.appendParagraph('Pauses       : ' + data.pauseMin + ' min');
  }
  body.appendParagraph('');

  if (data.notes) {
    body.appendParagraph('Notes : ' + data.notes);
    body.appendParagraph('');
  }

  if (data.rapport) {
    var h2 = body.appendParagraph('Rapport de fin de journée');
    h2.setHeading(DocumentApp.ParagraphHeading.HEADING2);
    if (data.rapport.travaux)    body.appendParagraph('Travaux effectués : ' + data.rapport.travaux);
    if (data.rapport.problemes)  body.appendParagraph('Problèmes         : ' + data.rapport.problemes);
    if (data.rapport.materiaux)  body.appendParagraph('Matériaux         : ' + data.rapport.materiaux);
    body.appendParagraph('');
  }

  doc.saveAndClose();

  var docUrl = 'https://docs.google.com/document/d/' + docId;
  return { status: 'ok', docId: docId, docUrl: docUrl, docName: docName };
}

function handlePunchReport(data) {
  try {
    return createPunchReport(data);
  } catch(e) {
    Logger.log('handlePunchReport error: ' + e);
    return { status: 'error', message: e.toString() };
  }
}

// ── UPLOAD FICHIER VERS DRIVE ─────────────────────────────────

/**
 * Reçoit un fichier en base64 et le dépose dans le bon sous-dossier.
 * data: { driveId, subfolder, fileName, mimeType, content (base64) }
 */
function uploadFileToDrive(data) {
  if (!data.driveId)  throw new Error('driveId manquant');
  if (!data.content)  throw new Error('Contenu du fichier manquant');
  if (!data.fileName) throw new Error('Nom de fichier manquant');

  var targetFolder = data.subfolder
    ? getChantierSubfolder(data.driveId, data.subfolder)
    : DriveApp.getFolderById(data.driveId);
  if (!targetFolder) throw new Error('Impossible d\'accéder au dossier');

  var bytes   = Utilities.base64Decode(data.content);
  var blob    = Utilities.newBlob(bytes, data.mimeType || 'application/octet-stream', data.fileName);
  var file    = targetFolder.createFile(blob);
  var fileId  = file.getId();
  var fileUrl = 'https://drive.google.com/file/d/' + fileId + '/view';

  return { status: 'ok', fileId: fileId, fileUrl: fileUrl, fileName: data.fileName };
}

// ── DOSSIER SOUMISSION (SOU-XXXX) ────────────────────────────────────────────

/**
 * Trouve ou crée le dossier SOU d'un projet.
 * Idempotent: utilise appProperties { projetId } pour retrouver un dossier
 * existant depuis n'importe quel appareil sans toucher au Sheet.
 * data: { projetId, clientNom, ville }
 */
function handleCreateSoumissionFolder(data) {
  var projetId  = data.projetId  || '';
  var clientNom = data.clientNom || 'Client';
  var ville     = data.ville     || '';

  // 1. Chercher un dossier existant dans tout le Drive partagé
  var existing = _findSouFolderByProjetId(projetId);
  if (existing) {
    return { status:'ok', exists:true,
             driveId: existing.id,
             driveUrl: 'https://drive.google.com/drive/folders/' + existing.id,
             souId: existing.name };
  }

  // 2. Créer le nouveau dossier dans 02 - Soumissions/En cours/
  var root    = _driveRoot();
  var souRoot = _getOrCreateFolder(root, '02 - Soumissions');
  var enCours = _getOrCreateFolder(souRoot, 'En cours');
  var nextNum = _getNextSouNumero(souRoot);
  var folderName = 'SOU-' + nextNum + ' - ' + clientNom + (ville ? ' - ' + ville : '');

  var newFolder = enCours.createFolder(folderName);
  var folderId  = newFolder.getId();

  // Sous-dossiers standards
  ['Plans & Devis','Photos','Contrats & Permis','Factures & Paiements','Heures'].forEach(function(n){
    newFolder.createFolder(n);
  });

  // Tag avec projetId pour retrouver cross-device (Drive Advanced Service)
  try {
    Drive.Files.update({ appProperties: { projetId: projetId } }, folderId, null,
      { supportsAllDrives: true });
  } catch(e) {
    Logger.log('appProperties tag: ' + e);
  }

  return { status:'ok', exists:false,
           driveId: folderId,
           driveUrl: 'https://drive.google.com/drive/folders/' + folderId,
           souId: folderName };
}

/**
 * Cherche dans tout le Drive partagé un dossier tagué projetId.
 * Retourne { id, name } ou null.
 */
function _findSouFolderByProjetId(projetId) {
  if (!projetId) return null;
  try {
    var res = Drive.Files.list({
      q: "mimeType='application/vnd.google-apps.folder'" +
         " and appProperties has { key='projetId' and value='" + projetId + "' }" +
         " and trashed=false",
      supportsAllDrives:         true,
      includeItemsFromAllDrives: true,
      fields: 'files(id,name)'
    });
    return (res.files && res.files.length > 0) ? res.files[0] : null;
  } catch(e) {
    Logger.log('_findSouFolderByProjetId: ' + e);
    return null;
  }
}

function _getNextSouNumero(souRoot) {
  var max = 0;
  var subs = souRoot.getFolders();
  while (subs.hasNext()) {
    var sub = subs.next().getFolders();
    while (sub.hasNext()) {
      var m = sub.next().getName().match(/^SOU-(\d+)/);
      if (m) max = Math.max(max, parseInt(m[1]));
    }
  }
  var n = String(max + 1); while (n.length < 4) n = '0' + n;
  return n;
}

/**
 * Déplace le dossier SOU vers Gagnées ou Refusées.
 * data: { driveId, newStatut }  'gagne' | 'perdu'
 */
function handleMoveDossierSoumission(data) {
  var driveId = data.driveId;
  if (!driveId) return { status:'error', message:'driveId manquant' };
  try {
    var root    = _driveRoot();
    var souRoot = _getOrCreateFolder(root, '02 - Soumissions');
    var dest    = _getOrCreateFolder(souRoot, data.newStatut === 'gagne' ? 'Gagnées' : 'Refusées');
    var file    = Drive.Files.get(driveId, { supportsAllDrives:true, fields:'id,name,parents' });
    Drive.Files.update({}, driveId, null, {
      addParents:        dest.getId(),
      removeParents:     (file.parents||[]).join(','),
      supportsAllDrives: true
    });
    return { status:'ok', driveId:driveId,
             driveUrl: 'https://drive.google.com/drive/folders/' + driveId,
             chanId: file.name };
  } catch(e) {
    Logger.log('handleMoveDossierSoumission: ' + e);
    return { status:'error', message:e.toString() };
  }
}

function handleDeleteDocument(data) {
  try {
    if (data.fileId) DriveApp.getFileById(data.fileId).setTrashed(true);
    return { status:'ok' };
  } catch(e) {
    Logger.log('handleDeleteDocument: ' + e);
    return { status:'error', message:e.toString() };
  }
}

function handleUploadFile(data) {
  try {
    return uploadFileToDrive(data);
  } catch(e) {
    Logger.log('handleUploadFile error: ' + e);
    return { status: 'error', message: e.toString() };
  }
}

// ── RAPPORT HEBDOMADAIRE D'HEURES ────────────────────────────

/**
 * Crée un Google Doc de rapport hebdomadaire dans 📁 Heures du chantier.
 * data: { weekStart, weekEnd, jobId, jobName, driveId,
 *         punches:[{ empId, empName, date, punchIn, punchOut, heures }] }
 */
function createWeeklyReport(data) {
  // 1. Trouver ou CRÉER le dossier Drive du chantier
  var driveId = data.driveId || _findChantierDriveId(data.jobId, data.jobName);
  var newDriveId = null, newDriveUrl = null;
  var heuresFolder = null;

  if (!driveId) {
    // Aucun dossier lié — on le crée automatiquement
    try {
      var created = createChantierDriveFolder(data.jobName || 'Chantier', '', data.jobId || '');
      driveId   = created.driveId;
      newDriveId  = created.driveId;
      newDriveUrl = created.driveUrl;
      Logger.log('Dossier Drive créé auto pour ' + data.jobName + ' : ' + driveId);
    } catch(e) {
      Logger.log('Création dossier auto échouée: ' + e);
    }
  }

  if (driveId) {
    // Synchroniser le Drive_ID dans le Sheet si pas déjà fait
    if (!newDriveId) {
      _updateChantierDriveId(data.jobId, driveId, 'https://drive.google.com/drive/folders/'+driveId);
    }
    try { heuresFolder = getChantierSubfolder(driveId, 'Heures'); } catch(e) {
      Logger.log('getChantierSubfolder: ' + e);
    }
  }

  // Fallback ultime : dossier RH & Paie
  if (!heuresFolder) {
    try {
      var root = _driveRoot();
      var rhFolder = _getOrCreateFolder(root, DRIVE_FOLDER_NAMES.RH_PAIE);
      heuresFolder = _getOrCreateFolder(rhFolder, 'Heures Ponctuels');
    } catch(e) {
      Logger.log('RH fallback folder: ' + e);
    }
  }

  // 2. Créer le doc (toujours dans Mon Drive d'abord)
  var docName = 'Heures_Semaine_' + (data.weekStart||'') + '_' + (data.jobName||data.jobId||'CHANTIER');
  var doc = DocumentApp.create(docName);
  var docId = doc.getId();

  var body = doc.getBody();
  body.clear();

  var titleP = body.appendParagraph('Rapport d\'heures hebdomadaire — Heuréka Construction');
  titleP.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph('');
  body.appendParagraph('Chantier : ' + (data.jobName || ''));
  body.appendParagraph('Semaine  : ' + (data.weekStart||'') + ' au ' + (data.weekEnd||''));
  body.appendParagraph('Approuvé le : ' + Utilities.formatDate(new Date(),'America/Toronto','yyyy-MM-dd HH:mm'));
  body.appendParagraph('');

  // Grouper par employé, colonnes = dates uniques triées
  var empMap = {}, allDates = [];
  (data.punches||[]).forEach(function(p) {
    var key = p.empId || p.empName;
    if (!empMap[key]) empMap[key] = { name: p.empName, days: {}, total: 0 };
    var hrs = parseFloat(p.heures) || 0;
    empMap[key].days[p.date] = (empMap[key].days[p.date]||0) + hrs;
    empMap[key].total += hrs;
    if (allDates.indexOf(p.date) < 0) allDates.push(p.date);
  });
  allDates.sort();

  var empKeys = Object.keys(empMap);
  var numRows = empKeys.length + 2; // header + lignes + total
  var numCols = allDates.length + 2; // Nom + dates + Total
  var table = body.appendTable();

  // En-tête
  var hRow = table.appendTableRow();
  var hCell0 = hRow.appendTableCell('Employé');
  hCell0.setBackgroundColor('#D4AF37');
  allDates.forEach(function(d) {
    var c = hRow.appendTableCell(d.slice(5).replace('-','/'));
    c.setBackgroundColor('#D4AF37');
  });
  var hCellT = hRow.appendTableCell('Total');
  hCellT.setBackgroundColor('#D4AF37');

  // Lignes employés
  var dayTotals = {}, grandTotal = 0;
  allDates.forEach(function(d){ dayTotals[d]=0; });

  empKeys.forEach(function(key) {
    var emp = empMap[key];
    var row = table.appendTableRow();
    row.appendTableCell(emp.name);
    allDates.forEach(function(d) {
      var hrs = emp.days[d] || 0;
      dayTotals[d] += hrs;
      row.appendTableCell(hrs > 0 ? hrs.toFixed(2)+'h' : '—');
    });
    row.appendTableCell(emp.total.toFixed(2)+'h');
    grandTotal += emp.total;
  });

  // Ligne total
  var tRow = table.appendTableRow();
  var tCell0 = tRow.appendTableCell('TOTAL');
  tCell0.setBackgroundColor('#f0f0f0');
  allDates.forEach(function(d) {
    var c = tRow.appendTableCell(dayTotals[d]>0?dayTotals[d].toFixed(2)+'h':'—');
    c.setBackgroundColor('#f0f0f0');
  });
  var tCellT = tRow.appendTableCell(grandTotal.toFixed(2)+'h');
  tCellT.setBackgroundColor('#f0f0f0');

  doc.saveAndClose();

  // 3. Déplacer vers le dossier cible si disponible, sinon reste dans Mon Drive
  var inFolder = 'Mon Drive (aucun dossier chantier lié)';
  if (heuresFolder) {
    try {
      var docFile = DriveApp.getFileById(docId);
      heuresFolder.addFile(docFile);
      try { DriveApp.getRootFolder().removeFile(docFile); } catch(e) {}
      inFolder = heuresFolder.getName();
    } catch(e) {
      Logger.log('Déplacement Drive: ' + e);
    }
  }

  var docUrl = 'https://docs.google.com/document/d/' + docId;
  return { docId:docId, docUrl:docUrl, docName:docName, inFolder:inFolder, newDriveId:newDriveId, newDriveUrl:newDriveUrl };
}

function handleApproveWeekPunchs(data) {
  try {
    var reports = [], errors = [];
    (data.chantiers||[]).forEach(function(ch) {
      if (!ch.punches || !ch.punches.length) return;
      try {
        var r = createWeeklyReport({
          weekStart: data.weekStart, weekEnd: data.weekEnd,
          jobId: ch.jobId, jobName: ch.jobName, driveId: ch.driveId,
          punches: ch.punches
        });
        reports.push({ jobId:ch.jobId, jobName:ch.jobName, docUrl:r.docUrl, docName:r.docName, inFolder:r.inFolder, newDriveId:r.newDriveId||null, newDriveUrl:r.newDriveUrl||null });
      } catch(e) {
        Logger.log('createWeeklyReport error for '+ch.jobName+': '+e);
        errors.push(ch.jobName + ': ' + e.toString());
      }
    });
    return { status:'ok', reports:reports, errors:errors };
  } catch(e) {
    Logger.log('handleApproveWeekPunchs: '+e);
    return { status:'error', message:e.toString() };
  }
}

// ── SYNC DRIVE IDs DEPUIS LOCALSTORAGE VERS LE SHEET ─────────

function handleSyncDriveIds(data) {
  try {
    var pairs = data.pairs || [];
    var updated = 0;
    pairs.forEach(function(p) {
      if (p.jobId && p.driveId) {
        _updateChantierDriveId(p.jobId, p.driveId, p.driveUrl || '');
        updated++;
      }
    });
    return { status: 'ok', updated: updated };
  } catch(e) {
    Logger.log('handleSyncDriveIds: ' + e);
    return { status: 'error', message: e.toString() };
  }
}

// ── LISTE DES FICHIERS D'UN CHANTIER ─────────────────────────

function listChantierFiles(driveId) {
  var result = [];
  try {
    var chanFolder = DriveApp.getFolderById(driveId);
    function scan(folder, prefix) {
      var fi = folder.getFiles();
      while (fi.hasNext()) {
        var f = fi.next();
        result.push({
          fileId:       f.getId(),
          fileName:     f.getName(),
          subfolder:    prefix,
          mimeType:     f.getMimeType(),
          viewUrl:      'https://drive.google.com/file/d/' + f.getId() + '/view',
          modifiedDate: Utilities.formatDate(f.getLastUpdated(), 'America/Toronto', 'yyyy-MM-dd')
        });
      }
      var si = folder.getFolders();
      while (si.hasNext()) {
        var sub = si.next();
        scan(sub, sub.getName());
      }
    }
    scan(chanFolder, '');
  } catch(e) {
    Logger.log('listChantierFiles: ' + e);
  }
  return result;
}

function handleGetChantierFiles(data) {
  try {
    var driveId = data.driveId || '';
    if (!driveId) return { status: 'error', message: 'driveId manquant' };
    return { status: 'ok', files: listChantierFiles(driveId) };
  } catch(e) {
    return { status: 'error', message: e.toString() };
  }
}

// ── LISTE DES PHOTOS D'UN CHANTIER (pour app Marketing) ──────

/**
 * Retourne la liste des images dans 📁 Photos d'un dossier chantier.
 * data: { driveId }
 */
function listChantierPhotos(driveId) {
  try {
    var chanFolder = DriveApp.getFolderById(driveId);
    var it = chanFolder.getFoldersByName('Photos');
    if (!it.hasNext()) return [];
    var photosFolder = it.next();
    var photos = [];
    var fIt = photosFolder.getFiles();
    while (fIt.hasNext()) {
      var f = fIt.next();
      if (f.getMimeType().startsWith('image/')) {
        photos.push({
          fileId:       f.getId(),
          fileName:     f.getName(),
          viewUrl:      'https://drive.google.com/file/d/' + f.getId() + '/view',
          thumbnailUrl: 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w300'
        });
      }
    }
    return photos;
  } catch(e) {
    Logger.log('listChantierPhotos error: ' + e);
    return [];
  }
}

function handleGetChantierPhotos(data) {
  try {
    var driveId = data.driveId || '';
    if (!driveId) return { status: 'error', message: 'driveId manquant' };
    var photos = listChantierPhotos(driveId);
    return { status: 'ok', photos: photos };
  } catch(e) {
    Logger.log('handleGetChantierPhotos error: ' + e);
    return { status: 'error', message: e.toString() };
  }
}
