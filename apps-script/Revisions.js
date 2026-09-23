/*************************************************************
 *  DIAGNOSTIC — export des révisions du classeur (lecture seule)
 *
 *  But : retrouver les doublons de la caisse qui ont été supprimés
 *  à la main. Chaque révision enregistrée par Google depuis la date
 *  DEPUIS est exportée en .xlsx dans le dossier Drive « 13h59_revisions ».
 *  Le Sheet n'est JAMAIS modifié.
 *
 *  Lancer exporterRevisions() depuis l'éditeur (▶ Exécuter).
 *  Si le journal indique « À relancer », relancer : les fichiers déjà
 *  exportés sont sautés (limite Apps Script de 6 min par exécution).
 *
 *  ⚠️ Pas de `const` global ici (scope partagé entre fichiers .gs).
 *************************************************************/

function exporterRevisions() {
  var DEPUIS = new Date('2026-09-11T00:00:00+02:00');   // mise en service caisse web
  var NOM_DOSSIER = '13h59_revisions';
  var BUDGET_MS = 4.5 * 60 * 1000;
  var debut = Date.now();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var id = ss.getId();
  var opts = { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };

  // 1) Liste des révisions (API Drive v3, appelée en REST : pas de service avancé à activer).
  var revs = [], page = '';
  do {
    var url = 'https://www.googleapis.com/drive/v3/files/' + id + '/revisions?pageSize=200'
      + '&fields=nextPageToken,revisions(id,modifiedTime,lastModifyingUser/displayName,exportLinks)'
      + (page ? '&pageToken=' + encodeURIComponent(page) : '');
    var j = JSON.parse(UrlFetchApp.fetch(url, opts).getContentText());
    if (j.error) throw new Error('Drive API : ' + j.error.message);
    revs = revs.concat(j.revisions || []);
    page = j.nextPageToken || '';
  } while (page);

  var cibles = revs.filter(function (v) { return new Date(v.modifiedTime) >= DEPUIS; });
  Logger.log(revs.length + ' révision(s) au total, ' + cibles.length + ' depuis ' + DEPUIS.toISOString());

  // 2) Dossier de sortie (réutilisé d'une exécution à l'autre).
  var it = DriveApp.getFoldersByName(NOM_DOSSIER);
  var dossier = it.hasNext() ? it.next() : DriveApp.createFolder(NOM_DOSSIER);
  var deja = {};
  var fi = dossier.getFiles();
  while (fi.hasNext()) deja[fi.next().getName()] = 1;

  // 3) Export de chaque révision. Le nom contient la date et l'auteur :
  //    les écritures de la caisse apparaissent sous TON nom (le script
  //    s'exécute en tant que toi), les suppressions manuelles sous celui
  //    de la personne qui les a faites.
  var faits = 0, sautes = 0;
  for (var i = 0; i < cibles.length; i++) {
    if (Date.now() - debut > BUDGET_MS) {
      Logger.log('⏸ Budget de temps atteint — À relancer (' + (cibles.length - i) + ' restante(s)).');
      break;
    }
    var v = cibles[i];
    var auteur = (v.lastModifyingUser && v.lastModifyingUser.displayName || 'inconnu').replace(/[^\w\- ]/g, '');
    var nom = 'rev_' + v.modifiedTime.replace(/[:.]/g, '-') + '_' + v.id + '_' + auteur + '.xlsx';
    if (deja[nom]) { sautes++; continue; }
    var lien = v.exportLinks && v.exportLinks['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
    if (!lien) { Logger.log('Pas d\'export pour la révision ' + v.id); continue; }
    var rep = UrlFetchApp.fetch(lien, opts);
    if (rep.getResponseCode() !== 200) {
      Logger.log('Échec révision ' + v.id + ' : HTTP ' + rep.getResponseCode());
      continue;
    }
    dossier.createFile(rep.getBlob().setName(nom));
    faits++;
  }
  Logger.log('✅ ' + faits + ' exportée(s), ' + sautes + ' déjà présente(s). Dossier : ' + dossier.getUrl());
}