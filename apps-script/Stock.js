/*************************************************************
 *  STOCK — étape 0 : Google Sheet de reprise du stock existant
 *  Fichier séparé « 13H59 · Reprise du stock », rempli à la main par
 *  l'équipe (créateur par créateur) le temps que les créateurs déclarent
 *  leur stock en ligne. Le site l'importera ensuite (produits, codes-barres,
 *  stock initial, étiquettes). Listes déroulantes tirées des onglets
 *  createurs et categories pour que l'import soit sans surprise.
 *  Rangé dans 13h59 ▸ Stock ; son identifiant est gardé (STOCK_REPRISE_ID).
 *************************************************************/

const COLONNES_REPRISE_STOCK = ['Créateur', 'Nom du produit', 'Référence du créateur (facultatif)', 'Catégorie', 'Prix de vente (€)', 'Quantité en boutique', 'Étiquette', 'Remarque'];
const ETIQUETTE_REPRISE = ['Étiquette', 'Petit article (planche à la caisse)'];

/** Libellés « Nom · C012 » des créateurs actifs, triés. */
function _createursPourListe(ss) {
  return _lireTable(_onglet(ss, SHEET_CREATEURS)).filter(function (c) { return c['id_createur'] && _norm(c['statut']) === 'actif'; })
    .map(function (c) { return _nomPropre(c['nom']) + ' · ' + c['id_createur']; })
    .sort(function (a, b) { return a.localeCompare(b, 'fr', { sensitivity: 'base' }); });
}

/** À lancer une fois depuis l'éditeur : crée (ou retrouve) la feuille de reprise et affiche son lien. */
function creerFeuilleRepriseStock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), props = PropertiesService.getScriptProperties();
  const existant = props.getProperty('STOCK_REPRISE_ID');
  if (existant) {
    try { const f = DriveApp.getFileById(existant); if (!f.isTrashed()) { Logger.log('La feuille existe déjà : ' + f.getUrl() + '\nPour mettre à jour les listes de créateurs et de catégories : majListesRepriseStock()'); return; } }
    catch (e) { /* supprimée : on recrée */ }
  }
  const r = SpreadsheetApp.create('13H59 · Reprise du stock');
  DriveApp.getFileById(r.getId()).moveTo(_sousDossier(_sousDossier(DriveApp.getRootFolder(), '13h59'), 'Stock'));
  props.setProperty('STOCK_REPRISE_ID', r.getId());

  const sh = r.getSheets()[0].setName('Stock');
  sh.getRange(1, 1, 1, COLONNES_REPRISE_STOCK.length).setValues([COLONNES_REPRISE_STOCK]).setFontWeight('bold').setBackground('#0f0f0f').setFontColor('#ffffff').setWrap(true);
  sh.setFrozenRows(1);
  [230, 240, 150, 150, 110, 110, 190, 240].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  const n = 2000;
  sh.getRange(2, 5, n, 1).setNumberFormat('0.00');
  sh.getRange(2, 6, n, 1).setNumberFormat('0');
  sh.getRange(2, 3, n, 1).setNumberFormat('@');
  sh.getRange(2, 5, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireNumberGreaterThan(0).setAllowInvalid(false).setHelpText('Prix de vente en euros, par exemple 12,50').build());
  sh.getRange(2, 6, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).setHelpText('Nombre d\'exemplaires en boutique').build());
  sh.getRange(2, 7, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ETIQUETTE_REPRISE, true).setAllowInvalid(false).build());
  sh.getRange(1, 1, 1, COLONNES_REPRISE_STOCK.length).protect().setWarningOnly(true).setDescription('En-têtes lus par le site : ne pas les modifier');

  const aide = r.insertSheet('Mode d\'emploi', 0);
  aide.getRange(1, 1, 12, 1).setValues([
    ['REPRISE DU STOCK · MODE D\'EMPLOI'],
    [''],
    ['1. Dans l\'onglet « Stock », une ligne par produit : choisis le créateur et la catégorie dans les listes.'],
    ['2. Nom du produit : court et clair, il sera imprimé sur l\'étiquette (par exemple « Boucles d\'oreilles lune »).'],
    ['3. Référence du créateur : facultative, seulement s\'il en utilise une.'],
    ['4. Prix de vente : le prix affiché en boutique, en euros.'],
    ['5. Quantité : le nombre d\'exemplaires présents en boutique.'],
    ['6. Étiquette : « Étiquette » pour un article qu\'on étiquette, « Petit article » pour ce qui est trop petit (bijoux) : son code-barres sera sur une planche à côté de la caisse.'],
    ['7. Un même produit en plusieurs exemplaires = une seule ligne avec la quantité.'],
    ['8. Ne modifie pas la première ligne de l\'onglet « Stock » (les titres des colonnes).'],
    [''],
    ['Quand un créateur est fini, préviens Rémi : le site importera ses lignes, créera les codes-barres et imprimera les étiquettes.']
  ]);
  aide.getRange(1, 1).setFontWeight('bold').setFontSize(13);
  aide.setColumnWidth(1, 900);
  aide.getRange(1, 1, 12, 1).setWrap(true);
  const listes = r.insertSheet('Listes');
  listes.hideSheet();
  _majListesReprise(ss, r);
  r.setActiveSheet(sh);
  Logger.log('✅ Feuille créée dans 13h59 ▸ Stock : ' + r.getUrl() + '\nPartage-la avec Mo, Caro et Solenne (bouton Partager, en modification).');
}

/** Met à jour les listes déroulantes (créateurs actifs, catégories) de la feuille de reprise. */
function majListesRepriseStock() {
  const id = PropertiesService.getScriptProperties().getProperty('STOCK_REPRISE_ID');
  if (!id) { Logger.log('Pas encore de feuille : lance creerFeuilleRepriseStock()'); return; }
  _majListesReprise(SpreadsheetApp.getActiveSpreadsheet(), SpreadsheetApp.openById(id));
  Logger.log('✅ Listes mises à jour.');
}
function _majListesReprise(ss, r) {
  const createurs = _createursPourListe(ss);
  const categories = _categoriesTriees(ss).filter(function (c) { return c['actif'] !== false; }).map(function (c) { return String(c['libelle'] || c['code']); });
  const listes = r.getSheetByName('Listes'), sh = r.getSheetByName('Stock');
  listes.clear();
  listes.getRange(1, 1, 1, 2).setValues([['Créateurs', 'Catégories']]);
  if (createurs.length) listes.getRange(2, 1, createurs.length, 1).setValues(createurs.map(function (x) { return [x]; }));
  if (categories.length) listes.getRange(2, 2, categories.length, 1).setValues(categories.map(function (x) { return [x]; }));
  const n = 2000;
  sh.getRange(2, 1, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(listes.getRange(2, 1, Math.max(1, createurs.length), 1), true).setAllowInvalid(false).build());
  sh.getRange(2, 4, n, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(listes.getRange(2, 2, Math.max(1, categories.length), 1), true).setAllowInvalid(false).build());
}
