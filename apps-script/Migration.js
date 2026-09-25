/*************************************************************
 *  MIGRATION « BACKOFFICE V1 » — étape 2 : créateurs
 *   apercuEtape2()   : liste ce qui serait fait, n'écrit RIEN
 *   etape2Createurs(): crée `createurs`, `categories` et `_correspondance_noms`
 *  Ne modifie ni ne supprime aucun onglet existant.
 *  `_correspondance_noms` : chaque nom trouvé dans les ventes → id_createur.
 *  Les filles vérifient les lignes « À VÉRIFIER » et corrigent id_createur
 *  si besoin, AVANT l'étape 3 (construction de la nouvelle table ventes).
 *************************************************************/

const SHEET_CREATEURS_V1  = 'createurs';
const SHEET_CATEGORIES    = 'categories';
const SHEET_CORRESPONDANCE = '_correspondance_noms';

const COLONNES_CREATEURS_V1 = [
  'id_createur', 'nom', 'email', 'statut', 'categorie', 'nom_legal', 'telephone', 'adresse',
  'code_postal', 'ville', 'siret', 'iban', 'benevole', 'perms_prevues', 'rc_pro',
  'adhesion_payee_le', 'cree_le', 'modifie_le'
];
const COLONNES_CATEGORIES = ['code', 'libelle', 'exemples', 'ordre', 'actif'];
const CATEGORIES_INITIALES = [
  ['bijoux',       'Bijoux',               'boucles d\'oreilles, colliers, bagues'],
  ['accessoires',  'Accessoires',          'sacs, pochettes, chouchous, maroquinerie'],
  ['illustration', 'Illustration',         'affiches, prints, cartes postales'],
  ['papeterie',    'Papeterie & stickers', 'carnets, stickers, marque-pages'],
  ['ceramique',    'Céramique',            'tasses, bols, porcelaine'],
  ['deco',         'Déco & maison',        'bougies, suncatchers, objets'],
  ['textile',      'Textile & mode',       'vêtements de créateur, tricot'],
  ['friperie',     'Friperie',             'seconde main, vintage'],
  ['beaute',       'Beauté & bien-être',   'savons, cosmétiques'],
  ['enfants',      'Enfants',              'jouets, doudous'],
  ['art',          'Art',                  'peintures, gravures, photos originales']
].map(function (c, i) { return [c[0], c[1], c[2], i + 1, true]; });
const COLONNES_CORRESPONDANCE = [
  'nom_dans_les_ventes', 'nb_ventes', 'premiere_vente', 'derniere_vente',
  'id_createur', 'nom_createur', 'methode', 'remarque'
];

function apercuEtape2() { _etape2(true); }
function etape2Createurs() { _etape2(false); }

/** Clé de rapprochement : casse, espaces et caractères invisibles ignorés. */
function _cleNom(s) { return _nomPropre(s).toLowerCase(); }
/** Clé « large » : ignore aussi accents et ponctuation (pour signaler, jamais pour fusionner). */
function _cleLarge(s) { return _cleNom(s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]/g, ''); }

function _levenshtein(a, b) {
  const d = [];
  for (let i = 0; i <= a.length; i++) d[i] = [i];
  for (let j = 0; j <= b.length; j++) d[0][j] = j;
  for (let i = 1; i <= a.length; i++)
    for (let j = 1; j <= b.length; j++)
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
  return d[a.length][b.length];
}

/** Lit un onglet en objets {entête: valeur}. */
function _lireOnglet(sh) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(_norm);
  return sh.getRange(2, 1, n, h.length).getValues().map(function (r) {
    const o = {}; h.forEach(function (k, i) { if (k) o[k] = r[i]; }); return o;
  });
}

function _etape2(apercu) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bilan = [];
  const dire = function (s) { bilan.push((apercu ? '[aperçu] ' : '') + s); };

  [SHEET_CREATEURS_V1, SHEET_CATEGORIES, SHEET_CORRESPONDANCE].forEach(function (nom) {
    if (ss.getSheetByName(nom)) throw new Error("L'onglet « " + nom + " » existe déjà : étape 2 déjà faite ? Supprime-le pour la relancer.");
  });

  // --- 1) Créateurs connus : onglet vendeurs ---
  const vendeurs = _lireOnglet(ss.getSheetByName(SHEET_CREATEURS));
  const createurs = [];                 // {id, nom, email, statut, cree_le, cles:Set}
  const parCle = {};
  let n = 0;
  const nouvelId = function () { return 'C' + String(++n).padStart(3, '0'); };
  let doublonsVendeurs = 0;
  vendeurs.forEach(function (v) {
    const nom = _nomPropre(v[COL_NOM_CREATEUR]);
    if (!nom) return;
    const cle = _cleNom(nom);
    const statut = _norm(v['status']) === 'actif' ? 'actif' : 'inactif';
    if (parCle[cle]) {                  // même créateur listé deux fois : on garde une fiche
      doublonsVendeurs++;
      const c = parCle[cle];
      if (statut === 'actif') c.statut = 'actif';
      if (!c.email && v['email_vendeur']) c.email = String(v['email_vendeur']).trim();
      return;
    }
    const c = { id: nouvelId(), nom: nom, email: String(v['email_vendeur'] || '').trim(), statut: statut,
                cree_le: v['date_de_modifcation'] instanceof Date ? v['date_de_modifcation'] : '' };
    createurs.push(c); parCle[cle] = c;
  });
  dire('createurs : ' + createurs.length + " créateurs repris de l'onglet « " + SHEET_CREATEURS + ' »' + (doublonsVendeurs ? ' (' + doublonsVendeurs + ' ligne(s) en double fusionnée(s))' : '') + '.');

  // --- 2) Tous les noms présents dans les ventes (historique + ventes actuelles) ---
  const toutes = _lireOnglet(ss.getSheetByName(_sheetData()));
  const noms = {};                      // nom brut → {n, premiere, derniere}
  toutes.forEach(function (l) {
    const brut = String(l[COL_CREATEUR] == null ? '' : l[COL_CREATEUR]);
    if (!brut.trim() || !(l['date'] instanceof Date)) return;
    const s = noms[brut] || (noms[brut] = { n: 0, premiere: l['date'], derniere: l['date'] });
    s.n++;
    if (l['date'] < s.premiere) s.premiere = l['date'];
    if (l['date'] > s.derniere) s.derniere = l['date'];
  });

  // Un nom qui ne diffère d'un créateur connu que par les accents ou la ponctuation
  // (« LSP CREATION » / « LSP création ») est rattaché à ce créateur.
  const parCleLarge = {};
  createurs.forEach(function (c) { const k = _cleLarge(c.nom); if (k && !parCleLarge[k]) parCleLarge[k] = c; });
  const trouver = function (brut) { return parCle[_cleNom(brut)] || parCleLarge[_cleLarge(brut)] || null; };

  // Créateurs présents uniquement dans les ventes : une fiche « inactif » par clé large,
  // au nom le plus utilisé, dans l'ordre de leur première vente.
  const groupes = {};
  Object.keys(noms).forEach(function (brut) {
    if (trouver(brut)) return;
    const cle = _cleLarge(brut) || _cleNom(brut);
    const g = groupes[cle] || (groupes[cle] = { variantes: [], premiere: noms[brut].premiere, n: 0 });
    g.variantes.push(brut); g.n += noms[brut].n;
    if (noms[brut].premiere < g.premiere) g.premiere = noms[brut].premiere;
  });
  const connus = createurs.slice();
  Object.keys(groupes).sort(function (a, b) { return groupes[a].premiere - groupes[b].premiere; }).forEach(function (cle) {
    const g = groupes[cle];
    const nom = _nomPropre(g.variantes.sort(function (a, b) { return noms[b].n - noms[a].n; })[0]);
    const c = { id: nouvelId(), nom: nom, email: '', statut: 'inactif', cree_le: g.premiere, nouveau: true };
    createurs.push(c);
    g.variantes.forEach(function (v) { parCle[_cleNom(v)] = c; });
  });
  dire('createurs : ' + Object.keys(groupes).length + ' créateur(s) présents seulement dans les ventes, ajoutés en « inactif ».');

  // --- 3) Correspondance nom → id, avec signalement des cas douteux ---
  const lignes = [];
  let nbVerif = 0;
  Object.keys(noms).sort(function (a, b) { return _cleNom(a).localeCompare(_cleNom(b), 'fr'); }).forEach(function (brut) {
    const c = trouver(brut);
    let methode, remarque = '';
    if (_nomPropre(brut) === c.nom) methode = c.nouveau ? '' : 'identique';
    else if (_cleNom(brut) === _cleNom(c.nom)) methode = 'variante (majuscules / espaces)';
    else if (_cleLarge(brut) === _cleLarge(c.nom)) methode = 'variante (accents / ponctuation)';
    if (!methode && c.nouveau) {
      const large = _cleLarge(brut);
      const proches = connus.filter(function (k) {
        const kl = _cleLarge(k.nom);
        if (!kl || !large) return false;
        if (kl === large) return true;
        if (kl.length >= 4 && large.length >= 4 && (large.indexOf(kl) !== -1 || kl.indexOf(large) !== -1)) return true;
        return Math.min(kl.length, large.length) >= 6 && _levenshtein(kl, large) <= 2;
      });
      if (proches.length) {
        methode = 'À VÉRIFIER';
        remarque = 'Ressemble à : ' + proches.slice(0, 3).map(function (k) { return k.nom + ' (' + k.id + ')'; }).join(', ') +
          '. Si c\'est la même personne, remplace ' + c.id + ' par son identifiant.';
        nbVerif++;
      } else {
        methode = 'nouveau créateur (ancien, inactif)';
      }
    }
    lignes.push([brut, noms[brut].n, noms[brut].premiere, noms[brut].derniere, c.id, c.nom, methode, remarque]);
  });
  dire(SHEET_CORRESPONDANCE + ' : ' + lignes.length + ' noms trouvés dans ' + toutes.length + ' ventes, dont ' + nbVerif + ' à vérifier par les filles.');
  dire(SHEET_CATEGORIES + ' : ' + CATEGORIES_INITIALES.length + ' catégories proposées.');

  if (!apercu) {
    const sC = _nouvelOnglet(ss, SHEET_CREATEURS_V1, COLONNES_CREATEURS_V1, createurs.map(function (c) {
      return [c.id, c.nom, c.email, c.statut, '', '', '', '', '', '', '', '', false, '', false, '', c.cree_le, new Date()];
    }));
    const col = function (nom) { return COLONNES_CREATEURS_V1.indexOf(nom) + 1; };
    const nb = createurs.length;
    ['telephone', 'code_postal', 'siret', 'iban'].forEach(function (k) { sC.getRange(2, col(k), nb, 1).setNumberFormat('@'); });
    ['benevole', 'rc_pro'].forEach(function (k) { sC.getRange(2, col(k), nb, 1).insertCheckboxes(); });
    ['adhesion_payee_le', 'cree_le', 'modifie_le'].forEach(function (k) { sC.getRange(2, col(k), nb, 1).setNumberFormat('dd/mm/yyyy'); });
    sC.getRange(2, col('statut'), nb, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['actif', 'inactif'], true).build());

    _nouvelOnglet(ss, SHEET_CATEGORIES, COLONNES_CATEGORIES, CATEGORIES_INITIALES)
      .getRange(2, 5, CATEGORIES_INITIALES.length, 1).insertCheckboxes();
    sC.getRange(2, col('categorie'), nb, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInRange(ss.getSheetByName(SHEET_CATEGORIES).getRange('A2:A200'), true).setAllowInvalid(false).build());

    const sM = _nouvelOnglet(ss, SHEET_CORRESPONDANCE, COLONNES_CORRESPONDANCE, lignes);
    sM.getRange(2, 3, lignes.length, 2).setNumberFormat('dd/mm/yyyy');
    sM.getRange(2, 5, lignes.length, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInRange(sC.getRange('A2:A' + (nb + 1)), true).setAllowInvalid(false).build());
    lignes.forEach(function (l, i) { if (l[6] === 'À VÉRIFIER') sM.getRange(i + 2, 1, 1, COLONNES_CORRESPONDANCE.length).setBackground('#fff3cd'); });
    sM.getRange(1, 5).setNote('Modifiable : si un nom correspond en réalité à un autre créateur, choisis son identifiant dans la liste.');

    const j = ss.getSheetByName(SHEET_JOURNAL) || _nouvelOnglet(ss, SHEET_JOURNAL, ['date', 'qui', 'action', 'detail'], []);
    j.appendRow([new Date(), Session.getEffectiveUser().getEmail(), 'migration_etape2', bilan.join(' | ')]);
  }
  Logger.log((apercu ? "APERÇU — rien n'a été écrit.\n" : '✅ Étape 2 terminée.\n') + bilan.map(function (l) { return '• ' + l; }).join('\n'));
}

function _nouvelOnglet(ss, nom, entetes, lignes) {
  const sh = ss.insertSheet(nom);
  sh.getRange(1, 1, 1, entetes.length).setValues([entetes]).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (lignes.length) sh.getRange(2, 1, lignes.length, entetes.length).setValues(lignes);
  sh.autoResizeColumns(1, entetes.length);
  return sh;
}
