/*************************************************************
 *  MIGRATION « BACKOFFICE V1 » — étape 2 : créateurs
 *   apercuEtape2()   : liste ce qui serait fait, n'écrit RIEN
 *   etape2Createurs(): crée `createurs`, `categories` et `_correspondance_noms`
 *  Ne modifie ni ne supprime aucun onglet existant.
 *  `_correspondance_noms` : chaque nom trouvé dans les ventes → id_createur.
 *  Les filles vérifient les lignes « À VÉRIFIER » et corrigent id_createur
 *  si besoin, AVANT l'étape 3 (construction de la nouvelle table ventes).
 *************************************************************/

// Noms d'AVANT la bascule (l'ancien Sheet) : utilisés uniquement par la migration.
const ANCIEN_VENDEURS       = 'vendeurs';
const ANCIEN_ALL_DATA       = 'All data';
const ANCIEN_VENTES         = 'ventes';
const ANCIEN_VENTES_ARCHIVE = '_ancien_ventes';
const ANCIEN_PAIEMENTS      = 'taxe_par_type_de_paiement';
const ANCIEN_COL_VENDEUR    = 'vendeur';
const ANCIEN_COL_NOM        = 'nom_vendeur';
const ANCIEN_COL_PRIME      = 'prime_vendeur';

const SHEET_CREATEURS_V1  = 'createurs';
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

function _etape2(apercu) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bilan = [];
  const dire = function (s) { bilan.push((apercu ? '[aperçu] ' : '') + s); };

  [SHEET_CREATEURS_V1, SHEET_CATEGORIES, SHEET_CORRESPONDANCE].forEach(function (nom) {
    if (ss.getSheetByName(nom)) throw new Error("L'onglet « " + nom + " » existe déjà : étape 2 déjà faite ? Supprime-le pour la relancer.");
  });

  // --- 1) Créateurs connus : onglet vendeurs ---
  const vendeurs = _lireTable(ss.getSheetByName(ANCIEN_VENDEURS));
  const createurs = [];                 // {id, nom, email, statut, cree_le, cles:Set}
  const parCle = {};
  let n = 0;
  const nouvelId = function () { return 'C' + String(++n).padStart(3, '0'); };
  let doublonsVendeurs = 0;
  vendeurs.forEach(function (v) {
    const nom = _nomPropre(v[ANCIEN_COL_NOM]);
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
  dire('createurs : ' + createurs.length + " créateurs repris de l'onglet « " + ANCIEN_VENDEURS + ' »' + (doublonsVendeurs ? ' (' + doublonsVendeurs + ' ligne(s) en double fusionnée(s))' : '') + '.');

  // --- 2) Tous les noms présents dans les ventes (historique + ventes actuelles) ---
  const toutes = _lireTable(ss.getSheetByName(ANCIEN_ALL_DATA));
  const noms = {};                      // nom brut → {n, premiere, derniere}
  toutes.forEach(function (l) {
    const brut = String(l[ANCIEN_COL_VENDEUR] == null ? '' : l[ANCIEN_COL_VENDEUR]);
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

/*************************************************************
 *  MIGRATION « BACKOFFICE V1 » — étape 3 : ventes
 *   apercuEtape3() : contrôles seuls, n'écrit RIEN
 *   etape3Ventes() : (re)construit `ventes_v1` et `_controle_migration`
 *  Source : l'historique figé de `All data` (lignes au-dessus de la formule
 *  FILTER) + l'onglet `ventes` actuel. Ne modifie aucun onglet existant ;
 *  relançable (les deux onglets qu'elle crée sont reconstruits à chaque fois).
 *************************************************************/

const SHEET_VENTES_V1 = 'ventes_v1';
const SHEET_CONTROLE  = '_controle_migration';

function apercuEtape3() { _etape3(true); }
function etape3Ventes() { _etape3(false); }

function _etape3(apercu) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(ANCIEN_VENTES_ARCHIVE)) throw new Error('La bascule a déjà été faite : l\'étape 3 ne sert plus.');
  const bilan = [], alertes = [];
  const dire = function (s) { bilan.push(s); };

  // --- Correspondance nom → id_createur (étape 2, éventuellement corrigée par les filles) ---
  const shM = ss.getSheetByName(SHEET_CORRESPONDANCE);
  if (!shM) throw new Error("Onglet « " + SHEET_CORRESPONDANCE + " » introuvable : lance d'abord l'étape 2.");
  const idParNom = {}, idParCle = {};
  _lireTable(shM).forEach(function (m) {
    const id = String(m['id_createur'] || '').trim();
    if (!id) return;
    idParNom[String(m['nom_dans_les_ventes'])] = id;
    idParCle[_cleNom(m['nom_dans_les_ventes'])] = id;
  });
  const idsCreateurs = {};
  _lireTable(ss.getSheetByName(SHEET_CREATEURS_V1)).forEach(function (c) { idsCreateurs[String(c['id_createur'])] = true; });
  const codesRemise = {}, codesPaiement = {};
  _lireTable(ss.getSheetByName(SHEET_REMISES)).forEach(function (r) { if (r['type_de_remise']) codesRemise[String(r['type_de_remise']).trim()] = true; });
  _lireTable(ss.getSheetByName(ANCIEN_PAIEMENTS)).forEach(function (r) { if (r['type_de_paiement']) codesPaiement[_norm(r['type_de_paiement'])] = true; });

  // --- Source 1 : historique figé de All data (avant la formule FILTER) ---
  const shA = ss.getSheetByName(ANCIEN_ALL_DATA);
  const nA = shA.getLastRow() - 1;
  const formulesA = shA.getRange(2, 1, nA, 1).getFormulas();
  const idxFormule = formulesA.findIndex(function (f) { return /FILTER\s*\(/i.test(f[0]); });
  if (idxFormule < 0) throw new Error("Formule FILTER introuvable dans la colonne A de « " + ANCIEN_ALL_DATA + " ».");
  const toutesA = _lireTable(shA);
  const historique = toutesA.slice(0, idxFormule).filter(function (l) { return l['date'] instanceof Date; });
  const recopieA = toutesA.slice(idxFormule).filter(function (l) { return l['date'] instanceof Date; });

  // --- Source 2 : onglet ventes actuel ---
  const actuelles = _lireTable(ss.getSheetByName(ANCIEN_VENTES)).filter(function (l) { return l['date'] instanceof Date; });
  dire('Sources : ' + historique.length + " ventes dans l'historique de « " + ANCIEN_ALL_DATA + ' » + ' + actuelles.length + ' dans « ' + ANCIEN_VENTES + ' » = ' + (historique.length + actuelles.length) + '.');
  if (recopieA.length !== actuelles.length) alertes.push('« ' + ANCIEN_ALL_DATA + ' » ne recopie que ' + recopieA.length + ' des ' + actuelles.length + ' ventes actuelles (plage de la formule FILTER trop courte).');

  // --- Construction, dans l'ordre chronologique ---
  const num = function (v) { const x = Number(v); return isNaN(x) ? 0 : x; };
  const lignes = historique.map(function (l, i) { return { l: l, ordre: i }; })
    .concat(actuelles.map(function (l, i) { return { l: l, ordre: historique.length + i }; }))
    .sort(function (a, b) { return (a.l['date'] - b.l['date']) || (a.ordre - b.ordre); });
  const inconnus = {}, remisesInconnues = {}, paiementsInconnus = {};
  const sortie = lignes.map(function (x, i) {
    const l = x.l, nom = String(l[ANCIEN_COL_VENDEUR] == null ? '' : l[ANCIEN_COL_VENDEUR]);
    const id = idParNom[nom] || idParCle[_cleNom(nom)] || '';
    if (!id || !idsCreateurs[id]) inconnus[nom] = (inconnus[nom] || 0) + 1;
    const remise = String(l['type_de_remise'] || '').trim() || 'pas_de_remise';
    const paiement = _norm(l['type_de_paiement']);
    if (!codesRemise[remise]) remisesInconnues[remise] = (remisesInconnues[remise] || 0) + 1;
    if (!codesPaiement[paiement]) paiementsInconnus[paiement] = (paiementsInconnus[paiement] || 0) + 1;
    const panier = l['numero_panier'];
    return [i + 1, (panier === '' || panier == null) ? '' : num(panier), l['date'], id, String(l['reference_produit'] || ''),
      remise, paiement, num(l['prix']), num(l['remise']), num(l['prix_client']), num(l['taxe']), num(l[ANCIEN_COL_PRIME]),
      String(l['id_transaction'] || '')];
  });
  const cles = function (o) { return Object.keys(o).map(function (k) { return '« ' + k + ' » (' + o[k] + ')'; }).join(', '); };
  if (Object.keys(inconnus).length) alertes.push('Créateurs non rattachés : ' + cles(inconnus));
  if (Object.keys(remisesInconnues).length) alertes.push('Codes remise absents de « ' + SHEET_REMISES + ' » : ' + cles(remisesInconnues));
  if (Object.keys(paiementsInconnus).length) alertes.push('Paiements absents de « ' + ANCIEN_PAIEMENTS + ' » : ' + cles(paiementsInconnus));

  // --- Contrôle : mêmes totaux, mois par mois, avant / après ---
  const cumul = function (rows, getDate, getCa, getPrime) {
    const t = {};
    rows.forEach(function (r) {
      const d = getDate(r), m = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
      const o = t[m] || (t[m] = { n: 0, ca: 0, prime: 0 });
      o.n++; o.ca += getCa(r); o.prime += getPrime(r);
    });
    return t;
  };
  // « avant » = ce que les tableaux de bord lisent aujourd'hui (tout All data).
  const avant = cumul(historique.concat(recopieA), function (l) { return l['date']; }, function (l) { return num(l['prix_client']); }, function (l) { return num(l[ANCIEN_COL_PRIME]); });
  const apres = cumul(sortie, function (r) { return r[2]; }, function (r) { return r[9]; }, function (r) { return r[11]; });
  const r2 = function (x) { return Math.round(x * 100) / 100; };
  const mois = Object.keys(avant).concat(Object.keys(apres)).filter(function (m, i, a) { return a.indexOf(m) === i; }).sort();
  let ecarts = 0;
  const controle = mois.map(function (m) {
    const a = avant[m] || { n: 0, ca: 0, prime: 0 }, b = apres[m] || { n: 0, ca: 0, prime: 0 };
    const ok = a.n === b.n && r2(a.ca) === r2(b.ca) && r2(a.prime) === r2(b.prime);
    if (!ok) ecarts++;
    return [m, a.n, b.n, r2(a.ca), r2(b.ca), r2(a.prime), r2(b.prime), ok ? 'OK' : 'ÉCART'];
  });
  const tot = function (k, src) { return r2(Object.keys(src).reduce(function (s, m) { return s + src[m][k]; }, 0)); };
  const tz = ss.getSpreadsheetTimeZone();
  dire(SHEET_VENTES_V1 + ' : ' + sortie.length + ' ventes, id_vente 1 → ' + sortie.length + ' (uniques), du ' +
    Utilities.formatDate(sortie[0][2], tz, 'dd/MM/yyyy') + ' au ' + Utilities.formatDate(sortie[sortie.length - 1][2], tz, 'dd/MM/yyyy') + '.');
  dire('Contrôle : ' + mois.length + ' mois comparés, ' + (ecarts ? ecarts + ' ÉCART(S)' : 'tous identiques') +
    ' · CA ' + tot('ca', avant) + ' € → ' + tot('ca', apres) + ' € · primes ' + tot('prime', avant) + ' € → ' + tot('prime', apres) + ' €.');

  if (!apercu) {
    [SHEET_VENTES_V1, SHEET_CONTROLE].forEach(function (nom) { const s = ss.getSheetByName(nom); if (s) ss.deleteSheet(s); });
    const sV = _nouvelOnglet(ss, SHEET_VENTES_V1, COLONNES_VENTES, []);
    const n = sortie.length;
    sV.getRange(2, 1, n, COLONNES_VENTES.length).setValues(sortie);
    const col = function (k) { return COLONNES_VENTES.indexOf(k) + 1; };
    sV.getRange(2, col('date'), n, 1).setNumberFormat('dd/mm/yyyy hh:mm');
    sV.getRange(2, col('reference'), n, 1).setNumberFormat('@');
    sV.getRange(2, col('prix'), n, 5).setNumberFormat('0.00');
    sV.autoResizeColumns(1, COLONNES_VENTES.length);

    const sK = _nouvelOnglet(ss, SHEET_CONTROLE, ['mois', 'nb_avant', 'nb_apres', 'ca_avant', 'ca_apres', 'prime_avant', 'prime_apres', 'resultat'], controle);
    controle.forEach(function (c, i) { if (c[7] !== 'OK') sK.getRange(i + 2, 1, 1, 8).setBackground('#fbeceb'); });

    ss.getSheetByName(SHEET_JOURNAL).appendRow([new Date(), Session.getEffectiveUser().getEmail(), 'migration_etape3', bilan.concat(alertes).join(' | ')]);
  }
  Logger.log((apercu ? "APERÇU — rien n'a été écrit.\n" : '✅ Étape 3 terminée.\n') +
    bilan.map(function (l) { return '• ' + l; }).join('\n') +
    (alertes.length ? '\n⚠️ ' + alertes.join('\n⚠️ ') : '\n✅ Aucune anomalie : tous les créateurs, remises et paiements sont reconnus.'));
  return { ecarts: ecarts, alertes: alertes, nb: sortie.length };
}

/*************************************************************
 *  MIGRATION « BACKOFFICE V1 » — étape 4 : la bascule
 *   apercuEtape4()  : reconstruit et contrôle, puis dit ce qui serait renommé
 *   etape4Bascule() : À LANCER BOUTIQUE FERMÉE, juste avant de déployer le
 *                     nouveau code. Bloque la caisse pendant l'opération,
 *                     reconstruit `ventes_v1` avec les toutes dernières ventes,
 *                     et n'effectue les renommages QUE si le contrôle est parfait.
 *************************************************************/

function apercuEtape4() { _etape4(true); }
function etape4Bascule() { _etape4(false); }

function _etape4(apercu) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(ANCIEN_VENTES_ARCHIVE)) throw new Error('La bascule a déjà été faite.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);                 // même verrou que la caisse : aucune vente ne peut s'écrire pendant la bascule
  try {
    const r = _etape3(apercu);
    if (r.ecarts || r.alertes.length) {
      throw new Error('Contrôle non conforme (' + r.ecarts + ' mois en écart, ' + r.alertes.length + ' alerte(s)) : bascule ANNULÉE, rien n\'a été renommé.');
    }
    const plan = [
      '« ' + ANCIEN_VENTES + ' » → « ' + ANCIEN_VENTES_ARCHIVE + ' » (gardé pour contrôle)',
      '« ' + SHEET_VENTES_V1 + ' » → « ' + SHEET_VENTES + ' » (' + r.nb + ' ventes)',
      '« ' + ANCIEN_PAIEMENTS + ' » → « ' + SHEET_PAIEMENTS + ' »',
      'en-têtes de « ' + SHEET_PAIEMENTS + ' » : code, taux_frais, statut, modifie_le',
      'en-têtes de « ' + SHEET_REMISES + ' » : code, valeur, type, statut, modifie_le'
    ];
    if (apercu) {
      Logger.log("APERÇU — contrôle parfait, rien n'a été renommé. La bascule ferait :\n• " + plan.join('\n• '));
      return;
    }
    ss.getSheetByName(ANCIEN_VENTES).setName(ANCIEN_VENTES_ARCHIVE);
    ss.getSheetByName(SHEET_VENTES_V1).setName(SHEET_VENTES);
    ss.getSheetByName(ANCIEN_PAIEMENTS).setName(SHEET_PAIEMENTS);
    _renommerEntetes(ss.getSheetByName(SHEET_PAIEMENTS), { type_de_paiement: 'code', valeur: 'taux_frais', status: 'statut', date_modification: 'modifie_le' });
    _renommerEntetes(ss.getSheetByName(SHEET_REMISES), { type_de_remise: 'code', status: 'statut', date_modification: 'modifie_le' });
    ss.getSheetByName(SHEET_JOURNAL).appendRow([new Date(), Session.getEffectiveUser().getEmail(), 'migration_etape4_bascule', plan.join(' | ')]);
    Logger.log('✅ Bascule faite :\n• ' + plan.join('\n• ') + '\n→ Déploie maintenant le nouveau code (serveur + site).');
  } finally {
    lock.releaseLock();
  }
}

function _renommerEntetes(sh, correspondances) {
  const r = sh.getRange(1, 1, 1, sh.getLastColumn());
  r.setValues([r.getValues()[0].map(function (h) { const k = _norm(h); return correspondances[k] || h; })]);
}

/*************************************************************
 *  MIGRATION « BACKOFFICE V1 » — étape 5 : le nettoyage
 *   apercuNettoyage() : liste les onglets qui seraient supprimés
 *   etape5Nettoyage() : les supprime. À lancer seulement après une semaine
 *                       de fonctionnement sans souci, et avec une copie de
 *                       sauvegarde du fichier (la suppression est définitive
 *                       dans ce fichier ; la copie, elle, garde tout).
 *************************************************************/

function apercuNettoyage() { _etape5(true); }
function etape5Nettoyage() { _etape5(false); }

function _etape5(apercu) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(ANCIEN_VENTES_ARCHIVE)) throw new Error("La bascule n'a pas encore été faite : rien à nettoyer.");
  const aSupprimer = ss.getSheets().map(function (s) { return s.getName(); }).filter(function (nom) {
    return [ANCIEN_ALL_DATA, ANCIEN_VENDEURS, ANCIEN_VENTES_ARCHIVE, SHEET_CORRESPONDANCE, SHEET_CONTROLE, 'Looker'].indexOf(nom) !== -1 ||
      /^archives?\b/i.test(nom);
  });
  if (apercu) {
    Logger.log('APERÇU — ' + aSupprimer.length + ' onglet(s) seraient supprimés :\n• ' + aSupprimer.join('\n• ') +
      '\nRestent : ' + ss.getSheets().map(function (s) { return s.getName(); }).filter(function (n) { return aSupprimer.indexOf(n) === -1; }).join(', '));
    return;
  }
  aSupprimer.forEach(function (nom) { ss.deleteSheet(ss.getSheetByName(nom)); });
  ss.getSheetByName(SHEET_JOURNAL).appendRow([new Date(), Session.getEffectiveUser().getEmail(), 'migration_etape5_nettoyage', aSupprimer.join(', ')]);
  Logger.log('✅ ' + aSupprimer.length + ' onglet(s) supprimés : ' + aSupprimer.join(', '));
}
