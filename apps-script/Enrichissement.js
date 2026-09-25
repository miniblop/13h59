/*************************************************************
 *  BACKOFFICE — étape 6 : enrichissement des créateurs
 *   apercuEnrichissement() : liste ce qui serait fait, n'écrit RIEN
 *   etape6Enrichissement() : complète `createurs`, crée `stands` et
 *                            `emplacements`, et écrit `_revue_enrichissement`
 *  Sources (fichiers de l'équipe, lus sans être modifiés) :
 *   - « Base de données » : créateurs présents, stand, identité, SIRET
 *   - « LISTE CREATEURS », onglet 2026 : catégorie, Instagram, adhésion
 *     (et onglet « Feuille 9 » : stands, en secours)
 *************************************************************/

const ID_FICHIER_BASE   = '1h538tdGm8Y-DzjdI0PND7XKkcYDArJdmYZWin20omDg';
const ID_FICHIER_LISTE  = '1Jg79ugNWQJqYCqYoR9mektHCgv8sDPMEtVVnufD9B2Q';
const SHEET_STANDS       = 'stands';
const SHEET_EMPLACEMENTS = 'emplacements';
const SHEET_REVUE        = '_revue_enrichissement';

const STANDS_INITIAUX = [
  // code, libelle, loyer, places, ref_facture, perms_loyer_gratuit
  ['illu',     'Exposant illustration · espace partagé',  28, 10, '',    ''],
  ['unique',   'Exposant créateur · petit espace unique', 48, 18, '',    ''],
  ['grand',    'Exposant créateur · grand stand',         68, 6,  'LGS', 4],
  ['friperie', 'Exposant friperie',                       70, 3,  '',    '']
];
const STAND_PAR_LOYER = { 28: 'illu', 48: 'unique', 68: 'grand', 70: 'friperie' };
const STAND_PAR_COLONNE_F9 = { 'S.UNIQUE': 'unique', 'STAND ILLU': 'illu', 'FRIPE': 'friperie', 'STAND 68': 'grand' };
const CATEGORIE_PAR_GROUPE = { 'BIJOUX': 'bijoux', 'ACCESSOIRES': 'accessoires', 'DECO': 'deco', 'ILLUSTRATION': 'illustration', 'FRIPE': 'friperie', '1.SHOP ET RESIDENT': 'shop' };
const COLONNES_EMPLACEMENTS = ['id_emplacement', 'id_createur', 'code_stand', 'debut', 'fin', 'preavis_recu_le', 'accueil_par', 'motif_fin', 'remarque'];

/** Valeur d'une colonne, quel que soit l'accent, l'apostrophe ou la casse de son en-tête. */
function _champ(o, nom) {
  const k = _cleLarge(nom);
  const cle = Object.keys(o).filter(function (x) { return _cleLarge(x) === k; })[0];
  return cle ? o[cle] : '';
}

function apercuEnrichissement() { _etape6(true); }
function etape6Enrichissement() { _etape6(false); }

function _telephone(v) {
  let d = String(v == null ? '' : v).replace(/\D/g, '');
  if (d.length === 9) d = '0' + d;
  if (d.length === 11 && d.indexOf('33') === 0) d = '0' + d.slice(2);
  return d.length === 10 ? d.replace(/(\d{2})(?=\d)/g, '$1 ') : String(v || '').trim();
}
function _siretValide(s) {
  if (!/^\d{14}$/.test(s)) return false;
  let t = 0;
  for (let i = 0; i < 14; i++) { let n = +s[13 - i]; if (i % 2) { n *= 2; if (n > 9) n -= 9; } t += n; }
  return t % 10 === 0 || s.indexOf('356000000') === 0;   // La Poste : exception connue
}
/** « 12 rue X, 59000 Lille » → {cp:'59000', ville:'Lille'} (dernier code postal trouvé). */
function _cpVille(adresse) {
  const a = String(adresse || '').replace(/\s+/g, ' ').trim();
  const m = a.match(/(\d{2}) ?(\d{3})(?!.*\d{2} ?\d{3})\s*,?\s*([^,\d]*)/);
  return m ? { cp: m[1] + m[2], ville: m[3].replace(/\b(france|app\w*.*)$/i, '').trim() } : { cp: '', ville: '' };
}

/** « sept. 26 » saisi dans Sheets devient le 26/09 de l'année en cours : le jour est en fait
 *  l'année (25 → 2025, 26 → 2026). On le ramène au 1er du mois de cette année-là. */
function _moisAdhesion(d) {
  if (!(d instanceof Date)) return null;
  const j = d.getDate();
  return (j >= 20 && j <= 30) ? new Date(2000 + j, d.getMonth(), 1) : d;
}

function _etape6(apercu) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const revue = [];                       // [id, nom, champ, avant, apres, remarque]
  const bilan = [];
  const noter = function (id, nom, champ, avant, apres, rem) { revue.push([id, nom, champ, avant === undefined ? '' : avant, apres === undefined ? '' : apres, rem || '']); };

  // --- Créateurs du backoffice ---
  const shC = _onglet(ss, SHEET_CREATEURS);
  let h = _headerMap(shC);
  if (h.map['instagram'] == null) {
    if (!apercu) { shC.getRange(1, h.nbCol + 1).setValue('instagram').setFontWeight('bold'); h = _headerMap(shC); }
    bilan.push('createurs : ajout de la colonne « instagram ».');
  }
  const n = shC.getLastRow() - 1;
  const lignes = shC.getRange(2, 1, n, Math.max(h.nbCol, (h.map['instagram'] == null ? h.nbCol + 1 : h.nbCol))).getValues();
  const col = function (k) { return h.map[k] != null ? h.map[k] : h.nbCol; };   // « instagram » en aperçu
  const createurs = lignes.map(function (r, i) {
    return { i: i, id: String(r[col('id_createur')]), nom: _nomPropre(r[col('nom')]), email: _norm(r[col('email')]), r: r };
  });
  const parEmail = {}, parLarge = {};
  createurs.forEach(function (c) { if (c.email) parEmail[c.email] = c; parLarge[_cleLarge(c.nom)] = c; });
  const trouver = function (nom, email) {
    const e = _norm(email);
    if (e && parEmail[e]) return { c: parEmail[e], par: 'e-mail' };
    const k = _cleLarge(nom);
    if (!k) return null;
    if (parLarge[k]) return { c: parLarge[k], par: 'nom' };
    const proches = createurs.filter(function (c) {
      const kc = _cleLarge(c.nom);
      if (!kc) return false;
      if (Math.min(kc.length, k.length) >= 3 && (kc.indexOf(k) === 0 || k.indexOf(kc) === 0)) return true;
      return Math.min(kc.length, k.length) >= 6 && _levenshtein(kc, k) <= 2;
    });
    return proches.length === 1 ? { c: proches[0], par: 'nom proche (' + proches[0].nom + ')' } : null;
  };
  const modifier = function (c, champ, valeur, rem) {
    if (valeur === '' || valeur == null) return;
    const idx = col(champ), avant = c.r[idx];
    const egal = (avant instanceof Date && valeur instanceof Date) ? avant.getTime() === valeur.getTime() : String(avant) === String(valeur);
    if (egal) return;
    c.r[idx] = valeur; c.modifie = true;
    noter(c.id, c.nom, champ, avant instanceof Date ? Utilities.formatDate(avant, tz, 'dd/MM/yyyy') : avant,
      valeur instanceof Date ? Utilities.formatDate(valeur, tz, 'dd/MM/yyyy') : valeur, rem);
  };

  // --- Source 1 : Base de données (présents, stand, identité) ---
  const base = _lireTable(SpreadsheetApp.openById(ID_FICHIER_BASE).getSheets()[0]);
  const presents = {}, standDe = {};
  let nbBase = 0;
  base.forEach(function (b) {
    const nom = _champ(b, "nom d'artiste"), email = _champ(b, 'mail');
    if (!_nomPropre(nom) && !_norm(email)) return;
    nbBase++;
    const t = trouver(nom, email);
    if (!t) { noter('', _nomPropre(nom), '—', '', '', 'Base de données : créateur introuvable dans « createurs » (à ajouter ou à rattacher à la main)'); return; }
    const c = t.c;
    presents[c.id] = true;
    if (t.par !== 'nom') noter(c.id, c.nom, 'rattachement', _nomPropre(nom), c.nom, 'Base de données, rattaché par ' + t.par);
    const loyer = Number(String(_champ(b, 'stand') || '').replace(',', '.'));
    if (STAND_PAR_LOYER[loyer]) standDe[c.id] = STAND_PAR_LOYER[loyer];
    const cv = _cpVille(_champ(b, 'adresse'));
    const siret = String(_champ(b, 'n°siret') || '').replace(/\D/g, '');
    modifier(c, 'nom_legal', _nomPropre(_champ(b, 'nom et prénom')));
    modifier(c, 'adresse', String(_champ(b, 'adresse') || '').replace(/\s+/g, ' ').trim());
    modifier(c, 'code_postal', cv.cp);
    modifier(c, 'ville', cv.ville);
    modifier(c, 'telephone', _telephone(_champ(b, 'numéro de tel')));
    modifier(c, 'email', _norm(email));
    if (siret) modifier(c, 'siret', siret, _siretValide(siret) ? '' : '⚠️ SIRET invalide (vérifier les chiffres)');
  });
  bilan.push('Base de données : ' + nbBase + ' créateurs, ' + Object.keys(presents).length + ' rattachés.');

  // --- Source 2 : LISTE CREATEURS / 2026 (catégorie, Instagram, adhésion) ---
  const liste = SpreadsheetApp.openById(ID_FICHIER_LISTE);
  const suivi = _lireTable(liste.getSheetByName('2026'));
  const equipe = {};
  let nbSuivi = 0;
  suivi.forEach(function (s) {
    const marque = _champ(s, 'marque');
    if (!_nomPropre(marque)) return;
    nbSuivi++;
    const nomCourt = String(marque).split(' - ')[0];
    const t = trouver(marque, _champ(s, 'contact')) || trouver(nomCourt, '');
    if (!t) { noter('', _nomPropre(marque), '—', '', '', 'LISTE CREATEURS / 2026 : introuvable dans « createurs »'); return; }
    const c = t.c;
    const groupe = String(_champ(s, 'colonne 1') || '').trim().toUpperCase();
    if (groupe === '1.SHOP ET RESIDENT') equipe[c.id] = true;
    modifier(c, 'categorie', CATEGORIE_PAR_GROUPE[groupe] || '');
    modifier(c, 'instagram', String(_champ(s, 'compte ig') || '').replace(/^@/, '').trim());
    const adh = _moisAdhesion(_champ(s, 'date adhésion'));
    if (adh) modifier(c, 'adhesion_payee_le', adh, "mois d'adhésion lu dans LISTE CREATEURS (« " +
      Utilities.formatDate(_champ(s, 'date adhésion'), tz, 'dd/MM/yyyy') + ' » compris comme ' + Utilities.formatDate(adh, tz, 'MM/yyyy') + ')');
  });
  bilan.push('LISTE CREATEURS / 2026 : ' + nbSuivi + ' marques lues.');

  // Stand de secours : « Feuille 9 » (répartition des stands) pour ceux sans stand dans la Base.
  const f9 = liste.getSheetByName('Feuille 9');
  if (f9) {
    const v = f9.getDataRange().getValues();
    v[0].forEach(function (titre, j) {
      const code = STAND_PAR_COLONNE_F9[String(titre).replace(/\s*\(.*\)$/, '').trim().toUpperCase()];
      if (!code) return;
      v.slice(1).forEach(function (r) {
        const t = r[j] ? trouver(r[j], '') : null;
        if (t && presents[t.c.id] && !standDe[t.c.id]) { standDe[t.c.id] = code; noter(t.c.id, t.c.nom, 'stand', '', code, 'stand pris dans « Feuille 9 »'); }
      });
    });
  }

  // --- Statut : présents (Base) + équipe = actif ; les autres passent inactif sauf vente récente ---
  const ventes = _lireTable(_onglet(ss, SHEET_VENTES));
  const derniereVente = {}, premiereVente = {};
  ventes.forEach(function (v) {
    const id = String(v['id_createur']), d = v['date'];
    if (!(d instanceof Date)) return;
    if (!derniereVente[id] || d > derniereVente[id]) derniereVente[id] = d;
    if (!premiereVente[id] || d < premiereVente[id]) premiereVente[id] = d;
  });
  const il30j = new Date(Date.now() - 30 * 864e5);
  let nbActifs = 0;
  createurs.forEach(function (c) {
    const avant = _norm(c.r[col('statut')]);
    let statut = (presents[c.id] || equipe[c.id]) ? 'actif' : 'inactif', rem = '';
    if (statut === 'inactif' && avant === 'actif' && derniereVente[c.id] && derniereVente[c.id] > il30j) {
      statut = 'actif';
      rem = '⚠️ absent de la Base mais a vendu le ' + Utilities.formatDate(derniereVente[c.id], tz, 'dd/MM') + ' : laissé actif, à vérifier';
      noter(c.id, c.nom, 'statut', avant, statut, rem);
    }
    if (statut !== avant) modifier(c, 'statut', statut, statut === 'inactif' ? 'absent de la Base de données' : '');
    if (statut === 'actif') nbActifs++;
  });
  createurs.forEach(function (c) { if (c.modifie) modifier(c, 'modifie_le', new Date()); });
  bilan.push('createurs : ' + createurs.filter(function (c) { return c.modifie; }).length + ' fiche(s) modifiée(s), ' + nbActifs + ' créateurs actifs.');

  // --- Emplacements (un par créateur présent avec un stand connu) ---
  const emplacements = Object.keys(presents).filter(function (id) { return standDe[id]; }).sort().map(function (id, i) {
    // L'adhésion est renouvelée chaque année : la première vente est un meilleur repère d'arrivée.
    const debut = premiereVente[id] || '';
    return ['E' + String(i + 1).padStart(3, '0'), id, standDe[id], debut, '', '', '', '',
      debut ? 'début = première vente (à vérifier)' : "début inconnu (aucune vente) : à compléter"];
  });
  Object.keys(presents).filter(function (id) { return !standDe[id]; }).forEach(function (id) {
    const c = createurs.filter(function (x) { return x.id === id; })[0];
    noter(id, c.nom, 'stand', '', '', '⚠️ présent dans la Base mais sans stand connu : emplacement à créer à la main');
  });
  const occupation = {};
  emplacements.forEach(function (e) { occupation[e[2]] = (occupation[e[2]] || 0) + 1; });
  bilan.push('emplacements : ' + emplacements.length + ' (' + STANDS_INITIAUX.map(function (s) { return s[0] + ' ' + (occupation[s[0]] || 0) + '/' + s[3]; }).join(', ') + ').');

  const aVerifier = revue.filter(function (r) { return String(r[5]).indexOf('⚠️') === 0 || !r[0]; }).length;
  bilan.push(SHEET_REVUE + ' : ' + revue.length + ' ligne(s), dont ' + aVerifier + ' à vérifier.');

  if (!apercu) {
    shC.getRange(2, 1, n, lignes[0].length).setValues(lignes.map(function (r) { return r.slice(0, lignes[0].length); }));
    const cats = _onglet(ss, SHEET_CATEGORIES);
    if (!_lireTable(cats).some(function (c) { return c['code'] === 'shop'; })) {
      cats.appendRow(['shop', 'Shop & résidents', "l'équipe et les résidents du shop", cats.getLastRow(), true]);
    }
    if (!ss.getSheetByName(SHEET_STANDS)) _nouvelOnglet(ss, SHEET_STANDS, ['code', 'libelle', 'loyer', 'places', 'ref_facture', 'perms_loyer_gratuit'], STANDS_INITIAUX);
    if (!ss.getSheetByName(SHEET_EMPLACEMENTS)) {
      const sE = _nouvelOnglet(ss, SHEET_EMPLACEMENTS, COLONNES_EMPLACEMENTS, emplacements);
      if (emplacements.length) sE.getRange(2, 4, emplacements.length, 3).setNumberFormat('dd/mm/yyyy');
    } else bilan.push(SHEET_EMPLACEMENTS + ' existe déjà : laissé tel quel.');
    const r = ss.getSheetByName(SHEET_REVUE); if (r) ss.deleteSheet(r);
    const sR = _nouvelOnglet(ss, SHEET_REVUE, ['id_createur', 'nom', 'champ', 'avant', 'apres', 'remarque'], revue);
    revue.forEach(function (l, i) { if (String(l[5]).indexOf('⚠️') === 0 || !l[0]) sR.getRange(i + 2, 1, 1, 6).setBackground('#fff3cd'); });
    _onglet(ss, SHEET_JOURNAL).appendRow([new Date(), Session.getEffectiveUser().getEmail(), 'etape6_enrichissement', bilan.join(' | ')]);
  }
  Logger.log((apercu ? "APERÇU — rien n'a été écrit.\n" : '✅ Enrichissement terminé.\n') + bilan.map(function (l) { return '• ' + l; }).join('\n') +
    '\nÀ vérifier :\n' + revue.filter(function (r) { return String(r[5]).indexOf('⚠️') === 0 || !r[0]; }).map(function (r) { return '  - ' + (r[1] || '?') + ' : ' + r[5]; }).join('\n'));
}
