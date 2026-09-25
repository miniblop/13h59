/*************************************************************
 *  RÉGLAGES — Gestion ▸ Réglages
 *  Catégories, stands, remises et moyens de paiement se gèrent depuis le
 *  site : le Sheet n'est que le stockage (future base de données).
 *  Règles : un code ne change jamais (clé étrangère d'autres tables) et
 *  rien n'est supprimé, on désactive. Chaque modification va dans `journal`.
 *  Les montants circulent en unités lisibles : 10 (%) ou 5 (€), 1,75 (%).
 *************************************************************/

const REMISE_PAR_DEFAUT = 'pas_de_remise';

/** Code technique à partir d'un libellé : « Déco & maison » → « deco_maison ». */
function _codeDepuis(libelle) {
  return _cleNom(libelle).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '').slice(0, 30);
}
function _nombre(v, nom, min, max) {
  const s = String(v == null ? '' : v).replace(',', '.').trim();
  const n = Number(s);
  if (s === '' || isNaN(n)) throw new Error(nom + ' : indique un nombre.');
  if (n < min || n > max) throw new Error(nom + ' : entre ' + min + ' et ' + max + '.');
  return n;
}
function _libelle(v, nom, max) {
  const s = _nomPropre(_textePublic(v, max));
  if (!s) throw new Error(nom + ' est obligatoire.');
  return s;
}

/** Nombre d'utilisations d'un code dans une colonne (pour afficher « utilisé par … »). */
function _compter(ss, onglet, colonne) {
  const out = {};
  _lireTable(_onglet(ss, onglet)).forEach(function (r) { const c = String(r[colonne] == null ? '' : r[colonne]).trim(); if (c) out[c] = (out[c] || 0) + 1; });
  return out;
}

function _gReglages() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parCategorie = _compter(ss, SHEET_CREATEURS, 'categorie');
  const ventes = _lireTable(_onglet(ss, SHEET_VENTES)), parRemise = {}, parPaiement = {};
  ventes.forEach(function (v) {
    const r = String(v['code_remise'] || '').trim(), p = String(v['code_paiement'] || '').trim();
    if (r) parRemise[r] = (parRemise[r] || 0) + 1;
    if (p) parPaiement[p] = (parPaiement[p] || 0) + 1;
  });
  const tE = _tableau(ss, SHEET_EMPLACEMENTS), auj = _aujourdhui(), occupes = function (code) { return _occupation(tE, code, auj, ''); };
  const pct = function (x) { return _round2(Math.abs(Number(x) || 0) * 100); };

  return {
    ok: true,
    categories: _lireTable(_onglet(ss, SHEET_CATEGORIES)).filter(function (c) { return c['code']; }).map(function (c) {
      return { code: String(c['code']), libelle: String(c['libelle'] || ''), exemples: String(c['exemples'] || ''), ordre: c['ordre'] === '' ? '' : Number(c['ordre']),
               actif: c['actif'] !== false, utilise: parCategorie[String(c['code'])] || 0 };
    }).sort(function (a, b) { return (Number(a.ordre) || 999) - (Number(b.ordre) || 999); }),
    stands: _lireTable(_onglet(ss, SHEET_STANDS)).filter(function (s) { return s['code']; }).map(function (s) {
      return { code: String(s['code']), libelle: String(s['libelle'] || ''), loyer: Number(s['loyer']) || 0, places: Number(s['places']) || 0,
               ref_facture: String(s['ref_facture'] || ''), perms_loyer_gratuit: s['perms_loyer_gratuit'] === '' ? '' : Number(s['perms_loyer_gratuit']),
               utilise: occupes(String(s['code'])) };
    }),
    remises: _lireTable(_onglet(ss, SHEET_REMISES)).filter(function (r) { return r['code']; }).map(function (r) {
      const pourcentage = _norm(r['type']) === 'pourcentage';
      return { code: String(r['code']), type: pourcentage ? 'pourcentage' : 'montant', valeur: pourcentage ? pct(r['valeur']) : _round2(Math.abs(Number(r['valeur']) || 0)),
               statut: _norm(r['statut']) === 'actif' ? 'actif' : 'inactif', verrouille: String(r['code']) === REMISE_PAR_DEFAUT, utilise: parRemise[String(r['code'])] || 0 };
    }),
    paiements: _lireTable(_onglet(ss, SHEET_PAIEMENTS)).filter(function (p) { return p['code']; }).map(function (p) {
      return { code: String(p['code']), taux_frais: pct(p['taux_frais']), statut: _norm(p['statut']) === 'actif' ? 'actif' : 'inactif', utilise: parPaiement[String(p['code'])] || 0 };
    }),
    modeles: _modelesEmails(ss), modelesDefaut: _emailsParDefaut()
  };
}

/* ---------- Validation, table par table ---------- */

/** Valeurs à écrire (colonnes du Sheet) pour les champs reçus du site. Lève une erreur lisible. */
function _valeursReglage(table, champs, creation) {
  const c = champs || {}, v = {}, a = function (k) { return Object.prototype.hasOwnProperty.call(c, k) || creation; };
  switch (table) {
    case 'categories':
      if (a('libelle')) v.libelle = _libelle(c.libelle, 'Le nom', 60);
      if (a('exemples')) v.exemples = _textePublic(c.exemples, 200);
      if (a('ordre')) v.ordre = String(c.ordre == null ? '' : c.ordre).trim() === '' ? '' : _nombre(c.ordre, 'Ordre', 0, 999);
      if (a('actif')) v.actif = c.actif !== false && c.actif !== 'false';
      break;
    case 'stands':
      if (a('libelle')) v.libelle = _libelle(c.libelle, 'Le nom', 80);
      if (a('loyer')) v.loyer = _nombre(c.loyer, 'Loyer', 0, 10000);
      if (a('places')) { v.places = _nombre(c.places, 'Places', 0, 500); if (v.places !== Math.round(v.places)) throw new Error('Places : un nombre entier.'); }
      if (a('ref_facture')) v.ref_facture = _textePublic(String(c.ref_facture || '').toUpperCase(), 10);
      if (a('perms_loyer_gratuit')) v.perms_loyer_gratuit = String(c.perms_loyer_gratuit == null ? '' : c.perms_loyer_gratuit).trim() === '' ? '' : _nombre(c.perms_loyer_gratuit, 'Permanences', 0, 31);
      break;
    case 'remises': {
      const type = a('type') ? String(c.type) : null;
      if (type != null && ['pourcentage', 'montant'].indexOf(type) === -1) throw new Error('Type de remise : pourcentage ou montant.');
      if (a('valeur')) {
        if (type == null) throw new Error('Indique le type de la remise avec sa valeur.');
        const n = type === 'pourcentage' ? _nombre(c.valeur, 'Pourcentage', 0, 100) : _nombre(c.valeur, 'Montant', 0, 1000);
        v.valeur = type === 'pourcentage' ? -_round2(n) / 100 : -_round2(n);
        v.type = type;
      }
      if (a('statut')) v.statut = c.statut === 'actif' ? 'actif' : 'inactif';
      v.modifie_le = new Date();
      break;
    }
    case 'paiements':
      if (a('taux_frais')) v.taux_frais = _round2(_nombre(c.taux_frais, 'Frais', 0, 20) * 100) / 10000;
      if (a('statut')) v.statut = c.statut === 'actif' ? 'actif' : 'inactif';
      v.modifie_le = new Date();
      break;
    default:
      throw new Error('Réglage inconnu : « ' + table + ' ».');
  }
  return v;
}
function _ongletReglage(table) {
  return { categories: SHEET_CATEGORIES, stands: SHEET_STANDS, remises: SHEET_REMISES, paiements: SHEET_PAIEMENTS }[table];
}
function _resumeValeurs(v) {
  return Object.keys(v).filter(function (k) { return k !== 'modifie_le'; }).map(function (k) { return k + '=' + (v[k] instanceof Date ? _iso(v[k]) : v[k]); }).join(', ');
}

function _gReglageMaj(body) {
  const table = String(body.table || ''), code = String(body.code || '');
  const onglet = _ongletReglage(table);
  if (!onglet) throw new Error('Réglage inconnu : « ' + table + ' ».');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const t = _tableau(ss, onglet);
  const i = t.lignes.findIndex(function (r) { return String(_val(t, r, 'code')) === code; });
  if (i < 0) throw new Error('Code introuvable : « ' + code + ' ».');
  if (table === 'remises' && code === REMISE_PAR_DEFAUT) throw new Error('« ' + REMISE_PAR_DEFAUT + ' » est la remise par défaut de la caisse : elle ne se modifie pas.');
  const v = _valeursReglage(table, body.champs, false);
  if (table === 'paiements' && v.statut === 'inactif') {
    const autres = t.lignes.filter(function (r, j) { return j !== i && _norm(_val(t, r, 'statut')) === 'actif'; }).length;
    if (!autres) throw new Error('Il faut garder au moins un moyen de paiement actif pour la caisse.');
  }
  const r = t.lignes[i];
  Object.keys(v).forEach(function (k) { if (t.M[k] != null) r[t.M[k]] = v[k]; });
  _ecrireLigne(t, i, r);
  _journaliser('reglage_' + table, code + ' : ' + _resumeValeurs(v));
  return { ok: true };
}

function _gReglageCreer(body) {
  const table = String(body.table || '');
  const onglet = _ongletReglage(table);
  if (!onglet) throw new Error('Réglage inconnu : « ' + table + ' ».');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const t = _tableau(ss, onglet);
  const c = body.champs || {};
  // Code : dérivé du nom pour catégories et stands ; saisi pour remises et paiements (c'est ce que la caisse affiche).
  let code;
  if (table === 'categories' || table === 'stands') code = _codeDepuis(c.libelle);
  else if (table === 'remises') code = _nomPropre(_textePublic(c.code, 40)).replace(/\s+/g, '_');
  else code = _nomPropre(_textePublic(c.code, 30)).toLowerCase().replace(/\s+/g, '_');
  if (!code || /^'/.test(code)) throw new Error(table === 'remises' || table === 'paiements' ? 'Indique le nom affiché en caisse.' : 'Le nom est obligatoire.');
  if (t.lignes.some(function (r) { return _norm(_val(t, r, 'code')) === code.toLowerCase(); })) throw new Error('« ' + code + ' » existe déjà : modifie-le ou réactive-le plutôt.');
  const v = _valeursReglage(table, table === 'remises' || table === 'paiements' ? Object.assign({ statut: 'actif' }, c) : c, true);
  v.code = code;
  if (table === 'categories' && v.ordre === '') v.ordre = t.lignes.length + 1;
  _ajouterLigne(t, v);
  if (table === 'categories' && t.M['actif'] != null) t.sh.getRange(t.lignes.length + 1, t.M['actif'] + 1).insertCheckboxes().setValue(v.actif);
  _journaliser('reglage_' + table + '_creation', code + ' : ' + _resumeValeurs(v));
  return { ok: true, code: code };
}
