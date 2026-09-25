/*************************************************************
 *  GESTION — créateurs, stands et emplacements (onglet Gestion du site)
 *  Réservé au mot de passe gestion. Chaque écriture est notée dans `journal`.
 *  Les dates circulent au format 'yyyy-MM-dd'.
 *************************************************************/

const CHAMPS_CREATEUR_MODIFIABLES = [
  'nom', 'email', 'statut', 'categorie', 'nom_legal', 'telephone', 'adresse', 'code_postal', 'ville',
  'siret', 'iban', 'benevole', 'perms_prevues', 'rc_pro', 'adhesion_payee_le', 'instagram'
];

/** Point d'entrée des actions « gestion_* ». */
function _gestion(body) {
  const role = _role(body.password);
  if (role !== 'gestion') return { ok: false, message: role ? "Ce mot de passe ne donne pas accès à la gestion." : 'Mot de passe incorrect.' };
  const actions = {
    gestion_createurs: _gCreateurs,
    gestion_createur_maj: _gCreateurMaj,
    gestion_createur_creer: _gCreateurCreer,
    gestion_preavis: _gPreavis,
    gestion_preavis_annuler: _gPreavisAnnuler,
    gestion_changer_stand: _gChangerStand,
    gestion_emplacement_creer: _gEmplacementCreer,
    gestion_candidatures: _gCandidatures,
    gestion_candidature_statut: _gCandidatureStatut,
    gestion_candidature_retenir: _gCandidatureRetenir,
    gestion_candidature_remarque: _gCandidatureRemarque,
    gestion_email_maj: _gEmailMaj,
    gestion_reglages: _gReglages,
    gestion_reglage_maj: _gReglageMaj,
    gestion_reglage_creer: _gReglageCreer,
    gestion_ventes: _gVentes,
    gestion_vente_annuler: _gVenteAnnuler,
    gestion_vente_corriger: _gVenteCorriger,
    gestion_panier_annuler: _gPanierAnnuler,
    gestion_vente_supprimer: _gVenteSupprimer,
    gestion_panier_supprimer: _gPanierSupprimer,
    gestion_journal: _gJournal
  };
  const f = actions[body.action];
  if (!f) return { ok: false, message: 'Action inconnue.' };
  _auteurJournal = _auteur(body.qui);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const r = f(body);
    SpreadsheetApp.flush();
    return r;
  } catch (e) {
    return { ok: false, message: String(e && e.message ? e.message : e) };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Outils ---------- */

function _tz() { return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(); }
function _iso(d) { return d instanceof Date ? Utilities.formatDate(d, _tz(), 'yyyy-MM-dd') : ''; }
function _dateIso(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
  if (!m) throw new Error('Date invalide : « ' + s + ' ».');
  return new Date(+m[1], +m[2] - 1, +m[3]);
}
function _aujourdhui() { const d = new Date(); return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
/** Même jour le mois suivant (31/01 → 28/02). */
function _plusUnMois(d) {
  const fin = new Date(d.getFullYear(), d.getMonth() + 2, 0).getDate();
  return new Date(d.getFullYear(), d.getMonth() + 1, Math.min(d.getDate(), fin));
}
function _veille(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate() - 1); }

/** Un onglet sous forme {sh, M: {entête: index}, lignes: [[…]]}. */
function _tableau(ss, nom) {
  const sh = _onglet(ss, nom);
  const h = _headerMap(sh);
  const n = sh.getLastRow() - 1;
  return { sh: sh, M: h.map, nbCol: h.nbCol, lignes: n > 0 ? sh.getRange(2, 1, n, h.nbCol).getValues() : [] };
}
function _val(t, r, champ) { return t.M[champ] != null ? r[t.M[champ]] : ''; }
function _ecrireLigne(t, index, ligne) { t.sh.getRange(index + 2, 1, 1, ligne.length).setValues([ligne]); }
function _ajouterLigne(t, valeurs) {
  const ligne = new Array(t.nbCol).fill('');
  Object.keys(valeurs).forEach(function (k) { if (t.M[k] != null) ligne[t.M[k]] = valeurs[k]; });
  t.sh.appendRow(ligne);
  t.lignes.push(ligne);
}
function _prochainId(t, champ, prefixe) {
  const max = t.lignes.reduce(function (m, r) { const x = new RegExp('^' + prefixe + '(\\d+)$').exec(String(_val(t, r, champ))); return x ? Math.max(m, +x[1]) : m; }, 0);
  return prefixe + String(max + 1).padStart(3, '0');
}
/** Auteur des lignes du journal : prénom saisi sur le site (nettoyé pour ne jamais être lu comme une formule). */
let _auteurJournal = 'gestion (site)';
function _auteur(qui) {
  const p = String(qui || '').replace(/[\u0000-\u001f<>"]/g, '').replace(/^[=+\-@\s]+/, '').trim().slice(0, 30);
  return p ? p + ' (site)' : 'gestion (site)';
}
function _journaliser(action, detail, qui) {
  _onglet(SpreadsheetApp.getActiveSpreadsheet(), SHEET_JOURNAL).appendRow([new Date(), qui || _auteurJournal, action, detail]);
}

/** Emplacement actif à une date donnée (début passé, fin absente ou à venir). */
function _actifLe(t, r, d) {
  const debut = _val(t, r, 'debut'), fin = _val(t, r, 'fin');
  return debut instanceof Date && debut <= d && (!(fin instanceof Date) || fin >= d);
}
function _occupation(tE, stand, d, sauf) {
  return tE.lignes.filter(function (r) {
    return String(_val(tE, r, 'code_stand')) === stand && String(_val(tE, r, 'id_createur')) !== sauf && _actifLe(tE, r, d);
  }).length;
}
function _placesStand(ss, stand) {
  const s = _lireTable(_onglet(ss, SHEET_STANDS)).filter(function (x) { return String(x['code']) === stand; })[0];
  if (!s) throw new Error('Stand inconnu : « ' + stand + ' ».');
  return Number(s['places']) || 0;
}

/* ---------- Lecture ---------- */

/** Aligne le statut des créateurs sur leurs emplacements (arrivée atteinte → actif, départ passé → inactif).
 *  L'équipe (catégorie « shop ») et les créateurs sans emplacement sont gérés à la main. */
function _synchroniserStatuts(ss) {
  const tC = _tableau(ss, SHEET_CREATEURS), tE = _tableau(ss, SHEET_EMPLACEMENTS);
  const auj = _aujourdhui(), empl = {}, changements = [];
  tE.lignes.forEach(function (r) { const id = String(_val(tE, r, 'id_createur')); (empl[id] = empl[id] || []).push(r); });
  tC.lignes.forEach(function (r, i) {
    const id = String(_val(tC, r, 'id_createur'));
    if (!empl[id] || String(_val(tC, r, 'categorie')) === 'shop') return;
    const voulu = empl[id].some(function (e) { return _actifLe(tE, e, auj); }) ? 'actif' : 'inactif';
    if (_norm(_val(tC, r, 'statut')) === voulu) return;
    r[tC.M['statut']] = voulu;
    if (tC.M['modifie_le'] != null) r[tC.M['modifie_le']] = new Date();
    _ecrireLigne(tC, i, r);
    changements.push(_val(tC, r, 'nom') + ' → ' + voulu);
  });
  if (changements.length) _journaliser('statut_automatique', changements.join(', '), 'automatique');
}

function _gCreateurs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  _synchroniserStatuts(ss);
  const tC = _tableau(ss, SHEET_CREATEURS), tE = _tableau(ss, SHEET_EMPLACEMENTS);
  const auj = _aujourdhui();

  const ventes = {};
  _lireTable(_onglet(ss, SHEET_VENTES)).forEach(function (v) {
    const id = String(v['id_createur'] || '');
    const o = ventes[id] || (ventes[id] = { ca: 0, n: 0, derniere: null });
    o.ca += Number(v['prix_client']) || 0; o.n++;
    if (v['date'] instanceof Date && (!o.derniere || v['date'] > o.derniere)) o.derniere = v['date'];
  });

  const createurs = tC.lignes.filter(function (r) { return String(_val(tC, r, 'id_createur')); }).map(function (r) {
    const id = String(_val(tC, r, 'id_createur')), v = ventes[id] || { ca: 0, n: 0, derniere: null };
    const iban = String(_val(tC, r, 'iban') || '').replace(/\s/g, '');
    return {
      id: id, nom: _nomPropre(_val(tC, r, 'nom')), email: String(_val(tC, r, 'email') || ''),
      statut: _norm(_val(tC, r, 'statut')), categorie: String(_val(tC, r, 'categorie') || ''),
      nomLegal: String(_val(tC, r, 'nom_legal') || ''), telephone: String(_val(tC, r, 'telephone') || ''),
      adresse: String(_val(tC, r, 'adresse') || ''), codePostal: String(_val(tC, r, 'code_postal') || ''),
      ville: String(_val(tC, r, 'ville') || ''), siret: String(_val(tC, r, 'siret') || ''),
      ibanFin: iban ? iban.slice(-4) : '', benevole: _val(tC, r, 'benevole') === true,
      permsPrevues: _val(tC, r, 'perms_prevues') === '' ? '' : Number(_val(tC, r, 'perms_prevues')),
      rcPro: _val(tC, r, 'rc_pro') === true, adhesion: _iso(_val(tC, r, 'adhesion_payee_le')),
      instagram: String(_val(tC, r, 'instagram') || ''), creeLe: _iso(_val(tC, r, 'cree_le')),
      modifieLe: _iso(_val(tC, r, 'modifie_le')),
      ca: _round2(v.ca), nbVentes: v.n, derniereVente: _iso(v.derniere)
    };
  });
  const emplacements = tE.lignes.filter(function (r) { return String(_val(tE, r, 'id_emplacement')); }).map(function (r) {
    return {
      id: String(_val(tE, r, 'id_emplacement')), idCreateur: String(_val(tE, r, 'id_createur')), stand: String(_val(tE, r, 'code_stand')),
      debut: _iso(_val(tE, r, 'debut')), fin: _iso(_val(tE, r, 'fin')), preavisRecuLe: _iso(_val(tE, r, 'preavis_recu_le')),
      accueilPar: String(_val(tE, r, 'accueil_par') || ''), motifFin: String(_val(tE, r, 'motif_fin') || ''), remarque: String(_val(tE, r, 'remarque') || '')
    };
  });
  const stands = _lireTable(_onglet(ss, SHEET_STANDS)).filter(function (s) { return s['code']; }).map(function (s) {
    return { code: String(s['code']), libelle: String(s['libelle'] || s['code']), loyer: Number(s['loyer']) || 0, places: Number(s['places']) || 0 };
  });
  const categories = _categoriesTriees(ss).map(function (c) {
    return { code: String(c['code']), libelle: String(c['libelle'] || c['code']), actif: c['actif'] !== false };
  });
  return { ok: true, aujourdhui: _iso(auj), createurs: createurs, emplacements: emplacements, stands: stands, categories: categories };
}

/* ---------- Fiche créateur ---------- */

/** Valeur prête à écrire pour un champ de `createurs` ; lève une erreur lisible si elle est invalide. */
function _valeurChamp(champ, v, ctx) {
  const s = String(v == null ? '' : v).trim();
  switch (champ) {
    case 'nom': {
      const nom = _nomPropre(s);
      if (!nom) throw new Error('Le nom de marque est obligatoire.');
      if (ctx.nomsPris[nom.toLowerCase()] && ctx.nomsPris[nom.toLowerCase()] !== ctx.id) throw new Error('Un autre créateur s\'appelle déjà « ' + nom + ' ».');
      return nom;
    }
    case 'email':
      if (s && !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(s)) throw new Error('E-mail invalide.');
      return s.toLowerCase();
    case 'statut':
      if (['actif', 'inactif'].indexOf(s) === -1) throw new Error('Statut invalide.');
      return s;
    case 'categorie':
      if (s && !ctx.categories[s]) throw new Error('Catégorie inconnue : « ' + s + ' ».');
      return s;
    case 'telephone': return s ? _telephone(s) : '';
    case 'code_postal':
      if (s && !/^\d{5}$/.test(s)) throw new Error('Code postal : 5 chiffres.');
      return s;
    case 'siret': {
      const d = s.replace(/\D/g, '');
      if (d && !_siretValide(d)) throw new Error('SIRET invalide (14 chiffres, vérifie la saisie).');
      return d;
    }
    case 'iban': {
      const i = s.replace(/\s/g, '').toUpperCase();
      if (i && !_ibanValide(i)) throw new Error('IBAN invalide.');
      return i;
    }
    case 'benevole': case 'rc_pro': return v === true || v === 'true';
    case 'perms_prevues':
      if (s === '') return '';
      if (!(Number(s) >= 0)) throw new Error('Nombre de permanences invalide.');
      return Number(s);
    case 'adhesion_payee_le': return s ? _dateIso(s) : '';
    case 'instagram': {
      if (!s) return '';
      const c = _compteInstagram(s);
      if (!c) throw new Error('Compte Instagram invalide : indique le nom du compte ou son lien.');
      return _lienInstagram(c);
    }
    default: return s;
  }
}
function _ibanValide(i) {
  if (!/^[A-Z]{2}\d{2}[A-Z0-9]{10,30}$/.test(i)) return false;
  const r = (i.slice(4) + i.slice(0, 4)).replace(/[A-Z]/g, function (c) { return c.charCodeAt(0) - 55; });
  let m = 0; for (let k = 0; k < r.length; k++) m = (m * 10 + +r[k]) % 97;
  return m === 1;
}
function _contexteCreateurs(ss, tC, id) {
  const nomsPris = {}, categories = {};
  tC.lignes.forEach(function (r) { nomsPris[_nomPropre(_val(tC, r, 'nom')).toLowerCase()] = String(_val(tC, r, 'id_createur')); });
  _lireTable(_onglet(ss, SHEET_CATEGORIES)).forEach(function (c) { if (c['code']) categories[String(c['code'])] = true; });
  return { id: id, nomsPris: nomsPris, categories: categories };
}

function _gCreateurMaj(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tC = _tableau(ss, SHEET_CREATEURS);
  const i = tC.lignes.findIndex(function (r) { return String(_val(tC, r, 'id_createur')) === String(body.id); });
  if (i < 0) throw new Error('Créateur introuvable : « ' + body.id + ' ».');
  const r = tC.lignes[i], ctx = _contexteCreateurs(ss, tC, String(body.id));
  const champs = body.champs || {}, modifies = [];
  Object.keys(champs).forEach(function (k) {
    if (CHAMPS_CREATEUR_MODIFIABLES.indexOf(k) === -1 || tC.M[k] == null) throw new Error('Champ non modifiable : « ' + k + ' ».');
    const v = _valeurChamp(k, champs[k], ctx), avant = r[tC.M[k]];
    const egal = (avant instanceof Date && v instanceof Date) ? avant.getTime() === v.getTime() : String(avant) === String(v);
    if (egal) return;
    r[tC.M[k]] = v; modifies.push(k);
  });
  if (!modifies.length) return { ok: true, modifies: [] };
  if (tC.M['modifie_le'] != null) r[tC.M['modifie_le']] = new Date();
  _ecrireLigne(tC, i, r);
  _journaliser('createur_maj', body.id + ' ' + _nomPropre(_val(tC, r, 'nom')) + ' : ' + modifies.join(', '));
  return { ok: true, modifies: modifies };
}

function _gCreateurCreer(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tC = _tableau(ss, SHEET_CREATEURS), tE = _tableau(ss, SHEET_EMPLACEMENTS);
  const ctx = _contexteCreateurs(ss, tC, '');
  const nom = _valeurChamp('nom', body.nom, ctx);
  const email = _valeurChamp('email', body.email, ctx);
  const categorie = _valeurChamp('categorie', body.categorie, ctx);
  const instagram = body.instagram ? _valeurChamp('instagram', body.instagram, ctx) : '';
  const id = _prochainId(tC, 'id_createur', 'C');
  let statut = 'inactif', detail = '';
  if (body.stand) {
    const debut = _dateIso(body.debut);
    if (_occupation(tE, body.stand, debut, '') >= _placesStand(ss, body.stand)) throw new Error('Plus de place libre en « ' + body.stand + ' » le ' + body.debut + '.');
    _ajouterLigne(tE, { id_emplacement: _prochainId(tE, 'id_emplacement', 'E'), id_createur: id, code_stand: body.stand, debut: debut, accueil_par: String(body.accueilPar || '') });
    if (debut <= _aujourdhui()) statut = 'actif';
    detail = ' · stand ' + body.stand + ' dès le ' + body.debut;
  }
  _ajouterLigne(tC, { id_createur: id, nom: nom, email: email, statut: statut, categorie: categorie, instagram: instagram, nom_legal: String(body.nomLegal || ''),
                      benevole: false, rc_pro: false, cree_le: new Date(), modifie_le: new Date() });
  _journaliser('createur_creer', id + ' ' + nom + detail);
  return { ok: true, id: id };
}

/* ---------- Emplacements ---------- */

function _ligneEmplacement(tE, idEmplacement) {
  const i = tE.lignes.findIndex(function (r) { return String(_val(tE, r, 'id_emplacement')) === String(idEmplacement); });
  if (i < 0) throw new Error('Emplacement introuvable : « ' + idEmplacement + ' ».');
  return i;
}

/** Préavis reçu par e-mail le `recuLe` : départ un mois plus tard. */
function _gPreavis(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tE = _tableau(ss, SHEET_EMPLACEMENTS);
  const i = _ligneEmplacement(tE, body.idEmplacement), r = tE.lignes[i];
  if (_val(tE, r, 'fin') instanceof Date) throw new Error('Cet emplacement a déjà une date de fin.');
  const recu = _dateIso(body.recuLe), fin = _plusUnMois(recu);
  r[tE.M['preavis_recu_le']] = recu;
  r[tE.M['fin']] = fin;
  r[tE.M['motif_fin']] = body.pause ? 'pause' : 'depart';
  _ecrireLigne(tE, i, r);
  _journaliser('preavis', body.idEmplacement + ' (' + _val(tE, r, 'id_createur') + ') : reçu le ' + body.recuLe + ', départ le ' + _iso(fin) + (body.pause ? ' (pause)' : ''));
  return { ok: true, fin: _iso(fin) };
}

function _gPreavisAnnuler(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tE = _tableau(ss, SHEET_EMPLACEMENTS);
  const i = _ligneEmplacement(tE, body.idEmplacement), r = tE.lignes[i];
  const fin = _val(tE, r, 'fin');
  if (!(fin instanceof Date)) throw new Error("Cet emplacement n'a pas de préavis.");
  if (fin < _aujourdhui()) throw new Error('Le départ a déjà eu lieu : crée plutôt un nouvel emplacement.');
  ['fin', 'preavis_recu_le', 'motif_fin'].forEach(function (k) { r[tE.M[k]] = ''; });
  _ecrireLigne(tE, i, r);
  _journaliser('preavis_annule', body.idEmplacement + ' (' + _val(tE, r, 'id_createur') + ')');
  return { ok: true };
}

/** Change de stand à partir de `date` : l'emplacement actuel se termine la veille. */
function _gChangerStand(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tE = _tableau(ss, SHEET_EMPLACEMENTS);
  const date = _dateIso(body.date), id = String(body.idCreateur);
  const i = tE.lignes.findIndex(function (r) { return String(_val(tE, r, 'id_createur')) === id && _actifLe(tE, r, date); });
  if (i < 0) throw new Error("Ce créateur n'a pas d'emplacement actif à cette date.");
  const r = tE.lignes[i], ancien = String(_val(tE, r, 'code_stand'));
  if (ancien === body.stand) throw new Error('Le créateur est déjà sur ce stand.');
  if (_occupation(tE, body.stand, date, id) >= _placesStand(ss, body.stand)) throw new Error('Plus de place libre en « ' + body.stand + ' » le ' + body.date + '.');
  if (_val(tE, r, 'debut').getTime() === date.getTime()) {
    // L'emplacement commence ce jour-là : on corrige simplement son stand.
    r[tE.M['code_stand']] = body.stand;
    _ecrireLigne(tE, i, r);
    _journaliser('changement_stand', id + ' : ' + ancien + ' → ' + body.stand + ' le ' + body.date + ' (correction)');
    return { ok: true };
  }
  r[tE.M['fin']] = _veille(date);
  r[tE.M['motif_fin']] = 'changement_stand';
  _ecrireLigne(tE, i, r);
  _ajouterLigne(tE, { id_emplacement: _prochainId(tE, 'id_emplacement', 'E'), id_createur: id, code_stand: body.stand, debut: date });
  _journaliser('changement_stand', id + ' : ' + _val(tE, r, 'code_stand') + ' → ' + body.stand + ' le ' + body.date);
  return { ok: true };
}

/** Nouvel emplacement pour un créateur qui n'en a pas (retour, ou créateur ajouté sans stand). */
function _gEmplacementCreer(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tC = _tableau(ss, SHEET_CREATEURS), tE = _tableau(ss, SHEET_EMPLACEMENTS);
  const id = String(body.idCreateur);
  if (!tC.lignes.some(function (r) { return String(_val(tC, r, 'id_createur')) === id; })) throw new Error('Créateur introuvable : « ' + id + ' ».');
  const debut = _dateIso(body.debut);
  const enCours = tE.lignes.filter(function (r) {
    const fin = _val(tE, r, 'fin');
    return String(_val(tE, r, 'id_createur')) === id && (!(fin instanceof Date) || fin >= debut);
  });
  if (enCours.length) throw new Error('Ce créateur a déjà un emplacement à cette date : un créateur = un seul stand.');
  if (_occupation(tE, body.stand, debut, '') >= _placesStand(ss, body.stand)) throw new Error('Plus de place libre en « ' + body.stand + ' » le ' + body.debut + '.');
  _ajouterLigne(tE, { id_emplacement: _prochainId(tE, 'id_emplacement', 'E'), id_createur: id, code_stand: body.stand, debut: debut, accueil_par: String(body.accueilPar || '') });
  _journaliser('emplacement_creer', id + ' : ' + body.stand + ' dès le ' + body.debut);
  return { ok: true };
}
