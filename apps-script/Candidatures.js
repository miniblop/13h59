/*************************************************************
 *  CANDIDATURES
 *   - page publique « candidater.html » : infos des forfaits + envoi
 *     (sans mot de passe, aucune donnée personnelle renvoyée)
 *   - Gestion ▸ Candidatures : liste, statut, retenir, remarque
 *  Statuts : a_traiter → liste_attente / non_retenu / retenu.
 *  Une même adresse e-mail = une seule candidature (mise à jour si elle recandidate).
 *************************************************************/

const STATUTS_CANDIDATURE = ['a_traiter', 'liste_attente', 'non_retenu', 'retenu'];
const CANDIDATURES_PAR_HEURE = 30;   // au-delà : on suspend (robots)

/** Texte saisi par le public : sans caractères de contrôle, tronqué, jamais lu comme une formule par le Sheet. */
function _textePublic(v, max) {
  const s = String(v == null ? '' : v).replace(/\r\n?/g, '\n').replace(/[\u0000-\u0009\u000b-\u001f\u007f]/g, '').trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function _email(v) {
  const s = String(v == null ? '' : v).trim().toLowerCase();
  return /^[^\s@'"<>]+@[^\s@'"<>]+\.[a-z]{2,}$/.test(s) && s.length <= 120 ? s : '';
}
/** Compte Instagram → lien complet ; sinon un lien web ; sinon ''. */
function _lienCandidat(v) {
  const s = String(v == null ? '' : v).trim();
  const c = _compteInstagram(s);
  if (c) return _lienInstagram(c);
  return /^https?:\/\/[^\s<>"]+\.[^\s<>"]+$/i.test(s) ? s.slice(0, 200) : '';
}

/* ---------- Page publique ---------- */

function _standsAvecPlaces(ss) {
  const tE = _tableau(ss, SHEET_EMPLACEMENTS), auj = _aujourdhui();
  return _lireTable(_onglet(ss, SHEET_STANDS)).filter(function (s) { return s['code']; }).map(function (s) {
    const code = String(s['code']), places = Number(s['places']) || 0;
    return { code: code, libelle: String(s['libelle'] || code), loyer: Number(s['loyer']) || 0, places: places, libres: Math.max(0, places - _occupation(tE, code, auj, '')) };
  });
}
function _categoriesPubliques(ss) {
  return _categoriesTriees(ss).filter(function (c) {
    return c['actif'] !== false && String(c['code']) !== 'shop';
  }).map(function (c) { return { code: String(c['code']), libelle: String(c['libelle'] || c['code']) }; });
}

function _candidatureInfos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stands = _standsAvecPlaces(ss).map(function (s) { return { code: s.code, libelle: s.libelle, loyer: s.loyer, complet: s.libres === 0 }; });
  return { ok: true, stands: stands, categories: _categoriesPubliques(ss) };
}

function _candidatureEnvoyer(body) {
  if (body.site) return { ok: true };                       // champ piège rempli : robot, on ne dit rien
  if (!(Number(body.duree) >= 4000)) return { ok: false, message: 'Envoi trop rapide : relis ta candidature puis renvoie-la.' };
  const cache = CacheService.getScriptCache();
  const nbHeure = Number(cache.get('candidatures_heure') || 0);
  if (nbHeure >= CANDIDATURES_PAR_HEURE) return { ok: false, message: 'Beaucoup de candidatures en ce moment : réessaie dans une heure, ou écris-nous à ' + EMAIL_SHOP + '.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const c = {
    prenom: _nomPropre(_textePublic(body.prenom, 60)), nom: _nomPropre(_textePublic(body.nom, 60)),
    marque: _nomPropre(_textePublic(body.marque, 80)), email: _email(body.email),
    instagram: _lienCandidat(body.instagram), description: _textePublic(body.description, 3000),
    categorie: String(body.categorie || ''), stand: String(body.stand || '')
  };
  const erreurs = [];
  if (!c.prenom) erreurs.push('ton prénom');
  if (!c.nom) erreurs.push('ton nom');
  if (!c.marque) erreurs.push("ton nom d'artiste ou de marque");
  if (!c.email) erreurs.push('une adresse e-mail valide');
  if (!c.instagram) erreurs.push('ton compte Instagram (ou le lien de ton site)');
  if (c.description.length < 30) erreurs.push('une description de tes créations (une ou deux phrases au moins)');
  if (!_categoriesPubliques(ss).some(function (x) { return x.code === c.categorie; })) erreurs.push('ta catégorie');
  if (!_standsAvecPlaces(ss).some(function (x) { return x.code === c.stand; })) erreurs.push('le forfait souhaité');
  if (body.accord !== true) erreurs.push("ton accord pour la conservation de ta candidature");
  if (erreurs.length) return { ok: false, message: 'Il manque ' + erreurs.join(', ') + '.' };

  const cleEmail = 'candidature_' + c.email;
  if (cache.get(cleEmail)) return { ok: true, deja: true };   // double clic ou renvoi immédiat

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, message: 'Le serveur est occupé : réessaie dans une minute.' };
  let deja = false, id = '';
  try {
    const tK = _tableau(ss, SHEET_CANDIDATURES), tC = _tableau(ss, SHEET_CREATEURS);
    const createur = tC.lignes.filter(function (r) { return _norm(_val(tC, r, 'email')) === c.email; })[0];
    const valeurs = {
      maj_le: new Date(), prenom: c.prenom, nom: c.nom, marque: c.marque, instagram: c.instagram,
      categorie: c.categorie, stand_souhaite: c.stand, description: c.description
    };
    const i = tK.lignes.findIndex(function (r) { return _norm(_val(tK, r, 'email')) === c.email; });
    if (i >= 0) {
      deja = true;
      const r = tK.lignes[i];
      id = String(_val(tK, r, 'id_candidature'));
      Object.keys(valeurs).forEach(function (k) { if (tK.M[k] != null) r[tK.M[k]] = valeurs[k]; });
      if (_val(tK, r, 'statut') === 'non_retenu') { r[tK.M['statut']] = 'a_traiter'; r[tK.M['traitee_le']] = ''; r[tK.M['traitee_par']] = ''; }
      if (createur && !_val(tK, r, 'id_createur')) r[tK.M['id_createur']] = _val(tC, createur, 'id_createur');
      _ecrireLigne(tK, i, r);
      _journaliser('candidature_maj', id + ' ' + c.marque + ' (nouvelle candidature de la même adresse)', 'site public');
    } else {
      id = _prochainId(tK, 'id_candidature', 'CA');
      valeurs.id_candidature = id; valeurs.recue_le = new Date(); valeurs.email = c.email;
      valeurs.statut = 'a_traiter'; valeurs.source = 'site';
      valeurs.id_createur = createur ? _val(tC, createur, 'id_createur') : '';
      _ajouterLigne(tK, valeurs);
      _journaliser('candidature_recue', id + ' ' + c.marque + ' · ' + c.stand, 'site public');
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  cache.put(cleEmail, '1', 120);
  cache.put('candidatures_heure', String(nbHeure + 1), 3600);
  try {
    envoyerModele(ss, 'accuse_reception', c.email, { prenom: c.prenom, marque: c.marque, stand: _libelleStand(ss, c.stand), date: '' });
  } catch (e) {
    _journaliser('email_echec', id + ' accusé de réception : ' + (e && e.message ? e.message : e), 'site public');
  }
  return { ok: true, deja: deja };
}

function _libelleStand(ss, code) {
  const s = _lireTable(_onglet(ss, SHEET_STANDS)).filter(function (x) { return String(x['code']) === String(code); })[0];
  return s ? String(s['libelle'] || code) : String(code || '');
}

/* ---------- Gestion ▸ Candidatures ---------- */

function _objetCandidature(tK, r) {
  const s = function (k) { return String(_val(tK, r, k) == null ? '' : _val(tK, r, k)); };
  return {
    id: s('id_candidature'), recueLe: _iso(_val(tK, r, 'recue_le')), majLe: _iso(_val(tK, r, 'maj_le')),
    prenom: s('prenom'), nom: s('nom'), marque: _nomPropre(s('marque')), email: s('email'), instagram: s('instagram'),
    categorie: s('categorie'), standSouhaite: s('stand_souhaite'), description: s('description'),
    statut: STATUTS_CANDIDATURE.indexOf(s('statut')) >= 0 ? s('statut') : 'a_traiter',
    traiteeLe: _iso(_val(tK, r, 'traitee_le')), traiteePar: s('traitee_par'), idCreateur: s('id_createur'),
    source: s('source'), remarque: s('remarque')
  };
}

function _gCandidatures() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tK = _tableau(ss, SHEET_CANDIDATURES), tC = _tableau(ss, SHEET_CREATEURS);

  const createurs = {}, parEmail = {}, parMarque = {};
  tC.lignes.forEach(function (r) {
    const id = String(_val(tC, r, 'id_createur'));
    if (!id) return;
    const c = { id: id, nom: _nomPropre(_val(tC, r, 'nom')), statut: _norm(_val(tC, r, 'statut')) };
    createurs[id] = c;
    if (_val(tC, r, 'email')) parEmail[_norm(_val(tC, r, 'email'))] = c;
    if (_cleLarge(c.nom)) parMarque[_cleLarge(c.nom)] = c;
  });

  const candidatures = tK.lignes.filter(function (r) { return _val(tK, r, 'id_candidature'); }).map(function (r) { return _objetCandidature(tK, r); });

  // Même personne ? même e-mail, même compte Instagram ou même marque (accents et ponctuation ignorés).
  const index = {};
  const cles = function (k) {
    const ig = _compteInstagram(k.instagram);
    return ['e:' + _norm(k.email), ig ? 'i:' + ig : '', _cleLarge(k.marque) ? 'm:' + _cleLarge(k.marque) : ''].filter(function (x) { return x && x !== 'e:'; });
  };
  candidatures.forEach(function (k) { cles(k).forEach(function (x) { (index[x] = index[x] || []).push(k.id); }); });
  candidatures.forEach(function (k) {
    const autres = {};
    cles(k).forEach(function (x) { index[x].forEach(function (id) { if (id !== k.id) autres[id] = true; }); });
    k.doublons = Object.keys(autres);
    const c = createurs[k.idCreateur] || parEmail[_norm(k.email)] || parMarque[_cleLarge(k.marque)];
    k.createur = c ? { id: c.id, nom: c.nom, statut: c.statut, lie: c.id === k.idCreateur } : null;
  });

  const modeles = _modelesEmails(ss);
  return {
    ok: true, aujourdhui: _iso(_aujourdhui()), candidatures: candidatures,
    stands: _standsAvecPlaces(ss),
    categories: _categoriesTriees(ss)
      .map(function (c) { return { code: String(c['code']), libelle: String(c['libelle'] || c['code']) }; }),
    modeles: modeles, modelesDefaut: _emailsParDefaut()
  };
}

function _ligneCandidature(tK, id) {
  const i = tK.lignes.findIndex(function (r) { return String(_val(tK, r, 'id_candidature')) === String(id); });
  if (i < 0) throw new Error('Candidature introuvable : « ' + id + ' ».');
  return i;
}
function _prenomPourEmail(tK, r) { return _nomPropre(_val(tK, r, 'prenom')) || _nomPropre(_val(tK, r, 'nom')) || _nomPropre(_val(tK, r, 'marque')); }
function _signataire() { return _auteurJournal.replace(/ \(site\)$/, ''); }

/** Envoie un modèle si demandé ; renvoie {email: true/false, emailErreur?} sans jamais annuler l'action déjà faite. */
function _emailApresAction(ss, envoyer, code, to, vars, id) {
  if (!envoyer) return { email: false };
  if (!to) return { email: false, emailErreur: "pas d'adresse e-mail sur cette candidature" };
  try {
    const envoye = envoyerModele(ss, code, to, vars);
    if (envoye) _journaliser('email_envoye', id + ' · modèle ' + code + ' → ' + to);
    return envoye ? { email: true } : { email: false, emailErreur: 'le modèle « ' + code + ' » est désactivé dans l\'onglet emails' };
  } catch (e) {
    const m = String(e && e.message ? e.message : e);
    _journaliser('email_echec', id + ' · modèle ' + code + ' : ' + m);
    return { email: false, emailErreur: m };
  }
}

function _gCandidatureStatut(body) {
  if (['a_traiter', 'liste_attente', 'non_retenu'].indexOf(body.statut) === -1) throw new Error('Statut invalide.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tK = _tableau(ss, SHEET_CANDIDATURES);
  const i = _ligneCandidature(tK, body.id), r = tK.lignes[i];
  const avant = String(_val(tK, r, 'statut'));
  if (avant === 'retenu') throw new Error('Candidature déjà retenue : le créateur se gère dans Gestion ▸ Créateurs.');
  r[tK.M['statut']] = body.statut;
  r[tK.M['traitee_le']] = body.statut === 'a_traiter' ? '' : new Date();
  r[tK.M['traitee_par']] = body.statut === 'a_traiter' ? '' : _signataire();
  _ecrireLigne(tK, i, r);
  const marque = _nomPropre(_val(tK, r, 'marque'));
  const mail = body.statut === 'a_traiter' ? { email: false } : _emailApresAction(ss, body.email === true, body.statut, String(_val(tK, r, 'email')),
    { prenom: _prenomPourEmail(tK, r), marque: marque, stand: _libelleStand(ss, _val(tK, r, 'stand_souhaite')), date: '' }, body.id);
  _journaliser('candidature_statut', body.id + ' ' + marque + ' : ' + avant + ' → ' + body.statut + (mail.email ? ' (e-mail envoyé)' : ''));
  return { ok: true, email: mail.email, emailErreur: mail.emailErreur || '' };
}

/** Retenir : crée le créateur (ou réintègre un ancien) avec son emplacement, puis prévient par e-mail si demandé. */
function _gCandidatureRetenir(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tK = _tableau(ss, SHEET_CANDIDATURES);
  const i = _ligneCandidature(tK, body.id), r = tK.lignes[i];
  if (String(_val(tK, r, 'statut')) === 'retenu') throw new Error('Cette candidature est déjà retenue.');
  if (!body.stand) throw new Error('Choisis le stand.');
  _dateIso(body.debut);
  const marque = _nomPropre(_val(tK, r, 'marque'));
  let idCreateur = String(body.idCreateur || '');
  if (idCreateur) {
    _gEmplacementCreer({ idCreateur: idCreateur, stand: body.stand, debut: body.debut, accueilPar: body.accueilPar });
  } else {
    const ig = _compteInstagram(_val(tK, r, 'instagram'));
    idCreateur = _gCreateurCreer({
      nom: _textePublic(marque, 80), email: _val(tK, r, 'email'), categorie: body.categorie || _val(tK, r, 'categorie'),
      stand: body.stand, debut: body.debut, accueilPar: body.accueilPar,
      instagram: ig ? _lienInstagram(ig) : '', nomLegal: _textePublic([_val(tK, r, 'prenom'), _val(tK, r, 'nom')].filter(String).join(' '), 120)
    }).id;
  }
  const tK2 = _tableau(ss, SHEET_CANDIDATURES), r2 = tK2.lignes[i];
  r2[tK2.M['statut']] = 'retenu';
  r2[tK2.M['traitee_le']] = new Date();
  r2[tK2.M['traitee_par']] = _signataire();
  r2[tK2.M['id_createur']] = idCreateur;
  _ecrireLigne(tK2, i, r2);
  const d = _dateIso(body.debut);
  const mail = _emailApresAction(ss, body.email === true, 'retenu', String(_val(tK2, r2, 'email')),
    { prenom: _prenomPourEmail(tK2, r2), marque: marque, stand: _libelleStand(ss, body.stand), date: Utilities.formatDate(d, _tz(), 'dd/MM/yyyy') }, body.id);
  _journaliser('candidature_retenue', body.id + ' ' + marque + ' → ' + idCreateur + ' · ' + body.stand + ' dès le ' + body.debut + (mail.email ? ' (e-mail envoyé)' : ''));
  return { ok: true, idCreateur: idCreateur, email: mail.email, emailErreur: mail.emailErreur || '' };
}

function _gCandidatureRemarque(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tK = _tableau(ss, SHEET_CANDIDATURES);
  const i = _ligneCandidature(tK, body.id), r = tK.lignes[i];
  r[tK.M['remarque']] = _textePublic(body.remarque, 1000);
  _ecrireLigne(tK, i, r);
  _journaliser('candidature_remarque', body.id + ' ' + _nomPropre(_val(tK, r, 'marque')));
  return { ok: true };
}
