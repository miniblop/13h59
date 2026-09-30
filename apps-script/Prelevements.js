/*************************************************************
 *  PRÉLÈVEMENTS SEPA — le loyer du mois M prélevé le 25 du mois M−1
 *  Semi-automatique : le site prépare tout, une personne du bureau dépose
 *  le fichier sur l'espace pro du CIC (les banques n'ouvrent pas d'API de
 *  prélèvement aux associations).
 *  1. Mandats (onglet `mandats`) : lien personnel envoyé par e-mail ; le
 *     créateur signe sur mandat.html en saisissant son IBAN (jamais renvoyé
 *     au navigateur). Un mandat signé sur papier peut aussi être saisi.
 *  2. Pré-notification : montant et date par e-mail, au moins 14 jours avant
 *     (délai SEPA par défaut) ; crée la ligne dans `prelevements`.
 *  3. Fichier SEPA (pain.008, CORE) à déposer au CIC : les lignes passent
 *     « dans un fichier ».
 *  4. Une fois l'argent arrivé, « Prélèvements encaissés » pointe les loyers
 *     (paiements_recus, source prelevement) ; un rejet dépointe et le loyer
 *     redevient dû (relance habituelle).
 *  L'ICS (identifiant créancier SEPA) se saisit dans Facturation ▸ Prélèvements
 *  et reste dans les propriétés du script (ASSO_ICS).
 *************************************************************/

const SHEET_MANDATS = 'mandats';
const COLONNES_MANDATS = ['rum', 'id_createur', 'statut', 'jeton', 'demande_le', 'demande_par', 'signe_le', 'mode', 'titulaire', 'adresse', 'code_postal', 'ville',
  'iban', 'bic', 'email_signature', 'premier_prelevement_le', 'revoque_le', 'remarque'];
const SHEET_PRELEVEMENTS = 'prelevements';
const COLONNES_PRELEVEMENTS = ['id_prelevement', 'mois_loyer', 'id_createur', 'rum', 'montant', 'date_prelevement', 'sequence', 'statut', 'prenotifie_le', 'lot',
  'fichier_le', 'id_paiement', 'par', 'remarque'];
const MANDAT_LIEN_JOURS = 30;          // validité du lien de signature
const PRENOTIFICATION_JOURS = 14;      // délai SEPA par défaut, sauf accord plus court
/** Créancier tel qu'il figure sur le mandat : siège social de l'association (convention) — à vérifier avec l'adresse déclarée pour l'ICS. */
const ASSO_CREANCIER = { nom: 'COLLECTIF 13H59', adresse: '17 rue Jean Jaurès', ville: '59160 Lomme', pays: 'France' };
/** Pays de la zone SEPA (préfixe d'IBAN). */
const PAYS_SEPA = 'AD AT BE BG CH CY CZ DE DK EE ES FI FR GB GI GR HR HU IE IS IT LI LT LU LV MC MT NL NO PL PT RO SE SI SK SM VA'.split(' ');

function _ics() { return String(PropertiesService.getScriptProperties().getProperty('ASSO_ICS') || ''); }
function _icsValide(s) { return /^FR\d{2}[A-Z0-9]{3}[A-Z0-9]{6}$/.test(String(s || '')); }

function _ongletsPrelevement(ss) {
  if (!ss.getSheetByName(SHEET_MANDATS)) {
    _creerOngletSiAbsent(ss, SHEET_MANDATS, COLONNES_MANDATS);
    const m = ss.getSheetByName(SHEET_MANDATS);
    m.getRange(1, 1, m.getMaxRows(), 4).setNumberFormat('@');     // rum, id, statut, jeton
    m.getRange(1, 11, m.getMaxRows(), 4).setNumberFormat('@');    // code postal, ville, iban, bic
  }
  if (!ss.getSheetByName(SHEET_PRELEVEMENTS)) {
    _creerOngletSiAbsent(ss, SHEET_PRELEVEMENTS, COLONNES_PRELEVEMENTS);
    const p = ss.getSheetByName(SHEET_PRELEVEMENTS);
    p.getRange(1, 1, p.getMaxRows(), 4).setNumberFormat('@');     // id, mois, créateur, rum
    p.getRange(1, 10, p.getMaxRows(), 1).setNumberFormat('@');    // lot
  }
}

/* ---------- Jours ouvrés SEPA (TARGET2 : week-ends, 1er janvier, Vendredi saint, lundi de Pâques, 1er mai, 25 et 26 décembre) ---------- */

function _paques(a) {
  const b = a % 19, c = Math.floor(a / 100), d = a % 100, e = Math.floor(c / 4), f = c % 4, g = Math.floor((c + 8) / 25), h = Math.floor((c - g + 1) / 3);
  const i = (19 * b + c - e - h + 15) % 30, k = Math.floor(d / 4), l = d % 4, m = (32 + 2 * f + 2 * k - i - l) % 7, n = Math.floor((b + 11 * i + 22 * m) / 451);
  const mois = Math.floor((i + m - 7 * n + 114) / 31), jour = ((i + m - 7 * n + 114) % 31) + 1;
  return new Date(a, mois - 1, jour);
}
function _jourOuvreSepa(d) {
  if (d.getDay() === 0 || d.getDay() === 6) return false;
  const md = (d.getMonth() + 1) * 100 + d.getDate();
  if ([101, 501, 1225, 1226].indexOf(md) !== -1) return false;
  const p = _paques(d.getFullYear()), t = _jour(d).getTime();
  return t !== new Date(p.getFullYear(), p.getMonth(), p.getDate() - 2).getTime() && t !== new Date(p.getFullYear(), p.getMonth(), p.getDate() + 1).getTime();
}
/** Date proposée pour le loyer du mois M : le 25 du mois M−1, ou le jour ouvré d'avant. */
function _datePrelevementDefaut(mois) {
  const d = _echeance(mois);
  while (!_jourOuvreSepa(d)) d.setDate(d.getDate() - 1);
  return d;
}

/* ---------- Lecture ---------- */

function _lienExpire(demandeLe) { return !(demandeLe instanceof Date) || _aujourdhui() > new Date(demandeLe.getTime() + MANDAT_LIEN_JOURS * 864e5); }

/** Mandat retenu par créateur : signé d'abord, sinon la dernière demande. */
function _mandatsParCreateur(ss) {
  const out = {};
  if (!ss.getSheetByName(SHEET_MANDATS)) return out;
  const t = _tableau(ss, SHEET_MANDATS), rang = function (s) { return s === 'signe' ? 3 : s === 'a_signer' ? 2 : 1; };
  t.lignes.forEach(function (r, i) {
    const id = String(_val(t, r, 'id_createur'));
    if (!id) return;
    let st = String(_val(t, r, 'statut'));
    if (st === 'a_signer' && _lienExpire(_val(t, r, 'demande_le'))) st = 'expire';
    const o = { i: i, rum: String(_val(t, r, 'rum')), statut: st, demandeLe: _iso(_val(t, r, 'demande_le')), signeLe: _iso(_val(t, r, 'signe_le')), mode: String(_val(t, r, 'mode') || ''),
      ibanFin: String(_val(t, r, 'iban') || '').slice(-4), premier: _iso(_val(t, r, 'premier_prelevement_le')), revoqueLe: _iso(_val(t, r, 'revoque_le')) };
    const p = out[id];
    if (!p || rang(st) > rang(p.statut) || (rang(st) === rang(p.statut) && i > p.i)) out[id] = o;
  });
  return out;
}

function _gPrelevements(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), mois = String(body.mois || '');
  if (!/^\d{4}-\d{2}$/.test(mois)) throw new Error('Mois invalide.');
  const mandats = _mandatsParCreateur(ss), tC = _tableau(ss, SHEET_CREATEURS);
  const createurs = tC.lignes.filter(function (r) {
    return _val(tC, r, 'id_createur') && String(_val(tC, r, 'categorie')) !== 'shop' && _norm(_val(tC, r, 'statut')) === 'actif';
  }).map(function (r) {
    const id = String(_val(tC, r, 'id_createur')), m = mandats[id];
    return { id: id, nom: _nomPropre(_val(tC, r, 'nom')), email: String(_val(tC, r, 'email') || ''), mandat: m ? { rum: m.rum, statut: m.statut, demandeLe: m.demandeLe, signeLe: m.signeLe, mode: m.mode, ibanFin: m.ibanFin, revoqueLe: m.revoqueLe } : null };
  }).sort(function (a, b) { return a.nom.localeCompare(b.nom, 'fr', { sensitivity: 'base' }); });

  const enc = _gEncaissements({ mois: mois }), resteDe = {};
  enc.lignes.forEach(function (l) { resteDe[l.idCreateur] = l; });
  const prl = {};
  _lireSi(ss, SHEET_PRELEVEMENTS).forEach(function (p) {
    if (_moisTexte(p['mois_loyer']) !== mois || String(p['statut']) === 'annule') return;
    prl[String(p['id_createur'])] = { id: String(p['id_prelevement']), montant: Number(p['montant']) || 0, date: _iso(p['date_prelevement']), statut: String(p['statut']),
      prenotifieLe: _iso(p['prenotifie_le']), lot: String(p['lot'] || ''), rum: String(p['rum'] || ''), sequence: String(p['sequence'] || ''), remarque: String(p['remarque'] || '') };
  });
  const noms = {};
  createurs.forEach(function (c) { noms[c.id] = c; });
  const ids = Object.keys(mandats).filter(function (id) { return mandats[id].statut === 'signe' && noms[id]; }).concat(Object.keys(prl).filter(function (id) { return !(mandats[id] && mandats[id].statut === 'signe' && noms[id]); }));
  const lignes = ids.map(function (id) {
    const p = prl[id] || null, e = resteDe[id], m = mandats[id];
    return { idCreateur: id, nom: noms[id] ? noms[id].nom : (e ? e.nom : id), rum: p ? p.rum : m.rum, ibanFin: m ? m.ibanFin : '', loyer: e ? e.loyer : 0, reste: e ? e.reste : 0,
      montant: p ? p.montant : (e ? e.reste : 0), statut: p ? p.statut : (e && e.reste > 0 ? 'a_prenotifier' : 'rien'), prelevement: p };
  }).sort(function (a, b) { return a.nom.localeCompare(b.nom, 'fr', { sensitivity: 'base' }); });
  const lots = {};
  lignes.forEach(function (l) {
    const p = l.prelevement;
    if (!p || !p.lot || p.statut !== 'fichier') return;
    const x = lots[p.lot] = lots[p.lot] || { lot: p.lot, nb: 0, total: 0, date: p.date };
    x.nb++; x.total = _round2(x.total + p.montant);
  });
  let banqueOk = true;
  try { _banqueAsso(); } catch (e) { banqueOk = false; }
  return { ok: true, mois: mois, libelle: _moisBornes(mois).libelle, aujourdhui: _iso(_aujourdhui()), ics: _ics(), banqueOk: banqueOk, creancier: ASSO_CREANCIER,
    dateDefaut: _iso(_datePrelevementDefaut(mois)), prenotificationJours: PRENOTIFICATION_JOURS, createurs: createurs, lignes: lignes,
    lots: Object.keys(lots).map(function (k) { return lots[k]; }) };
}

/* ---------- Réglage ---------- */

function _gPrelevementIcs(body) {
  const ics = String(body.ics || '').replace(/\s/g, '').toUpperCase();
  if (!_icsValide(ics)) throw new Error('ICS invalide : 13 caractères, par exemple FR12ZZZ123456.');
  const avant = _ics();
  PropertiesService.getScriptProperties().setProperty('ASSO_ICS', ics);
  _journaliser('prelevement_ics', 'ICS ' + (avant ? avant + ' → ' : '') + ics);
  return { ok: true, ics: ics };
}

/* ---------- Mandats ---------- */

function _nouvelleRum(t, id) {
  const n = t.lignes.filter(function (r) { return String(_val(t, r, 'id_createur')) === id; }).length + 1;
  return '13H59-' + id + '-' + _deux(n);
}
function _fichesParId(ss) {
  const out = {};
  _lireTable(_onglet(ss, SHEET_CREATEURS)).forEach(function (c) {
    if (c['id_createur']) out[String(c['id_createur'])] = { nom: _nomPropre(c['nom']), email: _email(c['email']), nomLegal: String(c['nom_legal'] || ''),
      adresse: String(c['adresse'] || ''), codePostal: String(c['code_postal'] || ''), ville: String(c['ville'] || '') };
  });
  return out;
}

/** Envoie le lien de signature (un ou plusieurs créateurs). Un lien encore en attente est remplacé. */
function _gMandatsDemander(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!_icsValide(_ics())) throw new Error("Saisis d'abord l'ICS (identifiant créancier SEPA) : il figure sur le mandat.");
  _ongletsPrelevement(ss);
  _modeleSiAbsent(ss, 'mandat_demande');
  const perso = body.emailPerso && (body.emailPerso.objet || body.emailPerso.texte) ? body.emailPerso : null;
  if (perso) _textePerso(perso, {});
  const fiches = _fichesParId(ss), t = _tableau(ss, SHEET_MANDATS), envoyes = [], erreurs = [];
  (body.ids || []).forEach(function (id) {
    id = String(id);
    const f = fiches[id];
    if (!f) { erreurs.push(id + ' : fiche introuvable'); return; }
    if (!f.email) { erreurs.push(f.nom + " : pas d'e-mail valide sur la fiche"); return; }
    try {
      const jeton = Utilities.getUuid().replace(/-/g, '');
      const i = t.lignes.findIndex(function (r) { return String(_val(t, r, 'id_createur')) === id && String(_val(t, r, 'statut')) === 'a_signer'; });
      if (i >= 0) {
        const r = t.lignes[i];
        r[t.M['jeton']] = _empreinte(jeton); r[t.M['demande_le']] = new Date(); r[t.M['demande_par']] = _signataire();
        _ecrireLigne(t, i, r);
      } else {
        _ajouterLigne(t, { rum: _nouvelleRum(t, id), id_createur: id, statut: 'a_signer', jeton: _empreinte(jeton), demande_le: new Date(), demande_par: _signataire() });
      }
      const vars = { marque: f.nom, prenom: f.nom }, boutons = [{ libelle: 'Signer mon mandat', url: URL_SITE + 'mandat.html?j=' + jeton, principal: true }];
      const ok = perso ? envoyerTexte(ss, 'mandat_demande', f.email, perso, vars, boutons) : envoyerModele(ss, 'mandat_demande', f.email, vars, boutons);
      if (!ok) throw new Error('modèle « mandat_demande » désactivé');
      envoyes.push(f.nom);
    } catch (e) { erreurs.push(f.nom + ' : ' + (e && e.message ? e.message : e)); }
  });
  if (envoyes.length) _journaliser('mandats_demandes', envoyes.join(', ') + (perso && perso.modifie ? ' (texte modifié)' : ''));
  return { ok: true, envoyes: envoyes.length, erreurs: erreurs };
}

/** Fin d'un mandat (demande du créateur, changement de compte, départ) : plus aucun prélèvement avec cette RUM. */
function _gMandatRevoquer(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), id = String(body.idCreateur || ''), motif = _motif(body);
  _ongletsPrelevement(ss);
  const t = _tableau(ss, SHEET_MANDATS), faits = [];
  t.lignes.forEach(function (r, i) {
    const st = String(_val(t, r, 'statut'));
    if (String(_val(t, r, 'id_createur')) !== id || (st !== 'signe' && st !== 'a_signer')) return;
    r[t.M['statut']] = 'revoque'; r[t.M['revoque_le']] = new Date(); r[t.M['jeton']] = ''; r[t.M['remarque']] = motif;
    _ecrireLigne(t, i, r); faits.push(String(_val(t, r, 'rum')));
  });
  if (!faits.length) throw new Error('Aucun mandat en cours pour ce créateur.');
  _journaliser('mandat_revoque', id + ' · ' + faits.join(', ') + ' : ' + motif);
  return { ok: true };
}

/** Termine les mandats signés d'un créateur quand un nouveau est signé (changement de compte). */
function _remplacerMandats(t, id, sauf) {
  t.lignes.forEach(function (r, i) {
    if (i === sauf || String(_val(t, r, 'id_createur')) !== id || String(_val(t, r, 'statut')) !== 'signe') return;
    r[t.M['statut']] = 'remplace'; r[t.M['revoque_le']] = new Date(); _ecrireLigne(t, i, r);
  });
}

function _ibanSepa(v) {
  const iban = String(v || '').replace(/\s/g, '').toUpperCase();
  if (!_ibanValide(iban)) throw new Error('IBAN invalide : vérifie la saisie.');
  if (PAYS_SEPA.indexOf(iban.slice(0, 2)) === -1) throw new Error('Ce compte n\'est pas dans la zone SEPA : le prélèvement est impossible.');
  return iban;
}
function _bicOptionnel(v) {
  const bic = String(v || '').replace(/\s/g, '').toUpperCase();
  if (bic && !/^[A-Z]{6}[A-Z0-9]{2}([A-Z0-9]{3})?$/.test(bic)) throw new Error('BIC invalide (8 ou 11 caractères) : laisse vide si tu ne le connais pas.');
  return bic;
}

/** Mandat signé sur papier (scan à ranger dans les documents du créateur). */
function _gMandatPapier(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), id = String(body.idCreateur || '');
  if (!_icsValide(_ics())) throw new Error("Saisis d'abord l'ICS (identifiant créancier SEPA).");
  const f = _fichesParId(ss)[id];
  if (!f) throw new Error('Créateur introuvable.');
  const titulaire = _textePublic(body.titulaire, 70), iban = _ibanSepa(body.iban), bic = _bicOptionnel(body.bic), signe = _dateIso(body.signeLe);
  if (!titulaire) throw new Error('Indique le titulaire du compte, comme sur le mandat.');
  if (signe > _aujourdhui()) throw new Error('La date de signature est dans le futur.');
  _ongletsPrelevement(ss);
  const t = _tableau(ss, SHEET_MANDATS), rum = _nouvelleRum(t, id);
  _ajouterLigne(t, { rum: rum, id_createur: id, statut: 'signe', demande_le: new Date(), demande_par: _signataire(), signe_le: signe, mode: 'papier', titulaire: titulaire,
    adresse: _textePublic(body.adresse, 120), code_postal: _textePublic(body.codePostal, 10), ville: _textePublic(body.ville, 60), iban: iban, bic: bic });
  _remplacerMandats(t, id, t.lignes.length - 1);
  // une demande en ligne encore ouverte n'a plus lieu d'être
  t.lignes.forEach(function (r, i) {
    if (String(_val(t, r, 'id_createur')) === id && String(_val(t, r, 'statut')) === 'a_signer') { r[t.M['statut']] = 'remplace'; r[t.M['jeton']] = ''; _ecrireLigne(t, i, r); }
  });
  _journaliser('mandat_papier', id + ' ' + f.nom + ' · ' + rum + ' · IBAN •••• ' + iban.slice(-4) + ' · signé le ' + _dateFr(signe));
  return { ok: true, rum: rum };
}

/* ---------- Page publique mandat.html ---------- */

function _mandatLigne(ss, jeton) {
  const j = String(jeton || '');
  if (!/^[0-9a-f]{32}$/.test(j) || !ss.getSheetByName(SHEET_MANDATS)) return null;
  const t = _tableau(ss, SHEET_MANDATS), h = _empreinte(j), i = t.lignes.findIndex(function (r) { return String(_val(t, r, 'jeton')) === h; });
  return i < 0 ? null : { t: t, i: i, r: t.lignes[i] };
}
function _creancierPublic() { return { nom: ASSO_CREANCIER.nom, adresse: ASSO_CREANCIER.adresse, ville: ASSO_CREANCIER.ville, pays: ASSO_CREANCIER.pays, ics: _ics() }; }

function _mandatInfos(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), x = _mandatLigne(ss, body.jeton);
  if (!x) return { ok: false, message: 'Lien invalide ou déjà utilisé. Écris-nous à ' + EMAIL_SHOP + ' pour en recevoir un nouveau.' };
  const t = x.t, r = x.r, st = String(_val(t, r, 'statut'));
  if (st === 'signe') return { ok: true, statut: 'signe', rum: String(_val(t, r, 'rum')), signeLe: _iso(_val(t, r, 'signe_le')), ibanFin: String(_val(t, r, 'iban')).slice(-4) };
  if (st !== 'a_signer') return { ok: false, message: "Ce lien n'est plus valable. Écris-nous à " + EMAIL_SHOP + '.' };
  if (_lienExpire(_val(t, r, 'demande_le'))) return { ok: false, message: 'Ce lien a expiré. Écris-nous à ' + EMAIL_SHOP + ' pour en recevoir un nouveau.' };
  if (!_icsValide(_ics())) return { ok: false, message: 'Le mandat n\'est pas encore prêt côté Collectif. Réessaie un peu plus tard.' };
  const f = _fichesParId(ss)[String(_val(t, r, 'id_createur'))] || {};
  return { ok: true, statut: 'a_signer', rum: String(_val(t, r, 'rum')), marque: f.nom || '', titulaire: f.nomLegal || '', adresse: f.adresse || '', codePostal: f.codePostal || '',
    ville: f.ville || '', creancier: _creancierPublic() };
}

function _mandatSigner(body) {
  const cache = CacheService.getScriptCache(), cle = 'mandat_n_' + _empreinte(String(body.jeton || '')).slice(0, 20), n = Number(cache.get(cle) || 0);
  if (n >= 10) return { ok: false, message: 'Trop d\'essais : réessaie dans une heure.' };
  cache.put(cle, String(n + 1), 3600);
  if (body.accepte !== true) return { ok: false, message: 'Coche la case d\'autorisation pour signer le mandat.' };
  let titulaire, adresse, codePostal, ville, iban, bic;
  try {
    titulaire = _textePublic(body.titulaire, 70); adresse = _textePublic(body.adresse, 120); codePostal = _textePublic(body.codePostal, 10); ville = _textePublic(body.ville, 60);
    if (!titulaire) throw new Error('Indique le nom du titulaire du compte.');
    if (!adresse || !codePostal || !ville) throw new Error('Indique ton adresse complète.');
    iban = _ibanSepa(body.iban); bic = _bicOptionnel(body.bic);
  } catch (e) { return { ok: false, message: e.message }; }
  const ss = SpreadsheetApp.getActiveSpreadsheet(), lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, message: 'Le serveur est occupé : réessaie dans une minute.' };
  let rum, id, f, maintenant = new Date();
  try {
    const x = _mandatLigne(ss, body.jeton);
    if (!x) return { ok: false, message: 'Lien invalide ou déjà utilisé. Écris-nous à ' + EMAIL_SHOP + '.' };
    const t = x.t, r = x.r;
    if (String(_val(t, r, 'statut')) === 'signe') return { ok: false, deja: true, message: 'Ton mandat est déjà signé.' };
    if (String(_val(t, r, 'statut')) !== 'a_signer' || _lienExpire(_val(t, r, 'demande_le'))) return { ok: false, message: "Ce lien n'est plus valable. Écris-nous à " + EMAIL_SHOP + '.' };
    if (!_icsValide(_ics())) return { ok: false, message: 'Le mandat n\'est pas encore prêt côté Collectif. Réessaie un peu plus tard.' };
    id = String(_val(t, r, 'id_createur')); rum = String(_val(t, r, 'rum')); f = _fichesParId(ss)[id] || { nom: id, email: '' };
    const v = { statut: 'signe', signe_le: maintenant, mode: 'en_ligne', titulaire: titulaire, adresse: adresse, code_postal: codePostal, ville: ville, iban: iban, bic: bic, email_signature: f.email };
    Object.keys(v).forEach(function (k) { r[t.M[k]] = v[k]; });
    _ecrireLigne(t, x.i, r);          // le jeton reste : le lien affiche ensuite « mandat signé »
    _remplacerMandats(t, id, x.i);
    _journaliser('mandat_signe', id + ' ' + f.nom + ' · ' + rum + ' · IBAN •••• ' + iban.slice(-4) + ' · signé en ligne', f.nom + ' (site)');
  } finally { lock.releaseLock(); }
  try {
    _modeleSiAbsent(ss, 'mandat_signe');
    if (f.email) envoyerModele(ss, 'mandat_signe', f.email, { marque: f.nom, rum: rum, ics: _ics(), iban_fin: iban.slice(-4), titulaire: titulaire,
      date: Utilities.formatDate(maintenant, _tz(), "dd/MM/yyyy 'à' HH:mm") });
  } catch (e) { _journaliser('email_echec', 'mandat ' + rum + ' : ' + (e && e.message ? e.message : e), 'site public'); }
  return { ok: true, rum: rum, ibanFin: iban.slice(-4), signeLe: _iso(maintenant) };
}

/* ---------- Pré-notification, fichier, encaissement, rejet ---------- */

function _datePrelevement(s) {
  const d = _dateIso(s);
  if (!_jourOuvreSepa(d)) throw new Error('Le ' + _dateFr(d) + ' n\'est pas un jour ouvré pour les banques (week-end ou jour férié) : choisis une autre date.');
  if (d <= _aujourdhui()) throw new Error('La date de prélèvement doit être dans le futur.');
  return d;
}

function _gPrelevementsPrenotifier(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), mois = String(body.mois || ''), date = _datePrelevement(body.date), ics = _ics();
  if (!_icsValide(ics)) throw new Error("Saisis d'abord l'ICS (identifiant créancier SEPA).");
  _ongletsPrelevement(ss);
  _modeleSiAbsent(ss, 'prelevement_prenotification');
  const perso = body.emailPerso && (body.emailPerso.objet || body.emailPerso.texte) ? body.emailPerso : null;
  if (perso) _textePerso(perso, {});
  const d = _gPrelevements({ mois: mois }), fiches = _fichesParId(ss), mandats = _mandatsParCreateur(ss), voulus = (body.ids || []).map(String);
  const tP = _tableau(ss, SHEET_PRELEVEMENTS), envoyes = [], erreurs = [];
  let n = _maxId(tP, 'id_prelevement', 'PR');
  voulus.filter(function (id) { return !d.lignes.some(function (l) { return l.idCreateur === id; }); })
    .forEach(function (id) { erreurs.push(((fiches[id] || {}).nom || id) + ' : pas de mandat signé'); });
  d.lignes.filter(function (l) { return voulus.indexOf(l.idCreateur) !== -1; }).forEach(function (l) {
    const f = fiches[l.idCreateur] || { nom: l.nom, email: '' }, m = mandats[l.idCreateur];
    try {
      if (!m || m.statut !== 'signe') throw new Error('pas de mandat signé');
      if (['a_prenotifier', 'prenotifie'].indexOf(l.statut) === -1) throw new Error('déjà dans un fichier ou traité');
      if (!(l.reste > 0)) throw new Error('rien à prélever (loyer déjà reçu)');
      if (!f.email) throw new Error("pas d'e-mail valide sur la fiche");
      const montant = _round2(l.reste);
      const vars = { marque: f.nom, montant: _eurFr(montant), date: _dateFr(date), mois: d.libelle, rum: m.rum, ics: ics, iban_fin: m.ibanFin };
      const ok = perso ? envoyerTexte(ss, 'prelevement_prenotification', f.email, perso, vars) : envoyerModele(ss, 'prelevement_prenotification', f.email, vars);
      if (!ok) throw new Error('modèle « prelevement_prenotification » désactivé');
      const v = { mois_loyer: mois, id_createur: l.idCreateur, rum: m.rum, montant: montant, date_prelevement: date, sequence: m.premier ? 'RCUR' : 'FRST', statut: 'prenotifie',
        prenotifie_le: new Date(), par: _signataire() };
      const i = l.prelevement ? tP.lignes.findIndex(function (r) { return String(_val(tP, r, 'id_prelevement')) === l.prelevement.id; }) : -1;
      if (i >= 0) { const r = tP.lignes[i]; Object.keys(v).forEach(function (k) { r[tP.M[k]] = v[k]; }); _ecrireLigne(tP, i, r); }
      else { n++; v.id_prelevement = 'PR' + String(n).padStart(4, '0'); _ajouterLigne(tP, v); }
      envoyes.push(f.nom + ' ' + _eurFr(montant));
    } catch (e) { erreurs.push(f.nom + ' : ' + (e && e.message ? e.message : e)); }
  });
  if (envoyes.length) _journaliser('prelevements_prenotifies', d.libelle + ' · prélèvement le ' + _dateFr(date) + ' : ' + envoyes.join(', ') + (perso && perso.modifie ? ' (texte modifié)' : ''));
  const jours = Math.round((date - _aujourdhui()) / 864e5);
  return { ok: true, envoyes: envoyes.length, erreurs: erreurs, joursAvant: jours };
}

/** Fichier SEPA pain.008.001.02 (CORE) des lignes pré-notifiées choisies ; elles passent « dans un fichier ». */
function _gPrelevementsFichier(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), mois = String(body.mois || ''), ics = _ics(), banque = _banqueAsso(), voulus = (body.ids || []).map(String);
  if (!_icsValide(ics)) throw new Error("Saisis d'abord l'ICS (identifiant créancier SEPA).");
  _ongletsPrelevement(ss);
  const tP = _tableau(ss, SHEET_PRELEVEMENTS), tM = _tableau(ss, SHEET_MANDATS), fiches = _fichesParId(ss), auj = _aujourdhui(), erreurs = [];
  const mandat = {};
  tM.lignes.forEach(function (r) { mandat[String(_val(tM, r, 'rum'))] = r; });
  const L = [];
  tP.lignes.forEach(function (r, i) {
    if (_moisTexte(_val(tP, r, 'mois_loyer')) !== mois || String(_val(tP, r, 'statut')) !== 'prenotifie' || voulus.indexOf(String(_val(tP, r, 'id_createur'))) === -1) return;
    const id = String(_val(tP, r, 'id_createur')), m = mandat[String(_val(tP, r, 'rum'))], nom = (fiches[id] || {}).nom || id, date = _val(tP, r, 'date_prelevement');
    if (!m || String(_val(tM, m, 'statut')) !== 'signe') { erreurs.push(nom + ' : mandat révoqué ou remplacé depuis la pré-notification'); return; }
    if (!(date instanceof Date) || date <= auj) { erreurs.push(nom + ' : date de prélèvement dépassée, refais la pré-notification'); return; }
    L.push({ i: i, r: r, id: String(_val(tP, r, 'id_prelevement')), idCreateur: id, nom: nom, montant: Number(_val(tP, r, 'montant')) || 0, date: date,
      sequence: String(_val(tP, r, 'sequence') || 'RCUR'), rum: String(_val(tP, r, 'rum')), titulaire: String(_val(tM, m, 'titulaire')), iban: String(_val(tM, m, 'iban')),
      bic: String(_val(tM, m, 'bic') || ''), signeLe: _val(tM, m, 'signe_le') });
  });
  if (!L.length) throw new Error('Aucun prélèvement pré-notifié parmi ceux choisis.' + (erreurs.length ? ' ' + erreurs.join(' ; ') : ''));
  const maintenant = new Date(), lot = 'PRL-' + Utilities.formatDate(maintenant, _tz(), 'yyyyMMdd-HHmmss'), mt = function (x) { return _round2(x).toFixed(2); };
  const total = _round2(L.reduce(function (s, l) { return s + l.montant; }, 0)), libMois = _moisBornes(mois).libelle;
  const groupes = {};
  L.forEach(function (l) { const k = _jourIso(l.date) + '|' + l.sequence; (groupes[k] = groupes[k] || []).push(l); });
  const blocs = Object.keys(groupes).sort().map(function (k, gi) {
    const G = groupes[k], somme = _round2(G.reduce(function (s, l) { return s + l.montant; }, 0));
    return '<PmtInf><PmtInfId>' + lot + '-' + (gi + 1) + '</PmtInfId><PmtMtd>DD</PmtMtd><BtchBookg>true</BtchBookg><NbOfTxs>' + G.length + '</NbOfTxs><CtrlSum>' + mt(somme) + '</CtrlSum>' +
      '<PmtTpInf><SvcLvl><Cd>SEPA</Cd></SvcLvl><LclInstrm><Cd>CORE</Cd></LclInstrm><SeqTp>' + k.split('|')[1] + '</SeqTp></PmtTpInf>' +
      '<ReqdColltnDt>' + k.split('|')[0] + '</ReqdColltnDt><Cdtr><Nm>' + banque.nom + '</Nm></Cdtr><CdtrAcct><Id><IBAN>' + banque.iban + '</IBAN></Id></CdtrAcct>' +
      '<CdtrAgt><FinInstnId><BIC>' + banque.bic + '</BIC></FinInstnId></CdtrAgt><ChrgBr>SLEV</ChrgBr>' +
      '<CdtrSchmeId><Id><PrvtId><Othr><Id>' + ics + '</Id><SchmeNm><Prtry>SEPA</Prtry></SchmeNm></Othr></PrvtId></Id></CdtrSchmeId>' +
      G.map(function (l) {
        return '<DrctDbtTxInf><PmtId><EndToEndId>' + _xml(_texteSepa(l.id + '-' + mois, 35)) + '</EndToEndId></PmtId><InstdAmt Ccy="EUR">' + mt(l.montant) + '</InstdAmt>' +
          '<DrctDbtTx><MndtRltdInf><MndtId>' + _xml(_texteSepa(l.rum, 35)) + '</MndtId><DtOfSgntr>' + _jourIso(l.signeLe) + '</DtOfSgntr></MndtRltdInf></DrctDbtTx>' +
          '<DbtrAgt><FinInstnId>' + (l.bic ? '<BIC>' + l.bic + '</BIC>' : '<Othr><Id>NOTPROVIDED</Id></Othr>') + '</FinInstnId></DbtrAgt>' +
          '<Dbtr><Nm>' + _xml(_texteSepa(l.titulaire || l.nom, 70)) + '</Nm></Dbtr><DbtrAcct><Id><IBAN>' + _xml(l.iban) + '</IBAN></Id></DbtrAcct>' +
          '<RmtInf><Ustrd>' + _xml(_texteSepa(('13H59 LOYER ' + libMois + ' ' + l.nom).toUpperCase(), 140)) + '</Ustrd></RmtInf></DrctDbtTxInf>';
      }).join('') + '</PmtInf>';
  }).join('');
  const xml = '<?xml version="1.0" encoding="UTF-8"?>\n<Document xmlns="urn:iso:std:iso:20022:tech:xsd:pain.008.001.02" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    '<CstmrDrctDbtInitn><GrpHdr><MsgId>' + lot + '</MsgId><CreDtTm>' + Utilities.formatDate(maintenant, _tz(), "yyyy-MM-dd'T'HH:mm:ss") + '</CreDtTm>' +
    '<NbOfTxs>' + L.length + '</NbOfTxs><CtrlSum>' + mt(total) + '</CtrlSum><InitgPty><Nm>' + banque.nom + '</Nm></InitgPty></GrpHdr>' + blocs + '</CstmrDrctDbtInitn></Document>\n';
  L.forEach(function (l) { l.r[tP.M['statut']] = 'fichier'; l.r[tP.M['lot']] = lot; l.r[tP.M['fichier_le']] = maintenant; _ecrireLigne(tP, l.i, l.r); });
  _journaliser('prelevements_fichier', lot + ' · ' + L.length + ' prélèvement(s) · ' + _eurFr(total) + ' : ' + L.map(function (l) { return l.nom + ' ' + _eurFr(l.montant); }).join(', '));
  return { ok: true, lot: lot, nb: L.length, total: total, xml: xml, nom: '13H59-prelevements-' + lot.slice(4) + '.xml', erreurs: erreurs,
    detail: L.map(function (l) { return { nom: l.titulaire || l.nom, ibanFin: l.iban.slice(-4), montant: l.montant, rum: l.rum, date: _iso(l.date), sequence: l.sequence }; }) };
}

/** L'argent du lot est arrivé : les loyers sont pointés reçus (source prelevement). */
function _gPrelevementsValider(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), lot = String(body.lot || '');
  _ongletsPrelevement(ss); _ongletsEncaissement(ss);
  const tP = _tableau(ss, SHEET_PRELEVEMENTS), tR = _tableau(ss, SHEET_PAIEMENTS_RECUS), tM = _tableau(ss, SHEET_MANDATS), faits = [];
  let n = Math.max(_maxId(tR, 'id_paiement', 'P'), _dernierId('DERNIER_ID_P'));
  tP.lignes.forEach(function (r, i) {
    if (String(_val(tP, r, 'lot')) !== lot || String(_val(tP, r, 'statut')) !== 'fichier') return;
    n++;
    const idP = 'P' + String(n).padStart(4, '0'), rum = String(_val(tP, r, 'rum'));
    _ajouterLigne(tR, { id_paiement: idP, mois_loyer: _moisTexte(_val(tP, r, 'mois_loyer')), id_createur: String(_val(tP, r, 'id_createur')), recu_le: _val(tP, r, 'date_prelevement'),
      montant: Number(_val(tP, r, 'montant')) || 0, source: 'prelevement', libelle: 'Prélèvement SEPA ' + rum + ' (' + lot + ')', statut: 'ok', pointe_par: _signataire(), pointe_le: new Date() });
    r[tP.M['statut']] = 'preleve'; r[tP.M['id_paiement']] = idP; _ecrireLigne(tP, i, r);
    const k = tM.lignes.findIndex(function (m) { return String(_val(tM, m, 'rum')) === rum; });
    if (k >= 0 && !(_val(tM, tM.lignes[k], 'premier_prelevement_le') instanceof Date)) { const m = tM.lignes[k]; m[tM.M['premier_prelevement_le']] = _val(tP, r, 'date_prelevement'); _ecrireLigne(tM, k, m); }
    faits.push(String(_val(tP, r, 'id_createur')) + ' ' + _eurFr(_val(tP, r, 'montant')));
  });
  if (!faits.length) throw new Error('Rien à valider dans ce fichier.');
  _journaliser('prelevements_encaisses', lot + ' : ' + faits.join(', '));
  return { ok: true, nb: faits.length };
}

/** Fichier jamais déposé (ou refusé par la banque) : les lignes redeviennent « pré-notifiées ». */
function _gPrelevementsAnnulerLot(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), lot = String(body.lot || ''), motif = _motif(body), tP = _tableau(ss, SHEET_PRELEVEMENTS);
  let n = 0;
  tP.lignes.forEach(function (r, i) {
    if (String(_val(tP, r, 'lot')) !== lot || String(_val(tP, r, 'statut')) !== 'fichier') return;
    r[tP.M['statut']] = 'prenotifie'; r[tP.M['lot']] = ''; r[tP.M['fichier_le']] = ''; r[tP.M['remarque']] = 'fichier ' + lot + ' annulé : ' + motif; _ecrireLigne(tP, i, r); n++;
  });
  if (!n) throw new Error('Rien à annuler dans ce fichier.');
  _journaliser('prelevements_fichier_annule', lot + ' (' + n + ') : ' + motif);
  return { ok: true, nb: n };
}

/** Prélèvement rejeté par la banque du créateur : le loyer redevient dû. */
function _gPrelevementRejete(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), motif = _motif(body), tP = _tableau(ss, SHEET_PRELEVEMENTS);
  const i = tP.lignes.findIndex(function (r) { return String(_val(tP, r, 'id_prelevement')) === String(body.idPrelevement); });
  if (i < 0) throw new Error('Prélèvement introuvable.');
  const r = tP.lignes[i], st = String(_val(tP, r, 'statut'));
  if (st !== 'preleve' && st !== 'fichier') throw new Error('Seul un prélèvement envoyé à la banque peut être rejeté.');
  const idP = String(_val(tP, r, 'id_paiement') || '');
  if (idP) {
    const tR = _tableau(ss, SHEET_PAIEMENTS_RECUS), k = tR.lignes.findIndex(function (x) { return String(_val(tR, x, 'id_paiement')) === idP; });
    if (k >= 0) { const x = tR.lignes[k]; x[tR.M['statut']] = 'annule'; x[tR.M['remarque']] = 'prélèvement rejeté : ' + motif; _ecrireLigne(tR, k, x); }
  }
  r[tP.M['statut']] = 'rejete'; r[tP.M['remarque']] = motif; _ecrireLigne(tP, i, r);
  _journaliser('prelevement_rejete', body.idPrelevement + ' (' + _val(tP, r, 'id_createur') + ', ' + _eurFr(_val(tP, r, 'montant')) + ') : ' + motif);
  return { ok: true };
}
