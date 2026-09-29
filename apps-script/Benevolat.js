/*************************************************************
 *  BÉNÉVOLAT
 *   - page publique « benevolat.html » : un créateur propose ses
 *     permanences (stand, nombre par mois, disponibilités, formation)
 *   - Gestion ▸ Bénévolat : demandes, rattachement à la fiche créateur,
 *     « Retenir » coche « bénévole » sur la fiche avec le nombre de permanences
 *  Les permanences elles-mêmes restent dans l'agenda BÉNÉVOLES (lu par la facturation).
 *  Une même adresse e-mail = une seule demande (mise à jour si elle revient).
 *************************************************************/

const SHEET_BENEVOLAT = 'benevolat';
const COLONNES_BENEVOLAT = ['id_demande', 'recue_le', 'maj_le', 'nom', 'marque', 'email', 'id_createur', 'stand', 'perms_souhaitees',
  'disponibilites', 'formation', 'statut', 'traitee_le', 'traitee_par', 'remarque', 'source'];
const STATUTS_BENEVOLAT = ['a_traiter', 'retenu', 'plus_tard', 'non_retenu'];
const JOURS_BENEVOLAT = ['mar', 'mer', 'jeu', 'ven', 'sam'];
const CRENEAUX_BENEVOLAT = ['matin', 'apres_midi', 'journee'];
/** Permanences au plus par mois selon le stand (formulaire de l'équipe) ; pas de bénévolat en friperie. */
const PERMS_MAX = { illu: 2, unique: 3, grand: 4 };

/** Stands ouverts au bénévolat : { code, libelle, loyer, perms (loyer offert), max }. */
function _standsBenevolat(ss) {
  return _lireTable(_onglet(ss, SHEET_STANDS)).filter(function (s) { return s['code']; }).map(function (s) {
    const code = String(s['code']);
    const perms = s['perms_loyer_gratuit'] !== '' && s['perms_loyer_gratuit'] != null ? Number(s['perms_loyer_gratuit']) : PERMS_PAR_DEFAUT[code];
    return { code: code, libelle: String(s['libelle'] || code), loyer: Number(s['loyer']) || 0, perms: perms, max: Math.max(perms || 0, PERMS_MAX[code] || 0) };
  }).filter(function (s) { return s.perms > 0; });
}

/** « mar:matin jeu:journee » → validé ; les entrées inconnues sont ignorées. */
function _disponibilites(v) {
  const out = {};
  String(v == null ? '' : v).split(/\s+/).forEach(function (x) {
    const p = x.split(':');
    if (JOURS_BENEVOLAT.indexOf(p[0]) >= 0 && CRENEAUX_BENEVOLAT.indexOf(p[1]) >= 0) out[p[0]] = p[1];
  });
  return JOURS_BENEVOLAT.filter(function (j) { return out[j]; }).map(function (j) { return j + ':' + out[j]; }).join(' ');
}

/* ---------- Page publique ---------- */

function _benevolatInfos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return { ok: true, stands: _standsBenevolat(ss).map(function (s) { return { code: s.code, libelle: s.libelle, loyer: s.loyer, perms: s.perms, max: s.max }; }) };
}

function _benevolatEnvoyer(body) {
  if (body.site) return { ok: true };                       // champ piège rempli : robot
  if (!(Number(body.duree) >= 4000)) return { ok: false, message: 'Envoi trop rapide : relis ta demande puis renvoie-la.' };
  const cache = CacheService.getScriptCache();
  const nbHeure = Number(cache.get('benevolat_heure') || 0);
  if (nbHeure >= CANDIDATURES_PAR_HEURE) return { ok: false, message: 'Beaucoup de demandes en ce moment : réessaie dans une heure, ou écris-nous à ' + EMAIL_SHOP + '.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stands = _standsBenevolat(ss);
  const d = {
    nom: _nomPropre(_textePublic(body.nom, 80)), marque: _nomPropre(_textePublic(body.marque, 80)), email: _email(body.email),
    stand: String(body.stand || ''), perms: Number(body.perms), dispos: _disponibilites(body.disponibilites),
    formation: body.formation === 'formation' ? 'formation' : body.formation === 'habitude' ? 'habitude' : ''
  };
  const s = stands.filter(function (x) { return x.code === d.stand; })[0];
  const erreurs = [];
  if (!d.nom) erreurs.push('ton nom et prénom');
  if (!d.marque) erreurs.push('ton nom de créateur');
  if (!d.email) erreurs.push('une adresse e-mail valide');
  if (!s) erreurs.push('ton stand actuel');
  else if (!(d.perms >= 1 && d.perms <= s.max && d.perms === Math.floor(d.perms))) erreurs.push('le nombre de permanences (1 à ' + s.max + ' pour ce stand)');
  if (!d.dispos) erreurs.push('au moins une disponibilité');
  if (!d.formation) erreurs.push('si tu as besoin d\'une journée de formation');
  if (erreurs.length) return { ok: false, message: 'Il manque ' + erreurs.join(', ') + '.' };

  const cleEmail = 'benevolat_' + d.email;
  if (cache.get(cleEmail)) return { ok: true, deja: true };   // double clic

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, message: 'Le serveur est occupé : réessaie dans une minute.' };
  let deja = false, id = '';
  try {
    _creerOngletSiAbsent(ss, SHEET_BENEVOLAT, COLONNES_BENEVOLAT);
    _modeleSiAbsent(ss, 'benevolat_recu');
    const tB = _tableau(ss, SHEET_BENEVOLAT), tC = _tableau(ss, SHEET_CREATEURS);
    const createur = _createurDeLaDemande(tC, d.email, d.marque);
    const valeurs = { maj_le: new Date(), nom: d.nom, marque: d.marque, stand: d.stand, perms_souhaitees: d.perms, disponibilites: d.dispos, formation: d.formation };
    const i = tB.lignes.findIndex(function (r) { return _norm(_val(tB, r, 'email')) === d.email; });
    if (i >= 0) {
      deja = true;
      const r = tB.lignes[i];
      id = String(_val(tB, r, 'id_demande'));
      Object.keys(valeurs).forEach(function (k) { if (tB.M[k] != null) r[tB.M[k]] = valeurs[k]; });
      if (_val(tB, r, 'statut') !== 'a_traiter') { r[tB.M['statut']] = 'a_traiter'; r[tB.M['traitee_le']] = ''; r[tB.M['traitee_par']] = ''; }
      if (createur && !_val(tB, r, 'id_createur')) r[tB.M['id_createur']] = createur;
      _ecrireLigne(tB, i, r);
      _journaliser('benevolat_maj', id + ' ' + d.marque + ' (nouvelle demande de la même adresse)', 'site public');
    } else {
      id = _prochainId(tB, 'id_demande', 'BV');
      valeurs.id_demande = id; valeurs.recue_le = new Date(); valeurs.email = d.email;
      valeurs.statut = 'a_traiter'; valeurs.source = 'site'; valeurs.id_createur = createur || '';
      _ajouterLigne(tB, valeurs);
      _journaliser('benevolat_recu', id + ' ' + d.marque + ' · ' + d.stand + ' · ' + d.perms + ' permanence(s)', 'site public');
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  cache.put(cleEmail, '1', 120);
  cache.put('benevolat_heure', String(nbHeure + 1), 3600);
  try {
    envoyerModele(ss, 'benevolat_recu', d.email, { prenom: String(d.nom).split(' ')[0], marque: d.marque, stand: s.libelle });
  } catch (e) {
    _journaliser('email_echec', id + ' bénévolat : ' + (e && e.message ? e.message : e), 'site public');
  }
  return { ok: true, deja: deja };
}

/** Fiche créateur de la demande : par e-mail, sinon par marque si un seul créateur porte ce nom. */
function _createurDeLaDemande(tC, email, marque) {
  const parEmail = tC.lignes.filter(function (r) { return email && _norm(_val(tC, r, 'email')) === email; });
  if (parEmail.length === 1) return String(_val(tC, parEmail[0], 'id_createur'));
  const k = _cleLarge(marque);
  const parNom = k ? tC.lignes.filter(function (r) { return _cleLarge(_val(tC, r, 'nom')) === k; }) : [];
  return parNom.length === 1 ? String(_val(tC, parNom[0], 'id_createur')) : '';
}

/* ---------- Gestion ▸ Bénévolat ---------- */

function _gBenevolat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _creerOngletSiAbsent(ss, SHEET_BENEVOLAT, COLONNES_BENEVOLAT);
  _modeleSiAbsent(ss, 'benevolat_recu');
  const tB = _tableau(ss, SHEET_BENEVOLAT), tC = _tableau(ss, SHEET_CREATEURS), tE = _tableau(ss, SHEET_EMPLACEMENTS), auj = _aujourdhui();
  // stand en cours de chaque créateur
  const standDe = {};
  tE.lignes.forEach(function (e) { if (_actifLe(tE, e, auj)) standDe[String(_val(tE, e, 'id_createur'))] = String(_val(tE, e, 'code_stand')); });
  const createurs = tC.lignes.filter(function (r) { return _val(tC, r, 'id_createur'); }).map(function (r) {
    const id = String(_val(tC, r, 'id_createur'));
    return { id: id, nom: _nomPropre(_val(tC, r, 'nom')), statut: _norm(_val(tC, r, 'statut')), stand: standDe[id] || '',
      benevole: _val(tC, r, 'benevole') === true, permsPrevues: Number(_val(tC, r, 'perms_prevues')) || '' };
  });
  const parId = {};
  createurs.forEach(function (c) { parId[c.id] = c; });
  const s = function (r, k) { const v = _val(tB, r, k); return v instanceof Date ? _jourIso(v) : String(v == null ? '' : v); };
  const demandes = tB.lignes.filter(function (r) { return _val(tB, r, 'id_demande'); }).map(function (r) {
    const idC = s(r, 'id_createur');
    return { id: s(r, 'id_demande'), recueLe: s(r, 'recue_le'), majLe: s(r, 'maj_le'), nom: s(r, 'nom'), marque: s(r, 'marque'), email: s(r, 'email'),
      idCreateur: idC, createur: parId[idC] || null, stand: s(r, 'stand'), perms: Number(_val(tB, r, 'perms_souhaitees')) || 0,
      disponibilites: s(r, 'disponibilites'), formation: s(r, 'formation'), statut: s(r, 'statut') || 'a_traiter',
      traiteeLe: s(r, 'traitee_le'), traiteePar: s(r, 'traitee_par'), remarque: s(r, 'remarque'), source: s(r, 'source') };
  });
  return { ok: true, aujourdhui: _jourIso(auj), demandes: demandes, stands: _standsBenevolat(ss),
    createurs: createurs.filter(function (c) { return c.statut === 'actif' || c.benevole; }) };
}

function _ligneDemande(tB, id) {
  const i = tB.lignes.findIndex(function (r) { return String(_val(tB, r, 'id_demande')) === String(id); });
  if (i < 0) throw new Error('Demande introuvable : « ' + id + ' ».');
  return i;
}

/** Statut d'une demande. « retenu » coche « bénévole » sur la fiche, avec le nombre de permanences prévu. */
function _gBenevolatStatut(body) {
  const statut = String(body.statut || '');
  if (STATUTS_BENEVOLAT.indexOf(statut) < 0) throw new Error('Statut inconnu : « ' + statut + ' ».');
  const ss = SpreadsheetApp.getActiveSpreadsheet(), tB = _tableau(ss, SHEET_BENEVOLAT), i = _ligneDemande(tB, body.id), r = tB.lignes[i];
  let fiche = '';
  if (statut === 'retenu') {
    const idC = String(_val(tB, r, 'id_createur') || '');
    if (!idC) throw new Error('Rattache d\'abord la demande à une fiche créateur.');
    const tC = _tableau(ss, SHEET_CREATEURS), j = tC.lignes.findIndex(function (x) { return String(_val(tC, x, 'id_createur')) === idC; });
    if (j < 0) throw new Error('Fiche créateur introuvable : « ' + idC + ' ».');
    const perms = Number(body.perms) > 0 ? Number(body.perms) : Number(_val(tB, r, 'perms_souhaitees')) || '';
    if (perms !== '' && !(perms >= 0.5 && perms <= 8)) throw new Error('Nombre de permanences invalide.');
    const c = tC.lignes[j];
    c[tC.M['benevole']] = true;
    if (tC.M['perms_prevues'] != null) c[tC.M['perms_prevues']] = perms;
    if (tC.M['modifie_le'] != null) c[tC.M['modifie_le']] = new Date();
    _ecrireLigne(tC, j, c);
    fiche = _nomPropre(_val(tC, c, 'nom')) + ' (' + idC + ') : bénévole, ' + perms + ' permanence(s) par mois';
  }
  r[tB.M['statut']] = statut;
  r[tB.M['traitee_le']] = statut === 'a_traiter' ? '' : new Date();
  r[tB.M['traitee_par']] = statut === 'a_traiter' ? '' : _signataire();
  _ecrireLigne(tB, i, r);
  _journaliser('benevolat_statut', body.id + ' ' + _val(tB, r, 'marque') + ' → ' + statut + (fiche ? ' · ' + fiche : ''));
  return { ok: true, fiche: fiche };
}

/** Planning : une permanence invite une adresse inconnue → on la rattache à une fiche (colonne emails_agenda). */
function _gBenevolatAssocier(body) {
  const email = _email(body.email), id = String(body.idCreateur || ''), ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!email) throw new Error('Adresse e-mail invalide.');
  const t = _assurerColonne(ss, SHEET_CREATEURS, 'emails_agenda'), i = t.lignes.findIndex(function (r) { return String(_val(t, r, 'id_createur')) === id; });
  if (i < 0) throw new Error('Créateur inconnu : « ' + id + ' ».');
  // une adresse ne désigne qu'une fiche : on la retire des autres
  t.lignes.forEach(function (r, k) {
    const l = _emailsAgenda(_val(t, r, 'emails_agenda'));
    if (k !== i && l.indexOf(email) >= 0) { r[t.M['emails_agenda']] = l.filter(function (m) { return m !== email; }).join(', '); _ecrireLigne(t, k, r); }
  });
  const r = t.lignes[i], l = _emailsAgenda(_val(t, r, 'emails_agenda'));
  if (_norm(_val(t, r, 'email')) !== email && l.indexOf(email) < 0) { l.push(email); r[t.M['emails_agenda']] = l.join(', '); _ecrireLigne(t, i, r); }
  if (body.mois) CacheService.getScriptCache().remove('permanences_' + body.mois);   // la facturation relit l'agenda
  _journaliser('benevolat_agenda', email + ' → ' + _nomPropre(_val(t, r, 'nom')) + ' (' + id + ')');
  return { ok: true };
}

function _gBenevolatLier(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), tB = _tableau(ss, SHEET_BENEVOLAT), i = _ligneDemande(tB, body.id), r = tB.lignes[i];
  const idC = String(body.idCreateur || '');
  if (idC && !_createurs(ss).parId[idC]) throw new Error('Créateur inconnu : « ' + idC + ' ».');
  r[tB.M['id_createur']] = idC;
  _ecrireLigne(tB, i, r);
  _journaliser('benevolat_lien', body.id + ' ' + _val(tB, r, 'marque') + ' → ' + (idC || 'aucune fiche'));
  return { ok: true };
}

function _gBenevolatRemarque(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), tB = _tableau(ss, SHEET_BENEVOLAT), i = _ligneDemande(tB, body.id), r = tB.lignes[i];
  r[tB.M['remarque']] = _textePublic(body.remarque, 1000);
  _ecrireLigne(tB, i, r);
  return { ok: true };
}

/* ---------- Gestion ▸ Bénévolat ▸ Planning (agenda BÉNÉVOLES) ----------
 * Le site lit les permanences du mois dans l'agenda et peut en poser ou en retirer.
 * Poser = créer un événement invitant le ou la bénévole avec l'e-mail de sa fiche :
 * Google lui envoie l'invitation. Les filles peuvent toujours modifier l'agenda à la main.
 * Il faut le droit « Apporter des modifications aux événements » sur l'agenda BÉNÉVOLES. */

const CRENEAUX_HEURES = { matin: [11, 15], apres_midi: [15, 19], journee: [11, 19] };

function _agendaBenevoles() {
  const src = _agendasPermanences();
  if (!src.dedie) throw new Error("Aucun agenda dont le nom contient « bénévole » n'est visible par le compte du site.");
  return src.agendas[0];
}
/** Créneau d'un événement : journée (≥ 6 h ou toute la journée), sinon matin ou après-midi selon l'heure de début. */
function _creneauEvenement(e) {
  if (e.isAllDayEvent() || (e.getEndTime() - e.getStartTime()) / 36e5 >= 6) return 'journee';
  return e.getStartTime().getHours() < 13 ? 'matin' : 'apres_midi';
}

function _gBenevolatPlanning(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), mois = String(body.mois || ''), b = _moisBornes(mois);
  const g = _gBenevolat();
  const stands = {};
  g.stands.forEach(function (s) { stands[s.code] = s; });
  const tC = _tableau(ss, SHEET_CREATEURS), emailDe = {}, autresDe = {};
  tC.lignes.forEach(function (r) { const id = String(_val(tC, r, 'id_createur')); emailDe[id] = _norm(_val(tC, r, 'email')); autresDe[id] = _emailsAgenda(_val(tC, r, 'emails_agenda')); });
  // disponibilités : la demande la plus récente de chaque créateur
  const dispoDe = {};
  g.demandes.slice().sort(function (x, y) { return String(x.majLe || x.recueLe).localeCompare(String(y.majLe || y.recueLe)); })
    .forEach(function (d) { if (d.idCreateur) dispoDe[d.idCreateur] = d.disponibilites; });
  const benevoles = g.createurs.filter(function (c) { return c.benevole && !_horsBenevolat(emailDe[c.id], ''); }).map(function (c) {
    const s = stands[c.stand];
    return { id: c.id, nom: c.nom, stand: c.stand, quota: s ? s.perms : null, permsPrevues: c.permsPrevues, email: emailDe[c.id] || '', disponibilites: dispoDe[c.id] || '' };
  });
  const parEmail = {}, parMots = {};
  benevoles.forEach(function (c) { if (c.email) parEmail[c.email] = c; (autresDe[c.id] || []).forEach(function (m) { parEmail[m] = c; }); const k = _cleMots(c.nom); if (k) (parMots[k] = parMots[k] || []).push(c); });

  let agenda, evenements = [], agendaErreur = '';
  try {
    agenda = _agendaBenevoles();
    const fin = new Date(b.fin.getFullYear(), b.fin.getMonth(), b.fin.getDate() + 1);
    evenements = agenda.getEvents(b.debut, fin).map(function (e) {
      const invites = e.getGuestList().map(function (x) { return { email: _norm(x.getEmail()), nom: x.getName() }; });
      const qui = [];
      invites.forEach(function (x) {
        const c = parEmail[x.email] || (x.nom && parMots[_cleMots(x.nom)] && parMots[_cleMots(x.nom)].length === 1 ? parMots[_cleMots(x.nom)][0] : null);
        if (c && qui.indexOf(c.id) < 0) qui.push(c.id);
      });
      const creneau = _creneauEvenement(e);
      return { id: e.getId(), date: _jourIso(e.getStartTime()), creneau: creneau, jours: creneau === 'journee' ? 1 : 0.5,
        heure: e.isAllDayEvent() ? '' : _heure(e.getStartTime()) + '-' + _heure(e.getEndTime()), titre: e.getTitle(),
        benevoles: qui, invites: qui.length ? [] : invites.map(function (x) { return x.email; }).filter(function (m) { return m !== EMAIL_SHOP; }),
        recurrent: e.isRecurringEvent() };
    });
  } catch (e) {
    agendaErreur = String(e && e.message ? e.message : e);
  }
  return { ok: true, mois: mois, libelle: b.libelle, aujourdhui: g.aujourdhui, agenda: agenda ? agenda.getName() : '', agendaErreur: agendaErreur,
    benevoles: benevoles, evenements: evenements };
}

function _gBenevolatPoser(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(body.date || ''));
  if (!m) throw new Error('Date invalide.');
  const jour = new Date(+m[1], +m[2] - 1, +m[3]);
  if (jour.getDay() < 2) throw new Error('La boutique est fermée le dimanche et le lundi.');
  const h = CRENEAUX_HEURES[body.creneau];
  if (!h) throw new Error('Créneau inconnu.');
  const c = _lireTable(_onglet(ss, SHEET_CREATEURS)).filter(function (r) { return String(r['id_createur']) === String(body.idCreateur); })[0];
  if (!c) throw new Error('Créateur inconnu : « ' + body.idCreateur + ' ».');
  if (c['benevole'] !== true) throw new Error("La fiche de " + _nomPropre(c['nom']) + " n'est pas cochée « bénévole » : retiens d'abord sa demande (ou coche la case sur la fiche).");
  const email = _email(c['email']);
  if (!email) throw new Error(_nomPropre(c['nom']) + " n'a pas d'adresse e-mail valide sur sa fiche : impossible de l'inviter.");
  const agenda = _agendaBenevoles();
  const debut = new Date(+m[1], +m[2] - 1, +m[3], h[0]), fin = new Date(+m[1], +m[2] - 1, +m[3], h[1]);
  const deja = agenda.getEventsForDay(jour).filter(function (e) { return e.getGuestList().some(function (g) { return _norm(g.getEmail()) === email; }); });
  if (deja.length) throw new Error(_nomPropre(c['nom']) + ' a déjà une permanence ce jour-là (' + deja[0].getTitle() + ').');
  let e;
  try {
    e = agenda.createEvent('Permanence ' + _nomPropre(c['nom']), debut, fin, {
      guests: email, sendInvites: fin > new Date(), description: 'Permanence posée depuis le site 13H59 par ' + _signataire() + '.'
    });
  } catch (err) {
    throw new Error("Impossible d'écrire dans l'agenda « " + agenda.getName() + ' » : le compte du site doit avoir le droit « Apporter des modifications aux événements » sur cet agenda. (' + (err && err.message ? err.message : err) + ')');
  }
  CacheService.getScriptCache().remove('permanences_' + m[1] + '-' + m[2]);
  _journaliser('permanence_posee', _nomPropre(c['nom']) + ' (' + c['id_createur'] + ') · ' + _dateFr(jour) + ' · ' + h[0] + 'h-' + h[1] + 'h');
  return { ok: true, id: e.getId(), invitation: fin > new Date() };
}

function _gBenevolatRetirer(body) {
  const agenda = _agendaBenevoles(), e = agenda.getEventById(String(body.idEvenement || ''));
  if (!e) throw new Error("Permanence introuvable dans l'agenda (déjà supprimée ?) : actualise.");
  if (e.isRecurringEvent()) throw new Error('Permanence répétée (événement récurrent) : modifie-la directement dans Google Agenda.');
  const titre = e.getTitle(), jour = e.getStartTime(), heure = e.isAllDayEvent() ? '' : ' · ' + _heure(e.getStartTime()) + '-' + _heure(e.getEndTime());
  try { e.deleteEvent(); } catch (err) {
    throw new Error("Impossible de modifier l'agenda « " + agenda.getName() + ' » : le compte du site doit avoir le droit « Apporter des modifications aux événements ». (' + (err && err.message ? err.message : err) + ')');
  }
  CacheService.getScriptCache().remove('permanences_' + _jourIso(jour).slice(0, 7));
  _journaliser('permanence_retiree', titre + ' · ' + _dateFr(jour) + heure);
  return { ok: true };
}

/** À lancer depuis l'éditeur : autorise l'écriture dans l'agenda et vérifie les droits sur l'agenda BÉNÉVOLES.
 *  Le test crée puis supprime aussitôt un événement sans invité daté du 1er janvier 2000 (invisible). */
function autoriserPlanning() {
  const agenda = _agendaBenevoles();
  Logger.log('Agenda : « ' + agenda.getName() + ' » · compte : ' + Session.getEffectiveUser().getEmail());
  try {
    const e = agenda.createEvent('Test 13H59 (supprimé aussitôt)', new Date(2000, 0, 1, 11), new Date(2000, 0, 1, 12));
    e.deleteEvent();
    Logger.log('✅ Droits OK : le site peut poser et retirer des permanences.');
  } catch (err) {
    Logger.log('❌ Lecture seule : il faut le droit « Apporter des modifications aux événements » sur « ' + agenda.getName() + ' » (à donner depuis le compte du shop). Détail : ' + (err && err.message ? err.message : err));
  }
}
