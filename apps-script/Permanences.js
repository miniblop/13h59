/*************************************************************
 *  PERMANENCES — lues dans le Google Agenda du shop
 *  Un événement par permanence, le ou la bénévole invité·e avec son e-mail.
 *  S'il existe un agenda dont le nom contient « bénévole », on lit celui-là ;
 *  sinon on parcourt tous les agendas visibles, en ne gardant que les
 *  événements où est invité un créateur marqué bénévole dans sa fiche. Une journée = 1, une demi-journée (moins de 6 h,
 *  ex. 11 h → 15 h) = 0,5. Le créateur est retrouvé par l'e-mail de l'invité.
 *  Lecture seule : le site ne modifie jamais l'agenda.
 *************************************************************/

const NOM_AGENDA_BENEVOLES = /b[ée]n[ée]vole/i;
/** Le compte du shop organise les événements (il figure parmi les invités) ; l'équipe n'est pas bénévole. */
function _horsBenevolat(email, categorie) { return _norm(email) === EMAIL_SHOP || String(categorie) === 'shop'; }

/** { agendas: [...], dedie: true si un agenda « bénévoles » existe (sinon tous les agendas, filtrés sur les bénévoles) }. */
function _agendasPermanences() {
  const tous = CalendarApp.getAllCalendars(), dedies = tous.filter(function (c) { return NOM_AGENDA_BENEVOLES.test(c.getName()); });
  return dedies.length ? { agendas: dedies, dedie: true } : { agendas: tous, dedie: false };
}

/** Permanences d'un mois ('yyyy-MM') : { parCreateur: {id: {jours, detail[]}}, nonRattachees: [...] }. */
function _permanencesDuMois(ss, mois) {
  const b = _moisBornes(mois), fin = new Date(b.fin.getFullYear(), b.fin.getMonth(), b.fin.getDate() + 1);
  const src = _agendasPermanences(), parEmail = {}, parMots = {};
  _lireTable(_onglet(ss, SHEET_CREATEURS)).forEach(function (c) {
    // agenda dédié : tout créateur invité compte ; sinon seulement les créateurs marqués bénévoles
    if (!(src.dedie || c['benevole'] === true) || _horsBenevolat(c['email'], c['categorie'])) return;
    const f = { id: String(c['id_createur']), nom: _nomPropre(c['nom']) };
    if (c['email']) parEmail[_norm(c['email'])] = f;
    [c['nom_legal'], c['nom']].forEach(function (n) { const k = _cleMots(n); if (k.indexOf(' ') > 0) (parMots[k] = parMots[k] || []).push(f); });
  });
  // invité retrouvé par son e-mail, sinon par son nom s'il ne désigne qu'une seule fiche
  const retrouver = function (g) {
    const f = parEmail[_norm(g.getEmail())];
    if (f) return f;
    const c = g.getName() ? parMots[_cleMots(g.getName())] : null;
    return c && c.length === 1 ? c[0] : null;
  };
  const parCreateur = {}, nonRattachees = [], vus = {};
  const evenements = [];
  src.agendas.forEach(function (a) { a.getEvents(b.debut, fin).forEach(function (e) { evenements.push({ e: e, agenda: a.getName() }); }); });
  evenements.forEach(function (x) {
    const e = x.e;
    if (vus[e.getId()]) return;                    // même événement visible dans deux agendas
    vus[e.getId()] = true;
    const heures = (e.getEndTime() - e.getStartTime()) / 36e5;
    const jours = e.isAllDayEvent() || heures >= 6 ? 1 : 0.5;
    const quand = _dateFr(e.getStartTime()) + (e.isAllDayEvent() ? '' : ' ' + _heure(e.getStartTime()) + '-' + _heure(e.getEndTime()));
    const guests = e.getGuestList(), invites = guests.map(function (g) { return _norm(g.getEmail()); });
    const trouves = guests.map(retrouver).filter(Boolean);
    if (!trouves.length) {
      if (src.dedie) nonRattachees.push(quand + ' « ' + e.getTitle() + ' »' + (invites.length ? ' (invité : ' + invites.join(', ') + ')' : ' (aucun invité)'));
      return;
    }
    trouves.forEach(function (c) {
      const p = parCreateur[c.id] = parCreateur[c.id] || { nom: c.nom, jours: 0, detail: [] };
      p.jours += jours; p.detail.push(quand + ' · ' + String(jours).replace('.', ',') + (src.dedie ? '' : ' · agenda « ' + x.agenda + ' »'));
    });
  });
  return { parCreateur: parCreateur, nonRattachees: nonRattachees, source: src.dedie ? 'agenda « ' + src.agendas[0].getName() + ' »' : 'tous les agendas visibles (' + src.agendas.length + ')' };
}

/** Permanences du mois par créateur { id: {jours, detail[]} }, gardées 5 minutes (l'agenda est lent à lire). */
function _permanencesCache(ss, mois, rafraichir) {
  const cache = CacheService.getScriptCache(), cle = 'permanences_' + mois;
  if (!rafraichir) { const d = cache.get(cle); if (d) return JSON.parse(d); }
  const p = _permanencesDuMois(ss, mois).parCreateur, out = {};
  Object.keys(p).forEach(function (id) { out[id] = { jours: p[id].jours, detail: p[id].detail }; });
  try { cache.put(cle, JSON.stringify(out), 300); } catch (e) { /* trop gros pour le cache : sans effet */ }
  return out;
}

/** À lancer depuis l'éditeur : vérifie l'accès à l'agenda et montre ce qui serait compté (n'écrit rien). */
function diagnosticPermanences() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Agendas visibles par ' + Session.getEffectiveUser().getEmail() + ' : ' + CalendarApp.getAllCalendars().map(function (c) { return c.getName(); }).join(' · '));
  const src = _agendasPermanences();
  Logger.log(src.dedie ? '✅ Agenda des bénévoles : « ' + src.agendas[0].getName() + ' »'
    : 'Pas d\'agenda nommé « bénévoles » : lecture de tous les agendas, en ne gardant que les créateurs marqués bénévoles (' +
      _lireTable(_onglet(ss, SHEET_CREATEURS)).filter(function (c) { return c['benevole'] === true; }).map(function (c) { return _nomPropre(c['nom']); }).join(', ') + ').');
  const auj = new Date();
  [-1, 0].forEach(function (k) {
    const d = new Date(auj.getFullYear(), auj.getMonth() + k, 1), mois = d.getFullYear() + '-' + _deux(d.getMonth() + 1);
    const p = _permanencesDuMois(ss, mois), ids = Object.keys(p.parCreateur);
    Logger.log('\n' + _moisBornes(mois).libelle.toUpperCase() + ' (' + p.source + ') : ' + ids.length + ' bénévole(s) retrouvé(s)' +
      ids.map(function (id) { const x = p.parCreateur[id]; return '\n  • ' + x.nom + ' (' + id + ') : ' + String(x.jours).replace('.', ',') + ' journée(s) — ' + x.detail.join(' ; '); }).join('') +
      (p.nonRattachees.length ? '\n  ⚠️ Permanences sans créateur retrouvé (' + p.nonRattachees.length + ') :\n    - ' + p.nonRattachees.join('\n    - ') : ''));
  });
}

/*************************************************************
 *  Bénévoles repérés dans l'agenda — à lancer depuis l'éditeur
 *   apercuBenevolesAgenda()    : liste ce qui serait fait, n'écrit RIEN
 *   appliquerBenevolesAgenda() : coche « bénévole » sur les fiches retrouvées
 *  Fenêtre : 6 mois en arrière, 2 mois en avant. Une personne invitée est
 *  retrouvée par son e-mail, sinon par son nom (mêmes mots que le nom légal,
 *  dans n'importe quel ordre : « Maïa Blanchot » = « Blanchot Maïa », à vérifier).
 *  On ne décoche jamais personne : les fiches cochées absentes de l'agenda sont signalées.
 *************************************************************/

function apercuBenevolesAgenda() { _benevolesAgenda(true); }
function appliquerBenevolesAgenda() { _benevolesAgenda(false); }

/** « Maïa Blanchot » et « Blanchot Maïa » → même clé (mots triés, sans accents ni casse). */
function _cleMots(s) {
  return _cleNom(s).normalize('NFD').replace(/[̀-ͯ]/g, '').split(/[^a-z0-9]+/).filter(String).sort().join(' ');
}

function _benevolesAgenda(apercu) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), src = _agendasPermanences();
  if (!src.dedie) throw new Error("Pas d'agenda dont le nom contient « bénévole » parmi : " + CalendarApp.getAllCalendars().map(function (c) { return c.getName(); }).join(', '));
  const auj = new Date(), debut = new Date(auj.getFullYear(), auj.getMonth() - 6, 1), fin = new Date(auj.getFullYear(), auj.getMonth() + 3, 1);

  // Invités de l'agenda : e-mail → { nom, nb permanences, dernière date }
  const invites = {};
  src.agendas[0].getEvents(debut, fin).forEach(function (e) {
    e.getGuestList().forEach(function (g) {
      const m = _norm(g.getEmail());
      if (!m || m === EMAIL_SHOP) return;
      const x = invites[m] = invites[m] || { email: m, nom: '', nb: 0, derniere: null };
      if (g.getName() && g.getName() !== g.getEmail()) x.nom = g.getName();
      x.nb++;
      if (!x.derniere || e.getStartTime() > x.derniere) x.derniere = e.getStartTime();
    });
  });

  const tC = _tableau(ss, SHEET_CREATEURS), parEmail = {}, parMots = {};
  tC.lignes.forEach(function (r, i) {
    if (_horsBenevolat(_val(tC, r, 'email'), _val(tC, r, 'categorie'))) return;
    if (_val(tC, r, 'email')) parEmail[_norm(_val(tC, r, 'email'))] = i;
    [_val(tC, r, 'nom_legal'), _val(tC, r, 'nom')].forEach(function (n) { const k = _cleMots(n); if (k.indexOf(' ') > 0) (parMots[k] = parMots[k] || []).push(i); });
  });

  const coches = [], deja = [], parNom = [], introuvables = [], retrouves = {};
  Object.keys(invites).forEach(function (m) {
    const x = invites[m];
    let i = parEmail[m], comment = 'e-mail';
    if (i == null && x.nom) { const c = parMots[_cleMots(x.nom)]; if (c && c.length === 1) { i = c[0]; comment = 'nom (à vérifier)'; } }
    const resume = (x.nom ? x.nom + ' <' + m + '>' : m) + ' · ' + x.nb + ' permanence(s), dernière le ' + _dateFr(x.derniere);
    if (i == null) { introuvables.push(resume); return; }
    const r = tC.lignes[i], nom = _nomPropre(_val(tC, r, 'nom')) + ' (' + _val(tC, r, 'id_createur') + ')';
    retrouves[i] = true;
    if (comment !== 'e-mail') parNom.push(resume + ' → ' + nom);
    if (_val(tC, r, 'benevole') === true) { deja.push(nom); return; }
    coches.push(nom + ' ← ' + resume + ' [' + comment + ']');
    if (!apercu) {
      r[tC.M['benevole']] = true;
      if (tC.M['modifie_le'] != null) r[tC.M['modifie_le']] = new Date();
      _ecrireLigne(tC, i, r);
    }
  });
  const absents = tC.lignes.filter(function (r, i) { return _val(tC, r, 'benevole') === true && !retrouves[i]; })
    .map(function (r) { return _nomPropre(_val(tC, r, 'nom')) + ' (' + _val(tC, r, 'id_createur') + ')'; });

  const bilan = [
    Object.keys(invites).length + ' personne(s) invitée(s) dans « ' + src.agendas[0].getName() + ' » du ' + _dateFr(debut) + ' au ' + _dateFr(new Date(fin - 864e5)),
    coches.length + ' fiche(s) ' + (apercu ? 'à cocher' : 'cochée(s)') + ' « bénévole »' + (coches.length ? ' :\n    - ' + coches.join('\n    - ') : ''),
    deja.length + ' déjà cochée(s)' + (deja.length ? ' : ' + deja.join(', ') : ''),
    parNom.length ? 'Retrouvées par le NOM (e-mail différent de la fiche, à vérifier) :\n    - ' + parNom.join('\n    - ') : '',
    introuvables.length ? 'Invités sans fiche créateur retrouvée (gérantes, bénévoles non exposants…) :\n    - ' + introuvables.join('\n    - ') : '',
    absents.length ? 'Cochées « bénévole » mais absentes de l\'agenda (rien n\'est décoché) : ' + absents.join(', ') : ''
  ].filter(String);
  if (!apercu && coches.length) _journaliser('benevoles_agenda', coches.length + ' fiche(s) cochée(s) bénévole d\'après l\'agenda : ' + coches.map(function (c) { return c.split(' ← ')[0]; }).join(', '), Session.getEffectiveUser().getEmail());
  Logger.log((apercu ? "APERÇU — rien n'a été écrit.\n" : '✅ Bénévoles mis à jour.\n') + bilan.map(function (l) { return '• ' + l; }).join('\n'));
}
