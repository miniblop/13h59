/*************************************************************
 *  CONTRATS À DURÉE FIXE (grand stand : 4 mois)
 *  stands.duree_contrat_mois : durée d'un contrat (vide = durée indéterminée).
 *  emplacements.echeance : dernier jour du contrat en cours.
 *  Pas de renouvellement automatique (réponse de Mo) : à l'échéance, la
 *  gestion renouvelle pour une nouvelle durée ou clôture sans préavis ; le
 *  Collectif peut ne pas renouveler un créateur qui ne vend pas (6 mois
 *  sans vente). Partir avant l'échéance reste soumis au préavis d'un mois.
 *************************************************************/

/** Ajoute une colonne à un onglet si elle manque ; `init(t)` remplit les lignes existantes. */
function _assurerColonne(ss, nom, colonne, init) {
  let t = _tableau(ss, nom);
  if (t.M[colonne] != null) return t;
  _assurerTaille(t.sh, Math.max(1, t.sh.getLastRow()), t.nbCol + 1);
  t.sh.getRange(1, t.nbCol + 1).setValue(colonne).setFontWeight('bold');
  SpreadsheetApp.flush();
  t = _tableau(ss, nom);
  if (init) init(t);
  return t;
}
/** Colonnes des contrats à durée fixe ; à la création, le grand stand reçoit 4 mois (réponse de Mo). */
function _colonnesContrat(ss) {
  _assurerColonne(ss, SHEET_STANDS, 'duree_contrat_mois', function (t) {
    t.lignes.forEach(function (r, i) { if (String(_val(t, r, 'code')) === 'grand') { r[t.M['duree_contrat_mois']] = 4; _ecrireLigne(t, i, r); } });
  });
  _assurerColonne(ss, SHEET_EMPLACEMENTS, 'echeance');
}
function _dureeContrat(ss, code) {
  const s = _lireTable(_onglet(ss, SHEET_STANDS)).filter(function (x) { return String(x['code']) === String(code); })[0];
  return s ? Number(s['duree_contrat_mois']) || 0 : 0;
}
/** Dernier jour d'un contrat de `mois` mois commençant le `debut` (1er mai + 4 mois → 31 août). */
function _finContrat(debut, mois) { return new Date(debut.getFullYear(), debut.getMonth() + mois, debut.getDate() - 1); }
/** Échéance d'un nouvel emplacement ('' si le stand est à durée indéterminée). */
function _echeanceNouvelle(ss, code, debut) { const n = _dureeContrat(ss, code); return n ? _finContrat(debut, n) : ''; }

function _gContratRenouveler(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _colonnesContrat(ss);
  const tE = _tableau(ss, SHEET_EMPLACEMENTS), i = _ligneEmplacement(tE, body.idEmplacement), r = tE.lignes[i];
  const n = _dureeContrat(ss, _val(tE, r, 'code_stand'));
  if (!n) throw new Error('Ce stand est à durée indéterminée : rien à renouveler.');
  if (_val(tE, r, 'fin') instanceof Date) throw new Error('Une fin est déjà enregistrée pour cet emplacement : annule-la avant de renouveler.');
  const e = _val(tE, r, 'echeance');
  if (!(e instanceof Date)) throw new Error("Indique d'abord l'échéance du contrat en cours.");
  const nouvelle = _finContrat(new Date(e.getFullYear(), e.getMonth(), e.getDate() + 1), n);
  r[tE.M['echeance']] = nouvelle;
  _ecrireLigne(tE, i, r);
  _journaliser('contrat_renouvele', body.idEmplacement + ' (' + _val(tE, r, 'id_createur') + ') : ' + n + ' mois, jusqu\'au ' + _iso(nouvelle));
  return { ok: true, echeance: _iso(nouvelle) };
}

function _gContratEcheance(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _colonnesContrat(ss);
  const tE = _tableau(ss, SHEET_EMPLACEMENTS), i = _ligneEmplacement(tE, body.idEmplacement), r = tE.lignes[i];
  const d = _dateIso(body.echeance), debut = _val(tE, r, 'debut');
  if (debut instanceof Date && d < debut) throw new Error("L'échéance est avant l'arrivée en boutique.");
  r[tE.M['echeance']] = d;
  _ecrireLigne(tE, i, r);
  _journaliser('contrat_echeance', body.idEmplacement + ' (' + _val(tE, r, 'id_createur') + ') : échéance le ' + body.echeance);
  return { ok: true, echeance: _iso(d) };
}

/*************************************************************
 *  RENOUVELLEMENT PAR E-MAIL (contrats à durée fixe)
 *  Au début du dernier mois du contrat, le créateur reçoit un e-mail avec
 *  deux boutons (« Je renouvelle » / « Je ne renouvelle pas ») qui mènent à
 *  renouvellement.html : il y confirme d'un clic (une page, pas un lien
 *  direct, car certaines messageries ouvrent les liens toutes seules).
 *  Il a 7 jours pour répondre ; sans réponse, le contrat n'est pas renouvelé
 *  et se termine à l'échéance (l'e-mail vaut préavis).
 *  Registre : onglet `renouvellements`. Envoi et clôture : tacheQuotidienne(),
 *  déclenchée chaque matin (installerRenouvellements() une fois dans l'éditeur).
 *************************************************************/

const SHEET_RENOUVELLEMENTS = 'renouvellements';
const COLONNES_RENOUVELLEMENTS = ['id_renouvellement', 'id_emplacement', 'id_createur', 'echeance', 'jeton', 'envoye_le', 'date_limite', 'reponse',
  'repondu_le', 'statut', 'note', 'commentaire', 'remarque'];

/** Nom et e-mail d'une fiche créateur (null si absente). */
function _ficheCreateur(ss, id) {
  const c = _lireTable(_onglet(ss, SHEET_CREATEURS)).filter(function (x) { return String(x['id_createur']) === String(id); })[0];
  return c ? { id: String(id), nom: _nomPropre(c['nom']), email: _email(c['email']) } : null;
}
const DELAI_REPONSE_JOURS = 7;
const URL_SITE = 'https://miniblop.github.io/13h59/';

/** Premier jour du dernier mois du contrat (échéance 31 août → 1er août). */
function _debutDernierMois(e) { return new Date(e.getFullYear(), e.getMonth() - 1, e.getDate() + 1); }
function _jour(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }

/** Prolonge le contrat d'un emplacement d'une durée ; renvoie la nouvelle échéance. */
function _prolongerContrat(ss, tE, i) {
  const r = tE.lignes[i], n = _dureeContrat(ss, _val(tE, r, 'code_stand')), e = _val(tE, r, 'echeance');
  const nouvelle = _finContrat(new Date(e.getFullYear(), e.getMonth(), e.getDate() + 1), n);
  r[tE.M['echeance']] = nouvelle;
  _ecrireLigne(tE, i, r);
  return nouvelle;
}
/** Clôture un emplacement à son échéance (fin de contrat, sans préavis). */
function _cloreAEcheance(tE, i) {
  const r = tE.lignes[i];
  r[tE.M['fin']] = _val(tE, r, 'echeance'); r[tE.M['motif_fin']] = 'fin_contrat'; r[tE.M['preavis_recu_le']] = '';
  _ecrireLigne(tE, i, r);
}

function _ongletRenouvellements(ss) {
  if (ss.getSheetByName(SHEET_RENOUVELLEMENTS)) return;
  _creerOngletSiAbsent(ss, SHEET_RENOUVELLEMENTS, COLONNES_RENOUVELLEMENTS);
  ss.getSheetByName(SHEET_RENOUVELLEMENTS).getRange(1, 5, 1000, 1).setNumberFormat('@');   // jeton en texte
}

/** Envoie (ou renvoie) la demande de renouvellement d'un emplacement. */
function _demanderRenouvellement(ss, idEmplacement, par) {
  _colonnesContrat(ss); _ongletRenouvellements(ss);
  _modeleSiAbsent(ss, 'renouvellement_demande'); _modeleSiAbsent(ss, 'renouvellement_fin');
  const tE = _tableau(ss, SHEET_EMPLACEMENTS), i = _ligneEmplacement(tE, idEmplacement), e = tE.lignes[i];
  const n = _dureeContrat(ss, _val(tE, e, 'code_stand')), ech = _val(tE, e, 'echeance');
  if (!n) throw new Error('Ce stand est à durée indéterminée : pas de renouvellement.');
  if (!(ech instanceof Date)) throw new Error("L'échéance du contrat n'est pas renseignée.");
  if (_val(tE, e, 'fin') instanceof Date) throw new Error('Une fin est déjà enregistrée pour cet emplacement.');
  const idC = String(_val(tE, e, 'id_createur')), c = _ficheCreateur(ss, idC);
  const email = c ? c.email : '';
  if (!email) throw new Error("Pas d'adresse e-mail valide sur la fiche du créateur.");
  const tR = _tableau(ss, SHEET_RENOUVELLEMENTS);
  const iR = tR.lignes.findIndex(function (r) { return String(_val(tR, r, 'id_emplacement')) === String(idEmplacement) && _val(tR, r, 'echeance') instanceof Date && _val(tR, r, 'echeance').getTime() === ech.getTime(); });
  let jeton, limite;
  if (iR >= 0) {
    const r = tR.lignes[iR];
    if (String(_val(tR, r, 'statut')) !== 'en_attente') throw new Error('Le renouvellement de cette échéance est déjà traité (' + _val(tR, r, 'statut') + ').');
    jeton = String(_val(tR, r, 'jeton')); limite = _val(tR, r, 'date_limite');
    r[tR.M['envoye_le']] = new Date(); _ecrireLigne(tR, iR, r);
  } else {
    const auj = _jour(new Date());
    limite = new Date(auj.getFullYear(), auj.getMonth(), auj.getDate() + DELAI_REPONSE_JOURS);
    if (limite > ech) limite = ech;
    jeton = Utilities.getUuid().replace(/-/g, '');
    _ajouterLigne(tR, { id_renouvellement: _prochainId(tR, 'id_renouvellement', 'R'), id_emplacement: idEmplacement, id_createur: idC, echeance: ech, jeton: jeton,
      envoye_le: new Date(), date_limite: limite, reponse: '', repondu_le: '', statut: 'en_attente', remarque: '' });
  }
  const lien = URL_SITE + 'renouvellement.html?j=' + jeton;
  const envoye = envoyerModele(ss, 'renouvellement_demande', email, { marque: c.nom, echeance: _dateFr(ech), date_limite: _dateFr(limite), duree: n },
    [{ libelle: 'Je renouvelle', url: lien + '&c=oui', principal: true }, { libelle: 'Je ne renouvelle pas', url: lien + '&c=non' }]);
  _journaliser('renouvellement_demande', idEmplacement + ' (' + idC + ' ' + c.nom + ') : échéance ' + _dateFr(ech) + ', réponse avant le ' + _dateFr(limite) + (envoye ? '' : ' · modèle d\'e-mail désactivé : rien envoyé'), par);
  return { ok: true, envoye: envoye, dateLimite: _iso(limite) };
}

/** Chaque matin : envoie les demandes du début de dernier mois, clôt les contrats restés sans réponse. */
function tacheQuotidienne() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  const bilan = [];
  try {
    _colonnesContrat(ss); _ongletRenouvellements(ss);
    const auj = _jour(new Date());
    const tE = _tableau(ss, SHEET_EMPLACEMENTS), tR = _tableau(ss, SHEET_RENOUVELLEMENTS);
    const deja = {};
    tR.lignes.forEach(function (r) { const e = _val(tR, r, 'echeance'); if (e instanceof Date) deja[_val(tR, r, 'id_emplacement') + '|' + e.getTime()] = true; });
    tE.lignes.forEach(function (e) {
      const ech = _val(tE, e, 'echeance'), id = String(_val(tE, e, 'id_emplacement'));
      if (!(ech instanceof Date) || _val(tE, e, 'fin') instanceof Date || !_dureeContrat(ss, _val(tE, e, 'code_stand'))) return;
      if (auj < _debutDernierMois(ech) || auj > ech || deja[id + '|' + ech.getTime()]) return;
      try { _demanderRenouvellement(ss, id, 'automatique'); bilan.push('demande ' + id); }
      catch (err) { _journaliser('renouvellement_echec', id + ' : ' + (err && err.message ? err.message : err), 'automatique'); }
    });
    // sans réponse à la date limite : contrat non renouvelé
    const tR2 = _tableau(ss, SHEET_RENOUVELLEMENTS), tE2 = _tableau(ss, SHEET_EMPLACEMENTS);
    tR2.lignes.forEach(function (r, k) {
      if (String(_val(tR2, r, 'statut')) !== 'en_attente') return;
      const lim = _val(tR2, r, 'date_limite');
      if (!(lim instanceof Date) || auj <= _jour(lim)) return;
      const i = tE2.lignes.findIndex(function (e) { return String(_val(tE2, e, 'id_emplacement')) === String(_val(tR2, r, 'id_emplacement')); });
      const e = i >= 0 ? tE2.lignes[i] : null, ech = _val(tR2, r, 'echeance');
      const inchange = e && !(_val(tE2, e, 'fin') instanceof Date) && _val(tE2, e, 'echeance') instanceof Date && _val(tE2, e, 'echeance').getTime() === ech.getTime();
      if (inchange) {
        _cloreAEcheance(tE2, i);
        const c = _ficheCreateur(ss, _val(tR2, r, 'id_createur'));
        try { if (c && c.email) envoyerModele(ss, 'renouvellement_fin', c.email, { marque: c.nom, echeance: _dateFr(ech), date_limite: _dateFr(lim), duree: _dureeContrat(ss, _val(tE2, e, 'code_stand')) }); }
        catch (err) { _journaliser('email_echec', 'fin de contrat ' + _val(tR2, r, 'id_emplacement') + ' : ' + (err && err.message ? err.message : err), 'automatique'); }
        _journaliser('renouvellement_expire', _val(tR2, r, 'id_emplacement') + ' (' + _val(tR2, r, 'id_createur') + ') : sans réponse, fin de contrat le ' + _dateFr(ech), 'automatique');
        bilan.push('fin ' + _val(tR2, r, 'id_emplacement'));
      }
      r[tR2.M['statut']] = inchange ? 'expire' : 'annule';
      if (!inchange) r[tR2.M['remarque']] = 'contrat modifié par la gestion avant la date limite';
      _ecrireLigne(tR2, k, r);
    });
  } finally {
    lock.releaseLock();
  }
  return bilan;
}

/* ---------- Page publique renouvellement.html (le jeton fait office d'accès) ---------- */

function _renouvellementLigne(ss, jeton) {
  const j = String(jeton || '');
  if (!/^[0-9a-f]{32}$/.test(j) || !ss.getSheetByName(SHEET_RENOUVELLEMENTS)) return null;
  const t = _tableau(ss, SHEET_RENOUVELLEMENTS), i = t.lignes.findIndex(function (r) { return String(_val(t, r, 'jeton')) === j; });
  return i < 0 ? null : { t: t, i: i, r: t.lignes[i] };
}
function _renouvellementInfos(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), x = _renouvellementLigne(ss, body.jeton);
  if (!x) return { ok: false, message: 'Lien invalide ou expiré. Écris-nous à ' + EMAIL_SHOP + '.' };
  const t = x.t, r = x.r, c = _ficheCreateur(ss, _val(t, r, 'id_createur'));
  const tE = _tableau(ss, SHEET_EMPLACEMENTS), e = tE.lignes.filter(function (l) { return String(_val(tE, l, 'id_emplacement')) === String(_val(t, r, 'id_emplacement')); })[0];
  const n = e ? _dureeContrat(ss, _val(tE, e, 'code_stand')) : 0;
  return { ok: true, marque: c ? c.nom : '', stand: e ? _libelleStand(ss, _val(tE, e, 'code_stand')) : '', duree: n, echeance: _iso(_val(t, r, 'echeance')),
    dateLimite: _iso(_val(t, r, 'date_limite')), statut: String(_val(t, r, 'statut')), reponse: String(_val(t, r, 'reponse') || ''),
    nouvelleEcheance: e && String(_val(t, r, 'statut')) === 'renouvele' ? _iso(_val(tE, e, 'echeance')) : '', avisDonne: _val(t, r, 'note') !== '' || !!_val(t, r, 'commentaire') };
}
function _renouvellementRepondre(body) {
  const choix = body.choix === 'oui' ? 'oui' : body.choix === 'non' ? 'non' : '';
  if (!choix) return { ok: false, message: 'Réponse invalide.' };
  const ss = SpreadsheetApp.getActiveSpreadsheet(), lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, message: 'Le serveur est occupé : réessaie dans une minute.' };
  try {
    const x = _renouvellementLigne(ss, body.jeton);
    if (!x) return { ok: false, message: 'Lien invalide ou expiré. Écris-nous à ' + EMAIL_SHOP + '.' };
    const t = x.t, r = x.r, statut = String(_val(t, r, 'statut'));
    if (statut !== 'en_attente') return { ok: false, deja: true, message: statut === 'renouvele' ? 'Ton contrat est déjà renouvelé.' : statut === 'non_renouvele' ? 'Tu as déjà indiqué ne pas renouveler.' : 'Cette demande est close : écris-nous à ' + EMAIL_SHOP + '.' };
    if (_jour(new Date()) > _jour(_val(t, r, 'date_limite'))) return { ok: false, message: 'Le délai pour répondre est dépassé : écris-nous vite à ' + EMAIL_SHOP + '.' };
    const tE = _tableau(ss, SHEET_EMPLACEMENTS), i = tE.lignes.findIndex(function (l) { return String(_val(tE, l, 'id_emplacement')) === String(_val(t, r, 'id_emplacement')); });
    const e = i >= 0 ? tE.lignes[i] : null, ech = _val(t, r, 'echeance');
    if (!e || _val(tE, e, 'fin') instanceof Date || !(_val(tE, e, 'echeance') instanceof Date) || _val(tE, e, 'echeance').getTime() !== ech.getTime())
      return { ok: false, message: "L'équipe a déjà traité ton contrat : écris-nous à " + EMAIL_SHOP + ' si besoin.' };
    let nouvelle = null;
    if (choix === 'oui') nouvelle = _prolongerContrat(ss, tE, i); else _cloreAEcheance(tE, i);
    r[t.M['reponse']] = choix; r[t.M['repondu_le']] = new Date(); r[t.M['statut']] = choix === 'oui' ? 'renouvele' : 'non_renouvele';
    _ecrireLigne(t, x.i, r);
    const idC = String(_val(t, r, 'id_createur')), c = _ficheCreateur(ss, idC);
    _journaliser(choix === 'oui' ? 'renouvellement_accepte' : 'renouvellement_refuse', _val(t, r, 'id_emplacement') + ' (' + idC + ' ' + (c ? c.nom : '') + ') : ' +
      (choix === 'oui' ? 'renouvelé jusqu\'au ' + _dateFr(nouvelle) : 'ne renouvelle pas, fin le ' + _dateFr(ech)), 'créateur (e-mail)');
    return { ok: true, choix: choix, echeance: _iso(ech), nouvelleEcheance: nouvelle ? _iso(nouvelle) : '' };
  } finally {
    lock.releaseLock();
  }
}

/** Avis facultatif d'un créateur qui ne renouvelle pas : gardé au registre et envoyé au shop. */
function _renouvellementAvis(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), x = _renouvellementLigne(ss, body.jeton);
  if (!x) return { ok: false, message: 'Lien invalide ou expiré.' };
  const t = x.t, r = x.r;
  if (String(_val(t, r, 'statut')) !== 'non_renouvele') return { ok: false, message: "L'avis se donne après avoir indiqué ne pas renouveler." };
  if (_val(t, r, 'note') !== '' || _val(t, r, 'commentaire')) return { ok: true, deja: true };
  const note = Number(body.note), commentaire = _textePublic(body.commentaire, 2000);
  if (!(note >= 1 && note <= 5 && note === Math.round(note)) && !commentaire) return { ok: false, message: 'Donne une note ou écris quelques mots.' };
  r[t.M['note']] = note >= 1 && note <= 5 ? Math.round(note) : ''; r[t.M['commentaire']] = commentaire;
  _ecrireLigne(t, x.i, r);
  const c = _ficheCreateur(ss, _val(t, r, 'id_createur')), nom = c ? c.nom : String(_val(t, r, 'id_createur'));
  const etoiles = r[t.M['note']] ? '★★★★★'.slice(0, r[t.M['note']]) + '☆☆☆☆☆'.slice(0, 5 - r[t.M['note']]) + ' (' + r[t.M['note']] + '/5)' : 'pas de note';
  const texte = nom + ' ne renouvelle pas son contrat (fin le ' + _dateFr(_val(t, r, 'echeance')) + ') et nous laisse un avis.\n\nNote : ' + etoiles + '\n\n' + (commentaire ? 'Commentaire :\n' + commentaire : 'Pas de commentaire.');
  try { envoyerEmailShop({ to: EMAIL_SHOP, subject: 'Avis de départ · ' + nom, texte: texte, html: _htmlShop(texte) }); }
  catch (e) { _journaliser('email_echec', 'avis de départ ' + nom + ' : ' + (e && e.message ? e.message : e), 'créateur (e-mail)'); }
  _journaliser('avis_depart', nom + ' : ' + etoiles, 'créateur (e-mail)');
  return { ok: true };
}

/* ---------- Gestion et installation ---------- */

function _gRenouvellementEnvoyer(body) { return _demanderRenouvellement(SpreadsheetApp.getActiveSpreadsheet(), String(body.idEmplacement || ''), _auteurJournal); }

/** Dernière demande de renouvellement par emplacement, pour la fiche. */
function _renouvellementsParEmplacement(ss) {
  const out = {};
  if (!ss.getSheetByName(SHEET_RENOUVELLEMENTS)) return out;
  _lireTable(_onglet(ss, SHEET_RENOUVELLEMENTS)).forEach(function (r) {
    const id = String(r['id_emplacement'] || ''), env = r['envoye_le'];
    if (!id) return;
    const o = { echeance: _iso(r['echeance']), envoyeLe: _iso(env), dateLimite: _iso(r['date_limite']), statut: String(r['statut'] || ''), reponse: String(r['reponse'] || ''), reponduLe: _iso(r['repondu_le']),
      note: r['note'] === '' || r['note'] == null ? '' : Number(r['note']), commentaire: String(r['commentaire'] || '') };
    if (!out[id] || o.echeance >= out[id].echeance) out[id] = o;
  });
  return out;
}

/** À lancer une fois depuis l'éditeur : liste ce que la tâche du matin ferait aujourd'hui, sans rien envoyer. */
function apercuRenouvellements() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); _colonnesContrat(ss);
  const auj = _jour(new Date()), tE = _tableau(ss, SHEET_EMPLACEMENTS), noms = _createurs(ss).parId, lignes = [];
  tE.lignes.forEach(function (e) {
    const ech = _val(tE, e, 'echeance'), n = _dureeContrat(ss, _val(tE, e, 'code_stand'));
    if (!n || _val(tE, e, 'fin') instanceof Date) return;
    const c = noms[String(_val(tE, e, 'id_createur'))], nom = c ? c.nom : _val(tE, e, 'id_createur');   // noms : _createurs(ss).parId
    if (!(ech instanceof Date)) { lignes.push('⚠️ ' + nom + ' : échéance à renseigner sur sa fiche'); return; }
    const d = _debutDernierMois(ech);
    lignes.push((auj >= d && auj <= ech ? '📧 e-mail envoyé dès demain matin : ' : '· ') + nom + ' : échéance ' + _dateFr(ech) + ', demande le ' + _dateFr(d));
  });
  Logger.log(lignes.length ? lignes.join('\n') : 'Aucun contrat à durée fixe en cours.');
}

/** À lancer une fois depuis l'éditeur : autorise et programme la tâche de chaque matin (8 h). */
function installerRenouvellements() {
  ScriptApp.getProjectTriggers().forEach(function (tr) { if (tr.getHandlerFunction() === 'tacheQuotidienne') ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('tacheQuotidienne').timeBased().everyDays(1).atHour(8).create();
  Logger.log('✅ Tâche programmée chaque matin vers 8 h : demandes de renouvellement et fins de contrat sans réponse.');
  apercuRenouvellements();
}
