/*************************************************************
 *  À ENCAISSER — loyers (créateurs → Collectif)
 *  Le loyer du mois M se paie par virement du 20 au 25 du mois M−1.
 *  Attendu : le loyer de la facture du mois si elle est générée (montant figé),
 *  sinon le loyer calculé. Reçu : les virements pointés, depuis l'extrait de
 *  compte (PDF du CIC, lu dans le navigateur) ou à la main.
 *   - paiements_recus : un virement pointé (jamais effacé : statut « annule »)
 *   - payeurs         : « tel nom de payeur = tel créateur », appris à chaque
 *                       pointage depuis la banque (ex. « BAILLET SARA » paie
 *                       le loyer de Jackline)
 *   - relances        : les e-mails de rappel envoyés
 *************************************************************/

const SHEET_PAIEMENTS_RECUS = 'paiements_recus';
const SHEET_PAYEURS = 'payeurs';
const SHEET_RELANCES = 'relances';
const COLONNES_PAIEMENTS_RECUS = ['id_paiement', 'mois_loyer', 'id_createur', 'recu_le', 'montant', 'source', 'libelle', 'ref_releve', 'statut', 'pointe_par', 'pointe_le', 'remarque'];
const COLONNES_PAYEURS = ['payeur', 'id_createur', 'appris_le'];
const COLONNES_RELANCES = ['envoyee_le', 'id_createur', 'mois_loyer', 'email', 'reste_du', 'envoyee_par'];
const JOUR_ECHEANCE = 25;

function _ongletsEncaissement(ss) {
  _creerOngletSiAbsent(ss, SHEET_PAIEMENTS_RECUS, COLONNES_PAIEMENTS_RECUS);
  _creerOngletSiAbsent(ss, SHEET_PAYEURS, COLONNES_PAYEURS);
  _creerOngletSiAbsent(ss, SHEET_RELANCES, COLONNES_RELANCES);
  // textes bruts : « 2026-10 » et les identifiants ne doivent pas devenir des dates
  const p = ss.getSheetByName(SHEET_PAIEMENTS_RECUS);
  p.getRange(1, 1, p.getMaxRows(), 3).setNumberFormat('@');
  p.getRange(1, 7, p.getMaxRows(), 2).setNumberFormat('@');
  ss.getSheetByName(SHEET_RELANCES).getRange(1, 3, ss.getSheetByName(SHEET_RELANCES).getMaxRows(), 1).setNumberFormat('@');
}
/** Échéance du loyer du mois M : le 25 du mois M−1. */
function _echeance(mois) { const b = _moisBornes(mois); return new Date(b.debut.getFullYear(), b.debut.getMonth() - 1, JOUR_ECHEANCE); }
function _lireSi(ss, nom) { return ss.getSheetByName(nom) ? _lireTable(_onglet(ss, nom)) : []; }

/** Loyers attendus du mois : { idCreateur: {loyer, figee, benevole, nom, email} }. */
function _loyersAttendus(ss, mois) {
  const ctx = _contexteFactures(ss, mois), out = {};
  _aFacturer(ctx).forEach(function (r) {
    const id = String(_val(ctx.tC, r, 'id_createur')), f = ctx.factures[id];
    const c = _calculFacture(ctx, r, f ? f['perms_faites'] : 0, f ? f['perms_requises'] : '');
    const loyer = f ? (Number(f['loyer']) || 0) : c.loyer;
    if (loyer > 0) out[id] = { loyer: _round2(loyer), figee: !!f, numero: f ? String(f['numero']) : '', benevole: c.benevole, nom: c.nom, email: c.email };
  });
  return out;
}

function _gEncaissements(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), mois = String(body.mois || '');
  const attendus = _loyersAttendus(ss, mois);
  const pay = {}, refs = [];
  _lireSi(ss, SHEET_PAIEMENTS_RECUS).forEach(function (p) {
    if (p['ref_releve']) refs.push(String(p['ref_releve']));
    if (String(p['statut']) === 'annule' || _moisTexte(p['mois_loyer']) !== mois) return;
    const id = String(p['id_createur']);
    (pay[id] = pay[id] || []).push({ id: String(p['id_paiement']), date: _iso(p['recu_le']), montant: Number(p['montant']) || 0, source: String(p['source'] || ''), libelle: String(p['libelle'] || '') });
  });
  const relances = {};
  _lireSi(ss, SHEET_RELANCES).forEach(function (r) { if (_moisTexte(r['mois_loyer']) === mois) relances[String(r['id_createur'])] = _iso(r['envoyee_le']); });
  const ech = _echeance(mois), auj = _aujourdhui();
  const ids = Object.keys(attendus).concat(Object.keys(pay).filter(function (id) { return !attendus[id]; }));
  const tC = _tableau(ss, SHEET_CREATEURS), fiches = {};
  tC.lignes.forEach(function (r) { fiches[String(_val(tC, r, 'id_createur'))] = r; });
  const lignes = ids.map(function (id) {
    const a = attendus[id] || { loyer: 0, figee: false, numero: '', benevole: false, nom: fiches[id] ? _nomPropre(_val(tC, fiches[id], 'nom')) : id, email: fiches[id] ? String(_val(tC, fiches[id], 'email') || '') : '' };
    const p = pay[id] || [], recu = _round2(p.reduce(function (s, x) { return s + x.montant; }, 0)), reste = _round2(a.loyer - recu);
    const statut = reste < -0.009 ? 'trop' : reste <= 0.009 ? 'recu' : recu > 0 ? 'partiel' : auj > ech ? 'retard' : 'attente';
    return { idCreateur: id, nom: a.nom, email: a.email, loyer: a.loyer, figee: a.figee, numero: a.numero, benevole: a.benevole, recu: recu, reste: Math.max(0, reste), paiements: p, statut: statut, relanceLe: relances[id] || '' };
  }).sort(function (a, b) { return a.nom.localeCompare(b.nom, 'fr', { sensitivity: 'base' }); });

  // Pour rapprocher l'extrait de compte dans le navigateur : fiches (hors équipe) et payeurs appris.
  const createurs = tC.lignes.filter(function (r) { return _val(tC, r, 'id_createur') && String(_val(tC, r, 'categorie')) !== 'shop'; }).map(function (r) {
    return { id: String(_val(tC, r, 'id_createur')), nom: _nomPropre(_val(tC, r, 'nom')), nomLegal: String(_val(tC, r, 'nom_legal') || ''), actif: _norm(_val(tC, r, 'statut')) === 'actif' };
  });
  const payeurs = {};
  _lireSi(ss, SHEET_PAYEURS).forEach(function (p) { if (p['payeur']) payeurs[String(p['payeur'])] = String(p['id_createur']); });
  const b = _moisBornes(mois);
  return { ok: true, mois: mois, libelle: b.libelle, echeance: _iso(ech), aujourdhui: _iso(auj), lignes: lignes, createurs: createurs, payeurs: payeurs, refs: refs };
}

/** Pointe des virements reçus (lot) ; apprend « payeur → créateur » pour les lignes venant de la banque. */
function _gPaiementsPointer(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _ongletsEncaissement(ss);
  const tP = _tableau(ss, SHEET_PAIEMENTS_RECUS), tY = _tableau(ss, SHEET_PAYEURS);
  const dejaRefs = {};
  tP.lignes.forEach(function (r) { if (_val(tP, r, 'ref_releve')) dejaRefs[String(_val(tP, r, 'ref_releve'))] = true; });
  const connus = {};
  tY.lignes.forEach(function (r, i) { connus[String(_val(tY, r, 'payeur'))] = i; });
  const ids = {};
  _lireTable(_onglet(ss, SHEET_CREATEURS)).forEach(function (c) { ids[String(c['id_createur'])] = _nomPropre(c['nom']); });
  let n = Math.max(_maxId(tP, 'id_paiement', 'P'), _dernierId('DERNIER_ID_P'));
  const nouvelles = [], appris = [], resume = [];
  (body.lignes || []).forEach(function (l) {
    const id = String(l.idCreateur || ''), mois = String(l.mois || ''), montant = _round2(Number(String(l.montant).replace(',', '.')));
    if (!ids[id]) throw new Error('Créateur inconnu : « ' + id + ' ».');
    if (!/^\d{4}-\d{2}$/.test(mois)) throw new Error('Mois du loyer invalide : « ' + mois + ' ».');
    if (!(montant > 0)) throw new Error('Montant invalide pour ' + ids[id] + '.');
    const ref = String(l.ref || '');
    if (ref && dejaRefs[ref]) return;                       // déjà pointé lors d'un import précédent
    if (ref) dejaRefs[ref] = true;
    n++;
    nouvelles.push({ id_paiement: 'P' + String(n).padStart(4, '0'), mois_loyer: mois, id_createur: id, recu_le: l.date ? _dateIso(l.date) : _aujourdhui(),
      montant: montant, source: l.source === 'banque' ? 'banque' : 'manuel', libelle: _textePublic(l.libelle, 300), ref_releve: ref, statut: 'ok', pointe_par: _signataire(), pointe_le: new Date() });
    resume.push(ids[id] + ' ' + _eurFr(montant) + ' (' + mois + ')');
    const payeur = _cleMots(l.payeur || '');
    if (l.source === 'banque' && payeur && connus[payeur] == null) { connus[payeur] = -1; appris.push([payeur, id, new Date()]); }
  });
  if (nouvelles.length) {
    const sh = tP.sh, der = sh.getLastRow();
    _assurerTaille(sh, der + nouvelles.length, COLONNES_PAIEMENTS_RECUS.length);
    sh.getRange(der + 1, 1, nouvelles.length, COLONNES_PAIEMENTS_RECUS.length).setValues(nouvelles.map(function (o) {
      return COLONNES_PAIEMENTS_RECUS.map(function (k) { return o[k] == null ? '' : o[k]; });
    }));
  }
  if (appris.length) {
    const sh = tY.sh, der = sh.getLastRow();
    _assurerTaille(sh, der + appris.length, COLONNES_PAYEURS.length);
    sh.getRange(der + 1, 1, appris.length, COLONNES_PAYEURS.length).setValues(appris);
  }
  if (nouvelles.length) _journaliser('loyers_pointes', nouvelles.length + ' virement(s) : ' + resume.join(', '));
  return { ok: true, pointes: nouvelles.length, ignores: (body.lignes || []).length - nouvelles.length, appris: appris.length };
}

function _gPaiementAnnuler(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), t = _tableau(ss, SHEET_PAIEMENTS_RECUS), motif = _motif(body);
  const i = t.lignes.findIndex(function (r) { return String(_val(t, r, 'id_paiement')) === String(body.idPaiement); });
  if (i < 0) throw new Error('Paiement introuvable.');
  const r = t.lignes[i];
  r[t.M['statut']] = 'annule'; r[t.M['remarque']] = motif;
  _ecrireLigne(t, i, r);
  _journaliser('loyer_depointe', body.idPaiement + ' (' + _val(t, r, 'id_createur') + ', ' + _eurFr(_val(t, r, 'montant')) + ') : ' + motif);
  return { ok: true };
}

/** Relance par e-mail les créateurs dont le loyer du mois n'est pas (entièrement) reçu. */
function _gRelancer(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _ongletsEncaissement(ss);
  _modeleSiAbsent(ss, 'relance_loyer');
  const d = _gEncaissements({ mois: body.mois }), voulus = body.ids || [];
  const tR = _tableau(ss, SHEET_RELANCES), envoyes = [], erreurs = [];
  d.lignes.filter(function (l) { return voulus.indexOf(l.idCreateur) !== -1 && l.reste > 0; }).forEach(function (l) {
    if (!l.email) { erreurs.push(l.nom + " : pas d'e-mail"); return; }
    try {
      if (!envoyerModele(ss, 'relance_loyer', l.email, { marque: l.nom, mois: d.libelle, montant: _eurFr(l.reste) })) throw new Error('modèle « relance_loyer » désactivé');
      _ajouterLigne(tR, { envoyee_le: new Date(), id_createur: l.idCreateur, mois_loyer: body.mois, email: l.email, reste_du: l.reste, envoyee_par: _signataire() });
      envoyes.push(l.nom);
    } catch (e) { erreurs.push(l.nom + ' : ' + (e && e.message ? e.message : e)); }
  });
  if (envoyes.length) _journaliser('loyers_relances', d.libelle + ' : ' + envoyes.join(', '));
  return { ok: true, envoyes: envoyes.length, erreurs: erreurs };
}
