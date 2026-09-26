/*************************************************************
 *  FACTURES — Gestion ▸ Facturation
 *  Facture du mois M = loyer du mois M + commission sur les ventes du mois M−1.
 *   - loyer : chaque emplacement présent pendant le mois, au prorata des jours
 *     d'ouverture (mardi → samedi) ; une arrivée le 3 quand le 1er et le 2 sont
 *     fermés paie donc le mois entier ;
 *   - bénévole : aucune commission, loyer réduit selon les permanences faites
 *     (journées pour un loyer offert = stands.perms_loyer_gratuit) ;
 *   - commission : 10 % dès 100 €, 15 % dès 250 € de prime du mois (prime =
 *     prix client − frais), même calcul que les tableaux de bord ;
 *   - l'équipe (catégorie « shop ») n'est pas facturée ; l'adhésion se paie
 *     à part (HelloAsso) et n'apparaît pas sur la facture.
 *  À la génération, les montants sont FIGÉS dans `factures` et `lignes_facture` ;
 *  le PDF est rangé dans Drive (13h59 / Factures / AAAA-MM).
 *************************************************************/

const SHEET_FACTURES = 'factures';
const SHEET_LIGNES_FACTURE = 'lignes_facture';
const COLONNES_FACTURES = ['numero', 'mois', 'id_createur', 'emise_le', 'total', 'loyer', 'commission', 'taux_commission',
  'ventes_mois_precedent', 'frais_mois_precedent', 'prime_mois_precedent', 'net_a_verser', 'perms_faites', 'perms_requises',
  'statut', 'pdf_id', 'envoyee_le', 'envoyee_a', 'emise_par', 'remarque'];
const COLONNES_LIGNES_FACTURE = ['numero', 'ordre', 'ref', 'libelle', 'montant'];
const ASSO = {
  nom: 'COLLECTIF 13H59', adresse: '27 rue du Plat', ville: '59800 LILLE', pays: 'FRANCE', email: 'contact : 13h59shop@gmail.com',
  tva: 'FR61988937876', siret: '98893787600015', naf: '91.02Z'
};
/* Journées de permanence pour un loyer offert, tant qu'elles ne sont pas réglées dans Réglages ▸ Stands (PROVISOIRE). */
const PERMS_PAR_DEFAUT = { illu: 1, unique: 2, grand: 4 };
const MOIS_FR = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];

/* ---------- Outils ---------- */

function _moisBornes(mois) {
  const m = /^(\d{4})-(\d{2})$/.exec(String(mois || ''));
  if (!m) throw new Error('Mois invalide : « ' + mois + ' ».');
  const y = +m[1], mo = +m[2] - 1;
  return { debut: new Date(y, mo, 1), fin: new Date(y, mo + 1, 0), libelle: MOIS_FR[mo] + ' ' + y };
}
function _moisPrecedent(mois) {
  const b = _moisBornes(mois), d = new Date(b.debut.getFullYear(), b.debut.getMonth() - 1, 1);
  return d.getFullYear() + '-' + _deux(d.getMonth() + 1);
}
/** Jours d'ouverture (mardi → samedi) entre deux dates incluses. */
function _joursOuverts(a, b) {
  let n = 0;
  for (let d = new Date(a.getFullYear(), a.getMonth(), a.getDate()); d <= b; d.setDate(d.getDate() + 1)) {
    const j = d.getDay();
    if (j >= 2 && j <= 6) n++;
  }
  return n;
}
function _tauxCommission(prime, mois) {
  if (mois < '2026-01') return prime >= 250 ? 0.12 : prime >= 100 ? 0.08 : 0;
  return prime >= 250 ? 0.15 : prime >= 100 ? 0.10 : 0;
}
function _eurFr(n) {
  const v = _round2(Number(n) || 0), s = Math.abs(v).toFixed(2).split('.');
  return (v < 0 ? '−' : '') + s[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + ',' + s[1] + ' €';
}
/** « de septembre », « d'octobre », « d'août ». */
function _de(mot) { return (/^[aeiouyàâéèêîôû]/i.test(String(mot)) ? "d'" : 'de ') + mot; }
function _dateFr(d) { return _deux(d.getDate()) + '/' + _deux(d.getMonth() + 1) + '/' + d.getFullYear(); }
/* Sheets lit « 26-10-001 » ou « 2026-10 » comme des dates : ces colonnes sont en texte brut,
 * et la lecture reconstruit la valeur si une ancienne cellule a déjà été convertie. */
function _colonnesTexteFactures(ss) {
  const f = ss.getSheetByName(SHEET_FACTURES), l = ss.getSheetByName(SHEET_LIGNES_FACTURE);
  if (f) f.getRange(1, 1, f.getMaxRows(), 2).setNumberFormat('@');   // numero, mois
  if (l) l.getRange(1, 1, l.getMaxRows(), 1).setNumberFormat('@');   // numero
}
function _moisTexte(v) { return v instanceof Date ? _jourIso(v).slice(0, 7) : String(v == null ? '' : v); }
/** « 26-10-001 » lu comme le 26/10/2001 → « 26-10-001 ». */
function _numeroTexte(v) {
  return v instanceof Date ? _deux(v.getDate()) + '-' + _deux(v.getMonth() + 1) + '-' + String(v.getFullYear() % 1000).padStart(3, '0') : String(v == null ? '' : v);
}
function _creerOngletSiAbsent(ss, nom, colonnes) {
  if (ss.getSheetByName(nom)) return;
  const sh = ss.insertSheet(nom);
  sh.getRange(1, 1, 1, colonnes.length).setValues([colonnes]).setFontWeight('bold');
  sh.setFrozenRows(1);
}
/** Modèle d'e-mail « facture » ajouté à l'onglet emails s'il manque. */
function _modeleFactureSiAbsent(ss) {
  _creerOngletEmails(ss);
  const t = _tableau(ss, SHEET_EMAILS);
  if (t.lignes.some(function (r) { return String(_val(t, r, 'code')) === 'facture'; })) return;
  const d = EMAILS_PAR_DEFAUT.filter(function (l) { return l[0] === 'facture'; })[0];
  _ajouterLigne(t, { code: d[0], objet: d[1], texte: d[2], actif: d[3], utilise_pour: d[4] });
  t.sh.getRange(t.lignes.length + 1, t.M['actif'] + 1).insertCheckboxes().setValue(true);
}

/* ---------- Calcul (sans rien écrire) ---------- */

/** Tout ce qu'il faut pour calculer les factures d'un mois, en une lecture par onglet. */
function _contexteFactures(ss, mois) {
  const b = _moisBornes(mois), prec = _moisPrecedent(mois);
  const tC = _tableau(ss, SHEET_CREATEURS), tE = _tableau(ss, SHEET_EMPLACEMENTS);
  const stands = {};
  _lireTable(_onglet(ss, SHEET_STANDS)).forEach(function (s) {
    if (s['code']) stands[String(s['code'])] = { code: String(s['code']), libelle: String(s['libelle'] || s['code']), loyer: Number(s['loyer']) || 0,
      ref: String(s['ref_facture'] || '') || String(s['code']).toUpperCase(), perms: s['perms_loyer_gratuit'] === '' ? null : Number(s['perms_loyer_gratuit']) };
  });
  const ventes = {};
  _lireTable(_onglet(ss, SHEET_VENTES)).forEach(function (v) {
    if (!(v['date'] instanceof Date) || _jourIso(v['date']).slice(0, 7) !== prec) return;
    const id = String(v['id_createur'] || '');
    (ventes[id] = ventes[id] || []).push(v);
  });
  const factures = {};
  if (ss.getSheetByName(SHEET_FACTURES)) {
    _lireTable(_onglet(ss, SHEET_FACTURES)).forEach(function (f) {
      f['mois'] = _moisTexte(f['mois']); f['numero'] = _numeroTexte(f['numero']);
      if (f['mois'] === mois && String(f['statut']) !== 'annulee') factures[String(f['id_createur'])] = f;
    });
  }
  return { b: b, mois: mois, prec: prec, precBornes: _moisBornes(prec), tC: tC, tE: tE, stands: stands, ventes: ventes, factures: factures,
           ouvertsMois: _joursOuverts(b.debut, b.fin) };
}

/** Calcule la facture d'un créateur (lignes, totaux, relevé des ventes). permsFaites : nombre saisi pour un bénévole. */
function _calculFacture(ctx, r, permsFaites) {
  const tC = ctx.tC, tE = ctx.tE, b = ctx.b, id = String(_val(tC, r, 'id_createur'));
  const benevole = _val(tC, r, 'benevole') === true, lignes = [], alertes = [];
  // Loyers : chaque emplacement présent pendant le mois.
  let loyer = 0, permsRequises = null, standPrincipal = null;
  tE.lignes.forEach(function (e) {
    if (String(_val(tE, e, 'id_createur')) !== id) return;
    const d = _val(tE, e, 'debut'), f = _val(tE, e, 'fin');
    if (!(d instanceof Date) || d > b.fin || (f instanceof Date && f < b.debut)) return;
    const s = ctx.stands[String(_val(tE, e, 'code_stand'))];
    if (!s) { alertes.push('stand inconnu : ' + _val(tE, e, 'code_stand')); return; }
    const de = d > b.debut ? d : b.debut, a = (f instanceof Date && f < b.fin) ? f : b.fin;
    const jours = _joursOuverts(de, a), part = ctx.ouvertsMois ? jours / ctx.ouvertsMois : 1;
    const montant = _round2(s.loyer * Math.min(1, part));
    const prorata = jours < ctx.ouvertsMois;
    lignes.push({ ref: s.ref, libelle: 'Loyer · ' + s.libelle + (prorata ? ' · du ' + _dateFr(de) + ' au ' + _dateFr(a) + ' (' + jours + ' jours d\'ouverture sur ' + ctx.ouvertsMois + ')' : ''), montant: montant });
    loyer += montant;
    if (!standPrincipal || montant > standPrincipal.montant) standPrincipal = { code: s.code, montant: montant, perms: s.perms };
  });
  // Bénévole : loyer réduit selon les permanences faites.
  if (benevole && standPrincipal) {
    permsRequises = standPrincipal.perms != null ? standPrincipal.perms : (PERMS_PAR_DEFAUT[standPrincipal.code] != null ? PERMS_PAR_DEFAUT[standPrincipal.code] : null);
    if (permsRequises == null) alertes.push('bénévolat impossible sur ce stand');
    else {
      if (standPrincipal.perms == null) alertes.push('journées de permanence provisoires (' + permsRequises + ') : à régler dans Réglages ▸ Stands');
      const faites = Math.max(0, Number(permsFaites) || 0);
      const remise = _round2(Math.min(loyer, loyer * faites / permsRequises));
      if (remise > 0) lignes.push({ ref: 'BEN', libelle: 'Remise bénévolat · ' + String(faites).replace('.', ',') + ' journée(s) de permanence sur ' + permsRequises, montant: -remise });
      loyer = _round2(loyer - remise);
    }
  }
  // Ventes du mois précédent et commission.
  const vs = (ctx.ventes[id] || []).slice().sort(function (x, y) { return x['date'] - y['date'] || Number(x['id_vente']) - Number(y['id_vente']); });
  const somme = function (k) { return _round2(vs.reduce(function (s, v) { return s + (Number(v[k]) || 0); }, 0)); };
  const ventesCA = somme('prix_client'), frais = somme('frais'), prime = somme('prime');
  const taux = benevole ? 0 : _tauxCommission(prime, ctx.prec), commission = _round2(prime * taux);
  if (commission) lignes.push({ ref: 'COM' + Math.round(taux * 100), libelle: 'Commission ' + Math.round(taux * 100) + ' % · ventes ' + _de(ctx.precBornes.libelle) + ' (prime ' + _eurFr(prime) + ')', montant: commission });
  const releve = vs.map(function (v) {
    return { date: _jourIso(v['date']), reference: String(v['reference'] || ''), paiement: String(v['code_paiement'] || ''), remise: String(v['code_remise'] || ''),
      prix: Number(v['prix']) || 0, prixClient: Number(v['prix_client']) || 0, frais: Number(v['frais']) || 0, type: _norm(v['type_ligne']) || 'vente' };
  });
  return {
    idCreateur: id, nom: _nomPropre(_val(tC, r, 'nom')), nomLegal: String(_val(tC, r, 'nom_legal') || ''), email: String(_val(tC, r, 'email') || ''),
    adresse: String(_val(tC, r, 'adresse') || ''), codePostal: String(_val(tC, r, 'code_postal') || ''), ville: String(_val(tC, r, 'ville') || ''),
    iban: String(_val(tC, r, 'iban') || '') ? true : false, benevole: benevole, permsFaites: benevole ? Math.max(0, Number(permsFaites) || 0) : '', permsRequises: permsRequises,
    lignes: lignes, loyer: _round2(loyer), commission: commission, taux: taux, total: _round2(lignes.reduce(function (s, l) { return s + l.montant; }, 0)),
    ventes: ventesCA, frais: frais, prime: prime, net: _round2(prime - commission), nbVentes: vs.length, releve: releve, alertes: alertes
  };
}

/** Créateurs à facturer pour le mois : un emplacement pendant le mois ou des ventes le mois précédent (hors équipe). */
function _aFacturer(ctx) {
  const tC = ctx.tC, tE = ctx.tE, b = ctx.b, avecEmpl = {};
  tE.lignes.forEach(function (e) {
    const d = _val(tE, e, 'debut'), f = _val(tE, e, 'fin');
    if (d instanceof Date && d <= b.fin && !(f instanceof Date && f < b.debut)) avecEmpl[String(_val(tE, e, 'id_createur'))] = true;
  });
  return tC.lignes.filter(function (r) {
    const id = String(_val(tC, r, 'id_createur'));
    return id && String(_val(tC, r, 'categorie')) !== 'shop' && (avecEmpl[id] || ctx.ventes[id] || ctx.factures[id]);
  });
}

/* ---------- Actions de la gestion ---------- */

function _gFactures(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), ctx = _contexteFactures(ss, body.mois);
  const perms = body.perms || {};
  const lignes = _aFacturer(ctx).map(function (r) {
    const id = String(_val(ctx.tC, r, 'id_createur')), f = ctx.factures[id];
    const c = _calculFacture(ctx, r, perms[id] != null ? perms[id] : (f ? f['perms_faites'] : 0));
    delete c.releve;
    c.facture = f ? { numero: String(f['numero']), statut: String(f['statut']), total: Number(f['total']) || 0, emiseLe: _iso(f['emise_le']),
                      envoyeeLe: _iso(f['envoyee_le']), pdfId: String(f['pdf_id'] || '') } : null;
    return c;
  }).sort(function (a, b) { return a.nom.localeCompare(b.nom, 'fr', { sensitivity: 'base' }); });
  return { ok: true, mois: ctx.mois, libelle: ctx.b.libelle, moisVentes: ctx.precBornes.libelle, joursOuverts: ctx.ouvertsMois, factures: lignes };
}

/** Aperçu : la facture générée (montants figés) si elle existe, sinon le calcul du moment. */
function _gFactureApercu(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), ctx = _contexteFactures(ss, body.mois);
  const r = ctx.tC.lignes.filter(function (x) { return String(_val(ctx.tC, x, 'id_createur')) === String(body.idCreateur); })[0];
  if (!r) throw new Error('Créateur introuvable.');
  const c = _calculFacture(ctx, r, body.permsFaites);
  const f = ctx.factures[c.idCreateur];
  if (f) _appliquerFigee(ss, c, f);
  return { ok: true, html: _htmlFacture(ctx, c, f ? String(f['numero']) : null, f ? f['emise_le'] : new Date()), email: _emailFacture(ss, ctx, c, f ? String(f['numero']) : '(numéro attribué à la génération)') };
}
/** Remplace le calcul par les montants figés d'une facture déjà générée. */
function _appliquerFigee(ss, c, f) {
  const num = String(f['numero']);
  c.lignes = _lireTable(_onglet(ss, SHEET_LIGNES_FACTURE)).filter(function (l) { return _numeroTexte(l['numero']) === num; })
    .sort(function (a, b) { return Number(a['ordre']) - Number(b['ordre']); })
    .map(function (l) { return { ref: String(l['ref']), libelle: String(l['libelle']), montant: Number(l['montant']) || 0 }; });
  c.total = Number(f['total']) || 0; c.loyer = Number(f['loyer']) || 0; c.commission = Number(f['commission']) || 0; c.taux = Number(f['taux_commission']) || 0;
  c.ventes = Number(f['ventes_mois_precedent']) || 0; c.frais = Number(f['frais_mois_precedent']) || 0; c.prime = Number(f['prime_mois_precedent']) || 0; c.net = Number(f['net_a_verser']) || 0;
}

/** Génère UNE facture (appelée créateur par créateur par le site) : numéro, montants figés, PDF dans Drive. */
function _gFactureGenerer(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _creerOngletSiAbsent(ss, SHEET_FACTURES, COLONNES_FACTURES);
  _creerOngletSiAbsent(ss, SHEET_LIGNES_FACTURE, COLONNES_LIGNES_FACTURE);
  _modeleFactureSiAbsent(ss);
  _colonnesTexteFactures(ss);
  const ctx = _contexteFactures(ss, body.mois);
  if (ctx.factures[String(body.idCreateur)]) throw new Error('Cette facture est déjà générée (n° ' + ctx.factures[String(body.idCreateur)]['numero'] + ').');
  const r = ctx.tC.lignes.filter(function (x) { return String(_val(ctx.tC, x, 'id_createur')) === String(body.idCreateur); })[0];
  if (!r) throw new Error('Créateur introuvable.');
  const c = _calculFacture(ctx, r, body.permsFaites);
  if (!c.lignes.length) throw new Error('Rien à facturer pour ' + c.nom + ' ce mois-ci.');
  if (!c.nomLegal && !c.adresse) throw new Error('Nom légal et adresse manquants pour ' + c.nom + ' : à compléter dans Gestion ▸ Créateurs.');

  // Numéro : AA-MM-NNN, compteur sur l'année (jamais réutilisé, même si une facture est annulée).
  const tF = _tableau(ss, SHEET_FACTURES), aa = ctx.mois.slice(2, 4);
  const n = tF.lignes.reduce(function (m, x) { const k = new RegExp('^' + aa + '-\\d{2}-(\\d+)$').exec(_numeroTexte(_val(tF, x, 'numero'))); return k ? Math.max(m, +k[1]) : m; }, 0);
  const numero = aa + '-' + ctx.mois.slice(5, 7) + '-' + String(n + 1).padStart(3, '0');
  const emiseLe = new Date();

  const pdf = HtmlService.createHtmlOutput(_htmlFacture(ctx, c, numero, emiseLe)).getBlob().getAs(MimeType.PDF)
    .setName('FACTURE-' + numero + '-' + _cleNom(c.nom).normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/[^A-Z0-9]+/g, '-') + '.pdf');
  const fichier = _dossierFactures(ctx.mois).createFile(pdf);

  _ajouterLigne(tF, {
    numero: numero, mois: ctx.mois, id_createur: c.idCreateur, emise_le: emiseLe, total: c.total, loyer: c.loyer, commission: c.commission,
    taux_commission: c.taux, ventes_mois_precedent: c.ventes, frais_mois_precedent: c.frais, prime_mois_precedent: c.prime, net_a_verser: c.net,
    perms_faites: c.permsFaites, perms_requises: c.permsRequises == null ? '' : c.permsRequises, statut: 'generee', pdf_id: fichier.getId(), emise_par: _signataire()
  });
  const tL = _tableau(ss, SHEET_LIGNES_FACTURE);
  _assurerTaille(tL.sh, tL.sh.getLastRow() + c.lignes.length, COLONNES_LIGNES_FACTURE.length);
  tL.sh.getRange(tL.sh.getLastRow() + 1, 1, c.lignes.length, COLONNES_LIGNES_FACTURE.length)
    .setValues(c.lignes.map(function (l, i) { return [numero, i + 1, l.ref, _textePublic(l.libelle, 300), l.montant]; }));
  _journaliser('facture_generee', numero + ' · ' + c.idCreateur + ' ' + c.nom + ' · ' + _eurFr(c.total));
  return { ok: true, numero: numero, total: c.total };
}

function _dossierFactures(mois) {
  const sous = function (parent, nom) { const it = parent.getFoldersByName(nom); return it.hasNext() ? it.next() : parent.createFolder(nom); };
  return sous(sous(sous(DriveApp.getRootFolder(), '13h59'), 'Factures'), mois);
}

function _ligneFacture(ss, numero) {
  const t = _tableau(ss, SHEET_FACTURES);
  const i = t.lignes.findIndex(function (r) { return _numeroTexte(_val(t, r, 'numero')) === String(numero); });
  if (i < 0) throw new Error('Facture introuvable : « ' + numero + ' ».');
  return { t: t, i: i, r: t.lignes[i] };
}

function _gFactureEnvoyer(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), x = _ligneFacture(ss, body.numero), t = x.t, r = x.r;
  if (String(_val(t, r, 'statut')) === 'annulee') throw new Error('Facture annulée : elle ne s\'envoie pas.');
  const mois = _moisTexte(_val(t, r, 'mois')), ctx = _contexteFactures(ss, mois);
  const rc = ctx.tC.lignes.filter(function (c) { return String(_val(ctx.tC, c, 'id_createur')) === String(_val(t, r, 'id_createur')); })[0];
  const email = rc ? _norm(_val(ctx.tC, rc, 'email')) : '';
  if (!email) throw new Error("Pas d'adresse e-mail pour ce créateur : à compléter dans Gestion ▸ Créateurs.");
  const c = _calculFacture(ctx, rc, _val(t, r, 'perms_faites'));
  _appliquerFigee(ss, c, { numero: body.numero, total: _val(t, r, 'total'), loyer: _val(t, r, 'loyer'), commission: _val(t, r, 'commission'), taux_commission: _val(t, r, 'taux_commission'),
    ventes_mois_precedent: _val(t, r, 'ventes_mois_precedent'), frais_mois_precedent: _val(t, r, 'frais_mois_precedent'), prime_mois_precedent: _val(t, r, 'prime_mois_precedent'), net_a_verser: _val(t, r, 'net_a_verser') });
  const e = _emailFacture(ss, ctx, c, String(body.numero));
  if (!e.actif) throw new Error("Le modèle d'e-mail « facture » est désactivé dans Réglages ▸ E-mails types.");
  const pdf = DriveApp.getFileById(String(_val(t, r, 'pdf_id'))).getBlob();
  envoyerEmailShop({ to: email, subject: e.objet, texte: e.texte, html: _htmlShop(e.texte), pieces: [pdf] });
  r[t.M['statut']] = 'envoyee'; r[t.M['envoyee_le']] = new Date(); r[t.M['envoyee_a']] = email;
  _ecrireLigne(t, x.i, r);
  _journaliser('facture_envoyee', body.numero + ' → ' + email);
  return { ok: true };
}

function _gFactureAnnuler(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), x = _ligneFacture(ss, body.numero), t = x.t, r = x.r;
  if (String(_val(t, r, 'statut')) === 'envoyee') throw new Error('Facture déjà envoyée : elle ne s\'annule pas ici (il faudra un avoir).');
  const motif = _motif(body);
  r[t.M['statut']] = 'annulee'; r[t.M['remarque']] = motif;
  _ecrireLigne(t, x.i, r);
  _journaliser('facture_annulee', body.numero + ' : ' + motif);
  return { ok: true };
}

function _gFacturePdf(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), x = _ligneFacture(ss, body.numero);
  const id = String(_val(x.t, x.r, 'pdf_id'));
  return { ok: true, url: 'https://drive.google.com/file/d/' + id + '/view' };
}

/* ---------- E-mail ---------- */

function _emailFacture(ss, ctx, c, numero) {
  const m = _modelesEmails(ss)['facture'] || { objet: EMAILS_PAR_DEFAUT.filter(function (l) { return l[0] === 'facture'; })[0][1], texte: EMAILS_PAR_DEFAUT.filter(function (l) { return l[0] === 'facture'; })[0][2], actif: true };
  const vars = { prenom: c.nom, marque: c.nom, mois: ctx.b.libelle, mois_ventes: ctx.precBornes.libelle, numero: numero };
  return { objet: _remplir(m.objet, vars), texte: _remplir(m.texte, vars), actif: m.actif, a: c.email };
}

/* ---------- Document (aperçu et PDF) ----------
 * Mise en page en tableaux : le convertisseur PDF de Google ne gère ni flex ni grid. */
function _htmlFacture(ctx, c, numero, emiseLe) {
  const e = _echapperHtml;
  const cell = 'padding:6px 5px;border-bottom:1px solid #eae7e1;font-size:11px;';
  const pied = '<div style="margin-top:18px;border-top:1px solid #eae7e1;padding-top:6px;font-size:8.5px;color:#8b887f;line-height:1.45">' +
    'Clause de réserve de propriété (loi 80.335 du 12 mai 1980). Pénalité de retard : 3 fois le taux d\'intérêt légal après échéance. Escompte pour règlement anticipé : 0 %. ' +
    'Indemnité forfaitaire pour frais de recouvrement : 40 €.</div>';
  const bandeau = function (droite) {
    return '<table style="width:100%;border-collapse:collapse;background:#000;color:#fff"><tr><td style="padding:12px 18px;font-size:17px;font-weight:bold;letter-spacing:4px">13H59</td>' +
      '<td style="padding:12px 18px;text-align:right;font-size:11px;font-weight:bold;letter-spacing:2px">' + droite + '</td></tr></table>';
  };
  const dest = [c.nomLegal, c.adresse, [c.codePostal, c.ville].filter(String).join(' ')].filter(String).map(e).join('<br>');
  let p1 = bandeau(e(ctx.b.libelle.toUpperCase())) + '<div style="padding:16px 18px">' +
    '<table style="width:100%;border-collapse:collapse"><tr>' +
      '<td style="vertical-align:top;width:50%;font-size:11px;line-height:1.5"><b style="text-transform:uppercase">' + ASSO.nom + '</b><br>' + ASSO.adresse + '<br>' + ASSO.ville + '<br>' + ASSO.pays + '<br>' + ASSO.email +
        '<div style="color:#6f6c65;font-size:10px;margin-top:5px">N° TVA intracommunautaire : ' + ASSO.tva + '<br>N° SIRET : ' + ASSO.siret + '<br>Code NAF : ' + ASSO.naf + '</div></td>' +
      '<td style="vertical-align:top;width:50%;background:#f3f1ec;padding:9px 11px;font-size:11px;line-height:1.5"><b style="text-transform:uppercase">' + e(c.nom) + '</b><br>' + dest +
        (c.email ? '<div style="color:#6f6c65;font-size:10px;margin-top:5px">' + e(c.email) + '</div>' : '') + '</td>' +
    '</tr></table>' +
    '<table style="width:100%;border-collapse:collapse;margin-top:16px;border-bottom:2px solid #0f0f0f"><tr><td style="font-size:14px;font-weight:bold;padding-bottom:5px">' +
      (numero ? 'FACTURE N° ' + e(numero) : 'APERÇU · FACTURE NON GÉNÉRÉE') + '</td><td style="text-align:right;font-size:11px;color:#8b887f;padding-bottom:5px">Le ' + _dateFr(emiseLe) + '</td></tr></table>' +
    '<table style="width:100%;border-collapse:collapse;margin-top:6px"><tr>' +
      ['Réf.', 'Désignation', 'Qté', 'PU', 'TVA', 'Montant HT'].map(function (h, i) { return '<th style="' + cell + 'font-size:9px;color:#8b887f;text-transform:uppercase;text-align:' + (i === 1 ? 'left' : i === 0 ? 'left' : 'right') + '">' + h + '</th>'; }).join('') + '</tr>' +
      c.lignes.map(function (l) {
        return '<tr><td style="' + cell + 'font-weight:bold">' + e(l.ref) + '</td><td style="' + cell + '">' + e(l.libelle) + '</td><td style="' + cell + 'text-align:right">1,00</td>' +
          '<td style="' + cell + 'text-align:right">' + _eurFr(l.montant) + '</td><td style="' + cell + 'text-align:right">0,00</td><td style="' + cell + 'text-align:right">' + _eurFr(l.montant) + '</td></tr>';
      }).join('') + '</table>' +
    (c.benevole ? '<div style="font-size:10px;color:#8b887f;margin-top:5px">Bénévole : aucune commission n\'est prise sur les ventes.</div>' : '') +
    (c.ventes > 0 ? '<div style="font-size:10px;color:#8b887f;margin-top:5px">Détail des ventes ' + e(_de(ctx.precBornes.libelle)) + ' et du calcul de la commission en page 2.</div>' : '') +
    '<table style="width:100%;border-collapse:collapse;margin-top:14px"><tr>' +
      '<td style="vertical-align:top;font-size:10px;color:#6f6c65;line-height:1.5;padding-right:18px">Le loyer est payé par virement entre le 20 et le 25 du mois précédent. Les commissions sont déduites des ventes reversées et n\'ont pas à être payées.</td>' +
      '<td style="vertical-align:top;width:210px"><table style="width:100%;border-collapse:collapse;font-size:11px">' +
        '<tr><td style="padding:3px 0">Total HT</td><td style="text-align:right">' + _eurFr(c.total) + '</td></tr>' +
        '<tr><td style="padding:3px 0">TVA (0 %)</td><td style="text-align:right">0,00 €</td></tr>' +
        '<tr><td style="padding:6px 0 0;border-top:2px solid #0f0f0f;font-weight:bold;font-size:13px">Total TTC</td><td style="padding:6px 0 0;border-top:2px solid #0f0f0f;text-align:right;font-weight:bold;font-size:13px">' + _eurFr(c.total) + '</td></tr>' +
      '</table></td></tr></table>' + pied +
    '<div style="text-align:right;font-size:8.5px;color:#8b887f;margin-top:5px">Page 1 / 2</div></div>';

  // Page 2 : relevé des ventes (information, ne fait pas partie de la facture).
  const rel = c.releve || [];
  let p2 = '<div style="page-break-before:always"></div>' + bandeau('RELEVÉ DES VENTES · ' + e(ctx.precBornes.libelle.toUpperCase())) + '<div style="padding:16px 18px">' +
    '<table style="width:100%;border-collapse:collapse;border-bottom:2px solid #0f0f0f"><tr><td style="font-size:13px;font-weight:bold;padding-bottom:5px;text-transform:uppercase">' + e(c.nom) + '</td>' +
      '<td style="text-align:right;font-size:10px;color:#8b887f;padding-bottom:5px">' + (numero ? 'Annexe à la facture n° ' + e(numero) : 'Aperçu') + '</td></tr></table>';
  if (rel.length) {
    p2 += '<table style="width:100%;border-collapse:collapse;margin-top:6px"><tr>' +
      ['Date', 'Référence', 'Paiement', 'Prix client', 'Frais'].map(function (h, i) { return '<th style="' + cell + 'font-size:9px;color:#8b887f;text-transform:uppercase;text-align:' + (i < 3 ? 'left' : 'right') + '">' + h + '</th>'; }).join('') + '</tr>' +
      rel.map(function (v) {
        const annul = v.type === 'annulation';
        return '<tr' + (annul ? ' style="color:#c0392b"' : '') + '><td style="' + cell + '">' + v.date.slice(8, 10) + '/' + v.date.slice(5, 7) + '</td><td style="' + cell + '">' + e(v.reference) +
          (v.remise && v.remise !== 'pas_de_remise' ? ' <span style="color:#8b887f">(' + e(v.remise) + ')</span>' : '') + '</td><td style="' + cell + '">' + e(v.paiement) + '</td>' +
          '<td style="' + cell + 'text-align:right">' + _eurFr(v.prixClient) + '</td><td style="' + cell + 'text-align:right">' + (v.frais ? _eurFr(v.frais) : '—') + '</td></tr>';
      }).join('') + '</table>';
    const comLib = c.benevole ? 'Commission : bénévole, aucune commission' : c.taux ? 'Commission ' + Math.round(c.taux * 100) + ' % (prime du mois ' + (c.taux >= 0.15 ? 'dès 250 €' : 'dès 100 €') + ')' : 'Commission : 0 % (prime du mois sous 100 €)';
    const ligneCalc = function (g, d, fort) { return '<tr><td style="padding:3px 0;' + (fort ? 'border-top:2px solid #0f0f0f;font-weight:bold;font-size:13px;padding-top:6px' : '') + '">' + g + '</td><td style="text-align:right;' + (fort ? 'border-top:2px solid #0f0f0f;font-weight:bold;font-size:13px;padding-top:6px' : '') + '">' + d + '</td></tr>'; };
    p2 += '<table style="width:100%;border-collapse:collapse;margin-top:12px;font-size:11px;border:1px solid #eae7e1"><tr><td style="padding:8px 10px"><table style="width:100%;border-collapse:collapse">' +
      ligneCalc('Ventes ' + e(_de(ctx.precBornes.libelle)) + ' (' + rel.length + ' ligne' + (rel.length > 1 ? 's' : '') + ')', _eurFr(c.ventes)) +
      ligneCalc('Frais de paiement (carte bancaire)', '− ' + _eurFr(c.frais)) +
      ligneCalc(comLib, '− ' + _eurFr(c.commission)) +
      ligneCalc('Net à te verser', _eurFr(c.net), true) + '</table></td></tr></table>';
  } else {
    p2 += '<div style="margin-top:12px;font-size:11px">Aucune vente en ' + e(ctx.precBornes.libelle) + '.</div>';
  }
  p2 += '<div style="margin-top:18px;border-top:1px solid #eae7e1;padding-top:6px;font-size:8.5px;color:#8b887f;line-height:1.45">Document d\'information joint à la facture : il détaille le calcul de ta commission et de ton virement, et ne fait pas partie de la facture. ' +
    'Barème : 0 % jusqu\'à 100 € de prime dans le mois, 10 % dès 100 €, 15 % dès 250 € (prime = prix client − frais de paiement).</div>' +
    '<div style="text-align:right;font-size:8.5px;color:#8b887f;margin-top:5px">Page 2 / 2</div></div>';
  return '<html><head><meta charset="utf-8"></head><body style="margin:0;font-family:Arial,Helvetica,sans-serif;color:#1a1a1a">' + p1 + p2 + '</body></html>';
}

/**
 * À lancer depuis l'éditeur, UNIQUEMENT tant qu'aucune facture n'a été envoyée : efface les factures
 * d'essai (lignes des onglets factures et lignes_facture, PDF mis à la corbeille de Drive), pour que la
 * vraie numérotation commence à 001. Refuse de tourner dès qu'une facture a été envoyée à un créateur.
 */
function effacerFacturesDEssai() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shF = ss.getSheetByName(SHEET_FACTURES), shL = ss.getSheetByName(SHEET_LIGNES_FACTURE);
  if (!shF) { Logger.log('Aucune facture.'); return; }
  const f = _lireTable(shF);
  if (f.some(function (x) { return String(x['statut']) === 'envoyee'; })) throw new Error("Une facture a déjà été envoyée : les factures ne s'effacent plus (numérotation légale).");
  let pdf = 0;
  f.forEach(function (x) { if (x['pdf_id']) { try { DriveApp.getFileById(String(x['pdf_id'])).setTrashed(true); pdf++; } catch (e) { /* déjà supprimé */ } } });
  if (shF.getLastRow() > 1) shF.deleteRows(2, shF.getLastRow() - 1);
  if (shL && shL.getLastRow() > 1) shL.deleteRows(2, shL.getLastRow() - 1);
  _colonnesTexteFactures(ss);
  _journaliser('factures_essai_effacees', f.length + ' facture(s) d\'essai effacée(s), ' + pdf + ' PDF mis à la corbeille', Session.getEffectiveUser().getEmail());
  Logger.log('✅ ' + f.length + " facture(s) d'essai effacée(s), " + pdf + ' PDF mis à la corbeille de Drive. La numérotation repart à 001.');
}

/** À lancer UNE fois depuis l'éditeur : autorise l'accès à Drive (dossier des factures) et vérifie le modèle d'e-mail. */
function autoriserFacturation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dossier = _dossierFactures(_jourIso(new Date()).slice(0, 7));
  _modeleFactureSiAbsent(ss);
  Logger.log('✅ Dossier des factures prêt : ' + dossier.getUrl() + '\n✅ Modèle d\'e-mail « facture » présent (Réglages ▸ E-mails types).');
}
