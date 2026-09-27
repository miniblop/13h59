/*************************************************************
 *  BILAN ANNUEL — Gestion ▸ Statistiques ▸ Bilan
 *  Compte de résultat de l'association, mois par mois, sur un exercice :
 *   - recettes calculées par le site (loyers, commissions, ventes du shop,
 *     adhésions) ; un montant saisi pour un mois remplace le calcul ;
 *   - charges décrites une fois (poste, catégorie, montant, fréquence, dates),
 *     avec le montant réel d'un mois quand il varie (facture ENGIE…) ;
 *   - hors résultat : prêts (apport, remboursements), investissements
 *     (amortis sur plusieurs années), dépôt de garantie → trésorerie ;
 *   - activité : ventes des créateurs, stands occupés, seuil de rentabilité,
 *     manque à gagner, bénévolat valorisé.
 *  Les ventes des créateurs ne sont pas des recettes de l'association (elle
 *  encaisse pour leur compte) ; les frais de carte bancaire, refacturés aux
 *  créateurs, sont neutres.
 *************************************************************/

const SHEET_CHARGES = 'charges';
const SHEET_CHARGES_MOIS = 'charges_mois';
const SHEET_RECETTES_MOIS = 'recettes_mois';
const SHEET_EXERCICES = 'exercices';
const SHEET_PARAMETRES_BILAN = 'parametres_bilan';
const COLONNES_CHARGES = ['id_charge', 'poste', 'categorie', 'nature', 'montant', 'frequence', 'debut', 'fin', 'montant_total', 'amortissement_annees', 'remarque', 'modifie_le'];
const COLONNES_CHARGES_MOIS = ['mois', 'id_charge', 'montant', 'remarque', 'source', 'modifie_le'];
const COLONNES_RECETTES_MOIS = ['mois', 'type', 'montant', 'remarque', 'source', 'modifie_le'];
const COLONNES_EXERCICES = ['code', 'libelle', 'debut', 'fin'];
const COLONNES_PARAMETRES_BILAN = ['cle', 'valeur', 'libelle'];

const CATEGORIES_CHARGES = [
  ['local', 'Local (loyer, charges, PNO)'], ['energie', 'Énergie'], ['assurances', 'Assurances'], ['banque', 'Banque et TPE'],
  ['telecom', 'Internet et téléphone'], ['fonctionnement', 'Fonctionnement (café, consommables)'], ['fournitures', 'Fournitures et emballages'],
  ['communication', 'Communication et PLV'], ['evenements', 'Événements'], ['reglementaire', 'Obligations réglementaires (DASRI, ARS)'], ['autres', 'Autres charges']
];
const NATURES_CHARGES = ['charge', 'pret', 'investissement', 'garantie'];
const FREQUENCES_CHARGES = ['mensuel', 'annuel', 'ponctuel'];
const TYPES_RECETTES = [
  ['loyers', 'Loyers des stands', true], ['commissions', 'Commissions sur les ventes', true], ['vente_shop', 'Ventes des produits du shop', true],
  ['adhesions', 'Adhésions', true], ['subventions', 'Subventions', false], ['dons', 'Dons', false], ['ateliers', 'Ateliers et événements', false], ['autres', 'Autres recettes', false]
];
const PARAMETRES_BILAN_DEFAUT = [
  ['adhesion', 15, "Montant de l'adhésion annuelle (€)"],
  ['smic_horaire', 11.88, 'SMIC horaire brut pour valoriser le bénévolat (€) : à mettre à jour chaque année']
];
const EXERCICES_DEFAUT = [['2025-2026', 'Exercice 2025-2026', '2025-09', '2026-12'], ['2027', 'Exercice 2027', '2027-01', '2027-12']];

function _moisPlus(mois, n) {
  const b = _moisBornes(mois), d = new Date(b.debut.getFullYear(), b.debut.getMonth() + n, 1);
  return d.getFullYear() + '-' + _deux(d.getMonth() + 1);
}
function _moisValide(m) { return /^\d{4}-(0[1-9]|1[0-2])$/.test(String(m || '')); }

/** Crée les onglets du bilan s'ils manquent (colonnes « mois » en texte). */
function _ongletsBilan(ss) {
  _creerOngletSiAbsent(ss, SHEET_CHARGES, COLONNES_CHARGES);
  _creerOngletSiAbsent(ss, SHEET_CHARGES_MOIS, COLONNES_CHARGES_MOIS);
  _creerOngletSiAbsent(ss, SHEET_RECETTES_MOIS, COLONNES_RECETTES_MOIS);
  if (!ss.getSheetByName(SHEET_EXERCICES)) {
    _creerOngletSiAbsent(ss, SHEET_EXERCICES, COLONNES_EXERCICES);
    const sh = ss.getSheetByName(SHEET_EXERCICES);
    sh.getRange(1, 1, 100, 4).setNumberFormat('@');
    sh.getRange(2, 1, EXERCICES_DEFAUT.length, 4).setValues(EXERCICES_DEFAUT);
  }
  if (!ss.getSheetByName(SHEET_PARAMETRES_BILAN)) {
    _creerOngletSiAbsent(ss, SHEET_PARAMETRES_BILAN, COLONNES_PARAMETRES_BILAN);
    ss.getSheetByName(SHEET_PARAMETRES_BILAN).getRange(2, 1, PARAMETRES_BILAN_DEFAUT.length, 3).setValues(PARAMETRES_BILAN_DEFAUT);
  }
  [[SHEET_CHARGES, [7, 8]], [SHEET_CHARGES_MOIS, [1]], [SHEET_RECETTES_MOIS, [1]]].forEach(function (x) {
    const sh = ss.getSheetByName(x[0]);
    x[1].forEach(function (c) { sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'); });
  });
}

function _parametresBilan(ss) {
  const p = {};
  PARAMETRES_BILAN_DEFAUT.forEach(function (d) { p[d[0]] = d[1]; });
  _lireTable(_onglet(ss, SHEET_PARAMETRES_BILAN)).forEach(function (r) { if (r['cle'] && r['valeur'] !== '') p[String(r['cle'])] = Number(r['valeur']); });
  return p;
}
function _exercices(ss) {
  return _lireTable(_onglet(ss, SHEET_EXERCICES)).filter(function (r) { return r['code']; }).map(function (r) {
    return { code: String(r['code']), libelle: String(r['libelle'] || r['code']), debut: _moisTexte(r['debut']), fin: _moisTexte(r['fin']) };
  }).filter(function (e) { return _moisValide(e.debut) && _moisValide(e.fin) && e.debut <= e.fin; }).sort(function (a, b) { return a.debut.localeCompare(b.debut); });
}
function _charges(ss) {
  return _lireTable(_onglet(ss, SHEET_CHARGES)).filter(function (r) { return r['id_charge']; }).map(function (r) {
    return { id: String(r['id_charge']), poste: String(r['poste'] || ''), categorie: String(r['categorie'] || 'autres'), nature: String(r['nature'] || 'charge'),
      montant: Number(r['montant']) || 0, frequence: String(r['frequence'] || 'mensuel'), debut: _moisTexte(r['debut']), fin: _moisTexte(r['fin']),
      montantTotal: Number(r['montant_total']) || 0, amortissement: Number(r['amortissement_annees']) || 0, remarque: String(r['remarque'] || '') };
  });
}
function _moisDansCharge(c, m) { return (!c.debut || m >= c.debut) && (!c.fin || m <= c.fin); }

/* ---------- Calcul ---------- */

/** Tout l'exercice en une lecture par onglet : [{ mois, recettes, charges, … }]. */
function _calculBilan(ss, ex) {
  const moisListe = [];
  for (let m = ex.debut; m <= ex.fin; m = _moisPlus(m, 1)) moisListe.push(m);
  const dans = {};
  moisListe.forEach(function (m) { dans[m] = true; });
  const param = _parametresBilan(ss), charges = _charges(ss);
  // saisies
  const reels = {}, saisies = {};
  _lireTable(_onglet(ss, SHEET_CHARGES_MOIS)).forEach(function (r) {
    const m = _moisTexte(r['mois']); if (dans[m] && r['id_charge'] && r['montant'] !== '') reels[m + '|' + r['id_charge']] = { montant: Number(r['montant']) || 0, remarque: String(r['remarque'] || '') };
  });
  _lireTable(_onglet(ss, SHEET_RECETTES_MOIS)).forEach(function (r) {
    const m = _moisTexte(r['mois']); if (dans[m] && r['type'] && r['montant'] !== '') saisies[m + '|' + r['type']] = { montant: Number(r['montant']) || 0, remarque: String(r['remarque'] || ''), source: String(r['source'] || '') };
  });
  // créateurs
  const fiches = {};
  _lireTable(_onglet(ss, SHEET_CREATEURS)).forEach(function (c) {
    if (c['id_createur']) fiches[String(c['id_createur'])] = { shop: String(c['categorie']) === 'shop', benevole: c['benevole'] === true, adhesion: c['adhesion_payee_le'] instanceof Date ? _jourIso(c['adhesion_payee_le']).slice(0, 7) : '' };
  });
  // ventes : un seul passage
  const V = {};
  [_moisPlus(ex.debut, -1)].concat(moisListe).forEach(function (m) { V[m] = { ca: 0, lignes: 0, paniers: {}, createurs: {}, primeParCreateur: {}, shop: 0 }; });
  _lireTable(_onglet(ss, SHEET_VENTES)).forEach(function (v) {
    if (!(v['date'] instanceof Date)) return;
    const m = _jourIso(v['date']).slice(0, 7), x = V[m];
    if (!x) return;
    const id = String(v['id_createur'] || ''), f = fiches[id] || {};
    const prime = Number(v['prime']) || 0, pc = Number(v['prix_client']) || 0;
    if (f.shop) { x.shop += prime; return; }
    x.ca += pc;
    x.lignes += _norm(v['type_ligne']) === 'annulation' ? -1 : 1;
    if (v['id_panier'] !== '' && v['id_panier'] != null) x.paniers[String(v['id_panier'])] = true;
    x.createurs[id] = true;
    x.primeParCreateur[id] = (x.primeParCreateur[id] || 0) + prime;
  });
  // factures (loyers figés du mois, commission des ventes du mois précédent) et encaissements
  const fact = {}, lignesBen = {}, recus = {};
  if (ss.getSheetByName(SHEET_FACTURES)) {
    _lireTable(_onglet(ss, SHEET_FACTURES)).forEach(function (f) {
      if (String(f['statut']) === 'annulee') return;
      const m = _moisTexte(f['mois']);
      const x = fact[m] = fact[m] || { loyer: 0, commission: 0, n: 0, parCreateur: {} };
      x.loyer += Number(f['loyer']) || 0; x.commission += Number(f['commission']) || 0; x.n++;
      x.parCreateur[String(f['id_createur'])] = { loyer: Number(f['loyer']) || 0, numero: _numeroTexte(f['numero']) };
    });
    if (ss.getSheetByName(SHEET_LIGNES_FACTURE)) {
      const numMois = {};
      Object.keys(fact).forEach(function (m) { Object.keys(fact[m].parCreateur).forEach(function (id) { numMois[fact[m].parCreateur[id].numero] = m; }); });
      _lireTable(_onglet(ss, SHEET_LIGNES_FACTURE)).forEach(function (l) {
        const m = numMois[_numeroTexte(l['numero'])];
        if (m && String(l['ref']) === 'BEN') lignesBen[m] = (lignesBen[m] || 0) - (Number(l['montant']) || 0);
      });
    }
  }
  if (ss.getSheetByName(SHEET_PAIEMENTS_RECUS)) {
    _lireTable(_onglet(ss, SHEET_PAIEMENTS_RECUS)).forEach(function (p) {
      if (String(p['statut']) === 'annule') return;
      const k = _moisTexte(p['mois_loyer']) + '|' + String(p['id_createur']);
      recus[k] = (recus[k] || 0) + (Number(p['montant']) || 0);
    });
  }
  // stands et emplacements
  const stands = {};
  _lireTable(_onglet(ss, SHEET_STANDS)).forEach(function (s) { if (s['code']) stands[String(s['code'])] = { loyer: Number(s['loyer']) || 0, places: Number(s['places']) || 0 }; });
  const tE = _tableau(ss, SHEET_EMPLACEMENTS);
  const bene = _heuresBenevolat(ex);
  const auj = _jourIso(new Date()).slice(0, 7);

  return moisListe.map(function (m) {
    const b = _moisBornes(m), milieu = new Date(b.debut.getFullYear(), b.debut.getMonth(), 15), x = V[m];
    const ouverts = _joursOuverts(b.debut, b.fin);
    // occupation au 15 du mois et loyers théoriques (au prorata des jours d'ouverture)
    const occupes = {}; let loyerEstime = 0;
    tE.lignes.forEach(function (e) {
      const code = String(_val(tE, e, 'code_stand')), s = stands[code], id = String(_val(tE, e, 'id_createur'));
      if (!s || (fiches[id] || {}).shop) return;
      if (_actifLe(tE, e, milieu)) occupes[code] = (occupes[code] || 0) + 1;
      const d = _val(tE, e, 'debut'), f = _val(tE, e, 'fin');
      if (!(d instanceof Date) || d > b.fin || (f instanceof Date && f < b.debut)) return;
      const de = d > b.debut ? d : b.debut, a = (f instanceof Date && f < b.fin) ? f : b.fin;
      loyerEstime += s.loyer * Math.min(1, ouverts ? _joursOuverts(de, a) / ouverts : 1);
    });
    const places = {}; let vides = 0;
    Object.keys(stands).forEach(function (code) { places[code] = stands[code].places; vides += Math.max(0, stands[code].places - (occupes[code] || 0)) * stands[code].loyer; });
    // recettes : saisie > factures > calcul
    const rec = {};
    const auto = {
      loyers: fact[m] ? { montant: fact[m].loyer, source: 'factures' } : { montant: loyerEstime, source: 'estime' },
      // commissions facturées ce mois-ci, sur les ventes du mois précédent (comme la facture et le fichier de Mo)
      commissions: fact[m] ? { montant: fact[m].commission, source: 'factures' } : (function () {
        const mp = _moisPlus(m, -1), xp = V[mp];
        return { montant: Object.keys(xp.primeParCreateur).reduce(function (s, id) {
          const p = xp.primeParCreateur[id]; return s + ((fiches[id] || {}).benevole ? 0 : _round2(p * _tauxCommission(p, mp)));
        }, 0), source: 'calcule' };
      })(),
      vente_shop: { montant: x.shop, source: 'calcule' },
      adhesions: { montant: Object.keys(fiches).filter(function (id) { return fiches[id].adhesion === m; }).length * (param.adhesion || 0), source: 'calcule' }
    };
    TYPES_RECETTES.forEach(function (t) {
      const s = saisies[m + '|' + t[0]];
      rec[t[0]] = s ? { montant: _round2(s.montant), source: 'saisi', remarque: s.remarque, origine: s.source } : auto[t[0]] ? { montant: _round2(auto[t[0]].montant), source: auto[t[0]].source } : { montant: 0, source: '' };
    });
    // charges
    const parCat = {}, detail = {}, treso = { apports: 0, remboursements: 0, investissements: 0, garanties: 0 };
    let amortissements = 0;
    charges.forEach(function (c) {
      const r = reels[m + '|' + c.id];
      let prevu = 0;
      if (c.nature === 'charge' || c.nature === 'pret') {
        if (_moisDansCharge(c, m)) prevu = c.frequence === 'annuel' ? c.montant / 12 : c.frequence === 'ponctuel' ? (m === c.debut ? c.montant : 0) : c.montant;
      }
      const montant = _round2(r ? r.montant : prevu);
      if (c.nature === 'charge') {
        if (montant) { parCat[c.categorie] = _round2((parCat[c.categorie] || 0) + montant); }
        if (montant || r) detail[c.id] = { montant: montant, reel: !!r, remarque: r ? r.remarque : '' };
      } else if (c.nature === 'pret') {
        treso.remboursements += montant;
        if (m === c.debut) treso.apports += c.montantTotal;
        if (montant || r) detail[c.id] = { montant: montant, reel: !!r, remarque: r ? r.remarque : '' };
      } else if (c.nature === 'investissement') {
        if (m === c.debut) treso.investissements += c.montantTotal;
        if (c.amortissement > 0 && c.debut && m >= c.debut && m < _moisPlus(c.debut, Math.round(c.amortissement * 12))) amortissements += c.montantTotal / (c.amortissement * 12);
        else if (!(c.amortissement > 0) && m === c.debut) parCat[c.categorie] = _round2((parCat[c.categorie] || 0) + c.montantTotal);
      } else if (c.nature === 'garantie') {
        if (m === c.debut) treso.garanties += c.montantTotal;
      }
    });
    // impayés : loyers facturés du mois non reçus (échéance le 25 du mois précédent, donc passée dès le début du mois)
    let impayes = 0;
    if (fact[m] && m <= auj) Object.keys(fact[m].parCreateur).forEach(function (id) { impayes += Math.max(0, fact[m].parCreateur[id].loyer - (recus[m + '|' + id] || 0)); });
    return {
      mois: m, recettes: rec, charges: parCat, detail: detail, amortissements: _round2(amortissements),
      tresorerie: { apports: _round2(treso.apports), remboursements: _round2(treso.remboursements), investissements: _round2(treso.investissements), garanties: _round2(treso.garanties) },
      activite: { ca: _round2(x.ca), ventes: x.lignes, paniers: Object.keys(x.paniers).length, createurs: Object.keys(x.createurs).length, occupes: occupes, places: places, ouverts: ouverts },
      manque: { standsVides: _round2(vides), benevoles: lignesBen[m] != null ? _round2(lignesBen[m]) : null, impayes: fact[m] ? _round2(impayes) : null },
      benevolat: bene ? { heures: _round2(bene[m] || 0) } : null,
      futur: m > auj, enCours: m === auj
    };
  });
}

/** Heures de permanence par mois (agenda BÉNÉVOLES), gardées une heure ; null si l'agenda est illisible. */
function _heuresBenevolat(ex) {
  const cache = CacheService.getScriptCache(), cle = 'benevolat_heures_' + ex.debut + '_' + ex.fin;
  const c = cache.get(cle);
  if (c) return JSON.parse(c);
  try {
    const src = _agendasPermanences();
    if (!src.dedie) return null;
    const out = {}, deb = _moisBornes(ex.debut).debut, f = _moisBornes(ex.fin).fin, fin = new Date(f.getFullYear(), f.getMonth(), f.getDate() + 1);
    src.agendas[0].getEvents(deb, fin).forEach(function (e) {
      if (!e.getGuestList().some(function (g) { return _norm(g.getEmail()) !== EMAIL_SHOP; })) return;
      const h = e.isAllDayEvent() ? 8 : (e.getEndTime() - e.getStartTime()) / 36e5;
      const m = _jourIso(e.getStartTime()).slice(0, 7);
      out[m] = (out[m] || 0) + h;
    });
    try { cache.put(cle, JSON.stringify(out), 3600); } catch (e) { /* sans effet */ }
    return out;
  } catch (e) {
    return null;
  }
}

/* ---------- Actions de Gestion ---------- */

function _gBilan(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _ongletsBilan(ss);
  const exercices = _exercices(ss);
  if (!exercices.length) throw new Error("Aucun exercice défini : ajoute-en un dans les paramètres du bilan.");
  const auj = _jourIso(new Date()).slice(0, 7);
  const ex = exercices.filter(function (e) { return e.code === body.exercice; })[0] ||
    exercices.filter(function (e) { return e.debut <= auj && auj <= e.fin; })[0] || exercices[exercices.length - 1];
  const i = exercices.indexOf(ex), prec = i > 0 ? exercices[i - 1] : null;
  return {
    ok: true, exercice: ex, exercices: exercices, mois: _calculBilan(ss, ex), precedent: prec ? { exercice: prec, mois: _calculBilan(ss, prec) } : null,
    charges: _charges(ss), parametres: _parametresBilan(ss), categories: CATEGORIES_CHARGES, typesRecettes: TYPES_RECETTES,
    stands: _lireTable(_onglet(ss, SHEET_STANDS)).filter(function (s) { return s['code']; }).map(function (s) { return { code: String(s['code']), libelle: String(s['libelle'] || s['code']), loyer: Number(s['loyer']) || 0, places: Number(s['places']) || 0 }; }),
    asso: ASSO, importPossible: !_charges(ss).length
  };
}

function _gBilanCharge(body) {
  const c = body.charge || {};
  const poste = _textePublic(c.poste, 80);
  if (!poste) throw new Error('Indique le nom du poste.');
  if (!CATEGORIES_CHARGES.some(function (x) { return x[0] === c.categorie; })) throw new Error('Catégorie inconnue.');
  if (NATURES_CHARGES.indexOf(c.nature) < 0) throw new Error('Nature inconnue.');
  if (FREQUENCES_CHARGES.indexOf(c.frequence) < 0) throw new Error('Fréquence inconnue.');
  if (!_moisValide(c.debut)) throw new Error('Indique le mois de début.');
  if (c.fin && (!_moisValide(c.fin) || c.fin < c.debut)) throw new Error('Mois de fin invalide.');
  const n = function (v, lib) { const x = Number(String(v == null || v === '' ? 0 : v).replace(',', '.')); if (isNaN(x) || x < 0 || x > 1e6) throw new Error(lib + ' invalide.'); return _round2(x); };
  const valeurs = { poste: poste, categorie: c.categorie, nature: c.nature, montant: n(c.montant, 'Montant'), frequence: c.frequence, debut: c.debut, fin: c.fin || '',
    montant_total: n(c.montantTotal, 'Montant total'), amortissement_annees: c.amortissement === '' || c.amortissement == null ? '' : n(c.amortissement, "Durée d'amortissement"),
    remarque: _textePublic(c.remarque, 300), modifie_le: new Date() };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _ongletsBilan(ss);
  const t = _tableau(ss, SHEET_CHARGES);
  let id = String(c.id || '');
  if (id) {
    const i = t.lignes.findIndex(function (r) { return String(_val(t, r, 'id_charge')) === id; });
    if (i < 0) throw new Error('Poste introuvable : « ' + id + ' ».');
    const r = t.lignes[i];
    Object.keys(valeurs).forEach(function (k) { if (t.M[k] != null) r[t.M[k]] = valeurs[k]; });
    _ecrireLigne(t, i, r);
  } else {
    id = _prochainId(t, 'id_charge', 'CH');
    valeurs.id_charge = id;
    _ajouterLigne(t, valeurs);
  }
  _journaliser('bilan_charge', id + ' ' + poste + ' · ' + valeurs.montant + ' € ' + valeurs.frequence + ' dès ' + valeurs.debut + (valeurs.fin ? ' jusqu\'à ' + valeurs.fin : ''));
  return { ok: true, id: id };
}

function _gBilanChargeSupprimer(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), t = _tableau(ss, SHEET_CHARGES), id = String(body.id || '');
  const i = t.lignes.findIndex(function (r) { return String(_val(t, r, 'id_charge')) === id; });
  if (i < 0) throw new Error('Poste introuvable.');
  const poste = _val(t, t.lignes[i], 'poste');
  // identifiant jamais réattribué
  PropertiesService.getScriptProperties().setProperty('DERNIER_ID_CH', String(Math.max(_maxId(t, 'id_charge', 'CH'), _dernierId('DERNIER_ID_CH'))));
  t.sh.deleteRow(i + 2);
  const tr = _tableau(ss, SHEET_CHARGES_MOIS);
  for (let j = tr.lignes.length - 1; j >= 0; j--) if (String(_val(tr, tr.lignes[j], 'id_charge')) === id) tr.sh.deleteRow(j + 2);
  _journaliser('bilan_charge_supprimee', id + ' ' + poste);
  return { ok: true };
}

/** Montant réel d'un poste pour un mois (montant vide = on revient au montant prévu). */
function _gBilanReel(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!_moisValide(body.mois)) throw new Error('Mois invalide.');
  if (!_charges(ss).some(function (c) { return c.id === String(body.idCharge); })) throw new Error('Poste introuvable.');
  return _saisieMois(ss, SHEET_CHARGES_MOIS, 'id_charge', String(body.idCharge), body, 'bilan_reel');
}
/** Recette saisie pour un mois (montant vide = on revient au calcul du site). */
function _gBilanRecette(body) {
  if (!_moisValide(body.mois)) throw new Error('Mois invalide.');
  if (!TYPES_RECETTES.some(function (t) { return t[0] === body.type; })) throw new Error('Type de recette inconnu.');
  return _saisieMois(SpreadsheetApp.getActiveSpreadsheet(), SHEET_RECETTES_MOIS, 'type', String(body.type), body, 'bilan_recette');
}
function _saisieMois(ss, onglet, cle, valeurCle, body, action) {
  _ongletsBilan(ss);
  const t = _tableau(ss, onglet);
  const i = t.lignes.findIndex(function (r) { return _moisTexte(_val(t, r, 'mois')) === body.mois && String(_val(t, r, cle)) === valeurCle; });
  const vide = body.montant === '' || body.montant == null;
  if (vide) {
    if (i >= 0) t.sh.deleteRow(i + 2);
    _journaliser(action, body.mois + ' ' + valeurCle + ' : saisie retirée');
    return { ok: true };
  }
  const montant = Number(String(body.montant).replace(',', '.'));
  if (isNaN(montant) || Math.abs(montant) > 1e6) throw new Error('Montant invalide.');
  const valeurs = { mois: body.mois, montant: _round2(montant), remarque: _textePublic(body.remarque, 200), source: 'site', modifie_le: new Date() };
  valeurs[cle] = valeurCle;
  if (i >= 0) { const r = t.lignes[i]; Object.keys(valeurs).forEach(function (k) { if (t.M[k] != null) r[t.M[k]] = valeurs[k]; }); _ecrireLigne(t, i, r); }
  else _ajouterLigne(t, valeurs);
  _journaliser(action, body.mois + ' ' + valeurCle + ' : ' + valeurs.montant + ' €');
  return { ok: true };
}

function _gBilanExercice(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _ongletsBilan(ss);
  const code = _textePublic(body.code, 20), libelle = _textePublic(body.libelle, 60) || 'Exercice ' + code;
  if (!code) throw new Error("Indique un code (ex. 2028).");
  if (!_moisValide(body.debut) || !_moisValide(body.fin) || body.fin < body.debut) throw new Error('Mois de début et de fin invalides.');
  const t = _tableau(ss, SHEET_EXERCICES);
  const autres = _exercices(ss).filter(function (e) { return e.code !== code; });
  const chevauche = autres.filter(function (e) { return !(body.fin < e.debut || body.debut > e.fin); })[0];
  if (chevauche) throw new Error('Ces mois chevauchent l\'exercice « ' + chevauche.libelle + ' ».');
  const i = t.lignes.findIndex(function (r) { return String(_val(t, r, 'code')) === code; });
  const v = { code: code, libelle: libelle, debut: body.debut, fin: body.fin };
  if (i >= 0) { const r = t.lignes[i]; Object.keys(v).forEach(function (k) { r[t.M[k]] = v[k]; }); _ecrireLigne(t, i, r); } else _ajouterLigne(t, v);
  _journaliser('bilan_exercice', code + ' ' + libelle + ' : ' + body.debut + ' → ' + body.fin);
  return { ok: true };
}

function _gBilanParametre(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _ongletsBilan(ss);
  const d = PARAMETRES_BILAN_DEFAUT.filter(function (x) { return x[0] === body.cle; })[0];
  if (!d) throw new Error('Paramètre inconnu.');
  const v = Number(String(body.valeur).replace(',', '.'));
  if (isNaN(v) || v < 0 || v > 10000) throw new Error('Valeur invalide.');
  const t = _tableau(ss, SHEET_PARAMETRES_BILAN), i = t.lignes.findIndex(function (r) { return String(_val(t, r, 'cle')) === body.cle; });
  if (i >= 0) { const r = t.lignes[i]; r[t.M['valeur']] = v; _ecrireLigne(t, i, r); } else _ajouterLigne(t, { cle: d[0], valeur: v, libelle: d[2] });
  _journaliser('bilan_parametre', body.cle + ' = ' + v);
  return { ok: true };
}

/* ---------- Reprise du fichier GESTION COMPTABLE de Mo ---------- */

const ID_GESTION_COMPTABLE = '1JjA3M0-WxHckdRaulEKEgWpCFX5MTYDIHMJTHzXJ770';

/** Lit l'onglet « 2025-2026 » : postes de charges, prêts, investissement, et chiffres mois par mois jusqu'au mois dernier. */
function _lireGestionComptable() {
  const sh = SpreadsheetApp.openById(ID_GESTION_COMPTABLE).getSheetByName('2025-2026');
  if (!sh) throw new Error("Onglet « 2025-2026 » introuvable dans GESTION COMPTABLE.");
  const v = sh.getDataRange().getValues();
  const num = function (x) { if (typeof x === 'number') return x; const s = String(x || '').replace(/[\s €]/g, '').replace(',', '.'); return s === '' || isNaN(Number(s)) ? null : Number(s); };
  const txt = function (x) { return _norm(x).normalize('NFD').replace(/[̀-ͯ]/g, ''); };
  // valeur à droite d'un libellé (premier nombre de la ligne après lui)
  const apres = function (libelle) {
    for (let i = 0; i < v.length; i++) for (let j = 0; j < v[i].length; j++) {
      if (txt(v[i][j]) !== txt(libelle)) continue;
      for (let k = j + 1; k < v[i].length && k <= j + 4; k++) { const n = num(v[i][k]); if (n != null) return n; }
    }
    return null;
  };
  const postes = [];
  const ajouter = function (poste, categorie, libelle, frequence) { const n = apres(libelle); if (n != null) postes.push({ poste: poste, categorie: categorie, montant: n, frequence: frequence || 'mensuel' }); };
  ajouter('Loyer du local', 'local', 'Loyer mensuel');
  ajouter('Provision charges foncières et PNO', 'local', 'provision charges fonciere et PNO');
  ajouter('Assurance', 'assurances', 'assurance');
  ajouter('Internet et téléphone', 'telecom', 'internet');
  ajouter('Contrat banque CIC', 'banque', 'contrat banque');
  ajouter('Café', 'fonctionnement', 'Cafe');
  ajouter('Consommables', 'fonctionnement', 'Consommable');
  ajouter('DASRI (déchets de soins)', 'reglementaire', 'Dasri');
  ajouter('ARS', 'reglementaire', 'ARS /ans', 'annuel');
  // prêts et investissement : lignes « <nom>, NN mois, montant, du jj/mm/aaaa au jj/mm/aaaa, …, mensualité, reste à charge »
  const prets = [];
  v.forEach(function (l) {
    l.forEach(function (c, j) {
      const d = /^(\d+) mois$/.exec(String(c).trim());
      if (!d || j === 0) return;
      let nom = ''; for (let k = j - 1; k >= 0 && !nom; k--) nom = String(l[k] || '').trim();
      const reste = l.slice(j + 1).map(function (x) { return x; });
      const montant = num(reste[0]);
      const per = /du (\d{2})\/(\d{2})\/(\d{4}) au (\d{2})\/(\d{2})\/(\d{4})/.exec(reste.join(' '));
      const nombres = reste.slice(2).map(num).filter(function (n) { return n != null; });
      prets.push({ nom: nom, mois: Number(d[1]), montant: montant, debut: per ? per[3] + '-' + per[2] : '', fin: per ? per[6] + '-' + per[5] : '',
        mensualite: nombres[0], resteACharge: nombres[1], note: String(l[l.length - 1] || '') });
    });
  });
  // tableau mensuel
  const iEntete = v.findIndex(function (l) { return txt(l[0]) === 'mois' && txt(l[1]) === 'charges'; });
  const mensuel = [];
  if (iEntete >= 0) {
    const col = {}; v[iEntete].forEach(function (h, j) { col[txt(h)] = j; });
    let m = '2025-09';
    for (let i = iEntete + 1; i < v.length && txt(v[i][0]) !== 'total'; i++) {
      if (!String(v[i][0]).trim()) continue;
      const g = function (h) { return col[h] != null ? num(v[i][col[h]]) : null; };
      mensuel.push({ mois: m, engie: g('engie'), virements: g('frais de virement'), loyers: g('loyer createurs'), commissions: g('commission'),
        venteShop: g('vente du shop'), adhesions: g('hello asso'), depenses: g('depenses') });
      m = _moisPlus(m, 1);
    }
  }
  return { postes: postes, prets: prets, mensuel: mensuel };
}

/** Bouton « Reprendre GESTION COMPTABLE » : n'agit que si aucun poste n'existe encore. */
function _gBilanImporter(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _ongletsBilan(ss);
  if (_charges(ss).length) throw new Error('Des postes de charges existent déjà : reprise annulée pour ne rien dupliquer.');
  const g = _lireGestionComptable(), debut = '2025-09', auj = _jourIso(new Date()).slice(0, 7);
  const apercu = body.apercu === true;
  const tC = _tableau(ss, SHEET_CHARGES), tR = _tableau(ss, SHEET_CHARGES_MOIS), tS = _tableau(ss, SHEET_RECETTES_MOIS);
  const nouveaux = [];
  const poste = function (v) {
    v.id_charge = 'CH' + String(Math.max(_maxId(tC, 'id_charge', 'CH'), _dernierId('DERNIER_ID_CH')) + 1 + nouveaux.length).padStart(3, '0');
    v.modifie_le = new Date(); nouveaux.push(v); return v.id_charge;
  };
  g.postes.forEach(function (p) { poste({ poste: p.poste, categorie: p.categorie, nature: 'charge', montant: p.montant, frequence: p.frequence, debut: debut, fin: '', montant_total: '', amortissement_annees: '', remarque: 'Repris de GESTION COMPTABLE' }); });
  const idEngie = poste({ poste: 'Électricité ENGIE', categorie: 'energie', nature: 'charge', montant: 0, frequence: 'mensuel', debut: debut, fin: '', montant_total: '', amortissement_annees: '', remarque: 'Montant réel à saisir chaque mois' });
  const idVir = poste({ poste: 'Frais de virement', categorie: 'banque', nature: 'charge', montant: 0, frequence: 'mensuel', debut: debut, fin: '', montant_total: '', amortissement_annees: '', remarque: '0,21 € par virement : montant réel à saisir chaque mois' });
  const idDiv = poste({ poste: 'Dépenses diverses (reprise GESTION COMPTABLE)', categorie: 'autres', nature: 'charge', montant: 0, frequence: 'ponctuel', debut: debut, fin: '', montant_total: '', amortissement_annees: '', remarque: 'Colonne « Dépenses » du fichier de Mo, à détailler si possible' });
  g.prets.forEach(function (p) {
    if (p.mensualite != null && p.mensualite > 0) {
      poste({ poste: 'Prêt ' + p.nom, categorie: 'autres', nature: 'pret', montant: p.mensualite, frequence: 'mensuel', debut: p.debut || debut, fin: p.fin || '', montant_total: p.montant || '', amortissement_annees: '', remarque: p.mois + ' mois' });
    } else if (p.resteACharge != null) {
      poste({ poste: p.nom, categorie: 'local', nature: 'investissement', montant: 0, frequence: 'ponctuel', debut: debut, fin: '', montant_total: p.resteACharge, amortissement_annees: 5,
        remarque: 'Coût ' + p.montant + ' € dont ' + Math.abs(p.mensualite || 0) + ' € pris en charge par le propriétaire. ' + p.note + '. Durée d\'amortissement à valider avec le comptable.' });
    }
  });
  const garantie = (function () { const v = SpreadsheetApp.openById(ID_GESTION_COMPTABLE).getSheetByName('2025-2026').getDataRange().getValues(); for (let i = 0; i < v.length; i++) for (let j = 0; j < v[i].length; j++) if (/^d[ée]p[ôo]ts? de garantie$/i.test(String(v[i][j]).trim())) { for (let k = j + 1; k < v[i].length; k++) if (typeof v[i][k] === 'number') return v[i][k]; } return null; })();
  if (garantie) poste({ poste: 'Dépôt de garantie du local', categorie: 'local', nature: 'garantie', montant: 0, frequence: 'ponctuel', debut: debut, fin: '', montant_total: garantie, amortissement_annees: '', remarque: 'Date de versement à vérifier' });
  // mois passés
  const reels = [], recettes = [];
  // mois écoulés et mois en cours (le fichier de Mo est à jour) ; une case vide vaut 0
  g.mensuel.filter(function (l) { return l.mois <= auj; }).forEach(function (l) {
    [[idEngie, l.engie], [idVir, l.virements], [idDiv, l.depenses]].forEach(function (x) { if (x[1] != null && x[1] !== 0) reels.push({ mois: l.mois, id_charge: x[0], montant: x[1], remarque: '', source: 'GESTION COMPTABLE', modifie_le: new Date() }); });
    [['loyers', l.loyers], ['commissions', l.commissions], ['vente_shop', l.venteShop], ['adhesions', l.adhesions]].forEach(function (x) { recettes.push({ mois: l.mois, type: x[0], montant: x[1] || 0, remarque: '', source: 'GESTION COMPTABLE', modifie_le: new Date() }); });
  });
  if (apercu) return { ok: true, apercu: true, postes: nouveaux, reels: reels.length, recettes: recettes.length, mois: g.mensuel.map(function (l) { return l.mois; }) };
  const ecrire = function (t, lignes) {
    if (!lignes.length) return;
    const rows = lignes.map(function (v) { const r = new Array(t.nbCol).fill(''); Object.keys(v).forEach(function (k) { if (t.M[k] != null) r[t.M[k]] = v[k]; }); return r; });
    const debutL = t.lignes.length + 2;
    _assurerTaille(t.sh, debutL + rows.length - 1, t.nbCol);
    t.sh.getRange(debutL, 1, rows.length, t.nbCol).setValues(rows);
  };
  ecrire(tC, nouveaux); ecrire(tR, reels); ecrire(tS, recettes);
  PropertiesService.getScriptProperties().setProperty('DERNIER_ID_CH', String(Number(nouveaux[nouveaux.length - 1].id_charge.slice(2))));
  _journaliser('bilan_import', nouveaux.length + ' postes, ' + reels.length + ' montants réels, ' + recettes.length + ' recettes (GESTION COMPTABLE)');
  return { ok: true, postes: nouveaux.length, reels: reels.length, recettes: recettes.length };
}
