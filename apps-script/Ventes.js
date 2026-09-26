/*************************************************************
 *  VENTES — Gestion ▸ Ventes : consulter, supprimer, annuler, corriger
 *  Règle : un mois passé (déjà facturé et versé) n'est jamais réécrit.
 *   - Vente du MOIS EN COURS : « supprimer » efface la ligne, « corriger »
 *     la modifie sur place. Le journal garde une copie complète de l'avant.
 *   - Vente d'un MOIS PASSÉ : « annuler » ajoute une ligne « annulation »
 *     aux montants opposés (même créateur, même panier) datée du jour ;
 *     « corriger » = annuler + ajouter une ligne « correction » juste.
 *  Un numéro de vente ou de panier supprimé n'est jamais réattribué
 *  (propriétés DERNIER_ID_VENTE / DERNIER_ID_PANIER, lues par la caisse).
 *  Colonnes ajoutées à `ventes` à la première régularisation : type_ligne
 *  (vide = vente, « annulation », « correction »), vente_origine, motif.
 *************************************************************/

const COLONNES_CORRECTION = ['type_ligne', 'vente_origine', 'motif'];

/** Ajoute les colonnes de correction à `ventes` si elles manquent ; renvoie l'en-tête à jour. */
function _colonnesCorrection(sh) {
  const h = _headerMap(sh);
  const manquantes = COLONNES_CORRECTION.filter(function (c) { return h.map[c] == null; });
  if (!manquantes.length) return h;
  _assurerTaille(sh, 1, h.nbCol + manquantes.length);
  sh.getRange(1, h.nbCol + 1, 1, manquantes.length).setValues([manquantes]).setFontWeight('bold');
  return _headerMap(sh);
}

function _typeLigne(v) { const t = _norm(v); return t === 'annulation' || t === 'correction' ? t : 'vente'; }

/** Toutes les ventes, en objets lisibles par le site. */
function _ventesDetaillees(ss) {
  const noms = _createurs(ss).parId, moisCourant = _jourIso(new Date()).slice(0, 7);
  const num = function (v) { const n = Number(v); return isNaN(n) ? 0 : n; };
  const lignes = [];
  _lireTable(_onglet(ss, SHEET_VENTES)).forEach(function (v) {
    if (!(v['date'] instanceof Date) || v['id_vente'] === '' || v['id_vente'] == null) return;
    const id = String(v['id_createur'] || ''), heure = _heure(v['date']), jour = _jourIso(v['date']);
    lignes.push({
      vente: num(v['id_vente']), panier: v['id_panier'] === '' ? null : num(v['id_panier']),
      date: jour, heure: heure === '00:00' ? '' : heure,
      idCreateur: id, createur: noms[id] ? noms[id].nom : id,
      reference: String(v['reference'] || ''), remise: String(v['code_remise'] || ''), paiement: String(v['code_paiement'] || ''),
      prix: num(v['prix']), montantRemise: num(v['remise']), prixClient: num(v['prix_client']), frais: num(v['frais']), prime: num(v['prime']),
      type: _typeLigne(v['type_ligne']), origine: v['vente_origine'] === '' || v['vente_origine'] == null ? null : num(v['vente_origine']),
      motif: String(v['motif'] || ''),
      moisOuvert: jour.slice(0, 7) === moisCourant
    });
  });
  // Qui a annulé / corrigé quoi.
  const parId = {};
  lignes.forEach(function (l) { parId[l.vente] = l; });
  lignes.forEach(function (l) {
    if (l.origine == null || !parId[l.origine]) return;
    if (l.type === 'annulation') parId[l.origine].annuleePar = l.vente;
    if (l.type === 'correction') parId[l.origine].corrigeePar = l.vente;
  });
  return lignes;
}

function _gVentes(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const du = /^\d{4}-\d{2}-\d{2}$/.test(String(body.du || '')) ? body.du : _iso(new Date(Date.now() - 6 * 864e5));
  const au = /^\d{4}-\d{2}-\d{2}$/.test(String(body.au || '')) ? body.au : _iso(new Date());
  const q = _norm(body.q), idCreateur = String(body.idCreateur || '');
  const toutes = _ventesDetaillees(ss);
  // Recherche par numéro de vente ou de panier : sur tout l'historique, sans tenir compte des dates.
  const numero = /^#?\d+$/.test(q) ? Number(q.replace('#', '')) : null;
  let lignes = toutes.filter(function (l) {
    if (numero != null) return l.vente === numero || l.panier === numero || l.origine === numero;
    if (l.date < du || l.date > au) return false;
    if (idCreateur && l.idCreateur !== idCreateur) return false;
    return !q || (l.reference + ' ' + l.createur + ' ' + l.motif).toLowerCase().indexOf(q) !== -1;
  }).sort(function (a, b) { return b.vente - a.vente; });
  const trop = lignes.length > 1500;
  if (trop) lignes = lignes.slice(0, 1500);

  const cr = _createurs(ss).parId;
  const remises = _lireTable(_onglet(ss, SHEET_REMISES)).filter(function (r) { return r['code']; }).map(function (r) {
    return { code: String(r['code']), valeur: Number(r['valeur']) || 0, estPourcentage: _norm(r['type']) === 'pourcentage', actif: _norm(r['statut']) === 'actif' };
  });
  const paiements = _lireTable(_onglet(ss, SHEET_PAIEMENTS)).filter(function (p) { return p['code']; }).map(function (p) {
    return { code: _norm(p['code']), taux: Number(p['taux_frais']) || 0, actif: _norm(p['statut']) === 'actif' };
  });
  return {
    ok: true, du: du, au: au, lignes: lignes, tronque: trop,
    createurs: Object.keys(cr).map(function (k) { return cr[k]; }).sort(function (a, b) { return a.nom.localeCompare(b.nom, 'fr', { sensitivity: 'base' }); }),
    remises: remises, paiements: paiements
  };
}

/* ---------- Écriture ---------- */

function _motif(body) {
  const m = _textePublic(body.motif, 300);
  if (m.replace(/^'/, '').length < 3) throw new Error('Indique le motif (ex. « mauvais créateur », « vente saisie deux fois »).');
  return m + ' (' + _signataire() + ')';
}

/** Lignes brutes de `ventes`. */
function _lignesVentes(sh, nb) {
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, nb).getValues() : [];
}

/** Ajoute des lignes (objets {colonne: valeur}) à la fin de `ventes`, au format des ventes de la caisse. */
function _ajouterVentes(sh, h, objets) {
  const lignes = objets.map(function (o) {
    const l = new Array(h.nbCol).fill('');
    Object.keys(o).forEach(function (k) { if (h.map[k] != null) l[h.map[k]] = o[k]; });
    return l;
  });
  const derniere = sh.getLastRow();
  _assurerTaille(sh, derniere + lignes.length, h.nbCol);
  const bloc = sh.getRange(derniere + 1, 1, lignes.length, h.nbCol);
  if (derniere >= 2) sh.getRange(derniere, 1, 1, h.nbCol).copyTo(bloc, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  sh.getRange(derniere + 1, h.map['reference'] + 1, lignes.length, 1).setNumberFormat('@');
  bloc.setValues(lignes);
}

/** Contexte commun : feuille, en-tête (colonnes ajoutées), lignes, index des ventes déjà annulées, prochain numéro. */
function _contexteVentes(ss) {
  const sh = _onglet(ss, SHEET_VENTES);
  const h = _colonnesCorrection(sh);
  const M = h.map, lignes = _lignesVentes(sh, h.nbCol);
  const annulees = {};
  let max = 0, maxPanier = 0;
  lignes.forEach(function (r) {
    const id = Number(r[M['id_vente']]), p = Number(r[M['id_panier']]);
    if (!isNaN(id)) max = Math.max(max, id);
    if (!isNaN(p)) maxPanier = Math.max(maxPanier, p);
    if (_typeLigne(r[M['type_ligne']]) === 'annulation') annulees[Number(r[M['vente_origine']])] = true;
  });
  return { ss: ss, sh: sh, h: h, M: M, lignes: lignes, annulees: annulees, max: max, maxPanier: maxPanier, prochain: Math.max(max, _dernierId('DERNIER_ID_VENTE')) + 1 };
}

/** Plus grand numéro déjà attribué puis supprimé (0 si aucun). */
function _dernierId(cle) { return Number(PropertiesService.getScriptProperties().getProperty(cle)) || 0; }
function _retenirDerniersIds(ctx) {
  const p = PropertiesService.getScriptProperties();
  p.setProperty('DERNIER_ID_VENTE', String(Math.max(ctx.max, _dernierId('DERNIER_ID_VENTE'))));
  p.setProperty('DERNIER_ID_PANIER', String(Math.max(ctx.maxPanier, _dernierId('DERNIER_ID_PANIER'))));
}

/** Le mois de cette date est-il le mois en cours (pas encore facturé) ? */
function _moisOuvert(ss, d) {
  return d instanceof Date && _jourIso(d).slice(0, 7) === _jourIso(new Date()).slice(0, 7);
}
/** Modifiable sur place : mois en cours, pas une annulation, pas déjà annulée. */
function _modifiableSurPlace(ctx, r) {
  return _moisOuvert(ctx.ss, r[ctx.M['date']]) && _typeLigne(r[ctx.M['type_ligne']]) !== 'annulation' && !ctx.annulees[Number(r[ctx.M['id_vente']])];
}
/** Copie lisible d'une ligne pour le journal (rien ne se perd, même supprimé). */
function _copieVente(ctx, r) {
  const M = ctx.M;
  return 'n°' + r[M['id_vente']] + ' panier ' + r[M['id_panier']] + ' du ' + _iso(r[M['date']]) + ' · ' + r[M['id_createur']] + ' · « ' + r[M['reference']] + ' » · ' +
    r[M['prix']] + ' € ' + r[M['code_remise']] + ' · ' + r[M['code_paiement']] + ' · client ' + r[M['prix_client']] + ' € · prime ' + r[M['prime']] + ' €';
}
/** Supprime des lignes de `ventes` (numéros jamais réattribués). */
function _supprimerLignes(ctx, lignes) {
  _retenirDerniersIds(ctx);
  lignes.map(function (r) { return ctx.lignes.indexOf(r) + 2; }).sort(function (a, b) { return b - a; })
    .forEach(function (n) { ctx.sh.deleteRow(n); });
}

function _gVenteSupprimer(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), ctx = _contexteVentes(ss);
  const motif = _motif(body), r = _venteAnnulable(ctx, body.idVente);
  if (!_modifiableSurPlace(ctx, r)) throw new Error("La vente n°" + body.idVente + " date d'un mois passé (déjà facturé) : annule-la plutôt.");
  if (ctx.lignes.some(function (x) { return Number(x[ctx.M['vente_origine']]) === Number(body.idVente); })) throw new Error('Cette vente a déjà été corrigée : supprime plutôt la ligne de correction.');
  const copie = _copieVente(ctx, r);
  _supprimerLignes(ctx, [r]);
  _journaliser('vente_supprimee', copie + ' : ' + motif);
  return { ok: true };
}

function _gPanierSupprimer(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), ctx = _contexteVentes(ss), M = ctx.M;
  const motif = _motif(body), panier = Number(body.idPanier);
  const lignes = ctx.lignes.filter(function (r) { return Number(r[M['id_panier']]) === panier; });
  if (!lignes.length) throw new Error('Panier n°' + body.idPanier + ' introuvable.');
  if (!lignes.every(function (r) { return _modifiableSurPlace(ctx, r) && !r[M['vente_origine']]; }))
    throw new Error('Le panier n°' + panier + " contient des ventes d'un mois passé ou déjà régularisées : annule-le plutôt.");
  const copies = lignes.map(function (r) { return _copieVente(ctx, r); });
  _supprimerLignes(ctx, lignes);
  _journaliser('panier_supprime', 'panier n°' + panier + ' (' + lignes.length + ' vente(s)) : ' + motif + ' | ' + copies.join(' | '));
  return { ok: true, nb: lignes.length };
}

/** Vérifie qu'une vente peut être annulée et renvoie sa ligne brute. */
function _venteAnnulable(ctx, idVente) {
  const M = ctx.M, id = Number(idVente);
  const r = ctx.lignes.filter(function (x) { return Number(x[M['id_vente']]) === id; })[0];
  if (!r) throw new Error('Vente n°' + idVente + ' introuvable.');
  if (_typeLigne(r[M['type_ligne']]) === 'annulation') throw new Error("La ligne n°" + id + " est déjà une annulation : on ne l'annule pas.");
  if (ctx.annulees[id]) throw new Error('La vente n°' + id + ' est déjà annulée.');
  return r;
}

/** Ligne d'annulation : montants opposés, même créateur, même panier. */
function _ligneAnnulation(ctx, r, motif, maintenant) {
  const M = ctx.M, neg = function (k) { const n = Number(r[M[k]]) || 0; return n ? -n : 0; };
  return {
    id_vente: ctx.prochain++, id_panier: r[M['id_panier']], date: maintenant, id_createur: r[M['id_createur']],
    reference: 'Annulation de la vente n°' + r[M['id_vente']], code_remise: r[M['code_remise']], code_paiement: r[M['code_paiement']],
    prix: neg('prix'), remise: neg('remise'), prix_client: neg('prix_client'), frais: neg('frais'), prime: neg('prime'),
    id_transaction: '', type_ligne: 'annulation', vente_origine: r[M['id_vente']], motif: motif
  };
}

function _gVenteAnnuler(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), ctx = _contexteVentes(ss);
  const motif = _motif(body), r = _venteAnnulable(ctx, body.idVente);
  const a = _ligneAnnulation(ctx, r, motif, new Date());
  _ajouterVentes(ctx.sh, ctx.h, [a]);
  _journaliser('vente_annulee', 'n°' + body.idVente + ' (' + r[ctx.M['id_createur']] + ', ' + r[ctx.M['prix_client']] + ' €) : ' + motif);
  return { ok: true, annulation: a.id_vente };
}

function _gPanierAnnuler(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), ctx = _contexteVentes(ss), M = ctx.M;
  const motif = _motif(body), panier = Number(body.idPanier), maintenant = new Date();
  const aAnnuler = ctx.lignes.filter(function (r) {
    return Number(r[M['id_panier']]) === panier && _typeLigne(r[M['type_ligne']]) !== 'annulation' && !ctx.annulees[Number(r[M['id_vente']])];
  });
  if (!aAnnuler.length) throw new Error('Rien à annuler dans le panier n°' + body.idPanier + '.');
  const lignes = aAnnuler.map(function (r) { return _ligneAnnulation(ctx, r, motif, maintenant); });
  _ajouterVentes(ctx.sh, ctx.h, lignes);
  _journaliser('panier_annule', 'panier n°' + panier + ' : ' + aAnnuler.length + ' vente(s) (' + aAnnuler.map(function (r) { return 'n°' + r[M['id_vente']]; }).join(', ') + ') : ' + motif);
  return { ok: true, nb: lignes.length };
}

function _gVenteCorriger(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), ctx = _contexteVentes(ss), M = ctx.M;
  const motif = _motif(body), r = _venteAnnulable(ctx, body.idVente);

  const cr = _createurs(ss).parId;
  const idCreateur = String(body.idCreateur || r[M['id_createur']]);
  if (!cr[idCreateur]) throw new Error('Créateur inconnu : « ' + idCreateur + ' ».');
  const prix = _round2(Number(String(body.prix == null ? r[M['prix']] : body.prix).replace(',', '.')));
  if (!(prix > 0)) throw new Error('Prix invalide.');
  const codeRemise = String(body.codeRemise || r[M['code_remise']] || REMISE_PAR_DEFAUT);
  const rem = _lireTable(_onglet(ss, SHEET_REMISES)).filter(function (x) { return String(x['code']) === codeRemise; })[0];
  if (!rem) throw new Error('Remise inconnue : « ' + codeRemise + ' ».');
  const codePaiement = _norm(body.codePaiement || r[M['code_paiement']]);
  const pai = _lireTable(_onglet(ss, SHEET_PAIEMENTS)).filter(function (x) { return _norm(x['code']) === codePaiement; })[0];
  if (!pai) throw new Error('Moyen de paiement inconnu : « ' + codePaiement + ' ».');
  const reference = body.reference == null ? String(r[M['reference']] || '') : _textePublic(body.reference, 80);

  const avant = [r[M['id_createur']], Number(r[M['prix']]), String(r[M['code_remise']]), _norm(r[M['code_paiement']]), String(r[M['reference']] || '')].join('|');
  if (avant === [idCreateur, prix, codeRemise, codePaiement, reference].join('|')) throw new Error("Rien n'a changé : modifie au moins un champ.");

  const m = _montants(prix, { valeur: Number(rem['valeur']) || 0, estPourcentage: _norm(rem['type']) === 'pourcentage' }, Number(pai['taux_frais']) || 0);
  // Même paiement, même montant payé : les frais d'origine sont gardés (le taux a pu changer depuis la vente).
  if (codePaiement === _norm(r[M['code_paiement']]) && m.prixClient === Number(r[M['prix_client']])) {
    m.frais = Number(r[M['frais']]) || 0; m.prime = _round2(m.prixClient - m.frais);
  }
  if (_modifiableSurPlace(ctx, r)) {
    // Mois en cours : la ligne est modifiée, le journal garde l'avant.
    const copie = _copieVente(ctx, r), i = ctx.lignes.indexOf(r);
    const nv = { id_createur: idCreateur, reference: reference, code_remise: codeRemise, code_paiement: codePaiement, prix: prix,
                 remise: m.remise, prix_client: m.prixClient, frais: m.frais, prime: m.prime };
    Object.keys(nv).forEach(function (k) { r[M[k]] = nv[k]; });
    ctx.sh.getRange(i + 2, M['reference'] + 1).setNumberFormat('@');
    ctx.sh.getRange(i + 2, 1, 1, r.length).setValues([r]);
    _journaliser('vente_corrigee', 'sur place, avant : ' + copie + ' → après : ' + _copieVente(ctx, r) + ' : ' + motif);
    return { ok: true, surPlace: true, correction: Number(body.idVente), prixClient: m.prixClient, prime: m.prime };
  }
  const maintenant = new Date();
  const a = _ligneAnnulation(ctx, r, motif, maintenant);
  const c = {
    id_vente: ctx.prochain++, id_panier: r[M['id_panier']], date: maintenant, id_createur: idCreateur, reference: reference,
    code_remise: codeRemise, code_paiement: codePaiement, prix: prix, remise: m.remise, prix_client: m.prixClient, frais: m.frais, prime: m.prime,
    id_transaction: '', type_ligne: 'correction', vente_origine: r[M['id_vente']], motif: motif
  };
  _ajouterVentes(ctx.sh, ctx.h, [a, c]);
  const change = [];
  if (idCreateur !== String(r[M['id_createur']])) change.push('créateur ' + r[M['id_createur']] + ' → ' + idCreateur);
  if (prix !== Number(r[M['prix']])) change.push('prix ' + r[M['prix']] + ' → ' + prix);
  if (codeRemise !== String(r[M['code_remise']])) change.push('remise ' + r[M['code_remise']] + ' → ' + codeRemise);
  if (codePaiement !== _norm(r[M['code_paiement']])) change.push('paiement ' + r[M['code_paiement']] + ' → ' + codePaiement);
  if (reference !== String(r[M['reference']] || '')) change.push('référence');
  _journaliser('vente_corrigee', 'n°' + body.idVente + ' → n°' + c.id_vente + ' (' + change.join(', ') + ') : ' + motif);
  return { ok: true, surPlace: false, annulation: a.id_vente, correction: c.id_vente, prixClient: m.prixClient, prime: m.prime };
}

/* ---------- Journal ---------- */

function _gJournal(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = _onglet(ss, SHEET_JOURNAL), n = sh.getLastRow() - 1;
  const limite = Math.min(Math.max(Number(body.limite) || 1000, 50), 5000);
  if (n < 1) return { ok: true, lignes: [], total: 0 };
  const debut = Math.max(2, n + 2 - limite);
  const lignes = sh.getRange(debut, 1, n + 2 - debut, 4).getValues().map(function (r) {
    return { date: r[0] instanceof Date ? _jourIso(r[0]) + ' ' + _heure(r[0]) : String(r[0]), qui: String(r[1] || ''), action: String(r[2] || ''), detail: String(r[3] || '') };
  }).reverse();
  return { ok: true, lignes: lignes, total: n };
}
