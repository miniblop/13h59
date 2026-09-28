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
