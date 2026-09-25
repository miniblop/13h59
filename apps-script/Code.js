/*************************************************************
 *  CAISSE 13h59 — enregistrement des ventes (backoffice v1)
 *  La caisse écrit des VALEURS dans l'onglet `ventes` (aucune formule) :
 *  remise, prix client, frais et prime sont calculés ici, à partir des
 *  onglets `remises` et `paiements`. Chaque vente référence son créateur
 *  par `id_createur` (clé de l'onglet `createurs`).
 *************************************************************/

// ===================== CONFIGURATION =======================

// Onglets du backoffice : une future table de base de données chacun.
const SHEET_VENTES    = 'ventes';
const SHEET_CREATEURS = 'createurs';
const SHEET_REMISES   = 'remises';
const SHEET_PAIEMENTS = 'paiements';
const SHEET_JOURNAL   = 'journal';
const SHEET_CATEGORIES   = 'categories';
const SHEET_STANDS       = 'stands';
const SHEET_EMPLACEMENTS = 'emplacements';

const COLONNES_EMPLACEMENTS = ['id_emplacement', 'id_createur', 'code_stand', 'debut', 'fin', 'preavis_recu_le', 'accueil_par', 'motif_fin', 'remarque'];

const SHEET_CANDIDATURES = 'candidatures';
const SHEET_EMAILS       = 'emails';
const COLONNES_CANDIDATURES = [
  'id_candidature', 'recue_le', 'maj_le', 'prenom', 'nom', 'marque', 'email', 'instagram', 'categorie',
  'stand_souhaite', 'description', 'statut', 'traitee_le', 'traitee_par', 'id_createur', 'source', 'remarque'
];

const COLONNES_VENTES = [
  'id_vente', 'id_panier', 'date', 'id_createur', 'reference', 'code_remise', 'code_paiement',
  'prix', 'remise', 'prix_client', 'frais', 'prime', 'id_transaction'
];

// ===========================================================

/** Menu ajouté à l'ouverture du classeur. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🛒 Caisse')
    .addItem('Ouvrir la caisse', 'ouvrirCaisse')
    .addToUi();
}

/** Ouvre l'interface caisse dans une fenêtre modale. */
function ouvrirCaisse() {
  const html = HtmlService.createHtmlOutputFromFile('Caisse')
    .setWidth(940)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Caisse 13h59');
}

/* ---------- Utilitaires ---------- */

function _norm(s) { return String(s == null ? '' : s).trim().toLowerCase(); }
function _round2(x) { return Math.round((Number(x) + Number.EPSILON) * 100) / 100; }

/** Mappe {nom_entete_normalisé : index 0-based} à partir de la ligne d'en-tête. */
function _headerMap(sheet) {
  const nbCol = sheet.getLastColumn();
  const entetes = sheet.getRange(1, 1, 1, nbCol).getValues()[0];
  const map = {};
  entetes.forEach(function (h, i) { map[_norm(h)] = i; });
  return { map: map, entetes: entetes, nbCol: nbCol };
}

/** Lit un onglet en objets {entête: valeur}. */
function _lireTable(sh) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(_norm);
  return sh.getRange(2, 1, n, h.length).getValues().map(function (r) {
    const o = {}; h.forEach(function (k, i) { if (k) o[k] = r[i]; }); return o;
  });
}

/** Agrandit l'onglet si besoin : écrire au-delà de sa dernière ligne ou colonne physique est refusé par Sheets.
 *  On ajoute une réserve pour ne pas le refaire à chaque vente. */
function _assurerTaille(sh, derniereLigne, derniereColonne) {
  const maxL = sh.getMaxRows(), maxC = sh.getMaxColumns();
  if (derniereLigne > maxL) sh.insertRowsAfter(maxL, derniereLigne - maxL + 500);
  if (derniereColonne && derniereColonne > maxC) sh.insertColumnsAfter(maxC, derniereColonne - maxC);
}

function _onglet(ss, nom) {
  const sh = ss.getSheetByName(nom);
  if (!sh) throw new Error("Onglet « " + nom + " » introuvable.");
  return sh;
}

/* ---------- Formats ---------- */

function _telephone(v) {
  let d = String(v == null ? '' : v).replace(/\D/g, '');
  if (d.length === 9) d = '0' + d;
  if (d.length === 11 && d.indexOf('33') === 0) d = '0' + d.slice(2);
  return d.length === 10 ? d.replace(/(\d{2})(?=\d)/g, '$1 ') : String(v || '').trim();
}
function _siretValide(s) {
  if (!/^\d{14}$/.test(s)) return false;
  let t = 0;
  for (let i = 0; i < 14; i++) { let n = +s[13 - i]; if (i % 2) { n *= 2; if (n > 9) n -= 9; } t += n; }
  return t % 10 === 0 || s.indexOf('356000000') === 0;   // La Poste : exception connue
}
/** « @nom », « nom » ou un lien instagram.com/nom?igsh=… → « nom » ; sinon null. */
function _compteInstagram(v) {
  const s = String(v == null ? '' : v).trim();
  const m = s.match(/instagram\.com\/([A-Za-z0-9._]+)/i);
  if (m) return m[1].toLowerCase();
  const h = s.replace(/^@/, '').trim();
  return /^[A-Za-z0-9._]{2,30}$/.test(h) && /[a-z]/i.test(h) ? h.toLowerCase() : null;
}
/** Catégories dans l'ordre choisi par l'équipe (colonne « ordre »). */
function _categoriesTriees(ss) {
  return _lireTable(_onglet(ss, SHEET_CATEGORIES)).filter(function (c) { return c['code']; })
    .sort(function (a, b) { return (Number(a['ordre']) || 999) - (Number(b['ordre']) || 999); });
}
function _lienInstagram(compte) { return 'https://www.instagram.com/' + compte + '/'; }
/** Clé de rapprochement : casse, espaces et caractères invisibles ignorés. */
function _cleNom(s) { return _nomPropre(s).toLowerCase(); }
/** Clé « large » : ignore aussi accents et ponctuation (pour signaler, jamais pour fusionner). */
function _cleLarge(s) { return _cleNom(s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]/g, ''); }

/* ---------- Référentiels ---------- */

/** Tous les créateurs : { parId: {id: {id, nom, actif}}, parCle: {nom normalisé: id} }. */
function _createurs(ss) {
  const parId = {}, parCle = {};
  _lireTable(_onglet(ss, SHEET_CREATEURS)).forEach(function (c) {
    const id = String(c['id_createur'] || '').trim();
    const nom = _nomPropre(c['nom']);
    if (!id || !nom) return;
    parId[id] = { id: id, nom: nom, actif: _norm(c['statut']) === 'actif' };
    parCle[nom.toLowerCase()] = id;
  });
  return { parId: parId, parCle: parCle };
}

/** Remises actives, par code : {code: {code, valeur, estPourcentage}} (valeur négative). */
function _remises(ss) {
  const out = {};
  _lireTable(_onglet(ss, SHEET_REMISES)).forEach(function (r) {
    const code = String(r['code'] || '').trim();
    if (!code || _norm(r['statut']) !== 'actif') return;
    out[code] = { code: code, valeur: Number(r['valeur']) || 0, estPourcentage: _norm(r['type']) === 'pourcentage' };
  });
  return out;
}

/** Moyens de paiement actifs, par code : {code: taux_frais}. */
function _paiements(ss) {
  const out = {};
  _lireTable(_onglet(ss, SHEET_PAIEMENTS)).forEach(function (p) {
    const code = _norm(p['code']);
    if (code && _norm(p['statut']) === 'actif') out[code] = Number(p['taux_frais']) || 0;
  });
  return out;
}

/** Montants d'une ligne : la remise est toujours à la charge du créateur. */
function _montants(prix, remise, tauxFrais) {
  const r = remise.estPourcentage ? _round2(prix * remise.valeur) : _round2(remise.valeur);
  const prixClient = _round2(prix - Math.abs(r));
  const frais = _round2(prixClient * tauxFrais);
  return { remise: r, prixClient: prixClient, frais: frais, prime: _round2(prixClient - frais) };
}

/* ---------- Données pour l'interface ---------- */

/** Renvoie tout ce dont l'interface caisse a besoin au chargement. */
function getCaisseData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cr = _createurs(ss);
  const createurs = Object.keys(cr.parId).map(function (id) { return cr.parId[id]; })
    .filter(function (c) { return c.actif; })
    .map(function (c) { return { id: c.id, nom: c.nom }; })
    .sort(function (a, b) { return a.nom.localeCompare(b.nom, 'fr', { sensitivity: 'base' }); });
  const rem = _remises(ss);
  const remises = Object.keys(rem).map(function (k) { return rem[k]; });

  const sh = _onglet(ss, SHEET_VENTES);
  const M = _headerMap(sh).map;
  const n = Math.max(0, sh.getLastRow() - 1);
  const max = function (col) {
    if (!n || M[col] == null) return 0;
    return sh.getRange(2, M[col] + 1, n, 1).getValues().reduce(function (m, r) { const x = Number(r[0]); return isNaN(x) ? m : Math.max(m, x); }, 0);
  };

  return {
    createurs: createurs,
    remises: remises,
    paiements: Object.keys(_paiements(ss)),
    // un numéro supprimé depuis Gestion ▸ Ventes n'est jamais réattribué
    prochainPanier: Math.max(max('id_panier'), _dernierId('DERNIER_ID_PANIER')) + 1,
    prochaineVente: Math.max(max('id_vente'), _dernierId('DERNIER_ID_VENTE')) + 1
  };
}

/* ---------- Anti-doublon (idempotence) ---------- */

// Durée pendant laquelle un identifiant de transaction reste mémorisé.
// 6 h = maximum autorisé par CacheService, largement assez pour un « réessayer ».
const ANTI_DOUBLON_TTL = 21600;

function _idTxValide(id) {
  return typeof id === 'string' && /^[A-Za-z0-9_-]{8,64}$/.test(id);
}

/**
 * Cette transaction a-t-elle déjà été écrite ? Renvoie le résultat d'origine, ou null.
 * 1) CacheService (rapide, mais « best effort ») ; 2) colonne id_transaction (filet durable).
 */
function _transactionDejaEcrite(sh, M, idTx) {
  const enCache = CacheService.getScriptCache().get('tx_' + idTx);
  if (enCache) return JSON.parse(enCache);
  const last = sh.getLastRow();
  if (last < 2) return null;
  const n = Math.min(500, last - 1);            // une retentative arrive dans la minute
  const vals = sh.getRange(last - n + 1, 1, n, sh.getLastColumn()).getValues();
  const lignes = vals.filter(function (r) { return String(r[M['id_transaction']]) === idTx; });
  if (!lignes.length) return null;
  return {
    ok: true,
    panier: lignes[0][M['id_panier']],
    ventes: lignes.map(function (r) { return r[M['id_vente']]; }),
    nbLignes: lignes.length
  };
}

/* ---------- Enregistrement d'une vente ---------- */

/**
 * Enregistre un panier dans l'onglet `ventes`.
 * @param {Object} data  { idTransaction, paiement, lignes:[{idCreateur, reference, prix, typeRemise}] }
 *   idTransaction : identifiant unique généré par le site pour CE panier. Si la même
 *   transaction arrive deux fois (réponse perdue puis « réessayer »), la seconde
 *   n'écrit rien et renvoie le résultat de la première (deja:true).
 * @return {Object} { ok, panier, ventes:[id_vente], nbLignes, deja }
 */
function enregistrerVente(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000); // deux saisies simultanées ne peuvent pas prendre les mêmes numéros
  try {
    if (!data || !data.lignes || !data.lignes.length) throw new Error('Panier vide.');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = _onglet(ss, SHEET_VENTES);
    const h = _headerMap(sh);
    const M = h.map;
    COLONNES_VENTES.forEach(function (c) { if (M[c] == null) throw new Error("Colonne « " + c + " » absente de l'onglet « " + SHEET_VENTES + " »."); });

    const idTx = _idTxValide(data.idTransaction) ? data.idTransaction : null;
    if (idTx) {
      const deja = _transactionDejaEcrite(sh, M, idTx);
      if (deja) { deja.deja = true; return deja; }
    }

    const cr = _createurs(ss);
    const remises = _remises(ss);
    const paiements = _paiements(ss);
    const paiement = _norm(data.paiement);
    if (!(paiement in paiements)) throw new Error('Moyen de paiement inconnu : « ' + data.paiement + ' ».');

    const info = getCaisseData();
    const panier = info.prochainPanier;
    let vente = info.prochaineVente;
    const maintenant = new Date();

    const lignes = data.lignes.map(function (l) {
      // Une page de caisse ouverte avant la mise à jour envoie le NOM (createur, ou vendeur avant le renommage).
      const nom = l.createur != null ? l.createur : l.vendeur;
      const id = (l.idCreateur && cr.parId[l.idCreateur]) ? l.idCreateur : cr.parCle[_nomPropre(nom).toLowerCase()];
      if (!id) throw new Error('Créateur inconnu : « ' + (l.idCreateur || nom || '?') + ' ».');
      const prix = Number(l.prix);
      if (!(prix > 0)) throw new Error('Prix invalide pour ' + cr.parId[id].nom + '.');
      const code = String(l.typeRemise || 'pas_de_remise').trim();
      const remise = remises[code];
      if (!remise) throw new Error('Remise inconnue ou inactive : « ' + code + ' ».');
      const m = _montants(prix, remise, paiements[paiement]);
      const valeurs = {
        id_vente: vente++, id_panier: panier, date: maintenant, id_createur: id,
        reference: String(l.reference || ''), code_remise: code, code_paiement: paiement,
        prix: prix, remise: m.remise, prix_client: m.prixClient, frais: m.frais, prime: m.prime,
        id_transaction: idTx || ''
      };
      const ligne = new Array(h.nbCol).fill('');
      COLONNES_VENTES.forEach(function (c) { ligne[M[c]] = valeurs[c]; });
      return ligne;
    });

    const derniere = sh.getLastRow();
    _assurerTaille(sh, derniere + lignes.length, h.nbCol);
    const bloc = sh.getRange(derniere + 1, 1, lignes.length, h.nbCol);
    if (derniere >= 2) sh.getRange(derniere, 1, 1, h.nbCol).copyTo(bloc, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    // Référence en texte brut : sinon Sheets convertit « 20-2 » ou « 3/4 » en date.
    sh.getRange(derniere + 1, M['reference'] + 1, lignes.length, 1).setNumberFormat('@');
    bloc.setValues(lignes);
    SpreadsheetApp.flush();

    const res = { ok: true, panier: panier, ventes: lignes.map(function (l) { return l[M['id_vente']]; }), nbLignes: lignes.length, deja: false };
    if (idTx) CacheService.getScriptCache().put('tx_' + idTx, JSON.stringify(res), ANTI_DOUBLON_TTL);
    return res;

  } catch (e) {
    return { ok: false, message: String(e && e.message ? e.message : e) };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Récap des ventes d'une journée (onglet Caisse du site) ---------- */

/**
 * Ventes d'un jour donné (lecture seule), avec le nom du créateur.
 * @param {string} jourStr  'yyyy-MM-dd' ; vide/invalide = aujourd'hui (fuseau du classeur).
 */
function getVentesDuJour(jourStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const jour = /^\d{4}-\d{2}-\d{2}$/.test(String(jourStr || '')) ? jourStr : Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const noms = _createurs(ss).parId;
  const num = function (v) { const n = Number(v); return isNaN(n) ? 0 : n; };
  const lignes = [];
  _lireTable(_onglet(ss, SHEET_VENTES)).forEach(function (v) {
    const d = v['date'];
    if (!(d instanceof Date) || Utilities.formatDate(d, tz, 'yyyy-MM-dd') !== jour) return;
    const heure = Utilities.formatDate(d, tz, 'HH:mm');
    const id = String(v['id_createur'] || '');
    lignes.push({
      panier: v['id_panier'] === '' || v['id_panier'] == null ? null : num(v['id_panier']),
      vente: num(v['id_vente']),
      heure: heure === '00:00' ? '' : heure,   // 00:00 = saisie manuelle sans heure
      idCreateur: id,
      createur: noms[id] ? noms[id].nom : id,
      reference: String(v['reference'] || ''),
      remiseType: String(v['code_remise'] || '').trim(),
      paiement: String(v['code_paiement'] || '').trim(),
      prix: num(v['prix']),
      remise: num(v['remise']),
      prixClient: num(v['prix_client']),
      sens: _norm(v['type_ligne']) === 'annulation' ? -1 : 1
    });
  });
  return { jour: jour, lignes: lignes };
}
