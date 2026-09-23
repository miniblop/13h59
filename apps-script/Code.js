/*************************************************************
 *  CAISSE 13h59 — système d'encaissement
 *  À coller dans : Extensions ▸ Apps Script ▸ fichier Code.gs
 *
 *  Principe : la caisse écrit les colonnes de SAISIE dans l'onglet
 *  "ventes" et RECOPIE tes formules existantes vers le bas pour les
 *  colonnes calculées (remise, prix_client, taxe, prime_vendeur).
 *  => c'est ta propre logique de calcul qui s'applique.
 *  Un calcul de secours est prévu si une colonne n'est pas une formule.
 *************************************************************/

// ===================== CONFIGURATION =======================

const SHEET_VENTES   = 'ventes';
const SHEET_VENDEURS = 'vendeurs';
const SHEET_REMISES  = 'remises';

// Remises PRISES EN CHARGE PAR LE MAGASIN
// (le vendeur touche sa prime sur le prix PLEIN, pas sur le prix client).
// >>> Vérifie / complète cette liste si besoin. <<<
const REMISES_MAGASIN = ['machine_cadeau_5€', 'machine_cadeau_10%', 'machine_cadeau_20%'];

// Taux de taxe par type de paiement — SECOURS uniquement
// (utilisé seulement si la colonne "taxe" n'est PAS une formule).
const TAUX_TAXE = { 'cb': 0.0175, 'espèces': 0, 'especes': 0 };

// Types de paiement proposés dans la caisse.
const TYPES_PAIEMENT = ['cb', 'espèces'];

// Colonne "validation" : la caisse n'y touche pas par défaut
// (nouvelle vente = non validée). Passe à true pour cocher automatiquement.
const COCHER_VALIDATION = false;

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

/* ---------- Données pour l'interface ---------- */

/** Renvoie tout ce dont l'interface a besoin au chargement. */
function getCaisseData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Vendeurs actifs ---
  const shV = ss.getSheetByName(SHEET_VENDEURS);
  const hV = _headerMap(shV);
  const cNomV = hV.map['nom_vendeur'];
  const cStatV = hV.map['status'];
  const dataV = shV.getRange(2, 1, Math.max(0, shV.getLastRow() - 1), hV.nbCol).getValues();
  const vendeurs = dataV
    .filter(function (r) { return _norm(r[cStatV]) === 'actif' && String(r[cNomV]).trim() !== ''; })
    .map(function (r) { return String(r[cNomV]).trim(); });
  // dédoublonnage + tri
  const vendeursUniques = Array.from(new Set(vendeurs)).sort(function (a, b) {
    return a.localeCompare(b, 'fr', { sensitivity: 'base' });
  });

  // --- Remises actives ---
  const shR = ss.getSheetByName(SHEET_REMISES);
  const hR = _headerMap(shR);
  const cNomR = hR.map['type_de_remise'];
  const cValR = hR.map['valeur'];
  const cTypeR = hR.map['type'];
  const cStatR = hR.map['status'];
  const dataR = shR.getRange(2, 1, Math.max(0, shR.getLastRow() - 1), hR.nbCol).getValues();
  const remises = dataR
    .filter(function (r) { return _norm(r[cStatR]) === 'actif' && String(r[cNomR]).trim() !== ''; })
    .map(function (r) {
      const nom = String(r[cNomR]).trim();
      return {
        nom: nom,
        valeur: Number(r[cValR]) || 0,            // ex : -0.10 (%) ou -5 (€)
        estPourcentage: _norm(r[cTypeR]) === 'pourcentage',
        magasin: REMISES_MAGASIN.indexOf(nom) !== -1
      };
    });

  // --- Prochains numéros (depuis l'onglet ventes) ---
  const shVe = ss.getSheetByName(SHEET_VENTES);
  const hVe = _headerMap(shVe);
  const nbLignes = Math.max(0, shVe.getLastRow() - 1);
  let maxPanier = 0, maxVente = 0;
  if (nbLignes > 0) {
    const colP = shVe.getRange(2, hVe.map['numero_panier'] + 1, nbLignes, 1).getValues();
    const colVn = shVe.getRange(2, hVe.map['numero_vente'] + 1, nbLignes, 1).getValues();
    colP.forEach(function (r) { const n = Number(r[0]); if (!isNaN(n)) maxPanier = Math.max(maxPanier, n); });
    colVn.forEach(function (r) { const n = Number(r[0]); if (!isNaN(n)) maxVente = Math.max(maxVente, n); });
  }

  return {
    vendeurs: vendeursUniques,
    remises: remises,
    paiements: TYPES_PAIEMENT,
    prochainPanier: maxPanier + 1,
    prochaineVente: maxVente + 1
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
 * 1) CacheService (rapide, mais « best effort » : Google peut l'effacer plus tôt) ;
 * 2) colonne « id_transaction » de l'onglet ventes, SI elle existe (filet durable).
 */
function _transactionDejaEcrite(sh, M, idTx) {
  const enCache = CacheService.getScriptCache().get('tx_' + idTx);
  if (enCache) return JSON.parse(enCache);

  const col = M['id_transaction'];
  if (col == null) return null;
  const last = sh.getLastRow();
  if (last < 2) return null;
  const n = Math.min(500, last - 1);            // une retentative arrive dans la minute
  const vals = sh.getRange(last - n + 1, 1, n, sh.getLastColumn()).getValues();
  const lignes = vals.filter(function (r) { return String(r[col]) === idTx; });
  if (!lignes.length) return null;
  return {
    ok: true,
    panier: lignes[0][M['numero_panier']],
    ventes: lignes.map(function (r) { return r[M['numero_vente']]; }),
    nbLignes: lignes.length
  };
}

/* ---------- Enregistrement d'une vente ---------- */

/**
 * Enregistre un panier dans l'onglet "ventes".
 * @param {Object} data  { idTransaction, paiement:'cb'|'espèces', lignes:[{vendeur,reference,prix,typeRemise}] }
 *   idTransaction : identifiant unique généré par le site pour CE panier. Si la même
 *   transaction arrive deux fois (clic répété, réponse perdue puis « réessayer »),
 *   la seconde ne réécrit rien et renvoie le résultat de la première (deja:true).
 * @return {Object} { ok, panier, ventes:[numeros], nbLignes, deja }
 */
function enregistrerVente(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000); // évite les collisions de numéros si 2 saisies simultanées
  try {
    if (!data || !data.lignes || !data.lignes.length) {
      throw new Error('Panier vide.');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_VENTES);
    const h = _headerMap(sh);
    const M = h.map;
    const nbCol = h.nbCol;

    // --- Anti-doublon : vérifié DANS le verrou, donc deux envois simultanés
    //     du même panier sont forcément traités l'un après l'autre.
    const idTx = _idTxValide(data.idTransaction) ? data.idTransaction : null;
    if (idTx) {
      const deja = _transactionDejaEcrite(sh, M, idTx);
      if (deja) { deja.deja = true; return deja; }
    }

    // Recalcul autoritaire des numéros (dans le verrou).
    const info = getCaisseData();
    const panier = info.prochainPanier;
    let vente = info.prochaineVente;

    // Table des remises (pour le calcul de secours).
    const remiseParNom = {};
    info.remises.forEach(function (r) { remiseParNom[r.nom] = r; });

    // Ligne modèle = dernière ligne de données (pour recopier formules + format).
    const derniereLigne = sh.getLastRow();
    const formulesR1C1 = derniereLigne >= 2
      ? sh.getRange(derniereLigne, 1, 1, nbCol).getFormulasR1C1()[0]
      : new Array(nbCol).fill('');

    const nb = data.lignes.length;
    const premiere = derniereLigne + 1;
    const bloc = sh.getRange(premiere, 1, nb, nbCol);

    // 1) FORMAT + VALIDATIONS du modèle, en une fois pour tout le panier
    //    (avant : 2 copies + N appels PAR ligne → exécution lente → délais réseau).
    if (derniereLigne >= 2) {
      const src = sh.getRange(derniereLigne, 1, 1, nbCol);
      src.copyTo(bloc, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      src.copyTo(bloc, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);

      // Assouplit les validations « Refuser la saisie » (menu conservé, valeur hors
      // liste acceptée) — sinon la caisse plante sur un nom absent de la liste.
      const dvs = src.getDataValidations()[0];
      for (var c = 0; c < nbCol; c++) {
        var dv = dvs[c];
        if (dv && dv.getAllowInvalid && !dv.getAllowInvalid()) {
          sh.getRange(premiere, c + 1, nb, 1).setDataValidation(dv.copy().setAllowInvalid(true).build());
        }
      }
    }

    // Référence en TEXTE brut : sinon Sheets convertit « 20-2 » ou « 3/4 » en date.
    if (M['reference_produit'] != null) {
      sh.getRange(premiere, M['reference_produit'] + 1, nb, 1).setNumberFormat('@');
    }

    const aFormule = function (nomCol) {
      const idx = M[nomCol];
      return idx != null && formulesR1C1[idx] && formulesR1C1[idx] !== '';
    };

    // 2) Valeurs de saisie (+ calcul de secours si une colonne n'est pas une formule).
    const numerosVente = [];
    const maintenant = new Date();
    const tableau = data.lignes.map(function (ligne) {
      const valeurs = new Array(nbCol).fill('');
      const prix = Number(ligne.prix) || 0;
      const typeRemise = ligne.typeRemise || 'pas_de_remise';

      if (M['numero_panier'] != null)     valeurs[M['numero_panier']] = panier;
      if (M['numero_vente'] != null)      valeurs[M['numero_vente']] = vente;
      if (M['date'] != null)              valeurs[M['date']] = maintenant;
      if (M['vendeur'] != null)           valeurs[M['vendeur']] = ligne.vendeur;
      if (M['reference_produit'] != null) valeurs[M['reference_produit']] = String(ligne.reference || '');
      if (M['type_de_remise'] != null)    valeurs[M['type_de_remise']] = typeRemise;
      if (M['type_de_paiement'] != null)  valeurs[M['type_de_paiement']] = data.paiement;
      if (M['prix'] != null)              valeurs[M['prix']] = prix;
      if (M['id_transaction'] != null)    valeurs[M['id_transaction']] = idTx || '';

      const rd = remiseParNom[typeRemise] || { valeur: 0, estPourcentage: false, magasin: false };
      const remise = rd.estPourcentage ? _round2(prix * rd.valeur) : _round2(rd.valeur);
      const prixClient = _round2(prix + remise);
      const taux = TAUX_TAXE[_norm(data.paiement)] || 0;
      const taxe = _round2(prixClient * taux);
      const prime = rd.magasin ? _round2(prix - taxe) : _round2(prixClient - taxe);

      if (M['remise'] != null && !aFormule('remise'))               valeurs[M['remise']] = remise;
      if (M['prix_client'] != null && !aFormule('prix_client'))     valeurs[M['prix_client']] = prixClient;
      if (M['taxe'] != null && !aFormule('taxe'))                   valeurs[M['taxe']] = taxe;
      if (M['prime_vendeur'] != null && !aFormule('prime_vendeur')) valeurs[M['prime_vendeur']] = prime;

      if (M['validation'] != null) valeurs[M['validation']] = COCHER_VALIDATION ? true : '';

      numerosVente.push(vente);
      vente++;
      return valeurs;
    });

    // 3) Écriture du panier en un seul appel.
    bloc.setValues(tableau);

    // 4) Formules du modèle (R1C1 = recopie relative correcte), une colonne à la fois.
    for (let c = 0; c < nbCol; c++) {
      if (formulesR1C1[c] && formulesR1C1[c] !== '') {
        sh.getRange(premiere, c + 1, nb, 1).setFormulaR1C1(formulesR1C1[c]);
      }
    }

    SpreadsheetApp.flush();
    const res = { ok: true, panier: panier, ventes: numerosVente, nbLignes: nb, deja: false };
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
 * Lignes de l'onglet "ventes" pour un jour donné (lecture seule).
 * @param {string} jourStr  'yyyy-MM-dd' ; vide/invalide = aujourd'hui (fuseau du classeur).
 */
function getVentesDuJour(jourStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_VENTES);
  const tz = ss.getSpreadsheetTimeZone();
  const jour = /^\d{4}-\d{2}-\d{2}$/.test(String(jourStr || ''))
    ? jourStr
    : Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  const h = _headerMap(sh);
  const M = h.map;
  const last = sh.getLastRow();
  if (last < 2) return { jour: jour, lignes: [] };

  const data = sh.getRange(2, 1, last - 1, h.nbCol).getValues();
  const get = function (r, nom) { return M[nom] != null ? r[M[nom]] : ''; };
  const num = function (v) { const n = Number(v); return isNaN(n) ? 0 : n; };

  const lignes = [];
  data.forEach(function (r) {
    const d = get(r, 'date');
    let j, heure = '';
    if (d instanceof Date) {
      j = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      heure = Utilities.formatDate(d, tz, 'HH:mm');
    } else {
      j = String(d || '').trim().slice(0, 10);
    }
    if (j !== jour) return;
    const p = get(r, 'numero_panier');
    lignes.push({
      panier: (p === '' || p == null) ? null : num(p),
      vente: num(get(r, 'numero_vente')),
      heure: heure === '00:00' ? '' : heure,   // 00:00 = saisie manuelle sans heure
      vendeur: String(get(r, 'vendeur') || ''),
      reference: String(get(r, 'reference_produit') || ''),
      remiseType: String(get(r, 'type_de_remise') || '').trim(),
      paiement: String(get(r, 'type_de_paiement') || '').trim(),
      prix: num(get(r, 'prix')),
      remise: num(get(r, 'remise')),
      prixClient: num(get(r, 'prix_client'))
    });
  });
  return { jour: jour, lignes: lignes };
}
