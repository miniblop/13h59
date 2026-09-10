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

/* ---------- Enregistrement d'une vente ---------- */

/**
 * Enregistre un panier dans l'onglet "ventes".
 * @param {Object} data  { paiement:'cb'|'espèces', lignes:[{vendeur,reference,prix,typeRemise}] }
 * @return {Object} { ok, panier, ventes:[numeros], nbLignes }
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

    const numerosVente = [];
    const maintenant = new Date();

    data.lignes.forEach(function (ligne) {
      const newRow = sh.getLastRow() + 1;

      // 1) Recopie le FORMAT + les VALIDATIONS (listes déroulantes, case à cocher…) du modèle.
      if (derniereLigne >= 2) {
        const src = sh.getRange(derniereLigne, 1, 1, nbCol);
        const dst = sh.getRange(newRow, 1, 1, nbCol);
        src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);

        // Assouplit les validations réglées sur « Refuser la saisie » :
        // on garde le menu déroulant, mais une valeur hors liste n'est plus
        // rejetée (sinon la caisse plante quand un vendeur/une remise n'est
        // pas exactement dans la liste de validation de la colonne).
        const dvs = dst.getDataValidations()[0];
        for (var c = 0; c < nbCol; c++) {
          var dv = dvs[c];
          if (dv && dv.getAllowInvalid && !dv.getAllowInvalid()) {
            sh.getRange(newRow, c + 1).setDataValidation(dv.copy().setAllowInvalid(true).build());
          }
        }
      }

      // 2) Prépare les valeurs de saisie.
      const valeurs = new Array(nbCol).fill('');
      const prix = Number(ligne.prix) || 0;
      const typeRemise = ligne.typeRemise || 'pas_de_remise';

      if (M['numero_panier'] != null)     valeurs[M['numero_panier']] = panier;
      if (M['numero_vente'] != null)      valeurs[M['numero_vente']] = vente;
      if (M['date'] != null)              valeurs[M['date']] = maintenant;
      if (M['vendeur'] != null)           valeurs[M['vendeur']] = ligne.vendeur;
      if (M['reference_produit'] != null) valeurs[M['reference_produit']] = ligne.reference || '';
      if (M['type_de_remise'] != null)    valeurs[M['type_de_remise']] = typeRemise;
      if (M['type_de_paiement'] != null)  valeurs[M['type_de_paiement']] = data.paiement;
      if (M['prix'] != null)              valeurs[M['prix']] = prix;

      // 3) Colonnes calculées : SI formule dans le modèle -> on la recopiera (étape 5).
      //    SINON -> calcul de secours ici.
      const aFormule = function (nomCol) {
        const idx = M[nomCol];
        return idx != null && formulesR1C1[idx] && formulesR1C1[idx] !== '';
      };

      const rd = remiseParNom[typeRemise] || { valeur: 0, estPourcentage: false, magasin: false };
      const remise = rd.estPourcentage ? _round2(prix * rd.valeur) : _round2(rd.valeur);
      const prixClient = _round2(prix + remise);
      const taux = TAUX_TAXE[_norm(data.paiement)] || 0;
      const taxe = _round2(prixClient * taux);
      const magasin = rd.magasin;
      const prime = magasin ? _round2(prix - taxe) : _round2(prixClient - taxe);

      if (M['remise'] != null && !aFormule('remise'))               valeurs[M['remise']] = remise;
      if (M['prix_client'] != null && !aFormule('prix_client'))     valeurs[M['prix_client']] = prixClient;
      if (M['taxe'] != null && !aFormule('taxe'))                   valeurs[M['taxe']] = taxe;
      if (M['prime_vendeur'] != null && !aFormule('prime_vendeur')) valeurs[M['prime_vendeur']] = prime;

      // validation : on laisse vide (non validée) sauf config.
      if (M['validation'] != null) valeurs[M['validation']] = COCHER_VALIDATION ? true : '';

      // 4) Écrit les valeurs de la ligne d'un coup.
      sh.getRange(newRow, 1, 1, nbCol).setValues([valeurs]);

      // 5) Recopie les formules (R1C1 = recopie relative correcte) par-dessus.
      for (let c = 0; c < nbCol; c++) {
        if (formulesR1C1[c] && formulesR1C1[c] !== '') {
          sh.getRange(newRow, c + 1).setFormulaR1C1(formulesR1C1[c]);
        }
      }

      numerosVente.push(vente);
      vente++;
    });

    SpreadsheetApp.flush();
    return { ok: true, panier: panier, ventes: numerosVente, nbLignes: numerosVente.length };

  } catch (e) {
    return { ok: false, message: String(e && e.message ? e.message : e) };
  } finally {
    lock.releaseLock();
  }
}