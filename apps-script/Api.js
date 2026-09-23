/*************************************************************
 *  API WEB — Dashboard 13h59
 *  À AJOUTER dans le MÊME projet Apps Script que la caisse
 *  (＋ ▸ Script ▸ colle ce fichier, nomme-le "Api").
 *
 *  Sécurité :
 *   - Les mots de passe ne sont PAS dans le code : ils sont
 *     stockés dans Paramètres du projet ▸ Propriétés du script :
 *        MDP_CREATEUR  = mot de passe des créateurs  → onglet Créateur
 *        MDP_CAISSE    = mot de passe de la caisse   → onglet Caisse
 *        MDP_GESTION   = mot de passe gestionnaire   → accès à tout
 *   - Déploie en "Exécuter en tant que MOI" + "Accès : tout le monde".
 *     Le Sheet peut alors rester PRIVÉ : c'est le script qui y accède.
 *************************************************************/

/*  ⚠️ Aucun `const` global ici : dans Apps Script, TOUS les fichiers .gs partagent
 *  le même scope. Si Code.gs déclare déjà SHEET_DATA, un second `const SHEET_DATA`
 *  fait planter le projet entier ("Identifier ... has already been declared").
 *  On passe donc par des fonctions, qui elles peuvent coexister sans erreur.      */

/** Nettoie un nom de créateur SANS en changer l'apparence :
 *   - normalise l'accentuation Unicode (é composé U+00E9 vs e + ´ U+0301)
 *   - supprime les caractères invisibles (zero-width, BOM)
 *   - ramène espaces insécables / doubles espaces à un espace simple
 *  Résultat : deux noms visuellement identiques deviennent une seule et même chaîne. */
function _nomPropre(v) {
  return String(v == null ? '' : v)
    .normalize('NFC')
    // Caract\u00E8res de formatage strictement invisibles \u00E0 l'\u00E9cran :
    //   00AD trait d'union conditionnel \u00B7 180E s\u00E9parateur mongol
    //   200B-200F zero-width + marques gauche/droite \u00B7 202A-202E encadrement bidi
    //   2060-2064 word joiner & invisibles math\u00E9matiques \u00B7 2066-2069 isolats bidi
    //   FE00-FE0F s\u00E9lecteurs de variante \u00B7 FEFF BOM
    .replace(/[\u00AD\u180E\u200B-\u200F\u202A-\u202E\u2060-\u2064\u2066-\u2069\uFE00-\uFE0F\uFEFF]/g, '')
    // \s couvre d\u00E9j\u00E0 l'espace ins\u00E9cable (00A0), les espaces typographiques
    // (2000-200A), l'espace id\u00E9ographique (3000) et les sauts de ligne.
    .replace(/\s+/g, ' ')
    .trim();
}

/** Nom de l'onglet source. Réutilise SHEET_DATA de Code.gs s'il existe. */
function _sheetData() {
  return (typeof SHEET_DATA !== 'undefined' && SHEET_DATA) ? SHEET_DATA : 'All data';
}

/** Zones accessibles par rôle. */
function _acces() {
  return {
    gestion:  ['caisse', 'createur', 'gestion'],
    caisse:   ['caisse'],
    createur: ['createur']
  };
}

/** Numéro de version du code — sert à vérifier ce qui est réellement DÉPLOYÉ. */
function _version() { return '2026-09-24-createurs'; }

/** Fait converger les variantes de CASSE d'un même créateur.
 *  "hello cloudy" (297 ventes) et "Hello Cloudy" (1 vente) désignent la même
 *  personne : on garde l'orthographe la plus fréquente pour les deux.
 *  Indispensable car l'interface affiche les noms en `text-transform: capitalize`,
 *  ce qui rend les différences de casse invisibles à l'écran. */
function _canoniserCasse(lignes) {
  const freq = {};
  lignes.forEach(function (l) {
    if (!l.createur) return;
    const k = l.createur.toLowerCase();
    if (!freq[k]) freq[k] = {};
    freq[k][l.createur] = (freq[k][l.createur] || 0) + 1;
  });
  const canon = {};
  Object.keys(freq).forEach(function (k) {
    let best = null, n = -1;
    Object.keys(freq[k]).forEach(function (nom) {
      if (freq[k][nom] > n) { n = freq[k][nom]; best = nom; }
    });
    canon[k] = best;
  });
  lignes.forEach(function (l) {
    if (l.createur) l.createur = canon[l.createur.toLowerCase()] || l.createur;
  });
  return lignes;
}

/** Point d'entrée des appels POST du site. */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'ping')        return _json(_ping());
    if (body.action === 'data')        return _apiData(body);
    if (body.action === 'caisse_data') return _caisseData(body);
    if (body.action === 'caisse_save') return _caisseSave(body);
    if (body.action === 'caisse_jour') return _caisseJour(body);
    return _json({ ok: false, message: 'Action inconnue.' });
  } catch (err) {
    return _json({ ok: false, message: 'Erreur : ' + (err && err.message ? err.message : err) });
  }
}

/** Renvoie le rôle ('gestion' / 'caisse' / 'createur') d'un mot de passe, ou null. */
function _role(password) {
  const props = PropertiesService.getScriptProperties();
  // on ignore les espaces parasites de part et d'autre (copier-coller, clavier mobile)
  const net = function (v) { return String(v == null ? '' : v).trim(); };
  const p    = net(password);
  const mdpCr = net(props.getProperty('MDP_CREATEUR'));
  const mdpCa = net(props.getProperty('MDP_CAISSE'));
  const mdpG  = net(props.getProperty('MDP_GESTION'));
  if (!p) return null;
  if (mdpG && p === mdpG) return 'gestion';
  if (mdpCa && p === mdpCa) return 'caisse';
  if (mdpCr && p === mdpCr) return 'createur';
  return null;
}

/** Le rôle a-t-il accès à cette zone ? */
function _peut(role, zone) {
  const a = _acces();
  return !!role && !!a[role] && a[role].indexOf(zone) !== -1;
}

/** Données pour l'interface caisse (réutilise getCaisseData de Code.gs, inchangé). */
function _caisseData(body) {
  const role = _role(body.password);
  if (!role) return _json({ ok: false, message: 'Mot de passe incorrect.' });
  if (!_peut(role, 'caisse')) return _json({ ok: false, message: "Ce mot de passe ne donne pas accès à la caisse." });

  const d = getCaisseData() || {};

  // Normalisation défensive : on garantit toujours une liste de créateurs exploitable.
  let createurs = d.createurs;
  if (!Array.isArray(createurs)) createurs = [];
  createurs = createurs
    .map(_nomPropre)
    .filter(function (v) { return v !== ''; });
  // dédoublonnage insensible à la casse (l'interface affiche en capitalize,
  // donc "hello cloudy" et "Hello Cloudy" seraient deux lignes identiques)
  const vus = {};
  createurs = createurs.filter(function (v) {
    const k = v.toLowerCase();
    if (vus[k]) return false;
    vus[k] = 1;
    return true;
  }).sort(function (a, b) { return a.localeCompare(b, 'fr'); });

  // Filet de secours : si getCaisseData ne renvoie rien, on déduit les créateurs de "All data".
  if (!createurs.length) {
    const vusSecours = {};
    _lireDonnees().forEach(function (l) { if (l.createur) vusSecours[l.createur] = 1; });
    createurs = Object.keys(vusSecours).sort(function (a, b) { return a.localeCompare(b, 'fr'); });
  }

  const remises = Array.isArray(d.remises) && d.remises.length
    ? d.remises
    : [{ nom: 'pas_de_remise', valeur: 0, estPourcentage: false, magasin: false }];

  const paiements = Array.isArray(d.paiements) && d.paiements.length ? d.paiements : ['cb', 'espèces'];

  return _json({
    ok: true,
    role: role,
    createurs: createurs,
    remises: remises,
    paiements: paiements,
    prochainPanier: d.prochainPanier,
    prochaineVente: d.prochaineVente
  });
}

/** Enregistre une vente (réutilise enregistrerVente de Code.gs, inchangé). */
function _caisseSave(body) {
  const role = _role(body.password);
  if (!role) return _json({ ok: false, message: 'Mot de passe incorrect.' });
  if (!_peut(role, 'caisse')) return _json({ ok: false, message: "Ce mot de passe ne donne pas accès à la caisse." });
  return _json(enregistrerVente(body.payload));
}

/* Ventes d'une journée pour le récap de l'onglet Caisse (lecture seule).
   body.date = 'yyyy-MM-dd' (optionnel, défaut : aujourd'hui). */
function _caisseJour(body) {
  const role = _role(body.password);
  if (!role) return _json({ ok: false, message: 'Mot de passe incorrect.' });
  if (!_peut(role, 'caisse')) return _json({ ok: false, message: "Ce mot de passe ne donne pas accès à la caisse." });
  const r = getVentesDuJour(body.date);
  r.lignes.forEach(function (l) { l.createur = _nomPropre(l.createur); });
  return _json({ ok: true, jour: r.jour, lignes: r.lignes });
}

/** Variante GET (JSONP) — filet de secours si un jour on a un souci CORS.
 *  Appel : ...exec?action=data&password=xxx&callback=cb  */
function doGet(e) {
  const p = e.parameter || {};
  const res = (p.action === 'ping') ? _ping() : _apiDataObj(p.password);
  const txt = JSON.stringify(res);
  if (p.callback) {
    return ContentService.createTextOutput(p.callback + '(' + txt + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return _json(res);
}

/** Diagnostic : dit quelle version est DÉPLOYÉE et quels mots de passe sont définis.
 *  Ne renvoie AUCUNE valeur de mot de passe, uniquement vrai/faux.
 *  Ouvre .../exec?action=ping dans le navigateur pour le lire. */
function _ping() {
  const props = PropertiesService.getScriptProperties();
  const def = function (k) {
    const v = props.getProperty(k);
    return { defini: !!v, longueur: v ? v.length : 0, espacesParasites: !!v && v !== v.trim() };
  };
  return {
    ok: true,
    version: _version(),
    ongletSource: _sheetData(),
    motsDePasse: {
      MDP_CREATEUR: def('MDP_CREATEUR'),
      MDP_CAISSE:   def('MDP_CAISSE'),
      MDP_GESTION:  def('MDP_GESTION')
    }
  };
}

function _apiData(body) {
  return _json(_apiDataObj(body.password));
}

function _apiDataObj(password) {
  const role = _role(password);
  if (!role) return { ok: false, message: 'Mot de passe incorrect.' };
  const zones = _acces()[role] || [];
  // Le rôle "caisse" n'a pas besoin des données de reporting : on ne les envoie pas.
  const rows = (zones.indexOf('createur') !== -1 || zones.indexOf('gestion') !== -1) ? _lireDonnees() : [];
  return { ok: true, role: role, zones: zones, rows: rows };
}

/** Lit l'onglet source et renvoie des lignes normalisées (clés courtes). */
function _lireDonnees() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(_sheetData());
  if (!sh) return [];
  const tz = ss.getSpreadsheetTimeZone();
  const nbCol = sh.getLastColumn();
  const nbLig = sh.getLastRow() - 1;
  if (nbLig < 1) return [];

  const entetes = sh.getRange(1, 1, 1, nbCol).getValues()[0]
    .map(function (h) { return String(h).trim().toLowerCase(); });
  const idx = {};
  entetes.forEach(function (h, i) { idx[h] = i; });

  const data = sh.getRange(2, 1, nbLig, nbCol).getValues();
  function num(v) { const n = Number(v); return isNaN(n) ? 0 : n; }
  function get(r, nom) { return idx[nom] != null ? r[idx[nom]] : ''; }

  const out = [];
  data.forEach(function (r) {
    const d = get(r, 'date');
    const dateStr = (d instanceof Date)
      ? Utilities.formatDate(d, tz, 'yyyy-MM-dd')
      : String(d || '').trim();
    if (!dateStr) return;
    const panierBrut = get(r, 'numero_panier');
    out.push({
      date: dateStr,
      createur: _nomPropre(get(r, COL_CREATEUR)),
      panier: (panierBrut === '' || panierBrut == null) ? null : num(panierBrut),
      vente: num(get(r, 'numero_vente')),
      reference: String(get(r, 'reference_produit') || ''),
      paiement: String(get(r, 'type_de_paiement') || '').trim(),
      remiseType: String(get(r, 'type_de_remise') || '').trim(),
      prix: num(get(r, 'prix')),
      remise: num(get(r, 'remise')),
      prixClient: num(get(r, 'prix_client')),
      taxe: num(get(r, 'taxe')),
      prime: num(get(r, COL_PRIME_CREATEUR))
    });
  });
  return _canoniserCasse(out);
}

/** Rend visibles les caractères non-ASCII d'une chaîne : "Hello Cloudy" (espace
 *  insécable) devient "Hello Cloudy". Sert au diagnostic ci-dessous. */
function _echappe(s) {
  return String(s).replace(/[^\x20-\x7E]/g, function (c) {
    return '\\u' + ('000' + c.charCodeAt(0).toString(16).toUpperCase()).slice(-4);
  });
}

/** DIAGNOSTIC (lecture seule, n'écrit RIEN dans le Sheet).
 *  Modifie la valeur ci-dessous puis lance la fonction depuis l'éditeur :
 *  elle liste chaque orthographe réellement présente dans les ventes, caractère
 *  par caractère, et dit si elles fusionnent après nettoyage.
 *  Exemple de sortie :  "Hello Cloudy"  ≠  "Hello Cloudy"  */
function inspecterNom() {
  const recherche = 'hello';   // ← mets ici un bout du nom à examiner

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(_sheetData());
  if (!sh) { Logger.log('Onglet introuvable : ' + _sheetData()); return; }
  const entetes = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h).trim().toLowerCase(); });
  const col = entetes.indexOf(COL_CREATEUR) + 1;
  if (!col) { Logger.log("Colonne '" + COL_CREATEUR + "' introuvable."); return; }

  const nbLig = sh.getLastRow() - 1;
  const vals = sh.getRange(2, col, nbLig, 1).getValues();
  const q = String(recherche).toLowerCase();

  const vus = {};
  vals.forEach(function (r) {
    const brut = String(r[0] == null ? '' : r[0]);
    if (brut.toLowerCase().indexOf(q) === -1) return;
    vus[brut] = (vus[brut] || 0) + 1;
  });

  const noms = Object.keys(vus);
  if (!noms.length) { Logger.log('Aucun nom ne contient "' + recherche + '".'); return; }

  Logger.log(noms.length + ' orthographe(s) contenant "' + recherche + '" :');
  noms.forEach(function (n, i) {
    Logger.log('  [' + (i + 1) + '] ' + vus[n] + ' vente(s)');
    Logger.log('      brut    : "' + _echappe(n) + '"');
    Logger.log('      nettoyé : "' + _echappe(_nomPropre(n)) + '"');
  });

  const apres = {};
  noms.forEach(function (n) { apres[_nomPropre(n)] = 1; });
  const restants = Object.keys(apres).length;
  Logger.log(restants === 1
    ? '→ Après nettoyage : 1 seule entrée, la fusion fonctionne ✅'
    : '→ Après nettoyage : ' + restants + ' entrées distinctes ❌ (compare les lignes "nettoyé" ci-dessus)');
}

/** Petit test à lancer depuis l'éditeur : vérifie que les 3 mots de passe existent
 *  et que getCaisseData() renvoie bien des créateurs. Voir Exécutions ▸ Journaux. */
function testConfig() {
  const props = PropertiesService.getScriptProperties();
  Logger.log('Onglet source utilisé : ' + _sheetData());
  ['MDP_CREATEUR', 'MDP_CAISSE', 'MDP_GESTION'].forEach(function (k) {
    Logger.log(k + ' : ' + (props.getProperty(k) ? 'défini ✅' : 'MANQUANT ❌'));
  });
  try {
    const d = getCaisseData() || {};
    Logger.log('getCaisseData().createurs: ' + JSON.stringify(d.createurs));
    Logger.log('getCaisseData().remises  : ' + JSON.stringify(d.remises));
    Logger.log('getCaisseData().paiements: ' + JSON.stringify(d.paiements));
  } catch (e) {
    Logger.log('getCaisseData() a échoué ❌ : ' + e.message);
  }
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
