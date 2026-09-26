/*************************************************************
 *  API WEB — Dashboard 13h59
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

/** Nettoie un nom SANS en changer l'apparence :
 *   - normalise l'accentuation Unicode (é composé U+00E9 vs e + ´ U+0301)
 *   - supprime les caractères invisibles (zero-width, BOM)
 *   - ramène espaces insécables / doubles espaces à un espace simple */
function _nomPropre(v) {
  return String(v == null ? '' : v)
    .normalize('NFC')
    .replace(/[­᠎​-‏‪-‮⁠-⁤⁦-⁩︀-️﻿]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
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
function _version() { return '2026-09-suppression-fiche'; }

/** Point d'entrée des appels POST du site. */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'ping')        return _json(_ping());
    if (body.action === 'data')        return _json(_apiDataObj(body.password, body.compact === true));
    if (body.action === 'caisse_data') return _caisseData(body);
    if (body.action === 'caisse_save') return _caisseSave(body);
    if (body.action === 'caisse_jour') return _caisseJour(body);
    if (body.action === 'candidature_infos')   return _json(_lecturePure(_candidatureInfos));
    if (body.action === 'candidature_envoyer') return _json(_candidatureEnvoyer(body));
    if (String(body.action).indexOf('gestion_') === 0) return _json(_gestion(body));
    return _json({ ok: false, message: 'Action inconnue.' });
  } catch (err) {
    return _json({ ok: false, message: 'Erreur : ' + (err && err.message ? err.message : err) });
  }
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

/** Renvoie le rôle ('gestion' / 'caisse' / 'createur') d'un mot de passe, ou null. */
function _role(password) {
  const props = PropertiesService.getScriptProperties();
  // on ignore les espaces parasites de part et d'autre (copier-coller, clavier mobile)
  const net = function (v) { return String(v == null ? '' : v).trim(); };
  const p     = net(password);
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

function _refusCaisse(body) {
  const role = _role(body.password);
  if (!role) return 'Mot de passe incorrect.';
  if (!_peut(role, 'caisse')) return "Ce mot de passe ne donne pas accès à la caisse.";
  return null;
}

/** Données pour l'interface caisse : créateurs actifs {id, nom}, remises, paiements, prochains numéros. */
function _caisseData(body) {
  const refus = _refusCaisse(body);
  if (refus) return _json({ ok: false, message: refus });
  // À l'ouverture de la caisse, les créateurs qui arrivent ou partent aujourd'hui sont mis à jour.
  // (une fois par jour : la gestion le refait de toute façon à chaque ouverture de Créateurs)
  const cache = CacheService.getScriptCache(), cle = 'statuts_' + _jourIso(new Date());
  if (!cache.get(cle)) {
    const lock = LockService.getScriptLock();
    if (lock.tryLock(5000)) { try { _synchroniserStatuts(SpreadsheetApp.getActiveSpreadsheet()); SpreadsheetApp.flush(); cache.put(cle, '1', 21600); } finally { lock.releaseLock(); } }
  }
  const d = getCaisseData();
  d.ok = true;
  d.role = _role(body.password);
  return _json(d);
}

/** Enregistre une vente. */
function _caisseSave(body) {
  const refus = _refusCaisse(body);
  if (refus) return _json({ ok: false, message: refus });
  return _json(enregistrerVente(body.payload));
}

/* Ventes d'une journée pour le récap de l'onglet Caisse (lecture seule).
   body.date = 'yyyy-MM-dd' (optionnel, défaut : aujourd'hui). */
function _caisseJour(body) {
  const refus = _refusCaisse(body);
  if (refus) return _json({ ok: false, message: refus });
  const r = getVentesDuJour(body.date);
  return _json({ ok: true, jour: r.jour, lignes: r.lignes });
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
    ongletSource: SHEET_VENTES,
    motsDePasse: {
      MDP_CREATEUR: def('MDP_CREATEUR'),
      MDP_CAISSE:   def('MDP_CAISSE'),
      MDP_GESTION:  def('MDP_GESTION')
    }
  };
}

function _apiDataObj(password, compact) {
  const role = _role(password);
  if (!role) return { ok: false, message: 'Mot de passe incorrect.' };
  const zones = _acces()[role] || [];
  // Le rôle "caisse" n'a pas besoin des données de reporting : on ne les envoie pas.
  const rows = (zones.indexOf('createur') !== -1 || zones.indexOf('gestion') !== -1) ? _lecturePure(_lireDonnees) : [];
  if (!compact) return { ok: true, role: role, zones: zones, rows: rows };
  return { ok: true, role: role, zones: zones, compact: _compacter(rows) };
}
/** Format compact (2 à 3 fois plus léger) : une ligne = un tableau, les noms des créateurs une seule fois. */
function _compacter(rows) {
  const noms = {};
  rows.forEach(function (r) { noms[r.idCreateur] = r.createur; });
  return { colonnes: COLONNES_DONNEES, noms: noms, lignes: rows.map(function (r) { return COLONNES_DONNEES.map(function (k) { return r[k]; }); }) };
}
const COLONNES_DONNEES = ['date', 'idCreateur', 'panier', 'vente', 'reference', 'paiement', 'remiseType', 'prix', 'remise', 'prixClient', 'taxe', 'prime', 'sens'];

/** Toutes les ventes, avec le nom du créateur (clés courtes attendues par le site). */
function _lireDonnees() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const noms = _createurs(ss).parId;
  const num = function (v) { const n = Number(v); return isNaN(n) ? 0 : n; };
  const out = [];
  _lireTable(_onglet(ss, SHEET_VENTES)).forEach(function (v) {
    if (!(v['date'] instanceof Date)) return;
    const id = String(v['id_createur'] || '');
    out.push({
      date: _jourIso(v['date']),
      idCreateur: id,
      createur: noms[id] ? noms[id].nom : id,
      panier: (v['id_panier'] === '' || v['id_panier'] == null) ? null : num(v['id_panier']),
      vente: num(v['id_vente']),
      reference: String(v['reference'] || ''),
      paiement: String(v['code_paiement'] || '').trim(),
      remiseType: String(v['code_remise'] || '').trim(),
      prix: num(v['prix']),
      remise: num(v['remise']),
      prixClient: num(v['prix_client']),
      taxe: num(v['frais']),
      prime: num(v['prime']),
      sens: _norm(v['type_ligne']) === 'annulation' ? -1 : 1
    });
  });
  return out;
}

/** Petit test à lancer depuis l'éditeur : mots de passe et données de caisse. Voir Exécutions ▸ Journaux. */
function testConfig() {
  const props = PropertiesService.getScriptProperties();
  ['MDP_CREATEUR', 'MDP_CAISSE', 'MDP_GESTION'].forEach(function (k) {
    Logger.log(k + ' : ' + (props.getProperty(k) ? 'défini ✅' : 'MANQUANT ❌'));
  });
  try {
    const d = getCaisseData();
    Logger.log('Créateurs actifs : ' + d.createurs.length + ' · remises : ' + d.remises.length + ' · paiements : ' + d.paiements.join(', ') +
      ' · prochain panier ' + d.prochainPanier + ' · prochaine vente ' + d.prochaineVente);
  } catch (e) {
    Logger.log('getCaisseData() a échoué ❌ : ' + e.message);
  }
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
