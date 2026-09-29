/*************************************************************
 *  ACCÈS NOMINATIF À LA GESTION — code à 6 chiffres par e-mail
 *  Onglet `utilisateurs` : les personnes autorisées (e-mail, prénom).
 *  1. connexion_code {email}        : envoie un code (10 min, 5 essais) si
 *     l'adresse est autorisée ; la réponse est la même dans tous les cas.
 *  2. connexion_verifier {email, code} : ouvre une session de 30 jours,
 *     renvoyée au site sous la forme « S:<jeton> » qui remplace le mot de passe.
 *  Les sessions sont gardées dans les propriétés du script (empreinte du
 *  jeton seulement). Retirer une personne ferme ses sessions. Le journal
 *  signe avec le prénom enregistré, plus avec un prénom saisi.
 *  Le mot de passe MDP_GESTION reste un accès de secours.
 *************************************************************/

const SHEET_UTILISATEURS = 'utilisateurs';
const COLONNES_UTILISATEURS = ['email', 'prenom', 'actif', 'ajoute_le', 'ajoute_par', 'derniere_connexion'];
const DUREE_SESSION_JOURS = 30;
let _sessionUtilisateur = null;   // personne connectée par code pendant la requête en cours

function _empreinte(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function _utilisateurs(ss) {
  if (!ss.getSheetByName(SHEET_UTILISATEURS)) return [];
  return _lireTable(_onglet(ss, SHEET_UTILISATEURS)).filter(function (u) { return u['email']; }).map(function (u) {
    return { email: _norm(u['email']), prenom: String(u['prenom'] || ''), actif: u['actif'] !== false, ajouteLe: _iso(u['ajoute_le']), ajoutePar: String(u['ajoute_par'] || ''), derniereConnexion: _iso(u['derniere_connexion']) };
  });
}

/** Session « S:<32 hexa> » → personne connectée, ou null (expirée, inconnue, personne désactivée). */
function _session(motDePasse) {
  const m = /^S:([0-9a-f]{32})$/.exec(String(motDePasse || '').trim());
  if (!m) return null;
  const props = PropertiesService.getScriptProperties(), cle = 'SESSION_' + _empreinte(m[1]).slice(0, 40);
  let s = null;
  try { s = JSON.parse(props.getProperty(cle) || 'null'); } catch (e) { s = null; }
  if (!s || !(s.exp > Date.now())) { if (s) props.deleteProperty(cle); return null; }
  // la personne a pu être retirée ou désactivée depuis (vérifié au plus une fois par heure)
  const cache = CacheService.getScriptCache(), ck = 'u_ok_' + _empreinte(s.email).slice(0, 20);
  if (!cache.get(ck)) {
    const u = _utilisateurs(SpreadsheetApp.getActiveSpreadsheet()).filter(function (x) { return x.email === s.email && x.actif; })[0];
    if (!u) { props.deleteProperty(cle); return null; }
    s.prenom = u.prenom;
    cache.put(ck, '1', 3600);
  }
  return { email: s.email, prenom: s.prenom, cle: cle };
}

function _connexionCode(body) {
  const email = _email(body.email), neutre = { ok: true, message: 'Si cette adresse a accès à la gestion, un code vient de lui être envoyé (pense aux spams).' };
  if (!email) return { ok: false, message: 'Adresse e-mail invalide.' };
  const cache = CacheService.getScriptCache(), cleN = 'code_n_' + _empreinte(email).slice(0, 20);
  const n = Number(cache.get(cleN) || 0);
  if (n >= 3) return { ok: false, message: 'Trop de demandes : réessaie dans 15 minutes.' };
  cache.put(cleN, String(n + 1), 900);
  const u = _utilisateurs(SpreadsheetApp.getActiveSpreadsheet()).filter(function (x) { return x.email === email && x.actif; })[0];
  if (!u) return neutre;   // même réponse : on ne dit pas quelles adresses existent
  const code = String(parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 10), 16) % 1000000).padStart(6, '0');
  cache.put('code_' + _empreinte(email).slice(0, 20), JSON.stringify({ h: _empreinte(code + email), essais: 0 }), 600);
  const texte = 'Bonjour ' + (u.prenom || '') + ',\n\nTon code pour te connecter à la gestion 13H59 : ' + code + '\n\nIl est valable 10 minutes. Si tu n\'as rien demandé, ignore cet e-mail.';
  envoyerEmailShop({ to: email, subject: 'Ton code de connexion à la gestion 13H59', texte: texte,
    html: _htmlShop(texte).replace(code, '<b style="font-size:22px;letter-spacing:4px">' + code + '</b>') });
  return neutre;
}

function _connexionVerifier(body) {
  const email = _email(body.email), code = String(body.code || '').replace(/\s/g, '');
  if (!email || !/^\d{6}$/.test(code)) return { ok: false, message: 'Indique le code à 6 chiffres reçu par e-mail.' };
  const cache = CacheService.getScriptCache(), cle = 'code_' + _empreinte(email).slice(0, 20);
  let c = null;
  try { c = JSON.parse(cache.get(cle) || 'null'); } catch (e) { c = null; }
  if (!c) return { ok: false, message: 'Code expiré : demandes-en un nouveau.' };
  if (c.h !== _empreinte(code + email)) {
    c.essais++;
    if (c.essais >= 5) { cache.remove(cle); return { ok: false, message: 'Trop d\'essais : demande un nouveau code.' }; }
    cache.put(cle, JSON.stringify(c), 600);
    return { ok: false, message: 'Code incorrect.' };
  }
  cache.remove(cle);
  const ss = SpreadsheetApp.getActiveSpreadsheet(), t = _tableau(ss, SHEET_UTILISATEURS);
  const i = t.lignes.findIndex(function (r) { return _norm(_val(t, r, 'email')) === email && _val(t, r, 'actif') !== false; });
  if (i < 0) return { ok: false, message: "Cette adresse n'a plus accès à la gestion." };
  const r = t.lignes[i], prenom = String(_val(t, r, 'prenom') || '');
  r[t.M['derniere_connexion']] = new Date(); _ecrireLigne(t, i, r);
  const jeton = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('SESSION_' + _empreinte(jeton).slice(0, 40),
    JSON.stringify({ email: email, prenom: prenom, exp: Date.now() + DUREE_SESSION_JOURS * 864e5 }));
  _journaliser('connexion', prenom + ' (' + email + ') : session de ' + DUREE_SESSION_JOURS + ' jours ouverte', prenom + ' (site)');
  return { ok: true, jeton: 'S:' + jeton, prenom: prenom };
}

function _connexionFin(body) {
  const s = _session(body.password);
  if (s) PropertiesService.getScriptProperties().deleteProperty(s.cle);
  return { ok: true };
}

/** Ferme toutes les sessions d'une adresse. */
function _fermerSessions(email) {
  const props = PropertiesService.getScriptProperties(), toutes = props.getProperties();
  let n = 0;
  Object.keys(toutes).forEach(function (k) {
    if (k.indexOf('SESSION_') !== 0) return;
    try { if (JSON.parse(toutes[k]).email === email) { props.deleteProperty(k); n++; } } catch (e) { props.deleteProperty(k); }
  });
  return n;
}

/* ---------- Réglages ▸ Accès à la gestion ---------- */

function _gUtilisateurs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _creerOngletSiAbsent(ss, SHEET_UTILISATEURS, COLONNES_UTILISATEURS);
  const toutes = PropertiesService.getScriptProperties().getProperties(), sessions = {};
  Object.keys(toutes).forEach(function (k) {
    if (k.indexOf('SESSION_') !== 0) return;
    try { const s = JSON.parse(toutes[k]); if (s.exp > Date.now()) sessions[s.email] = (sessions[s.email] || 0) + 1; } catch (e) { /* ignorée */ }
  });
  return { ok: true, utilisateurs: _utilisateurs(ss).map(function (u) { u.sessions = sessions[u.email] || 0; return u; }), moi: _sessionUtilisateur ? _sessionUtilisateur.email : '' };
}
function _gUtilisateurAjouter(body) {
  const email = _email(body.email), prenom = _nomPropre(_textePublic(body.prenom, 30));
  if (!email) throw new Error('Adresse e-mail invalide.');
  if (!prenom) throw new Error('Indique le prénom (il signe les modifications dans le journal).');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _creerOngletSiAbsent(ss, SHEET_UTILISATEURS, COLONNES_UTILISATEURS);
  const t = _tableau(ss, SHEET_UTILISATEURS), i = t.lignes.findIndex(function (r) { return _norm(_val(t, r, 'email')) === email; });
  if (i >= 0) {
    const r = t.lignes[i];
    if (_val(t, r, 'actif') !== false) throw new Error(email + ' a déjà accès.');
    r[t.M['actif']] = true; r[t.M['prenom']] = prenom; _ecrireLigne(t, i, r);
  } else _ajouterLigne(t, { email: email, prenom: prenom, actif: true, ajoute_le: new Date(), ajoute_par: _signataire() });
  _journaliser('acces_gestion', prenom + ' (' + email + ') : accès à la gestion donné');
  return { ok: true };
}
function _gUtilisateurRetirer(body) {
  const email = _email(body.email), ss = SpreadsheetApp.getActiveSpreadsheet(), t = _tableau(ss, SHEET_UTILISATEURS);
  const i = t.lignes.findIndex(function (r) { return _norm(_val(t, r, 'email')) === email; });
  if (i < 0) throw new Error('Adresse introuvable.');
  if (_sessionUtilisateur && _sessionUtilisateur.email === email) throw new Error('Tu ne peux pas retirer ton propre accès.');
  const r = t.lignes[i]; r[t.M['actif']] = false; _ecrireLigne(t, i, r);
  const n = _fermerSessions(email);
  CacheService.getScriptCache().remove('u_ok_' + _empreinte(email).slice(0, 20));
  _journaliser('acces_gestion', _val(t, r, 'prenom') + ' (' + email + ') : accès retiré, ' + n + ' session(s) fermée(s)');
  return { ok: true, sessions: n };
}
