/*************************************************************
 *  E-MAILS — envoi au nom du shop
 *  Tous les e-mails automatiques partent de l'alias du shop, avec
 *  le shop en « Répondre à » : les réponses arrivent dans sa boîte. Pas de copie
 *  cachée (elle remplissait la boîte du shop) : la trace est dans `journal`.
 *  L'alias doit être configuré dans le Gmail du compte qui exécute le
 *  script (Paramètres ▸ Comptes ▸ « Envoyer des e-mails en tant que »).
 *************************************************************/

const EMAIL_SHOP = '13h59shop@gmail.com';
const NOM_EXPEDITEUR = 'Collectif 13H59';

function _aliasShopDisponible() {
  return GmailApp.getAliases().map(function (a) { return a.toLowerCase(); }).indexOf(EMAIL_SHOP) !== -1;
}

/**
 * Envoie un e-mail au nom du shop.
 * @param {Object} m  { to, subject, html, texte?, pieces? (Blob[]) }
 */
function envoyerEmailShop(m) {
  if (!_aliasShopDisponible()) {
    throw new Error("L'alias " + EMAIL_SHOP + " n'est pas configuré dans le Gmail de " +
      Session.getEffectiveUser().getEmail() + ' : envoi annulé.');
  }
  GmailApp.sendEmail(m.to, m.subject, m.texte || m.subject, {
    from: EMAIL_SHOP,
    name: NOM_EXPEDITEUR,
    replyTo: EMAIL_SHOP,
    htmlBody: m.html,
    attachments: m.pieces || []
  });
}

/* ---------- Modèles (onglet `emails`) ----------
 * Une ligne par e-mail automatique : code, objet, texte, actif.
 * Le texte se modifie directement dans le Sheet ; les champs entre accolades
 * sont remplacés à l'envoi : {prenom} {marque} {stand} {date}.
 * Décocher « actif » coupe l'envoi de ce modèle sans toucher au code. */

const COLONNES_EMAILS = ['code', 'objet', 'texte', 'actif', 'utilise_pour'];
const EMAILS_PAR_DEFAUT = [
  ['accuse_reception', 'Ta candidature au 13H59 Shop est bien reçue',
    "Bonjour {prenom},\n\nMerci pour ta candidature {marque} ! On l'a bien reçue.\n\n" +
    "On prend le temps de découvrir chaque univers avec attention, et on te répond dans tous les cas, que ta candidature soit retenue ou non. " +
    "Si le forfait que tu as choisi est complet, on garde ta candidature et on te recontacte dès qu'une place se libère.\n\n" +
    "À très vite,\nL'équipe du 13H59 Shop",
    true, 'Envoyé automatiquement à chaque candidature reçue par le site.'],
  ['liste_attente', 'Ta candidature {marque} au 13H59 Shop',
    "Bonjour {prenom},\n\nMerci encore pour ta candidature {marque}. Ton univers nous plaît, mais on ne peut pas t'accueillir tout de suite sur le forfait « {stand} ».\n\n" +
    "On garde ta candidature et on te recontacte dès qu'une place se libère.\n\nÀ bientôt,\nL'équipe du 13H59 Shop",
    true, 'Gestion ▸ Candidatures : bouton « Liste d\'attente » (case e-mail cochée).'],
  ['non_retenu', 'Ta candidature {marque} au 13H59 Shop',
    "Bonjour {prenom},\n\nMerci pour ta candidature {marque} et pour le temps que tu y as consacré.\n\n" +
    "Après l'avoir étudiée avec attention, on ne peut pas te proposer de place au 13H59 Shop pour le moment : on veille à garder un équilibre entre les univers présents en boutique. " +
    "Ce n'est pas un jugement sur ton travail, et tu peux tout à fait candidater à nouveau plus tard.\n\nBelle continuation,\nL'équipe du 13H59 Shop",
    true, 'Gestion ▸ Candidatures : bouton « Non retenue » (case e-mail cochée).'],
  ['retenu', 'Bienvenue au 13H59 Shop, {marque} !',
    "Bonjour {prenom},\n\nBonne nouvelle : ta candidature {marque} est retenue ! On t'accueille en boutique à partir du {date}, sur le forfait « {stand} ».\n\n" +
    "Pour préparer ton arrivée, garde sous la main : ton numéro SIRET, ton RIB, ton attestation d'assurance RC Pro et une pièce d'identité. " +
    "L'adhésion à l'association (15 €, valable un an) est réglée avec le premier loyer.\n\n" +
    "On revient vers toi avec la convention à signer. Si tu as la moindre question, réponds simplement à cet e-mail.\n\nÀ très vite,\nL'équipe du 13H59 Shop",
    true, 'Gestion ▸ Candidatures : bouton « Retenir » (case e-mail cochée).']
];

/** Crée l'onglet `emails` avec les textes par défaut s'il n'existe pas. */
function _creerOngletEmails(ss) {
  if (ss.getSheetByName(SHEET_EMAILS)) return false;
  const sh = ss.insertSheet(SHEET_EMAILS);
  sh.getRange(1, 1, 1, COLONNES_EMAILS.length).setValues([COLONNES_EMAILS]).setFontWeight('bold');
  sh.getRange(2, 1, EMAILS_PAR_DEFAUT.length, COLONNES_EMAILS.length).setValues(EMAILS_PAR_DEFAUT);
  sh.getRange(2, 4, EMAILS_PAR_DEFAUT.length, 1).insertCheckboxes();
  sh.getRange(2, 3, EMAILS_PAR_DEFAUT.length, 1).setWrap(true);
  sh.setColumnWidth(2, 260); sh.setColumnWidth(3, 520); sh.setColumnWidth(5, 280);
  sh.setFrozenRows(1);
  return true;
}

/** Modèles par code : {code: {objet, texte, actif, utilisePour}}. */
function _modelesEmails(ss) {
  const sh = ss.getSheetByName(SHEET_EMAILS), out = {};
  if (!sh) return out;
  _lireTable(sh).forEach(function (m) {
    if (m['code']) out[String(m['code'])] = { objet: String(m['objet'] || ''), texte: String(m['texte'] || ''), actif: m['actif'] === true, utilisePour: String(m['utilise_pour'] || '') };
  });
  return out;
}

/** Textes d'origine, pour le bouton « Rétablir » du site. */
function _emailsParDefaut() {
  const out = {};
  EMAILS_PAR_DEFAUT.forEach(function (l) { out[l[0]] = { objet: l[1], texte: l[2] }; });
  return out;
}

/** Champs utilisables dans chaque modèle ({date} n'a de sens que pour l'arrivée d'un créateur retenu). */
const CHAMPS_EMAIL = ['prenom', 'marque', 'stand', 'date'];
function _champsEmailInvalides(code, s) {
  const out = [];
  String(s).replace(/\{([^{}]*)\}/g, function (t, k) {
    if (CHAMPS_EMAIL.indexOf(k) === -1 || (k === 'date' && code !== 'retenu')) out.push(t);
    return t;
  });
  return out;
}

/** Gestion ▸ Candidatures ▸ E-mails types : modifie un modèle de l'onglet `emails`. */
function _gEmailMaj(body) {
  const code = String(body.code || '');
  const defaut = EMAILS_PAR_DEFAUT.filter(function (l) { return l[0] === code; })[0];
  if (!defaut) throw new Error('Modèle inconnu : « ' + code + ' ».');
  const objet = _textePublic(String(body.objet || '').replace(/\n/g, ' '), 200);
  const texte = _textePublic(body.texte, 5000);
  if (!objet) throw new Error("L'objet de l'e-mail est vide.");
  if (texte.length < 20) throw new Error("Le texte de l'e-mail est vide ou trop court.");
  const mauvais = _champsEmailInvalides(code, objet + ' ' + texte);
  if (mauvais.length) throw new Error('Champ inconnu dans le texte : ' + mauvais.join(', ') + '. Champs possibles : {prenom} {marque} {stand}' + (code === 'retenu' ? ' {date}' : '') + '.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _creerOngletEmails(ss);
  const t = _tableau(ss, SHEET_EMAILS);
  const i = t.lignes.findIndex(function (r) { return String(_val(t, r, 'code')) === code; });
  const valeurs = { code: code, objet: objet, texte: texte, actif: body.actif === true, utilise_pour: defaut[4] };
  if (i < 0) _ajouterLigne(t, valeurs);
  else {
    const r = t.lignes[i];
    ['objet', 'texte', 'actif'].forEach(function (k) { if (t.M[k] != null) r[t.M[k]] = valeurs[k]; });
    _ecrireLigne(t, i, r);
  }
  _journaliser('email_modele_maj', code + (body.actif === true ? '' : ' (désactivé)'));
  return { ok: true };
}

function _remplir(modele, vars) {
  return String(modele).replace(/\{(\w+)\}/g, function (t, k) { return vars[k] != null ? String(vars[k]) : t; });
}
function _echapperHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
/** Mise en page du shop : bandeau noir, paragraphes du texte. */
function _htmlShop(texte) {
  const corps = String(texte).split(/\n{2,}/).map(function (p) {
    return '<p style="margin:0 0 12px">' + _echapperHtml(p).replace(/\n/g, '<br>') + '</p>';
  }).join('');
  return '<div style="font-family:Arial,sans-serif;max-width:560px;border:1px solid #eae7e1;border-radius:12px;overflow:hidden">' +
    '<div style="background:#000;color:#fff;font-weight:800;letter-spacing:4px;padding:14px 20px">13H59 SHOP</div>' +
    '<div style="padding:18px 20px 8px;font-size:14px;line-height:1.6;color:#0f0f0f">' + corps + '</div></div>';
}

/**
 * Envoie le modèle `code` à `to`. Renvoie true si envoyé, false si le modèle est absent ou désactivé.
 * Lève une erreur si l'envoi échoue (alias absent, quota…).
 */
function envoyerModele(ss, code, to, vars) {
  const m = _modelesEmails(ss)[code];
  if (!m || !m.actif || !m.texte) return false;
  const texte = _remplir(m.texte, vars);
  envoyerEmailShop({ to: to, subject: _remplir(m.objet, vars), texte: texte, html: _htmlShop(texte) });
  return true;
}

/** TEST — à lancer depuis l'éditeur. N'écrit à personne d'autre que toi. */
function testEnvoiAlias() {
  const moi = Session.getEffectiveUser().getEmail();
  Logger.log('Compte qui exécute le script : ' + moi);
  Logger.log('Alias déclarés dans ce Gmail : ' + JSON.stringify(GmailApp.getAliases()));
  if (!_aliasShopDisponible()) {
    Logger.log('❌ ' + EMAIL_SHOP + " n'apparaît pas dans les alias : vérifie Paramètres ▸ Comptes et importation.");
    return;
  }
  envoyerEmailShop({
    to: moi,
    subject: 'Test 13H59 · envoi depuis l’adresse du shop',
    texte: "Test d'envoi depuis l'alias du shop. Si tu lis ceci, l'envoi fonctionne.",
    html:
      '<div style="font-family:Arial,sans-serif;max-width:520px;border:1px solid #eae7e1;border-radius:12px;overflow:hidden">' +
        '<div style="background:#000;color:#fff;font-weight:800;letter-spacing:4px;padding:14px 20px">13H59</div>' +
        '<div style="padding:18px 20px;font-size:14px;line-height:1.6;color:#0f0f0f">' +
          '<p style="margin-top:0">Si tu lis ce message, l’envoi depuis l’adresse du shop fonctionne.</p>' +
          '<p>À vérifier :</p>' +
          '<ol style="padding-left:18px;margin:0">' +
            '<li>l’expéditeur affiché est <b>' + NOM_EXPEDITEUR + ' &lt;' + EMAIL_SHOP + '&gt;</b> ;</li>' +
            '<li>en cliquant sur « Répondre », le destinataire proposé est <b>' + EMAIL_SHOP + '</b> ;</li>' +
          '</ol>' +
        '</div>' +
      '</div>'
  });
  Logger.log('✅ E-mail de test envoyé à ' + moi + '.');
}
