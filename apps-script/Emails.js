/*************************************************************
 *  E-MAILS — envoi au nom du shop
 *  Tous les e-mails automatiques partent de l'alias du shop, avec
 *  le shop en « Répondre à » et en copie cachée (trace dans sa boîte).
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
    bcc: EMAIL_SHOP,
    htmlBody: m.html,
    attachments: m.pieces || []
  });
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
            '<li>une copie de ce message est arrivée dans la boîte du shop.</li>' +
          '</ol>' +
        '</div>' +
      '</div>'
  });
  Logger.log('✅ E-mail de test envoyé à ' + moi + ' (copie cachée : ' + EMAIL_SHOP + ').');
}
