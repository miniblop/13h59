/*************************************************************
 *  CORRECTION PONCTUELLE — doublons du 12/09/2026 (paniers 1671 et 1672)
 *
 *  Le panier 1670 a été enregistré 3 fois (réponse serveur perdue puis
 *  revalidation). On supprime les 2 copies : ventes 3183 à 3188.
 *
 *  Sécurités :
 *   - les lignes sont retrouvées par numero_vente, pas par numéro de ligne ;
 *   - le contenu de chaque ligne est vérifié avant toute suppression ;
 *   - au moindre écart, RIEN n'est supprimé et le journal dit pourquoi ;
 *   - relancer après coup ne fait rien (lignes introuvables → arrêt).
 *  Annulation possible via Fichier ▸ Historique des versions.
 *************************************************************/

function supprimerDoublonsPanier1670() {
  var ATTENDU = {                       // numero_vente : [panier, vendeur, référence, prix, paiement]
    3183: [1671, 'Nomi',         'sticker', 1.5, 'espèces'],
    3184: [1671, 'Dream club',   'pc',      2,   'espèces'],
    3185: [1671, 'Studio naimé', 'sticker', 2.5, 'espèces'],
    3186: [1672, 'Nomi',         'sticker', 1.5, 'espèces'],
    3187: [1672, 'Dream club',   'pc',      2,   'espèces'],
    3188: [1672, 'Studio naimé', 'sticker', 2.5, 'espèces']
  };
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);                 // pas de vente de la caisse pendant la correction
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ventes');
    var h = _headerMap(sh).map;
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();

    var lignes = [];                    // numéros de ligne Sheet à supprimer
    data.forEach(function (r, i) {
      var nv = Number(r[h['numero_vente']]);
      if (!ATTENDU[nv]) return;
      var a = ATTENDU[nv];
      var ok = Number(r[h['numero_panier']]) === a[0]
        && String(r[h['vendeur']]).trim() === a[1]
        && String(r[h['reference_produit']]).trim() === a[2]
        && Number(r[h['prix']]) === a[3]
        && String(r[h['type_de_paiement']]).trim() === a[4];
      if (!ok) throw new Error('Vente ' + nv + ' (ligne ' + (i + 2) + ') ne correspond pas au contenu attendu : ' + JSON.stringify(r.slice(0, 8)));
      lignes.push(i + 2);
    });

    if (lignes.length !== 6) throw new Error(lignes.length + ' ligne(s) trouvée(s) au lieu de 6 — déjà corrigé ? Rien supprimé.');

    // Suppression du bas vers le haut pour ne pas décaler les lignes restantes.
    lignes.sort(function (x, y) { return y - x; }).forEach(function (n) { sh.deleteRow(n); });
    SpreadsheetApp.flush();
    Logger.log('✅ 6 lignes supprimées (lignes ' + lignes.reverse().join(', ') + '). Paniers 1671 et 1672 retirés, 1670 conservé.');
  } catch (e) {
    Logger.log('⛔ ' + e.message);
  } finally {
    lock.releaseLock();
  }
}
