/*************************************************************
 *  À VERSER — ventes reversées aux créateurs (Collectif → créateurs)
 *  Pour chaque facture générée : net à verser = prime des ventes du mois
 *  précédent − commission (montant figé sur la facture). Le versement part :
 *   - seulement si le loyer du mois de la facture est reçu (convention) ;
 *   - seulement si l'IBAN de la fiche est valide.
 *  Un versement bloqué reste proposé les mois suivants jusqu'à ce qu'il parte.
 *  Le site ne fait AUCUN paiement : il prépare un fichier de virements groupés
 *  (SEPA, pain.001.001.03) à importer dans l'espace pro de la banque, puis on
 *  indique « validé à la banque » (ou « versé » pour un virement fait à la main).
 *   - versements : une ligne par virement préparé (statut fichier / verse / annule)
 *************************************************************/

const SHEET_VERSEMENTS = 'versements';
const COLONNES_VERSEMENTS = ['id_versement', 'numero_facture', 'id_createur', 'mois', 'montant', 'statut', 'fichier_le', 'verse_le', 'lot', 'par', 'remarque'];
const ASSO_BANQUE = { iban: 'FR7630027172150002142520125', bic: 'CMCIFRPP', nom: 'COLLECTIF 13H59' };

function _ongletVersements(ss) {
  _creerOngletSiAbsent(ss, SHEET_VERSEMENTS, COLONNES_VERSEMENTS);
  const sh = ss.getSheetByName(SHEET_VERSEMENTS);
  sh.getRange(1, 1, sh.getMaxRows(), 4).setNumberFormat('@');   // id, numero, créateur, mois
  sh.getRange(1, 9, sh.getMaxRows(), 1).setNumberFormat('@');   // lot
}

/** Toutes les lignes « à verser » jusqu'au mois donné (factures générées, net > 0), avec leur état. */
function _lignesAVerser(ss, mois) {
  const auj = _aujourdhui();
  const fiches = {};
  _lireTable(_onglet(ss, SHEET_CREATEURS)).forEach(function (c) {
    const iban = String(c['iban'] || '').replace(/\s/g, '').toUpperCase();
    fiches[String(c['id_createur'])] = { nom: _nomPropre(c['nom']), nomLegal: String(c['nom_legal'] || ''), iban: iban, ibanOk: !!iban && _ibanValide(iban), email: String(c['email'] || '') };
  });
  const recus = {};
  _lireSi(ss, SHEET_PAIEMENTS_RECUS).forEach(function (p) {
    if (String(p['statut']) === 'annule') return;
    const k = String(p['id_createur']) + '|' + _moisTexte(p['mois_loyer']);
    recus[k] = _round2((recus[k] || 0) + (Number(p['montant']) || 0));
  });
  const vers = {};
  _lireSi(ss, SHEET_VERSEMENTS).forEach(function (v) {
    if (String(v['statut']) === 'annule') return;
    vers[_numeroTexte(v['numero_facture'])] = { id: String(v['id_versement']), statut: String(v['statut']), fichierLe: _iso(v['fichier_le']), verseLe: _iso(v['verse_le']), lot: String(v['lot'] || '') };
  });
  return _lireSi(ss, SHEET_FACTURES).filter(function (f) {
    return String(f['statut']) !== 'annulee' && _moisTexte(f['mois']) <= mois && (Number(f['net_a_verser']) || 0) > 0;
  }).map(function (f) {
    const num = _numeroTexte(f['numero']), id = String(f['id_createur']), fm = _moisTexte(f['mois']), fi = fiches[id] || { nom: id, nomLegal: '', iban: '', ibanOk: false, email: '' };
    const loyer = Number(f['loyer']) || 0, recu = recus[id + '|' + fm] || 0, loyerOk = loyer <= 0.009 || recu >= loyer - 0.009;
    const v = vers[num];
    const statut = v && v.statut === 'verse' ? 'verse' : v && v.statut === 'fichier' ? 'fichier' : !fi.ibanOk ? 'iban' : !loyerOk ? (auj > _echeance(fm) ? 'reporte' : 'attente') : 'pret';
    return {
      numero: num, mois: fm, idCreateur: id, nom: fi.nom, nomLegal: fi.nomLegal, ibanFin: fi.iban ? fi.iban.slice(-4) : '', ibanOk: fi.ibanOk,
      ventes: Number(f['ventes_mois_precedent']) || 0, frais: Number(f['frais_mois_precedent']) || 0, commission: Number(f['commission']) || 0,
      net: Number(f['net_a_verser']) || 0, loyer: loyer, loyerRecu: recu, loyerOk: loyerOk, statut: statut, versement: v || null, _iban: fi.iban
    };
  }).filter(function (l) { return l.mois === mois || l.statut !== 'verse'; })   // les anciens déjà versés ne sont plus affichés
    .sort(function (a, b) { return a.nom.localeCompare(b.nom, 'fr', { sensitivity: 'base' }) || a.mois.localeCompare(b.mois); });
}

function _gVersements(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), mois = String(body.mois || '');
  const lignes = _lignesAVerser(ss, mois).map(function (l) { delete l._iban; return l; });
  // Créateurs avec des ventes le mois précédent mais sans facture générée pour ce mois : rien ne peut partir.
  const ctx = _contexteFactures(ss, mois), sansFacture = [];
  _aFacturer(ctx).forEach(function (r) {
    const id = String(_val(ctx.tC, r, 'id_createur'));
    if (!ctx.factures[id] && ctx.ventes[id]) sansFacture.push(_nomPropre(_val(ctx.tC, r, 'nom')));
  });
  return { ok: true, mois: mois, libelle: _moisBornes(mois).libelle, moisVentes: _moisBornes(_moisPrecedent(mois)).libelle, aujourdhui: _iso(_aujourdhui()), lignes: lignes, sansFacture: sansFacture };
}

/** Texte accepté dans un fichier SEPA : lettres sans accents, chiffres et / - ? : ( ) . , ' + espace. */
function _texteSepa(s, max) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Za-z0-9\/\-?:().,'+ ]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, max);
}
function _xml(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

/** Prépare le fichier de virements groupés pour les lignes « prêtes » choisies ; elles passent « dans un fichier ». */
function _gVersementsFichier(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), mois = String(body.mois || '');
  _ongletVersements(ss);
  const voulus = body.numeros || [];
  const L = _lignesAVerser(ss, mois).filter(function (l) { return voulus.indexOf(l.numero) !== -1 && l.statut === 'pret'; });
  if (!L.length) throw new Error('Aucun versement prêt parmi ceux choisis (loyer reçu et IBAN valide).');
  const maintenant = new Date(), lot = 'LOT-' + Utilities.formatDate(maintenant, _tz(), 'yyyyMMdd-HHmmss');
  const total = _round2(L.reduce(function (s, l) { return s + l.net; }, 0)), montant = function (n) { return _round2(n).toFixed(2); };
  const libMois = function (m) { return MOIS_FR[+m.slice(5, 7) - 1] + ' ' + m.slice(0, 4); };
  const tx = L.map(function (l) {
    return '<CdtTrfTxInf><PmtId><EndToEndId>' + _xml(_texteSepa(l.numero, 35)) + '</EndToEndId></PmtId>' +
      '<Amt><InstdAmt Ccy="EUR">' + montant(l.net) + '</InstdAmt></Amt>' +
      '<Cdtr><Nm>' + _xml(_texteSepa(l.nomLegal || l.nom, 70)) + '</Nm></Cdtr>' +
      '<CdtrAcct><Id><IBAN>' + _xml(l._iban) + '</IBAN></Id></CdtrAcct>' +
      '<RmtInf><Ustrd>' + _xml(_texteSepa(('13H59 VENTES ' + libMois(_moisPrecedent(l.mois)) + ' FACTURE ' + l.numero).toUpperCase(), 140)) + '</Ustrd></RmtInf></CdtTrfTxInf>';
  }).join('');
  const xml = '<?xml version="1.0" encoding="UTF-8"?>\n<Document xmlns="urn:iso:std:iso:20022:tech:xsd:pain.001.001.03" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    '<CstmrCdtTrfInitn><GrpHdr><MsgId>' + lot + '</MsgId><CreDtTm>' + Utilities.formatDate(maintenant, _tz(), "yyyy-MM-dd'T'HH:mm:ss") + '</CreDtTm>' +
    '<NbOfTxs>' + L.length + '</NbOfTxs><CtrlSum>' + montant(total) + '</CtrlSum><InitgPty><Nm>' + ASSO_BANQUE.nom + '</Nm></InitgPty></GrpHdr>' +
    '<PmtInf><PmtInfId>' + lot + '</PmtInfId><PmtMtd>TRF</PmtMtd><BtchBookg>false</BtchBookg><NbOfTxs>' + L.length + '</NbOfTxs><CtrlSum>' + montant(total) + '</CtrlSum>' +
    '<PmtTpInf><SvcLvl><Cd>SEPA</Cd></SvcLvl></PmtTpInf><ReqdExctnDt>' + _jourIso(maintenant) + '</ReqdExctnDt>' +
    '<Dbtr><Nm>' + ASSO_BANQUE.nom + '</Nm></Dbtr><DbtrAcct><Id><IBAN>' + ASSO_BANQUE.iban + '</IBAN></Id></DbtrAcct>' +
    '<DbtrAgt><FinInstnId><BIC>' + ASSO_BANQUE.bic + '</BIC></FinInstnId></DbtrAgt><ChrgBr>SLEV</ChrgBr>' + tx + '</PmtInf></CstmrCdtTrfInitn></Document>\n';

  const tV = _tableau(ss, SHEET_VERSEMENTS);
  let n = _maxId(tV, 'id_versement', 'V');
  const lignes = L.map(function (l) {
    n++;
    return ['V' + String(n).padStart(4, '0'), l.numero, l.idCreateur, l.mois, l.net, 'fichier', maintenant, '', lot, _signataire(), ''];
  });
  const der = tV.sh.getLastRow();
  _assurerTaille(tV.sh, der + lignes.length, COLONNES_VERSEMENTS.length);
  tV.sh.getRange(der + 1, 1, lignes.length, COLONNES_VERSEMENTS.length).setValues(lignes);
  _journaliser('versements_fichier', lot + ' · ' + L.length + ' virement(s) · ' + _eurFr(total) + ' : ' + L.map(function (l) { return l.nom + ' ' + _eurFr(l.net); }).join(', '));
  return { ok: true, lot: lot, nb: L.length, total: total, xml: xml, nom: '13H59-virements-' + lot.slice(4) + '.xml',
           detail: L.map(function (l) { return { nom: l.nomLegal || l.nom, ibanFin: l._iban.slice(-4), montant: l.net, libelle: _texteSepa(('13H59 VENTES ' + libMois(_moisPrecedent(l.mois)) + ' FACTURE ' + l.numero).toUpperCase(), 140) }; }) };
}

/** « J'ai validé les virements à la banque » (un lot) ou « Marquer versé » (une facture, virement fait à la main). */
function _gVersementsValider(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _ongletVersements(ss);
  const tV = _tableau(ss, SHEET_VERSEMENTS), maintenant = new Date(), faits = [];
  if (body.lot) {
    tV.lignes.forEach(function (r, i) {
      if (String(_val(tV, r, 'lot')) === String(body.lot) && String(_val(tV, r, 'statut')) === 'fichier') {
        r[tV.M['statut']] = 'verse'; r[tV.M['verse_le']] = maintenant; _ecrireLigne(tV, i, r); faits.push(_numeroTexte(_val(tV, r, 'numero_facture')));
      }
    });
  } else if (body.numero) {
    const l = _lignesAVerser(ss, String(body.mois || '9999-12')).filter(function (x) { return x.numero === String(body.numero); })[0];
    if (!l) throw new Error('Facture introuvable ou rien à verser : « ' + body.numero + ' ».');
    if (l.statut === 'verse') throw new Error('Déjà versé.');
    if (l.statut === 'fichier') {
      const i = tV.lignes.findIndex(function (r) { return _numeroTexte(_val(tV, r, 'numero_facture')) === l.numero && String(_val(tV, r, 'statut')) === 'fichier'; });
      const r = tV.lignes[i]; r[tV.M['statut']] = 'verse'; r[tV.M['verse_le']] = maintenant; _ecrireLigne(tV, i, r);
    } else {
      if (!l.loyerOk && body.malgreLoyer !== true) throw new Error('Le loyer de ' + l.nom + ' n\'est pas reçu : le versement est normalement reporté.');
      _ajouterLigne(tV, { id_versement: 'V' + String(_maxId(tV, 'id_versement', 'V') + 1).padStart(4, '0'), numero_facture: l.numero, id_createur: l.idCreateur, mois: l.mois,
        montant: l.net, statut: 'verse', verse_le: maintenant, lot: 'MANUEL', par: _signataire(), remarque: l.loyerOk ? '' : 'versé malgré le loyer non reçu' });
    }
    faits.push(l.numero);
  }
  if (!faits.length) throw new Error('Rien à valider.');
  _journaliser('versements_valides', (body.lot ? body.lot + ' · ' : 'à la main · ') + faits.join(', '));
  return { ok: true, nb: faits.length };
}

/** Un fichier jamais importé à la banque : les virements repassent « prêts ». */
function _gVersementsAnnulerLot(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), tV = _tableau(ss, SHEET_VERSEMENTS), motif = _motif(body);
  let n = 0;
  tV.lignes.forEach(function (r, i) {
    if (String(_val(tV, r, 'lot')) === String(body.lot) && String(_val(tV, r, 'statut')) === 'fichier') { r[tV.M['statut']] = 'annule'; r[tV.M['remarque']] = motif; _ecrireLigne(tV, i, r); n++; }
  });
  if (!n) throw new Error('Aucun virement en attente dans ce fichier.');
  _journaliser('versements_lot_annule', body.lot + ' · ' + n + ' virement(s) : ' + motif);
  return { ok: true, nb: n };
}
