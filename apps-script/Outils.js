/*************************************************************
 *  OUTILS DE MAINTENANCE — à lancer depuis l'éditeur (▶ Exécuter)
 *   inventaireClasseur()  : chaque onglet, sa taille, ses formules, et s'il
 *                           est une table connue du code (lecture seule)
 *   controleIntegrite()   : identifiants vides ou en double, références
 *                           cassées, dates, e-mails et IBAN invalides
 *                           (lecture seule ; le journal ne cite que des
 *                           identifiants, jamais de données personnelles)
 *   sauvegarderClasseur() : copie datée du classeur dans Drive
 *                           « 13h59/Sauvegardes » (avant tout nettoyage)
 *  SCHEMA décrit les tables comme de futures tables Postgres : clé
 *  primaire, clés étrangères, colonnes de dates / e-mails / IBAN.
 *************************************************************/

const SCHEMA = {
  createurs:        { pk: 'id_createur', fk: { categorie: 'categories.code' }, emails: ['email'], ibans: ['iban'], dates: ['adhesion_payee_le', 'cree_le', 'modifie_le'] },
  categories:       { pk: 'code' },
  stands:           { pk: 'code' },
  remises:          { pk: 'code' },
  paiements:        { pk: 'code' },
  ventes:           { pk: 'id_vente', fk: { id_createur: 'createurs.id_createur', code_remise: 'remises.code', code_paiement: 'paiements.code', vente_origine: 'ventes.id_vente' }, dates: ['date'] },
  emplacements:     { pk: 'id_emplacement', fk: { id_createur: 'createurs.id_createur', code_stand: 'stands.code' }, dates: ['debut', 'fin', 'preavis_recu_le', 'echeance'] },
  candidatures:     { pk: 'id_candidature', fk: { id_createur: 'createurs.id_createur', categorie: 'categories.code', stand_souhaite: 'stands.code' }, emails: ['email'], dates: ['recue_le', 'maj_le', 'traitee_le'] },
  benevolat:        { pk: 'id_demande', fk: { id_createur: 'createurs.id_createur', stand: 'stands.code' }, emails: ['email'], dates: ['recue_le', 'maj_le', 'traitee_le'] },
  benevolat_disponibilites: { fk: { id_demande: 'benevolat.id_demande' } },
  createurs_emails_agenda:  { pk: 'email', fk: { id_createur: 'createurs.id_createur' }, emails: ['email'], dates: ['ajoute_le'] },
  emails:           { pk: 'code' },
  factures:         { pk: 'numero', fk: { id_createur: 'createurs.id_createur' }, dates: ['emise_le', 'envoyee_le'] },
  lignes_facture:   { fk: { numero: 'factures.numero' } },
  paiements_recus:  { pk: 'id_paiement', fk: { id_createur: 'createurs.id_createur' }, dates: ['recu_le', 'pointe_le'] },
  payeurs:          { pk: 'payeur', fk: { id_createur: 'createurs.id_createur' } },
  relances:         { fk: { id_createur: 'createurs.id_createur' }, dates: ['envoyee_le'] },
  versements:       { pk: 'id_versement', fk: { numero_facture: 'factures.numero', id_createur: 'createurs.id_createur' }, dates: ['fichier_le', 'verse_le'] },
  documents:        { pk: 'id_document', fk: { id_createur: 'createurs.id_createur' }, dates: ['ajoute_le', 'verifie_le'] },
  renouvellements:  { pk: 'id_renouvellement', fk: { id_emplacement: 'emplacements.id_emplacement', id_createur: 'createurs.id_createur' }, dates: ['echeance', 'envoye_le', 'date_limite', 'repondu_le'] },
  mandats:          { pk: 'rum', fk: { id_createur: 'createurs.id_createur' }, ibans: ['iban'], dates: ['demande_le', 'signe_le', 'premier_prelevement_le', 'revoque_le'] },
  prelevements:     { pk: 'id_prelevement', fk: { id_createur: 'createurs.id_createur', rum: 'mandats.rum', id_paiement: 'paiements_recus.id_paiement' }, dates: ['date_prelevement', 'prenotifie_le', 'fichier_le'] },
  utilisateurs:     { pk: 'email', emails: ['email'], dates: ['ajoute_le', 'derniere_connexion'] },
  acces_drive:      { pk: 'email', emails: ['email'], dates: ['ajoute_le'] },
  charges:          { pk: 'id_charge', dates: ['modifie_le'] },
  charges_mois:     { fk: { id_charge: 'charges.id_charge' } },
  recettes_mois:    {},
  exercices:        { pk: 'code' },
  parametres_bilan: { pk: 'cle' },
  journal:          {}
};
/** Onglets de travail connus mais hors base (reprise du stock, en attente). */
const ONGLETS_DE_TRAVAIL = ['Stock', 'Listes'];

function inventaireClasseur() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), lignes = [], inconnus = [];
  ss.getSheets().forEach(function (sh) {
    const nom = sh.getName(), nl = sh.getLastRow(), nc = sh.getLastColumn();
    let formules = 0, exemple = '';
    if (nl && nc) {
      const f = sh.getRange(1, 1, nl, nc).getFormulas();
      f.forEach(function (r) { r.forEach(function (x) { if (x) { formules++; if (!exemple) exemple = x.slice(0, 60); } }); });
    }
    const role = SCHEMA[nom] ? 'table' : ONGLETS_DE_TRAVAIL.indexOf(nom) !== -1 ? 'travail (stock)' : 'INCONNU DU CODE';
    if (!SCHEMA[nom] && ONGLETS_DE_TRAVAIL.indexOf(nom) === -1) inconnus.push(nom);
    const entetes = nc ? sh.getRange(1, 1, 1, nc).getValues()[0].map(String) : [];
    const pasSnake = entetes.filter(function (h) { return h && !/^[a-z][a-z0-9_]*$/.test(h); });
    lignes.push('• ' + nom + ' — ' + role + ' · ' + Math.max(0, nl - 1) + ' ligne(s) × ' + nc + ' col.' + (sh.isSheetHidden() ? ' · masqué' : '') +
      (formules ? ' · ' + formules + ' formule(s), ex. ' + exemple : '') +
      '\n    colonnes : ' + entetes.join(', ') + (pasSnake.length ? '\n    ⚠️ noms de colonnes hors norme (minuscules_sans_accents attendues) : ' + pasSnake.join(', ') : ''));
  });
  const absentes = Object.keys(SCHEMA).filter(function (t) { return !ss.getSheetByName(t); });
  Logger.log('ONGLETS (' + lignes.length + ') :\n' + lignes.join('\n'));
  Logger.log('Tables du code pas encore créées (normal si la fonction n\'a jamais servi) : ' + (absentes.join(', ') || 'aucune'));
  Logger.log('Onglets inconnus du code, candidats au nettoyage : ' + (inconnus.join(', ') || 'aucun'));
}

function controleIntegrite() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), tables = {}, pbs = [];
  Object.keys(SCHEMA).forEach(function (t) { if (ss.getSheetByName(t)) tables[t] = _lireTable(_onglet(ss, t)); });
  const cles = {};
  Object.keys(SCHEMA).forEach(function (t) {
    const s = SCHEMA[t];
    if (!tables[t] || !s.pk) return;
    const vus = {}, doublons = [];
    let vides = 0;
    tables[t].forEach(function (r) {
      const k = String(r[s.pk] == null ? '' : r[s.pk]).trim();
      if (!k) { vides++; return; }
      if (vus[k]) doublons.push(k); vus[k] = true;
    });
    cles[t + '.' + s.pk] = vus;
    if (vides) pbs.push(t + ' : ' + vides + ' ligne(s) sans ' + s.pk);
    if (doublons.length) pbs.push(t + ' : ' + s.pk + ' en double → ' + doublons.slice(0, 15).join(', ') + (doublons.length > 15 ? '…' : ''));
  });
  const ident = function (t, r) { const pk = SCHEMA[t].pk; return pk ? String(r[pk]) : '(ligne sans clé)'; };
  Object.keys(SCHEMA).forEach(function (t) {
    const s = SCHEMA[t], L = tables[t];
    if (!L) return;
    Object.keys(s.fk || {}).forEach(function (col) {
      const cible = cles[s.fk[col]];
      if (!cible) return;
      const casses = L.filter(function (r) { const v = String(r[col] == null ? '' : r[col]).trim(); return v && !cible[v]; });
      if (casses.length) pbs.push(t + '.' + col + ' → ' + s.fk[col] + ' : ' + casses.length + ' référence(s) introuvable(s), ex. ' +
        casses.slice(0, 8).map(function (r) { return ident(t, r) + ' (' + String(r[col]).slice(0, 20) + ')'; }).join(', '));
    });
    (s.dates || []).forEach(function (col) {
      const mauvais = L.filter(function (r) { const v = r[col]; return v !== '' && v != null && !(v instanceof Date); });
      if (mauvais.length) pbs.push(t + '.' + col + ' : ' + mauvais.length + ' valeur(s) qui ne sont pas des dates, ex. ' + mauvais.slice(0, 8).map(function (r) { return ident(t, r); }).join(', '));
    });
    (s.emails || []).forEach(function (col) {
      const mauvais = L.filter(function (r) { const v = String(r[col] || '').trim(); return v && !_email(v); });
      if (mauvais.length) pbs.push(t + '.' + col + ' : ' + mauvais.length + ' e-mail(s) invalide(s), ex. ' + mauvais.slice(0, 8).map(function (r) { return ident(t, r); }).join(', '));
    });
    (s.ibans || []).forEach(function (col) {
      const mauvais = L.filter(function (r) { const v = String(r[col] || '').replace(/\s/g, '').toUpperCase(); return v && !_ibanValide(v); });
      if (mauvais.length) pbs.push(t + '.' + col + ' : ' + mauvais.length + ' IBAN invalide(s), ex. ' + mauvais.slice(0, 8).map(function (r) { return ident(t, r); }).join(', '));
    });
  });
  // règles métier utiles avant la migration
  if (tables.emplacements) {
    const parC = {};
    tables.emplacements.forEach(function (e) { (parC[String(e['id_createur'])] = parC[String(e['id_createur'])] || []).push(e); });
    Object.keys(parC).forEach(function (id) {
      const L = parC[id].filter(function (e) { return e['debut'] instanceof Date; }).sort(function (a, b) { return a['debut'] - b['debut']; });
      for (let k = 1; k < L.length; k++) {
        const f = L[k - 1]['fin'];
        if (!(f instanceof Date) || f >= L[k]['debut']) pbs.push('emplacements : ' + id + ' a deux emplacements qui se chevauchent (' + L[k - 1]['id_emplacement'] + ', ' + L[k]['id_emplacement'] + ')');
      }
    });
  }
  Logger.log(pbs.length ? '⚠️ ' + pbs.length + ' point(s) à regarder :\n• ' + pbs.join('\n• ') : '✅ Aucune incohérence trouvée.');
  Logger.log('Tables contrôlées : ' + Object.keys(tables).map(function (t) { return t + ' (' + tables[t].length + ')'; }).join(', '));
}

/** Restes de la migration de septembre 2026 (inventaire du 30/09 : aucune ligne de code ne les lit). */
const ONGLETS_MIGRATION = ['All data', '_ancien_ventes', 'vendeurs', '_correspondance_noms', '_controle_migration', '_revue_enrichissement'];

function apercuNettoyageOnglets() { _nettoyerOnglets(true); }
/** Sauvegarde le classeur, puis supprime les onglets de ONGLETS_MIGRATION présents. */
function nettoyerOnglets() { _nettoyerOnglets(false); }
function _nettoyerOnglets(apercu) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const L = ONGLETS_MIGRATION.filter(function (n) { return ss.getSheetByName(n); });
  if (!L.length) { Logger.log('Rien à supprimer : ces onglets n\'existent plus.'); return; }
  const code = L.filter(function (n) { return SCHEMA[n]; });
  if (code.length) throw new Error('Refusé : ' + code.join(', ') + ' est une table du code.');
  if (apercu) {
    Logger.log('APERÇU — seraient supprimés (après une copie de sauvegarde) :\n• ' + L.map(function (n) { return n + ' (' + Math.max(0, ss.getSheetByName(n).getLastRow() - 1) + ' lignes)'; }).join('\n• ') +
      '\nRestent : ' + ss.getSheets().map(function (s) { return s.getName(); }).filter(function (n) { return L.indexOf(n) === -1; }).join(', '));
    return;
  }
  sauvegarderClasseur();
  L.forEach(function (n) { ss.deleteSheet(ss.getSheetByName(n)); });
  _journaliser('nettoyage_onglets', 'Supprimés après sauvegarde : ' + L.join(', '), Session.getEffectiveUser().getEmail());
  Logger.log('✅ ' + L.length + ' onglet(s) supprimés : ' + L.join(', '));
}

/**
 * Mise aux normes (30/09/2026) : les colonnes à plusieurs valeurs deviennent des tables.
 *  benevolat.disponibilites → benevolat_disponibilites ; createurs.emails_agenda → createurs_emails_agenda.
 *  createurs.rc_pro (case en double avec le registre documents) : retirée, les cases sans attestation sont listées.
 *  Les valeurs existantes sont recopiées, puis la colonne est retirée.
 */
function apercuNormalisation() { _normaliser(true); }
function normaliserTables() { _normaliser(false); }
function _normaliser(apercu) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), faits = [];
  const colonne = function (nom, col) { const sh = ss.getSheetByName(nom); if (!sh || !sh.getLastColumn()) return null; const i = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String).indexOf(col); return i < 0 ? null : { sh: sh, i: i }; };
  const b = colonne(SHEET_BENEVOLAT, 'disponibilites');
  if (b) {
    const L = _lireTable(_onglet(ss, SHEET_BENEVOLAT)).filter(function (r) { return r['id_demande'] && r['disponibilites']; });
    faits.push('benevolat.disponibilites → ' + SHEET_BENEVOLAT_DISPOS + ' (' + L.length + ' demande(s) à recopier), puis colonne retirée');
    if (!apercu) { L.forEach(function (r) { _ecrireDisponibilites(ss, String(r['id_demande']), _disponibilites(r['disponibilites'])); }); b.sh.deleteColumn(b.i + 1); }
  }
  const c = colonne(SHEET_CREATEURS, 'emails_agenda');
  if (c) {
    const L = _lireTable(_onglet(ss, SHEET_CREATEURS)).filter(function (r) { return r['id_createur'] && r['emails_agenda']; });
    faits.push('createurs.emails_agenda → ' + SHEET_EMAILS_AGENDA + ' (' + L.length + ' fiche(s) à recopier), puis colonne retirée');
    if (!apercu) {
      _creerOngletSiAbsent(ss, SHEET_EMAILS_AGENDA, COLONNES_EMAILS_AGENDA);
      const t = _tableau(ss, SHEET_EMAILS_AGENDA);
      L.forEach(function (r) { String(r['emails_agenda']).split(/[\s,;]+/).map(_norm).filter(function (m) { return m.indexOf('@') > 0; })
        .forEach(function (m) { _ajouterLigne(t, { email: m, id_createur: String(r['id_createur']), ajoute_le: new Date(), ajoute_par: 'normalisation' }); }); });
      c.sh.deleteColumn(c.i + 1);
    }
  }
  const rc = colonne(SHEET_CREATEURS, 'rc_pro');
  if (rc) {
    // la case ne sert plus : « RC Pro » = attestation validée dans `documents`. On liste les cases cochées sans attestation.
    const ok = {};
    _lireSi(ss, SHEET_DOCUMENTS).forEach(function (d) { if (String(d['type']) === 'rc_pro' && String(d['statut']) === 'valide') ok[String(d['id_createur'])] = true; });
    const sans = _lireTable(_onglet(ss, SHEET_CREATEURS)).filter(function (r) { return r['id_createur'] && r['rc_pro'] === true && !ok[String(r['id_createur'])]; });
    faits.push('createurs.rc_pro retirée (la RC Pro se lit dans documents). Case cochée sans attestation dans le registre (' + sans.length + ') : ' +
      (sans.map(function (r) { return _nomPropre(r['nom']) + ' (' + r['id_createur'] + ')'; }).join(', ') || 'aucune') + ' — attestation à déposer dans leur fiche ▸ Documents');
    if (!apercu) rc.sh.deleteColumn(rc.i + 1);
  }
  if (!faits.length) { Logger.log('✅ Rien à faire : les tables sont déjà aux normes.'); return; }
  if (apercu) { Logger.log('APERÇU :\n• ' + faits.join('\n• ')); return; }
  _journaliser('normalisation', faits.join(' ; '), Session.getEffectiveUser().getEmail());
  Logger.log('✅ Fait :\n• ' + faits.join('\n• '));
}

function sauvegarderClasseur() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dossier = _sousDossier(_sousDossier(DriveApp.getRootFolder(), '13h59'), 'Sauvegardes');
  const nom = ss.getName() + ' — sauvegarde ' + Utilities.formatDate(new Date(), _tz(), 'yyyy-MM-dd HH-mm');
  const copie = DriveApp.getFileById(ss.getId()).makeCopy(nom, dossier);
  Logger.log('✅ Copie créée : « ' + nom + ' » dans 13h59/Sauvegardes (' + copie.getUrl() + '). Elle contient des données personnelles : ne la partage pas.');
}

/**
 * Bascule vers Athena — à lancer depuis l'éditeur JUSTE AVANT de télécharger le classeur en .xlsx.
 * Recopie dans l'onglet `_proprietes` les valeurs gardées hors du Sheet (propriétés du script) dont Athena a besoin :
 * les derniers identifiants attribués (DERNIER_ID_*, pour ne jamais réattribuer un numéro) et l'ICS s'il est saisi.
 * Les mots de passe, l'IBAN et le BIC de l'association ne sont PAS recopiés (ils sont dans le .env du serveur).
 */
function exporterProprietesPourAthena() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), props = PropertiesService.getScriptProperties().getProperties();
  const lignes = Object.keys(props).filter(function (k) { return /^DERNIER_ID_[A-Z]+$/.test(k) || k === 'ASSO_ICS'; }).sort()
    .map(function (k) { return [k, String(props[k])]; });
  const sh = ss.getSheetByName('_proprietes') || ss.insertSheet('_proprietes');
  sh.clear();
  sh.getRange(1, 1, 1, 2).setValues([['cle', 'valeur']]).setFontWeight('bold');
  sh.getRange(1, 1, lignes.length + 1, 2).setNumberFormat('@');
  if (lignes.length) sh.getRange(2, 1, lignes.length, 2).setValues(lignes);
  Logger.log('✅ ' + lignes.length + ' valeur(s) recopiée(s) dans l\'onglet _proprietes : ' + lignes.map(function (l) { return l[0] + ' = ' + l[1]; }).join(', ') +
    '. Télécharge maintenant le classeur : Fichier ▸ Télécharger ▸ Microsoft Excel (.xlsx).');
}
