/*************************************************************
 *  DOSSIERS DRIVE DES CRÉATEURS
 *  13h59 ▸ Créateurs ▸ « <nom> · <id> » ▸ Factures · Documents · Convention
 *   - créé avec la fiche (et à la demande pour les fiches existantes) ;
 *     son identifiant est gardé dans la colonne createurs.dossier_drive ;
 *   - renommé quand le nom de marque change ;
 *   - chaque facture reste dans 13h59 ▸ Factures ▸ AAAA-MM et reçoit un
 *     raccourci dans Factures du créateur (un seul fichier, deux accès) ;
 *   - le dossier « Créateurs » est partagé avec les comptes listés dans
 *     l'onglet acces_drive (Gestion ▸ Créateurs ▸ Dossiers Drive).
 *************************************************************/

const DOSSIER_CREATEURS = 'Créateurs';
const SOUS_DOSSIERS_CREATEUR = ['Factures', 'Documents', 'Convention'];
const SHEET_ACCES_DRIVE = 'acces_drive';
const COLONNES_ACCES_DRIVE = ['email', 'ajoute_le', 'ajoute_par'];

function _sousDossier(parent, nom) { const it = parent.getFoldersByName(nom); return it.hasNext() ? it.next() : parent.createFolder(nom); }
function _racineCreateurs() { return _sousDossier(_sousDossier(DriveApp.getRootFolder(), '13h59'), DOSSIER_CREATEURS); }
function _nomDossierCreateur(nom, id) { return String(nom || id).replace(/[\/\\]/g, '-').trim() + ' · ' + id; }
function _urlDossier(id) { return id ? 'https://drive.google.com/drive/folders/' + id : ''; }

/** Ajoute la colonne dossier_drive à createurs si elle manque ; renvoie le tableau à jour. */
function _createursAvecDossier(ss) {
  let t = _tableau(ss, SHEET_CREATEURS);
  if (t.M['dossier_drive'] == null) {
    _assurerTaille(t.sh, t.sh.getLastRow(), t.nbCol + 1);
    t.sh.getRange(1, t.nbCol + 1).setValue('dossier_drive').setFontWeight('bold');
    SpreadsheetApp.flush();
    t = _tableau(ss, SHEET_CREATEURS);
  }
  return t;
}

/** Dossier d'un créateur ; le crée (avec ses sous-dossiers) si `creer` et s'il n'existe pas ou a été supprimé. */
function _dossierCreateur(ss, id, creer) {
  const t = _createursAvecDossier(ss);
  const i = t.lignes.findIndex(function (r) { return String(_val(t, r, 'id_createur')) === String(id); });
  if (i < 0) throw new Error('Créateur introuvable : « ' + id + ' ».');
  const r = t.lignes[i], existant = String(_val(t, r, 'dossier_drive') || '');
  if (existant) {
    try { const d = DriveApp.getFolderById(existant); if (!d.isTrashed()) return d; } catch (e) { /* supprimé ou inaccessible : on recrée */ }
  }
  if (!creer) return null;
  const d = _racineCreateurs().createFolder(_nomDossierCreateur(_nomPropre(_val(t, r, 'nom')), id));
  SOUS_DOSSIERS_CREATEUR.forEach(function (n) { d.createFolder(n); });
  r[t.M['dossier_drive']] = d.getId();
  _ecrireLigne(t, i, r);
  return d;
}

/** Raccourci d'une facture dans Factures du créateur (sans doublon). */
function _raccourciFacture(ss, idCreateur, idPdf, nomPdf) {
  const d = _sousDossier(_dossierCreateur(ss, idCreateur, true), 'Factures');
  const it = d.getFiles();
  while (it.hasNext()) { const f = it.next(); try { if (f.getTargetId && f.getTargetId() === idPdf) return false; } catch (e) { /* pas un raccourci */ } }
  Drive.Files.create({ name: nomPdf, mimeType: 'application/vnd.google-apps.shortcut', shortcutDetails: { targetId: idPdf }, parents: [d.getId()] });
  return true;
}

/** Raccourcis des factures déjà faites d'un créateur. */
function _raccourcisFacturesExistantes(ss, idCreateur) {
  if (!ss.getSheetByName(SHEET_FACTURES)) return 0;
  let n = 0;
  _lireTable(_onglet(ss, SHEET_FACTURES)).forEach(function (f) {
    if (String(f['id_createur']) !== String(idCreateur) || !f['pdf_id'] || String(f['statut']) === 'annulee') return;
    try { const p = DriveApp.getFileById(String(f['pdf_id'])); if (!p.isTrashed() && _raccourciFacture(ss, idCreateur, p.getId(), p.getName())) n++; } catch (e) { /* PDF supprimé */ }
  });
  return n;
}

/* ---------- Actions de Gestion ---------- */

function _gDossierCreer(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), id = String(body.id || '');
  const d = _dossierCreateur(ss, id, true), n = _raccourcisFacturesExistantes(ss, id);
  _journaliser('dossier_drive', id + ' : dossier ' + d.getName() + (n ? ' · ' + n + ' facture(s) rangée(s)' : ''));
  return { ok: true, dossier: d.getId(), url: d.getUrl(), factures: n };
}

/** Crée les dossiers manquants des créateurs actifs ou déjà facturés (par lots pour rester sous la limite de temps). */
function _gDossiersCreerTous() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), t = _createursAvecDossier(ss), debut = Date.now();
  const factures = {};
  if (ss.getSheetByName(SHEET_FACTURES)) _lireTable(_onglet(ss, SHEET_FACTURES)).forEach(function (f) { factures[String(f['id_createur'])] = true; });
  const aFaire = t.lignes.filter(function (r) {
    const id = String(_val(t, r, 'id_createur'));
    return id && !_val(t, r, 'dossier_drive') && String(_val(t, r, 'categorie')) !== 'shop' && (_norm(_val(t, r, 'statut')) === 'actif' || factures[id]);
  }).map(function (r) { return String(_val(t, r, 'id_createur')); });
  const faits = [];
  for (let k = 0; k < aFaire.length && Date.now() - debut < 240000; k++) {
    _dossierCreateur(ss, aFaire[k], true); _raccourcisFacturesExistantes(ss, aFaire[k]); faits.push(aFaire[k]);
  }
  if (faits.length) _journaliser('dossier_drive', faits.length + ' dossier(s) créé(s) : ' + faits.join(', '));
  return { ok: true, crees: faits.length, restants: aFaire.length - faits.length };
}

function _gDossiersEtat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), t = _createursAvecDossier(ss);
  _creerOngletSiAbsent(ss, SHEET_ACCES_DRIVE, COLONNES_ACCES_DRIVE);
  const racine = _racineCreateurs();
  const manquants = t.lignes.filter(function (r) { return _val(t, r, 'id_createur') && !_val(t, r, 'dossier_drive') && _norm(_val(t, r, 'statut')) === 'actif' && String(_val(t, r, 'categorie')) !== 'shop'; }).length;
  const acces = _lireTable(_onglet(ss, SHEET_ACCES_DRIVE)).filter(function (a) { return a['email']; }).map(function (a) {
    return { email: String(a['email']), ajouteLe: _iso(a['ajoute_le']), ajoutePar: String(a['ajoute_par'] || '') };
  });
  return { ok: true, racine: racine.getUrl(), acces: acces, manquants: manquants };
}

function _gAccesDriveAjouter(body) {
  const email = _email(body.email);
  if (!email) throw new Error('Adresse e-mail invalide.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _creerOngletSiAbsent(ss, SHEET_ACCES_DRIVE, COLONNES_ACCES_DRIVE);
  const t = _tableau(ss, SHEET_ACCES_DRIVE);
  if (t.lignes.some(function (r) { return _norm(_val(t, r, 'email')) === email; })) throw new Error(email + ' a déjà accès.');
  _racineCreateurs().addEditor(email);
  _ajouterLigne(t, { email: email, ajoute_le: new Date(), ajoute_par: _signataire() });
  _journaliser('acces_drive', email + ' : accès aux dossiers des créateurs');
  return { ok: true };
}

function _gAccesDriveRetirer(body) {
  const email = _email(body.email), ss = SpreadsheetApp.getActiveSpreadsheet(), t = _tableau(ss, SHEET_ACCES_DRIVE);
  const i = t.lignes.findIndex(function (r) { return _norm(_val(t, r, 'email')) === email; });
  if (i < 0) throw new Error('Adresse introuvable.');
  try { _racineCreateurs().removeEditor(email); } catch (e) { /* déjà retiré dans Drive */ }
  t.sh.deleteRow(i + 2);
  _journaliser('acces_drive', email + ' : accès retiré');
  return { ok: true };
}

/*************************************************************
 *  DOCUMENTS DES CRÉATEURS — registre `documents` + fichier dans son dossier
 *  Ajoutés par une gérante depuis la fiche (validés d'office) ou, plus tard,
 *  par le créateur depuis son formulaire en ligne (« à vérifier »).
 *  Une attestation RC Pro validée coche « RC Pro » sur la fiche.
 *************************************************************/

const SHEET_DOCUMENTS = 'documents';
const COLONNES_DOCUMENTS = ['id_document', 'id_createur', 'type', 'nom_fichier', 'id_fichier', 'taille_ko', 'ajoute_le', 'ajoute_par', 'source',
  'statut', 'verifie_le', 'verifie_par', 'remarque'];
const TYPES_DOCUMENTS = [
  ['siret', 'Justificatif SIRET (avis INSEE ou extrait INPI)'], ['rc_pro', 'Attestation RC Pro'], ['rib', 'RIB'], ['identite', "Pièce d'identité"],
  ['demarche', 'Justificatif de démarche en cours (URSSAF, INPI, assureur)'], ['convention', 'Convention signée'], ['autre', 'Autre document']
];
const STATUTS_DOCUMENTS = ['a_verifier', 'valide', 'refuse'];
const MIMES_DOCUMENTS = ['application/pdf', 'image/jpeg', 'image/png', 'image/heic', 'image/heif', 'image/webp'];
const TAILLE_MAX_DOCUMENT = 10 * 1024 * 1024;

function _typeDocument(code) { const t = TYPES_DOCUMENTS.filter(function (x) { return x[0] === code; })[0]; return t ? t[1] : ''; }

/** Enregistre un fichier (base64) dans le dossier du créateur et l'inscrit au registre. `source` : 'gestion' ou 'formulaire'. */
function _ajouterDocument(ss, idCreateur, type, nom, mime, contenu, source, par, remarque) {
  if (!_typeDocument(type)) throw new Error('Type de document inconnu.');
  if (MIMES_DOCUMENTS.indexOf(String(mime)) < 0) throw new Error('Format non accepté : PDF ou photo (JPEG, PNG, HEIC) uniquement.');
  const octets = Utilities.base64Decode(String(contenu || ''));
  if (!octets.length) throw new Error('Fichier vide.');
  if (octets.length > TAILLE_MAX_DOCUMENT) throw new Error('Fichier trop lourd (10 Mo au plus).');
  const propre = String(nom || 'document').replace(/[\/\\<>:"|?*\x00-\x1f]/g, '-').slice(0, 120);
  const dossier = _sousDossier(_dossierCreateur(ss, idCreateur, true), type === 'convention' ? 'Convention' : 'Documents');
  const fichier = dossier.createFile(Utilities.newBlob(octets, mime, _typeDocument(type).split(' (')[0] + ' · ' + propre));
  _creerOngletSiAbsent(ss, SHEET_DOCUMENTS, COLONNES_DOCUMENTS);
  const t = _tableau(ss, SHEET_DOCUMENTS), id = _prochainId(t, 'id_document', 'DOC');
  const valide = source === 'gestion';
  _ajouterLigne(t, { id_document: id, id_createur: idCreateur, type: type, nom_fichier: propre, id_fichier: fichier.getId(), taille_ko: Math.ceil(octets.length / 1024),
    ajoute_le: new Date(), ajoute_par: par, source: source, statut: valide ? 'valide' : 'a_verifier', verifie_le: valide ? new Date() : '', verifie_par: valide ? par : '',
    remarque: _textePublic(remarque, 300) });
  if (valide && type === 'rc_pro') _cocherRcPro(ss, idCreateur);
  _journaliser('document_ajoute', idCreateur + ' · ' + id + ' ' + _typeDocument(type) + ' (' + propre + ')' + (valide ? '' : ' · à vérifier'), par);
  return { id: id, idFichier: fichier.getId() };
}
function _cocherRcPro(ss, idCreateur) {
  const t = _tableau(ss, SHEET_CREATEURS), i = t.lignes.findIndex(function (r) { return String(_val(t, r, 'id_createur')) === String(idCreateur); });
  if (i < 0 || _val(t, t.lignes[i], 'rc_pro') === true) return;
  const r = t.lignes[i]; r[t.M['rc_pro']] = true; if (t.M['modifie_le'] != null) r[t.M['modifie_le']] = new Date();
  _ecrireLigne(t, i, r);
}

function _gDocuments(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), id = String(body.idCreateur || '');
  _creerOngletSiAbsent(ss, SHEET_DOCUMENTS, COLONNES_DOCUMENTS);
  const docs = _lireTable(_onglet(ss, SHEET_DOCUMENTS)).filter(function (d) { return d['id_document'] && String(d['id_createur']) === id; }).map(function (d) {
    return { id: String(d['id_document']), type: String(d['type']), libelle: _typeDocument(String(d['type'])) || String(d['type']), nom: String(d['nom_fichier'] || ''),
      url: d['id_fichier'] ? 'https://drive.google.com/file/d/' + d['id_fichier'] + '/view' : '', tailleKo: Number(d['taille_ko']) || 0,
      ajouteLe: _iso(d['ajoute_le']), ajoutePar: String(d['ajoute_par'] || ''), source: String(d['source'] || ''), statut: String(d['statut'] || ''),
      verifieLe: _iso(d['verifie_le']), verifiePar: String(d['verifie_par'] || ''), remarque: String(d['remarque'] || '') };
  }).sort(function (a, b) { return b.ajouteLe.localeCompare(a.ajouteLe) || b.id.localeCompare(a.id); });
  return { ok: true, documents: docs, types: TYPES_DOCUMENTS };
}

function _gDocumentAjouter(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), id = String(body.idCreateur || '');
  if (!_createurs(ss).parId[id]) throw new Error('Créateur introuvable.');
  const r = _ajouterDocument(ss, id, String(body.type || ''), body.nom, body.mime, body.contenu, 'gestion', _signataire(), body.remarque);
  return { ok: true, id: r.id };
}

function _ligneDocument(t, id) {
  const i = t.lignes.findIndex(function (r) { return String(_val(t, r, 'id_document')) === String(id); });
  if (i < 0) throw new Error('Document introuvable.');
  return i;
}

function _gDocumentStatut(body) {
  if (STATUTS_DOCUMENTS.indexOf(body.statut) < 0) throw new Error('Statut inconnu.');
  const ss = SpreadsheetApp.getActiveSpreadsheet(), t = _tableau(ss, SHEET_DOCUMENTS), i = _ligneDocument(t, body.id), r = t.lignes[i];
  r[t.M['statut']] = body.statut; r[t.M['verifie_le']] = new Date(); r[t.M['verifie_par']] = _signataire();
  if (body.remarque != null) r[t.M['remarque']] = _textePublic(body.remarque, 300);
  _ecrireLigne(t, i, r);
  if (body.statut === 'valide' && String(_val(t, r, 'type')) === 'rc_pro') _cocherRcPro(ss, String(_val(t, r, 'id_createur')));
  _journaliser('document_statut', _val(t, r, 'id_createur') + ' · ' + body.id + ' → ' + body.statut);
  return { ok: true };
}

/** Retire un document du registre et met son fichier à la corbeille de Drive (récupérable 30 jours). */
function _gDocumentRetirer(body) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), t = _tableau(ss, SHEET_DOCUMENTS), i = _ligneDocument(t, body.id), r = t.lignes[i];
  try { DriveApp.getFileById(String(_val(t, r, 'id_fichier'))).setTrashed(true); } catch (e) { /* déjà supprimé dans Drive */ }
  const detail = _val(t, r, 'id_createur') + ' · ' + body.id + ' ' + _typeDocument(String(_val(t, r, 'type'))) + ' (' + _val(t, r, 'nom_fichier') + ')';
  PropertiesService.getScriptProperties().setProperty('DERNIER_ID_DOC', String(Math.max(_maxId(t, 'id_document', 'DOC'), _dernierId('DERNIER_ID_DOC'))));
  t.sh.deleteRow(i + 2);
  _journaliser('document_retire', detail + (body.motif ? ' : ' + _textePublic(body.motif, 200) : ''));
  return { ok: true };
}
