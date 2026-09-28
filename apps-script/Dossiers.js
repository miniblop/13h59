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
