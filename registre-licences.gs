/**
 * Registre des licences – Verhaeghe Irrigation
 * Un code d'activation = un seul téléphone.
 *
 * Feuille "Licences" :
 *   A Nom | B Empreinte du code | C Actif (OUI/NON) | D Téléphone | E Modèle | F Activé le | G Dernière vérification
 *
 *  - Désactiver une personne : mettre NON dans la colonne Actif.
 *  - Libérer un code (changement de téléphone) : vider la case Téléphone (colonne D).
 *  - Nouveau code : menu « Licences » > « Créer un code ».
 */
const FEUILLE = 'Licences';
const ALPHA = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';

const CODES_EXISTANTS = [
    ["Code 01", "00d2f722424b03d16dc94280c90afbed800382b47f493a783fda01814dfa2a9b"],
    ["Code 02", "283d308e272891b39a5128aff0dfd39719ce84fc6cda094cd69572bf054802a8"],
    ["Code 03", "a1e53bc9cebb6d030102d1cf81bdcdcc595cdbd48bfa2d425054a60936aef616"],
    ["Code 04", "97c002db35273d1336dc3ba0cc0b443a7469cff70113e76d71df72d44020389f"],
    ["Code 05", "b577455078954bcd5acffe713b851d7867dd90f7b1f84dc450a9be47202f93b5"],
    ["Code 06", "ca547de7b795c180d82e28977b38fa363d837b8c8e7f34c89d6c57c484f525d7"],
    ["Code 07", "56cd17219a4c27caad21ea0edc129d66b2284f4e5a12038d7eefd0b65a999d41"],
    ["Code 08", "1ec23356b700a1ddfd8ce200fe4a2a23db70917b28d549a6257ae10dadd78be8"],
    ["Code 09", "dae6a6017e93ef433c58a19256f3d4a77fc977180d32577ef295c37f05acab49"],
    ["Code 10", "31519c3f22c2bba5d428ba6d70e2955b8fc94a0bfa2fe9f76ad8322f5d32345c"]
];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Licences')
    .addItem('Créer un code', 'creerCode')
    .addItem('Installer / réparer la feuille', 'installer')
    .addToUi();
}

/** À lancer une seule fois : crée la feuille et y met les codes déjà distribués. */
function installer() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(FEUILLE);
  if (!sh) sh = ss.insertSheet(FEUILLE);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Nom', 'Empreinte du code', 'Actif', 'Téléphone', 'Modèle', 'Activé le', 'Dernière vérification']);
    sh.setFrozenRows(1);
    sh.getRange('A1:G1').setFontWeight('bold').setBackground('#e6f0fd');
  }
  const deja = new Set(sh.getRange(1, 2, sh.getLastRow(), 1).getValues().map(r => r[0]));
  CODES_EXISTANTS.forEach(([n, h]) => { if (!deja.has(h)) sh.appendRow([n, h, 'OUI', '', '', '', '']); });
  sh.hideColumns(2);
  sh.autoResizeColumns(1, 7);
  const v = SpreadsheetApp.newDataValidation().requireValueInList(['OUI', 'NON'], true).build();
  sh.getRange(2, 3, Math.max(1, sh.getMaxRows() - 1), 1).setDataValidation(v);
}

function norm_(c) { return String(c || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }

function empreinte_(code) {
  let b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'verhaeghe-licence:' + norm_(code), Utilities.Charset.UTF_8);
  for (let i = 0; i < 20000; i++) b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, b);
  return b.map(x => ('0' + ((x + 256) % 256).toString(16)).slice(-2)).join('');
}

function creerCode() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Nouveau code', 'Nom de la personne :', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const nom = r.getResponseText().trim() || 'Sans nom';
  const hex = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  let s = '';
  for (let i = 0; i < 12; i++) s += ALPHA[parseInt(hex.substr(i * 2, 2), 16) % 32];
  const code = 'VHG-' + s.slice(0, 4) + '-' + s.slice(4, 8) + '-' + s.slice(8, 12);
  SpreadsheetApp.getActive().getSheetByName(FEUILLE).appendRow([nom, empreinte_(code), 'OUI', '', '', '', '']);
  ui.alert('Code pour ' + nom, code + '\n\nNotez-le maintenant : il ne sera plus affiché (seule son empreinte est gardée).', ui.ButtonSet.OK);
}

/** Appelé par l'application. */
function doGet(e) {
  let res;
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); res = traiter_(e.parameter || {}); }
  catch (x) { res = { ok: false, err: 'serveur' }; }
  finally { try { lock.releaseLock(); } catch (x) {} }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

function traiter_(p) {
  const sh = SpreadsheetApp.getActive().getSheetByName(FEUILLE);
  const dev = String(p.dev || '').slice(0, 80);
  if (!dev) return { ok: false, err: 'serveur' };
  const h = p.a === 'act' ? empreinte_(p.code) : String(p.h || '');
  const data = sh.getDataRange().getValues();
  let i = -1;
  for (let k = 1; k < data.length; k++) if (data[k][1] === h) { i = k; break; }
  if (i < 0) return { ok: false, err: 'inconnu' };
  const row = data[i], ligne = i + 1, maintenant = new Date();
  if (String(row[2]).toUpperCase() !== 'OUI') return { ok: false, err: 'desactive' };
  const tel = String(row[3] || '');
  if (p.a === 'act' || (p.a === 'chk' && p.mig === '1')) {
    if (tel && tel !== dev) return { ok: false, err: 'autre_tel' };
    if (!tel) sh.getRange(ligne, 4, 1, 3).setValues([[dev, String(p.model || ''), maintenant]]);
  } else if (p.a === 'chk') {
    if (!tel) return { ok: false, err: 'libere' };
    if (tel !== dev) return { ok: false, err: 'autre_tel' };
  } else return { ok: false, err: 'serveur' };
  sh.getRange(ligne, 7).setValue(maintenant);
  return { ok: true, n: String(row[0] || '') };
}
