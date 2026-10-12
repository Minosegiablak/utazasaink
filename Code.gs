/***** UTAZÁSAINK – Google Apps Script háttér *****
 * A jelszavak ITT vannak, nem a nyilvános HTML-ben.
 * Új felhasználó: írj be egy új sort a USERS listába.
 */
const USERS = {
  '345198G': 'Gábor',
  '345198R': 'Réka'
};

const SHEET_NAME = 'Adatok';
const HEAD = ['id', 'trip', 'type', 'author', 'updated', 'deleted', 'data'];
const META = ['id', 'trip', 'type', 'author', 'updated', 'deleted'];

function doGet() {
  return out_({ ok: true, msg: 'Utazásaink API működik' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return out_({ ok: false, error: 'Hibás kérés' }); }

  const user = USERS[String(req.code || '').trim().toUpperCase()];
  if (!user) return out_({ ok: false, auth: false, error: 'Hibás jelszó' });

  try {
    switch (req.action) {
      case 'login':    return out_({ ok: true, user: user });
      case 'list':     return out_({ ok: true, user: user, items: list_() });
      case 'save':     return out_({ ok: true, items: saveMany_([req.item], user) });
      case 'saveMany': return out_({ ok: true, items: saveMany_(req.items || [], user) });
      case 'delete':   delete_(req.id); return out_({ ok: true });
      case 'upload':   return out_({ ok: true, id: upload_(req) });
      case 'file':     return out_(Object.assign({ ok: true }, file_(req.id)));
      default:         return out_({ ok: false, error: 'Ismeretlen művelet' });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

// Ha a script táblázatból lett létrehozva, azt használja.
// Ha külön (önálló) script, első futáskor létrehoz egy "Utazásaink adatok" táblázatot a Drive-on.
function ss_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (ss) return ss;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  ss = SpreadsheetApp.create('Utazásaink adatok');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

// Opcionális: futtasd egyszer a szerkesztőből, hogy előre létrejöjjön a táblázat.
function setup() {
  const sh = sheet_();
  Logger.log('Táblázat kész: ' + sh.getParent().getUrl());
  Logger.log('Fájlok mappája: ' + folder_().getUrl());
}

/***** Beszállókártyák / jegyek a Google Drive-on *****/
function folder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder('Utazásaink fájlok');
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}

function upload_(req) {
  if (!req.data) throw new Error('Nincs fájl');
  const name = String(req.name || 'jegy.jpg').replace(/[\\/:*?"<>|]/g, '_');
  const blob = Utilities.newBlob(Utilities.base64Decode(req.data), req.mime || 'image/jpeg', name);
  return folder_().createFile(blob).getId();
}

function file_(id) {
  const f = DriveApp.getFileById(String(id));
  const folderId = folder_().getId();
  let ok = false;
  const ps = f.getParents();
  while (ps.hasNext()) { if (ps.next().getId() === folderId) { ok = true; break; } }
  if (!ok) throw new Error('Ez a fájl nem érhető el');
  const b = f.getBlob();
  return { mime: b.getContentType(), data: Utilities.base64Encode(b.getBytes()) };
}

function sheet_() {
  const ss = ss_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
  }
  return sh;
}

function list_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const rows = sh.getRange(2, 1, last - 1, HEAD.length).getValues();
  const items = [];
  rows.forEach(function (r) {
    if (!r[0] || r[5] === true || r[5] === 'TRUE') return;
    let d = {};
    try { d = JSON.parse(r[6] || '{}'); } catch (e) {}
    d.id = String(r[0]);
    d.trip = String(r[1]);
    d.type = String(r[2]);
    d.author = String(r[3]);
    d.updated = r[4] instanceof Date ? r[4].toISOString() : String(r[4]);
    items.push(d);
  });
  return items;
}

function saveMany_(items, user) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const last = sh.getLastRow();
    const ids = last > 1
      ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); })
      : [];
    const now = new Date();
    const out = [];

    items.forEach(function (it) {
      if (!it || !it.type) return;
      const id = String(it.id || Utilities.getUuid());
      const data = {};
      Object.keys(it).forEach(function (k) { if (META.indexOf(k) < 0) data[k] = it[k]; });

      const idx = ids.indexOf(id);
      let author = user;
      if (idx >= 0) {
        const row = idx + 2;
        author = sh.getRange(row, 4).getValue() || user; // az eredeti szerző marad
        sh.getRange(row, 1, 1, HEAD.length)
          .setValues([[id, it.trip || '', it.type, author, now, false, JSON.stringify(data)]]);
      } else {
        sh.appendRow([id, it.trip || '', it.type, author, now, false, JSON.stringify(data)]);
        ids.push(id);
      }
      out.push(Object.assign({}, data, {
        id: id, trip: it.trip || '', type: it.type, author: author, updated: now.toISOString()
      }));
    });
    return out;
  } finally {
    lock.releaseLock();
  }
}

function delete_(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const last = sh.getLastRow();
    if (last < 2) return;
    const ids = sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); });
    const idx = ids.indexOf(String(id));
    if (idx >= 0) {
      sh.getRange(idx + 2, 5).setValue(new Date());
      sh.getRange(idx + 2, 6).setValue(true); // csak megjelöljük, nem töröljük végleg
    }
  } finally {
    lock.releaseLock();
  }
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
