/**
 * Inventory Helper: Google Sheet backend.
 *
 * Paste this into the sheet's Extensions > Apps Script editor and deploy it as a
 * web app (Execute as: Me, Who has access: Anyone). The page talks to that URL.
 *
 * Rows live on a tab named "Inventory". Photos are saved as JPEG files in a Drive
 * folder named "Inventory Helper Photos" and linked from the Photo column.
 */

// Set a passcode here, in your own copy only. Anyone who has the web app URL and
// this passcode can read, change and clear the count. Empty means no passcode.
const PASSCODE = '';

const TAB = 'Inventory';
const HEAD = ['Name', 'Qty', 'Counted at', 'Photo', 'ID', 'Created', 'Updated', 'File ID', 'Thumb'];
const C = { name: 0, qty: 1, at: 2, link: 3, id: 4, created: 5, updated: 6, file: 7, thumb: 8 };

function doGet() {
  return out_({ ok: true, service: 'inventory-helper' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ error: 'bad_request' }); }
  if (PASSCODE && req.key !== PASSCODE) return out_({ error: 'auth' });
  try {
    switch (req.action) {
      case 'list': return out_({ items: list_() });
      case 'thumbs': return out_({ thumbs: thumbs_(req.ids || []) });
      case 'photo': return out_({ data: photo_(req.id) });
      case 'upsert': return out_(locked_(function () { return upsert_(req.item, req.photo); }));
      case 'delete': return out_(locked_(function () { return remove_(req.id); }));
      case 'clear': return out_(locked_(clear_));
    }
    return out_({ error: 'unknown_action' });
  } catch (err) {
    return out_({ error: 'server', message: String(err) });
  }
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function locked_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(TAB);
  if (!sh) {
    sh = ss.insertSheet(TAB);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(2, C.at + 1, sh.getMaxRows() - 1, 1).setNumberFormat('m/d/yyyy h:mm am/pm');
    sh.getRange(1, C.thumb + 1, sh.getMaxRows(), 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    sh.setColumnWidth(C.name + 1, 260);
    sh.setColumnWidth(C.at + 1, 160);
    sh.hideColumns(C.id + 1, HEAD.length - C.id); // bookkeeping columns the page needs
  }
  return sh;
}

function rows_(sh) {
  const last = sh.getLastRow();
  return last < 2 ? [] : sh.getRange(2, 1, last - 1, HEAD.length).getValues();
}

function folder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (err) {} }
  const f = DriveApp.createFolder('Inventory Helper Photos');
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}

// Keep typed names from being read as formulas.
function text_(s) {
  s = String(s == null ? '' : s).slice(0, 200);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function list_() {
  return rows_(sheet_()).filter(function (r) { return r[C.id]; }).map(function (r) {
    return {
      id: String(r[C.id]), name: String(r[C.name]), qty: Number(r[C.qty]) || 0,
      created: Number(r[C.created]) || 0, updated: Number(r[C.updated]) || 0, hasPhoto: !!r[C.file],
    };
  });
}

function thumbs_(ids) {
  const want = {}, got = {};
  ids.forEach(function (id) { want[id] = true; });
  rows_(sheet_()).forEach(function (r) { if (want[r[C.id]]) got[r[C.id]] = String(r[C.thumb] || ''); });
  return got;
}

function photo_(id) {
  const row = rows_(sheet_()).filter(function (r) { return r[C.id] === id; })[0];
  if (!row || !row[C.file]) return '';
  const blob = DriveApp.getFileById(row[C.file]).getBlob();
  return 'data:image/jpeg;base64,' + Utilities.base64Encode(blob.getBytes());
}

function upsert_(item, photo) {
  if (!item || !item.id) throw new Error('missing item id');
  const sh = sheet_(), rows = rows_(sh);
  let i = -1;
  for (let k = 0; k < rows.length; k++) if (rows[k][C.id] === item.id) { i = k; break; }
  const old = i >= 0 ? rows[i] : null;

  let fileId = old ? old[C.file] : '', link = old ? old[C.link] : '';
  if (photo && !fileId) {
    const bytes = Utilities.base64Decode(String(photo).split(',').pop());
    const file = folder_().createFile(Utilities.newBlob(bytes, 'image/jpeg', item.id + '.jpg'));
    fileId = file.getId(); link = file.getUrl();
  }

  // A newer edit from another device wins; still keep a photo that just arrived.
  const stale = old && Number(old[C.updated]) > Number(item.updated);
  const thumb = item.thumb && String(item.thumb).length < 49000 ? item.thumb : (old ? old[C.thumb] : '');
  const row = stale
    ? [old[C.name], old[C.qty], old[C.at], link, old[C.id], old[C.created], old[C.updated], fileId, old[C.thumb] || thumb]
    : [text_(item.name), Number(item.qty) || 0, new Date(Number(item.created)), link, item.id,
       Number(item.created), Number(item.updated), fileId, thumb];
  if (old) sh.getRange(i + 2, 1, 1, HEAD.length).setValues([row]);
  else sh.appendRow(row);
  return { ok: true, hasPhoto: !!fileId, stale: !!stale };
}

function remove_(id) {
  const sh = sheet_(), rows = rows_(sh);
  for (let k = rows.length - 1; k >= 0; k--) {
    if (rows[k][C.id] !== id) continue;
    trash_(rows[k][C.file]);
    sh.deleteRow(k + 2);
  }
  return { ok: true };
}

function clear_() {
  const sh = sheet_(), rows = rows_(sh);
  rows.forEach(function (r) { trash_(r[C.file]); });
  if (rows.length) sh.getRange(2, 1, rows.length, HEAD.length).clearContent();
  return { ok: true, cleared: rows.length };
}

function trash_(fileId) {
  if (!fileId) return;
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (err) {}
}

/** Run this once from the editor (Run > setup) to grant Sheets and Drive access. */
function setup() {
  sheet_();
  folder_();
}
