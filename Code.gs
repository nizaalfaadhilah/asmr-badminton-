// ASMR Badminton — backend Google Apps Script + Google Sheets
// GANTI PIN di bawah ini sebelum deploy. Siapa pun yang tahu PIN = admin.
const ADMIN_PIN = 'GANTI-PIN-INI';

const TAB = 'docs';
const COLLS = ['sessions', 'players', 'matches', 'live'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ASMR Badminton')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(TAB);
  if (!s) {
    s = ss.insertSheet(TAB);
    s.appendRow(['coll', 'id', 'json']);
  }
  return s;
}

// Dipanggil semua orang (member & admin): hanya membaca data.
function getAll() {
  const v = sheet_().getDataRange().getValues();
  const out = {};
  COLLS.forEach(function (c) { out[c] = []; });
  for (let i = 1; i < v.length; i++) {
    const c = v[i][0], id = v[i][1];
    if (!out[c] || !id) continue;
    try {
      out[c].push(Object.assign({ id: String(id) }, JSON.parse(v[i][2])));
    } catch (e) {}
  }
  return out;
}

function checkPin(p) {
  return !!p && String(p) === ADMIN_PIN && ADMIN_PIN !== 'GANTI-PIN-INI';
}

// Hanya admin (PIN benar) yang bisa menulis.
function write(pin, ops) {
  if (!checkPin(pin)) throw new Error('PIN salah');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const s = sheet_();
    const v = s.getDataRange().getValues();
    const m = {};
    for (let i = 1; i < v.length; i++) {
      if (v[i][0]) m[v[i][0] + '/' + v[i][1]] = v[i][2];
    }
    ops.forEach(function (o) {
      if (COLLS.indexOf(o.c) < 0 || !o.id) return;
      const k = o.c + '/' + o.id;
      if (o.op === 'delete') {
        delete m[k];
      } else if (o.op === 'set') {
        m[k] = JSON.stringify(o.data);
      } else if (o.op === 'update') {
        const cur = m[k] ? JSON.parse(m[k]) : {};
        m[k] = JSON.stringify(Object.assign(cur, o.data));
      }
    });
    const rows = Object.keys(m).map(function (k) {
      const i = k.indexOf('/');
      return [k.slice(0, i), k.slice(i + 1), m[k]];
    });
    s.clearContents();
    s.getRange(1, 1, 1, 3).setValues([['coll', 'id', 'json']]);
    if (rows.length) s.getRange(2, 1, rows.length, 3).setValues(rows);
    return true;
  } finally {
    lock.releaseLock();
  }
}
