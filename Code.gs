/**
 * Timenest backend (Google Apps Script + Google Sheet)
 * Sheets are created automatically: accounts, appts, blocks, notes
 * Each sheet has two columns: id | data (JSON)
 */

function getSS_() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  // standalone script: create/remember a spreadsheet automatically
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('SSID');
  if (id) return SpreadsheetApp.openById(id);
  ss = SpreadsheetApp.create('Timenest data');
  p.setProperty('SSID', ss.getId());
  return ss;
}
const ALLOWED = ['accounts', 'appts', 'blocks', 'notes'];

function sheet_(name) {
  const SS = getSS_();
  let s = SS.getSheetByName(name);
  if (!s) { s = SS.insertSheet(name); s.appendRow(['id', 'data']); }
  return s;
}

function rows_(col) {
  const v = sheet_(col).getDataRange().getValues(), out = [];
  for (let i = 1; i < v.length; i++) {
    if (v[i][0] === '') continue;
    let d = {};
    try { d = JSON.parse(v[i][1] || '{}'); } catch (e) {}
    out.push({ row: i + 1, id: String(v[i][0]), data: d });
  }
  return out;
}

function find_(col, id) { return rows_(col).find(r => r.id === id) || null; }

function hash_(salt, pw) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + pw, Utilities.Charset.UTF_8);
  return b.map(x => ('0' + (x & 0xff).toString(16)).slice(-2)).join('');
}

const pub_ = d => d ? { username: d.username, name: d.name, role: d.role } : null;
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

// ---- Run this ONCE from the editor to create the MD and Checker accounts ----
function setup() {
  [['md', 'MD Sir', 'md', 'md123'], ['checker', 'Checker', 'checker', 'checker123']].forEach(a => {
    if (!find_('accounts', a[0])) {
      const salt = Utilities.getUuid();
      sheet_('accounts').appendRow([a[0], JSON.stringify({ username: a[0], name: a[1], role: a[2], salt: salt, hash: hash_(salt, a[3]) })]);
    }
  });
  ['appts', 'blocks', 'notes'].forEach(sheet_);
}

function doGet() { return out_({ ok: true }); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const o = JSON.parse(e.postData.contents);
    return out_(handle_(o));
  } catch (err) {
    return out_({ error: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function handle_(o) {
  const op = o.op, col = o.col;

  // ----- auth -----
  if (op === 'user') {
    const r = find_('accounts', String(o.username || '').toLowerCase());
    return { user: r ? pub_(r.data) : null };
  }
  if (op === 'login') {
    const r = find_('accounts', String(o.username || '').toLowerCase());
    if (!r || r.data.hash !== hash_(r.data.salt, String(o.password || ''))) return { error: 'BAD_LOGIN' };
    return { user: pub_(r.data) };
  }
  if (op === 'signup') {
    const u = String(o.username || '').toLowerCase(), n = String(o.name || '').trim(), p = String(o.password || '');
    if (!n || !/^[a-z0-9_]{3,20}$/.test(u) || p.length < 6) return { error: 'INVALID' };
    if (find_('accounts', u)) return { error: 'TAKEN' };
    const salt = Utilities.getUuid(), a = { username: u, name: n, role: 'emp', salt: salt, hash: hash_(salt, p) };
    sheet_('accounts').appendRow([u, JSON.stringify(a)]);
    return { user: pub_(a) };
  }
  if (op === 'changePw') {
    const r = find_('accounts', String(o.username || '').toLowerCase());
    if (!r || r.data.hash !== hash_(r.data.salt, String(o.old || ''))) return { error: 'BAD_PW' };
    if (String(o.password || '').length < 6) return { error: 'INVALID' };
    const salt = Utilities.getUuid();
    r.data.salt = salt; r.data.hash = hash_(salt, o.password);
    sheet_('accounts').getRange(r.row, 2).setValue(JSON.stringify(r.data));
    return { ok: true };
  }

  // ----- generic data -----
  if (ALLOWED.indexOf(col) < 0) return { error: 'BAD_COLLECTION' };
  const isAcc = col === 'accounts';
  const safe = d => { if (!isAcc) return d; const c = Object.assign({}, d); delete c.salt; delete c.hash; return c; };

  if (op === 'get') {
    if (isAcc) return { data: null };
    const r = find_(col, o.id);
    return { data: r ? r.data : null };
  }
  if (op === 'list') {
    let l = rows_(col);
    if (o.where && o.where.length === 2) l = l.filter(r => r.data[o.where[0]] === o.where[1]);
    return { docs: l.map(r => ({ id: r.id, data: safe(r.data) })) };
  }

  // writes: accounts can only be changed through login/signup/changePw
  if (isAcc) return { error: 'FORBIDDEN' };

  if (op === 'add') {
    const id = Utilities.getUuid();
    sheet_(col).appendRow([id, JSON.stringify(o.data || {})]);
    return { id: id };
  }
  if (op === 'set') {
    const r = find_(col, o.id);
    if (r) sheet_(col).getRange(r.row, 2).setValue(JSON.stringify(o.data || {}));
    else sheet_(col).appendRow([o.id, JSON.stringify(o.data || {})]);
    return { ok: true };
  }
  if (op === 'update') {
    const r = find_(col, o.id);
    if (!r) return { error: 'NOT_FOUND' };
    sheet_(col).getRange(r.row, 2).setValue(JSON.stringify(Object.assign({}, r.data, o.data || {})));
    return { ok: true };
  }
  return { error: 'BAD_OP' };
}
