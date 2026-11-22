/**
 * HRPCI PR CRM - Google Apps Script backend
 * Put this inside your Google Sheet (Extensions > Apps Script), run setup() once, then Deploy > Web app.
 * The sheet stays your database. The mobile app reads and writes the PR 1..PR 5 tabs.
 */
const DATA_ROWS = 400;     // rows per PR tab (rows 8 to 407 in your sheet)
const NCOLS = 25;          // A..Y
const USERS_SHEET = 'AppUsers';
// Input columns only (formula columns like Chat Link, Alert, Balance are never touched)
const COL = {date:2,name:3,phone:4,city:6,field:7,award:8,status:9,priority:10,calls:11,
  firstCalled:12,interestedOn:13,lastContact:14,nextFollowup:15,fee:18,received:19,
  payDate:21,payMode:22,docs:23,remarks:25};
const DATE_KEYS = ['date','firstCalled','interestedOn','lastContact','nextFollowup','payDate'];
const NUM_KEYS = ['calls','fee','received'];

/* ---------- ONE-TIME SETUP: run this once from the editor ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SECRET')) props.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  let sh = ss.getSheetByName(USERS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(USERS_SHEET);
    sh.getRange('C:C').setNumberFormat('@');
    sh.appendRow(['ID', 'Name', 'PIN', 'Role', 'Tab', 'Active (Y/N)']);
    sh.setFrozenRows(1);
    const set = ss.getSheetByName('Setup');
    const rows = set ? set.getRange(1, 1, 30, 3).getValues() : [];
    rows.forEach(function (r) {
      const id = String(r[0]).trim().toUpperCase();
      if (/^PR\d+$/.test(id)) sh.appendRow([id, r[1], randPin_(4), 'exec', 'PR ' + parseInt(id.replace(/\D/g, ''), 10), 'Y']);
    });
    sh.appendRow(['ADMIN', 'Admin', randPin_(6), 'admin', '', 'Y']);
    sh.getRange('A1:F1').setFontWeight('bold').setBackground('#14284b').setFontColor('#ffffff');
    sh.setColumnWidths(1, 6, 140);
  }
  SpreadsheetApp.getUi().alert('Setup done. Open the "' + USERS_SHEET + '" tab to see each executive\'s PIN. Now deploy as a Web app.');
}
function randPin_(n) { let s = ''; for (let i = 0; i < n; i++) s += Math.floor(Math.random() * 10); return s; }

/* ---------- HTTP entry points ---------- */
function doGet() { return ContentService.createTextOutput('HRPCI CRM API is running.'); }
function doPost(e) {
  let out;
  try { out = route_(JSON.parse((e.postData && e.postData.contents) || '{}')); }
  catch (err) { out = {ok: false, error: String((err && err.message) || err)}; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function route_(r) {
  if (r.action === 'users') return {ok: true, users: getUsers_().filter(function (u) { return u.active; }).map(function (u) { return {id: u.id, name: u.name, role: u.role}; })};
  if (r.action === 'login') return login_(r);
  const u = authUser_(r.token);
  if (!u) return {ok: false, error: 'auth'};
  if (r.action === 'data') return data_(u);
  if (r.action === 'save') return save_(u, r);
  return {ok: false, error: 'Unknown action'};
}

/* ---------- Users and login ---------- */
function getUsers_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USERS_SHEET);
  if (!sh) throw new Error('Run setup() first');
  const v = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues() : [];
  return v.filter(function (r) { return r[0]; }).map(function (r) {
    return {id: String(r[0]).trim(), name: String(r[1]).trim(), pin: String(r[2]).trim(), role: String(r[3]).trim().toLowerCase(),
      tab: String(r[4]).trim(), active: String(r[5]).trim().toUpperCase() !== 'N'};
  });
}
function pub_(u) { return {id: u.id, name: u.name, role: u.role, tab: u.tab}; }
function login_(r) {
  const cache = CacheService.getScriptCache(), key = 'fail_' + r.id;
  if (Number(cache.get(key) || 0) >= 6) return {ok: false, error: 'Too many wrong tries. Wait 15 minutes.'};
  const u = getUsers_().filter(function (x) { return x.id === r.id && x.active; })[0];
  if (!u || u.pin !== String(r.pin || '').trim()) {
    cache.put(key, String(Number(cache.get(key) || 0) + 1), 900);
    return {ok: false, error: 'Wrong PIN'};
  }
  cache.remove(key);
  return {ok: true, token: makeToken_(u.id), user: pub_(u)};
}
function sign_(s) {
  const key = PropertiesService.getScriptProperties().getProperty('SECRET');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(s, key));
}
function makeToken_(id) {
  const p = Utilities.base64EncodeWebSafe(Utilities.newBlob(JSON.stringify({id: id, exp: Date.now() + 30 * 864e5})).getBytes());
  return p + '.' + sign_(p);
}
function authUser_(t) {
  if (!t || String(t).indexOf('.') < 0) return null;
  const parts = String(t).split('.');
  if (sign_(parts[0]) !== parts[1]) return null;
  const o = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString());
  if (o.exp < Date.now()) return null;
  return getUsers_().filter(function (u) { return u.id === o.id && u.active; })[0] || null;
}

/* ---------- Reading ---------- */
function SS_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function headerRow_(sh) {
  const v = sh.getRange(1, COL.name, 20, 1).getValues();
  for (let i = 0; i < v.length; i++) if (String(v[i][0]).trim() === 'Lead Name') return i + 1;
  return 7;
}
function dstr_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const s = String(v || '').trim();
  return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0, 10) : '';
}
function rowObj_(v, x, row, tz) {
  const o = {tab: x.tab, row: row, exec: x.name, execId: x.id};
  Object.keys(COL).forEach(function (k) {
    const val = v[COL[k] - 1];
    if (DATE_KEYS.indexOf(k) >= 0) o[k] = dstr_(val, tz);
    else if (NUM_KEYS.indexOf(k) >= 0) o[k] = Number(val) || 0;
    else if (k === 'phone') o[k] = String(val).replace(/\D/g, '').slice(-10);
    else o[k] = String(val === null || val === undefined ? '' : val).trim();
  });
  return o;
}
function readTab_(ss, x, tz) {
  const sh = ss.getSheetByName(x.tab);
  if (!sh) return [];
  const hr = headerRow_(sh), vals = sh.getRange(hr + 1, 1, DATA_ROWS, NCOLS).getValues(), out = [];
  vals.forEach(function (v, i) {
    if (!String(v[COL.name - 1]).trim() && !String(v[COL.phone - 1]).trim()) return;
    if (/EXAMPLE ROW/i.test(String(v[NCOLS - 1]))) return;
    out.push(rowObj_(v, x, hr + 1 + i, tz));
  });
  return out;
}
function lists_(ss) {
  const sh = ss.getSheetByName('Setup'), res = {status: [], award: [], field: [], priority: [], docs: [], payMode: []};
  if (!sh) return res;
  const map = {'Status (fixed)': 'status', 'Award Interested In': 'award', 'Field of Work': 'field',
    'Priority (fixed)': 'priority', 'Documents (fixed)': 'docs', 'Payment Mode': 'payMode'};
  const all = sh.getRange(1, 1, 40, sh.getLastColumn()).getValues();
  for (let r = 0; r < all.length; r++) {
    let found = false;
    for (let c = 0; c < all[r].length; c++) {
      const key = map[String(all[r][c]).trim()];
      if (!key) continue;
      found = true;
      for (let i = r + 1; i < all.length && i < r + 21; i++) {
        const t = String(all[i][c]).trim();
        if (!t || t.length > 40) break;
        res[key].push(t);
      }
    }
    if (found) break;
  }
  return res;
}
function targets_(ss) {
  const sh = ss.getSheetByName('Setup'), t = {};
  if (!sh) return t;
  sh.getRange(1, 1, 30, 5).getValues().forEach(function (r) { if (/^PR\d+$/i.test(String(r[0]).trim())) t[String(r[0]).trim().toUpperCase()] = Number(r[4]) || 0; });
  return t;
}
function data_(u) {
  const ss = SS_(), tz = ss.getSpreadsheetTimeZone(), tg = targets_(ss);
  const execs = getUsers_().filter(function (x) { return x.active && x.role === 'exec' && (u.role === 'admin' || x.id === u.id); });
  let leads = [];
  execs.forEach(function (x) { leads = leads.concat(readTab_(ss, x, tz)); });
  return {ok: true, user: pub_(u), lists: lists_(ss), leads: leads, today: Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'),
    execs: execs.map(function (x) { return {id: x.id, name: x.name, tab: x.tab, target: tg[x.id.toUpperCase()] || 0}; })};
}

/* ---------- Writing ---------- */
function findDup_(ss, users, phone, tz) {
  const p = String(phone).replace(/\D/g, '').slice(-10);
  if (p.length < 10) return null;
  const ex = users.filter(function (x) { return x.role === 'exec' && x.active; });
  for (let i = 0; i < ex.length; i++) {
    const hit = readTab_(ss, ex[i], tz).filter(function (l) { return l.phone === p; })[0];
    if (hit) return {name: hit.name, exec: ex[i].name, tab: ex[i].tab};
  }
  return null;
}
function save_(u, r) {
  const ss = SS_(), tz = ss.getSpreadsheetTimeZone(), users = getUsers_(), f = r.fields || {};
  let target = u;
  if (u.role === 'admin') target = users.filter(function (x) { return x.role === 'exec' && x.tab === r.tab; })[0];
  else if (r.tab && r.tab !== u.tab) return {ok: false, error: 'Not allowed'};
  if (!target) return {ok: false, error: 'Unknown executive sheet'};
  const sh = ss.getSheetByName(target.tab);
  if (!sh) return {ok: false, error: 'Tab not found: ' + target.tab};
  const hr = headerRow_(sh), lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    let row = Number(r.row) || 0;
    if (!row) {
      if (!f.name || !f.phone) return {ok: false, error: 'Name and phone are required'};
      if (!r.force) { const d = findDup_(ss, users, f.phone, tz); if (d) return {ok: true, dup: d}; }
      const nv = sh.getRange(hr + 1, COL.name, DATA_ROWS, 2).getValues();
      for (let i = 0; i < nv.length; i++) if (!String(nv[i][0]).trim() && !String(nv[i][1]).trim()) { row = hr + 1 + i; break; }
      if (!row) return {ok: false, error: 'This sheet is full (' + DATA_ROWS + ' leads)'};
    } else if (row <= hr || row > hr + DATA_ROWS) return {ok: false, error: 'Bad row'};
    Object.keys(f).forEach(function (k) {
      if (!COL[k]) return;
      const v = f[k], cell = sh.getRange(row, COL[k]);
      if (v === '' || v === null || v === undefined) cell.clearContent();
      else if (DATE_KEYS.indexOf(k) >= 0) cell.setValue(Utilities.parseDate(String(v), tz, 'yyyy-MM-dd'));
      else if (NUM_KEYS.indexOf(k) >= 0) cell.setValue(Number(v) || 0);
      else if (k === 'phone') cell.setValue(Number(String(v).replace(/\D/g, '').slice(-10)));
      else cell.setValue(String(v));
    });
    SpreadsheetApp.flush();
    return {ok: true, lead: rowObj_(sh.getRange(row, 1, 1, NCOLS).getValues()[0], target, row, tz)};
  } finally { lock.releaseLock(); }
}
