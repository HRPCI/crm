/**
 * HRPCI PR CRM - Google Apps Script backend  (UPDATED)
 * NEW: WhatsApp leads inbox ("WA Leads" sheet) + round-robin + duplicate check + call-log endpoint
 *
 * Steps:
 * 1. Replace your old code with this file.
 * 2. Run setup() only if AppUsers does not exist. Then run setupWA() once (shows your API key).
 * 3. Deploy > Manage deployments > Edit (pencil) > Version: New version > Deploy.
 *    (Web app: Execute as Me, Who has access: Anyone)
 */
const DATA_ROWS = 400;     // rows per PR tab (rows 8 to 407 in your sheet)
const NCOLS = 25;          // A..Y
const USERS_SHEET = 'AppUsers';
const WA_SHEET = 'WA Leads';     // ONE inbox sheet for all new WhatsApp leads
const CALL_SHEET = 'CallLog';    // every call (agent phone -> server)
const AUTO_ASSIGN = true;        // true = round-robin; false = leave "Assigned To" blank and you assign by hand
const UPDATE_PR_TABS = true;     // call-log also updates Calls Made / First Called / Last Contact in PR tabs
const WA_HEAD = ['Received', 'Name', 'Phone', 'First Message', 'Assigned To', 'Moved to PR Tab?',
  'Calls', 'Last Call', 'Last Duration (sec)', 'Total Talk (sec)', 'Last Call Result'];
const CALL_HEAD = ['Time', 'Phone', 'Agent', 'Type', 'Duration (sec)', 'Duration (mm:ss)', 'Result', 'Lead Name', 'Lead Tab'];
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

/* ---------- ONE-TIME SETUP for WhatsApp leads + call log: run once ---------- */
function setupWA() {
  const ss = SS_(), p = PropertiesService.getScriptProperties();
  if (!p.getProperty('SECRET')) p.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  if (!p.getProperty('API_KEY')) p.setProperty('API_KEY', Utilities.getUuid().replace(/-/g, ''));
  if (!p.getProperty('VERIFY_TOKEN')) p.setProperty('VERIFY_TOKEN', 'hrpci' + Utilities.getUuid().slice(0, 8));
  waSheet_(ss); callSheet_(ss); refreshAgents();
  const msg = 'API KEY: ' + p.getProperty('API_KEY') + '\nMeta VERIFY TOKEN: ' + p.getProperty('VERIFY_TOKEN') +
    '\n\nWebhook URL = <your web app URL>?key=' + p.getProperty('API_KEY');
  Logger.log(msg);
  SpreadsheetApp.getUi().alert('Setup done.\n\n' + msg);
}
function onOpen() {
  SpreadsheetApp.getUi().createMenu('HRPCI CRM')
    .addItem('Send assigned leads to PR tabs', 'moveAssignedToPR')
    .addItem('Refresh agent dropdown', 'refreshAgents')
    .addToUi();
}
// Fills the "Assigned To" dropdown with active executives (run again if you add/remove users)
function refreshAgents() {
  const sh = waSheet_(SS_());
  const names = getUsers_().filter(function (x) { return x.role === 'exec' && x.active; }).map(function (x) { return x.name; });
  if (!names.length) return;
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(names, true).setAllowInvalid(false).build();
  sh.getRange(2, 5, 2000, 1).setDataValidation(rule);
}

/* ---------- HTTP entry points ---------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  // Meta webhook verification
  if (p['hub.mode'] === 'subscribe' && p['hub.verify_token'] &&
      p['hub.verify_token'] === PropertiesService.getScriptProperties().getProperty('VERIFY_TOKEN'))
    return ContentService.createTextOutput(p['hub.challenge']);
  return ContentService.createTextOutput('HRPCI CRM API is running.');
}
function doPost(e) {
  let out;
  try {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (e.parameter && e.parameter.key && !body.key) body.key = e.parameter.key;
    out = body.object === 'whatsapp_business_account' ? handleMeta_(body) : route_(body);
  } catch (err) { out = {ok: false, error: String((err && err.message) || err)}; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function route_(r) {
  if (r.action === 'users') return {ok: true, users: getUsers_().filter(function (u) { return u.active; }).map(function (u) { return {id: u.id, name: u.name, role: u.role}; })};
  if (r.action === 'login') return login_(r);
  // Server-to-server actions (secured by API key)
  if (r.action === 'wa_lead') return keyOk_(r.key) ? addLead_(r) : {ok: false, error: 'auth'};
  if (r.action === 'calllog') return keyOk_(r.key) ? callLog_(r) : {ok: false, error: 'auth'};
  const u = authUser_(r.token);
  if (!u) return {ok: false, error: 'auth'};
  if (r.action === 'data') return data_(u);
  if (r.action === 'save') return save_(u, r);
  return {ok: false, error: 'Unknown action'};
}
function keyOk_(k) {
  const real = PropertiesService.getScriptProperties().getProperty('API_KEY');
  return !!real && String(k || '') === real;
}

/* ---------- WhatsApp leads: inbox + duplicate check + round-robin ---------- */
// Accepts Meta's raw webhook payload (POST to <web app URL>?key=API_KEY)
function handleMeta_(r) {
  if (!keyOk_(r.key)) return {ok: false, error: 'auth'};
  let added = 0, dups = 0;
  (r.entry || []).forEach(function (en) {
    (en.changes || []).forEach(function (ch) {
      const v = ch.value || {}, contacts = v.contacts || [];
      (v.messages || []).forEach(function (m) {
        const c = contacts.filter(function (x) { return x.wa_id === m.from; })[0];
        const res = addLead_({
          name: c && c.profile ? c.profile.name : '',
          phone: m.from,
          message: m.text ? m.text.body : '[' + m.type + ']',
          time: Number(m.timestamp) * 1000
        });
        if (res.dup) dups++; else if (res.ok) added++;
      });
    });
  });
  return {ok: true, added: added, dups: dups};
}
function waSheet_(ss) {
  let sh = ss.getSheetByName(WA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(WA_SHEET);
    sh.appendRow(WA_HEAD);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, WA_HEAD.length).setFontWeight('bold').setBackground('#14284b').setFontColor('#ffffff');
    sh.getRange('A:A').setNumberFormat('dd-mmm-yyyy hh:mm');
    sh.getRange('H:H').setNumberFormat('dd-mmm-yyyy hh:mm');
    sh.getRange('C:C').setNumberFormat('0');
    sh.setColumnWidths(1, WA_HEAD.length, 130);
    sh.setColumnWidth(4, 300);
  }
  return sh;
}
function callSheet_(ss) {
  let sh = ss.getSheetByName(CALL_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CALL_SHEET);
    sh.appendRow(CALL_HEAD);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, CALL_HEAD.length).setFontWeight('bold').setBackground('#14284b').setFontColor('#ffffff');
    sh.getRange('A:A').setNumberFormat('dd-mmm-yyyy hh:mm:ss');
    sh.getRange('B:B').setNumberFormat('0');
    sh.setColumnWidths(1, CALL_HEAD.length, 130);
  }
  return sh;
}
function inboxFind_(sh, phone) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  const v = sh.getRange(2, 3, n, 1).getValues();
  for (let i = 0; i < n; i++) if (String(v[i][0]).replace(/\D/g, '').slice(-10) === phone) return i + 2;
  return 0;
}
function findLeadPR_(ss, users, phone, tz) {
  const ex = users.filter(function (x) { return x.role === 'exec' && x.active; });
  for (let i = 0; i < ex.length; i++) {
    const hit = readTab_(ss, ex[i], tz).filter(function (l) { return l.phone === phone; })[0];
    if (hit) return {x: ex[i], lead: hit};
  }
  return null;
}
function nextExec_() {
  const ex = getUsers_().filter(function (x) { return x.role === 'exec' && x.active; });
  if (!ex.length) return '';
  const p = PropertiesService.getScriptProperties();
  const i = Number(p.getProperty('RR_INDEX') || 0) % ex.length;
  p.setProperty('RR_INDEX', String(i + 1));
  return ex[i].name;
}
// o = {name, phone, message, time}
function addLead_(o) {
  const ss = SS_(), tz = ss.getSpreadsheetTimeZone();
  const phone = String(o.phone || '').replace(/\D/g, '').slice(-10);
  if (phone.length < 10) return {ok: false, error: 'Bad phone'};
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = waSheet_(ss), users = getUsers_();
    // duplicate check: WA Leads sheet + all PR tabs
    const wRow = inboxFind_(sh, phone);
    if (wRow) return {ok: true, dup: {where: WA_SHEET, row: wRow}};
    const d = findDup_(ss, users, phone, tz);
    if (d) return {ok: true, dup: {where: d.tab, exec: d.exec, name: d.name}};
    // new lead
    let when = o.time ? new Date(Number(o.time)) : new Date();
    if (isNaN(when.getTime())) when = new Date();
    const who = AUTO_ASSIGN ? nextExec_() : '';
    sh.appendRow([when, String(o.name || '').trim() || 'WhatsApp Lead', Number(phone),
      String(o.message || '').slice(0, 300), who, '', 0, '', '', 0, '']);
    return {ok: true, added: true, assigned: who};
  } finally { lock.releaseLock(); }
}

// Menu: copies every row that has "Assigned To" filled into that executive's PR tab
function moveAssignedToPR() {
  const ss = SS_(), tz = ss.getSpreadsheetTimeZone(), sh = waSheet_(ss), users = getUsers_();
  const n = sh.getLastRow() - 1;
  if (n < 1) { SpreadsheetApp.getUi().alert('No leads in "' + WA_SHEET + '".'); return; }
  const v = sh.getRange(2, 1, n, 6).getValues();
  let moved = 0, skipped = 0;
  v.forEach(function (r, i) {
    const who = String(r[4]).trim();
    if (!who || String(r[5]).trim()) return;
    const x = users.filter(function (u) { return u.role === 'exec' && u.active && u.name === who; })[0];
    if (!x) { sh.getRange(i + 2, 6).setValue('Unknown agent'); return; }
    const res = save_({role: 'admin'}, {tab: x.tab, fields: {
      date: Utilities.formatDate(r[0] instanceof Date ? r[0] : new Date(), tz, 'yyyy-MM-dd'),
      name: String(r[1]), phone: String(r[2]), status: 'New',
      remarks: 'WhatsApp: ' + String(r[3]).slice(0, 80)}});
    if (res.ok && res.lead) { sh.getRange(i + 2, 6).setValue('Moved to ' + x.tab + ' row ' + res.lead.row); moved++; }
    else if (res.ok && res.dup) { sh.getRange(i + 2, 6).setValue('Duplicate - skipped'); skipped++; }
    else sh.getRange(i + 2, 6).setValue('Error: ' + (res.error || ''));
  });
  SpreadsheetApp.getUi().alert('Moved: ' + moved + '\nDuplicates skipped: ' + skipped);
}

/* ---------- Call-log endpoint ----------
 * POST {action:'calllog', key:'API_KEY', phone:'919876543210', type:'outgoing|incoming|missed|rejected',
 *       duration:75 (seconds), time: 1732350000000 (ms, optional), agent:'PR01' (optional)}
 */
function callLog_(r) {
  const ss = SS_(), tz = ss.getSpreadsheetTimeZone();
  const phone = String(r.phone || '').replace(/\D/g, '').slice(-10);
  if (phone.length < 10) return {ok: false, error: 'Bad phone'};
  const type = String(r.type || 'outgoing').toLowerCase();
  const isOut = type === 'outgoing';
  const dur = Math.max(0, Math.round(Number(r.duration) || 0));
  let when = r.time ? new Date(isNaN(Number(r.time)) ? r.time : Number(r.time)) : new Date();
  if (isNaN(when.getTime())) when = new Date();
  let result;
  if (type === 'missed') result = 'Missed';
  else if (type === 'rejected') result = 'Rejected';
  else if (dur > 0) result = 'Answered';
  else result = type === 'incoming' ? 'Missed' : 'Not answered';

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const cs = callSheet_(ss), wa = waSheet_(ss), users = getUsers_();
    // ignore repeats of the same call (phone app may retry)
    const last = cs.getLastRow();
    if (last > 1) {
      const from = Math.max(2, last - 19);
      const rec = cs.getRange(from, 1, last - from + 1, 2).getValues();
      for (let i = 0; i < rec.length; i++)
        if (String(rec[i][1]).replace(/\D/g, '').slice(-10) === phone && Math.abs(new Date(rec[i][0]) - when) < 2000)
          return {ok: true, repeat: true};
    }
    const wRow = inboxFind_(wa, phone);
    const pr = findLeadPR_(ss, users, phone, tz);
    let agent = String(r.agent || '').trim();
    const byId = users.filter(function (u) { return u.id.toUpperCase() === agent.toUpperCase(); })[0];
    if (byId) agent = byId.name;
    let leadName = '', leadTab = '';
    if (pr) { leadName = pr.lead.name; leadTab = pr.x.tab; if (!agent) agent = pr.x.name; }
    if (wRow) {
      const wv = wa.getRange(wRow, 1, 1, WA_HEAD.length).getValues()[0];
      if (!leadName) leadName = String(wv[1]);
      if (!agent) agent = String(wv[4]);
      wa.getRange(wRow, 7, 1, 5).setValues([[(Number(wv[6]) || 0) + (isOut ? 1 : 0), when, dur, (Number(wv[9]) || 0) + dur, result]]);
    }
    if (pr && UPDATE_PR_TABS && isOut) {
      const sh = ss.getSheetByName(pr.x.tab), L = pr.lead;
      const day = Utilities.parseDate(Utilities.formatDate(when, tz, 'yyyy-MM-dd'), tz, 'yyyy-MM-dd');
      sh.getRange(L.row, COL.calls).setValue(L.calls + 1);
      if (!L.firstCalled) sh.getRange(L.row, COL.firstCalled).setValue(day);
      if (result === 'Answered') sh.getRange(L.row, COL.lastContact).setValue(day);
    }
    cs.appendRow([when, Number(phone), agent, type, dur, Math.floor(dur / 60) + ':' + ('0' + (dur % 60)).slice(-2), result, leadName, leadTab]);
    return {ok: true, result: result, duration: dur, matched: !!(wRow || pr)};
  } finally { lock.releaseLock(); }
}

/* ---------- Quick tests: run from the editor ---------- */
function testLead() { Logger.log(JSON.stringify(addLead_({name: 'Test Lead', phone: '9999900001', message: 'Hello, I want award details'}))); }
function testCall() { Logger.log(JSON.stringify(callLog_({phone: '9999900001', type: 'outgoing', duration: 45}))); }

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
