/**
 * Pour Tracker — shared state for ALL devices.
 *
 * Stores the dashboard's Done marks + manual Required/Rate in a "STATE" tab of THIS sheet,
 * so a pour closed on one device disappears on every device. The sheet is the single source of truth.
 *
 * ── SETUP (once) ────────────────────────────────────────────────────────────
 *  1. Open the PLAN sheet in Google Sheets.
 *  2. Extensions  ▸  Apps Script.
 *  3. Delete whatever is there, paste ALL of this file, then click Save (💾).
 *  4. Deploy  ▸  New deployment  ▸  (gear ⚙)  ▸  Web app
 *        Description   : Pour Tracker sync
 *        Execute as    : Me
 *        Who has access: Anyone            <-- must be "Anyone", not "Anyone with Google account"
 *     ▸ Deploy  ▸ Authorize access (approve the Google warning screens).
 *  5. Copy the "Web app" URL — it ends with /exec
 *  6. Paste it in the dashboard: ⚙ Settings ▸ "Sync URL" ▸ Load / Refresh.
 *
 *  The STATE tab is created automatically on first use. Don't rename it.
 * ────────────────────────────────────────────────────────────────────────────
 */

var SHEET_NAME = 'STATE';
var TYPES = ['done', 'req', 'rate'];

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, 3).setValues([['type', 'key', 'value']]);
  }
  return sh;
}

function readAll_() {
  var sh = sheet_();
  var out = { done: {}, req: {}, rate: {} };
  var last = sh.getLastRow();
  if (last < 2) return out;
  var rows = sh.getRange(2, 1, last - 1, 3).getValues();
  for (var i = 0; i < rows.length; i++) {
    var type = String(rows[i][0] || '').trim();
    var key = String(rows[i][1] || '');
    var val = String(rows[i][2] == null ? '' : rows[i][2]).trim();
    if (!type || !key || val === '') continue;
    if (TYPES.indexOf(type) < 0) continue;
    out[type][key] = (type === 'done') ? true : Number(val);
  }
  return out;
}

/** Insert / update / delete one (type,key) row. Empty value = delete. */
function writeOne_(type, key, value) {
  if (TYPES.indexOf(type) < 0 || !key) return;
  var sh = sheet_();
  var last = sh.getLastRow();
  var rowIdx = -1;
  if (last >= 2) {
    var keys = sh.getRange(2, 1, last - 1, 2).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (String(keys[i][0]).trim() === type && String(keys[i][1]) === String(key)) { rowIdx = i + 2; break; }
    }
  }
  var remove = (value === '' || value === null || value === undefined);
  if (rowIdx > 0) {
    if (remove) sh.deleteRow(rowIdx);
    else sh.getRange(rowIdx, 3).setValue(value);
  } else if (!remove) {
    sh.appendRow([type, key, value]);
  }
}

/** Delete every row of one type (used by "Reset done marks"). */
function clearType_(type) {
  if (TYPES.indexOf(type) < 0) return;
  var sh = sheet_();
  var last = sh.getLastRow();
  if (last < 2) return;
  var col = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = col.length - 1; i >= 0; i--) {            // bottom-up so row numbers stay valid
    if (String(col[i][0]).trim() === type) sh.deleteRow(i + 2);
  }
}

/* ── Presence: how many devices have the dashboard open right now ──────────────
   Deliberately kept OUT of the STATE sheet — it lives in the script cache, so a heartbeat
   every 30s never writes to the sheet and can never touch the Done / Required / Rate rows.
   Each device pings with its own id; a device counts as "open" until PRESENCE_TTL_MS passes
   with no ping (i.e. it drops off ~90s after the app is closed). */
var PRESENCE_TTL_MS = 90 * 1000;
var PRESENCE_KEY = 'presence';

/** Heartbeat. Stores {deviceId: [lastSeenMs, name]} in the cache and returns
 *  { online: <count>, who: [{name, seen}] } — the admin panel lists the names. */
function ping_(who, name, ip) {
  var cache = CacheService.getScriptCache();
  var lock = LockService.getScriptLock();
  var gotLock = false;
  try { lock.waitLock(5000); gotLock = true; } catch (e) {}
  var live = {};
  try {
    var map = {};
    var raw = cache.get(PRESENCE_KEY);
    if (raw) { try { map = JSON.parse(raw) || {}; } catch (e2) { map = {}; } }
    var now = new Date().getTime();
    if (who) map[who] = [now, String(name || ''), String(ip || '')];
    for (var k in map) {                                  // drop devices that stopped pinging
      var v = map[k];
      var t = (v && v.length) ? Number(v[0]) : Number(v);  // tolerate the old number-only format
      if (now - t < PRESENCE_TTL_MS) live[k] = [t, (v && v.length) ? String(v[1] || '') : '', (v && v.length > 2) ? String(v[2] || '') : ''];
    }
    cache.put(PRESENCE_KEY, JSON.stringify(live), 600);
  } finally {
    if (gotLock) { try { lock.releaseLock(); } catch (e3) {} }
  }
  var n = 0, list = [];
  for (var k2 in live) { n++; list.push({ name: live[k2][1] || '—', seen: live[k2][0], ip: live[k2][2] || '' }); }
  list.sort(function (a, b) { return b.seen - a.seen; });
  return { online: n, who: list };
}

/* ── Single admin lock: only one device may hold admin mode at a time ─────────
   Lives in the cache (never the sheet). The holder's heartbeat (ping with admin=1) keeps the lock
   alive; if they stop for ADMIN_TTL_MS the lock frees so someone else can take it. */
var ADMIN_KEY = 'adminHolder';
var ADMIN_TTL_MS = 120 * 1000;
function adminClaim_(dev, name) {
  var cache = CacheService.getScriptCache();
  var lock = LockService.getScriptLock(), got = false;
  try { lock.waitLock(5000); got = true; } catch (e) {}
  try {
    var now = new Date().getTime(), raw = cache.get(ADMIN_KEY), h = null;
    if (raw) { try { h = JSON.parse(raw); } catch (e2) {} }
    if (h && h.dev && h.dev !== dev && (now - Number(h.ts || 0) < ADMIN_TTL_MS)) {
      // NOTE: ok MUST be true — the dashboard's syncCall discards any reply with ok:false (treats it as a
      // network failure) and would then grant admin. The "held" flag carries the real answer.
      return { ok: true, held: true, by: (h.name || '') };     // someone else holds it
    }
    cache.put(ADMIN_KEY, JSON.stringify({ dev: dev, name: String(name || ''), ts: now }), 600);
    return { ok: true, held: false };
  } finally { if (got) { try { lock.releaseLock(); } catch (e3) {} } }
}
function adminRefresh_(dev, name) {                            // ping with admin=1 -> keep the holder alive
  var cache = CacheService.getScriptCache(), raw = cache.get(ADMIN_KEY), now = new Date().getTime();
  if (!raw) { cache.put(ADMIN_KEY, JSON.stringify({ dev: dev, name: String(name || ''), ts: now }), 600); return; }
  var h = null; try { h = JSON.parse(raw); } catch (e) {}
  if (h && h.dev === dev) { h.ts = now; if (name) h.name = String(name); cache.put(ADMIN_KEY, JSON.stringify(h), 600); }
}
function adminRelease_(dev) {
  var cache = CacheService.getScriptCache(), raw = cache.get(ADMIN_KEY);
  if (!raw) return; var h = null; try { h = JSON.parse(raw); } catch (e) {}
  if (h && h.dev === dev) cache.remove(ADMIN_KEY);
}

/* ── Users registry: remember each visitor's name by their IP, and let the admin block an IP ──
   Its own USERS tab (ip | name | lastSeen | blocked). The client sends the IP it got from a public
   echo service (Apps Script can't see the caller's IP), so this is a deterrent, not hard security.
   Never touches the STATE rows. */
var USERS_NAME = 'USERS';
function usersSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(USERS_NAME);
  if (!sh) { sh = ss.insertSheet(USERS_NAME); sh.getRange(1, 1, 1, 4).setValues([['ip', 'name', 'lastSeen', 'blocked']]); }
  return sh;
}
function userFindRow_(sh, ip) {
  var last = sh.getLastRow(); if (last < 2 || !ip) return -1;
  var ips = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ips.length; i++) if (String(ips[i][0]).trim() === String(ip).trim()) return i + 2;
  return -1;
}
function isBlocked_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function userUpsert_(ip, name) {
  if (!ip) return;
  var sh = usersSheet_(), row = userFindRow_(sh, ip), now = new Date();
  if (row > 0) { if (name) sh.getRange(row, 2).setValue(name); sh.getRange(row, 3).setValue(now); }
  else sh.appendRow([String(ip), String(name || ''), now, false]);
}
function userTouch_(ip) {                                   // bump lastSeen only (called by whoami)
  if (!ip) return; var sh = usersSheet_(), row = userFindRow_(sh, ip);
  if (row > 0) sh.getRange(row, 3).setValue(new Date());
}
function userLookup_(ip) {
  if (!ip) return null; var sh = usersSheet_(), row = userFindRow_(sh, ip);
  if (row < 2) return null; var v = sh.getRange(row, 1, 1, 4).getValues()[0];
  return { ip: String(v[0]), name: String(v[1] || ''), blocked: isBlocked_(v[3]) };
}
function usersAll_() {
  var sh = usersSheet_(), last = sh.getLastRow(); if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, 4).getValues(), out = [];
  for (var i = 0; i < rows.length; i++) {
    if (!String(rows[i][0]).trim()) continue; var ls = rows[i][2];
    out.push({ ip: String(rows[i][0]), name: String(rows[i][1] || ''),
      lastSeen: (ls instanceof Date) ? ls.getTime() : String(ls || ''), blocked: isBlocked_(rows[i][3]) });
  }
  out.sort(function (a, b) { return (b.lastSeen || 0) - (a.lastSeen || 0); });
  return out;
}
function userSetBlocked_(ip, blocked) {
  if (!ip) return; var sh = usersSheet_(), row = userFindRow_(sh, ip);
  if (row < 2) { sh.appendRow([String(ip), '', new Date(), !!blocked]); return; }
  sh.getRange(row, 4).setValue(!!blocked);
}

/* ── Edit log: who changed what, appended to a separate LOG tab ────────────────
   A real audit trail has to survive, so unlike presence this IS written to the sheet —
   but to its OWN tab. Nothing here ever touches the STATE rows. */
var LOG_NAME = 'LOG';
var LOG_MAX = 5000;                                       // trim oldest beyond this

function logSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(LOG_NAME);
  if (!sh) {
    sh = ss.insertSheet(LOG_NAME);
    sh.getRange(1, 1, 1, 5).setValues([['when', 'who', 'type', 'key', 'value']]);
  }
  return sh;
}

function logWrite_(by, type, key, value) {
  try {
    var sh = logSheet_();
    sh.appendRow([new Date(), String(by || '—'), String(type || ''), String(key || ''),
                  (value === '' || value === null || value === undefined) ? '(cleared)' : String(value)]);
    var last = sh.getLastRow();
    if (last > LOG_MAX + 1) sh.deleteRows(2, last - LOG_MAX - 1);   // drop oldest, keep the header
  } catch (e) {}                                          // logging must never break an edit
}

/** Most recent entries, newest first — for the admin panel. */
function logRead_(limit) {
  var sh = logSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var n = Math.min(limit || 100, last - 1);
  var rows = sh.getRange(last - n + 1, 1, n, 5).getValues();
  var out = [];
  for (var i = rows.length - 1; i >= 0; i--) {
    var w = rows[i][0];
    out.push({
      when: (w instanceof Date) ? w.getTime() : String(w),
      who: String(rows[i][1] || ''), type: String(rows[i][2] || ''),
      key: String(rows[i][3] || ''), value: String(rows[i][4] == null ? '' : rows[i][4])
    });
  }
  return out;
}

/** JSONP when a callback is given (the dashboard uses JSONP to avoid CORS), else plain JSON. */
function reply_(p, payload) {
  var json = JSON.stringify(payload);
  if (p.callback) {
    return ContentService.createTextOutput(p.callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var payload;
  try {
    var action = String(p.action || 'get');
    if (action === 'ping') {                              // heartbeat only — never reads the STATE sheet
      var pr = ping_(String(p.who || ''), String(p.name || ''), String(p.ip || ''));
      if (String(p.admin || '') === '1') adminRefresh_(String(p.who || ''), String(p.name || ''));
      payload = { ok: true, online: pr.online, who: pr.who };
      return reply_(p, payload);
    }
    if (action === 'adminclaim') {                        // request the single admin seat
      payload = adminClaim_(String(p.who || ''), String(p.name || ''));
      return reply_(p, payload);
    }
    if (action === 'adminrelease') {                      // give up the admin seat
      adminRelease_(String(p.who || ''));
      payload = { ok: true };
      return reply_(p, payload);
    }
    if (action === 'whoami') {                            // client asks by IP: my remembered name + am I blocked
      var u = userLookup_(String(p.ip || ''));
      if (u) { try { userTouch_(String(p.ip || '')); } catch (e) {} }
      payload = { ok: true, name: u ? u.name : '', blocked: u ? u.blocked : false };
      return reply_(p, payload);
    }
    if (action === 'register') {                          // user typed their name -> remember it against their IP
      var lkR = LockService.getScriptLock(); try { lkR.waitLock(10000); } catch (e) {}
      try { userUpsert_(String(p.ip || ''), String(p.name || '')); } finally { try { lkR.releaseLock(); } catch (e) {} }
      var ur = userLookup_(String(p.ip || ''));
      payload = { ok: true, blocked: ur ? ur.blocked : false };
      return reply_(p, payload);
    }
    if (action === 'users') {                             // admin panel: every known visitor
      payload = { ok: true, users: usersAll_() };
      return reply_(p, payload);
    }
    if (action === 'block' || action === 'unblock') {     // admin: block / unblock an IP
      var lkB = LockService.getScriptLock(); try { lkB.waitLock(10000); } catch (e) {}
      try { userSetBlocked_(String(p.ip || ''), action === 'block'); } finally { try { lkB.releaseLock(); } catch (e) {} }
      logWrite_(p.by, action === 'block' ? 'BLOCK ip' : 'UNBLOCK ip', String(p.ip || ''), '');
      payload = { ok: true };
      return reply_(p, payload);
    }
    if (action === 'log') {                               // admin panel: recent edits
      payload = { ok: true, log: logRead_(Number(p.limit || 100)) };
      return reply_(p, payload);
    }
    if (action === 'set' || action === 'clear') {
      var lock = LockService.getScriptLock();
      lock.waitLock(20000);                               // serialise concurrent devices
      try {
        if (action === 'set') writeOne_(String(p.type || ''), String(p.key || ''), p.value === undefined ? '' : p.value);
        else clearType_(String(p.type || ''));
      } finally {
        lock.releaseLock();
      }
      logWrite_(p.by, action === 'clear' ? ('clear:' + String(p.type || '')) : String(p.type || ''),
                action === 'clear' ? '(all)' : String(p.key || ''),
                action === 'clear' ? '' : p.value);
    }
    payload = { ok: true, state: readAll_() };
  } catch (err) {
    payload = { ok: false, error: String(err) };
  }
  var json = JSON.stringify(payload);
  if (p.callback) {                                       // JSONP -> no CORS problems in the dashboard
    return ContentService.createTextOutput(p.callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
