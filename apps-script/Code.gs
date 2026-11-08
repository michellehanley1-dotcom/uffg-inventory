/**
 * UFFG Inventory — back end.
 * One Google Sheet, one web app. Deploy: Execute as Me, Who has access: Anyone.
 * Columns are always found by header NAME, never by position.
 */

var VERSION   = '1.0.0';
var SHEET_ID  = ''; // leave blank to use the sheet this script is bound to
var CODE_TTL  = 10 * 60 * 1000;        // 6-digit code lifetime
var CODE_TRIES= 5;
var RESEND_MS = 45 * 1000;
var SESS_DAYS = 120;
var BRAND     = 'UFFG Construction';

var TABS = {
  Team:      ['EMAIL','NAME','TRUCK/LOCATION','ROLE'],
  Locations: ['NAME','TYPE','ACTIVE'],
  Tools:     ['TOOL ID','NAME','BRAND','MODEL','STATUS','LOCATION','CHECKED OUT BY','CHECKED OUT AT','NOTES','UPDATED AT'],
  Parts:     ['PART','BRAND','UNIT','LOW TARGET'],
  Stock:     ['PART','LOCATION','QTY'],
  Log:       ['WHEN','WHO','ACTION','KIND','ITEM','FROM','TO','QTY','NOTE'],
  Auth:      ['EMAIL','CODE','EXPIRES','TRIES','LAST SENT'],
  Sessions:  ['TOKEN','EMAIL','EXPIRES'],
  Processed: ['ID','WHEN']
};

var STATUSES = ['Good','Needs repair','Lost'];

/* ------------------------------------------------------------------ sheet */

function ss_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive();
}

function tab_(name) {
  var s = ss_().getSheetByName(name);
  if (!s) {
    s = ss_().insertSheet(name);
    s.getRange(1, 1, 1, TABS[name].length).setValues([TABS[name]]);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, TABS[name].length).setFontWeight('bold');
  }
  return s;
}

/** Header name -> 0-based column index. */
function headers_(sheet) {
  var last = Math.max(1, sheet.getLastColumn());
  var row  = sheet.getRange(1, 1, 1, last).getValues()[0];
  var map  = {};
  for (var i = 0; i < row.length; i++) {
    var k = String(row[i]).trim().toUpperCase();
    if (k) map[k] = i;
  }
  return map;
}

/** All data rows as objects keyed by header name. */
function rows_(name) {
  var sh = tab_(name);
  if (sh.getLastRow() < 2) return [];
  var h = headers_(sh);
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  return vals.map(function (r, i) {
    var o = { _row: i + 2 };
    for (var k in h) o[k] = r[h[k]];
    return o;
  });
}

/** Append one object, placing each value under its header. */
function append_(name, obj) {
  var sh = tab_(name), h = headers_(sh);
  var width = Math.max(sh.getLastColumn(), 1);
  var arr = new Array(width).fill('');
  for (var k in obj) {
    var key = k.toUpperCase();
    if (h[key] === undefined) continue;
    arr[h[key]] = obj[k];
  }
  sh.appendRow(arr);
}

/** Update one existing row by its _row number. */
function update_(name, rowNum, obj) {
  var sh = tab_(name), h = headers_(sh);
  for (var k in obj) {
    var key = k.toUpperCase();
    if (h[key] === undefined) continue;
    sh.getRange(rowNum, h[key] + 1).setValue(obj[k]);
  }
}

function norm_(v) { return String(v == null ? '' : v).trim(); }
function lower_(v) { return norm_(v).toLowerCase(); }
function now_()  { return new Date(); }

/* ------------------------------------------------------------------- auth */

function teamMember_(email) {
  var e = lower_(email);
  var t = rows_('Team');
  for (var i = 0; i < t.length; i++) {
    if (lower_(t[i]['EMAIL']) === e) {
      return {
        email:    lower_(t[i]['EMAIL']),
        name:     norm_(t[i]['NAME']) || lower_(t[i]['EMAIL']),
        base:     norm_(t[i]['TRUCK/LOCATION']),
        role:     norm_(t[i]['ROLE']) || 'Lead',
        isManager: lower_(t[i]['ROLE']) === 'manager'
      };
    }
  }
  return null;
}

function requestCode_(args) {
  var email = lower_(args.email);
  if (!email) throw new Error('Enter your email.');
  // Always answer the same way, so this cannot be used to discover who is on the team.
  var member = teamMember_(email);
  if (!member) return { sent: true };

  var sh = tab_('Auth'), all = rows_('Auth'), existing = null;
  for (var i = 0; i < all.length; i++) if (lower_(all[i]['EMAIL']) === email) existing = all[i];

  if (existing && existing['LAST SENT']) {
    var since = now_().getTime() - new Date(existing['LAST SENT']).getTime();
    if (since < RESEND_MS) {
      return { sent: true, waitMs: RESEND_MS - since };
    }
  }

  var code = String(Math.floor(100000 + Math.random() * 900000));
  var rec = {
    'EMAIL': email,
    'CODE': hash_(email + ':' + code),
    'EXPIRES': new Date(now_().getTime() + CODE_TTL),
    'TRIES': 0,
    'LAST SENT': now_()
  };
  if (existing) update_('Auth', existing._row, rec); else append_('Auth', rec);

  MailApp.sendEmail({
    to: email,
    subject: code + ' is your ' + BRAND + ' inventory code',
    body: 'Your sign-in code is ' + code + '\n\n' +
          'It expires in 10 minutes. If you did not ask for this, ignore this email.\n',
    name: BRAND + ' Inventory'
  });
  return { sent: true };
}

function hash_(s) {
  var salt = PropertiesService.getScriptProperties().getProperty('SALT');
  if (!salt) {
    salt = Utilities.getUuid();
    PropertiesService.getScriptProperties().setProperty('SALT', salt);
  }
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + s, Utilities.Charset.UTF_8);
  return raw.map(function (b) { return ((b & 0xff) + 0x100).toString(16).slice(1); }).join('');
}

function verifyCode_(args) {
  var email = lower_(args.email), code = norm_(args.code);
  var all = rows_('Auth'), rec = null;
  for (var i = 0; i < all.length; i++) if (lower_(all[i]['EMAIL']) === email) rec = all[i];
  if (!rec) throw new Error('Ask for a new code.');
  if (new Date(rec['EXPIRES']).getTime() < now_().getTime()) throw new Error('That code expired. Ask for a new one.');
  if (Number(rec['TRIES']) >= CODE_TRIES) throw new Error('Too many tries. Ask for a new code.');

  if (hash_(email + ':' + code) !== norm_(rec['CODE'])) {
    update_('Auth', rec._row, { 'TRIES': Number(rec['TRIES']) + 1 });
    throw new Error('That code is not right.');
  }
  var member = teamMember_(email);
  if (!member) throw new Error('You are not on the team list.');

  update_('Auth', rec._row, { 'CODE': '', 'TRIES': 0, 'EXPIRES': '' });
  var token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  append_('Sessions', {
    'TOKEN': hash_(token),
    'EMAIL': email,
    'EXPIRES': new Date(now_().getTime() + SESS_DAYS * 864e5)
  });
  return { token: token, me: member };
}

/** Resolve a token to a team member, or null. Deleting a Team row signs them out. */
function session_(token) {
  if (!token) return null;
  var h = hash_(token), all = rows_('Sessions');
  for (var i = 0; i < all.length; i++) {
    if (norm_(all[i]['TOKEN']) === h) {
      if (new Date(all[i]['EXPIRES']).getTime() < now_().getTime()) return null;
      return teamMember_(all[i]['EMAIL']);   // gone from Team => signed out
    }
  }
  return null;
}

function signOut_(args, me, token) {
  var h = hash_(token), sh = tab_('Sessions'), all = rows_('Sessions');
  for (var i = all.length - 1; i >= 0; i--) {
    if (norm_(all[i]['TOKEN']) === h) sh.deleteRow(all[i]._row);
  }
  return { ok: true };
}

/* --------------------------------------------------------- idempotency */

function alreadyDone_(id) {
  if (!id) return false;
  var all = rows_('Processed');
  for (var i = all.length - 1; i >= 0; i--) if (norm_(all[i]['ID']) === id) return true;
  return false;
}

function markDone_(id) {
  if (!id) return;
  append_('Processed', { 'ID': id, 'WHEN': now_() });
  var sh = tab_('Processed');
  if (sh.getLastRow() > 3000) sh.deleteRows(2, 1000);   // keep it from growing forever
}

/* ------------------------------------------------------------------- log */

function log_(who, action, kind, item, from, to, qty, note) {
  append_('Log', {
    'WHEN': now_(), 'WHO': who, 'ACTION': action, 'KIND': kind,
    'ITEM': item, 'FROM': from || '', 'TO': to || '',
    'QTY': (qty === undefined || qty === null) ? '' : qty, 'NOTE': note || ''
  });
}

/* ----------------------------------------------------------------- reads */

function bootstrap_(args, me) {
  var acting = me;
  if (args && args.viewAs && me.isManager) {
    var other = teamMember_(args.viewAs);
    if (other) acting = other;
  }
  var tools = rows_('Tools').map(function (t) {
    return {
      id: norm_(t['TOOL ID']), name: norm_(t['NAME']), brand: norm_(t['BRAND']),
      model: norm_(t['MODEL']), status: norm_(t['STATUS']) || 'Good',
      location: norm_(t['LOCATION']), by: norm_(t['CHECKED OUT BY']),
      at: t['CHECKED OUT AT'] ? new Date(t['CHECKED OUT AT']).toISOString() : '',
      notes: norm_(t['NOTES'])
    };
  });
  var parts = rows_('Parts').map(function (p) {
    return {
      part: norm_(p['PART']), brand: norm_(p['BRAND']),
      unit: norm_(p['UNIT']) || 'ea', low: Number(p['LOW TARGET']) || 0
    };
  });
  var stock = rows_('Stock').map(function (s) {
    return { part: norm_(s['PART']), location: norm_(s['LOCATION']), qty: Number(s['QTY']) || 0 };
  }).filter(function (s) { return s.part; });

  var locs = rows_('Locations').filter(function (l) {
    return String(l['ACTIVE']).toUpperCase() !== 'NO';
  }).map(function (l) {
    return { name: norm_(l['NAME']), type: norm_(l['TYPE']) || 'Job site' };
  });

  var team = rows_('Team').map(function (t) {
    return { email: lower_(t['EMAIL']), name: norm_(t['NAME']) || lower_(t['EMAIL']),
             base: norm_(t['TRUCK/LOCATION']), role: norm_(t['ROLE']) || 'Lead' };
  });

  var log = rows_('Log').slice(-120).reverse().map(function (l) {
    return { when: l['WHEN'] ? new Date(l['WHEN']).toISOString() : '', who: norm_(l['WHO']),
             action: norm_(l['ACTION']), kind: norm_(l['KIND']), item: norm_(l['ITEM']),
             from: norm_(l['FROM']), to: norm_(l['TO']), qty: l['QTY'], note: norm_(l['NOTE']) };
  });

  return {
    version: VERSION, me: me, acting: acting, team: team,
    locations: locs, tools: tools, parts: parts, stock: stock, log: log,
    statuses: STATUSES, serverTime: now_().toISOString()
  };
}

/* ---------------------------------------------------------------- writes */

function addTool_(a, me) {
  var id = norm_(a.id) || nextToolId_();
  if (!norm_(a.name)) throw new Error('A tool needs a name.');
  var dupes = rows_('Tools').filter(function (t) { return lower_(t['TOOL ID']) === lower_(id); });
  if (dupes.length) throw new Error('Tool ID ' + id + ' already exists.');
  append_('Tools', {
    'TOOL ID': id, 'NAME': norm_(a.name), 'BRAND': norm_(a.brand), 'MODEL': norm_(a.model),
    'STATUS': STATUSES.indexOf(norm_(a.status)) >= 0 ? norm_(a.status) : 'Good',
    'LOCATION': norm_(a.location), 'CHECKED OUT BY': '', 'CHECKED OUT AT': '',
    'NOTES': norm_(a.notes), 'UPDATED AT': now_()
  });
  log_(me.name, 'Added tool', 'Tool', norm_(a.name) + ' (' + id + ')', '', norm_(a.location), '', norm_(a.notes));
  return { ok: true, id: id };
}

function nextToolId_() {
  var n = rows_('Tools').length + 1;
  var ids = {};
  rows_('Tools').forEach(function (t) { ids[lower_(t['TOOL ID'])] = 1; });
  while (ids['t-' + n]) n++;
  return 'T-' + n;
}

function findTool_(id) {
  var all = rows_('Tools');
  for (var i = 0; i < all.length; i++) if (lower_(all[i]['TOOL ID']) === lower_(id)) return all[i];
  throw new Error('No tool with ID ' + id + '.');
}

function moveTool_(a, me, acting) {
  var t = findTool_(a.toolId);
  var from = norm_(t['LOCATION']);
  var to   = norm_(a.to);
  if (!to) throw new Error('Pick where it is going.');
  // A regular user may only move a tool to or from their own base.
  if (!me.isManager) {
    var mine = lower_(acting.base);
    if (lower_(from) !== mine && lower_(to) !== mine) {
      throw new Error('You can only move tools to or from ' + (acting.base || 'your own location') + '.');
    }
  }
  var holder = (norm_(a.to).toLowerCase() === lower_(acting.base)) ? acting.name : '';
  update_('Tools', t._row, {
    'LOCATION': to,
    'CHECKED OUT BY': holder,
    'CHECKED OUT AT': holder ? now_() : '',
    'UPDATED AT': now_()
  });
  log_(acting.name, holder ? 'Checked out' : 'Moved', 'Tool',
       norm_(t['NAME']) + ' (' + norm_(t['TOOL ID']) + ')', from, to, '', norm_(a.note));
  return { ok: true };
}

function setToolStatus_(a, me, acting) {
  var t = findTool_(a.toolId);
  var st = norm_(a.status);
  if (STATUSES.indexOf(st) < 0) throw new Error('Unknown status.');
  update_('Tools', t._row, { 'STATUS': st, 'NOTES': norm_(a.note) || norm_(t['NOTES']), 'UPDATED AT': now_() });
  log_(acting.name, 'Status ' + st, 'Tool', norm_(t['NAME']) + ' (' + norm_(t['TOOL ID']) + ')',
       norm_(t['STATUS']), st, '', norm_(a.note));
  return { ok: true };
}

function addPart_(a, me) {
  if (!me.isManager) throw new Error('Only a manager can add a part.');
  var name = norm_(a.part);
  if (!name) throw new Error('A part needs a name.');
  var dupes = rows_('Parts').filter(function (p) { return lower_(p['PART']) === lower_(name); });
  if (dupes.length) throw new Error(name + ' is already on the list.');
  append_('Parts', { 'PART': name, 'BRAND': norm_(a.brand), 'UNIT': norm_(a.unit) || 'ea',
                     'LOW TARGET': Number(a.low) || 0 });
  log_(me.name, 'Added part', 'Part', name, '', '', '', '');
  return { ok: true };
}

function setLowTarget_(a, me) {
  if (!me.isManager) throw new Error('Only a manager can change the low target.');
  var all = rows_('Parts');
  for (var i = 0; i < all.length; i++) {
    if (lower_(all[i]['PART']) === lower_(a.part)) {
      update_('Parts', all[i]._row, { 'LOW TARGET': Number(a.low) || 0 });
      log_(me.name, 'Low target', 'Part', norm_(a.part), '', String(Number(a.low) || 0), '', '');
      return { ok: true };
    }
  }
  throw new Error('No part called ' + a.part + '.');
}

function stockRow_(part, location) {
  var all = rows_('Stock');
  for (var i = 0; i < all.length; i++) {
    if (lower_(all[i]['PART']) === lower_(part) && lower_(all[i]['LOCATION']) === lower_(location)) return all[i];
  }
  return null;
}

function setQty_(part, location, qty) {
  var r = stockRow_(part, location);
  if (r) update_('Stock', r._row, { 'QTY': qty });
  else append_('Stock', { 'PART': part, 'LOCATION': location, 'QTY': qty });
}

function qtyAt_(part, location) {
  var r = stockRow_(part, location);
  return r ? (Number(r['QTY']) || 0) : 0;
}

/** reason: 'Picked up' | 'Dropped off' | 'New stock' | 'Used' | 'Count' */
function adjustStock_(a, me, acting) {
  var part = norm_(a.part), loc = norm_(a.location), reason = norm_(a.reason);
  if (!part || !loc) throw new Error('Pick a part and a location.');
  if (!me.isManager && lower_(loc) !== lower_(acting.base) && reason !== 'Picked up' && reason !== 'Dropped off') {
    throw new Error('You can only change counts at ' + (acting.base || 'your own location') + '.');
  }
  var before = qtyAt_(part, loc), after;
  if (reason === 'Count') {
    after = Math.max(0, Number(a.qty) || 0);
  } else {
    var delta = Number(a.qty) || 0;
    if (delta <= 0) throw new Error('Enter how many.');
    after = (reason === 'Used' || reason === 'Picked up') ? before - delta : before + delta;
    if (after < 0) throw new Error('Only ' + before + ' at ' + loc + '.');
  }
  setQty_(part, loc, after);
  log_(acting.name, reason, 'Part', part, String(before), String(after),
       String(after - before), norm_(a.note));
  return { ok: true, qty: after };
}

function transferStock_(a, me, acting) {
  var part = norm_(a.part), from = norm_(a.from), to = norm_(a.to), qty = Number(a.qty) || 0;
  if (!part || !from || !to) throw new Error('Pick a part and both locations.');
  if (from === to) throw new Error('Those are the same place.');
  if (qty <= 0) throw new Error('Enter how many.');
  if (!me.isManager && lower_(from) !== lower_(acting.base) && lower_(to) !== lower_(acting.base)) {
    throw new Error('One side has to be ' + (acting.base || 'your own location') + '.');
  }
  var have = qtyAt_(part, from);
  if (have < qty) throw new Error('Only ' + have + ' at ' + from + '.');
  setQty_(part, from, have - qty);
  setQty_(part, to, qtyAt_(part, to) + qty);
  log_(acting.name, 'Moved', 'Part', part, from, to, String(qty), norm_(a.note));
  return { ok: true };
}

function addLocation_(a, me) {
  var name = norm_(a.name);
  if (!name) throw new Error('Give the location a name.');
  var all = rows_('Locations');
  for (var i = 0; i < all.length; i++) {
    if (lower_(all[i]['NAME']) === lower_(name)) {
      update_('Locations', all[i]._row, { 'ACTIVE': 'YES' });
      return { ok: true };
    }
  }
  append_('Locations', { 'NAME': name, 'TYPE': norm_(a.type) || 'Job site', 'ACTIVE': 'YES' });
  log_(me.name, 'Added location', 'Location', name, '', '', '', '');
  return { ok: true };
}

function archiveLocation_(a, me) {
  if (!me.isManager) throw new Error('Only a manager can archive a location.');
  var all = rows_('Locations');
  for (var i = 0; i < all.length; i++) {
    if (lower_(all[i]['NAME']) === lower_(a.name)) {
      update_('Locations', all[i]._row, { 'ACTIVE': 'NO' });
      log_(me.name, 'Archived location', 'Location', norm_(a.name), '', '', '', '');
      return { ok: true };
    }
  }
  throw new Error('No location called ' + a.name + '.');
}

/** Paste-a-list import. kind: 'tools' | 'parts' */
function bulkImport_(a, me) {
  if (!me.isManager) throw new Error('Only a manager can import.');
  var kind = norm_(a.kind), lines = a.rows || [];
  var added = 0, skipped = [];
  for (var i = 0; i < lines.length; i++) {
    var c = lines[i];
    try {
      if (kind === 'tools') {
        addTool_({ name: c[0], brand: c[1], model: c[2], id: c[3], location: c[4] || '', status: c[5] || 'Good' }, me);
      } else {
        addPart_({ part: c[0], brand: c[1], unit: c[2], low: c[3] }, me);
      }
      added++;
    } catch (e) { skipped.push((c[0] || '(blank)') + ': ' + e.message); }
  }
  return { ok: true, added: added, skipped: skipped };
}

/* --------------------------------------------------------------- routing */

var PUBLIC = { requestCode: requestCode_, verifyCode: verifyCode_ };
var NEEDS_ID = { addTool:1, moveTool:1, setToolStatus:1, addPart:1, setLowTarget:1,
                 adjustStock:1, transferStock:1, addLocation:1, archiveLocation:1, bulkImport:1 };

function handlers_() {
  return {
    bootstrap: function (a, me) { return bootstrap_(a, me); },
    signOut: signOut_,
    addTool: addTool_, moveTool: moveTool_, setToolStatus: setToolStatus_,
    addPart: addPart_, setLowTarget: setLowTarget_,
    adjustStock: adjustStock_, transferStock: transferStock_,
    addLocation: addLocation_, archiveLocation: archiveLocation_,
    bulkImport: bulkImport_
  };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (e && e.parameter && e.parameter.ping) {
    var health = {};
    try {
      for (var t in TABS) health[t] = Math.max(0, tab_(t).getLastRow() - 1);
    } catch (err) { return json_({ ok: false, version: VERSION, error: String(err) }); }
    return json_({ ok: true, version: VERSION, sheet: ss_().getName(), rows: health, time: now_().toISOString() });
  }
  return json_({ ok: true, version: VERSION, hint: 'Add ?ping=1 for health.' });
}

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'Bad request.' }); }

  var fn = req.fn, args = req.args || {}, id = req.id || '', token = req.token || '';

  try {
    if (PUBLIC[fn]) return json_({ ok: true, data: PUBLIC[fn](args) });

    var me = session_(token);
    if (!me) return json_({ ok: false, code: 'AUTH', error: 'Please sign in again.' });

    var h = handlers_()[fn];
    if (!h) return json_({ ok: false, error: 'Unknown action.' });

    // Managers may act as another team member.
    var acting = me;
    if (args.viewAs && me.isManager) {
      var other = teamMember_(args.viewAs);
      if (other) acting = other;
    }

    if (NEEDS_ID[fn]) {
      var lock = LockService.getScriptLock();
      lock.waitLock(25000);
      try {
        if (alreadyDone_(id)) return json_({ ok: true, data: { ok: true, duplicate: true } });
        var out = h(args, me, acting, token);
        markDone_(id);
        return json_({ ok: true, data: out });
      } finally { lock.releaseLock(); }
    }
    return json_({ ok: true, data: h(args, me, acting, token) });

  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

/* ------------------------------------------------- one-time sheet set-up */

function setup() {
  for (var t in TABS) tab_(t);
  if (rows_('Team').length === 0) {
    append_('Team', { 'EMAIL': 'ricardo@uffgconstruction.com', 'NAME': 'Ricardo Uribe',
                      'TRUCK/LOCATION': "Ricky's Truck", 'ROLE': 'Manager' });
    append_('Team', { 'EMAIL': 'michellehanley1@gmail.com', 'NAME': 'Michelle Hanley',
                      'TRUCK/LOCATION': 'Garage', 'ROLE': 'Manager' });
  }
  if (rows_('Locations').length === 0) {
    append_('Locations', { 'NAME': 'Garage', 'TYPE': 'Shop', 'ACTIVE': 'YES' });
    append_('Locations', { 'NAME': "Ricky's Truck", 'TYPE': 'Truck', 'ACTIVE': 'YES' });
  }
  SpreadsheetApp.getActive().toast('Tabs ready.');
}
