/**
 * UFFG Inventory — back end.
 * One Google Sheet, one web app. Deploy: Execute as Me, Who has access: Anyone.
 * Columns are always found by header NAME, never by position.
 */

var VERSION   = '2.5.0';
var SHEET_ID  = ''; // leave blank to use the sheet this script is bound to
var CODE_TTL  = 10 * 60 * 1000;        // 6-digit code lifetime
var CODE_TRIES= 5;
var RESEND_MS = 45 * 1000;
var SESS_DAYS = 120;
var BRAND     = 'UFFG Construction';

var TABS = {
  Team:      ['EMAIL','NAME','TRUCK/LOCATION','ROLE','BUSINESSES'],
  Locations: ['NAME','TYPE','ACTIVE'],
  Tools:     ['TOOL ID','NAME','BRAND','MODEL','STATUS','LOCATION','CHECKED OUT BY','CHECKED OUT AT','NOTES','UPDATED AT'],
  Parts:     ['PART','BRAND','UNIT','LOW TARGET'],
  Stock:     ['PART','LOCATION','QTY'],
  Log:       ['WHEN','WHO','ACTION','KIND','ITEM','FROM','TO','QTY','NOTE'],
  Auth:      ['EMAIL','CODE','EXPIRES','TRIES','LAST SENT'],
  Sessions:  ['TOKEN','EMAIL','EXPIRES'],
  Processed: ['ID','WHEN'],
  Customers: ['ID','BUSINESS','NAME','COMPANY','PHONE','EMAIL','ADDRESS','NOTES','ACTIVE','CREATED'],
  Services:  ['ID','BUSINESS','NAME','DESCRIPTION','UNIT','PRICE','TAXABLE','ACTIVE'],
  Docs:      ['ID','BUSINESS','TYPE','NUMBER','CUSTOMER','TITLE','SITE','STATUS','DATE','DUE','PARENT','LINES',
              'TAX RATE','DISCOUNT','DEPOSIT','NOTES','TERMS','SUBTOTAL','TAX','TOTAL','PAID','BILLED',
              'SENT AT','APPROVED AT','CREATED BY','CREATED','UPDATED','DELETED'],
  Payments:  ['ID','INVOICE','DATE','AMOUNT','METHOD','NOTE','WHO','DELETED','BANK'],
  Expenses:  ['ID','BUSINESS','STATUS','DATE','VENDOR','ACCOUNT','AMOUNT','METHOD','JOB','NOTE','RECEIPT',
              'RECEIPT LINK','HASH','DUP OF','SCANNED','CREATED BY','CREATED','UPDATED','DELETED',
              'BANK','PAID BY','CHECK NO','CHECK FILE','CHECK LINK'],
  Company:   ['KEY','VALUE'],
  Banks:     ['ID','BUSINESS','NAME','TYPE','INSTITUTION','LAST4','OPENING BALANCE','OPENING DATE','ACTIVE','CREATED'],
  Vendors:   ['ID','BUSINESS','NAME','TYPE','PHONE','EMAIL','NOTES','ACTIVE','CREATED'],
  Banking:   ['ID','BUSINESS','TYPE','DATE','AMOUNT','BANK','TO BANK','PERSON','NOTE','WHO','CREATED','DELETED']
};

/** Tabs stored as plain text so Sheets never turns "2026-11-13" or "1/2" into a date. */
var TEXT_TABS = { Customers:1, Services:1, Docs:1, Payments:1, Expenses:1, Company:1, Banks:1, Vendors:1, Banking:1 };
/** Columns kept as real numbers so the sheet can add them up. */
var MONEY_COLS = { Expenses: { 'AMOUNT': 1 }, Payments: { 'AMOUNT': 1 }, Banking: { 'AMOUNT': 1 }, Banks: { 'OPENING BALANCE': 1 } };

function moneyCells_(name, sh, h, rowNum, arr) {
  var cols = MONEY_COLS[name]; if (!cols) return;
  for (var c in cols) {
    if (h[c] === undefined) continue;
    var v = Number(arr[h[c]]);
    sh.getRange(rowNum, h[c] + 1).setNumberFormat('$#,##0.00').setValue(isFinite(v) ? v : 0);
  }
}

var STATUSES = ['Good','Needs repair','Lost'];

/* ------------------------------------------------------------------ sheet */

function ss_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive();
}

function tab_(name) {
  var s = ss_().getSheetByName(name);
  if (!s) {
    s = ss_().insertSheet(name);
    if (TEXT_TABS[name]) s.getRange(1, 1, s.getMaxRows(), TABS[name].length).setNumberFormat('@');
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
    arr[h[key]] = TEXT_TABS[name] ? txt_(obj[k]) : obj[k];
  }
  if (TEXT_TABS[name]) {
    var r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, width).setNumberFormat('@').setValues([arr]);
    moneyCells_(name, sh, h, r, arr);
  } else {
    sh.appendRow(arr);
  }
}

function txt_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v.toISOString();
  return String(v);
}

/** Update one existing row by its _row number. */
function update_(name, rowNum, obj) {
  var sh = tab_(name), h = headers_(sh);
  for (var k in obj) {
    var key = k.toUpperCase();
    if (h[key] === undefined) continue;
    var cell = sh.getRange(rowNum, h[key] + 1);
    if (TEXT_TABS[name]) cell.setNumberFormat('@').setValue(txt_(obj[k]));
    else cell.setValue(obj[k]);
  }
}

function norm_(v) { return String(v == null ? '' : v).trim(); }
function lower_(v) { return norm_(v).toLowerCase(); }
function now_()  { return new Date(); }

/* ------------------------------------------------------------------- auth */

/** Team tab BUSINESSES column: blank = both; otherwise "Uribest", "UFFG", or "UFFG, Uribest". */
function allowedBiz_(v) {
  var l = lower_(v), out = [];
  if (l.indexOf('uffg') >= 0) out.push('uffg');
  if (l.indexOf('uribest') >= 0) out.push('uribest');
  return out.length ? out : ['uffg', 'uribest'];
}

/** Add any header a newer version expects to a tab that already exists (e.g. BUSINESSES on Team). */
function ensureHeaders_(name) {
  var sh = tab_(name), h = headers_(sh), add = TABS[name].filter(function (c) { return h[c] === undefined; });
  if (!add.length) return;
  var col = sh.getLastColumn() + 1;
  sh.getRange(1, col, 1, add.length).setValues([add]).setFontWeight('bold');
}

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
        isManager: lower_(t[i]['ROLE']) === 'manager',
        businesses: allowedBiz_(t[i]['BUSINESSES'])
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
             base: norm_(t['TRUCK/LOCATION']), role: norm_(t['ROLE']) || 'Lead', businesses: allowedBiz_(t['BUSINESSES']) };
  });

  var log = rows_('Log').slice(-120).reverse().map(function (l) {
    return { when: l['WHEN'] ? new Date(l['WHEN']).toISOString() : '', who: norm_(l['WHO']),
             action: norm_(l['ACTION']), kind: norm_(l['KIND']), item: norm_(l['ITEM']),
             from: norm_(l['FROM']), to: norm_(l['TO']), qty: l['QTY'], note: norm_(l['NOTE']) };
  });

  var out = {
    version: VERSION, me: me, acting: acting, team: team,
    locations: locs, tools: tools, parts: parts, stock: stock, log: log,
    statuses: STATUSES, serverTime: now_().toISOString()
  };
  if (me.isManager && acting.isManager) out.biz = business_(acting.businesses);
  return out;
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

/* ============================================================ business */
/* Customers, price book, estimates, change orders, invoices, payments.
   Managers only. Everything lives in text-formatted tabs.             */

var DOC_STATUSES = {
  estimate: ['Draft','Sent','Approved','Declined'],
  change:   ['Draft','Sent','Approved','Declined'],
  invoice:  ['Draft','Sent','Partial','Paid','Void']
};

/* Two businesses share one app. Everything money-related is tagged with one of these. */
var BUSINESSES = {
  uffg: {
    prefixE: 'E-', prefixI: 'INV-',
    defaults: {
      name: 'UFFG Construction LLC', tagline: '', contact: 'Ricardo Uribe', address: '', phone: '',
      email: 'ricardo@uffgconstruction.com', license: '', color: '#2E6B4F',
      taxRate: '0', estimateDays: '30', invoiceDays: '0',
      estimateTerms: 'This estimate is valid for 30 days. Work is scheduled after signed approval.',
      invoiceTerms: 'Thank you for your business.',
      changeTerms: 'This change order becomes part of the original agreement once signed.'
    }
  },
  uribest: {
    prefixE: 'UB-E-', prefixI: 'UB-INV-',
    defaults: {
      name: 'Uribest Carports & Repairs LLC', tagline: 'Built to Last', contact: 'Ricardo Uribe, Project Manager',
      address: 'Based in Texas', phone: '', email: 'uribestcarportsllc@gmail.com', license: '', color: '#1F4FB4',
      taxRate: '0', estimateDays: '30', invoiceDays: '0',
      estimateTerms: 'This estimate is valid for 30 days. Work is scheduled after signed approval.',
      invoiceTerms: 'Thank you for your business.',
      changeTerms: 'This change order becomes part of the original agreement once signed.'
    }
  }
};
function bizKey_(v) { return lower_(v) === 'uribest' ? 'uribest' : 'uffg'; }
/* Chart of accounts. 5000s are job costs (cost of goods sold), 6000s are overhead. */
var ACCOUNTS = [
  { code: '5000', name: 'Job materials',                type: 'cogs' },
  { code: '5100', name: 'Subcontractors',               type: 'cogs' },
  { code: '5200', name: 'Job labor',                    type: 'cogs' },
  { code: '5300', name: 'Equipment rental',             type: 'cogs' },
  { code: '5400', name: 'Permits & inspections',        type: 'cogs' },
  { code: '5500', name: 'Dump & disposal',              type: 'cogs' },
  { code: '5600', name: 'Job travel & lodging',         type: 'cogs' },
  { code: '5700', name: 'Job supplies & small tools',   type: 'cogs' },
  { code: '6000', name: 'Fuel',                         type: 'expense' },
  { code: '6100', name: 'Vehicle repairs & maintenance', type: 'expense' },
  { code: '6200', name: 'Tools & small equipment',      type: 'expense' },
  { code: '6300', name: 'Insurance',                    type: 'expense' },
  { code: '6400', name: 'Licenses & registrations',     type: 'expense' },
  { code: '6500', name: 'Office, software & phone',     type: 'expense' },
  { code: '6600', name: 'Meals',                        type: 'expense' },
  { code: '6700', name: 'Advertising & marketing',      type: 'expense' },
  { code: '6800', name: 'Legal & accounting',           type: 'expense' },
  { code: '6900', name: 'Bank & card fees',             type: 'expense' },
  { code: '6990', name: 'Other expenses',               type: 'expense' }
];
function accountName_(v) {
  v = norm_(v);
  for (var i = 0; i < ACCOUNTS.length; i++) if (ACCOUNTS[i].name === v || ACCOUNTS[i].code === v) return ACCOUNTS[i].name;
  return v ? 'Other expenses' : '';
}
function vendorKey_(v) { return lower_(v).replace(/^the\s+/, '').replace(/[^a-z0-9]/g, '').slice(0, 14); }

function mustManage_(me) {
  if (!me.isManager) throw new Error('Only a manager can do that.');
}

function num_(v) { var n = Number(v); return isFinite(n) ? n : 0; }
function r2_(n) { return Math.round(num_(n) * 100) / 100; }

function dstr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return norm_(v);
}

/** Same math as the app: discount comes off the subtotal, tax only on taxable lines. */
function totals_(lines, taxRate, discount) {
  var sub = 0, taxable = 0;
  (lines || []).forEach(function (l) {
    if (l.heading) return;
    var amt = r2_(num_(l.qty) * num_(l.rate));
    sub += amt;
    if (l.taxable) taxable += amt;
  });
  sub = r2_(sub);
  var disc = Math.min(Math.max(0, r2_(discount)), sub);
  var taxBase = sub ? taxable * (1 - disc / sub) : 0;
  var tax = r2_(taxBase * num_(taxRate) / 100);
  return { subtotal: sub, discount: disc, tax: tax, total: r2_(sub - disc + tax) };
}

function cleanLines_(lines) {
  if (!Array.isArray(lines)) return [];
  return lines.slice(0, 300).map(function (l) {
    if (l && l.heading) return { heading: true, name: norm_(l.name).slice(0, 200) };
    return {
      name: norm_(l && l.name).slice(0, 200),
      desc: norm_(l && l.desc).slice(0, 2000),
      qty: num_(l && l.qty),
      unit: norm_(l && l.unit).slice(0, 20),
      rate: r2_(l && l.rate),
      taxable: !!(l && l.taxable),
      service: norm_(l && l.service)
    };
  }).filter(function (l) { return l.heading ? l.name : (l.name || l.desc); });
}

/** Write a whole row in one call (much faster than one cell at a time). */
function writeRow_(name, rowNum, obj) {
  var sh = tab_(name), h = headers_(sh), width = sh.getLastColumn();
  var rng = sh.getRange(rowNum, 1, 1, width);
  var cur = rng.getValues()[0];
  for (var k in obj) {
    var key = k.toUpperCase();
    if (h[key] === undefined) continue;
    cur[h[key]] = txt_(obj[k]);
  }
  cur = cur.map(txt_);
  rng.setNumberFormat('@').setValues([cur]);
  moneyCells_(name, sh, h, rowNum, cur);
}

function findBy_(name, col, val) {
  var all = rows_(name);
  for (var i = 0; i < all.length; i++) if (norm_(all[i][col]) === norm_(val)) return all[i];
  return null;
}

/* ---- company ---- */

/** { uffg: {...}, uribest: {...} } — saved values over defaults. Keys are stored as "uribest.name". */
function company_() {
  var out = {};
  for (var b in BUSINESSES) {
    out[b] = {};
    for (var k in BUSINESSES[b].defaults) out[b][k] = BUSINESSES[b].defaults[k];
  }
  rows_('Company').forEach(function (r) {
    var key = norm_(r['KEY']); if (!key) return;
    var m = key.match(/^(uffg|uribest)\.(.+)$/), b = m ? m[1] : 'uffg', k = m ? m[2] : key;
    if (out[b].hasOwnProperty(k)) out[b][k] = norm_(r['VALUE']);
  });
  return out;
}

function saveCompany_(a, me) {
  mustManage_(me);
  var b = bizKey_(a.business), vals = a.values || {};
  for (var k in vals) {
    if (!BUSINESSES[b].defaults.hasOwnProperty(k)) continue;
    var key = b + '.' + k;
    var r = findBy_('Company', 'KEY', key);
    var v = norm_(vals[k]).slice(0, 2000);
    if (r) writeRow_('Company', r._row, { 'VALUE': v });
    else append_('Company', { 'KEY': key, 'VALUE': v });
  }
  log_(me.name, 'Updated company info', 'Company', BUSINESSES[b].defaults.name, '', '', '', '');
  return { ok: true };
}

/* ---- customers ---- */

function saveCustomer_(a, me) {
  mustManage_(me);
  var id = norm_(a.id) || Utilities.getUuid();
  var name = norm_(a.name);
  if (!name) throw new Error('A customer needs a name.');
  var rec = {
    'ID': id, 'BUSINESS': bizKey_(a.business), 'NAME': name.slice(0, 200), 'COMPANY': norm_(a.company).slice(0, 200),
    'PHONE': norm_(a.phone).slice(0, 50), 'EMAIL': norm_(a.email).slice(0, 200),
    'ADDRESS': norm_(a.address).slice(0, 500), 'NOTES': norm_(a.notes).slice(0, 2000), 'ACTIVE': 'YES'
  };
  var r = findBy_('Customers', 'ID', id);
  if (r) { delete rec['BUSINESS']; writeRow_('Customers', r._row, rec); }
  else {
    rec['CREATED'] = now_().toISOString();
    append_('Customers', rec);
    log_(me.name, 'Added customer', 'Customer', name, '', '', '', '');
  }
  return { ok: true, id: id };
}

function archiveCustomer_(a, me) {
  mustManage_(me);
  var r = findBy_('Customers', 'ID', a.id);
  if (!r) throw new Error('Customer not found.');
  writeRow_('Customers', r._row, { 'ACTIVE': 'NO' });
  log_(me.name, 'Archived customer', 'Customer', norm_(r['NAME']), '', '', '', '');
  return { ok: true };
}

/* ---- price book ---- */

function saveService_(a, me) {
  mustManage_(me);
  var id = norm_(a.id) || Utilities.getUuid();
  var name = norm_(a.name);
  if (!name) throw new Error('A service needs a name.');
  var rec = {
    'ID': id, 'BUSINESS': bizKey_(a.business), 'NAME': name.slice(0, 200), 'DESCRIPTION': norm_(a.description).slice(0, 2000),
    'UNIT': norm_(a.unit).slice(0, 20), 'PRICE': r2_(a.price), 'TAXABLE': a.taxable ? 'YES' : 'NO', 'ACTIVE': 'YES'
  };
  var r = findBy_('Services', 'ID', id);
  if (r) { delete rec['BUSINESS']; writeRow_('Services', r._row, rec); }
  else append_('Services', rec);
  return { ok: true, id: id };
}

function archiveService_(a, me) {
  mustManage_(me);
  var r = findBy_('Services', 'ID', a.id);
  if (!r) throw new Error('Service not found.');
  writeRow_('Services', r._row, { 'ACTIVE': 'NO' });
  return { ok: true };
}

/* ---- documents ---- */

function docRow_(id) {
  var r = findBy_('Docs', 'ID', id);
  return (r && lower_(r['DELETED']) !== 'yes') ? r : null;
}

function nextNumber_(type, parentRow, business) {
  var all = rows_('Docs');
  if (type === 'change') {
    var base = norm_(parentRow['NUMBER']), n = 0;
    all.forEach(function (d) {
      if (norm_(d['TYPE']) === 'change' && norm_(d['PARENT']) === norm_(parentRow['ID'])) {
        var m = norm_(d['NUMBER']).match(/CO(\d+)$/);
        if (m) n = Math.max(n, Number(m[1]));
      }
    });
    return base + '-CO' + (n + 1);
  }
  var B = BUSINESSES[bizKey_(business)];
  var prefix = type === 'invoice' ? B.prefixI : B.prefixE, max = 1000;
  all.forEach(function (d) {
    if (norm_(d['TYPE']) !== type || bizKey_(d['BUSINESS']) !== bizKey_(business)) return;
    var m = norm_(d['NUMBER']).match(/(\d+)$/);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return prefix + (max + 1);
}

function saveDoc_(a, me) {
  mustManage_(me);
  var d = a.doc || {};
  var id = norm_(d.id);
  if (!id) throw new Error('Missing document id.');
  var existing = docRow_(id);
  var type = existing ? norm_(existing['TYPE']) : norm_(d.type);
  if (!DOC_STATUSES[type]) throw new Error('Unknown document type.');
  if (!norm_(d.customer)) throw new Error('Pick a customer first.');

  var parent = null;
  if (type === 'change') {
    parent = docRow_(existing ? existing['PARENT'] : d.parent);
    if (!parent) throw new Error('A change order needs its estimate.');
  }

  var lines = cleanLines_(d.lines);
  var t = totals_(lines, d.taxRate, d.discount);
  var rec = {
    'CUSTOMER': norm_(d.customer), 'TITLE': norm_(d.title).slice(0, 200), 'SITE': norm_(d.site).slice(0, 500),
    'DATE': norm_(d.date), 'DUE': norm_(d.due), 'LINES': JSON.stringify(lines),
    'TAX RATE': num_(d.taxRate), 'DISCOUNT': t.discount, 'DEPOSIT': num_(d.deposit),
    'NOTES': norm_(d.notes).slice(0, 4000), 'TERMS': norm_(d.terms).slice(0, 4000),
    'SUBTOTAL': t.subtotal, 'TAX': t.tax, 'TOTAL': t.total, 'UPDATED': now_().toISOString()
  };

  var number;
  if (existing) {
    number = norm_(existing['NUMBER']);
    writeRow_('Docs', existing._row, rec);
  } else {
    var business = parent ? bizKey_(parent['BUSINESS']) : bizKey_(d.business);
    number = nextNumber_(type, parent, business);
    rec['ID'] = id; rec['BUSINESS'] = business; rec['TYPE'] = type; rec['NUMBER'] = number;
    rec['STATUS'] = 'Draft'; rec['PARENT'] = parent ? norm_(parent['ID']) : norm_(d.parent);
    rec['PAID'] = 0; rec['BILLED'] = ''; rec['CREATED BY'] = me.name;
    rec['CREATED'] = now_().toISOString(); rec['DELETED'] = '';
    append_('Docs', rec);
    log_(me.name, 'Created ' + (type === 'change' ? 'change order' : type), 'Doc', number, '', '', '', '');
  }

  // An invoice built from an estimate (and its change orders) marks them billed.
  if (type === 'invoice' && Array.isArray(a.billFrom)) {
    a.billFrom.forEach(function (sid) {
      var s = docRow_(sid);
      if (s && !norm_(s['BILLED'])) writeRow_('Docs', s._row, { 'BILLED': id });
    });
  }
  if (type === 'invoice') recalcPaid_(id);
  return { ok: true, id: id, number: number };
}

function setDocStatus_(a, me) {
  mustManage_(me);
  var r = docRow_(a.id);
  if (!r) throw new Error('Document not found.');
  var type = norm_(r['TYPE']), st = norm_(a.status);
  if (DOC_STATUSES[type].indexOf(st) < 0) throw new Error('That status does not apply here.');
  var upd = { 'STATUS': st, 'UPDATED': now_().toISOString() };
  if (st === 'Sent' && !norm_(r['SENT AT'])) upd['SENT AT'] = now_().toISOString();
  if (st === 'Approved') upd['APPROVED AT'] = now_().toISOString();
  writeRow_('Docs', r._row, upd);
  log_(me.name, 'Marked ' + st, 'Doc', norm_(r['NUMBER']), norm_(r['STATUS']), st, '', '');
  if (type === 'invoice' && st !== 'Void' && st !== 'Draft') recalcPaid_(a.id);
  return { ok: true };
}

function deleteDoc_(a, me) {
  mustManage_(me);
  var r = docRow_(a.id);
  if (!r) return { ok: true };
  var id = norm_(r['ID']), type = norm_(r['TYPE']);
  writeRow_('Docs', r._row, { 'DELETED': 'YES', 'UPDATED': now_().toISOString() });
  rows_('Docs').forEach(function (d) {
    if (lower_(d['DELETED']) === 'yes') return;
    if (norm_(d['BILLED']) === id) writeRow_('Docs', d._row, { 'BILLED': '' });         // free up what this invoice billed
    if (type === 'estimate' && norm_(d['TYPE']) === 'change' && norm_(d['PARENT']) === id)
      writeRow_('Docs', d._row, { 'DELETED': 'YES' });                                     // its change orders go too
  });
  log_(me.name, 'Deleted', 'Doc', norm_(r['NUMBER']), '', '', '', '');
  return { ok: true };
}

/* ---- payments ---- */

function recalcPaid_(invoiceId) {
  var r = docRow_(invoiceId);
  if (!r || norm_(r['TYPE']) !== 'invoice') return;
  var paid = 0;
  rows_('Payments').forEach(function (p) {
    if (norm_(p['INVOICE']) === norm_(invoiceId) && lower_(p['DELETED']) !== 'yes') paid += num_(p['AMOUNT']);
  });
  paid = r2_(paid);
  var total = num_(r['TOTAL']), st = norm_(r['STATUS']);
  if (st !== 'Void') {
    if (paid > 0 && paid >= total) st = 'Paid';
    else if (paid > 0) st = 'Partial';
    else if (st === 'Paid' || st === 'Partial') st = 'Sent';
  }
  writeRow_('Docs', r._row, { 'PAID': paid, 'STATUS': st });
}

function addPayment_(a, me) {
  mustManage_(me);
  var inv = docRow_(a.invoice);
  if (!inv || norm_(inv['TYPE']) !== 'invoice') throw new Error('Invoice not found.');
  var amt = r2_(a.amount);
  if (amt <= 0) throw new Error('Enter the amount paid.');
  var id = norm_(a.id) || Utilities.getUuid();
  if (findBy_('Payments', 'ID', id)) return { ok: true };
  append_('Payments', {
    'ID': id, 'INVOICE': norm_(a.invoice), 'DATE': norm_(a.date), 'AMOUNT': amt,
    'METHOD': norm_(a.method).slice(0, 50), 'NOTE': norm_(a.note).slice(0, 500), 'WHO': me.name, 'DELETED': '',
    'BANK': bankFor_(a.bank, bizKey_(inv['BUSINESS']))
  });
  recalcPaid_(a.invoice);
  log_(me.name, 'Payment', 'Doc', norm_(inv['NUMBER']), '', '', String(amt), norm_(a.method));
  return { ok: true };
}

function deletePayment_(a, me) {
  mustManage_(me);
  var p = findBy_('Payments', 'ID', a.id);
  if (!p) return { ok: true };
  writeRow_('Payments', p._row, { 'DELETED': 'YES' });
  recalcPaid_(norm_(p['INVOICE']));
  return { ok: true };
}

/* ---- expenses ---- */

/** Drive: UFFG App Receipts / UFFG (or Uribest) / 2026-11 */
function receiptsFolder_(business, ymd) {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('RECEIPTS_FOLDER'), root = null;
  if (id) { try { root = DriveApp.getFolderById(id); } catch (e) {} }
  if (!root) { root = DriveApp.createFolder('UFFG App Receipts'); props.setProperty('RECEIPTS_FOLDER', root.getId()); }
  if (!business) return root;
  var b = childFolder_(root, business === 'uribest' ? 'Uribest' : 'UFFG');
  return childFolder_(b, String(ymd || dstr_(now_())).slice(0, 7));
}
function childFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
/** File name + folder follow the expense, so Drive stays tidy for the accountant. */
function fileReceipt_(fileId, business, rec) {
  try {
    var f = DriveApp.getFileById(fileId);
    var label = [rec['DATE'], rec['VENDOR'] || rec['ACCOUNT'] || 'Receipt', '$' + r2_(rec['AMOUNT']).toFixed(2)].filter(String).join(' ')
      .replace(/[\\/:*?"<>|]+/g, '');
    f.setName(label + '.jpg');
    f.moveTo(receiptsFolder_(business, rec['DATE']));
    return f.getUrl();
  } catch (e) { return ''; }
}

/* ---- receipt reading ---- */

/** Google's own OCR: convert the photo into a temporary Google Doc and read its text. Needs the Drive advanced service. */
function ocr_(blob) {
  if (typeof Drive === 'undefined' || !Drive.Files) return '';
  var tmp = null;
  try {
    tmp = Drive.Files.create({ name: 'receipt-ocr-tmp', mimeType: 'application/vnd.google-apps.document' }, blob, { ocrLanguage: 'en', fields: 'id' });
    return DocumentApp.openById(tmp.id).getBody().getText() || '';
  } catch (e) {
    return '';
  } finally {
    if (tmp && tmp.id) { try { DriveApp.getFileById(tmp.id).setTrashed(true); } catch (e2) {} }
  }
}

var KNOWN_VENDORS = [
  ['home depot', 'The Home Depot', 'Job materials'], ["lowe's", "Lowe's", 'Job materials'], ['lowes', "Lowe's", 'Job materials'],
  ['menards', 'Menards', 'Job materials'], ['ace hardware', 'Ace Hardware', 'Job materials'], ['84 lumber', '84 Lumber', 'Job materials'],
  ['mccoy', "McCoy's", 'Job materials'], ['builders firstsource', 'Builders FirstSource', 'Job materials'], ['sherwin', 'Sherwin-Williams', 'Job materials'],
  ['ferguson', 'Ferguson', 'Job materials'], ['fastenal', 'Fastenal', 'Job materials'], ['grainger', 'Grainger', 'Job materials'],
  ['metal mart', 'Metal Mart', 'Job materials'], ['metals depot', 'Metals Depot', 'Job materials'], ['tractor supply', 'Tractor Supply', 'Job materials'],
  ['harbor freight', 'Harbor Freight', 'Tools & small equipment'], ['northern tool', 'Northern Tool', 'Tools & small equipment'],
  ['sunbelt', 'Sunbelt Rentals', 'Equipment rental'], ['united rentals', 'United Rentals', 'Equipment rental'], ['herc', 'Herc Rentals', 'Equipment rental'],
  ['buc-ee', "Buc-ee's", 'Fuel'], ['bucee', "Buc-ee's", 'Fuel'], ['shell', 'Shell', 'Fuel'], ['exxon', 'Exxon', 'Fuel'], ['mobil', 'Mobil', 'Fuel'],
  ['chevron', 'Chevron', 'Fuel'], ['valero', 'Valero', 'Fuel'], ['quiktrip', 'QuikTrip', 'Fuel'], ['qt ', 'QuikTrip', 'Fuel'], ['murphy', 'Murphy USA', 'Fuel'],
  ['circle k', 'Circle K', 'Fuel'], ['7-eleven', '7-Eleven', 'Fuel'], ['racetrac', 'RaceTrac', 'Fuel'], ['love\'s', "Love's", 'Fuel'], ['pilot', 'Pilot', 'Fuel'],
  ['sunoco', 'Sunoco', 'Fuel'], ['conoco', 'Conoco', 'Fuel'], ['phillips 66', 'Phillips 66', 'Fuel'], ['stripes', 'Stripes', 'Fuel'],
  ["o'reilly", "O'Reilly Auto Parts", 'Vehicle repairs & maintenance'], ['autozone', 'AutoZone', 'Vehicle repairs & maintenance'], ['advance auto', 'Advance Auto Parts', 'Vehicle repairs & maintenance'],
  ['walmart', 'Walmart', 'Job materials'], ['costco', 'Costco', 'Job materials'], ["sam's club", "Sam's Club", 'Job materials'],
  ['whataburger', 'Whataburger', 'Meals'], ['mcdonald', "McDonald's", 'Meals'], ['taco', '', 'Meals'], ['chick-fil-a', 'Chick-fil-A', 'Meals']
];

/** Pull vendor, date, total, payment type out of OCR text. Best guess; the person confirms before saving. */
function parseReceipt_(text) {
  var out = { vendor: '', date: '', total: 0, account: '', method: '' };
  var t = String(text || '').replace(/\r/g, '');
  if (!t.trim()) return out;
  var low = t.toLowerCase();
  var lines = t.split('\n').map(function (l) { return l.trim(); }).filter(String);

  for (var i = 0; i < KNOWN_VENDORS.length; i++) {
    if (low.indexOf(KNOWN_VENDORS[i][0]) >= 0) { out.vendor = KNOWN_VENDORS[i][1]; out.account = KNOWN_VENDORS[i][2]; break; }
  }
  if (!out.vendor) {
    for (var j = 0; j < Math.min(lines.length, 8); j++) {
      var l = lines[j];
      if (/[a-z]{3,}/i.test(l) && !/\d{3}[-.)\s]*\d{3}[-.\s]*\d{4}/.test(l) && !/^\d+\s+\w/.test(l) &&
          !/(receipt|welcome|store|thank|invoice|order|date|www\.|\.com|tel|phone)/i.test(l)) { out.vendor = l.replace(/[^\w&'.\- ]/g, '').trim().slice(0, 60); break; }
    }
  }

  var amt = /\$?\s*(\d{1,3}(?:,\d{3})+|\d+)\.(\d{2})(?!\d)/g;
  var amountsIn = function (l) { var a = [], m; amt.lastIndex = 0; while ((m = amt.exec(l))) a.push(Number(m[1].replace(/,/g, '') + '.' + m[2])); return a; };
  var best = 0;
  // Prefer the last "total"-type line that isn't subtotal / savings / tax.
  for (var k = 0; k < lines.length; k++) {
    var L = lines[k].toLowerCase();
    if (/(sub\s*-?\s*total|savings|you saved|discount|tax\b|change\b|tip\b|points)/.test(L)) continue;
    if (/(grand total|total due|amount due|balance due|total\b|amount\b|sale\b|purchase\b|debit\b|credit\b|visa|mastercard)/.test(L)) {
      var a = amountsIn(lines[k]);
      if (!a.length && k + 1 < lines.length) a = amountsIn(lines[k + 1]);     // amount on the next line
      if (a.length && /total|amount|balance/.test(L)) best = a[a.length - 1];
      else if (a.length && !best) best = a[a.length - 1];
    }
  }
  if (!best) {                                                              // fall back to the biggest number on the receipt
    var all = amountsIn(t).filter(function (n) { return n > 0 && n < 100000; });
    if (all.length) best = Math.max.apply(null, all);
  }
  out.total = r2_(best);

  var mon = { jan:1, feb:2, mar:3, apr:4, may:5, jun:6, jul:7, aug:8, sep:9, oct:10, nov:11, dec:12 };
  var pad = function (n) { return (n < 10 ? '0' : '') + n; };
  var dm, y, mo, d;
  if ((dm = t.match(/\b(20\d{2})[-\/.](\d{1,2})[-\/.](\d{1,2})\b/))) { y = +dm[1]; mo = +dm[2]; d = +dm[3]; }
  else if ((dm = t.match(/\b(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})\b/))) { mo = +dm[1]; d = +dm[2]; y = +dm[3]; if (y < 100) y += 2000; }
  else if ((dm = t.match(/\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+(\d{1,2}),?\s+(20\d{2})\b/i))) { mo = mon[dm[1].toLowerCase()]; d = +dm[2]; y = +dm[3]; }
  if (y && mo >= 1 && mo <= 12 && d >= 1 && d <= 31) {
    var ds = y + '-' + pad(mo) + '-' + pad(d), todayS = dstr_(now_());
    var age = (new Date(todayS) - new Date(ds)) / 864e5;
    if (age > -2 && age < 800) out.date = ds;
  }

  if (/\b(visa|mastercard|master card|amex|american express|discover|debit|credit|card #|chip|contactless|apple pay|google pay)\b/i.test(t)) out.method = 'Card';
  else if (/\bcash\b/i.test(t)) out.method = 'Cash';
  else if (/\bcheck\b/i.test(t)) out.method = 'Check';
  return out;
}

/** Possible duplicates: the exact same photo, or same vendor + amount within 3 days. */
function findDuplicates_(business, hash, vendor, amount, date, excludeId) {
  var out = [], vk = vendorKey_(vendor), amt = r2_(amount), t = date ? new Date(date).getTime() : NaN;
  rows_('Expenses').forEach(function (x) {
    if (lower_(x['DELETED']) === 'yes' || norm_(x['ID']) === norm_(excludeId) || bizKey_(x['BUSINESS']) !== business) return;
    var reason = '';
    if (hash && norm_(x['HASH']) === hash) reason = 'Same photo';
    else if (amt > 0 && Math.abs(num_(x['AMOUNT']) - amt) < 0.005) {
      var xv = vendorKey_(x['VENDOR']), xt = new Date(dstr_(x['DATE'])).getTime();
      var sameVendor = !vk || !xv || xv === vk;
      var close = isNaN(t) || isNaN(xt) || Math.abs(xt - t) <= 3 * 864e5;
      if (sameVendor && close) reason = 'Same vendor, amount and date';
    }
    if (reason) out.push({ id: norm_(x['ID']), vendor: norm_(x['VENDOR']), amount: num_(x['AMOUNT']), date: dstr_(x['DATE']), reason: reason });
  });
  return out;
}

/** The account this vendor was booked to last time, if any. */
function lastAccountFor_(business, vendor) {
  var vk = vendorKey_(vendor); if (!vk) return '';
  var all = rows_('Expenses'), best = '';
  all.forEach(function (x) {
    if (lower_(x['DELETED']) !== 'yes' && bizKey_(x['BUSINESS']) === business && norm_(x['STATUS']) !== 'Review' &&
        vendorKey_(x['VENDOR']) === vk && norm_(x['ACCOUNT'])) best = norm_(x['ACCOUNT']);
  });
  return best;
}

function expenseOut_(x) {
  return { id: norm_(x['ID']), business: bizKey_(x['BUSINESS']), status: norm_(x['STATUS']) || 'Booked', date: dstr_(x['DATE']),
           vendor: norm_(x['VENDOR']), account: norm_(x['ACCOUNT']), amount: num_(x['AMOUNT']), method: norm_(x['METHOD']),
           job: norm_(x['JOB']), note: norm_(x['NOTE']), receipt: !!(norm_(x['RECEIPT']) || norm_(x['RECEIPT LINK'])), link: norm_(x['RECEIPT LINK']),
           hash: norm_(x['HASH']), dupOf: norm_(x['DUP OF']), scanned: !!norm_(x['SCANNED']),
           createdBy: norm_(x['CREATED BY']), created: norm_(x['CREATED']),
           file: !!norm_(x['RECEIPT']), bank: norm_(x['BANK']), paidBy: norm_(x['PAID BY']), checkNo: norm_(x['CHECK NO']),
           checkFile: !!norm_(x['CHECK FILE']), checkLink: norm_(x['CHECK LINK']) };
}

/** Save the photo to Drive right away and read it. With createDraft, also file it as an expense "To review". */
function scanReceipt_(a, me) {
  mustManage_(me);
  var img = a.image || {};
  if (!img.data) throw new Error('No photo received.');
  var b = bizKey_(a.business), hash = norm_(a.hash).slice(0, 80);
  var blob = Utilities.newBlob(Utilities.base64Decode(img.data), img.type || 'image/jpeg',
    'Receipt ' + dstr_(now_()) + ' ' + Utilities.formatDate(now_(), Session.getScriptTimeZone(), 'HHmmss') + '.jpg');
  var text = ocr_(blob);
  var p = parseReceipt_(text);
  var dups = findDuplicates_(b, hash, p.vendor, p.total, p.date, a.id);
  var account = lastAccountFor_(b, p.vendor) || p.account;
  if (a.skipIfSamePhoto && dups.some(function (d) { return d.reason === 'Same photo'; })) {
    return { skipped: true, duplicates: dups, vendor: p.vendor, total: p.total, date: p.date };   // nothing saved
  }
  var file = receiptsFolder_(b, p.date || dstr_(now_())).createFile(blob);
  var out = { fileId: file.getId(), link: file.getUrl(), text: text.slice(0, 3000), read: !!text, duplicates: dups,
              vendor: p.vendor, date: p.date, total: p.total, account: account, method: p.method };
  if (a.createDraft) {
    var id = norm_(a.id) || Utilities.getUuid();
    var rec = { 'ID': id, 'BUSINESS': b, 'STATUS': 'Review', 'DATE': p.date || dstr_(now_()), 'VENDOR': p.vendor,
      'ACCOUNT': account, 'AMOUNT': p.total, 'METHOD': p.method, 'JOB': '', 'NOTE': '', 'RECEIPT': out.fileId,
      'HASH': hash, 'DUP OF': dups.length ? dups[0].id : '', 'SCANNED': out.text, 'CREATED BY': me.name,
      'CREATED': now_().toISOString(), 'UPDATED': now_().toISOString(), 'DELETED': '' };
    rec['RECEIPT LINK'] = fileReceipt_(out.fileId, b, { 'DATE': rec['DATE'], 'VENDOR': p.vendor, 'AMOUNT': p.total }) || out.link;
    append_('Expenses', rec);
    out.expense = expenseOut_(findBy_('Expenses', 'ID', id));
  }
  return out;
}

function saveExpense_(a, me) {
  mustManage_(me);
  var id = norm_(a.id) || Utilities.getUuid();
  var status = norm_(a.status) === 'Review' ? 'Review' : 'Booked';
  var amt = r2_(a.amount);
  if (status === 'Booked' && !amt) throw new Error('Enter the amount.');
  var existing = findBy_('Expenses', 'ID', id);
  if (existing && lower_(existing['DELETED']) === 'yes') return { ok: true };
  var rec = {
    'STATUS': status, 'DATE': norm_(a.date), 'VENDOR': norm_(a.vendor).slice(0, 200),
    'ACCOUNT': accountName_(a.account) || (status === 'Booked' ? 'Other expenses' : ''),
    'AMOUNT': amt, 'METHOD': norm_(a.method).slice(0, 50), 'JOB': norm_(a.job), 'NOTE': norm_(a.note).slice(0, 1000),
    'UPDATED': now_().toISOString()
  };
  var b = bizKey_(existing ? existing['BUSINESS'] : a.business);
  if (a.bank !== undefined) { rec['BANK'] = bankFor_(a.bank, b); }
  if (a.paidBy !== undefined) rec['PAID BY'] = rec['BANK'] ? '' : norm_(a.paidBy).slice(0, 100);
  if (a.checkNo !== undefined) rec['CHECK NO'] = norm_(a.checkNo).replace(/^#/, '').slice(0, 20);
  if (a.checkImage && a.checkImage.data) {
    rec['CHECK FILE'] = receiptsFolder_(b, rec['DATE']).createFile(Utilities.newBlob(Utilities.base64Decode(a.checkImage.data), a.checkImage.type || 'image/jpeg', 'check.jpg')).getId();
  } else if (norm_(a.checkFile)) rec['CHECK FILE'] = norm_(a.checkFile);
  else if (a.removeCheck) { rec['CHECK FILE'] = ''; rec['CHECK LINK'] = ''; }
  var cfid = rec['CHECK FILE'] !== undefined ? rec['CHECK FILE'] : (existing ? norm_(existing['CHECK FILE']) : '');
  if (cfid) rec['CHECK LINK'] = fileCheck_(cfid, b, rec, existing) || (existing ? norm_(existing['CHECK LINK']) : '');
  if (norm_(a.hash)) rec['HASH'] = norm_(a.hash).slice(0, 80);
  if (a.receipt && a.receipt.data) {
    var blob = Utilities.newBlob(Utilities.base64Decode(a.receipt.data), a.receipt.type || 'image/jpeg', 'receipt.jpg');
    rec['RECEIPT'] = receiptsFolder_(b, rec['DATE']).createFile(blob).getId();
  } else if (norm_(a.receiptFile)) {
    rec['RECEIPT'] = norm_(a.receiptFile);
  } else if (a.removeReceipt) {
    rec['RECEIPT'] = ''; rec['RECEIPT LINK'] = '';
  }
  if (norm_(a.scanText)) rec['SCANNED'] = norm_(a.scanText).slice(0, 3000);
  var fid = rec['RECEIPT'] !== undefined ? rec['RECEIPT'] : (existing ? norm_(existing['RECEIPT']) : '');
  if (fid) rec['RECEIPT LINK'] = fileReceipt_(fid, b, rec) || (existing ? norm_(existing['RECEIPT LINK']) : '');
  // duplicate flag: re-checked on every save unless someone said "not a duplicate"
  if (a.notDuplicate) rec['DUP OF'] = 'CLEARED';
  else if (!existing || norm_(existing['DUP OF']) !== 'CLEARED') {
    var d = findDuplicates_(b, rec['HASH'] || (existing ? norm_(existing['HASH']) : ''), rec['VENDOR'], amt, rec['DATE'], id);
    rec['DUP OF'] = d.length ? d[0].id : '';
  }
  if (existing) writeRow_('Expenses', existing._row, rec);
  else {
    rec['ID'] = id; rec['BUSINESS'] = b; rec['CREATED BY'] = me.name; rec['CREATED'] = now_().toISOString(); rec['DELETED'] = '';
    append_('Expenses', rec);
    log_(me.name, 'Expense', 'Expense', rec['VENDOR'] || rec['ACCOUNT'], '', '', String(amt), BUSINESSES[b].defaults.name);
  }
  return { ok: true, id: id, dupOf: rec['DUP OF'] === 'CLEARED' ? '' : rec['DUP OF'] };
}

function deleteExpense_(a, me) {
  mustManage_(me);
  var r = findBy_('Expenses', 'ID', a.id);
  if (r) writeRow_('Expenses', r._row, { 'DELETED': 'YES', 'UPDATED': now_().toISOString() });
  return { ok: true };
}

function getReceipt_(a, me) {
  mustManage_(me);
  var r = findBy_('Expenses', 'ID', a.id);
  var fid = r && norm_(r[a.which === 'check' ? 'CHECK FILE' : 'RECEIPT']);
  if (!fid) throw new Error(a.which === 'check' ? 'No check photo saved for this expense.' : 'No receipt saved for this expense.');
  var blob = DriveApp.getFileById(fid).getBlob();
  return { type: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) };
}

/** Run once from the editor after updating, so Google asks for Drive permission (for receipt photos). */
function authorizeDrive() {
  ensureHeaders_('Team');
  var f = receiptsFolder_();
  Logger.log('Receipts folder ready: ' + f.getName());
  if (typeof Drive !== 'undefined') Logger.log('Receipt reading is on.');
  else Logger.log('Receipt reading is OFF: add the "Drive API" service (Services +) to turn it on.');
  DocumentApp.getActiveDocument;   // makes Google ask for Docs permission too (used to read scanned text)
}

/* ---- banks, cards, checks ---- */

var BANK_TYPES = { checking: 'Checking', savings: 'Savings', card: 'Credit card' };
var BANKING_TYPES = {                       // what each kind of entry does to the bank, and where the other side goes
  contribution: 'Owner put money in', draw: 'Owner took money out', loan_in: 'Loan received', loan_pay: 'Loan payment',
  tax_pay: 'Sales tax paid', transfer: 'Transfer / card payment', reimburse: 'Paid back out-of-pocket expenses'
};

/** A bank id that belongs to this business, or '' — so nobody can pay a Uribest check from a UFFG account. */
function bankFor_(id, business) {
  id = norm_(id); if (!id) return '';
  var r = findBy_('Banks', 'ID', id);
  if (!r || bizKey_(r['BUSINESS']) !== business) throw new Error('That bank account isn’t part of ' + BUSINESSES[business].defaults.name + '.');
  return id;
}
function bankOut_(r) {
  return { id: norm_(r['ID']), business: bizKey_(r['BUSINESS']), name: norm_(r['NAME']), type: BANK_TYPES[lower_(r['TYPE'])] ? lower_(r['TYPE']) : 'checking',
           institution: norm_(r['INSTITUTION']), last4: norm_(r['LAST4']), opening: num_(r['OPENING BALANCE']),
           openingDate: dstr_(r['OPENING DATE']), active: lower_(r['ACTIVE']) !== 'no' };
}

function saveBank_(a, me) {
  mustManage_(me);
  var id = norm_(a.id) || Utilities.getUuid(), name = norm_(a.name);
  if (!name) throw new Error('Give the account a name.');
  var rec = { 'NAME': name.slice(0, 100), 'TYPE': BANK_TYPES[lower_(a.type)] ? lower_(a.type) : 'checking',
    'INSTITUTION': norm_(a.institution).slice(0, 200), 'LAST4': norm_(a.last4).replace(/\D/g, '').slice(-4),
    'OPENING BALANCE': r2_(a.opening), 'OPENING DATE': norm_(a.openingDate), 'ACTIVE': a.active === false ? 'NO' : 'YES' };
  var r = findBy_('Banks', 'ID', id);
  if (r) writeRow_('Banks', r._row, rec);
  else {
    rec['ID'] = id; rec['BUSINESS'] = bizKey_(a.business); rec['CREATED'] = now_().toISOString();
    append_('Banks', rec);
    log_(me.name, 'Added bank account', 'Bank', name, '', '', '', '');
  }
  return { ok: true, id: id };
}

/* ---- vendors & subcontractors ---- */

function saveVendor_(a, me) {
  mustManage_(me);
  var id = norm_(a.id) || Utilities.getUuid(), name = norm_(a.name);
  if (!name) throw new Error('Give the vendor a name.');
  var rec = { 'NAME': name.slice(0, 200), 'TYPE': lower_(a.type) === 'contractor' ? 'contractor' : 'vendor',
    'PHONE': norm_(a.phone).slice(0, 50), 'EMAIL': norm_(a.email).slice(0, 200), 'NOTES': norm_(a.notes).slice(0, 1000), 'ACTIVE': 'YES' };
  var r = findBy_('Vendors', 'ID', id);
  if (!r) {                                                       // same name in the same business = same vendor
    var b = bizKey_(a.business), vk = vendorKey_(name);
    rows_('Vendors').forEach(function (v) { if (!r && bizKey_(v['BUSINESS']) === b && lower_(v['ACTIVE']) !== 'no' && vendorKey_(v['NAME']) === vk) r = v; });
    if (r) id = norm_(r['ID']);
  }
  if (r) writeRow_('Vendors', r._row, rec);
  else { rec['ID'] = id; rec['BUSINESS'] = bizKey_(a.business); rec['CREATED'] = now_().toISOString(); append_('Vendors', rec); }
  return { ok: true, id: id };
}
function archiveVendor_(a, me) {
  mustManage_(me);
  var r = findBy_('Vendors', 'ID', a.id);
  if (r) writeRow_('Vendors', r._row, { 'ACTIVE': 'NO' });
  return { ok: true };
}

/* ---- money in and out of the bank that isn't a sale or an expense ---- */

function saveBankTx_(a, me) {
  mustManage_(me);
  var type = norm_(a.type);
  if (!BANKING_TYPES[type]) throw new Error('Pick what kind of entry this is.');
  var amt = r2_(a.amount);
  if (amt <= 0) throw new Error('Enter the amount.');
  var id = norm_(a.id) || Utilities.getUuid(), existing = findBy_('Banking', 'ID', id);
  var b = bizKey_(existing ? existing['BUSINESS'] : a.business);
  var bank = bankFor_(a.bank, b), to = type === 'transfer' ? bankFor_(a.toBank, b) : '';
  if (!bank) throw new Error('Pick the bank account.');
  if (type === 'transfer' && (!to || to === bank)) throw new Error('Pick the account the money went to.');
  var rec = { 'TYPE': type, 'DATE': norm_(a.date) || dstr_(now_()), 'AMOUNT': amt, 'BANK': bank, 'TO BANK': to,
    'PERSON': norm_(a.person).slice(0, 100), 'NOTE': norm_(a.note).slice(0, 500) };
  if (existing) { if (lower_(existing['DELETED']) === 'yes') return { ok: true }; writeRow_('Banking', existing._row, rec); }
  else {
    rec['ID'] = id; rec['BUSINESS'] = b; rec['WHO'] = me.name; rec['CREATED'] = now_().toISOString(); rec['DELETED'] = '';
    append_('Banking', rec);
    log_(me.name, BANKING_TYPES[type], 'Bank', '', '', '', String(amt), BUSINESSES[b].defaults.name);
  }
  return { ok: true, id: id };
}
function deleteBankTx_(a, me) {
  mustManage_(me);
  var r = findBy_('Banking', 'ID', a.id);
  if (r) writeRow_('Banking', r._row, { 'DELETED': 'YES' });
  return { ok: true };
}
/** Move a customer payment to a different account (or to "not deposited yet"). */
function setPaymentBank_(a, me) {
  mustManage_(me);
  var p = findBy_('Payments', 'ID', a.id);
  if (!p) throw new Error('Payment not found.');
  var inv = docRow_(norm_(p['INVOICE']));
  writeRow_('Payments', p._row, { 'BANK': bankFor_(a.bank, bizKey_(inv ? inv['BUSINESS'] : '')) });
  return { ok: true };
}

/* ---- check photos ---- */

function fileCheck_(fileId, business, rec, existing) {
  try {
    var get = function (k) { return rec[k] !== undefined ? rec[k] : (existing ? existing[k] : ''); };
    var f = DriveApp.getFileById(fileId);
    var label = ['Check', get('CHECK NO') ? '#' + norm_(get('CHECK NO')) : '', dstr_(get('DATE')), norm_(get('VENDOR')), '$' + r2_(get('AMOUNT')).toFixed(2)]
      .filter(String).join(' ').replace(/[\\/:*?"<>|]+/g, '');
    f.setName(label + '.jpg');
    f.moveTo(receiptsFolder_(business, dstr_(get('DATE'))));
    return f.getUrl();
  } catch (e) { return ''; }
}

var NUM_WORDS = { zero:0, one:1, two:2, three:3, four:4, five:5, six:6, seven:7, eight:8, nine:9, ten:10, eleven:11, twelve:12,
  thirteen:13, fourteen:14, fifteen:15, sixteen:16, seventeen:17, eighteen:18, nineteen:19, twenty:20, thirty:30, forty:40,
  fourty:40, fifty:50, sixty:60, seventy:70, eighty:80, ninety:90 };
/** "Three thousand two hundred fifty and 00/100" -> 3250. 0 when it can't tell. */
function wordsAmount_(line) {
  var l = lower_(line).replace(/-/g, ' '), cents = 0, m = l.match(/(\d{1,2})\s*\/\s*100/);
  if (m) cents = Number(m[1]) / 100;
  var words = l.replace(/\d{1,2}\s*\/\s*100.*$/, '').replace(/[^a-z ]/g, ' ').split(/\s+/).filter(String);
  var total = 0, cur = 0, seen = false;
  words.forEach(function (w) {
    if (NUM_WORDS.hasOwnProperty(w)) { cur += NUM_WORDS[w]; seen = true; }
    else if (w === 'hundred') { cur = (cur || 1) * 100; seen = true; }
    else if (w === 'thousand') { total += (cur || 1) * 1000; cur = 0; seen = true; }
  });
  return seen ? r2_(total + cur + cents) : 0;
}

/** Pull check number, date, payee, amount and memo out of a check photo's text. Best guess; the person confirms. */
function parseCheck_(text) {
  var out = { checkNo: '', date: '', payee: '', amount: 0, memo: '' };
  var t = String(text || '').replace(/\r/g, '');
  if (!t.trim()) return out;
  var lines = t.split('\n').map(function (l) { return l.trim(); }).filter(String);
  out.date = parseReceipt_(t).date;

  // payee: after "pay to the order of"
  for (var i = 0; i < lines.length; i++) {
    var m = lines[i].match(/(?:pay\s*to\s*the\s*order\s*of|order\s*of|pay\s*to)\s*[:\-]?\s*(.*)$/i);
    if (m) {
      var p = m[1].replace(/\$.*$/, '').replace(/[\d,]+\.\d{2}.*$/, '').replace(/[^\w&'.,\- ]/g, '').trim();
      if (!p && i + 1 < lines.length) p = lines[i + 1].replace(/\$.*$/, '').replace(/[^\w&'.,\- ]/g, '').trim();
      out.payee = p.slice(0, 80); break;
    }
  }
  // amount: the $ box first, then the written-out line
  var dollars = /\$\s*([\d,]+(?:\.\d{2})?)|(?:^|\s)\**([\d]{1,3}(?:,\d{3})*\.\d{2})\b/g, mm, best = 0;
  while ((mm = dollars.exec(t))) { var v = Number(String(mm[1] || mm[2]).replace(/,/g, '')); if (v > 0 && v < 1000000 && (mm[1] || !best)) best = mm[1] ? v : (best || v); }
  var worded = 0;
  for (var j = 0; j < lines.length; j++) if (/\/\s*100|dollars/i.test(lines[j])) { worded = wordsAmount_(lines[j]); if (worded) break; }
  out.amount = r2_(best || worded);
  out.amountWords = worded;

  // memo
  for (var k = 0; k < lines.length; k++) {
    var mo = lines[k].match(/^(?:memo|for)\s*[:\-]?\s*(.+)$/i);
    if (mo && !/^\W*$/.test(mo[1])) { out.memo = mo[1].replace(/[⑆⑈⑇⑉]/g, '').trim().slice(0, 100); break; }
  }

  // check number: a short number printed top-right that also shows up in the bank line at the bottom
  var nums = function (l) { return (l.replace(/[⑆⑈⑇⑉:;|]/g, ' ').match(/\b\d{3,6}\b/g) || []).map(function (n) { return String(Number(n)); }); };
  var bottom = lines.slice(-2).join(' '), botNums = nums(bottom), cand = [];
  lines.slice(0, Math.min(6, lines.length)).forEach(function (l) {
    if (/\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}|\$|\.\d{2}\b|\d{3}[-.)\s]\d{3}[-.\s]\d{4}|\b(tx|texas|houston|suite|ste|street|st\.|rd|ave|blvd|dr)\b/i.test(l)) return;
    nums(l).forEach(function (n) { cand.push(n); });
  });
  var both = cand.filter(function (n) { return botNums.indexOf(n) >= 0; });
  out.checkNo = both[0] || cand[0] || '';
  if (!out.checkNo && botNums.length) {                       // MICR: routing (9), account, check number — the short one
    var shortOnes = botNums.filter(function (n) { return n.length >= 3 && n.length <= 5; });
    out.checkNo = shortOnes[shortOnes.length - 1] || '';
  }
  return out;
}

/** Expenses this check probably pays: same amount, close in date, no check recorded yet. */
function checkMatches_(business, amount, date, payee, excludeId) {
  var amt = r2_(amount); if (!amt) return [];
  var t = date ? new Date(date).getTime() : NaN, pk = vendorKey_(payee), out = [];
  rows_('Expenses').forEach(function (x) {
    if (lower_(x['DELETED']) === 'yes' || bizKey_(x['BUSINESS']) !== business || norm_(x['ID']) === norm_(excludeId) || norm_(x['CHECK NO'])) return;
    if (Math.abs(num_(x['AMOUNT']) - amt) >= 0.005) return;
    var xt = new Date(dstr_(x['DATE'])).getTime(), days = isNaN(t) || isNaN(xt) ? 0 : Math.abs(xt - t) / 864e5;
    if (days > 60) return;
    var o = expenseOut_(x); o.score = (pk && vendorKey_(x['VENDOR']) === pk ? 2 : 0) + (days <= 7 ? 1 : 0); out.push(o);
  });
  return out.sort(function (a, b) { return b.score - a.score; }).slice(0, 5);
}
function sameCheckNo_(business, bank, checkNo, excludeId) {
  checkNo = norm_(checkNo).replace(/^#/, ''); if (!checkNo) return [];
  return rows_('Expenses').filter(function (x) {
    return lower_(x['DELETED']) !== 'yes' && bizKey_(x['BUSINESS']) === business && norm_(x['ID']) !== norm_(excludeId) &&
      norm_(x['CHECK NO']) === checkNo && (!bank || !norm_(x['BANK']) || norm_(x['BANK']) === bank);
  }).map(expenseOut_);
}

/** Save the check photo to Drive, read it, and find what it pays. */
function scanCheck_(a, me) {
  mustManage_(me);
  var img = a.image || {};
  if (!img.data) throw new Error('No photo received.');
  var b = bizKey_(a.business);
  var blob = Utilities.newBlob(Utilities.base64Decode(img.data), img.type || 'image/jpeg',
    'Check ' + dstr_(now_()) + ' ' + Utilities.formatDate(now_(), Session.getScriptTimeZone(), 'HHmmss') + '.jpg');
  var text = ocr_(blob), p = parseCheck_(text);
  var file = receiptsFolder_(b, p.date || dstr_(now_())).createFile(blob);
  var vendor = p.payee, best = '';
  if (p.payee) {                                               // tidy the payee to a saved vendor name when it's close
    var pk = vendorKey_(p.payee);
    rows_('Vendors').concat(rows_('Expenses')).forEach(function (v) {
      var n = norm_(v['NAME'] || v['VENDOR']);
      if (!best && bizKey_(v['BUSINESS']) === b && n && pk && (vendorKey_(n) === pk || vendorKey_(n).indexOf(pk.slice(0, 6)) === 0)) best = n;
    });
    if (best) vendor = best;
  }
  return { fileId: file.getId(), link: file.getUrl(), text: text.slice(0, 3000), read: !!text,
           checkNo: p.checkNo, date: p.date, payee: vendor, amount: p.amount, amountWords: p.amountWords || 0, memo: p.memo,
           account: lastAccountFor_(b, vendor),
           matches: checkMatches_(b, p.amount, p.date, vendor, a.id), sameNumber: sameCheckNo_(b, norm_(a.bank), p.checkNo, a.id) };
}

/* ---- bring in expenses tracked somewhere else (e.g. a job's own sheet) ---- */

function importExpenses_(a, me) {
  mustManage_(me);
  var b = bizKey_(a.business), list = Array.isArray(a.rows) ? a.rows.slice(0, 500) : [];
  var seen = {};
  rows_('Expenses').forEach(function (x) {
    if (lower_(x['DELETED']) !== 'yes' && bizKey_(x['BUSINESS']) === b)
      seen[[dstr_(x['DATE']), vendorKey_(x['VENDOR']), r2_(x['AMOUNT']).toFixed(2), lower_(x['NOTE']).slice(0, 24)].join('|')] = 1;
  });
  var added = 0, skipped = 0;
  list.forEach(function (r) {
    var amt = r2_(r.amount);
    if (!amt) { skipped++; return; }
    var key = [norm_(r.date), vendorKey_(r.vendor), amt.toFixed(2), lower_(r.note).slice(0, 24)].join('|');
    if (seen[key]) { skipped++; return; }
    seen[key] = 1;
    var link = norm_(r.link);
    var clink = norm_(r.checkLink);
    append_('Expenses', {
      'ID': Utilities.getUuid(), 'BUSINESS': b, 'STATUS': 'Booked', 'DATE': norm_(r.date), 'VENDOR': norm_(r.vendor).slice(0, 200),
      'ACCOUNT': accountName_(r.account) || 'Other expenses', 'AMOUNT': amt, 'METHOD': norm_(r.method).slice(0, 50),
      'JOB': norm_(r.job), 'NOTE': norm_(r.note).slice(0, 1000), 'RECEIPT': '', 'RECEIPT LINK': link,
      'BANK': r.bank ? bankFor_(r.bank, b) : '', 'PAID BY': r.bank ? '' : norm_(r.paidBy).slice(0, 100),
      'CHECK NO': norm_(r.checkNo), 'CHECK FILE': '', 'CHECK LINK': clink,
      'HASH': '', 'DUP OF': 'CLEARED', 'SCANNED': '', 'CREATED BY': me.name, 'CREATED': now_().toISOString(),
      'UPDATED': now_().toISOString(), 'DELETED': ''
    });
    added++;
  });
  log_(me.name, 'Imported expenses', 'Expense', String(added), '', '', '', norm_(a.source).slice(0, 100));
  return { ok: true, added: added, skipped: skipped };
}

/* ---- keep older sheets up to date: new columns, and the Uribest checking account ---- */

function migrate_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('SCHEMA') === VERSION) return;
  for (var t in TABS) { tab_(t); ensureHeaders_(t); }
  if (!props.getProperty('SEED_BANKS')) {
    if (!rows_('Banks').some(function (r) { return bizKey_(r['BUSINESS']) === 'uribest'; })) {
      append_('Banks', { 'ID': Utilities.getUuid(), 'BUSINESS': 'uribest', 'NAME': 'Uribest checking', 'TYPE': 'checking',
        'INSTITUTION': 'Houston Texas Firefighters FCU', 'LAST4': '1882', 'OPENING BALANCE': '', 'OPENING DATE': '', 'ACTIVE': 'YES',
        'CREATED': now_().toISOString() });
    }
    props.setProperty('SEED_BANKS', '1');
  }
  props.setProperty('SCHEMA', VERSION);
}

/* ---- read ---- */

function business_(allowed) {
  allowed = allowed || ['uffg', 'uribest'];
  var ok = function (o) { return allowed.indexOf(o.business) >= 0; };
  var customers = rows_('Customers').filter(function (c) { return lower_(c['ACTIVE']) !== 'no' && norm_(c['ID']); })
    .map(function (c) {
      return { id: norm_(c['ID']), business: bizKey_(c['BUSINESS']), name: norm_(c['NAME']), company: norm_(c['COMPANY']), phone: norm_(c['PHONE']),
               email: norm_(c['EMAIL']), address: norm_(c['ADDRESS']), notes: norm_(c['NOTES']) };
    });
  var services = rows_('Services').filter(function (s) { return lower_(s['ACTIVE']) !== 'no' && norm_(s['ID']); })
    .map(function (s) {
      return { id: norm_(s['ID']), business: bizKey_(s['BUSINESS']), name: norm_(s['NAME']), description: norm_(s['DESCRIPTION']),
               unit: norm_(s['UNIT']), price: num_(s['PRICE']), taxable: lower_(s['TAXABLE']) === 'yes' };
    });
  var docs = rows_('Docs').filter(function (d) { return lower_(d['DELETED']) !== 'yes' && norm_(d['ID']); })
    .map(function (d) {
      var lines = [];
      try { lines = JSON.parse(norm_(d['LINES']) || '[]'); } catch (e) {}
      return {
        id: norm_(d['ID']), business: bizKey_(d['BUSINESS']), type: norm_(d['TYPE']), number: norm_(d['NUMBER']), customer: norm_(d['CUSTOMER']),
        title: norm_(d['TITLE']), site: norm_(d['SITE']), status: norm_(d['STATUS']), date: dstr_(d['DATE']),
        due: dstr_(d['DUE']), parent: norm_(d['PARENT']), lines: lines, taxRate: num_(d['TAX RATE']),
        discount: num_(d['DISCOUNT']), deposit: num_(d['DEPOSIT']), notes: norm_(d['NOTES']), terms: norm_(d['TERMS']),
        subtotal: num_(d['SUBTOTAL']), tax: num_(d['TAX']), total: num_(d['TOTAL']), paid: num_(d['PAID']),
        billed: norm_(d['BILLED']), sentAt: norm_(d['SENT AT']), approvedAt: norm_(d['APPROVED AT']),
        createdBy: norm_(d['CREATED BY']), created: norm_(d['CREATED']), updated: norm_(d['UPDATED'])
      };
    });
  var payments = rows_('Payments').filter(function (p) { return lower_(p['DELETED']) !== 'yes' && norm_(p['ID']); })
    .map(function (p) {
      return { id: norm_(p['ID']), invoice: norm_(p['INVOICE']), date: dstr_(p['DATE']), amount: num_(p['AMOUNT']),
               method: norm_(p['METHOD']), note: norm_(p['NOTE']), who: norm_(p['WHO']), bank: norm_(p['BANK']) };
    });
  customers = customers.filter(ok); services = services.filter(ok); docs = docs.filter(ok);
  var docIds = {}; docs.forEach(function (d) { docIds[d.id] = 1; });
  payments = payments.filter(function (p) { return docIds[p.invoice]; });
  var expenses = rows_('Expenses').filter(function (x) { return lower_(x['DELETED']) !== 'yes' && norm_(x['ID']); }).map(expenseOut_).filter(ok);
  expenses.forEach(function (x) { if (x.dupOf === 'CLEARED') x.dupOf = ''; });
  var co = company_(), coOut = {};
  allowed.forEach(function (b) { coOut[b] = co[b]; });
  var banks = rows_('Banks').filter(function (r) { return norm_(r['ID']); }).map(bankOut_).filter(ok);
  var vendors = rows_('Vendors').filter(function (v) { return norm_(v['ID']) && lower_(v['ACTIVE']) !== 'no'; })
    .map(function (v) { return { id: norm_(v['ID']), business: bizKey_(v['BUSINESS']), name: norm_(v['NAME']), type: norm_(v['TYPE']) || 'vendor',
                                 phone: norm_(v['PHONE']), email: norm_(v['EMAIL']), notes: norm_(v['NOTES']) }; }).filter(ok);
  var banking = rows_('Banking').filter(function (x) { return norm_(x['ID']) && lower_(x['DELETED']) !== 'yes'; })
    .map(function (x) { return { id: norm_(x['ID']), business: bizKey_(x['BUSINESS']), type: norm_(x['TYPE']), date: dstr_(x['DATE']),
                                 amount: num_(x['AMOUNT']), bank: norm_(x['BANK']), toBank: norm_(x['TO BANK']), person: norm_(x['PERSON']),
                                 note: norm_(x['NOTE']), who: norm_(x['WHO']), created: norm_(x['CREATED']) }; }).filter(ok);
  return { company: coOut, businesses: allowed, customers: customers, services: services, docs: docs, payments: payments,
           expenses: expenses, accounts: ACCOUNTS, banks: banks, vendors: vendors, banking: banking, bankingTypes: BANKING_TYPES };
}

/* --------------------------------------------------------------- routing */

var PUBLIC = { requestCode: requestCode_, verifyCode: verifyCode_ };
var BIZ_FNS = { saveCompany:1, saveCustomer:1, archiveCustomer:1, saveService:1, archiveService:1, saveDoc:1, setDocStatus:1,
                deleteDoc:1, addPayment:1, deletePayment:1, saveExpense:1, deleteExpense:1, getReceipt:1, scanReceipt:1,
                saveBank:1, saveVendor:1, archiveVendor:1, saveBankTx:1, deleteBankTx:1, setPaymentBank:1, scanCheck:1, importExpenses:1 };
var NEEDS_ID = { addTool:1, moveTool:1, setToolStatus:1, addPart:1, setLowTarget:1,
                 adjustStock:1, transferStock:1, addLocation:1, archiveLocation:1, bulkImport:1,
                 saveCompany:1, saveCustomer:1, archiveCustomer:1, saveService:1, archiveService:1,
                 saveDoc:1, setDocStatus:1, deleteDoc:1, addPayment:1, deletePayment:1,
                 saveExpense:1, deleteExpense:1,
                 saveBank:1, saveVendor:1, archiveVendor:1, saveBankTx:1, deleteBankTx:1, setPaymentBank:1, importExpenses:1 };

function handlers_() {
  return {
    bootstrap: function (a, me) { return bootstrap_(a, me); },
    signOut: signOut_,
    addTool: addTool_, moveTool: moveTool_, setToolStatus: setToolStatus_,
    addPart: addPart_, setLowTarget: setLowTarget_,
    adjustStock: adjustStock_, transferStock: transferStock_,
    addLocation: addLocation_, archiveLocation: archiveLocation_,
    bulkImport: bulkImport_,
    saveCompany: saveCompany_, saveCustomer: saveCustomer_, archiveCustomer: archiveCustomer_,
    saveService: saveService_, archiveService: archiveService_,
    saveDoc: saveDoc_, setDocStatus: setDocStatus_, deleteDoc: deleteDoc_,
    addPayment: addPayment_, deletePayment: deletePayment_,
    saveExpense: saveExpense_, deleteExpense: deleteExpense_, getReceipt: getReceipt_, scanReceipt: scanReceipt_,
    saveBank: saveBank_, saveVendor: saveVendor_, archiveVendor: archiveVendor_, saveBankTx: saveBankTx_, deleteBankTx: deleteBankTx_,
    setPaymentBank: setPaymentBank_, scanCheck: scanCheck_, importExpenses: importExpenses_
  };
}

/** Which business a call touches — taken from the saved record when there is one, so it can't be spoofed. */
function bizOfCall_(fn, a) {
  a = a || {};
  var rowBiz = function (tab, id) { var r = id ? findBy_(tab, 'ID', id) : null; return r ? bizKey_(r['BUSINESS']) : ''; };
  switch (fn) {
    case 'saveCompany': case 'scanReceipt': case 'scanCheck': case 'importExpenses': return bizKey_(a.business);
    case 'saveBank': return rowBiz('Banks', a.id) || bizKey_(a.business);
    case 'saveVendor': return rowBiz('Vendors', a.id) || bizKey_(a.business);
    case 'archiveVendor': return rowBiz('Vendors', a.id);
    case 'saveBankTx': return rowBiz('Banking', a.id) || bizKey_(a.business);
    case 'deleteBankTx': return rowBiz('Banking', a.id);
    case 'setPaymentBank': var pp = findBy_('Payments', 'ID', a.id); return pp ? rowBiz('Docs', norm_(pp['INVOICE'])) : '';
    case 'saveCustomer': return rowBiz('Customers', a.id) || bizKey_(a.business);
    case 'archiveCustomer': return rowBiz('Customers', a.id);
    case 'saveService': return rowBiz('Services', a.id) || bizKey_(a.business);
    case 'archiveService': return rowBiz('Services', a.id);
    case 'saveDoc':
      var d = a.doc || {};
      return rowBiz('Docs', d.id) || rowBiz('Docs', d.parent) || bizKey_(d.business);
    case 'setDocStatus': case 'deleteDoc': return rowBiz('Docs', a.id);
    case 'addPayment': return rowBiz('Docs', a.invoice);
    case 'deletePayment': var p = findBy_('Payments', 'ID', a.id); return p ? rowBiz('Docs', norm_(p['INVOICE'])) : '';
    case 'saveExpense': return rowBiz('Expenses', a.id) || bizKey_(a.business);
    case 'deleteExpense': case 'getReceipt': return rowBiz('Expenses', a.id);
  }
  return '';
}

function runMigrate_() {
  if (PropertiesService.getScriptProperties().getProperty('SCHEMA') === VERSION) return;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try { migrate_(); } finally { lock.releaseLock(); }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (e && e.parameter && e.parameter.ping) {
    var health = {};
    try {
      runMigrate_();
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
    runMigrate_();

    var h = handlers_()[fn];
    if (!h) return json_({ ok: false, error: 'Unknown action.' });

    // Managers may act as another team member.
    var acting = me;
    if (args.viewAs && me.isManager) {
      var other = teamMember_(args.viewAs);
      if (other) acting = other;
    }

    // Money screens run as the person being viewed: their businesses, their name on what gets saved.
    // Only a real manager can view as someone else, and only a manager-level person can touch money.
    var who = me;
    if (BIZ_FNS[fn]) {
      who = acting === me ? me : { email: acting.email, name: acting.name, base: acting.base, role: acting.role,
                                   isManager: me.isManager && acting.isManager, businesses: acting.businesses };
      var bk = bizOfCall_(fn, args);
      if (bk && who.businesses.indexOf(bk) < 0) throw new Error((who === me ? 'You don’t' : who.name + ' doesn’t') + ' have access to ' + BUSINESSES[bk].defaults.name + '.');
    }

    if (NEEDS_ID[fn]) {
      var lock = LockService.getScriptLock();
      lock.waitLock(25000);
      try {
        if (alreadyDone_(id)) return json_({ ok: true, data: { ok: true, duplicate: true } });
        var out = h(args, who, acting, token);
        markDone_(id);
        return json_({ ok: true, data: out });
      } finally { lock.releaseLock(); }
    }
    return json_({ ok: true, data: h(args, who, acting, token) });

  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

/* ------------------------------------------------- one-time sheet set-up */

function setup() {
  migrate_();
  for (var t in TABS) { tab_(t); ensureHeaders_(t); }
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
