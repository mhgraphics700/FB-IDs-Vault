// FB IDs Vault - Google Sheet backend (Apps Script) + web page.
// Paste this whole file into Extensions > Apps Script, then Deploy > Web app (Execute as: Me, Who has access: Anyone).
// Note: the internal id 'UIDVault' in ping() is kept on purpose so older apps and sheets still recognise each other.
// ====== SETTINGS ======
const DEFAULT_PIN = '006464'; // Default PIN. If you changed the PIN from the app earlier, run resetPin() once from the editor.
const LOW_STOCK = 50;         // warn when remaining IDs fall below this
const VERSION = 6;            // the app checks this number to warn when the sheet script is old
const MAX_QTY = 1500;         // max IDs per order (a Sheets cell holds 50,000 characters)
// Telegram (optional): best to set TG_TOKEN and TG_CHAT in Project Settings > Script properties.
const TG_TOKEN = '';
const TG_CHAT  = '';

// CATEGORIES: the "Batch" column of Stock/Sold holds a 2-letter code.
// 1st letter = name type (E English, M Muslim, H Hindu, X Mix). 2nd = gender (M Male, F Female, X Mix). Blank = XX.
const CAT_RE = /^[EMHX][MFX]$/;
const catOf_ = r => { const c = String(r[1] || '').trim(); return CAT_RE.test(c) ? c : 'XX'; };
const catMatch_ = (r, ty, ge) => { const c = catOf_(r); return (!ty || c[0] === ty) && (!ge || c[1] === ge); };

// PASSWORD: Stock sheet column D = Pass, Sold sheet column F = Pass (each ID keeps its own password).
const SH = {
  stock: ['Stock', ['UID', 'Batch', 'Added', 'Pass']],
  sold:  ['Sold',  ['UID', 'Customer', 'Date', 'Order', 'Batch', 'Pass']],
  sales: ['Sales', ['Order', 'Date', 'Customer', 'Qty', 'Rate', 'Amount', 'UIDs', 'Mode', 'Status', 'OpId']],
  bad:   ['Bad',   ['UID', 'Customer', 'Order', 'Date']],
  returns: ['Returns', ['UID', 'Customer', 'Order', 'Date']],
  ops:   ['Ops',   ['OpId', 'Order']]
};

function doGet() {
  return HtmlService.createHtmlOutput(HTML).setTitle('FB IDs Vault')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {   // used by the offline-capable app (APK)
  let q = null;
  try { q = JSON.parse(e.postData.contents); } catch (x) {}
  return ContentService.createTextOutput(JSON.stringify(api(q))).setMimeType(ContentService.MimeType.JSON);
}

// ---------- sheet helpers ----------
const SHC_ = {};   // sheet objects and header checks are cached for one request (each lookup is slow in Apps Script)
function sh_(k) {
  if (SHC_[k]) return SHC_[k];
  const ss = SpreadsheetApp.getActiveSpreadsheet(), [name, head] = SH[k];
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(head);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else if (sh.getRange(1, head.length).getValue() === '') {
    sh.getRange(1, head.length).setValue(head[head.length - 1]).setFontWeight('bold'); // upgrade older sheets (Pass / OpId column)
  }
  SHC_[k] = sh;
  return sh;
}
function rows_(k) {
  const sh = sh_(k), n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, SH[k][1].length).getValues() : [];
}
function put_(k, rows, start) {
  if (!rows.length) return;
  const sh = sh_(k), end = start + rows.length - 1;
  if (end > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), end - sh.getMaxRows());
  sh.getRange(start, k === 'sales' ? 7 : 1, rows.length, 1).setNumberFormat('@'); // UID = text
  if (k === 'stock') sh.getRange(start, 4, rows.length, 1).setNumberFormat('@');  // password = text
  if (k === 'sold')  sh.getRange(start, 6, rows.length, 1).setNumberFormat('@');
  sh.getRange(start, 1, rows.length, SH[k][1].length).setValues(rows);
}
const push_ = (k, rows) => put_(k, rows, sh_(k).getLastRow() + 1);
function rewrite_(k, rows) {
  const sh = sh_(k), n = sh.getLastRow() - 1;
  if (n > 0) sh.getRange(2, 1, n, SH[k][1].length).clearContent();
  put_(k, rows, 2);
}
// Live stock = Stock tab minus blank rows, duplicates and IDs already sold.
// Top of the list = oldest, bottom = freshest (last added).
function stock_() {
  const sold = new Set(rows_('sold').map(r => String(r[0]).trim())), seen = new Set();
  return rows_('stock').filter(r => {
    const u = String(r[0]).trim();
    if (!u || sold.has(u) || seen.has(u)) return false;
    seen.add(u); return true;
  });
}
function prop_(k, fallback) {
  const v = PropertiesService.getScriptProperties().getProperty(k);
  return v ? v : fallback;
}
function notify_(msg) {
  const tok = prop_('TG_TOKEN', TG_TOKEN), chat = prop_('TG_CHAT', TG_CHAT);
  if (!tok || !chat) return;
  try { UrlFetchApp.fetch('https://api.telegram.org/bot' + tok + '/sendMessage',
    { method: 'post', payload: { chat_id: chat, text: msg }, muteHttpExceptions: true }); } catch (e) {}
}
function lowCheck_(before, after) {   // alert once, when stock crosses below the limit
  if (before >= LOW_STOCK && after < LOW_STOCK) notify_('Low stock: only ' + after + ' IDs left.');
}
function locked_(fn) {
  const l = LockService.getScriptLock(); l.waitLock(20000);
  try { return fn(); } finally { l.releaseLock(); }
}
// Run once from the Apps Script editor (pick this function, then Run): sets the PIN back to the default and clears the 5 minute lock.
function resetPin() {
  PropertiesService.getScriptProperties().setProperty('PIN', DEFAULT_PIN);
  const c = CacheService.getScriptCache(); c.remove('lock'); c.remove('fails');
}
function pin_() {
  const p = PropertiesService.getScriptProperties().getProperty('PIN');
  return p === null ? DEFAULT_PIN : p;
}
function nextOrder_() {   // highest order number + 1 (safe even if rows were deleted)
  let m = 0;
  rows_('sales').forEach(r => { const n = Number(r[0]); if (n > m) m = n; });
  return m + 1;
}

// ---------- API ----------
function api(q) {
  if (!q) return { ok: false, error: 'Bad request' };
  if (q.action === 'ping') return { ok: true, app: 'UIDVault', ver: VERSION };   // connection test from the app
  const PIN = pin_();
  if (PIN) {
    const c = CacheService.getScriptCache();
    if (c.get('lock')) return { ok: false, error: 'Too many wrong PINs. Try again in 5 minutes.' };
    if (!q.pin) return { ok: false, error: 'Enter your PIN' };
    if (String(q.pin) !== PIN) {
      const n = Number(c.get('fails') || 0) + 1;
      c.put('fails', String(n), 600);
      if (n >= 5) { c.put('lock', '1', 300); c.remove('fails'); }
      return { ok: false, error: 'Wrong PIN' };
    }
    c.remove('fails');
  }
  try {
    switch (q.action) {
      case 'login': case 'stats': return { ok: true, stats: stats_() };
      case 'sell':      return locked_(() => sell_(q));
      case 'add':       return locked_(() => add_(q.text, q.cat, q.hasPw, q.front, q.bad));
      case 'about':     return about_();
      case 'setpass':   return locked_(() => setPass_(q.uids, q.pass));
      case 'returns':   return locked_(() => returnMany_(q.uids));
      case 'replacemany': return locked_(() => replaceMany_(q.uids));
      case 'undo':      return locked_(() => undo_(q.order));
      case 'replace':   return locked_(() => replace_(q.uid));
      case 'find':      return find_(q.uid);
      case 'customers': return { ok: true, list: customers_() };
      case 'history':   return { ok: true, history: history_(q) };
      case 'changepin': return locked_(() => changePin_(q));
      case 'sync':      return locked_(() => sync_(q.ops));
      default:          return { ok: false, error: 'Unknown action' };
    }
  } catch (e) { return { ok: false, error: String(e.message || e) }; }
}

// Connection check: tells the app which spreadsheet this script is attached to and what it holds.
function about_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return { ok: true, ver: VERSION, name: ss.getName(), stock: stock_().length, sold: rows_('sold').length, orders: rows_('sales').length };
}

function changePin_(q) {
  const cur = pin_(), np = String(q.newPin || '');
  if (cur && String(q.cur || '') !== cur) return { ok: false, error: 'Current PIN is wrong' };
  if (!/^\d{4,10}$/.test(np)) return { ok: false, error: 'New PIN must be 4 to 10 digits' };
  PropertiesService.getScriptProperties().setProperty('PIN', np);
  return { ok: true };
}

function stats_() {
  const st = stock_(), so = rows_('sold'), sa = rows_('sales'), tz = Session.getScriptTimeZone();
  const t0 = new Date(); t0.setHours(0, 0, 0, 0);
  const m0 = new Date(t0.getFullYear(), t0.getMonth(), 1);
  const last7 = [];
  for (let i = 6; i >= 0; i--) {
    const d = new Date(t0); d.setDate(d.getDate() - i);
    last7.push({ k: d.getTime(), l: Utilities.formatDate(d, tz, 'EEE'), n: 0 });
  }
  let today = 0, month = 0, eT = 0, eM = 0, eA = 0;
  so.forEach(r => {
    const d = new Date(r[2]);
    if (d >= m0) { month++; if (d >= t0) today++; }
    const day = new Date(d); day.setHours(0, 0, 0, 0);
    const x = last7.find(y => y.k === day.getTime()); if (x) x.n++;
  });
  sa.forEach(r => {
    if (r[8] === 'Undone') return;
    const d = new Date(r[1]), a = Number(r[5]) || 0;
    eA += a; if (d >= m0) { eM += a; if (d >= t0) eT += a; }
  });
  return { available: st.length, sold: so.length, total: st.length + so.length, today, month,
    earnToday: eT, earnMonth: eM, earnAll: eA, last7, low: st.length < LOW_STOCK, lowAt: LOW_STOCK };
}

// Shared by sell_ and sellOp_. Write order matters: Sold first, Sales second, Stock last.
// stock_() already hides sold IDs, so a crash part-way never loses an ID.
function commitSale_(picked, rest, customer, when, rate, fresh, opId) {
  const uids = picked.map(r => String(r[0]).trim()), joined = uids.join('\n');
  if (joined.length > 45000) throw new Error('Order too large. Split it into smaller orders.');
  const order = nextOrder_();
  push_('sold', picked.map(r => [String(r[0]).trim(), customer, when, order, String(r[1] || ''), String(r[3] || '')]));
  push_('sales', [[order, when, customer, uids.length, rate, rate * uids.length, joined, fresh ? 'Fresh' : 'Normal', 'Active', opId || '']]);
  rewrite_('stock', rest);
  return { order, uids };
}

function sell_(q) {
  const customer = String(q.customer || '').trim(), qty = Number(q.qty), rate = Number(q.rate) || 0;
  const fresh = q.mode === 'fresh';
  if (!customer) return { ok: false, error: 'Enter customer name' };
  if (!qty || qty < 1 || qty % 1) return { ok: false, error: 'Enter a valid quantity' };
  if (qty > MAX_QTY) return { ok: false, error: 'Maximum ' + MAX_QTY + ' IDs per order' };
  const st = stock_(), pool = st.filter(r => catMatch_(r, q.ty, q.ge));
  if (pool.length < qty) return { ok: false, error: 'Only ' + pool.length + ' IDs left' + ((q.ty || q.ge) ? ' in this category' : '') + ', ' + qty + ' requested' };
  // Normal = oldest first, Fresh = most recently added (bottom of the list)
  const picked = fresh ? pool.slice(pool.length - qty) : pool.slice(0, qty);
  const take = new Set(picked.map(r => String(r[0]).trim())), rest = st.filter(r => !take.has(String(r[0]).trim()));
  const r = commitSale_(picked, rest, customer, new Date(), rate, fresh, '');
  SpreadsheetApp.flush();
  lowCheck_(st.length, rest.length);
  const pw = {}; picked.forEach(x => { if (x[3]) pw[String(x[0]).trim()] = String(x[3]); });
  return { ok: true, uids: r.uids, pw: pw, order: r.order, amount: rate * qty, stats: stats_() };
}

// Each line: "UID|password" (separator: | , ; : space or tab). With hasPw false only UIDs are read (old behaviour).
function parseEntries_(text, hasPw) {
  const m = new Map();
  if (hasPw) {
    String(text || '').split(/\r?\n/).forEach(l => {
      l = l.trim(); if (!l) return;
      const x = l.match(/^(\S+?)(?:[\s|,;:]+(.+))?$/);
      if (x && !m.has(x[1])) m.set(x[1], (x[2] || '').trim());
    });
  } else {
    String(text || '').split(/[\r\n,;\s]+/).map(s => s.trim()).filter(Boolean).forEach(u => { if (!m.has(u)) m.set(u, ''); });
  }
  return Array.from(m.entries());
}
function add_(text, cat, hasPw, front, allowBad) { return addEntries_(parseEntries_(text, hasPw), cat, front, allowBad); }
// front = true puts the new IDs at the top of the stock (used first); otherwise at the end (Fresh).
// IDs already in stock are never added twice (their category / password can be updated). Sold IDs are skipped (use Returns).
// IDs marked Bad are skipped unless allowBad is true.
function addEntries_(list, cat, front, allowBad) {
  const code = (CAT_RE.test(String(cat)) && cat !== 'XX') ? String(cat) : '';
  if (!list.length) return { ok: false, error: 'No UIDs found' };
  const raw = rows_('stock'), inStock = new Map(), soldSet = new Set(), badSet = new Set();
  raw.forEach((r, i) => { const u = String(r[0]).trim(); if (u && !inStock.has(u)) inStock.set(u, i); });
  rows_('sold').forEach(r => soldSet.add(String(r[0]).trim()));
  rows_('bad').forEach(r => badSet.add(String(r[0]).trim()));
  const fresh = [], skip = { stock: 0, sold: 0, bad: 0 }; let moved = 0;
  list.forEach(([u, p]) => {
    if (soldSet.has(u)) { skip.sold++; return; }
    if (inStock.has(u)) {                          // already in stock: only update category / password
      const r = raw[inStock.get(u)]; let ch = false;
      if (code && String(r[1] || '').trim() !== code) { r[1] = code; ch = true; }
      if (p && String(r[3] || '') !== p) { r[3] = p; ch = true; }
      if (ch) moved++; else skip.stock++;
      return;
    }
    if (badSet.has(u) && !allowBad) { skip.bad++; return; }
    fresh.push([u, code, new Date(), p]);
  });
  if (front && fresh.length) rewrite_('stock', fresh.concat(raw));
  else { if (moved) rewrite_('stock', raw); push_('stock', fresh); }
  return { ok: true, added: fresh.length, moved: moved, duplicates: skip.stock + skip.sold + skip.bad, skip: skip, stats: stats_() };
}

function undo_(order) {
  const sales = rows_('sales'), i = sales.findIndex(r => String(r[0]) === String(order));
  if (i < 0 || sales[i][8] === 'Undone') return { ok: false, error: 'Order not found or already undone' };
  const back = [], keep = [];
  rows_('sold').forEach(r => String(r[3]) === String(order) ? back.push([String(r[0]), String(r[4] || ''), new Date(), String(r[5] || '')]) : keep.push(r));
  // Stock first, then Sold: a crash in between can only create a harmless duplicate, never a lost ID.
  // Fresh orders return to the end (still fresh), normal orders return to the top (still old).
  const cur = stock_();
  if (sales[i][7] === 'Fresh') rewrite_('stock', cur.concat(back)); else rewrite_('stock', back.concat(cur));
  rewrite_('sold', keep);
  sh_('sales').getRange(i + 2, 9).setValue('Undone');
  return { ok: true, stats: stats_() };
}

function replace_(uid) {
  uid = String(uid || '').trim();
  const so = rows_('sold'), i = so.findIndex(r => String(r[0]) === uid);
  if (i < 0) return { ok: false, error: 'That UID is not in Sold' };
  const st = stock_();
  if (!st.length) return { ok: false, error: 'No stock left' };
  const old = so[i], sales = rows_('sales'), k = sales.findIndex(r => String(r[0]) === String(old[3]));
  const fresh = k >= 0 && sales[k][7] === 'Fresh';          // Fresh order gets a fresh replacement
  const code = String(old[4] || '').trim() || 'XX';          // same category first, else any
  const same = st.filter(r => catOf_(r) === code), pool = same.length ? same : st;
  const pick = fresh ? pool[pool.length - 1] : pool[0], nu = String(pick[0]).trim();
  push_('bad', [[uid, old[1], old[3], new Date()]]);
  so[i] = [nu, old[1], old[2], old[3], String(pick[1] || ''), String(pick[3] || '')];
  rewrite_('sold', so);
  if (k >= 0) sh_('sales').getRange(k + 2, 7).setValue(String(sales[k][6]).split('\n').map(x => x === uid ? nu : x).join('\n'));
  rewrite_('stock', st.filter(r => r !== pick));
  return { ok: true, newUid: nu, newPass: String(pick[3] || ''), customer: old[1], order: old[3], stats: stats_() };
}

// ---------- RETURNS & REPLACEMENTS ----------
const parseUids_ = a => Array.from(new Set((Array.isArray(a) ? a : String(a || '').split(/[\r\n,;\s]+/)).map(x => String(x).trim()).filter(Boolean)));

// A customer gives IDs back: they leave Sold and the customer's order, and go back into stock.
// Normal orders go back to the top (still old), Fresh orders go to the end (still fresh).
function returnMany_(arr) {
  const list = parseUids_(arr);
  if (!list.length) return { ok: false, error: 'No UIDs found' };
  const so = rows_('sold'), sales = rows_('sales'), idx = new Map(), sIdx = new Map();
  so.forEach((r, i) => { const u = String(r[0]).trim(); if (u && !idx.has(u)) idx.set(u, i); });
  sales.forEach((r, i) => sIdx.set(String(r[0]), i));
  const gone = new Set(), top = [], end = [], log = [], done = [], missing = [], perSale = new Map(), now = new Date();
  list.forEach(u => {
    if (!idx.has(u)) { missing.push(u); return; }
    const i = idx.get(u), r = so[i], k = sIdx.get(String(r[3])), fresh = k !== undefined && sales[k][7] === 'Fresh';
    (fresh ? end : top).push([u, String(r[4] || ''), now, String(r[5] || '')]);
    log.push([u, r[1], r[3], now]);
    done.push({ uid: u, customer: r[1], order: r[3] });
    gone.add(i);
    if (k !== undefined) { if (!perSale.has(k)) perSale.set(k, new Set()); perSale.get(k).add(u); }
  });
  if (!done.length) return { ok: true, done, missing, stats: stats_() };
  // Stock first, then Sold: a crash in between can only leave a harmless duplicate, never a lost ID.
  rewrite_('stock', top.concat(stock_(), end));
  rewrite_('sold', so.filter((r, i) => !gone.has(i)));
  push_('returns', log);
  const sh = sh_('sales');
  perSale.forEach((set, k) => {
    const left = String(sales[k][6]).split('\n').filter(x => x && !set.has(x.trim())), rate = Number(sales[k][4]) || 0;
    sh.getRange(k + 2, 4).setValue(left.length);
    sh.getRange(k + 2, 6).setValue(rate * left.length);
    sh.getRange(k + 2, 7).setValue(left.join('\n'));
    if (!left.length) sh.getRange(k + 2, 9).setValue('Undone');   // whole order came back
  });
  return { ok: true, done, missing, stats: stats_() };
}

// IDs that do not log in: each is moved to Bad and the customer gets a new ID
// (same category first, else any; Fresh orders get the freshest ID).
function replaceMany_(arr) {
  const list = parseUids_(arr);
  if (!list.length) return { ok: false, error: 'No UIDs found' };
  const so = rows_('sold'), sales = rows_('sales'), idx = new Map(), sIdx = new Map(), pool = stock_();
  so.forEach((r, i) => { const u = String(r[0]).trim(); if (u && !idx.has(u)) idx.set(u, i); });
  sales.forEach((r, i) => sIdx.set(String(r[0]), i));
  const done = [], missing = [], noStock = [], badRows = [], swaps = new Map(), now = new Date();
  list.forEach(u => {
    if (!idx.has(u)) { missing.push(u); return; }
    if (!pool.length) { noStock.push(u); return; }
    const i = idx.get(u), old = so[i], k = sIdx.get(String(old[3])), fresh = k !== undefined && sales[k][7] === 'Fresh';
    const code = String(old[4] || '').trim() || 'XX';
    const same = pool.filter(r => catOf_(r) === code), from = same.length ? same : pool;
    const pick = fresh ? from[from.length - 1] : from[0], nu = String(pick[0]).trim();
    pool.splice(pool.indexOf(pick), 1);
    badRows.push([u, old[1], old[3], now]);
    so[i] = [nu, old[1], old[2], old[3], String(pick[1] || ''), String(pick[3] || '')];
    if (k !== undefined) { if (!swaps.has(k)) swaps.set(k, new Map()); swaps.get(k).set(u, nu); }
    done.push({ old: u, nu: nu, customer: old[1], order: old[3] });
  });
  if (!done.length) return { ok: true, done, missing, noStock, stats: stats_() };
  // Sold first (the new IDs are hidden from stock right away), then Sales, then Stock last.
  push_('bad', badRows);
  rewrite_('sold', so);
  const sh = sh_('sales');
  swaps.forEach((m, k) => sh.getRange(k + 2, 7).setValue(String(sales[k][6]).split('\n').map(x => m.has(x.trim()) ? m.get(x.trim()) : x).join('\n')));
  rewrite_('stock', pool);
  return { ok: true, done, missing, noStock, stats: stats_() };
}

// Save one password on the given IDs, wherever they are now (Stock column D and/or Sold column F).
function setPass_(arr, pass) {
  const list = parseUids_(arr), p = String(pass == null ? '' : pass).trim();
  if (!list.length) return { ok: false, error: 'No UIDs found' };
  if (!p) return { ok: false, error: 'Enter a password' };
  const set = new Set(list), hit = new Set();
  const st = rows_('stock'); let c1 = false;
  st.forEach(r => { const u = String(r[0]).trim(); if (set.has(u)) { r[3] = p; hit.add(u); c1 = true; } });
  const so = rows_('sold'); let c2 = false;
  so.forEach(r => { const u = String(r[0]).trim(); if (set.has(u)) { r[5] = p; hit.add(u); c2 = true; } });
  if (c2) rewrite_('sold', so);
  if (c1) rewrite_('stock', st);
  return { ok: true, saved: hit.size, missing: list.filter(u => !hit.has(u)), stats: stats_() };
}

function find_(uid) {
  uid = String(uid || '').trim();
  if (!uid) return { ok: false, error: 'Enter a UID' };
  const tz = Session.getScriptTimeZone(), fmt = d => d ? Utilities.formatDate(new Date(d), tz, 'dd MMM yyyy, hh:mm a') : '';
  let r = rows_('sold').find(x => String(x[0]).trim() === uid);
  if (r) return { ok: true, where: 'sold', customer: r[1], date: fmt(r[2]), order: r[3], pass: String(r[5] || '') };
  r = stock_().find(x => String(x[0]).trim() === uid);
  if (r) return { ok: true, where: 'stock', date: fmt(r[2]), pass: String(r[3] || '') };
  r = rows_('bad').find(x => String(x[0]).trim() === uid);
  if (r) return { ok: true, where: 'bad', customer: r[1], order: r[2], date: fmt(r[3]) };
  return { ok: true, where: 'none' };
}

function customers_() {
  const tz = Session.getScriptTimeZone(), m = {};
  rows_('sales').forEach(r => {
    if (r[8] === 'Undone') return;
    const c = m[r[2]] = m[r[2]] || { name: r[2], orders: 0, qty: 0, amount: 0, last: 0 };
    c.orders++; c.qty += Number(r[3]) || 0; c.amount += Number(r[5]) || 0; c.last = Math.max(c.last, new Date(r[1]).getTime());
  });
  return Object.keys(m).map(k => m[k]).sort((a, b) => b.amount - a.amount || b.qty - a.qty)
    .map(c => (c.last = Utilities.formatDate(new Date(c.last), tz, 'dd MMM yyyy'), c));
}

function history_(q) {
  const tz = Session.getScriptTimeZone();
  const from = q.from ? new Date(q.from + 'T00:00:00') : null, to = q.to ? new Date(q.to + 'T23:59:59') : null;
  const cu = String(q.customer || '').toLowerCase();
  return rows_('sales').reverse().filter(r => {
    const d = new Date(r[1]);
    return (!from || d >= from) && (!to || d <= to) && (!cu || String(r[2]).toLowerCase().indexOf(cu) >= 0);
  }).slice(0, 200).map(r => ({ order: r[0], date: Utilities.formatDate(new Date(r[1]), tz, 'dd MMM yyyy, hh:mm a'),
    customer: r[2], qty: r[3], rate: r[4], amount: r[5], uids: String(r[6]).split('\n'), mode: r[7], status: r[8] }));
}

// ---------- OFFLINE SYNC ----------
// The app queues sales/stock-adds while offline and sends them here when internet is back.
// Each operation has a unique id (saved in the Ops tab AND on the Sales row), so a repeat never applies twice.
function sync_(ops) {
  const done = {};
  rows_('ops').forEach(r => done[String(r[0])] = r[1]);
  rows_('sales').forEach(r => { if (r[9]) done[String(r[9])] = r[0]; });
  const res = []; let before = null;
  (ops || []).forEach(op => {
    const id = String(op.id);
    if (id in done) { res.push({ id, ok: true, order: done[id], dup: true }); return; }
    let r;
    try {
      if (op.t === 'add') r = addEntries_((op.uids || []).map(u => [u, (op.pws && op.pws[u]) || '']), op.cat, op.front, op.bad);
      else if (op.t === 'sell') { if (before === null) before = stock_().length; r = sellOp_(op); }
      else r = { ok: false, error: 'Unknown operation' };
    } catch (e) { r = { ok: false, error: String(e.message || e) }; }
    delete r.stats;
    res.push(Object.assign({ id }, r));
    if (r.ok) { push_('ops', [[id, r.order || 0]]); done[id] = r.order || 0; }   // log right away
  });
  SpreadsheetApp.flush();
  const snap = snap_();
  if (before !== null) lowCheck_(before, snap.stock.length);
  return { ok: true, res, snap };
}

// Sell the exact IDs the phone handed out. If one was taken meanwhile (e.g. another device), swap it for a free one
// (from the same category the phone asked for).
function sellOp_(op) {
  const fresh = op.mode === 'Fresh', st = stock_(), want = op.uids || [], qty = want.length;
  if (!op.customer || !qty) return { ok: false, error: 'Bad sale' };
  if (qty > MAX_QTY) return { ok: false, error: 'Maximum ' + MAX_QTY + ' IDs per order' };
  const key = r => String(r[0]).trim(), by = {};
  st.forEach(r => by[key(r)] = r);
  let picked = want.filter(u => by[u]).map(u => by[u]);
  const used = new Set(picked.map(key));
  const extra = (fresh ? st.slice().reverse() : st).filter(r => !used.has(key(r)) && catMatch_(r, op.ty, op.ge)).slice(0, qty - picked.length);
  if (picked.length + extra.length < qty) return { ok: false, error: 'Not enough stock for ' + op.customer };
  picked = picked.concat(extra);
  const take = new Set(picked.map(key)), rest = st.filter(r => !take.has(key(r)));
  const when = new Date(op.date || Date.now()), rate = Number(op.rate) || 0;
  const r = commitSale_(picked, rest, String(op.customer).trim(), when, rate, fresh, String(op.id));
  return { ok: true, order: r.order, uids: r.uids };
}

function snap_() {
  const st = stock_(), tags = {}, pw = {};
  st.forEach(r => { const c = catOf_(r); if (c !== 'XX') tags[String(r[0]).trim()] = c; });
  st.forEach(r => { if (r[3]) pw[String(r[0]).trim()] = String(r[3]); });
  rows_('sold').forEach(r => { if (r[5]) pw[String(r[0]).trim()] = String(r[5]); });
  return {
    stock: st.map(r => String(r[0]).trim()),
    ver: VERSION,
    tags: tags,
    pw: pw,
    sales: rows_('sales').map(r => ({ order: r[0], date: new Date(r[1]).getTime(), customer: r[2], qty: r[3], rate: r[4],
      amount: r[5], uids: String(r[6]).split('\n'), mode: r[7], status: r[8] })),
    bad: rows_('bad').map(r => [String(r[0]), r[1], r[2]])
  };
}

// ---------- PAGE (HTML) ----------
const HTML = String.raw`<!DOCTYPE html>
<html lang="en">
<head>
  <base target="_top">
  <meta name="theme-color" content="#F3F5FA" media="(prefers-color-scheme: light)">
  <meta name="theme-color" content="#080D18" media="(prefers-color-scheme: dark)">
  <meta name="color-scheme" content="light dark">
  <meta charset="utf-8">
  <title>FB IDs Vault</title>
  <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <style>
    :root{--bg:#F1F4F9;--surface:#fff;--ink:#0F172A;--muted:#64748B;--line:#E2E8F0;--brand:#2447F5;--brand-d:#1A35C4;--brand-soft:#E8EDFF;--ok:#0E9F6E;--ok-soft:#DDF5EC;--warn:#C2410C;--warn-soft:#FFE9DC;--bad:#D92D4A;--side:#0B1222;--r:16px;--r-sm:10px}
    @media(prefers-color-scheme:dark){:root{--bg:#0A101C;--surface:#121A2B;--ink:#E8EDF7;--muted:#8C98B2;--line:#1F2B44;--brand:#5B78FF;--brand-d:#7B93FF;--brand-soft:#18234A;--ok-soft:#0F2E24;--warn-soft:#3A1F10;--side:#070C17}}
    *{box-sizing:border-box;margin:0}
    html,body{background:var(--bg);color:var(--ink);font:15px/1.5 'Plus Jakarta Sans',system-ui,-apple-system,'Segoe UI',Roboto,sans-serif;-webkit-text-size-adjust:100%}
    button,input,textarea{font:inherit;color:inherit}
    :focus-visible{outline:3px solid var(--brand);outline-offset:2px}
    .hidden{display:none!important}
    .ic{width:20px;height:20px;fill:none;stroke:currentColor;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;flex:none}
    .logo{width:40px;height:40px;border-radius:12px;background:linear-gradient(135deg,#4F6BFF,#2447F5);display:grid;place-items:center;color:#fff;flex:none}

    /* Login */
    #login{min-height:100vh;display:grid;place-items:center;padding:20px}
    .login-card{width:100%;max-width:380px;background:var(--surface);border:1px solid var(--line);border-radius:22px;padding:32px 26px;box-shadow:0 20px 50px -20px rgba(17,27,43,.25)}
    .login-card .logo{width:52px;height:52px;margin-bottom:18px}
    .login-card h1{font-size:24px;font-weight:800;letter-spacing:-.02em}
    .login-card p{color:var(--muted);margin:4px 0 22px}

    /* Shell */
    .shell{display:flex;min-height:100vh}
    .side{width:240px;background:var(--side);color:#CBD5E8;padding:20px 14px;display:flex;flex-direction:column;gap:6px;position:sticky;top:0;height:100vh;flex:none}
    .brand{display:flex;align-items:center;gap:10px;color:#fff;font-weight:800;font-size:18px;padding:0 8px 18px}
    .nav{display:flex;flex-direction:column;gap:4px;flex:1}
    .nav button{display:flex;align-items:center;gap:12px;border:0;background:transparent;color:inherit;padding:11px 12px;border-radius:12px;font-weight:600;cursor:pointer;text-align:left}
    .nav button:hover{background:rgba(255,255,255,.07)}
    .nav button.on{background:var(--brand);color:#fff}
    .nav button[data-t=sec]{margin-top:auto}
    .sfoot{background:rgba(255,255,255,.06);border-radius:14px;padding:14px;font-size:12.5px}
    .sfoot b{display:block;color:#fff;font-size:22px;font-weight:800}
    .sfoot .meter{background:rgba(255,255,255,.12);margin-top:8px}
    .content{flex:1;min-width:0}
    .top{display:flex;align-items:center;gap:12px;padding:18px 28px;position:sticky;top:0;z-index:5;background:color-mix(in srgb,var(--bg) 88%,transparent);backdrop-filter:blur(10px);border-bottom:1px solid var(--line)}
    .top>div{flex:1}
    .top>.mlogo{flex:none;width:34px;height:34px;border-radius:10px}
    .top>.mlogo .ic{width:18px;height:18px}
    .top h1{font-size:22px;font-weight:800;letter-spacing:-.02em;line-height:1.2}
    .top p{color:var(--muted);font-size:13.5px}
    .pill{font-size:12.5px;font-weight:700;padding:7px 13px;border-radius:99px;background:var(--ok-soft);color:var(--ok);white-space:nowrap}
    .pill.low{background:var(--warn-soft);color:var(--warn)}
    .icbtn{width:42px;height:42px;border-radius:12px;border:1px solid var(--line);background:var(--surface);display:grid;place-items:center;cursor:pointer;color:var(--muted)}
    .icbtn:hover{color:var(--brand);border-color:var(--brand)}
    main{padding:24px 28px;max-width:1240px;display:grid;gap:20px}
    .g2{display:grid;grid-template-columns:1.5fr 1fr;gap:20px;align-items:start}
    .stack{display:grid;gap:20px}
    @media(max-width:860px){
      .shell{display:block}
      .side{position:fixed;top:auto;bottom:0;left:0;right:0;width:auto;height:auto;flex-direction:row;padding:8px 8px calc(8px + env(safe-area-inset-bottom));z-index:10}
      .brand,.sfoot{display:none}
      .nav{flex-direction:row;justify-content:space-around}
      .nav button[data-t=sec]{margin:0}
      .nav button{padding:11px 14px}
      .nav button span{display:none}.nav button.on span{display:block;font-size:13px}
      .top{padding:14px 16px}.top h1{font-size:19px}
      main{padding:16px 16px 96px}.g2{grid-template-columns:1fr}
    }

    /* KPIs */
    .kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}
    @media(max-width:860px){.kpis{grid-template-columns:repeat(2,1fr)}}
    .kpi{background:var(--surface);border:1px solid var(--line);border-radius:var(--r);padding:16px;display:flex;gap:12px;align-items:center;min-width:0}
    .kic{width:42px;height:42px;border-radius:12px;display:grid;place-items:center;background:var(--brand-soft);color:var(--brand);flex:none}
    .kpi.ok .kic{background:var(--ok-soft);color:var(--ok)}
    .kpi span{font-size:12.5px;color:var(--muted);font-weight:600;display:block}
    .kpi b{font-size:24px;font-weight:800;letter-spacing:-.02em;line-height:1.2}
    .kpi.main{background:linear-gradient(135deg,#3A5BFF,#1A35C4);border:0;color:#fff}
    .kpi.main .kic{background:rgba(255,255,255,.18);color:#fff}
    .kpi.main span{color:#fff;opacity:.85}
    .kpi.main.low{background:linear-gradient(135deg,#EA6A1F,#C2410C)}

    /* Cards + forms */
    .card{background:var(--surface);border:1px solid var(--line);border-radius:var(--r);padding:22px}
    .card h3{font-size:17px;font-weight:700;letter-spacing:-.01em}
    .card .sub{color:var(--muted);font-size:13.5px;margin:2px 0 18px}
    label{display:block;font-size:13px;font-weight:600;margin:0 0 6px}
    .field{margin-bottom:14px}
    input,textarea,select{width:100%;background:var(--bg);border:1.5px solid var(--line);border-radius:var(--r-sm);padding:13px 14px;font-size:16px;transition:border-color .15s,background .15s}
    input:focus,textarea:focus{outline:none;border-color:var(--brand);background:var(--surface)}
    textarea{min-height:200px;resize:vertical;font-family:ui-monospace,Consolas,monospace;font-size:14px}
    .row3{display:grid;grid-template-columns:1fr;gap:14px}
    @media(min-width:520px){.row3{grid-template-columns:1.6fr 1fr 1fr}}
    .row{display:grid;grid-template-columns:1fr;gap:14px}
    @media(min-width:520px){.row{grid-template-columns:1.6fr 1fr}}
    .pw{position:relative}.pw input{padding-right:50px}
    .eye{position:absolute;right:6px;top:50%;transform:translateY(-50%);border:0;background:transparent;width:38px;height:38px;display:grid;place-items:center;color:var(--muted);cursor:pointer;border-radius:8px}
    .eye:hover{color:var(--brand)}
    .seg{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:14px}
    .seg button{display:grid;grid-template-columns:auto 1fr;column-gap:10px;align-items:center;text-align:left;padding:12px 14px;border:1.5px solid var(--line);background:var(--bg);border-radius:var(--r-sm);cursor:pointer;color:var(--muted)}
    .seg .ic{grid-row:span 2}.seg small{font-size:12px}.seg b{color:var(--ink)}
    .seg button.on{border-color:var(--brand);background:var(--brand-soft);color:var(--brand)}
    .chips{display:flex;gap:8px;flex-wrap:wrap;margin:-4px 0 16px;align-items:center}
    .chip{border:1.5px solid var(--line);background:transparent;border-radius:99px;padding:6px 14px;font-weight:600;font-size:13.5px;cursor:pointer;color:var(--muted)}
    .chip:hover{border-color:var(--brand);color:var(--brand)}
    .tot{margin-left:auto;font-weight:700;color:var(--ok)}
    .btn{display:inline-flex;align-items:center;justify-content:center;gap:8px;border:0;border-radius:var(--r-sm);padding:14px 20px;font-weight:700;cursor:pointer;background:var(--brand);color:#fff;width:100%;transition:background .15s,transform .05s}
    .btn:hover{background:var(--brand-d)}.btn:active{transform:scale(.98)}.btn:disabled{opacity:.55;cursor:wait}
    .btn.ghost{background:var(--brand-soft);color:var(--brand)}.btn.ghost:hover{background:var(--line)}
    .btn.sm{width:auto;padding:9px 14px;font-size:13.5px}.btn.sm .ic{width:16px;height:16px}
    .err{color:var(--bad);font-size:13.5px;font-weight:600;margin-top:10px;min-height:18px}

    /* Meter + chart */
    .meter{height:10px;border-radius:99px;background:var(--bg);overflow:hidden}
    .meter i{display:block;height:100%;width:0;background:var(--ok);border-radius:99px;transition:width .5s}
    .meter.low i{background:var(--warn)}
    .bars{display:flex;align-items:flex-end;gap:10px;height:170px}
    .bar{flex:1;display:flex;flex-direction:column;justify-content:flex-end;align-items:center;height:100%;gap:4px;font-size:12px;color:var(--muted)}
    .bar i{width:100%;max-width:34px;background:linear-gradient(var(--brand),var(--brand-d));border-radius:8px 8px 4px 4px;min-height:4px}
    .bar em{font-style:normal;font-weight:700;color:var(--ink)}

    /* Results + lists */
    .result{border:1.5px solid var(--ok);background:var(--ok-soft);border-radius:var(--r);padding:18px;margin-top:18px}
    .result-head{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:12px}
    .acts{display:flex;gap:8px;flex-wrap:wrap}
    .uids{background:var(--surface);border:1px solid var(--line);border-radius:var(--r-sm);padding:12px 14px;max-height:260px;overflow:auto;font-family:ui-monospace,Consolas,monospace;font-size:13.5px;white-space:pre;-webkit-user-select:all;user-select:all}
    .filters{display:grid;grid-template-columns:1fr;gap:10px;margin-bottom:16px}
    @media(min-width:700px){.filters{grid-template-columns:1fr 1fr 1.4fr auto auto}}
    .hist{display:grid;gap:10px}
    .h-item{border:1px solid var(--line);border-radius:var(--r-sm);padding:12px 14px;display:flex;align-items:center;gap:12px;flex-wrap:wrap}
    .h-item.undone{opacity:.5}
    .h-item .who{flex:1;min-width:140px}.h-item .who b{display:block}.h-item .who span{font-size:12.5px;color:var(--muted)}
    .qty{font-weight:800;background:var(--brand-soft);color:var(--brand);border-radius:8px;padding:4px 10px;font-size:14px}
    .tag{font-size:11.5px;font-weight:700;padding:3px 8px;border-radius:99px;background:var(--warn-soft);color:var(--warn)}
    .found{border:1.5px solid var(--line);border-radius:var(--r);padding:16px;margin-top:16px}
    .empty{color:var(--muted);text-align:center;padding:22px 0}
    .toast{position:fixed;left:50%;bottom:calc(90px + env(safe-area-inset-bottom));transform:translate(-50%,30px);background:var(--ink);color:var(--bg);padding:11px 18px;border-radius:99px;font-weight:600;font-size:14px;opacity:0;pointer-events:none;transition:.25s;z-index:20}
    .toast.show{opacity:1;transform:translate(-50%,0)}
    @media(min-width:861px){.toast{bottom:28px}}
    @media(prefers-reduced-motion:reduce){*{transition:none!important}}

    /* ===== PRO THEME (overrides) ===== */
    :root{--bg:#F3F5FA;--surface:#fff;--ink:#0B1220;--muted:#667085;--line:#E6EAF2;--brand:#2F54F5;--brand-d:#1E3BD0;--brand-soft:#EAEFFF;--side:#0A1020;--r:20px;--r-sm:14px;--sh:0 1px 2px rgba(16,24,40,.04),0 8px 24px -12px rgba(16,24,40,.12)}
    @media(prefers-color-scheme:dark){:root{--bg:#080D18;--surface:#101828;--ink:#EEF2FA;--muted:#93A0BA;--line:#1C2740;--brand:#6580FF;--brand-d:#86A0FF;--brand-soft:#16224A;--side:#0A1020;--sh:0 1px 0 rgba(255,255,255,.03) inset,0 10px 30px -14px rgba(0,0,0,.6)}}
    html,body{background:var(--bg);letter-spacing:-.005em}
    body{-webkit-tap-highlight-color:transparent}
    .card,.kpi{box-shadow:var(--sh);border-color:var(--line)}
    .card{padding:24px}
    .card h3{font-size:18px}
    .top{padding:16px 28px;background:color-mix(in srgb,var(--bg) 82%,transparent);backdrop-filter:blur(16px) saturate(1.4)}
    .top h1{font-size:24px}
    .pill{border:1px solid color-mix(in srgb,currentColor 20%,transparent);display:inline-flex;align-items:center;gap:7px}
    .pill::before{content:"";width:8px;height:8px;border-radius:50%;background:currentColor}
    .icbtn{box-shadow:var(--sh)}
    .mlogo{display:none;width:38px;height:38px;border-radius:12px}
    /* sidebar */
    .side{background:linear-gradient(180deg,#0C1328,#070B17);width:252px;padding:22px 14px}
    .brand{font-size:19px;letter-spacing:-.02em;padding-bottom:22px}
    .nav button{padding:12px 14px;border-radius:14px;color:#A9B6D3;transition:background .15s,color .15s}
    .nav button:hover{color:#fff}
    .nav button.on{background:linear-gradient(135deg,#4767FF,#2442D9);color:#fff;box-shadow:0 8px 20px -8px rgba(47,84,245,.8)}
    .sfoot{background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.07);border-radius:16px}
    /* KPIs */
    .kpis{gap:14px}
    .kpi{padding:18px;border-radius:var(--r);transition:transform .15s}
    .kpi:hover{transform:translateY(-2px)}
    .kpi b{font-size:26px}
    .kic{width:44px;height:44px;border-radius:14px}
    .kpi.main{position:relative;overflow:hidden;box-shadow:0 14px 30px -14px rgba(36,71,245,.7)}
    .kpi.main::after{content:"";position:absolute;right:-30px;top:-40px;width:130px;height:130px;border-radius:50%;background:rgba(255,255,255,.12)}
    .kpi.main b{font-size:32px}
    /* forms */
    input,textarea,select{background:color-mix(in srgb,var(--bg) 70%,var(--surface));border:1.5px solid var(--line);border-radius:var(--r-sm);padding:14px 15px}
    input:focus,textarea:focus{border-color:var(--brand);box-shadow:0 0 0 4px color-mix(in srgb,var(--brand) 18%,transparent)}
    label{font-size:12.5px;text-transform:uppercase;letter-spacing:.04em;color:var(--muted)}
    .btn{border-radius:var(--r-sm);padding:15px 20px;background:linear-gradient(135deg,#4767FF,#2442D9);box-shadow:0 10px 22px -10px rgba(47,84,245,.8)}
    .btn:hover{background:linear-gradient(135deg,#5876FF,#2E4DE6)}
    .btn.ghost{background:var(--brand-soft);color:var(--brand);box-shadow:none}
    .btn.sm{padding:10px 15px;border-radius:12px}
    .seg button{border-radius:var(--r-sm);padding:14px;transition:all .15s}
    .seg button.on{box-shadow:0 0 0 3px color-mix(in srgb,var(--brand) 16%,transparent)}
    .chip{background:var(--surface);padding:8px 16px}
    .chip:active{transform:scale(.96)}
    /* results + lists */
    .result{border-radius:var(--r);background:linear-gradient(180deg,var(--ok-soft),transparent 140%)}
    .uids{border-radius:var(--r-sm);line-height:1.7}
    .h-item{border-radius:var(--r-sm);padding:14px 16px;background:var(--surface);transition:border-color .15s}
    .h-item:hover{border-color:var(--brand)}
    .qty{border-radius:10px;padding:5px 11px}
    .bar i{background:linear-gradient(180deg,#6A85FF,#2442D9);border-radius:10px 10px 4px 4px}
    .meter{height:12px}
    .found{border-radius:var(--r);background:var(--bg)}
    /* login */
    #login{background:radial-gradient(900px 500px at 50% -10%,color-mix(in srgb,var(--brand) 28%,transparent),transparent 70%),var(--bg)}
    .login-card{border-radius:28px;padding:38px 28px;box-shadow:0 30px 70px -30px rgba(16,24,40,.45)}
    .login-card .logo{width:60px;height:60px;border-radius:18px;box-shadow:0 14px 28px -10px rgba(47,84,245,.8)}
    .login-card h1{font-size:28px}
    /* motion */
    [data-tab]:not(.hidden){animation:up .28s ease both}
    @keyframes up{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}
    /* mobile */
    @media(max-width:860px){
      .mlogo{display:grid}
      .top{padding:12px 16px calc(12px)}
      .top h1{font-size:20px}.top p{font-size:12.5px}
      .pill{font-size:11.5px;padding:6px 10px}
      .side{left:12px;right:12px;bottom:calc(10px + env(safe-area-inset-bottom));width:auto;border-radius:24px;padding:8px;background:rgba(10,16,32,.94);backdrop-filter:blur(16px);box-shadow:0 18px 40px -12px rgba(0,0,0,.55);border:1px solid rgba(255,255,255,.08)}
      .nav button{padding:11px 13px;border-radius:16px;gap:8px}
      .nav button.on span{font-size:13px;font-weight:700}
      main{padding:14px 14px 118px;gap:16px}
      .kpis{grid-template-columns:repeat(2,1fr);gap:10px}
      .kpi.main{grid-column:1/-1;padding:20px}
      .kpi{padding:14px;gap:10px}.kpi b{font-size:21px}.kic{width:38px;height:38px;border-radius:12px}
      .card{padding:18px;border-radius:22px}
      .toast{bottom:calc(104px + env(safe-area-inset-bottom))}
    }
    @media(prefers-reduced-motion:reduce){[data-tab]{animation:none!important}.kpi:hover{transform:none}}

    /* ===== COMPACT ===== */
    :root{--r:16px;--r-sm:11px}
    html,body{font-size:14px;line-height:1.45}
    main{padding:16px 20px;gap:14px;max-width:1180px}
    .g2,.stack{gap:14px}
    .top{padding:10px 20px}
    .top h1{font-size:20px}.top p{font-size:12px}
    .icbtn{width:36px;height:36px;border-radius:11px}
    .pill{font-size:11.5px;padding:5px 11px}
    .side{width:216px;padding:16px 10px;gap:4px}
    .brand{font-size:17px;padding:0 8px 14px}
    .logo{width:34px;height:34px;border-radius:10px}
    .nav{gap:2px}
    .nav button{padding:9px 12px;font-size:13.5px;border-radius:11px;gap:10px}
    .nav .ic{width:18px;height:18px}
    .sfoot{padding:11px 12px;font-size:11.5px}.sfoot b{font-size:19px}
    .kpis{gap:10px}
    .kpi{padding:11px 13px;gap:10px;border-radius:14px}
    .kpi:hover{transform:none}
    .kic{width:34px;height:34px;border-radius:10px}
    .kic .ic{width:17px;height:17px}
    .kpi span{font-size:11.5px}
    .kpi b{font-size:19px}
    .kpi.main b{font-size:24px}
    .kpi.main::after{width:90px;height:90px;right:-20px;top:-30px}
    .card{padding:16px;border-radius:16px}
    .card h3{font-size:15.5px}
    .card .sub{font-size:12.5px;margin:1px 0 12px}
    label{font-size:11px;margin-bottom:5px}
    .field{margin-bottom:10px}
    input,textarea,select{padding:10px 12px;font-size:15px}
    textarea{min-height:150px;font-size:13px}
    .row3,.row{gap:10px}
    .btn{padding:11px 16px;font-size:14px}
    .btn.sm{padding:7px 12px;font-size:12.5px}
    .btn.sm .ic{width:15px;height:15px}
    .seg{gap:8px;margin-bottom:10px}
    .seg button{padding:9px 12px}.seg small{font-size:11px}
    .chips{gap:6px;margin:-2px 0 12px}
    .chip{padding:5px 13px;font-size:12.5px}
    .result{padding:12px;margin-top:12px}
    .uids{padding:9px 12px;font-size:12.5px;line-height:1.55;max-height:220px}
    .hist{gap:8px}
    .h-item{padding:9px 12px;gap:8px}
    .h-item .who span{font-size:11.5px}
    .qty{font-size:12.5px;padding:3px 9px}
    .filters{gap:8px;margin-bottom:12px}
    .bars{height:130px;gap:8px}
    .bar{font-size:11px}
    .meter{height:9px}
    .found{padding:12px;margin-top:12px}
    .login-card{padding:28px 22px;max-width:340px}
    .login-card .logo{width:50px;height:50px;border-radius:15px;margin-bottom:14px}
    .login-card h1{font-size:24px}
    .login-card p{margin-bottom:16px}
    @media(max-width:860px){
      .mlogo{width:32px;height:32px;border-radius:10px}
      .top{padding:9px 14px}
      .top h1{font-size:18px}
      .side{left:10px;right:10px;bottom:calc(8px + env(safe-area-inset-bottom));padding:5px;border-radius:20px;width:auto}
      .nav button{padding:9px 11px;border-radius:14px}
      .nav button.on span{font-size:12.5px}
      main{padding:10px 12px 84px;gap:12px}
      .kpis{gap:8px}
      .kpi{padding:10px 11px}.kpi b{font-size:17px}.kic{width:30px;height:30px}
      .kpi.main{padding:13px 14px}.kpi.main b{font-size:24px}
      .card{padding:14px;border-radius:16px}
      .toast{bottom:calc(84px + env(safe-area-inset-bottom))}
    }

    /* ===== Menu drawer (replaces bottom bar) ===== */
    .top>.icbtn{display:none}
    .menubtn{border:0;cursor:pointer;padding:0}
    .menubtn:active{transform:scale(.94)}
    .side .logout{display:flex;align-items:center;gap:12px;width:100%;margin-top:10px;border:0;cursor:pointer;background:rgba(255,92,120,.12);color:#FF8DA0;padding:12px 14px;border-radius:14px;font-weight:700;font-size:14.5px;text-align:left}
    .side .logout:hover{background:rgba(255,92,120,.2)}
    .scrim{display:none}
    @media(max-width:860px){
      .scrim{display:block;position:fixed;inset:0;background:rgba(3,7,18,.6);backdrop-filter:blur(2px);z-index:25;opacity:0;pointer-events:none;transition:opacity .25s}
      .menu-open .scrim{opacity:1;pointer-events:auto}
      .shell .side{position:fixed;top:0;bottom:0;left:0;right:auto;width:288px;max-width:84vw;height:100%;flex-direction:column;gap:6px;
        padding:calc(18px + env(safe-area-inset-top)) 14px calc(16px + env(safe-area-inset-bottom));border-radius:0 26px 26px 0;overflow-y:auto;z-index:30;
        background:linear-gradient(180deg,#0C1328,#070B17);box-shadow:24px 0 60px -20px rgba(0,0,0,.7);border:0;
        transform:translateX(-104%);visibility:hidden;transition:transform .28s cubic-bezier(.2,.8,.2,1),visibility 0s .28s}
      .menu-open .side{transform:none;visibility:visible;transition:transform .28s cubic-bezier(.2,.8,.2,1),visibility 0s}
      .shell .brand{display:flex}
      .shell .sfoot{display:block}
      .shell .nav{flex-direction:column;justify-content:flex-start;gap:4px}
      .shell .nav button{padding:13px 14px;font-size:15px;gap:12px;border-radius:14px}
      .shell .nav button[data-t=sec]{margin:0}
      .shell .nav button span,.shell .nav button.on span{display:block;font-size:15px}
      main{padding-bottom:32px!important}
    }
    @media(prefers-reduced-motion:reduce){.shell .side,.scrim{transition:none!important}}

    /* ===== Compact stats ===== */
    .kpis2{display:grid;grid-template-columns:1fr;gap:10px}
    @media(min-width:861px){.kpis2{grid-template-columns:minmax(250px,.8fr) 1.8fr}}
    .hero{position:relative;overflow:hidden;border-radius:var(--r);padding:14px 16px;color:#fff;background:linear-gradient(135deg,#3A5BFF,#1A35C4);box-shadow:0 14px 30px -16px rgba(36,71,245,.7);display:grid;grid-template-columns:1fr auto;gap:8px 14px;align-items:end}
    .hero.low{background:linear-gradient(135deg,#EA6A1F,#C2410C)}
    .hero span{display:block;font-size:12px;font-weight:600;opacity:.85}
    .hero b{display:block;font-size:34px;font-weight:800;letter-spacing:-.03em;line-height:1.05}
    .hero .hr{text-align:right}.hero .hr b{font-size:18px;font-weight:700;letter-spacing:-.01em}
    .hbar{grid-column:1/-1;height:6px;border-radius:99px;background:rgba(255,255,255,.25);overflow:hidden}
    .hbar i{display:block;height:100%;width:0;background:#fff;border-radius:99px;transition:width .5s}
    .sc{background:var(--surface);border:1px solid var(--line);border-radius:var(--r);box-shadow:var(--sh);display:grid;align-content:center}
    .sr{display:grid;grid-template-columns:54px repeat(3,minmax(0,1fr));align-items:center;gap:6px;padding:10px 14px}
    .sr+.sr{border-top:1px solid var(--line)}
    .sl{font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:.07em}
    .sl.sold{color:var(--brand)}.sl.earn{color:var(--ok)}
    .sr .c b{display:block;font-size:19px;font-weight:800;letter-spacing:-.02em;line-height:1.2;overflow:hidden;text-overflow:ellipsis}
    .sr .c span{display:block;font-size:11.5px;color:var(--muted);font-weight:600}
    @media(max-width:380px){.sr{grid-template-columns:46px repeat(3,minmax(0,1fr));padding:9px 10px;gap:4px}.sr .c b{font-size:17px}.hero b{font-size:30px}}
    /* ===== Saved customer names ===== */
    .nchips{display:flex;gap:8px;overflow-x:auto;padding:0 0 4px;margin:-4px 0 14px;scrollbar-width:none}
    .nchips::-webkit-scrollbar{display:none}
    .nchips:empty{display:none}
    .nchip{flex:none;border:1px solid var(--line);background:var(--bg);color:var(--ink);border-radius:99px;padding:6px 13px;font-size:13px;font-weight:600;cursor:pointer;white-space:nowrap}
    .nchip:hover,.nchip:active{border-color:var(--brand);color:var(--brand)}
    .side .logout.rep{background:rgba(255,255,255,.07);color:#CBD5E8;margin-top:8px}.side .logout.rep:hover{background:rgba(255,255,255,.12)}
    .rfbtn{width:38px;height:38px;border-radius:12px;border:1px solid var(--line);background:var(--surface);color:var(--ink);display:grid;place-items:center;cursor:pointer;flex:none}
    .rfbtn:active{transform:scale(.94)}
    .rfbtn.spin .ic{animation:rfspin .8s linear infinite}
    @keyframes rfspin{to{transform:rotate(360deg)}}

    /* ===== Modals, order detail, sheet connections ===== */
    .toast{z-index:80}
    .modal{position:fixed;inset:0;z-index:60;background:rgba(6,10,20,.62);backdrop-filter:blur(3px);display:flex;align-items:center;justify-content:center;padding:16px}
    .sheetbox{width:100%;max-width:580px;max-height:92vh;overflow:auto;background:var(--surface);border:1px solid var(--line);border-radius:24px;padding:22px;box-shadow:0 30px 80px -20px rgba(0,0,0,.55);animation:up .22s ease both}
    @media(max-width:600px){.modal{align-items:flex-end;padding:0}.sheetbox{border-radius:26px 26px 0 0;max-height:94vh;padding:20px 18px calc(22px + env(safe-area-inset-bottom))}}
    .mhead{display:flex;align-items:flex-start;gap:10px;margin-bottom:14px}
    .mhead>div{flex:1;min-width:0}
    .mhead h3{font-size:19px;font-weight:800;letter-spacing:-.01em;word-break:break-word}
    .mhead .tag{flex:none;margin-top:4px}
    .mhead .icbtn{flex:none;width:38px;height:38px}
    .sheetbox .sub{color:var(--muted);font-size:13.5px;line-height:1.5}
    .mh2{font-size:15px;font-weight:700;margin:22px 0 10px}
    .kv{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px;margin-bottom:14px}
    .kv>div{background:var(--bg);border:1px solid var(--line);border-radius:12px;padding:9px 12px;min-width:0}
    .kv span{display:block;font-size:11.5px;color:var(--muted);font-weight:600}
    .kv b{font-size:14.5px;word-break:break-word}
    .found .kv>div{background:var(--surface)}
    .tag.ok{background:var(--ok-soft);color:var(--ok)}
    .fhead{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:12px}
    .fhead code{font-family:ui-monospace,Consolas,monospace;font-size:13.5px;color:var(--muted);word-break:break-all}
    .facts{display:flex;gap:8px;flex-wrap:wrap}
    .od-list{max-height:42vh}
    mark{background:#FDE68A;color:#1F2937;border-radius:4px;padding:1px 5px;font-weight:700}
    .clk{cursor:pointer}
    .h-item.conn.act{border-color:var(--ok)}
    .cmsg{font-size:13.5px;font-weight:600;margin:2px 0 10px;min-height:20px;color:var(--muted)}
    .cmsg.ok{color:var(--ok)}.cmsg.bad{color:var(--bad)}
    .linkbtn{display:flex;align-items:center;justify-content:center;gap:8px;width:100%;margin-top:14px;padding:10px;background:none;border:0;color:var(--muted);font-weight:600;font-size:13.5px;cursor:pointer;border-radius:12px}
    .linkbtn:hover{color:var(--brand)}
    .linkbtn b{color:var(--brand)}

    /* ===== Categories ===== */
    .cpill{border:1.5px solid var(--line);background:var(--bg);color:var(--ink);border-radius:99px;padding:8px 14px;font-weight:700;font-size:13.5px;cursor:pointer;display:inline-flex;gap:7px;align-items:center}
    .cpill small{font-weight:600;color:var(--muted);font-size:12.5px}
    .cpill.on{border-color:var(--brand);background:var(--brand-soft);color:var(--brand)}
    .cpill.on small{color:var(--brand)}
    .pills{display:flex;gap:8px;flex-wrap:wrap;margin:0 0 14px}
    .selinfo{font-size:13.5px;font-weight:600;color:var(--muted);margin:-2px 0 14px}
    .cgrid{width:100%;border-collapse:collapse;font-size:13.5px}
    .cgrid th,.cgrid td{padding:9px 6px;text-align:center;border-bottom:1px solid var(--line)}
    .cgrid th:first-child,.cgrid td:first-child{text-align:left;font-weight:700}
    .cgrid th{color:var(--muted);font-weight:600;font-size:12.5px}
    .cgrid td.z{color:var(--muted)}

    /* ===== Dropdowns (dark + light) ===== */
    :root{color-scheme:light}
    @media(prefers-color-scheme:dark){:root{color-scheme:dark}}
    select{font:inherit;color:var(--ink);-webkit-appearance:none;appearance:none;cursor:pointer;padding-right:44px;text-overflow:ellipsis;
      background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='18' height='18' viewBox='0 0 24 24' fill='none' stroke='%238C98B2' stroke-width='2.4' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='m6 9 6 6 6-6'/%3E%3C/svg%3E");
      background-repeat:no-repeat;background-position:right 14px center;background-size:18px}
    select:focus{outline:none;border-color:var(--brand);background-color:var(--surface)}
    option{background:var(--surface);color:var(--ink);padding:8px}
    .row2{display:grid;grid-template-columns:1fr 1fr;gap:12px}
    .selinfo{font-size:13.5px;font-weight:600;color:var(--muted);margin:-2px 0 14px}

    /* ===== Settings ===== */
    .setcard{padding:0;overflow:hidden}
    .sethead{display:flex;gap:14px;align-items:flex-start;padding:22px 22px 0}
    .seticon{width:46px;height:46px;border-radius:14px;display:grid;place-items:center;flex:none;background:linear-gradient(135deg,#4F6BFF,#2447F5);color:#fff;box-shadow:0 10px 22px -10px rgba(36,71,245,.7)}
    .seticon.sh{background:linear-gradient(135deg,#19B98A,#0E9F6E);box-shadow:0 10px 22px -10px rgba(14,159,110,.7)}
    .sethead h3{font-size:18px}
    .sethead .sub{margin:2px 0 0}
    .setbody{padding:18px 22px 22px}
    .connbox{display:flex;align-items:center;gap:12px;padding:14px 16px;border:1px solid var(--line);border-radius:14px;background:var(--bg);margin-bottom:14px}
    .connbox .dot{width:10px;height:10px;border-radius:50%;background:var(--ok);box-shadow:0 0 0 4px var(--ok-soft);flex:none}
    .cmeta{flex:1;min-width:0}
    .cmeta b{display:block;font-size:15.5px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .cmeta span{display:block;color:var(--muted);font-size:12.5px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .pinrow{display:grid;gap:0}
    .hint{display:flex;gap:8px;align-items:center;color:var(--muted);font-size:12.5px;margin:-4px 0 14px}
    .hint .ic{width:16px;height:16px}
    .pinrow input{letter-spacing:.2em}
    .pinrow input::placeholder{letter-spacing:0}
    @media(max-width:600px){.sethead{padding:18px 16px 0}.setbody{padding:16px}}

    /* ===== Success popup ===== */
    .modal.center{align-items:center;padding:18px}
    .modal.center .sheetbox{max-width:380px;border-radius:26px;text-align:center;padding:30px 24px 24px}
    .okring{width:78px;height:78px;border-radius:50%;margin:0 auto 18px;display:grid;place-items:center;background:var(--ok-soft);color:var(--ok);box-shadow:0 0 0 10px color-mix(in srgb,var(--ok-soft) 55%,transparent);animation:pop .45s cubic-bezier(.2,1.4,.4,1) both}
    .okring .ic{width:38px;height:38px;stroke-width:2.4}
    @keyframes pop{from{transform:scale(.3);opacity:0}to{transform:scale(1);opacity:1}}
    .okbox h3{font-size:21px;font-weight:800;margin-bottom:6px}
    .okbox .sub{margin:0 0 22px}

    /* ===== Small fixes ===== */
    .pinrow .field{margin-bottom:14px}
    .top>div{min-width:0}
    .top h1{white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    @media(max-width:600px){.top{gap:8px}.top p{display:none}.top h1{font-size:17px;white-space:normal;line-height:1.15}.top .pill{font-size:11.5px;padding:6px 10px}}
    /* ===== Add Stock (compact) ===== */
    .ahead{display:flex;justify-content:space-between;align-items:flex-start;gap:10px;margin-bottom:12px}
    .cbadge{font-size:12px;font-weight:700;padding:5px 11px;border-radius:99px;background:var(--brand-soft);color:var(--brand);white-space:nowrap}
    .arow{display:grid;grid-template-columns:78px minmax(0,1fr);gap:10px;align-items:center;margin-bottom:10px}
    .arow>label{margin:0}
    .sg{display:flex;gap:3px;background:var(--bg);border:1px solid var(--line);border-radius:12px;padding:3px}
    .sg button{flex:1;border:0;background:transparent;color:var(--muted);font-weight:700;font-size:12.5px;padding:8px 6px;border-radius:9px;cursor:pointer;white-space:nowrap;transition:background .15s,color .15s}
    .sg button.on{background:var(--surface);color:var(--brand);box-shadow:var(--sh)}
    .lbrow{display:flex;justify-content:space-between;align-items:center;margin:4px 0 6px}
    .lbrow label{margin:0}
    .tools{display:flex;gap:6px}
    .mini{border:1px solid var(--line);background:var(--bg);color:var(--ink);border-radius:9px;padding:4px 10px;font-size:12px;font-weight:700;cursor:pointer;display:inline-flex;gap:5px;align-items:center}
    .mini:hover{border-color:var(--brand);color:var(--brand)}
    .mini .ic{width:14px;height:14px}
    #addTxt{min-height:120px}
    .achips{display:flex;gap:6px;flex-wrap:wrap;margin:10px 0 2px;min-height:26px}
    .ac{font-size:12px;font-weight:700;padding:4px 10px;border-radius:99px;background:var(--bg);border:1px solid var(--line);color:var(--muted)}
    .ac.ok{background:var(--ok-soft);color:var(--ok);border-color:transparent}
    .ac.warn{background:var(--warn-soft);color:var(--warn);border-color:transparent}
    .chk{display:flex;gap:8px;align-items:center;text-transform:none;letter-spacing:0;font-size:12.5px;font-weight:600;color:var(--muted);margin:8px 0 12px}
    .chk input{width:auto}
    @media(max-width:480px){.arow{grid-template-columns:1fr;gap:5px}.sg button{font-size:12px;padding:8px 4px}}
    .pwbar{background:var(--surface);border:1px solid var(--warn);border-radius:var(--r-sm);padding:10px 12px;margin-bottom:10px}
    .pwfix{display:flex;gap:8px}.pwfix input{padding:8px 12px;font-size:14px}.pwfix .btn{flex:none}
    .pwp{margin:0 0 10px}
    .pwrow{display:flex;align-items:center;gap:8px;flex-wrap:wrap;background:var(--surface);border:1px solid var(--line);border-radius:var(--r-sm);padding:7px 10px;margin-bottom:6px;font-size:13px}
    .pwrow code{font-family:ui-monospace,Consolas,monospace;font-weight:700;background:var(--bg);padding:2px 8px;border-radius:7px;word-break:break-all}
    .pwrow span{color:var(--muted);font-weight:600}
    .pwrow .sp{flex:1}
    .verwarn{background:var(--warn-soft);color:var(--warn);border:1px solid var(--warn);border-radius:var(--r-sm);padding:10px 14px;font-size:13px;line-height:1.5}
    /* ===== Themed scrollbars ===== */
    *{scrollbar-width:thin;scrollbar-color:color-mix(in srgb,var(--muted) 40%,transparent) transparent}
    ::-webkit-scrollbar{width:10px;height:10px}
    ::-webkit-scrollbar-track{background:transparent}
    ::-webkit-scrollbar-thumb{background:color-mix(in srgb,var(--muted) 38%,transparent);border-radius:99px;border:2px solid transparent;background-clip:content-box}
    ::-webkit-scrollbar-thumb:hover{background:var(--brand);background-clip:content-box}
    ::-webkit-scrollbar-corner{background:transparent}
    .uids::-webkit-scrollbar,textarea::-webkit-scrollbar{width:8px;height:8px}
    /* ===== Progress bar ===== */
    #loader{position:fixed;top:0;left:0;right:0;height:3px;z-index:120;opacity:0;pointer-events:none;transition:opacity .2s;overflow:hidden;background:color-mix(in srgb,var(--brand) 14%,transparent)}
    #loader.on{opacity:1}
    #loader i{display:block;height:100%;width:38%;border-radius:99px;background:linear-gradient(90deg,transparent,var(--brand),#8AA2FF,transparent);animation:ldbar 1.05s ease-in-out infinite}
    @keyframes ldbar{from{transform:translateX(-110%)}to{transform:translateX(280%)}}
    @media(prefers-reduced-motion:reduce){#loader i{animation:none;width:100%}}
    .okring.sm{width:44px;height:44px;margin:0;flex:none;box-shadow:0 0 0 5px color-mix(in srgb,var(--ok-soft) 55%,transparent)}
    .okring.sm .ic{width:22px;height:22px}
    .skel{height:54px;border-radius:var(--r-sm);background:linear-gradient(90deg,var(--bg),color-mix(in srgb,var(--line) 70%,var(--bg)),var(--bg));background-size:200% 100%;animation:skl 1.2s linear infinite}
    @keyframes skl{to{background-position:-200% 0}}

    /* ===== FB IDs Vault polish ===== */
    html{overscroll-behavior-y:contain;-webkit-text-size-adjust:100%}
    body{-webkit-touch-callout:none}
    button,.nav,.chip,.seg,.sg,.brand,.pill,.sl,label{-webkit-user-select:none;user-select:none}
    input,textarea,.uids,code,.h-item,.found{-webkit-user-select:text;user-select:text}
    ::selection{background:color-mix(in srgb,var(--brand) 30%,transparent)}
    .hero b,.sc b,.kpi b,.tot,.qty,.cbadge,#sbAv,.pill{font-variant-numeric:tabular-nums}
    button{touch-action:manipulation}
    .btn,.chip,.mini,.icbtn,.rfbtn,.nav button,.seg button,.sg button{transition:transform .12s ease,background .15s,box-shadow .15s,border-color .15s,color .15s}
    .btn:active,.icbtn:active,.rfbtn:active,.mini:active,.seg button:active{transform:scale(.97)}
    .btn:disabled{opacity:.55;filter:saturate(.7);cursor:not-allowed;box-shadow:none}
    :is(button,select,input,textarea,a):focus-visible{outline:2px solid var(--brand);outline-offset:2px}
    input:focus-visible,textarea:focus-visible,select:focus-visible{outline:none}
    .logo{background:linear-gradient(145deg,#4D6DFF,#2036C9);box-shadow:0 8px 18px -8px rgba(47,84,245,.9),inset 0 1px 0 rgba(255,255,255,.25)}
    .logo .ic{stroke-width:1.7}
    .card{transition:border-color .2s,box-shadow .2s}
    .h-item:active{transform:scale(.995)}
    .toast{border-radius:14px;backdrop-filter:blur(10px)}
    textarea{resize:vertical;line-height:1.55}
    .modal .sheetbox{overscroll-behavior:contain}
    /* settings rows (Security) */
    .swlist{display:grid;gap:4px;padding-top:12px}
    .swrow{display:flex;align-items:center;justify-content:space-between;gap:16px;padding:14px 0;border-bottom:1px solid var(--line)}
    .swrow:last-child{border-bottom:0;padding-bottom:0}
    .swrow>div:first-child{min-width:0}
    .swrow b{display:block;font-size:15px;font-weight:700;color:var(--ink)}
    .swrow small{display:block;margin-top:2px;font-size:12.5px;line-height:1.4;color:var(--muted)}
    .swrow select{width:auto;min-width:132px;flex:none;padding-top:10px;padding-bottom:10px}
    .tgl{position:relative;flex:none;width:50px;height:30px;display:block;cursor:pointer}
    .tgl input{position:absolute;inset:0;width:100%;height:100%;opacity:0;margin:0;padding:0;cursor:pointer;z-index:1}
    .tgl i{position:absolute;inset:0;border-radius:99px;background:color-mix(in srgb,var(--muted) 35%,transparent);transition:background .2s}
    .tgl i::after{content:"";position:absolute;left:3px;top:3px;width:24px;height:24px;border-radius:50%;background:#fff;box-shadow:0 2px 6px rgba(0,0,0,.28);transition:transform .2s cubic-bezier(.3,.8,.3,1)}
    .tgl input:checked+i{background:linear-gradient(135deg,#4767FF,#2442D9)}
    .tgl input:checked+i::after{transform:translateX(20px)}
    .tgl input:focus-visible+i{outline:2px solid var(--brand);outline-offset:2px}
    .setbody .row2 .btn{white-space:nowrap;padding-left:12px;padding-right:12px}
    .swrow:has(+ .swrow.hidden){border-bottom:0;padding-bottom:0}
    @media(max-width:480px){.swrow{align-items:flex-start;flex-wrap:wrap}.swrow select{width:100%}.swrow .tgl{margin-top:2px}.swlist .swrow:has(.tgl){flex-wrap:nowrap}}
    .updbar{display:flex;align-items:center;gap:12px;flex-wrap:wrap;background:var(--brand-soft);color:var(--brand);border-color:color-mix(in srgb,var(--brand) 40%,transparent)}
    .updbar .t{flex:1;min-width:180px;line-height:1.45}
    .updbar .btn{width:auto;padding:9px 18px;border-radius:11px;font-size:13px;white-space:nowrap}
    .updbar .x{background:transparent;border:0;color:inherit;font-weight:700;cursor:pointer;padding:8px 6px}
    .hidden{display:none!important}
  </style>
</head>
<body>
<svg width="0" height="0" style="position:absolute" aria-hidden="true"><defs>
  <symbol id="i-vault" viewBox="0 0 24 24"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><rect x="8" y="8.6" width="8" height="6" rx="1.3"/><path d="M10 11.6h.01M12.6 10.9h1.9M12.6 12.7h1.9"/></symbol>
  <symbol id="i-menu" viewBox="0 0 24 24"><path d="M4 7h16M4 12h16M4 17h16"/></symbol>
  <symbol id="i-logout" viewBox="0 0 24 24"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4M16 17l5-5-5-5M21 12H9"/></symbol>
  <symbol id="i-share" viewBox="0 0 24 24"><circle cx="18" cy="5" r="3"/><circle cx="6" cy="12" r="3"/><circle cx="18" cy="19" r="3"/><path d="M8.6 13.5l6.8 4M15.4 6.5l-6.8 4"/></symbol>
  <symbol id="i-cut" viewBox="0 0 24 24"><circle cx="6" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><path d="M20 4L8.1 15.9M14.5 14.5L20 20M8.1 8.1L12 12"/></symbol>
  <symbol id="i-refresh" viewBox="0 0 24 24"><path d="M21 12a9 9 0 1 1-3-6.7"/><path d="M21 3v6h-6"/></symbol>
  <symbol id="i-lock" viewBox="0 0 24 24"><rect x="4" y="10" width="16" height="10" rx="2"/><path d="M8 10V7a4 4 0 0 1 8 0v3"/></symbol>
  <symbol id="i-sell" viewBox="0 0 24 24"><path d="M6 2 3 6v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2V6l-3-4z"/><path d="M3 6h18"/><path d="M16 10a4 4 0 0 1-8 0"/></symbol>
  <symbol id="i-orders" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M8 13h8M8 17h8"/></symbol>
  <symbol id="i-users" viewBox="0 0 24 24"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75"/></symbol>
  <symbol id="i-search" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/></symbol>
  <symbol id="i-add" viewBox="0 0 24 24"><rect x="3" y="3" width="18" height="18" rx="3"/><path d="M12 8v8M8 12h8"/></symbol>
  <symbol id="i-shield" viewBox="0 0 24 24"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><path d="m9 12 2 2 4-4"/></symbol>
  <symbol id="i-eye" viewBox="0 0 24 24"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8S1 12 1 12z"/><circle cx="12" cy="12" r="3"/></symbol>
  <symbol id="i-eyeoff" viewBox="0 0 24 24"><path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"/><path d="M1 1l22 22"/></symbol>
  <symbol id="i-box" viewBox="0 0 24 24"><path d="M21 8 12 3 3 8v8l9 5 9-5z"/><path d="m3 8 9 5 9-5M12 13v8"/></symbol>
  <symbol id="i-check" viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/></symbol>
  <symbol id="i-cal" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/></symbol>
  <symbol id="i-trend" viewBox="0 0 24 24"><path d="m3 17 6-6 4 4 8-8"/><path d="M14 7h7v7"/></symbol>
  <symbol id="i-coin" viewBox="0 0 24 24"><path d="M12 1v22"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/></symbol>
  <symbol id="i-zap" viewBox="0 0 24 24"><path d="M13 2 3 14h9l-1 8 10-12h-9z"/></symbol>
  <symbol id="i-clock" viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></symbol>
  <symbol id="i-copy" viewBox="0 0 24 24"><rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></symbol>
  <symbol id="i-down" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M7 10l5 5 5-5M12 15V3"/></symbol>
  <symbol id="i-undo" viewBox="0 0 24 24"><path d="M3 7v6h6"/><path d="M21 17a9 9 0 0 0-15-6.7L3 13"/></symbol>
  <symbol id="i-swap" viewBox="0 0 24 24"><path d="M17 1l4 4-4 4"/><path d="M3 11V9a4 4 0 0 1 4-4h14M7 23l-4-4 4-4"/><path d="M21 13v2a4 4 0 0 1-4 4H3"/></symbol>
  <symbol id="i-filter" viewBox="0 0 24 24"><path d="M22 3H2l8 9.46V19l4 2v-8.54z"/></symbol>
  <symbol id="i-gear" viewBox="0 0 24 24"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/></symbol>
  <symbol id="i-link" viewBox="0 0 24 24"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></symbol>
  <symbol id="i-x" viewBox="0 0 24 24"><path d="M18 6 6 18M6 6l12 12"/></symbol>
  <symbol id="i-trash" viewBox="0 0 24 24"><path d="M3 6h18M8 6V4h8v2M19 6l-1 14H6L5 6"/></symbol>
  <symbol id="i-sheet" viewBox="0 0 24 24"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18"/></symbol>
</defs></svg>

<!-- LOGIN -->
<section id="login">
  <form class="login-card" onsubmit="return doLogin(event)">
    <div class="logo"><svg class="ic"><use href="#i-vault"/></svg></div>
    <h1>FB IDs Vault</h1>
    <p>Enter your PIN to continue.</p>
    <div class="field"><label for="pin">PIN</label>
      <div class="pw"><input id="pin" type="password" inputmode="numeric" autocomplete="current-password" placeholder="••••" required>
        <button type="button" class="eye" onclick="eye(this)" aria-label="Show or hide PIN"><svg class="ic"><use href="#i-eye"/></svg></button></div></div>
    <button class="btn" id="loginBtn" type="submit">Unlock</button>
    <button class="btn ghost hidden" id="bioBtn" type="button" style="margin-top:10px" onclick="bioLogin()"><svg class="ic"><use href="#i-shield"/></svg>Unlock with fingerprint</button>
    <div class="err" id="loginErr"></div>
    <button type="button" class="linkbtn" onclick="openConn()"><svg class="ic"><use href="#i-sheet"/></svg><span>Google Sheet: <b id="connLoginName">Main sheet</b>. Change</span></button>
  </form>
</section>

<!-- APP -->
<div id="app" class="shell hidden">
  <div class="scrim" id="scrim" onclick="toggleMenu(false)"></div>
  <aside class="side">
    <div class="brand"><div class="logo"><svg class="ic"><use href="#i-vault"/></svg></div>FB IDs Vault</div>
    <nav class="nav" id="nav">
      <button data-t="sell" class="on" onclick="showTab('sell')"><svg class="ic"><use href="#i-sell"/></svg><span>Sell</span></button>
      <button data-t="orders" onclick="showTab('orders')"><svg class="ic"><use href="#i-orders"/></svg><span>Orders</span></button>
      <button data-t="cust" onclick="showTab('cust')"><svg class="ic"><use href="#i-users"/></svg><span>Customers</span></button>
      <button data-t="find" onclick="showTab('find')"><svg class="ic"><use href="#i-search"/></svg><span>UID search</span></button>
      <button data-t="add" onclick="showTab('add')"><svg class="ic"><use href="#i-add"/></svg><span>Add New Stock</span></button>
      <button data-t="ret" onclick="showTab('ret')"><svg class="ic"><use href="#i-swap"/></svg><span>Returns</span></button>
      <button data-t="sec" onclick="showTab('sec')"><svg class="ic"><use href="#i-gear"/></svg><span>Settings</span></button>
    </nav>
    <div class="sfoot">Remaining IDs<b id="sbAv">0</b><div class="meter" id="meter2"><i></i></div></div>
    <button type="button" class="logout rep" onclick="shareReport()"><svg class="ic"><use href="#i-share"/></svg>Share today's report</button>
    <button type="button" class="logout" onclick="logoutBtn()"><svg class="ic"><use href="#i-logout"/></svg>Logout</button>
  </aside>

  <div class="content">
    <header class="top">
      <button type="button" class="logo mlogo menubtn" onclick="toggleMenu()" aria-label="Menu"><svg class="ic"><use href="#i-menu"/></svg></button>
      <div><h1 id="pgT">Sell</h1><p id="pgS"></p></div>
      <span class="pill" id="pill">-</span>
      <button type="button" class="rfbtn" id="rfBtn" onclick="refreshNow()" title="Refresh from sheet" aria-label="Refresh"><svg class="ic"><use href="#i-refresh"/></svg></button>
      <button class="icbtn" onclick="lockApp()" title="Lock" aria-label="Lock"><svg class="ic"><use href="#i-lock"/></svg></button>
    </header>

    <main>
      <div class="verwarn hidden" id="verWarn"><b>Your Google Sheet script is old.</b> Passwords, Returns and categories will not work correctly until you update it: open the sheet, Extensions, Apps Script, paste the latest Code.gs, then Deploy, Manage deployments, Edit, New version, Deploy.<div id="verInfo" style="margin-top:6px;opacity:.85"></div><button type="button" class="btn sm ghost" style="margin-top:8px" onclick="showTab('sec')">Open Settings, Check connection</button></div>
      <section data-tab="sell" class="kpis2">
        <div class="hero" id="kMain">
          <div><span>Remaining stock</span><b id="sAv">0</b></div>
          <div class="hr"><span>Total IDs ever</span><b id="sTot">0</b></div>
          <div class="hbar"><i id="heroBar"></i></div>
        </div>
        <div class="sc">
          <div class="sr"><div class="sl sold">Sold</div>
            <div class="c"><b id="sToday">0</b><span>Today</span></div>
            <div class="c"><b id="sMonth">0</b><span>This month</span></div>
            <div class="c"><b id="sSold">0</b><span>Total</span></div></div>
          <div class="sr"><div class="sl earn">Earned</div>
            <div class="c"><b id="eToday">0</b><span>Today</span></div>
            <div class="c"><b id="eMonth">0</b><span>This month</span></div>
            <div class="c"><b id="eAll">0</b><span>Total</span></div></div>
        </div>
      </section>

      <div class="g2" data-tab="sell">
        <section class="card">
          <h3>Give IDs to customer</h3>
          <p class="sub">That many IDs are taken from stock and marked as Sold.</p>
          <div class="row3">
            <div class="field"><label for="cust">Customer name</label><input id="cust" list="custNames" autocomplete="off" placeholder="e.g. Ali Khan" oninput="drawNames()"><datalist id="custNames"></datalist></div>
            <div class="field"><label for="qty">Quantity</label><input id="qty" type="number" inputmode="numeric" min="1" placeholder="50" oninput="updTot()"></div>
            <div class="field"><label for="rate">Rate per ID</label><input id="rate" type="number" inputmode="decimal" min="0" placeholder="0" oninput="updTot()"></div>
          </div>
          <div class="nchips" id="nchips"></div>
          <label>Which IDs?</label>
          <div class="seg" id="seg">
            <button type="button" class="on" data-m="normal" onclick="setMode('normal')"><svg class="ic"><use href="#i-clock"/></svg><b>Normal</b><small>Oldest IDs first</small></button>
            <button type="button" data-m="fresh" onclick="setMode('fresh')"><svg class="ic"><use href="#i-zap"/></svg><b>Fresh</b><small>Newest added IDs</small></button>
          </div>
          <div class="row2">
            <div class="field"><label for="selTyS">Names type</label><select id="selTyS" onchange="selTy(this.value)"></select></div>
            <div class="field"><label for="selGeS">Gender</label><select id="selGeS" onchange="selGe(this.value)"></select></div>
          </div>
          <div class="selinfo" id="selInfo"></div>
          <div class="chips">
            <button class="chip" onclick="setQty(10)">10</button><button class="chip" onclick="setQty(25)">25</button>
            <button class="chip" onclick="setQty(50)">50</button><button class="chip" onclick="setQty(100)">100</button>
            <button class="chip" onclick="setQty(500)">500</button><span class="tot" id="tot"></span>
          </div>
          <button class="btn" id="sellBtn" onclick="sell()">Get IDs</button>
          <div class="err" id="sellErr"></div>
          <button type="button" class="btn ghost hidden" id="lastBtn" style="margin-top:8px" onclick="showModal('resModal')"><svg class="ic"><use href="#i-eye"/></svg>Show last IDs again</button>

        </section>
        <div class="stack">
          <section class="card"><h3>Last 7 days</h3><p class="sub">IDs sold per day</p><div class="bars" id="chart"></div></section>
          <section class="card"><h3>Stock by category</h3><p class="sub">Remaining IDs. Older untagged IDs are counted under Mix / Mix.</p><table class="cgrid" id="cgrid"></table></section>
          <section class="card"><h3>Stock level</h3><p class="sub" id="stTxt"></p><div class="meter" id="meter"><i></i></div></section>
        </div>
      </div>

      <section class="card hidden" data-tab="orders">
        <h3>Orders</h3><p class="sub">Filter by date or customer, export CSV, or undo a wrong order.</p>
        <div class="filters">
          <input id="fFrom" type="date" aria-label="From date"><input id="fTo" type="date" aria-label="To date">
          <input id="fCust" placeholder="Customer name" aria-label="Customer">
          <button class="btn sm" onclick="loadHistory()"><svg class="ic"><use href="#i-filter"/></svg>Filter</button>
          <button class="btn sm ghost" onclick="exportCsv()"><svg class="ic"><use href="#i-down"/></svg>CSV</button>
        </div>
        <div class="hist" id="hist"></div>
      </section>

      <section class="card hidden" data-tab="cust">
        <h3>Customers</h3><p class="sub">Top buyers first.</p><div class="field"><input id="custQ" placeholder="Search customer" oninput="drawCust()" aria-label="Search customer"></div><div class="hist" id="custList"></div>
      </section>

      <section class="card hidden" data-tab="find">
        <h3>UID search</h3><p class="sub">Enter any UID to see which customer got it, or whether it is still in stock.</p>
        <div class="row"><div class="field"><input id="fUid" placeholder="UID" aria-label="UID"></div><button class="btn" style="height:50px" onclick="findUid()"><svg class="ic"><use href="#i-search"/></svg>Search</button></div>
        <div class="found hidden" id="fRes"></div>
      </section>

      <section class="card hidden" data-tab="add" style="max-width:720px">
        <div class="ahead">
          <div><h3>Add New Stock</h3><p class="sub" style="margin:2px 0 0">Pick a category, paste the IDs, done.</p></div>
          <span class="cbadge" id="addInfo"></span>
        </div>
        <div class="err hidden" id="addWarn" style="margin:0 0 10px">Sheet script is old. Deploy the latest Code.gs, otherwise categories are not saved.</div>
        <div class="row2">
          <div class="field"><label for="addTyS">Names type</label><select id="addTyS" onchange="addTy(this.value)"></select></div>
          <div class="field"><label for="addGeS">Gender</label><select id="addGeS" onchange="addGe(this.value)"></select></div>
        </div>
        <div class="arow"><label>Password</label>
          <div class="sg" id="pwSeg">
            <button type="button" class="on" data-v="none" onclick="setPw('none')">None</button>
            <button type="button" data-v="same" onclick="setPw('same')">Same for all</button>
            <button type="button" data-v="line" onclick="setPw('line')">Different per ID</button>
          </div></div>
        <div class="field hidden" id="pwSame"><input id="addPass" oninput="cnt()" type="text" autocomplete="off" autocapitalize="off" spellcheck="false" placeholder="Password for every ID in this batch"></div>
        <div class="hint hidden" id="pwLine"><svg class="ic"><use href="#i-check"/></svg>Write each line as <b>UID|password</b></div>
        <div class="arow"><label>Position</label>
          <div class="sg" id="posSeg">
            <button type="button" class="on" data-v="end" onclick="setPos('end')">End &middot; Fresh</button>
            <button type="button" data-v="front" onclick="setPos('front')">Front &middot; used first</button>
          </div></div>
        <div class="lbrow"><label for="addTxt">UIDs</label>
          <div class="tools"><button type="button" class="mini" onclick="pasteIds()"><svg class="ic"><use href="#i-copy"/></svg>Paste</button><button type="button" class="mini" onclick="clearIds()"><svg class="ic"><use href="#i-x"/></svg>Clear</button></div></div>
        <textarea id="addTxt" placeholder="Paste one UID per line" oninput="cnt()"></textarea>
        <div class="achips" id="addChips"></div>
        <label class="chk"><input type="checkbox" id="addBad" onchange="cnt()">Also re-add IDs that were marked Bad</label>
        <button class="btn" id="addBtn" onclick="addStock()"></button>
        <div class="err" id="addErr"></div>
      </section>

      <section class="card hidden" data-tab="ret" style="max-width:720px">
        <h3>Returns &amp; Replacements</h3>
        <p class="sub" id="retSub"></p>
        <div class="seg" id="retSeg">
          <button type="button" class="on" data-m="return" onclick="setRet('return')"><svg class="ic"><use href="#i-undo"/></svg><b>Return</b><small>Customer gave IDs back</small></button>
          <button type="button" data-m="replace" onclick="setRet('replace')"><svg class="ic"><use href="#i-swap"/></svg><b>Replace</b><small>IDs that do not log in</small></button>
        </div>
        <div class="field"><label for="retTxt">UIDs</label>
          <textarea id="retTxt" placeholder="Paste one UID per line (a password after | is ignored)" oninput="cntRet()"></textarea></div>
        <div class="chips"><span id="retCnt" class="tot" style="margin:0;color:var(--muted)">0 IDs detected</span></div>
        <button class="btn" id="retBtn" onclick="runRet()"></button>
        <div class="err" id="retErr"></div>
        <div class="result hidden" id="retRes">
          <div class="result-head"><b id="retTitle"></b>
            <div class="acts">
              <button class="btn sm" onclick="copyText(RETNEW.join('\n'))"><svg class="ic"><use href="#i-copy"/></svg>Copy new IDs</button>
              <button class="btn sm ghost" onclick="shareText(RETNEW.join('\n'))"><svg class="ic"><use href="#i-share"/></svg>Share</button>
              <button class="btn sm ghost" onclick="saveTxt(RETNEW.join('\n'),'replacements')"><svg class="ic"><use href="#i-down"/></svg>.txt</button>
            </div></div>
          <div class="uids" id="retOut"></div>
          <p class="sub" id="retNote" style="margin-top:10px"></p>
        </div>
      </section>

      <div class="stack hidden" data-tab="sec" style="max-width:600px">
        <section class="card setcard">
          <div class="sethead"><div class="seticon"><svg class="ic"><use href="#i-check"/></svg></div>
            <div><h3>Check connection</h3><p class="sub">See which Google Sheet this app is talking to.</p></div></div>
          <div class="setbody">
            <button class="btn ghost" id="chkBtn" onclick="checkConn()"><svg class="ic"><use href="#i-refresh"/></svg>Check now</button>
            <div class="sub" id="chkOut" style="margin-top:12px;line-height:1.7"></div>
          </div>
        </section>
        <section class="card setcard">
          <div class="sethead"><div class="seticon sh"><svg class="ic"><use href="#i-sheet"/></svg></div>
            <div><h3>Google Sheet connection</h3><p class="sub">Every sheet keeps its own stock, orders and PIN.</p></div></div>
          <div class="setbody">
            <div class="connbox"><span class="dot"></span><div class="cmeta"><b id="connNow">-</b><span id="connNowUrl"></span></div><span class="tag ok">Active</span></div>
            <button class="btn ghost" onclick="openConn()"><svg class="ic"><use href="#i-link"/></svg>Manage Google Sheets</button>
          </div>
        </section>
        <section class="card setcard">
          <div class="sethead"><div class="seticon"><svg class="ic"><use href="#i-shield"/></svg></div>
            <div><h3>Change PIN</h3><p class="sub">Works on every device that uses this sheet.</p></div></div>
          <div class="setbody pinrow">
            <div class="field"><label for="curP">Current PIN</label><div class="pw"><input id="curP" type="password" inputmode="numeric" autocomplete="off" placeholder="Enter current PIN"><button type="button" class="eye" onclick="eye(this)" aria-label="Show or hide PIN"><svg class="ic"><use href="#i-eye"/></svg></button></div></div>
            <div class="field"><label for="newP">New PIN</label><div class="pw"><input id="newP" type="password" inputmode="numeric" autocomplete="off" placeholder="4 to 10 digits"><button type="button" class="eye" onclick="eye(this)" aria-label="Show or hide PIN"><svg class="ic"><use href="#i-eye"/></svg></button></div></div>
            <div class="field"><label for="cnfP">Confirm new PIN</label><div class="pw"><input id="cnfP" type="password" inputmode="numeric" autocomplete="off" placeholder="Type it again"><button type="button" class="eye" onclick="eye(this)" aria-label="Show or hide PIN"><svg class="ic"><use href="#i-eye"/></svg></button></div></div>
            <div class="hint"><svg class="ic"><use href="#i-check"/></svg>Use 4 to 10 digits. Avoid easy PINs like 1234.</div>
            <button class="btn" id="pinBtn" onclick="changePin()"><svg class="ic"><use href="#i-shield"/></svg>Save new PIN</button>
            <div class="err" id="pinErr"></div>
          </div>
        </section>
      <section class="card setcard">
        <div class="sethead"><div class="seticon"><svg class="ic"><use href="#i-lock"/></svg></div>
          <div><h3>Security</h3><p class="sub">Protect the IDs on this phone.</p></div></div>
        <div class="setbody swlist">
          <div class="swrow"><div><b>Auto-lock</b><small>Ask for the PIN again after the app was in the background</small></div>
            <select id="alSel" onchange="fbSetAutoLock(this.value)" aria-label="Auto-lock time">
              <option value="0">Off</option><option value="1">1 minute</option><option value="5">5 minutes</option><option value="15">15 minutes</option><option value="30">30 minutes</option>
            </select></div>
          <div class="swrow hidden" id="secRow"><div><b>Block screenshots</b><small>Hides IDs from screenshots and the recent apps screen</small></div>
            <label class="tgl"><input type="checkbox" id="secChk" onchange="fbSetSecure(this.checked)" aria-label="Block screenshots"><i></i></label></div>
        </div>
      </section>
      <section class="card setcard">
        <div class="sethead"><div class="seticon sh"><svg class="ic"><use href="#i-down"/></svg></div>
          <div><h3>Backup</h3><p class="sub">Save your Google Sheet links, saved names and settings to a file. Your IDs stay in the Google Sheet. The PIN is never saved in the file.</p></div></div>
        <div class="setbody">
          <div class="row2">
            <button class="btn ghost" onclick="fbBackupExport()"><svg class="ic"><use href="#i-down"/></svg>Export</button>
            <button class="btn ghost" onclick="document.getElementById('bkFile').click()"><svg class="ic"><use href="#i-refresh"/></svg>Restore</button>
          </div>
          <input type="file" id="bkFile" accept=".json,application/json" class="hidden" onchange="fbBackupImport(this)">
          <div class="err" id="bkMsg"></div>
        </div>
      </section>
      <section class="card setcard hidden" id="updCard">
        <div class="sethead"><div class="seticon"><svg class="ic"><use href="#i-refresh"/></svg></div>
          <div><h3>App updates</h3><p class="sub">Version <b id="verNow">-</b>. New versions are published on GitHub.</p></div></div>
        <div class="setbody">
          <div class="field"><label for="repoIn">GitHub repository (username/repo)</label>
            <input id="repoIn" placeholder="username/FB-IDs-Vault" autocomplete="off" autocapitalize="off" spellcheck="false" onchange="fbSetRepo(this.value)"></div>
          <button class="btn ghost" id="updBtn" onclick="fbCheckUpdate(true)"><svg class="ic"><use href="#i-refresh"/></svg>Check for updates</button>
          <div class="sub" id="updMsg" style="margin-top:12px;line-height:1.6"></div>
        </div>
      </section>
      </div>
    </main>
  </div>
</div>
<!-- PIN CHANGED -->
<div class="modal center hidden" id="pinModal" onclick="if(event.target===this)closeModal()" role="dialog" aria-modal="true" aria-labelledby="pmT">
  <div class="sheetbox okbox">
    <div class="okring"><svg class="ic"><use href="#i-check"/></svg></div>
    <h3 id="pmT">PIN changed</h3>
    <p class="sub">Your new PIN is active now. Use it the next time you unlock the app, on every device.</p>
    <button class="btn" onclick="closeModal()">Done</button>
  </div>
</div>
<!-- STOCK ADDED -->
<div class="modal center hidden" id="addModal" onclick="if(event.target===this)closeModal()" role="dialog" aria-modal="true" aria-labelledby="amT">
  <div class="sheetbox okbox">
    <div class="okring"><svg class="ic"><use href="#i-check"/></svg></div>
    <h3 id="amT">Stock add ho gaya</h3>
    <div class="sub" id="amBody" style="text-align:left;line-height:1.8"></div>
    <button class="btn" onclick="closeModal()">Done</button>
  </div>
</div>
<!-- GET IDS RESULT -->
<div class="modal hidden" id="resModal" onclick="if(event.target===this)closeModal()" role="dialog" aria-modal="true" aria-labelledby="resH">
  <div class="sheetbox">
    <div class="mhead">
      <div style="display:flex;gap:12px;align-items:center"><div class="okring sm"><svg class="ic"><use href="#i-check"/></svg></div>
        <div><h3 id="resH">IDs ready</h3><p class="sub" id="resTitle" style="margin:2px 0 0"></p></div></div>
      <button type="button" class="icbtn" onclick="closeModal()" aria-label="Close"><svg class="ic"><use href="#i-x"/></svg></button></div>
    <div class="kv" id="resMeta"></div>
    <div id="result">
      <div class="result-head">
        <div class="acts">
          <button class="btn sm" id="rcAll" onclick="copyUids()"><svg class="ic"><use href="#i-copy"/></svg>Copy all</button>
          <button class="btn sm ghost hidden" id="rcIds" onclick="copyIdsOnly(LASTIDS)"><svg class="ic"><use href="#i-copy"/></svg>Copy IDs</button>
          <button class="btn sm ghost hidden" id="rcPw" onclick="copyPassOnly(LASTIDS)"><svg class="ic"><use href="#i-lock"/></svg>Copy passwords</button>
          <button class="btn sm ghost" onclick="cutUids()"><svg class="ic"><use href="#i-cut"/></svg>Cut</button>
          <button class="btn sm ghost" onclick="shareUids()"><svg class="ic"><use href="#i-share"/></svg>Share</button>
          <button class="btn sm ghost" onclick="downloadUids()"><svg class="ic"><use href="#i-down"/></svg>.txt</button>
        </div></div>
      <div id="resPw"></div>
      <div class="pwbar hidden" id="pwBar">
        <div class="hint" id="pwBarNote" style="margin:0 0 8px"></div>
        <div class="pwfix" id="pwFixRow"><input id="pwFix" type="text" autocomplete="off" autocapitalize="off" spellcheck="false" placeholder="Password for these IDs"><button type="button" class="btn sm" id="pwFixBtn" onclick="fixPass()">Save password</button></div>
      </div>
      <div class="uids" id="uids"></div>
    </div>
    <button class="btn" style="margin-top:14px" onclick="closeModal()">Done</button>
  </div>
</div>
<!-- ORDER DETAIL -->
<div class="modal hidden" id="ordModal" onclick="if(event.target===this)closeModal()" role="dialog" aria-modal="true" aria-labelledby="odTitle">
  <div class="sheetbox">
    <div class="mhead"><div><h3 id="odTitle"></h3><p class="sub" id="odSub" style="margin:2px 0 0"></p></div><span class="tag" id="odTag"></span>
      <button type="button" class="icbtn" onclick="closeModal()" aria-label="Close"><svg class="ic"><use href="#i-x"/></svg></button></div>
    <div class="kv" id="odMeta"></div>
    <div class="result-head"><b id="odCount"></b>
      <div class="acts">
        <button class="btn sm" id="odAll" onclick="copyText(lines(OD.uids))"><svg class="ic"><use href="#i-copy"/></svg>Copy all</button>
        <button class="btn sm ghost hidden" id="odIds" onclick="copyIdsOnly(OD.uids)"><svg class="ic"><use href="#i-copy"/></svg>Copy IDs</button>
        <button class="btn sm ghost hidden" id="odPw" onclick="copyPassOnly(OD.uids)"><svg class="ic"><use href="#i-lock"/></svg>Copy passwords</button>
        <button class="btn sm ghost" onclick="saveTxt(lines(OD.uids),OD.customer+'_order'+OD.order)"><svg class="ic"><use href="#i-down"/></svg>.txt</button>
        <button class="btn sm ghost" onclick="shareText(lines(OD.uids))"><svg class="ic"><use href="#i-share"/></svg>Share</button>
      </div></div>
    <div id="odPwp"></div>
    <div class="uids od-list" id="odUids"></div>
    <div class="facts" style="margin-top:14px"><button id="odUndo" class="btn sm ghost" onclick="undo(OD.order,true)"><svg class="ic"><use href="#i-undo"/></svg>Undo this order</button></div>
  </div>
</div>

<!-- GOOGLE SHEET CONNECTIONS -->
<div class="modal hidden" id="connModal" onclick="if(event.target===this)closeModal()" role="dialog" aria-modal="true" aria-labelledby="connTitle">
  <div class="sheetbox">
    <div class="mhead"><div><h3 id="connTitle">Google Sheets</h3><p class="sub" style="margin:2px 0 0">Connect this app to any Google Sheet that has the FB IDs Vault script.</p></div>
      <button type="button" class="icbtn" onclick="closeModal()" aria-label="Close"><svg class="ic"><use href="#i-x"/></svg></button></div>
    <div class="hist" id="connList"></div>
    <h3 class="mh2">Connect a new sheet</h3>
    <div class="field"><label for="cName">Name (optional)</label><input id="cName" placeholder="e.g. Shop 2" autocomplete="off"></div>
    <div class="field"><label for="cUrl">Web app link (ends with /exec)</label><input id="cUrl" placeholder="https://script.google.com/macros/s/.../exec" autocomplete="off" autocapitalize="off" spellcheck="false" inputmode="url"></div>
    <div class="cmsg" id="connMsg"></div>
    <div class="facts"><button class="btn sm ghost" id="cTest" onclick="testConn()">Test link</button><button class="btn sm" id="cAdd" onclick="addConn()"><svg class="ic"><use href="#i-link"/></svg>Save &amp; connect</button></div>
    <p class="sub" style="margin:16px 0 0">Setup: open your Google Sheet, then Extensions, Apps Script. Paste the FB IDs Vault script (Code.gs), then Deploy, New deployment, Web app, Execute as Me, Who has access Anyone. Copy the link that ends with /exec and paste it above. First PIN is 1234.</p>
  </div>
</div>

<div id="loader" aria-hidden="true"><i></i></div>
<div class="toast" id="toast"></div>

<script>
  const APP_VER = 6; let SCRIPT_VER = 0;
  let LASTIDS = [], LASTMISS = [], PIN = '', LAST = [], LASTNAME = '', MODE = 'normal', ROWS = [];
  const $ = id => document.getElementById(id);
  const ic = n => '<svg class="ic"><use href="#i-' + n + '"/></svg>';
  /* ===== OFFLINE ENGINE: local copy of the sheet + sync queue ===== */
  /* ===== Google Sheet connections: the app can talk to any sheet that has the FB IDs Vault script ===== */
  const DEFAULT_URL = 'https://script.google.com/macros/s/AKfycby_INdVNu3yBg-nMJTPbQSVNysJz738URMbLsW9Zxa_VRoni0d66JtIBu_QfsyOrpA/exec';
  const CACHE_KEYS = ['uvs','uvq','uvp','uv','uvt','uvf','uvl','uvn','uvnames'];
  const rawGet = k => { try { return localStorage.getItem(k); } catch (e) { return null; } };
  function savedConns(){ try{ const l=JSON.parse(rawGet('uvconns')||'[]'); return Array.isArray(l)?l.filter(c=>c&&c.id&&c.url):[]; }catch(e){ return []; } }
  function allConns(){ return [{id:'default',name:'Main sheet',url:DEFAULT_URL,builtin:true}].concat(savedConns()); }
  function activeConn(){ const id=rawGet('uvact')||'default', l=allConns(); return l.find(c=>c.id===id)||l[0]; }
  const APP_URL = activeConn().url;
  const LOW = 50;
  let PW_OLD = false, TAGS_OLD = false, CATS = {}, SEL = {ty:'',ge:''}, ADDC = {ty:'X',ge:'X'}, LASTERR = '', DB = null, SNAP = null, Q = [], RES = {}, ONLINE = false, SP = Promise.resolve(), LASTPULL = 0;
  const LS = {
    get: k => { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } },
    set: (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { toast('Phone storage is full'); } }
  };
  const newId = () => 'o' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  function nextN(){let n=1;try{n=Number(localStorage.getItem('uvn')||0)+1}catch(e){}try{localStorage.setItem('uvn',n)}catch(e){}return n}
  const TYPES=[['E','English'],['M','Muslim'],['H','Hindu'],['X','Mix']], GENS=[['M','Male'],['F','Female'],['X','Mix']];
  const catOf=u=>(DB&&DB.tags&&DB.tags[u])||'XX';
  const fmt=u=>(DB&&DB.pw&&DB.pw[u])?u+'|'+DB.pw[u]:u;   // change '|' here if you want a different separator
  const lines=a=>a.map(fmt).join('\n');
  function parseEntries(text,hasPw){
    const m=new Map();
    if(hasPw){String(text||'').split(/\r?\n/).forEach(l=>{l=l.trim();if(!l)return;const x=l.match(/^(\S+?)(?:[\s|,;:]+(.+))?$/);if(x&&!m.has(x[1]))m.set(x[1],(x[2]||'').trim())})}
    else String(text||'').split(/[\s,;]+/).filter(Boolean).forEach(u=>{if(!m.has(u))m.set(u,'')});
    return Array.from(m.entries());
  }
  const catMatch=(u,ty,ge)=>{const c=catOf(u);return(!ty||c[0]===ty)&&(!ge||c[1]===ge)};
  const cc=(c,ty,ge)=>{let n=0;for(const k in c)if((!ty||k[0]===ty)&&(!ge||k[1]===ge))n+=c[k];return n};
  const catName=(ty,ge)=>(ty?TYPES.find(x=>x[0]===ty)[1]:'Any name')+', '+(ge?GENS.find(x=>x[0]===ge)[1]:'any gender');
  function index(){
    DB.set=new Set(DB.stock); DB.sold={};
    DB.sales.forEach(x=>{if(x.status!=='Undone')x.uids.forEach(u=>{DB.sold[u]=x})});
  }
  function applyOp(op){
    if(op.t==='add'){if(op.pws)Object.assign(DB.pw,op.pws);const rc=op.cat&&op.cat!=='XX',nw=[];op.uids.forEach(u=>{if(DB.sold[u])return;if(!DB.set.has(u)){nw.push(u);DB.set.add(u)}if(rc)DB.tags[u]=op.cat});DB.stock=op.front?nw.concat(DB.stock):DB.stock.concat(nw)}
    else if(op.t==='sell'){
      const take=new Set(op.uids); DB.stock=DB.stock.filter(u=>!take.has(u));
      DB.sales.push({order:'P'+op.n,date:op.date,customer:op.customer,qty:op.uids.length,rate:op.rate,amount:op.rate*op.uids.length,uids:op.uids.slice(),mode:op.mode,status:'Pending',opId:op.id});
    }
    index();
  }
  function rebuild(){
    const s=SNAP||{stock:[],sales:[],bad:[]};
    DB={stock:s.stock.slice(),sales:s.sales.slice(),bad:s.bad||[],tags:Object.assign({},s.tags||{}),pw:Object.assign({},s.pw||{})}; index();
    Q.forEach(applyOp);
  }
  const netMode=()=>(typeof google!=='undefined'&&google.script&&google.script.run)?'page':'link';
  let BUSY=0, LT=0;
  function busy(d){BUSY=Math.max(0,BUSY+d);clearTimeout(LT);if(BUSY>0)LT=setTimeout(()=>$('loader').classList.add('on'),120);else $('loader').classList.remove('on')}
  async function net(req){busy(1);try{return await net0(req)}finally{busy(-1)}}
  async function net0(req){
    const c=new AbortController(), t=setTimeout(()=>c.abort(),60000);
    try{
      const r=await fetch(APP_URL,{method:'POST',redirect:'follow',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify(Object.assign({},req,{pin:PIN})),signal:c.signal});
      const txt=await r.text();
      try{return JSON.parse(txt)}catch(x){throw new Error('Server did not return data. Check that web app access is Anyone and the latest version is deployed.')}
    }finally{clearTimeout(t)}
  }
  async function doSync(){
    const sent=Q.slice();
    try{
      const r=await net({action:'sync',ops:sent});
      if(!r.ok){if(!$('app').classList.contains('hidden'))toast(r.error);return{err:r.error}}
      r.res.forEach(x=>{RES[x.id]=x;if(!x.ok)toast(x.error)});
      Q=Q.slice(sent.length); LS.set('uvq',Q);
      TAGS_OLD=!!(r.snap&&r.snap.tags===undefined); PW_OLD=!!(r.snap&&r.snap.pw===undefined); SCRIPT_VER=(r.snap&&r.snap.ver)||0; SNAP=r.snap; LS.set('uvs',SNAP); LASTPULL=Date.now(); ONLINE=true; rebuild();
      return{};
    }catch(e){ONLINE=false;LASTERR=String((e&&e.message)||e);return{net:true}}
  }
  function sync(){const p=SP.then(doSync);SP=p;return p}
  function stats(){
    const t0=new Date();t0.setHours(0,0,0,0);
    const T=t0.getTime(), m0=new Date(t0.getFullYear(),t0.getMonth(),1).getTime();
    let today=0,month=0,eT=0,eM=0,eA=0,sold=0;
    const last7=[];for(let i=6;i>=0;i--){const d=new Date(t0);d.setDate(d.getDate()-i);last7.push({k:d.getTime(),l:d.toLocaleDateString('en-US',{weekday:'short'}),n:0})}
    DB.sales.forEach(x=>{
      if(x.status==='Undone')return;
      const n=x.uids.length,a=Number(x.amount)||0; sold+=n; eA+=a;
      if(x.date>=m0){month+=n;eM+=a;if(x.date>=T){today+=n;eT+=a}}
      const dd=new Date(x.date);dd.setHours(0,0,0,0);const y=last7.find(z=>z.k===dd.getTime());if(y)y.n+=n;
    });
    const bc={};DB.stock.forEach(u=>{const c=catOf(u);bc[c]=(bc[c]||0)+1});CATS=bc;
    const av=DB.stock.length;
    return{available:av,sold,total:av+sold,today,month,earnToday:eT,earnMonth:eM,earnAll:eA,last7,low:av<LOW,lowAt:LOW,pending:Q.length,online:ONLINE};
  }
  const fmtD=(t,full)=>new Date(t).toLocaleString('en-GB',full?{day:'2-digit',month:'short',year:'numeric',hour:'2-digit',minute:'2-digit',hour12:true}:{day:'2-digit',month:'short',year:'numeric'});
  function localSell(q){
    const customer=String(q.customer||'').trim(), qty=Number(q.qty), rate=Number(q.rate)||0, fresh=q.mode==='fresh';
    if(!customer)return{ok:false,error:'Enter customer name'};
    if(!qty||qty<1||qty%1)return{ok:false,error:'Enter a valid quantity'};
    const ty=q.ty||'',ge=q.ge||'', pool=(ty||ge)?DB.stock.filter(u=>catMatch(u,ty,ge)):DB.stock;
    if(pool.length<qty)return{ok:false,error:'Only '+pool.length+' IDs left'+((ty||ge)?' in this category':'')+', '+qty+' requested'};
    const uids=fresh?pool.slice(pool.length-qty):pool.slice(0,qty);
    const op={id:newId(),t:'sell',customer,rate,mode:fresh?'Fresh':'Normal',date:Date.now(),uids,n:nextN(),ty,ge};
    Q.push(op); LS.set('uvq',Q); applyOp(op); return op;
  }
  function localAdd(text,cat,hasPw,pass,front,allowBad){
    const ent=parseEntries(text,hasPw).map(([u,p])=>[u,p||pass||'']);
    if(!ent.length)return{ok:false,error:'No UIDs found'};
    const code=/^[EMHX][MFX]$/.test(cat)?cat:'XX', pws={}, fresh=[], moved=[], sk={stock:0,sold:0,bad:0};
    const badSet=new Set((DB.bad||[]).map(x=>x[0]));
    ent.forEach(([u,p])=>{
      if(DB.sold[u]){sk.sold++;return}
      if(DB.set.has(u)){
        if((code!=='XX'&&catOf(u)!==code)||(p&&DB.pw[u]!==p)){moved.push(u);if(p)pws[u]=p}else sk.stock++;
        return;
      }
      if(badSet.has(u)&&!allowBad){sk.bad++;return}
      fresh.push(u);if(p)pws[u]=p;
    });
    const all=fresh.concat(moved);
    if(all.length){const op={id:newId(),t:'add',uids:all,date:Date.now(),cat:code,pws,front:!!front,bad:!!allowBad};Q.push(op);LS.set('uvq',Q);applyOp(op)}
    return{ok:true,added:fresh.length,moved:moved.length,duplicates:sk.stock+sk.sold+sk.bad,skip:sk,uids:all};
  }
  async function call(req){
    const a=req.action;
    if(a==='login'){
      const saved=localStorage.getItem('uvp');
      if(SNAP&&saved!==null&&PIN===saved&&localStorage.getItem('uv')===PIN){ // auto sign-in: open instantly, sync in background
        sync().then(()=>{if(!$('app').classList.contains('hidden'))renderStats(stats())});
        return{ok:true,stats:stats()};
      }
      const r=await sync();
      if(r.err)return{ok:false,error:r.err};
      if(r.net){
        if(!SNAP||saved===null)return{ok:false,error:'Connect to the internet for the first login'+(LASTERR?' ('+LASTERR+')':'')};
        if(PIN!==saved)return{ok:false,error:'Wrong PIN'};
      }else{try{localStorage.setItem('uvp',PIN)}catch(e){}}
      return{ok:true,stats:stats()};
    }
    if(a==='stats')return{ok:true,stats:stats()};
    if(a==='sell'){
      const op=localSell(req); if(op.ok===false)return op;
      await sync();
      const r=RES[op.id];
      if(r&&!r.ok)return{ok:false,error:r.error};
      return{ok:true,uids:(r&&r.uids)||op.uids,order:r?r.order:'pending',amount:op.rate*op.uids.length,stats:stats()};
    }
    if(a==='add'){const r=localAdd(req.text,req.cat,req.hasPw,req.pass,req.front,req.bad); if(!r.ok)return r; await sync(); r.stats=stats(); return r}
    if(a==='history'){
      const from=req.from?new Date(req.from+'T00:00:00').getTime():0, to=req.to?new Date(req.to+'T23:59:59').getTime():Infinity, cu=String(req.customer||'').toLowerCase();
      return{ok:true,history:DB.sales.slice().reverse().filter(x=>x.date>=from&&x.date<=to&&(!cu||String(x.customer).toLowerCase().indexOf(cu)>=0)).slice(0,200)
        .map(x=>({order:x.order,date:fmtD(x.date,1),customer:x.customer,qty:x.qty,rate:x.rate,amount:x.amount,uids:x.uids,mode:x.mode,status:x.status}))};
    }
    if(a==='customers'){
      const m={};
      DB.sales.forEach(x=>{if(x.status==='Undone')return;const c=m[x.customer]=m[x.customer]||{name:x.customer,orders:0,qty:0,amount:0,last:0};c.orders++;c.qty+=Number(x.qty)||0;c.amount+=Number(x.amount)||0;c.last=Math.max(c.last,x.date)});
      return{ok:true,list:Object.keys(m).map(k=>m[k]).sort((p,q)=>q.amount-p.amount||q.qty-p.qty).map(c=>(c.last=fmtD(c.last),c))};
    }
    if(a==='find'){
      const u=String(req.uid||'').trim(); if(!u)return{ok:false,error:'Enter a UID'};
      const s=DB.sold[u]; if(s)return{ok:true,where:'sold',customer:s.customer,date:fmtD(s.date,1),order:s.order};
      if(DB.set.has(u))return{ok:true,where:'stock'};
      const b=DB.bad.find(x=>x[0]===u); if(b)return{ok:true,where:'bad',customer:b[1],order:b[2]};
      return{ok:true,where:'none'};
    }
    // undo, replace, change PIN: these need internet
    try{
      const sy=await sync();
      if(sy.net)return{ok:false,error:'This needs an internet connection'};
      if(Q.length)return{ok:false,error:'Waiting to sync. Try again in a moment'};
      const r=await net(req);
      if(r.ok){
        if(a==='changepin'){PIN=req.newPin;try{localStorage.setItem('uvp',PIN)}catch(e){}}
        await sync(); r.stats=stats();
      }
      return r;
    }catch(e){return{ok:false,error:'This needs an internet connection'}}
  }
  SNAP=LS.get('uvs'); Q=LS.get('uvq')||[]; rebuild();
  window.addEventListener('online',()=>{sync().then(()=>{if(!$('app').classList.contains('hidden'))renderStats(stats())})});

  const esc = v => String(v).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
  const money = n => Number(n || 0).toLocaleString();
  const PG = {sell:['Sell','Give IDs to a customer'], orders:['Orders','History, export and undo'], cust:['Customers','Your top buyers'],
    find:['UID search','Find where any UID went'], add:['Add New Stock','Load fresh IDs'], sec:['Settings','Google Sheet and PIN'], ret:['Returns & Replace','Take IDs back or replace bad ones']};

  function toast(m){const t=$('toast');t.textContent=m;t.classList.add('show');setTimeout(()=>t.classList.remove('show'),1800)}
  function toggleMenu(f){const a=$('app');a.classList.toggle('menu-open',f===undefined?!a.classList.contains('menu-open'):f)}
  document.addEventListener('keydown',e=>{if(e.key==='Escape')toggleMenu(false)});
  function lockApp(){
    toggleMenu(false);
    try{localStorage.removeItem('uv');localStorage.removeItem('uvt')}catch(e){}
    PIN=''; LAST=[]; ROWS=[];
    ['pin','cust','qty','rate','addTxt','addPass','retTxt','curP','newP','cnfP','fUid'].forEach(i=>{const e=$(i); if(e)e.value='';});
    ['fRes','retRes'].forEach(i=>$(i).classList.add('hidden')); hideModals(); LASTIDS=[]; $('lastBtn').classList.add('hidden');
    $('loginErr').textContent=''; $('loginBtn').disabled=false;
    const p=$('pin'); p.type='password'; p.nextElementSibling.innerHTML=ic('eye');
    showTab('sell'); setMode('normal'); updTot(); cnt();
    $('app').classList.add('hidden'); $('login').classList.remove('hidden');
    setTimeout(()=>p.focus(),50);
  }
  function setQty(n){$('qty').value=n; updTot()}
  function updTot(){const q=Number($('qty').value),r=Number($('rate').value);$('tot').textContent=(q&&r)?'Total '+money(q*r):''}
  function setMode(m){MODE=m; document.querySelectorAll('#seg button').forEach(b=>b.classList.toggle('on',b.dataset.m===m))}
  function eye(b){const i=b.previousElementSibling, show=i.type==='password'; i.type=show?'text':'password'; b.innerHTML=ic(show?'eyeoff':'eye')}
  const ADDS={pos:'end',pw:'none'};
  const segSet=(id,v)=>document.querySelectorAll('#'+id+' button').forEach(b=>b.classList.toggle('on',b.dataset.v===v));
  function setPos(v){ADDS.pos=v;segSet('posSeg',v);cnt()}
  function setPw(v){
    ADDS.pw=v;segSet('pwSeg',v);$('pwSame').classList.toggle('hidden',v!=='same');$('pwLine').classList.toggle('hidden',v!=='line');
    $('addTxt').placeholder=v==='line'?'UID|password, one per line':'Paste one UID per line';
    if(v==='same')setTimeout(()=>$('addPass').focus(),30); cnt();
  }
  function clearIds(){$('addTxt').value='';cnt()}
  async function pasteIds(){
    try{const t=await navigator.clipboard.readText();if(t){$('addTxt').value=t;cnt();return}}catch(e){}
    toast('Long-press the box and choose Paste'); $('addTxt').focus();
  }
  function cnt(){
    const ent=parseEntries($('addTxt').value,ADDS.pw==='line'), allowBad=$('addBad').checked;
    const badSet=new Set(((DB&&DB.bad)||[]).map(x=>x[0]));
    let nw=0,st=0,so=0,bd=0;
    if(DB)ent.forEach(([u])=>{if(DB.sold[u])so++;else if(DB.set.has(u))st++;else if(badSet.has(u)&&!allowBad)bd++;else nw++});
    const c=[];
    if(!ent.length)c.push('<span class="ac">Paste UIDs above</span>');
    else{
      c.push('<span class="ac ok">'+money(nw)+' new</span>');
      if(st)c.push('<span class="ac warn">'+money(st)+' already in stock</span>');
      if(so)c.push('<span class="ac warn">'+money(so)+' already sold</span>');
      if(bd)c.push('<span class="ac warn">'+money(bd)+' marked Bad</span>');
      if(ADDS.pw==='line'){const wp=ent.filter(e=>e[1]),dp=new Set(wp.map(e=>e[1]));c.push('<span class="ac '+(wp.length===ent.length?'ok':'warn')+'">'+wp.length+' with password'+(dp.size>1?' ('+dp.size+' different)':'')+'</span>');if(wp.length<ent.length)c.push('<span class="ac warn">'+(ent.length-wp.length)+' without</span>')}
      else if(ADDS.pw==='same'&&$('addPass').value.trim())c.push('<span class="ac ok">same password on all</span>');
      c.push('<span class="ac">'+(ADDS.pos==='front'?'to the front':'to the end')+'</span>');
    }
    $('addChips').innerHTML=c.join('');
    $('addBtn').innerHTML=ic('add')+(nw?'Add '+money(nw)+' new ID'+(nw===1?'':'s'):(st?'Update existing IDs':'Add to stock'));
  }

  /* ===== 24h session + wrong-PIN lockout ===== */
  const SESSION_MS=24*60*60*1000;
  const lsSet=(k,v)=>{try{localStorage.setItem(k,v)}catch(e){}};
  const lsDel=k=>{try{localStorage.removeItem(k)}catch(e){}};
  function sessionExpired(){const t=Number(localStorage.getItem('uvt')||0);return !!t&&Date.now()-t>SESSION_MS}
  async function refreshNow(){
    const b=$('rfBtn'); if(b.classList.contains('spin'))return; b.classList.add('spin');
    const r=await sync(); b.classList.remove('spin');
    renderStats(stats());
    const ot=document.querySelector('[data-tab=orders]'); if(ot&&!ot.classList.contains('hidden')&&typeof loadHistory==='function')try{loadHistory()}catch(e){}
    if(r.net)toast('No internet, showing saved data'); else if(!r.err)toast('Updated from sheet');
  }
  function checkSession(){
    if($('app').classList.contains('hidden'))return;
    if(sessionExpired()){lockApp();refreshBio();$('loginErr').textContent='Session expired after 24 hours. Enter your PIN again.'}
  }
  setInterval(checkSession,60000);
  document.addEventListener('visibilitychange',()=>{if(!document.hidden){checkSession();if(!$('app').classList.contains('hidden'))sync().then(()=>renderStats(stats()))}});
  function failPin(){
    let n=Number(localStorage.getItem('uvf')||0)+1;
    if(n>=5){lsSet('uvl',Date.now()+60000);n=0;$('loginErr').textContent='Too many wrong tries. Locked for 1 minute.'}
    lsSet('uvf',n);
  }
  function logoutBtn(){
    if(Q.length&&!confirm(Q.length+' sale(s) are still waiting to sync. They stay saved and will sync after you log in again. Logout now?'))return;
    lockApp(); refreshBio();
  }
  /* ===== Cut / Share / Report / Fingerprint / Smart helpers ===== */
  function shareText(t){
    if(window.Android&&Android.shareText)Android.shareText(t);
    else if(navigator.share)navigator.share({text:t}).catch(()=>{});
    else copyText(t);
  }
  function cutUids(){copyText(LAST.join('\n'));$('uids').textContent='';LAST=[];LASTIDS=[];closeModal();$('lastBtn').classList.add('hidden');toast('Cut: copied and cleared from screen')}
  function shareUids(){shareText(LAST.join('\n'))}
  const shareRow=i=>shareText(lines(ROWS[i].uids));
  function shareReport(){
    toggleMenu(false); const s=stats();
    shareText('FB IDs Vault report, '+fmtD(Date.now())+'\nSold today: '+money(s.today)+' IDs, earned '+money(s.earnToday)+'\nThis month: '+money(s.month)+' IDs, earned '+money(s.earnMonth)+'\nTotal sold: '+money(s.sold)+' IDs, earned '+money(s.earnAll)+'\nStock left: '+money(s.available)+' of '+money(s.total));
  }
  let AUTO='';
  function fillRate(){
    const n=($('cust').value||'').trim().toLowerCase(); if(!n||!DB)return;
    let best=null; DB.sales.forEach(x=>{if(x.status!=='Undone'&&String(x.customer).trim().toLowerCase()===n&&Number(x.rate)>0&&(!best||Number(x.date)>Number(best.date)))best=x});
    const r=$('rate'); if(best&&(r.value===''||r.value===AUTO)){r.value=best.rate;AUTO=String(best.rate);updTot()}
  }
  function refreshBio(){
    let ok=false; try{ok=!!(window.Android&&Android.canBio&&Android.canBio()&&localStorage.getItem('uvp')!==null)}catch(e){}
    $('bioBtn').classList.toggle('hidden',!ok); return ok;
  }
  function bioLogin(){ if(window.Android&&Android.bio)Android.bio(); }
  window.onBio=ok=>{ if(!ok)return; $('pin').value=localStorage.getItem('uvp')||''; doLogin({preventDefault(){}}); };
  async function doLogin(e){
    e.preventDefault();
    const wait=Number(localStorage.getItem('uvl')||0)-Date.now();
    if(wait>0){$('loginErr').textContent='Too many wrong tries. Try again in '+Math.ceil(wait/1000)+' sec.';return false}
    PIN=$('pin').value; $('loginBtn').disabled=true; $('loginBtn').textContent='Unlocking...'; $('loginErr').textContent='';
    try{
      const r=await call({action:'login'});
      if(!r.ok){$('loginErr').textContent=r.error; if(/pin/i.test(r.error||''))failPin(); PIN='';}
      else{lsSet('uv',PIN); lsSet('uvt',Date.now()); lsSet('uvf',0); enter(r.stats);}
    }catch(x){$('loginErr').textContent='Error: '+((x&&x.message)||x);}
    $('loginBtn').disabled=false; $('loginBtn').textContent='Unlock'; return false;
  }
  function enter(s){$('login').classList.add('hidden');$('app').classList.remove('hidden');renderStats(s);if(s.low)setTimeout(()=>toast('Low stock: only '+s.available+' IDs left'),700)}
  async function run(req,errId){
    try{const r=await call(req); if(!r.ok){const m=/Unknown action/i.test(r.error||'')?'The sheet script is old. Deploy the latest Code.gs (Manage deployments, Edit, New version).':r.error; errId?($(errId).textContent=m):toast(m); return null;} return r;}
    catch(e){const m='Error: '+((e&&e.message)||e); errId?($(errId).textContent=m):toast(m); return null;}
  }
  function showTab(t){
    toggleMenu(false);
    document.querySelectorAll('[data-tab]').forEach(e=>e.classList.toggle('hidden',e.dataset.tab!==t));
    document.querySelectorAll('#nav button').forEach(b=>b.classList.toggle('on',b.dataset.t===t));
    $('pgT').textContent=PG[t][0]; $('pgS').textContent=PG[t][1]; window.scrollTo(0,0);
    if(t==='orders')loadHistory(); if(t==='cust')loadCust();
  }
  /* ===== Saved customer names ===== */
  let NAMES=[];
  function saveName(n){n=String(n||'').trim();if(!n)return;const l=(LS.get('uvnames')||[]).filter(x=>String(x).toLowerCase()!==n.toLowerCase());l.unshift(n);LS.set('uvnames',l.slice(0,300))}
  function allNames(){
    const seen=new Set(),out=[],add=n=>{n=String(n||'').trim();const k=n.toLowerCase();if(n&&!seen.has(k)){seen.add(k);out.push(n)}};
    if(DB&&DB.sales)DB.sales.slice().sort((a,b)=>(Number(b.date)||0)-(Number(a.date)||0)).forEach(x=>add(x.customer));
    (LS.get('uvnames')||[]).forEach(add); return out;
  }
  function drawNames(){
    const all=allNames(), q=($('cust').value||'').trim().toLowerCase();
    $('custNames').innerHTML=all.slice(0,300).map(n=>'<option value="'+esc(n)+'"></option>').join('');
    NAMES=all.filter(n=>!q||(n.toLowerCase().includes(q)&&n.toLowerCase()!==q)).slice(0,10);
    try{fillRate()}catch(e){}
    $('nchips').innerHTML=NAMES.map((n,i)=>'<button type="button" class="nchip" onclick="pickName('+i+')">'+esc(n)+'</button>').join('');
  }
  function pickName(i){$('cust').value=NAMES[i]||'';drawNames();$('qty').focus()}
  function drawCats(){
    const fillSel=(id,arr,cur,count)=>{
      const el=$(id); if(!el)return;
      if(!el.options.length)arr.forEach(x=>{const o=document.createElement('option');o.value=x[0];el.appendChild(o)});
      Array.from(el.options).forEach((o,i)=>{const t=arr[i][1]+(count?' ('+money(count(arr[i][0]))+')':'');if(o.textContent!==t)o.textContent=t});
      if(el.value!==cur)el.value=cur;
    };
    fillSel('selTyS',[['','Any']].concat(TYPES),SEL.ty,k=>cc(CATS,k,SEL.ge));
    fillSel('selGeS',[['','Any']].concat(GENS),SEL.ge,k=>cc(CATS,SEL.ty,k));
    fillSel('addTyS',[TYPES[3]].concat(TYPES.slice(0,3)),ADDC.ty);
    fillSel('addGeS',[GENS[2]].concat(GENS.slice(0,2)),ADDC.ge);
    $('addInfo').textContent=catName(ADDC.ty,ADDC.ge).replace(', ',' \u00b7 ');
    $('addWarn').classList.toggle('hidden',!TAGS_OLD); try{cnt()}catch(e){}
    $('selInfo').textContent=money(cc(CATS,SEL.ty,SEL.ge))+' IDs available'+((SEL.ty||SEL.ge)?' in this category':'');
    $('cgrid').innerHTML='<tr><th></th>'+GENS.map(g=>'<th>'+g[1]+'</th>').join('')+'<th>Total</th></tr>'+TYPES.map(t=>'<tr><td>'+t[1]+'</td>'+GENS.map(g=>{const n=CATS[t[0]+g[0]]||0;return'<td class="'+(n?'':'z')+'">'+n+'</td>'}).join('')+'<td><b>'+cc(CATS,t[0],'')+'</b></td></tr>').join('')
      +'<tr><td>All</td>'+GENS.map(g=>'<td><b>'+cc(CATS,'',g[0])+'</b></td>').join('')+'<td><b>'+cc(CATS,'','')+'</b></td></tr>';
  }
  const selTy=k=>{SEL.ty=k;drawCats()}, selGe=k=>{SEL.ge=k;drawCats()}, addTy=k=>{ADDC.ty=k;drawCats()}, addGe=k=>{ADDC.ge=k;drawCats()};
  function renderStats(s){
    $('verWarn').classList.toggle('hidden',!s.online||SCRIPT_VER>=APP_VER);
    $('verInfo').textContent='Found: '+(SCRIPT_VER?'script version '+SCRIPT_VER:'no version number')+', needed: '+APP_VER+'. Connected through: '+(netMode()==='page'?'this web page':'the link '+shortUrl(APP_URL))+'.';
    const set=(i,v)=>$(i).textContent=money(v);
    set('sAv',s.available);set('sSold',s.sold);set('sTot',s.total);set('sToday',s.today);set('sMonth',s.month);
    set('eToday',s.earnToday);set('eMonth',s.earnMonth);set('eAll',s.earnAll);set('sbAv',s.available);
    const pc=s.total?Math.round(s.available/s.total*100):0;
    ['meter','meter2'].forEach(i=>{$(i).classList.toggle('low',s.low);$(i).firstElementChild.style.width=pc+'%'});
    $('kMain').classList.toggle('low',s.low); $('heroBar').style.width=pc+'%'; try{drawNames()}catch(e){} try{drawCats()}catch(e){}
    $('stTxt').textContent=money(s.available)+' of '+money(s.total)+' IDs left ('+pc+'%)';
    const p=$('pill'); p.textContent=(s.online?'':'Offline, ')+(s.pending?s.pending+' waiting to sync, ':'')+(s.low?'Low stock: '+s.available+' left':s.available+' IDs ready'); p.classList.toggle('low',s.low||!s.online);
    const mx=Math.max.apply(null,s.last7.map(x=>x.n).concat(1));
    $('chart').innerHTML=s.last7.map(x=>'<div class="bar"><em>'+x.n+'</em><i style="height:'+Math.round(x.n/mx*110)+'px"></i>'+x.l+'</div>').join('');
  }
  async function sell(){
    const c0=$('cust').value.trim().toLowerCase(), q0=Number($('qty').value);
    if(c0&&q0&&DB&&DB.sales.some(x=>x.status!=='Undone'&&String(x.customer).trim().toLowerCase()===c0&&Number(x.qty)===q0&&Date.now()-Number(x.date)<120000)&&!confirm('The same order ('+q0+' IDs) was just given to this customer. Give again?'))return;
    $('sellErr').textContent=''; $('sellBtn').disabled=true; $('sellBtn').textContent='Processing…';
    const customer=$('cust').value.trim();
    const r=await run({action:'sell',customer,qty:$('qty').value,rate:$('rate').value,mode:MODE,ty:SEL.ty,ge:SEL.ge},'sellErr');
    if(r){
      LAST=r.uids.map(fmt); LASTNAME=customer; saveName(customer);
      $('resTitle').textContent=r.uids.length+' IDs for '+customer+((SEL.ty||SEL.ge)?' ['+catName(SEL.ty,SEL.ge)+']':'')+(r.amount?' ('+money(r.amount)+')':'')+(typeof r.order==='number'?', Order #'+r.order:', waiting to sync');
      LASTIDS=r.uids.slice(); LASTMISS=r.uids.filter(u=>!(DB.pw&&DB.pw[u]));
      $('pwBar').classList.toggle('hidden',!LASTMISS.length); $('pwFix').value='';
      if(LASTMISS.length){
        $('resTitle').textContent+=' | '+LASTMISS.length+' IDs have no password';
        $('pwBarNote').textContent=PW_OLD?'The sheet script is old, so passwords cannot be read. Deploy the latest Code.gs (Manage deployments, Edit, New version).':LASTMISS.length+' of these IDs have no password saved in the sheet. Type one below to save it on them.';
        $('pwFixRow').classList.toggle('hidden',PW_OLD); toast(LASTMISS.length+' IDs have no password');
      }
      $('uids').textContent=LAST.join('\n');
      drawResPw(); drawResMeta(customer,r); showModal('resModal'); $('lastBtn').classList.remove('hidden');
      renderStats(r.stats); $('cust').value=''; $('qty').value=''; $('rate').value=''; AUTO=''; updTot(); drawNames(); if(r.stats.low)toast('Low stock: only '+r.stats.available+' IDs left');
    }
    $('sellBtn').disabled=false; $('sellBtn').textContent='Get IDs';
  }
  async function addStock(){
    $('addErr').textContent='';
    const ty=$('addTyS').value||'X', ge=$('addGeS').value||'X', code=ty+ge;   // read the category straight from the dropdowns
    ADDC.ty=ty; ADDC.ge=ge;
    const front=ADDS.pos==='front', pass=ADDS.pw==='same'?$('addPass').value.trim():'', hadPass=ADDS.pw==='line'||!!pass;
    if(ADDS.pw==='same'&&!pass){$('addErr').textContent='Type the password, or choose None';return}
    $('addBtn').disabled=true;
    const r=await run({action:'add',text:$('addTxt').value,cat:code,hasPw:ADDS.pw==='line',pass,front,bad:$('addBad').checked},'addErr');
    if(r){
      $('addTxt').value=''; $('addPass').value=''; cnt(); renderStats(r.stats);
      const ids=r.uids||[], pending=Q.length>0, sk=r.skip||{stock:0,sold:0,bad:0};
      const miss=(!pending&&code!=='XX')?ids.filter(u=>catOf(u)!==code).length:0;
      const L=[];
      L.push('Category: <b>'+esc(catName(ty,ge))+'</b> (Batch in sheet: <b>'+(code==='XX'?'blank':code)+'</b>)');
      L.push('New IDs added: <b>'+money(r.added)+'</b> at the <b>'+(front?'front':'end')+'</b> of the stock');
      if(r.moved)L.push('Category / password updated on IDs already in stock: <b>'+money(r.moved)+'</b>');
      if(sk.stock)L.push('Skipped, already in stock: <b>'+money(sk.stock)+'</b>');
      if(sk.sold)L.push('Skipped, already sold: <b>'+money(sk.sold)+'</b> (use the Returns tab to take them back)');
      if(sk.bad)L.push('Skipped, marked Bad: <b>'+money(sk.bad)+'</b> (tick "re-add Bad IDs" to add them)');
      L.push('Now in this category: <b>'+money(CATS[code]||0)+'</b> IDs');
      L.push('Total stock: <b>'+money(r.stats.available)+'</b> IDs');
      const wp=ids.filter(u=>pwOf(u)), dp=new Set(wp.map(pwOf));
      if(wp.length)L.push('Passwords saved on <b>'+money(wp.length)+'</b> IDs ('+dp.size+' different)');
      else if(hadPass&&!pending)L.push('<span style="color:var(--warn)">No passwords were saved.</span>');
      if(pending)L.push('<span style="color:var(--warn)">No internet. The IDs will reach the sheet after the next sync.</span>');
      if(miss)L.push('<span style="color:var(--bad)"><b>Warning:</b> the category of '+miss+' IDs was not saved in the sheet. Deploy the latest Code.gs (Manage deployments, Edit, New version).</span>');
      const pmiss=(!pending&&ADDS.pw==='same')?ids.filter(u=>!(DB.pw&&DB.pw[u])).length:0;
      if(pmiss)L.push('<span style="color:var(--bad)"><b>Warning:</b> the password of '+pmiss+' IDs was not saved in the sheet. Deploy the latest Code.gs (Manage deployments, Edit, New version).</span>');
      $('amT').textContent=(r.added||r.moved)?'Stock added':'No new IDs were added';
      $('amBody').innerHTML=L.join('<br>');
      showModal('addModal');
    }
    $('addBtn').disabled=false;
  }
  async function changePin(){
    const n=$('newP').value; $('pinErr').textContent='';
    if(!/^\d{4,10}$/.test(n)){$('pinErr').textContent='New PIN must be 4 to 10 digits';return}
    if(n!==$('cnfP').value){$('pinErr').textContent='PINs do not match';return}
    $('pinBtn').disabled=true;
    const r=await run({action:'changepin',cur:$('curP').value,newPin:n},'pinErr');
    $('pinBtn').disabled=false;
    if(r){PIN=n; try{localStorage.setItem('uv',n)}catch(e){} ['curP','newP','cnfP'].forEach(i=>{$(i).value='';const b=$(i).nextElementSibling;$(i).type='password';if(b)b.innerHTML=ic('eye')}); showModal('pinModal');}
  }
  async function loadHistory(){
    $('hist').innerHTML='<div class="skel"></div><div class="skel"></div><div class="skel"></div>';
    const r=await run({action:'history',from:$('fFrom').value,to:$('fTo').value,customer:$('fCust').value}); if(!r)return;
    ROWS=r.history; const box=$('hist');
    if(!ROWS.length){box.innerHTML='<div class="empty">No orders found.</div>';return;}
    box.innerHTML=ROWS.map((h,i)=>'<div class="h-item'+(h.status==='Undone'?' undone':'')+'"><div class="who clk" onclick="openOrderRow('+i+')"><b>'+esc(h.customer)+' '+(h.mode==='Fresh'?'<span class="tag">Fresh</span>':'')+'</b><span>Order #'+h.order+', '+esc(h.date)+(h.amount?', '+money(h.amount):'')+(h.status==='Undone'?', undone':h.status==='Pending'?', waiting to sync':'')+'</span></div><span class="qty">'+h.qty+'</span><button class="btn sm" onclick="openOrderRow('+i+')">'+ic('eye')+'View</button><button class="btn sm ghost" onclick="copyRow('+i+')">'+ic('copy')+'Copy</button><button class="btn sm ghost" onclick="saveRow('+i+')">'+ic('down')+'.txt</button><button class="btn sm ghost" onclick="shareRow('+i+')">'+ic('share')+'Share</button>'+(h.status==='Undone'||h.status==='Pending'?'':'<button class="btn sm ghost" onclick="undo('+h.order+')">'+ic('undo')+'Undo</button>')+'</div>').join('');
  }
  const copyRow=i=>copyText(lines(ROWS[i].uids));
  const saveRow=i=>saveTxt(lines(ROWS[i].uids),ROWS[i].customer+'_order'+ROWS[i].order);
  async function undo(o,fromModal){
    if(!confirm('Undo order #'+o+'? Its IDs will go back to stock.'))return;
    const r=await run({action:'undo',order:o}); if(r){toast('Order #'+o+' undone'); renderStats(r.stats); if(fromModal)closeModal(); loadHistory(); if(!$('fRes').classList.contains('hidden')&&$('fUid').value.trim())findUid();}
  }
  async function loadCust(){
    $('custList').innerHTML='<div class="skel"></div><div class="skel"></div><div class="skel"></div>';
    const r=await run({action:'customers'}); if(!r){$('custList').innerHTML='<div class="empty">Could not load customers.</div>';return} CUSTL=r.list; drawCust();
  }
  let CUSTL=[];
  function drawCust(){
    const q=($('custQ').value||'').trim().toLowerCase(), box=$('custList'), list=CUSTL.filter(c=>!q||String(c.name).toLowerCase().includes(q)); const r={list};
    if(!r.list.length){box.innerHTML='<div class="empty">'+(CUSTL.length?'No customer matches.':'No customers yet.')+'</div>';return;}
    box.innerHTML=r.list.map(c=>'<div class="h-item"><div class="who"><b>'+esc(c.name)+'</b><span>'+c.orders+' orders, last '+esc(c.last)+(c.amount?', '+money(c.amount):'')+'</span></div><span class="qty">'+money(c.qty)+' IDs</span><button class="btn sm ghost" data-n="'+esc(c.name)+'" onclick="custOrders(this.dataset.n)">'+ic('orders')+'Orders</button></div>').join('');
  }
  function custOrders(n){$('fCust').value=n; showTab('orders');}
  let FS=null;
  async function findUid(){
    const uid=$('fUid').value.trim(), box=$('fRes'); const r=await run({action:'find',uid}); if(!r)return;
    FS=null; let h='';
    if(r.where==='sold'){
      const sale=DB.sold[uid]; FS={sale,uid}; const o=saleRow(sale), pend=o.status==='Pending';
      h='<div class="fhead"><span class="tag ok">Sold'+(pend?', waiting to sync':'')+'</span><code>'+esc(uid)+'</code></div>'
       +'<div class="kv">'+kv('Given to',esc(o.customer))+kv('Date &amp; time',esc(o.date))+kv('Order #',esc(o.order))+kv('Password',esc(DB.pw[uid]||'none'))+kv('Order size',money(o.qty)+' IDs')+kv('Rate per ID',money(o.rate))+kv('Order total',money(o.amount))+'</div>'
       +'<div class="facts"><button class="btn sm" onclick="viewFound()">'+ic('orders')+'View full order ('+money(o.qty)+' IDs)</button><button class="btn sm ghost" onclick="replaceUid()">'+ic('swap')+'Replace this ID</button></div>';
    }else if(r.where==='stock'){
      h='<div class="fhead"><span class="tag ok">In stock</span><code>'+esc(uid)+'</code></div>This ID is available and has not been given to any customer.<br>Password: <b>'+esc(DB.pw[uid]||'none')+'</b>';
    }else if(r.where==='bad'){
      const sale=DB.sales.find(x=>String(x.order)===String(r.order)); if(sale)FS={sale,uid};
      h='<div class="fhead"><span class="tag">Marked bad</span><code>'+esc(uid)+'</code></div>'
       +'<div class="kv">'+kv('Was given to',esc(r.customer))+kv('Order #',esc(r.order))+(sale?kv('Date &amp; time',esc(fmtD(sale.date,1))):'')+'</div>'
       +(sale?'<div class="facts"><button class="btn sm" onclick="viewFound()">'+ic('orders')+'View full order</button></div>':'');
    }else h='UID not found in this sheet.';
    box.innerHTML=h; box.classList.remove('hidden');
  }
  function viewFound(){ if(FS)openOrder(saleRow(FS.sale),FS.uid); }
  async function replaceUid(){
    const uid=$('fUid').value.trim();
    if(!confirm('Mark this ID as bad and give the customer a new one?'))return;
    const r=await run({action:'replace',uid}); if(!r)return;
    $('fRes').innerHTML='<b>Replaced</b><br>New ID for '+esc(r.customer)+' (Order #'+r.order+'):<div class="uids" id="newU" style="margin-top:8px">'+esc(fmt(r.newUid))+'</div><br><button class="btn sm" onclick="copyText($(\'newU\').textContent)">'+ic('copy')+'Copy</button>';
    renderStats(r.stats);
  }
  function exportCsv(){
    if(!ROWS.length)return toast('Load orders first');
    const q=v=>'"'+String(v==null?'':v).replace(/"/g,'""')+'"';
    const lines=['Order,Date,Customer,Qty,Rate,Amount,Mode,Status,UIDs'].concat(ROWS.map(h=>[h.order,h.date,h.customer,h.qty,h.rate,h.amount,h.mode,h.status,h.uids.map(fmt).join(' ')].map(q).join(',')));
    saveFile('uid-vault-orders.csv','\ufeff'+lines.join('\n'),'text/csv');
  }
  const copied=t=>{const n=String(t).split('\n').filter(Boolean).length;return n>1?'Copied '+n+' lines':'Copied'};
  function copyText(t){
    if(navigator.clipboard && window.isSecureContext){ navigator.clipboard.writeText(t).then(()=>toast(copied(t))).catch(()=>fallback(t)); }
    else fallback(t);
  }
  function fallback(t){const a=document.createElement('textarea');a.value=t;document.body.appendChild(a);a.select();try{document.execCommand('copy');toast(copied(t))}catch(e){toast('Copy failed, select and copy manually')}a.remove()}
  function copyUids(){copyText(LAST.join('\n'))}
  function saveFile(name,text,mime){
    if(window.Android&&Android.saveFile){Android.saveFile(name,text,mime);return}
    const u=URL.createObjectURL(new Blob([text],{type:mime}));
    const a=document.createElement('a');a.href=u;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),1000);
  }
  function saveTxt(t,name){saveFile((name||'uids').replace(/[^\w-]+/g,'_')+'.txt',t,'text/plain')}
  function downloadUids(){saveTxt(LAST.join('\n'),LASTNAME)}


  /* ===== Connection check ===== */
  async function checkConn(){
    const o=$('chkOut'); o.textContent='Checking...'; $('chkBtn').disabled=true;
    const via=netMode()==='page'?'this web page (Apps Script)':'the link '+shortUrl(APP_URL);
    try{
      const r=await net({action:'about'});
      if(!r||!r.ok){
        o.innerHTML='Connected through: <b>'+esc(via)+'</b><br><span style="color:var(--bad)">'+(/Unknown action/i.test((r&&r.error)||'')?'This script is OLD (it does not know the latest commands). Paste the latest Code.gs and deploy a New version.':esc((r&&r.error)||'No answer'))+'</span>';
      }else{
        o.innerHTML='Connected through: <b>'+esc(via)+'</b><br>Spreadsheet: <b>'+esc(r.name)+'</b><br>Script version: <b>'+r.ver+'</b> '+(r.ver>=APP_VER?'(up to date)':'(old, needs '+APP_VER+')')+'<br>IDs in stock: <b>'+money(r.stock)+'</b>, rows sold: <b>'+money(r.sold)+'</b>, orders: <b>'+money(r.orders)+'</b>';
      }
    }catch(e){o.innerHTML='Connected through: <b>'+esc(via)+'</b><br><span style="color:var(--bad)">Could not reach the script: '+esc((e&&e.message)||e)+'</span>'}
    $('chkBtn').disabled=false;
  }

  /* ===== Passwords: grouped copy helpers ===== */
  const PWP={};
  const pwOf=u=>(DB&&DB.pw&&DB.pw[u])||'';
  function pwGroups(uids){const m=new Map();uids.forEach(u=>{const p=pwOf(u);if(!m.has(p))m.set(p,[]);m.get(p).push(u)});return Array.from(m,([pw,ids])=>({pw,ids}))}
  const copyIdsOnly=a=>copyText(a.join('\n'));
  function copyPassOnly(a){const g=pwGroups(a).filter(x=>x.pw);copyText(g.length===1&&g[0].ids.length===a.length?g[0].pw:a.map(pwOf).join('\n'))}
  function pwCopy(key,i,w){const g=PWP[key][i];copyText(w==='ids'?g.ids.join('\n'):w==='pw'?g.pw:lines(g.ids))}
  function pwPanelHtml(key,uids){
    const g=pwGroups(uids); PWP[key]=g; const withPw=g.filter(x=>x.pw);
    if(!withPw.length)return '';
    const b=(i,w,t)=>'<button type="button" class="mini" onclick="pwCopy(\''+key+'\','+i+',\''+w+'\')">'+t+'</button>';
    if(g.length===1)return '<div class="pwp"><div class="pwrow"><span>Password for all '+money(uids.length)+' IDs</span><code>'+esc(g[0].pw)+'</code><i class="sp"></i>'+b(0,'pw',ic('copy')+'Copy password')+'</div></div>';
    return '<div class="pwp"><div class="sub" style="margin:0 0 6px">'+withPw.length+' different password'+(withPw.length===1?'':'s')+' in this list. Copy by group:</div>'
      +g.slice(0,10).map((x,i)=>'<div class="pwrow">'+(x.pw?'<code>'+esc(x.pw)+'</code>':'<span>No password</span>')+'<span>'+money(x.ids.length)+' IDs</span><i class="sp"></i>'+b(i,'ids','IDs')+(x.pw?b(i,'pw','Pass')+b(i,'both','Both'):'')+'</div>').join('')
      +(g.length>10?'<div class="sub">+ '+(g.length-10)+' more groups. Use Copy IDs + passwords above.</div>':'')+'</div>';
  }
  function drawResMeta(customer,r){
    const num=typeof r.order==='number';
    $('resMeta').innerHTML=kv('Customer',esc(customer))+kv('Order',num?'#'+r.order:'Waiting to sync')+kv('Quantity',money(r.uids.length)+' IDs')+kv('Total',r.amount?money(r.amount):'-')
      +kv('Which IDs',MODE==='fresh'?'Fresh (newest)':'Normal (oldest first)')+kv('Category',(SEL.ty||SEL.ge)?esc(catName(SEL.ty,SEL.ge)):'Any');
  }
  function drawResPw(){
    const any=LASTIDS.some(u=>pwOf(u));
    ['rcIds','rcPw'].forEach(i=>$(i).classList.toggle('hidden',!any));
    $('rcAll').innerHTML=ic('copy')+(any?'Copy IDs + passwords':'Copy all');
    $('resPw').innerHTML=pwPanelHtml('res',LASTIDS);
  }

  /* ===== Save a password on IDs that have none ===== */
  async function fixPass(){
    const pw=$('pwFix').value.trim(); if(!pw){toast('Type a password first');return}
    $('pwFixBtn').disabled=true;
    const r=await run({action:'setpass',uids:LASTMISS,pass:pw});
    $('pwFixBtn').disabled=false; if(!r)return;
    LAST=LASTIDS.map(fmt); $('uids').textContent=LAST.join('\n');
    LASTMISS=LASTIDS.filter(u=>!(DB.pw&&DB.pw[u]));
    $('pwBar').classList.toggle('hidden',!LASTMISS.length); $('pwFix').value='';
    drawResPw(); renderStats(r.stats); toast('Password saved on '+r.saved+' IDs');
  }

  /* ===== Returns & Replacements ===== */
  let RETM='return', RETNEW=[];
  function parseUidList(text){
    const out=[],seen=new Set();
    String(text||'').split(/\r?\n/).forEach(l=>{l=l.trim();if(!l)return;const parts=/[|\t]/.test(l)?[l.split(/[|\t]/)[0].trim()]:l.split(/[\s,;]+/);parts.forEach(u=>{if(u&&!seen.has(u)){seen.add(u);out.push(u)}})});
    return out;
  }
  function cntRet(){$('retCnt').textContent=parseUidList($('retTxt').value).length+' IDs detected'}
  function setRet(m){
    RETM=m; document.querySelectorAll('#retSeg button').forEach(b=>b.classList.toggle('on',b.dataset.m===m));
    $('retSub').textContent=m==='return'
      ?'A customer gave IDs back? Paste them here. They are removed from the order and go back into stock (Normal orders go to the front, Fresh orders to the end). The order amount is reduced.'
      :'IDs that do not log in? Paste them here. Each one is marked Bad and the customer gets a new ID from the same category (Fresh orders get a Fresh ID). The old ID is saved in the Bad sheet.';
    $('retBtn').innerHTML=ic(m==='return'?'undo':'swap')+(m==='return'?'Return to stock':'Replace IDs');
    $('retRes').classList.add('hidden'); $('retErr').textContent='';
  }
  async function runRet(){
    $('retErr').textContent=''; const ids=parseUidList($('retTxt').value);
    if(!ids.length){$('retErr').textContent='Paste at least one UID';return}
    if(!confirm(RETM==='return'?'Return '+ids.length+' ID(s) to stock? They will be removed from the customers\' orders.':'Replace '+ids.length+' ID(s)? Each one is marked Bad and a new ID is given.'))return;
    $('retBtn').disabled=true;
    const r=await run({action:RETM==='return'?'returns':'replacemany',uids:ids},'retErr');
    $('retBtn').disabled=false; if(!r)return;
    renderStats(r.stats); $('retTxt').value=''; cntRet();
    const miss=r.missing||[], cut=a=>a.slice(0,5).join(', ')+(a.length>5?' and '+(a.length-5)+' more':'');
    if(RETM==='return'){
      const by={}; (r.done||[]).forEach(x=>{const k=x.customer+' (Order #'+x.order+')';by[k]=(by[k]||0)+1});
      const L=[]; L.push('<b>'+money((r.done||[]).length)+'</b> IDs are back in stock');
      Object.keys(by).forEach(k=>L.push(esc(k)+': <b>'+by[k]+'</b>'));
      if(miss.length)L.push('<span style="color:var(--warn)">Not found in Sold (never sold, or already returned): <b>'+miss.length+'</b><br>'+esc(cut(miss))+'</span>');
      L.push('Total stock now: <b>'+money(r.stats.available)+'</b> IDs');
      $('amT').textContent=(r.done||[]).length?'IDs returned to stock':'Nothing was returned';
      $('amBody').innerHTML=L.join('<br>'); showModal('addModal'); $('retRes').classList.add('hidden');
    }else{
      const d=r.done||[]; RETNEW=d.map(x=>fmt(x.nu));
      $('retTitle').textContent=d.length+' ID(s) replaced';
      $('retOut').textContent=d.map(x=>x.old+'  ->  '+fmt(x.nu)+'   ('+x.customer+', #'+x.order+')').join('\n');
      const nt=[]; if(miss.length)nt.push('Not found in Sold: '+cut(miss)+'.'); if((r.noStock||[]).length)nt.push('No stock left for: '+cut(r.noStock)+'.');
      $('retNote').textContent=nt.join(' ');
      $('retRes').classList.toggle('hidden',!d.length&&!nt.length);
      if(!d.length)$('retErr').textContent=nt.join(' ')||'Nothing was replaced';
      else toast(d.length+' replaced. Old IDs saved as Bad.');
    }
  }

  /* ===== Modals ===== */
  function showModal(id){hideModals();$(id).classList.remove('hidden');document.body.style.overflow='hidden';try{history.pushState({m:id},'')}catch(e){}}
  function hideModals(){document.querySelectorAll('.modal').forEach(m=>m.classList.add('hidden'));document.body.style.overflow=''}
  function closeModal(){hideModals();try{if(history.state&&history.state.m)history.back()}catch(e){}}
  window.addEventListener('popstate',hideModals);
  document.addEventListener('keydown',e=>{if(e.key==='Escape')closeModal()});

  /* ===== Order detail (click any order, or "View full order" in UID search) ===== */
  const kv=(k,v)=>'<div><span>'+k+'</span><b>'+v+'</b></div>';
  let OD=null;
  function saleRow(x){return{order:x.order,date:typeof x.date==='number'?fmtD(x.date,1):x.date,customer:x.customer,qty:(x.uids||[]).length,rate:x.rate,amount:x.amount,uids:x.uids||[],mode:x.mode,status:x.status}}
  const openOrderRow=i=>openOrder(ROWS[i]);
  function openOrder(o,hl){
    OD=o; const undone=o.status==='Undone', pend=o.status==='Pending';
    $('odTitle').textContent=o.customer; $('odSub').textContent='Order #'+o.order;
    const t=$('odTag'); t.textContent=undone?'Undone':pend?'Waiting to sync':'Sold'; t.className='tag'+(undone||pend?'':' ok');
    $('odMeta').innerHTML=kv('Given on',esc(o.date))+kv('Quantity',money(o.qty)+' IDs')+kv('Rate per ID',money(o.rate))+kv('Total amount',money(o.amount))
      +kv('Type',o.mode==='Fresh'?'Fresh IDs':'Normal (oldest first)')+kv('Status',undone?'Undone, IDs back in stock':pend?'Waiting to sync':'Completed');
    $('odCount').textContent=money(o.uids.length)+' UIDs in this order';
    {const any=o.uids.some(u=>pwOf(u)); ['odIds','odPw'].forEach(i=>$(i).classList.toggle('hidden',!any)); $('odAll').innerHTML=ic('copy')+(any?'Copy IDs + passwords':'Copy all'); $('odPwp').innerHTML=pwPanelHtml('od',o.uids)}
    $('odUids').innerHTML=o.uids.map(u=>u===hl?'<mark id="odHit">'+esc(fmt(u))+'</mark>':esc(fmt(u))).join('\n');
    $('odUndo').classList.toggle('hidden',undone||pend);
    showModal('ordModal'); $('odUids').scrollTop=0;
    const hit=$('odHit'); if(hit)setTimeout(()=>hit.scrollIntoView({block:'center'}),80);
  }

  /* ===== Google Sheet connection manager ===== */
  const shortUrl=u=>String(u).replace(/^https:\/\/script\.google\.com\//,'').replace(/(macros\/s\/[\w-]{6})[\w-]*(\/exec)$/,'$1…$2');
  function drawConnInfo(){
    const c=activeConn();
    ['connNow','connLoginName'].forEach(i=>{const e=$(i); if(e)e.textContent=c.name});
    const u=$('connNowUrl'); if(u)u.textContent=shortUrl(c.url);
  }
  function cmsg(t,ok){const e=$('connMsg'); e.textContent=t; e.className='cmsg'+(ok===true?' ok':ok===false?' bad':'')}
  function openConn(){ $('cName').value=''; $('cUrl').value=''; cmsg(''); renderConn(); showModal('connModal'); }
  function renderConn(){
    const act=activeConn().id;
    $('connList').innerHTML=allConns().map(c=>{
      const on=c.id===act;
      return '<div class="h-item conn'+(on?' act':'')+'"><div class="who"><b>'+esc(c.name)+(on?' <span class="tag ok">Connected</span>':'')+'</b><span>'+esc(shortUrl(c.url))+'</span></div>'
        +(on?'':'<button class="btn sm" data-id="'+esc(c.id)+'" onclick="switchConn(this.dataset.id)">'+ic('link')+'Connect</button>')
        +(c.builtin||on?'':'<button class="btn sm ghost" data-id="'+esc(c.id)+'" onclick="delConn(this.dataset.id)" aria-label="Remove">'+ic('trash')+'</button>')+'</div>';
    }).join('');
  }
  function checkUrl(raw){
    const u=String(raw||'').trim().replace(/\s+/g,'');
    if(!u)return{error:'Paste the Web app link first'};
    if(/docs\.google\.com\/spreadsheets/i.test(u))return{error:'That is the Google Sheet link. You need the Apps Script Web app link that ends with /exec'};
    if(/\/dev(\?.*)?$/.test(u))return{error:'That is a test link (/dev). Use the link from Deploy, New deployment, that ends with /exec'};
    if(!/^https:\/\/script\.google\.com\/(a\/[^\/]+\/)?macros\/s\/[\w-]+\/exec$/.test(u))return{error:'Link must look like https://script.google.com/macros/s/.../exec'};
    return{url:u};
  }
  async function pingUrl(url){
    const c=new AbortController(), t=setTimeout(()=>c.abort(),20000);
    try{
      const r=await fetch(url,{method:'POST',redirect:'follow',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify({action:'ping'}),signal:c.signal});
      const j=JSON.parse(await r.text());
      if(j&&typeof j==='object'&&('ok' in j))return{ok:true};
      return{ok:false,error:'This link did not answer like a FB IDs Vault script'};
    }catch(e){
      return{ok:false,error:(e&&e.name==='AbortError')?'No answer from this link (timeout). Check your internet.':'Could not reach this link. Check internet, that Who has access is Anyone, and that the latest version is deployed.'};
    }finally{clearTimeout(t)}
  }
  async function testConn(){
    const v=checkUrl($('cUrl').value); if(v.error){cmsg(v.error,false);return}
    cmsg('Testing…'); $('cTest').disabled=true;
    const r=await pingUrl(v.url); $('cTest').disabled=false;
    cmsg(r.ok?'Connected. This sheet is ready.':r.error,r.ok);
  }
  async function addConn(){
    const v=checkUrl($('cUrl').value); if(v.error){cmsg(v.error,false);return}
    const ex=allConns().find(c=>c.url===v.url); if(ex){switchConn(ex.id);return}
    cmsg('Testing…'); $('cAdd').disabled=true;
    const r=await pingUrl(v.url); $('cAdd').disabled=false;
    if(!r.ok){cmsg(r.error,false);return}
    const l=savedConns(), id='c'+Date.now().toString(36);
    l.push({id,name:($('cName').value||'').trim()||('Sheet '+(l.length+1)),url:v.url});
    try{localStorage.setItem('uvconns',JSON.stringify(l))}catch(e){cmsg('Phone storage is full',false);return}
    switchConn(id);
  }
  function switchConn(id){
    if(id===activeConn().id){closeModal();return}
    if(Q.length&&!confirm(Q.length+' change(s) have not synced to the current sheet yet. Switching will discard them. Continue?'))return;
    try{CACHE_KEYS.forEach(k=>localStorage.removeItem(k)); localStorage.setItem('uvact',id)}catch(e){}
    try{if(history.state&&history.state.m)history.replaceState(null,'')}catch(e){}
    location.reload();
  }
  function delConn(id){
    const c=savedConns().find(x=>x.id===id); if(!c||id===activeConn().id)return;
    if(!confirm('Remove "'+c.name+'" from this app? The Google Sheet and its data are not deleted.'))return;
    try{localStorage.setItem('uvconns',JSON.stringify(savedConns().filter(x=>x.id!==id)))}catch(e){}
    renderConn();
  }
  drawConnInfo();
  setRet('return'); cnt();
  ['qty','rate'].forEach(i=>$(i).addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();sell()}}));
  $('fUid').addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();findUid()}});

  // Auto refresh every 15 seconds (picks up IDs pasted into the Sheet)
  setInterval(async()=>{
    if(document.hidden||$('app').classList.contains('hidden'))return;
    if(Q.length||Date.now()-LASTPULL>30000)await sync();
    renderStats(stats());
  },15000);

  // Auto sign-in (PIN saved on this device, valid for 24 hours)
  (async function(){
    try{
      if(localStorage.getItem('uv')!==null){
        if(!localStorage.getItem('uvt'))lsSet('uvt',Date.now());
        if(sessionExpired()){lsDel('uv');lsDel('uvt');$('loginErr').textContent='Session expired after 24 hours. Enter your PIN again.';if(refreshBio())setTimeout(bioLogin,500);return}
      }
      PIN=localStorage.getItem('uv')||'';
    }catch(e){}
    refreshBio();
    try{
      const r=await call({action:'login'});
      if(r.ok) enter(r.stats);
      else { PIN=''; lsDel('uv'); lsDel('uvt'); }
    }catch(e){}
  })();
</script>

<script>
/* ===== FB IDs Vault additions: auto-lock, screenshot block, backup, pull to refresh, back button ===== */
(function(){
  const g=id=>document.getElementById(id);
  const lsGet=k=>{try{return localStorage.getItem(k)}catch(e){return null}};
  const lsPut=(k,v)=>{try{localStorage.setItem(k,v)}catch(e){}};
  const appOpen=()=>{const a=g('app');return !!a&&!a.classList.contains('hidden')};

  /* Auto-lock (minutes, 0 = off). Default 5. */
  const alMin=()=>{const v=lsGet('uval');return v===null?5:Number(v)||0};
  window.fbSetAutoLock=v=>{lsPut('uval',String(Number(v)||0));toast(Number(v)?'Auto-lock: '+v+' min':'Auto-lock is off')};
  let hiddenAt=0;
  document.addEventListener('visibilitychange',()=>{
    if(document.hidden){hiddenAt=Date.now();return}
    const m=alMin();
    if(hiddenAt&&m>0&&appOpen()&&Date.now()-hiddenAt>m*60000){hiddenAt=0;lockApp();return}
    hiddenAt=0;
  },true);
  const al=g('alSel'); if(al)al.value=String(alMin());

  /* Screenshot block (Android only) */
  const hasAnd=()=>!!(window.Android&&Android.setSecure);
  if(hasAnd()){const r=g('secRow'),c=g('secChk');if(r&&c){r.classList.remove('hidden');try{c.checked=!!Android.isSecure()}catch(e){}}}
  window.fbSetSecure=on=>{try{Android.setSecure(!!on);toast(on?'Screenshots are blocked':'Screenshots are allowed')}catch(e){}};

  /* Back button (called from the Android app) */
  window.fbBack=function(){
    if(document.querySelector('.modal:not(.hidden)')){closeModal();return true}
    const a=g('app');
    if(a&&a.classList.contains('menu-open')){toggleMenu(false);return true}
    if(appOpen()){const on=document.querySelector('#nav button.on');if(on&&on.dataset.t!=='sell'){showTab('sell');return true}}
    return false;
  };

  /* Backup / restore of links and settings (never the PIN) */
  const KEYS=['uvconns','uvact','uvnames','uval','uvrepo'];
  window.fbBackupExport=function(){
    const data={};KEYS.forEach(k=>{const v=lsGet(k);if(v!==null)data[k]=v});
    const o={app:'FBIDsVault',v:1,date:new Date().toISOString(),data};
    const d=new Date(),p=n=>String(n).padStart(2,'0');
    saveFile('FBIDsVault-backup-'+d.getFullYear()+p(d.getMonth()+1)+p(d.getDate())+'.json',JSON.stringify(o,null,1),'application/json');
  };
  window.fbBackupImport=function(inp){
    const m=g('bkMsg'),f=inp.files&&inp.files[0];m.textContent='';m.style.color='';
    if(!f)return;
    const r=new FileReader();
    r.onload=()=>{
      try{
        const o=JSON.parse(String(r.result));
        if(!o||o.app!=='FBIDsVault'||!o.data||typeof o.data!=='object')throw 0;
        let n=0;KEYS.forEach(k=>{if(typeof o.data[k]==='string'){lsPut(k,o.data[k]);n++}});
        m.style.color='var(--ok)';m.textContent='Backup restored. Reloading...';
        setTimeout(()=>location.reload(),700);
      }catch(e){m.textContent='This file is not an FB IDs Vault backup.'}
      inp.value='';
    };
    r.onerror=()=>{m.textContent='Could not read the file.'};
    r.readAsText(f);
  };

  /* Pull down to refresh (top of the page only) */
  let y0=null,x0=0,t0=null;
  document.addEventListener('touchstart',e=>{
    if(!appOpen()||window.scrollY>0||document.querySelector('.modal:not(.hidden)')){y0=null;return}
    const t=e.touches[0];y0=t.clientY;x0=t.clientX;t0=e.target;
  },{passive:true});
  document.addEventListener('touchend',e=>{
    if(y0===null)return;const t=e.changedTouches[0],dy=t.clientY-y0,dx=Math.abs(t.clientX-x0);y0=null;
    if(dy>120&&dx<60&&window.scrollY<=0&&!(t0&&t0.closest&&t0.closest('textarea,input,select,.uids,.nchips,.bars'))){try{refreshNow()}catch(er){}}
  },{passive:true});
})();
</script>

<script>
/* ===== FB IDs Vault: background loading + update check (GitHub Releases) ===== */
(function(){
  /* >>> Set this to your GitHub "username/repository" (the repository must be public) <<< */
  const DEFAULT_REPO='YOUR_GITHUB_USERNAME/FB-IDs-Vault';

  const g=id=>document.getElementById(id);
  const lsGet=k=>{try{return localStorage.getItem(k)}catch(e){return null}};
  const lsPut=(k,v)=>{try{localStorage.setItem(k,v)}catch(e){}};
  const appOpen=()=>{const a=g('app');return !!a&&!a.classList.contains('hidden')};
  const hasAnd=()=>!!(window.Android&&Android.appVersion);

  /* --- Load everything in the background so every tab opens instantly --- */
  function warm(){setTimeout(()=>{try{loadHistory()}catch(e){}try{loadCust()}catch(e){}},600)}
  let stopped=false;
  setInterval(()=>{
    if(stopped||!appOpen()||document.hidden||(typeof BUSY!=='undefined'&&BUSY>0))return;
    sync().then(r=>{
      if(r&&r.err){stopped=true;return}          // wrong PIN or script problem: stop quietly, the normal sync will report it
      try{renderStats(stats())}catch(e){} warm();
    }).catch(()=>{});
  },90000);
  document.addEventListener('visibilitychange',()=>{if(!document.hidden)stopped=false});

  /* --- Update check --- */
  const nums=v=>String(v||'').replace(/^[^0-9]*/,'').split('.').map(x=>parseInt(x,10)||0);
  function newer(a,b){const x=nums(a),y=nums(b);for(let i=0;i<Math.max(x.length,y.length);i++){const d=(x[i]||0)-(y[i]||0);if(d)return d>0}return false}
  const getRepo=()=>(lsGet('uvrepo')||DEFAULT_REPO).trim();
  const repoSet=()=>/^[\w.-]+\/[\w.-]+$/.test(getRepo())&&!/^YOUR_/.test(getRepo());
  window.fbSetRepo=v=>{v=String(v||'').trim().replace(/^https?:\/\/github\.com\//i,'').replace(/\.git$/i,'').replace(/\/$/,'');lsPut('uvrepo',v);g('repoIn').value=v;lsPut('uvupd','0');toast(repoSet()?'Update source saved':'Use the form username/repo')};

  async function latest(){
    const c=new AbortController(),t=setTimeout(()=>c.abort(),12000);
    try{
      const r=await fetch('https://api.github.com/repos/'+getRepo()+'/releases/latest',{headers:{Accept:'application/vnd.github+json'},signal:c.signal});
      if(r.status===404)return{none:true};
      if(!r.ok)throw 0;
      const j=await r.json();
      const apk=(j.assets||[]).find(a=>/\.apk$/i.test(a.name));
      return{ver:String(j.tag_name||''),url:apk?apk.browser_download_url:j.html_url,page:j.html_url,notes:String(j.body||'').slice(0,300)};
    }finally{clearTimeout(t)}
  }
  function banner(u){
    const old=g('updBar');if(old)old.remove();
    const m=document.querySelector('main');if(!m)return;
    const d=document.createElement('div');d.id='updBar';d.className='verwarn updbar';
    const t=document.createElement('div');t.className='t';
    const b=document.createElement('b');b.textContent='New version '+u.ver.replace(/^v/i,'')+' is ready.';
    t.appendChild(b);t.appendChild(document.createTextNode(' Download it and install over the current app. Your data stays in your Google Sheet.'));
    const go=document.createElement('button');go.className='btn';go.textContent='Update';go.onclick=()=>{try{Android.openUrl(u.url)}catch(e){}};
    const x=document.createElement('button');x.className='x';x.textContent='Later';x.onclick=()=>{lsPut('uvskip',u.ver);d.remove()};
    d.appendChild(t);d.appendChild(go);d.appendChild(x);m.insertBefore(d,m.firstChild);
  }
  window.fbCheckUpdate=async function(manual){
    if(!hasAnd())return;
    const msg=g('updMsg'),btn=g('updBtn');let cur='';try{cur=Android.appVersion()}catch(e){}
    if(manual){msg.style.color='';msg.textContent='Checking...';btn.disabled=true}
    try{
      if(!repoSet()){if(manual)msg.textContent='Enter your GitHub repository above (username/repo) first.';return}
      const u=await latest();
      lsPut('uvupd',String(Date.now()));
      if(u.none||!u.ver){if(manual)msg.textContent='No release published yet.';return}
      if(newer(u.ver,cur)){
        if(manual||lsGet('uvskip')!==u.ver)banner(u);
        if(manual){msg.style.color='var(--brand)';msg.textContent='Version '+u.ver.replace(/^v/i,'')+' is available.'}
      }else if(manual){msg.style.color='var(--ok)';msg.textContent='You have the latest version ('+cur+').'}
    }catch(e){if(manual){msg.style.color='';msg.textContent='Could not check. Check your internet.'}}
    finally{if(manual)btn.disabled=false}
  };
  if(hasAnd()){
    const c=g('updCard');if(c)c.classList.remove('hidden');
    try{g('verNow').textContent=Android.appVersion()}catch(e){}
    {const r=lsGet('uvrepo');if(r)g('repoIn').value=r}
  }
  function autoUpd(){
    if(!hasAnd()||!repoSet())return;
    if(Date.now()-Number(lsGet('uvupd')||0)<12*3600*1000)return;
    setTimeout(()=>fbCheckUpdate(false),4000);
  }

  /* run when the app opens (after login / auto sign-in) */
  const _enter=window.enter;
  if(typeof _enter==='function'){window.enter=function(){const r=_enter.apply(this,arguments);try{warm();autoUpd()}catch(e){}return r}}
})();
</script>
</body>
</html>`;
