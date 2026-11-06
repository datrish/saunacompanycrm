/**
 * Sauna Company CRM — lead board.
 * Google Apps Script bound to the "Sauna Lead Tracker" Google Sheet.
 * The sheet tab "Leads" is the database; this script serves the board (Index.html)
 * and reads/writes rows. Deploy as a web app: Execute as "User accessing the web app",
 * Who has access "Anyone with Google account". Only people the sheet is shared with
 * (as Editor) can use the board.
 */

const SHEET_NAME = 'Leads';
const TZ = 'Europe/London';

// field key -> sheet header. Extra headers are added on the right by setup().
const FIELDS = {
  id: 'Lead ID', date: 'Date', source: 'Source', name: 'Name', email: 'Email', phone: 'Phone',
  postcode: 'Postcode', enquiry: 'Enquiry', status: 'Status', quote: 'Quote £ (ex VAT)',
  priority: 'Priority', owner: 'Owner', lastContact: 'Last contact', nextFollowUp: 'Next follow-up',
  lostReason: 'Lost reason', notes: 'Notes', chases: 'Chases', lastChased: 'Last chased',
  wixId: 'Wix ID', lastActivity: 'Last activity', history: 'History', drawingSent: 'Drawing sent', payments: 'Payments'
};
const DATE_FIELDS = ['date', 'lastContact', 'nextFollowUp', 'lastChased', 'drawingSent'];
const EDITABLE = ['name', 'email', 'phone', 'postcode', 'enquiry', 'source', 'status', 'quote', 'priority', 'owner',
  'nextFollowUp', 'lostReason', 'notes', 'chases', 'lastChased', 'lastContact', 'drawingSent', 'payments'];
const STATUSES = ['New', 'Phoned - no answer', 'Replied - no quote', 'Quote sent', 'Chased', 'Won', 'Lost'];
const LOST_REASONS = ['Quote too high', 'Not going ahead', 'Postponed', 'Went with someone else', 'We were too slow', 'No response', 'Not a fit', 'Other'];
const OLD_STATUS = { 'Not contacted': 'New', 'Contacted - no quote': 'Replied - no quote' };
const STATUS_COLOURS = {
  'New': '#F19C9C', 'Phoned - no answer': '#F8C9B5', 'Replied - no quote': '#FCE3A6', 'Quote sent': '#CFE2F3',
  'Chased': '#DCD6F7', 'Won': '#B7E1C4', 'Lost': '#E0E0E0'
};

/* ---------------- web app ---------------- */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Sauna Leads Board')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function sheet_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('No "' + SHEET_NAME + '" tab in this spreadsheet.');
  return sh;
}

function headerMap_(sh) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  const map = {};
  for (const [k, h] of Object.entries(FIELDS)) {
    const i = headers.indexOf(h);
    if (i >= 0) map[k] = i;
  }
  return { map, width: headers.length };
}

function fmtDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return v ? String(v) : '';
}

function rowToLead_(row, map) {
  const l = {};
  for (const [k, i] of Object.entries(map)) {
    let v = row[i];
    if (DATE_FIELDS.includes(k)) v = fmtDate_(v);
    else if (k === 'quote') v = (v === '' || v == null) ? null : Number(String(v).replace(/[£,\s]/g, '')) || null;
    else if (k === 'chases') v = Number(v) || 0;
    else if (k === 'history' || k === 'payments') { try { v = v ? JSON.parse(v) : []; } catch (e) { v = []; } }
    else if (k === 'lastActivity') v = v instanceof Date ? v.toISOString() : String(v || '');
    else v = v == null ? '' : String(v);
    l[k] = v;
  }
  if (OLD_STATUS[l.status]) l.status = OLD_STATUS[l.status];
  if (!STATUSES.includes(l.status)) l.status = l.status || 'New';
  if (!l.lastActivity) l.lastActivity = (l.lastContact || l.date || '') ? (l.lastContact || l.date) + 'T09:00:00Z' : '';
  return l;
}

/** Called by the page on load and every minute. */
function getBoard() {
  const sh = sheet_();
  const { map } = headerMap_(sh);
  const last = sh.getLastRow();
  const rows = last > 1 ? sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues() : [];
  const leads = rows.filter((r) => r[map.id]).map((r) => rowToLead_(r, map));
  return { me: Session.getActiveUser().getEmail() || '', leads, at: new Date().toISOString() };
}

function findRow_(sh, map, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, map.id + 1, last - 1, 1).getDisplayValues();
  for (let i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  return -1;
}

function writeValue_(k, v) {
  if (k === 'payments') return JSON.stringify(Array.isArray(v) ? v : []);
  if (v == null) return '';
  if (k === 'quote') return v === '' ? '' : Number(v);
  if (k === 'chases') return Number(v) || 0;
  return v;
}

/**
 * Update one lead. patch = {field: value}; action = {type, from, to, field} for the history log.
 * Returns the updated lead.
 */
function updateLead(id, patch, action) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const { map } = headerMap_(sh);
    const r = findRow_(sh, map, id);
    if (r < 0) throw new Error('Lead ' + id + ' not found. Someone may have deleted the row.');
    const width = sh.getLastColumn();
    const row = sh.getRange(r, 1, 1, width).getValues()[0];
    const now = new Date();
    const today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
    for (const [k, v] of Object.entries(patch || {})) {
      if (!EDITABLE.includes(k) || map[k] == null) continue;
      if (k === 'status' && !STATUSES.includes(v)) throw new Error('Unknown status ' + v);
      row[map[k]] = writeValue_(k, v);
    }
    if (action && ['status', 'chase', 'call', 'email', 'drawing', 'visit'].includes(action.type) && map.lastContact != null) row[map.lastContact] = today;
    if (map.history != null) {
      let hist = [];
      try { hist = row[map.history] ? JSON.parse(row[map.history]) : []; } catch (e) { hist = []; }
      if (action) hist.push(Object.assign({ at: now.toISOString(), by: Session.getActiveUser().getEmail() || '' }, action));
      row[map.history] = JSON.stringify(hist.slice(-40));
    }
    if (map.lastActivity != null) row[map.lastActivity] = now.toISOString();
    sh.getRange(r, 1, 1, width).setValues([row]);
    const out = rowToLead_(sh.getRange(r, 1, 1, width).getValues()[0], map);
    if (patch && ('payments' in patch || 'status' in patch || 'quote' in patch)) {
      try { refreshPaymentsTab(); } catch (e) { console.warn('Payments tab not refreshed: ' + e.message); }
    }
    return out;
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- Payments tab (a readable copy of every payment stage) ---------------- */

/** Rebuilds the "Payments" tab from the board. Runs by itself after every payment change; can also be run from the editor. */
function refreshPaymentsTab() {
  const ss = SpreadsheetApp.getActive();
  const leads = getBoard().leads.filter((l) => l.status === 'Won' || (Array.isArray(l.payments) && l.payments.length));
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const head = ['Lead ID', 'Client', 'Status', 'Stage', 'Amount £ (ex VAT)', 'Amount £ (inc VAT)', 'Invoice ref', 'Invoice sent', 'Due', 'Payment received', 'Payment status', 'Days late', 'Times chased', 'Last chased'];
  const rows = [];
  leads.forEach((l) => {
    const pays = Array.isArray(l.payments) && l.payments.length ? l.payments : defaultStages_(l.quote || 0);
    pays.forEach((p) => {
      const amt = Number(p.amount) || 0;
      const late = p.invSent && !p.paid && p.due && p.due < today ? Math.round((new Date(today) - new Date(p.due)) / 864e5) : '';
      const st = p.paid ? 'Paid' : p.invSent ? (late ? 'Late' : 'Invoiced, awaiting payment') : 'Not invoiced yet';
      rows.push([l.id, l.name || '', l.status, p.stage || '', amt, Math.round(amt * 1.2), p.invRef || '', p.invSent || '', p.due || '', p.paid || '', st, late, p.chased || '', p.lastChased || '']);
    });
  });
  const order = { 'Late': 0, 'Invoiced, awaiting payment': 1, 'Not invoiced yet': 2, 'Paid': 3 };
  rows.sort((a, b) => order[a[10]] - order[b[10]] || String(a[8] || '9').localeCompare(String(b[8] || '9')));
  let sh = ss.getSheetByName('Payments');
  if (!sh) { sh = ss.insertSheet('Payments', 1); sh.setFrozenRows(1); }
  sh.clearContents();
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  if (rows.length) {
    sh.getRange(2, 1, rows.length, head.length).setValues(rows);
    sh.getRange(2, 5, rows.length, 2).setNumberFormat('£#,##0');
  }
  const note = rows.length + 2;
  sh.getRange(note, 1).setValue('Updated automatically from the board (' + Utilities.formatDate(new Date(), TZ, 'd MMM yyyy HH:mm') + '). Log invoices and payments on the board, not here.').setFontStyle('italic');
}
function defaultStages_(q) {
  const plan = [['Advance 25% (on order)', 0.25], ['40% – frame completed', 0.40], ['30% – after cladding', 0.30], ['Final 5% (handover)', 0.05]];
  let left = q;
  return plan.map(([stage, pct], i) => { const amount = i === plan.length - 1 ? left : Math.round(q * pct); left -= amount; return { stage, amount }; });
}

/** Add a lead from the page (phone calls, info@ emails). Inserted at the top (newest first). */
function addLead(lead) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const { map, width } = headerMap_(sh);
    const id = nextId_(sh, map);
    const now = new Date();
    const today = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
    const tomorrow = Utilities.formatDate(new Date(now.getTime() + 86400000), TZ, 'yyyy-MM-dd');
    const l = Object.assign({ status: 'New', priority: 'Medium', source: 'Phone / other', chases: 0 }, lead, {
      id, date: today, nextFollowUp: (lead && lead.nextFollowUp) || tomorrow, lastActivity: now.toISOString(),
      history: JSON.stringify([{ at: now.toISOString(), by: Session.getActiveUser().getEmail() || '', type: 'created' }])
    });
    const row = new Array(width).fill('');
    for (const [k, i] of Object.entries(map)) if (l[k] != null) row[i] = writeValue_(k, l[k]);
    sh.insertRowBefore(2);
    sh.getRange(2, 1, 1, width).setValues([row]);
    return rowToLead_(sh.getRange(2, 1, 1, width).getValues()[0], map);
  } finally {
    lock.releaseLock();
  }
}

function nextId_(sh, map) {
  const last = sh.getLastRow();
  let max = 0;
  if (last > 1) sh.getRange(2, map.id + 1, last - 1, 1).getDisplayValues()
    .forEach(([v]) => { const n = Number(String(v).replace(/\D/g, '')); if (n > max) max = n; });
  return 'L' + String(max + 1).padStart(3, '0');
}

/* ---------------- one-time setup (run from the editor) ---------------- */

/**
 * Run once: adds the board's extra columns, renames old status labels,
 * updates the Status drop-down and the row colours.
 */
function setup() {
  const sh = sheet_();
  // 1. extra columns
  let headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
  for (const h of Object.values(FIELDS)) {
    if (!headers.includes(h)) {
      sh.getRange(1, headers.length + 1).setValue(h).setFontWeight('bold');
      headers.push(h);
    }
  }
  const { map } = headerMap_(sh);
  const last = Math.max(sh.getLastRow(), 2);
  // 2. drop-down first (the old one would reject the new labels)
  const dv = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build();
  sh.getRange(2, map.status + 1, sh.getMaxRows() - 1, 1).setDataValidation(dv);
  // 3. status labels
  const st = sh.getRange(2, map.status + 1, last - 1, 1);
  st.setValues(st.getValues().map(([v]) => [OLD_STATUS[v] || v]));
  sh.getRange(2, map.chases + 1, last - 1, 1).setValues(
    sh.getRange(2, map.chases + 1, last - 1, 1).getValues().map(([v]) => [v === '' ? 0 : v]));
  // 4. row colours by status
  const col = String.fromCharCode(65 + map.status);
  const full = sh.getRange(2, 1, sh.getMaxRows() - 1, sh.getLastColumn());
  const rules = Object.entries(STATUS_COLOURS).map(([s, c]) => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$' + col + '2="' + s + '"').setBackground(c).setRanges([full]).build());
  sh.setConditionalFormatRules(rules);
  // 5. keep the long History column narrow
  sh.setColumnWidth(map.history + 1, 120);
  // 6. Lost reason drop-down
  sh.getRange(2, map.lostReason + 1, sh.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(LOST_REASONS, true).setAllowInvalid(true).build());
  // 7. readable Payments tab
  refreshPaymentsTab();
  SpreadsheetApp.getActive().toast('Lead board setup done. Now Deploy > New deployment > Web app.');
}

/* ---------------- daily & weekly email reports ---------------- */

const REPORT_TO = 'datrish@gmail.com';
const BOARD_URL = 'https://saunaleads.vercel.app';
const PEOPLE = { 'datrish@gmail.com': 'Deepti', 'hemantatrish@gmail.com': 'Hemant', 'spaziolondondesigners@gmail.com': 'Kiana' };

function day_(d) { return Utilities.formatDate(d instanceof Date ? d : new Date(d), TZ, 'yyyy-MM-dd'); }
function addDays_(ymd, n) { const d = new Date(ymd + 'T12:00:00Z'); d.setUTCDate(d.getUTCDate() + n); return day_(d); }
function who_(by) { if (!by) return 'Automatic'; const k = String(by).toLowerCase(); return PEOPLE[k] || (k.indexOf('@') > 0 ? k.split('@')[0] : String(by)); }
function gbp_(v) { return '£' + Math.round(v || 0).toLocaleString('en-GB'); }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

function kindOf_(e) {
  if (!e) return null;
  if (e.type === 'paychase') return 'paychase';
  if (['call', 'email', 'drawing', 'chase', 'visit', 'invoice', 'paid'].indexOf(e.type) >= 0) return e.type;
  if (e.type !== 'status') return null;
  if (e.to === 'Quote sent') return 'quote';
  if (e.to === 'Won') return 'won';
  if (e.to === 'Lost') return 'lost';
  if (e.to === 'Phoned - no answer') return 'call';
  if (e.to === 'Replied - no quote') return 'email';
  return null;
}

const REPORT_COLS = [
  ['call', 'Calls'], ['noAnswer', '…no answer'], ['email', 'Emails'], ['drawing', 'Drawings'], ['revision', '…revisions'],
  ['visit', 'Site visits'], ['quote', 'Quotes sent'], ['quoteValue', 'Quoted £'], ['chase', 'Chases'],
  ['won', 'New clients (won)'], ['wonValue', 'Won £'], ['lost', 'Lost'], ['invoice', 'Invoices sent'], ['paychase', 'Payment chases'], ['paidValue', 'Payments in £']
];
const LABEL = { call: 'Call', email: 'Email', drawing: 'Drawing sent', chase: 'Chased', visit: 'Site visit', quote: 'Quote sent', won: 'WON', lost: 'Lost', invoice: 'Invoice sent', paid: 'Payment received', paychase: 'Payment chased' };

function buildReport_(from, to, title) {
  const leads = getBoard().leads;
  const blank = () => { const o = {}; REPORT_COLS.forEach(([k]) => { o[k] = 0; }); o.paid = 0; return o; };
  const people = {}; const items = [];
  leads.forEach((l) => (l.history || []).forEach((e) => {
    const k = kindOf_(e); if (!k || !e.at) return;
    const d = day_(e.at); if (d < from || d > to) return;
    const w = who_(e.by); const p = people[w] = people[w] || blank();
    p[k]++;
    if (k === 'call' && (e.outcome === 'no answer' || e.to === 'Phoned - no answer')) p.noAnswer++;
    if (k === 'drawing' && (e.rev || 1) > 1) p.revision++;
    if (k === 'quote') p.quoteValue += l.quote || 0;
    if (k === 'won') p.wonValue += l.quote || 0;
    if (k === 'paid') p.paidValue += Number(e.amount) || 0;
    items.push({ at: e.at, w, k, l, e });
  }));
  const newLeads = leads.filter((l) => l.date >= from && l.date <= to);
  const bySource = {}; newLeads.forEach((l) => { bySource[l.source || 'Other'] = (bySource[l.source || 'Other'] || 0) + 1; });
  const open = ['New', 'Phoned - no answer', 'Replied - no quote', 'Quote sent', 'Chased'];
  const notContacted = leads.filter((l) => l.status === 'New');
  const dueTomorrow = leads.filter((l) => open.indexOf(l.status) >= 0 && l.nextFollowUp && l.nextFollowUp <= addDays_(to, 1));
  const names = Object.keys(people).sort();

  const th = 'style="text-align:right;padding:6px 8px;border-bottom:1px solid #ddd;font-size:12px;color:#555"';
  const td = 'style="text-align:right;padding:6px 8px;border-bottom:1px solid #eee"';
  const cell = (k, v) => (k === 'quoteValue' || k === 'wonValue' || k === 'paidValue') ? (v ? gbp_(v) : '–') : (v || '–');
  let html = '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1d2328;max-width:900px">';
  html += '<h2 style="margin:0 0 4px">' + esc_(title) + '</h2>';
  html += '<p style="margin:0 0 14px;color:#5f676e">' + (from === to ? from : from + ' to ' + to) + ' · <a href="' + BOARD_URL + '">Open the board</a></p>';
  html += '<p><b>New enquiries:</b> ' + newLeads.length + (newLeads.length ? ' (' + Object.keys(bySource).map((s) => esc_(s) + ' ' + bySource[s]).join(', ') + ')' : '') + '</p>';
  if (names.length) {
    html += '<table style="border-collapse:collapse;font-size:13px"><tr><th style="text-align:left;padding:6px 8px;border-bottom:1px solid #ddd;font-size:12px;color:#555">Who</th>' + REPORT_COLS.map(([, lbl]) => '<th ' + th + '>' + lbl + '</th>').join('') + '</tr>';
    names.forEach((n) => { html += '<tr><td style="padding:6px 8px;border-bottom:1px solid #eee"><b>' + esc_(n) + '</b></td>' + REPORT_COLS.map(([k]) => '<td ' + td + '>' + cell(k, people[n][k]) + '</td>').join('') + '</tr>'; });
    html += '</table>';
  } else {
    html += '<p style="padding:10px;background:#fbe4e1;border-radius:6px"><b>Nothing was logged on the board in this period.</b></p>';
  }
  const lostWhy = {}; items.filter((it) => it.k === 'lost').forEach((it) => { const r = it.e.reason || it.l.lostReason || 'No reason given'; lostWhy[r] = (lostWhy[r] || 0) + 1; });
  if (Object.keys(lostWhy).length) html += '<p><b>Why we lost:</b> ' + Object.keys(lostWhy).sort((a, b) => lostWhy[b] - lostWhy[a]).map((r) => esc_(r) + ' ' + lostWhy[r]).join(' · ') + '</p>';
  html += '<p style="margin-top:16px"><b>Still waiting:</b> ' + notContacted.length + ' new leads not contacted yet · ' + dueTomorrow.length + ' follow-ups due by tomorrow.</p>';
  if (items.length) {
    items.sort((a, b) => String(a.at).localeCompare(String(b.at)));
    html += '<h3 style="margin:18px 0 6px">What was done</h3><table style="border-collapse:collapse;font-size:13px">';
    items.slice(-150).forEach((it) => {
      const t = Utilities.formatDate(new Date(it.at), TZ, from === to ? 'HH:mm' : 'EEE d MMM HH:mm');
      let what = LABEL[it.k] || it.k;
      if (it.k === 'call') what += (it.e.outcome === 'no answer' || it.e.to === 'Phoned - no answer') ? ' (no answer)' : ' (spoke)';
      if (it.k === 'drawing' && (it.e.rev || 1) > 1) what += ' (rev ' + it.e.rev + ')';
      if ((it.k === 'quote' || it.k === 'won') && it.l.quote) what += ' ' + gbp_(it.l.quote);
      if ((it.k === 'invoice' || it.k === 'paid' || it.k === 'paychase') && (it.e.stage || it.e.amount)) what += ': ' + esc_(it.e.stage || '') + ' ' + gbp_(it.e.amount);
      if (it.k === 'lost') what += ' – ' + esc_(it.e.reason || it.l.lostReason || 'no reason given') + (it.e.detail ? ': ' + esc_(it.e.detail) : '');
      html += '<tr><td style="padding:4px 8px;color:#5f676e;white-space:nowrap">' + t + '</td><td style="padding:4px 8px">' + esc_(it.w) + '</td><td style="padding:4px 8px"><b>' + what + '</b></td><td style="padding:4px 8px">' + esc_(it.l.id + ' ' + (it.l.name || '')) + '</td></tr>';
    });
    html += '</table>';
  }
  html += '<p style="margin-top:18px;color:#8a8f94;font-size:12px">Sent automatically by the Sauna Leads Board. Counts only what was tapped on the board.</p></div>';
  return { html, total: items.length };
}

/** Weekdays at ~18:30 London: today's activity. */
function sendDailyReport() {
  const today = day_(new Date());
  const dow = Number(Utilities.formatDate(new Date(), TZ, 'u')); // 1 = Mon … 7 = Sun
  if (dow >= 6) return;
  const r = buildReport_(today, today, 'Sauna leads – daily report');
  MailApp.sendEmail({ to: REPORT_TO, subject: 'Sauna leads – daily report ' + Utilities.formatDate(new Date(), TZ, 'EEE d MMM') + ' (' + r.total + ' actions)', htmlBody: r.html, name: 'Sauna Leads Board' });
}

/** Fridays ~19:00 London: Monday to today. */
function sendWeeklyReport() {
  const today = day_(new Date());
  const dow = Number(Utilities.formatDate(new Date(), TZ, 'u'));
  const monday = addDays_(today, -(dow - 1));
  const r = buildReport_(monday, today, 'Sauna leads – weekly report');
  MailApp.sendEmail({ to: REPORT_TO, subject: 'Sauna leads – weekly report w/c ' + monday + ' (' + r.total + ' actions)', htmlBody: r.html, name: 'Sauna Leads Board' });
}

/** Run once from the editor to switch the reports on. */
function installReportTriggers() {
  ScriptApp.getProjectTriggers().forEach((t) => {
    if (['sendDailyReport', 'sendWeeklyReport'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDailyReport').timeBased().everyDays(1).atHour(18).nearMinute(30).inTimezone(TZ).create();
  ScriptApp.newTrigger('sendWeeklyReport').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(19).inTimezone(TZ).create();
}
