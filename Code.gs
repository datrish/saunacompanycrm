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
  wixId: 'Wix ID', lastActivity: 'Last activity', history: 'History'
};
const DATE_FIELDS = ['date', 'lastContact', 'nextFollowUp', 'lastChased'];
const EDITABLE = ['name', 'email', 'phone', 'postcode', 'enquiry', 'source', 'status', 'quote', 'priority', 'owner',
  'nextFollowUp', 'lostReason', 'notes', 'chases', 'lastChased', 'lastContact'];
const STATUSES = ['New', 'Phoned - no answer', 'Replied - no quote', 'Quote sent', 'Chased', 'Won', 'Lost'];
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
    else if (k === 'history') { try { v = v ? JSON.parse(v) : []; } catch (e) { v = []; } }
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
    if (action && (action.type === 'status' || action.type === 'chase') && map.lastContact != null) row[map.lastContact] = today;
    if (map.history != null) {
      let hist = [];
      try { hist = row[map.history] ? JSON.parse(row[map.history]) : []; } catch (e) { hist = []; }
      if (action) hist.push(Object.assign({ at: now.toISOString(), by: Session.getActiveUser().getEmail() || '' }, action));
      row[map.history] = JSON.stringify(hist.slice(-40));
    }
    if (map.lastActivity != null) row[map.lastActivity] = now.toISOString();
    sh.getRange(r, 1, 1, width).setValues([row]);
    return rowToLead_(sh.getRange(r, 1, 1, width).getValues()[0], map);
  } finally {
    lock.releaseLock();
  }
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
  SpreadsheetApp.getActive().toast('Lead board setup done. Now Deploy > New deployment > Web app.');
}
