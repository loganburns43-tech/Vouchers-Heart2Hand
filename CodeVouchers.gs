/***** CLIENT FETCH FOR HTML MENU *****/
function getAllClients() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Donor Summary');
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const idxL = h.findIndex(x => /last/i.test(x));
  const idxF = h.findIndex(x => /first/i.test(x));
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const l = String(data[i][idxL] || '').trim();
    const f = String(data[i][idxF] || '').trim();
    if (l && f) list.push({ last: l, first: f });
  }
  return list.sort((a, b) => a.last.localeCompare(b.last));
}

/***** CONFIG *****/
const RESPONSES_SHEET    = 'Form responses 1';
const RESPONSES_MERGED   = 'Form responses (Merged)';
const DONOR_SUMMARY      = 'Donor Summary';
const PROGRAM_SUMMARY    = 'Program Summary';
const FORMLOG            = 'FormLog';
const RAWFORMLOG         = 'RawFormLog';
const BUBBLE_FACES       = [5, 10, 15, 20];
const BUBBLE_CREDIT_MULTIPLIER = 6; // $5 → +30 credits (face × 6)
const BUBBLE_FACE_SET    = new Set(BUBBLE_FACES);

// The signature pad returns this 100% transparent PNG data URL when the user
// never touches the canvas. Treat it as "not signed" so logs remain clean.
const BLANK_SIGNATURE_PLACEHOLDER = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAeAAAACMCAYAAACklXoAAAAAAXNSR0IArs4c6QAAAERlWElmTU0AKgAAAAgAAYdpAAQAAAABAAAAGgAAAAAAA6ABAAMAAAABAAEAAKACAAQAAAABAAAB4KADAAQAAAABAAAAjAAAAAARyQiZAAAErElEQVR4Ae3QMQEAAADCoPVPbQsviEBhwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBgwIABAwYMGDBg4DMwGsgAAez54JgAAAAASUVORK5CYII=';

/***** PAGE 2 (Voucher Book) — standalone storage *****/
const VOUCHER_SHEET   = 'Voucher Transcribe';
const VOUCHER_HEADERS = [
  'Group ID','Saved At','First','Last','Prev Forward ($)','Row#',
  'Date','Added ($)','Used ($)','Running Balance ($)','Clerk'
];

/***** WEB APP ENTRY *****/
function doGet() {
  return HtmlService.createHtmlOutputFromFile('lookup')
    .setTitle('Client Balance Lookup');
}

/***** MENUS (minimal: rebuild + remove person) *****/
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Vouchers')
    .addItem('Rebuild Summaries Now', 'Rebuild_Now')
    .addItem('🧹 Remove Person…', 'promptAndRemovePerson')
    .addToUi();
}

/***** SIDEBARS (admin tools only) *****/
function showAdminAddPerson() {
  const html = HtmlService.createHtmlOutputFromFile('admin_add_person')
    .setTitle('Add Person • Manual Balance');
  SpreadsheetApp.getUi().showSidebar(html);
}
function showExactAdder() {
  const html = HtmlService.createHtmlOutputFromFile('exact_balance_adder')
    .setTitle('Exact Balance Adder');
  SpreadsheetApp.getUi().showSidebar(html);
}

/***** HELPERS *****/
function normHeader(s){
  return String(s || '')
    .toLowerCase()
    .replace(/\s+/g,'')
    .replace(/[\[\]\(\):_$€£¥%@#~*.,\-|]/g,'')
    .replace(/[^\p{L}\p{N}]+/gu,'');
}
function findHeaderIndex(headers, aliases){
  const H = headers.map(normHeader);
  const A = aliases.map(normHeader);
  for (let i=0;i<H.length;i++){
    for (let j=0;j<A.length;j++){
      const a = A[j];
      if (H[i] === a || H[i].includes(a)) return i;
    }
  }
  return -1;
}
function toNum(v){
  if (v === '' || v == null) return 0;
  const n = Number(String(v).replace(/[^0-9.\-]/g,''));
  return isNaN(n) ? 0 : n;
}
function monthParts(dt){
  if (!dt) return { month:'', year:'', monthYear:'', monthLabel:'' };
  const tz = Session.getScriptTimeZone();
  const month = Utilities.formatDate(dt, tz, 'MMMM');
  const year  = Utilities.formatDate(dt, tz, 'yyyy');
  const monthYear = Utilities.formatDate(dt, tz, 'MMMM yyyy');
  return { month, year, monthYear, monthLabel:monthYear };
}
function ensureSheetWithHeaders_(name, headers){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  } else if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  }
  return sh;
}

function ensureColumns_(sheet, definitions, minimumColumns) {
  if (!sheet) throw new Error('Missing sheet reference.');
  const defs = definitions || [];
  if (sheet.getLastRow() === 0 && defs.length) {
    sheet.getRange(1, 1, 1, defs.length).setValues([
      defs.map(d => d.label)
    ]);
  }

  const width = Math.max(sheet.getLastColumn(), minimumColumns || 0, defs.length || 0, 1);
  const headerRow = sheet.getRange(1, 1, 1, width).getValues()[0];
  const map = {};

  defs.forEach(def => {
    const aliases = (def.aliases || []).concat([def.label]);
    let idx = findHeaderIndex(headerRow, aliases);
    if (idx === -1) {
      idx = headerRow.length;
      sheet.getRange(1, idx + 1).setValue(def.label);
      headerRow[idx] = def.label;
    }
    map[def.key] = idx;
  });

  return { headerRow, map };
}
function cleanPhone(v){
  const s = String(v || '').trim();
  if (!s) return '';
  const digits = s.replace(/\D+/g,'');
  if (digits.length === 10) {
    return '('+digits.slice(0,3)+') '+digits.slice(3,6)+'-'+digits.slice(6);
  }
  return s;
}
function bubbleCredits(face){
  const amount = Number(face);
  return BUBBLE_FACES.includes(amount) ? amount * BUBBLE_CREDIT_MULTIPLIER : 0;
}
function fmtDollar_(n){
  const num = Number(n || 0);
  return '$' + num.toFixed(2);
}
function normalizeBubbleLabel_(raw){
  const cleaned = String(raw || '').trim();
  if (!cleaned) return '';

  const digitsOnly = cleaned.replace(/[^0-9.\-]+/g, '');
  const numeric = Number(digitsOnly);
  if (Number.isFinite(numeric) && BUBBLE_FACE_SET.has(numeric)) {
    return fmtDollar_(numeric);
  }

  const range = cleaned
    .replace(/[$+\s]/g, '')
    .replace(/[–—]/g, '-') // normalize en/em dash
    .trim();
  const m = range.match(/^(\d{1,2})-(\d{1,2})$/);
  if (m) {
    const a = Number(m[1]);
    const b = Number(m[2]);
    if (BUBBLE_FACE_SET.has(a) && BUBBLE_FACE_SET.has(b)) {
      return `$${a}-${b}`;
    }
  }

  return '';
}
function cap_(s){
  return String(s||'')
   .split(' ')
   .map(function(w){ return w ? (w[0].toUpperCase()+w.slice(1)) : ''; })
   .join(' ');
}

function normalizeSignature_(sig) {
  const cleaned = String(sig || '').trim();
  if (!cleaned) return '';
  if (cleaned === BLANK_SIGNATURE_PLACEHOLDER) return '';
  if (cleaned.startsWith('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAeAAAACMCAYAAACklXoAAAAAAXNSR0IArs4c6QAAAERlWElmTU0AKgAAAAgAAYdpAAQAAAABAAAAGgAAAAAAA6ABAAMAAAABAAEAAKACAAQAAAABAAAB4KADAAQAAAABAAAAjAAAAAARyQiZAAAErElEQVR4Ae3QMQEAAADCoPVPbQsviEBhwIAB')) return '';
  return cleaned;
}

/***** LOG WRITER — now includes optional Signature support *****/
function writeFormLogs(obj) {
  var now = obj.now,
      clerk = obj.clerk,
      last = obj.last,
      first = obj.first,
      cashNum = obj.cashNum,
      spentNum = obj.spentNum,
      earnedNum = obj.earnedNum,
      newBalance = obj.newBalance,
      responsesRow = obj.responsesRow,
      sig = normalizeSignature_(obj.signature), // ✅ optional signature (Base64)
      note = obj.note == null ? '' : String(obj.note);

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let fl = ss.getSheetByName(FORMLOG);
  if (!fl) {
    fl = ss.insertSheet(FORMLOG);
  }
  const flInfo = ensureColumns_(fl, [
    { key: 'timestamp', label: 'Timestamp', aliases: ['timestamp', 'time'] },
    { key: 'clerk', label: 'Clerk Initials', aliases: ['clerkinitials', 'clerk'] },
    { key: 'last', label: 'Last', aliases: ['last', 'lastname'] },
    { key: 'first', label: 'First', aliases: ['first', 'firstname'] },
    { key: 'cash', label: 'Cash Added', aliases: ['cashadded', 'cash'] },
    { key: 'spent', label: 'Credits Spent', aliases: ['creditsspent', 'spent'] },
    { key: 'earned', label: 'Credits Earned', aliases: ['creditsearned', 'earn'] },
    { key: 'notes', label: 'Clerk Notes', aliases: ['clerknotes', 'notes'] },
    { key: 'signature', label: 'Signature (Base64)', aliases: ['signature', 'signaturebase64'] }
  ], 8);
  const flRowLen = flInfo.headerRow.length;
  const flRow = new Array(flRowLen).fill('');
  flRow[flInfo.map.timestamp] = now;
  flRow[flInfo.map.clerk] = clerk;
  flRow[flInfo.map.last] = last;
  flRow[flInfo.map.first] = first;
  flRow[flInfo.map.cash] = cashNum;
  flRow[flInfo.map.spent] = spentNum;
  flRow[flInfo.map.earned] = earnedNum;
  flRow[flInfo.map.notes] = note;
  flRow[flInfo.map.signature] = sig;
  fl.appendRow(flRow);

  let rl = ss.getSheetByName(RAWFORMLOG);
  if (!rl) {
    rl = ss.insertSheet(RAWFORMLOG);
  }
  const rlInfo = ensureColumns_(rl, [
    { key: 'timestamp', label: 'Timestamp', aliases: ['timestamp', 'time'] },
    { key: 'clerk', label: 'Clerk Initials', aliases: ['clerkinitials', 'clerk'] },
    { key: 'last', label: 'Last', aliases: ['last', 'lastname'] },
    { key: 'first', label: 'First', aliases: ['first', 'firstname'] },
    { key: 'cash', label: 'Cash Added', aliases: ['cashadded', 'cash'] },
    { key: 'spent', label: 'Credits Spent', aliases: ['creditsspent', 'spent'] },
    { key: 'earned', label: 'Credits Earned', aliases: ['creditsearned', 'earn'] },
    { key: 'after', label: 'Credits After Transaction', aliases: ['creditsaftertransaction', 'creditsafter'] },
    { key: 'helper', label: 'Helper Column', aliases: ['helpercolumn'] },
    { key: 'row', label: 'Row#', aliases: ['row', 'row#', 'rownumber'] },
    { key: 'notes', label: 'Clerk Notes', aliases: ['clerknotes', 'notes'] },
    { key: 'signature', label: 'Signature (Base64)', aliases: ['signature', 'signaturebase64'] }
  ], 11);
  const rlRowLen = rlInfo.headerRow.length;
  const rlRow = new Array(rlRowLen).fill('');
  rlRow[rlInfo.map.timestamp] = now;
  rlRow[rlInfo.map.clerk] = clerk;
  rlRow[rlInfo.map.last] = last;
  rlRow[rlInfo.map.first] = first;
  rlRow[rlInfo.map.cash] = cashNum;
  rlRow[rlInfo.map.spent] = spentNum;
  rlRow[rlInfo.map.earned] = earnedNum;
  rlRow[rlInfo.map.after] = (newBalance != null ? newBalance : '');
  rlRow[rlInfo.map.helper] = note;
  rlRow[rlInfo.map.row] = responsesRow || '';
  rlRow[rlInfo.map.notes] = note;
  rlRow[rlInfo.map.signature] = sig;
  rl.appendRow(rlRow);
}

function clearBlankSignatures() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targets = [
    FORMLOG,
    RAWFORMLOG,
    'Signatures',
  ];

  let cleared = 0;

  targets.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;

    const dataRange = sh.getDataRange();
    const values = dataRange.getValues();
    if (!values || values.length < 2) return;

    const headers = values[0];
    const sigIdx = findHeaderIndex(headers, ['signature', 'signaturebase64']);
    if (sigIdx === -1) return;

    let changed = false;
    for (let r = 1; r < values.length; r++) {
      const original = values[r][sigIdx];
      const normalized = normalizeSignature_(original);
      if (normalized !== original) {
        if (original && !normalized) cleared++;
        values[r][sigIdx] = normalized;
        changed = true;
      }
    }

    if (changed) {
      dataRange.setValues(values);
    }
  });

  let ui = null;
  try {
    if (SpreadsheetApp && typeof SpreadsheetApp.getUi === 'function') {
      ui = SpreadsheetApp.getUi();
    }
  } catch (_err) {
    ui = null;
  }

  if (ui && typeof ui.alert === 'function') {
    ui.alert(`Cleared ${cleared} blank signature${cleared === 1 ? '' : 's'}.`);
  }

  return cleared;
}

/***** BACKEND: RETURN ALL CLIENT BALANCES (for offline cache warm) *****/
function getAllClientBalances(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(DONOR_SUMMARY);
  if (!sh) return [];

  const data = sh.getDataRange().getValues();
  if (data.length < 2) return [];

  const h = data[0];
  const idxLast  = findHeaderIndex(h, ['last','lastname','last name']);
  const idxFirst = findHeaderIndex(h, ['first','firstname','first name']);
  const idxBal   = findHeaderIndex(h, ['balance']);
  const idxPhone = findHeaderIndex(h, ['phone','mobile','tel']);

  const out = [];

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const last  = String(r[idxLast]  || '').trim();
    const first = String(r[idxFirst] || '').trim();
    if (!last || !first) continue;

    out.push({
      last,
      first,
      balance: Number(r[idxBal] || 0),
      phone: idxPhone > -1 ? cleanPhone(r[idxPhone] || '') : ''
    });
  }

  return out;
}

/***** RECEIPT HISTORY (last N transactions) *****/
function getRecentReceipts(limit) {
  const max = Math.min(Math.max(Number(limit) || 5, 1), 25);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(RAWFORMLOG) || ss.getSheetByName(FORMLOG);

  if (!sh) return { items: [] };
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { items: [] };

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const idxTS    = findHeaderIndex(headers, ['timestamp', 'time']);
  const idxClerk = findHeaderIndex(headers, ['clerk', 'clerkinitials']);
  const idxLast  = findHeaderIndex(headers, ['last', 'lastname']);
  const idxFirst = findHeaderIndex(headers, ['first', 'firstname']);
  const idxCash  = findHeaderIndex(headers, ['cash', 'cashadded']);
  const idxSpent = findHeaderIndex(headers, ['spent', 'creditsspent']);
  const idxEarn  = findHeaderIndex(headers, ['earned', 'creditsearned']);
  const idxAfter = findHeaderIndex(headers, ['creditsaftertransaction', 'creditsafter', 'after']);
  const idxNote  = findHeaderIndex(headers, ['notes', 'clerknotes']);

  const startRow = Math.max(2, lastRow - max + 1);
  const rows = sh.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();
  const tz = Session.getScriptTimeZone();

  const items = [];
  for (let i = rows.length - 1; i >= 0 && items.length < max; i--) {
    const r = rows[i];
    const rawTs = idxTS > -1 ? r[idxTS] : '';
    const when = rawTs ? Utilities.formatDate(new Date(rawTs), tz, 'MMM d • h:mm a') : '';
    const notes = idxNote > -1 ? String(r[idxNote] || '') : '';
    const clerk = idxClerk > -1 ? String(r[idxClerk] || '') : '';
    const isSysAdj = /sys\s*adj/i.test(notes) || /sys\s*adj/i.test(clerk);
    if (isSysAdj) continue;

    const rawCash = idxCash > -1 ? r[idxCash] : '';
    const cashNum = Number(rawCash);
    const cash = Number.isFinite(cashNum) ? cashNum : 0;
    const cashDisplay = rawCash == null ? '' : String(rawCash).trim();
    const bubbleLabel = normalizeBubbleLabel_(rawCash);
    const spent = idxSpent > -1 ? Number(r[idxSpent] || 0) : 0;
    const earned = idxEarn > -1 ? Number(r[idxEarn] || 0) : 0;
    const after = idxAfter > -1 ? Number(r[idxAfter] || 0) : null;
    const before = after == null ? null : Number((after - cash - earned + spent).toFixed(2));

    items.push({
      when,
      last: idxLast > -1 ? String(r[idxLast] || '') : '',
      first: idxFirst > -1 ? String(r[idxFirst] || '') : '',
      cash,
      cashDisplay,
      spent,
      earned,
      after,
      before,
      bubbleLabel,
      notes
    });
  }

  return { items };
}

/***** PHONE PRESERVATION (Donor Summary) *****/
function readExistingPhones_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(DONOR_SUMMARY);
  const map = new Map(); // "last|first" -> phone
  if (!sh || sh.getLastRow() < 2) return map;
  const data = sh.getDataRange().getValues();
  const h = data[0];
  const idxLast  = findHeaderIndex(h, ['last','lastname','last name']);
  const idxFirst = findHeaderIndex(h, ['first','firstname','first name']);
  const idxPhone = findHeaderIndex(h, ['phone','mobile','tel']);
  for (let i=1;i<data.length;i++){
    const r = data[i];
    const l = String(r[idxLast] || '').trim();
    const f = String(r[idxFirst]|| '').trim();
    if (!l || !f) continue;
    const key = l.toLowerCase()+'|'+f.toLowerCase();
    const phone = cleanPhone(r[idxPhone] || '');
    if (phone) map.set(key, phone);
  }
  return map;
}

// Optional: update (or insert) a phone number in Donor Summary from the UI.
function savePhoneIfChanged(obj) {
  const last  = obj.last;
  const first = obj.first;
  const p0    = obj.phone;
  const p     = cleanPhone(p0 || '');
  if (!last || !first) return { ok:false, msg:'Missing name.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(DONOR_SUMMARY) || ss.insertSheet(DONOR_SUMMARY);

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 8).setValues([[
      'Last Name','First Name','Phone',
      'Total Donated ($)','Credits Earned ($)',
      'Credits Spent ($)','System Adjustment ($)','Balance ($)'
    ]]);
  }

  let headerWidth = Math.max(sh.getLastColumn(), 3);
  let header = sh.getRange(1, 1, 1, headerWidth).getValues()[0];

  const idxLast  = findHeaderIndex(header, ['last','lastname','last name']);
  const idxFirst = findHeaderIndex(header, ['first','firstname','first name']);
  if (idxLast === -1 || idxFirst === -1) {
    return { ok:false, msg:'Missing name columns.' };
  }

  let idxPhone = findHeaderIndex(header, ['phone','mobile','tel']);
  if (idxPhone === -1) {
    idxPhone = header.length;
    sh.getRange(1, idxPhone + 1).setValue('Phone');
    headerWidth = Math.max(headerWidth, idxPhone + 1);
    header = sh.getRange(1, 1, 1, headerWidth).getValues()[0];
  }

  const data = sh.getDataRange().getValues();

  let row = -1;
  for (let i = 1; i < data.length; i++) {
    const l = String(data[i][idxLast] || '').trim().toLowerCase();
    const f = String(data[i][idxFirst] || '').trim().toLowerCase();
    if (l === String(last).trim().toLowerCase() &&
        f === String(first).trim().toLowerCase()) {
      row = i + 1;
      break;
    }
  }

  if (row === -1) {
    const width = Math.max(sh.getLastColumn(), header.length);
    const newRow = new Array(width).fill('');
    const numericTokens = ['total', 'donated', 'earned', 'spent', 'adjust', 'balance', 'credit'];
    header.forEach((label, idx) => {
      const norm = normHeader(label);
      if (numericTokens.some(token => norm.includes(token))) {
        newRow[idx] = 0;
      }
    });
    newRow[idxLast] = cap_(last);
    newRow[idxFirst] = cap_(first);
    if (idxPhone > -1) newRow[idxPhone] = p;
    sh.appendRow(newRow);
    return { ok:true, msg:'Phone saved (created row).' };
  }

  if (idxPhone > -1) {
    sh.getRange(row, idxPhone + 1).setValue(p);
  }
  return { ok:true, msg:'Phone updated.' };
}

// Update a client's profile (name and/or phone) in Donor Summary.
function updateClientProfile(obj) {
  const currentLast  = cap_(obj.last || '');
  const currentFirst = cap_(obj.first || '');
  const newLast      = cap_(obj.newLast || currentLast);
  const newFirst     = cap_(obj.newFirst || currentFirst);
  const phoneClean   = cleanPhone(obj.phone || '');

  if (!currentLast || !currentFirst) {
    return { ok:false, msg:'Missing current name.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(DONOR_SUMMARY);
  if (!sh) {
    return { ok:false, msg:'Donor Summary not found.' };
  }

  const data = sh.getDataRange().getValues();
  if (!data.length) return { ok:false, msg:'No data available.' };

  const h = data[0];
  const idxLast  = findHeaderIndex(h, ['last','lastname','last name']);
  const idxFirst = findHeaderIndex(h, ['first','firstname','first name']);
  let idxPhone   = findHeaderIndex(h, ['phone','mobile','tel']);
  if (idxLast === -1 || idxFirst === -1) {
    return { ok:false, msg:'Missing name columns.' };
  }

  if (idxPhone === -1) {
    idxPhone = h.length;
    sh.getRange(1, idxPhone + 1).setValue('Phone');
  }

  let row = -1;
  for (let i = 1; i < data.length; i++) {
    const l = String(data[i][idxLast] || '').trim().toLowerCase();
    const f = String(data[i][idxFirst] || '').trim().toLowerCase();
    if (l === currentLast.toLowerCase() && f === currentFirst.toLowerCase()) {
      row = i + 1;
      break;
    }
  }

  if (row === -1) {
    return { ok:false, msg:'Person not found.' };
  }

  const finalLast  = newLast || currentLast;
  const finalFirst = newFirst || currentFirst;
  const oldPhone   = cleanPhone(data[row - 1][idxPhone] || '');
  const finalPhone = phoneClean || oldPhone;

  sh.getRange(row, idxLast + 1).setValue(finalLast);
  sh.getRange(row, idxFirst + 1).setValue(finalFirst);
  if (idxPhone > -1 && finalPhone !== undefined) {
    sh.getRange(row, idxPhone + 1).setValue(finalPhone);
  }

  return {
    ok: true,
    msg: 'Profile updated.',
    last: finalLast,
    first: finalFirst,
    phone: finalPhone
  };
}

function makeMonthlySnapshot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const stamp = Utilities.formatDate(now, tz, 'yyyy-MM');
  const src = ss.getSheetByName(PROGRAM_SUMMARY);
  if (!src) throw new Error('Missing Program Summary');

  // Rebuild everything first to ensure fresh data
  rebuildFromResponses_();

  // Create a new sheet for the snapshot
  const name = `Monthly Snapshot ${stamp}`;
  const sh = ss.insertSheet(name);

  // Copy header and main data area (A1:K40 should hold your dashboard tables)
  src.getRange('A1:K40').copyTo(sh.getRange('A1'), {contentsOnly: true});

  // Label and format
  sh.getRange('A1').setValue(`Program Summary Snapshot — ${stamp}`);
  sh.autoResizeColumns(1, 11);
  sh.setFrozenRows(1);

  SpreadsheetApp.getUi().alert(`✅ Monthly snapshot created: ${name}`);
}


/***** BACKEND: LOOKUP *****/
function lookupClient(obj){
  const last  = obj.last;
  const first = obj.first;

  if (!last || !first) return { error:'Missing name.' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const donor = ss.getSheetByName(DONOR_SUMMARY);
  if (!donor) return { error:'Missing Donor Summary.' };

  const data = donor.getDataRange().getValues();
  const h = data[0];
  const idxLast  = findHeaderIndex(h, ['last','lastname','last name']);
  const idxFirst = findHeaderIndex(h, ['first','firstname','first name']);
  const idxBal   = findHeaderIndex(h, ['balance']); // NOTE: stored in "money units" which equal your $ for UI
  const idxPhone = findHeaderIndex(h, ['phone','mobile','tel']);

  for (let i=1;i<data.length;i++){
    const r = data[i];
    const l = String(r[idxLast]||'').trim().toLowerCase();
    const f = String(r[idxFirst]||'').trim().toLowerCase();
    if (l === String(last).trim().toLowerCase() &&
        f === String(first).trim().toLowerCase()){
      const phone = idxPhone>-1 ? cleanPhone(r[idxPhone]||'') : '';
      return { last:cap_(last), first:cap_(first), balance:Number(r[idxBal]||0), phone:phone };
    }
  }
  return { error:'Person not found.' };
}

/***** HELPER: DIRECTLY SET STARTING BALANCE (optional) *****/
function applyStartingBalanceOverride(obj){
  // obj: { last, first, clerk, desiredStartD }
  const last  = obj.last;
  const first = obj.first;
  const clerk = obj.clerk || 'SYS (Adj)';
  const wantD = Number(obj.desiredStartD);

  if (!last || !first || isNaN(wantD)) return { success:false, message:'Missing name or starting balance.' };

  const now = new Date();
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sh  = ss.getSheetByName(RESPONSES_SHEET);

  if (!sh) return { success:false, message:'Missing Form responses 1' };

  // Find current balance
  const cur = lookupClient({ last:last, first:first });
  const curBalD = Number(cur && !cur.error ? cur.balance || 0 : 0);
  const deltaD  = wantD - curBalD;

  if (Math.abs(deltaD) < 0.000001){
    return { success:true, message:'No change to starting balance.' };
  }

  sh.appendRow([now, clerk, last, first, deltaD, 0]);
  writeFormLogs({
    now: now,
    clerk: clerk,
    last: last,
    first: first,
    cashNum: deltaD,
    spentNum: 0,
    earnedNum: 0,
    newBalance: curBalD + deltaD
  });

  rebuildFromResponses_();
  return { success:true, message:'Starting balance override applied.' };
}

/***** FAST SAVE → INSTANT REBUILD + HONOR OVERRIDES (Stable 2025-10) *****/
/*
  Accepts:
    { last, first, clerk, cash, spent, desiredStartD? }
  - cash  = face dollars ($5 → adds 30 credits)
  - spent = already in credits (UI sends ×6)
  - desiredStartD = override balance before transaction (in dollars)
*/
function addToFormResponses1(e) {
  try {
    const last = String(e.last || '').trim();
    const first = String(e.first || '').trim();
    const clerk = String(e.clerk || '').trim();
    if (!last || !first) {
      return { success:false, message:'Missing client name.' };
    }

    const rawCash = toNum(e.cash);
    const cashFace = BUBBLE_FACES.includes(rawCash) ? rawCash : 0;
    const rawSpent = toNum(e.spent);
    const spentNum = rawSpent > 0 ? rawSpent : 0;
    const note = typeof e.note === 'string' ? e.note.trim() : '';
    const signature = normalizeSignature_(e.signature);
    const phoneRaw = typeof e.phone === 'string' ? e.phone.trim() : '';
    const phoneNormalized = cleanPhone(phoneRaw);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(RESPONSES_SHEET);
    if (!sh) return { success:false, message:'Missing Form responses 1' };

    const now = new Date();

    const headerWidth = Math.max(sh.getLastColumn(), 6);
    const headerRow = sh.getRange(1, 1, 1, headerWidth).getValues()[0];
    let noteCol = findHeaderIndex(headerRow, ['note','notes','clerknotes']);
    if (noteCol === -1) {
      noteCol = headerRow.length;
      sh.getRange(1, noteCol + 1).setValue('Clerk Notes');
      headerRow[noteCol] = 'Clerk Notes';
    }

    let phoneCol = findHeaderIndex(headerRow, ['phone','mobile','tel']);
    if (phoneCol === -1) {
      phoneCol = headerRow.length;
      sh.getRange(1, phoneCol + 1).setValue('Phone');
      headerRow[phoneCol] = 'Phone';
    }

    // --- 0) Current balance in credit dollars ---
    const cur = lookupClient({ last, first });
    const curBalD = Number(cur && !cur.error ? cur.balance || 0 : 0);

    // --- 1) Handle override (treat "0" as valid) ---
    let startBalance = curBalD;
    const rawDesired = e.desiredStartD;
    const hasOverride = rawDesired !== undefined && rawDesired !== null && String(rawDesired).trim() !== '';
    if (hasOverride) {
      const wantD = Number(rawDesired);
      if (!isNaN(wantD)) {
        startBalance = wantD;
        const deltaD = wantD - curBalD;
        if (Math.abs(deltaD) > 0.009) {
          sh.appendRow([now, 'SYS (Adj)', last, first, deltaD, 0]);
          writeFormLogs({
            now,
            clerk: 'SYS (Adj)',
            last,
            first,
            cashNum: deltaD,
            spentNum: 0,
            earnedNum: 0,
            newBalance: startBalance
          });
        }
      } else {
        startBalance = curBalD;
      }
    }

    // --- 2) Record the actual visit row ---
    const rowLen = Math.max(headerWidth, noteCol + 1, phoneCol + 1);
    const newRow = new Array(rowLen).fill('');
    newRow[0] = now;
    newRow[1] = clerk;
    newRow[2] = last;
    newRow[3] = first;
    newRow[4] = cashFace;
    newRow[5] = spentNum;
    newRow[noteCol] = note;
    if (phoneCol > -1) {
      newRow[phoneCol] = phoneNormalized;
    }
    sh.appendRow(newRow);
    const responsesRow = sh.getLastRow();

    // --- 3) Rebuild summaries immediately ---
    rebuildFromResponses_();

    // --- 4) Log the visit using the contract math ---
    const earnedNum = bubbleCredits(cashFace);
    const newBalance = startBalance + earnedNum - spentNum;

    writeFormLogs({
      now,
      clerk,
      last,
      first,
      cashNum: cashFace,
      spentNum,
      earnedNum,
      newBalance,
      responsesRow,
      signature,
      note
    });

    return { success:true, newBalance };
  } catch (err) {
    return { success:false, message:String(err && err.message ? err.message : err) };
  }
}

/***** ADMIN VOUCHER FUNCTIONS (backend only) *****/
function adminVoucherLedger$(obj){
  const last  = obj.last;
  const first = obj.first;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const donor = ss.getSheetByName(DONOR_SUMMARY);
  if (!donor) return { balance:0 };
  const data = donor.getDataRange().getValues();
  const idxL = findHeaderIndex(data[0], ['last']);
  const idxF = findHeaderIndex(data[0], ['first']);
  const idxB = findHeaderIndex(data[0], ['balance']);
  for (let i=1;i<data.length;i++){
    if (String(data[i][idxL]).trim().toLowerCase()===String(last).trim().toLowerCase() &&
        String(data[i][idxF]).trim().toLowerCase()===String(first).trim().toLowerCase()){
      return { balance:Number(data[i][idxB]||0) };
    }
  }
  return { balance:0 };
}
function adminVoucherTranscribeSave(obj){
  const last  = obj.last;
  const first = obj.first;
  const previousBalance = obj.previousBalance;
  const rows  = obj.rows;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ensureSheetWithHeaders_(VOUCHER_SHEET, VOUCHER_HEADERS);
  const gid = Utilities.getUuid().slice(0,8);
  const now = new Date();
  let run = Number(previousBalance||0);
  let i = sh.getLastRow()+1;
  for (let k=0;k<rows.length;k++){
    const r = rows[k];
    run += (Number(r.added||0) - Number(r.used||0));
    sh.appendRow([gid, now, first, last, previousBalance, i - 1, r.date, r.added, r.used, run, 'Admin']);
    i++;
  }
  return { ok:true, msg:'Saved voucher entries.' };
}

// Lightweight JSON feed for voucher ledger UI
// Returns vouchers grouped by Group ID with historical rows attached
function fetchVoucherJson(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(VOUCHER_SHEET);
  if(!sh){
    return { vouchers: [] };
  }

  const data = sh.getDataRange().getValues();
  if(!data.length){
    return { vouchers: [] };
  }

  const headers = data[0];
  const idxGroup  = findHeaderIndex(headers, ['group id','groupid','gid','voucher id','id']);
  const idxFirst  = findHeaderIndex(headers, ['first']);
  const idxLast   = findHeaderIndex(headers, ['last']);
  const idxPrev   = findHeaderIndex(headers, ['prev forward','previous','starting balance']);
  const idxRowNum = findHeaderIndex(headers, ['row#','row','row number']);
  const idxDate   = findHeaderIndex(headers, ['date']);
  const idxAdded  = findHeaderIndex(headers, ['added ($)','added','cash']);
  const idxUsed   = findHeaderIndex(headers, ['used ($)','used','spent']);
  const idxRun    = findHeaderIndex(headers, ['running balance','running balance ($)','running']);

  const map = new Map();
  for(let i = 1; i < data.length; i++){
    const row = data[i];
    if(!row || row.every(cell => String(cell || '').trim() === '')) continue;

    const groupId = idxGroup === -1 ? `gid-${i}` : (row[idxGroup] || `gid-${i}`);
    const entry = map.get(groupId) || {
      id: groupId,
      groupId,
      first: String(idxFirst > -1 ? row[idxFirst] : '').trim(),
      last: String(idxLast > -1 ? row[idxLast] : '').trim(),
      previous: toNum(idxPrev > -1 ? row[idxPrev] : 0),
      rows: []
    };

    entry.rows.push({
      rowNumber: idxRowNum > -1 ? row[idxRowNum] : '',
      date: idxDate > -1 ? row[idxDate] : '',
      added: toNum(idxAdded > -1 ? row[idxAdded] : 0),
      spent: toNum(idxUsed > -1 ? row[idxUsed] : 0),
      runningBalance: toNum(idxRun > -1 ? row[idxRun] : 0)
    });

    map.set(groupId, entry);
  }

  return { vouchers: Array.from(map.values()) };
}

/***** CORE — REBUILD (Merged + Donor Summary) *****/
function rebuildFromResponses_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(RESPONSES_SHEET);
  if (!src) throw new Error('Missing sheet: ' + RESPONSES_SHEET);

  const data = src.getDataRange().getValues();
  const hasData = data.length >= 2 && data.slice(1).some(r => r.join('').trim() !== '');

  // Preserve any phone numbers that were manually saved before we rebuild.
  const existingPhones = readExistingPhones_();

  // Always clear and rebuild the Donor Summary — even if Form responses is empty
  const donor = ss.getSheetByName(DONOR_SUMMARY) || ss.insertSheet(DONOR_SUMMARY);
  donor.clear({ contentsOnly: true });

  // Set headers so the tab is never blank
  const donorHeader = [
    'Last Name','First Name','Phone',
    'Total Donated ($)','Credits Earned ($)',
    'Credits Spent ($)','System Adjustment ($)','Balance ($)'
  ];
  donor.getRange(1, 1, 1, donorHeader.length).setValues([donorHeader]);

  // If no Form responses data, skip the rest safely
  if (!hasData) {
    buildOrUpdateDashboard_();
    SpreadsheetApp.flush();
    return;
  }

  const headers = data[0];
  const rows = data.slice(1);
  const tz = Session.getScriptTimeZone();

  const idxTS    = findHeaderIndex(headers, ['timestamp','time','date']);
  const idxClerk = findHeaderIndex(headers, ['clerkinitials','clerk','initials']);
  const idxLast  = findHeaderIndex(headers, ['last','lastname']);
  const idxFirst = findHeaderIndex(headers, ['first','firstname']);
  const idxCash  = findHeaderIndex(headers, ['cashadded','cash','amount']);
  const idxSpent = findHeaderIndex(headers, ['creditsspent','spent']);
  const idxPhoneResp = findHeaderIndex(headers, ['phone','mobile','tel']);

  const byKey = new Map();
  const keyOf = (l,f) => `${String(l).trim().toLowerCase()}|${String(f).trim().toLowerCase()}`;

  // --- Go through all form response rows ---
  rows.forEach(r => {
    const last  = String(r[idxLast]  || '').trim();
    const first = String(r[idxFirst] || '').trim();
    if (!last || !first) return;
    const key = keyOf(last, first);
    const cashRaw  = toNum(r[idxCash]);
    const spentRaw = toNum(r[idxSpent]);
    const cashFace = BUBBLE_FACES.includes(cashRaw) ? cashRaw : 0;
    const spent = spentRaw > 0 ? spentRaw : 0;
    const clerk = String(r[idxClerk] || '').trim().toUpperCase();

    if (!byKey.has(key)) {
      byKey.set(key, {
        last,
        first,
        phone: existingPhones.get(key) || '',
        totalDonated: 0,
        earned: 0,
        spent: 0,
        adjustment: 0,
      });
    }

    const p = byKey.get(key);

    if (idxPhoneResp > -1) {
      const normalizedPhone = cleanPhone(r[idxPhoneResp] || '');
      if (normalizedPhone) {
        p.phone = normalizedPhone;
      }
    }

    // Real bubble donations (not SYS)
    if (clerk !== 'SYS (ADJ)') {
      const earned = bubbleCredits(cashFace);
      if (earned > 0) {
        p.totalDonated += earned / BUBBLE_CREDIT_MULTIPLIER;
        p.earned += earned;
      }
    }

    // ALWAYS count spent in Credits Spent ($)
    p.spent += spent;

    // Track manual adjustments separately
    if (clerk === 'SYS (ADJ)') {
      // +cashRaw = manual credit increase, +spentRaw = manual credit decrease
      p.adjustment += cashRaw - spentRaw;
    }
  }); // ✅ properly close the forEach loop

  const donorRows = Array.from(byKey.values()).map(p => {
    return [
      cap_(p.last),
      cap_(p.first),
      p.phone || '',
      p.totalDonated,
      p.earned,
      p.spent,
      p.adjustment || 0,
      p.earned - p.spent + (p.adjustment || 0) // ✅ includes system adjustment in balance
    ];
  });

  if (donorRows.length) donor.getRange(2, 1, donorRows.length, donorHeader.length).setValues(donorRows);

  // --- Build Form responses (Merged) ---
  const mergedHeader = [
    'Last','First','Timestamp','Clerk',
    'Cash Added ($)','Credits Spent','Credits Earned',
    'Net (Earned-Spent)','Month','Year','Month-Year','Month Label'
  ];
  const mergedRows = [];

  rows.forEach(r => {
    const last  = String(r[idxLast]  || '').trim();
    const first = String(r[idxFirst] || '').trim();
    if (!last || !first) return;
    const clerkRaw = String(r[idxClerk] || '').trim();
    const clerk = clerkRaw || '';
    const isAdjustment = clerkRaw.toUpperCase() === 'SYS (ADJ)';
    const cashRaw = toNum(r[idxCash]);
    const spentRaw = toNum(r[idxSpent]);
    const cashFace = BUBBLE_FACES.includes(cashRaw) ? cashRaw : 0;
    const spent = spentRaw > 0 ? spentRaw : 0;
    let earned = 0;
    if (isAdjustment) {
      earned = cashRaw;
    } else {
      earned = bubbleCredits(cashFace);
    }
    const net = isAdjustment ? (cashRaw - spentRaw) : (earned - spent);

    const ts = r[idxTS] ? new Date(r[idxTS]) : null;
    const tsStr = ts ? Utilities.formatDate(ts, tz, 'yyyy-MM-dd HH:mm') : '';
    const mp = monthParts(ts);
    mergedRows.push([last, first, tsStr, clerk, isAdjustment ? cashRaw : cashFace, spent, earned, net, mp.month, mp.year, mp.monthYear, mp.monthLabel]);
  });

  const merged = ensureSheetWithHeaders_(RESPONSES_MERGED, mergedHeader);
  merged.clear({ contentsOnly: true });
  merged.getRange(1, 1, 1, mergedHeader.length).setValues([mergedHeader]);
  if (mergedRows.length) merged.getRange(2, 1, mergedRows.length, mergedHeader.length).setValues(mergedRows);

  SpreadsheetApp.flush();
  buildOrUpdateDashboard_();
}

/***** VISUALS: remove all charts/images/drawings (prevents duplicates) *****/
function removeAllVisuals_(sh){
  try { sh.getCharts().forEach(function(c){ sh.removeChart(c); }); } catch(e){}
  try { sh.getImages().forEach(function(img){ img.remove(); }); } catch(e){}
  try { if (sh.getDrawings) sh.getDrawings().forEach(function(d){ d.remove(); }); } catch(e){}
}

/***** DASHBOARD DATA (metrics + tables without sheet formulas) *****/
function computeDashboardData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mergedSheet = ss.getSheetByName(RESPONSES_MERGED);
  const donorSheet = ss.getSheetByName(DONOR_SUMMARY);
  const rawLogSheet = ss.getSheetByName(RAWFORMLOG);

  const mergedRows = mergedSheet ? mergedSheet.getDataRange().getValues().slice(1) : [];
  const donorRows = donorSheet ? donorSheet.getDataRange().getValues().slice(1) : [];
  const rawLogRows = rawLogSheet ? rawLogSheet.getDataRange().getValues().slice(1) : [];
  const tz = Session.getScriptTimeZone();

  const normalizeClerk = (v) => String(v || '').trim().toUpperCase();

  const asDate = (value) => {
    if (!value) return null;
    if (value instanceof Date && !isNaN(value)) return value;
    const str = String(value).trim();
    if (!str) return null;
    const parsed = new Date(str);
    return isNaN(parsed) ? null : parsed;
  };

  const monthFromTimestamp = (value) => {
    const ts = asDate(value);
    if (!ts) {
      return {
        key: 'unknown',
        label: 'Unknown Month',
        sortKey: Number.MAX_SAFE_INTEGER,
      };
    }
    const mp = monthParts(ts);
    const label = mp.monthLabel || mp.monthYear || 'Unknown Month';
    const sortKey = ts.getFullYear() * 100 + (ts.getMonth() + 1);
    const key = `${ts.getFullYear()}-${String(ts.getMonth() + 1).padStart(2, '0')}`;
    return { key, label, sortKey };
  };

  const resolveMonthBucketFromMerged = (row) => {
    if (!row || !row.length) {
      return { key: 'unknown', label: 'Unknown Month', sortKey: Number.MAX_SAFE_INTEGER };
    }
    const tsDate = asDate(row[2]);
    const labelDate = asDate(row[11]);
    const source = labelDate || tsDate;
    if (source) {
      return monthFromTimestamp(source);
    }
    const rawMonth = String(row[8] || '').trim();
    const rawYear = String(row[9] || '').trim();
    if (rawMonth && rawYear) {
      const parsed = asDate(`${rawMonth} 1 ${rawYear}`);
      if (parsed) return monthFromTimestamp(parsed);
    }
    const fallback = String(row[11] || '').trim() || `${rawMonth} ${rawYear}`.trim();
    if (fallback) {
      const parsed = asDate(fallback);
      if (parsed) return monthFromTimestamp(parsed);
      return {
        key: fallback,
        label: fallback,
        sortKey: Number.MAX_SAFE_INTEGER,
      };
    }
    return {
      key: 'unknown',
      label: 'Unknown Month',
      sortKey: Number.MAX_SAFE_INTEGER,
    };
  };

  const rowsForMetrics = rawLogRows.length ? rawLogRows : mergedRows;
  const usingRawLog = rawLogRows.length > 0;
  const metricsSource = usingRawLog ? RAWFORMLOG : RESPONSES_MERGED;

  const metricRows = rowsForMetrics
    .filter(row => row && row.length)
    .map(row => {
      const clerk = normalizeClerk(row[usingRawLog ? 1 : 3]);
      return {
        clerk,
        isSystem: clerk === 'SYS (ADJ)',     // ← used for filtering
        cash: toNum(row[usingRawLog ? 4 : 4]),
        spent: toNum(row[usingRawLog ? 5 : 5]),
        bucket: usingRawLog ? monthFromTimestamp(row[0]) : resolveMonthBucketFromMerged(row),
      };
    });

  // Pick latest non-empty month
  let selectedBucket = null;
  metricRows.forEach(info => {
    if (info.isSystem) return; // <-- EXCLUDE SYS (Adj)
    if (!info.bucket || info.bucket.key === 'unknown') return;
    if (info.cash <= 0 && info.spent <= 0) return;
    if (!selectedBucket || info.bucket.sortKey > selectedBucket.sortKey) {
      selectedBucket = info.bucket;
    }
  });
  if (!selectedBucket) {
    metricRows.forEach(info => {
      if (info.isSystem) return; // <-- EXCLUDE SYS (Adj)
      if (!info.bucket || info.bucket.key === 'unknown') return;
      if (!selectedBucket || info.bucket.sortKey > selectedBucket.sortKey) {
        selectedBucket = info.bucket;
      }
    });
  }

  const filterByMonth = Boolean(selectedBucket);

  let totalAdded = 0;
  let totalUsed = 0;
  const usedPerTxn = [];
  const monthly = new Map();

  const accumulate = (bucketInfo, cash, spent) => {
    if (!bucketInfo || !bucketInfo.label) return;
    const { key, label, sortKey } = bucketInfo;
    if (!monthly.has(key)) {
      monthly.set(key, { label, added: 0, used: 0, sort: sortKey });
    }
    const target = monthly.get(key);
    if (cash > 0) target.added += cash;
    if (spent > 0) target.used += spent;
  };

  metricRows.forEach(info => {
    // 🔥 FULL EXCLUSION OF SYS (ADJ) — this is the key change
    if (info.isSystem) return;

    const { cash, spent, bucket } = info;

    const bucketMatches = !filterByMonth || (bucket && bucket.key === selectedBucket.key);

    // monthly table (clean)
    accumulate(bucket, cash, spent);

    // tile totals (clean)
    if (bucketMatches) {
      if (cash > 0) totalAdded += cash;
      if (spent > 0) {
        totalUsed += spent;
        usedPerTxn.push(spent);
      }
    }
  });

  const metricsPeriod = (selectedBucket && selectedBucket.key !== 'unknown')
    ? { key: selectedBucket.key, label: selectedBucket.label }
    : { key: 'all-time', label: 'All Time' };

  const balances = donorRows.map(row => toNum(row[7]));
  const avgBalance = balances.length ? balances.reduce((sum, v) => sum + v, 0) / balances.length : 0;
  const avgUsed = usedPerTxn.length ? usedPerTxn.reduce((sum, v) => sum + v, 0) / usedPerTxn.length : 0;

  const monthlyRows = Array.from(monthly.values())
    .sort((a, b) => {
      const diff = a.sort - b.sort;
      return diff !== 0 ? diff : a.label.localeCompare(b.label);
    })
    .map(item => [
      item.label,
      Number(item.added.toFixed(2)),
      Number(item.used.toFixed(2)),
    ]);

  const clientRows = donorRows
    .map(row => {
      const last = String(row[0] || '').trim();
      const first = String(row[1] || '').trim();
      if (!last && !first) return null;
      const name = first && last ? `${first}, ${last}` : (first || last);
      return [
        name,
        Number(toNum(row[3]).toFixed(2)), // Total Donated ($)
        Number(toNum(row[5]).toFixed(2)), // Credits Spent ($)
      ];
    })
    .filter(Boolean)
    .sort((a, b) => {
      const usedDiff = b[2] - a[2];
      if (usedDiff !== 0) return usedDiff;
      return a[0].localeCompare(b[0]);
    })
    .slice(0, 10);

  return {
    metrics: {
      totalAdded: Number(totalAdded.toFixed(2)),
      totalUsed: Number(totalUsed.toFixed(2)),
      avgBalance: Number(avgBalance.toFixed(2)),
      avgUsed: Number(avgUsed.toFixed(2)),
    },
    systemTotals: {
      totalAdded: 0, // dashboard ignores sys adj entirely
      totalUsed: 0,
    },
    metricsPeriod,
    monthlyRows,
    clientRows,
    sources: {
      metrics: metricsSource,
      monthly: usingRawLog ? RAWFORMLOG : RESPONSES_MERGED,
      clients: DONOR_SUMMARY,
    },
  };
}

function showTopSpendersDashboard() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const donor = ss.getSheetByName(DONOR_SUMMARY);

  if (!donor) {
    ui.alert('Missing Donor Summary sheet.');
    return;
  }

  const values = donor.getDataRange().getValues();
  if (!values.length || values.length === 1) {
    ui.alert('No donor data found.');
    return;
  }

  const header = values[0];
  const idxLast = findHeaderIndex(header, ['last', 'lastname', 'last name']);
  const idxFirst = findHeaderIndex(header, ['first', 'firstname', 'first name']);
  const idxUsed = findHeaderIndex(header, ['used', 'money used', 'spent', 'credits spent']);
  const usedIdx = idxUsed > -1 ? idxUsed : 5; // fallback to common "Used ($)" column

  if (idxLast === -1 && idxFirst === -1) {
    ui.alert('Missing first/last name columns in Donor Summary.');
    return;
  }

  const rows = values.slice(1)
    .map(row => {
      const last = String(row[idxLast] || '').trim();
      const first = String(row[idxFirst] || '').trim();
      const used = Number(toNum(row[usedIdx]).toFixed(2));
      if (!last && !first) return null;
      const name = first && last ? `${first} ${last}` : (first || last);
      return { name, used };
    })
    .filter(Boolean)
    .sort((a, b) => b.used - a.used || a.name.localeCompare(b.name))
    .slice(0, 10);

  if (!rows.length) {
    ui.alert('No donor activity yet.');
    return;
  }

  const lines = rows.map((row, idx) => `${idx + 1}. ${row.name} — $${row.used.toFixed(2)}`);
  const body = lines.concat('', 'Read-only view — no sheet data was changed.').join('\n');
  ui.alert('Top 10 Spenders', body, ui.ButtonSet.OK);
}

/***** DASHBOARD (Rose-Pine + Legend row + Single IMAGE chart) *****/
function buildOrUpdateDashboard_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(PROGRAM_SUMMARY);
  if (!sh) return;

  const data = computeDashboardData_();
  const RP = { base:'#191724',surface:'#1f1d2e',overlay:'#26233a',muted:'#6e6a86',text:'#e0def4',love:'#eb6f92',foam:'#9ccfd8',gold:'#f6c177' };

  removeAllVisuals_(sh);
  sh.clear();
  sh.clearFormats();

  const maxDataRows = Math.max(data.monthlyRows.length, data.clientRows.length, 1);
  const fillRows = Math.max(60, 12 + maxDataRows);

  sh.setFrozenRows(5);
  sh.setColumnWidths(1, 11, 180);
  sh.setRowHeights(1, 5, 48);
  if (fillRows > 5) {
    sh.setRowHeights(6, fillRows - 5, 28);
  }

  sh.getRange(1, 1, fillRows, 11)
    .setBackground(RP.base)
    .setFontColor(RP.text);

  // Title & legend
  sh.getRange('A1:K1').merge().setValue("Nancy's Hope — Program Summary")
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setFontWeight('bold').setFontSize(28).setFontColor(RP.love);

  sh.getRange('A4:K4').merge()
    .setValue('Legend — 💙 Money Added • 💗 Money Used')
    .setHorizontalAlignment('center')
    .setFontSize(12).setFontWeight('bold').setFontColor(RP.text)
    .setBackground(RP.surface);

  const metricsSourceLabel = (data.sources && data.sources.metrics) || RESPONSES_MERGED;
  const metricsSourceDetail = metricsSourceLabel === RAWFORMLOG
    ? 'RawFormLog — Cash Added & Credits Spent'
    : 'Form responses (Merged) — Cash Added & Credits Spent';
  const systemTotals = data.systemTotals || { totalAdded: 0, totalUsed: 0 };
  const systemNote = (systemTotals.totalAdded || systemTotals.totalUsed)
    ? ` — excluded SYS (Adj): +$${Number(systemTotals.totalAdded || 0).toFixed(2)} added, +$${Number(systemTotals.totalUsed || 0).toFixed(2)} used`
    : '';
  const metricsPeriodLabel = (data.metricsPeriod && data.metricsPeriod.label) || 'All Time';
  const periodNote = metricsPeriodLabel ? ` — period: ${metricsPeriodLabel}` : '';
  sh.getRange('A6:K6').merge()
    .setValue(`Money Added/Used source: ${metricsSourceDetail}${periodNote}${systemNote}`)
    .setHorizontalAlignment('center')
    .setFontSize(10)
    .setFontColor(RP.muted)
    .setBackground(RP.base);

  // KPI tiles
  const tileCols = ['C', 'E', 'G', 'I'];
  const labels = [
    'Total Added ($)',
    'Total Used ($)',
    'Avg Balance ($)',
    'Avg Used / Txn ($)'
  ];
  const values = [
    data.metrics.totalAdded,
    data.metrics.totalUsed,
    data.metrics.avgBalance,
    data.metrics.avgUsed,
  ];

  for (let i = 0; i < tileCols.length; i++) {
    const col = tileCols[i];
    sh.getRange(`${col}2`).setValue(labels[i])
      .setFontWeight('bold').setFontSize(14).setFontColor(RP.muted)
      .setHorizontalAlignment('center');
    sh.getRange(`${col}3`).setValue(values[i])
      .setNumberFormat('$#,##0.00')
      .setFontWeight('bold').setFontSize(22)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontColor(RP.text)
      .setBackground(RP.overlay);
  }

  // Section headers
  sh.getRange('A5').setValue('MONTHLY OVERVIEW').setFontWeight('bold').setFontColor(RP.muted);
  sh.getRange('F5').setValue('TOP 10 SPENDERS').setFontWeight('bold').setFontColor(RP.muted);

  // Monthly table
  sh.getRange('A7:C7').setValues([['Month','Money Added ($)','Money Used ($)']])
    .setFontWeight('bold').setFontSize(14).setFontColor(RP.text).setBackground(RP.overlay)
    .setHorizontalAlignment('center');

  const monthlyDataRows = Math.max(data.monthlyRows.length, 1);
  const monthlyFillRows = Math.max(fillRows - 7, monthlyDataRows);
  if (monthlyFillRows > 0) {
    sh.getRange(8, 1, monthlyFillRows, 3)
      .setBackground(RP.surface)
      .setFontColor(RP.text)
      .setFontSize(12);
  }
  const monthlyRows = data.monthlyRows.length ? data.monthlyRows : [['No monthly data yet', '', '']];
  const monthlyRange = sh.getRange(8, 1, monthlyRows.length, 3);
  monthlyRange.setValues(monthlyRows);
  if (monthlyFillRows > 0) {
    sh.getRange(8, 2, monthlyFillRows, 2).setNumberFormat('$#,##0.00');
  }

  // Client snapshot table
  sh.getRange('G7:I7').setValues([['Name','Money Added ($)','Money Used ($)']])
    .setFontWeight('bold').setFontSize(14).setFontColor(RP.text).setBackground(RP.overlay)
    .setHorizontalAlignment('center');

  const clientDataRows = Math.max(data.clientRows.length, 1);
  const clientFillRows = Math.max(fillRows - 7, clientDataRows);
  if (clientFillRows > 0) {
    sh.getRange(8, 7, clientFillRows, 3)
      .setBackground(RP.surface)
      .setFontColor(RP.text)
      .setFontSize(13);
  }
  const clientRows = data.clientRows.length ? data.clientRows : [['No donor activity yet', '', '']];
  const clientRange = sh.getRange(8, 7, clientRows.length, 3);
  clientRange.setValues(clientRows);
  if (clientFillRows > 0) {
    sh.getRange(8, 8, clientFillRows, 2).setNumberFormat('$#,##0.00');
  }

  removeAllVisuals_(sh);
  SpreadsheetApp.flush();
  addMonthlySummaryChartAsImage_(sh);

  sh.getRange(1, 1, fillRows, 11).setFontColor(RP.text);
}

/***** IMAGE CHART (blue = Added, pink = Used) *****/
function addMonthlySummaryChartAsImage_(sh){
  const RP = { base:'#191724', surface:'#1f1d2e', overlay:'#26233a', text:'#e0def4', love:'#eb6f92', foam:'#9ccfd8' };
  const HEADER_ROW = 7;
  const DATA_START = 8;
  const SNAP_COL = 'AQ';
  const SNAP_COLS = 3;

  removeAllVisuals_(sh);
  SpreadsheetApp.flush();

  const maxRows = Math.max(50, sh.getMaxRows() - (DATA_START - 1));
  const months = sh.getRange(DATA_START, 1, maxRows, 1).getValues().flat();
  let lastIdx = -1;
  for (let i = months.length - 1; i >= 0; i--) {
    if (months[i] !== '' && months[i] != null) { lastIdx = i; break; }
  }

  const out = [['Month', 'Money Added ($)', 'Money Used ($)']];
  if (lastIdx >= 0) {
    const rows = lastIdx + 1;
    const adds = sh.getRange(DATA_START, 2, rows, 1).getValues();
    const used = sh.getRange(DATA_START, 3, rows, 1).getValues();
    for (let i = 0; i < rows; i++) {
      out.push([String(months[i] || ''), Number(adds[i][0] || 0), Number(used[i][0] || 0)]);
    }
  } else {
    out.push(['', 0, 0]);
  }

  const snapColIndex = (function(a1){
    let c=0,s=a1.toUpperCase().replace(/[^A-Z]/g,'');
    for (let i=0;i<s.length;i++) c=c*26+(s.charCodeAt(i)-64);
    return c;
  })(SNAP_COL);

  const needRows = HEADER_ROW + out.length - 1;
  if (sh.getMaxRows() < needRows) sh.insertRowsAfter(sh.getMaxRows(), needRows - sh.getMaxRows());

  sh.getRange(HEADER_ROW, snapColIndex, Math.max(1000, out.length + 20), SNAP_COLS).clearContent();
  sh.getRange(HEADER_ROW, snapColIndex, out.length, SNAP_COLS).setValues(out);
  ['AQ','AR','AS'].forEach(function(c){ sh.hideColumn(sh.getRange(c+':'+c)); });

  const dt = Charts.newDataTable()
    .addColumn(Charts.ColumnType.STRING, 'Month')
    .addColumn(Charts.ColumnType.NUMBER, 'Money Added ($)')
    .addColumn(Charts.ColumnType.NUMBER, 'Money Used ($)');

  for (let i=1;i<out.length;i++){
    dt.addRow([String(out[i][0]), Number(out[i][1]), Number(out[i][2])]);
  }

  const chart = Charts.newColumnChart()
    .setDataTable(dt.build())
    .setTitle('Monthly Summary ($)')
    .setDimensions(860, 380)
    .setLegendPosition(Charts.Position.BOTTOM)
    .setOption('fontName', 'Inter, Segoe UI, Roboto, Arial, sans-serif')
    .setOption('backgroundColor', { fill: RP.surface })
    .setOption('chartArea', { left: 40, top: 24, width: '88%', height: '72%', backgroundColor: RP.surface })
    .setOption('titleTextStyle', { color: RP.love, fontSize: 16, bold: true })
    .setOption('legend', { textStyle: { color: RP.text, fontSize: 12 } })
    .setOption('hAxis', { textStyle: { color: RP.text, fontSize: 12 }, slantedText: false, gridlines: { color: RP.overlay }, baselineColor: RP.overlay })
    .setOption('vAxis', { textStyle: { color: RP.text, fontSize: 12 }, gridlines: { color: RP.overlay }, minorGridlines: { color: RP.overlay }, baselineColor: RP.overlay, format: '$#,##0' })
    .setOption('colors', [RP.foam, RP.love])
    .setOption('bar', { groupWidth: '60%' })
    .build();

  const blob = chart.getAs('image/png');
  removeAllVisuals_(sh);
  sh.insertImage(blob, 2, 15); // B15
}

/***** PEOPLE — REMOVE PERSON (UI + full cleanup) *****/
function promptAndRemovePerson() {
  const ui = SpreadsheetApp.getUi();

  const name = ui.prompt(
    'Remove Person',
    'Enter name (e.g. "Jane Doe" or "Doe, Jane"):',
    ui.ButtonSet.OK_CANCEL
  );
  if (name.getSelectedButton() !== ui.Button.OK) return;

  let raw = String(name.getResponseText() || '').trim();
  if (!raw) return ui.alert('No name entered.');

  // Handle both “Doe, Jane” and “Jane Doe” formats
  let [last, first] = raw.includes(',')
    ? raw.split(',').map(s => s.trim())
    : raw.split(' ').length > 1
      ? [raw.split(' ').pop(), raw.split(' ').slice(0, -1).join(' ')]
      : [raw, ''];

  if (!last) return ui.alert('Missing last name.');

  const confirm = ui.alert(
    'Confirm Deletion',
    `Delete ALL records for: ${cap_(last)}, ${cap_(first || '(no first)')}?`,
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const removed = removePersonEverywhere_({ last, first });
  ui.alert(`✅ Removed ${removed} record(s) for ${cap_(last)}, ${cap_(first)} from all sheets and rebuilt summaries.`);
}
/***** Deletes all rows for a person across every related sheet (fuzzy match) *****/
function removePersonEverywhere_(obj, options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targets = [
    RESPONSES_SHEET,
    RESPONSES_MERGED,
    DONOR_SUMMARY,
    FORMLOG,
    RAWFORMLOG
  ];

  const norm = s => String(s || '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '')
    .trim();

  const first = norm(obj.first);
  const last = norm(obj.last);

  let total = 0;

  targets.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return;

    const headers = data[0];
    const idxL = findHeaderIndex(headers, ['last', 'lastname']);
    const idxF = findHeaderIndex(headers, ['first', 'firstname']);
    if (idxL === -1 || idxF === -1) return;

    const keep = [headers];
    for (let i = 1; i < data.length; i++) {
      const rowLast  = norm(data[i][idxL]);
      const rowFirst = norm(data[i][idxF]);
      const match =
        rowLast === last &&
        (rowFirst === first || rowFirst.startsWith(first) || first.startsWith(rowFirst));
      if (!match) keep.push(data[i]);
      else total++;
    }

    sh.clearContents();
    sh.getRange(1, 1, keep.length, headers.length).setValues(keep);
  });

  rebuildFromResponses_();

  const silent = options && options.silent;
  if (!silent) {
    SpreadsheetApp.getUi().alert(`✅ Removed ${total} record(s) for ${cap_(first)}, ${cap_(last)}.`);
  }
  return total;
}

// Delete a client everywhere with a simple passcode gate for kiosk use.
function removePersonWithPassword(obj) {
  const pass = String(obj && obj.password ? obj.password : '').trim();
  if (pass !== '1944') {
    return { ok:false, msg:'Incorrect password.' };
  }

  const last  = cap_(obj && obj.last ? obj.last : '');
  const first = cap_(obj && obj.first ? obj.first : '');
  if (!last || !first) {
    return { ok:false, msg:'Missing name.' };
  }

  const removed = removePersonEverywhere_({ last, first }, { silent:true });
  if (removed === 0) {
    return { ok:false, removed: 0, msg:'No matching records found.' };
  }

  return { ok:true, removed: removed, msg:`Removed ${removed} record(s).` };
}

/***** DUPLICATE REPORT (no CUSTOM_FORMULA; safe in Apps Script) *****/
function showDupeReport(){
  const WINDOW_MINUTES = 10; // look for duplicates within last 10 minutes
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(RESPONSES_SHEET);
  if (!src){ SpreadsheetApp.getUi().alert('Missing sheet: ' + RESPONSES_SHEET); return; }

  const tz = Session.getScriptTimeZone();
  const data = src.getDataRange().getValues();
  if (data.length < 2){ SpreadsheetApp.getUi().alert('No data.'); return; }

  const h = data[0];
  const idxTS    = findHeaderIndex(h, ['timestamp','time','date']);
  const idxLast  = findHeaderIndex(h, ['last','lastname']);
  const idxFirst = findHeaderIndex(h, ['first','firstname']);

  const cutoff = Date.now() - WINDOW_MINUTES*60*1000;
  const bucket = new Map(); // key -> {count, firstAt, lastAt}

  for (let i=1;i<data.length;i++){
    const row = data[i];
    const ts = row[idxTS] ? new Date(row[idxTS]).getTime() : 0;
    if (!ts || ts < cutoff) continue;
    const last = String(row[idxLast]||'').trim();
    const first = String(row[idxFirst]||'').trim();
    if (!last || !first) continue;
    const key = last.toLowerCase()+'|'+first.toLowerCase();
    const cur = bucket.get(key) || {count:0, firstAt:ts, lastAt:ts, last:last, first:first};
    cur.count += 1;
    cur.firstAt = Math.min(cur.firstAt, ts);
    cur.lastAt  = Math.max(cur.lastAt, ts);
    bucket.set(key, cur);
  }

  const out = [['Last','First','Count (last 10 min)','First Seen','Last Seen']];
  const vals = Array.from(bucket.values());
  for (let i=0;i<vals.length;i++){
    const v = vals[i];
    if (v.count > 1){
      out.push([
        v.last, v.first, v.count,
        Utilities.formatDate(new Date(v.firstAt), tz, 'yyyy-MM-dd HH:mm:ss'),
        Utilities.formatDate(new Date(v.lastAt),  tz, 'yyyy-MM-dd HH:mm:ss')
      ]);
    }
  }

  const sh = ss.getSheetByName('Dupe Report') || ss.insertSheet('Dupe Report');
  sh.clearContents();
  sh.getRange(1,1,out.length, out[0].length).setValues(out);
  sh.autoResizeColumns(1, out[0].length);
  ss.setActiveSheet(sh);
}

/***** UTIL *****/
function Rebuild_Now(){
  let ui = null;
  try {
    ui = typeof SpreadsheetApp.getUi === 'function' ? SpreadsheetApp.getUi() : null;
  } catch (_err) {
    ui = null;
  }

  try {
    rebuildFromResponses_();
    if (ui) ui.alert('Summaries rebuilt successfully.');
  } catch (err) {
    if (ui) ui.alert('Rebuild failed: ' + (err && err.message ? err.message : err));
    throw err;
  }
}

function Update_Dashboard(){
  buildOrUpdateDashboard_();
}  // closes Update_Dashboard

/***** SAVE SIGNED ENTRY (for HTML signature pad) *****/
function saveSignedEntry(entry) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Signatures') || ss.insertSheet('Signatures');

  // Add headers if the sheet is new
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Timestamp', 'Last Name', 'First Name', 'Clerk', 'Signature (Base64)']);
  }

  // Add a new record
  sh.appendRow([
    new Date(),
    entry.last || '',
    entry.first || '',
    entry.clerk || '',
    normalizeSignature_(entry.signature)
  ]);

  return { success: true };
}
