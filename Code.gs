// ==============================
// Heart to Hand – Backend (GAS)
// ==============================

const HH_SHEET_NAME = 'Heart to Hand Log';
const HH_HEADERS = [
  'Timestamp',
  'Last Name',
  'First Name',
  'Phone',
  'Number in Family',
  'Adults',
  'Children',
  'Items',
  'Clerk Initials',
  'Signature (Base64 Image)',
];

// ---------- Helpers ----------
function _getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function _ensureSheetWithHeaders_(name, headers) {
  const ss = _getSS();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  let mismatch = false;

  headers.forEach((h, i) => {
    if (existing[i] !== h) mismatch = true;
  });

  if (mismatch) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  return sheet;
}

function _withLock_(label, fn) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(20000)) {
    throw new Error('Could not obtain lock: ' + label);
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// ---------- Setup ----------
function setupHeartToHand_() {
  _ensureSheetWithHeaders_(HH_SHEET_NAME, HH_HEADERS);
}

function onOpen() {
  setupHeartToHand_();
}

function onInstall() {
  onOpen();
}

// ---------- Web App ----------
function doGet() {
  setupHeartToHand_();
  return HtmlService.createHtmlOutputFromFile('HeartToHand')
    .setTitle('Heart to Hand')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------- Save Entry ----------
function saveHHEntry(entry) {
  if (!entry) throw new Error('Missing entry data');

  ['lastName', 'firstName', 'items', 'clerkInitials'].forEach((f) => {
    if (!entry[f] || !String(entry[f]).trim()) {
      throw new Error('Missing required field: ' + f);
    }
  });

  const sheet = _ensureSheetWithHeaders_(HH_SHEET_NAME, HH_HEADERS);

  const idx = {
    timestamp: 0,
    lastName: 1,
    firstName: 2,
    phone: 3,
    numberInFamily: 4,
    adults: 5,
    children: 6,
    items: 7,
    clerkInitials: 8,
    signatureImage: 9,
  };

  return _withLock_('saveHHEntry', () => {
    const adults = Number(entry.adults) || 0;
    const children = Number(entry.children) || 0;

    const row = new Array(HH_HEADERS.length).fill('');

    row[idx.timestamp] = entry.date
      ? new Date(entry.date)
      : new Date();

    row[idx.lastName] = entry.lastName;
    row[idx.firstName] = entry.firstName;
    row[idx.phone] = entry.phone || '';
    row[idx.adults] = adults || '';
    row[idx.children] = children || '';
    row[idx.numberInFamily] = adults + children || '';
    row[idx.items] = entry.items;
    row[idx.clerkInitials] = entry.clerkInitials;
    row[idx.signatureImage] = entry.signatureImage || '';

    sheet.appendRow(row);
    return { ok: true };
  });
}

// ---------- History ----------
function getHHHistory(lastName, firstName) {
  const ln = String(lastName || '').toLowerCase().trim();
  const fn = String(firstName || '').toLowerCase().trim();
  if (!ln || !fn) throw new Error('Both names required');

  const sheet = _ensureSheetWithHeaders_(HH_SHEET_NAME, HH_HEADERS);
  const rows = sheet.getDataRange().getValues().slice(1);

  const matches = rows
    .filter((r) => {
      return (
        String(r[1]).toLowerCase() === ln &&
        String(r[2]).toLowerCase().includes(fn)
      );
    })
    .map((r) => ({
      timestamp: r[0],
      lastName: r[1],
      firstName: r[2],
      phone: r[3],
      numberInFamily: r[4],
      adults: r[5],
      children: r[6],
      items: r[7],
      clerkInitials: r[8],
      signatureImage: r[9],
    }))
    .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

  return matches;
}

// ---------- Clients ----------
function getAllHHClients() {
  const sheet = _ensureSheetWithHeaders_(HH_SHEET_NAME, HH_HEADERS);
  const rows = sheet.getDataRange().getValues().slice(1);
  const map = new Map();

  rows.forEach((r) => {
    const ln = String(r[1] || '').trim();
    const fn = String(r[2] || '').trim();
    if (!ln || !fn) return;
    map.set(`${ln.toLowerCase()}|${fn.toLowerCase()}`, {
      lastName: ln,
      firstName: fn,
    });
  });

  return [...map.values()].sort((a, b) =>
    a.lastName.localeCompare(b.lastName) ||
    a.firstName.localeCompare(b.firstName)
  );
}
