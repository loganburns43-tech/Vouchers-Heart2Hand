const CLIENT_SHEET_NAME = 'DFS Clients';
const LOG_SHEET_NAME = 'Dress for Success Log';

const CLIENT_HEADERS = [
  'PersonID',
  'First Name',
  'Last Name',
  'Phone',
  'Notes',
  'Created Timestamp',
  'Total Visits',
  'Total Items Received',
];

const CLIENT_INDEX = {
  PERSON_ID: 0,
  FIRST_NAME: 1,
  LAST_NAME: 2,
  PHONE: 3,
  NOTES: 4,
  CREATED_TS: 5,
  TOTAL_VISITS: 6,
  TOTAL_ITEMS: 7,
};

const LOG_HEADERS = [
  'Timestamp',
  'Visit ID',
  'Visit Date',
  'PersonID',
  'First Name',
  'Last Name',
  'Phone',
  'Reason / Visit Type',
  'Items (Checkbox Summary)',
  'Items (Free Text)',
  'Clerk Initials',
  'Signature (Base64 PNG)',
  'Last Edited Timestamp',
  'Last Edited By',
];

const LOG_INDEX = {
  TIMESTAMP: 0,
  VISIT_ID: 1,
  VISIT_DATE: 2,
  PERSON_ID: 3,
  FIRST_NAME: 4,
  LAST_NAME: 5,
  PHONE: 6,
  REASON: 7,
  ITEMS_SUMMARY: 8,
  ITEMS_TEXT: 9,
  CLERK: 10,
  SIGNATURE: 11,
  LAST_EDITED_TS: 12,
  LAST_EDITED_BY: 13,
};

function _normalizeString(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function doGet() {
  initializeSheets();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dress for Success')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Dress for Success')
    .addItem('Initialize Sheets', 'initializeSheets')
    .addToUi();
}

function initializeSheets() {
  _ensureSheetWithHeaders_(CLIENT_SHEET_NAME, CLIENT_HEADERS);
  _ensureSheetWithHeaders_(LOG_SHEET_NAME, LOG_HEADERS);
}

function _getSS() {
  return SpreadsheetApp.getActive();
}

function _getSheet(name) {
  const ss = _getSS();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function _ensureSheetWithHeaders_(name, headers) {
  const sheet = _getSheet(name);
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  const existing = headerRange.getValues()[0];
  const needsHeaders = existing.some((cell, idx) => cell !== headers[idx]);

  if (needsHeaders) {
    headerRange.setValues([headers]);
  }

  headerRange
    .setBackground('#1f1d2e')
    .setFontColor('#e0def4')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true, '#403d52', SpreadsheetApp.BorderStyle.SOLID);

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

function getClients() {
  initializeSheets();
  _recalculateClientStats();
  const sheet = _getSheet(CLIENT_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  const values = sheet.getRange(2, 1, lastRow - 1, CLIENT_HEADERS.length).getValues();
  return values
    .filter(row => row[0])
    .map(row => ({
      personId: _normalizeString(row[CLIENT_INDEX.PERSON_ID]),
      firstName: _normalizeString(row[CLIENT_INDEX.FIRST_NAME]),
      lastName: _normalizeString(row[CLIENT_INDEX.LAST_NAME]),
      phone: _normalizeString(row[CLIENT_INDEX.PHONE]),
      notes: row[CLIENT_INDEX.NOTES],
      totalVisits: Number(row[CLIENT_INDEX.TOTAL_VISITS]) || 0,
      totalItems: Number(row[CLIENT_INDEX.TOTAL_ITEMS]) || 0,
    }));
}

function addClient(data) {
  const firstName = (data.firstName || '').trim();
  const lastName = (data.lastName || '').trim();
  const phone = (data.phone || '').trim();
  const notes = (data.notes || '').trim();

  if (!firstName || !lastName || !phone) {
    throw new Error('First Name, Last Name, and Phone are required to add a client.');
  }

  initializeSheets();
  const sheet = _getSheet(CLIENT_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const existing = sheet.getRange(2, 1, lastRow - 1, CLIENT_HEADERS.length).getValues();
    const duplicate = existing.find(row =>
      String(row[1]).trim().toLowerCase() === firstName.toLowerCase() &&
      String(row[2]).trim().toLowerCase() === lastName.toLowerCase() &&
      String(row[3]).trim() === phone
    );
    if (duplicate) {
      return {
        personId: duplicate[0],
        firstName,
        lastName,
        phone,
        notes,
      };
    }
  }

  const personId = _generatePersonId(sheet);
  const createdTimestamp = new Date();
  sheet.appendRow([personId, firstName, lastName, phone, notes, createdTimestamp, 0, 0]);

  return {
    personId,
    firstName,
    lastName,
    phone,
    notes,
  };
}

function _generatePersonId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 'DFS-0001';
  }
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  return _generateSequentialId('DFS-', ids);
}

function _generateVisitId(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 'DFS-V-0001';
  }
  const ids = sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
  return _generateSequentialId('DFS-V-', ids);
}

function _generateSequentialId(prefix, values) {
  const maxNumber = values
    .map(id => {
      if (!id || typeof id !== 'string' || !id.startsWith(prefix)) return 0;
      const numeric = parseInt(id.replace(prefix, ''), 10);
      return Number.isFinite(numeric) ? numeric : 0;
    })
    .reduce((max, num) => Math.max(max, num), 0);

  const nextNumber = maxNumber + 1;
  return `${prefix}${String(nextNumber).padStart(4, '0')}`;
}

function _countItems(itemsString) {
  if (!itemsString) return 0;
  return itemsString
    .split(',')
    .map(value => _normalizeString(value))
    .filter(Boolean)
    .length;
}

function _recalculateClientStats() {
  const logSheet = _getSheet(LOG_SHEET_NAME);
  const logLastRow = logSheet.getLastRow();
  const statsMap = new Map();

  if (logLastRow >= 2) {
    const logValues = logSheet.getRange(2, 1, logLastRow - 1, LOG_HEADERS.length).getValues();
    logValues.forEach(row => {
      const personId = _normalizeString(row[LOG_INDEX.PERSON_ID]);
      if (!personId) return;
      const itemsString = _normalizeString(row[LOG_INDEX.ITEMS_SUMMARY]) || _normalizeString(row[LOG_INDEX.ITEMS_TEXT]);
      const existing = statsMap.get(personId) || { visits: 0, items: 0 };
      statsMap.set(personId, {
        visits: existing.visits + 1,
        items: existing.items + _countItems(itemsString),
      });
    });
  }

  const clientSheet = _getSheet(CLIENT_SHEET_NAME);
  const clientLastRow = clientSheet.getLastRow();
  if (clientLastRow < 2) {
    return statsMap;
  }

  const range = clientSheet.getRange(2, 1, clientLastRow - 1, CLIENT_HEADERS.length);
  const clientValues = range.getValues();
  const updated = clientValues.map(row => {
    const normalizedId = _normalizeString(row[CLIENT_INDEX.PERSON_ID]);
    const stats = statsMap.get(normalizedId) || { visits: 0, items: 0 };
    row[CLIENT_INDEX.TOTAL_VISITS] = stats.visits;
    row[CLIENT_INDEX.TOTAL_ITEMS] = stats.items;
    return row;
  });

  range.setValues(updated);
  return statsMap;
}

function getClientHistory(personId) {
  initializeSheets();
  const normalizedPersonId = _normalizeString(personId);
  const normalizedPersonIdLower = normalizedPersonId.toLowerCase();
  const clientSheet = _getSheet(CLIENT_SHEET_NAME);
  const clientLastRow = clientSheet.getLastRow();
  const clientValues = clientLastRow >= 2
    ? clientSheet.getRange(2, 1, clientLastRow - 1, CLIENT_HEADERS.length).getValues()
    : [];

  const matchingClient = clientValues.find(row => _normalizeString(row[0]) === normalizedPersonId);
  const targetFirst = matchingClient ? _normalizeString(matchingClient[1]).toLowerCase() : '';
  const targetLast = matchingClient ? _normalizeString(matchingClient[2]).toLowerCase() : '';
  const targetPhone = matchingClient ? _normalizeString(matchingClient[3]) : '';

  const sheet = _getSheet(LOG_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }

  const values = sheet.getRange(2, 1, lastRow - 1, LOG_HEADERS.length).getValues();
  return values
    .filter(row => {
      const rowPersonId = _normalizeString(row[LOG_INDEX.PERSON_ID]);
      const rowPersonIdLower = rowPersonId.toLowerCase();
      const rowFirst = _normalizeString(row[LOG_INDEX.FIRST_NAME]).toLowerCase();
      const rowLast = _normalizeString(row[LOG_INDEX.LAST_NAME]).toLowerCase();
      const rowPhone = _normalizeString(row[LOG_INDEX.PHONE]);

      const matchesPersonId = normalizedPersonId && rowPersonId && rowPersonIdLower === normalizedPersonIdLower;
      const matchesByDetails = targetFirst && targetLast && targetPhone &&
        rowFirst === targetFirst && rowLast === targetLast && rowPhone === targetPhone;
      const matchesByNameOnly = targetFirst && targetLast && rowFirst === targetFirst && rowLast === targetLast;

      return matchesPersonId || (!matchesPersonId && (matchesByDetails || matchesByNameOnly));
    })
    .map(row => ({
      timestamp: row[LOG_INDEX.TIMESTAMP],
      visitId: row[LOG_INDEX.VISIT_ID],
      visitDate: row[LOG_INDEX.VISIT_DATE] instanceof Date ? row[LOG_INDEX.VISIT_DATE].toISOString() : row[LOG_INDEX.VISIT_DATE],
      personId: row[LOG_INDEX.PERSON_ID],
      firstName: row[LOG_INDEX.FIRST_NAME],
      lastName: row[LOG_INDEX.LAST_NAME],
      phone: row[LOG_INDEX.PHONE],
      reason: row[LOG_INDEX.REASON],
      itemsSummary: row[LOG_INDEX.ITEMS_SUMMARY],
      itemsText: row[LOG_INDEX.ITEMS_TEXT],
      clerkInitials: row[LOG_INDEX.CLERK],
      signature: row[LOG_INDEX.SIGNATURE],
      lastEditedTimestamp: row[LOG_INDEX.LAST_EDITED_TS],
      lastEditedBy: row[LOG_INDEX.LAST_EDITED_BY],
    }))
    .sort((a, b) => new Date(b.visitDate || b.timestamp) - new Date(a.visitDate || a.timestamp));
}

function saveNewVisit(visitData) {
  const requiredFields = ['firstName', 'lastName', 'phone', 'itemsText', 'clerkInitials', 'signature'];
  requiredFields.forEach(field => {
    if (!visitData[field] || String(visitData[field]).trim() === '') {
      throw new Error(`Missing required field: ${field}`);
    }
  });

  if (!visitData.personId) {
    throw new Error('A client must be selected before saving a visit.');
  }

  initializeSheets();
  const sheet = _getSheet(LOG_SHEET_NAME);
  const timestamp = new Date();
  const visitId = _generateVisitId(sheet);
  const visitDateValue = visitData.visitDate ? new Date(visitData.visitDate) : new Date();

  const row = [
    timestamp,
    visitId,
    visitDateValue,
    visitData.personId,
    visitData.firstName,
    visitData.lastName,
    visitData.phone,
    visitData.reason || '',
    visitData.itemsSummary || '',
    visitData.itemsText,
    visitData.clerkInitials,
    visitData.signature,
    timestamp,
    visitData.clerkInitials,
  ];

  sheet.appendRow(row);
  _recalculateClientStats();
  return { visitId, timestamp };
}

function updateVisit(visitId, visitData) {
  initializeSheets();
  const sheet = _getSheet(LOG_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('No visits found to update.');
  }

  const values = sheet.getRange(2, 1, lastRow - 1, LOG_HEADERS.length).getValues();
  const targetIndex = values.findIndex(row => String(row[LOG_INDEX.VISIT_ID]) === String(visitId));

  if (targetIndex === -1) {
    throw new Error('Visit not found.');
  }

  const visitDateValue = visitData.visitDate ? new Date(visitData.visitDate) : new Date();
  const lastEdited = new Date();
  const rowNumber = targetIndex + 2;
  const existingRow = sheet.getRange(rowNumber, 1, 1, LOG_HEADERS.length).getValues()[0];

  const updatedRow = [
    existingRow[LOG_INDEX.TIMESTAMP],
    visitId,
    visitDateValue,
    visitData.personId || existingRow[LOG_INDEX.PERSON_ID],
    visitData.firstName,
    visitData.lastName,
    visitData.phone,
    visitData.reason || '',
    visitData.itemsSummary || '',
    visitData.itemsText,
    visitData.clerkInitials,
    visitData.signature,
    lastEdited,
    visitData.clerkInitials,
  ];

  sheet.getRange(rowNumber, 1, 1, LOG_HEADERS.length).setValues([updatedRow]);
  _recalculateClientStats();
  return { visitId, lastEdited };
}

/*************************************************************
  🟢 OFFLINE QUEUE + ONLINE SAVE ADAPTERS FOR THE NEW HTML
*************************************************************/

/**
 * clientName comes from the frontend as:
 *   "First Last (phone)"
 * This helper tries to split that out again.
 */
function _parseClientNameFromQueue_(clientName) {
  const raw = _normalizeString(clientName);
  if (!raw) {
    return { firstName: '', lastName: '', phone: '' };
  }

  let phone = '';
  let base = raw;
  const phoneMatch = raw.match(/\((.*)\)\s*$/);
  if (phoneMatch) {
    phone = phoneMatch[1].trim();
    base = raw.slice(0, phoneMatch.index).trim();
  }

  const parts = base.split(/\s+/).filter(Boolean);
  if (!parts.length) {
    return { firstName: '', lastName: '', phone };
  }

  const firstName = parts.shift();
  const lastName = parts.join(' ');
  return { firstName, lastName, phone };
}

/**
 * Find a client row by first/last/phone. Returns personId or ''.
 */
function _findClientPersonIdByNamePhone_(firstName, lastName, phone) {
  const sheet = _getSheet(CLIENT_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return '';

  const values = sheet.getRange(2, 1, lastRow - 1, CLIENT_HEADERS.length).getValues();
  const targetFirst = _normalizeString(firstName).toLowerCase();
  const targetLast  = _normalizeString(lastName).toLowerCase();
  const targetPhone = _normalizeString(phone);

  const match = values.find(row =>
    _normalizeString(row[CLIENT_INDEX.FIRST_NAME]).toLowerCase() === targetFirst &&
    _normalizeString(row[CLIENT_INDEX.LAST_NAME]).toLowerCase() === targetLast &&
    _normalizeString(row[CLIENT_INDEX.PHONE]) === targetPhone
  );

  return match ? _normalizeString(match[CLIENT_INDEX.PERSON_ID]) : '';
}

/**
 * Core helper that takes a queue-style entry from the HTML and
 * actually writes a log row.
 *
 * entry shape (from the HTML):
 * {
 *   id: visitId,
 *   timestamp: visitDate || visitTimestamp,
 *   clientName: "First Last (phone)",
 *   itemsProvided: "...",
 *   staffMember: "AB",
 *   notes: "Interview" // reason
 * }
 */
function _appendVisitFromQueueEntry_(entry, nowOpt) {
  initializeSheets();
  const logSheet = _getSheet(LOG_SHEET_NAME);
  const now = nowOpt || new Date();

  const clientBits = _parseClientNameFromQueue_(entry.clientName || '');
  const personId = _findClientPersonIdByNamePhone_(
    clientBits.firstName,
    clientBits.lastName,
    clientBits.phone
  );

  const visitDateValue = entry.timestamp ? new Date(entry.timestamp) : now;
  const visitId = entry.id || entry.visitId || _generateVisitId(logSheet);

  const row = [
    now,                        // Timestamp
    visitId,                    // Visit ID
    visitDateValue,             // Visit Date
    personId,                   // PersonID (looked up)
    clientBits.firstName,       // First Name
    clientBits.lastName,        // Last Name
    clientBits.phone,           // Phone
    entry.notes || '',          // Reason / Visit Type
    entry.itemsProvided || '',  // Items (Checkbox Summary)
    entry.itemsProvided || '',  // Items (Free Text)
    entry.staffMember || '',    // Clerk Initials
    '',                         // Signature
    now,                        // Last Edited Timestamp
    entry.staffMember || '',    // Last Edited By
  ];

  logSheet.appendRow(row);
  _recalculateClientStats();
  return { visitId: visitId, timestamp: now };
}

/**
 * This is what the HTML calls when ONLINE:
 * google.script.run.submitRecord(queuePayload)
 */
function submitRecord(entry) {
  return _appendVisitFromQueueEntry_(entry);
}

/**
 * Frontend sends queued visits here when back online:
 * google.script.run.syncQueuedRecords(queue)
 */
function syncQueuedDFSVisits(entries) {
  if (!entries || !entries.length) {
    return { success: true, savedCount: 0 };
  }

  const now = new Date();
  entries.forEach(entry => _appendVisitFromQueueEntry_(entry, now));

  return { success: true, savedCount: entries.length, savedAt: now };
}

/**
 * Adapter name the HTML expects: syncQueuedRecords(queue)
 */
function syncQueuedRecords(entries) {
  return syncQueuedDFSVisits(entries);
}
