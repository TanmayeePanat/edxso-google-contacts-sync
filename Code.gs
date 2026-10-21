/**
 * Two-way sync: Google Sheets <-> Google Contacts
 *
 * Sheet layout (first tab):
 *   A = Name | B = Phone | C = Contact ID (auto) | D = Sync Key (auto)
 *
 * Requires the "Google People API" advanced service (identifier: People).
 */

const FIELDS = 'names,phoneNumbers';
const SYNC_EVERY_MINUTES = 1; // allowed: 1, 5, 10, 15, 30

/* ---------- Menu ---------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Contacts Sync')
    .addItem('Sync now', 'syncAll')
    .addItem('Set up automatic sync', 'setup')
    .addToUi();
}

/* ---------- One-time setup ---------- */

function setup() {
  const sheet = getSheet_();
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Name', 'Phone', 'Contact ID (auto)', 'Sync Key (auto)']);
  }
  sheet.getRange('B:B').setNumberFormat('@'); // keep "+" and leading zeros
  sheet.hideColumns(3, 2);                    // hide helper columns

  // Remove old triggers so we never create duplicates
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // Contacts -> Sheet (Google Contacts has no "on change" trigger, so we poll)
  ScriptApp.newTrigger('syncAll')
    .timeBased()
    .everyMinutes(SYNC_EVERY_MINUTES)
    .create();

  // Sheet -> Contacts (runs right after you edit the sheet)
  ScriptApp.newTrigger('syncAll')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  syncAll();
}

/* ---------- Main entry point ---------- */

function syncAll() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return; // another sync is running
  try {
    runSync_();
  } finally {
    lock.releaseLock();
  }
}

function runSync_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  const rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 4).getValues() : [];

  // Fresh (consistent) copy of every contact already linked to a sheet row
  const linkedIds = rows.map(r => String(r[2]).trim()).filter(Boolean);
  const linked = fetchLinked_(linkedIds);

  const knownIds = new Set(linkedIds);
  const rowsToDelete = [];

  // ---- Step 1: go through every sheet row ----
  rows.forEach((r, i) => {
    const rowNum = i + 2;
    const name = String(r[0]).trim();
    const phone = String(r[1]).trim();
    const id = String(r[2]).trim();
    const storedKey = String(r[3]);

    // New row -> create contact
    if (!id) {
      if (!name && !phone) return;
      const created = People.People.createContact(buildPerson_(name, phone));
      knownIds.add(created.resourceName);
      sheet.getRange(rowNum, 3, 1, 2)
        .setValues([[created.resourceName, makeKey_(name, phone)]]);
      return;
    }

    const person = linked[id];

    // Contact was deleted in Google Contacts -> remove the row
    if (!person) {
      rowsToDelete.push(rowNum);
      return;
    }

    const sheetKey = makeKey_(name, phone);
    const contact = readPerson_(person);
    const contactKey = makeKey_(contact.name, contact.phone);

    if (sheetKey !== storedKey) {
      // Sheet changed -> update the contact (sheet wins if both changed)
      const body = buildPerson_(name, phone);
      body.etag = person.etag;
      People.People.updateContact(body, id, { updatePersonFields: FIELDS });
      sheet.getRange(rowNum, 4).setValue(sheetKey);
    } else if (contactKey !== storedKey) {
      // Contact changed -> update the sheet
      sheet.getRange(rowNum, 1, 1, 2).setValues([[contact.name, contact.phone]]);
      sheet.getRange(rowNum, 4).setValue(contactKey);
    }
  });

  // Delete rows bottom-up so row numbers stay valid
  rowsToDelete.reverse().forEach(n => sheet.deleteRow(n));

  // ---- Step 2: contacts that are not in the sheet yet -> add rows ----
  const newRows = [];
  const params = { personFields: FIELDS, pageSize: 1000 };
  do {
    const res = People.People.Connections.list('people/me', params);
    (res.connections || []).forEach(p => {
      if (knownIds.has(p.resourceName)) return;
      const c = readPerson_(p);
      if (!c.name && !c.phone) return;
      newRows.push([c.name, c.phone, p.resourceName, makeKey_(c.name, c.phone)]);
    });
    params.pageToken = res.nextPageToken;
  } while (params.pageToken);

  if (newRows.length) {
    const start = sheet.getLastRow() + 1;
    sheet.getRange(start, 1, newRows.length, 4).setValues(newRows);
  }
}

/* ---------- Helpers ---------- */

function getSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

// Returns { resourceName: person } for contacts that still exist
function fetchLinked_(ids) {
  const map = {};
  for (let i = 0; i < ids.length; i += 50) {
    const res = People.People.getBatchGet({
      resourceNames: ids.slice(i, i + 50),
      personFields: FIELDS
    });
    (res.responses || []).forEach(r => {
      const p = r.person;
      if (p && !(p.metadata && p.metadata.deleted)) {
        map[r.requestedResourceName] = p;
      }
    });
  }
  return map;
}

function buildPerson_(name, phone) {
  return {
    names: [{ givenName: name }],
    phoneNumbers: phone ? [{ value: phone }] : []
  };
}

function readPerson_(p) {
  return {
    name: p.names && p.names[0] ? (p.names[0].displayName || '') : '',
    phone: p.phoneNumbers && p.phoneNumbers[0] ? (p.phoneNumbers[0].value || '') : ''
  };
}

// Fingerprint of "what both sides looked like at the last sync".
// Phone is reduced to digits so "+91 98765-43210" equals "919876543210".
function makeKey_(name, phone) {
  return String(name).trim() + '|' + String(phone).replace(/\D/g, '');
}
