/**
 * SE Practice Partners · Google Apps Script web app
 * Bound to the spreadsheet that collects the form responses.
 * Reads availability from the responses sheet, stores bookings in a "Bookings" sheet.
 */

const CONFIG = {
  FORM_URL: 'https://forms.gle/TUO_LINK',   // link to the Google Form (Send → link)
  APP_URL: '',                                // short link to the app (e.g. https://pratica.mnesti.it); empty = Apps Script URL
  RESPONSES_SHEET: '',      // leave empty to auto-detect the sheet with the form responses
  BOOKINGS_SHEET: 'Bookings',
  TIMEZONE: 'Europe/Rome',
  MIN_GAP_MINUTES: 90,        // a person can't have two sessions closer than this
};

const BOOKING_HEADERS = ['ID', 'Created', 'Proposed by', 'With', 'Date', 'Time', 'Format', 'Note', 'Status'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SE Practice Partners')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- Data for the page ---------- */

function getData() {
  return { people: readPeople_(), bookings: readBookings_(), today: today_(), formUrl: CONFIG.FORM_URL,
           appUrl: CONFIG.APP_URL || ScriptApp.getService().getUrl() };
}

function readPeople_() {
  const sheet = responsesSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const col = mapColumns_(values[0]);
  const byName = {};

  values.slice(1).forEach(function (row) {
    const name = String(row[col.name] || '').trim();
    if (!name) return;
    byName[key_(name)] = {               // later submissions replace earlier ones
      name: name,
      region: cell_(row, col.region),
      format: cell_(row, col.format),
      days: list_(row, col.days),
      slots: list_(row, col.slots),
      frequency: cell_(row, col.frequency),
      setup: cell_(row, col.setup),
      languages: list_(row, col.languages),
      phone: cell_(row, col.phone),
      notes: cell_(row, col.notes),
      photo: photo_(cell_(row, col.photo)),
    };
  });

  return Object.keys(byName).map(function (k) { return byName[k]; })
    .sort(function (a, b) { return a.name.localeCompare(b.name); });
}

/* Photo uploaded through the form: the cell holds a Drive link. The script runs as the owner,
   so it reads a small thumbnail and sends it as an image; the files stay private in Drive. */
function photo_(link) {
  const m = String(link || '').match(/[?&]id=([\w-]+)|\/d\/([\w-]+)/);
  if (!m) return '';
  const id = m[1] || m[2];
  const cache = CacheService.getScriptCache();
  const hit = cache.get('ph_' + id);
  if (hit !== null) return hit;
  var uri = '';
  try {
    const thumb = DriveApp.getFileById(id).getThumbnail();
    if (thumb) uri = 'data:' + (thumb.getContentType() || 'image/png') + ';base64,' + Utilities.base64Encode(thumb.getBytes());
  } catch (e) { uri = ''; }
  if (uri.length < 95000) cache.put('ph_' + id, uri, 21600); else uri = '';   // cache 6 h
  return uri;
}

function readBookings_() {
  const sheet = bookingsSheet_();
  const values = sheet.getDataRange().getValues();
  const now = today_();
  return values.slice(1)
    .map(function (r) {
      return {
        id: String(r[0]),
        createdAt: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
        by: String(r[2]),
        with: String(r[3]),
        date: asDate_(r[4]),
        time: asTime_(r[5]),
        format: String(r[6]),
        note: String(r[7]),
        status: String(r[8] || 'proposed'),
      };
    })
    .filter(function (b) { return b.id && b.date >= now && b.status !== 'cancelled'; })
    .sort(function (a, b) { return (a.date + a.time).localeCompare(b.date + b.time); });
}

/* ---------- Actions ---------- */

function createBooking(input) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const people = readPeople_();
    const names = {};
    people.forEach(function (p) { names[key_(p.name)] = p.name; });

    const by = names[key_(input.by)];
    const withName = names[key_(input.with)];
    if (!by || !withName) throw new Error('ERR_PICK_TWO');
    if (by === withName) throw new Error('ERR_SELF');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(input.date || '')) throw new Error('ERR_DATE');
    if (!/^\d{2}:\d{2}$/.test(input.time || '')) throw new Error('ERR_TIME');
    if (input.date < today_()) throw new Error('ERR_PAST');

    const start = minutes_(input.time);
    const clash = readBookings_().find(function (b) {
      const involved = [b.by, b.with].some(function (n) { return n === by || n === withName; });
      return involved && b.date === input.date && Math.abs(minutes_(b.time) - start) < CONFIG.MIN_GAP_MINUTES;
    });
    if (clash) {
      throw new Error(['ERR_CLASH', clash.by, clash.with, clash.time].join('|'));
    }

    bookingsSheet_().appendRow([
      Utilities.getUuid().slice(0, 8),
      new Date(),
      by,
      withName,
      input.date,
      input.time,
      String(input.format || '').slice(0, 40),
      String(input.note || '').slice(0, 300),
      'proposed',
    ]);
    return getData();
  } finally {
    lock.releaseLock();
  }
}

function confirmBooking(id, name) {
  return updateStatus_(id, name, 'confirmed', function (b) { return key_(b.with) === key_(name); },
    'ERR_ONLY_INVITED');
}

function cancelBooking(id, name) {
  return updateStatus_(id, name, 'cancelled', function (b) {
    return key_(b.with) === key_(name) || key_(b.by) === key_(name);
  }, 'ERR_ONLY_PAIR');
}

function updateStatus_(id, name, status, allowed, refusal) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = bookingsSheet_();
    const values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]) !== String(id)) continue;
      if (!allowed({ by: values[i][2], with: values[i][3] })) throw new Error(refusal);
      sheet.getRange(i + 1, 9).setValue(status);
      return getData();
    }
    throw new Error('ERR_NOT_FOUND');
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Sheets ---------- */

function responsesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (CONFIG.RESPONSES_SHEET) return ss.getSheetByName(CONFIG.RESPONSES_SHEET);
  const found = ss.getSheets().find(function (s) {
    if (s.getName() === CONFIG.BOOKINGS_SHEET || s.getLastColumn() === 0) return false;
    const header = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
    return header.some(function (h) { return /^name\b/i.test(String(h).trim()); });
  });
  if (!found) throw new Error('No sheet with a "Name" column was found.');
  return found;
}

function bookingsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.BOOKINGS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.BOOKINGS_SHEET);
    sheet.getRange(1, 1, 1, BOOKING_HEADERS.length).setValues([BOOKING_HEADERS]).setFontWeight('bold');
    sheet.getRange('E:F').setNumberFormat('@');   // keep date and time as plain text
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function mapColumns_(header) {
  const h = header.map(function (x) { return String(x).trim().toLowerCase(); });
  function find(test) { return h.findIndex(test); }
  return {
    name: find(function (x) { return /^name\b/.test(x); }),
    region: find(function (x) { return x.indexOf('region') !== -1 || x.indexOf('city') !== -1; }),
    format: find(function (x) { return x.indexOf('session format') !== -1; }),
    days: find(function (x) { return x.indexOf('available days') !== -1; }),
    slots: find(function (x) { return x.indexOf('time slot') !== -1; }),
    frequency: find(function (x) { return x.indexOf('how often') !== -1; }),
    setup: find(function (x) { return x.indexOf('setup') !== -1; }),
    languages: find(function (x) { return x.indexOf('language') !== -1; }),
    phone: find(function (x) { return x.indexOf('phone') !== -1; }),
    notes: find(function (x) { return /^notes\b/.test(x); }),
    photo: find(function (x) { return x.indexOf('photo') !== -1; }),
  };
}

/* ---------- Helpers ---------- */

function cell_(row, i) { return i < 0 ? '' : String(row[i] || '').trim(); }
function list_(row, i) { return cell_(row, i).split(',').map(function (s) { return s.trim(); }).filter(String); }
function key_(s) { return String(s || '').trim().toLowerCase(); }
function today_() { return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd'); }
function minutes_(t) { const p = String(t).split(':'); return Number(p[0]) * 60 + Number(p[1]); }
function asDate_(v) { return v instanceof Date ? Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyy-MM-dd') : String(v); }
function asTime_(v) { return v instanceof Date ? Utilities.formatDate(v, CONFIG.TIMEZONE, 'HH:mm') : String(v); }
