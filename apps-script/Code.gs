/**
 * SE Practice Partners · Google Apps Script web app
 * Bound to the spreadsheet that collects the form responses.
 * Reads availability from the responses sheet, stores bookings in a "Bookings" sheet.
 */

const CONFIG = {
  FORM_URL: 'https://forms.gle/TUO_LINK',   // link to the Google Form (Send → link)
  APP_URL: 'https://se-practice.mnesti.it',                                // short link to the app (e.g. https://pratica.mnesti.it); empty = Apps Script URL
  RESPONSES_SHEET: '',      // leave empty to auto-detect the sheet with the form responses
  BOOKINGS_SHEET: 'Bookings',
  TIMEZONE: 'Europe/Rome',
  MIN_GAP_MINUTES: 90,        // a person can't have two sessions closer than this
  SEND_EMAILS: true,          // email notifications for proposals, confirmations and cancellations
  SENDER_NAME: 'SE Practice Partners',
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
  // Email addresses never leave the server: the page only learns whether someone can be notified.
  const people = readPeople_().map(function (p) {
    const out = Object.assign({}, p);
    out.hasEmail = !!p.email;
    delete out.email;
    return out;
  });
  return { people: people, bookings: readBookings_(), today: today_(), formUrl: CONFIG.FORM_URL, appUrl: appUrl_() };
}

function appUrl_() { return CONFIG.APP_URL || ScriptApp.getService().getUrl(); }

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
      email: /@/.test(cell_(row, col.email)) ? cell_(row, col.email) : '',
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

    const booking = {
      id: Utilities.getUuid().slice(0, 8), by: by, with: withName, date: input.date, time: input.time,
      format: String(input.format || '').slice(0, 40), note: String(input.note || '').slice(0, 300),
    };
    bookingsSheet_().appendRow([booking.id, new Date(), booking.by, booking.with, booking.date, booking.time,
      booking.format, booking.note, 'proposed']);
    const data = getData();
    data.emailed = notify_('proposed', booking, people, input.lang);
    return data;
  } finally {
    lock.releaseLock();
  }
}

function confirmBooking(id, name, lang) {
  return updateStatus_(id, name, 'confirmed', function (b) { return key_(b.with) === key_(name); },
    'ERR_ONLY_INVITED', lang);
}

function cancelBooking(id, name, lang) {
  return updateStatus_(id, name, 'cancelled', function (b) {
    return key_(b.with) === key_(name) || key_(b.by) === key_(name);
  }, 'ERR_ONLY_PAIR', lang);
}

function updateStatus_(id, name, status, allowed, refusal, lang) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = bookingsSheet_();
    const values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]) !== String(id)) continue;
      if (!allowed({ by: values[i][2], with: values[i][3] })) throw new Error(refusal);
      sheet.getRange(i + 1, 9).setValue(status);
      const r = values[i];
      const booking = { id: String(r[0]), by: String(r[2]), with: String(r[3]), date: asDate_(r[4]),
                        time: asTime_(r[5]), format: String(r[6]), note: String(r[7]), actor: name };
      const data = getData();
      data.emailed = notify_(status, booking, readPeople_(), lang);
      return data;
    }
    throw new Error('ERR_NOT_FOUND');
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Email notifications ---------- */

const MAIL = {
  it: {
    proposed: { to: 'with', subject: function (b) { return b.byFirst + ' ti propone una sessione di pratica SE'; },
      lead: function (b) { return b.by + ' ti ha proposto una sessione di pratica:'; },
      action: 'Apri l\'app per confermarla o annullarla.', button: 'Conferma la sessione' },
    confirmed: { to: 'by', subject: function (b) { return b.withFirst + ' ha confermato la sessione di pratica'; },
      lead: function (b) { return b.with + ' ha confermato la sessione che gli/le hai proposto:'; },
      action: 'Buona pratica!', button: 'Apri l\'app' },
    cancelled: { to: 'other', subject: function (b) { return 'Sessione di pratica annullata · ' + b.whenShort; },
      lead: function (b) { return b.actorName + ' ha annullato questa sessione di pratica:'; },
      action: 'Se volete, potete proporne un\'altra dall\'app.', button: 'Apri l\'app' },
    labels: { when: 'Quando', format: 'Formato', note: 'Nota', at: 'alle' },
    formats: { 'Online': 'Online', 'In person': 'In presenza' },
    footer: 'Ricevi questa email perché hai inserito le tue disponibilità per la pratica SE tra compagni di corso.',
    locale: 'it',
  },
  en: {
    proposed: { to: 'with', subject: function (b) { return b.byFirst + ' proposed an SE practice session'; },
      lead: function (b) { return b.by + ' proposed a practice session with you:'; },
      action: 'Open the app to confirm or cancel it.', button: 'Confirm the session' },
    confirmed: { to: 'by', subject: function (b) { return b.withFirst + ' confirmed your practice session'; },
      lead: function (b) { return b.with + ' confirmed the session you proposed:'; },
      action: 'Enjoy the practice!', button: 'Open the app' },
    cancelled: { to: 'other', subject: function (b) { return 'Practice session cancelled · ' + b.whenShort; },
      lead: function (b) { return b.actorName + ' cancelled this practice session:'; },
      action: 'You can propose another one from the app.', button: 'Open the app' },
    labels: { when: 'When', format: 'Format', note: 'Note', at: 'at' },
    formats: { 'Online': 'Online', 'In person': 'In person' },
    footer: 'You receive this email because you shared your availability for SE peer practice.',
    locale: 'en',
  },
};

/* Sends one email for the event; returns the recipient's name if sent, '' otherwise. Never blocks the booking. */
function notify_(event, b, people, lang) {
  if (!CONFIG.SEND_EMAILS) return '';
  try {
    const t = MAIL[lang] || MAIL.it;
    const tpl = t[event];
    const find = function (n) { return people.find(function (p) { return key_(p.name) === key_(n); }) || {}; };
    const byP = find(b.by), withP = find(b.with);
    var to, from;
    if (tpl.to === 'with') { to = withP; from = byP; }
    else if (tpl.to === 'by') { to = byP; from = withP; }
    else {   // cancellation: tell the person who did NOT cancel
      const actorIsBy = key_(b.actor) === key_(b.by);
      to = actorIsBy ? withP : byP; from = actorIsBy ? byP : withP;
    }
    if (!to.email || MailApp.getRemainingDailyQuota() < 1) return '';

    const first = function (n) { return String(n).split(/\s+/)[0]; };
    const d = dateLabel_(b.date, t.locale);
    const v = Object.assign({}, b, {
      byFirst: first(b.by), withFirst: first(b.with), actorName: from.name || b.actor,
      whenShort: d.short + ' ' + b.time,
    });
    const when = d.long + ' ' + t.labels.at + ' ' + b.time;
    const format = t.formats[b.format] || b.format;
    const url = appUrl_();

    const rows = [[t.labels.when, when], [t.labels.format, format]];
    if (b.note) rows.push([t.labels.note, b.note]);

    const text = [tpl.lead(v), ''].concat(rows.map(function (r) { return r[0] + ': ' + r[1]; }),
      ['', tpl.action, url, '', '—', t.footer]).join('\n');

    const html =
      '<div style="background:#F5F1E7;padding:28px 16px;font-family:Helvetica,Arial,sans-serif;color:#262419">' +
      '<div style="max-width:520px;margin:0 auto;background:#fff;border:1px solid #E2DBC9;border-radius:16px;padding:28px">' +
      '<p style="margin:0 0 6px;font-size:11px;letter-spacing:.12em;text-transform:uppercase;font-weight:bold;color:#4B6B57">SE Practice Partners</p>' +
      '<p style="margin:0 0 18px;font-family:Georgia,serif;font-size:21px;line-height:1.3">' + esc_(tpl.lead(v)) + '</p>' +
      '<table style="border-collapse:collapse;width:100%;font-size:15px;margin-bottom:20px">' +
      rows.map(function (r) {
        return '<tr><td style="padding:8px 0;border-top:1px solid #ECE6D7;color:#6B6553;width:90px;vertical-align:top">' + esc_(r[0]) +
          '</td><td style="padding:8px 0;border-top:1px solid #ECE6D7">' + esc_(r[1]) + '</td></tr>';
      }).join('') + '</table>' +
      '<p style="margin:0 0 18px;font-size:15px">' + esc_(tpl.action) + '</p>' +
      '<a href="' + esc_(url) + '" style="display:inline-block;background:#3F5443;color:#fff;text-decoration:none;font-weight:bold;padding:12px 22px;border-radius:12px">' + esc_(tpl.button) + '</a>' +
      '</div><p style="max-width:520px;margin:16px auto 0;font-size:12px;color:#6B6553;text-align:center">' + esc_(t.footer) + '</p></div>';

    const opts = { name: CONFIG.SENDER_NAME, htmlBody: html };
    if (from.email) opts.replyTo = from.email;   // replies go to the other student, not to the app owner
    MailApp.sendEmail(to.email, tpl.subject(v), text, opts);
    return to.name;
  } catch (e) {
    console.error('Email not sent: ' + e);
    return '';
  }
}

function dateLabel_(iso, locale) {
  const p = iso.split('-').map(Number);
  const d = new Date(p[0], p[1] - 1, p[2], 12);
  const it = locale === 'it';
  const days = it ? ['domenica','lunedì','martedì','mercoledì','giovedì','venerdì','sabato']
                  : ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  const months = it ? ['gennaio','febbraio','marzo','aprile','maggio','giugno','luglio','agosto','settembre','ottobre','novembre','dicembre']
                    : ['January','February','March','April','May','June','July','August','September','October','November','December'];
  return { long: days[d.getDay()] + ' ' + d.getDate() + ' ' + months[d.getMonth()], short: d.getDate() + '/' + (d.getMonth() + 1) };
}

function esc_(s) {
  return String(s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; });
}

/* Run once from the editor to authorize email sending and check the setup. */
function testEmail() {
  MailApp.sendEmail(Session.getActiveUser().getEmail(), 'SE Practice Partners · test', 'Email notifications are working.');
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
    email: find(function (x) { return x.indexOf('email') !== -1 || x.indexOf('e-mail') !== -1; }),
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
