/** SE Practice Partners · Google Apps Script web app (published automatically from GitHub)
 * SE Practice Partners · Google Apps Script web app
 * Bound to the spreadsheet that collects the form responses.
 * Reads availability from the responses sheet, stores bookings in a "Bookings" sheet.
 * Serves a small JSON API; the page itself is hosted on GitHub Pages (index.html in the repository).
 */

const CONFIG = {
  FORM_URL: 'https://forms.gle/Je4gaWUF94cij7Vb8',   // link to the Google Form (Send → link)
  GOOGLE_CLIENT_ID: '845019442597-j5b4h0tj7p3jc5pnnv9scgr0rqj7r8hv.apps.googleusercontent.com',   // «Accedi con Google»
  APP_URL: 'https://se-practice.mnesti.it',                                // short link to the app (e.g. https://pratica.mnesti.it); empty = Apps Script URL
  RESPONSES_SHEET: '',      // leave empty to auto-detect the sheet with the form responses
  BOOKINGS_SHEET: 'Bookings',
  PHOTOS_SHEET: 'Photos',
  NOW_SHEET: 'Now',
  ZONES_SHEET: 'Zones',       // time zone of each person, taken from their browser when they sign in           // "free right now" switches: one row per person with the time it ends     // photos uploaded from the app (small JPEGs, one row per person)
  TIMEZONE: 'Europe/Rome',
  MIN_GAP_MINUTES: 90,        // a person can't have two sessions closer than this
  SEND_EMAILS: true,
  CREATE_MEET: true,          // when everyone has confirmed: calendar event + Google Meet link (needs the Calendar service, see ISTRUZIONI)
  SESSION_MINUTES: 90,
  BACKUP_FOLDER: 'SE Practice · backup',   // Drive folder for the nightly .xlsx copies
  BACKUP_KEEP: 30,            // how many backups to keep        // length of the calendar event          // email notifications for proposals, confirmations and cancellations
  SENDER_NAME: 'SE Practice Partners',
};

const BOOKING_HEADERS = ['ID', 'Created', 'Proposed by', 'With', 'Date', 'Time', 'Format', 'Note', 'Status', 'Third', 'Confirmed by', 'Meet', 'Event'];

/* ---------- JSON API used by the page on GitHub Pages ----------
   GET  …/exec?fn=getData                      → data
   POST …/exec  body: {"fn": "...", "args": [...]}  (sent as text/plain to avoid CORS preflight) */

const API = { getData: getData, createBooking: createBooking, confirmBooking: confirmBooking, cancelBooking: cancelBooking,
              whoAmI: whoAmI, signIn: signIn, requestLink: requestLink, uploadPhoto: uploadPhoto, getNow: getNow, setFreeNow: setFreeNow,
              setTimeZone: setTimeZone };

function doGet(e) {
  const fn = e && e.parameter && e.parameter.fn;
  if (!fn) return json_({ ok: true, info: 'SE Practice Partners API' });
  return run_(fn, []);
}

function doPost(e) {
  var body = {};
  try { body = JSON.parse(e.postData.contents || '{}'); } catch (err) { return json_({ error: 'ERR_BAD_REQUEST' }); }
  return run_(body.fn, Array.isArray(body.args) ? body.args : []);
}

function run_(fn, args) {
  if (!Object.prototype.hasOwnProperty.call(API, fn)) return json_({ error: 'ERR_BAD_REQUEST' });
  try { return json_({ data: API[fn].apply(null, args) }); }
  catch (err) { return json_({ error: String(err && err.message || err) }); }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Data for the page ---------- */

/* k = personal key of whoever is looking. "Free right now" is shared only with signed-in students. */
function getData(k) {
  // Email addresses never leave the server: the page only learns whether someone can be notified.
  const people = readPeople_().map(function (p) {
    const out = Object.assign({}, p);
    out.hasEmail = !!p.email;
    delete out.email;
    return out;
  });
  const now = k && personByKey_(k) ? readNow_() : {};
  const zones = readZones_();
  people.forEach(function (p) {
    if (now[p.name]) p.freeUntil = now[p.name];
    if (zones[p.name] && zones[p.name] !== CONFIG.TIMEZONE) p.tz = zones[p.name];
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

  const uploaded = readPhotos_();
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
      photo: uploaded[personKey_(name)] || photo_(cell_(row, col.photo)),
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
        third: String(r[9] || ''),
        confirmedBy: String(r[10] || '').split(',').map(function (x) { return x.trim(); }).filter(String),
      };
    })
    .filter(function (b) { return b.id && b.date >= now && b.status !== 'cancelled'; })
    .sort(function (a, b) { return (a.date + a.time).localeCompare(b.date + b.time); });
}

function all_(b) { return participants_(b); }
function participants_(b) { return [b.by, b.with].concat(b.third ? [b.third] : []); }
function invitees_(b) { return [b.with].concat(b.third ? [b.third] : []); }

/* ---------- Personal links ----------
   Each person has a secret key derived from their name and a random salt kept in the script properties.
   Keys never reach the page in bulk: they only travel in emails sent to that person (…/?me=<key>). */

function salt_() {
  const props = PropertiesService.getScriptProperties();
  var salt = props.getProperty('SALT');
  if (!salt) { salt = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('SALT', salt); }
  return salt;
}

function personKey_(name) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt_() + '|' + key_(name));
  return Utilities.base64EncodeWebSafe(bytes).replace(/[^A-Za-z0-9]/g, '').slice(0, 16);
}

function personalUrl_(name) { return appUrl_().replace(/\/?$/, '/') + '?me=' + personKey_(name); }

function personByKey_(k) {
  if (!k) return null;
  return readPeople_().find(function (p) { return personKey_(p.name) === k; }) || null;
}

/* The page calls this when opened from a personal link */
function whoAmI(k) {
  const p = personByKey_(k);
  return p ? { name: p.name, key: k } : null;
}

/* Every action that changes something goes through here: the name comes from the key, never from the page */
function actor_(k) {
  const p = personByKey_(k);
  if (!p) throw new Error('ERR_NOT_YOU');
  return p;
}

/* «Accedi con Google»: the page sends Google's signed ID token; we check it with Google and match the email
   with the one in the form responses. */
function signIn(idToken) {
  if (!idToken) throw new Error('ERR_BAD_REQUEST');
  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('ERR_SIGNIN');
  const info = JSON.parse(res.getContentText());
  const verified = info.email_verified === true || info.email_verified === 'true';
  if (info.aud !== CONFIG.GOOGLE_CLIENT_ID || !verified || Number(info.exp) * 1000 < Date.now()) throw new Error('ERR_SIGNIN');
  const p = readPeople_().find(function (x) { return x.email && key_(x.email) === key_(info.email); });
  if (!p) throw new Error('ERR_NOT_IN_LIST|' + info.email);
  return { name: p.name, key: personKey_(p.name) };
}

/* Fallback for people who don't use Google: the personal link by email. Same answer whether or not the email is known. */
function requestLink(email, lang) {
  const e = String(email || '').trim();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e)) throw new Error('ERR_EMAIL');
  const p = readPeople_().find(function (x) { return x.email && key_(x.email) === key_(e); });
  if (p) sendLink_(p, lang);
  return { sent: true };
}

/* Emails someone their personal link (no more than once every 10 minutes per person) */
function sendLink_(p, lang) {
  if (!p.email) return { sent: false };
  const cache = CacheService.getScriptCache();
  if (cache.get('link_' + key_(p.name))) return { sent: true };
  const it = (lang || 'it') === 'it';
  const url = personalUrl_(p.name);
  const lead = it ? 'Ecco il tuo link personale per l\'app dei compagni di pratica SE.'
                  : 'Here is your personal link to the SE practice partners app.';
  const body = it ? 'Aprendolo, l\'app ti riconosce su questo dispositivo: puoi proporre sessioni, confermarle e caricare la tua foto. Non inoltrarlo: chi lo apre agisce a nome tuo.'
                  : 'Opening it, the app recognises you on that device: you can propose and confirm sessions and upload your photo. Don\'t forward it: whoever opens it acts as you.';
  const button = it ? 'Apri l\'app' : 'Open the app';
  const html = mailFrame_(lead, '<p style="margin:0 0 18px;font-size:15px">' + esc_(body) + '</p>', url, button, MAIL[it ? 'it' : 'en'].footer);
  MailApp.sendEmail(p.email, it ? 'Il tuo link personale · SE Practice Partners' : 'Your personal link · SE Practice Partners',
    lead + '\n\n' + body + '\n' + url, { name: CONFIG.SENDER_NAME, htmlBody: html });
  cache.put('link_' + key_(p.name), '1', 600);
  return { sent: true };
}

/* Photo uploaded from the app: a small square JPEG prepared in the browser, saved in the Photos sheet */
function uploadPhoto(k, dataUrl) {
  const p = personByKey_(k);
  if (!p) throw new Error('ERR_NOT_YOU');
  if (!/^data:image\/jpeg;base64,[A-Za-z0-9+\/=]+$/.test(String(dataUrl || '')) || dataUrl.length > 48000) {
    throw new Error('ERR_PHOTO');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = photosSheet_();
    const pk = personKey_(p.name);
    const values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]) === pk) {
        sheet.getRange(i + 1, 2, 1, 3).setValues([[p.name, dataUrl, new Date()]]);
        return getData(k);
      }
    }
    sheet.appendRow([pk, p.name, dataUrl, new Date()]);
    return getData(k);
  } finally {
    lock.releaseLock();
  }
}

function readPhotos_() {
  const out = {};
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.PHOTOS_SHEET);
  if (!sheet) return out;
  sheet.getDataRange().getValues().slice(1).forEach(function (r) { if (r[0] && r[2]) out[String(r[0])] = String(r[2]); });
  return out;
}

function photosSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.PHOTOS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.PHOTOS_SHEET);
    sheet.getRange(1, 1, 1, 4).setValues([['Key', 'Name', 'Photo', 'Updated']]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/* ---------- "Free right now" ----------
   A light switch for sessions organised on the spot. It switches itself off when its time is up. */

const FREE_MINUTES = [30, 60, 120, 180];

/* Polled by the page every minute: only names and end times, nothing else */
function getNow(k) {
  const ok = !!(k && personByKey_(k));
  return { now: ok ? readNow_() : {}, signedIn: ok, serverTime: new Date().toISOString() };
}

function setFreeNow(k, minutes) {
  const p = actor_(k);
  const mins = Number(minutes) || 0;
  if (mins && FREE_MINUTES.indexOf(mins) === -1) throw new Error('ERR_BAD_REQUEST');
  const until = mins ? new Date(Date.now() + mins * 60000) : '';
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = nowSheet_();
    const values = sheet.getDataRange().getValues();
    var row = -1;
    for (var i = 1; i < values.length; i++) { if (key_(values[i][0]) === key_(p.name)) { row = i + 1; break; } }
    if (row > 0) sheet.getRange(row, 2).setValue(until);
    else if (until) sheet.appendRow([p.name, until]);
    return getNow(k);
  } finally {
    lock.releaseLock();
  }
}

function readNow_() {
  const out = {};
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.NOW_SHEET);
  if (!sheet) return out;
  const t = Date.now();
  sheet.getDataRange().getValues().slice(1).forEach(function (r) {
    const until = r[1] instanceof Date ? r[1] : (r[1] ? new Date(r[1]) : null);
    if (r[0] && until && until.getTime() > t) out[String(r[0])] = until.toISOString();
  });
  return out;
}

function nowSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.NOW_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.NOW_SHEET);
    sheet.getRange(1, 1, 1, 2).setValues([['Name', 'Free until']]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/* ---------- Time zones ----------
   Sessions are always stored in Italian time (CONFIG.TIMEZONE); each person's zone is used only to show local times. */

function setTimeZone(k, tz) {
  const p = actor_(k);
  tz = String(tz || '');
  try { Utilities.formatDate(new Date(), tz, 'Z'); } catch (e) { throw new Error('ERR_BAD_REQUEST'); }
  if (!/^[A-Za-z_]+(\/[A-Za-z0-9_+\-]+)+$/.test(tz) && tz !== 'UTC') throw new Error('ERR_BAD_REQUEST');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = zonesSheet_();
    const values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (key_(values[i][0]) === key_(p.name)) { sheet.getRange(i + 1, 2).setValue(tz); return { tz: tz }; }
    }
    sheet.appendRow([p.name, tz]);
    return { tz: tz };
  } finally {
    lock.releaseLock();
  }
}

function readZones_() {
  const out = {};
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ZONES_SHEET);
  if (!sheet) return out;
  sheet.getDataRange().getValues().slice(1).forEach(function (r) { if (r[0] && r[1]) out[String(r[0])] = String(r[1]); });
  return out;
}

function zonesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.ZONES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.ZONES_SHEET);
    sheet.getRange(1, 1, 1, 2).setValues([['Name', 'Time zone']]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/* Italian date + time of a session → the same moment written in another zone */
function inZone_(date, time, tz) {
  const instant = Utilities.parseDate(date + ' ' + time, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm');
  return { date: Utilities.formatDate(instant, tz, 'yyyy-MM-dd'), time: Utilities.formatDate(instant, tz, 'HH:mm') };
}

/* ---------- Actions ---------- */

function createBooking(input) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const people = readPeople_();
    const names = {};
    people.forEach(function (p) { names[key_(p.name)] = p.name; });

    const by = actor_(input.key).name;   // who proposes = whoever holds the key
    const withName = names[key_(input.with)];
    const third = input.third ? names[key_(input.third)] : '';
    if (!by || !withName || (input.third && !third)) throw new Error('ERR_PICK_TWO');
    if (by === withName) throw new Error('ERR_SELF');
    if (third && (third === by || third === withName)) throw new Error('ERR_THIRD');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(input.date || '')) throw new Error('ERR_DATE');
    if (!/^\d{2}:\d{2}$/.test(input.time || '')) throw new Error('ERR_TIME');
    if (input.date < today_()) throw new Error('ERR_PAST');

    const booking = {
      id: Utilities.getUuid().slice(0, 8), by: by, with: withName, third: third, date: input.date, time: input.time,
      format: String(input.format || '').slice(0, 40), note: String(input.note || '').slice(0, 300),
    };
    const mine = participants_(booking);
    const start = minutes_(input.time);
    const clash = readBookings_().find(function (b) {
      const involved = participants_(b).some(function (n) { return mine.indexOf(n) !== -1; });
      return involved && b.date === input.date && Math.abs(minutes_(b.time) - start) < CONFIG.MIN_GAP_MINUTES;
    });
    if (clash) throw new Error(['ERR_CLASH', clash.by, clash.with, clash.time].join('|'));

    bookingsSheet_().appendRow([booking.id, new Date(), booking.by, booking.with, booking.date, booking.time,
      booking.format, booking.note, 'proposed', booking.third, '']);
    const data = getData(input.key);
    data.emailed = notify_('proposed', booking, people, input.lang, by);
    return data;
  } finally {
    lock.releaseLock();
  }
}

/* Each invited person confirms for themselves; the session is "confirmed" once everyone invited has confirmed. */
function confirmBooking(id, k, lang) {
  const name = actor_(k).name;
  return withBooking_(id, function (b, row, sheet) {
    const who = invitees_(b).find(function (n) { return key_(n) === key_(name); });
    if (!who) throw new Error('ERR_ONLY_INVITED');
    if (b.confirmedBy.indexOf(who) === -1) b.confirmedBy.push(who);
    const all = invitees_(b).every(function (n) { return b.confirmedBy.indexOf(n) !== -1; });
    sheet.getRange(row, 11).setValue(b.confirmedBy.join(', '));
    sheet.getRange(row, 9).setValue(all ? 'confirmed' : 'proposed');
    b.allConfirmed = all;
    if (!all) return { event: 'confirmed', actor: who };
    // Everyone is in: calendar event with a Meet link, then one email to all participants
    if (!b.eventId) {
      const ev = createEvent_(b, readPeople_());
      if (ev.eventId) { sheet.getRange(row, 12, 1, 2).setValues([[ev.meet || '', ev.eventId]]); b.meet = ev.meet || ''; b.eventId = ev.eventId; }
    }
    return { event: 'scheduled', actor: who, everyone: true };
  }, lang, k);
}

function cancelBooking(id, k, lang) {
  const name = actor_(k).name;
  return withBooking_(id, function (b, row, sheet) {
    const who = participants_(b).find(function (n) { return key_(n) === key_(name); });
    if (!who) throw new Error('ERR_ONLY_PAIR');
    sheet.getRange(row, 9).setValue('cancelled');
    if (b.eventId) removeEvent_(b.eventId);
    return { event: 'cancelled', actor: who };
  }, lang, k);
}

function withBooking_(id, change, lang, k) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = bookingsSheet_();
    const values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      const r = values[i];
      if (String(r[0]) !== String(id)) continue;
      const b = { id: String(r[0]), by: String(r[2]), with: String(r[3]), date: asDate_(r[4]), time: asTime_(r[5]),
        format: String(r[6]), note: String(r[7]), third: String(r[9] || ''),
        confirmedBy: String(r[10] || '').split(',').map(function (x) { return x.trim(); }).filter(String),
        meet: String(r[11] || ''), eventId: String(r[12] || '') };
      const out = change(b, i + 1, sheet);
      const data = getData(k);
      data.emailed = notify_(out.event, b, readPeople_(), lang, out.actor, out.everyone);
      data.scheduled = out.event === 'scheduled';
      data.meet = !!b.meet;
      return data;
    }
    throw new Error('ERR_NOT_FOUND');
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Calendar event with Google Meet ----------
   Uses the advanced "Google Calendar API" service (identifier: Calendar). Without it, the session is still
   confirmed and emailed, just without the Meet link. */

function createEvent_(b, people) {
  if (!CONFIG.CREATE_MEET || typeof Calendar === 'undefined') return {};
  try {
    const find = function (n) { return people.find(function (p) { return key_(p.name) === key_(n); }) || {}; };
    const emails = participants_(b).map(function (n) { return find(n).email; }).filter(Boolean);
    const online = b.format !== 'In person';
    const resource = {
      summary: 'Pratica SE · ' + participants_(b).join(', '),
      description: (b.note ? b.note + '\n\n' : '') + 'Organizzata con SE Practice Partners · ' + appUrl_(),
      start: { dateTime: localDateTime_(b.date, b.time, 0), timeZone: CONFIG.TIMEZONE },
      end: { dateTime: localDateTime_(b.date, b.time, CONFIG.SESSION_MINUTES), timeZone: CONFIG.TIMEZONE },
      attendees: emails.map(function (e) { return { email: e }; }),
      guestsCanSeeOtherGuests: true,
      reminders: { useDefault: false, overrides: [{ method: 'popup', minutes: 15 }] },
    };
    if (online) {
      resource.conferenceData = { createRequest: { requestId: b.id + '-' + Date.now(), conferenceSolutionKey: { type: 'hangoutsMeet' } } };
    }
    // sendUpdates "none": our own email carries the details, so nobody gets two messages
    const ev = Calendar.Events.insert(resource, 'primary', { conferenceDataVersion: 1, sendUpdates: 'none' });
    const video = ((ev.conferenceData && ev.conferenceData.entryPoints) || []).find(function (e) { return e.entryPointType === 'video'; });
    return { eventId: ev.id, meet: ev.hangoutLink || (video && video.uri) || '' };
  } catch (e) {
    console.error('Calendar event not created: ' + e);
    return {};
  }
}

function removeEvent_(eventId) {
  if (typeof Calendar === 'undefined') return;
  try { Calendar.Events.remove('primary', eventId, { sendUpdates: 'none' }); }
  catch (e) { console.error('Calendar event not removed: ' + e); }
}

/* "2026-11-04", "19:00", +90 → "2026-11-04T20:30:00", computed without time-zone surprises */
function localDateTime_(date, time, plusMinutes) {
  const d = date.split('-').map(Number), t = time.split(':').map(Number);
  const x = new Date(Date.UTC(d[0], d[1] - 1, d[2], t[0], t[1] + plusMinutes));
  const pad = function (n) { return ('0' + n).slice(-2); };
  return x.getUTCFullYear() + '-' + pad(x.getUTCMonth() + 1) + '-' + pad(x.getUTCDate()) + 'T' + pad(x.getUTCHours()) + ':' + pad(x.getUTCMinutes()) + ':00';
}

/* ---------- Email notifications ---------- */

const MAIL = {
  it: {
    proposed: { subject: function (v) { return v.actorFirst + ' ti propone una sessione di pratica SE'; },
      lead: function (v) { return v.actor + ' ti ha proposto una sessione di pratica' + (v.others ? ' insieme a ' + v.others : '') + ':'; },
      action: 'Apri l\'app per confermarla o annullarla.', button: 'Conferma la sessione' },
    confirmed: { subject: function (v) { return v.actorFirst + ' ha confermato la sessione di pratica'; },
      lead: function (v) { return v.actor + ' ha confermato la sessione di pratica' + (v.allConfirmed ? ': ora è confermata da tutti.' : '. Si attende ancora la conferma di ' + v.pending + '.'); },
      action: 'Buona pratica!', button: 'Apri l\'app' },
    scheduled: { subject: function (v) { return 'Sessione confermata · ' + v.whenShort + (v.meet ? ' · link Meet' : ''); },
      lead: function (v) { return 'La sessione di pratica è confermata da tutti.' + (v.meet ? ' Ecco il link per collegarvi.' : ''); },
      action: 'Buona pratica!', button: 'Entra su Google Meet', buttonNoMeet: 'Apri l\'app' },
    cancelled: { subject: function (v) { return 'Sessione di pratica annullata · ' + v.whenShort; },
      lead: function (v) { return v.actor + ' ha annullato questa sessione di pratica:'; },
      action: 'Se volete, potete proporne un\'altra dall\'app.', button: 'Apri l\'app' },
    labels: { when: 'Quando', format: 'Formato', who: 'Partecipanti', note: 'Nota', at: 'alle', and: 'e', meet: 'Google Meet', app: 'App',
              italy: 'ora italiana', local: 'Da te' },
    formats: { 'Online': 'Online', 'In person': 'In presenza' },
    footer: 'Ricevi questa email perché hai inserito le tue disponibilità per la pratica SE tra compagni di corso.',
    locale: 'it',
  },
  en: {
    proposed: { subject: function (v) { return v.actorFirst + ' proposed an SE practice session'; },
      lead: function (v) { return v.actor + ' proposed a practice session with you' + (v.others ? ' and ' + v.others : '') + ':'; },
      action: 'Open the app to confirm or cancel it.', button: 'Confirm the session' },
    confirmed: { subject: function (v) { return v.actorFirst + ' confirmed the practice session'; },
      lead: function (v) { return v.actor + ' confirmed the practice session' + (v.allConfirmed ? ': everyone has now confirmed.' : '. Still waiting for ' + v.pending + '.'); },
      action: 'Enjoy the practice!', button: 'Open the app' },
    scheduled: { subject: function (v) { return 'Session confirmed · ' + v.whenShort + (v.meet ? ' · Meet link' : ''); },
      lead: function (v) { return 'The practice session is confirmed by everyone.' + (v.meet ? ' Here is the link to join.' : ''); },
      action: 'Enjoy the practice!', button: 'Join on Google Meet', buttonNoMeet: 'Open the app' },
    cancelled: { subject: function (v) { return 'Practice session cancelled · ' + v.whenShort; },
      lead: function (v) { return v.actor + ' cancelled this practice session:'; },
      action: 'You can propose another one from the app.', button: 'Open the app' },
    labels: { when: 'When', format: 'Format', who: 'People', note: 'Note', at: 'at', and: 'and', meet: 'Google Meet', app: 'App',
              italy: 'Italian time', local: 'Your time' },
    formats: { 'Online': 'Online', 'In person': 'In person' },
    footer: 'You receive this email because you shared your availability for SE peer practice.',
    locale: 'en',
  },
};

/* Sends the emails for an event and returns the names of the people emailed. Never blocks the booking.
   proposed → everyone invited · confirmed → everyone except the person confirming · cancelled → everyone except who cancelled */
function notify_(event, b, people, lang, actor, everyone) {
  if (!CONFIG.SEND_EMAILS) return [];
  const t = MAIL[lang] || MAIL.it;
  const tpl = t[event];
  const find = function (n) { return people.find(function (p) { return key_(p.name) === key_(n); }) || { name: n }; };
  const first = function (n) { return String(n).split(/\s+/)[0]; };
  const zones = readZones_();
  const abroad = all_(b).some(function (n) { return zones[n] && zones[n] !== CONFIG.TIMEZONE; });
  const joinNames = function (list) {
    return list.length < 2 ? list.join('') : list.slice(0, -1).join(', ') + ' ' + t.labels.and + ' ' + list[list.length - 1];
  };
  const all = participants_(b);
  const recipients = everyone ? all : all.filter(function (n) { return key_(n) !== key_(actor); });
  const actorP = find(actor);
  const d = dateLabel_(b.date, t.locale);
  const sent = [];

  recipients.forEach(function (name) {
    try {
      const to = find(name);
      if (!to.email || MailApp.getRemainingDailyQuota() < 1) return;
      const others = event === 'proposed' ? all.filter(function (n) { return n !== actor && n !== name; }) : [];
      const pending = invitees_(b).filter(function (n) { return (b.confirmedBy || []).indexOf(n) === -1; });
      const v = { actor: actor, actorFirst: first(actor), others: joinNames(others), allConfirmed: !!b.allConfirmed,
                  pending: joinNames(pending), whenShort: d.short + ' ' + b.time, meet: b.meet || '' };
      const rows = [[t.labels.when, d.long + ' ' + t.labels.at + ' ' + b.time + (abroad ? ' (' + t.labels.italy + ')' : '')]];
      const tz = zones[name];
      if (tz && tz !== CONFIG.TIMEZONE) {
        const loc = inZone_(b.date, b.time, tz), ld = dateLabel_(loc.date, t.locale);
        rows.push([t.labels.local, ld.long + ' ' + t.labels.at + ' ' + loc.time + ' · ' + tz.split('/').pop().replace(/_/g, ' ')]);
      }
      rows.push([t.labels.format, t.formats[b.format] || b.format]);
      if (all.length > 2 || event === 'scheduled') rows.push([t.labels.who, joinNames(all)]);
      if (event === 'scheduled' && b.meet) rows.push([t.labels.meet, b.meet]);
      if (b.note) rows.push([t.labels.note, b.note]);

      const appLink = personalUrl_(to.name);   // opens the app with their name already chosen
      const url = event === 'scheduled' && b.meet ? b.meet : appLink;
      const button = event === 'scheduled' && !b.meet ? tpl.buttonNoMeet : tpl.button;
      const text = [tpl.lead(v), ''].concat(rows.map(function (r) { return r[0] + ': ' + r[1]; }),
        ['', tpl.action, url].concat(url !== appLink ? [t.labels.app + ': ' + appLink] : [], ['', '—', t.footer])).join('\n');
      const table = '<table style="border-collapse:collapse;width:100%;font-size:15px;margin-bottom:20px">' +
        rows.map(function (r) {
          return '<tr><td style="padding:8px 0;border-top:1px solid #EFE8EC;color:#6E686B;width:110px;vertical-align:top">' + esc_(r[0]) +
            '</td><td style="padding:8px 0;border-top:1px solid #EFE8EC">' + esc_(r[1]) + '</td></tr>';
        }).join('') + '</table>' +
        '<p style="margin:0 0 18px;font-size:15px">' + esc_(tpl.action) + '</p>';
      const html = mailFrame_(tpl.lead(v), table, url, button, t.footer);

      const opts = { name: CONFIG.SENDER_NAME, htmlBody: html };
      if (actorP.email) opts.replyTo = actorP.email;   // replies go to the student who acted, not to the app owner
      MailApp.sendEmail(to.email, tpl.subject(v), text, opts);
      sent.push(to.name);
    } catch (e) {
      console.error('Email to ' + name + ' not sent: ' + e);
    }
  });
  return sent;
}

function mailFrame_(lead, inner, url, button, footer) {
  return '<div style="background:#F7F3F5;padding:28px 16px;font-family:Roboto,Helvetica,Arial,sans-serif;color:#3D3D3D">' +
    '<div style="max-width:520px;margin:0 auto;background:#fff;border:1px solid #E5DCE1;border-radius:16px;padding:28px">' +
    '<p style="margin:0 0 6px;font-size:11px;letter-spacing:.12em;text-transform:uppercase;font-weight:bold;color:#D86018">SE Practice Partners</p>' +
    '<p style="margin:0 0 18px;font-family:Comfortaa,Trebuchet MS,Helvetica,sans-serif;font-size:20px;color:#6C1D45;line-height:1.3">' + esc_(lead) + '</p>' +
    inner +
    '<a href="' + esc_(url) + '" style="display:inline-block;background:#6C1D45;color:#fff;text-decoration:none;font-weight:bold;padding:12px 22px;border-radius:12px">' + esc_(button) + '</a>' +
    '</div><p style="max-width:520px;margin:16px auto 0;font-size:12px;color:#6E686B;text-align:center">' + esc_(footer) + '</p></div>';
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

/* ---------- Nightly backup ----------
   Run installBackup() once from the editor: it schedules backupNow() every night at about 3:00 and makes a first copy.
   The copy is an .xlsx export (copying a form-linked spreadsheet would also duplicate the form). */

function installBackup() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'backupNow') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('backupNow').timeBased().everyDays(1).atHour(3).inTimezone(CONFIG.TIMEZONE).create();
  return backupNow();
}

function backupNow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
  const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('Backup failed: HTTP ' + res.getResponseCode());
  const name = ss.getName() + ' · backup ' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH.mm') + '.xlsx';
  const folder = backupFolder_();
  const file = folder.createFile(res.getBlob().setName(name));

  // Keep only the most recent BACKUP_KEEP copies (older ones go to the Drive bin, recoverable for 30 days)
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) files.push(it.next());
  files.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  files.slice(CONFIG.BACKUP_KEEP).forEach(function (f) { f.setTrashed(true); });
  console.log('Backup saved: ' + file.getName() + ' · ' + file.getUrl());
  return file.getUrl();
}

function backupFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('BACKUP_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* deleted: create it again */ } }
  const folder = DriveApp.createFolder(CONFIG.BACKUP_FOLDER);
  props.setProperty('BACKUP_FOLDER_ID', folder.getId());
  return folder;
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
  } else if (sheet.getLastColumn() < BOOKING_HEADERS.length) {   // sheets created before triads existed
    sheet.getRange(1, 1, 1, BOOKING_HEADERS.length).setValues([BOOKING_HEADERS]).setFontWeight('bold');
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
