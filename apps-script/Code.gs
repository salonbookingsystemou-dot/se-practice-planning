/** SE Practice Partners · Google Apps Script web app (published automatically from GitHub)
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
  PHOTOS_SHEET: 'Photos',     // photos uploaded from the app (small JPEGs, one row per person)
  NOW_SHEET: 'Now',           // "free right now" switches: one row per person with the time it ends
  ZONES_SHEET: 'Zones',       // per person: time zone, dashboard language, reminders on/off
  FEEDBACK_SHEET: 'Feedback',  // post-session comments: one row per comment
  FEEDBACK_DAYS: 90,          // sessions shown on the feedback board (days back)
  FEEDBACK_PUBLIC: false,     // false = only the session's participants see and write its board
  MAIL_LOG_SHEET: 'Mail log', // one row per email: when, type, to, subject, result
  TIMEZONE: 'Europe/Rome',
  MIN_GAP_MINUTES: 90,        // a person can't have two sessions closer than this
  SEND_EMAILS: true,          // email notifications for proposals, confirmations and cancellations
  CREATE_MEET: true,          // when everyone has confirmed: calendar event + Google Meet link (needs the Calendar service, see ISTRUZIONI)
  SESSION_MINUTES: 90,        // length of a session (calendar event, "still visible" window, digest timing)
  SOON_HOURS: 2,
  START_MINUTES: 15,          // "starting soon" email with the Meet link, about this many minutes before (0 = off)              // reminder to everyone, this many hours before a CONFIRMED session (0 = off)
  REMIND_HOURS: 24,           // reminder to confirm, this many hours before an unconfirmed session (0 = off)
  DIGEST: true,               // after confirmed sessions end: "X and Y just practised, what about you?" to everyone else
  DIGEST_EVERY_HOURS: 20,     // at most one such email per person per day
  DIGEST_RESERVE: 25,         // daily Gmail quota kept free for proposals/confirmations
  BACKUP_FOLDER: 'SE Practice · backup',   // Drive folder for the nightly .xlsx copies
  BACKUP_KEEP: 30,            // how many backups to keep
  SENDER_NAME: 'SE Practice Partners',
};

const BOOKING_HEADERS = ['ID', 'Created', 'Proposed by', 'With', 'Date', 'Time', 'Format', 'Note', 'Status', 'Third', 'Confirmed by', 'Meet', 'Event', 'Shared', 'Reminded', 'Soon', 'Starting'];

/* ---------- JSON API used by the page on GitHub Pages ----------
   GET  …/exec?fn=getData                      → data
   POST …/exec  body: {"fn": "...", "args": [...]}  (sent as text/plain to avoid CORS preflight) */

const API = { getData: getData, createBooking: createBooking, confirmBooking: confirmBooking, cancelBooking: cancelBooking,
              whoAmI: whoAmI, signIn: signIn, requestLink: requestLink, uploadPhoto: uploadPhoto, getNow: getNow, setFreeNow: setFreeNow,
              setTimeZone: setTimeZone, setLanguage: setLanguage, leave: leave, setDigest: setDigest, status: status,
              getFeedback: getFeedback, addFeedback: addFeedback, deleteFeedback: deleteFeedback };

/* Health check without personal data: is the scheduler alive, are Gmail/Calendar usable, what happened to recent sessions */
function status() {
  const props = PropertiesService.getScriptProperties();
  var triggers = [];
  try { triggers = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); }); } catch (e) { triggers = ['error: ' + e.message]; }
  var quota = null; try { quota = MailApp.getRemainingDailyQuota(); } catch (e) { quota = 'error: ' + e.message; }
  const t = Date.now();
  const sessions = bookingsSheet_().getDataRange().getValues().slice(1).map(function (r) {
    const b = { date: asDate_(r[4]), time: asTime_(r[5]) };
    const st = startOf_(b);
    return { id: String(r[0]), date: b.date, time: b.time, status: String(r[8] || 'proposed'), invitees: 1 + (r[9] ? 1 : 0),
             confirmations: String(r[10] || '').split(',').filter(function (x) { return x.trim(); }).length,
             meet: !!r[11], event: !!r[12], reminded: !!r[14], soon: !!r[15], starting: !!r[16], hoursFromNow: st ? Math.round((st.getTime() - t) / 360000) / 10 : null };
  }).filter(function (x) { return x.hoursFromNow !== null && x.hoursFromNow > -48 && x.hoursFromNow < 72; });
  return { now: new Date().toISOString(), triggers: triggers, triggerVersion: props.getProperty('TRIGGER_VERSION'),
           lastRun: props.getProperty('LAST_RUN'), lastRunResult: props.getProperty('LAST_RUN_RESULT'),
           calendarService: typeof Calendar !== 'undefined', mailQuota: quota, sessions: sessions, mailLog: mailLogSummary_() };
}

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
  MEMO_ = {};
  ensureScheduler_();
  if (!Object.prototype.hasOwnProperty.call(API, fn)) return json_({ error: 'ERR_BAD_REQUEST' });
  try { return json_({ data: API[fn].apply(null, args) }); }
  catch (err) { return json_({ error: String(err && err.message || err) }); }
}

/* Last emails: type, result, time, session — no names or addresses */
function mailLogSummary_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MAIL_LOG_SHEET);
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1).slice(-40).map(function (r) {
    return { at: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), type: String(r[1]), result: String(r[5]), session: String(r[6] || '') };
  });
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
  const viewer = k ? personByKey_(k) : null;
  const now = viewer ? readNow_() : {};
  const zones = readZones_();
  const done = sessionsDone_();
  people.forEach(function (p) {
    p.done = done[key_(p.name)] || 0;
    if (now[p.name]) p.freeUntil = now[p.name];
    if (zones[p.name] && zones[p.name] !== CONFIG.TIMEZONE) p.tz = zones[p.name];
  });
  const out = { people: people, bookings: readBookings_(), today: today_(), formUrl: CONFIG.FORM_URL, appUrl: appUrl_() };
  if (viewer) { touch_(viewer); out.online = readOnline_(people); }   // presence: only for signed-in people
  return out;
}

/* What the page needs after proposing / confirming / cancelling: sessions and presence, not everyone's photos again */
function sessionsData_(k) {
  const viewer = k ? personByKey_(k) : null;
  const out = { bookings: readBookings_(), today: today_() };
  if (viewer) { touch_(viewer); out.online = readOnline_(readPeople_()); }
  return out;
}

function appUrl_() { return CONFIG.APP_URL || ScriptApp.getService().getUrl(); }

/* Each request is a fresh execution: these memos avoid reading the same sheet or hashing the same name twice */
var MEMO_ = {};

function readPeople_() {
  if (MEMO_.people) return MEMO_.people;
  MEMO_.people = readPeopleFromSheet_();
  return MEMO_.people;
}

function readPeopleFromSheet_() {
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

/* Upcoming sessions for the page.
   - proposed (not everyone confirmed) and start time passed → marked "expired" in the sheet and hidden
   - confirmed → visible until the end of the session (start + SESSION_MINUTES), then hidden
   Nothing is deleted: the sheet keeps the history. */
function readBookings_() {
  const sheet = bookingsSheet_();
  const values = sheet.getDataRange().getValues();
  const t = Date.now();
  const out = [];
  values.slice(1).forEach(function (r, i) {
    const b = {
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
    if (!b.id || b.status === 'cancelled' || b.status === 'expired') return;
    const start = startOf_(b);
    if (!start) return;
    if (b.status !== 'confirmed' && start.getTime() <= t) {
      try { sheet.getRange(i + 2, 9).setValue('expired'); } catch (e) { /* hidden anyway */ }
      return;
    }
    if (b.status === 'confirmed' && start.getTime() + CONFIG.SESSION_MINUTES * 60000 <= t) return;
    out.push(b);
  });
  return out.sort(function (a, b) { return (a.date + a.time).localeCompare(b.date + b.time); });
}

/* How many confirmed sessions each person has completed (ended), by lower-case name */
function sessionsDone_() {
  if (MEMO_.done) return MEMO_.done;
  const out = {};
  const t = Date.now();
  bookingsSheet_().getDataRange().getValues().slice(1).forEach(function (r) {
    if (String(r[8]) !== 'confirmed') return;
    const st = startOf_({ date: asDate_(r[4]), time: asTime_(r[5]) });
    if (!st || st.getTime() + CONFIG.SESSION_MINUTES * 60000 > t) return;
    [r[2], r[3], r[9]].forEach(function (n) { if (n) out[key_(n)] = (out[key_(n)] || 0) + 1; });
  });
  MEMO_.done = out;
  return out;
}

/* The moment a session starts (stored as Italian date and time) */
function startOf_(b) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(b.date) || !/^\d{2}:\d{2}$/.test(b.time)) return null;
  try { return Utilities.parseDate(b.date + ' ' + b.time, CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm'); } catch (e) { return null; }
}

function all_(b) { return participants_(b); }
function participants_(b) { return [b.by, b.with].concat(b.third ? [b.third] : []); }
function invitees_(b) { return [b.with].concat(b.third ? [b.third] : []); }

/* ---------- Personal links ----------
   Each person has a secret key derived from their name and a random salt kept in the script properties.
   Keys never reach the page in bulk: they only travel in emails sent to that person (…/?me=<key>). */

function salt_() {
  if (MEMO_.salt) return MEMO_.salt;
  const props = PropertiesService.getScriptProperties();
  var salt = props.getProperty('SALT');
  if (!salt) { salt = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('SALT', salt); }
  MEMO_.salt = salt;
  return salt;
}

function personKey_(name) {
  const keys = MEMO_.keys || (MEMO_.keys = {});
  const n = key_(name);
  if (!keys[n]) {
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt_() + '|' + n);
    keys[n] = Utilities.base64EncodeWebSafe(bytes).replace(/[^A-Za-z0-9]/g, '').slice(0, 16);
  }
  return keys[n];
}

function personalUrl_(name) { return appUrl_().replace(/\/?$/, '/') + '?me=' + personKey_(name); }

function personByKey_(k) {
  if (!k) return null;
  if (!MEMO_.byKey) {
    MEMO_.byKey = {};
    readPeople_().forEach(function (p) { MEMO_.byKey[personKey_(p.name)] = p; });
  }
  return MEMO_.byKey[k] || null;
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
  const issuer = info.iss === 'accounts.google.com' || info.iss === 'https://accounts.google.com';
  if (info.aud !== CONFIG.GOOGLE_CLIENT_ID || !issuer || !verified || Number(info.exp) * 1000 < Date.now()) throw new Error('ERR_SIGNIN');
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
  sendLogged_('link', p.name, p.email, it ? 'Il tuo link personale · SE Practice Partners' : 'Your personal link · SE Practice Partners',
    lead + '\n\n' + body + '\n' + url, { name: CONFIG.SENDER_NAME, htmlBody: html }, '');
  cache.put('link_' + key_(p.name), '1', 600);
  return { sent: true };
}

/* Photo uploaded from the app: a small square JPEG prepared in the browser, saved in the Photos sheet */
function uploadPhoto(k, dataUrl) {
  const p = personByKey_(k);
  if (!p) throw new Error('ERR_NOT_YOU');
  if (!/^data:image\/jpeg;base64,[A-Za-z0-9+\/=]+$/.test(String(dataUrl || '')) || dataUrl.length > 48000) {   // the page sends ~10–25 KB
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
  const p = k ? personByKey_(k) : null;
  if (p) touch_(p);   // the page calls this every minute: it doubles as "I'm still here"
  return { now: p ? readNow_() : {}, online: p ? readOnline_(readPeople_()) : [], signedIn: !!p, serverTime: new Date().toISOString() };
}

/* ---------- Who is connected ----------
   Kept in the script cache only (nothing written to the sheet): a timestamp per person, refreshed by the page every minute. */
const ONLINE_MS = 150000;   // seen in the last 2.5 minutes = connected

function touch_(p) { CacheService.getScriptCache().put('seen_' + personKey_(p.name), String(Date.now()), 600); }

function readOnline_(people) {
  const keys = people.map(function (p) { return 'seen_' + personKey_(p.name); });
  const seen = keys.length ? CacheService.getScriptCache().getAll(keys) : {};
  const t = Date.now();
  return people.filter(function (p, i) { return Number(seen[keys[i]] || 0) > t - ONLINE_MS; }).map(function (p) { return p.name; });
}

/* Called when the page is closed */
function leave(k) {
  const p = personByKey_(k);
  if (p) CacheService.getScriptCache().remove('seen_' + personKey_(p.name));
  return { ok: true };
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
    sheet.appendRow([p.name, tz, '']);
    return { tz: tz };
  } finally {
    lock.releaseLock();
  }
}

/* Dashboard language of each person (column C of the Zones sheet): emails to them use it */
function setLanguage(k, lang) {
  const p = actor_(k);
  if (['it', 'en'].indexOf(lang) === -1) throw new Error('ERR_BAD_REQUEST');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = zonesSheet_();
    const values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (key_(values[i][0]) === key_(p.name)) { sheet.getRange(i + 1, 3).setValue(lang); return { lang: lang }; }
    }
    sheet.appendRow([p.name, '', lang]);
    return { lang: lang };
  } finally {
    lock.releaseLock();
  }
}

function readLangs_() {
  const out = {};
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ZONES_SHEET);
  if (!sheet) return out;
  sheet.getDataRange().getValues().slice(1).forEach(function (r) { if (r[0] && MAIL[r[2]]) out[String(r[0])] = String(r[2]); });
  return out;
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
    sheet.getRange(1, 1, 1, 4).setValues([['Name', 'Time zone', 'Language', 'Reminders']]).setFontWeight('bold');
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
    const st = startOf_({ date: input.date, time: input.time });
    if (st && st.getTime() <= Date.now()) throw new Error('ERR_PAST');

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

    bookingsSheet_().appendRow([booking.id, new Date(), booking.by, booking.with, "'" + booking.date, "'" + booking.time,
      booking.format, booking.note, 'proposed', booking.third, '']);
    const data = sessionsData_(input.key);
    data.emailed = notify_('proposed', booking, people, input.lang, by);
    return data;
  } finally {
    lock.releaseLock();
  }
}

/* Each invited person confirms for themselves; the session is "confirmed" once everyone invited has confirmed. */
function confirmBooking(id, k, lang) {
  lang = MAIL[lang] ? lang : 'it';
  const name = actor_(k).name;
  return withBooking_(id, function (b, row, sheet) {
    const who = invitees_(b).find(function (n) { return key_(n) === key_(name); });
    if (!who) throw new Error('ERR_ONLY_INVITED');
    if (b.status === 'cancelled') throw new Error('ERR_CANCELLED');
    const start = startOf_(b);
    if (b.status === 'expired' || (start && start.getTime() <= Date.now())) throw new Error('ERR_EXPIRED');
    if (b.confirmedBy.indexOf(who) !== -1) return { unchanged: true };   // double click / second tab: no new emails
    b.confirmedBy.push(who);
    const all = invitees_(b).every(function (n) { return b.confirmedBy.indexOf(n) !== -1; });
    sheet.getRange(row, 11).setValue(b.confirmedBy.join(', '));
    sheet.getRange(row, 9).setValue(all ? 'confirmed' : 'proposed');
    b.allConfirmed = all;
    if (!all) return { event: 'confirmed', actor: who };
    // Everyone is in: calendar event with a Meet link, then one email to all participants
    if (!b.eventId) {
      const ev = createEvent_(b, readPeople_(), lang);
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
    if (b.status === 'cancelled' || b.status === 'expired') return { unchanged: true };
    sheet.getRange(row, 9).setValue('cancelled');
    if (b.eventId) removeEvent_(b.eventId);
    const declined = key_(who) !== key_(b.by) && b.confirmedBy.indexOf(who) === -1;   // invited and never accepted
    return { event: declined ? 'declined' : 'cancelled', actor: who };
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
        meet: String(r[11] || ''), eventId: String(r[12] || ''), status: String(r[8] || 'proposed') };
      const out = change(b, i + 1, sheet);
      if (out.unchanged) { const same = sessionsData_(k); same.emailed = []; return same; }
      const data = sessionsData_(k);
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

/* ---------- Post-session feedback board ----------
   For each confirmed session that has ended: a thread of comments. By default only its participants can see and write. */

function pastSessions_() {
  const t = Date.now(), from = t - CONFIG.FEEDBACK_DAYS * 86400000;
  const out = [];
  bookingsSheet_().getDataRange().getValues().slice(1).forEach(function (r) {
    if (String(r[8]) !== 'confirmed') return;
    const b = { id: String(r[0]), by: String(r[2]), with: String(r[3]), third: String(r[9] || ''), date: asDate_(r[4]), time: asTime_(r[5]),
                format: String(r[6]) };
    const st = startOf_(b);
    if (!st || st.getTime() + CONFIG.SESSION_MINUTES * 60000 > t || st.getTime() < from) return;
    out.push(b);
  });
  return out.sort(function (a, b) { return (b.date + b.time).localeCompare(a.date + a.time); });   // newest first
}

function feedbackSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.FEEDBACK_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.FEEDBACK_SHEET);
    sheet.getRange(1, 1, 1, 5).setValues([['ID', 'Session', 'Author', 'When', 'Text']]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function canSee_(b, name) {
  return CONFIG.FEEDBACK_PUBLIC || participants_(b).some(function (n) { return key_(n) === key_(name); });
}

function getFeedback(k) {
  const me = actor_(k);
  const sessions = pastSessions_().filter(function (b) { return canSee_(b, me.name); });
  const ids = {};
  sessions.forEach(function (b) { ids[b.id] = true; b.comments = []; });
  const byId = {};
  sessions.forEach(function (b) { byId[b.id] = b; });
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.FEEDBACK_SHEET);
  if (sheet) {
    sheet.getDataRange().getValues().slice(1).forEach(function (r) {
      const sid = String(r[1]);
      if (!ids[sid]) return;
      byId[sid].comments.push({ id: String(r[0]), author: String(r[2]), at: r[3] instanceof Date ? r[3].toISOString() : String(r[3]),
                                text: String(r[4]), mine: key_(r[2]) === key_(me.name) });
    });
  }
  sessions.forEach(function (b) {
    b.canWrite = participants_(b).some(function (n) { return key_(n) === key_(me.name); });
    b.comments.sort(function (a, c) { return a.at.localeCompare(c.at); });
  });
  return { sessions: sessions };
}

function addFeedback(k, sessionId, text, lang) {
  const me = actor_(k);
  const body = String(text || '').trim();
  if (!body) throw new Error('ERR_EMPTY');
  if (body.length > 1500) throw new Error('ERR_TOO_LONG');
  const b = pastSessions_().find(function (x) { return x.id === String(sessionId); });
  if (!b) throw new Error('ERR_NOT_FOUND');
  if (!participants_(b).some(function (n) { return key_(n) === key_(me.name); })) throw new Error('ERR_ONLY_PAIR');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    feedbackSheet_().appendRow([Utilities.getUuid().slice(0, 8), b.id, me.name, new Date(), body]);
  } finally { lock.releaseLock(); }
  notifyFeedback_(b, me, body, lang);
  return getFeedback(k);
}

function deleteFeedback(k, commentId) {
  const me = actor_(k);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = feedbackSheet_();
    const values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]) !== String(commentId)) continue;
      if (key_(values[i][2]) !== key_(me.name)) throw new Error('ERR_NOT_YOURS');
      sheet.deleteRow(i + 1);
      break;
    }
  } finally { lock.releaseLock(); }
  return getFeedback(k);
}

const FEEDBACK_MAIL = {
  it: { subject: function (v) { return v.author + ' ha lasciato un feedback sulla sessione del ' + v.when; },
        lead: function (v) { return v.author + ' ha scritto sulla bacheca della vostra sessione di pratica del ' + v.when + ':'; },
        button: 'Leggi e rispondi' },
  en: { subject: function (v) { return v.author + ' left feedback on the session of ' + v.when; },
        lead: function (v) { return v.author + ' wrote on the board of your practice session of ' + v.when + ':'; },
        button: 'Read and reply' },
};

function notifyFeedback_(b, me, body, lang) {
  if (!CONFIG.SEND_EMAILS) return;
  const people = readPeople_(), langs = readLangs_();
  const first = function (n) { return String(n).split(/\s+/)[0]; };
  participants_(b).filter(function (n) { return key_(n) !== key_(me.name); }).forEach(function (name) {
    const p = people.find(function (x) { return key_(x.name) === key_(name); }) || { name: name };
    const l = MAIL[langs[name]] ? langs[name] : (MAIL[lang] ? lang : 'it');
    const t1 = MAIL[l], m = FEEDBACK_MAIL[l];
    const v = { author: first(me.name), when: dateLabel_(b.date, t1.locale).long };
    const quote = '<blockquote style="margin:0 0 20px;padding:12px 16px;border-left:3px solid #D86018;background:#F7F3F5;border-radius:8px;font-size:15px;white-space:pre-wrap">' + esc_(body) + '</blockquote>';
    const url = personalUrl_(name) + '&tab=feedback';
    const html = mailFrame_(m.lead(v), quote, url, m.button, t1.footer);
    try { sendLogged_('feedback', p.name, p.email, m.subject(v), m.lead(v) + '\n\n' + body + '\n\n' + url, replyOpts_(html, me.email), b.id); } catch (e) { /* logged */ }
  });
}

/* ---------- Email log ---------- */

function logMail_(type, toName, toEmail, subject, result, bookingId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.MAIL_LOG_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.MAIL_LOG_SHEET);
      sheet.getRange(1, 1, 1, 7).setValues([['When', 'Type', 'To', 'Email', 'Subject', 'Result', 'Session']]).setFontWeight('bold');
      sheet.setFrozenRows(1);
      try { sheet.setColumnWidths(1, 7, 150); sheet.setColumnWidth(5, 340); } catch (e) { /* cosmetic only */ }
    }
    sheet.appendRow([new Date(), type, toName || '', toEmail || '', subject || '', result, bookingId || '']);
  } catch (e) { console.error('Mail log not written: ' + e); }
}

/* Sends through MailApp and writes the result in the log */
function sendLogged_(type, toName, toEmail, subject, text, opts, bookingId) {
  if (!toEmail) { logMail_(type, toName, '', subject, 'skipped: no email', bookingId); return false; }
  if (MailApp.getRemainingDailyQuota() < 1) { logMail_(type, toName, toEmail, subject, 'skipped: daily quota reached', bookingId); return false; }
  try {
    MailApp.sendEmail(toEmail, subject, text, opts);
    logMail_(type, toName, toEmail, subject, 'sent', bookingId);
    return true;
  } catch (e) {
    logMail_(type, toName, toEmail, subject, 'error: ' + e.message, bookingId);
    throw e;
  }
}

/* ---------- Calendar event with Google Meet ----------
   Uses the advanced "Google Calendar API" service (identifier: Calendar). Without it, the session is still
   confirmed and emailed, just without the Meet link. */

function createEvent_(b, people, lang) {
  const en = lang === 'en';
  if (!CONFIG.CREATE_MEET || typeof Calendar === 'undefined') return {};
  try {
    const find = function (n) { return people.find(function (p) { return key_(p.name) === key_(n); }) || {}; };
    const emails = participants_(b).map(function (n) { return find(n).email; }).filter(Boolean);
    const online = b.format !== 'In person';
    const resource = {
      summary: (en ? 'SE practice · ' : 'Pratica SE · ') + participants_(b).join(', '),
      description: (b.note ? b.note + '\n\n' : '') + (en ? 'Organised with SE Practice Partners · ' : 'Organizzata con SE Practice Partners · ') + appUrl_(),
      start: { dateTime: localDateTime_(b.date, b.time, 0), timeZone: CONFIG.TIMEZONE },
      end: { dateTime: localDateTime_(b.date, b.time, CONFIG.SESSION_MINUTES), timeZone: CONFIG.TIMEZONE },
      attendees: emails.map(function (e) { return { email: e }; }),
      guestsCanSeeOtherGuests: true,
      // organiser's copy: popup + email 15 min before (participants get their own Calendar defaults with the invitation)
      reminders: { useDefault: false, overrides: [{ method: 'popup', minutes: 15 }, { method: 'email', minutes: 15 }] },
    };
    if (online) {
      resource.conferenceData = { createRequest: { requestId: b.id + '-' + Date.now(), conferenceSolutionKey: { type: 'hangoutsMeet' } } };
    }
    // sendUpdates "all": Google sends its own invitation, so the event (with the Meet link and Google's reminders)
    // shows up in every participant's calendar, not only in the organiser's
    const ev = Calendar.Events.insert(resource, 'primary', { conferenceDataVersion: 1, sendUpdates: 'all' });
    const video = ((ev.conferenceData && ev.conferenceData.entryPoints) || []).find(function (e) { return e.entryPointType === 'video'; });
    return { eventId: ev.id, meet: ev.hangoutLink || (video && video.uri) || '' };
  } catch (e) {
    console.error('Calendar event not created: ' + e);
    return {};
  }
}

function removeEvent_(eventId) {
  if (typeof Calendar === 'undefined') return;
  try { Calendar.Events.remove('primary', eventId, { sendUpdates: 'all' }); }
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
      action: 'Dall\'app puoi confermarla o rifiutarla con un clic.', button: 'Conferma o rifiuta' },
    confirmed: { subject: function (v) { return v.actorFirst + ' ha confermato la sessione di pratica'; },
      lead: function (v) { return v.actor + ' ha confermato la sessione di pratica' + (v.allConfirmed ? ': ora è confermata da tutti.' : '. Si attende ancora la conferma di ' + v.pending + '.'); },
      action: 'Buona pratica!', button: 'Apri l\'app' },
    scheduled: { subject: function (v) { return 'Sessione confermata · ' + v.whenShort + (v.meet ? ' · link Meet' : ''); },
      lead: function (v) { return 'La sessione di pratica è confermata da tutti.' + (v.meet ? ' Ecco il link per collegarvi.' : ''); },
      action: 'Buona pratica!', button: 'Entra su Google Meet', buttonNoMeet: 'Apri l\'app' },
    declined: { subject: function (v) { return v.actorFirst + ' ha rifiutato l\'invito · ' + v.whenShort; },
      lead: function (v) { return v.actor + ' ha rifiutato l\'invito a questa sessione di pratica, che quindi è annullata:'; },
      action: 'Se volete, potete proporne un\'altra in un altro orario dall\'app.', button: 'Apri l\'app' },
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
      action: 'In the app you can confirm or decline it with one click.', button: 'Confirm or decline' },
    confirmed: { subject: function (v) { return v.actorFirst + ' confirmed the practice session'; },
      lead: function (v) { return v.actor + ' confirmed the practice session' + (v.allConfirmed ? ': everyone has now confirmed.' : '. Still waiting for ' + v.pending + '.'); },
      action: 'Enjoy the practice!', button: 'Open the app' },
    scheduled: { subject: function (v) { return 'Session confirmed · ' + v.whenShort + (v.meet ? ' · Meet link' : ''); },
      lead: function (v) { return 'The practice session is confirmed by everyone.' + (v.meet ? ' Here is the link to join.' : ''); },
      action: 'Enjoy the practice!', button: 'Join on Google Meet', buttonNoMeet: 'Open the app' },
    declined: { subject: function (v) { return v.actorFirst + ' declined the invitation · ' + v.whenShort; },
      lead: function (v) { return v.actor + ' declined the invitation to this practice session, so it is cancelled:'; },
      action: 'You can propose another time from the app.', button: 'Open the app' },
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
  const find = function (n) { return people.find(function (p) { return key_(p.name) === key_(n); }) || { name: n }; };
  const first = function (n) { return String(n).split(/\s+/)[0]; };
  const zones = readZones_();
  const langs = readLangs_();
  const abroad = all_(b).some(function (n) { return zones[n] && zones[n] !== CONFIG.TIMEZONE; });
  const all = participants_(b);
  const recipients = everyone ? all : all.filter(function (n) { return key_(n) !== key_(actor); });
  const actorP = find(actor);
  const sent = [];

  recipients.forEach(function (name) {
    try {
      // the language this person uses on the dashboard; otherwise the one of whoever acted
      const t = MAIL[langs[name]] || MAIL[lang] || MAIL.it;
      const tpl = t[event];
      const joinNames = function (list) {
        return list.length < 2 ? list.join('') : list.slice(0, -1).join(', ') + ' ' + t.labels.and + ' ' + list[list.length - 1];
      };
      const d = dateLabel_(b.date, t.locale);
      const to = find(name);
      if (!to.email) { logMail_(event, name, '', '', 'skipped: no email', b.id); return; }
      const others = event === 'proposed' ? all.filter(function (n) { return n !== actor && n !== name; }) : [];
      const pending = invitees_(b).filter(function (n) { return (b.confirmedBy || []).indexOf(n) === -1; });
      const v = { actor: actor, actorFirst: first(actor), others: joinNames(others), allConfirmed: !!b.allConfirmed,
                  pending: joinNames(pending), whenShort: d.short + ' ' + timeLabel_(b.time, t.locale), meet: b.meet || '' };
      const rows = [[t.labels.when, d.long + ' ' + t.labels.at + ' ' + timeLabel_(b.time, t.locale) + (abroad ? ' · ' + t.labels.italy : '')]];
      const tz = zones[name];
      if (tz && tz !== CONFIG.TIMEZONE) {
        const loc = inZone_(b.date, b.time, tz), ld = dateLabel_(loc.date, t.locale);
        rows.push([t.labels.local, ld.long + ' ' + t.labels.at + ' ' + timeLabel_(loc.time, t.locale) + ' · ' + tz.split('/').pop().replace(/_/g, ' ')]);
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
      if (sendLogged_(event, to.name, to.email, tpl.subject(v), text, opts, b.id)) sent.push(to.name);
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

/* "19:00" → "19:00 (sera)" · "07:00" → "7:00 (mattina)" · EN "7:00 pm (evening)": no doubt about morning or evening */
function timeLabel_(hhmm, locale) {
  const p = String(hhmm).split(':').map(Number), h = p[0], m = ('0' + (p[1] || 0)).slice(-2);
  if (isNaN(h)) return String(hhmm);
  const part = h >= 5 && h < 12 ? 0 : h >= 12 && h < 14 ? 1 : h >= 14 && h < 18 ? 2 : h >= 18 && h < 23 ? 3 : 4;   // same bands as the form
  if (locale === 'en') {
    const h12 = h % 12 || 12;
    return h12 + ':' + m + ' ' + (h < 12 ? 'am' : 'pm') + ' (' + ['morning', 'lunchtime', 'afternoon', 'evening', 'night'][part] + ')';
  }
  return h + ':' + m + ' (' + ['mattina', 'pranzo', 'pomeriggio', 'sera', 'notte'][part] + ')';
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

/* ---------- "They just practised: what about you?" ----------
   Hourly trigger. When confirmed sessions have ended, everyone else gets ONE email (at most one a day each),
   naming who practised (first names only), with their personal link and a WhatsApp share for classmates not yet in. */

const DIGEST_MAIL = {
  it: {
    subject: function (n) { return n > 1 ? n + ' sessioni di pratica SE appena concluse: tocca a te? 🌿' : 'Una sessione di pratica SE appena conclusa: tocca a te? 🌿'; },
    lead: function (groups) { return groups + (groups.indexOf(';') !== -1 ? ' hanno appena concluso le loro sessioni di pratica.' : ' hanno appena concluso una sessione di pratica.'); },
    ask: 'Che ne dici di programmarne una anche tu? Guarda chi è libero negli stessi giorni e proponi una sessione in un minuto.',
    button: 'Trova un compagno di pratica',
    shareTitle: 'Conosci un compagno del corso SE che non ha ancora inserito le sue disponibilità?',
    shareText: 'Condividi l\'app su WhatsApp: più siamo, più è facile trovare un compagno.',
    shareButton: 'Condividi su WhatsApp',
    shareMsg: function (app, form) { return 'Ciao! Per trovare compagni di pratica SE c\'è questa app: ' + app + '\nPrima inserisci le tue disponibilità qui: ' + form; },
    stop: 'Non voglio più ricevere questi promemoria', and: 'e',
    footer: 'Ricevi questa email perché hai inserito le tue disponibilità per la pratica SE tra compagni di corso.',
  },
  en: {
    subject: function (n) { return n > 1 ? n + ' SE practice sessions just finished: your turn? 🌿' : 'An SE practice session just finished: your turn? 🌿'; },
    lead: function (groups) { return groups + (groups.indexOf(';') !== -1 ? ' have just finished their practice sessions.' : ' have just finished a practice session.'); },
    ask: 'How about planning one too? See who is free on the same days and propose a session in a minute.',
    button: 'Find a practice partner',
    shareTitle: 'Do you know an SE classmate who hasn\'t shared their availability yet?',
    shareText: 'Share the app on WhatsApp: the more we are, the easier it is to find a partner.',
    shareButton: 'Share on WhatsApp',
    shareMsg: function (app, form) { return 'Hi! To find SE practice partners there is this app: ' + app + '\nFirst add your availability here: ' + form; },
    stop: 'Stop these reminders', and: 'and',
    footer: 'You receive this email because you shared your availability for SE peer practice.',
  },
};

const TRIGGER_VERSION = '2';   // 2 = every 15 minutes (needed for the "in 2 hours" reminder)

function digestNow() {   // periodic trigger: reminders, then the digest
  MEMO_ = {};
  upgradeTriggers_();
  const starting = remindStarting_();
  const soon = remindSoon_();
  const reminders = remindPending_();
  const digest = runDigest_();
  const result = 'starting: ' + starting + ' · soon: ' + soon + ' · reminders: ' + reminders + ' · digest: ' + digest;
  const props = PropertiesService.getScriptProperties();
  props.setProperty('LAST_RUN', new Date().toISOString()); props.setProperty('LAST_RUN_RESULT', result);
  return result;
}

/* The reminders need the 15-minute trigger. If it is missing (installTriggers never run), any visit to the app installs it.
   Checked at most once every 6 hours, so it costs nothing on normal requests. */
function ensureScheduler_() {
  const props = PropertiesService.getScriptProperties();
  const checked = Number(props.getProperty('SCHEDULER_CHECKED') || 0);
  if (props.getProperty('TRIGGER_VERSION') === TRIGGER_VERSION && Date.now() - checked < 6 * 3600000) return;
  try {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(3000)) return;
    try {
      const handlers = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
      if (handlers.indexOf('digestNow') === -1 || props.getProperty('TRIGGER_VERSION') !== TRIGGER_VERSION) {
        ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'digestNow') ScriptApp.deleteTrigger(t); });
        ScriptApp.newTrigger('digestNow').timeBased().everyMinutes(15).create();
        props.setProperty('TRIGGER_VERSION', TRIGGER_VERSION);
      }
      if (handlers.indexOf('backupNow') === -1) ScriptApp.newTrigger('backupNow').timeBased().everyDays(1).atHour(3).inTimezone(CONFIG.TIMEZONE).create();
      props.setProperty('SCHEDULER_CHECKED', String(Date.now()));
    } finally { lock.releaseLock(); }
  } catch (e) { console.error('Scheduler check failed: ' + e); }
}

/* Older installs ran this hourly: switch to every 15 minutes by itself, once */
function upgradeTriggers_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('TRIGGER_VERSION') === TRIGGER_VERSION) return;
  try {
    ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'digestNow') ScriptApp.deleteTrigger(t); });
    ScriptApp.newTrigger('digestNow').timeBased().everyMinutes(15).create();
    props.setProperty('TRIGGER_VERSION', TRIGGER_VERSION);
  } catch (e) { console.error('Trigger upgrade failed: ' + e); }
}

function runDigest_() {
  if (!CONFIG.DIGEST || !CONFIG.SEND_EMAILS) return 'off';
  const hour = Number(Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'H'));
  if (hour < 8 || hour >= 22) return 'quiet hours';
  const props = PropertiesService.getScriptProperties();
  const last = Number(props.getProperty('DIGEST_LAST') || 0);
  if (Date.now() - last < CONFIG.DIGEST_EVERY_HOURS * 3600000) return 'already sent today';

  const sheet = bookingsSheet_();
  const values = sheet.getDataRange().getValues();
  const t = Date.now();
  const done = [];
  values.slice(1).forEach(function (r, i) {
    if (String(r[8]) !== 'confirmed' || r[13]) return;
    const b = { by: String(r[2]), with: String(r[3]), third: String(r[9] || ''), date: asDate_(r[4]), time: asTime_(r[5]) };
    const start = startOf_(b);
    if (!start) return;
    const end = start.getTime() + CONFIG.SESSION_MINUTES * 60000;
    if (end > t) return;
    if (t - end > 3 * 86400000) { sheet.getRange(i + 2, 14).setValue('too old'); return; }   // never announce stale sessions
    done.push({ row: i + 2, b: b });
  });
  if (!done.length) return 'nothing new';

  const people = readPeople_();
  const langs = readLangs_();
  const optOut = readOptOut_();
  const practised = {};
  done.forEach(function (d) { participants_(d.b).forEach(function (n) { practised[key_(n)] = true; }); });
  const first = function (n) { return String(n).split(/\s+/)[0]; };
  const app = appUrl_(), form = CONFIG.FORM_URL;
  var count = 0;

  people.forEach(function (p) {
    if (!p.email || practised[key_(p.name)] || optOut[p.name]) return;
    if (MailApp.getRemainingDailyQuota() <= CONFIG.DIGEST_RESERVE) return;
    const lang = DIGEST_MAIL[langs[p.name]] ? langs[p.name] : 'it';
    const m = DIGEST_MAIL[lang];
    const join = function (list) { return list.length < 2 ? list.join('') : list.slice(0, -1).join(', ') + ' ' + m.and + ' ' + list[list.length - 1]; };
    const groups = done.map(function (d) { return join(participants_(d.b).map(first)); }).join('; ');
    const url = personalUrl_(p.name);
    const stopUrl = url + '&reminders=off';
    const wa = 'https://wa.me/?text=' + encodeURIComponent(m.shareMsg(app, form));
    const inner =
      '<p style="margin:0 0 20px;font-size:15px;line-height:1.5">' + esc_(m.ask) + '</p>' +
      '<a href="' + esc_(url) + '" style="display:inline-block;background:#D86018;color:#fff;text-decoration:none;font-weight:bold;padding:12px 22px;border-radius:12px">' + esc_(m.button) + '</a>' +
      '<div style="margin-top:26px;padding:16px 18px;border-radius:12px;background:#F3E5EC">' +
      '<p style="margin:0 0 6px;font-weight:bold;color:#6C1D45">' + esc_(m.shareTitle) + '</p>' +
      '<p style="margin:0 0 12px;font-size:14px">' + esc_(m.shareText) + '</p>' +
      '<a href="' + esc_(wa) + '" style="display:inline-block;background:#25D366;color:#fff;text-decoration:none;font-weight:bold;padding:10px 18px;border-radius:10px">' + esc_(m.shareButton) + '</a>' +
      '<p style="margin:10px 0 0;font-size:13px;color:#6E686B">' + esc_(app) + '</p></div>';
    const html = digestFrame_(m.lead(groups), inner, m.footer + ' · <a href="' + esc_(stopUrl) + '" style="color:#6E686B">' + esc_(m.stop) + '</a>');
    const text = [m.lead(groups), '', m.ask, url, '', m.shareTitle, m.shareText, app, '', '—', m.footer, m.stop + ': ' + stopUrl].join('\n');
    try {
      if (sendLogged_('digest', p.name, p.email, m.subject(done.length), text, { name: CONFIG.SENDER_NAME, htmlBody: html }, '')) count++;
    } catch (e) { /* logged */ }
  });

  done.forEach(function (d) { sheet.getRange(d.row, 14).setValue(new Date()); });
  props.setProperty('DIGEST_LAST', String(Date.now()));
  return count + ' emails';
}

function digestFrame_(lead, inner, footerHtml) {
  return '<div style="background:#F7F3F5;padding:28px 16px;font-family:Roboto,Helvetica,Arial,sans-serif;color:#3D3D3D">' +
    '<div style="max-width:520px;margin:0 auto;background:#fff;border:1px solid #E5DCE1;border-radius:16px;padding:28px">' +
    '<p style="margin:0 0 6px;font-size:11px;letter-spacing:.12em;text-transform:uppercase;font-weight:bold;color:#D86018">SE Practice Partners</p>' +
    '<p style="margin:0 0 14px;font-family:Comfortaa,Trebuchet MS,Helvetica,sans-serif;font-size:20px;color:#6C1D45;line-height:1.35">' + esc_(lead) + '</p>' +
    inner + '</div><p style="max-width:520px;margin:16px auto 0;font-size:12px;color:#6E686B;text-align:center">' + footerHtml + '</p></div>';
}

/* ---------- 24 h reminder for sessions not yet confirmed ---------- */

const REMIND_MAIL = {
  it: {
    invSubject: function (v) { return 'Conferma la sessione di ' + v.whenShort + ' con ' + v.byFirst; },
    invLead: function (v) { return 'La sessione di pratica proposta da ' + v.by + ' non è ancora confermata.'; },
    invAction: 'Dall\'app puoi confermarla o rifiutarla con un clic. Se non rispondi entro l\'orario d\'inizio, verrà annullata automaticamente.',
    invButton: 'Conferma o rifiuta',
    bySubject: function (v) { return 'Sessione di ' + v.whenShort + ': manca la conferma di ' + v.pending; },
    byLead: function (v) { return v.pending + (v.many ? ' non hanno' : ' non ha') + ' ancora confermato la sessione che hai proposto.'; },
    byAction: 'Se la conferma non arriva entro l\'orario d\'inizio, la sessione verrà annullata. Puoi ricordarlo anche su WhatsApp.',
    byButton: 'Apri l\'app', tomorrow: 'domani alle', today: 'oggi alle',
  },
  en: {
    invSubject: function (v) { return 'Confirm the session ' + v.whenShort + ' with ' + v.byFirst; },
    invLead: function (v) { return 'The practice session proposed by ' + v.by + ' is not confirmed yet.'; },
    invAction: 'In the app you can confirm or decline it with one click. If you don\'t answer before it starts, it will be cancelled automatically.',
    invButton: 'Confirm or decline',
    bySubject: function (v) { return 'Session ' + v.whenShort + ': waiting for ' + v.pending; },
    byLead: function (v) { return v.pending + (v.many ? ' have' : ' has') + ' not confirmed the session you proposed yet.'; },
    byAction: 'If the confirmation doesn\'t arrive before it starts, the session will be cancelled. You can also remind them on WhatsApp.',
    byButton: 'Open the app', tomorrow: 'tomorrow at', today: 'today at',
  },
};

function replyOpts_(html, replyTo) {
  const o = { name: CONFIG.SENDER_NAME, htmlBody: html };
  if (replyTo) o.replyTo = replyTo;
  return o;
}

function remindPending_() {
  if (!CONFIG.REMIND_HOURS || !CONFIG.SEND_EMAILS) return 'off';
  const hour = Number(Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'H'));
  if (hour < 8 || hour >= 22) return 'quiet hours';
  const sheet = bookingsSheet_();
  const values = sheet.getDataRange().getValues();
  const t = Date.now();
  var people, langs, zones;
  const load = function () { if (!people) { people = readPeople_(); langs = readLangs_(); zones = readZones_(); } };
  const find = function (n) { load(); return people.find(function (p) { return key_(p.name) === key_(n); }) || { name: n }; };
  const first = function (n) { return String(n).split(/\s+/)[0]; };
  const todayRome = today_();
  var sent = 0;

  values.slice(1).forEach(function (r, i) {
    if (String(r[8] || 'proposed') !== 'proposed' || r[14]) return;
    const b = { id: String(r[0]), by: String(r[2]), with: String(r[3]), date: asDate_(r[4]), time: asTime_(r[5]),
      format: String(r[6]), note: String(r[7]), third: String(r[9] || ''),
      confirmedBy: String(r[10] || '').split(',').map(function (x) { return x.trim(); }).filter(String) };
    const start = startOf_(b);
    if (!start || start.getTime() <= t || start.getTime() - t > CONFIG.REMIND_HOURS * 3600000) return;
    const created = r[1] && typeof r[1].getTime === 'function' ? r[1].getTime() : (Date.parse(r[1]) || 0);
    if (created && t - created < 3 * 3600000) return;   // just proposed: the proposal email is enough
    const pending = invitees_(b).filter(function (n) { return b.confirmedBy.indexOf(n) === -1; });
    if (!pending.length) return;

    load();
    const send = function (name, isProposer) {
      const p = find(name);
      const lang = MAIL[langs[name]] ? langs[name] : 'it';
      const t1 = MAIL[lang], m = REMIND_MAIL[lang];
      const join = function (list) { return list.length < 2 ? list.join('') : list.slice(0, -1).join(', ') + ' ' + t1.labels.and + ' ' + list[list.length - 1]; };
      const d = dateLabel_(b.date, t1.locale);
      const v = { by: b.by, byFirst: first(b.by), pending: join(pending.map(first)), many: pending.length > 1,
                  whenShort: (b.date === todayRome ? m.today : m.tomorrow) + ' ' + timeLabel_(b.time, t1.locale) };
      if (b.date !== todayRome && start.getTime() - t > 36 * 3600000) v.whenShort = d.short + ' ' + timeLabel_(b.time, t1.locale);
      const rows = [[t1.labels.when, d.long + ' ' + t1.labels.at + ' ' + timeLabel_(b.time, t1.locale)], [t1.labels.format, t1.formats[b.format] || b.format],
                    [t1.labels.who, join(participants_(b))]];
      const tz = zones[name];
      if (tz && tz !== CONFIG.TIMEZONE) {
        const loc = inZone_(b.date, b.time, tz), ld = dateLabel_(loc.date, t1.locale);
        rows.splice(1, 0, [t1.labels.local, ld.long + ' ' + t1.labels.at + ' ' + timeLabel_(loc.time, t1.locale) + ' · ' + tz.split('/').pop().replace(/_/g, ' ')]);
      }
      if (b.note) rows.push([t1.labels.note, b.note]);
      const lead = isProposer ? m.byLead(v) : m.invLead(v);
      const action = isProposer ? m.byAction : m.invAction;
      const url = personalUrl_(name);
      const table = '<table style="border-collapse:collapse;width:100%;font-size:15px;margin-bottom:20px">' +
        rows.map(function (x) {
          return '<tr><td style="padding:8px 0;border-top:1px solid #EFE8EC;color:#6E686B;width:110px;vertical-align:top">' + esc_(x[0]) +
            '</td><td style="padding:8px 0;border-top:1px solid #EFE8EC">' + esc_(x[1]) + '</td></tr>';
        }).join('') + '</table>' +
        '<p style="margin:0 0 18px;font-size:15px;padding:10px 12px;border-radius:10px;background:#FBE9DD;color:#9A410C">' + esc_(action) + '</p>';
      const html = mailFrame_(lead, table, url, isProposer ? m.byButton : m.invButton, t1.footer);
      const text = [lead, ''].concat(rows.map(function (x) { return x[0] + ': ' + x[1]; }), ['', action, url, '', '—', t1.footer]).join('\n');
      try {
        if (sendLogged_('reminder', p.name, p.email, isProposer ? m.bySubject(v) : m.invSubject(v), text,
            replyOpts_(html, isProposer ? '' : find(b.by).email), b.id)) sent++;
      } catch (e) { /* logged */ }
    };
    pending.forEach(function (n) { send(n, false); });
    send(b.by, true);
    sheet.getRange(i + 2, 15).setValue(new Date());
  });
  return sent + ' emails';
}

/* ---------- "In 2 hours" reminder for confirmed sessions ----------
   To every participant. Early sessions (before 9:00) get it the evening before, around 21:00, not at dawn. */

const SOON_MAIL = {
  it: {
    subject: function (v) { return (v.evening ? 'Domani' : 'Tra poco') + ': sessione di pratica SE alle ' + v.time + ' con ' + v.others; },
    lead: function (v) { return (v.evening ? 'Promemoria per domani: ' : 'Tra circa ' + v.hours + ' ore: ') + 'sessione di pratica con ' + v.others + '.'; },
    action: 'Buona pratica! Se non potete più, avvisatevi e annullatela dall\'app.',
    meetButton: 'Entra su Google Meet', appButton: 'Apri l\'app',
  },
  en: {
    subject: function (v) { return (v.evening ? 'Tomorrow' : 'Soon') + ': SE practice session at ' + v.time + ' with ' + v.others; },
    lead: function (v) { return (v.evening ? 'Reminder for tomorrow: ' : 'In about ' + v.hours + ' hours: ') + 'practice session with ' + v.others + '.'; },
    action: 'Enjoy the practice! If you can no longer make it, let each other know and cancel it in the app.',
    meetButton: 'Join on Google Meet', appButton: 'Open the app',
  },
};

function remindSoon_() {
  if (!CONFIG.SOON_HOURS || !CONFIG.SEND_EMAILS) return 'off';
  const nowHour = Number(Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'H'));
  if (nowHour < 7 || nowHour >= 22) return 'quiet hours';
  const sheet = bookingsSheet_();
  const values = sheet.getDataRange().getValues();
  const t = Date.now();
  var people, langs, zones;
  const load = function () { if (!people) { people = readPeople_(); langs = readLangs_(); zones = readZones_(); } };
  const find = function (n) { load(); return people.find(function (p) { return key_(p.name) === key_(n); }) || { name: n }; };
  const first = function (n) { return String(n).split(/\s+/)[0]; };
  var sent = 0;

  values.slice(1).forEach(function (r, i) {
    if (String(r[8]) !== 'confirmed' || r[15]) return;
    const b = { id: String(r[0]), by: String(r[2]), with: String(r[3]), date: asDate_(r[4]), time: asTime_(r[5]),
      format: String(r[6]), note: String(r[7]), third: String(r[9] || ''), meet: String(r[11] || '') };
    const start = startOf_(b);
    if (!start || start.getTime() <= t) return;
    const left = start.getTime() - t;
    const startHour = Number(b.time.split(':')[0]);
    const early = startHour < 9;
    // normal: within SOON_HOURS · early mornings: the evening before (from 20:00, within 13 h)
    const evening = early && nowHour >= 20 && left <= 13 * 3600000 && left > CONFIG.SOON_HOURS * 3600000;
    if (!(left <= CONFIG.SOON_HOURS * 3600000 || evening)) return;

    load();
    participants_(b).forEach(function (name) {
      const p = find(name);
      const lang = MAIL[langs[name]] ? langs[name] : 'it';
      const t1 = MAIL[lang], m = SOON_MAIL[lang];
      const join = function (list) { return list.length < 2 ? list.join('') : list.slice(0, -1).join(', ') + ' ' + t1.labels.and + ' ' + list[list.length - 1]; };
      const others = participants_(b).filter(function (n) { return n !== name; }).map(first);
      const v = { others: join(others), time: timeLabel_(b.time, t1.locale), evening: evening, hours: Math.max(1, Math.round(left / 3600000)) };
      const d = dateLabel_(b.date, t1.locale);
      const rows = [[t1.labels.when, d.long + ' ' + t1.labels.at + ' ' + timeLabel_(b.time, t1.locale)]];
      const tz = zones[name];
      if (tz && tz !== CONFIG.TIMEZONE) {
        const loc = inZone_(b.date, b.time, tz), ld = dateLabel_(loc.date, t1.locale);
        rows[0][1] += ' · ' + t1.labels.italy;
        rows.push([t1.labels.local, ld.long + ' ' + t1.labels.at + ' ' + timeLabel_(loc.time, t1.locale) + ' · ' + tz.split('/').pop().replace(/_/g, ' ')]);
      }
      rows.push([t1.labels.format, t1.formats[b.format] || b.format], [t1.labels.who, join(participants_(b))]);
      if (b.meet) rows.push([t1.labels.meet, b.meet]);
      if (b.note) rows.push([t1.labels.note, b.note]);
      const url = b.meet || personalUrl_(name);
      const table = '<table style="border-collapse:collapse;width:100%;font-size:15px;margin-bottom:20px">' +
        rows.map(function (x) {
          return '<tr><td style="padding:8px 0;border-top:1px solid #EFE8EC;color:#6E686B;width:110px;vertical-align:top">' + esc_(x[0]) +
            '</td><td style="padding:8px 0;border-top:1px solid #EFE8EC">' + esc_(x[1]) + '</td></tr>';
        }).join('') + '</table><p style="margin:0 0 18px;font-size:15px">' + esc_(m.action) + '</p>';
      const html = mailFrame_(m.lead(v), table, url, b.meet ? m.meetButton : m.appButton, t1.footer);
      const text = [m.lead(v), ''].concat(rows.map(function (x) { return x[0] + ': ' + x[1]; }), ['', m.action, url, '', '—', t1.footer]).join('\n');
      try { if (sendLogged_('soon', p.name, p.email, m.subject(v), text, replyOpts_(html, ''), b.id)) sent++; } catch (e) { /* logged */ }
    });
    sheet.getRange(i + 2, 16).setValue(new Date());
  });
  return sent + ' emails';
}

/* ---------- "Starting soon" email, about 15 minutes before a confirmed session ----------
   The scheduler runs every 15 minutes, so it goes out between ~15 and ~30 minutes before the start. Any hour of the day:
   it is what people asked for, right before the session. */

const START_MAIL = {
  it: { subject: function (v) { return 'Sta per iniziare: sessione di pratica SE alle ' + v.time; },
        lead: function (v) { return 'Tra pochi minuti inizia la tua sessione di pratica con ' + v.others + '.'; },
        meetButton: 'Entra su Google Meet', appButton: 'Apri l\'app', inPerson: 'La sessione è in presenza.' },
  en: { subject: function (v) { return 'Starting soon: SE practice session at ' + v.time; },
        lead: function (v) { return 'Your practice session with ' + v.others + ' starts in a few minutes.'; },
        meetButton: 'Join on Google Meet', appButton: 'Open the app', inPerson: 'The session is in person.' },
};

function remindStarting_() {
  if (!CONFIG.START_MINUTES || !CONFIG.SEND_EMAILS) return 'off';
  const sheet = bookingsSheet_();
  const values = sheet.getDataRange().getValues();
  const t = Date.now();
  const due = [];
  values.slice(1).forEach(function (r, i) {
    if (String(r[8]) !== 'confirmed' || r[16]) return;
    const b = { id: String(r[0]), by: String(r[2]), with: String(r[3]), date: asDate_(r[4]), time: asTime_(r[5]),
      format: String(r[6]), note: String(r[7]), third: String(r[9] || ''), meet: String(r[11] || '') };
    const st = startOf_(b);
    if (!st) return;
    const left = st.getTime() - t;
    if (left > 0 && left <= (CONFIG.START_MINUTES + 15) * 60000) due.push({ row: i + 2, b: b });
  });
  if (!due.length) return '0 emails';
  const people = readPeople_(), langs = readLangs_(), zones = readZones_();
  const find = function (n) { return people.find(function (p) { return key_(p.name) === key_(n); }) || { name: n }; };
  const first = function (n) { return String(n).split(/\s+/)[0]; };
  var sent = 0;
  due.forEach(function (d) {
    const b = d.b;
    participants_(b).forEach(function (name) {
      const p = find(name);
      const lang = MAIL[langs[name]] ? langs[name] : 'it';
      const t1 = MAIL[lang], m = START_MAIL[lang];
      const join = function (list) { return list.length < 2 ? list.join('') : list.slice(0, -1).join(', ') + ' ' + t1.labels.and + ' ' + list[list.length - 1]; };
      const tz = zones[name];
      const shown = tz && tz !== CONFIG.TIMEZONE ? inZone_(b.date, b.time, tz).time : b.time;   // their own clock
      const v = { others: join(participants_(b).filter(function (n) { return n !== name; }).map(first)), time: timeLabel_(shown, t1.locale) };
      const rows = [];
      if (b.meet) rows.push([t1.labels.meet, b.meet]); else rows.push([t1.labels.format, m.inPerson]);
      if (b.note) rows.push([t1.labels.note, b.note]);
      const url = b.meet || personalUrl_(name);
      const table = '<table style="border-collapse:collapse;width:100%;font-size:15px;margin-bottom:20px">' +
        rows.map(function (x) {
          return '<tr><td style="padding:8px 0;border-top:1px solid #EFE8EC;color:#6E686B;width:110px;vertical-align:top">' + esc_(x[0]) +
            '</td><td style="padding:8px 0;border-top:1px solid #EFE8EC">' + esc_(x[1]) + '</td></tr>';
        }).join('') + '</table>';
      const html = mailFrame_(m.lead(v), table, url, b.meet ? m.meetButton : m.appButton, t1.footer);
      const text = [m.lead(v), ''].concat(rows.map(function (x) { return x[0] + ': ' + x[1]; }), ['', url, '', '—', t1.footer]).join('\n');
      try { if (sendLogged_('starting', p.name, p.email, m.subject(v), text, replyOpts_(html, ''), b.id)) sent++; } catch (e) { /* logged */ }
    });
    sheet.getRange(d.row, 17).setValue(new Date());
  });
  return sent + ' emails';
}

/* Opt-out of the digest (column D of the Zones sheet) */
function setDigest(k, on) {
  const p = actor_(k);
  const v = on ? '' : 'off';
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = zonesSheet_();
    const values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (key_(values[i][0]) === key_(p.name)) { sheet.getRange(i + 1, 4).setValue(v); return { reminders: !!on }; }
    }
    sheet.appendRow([p.name, '', '', v]);
    return { reminders: !!on };
  } finally {
    lock.releaseLock();
  }
}

function readOptOut_() {
  const out = {};
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.ZONES_SHEET);
  if (!sheet) return out;
  sheet.getDataRange().getValues().slice(1).forEach(function (r) { if (r[0] && String(r[3]) === 'off') out[String(r[0])] = true; });
  return out;
}

/* ---------- Nightly backup ----------
   Run installBackup() once from the editor: it schedules backupNow() every night at about 3:00 and makes a first copy.
   The copy is an .xlsx export (copying a form-linked spreadsheet would also duplicate the form). */

/* Run once from the editor: nightly backup + hourly check for finished sessions (digest email) */
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (['backupNow', 'digestNow'].indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('backupNow').timeBased().everyDays(1).atHour(3).inTimezone(CONFIG.TIMEZONE).create();
  ScriptApp.newTrigger('digestNow').timeBased().everyMinutes(15).create();
  PropertiesService.getScriptProperties().setProperty('TRIGGER_VERSION', TRIGGER_VERSION);
  return backupNow();
}
function installBackup() { return installTriggers(); }   // older name, same thing

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
/* Sheets may turn "2026-11-05" / "19:00" into real dates and times. Those must be read back in the SPREADSHEET's own
   time zone: a time alone is stored on 30/12/1899, when Rome's offset was not yet a whole hour (19:00 became 20:21). */
function sheetTz_() { return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || CONFIG.TIMEZONE; }
function asDate_(v) { return v instanceof Date ? Utilities.formatDate(v, sheetTz_(), 'yyyy-MM-dd') : String(v).replace(/^'/, ''); }
function asTime_(v) { return v instanceof Date ? Utilities.formatDate(v, sheetTz_(), 'HH:mm') : String(v).replace(/^'/, ''); }
