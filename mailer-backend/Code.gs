/**
 * מערכת מיילים PRIME — שרת שליחה (Google Apps Script)
 * פרויקט תחת primerct2026@gmail.com
 *
 * התקנה: להריץ פעם אחת את הפונקציה setup()
 * היא יוצרת את תיקיית הקבצים, מתקינה את טריגר התזמון, ומדפיסה את הקישור לתיקייה.
 */

const SECRET = 'lgGnSJZnAsfIs4W822y0k7F6';
const SENDER_NAME = 'מחקר PRIME';
const FOLDER_NAME = 'PRIME Mailer Files';

const ALLOWED_EMAILS = ['amirsportdiet@gmail.com', 'primerct2026@gmail.com'];
const FIREBASE_API_KEY = 'AIzaSyCZR4jxDQ8hcmfslbgx06dlRQNIoTf3Wss';

// ===================== התקנה =====================
function setup() {
  const folder = filesFolder_();
  installTrigger_();
  const url = 'https://drive.google.com/drive/folders/' + folder.getId();
  Logger.log('✅ הכל מוכן!');
  Logger.log('תיקיית הקבצים: ' + url);
  Logger.log('יש להעלות לתיקייה את 7 קובצי ה-PDF.');
  Logger.log('קבצים שכבר בתיקייה: ' + (listFiles_().join(', ') || '(ריקה)'));
  Logger.log('יומן שליחות: https://docs.google.com/spreadsheets/d/' + logSheet_().getParent().getId());
  return url;
}

/** בדיקה: מציג אילו קבצים נמצאים בתיקייה */
function listFiles_() {
  const it = filesFolder_().getFiles();
  const out = [];
  while (it.hasNext()) out.push(it.next().getName());
  return out;
}
function whatFilesDoIHave() { Logger.log(listFiles_().join('\n') || '(התיקייה ריקה)'); }

/** מאתר את תיקיית הקבצים, ויוצר אותה אם אינה קיימת */
function filesFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FILES_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* נמחקה — ניצור מחדש */ }
  }
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('FILES_FOLDER_ID', folder.getId());
  return folder;
}

// ===================== אבטחה =====================
/**
 * מאמת את אסימון הזהות מול Google ומחזיר את המייל, או null.
 * זו שכבת האבטחה האמיתית — המפתח הסודי גלוי בקוד העמוד ולכן אינו מספיק לבדו.
 */
function verifiedEmail_(idToken) {
  if (!idToken) return { err: 'האסימון לא הגיע לשרת' };
  var res;
  try {
    res = UrlFetchApp.fetch(
      'https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + FIREBASE_API_KEY,
      { method: 'post', contentType: 'application/json',
        payload: JSON.stringify({ idToken: idToken }), muteHttpExceptions: true });
  } catch (e) {
    return { err: 'הקריאה לגוגל נכשלה (ייתכן שחסרה הרשאה): ' + e };
  }
  const code = res.getResponseCode();
  const txt = res.getContentText();
  if (code !== 200) {
    return { err: 'גוגל החזירה ' + code + ' — ' + txt.slice(0, 180) + ' | אורך אסימון: ' + idToken.length };
  }
  const users = (JSON.parse(txt) || {}).users || [];
  if (!users.length || !users[0].email) return { err: 'לא נמצא מייל בתשובה של גוגל' };
  return { email: String(users[0].email).toLowerCase() };
}

// ===================== נקודת הכניסה מהעמוד =====================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.secret !== SECRET) return json({ ok: false, error: 'unauthorized' });

    const v = verifiedEmail_(data.idToken);
    if (!v.email) return json({ ok: false, error: 'זיהוי נכשל — ' + v.err });
    if (ALLOWED_EMAILS.indexOf(v.email) < 0) {
      return json({ ok: false, error: 'החשבון ' + v.email + ' אינו מורשה' });
    }
    data.sentBy = v.email;

    // --- ניהול מיילים ממתינים ---
    if (data.mode === 'history')    return json({ ok: true, items: historyFor_(data.email) });
    if (data.mode === 'list')       return json({ ok: true, items: listPending_() });
    if (data.mode === 'cancel')     return json(cancelPending_(data.row));
    if (data.mode === 'reschedule') return json(reschedulePending_(data.row, data.sendAt));

    // --- בדיקת כפילות: אותו נחקר + אותו שלב + אותו סוג שליחה ---
    if (!data.force) {
      const dup = findDuplicate_(data);
      if (dup) {
        return json({ ok: false, duplicate: true,
          when: dup.when, by: dup.by, mode: dup.mode });
      }
    }

    if (data.mode === 'now') {
      sendMail_(data);
      logSend_(data, 'נשלח');
      return json({ ok: true, sent: true });
    }
    queueScheduled_(data);
    logSend_(data, 'תוזמן');
    return json({ ok: true, scheduled: true, sendAt: data.sendAt });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// ===================== שליחה =====================
function sendMail_(data) {
  GmailApp.sendEmail(data.to.join(','), data.subject, data.body, {
    name: SENDER_NAME,
    htmlBody: rtlHtml_(data.body),          // עברית מיושרת לימין
    attachments: attachmentsFor_(data.attachments)
  });
}

/** עוטף את הטקסט ב-HTML עם כיווניות RTL ויישור לימין */
function rtlHtml_(text) {
  var esc = String(text || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  return '<div dir="rtl" style="direction:rtl;text-align:right;unicode-bidi:embed;' +
         'font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.75;' +
         'color:#1a1a1a;white-space:pre-wrap;">' + esc + '</div>';
}

/** מאתר בתיקייה את הקבצים לפי שמם */
function attachmentsFor_(names) {
  if (!names || !names.length) return [];
  const folder = filesFolder_();
  const out = [];
  for (var i = 0; i < names.length; i++) {
    const it = folder.getFilesByName(names[i]);
    if (it.hasNext()) out.push(it.next().getBlob());
    else throw new Error('לא נמצא קובץ בתיקייה: ' + names[i]);
  }
  return out;
}

// ===================== יומן שליחות + כפילויות =====================
const LOG_HEADERS = ['מתי', 'אל', 'שם', 'שלב', 'קבוצה', 'מוקד', 'סוג', 'מועד מתוזמן', 'נשלח ע״י', 'נושא'];

function logSheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('LOG_SHEET_ID');
  let ss;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('PRIME Mailer — יומן שליחות');
    props.setProperty('LOG_SHEET_ID', ss.getId());
    ss.getSheets()[0].appendRow(LOG_HEADERS);
    ss.getSheets()[0].setFrozenRows(1);
  }
  return ss.getSheets()[0];
}

/** מפתח ייחודי לשליחה: נמען + שלב + סוג */
function sendKey_(d) {
  return String((d.to || []).join(',')).toLowerCase() + '|' + d.phase + '|' + d.mode;
}

/** מחזיר את הרשומה הקודמת אם כבר בוצעה שליחה זהה, אחרת null */
function findDuplicate_(d) {
  const sh = logSheet_();
  const rows = sh.getDataRange().getValues();
  const key = sendKey_(d);
  for (let i = rows.length - 1; i >= 1; i--) {
    const r = rows[i];
    const rowKey = String(r[1]).toLowerCase() + '|' + r[3] + '|' + (r[6] === 'תוזמן' ? 'schedule' : 'now');
    if (rowKey === key) {
      return { when: Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
               by: r[8], mode: r[6] };
    }
  }
  return null;
}

function logSend_(d, kind) {
  logSheet_().appendRow([new Date(), (d.to || []).join(','), d.name || '', d.phase || '',
    d.group || '—', d.site || '', kind, d.sendAt || '', d.sentBy || '', d.subject || '']);
}

/** פתיחת היומן — הרץ ידנית כדי לקבל את הקישור */
function openLog() { Logger.log('https://docs.google.com/spreadsheets/d/' + logSheet_().getParent().getId()); }

/** היסטוריית השליחות של נחקר לפי כתובת מייל */
function historyFor_(email) {
  const e = String(email || '').toLowerCase().trim();
  if (!e) return [];
  const sh = logSheet_();
  const rows = sh.getDataRange().getValues();
  const out = [];
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).toLowerCase().indexOf(e) < 0) continue;
    out.push({
      when: rows[i][0] instanceof Date
        ? Utilities.formatDate(rows[i][0], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
        : String(rows[i][0]),
      name: rows[i][2], phase: rows[i][3], group: rows[i][4],
      site: rows[i][5], kind: rows[i][6], sendAt: normSendAt_(rows[i][7])
    });
  }
  return out.reverse();   // החדש ביותר ראשון
}

// ===================== יומן PRIME =====================
const CALENDAR_NAME = 'PRIME';

function primeCalendar_() {
  const byName = CalendarApp.getCalendarsByName(CALENDAR_NAME);
  if (byName.length) return byName[0];
  // גיבוי: יומן שהשם שלו מכיל PRIME
  const all = CalendarApp.getAllCalendars();
  for (var i = 0; i < all.length; i++) {
    if (all[i].getName().toUpperCase().indexOf('PRIME') >= 0) return all[i];
  }
  return null;
}

/** מזהה שלב מתוך טקסט האירוע: T0 / T3 / T6 */
function phaseFromText_(t) {
  const s = String(t || '').toUpperCase();
  if (/\bT0\b/.test(s)) return 'start';
  if (/\bT3\b/.test(s)) return 'middle';
  if (/\bT6\b/.test(s)) return 'end';
  return '';
}

/** מזהה מוקד מתוך טקסט: אסותא / איכילוב */
function siteFromText_(t) {
  const s = String(t || '');
  if (s.indexOf('איכילוב') >= 0 || /ichilov|sourasky|tasmc/i.test(s)) return 'ichilov';
  if (s.indexOf('אסותא') >= 0 || /assuta/i.test(s)) return 'assuta';
  return '';
}

/** מחפש ביומן PRIME את האירוע הקרוב שמתאים לשם */
function lookupCalendar_(name) {
  const cal = primeCalendar_();
  if (!cal) return { ok: false, error: 'לא נמצא יומן בשם ' + CALENDAR_NAME };
  const now = new Date();
  const end = new Date(now.getTime() + 400 * 24 * 3600 * 1000);
  const q = String(name || '').trim();
  const evs = cal.getEvents(now, end);
  const out = [];
  for (var i = 0; i < evs.length && out.length < 10; i++) {
    const ev = evs[i];
    const title = ev.getTitle() || '';
    const desc = ev.getDescription() || '';
    const loc = ev.getLocation() || '';
    if (q && title.indexOf(q) < 0 && desc.indexOf(q) < 0) continue;
    var guests = [];
    try { guests = ev.getGuestList().map(function (g) { return g.getEmail(); }); } catch (e) {}
    // מייל: מהאורחים, ואם אין — מתוך הכותרת/התיאור
    var email = guests.filter(function (g) { return ALLOWED_EMAILS.indexOf(String(g).toLowerCase()) < 0; })[0] || '';
    if (!email) {
      const m = (title + ' ' + desc).match(/[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/);
      if (m) email = m[0];
    }
    out.push({
      title: title,
      start: Utilities.formatDate(ev.getStartTime(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm"),
      location: loc,
      email: email,
      phase: phaseFromText_(title + ' ' + desc),
      site: siteFromText_(title + ' ' + loc + ' ' + desc)
    });
  }
  return { ok: true, items: out };
}

/** הרץ ידנית — מציג את מבנה 10 האירועים הקרובים ביומן PRIME */
function scanCalendar() {
  const cal = primeCalendar_();
  if (!cal) {
    Logger.log('❌ לא נמצא יומן בשם PRIME. היומנים הזמינים:');
    CalendarApp.getAllCalendars().forEach(function (c) { Logger.log('   • ' + c.getName()); });
    return;
  }
  Logger.log('📅 יומן: ' + cal.getName());
  const now = new Date();
  const evs = cal.getEvents(now, new Date(now.getTime() + 400 * 24 * 3600 * 1000));
  Logger.log('נמצאו ' + evs.length + ' אירועים עתידיים. 10 הראשונים:');
  for (var i = 0; i < Math.min(10, evs.length); i++) {
    const ev = evs[i];
    var guests = [];
    try { guests = ev.getGuestList().map(function (g) { return g.getEmail(); }); } catch (e) {}
    Logger.log('──────────────');
    Logger.log('כותרת : ' + ev.getTitle());
    Logger.log('מתי   : ' + Utilities.formatDate(ev.getStartTime(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'));
    Logger.log('מיקום : ' + (ev.getLocation() || '(ריק)'));
    Logger.log('אורחים: ' + (guests.join(', ') || '(אין)'));
    Logger.log('תיאור : ' + (ev.getDescription() || '(ריק)').slice(0, 120));
  }
}

// ===================== תזמון =====================
function queueScheduled_(data) {
  queueSheet_().appendRow([new Date(), data.sendAt, 'pending', JSON.stringify(data)]);
}

/** מנרמל את מועד השליחה למחרוזת YYYY-MM-DDTHH:mm */
function normSendAt_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm");
  }
  return String(v || '');
}

/** רשימת המיילים הממתינים לשליחה */
function listPending_() {
  const sh = queueSheet_();
  const rows = sh.getDataRange().getValues();
  const out = [];
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][2]) !== 'pending') continue;
    var d = {};
    try { d = JSON.parse(rows[i][3]) || {}; } catch (e) {}
    out.push({
      row: i + 1,
      sendAt: normSendAt_(rows[i][1]),
      to: (d.to || []).join(','),
      name: d.name || '',
      phase: d.phase || '',
      group: d.group || '',
      site: d.site || '',
      gender: d.gender || 'f',
      researchAt: d.researchAt || '',
      subject: d.subject || ''
    });
  }
  out.sort(function (a, b) { return a.sendAt < b.sendAt ? -1 : 1; });
  return out;
}

/** ביטול מייל ממתין */
function cancelPending_(row) {
  const sh = queueSheet_();
  if (!row || row < 2 || row > sh.getLastRow()) return { ok: false, error: 'שורה לא תקינה' };
  if (String(sh.getRange(row, 3).getValue()) !== 'pending') return { ok: false, error: 'המייל כבר נשלח או בוטל' };
  sh.getRange(row, 3).setValue('בוטל ' + new Date().toISOString());
  return { ok: true, cancelled: true };
}

/** שינוי מועד של מייל ממתין */
function reschedulePending_(row, sendAt) {
  const sh = queueSheet_();
  if (!row || row < 2 || row > sh.getLastRow()) return { ok: false, error: 'שורה לא תקינה' };
  if (String(sh.getRange(row, 3).getValue()) !== 'pending') return { ok: false, error: 'המייל כבר נשלח או בוטל' };
  if (!sendAt) return { ok: false, error: 'חסר מועד חדש' };
  sh.getRange(row, 2).setValue(sendAt);
  var d = {};
  try { d = JSON.parse(sh.getRange(row, 4).getValue()) || {}; } catch (e) {}
  d.sendAt = sendAt;
  sh.getRange(row, 4).setValue(JSON.stringify(d));
  return { ok: true, sendAt: sendAt };
}

function checkScheduled() {
  const sh = queueSheet_();
  const rows = sh.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2]) !== 'pending') continue;
    if (new Date(normSendAt_(rows[i][1])) <= now) {
      try {
        sendMail_(JSON.parse(rows[i][3]));
        sh.getRange(i + 1, 3).setValue('sent ' + new Date().toISOString());
      } catch (err) {
        sh.getRange(i + 1, 3).setValue('error: ' + err);
      }
    }
  }
}

function queueSheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('QUEUE_SHEET_ID');
  let ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create('PRIME Mailer Queue');
    props.setProperty('QUEUE_SHEET_ID', ss.getId());
    ss.getSheets()[0].appendRow(['created', 'sendAt', 'status', 'payload']);
  }
  return ss.getSheets()[0];
}

function installTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'checkScheduled') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkScheduled').timeBased().everyMinutes(15).create();
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
