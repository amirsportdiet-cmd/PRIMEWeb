/**
 * מערכת מיילים PRIME — שרת שליחה (Google Apps Script)
 * פרויקט תחת primerct2026@gmail.com
 *
 * התקנה: להריץ פעם אחת את הפונקציה setup()
 * היא יוצרת את תיקיית הקבצים, מתקינה את טריגר התזמון, ומדפיסה את הקישור לתיקייה.
 */

const CODE_VERSION = 'v19-auto-full';
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

    /* דופק חיצוני: מריץ את בדיקת המתוזמנים עכשיו, ומתקין מחדש את הטריגר אם נעלם.
       כך התזמון עובד גם כשהטריגר של גוגל מת בשקט (הסיבה שהמיילים "לא יצאו בפועל").
       מוגן בסוד בלבד — שולח רק מה שכבר תוזמן ואושר, לא מקבל תוכן חדש. */
    if (data.mode === 'tick') {
      var fixed = false;
      try {
        var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'checkScheduled'; });
        if (!has) { installTrigger_(); fixed = true; }
      } catch (e) { /* ריצה בהקשר בלי הרשאת טריגרים — הדופק עצמו עדיין שולח */ }
      var r = checkScheduled();
      return json({ ok: true, version: CODE_VERSION, triggerReinstalled: fixed,
        pending: r.pending, sent: r.sent, errors: r.errors });
    }

    // בדיקת גרסה + אבחון יומן (ללא פרטי נחקרים)
    if (data.mode === 'ping') {
      // מצב הטריגרים ותור ההמתנה — כדי שאפשר יהיה לראות מבחוץ שהתזמון חי
      var trig = [];
      try {
        trig = ScriptApp.getProjectTriggers().map(function (t) {
          return t.getHandlerFunction() + ' (' + t.getEventType() + ')';
        });
      } catch (e) { trig = ['unreadable: ' + e]; }
      var qStat = { pending: 0, overdue: 0 };
      try {
        var qRows = queueSheet_().getDataRange().getValues();
        var nowT = new Date();
        for (var qi = 1; qi < qRows.length; qi++) {
          if (String(qRows[qi][2]) !== 'pending') continue;
          qStat.pending++;
          if (new Date(normSendAt_(qRows[qi][1])) <= nowT) qStat.overdue++;
        }
      } catch (e) { qStat.error = String(e); }
      var cal = {};
      try {
        var c = primeCalendar_();
        if (c) {
          cal = { name: c.getName(),
                  events: c.getEvents(new Date(), new Date(Date.now() + 400 * 24 * 3600 * 1000)).length };
        } else {
          cal = { name: null, available: CalendarApp.getAllCalendars().map(function (x) { return x.getName(); }) };
        }
      } catch (e) { cal = { error: String(e) }; }
      // בדיקת הפונקציה עצמה — ספירות בלבד, ללא פרטי נחקרים
      var probe = {};
      try {
        var p = lookupCalendar_(data.probe || '');
        probe = { ok: p.ok, items: (p.items || []).length, total: p.total, name: p.calendarName, err: p.error };
      } catch (e2) { probe = { thrown: String(e2) }; }
      return json({ ok: true, version: CODE_VERSION, calendar: cal, probe: probe,
        triggers: trig, queue: qStat });
    }

    const v = verifiedEmail_(data.idToken);
    if (!v.email) return json({ ok: false, error: 'זיהוי נכשל — ' + v.err });
    if (ALLOWED_EMAILS.indexOf(v.email) < 0) {
      return json({ ok: false, error: 'החשבון ' + v.email + ' אינו מורשה' });
    }
    data.sentBy = v.email;

    // --- חיפוש ביומן PRIME ---
    if (data.mode === 'calendar')   return json(lookupCalendar_(data.name));

    // --- ניהול מיילים ממתינים ---
    if (data.mode === 'history')    return json({ ok: true, items: historyFor_(data.email) });
    if (data.mode === 'list')       return json({ ok: true, items: listPending_() });
    if (data.mode === 'sentlog')    return json({ ok: true, items: sentLog_(data.limit) });
    if (data.mode === 'cancel')     return json(cancelPending_(data.row));
    if (data.mode === 'reschedule') return json(reschedulePending_(data.row, data.sendAt));

    // --- קריאת מיילי פניות נחקרים מהתיבה ---
    if (data.mode === 'inbox')      return json(fetchAlertMails_(data.days, data.since));

    // --- שליחת דו״ח תוצאות אישי, עם PDF שנבנה בדפדפן ---
    // לפני בדיקת הכפילויות: לדו״ח אין שלב וקבוצה, והוא עשוי להישלח יותר מפעם אחת
    if (data.mode === 'report')     return json(sendReport_(data));

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
    if (data.mode === 'schedule') {
      queueScheduled_(data);
      logSend_(data, 'תוזמן');
      return json({ ok: true, scheduled: true, sendAt: data.sendAt });
    }
    return json({ ok: false, error: 'mode לא מוכר: ' + data.mode });
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

/**
 * שולח דו״ח תוצאות אישי לנחקר/ת. ה-PDF נבנה בדפדפן ומגיע כאן כ-base64,
 * כדי שמה שנשלח יהיה בדיוק מה שהחוקר ראה על המסך.
 */
function sendReport_(data) {
  var to = (data.to || []).filter(function (x) { return x; });
  if (!to.length) return { ok: false, error: 'לא צוינה כתובת מייל' };
  if (!/^[^\s@]+@[^\s@]+\.[a-zA-Z]{2,}$/.test(to[0])) {
    return { ok: false, error: 'כתובת המייל אינה תקינה' };
  }
  if (!data.pdfBase64) return { ok: false, error: 'לא התקבל קובץ הדו״ח' };

  var bytes = Utilities.base64Decode(data.pdfBase64);
  var name  = data.pdfName || 'PRIME.pdf';
  if (bytes.length > 20 * 1024 * 1024) return { ok: false, error: 'קובץ הדו״ח גדול מדי' };
  var pdf = Utilities.newBlob(bytes, 'application/pdf', name);

  GmailApp.sendEmail(to.join(','), data.subject, data.body, {
    name: SENDER_NAME,
    htmlBody: rtlHtml_(data.body),
    attachments: [pdf]
  });

  logSend_({ to: to, name: '(דו״ח תוצאות)', phase: 'דו״ח', group: '', site: '',
             sendAt: '', sentBy: data.sentBy, subject: data.subject }, 'דו״ח נשלח');
  return { ok: true, sent: true, to: to[0], kb: Math.round(bytes.length / 1024) };
}

// ===================== קריאת מיילי פניות =====================
/* מיילי ההתראה מגיעים תמיד מאותו שולח. הם נשלחים במקור לתיבה האישית,
   ומועברים לכאן בכלל אוטומטי, כי הסקריפט הזה רץ תחת primerct2026.
   ההרשאה לקריאה כבר קיימת (https://mail.google.com/) ולכן אין צורך באישור נוסף. */
const ALERT_SENDER = 'send.vpcontact.com';
// תווית שהמסנן ב-primerct2026 מדביק למיילים שהועברו מהתיבה האישית.
// החיפוש מקבל גם תווית וגם שולח, כדי שיעבוד עם המסנן ובלעדיו.
const ALERT_LABEL = 'prime-alerts';

/**
 * מחזיר את מיילי ההתראה מהתקופה האחרונה.
 * הפרסור עצמו נעשה בדפדפן, כדי שנוסח חדש יטופל בלי לפרוס מחדש.
 * since — חותמת ISO, מחזיר רק מה שחדש ממנה.
 */
function fetchAlertMails_(days, since) {
  var d = Math.max(1, Math.min(365, days || 30));
  var q = '(label:' + ALERT_LABEL + ' OR from:' + ALERT_SENDER + ') newer_than:' + d + 'd';
  var threads = GmailApp.search(q, 0, 200);
  var cut = since ? new Date(since).getTime() : 0;
  var items = [];
  for (var i = 0; i < threads.length; i++) {
    var msgs = threads[i].getMessages();
    for (var j = 0; j < msgs.length; j++) {
      var m = msgs[j];
      if (cut && m.getDate().getTime() <= cut) continue;
      items.push({
        id: m.getId(),
        at: m.getDate().toISOString(),
        subject: m.getSubject(),
        body: m.getPlainBody().slice(0, 3000)
      });
    }
  }
  items.sort(function (a, b) { return a.at < b.at ? 1 : -1; });
  return { ok: true, query: q, threads: threads.length, count: items.length, items: items };
}

/** בדיקה ידנית: כמה מיילי התראה יש בתיבה */
function alertsSelfTest() {
  var r = fetchAlertMails_(120, null);
  Logger.log('שאילתה: ' + r.query);
  Logger.log('נמצאו ' + r.count + ' מיילים ב-' + r.threads + ' שרשורים');
  return r.count;
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

/**
 * יומן שליחות לעמוד — אמיר: "שיהיה לוג פשוט שמראה שהמייל שתוזמן יצא".
 * ממזג את יומן השליחות (מיידי + מתוזמן) עם שורות התור שכבר טופלו
 * (sent / skipped / error), החדש ביותר ראשון.
 */
function sentLog_(limit) {
  var max = Math.max(5, Math.min(100, limit || 30));
  var out = [];
  var rows = logSheet_().getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    out.push({
      when: rows[i][0] instanceof Date
        ? Utilities.formatDate(rows[i][0], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : String(rows[i][0]),
      ts: rows[i][0] instanceof Date ? rows[i][0].getTime() : 0,
      to: String(rows[i][1] || ''), name: String(rows[i][2] || ''), phase: String(rows[i][3] || ''),
      kind: String(rows[i][6] || ''), sendAt: normSendAt_(rows[i][7]), by: String(rows[i][8] || '')
    });
  }
  // שורות תור שטופלו — מהן רואים גם דילוגים ותקלות של המתוזמנים
  var qRows = queueSheet_().getDataRange().getValues();
  for (var k = 1; k < qRows.length; k++) {
    var st = String(qRows[k][2] || '');
    if (st === 'pending') continue;
    var d = {};
    try { d = JSON.parse(qRows[k][3]) || {}; } catch (e) {}
    var kind = st.indexOf('sent') === 0 ? 'נשלח (מתוזמן)'
      : st.indexOf('skipped') === 0 ? 'דולג — יום המחקר עבר'
      : st.indexOf('error') === 0 ? 'שגיאה' : st;
    var tsM = st.match(/\d{4}-\d{2}-\d{2}T[\d:.]+Z/);
    var ts = tsM ? new Date(tsM[0]).getTime() : 0;
    out.push({
      when: ts ? Utilities.formatDate(new Date(ts), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : st,
      ts: ts, to: (d.to || []).join(','), name: d.name || '', phase: d.phase || '',
      kind: kind, sendAt: normSendAt_(qRows[k][1]), by: 'מערכת (תזמון)', queue: true
    });
  }
  // 'תוזמן' מהיומן הישן כפול מול שורת התור — היומן מציג את שתיהן, זה בסדר: אחת "נכנס לתור" ואחת "יצא"
  out.sort(function (a, b) { return b.ts - a.ts; });
  return out.slice(0, max);
}

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
  if (/T0/.test(s)) return 'start';
  if (/T3/.test(s)) return 'middle';
  if (/T6/.test(s)) return 'end';
  return '';
}

/** שם הנחקר: מתוך "Invitee:" בתיאור, או מהכותרת שלפני ":" */
function nameFromEvent_(title, desc) {
  const m = String(desc || '').match(/Invitee\s*:\s*([^\r\n]+)/i);
  if (m && m[1].trim()) return m[1].trim();
  const t = String(title || '');
  const i = t.indexOf(':');
  return (i > 0 ? t.slice(0, i) : t).trim();
}

/** מייל הנחקר: מתוך "Invitee Email:" בתיאור */
function emailFromEvent_(desc) {
  const m = String(desc || '').match(/Invitee\s*Email\s*:\s*([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})/i);
  return m ? m[1] : '';
}

/** קבוצת המחקר של הנחקר/ת.
    ביומן יש סוג פגישה אחד, "מחקר PRIME", ובטופס ההזמנה שלו יש שאלת חובה
    בשם "קבוצת מחקר" שהצוות עונה עליה בזמן קביעת הפגישה. קלנדלי כותב את
    התשובה בתיאור האירוע, ומכאן היא נקראת. "טרם שובץ" (T0) מחזיר ריק בכוונה. */
function groupFromEvent_(desc) {
  const m = String(desc || '').match(/קבוצת\s*מחקר\s*[:\uFF1A]\s*([^\r\n]+)/);
  const ans = m ? m[1].trim() : '';
  if (ans.indexOf('התערבות') >= 0) return 'intervention';
  if (ans.indexOf('ביקורת') >= 0)  return 'control';
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
  if (!cal) {
    const names = CalendarApp.getAllCalendars().map(function (c) { return c.getName(); });
    return { ok: false, error: 'לא נמצא יומן בשם ' + CALENDAR_NAME + '. יומנים זמינים: ' + names.join(' · ') };
  }
  const now = new Date();
  const end = new Date(now.getTime() + 400 * 24 * 3600 * 1000);
  const q = String(name || '').trim();
  const words = q.split(/\s+/).filter(function (w) { return w.length >= 2; });
  const evs = cal.getEvents(now, end);
  const out = [];
  const all = [];   // כל האירועים — לגיבוי אם אין התאמה
  for (var i = 0; i < evs.length; i++) {
    const ev = evs[i];
    const title = ev.getTitle() || '';
    const desc = ev.getDescription() || '';
    const loc = ev.getLocation() || '';
    const person = nameFromEvent_(title, desc);
    const hay = (person + ' ' + title + ' ' + desc);
    // התאמה אם אחת ממילות החיפוש מופיעה
    const hit = !q || words.some(function (w) { return hay.indexOf(w) >= 0; });
    if (all.length < 15) all.push({ title: title, name: person, loc: loc });
    if (!hit) continue;
    if (out.length >= 60) continue;
    var email = emailFromEvent_(desc);
    if (!email) {
      try {
        email = ev.getGuestList().map(function (g) { return g.getEmail(); })
          .filter(function (g) { return ALLOWED_EMAILS.indexOf(String(g).toLowerCase()) < 0; })[0] || '';
      } catch (e) {}
    }
    out.push({
      title: title,
      name: person,
      start: Utilities.formatDate(ev.getStartTime(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm"),
      location: loc,
      email: email,
      phase: phaseFromText_(loc + ' ' + title + ' ' + desc),
      site: siteFromText_(loc + ' ' + title + ' ' + desc),
      group: groupFromEvent_(desc)
    });
  }
  return {
    ok: true, items: out,
    calendarName: cal.getName(),
    total: evs.length,
    sample: out.length ? [] : all.slice(0, 8)   // אם אין התאמה — מציגים מה כן קיים
  };
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
  const out = { pending: 0, sent: 0, errors: 0, skippedStale: 0 };
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2]) !== 'pending') continue;
    const due = new Date(normSendAt_(rows[i][1]));
    if (due > now) { out.pending++; continue; }
    var d = null;
    try { d = JSON.parse(rows[i][3]); } catch (e) {}
    /* אמיר (20.8.2026): מייל שנתקע ויום המחקר שלו כבר עבר — לא שולחים. תזכורת
       לפגישה שהייתה היא גרועה משתיקה. נחשב עבר-זמנו אם מועד המחקר מאחורינו,
       או (כשאין מועד מחקר בנתונים) אם מועד השליחה חלף לפני יותר מ-6 ימים —
       התזכורת נשלחת שבוע לפני, כך שאחרי 6 ימים הפגישה כבר מאחורינו ממילא. */
    var research = d && d.researchAt ? new Date(d.researchAt) : null;
    var stale = (research && !isNaN(research) && research < now)
      || (!(research && !isNaN(research)) && (now - due) > 6 * 24 * 3600 * 1000);
    if (stale) {
      sh.getRange(i + 1, 3).setValue('skipped ' + new Date().toISOString() + ' (יום המחקר עבר)');
      out.skippedStale++;
      continue;
    }
    /* הפגישה הוזזה מאז התזמון הידני? לא שולחים תאריך שגוי — מבטלים, והאוטומט
       שולח את המעודכן (הוא סורק את היומן כל רבע שעה) */
    if (d && d.name && d.researchAt && !d.auto) {
      var cur = eventStartForName_(d.name);
      if (cur && Math.abs(cur.getTime() - new Date(d.researchAt).getTime()) > 36 * 3600 * 1000) {
        sh.getRange(i + 1, 3).setValue('בוטל אוטומטית ' + new Date().toISOString() + ' (הפגישה הוזזה - תישלח תזכורת מעודכנת)');
        continue;
      }
    }
    /* שורות היסטוריות של עודד חיימוב נכשלו 17 פעמים ב-TypeError כי המטען נשמר בלי
       'to' — מעכשיו שורה פגומה מסומנת בבירור במקום להתפוצץ שוב ושוב. */
    if (!d || !Array.isArray(d.to) || !d.to.length) {
      sh.getRange(i + 1, 3).setValue('error: המטען חסר נמען (to) — לתזמן מחדש מהעמוד');
      out.errors++;
      continue;
    }
    try {
      sendMail_(d || JSON.parse(rows[i][3]));
      sh.getRange(i + 1, 3).setValue('sent ' + new Date().toISOString());
      // נרשם גם ביומן השליחות — כדי שיהיה מקום אחד שרואים בו שהמייל באמת יצא
      try { logSend_(Object.assign({}, d, { sentBy: 'מערכת (תזמון)' }), 'נשלח (מתוזמן)'); } catch (e2) {}
      out.sent++;
    } catch (err) {
      sh.getRange(i + 1, 3).setValue('error: ' + err);
      out.errors++;
    }
  }
  try { autoReminders_(); } catch (e) { /* אוטומט התזכורות לא מפיל את שליחת התור */ }
  reportQueue_(sh.getDataRange().getValues());
  return out;
}

/** מועד האירוע העתידי הקרוב של נחקר/ת ביומן PRIME (לפי שם), או null. */
function eventStartForName_(name) {
  try {
    var cal = primeCalendar_(); if (!cal) return null;
    var now = new Date();
    var evs = cal.getEvents(now, new Date(now.getTime() + 120 * 24 * 3600 * 1000));
    var norm = function (x) { return String(x || '').replace(/["'׳״׳״-]/g, ' ').replace(/\s+/g, ' ').trim(); };
    var q = norm(name).split(' ').filter(String);
    if (!q.length) return null;
    for (var i = 0; i < evs.length; i++) {
      var person = norm(nameFromEvent_(evs[i].getTitle() || '', evs[i].getDescription() || ''));
      var w = person.split(' ').filter(String);
      var pair = q.length <= w.length ? [q, w] : [w, q];
      if (pair[0].length && pair[0].every(function (x) { return pair[1].indexOf(x) >= 0; })) return evs[i].getStartTime();
    }
  } catch (e) {}
  return null;
}

/* ===== תזכורת אוטומטית — אמיר (20.8.2026): "שישלח אוטומטית בלי אישור; מי
   שבתוך השבוע הקרוב מקבל גם אם זה כבר לא בדיוק שבוע לפני; ההודעה זהה למייל
   המקורי — כל ההסברים והקבצים, לא נוסח הוואטסאפ".
   כל ריצת טריגר: אירועי יומן PRIME ב-8 הימים הקרובים. נחקר שיום המחקר שלו
   בעוד ≤7 ימים ואין לו תזכורת בתור — מקבל את מייל ההכנה המלא (תבנית לפי
   שלב+קבוצה+מוקד, אותם קבצים מצורפים). מפתח הדדופ כולל את מועד האירוע —
   אם הפגישה זזה ביומן, נשלחת תזכורת מעודכנת אוטומטית. */
var WD_HE = ['ראשון', 'שני', 'שלישי', 'רביעי', 'חמישי', 'שישי', 'שבת'];
var SITES_INFO = {
  assuta: { location: 'אסותא רמת החייל, רחוב הברזל 20, בניין אשפוז, קומה 1, חדר 140',
    parking: '** ביום המחקר נסדיר עבורך חניה - יש לוודא שחנית בחניון בית החולים, ברחוב הברזל 20, רמת החייל.' },
  ichilov: { location: 'המרכז הרפואי תל אביב ע״ש סוראסקי (איכילוב), דרך וייצמן 6, מגדל האשפוז ע״ש אריסון, קומה 13 - המכון האנדוקריני',
    parking: '** ביום המחקר נסדיר עבורך חניה - יש לוודא שחנית בחניון בית החולים, דרך וייצמן 6, תל אביב.' }
};
var FILES_BY_KEY = {
  'assuta:start': ['T0 - אסותא.pdf', 'איסוף שתן 24 שעות - הנחיות ליום מחקר.pdf', 'הסכמה מדעת אסותא.pdf'],
  'assuta:middle:control': ['T3 - ביקורת אסותא.pdf'],
  'assuta:middle:intervention': ['T3 - התערבות אסותא.pdf'],
  'assuta:end:control': ['T6 - ביקורת אסותא.pdf', 'איסוף שתן 24 שעות - הנחיות ליום מחקר.pdf'],
  'assuta:end:intervention': ['T6 - התערבות אסותא.pdf', 'איסוף שתן 24 שעות - הנחיות ליום מחקר.pdf'],
  'ichilov:start': ['T0 - איכילוב.pdf', 'איסוף שתן 24 שעות - הנחיות ליום מחקר.pdf', 'הסכמה מדעת איכילוב.pdf'],
  'ichilov:middle:control': ['T3 - ביקורת איכילוב.pdf'],
  'ichilov:middle:intervention': ['T3 - התערבות איכילוב.pdf'],
  'ichilov:end:control': ['T6 - ביקורת איכילוב.pdf', 'איסוף שתן 24 שעות - הנחיות ליום מחקר.pdf'],
  'ichilov:end:intervention': ['T6 - התערבות איכילוב.pdf', 'איסוף שתן 24 שעות - הנחיות ליום מחקר.pdf']
};
function p2_(n) { return ('0' + n).slice(-2); }
function ddmm_(d) { return p2_(d.getDate()) + '/' + p2_(d.getMonth() + 1); }
function ddmmyyyy_(d) { return ddmm_(d) + '/' + d.getFullYear(); }
function addDays_(d, n) { var x = new Date(d); x.setDate(x.getDate() + n); return x; }
function diaryDays_(research) {
  var out = []; var d = addDays_(research, -2);
  while (out.length < 3) { var wd = d.getDay(); if (wd >= 0 && wd <= 4) out.push(new Date(d)); d = addDays_(d, -1); }
  return out.reverse();
}
function genderApply_(text, g) {
  return String(text || '').replace(/\[([^\[\]|]*)\|([^\[\]|]*)\]/g, function (m, a, b) { return g === 'm' ? a : b; });
}

/* אותן תבניות בדיוק כמו בעמוד המיילר (prime-mailer.html) — שינוי נוסח משנים בשניהם */
function fullReminder_(name, phase, group, site, gender, research) {
  var first = String(name || '').trim().split(/\s+/)[0] || '';
  var fast = new Date(research.getTime() - 12 * 3600 * 1000);
  var urine = addDays_(research, -2), urineNext = addDays_(research, -1);
  var diary = diaryDays_(research);
  var si = SITES_INFO[site] || SITES_INFO.assuta;
  var map = {
    'שם': first, 'תאריך_מחקר': ddmmyyyy_(research), 'יום_מחקר': WD_HE[research.getDay()],
    'שעת_מחקר': p2_(research.getHours()) + ':' + p2_(research.getMinutes()),
    'יום_צום': WD_HE[fast.getDay()], 'תאריך_צום': ddmmyyyy_(fast),
    'שעת_צום': p2_(fast.getHours()) + ':' + p2_(fast.getMinutes()),
    'איסוף_שתן': ddmm_(urine) + ' (יום ' + WD_HE[urine.getDay()] + ' עד ' + WD_HE[urineNext.getDay()] + ' בבוקר)',
    'יומן_אכילה': diary.map(function (x) { return ddmm_(x) + ' (יום ' + WD_HE[x.getDay()] + ')'; }).join(' + '),
    'תאריך_קצר': research.getDate() + '.' + (research.getMonth() + 1),
    'מיקום': si.location, 'חניה': si.parking,
    'מספר_יום': phase === 'start' ? '1' : phase === 'middle' ? '2' : '3',
    'משך': phase === 'middle' ? 'כ-3' : '3-4'
  };
  var SIGNATURE = '\n--\nתודה רבה על השתתפותך ותרומתך למחקר,\nצוות מחקר PRIME';
  var RULES_COMMON = 'מצורף דף הנחיות מפורט להכנה ליום המחקר.\n' +
    'יום המחקר מתקיים ב- {{מיקום}}.\n' +
    'משך כל הבדיקות צפוי להיות {{משך}} שעות, ולכן מומלץ להביא עמך ארוחה קלה ושתייה לאחר סיום הבדיקות הדורשות צום.\n' +
    'לקראת יום המחקר, נבקש להקפיד על ההנחיות הבאות:\n\n' +
    '1. להגיע בצום של 12 שעות. בבוקר יום המחקר ניתן לשתות מים בלבד (צום {{יום_צום}} מ-{{שעת_צום}}).\n' +
    '2. לשמור על תזונה שגרתית ככל הניתן ביממה שלפני יום המחקר.\n' +
    '3. להימנע מפעילות גופנית במשך 24 שעות לפני ההגעה.\n' +
    '4. להימנע מצריכת אלכוהול במשך 24 שעות לפני ההגעה.\n' +
    '5. להימנע מצריכת קפאין במשך 12 שעות לפני ההגעה.\n' +
    '6. להימנע ממריחת קרמים או שמנים על הידיים והרגליים ביום המחקר.\n' +
    '7. להגיע בביגוד נוח ונעליים סגורות, עדיפות לנעלי ספורט.\n' +
    '8. להביא משקפי קריאה, אם יש צורך.\n' +
    '9. להביא רשימת תרופות ותוספי תזונה עדכנית, או צילום שלהם.\n' +
    '10. להביא יומן אכילה של 3 ימים, בהתאם להנחיות המצורפות.';
  var DIARY_RULES = 'יש לרשום 3 ימי אכילה מלאים, לא כולל שישי/שבת ולא כולל היום שלפני יום המחקר, משום שבאותו ערב מתחיל הצום.\n' +
    'ניתן לרשום את היומן בטלפון הנייד, בדף או במחברת - לפי הנוחות שלך.\n' +
    'חשוב לציין ככל האפשר את שעת האכילה, סוג המזון, הכמות, אופן ההכנה, כמה שיותר פרטים.';
  var VISIT_HEAD = '{{שם}} [היקר|היקרה],\nלקראת הגעתך ליום מחקר {{מספר_יום}} במסגרת מחקר PRIME,\nהנך [מוזמן|מוזמנת] ליום המחקר בתאריך:\nיום {{יום_מחקר}}, {{תאריך_מחקר}}\nבשעה {{שעת_מחקר}}\n\n';
  var MIDDLE_TAIL = '\nלגבי יומן האכילה:\n' + DIARY_RULES +
    '\n\n* [מוזמן|מוזמנת] לשלוח למייל 3 ימי חול, בשבוע שלפני ההגעה לאסותא:\n{{יומן_אכילה}}\n\n{{חניה}}' + SIGNATURE;
  var FULL_VISIT_TAIL = '11. להביא מיכל/י חלבון ריק.\n' +
    '12. בדיקות דם עדכניות (משלושת החודשים האחרונים).\n' +
    '13. לקרוא היטב הנחיות לאיסוף שתן (איסוף 24 שעות) ולהגיע עם המיכל!\n' +
    'את האיסוף יש לבצע בדיוק לפי ההנחיות המצורפות.\n' +
    '[שים|שימי] לב: בערכה מצורף ברקוד/QR לסרטון הדרכה קצר.\n\n' +
    'הנחיות ליומן האכילה:\n' + DIARY_RULES +
    '\n\n* תאריכי יומן אכילה: {{יומן_אכילה}}\n* תאריך איסוף שתן (24 שעות): {{איסוף_שתן}}\n\n' +
    '* [שים|שימי] לב - ביום {{יום_צום}} ({{תאריך_צום}}), [אתה מתחיל|את מתחילה] צום בשעות הערב (לפחות 12 שעות), כלומר החל מהשעה {{שעת_צום}} - ניתן לשתות מים בלבד.\n' +
    '** נשמח בבקשה לשלוח אלינו - בדיקות דם (משלושת החודשים האחרונים).\n' +
    'נתראה ביום {{יום_מחקר}} {{תאריך_מחקר}} בשעה {{שעת_מחקר}}\n' +
    '*נא לקרוא טוב את כל ההנחיות המצורפות :)\n{{חניה}}' + SIGNATURE;
  var START_BODY = 'היי {{שם}},\n\n' +
    'שמחים מאוד שהצטרפת אלינו למחקר PRIME - תודה רבה על שיתוף הפעולה.\n' +
    'כחלק מההכנה ליום המחקר הראשון, חשוב לנו [שתקרא|שתקראי] בעיון את דף ההנחיות המצורף [ותפעל|ותפעלי] לפיו.\n' +
    'בדף [תמצא|תמצאי] הנחיות לגבי:\n\n' +
    '1) הנחיות הגעה ליום המחקר\nתאריך: {{תאריך_מחקר}} יום {{יום_מחקר}}\nבשעה: {{שעת_מחקר}}\nמיקום: {{מיקום}}\n*מומלץ להגיע 10-15 דקות לפני.\n**ניתן לחנות בבניין - יש הסדר חנייה.\n\n' +
    '2) טופס הסכמה מדעת - לקריאה מראש (חתימה תתבצע ביום המחקר)\n\n' +
    '3) הנחיות לאיסוף שתן (איסוף 24 שעות)\nאת האיסוף יש לבצע בדיוק לפי ההנחיות המצורפות.\n[שים|שימי] לב: בערכה מצורף ברקוד/QR לסרטון הדרכה קצר.\n\n' +
    '4) הנחיות ליומן אכילה\nההנחיות ליומן אכילה מופיעות בתוך דף ההנחיות הכללי המצורף.\n\n' +
    '* תאריכי יומן אכילה: {{יומן_אכילה}}\n* תאריך איסוף שתן (24 שעות): {{איסוף_שתן}}\n\n' +
    '* [שים|שימי] לב - ביום {{יום_צום}} ({{תאריך_צום}}), [אתה אמור|את אמורה] להתחיל צום בשעות הערב (לפחות 12 שעות), כלומר החל מהשעה {{שעת_צום}} - ניתן לשתות מים בלבד.\n' +
    '** אשמח בבקשה לשלוח אליי - בדיקות דם (3 חודשים האחרונים) + אישור פעילות גופנית (לבקש מרופא/ת המשפחה) + צילומים של תרופות/ויטמינים [שנוטל|שנוטלת] (כל אלו מופיעים בדף ההנחיות).\n\n' +
    'נתראה ביום {{יום_מחקר}} {{תאריך_מחקר}} בשעה {{שעת_מחקר}} (להגיע בבקשה עם תעודת זהות/רישיון)\n' +
    '*נא לקרוא טוב טוב את כל ההנחיות המצורפות :)';
  var body, subject;
  if (phase === 'start') { subject = 'הכנה ליום מחקר 1 - מחקר PRIME'; body = START_BODY; }
  else {
    subject = 'הכנה ליום מחקר {{מספר_יום}} - מחקר PRIME';
    if (phase === 'middle') body = VISIT_HEAD + RULES_COMMON + (group === 'intervention' ? '\n11. להביא מיכל חלבון ריק (במידה ויש)\n' : '\n') + MIDDLE_TAIL;
    else body = VISIT_HEAD + RULES_COMMON + '\n' + FULL_VISIT_TAIL;
  }
  var sub = function (t) { return genderApply_(String(t).replace(/\{\{\s*([^}]+?)\s*\}\}/g, function (m, k) { return (k in map) ? map[k] : m; }), gender); };
  var fkey = site + ':' + (phase === 'start' ? 'start' : phase + ':' + group);
  return { subject: sub(subject), body: sub(body), attachments: FILES_BY_KEY[fkey] || [] };
}

/* מין מ-REDCap (cand_sex: 1=זכר, 2=נקבה; ייצוא 17.8.2026) — לניסוח נכון של המייל.
   אותה התאמה גמישה כמו הקבוצות. מתעדכן עם כל ייצוא חדש. */
var SEX_FALLBACK = {
  'אביב ברשף': 'm',
  'אביבה פנחסוב': 'f',
  'אביבה רובין': 'f',
  'אברהם איגל': 'm',
  'אורה שורץ': 'f',
  'אורי קטרי': 'm',
  'אורן גונן': 'm',
  'אילנה נתוביץ': 'f',
  'אירית שרון': 'f',
  'איתי שגיא': 'm',
  'אלה נגרו יוסף': 'f',
  'אלי מלכה': 'm',
  'אלי שוורצר': 'm',
  'אלישבע גרינשטיין': 'f',
  'אלמוג קובריגרו': 'm',
  'אמנון אגסי': 'm',
  'אפרת אגסי פלדמן': 'f',
  'אריאלה קלפנר': 'f',
  'אשרת הרבסט': 'f',
  'אתי כובש': 'f',
  'ברטה ניסימוב': 'f',
  'בתיה רחום': 'f',
  "ג'ני פקגוז": 'f',
  'גיל בן צבי': 'm',
  'גילה מעתוק': 'f',
  'גלית רז': 'f',
  'גלית ריכטר': 'f',
  'דביר שרעבי': 'm',
  'דוד בכר': 'm',
  'דן בוקאי': 'm',
  'דניאלה נוסבאום': 'f',
  'ורדית עופר': 'f',
  'זלמן שטרנליכט': 'f',
  'חדווה מלמד פגדאו': 'f',
  'חיים זאב שהם': 'm',
  'חנה הכרמי': 'f',
  'טטיאנה ברדר': 'f',
  'טטיאנה חבין': 'f',
  'טלי כהנא': 'f',
  'יהודית מזרחי': 'f',
  'יוסף בוקסבוים': 'm',
  'יוספה בלקין': 'f',
  'יעל ברגשטיין': 'f',
  'יעל דורון': 'f',
  'יפית כרפסי': 'f',
  'יצחק אדלר': 'm',
  'יצחק רוגק': 'm',
  'ישראלה רבינא יפעת': 'f',
  'כרמלה גוטל': 'f',
  'לי- רון עופר': 'f',
  'ליטל ניר': 'f',
  'מאיה מרגולין': 'f',
  'מאיה סלומונס': 'f',
  'מיכל לב ארי': 'f',
  'מיכל פרידמן רם': 'f',
  'מירי רפאלי': 'f',
  'נחמה שפירא': 'f',
  'סוניה ביטרמן גפנר': 'f',
  'סימון בר ציון': 'f',
  'סמדר דוד': 'f',
  'ספי פלג': 'f',
  'עדנה ברזוזה': 'f',
  'עודד חיימוב': 'm',
  'עמית גיא': 'f',
  'עמליה דגן': 'f',
  'ענת בצרי': 'f',
  'ערן מונרוב': 'm',
  'פליקס שופמן': 'm',
  'פנינה שטראוס': 'f',
  'קרן קשת גופמן': 'f',
  'רביב שוורץ': 'm',
  'רואי ליטרט': 'm',
  'רונית אלגרבלי': 'f',
  'רונית מרסה': 'f',
  'רותי שליו': 'f',
  'רחל בן יצחק': 'f',
  'רם שטיינר': 'm',
  'שחר ברטל': 'm',
  'שירלי אלף': 'f',
  'שלום קמיל': 'm',
  'שרון בונקר': 'm',
  'תחיה אבידן': 'f',
  'תמי מלכה': 'f'
};
function sexFromRoster_(name) {
  var norm = function (x) { return String(x || '').replace(/["'׳״׳״-]/g, ' ').replace(/\s+/g, ' ').trim(); };
  var q = norm(name).split(' ').filter(String);
  if (!q.length) return '';
  var simLast = function (a, b) {
    if (a === b || a.indexOf(b) >= 0 || b.indexOf(a) >= 0) return true;
    var pre = 0; while (pre < Math.min(a.length, b.length) && a[pre] === b[pre]) pre++;
    var suf = 0; while (suf < Math.min(a.length, b.length) - pre && a[a.length - 1 - suf] === b[b.length - 1 - suf]) suf++;
    return pre >= 2 && suf >= 4;
  };
  for (var key in SEX_FALLBACK) {
    var w = norm(key).split(' ').filter(String);
    if (w[0] !== q[0]) continue;
    if (q.length === 1 || w.length === 1) continue;
    var contained = q.slice(1).every(function (x) { return w.indexOf(x) >= 0; });
    if (contained || simLast(q[q.length - 1], w[w.length - 1])) return SEX_FALLBACK[key];
  }
  return '';
}

/* קבוצות מחקר מ-REDCap (ייצוא 17.8.2026) — גיבוי כשהתשובה חסרה באירוע היומן.
   1=ביקורת, 2=התערבות. שמות היומן לעיתים בכתיב שונה (גרשטיין/גרינשטיין) —
   ההתאמה: שם פרטי זהה + שם משפחה זהה/מוכל/דומה. מתעדכן עם כל ייצוא חדש. */
var GROUP_FALLBACK = {
  "אביב ברשף': 'intervention', 'אורן גונן': 'control', 'אלה נגרו יוסף": 'intervention',
  "אלי שוורצר': 'control', 'אלישבע גרינשטיין': 'control', 'אלמוג קובריגרו": 'intervention',
  "אמנון אגסי': 'intervention', 'אתי כובש': 'control', 'ברטה ניסימוב": 'control',
  "גיל בן צבי': 'intervention', 'גלית ריכטר': 'control', 'דביר שרעבי": 'control',
  "דוד בכר': 'intervention', 'דן בוקאי': 'control', 'ורדית עופר": 'intervention',
  "חיים זאב שהם': 'control', 'חנה הכרמי': 'control', 'טטיאנה חבין": 'control',
  "טלי כהנא': 'control', 'יעל ברגשטיין': 'control', 'יפית כרפסי": 'control',
  "כרמלה גוטל': 'control', 'ליטל ניר': 'intervention', 'מיכל לב ארי": 'intervention',
  "מיכל פרידמן רם': 'intervention', 'מירי רפאלי': 'intervention', 'נחמה שפירא": 'intervention',
  "סוניה ביטרמן גפנר': 'intervention', 'סמדר דוד': 'control', 'עדנה ברזוזה": 'intervention',
  "עודד חיימוב': 'control', 'עמית גיא': 'control', 'עמליה דגן": 'intervention',
  "פליקס שופמן': 'intervention', 'רביב שוורץ': 'intervention', 'רואי ליטרט": 'intervention',
  "רחל בן יצחק': 'control', 'רינת שלום': 'control', 'שחר ברטל': 'control', 'שירלי אלף": 'intervention'
};
function groupFromRoster_(name) {
  var norm = function (x) { return String(x || '').replace(/["'׳״\-]/g, ' ').replace(/\s+/g, ' ').trim(); };
  var q = norm(name).split(' ').filter(String);
  if (!q.length) return '';
  var simLast = function (a, b) {
    if (a === b || a.indexOf(b) >= 0 || b.indexOf(a) >= 0) return true;
    var pre = 0; while (pre < Math.min(a.length, b.length) && a[pre] === b[pre]) pre++;
    var suf = 0; while (suf < Math.min(a.length, b.length) - pre && a[a.length - 1 - suf] === b[b.length - 1 - suf]) suf++;
    return pre >= 2 && suf >= 4;
  };
  for (var key in GROUP_FALLBACK) {
    var w = norm(key).split(' ').filter(String);
    if (w[0] !== q[0]) continue;
    if (q.length === 1 || w.length === 1) continue;
    var ql = q[q.length - 1], wl = w[w.length - 1];
    // כל מילות היומן מוכלות ברשומה, או שם משפחה דומה
    var contained = q.slice(1).every(function (x) { return w.indexOf(x) >= 0; });
    if (contained || simLast(ql, wl)) return GROUP_FALLBACK[key];
  }
  return '';
}

/** מין מהתיאור של האירוע (אם הטופס שואל), ברירת מחדל נקבה — כמו בעמוד. */
function genderFromEvent_(desc) {
  var m = String(desc || '').match(/(?:מין|מגדר)\s*[:：]\s*([^\r\n]+)/);
  var ans = m ? m[1] : '';
  if (ans.indexOf('זכר') >= 0 || /male/i.test(ans)) return 'm';
  return 'f';
}

function autoReminders_() {
  const cal = primeCalendar_(); if (!cal) return;
  const now = new Date();
  const evs = cal.getEvents(now, new Date(now.getTime() + 8 * 24 * 3600 * 1000));
  if (!evs.length) return;
  const sh = queueSheet_();
  const rows = sh.getDataRange().getValues();
  const normN = function (s) { return String(s || '').replace(/["'׳״\-]/g, ' ').replace(/\s+/g, ' ').trim(); };
  const have = {};
  const manualCover = [];   // {key, researchMs} — תזמון ידני מכסה רק אם מועדו תואם ליומן
  const errKeys = {};
  for (var i = 1; i < rows.length; i++) {
    var d = {}; try { d = JSON.parse(rows[i][3]) || {}; } catch (e) {}
    var st = String(rows[i][2] || '');
    /* רק שליחה מוצלחת חוסמת לתמיד; שורת שגיאה רק מונעת כפילות של אותה שגיאה,
       וברגע שהבעיה נפתרת (קבוצה הושלמה, מייל נוסף) — המייל יוצא בריצה הבאה */
    if (d.autoKey) { if (st.indexOf('sent') === 0) have[d.autoKey] = true; else if (st.indexOf('error') === 0) errKeys[d.autoKey] = true; }
    if ((st === 'pending' || st.indexOf('sent') === 0) && d.name && !d.auto && !d.autoKey) {
      var rMs = d.researchAt ? new Date(d.researchAt).getTime() : 0;
      manualCover.push({ key: 'np_' + normN(d.name) + '|' + (d.phase || ''), researchMs: rMs });
    }
  }
  for (var j = 0; j < evs.length; j++) {
    var ev = evs[j];
    var start = ev.getStartTime();
    if (start <= now) continue;
    var sendAt = new Date(start.getTime() - 7 * 24 * 3600 * 1000);
    /* אירוע שהוזז אחרי שכבר יצאה לו תזכורת: לא מחכים לחלון החדש — שולחים מיד
       תיקון עם התאריך הנכון, אחרת הנחקר נשאר עם תאריך שגוי ביד */
    var evPrefix = 'auto_' + ev.getId() + '_';
    var sentOtherDate = false;
    for (var hk in have) { if (hk.indexOf(evPrefix) === 0) { sentOtherDate = true; break; } }
    if (sendAt > now && !sentOtherDate) continue;
    var desc = ev.getDescription() || '', title = ev.getTitle() || '', loc = ev.getLocation() || '';
    var name = nameFromEvent_(title, desc);
    if (!name || name.length < 2) continue;
    var phase = phaseFromText_(loc + ' ' + title + ' ' + desc);
    /* המפתח כולל את מועד האירוע: הפגישה זזה ביומן ⇒ מפתח חדש ⇒ תזכורת מעודכנת */
    var key = 'auto_' + ev.getId() + '_' + Utilities.formatDate(start, Session.getScriptTimeZone(), 'yyyyMMddHHmm');
    if (have[key]) continue;
    /* תזמון ידני קיים מכסה רק כשהמועד שנשמר בו קרוב (עד 36 שעות) למועד ביומן —
       אם הפגישה הוזזה, נשלחת תזכורת מעודכנת */
    var covKey = 'np_' + normN(name) + '|' + phase;
    var covered = manualCover.some(function (c) {
      return c.key === covKey && (!c.researchMs || Math.abs(c.researchMs - start.getTime()) < 36 * 3600 * 1000);
    });
    if (covered) continue;
    var email = emailFromEvent_(desc);
    if (!email) {
      try {
        email = ev.getGuestList().map(function (g) { return g.getEmail(); })
          .filter(function (g) { return ALLOWED_EMAILS.indexOf(String(g).toLowerCase()) < 0; })[0] || '';
      } catch (e) {}
    }
    var site = siteFromText_(loc + ' ' + title + ' ' + desc);
    var group = groupFromEvent_(desc) || groupFromRoster_(name);
    var payload = { to: email ? [email] : [], name: name, phase: phase, site: site, group: group,
      sendAt: normSendAt_(sendAt),
      researchAt: Utilities.formatDate(start, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm"),
      autoKey: key, auto: true };
    var problem = !email ? 'אין כתובת מייל באירוע היומן'
      : !phase ? 'לא זוהה שלב (T0/T3/T6) באירוע'
      : (phase !== 'start' && !group) ? 'אין קבוצת מחקר באירוע היומן'
      : '';
    if (problem) {
      if (!errKeys[key]) {
        payload.subject = 'תזכורת ' + name;
        sh.appendRow([new Date(), payload.sendAt, 'error: ' + problem + ' — לשלוח ידנית מהעמוד', JSON.stringify(payload)]);
      }
      continue;
    }
    var mail = fullReminder_(name, phase, group, site, sexFromRoster_(name) || genderFromEvent_(desc), start);
    payload.subject = mail.subject;
    payload.body = mail.body;
    payload.attachments = mail.attachments;
    try {
      sendMail_(payload);
      sh.appendRow([new Date(), payload.sendAt, 'sent ' + new Date().toISOString() + ' (auto)',
        JSON.stringify({ autoKey: key, name: name, phase: phase, site: site, group: group, to: payload.to, researchAt: payload.researchAt, sendAt: payload.sendAt, subject: payload.subject })]);
      try { logSend_(Object.assign({}, payload, { sentBy: 'מערכת (אוטומטי)' }), 'נשלח (תזכורת אוטומטית)'); } catch (e2) {}
    } catch (err) {
      sh.appendRow([new Date(), payload.sendAt, 'error: ' + err, JSON.stringify({ autoKey: key, name: name, phase: phase, to: payload.to })]);
    }
    have[key] = true;
  }
  /* טסט מלא על אמיר (בקשתו 20.8): כאילו יום המחקר שלו חמישי 27.8 07:30, T0 —
     שני מיילים במקביל: המיידי והתזכורת של שבוע לפני (זהים בתוכן, זו הנקודה),
     שניהם עם כל הקבצים. חד-פעמי. */
  try {
    var props = PropertiesService.getScriptProperties();
    if (!props.getProperty('TEST_MAIL_V20')) {
      props.setProperty('TEST_MAIL_V20', String(Date.now()));
      var tStart = new Date(2026, 7, 27, 7, 30, 0);
      var tm = fullReminder_('אמיר', 'start', '', 'assuta', 'm', tStart);
      GmailApp.sendEmail('amirsportdiet@gmail.com', '[בדיקה 1/2 - המייל המיידי] ' + tm.subject, tm.body, {
        name: SENDER_NAME, htmlBody: rtlHtml_(tm.body), attachments: attachmentsFor_(tm.attachments) });
      GmailApp.sendEmail('amirsportdiet@gmail.com', '[בדיקה 2/2 - התזכורת שנשלחת שבוע לפני] ' + tm.subject, tm.body, {
        name: SENDER_NAME, htmlBody: rtlHtml_(tm.body), attachments: attachmentsFor_(tm.attachments) });
    }
  } catch (eT) { /* בדיקה בלבד */ }}

/* תמונת-מצב מצומצמת של התור אל הדשבורד (שם, מועד, סטטוס — בלי כתובות ותוכן),
   כדי שהבוט יוכל לענות "מה ממתין ומה יצא". מאומת בסוד של המיילר עצמו. */
const DASH_SNAP_URL = 'https://us-central1-wellnessprojectar.cloudfunctions.net/waTest?action=primeQueueSnap';
function reportQueue_(rows) {
  try {
    const items = [];
    for (var i = Math.max(1, rows.length - 60); i < rows.length; i++) {
      var d = {};
      try { d = JSON.parse(rows[i][3]) || {}; } catch (e) {}
      items.push({ row: i + 1, name: d.name || '', phase: d.phase || '',
        sendAt: normSendAt_(rows[i][1]), status: String(rows[i][2] || '').slice(0, 60) });
    }
    UrlFetchApp.fetch(DASH_SNAP_URL, { method: 'post', contentType: 'application/json',
      headers: { 'x-mailer-secret': SECRET },
      payload: JSON.stringify({ version: CODE_VERSION, items: items }), muteHttpExceptions: true });
  } catch (e) { /* דיווח בלבד — כישלון שלו לא נוגע בשליחות */ }
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
