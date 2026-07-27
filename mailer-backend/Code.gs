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
  if (!idToken) return null;
  const res = UrlFetchApp.fetch(
    'https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + FIREBASE_API_KEY,
    { method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ idToken: idToken }), muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) return null;
  const users = (JSON.parse(res.getContentText()) || {}).users || [];
  if (!users.length || !users[0].email) return null;
  return String(users[0].email).toLowerCase();
}

// ===================== נקודת הכניסה מהעמוד =====================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.secret !== SECRET) return json({ ok: false, error: 'unauthorized' });

    const email = verifiedEmail_(data.idToken);
    if (!email) return json({ ok: false, error: 'לא זוהה משתמש מחובר' });
    if (ALLOWED_EMAILS.indexOf(email) < 0) {
      return json({ ok: false, error: 'החשבון ' + email + ' אינו מורשה' });
    }
    data.sentBy = email;

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

// ===================== תזמון =====================
function queueScheduled_(data) {
  queueSheet_().appendRow([new Date(), data.sendAt, 'pending', JSON.stringify(data)]);
}

function checkScheduled() {
  const sh = queueSheet_();
  const rows = sh.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][2] !== 'pending') continue;
    if (new Date(rows[i][1]) <= now) {
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
