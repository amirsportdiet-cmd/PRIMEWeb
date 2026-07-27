/**
 * מערכת מיילים PRIME — שרת שליחה (Google Apps Script)
 * לפריסה תחת החשבון primerct2026@gmail.com
 */

// ===== הגדרות =====
const SECRET = 'lgGnSJZnAsfIs4W822y0k7F6';   // מפתח אבטחה (זהה לזה שבעמוד)
const SENDER_NAME = 'מחקר PRIME';

// מזהה תיקיית ה-Drive שבה נמצאים הקבצים המצורפים לכל שלב.
// כל קבצי ה-PDF שבתיקייה יצורפו אוטומטית. יתמלא אחרי יצירת התיקייה.
// מפתח = site:phase[:group] — כמו fileKey() בעמוד
const ATTACH_FOLDER = {
  'assuta:start':             'PASTE_T0_FOLDER_ID',
  'assuta:end:control':       'PASTE_T6_CONTROL_FOLDER_ID',
  'assuta:end:intervention':  'PASTE_T6_INTERVENTION_FOLDER_ID'
  // אמצע-מחקר ואיכילוב — יתווספו בהמשך
};

// מיילים מורשים לשלוח (חייב להיות זהה לרשימה בעמוד)
const ALLOWED_EMAILS = ['amirsportdiet@gmail.com', 'primerct2026@gmail.com'];
// מפתח ה-Web API של פרויקט Firebase (ציבורי — משמש רק לאימות האסימון)
const FIREBASE_API_KEY = 'AIzaSyCZR4jxDQ8hcmfslbgx06dlRQNIoTf3Wss';

/**
 * מאמת את אסימון הזהות מול Google ומחזיר את המייל, או null אם לא תקין.
 * זו שכבת האבטחה האמיתית: המפתח הסודי גלוי בקוד העמוד, ולכן לבדו אינו מספיק.
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
  if (users[0].emailVerified === false) return null;
  return String(users[0].email).toLowerCase();
}

// ===== נקודת הכניסה מהעמוד =====
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.secret !== SECRET) return json({ ok: false, error: 'unauthorized' });

    // אימות זהות אמיתי — מי באמת שולח את הבקשה
    const email = verifiedEmail_(data.idToken);
    if (!email) return json({ ok: false, error: 'לא זוהה משתמש מחובר' });
    if (ALLOWED_EMAILS.indexOf(email) < 0) {
      return json({ ok: false, error: 'החשבון ' + email + ' אינו מורשה' });
    }
    data.sentBy = email;

    if (data.mode === 'now') {
      sendMail_(data);
      return json({ ok: true, sent: true });
    } else {
      queueScheduled_(data);
      return json({ ok: true, scheduled: true, sendAt: data.sendAt });
    }
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// ===== שליחת מייל =====
function sendMail_(data) {
  GmailApp.sendEmail(data.to.join(','), data.subject, data.body, {
    name: SENDER_NAME,
    attachments: attachmentsFor_(data.fileKey)
  });
}

function attachmentsFor_(fileKey) {
  const folderId = ATTACH_FOLDER[fileKey];
  if (!folderId || folderId.indexOf('PASTE_') === 0) return [];
  const files = DriveApp.getFolderById(folderId).getFiles();
  const out = [];
  while (files.hasNext()) out.push(files.next().getBlob());
  return out;
}

// ===== תור לתזמון =====
function queueScheduled_(data) {
  queueSheet_().appendRow([new Date(), data.sendAt, 'pending', JSON.stringify(data)]);
}

// רץ כל 15 דקות (טריגר) — שולח את מה שהגיע זמנו
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
  let id = props.getProperty('QUEUE_SHEET_ID');
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

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// הרץ פעם אחת ידנית — מתקין את הטריגר לתזמון
function installTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'checkScheduled') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkScheduled').timeBased().everyMinutes(15).create();
}

// בדיקה ידנית — שולח מייל בדיקה לעצמך (הרץ אותו כדי לאשר הרשאות)
function testSend() {
  GmailApp.sendEmail('primerct2026@gmail.com', 'בדיקה — מערכת מיילים PRIME',
    'זו בדיקה. אם קיבלת את זה, השליחה עובדת.', { name: SENDER_NAME });
}
