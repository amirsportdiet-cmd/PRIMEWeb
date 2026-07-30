/**
 * קורא פניות נחקרים מהמייל — Google Apps Script
 *
 * ⚠️ הפרויקט הזה חייב לשבת תחת amirsportdiet@gmail.com ולא תחת primerct2026,
 * משום שמיילי ההתראה מגיעים לתיבה האישית. אפליקציית ווב שנפרסת עם
 * "הפעלה בתור: אני" קוראת את התיבה של בעל הפרויקט בלבד.
 *
 * התקנה: להריץ פעם אחת את selfTest(), לאשר הרשאות, ואז לפרוס כאפליקציית ווב
 * (הפעלה בתור: אני · גישה: כולם) ולהעתיק את כתובת ה-exec ל-prime-alerts.html.
 */
const CODE_VERSION = 'alerts-v1';
const SECRET = 'lgGnSJZnAsfIs4W822y0k7F6';
const ALLOWED_EMAILS = ['amirsportdiet@gmail.com', 'primerct2026@gmail.com'];
const FIREBASE_API_KEY = 'AIzaSyCZR4jxDQ8hcmfslbgx06dlRQNIoTf3Wss';

// המיילים מגיעים תמיד מאותו שולח, ולרוב נוחתים בלשונית "קידום מכירות".
// החיפוש כולל את כל הלשוניות במפורש, אחרת Gmail מדלג עליהן.
const SENDER = 'send.vpcontact.com';

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

/** מאמת את אסימון הזיהוי מול Firebase ומחזיר את המייל */
function verifiedEmail_(idToken) {
  if (!idToken) return { email: null, err: 'חסר אסימון' };
  try {
    const res = UrlFetchApp.fetch(
      'https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + FIREBASE_API_KEY,
      { method: 'post', contentType: 'application/json',
        payload: JSON.stringify({ idToken: idToken }), muteHttpExceptions: true });
    const body = JSON.parse(res.getContentText());
    if (!body.users || !body.users.length) return { email: null, err: 'אסימון לא תקף' };
    return { email: String(body.users[0].email || '').toLowerCase(), err: null };
  } catch (e) {
    return { email: null, err: String(e) };
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.secret !== SECRET) return json({ ok: false, error: 'unauthorized' });

    if (data.mode === 'ping') {
      return json({ ok: true, version: CODE_VERSION, mailbox: Session.getEffectiveUser().getEmail(),
                    probe: countMessages_(30) });
    }

    const v = verifiedEmail_(data.idToken);
    if (!v.email) return json({ ok: false, error: 'זיהוי נכשל — ' + v.err });
    if (ALLOWED_EMAILS.indexOf(v.email) < 0) {
      return json({ ok: false, error: 'החשבון ' + v.email + ' אינו מורשה' });
    }

    if (data.mode === 'inbox') return json(fetchAlerts_(data.days, data.since));
    return json({ ok: false, error: 'mode לא מוכר: ' + data.mode });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

/** בונה את שאילתת החיפוש. category:promotions נכלל במפורש כי שם המיילים נוחתים. */
function query_(days) {
  const d = Math.max(1, Math.min(365, days || 30));
  return 'from:' + SENDER + ' newer_than:' + d + 'd ' +
         '{category:promotions category:primary category:updates category:social -category:promotions}';
}

/** ספירה בלבד, בלי תוכן — משמשת לבדיקת חיבור */
function countMessages_(days) {
  try {
    const q = 'from:' + SENDER + ' newer_than:' + Math.max(1, days || 30) + 'd';
    return { query: q, threads: GmailApp.search(q, 0, 200).length };
  } catch (e) {
    return { error: String(e) };
  }
}

/**
 * מחזיר את מיילי ההתראה. הפרסור עצמו נעשה בדפדפן,
 * כדי שנוסח חדש יטופל בלי לפרוס מחדש את הסקריפט.
 * since = חותמת זמן ISO, מחזיר רק מה שחדש ממנה.
 */
function fetchAlerts_(days, since) {
  const q = 'from:' + SENDER + ' newer_than:' + Math.max(1, Math.min(365, days || 30)) + 'd';
  const threads = GmailApp.search(q, 0, 200);
  const cut = since ? new Date(since).getTime() : 0;
  const items = [];
  for (var i = 0; i < threads.length; i++) {
    const msgs = threads[i].getMessages();
    for (var j = 0; j < msgs.length; j++) {
      const m = msgs[j];
      const when = m.getDate().getTime();
      if (cut && when <= cut) continue;
      items.push({
        id: m.getId(),
        at: m.getDate().toISOString(),
        subject: m.getSubject(),
        body: m.getPlainBody().slice(0, 3000)
      });
    }
  }
  items.sort(function (a, b) { return a.at < b.at ? 1 : -1; });
  return { ok: true, query: q, mailbox: Session.getEffectiveUser().getEmail(),
           threads: threads.length, count: items.length, items: items };
}

/** להריץ ידנית פעם אחת: מאשר הרשאות ומדפיס כמה מיילים נמצאו */
function selfTest() {
  const r = fetchAlerts_(60, null);
  Logger.log('תיבה: ' + r.mailbox);
  Logger.log('שאילתה: ' + r.query);
  Logger.log('נמצאו ' + r.count + ' מיילים ב-' + r.threads + ' שרשורים');
  if (r.items.length) Logger.log('האחרון: ' + r.items[0].subject);
  return r.count;
}
