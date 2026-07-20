/**
 * שולח וואטסאפ אוטומטי לספקי סלטי גן — Google Apps Script.
 * ------------------------------------------------------------------
 * קורא שורות "pending" מלשונית WA_OUTBOX בגיליון הרכש, שולח כל אחת דרך
 * WhatsApp Cloud API (מטא), ומעדכן את הסטטוס ל-"sent"/"failed" + מזהה ההודעה.
 * מריצים כטריגר מתוזמן (כל דקה/כמה דקות) — ראה setup ב-README.md.
 *
 * אבטחה: הטוקן של מטא ומזהה מספר הטלפון נשמרים ב-Script Properties בלבד,
 * לעולם לא בקובץ הזה ולא בקוד הציבורי של האתר.
 *
 * מבנה WA_OUTBOX (נכתב ע"י purchasing.html → waAuto):
 *   A תאריך | B מפתח ספק | C שם ספק | D טלפון (E.164 בלי +) | E הודעה | F סטטוס | G מזהה הודעה
 *
 * Script Properties הנדרשים (File ▸ Project properties ▸ Script properties):
 *   SHEET_ID          מזהה גיליון הרכש (1rWHMhO8zCB8KKzAJwyFYpuKfo_EQ_-rZB8afaiqUv9Q)
 *   META_TOKEN        טוקן קבוע (System User) מ-WhatsApp Business Platform
 *   PHONE_NUMBER_ID   מזהה מספר הטלפון של הממשק
 *   TEMPLATE_NAME     שם התבנית המאושרת (למשל: purchase_order)
 *   TEMPLATE_LANG     קוד שפת התבנית (למשל: he)
 *
 * התבנית (קטגוריית Utility) מצופה עם שלושה משתני body בסדר:
 *   {{1}} שם הספק · {{2}} טקסט ההזמנה (רשימת הפריטים) · {{3}} תאריך
 */

var OUTBOX_TAB = 'WA_OUTBOX';

function sendPendingOrders() {
  var props = PropertiesService.getScriptProperties();
  var SHEET_ID = props.getProperty('SHEET_ID');
  var TOKEN = props.getProperty('META_TOKEN');
  var PHONE_ID = props.getProperty('PHONE_NUMBER_ID');
  var TEMPLATE = props.getProperty('TEMPLATE_NAME');
  var LANG = props.getProperty('TEMPLATE_LANG') || 'he';
  if (!SHEET_ID || !TOKEN || !PHONE_ID || !TEMPLATE) {
    Logger.log('חסרים Script Properties — ראה README');
    return;
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(OUTBOX_TAB);
  if (!sh) { Logger.log('אין לשונית ' + OUTBOX_TAB); return; }

  var rng = sh.getDataRange();
  var vals = rng.getValues();
  if (vals.length < 2) return;                 // רק כותרת
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;            // מונע ריצה כפולה של הטריגר

  try {
    for (var r = 1; r < vals.length; r++) {
      var row = vals[r];
      var status = String(row[5] || '').toLowerCase();
      if (status !== 'pending') continue;       // כבר נשלח / נכשל / ריק
      var name = String(row[2] || '');
      var phone = String(row[3] || '').replace(/\D/g, '');
      var body = String(row[4] || '');
      var date = String(row[0] || '');
      if (!phone) { setStatus(sh, r, 'failed', 'no-phone'); continue; }

      // מטא אוסרת מעברי-שורה/טאבים/>4 רווחים בתוך משתנה תבנית — ממירים לרשימה עם מפריד "•"
      var bodyParam = body.replace(/[\r\n\t]+/g, '  •  ').replace(/ {4,}/g, '   ').trim();
      var res = sendTemplate(PHONE_ID, TOKEN, TEMPLATE, LANG, phone, [name, bodyParam, date]);
      if (res.ok) setStatus(sh, r, 'sent', res.id);
      else setStatus(sh, r, 'failed', res.err);
    }
  } finally {
    lock.releaseLock();
  }
}

function sendTemplate(phoneId, token, template, lang, toPhone, bodyParams) {
  var url = 'https://graph.facebook.com/v21.0/' + phoneId + '/messages';
  var payload = {
    messaging_product: 'whatsapp',
    to: toPhone,
    type: 'template',
    template: {
      name: template,
      language: { code: lang },
      components: [{
        type: 'body',
        parameters: bodyParams.map(function (p) { return { type: 'text', text: String(p || '') }; })
      }]
    }
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    var resp = UrlFetchApp.fetch(url, options);
    var code = resp.getResponseCode();
    var j = JSON.parse(resp.getContentText() || '{}');
    if (code >= 200 && code < 300 && j.messages && j.messages[0]) {
      return { ok: true, id: j.messages[0].id };
    }
    return { ok: false, err: (j.error && j.error.message) ? j.error.message : ('HTTP ' + code) };
  } catch (e) {
    return { ok: false, err: String(e) };
  }
}

function setStatus(sh, rowIdx, status, note) {
  sh.getRange(rowIdx + 1, 6).setValue(status);              // F סטטוס
  sh.getRange(rowIdx + 1, 7).setValue(note || '');          // G מזהה הודעה / שגיאה
}
