/**
 * מערכת אישור איסוף ארבעת המינים
 * -------------------------------
 * הסקריפט הזה מתחבר לגיליון ההזמנות ומאפשר לדף index.html
 * (המתארח ב-GitHub Pages) לבדוק מספר טלפון מול הגיליון,
 * ולסמן שהאדם אסף את הסט שהזמין.
 *
 * התקנה (מהאייפון, בלי צורך במחשב):
 * 1. פתחו את גיליון ההזמנות ב-Safari.
 * 2. תפריט ה-☰ למעלה > "שלוחות" (Extensions) > "Apps Script".
 * 3. מחקו את הקוד לדוגמה שכבר שם, והדביקו את כל הקובץ הזה במקומו.
 * 4. שמרו (סמל הדיסקט).
 * 5. "פרוס" (Deploy) > "פריסה חדשה" (New deployment).
 * 6. ליד "בחר סוג" (Select type) בחרו "אפליקציית אינטרנט" (Web app).
 * 7. "מבצע הפעולה" (Execute as): אני (Me).
 *    "מי יכול לגשת" (Who has access): כל אחד (Anyone).
 * 8. "פרוס" (Deploy) ואז אשרו את בקשות ההרשאה
 *    (אם מופיעה אזהרה, לחצו "מתקדם" ואז "המשך").
 * 9. העתיקו את הכתובת (URL) שמסתיימת ב-/exec.
 * 10. הדביקו אותה בקובץ index.html במקום המסומן WEBAPP_URL.
 *
 * אם בעתיד תעדכנו את הקוד הזה, בחרו "פריסה חדשה" שוב כדי
 * שהשינויים ייכנסו לתוקף (עריכה בלבד לא מספיקה).
 */

var NAME_HEADER = 'שם מלא';
var PHONE_HEADER = 'טלפון';
var CONFIRM_HEADER = 'אישור איסוף';
var CONFIRM_TIME_HEADER = 'זמן אישור';

function doPost(e) {
  try {
    var params = (e && e.parameter) || {};
    var action = params.action;
    var phone = normalizePhone(params.phone);

    if (!phone) {
      return jsonResponse({ ok: false, error: 'לא התקבל מספר טלפון.' });
    }

    var sheet = findResponseSheet();
    if (!sheet) {
      return jsonResponse({ ok: false, error: 'לא נמצא בגיליון עמודות בשם "שם מלא" ו"טלפון".' });
    }

    var headers = sheet.getDataRange().getValues()[0];
    var nameCol = headers.indexOf(NAME_HEADER);
    var phoneCol = headers.indexOf(PHONE_HEADER);
    var cols = ensureConfirmColumns(sheet, headers);

    // קוראים מחדש אחרי שאולי נוספו עמודות
    var data = sheet.getDataRange().getValues();

    var matches = [];
    for (var i = 1; i < data.length; i++) {
      var rowPhone = data[i][phoneCol];
      if (rowPhone !== '' && normalizePhone(rowPhone) === phone) {
        matches.push(i);
      }
    }

    if (matches.length === 0) {
      return jsonResponse({ ok: true, found: false });
    }

    var firstRow = matches[0];
    var name = data[firstRow][nameCol];

    if (action === 'confirm') {
      var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
      matches.forEach(function (r) {
        sheet.getRange(r + 1, cols.confirmCol + 1).setValue('כן');
        sheet.getRange(r + 1, cols.confirmTimeCol + 1).setValue(now);
      });
      return jsonResponse({ ok: true, found: true, name: name, confirmed: true });
    }

    var alreadyConfirmed = matches.every(function (r) {
      return data[r][cols.confirmCol] === 'כן';
    });
    return jsonResponse({ ok: true, found: true, name: name, alreadyConfirmed: alreadyConfirmed });

  } catch (err) {
    return jsonResponse({ ok: false, error: 'שגיאת מערכת: ' + err.message });
  }
}

function doGet(e) {
  return ContentService.createTextOutput('המערכת פעילה. יש לגשת אליה דרך דף האינטרנט של הדיווח.');
}

function findResponseSheet() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var values = sheets[i].getDataRange().getValues();
    var headers = values[0] || [];
    if (headers.indexOf(NAME_HEADER) !== -1 && headers.indexOf(PHONE_HEADER) !== -1) {
      return sheets[i];
    }
  }
  return null;
}

function ensureConfirmColumns(sheet, headers) {
  var confirmCol = headers.indexOf(CONFIRM_HEADER);
  var confirmTimeCol = headers.indexOf(CONFIRM_TIME_HEADER);

  if (confirmCol === -1) {
    confirmCol = sheet.getLastColumn();
    sheet.getRange(1, confirmCol + 1).setValue(CONFIRM_HEADER);
  }
  if (confirmTimeCol === -1) {
    confirmTimeCol = sheet.getLastColumn();
    sheet.getRange(1, confirmTimeCol + 1).setValue(CONFIRM_TIME_HEADER);
  }
  return { confirmCol: confirmCol, confirmTimeCol: confirmTimeCol };
}

function normalizePhone(p) {
  return String(p || '').replace(/[^0-9]/g, '').replace(/^972/, '0');
}
