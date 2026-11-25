/**
 * Google Apps Script — שאלון סיום למסע "להעיר את השפע"
 *
 * כולל:
 * 1. שמירה מסודרת בטאב "משוב" בימין-לשמאל (RTL) עם גלישת טקסט ועמודות רחבות וקריאות.
 * 2. שליחת מייל מעוצב ונוח לקריאה בכל פעם שמשתתף/ת שולחים משוב.
 *
 * לעדכון סקריפט קיים (בלי לשנות את הכתובת):
 * 1. הדביקי את הקוד הזה במקום הקוד הקיים ושמרי (Cmd+S).
 * 2. כדי לסדר מיד את הטאב הקיים: בחרי למעלה בפונקציה formatExistingSheet ולחצי Run (אשרי הרשאות מייל אם יתבקש).
 * 3. לחצי Deploy -> Manage deployments -> אייקון העיפרון -> Version: New version -> Deploy.
 */

const SHEET_NAME = 'משוב';
const NOTIFY_EMAIL = 'helitsayfan@gmail.com'; // כתובת המייל לקבלת סיכום מעוצב של כל משוב (השאירי ריק "" כדי לכבות)

const HEADERS = [
  'תאריך ושעה',
  '1. כוונה ומטרה',
  '2. התמדה ב-21 הימים',
  '3. חוויה כללית (1-5)',
  '4. שבוע שהכי התחברו אליו',
  '5. מה הכי עבד במדיטציות',
  '6. שבוע/פרק שפחות עבד או שנשאר במיוחד',
  '7. משימות העומק היומיות',
  '8. נושא שהיה חסר או תפס יותר מדי מקום',
  '9. המעטפת (אורך, קבוצה, סקר ערב)',
  '10. מה זז, נפתח או השתנה',
  '11. המלצה לחבר/ה (1-10)',
  '12. תיאור המסע בשני משפטים (ציטוט)',
  '13. אישור ציטוט',
  '14. שם'
];

const COLUMN_WIDTHS = [
  135, // תאריך ושעה
  240, // 1. כוונה
  200, // 2. התמדה
  100, // 3. חוויה 1-5
  230, // 4. שבוע
  240, // 5. מדיטציות
  260, // 6. שבוע/פרק פתוח
  240, // 7. משימות
  250, // 8. חסר/מיותר
  250, // 9. מעטפת
  260, // 10. מה השתנה
  100, // 11. המלצה 1-10
  280, // 12. ציטוט
  190, // 13. אישור ציטוט
  140  // 14. שם
];

function applySheetFormatting(sheet) {
  sheet.setRightToLeft(true);
  sheet.setFrozenRows(1);

  COLUMN_WIDTHS.forEach(function (w, idx) {
    sheet.setColumnWidth(idx + 1, w);
  });

  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange
    .setFontWeight('bold')
    .setBackground('#234444')
    .setFontColor('#ffffff')
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('right');
  sheet.setRowHeight(1, 44);

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const dataRange = sheet.getRange(2, 1, lastRow - 1, HEADERS.length);
    dataRange
      .setWrap(true)
      .setVerticalAlignment('top')
      .setHorizontalAlignment('right')
      .setFontSize(10);

    // יישור מרכזי לעמודות הציון (1-5 ו-1-10)
    sheet.getRange(2, 4, lastRow - 1, 1).setHorizontalAlignment('center');
    sheet.getRange(2, 12, lastRow - 1, 1).setHorizontalAlignment('center');
  }
}

/**
 * הריצי פונקציה זו פעם אחת מתוך Apps Script (Run) כדי לסדר מיד את הטאב הקיים
 */
function formatExistingSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (sheet) {
    applySheetFormatting(sheet);
  }
}

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\n/g, '<br>');
}

function sendFeedbackEmail(data, formattedDate) {
  if (!NOTIFY_EMAIL) return;

  const rows = [
    ['1. כוונה ומטרה', data.q1_intent],
    ['2. התמדה ב-21 הימים', data.q2_consistency],
    ['3. חוויה כללית (1–5)', data.q3_overall ? data.q3_overall + ' / 5' : ''],
    ['4. שבוע שהכי התחברו אליו', data.q4_weeks],
    ['5. מה הכי עבד במדיטציות', data.q5_meditations],
    ['6. שבוע/פרק שפחות עבד או שנשאר במיוחד', data.q6_weeks_chapters],
    ['7. משימות העומק היומיות', data.q7_tasks],
    ['8. נושא שהיה חסר או תפס יותר מדי מקום', data.q8_missing_topic],
    ['9. המעטפת (אורך, קבוצה, סקר ערב)', data.q9_envelope],
    ['10. מה זז, נפתח או השתנה', data.q10_shift],
    ['11. המלצה לחבר/ה (1–10)', data.q11_nps ? data.q11_nps + ' / 10' : ''],
    ['12. תיאור המסע בשני משפטים (ציטוט)', data.q12_quote],
    ['13. אישור ציטוט', data.q13_quote_permission],
    ['14. שם', data.q14_name]
  ];

  const itemsHtml = rows
    .filter(function (pair) { return pair[1] && String(pair[1]).trim() !== ''; })
    .map(function (pair) {
      const val = String(pair[1]).split(' | ').map(escapeHtml).join('<br>• ');
      const formattedVal = String(pair[1]).indexOf(' | ') !== -1 ? '• ' + val : escapeHtml(pair[1]);
      return (
        '<div style="padding:12px 0;border-bottom:1px solid #e7f0ec;">' +
        '<div style="font-size:13px;font-weight:700;color:#375c5c;margin-bottom:4px;">' + escapeHtml(pair[0]) + '</div>' +
        '<div style="font-size:15px;color:#1f3b3b;line-height:1.6;">' + formattedVal + '</div>' +
        '</div>'
      );
    })
    .join('');

  const htmlBody =
    '<div dir="rtl" style="font-family:Arial,sans-serif;max-width:620px;margin:0 auto;background:#fffdfb;border:1px solid #e7f0ec;border-radius:16px;overflow:hidden;">' +
      '<div style="background:#234444;color:#ffffff;padding:20px 24px;">' +
        '<div style="font-size:12px;opacity:.8;margin-bottom:4px;">להעיר את השפע · משוב חדש</div>' +
        '<div style="font-size:20px;font-weight:700;">התקבל משוב חדש (' + escapeHtml(data.q3_overall || '-') + '/5 · המלצה ' + escapeHtml(data.q11_nps || '-') + '/10)</div>' +
        '<div style="font-size:12px;opacity:.75;margin-top:6px;">' + escapeHtml(formattedDate) + (data.q14_name ? ' · ' + escapeHtml(data.q14_name) : ' · בעילום שם') + '</div>' +
      '</div>' +
      '<div style="padding:12px 24px 20px;">' + itemsHtml + '</div>' +
    '</div>';

  const subject = 'משוב חדש: להעיר את השפע (' + (data.q3_overall || '-') + '/5)' + (data.q14_name ? ' — ' + data.q14_name : '');

  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: subject,
    htmlBody: htmlBody
  });
}

function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    const data = JSON.parse(raw);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    const isNewSheet = !sheet;
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      applySheetFormatting(sheet);
    } else if (isNewSheet) {
      applySheetFormatting(sheet);
    }

    const formattedDate = Utilities.formatDate(
      new Date(),
      'Asia/Jerusalem',
      'dd/MM/yyyy HH:mm'
    );

    sheet.appendRow([
      formattedDate,
      data.q1_intent || '',
      data.q2_consistency || '',
      data.q3_overall || '',
      data.q4_weeks || '',
      data.q5_meditations || '',
      data.q6_weeks_chapters || '',
      data.q7_tasks || '',
      data.q8_missing_topic || '',
      data.q9_envelope || '',
      data.q10_shift || '',
      data.q11_nps || '',
      data.q12_quote || '',
      data.q13_quote_permission || '',
      data.q14_name || ''
    ]);

    const newRowIdx = sheet.getLastRow();
    const rowRange = sheet.getRange(newRowIdx, 1, 1, HEADERS.length);
    rowRange
      .setWrap(true)
      .setVerticalAlignment('top')
      .setHorizontalAlignment('right')
      .setFontSize(10)
      .setBackground(newRowIdx % 2 === 0 ? '#ffffff' : '#f7faf8');
    sheet.getRange(newRowIdx, 4).setHorizontalAlignment('center');
    sheet.getRange(newRowIdx, 12).setHorizontalAlignment('center');

    try {
      sendFeedbackEmail(data, formattedDate);
    } catch (mailErr) {
      console.error('Email notification failed:', mailErr);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Abundance 21 Feedback endpoint is live' }))
    .setMimeType(ContentService.MimeType.JSON);
}
