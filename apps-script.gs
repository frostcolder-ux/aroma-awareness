// ══════════════════════════════════════════════════════════════
//  精油香氛覺察表 — Google Apps Script 後端
//  使用方式：
//    1. 開啟 Google 試算表 → 擴充功能 → Apps Script
//    2. 貼上此程式碼，儲存
//    3. 部署 → 新增部署 → 類型選「網頁應用程式」
//       執行身分：「我自己」
//       誰可以存取：「所有人」
//    4. 複製部署後的 URL（格式：https://script.google.com/macros/s/...）
//    5. 貼入 awareness-form.html 和 admin.html 的 APPS_SCRIPT_URL 變數
// ══════════════════════════════════════════════════════════════

const SHEET_NAME = '覺察表資料';

// ── 接收表單送出（POST）────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let sheet   = ss.getSheetByName(SHEET_NAME);

    // 第一次執行自動建立工作表並加標題
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        '填答時間', '姓名', 'Email', '職業', '年齡',
        '壓力指數', '疲憊指數',
        '困擾第1名', '困擾第2名', '困擾第3名',
        '薰衣草', '玫瑰', '依蘭依蘭', '甜橙', '佛手柑',
        '薄荷', '迷迭香', '茶樹', '雪松', '乳香', '廣藿香', '尤加利',
        '原始JSON'
      ]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 23).setFontWeight('bold').setBackground('#0a2218').setFontColor('#ffffff');
    }

    const concerns = data.concerns || (data.concern ? [data.concern] : []);
    const oils     = data.oils || new Array(12).fill(0);

    sheet.appendRow([
      new Date(data.ts || Date.now()).toLocaleString('zh-TW'),
      data.name   || '',
      data.email  || '',
      data.job    || '',
      data.age    || '',
      data.stress  || 0,
      data.fatigue || 0,
      concerns[0] || '', concerns[1] || '', concerns[2] || '',
      ...oils.slice(0, 12),
      JSON.stringify(data)   // 完整備份
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 提供後台讀取資料（GET）────────────────────────────────
function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) {
      return ContentService
        .createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 從「原始JSON」欄位還原完整記錄（最後一欄）
    const lastCol  = sheet.getLastColumn();
    const lastRow  = sheet.getLastRow();
    const jsonCol  = sheet.getRange(2, lastCol, lastRow - 1, 1).getValues();

    const records = jsonCol
      .map(row => {
        try { return JSON.parse(row[0]); } catch { return null; }
      })
      .filter(Boolean);

    return ContentService
      .createTextOutput(JSON.stringify(records))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
