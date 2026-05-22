/**
 * 阿嬤的味道 — 粽子預購訂單收集
 *
 * 這段程式貼進 Google Apps Script，部署成「網頁應用程式」後，
 * 網站的訂購表單送出時，訂單就會自動寫進這份 Google 試算表。
 *
 * 設定步驟詳見同資料夾的「Google試算表設定說明.txt」
 */

// 訂單會寫進這個分頁，沒有的話程式會自動建立
var SHEET_NAME = '訂單';

// 表單欄位 → 試算表欄位的對應與順序
var HEADERS = ['訂購時間', '姓名', '電話', '7-11取貨門市',
               '1顆(份)', '10顆(份)', '20顆(份)', '加購小菜(份)',
               '冷凍運費', '應付總額', '轉帳後五碼', '備註'];

function doPost(e) {
  // 加鎖，避免兩張訂單同時寫入時互相覆蓋
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);

    // 第一次執行時，自動建立分頁與標題列
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    var p = e.parameter; // 表單送來的欄位

    sheet.appendRow([
      new Date(),
      p.name    || '',
      p.phone   || '',
      p.store   || '',
      p.qty1    || '0',
      p.qty10   || '0',
      p.qty20   || '0',
      p.addon   || '0',
      p.freight || '0',
      p.total   || '0',
      p.last5   || '',
      p.note    || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

// 用瀏覽器直接開 Web App 網址時會看到這句，代表部署成功
function doGet() {
  return ContentService.createTextOutput('阿嬤的味道訂單系統運作中 ❤');
}
