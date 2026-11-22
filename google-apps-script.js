/**
 * ====================================================================
 * Google 試算表後端程式碼 (Google Apps Script)
 * ====================================================================
 * 
 * 【使用步驟】
 * 1. 打開你的 Google 試算表（或 Google 表單連結的試算表）。
 * 2. 點選上方選單：「擴充功能」 -> 「Apps Script」。
 * 3. 刪除原本所有內容，將這整份程式碼完整複製並貼上。
 * 4. 點選右上角「部署」 -> 「新增部署作業」。
 * 5. 類型選擇「網頁應用程式 (Web App)」：
 *    - 說明：記帳小幫手 API
 *    - 執行身分：我 (你的帳號)
 *    - 誰可以存取：任何人 (Anyone)  <--- 務必選擇這個，網頁才能存取！
 * 6. 點選「部署」，並授予 Google 帳號權限。
 * 7. 複製產生的「網頁應用程式網址 (Web App URL)」，貼到記帳網頁的設定中即可！
 */

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return responseJSON({ success: true, data: [] });
    }
    
    // 第一列是標題列
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    
    // 尋找各欄位索引（支援中文與常見英文）
    const dateIdx = headers.findIndex(h => h.includes('日期') || h.includes('date') || h.includes('時間'));
    const catIdx = headers.findIndex(h => h.includes('類別') || h.includes('分類') || h.includes('category') || h.includes('項目'));
    const amtIdx = headers.findIndex(h => h.includes('金額') || h.includes('amount') || h.includes('花費') || h.includes('價錢'));
    const noteIdx = headers.findIndex(h => h.includes('備註') || h.includes('說明') || h.includes('note') || h.includes('品項') || h.includes('內容'));

    const records = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // 如果該列全部為空則略過
      if (row.every(cell => cell === '')) continue;
      
      let rawDate = dateIdx !== -1 ? row[dateIdx] : row[0];
      let formattedDate = '';
      if (rawDate instanceof Date) {
        formattedDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        formattedDate = String(rawDate || '').split('T')[0];
      }
      
      let category = catIdx !== -1 ? String(row[catIdx] || '').trim() : '食';
      let amount = amtIdx !== -1 ? Number(row[amtIdx]) || 0 : 0;
      let note = noteIdx !== -1 ? String(row[noteIdx] || '').trim() : '';
      
      records.push({
        id: 'gs_' + (i + 1),
        date: formattedDate,
        category: category,
        amount: amount,
        note: note
      });
    }

    return responseJSON({ success: true, data: records });
  } catch (error) {
    return responseJSON({ success: false, error: error.toString() });
  }
}

function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let body = {};
    if (e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      body = e.parameter;
    }
    
    const date = body.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const category = body.category || '食';
    const amount = Number(body.amount) || 0;
    const note = body.note || '';
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    
    // 如果是全新的表單，先建立標題
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['時間戳記', '消費日期', '消費類別', '金額', '備註說明']);
    }
    
    // 新增資料
    sheet.appendRow([timestamp, date, category, amount, note]);
    
    return responseJSON({
      success: true,
      message: '記帳成功！',
      data: { date: date, category: category, amount: amount, note: note }
    });
  } catch (error) {
    return responseJSON({ success: false, error: error.toString() });
  }
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
