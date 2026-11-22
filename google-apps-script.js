/**
 * ====================================================================
 * Google 試算表後端程式碼 (Google Apps Script) - 增強版
 * ====================================================================
 * 
 * 💡 特點：
 * 1. 自動辨識「表單回應 1」或資料最多的工作表分頁（不再誤抓空白工作表）。
 * 2. 智慧區分「消費日期」與「時間戳記」，日期自動標準化為 YYYY-MM-DD。
 * 3. 類別名稱自動模糊比對「食、衣、住、行、育、樂」。
 * 4. 金額欄位自動去除 $ 符號與逗號。
 * 
 * 【重要更新步驟】
 * 每次修改此程式碼後：
 * 1. 打開 Google 試算表 -> 「擴充功能」 -> 「Apps Script」，貼上並儲存。
 * 2. 點選右上角「部署」 -> **「管理部署作業」**。
 * 3. 點選右上角鉛筆圖示（編輯） -> 版本下拉選**「新版本 (New Version)」** -> 點選「部署」！
 *    （⚠️ 務必建立新版本部署，網頁才能讀到最新程式碼！）
 */

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getTargetSheet(ss);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return responseJSON({ 
        success: true, 
        sheetName: sheet.getName(),
        recordsCount: 0,
        data: [] 
      });
    }
    
    // 第一列是標題列
    const rawHeaders = data[0].map(h => String(h).trim());
    const headers = rawHeaders.map(h => h.toLowerCase());
    
    // 智慧尋找欄位：
    // 1. 日期欄位：優先找包含「消費日」、「日期」、「date」的欄位（排除純時間戳記），若沒有才找「時間戳記」
    let dateIdx = headers.findIndex(h => (h.includes('消費日') || h.includes('日期') || h.includes('date')) && !h.includes('戳記'));
    if (dateIdx === -1) {
      dateIdx = headers.findIndex(h => h.includes('時間') || h.includes('戳記') || h.includes('time'));
    }

    // 2. 類別欄位：包含「類別」、「分類」、「category」、「項目」
    const catIdx = headers.findIndex(h => h.includes('類別') || h.includes('分類') || h.includes('category') || h.includes('項目'));

    // 3. 金額欄位：包含「金額」、「amount」、「花費」、「價錢」、「費用」
    const amtIdx = headers.findIndex(h => h.includes('金額') || h.includes('amount') || h.includes('花費') || h.includes('價錢') || h.includes('費用'));

    // 4. 備註欄位：包含「備註」、「說明」、「note」、「品項」、「內容」、「名稱」
    const noteIdx = headers.findIndex(h => h.includes('備註') || h.includes('說明') || h.includes('note') || h.includes('品項') || h.includes('內容') || h.includes('名稱'));

    const records = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row.every(cell => cell === '')) continue;
      
      let rawDate = dateIdx !== -1 ? row[dateIdx] : row[0];
      let formattedDate = normalizeDate(rawDate);
      
      let rawCategory = catIdx !== -1 ? String(row[catIdx] || '').trim() : '';
      let category = normalizeCategory(rawCategory);
      
      let rawAmt = amtIdx !== -1 ? row[amtIdx] : 0;
      let cleanAmt = String(rawAmt).replace(/[$,NT\s]/gi, '').replace(/,/g, '');
      let amount = Number(cleanAmt) || 0;

      let note = noteIdx !== -1 ? String(row[noteIdx] || '').trim() : '';
      if (!note) {
        note = category ? `${category}類消費` : '日常消費';
      }
      
      records.push({
        id: 'gs_' + (i + 1),
        date: formattedDate,
        category: category,
        amount: amount,
        note: note
      });
    }

    return responseJSON({ 
      success: true, 
      sheetName: sheet.getName(),
      recordsCount: records.length,
      headers: rawHeaders,
      data: records 
    });
  } catch (error) {
    return responseJSON({ success: false, error: error.toString() });
  }
}

// 智慧選擇目標工作表（優先選取表單回應分頁，或最多資料的分頁）
function getTargetSheet(ss) {
  const sheets = ss.getSheets();
  if (sheets.length === 1) return sheets[0];

  // 1. 優先尋找名稱含有「表單回應」或「Form Responses」的工作表
  for (let i = 0; i < sheets.length; i++) {
    const name = sheets[i].getName();
    if (name.includes('表單回應') || name.toLowerCase().includes('form responses')) {
      return sheets[i];
    }
  }

  // 2. 尋找資料列數最多且有內容的工作表
  let maxRows = -1;
  let target = sheets[0];
  for (let i = 0; i < sheets.length; i++) {
    const s = sheets[i];
    const rows = s.getLastRow();
    if (rows > maxRows) {
      maxRows = rows;
      target = s;
    }
  }
  return target;
}

// 日期標準化轉為 YYYY-MM-DD
function normalizeDate(raw) {
  if (!raw) return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  if (raw instanceof Date) {
    return Utilities.formatDate(raw, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  const str = String(raw).trim();
  // 比對 YYYY/M/D 或 YYYY-M-D 或 YYYY.M.D
  const match = str.match(/(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);
  if (match) {
    const y = match[1];
    const m = match[2].padStart(2, '0');
    const d = match[3].padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    const y = parsed.getFullYear();
    const m = String(parsed.getMonth() + 1).padStart(2, '0');
    const d = String(parsed.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  return str.split(' ')[0].replace(/\//g, '-');
}

// 類別模糊比對標準化為「食、衣、住、行、育、樂」
function normalizeCategory(raw) {
  if (!raw) return '食';
  const str = String(raw).trim();
  if (str.includes('食') || str.includes('餐') || str.includes('飲')) return '食';
  if (str.includes('衣') || str.includes('服') || str.includes('穿')) return '衣';
  if (str.includes('住') || str.includes('房') || str.includes('水電') || str.includes('瓦斯')) return '住';
  if (str.includes('行') || str.includes('車') || str.includes('交通') || str.includes('油')) return '行';
  if (str.includes('育') || str.includes('書') || str.includes('學') || str.includes('課')) return '育';
  if (str.includes('樂') || str.includes('玩') || str.includes('遊') || str.includes('休閒')) return '樂';
  return '食';
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getTargetSheet(ss);
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
