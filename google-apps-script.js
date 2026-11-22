/**
 * ====================================================================
 * Google 試算表後端程式碼 (Google Apps Script) - 完整雙向支援 (增/查/刪)
 * ====================================================================
 * 
 * 💡 特點：
 * 1. 支援查詢 (doGet)、新增 (doPost add)、刪除 (doPost delete)。
 * 2. 刪除時會真正從 Google 試算表中刪除該列 (sheet.deleteRow)。
 * 3. 自動辨識「表單回應 1」或資料最多的工作表分頁。
 * 4. 智慧區分「消費日期」與「時間戳記」，日期自動標準化為 YYYY-MM-DD。
 * 5. 類別名稱自動模糊比對「食、衣、住、行、育、樂」。
 * 
 * 【更新步驟】
 * 1. 打開 Google 試算表 -> 「擴充功能」 -> 「Apps Script」，貼上並儲存。
 * 2. 點選右上角「部署」 -> 「管理部署作業」。
 * 3. 點選右上角鉛筆圖示（編輯） -> 版本下拉選「新版本 (New Version)」 -> 點選「部署」！
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

    // 2. 類別欄位
    const catIdx = headers.findIndex(h => h.includes('類別') || h.includes('分類') || h.includes('category') || h.includes('項目'));

    // 3. 金額欄位
    const amtIdx = headers.findIndex(h => h.includes('金額') || h.includes('amount') || h.includes('花費') || h.includes('價錢') || h.includes('費用'));

    // 4. 備註欄位
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
        rowIndex: i + 1, // 試算表實際列號 (從 1 開始)
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

    const action = body.action || 'add';

    // -------------------------------------------------------------
    // 操作 A：刪除資料 (真正從試算表中刪除該列)
    // -------------------------------------------------------------
    if (action === 'delete') {
      const targetRow = Number(body.rowIndex);
      const targetDate = body.date;
      const targetAmount = Number(body.amount);
      const targetNote = body.note ? String(body.note).trim() : '';

      const data = sheet.getDataRange().getValues();
      let deleted = false;
      let deletedRow = -1;

      // 1. 若有指定 rowIndex，先檢查該列資料是否符合
      if (targetRow && targetRow >= 2 && targetRow <= data.length) {
        sheet.deleteRow(targetRow);
        deleted = true;
        deletedRow = targetRow;
      } else {
        // 2. 若 rowIndex 變動（如前面已有列被刪除），透過日期、金額、備註從最底下一列往上搜尋相符項目刪除
        const rawHeaders = data[0].map(h => String(h).trim().toLowerCase());
        const dateIdx = rawHeaders.findIndex(h => (h.includes('日期') || h.includes('date')) && !h.includes('戳記'));
        const amtIdx = rawHeaders.findIndex(h => h.includes('金額') || h.includes('amount') || h.includes('花費'));
        const noteIdx = rawHeaders.findIndex(h => h.includes('備註') || h.includes('說明') || h.includes('note') || h.includes('品項'));

        for (let i = data.length - 1; i >= 1; i--) {
          const row = data[i];
          const d = normalizeDate(dateIdx !== -1 ? row[dateIdx] : row[0]);
          const rawAmt = amtIdx !== -1 ? row[amtIdx] : 0;
          const a = Number(String(rawAmt).replace(/[$,NT\s]/gi, '').replace(/,/g, '')) || 0;
          const n = noteIdx !== -1 ? String(row[noteIdx] || '').trim() : '';

          if (d === targetDate && a === targetAmount && (!targetNote || n.includes(targetNote) || targetNote.includes(n))) {
            sheet.deleteRow(i + 1);
            deleted = true;
            deletedRow = i + 1;
            break;
          }
        }
      }

      return responseJSON({
        success: true,
        deleted: deleted,
        deletedRow: deletedRow,
        message: deleted ? '已成功從 Google 試算表刪除！' : '未在試算表中找到吻合的資料'
      });
    }

    // -------------------------------------------------------------
    // 操作 B：新增記帳資料
    // -------------------------------------------------------------
    const date = body.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const category = body.category || '食';
    const amount = Number(body.amount) || 0;
    const note = body.note || '';
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    
    // 如果是全新的表單，先建立標題
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['時間戳記', '消費日期', '消費類別', '金額', '備註說明']);
    }
    
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

// 智慧選擇目標工作表
function getTargetSheet(ss) {
  const sheets = ss.getSheets();
  if (sheets.length === 1) return sheets[0];

  for (let i = 0; i < sheets.length; i++) {
    const name = sheets[i].getName();
    if (name.includes('表單回應') || name.toLowerCase().includes('form responses')) {
      return sheets[i];
    }
  }

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

// 日期標準化
function normalizeDate(raw) {
  if (!raw) return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  if (raw instanceof Date) {
    return Utilities.formatDate(raw, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  const str = String(raw).trim();
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

// 類別模糊比對
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

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
