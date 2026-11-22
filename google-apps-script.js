/**
 * ====================================================================
 * Google 試算表後端程式碼 (Google Apps Script) - 旗艦智慧統計版
 * ====================================================================
 * 
 * 💡 本版本新增兩大核心功能：
 * 1. 【記錄新增與刪除狀態】：
 *    - 試算表會有一欄「狀態」，清楚標示「🟢 正常」或「🔴 已刪除 (含刪除時間)」，保留完整流水帳歷程。
 * 2. 【自動計算總金額與統計看板】：
 *    - 自動在試算表右側建立「📊 即時消費統計看板」，包含累積總金額、有效筆數、食衣住行育樂各類小計！
 * 
 * 【更新步驟】
 * 1. 打開 Google 試算表 -> 「擴充功能」 -> 「Apps Script」，將此程式碼覆蓋儲存。
 * 2. 點選右上角「部署」 -> 「管理部署作業」。
 * 3. 點選右上角鉛筆圖示（編輯） -> 版本選擇「新版本 (New Version)」 -> 點選「部署」完成！
 */

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getTargetSheet(ss);
    
    // 確保右側統計看板存在
    ensureDashboard(sheet);
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return responseJSON({ 
        success: true, 
        sheetName: sheet.getName(),
        recordsCount: 0,
        totalAmount: 0,
        data: [] 
      });
    }
    
    // 第一列是標題列
    const rawHeaders = data[0].map(h => String(h).trim());
    const headers = rawHeaders.map(h => h.toLowerCase());
    
    let dateIdx = headers.findIndex(h => (h.includes('消費日') || h.includes('日期') || h.includes('date')) && !h.includes('戳記'));
    if (dateIdx === -1) {
      dateIdx = headers.findIndex(h => h.includes('時間') || h.includes('戳記') || h.includes('time'));
    }

    const catIdx = headers.findIndex(h => h.includes('類別') || h.includes('分類') || h.includes('category') || h.includes('項目'));
    const amtIdx = headers.findIndex(h => h.includes('金額') || h.includes('amount') || h.includes('花費') || h.includes('價錢') || h.includes('費用'));
    const noteIdx = headers.findIndex(h => h.includes('備註') || h.includes('說明') || h.includes('note') || h.includes('品項') || h.includes('內容') || h.includes('名稱'));
    const statusIdx = headers.findIndex(h => h.includes('狀態') || h.includes('status'));

    const records = [];
    let grandTotal = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row.every(cell => cell === '')) continue;
      
      // 檢查狀態：若已標記為刪除，則略過（不傳入網頁統計）
      const statusText = statusIdx !== -1 ? String(row[statusIdx] || '') : '';
      if (statusText.includes('刪除') || statusText.includes('delete')) {
        continue;
      }
      
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
      
      grandTotal += amount;
      
      records.push({
        id: 'gs_' + (i + 1),
        rowIndex: i + 1,
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
      totalAmount: grandTotal,
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
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    // 檢查欄位標題與狀態欄位
    ensureHeadersAndStatus(sheet);

    // -------------------------------------------------------------
    // 操作 A：刪除操作 (將狀態標記為「🔴 已刪除」並加刪除線，保留操作紀錄)
    // -------------------------------------------------------------
    if (action === 'delete') {
      const targetRow = Number(body.rowIndex);
      const targetDate = body.date;
      const targetAmount = Number(body.amount);
      const targetNote = body.note ? String(body.note).trim() : '';

      const data = sheet.getDataRange().getValues();
      let matchedRow = -1;

      // 1. 若有 rowIndex 優先比對
      if (targetRow && targetRow >= 2 && targetRow <= sheet.getLastRow()) {
        matchedRow = targetRow;
      } else {
        // 2. 倒序比對日期、金額
        const rawHeaders = data[0].map(h => String(h).trim().toLowerCase());
        const dateIdx = rawHeaders.findIndex(h => (h.includes('日期') || h.includes('date')) && !h.includes('戳記'));
        const amtIdx = rawHeaders.findIndex(h => h.includes('金額') || h.includes('amount') || h.includes('花費'));
        const noteIdx = rawHeaders.findIndex(h => h.includes('備註') || h.includes('說明') || h.includes('note'));

        for (let i = data.length - 1; i >= 1; i--) {
          const row = data[i];
          const d = normalizeDate(dateIdx !== -1 ? row[dateIdx] : row[0]);
          const rawAmt = amtIdx !== -1 ? row[amtIdx] : 0;
          const a = Number(String(rawAmt).replace(/[$,NT\s]/gi, '').replace(/,/g, '')) || 0;
          const n = noteIdx !== -1 ? String(row[noteIdx] || '').trim() : '';

          if (d === targetDate && a === targetAmount && (!targetNote || n.includes(targetNote) || targetNote.includes(n))) {
            matchedRow = i + 1;
            break;
          }
        }
      }

      if (matchedRow !== -1) {
        const statusCol = getStatusColumnIndex(sheet);
        // 將狀態改為已刪除
        sheet.getRange(matchedRow, statusCol).setValue(`🔴 已刪除 (${timestamp})`);
        // 為該列文字劃上刪除線
        sheet.getRange(matchedRow, 1, 1, Math.min(statusCol, 6)).setFontLine('line-through');
        
        ensureDashboard(sheet);
        return responseJSON({
          success: true,
          deleted: true,
          row: matchedRow,
          message: `已成功將第 ${matchedRow} 列標記為「已刪除」！`
        });
      }

      return responseJSON({
        success: false,
        deleted: false,
        message: '未在試算表中找到匹配的消費列'
      });
    }

    // -------------------------------------------------------------
    // 操作 B：新增消費紀錄 (標記為「🟢 正常記帳」)
    // -------------------------------------------------------------
    if (action === 'add' || !action) {
      const date = body.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const category = body.category || '食';
      const amount = Number(body.amount) || 0;
      const note = body.note || '';

      // 新增一列：[時間戳記, 消費日期, 消費類別, 金額, 備註說明, 狀態]
      const newRowNumber = sheet.getLastRow() + 1;
      sheet.appendRow([timestamp, date, category, amount, note, `🟢 正常 (${timestamp})`]);
      
      // 確保樣式無刪除線
      sheet.getRange(newRowNumber, 1, 1, 6).setFontLine('none');

      // 更新即時統計看板
      ensureDashboard(sheet);

      return responseJSON({
        success: true,
        message: '記帳成功！已寫入試算表並更新總金額統計。',
        data: { date: date, category: category, amount: amount, note: note }
      });
    }

    // -------------------------------------------------------------
    // 操作 C：清空試算表所有消費紀錄 (保留標題與統計看板)
    // -------------------------------------------------------------
    if (action === 'clear_all') {
      const lastRow = sheet.getLastRow();
      let clearedCount = 0;
      if (lastRow >= 2) {
        clearedCount = lastRow - 1;
        sheet.deleteRows(2, lastRow - 1);
      }

      ensureHeadersAndStatus(sheet);
      ensureDashboard(sheet);

      return responseJSON({
        success: true,
        cleared: true,
        clearedCount: clearedCount,
        message: `已成功清空 Google 試算表共 ${clearedCount} 筆資料！`
      });
    }
  } catch (error) {
    return responseJSON({ success: false, error: error.toString() });
  }
}

// -----------------------------------------------------------------
// 輔助函式：確保標題列具備「狀態」欄位
// -----------------------------------------------------------------
function ensureHeadersAndStatus(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['時間戳記', '消費日期', '消費類別', '金額', '備註說明', '狀態 (新增/刪除)']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#E2E8F0');
    return;
  }

  const firstRow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 6)).getValues()[0];
  const hasStatus = firstRow.some(h => String(h).includes('狀態'));
  if (!hasStatus) {
    // 在第 6 欄補上狀態標題
    sheet.getRange(1, 6).setValue('狀態 (新增/刪除)').setFontWeight('bold').setBackground('#E2E8F0');
  }
}

function getStatusColumnIndex(sheet) {
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 6)).getValues()[0];
  for (let c = 0; c < headers.length; c++) {
    if (String(headers[c]).includes('狀態')) return c + 1;
  }
  return 6;
}

// -----------------------------------------------------------------
// 輔助函式：在試算表右側建立「📊 即時消費統計看板」
// -----------------------------------------------------------------
function ensureDashboard(sheet) {
  try {
    // 看板設置於 H 欄與 I 欄 (第 8, 9 欄)
    const startCol = 8;
    
    // 檢查看板是否已經建立過標題
    const titleVal = sheet.getRange(1, startCol).getValue();
    if (titleVal === '📊 消費即時統計看板') {
      return; // 已建立，公式會自動動態計算，無需重複覆寫
    }

    // 建立看板標題
    sheet.getRange(1, startCol, 1, 2)
      .merge()
      .setValue('📊 消費即時統計看板')
      .setFontWeight('bold')
      .setFontSize(11)
      .setBackground('#10B981')
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('center');

    // 看板項目與動態公式 (只統計狀態未被刪除的有效項目)
    // 假設 D 欄為金額、C 欄為類別、F 欄為狀態
    const dashboardItems = [
      ['💰 累積總支出', '=SUMIF(F2:F, "<>*已刪除*", D2:D)'],
      ['📝 有效總筆數', '=COUNTIF(F2:F, "<>*已刪除*")'],
      ['🍜 食 (飲食三餐)', '=SUMIFS(D2:D, C2:C, "*食*", F2:F, "<>*已刪除*")'],
      ['👕 衣 (服飾鞋包)', '=SUMIFS(D2:D, C2:C, "*衣*", F2:F, "<>*已刪除*")'],
      ['🏠 住 (房租水電)', '=SUMIFS(D2:D, C2:C, "*住*", F2:F, "<>*已刪除*")'],
      ['🚗 行 (交通通勤)', '=SUMIFS(D2:D, C2:C, "*行*", F2:F, "<>*已刪除*")'],
      ['📚 育 (教育進修)', '=SUMIFS(D2:D, C2:C, "*育*", F2:F, "<>*已刪除*")'],
      ['🎮 樂 (休閒玩樂)', '=SUMIFS(D2:D, C2:C, "*樂*", F2:F, "<>*已刪除*")']
    ];

    for (let r = 0; r < dashboardItems.length; r++) {
      const rowIdx = r + 2;
      sheet.getRange(rowIdx, startCol).setValue(dashboardItems[r][0]).setFontWeight('bold').setBackground('#F8FAFC');
      sheet.getRange(rowIdx, startCol + 1).setFormula(dashboardItems[r][1]).setFontWeight('bold').setFontColor('#0F172A').setHorizontalAlignment('right');
    }

    // 設定總金額格式為貨幣或千分位
    sheet.getRange(2, startCol + 1).setNumberFormat('"NT$"#,##0').setFontColor('#059669').setFontSize(12);
    // 食衣住行育樂數字格式
    sheet.getRange(4, startCol + 1, 6, 1).setNumberFormat('"NT$"#,##0');

    // 邊框修飾
    sheet.getRange(1, startCol, dashboardItems.length + 1, 2).setBorder(true, true, true, true, true, true, '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID);
  } catch (err) {
    // 避免看板建立錯誤影響主流程
  }
}

// -----------------------------------------------------------------
// 智慧選擇目標工作表 (優先選取表單回應分頁)
// -----------------------------------------------------------------
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
