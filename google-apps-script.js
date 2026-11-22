/**
 * ====================================================================
 * Google 試算表後端程式碼 (Google Apps Script) - 穩定修正版
 * ====================================================================
 * 
 * 【本版修正】
 * - 移除動態統計看板（防止看板資料被誤判為消費紀錄、防止清空後產生幽靈 0 元資料）
 * - 讀取時嚴格只讀 A~F 欄「有金額」的列
 * - 清空操作不再呼叫 ensureDashboard，避免殘留 H 欄干擾 getLastRow()
 * 
 * 【更新步驟】
 * 1. 打開 Google 試算表 -> 「擴充功能」 -> 「Apps Script」，整份覆蓋儲存。
 * 2. 點選「部署」-> 「管理部署作業」-> 鉛筆圖示 -> 版本選「新版本」-> 部署。
 */

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getTargetSheet(ss);

    // 只讀 A~F 欄，不用 getDataRange()，避免 H、I 欄看板被誤抓
    const lastDataRow = getLastDataRow(sheet);
    if (lastDataRow <= 1) {
      return responseJSON({
        success: true,
        sheetName: sheet.getName(),
        recordsCount: 0,
        data: []
      });
    }

    const data = sheet.getRange(1, 1, lastDataRow, 6).getValues();
    const rawHeaders = data[0].map(h => String(h).trim());
    const headers = rawHeaders.map(h => h.toLowerCase());

    let dateIdx = headers.findIndex(h =>
      (h.includes('消費日') || h.includes('日期') || h.includes('date')) && !h.includes('戳記')
    );
    if (dateIdx === -1) {
      dateIdx = headers.findIndex(h => h.includes('時間') || h.includes('戳記') || h.includes('time'));
    }
    const catIdx  = headers.findIndex(h => h.includes('類別') || h.includes('分類') || h.includes('category') || h.includes('項目'));
    const amtIdx  = headers.findIndex(h => h.includes('金額') || h.includes('amount') || h.includes('花費') || h.includes('費用'));
    const noteIdx = headers.findIndex(h => h.includes('備註') || h.includes('說明') || h.includes('note') || h.includes('品項') || h.includes('內容'));
    const statusIdx = headers.findIndex(h => h.includes('狀態') || h.includes('status'));

    const records = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // 若 A~E 欄全空，跳過
      if (row.slice(0, 5).every(cell => cell === '' || cell === null)) continue;

      // 跳過已刪除列
      const statusText = statusIdx !== -1 ? String(row[statusIdx] || '') : '';
      if (statusText.includes('刪除')) continue;

      // 解析金額，0 元直接跳過
      const rawAmt = amtIdx !== -1 ? row[amtIdx] : 0;
      const amount = Number(String(rawAmt).replace(/[$,NT\s]/gi, '')) || 0;
      if (amount <= 0) continue;

      const rawDate = dateIdx !== -1 ? row[dateIdx] : row[0];
      const date = normalizeDate(rawDate);
      if (!date || !/^\d{4}-\d{2}-\d{2}$/.test(date)) continue;

      const rawCat = catIdx !== -1 ? String(row[catIdx] || '').trim() : '';
      const category = normalizeCategory(rawCat);
      const note = noteIdx !== -1 ? String(row[noteIdx] || '').trim() : '';

      records.push({
        id: 'gs_' + (i + 1),
        rowIndex: i + 1,
        date: date,
        category: category,
        amount: amount,
        note: note || `${category}類消費`
      });
    }

    return responseJSON({
      success: true,
      sheetName: sheet.getName(),
      recordsCount: records.length,
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
    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    // ----------------------------------------------------------------
    // 操作 A：新增
    // ----------------------------------------------------------------
    if (action === 'add') {
      ensureHeaders(sheet);
      const date     = body.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const category = body.category || '食';
      const amount   = Number(body.amount) || 0;
      const note     = body.note || '';

      if (amount <= 0) {
        return responseJSON({ success: false, message: '金額必須大於 0' });
      }

      const newRow = sheet.getLastRow() + 1;
      sheet.appendRow([ts, date, category, amount, note, '🟢 正常']);
      sheet.getRange(newRow, 1, 1, 6).setFontLine('none');

      return responseJSON({ success: true, message: '記帳成功！' });
    }

    // ----------------------------------------------------------------
    // 操作 B：標記刪除（加刪除線，不真正刪列）
    // ----------------------------------------------------------------
    if (action === 'delete') {
      const targetRow    = Number(body.rowIndex);
      const targetDate   = body.date;
      const targetAmount = Number(body.amount);
      const targetNote   = String(body.note || '').trim();

      const lastDataRow = getLastDataRow(sheet);
      if (lastDataRow <= 1) {
        return responseJSON({ success: false, message: '試算表無資料' });
      }

      const data = sheet.getRange(1, 1, lastDataRow, 6).getValues();
      let matchedRow = -1;

      if (targetRow >= 2 && targetRow <= lastDataRow) {
        matchedRow = targetRow;
      } else {
        const h = data[0].map(h => String(h).trim().toLowerCase());
        const dIdx = h.findIndex(x => (x.includes('日期') || x.includes('date')) && !x.includes('戳記'));
        const aIdx = h.findIndex(x => x.includes('金額') || x.includes('amount'));
        const nIdx = h.findIndex(x => x.includes('備註') || x.includes('說明') || x.includes('note'));

        for (let i = data.length - 1; i >= 1; i--) {
          const row = data[i];
          const d = normalizeDate(dIdx !== -1 ? row[dIdx] : row[0]);
          const a = Number(String(aIdx !== -1 ? row[aIdx] : 0).replace(/[$,NT\s]/gi, '')) || 0;
          const n = String(nIdx !== -1 ? row[nIdx] : '').trim();
          if (d === targetDate && a === targetAmount && (!targetNote || n === targetNote || n.includes(targetNote))) {
            matchedRow = i + 1;
            break;
          }
        }
      }

      if (matchedRow !== -1) {
        const statusCol = getStatusCol(sheet);
        sheet.getRange(matchedRow, statusCol).setValue(`🔴 已刪除 (${ts})`);
        sheet.getRange(matchedRow, 1, 1, Math.min(statusCol, 6)).setFontLine('line-through');
        return responseJSON({ success: true, deleted: true, row: matchedRow });
      }
      return responseJSON({ success: false, deleted: false, message: '找不到匹配資料' });
    }

    // ----------------------------------------------------------------
    // 操作 C：清空所有資料列（只刪 A~F 欄的資料，保留標題）
    // ----------------------------------------------------------------
    if (action === 'clear_all') {
      const lastDataRow = getLastDataRow(sheet);
      let clearedCount = 0;
      if (lastDataRow >= 2) {
        clearedCount = lastDataRow - 1;
        sheet.deleteRows(2, lastDataRow - 1);
      }
      return responseJSON({
        success: true,
        cleared: true,
        clearedCount: clearedCount,
        message: `已成功清空 ${clearedCount} 筆資料！`
      });
    }

    return responseJSON({ success: false, message: '未知的操作：' + action });
  } catch (error) {
    return responseJSON({ success: false, error: error.toString() });
  }
}

// 只計算 A~F 欄有內容的最後一列（排除 H、I 欄看板的干擾）
function getLastDataRow(sheet) {
  const lastCol = Math.min(sheet.getLastColumn(), 6);
  if (lastCol === 0) return 0;
  const colAValues = sheet.getRange(1, 1, sheet.getMaxRows(), 1).getValues();
  let lastRow = 0;
  for (let i = 0; i < colAValues.length; i++) {
    if (colAValues[i][0] !== '' && colAValues[i][0] !== null) {
      lastRow = i + 1;
    }
  }
  return lastRow;
}

function ensureHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['時間戳記', '消費日期', '消費類別', '金額', '備註說明', '狀態']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#E2E8F0');
  }
}

function getStatusCol(sheet) {
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 6)).getValues()[0];
  for (let c = 0; c < headers.length; c++) {
    if (String(headers[c]).includes('狀態')) return c + 1;
  }
  return 6;
}

function getTargetSheet(ss) {
  const sheets = ss.getSheets();
  if (sheets.length === 1) return sheets[0];
  for (const s of sheets) {
    if (s.getName().includes('表單回應') || s.getName().toLowerCase().includes('form responses')) return s;
  }
  return sheets.reduce((a, b) => (a.getLastRow() >= b.getLastRow() ? a : b));
}

function normalizeDate(raw) {
  if (!raw) return '';
  if (raw instanceof Date) {
    return Utilities.formatDate(raw, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const str = String(raw).trim();
  const m = str.match(/(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);
  if (m) {
    const y = parseInt(m[1]), mo = parseInt(m[2]), d = parseInt(m[3]);
    if (y >= 2000 && mo >= 1 && mo <= 12 && d >= 1 && d <= 31) {
      return `${y}-${String(mo).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
    }
  }
  return '';
}

function normalizeCategory(raw) {
  if (!raw) return '食';
  const s = String(raw).trim();
  if (s.includes('食') || s.includes('餐') || s.includes('飲')) return '食';
  if (s.includes('衣') || s.includes('服') || s.includes('穿')) return '衣';
  if (s.includes('住') || s.includes('房') || s.includes('水電') || s.includes('瓦斯')) return '住';
  if (s.includes('行') || s.includes('車') || s.includes('交通') || s.includes('油')) return '行';
  if (s.includes('育') || s.includes('書') || s.includes('學') || s.includes('課')) return '育';
  if (s.includes('樂') || s.includes('玩') || s.includes('遊') || s.includes('休閒')) return '樂';
  return '食';
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
