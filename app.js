/**
 * ====================================================================
 * 日常記帳本 - 主應用程式核心邏輯 (app.js)
 * ====================================================================
 */

// 1. 六大分類設定
const CATEGORIES = {
  '食': { color: '#F97316', icon: 'fa-utensils', label: '食 (飲食三餐)', lightBg: '#FFEDD5' },
  '衣': { color: '#8B5CF6', icon: 'fa-shirt', label: '衣 (服飾鞋包)', lightBg: '#EDE9FE' },
  '住': { color: '#0284C7', icon: 'fa-house', label: '住 (房租水電)', lightBg: '#E0F2FE' },
  '行': { color: '#10B981', icon: 'fa-car', label: '行 (交通通勤)', lightBg: '#D1FAE5' },
  '育': { color: '#F59E0B', icon: 'fa-book-open', label: '育 (教育進修)', lightBg: '#FEF3C7' },
  '樂': { color: '#EC4899', icon: 'fa-gamepad', label: '樂 (休閒玩樂)', lightBg: '#FCE7F3' }
};

const CATEGORY_KEYS = ['食', '衣', '住', '行', '育', '樂'];

// 2. 本地儲存 Key
const STORAGE_DATA_KEY = 'EXPENSE_TRACKER_DATA_V1';
const STORAGE_CONFIG_KEY = 'EXPENSE_TRACKER_CONFIG_V1';
const STORAGE_INIT_KEY = 'EXPENSE_TRACKER_INITIALIZED_V1';
const STORAGE_DELETED_KEY = 'EXPENSE_TRACKER_DELETED_KEYS_V1';

// 3. 全域狀態
let expenses = [];
let chartInstance = null;
let chartType = 'doughnut'; // 'doughnut' or 'pie'
let selectedCategory = '食';
let deletedKeys = new Set(); // 已刪除紀錄黑名單 (防止重新整理時拉回)

// 當前檢視的年月 (YYYY-MM)
let currentYearMonth = getTodayString().slice(0, 7);

// Google 雲端設定
let cloudConfig = {
  mode: 'gas', // 'gas' or 'form'
  gasUrl: '',
  formUrl: '',
  formEntries: {
    date: '',
    category: '',
    amount: '',
    note: ''
  }
};

// ====================================================================
// 初始化與 DOM 監聽
// ====================================================================
document.addEventListener('DOMContentLoaded', () => {
  // 清除舊版 V1 的壞資料 key（防止舊 key 殘留）
  localStorage.removeItem('EXPENSE_TRACKER_DATA_V1');
  localStorage.removeItem('EXPENSE_TRACKER_INITIALIZED_V1');

  loadConfig();
  loadDeletedKeys();
  loadData();
  initDateInputs();
  initCategorySelector();
  initMonthNavigation();
  initChart();
  initForm();
  initHistoryControls();
  initModal();
  updateSyncBadge();
  renderAll();

  // 若有設定 Google Apps Script，網頁開啟時自動在背景同步最新的雲端試算表資料
  if (cloudConfig.mode === 'gas' && cloudConfig.gasUrl) {
    fetchCloudData(true);
  } else {
    // 只有在第一次造訪（從未初始化過）且未串接雲端時，才載入預設示範資料
    const hasInitialized = localStorage.getItem(STORAGE_INIT_KEY);
    if (!hasInitialized) {
      localStorage.setItem(STORAGE_INIT_KEY, 'true');
      if (expenses.length === 0) {
        loadDemoData(false);
      }
    }
  }
});

// ====================================================================
// 日期工具函式
// ====================================================================
function getTodayString() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function formatYearMonth(ym) {
  const [year, month] = ym.split('-');
  return `${year} 年 ${parseInt(month, 10)} 月`;
}

function initDateInputs() {
  const dateInput = document.getElementById('expense-date');
  if (dateInput) {
    dateInput.value = getTodayString();
  }

  const monthPicker = document.getElementById('month-picker');
  if (monthPicker) {
    monthPicker.value = currentYearMonth;
  }
}

// ====================================================================
// 類別選擇按鈕邏輯
// ====================================================================
function initCategorySelector() {
  const buttons = document.querySelectorAll('.category-btn');
  buttons.forEach(btn => {
    btn.addEventListener('click', () => {
      buttons.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      selectedCategory = btn.getAttribute('data-cat');
    });
  });

  // 快速金額按鈕 (+50, +100, etc.)
  const amtInput = document.getElementById('expense-amount');
  const quickBtns = document.querySelectorAll('.quick-amt-btn');
  quickBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const addVal = parseInt(btn.getAttribute('data-add'), 10) || 0;
      const currentVal = parseInt(amtInput.value, 10) || 0;
      amtInput.value = currentVal + addVal;
      amtInput.focus();
    });
  });
}

// ====================================================================
// 月份切換控制
// ====================================================================
function initMonthNavigation() {
  const prevBtn = document.getElementById('btn-prev-month');
  const nextBtn = document.getElementById('btn-next-month');
  const thisMonthBtn = document.getElementById('btn-this-month');
  const monthTrigger = document.getElementById('month-picker-trigger');
  const monthPicker = document.getElementById('month-picker');

  prevBtn.addEventListener('click', () => changeMonth(-1));
  nextBtn.addEventListener('click', () => changeMonth(1));

  thisMonthBtn.addEventListener('click', () => {
    currentYearMonth = getTodayString().slice(0, 7);
    monthPicker.value = currentYearMonth;
    renderAll();
  });

  monthTrigger.addEventListener('click', () => {
    monthPicker.showPicker ? monthPicker.showPicker() : monthPicker.focus();
  });

  monthPicker.addEventListener('change', (e) => {
    if (e.target.value) {
      currentYearMonth = e.target.value;
      renderAll();
    }
  });
}

function changeMonth(delta) {
  const [yearStr, monthStr] = currentYearMonth.split('-');
  let year = parseInt(yearStr, 10);
  let month = parseInt(monthStr, 10) + delta;

  if (month < 1) {
    month = 12;
    year -= 1;
  } else if (month > 12) {
    month = 1;
    year += 1;
  }

  currentYearMonth = `${year}-${String(month).padStart(2, '0')}`;
  document.getElementById('month-picker').value = currentYearMonth;
  renderAll();
}

// ====================================================================
// 表單送出與記帳
// ====================================================================
function initForm() {
  const form = document.getElementById('expense-form');
  const amtInput = document.getElementById('expense-amount');
  const noteInput = document.getElementById('expense-note');
  const dateInput = document.getElementById('expense-date');

  form.addEventListener('submit', async (e) => {
    e.preventDefault();

    const amount = parseInt(amtInput.value, 10);
    if (!amount || amount <= 0) {
      showToast('請輸入有效的消費金額！', 'error');
      amtInput.focus();
      return;
    }

    const dateVal = dateInput.value || getTodayString();
    const noteVal = noteInput.value.trim() || `${selectedCategory}類消費`;

    const newExpense = {
      id: 'local_' + Date.now(),
      date: dateVal,
      category: selectedCategory,
      amount: amount,
      note: noteVal,
      createdAt: new Date().toISOString()
    };

    // 新增至本地
    expenses.unshift(newExpense);
    saveData();

    // 如果該筆資料的月份就是當前正在看的月份，或切換到該月
    const expMonth = dateVal.slice(0, 7);
    if (expMonth !== currentYearMonth) {
      currentYearMonth = expMonth;
      document.getElementById('month-picker').value = currentYearMonth;
    }

    // 重設表單輸入
    amtInput.value = '';
    noteInput.value = '';
    amtInput.focus();

    // 更新畫面
    renderAll();
    showToast(`已成功記錄 NT$ ${amount.toLocaleString()}！`, 'success');

    // 背景同步至 Google 雲端
    syncToCloud(newExpense);
  });
}

// ====================================================================
// Google 雲端寫入 (Apps Script / Form Response)
// ====================================================================
async function syncToCloud(item) {
  if (cloudConfig.mode === 'gas' && cloudConfig.gasUrl) {
    try {
      // Google Apps Script Web App POST
      fetch(cloudConfig.gasUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'text/plain;charset=utf-8' // GAS 建議用 text/plain 避免 CORS 預檢失敗
        },
        body: JSON.stringify({
          date: item.date,
          category: item.category,
          amount: item.amount,
          note: item.note
        })
      }).then(res => res.json()).then(result => {
        if (result && result.success) {
          showToast('☁️ 已即時同步寫入 Google 試算表！', 'success');
        }
      }).catch(err => {
        console.warn('雲端同步提示:', err);
      });
    } catch (err) {
      console.warn('GAS POST 失敗:', err);
    }
  } else if (cloudConfig.mode === 'form' && cloudConfig.formUrl) {
    try {
      const formData = new FormData();
      if (cloudConfig.formEntries.date) formData.append(cloudConfig.formEntries.date, item.date);
      if (cloudConfig.formEntries.category) formData.append(cloudConfig.formEntries.category, item.category);
      if (cloudConfig.formEntries.amount) formData.append(cloudConfig.formEntries.amount, item.amount);
      if (cloudConfig.formEntries.note) formData.append(cloudConfig.formEntries.note, item.note);

      // mode: 'no-cors' 允許跨域送出 Google Form
      fetch(cloudConfig.formUrl, {
        method: 'POST',
        mode: 'no-cors',
        body: formData
      }).then(() => {
        showToast('☁️ 已成功提交至 Google 表單！', 'success');
      }).catch(err => {
        console.warn('Google Form 送出錯誤:', err);
      });
    } catch (err) {
      console.warn('Google Form 提交失敗:', err);
    }
  }
}

// ====================================================================
// 從 Google Apps Script 讀取全量歷史資料 (雙向同步)
// ====================================================================
async function fetchCloudData(isSilent = false) {
  if (!cloudConfig.gasUrl) {
    if (!isSilent) {
      showToast('尚未設定 Google Apps Script 網址！', 'error');
      openSettingsModal();
    }
    return;
  }

  const syncBtn = document.getElementById('btn-sync-cloud');
  if (syncBtn) syncBtn.classList.add('rotating');

  if (!isSilent) {
    showToast('正在從 Google 試算表同步資料中...', 'info');
  }

  try {
    // 加入時間戳記破壞快取，避免手機或瀏覽器讀取到舊的快取內容
    const separator = cloudConfig.gasUrl.includes('?') ? '&' : '?';
    const noCacheUrl = `${cloudConfig.gasUrl}${separator}_t=${Date.now()}`;

    const res = await fetch(noCacheUrl);
    const result = await res.json();

    if (result && result.success && Array.isArray(result.data)) {
      if (result.data.length === 0) {
        if (!isSilent) {
          showToast(`已連線到分頁「${result.sheetName || '試算表'}」，目前尚無資料`, 'info');
        }
        return;
      }

      // 合併雲端資料（自動排除本地已刪除的紀錄、無效日期、金額為 0 的空白列）
      const validRows = result.data.filter(item => {
        // 排除金額 <= 0 的無效列（統計看板被誤抓時產生的 0 元空資料）
        if ((Number(item.amount) || 0) <= 0) return false;

        // 排除日期異常（包含 NaN、Sun 等非 YYYY-MM-DD 格式）
        const cleanDate = normalizeDateClient(item.date);
        if (!cleanDate || cleanDate.includes('NaN') || !/^\d{4}-\d{2}-\d{2}$/.test(cleanDate)) return false;

        // 排除本地已刪除的黑名單
        const cleanCat = normalizeCategoryClient(item.category);
        const itemKey = `${cleanDate}_${cleanCat}_${Number(item.amount) || 0}_${(item.note || '').trim()}`;
        return !deletedKeys.has(itemKey) && !deletedKeys.has(item.id);
      });

      expenses = validRows.map(item => {
        const cleanCat = normalizeCategoryClient(item.category);
        const cleanDate = normalizeDateClient(item.date);
        return {
          id: item.id || ('gs_' + Math.random().toString(36).substr(2, 9)),
          rowIndex: item.rowIndex || null,
          date: cleanDate,
          category: cleanCat,
          amount: Number(item.amount) || 0,
          note: item.note || `${cleanCat}類消費`
        };
      });

      localStorage.setItem(STORAGE_INIT_KEY, 'true');
      saveData();

      // 檢查是否需要自動切換到有資料的月份
      const currentMonthHasData = expenses.some(i => i.date.startsWith(currentYearMonth));
      if (!currentMonthHasData && expenses.length > 0) {
        // 取得最新一筆資料的年月
        const latestDate = expenses.map(i => i.date).sort().reverse()[0];
        if (latestDate) {
          currentYearMonth = latestDate.slice(0, 7);
          const monthPicker = document.getElementById('month-picker');
          if (monthPicker) monthPicker.value = currentYearMonth;
        }
      }

      renderAll();
      if (!isSilent) {
        const sheetInfo = result.sheetName ? ` [${result.sheetName}]` : '';
        showToast(`🎉 成功同步 ${result.data.length} 筆資料${sheetInfo}！`, 'success');
      }
    } else {
      if (!isSilent) showToast('讀取失敗：' + (result.error || '回傳格式異常'), 'error');
    }
  } catch (err) {
    console.error('GAS 讀取錯誤:', err);
    if (!isSilent) {
      showToast('連線失敗，請確認 Web App 是否已設定為「任何人」可存取。', 'error');
    }
  } finally {
    if (syncBtn) syncBtn.classList.remove('rotating');
  }
}

// 前端輔助：標準化分類名稱
function normalizeCategoryClient(raw) {
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

// 前端輔助：確保日期字串為嚴格的 YYYY-MM-DD，絕不產生 Sun 或 NaN
function normalizeDateClient(raw) {
  if (!raw) return getTodayString();
  const str = String(raw).trim();

  // 優先比對 YYYY-MM-DD 或 YYYY/MM/DD 或 YYYY.MM.DD 格式
  const match = str.match(/(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);
  if (match) {
    const y = parseInt(match[1], 10);
    const m = parseInt(match[2], 10);
    const d = parseInt(match[3], 10);
    // 年份合理性驗證
    if (y >= 2000 && y <= 2100 && m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      return `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    }
  }

  // 嘗試 JS Date 解析，只接受合理年份
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime()) && parsed.getFullYear() >= 2000) {
    const y = parsed.getFullYear();
    const m = String(parsed.getMonth() + 1).padStart(2, '0');
    const d = String(parsed.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  // 無法解析時回傳今天，避免顯示 Sun/NaN
  return getTodayString();
}

// ====================================================================
// 圓餅圖 (Chart.js)
// ====================================================================
function initChart() {
  const ctx = document.getElementById('expensePieChart').getContext('2d');
  
  const initialData = {
    labels: CATEGORY_KEYS,
    datasets: [{
      data: [0, 0, 0, 0, 0, 0],
      backgroundColor: CATEGORY_KEYS.map(k => CATEGORIES[k].color),
      hoverBackgroundColor: CATEGORY_KEYS.map(k => CATEGORIES[k].color),
      borderColor: '#FFFFFF',
      borderWidth: 2,
      hoverOffset: 8
    }]
  };

  chartInstance = new Chart(ctx, {
    type: chartType,
    data: initialData,
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: false // 自訂右側清單圖例
        },
        tooltip: {
          callbacks: {
            label: function(context) {
              const label = context.label || '';
              const value = context.parsed || 0;
              const total = context.dataset.data.reduce((a, b) => a + b, 0);
              const percentage = total > 0 ? ((value / total) * 100).toFixed(1) : '0';
              return ` ${label}：NT$ ${value.toLocaleString()} (${percentage}%)`;
            }
          }
        }
      },
      cutout: chartType === 'doughnut' ? '65%' : '0%'
    }
  });

  // 圖表類型切換按鈕
  const btnDoughnut = document.getElementById('btn-chart-doughnut');
  const btnPie = document.getElementById('btn-chart-pie');

  btnDoughnut.addEventListener('click', () => {
    chartType = 'doughnut';
    btnDoughnut.classList.add('active');
    btnPie.classList.remove('active');
    if (chartInstance) {
      chartInstance.destroy();
      initChart();
      updateChart(getFilteredExpenses(currentYearMonth));
    }
  });

  btnPie.addEventListener('click', () => {
    chartType = 'pie';
    btnPie.classList.add('active');
    btnDoughnut.classList.remove('active');
    if (chartInstance) {
      chartInstance.destroy();
      initChart();
      updateChart(getFilteredExpenses(currentYearMonth));
    }
  });
}

function updateChart(monthExpenses) {
  if (!chartInstance) return;

  const categoryTotals = calculateCategoryTotals(monthExpenses);
  const data = CATEGORY_KEYS.map(k => categoryTotals[k]);
  const total = data.reduce((a, b) => a + b, 0);

  const emptyNotice = document.getElementById('empty-chart-notice');
  if (total === 0) {
    emptyNotice.classList.remove('hidden');
    chartInstance.data.datasets[0].data = [0, 0, 0, 0, 0, 0];
  } else {
    emptyNotice.classList.add('hidden');
    chartInstance.data.datasets[0].data = data;
  }

  chartInstance.update();
  renderCategoryBreakdown(categoryTotals, total);
}

// 渲染右側六大分類進度條與佔比
function renderCategoryBreakdown(categoryTotals, totalMonthExpense) {
  const container = document.getElementById('category-breakdown');
  container.innerHTML = '';

  CATEGORY_KEYS.forEach(key => {
    const config = CATEGORIES[key];
    const amount = categoryTotals[key] || 0;
    const percent = totalMonthExpense > 0 ? ((amount / totalMonthExpense) * 100).toFixed(1) : 0;

    const row = document.createElement('div');
    row.className = 'cat-row';
    row.innerHTML = `
      <div class="cat-row-top">
        <div class="cat-title-badge">
          <span class="cat-indicator-dot" style="background-color: ${config.color}"></span>
          <span>${config.label}</span>
        </div>
        <div class="cat-numbers">
          <span class="cat-amt">NT$ ${amount.toLocaleString()}</span>
          <span class="cat-percent">${percent}%</span>
        </div>
      </div>
      <div class="cat-bar-bg">
        <div class="cat-bar-fill" style="width: ${percent}%; background-color: ${config.color};"></div>
      </div>
    `;
    container.appendChild(row);
  });
}

// ====================================================================
// 數據計算與畫面整體渲染
// ====================================================================
function getFilteredExpenses(yearMonth) {
  return expenses.filter(item => {
    return item.date && item.date.startsWith(yearMonth);
  });
}

function calculateCategoryTotals(expenseList) {
  const totals = { '食': 0, '衣': 0, '住': 0, '行': 0, '育': 0, '樂': 0 };
  expenseList.forEach(item => {
    if (totals[item.category] !== undefined) {
      totals[item.category] += Number(item.amount) || 0;
    }
  });
  return totals;
}

function renderAll() {
  const monthTitle = formatYearMonth(currentYearMonth);
  document.getElementById('current-month-text').textContent = monthTitle;
  document.getElementById('month-metric-title').textContent = `${monthTitle} 總支出`;
  document.getElementById('history-month-title').textContent = monthTitle;

  const monthExpenses = getFilteredExpenses(currentYearMonth);
  const totalAmount = monthExpenses.reduce((sum, item) => sum + (Number(item.amount) || 0), 0);
  const count = monthExpenses.length;

  // 1. 更新卡片指標
  document.getElementById('total-monthly-expense').textContent = totalAmount.toLocaleString();
  document.getElementById('monthly-count').textContent = count;

  // 計算最高花費類別
  const totals = calculateCategoryTotals(monthExpenses);
  let maxCat = '-';
  let maxAmt = 0;
  for (const cat of CATEGORY_KEYS) {
    if (totals[cat] > maxAmt) {
      maxAmt = totals[cat];
      maxCat = `${cat} (NT$ ${maxAmt.toLocaleString()})`;
    }
  }
  document.getElementById('top-category').textContent = maxCat;

  // 2. 更新圓餅圖與分類統計
  updateChart(monthExpenses);

  // 3. 渲染歷史列表
  renderHistoryList();
}

// ====================================================================
// 消費明細列表渲染
// ====================================================================
function initHistoryControls() {
  const searchInput = document.getElementById('history-search');
  const filterCat = document.getElementById('filter-category');

  searchInput.addEventListener('input', () => renderHistoryList());
  filterCat.addEventListener('change', () => renderHistoryList());
}

function renderHistoryList() {
  const container = document.getElementById('expense-list');
  const emptyNotice = document.getElementById('empty-history-notice');
  const countBadge = document.getElementById('history-count-badge');
  const searchKeyword = document.getElementById('history-search').value.trim().toLowerCase();
  const catFilter = document.getElementById('filter-category').value;

  let list = getFilteredExpenses(currentYearMonth);

  // 類別過濾
  if (catFilter !== 'ALL') {
    list = list.filter(item => item.category === catFilter);
  }

  // 搜尋關鍵字過濾
  if (searchKeyword) {
    list = list.filter(item => {
      const matchNote = (item.note || '').toLowerCase().includes(searchKeyword);
      const matchAmt = String(item.amount).includes(searchKeyword);
      const matchCat = item.category.includes(searchKeyword);
      return matchNote || matchAmt || matchCat;
    });
  }

  // 依日期降序排列
  list.sort((a, b) => new Date(b.date) - new Date(a.date));

  countBadge.textContent = `${list.length} 筆`;
  container.innerHTML = '';

  if (list.length === 0) {
    emptyNotice.classList.remove('hidden');
    return;
  }

  emptyNotice.classList.add('hidden');

  list.forEach(item => {
    const catConfig = CATEGORIES[item.category] || CATEGORIES['食'];
    const el = document.createElement('div');
    el.className = 'expense-item';
    el.innerHTML = `
      <div class="item-left">
        <div class="item-cat-icon" style="background-color: ${catConfig.lightBg}; color: ${catConfig.color};">
          <i class="fa-solid ${catConfig.icon}"></i>
        </div>
        <div class="item-details">
          <h4>${escapeHTML(item.note || catConfig.label)}</h4>
          <div class="item-meta">
            <span><i class="fa-regular fa-calendar"></i> ${item.date}</span>
            <span class="item-cat-badge" style="background-color: ${catConfig.lightBg}; color: ${catConfig.color};">
              ${item.category}
            </span>
          </div>
        </div>
      </div>
      <div class="item-right">
        <div class="item-amount">- NT$ ${(Number(item.amount) || 0).toLocaleString()}</div>
        <button class="item-delete-btn" data-id="${item.id}" title="刪除這筆消費">
          <i class="fa-regular fa-trash-can"></i>
        </button>
      </div>
    `;

    el.querySelector('.item-delete-btn').addEventListener('click', (e) => {
      e.stopPropagation();
      deleteExpense(item.id);
    });

    container.appendChild(el);
  });
}

function deleteExpense(id) {
  if (!confirm('確定要刪除這筆記帳紀錄嗎？')) return;

  const targetIndex = expenses.findIndex(i => i.id === id);
  if (targetIndex === -1) return;
  const targetItem = expenses[targetIndex];

  // 1. 加入已刪除黑名單 (保證重新整理時絕不會再次抓回)
  const itemKey = getItemKey(targetItem);
  deletedKeys.add(itemKey);
  if (targetItem.id) deletedKeys.add(targetItem.id);
  saveDeletedKeys();

  // 2. 本機陣列移除並儲存
  expenses.splice(targetIndex, 1);
  localStorage.setItem(STORAGE_INIT_KEY, 'true');
  saveData();
  renderAll();
  showToast('已刪除該筆消費紀錄', 'info');

  // 3. 若有連線 Google 試算表，發送同步刪除通知
  if (cloudConfig.mode === 'gas' && cloudConfig.gasUrl) {
    try {
      fetch(cloudConfig.gasUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({
          action: 'delete',
          rowIndex: targetItem.rowIndex || null,
          date: targetItem.date,
          category: targetItem.category,
          amount: targetItem.amount,
          note: targetItem.note
        })
      }).then(res => res.json()).then(result => {
        if (result && result.success && result.deleted) {
          showToast('☁️ 已同步從 Google 試算表移除該列！', 'success');
        }
      }).catch(err => {
        console.warn('雲端刪除通知失敗:', err);
      });
    } catch (e) {
      console.warn('發送刪除請求錯誤:', e);
    }
  }
}

// 輔助函式：產生唯一特徵 Key
function getItemKey(item) {
  if (!item) return '';
  const d = normalizeDateClient(item.date);
  const c = normalizeCategoryClient(item.category);
  const a = Number(item.amount) || 0;
  const n = (item.note || '').trim();
  return `${d}_${c}_${a}_${n}`;
}

function saveDeletedKeys() {
  try {
    localStorage.setItem(STORAGE_DELETED_KEY, JSON.stringify(Array.from(deletedKeys)));
  } catch (e) {
    console.error('儲存刪除清單失敗:', e);
  }
}

function loadDeletedKeys() {
  try {
    const raw = localStorage.getItem(STORAGE_DELETED_KEY);
    if (raw) {
      const arr = JSON.parse(raw);
      deletedKeys = new Set(Array.isArray(arr) ? arr : []);
    }
  } catch (e) {
    deletedKeys = new Set();
  }
}

// ====================================================================
// 設定 Modal 邏輯
// ====================================================================
function initModal() {
  const modal = document.getElementById('settings-modal');
  const btnOpen = document.getElementById('btn-open-settings');
  const btnClose = document.getElementById('btn-close-settings');
  const btnCancel = document.getElementById('btn-cancel-settings');
  const btnSave = document.getElementById('btn-save-settings');
  const btnSync = document.getElementById('btn-sync-cloud');
  const btnTestGas = document.getElementById('btn-test-gas');
  const btnViewCode = document.getElementById('btn-view-script-code');
  const btnCopyCode = document.getElementById('btn-copy-code');
  const tabButtons = document.querySelectorAll('.tab-btn');

  btnOpen.addEventListener('click', openSettingsModal);
  btnClose.addEventListener('click', closeSettingsModal);
  btnCancel.addEventListener('click', closeSettingsModal);
  btnSync.addEventListener('click', fetchCloudData);

  modal.addEventListener('click', (e) => {
    if (e.target === modal) closeSettingsModal();
  });

  // 分頁切換
  tabButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      tabButtons.forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      btn.classList.add('active');
      const tabId = btn.getAttribute('data-tab');
      document.getElementById(tabId).classList.add('active');
      cloudConfig.mode = tabId === 'tab-gas' ? 'gas' : 'form';
    });
  });

  // 儲存設定
  btnSave.addEventListener('click', () => {
    cloudConfig.gasUrl = document.getElementById('gas-webapp-url').value.trim();
    cloudConfig.formUrl = document.getElementById('form-action-url').value.trim();
    cloudConfig.formEntries = {
      date: document.getElementById('entry-date').value.trim(),
      category: document.getElementById('entry-category').value.trim(),
      amount: document.getElementById('entry-amount').value.trim(),
      note: document.getElementById('entry-note').value.trim()
    };

    saveConfig();
    updateSyncBadge();
    closeSettingsModal();
    showToast('雲端設定已成功儲存！', 'success');
  });

  // 測試連線
  btnTestGas.addEventListener('click', async () => {
    cloudConfig.gasUrl = document.getElementById('gas-webapp-url').value.trim();
    saveConfig();
    await fetchCloudData();
  });

  // 查看 GAS 程式碼
  btnViewCode.addEventListener('click', async () => {
    const container = document.getElementById('script-code-container');
    const codeDisplay = document.getElementById('gas-code-display');
    if (container.classList.contains('hidden')) {
      container.classList.remove('hidden');
      try {
        const res = await fetch('google-apps-script.js');
        const code = await res.text();
        codeDisplay.textContent = code;
      } catch (e) {
        codeDisplay.textContent = '// 無法讀取檔案，請直接開啟 google-apps-script.js 複製';
      }
    } else {
      container.classList.add('hidden');
    }
  });

  // 複製 GAS 程式碼
  btnCopyCode.addEventListener('click', () => {
    const code = document.getElementById('gas-code-display').textContent;
    navigator.clipboard.writeText(code).then(() => {
      showToast('已複製 Google Apps Script 程式碼！', 'success');
    });
  });

  // 本機資料管理按鈕
  document.getElementById('btn-load-demo').addEventListener('click', () => {
    loadDemoData(true);
    closeSettingsModal();
  });

  document.getElementById('btn-export-json').addEventListener('click', () => {
    exportDataAsJSON();
  });

  // 清空 Google 雲端試算表所有資料
  const btnClearCloud = document.getElementById('btn-clear-cloud-sheet');
  if (btnClearCloud) {
    btnClearCloud.addEventListener('click', async () => {
      if (!cloudConfig.gasUrl) {
        showToast('尚未設定 Google Apps Script 網址，無法操作雲端試算表！', 'error');
        return;
      }

      const confirmFirst = confirm('⚠️ 確定要清空 Google 試算表中的「所有消費紀錄」嗎？\n\n此操作將會清空試算表第 2 列之後的所有資料，此動作無法復原！');
      if (!confirmFirst) return;

      const confirmDouble = confirm('🚨 再次確認：真的要清空 Google 雲端試算表全部資料嗎？\n（確定後將同步歸零總金額看板與本機資料）');
      if (!confirmDouble) return;

      showToast('正在請求清空 Google 試算表資料中...', 'info');

      try {
        const res = await fetch(cloudConfig.gasUrl, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify({ action: 'clear_all' })
        });
        const result = await res.json();

        if (result && result.success) {
          // 清空本地陣列與已刪除黑名單
          expenses = [];
          deletedKeys.clear();
          localStorage.setItem(STORAGE_INIT_KEY, 'true');
          saveData();
          saveDeletedKeys();
          renderAll();
          closeSettingsModal();
          showToast(`🎉 ${result.message || '已成功清空 Google 試算表全部資料！'}`, 'success');
        } else {
          showToast('清空失敗：' + (result.error || result.message || '未知錯誤'), 'error');
        }
      } catch (err) {
        console.error('清空試算表失敗:', err);
        showToast('清空請求失敗，請確認 Apps Script 權限與連線！', 'error');
      }
    });
  }

  document.getElementById('btn-clear-all').addEventListener('click', () => {
    if (confirm('警告：確定要清空所有記帳資料嗎？此操作無法還原！')) {
      // 將當前所有紀錄加入已刪除清單
      expenses.forEach(item => {
        deletedKeys.add(getItemKey(item));
        if (item.id) deletedKeys.add(item.id);
      });
      saveDeletedKeys();
      expenses = [];
      localStorage.setItem(STORAGE_INIT_KEY, 'true');
      saveData();
      renderAll();
      closeSettingsModal();
      showToast('已清空所有記帳資料', 'info');
    }
  });
}

function openSettingsModal() {
  document.getElementById('gas-webapp-url').value = cloudConfig.gasUrl || '';
  document.getElementById('form-action-url').value = cloudConfig.formUrl || '';
  document.getElementById('entry-date').value = cloudConfig.formEntries?.date || '';
  document.getElementById('entry-category').value = cloudConfig.formEntries?.category || '';
  document.getElementById('entry-amount').value = cloudConfig.formEntries?.amount || '';
  document.getElementById('entry-note').value = cloudConfig.formEntries?.note || '';

  // 預設 tab
  const gasTab = document.querySelector('[data-tab="tab-gas"]');
  const formTab = document.querySelector('[data-tab="tab-form"]');
  if (cloudConfig.mode === 'form') {
    formTab.click();
  } else {
    gasTab.click();
  }

  document.getElementById('settings-modal').classList.remove('hidden');
}

function closeSettingsModal() {
  document.getElementById('settings-modal').classList.add('hidden');
}

function updateSyncBadge() {
  const badge = document.getElementById('sync-status-badge');
  const isCloud = (cloudConfig.mode === 'gas' && cloudConfig.gasUrl) ||
                  (cloudConfig.mode === 'form' && cloudConfig.formUrl);

  if (isCloud) {
    badge.className = 'badge badge-cloud';
    badge.querySelector('.status-text').textContent =
      cloudConfig.mode === 'gas' ? 'Google 試算表連線中' : 'Google 表單已綁定';
  } else {
    badge.className = 'badge badge-local';
    badge.querySelector('.status-text').textContent = '本機儲存模式';
  }
}

// ====================================================================
// 示範資料載入器
// ====================================================================
function loadDemoData(isUserTriggered = false) {
  const currentYM = currentYearMonth; // 採用當前選取之年月
  const demoList = [
    { id: 'demo_1', date: `${currentYM}-02`, category: '食', amount: 150, note: '早餐美而美牛肉堡套餐' },
    { id: 'demo_2', date: `${currentYM}-03`, category: '行', amount: 1200, note: '北捷定期月票 TPASS' },
    { id: 'demo_3', date: `${currentYM}-05`, category: '食', amount: 320, note: '義大利麵商業午餐' },
    { id: 'demo_4', date: `${currentYM}-08`, category: '衣', amount: 1980, note: 'UNIQLO 保暖外套與長褲' },
    { id: 'demo_5', date: `${currentYM}-10`, category: '住', amount: 1250, note: '台電夏季電費分攤' },
    { id: 'demo_6', date: `${currentYM}-12`, category: '育', amount: 650, note: '博客來理財與程式設計書籍' },
    { id: 'demo_7', date: `${currentYM}-15`, category: '樂', amount: 890, note: '威秀影城 IMAX 電影與爆米花' },
    { id: 'demo_8', date: `${currentYM}-18`, category: '食', amount: 280, note: '晚餐日式拉麵加叉燒' },
    { id: 'demo_9', date: `${currentYM}-20`, category: '行', amount: 450, note: '中油自助加油 95 無鉛' },
    { id: 'demo_10', date: `${currentYM}-22`, category: '樂', amount: 390, note: '任天堂 Switch 特價遊戲' }
  ];

  expenses = demoList;
  saveData();
  renderAll();

  if (isUserTriggered) {
    showToast('已成功載入食衣住行育樂示範資料！', 'success');
  }
}

// 匯出 JSON
function exportDataAsJSON() {
  const dataStr = 'data:text/json;charset=utf-8,' + encodeURIComponent(JSON.stringify(expenses, null, 2));
  const downloadAnchor = document.createElement('a');
  downloadAnchor.setAttribute('href', dataStr);
  downloadAnchor.setAttribute('download', `記帳本備份_${currentYearMonth}.json`);
  document.body.appendChild(downloadAnchor);
  downloadAnchor.click();
  downloadAnchor.remove();
}

// ====================================================================
// 本地資料儲存 / 載入
// ====================================================================
function saveData() {
  try {
    localStorage.setItem(STORAGE_DATA_KEY, JSON.stringify(expenses));
  } catch (e) {
    console.error('儲存至 LocalStorage 失敗:', e);
  }
}

function loadData() {
  try {
    const raw = localStorage.getItem(STORAGE_DATA_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      // 讀取時過濾掉無效紀錄（金額 ≤ 0 或日期無效）
      expenses = parsed.filter(item => {
        const amount = Number(item.amount) || 0;
        if (amount <= 0) return false;
        const date = normalizeDateClient(item.date);
        if (!date || date.includes('NaN') || !/^\d{4}-\d{2}-\d{2}$/.test(date)) return false;
        return true;
      }).map(item => ({
        ...item,
        date: normalizeDateClient(item.date)
      }));
      // 若過濾後資料有變動，自動寫回 localStorage 清掉壞資料
      if (expenses.length !== parsed.length) {
        console.info(`已清除 ${parsed.length - expenses.length} 筆無效本機資料`);
        saveData();
      }
    }
  } catch (e) {
    console.error('讀取 LocalStorage 失敗:', e);
    expenses = [];
  }
}

function saveConfig() {
  try {
    localStorage.setItem(STORAGE_CONFIG_KEY, JSON.stringify(cloudConfig));
  } catch (e) {
    console.error('儲存設定失敗:', e);
  }
}

function loadConfig() {
  // 1. 若有 config.js 預設設定，先載入預設值
  if (window.DEFAULT_CONFIG) {
    if (window.DEFAULT_CONFIG.gasUrl) cloudConfig.gasUrl = window.DEFAULT_CONFIG.gasUrl;
    if (window.DEFAULT_CONFIG.formUrl) cloudConfig.formUrl = window.DEFAULT_CONFIG.formUrl;
    if (window.DEFAULT_CONFIG.formEntries) {
      cloudConfig.formEntries = Object.assign(cloudConfig.formEntries, window.DEFAULT_CONFIG.formEntries);
    }
  }

  // 2. 嘗試從 LocalStorage 讀取（使用者若在前端介面手動覆蓋設定）
  try {
    const raw = localStorage.getItem(STORAGE_CONFIG_KEY);
    if (raw) {
      const stored = JSON.parse(raw);
      if (stored.gasUrl) cloudConfig.gasUrl = stored.gasUrl;
      if (stored.formUrl) cloudConfig.formUrl = stored.formUrl;
      if (stored.mode) cloudConfig.mode = stored.mode;
      if (stored.formEntries) {
        cloudConfig.formEntries = Object.assign(cloudConfig.formEntries, stored.formEntries);
      }
    }
  } catch (e) {
    console.error('讀取設定失敗:', e);
  }

  // 3. 支援 URL 參數快速設定 (例如: https://xxx.github.io/?gas=https://script.google.com/...)
  try {
    const params = new URLSearchParams(window.location.search);
    const gasParam = params.get('gas');
    if (gasParam) {
      cloudConfig.gasUrl = decodeURIComponent(gasParam);
      cloudConfig.mode = 'gas';
      saveConfig();
    }
  } catch (e) {
    // 忽略 URL 參數讀取異常
  }
}

// ====================================================================
// Toast 提示通知
// ====================================================================
let toastTimer = null;
function showToast(msg, type = 'info') {
  const toast = document.getElementById('toast');
  const msgEl = toast.querySelector('.toast-message');
  const iconEl = toast.querySelector('.toast-icon');

  msgEl.textContent = msg;
  toast.className = `toast toast-${type}`;

  if (type === 'success') {
    iconEl.className = 'toast-icon fa-solid fa-circle-check';
  } else if (type === 'error') {
    iconEl.className = 'toast-icon fa-solid fa-triangle-exclamation';
  } else {
    iconEl.className = 'toast-icon fa-solid fa-circle-info';
  }

  toast.classList.remove('hidden');

  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    toast.classList.add('hidden');
  }, 3500);
}

// 輔助函式：跳脫 HTML
function escapeHTML(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
