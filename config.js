/**
 * ====================================================================
 * 全域預設設定檔 (config.js)
 * ====================================================================
 * 
 * 💡 為什麼需要這個檔案？
 * 當您把網頁上傳到 GitHub Pages 後，若只透過網頁介面輸入設定，設定只會保存在
 * 該瀏覽器的「本機快取 (LocalStorage)」中。一旦換了手機、平板或清除快取，就必須重新輸入。
 * 
 * 👉 最佳解決方案：
 * 直接將您的 Google Apps Script 網址填入下方的 gasUrl 中，
 * 上傳到 GitHub 後，無論在任何手機、電腦或任何瀏覽器打開，都會「自動連線」，再也不需要重複設定！
 */

window.DEFAULT_CONFIG = {
  // 請將您的 Google Apps Script Web App URL 填入下方引號中：
  // 例如: "https://script.google.com/macros/s/AKfycbxxxxxxxxx/exec"
  gasUrl: "",

  // 若使用 Google 表單單向提交模式，可在下方填寫：
  formUrl: "",
  formEntries: {
    date: "",
    category: "",
    amount: "",
    note: ""
  }
};
