// ================================================================
//  Helpers.gs - 系統共用工具與輔助函式
// ================================================================

// ── HTML 拆分檔案專用 Helper ────────────────────────────────────
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── Spreadsheet Helpers ──────────────────────────────────────────
function getSheet(name) {
  var ss = getDb();
  var sh = ss.getSheetByName(name);
  if (!sh) { 
    setupSheets();
    sh = ss.getSheetByName(name); 
  }
  return sh;
}

function findRowIndex(sh, colIndex, value) {
  var vals = sh.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][colIndex]) === String(value)) return i + 1;
  }
  return -1;
}

// ── Date & Formatting Helpers ────────────────────────────────────
function formatDate(d) {
  if (!d) return '';
  var dt = new Date(d);
  if (isNaN(dt.getTime())) return '';
  return dt.getFullYear() + '/' +
    ('0'+(dt.getMonth()+1)).slice(-2) + '/' +
    ('0'+dt.getDate()).slice(-2);
}

function formatDateTime(d) {
  if (!d) return '';
  var dt = new Date(d);
  if (isNaN(dt.getTime())) return '';
  return formatDate(dt) + ' ' + ('0'+dt.getHours()).slice(-2) + ':' + ('0'+dt.getMinutes()).slice(-2);
}

// ── ID Generator ─────────────────────────────────────────────────
function genId(prefix) {
  return prefix + Date.now() + Math.random().toString(36).slice(2,6).toUpperCase();
}

// ================================================================
//  系統環境設定
// ================================================================
// ★ 將這裡換成您試算表的真實 ID
var DATABASE_SHEET_ID = '1MLSP6VpPhzqFuey6BUQDk1YTJ2uuPc9Lhyad5mqJOPA'; 

// 建立一個取代 getActiveSpreadsheet() 的連線函式
function getDb() {
  return SpreadsheetApp.openById(DATABASE_SHEET_ID);
}
