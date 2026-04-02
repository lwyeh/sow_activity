// ================================================================
//  Main.gs - 系統常數、入口與初始化設定
// ================================================================

// ── 系統常數 (資料表名稱) ─────────────────────────────────────────
var SHEET_MEMBERS     = '成員資料';
var SHEET_FAMILIES    = '家庭資料';
var SHEET_ACTIVITIES  = '活動列表';
var SHEET_RSVP        = '報名紀錄';
var SHEET_PAY_ACTS    = '收費活動';
var SHEET_PAY_EQUIPS  = '收費設備';   // ★ 新增
var SHEET_PAY_DETAILS = '收費明細';
var SHEET_PAY_RECORDS = '繳費紀錄';

var ADMIN_PASSWORD    = 'admin1234';

// ── 系統常數 (選單選項) ──────────────────────────────────────────
var TROOP_LIST   = ['蟻','蜂','鹿','鷹','育','複式'];
var GRADE_LIST   = ['一歲','兩歲','三歲','四歲','五歲','小班','中班','大班','小一','小二','小三','小四','小五','小六','國一','國二','國三','高一','高二','高三','成人'];
var ROLE_LIST    = ['複式團長','會長','蟻育副會長','蜂育副會長','鹿育副會長','鷹育副會長','手作組組長','財務組組長','資訊組組長','文書組組長','值星官','財務','安心營營長','安心營','副團長','團長','季總召導引員','活動總召','助理導引員','導引員','小小蟻','小蟻','小蜂','小鹿','小鷹','蟻蜂育','鹿育','鷹育','離團'];
var SQUAD_LIST   = ['花叢小隊','天空小隊','草原小隊','大地小隊','草原鹿','森林鹿','高地鹿','湖泊鹿','泥壺蜂','虎頭蜂','長腳蜂','細腰蜂','小黑蟻','小黃蟻','小綠蟻','小紅蟻','無'];
var GENDER_LIST  = ['男','女','其他'];

var TROOP_SQUADS = {
  '蟻': ['小黑蟻','小黃蟻','小綠蟻','小紅蟻'],
  '蜂': ['泥壺蜂','虎頭蜂','長腳蜂','細腰蜂'],
  '鹿': ['草原鹿','森林鹿','高地鹿','湖泊鹿'],
  '鷹': [],
  '育': ['花叢小隊','天空小隊','草原小隊','大地小隊'],
  '複式': []
};

// ── Entry Point (網頁入口) ───────────────────────────────────────
function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('荒野親子團報名系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width,initial-scale=1,maximum-scale=5,user-scalable=yes');
}

// ── Sheet Setup (初始化資料表) ───────────────────────────────────
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  function ensureSheet(name, headers) {
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.appendRow(headers);
      sh.getRange(1, 1, 1, headers.length)
        .setBackground('#1a3a2a')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    return sh;
  }

  var famSh = ensureSheet(SHEET_FAMILIES, ['家庭ID','家庭名稱','建立時間','家庭編號']);
  var famData = famSh.getDataRange().getValues();
  for (var i = 1; i < famData.length; i++) {
    if (!famData[i][3]) {
      famSh.getRange(i+1, 4).setValue(String(i).padStart(3,'0'));
    }
  }

  // ★ 收費活動：新增 type 欄
  ensureSheet(SHEET_PAY_ACTS, ['活動ID','活動名稱','說明','類型','建立時間']);

  // ★ 收費設備（全新資料表）
  ensureSheet(SHEET_PAY_EQUIPS, ['設備ID','活動ID','設備名稱','金額','數量上限','種類']);

  // ★ 收費明細：重新設計欄位
  ensureSheet(SHEET_PAY_DETAILS, ['明細ID','活動ID','成員ID','家庭ID','設備ID','設備名稱','數量','單價','小計']);

  // ★ 繳費紀錄：簡化為家庭+活動維度
  ensureSheet(SHEET_PAY_RECORDS, [
    '紀錄ID','活動ID','家庭ID',
    '繳費金額','繳費方式','備註','狀態','退回原因',
    '送出時間','確認時間','建立時間'
  ]);

  ensureSheet(SHEET_MEMBERS, [
    '成員ID','家庭ID','家庭名稱','姓名','自然名','性別','角色',
    '電話','Email','團別','年級','職位','隊名','建立時間'
  ]);
  ensureSheet(SHEET_ACTIVITIES, [
    '活動ID','活動名稱','開始日期','結束日期','截止報名日期','狀態','建立時間'
  ]);
  ensureSheet(SHEET_RSVP, [
    '報名ID','活動ID','活動名稱','成員ID','成員姓名','家庭ID','家庭名稱',
    '出席狀態','備註','報名時間','最後更新'
  ]);

  return { success: true };
}

// =================================================================
// 1. 小工具：產生密碼的 MD5 雜湊值 (供開發者手動新增管理員使用)
// =================================================================
function generatePasswordHashTool() {
  // 👉 在這裡填入您想設定的密碼
  var rawPassword = "test123"; 
  
  // 計算 MD5 雜湊值
  var signature = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, rawPassword, Utilities.Charset.UTF_8);
  
  // 將 Byte 陣列轉換成 16 進位字串
  var hash = signature.map(function(byte) {
    var v = (byte < 0) ? 256 + byte : byte;
    return ("0" + v.toString(16)).slice(-2);
  }).join("");
  
  Logger.log("===============================");
  Logger.log("原始密碼: " + rawPassword);
  Logger.log("MD5 雜湊: " + hash);
  Logger.log("===============================");
  Logger.log("請將上方的 MD5 雜湊值複製，並貼到「管理員帳號」工作表的 B 欄中。");
}

// =================================================================
// 2. 核心函式：後端 MD5 轉換器 (供登入比對使用)
// =================================================================
function getMD5Hash_(text) {
  var signature = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text, Utilities.Charset.UTF_8);
  return signature.map(function(byte) {
    var v = (byte < 0) ? 256 + byte : byte;
    return ("0" + v.toString(16)).slice(-2);
  }).join("");
}

// =================================================================
// 3. API：管理員登入驗證
// =================================================================
function verifyAdminLogin(username, password) {
  var sh = getSheet('管理員帳號');
  if (!sh) return { success: false, msg: '系統未設定管理員帳號表' };

  var data = sh.getDataRange().getValues();
  
  // 將使用者輸入的密碼，轉換成 MD5 雜湊值
  var inputHash = getMD5Hash_(password);

  // 從第 2 列開始比對 (避開標題)
  for (var i = 1; i < data.length; i++) {
    var sheetUser = String(data[i][0]).trim();
    var sheetHash = String(data[i][1]).trim();
    
    // 比對帳號與雜湊值是否完全一致
    if (sheetUser === String(username).trim() && sheetHash === inputHash) {
      return { success: true };
    }
  }
  
  // 跑完迴圈沒找到，代表帳密錯誤
  return { success: false, msg: '帳號或密碼錯誤' };
}

// 執行此函式可自動修復或建立管理員帳號表
function fixAdminSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = '管理員帳號';
  var sh = ss.getSheetByName(sheetName);
  
  if (!sh) {
    // 如果找不到，就建立一個
    sh = ss.insertSheet(sheetName);
    sh.appendRow(['帳號', '密碼雜湊']);
    sh.getRange("A1:B1").setFontWeight("bold");
    sh.setFrozenRows(1);
    Logger.log("已為您自動建立「管理員帳號」工作表！");
  } else {
    Logger.log("工作表已存在，請確認名稱是否包含多餘空格。");
    // 強制重整一次名稱，移除可能的隱藏空格
    sh.setName(sheetName);
  }
}

// ── 終極除錯照妖鏡：測試 getAllData 是否正常運作 ──
function testGetAllData() {
  try {
    var data = getAllData();
    Logger.log("【測試成功】資料打包沒有問題！");
    Logger.log("家庭資料數: " + (data.families ? data.families.length : "抓不到"));
    Logger.log("成員資料數: " + (data.members ? data.members.length : "抓不到"));
    Logger.log("活動資料數: " + (data.activities ? data.activities.length : "抓不到"));
    Logger.log("系統選項數: " + (Object.keys(data.options).length));
  } catch(e) {
    Logger.log("❌ 【測試失敗】程式碼當機了！");
    Logger.log("錯誤原因: " + e.message);
    Logger.log("錯誤發生在第幾行: " + e.lineNumber);
  }
}