// ================================================================
//  Api_Members.gs - 成員與家庭資料 API
// ================================================================

// ── FAMILIES API (家庭管理) ──────────────────────────────────────

function getFamilies() {
  var sh   = getSheet(SHEET_FAMILIES);
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).filter(function(r){ return r[0]; }).map(function(r){
    return { 
      id:   String(r[0]), 
      name: r[1], 
      no:   r[3] ? String(r[3]) : '' 
    };
  });
}

function saveFamily(data) {
  var sh  = getSheet(SHEET_FAMILIES);
  var id  = data.id ? String(data.id) : genId('F');
  var now = new Date();

  if (data.id) {
    var idx = findRowIndex(sh, 0, data.id);
    if (idx > 0) {
      // 若是更新現有家庭，只寫入前三欄，保留原有的第 4 欄 (家庭編號)
      var row = [id, data.name, now];
      sh.getRange(idx, 1, 1, row.length).setValues([row]);
      return { success: true, id: id };
    }
  }

  // ★ 修復 3：若是新增家庭，自動依照列數產生 3 位數的「家庭編號」
  var newRow = sh.getLastRow() + 1;
  var newNo = String(newRow - 1).padStart(3, '0');
  var row = [id, data.name, now, newNo];
  sh.getRange(newRow, 1, 1, row.length).setValues([row]);

  return { success: true, id: id, no: newNo };
}

function deleteFamily(familyId) {
  var sh  = getSheet(SHEET_FAMILIES);
  var idx = findRowIndex(sh, 0, familyId);
  if (idx > 0) sh.deleteRow(idx);
  return { success: true };
}

function initFamilyNumbers() {
  // 為沒有編號的家庭補齊編號
  var sh   = getSheet(SHEET_FAMILIES);
  var data = sh.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (!data[i][3]) {
      sh.getRange(i+1, 4).setValue(String(i).padStart(3,'0'));
      count++;
    }
  }
  return { success: true, count: count };
}

function repairMissingFamilies() {
  // 補齊：成員裡有 familyId 但家庭資料表沒有的家庭
  var families = getFamilies();
  var members  = getMembers();
  var famIds   = {};
  families.forEach(function(f){ famIds[f.id] = true; });
  var added = {};
  members.forEach(function(m){
    if (!famIds[m.familyId] && !added[m.familyId]) {
      added[m.familyId] = true;
      saveFamily({ id: m.familyId, name: m.familyName || m.familyId });
    }
  });
  return { repaired: Object.keys(added).length };
}


// ── MEMBERS API (成員管理) ───────────────────────────────────────

// ── 1. 取得成員資料 (加上防呆與新欄位) ──
function getMembers() {
  var sh   = getSheet(SHEET_MEMBERS);
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1).filter(function(r){ return r[0]; }).map(function(r){
    // 日期防呆，確保傳給前端的是字串
    var bDate = r[15] instanceof Date ? Utilities.formatDate(r[15], "GMT+8", "yyyy/MM/dd") : String(r[15] || '');
    var jDate = r[16] instanceof Date ? Utilities.formatDate(r[16], "GMT+8", "yyyy/MM/dd") : String(r[16] || '');

    return {
      id:          String(r[0]),
      familyId:    String(r[1]),
      familyName:  r[2],
      name:        r[3],
      naturalName: r[4],
      gender:      r[5],
      role:        r[6],
      phone:       r[7] || r[7] === 0 ? String(r[7]) : '',
      email:       r[8],
      troop:       r[9],
      grade:       r[10],
      position:    r[11],
      squad:       r[12],
      // r[13] 是建立時間
      memberNo:    String(r[14] || ''), // 第15欄
      birthDate:   bDate,               // 第16欄
      joinDate:    jDate,               // 第17欄
      remark:      String(r[17] || '')  // ★ 個人備註
    };
  });
}

function getMembersByFamily(familyId) {
  return getMembers().filter(function(m){ return String(m.familyId) === String(familyId); });
}

function getMembersBySquad(squad) {
  return getMembers().filter(function(m){ return m.squad === squad; });
}

function getMembersByTroop(troop) {
  return getMembers().filter(function(m){ return m.troop === troop; });
}

// ── 2. 儲存成員資料 (支援新欄位存檔) ──
function saveMember(data) {
  var sh  = getSheet(SHEET_MEMBERS);
  var id  = data.id ? String(data.id) : genId('M');
  var now = new Date();
  
  // 日期防呆
  var bDate = data.birthDate instanceof Date ? Utilities.formatDate(data.birthDate, "GMT+8", "yyyy/MM/dd") : String(data.birthDate || '');
  var jDate = data.joinDate instanceof Date ? Utilities.formatDate(data.joinDate, "GMT+8", "yyyy/MM/dd") : String(data.joinDate || '');

  var row = [
    id,
    String(data.familyId || ''),
    data.familyName  || '',
    data.name        || '',
    data.naturalName || '',
    data.gender      || '',
    data.role        || '',
    data.phone       || '',
    data.email       || '',
    data.troop       || '',
    data.grade       || '',
    data.position    || '',
    data.squad       || '',
    now,
    String(data.memberNo || ''),
    bDate,
    jDate,
    String(data.remark || '')   // ★ 個人備註
  ];

  if (data.id) {
    var idx = findRowIndex(sh, 0, data.id);
    if (idx > 0) {
      sh.getRange(idx, 8, 1, 1).setNumberFormat('@'); // 電話純文字
      sh.getRange(idx, 1, 1, row.length).setValues([row]);
      return { success: true, id: id };
    }
  }
  var newRow = sh.getLastRow() + 1;
  sh.getRange(newRow, 1, 1, row.length).setValues([row]);
  sh.getRange(newRow, 8, 1, 1).setNumberFormat('@'); 
  return { success: true, id: id };
}

function deleteMember(memberId) {
  var sh  = getSheet(SHEET_MEMBERS);
  var idx = findRowIndex(sh, 0, memberId);
  if (idx > 0) sh.deleteRow(idx);
  return { success: true };
}

function dataCleanup() {
  var ss = getDb();
  var sheets = [SHEET_MEMBERS, SHEET_FAMILIES];
  
  sheets.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    var range = sh.getDataRange();
    var values = range.getValues();
    
    var cleaned = values.map(function(row) {
      return row.map(function(cell) {
        // 去除字串前後空格、換行符號
        return (typeof cell === 'string') ? cell.trim().replace(/\n|\r/g, "") : cell;
      });
    });
    range.setValues(cleaned);
  });
  console.log("資料清理完成");
}