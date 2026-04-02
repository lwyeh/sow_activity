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
  var row = [id, data.name, now];
  
  if (data.id) {
    var idx = findRowIndex(sh, 0, data.id);
    if (idx > 0) {
      sh.getRange(idx, 8, 1, 1).setNumberFormat('@'); // 保留原有邏輯：設定第8欄為純文字
      sh.getRange(idx, 1, 1, row.length).setValues([row]);
      return { success: true, id: id };
    }
  }
  var newRow = sh.getLastRow() + 1;
  sh.getRange(newRow, 1, 1, row.length).setValues([row]);
  sh.getRange(newRow, 8, 1, 1).setNumberFormat('@');
  sh.getRange(newRow, 8, 1, 1).setValue(data.phone || '');
  return { success: true, id: id };
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

function getMembers() {
  var sh   = getSheet(SHEET_MEMBERS);
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).filter(function(r){ return r[0]; }).map(function(r){
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
      squad:       r[12]
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

function saveMember(data) {
  var sh  = getSheet(SHEET_MEMBERS);
  var id  = data.id ? String(data.id) : genId('M');
  var now = new Date();
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
    now
  ];

  if (data.id) {
    var idx = findRowIndex(sh, 0, data.id);
    if (idx > 0) {
      sh.getRange(idx, 8, 1, 1).setNumberFormat('@'); // 設定電話欄位為純文字
      sh.getRange(idx, 1, 1, row.length).setValues([row]);
      return { success: true, id: id };
    }
  }
  var newRow = sh.getLastRow() + 1;
  sh.getRange(newRow, 1, 1, row.length).setValues([row]);
  sh.getRange(newRow, 8, 1, 1).setNumberFormat('@');
  sh.getRange(newRow, 8, 1, 1).setValue(data.phone || '');
  return { success: true, id: id };
}

function deleteMember(memberId) {
  var sh  = getSheet(SHEET_MEMBERS);
  var idx = findRowIndex(sh, 0, memberId);
  if (idx > 0) sh.deleteRow(idx);
  return { success: true };
}

function dataCleanup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
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