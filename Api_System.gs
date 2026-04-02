// ================================================================
//  Api_System.gs - 系統設定、升級與匯出 API
// ================================================================

// ── META API (全域選單與資料初始化) ───────────────────────────

function getOptions() {
  return {
    troops:  TROOP_LIST,
    grades:  GRADE_LIST,
    roles:   ROLE_LIST,
    squads:  SQUAD_LIST,
    genders: GENDER_LIST
  };
}

function getAllData() {
  // 一次性讀取前端所需的基礎資料 (跨檔案呼叫 Api_Members 與 Api_Activities)
  return {
    families:   getFamilies(),
    members:    getMembers(),
    activities: getActivities(),
    options:    getOptions()
  };
}

function verifyAdmin(password) {
  return { success: password === ADMIN_PASSWORD };
}

// ── 年度升級 (UPGRADE API) ────────────────────────────────────

function getUpgradePreview() {
  var members = getMembers();
  var GRADE_MAP = {
    '中班': { grade: '大班', position: '小蟻' },
    '大班': { grade: '小一', position: '小蟻' },
    '小一': { grade: '小二', position: '小蟻' },
    '小二': { grade: '小三', position: '小蜂' },
    '小三': { grade: '小四', position: '小蜂' },
    '小四': { grade: '小五', position: '小蜂' },
    '小五': { grade: '小六', position: '小鹿' },
    '小六': { grade: '國一', position: '小鹿' },
    '國一': { grade: '國二', position: '小鹿' },
    '國二': { grade: '國三', position: '小鷹' },
    '國三': { grade: '高一', position: '小鷹' },
    '高一': { grade: '高二', position: '小鷹' },
    '高二': { grade: '高三', position: '離團' },
  };
  
  var preview = [];
  members.forEach(function(m) {
    if (m.position === '離團') return;
    if (m.role === '家長') return;
    var map = GRADE_MAP[m.grade];
    if (!map) return;
    preview.push({
      id:          m.id,
      name:        m.name,
      naturalName: m.naturalName || '',
      oldGrade:    m.grade,
      newGrade:    map.grade,
      oldPosition: m.position,
      newPosition: map.position
    });
  });
  return preview;
}

function executeUpgrade() {
  var members = getMembers();
  var sh = getSheet(SHEET_MEMBERS);
  var GRADE_MAP = {
    '中班': { grade: '大班', position: '小蟻' },
    '大班': { grade: '小一', position: '小蟻' },
    '小一': { grade: '小二', position: '小蟻' },
    '小二': { grade: '小三', position: '小蜂' },
    '小三': { grade: '小四', position: '小蜂' },
    '小四': { grade: '小五', position: '小蜂' },
    '小五': { grade: '小六', position: '小鹿' },
    '小六': { grade: '國一', position: '小鹿' },
    '國一': { grade: '國二', position: '小鹿' },
    '國二': { grade: '國三', position: '小鷹' },
    '國三': { grade: '高一', position: '小鷹' },
    '高一': { grade: '高二', position: '小鷹' },
    '高二': { grade: '高三', position: '離團' },
  };
  
  var count = 0;
  var all = sh.getDataRange().getValues();
  members.forEach(function(m) {
    if (m.position === '離團') return;
    if (m.role === '家長') return;
    var map = GRADE_MAP[m.grade];
    if (!map) return;
    var idx = findRowIndex(sh, 0, m.id);
    if (idx < 0) return;
    // 年級在第11欄(index 10)，職位在第12欄(index 11)
    sh.getRange(idx, 11, 1, 1).setValue(map.grade);
    sh.getRange(idx, 12, 1, 1).setValue(map.position);
    count++;
  });
  return { success: true, count: count };
}

// ── CSV EXPORT (匯出報名名單) ─────────────────────────────────

function exportRsvpCsv(activityId) {
  var records = getRsvpByActivity(activityId);
  var members = getMembers();
  var memberMap = {};
  members.forEach(function(m){ memberMap[m.id] = m; });

  var headers = ['活動名稱','家庭名稱','姓名','自然名','性別','團別','小隊','年級','職位','出席狀態','備註','報名時間'];
  var rows = records.map(function(r) {
    var m = memberMap[r.memberId] || {};
    return [r.activityName, r.familyName, r.memberName, m.naturalName||'', m.gender||'', m.troop||'', m.squad||'', m.grade||'', m.position||'', r.status, r.note, r.rsvpTime];
  });
  
  var csv = [headers].concat(rows).map(function(row){
    return row.map(function(c){ return '"' + String(c).replace(/"/g,'""') + '"'; }).join(',');
  }).join('\n');

  return { success: true, csv: csv };
}