// ================================================================
//  Api_Activities.gs - 活動與報名統計 API
// ================================================================

// ── ACTIVITIES API (活動管理) ───────────────────────────────────
// ================================================================
// Api_Activities.gs - 活動與報名統計 API
// ================================================================

// ── 狀態推導工具函式 ────────────────────────────────────────────
function calcActivityStatus(openDate, deadline) {
  var now = new Date();
  now.setHours(0, 0, 0, 0);

  if (deadline && deadline < now) {
    return { status: '已結束',   subLabel: '已截止報名',   isOpen: false };
  }
  if (openDate && openDate > now) {
    return { status: '暫停報名', subLabel: '尚未開放報名', isOpen: false };
  }
  return { status: '開放報名',   subLabel: '',             isOpen: true  };
}

// ── ACTIVITIES API ──────────────────────────────────────────────
function getActivities() {
  var sh   = getSheet(SHEET_ACTIVITIES);
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1).filter(function(r){ return r[0]; }).map(function(r){
    var deadline = r[4] ? new Date(r[4]) : null;
    var openDate = r[6] ? new Date(r[6]) : null;  // col G：開放報名日期
    var calc     = calcActivityStatus(openDate, deadline);

    return {
      id:          String(r[0]),
      name:        r[1],
      startDate:   formatDate(r[2]),
      endDate:     formatDate(r[3]),
      deadline:    deadline ? formatDate(deadline) : '',
      deadlineRaw: deadline ? deadline.toISOString() : '',
      openDate:    openDate ? formatDate(openDate)  : '',
      openDateRaw: openDate ? openDate.toISOString() : '',
      status:      calc.status,
      statusLabel: calc.status,
      subLabel:    calc.subLabel,
      isOpen:      calc.isOpen
    };
  });
}

function saveActivity(data) {
  var sh  = getSheet(SHEET_ACTIVITIES);
  var id  = data.id ? String(data.id) : genId('A');
  var now = new Date();

  var row = [
    id,
    data.name,
    data.startDate ? new Date(data.startDate) : '',
    data.endDate   ? new Date(data.endDate)   : '',
    data.deadline  ? new Date(data.deadline)  : '',
    now,                                           // col F：建立/更新時間
    data.openDate  ? new Date(data.openDate)  : '' // col G：開放報名日期
  ];

  if (data.id) {
    var idx = findRowIndex(sh, 0, data.id);
    if (idx > 0) {
      sh.getRange(idx, 1, 1, row.length).setValues([row]);
      return { success: true, id: id };
    }
  }
  var newRow = sh.getLastRow() + 1;
  sh.getRange(newRow, 1, 1, row.length).setValues([row]);
  return { success: true, id: id };
}

function deleteActivity(activityId) {
  var sh  = getSheet(SHEET_ACTIVITIES);
  var idx = findRowIndex(sh, 0, activityId);
  if (idx > 0) sh.deleteRow(idx);
  return { success: true };
}

function deleteActivity(activityId) {
  var sh  = getSheet(SHEET_ACTIVITIES);
  var idx = findRowIndex(sh, 0, activityId);
  if (idx > 0) sh.deleteRow(idx);
  return { success: true };
}

// ── RSVP API (報名紀錄) ────────────────────────────────────────

function getRsvpByActivity(activityId) {
  var sh   = getSheet(SHEET_RSVP);
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  return data.slice(1).filter(function(r){ return String(r[1]) === String(activityId); }).map(function(r){
    return {
      id:           String(r[0]),
      activityId:   String(r[1]),
      activityName: r[2],
      memberId:     String(r[3]),
      memberName:   r[4],
      familyId:     String(r[5]),
      familyName:   r[6],
      status:       r[7],
      note:         r[8],
      rsvpTime:     formatDateTime(r[9]),
      updatedAt:    formatDateTime(r[10])
    };
  });
}

function submitRsvp(payload) {
  // payload: { activityId, activityName, familyId, familyName, members: [{id, name, status, note}] }
  var sh   = getSheet(SHEET_RSVP);
  var now  = new Date();
  var all  = sh.getDataRange().getValues();
  
  payload.members.forEach(function(m) {
    var existIdx = -1;
    for (var i = 1; i < all.length; i++) {
      if (String(all[i][1]) === String(payload.activityId) &&
          String(all[i][3]) === String(m.id)) {
        existIdx = i;
        break;
      }
    }

    var rsvpId = existIdx >= 0 ? String(all[existIdx][0]) : genId('R');
    var rsvpTime = existIdx >= 0 ? all[existIdx][9] : now;

    var row = [
      rsvpId,
      String(payload.activityId),
      payload.activityName,
      String(m.id),
      m.name,
      String(payload.familyId),
      payload.familyName,
      m.status,
      m.note || '',
      rsvpTime,
      now
    ];

    if (existIdx >= 0) {
      sh.getRange(existIdx + 1, 1, 1, row.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
  });

  return { success: true };
}

// ── STATS API (活動統計運算) ────────────────────────────────────

function getActivityStats(activityId) {
  var records   = getRsvpByActivity(activityId);
  var members   = getMembers();
  var families  = getFamilies();
  
  var memberMap = {};
  members.forEach(function(m){ memberMap[m.id] = m; });
  
  var total   = records.filter(function(r){ return r.memberName !== '整體備註'; }).length;
  var attend  = 0, absent = 0, pending = 0;
  var byFamily = {}, byTroop = {}, bySquad = {};
  var rsvpFamilyIds = {};
  
  records.forEach(function(r) {
    var fk = r.familyId;
    var fname = r.familyName || r.familyId;
    if (!byFamily[fk]) byFamily[fk] = { id: fk, name: fname, attend:0, absent:0, pending:0, unregistered:0, members:[], hasRsvp: true };

    if (r.memberName === '整體備註') {
      byFamily[fk].globalNote = r.note || '';
      return;
    }

    var m = memberMap[r.memberId] || {};
    if (m.position === '離團' || m.position === '無') return;

    rsvpFamilyIds[r.familyId] = true;
    if (r.status === '出席')    attend++;
    else if (r.status === '不出席') absent++;
    else pending++;
    
    byFamily[fk][r.status==='出席'?'attend':r.status==='不出席'?'absent':'pending']++;
    byFamily[fk].members.push({ name: r.memberName, naturalName: m.naturalName || '', status: r.status, note: r.note, rsvpTime: r.rsvpTime || '' });

    var tk = m.troop || '未分配';
    if (!byTroop[tk]) byTroop[tk] = { name: tk, attend:0, absent:0, pending:0, unregistered:0, members:[] };
    byTroop[tk][r.status==='出席'?'attend':r.status==='不出席'?'absent':'pending']++;
    byTroop[tk].members.push({ name: r.memberName, naturalName: m.naturalName || '', status: r.status, squad: m.squad || '', position: m.position || '', note: r.note || '' });
    
    var sk = m.squad || '未分配';
    if (!bySquad[sk]) bySquad[sk] = { name: sk, attend:0, absent:0, pending:0, unregistered:0 };
    bySquad[sk][r.status==='出席'?'attend':r.status==='不出席'?'absent':'pending']++;
  });
  
  var totalFamilies = families.length;
  var registeredFamilyCount = Object.keys(rsvpFamilyIds).length;
  var unregisteredFamilyCount = 0;
  
  families.forEach(function(f) {
    if (!rsvpFamilyIds[f.id]) {
      unregisteredFamilyCount++;
      var fmembers = members.filter(function(m){ return String(m.familyId) === String(f.id) && m.position !== '離團' && m.position !== '無'; });
      byFamily[f.id] = {
        id: f.id, name: f.name,
        attend: 0, absent: 0, pending: 0,
        unregistered: fmembers.length, 
        members: fmembers.map(function(m){
          return { name: m.name, naturalName: m.naturalName || '', status: '未報名', note: '', rsvpTime: '' };
        }),
        hasRsvp: false
      };
    }
  });
  
  var totalMembers = members.filter(function(m){ return m.position !== '離團' && m.position !== '無'; }).length;

  // 依團別整理所有成員（含未報名），供前端顯示完整出席狀況
  var allMembersByTroop = {};
  members.forEach(function(m) {
    if (m.position === '離團' || m.position === '無') return;
    var tk = m.troop || '未分配';
    if (!allMembersByTroop[tk]) allMembersByTroop[tk] = [];
    allMembersByTroop[tk].push({ name: m.name, naturalName: m.naturalName || '', squad: m.squad || '', position: m.position || '' });
  });
  
  // 計算各團/各隊未填寫人數
  Object.keys(allMembersByTroop).forEach(function(tk) {
    var allCount = allMembersByTroop[tk].length;
    var rsvpCount = byTroop[tk] ? (byTroop[tk].attend + byTroop[tk].absent + byTroop[tk].pending) : 0;
    if (byTroop[tk]) byTroop[tk].unregistered = allCount - rsvpCount;
  });
  
  return {
    total: total, attend: attend, absent: absent, pending: pending,
    attendRate: totalMembers ? Math.round(attend/totalMembers*100) : 0,
    totalFamilies: totalFamilies,
    registeredFamilyCount: registeredFamilyCount,
    unregisteredFamilyCount: unregisteredFamilyCount,
    familyRegRate: totalFamilies ? Math.round(registeredFamilyCount/totalFamilies*100) : 0,
    totalMembers: totalMembers,
    memberRegRate: totalMembers ? Math.round(total/totalMembers*100) : 0,
    allMembersByTroop: allMembersByTroop,
    byFamily: Object.values(byFamily),
    byTroop:  Object.values(byTroop),
    bySquad:  Object.values(bySquad),
    records:  records
  };
}

function testGetActivities() {
  var result = getActivities();
  Logger.log(JSON.stringify(result));
}