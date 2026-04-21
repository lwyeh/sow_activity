// ================================================================
//  Api_Payment.gs - 收費活動與繳費系統 API（重新設計版）
// ================================================================
//
// 資料表結構說明：
//
// [收費活動] SHEET_PAY_ACTS
//   活動ID | 活動名稱 | 說明 | 類型(type1/type2) | 建立時間
//
// [收費設備] SHEET_PAY_EQUIPS  (新)
//   設備ID | 活動ID | 設備名稱 | 金額 | 數量上限 | 種類(使用者/管理者)
//
// [收費明細] SHEET_PAY_DETAILS  (重新設計)
//   明細ID | 活動ID | 成員ID | 家庭ID | 設備ID | 設備名稱 | 數量 | 單價 | 小計
//   (type1: 使用者自選設備 or 管理者強加; type2: 管理者指定批次)
//
// [繳費紀錄] SHEET_PAY_RECORDS  (欄位同原本，但 detailId → actId+familyId 合併繳)
//   紀錄ID | 活動ID | 家庭ID | 繳費金額 | 繳費方式 | 備註 | 狀態 | 退回原因 | 送出時間 | 確認時間 | 建立時間
//
// ================================================================

// ── 收費活動 CRUD ─────────────────────────────────────────────

// ── 儲存/更新收費活動 ──
function savePayActivity(data) {
  var sh = getSheet('收費活動');
  var dataRange = sh.getDataRange().getValues();
  var methodsStr = (data.methods || []).join(',');

  if (data.id) {
    for (var i = 1; i < dataRange.length; i++) {
      if (String(dataRange[i][0]) === String(data.id)) {
        sh.getRange(i + 1, 2).setValue(data.name);
        sh.getRange(i + 1, 3).setValue(data.note);
        sh.getRange(i + 1, 4).setValue(data.type);
        sh.getRange(i + 1, 6).setValue(methodsStr);
        sh.getRange(i + 1, 7).setValue(data.openDate  ? new Date(data.openDate)  : ''); // ★ col G
        sh.getRange(i + 1, 8).setValue(data.deadline  ? new Date(data.deadline)  : ''); // ★ col H
        return { id: data.id };
      }
    }
  } else {
    var newId = genId('PA');
    var now = new Date();
    // [ID, 名稱, 說明, 類型, 建立時間, 繳費方式, 開放日期, 截止日期]
    sh.appendRow([
      newId,
      data.name,
      data.note,
      data.type,
      now,
      methodsStr,
      data.openDate ? new Date(data.openDate) : '',  // ★ col G
      data.deadline ? new Date(data.deadline) : ''   // ★ col H
    ]);
    return { id: newId };
  }
}

// ── 取得所有收費活動 ──
function getPayActivities() {
  var sh = getSheet('收費活動');
  var data = sh.getDataRange().getValues();
  var res = [];

  var now = new Date();
  now.setHours(0, 0, 0, 0);

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;

    var methodsStr = data[i][5] ? String(data[i][5]) : '現金,LinePay,iPassMoney,轉帳,其他';
    var openDate   = data[i][6] ? new Date(data[i][6]) : null;  // ★ col G
    var deadline   = data[i][7] ? new Date(data[i][7]) : null;  // ★ col H

    // 與報名活動相同的狀態推導邏輯
    var calc = calcActivityStatus(openDate, deadline);

    res.push({
      id:          data[i][0],
      name:        data[i][1],
      note:        data[i][2],
      type:        data[i][3],
      createdAt:   Utilities.formatDate(new Date(data[i][4]), "GMT+8", "yyyy/MM/dd"),
      methods:     methodsStr.split(','),
      openDate:    openDate ? Utilities.formatDate(openDate, "GMT+8", "yyyy/MM/dd") : '',
      deadline:    deadline ? Utilities.formatDate(deadline, "GMT+8", "yyyy/MM/dd") : '',
      status:      calc.status,    // ★ 開放報名 / 暫停報名 / 已結束
      subLabel:    calc.subLabel,  // ★ 尚未開放 / 已截止
      isOpen:      calc.isOpen     // ★ 是否可操作
    });
  }
  return res;
}


function deletePayActivity(actId) {
  // 刪除活動
  var sh  = getSheet(SHEET_PAY_ACTS);
  var idx = findRowIndex(sh, 0, actId);
  if (idx > 0) sh.deleteRow(idx);

  // 刪除設備
  var esh  = getSheet(SHEET_PAY_EQUIPS);
  var edata = esh.getDataRange().getValues();
  for (var i = edata.length - 1; i >= 1; i--) {
    if (String(edata[i][1]) === String(actId)) esh.deleteRow(i + 1);
  }

  // 刪除明細
  var dsh  = getSheet(SHEET_PAY_DETAILS);
  var ddata = dsh.getDataRange().getValues();
  for (var j = ddata.length - 1; j >= 1; j--) {
    if (String(ddata[j][1]) === String(actId)) dsh.deleteRow(j + 1);
  }

  // 刪除繳費紀錄
  var rsh  = getSheet(SHEET_PAY_RECORDS);
  var rdata = rsh.getDataRange().getValues();
  for (var k = rdata.length - 1; k >= 1; k--) {
    if (String(rdata[k][1]) === String(actId)) rsh.deleteRow(k + 1);
  }

  return { success: true };
}

// ── 設備管理 (type1 & type2 共用) ────────────────────────────

function getPayEquips(actId) {
  var sh   = getSheet(SHEET_PAY_EQUIPS);
  var data = sh.getDataRange().getValues();
  return data.slice(1).filter(function(r){ return r[0] && String(r[1]) === String(actId); }).map(function(r) {
    return {
      id:       String(r[0]),
      actId:    String(r[1]),
      name:     r[2] || '',
      price:    Number(r[3]) || 0,
      maxQty:   Number(r[4]) || 0,
      category: r[5] || '使用者'   // '使用者' | '管理者'
    };
  });
}

// 新增一筆設備
function addPayEquip(data) {
  var sh  = getSheet(SHEET_PAY_EQUIPS);
  var id  = genId('PE');
  sh.appendRow([id, String(data.actId), data.name || '', Number(data.price) || 0, Number(data.maxQty) || 0, data.category || '使用者']);
  return { success: true, id: id };
}

// 刪除一筆設備（同時刪除對應明細）
function deletePayEquip(equipId) {
  var sh  = getSheet(SHEET_PAY_EQUIPS);
  var idx = findRowIndex(sh, 0, equipId);
  if (idx > 0) sh.deleteRow(idx);

  // 刪除對應明細
  var dsh  = getSheet(SHEET_PAY_DETAILS);
  var ddata = dsh.getDataRange().getValues();
  for (var i = ddata.length - 1; i >= 1; i--) {
    if (String(ddata[i][4]) === String(equipId)) dsh.deleteRow(i + 1);
  }

  return { success: true };
}

// ── 收費明細 ──────────────────────────────────────────────────

// 取得一個活動的所有明細（加上設備與成員名稱）
function getPayDetailsByAct(actId) {
  var dsh  = getSheet(SHEET_PAY_DETAILS);
  var ddata = dsh.getDataRange().getValues();

  var equips  = getPayEquips(actId);
  var equipMap = {};
  equips.forEach(function(e){ equipMap[e.id] = e; });

  var members = getMembers();
  var mMap = {};
  members.forEach(function(m){ mMap[m.id] = m; });

  var families = getFamilies();
  var fMap = {};
  families.forEach(function(f){ fMap[f.id] = f; });

  return ddata.slice(1).filter(function(r){ return r[0] && String(r[1]) === String(actId); }).map(function(r) {
    var m = mMap[String(r[2])] || {};
    var f = fMap[String(r[3])] || {};
    return {
      id:          String(r[0]),
      actId:       String(r[1]),
      memberId:    String(r[2]),
      familyId:    String(r[3]),
      equipId:     String(r[4]),
      equipName:   r[5] || '',
      qty:         Number(r[6]) || 0,
      unitPrice:   Number(r[7]) || 0,
      subtotal:    Number(r[8]) || 0,
      memberName:  m.name || '',
      naturalName: m.naturalName || '',
      position:    m.position || '',
      familyName:  f.name || '',
      familyNo:    f.no || ''
    };
  });
}

// type1 - 使用者選擇設備與數量（儲存 / 更新）
// items = [{ equipId, equipName, qty, unitPrice }]
function saveType1Selection(actId, memberId, familyId, items) {
  var dsh  = getSheet(SHEET_PAY_DETAILS);
  var ddata = dsh.getDataRange().getValues();

  // 刪除該成員在此活動的「使用者」類設備明細（category=使用者）
  // 先取得此活動的 使用者 設備 ID 清單
  var equips = getPayEquips(actId);
  var userEquipIds = {};
  equips.forEach(function(e){ if (e.category === '使用者') userEquipIds[e.id] = true; });

  for (var i = ddata.length - 1; i >= 1; i--) {
    var r = ddata[i];
    if (String(r[1]) === String(actId) && String(r[2]) === String(memberId) && userEquipIds[String(r[4])]) {
      dsh.deleteRow(i + 1); // 這裡如果 memberId 對不上，就不會刪除舊項目
    }
  }

  // 寫入新選擇
  items.forEach(function(item) {
    if (!item.qty || item.qty <= 0) return;
    var subtotal = (Number(item.unitPrice) || 0) * (Number(item.qty) || 0);
    dsh.appendRow([genId('PD'), String(actId), String(memberId), String(familyId),
                   String(item.equipId), item.equipName || '', Number(item.qty), Number(item.unitPrice) || 0, subtotal]);
  });

  return { success: true };
}

// 管理員調整明細（type1 d. / type2 d.）：新增一筆管理者明細
function addAdminDetail(actId, memberId, familyId, equipId, equipName, qty, unitPrice) {
  var dsh = getSheet(SHEET_PAY_DETAILS);
  var subtotal = (Number(unitPrice) || 0) * (Number(qty) || 0);
  dsh.appendRow([genId('PD'), String(actId), String(memberId), String(familyId),
                 String(equipId || 'ADM'), equipName || '', Number(qty), Number(unitPrice) || 0, subtotal]);
  return { success: true };
}

// 刪除指定明細
function deletePayDetail(detailId) {
  var dsh = getSheet(SHEET_PAY_DETAILS);
  var idx = findRowIndex(dsh, 0, detailId);
  if (idx > 0) dsh.deleteRow(idx);
  return { success: true };
}

// type2 - 管理者批次指派明細
// items = [{ memberId, familyId, equipId, equipName, qty, unitPrice }]
function saveType2Batch(actId, items) {
  var dsh = getSheet(SHEET_PAY_DETAILS);
  items.forEach(function(item) {
    var subtotal = (Number(item.unitPrice) || 0) * (Number(item.qty) || 0);
    dsh.appendRow([genId('PD'), String(actId), String(item.memberId), String(item.familyId),
                   String(item.equipId), item.equipName || '', Number(item.qty), Number(item.unitPrice) || 0, subtotal]);
  });
  return { success: true };
}

// ── 繳費紀錄（家長端）────────────────────────────────────────
// 一個家庭一個活動合計成一筆

function getPaySummaryByFamily(familyId) {
  var acts    = getPayActivities();
  var actMap  = {};
  acts.forEach(function(a){ actMap[a.id] = a; });

  var dsh   = getSheet(SHEET_PAY_DETAILS);
  var ddata = dsh.getDataRange().getValues();

  var rsh   = getSheet(SHEET_PAY_RECORDS);
  var rdata = rsh.getDataRange().getValues();

  var members = getMembers();
  var mMap = {};
  members.forEach(function(m){ mMap[m.id] = m; });

  // 按 活動 分組，計算此家庭的應繳總計
  var actGroups = {};
  ddata.slice(1).forEach(function(r) {
    if (!r[0] || String(r[3]) !== String(familyId)) return;
    var aId = String(r[1]);
    if (!actGroups[aId]) actGroups[aId] = { details: [], records: [] };
    actGroups[aId].details.push({
      id:        String(r[0]),
      memberId:  String(r[2]),
      equipId:   String(r[4]),
      equipName: r[5] || '',
      qty:       Number(r[6]) || 0,
      unitPrice: Number(r[7]) || 0,
      subtotal:  Number(r[8]) || 0
    });
  });

  // 取得此家庭的繳費紀錄
  rdata.slice(1).forEach(function(r) {
    if (!r[0] || String(r[2]) !== String(familyId)) return;
    var aId = String(r[1]);
    if (!actGroups[aId]) actGroups[aId] = { details: [], records: [] };
    actGroups[aId].records.push({
      id:         String(r[0]),
      amount:     Number(r[3]) || 0,
      method:     r[4] || '',
      note:       r[5] || '',
      status:     r[6] || '未繳',
      rejectNote: r[7] || '',
      submitAt:   r[8] ? formatDateTime(new Date(r[8])) : '',
      confirmAt:  r[9] ? formatDateTime(new Date(r[9])) : ''
    });
  });

  // type1 活動即使家庭尚未選設備（無明細），也要顯示選設備按鈕
  acts.forEach(function(a) {
    if (a.type === 'type1' && !actGroups[a.id]) {
      actGroups[a.id] = { details: [], records: [] };
    }
  });

  var result = [];
  Object.keys(actGroups).forEach(function(aId) {
    var group  = actGroups[aId];
    var act    = actMap[aId] || { name: aId, type: 'type2' };
    // type2 沒有明細代表管理者尚未指定此家庭，不顯示
    if (!group.details.length && act.type !== 'type1') return;

    var totalAmount = group.details.reduce(function(s, d){ return s + d.subtotal; }, 0);
    var paidTotal   = group.records.filter(function(r){ return r.status === '已確認'; })
                              .reduce(function(s, r){ return s + r.amount; }, 0);
    var remaining   = totalAmount - paidTotal;

    // 整理明細列表（含成員名）
    var detailRows = group.details.map(function(d) {
      var m = mMap[d.memberId] || {};
      return {
        id:          d.id,
        memberId:    d.memberId,
        memberName:  m.name || '',
        naturalName: m.naturalName || '',
        equipName:   d.equipName,
        qty:         d.qty,
        unitPrice:   d.unitPrice,
        subtotal:    d.subtotal
      };
    });

    result.push({
      actId:       aId,
      actName:     act.name || aId,
      actType:     act.type || 'type2',
      actNote:     act.note || '',   // ★ 請務必加上這一行！把說明文字傳給前端
      actMethods:  act.methods || ['現金','LinePay','iPassMoney','轉帳','其他'], // ★ 補上這行，把允許的付款方式傳給家長頁面
      actStatus:   act.status  || '開放報名',  // ★
      isOpen:      act.isOpen !== false,        // ★（預設 true 向下相容舊資料）
      familyId:    familyId,
      totalAmount: totalAmount,
      paidTotal:   paidTotal,
      remaining:   remaining,
      details:     detailRows,
      records:     group.records
    });
  });

  result.sort(function(a, b){ return a.actName.localeCompare(b.actName); });
  return result;
}

// 使用者送出繳費（一個家庭一個活動一筆）
function submitPayRecord(data) {
  var sh  = getSheet(SHEET_PAY_RECORDS);
  var now = new Date();
  var row = [
    genId('PR'),
    String(data.actId),
    String(data.familyId),
    Number(data.amount) || 0,
    data.method || '',
    data.note   || '',
    '已送出',
    '',           // 退回原因
    now,          // 送出時間
    '',           // 確認時間
    now           // 建立時間
  ];
  sh.appendRow(row);
  return { success: true };
}

// ── 對帳管理 (管理員端) ───────────────────────────────────────

function getPendingPayRecords(filterActId) {
  var sh      = getSheet(SHEET_PAY_RECORDS);
  var data    = sh.getDataRange().getValues();
  var acts    = getPayActivities();
  var actMap  = {};
  acts.forEach(function(a){ actMap[a.id] = a; });
  var families = getFamilies();
  var fMap = {};
  families.forEach(function(f){ fMap[f.id] = f; });

  return data.slice(1).filter(function(r){
    if (!r[0] || r[6] !== '已送出') return false;
    if (filterActId && String(r[1]) !== String(filterActId)) return false;
    return true;
  }).map(function(r) {
    var act = actMap[String(r[1])] || {};
    var f   = fMap[String(r[2])] || {};
    return {
      id:         String(r[0]),
      actId:      String(r[1]),
      actName:    act.name || String(r[1]),
      familyId:   String(r[2]),
      familyNo:   f.no || '',
      familyName: f.name || String(r[2]),
      amount:     Number(r[3]) || 0,
      method:     r[4] || '',
      note:       r[5] || '',
      status:     r[6] || '',
      submitAt:   r[8] ? formatDateTime(new Date(r[8])) : ''
    };
  });
}

function confirmPayRecord(recordId) {
  var sh  = getSheet(SHEET_PAY_RECORDS);
  var idx = findRowIndex(sh, 0, recordId);
  if (idx < 0) return { success: false };
  sh.getRange(idx, 7).setValue('已確認');
  sh.getRange(idx, 10).setValue(new Date());
  return { success: true };
}

// ── 退回繳費紀錄 (包含自動發送 Email 通知，並修正退回原因寫入欄位) ────────────────────────────────
function rejectPayRecord(recordId, reason) {
  var sh = getSheet('繳費紀錄'); // 確保工作表名稱正確
  var data = sh.getDataRange().getValues();
  var targetRow = -1;
  
  var actId = '';
  var familyId = '';
  var amount = 0;
  var method = '';
  var submitDate = '';

  // 1. 尋找該筆紀錄
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(recordId)) {
      targetRow = i + 1;
      actId = data[i][1];
      familyId = data[i][2];
      amount = data[i][3];
      method = data[i][4];
      submitDate = data[i][8] ? Utilities.formatDate(new Date(data[i][8]), "GMT+8", "yyyy/MM/dd HH:mm") : '未知時間';
      break;
    }
  }

  if (targetRow === -1) throw new Error('找不到此繳費紀錄');

  // 2. 更新狀態與退回原因
  // ★ 修正：狀態是 G 欄 (第 7 欄)，退回原因應該是 H 欄 (第 8 欄)！
  sh.getRange(targetRow, 7).setValue('已退回');
  sh.getRange(targetRow, 8).setValue(reason); 

  // 3. 嘗試發送 Email 通知
  try {
    var actName = '收費活動';
    var actSh = getSheet('收費活動');
    if (actSh) {
      var actData = actSh.getDataRange().getValues();
      for (var j = 1; j < actData.length; j++) {
        if (String(actData[j][0]) === String(actId)) { 
          actName = actData[j][1]; 
          break; 
        }
      }
    }

    var families = getFamilies();
    var famName = '家長';
    families.forEach(function(f) { 
      if (String(f.id) === String(familyId)) famName = f.name; 
    });

    var members = getMembers();
    var emails = [];
    members.forEach(function(m) {
      if (String(m.familyId) === String(familyId) && m.email) {
        if (emails.indexOf(m.email) === -1) emails.push(m.email);
      }
    });

    if (emails.length > 0) {
      var subject = '【荒野親子團】繳費退回通知：' + actName;
      var body = famName + ' 您好，\n\n' +
                 '您於 ' + submitDate + ' 提交的「' + actName + '」繳費回報（' + method + ' $' + amount + '），' +
                 '已由管理員退回。\n\n' +
                 '【退回原因】：\n' + reason + '\n\n' +
                 '煩請您登入報名系統，進入「登記與繳費」頁面查看明細，並重新回報，謝謝您的配合！\n\n' +
                 '荒野親子團 系統自動通知';
      
      MailApp.sendEmail({
        to: emails.join(','),
        subject: subject,
        body: body,
        name: '荒野親子團報名系統'
      });
    }
  } catch (e) {
    console.error('退回通知 Email 發送失敗:', e);
  }

  return { success: true };
}


// ── 收費報表 ─────────────────────────────────────────────────

// ── 取得收費報表 (加入支援分次/多種繳費方式的精準統計) ────────────────────────────────
function getPayReport(actId, reportMode, posFilter) {
  reportMode = String(reportMode || '家庭').trim();
  posFilter = String(posFilter || '').trim();
  var validRoles = ['小蟻', '小蜂', '小鹿', '小鷹', '家長', '戶長'];
  if (posFilter && validRoles.indexOf(posFilter) === -1) {
    posFilter = '';
  }

  var dsh = getSheet('收費明細');
  var ddata = dsh.getDataRange().getValues();
  var rsh = getSheet('繳費紀錄'); 
  var rdata = rsh.getDataRange().getValues();

  var families = getFamilies();
  var members = getMembers();
  var fMap = {}; families.forEach(function(f){ fMap[f.id] = f; });

  var famStatus = {};
  var targetActId = String(actId).trim();

  // 讀取管理員專屬的「對帳備註」
  var adminNotesMap = {};
  var noteSh = getDb().getSheetByName('對帳備註');
  if (noteSh) {
    var nData = noteSh.getDataRange().getValues();
    for (var i = 1; i < nData.length; i++) {
      if (String(nData[i][0]).trim() === targetActId) {
        adminNotesMap[String(nData[i][1]).trim()] = String(nData[i][2]);
      }
    }
  }

  rdata.slice(1).forEach(function(r) {
    if (!r[0] || String(r[1]).trim() !== targetActId) return; 
    var fid = String(r[2]).trim();
    if (!famStatus[fid]) {
      famStatus[fid] = { paidTotal: 0, method: '', submitAt: '', confirmAt: '', note: '', status: '', activeRecords: [] };
    }
    var currentStatus = String(r[6]).trim();
    
    // ★ 核心修正：將所有有效的繳費紀錄都保存下來，供前端準確拆解各付款方式
    if (currentStatus === '已確認' || currentStatus === '待確認' || currentStatus === '已送出') {
      famStatus[fid].activeRecords.push({
        method: String(r[4]).trim(),
        amount: Number(r[3]) || 0,
        status: currentStatus
      });
    }

    if (currentStatus === '已確認') {
      famStatus[fid].paidTotal += (Number(r[3]) || 0);
      famStatus[fid].status = '已完成';
      famStatus[fid].method = r[4];
      famStatus[fid].submitAt = r[8] ? Utilities.formatDate(new Date(r[8]), "GMT+8", "yyyy/MM/dd HH:mm") : '';
      famStatus[fid].confirmAt = r[9] ? Utilities.formatDate(new Date(r[9]), "GMT+8", "yyyy/MM/dd HH:mm") : '';
      if (r[5]) famStatus[fid].note = r[5];
    } else if (currentStatus === '待確認' || currentStatus === '已送出') {
      if (famStatus[fid].status !== '已完成') {
        famStatus[fid].status = '待確認';
        famStatus[fid].method = r[4];
        famStatus[fid].submitAt = r[8] ? Utilities.formatDate(new Date(r[8]), "GMT+8", "yyyy/MM/dd HH:mm") : '';
        if (r[5]) famStatus[fid].note = r[5];
      }
    }
  });

  var memDetails = {};
  var famTotals = {};
  ddata.slice(1).forEach(function(r) {
    if (!r[0] || String(r[1]).trim() !== targetActId) return;
    var mid = String(r[2]).trim();
    var fid = String(r[3]).trim();
    var amt = Number(r[8]) || 0;
    
    if (!mid || !fid) return; // ★ 防呆：跳過 ID 為空的異常明細

    if (!memDetails[mid]) memDetails[mid] = [];
    memDetails[mid].push({ equipName: r[5], qty: Number(r[6]) || 1, amount: amt });
    if (!famTotals[fid]) famTotals[fid] = 0;
    famTotals[fid] += amt;
  });

  var results = [];

  if (reportMode === '個人') {
    members.forEach(function(mem) {
      if (mem.position === '離團') return;
      if (posFilter && mem.position !== posFilter) return;

      var fid = mem.familyId;
      var fam = fMap[fid] || {};
      var fs = famStatus[fid] || { paidTotal: 0, method: '', submitAt: '', confirmAt: '', note: '', status: '', activeRecords: [] };
      var dispName = mem.naturalName || mem.name || '無姓名';
      var details = memDetails[String(mem.id).trim()];
      var aNote = adminNotesMap[fid] || '';

      // ⭐ 新增抓取會員編號與入團日期 (加入自動格式化日期防呆)
      var mNo = mem.memberNo || '';
      var jDate = mem.joinDate;
      if (jDate instanceof Date) {
        jDate = Utilities.formatDate(jDate, "GMT+8", "yyyy/MM/dd");
      } else {
        jDate = jDate || '';
      }

      if (details && details.length > 0) {
        details.forEach(function(d) {
          // 👉 這裡的 push 多加了 memberNo: mNo, joinDate: jDate
          results.push({ 
            familyId: fid, familyNo: fam.no || '', familyName: fam.name || fid, 
            memberNo: mem.memberNo || '', // 👉 確保這裡有加上
            memberName: dispName, equipName: d.equipName, qty: d.qty, amount: d.amount, status: fs.status || '未繳', submitAt: fs.submitAt, note: fs.note, adminNote: aNote, activeRecords: fs.activeRecords 
          });
        });
      } else {
        // 👉 這裡的 push 也多加了 memberNo: mNo, joinDate: jDate
        results.push({ 
          familyId: fid, familyNo: fam.no || '', familyName: fam.name || fid, 
          memberNo: mem.memberNo || '', // 👉 確保這裡有加上
          memberName: dispName, equipName: '無項目', qty: 0, amount: 0, status: fs.status || '未產生費用', submitAt: fs.submitAt, note: fs.note, adminNote: aNote, activeRecords: fs.activeRecords 
        });
      }
    });
    return results;
  } 
  else {
    families.forEach(function(fam) {
      var fid = String(fam.id).trim();
      if (posFilter) {
        var hasMatch = members.some(function(m){ return m.familyId === fid && m.position === posFilter && m.position !== '離團'; });
        if (!hasMatch) return;
      }
      var activeCount = members.filter(function(m){ return m.familyId === fid && m.position !== '離團'; }).length;
      if (activeCount === 0) return;

      var total = famTotals[fid];
      var fs = famStatus[fid] || { paidTotal: 0, method: '', submitAt: '', confirmAt: '', note: '', status: '', activeRecords: [] };
      var aNote = adminNotesMap[fid] || '';

      if (total === undefined && !famStatus[fid]) {
        fs.status = '未產生費用';
        total = 0;
      } else {
        total = total || 0;
        var remaining = total - fs.paidTotal;
        if (remaining <= 0 && (!fs.status || fs.status === '未繳') && total !== 0) fs.status = '已完成';
        else if (remaining === 0 && total === 0 && (!fs.status || fs.status === '未繳')) fs.status = '已結清';
        else if (!fs.status) fs.status = '未繳';
      }

      results.push({ familyId: fid, familyNo: fam.no || '', familyName: fam.name || fid, amount: total, paidTotal: fs.paidTotal, remaining: total - fs.paidTotal, status: fs.status, method: fs.method, submitAt: fs.submitAt, confirmAt: fs.confirmAt, note: fs.note, adminNote: aNote, activeRecords: fs.activeRecords });
    });
    return results;
  }
}

// ── 管理員對帳調整：取得某活動某家庭的明細 ──────────────────

function getAdminFamilyDetails(actId, familyId) {
  var dsh   = getSheet(SHEET_PAY_DETAILS);
  var ddata = dsh.getDataRange().getValues();

  var members = getMembers();
  var mMap = {};
  members.forEach(function(m){ mMap[m.id] = m; });

  var equips  = getPayEquips(actId);
  var equipMap = {};
  equips.forEach(function(e){ equipMap[e.id] = e; });

  return ddata.slice(1).filter(function(r){
    return r[0] && String(r[1]) === String(actId) && String(r[3]) === String(familyId);
  }).map(function(r) {
    var m = mMap[String(r[2])] || {};
    return {
      id:          String(r[0]),
      memberId:    String(r[2]),
      memberName:  m.name || '',
      naturalName: m.naturalName || '',
      equipId:     String(r[4]),
      equipName:   r[5] || '',
      qty:         Number(r[6]) || 0,
      unitPrice:   Number(r[7]) || 0,
      subtotal:    Number(r[8]) || 0
    };
  });
}

// ── 新增：管理員直接紀錄退款 ─────────────────────────────────────
function adminSubmitRefund(actId, familyId, amount, note) {
  var sh = getSheet(SHEET_PAY_RECORDS);
  var now = new Date();
  // amount 會是負數，例如 -66
  var row = [
    genId('PR'), 
    String(actId), 
    String(familyId), 
    Number(amount) || 0, 
    '系統手動退款', 
    note || '', 
    '已確認',    // 直接設為已確認，不需要再對帳
    '', 
    now, 
    now, 
    now
  ];
  sh.appendRow(row);
  return { success: true };
}

// ── 新增：取得所有待退費的清單（供對帳區使用）────────────────────────────────
function getPendingRefunds(filterActId) {
  var dsh = getSheet(SHEET_PAY_DETAILS);
  var ddata = dsh.getDataRange().getValues();
  var rsh = getSheet(SHEET_PAY_RECORDS);
  var rdata = rsh.getDataRange().getValues();

  var acts = getPayActivities();
  var actMap = {}; acts.forEach(function(a){ actMap[a.id] = a; });
  var families = getFamilies();
  var fMap = {}; families.forEach(function(f){ fMap[f.id] = f; });

  var balances = {};

  // 1. 計算應繳總額
  ddata.slice(1).forEach(function(r) {
    if (!r[0]) return;
    var aId = String(r[1]);
    if (filterActId && aId !== String(filterActId)) return;
    var fId = String(r[3]);
    var key = aId + '_' + fId;
    if (!balances[key]) balances[key] = { actId: aId, familyId: fId, total: 0, paid: 0 };
    balances[key].total += (Number(r[8]) || 0);
  });

  // 2. 計算已繳總額 (只算已確認的)
  rdata.slice(1).forEach(function(r) {
    if (!r[0] || r[6] !== '已確認') return;
    var aId = String(r[1]);
    if (filterActId && aId !== String(filterActId)) return;
    var fId = String(r[2]);
    var key = aId + '_' + fId;
    if (!balances[key]) balances[key] = { actId: aId, familyId: fId, total: 0, paid: 0 };
    balances[key].paid += (Number(r[3]) || 0);
  });

  var refunds = [];
  Object.keys(balances).forEach(function(k) {
    var b = balances[k];
    var remaining = b.total - b.paid;
    // 如果剩餘金額 < 0，代表系統欠家長錢，是待退費
    if (remaining < 0) {
      var act = actMap[b.actId] || {};
      var fam = fMap[b.familyId] || {};
      refunds.push({
        id: b.actId + '_' + b.familyId,
        actId: b.actId,
        actName: act.name || b.actId,
        familyId: b.familyId,
        familyName: fam.name || b.familyId,
        familyNo: fam.no || '',
        amount: remaining // 負數
      });
    }
  });

  return refunds;
}

// ── 新增：儲存管理員對帳備註 (修正日期自動轉換問題) ────────────────────────────────
function saveAdminNote(actId, familyId, note) {
  var ss = getDb();
  var sh = ss.getSheetByName('對帳備註');
  
  // 如果還沒有這張表，系統自動建立並寫入標題
  if (!sh) {
    sh = ss.insertSheet('對帳備註');
    sh.appendRow(['活動ID', '家庭ID', '管理員備註']);
    sh.getRange("A1:C1").setFontWeight("bold");
    sh.getRange("C:C").setNumberFormat('@'); // ★ 強制將 C 欄 (備註欄) 預設為純文字格式
    sh.setFrozenRows(1);
    sh.hideSheet(); // 預設隱藏
  }
  
  var data = sh.getDataRange().getValues();
  // 尋找是否已有該活動與家庭的紀錄，有則更新，無則新增
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(actId).trim() && String(data[i][1]).trim() === String(familyId).trim()) {
      // ★ 寫入前，再三確認將該儲存格格式設為純文字 ('@')，防止 Google 雞婆轉換
      sh.getRange(i + 1, 3).setNumberFormat('@').setValue(note);
      return { success: true };
    }
  }
  
  // 如果找不到舊紀錄，就新增一列
  sh.appendRow([actId, familyId, note]);
  // ★ 針對剛新增的那一列，強制把備註欄設為純文字
  var lastRow = sh.getLastRow();
  sh.getRange(lastRow, 3).setNumberFormat('@').setValue(note); 
  
  return { success: true };
}

// ── 新增：取得活動可選設備 (支援「全域數量上限」與「個人已選數量」盤點) ──
function getPayEquipsWithLimit(actId, memberId) {
  // 1. 先取得原本的設備定義 (呼叫您原本寫好的 getPayEquips)
  var equips = getPayEquips(actId);
  if (!equips || equips.length === 0) return [];
  
  // 2. 掃描收費明細，計算全團目前被選走的總量
  var dsh = getSheet('收費明細');
  var ddata = dsh.getDataRange().getValues();
  
  var consumed = {};
  var myConsumed = {};
  
  ddata.slice(1).forEach(function(r) {
    if (!r[0] || String(r[1]).trim() !== String(actId).trim()) return;
    
    var mId = String(r[2]).trim();
    var eqName = String(r[5]).trim(); // 使用名稱來比對最安全
    var qty = Number(r[6]) || 0;
    
    if (!consumed[eqName]) consumed[eqName] = 0;
    consumed[eqName] += qty;
    
    // 如果有傳入成員ID，順便記錄該成員自己已經選了幾個
    if (memberId && mId === String(memberId).trim()) {
      if (!myConsumed[eqName]) myConsumed[eqName] = 0;
      myConsumed[eqName] += qty;
    }
  });
  
  // 3. 將計算好的「剩餘庫存」與「自己已選數量」合併回 equips 陣列
  equips.forEach(function(e) {
    var c = consumed[String(e.name).trim()] || 0;
    var m = myConsumed[String(e.name).trim()] || 0;
    
    e.myQty = m; // 該成員上次已經選的數量
    
    // 剩餘數量 = (總量上限) - (全團已選走總數) + (我自己已經選的，因為我在編輯時可以保有自己的額度)
    if (e.maxQty && Number(e.maxQty) > 0) {
      e.remainQty = Number(e.maxQty) - c + m;
      if (e.remainQty < 0) e.remainQty = 0;
    } else {
      e.remainQty = 9999; // 沒設上限就給一個極大值
    }
  });
  
  return equips;
}

function testSendEmail() {
  // 將下方的 Email 改成您自己能收信的信箱，用來測試
  var myEmail = "您的信箱@gmail.com"; 
  
  MailApp.sendEmail({
    to: "lunwu.yeh@gmail.com",
    subject: "荒野親子團 - 寄信功能測試",
    body: "如果您收到這封信，代表寄信權限已經成功開啟囉！"
  });
  Logger.log("測試信件已發送");
}

function testGetPayActivities() {
  var result = getPayActivities();
  Logger.log(JSON.stringify(result));
}

function debugPayActs() {
  var sh = getSheet('收費活動');
  var data = sh.getDataRange().getValues();
  
  // 印出標題列（看欄位順序）
  Logger.log('標題列: ' + JSON.stringify(data[0]));
  
  // 印出第一筆資料的每個欄位與索引
  if (data.length > 1) {
    data[1].forEach(function(val, idx) {
      Logger.log('索引 ' + idx + ' (欄 ' + String.fromCharCode(65 + idx) + '): ' + val);
    });
  }
}