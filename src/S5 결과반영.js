/** S5. 피킹라인의 최소 입력(O/X)을 주문 단위로 자동 반영한다. */
function S5_1_결과반영() {
  return withLock_(function () {
    var exceptions = findPickingExceptionOrders_(), cancelled = [];
    exceptions.forEach(function (item) {
      // X만 입력하고 사유를 비운 경우에는 사람이 사유를 고를 때까지 기다린다.
      if (!item.reason) return;
      var result = cancelOrder_(item.orderNo, item.reason, 'PICKING_X', { refresh: false });
      if (result.취소) cancelled.push({ 주문번호: item.orderNo, 사유: item.reason });
    });
    var completed = completePickingOrders_();
    var summary = {
      차감라인: completed.lines, 차감수량: completed.quantity, 취소주문: cancelled,
      복원라인: 0, 복원수량: 0, 완료주문: completed.orders, 담당자전파: completed.propagated
    };
    if (cancelled.length || completed.orders.length) D0_대시보드전체갱신(true);
    writeOpLog_('S5_1_결과반영', '성공', '출고완료 ' + completed.orders.length + ' / 취소 ' + cancelled.length);
    return summary;
  });
}

function findPickingExceptionOrders_() {
  var 라인 = readTable_(ROLE.라인), 주문 = readTable_(ROLE.주문);
  var L = { 주문번호: col_(라인, COL.주문번호, false), 품목별: col_(라인, COL.품목별주문번호, true),
    확인: col_(라인, COL.확인, true), 사유: col_(라인, COL.예외사유, true), 상태: col_(라인, COL.라인상태, true) };
  var itemOrder = {}, oItem = col_(주문, COL.품목별주문번호, true), oNo = col_(주문, COL.주문번호, true);
  주문.rows.forEach(function (row) { itemOrder[toStr_(row[oItem])] = toStr_(row[oNo]); });
  var found = {};
  라인.rows.forEach(function (row) {
    if (toStr_(row[L.상태]) !== ENUM.라인상태.미처리 || toStr_(row[L.확인]).toUpperCase() !== ENUM.확인.예외) return;
    var no = (L.주문번호 >= 0 ? toStr_(row[L.주문번호]) : '') || itemOrder[toStr_(row[L.품목별])];
    var reason = toStr_(row[L.사유]);
    if (no && (!found[no] || (!found[no].reason && reason))) found[no] = { orderNo: no, reason: reason };
  });
  return Object.keys(found).map(function (no) { return found[no]; });
}

/** 모든 필수 라인이 O인 주문만 한 번에 출고 처리한다. */
function completePickingOrders_() {
  var 라인 = readTable_(ROLE.라인), 주문 = readTable_(ROLE.주문), 헤더 = readTable_(ROLE.헤더), 마스터 = readTable_(ROLE.마스터);
  var L = {
    주문번호: col_(라인, COL.주문번호, false), 품목별: col_(라인, COL.품목별주문번호, true),
    상품코드: col_(라인, COL.상품코드, true), 필요: col_(라인, COL.필요수량, true),
    확인: col_(라인, COL.확인, true), 실제: col_(라인, COL.실제수량, true),
    담당자: col_(라인, COL.담당자, false), 지시: col_(라인, COL.피킹지시번호, true),
    상태: col_(라인, COL.라인상태, true), 처리일시: col_(라인, COL.처리일시, true)
  };
  var O = { 주문번호: col_(주문, COL.주문번호, true), 품목별: col_(주문, COL.품목별주문번호, true),
    상태: col_(주문, COL.주문상태, true), 출고: col_(주문, COL.출고완료, true) };
  var H = { 주문번호: col_(헤더, COL.주문번호, true), 담당자: col_(헤더, COL.피킹담당자, true),
    상태: col_(헤더, COL.상태, true), 지시: col_(헤더, COL.피킹지시번호, true) };
  var M = { 코드: col_(마스터, COL.상품품목코드, true), 예약: col_(마스터, COL.예약재고, true), 가용: col_(마스터, COL.가용재고, true) };
  var orderState = {}, orderRows = {}, itemOrder = {};
  주문.rows.forEach(function (row, idx) {
    var no = toStr_(row[O.주문번호]); orderState[no] = toStr_(row[O.상태]);
    (orderRows[no] = orderRows[no] || []).push(idx); itemOrder[toStr_(row[O.품목별])] = no;
  });
  var headerByOrder = {};
  헤더.rows.forEach(function (row, idx) { headerByOrder[toStr_(row[H.주문번호])] = idx; });
  var groups = {};
  라인.rows.forEach(function (row, idx) {
    var no = (L.주문번호 >= 0 ? toStr_(row[L.주문번호]) : '') || itemOrder[toStr_(row[L.품목별])];
    if (no) (groups[no] = groups[no] || []).push(idx);
  });
  var masterIndex = {}, reserved = [];
  마스터.rows.forEach(function (row, idx) { masterIndex[toStr_(row[M.코드])] = idx; reserved[idx] = toNum_(row[M.예약]); });
  var now = new Date(), logs = [], orders = [], lines = 0, quantity = 0, propagated = 0;

  Object.keys(groups).forEach(function (no) {
    if (orderState[no] !== ENUM.주문상태.처리완료) return;
    var indexes = groups[no];
    if (!indexes.length || !indexes.every(function (idx) {
      var row = 라인.rows[idx];
      return toStr_(row[L.상태]) === ENUM.라인상태.완료 || toStr_(row[L.확인]).toUpperCase() === ENUM.확인.정상;
    })) return;
    var required = {}, pending = [];
    indexes.forEach(function (idx) {
      var row = 라인.rows[idx];
      if (toStr_(row[L.상태]) === ENUM.라인상태.완료) return;
      pending.push(idx); var code = toStr_(row[L.상품코드]);
      required[code] = (required[code] || 0) + toNum_(row[L.필요]);
    });
    if (!pending.length) return;
    var shortage = Object.keys(required).filter(function (code) {
      var mi = masterIndex[code]; return mi === undefined || reserved[mi] < required[code];
    });
    if (shortage.length) { writeOpLog_('S5_1_결과반영', '경고', no + ' 예약재고 부족: ' + shortage.join(', ')); return; }
    Object.keys(required).forEach(function (code) { reserved[masterIndex[code]] -= required[code]; });
    var hIdx = headerByOrder[no], worker = hIdx === undefined ? '' : toStr_(헤더.rows[hIdx][H.담당자]);
    pending.forEach(function (idx) {
      var row = 라인.rows[idx], qty = toNum_(row[L.필요]), code = toStr_(row[L.상품코드]);
      row[L.실제] = qty; row[L.상태] = ENUM.라인상태.완료; row[L.처리일시] = now;
      if (L.담당자 >= 0 && worker && toStr_(row[L.담당자]) !== worker) { row[L.담당자] = worker; propagated++; }
      logs.push({ 구분: ENUM.로그구분.차감, 피킹지시번호: toStr_(row[L.지시]), 주문번호: no,
        품목별주문번호: toStr_(row[L.품목별]), 상품코드: code, 변동량: -qty,
        변동후재고: toNum_(마스터.rows[masterIndex[code]][M.가용]), 담당자: worker, 사유: '피킹 출고 · 예약재고 소진' });
      lines++; quantity += qty;
    });
    if (hIdx !== undefined) 헤더.rows[hIdx][H.상태] = ENUM.헤더상태.완료;
    (orderRows[no] || []).forEach(function (idx) { 주문.rows[idx][O.상태] = ENUM.주문상태.출고완료; 주문.rows[idx][O.출고] = 1; });
    orders.push(no);
  });
  if (라인.rows.length) {
    writeColumn_(라인.sheet, L.실제, 라인.rows); writeColumn_(라인.sheet, L.상태, 라인.rows);
    writeColumn_(라인.sheet, L.처리일시, 라인.rows); if (L.담당자 >= 0) writeColumn_(라인.sheet, L.담당자, 라인.rows);
  }
  if (헤더.rows.length) writeColumn_(헤더.sheet, H.상태, 헤더.rows);
  if (주문.rows.length) { writeColumn_(주문.sheet, O.상태, 주문.rows); writeColumn_(주문.sheet, O.출고, 주문.rows); }
  if (마스터.rows.length) 마스터.sheet.getRange(2, M.예약 + 1, reserved.length, 1).setValues(reserved.map(function (v) { return [v]; }));
  writeStockLog_(logs);
  return { lines: lines, quantity: quantity, orders: orders, propagated: propagated };
}

function S5_2_수동반영() {
  var r = S5_1_결과반영();
  alert_('피킹 결과 반영 완료\n출고완료 ' + r.완료주문.length + '건 / 취소 ' + r.취소주문.length + '건');
}

/** 결과 반영과 대시보드 갱신을 한 트리거에서 수행한다. */
function syncAndRefresh() {
  try { S5_1_결과반영(); }
  finally { try { D0_대시보드전체갱신(true); } catch (e) { writeOpLog_('D0_대시보드전체갱신', '실패', e.message); } }
}
