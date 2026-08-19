/** S8. 예약 주문은 가용성만 자동 계산하고 release 결정은 운영자가 한다. */
function collectPreorderData_() {
  var 주문 = readTable_(ROLE.주문), 마스터 = readTable_(ROLE.마스터);
  var O = { 주문번호: col_(주문, COL.주문번호, true), 상품코드: col_(주문, COL.상품품목코드, true),
    수량: col_(주문, COL.수량, true), 상태: col_(주문, COL.주문상태, true) };
  var M = { 코드: col_(마스터, COL.상품품목코드, true), 상품명: col_(마스터, COL.상품명, true),
    옵션: col_(마스터, COL.옵션명, false), 가용: col_(마스터, COL.가용재고, true) };
  var stock = {};
  마스터.rows.forEach(function (row) { stock[toStr_(row[M.코드])] = { 가용: toNum_(row[M.가용]), 상품명: toStr_(row[M.상품명]), 옵션: M.옵션 >= 0 ? toStr_(row[M.옵션]) : '' }; });
  var groups = {};
  주문.rows.forEach(function (row) {
    if (toStr_(row[O.상태]) !== ENUM.주문상태.예약) return;
    var no = toStr_(row[O.주문번호]), code = toStr_(row[O.상품코드]);
    if (!no || !code) return;
    var group = groups[no] = groups[no] || {};
    group[code] = (group[code] || 0) + toNum_(row[O.수량]);
  });
  var out = { 전체: 0, 출고가능: 0, 재고부족: 0, 주문: [], 상품: [], 주문수: 0, 총수량: 0, 총부족: 0, 주문목록: [] };
  var productTotals = {};
  Object.keys(groups).sort().forEach(function (no) {
    var shortages = [], total = 0;
    Object.keys(groups[no]).forEach(function (code) {
      var qty = groups[no][code], available = stock[code] ? stock[code].가용 : 0;
      total += qty; productTotals[code] = (productTotals[code] || 0) + qty;
      if (available < qty) shortages.push({ 코드: code, 필요: qty, 가용: available, 부족: qty - available });
    });
    var releasable = shortages.length === 0;
    out.주문.push({ 주문번호: no, 수량: total, 출고가능: releasable, 부족: shortages });
    out.전체++; out.총수량 += total;
    if (releasable) out.출고가능++; else out.재고부족++;
    shortages.forEach(function (s) { out.총부족 += s.부족; });
  });
  Object.keys(productTotals).forEach(function (code) {
    var info = stock[code] || { 가용: 0, 상품명: '(미등록)', 옵션: '' }, qty = productTotals[code];
    out.상품.push({ 코드: code, 상품명: info.상품명, 옵션: info.옵션, 수량: qty,
      현재고: info.가용, 부족: qty - info.가용, 주문수: out.주문.filter(function (o) { return groups[o.주문번호][code]; }).length });
  });
  out.상품.sort(function (a, b) { return Math.max(b.부족, 0) - Math.max(a.부족, 0); });
  out.주문수 = out.전체; out.주문목록 = out.주문.map(function (o) { return o.주문번호; });
  return out;
}

function 예약_주문피킹서생성() {
  var range = SpreadsheetApp.getActiveRange();
  if (!range || range.getSheet().getName() !== ROLE.주문 || range.getRow() < 2) {
    alert_('주문(완료)에서 release할 예약 주문의 셀을 선택하세요.'); return;
  }
  var table = readTable_(ROLE.주문), cNo = col_(table, COL.주문번호, true);
  var orderNo = toStr_(range.getSheet().getRange(range.getRow(), cNo + 1).getValue());
  return releaseReservationOrder_(orderNo);
}

function releaseReservationOrder_(orderNo) {
  return withLock_(function () {
    var confirmation = S3_1_주문확정([orderNo], { manualRelease: true, silent: true });
    if (!confirmation.준비주문.length) {
      var detail = confirmation.부족목록.length ? confirmation.부족목록[0].상세.join('\n') : '예약 상태가 아니거나 이미 처리되었습니다.';
      alert_('예약 주문을 release하지 못했습니다.\n' + detail); return { 생성: false, 부족: true };
    }
    var picking = S4_1_피킹지시생성([orderNo], { silent: true });
    if (!picking.지시번호) throw new Error('피킹지시를 만들지 못했습니다: ' + orderNo);
    try {
      var pdf = S9_피킹PDF생성(picking.지시번호, inputFolder_('Output폴더ID'));
      markPickingOutputState_(picking.지시번호, ENUM.헤더상태.대기);
      markOrdersReady_([orderNo]); D0_대시보드전체갱신(true);
      alert_('예약 주문 피킹서가 준비되었습니다.\n' + orderNo + '\n' + picking.지시번호);
      return { 생성: true, 지시번호: picking.지시번호, pdf: pdf };
    } catch (e) {
      markPickingOutputState_(picking.지시번호, ENUM.헤더상태.출력오류);
      writeOpLog_('releaseReservationOrder_', '실패', orderNo + ' / ' + e.message); throw e;
    }
  });
}
