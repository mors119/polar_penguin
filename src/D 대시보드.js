/** 통합 대시보드는 운영자가 지금 알아야 할 상태와 행동만 보여 준다. */
function collectOrderStatus_() {
  var out = { 예약: 0, 출고완료: 0, 취소: 0, 전체: 0,
    예약전체: 0, 예약수량: 0, 예약출고가능: 0, 예약출고가능SKU: 0, 예약부족: 0, 예약목록: [],
    피킹: { 속도: 0, 예상분: null }, 권고: { 제목: '정상 운영', 내용: ['PDF 출력 성공과 동시에 재고와 출고가 반영됩니다.'], 색: DASHCOLOR.좋음 } };
  var table = readTable_(ROLE.주문), cNo = col_(table, COL.주문번호, true), cState = col_(table, COL.주문상태, true), seen = {};
  table.rows.forEach(function (row) {
    var no = toStr_(row[cNo]); if (!no || seen[no]) return; seen[no] = true; out.전체++;
    var state = toStr_(row[cState]); if (out[state] !== undefined) out[state]++;
  });
  var reservation = collectPreorderData_();
  out.예약전체 = reservation.전체; out.예약수량 = reservation.총수량;
  out.예약출고가능 = reservation.출고가능; out.예약출고가능SKU = reservation.출고가능SKU; out.예약부족 = reservation.재고부족;
  out.예약목록 = reservation.주문;
  if (out.예약출고가능) out.권고 = { 제목: '예약상품 피킹 가능', 내용: [out.예약출고가능SKU + '개 SKU에서 FIFO 기준 ' + out.예약출고가능 + '건을 출고할 수 있습니다.', '메뉴의 예약상품 피킹 관리를 여세요.'], 색: DASHCOLOR.경고 };
  return out;
}

function collectStockStatus_() {
  var out = { 상품수: 0, 총가용: 0, 총예약: 0, 총불량: 0, 품절: 0, 부족: [], 예약부족SKU: [] };
  var table = readTable_(ROLE.마스터);
  var C = { 코드: col_(table, COL.상품품목코드, true), 상품명: col_(table, COL.상품명, true),
    가용: col_(table, COL.가용재고, true), 예약: col_(table, COL.예약재고, true), 불량: col_(table, COL.불량재고, false) };
  var threshold = Number(param_('재고경고임계치', 3));
  table.rows.forEach(function (row) {
    var code = toStr_(row[C.코드]); if (!code) return;
    var available = toNum_(row[C.가용]), reserved = toNum_(row[C.예약]), bad = C.불량 >= 0 ? toNum_(row[C.불량]) : 0;
    out.상품수++; out.총가용 += available; out.총예약 += reserved; out.총불량 += bad;
    if (available <= 0) out.품절++;
    if (available <= threshold) out.부족.push({ 코드: code, 상품명: toStr_(row[C.상품명]), 가용: available, 예약: reserved });
  });
  var reservation = collectPreorderData_();
  out.예약부족SKU = reservation.예약부족SKU || [];
  return out;
}

function collectPickingStatus_() {
  var out = { 전체라인: 0, 처리라인: 0, 미처리라인: 0, 진행률: 0, 작업대상주문: 0,
    오늘지시: 0, 오늘출고수량: 0,
    kpi: { 대기: 0, 완료: 0, 취소: 0, 출력오류: 0 } };
  var today = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
  var lines = readTable_(ROLE.라인), cState = col_(lines, COL.라인상태, true);
  var cQty = col_(lines, COL.실제수량, true), cTime = col_(lines, COL.처리일시, true);
  lines.rows.forEach(function (row) {
    var state = toStr_(row[cState]); if (!state) return; out.전체라인++;
    if (state === ENUM.라인상태.완료) {
      out.처리라인++;
      if (toStr_(row[cTime]) === today) out.오늘출고수량 += toNum_(row[cQty]);
    }
  });
  out.미처리라인 = out.전체라인 - out.처리라인;
  out.진행률 = out.전체라인 ? Math.round(out.처리라인 / out.전체라인 * 100) : 0;
  var headers = readTable_(ROLE.헤더), hState = col_(headers, COL.상태, true);
  var hInstruction = col_(headers, COL.피킹지시번호, true), hOutput = col_(headers, COL.출력일시, false), todayInstructions = {};
  headers.rows.forEach(function (row) {
    var state = toStr_(row[hState]) || ENUM.헤더상태.대기;
    if (out.kpi[state] !== undefined) out.kpi[state]++;
    if (state === ENUM.헤더상태.대기) out.작업대상주문++;
    if (hOutput >= 0 && toStr_(row[hOutput]) === today) todayInstructions[toStr_(row[hInstruction])] = true;
  });
  out.오늘지시 = Object.keys(todayInstructions).filter(String).length;
  return out;
}

function collectOperationStatus_() {
  var out = { latestTime: '기록 없음', latestResult: '기록 없음', warning: '없음', recentErrors: [] };
  var sheet = consoleSS_().getSheetByName(CONSOLE.입력처리로그);
  if (!sheet || sheet.getLastRow() < 2) return out;
  var rows = sheet.getRange(2, 1, Math.min(sheet.getLastRow() - 1, 20), INPUT_LOG_HEADERS.length).getValues();
  if (rows.length) {
    out.latestTime = rows[0][0] instanceof Date ? Utilities.formatDate(rows[0][0], tz_(), 'yyyy-MM-dd HH:mm') : toStr_(rows[0][0]);
    out.latestResult = toStr_(rows[0][5]) + ' · ' + toStr_(rows[0][2]);
  }
  rows.forEach(function (row) { if (toStr_(row[5]) === 'ERROR') out.recentErrors.push(toStr_(row[2]) + ': ' + toStr_(row[7])); });
  if (out.recentErrors.length) out.warning = out.recentErrors.length + '개 입력 오류';
  return out;
}

function D0_대시보드전체갱신(silent) {
  try {
    renderIntegratedDashboard_(consoleSS_(), collectOrderStatus_(), collectStockStatus_(), collectPickingStatus_(), collectOperationStatus_());
    if (!silent) alert_('대시보드 갱신 완료');
    return '✅ 통합 대시보드';
  } catch (e) {
    var message = '❌ 통합 대시보드: ' + e.message; if (!silent) alert_(message); return message;
  }
}

function renderIntegratedDashboard_(ss, order, stock, picking, operation) {
  var sh = ensureDashSheet_(ss, '📊 대시보드'); sh.clear(); sh.clearFormats();
  if (sh.getMaxColumns() < 10) sh.insertColumnsAfter(sh.getMaxColumns(), 10 - sh.getMaxColumns());
  sh.getRange(1, 1, 1, 10).merge().setValue('📊  Polar Penguin 통합 대시보드')
    .setFontSize(18).setFontWeight('bold').setFontColor('FFFFFF').setBackground(DASHCOLOR.제목);
  sh.getRange(2, 1, 1, 10).merge().setValue('최근 갱신 ' + Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss'));
  dashboardSection_(sh, 4, '입력', [['최근 처리', operation.latestTime], ['최근 결과', operation.latestResult], ['경고', operation.warning]]);
  dashboardSection_(sh, 9, '주문', [['예약', order.예약], ['출고완료', order.출고완료], ['취소', order.취소], ['전체', order.전체]]);
  dashboardSection_(sh, 14, '예약', [['예약 주문', order.예약전체], ['예약 수량', order.예약수량], ['출고 가능 주문', order.예약출고가능], ['가용 예약 SKU', order.예약출고가능SKU], ['재고 부족 주문', order.예약부족]]);
  dashboardSection_(sh, 19, '출력', [['오늘 생성된 피킹지시', picking.오늘지시], ['오늘 출고 처리 수량', picking.오늘출고수량], ['출력 오류', picking.kpi.출력오류]]);
  dashboardSection_(sh, 24, '재고', [['가용 재고', stock.총가용], ['예약 재고', stock.총예약], ['부족 상품', stock.부족.length], ['재고 경고', stock.품절], ['예약 부족 SKU', stock.예약부족SKU.length]]);
  sh.getRange(28, 1, 1, 10).merge().setValue('예약 부족 상위 SKU: ' +
    (stock.예약부족SKU.slice(0, 5).map(function (item) { return item.코드 + ' (' + item.주문수 + '건)'; }).join(' · ') || '없음'))
    .setWrap(true).setBackground(DASHCOLOR.카드);
  sh.getRange(30, 1, 1, 10).merge().setValue('최근 오류').setFontWeight('bold').setBackground(DASHCOLOR.제목).setFontColor('FFFFFF');
  sh.getRange(31, 1, 1, 10).merge().setValue(operation.recentErrors.join(' / ') || '없음').setWrap(true)
    .setBackground(operation.recentErrors.length ? DASHCOLOR.경고 : DASHCOLOR.좋음);
  sh.getRange(33, 1, 1, 10).merge().setValue('권고: ' + order.권고.제목 + ' — ' + order.권고.내용.join(' / ')).setWrap(true).setBackground(order.권고.색);
  sh.getRange(35, 1, 1, 10).merge().setValue('빠른 작업 (셀은 버튼이 아닙니다 — 상단 메뉴에서 실행)').setFontWeight('bold').setBackground(DASHCOLOR.경고);
  sh.getRange(36, 1, 1, 8).setValues([['파일 입력: Drive Input', '', 'Input 지금 처리', '', '예약상품 피킹 관리', '', '작업지시서 재출력', '']]).setFontWeight('bold');
  [1, 3, 5, 7].forEach(function (col) { sh.getRange(36, col, 1, 2).merge(); });
  for (var c = 1; c <= 10; c++) sh.setColumnWidth(c, c % 2 ? 120 : 90);
  sh.setFrozenRows(2); sh.setHiddenGridlines(true); return sh;
}

function dashboardSection_(sh, row, title, items) {
  sh.getRange(row, 1, 1, 10).merge().setValue(title).setFontWeight('bold').setFontColor('FFFFFF').setBackground(DASHCOLOR.제목);
  items.forEach(function (item, i) {
    var col = 1 + i * 2;
    sh.getRange(row + 1, col, 1, 2).merge().setValue(item[0]).setFontWeight('bold').setBackground(DASHCOLOR.헤더);
    sh.getRange(row + 2, col, 1, 2).merge().setValue(item[1]).setFontSize(16).setFontWeight('bold').setBackground(DASHCOLOR.카드);
  });
}
