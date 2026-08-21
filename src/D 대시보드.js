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
  if (out.예약출고가능) out.권고 = { 제목: '예약 주문 피킹 가능', 내용: [out.예약출고가능SKU + '개 SKU에서 FIFO 기준 ' + out.예약출고가능 + '건을 출고할 수 있습니다.', '메뉴의 예약상품 입고 관리를 여세요.'], 색: DASHCOLOR.경고 };
  return out;
}

function collectStockStatus_() {
  var out = { 상품수: 0, 총가용: 0, 품절: 0, 음수허용: 0, 위치미지정: 0, 부족: [], 예약부족SKU: [] };
  var table = readTable_(ROLE.마스터);
  var C = { 코드: col_(table, COL.상품품목코드, true), 상품명: col_(table, COL.상품명, true),
    가용: col_(table, COL.가용재고, true), 관리: col_(table, COL.재고관리, true),
    위치: col_(table, COL.기본보관위치, true) };
  var threshold = Number(param_('재고경고임계치', 3));
  table.rows.forEach(function (row) {
    var code = toStr_(row[C.코드]); if (!code) return;
    var available = toNum_(row[C.가용]), managed = toStr_(row[C.관리]).toUpperCase() !== 'F';
    out.상품수++; out.총가용 += available;
    if (!toStr_(row[C.위치])) out.위치미지정++;
    if (managed && available <= 0) out.품절++;
    if (!managed && available < 0) out.음수허용++;
    if (managed && available <= threshold) out.부족.push({ 코드: code, 상품명: toStr_(row[C.상품명]), 가용: available });
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
  var out = { latestTime: '기록 없음', latestResult: '기록 없음', warning: '없음', recentErrors: [], lastBackup: '백업 없음' };
  try {
    var backup = PropertiesService.getScriptProperties().getProperty('최근백업일시');
    if (backup) out.lastBackup = formatMaintenanceDate_(new Date(backup));
  } catch (ignore) { }
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
  var sh = ensureDashSheet_(ss, '📊 대시보드');
  dashboardResetSheet_(sh);
  try {
    return dashboardRenderLayout_(sh, order, stock, picking, operation);
  } catch (e) {
    throw new Error('[렌더링] ' + e.message);
  }
}

/** 기존 부분 렌더링을 포함해 시트 전체의 병합과 표시 상태를 초기화한다. */
function dashboardResetSheet_(sheet) {
  try {
    if (sheet.getMaxRows() < 30) sheet.insertRowsAfter(sheet.getMaxRows(), 30 - sheet.getMaxRows());
    if (sheet.getMaxColumns() < 12) sheet.insertColumnsAfter(sheet.getMaxColumns(), 12 - sheet.getMaxColumns());
  } catch (e) {
    throw new Error('[시트 크기 확보] ' + e.message);
  }
  try {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  } catch (e2) {
    throw new Error('[병합 초기화] ' + e2.message);
  }
  try {
    sheet.clearContents();
    sheet.clearFormats();
  } catch (e3) {
    throw new Error('[시트 초기화] ' + e3.message);
  }
}

/** 대시보드 병합은 이 함수에서만 수행해 범위를 일관되게 관리한다. */
function dashboardMerge_(sheet, row, column, numRows, numColumns) {
  return sheet.getRange(row, column, numRows, numColumns).merge();
}

function dashboardRenderLayout_(sh, order, stock, picking, operation) {
  dashboardMerge_(sh, 1, 1, 1, 12).setValue('📊  Polar Penguin 통합 대시보드')
    .setFontSize(18).setFontWeight('bold').setFontColor('FFFFFF').setBackground(DASHCOLOR.제목)
    .setVerticalAlignment('middle');
  dashboardMerge_(sh, 2, 1, 1, 12).setValue('최근 갱신 ' + Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss') +
    '  ·  최근 백업 ' + (operation.lastBackup || '백업 없음'));
  var pendingOrders = Math.max(0, Number(order.전체 || 0) - Number(order.출고완료 || 0) - Number(order.취소 || 0));
  var errorCount = Number(picking.kpi.출력오류 || 0) + (operation.recentErrors || []).length;
  var kpis = [
    ['대기 주문', pendingOrders, pendingOrders ? DASHCOLOR.경고 : DASHCOLOR.카드],
    ['피킹 대기', picking.작업대상주문, picking.작업대상주문 ? DASHCOLOR.경고 : DASHCOLOR.카드],
    ['출고 완료', order.출고완료, DASHCOLOR.좋음],
    ['예약 주문', order.예약전체, order.예약전체 ? DASHCOLOR.경고 : DASHCOLOR.카드],
    ['재고 경고', stock.부족.length, stock.부족.length ? DASHCOLOR.경고 : DASHCOLOR.카드],
    ['처리 오류', errorCount, errorCount ? DASHCOLOR.위험 : DASHCOLOR.좋음]
  ];
  kpis.forEach(function (item, index) { dashboardKpiCard_(sh, 4, index * 2 + 1, item[0], item[1], item[2]); });

  dashboardDetailSection_(sh, 8, 1, '처리 필요', [
    ['대기 주문', pendingOrders], ['피킹 대기', picking.작업대상주문],
    ['오늘 생성 피킹지시', picking.오늘지시], ['오늘 출고 수량', picking.오늘출고수량],
    ['위치 미지정', stock.위치미지정]
  ]);
  dashboardDetailSection_(sh, 8, 7, '재고 경고', [
    ['경고 SKU', stock.부족.length], ['품절(T)', stock.품절], ['음수 순재고(F)', stock.음수허용],
    ['예약 부족 SKU', stock.예약부족SKU.length], ['총 가용재고', stock.총가용]
  ]);
  dashboardDetailSection_(sh, 15, 1, '예약 대기', [
    ['예약 주문', order.예약전체], ['예약 수량', order.예약수량], ['출고 가능 주문', order.예약출고가능],
    ['가용 예약 SKU', order.예약출고가능SKU], ['재고 부족 주문', order.예약부족]
  ]);
  dashboardDetailSection_(sh, 15, 7, '최근 처리', [
    ['처리 시각', operation.latestTime], ['처리 결과', operation.latestResult], ['입력 경고', operation.warning],
    ['전체 주문', order.전체], ['취소 주문', order.취소]
  ]);

  dashboardMerge_(sh, 22, 1, 1, 12).setValue('운영 알림').setFontWeight('bold')
    .setBackground(DASHCOLOR.제목).setFontColor('FFFFFF');
  dashboardMerge_(sh, 23, 1, 1, 12).setValue(operation.recentErrors.join(' / ') || '최근 처리 오류 없음').setWrap(true)
    .setBackground(errorCount ? DASHCOLOR.위험 : DASHCOLOR.좋음);
  dashboardMerge_(sh, 25, 1, 1, 12).setValue('권고: ' + order.권고.제목 + ' — ' + order.권고.내용.join(' / '))
    .setWrap(true).setBackground(order.권고.색);
  dashboardMerge_(sh, 27, 1, 1, 12).setValue('빠른 작업 · 셀은 버튼이 아닙니다. 상단 📦 Polar Penguin 메뉴에서 실행하세요.')
    .setFontWeight('bold').setBackground(DASHCOLOR.헤더);
  [
    [1, 'Drive Input'], [4, 'Input 지금 처리'],
    [7, '예약상품 입고 관리'], [10, '피킹지시서 재출력']
  ].forEach(function (action) {
    dashboardMerge_(sh, 28, action[0], 1, 3).setValue(action[1])
      .setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  for (var c = 1; c <= 12; c++) sh.setColumnWidth(c, 78);
  sh.setRowHeight(1, 38); sh.setRowHeight(4, 30); sh.setRowHeight(5, 48);
  [8, 15, 22, 27].forEach(function (row) { sh.setRowHeight(row, 28); });
  sh.setRowHeight(23, 40); sh.setRowHeight(25, 42); sh.setRowHeight(28, 34);
  sh.setFrozenRows(2); sh.setHiddenGridlines(true); return sh;
}

function dashboardKpiCard_(sheet, row, column, label, value, color) {
  dashboardMerge_(sheet, row, column, 1, 2).setValue(label).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(DASHCOLOR.헤더);
  dashboardMerge_(sheet, row + 1, column, 1, 2).setValue(value).setFontSize(19).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setBackground(color);
}

function dashboardDetailSection_(sheet, row, column, title, items) {
  dashboardMerge_(sheet, row, column, 1, 6).setValue(title).setFontWeight('bold')
    .setFontColor('FFFFFF').setBackground(DASHCOLOR.제목).setVerticalAlignment('middle');
  items.forEach(function (item, index) {
    var itemRow = row + index + 1;
    dashboardMerge_(sheet, itemRow, column, 1, 4).setValue(item[0]).setBackground(DASHCOLOR.카드)
      .setVerticalAlignment('middle');
    dashboardMerge_(sheet, itemRow, column + 4, 1, 2).setValue(item[1]).setFontWeight('bold')
      .setHorizontalAlignment(typeof item[1] === 'number' ? 'right' : 'left').setVerticalAlignment('middle');
  });
}
