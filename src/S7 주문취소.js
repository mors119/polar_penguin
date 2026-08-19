/** 주문 취소의 유일한 재고 복원 경로. 모든 호출자는 주문번호 전체를 넘긴다. */
function cancelOrder_(orderNo, reason, source, options) {
  options = options || {};
  return withLock_(function () {
    orderNo = toStr_(orderNo); reason = toStr_(reason) || '기타'; source = toStr_(source) || 'SYSTEM';
    var 주문 = readTable_(ROLE.주문);
    var O = {
      주문번호: col_(주문, COL.주문번호, true), 품목별: col_(주문, COL.품목별주문번호, true),
      상품코드: col_(주문, COL.상품품목코드, true), 수량: col_(주문, COL.수량, true),
      상태: col_(주문, COL.주문상태, true), 출고: col_(주문, COL.출고완료, true),
      취소사유: col_(주문, COL.취소사유, true), 취소일시: col_(주문, COL.취소일시, true),
      취소경로: col_(주문, COL.취소경로, false), 확정일시: col_(주문, COL.확정일시, false),
      대기사유: col_(주문, COL.대기사유, false)
    };
    var items = [];
    주문.rows.forEach(function (row, idx) { if (toStr_(row[O.주문번호]) === orderNo) items.push({ row: row, idx: idx }); });
    if (!items.length) throw new Error('주문번호를 찾을 수 없습니다: ' + orderNo);
    var state = toStr_(items[0].row[O.상태]);
    if (state === ENUM.주문상태.취소) return { 취소: false, 이미취소: true, 메시지: '이미 취소된 주문입니다.' };
    if (state === ENUM.주문상태.출고완료 && !options.confirmReturn) {
      return { 취소: false, 확인필요: true, 메시지: '이미 출고완료된 주문입니다. 취소하면 출고 수량을 가용재고로 복원합니다. 실제 상품이 창고에 반환되었는지 확인하세요.' };
    }

    var required = {};
    items.forEach(function (item) {
      var code = toStr_(item.row[O.상품코드]);
      required[code] = (required[code] || 0) + toNum_(item.row[O.수량]);
    });
    if (state === ENUM.주문상태.출고완료) {
      var shipped = shippedQuantityBySku_(orderNo, items, O);
      if (Object.keys(shipped).length) required = shipped;
    }
    var 마스터 = readTable_(ROLE.마스터);
    var M = { 코드: col_(마스터, COL.상품품목코드, true), 가용: col_(마스터, COL.가용재고, true), 예약: col_(마스터, COL.예약재고, true) };
    var index = {}, available = [], reserved = [];
    마스터.rows.forEach(function (row, idx) {
      index[toStr_(row[M.코드])] = idx; available[idx] = toNum_(row[M.가용]); reserved[idx] = toNum_(row[M.예약]);
    });
    var logs = [], actor = 사용자_(), instruction = toStr_(items[0].row[col_(주문, COL.피킹지시번호, true)]);
    var reservationWasMade = state === ENUM.주문상태.처리완료 ||
      (state === ENUM.주문상태.예약 && O.확정일시 >= 0 && !isBlank_(items[0].row[O.확정일시]));
    Object.keys(required).forEach(function (code) {
      var mi = index[code];
      if (mi === undefined) return;
      var restore = 0, type = ENUM.로그구분.복원;
      if (state === ENUM.주문상태.출고완료) {
        restore = required[code];
      } else if (reservationWasMade) {
        restore = Math.min(required[code], reserved[mi]);
        reserved[mi] -= restore;
        type = ENUM.로그구분.예약해제;
      }
      if (!restore) return;
      available[mi] += restore;
      logs.push({ 구분: type, 피킹지시번호: instruction, 주문번호: orderNo, 품목별주문번호: '',
        상품코드: code, 변동량: restore, 변동후재고: available[mi], 담당자: actor,
        사유: source + ' 취소 복원 · ' + reason });
    });

    var now = new Date();
    items.forEach(function (item) {
      item.row[O.상태] = ENUM.주문상태.취소; item.row[O.출고] = 0;
      item.row[O.취소사유] = reason; item.row[O.취소일시] = now;
      if (O.취소경로 >= 0) item.row[O.취소경로] = source;
      if (O.대기사유 >= 0) item.row[O.대기사유] = '';
    });
    writeColumn_(주문.sheet, O.상태, 주문.rows); writeColumn_(주문.sheet, O.출고, 주문.rows);
    writeColumn_(주문.sheet, O.취소사유, 주문.rows); writeColumn_(주문.sheet, O.취소일시, 주문.rows);
    if (O.취소경로 >= 0) writeColumn_(주문.sheet, O.취소경로, 주문.rows);
    if (O.대기사유 >= 0) writeColumn_(주문.sheet, O.대기사유, 주문.rows);
    if (마스터.rows.length) {
      마스터.sheet.getRange(2, M.가용 + 1, available.length, 1).setValues(available.map(function (v) { return [v]; }));
      마스터.sheet.getRange(2, M.예약 + 1, reserved.length, 1).setValues(reserved.map(function (v) { return [v]; }));
    }
    closePickingForCancellation_(orderNo, now);
    writeStockLog_(logs);
    writeOpLog_('cancelOrder_', '성공', orderNo + ' / ' + source + ' / ' + reason);
    if (options.refresh !== false) try { D0_대시보드전체갱신(true); } catch (ignore) { }
    return { 취소: true, 주문번호: orderNo, 복원건: logs.length };
  });
}

function shippedQuantityBySku_(orderNo, orderItems, O) {
  var itemCode = {};
  orderItems.forEach(function (item) { itemCode[toStr_(item.row[O.품목별])] = toStr_(item.row[O.상품코드]); });
  var lines = readTable_(ROLE.라인), cItem = col_(lines, COL.품목별주문번호, true), cActual = col_(lines, COL.실제수량, true);
  var result = {};
  lines.rows.forEach(function (row) {
    var itemNo = toStr_(row[cItem]), code = itemCode[itemNo]; if (!code) return;
    var qty = toNum_(row[cActual]); if (qty > 0) result[code] = (result[code] || 0) + qty;
  });
  return result;
}

function closePickingForCancellation_(orderNo, now) {
  var 헤더 = readTable_(ROLE.헤더);
  var hNo = col_(헤더, COL.주문번호, true), hState = col_(헤더, COL.상태, true), instructions = {};
  헤더.rows.forEach(function (row) {
    if (toStr_(row[hNo]) !== orderNo) return;
    row[hState] = ENUM.헤더상태.취소;
    instructions[toStr_(row[col_(헤더, COL.피킹지시번호, true)])] = true;
  });
  if (헤더.rows.length) writeColumn_(헤더.sheet, hState, 헤더.rows);
  var 주문 = readTable_(ROLE.주문), itemOrder = {}, oItem = col_(주문, COL.품목별주문번호, true), oNo = col_(주문, COL.주문번호, true);
  주문.rows.forEach(function (row) { itemOrder[toStr_(row[oItem])] = toStr_(row[oNo]); });
  var 라인 = readTable_(ROLE.라인);
  var lNo = col_(라인, COL.주문번호, false), lItem = col_(라인, COL.품목별주문번호, true);
  var lState = col_(라인, COL.라인상태, true), lTime = col_(라인, COL.처리일시, true);
  라인.rows.forEach(function (row) {
    var lineOrder = (lNo >= 0 ? toStr_(row[lNo]) : '') || itemOrder[toStr_(row[lItem])];
    if (lineOrder === orderNo) {
      row[lState] = ENUM.라인상태.취소; row[lTime] = now;
    }
  });
  if (라인.rows.length) { writeColumn_(라인.sheet, lState, 라인.rows); writeColumn_(라인.sheet, lTime, 라인.rows); }
}

function 선택_주문취소() {
  var ui = SpreadsheetApp.getUi(), range = SpreadsheetApp.getActiveRange();
  if (!range || range.getSheet().getName() !== ROLE.주문 || range.getRow() < 2) {
    alert_('주문(완료)에서 취소할 주문의 셀을 선택하세요.'); return;
  }
  var table = readTable_(ROLE.주문), cNo = col_(table, COL.주문번호, true);
  var orderNo = toStr_(range.getSheet().getRange(range.getRow(), cNo + 1).getValue());
  if (!orderNo) { alert_('선택한 행에 주문번호가 없습니다.'); return; }
  var prompt = ui.prompt('선택 주문 취소', '사유를 입력하세요: 고객 요청 / 중복 주문 / 주소 오류 / 재고 오류 / 판매자 취소 / 기타', ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() !== ui.Button.OK) return;
  var reason = toStr_(prompt.getResponseText()) || '기타';
  if (ui.alert('주문 전체 취소', orderNo + '\n사유: ' + reason + '\n주문 전체를 취소할까요?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var result = cancelOrder_(orderNo, reason, 'MANUAL', {});
  if (result.확인필요) {
    if (ui.alert('출고완료 주문 취소', result.메시지, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
    result = cancelOrder_(orderNo, reason, 'MANUAL_RETURN', { confirmReturn: true });
  }
  alert_(result.메시지 || (result.취소 ? '주문 전체를 취소했습니다.' : '변경 사항이 없습니다.'));
}
