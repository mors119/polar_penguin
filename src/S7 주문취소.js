var CANCELLATION_REASONS = ['고객 요청', '중복 주문', '주소 오류', '재고 오류', '판매자 취소', '기타'];

/** 주문 취소의 유일한 재고 복원 경로. 모든 호출자는 주문번호 전체를 넘긴다. */
function cancelOrder_(orderNo, reason, source, options) {
  var originalState = '확인 불가';
  try {
    var before = readTable_(ROLE.주문), beforeNo = col_(before, COL.주문번호, true), beforeState = col_(before, COL.주문상태, true);
    before.rows.some(function (row) {
      if (toStr_(row[beforeNo]) !== toStr_(orderNo)) return false;
      originalState = toStr_(row[beforeState]) || originalState; return true;
    });
    return cancelOrderCore_(orderNo, reason, source, options);
  } catch (e) {
    sendSystemNotification_('ERROR', '주문 취소 실패', {
      주문번호: toStr_(orderNo), 원래상태: originalState, 시도작업: toStr_(source) || 'SYSTEM',
      오류: e.message, 재고변경여부: '확인 필요 — 작업로그와 재고이동로그를 확인하세요.'
    });
    throw e;
  }
}

function cancelOrderCore_(orderNo, reason, source, options) {
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
      return { 취소: false, 확인필요: true, 메시지: '이미 피킹지시서가 출력되어 출고 처리된 주문입니다.\n취소하면 해당 출고 수량을 가용재고로 복원합니다.\n계속하시겠습니까?' };
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
    var M = { 코드: col_(마스터, COL.상품품목코드, true), 가용: col_(마스터, COL.가용재고, true) };
    var index = {}, available = [];
    마스터.rows.forEach(function (row, idx) {
      index[toStr_(row[M.코드])] = idx; available[idx] = toNum_(row[M.가용]);
    });
    var logs = [], actor = 사용자_(), instruction = toStr_(items[0].row[col_(주문, COL.피킹지시번호, true)]);
    var reservationWasMade = state === '처리완료' ||
      (state === ENUM.주문상태.예약 && O.확정일시 >= 0 && !isBlank_(items[0].row[O.확정일시]));
    Object.keys(required).forEach(function (code) {
      var mi = index[code];
      if (mi === undefined) return;
      var restore = 0;
      if (state === ENUM.주문상태.출고완료) {
        restore = required[code];
      } else if (reservationWasMade) {
        restore = required[code];
      }
      if (!restore) return;
      available[mi] += restore;
      logs.push({ 구분: ENUM.로그구분.복원, 피킹지시번호: instruction, 주문번호: orderNo, 품목별주문번호: '',
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
    }
    closePickingForCancellation_(orderNo, now);
    writeStockLog_(logs);
    writeOpLog_('cancelOrder_', '성공', orderNo + ' / ' + source + ' / ' + reason);
    if (options.refresh !== false) try { D0_대시보드전체갱신(true); } catch (ignore) { }
    return {
      취소: true,
      주문번호: orderNo,
      복원건: logs.length,
      복원수량: logs.reduce(function (sum, log) { return sum + toNum_(log.변동량); }, 0)
    };
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
  var template = HtmlService.createTemplateFromFile('CancellationScope');
  template.orderNoJson = JSON.stringify(orderNo).replace(/</g, '\\u003c');
  ui.showModalDialog(template.evaluate().setWidth(620).setHeight(680), '주문 취소');
}

function normalizeCancellationRecipientName_(value) {
  return toStr_(value).replace(/[\s\u00A0]/g, '').toLowerCase();
}

function normalizeCancellationPhone_(value) {
  return toStr_(value).replace(/\D/g, '');
}

function maskCancellationPhone_(value) {
  var normalized = normalizeCancellationPhone_(value);
  return normalized ? '***-****-' + normalized.slice(-4) : '';
}

/** 주문 테이블을 주문번호별로 한 번만 묶어 취소 화면용 요약을 만든다. */
function cancellationOrdersFromTable_(table) {
  var O = {
    no: col_(table, COL.주문번호, true), state: col_(table, COL.주문상태, true),
    quantity: col_(table, COL.수량, true), recipient: col_(table, COL.수령인, true),
    phone: col_(table, COL.수령인휴대전화, true), sku: col_(table, COL.상품품목코드, true),
    name: col_(table, COL.주문상품명, false), option: col_(table, COL.상품옵션기본, false)
  };
  var byNo = {}, ordered = [];
  table.rows.forEach(function (row) {
    var no = toStr_(row[O.no]);
    if (!no) return;
    if (!byNo[no]) {
      byNo[no] = {
        orderNo: no, state: toStr_(row[O.state]), recipientName: toStr_(row[O.recipient]),
        recipientKey: normalizeCancellationRecipientName_(row[O.recipient]),
        phoneKey: normalizeCancellationPhone_(row[O.phone]), phoneMasked: maskCancellationPhone_(row[O.phone]),
        totalQuantity: 0, items: [], _items: {}
      };
      ordered.push(byNo[no]);
    }
    var order = byNo[no], quantity = toNum_(row[O.quantity]);
    order.totalQuantity += quantity;
    var product = O.name >= 0 ? toStr_(row[O.name]) : '';
    var option = O.option >= 0 ? toStr_(row[O.option]) : '';
    var label = product || toStr_(row[O.sku]) || '상품';
    if (option) label += ' / ' + option;
    if (!order._items[label]) {
      order._items[label] = { name: label, quantity: 0 };
      order.items.push(order._items[label]);
    }
    order._items[label].quantity += quantity;
  });
  ordered.forEach(function (order) { delete order._items; });
  return { ordered: ordered, byNo: byNo };
}

function cancellationOrderPreview_(order) {
  return {
    orderNo: order.orderNo,
    state: order.state,
    totalQuantity: order.totalQuantity,
    items: order.items.map(function (item) { return { name: item.name, quantity: item.quantity }; })
  };
}

function cancellationSummary_(orders) {
  var summary = { orderCount: orders.length, totalQuantity: 0, waitingOrderCount: 0,
    completedOrderCount: 0, cancelledOrderCount: 0, orders: orders.map(cancellationOrderPreview_) };
  orders.forEach(function (order) {
    summary.totalQuantity += order.totalQuantity;
    if (order.state === ENUM.주문상태.출고완료) summary.completedOrderCount++;
    else if (order.state === ENUM.주문상태.취소) summary.cancelledOrderCount++;
    else summary.waitingOrderCount++;
  });
  return summary;
}

function cancellationContextFromTable_(table, orderNo) {
  orderNo = toStr_(orderNo);
  var grouped = cancellationOrdersFromTable_(table), selected = grouped.byNo[orderNo];
  if (!selected) throw new Error('주문번호를 찾을 수 없습니다: ' + orderNo);
  var phoneAvailable = !!selected.phoneKey;
  var related = grouped.ordered.filter(function (order) {
    if (order.state === ENUM.주문상태.취소 || !phoneAvailable) return false;
    return order.recipientKey === selected.recipientKey && order.phoneKey === selected.phoneKey;
  });
  var selectedActive = selected.state !== ENUM.주문상태.취소;
  var safeSelected = cancellationOrderPreview_(selected);
  var safeRelated = related.map(cancellationOrderPreview_);
  var single = cancellationSummary_([selected]);
  var bulk = cancellationSummary_(related);
  return {
    selectedOrderNo: orderNo,
    selectedOrder: safeSelected,
    recipient: { name: selected.recipientName, phoneMasked: selected.phoneMasked },
    singleSummary: single,
    relatedOrders: safeRelated,
    bulkSummary: bulk,
    bulkAvailable: selectedActive && related.length >= 2,
    phoneAvailable: phoneAvailable,
    note: phoneAvailable ? '' : '수령인 휴대전화가 없어 동일 수령인 주문 전체 취소를 제공하지 않습니다.',
    reasons: CANCELLATION_REASONS.slice(),
    hasCompletedSingle: selected.state === ENUM.주문상태.출고완료,
    hasCompletedBulk: related.some(function (order) { return order.state === ENUM.주문상태.출고완료; })
  };
}

function getCancellationContext_(orderNo) {
  return cancellationContextFromTable_(readTable_(ROLE.주문), orderNo);
}

/** HtmlService에서 호출하는 공개 RPC. */
function getCancellationContext(orderNo) {
  return getCancellationContext_(orderNo);
}

/** 브라우저가 보낸 주문 목록은 받지 않고 현재 주문표에서 취소 대상을 다시 계산한다. */
function executeCancellationScope_(payload) {
  payload = payload || {};
  var selectedOrderNo = toStr_(payload.selectedOrderNo);
  var scope = toStr_(payload.scope).toUpperCase();
  var reason = toStr_(payload.reason);
  if (!selectedOrderNo) throw new Error('선택 주문번호가 없습니다.');
  if (scope !== 'SINGLE' && scope !== 'RECIPIENT') throw new Error('올바른 취소 범위를 선택하세요.');
  if (CANCELLATION_REASONS.indexOf(reason) < 0) throw new Error('올바른 취소 사유를 선택하세요.');

  // cancelOrder_() owns the mutation lock. Holding another scope-level lock here
  // would require nested ScriptLock acquisition for every target.
  var context = cancellationContextFromTable_(readTable_(ROLE.주문), selectedOrderNo);
  if (scope === 'RECIPIENT' && !context.phoneAvailable) {
    throw new Error('수령인 휴대전화가 없어 동일 수령인 주문 전체 취소를 실행할 수 없습니다.');
  }
  var targets = scope === 'SINGLE'
    ? [selectedOrderNo]
    : context.relatedOrders.map(function (order) { return order.orderNo; });
  if (!targets.length) targets = [selectedOrderNo];
  var selectedSummary = scope === 'SINGLE' ? context.singleSummary : context.bulkSummary;
  if (selectedSummary.completedOrderCount > 0 && payload.confirmCompleted !== true) {
    throw new Error('출고완료 주문 ' + selectedSummary.completedOrderCount + '건의 재고 복원 확인이 필요합니다. 취소 대상을 다시 확인하세요.');
  }
  var source = scope === 'SINGLE' ? 'MENU_SINGLE_CANCEL' : 'MENU_RECIPIENT_BULK_CANCEL';
  var summary = { requested: targets.length, success: 0, alreadyCancelled: 0, failed: [], restoredQuantity: 0 };
  targets.forEach(function (orderNo) {
    try {
      var result = cancelOrder_(orderNo, reason, source, {
        confirmReturn: payload.confirmCompleted === true, refresh: false
      });
      if (result.취소) {
        summary.success++;
        summary.restoredQuantity += toNum_(result.복원수량);
      } else if (result.이미취소) {
        summary.alreadyCancelled++;
      } else {
        summary.failed.push({ orderNo: orderNo, reason: result.메시지 || '취소되지 않았습니다.' });
      }
    } catch (e) {
      summary.failed.push({ orderNo: orderNo, reason: e.message || String(e) });
    }
  });
  try { D0_대시보드전체갱신(true); } catch (ignore) { }
  writeOpLog_('executeCancellationScope_', summary.failed.length ? '부분실패' : '성공',
    '선택주문=' + selectedOrderNo + ' / 범위=' + scope + ' / 대상=' + summary.requested +
    ' / 성공=' + summary.success + ' / 실패=' + summary.failed.length + ' / 사유=' + reason);
  return summary;
}

/** HtmlService에서 호출하는 공개 RPC. */
function executeCancellationScope(payload) {
  return executeCancellationScope_(payload);
}

/** HtmlService / 외부 호출 호환 공개 RPC */
function executeCancellation(payload) {
  return executeCancellationScope_(payload);
}
