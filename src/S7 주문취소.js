var CANCELLATION_REASONS = ['고객 요청', '중복 주문', '주소 오류', '재고 오류', '판매자 취소', '기타'];

/** 주문번호 전체 취소의 재고 복원 경로. */
function cancelOrder_(orderNo, reason, source, options) {
  return notifyCancellationFailure_(orderNo, source, function () {
    return cancelOrderCore_(orderNo, reason, source, options);
  });
}

function notifyCancellationFailure_(orderNo, source, action) {
  var originalState = '확인 불가';
  try {
    var before = readTable_(ROLE.주문), beforeNo = col_(before, COL.주문번호, true);
    var beforeState = col_(before, COL.주문상태, true);
    before.rows.some(function (row) {
      if (toStr_(row[beforeNo]) !== toStr_(orderNo)) return false;
      originalState = toStr_(row[beforeState]) || originalState;
      return true;
    });
    return action();
  } catch (e) {
    sendSystemNotification_('ERROR', '주문 취소 실패', {
      주문번호: toStr_(orderNo), 원래상태: originalState, 시도작업: toStr_(source) || 'SYSTEM',
      오류: e.message, 재고변경여부: '확인 필요 — 작업로그와 재고이동로그를 확인하세요.'
    });
    throw e;
  }
}

function cancellationOrderColumns_(table) {
  return {
    주문번호: col_(table, COL.주문번호, true), 품목별: col_(table, COL.품목별주문번호, true),
    상품코드: col_(table, COL.상품품목코드, true), 상품명: optionalCancellationCol_(table, COL.주문상품명),
    옵션: optionalCancellationCol_(table, COL.상품옵션기본), 수량: col_(table, COL.수량, true),
    상태: col_(table, COL.주문상태, true), 출고: col_(table, COL.출고완료, true),
    피킹지시: optionalCancellationCol_(table, COL.피킹지시번호), 취소사유: col_(table, COL.취소사유, true),
    취소일시: col_(table, COL.취소일시, true), 취소경로: optionalCancellationCol_(table, COL.취소경로),
    확정일시: optionalCancellationCol_(table, COL.확정일시), 대기사유: optionalCancellationCol_(table, COL.대기사유)
  };
}

function optionalCancellationCol_(table, canonical) {
  var index = table.headerIndex[normKey_(canonical)];
  return index === undefined ? -1 : index;
}

function orderRowsForCancellation_(table, O, orderNo) {
  var rows = [];
  table.rows.forEach(function (row, idx) {
    if (toStr_(row[O.주문번호]) === orderNo) rows.push({ row: row, idx: idx });
  });
  return rows;
}

function uniqueStrings_(values) {
  var seen = {}, result = [];
  (values || []).forEach(function (value) {
    var text = toStr_(value);
    if (!text || seen[text]) return;
    seen[text] = true;
    result.push(text);
  });
  return result;
}

function groupCancellationItems_(items, O) {
  var byItem = {}, ordered = [];
  items.forEach(function (item) {
    var itemOrderNo = toStr_(item.row[O.품목별]);
    if (!itemOrderNo) return;
    if (!byItem[itemOrderNo]) {
      byItem[itemOrderNo] = { itemOrderNo: itemOrderNo, rows: [], quantity: 0, skus: [], names: [], options: [] };
      ordered.push(byItem[itemOrderNo]);
    }
    var group = byItem[itemOrderNo];
    group.rows.push(item);
    group.quantity += toNum_(item.row[O.수량]);
    group.skus.push(toStr_(item.row[O.상품코드]));
    if (O.상품명 >= 0) group.names.push(toStr_(item.row[O.상품명]));
    if (O.옵션 >= 0) group.options.push(toStr_(item.row[O.옵션]));
  });
  return { byItem: byItem, ordered: ordered };
}

function cancellationGroupState_(group, O) {
  var states = uniqueStrings_(group.rows.map(function (item) { return item.row[O.상태]; }));
  if (states.length === 1) return states[0];
  if (!group.rows.some(function (item) { return toStr_(item.row[O.상태]) !== ENUM.주문상태.취소; })) return ENUM.주문상태.취소;
  return '혼합';
}

function cancellationGroupCompleted_(group, O) {
  return group.rows.some(function (item) {
    return toStr_(item.row[O.상태]) !== ENUM.주문상태.취소 &&
      (toStr_(item.row[O.상태]) === ENUM.주문상태.출고완료 || toNum_(item.row[O.출고]) > 0);
  });
}

function cancelOrderCore_(orderNo, reason, source, options) {
  options = options || {};
  return withLock_(function () {
    orderNo = toStr_(orderNo); reason = toStr_(reason) || '기타'; source = toStr_(source) || 'SYSTEM';
    var 주문 = readTable_(ROLE.주문), O = cancellationOrderColumns_(주문);
    var orderItems = orderRowsForCancellation_(주문, O, orderNo);
    if (!orderItems.length) throw new Error('주문번호를 찾을 수 없습니다: ' + orderNo);
    var active = orderItems.filter(function (item) { return toStr_(item.row[O.상태]) !== ENUM.주문상태.취소; });
    if (!active.length) return { 취소: false, 이미취소: true, 메시지: '이미 취소된 주문입니다.' };
    var activeGroups = groupCancellationItems_(active, O).ordered;
    if (activeGroups.some(function (group) { return cancellationGroupCompleted_(group, O); }) && !options.confirmReturn) {
      return { 취소: false, 확인필요: true, 메시지: '이미 피킹지시서가 출력되어 출고 처리된 주문 품목이 있습니다.\n취소하면 해당 출고 수량을 가용재고로 복원합니다.\n계속하시겠습니까?' };
    }
    var result = applyCancellationRows_(주문, O, orderNo, activeGroups, reason, source, new Date());
    writeOpLog_('cancelOrder_', '성공', orderNo + ' / ' + source + ' / ' + reason);
    if (options.refresh !== false) try { D0_대시보드전체갱신(true); } catch (ignore) { }
    return { 취소: true, 주문번호: orderNo, 복원건: result.logCount, 복원수량: result.restoredQuantity };
  });
}

/** 선택한 품목별 주문번호만 취소하는 독립 서비스. */
function cancelOrderItems_(orderNo, itemOrderNos, reason, source, options) {
  return notifyCancellationFailure_(orderNo, source, function () {
    return cancelOrderItemsCore_(orderNo, itemOrderNos, reason, source, options);
  });
}

function cancelOrderItemsCore_(orderNo, itemOrderNos, reason, source, options) {
  options = options || {};
  return withLock_(function () {
    orderNo = toStr_(orderNo); reason = toStr_(reason) || '기타'; source = toStr_(source) || 'SYSTEM';
    var requested = uniqueStrings_(itemOrderNos);
    if (!requested.length) throw new Error('취소할 품목을 하나 이상 선택하세요.');
    var 주문 = readTable_(ROLE.주문), O = cancellationOrderColumns_(주문);
    var orderItems = orderRowsForCancellation_(주문, O, orderNo);
    if (!orderItems.length) throw new Error('주문번호를 찾을 수 없습니다: ' + orderNo);
    var grouped = groupCancellationItems_(orderItems, O), selected = [], already = [], failed = [];
    requested.forEach(function (itemOrderNo) {
      var group = grouped.byItem[itemOrderNo];
      if (!group) {
        failed.push({ itemOrderNo: itemOrderNo, reason: '현재 주문에 속하지 않는 품목별 주문번호입니다.' });
        return;
      }
      var activeRows = group.rows.filter(function (item) { return toStr_(item.row[O.상태]) !== ENUM.주문상태.취소; });
      if (!activeRows.length) { already.push(itemOrderNo); return; }
      group.rows = activeRows;
      selected.push(group);
    });
    if (!selected.length) return { scope: 'ITEMS', requestedItems: requested, successItems: [],
      alreadyCancelledItems: already, failedItems: failed, restoredQuantity: 0 };
    if (selected.some(function (group) { return cancellationGroupCompleted_(group, O); }) && !options.confirmReturn) {
      throw new Error('선택한 품목에 출고 처리된 항목이 있어 재고 복원 확인이 필요합니다.');
    }
    var result = applyCancellationRows_(주문, O, orderNo, selected, reason, source, new Date());
    var success = selected.map(function (group) { return group.itemOrderNo; });
    writeOpLog_('cancelOrderItems_', failed.length ? '부분실패' : '성공',
      orderNo + ' / 품목=' + success.join(',') + ' / ' + source + ' / ' + reason);
    if (options.refresh !== false) try { D0_대시보드전체갱신(true); } catch (ignore) { }
    return { scope: 'ITEMS', requestedItems: requested, successItems: success,
      alreadyCancelledItems: already, failedItems: failed, restoredQuantity: result.restoredQuantity };
  });
}

function applyCancellationRows_(주문, O, orderNo, groups, reason, source, now) {
  var restoration = restorationBySkuForGroups_(groups, O), 마스터 = readTable_(ROLE.마스터);
  var M = { 코드: col_(마스터, COL.상품품목코드, true), 가용: col_(마스터, COL.가용재고, true) };
  var masterIndex = {}, available = [];
  마스터.rows.forEach(function (row, idx) { masterIndex[toStr_(row[M.코드])] = idx; available[idx] = toNum_(row[M.가용]); });
  var actor = 사용자_(), logs = [], itemOrderNos = groups.map(function (group) { return group.itemOrderNo; });
  var instructionByItem = {};
  groups.forEach(function (group) {
    group.rows.forEach(function (item) {
      var instruction = O.피킹지시 >= 0 ? toStr_(item.row[O.피킹지시]) : '';
      if (instruction) instructionByItem[group.itemOrderNo] = instruction;
    });
  });
  Object.keys(restoration).forEach(function (code) {
    var restore = restoration[code], mi = masterIndex[code];
    if (!restore || mi === undefined) return;
    available[mi] += restore;
    logs.push({ 구분: ENUM.로그구분.복원, 피킹지시번호: '', 주문번호: orderNo,
      품목별주문번호: itemOrderNos.join(','), 상품코드: code, 변동량: restore,
      변동후재고: available[mi], 담당자: actor, 사유: source + ' 취소 복원 · ' + reason });
  });
  groups.forEach(function (group) {
    group.rows.forEach(function (item) {
      item.row[O.상태] = ENUM.주문상태.취소; item.row[O.출고] = 0;
      item.row[O.취소사유] = reason; item.row[O.취소일시] = now;
      if (O.취소경로 >= 0) item.row[O.취소경로] = source;
      if (O.대기사유 >= 0) item.row[O.대기사유] = '';
    });
  });
  writeColumn_(주문.sheet, O.상태, 주문.rows); writeColumn_(주문.sheet, O.출고, 주문.rows);
  writeColumn_(주문.sheet, O.취소사유, 주문.rows); writeColumn_(주문.sheet, O.취소일시, 주문.rows);
  if (O.취소경로 >= 0) writeColumn_(주문.sheet, O.취소경로, 주문.rows);
  if (O.대기사유 >= 0) writeColumn_(주문.sheet, O.대기사유, 주문.rows);
  if (마스터.rows.length) 마스터.sheet.getRange(2, M.가용 + 1, available.length, 1)
    .setValues(available.map(function (value) { return [value]; }));
  closePickingItemsForCancellation_(itemOrderNos, now, instructionByItem);
  writeStockLog_(logs);
  return { logCount: logs.length, restoredQuantity: logs.reduce(function (sum, log) { return sum + toNum_(log.변동량); }, 0) };
}

function restorationBySkuForGroups_(groups, O) {
  var result = {};
  groups.forEach(function (group) {
    var completed = cancellationGroupCompleted_(group, O);
    if (completed) {
      var shipped = shippedQuantityBySkuForItems_([group], O);
      if (Object.keys(shipped).length) {
        Object.keys(shipped).forEach(function (code) { result[code] = (result[code] || 0) + shipped[code]; });
        return;
      }
    }
    group.rows.forEach(function (item) {
      var state = toStr_(item.row[O.상태]);
      var reserved = state === '처리완료' ||
        (state === ENUM.주문상태.예약 && O.확정일시 >= 0 && !isBlank_(item.row[O.확정일시]));
      if (!completed && !reserved) return;
      var code = toStr_(item.row[O.상품코드]);
      if (code) result[code] = (result[code] || 0) + toNum_(item.row[O.수량]);
    });
  });
  return result;
}

function shippedQuantityBySkuForItems_(groups, O) {
  var selected = {}, uniqueSku = {};
  groups.forEach(function (group) {
    selected[group.itemOrderNo] = true;
    var skus = uniqueStrings_(group.rows.map(function (item) { return item.row[O.상품코드]; }));
    uniqueSku[group.itemOrderNo] = skus.length === 1 ? skus[0] : '';
  });
  var lines = readTable_(ROLE.라인), cItem = col_(lines, COL.품목별주문번호, true);
  var cActual = col_(lines, COL.실제수량, true), cSku = optionalCancellationCol_(lines, COL.상품코드), result = {};
  lines.rows.forEach(function (row) {
    var itemNo = toStr_(row[cItem]);
    if (!selected[itemNo]) return;
    var code = cSku >= 0 ? toStr_(row[cSku]) : uniqueSku[itemNo], quantity = toNum_(row[cActual]);
    if (code && quantity > 0) result[code] = (result[code] || 0) + quantity;
  });
  return result;
}

/** 레거시 내부 호출 호환. */
function shippedQuantityBySku_(orderNo, orderItems, O) {
  return shippedQuantityBySkuForItems_(groupCancellationItems_(orderItems, O).ordered, O);
}

function closePickingItemsForCancellation_(itemOrderNos, now, instructionByItem) {
  var selected = {};
  uniqueStrings_(itemOrderNos).forEach(function (itemNo) { selected[itemNo] = true; });
  instructionByItem = instructionByItem || {};
  var 라인 = readTable_(ROLE.라인), lItem = col_(라인, COL.품목별주문번호, true);
  var lState = col_(라인, COL.라인상태, true), lTime = col_(라인, COL.처리일시, true);
  var lInstruction = optionalCancellationCol_(라인, COL.피킹지시번호), affectedInstructions = {};
  라인.rows.forEach(function (row) {
    var itemNo = toStr_(row[lItem]);
    if (!selected[itemNo]) return;
    var instruction = (lInstruction >= 0 ? toStr_(row[lInstruction]) : '') || instructionByItem[itemNo];
    if (instruction) affectedInstructions[instruction] = true;
    row[lState] = ENUM.라인상태.취소; row[lTime] = now;
  });
  Object.keys(selected).forEach(function (itemNo) {
    if (instructionByItem[itemNo]) affectedInstructions[instructionByItem[itemNo]] = true;
  });
  if (라인.rows.length) { writeColumn_(라인.sheet, lState, 라인.rows); writeColumn_(라인.sheet, lTime, 라인.rows); }
  var hasRemaining = {};
  라인.rows.forEach(function (row) {
    var itemNo = toStr_(row[lItem]);
    var instruction = (lInstruction >= 0 ? toStr_(row[lInstruction]) : '') || instructionByItem[itemNo];
    if (affectedInstructions[instruction] && toStr_(row[lState]) !== ENUM.라인상태.취소) hasRemaining[instruction] = true;
  });
  var 헤더 = readTable_(ROLE.헤더), hInstruction = col_(헤더, COL.피킹지시번호, true);
  var hState = col_(헤더, COL.상태, true);
  헤더.rows.forEach(function (row) {
    var instruction = toStr_(row[hInstruction]);
    if (affectedInstructions[instruction] && !hasRemaining[instruction]) row[hState] = ENUM.헤더상태.취소;
  });
  if (헤더.rows.length) writeColumn_(헤더.sheet, hState, 헤더.rows);
}

function closePickingForCancellation_(orderNo, now) {
  var 주문 = readTable_(ROLE.주문), O = cancellationOrderColumns_(주문);
  var items = orderRowsForCancellation_(주문, O, toStr_(orderNo)), instructions = {};
  items.forEach(function (item) {
    if (O.피킹지시 >= 0) instructions[toStr_(item.row[O.품목별])] = toStr_(item.row[O.피킹지시]);
  });
  closePickingItemsForCancellation_(items.map(function (item) { return item.row[O.품목별]; }), now, instructions);
}

function 선택_주문취소() {
  var ui = SpreadsheetApp.getUi(), range = SpreadsheetApp.getActiveRange();
  if (!range || range.getSheet().getName() !== ROLE.주문 || range.getRow() < 2) {
    alert_('주문(완료)에서 취소할 주문의 셀을 선택하세요.'); return;
  }
  var table = readTable_(ROLE.주문), cNo = col_(table, COL.주문번호, true);
  var cItem = col_(table, COL.품목별주문번호, true), row = range.getRow();
  var orderNo = toStr_(range.getSheet().getRange(row, cNo + 1).getValue());
  var itemOrderNo = toStr_(range.getSheet().getRange(row, cItem + 1).getValue());
  if (!orderNo) { alert_('선택한 행에 주문번호가 없습니다.'); return; }
  var template = HtmlService.createTemplateFromFile('CancellationScope');
  template.orderNoJson = JSON.stringify(orderNo).replace(/</g, '\\u003c');
  template.itemOrderNoJson = JSON.stringify(itemOrderNo).replace(/</g, '\\u003c');
  ui.showModalDialog(template.evaluate().setWidth(620).setHeight(720), '주문 취소');
}

function cancellationContextFromTable_(table, orderNo, selectedItemOrderNo) {
  orderNo = toStr_(orderNo); selectedItemOrderNo = toStr_(selectedItemOrderNo);
  var O = cancellationOrderColumns_(table), orderRows = orderRowsForCancellation_(table, O, orderNo);
  if (!orderRows.length) throw new Error('주문번호를 찾을 수 없습니다: ' + orderNo);
  var groups = groupCancellationItems_(orderRows, O).ordered;
  var items = groups.map(function (group) {
    var state = cancellationGroupState_(group, O);
    return { itemOrderNo: group.itemOrderNo, sku: uniqueStrings_(group.skus).join(', '),
      productName: uniqueStrings_(group.names).join(' / ') || uniqueStrings_(group.skus).join(', ') || '상품',
      option: uniqueStrings_(group.options).join(' / '), quantity: group.quantity, state: state,
      shippedOrCompleted: cancellationGroupCompleted_(group, O),
      cancellable: group.rows.some(function (item) { return toStr_(item.row[O.상태]) !== ENUM.주문상태.취소; }),
      selected: selectedItemOrderNo === group.itemOrderNo };
  });
  var active = items.filter(function (item) { return item.cancellable; });
  return { orderNo: orderNo, state: uniqueStrings_(items.map(function (item) { return item.state; })).join(', '),
    itemCount: items.length, totalQuantity: items.reduce(function (sum, item) { return sum + item.quantity; }, 0),
    cancellableItemCount: active.length,
    cancellableQuantity: active.reduce(function (sum, item) { return sum + item.quantity; }, 0),
    items: items, reasons: CANCELLATION_REASONS.slice() };
}

function getCancellationContext_(orderNo, selectedItemOrderNo) {
  return cancellationContextFromTable_(readTable_(ROLE.주문), orderNo, selectedItemOrderNo);
}

/** HtmlService에서 호출하는 공개 RPC. */
function getCancellationContext(orderNo, selectedItemOrderNo) { return getCancellationContext_(orderNo, selectedItemOrderNo); }

function executeCancellation_(payload) {
  payload = payload || {};
  var selectedOrderNo = toStr_(payload.selectedOrderNo), scope = toStr_(payload.scope).toUpperCase();
  var reason = toStr_(payload.reason);
  if (!selectedOrderNo) throw new Error('선택 주문번호가 없습니다.');
  if (scope !== 'ITEMS' && scope !== 'ORDER') throw new Error('올바른 취소 범위를 선택하세요.');
  if (CANCELLATION_REASONS.indexOf(reason) < 0) throw new Error('올바른 취소 사유를 선택하세요.');
  if (scope === 'ITEMS') return cancelOrderItems_(selectedOrderNo, payload.selectedItemOrderNos, reason, 'MENU_ITEM_CANCEL', {
    confirmReturn: payload.confirmCompleted === true
  });
  var result = cancelOrder_(selectedOrderNo, reason, 'MENU_ORDER_CANCEL', { confirmReturn: payload.confirmCompleted === true });
  if (result.확인필요) throw new Error(result.메시지);
  return { scope: 'ORDER', requestedItems: [selectedOrderNo], successItems: result.취소 ? [selectedOrderNo] : [],
    alreadyCancelledItems: result.이미취소 ? [selectedOrderNo] : [], failedItems: [], restoredQuantity: toNum_(result.복원수량) };
}

/** 이전 HtmlService 호출명 호환. */
function executeCancellationScope_(payload) { return executeCancellation_(payload); }
function executeCancellationScope(payload) { return executeCancellation_(payload); }

/** HtmlService에서 호출하는 공개 RPC. */
function executeCancellation(payload) { return executeCancellation_(payload); }
