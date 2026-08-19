/**
 * S8. 예약상품 피킹 관리.
 *
 * 예약 주문은 고객이 아니라 주문번호가 처리 단위다. 화면은 예약상품 SKU만 선택하며,
 * 서버가 주문일시 FIFO와 주문 전체 재고를 계산한다. 미리보기는 읽기 전용이고 실제
 * 생성 시에는 lock 안에서 주문/재고를 다시 읽어 같은 계산을 수행한다.
 */

function 예약상품_피킹관리() {
  var html = HtmlService.createHtmlOutputFromFile('ReservationPicking').setWidth(1180).setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, '예약상품 피킹 관리');
}

/** 기존 설치/북마크 호환용. 주문 행 선택 방식은 더 이상 사용하지 않는다. */
function 예약_주문피킹서생성() { return 예약상품_피킹관리(); }

function getReservationProductSummary_() {
  return reservationProductSummaries_(readReservationSnapshot_());
}

function reservationProductSummaries_(snapshot) {
  var summaries = [];
  Object.keys(snapshot.reservationProducts).sort().forEach(function (sku) {
    var demandOrders = snapshot.orders.filter(function (order) { return order.required[sku] > 0; });
    if (!demandOrders.length) return;
    var calculated = calculateReservationBatch_(sku, snapshot.inventory, demandOrders);
    var product = snapshot.inventory[sku] || { name: '(미등록)', option: '', available: 0 };
    summaries.push({
      sku: sku, name: product.name, option: product.option,
      reservationOrderCount: demandOrders.length,
      reservationQuantity: demandOrders.reduce(function (sum, order) { return sum + order.required[sku]; }, 0),
      available: product.available, releasableOrderCount: calculated.selected.length,
      releasableQuantity: calculated.targetAllocated, remaining: calculated.remaining[sku] || 0,
      status: reservationProductStatus_(calculated, product.available)
    });
  });
  summaries.sort(function (a, b) {
    if (a.releasableOrderCount !== b.releasableOrderCount) return b.releasableOrderCount - a.releasableOrderCount;
    return a.sku < b.sku ? -1 : a.sku > b.sku ? 1 : 0;
  });
  return summaries;
}

function getReservationProductPreview_(sku) {
  sku = toStr_(sku);
  if (!sku) throw new Error('상품코드를 선택하세요.');
  var snapshot = readReservationSnapshot_();
  if (!snapshot.reservationProducts[sku]) throw new Error('예약상품이 아닙니다: ' + sku);
  return buildReservationPreview_(sku, snapshot);
}

/**
 * Apps Script 서비스에 의존하지 않는 FIFO 계산기.
 * inventory: SKU -> {available, name, option, disabled}
 * orders: [{orderNo, orderDate, itemOrderNo, required:{SKU:qty}}]
 */
function calculateReservationBatch_(targetSku, inventory, orders) {
  targetSku = String(targetSku || '').trim();
  var remaining = {}, selected = [], waiting = [], stopped = false;
  Object.keys(inventory || {}).forEach(function (sku) {
    remaining[sku] = Math.max(0, Number(inventory[sku].available) || 0);
  });
  var candidates = (orders || []).filter(function (order) {
    return order && order.required && Number(order.required[targetSku]) > 0;
  }).slice().sort(compareReservationOrders_);

  candidates.forEach(function (order) {
    var targetRequired = Number(order.required[targetSku]) || 0;
    if (stopped) {
      waiting.push(reservationWait_(order, targetRequired,
        'FIFO 대기: 앞선 주문의 ' + targetSku + ' 재고 부족', [{ sku: targetSku, type: 'FIFO' }]));
      return;
    }
    if ((remaining[targetSku] || 0) < targetRequired) {
      waiting.push(reservationWait_(order, targetRequired,
        targetSku + ' 재고 부족: 필요 ' + targetRequired + ' / 가용 ' + (remaining[targetSku] || 0),
        [{ sku: targetSku, required: targetRequired, available: remaining[targetSku] || 0, type: 'TARGET' }]));
      stopped = true;
      return;
    }

    var otherShortages = [], shortageDetails = [];
    Object.keys(order.required).sort().forEach(function (sku) {
      if (sku === targetSku) return;
      var required = Number(order.required[sku]) || 0, item = inventory[sku];
      if (!item) { otherShortages.push(sku + ' 미등록'); shortageDetails.push({ sku: sku, type: 'MISSING' }); }
      else if (item.disabled) { otherShortages.push(sku + ' 사용중지'); shortageDetails.push({ sku: sku, type: 'DISABLED' }); }
      else if ((remaining[sku] || 0) < required) {
        otherShortages.push(sku + ' 필요 ' + required + ' / 가용 ' + (remaining[sku] || 0));
        shortageDetails.push({ sku: sku, required: required, available: remaining[sku] || 0, type: 'OTHER' });
      }
    });
    if (otherShortages.length) {
      waiting.push(reservationWait_(order, targetRequired, '다른 상품 재고 부족: ' + otherShortages.join(', '), shortageDetails));
      return;
    }

    Object.keys(order.required).forEach(function (sku) {
      remaining[sku] = (remaining[sku] || 0) - (Number(order.required[sku]) || 0);
    });
    selected.push({ orderNo: order.orderNo, orderDate: order.orderDate, itemOrderNo: order.itemOrderNo,
      targetQuantity: targetRequired, totalQuantity: sumRequired_(order.required),
      required: copyNumberMap_(order.required), recipient: order.recipient || '' });
  });
  return {
    targetSku: targetSku, selected: selected, waiting: waiting, remaining: remaining,
    targetAllocated: selected.reduce(function (sum, order) { return sum + order.targetQuantity; }, 0),
    totalAllocated: selected.reduce(function (sum, order) { return sum + order.totalQuantity; }, 0)
  };
}

function compareReservationOrders_(a, b) {
  var at = reservationSortTime_(a.orderDate), bt = reservationSortTime_(b.orderDate);
  if (at !== bt) return at - bt;
  var ano = String(a.orderNo || ''), bno = String(b.orderNo || '');
  if (ano !== bno) return ano < bno ? -1 : 1;
  var ai = String(a.itemOrderNo || ''), bi = String(b.itemOrderNo || '');
  return ai < bi ? -1 : ai > bi ? 1 : 0;
}

function reservationSortTime_(value) {
  if (value instanceof Date) return value.getTime();
  if (value === '' || value === null || value === undefined) return 0;
  var parsed = new Date(value).getTime();
  return isNaN(parsed) ? 0 : parsed;
}

function reservationWait_(order, qty, reason, shortages) {
  return { orderNo: order.orderNo, orderDate: order.orderDate, itemOrderNo: order.itemOrderNo,
    targetQuantity: qty, totalQuantity: sumRequired_(order.required), reason: reason,
    shortages: shortages || [], recipient: order.recipient || '' };
}

function sumRequired_(required) {
  return Object.keys(required || {}).reduce(function (sum, sku) { return sum + (Number(required[sku]) || 0); }, 0);
}

function copyNumberMap_(source) {
  var copy = {};
  Object.keys(source || {}).forEach(function (key) { copy[key] = Number(source[key]) || 0; });
  return copy;
}

/** 예약 상태이고 아직 지시번호가 없는 주문만 계산 후보로 읽는다. */
function readReservationSnapshot_() {
  var orderTable = readTable_(ROLE.주문);
  var O = {
    no: col_(orderTable, COL.주문번호, true), itemNo: col_(orderTable, COL.품목별주문번호, true),
    sku: col_(orderTable, COL.상품품목코드, true), qty: col_(orderTable, COL.수량, true),
    state: col_(orderTable, COL.주문상태, true), instruction: col_(orderTable, COL.피킹지시번호, true),
    date: reservationOptionalColumn_(orderTable, ['주문일시', '주문일자', '주문일', '결제일시']),
    recipientName: reservationOptionalColumn_(orderTable, ['수령인', '수령인명', '수령인 이름']),
    recipientMobile: reservationOptionalColumn_(orderTable, ['수령인 휴대전화', '수령인휴대전화'])
  };
  var grouped = {};
  orderTable.rows.forEach(function (row) {
    if (toStr_(row[O.state]) !== ENUM.주문상태.예약 || !isBlank_(row[O.instruction])) return;
    var no = toStr_(row[O.no]), sku = toStr_(row[O.sku]); if (!no || !sku) return;
    if (!grouped[no]) {
      grouped[no] = { orderNo: no, orderDate: O.date >= 0 ? reservationDateValue_(row[O.date]) : '',
        itemOrderNo: toStr_(row[O.itemNo]), required: {}, items: [], recipient: reservationRecipient_(row, O) };
    }
    var qty = toNum_(row[O.qty]), itemNo = toStr_(row[O.itemNo]);
    grouped[no].required[sku] = (grouped[no].required[sku] || 0) + qty;
    grouped[no].items.push({ sku: sku, quantity: qty, itemOrderNo: itemNo });
    if (!grouped[no].itemOrderNo || (itemNo && itemNo < grouped[no].itemOrderNo)) grouped[no].itemOrderNo = itemNo;
  });

  var master = readTable_(ROLE.마스터);
  var M = {
    sku: col_(master, COL.상품품목코드, true), name: col_(master, COL.상품명, true),
    option: col_(master, COL.옵션명, false), available: col_(master, COL.가용재고, true),
    reservationProduct: col_(master, COL.예약상품, false), state: col_(master, COL.상품상태, false)
  };
  var inventory = {}, reservationProducts = {};
  master.rows.forEach(function (row) {
    var sku = toStr_(row[M.sku]); if (!sku) return;
    inventory[sku] = { available: toNum_(row[M.available]), name: toStr_(row[M.name]),
      option: M.option >= 0 ? toStr_(row[M.option]) : '',
      disabled: M.state >= 0 && toStr_(row[M.state]) === '사용중지' };
    if (M.reservationProduct >= 0 && toStr_(row[M.reservationProduct]).toUpperCase() === 'Y') reservationProducts[sku] = true;
  });
  // 재고 부족으로 예약된 일반상품 주문도 상품 단위 release 경로를 잃지 않게 한다.
  Object.keys(grouped).forEach(function (no) {
    Object.keys(grouped[no].required).forEach(function (sku) { reservationProducts[sku] = true; });
  });
  return { orders: Object.keys(grouped).map(function (no) { return grouped[no]; }),
    inventory: inventory, reservationProducts: reservationProducts };
}

function reservationOptionalColumn_(table, names) {
  for (var i = 0; i < names.length; i++) {
    var index = col_(table, names[i], false); if (index >= 0) return index;
  }
  return -1;
}

function reservationDateValue_(value) {
  return value instanceof Date ? value.getTime() : toStr_(value);
}

/** 고객 정보는 표시 전용이며 fulfillment 판단에는 사용하지 않는다. */
function reservationRecipient_(row, O) {
  var name = O.recipientName >= 0 ? toStr_(row[O.recipientName]) : '';
  var mobile = O.recipientMobile >= 0 ? toStr_(row[O.recipientMobile]).replace(/\D/g, '') : '';
  return name && mobile ? name + ' · ' + mobile : name || mobile;
}

function buildReservationPreview_(sku, snapshot) {
  var orders = snapshot.orders.filter(function (order) { return order.required[sku] > 0; });
  var calculated = calculateReservationBatch_(sku, snapshot.inventory, orders);
  var product = snapshot.inventory[sku] || { name: '(미등록)', option: '', available: 0 };
  return {
    sku: sku, name: product.name, option: product.option, available: product.available,
    reservationOrderCount: orders.length,
    reservationQuantity: orders.reduce(function (sum, order) { return sum + order.required[sku]; }, 0),
    releaseOrderCount: calculated.selected.length, releaseQuantity: calculated.targetAllocated,
    totalPickingQuantity: calculated.totalAllocated, remaining: calculated.remaining[sku] || 0,
    selected: calculated.selected, waiting: calculated.waiting,
    status: reservationProductStatus_(calculated, product.available)
  };
}

function reservationProductStatus_(calculated, available) {
  if (calculated.selected.length) return '출고 가능';
  if (available <= 0) return '재고 없음';
  return calculated.waiting.length ? '주문 전체 재고 확인 필요' : '대기 없음';
}

/** 예약 FIFO 실패 알림을 한곳에서 처리하는 공개/내부 진입점. */
function createReservationPickingBatch(sku) { return createReservationPickingBatch_(sku); }

function createReservationPickingBatch_(sku) {
  try { return createReservationPickingBatchCore_(sku); }
  catch (e) {
    var inventoryState = '확인 불가';
    try {
      var snapshot = readReservationSnapshot_(), product = snapshot.inventory[toStr_(sku)];
      if (product) inventoryState = '가용재고 ' + product.available;
    } catch (ignore) { }
    sendSystemNotification_('ERROR', '예약상품 피킹 실패', {
      선택SKU: toStr_(sku), 피킹지시번호: e.pickingInstructionNo || '생성 전 실패',
      재고상태: inventoryState, 오류: e.message,
      조치: '재고 상태를 확인하고 예약상품 피킹 관리에서 다시 시도하세요.'
    });
    throw e;
  }
}

/** 브라우저의 주문번호를 신뢰하지 않고 lock 안에서 최신 FIFO를 다시 계산한다. */
function createReservationPickingBatchCore_(sku) {
  sku = toStr_(sku); if (!sku) throw new Error('상품코드를 선택하세요.');
  return withLock_(function () {
    var snapshot = readReservationSnapshot_();
    if (!snapshot.reservationProducts[sku]) throw new Error('예약상품이 아닙니다: ' + sku);
    var preview = buildReservationPreview_(sku, snapshot);
    var orderNos = preview.selected.map(function (order) { return order.orderNo; });
    if (!orderNos.length) return { created: false, message: '현재 FIFO 기준으로 출고 가능한 주문이 없습니다.', preview: preview };

    var confirmation = S3_1_주문확정(orderNos, { manualRelease: true, silent: true });
    if (confirmation.준비주문.length !== orderNos.length) {
      rollbackReservationAllocation_(confirmation.준비주문);
      throw new Error('재고가 변경되어 예약 배치를 만들지 못했습니다. 화면을 새로고침하세요.');
    }
    var picking;
    try {
      picking = S4_1_피킹지시생성(orderNos, { silent: true, reservationBatch: true });
      if (!picking.지시번호) throw new Error('피킹 헤더/라인을 만들지 못했습니다.');
    } catch (pickingError) {
      rollbackReservationAllocation_(orderNos); throw pickingError;
    }

    var pdf;
    try {
      pdf = S9_피킹PDF생성(picking.지시번호, inputFolder_('Output폴더ID'));
    } catch (pdfError) {
      markPickingOutputState_(picking.지시번호, ENUM.헤더상태.출력오류);
      writeOpLog_('createReservationPickingBatch_', '실패', picking.지시번호 + ' / PDF / ' + pdfError.message);
      var recoveryError = new Error('PDF 생성에 실패했습니다. 주문은 출고완료로 바뀌지 않았습니다. ' +
        '피킹지시서 조회/재출력에서 ' + picking.지시번호 + '을(를) 복구하세요.\n' + pdfError.message);
      recoveryError.pickingInstructionNo = picking.지시번호;
      throw recoveryError;
    }
    try { D0_대시보드전체갱신(true); } catch (dashboardError) {
      writeOpLog_('createReservationPickingBatch_', '경고', picking.지시번호 + ' / 대시보드 / ' + dashboardError.message);
    }
    var latest = buildReservationPreview_(sku, readReservationSnapshot_());
    writeOpLog_('createReservationPickingBatch_', '성공', picking.지시번호 + ' / ' + sku + ' / 주문 ' + orderNos.length + '건');
    return { created: true, sku: sku, name: preview.name, instructionNo: picking.지시번호,
      orderCount: orderNos.length, releaseQuantity: preview.releaseQuantity,
      totalPickingQuantity: preview.totalPickingQuantity, remaining: latest.available,
      waitingOrderCount: latest.reservationOrderCount, pdf: pdf };
  });
}

/** S3 성공 뒤 S4 생성 전 실패한 경우에만 내부 예약 이동을 되돌린다. */
function rollbackReservationAllocation_(orderNos) {
  if (!orderNos || !orderNos.length) return;
  var target = {}; orderNos.forEach(function (no) { target[toStr_(no)] = true; });
  var orders = readTable_(ROLE.주문);
  var O = { no: col_(orders, COL.주문번호, true), sku: col_(orders, COL.상품품목코드, true),
    qty: col_(orders, COL.수량, true), confirmed: col_(orders, COL.확정일시, false), reason: col_(orders, COL.대기사유, false) };
  var required = {};
  orders.rows.forEach(function (row) {
    var no = toStr_(row[O.no]); if (!target[no] || O.confirmed < 0 || isBlank_(row[O.confirmed])) return;
    var sku = toStr_(row[O.sku]); required[sku] = (required[sku] || 0) + toNum_(row[O.qty]);
    row[O.confirmed] = ''; if (O.reason >= 0) row[O.reason] = '';
  });
  var master = readTable_(ROLE.마스터);
  var M = { sku: col_(master, COL.상품품목코드, true), available: col_(master, COL.가용재고, true), reserved: col_(master, COL.예약재고, true) };
  var index = {}; master.rows.forEach(function (row, i) { index[toStr_(row[M.sku])] = i; });
  Object.keys(required).forEach(function (sku) {
    var i = index[sku]; if (i === undefined) return;
    master.rows[i][M.available] = toNum_(master.rows[i][M.available]) + required[sku];
    master.rows[i][M.reserved] = Math.max(0, toNum_(master.rows[i][M.reserved]) - required[sku]);
  });
  if (orders.rows.length && O.confirmed >= 0) writeColumn_(orders.sheet, O.confirmed, orders.rows);
  if (orders.rows.length && O.reason >= 0) writeColumn_(orders.sheet, O.reason, orders.rows);
  if (master.rows.length) { writeColumn_(master.sheet, M.available, master.rows); writeColumn_(master.sheet, M.reserved, master.rows); }
}

/** 대시보드 호환 DTO. 별도의 예약대기 시트는 만들지 않는다. */
function collectPreorderData_() {
  var snapshot = readReservationSnapshot_(), summaries = reservationProductSummaries_(snapshot);
  var selectedOrders = {}, shortageSku = {};
  summaries.forEach(function (summary) {
    var calculation = calculateReservationBatch_(summary.sku, snapshot.inventory,
      snapshot.orders.filter(function (order) { return order.required[summary.sku] > 0; }));
    calculation.selected.forEach(function (order) { selectedOrders[order.orderNo] = true; });
    calculation.waiting.forEach(function (order) {
      (order.shortages || []).forEach(function (shortage) {
        shortageSku[shortage.sku] = (shortageSku[shortage.sku] || 0) + 1;
      });
    });
  });
  var totalUnits = snapshot.orders.reduce(function (sum, order) { return sum + sumRequired_(order.required); }, 0);
  return {
    전체: snapshot.orders.length, 주문수: snapshot.orders.length, 총수량: totalUnits,
    출고가능: Object.keys(selectedOrders).length, 재고부족: snapshot.orders.length - Object.keys(selectedOrders).length,
    출고가능SKU: summaries.filter(function (summary) { return summary.releasableOrderCount > 0; }).length,
    상품: summaries, 주문: snapshot.orders, 주문목록: snapshot.orders.map(function (order) { return order.orderNo; }),
    예약부족SKU: Object.keys(shortageSku).map(function (sku) { return { 코드: sku, 주문수: shortageSku[sku] }; })
      .sort(function (a, b) { return b.주문수 - a.주문수; })
  };
}
