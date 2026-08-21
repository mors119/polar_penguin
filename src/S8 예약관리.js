/**
 * S8. 예약상품 입고 관리.
 *
 * F 상품의 가용재고는 물리 재고에서 확정된 미완료 수요를 뺀 순 재고 위치다.
 * 주문 확정(S3)이 수요를 한 번만 차감하므로 예약 피킹에서는 재고를 다시 차감하지 않는다.
 * FIFO로 막혀 창고에 남은 물리 재고는 다음처럼 안전하게 다시 파생한다.
 *
 *   미배정 물리 재고 = 순 가용재고 + 아직 피킹지시가 없는 확정 예약 수요
 *
 * 따라서 수동 입고 delta와 이전 미사용 재고를 합쳐 FIFO를 평가할 수 있으며,
 * 별도의 예약재고 열이나 피킹 시 재고 차감이 필요하지 않다.
 */

var RESERVATION_INBOUND_SOURCE = 'MANUAL_RESERVATION_INBOUND';
var RESERVATION_INBOUND_REASON = '예약상품 수동 입고';
var RESERVATION_INBOUND_IDEMPOTENCY_PREFIX = 'reservationInbound:';
var RESERVATION_PHYSICAL_CARRY_PREFIX = 'reservationPhysicalCarry:';

function 예약상품_입고관리() {
  var html = HtmlService.createHtmlOutputFromFile('ReservationInbound').setWidth(1180).setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, '예약상품 입고 관리');
}

/** 이전 배포 메뉴/매크로 호환용. */
function 예약_주문피킹관리() { return 예약상품_입고관리(); }

function getReservationProductSummary_() {
  return reservationSkuSummaries_(readReservationSnapshot_(), true);
}

function reservationSkuSummaries_(snapshot, fOnly) {
  var summaries = [], skus = {};
  if (fOnly) {
    Object.keys(snapshot.inventory).forEach(function (sku) {
      if (snapshot.inventory[sku].managed === false) skus[sku] = true;
    });
  } else {
    Object.keys(snapshot.candidateSkus).forEach(function (sku) { skus[sku] = true; });
  }
  Object.keys(skus).sort().forEach(function (sku) {
    var demandOrders = reservationOrdersForSku_(snapshot.orders, sku);
    var product = snapshot.inventory[sku] || { name: '(미등록)', option: '', available: 0, managed: true };
    var physical = product.managed === false
      ? deriveReservationPhysicalStock_(sku, snapshot)
      : Math.max(0, product.available);
    var calculated = calculateReservationBatch_(sku, physical, snapshot.inventory, demandOrders);
    summaries.push({
      sku: sku, name: product.name, option: product.option,
      reservationOrderCount: demandOrders.length,
      reservationQuantity: reservationQuantityForSku_(demandOrders, sku),
      available: product.available, securedAvailable: physical,
      releasableOrderCount: calculated.selected.length,
      releasableQuantity: calculated.targetAllocated,
      unusedSecuredQuantity: calculated.unusedTargetQuantity,
      status: reservationSkuStatus_(calculated, demandOrders.length)
    });
  });
  summaries.sort(function (a, b) {
    var aw = a.reservationOrderCount > 0 ? 1 : 0, bw = b.reservationOrderCount > 0 ? 1 : 0;
    if (aw !== bw) return bw - aw;
    if (a.releasableOrderCount !== b.releasableOrderCount) return b.releasableOrderCount - a.releasableOrderCount;
    return a.sku < b.sku ? -1 : a.sku > b.sku ? 1 : 0;
  });
  return summaries;
}

function getReservationInboundPreview_(sku, inboundQty) {
  sku = toStr_(sku);
  var qty = validateManualReservationInboundQuantity_(inboundQty);
  if (!sku) throw new Error('상품코드를 선택하세요.');
  var snapshot = readReservationSnapshot_();
  assertManualReservationInboundProduct_(sku, snapshot);
  return buildReservationInboundPreview_(sku, qty, snapshot);
}

/** 이전 내부 호출자의 읽기 전용 미리보기 호환용. */
function getReservationProductPreview_(sku) { return getReservationInboundPreview_(sku, 0); }

function buildReservationInboundPreview_(sku, inboundQty, snapshot) {
  var orders = reservationOrdersForSku_(snapshot.orders, sku);
  var product = snapshot.inventory[sku];
  var existingPhysical = deriveReservationPhysicalStock_(sku, snapshot);
  var effectivePhysical = existingPhysical + inboundQty;
  var calculated = calculateReservationBatch_(sku, effectivePhysical, snapshot.inventory, orders);
  return {
    sku: sku, name: product.name, option: product.option,
    available: product.available, inboundQuantity: inboundQty,
    inventoryAfterInbound: product.available + inboundQty,
    existingSecuredQuantity: existingPhysical,
    effectiveReleasableQuantity: effectivePhysical,
    reservationOrderCount: orders.length,
    reservationQuantity: reservationQuantityForSku_(orders, sku),
    releaseOrderCount: calculated.selected.length,
    releaseQuantity: calculated.targetAllocated,
    totalPickingQuantity: calculated.totalAllocated,
    unusedInboundQuantity: calculated.unusedTargetQuantity,
    selected: calculated.selected, waiting: calculated.waiting,
    status: reservationSkuStatus_(calculated, orders.length)
  };
}

/**
 * Apps Script 서비스에 의존하지 않는 FIFO 계산기.
 * targetAvailable은 순 가용재고가 아니라 target SKU의 실제 미배정 물리 수량이다.
 * 이미 확정된 주문은 S3에서 다른 SKU까지 재고가 예약되었으므로 교차 SKU를 다시 차감하지 않는다.
 */
function calculateReservationBatch_(targetSku, targetAvailable, inventory, orders) {
  // 이전 테스트/내부 호출의 (targetSku, inventory, orders) 형태도 안전하게 읽는다.
  if (arguments.length === 3 && targetAvailable && typeof targetAvailable === 'object') {
    orders = inventory;
    inventory = targetAvailable;
    targetAvailable = inventory[targetSku] ? inventory[targetSku].available : 0;
  }
  targetSku = String(targetSku || '').trim();
  var targetRemaining = Math.max(0, Number(targetAvailable) || 0);
  var remaining = {}, selected = [], waiting = [], stopped = false;
  Object.keys(inventory || {}).forEach(function (sku) {
    remaining[sku] = Number(inventory[sku].available) || 0;
  });
  var candidates = (orders || []).filter(function (order) {
    return order && order.required && Number(order.required[targetSku]) > 0;
  }).slice().sort(compareReservationOrders_);

  candidates.forEach(function (order) {
    var targetRequired = Number(order.required[targetSku]) || 0;
    if (stopped) {
      waiting.push(reservationWait_(order, targetRequired,
        'FIFO 대기: 앞선 주문의 ' + targetSku + ' 물리 재고 부족', [{ sku: targetSku, type: 'FIFO' }]));
      return;
    }
    if (!inventory[targetSku] || targetRemaining < targetRequired) {
      waiting.push(reservationWait_(order, targetRequired,
        targetSku + ' 물리 재고 부족: 필요 ' + targetRequired + ' / 사용 가능 ' + targetRemaining,
        [{ sku: targetSku, required: targetRequired, available: targetRemaining, type: 'TARGET' }]));
      stopped = true;
      return;
    }

    var otherShortages = [], shortageDetails = [];
    if (!order.committed) {
      Object.keys(order.required).sort().forEach(function (sku) {
        if (sku === targetSku) return;
        var required = Number(order.required[sku]) || 0, item = inventory[sku];
        if (!item) {
          otherShortages.push(sku + ' 미등록'); shortageDetails.push({ sku: sku, type: 'MISSING' });
        } else if (item.managed !== false && (remaining[sku] || 0) < required) {
          otherShortages.push(sku + ' 필요 ' + required + ' / 가용 ' + (remaining[sku] || 0));
          shortageDetails.push({ sku: sku, required: required, available: remaining[sku] || 0, type: 'OTHER' });
        }
      });
    }
    if (otherShortages.length) {
      waiting.push(reservationWait_(order, targetRequired, '다른 상품 재고 부족: ' + otherShortages.join(', '), shortageDetails));
      return;
    }

    targetRemaining -= targetRequired;
    if (!order.committed) {
      Object.keys(order.required).forEach(function (sku) {
        if (sku !== targetSku) remaining[sku] = (remaining[sku] || 0) - (Number(order.required[sku]) || 0);
      });
    }
    selected.push({ orderNo: order.orderNo, orderDate: order.orderDate, itemOrderNo: order.itemOrderNo,
      targetQuantity: targetRequired, totalQuantity: sumRequired_(order.required), committed: !!order.committed,
      required: copyNumberMap_(order.required), recipient: order.recipient || '' });
  });
  remaining[targetSku] = targetRemaining;
  return {
    targetSku: targetSku, selected: selected, waiting: waiting, remaining: remaining,
    targetAllocated: selected.reduce(function (sum, order) { return sum + order.targetQuantity; }, 0),
    totalAllocated: selected.reduce(function (sum, order) { return sum + order.totalQuantity; }, 0),
    unusedTargetQuantity: targetRemaining
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

function reservationOrdersForSku_(orders, sku) {
  return (orders || []).filter(function (order) { return Number(order.required[sku]) > 0; });
}

function reservationQuantityForSku_(orders, sku) {
  return (orders || []).reduce(function (sum, order) { return sum + (Number(order.required[sku]) || 0); }, 0);
}

/** 예약 상태이고 아직 지시번호가 없는 실제 주문 행만 FIFO 후보로 읽는다. */
function readReservationSnapshot_() {
  var orderTable = readTable_(ROLE.주문);
  var O = {
    no: col_(orderTable, COL.주문번호, true), itemNo: col_(orderTable, COL.품목별주문번호, true),
    sku: col_(orderTable, COL.상품품목코드, true), qty: col_(orderTable, COL.수량, true),
    state: col_(orderTable, COL.주문상태, true), instruction: col_(orderTable, COL.피킹지시번호, true),
    confirmed: col_(orderTable, COL.확정일시, false),
    date: reservationOptionalColumn_(orderTable, ['주문일시', '주문일자', '주문일', '결제일시']),
    recipientName: col_(orderTable, COL.수령인, true),
    recipientMobile: col_(orderTable, COL.수령인휴대전화, true)
  };
  var grouped = {};
  orderTable.rows.forEach(function (row) {
    if (toStr_(row[O.state]) !== ENUM.주문상태.예약 || !isBlank_(row[O.instruction])) return;
    var no = toStr_(row[O.no]), sku = toStr_(row[O.sku]); if (!no || !sku) return;
    if (!grouped[no]) {
      grouped[no] = { orderNo: no, orderDate: O.date >= 0 ? reservationDateValue_(row[O.date]) : '',
        itemOrderNo: toStr_(row[O.itemNo]), required: {}, items: [], recipient: reservationRecipient_(row, O),
        committed: O.confirmed >= 0 && !isBlank_(row[O.confirmed]) };
    }
    var qty = toNum_(row[O.qty]), itemNo = toStr_(row[O.itemNo]);
    grouped[no].required[sku] = (grouped[no].required[sku] || 0) + qty;
    grouped[no].items.push({ sku: sku, quantity: qty, itemOrderNo: itemNo });
    if (O.confirmed < 0 || isBlank_(row[O.confirmed])) grouped[no].committed = false;
    if (!grouped[no].itemOrderNo || (itemNo && itemNo < grouped[no].itemOrderNo)) grouped[no].itemOrderNo = itemNo;
  });

  var master = readTable_(ROLE.마스터);
  var M = {
    sku: col_(master, COL.상품품목코드, true), name: col_(master, COL.상품명, true),
    option: col_(master, COL.옵션명, false), available: col_(master, COL.가용재고, true),
    managed: col_(master, COL.재고관리, true)
  };
  var inventory = {}, candidateSkus = {};
  master.rows.forEach(function (row, rowIndex) {
    var sku = toStr_(row[M.sku]); if (!sku) return;
    inventory[sku] = { available: toNum_(row[M.available]), name: toStr_(row[M.name]),
      option: M.option >= 0 ? toStr_(row[M.option]) : '', rowIndex: rowIndex,
      managed: toStr_(row[M.managed]).toUpperCase() !== 'F' };
  });
  Object.keys(grouped).forEach(function (no) {
    Object.keys(grouped[no].required).forEach(function (sku) { candidateSkus[sku] = true; });
  });
  return { orders: Object.keys(grouped).map(function (no) { return grouped[no]; }),
    inventory: inventory, candidateSkus: candidateSkus };
}

function deriveReservationPhysicalStock_(sku, snapshot) {
  var product = snapshot.inventory[sku];
  if (!product) return 0;
  var confirmedWaitingDemand = reservationOrdersForSku_(snapshot.orders, sku).reduce(function (sum, order) {
    return sum + (order.committed ? Number(order.required[sku]) || 0 : 0);
  }, 0);
  var derived = Math.max(0, product.available + confirmedWaitingDemand);
  // 순 재고와 현재 FIFO 행이 과거 데이터 때문에 완전히 맞지 않아도, 이전 release에서
  // 명시적으로 남긴 실물 수량은 잃지 않는다. Spreadsheet 열은 추가하지 않는다.
  return Math.max(derived, readReservationPhysicalCarry_(sku));
}

function readReservationPhysicalCarry_(sku) {
  try {
    var value = PropertiesService.getScriptProperties().getProperty(RESERVATION_PHYSICAL_CARRY_PREFIX + sku);
    var number = Number(value);
    return isFinite(number) && number > 0 ? number : 0;
  } catch (e) { return 0; }
}

function writeReservationPhysicalCarry_(sku, quantity) {
  var properties = PropertiesService.getScriptProperties();
  var key = RESERVATION_PHYSICAL_CARRY_PREFIX + sku;
  var value = Math.max(0, Number(quantity) || 0);
  if (value > 0) properties.setProperty(key, String(value));
  else properties.deleteProperty(key);
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

function reservationSkuStatus_(calculated, waitingOrderCount) {
  if (calculated.selected.length) return 'FIFO 출고 가능';
  if (!waitingOrderCount) return '대기 없음';
  if (calculated.unusedTargetQuantity <= 0) return '입고 대기';
  return '주문 전체 재고 확인 필요';
}

function validateManualReservationInboundQuantity_(value) {
  if (value === '' || value === null || value === undefined || typeof value === 'boolean') {
    throw new Error('이번 입고 수량은 0 이상의 정수여야 합니다.');
  }
  var number = typeof value === 'number' ? value : Number(String(value).replace(/,/g, '').trim());
  if (!isFinite(number) || Math.floor(number) !== number || number < 0) {
    throw new Error('이번 입고 수량은 0 이상의 정수여야 합니다.');
  }
  return number;
}

function assertManualReservationInboundProduct_(sku, snapshot) {
  var product = snapshot.inventory[sku];
  if (!product) throw new Error('상품마스터에 없는 SKU입니다: ' + sku);
  if (product.managed !== false) throw new Error('예약상품 수동 입고는 재고관리=F 상품만 처리할 수 있습니다.');
  return product;
}

function validateReservationInboundRequestId_(requestId) {
  requestId = String(requestId || '').trim();
  if (!/^[A-Za-z0-9_-]{12,100}$/.test(requestId)) throw new Error('요청 식별자가 올바르지 않습니다. 화면을 새로고침하세요.');
  return requestId;
}

function applyManualReservationInbound(sku, inboundQty, requestId) {
  return applyManualReservationInbound_(sku, inboundQty, requestId);
}

function applyManualReservationInbound_(sku, inboundQty, requestId) {
  try { return applyManualReservationInboundCore_(sku, inboundQty, requestId); }
  catch (e) {
    sendSystemNotification_('ERROR', '예약상품 수동 입고 실패', {
      SKU: toStr_(sku), source: RESERVATION_INBOUND_SOURCE,
      입고수량: String(inboundQty), 피킹지시번호: e.pickingInstructionNo || '생성 전 실패',
      오류: e.message, 조치: '재고이동로그와 피킹지시서 조회/재출력을 확인하세요.'
    });
    throw e;
  }
}

/** 수동 delta와 FIFO 재검증을 하나의 Script Lock 안에서 직렬화한다. */
function applyManualReservationInboundCore_(sku, inboundQty, requestId) {
  sku = toStr_(sku);
  var qty = validateManualReservationInboundQuantity_(inboundQty);
  requestId = validateReservationInboundRequestId_(requestId);
  return withLock_(function () {
    var properties = PropertiesService.getScriptProperties();
    var propertyKey = RESERVATION_INBOUND_IDEMPOTENCY_PREFIX + requestId;
    var previous = parseReservationInboundRecord_(properties.getProperty(propertyKey));
    if (previous) return duplicateReservationInboundResult_(previous);

    // 잘못된 SKU나 T 상품은 idempotency 기록도 만들지 않고 먼저 거부한다.
    var snapshot = readReservationSnapshot_();
    var product = assertManualReservationInboundProduct_(sku, snapshot);
    var existingPhysical = deriveReservationPhysicalStock_(sku, snapshot);
    pruneReservationInboundRecords_(properties, new Date().getTime());
    properties.setProperty(propertyKey, JSON.stringify({ state: 'PENDING', sku: sku, inboundQuantity: qty,
      inboundApplied: false, timestamp: new Date().toISOString() }));
    var netAfter, mutationStarted = false, inboundApplied = false;
    try {
      netAfter = product.available + qty;
      if (qty > 0) {
        mutationStarted = true;
        applyManualReservationInboundStock_(sku, qty, netAfter, requestId);
        inboundApplied = true;
      }
      properties.setProperty(propertyKey, JSON.stringify({ state: 'STOCK_APPLIED', sku: sku,
        inboundQuantity: qty, inboundApplied: inboundApplied,
        inventoryAfterInbound: netAfter, timestamp: new Date().toISOString() }));

      var effectivePhysical = existingPhysical + qty;
      // 이후 피킹/PDF가 실패해도 입고된 실물과 기존 미사용분은 복구 가능해야 한다.
      writeReservationPhysicalCarry_(sku, effectivePhysical);
      var result = releaseReservationQueueCore_(sku, effectivePhysical, {
        source: RESERVATION_INBOUND_SOURCE, silent: true, outputFolder: inputFolder_('Output폴더ID')
      });
      result.inboundApplied = qty > 0;
      result.inboundQuantity = qty;
      result.inventoryAfterInbound = netAfter;
      result.source = RESERVATION_INBOUND_SOURCE;
      var storedResult = reservationInboundStoredResult_(result);
      properties.setProperty(propertyKey, JSON.stringify({ state: 'COMPLETE', sku: sku,
        inboundQuantity: qty, inboundApplied: inboundApplied,
        inventoryAfterInbound: netAfter, timestamp: new Date().toISOString(), result: storedResult }));
      return result;
    } catch (e) {
      properties.setProperty(propertyKey, JSON.stringify({ state: mutationStarted ? 'MUTATION_REVIEW' : 'REEVALUATION_ERROR', sku: sku,
        inboundQuantity: qty, inboundApplied: inboundApplied, inventoryAfterInbound: netAfter,
        instructionNo: e.pickingInstructionNo || '',
        message: e.message, timestamp: new Date().toISOString() }));
      throw e;
    }
  });
}

function applyManualReservationInboundStock_(sku, qty, netAfter, requestId) {
  var master = readTable_(ROLE.마스터);
  var M = { sku: col_(master, COL.상품품목코드, true), available: col_(master, COL.가용재고, true),
    managed: col_(master, COL.재고관리, true) };
  var rowIndex = -1;
  master.rows.some(function (row, i) { if (toStr_(row[M.sku]) !== sku) return false; rowIndex = i; return true; });
  if (rowIndex < 0) throw new Error('상품마스터에 없는 SKU입니다: ' + sku);
  if (toStr_(master.rows[rowIndex][M.managed]).toUpperCase() !== 'F') {
    throw new Error('예약상품 수동 입고는 재고관리=F 상품만 처리할 수 있습니다.');
  }
  master.rows[rowIndex][M.available] = netAfter;
  master.sheet.getRange(rowIndex + 2, M.available + 1).setValue(netAfter);
  writeStockLog_([{ 구분: ENUM.로그구분.입고, 피킹지시번호: '', 주문번호: '', 품목별주문번호: '',
    상품코드: sku, 변동량: qty, 변동후재고: netAfter, 담당자: 사용자_(),
    사유: RESERVATION_INBOUND_REASON + ' · source=' + RESERVATION_INBOUND_SOURCE + ' · request=' + requestId }]);
}

/**
 * 공유 예약 release 서비스. normal Input이 자동 release를 도입할 때도 S1의 SKU별
 * physical snapshot 반영 후 이 진입점에 새 물리 수량(기존 미사용 포함)을 전달한다.
 */
function releaseReservationQueue_(sku, releasableQty, options) {
  return withLock_(function () { return releaseReservationQueueCore_(sku, releasableQty, options || {}); });
}

function releaseReservationQueueCore_(sku, releasableQty, options) {
  options = options || {};
  sku = toStr_(sku); if (!sku) throw new Error('상품코드를 선택하세요.');
  var snapshot = readReservationSnapshot_();
  assertManualReservationInboundProduct_(sku, snapshot);
  var effective = releasableQty === undefined || releasableQty === null
    ? deriveReservationPhysicalStock_(sku, snapshot)
    : Math.max(0, Number(releasableQty) || 0);
  var preview = buildReservationReleasePreview_(sku, effective, snapshot);
  var orderNos = preview.selected.map(function (order) { return order.orderNo; });
  if (!orderNos.length) {
    writeReservationPhysicalCarry_(sku, preview.unusedSecuredQuantity);
    writeReservationReleaseLog_(sku, options.source, preview, '피킹 없음', '');
    return { created: false, sku: sku, name: preview.name,
      message: '현재 FIFO 기준으로 출고 가능한 주문이 없습니다.',
      orderCount: 0, releaseQuantity: 0, totalPickingQuantity: 0,
      unusedSecuredQuantity: preview.unusedSecuredQuantity,
      waitingOrderCount: preview.reservationOrderCount, waitingQuantity: preview.reservationQuantity,
      preview: preview };
  }

  var newlyCommitted = preview.selected.filter(function (order) { return !order.committed; })
    .map(function (order) { return order.orderNo; });
  var confirmation = S3_1_주문확정(orderNos, { silent: true });
  if (confirmation.준비주문.length !== orderNos.length) {
    rollbackStockCommitment_(newlyCommitted);
    throw new Error('재고가 변경되어 예약 배치를 만들지 못했습니다. 화면을 새로고침하세요.');
  }
  var picking;
  try {
    picking = S4_1_피킹지시생성(orderNos, { silent: true, reservationBatch: true });
    if (!picking.지시번호) throw new Error('피킹 헤더/라인을 만들지 못했습니다.');
  } catch (pickingError) {
    rollbackStockCommitment_(newlyCommitted);
    writeReservationPhysicalCarry_(sku, effective);
    throw pickingError;
  }

  // 피킹지시가 확보한 수량은 carry에서 제외하고, FIFO로 막힌 실물만 다음 이벤트에 넘긴다.
  writeReservationPhysicalCarry_(sku, preview.unusedSecuredQuantity);

  var pdf;
  try {
    pdf = S9_피킹PDF생성(picking.지시번호, options.outputFolder || inputFolder_('Output폴더ID'));
  } catch (pdfError) {
    markPickingOutputState_(picking.지시번호, ENUM.헤더상태.출력오류);
    writeReservationReleaseLog_(sku, options.source, preview, 'PDF 실패', picking.지시번호);
    var recoveryError = new Error('PDF 생성에 실패했습니다. 수동 입고는 다시 적용하지 마세요. ' +
      '피킹지시서 조회/재출력에서 ' + picking.지시번호 + '을(를) 복구하세요.\n' + pdfError.message);
    recoveryError.pickingInstructionNo = picking.지시번호;
    throw recoveryError;
  }
  try { D0_대시보드전체갱신(true); } catch (dashboardError) {
    writeOpLog_('releaseReservationQueue_', '경고', picking.지시번호 + ' / 대시보드 / ' + dashboardError.message);
  }
  var latest = readReservationSnapshot_();
  var latestOrders = reservationOrdersForSku_(latest.orders, sku);
  var unused = preview.unusedSecuredQuantity;
  writeReservationReleaseLog_(sku, options.source, preview, '성공', picking.지시번호);
  return { created: true, sku: sku, name: preview.name, instructionNo: picking.지시번호,
    orderCount: orderNos.length, releaseQuantity: preview.releaseQuantity,
    totalPickingQuantity: preview.totalPickingQuantity, unusedSecuredQuantity: unused,
    waitingOrderCount: latestOrders.length, waitingQuantity: reservationQuantityForSku_(latestOrders, sku), pdf: pdf };
}

function buildReservationReleasePreview_(sku, releasableQty, snapshot) {
  var orders = reservationOrdersForSku_(snapshot.orders, sku);
  var calculated = calculateReservationBatch_(sku, releasableQty, snapshot.inventory, orders);
  var product = snapshot.inventory[sku];
  return { sku: sku, name: product.name, option: product.option, available: product.available,
    reservationOrderCount: orders.length, reservationQuantity: reservationQuantityForSku_(orders, sku),
    releaseOrderCount: calculated.selected.length, releaseQuantity: calculated.targetAllocated,
    totalPickingQuantity: calculated.totalAllocated, unusedSecuredQuantity: calculated.unusedTargetQuantity,
    selected: calculated.selected, waiting: calculated.waiting };
}

function writeReservationReleaseLog_(sku, source, preview, outcome, instructionNo) {
  writeOpLog_('releaseReservationQueue_', outcome,
    'SKU ' + sku + ' / source=' + (source || '') +
    ' / 대기주문 ' + preview.reservationOrderCount + ' / 대기수량 ' + preview.reservationQuantity +
    ' / 출고주문 ' + preview.releaseOrderCount + ' / 출고수량 ' + preview.releaseQuantity +
    ' / 미사용수량 ' + preview.unusedSecuredQuantity +
    (instructionNo ? ' / 지시번호 ' + instructionNo : ''));
}

/** 이전 공개/내부 호출은 입고 0의 안전한 재평가로 유지한다. */
function createReservationPickingBatch(sku) { return createReservationPickingBatch_(sku); }
function createReservationPickingBatch_(sku) {
  try { return releaseReservationQueue_(sku, null, { source: 'RESERVATION_RECOVERY', silent: true }); }
  catch (e) {
    sendSystemNotification_('ERROR', '예약상품 FIFO 피킹 실패', {
      선택SKU: toStr_(sku), 피킹지시번호: e.pickingInstructionNo || '생성 전 실패', 오류: e.message,
      조치: '예약상품 입고 관리 또는 피킹지시서 조회/재출력에서 복구하세요.'
    });
    throw e;
  }
}

function parseReservationInboundRecord_(value) {
  if (!value) return null;
  try { return JSON.parse(value); } catch (e) { return { state: 'UNKNOWN', message: '기존 요청 기록을 확인할 수 없습니다.' }; }
}

function pruneReservationInboundRecords_(properties, nowMillis) {
  try {
    var all = properties.getProperties(), cutoff = nowMillis - 30 * 24 * 60 * 60 * 1000;
    Object.keys(all).forEach(function (key) {
      if (key.indexOf(RESERVATION_INBOUND_IDEMPOTENCY_PREFIX) !== 0) return;
      var record = parseReservationInboundRecord_(all[key]);
      var time = record && record.timestamp ? new Date(record.timestamp).getTime() : NaN;
      if (!isNaN(time) && time < cutoff) properties.deleteProperty(key);
    });
  } catch (e) { }
}

function duplicateReservationInboundResult_(record) {
  if (record.result) {
    record.result.duplicate = true;
    record.result.message = '이미 처리된 요청입니다. 입고 수량은 다시 반영하지 않았습니다.';
    return record.result;
  }
  return { created: false, duplicate: true, inboundApplied: record.inboundApplied === true,
    inboundQuantity: record.inboundQuantity || 0, inventoryAfterInbound: record.inventoryAfterInbound,
    instructionNo: record.instructionNo || '',
    message: record.message || '같은 요청이 이미 처리 중이거나 반영되었습니다. 재고이동로그를 확인하세요.' };
}

function reservationInboundStoredResult_(result) {
  return { created: !!result.created, sku: result.sku, name: result.name,
    instructionNo: result.instructionNo || '', orderCount: result.orderCount || 0,
    releaseQuantity: result.releaseQuantity || 0, totalPickingQuantity: result.totalPickingQuantity || 0,
    unusedSecuredQuantity: result.unusedSecuredQuantity || 0,
    waitingOrderCount: result.waitingOrderCount || 0, waitingQuantity: result.waitingQuantity || 0,
    inboundApplied: !!result.inboundApplied, inboundQuantity: result.inboundQuantity || 0,
    inventoryAfterInbound: result.inventoryAfterInbound, source: result.source || '' };
}

/** S3 성공 뒤 S4 생성 전 실패한 경우, 이번 release에서 새로 확정한 주문만 되돌린다. */
function rollbackStockCommitment_(orderNos) {
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
  var M = { sku: col_(master, COL.상품품목코드, true), available: col_(master, COL.가용재고, true) };
  var index = {}; master.rows.forEach(function (row, i) { index[toStr_(row[M.sku])] = i; });
  Object.keys(required).forEach(function (sku) {
    var i = index[sku]; if (i === undefined) return;
    master.rows[i][M.available] = toNum_(master.rows[i][M.available]) + required[sku];
  });
  if (orders.rows.length && O.confirmed >= 0) writeColumn_(orders.sheet, O.confirmed, orders.rows);
  if (orders.rows.length && O.reason >= 0) writeColumn_(orders.sheet, O.reason, orders.rows);
  if (master.rows.length) writeColumn_(master.sheet, M.available, master.rows);
}

/** 대시보드 호환 DTO. 별도의 예약대기 시트는 만들지 않는다. */
function collectPreorderData_() {
  var snapshot = readReservationSnapshot_(), summaries = reservationSkuSummaries_(snapshot, false);
  var selectedOrders = {}, shortageSku = {};
  summaries.forEach(function (summary) {
    var product = snapshot.inventory[summary.sku];
    var physical = product && product.managed === false
      ? deriveReservationPhysicalStock_(summary.sku, snapshot)
      : Math.max(0, product ? product.available : 0);
    var calculation = calculateReservationBatch_(summary.sku, physical, snapshot.inventory,
      reservationOrdersForSku_(snapshot.orders, summary.sku));
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
