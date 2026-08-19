/**
 * S5. 출력 성공 후 fulfillment 최종 확정.
 *
 * S3가 이미 가용재고를 줄이고 예약재고를 늘렸으므로 여기서는 예약재고만
 * 소진한다. 피킹지시번호와 주문/라인 상태가 idempotency 표식이며, 재출력은
 * 이 함수를 다시 호출해도 재고나 로그를 다시 만들지 않는다.
 */
function finalizePickingAfterOutput_(instructionNo, options) {
  options = options || {};
  return withLock_(function () {
    instructionNo = toStr_(instructionNo);
    if (!instructionNo) throw new Error('피킹지시번호가 없습니다.');

    var 주문 = readTable_(ROLE.주문), 라인 = readTable_(ROLE.라인);
    var 헤더 = readTable_(ROLE.헤더), 마스터 = readTable_(ROLE.마스터);
    var O = {
      주문번호: col_(주문, COL.주문번호, true), 품목별: col_(주문, COL.품목별주문번호, true),
      상품코드: col_(주문, COL.상품품목코드, true), 수량: col_(주문, COL.수량, true),
      상태: col_(주문, COL.주문상태, true), 출고: col_(주문, COL.출고완료, true),
      지시: col_(주문, COL.피킹지시번호, true), 확정일시: col_(주문, COL.확정일시, false),
      대기사유: col_(주문, COL.대기사유, false)
    };
    var L = {
      주문번호: col_(라인, COL.주문번호, false), 품목별: col_(라인, COL.품목별주문번호, true),
      상품코드: col_(라인, COL.상품코드, true), 필요: col_(라인, COL.필요수량, true),
      확인: col_(라인, COL.확인, true), 실제: col_(라인, COL.실제수량, true),
      담당자: col_(라인, COL.담당자, false), 지시: col_(라인, COL.피킹지시번호, true),
      상태: col_(라인, COL.라인상태, true), 처리일시: col_(라인, COL.처리일시, true)
    };
    var H = {
      주문번호: col_(헤더, COL.주문번호, true), 담당자: col_(헤더, COL.피킹담당자, true),
      상태: col_(헤더, COL.상태, true), 지시: col_(헤더, COL.피킹지시번호, true),
      출력일시: col_(헤더, COL.출력일시, false)
    };
    var M = {
      코드: col_(마스터, COL.상품품목코드, true), 예약: col_(마스터, COL.예약재고, true),
      가용: col_(마스터, COL.가용재고, true)
    };

    var selectedOrders = null;
    if (options.orderNos && options.orderNos.length) {
      selectedOrders = {};
      options.orderNos.forEach(function (no) { selectedOrders[toStr_(no)] = true; });
    }
    var headerOrders = {}, headerRows = {}, workers = {};
    헤더.rows.forEach(function (row, idx) {
      if (toStr_(row[H.지시]) !== instructionNo) return;
      var no = toStr_(row[H.주문번호]);
      if (!no || (selectedOrders && !selectedOrders[no])) return;
      headerOrders[no] = true; (headerRows[no] = headerRows[no] || []).push(idx);
      workers[no] = toStr_(row[H.담당자]);
    });
    if (!Object.keys(headerOrders).length) throw new Error('피킹헤더가 없습니다: ' + instructionNo);

    var orderRows = {}, itemOrder = {};
    주문.rows.forEach(function (row, idx) {
      var no = toStr_(row[O.주문번호]);
      if (headerOrders[no]) (orderRows[no] = orderRows[no] || []).push(idx);
      itemOrder[toStr_(row[O.품목별])] = no;
    });
    var lineRows = {};
    라인.rows.forEach(function (row, idx) {
      if (toStr_(row[L.지시]) !== instructionNo) return;
      var no = (L.주문번호 >= 0 ? toStr_(row[L.주문번호]) : '') || itemOrder[toStr_(row[L.품목별])];
      if (no) (lineRows[no] = lineRows[no] || []).push(idx);
    });

    var targetOrders = [], required = {};
    Object.keys(headerOrders).forEach(function (no) {
      var indexes = orderRows[no] || [], lines = lineRows[no] || [];
      if (!indexes.length) throw new Error('피킹 주문을 찾을 수 없습니다: ' + no);
      var state = toStr_(주문.rows[indexes[0]][O.상태]);
      if (state === ENUM.주문상태.취소 || state === ENUM.주문상태.출고완료) return;
      // 처리완료는 이전 버전에서 PDF와 O 확인 사이에 쓰던 값이다. 기존 데이터만 호환한다.
      if (state !== ENUM.주문상태.예약 && state !== '처리완료') {
        throw new Error('출고 확정할 수 없는 주문 상태입니다: ' + no + ' / ' + state);
      }
      if (!lines.length) throw new Error('피킹라인이 없습니다: ' + no);
      indexes.forEach(function (idx) {
        var row = 주문.rows[idx];
        if (toStr_(row[O.지시]) !== instructionNo) throw new Error('주문과 피킹지시 연결이 다릅니다: ' + no);
        if (O.확정일시 < 0 || isBlank_(row[O.확정일시])) throw new Error('예약재고 확보 기록이 없습니다: ' + no);
      });
      var ordered = {}, picked = {};
      indexes.forEach(function (idx) {
        var row = 주문.rows[idx], code = toStr_(row[O.상품코드]);
        ordered[code] = (ordered[code] || 0) + toNum_(row[O.수량]);
      });
      lines.forEach(function (idx) {
        var row = 라인.rows[idx], code = toStr_(row[L.상품코드]);
        if (toStr_(row[L.상태]) === ENUM.라인상태.완료) {
          throw new Error('주문 완료 전 이미 완료된 피킹라인이 있습니다: ' + no);
        }
        picked[code] = (picked[code] || 0) + toNum_(row[L.필요]);
      });
      var skus = {};
      Object.keys(ordered).concat(Object.keys(picked)).forEach(function (code) { skus[code] = true; });
      Object.keys(skus).forEach(function (code) {
        if (ordered[code] !== picked[code]) throw new Error('주문/피킹 수량이 다릅니다: ' + no + ' / ' + code);
        required[code] = (required[code] || 0) + picked[code];
      });
      targetOrders.push(no);
    });

    if (!targetOrders.length) {
      return { 이미완료: true, 지시번호: instructionNo, 완료주문: [], 차감라인: 0, 차감수량: 0 };
    }

    var masterIndex = {}, reserved = [];
    마스터.rows.forEach(function (row, idx) {
      masterIndex[toStr_(row[M.코드])] = idx; reserved[idx] = toNum_(row[M.예약]);
    });
    Object.keys(required).forEach(function (code) {
      var idx = masterIndex[code];
      if (idx === undefined) throw new Error('상품마스터에 없는 피킹 상품입니다: ' + code);
      if (reserved[idx] < required[code]) {
        throw new Error('예약재고가 부족합니다: ' + code + ' / 필요 ' + required[code] + ' / 예약 ' + reserved[idx]);
      }
    });

    var now = new Date(), actor = 사용자_(), logs = [], lineCount = 0, quantity = 0;
    Object.keys(required).forEach(function (code) { reserved[masterIndex[code]] -= required[code]; });
    targetOrders.forEach(function (no) {
      (lineRows[no] || []).forEach(function (idx) {
        var row = 라인.rows[idx], qty = toNum_(row[L.필요]), code = toStr_(row[L.상품코드]);
        row[L.확인] = ENUM.확인.정상; row[L.실제] = qty;
        row[L.상태] = ENUM.라인상태.완료; row[L.처리일시] = now;
        if (L.담당자 >= 0 && workers[no]) row[L.담당자] = workers[no];
        logs.push({ 구분: ENUM.로그구분.출고, 피킹지시번호: instructionNo, 주문번호: no,
          품목별주문번호: toStr_(row[L.품목별]), 상품코드: code, 변동량: -qty,
          변동후재고: toNum_(마스터.rows[masterIndex[code]][M.가용]), 담당자: workers[no] || actor,
          사유: '출력 성공 · 예약재고 소진' });
        lineCount++; quantity += qty;
      });
      (headerRows[no] || []).forEach(function (idx) {
        헤더.rows[idx][H.상태] = ENUM.헤더상태.완료;
        if (H.출력일시 >= 0 && isBlank_(헤더.rows[idx][H.출력일시])) 헤더.rows[idx][H.출력일시] = now;
      });
      (orderRows[no] || []).forEach(function (idx) {
        주문.rows[idx][O.상태] = ENUM.주문상태.출고완료; 주문.rows[idx][O.출고] = 1;
        if (O.대기사유 >= 0) 주문.rows[idx][O.대기사유] = '';
      });
    });

    if (라인.rows.length) {
      writeColumn_(라인.sheet, L.확인, 라인.rows); writeColumn_(라인.sheet, L.실제, 라인.rows);
      writeColumn_(라인.sheet, L.상태, 라인.rows); writeColumn_(라인.sheet, L.처리일시, 라인.rows);
      if (L.담당자 >= 0) writeColumn_(라인.sheet, L.담당자, 라인.rows);
    }
    if (헤더.rows.length) {
      writeColumn_(헤더.sheet, H.상태, 헤더.rows);
      if (H.출력일시 >= 0) writeColumn_(헤더.sheet, H.출력일시, 헤더.rows);
    }
    if (주문.rows.length) {
      writeColumn_(주문.sheet, O.상태, 주문.rows); writeColumn_(주문.sheet, O.출고, 주문.rows);
      if (O.대기사유 >= 0) writeColumn_(주문.sheet, O.대기사유, 주문.rows);
    }
    if (마스터.rows.length) {
      마스터.sheet.getRange(2, M.예약 + 1, reserved.length, 1).setValues(reserved.map(function (v) { return [v]; }));
    }
    writeStockLog_(logs);
    writeOpLog_('finalizePickingAfterOutput_', '성공', instructionNo + ' / 주문 ' + targetOrders.length + ' / 수량 ' + quantity);
    if (options.refresh !== false) try { D0_대시보드전체갱신(true); } catch (ignore) { }
    return { 이미완료: false, 지시번호: instructionNo, 완료주문: targetOrders, 차감라인: lineCount, 차감수량: quantity };
  });
}

/** 이전 설치/호출자 호환용. O는 더 이상 정상 출고의 입력 조건이 아니다. */
function S5_1_결과반영() {
  var exceptions = findPickingExceptionOrders_(), cancelled = [];
  exceptions.forEach(function (item) {
    if (!item.reason) return;
    var result = cancelOrder_(item.orderNo, item.reason, 'PICKING_X', { refresh: false });
    if (result.취소) cancelled.push({ 주문번호: item.orderNo, 사유: item.reason });
  });
  if (cancelled.length) try { D0_대시보드전체갱신(true); } catch (ignore) { }
  return { 차감라인: 0, 차감수량: 0, 취소주문: cancelled, 복원라인: 0, 복원수량: 0, 완료주문: [] };
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
