/** S3. 사용자에게 노출하지 않는 주문 재고 판정/예약 서비스. */
function S3_1_주문확정(orderNos, options) {
  options = options || {};
  return withLock_(function () {
    var 주문 = readTable_(ROLE.주문);
    var O = {
      주문번호: col_(주문, COL.주문번호, true), 상품코드: col_(주문, COL.상품품목코드, true),
      수량: col_(주문, COL.수량, true), 주문상태: col_(주문, COL.주문상태, true),
      확정일시: col_(주문, COL.확정일시, false), 대기사유: col_(주문, COL.대기사유, false)
    };
    var 대상 = {};
    (orderNos || []).forEach(function (no) { 대상[toStr_(no)] = true; });
    var 제한 = true; // 전체 예약 주문을 암묵적으로 release하지 않는다.
    var 그룹 = {}, 순서 = [];
    주문.rows.forEach(function (row) {
      var no = toStr_(row[O.주문번호]);
      if (!no || (제한 && !대상[no]) || toStr_(row[O.주문상태]) !== ENUM.주문상태.예약) return;
      if (!그룹[no]) { 그룹[no] = []; 순서.push(no); }
      그룹[no].push(row);
    });

    var 마스터 = readTable_(ROLE.마스터);
    var M = {
      코드: col_(마스터, COL.상품품목코드, true), 가용: col_(마스터, COL.가용재고, true),
      예약: col_(마스터, COL.예약재고, true), 상품상태: col_(마스터, COL.상품상태, false),
      예약상품: col_(마스터, COL.예약상품, false)
    };
    var 행 = {}, 가용 = [], 예약 = [], 예약상품 = {}, 사용중지 = {};
    마스터.rows.forEach(function (row, idx) {
      var code = toStr_(row[M.코드]);
      가용[idx] = toNum_(row[M.가용]); 예약[idx] = toNum_(row[M.예약]);
      if (!code) return;
      행[code] = idx;
      예약상품[code] = M.예약상품 >= 0 && toStr_(row[M.예약상품]).toUpperCase() === 'Y';
      사용중지[code] = M.상품상태 >= 0 && toStr_(row[M.상품상태]) === '사용중지';
    });

    var result = { 준비: 0, 예약: 0, 준비주문: [], 부족목록: [] };
    var logs = [], now = new Date(), actor = 사용자_();
    순서.sort().forEach(function (no) {
      var items = 그룹[no], required = {};
      items.forEach(function (row) {
        var code = toStr_(row[O.상품코드]);
        required[code] = (required[code] || 0) + toNum_(row[O.수량]);
      });

      // 확정일시는 예약재고 이동이 이미 끝났음을 나타내는 내부 idempotency 표식이다.
      var alreadyReserved = O.확정일시 >= 0 && items.some(function (row) { return !isBlank_(row[O.확정일시]); });
      if (alreadyReserved) { result.준비++; result.준비주문.push(no); return; }

      var shortage = [];
      Object.keys(required).forEach(function (code) {
        var mi = 행[code];
        if (mi === undefined) shortage.push(code + ' 미등록');
        else if (사용중지[code]) shortage.push(code + ' 사용중지');
        else if (!options.manualRelease && 예약상품[code]) shortage.push(code + ' 예약상품');
        else if (가용[mi] < required[code]) shortage.push(code + ' 필요 ' + required[code] + ' / 가용 ' + 가용[mi]);
      });
      if (shortage.length) {
        items.forEach(function (row) { if (O.대기사유 >= 0) row[O.대기사유] = shortage.join(' | ').substring(0, 250); });
        result.예약++;
        result.부족목록.push({ 주문번호: no, 상세: shortage });
        return;
      }

      Object.keys(required).forEach(function (code) {
        var mi = 행[code], qty = required[code];
        가용[mi] -= qty; 예약[mi] += qty;
        logs.push({ 구분: ENUM.로그구분.예약, 피킹지시번호: '', 주문번호: no, 품목별주문번호: '',
          상품코드: code, 변동량: -qty, 변동후재고: 가용[mi], 담당자: actor,
          사유: (options.manualRelease ? '예약 주문 release' : '자동 주문 처리') + ' · 예약 ' + qty + '개' });
      });
      items.forEach(function (row) {
        if (O.확정일시 >= 0) row[O.확정일시] = now;
        if (O.대기사유 >= 0) row[O.대기사유] = '';
      });
      result.준비++; result.준비주문.push(no);
    });

    if (주문.rows.length) {
      if (O.확정일시 >= 0) writeColumn_(주문.sheet, O.확정일시, 주문.rows);
      if (O.대기사유 >= 0) writeColumn_(주문.sheet, O.대기사유, 주문.rows);
    }
    if (마스터.rows.length) {
      마스터.sheet.getRange(2, M.가용 + 1, 가용.length, 1).setValues(가용.map(function (v) { return [v]; }));
      마스터.sheet.getRange(2, M.예약 + 1, 예약.length, 1).setValues(예약.map(function (v) { return [v]; }));
    }
    writeStockLog_(logs);
    writeOpLog_('S3_1_주문확정', '성공', '처리 준비 ' + result.준비 + ' / 예약 ' + result.예약);
    if (!options.silent) alert_('재고 판정 완료\n처리 준비 ' + result.준비 + '건 / 예약 ' + result.예약 + '건');
    return result;
  });
}
