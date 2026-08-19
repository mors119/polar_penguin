/**
 * ============================================================
 *  S8. 예약(선주문) 관리
 * ============================================================
 *  예약상품이 하나라도 포함된 주문은 부분 확정하지 않고 주문 전체를
 *  「예약대기」 상태로 보관한다. 이 모듈은 대기 현황 조회와 공급 불가 시
 *  주문 단위 강제 취소를 담당하며, 재고 예약·해제는 수행하지 않는다.
 *
 *  S3는 상품마스터의 예약상품 값이 N으로 바뀌고 가용재고가 충분해진 주문만
 *  확정할 수 있다. 현황은 통합 대시보드와 예약대기 조회 탭에도 표시된다.
 * ============================================================
 */

/**
 * 예약대기 주문을 상품별로 집계해 사람이 읽을 수 있는 현황을 표시한다.
 *
 * @return {string} 현황 메시지. 대기 주문이 없으면 빈 문자열
 */
function S8_2_예약대기조회() {
  var d = collectPreorderData_();

  if (!d.상품.length) {
    alert_('예약대기 중인 주문이 없습니다.');
    return '';
  }

  var out = [
    '예약대기 현황',
    '주문 ' + d.주문수 + '건 · 품목 ' + d.상품.length + '종 · 대기수량 ' + d.총수량 +
      '개 · 부족 ' + d.총부족 + '개',
    ''
  ];

  d.상품.forEach(function (p) {
    var 표시 = p.부족 <= 0 ? '✅ 출고가능'
             : (p.현재고 > 0 ? '⚠ 일부가능' : '⏳ 입고대기');
    out.push('  ' + 표시 + '  ' + p.코드 +
      '  대기 ' + p.수량 + '개 (주문 ' + p.주문수 + '건)' +
      '  재고 ' + p.현재고 +
      (p.부족 > 0 ? '  부족 ' + p.부족 : '') +
      '\n      ' + p.상품명.substring(0, 40) + (p.옵션 ? ' / ' + p.옵션 : ''));
  });

  out.push('');
  out.push('재고가 들어오면 「S3. 주문 확정」을 실행하세요.');
  out.push('출고 가능한 주문이 자동으로 확정으로 전환됩니다.');

  var msg = out.join('\n');
  Logger.log(msg);
  alert_(msg.substring(0, 1500));
  return msg;
}

/** 예약대기 데이터를 상품별로 집계한다. 현재고 표시는 운영 판단용이며 상태를 변경하지 않는다. */
function collectPreorderData_() {
  var out = { 상품: [], 주문수: 0, 총수량: 0, 총부족: 0, 주문목록: [] };

  var 주문 = readTable_(ROLE.주문);
  var O = {
    주문번호: col_(주문, COL.주문번호, true),
    상품코드: col_(주문, COL.상품품목코드, true),
    수량: col_(주문, COL.수량, true),
    주문상태: col_(주문, COL.주문상태, true)
  };

  var 마스터 = readTable_(ROLE.마스터);
  var M = {
    코드: col_(마스터, COL.상품품목코드, true),
    상품명: col_(마스터, COL.상품명, true),
    옵션: col_(마스터, COL.옵션명, false),
    가용: col_(마스터, COL.가용재고, true)
  };

  var 정보 = {};
  마스터.rows.forEach(function (r) {
    var c = toStr_(r[M.코드]);
    if (!c) return;
    정보[c] = {
      상품명: toStr_(r[M.상품명]),
      옵션: M.옵션 >= 0 ? toStr_(r[M.옵션]) : '',
      가용: toNum_(r[M.가용])
    };
  });

  var 집계 = {}, 주문집합 = {};
  주문.rows.forEach(function (r) {
    if (toStr_(r[O.주문상태]) !== ENUM.주문상태.예약대기) return;

    var no = toStr_(r[O.주문번호]);
    var code = toStr_(r[O.상품코드]);
    var qty = toNum_(r[O.수량]);
    if (!code) return;

    주문집합[no] = true;

    if (!집계[code]) 집계[code] = { 수량: 0, 주문: {}, 최초: null, 최근: null };
    집계[code].수량 += qty;
    집계[code].주문[no] = true;

    var m = no.match(/^(\d{4})(\d{2})(\d{2})/);
    if (m) {
      var dt = m[1] + '-' + m[2] + '-' + m[3];
      if (!집계[code].최초 || dt < 집계[code].최초) 집계[code].최초 = dt;
      if (!집계[code].최근 || dt > 집계[code].최근) 집계[code].최근 = dt;
    }
  });

  Object.keys(집계).forEach(function (code) {
    var a = 집계[code];
    var info = 정보[code] || { 상품명: '(마스터 미등록)', 옵션: '', 가용: 0 };

    out.상품.push({
      코드: code,
      상품명: info.상품명,
      옵션: info.옵션 === '-' ? '' : info.옵션,
      주문수: Object.keys(a.주문).length,
      수량: a.수량,
      현재고: info.가용,
      부족: a.수량 - info.가용,
      최초: a.최초,
      최근: a.최근
    });

    out.총수량 += a.수량;
    out.총부족 += Math.max(a.수량 - info.가용, 0);
  });

  // 운영자가 입고 우선순위를 보기 쉽도록 부족한 상품과 대기 수량이 큰 상품을 먼저 표시한다.
  out.상품.sort(function (a, b) {
    if ((a.부족 > 0) !== (b.부족 > 0)) return a.부족 > 0 ? -1 : 1;
    return b.수량 - a.수량;
  });

  out.주문수 = Object.keys(주문집합).length;
  out.주문목록 = Object.keys(주문집합).sort();
  return out;
}

/**
 * 더 이상 공급할 수 없는 상품이 포함된 예약대기 주문을 주문 전체 단위로 취소한다.
 *
 * @param {string=} 상품코드 취소 기준 상품코드. 생략하면 UI에서 입력받는다.
 * @param {string=} 사유 취소사유. 생략하면 "예약 공급 불가"를 사용한다.
 * @return {number|undefined} 취소한 주문 수. UI에서 취소하면 undefined
 * @sideEffect 대상 주문의 주문상태·취소사유·취소일시를 갱신한다.
 */
function S8_3_예약대기취소(상품코드, 사유) {
  return withLock_(function () {
    if (!상품코드) {
      var ui = null;
      try { ui = SpreadsheetApp.getUi(); } catch (e) { }
      if (!ui) throw new Error('상품코드를 인자로 넘겨주세요. 예: S8_3_예약대기취소("P00000KV000A", "공급 중단")');
      var resp = ui.prompt('예약대기 취소',
        '취소할 상품코드를 입력하세요.\n그 상품이 포함된 예약대기 주문이 모두 취소됩니다.',
        ui.ButtonSet.OK_CANCEL);
      if (resp.getSelectedButton() !== ui.Button.OK) return;
      상품코드 = resp.getResponseText().trim();
      if (!상품코드) return;
    }
    사유 = 사유 || '예약 공급 불가';

    var 주문 = readTable_(ROLE.주문);
    var O = {
      주문번호: col_(주문, COL.주문번호, true),
      상품코드: col_(주문, COL.상품품목코드, true),
      주문상태: col_(주문, COL.주문상태, true),
      취소사유: col_(주문, COL.취소사유, true),
      취소일시: col_(주문, COL.취소일시, true),
      대기사유: col_(주문, COL.대기사유, false)
    };

    // 해당 상품의 품목만 지우지 않고, 함께 주문된 모든 품목을 같은 주문번호로 취소한다.
    var 대상주문 = {};
    주문.rows.forEach(function (r) {
      if (toStr_(r[O.주문상태]) !== ENUM.주문상태.예약대기) return;
      if (toStr_(r[O.상품코드]) !== 상품코드) return;
      대상주문[toStr_(r[O.주문번호])] = true;
    });

    var 건수 = Object.keys(대상주문).length;
    if (!건수) {
      alert_('해당 상품의 예약대기 주문이 없습니다: ' + 상품코드);
      return 0;
    }

    var now = new Date();
    주문.rows.forEach(function (r) {
      var no = toStr_(r[O.주문번호]);
      if (!대상주문[no]) return;
      if (toStr_(r[O.주문상태]) !== ENUM.주문상태.예약대기) return;
      r[O.주문상태] = ENUM.주문상태.취소;
      r[O.취소사유] = 사유 + ' (' + 상품코드 + ')';
      r[O.취소일시] = now;
      if (O.대기사유 >= 0) r[O.대기사유] = '';
    });

    writeColumn_(주문.sheet, O.주문상태, 주문.rows);
    writeColumn_(주문.sheet, O.취소사유, 주문.rows);
    writeColumn_(주문.sheet, O.취소일시, 주문.rows);
    if (O.대기사유 >= 0) writeColumn_(주문.sheet, O.대기사유, 주문.rows);

    var msg = '예약대기 주문 ' + 건수 + '건을 취소했습니다.\n상품: ' + 상품코드 + '\n사유: ' + 사유;
    alert_(msg);
    writeOpLog_('S8_3_예약대기취소', '성공', 상품코드 + ' / ' + 건수 + '건');
    return 건수;
  });
}
