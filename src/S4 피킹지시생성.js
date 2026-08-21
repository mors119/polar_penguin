/**
 * ============================================================
 *  S4. 피킹지시 생성
 * ============================================================
 *  재고 확정이 끝난 지정 주문을 찾아 주문별 피킹헤더와 품목별 피킹라인을
 *  생성한다. 피킹지시번호가 재생성 방지와 출력 단위를 나타내며, 카트/작업자
 *  배정은 창고의 물리 운영에 맡기고 시스템 데이터로 만들지 않는다.
 * ============================================================
 */

/**
 * 확정된 미지시 주문을 주문 단위의 피킹헤더와 품목 단위의 피킹라인으로 만든다.
 *
 * @return {Object} 생성 여부와 생성된 지시번호를 포함한 결과
 * @sideEffect 피킹헤더·피킹라인을 추가하고 주문완료에 피킹지시번호를 기록한다.
 */
function S4_1_피킹지시생성(orderNos, options) {
  options = options || {};
  return withLock_(function () {
    // ---------- 주문 적재 ----------
    var 주문 = readTable_(ROLE.주문);
    var o = {
      주문번호: col_(주문, COL.주문번호, true),
      품목별: col_(주문, COL.품목별주문번호, true),
      상품코드: col_(주문, COL.상품품목코드, true),
      수량: col_(주문, COL.수량, true),
      출고완료: col_(주문, COL.출고완료, true),
      지시번호: col_(주문, COL.피킹지시번호, true),
      주문상태: col_(주문, COL.주문상태, true)
    };

    var 지정 = {};
    (orderNos || []).forEach(function (no) { 지정[toStr_(no)] = true; });
    var 제한 = true; // 자동/수동 오케스트레이터가 명시한 주문만 생성한다.
    var 기존지시 = {};
    주문.rows.forEach(function (row) {
      var no = toStr_(row[o.주문번호]);
      var instruction = toStr_(row[o.지시번호]);
      if (제한 && 지정[no] && instruction) 기존지시[instruction] = true;
    });
    if (Object.keys(기존지시).length === 1) {
      return { 생성: false, 재사용: true, 지시번호: Object.keys(기존지시)[0], 주문수: Object.keys(지정).length };
    }

    // ---------- 대상: 재고 확정이 끝난 지정 주문 중 아직 지시가 없는 것 ----------
    var 대상 = [];
    for (var i = 0; i < 주문.rows.length; i++) {
      var r = 주문.rows[i];
      if (isBlank_(r[o.주문번호])) continue;
      if (제한 && !지정[toStr_(r[o.주문번호])]) continue;
      if (toNum_(r[o.출고완료]) === 1) continue;
      if (!isBlank_(r[o.지시번호])) continue;
      if (toStr_(r[o.주문상태]) !== ENUM.주문상태.예약) continue;
      대상.push({ rowIdx: i, row: r });
    }

    if (!대상.length) {
      if (!options.silent) alert_('새로 생성할 피킹지시가 없습니다.');
      return { 생성: false };
    }

    // ---------- 마스터 조회 맵 ----------
    var 마스터맵 = buildMasterMap_();

    // ---------- 상품코드 존재 검증 ----------
    var 제외주문 = {};
    대상.forEach(function (t) {
      var code = toStr_(t.row[o.상품코드]);
      if (!마스터맵[code]) {
        var no = toStr_(t.row[o.주문번호]);
        if (!제외주문[no]) 제외주문[no] = [];
        if (제외주문[no].indexOf(code) < 0) 제외주문[no].push(code);
      }
    });

    var 유효 = 대상.filter(function (t) { return !제외주문[toStr_(t.row[o.주문번호])]; });
    if (!유효.length) {
      var 사유 = Object.keys(제외주문).map(function (k) {
        return '  · ' + k + ' → ' + 제외주문[k].join(', ');
      }).join('\n');
      if (!options.silent) alert_('모든 주문이 상품마스터 미등록 코드를 포함해 제외되었습니다.\n\n' + 사유);
      return { 생성: false };
    }

    // ---------- 헤더 적재 + 채번 ----------
    var 헤더 = readTable_(ROLE.헤더);
    var h = {
      지시번호: col_(헤더, COL.피킹지시번호, true),
      주문번호: col_(헤더, COL.주문번호, true),
      슬롯: col_(헤더, COL.카트슬롯, false),
      품목수: col_(헤더, COL.품목수, true),
      총수량: col_(헤더, COL.총수량, true),
      담당자: col_(헤더, COL.피킹담당자, false),
      상태: col_(헤더, COL.상태, true),
      생성일시: col_(헤더, COL.생성일시, false),
      출력일시: col_(헤더, COL.출력일시, false)
    };

    var 지시번호 = nextInstructionNo_(헤더, h.지시번호, options.reservationBatch ? 'RES' : '');

    // ---------- 주문번호 그룹핑 ----------
    var 그룹 = {}, 순서 = [];
    유효.forEach(function (t) {
      var no = toStr_(t.row[o.주문번호]);
      if (!그룹[no]) { 그룹[no] = []; 순서.push(no); }
      그룹[no].push(t);
    });
    순서.sort();

    // ---------- 헤더 행 구성 ----------
    var 헤더행 = [];

    순서.forEach(function (no) {
      var 총수량 = 0;
      그룹[no].forEach(function (t) {
        총수량 += toNum_(t.row[o.수량]);
      });

      var row = new Array(헤더.width).fill('');
      row[h.지시번호] = 지시번호;
      row[h.주문번호] = no;
      if (h.슬롯 >= 0) row[h.슬롯] = '';       // 기존 시트 호환용 legacy 열
      row[h.품목수] = 그룹[no].length;
      row[h.총수량] = 총수량;
      if (h.담당자 >= 0) row[h.담당자] = '';   // 기존 시트 호환용 선택 감사 열
      row[h.상태] = ENUM.헤더상태.대기;
      if (h.생성일시 >= 0) row[h.생성일시] = new Date();
      헤더행.push(row);
    });

    // ---------- 라인 구성 ----------
    var 라인결과 = buildPickingLines_(지시번호, 순서, 그룹, o, 마스터맵);

    // ---------- 교차 검증 ----------
    var 검증오류 = [];
    헤더행.forEach(function (hr) {
      var no = hr[h.주문번호];
      var stat = 라인결과.집계[no] || { 건수: 0, 수량: 0 };
      if (stat.건수 !== hr[h.품목수]) 검증오류.push(no + ': 라인 ' + stat.건수 + ' ≠ 품목수 ' + hr[h.품목수]);
      if (stat.수량 !== hr[h.총수량]) 검증오류.push(no + ': 수량합 ' + stat.수량 + ' ≠ 총수량 ' + hr[h.총수량]);
    });
    if (검증오류.length) {
      alert_('검증 실패로 생성을 중단했습니다.\n\n' + 검증오류.join('\n'));
      writeOpLog_('S4_1_피킹지시생성', '실패', 검증오류.join(' / '));
      return { 생성: false };
    }

    // ---------- 쓰기 (최신이 위로) ----------
    var 헤더삽입 = false, 라인삽입 = false, 라인행수 = 0, 라인시트 = null;
    try {
      prependRows_(헤더.sheet, 헤더행, [h.지시번호, h.주문번호]);
      헤더삽입 = true;

      var 라인 = readTable_(ROLE.라인);
      var 생성라인 = 라인결과.행(라인);
      prependRows_(라인.sheet, 생성라인, [
        col_(라인, COL.상품코드, true),
        col_(라인, COL.품목별주문번호, true),
        col_(라인, COL.피킹지시번호, true),
        col_(라인, COL.보관위치, true)
      ]);
      라인삽입 = true; 라인행수 = 생성라인.length; 라인시트 = 라인.sheet;

      // 주문 지시번호까지 성공해야 이번 피킹 데이터 쓰기를 완료한 것으로 본다.
      var 지시컬럼 = 주문.rows.map(function (r) { return [r[o.지시번호]]; });
      유효.forEach(function (t) { 지시컬럼[t.rowIdx][0] = 지시번호; });
      주문.sheet.getRange(2, o.지시번호 + 1, 지시컬럼.length, 1).setValues(지시컬럼);
    } catch (e) {
      // 헤더·라인·주문 연결 중 하나라도 실패하면 이번에 prepend한 피킹 행을 되돌린다.
      if (라인삽입 && 라인시트) {
        try { 라인시트.deleteRows(2, 라인행수); } catch (e1) { }
      }
      if (헤더삽입) {
        try { 헤더.sheet.deleteRows(2, 헤더행.length); } catch (e2) { }
      }
      writeOpLog_('S4_1_피킹지시생성', '실패', e.message);
      throw new Error('피킹라인 생성 중 오류가 발생해 전체를 취소했습니다.\n' + e.message);
    }

    // ---------- 결과 ----------
    var 총품목 = 헤더행.reduce(function (a, r) { return a + r[h.품목수]; }, 0);
    var 총수량합 = 헤더행.reduce(function (a, r) { return a + r[h.총수량]; }, 0);

    var msg = '피킹지시 생성 완료\n\n' +
      '배치번호: ' + 지시번호 + '\n' +
      '주문 ' + 순서.length + '건 / 품목라인 ' + 총품목 + '개 / 총수량 ' + 총수량합 + '\n' +
      'PDF에는 상품별 총 피킹수량과 주문별 포장 내역이 함께 표시됩니다.';

    if (Object.keys(제외주문).length) {
      msg += '\n\n⚠ 제외된 주문 (마스터 미등록)\n' +
        Object.keys(제외주문).map(function (k) { return '  · ' + k + ' → ' + 제외주문[k].join(', '); }).join('\n');
    }

    if (!options.silent) alert_(msg);
    writeOpLog_('S4_1_피킹지시생성', '성공', 지시번호 + ' / 주문 ' + 순서.length + '건');
    return { 생성: true, 지시번호: 지시번호, 주문수: 순서.length };
  });
}

/* ============================================================
 *  라인 구성
 * ============================================================ */

function buildPickingLines_(지시번호, 순서, 그룹, o, 마스터맵) {
  var 집계 = {}, 원본 = [];

  순서.forEach(function (no) {
    var items = 그룹[no].slice().sort(function (a, b) {
      return toStr_(a.row[o.품목별]) < toStr_(b.row[o.품목별]) ? -1 : 1;
    });

    집계[no] = { 건수: 0, 수량: 0 };

    items.forEach(function (t, i) {
      var code = toStr_(t.row[o.상품코드]);
      var qty = toNum_(t.row[o.수량]);
      원본.push({
        주문번호: no,
        순번: i + 1,
        상품코드: code,
        필요수량: qty,
        품목별: toStr_(t.row[o.품목별]),
        m: 마스터맵[code] || {}
      });
      집계[no].건수++;
      집계[no].수량 += qty;
    });
  });

  return {
    집계: 집계,
    행: function (라인) {
      var c = {
        순번: col_(라인, COL.순번, true),
        주문번호: col_(라인, COL.주문번호, true),
        보관위치: col_(라인, COL.보관위치, true),
        상품코드: col_(라인, COL.상품코드, true),
        이미지: col_(라인, COL.이미지, true),
        상품명: col_(라인, COL.상품명, true),
        옵션: col_(라인, COL.옵션, true),
        필요수량: col_(라인, COL.필요수량, true),
        확인: col_(라인, COL.확인, true),
        실제수량: col_(라인, COL.실제수량, true),
        예외사유: col_(라인, COL.예외사유, true),
        품목별: col_(라인, COL.품목별주문번호, true),
        지시번호: col_(라인, COL.피킹지시번호, true),
        담당자: col_(라인, COL.담당자, false),
        라인상태: col_(라인, COL.라인상태, true),
        처리일시: col_(라인, COL.처리일시, true)
      };

      return 원본.map(function (x) {
        var row = new Array(라인.width).fill('');
        row[c.순번] = x.순번;
        row[c.주문번호] = x.주문번호;
        row[c.보관위치] = x.m.보관위치 || '';
        row[c.이미지] = ''; // 피킹 라인의 레거시 필드이며 상품마스터에서는 더 이상 공급하지 않는다.
        row[c.상품명] = x.m.상품명 || '';
        row[c.옵션] = x.m.옵션명 || '';
        row[c.상품코드] = x.상품코드;
        row[c.필요수량] = x.필요수량;
        row[c.확인] = '';
        row[c.실제수량] = '';
        row[c.예외사유] = '';
        row[c.품목별] = x.품목별;
        row[c.지시번호] = 지시번호;
        if (c.담당자 >= 0) row[c.담당자] = '';
        row[c.라인상태] = ENUM.라인상태.미처리;
        row[c.처리일시] = '';
        return row;
      });
    }
  };
}

/* ============================================================
 *  공용 헬퍼
 * ============================================================ */

function buildMasterMap_() {
  var m = readTable_(ROLE.마스터);
  var c = {
    코드: col_(m, COL.상품품목코드, true),
    상품명: col_(m, COL.상품명, true),
    옵션명: col_(m, COL.옵션명, false),
    위치: col_(m, COL.기본보관위치, true),
    재고: col_(m, COL.가용재고, true)
  };

  var map = {};
  for (var i = 0; i < m.rows.length; i++) {
    var code = toStr_(m.rows[i][c.코드]);
    if (!code) continue;
    map[code] = {
      상품명: toStr_(m.rows[i][c.상품명]),
      옵션명: c.옵션명 >= 0 ? toStr_(m.rows[i][c.옵션명]) : '',
      보관위치: toStr_(m.rows[i][c.위치]),
      가용재고: toNum_(m.rows[i][c.재고]),
      rowIdx: i
    };
  }
  return map;
}

/** 당일 다음 배치번호 */
function nextInstructionNo_(헤더, c지시번호, batchType) {
  var 접두어 = String(param_('지시번호접두어', 'PK'));
  var 오늘 = Utilities.formatDate(new Date(), tz_(), 'yyyyMMdd');
  var prefix = 접두어 + '-' + 오늘 + '-' + (batchType ? batchType + '-' : '');

  var max = 0;
  헤더.rows.forEach(function (r) {
    var v = toStr_(r[c지시번호]);
    if (v.indexOf(prefix) === 0) {
      var n = parseInt(v.substring(prefix.length), 10);
      if (!isNaN(n) && n > max) max = n;
    }
  });
  return prefix + ('00' + (max + 1)).slice(-3);
}
