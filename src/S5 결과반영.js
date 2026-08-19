/**
 * ============================================================
 *  S5. 피킹 결과 반영
 * ============================================================
 *  피킹라인의 O/X 결과를 주문 단위로 반영한다. 하나라도 X이면 부분 출고하지
 *  않고 주문 전체를 취소하며, O인 주문은 S3에서 확보한 예약재고를 소진한다.
 *
 *  취소가 차감보다 먼저 처리되므로 같은 주문의 일부 품목만 재고에서 빠지지 않는다.
 *  이미 차감된 주문을 취소할 때는 재고이동로그의 순액을 근거로 복원한다.
 *  실행: 설정의 폴링주기(분) 트리거(syncAndRefresh) 또는 메뉴 수동 실행
 * ============================================================
 */

/**
 * 피킹라인의 확인 결과를 주문·재고·피킹 상태에 반영한다.
 *
 * @return {Object} 차감, 취소, 복원, 완료 건수를 담은 처리 요약
 * @sideEffect 예약재고를 소진하거나 해제하고 관련 시트와 재고이동로그를 갱신한다.
 */
function S5_1_결과반영() {
  return withLock_(function () {
    var 시작 = new Date();

    // ---------- 적재 ----------
    var 라인 = readTable_(ROLE.라인);
    var L = {
      상품코드: col_(라인, COL.상품코드, true),
      상품명: col_(라인, COL.상품명, true),
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
    if (!라인.rows.length) return 빈결과_();

    var 주문 = readTable_(ROLE.주문);
    var O = {
      주문번호: col_(주문, COL.주문번호, true),
      품목별: col_(주문, COL.품목별주문번호, true),
      출고완료: col_(주문, COL.출고완료, true),
      주문상태: col_(주문, COL.주문상태, true),
      취소사유: col_(주문, COL.취소사유, true),
      취소일시: col_(주문, COL.취소일시, true)
    };

    var 헤더 = readTable_(ROLE.헤더);
    var H = {
      지시번호: col_(헤더, COL.피킹지시번호, true),
      주문번호: col_(헤더, COL.주문번호, true),
      슬롯: col_(헤더, COL.카트슬롯, true),
      담당자: col_(헤더, COL.피킹담당자, true),
      상태: col_(헤더, COL.상태, true)
    };

    var 마스터 = readTable_(ROLE.마스터);
    var M = {
      코드: col_(마스터, COL.상품품목코드, true),
      가용: col_(마스터, COL.가용재고, true),
      예약: col_(마스터, COL.예약재고, true),
      불량: col_(마스터, COL.불량재고, false)
    };

    // ---------- 조인 맵 ----------
    var 품목별_주문 = {}, 주문번호_행 = {};
    for (var i = 0; i < 주문.rows.length; i++) {
      var k = toStr_(주문.rows[i][O.품목별]);
      var no = toStr_(주문.rows[i][O.주문번호]);
      if (k) 품목별_주문[k] = no;
      if (no) { (주문번호_행[no] = 주문번호_행[no] || []).push(i); }
    }

    var 헤더행 = {};
    for (var j = 0; j < 헤더.rows.length; j++) {
      헤더행[toStr_(헤더.rows[j][H.지시번호]) + '|' + toStr_(헤더.rows[j][H.주문번호])] = j;
    }

    var 마스터행 = {};
    for (var k2 = 0; k2 < 마스터.rows.length; k2++) {
      var code = toStr_(마스터.rows[k2][M.코드]);
      if (code) 마스터행[code] = k2;
    }

    var 로그순액 = readStockLogNet_();

    // ---------- 라인 그룹핑 ----------
    var 그룹 = {};
    for (var n = 0; n < 라인.rows.length; n++) {
      var r = 라인.rows[n];
      var 품목별키 = toStr_(r[L.품목별]);
      if (!품목별키) continue;
      var 주문no = 품목별_주문[품목별키];
      if (!주문no) continue;

      var key = toStr_(r[L.지시번호]) + '|' + 주문no;
      if (!그룹[key]) 그룹[key] = { 지시번호: toStr_(r[L.지시번호]), 주문번호: 주문no, lines: [] };
      그룹[key].lines.push(n);
    }

    // ---------- 버퍼 ----------
    var 가용 = 마스터.rows.map(function (r) { return toNum_(r[M.가용]); });
    var 예약 = 마스터.rows.map(function (r) { return toNum_(r[M.예약]); });
    var 로그 = [], now = new Date();
    var 요약 = { 차감라인: 0, 차감수량: 0, 취소주문: [], 복원라인: 0, 복원수량: 0, 완료주문: [], 담당자전파: 0 };

    // ---------- 0단계. 담당자 전파 + 실제수량 정규화 ----------
    Object.keys(그룹).forEach(function (key) {
      var g = 그룹[key];
      var hIdx = 헤더행[key];
      var 담당 = hIdx !== undefined ? toStr_(헤더.rows[hIdx][H.담당자]) : '';

      g.lines.forEach(function (idx) {
        var r = 라인.rows[idx];

        // 헤더에 한 번 입력한 피킹담당자를 같은 주문의 모든 라인에 전파한다.
        if (L.담당자 >= 0 && 담당 && toStr_(r[L.담당자]) !== 담당) {
          r[L.담당자] = 담당;
          요약.담당자전파++;
        }

        if (toStr_(r[L.라인상태]) !== ENUM.라인상태.미처리) return;
        var 확인값 = toStr_(r[L.확인]).toUpperCase();
        if (확인값 === ENUM.확인.정상) r[L.실제수량] = toNum_(r[L.필요수량]);
        else if (확인값 === ENUM.확인.예외) r[L.실제수량] = 0;
      });
    });

    // ---------- 1단계. 주문 전체 취소 판정 (차감보다 먼저) ----------
    var 취소대상 = [];
    Object.keys(그룹).forEach(function (key) {
      var 예외라인 = null;
      그룹[key].lines.forEach(function (idx) {
        var r = 라인.rows[idx];
        if (toStr_(r[L.확인]).toUpperCase() === ENUM.확인.예외 &&
            toStr_(r[L.라인상태]) !== ENUM.라인상태.취소마감) {
          if (예외라인 === null) 예외라인 = idx;
        }
      });
      if (예외라인 !== null) 취소대상.push({ key: key, 예외라인: 예외라인 });
    });

    취소대상.forEach(function (t) {
      var g = 그룹[t.key];
      var hIdx = 헤더행[t.key];
      var 담당 = hIdx !== undefined ? toStr_(헤더.rows[hIdx][H.담당자]) : '';
      var 사유 = toStr_(라인.rows[t.예외라인][L.예외사유]) || '미기재';
      var 복원건 = 0, 복원량 = 0;

      g.lines.forEach(function (idx) {
        var r = 라인.rows[idx];
        var 상태 = toStr_(r[L.라인상태]);
        var 품목별키 = toStr_(r[L.품목별]);
        var code = toStr_(r[L.상품코드]);

        if (상태 === ENUM.라인상태.차감완료) {
          // 복원량은 재고이동로그의 실제 차감 기록을 근거로 한다
          var 순액 = 로그순액[품목별키] || 0;
          var 복원 = 순액 < 0 ? -순액 : 0;
          var mi = 마스터행[code];

          if (복원 > 0 && mi !== undefined) {
            가용[mi] += 복원;      // 선반으로 복귀
            로그.push({
              구분: ENUM.로그구분.복원, 피킹지시번호: g.지시번호, 주문번호: g.주문번호,
              품목별주문번호: 품목별키, 상품코드: code,
              변동량: 복원, 변동후재고: 가용[mi],
              담당자: 담당, 사유: '주문취소-복원 (' + 사유 + ')'
            });
            복원건++; 복원량 += 복원;
          }
          r[L.라인상태] = ENUM.라인상태.복원완료;
          r[L.처리일시] = now;

        } else if (상태 !== ENUM.라인상태.복원완료) {
          // 아직 차감 전이면 예약을 풀어 가용으로 되돌린다
          var mi2 = 마스터행[code];
          var 필요 = toNum_(r[L.필요수량]);
          if (mi2 !== undefined && 예약[mi2] > 0 && 필요 > 0) {
            var 해제 = Math.min(필요, 예약[mi2]);
            예약[mi2] -= 해제;
            가용[mi2] += 해제;
            로그.push({
              구분: ENUM.로그구분.예약해제, 피킹지시번호: g.지시번호, 주문번호: g.주문번호,
              품목별주문번호: 품목별키, 상품코드: code,
              변동량: 해제, 변동후재고: 가용[mi2],
              담당자: 담당, 사유: '주문취소-예약해제 (' + 사유 + ')'
            });
          }
          r[L.라인상태] = ENUM.라인상태.취소마감;
          r[L.처리일시] = now;
        }
      });

      if (hIdx !== undefined) 헤더.rows[hIdx][H.상태] = ENUM.헤더상태.예외;

      (주문번호_행[g.주문번호] || []).forEach(function (oi) {
        주문.rows[oi][O.주문상태] = ENUM.주문상태.취소;
        주문.rows[oi][O.취소사유] = 사유;
        주문.rows[oi][O.취소일시] = now;
        주문.rows[oi][O.출고완료] = 0;
      });

      요약.취소주문.push({
        주문번호: g.주문번호,
        슬롯: hIdx !== undefined ? 헤더.rows[hIdx][H.슬롯] : '',
        담당자: 담당, 사유: 사유, 복원건: 복원건, 복원량: 복원량
      });
      요약.복원라인 += 복원건;
      요약.복원수량 += 복원량;
    });

    var 취소키 = {};
    취소대상.forEach(function (t) { 취소키[t.key] = true; });

    // ---------- 2단계. 정상 피킹 차감 (예약재고에서) ----------
    Object.keys(그룹).forEach(function (key) {
      if (취소키[key]) return;
      var g = 그룹[key];
      var hIdx = 헤더행[key];
      var 담당 = hIdx !== undefined ? toStr_(헤더.rows[hIdx][H.담당자]) : '';

      g.lines.forEach(function (idx) {
        var r = 라인.rows[idx];
        if (toStr_(r[L.라인상태]) !== ENUM.라인상태.미처리) return;
        if (toStr_(r[L.확인]).toUpperCase() !== ENUM.확인.정상) return;

        var code = toStr_(r[L.상품코드]);
        var qty = toNum_(r[L.실제수량]);
        var mi = 마스터행[code];

        if (mi === undefined) {
          writeOpLog_('S5_1_결과반영', '경고', '마스터 미등록 코드: ' + code);
          return;
        }
        if (qty <= 0) return;

        예약[mi] -= qty;               // S3에서 확보한 예약재고를 실제 출고 시점에 소진한다.
        if (예약[mi] < 0) 예약[mi] = 0;

        로그.push({
          구분: ENUM.로그구분.차감, 피킹지시번호: g.지시번호, 주문번호: g.주문번호,
          품목별주문번호: toStr_(r[L.품목별]), 상품코드: code,
          변동량: -qty, 변동후재고: 가용[mi],
          담당자: 담당, 사유: '피킹 출고 · 예약 ' + qty + '개 소진'
        });

        r[L.라인상태] = ENUM.라인상태.차감완료;
        r[L.처리일시] = now;
        요약.차감라인++;
        요약.차감수량 += qty;
      });
    });

    // ---------- 3단계. 헤더 상태 · 출고완료 ----------
    Object.keys(그룹).forEach(function (key) {
      var g = 그룹[key];
      var hIdx = 헤더행[key];
      if (hIdx === undefined || 취소키[key]) return;

      var 전체 = g.lines.length, 차감 = 0, 취소 = 0;
      g.lines.forEach(function (idx) {
        var s = toStr_(라인.rows[idx][L.라인상태]);
        if (s === ENUM.라인상태.차감완료) 차감++;
        if (s === ENUM.라인상태.취소마감 || s === ENUM.라인상태.복원완료) 취소++;
      });

      var 상태 = 취소 > 0 ? ENUM.헤더상태.예외
               : 차감 === 전체 ? ENUM.헤더상태.완료
               : 차감 > 0 ? ENUM.헤더상태.진행
               : ENUM.헤더상태.대기;

      헤더.rows[hIdx][H.상태] = 상태;

      if (상태 === ENUM.헤더상태.완료) {
        (주문번호_행[g.주문번호] || []).forEach(function (oi) {
          주문.rows[oi][O.출고완료] = 1;
        });
        요약.완료주문.push(g.주문번호);
      }
    });

    // ---------- 일괄 쓰기 ----------
    writeColumn_(라인.sheet, L.실제수량, 라인.rows);
    writeColumn_(라인.sheet, L.라인상태, 라인.rows);
    writeColumn_(라인.sheet, L.처리일시, 라인.rows);
    if (L.담당자 >= 0) writeColumn_(라인.sheet, L.담당자, 라인.rows);

    writeColumn_(헤더.sheet, H.상태, 헤더.rows);

    writeColumn_(주문.sheet, O.출고완료, 주문.rows);
    writeColumn_(주문.sheet, O.주문상태, 주문.rows);
    writeColumn_(주문.sheet, O.취소사유, 주문.rows);
    writeColumn_(주문.sheet, O.취소일시, 주문.rows);

    마스터.sheet.getRange(2, M.가용 + 1, 가용.length, 1).setValues(가용.map(function (v) { return [v]; }));
    마스터.sheet.getRange(2, M.예약 + 1, 예약.length, 1).setValues(예약.map(function (v) { return [v]; }));

    writeStockLog_(로그);

    // ---------- 결과 ----------
    var 소요 = ((new Date() - 시작) / 1000).toFixed(1);
    var msg = '차감 ' + 요약.차감라인 + '라인(' + 요약.차감수량 + '개) / ' +
              '취소 ' + 요약.취소주문.length + '건 / ' +
              '복원 ' + 요약.복원라인 + '라인 / ' +
              '완료 ' + 요약.완료주문.length + '건 · ' + 소요 + '초';

    if (요약.취소주문.length) {
      msg += '\n\n⚠ 취소된 주문 — 카트의 상품을 원위치하세요\n' +
        요약.취소주문.map(function (c) {
          return '  · 슬롯 ' + c.슬롯 + ' · ' + c.주문번호 + ' · ' + (c.담당자 || '담당 미지정') + ' · ' + c.사유;
        }).join('\n');
    }

    toast_(msg, '재고 반영 완료');
    writeOpLog_('S5_1_결과반영', '성공', msg.replace(/\n/g, ' | '));
    return 요약;
  });
}

function 빈결과_() {
  return { 차감라인: 0, 차감수량: 0, 취소주문: [], 복원라인: 0, 복원수량: 0, 완료주문: [], 담당자전파: 0 };
}

/** 메뉴용 수동 실행 */
function S5_2_수동반영() {
  var r = S5_1_결과반영();
  alert_(
    '재고 반영 · 취소 처리 완료\n\n' +
    '차감: ' + r.차감라인 + '라인 (' + r.차감수량 + '개)\n' +
    '취소: ' + r.취소주문.length + '건\n' +
    '복원: ' + r.복원라인 + '라인\n' +
    '완료 처리된 주문: ' + r.완료주문.length + '건' +
    (r.담당자전파 ? '\n담당자 전파: ' + r.담당자전파 + '라인' : '') +
    (r.취소주문.length
      ? '\n\n⚠ 취소된 주문\n' + r.취소주문.map(function (c) {
          return '  · 슬롯 ' + c.슬롯 + ' · ' + c.주문번호 + ' (' + c.사유 + ')';
        }).join('\n')
      : '')
  );
}

/** 트리거 진입점 — 반영 후 대시보드 3종 갱신 */
function syncAndRefresh() {
  try {
    S5_1_결과반영();
  } finally {
    try { D1_주문현황갱신(); } catch (e) { writeOpLog_('D1_주문현황갱신', '실패', e.message); }
    try { D2_재고현황갱신(); } catch (e) { writeOpLog_('D2_재고현황갱신', '실패', e.message); }
    try { D3_피킹현황갱신(); } catch (e) { writeOpLog_('D3_피킹현황갱신', '실패', e.message); }
  }
}
