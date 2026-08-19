/**
 * ============================================================
 *  S3. 주문 확정 (재고 검증)
 * ============================================================
 *  주문 시트의 「접수」·「예약대기」 주문을 읽어 재고를 검증한다.
 *
 *   일반 상품 — 재고 충분   →  확정      (가용 → 예약재고)
 *   일반 상품 — 재고 부족   →  취소
 *   예약 상품 — 재고 무관   →  예약대기  (재고 차감 없음)
 *
 *  한 주문에 일반 상품과 예약 상품이 섞여 있으면 주문 전체를 「예약대기」로 둔다.
 *  부분 출고를 하지 않기 때문이다.
 *
 *  예약대기 주문은 매 실행마다 다시 검증되므로,
 *  재고가 들어오면 자동으로 확정으로 전환된다.
 * ============================================================
 */

function S3_1_주문확정() {
  return withLock_(function () {
    // ---------- 적재 ----------
    var 주문 = readTable_(ROLE.주문);
    var O = {
      주문번호: col_(주문, COL.주문번호, true),
      품목별: col_(주문, COL.품목별주문번호, true),
      상품코드: col_(주문, COL.상품품목코드, true),
      수량: col_(주문, COL.수량, true),
      출고완료: col_(주문, COL.출고완료, true),
      지시번호: col_(주문, COL.피킹지시번호, true),
      주문상태: col_(주문, COL.주문상태, true),
      취소사유: col_(주문, COL.취소사유, true),
      취소일시: col_(주문, COL.취소일시, true),
      확정일시: col_(주문, COL.확정일시, false),
      대기사유: col_(주문, COL.대기사유, false)
    };

    // ---------- 대상: 접수 또는 예약대기 ----------
    var 그룹 = {}, 순서 = [];
    for (var i = 0; i < 주문.rows.length; i++) {
      var r = 주문.rows[i];
      var no = toStr_(r[O.주문번호]);
      if (!no) continue;

      var 상태 = toStr_(r[O.주문상태]);
      if (상태 && 상태 !== ENUM.주문상태.접수 && 상태 !== ENUM.주문상태.예약대기) continue;
      if (toNum_(r[O.출고완료]) === 1) continue;
      if (!isBlank_(r[O.지시번호])) continue;

      if (!그룹[no]) { 그룹[no] = []; 순서.push(no); }
      그룹[no].push({ rowIdx: i, row: r, 이전상태: 상태 });
    }

    if (!순서.length) {
      alert_('확정할 주문이 없습니다.\n\n(주문상태가 「접수」 또는 「예약대기」인 주문이 대상입니다)');
      return { 확정: 0, 예약대기: 0, 취소: 0 };
    }

    순서.sort();   // 접수 순 = 선착순으로 재고를 잡는다

    // ---------- 상품마스터 ----------
    var 마스터 = readTable_(ROLE.마스터);
    var M = {
      코드: col_(마스터, COL.상품품목코드, true),
      상품명: col_(마스터, COL.상품명, true),
      가용: col_(마스터, COL.가용재고, true),
      예약: col_(마스터, COL.예약재고, true),
      상품상태: col_(마스터, COL.상품상태, false),
      예약상품: col_(마스터, COL.예약상품, false)
    };

    var 행번호 = {}, 가용 = [], 예약재고 = [], 상품명 = {}, 예약구분 = {};
    마스터.rows.forEach(function (r, i) {
      가용[i] = toNum_(r[M.가용]);
      예약재고[i] = toNum_(r[M.예약]);
      var c = toStr_(r[M.코드]);
      if (!c) return;
      행번호[c] = i;
      상품명[c] = toStr_(r[M.상품명]);
      예약구분[c] = (M.예약상품 >= 0)
        ? (toStr_(r[M.예약상품]).toUpperCase() === 'Y')
        : isPreorderName_(상품명[c]);
    });

    // ---------- 주문별 판정 ----------
    var 로그 = [], 확정목록 = [], 취소목록 = [], 대기목록 = [];
    var now = new Date();
    var 사용자 = 사용자_();

    순서.forEach(function (no) {
      var items = 그룹[no];

      var 필요 = {};
      items.forEach(function (t) {
        var c = toStr_(t.row[O.상품코드]);
        필요[c] = (필요[c] || 0) + toNum_(t.row[O.수량]);
      });

      var 미등록 = [], 판매중지 = [], 부족 = [], 예약품목 = [];

      Object.keys(필요).forEach(function (code) {
        var mi = 행번호[code];
        if (mi === undefined) { 미등록.push(code); return; }

        if (M.상품상태 >= 0 && toStr_(마스터.rows[mi][M.상품상태]) === '사용중지') {
          판매중지.push(code);
          return;
        }

        // ★ 예약 상품은 재고를 보지 않는다
        if (예약구분[code]) {
          예약품목.push(code + '(' + 필요[code] + '개' +
            (가용[mi] > 0 ? ', 현재고 ' + 가용[mi] : '') + ')');
          return;
        }

        if (가용[mi] < 필요[code]) {
          부족.push(code + ' 필요 ' + 필요[code] + ' / 가용 ' + 가용[mi]);
        }
      });

      // ---------- 판정 1: 취소 ----------
      if (미등록.length || 판매중지.length || 부족.length) {
        var 사유 = []
          .concat(미등록.length ? ['미등록: ' + 미등록.join(', ')] : [])
          .concat(판매중지.length ? ['판매중지: ' + 판매중지.join(', ')] : [])
          .concat(부족.length ? ['재고부족: ' + 부족.join(' / ')] : [])
          .join(' | ');

        items.forEach(function (t) {
          t.row[O.주문상태] = ENUM.주문상태.취소;
          t.row[O.취소사유] = 사유.substring(0, 250);
          t.row[O.취소일시] = now;
          t.row[O.출고완료] = 0;
          if (O.대기사유 >= 0) t.row[O.대기사유] = '';
        });
        취소목록.push({ 주문번호: no, 사유: 사유 });
        return;
      }

      // ---------- 판정 2: 예약대기 ----------
      if (예약품목.length) {
        var 대기사유 = '예약상품 ' + 예약품목.join(', ');
        items.forEach(function (t) {
          t.row[O.주문상태] = ENUM.주문상태.예약대기;
          if (O.대기사유 >= 0) t.row[O.대기사유] = 대기사유.substring(0, 250);
          t.row[O.취소사유] = '';
          t.row[O.취소일시] = '';
        });
        대기목록.push({
          주문번호: no, 품목: 예약품목,
          신규: items[0].이전상태 !== ENUM.주문상태.예약대기
        });
        return;
      }

      // ---------- 판정 3: 확정 ----------
      Object.keys(필요).forEach(function (code) {
        var mi = 행번호[code];
        가용[mi] -= 필요[code];
        예약재고[mi] += 필요[code];

        로그.push({
          구분: ENUM.로그구분.예약, 피킹지시번호: '', 주문번호: no, 품목별주문번호: '',
          상품코드: code, 변동량: -필요[code], 변동후재고: 가용[mi],
          담당자: 사용자, 사유: '주문 확정 · 예약 ' + 필요[code] + '개'
        });
      });

      items.forEach(function (t) {
        t.row[O.주문상태] = ENUM.주문상태.확정;
        if (O.확정일시 >= 0) t.row[O.확정일시] = now;
        if (O.대기사유 >= 0) t.row[O.대기사유] = '';
        t.row[O.취소사유] = '';
        t.row[O.취소일시] = '';
      });

      var 총수량 = 0;
      Object.keys(필요).forEach(function (c) { 총수량 += 필요[c]; });
      확정목록.push({
        주문번호: no, 품목수: Object.keys(필요).length, 총수량: 총수량,
        예약해소: items[0].이전상태 === ENUM.주문상태.예약대기
      });
    });

    // ---------- 일괄 쓰기 ----------
    writeColumn_(주문.sheet, O.주문상태, 주문.rows);
    writeColumn_(주문.sheet, O.취소사유, 주문.rows);
    writeColumn_(주문.sheet, O.취소일시, 주문.rows);
    if (O.확정일시 >= 0) writeColumn_(주문.sheet, O.확정일시, 주문.rows);
    if (O.대기사유 >= 0) writeColumn_(주문.sheet, O.대기사유, 주문.rows);

    마스터.sheet.getRange(2, M.가용 + 1, 가용.length, 1)
      .setValues(가용.map(function (x) { return [x]; }));
    마스터.sheet.getRange(2, M.예약 + 1, 예약재고.length, 1)
      .setValues(예약재고.map(function (x) { return [x]; }));

    writeStockLog_(로그);

    // ---------- 보고 ----------
    var 해소 = 확정목록.filter(function (c) { return c.예약해소; }).length;

    var msg = '주문 확정 완료\n\n' +
      '확정 ' + 확정목록.length + '건' + (해소 ? ' (예약 해소 ' + 해소 + '건 포함)' : '') +
      '  /  예약대기 ' + 대기목록.length + '건' +
      '  /  취소 ' + 취소목록.length + '건\n';

    if (확정목록.length) {
      msg += '\n[확정]\n' + 확정목록.slice(0, 10).map(function (c) {
        return '  · ' + c.주문번호 + ' (품목 ' + c.품목수 + ' / 수량 ' + c.총수량 + ')' +
               (c.예약해소 ? '  ← 예약 해소' : '');
      }).join('\n');
      if (확정목록.length > 10) msg += '\n  … 외 ' + (확정목록.length - 10) + '건';
    }
    if (대기목록.length) {
      msg += '\n\n[예약대기 — 입고되면 자동 확정]\n' + 대기목록.slice(0, 10).map(function (c) {
        return '  · ' + c.주문번호 + ' → ' + c.품목.join(', ');
      }).join('\n');
      if (대기목록.length > 10) msg += '\n  … 외 ' + (대기목록.length - 10) + '건';
    }
    if (취소목록.length) {
      msg += '\n\n[취소]\n' + 취소목록.slice(0, 10).map(function (c) {
        return '  · ' + c.주문번호 + ' → ' + c.사유.substring(0, 60);
      }).join('\n');
      if (취소목록.length > 10) msg += '\n  … 외 ' + (취소목록.length - 10) + '건';
    }

    alert_(msg);
    writeOpLog_('S3_1_주문확정', '성공',
      '확정 ' + 확정목록.length + ' / 대기 ' + 대기목록.length + ' / 취소 ' + 취소목록.length);

    return {
      확정: 확정목록.length, 예약대기: 대기목록.length, 취소: 취소목록.length,
      확정목록: 확정목록, 대기목록: 대기목록, 취소목록: 취소목록
    };
  });
}

/**
 * S3_2. 확정 취소 — 확정된 주문을 되돌린다.
 *   예약재고를 풀어 가용재고로 되돌리고 주문상태를 접수로 되돌린다.
 *   이미 피킹지시가 나간 주문은 되돌릴 수 없다.
 */
function S3_2_확정취소(주문번호) {
  return withLock_(function () {
    if (!주문번호) {
      var ui = null;
      try { ui = SpreadsheetApp.getUi(); } catch (e) { }
      if (!ui) throw new Error('주문번호를 인자로 넘겨주세요. 예: S3_2_확정취소("20990101-0000001")');
      var resp = ui.prompt('확정 취소', '되돌릴 주문번호를 입력하세요', ui.ButtonSet.OK_CANCEL);
      if (resp.getSelectedButton() !== ui.Button.OK) return;
      주문번호 = resp.getResponseText().trim();
      if (!주문번호) return;
    }

    var 주문 = readTable_(ROLE.주문);
    var O = {
      주문번호: col_(주문, COL.주문번호, true),
      상품코드: col_(주문, COL.상품품목코드, true),
      수량: col_(주문, COL.수량, true),
      지시번호: col_(주문, COL.피킹지시번호, true),
      주문상태: col_(주문, COL.주문상태, true),
      확정일시: col_(주문, COL.확정일시, false)
    };

    var 대상 = [];
    주문.rows.forEach(function (r, i) {
      if (toStr_(r[O.주문번호]) === 주문번호) 대상.push({ rowIdx: i, row: r });
    });
    if (!대상.length) throw new Error('주문번호 "' + 주문번호 + '"를 찾을 수 없습니다.');

    if (toStr_(대상[0].row[O.주문상태]) !== ENUM.주문상태.확정) {
      throw new Error('확정 상태가 아닙니다. 현재: ' + toStr_(대상[0].row[O.주문상태]));
    }
    if (!isBlank_(대상[0].row[O.지시번호])) {
      throw new Error('이미 피킹지시(' + toStr_(대상[0].row[O.지시번호]) + ')가 발행되어 되돌릴 수 없습니다.');
    }

    var 마스터 = readTable_(ROLE.마스터);
    var M = {
      코드: col_(마스터, COL.상품품목코드, true),
      가용: col_(마스터, COL.가용재고, true),
      예약: col_(마스터, COL.예약재고, true)
    };
    var 행번호 = {}, 가용 = [], 예약 = [];
    마스터.rows.forEach(function (r, i) {
      var c = toStr_(r[M.코드]);
      if (c) 행번호[c] = i;
      가용[i] = toNum_(r[M.가용]);
      예약[i] = toNum_(r[M.예약]);
    });

    var 필요 = {};
    대상.forEach(function (t) {
      var c = toStr_(t.row[O.상품코드]);
      필요[c] = (필요[c] || 0) + toNum_(t.row[O.수량]);
    });

    var 로그 = [];
    var 사용자 = 사용자_();

    Object.keys(필요).forEach(function (code) {
      var mi = 행번호[code];
      if (mi === undefined) return;
      var 해제 = Math.min(필요[code], 예약[mi]);
      예약[mi] -= 해제;
      가용[mi] += 해제;
      로그.push({
        구분: ENUM.로그구분.예약해제, 피킹지시번호: '', 주문번호: 주문번호, 품목별주문번호: '',
        상품코드: code, 변동량: 해제, 변동후재고: 가용[mi],
        담당자: 사용자, 사유: '확정 취소 · 예약 해제'
      });
    });

    대상.forEach(function (t) {
      t.row[O.주문상태] = ENUM.주문상태.접수;
      if (O.확정일시 >= 0) t.row[O.확정일시] = '';
    });

    writeColumn_(주문.sheet, O.주문상태, 주문.rows);
    if (O.확정일시 >= 0) writeColumn_(주문.sheet, O.확정일시, 주문.rows);
    마스터.sheet.getRange(2, M.가용 + 1, 가용.length, 1).setValues(가용.map(function (x) { return [x]; }));
    마스터.sheet.getRange(2, M.예약 + 1, 예약.length, 1).setValues(예약.map(function (x) { return [x]; }));
    writeStockLog_(로그);

    var msg = '[' + 주문번호 + '] 확정을 취소했습니다.\n예약 재고를 가용재고로 되돌렸습니다.';
    alert_(msg);
    writeOpLog_('S3_2_확정취소', '성공', 주문번호);
    return msg;
  });
}

/**
 * S3_9. 확정 시뮬레이션 — 데이터를 바꾸지 않고 결과만 미리 본다.
 */
function S3_9_확정대기조회() {
  var 주문 = readTable_(ROLE.주문);
  var O = {
    주문번호: col_(주문, COL.주문번호, true),
    상품코드: col_(주문, COL.상품품목코드, true),
    수량: col_(주문, COL.수량, true),
    출고완료: col_(주문, COL.출고완료, true),
    지시번호: col_(주문, COL.피킹지시번호, true),
    주문상태: col_(주문, COL.주문상태, true)
  };

  var 마스터 = readTable_(ROLE.마스터);
  var M = {
    코드: col_(마스터, COL.상품품목코드, true),
    상품명: col_(마스터, COL.상품명, true),
    가용: col_(마스터, COL.가용재고, true),
    예약상품: col_(마스터, COL.예약상품, false)
  };

  var 잔여 = {}, 예약구분 = {}, 있음 = {};
  마스터.rows.forEach(function (r) {
    var c = toStr_(r[M.코드]);
    if (!c) return;
    있음[c] = true;
    잔여[c] = toNum_(r[M.가용]);
    예약구분[c] = (M.예약상품 >= 0)
      ? toStr_(r[M.예약상품]).toUpperCase() === 'Y'
      : isPreorderName_(toStr_(r[M.상품명]));
  });

  var 그룹 = {}, 순서 = [];
  주문.rows.forEach(function (r) {
    var no = toStr_(r[O.주문번호]);
    if (!no) return;
    var s = toStr_(r[O.주문상태]);
    if (s && s !== ENUM.주문상태.접수 && s !== ENUM.주문상태.예약대기) return;
    if (toNum_(r[O.출고완료]) === 1) return;
    if (!isBlank_(r[O.지시번호])) return;
    if (!그룹[no]) { 그룹[no] = {}; 순서.push(no); }
    var c = toStr_(r[O.상품코드]);
    그룹[no][c] = (그룹[no][c] || 0) + toNum_(r[O.수량]);
  });

  if (!순서.length) {
    alert_('확정 대기 중인 주문이 없습니다.');
    return '';
  }

  순서.sort();
  var out = ['확정 시뮬레이션 — 대상 ' + 순서.length + '건  (선착순, 데이터 변경 없음)', ''];
  var 통계 = { 확정: 0, 대기: 0, 취소: 0 };

  순서.forEach(function (no) {
    var 필요 = 그룹[no];
    var 문제 = [], 예약 = [];

    Object.keys(필요).forEach(function (c) {
      if (!있음[c]) { 문제.push(c + ' 미등록'); return; }
      if (예약구분[c]) { 예약.push(c); return; }
      if (잔여[c] < 필요[c]) 문제.push(c + ' 부족(' + 필요[c] + '/' + 잔여[c] + ')');
    });

    if (문제.length) {
      out.push('  ❌ 취소   ' + no + ' → ' + 문제.join(', '));
      통계.취소++;
    } else if (예약.length) {
      out.push('  ⏳ 대기   ' + no + ' → 예약상품 ' + 예약.join(', '));
      통계.대기++;
    } else {
      Object.keys(필요).forEach(function (c) { 잔여[c] -= 필요[c]; });
      out.push('  ✅ 확정   ' + no);
      통계.확정++;
    }
  });

  out.push('', '예상 — 확정 ' + 통계.확정 + ' / 예약대기 ' + 통계.대기 + ' / 취소 ' + 통계.취소);
  var msg = out.join('\n');
  Logger.log(msg);
  alert_(msg.substring(0, 1500));
  return msg;
}
