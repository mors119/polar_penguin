/**
 * ============================================================
 *  S1. 카페24 재고 동기화
 * ============================================================
 *  카페24에서 내려받은 재고 CSV를 상품마스터에 반영한다.
 *  통합 파이프라인에서는 헤더 검증이 끝난 단일 파일을 인자로 받는다.
 *  인자 없이 실행하면 통합 Input 파이프라인으로 연결한다.
 *  신규 상품은 등록하고, 기존 상품은 최신 정보로 갱신한다.
 *
 *  카페24 CSV 열
 *    상품코드 | 자체 상품코드 | 상품명 | 판매가 | 총 재고량 |
 *    품목코드 | 품목명 | 자체 품목코드 | 재고관리 사용 | 재고수량
 *
 *  매핑
 *    품목코드      → 상품품목코드  (PK)
 *    상품명        → 상품명
 *    품목명        → 옵션명        (비면 '-')
 *    재고수량      → 물리 재고 snapshot
 *    재고관리 사용 → 재고관리      (T/F)
 *
 *  ★ 보관위치는 창고 고유 정보이므로 절대 덮어쓰지 않는다.
 * ============================================================
 */

var CAFE24 = {
  열: {
    상품코드: '상품코드',
    상품명: '상품명',
    판매가: '판매가',
    총재고량: '총 재고량',
    품목코드: '품목코드',
    품목명: '품목명',
    재고관리: '재고관리 사용',
    재고수량: '재고수량'
  }
};

/**
 * S1_1. 카페24 재고 CSV/Google Spreadsheet → 상품마스터 동기화
 * @param {GoogleAppsScript.Drive.File=} 입력파일 processInput이 판별한 단일 파일
 * @param {Boolean=} silent 수동 알림 생략 여부
 * @return {Object} 신규/갱신/재고변동 건수와 위치 미지정·경고 목록
 * @sideEffect 상품마스터와 재고이동로그를 갱신함. 기본보관위치는 절대 덮어쓰지 않음
 */
function S1_1_카페24재고동기화(입력파일, silent) {
  // 운영자가 이 내부 단계를 직접 실행해도 헤더 판별을 포함한 표준 자동 흐름을 사용한다.
  if (!입력파일) return processInput();
  return withLock_(function () {
    var 대상 = 입력파일;

    // ---------- 파싱 (형식에 따라 분기) ----------
    var parsed;
    if (대상.getMimeType() === MimeType.GOOGLE_SHEETS) {
      var sh = SpreadsheetApp.openById(대상.getId()).getSheets()[0];
      parsed = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues()
        .map(function (row) { return row.map(function (c) { return toStr_(c); }); });
    } else {
      parsed = Utilities.parseCsv(readCsvText_(대상));
    }
    if (!parsed || parsed.length < 2) throw new Error('데이터가 없습니다.');

    // ---------- 헤더 매핑 ----------
    var ch = parsed[0].map(function (h) {
      return String(h || '').replace(/^\uFEFF/, '').trim();
    });
    var norm = {};
    ch.forEach(function (h, i) {
      var k = h.replace(/[\s\u00A0]/g, '');
      if (k && norm[k] === undefined) norm[k] = i;
    });

    var ci = {};
    Object.keys(CAFE24.열).forEach(function (key) {
      var 원본명 = CAFE24.열[key];
      ci[원본명] = norm[원본명.replace(/[\s\u00A0]/g, '')];
    });

    var 필수 = [CAFE24.열.품목코드, CAFE24.열.상품명, CAFE24.열.재고수량];
    var 없는열 = 필수.filter(function (n) { return ci[n] === undefined; });
    if (없는열.length) {
      throw new Error('필요한 열이 없습니다: ' + 없는열.join(', ') +
        '\n\n실제 열 ' + ch.length + '개:\n' + ch.join(' | '));
    }


    // ---------- 마스터 적재 ----------
    var 마스터 = readTable_(ROLE.마스터);
    var M = {
      코드: col_(마스터, COL.상품품목코드, true),
      내부SKU: col_(마스터, '내부SKU', false),
      관리코드: col_(마스터, '관리코드', false),
      상품명: col_(마스터, COL.상품명, true),
      옵션명: col_(마스터, COL.옵션명, false),
      상품구분: col_(마스터, '상품구분', false),
      위치: col_(마스터, COL.기본보관위치, true),
      가용: col_(마스터, COL.가용재고, true),
      등록일: col_(마스터, '등록일', false),
      승인자: col_(마스터, '승인자', false),
      재고관리: col_(마스터, COL.재고관리, true),
      판매가: col_(마스터, COL.판매가, false),
      동기화: col_(마스터, COL.최종동기화, false)
    };

    var 행번호 = {};
    마스터.rows.forEach(function (r, i) {
      var c = toStr_(r[M.코드]);
      if (c) 행번호[c] = i;
    });

    // ---------- 처리 ----------
    var now = new Date();
    var 사용자 = 사용자_();
    var 신규행 = [], 로그 = [];
    var 요약 = { 신규: 0, 갱신: 0, 재고변동: 0, 위치없음: [], 경고: [] };
    var 확정수요 = committedDemandBySku_();
    var CSV코드 = {};

    for (var r = 1; r < parsed.length; r++) {
      var row = parsed[r];
      if (!row.join('').trim()) continue;

      var code = String(row[ci[CAFE24.열.품목코드]] || '').trim();
      if (!code) continue;
      CSV코드[code] = true;

      var 상품명 = String(row[ci[CAFE24.열.상품명]] || '').trim();
      var 품목명 = ci[CAFE24.열.품목명] !== undefined
        ? String(row[ci[CAFE24.열.품목명]] || '').trim() : '';
      var 재고 = toNum_(row[ci[CAFE24.열.재고수량]]);
      var 관리 = ci[CAFE24.열.재고관리] !== undefined
        ? String(row[ci[CAFE24.열.재고관리]] || '').trim().toUpperCase() : 'T';
      if (!관리) 관리 = 'T';
      var 판매가 = ci[CAFE24.열.판매가] !== undefined ? toNum_(row[ci[CAFE24.열.판매가]]) : '';
      var 부모코드 = ci[CAFE24.열.상품코드] !== undefined
        ? String(row[ci[CAFE24.열.상품코드]] || '').trim() : '';

      // ---------- 신규 등록 ----------
      if (행번호[code] === undefined) {
        var nr = new Array(마스터.width).fill('');
        nr[M.코드] = code;
        if (M.내부SKU >= 0) nr[M.내부SKU] = code;
        if (M.관리코드 >= 0) nr[M.관리코드] = 부모코드;
        nr[M.상품명] = 상품명;
        if (M.옵션명 >= 0) nr[M.옵션명] = 품목명 || '-';
        nr[M.위치] = '';                     // ★ 관리자 입력 필요
        nr[M.가용] = 재고;
        if (M.등록일 >= 0) nr[M.등록일] = now;
        if (M.승인자 >= 0) nr[M.승인자] = 사용자;
        nr[M.재고관리] = 관리;
        if (M.판매가 >= 0) nr[M.판매가] = 판매가;
        if (M.동기화 >= 0) nr[M.동기화] = now;
        신규행.push(nr);

        요약.신규++;
        요약.위치없음.push(code);
        if (재고 > 0) {
          로그.push({
            구분: ENUM.로그구분.동기화, 피킹지시번호: '', 주문번호: '', 품목별주문번호: '',
            상품코드: code, 변동량: 재고, 변동후재고: 재고,
            담당자: 사용자, 사유: '카페24 신규 등록'
          });
        }
        행번호[code] = -1;
        continue;
      }

      if (행번호[code] < 0) continue;

      // ---------- 기존 갱신 ----------
      var mi = 행번호[code];
      var mr = 마스터.rows[mi];

      mr[M.상품명] = 상품명;
      if (M.옵션명 >= 0) mr[M.옵션명] = 품목명 || '-';
      if (M.관리코드 >= 0 && 부모코드) mr[M.관리코드] = 부모코드;
      if (M.판매가 >= 0) mr[M.판매가] = 판매가;
      mr[M.재고관리] = 관리;
      if (M.동기화 >= 0) mr[M.동기화] = now;

      // 카페24 재고수량은 가산 delta가 아닌 물리 재고 snapshot이다.
      // 순 재고 = snapshot - 확정되었지만 아직 출고완료가 아닌 주문 수요.
      // 같은 snapshot을 다시 처리해도 같은 값을 계산하며, F는 결과가 음수여도 그대로 보존한다.
      var 미완료수요 = 확정수요[code] || 0;
      var 이전가용 = toNum_(mr[M.가용]);
      var 새가용 = netStockFromSnapshot_(재고, 미완료수요, 관리);

      if (관리 !== 'F' && 재고 - 미완료수요 < 0) {
        요약.경고.push(code + ': 카페24 재고 ' + 재고 + ' < 확정 수요 ' + 미완료수요 + ' → 가용 0');
      }

      if (새가용 !== 이전가용) {
        mr[M.가용] = 새가용;
        로그.push({
          구분: ENUM.로그구분.동기화, 피킹지시번호: '', 주문번호: '', 품목별주문번호: '',
          상품코드: code, 변동량: 새가용 - 이전가용, 변동후재고: 새가용,
          담당자: 사용자, 사유: '카페24 snapshot ' + 재고 + ' − 미완료 확정 수요 ' + 미완료수요
        });
        요약.재고변동++;
      }

      요약.갱신++;
      if (isBlank_(mr[M.위치])) 요약.위치없음.push(code);
    }

    // ---------- 마스터에만 있는 코드 ----------
    var 잔여 = [];
    마스터.rows.forEach(function (mr2) {
      var c = toStr_(mr2[M.코드]);
      if (c && !CSV코드[c]) 잔여.push(c);
    });

    // ---------- 일괄 쓰기 ----------
    [M.상품명, M.옵션명, M.가용, M.관리코드, M.판매가, M.재고관리, M.동기화]
      .forEach(function (idx) {
        if (idx >= 0) writeColumn_(마스터.sheet, idx, 마스터.rows);
      });

    if (신규행.length) {
      appendRows_(마스터.sheet, 신규행, [M.코드, M.내부SKU, M.관리코드]);
    }

    writeStockLog_(로그);

    // 카페24 값은 권위 있는 물리 snapshot이다. 수동 입고 release에서 남긴 carry가
    // 더 크더라도 snapshot 이후까지 살아 이중 배정되지 않도록 현재 값으로 재조정한다.
    if (typeof synchronizeReservationPhysicalCarryFromSnapshot_ === 'function') {
      synchronizeReservationPhysicalCarryFromSnapshot_(CSV코드);
    }

    // ---------- 보고 ----------
    var msg = '카페24 재고 동기화 완료\n파일: ' + 대상.getName() + '\n\n' +
      '신규 등록 ' + 요약.신규 + '건 / 갱신 ' + 요약.갱신 + '건 / 재고 변동 ' + 요약.재고변동 + '건';

    if (잔여.length) msg += '\n\n카페24에 없는 마스터 코드 ' + 잔여.length + '건 (그대로 유지)';

    if (요약.위치없음.length) {
      msg += '\n\n⚠ 보관위치가 비어 있는 상품 ' + 요약.위치없음.length + '건\n' +
        '   📍 위치 관리 → 위치 미지정 상품에서 입력해주세요.\n   ' +
        요약.위치없음.slice(0, 10).join(', ') +
        (요약.위치없음.length > 10 ? ' 외 ' + (요약.위치없음.length - 10) + '건' : '');
    }
    if (요약.경고.length) {
      msg += '\n\n⚠ 재고 경고\n   ' + 요약.경고.slice(0, 5).join('\n   ');
    }

    if (!silent) alert_(msg);
    writeOpLog_('S1_1_카페24재고동기화', '성공',
      '신규 ' + 요약.신규 + ' / 갱신 ' + 요약.갱신 + ' / 재고변동 ' + 요약.재고변동);
    return 요약;
  });
}

/** 카페24 snapshot을 누적하지 않고 현재 순 재고로 재계산한다. */
function netStockFromSnapshot_(physicalSnapshot, committedDemand, stockManagement) {
  var net = toNum_(physicalSnapshot) - toNum_(committedDemand);
  return toStr_(stockManagement).toUpperCase() === 'F' ? net : Math.max(0, net);
}

/** 재고에 이미 반영되었고 아직 출고완료/취소되지 않은 주문 수요를 SKU별로 집계한다. */
function committedDemandBySku_() {
  var 주문 = readTable_(ROLE.주문);
  var O = {
    상품코드: col_(주문, COL.상품품목코드, true), 수량: col_(주문, COL.수량, true),
    상태: col_(주문, COL.주문상태, true), 확정일시: col_(주문, COL.확정일시, false)
  };
  var demand = {};
  주문.rows.forEach(function (row) {
    var state = toStr_(row[O.상태]);
    if (O.확정일시 < 0 || isBlank_(row[O.확정일시])) return;
    if (state !== ENUM.주문상태.예약 && state !== '처리완료') return;
    var code = toStr_(row[O.상품코드]);
    if (code) demand[code] = (demand[code] || 0) + toNum_(row[O.수량]);
  });
  return demand;
}

/**
 * S1_2. 보관위치가 비어 있는 상품 목록
 */
function S1_2_보관위치미지정조회() {
  var 마스터 = readTable_(ROLE.마스터);
  var M = {
    코드: col_(마스터, COL.상품품목코드, true),
    상품명: col_(마스터, COL.상품명, true),
    옵션: col_(마스터, COL.옵션명, false),
    위치: col_(마스터, COL.기본보관위치, true),
    가용: col_(마스터, COL.가용재고, true)
  };

  var out = [];
  마스터.rows.forEach(function (r, i) {
    var code = toStr_(r[M.코드]);
    if (!code) return;
    if (!isBlank_(r[M.위치])) return;
    out.push('  행' + (i + 2) + '  ' + code +
      '  재고' + toNum_(r[M.가용]) +
      '  ' + toStr_(r[M.상품명]).substring(0, 30) +
      (M.옵션 >= 0 && toStr_(r[M.옵션]) !== '-' ? ' / ' + toStr_(r[M.옵션]) : ''));
  });

  var msg = out.length
    ? '보관위치 미지정 ' + out.length + '건\n\n' + out.slice(0, 40).join('\n') +
      (out.length > 40 ? '\n... 외 ' + (out.length - 40) + '건' : '')
    : '모든 상품에 보관위치가 지정되어 있습니다.';

  Logger.log(msg);
  alert_(msg.substring(0, 1500));
  return msg;
}

/**
 * CSV 텍스트를 읽는다. UTF-8로 읽어 깨지면 CP949(EUC-KR)로 재시도한다.
 * S2 주문취입에서도 함께 쓴다.
 */
function readCsvText_(file) {
  var blob = file.getBlob();
  var text;

  try { text = blob.getDataAsString('UTF-8'); } catch (e) { text = null; }

  if (!text || text.indexOf('\uFFFD') >= 0) {
    try { text = blob.getDataAsString('EUC-KR'); }
    catch (e2) {
      if (!text) throw new Error('"' + file.getName() + '" 인코딩을 해석하지 못했습니다.');
    }
  }

  return text.replace(/^\uFEFF/, '');
}
