/**
 * ============================================================
 *  S1. 카페24 재고 동기화
 * ============================================================
 *  카페24에서 내려받은 재고 CSV를 상품마스터에 반영한다.
 *  통합 파이프라인에서는 헤더 검증이 끝난 단일 파일을 인자로 받는다.
 *  인자 없이 실행하는 기존 호환 모드만 파일명에서 '카페24'를 찾는다.
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
 *    재고수량      → 총 보유량     (가용 = 총보유 − 예약)
 *    재고관리 사용 → 재고관리      (T/F)
 *    상품명에 [설정]의 예약키워드 포함 → 예약상품 = Y
 *
 *  ★ 보관위치는 창고 고유 정보이므로 절대 덮어쓰지 않는다.
 * ============================================================
 */

var CAFE24 = {
  폴더파라미터: '재고CSV폴더ID',
  파일접두어: '카페24',
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
  return withLock_(function () {
    // ---------- CSV 찾기 ----------
    var 대상 = 입력파일 || null, 후보 = [];
    if (!대상) {
      var 폴더ID = String(param_(CAFE24.폴더파라미터, param_('CSV폴더ID', DEFAULT_FOLDER_ID)));
      var folder;
      try {
        folder = DriveApp.getFolderById(폴더ID);
      } catch (e) {
        throw new Error('재고 CSV 폴더를 열 수 없습니다. [설정] 탭의 ' + CAFE24.폴더파라미터 + ' 확인. (' + 폴더ID + ')');
      }

      // ---------- 파일 찾기 (CSV + 변환된 구글시트 모두) ----------
      var it = folder.getFiles();
      while (it.hasNext()) {
        var f = it.next();
        var n = f.getName();
        if (n.indexOf(CAFE24.파일접두어) < 0) continue;

        var mime = f.getMimeType();
        var csv파일 = /\.csv$/i.test(n) || mime.indexOf('csv') >= 0;
        var 시트파일 = mime === MimeType.GOOGLE_SHEETS;
        if (!csv파일 && !시트파일) continue;

        후보.push(n + (시트파일 ? ' [구글시트]' : ' [CSV]'));
        if (!대상 || f.getLastUpdated() > 대상.getLastUpdated()) 대상 = f;
      }
    }

    if (!대상) {
      throw new Error('폴더에서 "' + CAFE24.파일접두어 +
        '"가 들어간 CSV 또는 스프레드시트를 찾지 못했습니다.\n진단_재고폴더() 로 폴더 내용을 확인하세요.');
    }

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
      이미지: col_(마스터, COL.이미지, false),
      위치: col_(마스터, COL.기본보관위치, true),
      가용: col_(마스터, COL.가용재고, true),
      예약: col_(마스터, COL.예약재고, true),
      불량: col_(마스터, COL.불량재고, false),
      상품상태: col_(마스터, COL.상품상태, false),
      등록일: col_(마스터, '등록일', false),
      승인자: col_(마스터, '승인자', false),
      예약상품: col_(마스터, COL.예약상품, true),
      재고관리: col_(마스터, COL.재고관리, false),
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
    var 요약 = { 신규: 0, 갱신: 0, 재고변동: 0, 위치없음: [], 예약전환: 0, 경고: [] };
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
      var 판매가 = ci[CAFE24.열.판매가] !== undefined ? toNum_(row[ci[CAFE24.열.판매가]]) : '';
      var 부모코드 = ci[CAFE24.열.상품코드] !== undefined
        ? String(row[ci[CAFE24.열.상품코드]] || '').trim() : '';

      var 예약상품 = isPreorderName_(상품명) ? 'Y' : 'N';

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
        nr[M.예약] = 0;
        if (M.불량 >= 0) nr[M.불량] = 0;
        if (M.상품상태 >= 0) nr[M.상품상태] = '판매중';
        if (M.등록일 >= 0) nr[M.등록일] = now;
        if (M.승인자 >= 0) nr[M.승인자] = 사용자;
        nr[M.예약상품] = 예약상품;
        if (M.재고관리 >= 0) nr[M.재고관리] = 관리;
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
      if (M.재고관리 >= 0) mr[M.재고관리] = 관리;
      if (M.동기화 >= 0) mr[M.동기화] = now;

      var 이전예약상품 = toStr_(mr[M.예약상품]);
      mr[M.예약상품] = 예약상품;
      if (이전예약상품 && 이전예약상품 !== 예약상품) 요약.예약전환++;

      // 카페24 재고수량은 창고의 총 물리재고다.
      // S3에서 이미 확보한 예약재고를 빼야 새 주문에 사용할 수 있는 가용재고가 된다.
      var 예약중 = toNum_(mr[M.예약]);
      var 이전가용 = toNum_(mr[M.가용]);
      var 새가용 = 재고 - 예약중;

      if (새가용 < 0) {
        요약.경고.push(code + ': 카페24 재고 ' + 재고 + ' < 예약 ' + 예약중 + ' → 가용 0');
        새가용 = 0;
      }

      if (새가용 !== 이전가용) {
        mr[M.가용] = 새가용;
        로그.push({
          구분: ENUM.로그구분.동기화, 피킹지시번호: '', 주문번호: '', 품목별주문번호: '',
          상품코드: code, 변동량: 새가용 - 이전가용, 변동후재고: 새가용,
          담당자: 사용자, 사유: '카페24 동기화 (총 ' + 재고 + ' − 예약 ' + 예약중 + ')'
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
    [M.상품명, M.옵션명, M.가용, M.예약상품, M.관리코드, M.판매가, M.재고관리, M.동기화]
      .forEach(function (idx) {
        if (idx >= 0) writeColumn_(마스터.sheet, idx, 마스터.rows);
      });

    if (신규행.length) {
      appendRows_(마스터.sheet, 신규행, [M.코드, M.내부SKU, M.관리코드]);
    }

    writeStockLog_(로그);

    // ---------- 보고 ----------
    var msg = '카페24 재고 동기화 완료\n파일: ' + 대상.getName() + '\n\n' +
      '신규 등록 ' + 요약.신규 + '건 / 갱신 ' + 요약.갱신 + '건 / 재고 변동 ' + 요약.재고변동 + '건';

    if (요약.예약전환) msg += '\n예약상품 구분 변경 ' + 요약.예약전환 + '건';
    if (잔여.length) msg += '\n\n카페24에 없는 마스터 코드 ' + 잔여.length + '건 (그대로 유지)';

    if (요약.위치없음.length) {
      msg += '\n\n⚠ 보관위치가 비어 있는 상품 ' + 요약.위치없음.length + '건\n' +
        '   피킹 시 위치가 표시되지 않습니다. 상품마스터에서 채워주세요.\n   ' +
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
    가용: col_(마스터, COL.가용재고, true),
    예약상품: col_(마스터, COL.예약상품, false)
  };

  var out = [];
  마스터.rows.forEach(function (r, i) {
    var code = toStr_(r[M.코드]);
    if (!code) return;
    if (!isBlank_(r[M.위치])) return;
    out.push('  행' + (i + 2) + '  ' + code +
      '  재고' + toNum_(r[M.가용]) +
      (M.예약상품 >= 0 && toStr_(r[M.예약상품]) === 'Y' ? '  [예약]' : '') +
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
