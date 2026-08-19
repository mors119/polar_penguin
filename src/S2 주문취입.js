
/**
 * ============================================================
 *  S2. 주문 CSV 취입
 * ============================================================
 *  주문 CSV/Google Spreadsheet를 읽어 주문(완료) 시트에 적재한다.
 *  통합 파이프라인은 헤더 검증이 끝난 단일 파일을 넘기고 원본 이동을 직접 관리한다.
 *  인자 없이 실행하면 통합 Input 파이프라인으로 연결한다.
 *
 *   · 품목별 주문번호 기준으로 중복을 제거한다
 *     (같은 파일을 두 번 올려도 중복 적재되지 않는다)
 *   · 주문번호·우편번호 등은 문자열로 강제해 앞자리 0을 보존한다
 *   · 신규 행은 시트 최상단에 삽입한다 (최신이 위로)
 *   · 초기 상태: 출고완료 0 · 피킹지시번호 공란 · 주문상태 예약
 *     (같은 실행 안에서 재고 확보와 PDF가 끝나면 출고완료로 전환)
 * ============================================================
 */
 
/** 문자열로 보존해야 하는 컬럼 */
var TEXT_COLUMNS_ORDER = [
  COL.주문번호,
  COL.품목별주문번호,
  COL.상품품목코드,
  '수령인 우편번호',
  '수령인 휴대전화'
];
 
/**
 * @param {GoogleAppsScript.Drive.File=} 입력파일 processInput이 판별한 단일 파일
 * @param {{silent:Boolean=}=} options 파이프라인의 알림 제어
 * @return {Object} 처리한 파일 수와 신규/중복/오류 행 수
 * @sideEffect 신규 주문을 예약 상태로 적재. 원본 이동은 통합 Input 파이프라인이 담당
 */
function S2_1_주문CSV취입(입력파일, options) {
  // 운영자가 이 내부 단계를 직접 실행해도 별도 폴더를 만들지 않고 표준 자동 흐름을 사용한다.
  if (!입력파일) return processInput();
  return withLock_(function () {
    options = options || {};
    var csvFiles = [입력파일];

    // 카페24 주문 export의 모든 열을 먼저 보충한다. 내부 운영 열만 남기고
    // 원본 필드를 버리는 회귀를 막으며, 기존 열과 데이터는 이동하거나 지우지 않는다.
    ensureOrderImportColumns_(csvFiles);
 
    // ---------- 주문 시트 준비 ----------
    var 주문 = readTable_(ROLE.주문);
    var c품목별 = col_(주문, COL.품목별주문번호, true);
    var c출고완료 = col_(주문, COL.출고완료, true);
    var c지시번호 = col_(주문, COL.피킹지시번호, true);
    var c주문상태 = col_(주문, COL.주문상태, true);
 
    var 기존키 = {};
    주문.rows.forEach(function (r) {
      var k = toStr_(r[c품목별]);
      if (k) 기존키[k] = true;
    });
 
    var 텍스트열 = TEXT_COLUMNS_ORDER
      .map(function (n) { return col_(주문, n, false); })
      .filter(function (i) { return i >= 0; });
 
    var 리포트 = [], 신규행 = [], 주문번호집합 = {};
    var 총신규 = 0, 총중복 = 0, 총오류 = 0;
 
    csvFiles.forEach(function (file) {
      var r = parseCsvFile_(file, 주문, c품목별, 기존키);
      신규행 = 신규행.concat(r.rows);
      (r.주문번호 || []).forEach(function (no) { 주문번호집합[no] = true; });
      총신규 += r.rows.length;
      총중복 += r.중복;
      총오류 += r.오류;
      리포트.push('· ' + file.getName() + ' — 총 ' + r.총행 + '행 / 신규 ' + r.rows.length +
                  ' / 중복 ' + r.중복 + (r.오류 ? ' / 오류 ' + r.오류 : ''));
    });
 
    // ---------- 초기값 세팅 후 최상단에 삽입 ----------
    if (신규행.length) {
      신규행.forEach(function (row) {
        if (isBlank_(row[c출고완료])) row[c출고완료] = 0;
        row[c지시번호] = '';
        row[c주문상태] = ENUM.주문상태.예약;
      });
      // 최신 주문이 위로 오도록 역순 정렬 후 삽입
      신규행.reverse();
      prependRows_(주문.sheet, 신규행, 텍스트열);
    }
 
    var msg = 'CSV 취입 완료\n\n' + 리포트.join('\n') +
      '\n\n합계 — 신규 ' + 총신규 + '건 / 중복 제외 ' + 총중복 + '건' +
      (총오류 ? ' / 오류 ' + 총오류 + '건' : '') +
      '\n\n통합 Input 파이프라인이 이어서 재고 검증·확정·피킹 생성을 자동 수행합니다.';
 
    if (!options.silent) alert_(msg);
    writeOpLog_('S2_1_주문CSV취입', '성공', msg.replace(/\n/g, ' | '));
    return { 파일수: csvFiles.length, 신규: 총신규, 중복: 총중복, 오류: 총오류,
      주문번호: Object.keys(주문번호집합) };
  });
}
 
/* ============================================================
 *  내부 헬퍼
 * ============================================================ */
 
function parseCsvFile_(file, 주문, c품목별, 기존키) {
  var parsed;
  try {
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      var sh = SpreadsheetApp.openById(file.getId()).getSheets()[0];
      parsed = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues()
        .map(function (row) { return row.map(function (c) { return toStr_(c); }); });
    } else {
      parsed = Utilities.parseCsv(readCsvText_(file));
    }
  } catch (e) {
    throw new Error('"' + file.getName() + '" 읽기 실패: ' + e.message);
  }
 
  if (!parsed || parsed.length < 2) return { rows: [], 중복: 0, 오류: 0, 총행: 0 };
 
  // CSV 헤더 → 주문 시트 열 인덱스 매핑
  var csvHeader = parsed[0];
  var 매핑 = [];
  for (var c = 0; c < csvHeader.length; c++) {
    var name = String(csvHeader[c] || '').replace(/^\uFEFF/, '').trim();
    if (!name) { 매핑.push(-1); continue; }
    매핑.push(col_(주문, name, false));
  }
 
  var rows = [], 중복 = 0, 오류 = 0, 주문번호 = [];
  var c주문번호 = col_(주문, COL.주문번호, true);
 
  for (var i = 1; i < parsed.length; i++) {
    var src = parsed[i];
    if (!src.join('').trim()) continue;
 
    var row = new Array(주문.width).fill('');
    for (var j = 0; j < src.length && j < 매핑.length; j++) {
      if (매핑[j] >= 0) row[매핑[j]] = src[j];
    }
 
    var key = toStr_(row[c품목별]);
    var no = toStr_(row[c주문번호]);
    if (no && 주문번호.indexOf(no) < 0) 주문번호.push(no);
    if (!key) { 오류++; continue; }
    if (기존키[key]) { 중복++; continue; }
 
    기존키[key] = true;
    rows.push(row);
  }
 
  return { rows: rows, 중복: 중복, 오류: 오류, 총행: parsed.length - 1, 주문번호: 주문번호 };
}
 
function getOrCreateSubFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function ensureOrderImportColumns_(files) {
  var headers = [];
  (files || []).forEach(function (file) {
    var first;
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      var sh = SpreadsheetApp.openById(file.getId()).getSheets()[0];
      if (sh.getLastColumn() > 0) first = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    } else {
      var parsed = Utilities.parseCsv(readCsvText_(file));
      first = parsed && parsed[0];
    }
    (first || []).forEach(function (value) {
      var name = String(value || '').replace(/^\uFEFF/, '').trim();
      if (name && headers.indexOf(name) < 0) headers.push(name);
    });
  });
  if (!headers.length) return [];
  var order = readTable_(ROLE.주문);
  return ensureColumns_(order.sheet, order.headers, headers.filter(function (name) {
    return col_(order, name, false) < 0;
  }));
}
 
