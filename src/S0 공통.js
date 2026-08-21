

/**
 * ============================================================
 *  S0. 공통 정의와 런타임 유틸리티
 * ============================================================
 *  시트/열/상태의 표준 정의, 단일 운영 Spreadsheet 연결,
 *  테이블 IO, 재고/작업 로깅, LockService, UI 알림을 한곳에서 제공한다.
 *  다른 모듈은 시트명과 컬럼명을 직접 중복하지 않고 이 정의를 사용한다.
 * ============================================================
 */

/** setup 이전 환경을 위한 기존 fallback. 정상 설치는 OPERATION_SPREADSHEET_ID를 사용한다. */
var CONSOLE_SS_ID = '';

var CONSOLE = {
  설정: '설정',
  재고이동로그: '재고이동로그',
  작업로그: '작업로그',
  입력처리로그: '입력처리로그',
  처리주문아카이브: '처리주문아카이브'
};

var ORDER_ARCHIVE_HEADERS = [
  '아카이브일시', '품목별 주문번호', '주문번호', '상품품목코드',
  '최종상태', '최종처리일시', '피킹지시번호', '주문수량'
];

var INPUT_LOG_HEADERS = ['처리시각', '파일ID', '파일명', '체크섬', '유형', '상태', '오류코드', '메시지'];

var ROLE = {
  마스터: '상품마스터',
  주문: '주문(완료)',
  헤더: '피킹(헤더)',
  라인: '피킹(라인)'
};

/** 표준 컬럼명 */
var COL = {
  // ---------- 상품마스터 ----------
  상품품목코드: '상품품목코드',
  상품명: '상품명',
  옵션명: '옵션명',
  기본보관위치: '기본보관위치',
  가용재고: '가용재고',
  재고관리: '재고관리',
  판매가: '판매가',
  최종동기화: '최종동기화',

  // ---------- 주문 (완료) ----------
  쇼핑몰: '쇼핑몰',
  쇼핑몰번호: '쇼핑몰번호',
  주문번호: '주문번호',
  품목별주문번호: '품목별 주문번호',
  배송메시지: '배송메시지',
  총주문금액: '총 주문금액(KRW)',
  총결제금액: '총 결제금액(KRW)',
  주문상품명: '주문상품명(기본)',
  상품옵션기본: '상품옵션(기본)',
  수량: '수량',
  수령인: '수령인',
  수령인휴대전화: '수령인 휴대전화',
  수령인우편번호: '수령인 우편번호',
  수령인주소전체: '수령인 주소(전체)',
  출고완료: '출고완료',
  피킹지시번호: '피킹지시번호',
  주문상태: '주문상태',
  취소사유: '취소사유',
  취소일시: '취소일시',
  취소경로: '취소경로',
  확정일시: '확정일시',
  대기사유: '대기사유',
  운영메모: '운영메모',

  // ---------- 피킹 (헤더) ----------
  카트슬롯: '카트 슬롯',
  품목수: '품목수',
  총수량: '총수량',
  피킹담당자: '피킹담당자',
  상태: '상태',
  생성일시: '생성일시',
  출력일시: '출력일시',

  // ---------- 피킹 (라인) ----------
  순번: '순번',
  보관위치: '보관위치',
  상품코드: '상품코드',
  이미지: '이미지', // 피킹(라인) 레거시 감사 필드
  옵션: '옵션',
  필요수량: '필요수량',
  확인: '확인',
  실제수량: '실제수량',
  예외사유: '예외사유',
  담당자: '담당자',
  라인상태: '라인상태',
  처리일시: '처리일시'
};

/** 주문(완료)의 앞 25개 고정 열. CSV에서 관찰된 추가 열은 이 목록 뒤에만 붙는다. */
var ORDER_FIXED_HEADERS = [
  COL.쇼핑몰, COL.쇼핑몰번호, COL.주문번호, COL.품목별주문번호,
  COL.배송메시지, COL.총주문금액, COL.총결제금액, COL.상품품목코드,
  COL.주문상품명, COL.상품옵션기본, COL.수량, COL.판매가,
  COL.수령인, COL.수령인휴대전화, COL.수령인우편번호, COL.수령인주소전체,
  COL.출고완료, COL.피킹지시번호, COL.주문상태, COL.취소사유,
  COL.취소일시, COL.취소경로, COL.확정일시, COL.대기사유, COL.운영메모
];

/** CSV/레거시 주문 열을 중복 의미 열 없이 고정 열로 흡수하기 위한 별칭. */
var ORDER_FIXED_HEADER_ALIASES = {};
ORDER_FIXED_HEADER_ALIASES[COL.품목별주문번호] = ['품목별주문번호', '주문상세번호'];
ORDER_FIXED_HEADER_ALIASES[COL.상품품목코드] = ['품목코드', '상품코드', 'SKU'];
ORDER_FIXED_HEADER_ALIASES[COL.주문상품명] = ['상품명'];
ORDER_FIXED_HEADER_ALIASES[COL.상품옵션기본] = ['옵션명', '옵션'];
ORDER_FIXED_HEADER_ALIASES[COL.수량] = ['주문수량'];
ORDER_FIXED_HEADER_ALIASES[COL.수령인주소전체] = ['수령인 주소', '수령인주소'];

/** 이전 버전에서 고정 생성했지만 이제 실제 값이 있을 때만 동적으로 보존하는 열. */
var ORDER_LEGACY_OPTIONAL_HEADERS = [
  '주문일시', '주문경로', '결제수단', '상품번호', '상품구매금액', '할인금액',
  '실결제금액', '주문자명', '주문자 이메일', '주문자 휴대전화', '수령인 일반전화',
  '배송업체', '송장번호', '배송비', '배송유형'
];

function canonicalOrderHeader_(header) {
  var key = normKey_(header);
  if (!key) return '';
  for (var i = 0; i < ORDER_FIXED_HEADERS.length; i++) {
    var canonical = ORDER_FIXED_HEADERS[i];
    if (normKey_(canonical) === key) return canonical;
    var aliases = ORDER_FIXED_HEADER_ALIASES[canonical] || [];
    for (var j = 0; j < aliases.length; j++) {
      if (normKey_(aliases[j]) === key) return canonical;
    }
  }
  return '';
}

var ENUM = {
  헤더상태: { 대기: '대기', 완료: '완료', 취소: '취소', 출력오류: '출력오류' },
  라인상태: { 미처리: '미처리', 완료: '완료', 취소: '취소' },
  주문상태: { 예약: '예약', 출고완료: '출고완료', 취소: '취소' },
  확인: { 정상: 'O', 예외: 'X' },
  예외사유: ['재고없음'],
  로그구분: { 복원: '복원', 확정: '확정', 동기화: '동기화' }
};

var _cache = { config: null, ss: {}, consoleSS: null };
var _scriptLockDepth = 0;

/* ============================================================
 *  단일 운영 스프레드시트 접근
 * ============================================================ */

function consoleSS_() {
  if (_cache.consoleSS) return _cache.consoleSS;

  var ss = null;
  var props = PropertiesService.getScriptProperties();
  var propertyId = props.getProperty('OPERATION_SPREADSHEET_ID') || props.getProperty('CONSOLE_SS_ID');
  if (propertyId) {
    try { ss = SpreadsheetApp.openById(propertyId); }
    catch (e) { throw new Error('Script Property의 운영 Spreadsheet를 열 수 없습니다. (' + propertyId + ')'); }
  }

  if (!ss) {
    try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e2) { ss = null; }
  }

  if (!ss) {
    var id = CONSOLE_SS_ID;
    if (!id) {
      throw new Error('설정·로그를 둘 스프레드시트가 지정되지 않았습니다.\n' +
        'CONSOLE_SS_ID 를 입력한 뒤 설정_콘솔파일지정() 을 실행하세요.');
    }
    try { ss = SpreadsheetApp.openById(id); }
    catch (e2) { throw new Error('지정된 스프레드시트를 열 수 없습니다. (' + id + ')'); }
  }

  _cache.consoleSS = ss;
  return ss;
}

/** 레거시 수동 설치 호환용. 새 설치는 setRootFolder()와 setupSystem()을 사용한다. */
function 설정_콘솔파일지정() {
  var 입력 = PropertiesService.getScriptProperties().getProperty('CONSOLE_SS_ID') || CONSOLE_SS_ID;
  if (!입력) throw new Error('CONSOLE_SS_ID 에 스프레드시트 ID나 URL을 입력하세요.');

  var m = String(입력).match(/[-\w]{25,}/);
  if (!m) throw new Error('ID를 찾지 못했습니다: ' + 입력);

  var ss = SpreadsheetApp.openById(m[0]);
  PropertiesService.getScriptProperties().setProperty('CONSOLE_SS_ID', m[0]);
  _cache.consoleSS = null;

  var msg = '운영 Spreadsheet 지정 완료: ' + ss.getName();
  Logger.log(msg);
  return msg;
}

/* ============================================================
 *  설정
 * ============================================================ */

function getConfig_() {
  if (_cache.config) return _cache.config;

  var sh = consoleSS_().getSheetByName(CONSOLE.설정);
  if (!sh) throw new Error('[' + CONSOLE.설정 + '] 탭이 없습니다. 설치_1_초기설정() 을 실행하세요.');

  var values = sh.getDataRange().getValues();
  var cfg = { 파일ID: {}, 시트명: {}, 파라미터: {}, 별칭: {} };

  for (var i = 1; i < values.length; i++) {
    var 구분 = String(values[i][0] || '').trim();
    var 키 = String(values[i][1] || '').trim();
    var 값 = values[i][2];
    if (!구분 || !키) continue;

    if (구분 === '별칭') {
      cfg.별칭[키] = String(값 || '').split(',').map(function (s) { return s.trim(); }).filter(String);
    } else if (cfg[구분]) {
      cfg[구분][키] = (typeof 값 === 'string') ? 값.trim() : 값;
    }
  }

  _cache.config = cfg;
  return cfg;
}

/** 같은 실행 안에서 setup 등으로 설정 행이 바뀐 경우 다음 조회가 시트를 다시 읽게 한다. */
function resetConfigCache_() {
  _cache.config = null;
}

function param_(키, 기본값) {
  var v = getConfig_().파라미터[키];
  return (v === undefined || v === null || v === '') ? 기본값 : v;
}

/* ============================================================
 *  역할별 시트 접근
 * ============================================================ */

function openSS_(role) {
  if (_cache.ss.operation) return _cache.ss.operation;
  var ss = consoleSS_();
  _cache.ss.operation = ss;
  return ss;
}

function getSheet_(role) {
  var ss = openSS_(role);
  var name = getConfig_().시트명[role] || role;
  var sh = name ? ss.getSheetByName(name) : null;
  if (!sh) throw new Error('운영 Spreadsheet에 "' + name + '" 탭이 없습니다. setupSystem()으로 복구하세요.');
  return sh;
}

/* ============================================================
 *  테이블 IO
 * ============================================================ */

function normKey_(s) {
  return String(s || '').replace(/[\s\u00A0]/g, '').toLowerCase();
}

function readTable_(role) {
  var sh = getSheet_(role);
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 1 || lastCol < 1) throw new Error('"' + role + '" 시트가 비어 있습니다.');

  var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0];
  var rows = values.slice(1);

  var headerIndex = {};
  for (var c = 0; c < headers.length; c++) {
    var k = normKey_(headers[c]);
    if (k && headerIndex[k] === undefined) headerIndex[k] = c;
  }

  return { sheet: sh, headers: headers, rows: rows, headerIndex: headerIndex, width: headers.length, role: role };
}

function col_(table, canonical, required) {
  var idx = table.headerIndex[normKey_(canonical)];
  if (idx !== undefined) return idx;

  var aliases = getConfig_().별칭[canonical] || [];
  for (var i = 0; i < aliases.length; i++) {
    var a = table.headerIndex[normKey_(aliases[i])];
    if (a !== undefined) return a;
  }

  var target = normKey_(canonical);
  for (var key in table.headerIndex) {
    if (key.indexOf(target) === 0) return table.headerIndex[key];
  }

  if (required) {
    throw new Error('[' + table.role + '] 시트에서 "' + canonical + '" 열을 찾지 못했습니다.');
  }
  return -1;
}

/** 신규 행을 헤더 바로 아래(2행)에 삽입해 최신 업무 데이터를 위에 두는다. */
function prependRows_(sheet, rows, textColumns) {
  if (!rows || !rows.length) return 2;
  var width = rows[0].length;

  sheet.insertRowsAfter(1, rows.length);

  if (textColumns && textColumns.length) {
    textColumns.forEach(function (c) {
      if (c >= 0 && c < width) sheet.getRange(2, c + 1, rows.length, 1).setNumberFormat('@');
    });
  }

  sheet.getRange(2, 1, rows.length, width).setValues(rows);
  return 2;
}

/** 기존 행 뒤에 추가한다. 현재는 S1의 신규 상품 등록에 사용한다. */
function appendRows_(sheet, rows, textColumns) {
  if (!rows || !rows.length) return 0;
  var start = sheet.getLastRow() + 1;
  var width = rows[0].length;

  if (textColumns && textColumns.length) {
    textColumns.forEach(function (c) {
      if (c >= 0 && c < width) sheet.getRange(start, c + 1, rows.length, 1).setNumberFormat('@');
    });
  }

  sheet.getRange(start, 1, rows.length, width).setValues(rows);
  return start;
}

/** 특정 열 하나를 2행부터 일괄 기록 */
function writeColumn_(sheet, colIdx, rows) {
  if (colIdx < 0 || !rows.length) return;
  var v = rows.map(function (r) { return [r[colIdx] === undefined ? '' : r[colIdx]]; });
  sheet.getRange(2, colIdx + 1, v.length, 1).setValues(v);
}

/* ============================================================
 *  값 헬퍼
 * ============================================================ */

function isBlank_(v) { return v === null || v === undefined || String(v).trim() === ''; }

function toStr_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  return String(v).trim();
}

function toNum_(v) {
  if (isBlank_(v)) return 0;
  var n = Number(String(v).replace(/,/g, ''));
  return isNaN(n) ? 0 : n;
}

function tz_() { return Session.getScriptTimeZone() || 'Asia/Seoul'; }

function 사용자_() {
  try { return Session.getActiveUser().getEmail(); } catch (e) { return ''; }
}

/* ============================================================
 *  로그 (최신이 위로)
 * ============================================================ */

function writeStockLog_(entries) {
  if (!entries || !entries.length) return;
  var sh = consoleSS_().getSheetByName(CONSOLE.재고이동로그);
  if (!sh) throw new Error('[' + CONSOLE.재고이동로그 + '] 탭이 없습니다.');

  var now = new Date();
  var rows = entries.map(function (e) {
    return [now, e.구분, e.피킹지시번호 || '', e.주문번호 || '', e.품목별주문번호 || '',
            e.상품코드 || '', e.변동량, e.변동후재고, e.담당자 || '', e.사유 || ''];
  });

  prependRows_(sh, rows);
}

function writeOpLog_(함수명, 결과, 메시지) {
  try {
    var sh = consoleSS_().getSheetByName(CONSOLE.작업로그);
    if (!sh) return;
    prependRows_(sh, [[new Date(), 함수명, 결과, String(메시지 || '').substring(0, 2000), 사용자_()]]);
  } catch (e) { }
}

/* ============================================================
 *  동시성 · 알림
 * ============================================================ */

function withLock_(fn) {
  // processInput → S1/S2 → S3 → S4/S9처럼 한 실행 안에서 lock 함수가 중첩 호출될 수 있다.
  // 같은 실행 컨텍스트의 내부 호출은 깊이만 증가시켜 ScriptLock 재획득 대기를 피한다.
  if (_scriptLockDepth > 0) {
    _scriptLockDepth++;
    try { return fn(); } finally { _scriptLockDepth--; }
  }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) throw new Error('다른 작업이 실행 중입니다. 30초 후 다시 시도하세요.');
  _scriptLockDepth = 1;
  try { return fn(); } finally { _scriptLockDepth = 0; lock.releaseLock(); }
}

function alert_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); return; } catch (e) { }
  try { consoleSS_().toast(String(msg).substring(0, 300), '피킹 시스템', 15); } catch (e2) { }
  Logger.log(msg);
  writeOpLog_('알림', '정보', msg);
}

/* ============================================================
 *  대시보드 공용 유틸
 * ============================================================ */

/** 단일 운영 Spreadsheet의 대시보드 탭을 확보한다. */
function ensureDashSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name, 0);
    sh.setHiddenGridlines(true);
  }
  return sh;
}

var DASHCOLOR = {
  제목: '1F3864',
  카드: 'F2F6FA',
  헤더: 'E8EDF3',
  대기: 'D9D9D9',
  진행: 'BDD7EE',
  완료: 'C6E0B4',
  예외: 'F8CBAD',
  경고: 'FCE4D6',
  위험: 'FFC7CE',
  좋음: 'E2EFDA',
  선: 'BFC9D4'
};

/** onOpen 트리거만 제거 — 시트가 안 열릴 때 응급조치 */
function 응급_onOpen트리거제거() {
  var 제거 = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onOpen') {
      ScriptApp.deleteTrigger(t);
      제거++;
    }
  });
  var msg = 'onOpen 트리거 ' + 제거 + '개 제거. 이제 시트를 다시 열어보세요.\n' +
            '(메뉴는 사라집니다. 나중에 설치_3_트리거등록()으로 복구)';
  Logger.log(msg);
  return msg;
}
