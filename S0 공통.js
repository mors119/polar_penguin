

/**
 * ============================================================
 *  00. 공통  (v4.0)
 * ============================================================
 *  v4 변경
 *   · 신규 데이터를 시트 최상단에 쌓는다 (prependRows_)
 *   · 피킹라인에 담당자 열 추가
 *   · 입고 관련 상수 제거 (카페24 동기화로 일원화)
 * ============================================================
 */

/** 설정·로그 탭이 있는 스프레드시트 ID */
var CONSOLE_SS_ID = '1PwHJOkP-xRe7rDaZMeicRhg-BKviCYhLytJG_zpBLO8';

var CONSOLE = {
  설정: '설정',
  재고이동로그: '재고이동로그',
  작업로그: '작업로그'
};

var ROLE = {
  마스터: '상품마스터',
  주문: '주문완료',
  헤더: '피킹헤더',
  라인: '피킹라인'
};

/** 표준 컬럼명 */
var COL = {
  // ---------- 상품마스터 ----------
  상품품목코드: '상품품목코드',
  상품명: '상품명',
  옵션명: '옵션명',
  이미지: '이미지',
  기본보관위치: '기본보관위치',
  가용재고: '가용재고',
  예약재고: '예약재고',
  불량재고: '불량재고',
  상품상태: '상품상태',
  예약상품: '예약상품',
  재고관리: '재고관리',
  판매가: '판매가',
  최종동기화: '최종동기화',

  // ---------- 주문 (완료) ----------
  주문번호: '주문번호',
  품목별주문번호: '품목별 주문번호',
  수량: '수량',
  출고완료: '출고완료',
  피킹지시번호: '피킹지시번호',
  주문상태: '주문상태',
  취소사유: '취소사유',
  취소일시: '취소일시',
  확정일시: '확정일시',
  대기사유: '대기사유',

  // ---------- 피킹 (헤더) ----------
  카트슬롯: '카트 슬롯',
  품목수: '품목수',
  총수량: '총수량',
  피킹담당자: '피킹담당자',
  상태: '상태',
  출력일시: '출력일시',

  // ---------- 피킹 (라인) ----------
  순번: '순번',
  보관위치: '보관위치',
  상품코드: '상품코드',
  옵션: '옵션',
  필요수량: '필요수량',
  확인: '확인',
  실제수량: '실제수량',
  예외사유: '예외사유',
  담당자: '담당자',
  라인상태: '라인상태',
  처리일시: '처리일시'
};

var ENUM = {
  헤더상태: { 대기: '대기', 진행: '진행', 완료: '완료', 예외: '예외' },
  라인상태: { 미처리: '미처리', 차감완료: '차감완료', 복원완료: '복원완료', 취소마감: '취소마감' },
  주문상태: { 접수: '접수', 확정: '확정', 예약대기: '예약대기', 취소: '취소' },
  확인: { 정상: 'O', 예외: 'X' },
  예외사유: ['재고없음', '불량재고'],
  로그구분: { 차감: '차감', 복원: '복원', 예약: '예약', 예약해제: '예약해제', 동기화: '동기화' }
};

var _cache = { config: null, ss: {}, consoleSS: null };

/* ============================================================
 *  콘솔 스프레드시트 접근
 * ============================================================ */

function consoleSS_() {
  if (_cache.consoleSS) return _cache.consoleSS;

  var ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }

  if (!ss) {
    var id = PropertiesService.getScriptProperties().getProperty('CONSOLE_SS_ID') || CONSOLE_SS_ID;
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

function 설정_콘솔파일지정() {
  var 입력 = PropertiesService.getScriptProperties().getProperty('CONSOLE_SS_ID') || CONSOLE_SS_ID;
  if (!입력) throw new Error('CONSOLE_SS_ID 에 스프레드시트 ID나 URL을 입력하세요.');

  var m = String(입력).match(/[-\w]{25,}/);
  if (!m) throw new Error('ID를 찾지 못했습니다: ' + 입력);

  var ss = SpreadsheetApp.openById(m[0]);
  PropertiesService.getScriptProperties().setProperty('CONSOLE_SS_ID', m[0]);
  _cache.consoleSS = null;

  var msg = '콘솔 파일 지정 완료: ' + ss.getName();
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

function param_(키, 기본값) {
  var v = getConfig_().파라미터[키];
  return (v === undefined || v === null || v === '') ? 기본값 : v;
}

/* ============================================================
 *  외부 파일 접근
 * ============================================================ */

function openSS_(role) {
  if (_cache.ss[role]) return _cache.ss[role];

  var raw = getConfig_().파일ID[role];
  var m = String(raw || '').match(/[-\w]{25,}/);
  if (!m) throw new Error('[설정] 탭의 파일ID에 "' + role + '" 가 없습니다.');

  var ss;
  try { ss = SpreadsheetApp.openById(m[0]); }
  catch (e) { throw new Error('"' + role + '" 파일을 열 수 없습니다. (ID: ' + m[0] + ')'); }

  _cache.ss[role] = ss;
  return ss;
}

function getSheet_(role) {
  var ss = openSS_(role);
  var name = getConfig_().시트명[role];
  var sh = name ? ss.getSheetByName(name) : null;
  if (!sh) sh = ss.getSheets()[0];
  if (!sh) throw new Error('"' + role + '" 파일에 시트가 없습니다.');
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

/**
 * ★ v4 — 신규 행을 시트 최상단(2행)에 삽입한다.
 *   최신 데이터가 위로 쌓여 관리자·작업자가 스크롤 없이 확인할 수 있다.
 */
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

/** 하단에 추가 (로그 등 시간순 보존이 필요한 곳에서만 사용) */
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

/** 상품명이 예약판매 상품인지 판단 */
function isPreorderName_(상품명) {
  var 키워드 = String(param_('예약키워드', '예약')).split(',')
    .map(function (s) { return s.trim(); }).filter(String);
  var n = String(상품명 || '');
  for (var i = 0; i < 키워드.length; i++) {
    if (n.indexOf(키워드[i]) >= 0) return true;
  }
  return false;
}

/* ============================================================
 *  로그  (★ v4 — 최신이 위로)
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

/** 품목별 주문번호별 순변동 집계 (복원 계산의 근거) */
function readStockLogNet_() {
  var sh = consoleSS_().getSheetByName(CONSOLE.재고이동로그);
  var net = {};
  if (!sh || sh.getLastRow() < 2) return net;

  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
  for (var i = 0; i < v.length; i++) {
    var key = toStr_(v[i][4]);
    if (!key) continue;
    net[key] = (net[key] || 0) + toNum_(v[i][6]);
  }
  return net;
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
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) throw new Error('다른 작업이 실행 중입니다. 30초 후 다시 시도하세요.');
  try { return fn(); } finally { lock.releaseLock(); }
}

function alert_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); return; } catch (e) { }
  try { consoleSS_().toast(String(msg).substring(0, 300), '피킹 시스템', 15); } catch (e2) { }
  Logger.log(msg);
  writeOpLog_('알림', '정보', msg);
}

function toast_(msg, title) {
  try { consoleSS_().toast(String(msg).substring(0, 300), title || '피킹 시스템', 10); } catch (e) { }
  Logger.log(msg);
}

function 설정_캐시초기화() {
  _cache.config = null;
  _cache.ss = {};
  _cache.consoleSS = null;
  alert_('설정 캐시를 초기화했습니다.');
}

/* ============================================================
 *  대시보드 공용 유틸
 * ============================================================ */

/** 진행률 막대 */
function bar_(pct, len) {
  len = len || 12;
  var filled = Math.max(0, Math.min(len, Math.round((pct / 100) * len)));
  return new Array(filled + 1).join('█') + new Array(len - filled + 1).join('░');
}

/** 대시보드 탭을 확보한다 (외부 파일에도 만들 수 있다) */
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

/** v3 잔재 탭 삭제 + 작업로그 정리 (시트를 열지 않고 실행) */
function 정리_불필요탭() {
  var ss = consoleSS_();
  var out = [];

  ['입고', '대시보드'].forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) { out.push('· ' + name + ' — 이미 없음'); return; }
    ss.deleteSheet(sh);
    out.push('✅ ' + name + ' 탭 삭제');
  });

  // 작업로그를 최근 500행만 남긴다 (최신이 위)
  var log = ss.getSheetByName(CONSOLE.작업로그);
  if (log && log.getLastRow() > 501) {
    var 지울행 = log.getLastRow() - 501;
    log.deleteRows(502, 지울행);
    out.push('✅ 작업로그 ' + 지울행 + '행 삭제 → ' + log.getLastRow() + '행');
  }

  var msg;
}
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