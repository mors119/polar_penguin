/**
 * ============================================================
 *  01. 설치 호환 함수 및 운영 시트 정비
 * ============================================================
 *  공식 설치 진입점은 setupSystem()이다. 이 파일은 기존 설치 함수명과의 호환,
 *  시트 Validation·서식 정비, 중복 없는 트리거 재구성, 메뉴와 진단 도구를 맡는다.
 * ============================================================
 */

var DEFAULT_FOLDER_ID = '';

function 설치_1_초기설정() { return setupSystem(); }

function 설정_파일URL등록() {
  throw new Error('수동 Spreadsheet URL 등록은 더 이상 사용하지 않습니다. setRootFolder() 후 setupSystem()을 실행하세요.');
}

function 설정_파라미터보충() { return setupSystem(); }

/* ============================================================
 *  설치 2. 시트 정비
 * ============================================================ */

/**
 * 기존 데이터를 유지하면서 필수 운영 컬럼, 기본값, Validation과 입력 서식을 보장한다.
 *
 * @param {boolean=} silent true면 완료 알림을 표시하지 않는다.
 * @return {string} 역할별 정비 결과
 * @sideEffect 누락 열을 오른쪽에 추가하고 주문·피킹·상품마스터 입력 규칙을 갱신한다.
 */
function 설치_2_시트정비(silent) {
  var 결과 = [];

  /* ---------- 주문 (완료) ---------- */
  var 주문 = readTable_(ROLE.주문);
  var 추가주문 = ensureColumns_(주문.sheet, 주문.headers, [
    COL.피킹지시번호, COL.주문상태, COL.취소사유, COL.취소일시, COL.확정일시, COL.대기사유
  ]);
  결과.push('주문(완료): ' + (추가주문.length ? '열 추가 ' + 추가주문.join(', ') : '변경 없음'));

  주문 = readTable_(ROLE.주문);
  var c주문상태 = col_(주문, COL.주문상태, true);
  if (주문.rows.length) {
    var 상태값 = 주문.rows.map(function (r) {
      var s = toStr_(r[c주문상태]);
      if (!s || s === '정상') return [ENUM.주문상태.접수];
      return [r[c주문상태]];
    });
    주문.sheet.getRange(2, c주문상태 + 1, 상태값.length, 1).setValues(상태값);
  }

  주문.sheet.getRange(2, c주문상태 + 1, Math.max(주문.sheet.getMaxRows() - 1, 1), 1)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList([ENUM.주문상태.접수, ENUM.주문상태.확정,
                           ENUM.주문상태.예약대기, ENUM.주문상태.취소], true)
      .setAllowInvalid(false).build());

  [COL.주문번호, COL.품목별주문번호, COL.상품품목코드, '수령인 우편번호', '수령인 휴대전화']
    .forEach(function (name) {
      var idx = col_(주문, name, false);
      if (idx >= 0) 주문.sheet.getRange(1, idx + 1, 주문.sheet.getMaxRows(), 1).setNumberFormat('@');
    });

  /* ---------- 피킹 (헤더) ---------- */
  var 헤더 = readTable_(ROLE.헤더);
  var 추가헤더 = ensureColumns_(헤더.sheet, 헤더.headers, [COL.출력일시]);
  결과.push('피킹(헤더): ' + (추가헤더.length ? '열 추가 ' + 추가헤더.join(', ') : '변경 없음'));

  헤더 = readTable_(ROLE.헤더);
  var c헤더상태 = col_(헤더, COL.상태, true);
  헤더.sheet.getRange(2, c헤더상태 + 1, Math.max(헤더.sheet.getMaxRows() - 1, 1), 1)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList([ENUM.헤더상태.대기, ENUM.헤더상태.진행,
                           ENUM.헤더상태.완료, ENUM.헤더상태.예외], true)
      .setAllowInvalid(false).build());

  // 담당자 칸을 노란색으로 — 여기 이름을 적으면 라인까지 전파된다
  var c담당 = col_(헤더, COL.피킹담당자, true);
  헤더.sheet.getRange(2, c담당 + 1, Math.max(헤더.sheet.getMaxRows() - 1, 1), 1)
    .setBackground('FFF9E6');
  결과.push('피킹(헤더): 담당자 칸 표시 (여기 이름을 적으면 라인에 자동 전파)');

  /* ---------- 피킹 (라인) ---------- */
  var 라인 = readTable_(ROLE.라인);
  var 추가라인 = ensureColumns_(라인.sheet, 라인.headers, [
    COL.피킹지시번호, COL.담당자, COL.라인상태, COL.처리일시
  ]);
  결과.push('피킹(라인): ' + (추가라인.length ? '열 추가 ' + 추가라인.join(', ') : '변경 없음'));

  라인 = readTable_(ROLE.라인);
  var c확인 = col_(라인, COL.확인, true);
  var c예외 = col_(라인, COL.예외사유, true);
  var c라인상태 = col_(라인, COL.라인상태, true);

  if (라인.rows.length) {
    var 라인상태값 = 라인.rows.map(function (r) {
      return [isBlank_(r[c라인상태]) ? ENUM.라인상태.미처리 : r[c라인상태]];
    });
    라인.sheet.getRange(2, c라인상태 + 1, 라인상태값.length, 1).setValues(라인상태값);
  }

  var maxRows = Math.max(라인.sheet.getMaxRows() - 1, 1);
  라인.sheet.getRange(2, c확인 + 1, maxRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList([ENUM.확인.정상, ENUM.확인.예외], true)
      .setAllowInvalid(false)
      .setHelpText('O = 정상 / X = 예외(주문 전체 취소)').build());
  라인.sheet.getRange(2, c예외 + 1, maxRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(ENUM.예외사유, true)
      .setAllowInvalid(false).build());
  라인.sheet.getRange(1, c확인 + 1).setNote(
    'O = 정상, X = 예외입니다. X 입력 시 예외사유를 선택한 뒤 상단 📦 Polar Penguin → 주문/피킹 → 결과 반영을 실행하세요.');
  라인.sheet.getRange(1, c예외 + 1).setNote('X인 행은 재고없음 또는 불량재고를 선택하세요.');
  라인.sheet.setColumnWidth(c확인 + 1, 80);
  라인.sheet.setColumnWidth(c예외 + 1, 120);
  결과.push('피킹(라인): O/X · 예외사유 드롭다운 적용');

  /* ---------- 상품마스터 ---------- */
  var 마스터 = readTable_(ROLE.마스터);
  var 추가마스터 = ensureColumns_(마스터.sheet, 마스터.headers, [
    COL.예약재고, COL.예약상품, COL.재고관리, COL.판매가, COL.최종동기화
  ]);
  마스터 = readTable_(ROLE.마스터);

  [COL.예약재고, COL.불량재고].forEach(function (n) {
    var c = col_(마스터, n, false);
    if (c >= 0 && 마스터.rows.length) {
      마스터.sheet.getRange(2, c + 1, 마스터.rows.length, 1)
        .setValues(마스터.rows.map(function (r) {
          return [isBlank_(r[c]) ? 0 : toNum_(r[c])];
        }));
    }
  });

  [COL.상품품목코드, COL.상품명, COL.기본보관위치, COL.가용재고].forEach(function (n) {
    col_(마스터, n, true);
  });
  결과.push('상품마스터: ' + (추가마스터.length ? '열 추가 ' + 추가마스터.join(', ') : '변경 없음'));

  var msg = '시트 정비 완료\n\n' + 결과.join('\n') +
            '\n\n참고: 공식 설치는 setupSystem()이 트리거와 대시보드까지 연속 구성합니다.';
  if (!silent) alert_(msg);
  writeOpLog_('설치_2_시트정비', '성공', 결과.join(' / '));
  return msg;
}

/* ============================================================
 *  설치 3. 트리거
 * ============================================================ */

/**
 * 설정의 폴링주기에 맞춰 통합 Input, 결과 반영/대시보드 및 메뉴 트리거를 등록한다.
 * 같은 역할의 현재·레거시 handler를 먼저 제거한 뒤 하나씩 다시 만들어 중복을 막는다.
 *
 * @param {boolean=} silent true면 완료 알림을 표시하지 않는다.
 * @return {string} 적용한 주기와 메뉴 트리거 결과
 * @sideEffect Apps Script 프로젝트 트리거를 안전하게 재구성한다.
 */
function 설치_3_트리거등록(silent) {
  var ss = consoleSS_();

  ScriptApp.getProjectTriggers().forEach(function (t) {
    var f = t.getHandlerFunction();
    if (['S5_1_결과반영', 'syncAndRefresh', 'syncPickingResults', 'processInput', 'onOpen'].indexOf(f) >= 0) {
      ScriptApp.deleteTrigger(t);
    }
  });

  var 주기 = Number(param_('폴링주기(분)', 5));
  if ([1, 5, 10, 15, 30].indexOf(주기) < 0) 주기 = 5;

  ScriptApp.newTrigger('syncAndRefresh').timeBased().everyMinutes(주기).create();
  ScriptApp.newTrigger('processInput').timeBased().everyMinutes(주기).create();

  var 메뉴 = true;
  try { ScriptApp.newTrigger('onOpen').forSpreadsheet(ss).onOpen().create(); }
  catch (e) { 메뉴 = false; }

  var msg = 주기 + '분 주기 트리거를 등록했습니다.\n' +
    '통합 Input 처리 + 재고 반영 + 대시보드 갱신이 함께 실행됩니다.\n' +
    (메뉴 ? '"' + ss.getName() + '" 을 열면 메뉴가 표시됩니다.' : '⚠ 메뉴 트리거 등록 실패');

  if (!silent) alert_(msg);
  writeOpLog_('설치_3_트리거등록', '성공', 주기 + '분');
  return msg;
}

/* ============================================================
 *  진단
 * ============================================================ */

function 점검_연결확인() {
  var out = ['운영 Spreadsheet: ' + consoleSS_().getName()];
  [ROLE.마스터, ROLE.주문, ROLE.헤더, ROLE.라인].forEach(function (role) {
    try {
      var sh = getSheet_(role);
      out.push('· ' + role + ' → ' + sh.getParent().getName() + ' / ' + sh.getName() +
               ' (' + sh.getLastRow() + '행 × ' + sh.getLastColumn() + '열)');
    } catch (e) {
      out.push('· ' + role + ' → ❌ ' + e.message);
    }
  });
  var msg = out.join('\n');
  Logger.log(msg);
  alert_(msg);
  return msg;
}

function 진단_시트구조() {
  var out = [];

  function 확인(role, 필수) {
    var t;
    try { t = readTable_(role); }
    catch (e) { out.push('❌ ' + role + ': ' + e.message); return; }
    var 없음 = 필수.filter(function (n) { return col_(t, n, false) < 0; });
    out.push((없음.length ? '❌ ' : '✅ ') + role + ' (' + t.rows.length + '행 × ' + t.width + '열)' +
             (없음.length ? '  누락: ' + 없음.join(', ') : ''));
  }

  확인(ROLE.마스터, [COL.상품품목코드, COL.상품명, COL.기본보관위치, COL.가용재고, COL.예약재고, COL.예약상품]);
  확인(ROLE.주문, [COL.주문번호, COL.품목별주문번호, COL.상품품목코드, COL.수량,
                  COL.출고완료, COL.피킹지시번호, COL.주문상태, COL.확정일시]);
  확인(ROLE.헤더, [COL.피킹지시번호, COL.주문번호, COL.카트슬롯, COL.품목수, COL.총수량,
                  COL.피킹담당자, COL.상태]);
  확인(ROLE.라인, [COL.순번, COL.보관위치, COL.상품코드, COL.필요수량, COL.확인,
                  COL.실제수량, COL.품목별주문번호, COL.피킹지시번호, COL.담당자, COL.라인상태]);

  var ss = consoleSS_();
  ['📖 안내', '📊 대시보드', '예약대기', '주문반려', CONSOLE.설정,
   CONSOLE.재고이동로그, CONSOLE.작업로그, CONSOLE.입력처리로그].forEach(function (n) {
    out.push((ss.getSheetByName(n) ? '✅ ' : '❌ ') + '운영 탭: ' + n);
  });

  out.push((ss.getSheetByName('📊 대시보드') ? '✅ ' : '❌ ') + '통합 대시보드');

  var trg = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  out.push((trg.length ? '✅ ' : '❌ ') + '트리거: ' + (trg.join(', ') || '없음'));

  var msg = out.join('\n');
  Logger.log(msg);
  alert_(msg);
  return msg;
}

/* ============================================================
 *  헬퍼
 * ============================================================ */

function ensureColumns_(sheet, headers, names) {
  var 존재 = {};
  headers.forEach(function (h) { 존재[normKey_(h)] = true; });

  var 추가 = names.filter(function (n) { return !존재[normKey_(n)]; });
  if (!추가.length) return [];

  // 사용 중인 열을 이동하지 않고 현재 헤더 오른쪽에 누락 열만 덧붙인다.
  var start = headers.length + 1;
  var 필요 = start + 추가.length - 1;
  if (sheet.getMaxColumns() < 필요) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 필요 - sheet.getMaxColumns());
  }
  sheet.getRange(1, start, 1, 추가.length).setValues([추가]).setFontWeight('bold');
  return 추가;
}

/* ============================================================
 *  메뉴
 * ============================================================ */

function onOpen(e) {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('📦 Polar Penguin')
      .addSubMenu(ui.createMenu('📥 Input 처리')
        .addItem('Input 지금 처리', 'processInput'))
      .addSubMenu(ui.createMenu('📦 주문/피킹')
        .addItem('주문 확정', 'S3_1_주문확정')
        .addItem('피킹지시 생성', 'S4_1_피킹지시생성')
        .addItem('결과 반영', 'S5_2_수동반영')
        .addItem('작업지시서 출력', 'S9_1_작업지시서출력'))
      .addSubMenu(ui.createMenu('📊 운영')
        .addItem('대시보드 갱신', 'D0_대시보드전체갱신')
        .addItem('예약대기 조회', '운영_예약대기조회')
        .addItem('시스템 상태 확인', '진단_시트구조'))
      .addSubMenu(ui.createMenu('⚙ 관리')
        .addItem('시스템 설치/복구', 'setupSystem')
        .addItem('로그 정리', '정리_로그')
        .addItem('설정 캐시 초기화', '설정_캐시초기화'))
      .addToUi();
  } catch (err) {
    Logger.log('메뉴 생성 실패: ' + err.message);
  }
}

function 운영_예약대기조회() {
  refreshOperationalViews_();
  var ss = consoleSS_();
  var sh = ss.getSheetByName('예약대기');
  if (sh) ss.setActiveSheet(sh);
  return S8_2_예약대기조회();
}

/** 재고이동로그를 최근 N행만 남기고 정리 */
function 정리_로그(남길행수) {
  남길행수 = 남길행수 || 2000;
  var ss = consoleSS_();

  [CONSOLE.재고이동로그, CONSOLE.작업로그].forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    var last = sh.getLastRow();
    if (last <= 남길행수 + 1) return;

    // 최신이 위에 있으므로 아래쪽(오래된 것)을 지운다
    var 지울행 = last - 남길행수 - 1;
    sh.deleteRows(남길행수 + 2, 지울행);
    Logger.log(name + ': ' + 지울행 + '행 삭제 (남은 ' + sh.getLastRow() + '행)');
  });
}



/** 호환 함수명을 유지하며 통합 Input의 파일별 판별 결과를 보여준다. */
function 진단_주문폴더() {
  var raw = param_('통합Input폴더ID', param_('CSV폴더ID', DEFAULT_FOLDER_ID));
  var m = String(raw).match(/[-\w]{25,}/);
  var 폴더ID = m ? m[0] : String(raw);

  var out = ['설정값: ' + raw, '추출 ID: ' + 폴더ID];

  var folder;
  try {
    folder = DriveApp.getFolderById(폴더ID);
    out.push('폴더명: ' + folder.getName());
  } catch (e) {
    out.push('❌ 폴더를 열 수 없음: ' + e.message);
    Logger.log(out.join('\n'));
    return;
  }

  out.push('');
  var it = folder.getFiles(), n = 0;
  while (it.hasNext()) {
    var f = it.next();
    var name = f.getName();
    var mime = f.getMimeType();
    var 타입 = mime.indexOf('spreadsheet') >= 0 ? '구글시트'
             : mime.indexOf('csv') >= 0 ? 'CSV'
             : mime.indexOf('excel') >= 0 ? '엑셀' : mime;
    var 판별 = INPUT_TYPE.UNKNOWN;
    var 오류 = '';
    try { 판별 = detectInputType_(readUnifiedInput_(f)[0] || []); }
    catch (e2) { 오류 = e2.inputCode || e2.message; }

    out.push((판별 !== INPUT_TYPE.UNKNOWN && !오류 ? '✅ ' : '❌ ') + name);
    out.push('     ' + 타입 + '  |  판별: ' + 판별 + (오류 ? '  |  오류: ' + 오류 : ''));
    n++;
  }

  if (!n) out.push('(폴더가 비어 있습니다)');
  out.push('');
  out.push('※ processInput()과 동일하게 Input 직하위 파일만 검사합니다.');

  Logger.log(out.join('\n'));
}
