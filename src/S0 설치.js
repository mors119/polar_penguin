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
  var 주문이전 = migrateOrderSheetSchema_(주문);
  결과.push('주문(완료): ' + (주문이전.changed
    ? '고정 25열 정렬 / 동적 ' + 주문이전.dynamicHeaders.length + '열 보존'
    : '변경 없음'));

  주문 = readTable_(ROLE.주문);
  var c주문상태 = col_(주문, COL.주문상태, true);
  if (주문.rows.length) {
    var 상태값 = 주문.rows.map(function (r) {
      var s = toStr_(r[c주문상태]);
      if (toNum_(r[col_(주문, COL.출고완료, true)]) === 1) return [ENUM.주문상태.출고완료];
      if (!s || s === '정상' || s === '접수' || s === '예약대기') return [ENUM.주문상태.예약];
      if (s === '확정' || s === '처리완료') return [ENUM.주문상태.예약];
      return [r[c주문상태]];
    });
    주문.sheet.getRange(2, c주문상태 + 1, 상태값.length, 1).setValues(상태값);
  }

  주문.sheet.getRange(2, c주문상태 + 1, Math.max(주문.sheet.getMaxRows() - 1, 1), 1)
    .setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList([ENUM.주문상태.예약, ENUM.주문상태.출고완료, ENUM.주문상태.취소], true)
      .setAllowInvalid(false).build());

  [COL.쇼핑몰번호, COL.주문번호, COL.품목별주문번호, COL.상품품목코드,
   COL.수령인우편번호, COL.수령인휴대전화]
    .forEach(function (name) {
      var idx = col_(주문, name, false);
      if (idx >= 0) 주문.sheet.getRange(1, idx + 1, 주문.sheet.getMaxRows(), 1).setNumberFormat('@');
    });

  /* ---------- 피킹 (헤더) ---------- */
  var 헤더 = readTable_(ROLE.헤더);
  var 추가헤더 = ensureColumns_(헤더.sheet, 헤더.headers, [COL.생성일시, COL.출력일시]);
  결과.push('피킹(헤더): ' + (추가헤더.length ? '열 추가 ' + 추가헤더.join(', ') : '변경 없음'));

  헤더 = readTable_(ROLE.헤더);
  var c헤더상태 = col_(헤더, COL.상태, true);
  if (헤더.rows.length) {
    헤더.sheet.getRange(2, c헤더상태 + 1, 헤더.rows.length, 1).setValues(헤더.rows.map(function (r) {
      var s = toStr_(r[c헤더상태]);
      if (s === '예외') return [ENUM.헤더상태.취소];
      if (s === '진행') return [ENUM.헤더상태.대기];
      return [s || ENUM.헤더상태.대기];
    }));
  }
  헤더.sheet.getRange(2, c헤더상태 + 1, Math.max(헤더.sheet.getMaxRows() - 1, 1), 1)
    .clearDataValidations().setBackground('F2F2F2');
  헤더.sheet.getRange(1, c헤더상태 + 1).setNote('시스템 관리 필드입니다. 출력 결과에 따라 자동 변경됩니다.');

  // 기존 담당자 열은 감사 메모용으로만 남길 수 있으며 출력의 선행조건이 아니다.
  var c담당 = col_(헤더, COL.피킹담당자, false);
  if (c담당 >= 0) {
    헤더.sheet.getRange(2, c담당 + 1, Math.max(헤더.sheet.getMaxRows() - 1, 1), 1)
      .setBackground('F2F2F2');
    헤더.sheet.getRange(1, c담당 + 1).setNote('선택적 감사 메모입니다. 피킹지시서 출력에 필요하지 않습니다.');
  }

  /* ---------- 피킹 (라인) ---------- */
  var 라인 = readTable_(ROLE.라인);
  var 추가라인 = ensureColumns_(라인.sheet, 라인.headers, [
    COL.주문번호, COL.피킹지시번호, COL.담당자, COL.라인상태, COL.처리일시
  ]);
  결과.push('피킹(라인): ' + (추가라인.length ? '열 추가 ' + 추가라인.join(', ') : '변경 없음'));

  라인 = readTable_(ROLE.라인);
  var c확인 = col_(라인, COL.확인, true);
  var c예외 = col_(라인, COL.예외사유, true);
  var c라인상태 = col_(라인, COL.라인상태, true);

  if (라인.rows.length) {
    var 라인상태값 = 라인.rows.map(function (r) {
      var s = toStr_(r[c라인상태]);
      if (s === '차감완료') return [ENUM.라인상태.완료];
      if (s === '복원완료' || s === '취소마감') return [ENUM.라인상태.취소];
      return [s || ENUM.라인상태.미처리];
    });
    라인.sheet.getRange(2, c라인상태 + 1, 라인상태값.length, 1).setValues(라인상태값);
  }

  var maxRows = Math.max(라인.sheet.getMaxRows() - 1, 1);
  라인.sheet.getRange(2, c확인 + 1, maxRows, 1).clearDataValidations();
  라인.sheet.getRange(2, c예외 + 1, maxRows, 1).clearDataValidations();
  라인.sheet.getRange(1, c확인 + 1).setNote(
    '시스템 관리 필드입니다. 출력 성공 시 자동으로 O가 기록됩니다.');
  라인.sheet.getRange(1, c예외 + 1).setNote('이전 버전 호환용 감사 필드입니다. 문제 처리는 선택 주문 취소를 사용하세요.');
  라인.headers.forEach(function (header, idx) {
    라인.sheet.getRange(1, idx + 1).setNote('시스템 생성/관리 필드입니다. 직접 입력하지 마세요.');
    라인.sheet.getRange(2, idx + 1, maxRows, 1).setBackground('F2F2F2');
  });
  라인.sheet.setColumnWidth(c확인 + 1, 80);
  라인.sheet.setColumnWidth(c예외 + 1, 120);
  [COL.순번, COL.보관위치, COL.상품코드, COL.상품명, COL.옵션, COL.필요수량,
   COL.실제수량, COL.담당자, COL.라인상태].forEach(function (name) {
    var idx = col_(라인, name, false);
    if (idx < 0) return;
    var widths = {};
    widths[COL.순번] = 55; widths[COL.보관위치] = 110; widths[COL.상품코드] = 130;
    widths[COL.상품명] = 240; widths[COL.옵션] = 160; widths[COL.필요수량] = 80;
    widths[COL.실제수량] = 80; widths[COL.담당자] = 90; widths[COL.라인상태] = 100;
    라인.sheet.setColumnWidth(idx + 1, widths[name]);
  });
  라인.sheet.setFrozenRows(1);
  결과.push('피킹(라인): 시스템 관리 감사 시트로 정비');

  /* ---------- 상품마스터 ---------- */
  var 마스터 = readTable_(ROLE.마스터);
  var 제거마스터 = migrateLegacyProductMasterSchema_(마스터);
  마스터 = readTable_(ROLE.마스터);
  var 추가마스터 = ensureColumns_(마스터.sheet, 마스터.headers, [
    COL.재고관리, COL.판매가, COL.최종동기화, '창고메모'
  ]);
  마스터 = readTable_(ROLE.마스터);

  [COL.상품품목코드, COL.상품명, COL.기본보관위치, COL.가용재고].forEach(function (n) {
    col_(마스터, n, true);
  });
  결과.push('상품마스터: ' + ([
    제거마스터.length ? '레거시 열 제거 ' + 제거마스터.join(', ') : '',
    추가마스터.length ? '열 추가 ' + 추가마스터.join(', ') : ''
  ].filter(String).join(' / ') || '변경 없음'));
  var c위치 = col_(마스터, COL.기본보관위치, true);
  마스터.sheet.getRange(1, c위치 + 1).setNote('창고 소유 필드입니다. 재고 동기화가 덮어쓰지 않으며 운영자가 직접 관리합니다.');
  마스터.sheet.getRange(2, c위치 + 1, Math.max(마스터.sheet.getMaxRows() - 1, 1), 1).setBackground('FFF9E6');
  var c창고메모 = col_(마스터, '창고메모', false);
  if (c창고메모 >= 0) {
    마스터.sheet.getRange(1, c창고메모 + 1).setNote('창고 소유 메모입니다. 자동 동기화가 덮어쓰지 않습니다.');
    마스터.sheet.getRange(2, c창고메모 + 1, Math.max(마스터.sheet.getMaxRows() - 1, 1), 1).setBackground('FFF9E6');
  }

  [COL.주문상태, COL.출고완료, COL.피킹지시번호, COL.확정일시, COL.취소일시, COL.취소경로]
    .forEach(function (name) {
      var idx = col_(주문, name, false); if (idx < 0) return;
      주문.sheet.getRange(1, idx + 1).setNote('시스템 관리 필드입니다. 메뉴와 자동 처리 결과로만 변경하세요.');
      주문.sheet.getRange(2, idx + 1, Math.max(주문.sheet.getMaxRows() - 1, 1), 1).setBackground('F2F2F2');
    });
  var c메모 = col_(주문, COL.운영메모, false);
  if (c메모 >= 0) 주문.sheet.getRange(2, c메모 + 1, Math.max(주문.sheet.getMaxRows() - 1, 1), 1).setBackground('FFF9E6');

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
 * 설정의 폴링주기에 맞춰 통합 Input 및 메뉴 트리거를 등록한다.
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

  ScriptApp.newTrigger('processInput').timeBased().everyMinutes(주기).create();

  var 메뉴 = true;
  try { ScriptApp.newTrigger('onOpen').forSpreadsheet(ss).onOpen().create(); }
  catch (e) { 메뉴 = false; }

  var msg = 주기 + '분 주기 트리거를 등록했습니다.\n' +
    '통합 Input 처리가 실행되며 각 변경 직후 대시보드가 갱신됩니다.\n' +
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

  확인(ROLE.마스터, [COL.상품품목코드, COL.상품명, COL.기본보관위치, COL.가용재고, COL.재고관리]);
  확인(ROLE.주문, [COL.주문번호, COL.품목별주문번호, COL.상품품목코드, COL.수량,
                  COL.출고완료, COL.피킹지시번호, COL.주문상태, COL.확정일시]);
  확인(ROLE.헤더, [COL.피킹지시번호, COL.주문번호, COL.품목수, COL.총수량, COL.상태]);
  확인(ROLE.라인, [COL.순번, COL.보관위치, COL.상품코드, COL.필요수량, COL.확인,
                  COL.실제수량, COL.품목별주문번호, COL.피킹지시번호, COL.담당자, COL.라인상태]);

  var ss = consoleSS_();
  ['📖 안내', '📊 대시보드', CONSOLE.설정,
   CONSOLE.재고이동로그, CONSOLE.작업로그, CONSOLE.입력처리로그, CONSOLE.처리주문아카이브].forEach(function (n) {
    out.push((ss.getSheetByName(n) ? '✅ ' : '❌ ') + '운영 탭: ' + n);
  });

  out.push('✅ 운영 Spreadsheet: ' + ss.getName());
  [
    ['Input 폴더', '통합Input폴더ID'], ['Success 폴더', 'Success폴더ID'],
    ['Error 폴더', 'Error폴더ID'], ['Output 폴더', 'Output폴더ID'], ['Backup 폴더', '백업폴더ID']
  ].forEach(function (item) {
    try { inputFolder_(item[1]); out.push('✅ ' + item[0]); }
    catch (e) { out.push('❌ ' + item[0] + ': ' + e.message); }
  });
  out.push(toStr_(param_('알림이메일', ''))
    ? '✅ 알림이메일'
    : '⚠ 알림이메일 미설정 → ⚙ 관리 → 설정에서 입력하세요.');

  var trg = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  out.push((trg.indexOf('processInput') >= 0 ? '✅ ' : '❌ ') + 'processInput 트리거');
  out.push((trg.indexOf('onOpen') >= 0 ? '✅ ' : '❌ ') + 'onOpen 트리거');

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

/**
 * 주문 시트를 고정 25열 + 실제 동적 열 구조로 재구성할 값을 계산한다.
 * 고정 열은 이름으로 찾아 순서를 바로잡고, 레거시 상품명/옵션명/주소는 정식 열이
 * 비어 있을 때만 보충한다. 이전 버전의 선택 열은 값이 있는 경우에만 동적으로 남긴다.
 */
function buildOrderSchemaMigration_(headers, rows) {
  headers = (headers || []).map(function (header) { return toStr_(header); });
  rows = rows || [];
  var sourceByKey = {};
  headers.forEach(function (header, index) {
    var key = normKey_(header);
    if (!key) return;
    if (!sourceByKey[key]) sourceByKey[key] = [];
    sourceByKey[key].push(index);
  });

  function sourceIndexes(canonical) {
    var names = [canonical].concat(ORDER_FIXED_HEADER_ALIASES[canonical] || []), indexes = [];
    names.forEach(function (name) {
      (sourceByKey[normKey_(name)] || []).forEach(function (index) {
        if (indexes.indexOf(index) < 0) indexes.push(index);
      });
    });
    return indexes;
  }

  var legacyOptional = {};
  ORDER_LEGACY_OPTIONAL_HEADERS.forEach(function (header) { legacyOptional[normKey_(header)] = true; });
  var dynamic = [], dynamicKeys = {}, removed = [];
  headers.forEach(function (header, index) {
    var key = normKey_(header);
    if (!key) return;
    if (canonicalOrderHeader_(header)) { if (ORDER_FIXED_HEADERS.indexOf(header) < 0) removed.push(header); return; }
    if (dynamicKeys[key]) { removed.push(header); return; }
    var hasValue = rows.some(function (row) { return !isBlank_(row[index]); });
    if (legacyOptional[key] && !hasValue) { removed.push(header); return; }
    dynamicKeys[key] = true;
    dynamic.push({ header: header, index: index });
  });

  var fixedSources = ORDER_FIXED_HEADERS.map(sourceIndexes);
  var migratedRows = rows.map(function (row) {
    var fixed = fixedSources.map(function (indexes) {
      for (var i = 0; i < indexes.length; i++) {
        if (!isBlank_(row[indexes[i]])) return row[indexes[i]];
      }
      return indexes.length ? row[indexes[0]] : '';
    });
    return fixed.concat(dynamic.map(function (column) { return row[column.index]; }));
  });
  var targetHeaders = ORDER_FIXED_HEADERS.concat(dynamic.map(function (column) { return column.header; }));
  var changed = headers.length !== targetHeaders.length || headers.some(function (header, index) {
    return header !== targetHeaders[index];
  });
  if (!changed) {
    changed = migratedRows.some(function (row, rowIndex) {
      return row.some(function (value, columnIndex) { return value !== rows[rowIndex][columnIndex]; });
    });
  }
  return { headers: targetHeaders, rows: migratedRows, dynamicHeaders: dynamic.map(function (column) {
    return column.header;
  }), removedHeaders: removed, changed: changed };
}

/** 기존 행을 지우지 않고 메모리에 보존한 뒤 표준 주문 스키마로 한 번에 다시 쓴다. */
function migrateOrderSheetSchema_(table) {
  var migrated = buildOrderSchemaMigration_(table.headers, table.rows);
  if (!migrated.changed) return migrated;
  var sheet = table.sheet, width = migrated.headers.length;
  if (sheet.getMaxColumns() < width) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), width - sheet.getMaxColumns());
  }
  sheet.clearContents();
  var values = [migrated.headers].concat(migrated.rows);
  sheet.getRange(1, 1, values.length, width).setValues(values);
  sheet.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('E8EDF3');
  return migrated;
}

/**
 * 기존 이중 재고 모델을 단일 가용재고로 이전한다.
 * 기존 가용재고는 S3에서 예약 수량을 이미 차감한 순 재고이므로 값을
 * 그대로 보존한다. 예약재고를 한 번 더 빼면 기존 확정 주문이 중복 차감된다.
 * 레거시 열은 헤더 이름으로 찾고 인덱스 변동을 피하기 위해 오른쪽부터 제거한다.
 */
function migrateLegacyProductMasterSchema_(table) {
  var removedNames = ['이미지', '예약재고', '상품상태', '예약상품', '불량재고'];
  var targets = [];
  (table.headers || []).forEach(function (header, index) {
    var name = toStr_(header);
    if (removedNames.indexOf(name) >= 0) targets.push({ name: name, column: index + 1 });
  });
  targets.sort(function (a, b) { return b.column - a.column; });
  targets.forEach(function (target) { table.sheet.deleteColumn(target.column); });
  return targets.map(function (target) { return target.name; });
}

/* ============================================================
 *  메뉴
 * ============================================================ */

function onOpen(e) {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('📦 Polar Penguin')
      .addSubMenu(ui.createMenu('📥 Input')
        .addItem('Input 지금 처리', 'processInput'))
      .addSubMenu(ui.createMenu('📋 주문')
        .addItem('선택 주문 취소', '선택_주문취소')
        .addItem('예약상품 입고 관리', '예약상품_입고관리'))
      .addSubMenu(ui.createMenu('📄 피킹지시서')
        .addItem('피킹지시서 조회 / 재출력', 'S9_1_작업지시서출력'))
      .addSubMenu(ui.createMenu('📍 위치 관리')
        .addItem('위치 미지정 상품', '위치_미지정상품관리'))
      .addSubMenu(ui.createMenu('📊 운영')
        .addItem('대시보드 갱신', 'D0_대시보드전체갱신')
        .addItem('시스템 상태 확인', '진단_시트구조'))
      .addSubMenu(ui.createMenu('⚙ 관리')
        .addItem('시스템 설치 / 복구', 'setupSystem')
        .addItem('설정', '설정_열기')
        .addItem('백업 및 정리', '백업_및_정리')
        .addItem('로그 정리', '정리_로그'))
      .addToUi();
  } catch (err) {
    Logger.log('메뉴 생성 실패: ' + err.message);
  }
}

/** 숨김 설정 탭만 표시하고 운영자가 가장 먼저 확인할 알림 설정으로 이동한다. */
function 설정_열기() {
  var ss = consoleSS_(), sheet = ss.getSheetByName(CONSOLE.설정);
  if (!sheet) {
    alert_('설정 시트를 찾을 수 없습니다.\n⚙ 관리 → 시스템 설치 / 복구를 실행하세요.');
    return null;
  }

  formatSettingsSheet_(sheet);
  if (typeof sheet.showSheet === 'function') sheet.showSheet();
  if (typeof ss.setActiveSheet === 'function') ss.setActiveSheet(sheet);

  var targetRow = findSettingsRow_(sheet, '파라미터', '알림이메일') || 1;
  var targetColumn = targetRow === 1 ? 1 : 3;
  var target = sheet.getRange(targetRow, targetColumn);
  if (target && typeof target.activate === 'function') target.activate();
  return sheet;
}

/** 이전 배포나 직접 호출에서 사용하던 함수명은 동일 동작으로 유지한다. */
function 설정_보기() { return 설정_열기(); }

function findSettingsRow_(sheet, section, key) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === section && String(values[i][1] || '').trim() === key) return i + 2;
  }
  return 0;
}

function isSystemManagedSetting_(section, key) {
  return section === '파일ID' || section === '시트명' || section === '별칭' ||
    ['통합Input폴더ID', 'Success폴더ID', 'Error폴더ID', 'Output폴더ID', '백업폴더ID'].indexOf(key) >= 0;
}

/** 기존 비고가 비어 있을 때만 운영 도움말을 보충한다. */
function ensureSettingsHelpText_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return;
  var help = {
    '알림이메일': '오류 및 백업/정리 결과를 받을 이메일 주소',
    '정리보존일수': '완료/취소 주문과 오래된 Success/Error/Output 정리 기준',
    '폴링주기(분)': 'Input 자동 확인 주기. 변경 후 시스템 설치 / 복구를 실행하세요.',
    '통합Input폴더ID': 'setupSystem()이 ROOT/Input 폴더를 자동 연결',
    'Success폴더ID': 'setupSystem()이 ROOT/Success 폴더를 자동 연결',
    'Error폴더ID': 'setupSystem()이 ROOT/Error 폴더를 자동 연결',
    'Output폴더ID': 'setupSystem()이 ROOT/Output 폴더를 자동 연결',
    '백업폴더ID': 'setupSystem()이 ROOT/Backup 폴더를 자동 연결'
  };
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  values.forEach(function (row, index) {
    var key = String(row[1] || '').trim();
    if (!String(row[3] || '').trim() && help[key]) sheet.getRange(index + 2, 4).setValue(help[key]);
  });
}

/** 4열 설정 테이블을 유지하면서 운영자용 시각 구분과 경고형 validation을 적용한다. */
function formatSettingsSheet_(sheet) {
  if (!sheet) return null;
  var lastRow = Math.max(sheet.getLastRow(), 1);
  sheet.getRange(1, 1, 1, 4)
    .setFontWeight('bold').setFontColor('FFFFFF').setBackground('274C5E')
    .setVerticalAlignment('middle').setWrap(true)
    .setNote('⚙ Polar Penguin 설정\n운영 환경과 알림·백업 정책을 관리합니다.');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 190);
  sheet.setColumnWidth(3, 280);
  sheet.setColumnWidth(4, 430);
  sheet.setHiddenGridlines(true);
  if (lastRow < 2) return sheet;

  var values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  values.forEach(function (row, index) {
    var sheetRow = index + 2;
    var section = String(row[0] || '').trim(), key = String(row[1] || '').trim();
    var systemManaged = isSystemManagedSetting_(section, key);
    var sectionColor = section === '파라미터' ? 'DDEBF7' :
      (section === '별칭' ? 'EDEDED' : 'E2F0D9');
    sheet.getRange(sheetRow, 1, 1, 2).setBackground(sectionColor).setFontWeight('bold');
    sheet.getRange(sheetRow, 1, 1, 4).setVerticalAlignment('middle').setWrap(true);

    var valueCell = sheet.getRange(sheetRow, 3).clearDataValidations()
      .setBackground(systemManaged ? 'F2F2F2' : 'FFF2CC')
      .setNote(systemManaged
        ? '시스템 관리 값입니다. setupSystem()이 유효성을 관리하므로 일반적으로 수정하지 않아도 됩니다.'
        : '운영자 설정입니다. 변경한 값은 다음 실행부터 적용됩니다.');

    if (key === '알림이메일') {
      valueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireTextIsEmail()
        .setAllowInvalid(true).setHelpText('비워 둘 수 있습니다. 입력할 때는 올바른 이메일 주소 형식을 사용하세요.').build());
    } else if (key === '정리보존일수') {
      valueCell.setDataValidation(SpreadsheetApp.newDataValidation()
        .requireFormulaSatisfied('=OR(C' + sheetRow + '="",AND(ISNUMBER(C' + sheetRow + '),C' + sheetRow + '>=1,MOD(C' + sheetRow + ',1)=0))')
        .setAllowInvalid(true).setHelpText('1 이상의 정수를 입력하세요. 비어 있으면 기본값 30일을 사용합니다.').build());
    } else if (key === '폴링주기(분)') {
      valueCell.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([1, 5, 10, 15, 30], true)
        .setAllowInvalid(true).setHelpText('지원 주기: 1, 5, 10, 15, 30분. 변경 후 시스템 설치 / 복구를 실행하세요.').build());
    }
  });
  return sheet;
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
