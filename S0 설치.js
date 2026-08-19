/**
 * ============================================================
 *  01. 설치  (v4.0)
 * ============================================================
 *  실행 순서
 *    0) 설정_콘솔파일지정()      ← 00_공통.gs 의 CONSOLE_SS_ID 입력 후
 *    1) 설치_1_초기설정()
 *    2) 설정_파일URL등록()
 *    3) 점검_연결확인()
 *    4) 설치_2_시트정비()        ← 신규 열·드롭다운·정렬
 *    5) 설치_3_트리거등록()      ← 5분 주기 + 메뉴
 *    6) D0_대시보드전체갱신()    ← 대시보드 3종 생성
 * ============================================================
 */

var DEFAULT_FOLDER_ID = '1NqqOtvW9ZG5ddoXmXoi9N1zVNuNjG26l';

var ROLE_HINTS = [
  { role: ROLE.마스터, keys: ['상품마스터', '상품 마스터', '마스터'] },
  { role: ROLE.주문,   keys: ['주문완료', '주문 완료', '주문'] },
  { role: ROLE.헤더,   keys: ['피킹해더', '피킹헤더', '피킹 해더', '피킹 헤더', '해더', '헤더'] },
  { role: ROLE.라인,   keys: ['피킹라인', '피킹 라인', '라인'] }
];

/* ============================================================
 *  설치 1
 * ============================================================ */

function 설치_1_초기설정() {
  var ss = consoleSS_();

  ensureSheet_(ss, CONSOLE.설정, ['구분', '키', '값', '비고']);
  ensureSheet_(ss, CONSOLE.재고이동로그, [
    '시각', '구분', '피킹지시번호', '주문번호', '품목별 주문번호',
    '상품코드', '변동량', '변동 후 재고', '담당자', '사유'
  ]);
  ensureSheet_(ss, CONSOLE.작업로그, ['시각', '함수', '결과', '메시지', '실행계정']);

  var 설정 = ss.getSheetByName(CONSOLE.설정);
  if (설정.getLastRow() > 1) {
    var m = '[설정] 탭에 이미 값이 있어 덮어쓰지 않았습니다.\n초기화하려면 2행 이하를 지우고 다시 실행하세요.';
    alert_(m);
    return m;
  }

  var 발견 = scanFolder_(DEFAULT_FOLDER_ID);
  var rows = [];

  [ROLE.마스터, ROLE.주문, ROLE.헤더, ROLE.라인].forEach(function (role) {
    rows.push(['파일ID', role, 발견[role] ? 발견[role].id : '',
      발견[role] ? 발견[role].name : '⚠ 설정_파일URL등록() 실행']);
  });
  [ROLE.마스터, ROLE.주문, ROLE.헤더, ROLE.라인].forEach(function (role) {
    rows.push(['시트명', role, 발견[role] ? 발견[role].firstSheet : '', '비우면 첫 시트 사용']);
  });

  rows.push(['파라미터', 'CSV폴더ID', DEFAULT_FOLDER_ID, '주문 CSV 업로드 폴더']);
  rows.push(['파라미터', '재고CSV폴더ID', DEFAULT_FOLDER_ID, '카페24 재고 CSV 폴더']);
  rows.push(['파라미터', 'CSV처리완료폴더명', '처리완료', '취입 후 CSV 이동 대상']);
  rows.push(['파라미터', '폴링주기(분)', 5, '변경 후 설치_3 재실행']);
  rows.push(['파라미터', '지시번호접두어', 'PK', '배치번호 형식']);
  rows.push(['파라미터', '예약키워드', '예약', '상품명에 이 단어가 있으면 예약상품']);
  rows.push(['파라미터', '재고경고임계치', 3, '재고현황 대시보드 경고 기준']);
  rows.push(['파라미터', '추가투입임계(분)', 45, '이 시간 내 완료 예상이면 주문 추가 투입 권고']);

  var 별칭 = [
    ['상품품목코드', '상품코드,품목코드,SKU,내부SKU'],
    ['상품코드', '상품품목코드,품목코드,SKU'],
    ['품목별 주문번호', '품목별주문번호,주문상세번호'],
    ['카트 슬롯', '카트 술룻,카트술룻,카트슬롯,슬롯'],
    ['상태', '상태(대기/진행/완료/예외),피킹상태'],
    ['기본보관위치', '보관위치,로케이션,위치'],
    ['보관위치', '기본보관위치,로케이션,위치'],
    ['옵션', '옵션명'],
    ['옵션명', '옵션'],
    ['수량', '주문수량'],
    ['필요수량', '주문수량,지시수량'],
    ['예약재고', '예약,확보재고'],
    ['담당자', '피킹담당자,작업자']
  ];
  별칭.forEach(function (a) { rows.push(['별칭', a[0], a[1], '']); });

  설정.getRange(2, 1, rows.length, 4).setValues(rows);
  설정.setFrozenRows(1);
  설정.autoResizeColumns(1, 4);
  _cache.config = null;

  var msg = '초기 설정 완료 — ' + ss.getName() + '\n\n' + describeScan_(발견) +
            '\n\n다음: 설정_파일URL등록() → 점검_연결확인() → 설치_2_시트정비()';
  alert_(msg);
  return msg;
}

/* ============================================================
 *  파일 URL 등록
 * ============================================================ */

function 설정_파일URL등록() {
  var 입력 = {
    '상품마스터': 'https://docs.google.com/spreadsheets/d/1XVnoHk4DauGSOSfO1-TyZXwBzQ0NGevy6SE7icOTMPQ/edit',
    '주문완료':   'https://docs.google.com/spreadsheets/d/1SOFqzaPoqmiJpt3MOhB1Hk7qZPoWdoE6vXy5kSOgJzs/edit',
    '피킹헤더':   'https://docs.google.com/spreadsheets/d/16gJieIY8vZ3gRzMgAvuG_6WuCaad0yTZd0lbgdYZuIY/edit',
    '피킹라인':   'https://docs.google.com/spreadsheets/d/1PwHJOkP-xRe7rDaZMeicRhg-BKviCYhLytJG_zpBLO8/edit'
  };

  var 설정 = consoleSS_().getSheetByName(CONSOLE.설정);
  if (!설정) throw new Error('[설정] 탭이 없습니다. 설치_1_초기설정() 을 먼저 실행하세요.');

  var v = 설정.getDataRange().getValues();
  var out = [];

  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]).trim() !== '파일ID') continue;
    var 역할 = String(v[i][1]).trim();
    if (!입력[역할]) continue;

    var m = String(입력[역할]).match(/[-\w]{25,}/);
    if (!m) { out.push('❌ ' + 역할 + ': ID를 못 찾음'); continue; }

    var name;
    try { name = SpreadsheetApp.openById(m[0]).getName(); }
    catch (e) { out.push('❌ ' + 역할 + ': 열 수 없음 (권한 확인)'); continue; }

    설정.getRange(i + 1, 3).setValue(m[0]);
    out.push('✅ ' + 역할 + ' → ' + name);
  }

  _cache.config = null;
  _cache.ss = {};
  var msg = out.join('\n');
  Logger.log(msg);
  alert_(msg);
  return msg;
}

/**
 * 설정 탭에 빠진 파라미터·별칭을 보충한다.
 * 이미 있는 키는 건드리지 않으므로 여러 번 실행해도 안전하다.
 */
function 설정_파라미터보충() {
  var 설정 = consoleSS_().getSheetByName(CONSOLE.설정);
  if (!설정) throw new Error('[설정] 탭이 없습니다. 설치_1_초기설정() 을 먼저 실행하세요.');

  var 추가할것 = [
    ['파라미터', 'CSV폴더ID', DEFAULT_FOLDER_ID, '주문 CSV 업로드 폴더'],
    ['파라미터', '재고CSV폴더ID', DEFAULT_FOLDER_ID, '카페24 재고 CSV 폴더'],
    ['파라미터', 'CSV처리완료폴더명', '처리완료', '취입 후 CSV 이동 대상'],
    ['파라미터', '폴링주기(분)', 5, '변경 후 설치_3 재실행'],
    ['파라미터', '지시번호접두어', 'PK', '배치번호 형식'],
    ['파라미터', '예약키워드', '예약', '상품명에 이 단어가 있으면 예약상품 (쉼표 구분)'],
    ['파라미터', '재고경고임계치', 3, '재고현황 대시보드 경고 기준'],
    ['파라미터', '추가투입임계(분)', 45, '이 시간 내 완료 예상이면 주문 추가 투입 권고'],
    ['별칭', '예약재고', '예약,확보재고', ''],
    ['별칭', '담당자', '피킹담당자,작업자', '']
  ];

  var v = 설정.getDataRange().getValues();
  var 있음 = {};
  for (var i = 1; i < v.length; i++) {
    있음[String(v[i][0]).trim() + '|' + String(v[i][1]).trim()] = true;
  }

  var 신규 = 추가할것.filter(function (r) { return !있음[r[0] + '|' + r[1]]; });
  if (!신규.length) {
    alert_('추가할 파라미터가 없습니다. 이미 모두 있습니다.');
    return '변경 없음';
  }

  설정.getRange(설정.getLastRow() + 1, 1, 신규.length, 4).setValues(신규);
  _cache.config = null;

  var msg = '파라미터 ' + 신규.length + '건 추가\n\n' +
    신규.map(function (r) { return '  · ' + r[1] + ' = ' + r[2]; }).join('\n');
  alert_(msg);
  return msg;
}

/* ============================================================
 *  설치 2. 시트 정비
 * ============================================================ */

function 설치_2_시트정비() {
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
            '\n\n다음: 설치_3_트리거등록() → D0_대시보드전체갱신()';
  alert_(msg);
  writeOpLog_('설치_2_시트정비', '성공', 결과.join(' / '));
  return msg;
}

/* ============================================================
 *  설치 3. 트리거
 * ============================================================ */

function 설치_3_트리거등록() {
  var ss = consoleSS_();

  ScriptApp.getProjectTriggers().forEach(function (t) {
    var f = t.getHandlerFunction();
    if (['S5_1_결과반영', 'syncAndRefresh', 'syncPickingResults', 'onOpen'].indexOf(f) >= 0) {
      ScriptApp.deleteTrigger(t);
    }
  });

  var 주기 = Number(param_('폴링주기(분)', 5));
  if ([1, 5, 10, 15, 30].indexOf(주기) < 0) 주기 = 5;

  ScriptApp.newTrigger('syncAndRefresh').timeBased().everyMinutes(주기).create();

  var 메뉴 = true;
  try { ScriptApp.newTrigger('onOpen').forSpreadsheet(ss).onOpen().create(); }
  catch (e) { 메뉴 = false; }

  var msg = 주기 + '분 주기 트리거를 등록했습니다.\n' +
    '재고 반영 + 대시보드 3종 갱신이 함께 실행됩니다.\n' +
    (메뉴 ? '"' + ss.getName() + '" 을 열면 메뉴가 표시됩니다.' : '⚠ 메뉴 트리거 등록 실패');

  alert_(msg);
  writeOpLog_('설치_3_트리거등록', '성공', 주기 + '분');
  return msg;
}

/* ============================================================
 *  진단
 * ============================================================ */

function 점검_연결확인() {
  var out = ['콘솔 파일: ' + consoleSS_().getName()];
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
  [CONSOLE.설정, CONSOLE.재고이동로그, CONSOLE.작업로그].forEach(function (n) {
    out.push((ss.getSheetByName(n) ? '✅ ' : '❌ ') + '콘솔 탭: ' + n);
  });

  [[ROLE.주문, '📊 주문현황'], [ROLE.마스터, '📊 재고현황'], [ROLE.헤더, '📊 피킹현황']]
    .forEach(function (p) {
      try {
        out.push((openSS_(p[0]).getSheetByName(p[1]) ? '✅ ' : '❌ ') + '대시보드: ' + p[1]);
      } catch (e) { out.push('❌ 대시보드: ' + p[1]); }
    });

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

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function scanFolder_(folderId) {
  var 결과 = {}, folder;
  try { folder = DriveApp.getFolderById(folderId); } catch (e) { return 결과; }

  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  var 목록 = [];
  while (files.hasNext()) {
    var f = files.next();
    목록.push({ id: f.getId(), name: f.getName() });
  }

  ROLE_HINTS.forEach(function (hint) {
    for (var i = 0; i < 목록.length; i++) {
      var n = normKey_(목록[i].name);
      for (var k = 0; k < hint.keys.length; k++) {
        if (n.indexOf(normKey_(hint.keys[k])) >= 0 && !isTaken_(결과, 목록[i].id)) {
          결과[hint.role] = 목록[i];
          return;
        }
      }
    }
  });

  for (var role in 결과) {
    try { 결과[role].firstSheet = SpreadsheetApp.openById(결과[role].id).getSheets()[0].getName(); }
    catch (e) { 결과[role].firstSheet = ''; }
  }
  return 결과;
}

function isTaken_(결과, id) {
  for (var k in 결과) if (결과[k].id === id) return true;
  return false;
}

function describeScan_(발견) {
  return [ROLE.마스터, ROLE.주문, ROLE.헤더, ROLE.라인].map(function (role) {
    return '  · ' + role + ' → ' + (발견[role] ? 발견[role].name : '❌ 설정_파일URL등록() 실행');
  }).join('\n');
}

function ensureColumns_(sheet, headers, names) {
  var 존재 = {};
  headers.forEach(function (h) { 존재[normKey_(h)] = true; });

  var 추가 = names.filter(function (n) { return !존재[normKey_(n)]; });
  if (!추가.length) return [];

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
    ui.createMenu('📦 피킹 운영')
      .addItem('S1. 카페24 재고 동기화', 'S1_1_카페24재고동기화')
      .addItem('S2. 주문 CSV 취입', 'S2_1_주문CSV취입')
      .addItem('S3. 주문 확정 (재고 검증)', 'S3_1_주문확정')
      .addItem('S4. 피킹지시 생성', 'S4_1_피킹지시생성')
      .addSeparator()
      .addItem('🖨  작업지시서 출력', 'S9_1_작업지시서출력')
      .addItem('S5. 재고 반영 · 취소 처리', 'S5_2_수동반영')
      .addSeparator()
      .addItem('📊 대시보드 전체 갱신', 'D0_대시보드전체갱신')
      .addSubMenu(ui.createMenu('조회')
        .addItem('확정 시뮬레이션', 'S3_9_확정대기조회')
        .addItem('예약대기 현황', 'S8_2_예약대기조회')
        .addItem('보관위치 미지정', 'S1_2_보관위치미지정조회'))
      .addSubMenu(ui.createMenu('관리')
        .addItem('확정 취소', 'S3_2_확정취소')
        .addItem('예약대기 취소', 'S8_3_예약대기취소')
        .addItem('작업자 배정 해제', 'S9_2_배정해제')
        .addSeparator()
        .addItem('연결 상태 확인', '점검_연결확인')
        .addItem('시트 구조 점검', '진단_시트구조')
        .addItem('설정 캐시 초기화', '설정_캐시초기화'))
      .addToUi();
  } catch (err) {
    Logger.log('메뉴 생성 실패: ' + err.message);
  }
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



/** 주문 CSV 폴더 안을 그대로 보여준다 */
function 진단_주문폴더() {
  var raw = param_('CSV폴더ID', DEFAULT_FOLDER_ID);
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
    var 타입 = mime.indexOf('spreadsheet') >= 0 ? '구글시트 (⚠CSV 아님)'
             : mime.indexOf('csv') >= 0 ? 'CSV'
             : mime.indexOf('excel') >= 0 ? '엑셀' : mime;

    var csv조건 = /\.csv$/i.test(name);
    var 재고제외 = name.indexOf('카페24') < 0;
    var 통과 = csv조건 && 재고제외;

    out.push((통과 ? '✅ ' : '❌ ') + name);
    out.push('     ' + 타입 +
             '  |  .csv: ' + (csv조건 ? 'O' : 'X') +
             '  |  주문파일: ' + (재고제외 ? 'O' : 'X (카페24 포함)'));
    n++;
  }

  if (!n) out.push('(폴더가 비어 있습니다)');
  out.push('');
  out.push('※ 하위 폴더는 검색하지 않습니다.');

  Logger.log(out.join('\n'));
}