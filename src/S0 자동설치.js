/**
 * ============================================================
 *  ROOT_FOLDER 기반 자동 설치
 * ============================================================
 *  ROOT 폴더 하나를 기준으로 Drive 폴더, 단일 운영 Spreadsheet,
 *  표준 시트·헤더·설정, validation, trigger, 대시보드를 구성한다.
 *  모든 ensure 단계는 재실행을 전제로 하며 기존 데이터를 초기화하지 않는다.
 * ============================================================
 */

var ROOT_FOLDER_PROPERTY = 'ROOT_FOLDER_ID';
var ROOT_FOLDER_URL_PROPERTY = 'ROOT_FOLDER_URL';
var INSTALL_ROOT_MARKER_PROPERTY = 'POLAR_PENGUIN_ROOT_FOLDER_ID';
var OPERATION_SPREADSHEET_PROPERTY = 'OPERATION_SPREADSHEET_ID';

/**
 * 설치 기준이 될 Google Drive 폴더를 검증하고 Script Property에 저장한다.
 * @param {String} urlOrId Drive 폴더 URL 또는 ID
 * @return {String} 저장된 폴더명과 ID를 포함한 확인 메시지
 * @throws {Error} ID를 추출할 수 없거나 폴더 접근 권한이 없을 때
 */
function setRootFolder(urlOrId) {
  var id = extractDriveId_(urlOrId);
  if (!id) throw new Error('Google Drive 폴더 URL 또는 ID에서 폴더 ID를 찾지 못했습니다.');

  var folder;
  try { folder = DriveApp.getFolderById(id); }
  catch (e) { throw new Error('ROOT 폴더를 열 수 없습니다. ID와 접근 권한을 확인하세요. (' + id + ')'); }

  PropertiesService.getScriptProperties().setProperty(ROOT_FOLDER_PROPERTY, id);
  var msg = 'ROOT folder configured: ' + folder.getName() + ' (' + id + ')';
  Logger.log(msg);
  return msg;
}

/**
 * 신규 설치와 부분 복구를 같은 순서로 수행하는 공식 설치 진입점.
 * 기존 리소스는 재사용하고, 누락된 폴더·파일·시트·헤더·설정만 보충한다.
 * @return {Object} 생성/재사용 리소스, 구성 결과, 경고를 담은 설치 보고서
 * @throws {Error} 단계명이 포함된 오류. 이미 완료된 이전 단계는 삭제하지 않음
 */
function setupSystem() {
  return withLock_(function () {
    var report = { root: null, folders: [], spreadsheets: [], configuration: [], warnings: [] };
    try {
      var root = setupStep_('ROOT folder 확인', function () { return resolveRootFolder_(); });
      report.root = { id: root.getId(), name: root.getName() };

      var folders = setupStep_('Drive 폴더 구성', function () {
        return ensureProjectFolders_(root, report);
      });

      var operationResource = setupStep_('운영 Spreadsheet 구성', function () {
        return ensureOperationSpreadsheet_(root, root.getId(), report);
      });

      persistOperation_(operationResource.ss, root.getId(), folders);
      var resources = setupStep_('운영 탭 구성', function () {
        return ensureOperationalSheets_(operationResource.ss);
      });

      setupStep_('설정 등록', function () {
        ensureSetupConfig_(operationResource.ss, folders, resources, report);
      });

      _cache.consoleSS = operationResource.ss;
      _cache.config = null;
      _cache.ss = {};

      setupStep_('Validation 및 서식', function () { 설치_2_시트정비(true); });
      report.configuration.push('Columns verified');
      report.configuration.push('Validation applied');

      var triggerResult = setupStep_('Trigger 등록', function () { return 설치_3_트리거등록(true); });
      if (String(triggerResult).indexOf('⚠') >= 0) report.warnings.push(triggerResult);
      else report.configuration.push('Triggers registered');

      setupStep_('안내 구성', function () { renderGuideSheet_(operationResource.ss, folders); });
      report.configuration.push('Guide rendered');

      var dashboardResult = setupStep_('Dashboard 갱신', function () {
        return D0_대시보드전체갱신(true);
      });
      if (String(dashboardResult).indexOf('❌') >= 0) report.warnings.push(dashboardResult);
      else report.configuration.push('Dashboard refreshed');

      PropertiesService.getScriptProperties().setProperty(INSTALL_ROOT_MARKER_PROPERTY, root.getId());
      var message = buildSetupReport_(report);
      report.message = message;
      writeOpLog_('setupSystem', '성공', message.replace(/\n/g, ' | '));
      alert_(message);
      return report;
    } catch (e) {
      var failure = 'Polar Penguin setup failed\n\n' + e.message;
      try { writeOpLog_('setupSystem', '실패', e.message); } catch (ignore) { }
      alert_(failure);
      throw e;
    }
  });
}

function setupStep_(name, fn) {
  try { return fn(); }
  catch (e) { throw new Error('[' + name + '] ' + e.message); }
}

function extractDriveId_(value) {
  var match = String(value || '').trim().match(/[-\w]{25,}/);
  return match ? match[0] : '';
}

function resolveRootFolder_() {
  var properties = PropertiesService.getScriptProperties();
  var id = extractDriveId_(properties.getProperty(ROOT_FOLDER_PROPERTY));
  if (!id) id = extractDriveId_(properties.getProperty(ROOT_FOLDER_URL_PROPERTY));
  if (!id) {
    throw new Error('Script Properties에 ROOT_FOLDER_ID 또는 ROOT_FOLDER_URL을 지정하거나 setRootFolder(urlOrId)를 먼저 실행하세요.');
  }
  try { return DriveApp.getFolderById(id); }
  catch (e) { throw new Error('ROOT_FOLDER_ID 폴더를 열 수 없습니다. ID와 접근 권한을 확인하세요. (' + id + ')'); }
}

function ensureProjectFolders_(root, report) {
  var folders = {};
  folders.input = ensureChildFolder_(root, 'Input', 'Input', report);
  folders.processed = ensureChildFolder_(root, 'Processed', 'Processed', report);
  folders.error = ensureChildFolder_(root, 'Error', 'Error', report);
  folders.output = ensureChildFolder_(root, 'Output', 'Output', report);
  folders.backup = ensureChildFolder_(root, 'Backup', 'Backup', report);
  return folders;
}

function ensureChildFolder_(parent, name, label, report) {
  // 같은 이름의 직접 하위 폴더가 있으면 재사용해 setup 재실행 시 중복 생성을 막는다.
  var iterator = parent.getFoldersByName(name);
  var folder;
  var created = false;
  if (iterator.hasNext()) folder = iterator.next();
  else { folder = parent.createFolder(name); created = true; }
  report.folders.push({ label: label, id: folder.getId(), name: name, created: created });
  return folder;
}

function installResourceDefinitions_() {
  return [
    {
      role: ROLE.마스터, sheetName: '상품마스터',
      headers: [COL.상품품목코드, COL.상품명, COL.옵션명, COL.이미지, COL.기본보관위치,
        COL.가용재고, COL.예약재고, COL.불량재고, COL.상품상태, COL.예약상품,
        COL.재고관리, COL.판매가, COL.최종동기화]
    },
    {
      role: ROLE.주문, sheetName: '주문(완료)',
      headers: [COL.주문번호, COL.품목별주문번호, '주문일시', '주문경로', '결제수단',
        '상품번호', COL.상품품목코드, COL.상품명, COL.옵션명, COL.수량,
        '판매가', '상품구매금액', '할인금액', '실결제금액', '주문자명', '주문자 이메일',
        '주문자 휴대전화', '수령인', '수령인 일반전화', '수령인 휴대전화',
        '수령인 우편번호', '수령인 주소', '배송메시지', '배송업체', '송장번호',
        '배송비', '배송유형',
        COL.출고완료, COL.피킹지시번호, COL.주문상태, COL.취소사유,
        COL.취소일시, COL.확정일시, COL.대기사유]
    },
    {
      role: ROLE.헤더, sheetName: '피킹(헤더)',
      headers: [COL.피킹지시번호, COL.주문번호, COL.카트슬롯, COL.품목수,
        COL.총수량, COL.피킹담당자, COL.상태, COL.출력일시]
    },
    {
      role: ROLE.라인, sheetName: '피킹(라인)',
      headers: [COL.순번, COL.보관위치, COL.상품코드, COL.이미지, COL.상품명,
        COL.옵션, COL.필요수량, COL.확인, COL.실제수량, COL.예외사유,
        COL.품목별주문번호, COL.피킹지시번호, COL.담당자, COL.라인상태, COL.처리일시]
    }
  ];
}

function installAliases_() {
  return [
    ['상품품목코드', '상품코드,품목코드,SKU,내부SKU'],
    ['상품코드', '상품품목코드,품목코드,SKU'],
    ['품목별 주문번호', '품목별주문번호,주문상세번호'],
    ['카트 슬롯', '카트 술룻,카트술룻,카트슬롯,슬롯'],
    ['상태', '상태(대기/진행/완료/예외),피킹상태'],
    ['기본보관위치', '보관위치,로케이션,위치'],
    ['보관위치', '기본보관위치,로케이션,위치'],
    ['옵션', '옵션명'], ['옵션명', '옵션'],
    ['수량', '주문수량'], ['필요수량', '주문수량,지시수량'],
    ['예약재고', '예약,확보재고'], ['담당자', '피킹담당자,작업자']
  ];
}

function ensureOperationSpreadsheet_(root, rootId, report) {
  var props = PropertiesService.getScriptProperties();
  var candidate = props.getProperty(INSTALL_ROOT_MARKER_PROPERTY) === rootId
    ? props.getProperty(OPERATION_SPREADSHEET_PROPERTY) : '';
  var resource = ensureSpreadsheetResource_(root, 'Polar Penguin', '📖 안내', candidate);
  report.spreadsheets.push({ role: 'Operational', id: resource.ss.getId(), name: resource.ss.getName(), created: resource.created });
  return resource;
}

function persistOperation_(ss, rootId, folders) {
  var values = {
    ROOT_FOLDER_ID: rootId,
    OPERATION_SPREADSHEET_ID: ss.getId(),
    CONSOLE_SS_ID: ss.getId(),
    POLAR_PENGUIN_ROOT_FOLDER_ID: rootId,
    FOLDER_ID_INPUT: folders.input.getId(),
    FOLDER_ID_PROCESSED: folders.processed.getId(),
    FOLDER_ID_ERROR: folders.error.getId(),
    FOLDER_ID_OUTPUT: folders.output.getId(),
    FOLDER_ID_BACKUP: folders.backup.getId()
  };
  PropertiesService.getScriptProperties().setProperties(values, false);
  _cache.consoleSS = ss;
  _cache.config = null;
  _cache.ss = {};
}

function ensureOperationalSheets_(ss) {
  var resources = {};
  ensureInstallSheet_(ss, '📖 안내', []);
  ensureInstallSheet_(ss, '📊 대시보드', []);
  installResourceDefinitions_().forEach(function (def) {
    resources[def.role] = { ss: ss, sheet: ensureInstallSheet_(ss, def.sheetName, def.headers) };
  });
  ensureInstallSheet_(ss, '예약대기', [COL.주문번호, COL.품목별주문번호, COL.상품품목코드, COL.상품명, COL.수량, COL.대기사유]);
  ensureInstallSheet_(ss, '주문반려', [COL.주문번호, COL.품목별주문번호, COL.상품품목코드, COL.상품명, COL.수량, COL.취소사유, COL.취소일시]);
  ensureInstallSheet_(ss, CONSOLE.설정, ['구분', '키', '값', '비고']);
  ensureInstallSheet_(ss, CONSOLE.재고이동로그, [
    '시각', '구분', COL.피킹지시번호, COL.주문번호, COL.품목별주문번호,
    COL.상품코드, '변동량', '변동 후 재고', COL.담당자, '사유'
  ]);
  ensureInstallSheet_(ss, CONSOLE.작업로그, ['시각', '함수', '결과', '메시지', '실행계정']);
  ensureInstallSheet_(ss, CONSOLE.입력처리로그, INPUT_LOG_HEADERS);
  return resources;
}

function ensureSpreadsheetResource_(folder, fileName, sheetName, candidateId) {
  var ss = openSpreadsheetIfValid_(candidateId);
  var created = false;
  if (!ss) ss = findSpreadsheetByName_(folder, fileName);
  if (!ss) {
    ss = SpreadsheetApp.create(fileName);
    DriveApp.getFileById(ss.getId()).moveTo(folder);
    created = true;
  }
  ensureInstallSheet_(ss, sheetName, []);
  return { ss: ss, created: created };
}

function openSpreadsheetIfValid_(raw) {
  var id = extractDriveId_(raw);
  if (!id) return null;
  try { return SpreadsheetApp.openById(id); } catch (e) { return null; }
}

function validSpreadsheetId_(raw) { return !!openSpreadsheetIfValid_(raw); }

function findSpreadsheetByName_(folder, name) {
  var files = folder.getFilesByName(name);
  while (files.hasNext()) {
    var file = files.next();
    if (file.getMimeType() !== MimeType.GOOGLE_SHEETS) continue;
    var ss = openSpreadsheetIfValid_(file.getId());
    if (ss) return ss;
  }
  return null;
}

function ensureInstallSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    var sheets = ss.getSheets();
    if (sheets.length === 1 && sheets[0].getName() === 'Sheet1' &&
        sheets[0].getLastRow() === 0 && sheets[0].getLastColumn() === 0) {
      sheet = sheets[0].setName(name);
    } else {
      sheet = ss.insertSheet(name);
    }
  }

  var current = sheet.getLastColumn() > 0
    ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
  // 데이터가 있는 시트는 별칭도 기존 컬럼으로 인정해 중복 컬럼을 만들지 않는다.
  var missing = missingInstallHeaders_(current, headers, sheet.getLastRow() > 1);
  if (missing.length) {
    var start = lastInstallHeaderColumn_(current) + 1;
    var requiredLast = start + missing.length - 1;
    if (sheet.getMaxColumns() < requiredLast) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredLast - sheet.getMaxColumns());
    }
    sheet.getRange(1, start, 1, missing.length).setValues([missing]);
  }
  if (headers.length) {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length))
      .setFontWeight('bold').setBackground('E8EDF3');
  }
  return sheet;
}

/** 안내 탭은 시스템이 관리하는 설명 영역이며 setup 재실행 때 현재 구조로 다시 그린다. */
function renderGuideSheet_(ss, folders) {
  var sh = ensureInstallSheet_(ss, '📖 안내', []);
  sh.clearContents();
  sh.clearFormats();
  if (sh.getMaxColumns() < 8) sh.insertColumnsAfter(sh.getMaxColumns(), 8 - sh.getMaxColumns());

  sh.getRange(1, 1, 1, 8).merge().setValue('📖  Polar Penguin 운영 안내')
    .setFontSize(18).setFontWeight('bold').setFontColor('FFFFFF').setBackground(DASHCOLOR.제목);
  sh.getRange(3, 1, 1, 8).merge().setValue(
    'Input 파일 업로드  →  자동 주문/재고 판별  →  검증  →  재고 또는 주문 처리  →  주문 확정  →  피킹지시 생성  →  PDF 자동 생성  →  현장 피킹  →  O/X 입력  →  결과 반영'
  ).setWrap(true).setBackground(DASHCOLOR.카드).setFontWeight('bold');

  var rows = [
    ['폴더', '용도'],
    ['Input', '주문·재고 파일을 함께 올리는 유일한 입력 폴더'],
    ['Processed', '정상 처리된 원본이 날짜별로 이동'],
    ['Error', '검증/처리 실패 원본이 날짜별로 이동'],
    ['Output', 'Output/YYYY-MM-DD/<피킹지시번호>.pdf'],
    ['Backup', '보관 및 향후 백업용'],
    ['', ''],
    ['핵심 용어', '설명'],
    ['주문상태', '접수 → 확정 또는 예약대기/취소 → 출고완료'],
    ['가용재고', '새 주문에 배정할 수 있는 수량'],
    ['예약재고', '확정 주문에 확보되어 출고를 기다리는 수량'],
    ['예약상품', 'Y인 동안 주문 전체를 예약대기로 유지'],
    ['O / X', 'O=정상 피킹, X=예외. X이면 예외사유를 선택하고 주문 전체 취소 규칙 적용'],
    ['대시보드', '주문·재고·피킹·예약·운영 현황과 권고를 한 화면에서 확인'],
    ['', ''],
    ['운영 방법', '상단 「📦 Polar Penguin」 메뉴에서 Input 처리, 주문 확정, 피킹지시 생성, 결과 반영, 작업지시서 출력, 대시보드 갱신을 실행하세요.']
  ];
  sh.getRange(5, 1, rows.length, 2).setValues(rows).setWrap(true);
  sh.getRange(5, 1, 1, 2).setFontWeight('bold').setBackground(DASHCOLOR.헤더);
  sh.getRange(12, 1, 1, 2).setFontWeight('bold').setBackground(DASHCOLOR.헤더);
  sh.getRange(20, 1, 1, 2).setFontWeight('bold').setBackground(DASHCOLOR.헤더);
  sh.setColumnWidth(1, 150);
  sh.setColumnWidth(2, 620);
  sh.setFrozenRows(1);
  sh.setHiddenGridlines(true);
  return sh;
}

function missingInstallHeaders_(current, required, acceptAliases) {
  var present = {};
  (current || []).forEach(function (header) {
    var key = normKey_(header);
    if (key) present[key] = true;
  });
  var aliases = installAliasMap_();
  return (required || []).filter(function (header) {
    if (present[normKey_(header)]) return false;
    if (!acceptAliases) return true;
    var list = aliases[header] || [];
    return !list.some(function (alias) { return present[normKey_(alias)]; });
  });
}

function installAliasMap_() {
  var map = {};
  installAliases_().forEach(function (entry) {
    map[entry[0]] = entry[1].split(',').map(function (value) { return value.trim(); });
  });
  return map;
}

function lastInstallHeaderColumn_(headers) {
  for (var i = (headers || []).length - 1; i >= 0; i--) {
    if (String(headers[i] || '').trim()) return i + 1;
  }
  return 0;
}

function readSetupConfig_(sheet) {
  var config = { 파일ID: {}, 시트명: {}, 파라미터: {}, 별칭: {} };
  if (!sheet || sheet.getLastRow() < 2) return config;
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  values.forEach(function (row) {
    var section = String(row[0] || '').trim();
    var key = String(row[1] || '').trim();
    if (section && key && config[section]) config[section][key] = row[2];
  });
  return config;
}

function ensureSetupConfig_(consoleSs, folders, resources, report) {
  var sheet = consoleSs.getSheetByName(CONSOLE.설정);
  var managed = [];
  installResourceDefinitions_().forEach(function (def) {
    managed.push(['파일ID', def.role, resources[def.role].ss.getId(), '단일 운영 Spreadsheet',
      function (value) { return extractDriveId_(value) === resources[def.role].ss.getId(); }]);
    managed.push(['시트명', def.role, resources[def.role].sheet.getName(), '표준 업무 시트',
      function (value) { return !!resources[def.role].ss.getSheetByName(String(value || '')); }]);
  });
  var exactFolder = function (folder) {
    return function (value) { return extractDriveId_(value) === folder.getId() && validFolderId_(value); };
  };
  managed.push(['파라미터', '통합Input폴더ID', folders.input.getId(), '유일한 입력 폴더', exactFolder(folders.input)]);
  managed.push(['파라미터', 'Processed폴더ID', folders.processed.getId(), '정상 처리 원본', exactFolder(folders.processed)]);
  managed.push(['파라미터', 'Error폴더ID', folders.error.getId(), '처리 실패 원본', exactFolder(folders.error)]);
  managed.push(['파라미터', 'Output폴더ID', folders.output.getId(), '피킹 PDF 출력', exactFolder(folders.output)]);
  managed.push(['파라미터', 'Backup폴더ID', folders.backup.getId(), '백업 폴더', exactFolder(folders.backup)]);
  // 기존 S1/S2 수동 진입점도 통합 Input을 보도록 호환 키를 유지한다.
  managed.push(['파라미터', 'CSV폴더ID', folders.input.getId(), '통합 Input (호환)', exactFolder(folders.input)]);
  managed.push(['파라미터', '재고CSV폴더ID', folders.input.getId(), '통합 Input (호환)', exactFolder(folders.input)]);

  var defaults = [
    ['파라미터', 'CSV처리완료폴더명', '처리완료', '기존 S2 단독 실행용 호환 설정'],
    ['파라미터', '폴링주기(분)', 5, '변경 후 setupSystem 재실행'],
    ['파라미터', '지시번호접두어', 'PK', '배치번호 형식'],
    ['파라미터', '예약키워드', '예약', '상품명 예약 판정 문자열 (쉼표 구분)'],
    ['파라미터', '재고경고임계치', 3, '재고현황 대시보드 경고 기준'],
    ['파라미터', '추가투입임계(분)', 45, '주문 추가 투입 권고 기준'],
    ['파라미터', '알림이메일', '', '입력 처리 실패 알림 수신자']
  ];
  installAliases_().forEach(function (entry) { defaults.push(['별칭', entry[0], entry[1], '']); });

  var result = mergeSetupConfig_(sheet.getDataRange().getValues(), managed, defaults);
  result.updates.forEach(function (update) { sheet.getRange(update.row, 3).setValue(update.value); });
  if (result.additions.length) {
    sheet.getRange(Math.max(sheet.getLastRow() + 1, 2), 1, result.additions.length, 4).setValues(result.additions);
  }
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 4);
  _cache.config = null;
  _cache.ss = {};
  report.configuration.push('Single operational Spreadsheet registered');
  report.configuration.push('Folder IDs registered');
}

function mergeSetupConfig_(values, managed, defaults) {
  // managed 값은 없거나 유효하지 않을 때만 복구하고, defaults는 키가 없을 때만 추가한다.
  // 따라서 폴링 주기 같은 사용자 설정값은 setup 재실행으로 덮어쓰지 않는다.
  var index = {};
  (values || []).forEach(function (row, offset) {
    var section = String(row[0] || '').trim();
    var key = String(row[1] || '').trim();
    if (section && key && index[section + '|' + key] === undefined) {
      index[section + '|' + key] = { row: offset + 1, value: row[2] };
    }
  });

  var updates = [];
  var additions = [];
  (managed || []).forEach(function (row) {
    var item = index[row[0] + '|' + row[1]];
    if (!item) { additions.push(row.slice(0, 4)); return; }
    var validator = row[4];
    if (item.value === '' || item.value === null || item.value === undefined ||
        (validator && !validator(item.value))) {
      updates.push({ row: item.row, value: row[2] });
    }
  });
  (defaults || []).forEach(function (row) {
    if (!index[row[0] + '|' + row[1]]) additions.push(row);
  });
  return { additions: additions, updates: updates };
}

function validFolderId_(raw) {
  var id = extractDriveId_(raw);
  if (!id) return false;
  try { DriveApp.getFolderById(id).getName(); return true; } catch (e) { return false; }
}

function buildSetupReport_(report) {
  var lines = ['Polar Penguin setup complete', '', 'ROOT',
    '✓ ' + report.root.name + ' (' + report.root.id + ')', '', 'Folders'];
  report.folders.forEach(function (item) {
    lines.push('✓ ' + item.label + ' (' + (item.created ? 'created' : 'reused') + ')');
  });
  lines.push('', 'Spreadsheets');
  report.spreadsheets.forEach(function (item) {
    lines.push('✓ ' + item.role + ' — ' + item.name + ' (' + (item.created ? 'created' : 'reused') + ')');
  });
  lines.push('', 'Configuration');
  report.configuration.forEach(function (item) { lines.push('✓ ' + item); });
  if (report.warnings.length) {
    lines.push('', 'Warnings');
    report.warnings.forEach(function (item) { lines.push('⚠ ' + item); });
  }
  return lines.join('\n');
}
