/**
 * ============================================================
 *  ROOT_FOLDER 기반 자동 설치
 * ============================================================
 *  외부 공개 함수
 *    setRootFolder(urlOrId)
 *    setupSystem()
 * ============================================================
 */

var ROOT_FOLDER_PROPERTY = 'ROOT_FOLDER_ID';
var ROOT_FOLDER_URL_PROPERTY = 'ROOT_FOLDER_URL';
var INSTALL_ROOT_MARKER_PROPERTY = 'POLAR_PENGUIN_ROOT_FOLDER_ID';

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

function setupSystem() {
  return withLock_(function () {
    var report = { root: null, folders: [], spreadsheets: [], configuration: [], warnings: [] };
    try {
      var root = setupStep_('ROOT folder 확인', function () { return resolveRootFolder_(); });
      report.root = { id: root.getId(), name: root.getName() };

      var folders = setupStep_('Drive 폴더 구성', function () {
        return ensureProjectFolders_(root, report);
      });

      var consoleResource = setupStep_('Console Spreadsheet 구성', function () {
        return ensureConsoleSpreadsheet_(folders.console, root.getId(), report);
      });

      persistConsole_(consoleResource.ss, root.getId(), folders);
      ensureConsoleSheets_(consoleResource.ss);
      var existingConfig = readSetupConfig_(consoleResource.ss.getSheetByName(CONSOLE.설정));

      var resources = setupStep_('업무 Spreadsheet 구성', function () {
        return ensureProjectSpreadsheets_(folders, existingConfig, report);
      });

      setupStep_('설정 등록', function () {
        ensureSetupConfig_(consoleResource.ss, folders, resources, report);
      });

      _cache.consoleSS = consoleResource.ss;
      _cache.config = null;
      _cache.ss = {};

      setupStep_('Validation 및 서식', function () { 설치_2_시트정비(true); });
      report.configuration.push('Columns verified');
      report.configuration.push('Validation applied');

      var triggerResult = setupStep_('Trigger 등록', function () { return 설치_3_트리거등록(true); });
      if (String(triggerResult).indexOf('⚠') >= 0) report.warnings.push(triggerResult);
      else report.configuration.push('Triggers registered');

      var dashboardResult = setupStep_('Dashboard 갱신', function () {
        return D0_대시보드전체갱신(true);
      });
      if (String(dashboardResult).indexOf('❌') >= 0) report.warnings.push(dashboardResult);
      else report.configuration.push('Dashboards refreshed');

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
  folders.console = ensureChildFolder_(root, '01 Console', 'Console', report);
  folders.master = ensureChildFolder_(root, '02 Master', 'Master', report);
  folders.orders = ensureChildFolder_(root, '03 Orders', 'Orders', report);
  folders.picking = ensureChildFolder_(root, '04 Picking', 'Picking', report);
  folders.input = ensureChildFolder_(root, 'Input', 'Input', report);
  folders.processed = ensureChildFolder_(root, 'Processed', 'Processed', report);
  folders.error = ensureChildFolder_(root, 'Error', 'Error', report);
  folders.output = ensureChildFolder_(root, 'Output', 'Output', report);
  folders.backup = ensureChildFolder_(root, 'Backup', 'Backup', report);
  return folders;
}

function ensureChildFolder_(parent, name, label, report) {
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
      role: ROLE.마스터, property: 'SPREADSHEET_ID_MASTER', folder: 'master',
      fileName: '상품마스터', sheetName: '상품마스터',
      headers: [COL.상품품목코드, COL.상품명, COL.옵션명, COL.이미지, COL.기본보관위치,
        COL.가용재고, COL.예약재고, COL.불량재고, COL.상품상태, COL.예약상품,
        COL.재고관리, COL.판매가, COL.최종동기화]
    },
    {
      role: ROLE.주문, property: 'SPREADSHEET_ID_ORDERS', folder: 'orders',
      fileName: '주문완료', sheetName: '주문(완료)',
      headers: [COL.주문번호, COL.품목별주문번호, COL.상품품목코드, COL.수량,
        COL.출고완료, COL.피킹지시번호, COL.주문상태, COL.취소사유,
        COL.취소일시, COL.확정일시, COL.대기사유]
    },
    {
      role: ROLE.헤더, property: 'SPREADSHEET_ID_PICKING_HEADER', folder: 'picking',
      fileName: '피킹헤더', sheetName: '피킹(헤더)',
      headers: [COL.피킹지시번호, COL.주문번호, COL.카트슬롯, COL.품목수,
        COL.총수량, COL.피킹담당자, COL.상태, COL.출력일시]
    },
    {
      role: ROLE.라인, property: 'SPREADSHEET_ID_PICKING_LINE', folder: 'picking',
      fileName: '피킹라인', sheetName: '피킹(라인)',
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

function ensureConsoleSpreadsheet_(folder, rootId, report) {
  var props = PropertiesService.getScriptProperties();
  var candidate = props.getProperty(INSTALL_ROOT_MARKER_PROPERTY) === rootId
    ? props.getProperty('CONSOLE_SS_ID') : '';
  var resource = ensureSpreadsheetResource_(folder, 'Polar Penguin Console', '설정', candidate);
  report.spreadsheets.push({ role: 'Console', id: resource.ss.getId(), name: resource.ss.getName(), created: resource.created });
  return resource;
}

function persistConsole_(ss, rootId, folders) {
  var values = {
    ROOT_FOLDER_ID: rootId,
    CONSOLE_SS_ID: ss.getId(),
    POLAR_PENGUIN_ROOT_FOLDER_ID: rootId,
    FOLDER_ID_CONSOLE: folders.console.getId(),
    FOLDER_ID_MASTER: folders.master.getId(),
    FOLDER_ID_ORDERS: folders.orders.getId(),
    FOLDER_ID_PICKING: folders.picking.getId(),
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

function ensureConsoleSheets_(ss) {
  ensureInstallSheet_(ss, CONSOLE.설정, ['구분', '키', '값', '비고']);
  ensureInstallSheet_(ss, CONSOLE.재고이동로그, [
    '시각', '구분', COL.피킹지시번호, COL.주문번호, COL.품목별주문번호,
    COL.상품코드, '변동량', '변동 후 재고', COL.담당자, '사유'
  ]);
  ensureInstallSheet_(ss, CONSOLE.작업로그, ['시각', '함수', '결과', '메시지', '실행계정']);
  ensureInstallSheet_(ss, CONSOLE.입력처리로그, INPUT_LOG_HEADERS);
}

function ensureProjectSpreadsheets_(folders, config, report) {
  var props = PropertiesService.getScriptProperties();
  var resources = {};

  installResourceDefinitions_().forEach(function (def) {
    var configuredId = config.파일ID[def.role] || '';
    var candidate = validSpreadsheetId_(configuredId) ? configuredId : '';
    var resource = ensureSpreadsheetResource_(folders[def.folder], def.fileName, def.sheetName, candidate);
    resource.sheet = ensureInstallSheet_(resource.ss, def.sheetName, def.headers);
    resources[def.role] = resource;
    props.setProperty(def.property, resource.ss.getId());
    report.spreadsheets.push({ role: def.role, id: resource.ss.getId(), name: resource.ss.getName(), created: resource.created });
  });
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
    if (sheets.length === 1 && sheets[0].getLastRow() === 0 && sheets[0].getLastColumn() === 0) {
      sheet = sheets[0].setName(name);
    } else {
      sheet = ss.insertSheet(name);
    }
  }

  var current = sheet.getLastColumn() > 0
    ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] : [];
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
    managed.push(['파일ID', def.role, resources[def.role].ss.getId(), def.fileName,
      function (value) { return validSpreadsheetId_(value); }]);
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
  report.configuration.push('File IDs registered');
  report.configuration.push('Folder IDs registered');
}

function mergeSetupConfig_(values, managed, defaults) {
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
