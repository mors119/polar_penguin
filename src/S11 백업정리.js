/**
 * 전체 운영 Spreadsheet 백업을 안전 경계로 삼는 보존기간 정리 서비스.
 * 처리주문아카이브는 중복 방지용 최소 색인이며 전체 복구본은 Backup의 원본 사본이다.
 */

function 백업_및_정리() {
  var ui = SpreadsheetApp.getUi();
  var message = '현재 운영 Spreadsheet를 전체 백업한 후,\n' +
    '보존기간이 지난 완료/취소 주문 및\nSuccess / Error / Output 파일을 정리합니다.\n\n' +
    '백업에 실패하면 아무 데이터도 삭제하지 않습니다.\n\n계속하시겠습니까?';
  if (ui.alert('Polar Penguin 백업 및 정리', message, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
    return { cancelled: true };
  }
  var result = runBackupAndCleanup_({ silent: false });
  alert_(result.aborted ? '백업에 실패하여 정리를 중단했습니다. 삭제된 데이터는 없습니다.' :
    '백업 및 정리가 완료되었습니다.\n\n' + buildMaintenanceSummary_(result));
  return result;
}

function runBackupAndCleanup_(options) {
  options = options || {};
  return withLock_(function () {
    var now = options.now || new Date();
    var retentionDays = Math.max(1, Number(options.retentionDays || param_('정리보존일수', 30)) || 30);
    var cutoff = new Date(now.getTime() - retentionDays * 24 * 60 * 60 * 1000);
    var result = {
      backup: null,
      archive: { inserted: 0, skippedExisting: 0, failed: 0 },
      sheets: { ordersDeleted: 0, pickingHeadersDeleted: 0, pickingLinesDeleted: 0 },
      files: { successDeleted: 0, errorDeleted: 0, outputDeleted: 0, foldersDeleted: 0 },
      retentionDays: retentionDays, cutoff: cutoff, errors: [], emailSent: false, aborted: false
    };
    writeOpLog_('백업_및_정리', '시작', '보존 ' + retentionDays + '일 / cutoff ' + formatMaintenanceDate_(cutoff));

    try {
      maintenancePreflight_(options);
      result.backup = createBackupSnapshot_({ now: now, spreadsheet: options.spreadsheet,
        backupRoot: options.backupRoot });
      writeOpLog_('백업_및_정리', '백업성공', result.backup.fileId + ' / ' + result.backup.fileName);
    } catch (backupError) {
      result.aborted = true;
      result.errors.push('백업: ' + backupError.message);
      writeOpLog_('백업_및_정리', '실패', '백업 실패 / 정리 미실행 / ' + backupError.message);
      var failureNotice = sendSystemNotification_('ERROR', '백업 실패 - 정리 중단', {
        결과: '백업에 실패하여 정리를 중단했습니다. 삭제된 데이터는 없습니다.', 오류: backupError.message
      });
      result.emailSent = !!failureNotice.sent;
      return result;
    }

    var cleaned = null;
    try {
      cleaned = archiveAndCleanupFinalizedOrders_(cutoff, now);
      result.archive = cleaned.archive;
      result.sheets.ordersDeleted = cleaned.ordersDeleted;
      writeOpLog_('백업_및_정리', '아카이브', '삽입 ' + result.archive.inserted + ' / 주문행 삭제 ' + cleaned.ordersDeleted);
    } catch (orderError) {
      if (orderError.archiveResult) result.archive = orderError.archiveResult;
      if (orderError.deletedRows) result.sheets.ordersDeleted = orderError.deletedRows;
      result.errors.push('주문/아카이브: ' + orderError.message);
      writeOpLog_('백업_및_정리', '부분실패', '주문/아카이브 / ' + orderError.message);
    }

    var picking = { headerRowsDeleted: 0, lineRowsDeleted: 0, removedInstructions: {} };
    if (cleaned) {
      try {
        picking = cleanupPickingHistory_(cleaned.deletedInstructions);
        result.sheets.pickingHeadersDeleted = picking.headerRowsDeleted;
        result.sheets.pickingLinesDeleted = picking.lineRowsDeleted;
        writeOpLog_('백업_및_정리', '시트정리', '헤더 ' + picking.headerRowsDeleted + ' / 라인 ' + picking.lineRowsDeleted);
      } catch (pickingError) {
        if (pickingError.headerRowsDeleted) result.sheets.pickingHeadersDeleted = pickingError.headerRowsDeleted;
        if (pickingError.lineRowsDeleted) result.sheets.pickingLinesDeleted = pickingError.lineRowsDeleted;
        result.errors.push('피킹: ' + pickingError.message);
        writeOpLog_('백업_및_정리', '부분실패', '피킹 / ' + pickingError.message);
      }
    }

    [
      { label: 'Success', key: 'Success폴더ID', field: 'successDeleted', options: {} },
      { label: 'Error', key: 'Error폴더ID', field: 'errorDeleted', options: {} },
      { label: 'Output', key: 'Output폴더ID', field: 'outputDeleted',
        options: { allowedNames: maintenancePdfNames_(picking.removedInstructions) } }
    ].forEach(function (stage) {
      try {
        var cleanedFiles = cleanupFolderByRetention_(inputFolder_(stage.key), cutoff, stage.options);
        result.files[stage.field] = cleanedFiles.filesDeleted;
        result.files.foldersDeleted += cleanedFiles.foldersDeleted;
      } catch (fileError) {
        var partial = fileError.cleanupResult || { filesDeleted: 0, foldersDeleted: 0 };
        result.files[stage.field] = partial.filesDeleted;
        result.files.foldersDeleted += partial.foldersDeleted;
        result.errors.push(stage.label + ' Drive: ' + fileError.message);
        writeOpLog_('백업_및_정리', '부분실패', stage.label + ' Drive / ' + fileError.message);
      }
    });
    writeOpLog_('백업_및_정리', '파일정리', 'Success ' + result.files.successDeleted + ' / Error ' +
      result.files.errorDeleted + ' / Output ' + result.files.outputDeleted + ' / 폴더 ' + result.files.foldersDeleted);

    try { D0_대시보드전체갱신(true); }
    catch (dashboardError) { result.errors.push('대시보드: ' + dashboardError.message); }

    var level = result.errors.length ? 'WARNING' : 'INFO';
    var notification = sendSystemNotification_(level,
      '백업 및 정리 완료 - ' + Utilities.formatDate(now, tz_(), 'yyyy-MM-dd'), buildMaintenanceSummary_(result));
    result.emailSent = !!notification.sent;
    if (!notification.sent) result.notificationError = notification.reason || '알림 미발송';
    writeOpLog_('백업_및_정리', result.errors.length ? '부분성공' : '성공', buildMaintenanceSummary_(result));
    return result;
  });
}

/** 백업 원본의 필수 구조와 정리 대상 폴더 접근성을 쓰기 전에 검증한다. */
function maintenancePreflight_(options) {
  var ss = (options && options.spreadsheet) || consoleSS_();
  var required = [ROLE.마스터, ROLE.주문, ROLE.라인, ROLE.헤더, CONSOLE.재고이동로그,
    CONSOLE.작업로그, CONSOLE.입력처리로그, CONSOLE.설정, CONSOLE.처리주문아카이브];
  var missing = required.filter(function (name) { return !ss.getSheetByName(name); });
  if (missing.length) throw new Error('필수 탭 누락: ' + missing.join(', '));
  if (!(options && options.backupRoot)) inputFolder_('백업폴더ID');
  ['Success폴더ID', 'Error폴더ID', 'Output폴더ID'].forEach(function (key) {
    if (!(options && options.skipFolderPreflight)) inputFolder_(key);
  });
  return true;
}

function createBackupSnapshot_(options) {
  options = options || {};
  var now = options.now || new Date(), ss = options.spreadsheet || consoleSS_();
  var root = options.backupRoot || inputFolder_('백업폴더ID');
  var month = Utilities.formatDate(now, tz_(), 'yyyy-MM');
  var monthFolder = ensureMaintenanceMonthFolder_(root, month);
  var fileName = 'Polar Penguin Backup ' + Utilities.formatDate(now, tz_(), 'yyyy-MM-dd HHmmss');
  var source = DriveApp.getFileById(ss.getId());
  var copy = source.makeCopy(fileName, monthFolder);
  var fileId = copy && copy.getId ? copy.getId() : '';
  if (!fileId || fileId === ss.getId()) throw new Error('백업 파일 ID 검증 실패');

  var verified = DriveApp.getFileById(fileId);
  if (!verified || verified.getId() !== fileId) throw new Error('백업 파일 접근 검증 실패');
  var inside = false, parents = verified.getParents();
  while (parents.hasNext()) { if (parents.next().getId() === monthFolder.getId()) inside = true; }
  if (!inside) throw new Error('백업 파일이 Backup/' + month + ' 폴더에 없습니다.');
  var copiedSpreadsheet = SpreadsheetApp.openById(fileId);
  if (!copiedSpreadsheet || copiedSpreadsheet.getId() !== fileId) throw new Error('백업 Spreadsheet 열기 검증 실패');
  var copiedRequired = [ROLE.마스터, ROLE.주문, ROLE.라인, ROLE.헤더, CONSOLE.재고이동로그,
    CONSOLE.작업로그, CONSOLE.입력처리로그, CONSOLE.설정, CONSOLE.처리주문아카이브];
  var copiedMissing = copiedRequired.filter(function (name) { return !copiedSpreadsheet.getSheetByName(name); });
  if (copiedMissing.length) throw new Error('백업 필수 탭 검증 실패: ' + copiedMissing.join(', '));

  var properties = PropertiesService.getScriptProperties();
  properties.setProperties({ 최근백업일시: now.toISOString(), 최근백업파일ID: fileId }, false);
  return { success: true, fileId: fileId, fileName: fileName,
    url: verified.getUrl ? verified.getUrl() : 'https://docs.google.com/spreadsheets/d/' + fileId,
    monthFolderId: monthFolder.getId(), sourceSpreadsheetId: ss.getId() };
}

function ensureMaintenanceMonthFolder_(root, name) {
  var found = root.getFoldersByName(name);
  return found.hasNext() ? found.next() : root.createFolder(name);
}

function collectEligibleFinalizedOrders_(cutoff) {
  var table = readTable_(ROLE.주문);
  var C = { orderNo: col_(table, COL.주문번호, true), itemNo: col_(table, COL.품목별주문번호, true),
    sku: col_(table, COL.상품품목코드, true), qty: col_(table, COL.수량, true), state: col_(table, COL.주문상태, true),
    completedAt: col_(table, COL.확정일시, false), cancelledAt: col_(table, COL.취소일시, false),
    instruction: col_(table, COL.피킹지시번호, false) };
  var grouped = {};
  table.rows.forEach(function (row, index) {
    var orderNo = toStr_(row[C.orderNo]); if (!orderNo) return;
    if (!grouped[orderNo]) grouped[orderNo] = [];
    grouped[orderNo].push({ row: row, index: index, sheetRow: index + 2 });
  });
  var eligible = [];
  Object.keys(grouped).forEach(function (orderNo) {
    var items = grouped[orderNo], state = toStr_(items[0].row[C.state]);
    if ([ENUM.주문상태.출고완료, ENUM.주문상태.취소].indexOf(state) < 0) return;
    var valid = items.every(function (item) {
      if (toStr_(item.row[C.state]) !== state || !toStr_(item.row[C.itemNo])) return false;
      var value = state === ENUM.주문상태.취소 ? item.row[C.cancelledAt] : item.row[C.completedAt];
      var date = maintenanceDate_(value);
      item.finalDate = date;
      return date && date.getTime() < cutoff.getTime();
    });
    if (valid) eligible.push({ orderNo: orderNo, state: state, items: items, columns: C });
  });
  return { table: table, groups: eligible };
}

function archiveAndCleanupFinalizedOrders_(cutoff, now) {
  var collected = collectEligibleFinalizedOrders_(cutoff), archiveSheet = consoleSS_().getSheetByName(CONSOLE.처리주문아카이브);
  if (!archiveSheet) throw new Error(CONSOLE.처리주문아카이브 + ' 탭이 없습니다.');
  var existing = getArchivedItemOrderKeys_(), additions = [], skipped = 0;
  collected.groups.forEach(function (group) {
    group.items.forEach(function (item) {
      var C = group.columns, key = toStr_(item.row[C.itemNo]);
      if (existing[key]) { skipped++; return; }
      additions.push([now, key, group.orderNo, toStr_(item.row[C.sku]), group.state, item.finalDate,
        C.instruction >= 0 ? toStr_(item.row[C.instruction]) : '', toNum_(item.row[C.qty])]);
      existing[key] = true;
    });
  });
  if (additions.length) appendRows_(archiveSheet, additions, [1, 2, 3, 6]);
  var persisted = getArchivedItemOrderKeys_(), deleteRows = [], deletedInstructions = {};
  collected.groups.forEach(function (group) {
    if (!group.items.every(function (item) { return !!persisted[toStr_(item.row[group.columns.itemNo])]; })) return;
    group.items.forEach(function (item) {
      deleteRows.push(item.sheetRow);
      var instruction = group.columns.instruction >= 0 ? toStr_(item.row[group.columns.instruction]) : '';
      if (instruction) deletedInstructions[instruction] = true;
    });
  });
  var archiveResult = { inserted: additions.length, skippedExisting: skipped,
    failed: collected.groups.reduce(function (sum, group) { return sum + group.items.length; }, 0) - deleteRows.length };
  try { deleteSheetRowsByIndexes_(collected.table.sheet, deleteRows); }
  catch (e) { e.archiveResult = archiveResult; throw e; }
  var retryableInstructions = getArchivedInactiveInstructions_();
  Object.keys(retryableInstructions).forEach(function (no) { deletedInstructions[no] = true; });
  return { archive: archiveResult,
    ordersDeleted: deleteRows.length, deletedInstructions: deletedInstructions };
}

function getArchivedItemOrderKeys_() {
  var sheet = consoleSS_().getSheetByName(CONSOLE.처리주문아카이브), keys = {};
  if (!sheet || sheet.getLastRow() < 2) return keys;
  var values = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
  values.forEach(function (row) { var key = toStr_(row[0]); if (key) keys[key] = true; });
  return keys;
}

/** 이전 실행에서 주문 정리 후 피킹 정리가 실패했어도 다음 실행이 안전하게 이어받는다. */
function getArchivedInactiveInstructions_() {
  var orders = readTable_(ROLE.주문), active = {}, orderInstruction = col_(orders, COL.피킹지시번호, false);
  if (orderInstruction >= 0) orders.rows.forEach(function (row) {
    var instruction = toStr_(row[orderInstruction]); if (instruction) active[instruction] = true;
  });
  var sheet = consoleSS_().getSheetByName(CONSOLE.처리주문아카이브), result = {};
  if (!sheet || sheet.getLastRow() < 2) return result;
  var values = sheet.getDataRange().getValues(), headers = values[0];
  var instructionColumn = headers.indexOf('피킹지시번호');
  if (instructionColumn < 0) return result;
  values.slice(1).forEach(function (row) {
    var instruction = toStr_(row[instructionColumn]);
    if (instruction && !active[instruction]) result[instruction] = true;
  });
  return result;
}

function cleanupPickingHistory_(candidateInstructions) {
  candidateInstructions = candidateInstructions || {};
  var orders = readTable_(ROLE.주문), active = {}, oInstruction = col_(orders, COL.피킹지시번호, false);
  if (oInstruction >= 0) orders.rows.forEach(function (row) { var no = toStr_(row[oInstruction]); if (no) active[no] = true; });
  var headers = readTable_(ROLE.헤더), hInstruction = col_(headers, COL.피킹지시번호, true), hState = col_(headers, COL.상태, true);
  var lines = readTable_(ROLE.라인), lInstruction = col_(lines, COL.피킹지시번호, true), lState = col_(lines, COL.라인상태, true);
  var headerGroups = {}, lineGroups = {};
  headers.rows.forEach(function (row, index) {
    var no = toStr_(row[hInstruction]); if (!headerGroups[no]) headerGroups[no] = [];
    headerGroups[no].push({ row: row, sheetRow: index + 2 });
  });
  lines.rows.forEach(function (row, index) {
    var no = toStr_(row[lInstruction]); if (!lineGroups[no]) lineGroups[no] = [];
    lineGroups[no].push({ row: row, sheetRow: index + 2 });
  });
  var removable = {}, headerRows = [], lineRows = [];
  Object.keys(candidateInstructions).forEach(function (no) {
    if (active[no]) return;
    var instructionHeaders = headerGroups[no] || [], instructionLines = lineGroups[no] || [];
    var headersFinal = instructionHeaders.every(function (item) {
      return [ENUM.헤더상태.완료, ENUM.헤더상태.취소].indexOf(toStr_(item.row[hState])) >= 0;
    });
    var linesFinal = instructionLines.every(function (item) {
      return [ENUM.라인상태.완료, ENUM.라인상태.취소].indexOf(toStr_(item.row[lState])) >= 0;
    });
    if (!headersFinal || !linesFinal) return;
    removable[no] = true;
    instructionHeaders.forEach(function (item) { headerRows.push(item.sheetRow); });
    instructionLines.forEach(function (item) { lineRows.push(item.sheetRow); });
  });
  try { deleteSheetRowsByIndexes_(headers.sheet, headerRows); }
  catch (headerError) { headerError.headerRowsDeleted = headerError.deletedRows || 0; throw headerError; }
  try { deleteSheetRowsByIndexes_(lines.sheet, lineRows); }
  catch (lineError) {
    lineError.headerRowsDeleted = headerRows.length;
    lineError.lineRowsDeleted = lineError.deletedRows || 0;
    throw lineError;
  }
  return { headerRowsDeleted: headerRows.length, lineRowsDeleted: lineRows.length, removedInstructions: removable };
}

function deleteSheetRowsByIndexes_(sheet, indexes) {
  var sorted = (indexes || []).slice().sort(function (a, b) { return b - a; });
  if (!sorted.length) return 0;
  var end = sorted[0], start = end, deleted = 0;
  for (var i = 1; i <= sorted.length; i++) {
    var next = i < sorted.length ? sorted[i] : null;
    if (next !== null && next === start - 1) { start = next; continue; }
    try { sheet.deleteRows(start, end - start + 1); deleted += end - start + 1; }
    catch (e) { e.deletedRows = deleted; throw e; }
    if (next !== null) { start = next; end = next; }
  }
  return deleted;
}

function cleanupFolderByRetention_(root, cutoff, options) {
  options = options || {};
  var result = { filesDeleted: 0, foldersDeleted: 0 };
  cleanupMaintenanceFolder_(root, cutoff, options, result, true);
  return result;
}

function cleanupMaintenanceFolder_(folder, cutoff, options, result, isRoot) {
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next(), updated = file.getLastUpdated();
    if (!updated || updated.getTime() >= cutoff.getTime()) continue;
    if (options.allowedNames && !options.allowedNames[file.getName()]) continue;
    try { file.setTrashed(true); result.filesDeleted++; }
    catch (e) { e.cleanupResult = result; throw e; }
  }
  var children = [], iterator = folder.getFolders();
  while (iterator.hasNext()) children.push(iterator.next());
  children.forEach(function (child) {
    try { cleanupMaintenanceFolder_(child, cutoff, options, result, false); }
    catch (e) { e.cleanupResult = result; throw e; }
  });
  var dateFolder = /^\d{4}-\d{2}(?:-\d{2})?$/.test(folder.getName ? folder.getName() : '');
  if (!isRoot && dateFolder && !folder.getFiles().hasNext() && !folder.getFolders().hasNext()) {
    try { folder.setTrashed(true); result.foldersDeleted++; }
    catch (e) { e.cleanupResult = result; throw e; }
  }
}

function maintenancePdfNames_(instructions) {
  var names = {};
  Object.keys(instructions || {}).forEach(function (no) { names[no + '.pdf'] = true; });
  return names;
}

function maintenanceDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (isBlank_(value)) return null;
  var parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function formatMaintenanceDate_(value) {
  return value instanceof Date ? Utilities.formatDate(value, tz_(), 'yyyy-MM-dd HH:mm:ss') : toStr_(value);
}

function buildMaintenanceSummary_(result) {
  var lines = [
    '백업일시: ' + (result.backup ? result.backup.fileName.replace('Polar Penguin Backup ', '') : '-'),
    '백업파일명: ' + (result.backup ? result.backup.fileName : '-'),
    '백업링크: ' + (result.backup ? result.backup.url : '-'),
    '보존기간: ' + result.retentionDays + '일', '', '정리 결과:',
    '- 아카이브 주문 라인: ' + result.archive.inserted,
    '- 삭제 주문 라인: ' + result.sheets.ordersDeleted,
    '- 삭제 피킹 헤더: ' + result.sheets.pickingHeadersDeleted,
    '- 삭제 피킹 라인: ' + result.sheets.pickingLinesDeleted,
    '- 삭제 Success 파일: ' + result.files.successDeleted,
    '- 삭제 Error 파일: ' + result.files.errorDeleted,
    '- 삭제 Output 파일: ' + result.files.outputDeleted,
    '- 삭제 빈 폴더: ' + result.files.foldersDeleted
  ];
  if (result.errors.length) lines.push('', '경고/오류:', result.errors.join('\n'));
  return lines.join('\n');
}

function sendSystemNotification_(level, title, details) {
  level = String(level || 'INFO').toUpperCase();
  var recipient = '';
  try { recipient = String(param_('알림이메일', '') || '').trim(); }
  catch (configError) {
    try { writeOpLog_('sendSystemNotification_', '경고', '알림 설정 조회 실패 / ' + configError.message); } catch (ignore) { }
    return { sent: false, reason: '알림 설정 조회 실패: ' + configError.message };
  }
  if (!recipient) {
    try { writeOpLog_('sendSystemNotification_', '경고', '알림이메일 미설정 / ' + level + ' / ' + title); }
    catch (ignoreBlankEmailLog) { }
    return { sent: false, reason: '알림이메일 미설정' };
  }
  var subject = '[Polar Penguin]' + (level === 'INFO' ? '' : '[' + level + ']') + ' ' + title, body = '';
  try {
    body = typeof details === 'string' ? details : Object.keys(details || {}).map(function (key) {
      return key + ': ' + details[key];
    }).join('\n');
    body += '\n\n발생시각: ' + Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss');
    GmailApp.sendEmail(recipient, subject, body);
    return { sent: true, recipient: recipient, subject: subject };
  } catch (e) {
    try { writeOpLog_('sendSystemNotification_', '경고', level + ' 알림 발송 실패 / ' + e.message); }
    catch (ignoreSendFailureLog) { }
    return { sent: false, reason: e.message, subject: subject };
  }
}
