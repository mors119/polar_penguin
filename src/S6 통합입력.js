/**
 * ============================================================
 *  S6. 통합 Input Pipeline
 * ============================================================
 *  Input 폴더의 CSV/구글 시트를 헤더로 판별해 기존 S1~S4/S9에 연결한다.
 *
 *  처리 순서
 *    Input → 파일 판별/검증 → S1 또는 S2 → S3 → S4 → S9 PDF
 *    성공 원본 → Processed/YYYY-MM-DD, 실패 원본 → Error/YYYY-MM-DD
 *
 *  재고 파일을 주문보다 먼저 처리해 예약대기 주문이 최신 재고로 재평가되게 한다.
 *  성공 체크섬은 입력처리로그에 남기며, 같은 내용의 재업로드는 오류로 격리한다.
 */

var INPUT_TYPE = { ORDER: 'ORDER', INVENTORY: 'INVENTORY', UNKNOWN: 'UNKNOWN' };

function processInput() {
  return withLock_(function () {
    var inputFolder = inputFolder_('통합Input폴더ID');
    var processedFolder = inputFolder_('Processed폴더ID');
    var errorFolder = inputFolder_('Error폴더ID');
    var outputFolder = inputFolder_('Output폴더ID');
    var files = [], iterator = inputFolder.getFiles();
    while (iterator.hasNext()) files.push(iterator.next());
    files.sort(function (a, b) {
      var typeOrder = inputFilePriority_(a) - inputFilePriority_(b);
      if (typeOrder) return typeOrder;
      return a.getLastUpdated().getTime() - b.getLastUpdated().getTime();
    });

    var report = { 감지: files.length, 성공: 0, 실패: 0, 주문: 0, 재고: 0, 결과: [] };
    files.forEach(function (file) {
      try {
        var result = processInputFile_(file, processedFolder, outputFolder);
        report.성공++;
        if (result.type === INPUT_TYPE.ORDER) report.주문++;
        if (result.type === INPUT_TYPE.INVENTORY) report.재고++;
        report.결과.push({ file: file.getName(), status: 'PROCESSED', type: result.type });
      } catch (e) {
        var code = e.inputCode || 'PROCESSING_FAILED';
        report.실패++;
        report.결과.push({ file: file.getName(), status: 'ERROR', code: code, message: e.message });
        recordInputLog_(file, e.inputFingerprint || '', e.inputType || INPUT_TYPE.UNKNOWN, 'ERROR', code, e.message);
        try { moveInputFile_(file, errorFolder); }
        catch (moveError) { writeOpLog_('processInput', '실패', file.getName() + ' / Error 이동 실패 / ' + moveError.message); }
        notifyInputFailure_(file.getName(), code, e.message);
      }
    });

    try { D0_대시보드전체갱신(true); }
    catch (dashboardError) { writeOpLog_('processInput', '경고', '대시보드 갱신 실패 / ' + dashboardError.message); }
    writeOpLog_('processInput', report.실패 ? '부분성공' : '성공',
      '감지 ' + report.감지 + ' / 성공 ' + report.성공 + ' / 실패 ' + report.실패);
    return report;
  });
}

/** 재고(0) → 주문(1) → 미판별/손상(2) 순으로 처리한다. */
function inputFilePriority_(file) {
  try {
    var type = detectInputType_(readUnifiedInput_(file)[0] || []);
    return type === INPUT_TYPE.INVENTORY ? 0 : (type === INPUT_TYPE.ORDER ? 1 : 2);
  } catch (e) {
    return 2;
  }
}

function processInputFile_(file, processedFolder, outputFolder) {
  var parsed = readUnifiedInput_(file);
  var fingerprint = inputFingerprint_(file, parsed);
  if (hasProcessedFingerprint_(fingerprint)) {
    throw inputError_('DUPLICATE_FILE', '이미 성공적으로 처리한 파일과 내용이 같습니다.', INPUT_TYPE.UNKNOWN, fingerprint);
  }

  var type = detectInputType_(parsed[0] || []);
  try { validateInput_(type, parsed); }
  catch (e) {
    e.inputType = type;
    e.inputFingerprint = fingerprint;
    throw e;
  }

  var business = runInputBusiness_(type, file, outputFolder);
  recordInputLog_(file, fingerprint, type, 'PROCESSED', '', inputBusinessMessage_(business));
  moveInputFile_(file, processedFolder);
  return { type: type, business: business };
}

/** 파일명을 보지 않고 첫 행의 필수 헤더 조합만으로 유형을 판별한다. */
function detectInputType_(headers) {
  var present = {};
  (headers || []).forEach(function (header) { present[normKey_(header)] = true; });
  var hasAny = function (names) {
    return names.some(function (name) { return present[normKey_(name)]; });
  };
  var order = [
    ['주문번호'], ['품목별 주문번호', '품목별주문번호', '주문상세번호'],
    ['상품품목코드', '품목코드', '상품코드'], ['수량', '주문수량']
  ].every(hasAny);
  if (order) return INPUT_TYPE.ORDER;

  var inventory = [['품목코드', '상품품목코드'], ['상품명'], ['재고수량']].every(hasAny);
  return inventory ? INPUT_TYPE.INVENTORY : INPUT_TYPE.UNKNOWN;
}

/** 필수값과 수량/재고의 숫자 형식을 비즈니스 처리 전에 검증한다. */
function validateInput_(type, data) {
  if (!data || data.length < 2) throw inputError_('EMPTY_FILE', '헤더 아래에 처리할 데이터가 없습니다.', type);
  if (type === INPUT_TYPE.UNKNOWN) {
    throw inputError_('UNKNOWN_TYPE', '주문 또는 재고 파일로 판별할 필수 헤더가 없습니다.', type);
  }

  var index = inputHeaderIndex_(data[0]);
  var fields = type === INPUT_TYPE.ORDER ? [
    { name: '주문번호', aliases: ['주문번호'] },
    { name: '품목별 주문번호', aliases: ['품목별 주문번호', '품목별주문번호', '주문상세번호'] },
    { name: '상품품목코드', aliases: ['상품품목코드', '품목코드', '상품코드'] },
    { name: '수량', aliases: ['수량', '주문수량'], number: true, positive: true }
  ] : [
    { name: '품목코드', aliases: ['품목코드', '상품품목코드'] },
    { name: '상품명', aliases: ['상품명'] },
    { name: '재고수량', aliases: ['재고수량'], number: true }
  ];
  fields.forEach(function (field) { field.index = inputAliasIndex_(index, field.aliases); });

  var rows = data.slice(1).filter(function (row) {
    return row.some(function (value) { return String(value == null ? '' : value).trim() !== ''; });
  });
  if (!rows.length) throw inputError_('EMPTY_FILE', '처리할 데이터 행이 없습니다.', type);
  rows.forEach(function (row, offset) {
    fields.forEach(function (field) {
      var value = row[field.index];
      if (value === '' || value === null || value === undefined) {
        throw inputError_('MISSING_VALUE', (offset + 2) + '행의 "' + field.name + '" 값이 비어 있습니다.', type);
      }
      if (field.number) {
        var number = Number(String(value).replace(/,/g, ''));
        if (!isFinite(number) || (field.positive && number <= 0)) {
          throw inputError_('INVALID_VALUE', (offset + 2) + '행의 "' + field.name + '" 값이 올바른 숫자가 아닙니다.', type);
        }
      }
    });
  });
  return { type: type, rows: rows.length };
}

/** 기존 S1~S4의 재고/주문/피킹 규칙을 그대로 조합하는 오케스트레이터. */
function runInputBusiness_(type, file, outputFolder) {
  var result = {};
  if (type === INPUT_TYPE.ORDER) result.ingest = S2_1_주문CSV취입(file, { skipMove: true, silent: true });
  else if (type === INPUT_TYPE.INVENTORY) result.inventory = S1_1_카페24재고동기화(file, true);
  else throw inputError_('UNKNOWN_TYPE', '지원하지 않는 입력 유형입니다.', type);

  result.confirm = S3_1_주문확정();
  result.picking = S4_1_피킹지시생성();
  if (result.picking && result.picking.생성) {
    result.pdf = S9_피킹PDF생성(result.picking.지시번호, outputFolder);
  }
  return result;
}

function readUnifiedInput_(file) {
  var name = file.getName();
  var mime = file.getMimeType();
  var csv = /\.csv$/i.test(name) || String(mime).indexOf('csv') >= 0;
  var sheet = mime === MimeType.GOOGLE_SHEETS;
  if (!csv && !sheet) throw inputError_('UNSUPPORTED_FORMAT', 'CSV 또는 Google Spreadsheet만 처리할 수 있습니다.', INPUT_TYPE.UNKNOWN);
  try {
    if (sheet) {
      var sh = SpreadsheetApp.openById(file.getId()).getSheets()[0];
      if (!sh || sh.getLastRow() < 1 || sh.getLastColumn() < 1) return [];
      return sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    }
    return Utilities.parseCsv(readCsvText_(file));
  } catch (e) {
    if (e.inputCode) throw e;
    throw inputError_('CORRUPT_FILE', '파일을 해석할 수 없습니다: ' + e.message, INPUT_TYPE.UNKNOWN);
  }
}

/** 파일 ID가 아닌 파싱된 내용의 SHA-256을 쓰므로 다시 업로드해도 중복을 검출한다. */
function inputFingerprint_(file, parsed) {
  var payload = JSON.stringify(parsed || []);
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, payload, Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(digest).replace(/=+$/, '');
}

function inputHeaderIndex_(headers) {
  var index = {};
  (headers || []).forEach(function (header, position) {
    var key = normKey_(header);
    if (key && index[key] === undefined) index[key] = position;
  });
  return index;
}

function inputAliasIndex_(index, aliases) {
  for (var i = 0; i < aliases.length; i++) {
    var found = index[normKey_(aliases[i])];
    if (found !== undefined) return found;
  }
  return -1;
}

/** 실패 이력은 재시도를 막지 않고 PROCESSED 이력만 중복으로 판정한다. */
function hasProcessedFingerprint_(fingerprint) {
  if (!fingerprint) return false;
  var sheet = ensureInputLogSheet_();
  if (sheet.getLastRow() < 2) return false;
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, INPUT_LOG_HEADERS.length).getValues();
  return rows.some(function (row) { return String(row[3]) === fingerprint && String(row[5]) === 'PROCESSED'; });
}

function recordInputLog_(file, fingerprint, type, status, code, message) {
  var row = [new Date(), file.getId(), file.getName(), fingerprint, type, status, code || '', String(message || '').substring(0, 2000)];
  prependRows_(ensureInputLogSheet_(), [row], [1, 2, 3]);
  if (status === 'ERROR') {
    writeOpLog_('processInput', '실패', file.getName() + ' / ' + code + ' / ' + message);
  }
}

function ensureInputLogSheet_() {
  var ss = consoleSS_();
  var sheet = ss.getSheetByName(CONSOLE.입력처리로그);
  if (!sheet) sheet = ss.insertSheet(CONSOLE.입력처리로그);
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, INPUT_LOG_HEADERS.length).setValues([INPUT_LOG_HEADERS]);
  sheet.setFrozenRows(1);
  return sheet;
}

function moveInputFile_(file, destinationRoot) {
  var dateFolder = getOrCreateSubFolder_(destinationRoot, Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'));
  file.moveTo(dateFolder);
}

function inputFolder_(key) {
  var id = extractDriveId_(param_(key, ''));
  if (!id) throw new Error('[설정] 파라미터에 "' + key + '"가 없습니다. setupSystem()을 실행하세요.');
  try { return DriveApp.getFolderById(id); }
  catch (e) { throw new Error('"' + key + '" 폴더를 열 수 없습니다. (' + id + ')'); }
}

function inputError_(code, message, type, fingerprint) {
  var error = new Error(message);
  error.inputCode = code;
  error.inputType = type || INPUT_TYPE.UNKNOWN;
  error.inputFingerprint = fingerprint || '';
  return error;
}

function inputBusinessMessage_(result) {
  var picking = result && result.picking;
  return picking && picking.생성 ? '피킹지시 ' + picking.지시번호 + ' 생성' : '입력 처리 완료 (신규 피킹 없음)';
}

function notifyInputFailure_(fileName, code, message) {
  var recipient = String(param_('알림이메일', '') || '').trim();
  if (!recipient) {
    writeOpLog_('processInput', '경고', '알림이메일 미설정 / ' + fileName + ' / ' + code);
    return false;
  }
  var processedAt = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss');
  try {
    GmailApp.sendEmail(recipient, '[Polar Penguin] Input Processing Failed',
      '파일명: ' + fileName + '\n오류유형: ' + code + '\n오류메시지: ' + message + '\n처리시각: ' + processedAt);
    return true;
  } catch (e) {
    writeOpLog_('processInput', '경고', '오류 알림 발송 실패 / ' + e.message);
    return false;
  }
}
