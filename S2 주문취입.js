
/**
 * ============================================================
 *  S2. 주문 CSV 취입
 * ============================================================
 *  Drive 폴더의 주문 CSV를 읽어 주문(완료) 시트에 적재한다.
 *
 *   · 품목별 주문번호 기준으로 중복을 제거한다
 *     (같은 파일을 두 번 올려도 중복 적재되지 않는다)
 *   · 주문번호·우편번호 등은 문자열로 강제해 앞자리 0을 보존한다
 *   · 신규 행은 시트 최상단에 삽입한다 (최신이 위로)
 *   · 초기 상태: 출고완료 0 · 피킹지시번호 공란 · 주문상태 접수
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
 
function S2_1_주문CSV취입() {
  return withLock_(function () {
    var 폴더ID = String(param_('CSV폴더ID', DEFAULT_FOLDER_ID));
    var 완료폴더명 = String(param_('CSV처리완료폴더명', '처리완료'));
 
    var folder;
    try {
      folder = DriveApp.getFolderById(폴더ID);
    } catch (e) {
      throw new Error('CSV 폴더를 열 수 없습니다. [설정] 탭의 CSV폴더ID를 확인하세요. (' + 폴더ID + ')');
    }
 
    // ---------- CSV 수집 (카페24 재고 파일은 제외) ----------
    var csvFiles = [];
    var it = folder.getFiles();
    while (it.hasNext()) {
      var f = it.next();
      var n = f.getName();
      if (n.indexOf(CAFE24.파일접두어) >= 0) continue;   // 재고 파일 제외

      var mime = f.getMimeType();
      var csv파일 = /\.csv$/i.test(n) || mime.indexOf('csv') >= 0;
      var 시트파일 = mime === MimeType.GOOGLE_SHEETS;
      if (!csv파일 && !시트파일) continue;

      csvFiles.push(f);
    }
    if (!csvFiles.length) {
      alert_('폴더에 처리할 주문 CSV가 없습니다.\n(카페24 재고 CSV는 S1에서 처리합니다)');
      return { 파일수: 0 };
    }
 
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
 
    var 리포트 = [], 신규행 = [];
    var 총신규 = 0, 총중복 = 0, 총오류 = 0;
 
    csvFiles.forEach(function (file) {
      var r = parseCsvFile_(file, 주문, c품목별, 기존키);
      신규행 = 신규행.concat(r.rows);
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
        row[c주문상태] = ENUM.주문상태.접수;
      });
      // 최신 주문이 위로 오도록 역순 정렬 후 삽입
      신규행.reverse();
      prependRows_(주문.sheet, 신규행, 텍스트열);
    }
 
    // ---------- 처리한 CSV 이동 ----------
    var 완료폴더 = getOrCreateSubFolder_(folder, 완료폴더명);
    csvFiles.forEach(function (file) {
      try {
        완료폴더.addFile(file);
        folder.removeFile(file);
      } catch (e) {
        리포트.push('⚠ ' + file.getName() + ' 이동 실패: ' + e.message);
      }
    });
 
    var msg = 'CSV 취입 완료\n\n' + 리포트.join('\n') +
      '\n\n합계 — 신규 ' + 총신규 + '건 / 중복 제외 ' + 총중복 + '건' +
      (총오류 ? ' / 오류 ' + 총오류 + '건' : '') +
      '\n\n다음: 「S3. 주문 확정」으로 재고를 검증하세요.';
 
    alert_(msg);
    writeOpLog_('S2_1_주문CSV취입', '성공', msg.replace(/\n/g, ' | '));
    return { 파일수: csvFiles.length, 신규: 총신규, 중복: 총중복, 오류: 총오류 };
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
 
  var rows = [], 중복 = 0, 오류 = 0;
 
  for (var i = 1; i < parsed.length; i++) {
    var src = parsed[i];
    if (!src.join('').trim()) continue;
 
    var row = new Array(주문.width).fill('');
    for (var j = 0; j < src.length && j < 매핑.length; j++) {
      if (매핑[j] >= 0) row[매핑[j]] = src[j];
    }
 
    var key = toStr_(row[c품목별]);
    if (!key) { 오류++; continue; }
    if (기존키[key]) { 중복++; continue; }
 
    기존키[key] = true;
    rows.push(row);
  }
 
  return { rows: rows, 중복: 중복, 오류: 오류, 총행: parsed.length - 1 };
}
 
function getOrCreateSubFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
 

