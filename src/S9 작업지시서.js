/**
 * S9. 상품 중심 피킹지시서.
 *
 * 한 문서는 A) SKU별 창고 피킹 요약과 B) 주문별 포장/송장 확인표로 구성한다.
 * 카트와 작업자 배정은 창고의 물리 운영에 맡기며 출력 조건으로 사용하지 않는다.
 * PDF 파일 또는 수동 인쇄 화면이 처음 준비된 뒤 공용 finalizer가 출고를 확정한다.
 */

function S9_1_작업지시서출력() {
  var html = HtmlService.createHtmlOutputFromFile('PickingInstructions').setWidth(1080).setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, '피킹지시서 조회 / 재출력');
}

// HtmlService에서 호출할 공개 endpoint. 계산 함수는 아래의 내부 helper에 둔다.
function getPickingInstructionList() { return getPickingInstructionList_(); }
function preparePickingInstructionOutput(instructionNo) {
  try { return S9_수동출력준비_(instructionNo); }
  catch (e) {
    sendSystemNotification_('ERROR', '피킹지시서 출력 준비 실패', {
      피킹지시번호: toStr_(instructionNo), 오류: e.message,
      조치: '피킹지시서 조회 / 재출력에서 상태를 확인하고 다시 시도하세요.'
    });
    throw e;
  }
}
function retryPickingInstructionPdf(instructionNo) { return S9_피킹PDF재시도(instructionNo); }

/** 피킹라인을 SKU별로 합치며 주문 건수는 주문번호 unique count로 센다. */
function aggregatePickingItems_(items) {
  var grouped = {};
  (items || []).forEach(function (item) {
    var sku = toStr_(item.sku);
    if (!sku) return;
    if (!grouped[sku]) grouped[sku] = {
      location: toStr_(item.location), sku: sku, productName: toStr_(item.productName),
      option: toStr_(item.option), orderCount: 0, quantity: 0, _orders: {}
    };
    var entry = grouped[sku];
    entry.quantity += toNum_(item.quantity);
    entry._orders[toStr_(item.orderNo)] = true;
    if (!entry.location && item.location) entry.location = toStr_(item.location);
  });
  var result = Object.keys(grouped).map(function (sku) {
    var entry = grouped[sku];
    entry.orderCount = Object.keys(entry._orders).filter(String).length;
    delete entry._orders;
    if (!entry.location) entry.location = '(위치 미지정)';
    return entry;
  });
  result.sort(function (a, b) {
    var aMissing = a.location === '(위치 미지정)', bMissing = b.location === '(위치 미지정)';
    if (aMissing !== bMissing) return aMissing ? 1 : -1;
    var location = a.location.localeCompare(b.location); if (location) return location;
    var sku = a.sku.localeCompare(b.sku); if (sku) return sku;
    return (a.productName + ' ' + a.option).localeCompare(b.productName + ' ' + b.option);
  });
  return result;
}

/** 고객 메타데이터와 품목별 피킹라인을 주문번호 단위로 결합한다. */
function buildPackingOrders_(orderMeta, items) {
  var grouped = {};
  (items || []).forEach(function (item) {
    var no = toStr_(item.orderNo); if (!no) return;
    if (!grouped[no]) {
      var meta = orderMeta[no] || {};
      grouped[no] = {
        orderNo: no, orderDate: meta.orderDate || '', orderDateText: meta.orderDateText || '',
        recipient: meta.recipient || '', phoneLast4: meta.phoneLast4 || '', postalCode: meta.postalCode || '',
        address: meta.address || '', message: meta.message || '', items: []
      };
    }
    grouped[no].items.push({
      itemOrderNo: toStr_(item.itemOrderNo), sku: toStr_(item.sku),
      productName: toStr_(item.productName), option: toStr_(item.option), quantity: toNum_(item.quantity)
    });
  });
  var result = Object.keys(grouped).map(function (no) {
    grouped[no].items.sort(function (a, b) { return a.itemOrderNo.localeCompare(b.itemOrderNo); });
    return grouped[no];
  });
  result.sort(function (a, b) {
    var ad = S9_sortDate_(a.orderDate), bd = S9_sortDate_(b.orderDate);
    if (ad !== bd) return ad < bd ? -1 : 1;
    return a.orderNo.localeCompare(b.orderNo);
  });
  return result;
}

/** HTML과 PDF가 공유하는 순수 문서 DTO를 만든다. */
function buildPickingDocumentData_(instructionNo) {
  instructionNo = toStr_(instructionNo);
  if (!instructionNo) throw new Error('피킹지시번호를 입력하세요.');

  var headers = readTable_(ROLE.헤더);
  var H = {
    instruction: col_(headers, COL.피킹지시번호, true), orderNo: col_(headers, COL.주문번호, true),
    state: col_(headers, COL.상태, true), createdAt: S9_optionalColumn_(headers, [COL.생성일시]),
    outputAt: S9_optionalColumn_(headers, [COL.출력일시])
  };
  var selectedOrders = {}, createdAt = '', foundHeader = false;
  headers.rows.forEach(function (row) {
    if (toStr_(row[H.instruction]) !== instructionNo) return;
    foundHeader = true;
    if (toStr_(row[H.state]) === ENUM.헤더상태.취소) return;
    selectedOrders[toStr_(row[H.orderNo])] = true;
    if (!createdAt && H.createdAt >= 0 && !isBlank_(row[H.createdAt])) createdAt = row[H.createdAt];
    if (!createdAt && H.outputAt >= 0 && !isBlank_(row[H.outputAt])) createdAt = row[H.outputAt];
  });
  if (!foundHeader) throw new Error('피킹지시번호를 찾을 수 없습니다: ' + instructionNo);

  var orders = readTable_(ROLE.주문);
  var O = {
    orderNo: col_(orders, COL.주문번호, true), itemNo: col_(orders, COL.품목별주문번호, true),
    orderDate: S9_optionalColumn_(orders, ['주문일시', '주문일자', '주문일', '결제일시']),
    recipient: S9_optionalColumn_(orders, ['수령인', '수령인명', '수령인 이름']),
    phone: S9_optionalColumn_(orders, ['수령인 휴대전화', '수령인 휴대폰', '수령인휴대전화', '휴대전화']),
    postal: S9_optionalColumn_(orders, ['수령인 우편번호', '수령인우편번호', '우편번호']),
    address: S9_optionalColumn_(orders, ['수령인 주소', '수령인 주소(전체)', '수령인주소', '기본주소']),
    detailAddress: S9_optionalColumn_(orders, ['상세주소', '수령인 상세주소']),
    message: S9_optionalColumn_(orders, ['배송메시지', '배송 메세지', '배송 요청사항'])
  };
  var itemOrder = {}, orderMeta = {};
  orders.rows.forEach(function (row) {
    var no = toStr_(row[O.orderNo]);
    itemOrder[toStr_(row[O.itemNo])] = no;
    if (!selectedOrders[no] || orderMeta[no]) return;
    var phone = O.phone >= 0 ? toStr_(row[O.phone]).replace(/\D/g, '') : '';
    var address = O.address >= 0 ? toStr_(row[O.address]) : '';
    var detail = O.detailAddress >= 0 ? toStr_(row[O.detailAddress]) : '';
    orderMeta[no] = {
      orderDate: O.orderDate >= 0 ? row[O.orderDate] : '',
      orderDateText: O.orderDate >= 0 ? S9_formatDateTime_(row[O.orderDate]) : '',
      recipient: O.recipient >= 0 ? toStr_(row[O.recipient]) : '',
      phoneLast4: phone ? phone.slice(-4) : '', postalCode: O.postal >= 0 ? toStr_(row[O.postal]) : '',
      address: [address, detail].filter(String).join(' '), message: O.message >= 0 ? toStr_(row[O.message]) : ''
    };
  });

  var lines = readTable_(ROLE.라인);
  var L = {
    instruction: col_(lines, COL.피킹지시번호, true), orderNo: S9_optionalColumn_(lines, [COL.주문번호]),
    itemNo: col_(lines, COL.품목별주문번호, true), sku: col_(lines, COL.상품코드, true),
    location: col_(lines, COL.보관위치, true), name: col_(lines, COL.상품명, true),
    option: col_(lines, COL.옵션, true), quantity: col_(lines, COL.필요수량, true),
    state: col_(lines, COL.라인상태, true)
  };
  var items = [];
  lines.rows.forEach(function (row) {
    if (toStr_(row[L.instruction]) !== instructionNo || toStr_(row[L.state]) === ENUM.라인상태.취소) return;
    var no = (L.orderNo >= 0 ? toStr_(row[L.orderNo]) : '') || itemOrder[toStr_(row[L.itemNo])];
    if (!selectedOrders[no]) return;
    items.push({
      orderNo: no, itemOrderNo: toStr_(row[L.itemNo]), location: toStr_(row[L.location]),
      sku: toStr_(row[L.sku]), productName: toStr_(row[L.name]),
      option: toStr_(row[L.option]) === '-' ? '' : toStr_(row[L.option]), quantity: toNum_(row[L.quantity])
    });
  });
  if (!items.length) throw new Error('출력할 피킹 품목이 없습니다: ' + instructionNo);

  var pickSummary = aggregatePickingItems_(items), packingOrders = buildPackingOrders_(orderMeta, items);
  var totalQuantity = pickSummary.reduce(function (sum, item) { return sum + item.quantity; }, 0);
  return {
    instructionNo: instructionNo, title: instructionNo.indexOf('-RES-') >= 0 ? '예약 피킹지시서' : '피킹지시서',
    createdAt: S9_formatDateTime_(createdAt || new Date()),
    summary: { orderCount: packingOrders.length, skuCount: pickSummary.length, totalQuantity: totalQuantity },
    missingLocationCount: pickSummary.filter(function (item) { return item.location === '(위치 미지정)'; }).length,
    pickSummary: pickSummary, orders: packingOrders,
    hasAddress: packingOrders.some(function (order) { return !!order.address; }),
    hasMessage: packingOrders.some(function (order) { return !!order.message; })
  };
}

/** A4 landscape용 compact HTML. 자동 PDF와 수동 출력이 이 템플릿을 함께 쓴다. */
function renderPickingDocumentHTML_(data) {
  var e = S9_escapeHtml_, parts = [
    '<!doctype html><html><head><meta charset="utf-8"><style>',
    '@page{size:A4 landscape;margin:6mm 7mm}*{box-sizing:border-box}',
    'body{font-family:"Malgun Gothic",Arial,sans-serif;color:#20242a;font-size:8.5pt;line-height:1.2;margin:0}',
    'h1{font-size:14pt;margin:0 0 2px}.meta{font-size:9pt;margin-bottom:5px}.warn{color:#b42318;font-weight:700}',
    'h2{font-size:10pt;margin:7px 0 3px;padding:3px 5px;background:#1f3864;color:#fff}',
    'table{width:100%;border-collapse:collapse;table-layout:fixed;margin-bottom:5px}',
    'thead{display:table-header-group}th{background:#e8edf3;font-size:8pt;font-weight:700}',
    'th,td{border:1px solid #9da7b3;padding:2px 3px;vertical-align:top;word-break:break-word}',
    '.num{text-align:right;font-weight:700}.center{text-align:center}.sku{font-family:Consolas,monospace}',
    '.missing{background:#fff1f0;color:#b42318;font-weight:700}.address,.message{font-size:7.5pt}',
    'tr{page-break-inside:avoid}</style></head><body>',
    '<h1>', e(data.title), ' &nbsp; ', e(data.instructionNo), '</h1>',
    '<div class="meta">', e(data.createdAt), ' &nbsp;·&nbsp; 주문 ', data.summary.orderCount,
    '건 &nbsp;·&nbsp; SKU ', data.summary.skuCount, '종 &nbsp;·&nbsp; 총 ', data.summary.totalQuantity, '개',
    data.missingLocationCount ? ' &nbsp;·&nbsp; <span class="warn">위치 미지정 상품 ' + data.missingLocationCount + '종</span>' : '',
    '</div><h2>창고 피킹 요약</h2>',
    '<table><thead><tr><th style="width:13%">보관위치</th><th style="width:17%">상품코드 / SKU</th>',
    '<th>상품명</th><th style="width:16%">옵션</th><th style="width:9%">주문 건수</th><th style="width:9%">총 필요수량</th>',
    '</tr></thead><tbody>'
  ];
  data.pickSummary.forEach(function (item) {
    var missing = item.location === '(위치 미지정)';
    parts.push('<tr><td class="', missing ? 'missing' : '', '">', missing ? '⚠ ' : '', e(item.location),
      '</td><td class="sku">', e(item.sku), '</td><td>', e(item.productName), '</td><td>', e(item.option || '-'),
      '</td><td class="center">', item.orderCount, '</td><td class="num">', item.quantity, '</td></tr>');
  });
  parts.push('</tbody></table><h2>주문별 포장 / 송장 확인</h2><table><thead><tr>',
    '<th style="width:10%">주문일시</th><th style="width:12%">주문번호</th><th style="width:8%">수령인</th>',
    '<th style="width:6%">전화 끝4</th><th style="width:7%">우편번호</th><th>상품명</th>',
    '<th style="width:11%">옵션</th><th style="width:5%">수량</th>');
  if (data.hasAddress) parts.push('<th style="width:18%">배송주소</th>');
  if (data.hasMessage) parts.push('<th style="width:14%">배송메시지</th>');
  parts.push('</tr></thead><tbody>');
  data.orders.forEach(function (order) {
    order.items.forEach(function (item) {
      parts.push('<tr><td>', e(order.orderDateText), '</td><td>', e(order.orderNo), '</td><td>', e(order.recipient),
        '</td><td class="center">', e(order.phoneLast4), '</td><td>', e(order.postalCode), '</td><td>', e(item.productName),
        '</td><td>', e(item.option || '-'), '</td><td class="num">', item.quantity, '</td>');
      if (data.hasAddress) parts.push('<td class="address">', e(order.address), '</td>');
      if (data.hasMessage) parts.push('<td class="message">', e(order.message), '</td>');
      parts.push('</tr>');
    });
  });
  parts.push('</tbody></table></body></html>');
  return parts.join('');
}

function S9_피킹PDF생성(instructionNo, outputRoot) {
  instructionNo = toStr_(instructionNo);
  if (!instructionNo) return { 생성: false, 사유: '지시번호 없음' };
  outputRoot = outputRoot || DriveApp.getFolderById(String(param_('Output폴더ID', '')));
  var fileName = instructionNo + '.pdf', existing = S9_findPDFFile_(outputRoot, fileName);
  if (existing) return { 생성: false, 재사용: true, 파일명: fileName, 파일ID: existing.getId ? existing.getId() : '',
    출고확정: finalizePickingAfterOutput_(instructionNo, { refresh: false, source: 'PDF_REUSE' }) };

  var data = buildPickingDocumentData_(instructionNo);
  var html = renderPickingDocumentHTML_(data);
  var pdf = HtmlService.createHtmlOutput(html).getBlob().getAs(MimeType.PDF).setName(fileName);
  var dateFolder = getOrCreateSubFolder_(outputRoot, Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'));
  var file = dateFolder.createFile(pdf);
  var finalization = finalizePickingAfterOutput_(instructionNo, { refresh: false, source: 'PDF_OUTPUT' });
  writeOpLog_('S9_피킹PDF생성', '성공', fileName + ' / SKU ' + data.summary.skuCount + ' / 수량 ' + data.summary.totalQuantity);
  return { 생성: true, 파일ID: file.getId(), 파일명: fileName, 출고확정: finalization };
}

/** 수동 화면의 첫 출력과 재출력도 동일 DTO/HTML/finalizer를 사용한다. */
function S9_수동출력준비_(instructionNo) {
  return withLock_(function () {
    var data = buildPickingDocumentData_(instructionNo), html = renderPickingDocumentHTML_(data);
    var finalization = finalizePickingAfterOutput_(instructionNo, { refresh: false, source: 'MANUAL_OUTPUT' });
    try { D0_대시보드전체갱신(true); } catch (ignore) { }
    writeOpLog_('S9_수동출력준비_', '성공', instructionNo + ' / ' + (finalization.이미완료 ? '재출력' : '첫 출력'));
    return { html: html, summary: data.summary, 이미완료: finalization.이미완료 };
  });
}

function S9_피킹PDF재시도(instructionNo) {
  instructionNo = toStr_(instructionNo);
  if (!instructionNo) throw new Error('피킹지시번호를 입력하세요.');
  if (!S9_ordersForInstruction_(instructionNo).length) throw new Error('피킹지시번호를 찾을 수 없습니다: ' + instructionNo);
  try {
    var result = S9_피킹PDF생성(instructionNo, inputFolder_('Output폴더ID'));
    D0_대시보드전체갱신(true);
    return { 메시지: (result.재사용 ? '기존 PDF를 확인했습니다.' : 'PDF를 다시 생성했습니다.') + ' ' + instructionNo };
  } catch (e) {
    markPickingOutputState_(instructionNo, ENUM.헤더상태.출력오류);
    sendSystemNotification_('ERROR', '피킹 PDF 재시도 실패', {
      피킹지시번호: instructionNo, 오류: e.message,
      조치: '피킹지시서 조회 / 재출력에서 상태를 확인하고 다시 시도하세요.'
    });
    throw e;
  }
}

/** 조회 화면용 최근 지시 목록. 카트/작업자 정보는 반환하지 않는다. */
function getPickingInstructionList_() {
  var headers = readTable_(ROLE.헤더), lines = readTable_(ROLE.라인);
  var H = { instruction: col_(headers, COL.피킹지시번호, true), orderNo: col_(headers, COL.주문번호, true),
    state: col_(headers, COL.상태, true), createdAt: S9_optionalColumn_(headers, [COL.생성일시]),
    outputAt: S9_optionalColumn_(headers, [COL.출력일시]) };
  var L = { instruction: col_(lines, COL.피킹지시번호, true), sku: col_(lines, COL.상품코드, true),
    qty: col_(lines, COL.필요수량, true), state: col_(lines, COL.라인상태, true) };
  var grouped = {};
  headers.rows.forEach(function (row) {
    var no = toStr_(row[H.instruction]); if (!no) return;
    if (!grouped[no]) grouped[no] = { instructionNo: no, createdAt: '', orders: {}, skus: {}, totalQuantity: 0, states: {} };
    grouped[no].orders[toStr_(row[H.orderNo])] = true; grouped[no].states[toStr_(row[H.state])] = true;
    if (!grouped[no].createdAt && H.createdAt >= 0 && !isBlank_(row[H.createdAt])) grouped[no].createdAt = S9_formatDateTime_(row[H.createdAt]);
    if (!grouped[no].createdAt && H.outputAt >= 0 && !isBlank_(row[H.outputAt])) grouped[no].createdAt = S9_formatDateTime_(row[H.outputAt]);
  });
  lines.rows.forEach(function (row) {
    var no = toStr_(row[L.instruction]); if (!grouped[no] || toStr_(row[L.state]) === ENUM.라인상태.취소) return;
    grouped[no].skus[toStr_(row[L.sku])] = true; grouped[no].totalQuantity += toNum_(row[L.qty]);
  });
  var pdfNames = S9_existingPDFNames_();
  return Object.keys(grouped).sort().reverse().slice(0, 100).map(function (no) {
    var item = grouped[no], state = item.states[ENUM.헤더상태.출력오류] ? ENUM.헤더상태.출력오류 :
      (item.states[ENUM.헤더상태.대기] ? ENUM.헤더상태.대기 :
        (item.states[ENUM.헤더상태.완료] ? ENUM.헤더상태.완료 : ENUM.헤더상태.취소));
    return { instructionNo: no, createdAt: item.createdAt || '-', orderCount: Object.keys(item.orders).filter(String).length,
      skuCount: Object.keys(item.skus).filter(String).length, totalQuantity: item.totalQuantity,
      state: state, pdfExists: !!pdfNames[no + '.pdf'] };
  });
}

function S9_ordersForInstruction_(instructionNo) {
  var table = readTable_(ROLE.헤더), cInstruction = col_(table, COL.피킹지시번호, true);
  var cOrder = col_(table, COL.주문번호, true), found = {};
  table.rows.forEach(function (row) { if (toStr_(row[cInstruction]) === instructionNo) found[toStr_(row[cOrder])] = true; });
  return Object.keys(found).filter(String);
}

function S9_findPDFFile_(outputRoot, fileName) {
  var direct = outputRoot.getFilesByName(fileName); if (direct.hasNext()) return direct.next();
  var folders = outputRoot.getFolders();
  while (folders.hasNext()) {
    var files = folders.next().getFilesByName(fileName); if (files.hasNext()) return files.next();
  }
  return null;
}

function S9_existingPDFNames_() {
  var names = {};
  try {
    var root = inputFolder_('Output폴더ID'), files = root.getFiles();
    while (files.hasNext()) names[files.next().getName()] = true;
    var folders = root.getFolders();
    while (folders.hasNext()) {
      files = folders.next().getFiles(); while (files.hasNext()) names[files.next().getName()] = true;
    }
  } catch (ignore) { }
  return names;
}

function S9_optionalColumn_(table, names) {
  for (var i = 0; i < names.length; i++) {
    var exact = table.headerIndex ? table.headerIndex[normKey_(names[i])] : undefined;
    if (exact !== undefined) return exact;
    var found = col_(table, names[i], false); if (found >= 0) return found;
  }
  return -1;
}

function S9_sortDate_(value) {
  if (value instanceof Date) return value.getTime();
  var parsed = Date.parse(toStr_(value)); return isNaN(parsed) ? Number.MAX_SAFE_INTEGER : parsed;
}

function S9_formatDateTime_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, tz_(), 'yyyy-MM-dd HH:mm');
  return toStr_(value);
}

function S9_escapeHtml_(value) {
  return String(value == null ? '' : value).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}
