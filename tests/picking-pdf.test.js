const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync(path.resolve(__dirname, '..', 'src', 'S9 작업지시서.js'), 'utf8');

function table(headers, rows) {
  const headerIndex = {};
  headers.forEach((header, index) => { headerIndex[String(header).replace(/\s/g, '').toLowerCase()] = index; });
  return { headers, rows, headerIndex, sheet: {} };
}

function createContext(tables = {}) {
  const context = vm.createContext({ console });
  vm.runInContext(source, context);
  context.ROLE = { 헤더: 'header', 주문: 'orders', 라인: 'lines' };
  context.COL = {
    피킹지시번호: '피킹지시번호', 주문번호: '주문번호', 상태: '상태', 생성일시: '생성일시', 출력일시: '출력일시',
    품목별주문번호: '품목별 주문번호', 보관위치: '보관위치', 상품코드: '상품코드',
    상품명: '상품명', 옵션: '옵션', 필요수량: '필요수량', 라인상태: '라인상태',
    수령인: '수령인', 수령인휴대전화: '수령인 휴대전화', 수령인우편번호: '수령인 우편번호',
    수령인주소전체: '수령인 주소(전체)', 배송메시지: '배송메시지', 확정일시: '확정일시'
  };
  context.ENUM = { 헤더상태: { 대기: '대기', 완료: '완료', 취소: '취소', 출력오류: '출력오류' }, 라인상태: { 취소: '취소' } };
  context.normKey_ = value => String(value == null ? '' : value).replace(/\s/g, '').toLowerCase();
  context.col_ = (sourceTable, name, required) => {
    const index = sourceTable.headers.indexOf(name);
    if (index < 0 && required) throw new Error(`missing ${name}`);
    return index;
  };
  context.toStr_ = value => String(value == null ? '' : value).trim();
  context.toNum_ = value => Number(value) || 0;
  context.isBlank_ = value => value == null || String(value).trim() === '';
  context.readTable_ = role => tables[role];
  context.Utilities = { formatDate: (_date, _tz, format) => {
    if (format === 'yyyy-MM-dd') return '2026-08-19';
    if (format === 'MM/dd HH:mm') return '08/21 10:30';
    return '2026-08-19 10:00';
  } };
  context.tz_ = () => 'Asia/Seoul';
  context.withLock_ = fn => fn();
  context.writeOpLog_ = () => {};
  context.D0_대시보드전체갱신 = () => {};
  return context;
}

function documentTables(orderSpecs, instructionNo = 'PK-1') {
  const headerRows = [], orderRows = [], lineRows = [];
  orderSpecs.forEach((order, orderIndex) => {
    headerRows.push([instructionNo, order.orderNo, '대기', '']);
    order.items.forEach((item, itemIndex) => {
      const itemNo = `${order.orderNo}-I${itemIndex + 1}`;
      orderRows.push([order.orderNo, itemNo, order.orderDate || '', order.confirmedAt === undefined ?
        `2026-08-${String(orderIndex + 1).padStart(2, '0')} 09:00` : order.confirmedAt,
        order.recipient || '', order.phone || '', order.postal || '', order.address || '', order.message || '']);
      lineRows.push([instructionNo, order.orderNo, itemNo, item.sku, item.location || '', item.name, item.option || '', item.quantity,
        item.state || '미처리']);
    });
  });
  return {
    header: table(['피킹지시번호', '주문번호', '상태', '출력일시'], headerRows),
    orders: table(['주문번호', '품목별 주문번호', '주문일시', '확정일시', '수령인', '수령인 휴대전화', '수령인 우편번호', '수령인 주소(전체)', '배송메시지'], orderRows),
    lines: table(['피킹지시번호', '주문번호', '품목별 주문번호', '상품코드', '보관위치', '상품명', '옵션', '필요수량', '라인상태'], lineRows)
  };
}

test('73 reservation orders aggregate one SKU to an explicit total of 100', () => {
  const orders = Array.from({ length: 73 }, (_, index) => ({
    orderNo: `O-${String(index + 1).padStart(3, '0')}`,
    items: [{ sku: 'BOOK-1', location: 'A-03-02', name: '만화책1', quantity: index < 27 ? 2 : 1 }]
  }));
  const context = createContext(documentTables(orders, 'PK-RES-20260819-001'));
  const data = context.buildPickingDocumentData_('PK-RES-20260819-001');
  assert.equal(data.pickSummary.length, 1);
  assert.equal(data.pickSummary[0].quantity, 100);
  assert.equal(data.pickSummary[0].orderCount, 73);
  assert.equal(data.orders.length, 73);
  assert.equal(data.summary.totalQuantity, 100);
  assert.equal(data.orders[0].orderNo, 'O-001');
  assert.equal(data.orders[72].orderNo, 'O-073');
});

test('all order items aggregate by SKU while different option SKUs stay separate', () => {
  const context = createContext(documentTables([
    { orderNo: 'A001', items: [
      { sku: 'BOOK000A', location: 'A-01', name: '책', option: '1권', quantity: 2 },
      { sku: 'CARD', location: 'B-01', name: '포토카드', quantity: 1 }
    ] },
    { orderNo: 'A002', items: [
      { sku: 'BOOK000A', location: 'A-01', name: '책', option: '1권', quantity: 3 },
      { sku: 'BOOK000B', location: 'A-02', name: '책', option: '2권', quantity: 1 }
    ] }
  ]));
  const summary = context.buildPickingDocumentData_('PK-1').pickSummary;
  assert.deepEqual(JSON.parse(JSON.stringify(summary.map(item => [item.sku, item.quantity]))),
    [['BOOK000A', 5], ['BOOK000B', 1], ['CARD', 1]]);
});

test('packing data preserves customer match fields without exposing a full phone number', () => {
  const context = createContext(documentTables([{ orderNo: 'A001', confirmedAt: '2026-08-01 08:30',
    recipient: '김희성', phone: '010-9876-1234', postal: '01234', address: '서울시 중구', message: '문 앞',
    items: [{ sku: 'BOOK', location: 'A-01', name: '만화책', quantity: 2 }] }]));
  const order = context.buildPickingDocumentData_('PK-1').orders[0];
  assert.equal(order.orderNo, 'A001');
  assert.equal(order.recipient, '김희성');
  assert.equal(order.phoneLast4, '1234');
  assert.equal(order.postalCode, '01234');
  assert.equal(order.confirmedAt, '2026-08-01 08:30');
  assert.equal(order.orderDateText, '08/01 08:30');
  assert.equal(order.items[0].quantity, 2);
  assert.equal(JSON.stringify(order).includes('010-9876-1234'), false);
});

test('포장표 주문일시는 레거시 주문일시 없이 확정일시를 사용한다', () => {
  const tables = documentTables([{ orderNo: 'A001', orderDate: '1999-01-01 00:00', confirmedAt: '2026-08-21 19:42',
    items: [{ sku: 'BOOK', location: 'A-01', name: '만화책', quantity: 1 }] }]);
  tables.orders.headers.splice(2, 1);
  tables.orders.rows.forEach(row => row.splice(2, 1));
  const data = createContext(tables).buildPickingDocumentData_('PK-1');
  assert.equal(data.orders[0].confirmedAt, '2026-08-21 19:42');
  assert.equal(data.orders[0].orderDateText, '08/21 19:42');
});

test('포장표의 Date 확정일시는 MM/dd HH:mm으로 형식화하고 빈 값은 하이픈이다', () => {
  const context = createContext();
  const date = vm.runInContext("new Date('2026-08-21T10:30:00+09:00')", context);
  assert.equal(context.S9_formatPackingDateTime_(date), '08/21 10:30');
  assert.equal(context.S9_formatPackingDateTime_(''), '-');
  assert.equal(context.S9_formatPackingDateTime_('not-a-date'), '-');
});

test('포장표는 주문 첫 행만 order-start로 표시하고 공백 구분 행을 추가하지 않는다', () => {
  const context = createContext(documentTables([
    { orderNo: 'A', confirmedAt: '2026-08-21 10:30', items: [
      { sku: 'A1', location: 'A-01', name: '상품 A1', quantity: 1 },
      { sku: 'A2', location: 'A-02', name: '상품 A2', quantity: 1 }
    ] },
    { orderNo: 'B', confirmedAt: '2026-08-21 11:15', items: [
      { sku: 'B1', location: 'B-01', name: '상품 B1', quantity: 1 },
      { sku: 'B2', location: 'B-02', name: '상품 B2', quantity: 1 }
    ] }
  ]));
  const html = context.renderPickingDocumentHTML_(context.buildPickingDocumentData_('PK-1'));
  const packingBody = html.split('<table class="packing">')[1];
  assert.equal((packingBody.match(/<tr class="order-start/g) || []).length, 2);
  assert.equal((packingBody.match(/<tr class="">/g) || []).length, 2);
  assert.match(packingBody, /order-start first-order/);
  assert.match(html, /order-start:not\(\.first-order\) td\{border-top:2px solid #333\}/);
  assert.doesNotMatch(packingBody, /<tr[^>]*>\s*<\/tr>/);
});

test('확정일시로 바꿔 표시해도 기존 주문 정렬은 변하지 않는다', () => {
  const context = createContext(documentTables([
    { orderNo: 'B', orderDate: '2026-08-01 09:00', confirmedAt: '2026-08-21 11:15',
      items: [{ sku: 'B1', location: 'B-01', name: '상품 B', quantity: 1 }] },
    { orderNo: 'A', orderDate: '2026-08-02 09:00', confirmedAt: '2026-08-21 10:30',
      items: [{ sku: 'A1', location: 'A-01', name: '상품 A', quantity: 1 }] }
  ]));
  const orders = context.buildPickingDocumentData_('PK-1').orders;
  assert.deepEqual(Array.from(orders, order => order.orderNo), ['B', 'A']);
});

test('취소된 피킹 라인은 기존처럼 포장표에서 제외된다', () => {
  const context = createContext(documentTables([{ orderNo: 'A001', items: [
    { sku: 'KEEP', location: 'A-01', name: '정상 상품', quantity: 1 },
    { sku: 'CANCEL', location: 'A-02', name: '취소 상품', quantity: 1, state: '취소' }
  ] }]));
  const data = context.buildPickingDocumentData_('PK-1');
  assert.equal(data.orders[0].items.length, 1);
  assert.equal(data.orders[0].items[0].sku, 'KEEP');
});

test('compact shared HTML has two tables, missing-location warning, and no cart or O/X workflow', () => {
  const context = createContext(documentTables([{ orderNo: 'A001', items: [
    { sku: 'SKU-1', name: '위치없는 상품', quantity: 1 }
  ] }]));
  const html = context.renderPickingDocumentHTML_(context.buildPickingDocumentData_('PK-1'));
  assert.match(html, /@page\{size:A4 landscape;margin:6mm 7mm\}/);
  assert.match(html, /창고 피킹 요약/);
  assert.match(html, /주문별 포장 \/ 송장 확인/);
  assert.match(html, /위치 미지정/);
  assert.doesNotMatch(html, /카트 슬롯|피킹담당자|확인란|O\/X/);
});

test('PDF is created once, then reused without a second shipment mutation', () => {
  const context = createContext(documentTables([{ orderNo: 'A001', items: [
    { sku: 'SKU-1', location: 'A-01', name: '상품', quantity: 1 }
  ] }]));
  const files = new Map();
  const emptyIterator = () => ({ hasNext: () => false, next: () => undefined });
  const dateFolder = {
    getFilesByName: name => ({ hasNext: () => files.has(name), next: () => files.get(name) }),
    createFile: blob => { const file = { getId: () => 'pdf-id' }; files.set(blob.name, file); return file; }
  };
  const outputRoot = {
    getFilesByName: () => emptyIterator(),
    getFolders: () => { let used = false; return { hasNext: () => !used, next: () => { used = true; return dateFolder; } }; }
  };
  context.MimeType = { PDF: 'application/pdf' };
  context.getOrCreateSubFolder_ = () => dateFolder;
  context.HtmlService = { createHtmlOutput: () => ({
    getBlob: () => ({ getAs: () => ({ setName(name) { this.name = name; return this; } }) })
  }) };
  let committed = false;
  let mutations = 0;
  context.finalizePickingAfterOutput_ = () => {
    if (committed) return { 이미완료: true };
    committed = true; mutations++; return { 이미완료: false };
  };
  const first = context.S9_피킹PDF생성('PK-1', outputRoot);
  const second = context.S9_피킹PDF생성('PK-1', outputRoot);
  assert.equal(first.생성, true);
  assert.equal(second.재사용, true);
  assert.equal(files.size, 1);
  assert.equal(mutations, 1);
});

test('manual first output needs neither worker nor slot and shares the finalizer', () => {
  const context = createContext(documentTables([{ orderNo: 'A001', items: [
    { sku: 'SKU-1', location: 'A-01', name: '상품', quantity: 2 }
  ] }]));
  const calls = [];
  context.finalizePickingAfterOutput_ = (instructionNo, options) => {
    calls.push([instructionNo, options.source]); return { 이미완료: false };
  };
  const output = context.S9_수동출력준비_('PK-1');
  assert.match(output.html, /창고 피킹 요약/);
  assert.deepEqual(JSON.parse(JSON.stringify(calls)), [['PK-1', 'MANUAL_OUTPUT']]);
});
