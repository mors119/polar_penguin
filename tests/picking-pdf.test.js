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
    수령인주소전체: '수령인 주소(전체)', 배송메시지: '배송메시지'
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
  context.Utilities = { formatDate: (_date, _tz, format) => format === 'yyyy-MM-dd' ? '2026-08-19' : '2026-08-19 10:00' };
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
      orderRows.push([order.orderNo, itemNo, order.orderDate || `2026-08-${String(orderIndex + 1).padStart(2, '0')} 09:00`,
        order.recipient || '', order.phone || '', order.postal || '', order.address || '', order.message || '']);
      lineRows.push([instructionNo, order.orderNo, itemNo, item.sku, item.location || '', item.name, item.option || '', item.quantity, '미처리']);
    });
  });
  return {
    header: table(['피킹지시번호', '주문번호', '상태', '출력일시'], headerRows),
    orders: table(['주문번호', '품목별 주문번호', '주문일시', '수령인', '수령인 휴대전화', '수령인 우편번호', '수령인 주소(전체)', '배송메시지'], orderRows),
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
  const context = createContext(documentTables([{ orderNo: 'A001', orderDate: '2026-08-01 08:30',
    recipient: '김희성', phone: '010-9876-1234', postal: '01234', address: '서울시 중구', message: '문 앞',
    items: [{ sku: 'BOOK', location: 'A-01', name: '만화책', quantity: 2 }] }]));
  const order = context.buildPickingDocumentData_('PK-1').orders[0];
  assert.equal(order.orderNo, 'A001');
  assert.equal(order.recipient, '김희성');
  assert.equal(order.phoneLast4, '1234');
  assert.equal(order.postalCode, '01234');
  assert.equal(order.items[0].quantity, 2);
  assert.equal(JSON.stringify(order).includes('010-9876-1234'), false);
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
