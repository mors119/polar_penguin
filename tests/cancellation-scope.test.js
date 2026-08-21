const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const context = vm.createContext({ console });
for (const file of ['S0 공통.js', 'S7 주문취소.js']) {
  vm.runInContext(fs.readFileSync(path.join(root, 'src', file), 'utf8'), context);
}
const realCancelOrder = context.cancelOrder_;
const realCancelOrderItems = context.cancelOrderItems_;

function table(role, headers, rows) {
  const result = {
    role, headers, rows, width: headers.length,
    headerIndex: Object.fromEntries(headers.map((header, index) => [context.normKey_(header), index])),
    sheet: { getRange: (row, col) => ({ setValues(values) {
      values.forEach((line, r) => line.forEach((value, c) => { result.rows[row - 2 + r][col - 1 + c] = value; }));
      return this;
    } }) }
  };
  return result;
}

const orderHeaders = [
  '주문번호', '품목별 주문번호', '상품품목코드', '주문상품명(기본)', '상품옵션(기본)', '수량',
  '주문상태', '출고완료', '피킹지시번호', '취소사유', '취소일시', '취소경로', '확정일시', '대기사유'
];

function orderRow(no, itemNo, sku, quantity, state = '예약', instruction = '', options = {}) {
  return [no, itemNo, sku, options.name || `상품 ${sku}`, options.option || '', quantity, state,
    state === '출고완료' ? 1 : 0, instruction, '', '', '',
    options.confirmed === false ? '' : (state === '예약' ? new Date('2026-08-21T00:00:00Z') : ''), options.wait || ''];
}

function fixtures(orderRows, { masterRows, headerRows = [], lineRows = [] } = {}) {
  const skus = [...new Set(orderRows.map(row => row[2]))];
  return {
    [context.ROLE.주문]: table(context.ROLE.주문, orderHeaders, orderRows),
    [context.ROLE.마스터]: table(context.ROLE.마스터, ['상품품목코드', '가용재고'],
      masterRows || skus.map(sku => [sku, 100])),
    [context.ROLE.헤더]: table(context.ROLE.헤더, ['피킹지시번호', '주문번호', '상태'], headerRows),
    [context.ROLE.라인]: table(context.ROLE.라인,
      ['주문번호', '품목별 주문번호', '상품코드', '피킹지시번호', '실제수량', '라인상태', '처리일시'], lineRows)
  };
}

function install(tables) {
  context.getConfig_ = () => ({ 별칭: {} });
  context.readTable_ = role => tables[role];
  context.writeColumn_ = () => {};
  context.writeStockLog_ = logs => { install.stockLogs.push(...logs); };
  context.writeOpLog_ = (name, state, detail) => { install.opLogs.push({ name, state, detail }); };
  context.sendSystemNotification_ = () => ({ sent: true });
  context.사용자_ = () => 'tester';
  context.withLock_ = fn => fn();
  context.D0_대시보드전체갱신 = () => {};
  context.cancelOrder_ = realCancelOrder;
  context.cancelOrderItems_ = realCancelOrderItems;
  install.stockLogs = [];
  install.opLogs = [];
  return tables;
}

test('modal exposes only ITEMS/ORDER, preserves state on RPC failure, and passes selected item context', () => {
  const html = fs.readFileSync(path.join(root, 'src', 'CancellationScope.html'), 'utf8');
  assert.match(html, /name="scope" value="ITEMS" checked/);
  assert.match(html, /name="scope" value="ORDER"/);
  assert.match(html, /selectedItemOrderNos:itemNos/);
  assert.match(html, /withSuccessHandler\(renderResult\)\.withFailureHandler\(fail\)\.executeCancellation/);
  assert.match(html, /withSuccessHandler\(render\)\.withFailureHandler\(fail\)\.getCancellationContext\(selectedOrderNo,selectedItemOrderNo\)/);
  assert.match(html, /btn\.disabled=!context\|\|context\.cancellableItemCount===0/);
  assert.match(html, /if\(scope==='ITEMS'&&!itemNos\.length\)/);
});

test('context returns every unique item in the current order, aggregates duplicate rows, and preselects the entry item', () => {
  const tables = install(fixtures([
    orderRow('A001', 'A001-01', 'SKU-A', 1, '예약', '', { option: '빨강' }),
    orderRow('A001', 'A001-02', 'SKU-B', 1),
    orderRow('A001', 'A001-02', 'SKU-B', 2),
    orderRow('A001', 'A001-03', 'SKU-C', 1),
    orderRow('B001', 'B001-01', 'SKU-X', 9)
  ]));
  const result = context.getCancellationContext_('A001', 'A001-02');
  assert.deepEqual(Array.from(result.items, item => item.itemOrderNo), ['A001-01', 'A001-02', 'A001-03']);
  assert.equal(result.itemCount, 3);
  assert.equal(result.totalQuantity, 5);
  assert.equal(result.items[1].quantity, 3);
  assert.equal(result.items[1].selected, true);
  assert.equal(result.items[0].option, '빨강');
  assert.equal(tables[context.ROLE.주문].rows.length, 5);
  assert.equal(Object.prototype.hasOwnProperty.call(result, 'recipient'), false);
});

test('context keeps cancelled items visible and disabled while reporting active totals', () => {
  install(fixtures([
    orderRow('A001', 'A001-01', 'SKU-A', 1),
    orderRow('A001', 'A001-02', 'SKU-B', 2, '취소'),
    orderRow('A001', 'A001-03', 'SKU-C', 3)
  ]));
  const result = context.getCancellationContext_('A001', 'A001-02');
  assert.equal(result.itemCount, 3);
  assert.equal(result.totalQuantity, 6);
  assert.equal(result.cancellableItemCount, 2);
  assert.equal(result.cancellableQuantity, 4);
  assert.equal(result.items[1].cancellable, false);
  assert.equal(result.items[1].state, '취소');
});

test('ITEMS cancels only one selected logical item and restores only its committed quantity', () => {
  const tables = install(fixtures([
    orderRow('A001', 'A001-01', 'SKU-A', 1),
    orderRow('A001', 'A001-02', 'SKU-B', 2),
    orderRow('A001', 'A001-03', 'SKU-C', 3)
  ], { masterRows: [['SKU-A', 10], ['SKU-B', 20], ['SKU-C', 30]] }));
  const result = context.executeCancellation_({ selectedOrderNo: 'A001', scope: 'ITEMS',
    selectedItemOrderNos: ['A001-02'], reason: '중복 주문' });
  assert.deepEqual(Array.from(result.successItems), ['A001-02']);
  assert.deepEqual(tables[context.ROLE.주문].rows.map(row => row[6]), ['예약', '취소', '예약']);
  assert.deepEqual(tables[context.ROLE.마스터].rows.map(row => row[1]), [10, 22, 30]);
  assert.equal(tables[context.ROLE.주문].rows[1][11], 'MENU_ITEM_CANCEL');
  assert.equal(result.restoredQuantity, 2);
});

test('ITEMS cancels multiple item IDs without using SKU as identity and can cancel the final active items', () => {
  const tables = install(fixtures([
    orderRow('A001', 'A001-01', 'SKU-SAME', 1),
    orderRow('A001', 'A001-02', 'SKU-SAME', 2, '취소'),
    orderRow('A001', 'A001-03', 'SKU-SAME', 3)
  ], { masterRows: [['SKU-SAME', 50]] }));
  const result = context.executeCancellation_({ selectedOrderNo: 'A001', scope: 'ITEMS',
    selectedItemOrderNos: ['A001-01', 'A001-03'], reason: '고객 요청' });
  assert.deepEqual(Array.from(result.successItems), ['A001-01', 'A001-03']);
  assert.deepEqual(tables[context.ROLE.주문].rows.map(row => row[6]), ['취소', '취소', '취소']);
  assert.equal(tables[context.ROLE.마스터].rows[0][1], 54);
});

test('ITEMS rejects injected IDs from another order without touching that order', () => {
  const tables = install(fixtures([
    orderRow('A001', 'A001-01', 'SKU-A', 1),
    orderRow('B001', 'B001-01', 'SKU-B', 4)
  ]));
  const result = context.executeCancellation_({ selectedOrderNo: 'A001', scope: 'ITEMS',
    selectedItemOrderNos: ['A001-01', 'B001-01'], reason: '고객 요청' });
  assert.deepEqual(Array.from(result.successItems), ['A001-01']);
  assert.equal(result.failedItems.length, 1);
  assert.equal(result.failedItems[0].itemOrderNo, 'B001-01');
  assert.deepEqual(tables[context.ROLE.주문].rows.map(row => row[6]), ['취소', '예약']);
});

test('repeated ITEMS cancellation is inventory-idempotent and reports already cancelled', () => {
  const tables = install(fixtures([orderRow('A001', 'A001-01', 'SKU-A', 2)], { masterRows: [['SKU-A', 10]] }));
  const payload = { selectedOrderNo: 'A001', scope: 'ITEMS', selectedItemOrderNos: ['A001-01'], reason: '고객 요청' };
  assert.equal(context.executeCancellation_(payload).restoredQuantity, 2);
  const second = context.executeCancellation_(payload);
  assert.deepEqual(Array.from(second.alreadyCancelledItems), ['A001-01']);
  assert.equal(second.restoredQuantity, 0);
  assert.equal(tables[context.ROLE.마스터].rows[0][1], 12);
});

test('completed ITEMS requires confirmation and restores authoritative actual shipped quantity', () => {
  const tables = install(fixtures([
    orderRow('A001', 'A001-01', 'SKU-A', 5, '출고완료', 'PK-1'),
    orderRow('A001', 'A001-02', 'SKU-B', 7, '출고완료', 'PK-1')
  ], { masterRows: [['SKU-A', 10], ['SKU-B', 20]], headerRows: [['PK-1', 'A001', '완료']], lineRows: [
    ['A001', 'A001-01', 'SKU-A', 'PK-1', 3, '완료', ''],
    ['A001', 'A001-02', 'SKU-B', 'PK-1', 6, '완료', '']
  ] }));
  const payload = { selectedOrderNo: 'A001', scope: 'ITEMS', selectedItemOrderNos: ['A001-01'], reason: '고객 요청' };
  assert.throws(() => context.executeCancellation_(payload), /재고 복원 확인/);
  payload.confirmCompleted = true;
  const result = context.executeCancellation_(payload);
  assert.equal(result.restoredQuantity, 3);
  assert.deepEqual(tables[context.ROLE.마스터].rows.map(row => row[1]), [13, 20]);
});

test('ITEMS cancels only matching picking lines and preserves a shared header while any line remains', () => {
  const tables = install(fixtures([
    orderRow('A001', 'A001-01', 'SKU-A', 1, '예약', 'PK-SHARED'),
    orderRow('A001', 'A001-02', 'SKU-B', 2, '예약', 'PK-SHARED'),
    orderRow('B001', 'B001-01', 'SKU-C', 3, '예약', 'PK-SHARED')
  ], { headerRows: [['PK-SHARED', 'A001', '대기'], ['PK-SHARED', 'B001', '대기']], lineRows: [
    ['A001', 'A001-01', 'SKU-A', 'PK-SHARED', '', '미처리', ''],
    ['A001', 'A001-02', 'SKU-B', 'PK-SHARED', '', '미처리', ''],
    ['B001', 'B001-01', 'SKU-C', 'PK-SHARED', '', '미처리', '']
  ] }));
  context.executeCancellation_({ selectedOrderNo: 'A001', scope: 'ITEMS',
    selectedItemOrderNos: ['A001-02'], reason: '고객 요청' });
  assert.deepEqual(tables[context.ROLE.라인].rows.map(row => row[5]), ['미처리', '취소', '미처리']);
  assert.deepEqual(tables[context.ROLE.헤더].rows.map(row => row[2]), ['대기', '대기']);
  assert.equal(tables[context.ROLE.주문].rows[2][6], '예약');
});

test('picking header is cancelled only after every line in the instruction is cancelled', () => {
  const tables = install(fixtures([
    orderRow('A001', 'A001-01', 'SKU-A', 1, '예약', 'PK-1'),
    orderRow('A001', 'A001-02', 'SKU-B', 2, '예약', 'PK-1')
  ], { headerRows: [['PK-1', 'A001', '대기']], lineRows: [
    ['A001', 'A001-01', 'SKU-A', 'PK-1', '', '미처리', ''],
    ['A001', 'A001-02', 'SKU-B', 'PK-1', '', '미처리', '']
  ] }));
  context.executeCancellation_({ selectedOrderNo: 'A001', scope: 'ITEMS', selectedItemOrderNos: ['A001-01'], reason: '고객 요청' });
  assert.equal(tables[context.ROLE.헤더].rows[0][2], '대기');
  context.executeCancellation_({ selectedOrderNo: 'A001', scope: 'ITEMS', selectedItemOrderNos: ['A001-02'], reason: '고객 요청' });
  assert.equal(tables[context.ROLE.헤더].rows[0][2], '취소');
});

test('ORDER scope delegates to cancelOrder_ with the whole-order source and ignores checkbox subset', () => {
  install(fixtures([orderRow('A001', 'A001-01', 'SKU-A', 1), orderRow('A001', 'A001-02', 'SKU-B', 2)]));
  const calls = [];
  context.cancelOrder_ = (orderNo, reason, source, options) => {
    calls.push({ orderNo, reason, source, options });
    return { 취소: true, 복원수량: 3 };
  };
  const result = context.executeCancellation_({ selectedOrderNo: 'A001', scope: 'ORDER',
    selectedItemOrderNos: ['A001-01'], reason: '판매자 취소', confirmCompleted: true });
  assert.equal(calls.length, 1);
  assert.equal(calls[0].source, 'MENU_ORDER_CANCEL');
  assert.equal(calls[0].options.confirmReturn, true);
  assert.equal(result.scope, 'ORDER');
  assert.equal(result.restoredQuantity, 3);
});

test('real ORDER cancellation handles a partially cancelled order and keeps a shared instruction active', () => {
  const tables = install(fixtures([
    orderRow('A001', 'A001-01', 'SKU-A', 1, '취소', 'PK-SHARED'),
    orderRow('A001', 'A001-02', 'SKU-B', 2, '예약', 'PK-SHARED'),
    orderRow('B001', 'B001-01', 'SKU-C', 3, '예약', 'PK-SHARED')
  ], { headerRows: [['PK-SHARED', 'A001', '대기'], ['PK-SHARED', 'B001', '대기']], lineRows: [
    ['A001', 'A001-01', 'SKU-A', 'PK-SHARED', '', '취소', ''],
    ['A001', 'A001-02', 'SKU-B', 'PK-SHARED', '', '미처리', ''],
    ['B001', 'B001-01', 'SKU-C', 'PK-SHARED', '', '미처리', '']
  ] }));
  const result = context.executeCancellation_({ selectedOrderNo: 'A001', scope: 'ORDER', reason: '고객 요청' });
  assert.deepEqual(tables[context.ROLE.주문].rows.map(row => row[6]), ['취소', '취소', '예약']);
  assert.deepEqual(tables[context.ROLE.라인].rows.map(row => row[5]), ['취소', '취소', '미처리']);
  assert.deepEqual(tables[context.ROLE.헤더].rows.map(row => row[2]), ['대기', '대기']);
  assert.equal(result.successItems.length, 1);
});

test('invalid scopes and empty item selections are rejected', () => {
  install(fixtures([orderRow('A001', 'A001-01', 'SKU-A', 1)]));
  assert.throws(() => context.executeCancellation_({ selectedOrderNo: 'A001', scope: 'UNKNOWN', reason: '고객 요청' }), /올바른 취소 범위/);
  assert.throws(() => context.executeCancellation_({ selectedOrderNo: 'A001', scope: 'ITEMS', selectedItemOrderNos: [], reason: '고객 요청' }), /하나 이상 선택/);
});
