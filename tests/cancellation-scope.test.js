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
const centralCancelOrder = context.cancelOrder_;

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
  '주문상태', '출고완료', '취소사유', '취소일시', '취소경로', '확정일시', '대기사유',
  '피킹지시번호', '수령인', '수령인 휴대전화'
];

function orderRow(no, itemNo, sku, name, quantity, state, recipient, phone, instruction = '') {
  return [no, itemNo, sku, name, '', quantity, state, state === '출고완료' ? 1 : 0,
    '', '', '', state === '예약' ? new Date('2026-08-21T00:00:00Z') : '', '', instruction, recipient, phone];
}

function install(tables) {
  context.getConfig_ = () => ({ 별칭: {} });
  context.readTable_ = role => tables[role];
  context.writeColumn_ = () => {};
  context.writeStockLog_ = () => {};
  context.writeOpLog_ = () => {};
  context.sendSystemNotification_ = () => ({ sent: true });
  context.사용자_ = () => 'tester';
  context.withLock_ = fn => fn();
}

test('context groups unique order numbers by normalized recipient and phone and aggregates multi-line quantities', () => {
  const orders = table(context.ROLE.주문, orderHeaders, [
    orderRow('A001', 'A001-1', 'SKU-A', '상품 A', 2, '예약', '홍 길동', '010-1111-2222'),
    orderRow('A001', 'A001-2', 'SKU-B', '상품 B', 3, '예약', '홍 길동', '010-1111-2222'),
    orderRow('A002', 'A002-1', 'SKU-C', '상품 C', 1, '예약', '홍길동', '01011112222'),
    orderRow('A003', 'A003-1', 'SKU-D', '상품 D', 4, '출고완료', '홍길동', '010 1111 2222'),
    orderRow('A004', 'A004-1', 'SKU-E', '상품 E', 7, '취소', '홍길동', '010-1111-2222'),
    orderRow('B001', 'B001-1', 'SKU-F', '상품 F', 8, '예약', '홍길동', '010-9999-8888')
  ]);
  install({ [context.ROLE.주문]: orders });

  const result = context.getCancellationContext_('A001');
  assert.equal(result.bulkAvailable, true);
  assert.equal(result.recipient.phoneMasked, '***-****-2222');
  assert.equal(Object.prototype.hasOwnProperty.call(result.selectedOrder, 'phoneKey'), false);
  assert.equal(JSON.stringify(result).includes('01011112222'), false);
  assert.deepEqual(Array.from(result.relatedOrders, order => order.orderNo), ['A001', 'A002', 'A003']);
  assert.equal(result.singleSummary.orderCount, 1);
  assert.equal(result.singleSummary.totalQuantity, 5);
  assert.equal(result.singleSummary.orders[0].items.length, 2);
  assert.equal(result.bulkSummary.orderCount, 3);
  assert.equal(result.bulkSummary.totalQuantity, 10);
  assert.equal(result.hasCompletedBulk, true);
});

test('blank recipient phone never enables name-only recipient-wide cancellation', () => {
  const orders = table(context.ROLE.주문, orderHeaders, [
    orderRow('A001', 'A001-1', 'SKU-A', '상품 A', 2, '예약', '홍길동', ''),
    orderRow('A002', 'A002-1', 'SKU-B', '상품 B', 3, '예약', '홍길동', '')
  ]);
  install({ [context.ROLE.주문]: orders });

  const result = context.getCancellationContext_('A001');
  assert.equal(result.bulkAvailable, false);
  assert.equal(result.relatedOrders.length, 0);
  assert.match(result.note, /휴대전화가 없어/);
  assert.throws(() => context.executeCancellationScope_({
    selectedOrderNo: 'A001', scope: 'RECIPIENT', reason: '고객 요청'
  }), /휴대전화가 없어/);
});

test('single related order keeps scope simple', () => {
  const orders = table(context.ROLE.주문, orderHeaders, [
    orderRow('A001', 'A001-1', 'SKU-A', '상품 A', 2, '예약', '김희성', '010-1234-5678')
  ]);
  install({ [context.ROLE.주문]: orders });
  const result = context.getCancellationContext_('A001');
  assert.equal(result.bulkAvailable, false);
  assert.equal(result.bulkSummary.orderCount, 1);
});

test('bulk execution re-derives targets, ignores a browser order list, and reports partial results', () => {
  const orders = table(context.ROLE.주문, orderHeaders, [
    orderRow('A001', 'A001-1', 'SKU-A', '상품 A', 2, '예약', '홍길동', '010-1111-2222'),
    orderRow('A002', 'A002-1', 'SKU-B', '상품 B', 3, '예약', '홍길동', '01011112222'),
    orderRow('A003', 'A003-1', 'SKU-C', '상품 C', 4, '출고완료', '홍길동', '010 1111 2222'),
    orderRow('B001', 'B001-1', 'SKU-X', '상품 X', 99, '예약', '다른 고객', '010-9999-8888')
  ]);
  install({ [context.ROLE.주문]: orders });
  const calls = [];
  let refreshes = 0;
  context.D0_대시보드전체갱신 = () => { refreshes++; };
  context.cancelOrder_ = (orderNo, reason, source, options) => {
    calls.push({ orderNo, reason, source, options });
    if (orderNo === 'A001') return { 취소: true, 복원수량: 5 };
    if (orderNo === 'A002') return { 취소: false, 이미취소: true };
    throw new Error('테스트 실패');
  };

  const result = context.executeCancellationScope_({
    selectedOrderNo: 'A001', scope: 'RECIPIENT', reason: '고객 요청', targetOrderNos: ['B001']
  });
  assert.deepEqual(calls.map(call => call.orderNo), ['A001', 'A002', 'A003']);
  assert.ok(calls.every(call => call.source === 'MENU_RECIPIENT_BULK_CANCEL'));
  assert.ok(calls.every(call => call.options.confirmReturn && call.options.refresh === false));
  assert.equal(result.requested, 3);
  assert.equal(result.success, 1);
  assert.equal(result.alreadyCancelled, 1);
  assert.equal(result.failed.length, 1);
  assert.equal(result.restoredQuantity, 5);
  assert.equal(refreshes, 1);
  context.cancelOrder_ = centralCancelOrder;
});

test('real repeated bulk cancellation is inventory-idempotent and leaves an unrelated shared instruction order valid', () => {
  const orders = table(context.ROLE.주문, orderHeaders, [
    orderRow('A001', 'A001-1', 'SKU-1', '상품', 2, '예약', '홍길동', '010-1111-2222', 'PK-SHARED'),
    orderRow('A002', 'A002-1', 'SKU-1', '상품', 3, '예약', '홍길동', '01011112222', 'PK-A002'),
    orderRow('B001', 'B001-1', 'SKU-1', '상품', 4, '예약', '다른 고객', '010-9999-8888', 'PK-SHARED')
  ]);
  const master = table(context.ROLE.마스터, ['상품품목코드', '가용재고'], [['SKU-1', 91]]);
  const headers = table(context.ROLE.헤더, ['피킹지시번호', '주문번호', '상태'], [
    ['PK-SHARED', 'A001', '대기'], ['PK-A002', 'A002', '대기'], ['PK-SHARED', 'B001', '대기']
  ]);
  const lines = table(context.ROLE.라인,
    ['주문번호', '품목별 주문번호', '피킹지시번호', '실제수량', '라인상태', '처리일시'], [
      ['A001', 'A001-1', 'PK-SHARED', '', '미처리', ''],
      ['A002', 'A002-1', 'PK-A002', '', '미처리', ''],
      ['B001', 'B001-1', 'PK-SHARED', '', '미처리', '']
    ]);
  install({
    [context.ROLE.주문]: orders, [context.ROLE.마스터]: master,
    [context.ROLE.헤더]: headers, [context.ROLE.라인]: lines
  });
  context.cancelOrder_ = centralCancelOrder;
  context.D0_대시보드전체갱신 = () => {};

  const first = context.executeCancellationScope_({ selectedOrderNo: 'A001', scope: 'RECIPIENT', reason: '고객 요청' });
  assert.equal(first.success, 2);
  assert.equal(first.restoredQuantity, 5);
  assert.equal(master.rows[0][1], 96);
  assert.deepEqual(orders.rows.map(row => row[6]), ['취소', '취소', '예약']);
  assert.deepEqual(headers.rows.map(row => row[2]), ['취소', '취소', '대기']);
  assert.deepEqual(lines.rows.map(row => row[4]), ['취소', '취소', '미처리']);

  const second = context.executeCancellationScope_({ selectedOrderNo: 'A001', scope: 'RECIPIENT', reason: '고객 요청' });
  assert.equal(second.success, 0);
  assert.equal(second.alreadyCancelled, 1);
  assert.equal(second.restoredQuantity, 0);
  assert.equal(master.rows[0][1], 96);
  assert.equal(orders.rows[2][6], '예약');
});
