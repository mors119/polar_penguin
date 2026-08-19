const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const context = vm.createContext({ console });
for (const file of ['S0 공통.js', 'S3 주문확정.js', 'S8 예약관리.js']) {
  vm.runInContext(fs.readFileSync(path.join(root, 'src', file), 'utf8'), context);
}
const plain = value => JSON.parse(JSON.stringify(value));
const inventory = values => Object.fromEntries(Object.entries(values).map(([sku, available]) =>
  [sku, { available, name: sku, option: '', disabled: false }]));
const order = (orderNo, date, required, itemOrderNo = `${orderNo}-1`) =>
  ({ orderNo, orderDate: date, itemOrderNo, required, recipient: '김희성 · 01012345678' });

test('basic FIFO selects complete orders and does not partially release the next order', () => {
  const result = context.calculateReservationBatch_('BOOK', inventory({ BOOK: 100 }), [
    order('A001', '2026-08-01T08:00:00Z', { BOOK: 30 }),
    order('A002', '2026-08-01T09:00:00Z', { BOOK: 20 }),
    order('A003', '2026-08-02T08:00:00Z', { BOOK: 40 }),
    order('A004', '2026-08-03T08:00:00Z', { BOOK: 30 })
  ]);
  assert.deepEqual(plain(result.selected.map(item => item.orderNo)), ['A001', 'A002', 'A003']);
  assert.equal(result.targetAllocated, 90);
  assert.equal(result.remaining.BOOK, 10);
  assert.match(result.waiting[0].reason, /BOOK 재고 부족/);
});

test('target SKU shortage stops FIFO and prevents a smaller newer order leapfrog', () => {
  const result = context.calculateReservationBatch_('BOOK', inventory({ BOOK: 10 }), [
    order('A001', '2026-08-01', { BOOK: 20 }), order('A002', '2026-08-02', { BOOK: 5 })
  ]);
  assert.equal(result.selected.length, 0);
  assert.match(result.waiting[1].reason, /FIFO 대기/);
});

test('other-SKU shortage is skipped without consuming target stock', () => {
  const result = context.calculateReservationBatch_('BOOK', inventory({ BOOK: 10, B: 0 }), [
    order('A001', '2026-08-01', { BOOK: 2, B: 1 }), order('A002', '2026-08-02', { BOOK: 2 })
  ]);
  assert.deepEqual(plain(result.selected.map(item => item.orderNo)), ['A002']);
  assert.equal(result.remaining.BOOK, 8);
  assert.match(result.waiting[0].reason, /다른 상품 재고 부족: B 필요 1 \/ 가용 0/);
});

test('batch allocation consumes shared secondary inventory cumulatively', () => {
  const result = context.calculateReservationBatch_('BOOK', inventory({ BOOK: 10, B: 5 }), [
    order('A001', '2026-08-01', { BOOK: 2, B: 3 }), order('A002', '2026-08-02', { BOOK: 2, B: 3 })
  ]);
  assert.deepEqual(plain(result.selected.map(item => item.orderNo)), ['A001']);
  assert.equal(result.remaining.B, 2);
  assert.match(result.waiting[0].reason, /B 필요 3 \/ 가용 2/);
});

test('duplicate SKU lines are aggregated before FIFO evaluation', () => {
  const result = context.calculateReservationBatch_('BOOK', inventory({ BOOK: 4 }), [
    order('A001', '2026-08-01', { BOOK: 5 })
  ]);
  assert.equal(result.selected.length, 0);
  assert.equal(result.waiting[0].targetQuantity, 5);
});

test('an order with multiple reservation products releases only when every SKU is available', () => {
  const blocked = context.calculateReservationBatch_('BOOK', inventory({ BOOK: 5, BONUS: 0 }), [
    order('A001', '2026-08-01', { BOOK: 2, BONUS: 1 })
  ]);
  assert.equal(blocked.selected.length, 0);
  const released = context.calculateReservationBatch_('BOOK', inventory({ BOOK: 5, BONUS: 1 }), [
    order('A001', '2026-08-01', { BOOK: 2, BONUS: 1 })
  ]);
  assert.deepEqual(plain(released.selected.map(item => item.orderNo)), ['A001']);
  assert.equal(released.totalAllocated, 3);
});

function table(role, headers, rows) {
  const result = { role, headers, rows, width: headers.length,
    headerIndex: Object.fromEntries(headers.map((header, index) => [context.normKey_(header), index])),
    sheet: { getRange: (row, col) => ({ setValues(values) {
      values.forEach((line, r) => line.forEach((value, c) => { result.rows[row - 2 + r][col - 1 + c] = value; }));
    } }) } };
  return result;
}

function installConfirmationTables(orderRows) {
  const tables = {
    [context.ROLE.주문]: table(context.ROLE.주문,
      ['주문번호', '상품품목코드', '수량', '주문상태', '확정일시', '대기사유'], orderRows),
    [context.ROLE.마스터]: table(context.ROLE.마스터,
      ['상품품목코드', '가용재고', '예약재고', '상품상태', '예약상품'],
      [['NORMAL', 10, 0, '판매중', 'N'], ['BOOK', 10, 0, '판매중', 'Y']])
  };
  context.readTable_ = role => tables[role];
  context.writeColumn_ = () => {};
  context.writeStockLog_ = () => {};
  context.writeOpLog_ = () => {};
  context.alert_ = () => {};
  context.사용자_ = () => 'tester';
  context.withLock_ = fn => fn();
  return tables;
}

test('same customer separate normal and reservation orders remain independent fulfillment units', () => {
  const tables = installConfirmationTables([
    ['NORMAL-ORDER', 'NORMAL', 2, '예약', '', ''],
    ['RES-ORDER', 'BOOK', 3, '예약', '', '']
  ]);
  const result = context.S3_1_주문확정(['NORMAL-ORDER', 'RES-ORDER'], { silent: true });
  assert.deepEqual(plain(result.준비주문), ['NORMAL-ORDER']);
  assert.equal(tables[context.ROLE.마스터].rows[0][1], 8);
  assert.equal(tables[context.ROLE.마스터].rows[1][1], 10);
});

test('mixed normal and reservation items inside one order never reserve partially', () => {
  const tables = installConfirmationTables([
    ['MIXED', 'NORMAL', 2, '예약', '', ''], ['MIXED', 'BOOK', 1, '예약', '', '']
  ]);
  const result = context.S3_1_주문확정(['MIXED'], { silent: true });
  assert.equal(result.준비주문.length, 0);
  assert.equal(tables[context.ROLE.마스터].rows[0][1], 10);
  assert.equal(tables[context.ROLE.마스터].rows[0][2], 0);
});

test('generation revalidates under lock and never accepts browser-selected order IDs', () => {
  let locked = false;
  context.withLock_ = fn => { locked = true; return fn(); };
  context.toStr_ = value => String(value || '').trim();
  context.readReservationSnapshot_ = () => ({
    reservationProducts: { BOOK: true }, inventory: inventory({ BOOK: 0 }),
    orders: [order('A001', '2026-08-01', { BOOK: 1 })]
  });
  const result = context.createReservationPickingBatch_('BOOK');
  assert.equal(locked, true);
  assert.equal(result.created, false);
  assert.equal(result.preview.selected.length, 0);
});

test('successful generation makes the order immediately ineligible for a second batch', () => {
  let eligible = true;
  let pickingCalls = 0;
  context.withLock_ = fn => fn();
  context.toStr_ = value => String(value || '').trim();
  context.readReservationSnapshot_ = () => ({
    reservationProducts: { BOOK: true }, inventory: inventory({ BOOK: eligible ? 2 : 1 }),
    orders: eligible ? [order('A001', '2026-08-01', { BOOK: 1 })] : []
  });
  context.S3_1_주문확정 = () => ({ 준비주문: ['A001'] });
  context.S4_1_피킹지시생성 = () => { pickingCalls += 1; return { 지시번호: 'PK-20260819-RES-001' }; };
  context.S9_피킹PDF생성 = () => ({ 생성: true });
  context.inputFolder_ = () => ({});
  context.markPickingOutputState_ = () => {};
  context.markOrdersReady_ = () => { eligible = false; };
  context.D0_대시보드전체갱신 = () => {};
  context.writeOpLog_ = () => {};
  assert.equal(context.createReservationPickingBatch_('BOOK').created, true);
  assert.equal(context.createReservationPickingBatch_('BOOK').created, false);
  assert.equal(pickingCalls, 1);
});
