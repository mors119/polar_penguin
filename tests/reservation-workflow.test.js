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
const originalReadReservationSnapshot = context.readReservationSnapshot_;
const originalReadReservationPhysicalCarry = context.readReservationPhysicalCarry_;
const originalWriteReservationPhysicalCarry = context.writeReservationPhysicalCarry_;
const plain = value => JSON.parse(JSON.stringify(value));
const inventory = values => Object.fromEntries(Object.entries(values).map(([sku, available]) =>
  [sku, { available, name: sku, option: '', managed: true }]));
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
  assert.match(result.waiting[0].reason, /BOOK 물리 재고 부족/);
});

test('target SKU shortage stops FIFO and prevents a smaller newer order leapfrog', () => {
  const result = context.calculateReservationBatch_('BOOK', inventory({ BOOK: 10 }), [
    order('A001', '2026-08-01', { BOOK: 20 }), order('A002', '2026-08-02', { BOOK: 5 })
  ]);
  assert.equal(result.selected.length, 0);
  assert.match(result.waiting[1].reason, /FIFO 대기/);
});

test('new inbound releases complete FIFO orders while net inventory remains negative', () => {
  const snapshot = {
    inventory: { BOOK: { available: -20, name: '책', option: '', managed: false } },
    orders: [
      { ...order('A001', '2026-08-01', { BOOK: 4 }), committed: true },
      { ...order('A002', '2026-08-02', { BOOK: 3 }), committed: true },
      { ...order('A003', '2026-08-03', { BOOK: 5 }), committed: true },
      { ...order('A004', '2026-08-04', { BOOK: 2 }), committed: true }
    ]
  };
  const preview = context.buildReservationInboundPreview_('BOOK', 10, snapshot);
  assert.equal(preview.inventoryAfterInbound, -10);
  assert.deepEqual(plain(preview.selected.map(item => item.orderNo)), ['A001', 'A002']);
  assert.equal(preview.releaseQuantity, 7);
  assert.equal(preview.unusedInboundQuantity, 3);
  assert.match(preview.waiting[1].reason, /FIFO 대기/);
});

test('oldest oversized order blocks a smaller newer order from using inbound stock', () => {
  const result = context.calculateReservationBatch_('BOOK', 10, inventory({ BOOK: -10 }), [
    order('A001', '2026-08-01', { BOOK: 15 }), order('A002', '2026-08-02', { BOOK: 5 })
  ]);
  assert.equal(result.selected.length, 0);
  assert.equal(result.unusedTargetQuantity, 10);
  assert.match(result.waiting[1].reason, /FIFO 대기/);
});

test('unused secured stock combines with the next inbound event', () => {
  const props = new Map([[`${context.RESERVATION_PHYSICAL_CARRY_PREFIX}BOOK`, '3']]);
  context.PropertiesService = { getScriptProperties: () => ({
    getProperty: key => props.get(key) || null,
    setProperty: (key, value) => props.set(key, value),
    deleteProperty: key => props.delete(key)
  }) };
  const snapshot = {
    inventory: { BOOK: { available: -10, name: '책', option: '', managed: false } },
    orders: [{ ...order('A003', '2026-08-03', { BOOK: 5 }), committed: true }]
  };
  const preview = context.buildReservationInboundPreview_('BOOK', 2, snapshot);
  assert.equal(preview.existingSecuredQuantity, 3);
  assert.deepEqual(plain(preview.selected.map(item => item.orderNo)), ['A003']);
  assert.equal(preview.releaseQuantity, 5);
  assert.equal(preview.unusedInboundQuantity, 0);
});

test('waiting order count and waiting quantity remain distinct', () => {
  context.PropertiesService = undefined;
  const snapshot = {
    candidateSkus: { BOOK: true },
    inventory: { BOOK: { available: -20, name: '책', option: '', managed: false } },
    orders: [
      { ...order('A001', '2026-08-01', { BOOK: 5 }), committed: true },
      { ...order('A002', '2026-08-02', { BOOK: 3 }), committed: true },
      { ...order('A003', '2026-08-03', { BOOK: 12 }), committed: true }
    ]
  };
  const summary = context.reservationSkuSummaries_(snapshot, true)[0];
  assert.equal(summary.reservationOrderCount, 3);
  assert.equal(summary.reservationQuantity, 20);
});

test('manual inbound quantity validation accepts only nonnegative integers', () => {
  assert.equal(context.validateManualReservationInboundQuantity_(0), 0);
  assert.equal(context.validateManualReservationInboundQuantity_('10'), 10);
  for (const invalid of [-1, 'abc', NaN, 1.5, '']) {
    assert.throws(() => context.validateManualReservationInboundQuantity_(invalid), /0 이상의 정수/);
  }
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

test('an order with multiple waiting SKUs releases only when every managed SKU is available', () => {
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
    sheet: { getRange: (row, col) => ({
      setValue(value) { result.rows[row - 2][col - 1] = value; return this; },
      setValues(values) {
        values.forEach((line, r) => line.forEach((value, c) => { result.rows[row - 2 + r][col - 1 + c] = value; }));
        return this;
      }
    }) } };
  return result;
}

function installConfirmationTables(orderRows) {
  const tables = {
    [context.ROLE.주문]: table(context.ROLE.주문,
      ['주문번호', '상품품목코드', '수량', '주문상태', '확정일시', '대기사유'], orderRows),
    [context.ROLE.마스터]: table(context.ROLE.마스터,
      ['상품품목코드', '가용재고', '재고관리'],
      [['NORMAL', 10, 'T'], ['BOOK', 10, 'T']])
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

test('reservation queue candidates come from waiting order rows without product flags', () => {
  const tables = {
    [context.ROLE.주문]: table(context.ROLE.주문,
      ['주문번호', '품목별 주문번호', '상품품목코드', '수량', '주문상태', '피킹지시번호',
        '수령인', '수령인 휴대전화'],
      [['WAIT-1', 'ITEM-1', 'BOOK', 2, '예약', '', '김희성', '01012345678']]),
    [context.ROLE.마스터]: table(context.ROLE.마스터,
      ['상품품목코드', '상품명', '옵션명', '가용재고', '재고관리'],
      [['BOOK', '책', '-', 5, 'T']])
  };
  context.readTable_ = role => tables[role];
  context.getConfig_ = () => ({ 별칭: {} });
  const snapshot = context.readReservationSnapshot_();
  assert.deepEqual(plain(Object.keys(snapshot.candidateSkus)), ['BOOK']);
  assert.equal(snapshot.orders[0].required.BOOK, 2);
});

test('separate orders remain independent fulfillment units without a product reservation flag', () => {
  const tables = installConfirmationTables([
    ['NORMAL-ORDER', 'NORMAL', 2, '예약', '', ''],
    ['RES-ORDER', 'BOOK', 3, '예약', '', '']
  ]);
  const result = context.S3_1_주문확정(['NORMAL-ORDER', 'RES-ORDER'], { silent: true });
  assert.deepEqual(plain(result.준비주문), ['NORMAL-ORDER', 'RES-ORDER']);
  assert.equal(tables[context.ROLE.마스터].rows[0][1], 8);
  assert.equal(tables[context.ROLE.마스터].rows[1][1], 7);
});

test('mixed items inside one order commit atomically when every managed SKU is available', () => {
  const tables = installConfirmationTables([
    ['MIXED', 'NORMAL', 2, '예약', '', ''], ['MIXED', 'BOOK', 1, '예약', '', '']
  ]);
  const result = context.S3_1_주문확정(['MIXED'], { silent: true });
  assert.equal(result.준비주문.length, 1);
  assert.equal(tables[context.ROLE.마스터].rows[0][1], 8);
  assert.equal(tables[context.ROLE.마스터].rows[1][1], 9);
});

test('manual F-product inbound writes one delta movement and repeated submission is idempotent', () => {
  const tables = {
    [context.ROLE.주문]: table(context.ROLE.주문,
      ['주문번호', '품목별 주문번호', '상품품목코드', '수량', '주문상태', '피킹지시번호',
        '확정일시', '수령인', '수령인 휴대전화'], []),
    [context.ROLE.마스터]: table(context.ROLE.마스터,
      ['상품품목코드', '상품명', '옵션명', '가용재고', '재고관리'],
      [['BOOK', '책', '-', -20, 'F']])
  };
  const props = new Map(), stockLogs = [];
  context.readTable_ = role => tables[role];
  context.PropertiesService = { getScriptProperties: () => ({
    getProperty: key => props.get(key) || null,
    setProperty: (key, value) => { props.set(key, value); },
    deleteProperty: key => { props.delete(key); }
  }) };
  context.withLock_ = fn => fn();
  context.writeStockLog_ = logs => stockLogs.push(...logs);
  context.writeOpLog_ = () => {};
  context.사용자_ = () => 'tester@example.com';
  context.inputFolder_ = () => ({});
  const requestId = 'request_1234567890';
  const first = context.applyManualReservationInboundCore_('BOOK', 10, requestId);
  const second = context.applyManualReservationInboundCore_('BOOK', 10, requestId);
  assert.equal(tables[context.ROLE.마스터].rows[0][3], -10);
  assert.equal(first.inventoryAfterInbound, -10);
  assert.equal(second.duplicate, true);
  assert.equal(stockLogs.length, 1);
  assert.equal(stockLogs[0].변동량, 10);
  assert.equal(stockLogs[0].변동후재고, -10);
  assert.equal(stockLogs[0].담당자, 'tester@example.com');
  assert.match(stockLogs[0].사유, /예약상품 수동 입고.*MANUAL_RESERVATION_INBOUND/);

  context.applyManualReservationInboundCore_('BOOK', 0, 'request_zero_123456');
  assert.equal(stockLogs.length, 1, 'zero re-evaluation must not create a stock movement');
});

test('invalid or T-product manual inbound never mutates inventory', () => {
  assert.throws(() => context.applyManualReservationInboundCore_('BOOK', -1, 'request_invalid_123'), /0 이상의 정수/);
  const master = context.readTable_(context.ROLE.마스터);
  master.rows[0][4] = 'T';
  assert.throws(() => context.applyManualReservationInboundCore_('BOOK', 3, 'request_tproduct_123'), /재고관리=F/);
  assert.equal(master.rows[0][3], -10);
});

test('generation revalidates under lock and never accepts browser-selected order IDs', () => {
  let locked = false;
  context.withLock_ = fn => { locked = true; return fn(); };
  context.toStr_ = value => String(value || '').trim();
  context.readReservationSnapshot_ = () => ({
    candidateSkus: { BOOK: true }, inventory: { BOOK: { available: 0, name: 'BOOK', option: '', managed: false } },
    orders: [order('A001', '2026-08-01', { BOOK: 1 })]
  });
  context.writeReservationPhysicalCarry_ = () => {};
  context.readReservationPhysicalCarry_ = () => 0;
  context.writeOpLog_ = () => {};
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
    candidateSkus: { BOOK: true }, inventory: { BOOK: { available: eligible ? 2 : 1, name: 'BOOK', option: '', managed: false } },
    orders: eligible ? [order('A001', '2026-08-01', { BOOK: 1 })] : []
  });
  context.S3_1_주문확정 = () => ({ 준비주문: ['A001'] });
  context.S4_1_피킹지시생성 = () => { pickingCalls += 1; return { 지시번호: 'PK-20260819-RES-001' }; };
  context.S9_피킹PDF생성 = () => { eligible = false; return { 생성: true, 출고확정: { 완료주문: ['A001'] } }; };
  context.inputFolder_ = () => ({});
  context.markPickingOutputState_ = () => {};
  context.D0_대시보드전체갱신 = () => {};
  context.writeOpLog_ = () => {};
  context.writeReservationPhysicalCarry_ = () => {};
  assert.equal(context.createReservationPickingBatch_('BOOK').created, true);
  assert.equal(context.createReservationPickingBatch_('BOOK').created, false);
  assert.equal(pickingCalls, 1);
});

test('PDF-stage failure preserves inbound idempotency and a recoverable instruction number', () => {
  const tables = {
    [context.ROLE.주문]: table(context.ROLE.주문,
      ['주문번호', '품목별 주문번호', '상품품목코드', '수량', '주문상태', '피킹지시번호',
        '확정일시', '수령인', '수령인 휴대전화'], []),
    [context.ROLE.마스터]: table(context.ROLE.마스터,
      ['상품품목코드', '상품명', '옵션명', '가용재고', '재고관리'], [['BOOK', '책', '-', -20, 'F']])
  };
  const props = new Map(), logs = [];
  context.readTable_ = role => tables[role];
  context.readReservationSnapshot_ = originalReadReservationSnapshot;
  context.readReservationPhysicalCarry_ = originalReadReservationPhysicalCarry;
  context.writeReservationPhysicalCarry_ = originalWriteReservationPhysicalCarry;
  context.PropertiesService = { getScriptProperties: () => ({
    getProperty: key => props.get(key) || null,
    setProperty: (key, value) => { props.set(key, value); },
    deleteProperty: key => { props.delete(key); }
  }) };
  context.writeStockLog_ = entries => logs.push(...entries);
  context.withLock_ = fn => fn();
  context.inputFolder_ = () => ({});
  context.releaseReservationQueueCore_ = () => {
    const error = new Error('PDF failed'); error.pickingInstructionNo = 'PK-RES-ERROR-1'; throw error;
  };
  assert.throws(() => context.applyManualReservationInboundCore_('BOOK', 10, 'request_pdf_fail_123'), /PDF failed/);
  const retry = context.applyManualReservationInboundCore_('BOOK', 10, 'request_pdf_fail_123');
  assert.equal(tables[context.ROLE.마스터].rows[0][3], -10);
  assert.equal(logs.length, 1);
  assert.equal(retry.duplicate, true);
  assert.equal(retry.instructionNo, 'PK-RES-ERROR-1');
});
