const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const context = vm.createContext({ console });
for (const file of ['S0 공통.js', 'S3 주문확정.js', 'S4 피킹지시생성.js', 'S5 결과반영.js', 'S7 주문취소.js', 'S9 작업지시서.js']) {
  vm.runInContext(fs.readFileSync(path.join(root, 'src', file), 'utf8'), context);
}

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

function installTables(tables) {
  const stockLogs = [];
  context.readTable_ = (role) => tables[role];
  context.writeColumn_ = () => {};
  context.writeStockLog_ = (logs) => stockLogs.push(...logs);
  context.writeOpLog_ = () => {};
  context.alert_ = () => {};
  context.D0_대시보드전체갱신 = () => {};
  context.사용자_ = () => 'tester';
  context.withLock_ = (fn) => fn();
  return stockLogs;
}

function reservationTables(quantity = 8, available = 10, preorder = 'N') {
  return {
    [context.ROLE.주문]: table(context.ROLE.주문,
      ['주문번호', '상품품목코드', '수량', '주문상태', '확정일시', '대기사유'],
      [['O-1', 'SKU-1', 3, '예약', '', ''], ['O-1', 'SKU-1', quantity - 3, '예약', '', '']]),
    [context.ROLE.마스터]: table(context.ROLE.마스터,
      ['상품품목코드', '가용재고', '예약재고', '상품상태', '예약상품'],
      [['SKU-1', available, 0, '판매중', preorder]])
  };
}

test('normal order aggregates duplicate SKU demand and reserves inventory exactly once', () => {
  const tables = reservationTables();
  const logs = installTables(tables);
  const first = context.S3_1_주문확정(['O-1'], { silent: true });
  assert.equal(first.준비, 1);
  assert.equal(tables[context.ROLE.마스터].rows[0][1], 2);
  assert.equal(tables[context.ROLE.마스터].rows[0][2], 8);
  assert.equal(logs.length, 1);
  const second = context.S3_1_주문확정(['O-1'], { silent: true });
  assert.equal(second.준비, 1);
  assert.equal(tables[context.ROLE.마스터].rows[0][1], 2);
  assert.equal(tables[context.ROLE.마스터].rows[0][2], 8);
  assert.equal(logs.length, 1, 'retry must not reserve twice');
});

test('insufficient or preorder order remains reservation without inventory mutation', () => {
  for (const tables of [reservationTables(12, 10, 'N'), reservationTables(2, 10, 'Y')]) {
    installTables(tables);
    const result = context.S3_1_주문확정(['O-1'], { silent: true });
    assert.equal(result.예약, 1);
    assert.equal(tables[context.ROLE.마스터].rows[0][1], 10);
    assert.equal(tables[context.ROLE.마스터].rows[0][2], 0);
  }
});

test('manual reservation release requires complete stock and then reserves once', () => {
  const insufficient = reservationTables(12, 10, 'Y');
  installTables(insufficient);
  assert.equal(context.S3_1_주문확정(['O-1'], { manualRelease: true, silent: true }).준비, 0);
  assert.deepEqual([...insufficient[context.ROLE.마스터].rows[0].slice(1, 3)], [10, 0]);

  const sufficient = reservationTables(8, 10, 'Y');
  installTables(sufficient);
  assert.equal(context.S3_1_주문확정(['O-1'], { manualRelease: true, silent: true }).준비, 1);
  assert.deepEqual([...sufficient[context.ROLE.마스터].rows[0].slice(1, 3)], [2, 8]);
});

test('reservation release retry reuses the existing picking instruction', () => {
  const orders = table(context.ROLE.주문,
    ['주문번호', '품목별 주문번호', '상품품목코드', '수량', '출고완료', '피킹지시번호', '주문상태'],
    [['O-1', 'I-1', 'SKU-1', 2, 0, 'PK-1', '예약']]);
  installTables({ [context.ROLE.주문]: orders });
  const result = context.S4_1_피킹지시생성(['O-1'], { silent: true });
  assert.equal(result.재사용, true);
  assert.equal(result.지시번호, 'PK-1');
});

test('reservation picking batches use the RES batch-number segment', () => {
  context.param_ = () => 'PK';
  context.Utilities = { formatDate: () => '20260819' };
  context.tz_ = () => 'Asia/Seoul';
  const header = { rows: [['PK-20260819-RES-001']], width: 1 };
  assert.equal(context.nextInstructionNo_(header, 0, 'RES'), 'PK-20260819-RES-002');
  assert.equal(context.nextInstructionNo_(header, 0, ''), 'PK-20260819-001');
});

function pickingTables() {
  return {
    [context.ROLE.주문]: table(context.ROLE.주문,
      ['주문번호', '품목별 주문번호', '상품품목코드', '수량', '주문상태', '출고완료', '피킹지시번호', '확정일시', '대기사유'],
      [['O-1', 'I-1', 'SKU-1', 4, '예약', 0, 'PK-1', new Date(), ''],
       ['O-1', 'I-2', 'SKU-1', 6, '예약', 0, 'PK-1', new Date(), '']]),
    [context.ROLE.라인]: table(context.ROLE.라인,
      ['주문번호', '품목별 주문번호', '상품코드', '필요수량', '확인', '실제수량', '예외사유', '담당자', '피킹지시번호', '라인상태', '처리일시'],
      [['O-1', 'I-1', 'SKU-1', 4, '', '', '', '', 'PK-1', '미처리', ''],
       ['O-1', 'I-2', 'SKU-1', 6, '', '', '', '', 'PK-1', '미처리', '']]),
    [context.ROLE.헤더]: table(context.ROLE.헤더,
      ['피킹지시번호', '주문번호', '피킹담당자', '상태', '출력일시'], [['PK-1', 'O-1', 'Kim', '대기', '']]),
    [context.ROLE.마스터]: table(context.ROLE.마스터,
      ['상품품목코드', '예약재고', '가용재고'], [['SKU-1', 10, 90]])
  };
}

test('successful output auto-completes without manual O and consumes only reserved inventory once', () => {
  const tables = pickingTables();
  const logs = installTables(tables);
  const result = context.finalizePickingAfterOutput_('PK-1');
  assert.deepEqual([...result.완료주문], ['O-1']);
  assert.equal(tables[context.ROLE.마스터].rows[0][1], 0);
  assert.equal(tables[context.ROLE.마스터].rows[0][2], 90, 'available stock must not be decremented twice');
  assert.equal(tables[context.ROLE.주문].rows[0][4], '출고완료');
  assert.deepEqual(tables[context.ROLE.라인].rows.map((row) => [row[4], row[5], row[9]]),
    [['O', 4, '완료'], ['O', 6, '완료']]);
  assert.equal(tables[context.ROLE.헤더].rows[0][3], '완료');
  assert.equal(logs.length, 2);

  const reprint = context.finalizePickingAfterOutput_('PK-1');
  assert.equal(reprint.이미완료, true);
  assert.deepEqual(tables[context.ROLE.마스터].rows[0].slice(1), [0, 90]);
  assert.equal(logs.length, 2, 'reprint must not consume inventory or duplicate shipment logs');
});

test('PDF failure preserves reservation and successful retry finalizes exactly once', () => {
  const tables = pickingTables();
  const logs = installTables(tables);
  const emptyIterator = () => ({ hasNext: () => false, next: () => undefined });
  const outputRoot = { getFilesByName: () => emptyIterator(), getFolders: () => emptyIterator() };
  context.getOrCreateSubFolder_ = () => ({ createFile: () => ({ getId: () => 'pdf-id' }) });
  context.buildPickingDocumentData_ = () => ({
    title: '피킹지시서', instructionNo: 'PK-1', createdAt: '2026-08-19 10:00',
    summary: { orderCount: 1, skuCount: 1, totalQuantity: 10 }, missingLocationCount: 0,
    pickSummary: [{ location: 'A-01', sku: 'SKU-1', productName: '상품', option: '', orderCount: 1, quantity: 10 }],
    orders: [{ orderNo: 'O-1', orderDateText: '', recipient: '', phoneLast4: '', postalCode: '', address: '', message: '',
      items: [{ productName: '상품', option: '', quantity: 10 }] }], hasAddress: false, hasMessage: false
  });
  context.Utilities = { formatDate: () => '2026-08-19' };
  context.tz_ = () => 'Asia/Seoul';
  context.MimeType = { PDF: 'application/pdf' };
  let fail = true;
  context.HtmlService = { createHtmlOutput: () => ({ getBlob: () => ({ getAs: () => {
    if (fail) throw new Error('render failed');
    return { setName() { return this; } };
  } }) }) };

  assert.throws(() => context.S9_피킹PDF생성('PK-1', outputRoot), /render failed/);
  assert.deepEqual(tables[context.ROLE.마스터].rows[0].slice(1), [10, 90]);
  assert.equal(tables[context.ROLE.주문].rows[0][4], '예약');

  fail = false;
  context.S9_피킹PDF생성('PK-1', outputRoot);
  assert.deepEqual(tables[context.ROLE.마스터].rows[0].slice(1), [0, 90]);
  assert.equal(tables[context.ROLE.주문].rows[0][4], '출고완료');
  assert.equal(logs.length, 2);
});

test('picking X delegates whole-order cancellation to the shared service', () => {
  const tables = pickingTables(); tables[context.ROLE.라인].rows[0][4] = 'X';
  tables[context.ROLE.라인].rows[0][6] = '재고없음'; installTables(tables);
  const calls = [];
  const original = context.cancelOrder_;
  context.cancelOrder_ = (orderNo, reason, source) => { calls.push({ orderNo, reason, source }); return { 취소: true }; };
  context.S5_1_결과반영();
  context.cancelOrder_ = original;
  assert.deepEqual(calls, [{ orderNo: 'O-1', reason: '재고없음', source: 'PICKING_X' }]);
});

test('reservation FIFO output finalizes selected batch orders and leaves unselected reservation untouched', () => {
  const tables = {
    [context.ROLE.주문]: table(context.ROLE.주문,
      ['주문번호', '품목별 주문번호', '상품품목코드', '수량', '주문상태', '출고완료', '피킹지시번호', '확정일시', '대기사유'],
      [['A001', 'I-1', 'SKU-1', 40, '예약', 0, 'PK-RES-1', new Date(), ''],
       ['A002', 'I-2', 'SKU-1', 50, '예약', 0, 'PK-RES-1', new Date(), ''],
       ['A003', 'I-3', 'SKU-1', 10, '예약', 0, '', '', '']]),
    [context.ROLE.라인]: table(context.ROLE.라인,
      ['주문번호', '품목별 주문번호', '상품코드', '필요수량', '확인', '실제수량', '예외사유', '담당자', '피킹지시번호', '라인상태', '처리일시'],
      [['A001', 'I-1', 'SKU-1', 40, '', '', '', '', 'PK-RES-1', '미처리', ''],
       ['A002', 'I-2', 'SKU-1', 50, '', '', '', '', 'PK-RES-1', '미처리', '']]),
    [context.ROLE.헤더]: table(context.ROLE.헤더,
      ['피킹지시번호', '주문번호', '피킹담당자', '상태', '출력일시'],
      [['PK-RES-1', 'A001', 'Kim', '대기', ''], ['PK-RES-1', 'A002', 'Kim', '대기', '']]),
    [context.ROLE.마스터]: table(context.ROLE.마스터,
      ['상품품목코드', '예약재고', '가용재고'], [['SKU-1', 90, 10]])
  };
  const logs = installTables(tables);
  context.finalizePickingAfterOutput_('PK-RES-1');
  assert.deepEqual(tables[context.ROLE.마스터].rows[0].slice(1), [0, 10]);
  assert.deepEqual(tables[context.ROLE.주문].rows.map(row => row[4]), ['출고완료', '출고완료', '예약']);
  context.finalizePickingAfterOutput_('PK-RES-1');
  assert.equal(logs.length, 2, 'reservation batch reprint must be inventory-idempotent');
});

function cancellationTables(state, available, reserved, confirmed = '') {
  return {
    [context.ROLE.주문]: table(context.ROLE.주문,
      ['주문번호', '품목별 주문번호', '상품품목코드', '수량', '주문상태', '출고완료', '취소사유', '취소일시', '취소경로', '확정일시', '대기사유', '피킹지시번호'],
      [['O-1', 'I-1', 'SKU-1', 2, state, state === '출고완료' ? 1 : 0, '', '', '', confirmed, '', 'PK-1']]),
    [context.ROLE.마스터]: table(context.ROLE.마스터,
      ['상품품목코드', '가용재고', '예약재고'], [['SKU-1', available, reserved]]),
    [context.ROLE.헤더]: table(context.ROLE.헤더,
      ['피킹지시번호', '주문번호', '상태'], [['PK-1', 'O-1', state === '출고완료' ? '완료' : '대기']]),
    [context.ROLE.라인]: table(context.ROLE.라인,
      ['주문번호', '품목별 주문번호', '피킹지시번호', '실제수량', '라인상태', '처리일시'],
      [['O-1', 'I-1', 'PK-1', state === '출고완료' ? 2 : '', state === '출고완료' ? '완료' : '미처리', '']])
  };
}

test('central cancellation restores according to state and is idempotent', () => {
  let tables = cancellationTables('예약', 10, 0); installTables(tables);
  context.cancelOrder_('O-1', '고객 요청', 'TEST', {});
  assert.deepEqual([...tables[context.ROLE.마스터].rows[0].slice(1)], [10, 0]);

  tables = cancellationTables('예약', 8, 2, new Date()); installTables(tables);
  context.cancelOrder_('O-1', '고객 요청', 'TEST', {});
  assert.deepEqual([...tables[context.ROLE.마스터].rows[0].slice(1)], [10, 0]);
  context.cancelOrder_('O-1', '고객 요청', 'TEST', {});
  assert.deepEqual([...tables[context.ROLE.마스터].rows[0].slice(1)], [10, 0]);

  tables = cancellationTables('출고완료', 8, 0); installTables(tables);
  const warning = context.cancelOrder_('O-1', '반품', 'TEST', {});
  assert.equal(warning.확인필요, true);
  assert.match(warning.메시지, /피킹지시서가 출력되어 출고 처리된 주문/);
  assert.equal(tables[context.ROLE.마스터].rows[0][1], 8);
  context.cancelOrder_('O-1', '반품', 'TEST', { confirmReturn: true });
  assert.equal(tables[context.ROLE.마스터].rows[0][1], 10);
  context.cancelOrder_('O-1', '반품', 'TEST', { confirmReturn: true });
  assert.equal(tables[context.ROLE.마스터].rows[0][1], 10);
});

test('cancellation failure reports order context through the centralized notifier', () => {
  const tables = cancellationTables('예약', 10, 0); installTables(tables);
  const notices = [];
  context.sendSystemNotification_ = (level, title, details) => { notices.push({ level, title, details }); return { sent: true }; };
  assert.throws(() => context.cancelOrder_('MISSING', 'test', 'TEST', {}), /주문번호를 찾을 수 없습니다/);
  assert.equal(notices.length, 1);
  assert.equal(notices[0].title, '주문 취소 실패');
  assert.equal(notices[0].details.주문번호, 'MISSING');
});
