const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const rootPath = path.resolve(__dirname, '..');
const context = vm.createContext({ console });
for (const file of ['S0 공통.js', 'S6 통합입력.js']) {
  vm.runInContext(fs.readFileSync(path.join(rootPath, 'src', file), 'utf8'), context);
}
const originalMarkPickingOutputState = context.markPickingOutputState_;

function plain(value) { return JSON.parse(JSON.stringify(value)); }

const order = [
  ['주문번호', '품목별 주문번호', '상품품목코드', '수량'],
  ['O-1', 'I-1', 'SKU-1', '2']
];
const inventory = [
  ['품목코드', '상품명', '재고수량'],
  ['SKU-1', '상품 1', '10']
];

test('header detection classifies order and inventory without using filenames', () => {
  assert.equal(context.detectInputType_(order[0]), 'ORDER');
  assert.equal(context.detectInputType_(inventory[0]), 'INVENTORY');
  assert.equal(context.detectInputType_(['foo', 'bar']), 'UNKNOWN');
});

test('inventory receives processing priority when mixed with order input', () => {
  const original = context.readUnifiedInput_;
  context.readUnifiedInput_ = (file) => file.data;
  assert.equal(context.inputFilePriority_({ data: inventory }), 0);
  assert.equal(context.inputFilePriority_({ data: order }), 1);
  assert.equal(context.inputFilePriority_({ data: [['unknown']] }), 2);
  context.readUnifiedInput_ = original;
});

test('validation accepts normal order/inventory and rejects empty, bad header, and missing values', () => {
  assert.deepEqual(plain(context.validateInput_('ORDER', order)), { type: 'ORDER', rows: 1 });
  assert.deepEqual(plain(context.validateInput_('INVENTORY', inventory)), { type: 'INVENTORY', rows: 1 });
  assert.throws(() => context.validateInput_('ORDER', [order[0]]), (error) => error.inputCode === 'EMPTY_FILE');
  assert.throws(() => context.validateInput_('UNKNOWN', [['foo'], ['x']]), (error) => error.inputCode === 'UNKNOWN_TYPE');
  assert.throws(
    () => context.validateInput_('ORDER', [order[0], ['O-1', '', 'SKU-1', 1]]),
    (error) => error.inputCode === 'MISSING_VALUE'
  );
});

test('successful checksum detection prevents the same file content from being processed twice', () => {
  const rows = [[new Date(), 'file-1', 'orders.csv', 'same-checksum', 'ORDER', 'PROCESSED', '', '']];
  context.ensureInputLogSheet_ = () => ({
    getLastRow: () => 2,
    getRange: () => ({ getValues: () => rows })
  });
  assert.equal(context.hasProcessedFingerprint_('same-checksum'), true);
  assert.equal(context.hasProcessedFingerprint_('different-checksum'), false);
});

test('order flow reserves and reaches fulfillment through successful PDF output', () => {
  const calls = [];
  context.S2_1_주문CSV취입 = () => { calls.push('S2'); return { 신규: 1, 주문번호: ['O-1'] }; };
  context.S1_1_카페24재고동기화 = () => { calls.push('S1'); };
  context.S3_1_주문확정 = () => { calls.push('S3'); return { 준비: 1, 예약: 0, 준비주문: ['O-1'] }; };
  context.S4_1_피킹지시생성 = () => { calls.push('S4'); return { 생성: true, 지시번호: 'PK-1' }; };
  context.S9_피킹PDF생성 = () => { calls.push('S9_FINALIZE'); return { 생성: true, 출고확정: { 완료주문: ['O-1'] } }; };
  context.runInputBusiness_('ORDER', {}, {});
  assert.deepEqual(calls, ['S2', 'S3', 'S4', 'S9_FINALIZE']);
  context.markPickingOutputState_ = originalMarkPickingOutputState;
});

test('inventory flow synchronizes stock but never releases reservation orders', () => {
  const calls = [];
  context.S1_1_카페24재고동기화 = () => { calls.push('S1'); return { 신규: 1 }; };
  context.collectPreorderData_ = () => { calls.push('AVAILABILITY'); return { 전체: 1 }; };
  context.runInputBusiness_('INVENTORY', {}, {});
  assert.deepEqual(calls, ['S1', 'AVAILABILITY']);
});

test('reservation is a successful order result without picking or PDF', () => {
  const calls = [];
  context.S2_1_주문CSV취입 = () => ({ 신규: 1, 주문번호: ['O-1'] });
  context.S3_1_주문확정 = () => ({ 준비: 0, 예약: 1, 준비주문: [] });
  context.S4_1_피킹지시생성 = () => { calls.push('S4'); };
  context.S9_피킹PDF생성 = () => { calls.push('S9'); };
  const result = context.runInputBusiness_('ORDER', {}, {});
  assert.equal(result.confirm.예약, 1);
  assert.deepEqual(calls, []);
});

test('PDF failure records output error and never finalizes shipment', () => {
  const calls = [];
  context.sendSystemNotification_ = (level, title) => { calls.push(level + ':' + title); return { sent: true }; };
  context.S2_1_주문CSV취입 = () => ({ 신규: 1, 주문번호: ['O-1'] });
  context.S3_1_주문확정 = () => ({ 준비: 1, 예약: 0, 준비주문: ['O-1'] });
  context.S4_1_피킹지시생성 = () => ({ 생성: true, 지시번호: 'PK-1' });
  context.S9_피킹PDF생성 = () => { throw new Error('render failed'); };
  context.markPickingOutputState_ = (no, state) => calls.push(state);
  context.writeOpLog_ = () => {};
  assert.throws(() => context.runInputBusiness_('ORDER', {}, {}), /render failed/);
  assert.deepEqual(calls, ['출력오류', 'ERROR:피킹 PDF 생성 실패']);
  context.markPickingOutputState_ = originalMarkPickingOutputState;
});

test('unsupported and corrupt files receive stable validation error codes', () => {
  context.MimeType = { GOOGLE_SHEETS: 'sheet' };
  assert.throws(
    () => context.readUnifiedInput_({ getName: () => 'photo.png', getMimeType: () => 'image/png' }),
    (error) => error.inputCode === 'UNSUPPORTED_FORMAT'
  );
  context.Utilities = { parseCsv: () => { throw new Error('broken'); } };
  context.readCsvText_ = () => 'broken';
  assert.throws(
    () => context.readUnifiedInput_({ getName: () => 'broken.csv', getMimeType: () => 'text/csv' }),
    (error) => error.inputCode === 'CORRUPT_FILE'
  );
});

test('input failure delegates to the centralized system notifier', () => {
  const calls = [];
  context.sendSystemNotification_ = (level, title, details) => { calls.push({ level, title, details }); return { sent: true }; };
  assert.equal(context.notifyInputFailure_('orders.csv', 'BAD_ROWS', 'invalid quantity'), true);
  assert.equal(calls[0].level, 'ERROR');
  assert.equal(calls[0].title, 'Input 처리 실패');
  assert.equal(calls[0].details.파일명, 'orders.csv');
  assert.equal(calls[0].details.오류코드, 'BAD_ROWS');
});
