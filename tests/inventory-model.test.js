const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const context = vm.createContext({ console });
for (const file of ['S0 공통.js', 'S1 카페24동기화.js']) {
  vm.runInContext(fs.readFileSync(path.join(root, 'src', file), 'utf8'), context);
}

test('Cafe24 snapshots reconcile unmanaged net stock without additive repetition', () => {
  assert.equal(context.netStockFromSnapshot_(0, 50, 'F'), -50);
  assert.equal(context.netStockFromSnapshot_(40, 50, 'F'), -10);
  assert.equal(context.netStockFromSnapshot_(60, 50, 'F'), 10);
  assert.equal(context.netStockFromSnapshot_(60, 50, 'F'), 10, 'same snapshot must not add another 60');
});

test('managed snapshot keeps the prior zero floor while unmanaged stock remains negative', () => {
  assert.equal(context.netStockFromSnapshot_(5, 10, 'T'), 0);
  assert.equal(context.netStockFromSnapshot_(5, 10, 'F'), -5);
});

test('only confirmed, active order rows contribute to snapshot demand', () => {
  const headers = ['상품품목코드', '수량', '주문상태', '확정일시'];
  const table = {
    role: '주문(완료)', headers,
    headerIndex: Object.fromEntries(headers.map((header, index) => [context.normKey_(header), index])),
    rows: [
      ['SKU-A', 30, '예약', new Date()],
      ['SKU-A', 20, '예약', new Date()],
      ['SKU-A', 7, '예약', ''],
      ['SKU-A', 4, '출고완료', new Date()],
      ['SKU-A', 3, '취소', new Date()]
    ]
  };
  context.readTable_ = () => table;
  context.getConfig_ = () => ({ 별칭: {} });
  assert.deepEqual(JSON.parse(JSON.stringify(context.committedDemandBySku_())), { 'SKU-A': 50 });
});
