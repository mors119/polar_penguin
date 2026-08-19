const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const rootPath = path.resolve(__dirname, '..');
const context = vm.createContext({ console });
for (const file of ['S0 공통.js', 'S2 주문취입.js']) {
  vm.runInContext(fs.readFileSync(path.join(rootPath, 'src', file), 'utf8'), context);
}

test('order ingest keeps item-order-number deduplication for reuploaded orders', () => {
  const headers = ['주문번호', '품목별 주문번호', '상품품목코드', '수량'];
  const table = {
    width: headers.length,
    role: '주문완료',
    headerIndex: Object.fromEntries(headers.map((header, index) => [context.normKey_(header), index]))
  };
  context.MimeType = { GOOGLE_SHEETS: 'google-sheet' };
  context.Utilities = { parseCsv: () => [headers, ['O-1', 'I-1', 'SKU-1', 1], ['O-1', 'I-2', 'SKU-2', 2]] };
  context.readCsvText_ = () => 'ignored';
  const result = context.parseCsvFile_(
    { getMimeType: () => 'text/csv', getName: () => 'orders.csv' },
    table,
    1,
    { 'I-1': true }
  );
  assert.equal(result.rows.length, 1);
  assert.equal(result.rows[0][1], 'I-2');
  assert.equal(result.중복, 1);
});
