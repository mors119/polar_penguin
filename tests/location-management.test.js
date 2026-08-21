const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');

function createContext(rows) {
  const headers = ['상품품목코드', '상품명', '옵션명', '가용재고', '기본보관위치'];
  const master = { headers, rows, sheet: {} };
  const context = vm.createContext({ console });
  vm.runInContext(fs.readFileSync(path.join(root, 'src', 'S10 위치관리.js'), 'utf8'), context);
  context.ROLE = { 마스터: 'master' };
  context.COL = { 상품품목코드: '상품품목코드', 상품명: '상품명', 옵션명: '옵션명',
    가용재고: '가용재고', 기본보관위치: '기본보관위치' };
  context.readTable_ = () => master;
  context.col_ = (table, name, required) => {
    const index = table.headers.indexOf(name);
    if (index < 0 && required) throw new Error(`missing ${name}`);
    return index;
  };
  context.toStr_ = value => String(value == null ? '' : value).trim();
  context.toNum_ = value => Number(value) || 0;
  context.withLock_ = fn => fn();
  context.writeColumn_ = () => {};
  context.writeOpLog_ = () => {};
  context.D0_대시보드전체갱신 = () => {};
  return { context, master };
}

test('location-management default query returns only products without a location', () => {
  const { context } = createContext([
    ['SKU1', '상품1', '', 10, 'A-01'],
    ['SKU2', '상품2', '빨강', 20, ''],
    ['SKU3', '상품3', '', 30, '   ']
  ]);
  const result = context.getProductsForLocationManagement_('', false);
  assert.deepEqual(JSON.parse(JSON.stringify(result.map(item => item.sku))), ['SKU2', 'SKU3']);
  assert.deepEqual(JSON.parse(JSON.stringify(context.getProductsForLocationManagement_('a-01', true).map(item => item.sku))), ['SKU1']);
});

test('location save trims input, updates only the location column, and reports invalid SKU', () => {
  const initial = [
    ['SKU1', '상품1', '', 10, 'A-01'],
    ['SKU2', '상품2', '빨강', 20, '']
  ];
  const { context, master } = createContext(initial.map(row => row.slice()));
  const beforeOtherColumns = master.rows.map(row => row.slice(0, 4));
  const result = context.saveProductLocations_([
    { sku: 'SKU2', location: '  B-02-03  ' },
    { sku: 'UNKNOWN', location: 'C-01' }
  ]);
  assert.equal(result.updated, 1);
  assert.equal(result.failed.length, 1);
  assert.equal(master.rows[1][4], 'B-02-03');
  assert.deepEqual(master.rows.map(row => row.slice(0, 4)), beforeOtherColumns);
});

test('inventory synchronization documents warehouse ownership and never writes location', () => {
  const source = fs.readFileSync(path.join(root, 'src', 'S1 카페24동기화.js'), 'utf8');
  assert.match(source, /기본보관위치는 절대 덮어쓰지 않음/);
  const writeList = source.match(/\[M\.상품명,[\s\S]*?\.forEach\(function \(idx\) \{[\s\S]*?writeColumn_\(마스터\.sheet, idx, 마스터\.rows\);[\s\S]*?\}\);/);
  assert.ok(writeList, 'inventory sync write list should be discoverable');
  assert.doesNotMatch(writeList[0], /M\.위치/);
});
