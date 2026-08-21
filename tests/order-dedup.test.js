const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const rootPath = path.resolve(__dirname, '..');
const context = vm.createContext({ console });
for (const file of ['S0 공통.js', 'S0 설치.js', 'S2 주문취입.js']) {
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

test('archived item-order keys continue to block reimport after active rows are cleaned', () => {
  context.getArchivedItemOrderKeys_ = () => ({ 'ITEM-001': true });
  const keys = context.buildExistingItemOrderKeys_({ rows: [['ITEM-ACTIVE']] }, 0);
  assert.equal(keys['ITEM-ACTIVE'], true);
  assert.equal(keys['ITEM-001'], true);
});

test('canonical order schema is exactly 25 fixed columns in the required order', () => {
  assert.deepEqual(Array.from(context.ORDER_FIXED_HEADERS), [
    '쇼핑몰', '쇼핑몰번호', '주문번호', '품목별 주문번호', '배송메시지',
    '총 주문금액(KRW)', '총 결제금액(KRW)', '상품품목코드', '주문상품명(기본)',
    '상품옵션(기본)', '수량', '판매가', '수령인', '수령인 휴대전화',
    '수령인 우편번호', '수령인 주소(전체)', '출고완료', '피킹지시번호',
    '주문상태', '취소사유', '취소일시', '취소경로', '확정일시', '대기사유', '운영메모'
  ]);
});

test('dynamic CSV columns append only when observed and never duplicate fixed semantics', () => {
  const fixed = Array.from(context.ORDER_FIXED_HEADERS);
  assert.deepEqual(Array.from(context.normalizeOrderImportHeaders_([
    '주문번호', '주문상품명(기본)', '상품명', '옵션명', '수령인 주소',
    '주문일시', '주문경로', '결제수단', '배송업체'
  ], fixed)), ['주문일시', '주문경로', '결제수단', '배송업체']);
  assert.deepEqual(Array.from(context.normalizeOrderImportHeaders_([
    '주문번호', '송장번호'
  ], fixed.concat(['주문일시']))), ['송장번호']);
  assert.deepEqual(Array.from(context.normalizeOrderImportHeaders_(fixed, fixed)), []);
});

test('legacy duplicate fields migrate without overwriting canonical values and meaningful dynamic data remains', () => {
  const headers = ['주문번호', '상품명', '주문상품명(기본)', '옵션명', '상품옵션(기본)',
    '수령인 주소', '수령인 주소(전체)', '주문일시', '배송업체', '사용자필드'];
  const rows = [
    ['O-1', '레거시 상품', '', '빨강', '', '서울', '', '2026-08-01', '', 'KEEP-1'],
    ['O-2', '덮으면 안 됨', '정식 상품', '구 옵션', '정식 옵션', '구 주소', '정식 주소', '', '', 'KEEP-2']
  ];
  const migrated = context.buildOrderSchemaMigration_(headers, rows);
  assert.deepEqual(Array.from(migrated.headers.slice(0, 25)), Array.from(context.ORDER_FIXED_HEADERS));
  assert.deepEqual(Array.from(migrated.headers.slice(25)), ['주문일시', '사용자필드']);
  assert.equal(migrated.rows[0][migrated.headers.indexOf('주문상품명(기본)')], '레거시 상품');
  assert.equal(migrated.rows[0][migrated.headers.indexOf('상품옵션(기본)')], '빨강');
  assert.equal(migrated.rows[0][migrated.headers.indexOf('수령인 주소(전체)')], '서울');
  assert.equal(migrated.rows[1][migrated.headers.indexOf('주문상품명(기본)')], '정식 상품');
  assert.equal(migrated.rows[1][migrated.headers.indexOf('상품옵션(기본)')], '정식 옵션');
  assert.equal(migrated.rows[1][migrated.headers.indexOf('수령인 주소(전체)')], '정식 주소');
  assert.equal(migrated.headers.includes('배송업체'), false, 'empty former fixed columns are removed');

  const repeated = context.buildOrderSchemaMigration_(migrated.headers, migrated.rows);
  assert.equal(repeated.changed, false);
  assert.deepEqual(Array.from(repeated.headers), Array.from(migrated.headers));
});

test('CSV legacy name option and address headers write into canonical fixed columns', () => {
  const headers = Array.from(context.ORDER_FIXED_HEADERS).concat(['주문일시']);
  const table = {
    width: headers.length,
    role: '주문(완료)',
    headers,
    headerIndex: Object.fromEntries(headers.map((header, index) => [context.normKey_(header), index]))
  };
  const csvHeaders = ['주문번호', '품목별 주문번호', '상품코드', '상품명', '옵션명', '수량', '수령인 주소', '주문일시'];
  context.MimeType = { GOOGLE_SHEETS: 'google-sheet' };
  context.Utilities = { parseCsv: () => [csvHeaders,
    ['O-2', 'I-2', 'SKU-2', '상품 B', '파랑', 3, '부산시', '2026-08-02']] };
  context.readCsvText_ = () => 'ignored';
  context.getConfig_ = () => ({ 별칭: {} });
  const result = context.parseCsvFile_(
    { getMimeType: () => 'text/csv', getName: () => 'orders.csv' }, table,
    headers.indexOf('품목별 주문번호'), {}
  );
  const row = result.rows[0];
  assert.equal(row[headers.indexOf('상품품목코드')], 'SKU-2');
  assert.equal(row[headers.indexOf('주문상품명(기본)')], '상품 B');
  assert.equal(row[headers.indexOf('상품옵션(기본)')], '파랑');
  assert.equal(row[headers.indexOf('수령인 주소(전체)')], '부산시');
  assert.equal(row[headers.indexOf('주문일시')], '2026-08-02');
});

test('CSV canonical values win when the same file also contains legacy duplicate headers', () => {
  const headers = Array.from(context.ORDER_FIXED_HEADERS);
  const table = { width: headers.length, role: '주문(완료)', headers,
    headerIndex: Object.fromEntries(headers.map((header, index) => [context.normKey_(header), index])) };
  const csvHeaders = ['주문번호', '품목별 주문번호', '상품품목코드', '주문상품명(기본)', '상품명',
    '상품옵션(기본)', '옵션명', '수량'];
  context.MimeType = { GOOGLE_SHEETS: 'google-sheet' };
  context.Utilities = { parseCsv: () => [csvHeaders,
    ['O-3', 'I-3', 'SKU-3', '정식 상품', '레거시 상품', '정식 옵션', '레거시 옵션', 1]] };
  context.readCsvText_ = () => 'ignored';
  const result = context.parseCsvFile_(
    { getMimeType: () => 'text/csv', getName: () => 'orders.csv' }, table,
    headers.indexOf('품목별 주문번호'), {}
  );
  assert.equal(result.rows[0][headers.indexOf('주문상품명(기본)')], '정식 상품');
  assert.equal(result.rows[0][headers.indexOf('상품옵션(기본)')], '정식 옵션');
});
