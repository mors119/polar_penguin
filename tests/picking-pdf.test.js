const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync(path.resolve(__dirname, '..', 'src', 'S9 작업지시서.js'), 'utf8');
const context = vm.createContext({ console });
vm.runInContext(source, context);
const originalInstructionBuilder = context.S9_지시서생성;

test('picking PDF is created once and reused on rerun', () => {
  const files = new Map();
  const dateFolder = {
    getFilesByName: (name) => ({ hasNext: () => files.has(name) }),
    createFile: (blob) => {
      const file = { getId: () => 'pdf-id' };
      files.set(blob.name, file);
      return file;
    }
  };
  const outputRoot = {
    getFilesByName: () => ({ hasNext: () => false }),
    getFolders: () => {
      let used = false;
      return { hasNext: () => !used, next: () => { used = true; return dateFolder; } };
    }
  };
  context.Utilities = { formatDate: () => '2026-08-19' };
  context.MimeType = { PDF: 'application/pdf' };
  context.tz_ = () => 'Asia/Seoul';
  context.getOrCreateSubFolder_ = () => dateFolder;
  context.S9_지시서생성 = () => ({ 출력시각: '2026-08-19 10:00', 슬롯: [], 총수량: 0 });
  context.HtmlService = { createHtmlOutput: () => ({
    getBlob: () => ({ getAs: () => ({ setName(name) { this.name = name; return this; } }) })
  }) };
  context.writeOpLog_ = () => {};
  let finalized = false;
  let shipmentMutations = 0;
  context.finalizePickingAfterOutput_ = () => {
    if (!finalized) { finalized = true; shipmentMutations++; return { 이미완료: false }; }
    return { 이미완료: true };
  };

  const first = context.S9_피킹PDF생성('PK-20260819-001', outputRoot);
  const second = context.S9_피킹PDF생성('PK-20260819-001', outputRoot);
  assert.equal(first.생성, true);
  assert.equal(second.생성, false);
  assert.equal(second.재사용, true);
  assert.equal(files.size, 1);
  assert.equal(shipmentMutations, 1, 'reprint must not create a second shipment mutation');
  context.S9_지시서생성 = originalInstructionBuilder;
});

test('first manual printable output delegates selected orders to the shared finalizer', () => {
  context.ROLE = { 헤더: 'header', 주문: 'orders', 라인: 'lines' };
  context.COL = {
    피킹지시번호: '피킹지시번호', 주문번호: '주문번호', 카트슬롯: '카트 슬롯', 품목수: '품목수',
    총수량: '총수량', 피킹담당자: '피킹담당자', 상태: '상태', 출력일시: '출력일시',
    품목별주문번호: '품목별 주문번호', 순번: '순번', 보관위치: '보관위치', 상품코드: '상품코드',
    상품명: '상품명', 옵션: '옵션', 필요수량: '필요수량', 라인상태: '라인상태'
  };
  context.ENUM = { 헤더상태: { 완료: '완료', 취소: '취소' }, 라인상태: { 취소: '취소' } };
  const tables = {
    header: { headers: ['피킹지시번호', '주문번호', '카트 슬롯', '품목수', '총수량', '피킹담당자', '상태', '출력일시'],
      rows: [['PK-1', 'O-1', 1, 1, 2, '', '대기', '']], sheet: {} },
    orders: { headers: ['주문번호', '품목별 주문번호'], rows: [['O-1', 'I-1']], sheet: {} },
    lines: { headers: ['순번', '보관위치', '상품코드', '상품명', '옵션', '필요수량', '품목별 주문번호', '라인상태'],
      rows: [[1, 'A-1', 'SKU-1', 'Book', '', 2, 'I-1', '미처리']], sheet: {} }
  };
  context.readTable_ = role => tables[role];
  context.col_ = (table, name) => table.headers.indexOf(name);
  context.toStr_ = value => String(value == null ? '' : value).trim();
  context.toNum_ = value => Number(value) || 0;
  context.isBlank_ = value => value === '' || value == null;
  context.withLock_ = fn => fn();
  context.writeColumn_ = () => {};
  context.writeOpLog_ = () => {};
  context.D0_대시보드전체갱신 = () => {};
  context.Utilities = { formatDate: () => '2026-08-19 10:00' };
  context.tz_ = () => 'Asia/Seoul';
  const calls = [];
  context.finalizePickingAfterOutput_ = (instructionNo, options) => {
    calls.push({ instructionNo, orderNos: options.orderNos }); return { 완료주문: options.orderNos };
  };

  const result = context.S9_지시서생성('Kim', 1);
  assert.equal(result.슬롯.length, 1);
  assert.deepEqual(JSON.parse(JSON.stringify(calls)), [{ instructionNo: 'PK-1', orderNos: ['O-1'] }]);
});
