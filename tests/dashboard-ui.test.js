const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const rootPath = path.resolve(__dirname, '..');
const context = vm.createContext({ console });
for (const file of ['S0 공통.js', 'D 대시보드.js', 'S0 설치.js']) {
  vm.runInContext(fs.readFileSync(path.join(rootPath, 'src', file), 'utf8'), context);
}

test('integrated dashboard renderer writes the real dashboard tab and visible action guidance', () => {
  const cells = new Map();
  class Range {
    constructor(row, col) { this.row = row; this.col = col; }
    setValue(value) { cells.set(`${this.row}:${this.col}`, value); return this; }
    setValues(values) {
      values.forEach((line, r) => line.forEach((value, c) => cells.set(`${this.row + r}:${this.col + c}`, value)));
      return this;
    }
  }
  for (const method of ['merge', 'setFontSize', 'setFontWeight', 'setFontColor', 'setBackground',
    'setVerticalAlignment', 'setHorizontalAlignment', 'setFontFamily', 'setWrap']) {
    Range.prototype[method] = function noop() { return this; };
  }
  const sheet = {
    clear() {}, clearFormats() {}, getMaxColumns: () => 10,
    getRange: (row, col) => new Range(row, col),
    setRowHeight() {}, setColumnWidth() {}, setFrozenRows() {}, setHiddenGridlines() {}
  };
  const requested = [];
  const ss = { getSheetByName(name) { requested.push(name); return name === '📊 대시보드' ? sheet : null; } };
  context.Utilities = { formatDate: () => '2026-08-19 10:00:00' };
  context.tz_ = () => 'Asia/Seoul';
  context.renderIntegratedDashboard_(ss,
    { 접수: 1, 확정: 2, 예약대기: 3, 취소: 4, 출고완료: 5, 전체: 15,
      피킹: { 속도: 20, 예상분: 30 }, 예약목록: [{ 수량: 2, 부족: 0 }],
      권고: { 제목: '주문 확정', 내용: ['확인'], 색: 'FFFFFF' } },
    { 총가용: 100, 총예약: 10, 부족: [1], 품절: 2 },
    { 전체라인: 8, 처리라인: 6, 진행률: 75, kpi: { 대기: 1, 진행: 2 } },
    { input: 'PROCESSED', failure: '없음' });
  assert.deepEqual(requested, ['📊 대시보드']);
  assert.equal(cells.get('1:1'), '📊  Polar Penguin 통합 대시보드');
  assert.match(cells.get('35:1'), /셀은 버튼이 아닙니다/);
  assert.equal(cells.get('36:1'), '📥 Input 지금 처리');
});

test('onOpen registers the current menu and every handler exists in source', () => {
  const menuNames = [];
  const handlers = [];
  function menu(name) {
    menuNames.push(name);
    return {
      addItem(label, handler) { handlers.push(handler); return this; },
      addSubMenu() { return this; }, addSeparator() { return this; }, addToUi() { return this; }
    };
  }
  context.SpreadsheetApp = { getUi: () => ({ createMenu: menu }) };
  context.Logger = { log() {} };
  context.onOpen();
  assert.ok(menuNames.includes('📦 Polar Penguin'));
  for (const expected of ['processInput', 'S3_1_주문확정', 'S4_1_피킹지시생성', 'S5_2_수동반영',
    'S9_1_작업지시서출력', 'D0_대시보드전체갱신', '운영_예약대기조회', '진단_시트구조',
    'setupSystem', '정리_로그', '설정_캐시초기화']) {
    assert.ok(handlers.includes(expected), `${expected} is missing from the menu`);
  }

  const allSource = fs.readdirSync(path.join(rootPath, 'src')).filter((name) => name.endsWith('.js'))
    .map((name) => fs.readFileSync(path.join(rootPath, 'src', name), 'utf8')).join('\n');
  for (const handler of handlers) {
    assert.match(allSource, new RegExp(`function\\s+${handler.replace(/[.*+?^${}()|[\\]\\]/g, '\\$&')}\\s*\\(`));
  }
});
