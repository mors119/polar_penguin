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
    { 예약: 3, 취소: 4, 출고완료: 5, 전체: 12,
      예약전체: 3, 예약수량: 8, 예약출고가능: 1, 예약출고가능SKU: 1, 예약부족: 2,
      권고: { 제목: '주문 확정', 내용: ['확인'], 색: 'FFFFFF' } },
    { 총가용: 100, 총예약: 10, 부족: [1], 품절: 2, 예약부족SKU: [1] },
    { 전체라인: 8, 처리라인: 6, 미처리라인: 2, 진행률: 75, 작업대상주문: 1,
      오늘지시: 2, 오늘출고수량: 10, kpi: { 출력오류: 0 } },
    { latestTime: '2026-08-19 09:55', latestResult: 'PROCESSED · orders.csv', warning: '없음', recentErrors: [] });
  assert.deepEqual(requested, ['📊 대시보드']);
  assert.equal(cells.get('1:1'), '📊  Polar Penguin 통합 대시보드');
  assert.equal(cells.get('31:1'), '없음');
  assert.match(cells.get('35:1'), /셀은 버튼이 아닙니다/);
  assert.equal(cells.get('36:3'), 'Input 지금 처리');
});

test('onOpen registers the current menu and every handler exists in source', () => {
  const menuNames = [];
  const handlers = [];
  const labels = [];
  function menu(name) {
    menuNames.push(name);
    return {
      addItem(label, handler) { labels.push(label); handlers.push(handler); return this; },
      addSubMenu() { return this; }, addSeparator() { return this; }, addToUi() { return this; }
    };
  }
  context.SpreadsheetApp = { getUi: () => ({ createMenu: menu }) };
  context.Logger = { log() {} };
  context.onOpen();
  assert.ok(menuNames.includes('📦 Polar Penguin'));
  for (const expected of ['processInput', '선택_주문취소', '예약상품_피킹관리',
    'S9_1_작업지시서출력', 'D0_대시보드전체갱신', '진단_시트구조',
    'setupSystem', '정리_로그', '설정_보기']) {
    assert.ok(handlers.includes(expected), `${expected} is missing from the menu`);
  }
  for (const internal of ['S1_1_카페24재고동기화', 'S2_1_주문CSV취입', 'S3_1_주문확정', 'S4_1_피킹지시생성']) {
    assert.equal(handlers.includes(internal), false, `${internal} must stay out of the operator menu`);
  }
  for (const expected of ['Input 지금 처리', '작업지시서 조회 / 재출력',
    '선택 주문 취소', '예약상품 피킹 관리', '대시보드 갱신', '시스템 상태 확인',
    '시스템 설치 / 복구', '설정 보기', '로그 정리']) {
    assert.ok(labels.includes(expected), `${expected} is missing from the operator menu`);
  }
  for (const internal of ['카페24 재고 동기화', '주문 CSV 취입', '주문 확정', '피킹지시 생성']) {
    assert.equal(labels.includes(internal), false, `${internal} must stay out of the operator menu`);
  }

  const allSource = fs.readdirSync(path.join(rootPath, 'src')).filter((name) => name.endsWith('.js'))
    .map((name) => fs.readFileSync(path.join(rootPath, 'src', name), 'utf8')).join('\n');
  for (const handler of handlers) {
    assert.match(allSource, new RegExp(`function\\s+${handler.replace(/[.*+?^${}()|[\\]\\]/g, '\\$&')}\\s*\\(`));
  }
});
