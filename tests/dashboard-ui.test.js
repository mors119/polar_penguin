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
    clear() {}, clearFormats() {}, getMaxColumns: () => 10, insertColumnsAfter() {},
    getRange: (row, col) => new Range(row, col),
    setRowHeight() {}, setColumnWidth() {}, setFrozenRows() {}, setHiddenGridlines() {}
  };
  const requested = [];
  const ss = { getSheetByName(name) { requested.push(name); return name === '📊 대시보드' ? sheet : null; } };
  context.Utilities = { formatDate: () => '2026-08-19 10:00:00' };
  context.tz_ = () => 'Asia/Seoul';
  const order = { 예약: 3, 취소: 4, 출고완료: 5, 전체: 12,
      예약전체: 3, 예약수량: 8, 예약출고가능: 1, 예약출고가능SKU: 1, 예약부족: 2,
      권고: { 제목: '주문 확정', 내용: ['확인'], 색: 'FFFFFF' } };
  const stock = { 총가용: 100, 부족: [1], 품절: 2, 음수허용: 1, 위치미지정: 3, 예약부족SKU: [1] };
  const picking = { 전체라인: 8, 처리라인: 6, 미처리라인: 2, 진행률: 75, 작업대상주문: 1,
    오늘지시: 2, 오늘출고수량: 10, kpi: { 출력오류: 0 } };
  const operation = { latestTime: '2026-08-19 09:55', latestResult: 'PROCESSED · orders.csv', warning: '없음', recentErrors: [] };
  const before = JSON.stringify({ order, stock, picking, operation });
  context.renderIntegratedDashboard_(ss, order, stock, picking, operation);
  assert.deepEqual(requested, ['📊 대시보드']);
  assert.equal(cells.get('1:1'), '📊  Polar Penguin 통합 대시보드');
  assert.equal(cells.get('4:1'), '대기 주문');
  assert.equal(cells.get('5:1'), 3);
  assert.equal(cells.get('4:11'), '처리 오류');
  assert.equal(cells.get('5:11'), 0);
  assert.equal(cells.get('8:1'), '처리 필요');
  assert.equal(cells.get('8:7'), '재고 경고');
  assert.equal(cells.get('15:1'), '예약 대기');
  assert.equal(cells.get('15:7'), '최근 처리');
  assert.match(cells.get('27:1'), /셀은 버튼이 아닙니다/);
  assert.equal(cells.get('28:4'), 'Input 지금 처리');
  assert.equal(JSON.stringify({ order, stock, picking, operation }), before,
    'rendering must not mutate or recalculate dashboard metrics');
});

test('stock dashboard count reflects products whose warehouse location is blank', () => {
  const headers = ['상품품목코드', '상품명', '가용재고', '재고관리', '기본보관위치'];
  const headerIndex = Object.fromEntries(headers.map((header, index) => [context.normKey_(header), index]));
  context.readTable_ = () => ({ headers, headerIndex, role: '상품마스터', rows: [
    ['SKU1', '상품1', 10, 'T', 'A-01'],
    ['SKU2', '상품2', 5, 'T', ''],
    ['SKU3', '상품3', -2, 'F', '   ']
  ] });
  context.param_ = () => 3;
  context.collectPreorderData_ = () => ({ 예약부족SKU: [] });
  const stock = context.collectStockStatus_();
  assert.equal(stock.위치미지정, 2);
  assert.equal(stock.음수허용, 1);
  assert.equal(stock.부족.some(item => item.코드 === 'SKU3'), false);
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
  for (const expected of ['processInput', '선택_주문취소', '예약상품_입고관리',
    'S9_1_작업지시서출력', 'D0_대시보드전체갱신', '진단_시트구조',
    '위치_미지정상품관리', 'setupSystem', '설정_열기', '백업_및_정리', '정리_로그']) {
    assert.ok(handlers.includes(expected), `${expected} is missing from the menu`);
  }
  for (const internal of ['S1_1_카페24재고동기화', 'S2_1_주문CSV취입', 'S3_1_주문확정', 'S4_1_피킹지시생성']) {
    assert.equal(handlers.includes(internal), false, `${internal} must stay out of the operator menu`);
  }
  for (const expected of ['Input 지금 처리', '피킹지시서 조회 / 재출력', '위치 미지정 상품',
    '선택 주문 취소', '예약상품 입고 관리', '대시보드 갱신', '시스템 상태 확인',
    '시스템 설치 / 복구', '설정', '백업 및 정리', '로그 정리']) {
    assert.ok(labels.includes(expected), `${expected} is missing from the operator menu`);
  }
  assert.equal(labels.filter(label => label === '설정').length, 1);
  for (const internal of ['카페24 재고 동기화', '주문 CSV 취입', '주문 확정', '피킹지시 생성']) {
    assert.equal(labels.includes(internal), false, `${internal} must stay out of the operator menu`);
  }

  const allSource = fs.readdirSync(path.join(rootPath, 'src')).filter((name) => name.endsWith('.js'))
    .map((name) => fs.readFileSync(path.join(rootPath, 'src', name), 'utf8')).join('\n');
  assert.match(allSource, /⚠ 알림이메일 미설정/);
  assert.doesNotMatch(allSource, /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i,
    'runtime source must not contain a hard-coded notification recipient');
  for (const variable of ['CONSOLE_SS_ID', 'DEFAULT_FOLDER_ID']) {
    const assignment = allSource.match(new RegExp(`\\b${variable}\\s*=\\s*['\"]([^'\"]*)['\"]`));
    assert.ok(assignment && assignment[1] === '', `${variable} must not embed an environment-specific Drive ID`);
  }
  for (const handler of handlers) {
    assert.match(allSource, new RegExp(`function\\s+${handler.replace(/[.*+?^${}()|[\\]\\]/g, '\\$&')}\\s*\\(`));
  }
});

test('settings menu opens the HTML modal without exposing or changing the raw settings sheet', () => {
  const values = [
    ['구분', '키', '값', '비고'],
    ['파라미터', '알림이메일', 'operator@example.com', '알림 수신자'],
    ['파라미터', '정리보존일수', 45, '보존 기준']
  ];
  const settings = {
    hidden: true,
    getLastRow: () => values.length,
    getRange(row, col, rows = 1, cols = 1) {
      return {
        getValues: () => Array.from({ length: rows }, (_, r) =>
          Array.from({ length: cols }, (_, c) => values[row - 1 + r][col - 1 + c]))
      };
    }
  };
  const internal = { hidden: true };
  const ss = { getSheetByName: name => name === '설정' ? settings : internal };
  const before = JSON.stringify(values);
  let dialog = null;
  context.consoleSS_ = () => ss;
  context.HtmlService = { createHtmlOutputFromFile: name => ({
    name, setWidth(width) { this.width = width; return this; }, setHeight(height) { this.height = height; return this; }
  }) };
  context.SpreadsheetApp = { getUi: () => ({ showModalDialog: (html, title) => { dialog = { html, title }; } }) };
  const result = context.설정_열기();
  assert.equal(result, true);
  assert.equal(dialog.html.name, 'Settings');
  assert.equal(dialog.title, 'Polar Penguin 설정');
  assert.equal(dialog.html.width, 560);
  assert.equal(settings.hidden, true);
  assert.equal(internal.hidden, true);
  assert.equal(JSON.stringify(values), before);
});

test('settings menu reports the established repair action when the settings sheet is missing', () => {
  let message = '';
  context.consoleSS_ = () => ({ getSheetByName: () => null });
  context.alert_ = value => { message = value; };
  assert.equal(context.설정_열기(), null);
  assert.match(message, /⚙ 관리 → 시스템 설치 \/ 복구/);
});
