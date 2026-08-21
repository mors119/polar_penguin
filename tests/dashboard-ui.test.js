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
    { 총가용: 100, 부족: [1], 품절: 2, 음수허용: 1, 위치미지정: 3, 예약부족SKU: [1] },
    { 전체라인: 8, 처리라인: 6, 미처리라인: 2, 진행률: 75, 작업대상주문: 1,
      오늘지시: 2, 오늘출고수량: 10, kpi: { 출력오류: 0 } },
    { latestTime: '2026-08-19 09:55', latestResult: 'PROCESSED · orders.csv', warning: '없음', recentErrors: [] });
  assert.deepEqual(requested, ['📊 대시보드']);
  assert.equal(cells.get('1:1'), '📊  Polar Penguin 통합 대시보드');
  assert.equal(cells.get('31:1'), '없음');
  assert.match(cells.get('35:1'), /셀은 버튼이 아닙니다/);
  assert.match(cells.get('28:1'), /위치 미지정 상품 3건/);
  assert.equal(cells.get('36:3'), 'Input 지금 처리');
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
    '위치_미지정상품관리', 'setupSystem', '백업_및_정리', '정리_로그']) {
    assert.ok(handlers.includes(expected), `${expected} is missing from the menu`);
  }
  for (const internal of ['S1_1_카페24재고동기화', 'S2_1_주문CSV취입', 'S3_1_주문확정', 'S4_1_피킹지시생성']) {
    assert.equal(handlers.includes(internal), false, `${internal} must stay out of the operator menu`);
  }
  for (const expected of ['Input 지금 처리', '피킹지시서 조회 / 재출력', '위치 미지정 상품',
    '선택 주문 취소', '예약상품 입고 관리', '대시보드 갱신', '시스템 상태 확인',
    '시스템 설치 / 복구', '백업 및 정리', '로그 정리']) {
    assert.ok(labels.includes(expected), `${expected} is missing from the operator menu`);
  }
  for (const internal of ['카페24 재고 동기화', '주문 CSV 취입', '주문 확정', '피킹지시 생성']) {
    assert.equal(labels.includes(internal), false, `${internal} must stay out of the operator menu`);
  }

  const allSource = fs.readdirSync(path.join(rootPath, 'src')).filter((name) => name.endsWith('.js'))
    .map((name) => fs.readFileSync(path.join(rootPath, 'src', name), 'utf8')).join('\n');
  assert.match(allSource, /⚠ 알림이메일 미설정/);
  for (const handler of handlers) {
    assert.match(allSource, new RegExp(`function\\s+${handler.replace(/[.*+?^${}()|[\\]\\]/g, '\\$&')}\\s*\\(`));
  }
});

test('reservation inbound modal uses explicit stock, order-count, quantity, and FIFO wording', () => {
  const html = fs.readFileSync(path.join(rootPath, 'src', 'ReservationInbound.html'), 'utf8');
  for (const label of ['예약상품 입고 관리', '현재 재고', '대기 주문', '대기 수량', '이번 입고 수량',
    '입고 반영 후 재고', 'FIFO 출고 가능 주문', 'FIFO 출고 가능 수량', '미사용 입고 수량',
    '미리보기', '입고 반영 및 피킹']) {
    assert.match(html, new RegExp(label));
  }
  assert.doesNotMatch(html, /예약재고|예약상품 피킹 관리/);
});
