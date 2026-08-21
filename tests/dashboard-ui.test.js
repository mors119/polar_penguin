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

function dashboardData() {
  return {
    order: { 예약: 3, 취소: 4, 출고완료: 5, 전체: 12,
      예약전체: 3, 예약수량: 8, 예약출고가능: 1, 예약출고가능SKU: 1, 예약부족: 2,
      권고: { 제목: '주문 확정', 내용: ['확인'], 색: 'FFFFFF' } },
    stock: { 총가용: 100, 부족: [1], 품절: 2, 음수허용: 1, 위치미지정: 3, 예약부족SKU: [1] },
    picking: { 전체라인: 8, 처리라인: 6, 미처리라인: 2, 진행률: 75, 작업대상주문: 1,
      오늘지시: 2, 오늘출고수량: 10, kpi: { 출력오류: 0 } },
    operation: { latestTime: '2026-08-19 09:55', latestResult: 'PROCESSED · orders.csv',
      warning: '없음', recentErrors: [], lastBackup: '백업 없음' }
  };
}

function createDashboardSheet({ maxRows = 40, maxColumns = 15, oldMerges = [] } = {}) {
  const cells = new Map();
  const merges = oldMerges.map(range => ({ ...range }));
  const events = [];
  const overlaps = (a, b) => a.row <= b.row + b.rows - 1 && b.row <= a.row + a.rows - 1 &&
    a.col <= b.col + b.cols - 1 && b.col <= a.col + a.cols - 1;
  const contains = (outer, inner) => outer.row <= inner.row && outer.col <= inner.col &&
    outer.row + outer.rows >= inner.row + inner.rows && outer.col + outer.cols >= inner.col + inner.cols;
  const key = range => `${range.row}:${range.col}:${range.rows}:${range.cols}`;

  class Range {
    constructor(row, col, rows = 1, cols = 1) { Object.assign(this, { row, col, rows, cols }); }
    setValue(value) { cells.set(`${this.row}:${this.col}`, value); events.push(`value:${key(this)}`); return this; }
    setValues(values) {
      values.forEach((line, r) => line.forEach((value, c) => cells.set(`${this.row + r}:${this.col + c}`, value)));
      events.push(`values:${key(this)}`);
      return this;
    }
    merge() {
      if (merges.some(existing => overlaps(this, existing))) throw new Error(`overlapping merge ${key(this)}`);
      merges.push({ row: this.row, col: this.col, rows: this.rows, cols: this.cols });
      events.push(`merge:${key(this)}`);
      return this;
    }
    breakApart() {
      const intersecting = merges.filter(existing => overlaps(this, existing));
      if (intersecting.some(existing => !contains(this, existing))) {
        throw new Error('셀을 병합하거나 병합을 취소하려면 해당 범위의 모든 셀을 선택해야 합니다.');
      }
      for (const existing of intersecting) merges.splice(merges.indexOf(existing), 1);
      events.push(`breakApart:${key(this)}`);
      return this;
    }
  }
  for (const method of ['setFontSize', 'setFontWeight', 'setFontColor', 'setBackground',
    'setVerticalAlignment', 'setHorizontalAlignment', 'setFontFamily', 'setWrap']) {
    Range.prototype[method] = function noop() { return this; };
  }
  const sheet = {
    cells, merges, events,
    getMaxRows: () => maxRows,
    getMaxColumns: () => maxColumns,
    insertRowsAfter(after, count) { maxRows = Math.max(maxRows, after + count); events.push(`insertRows:${count}`); },
    insertColumnsAfter(after, count) { maxColumns = Math.max(maxColumns, after + count); events.push(`insertColumns:${count}`); },
    getRange: (row, col, rows = 1, cols = 1) => new Range(row, col, rows, cols),
    getDataRange: () => new Range(1, 1, 1, 1),
    clearContents() { cells.clear(); events.push('clearContents'); },
    clearFormats() { events.push('clearFormats'); },
    setRowHeight() {}, setColumnWidth() {}, setFrozenRows() {}, setHiddenGridlines() {}
  };
  return sheet;
}

function dashboardSpreadsheet(sheet) {
  const requested = [];
  const ss = { getSheetByName(name) { requested.push(name); return name === '📊 대시보드' ? sheet : null; } };
  return { ss, requested };
}

context.Utilities = { formatDate: () => '2026-08-19 10:00:00' };
context.tz_ = () => 'Asia/Seoul';

test('integrated dashboard renderer preserves metrics and expected labels', () => {
  const sheet = createDashboardSheet({ maxRows: 20, maxColumns: 10 });
  const { ss, requested } = dashboardSpreadsheet(sheet);
  const { order, stock, picking, operation } = dashboardData();
  context.Utilities = { formatDate: () => '2026-08-19 10:00:00' };
  const before = JSON.stringify({ order, stock, picking, operation });
  context.renderIntegratedDashboard_(ss, order, stock, picking, operation);
  assert.deepEqual(requested, ['📊 대시보드']);
  assert.equal(sheet.cells.get('1:1'), '📊  Polar Penguin 통합 대시보드');
  assert.equal(sheet.cells.get('4:1'), '대기 주문');
  assert.equal(sheet.cells.get('5:1'), 3);
  assert.equal(sheet.cells.get('4:11'), '처리 오류');
  assert.equal(sheet.cells.get('5:11'), 0);
  assert.equal(sheet.cells.get('8:1'), '처리 필요');
  assert.equal(sheet.cells.get('8:7'), '재고 경고');
  assert.equal(sheet.cells.get('15:1'), '예약 대기');
  assert.equal(sheet.cells.get('15:7'), '최근 처리');
  assert.equal(sheet.cells.get('22:1'), '운영 알림');
  assert.match(sheet.cells.get('25:1'), /^권고:/);
  assert.match(sheet.cells.get('27:1'), /셀은 버튼이 아닙니다/);
  assert.equal(sheet.cells.get('28:4'), 'Input 지금 처리');
  assert.ok(sheet.events.includes('insertRows:10'));
  assert.ok(sheet.events.includes('insertColumns:2'));
  assert.equal(JSON.stringify({ order, stock, picking, operation }), before,
    'rendering must not mutate or recalculate dashboard metrics');
});

test('renderer clears old full-sheet merges and succeeds ten times without overlaps or stale content', () => {
  const sheet = createDashboardSheet({ oldMerges: [
    { row: 1, col: 1, rows: 1, cols: 12 },
    { row: 4, col: 1, rows: 2, cols: 2 },
    { row: 28, col: 10, rows: 1, cols: 3 }
  ] });
  sheet.cells.set('40:15', 'stale');
  const { ss } = dashboardSpreadsheet(sheet);
  const data = dashboardData();
  for (let index = 0; index < 10; index++) {
    assert.doesNotThrow(() => context.renderIntegratedDashboard_(ss,
      data.order, data.stock, data.picking, data.operation));
  }
  assert.equal(sheet.cells.has('40:15'), false, 'full-sheet content reset must remove stale cells');
  assert.equal(sheet.merges.length, 66, 'each refresh must leave only the current non-overlapping merges');
  assert.equal(sheet.events.filter(event => event === 'breakApart:1:1:40:15').length, 10);
  assert.equal(sheet.events.some(event => event.startsWith('breakApart:1:1:1:1')), false,
    'getDataRange must not be used for the critical merge reset');
});

test('all merge ranges stay within the 12-column grid and quick actions merge before writing', () => {
  const sheet = createDashboardSheet();
  const { ss } = dashboardSpreadsheet(sheet);
  const data = dashboardData();
  context.renderIntegratedDashboard_(ss, data.order, data.stock, data.picking, data.operation);
  for (const range of sheet.merges) assert.ok(range.col >= 1 && range.col + range.cols - 1 <= 12);
  for (const start of [1, 4, 7, 10]) {
    const rangeKey = `28:${start}:1:3`;
    assert.ok(sheet.merges.some(range => `${range.row}:${range.col}:${range.rows}:${range.cols}` === rangeKey));
    assert.ok(sheet.events.indexOf(`merge:${rangeKey}`) < sheet.events.indexOf(`value:${rangeKey}`));
  }
  for (const row of [9, 10, 11, 12, 13, 16, 17, 18, 19, 20]) {
    assert.ok(sheet.merges.some(range => range.row === row && range.col === 1 && range.cols === 4));
    assert.ok(sheet.merges.some(range => range.row === row && range.col === 5 && range.cols === 2));
    assert.ok(sheet.merges.some(range => range.row === row && range.col === 7 && range.cols === 4));
    assert.ok(sheet.merges.some(range => range.row === row && range.col === 11 && range.cols === 2));
  }
});

test('merge reset failure is reported with its rendering stage and rendering does not continue', () => {
  const sheet = createDashboardSheet();
  sheet.getRange = () => ({ breakApart() { throw new Error('reset denied'); } });
  const { ss } = dashboardSpreadsheet(sheet);
  const data = dashboardData();
  assert.throws(() => context.renderIntegratedDashboard_(ss, data.order, data.stock, data.picking, data.operation),
    /\[병합 초기화\] reset denied/);
  assert.equal(sheet.merges.length, 0);
});

test('dashboard refresh returns success after an already-rendered dashboard', () => {
  const sheet = createDashboardSheet();
  const { ss } = dashboardSpreadsheet(sheet);
  const data = dashboardData();
  const originals = {
    consoleSS_: context.consoleSS_, collectOrderStatus_: context.collectOrderStatus_,
    collectStockStatus_: context.collectStockStatus_, collectPickingStatus_: context.collectPickingStatus_,
    collectOperationStatus_: context.collectOperationStatus_, alert_: context.alert_
  };
  context.consoleSS_ = () => ss;
  context.collectOrderStatus_ = () => data.order;
  context.collectStockStatus_ = () => data.stock;
  context.collectPickingStatus_ = () => data.picking;
  context.collectOperationStatus_ = () => data.operation;
  context.alert_ = () => {};
  try {
    assert.equal(context.D0_대시보드전체갱신(true), '✅ 통합 대시보드');
    assert.equal(context.D0_대시보드전체갱신(true), '✅ 통합 대시보드');
  } finally {
    Object.assign(context, originals);
  }
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
