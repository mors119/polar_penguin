const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const rootPath = path.resolve(__dirname, '..');
const context = vm.createContext({ console });
for (const file of ['S0 공통.js', 'S0 자동설치.js', 'S0 설치.js']) {
  vm.runInContext(fs.readFileSync(path.join(rootPath, 'src', file), 'utf8'), context);
}

function plain(value) { return JSON.parse(JSON.stringify(value)); }
function iterator(items) {
  let index = 0;
  return { hasNext: () => index < items.length, next: () => items[index++] };
}

test('Drive ID extraction accepts a folder URL or raw ID', () => {
  const id = '1AbCdEfGhIjKlMnOpQrStUvWxYz123';
  assert.equal(context.extractDriveId_(id), id);
  assert.equal(context.extractDriveId_(`https://drive.google.com/drive/folders/${id}?usp=sharing`), id);
  assert.equal(context.extractDriveId_('invalid'), '');
});

test('config merge adds missing values, repairs invalid managed values, and preserves user defaults', () => {
  const values = [
    ['구분', '키', '값', '비고'],
    ['파일ID', '피킹헤더', 'broken', ''],
    ['파라미터', '폴링주기(분)', 15, '사용자 설정']
  ];
  const managed = [['파일ID', '피킹헤더', 'new-id', '', (value) => value === 'valid']];
  const defaults = [
    ['파라미터', '폴링주기(분)', 5, '기본값'],
    ['파라미터', '예약키워드', '예약', '기본값']
  ];
  const result = plain(context.mergeSetupConfig_(values, managed, defaults));
  assert.deepEqual(result.updates, [{ row: 2, value: 'new-id' }]);
  assert.deepEqual(result.additions, [defaults[1]]);
  assert.equal(values[2][2], 15);
});

test('missing header detection is idempotent and accepts aliases only for populated sheets', () => {
  const required = ['상품품목코드', '상품명', '가용재고'];
  assert.deepEqual(
    plain(context.missingInstallHeaders_(['품목코드', '상품명'], required, true)),
    ['가용재고']
  );
  assert.deepEqual(
    plain(context.missingInstallHeaders_(required, required, true)),
    []
  );
});

test('setupSystem installs, reruns safely, repairs one missing spreadsheet, and preserves data', () => {
  let sequence = 1;
  const spreadsheets = new Map();
  const folders = new Map();
  const files = new Map();
  const properties = new Map();
  let triggers = [];
  let dashboardRefreshes = 0;

  class FakeRange {
    constructor(sheet, row, col, rows = 1, cols = 1) {
      this.sheet = sheet; this.row = row; this.col = col; this.rows = rows; this.cols = cols;
    }
    getValues() {
      return Array.from({ length: this.rows }, (_, r) =>
        Array.from({ length: this.cols }, (_, c) => this.sheet.value(this.row + r, this.col + c)));
    }
    getValue() { return this.sheet.value(this.row, this.col); }
    setValues(values) {
      values.forEach((line, r) => line.forEach((value, c) => this.sheet.set(this.row + r, this.col + c, value)));
      return this;
    }
    setValue(value) { this.sheet.set(this.row, this.col, value); return this; }
  }
  ['setFontWeight', 'setBackground', 'setFontColor', 'setVerticalAlignment', 'setNumberFormat',
    'setDataValidation', 'setHelpText'].forEach((method) => {
    FakeRange.prototype[method] = function noop() { return this; };
  });

  class FakeSheet {
    constructor(name, parent) {
      this.name = name; this.parent = parent; this.cells = new Map(); this.maxRows = 1000; this.maxCols = 26;
    }
    key(row, col) { return `${row}:${col}`; }
    value(row, col) { return this.cells.get(this.key(row, col)) ?? ''; }
    set(row, col, value) { this.cells.set(this.key(row, col), value); }
    getName() { return this.name; }
    setName(name) { this.name = name; return this; }
    getParent() { return this.parent; }
    getLastRow() { return Math.max(0, ...[...this.cells.keys()].map((key) => Number(key.split(':')[0]))); }
    getLastColumn() { return Math.max(0, ...[...this.cells.keys()].map((key) => Number(key.split(':')[1]))); }
    getMaxRows() { return this.maxRows; }
    getMaxColumns() { return this.maxCols; }
    insertColumnsAfter(after, count) { this.maxCols = Math.max(this.maxCols, after + count); }
    getRange(row, col, rows, cols) { return new FakeRange(this, row, col, rows || 1, cols || 1); }
    getDataRange() { return this.getRange(1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); }
    setFrozenRows() { return this; }
    autoResizeColumns() { return this; }
  }

  class FakeSpreadsheet {
    constructor(name) {
      this.id = `spreadsheet-${String(sequence++).padStart(20, '0')}`;
      this.name = name;
      this.sheets = [new FakeSheet('Sheet1', this)];
      spreadsheets.set(this.id, this);
    }
    getId() { return this.id; }
    getName() { return this.name; }
    getSheets() { return this.sheets.slice(); }
    getSheetByName(name) { return this.sheets.find((sheet) => sheet.name === name) || null; }
    insertSheet(name) { const sheet = new FakeSheet(name, this); this.sheets.push(sheet); return sheet; }
  }

  class FakeFile {
    constructor(ss) { this.id = ss.getId(); this.name = ss.getName(); this.parent = null; files.set(this.id, this); }
    getId() { return this.id; }
    getName() { return this.name; }
    getMimeType() { return 'application/vnd.google-apps.spreadsheet'; }
    moveTo(folder) {
      if (this.parent) this.parent.files = this.parent.files.filter((file) => file !== this);
      this.parent = folder;
      if (!folder.files.includes(this)) folder.files.push(this);
      return this;
    }
  }

  class FakeFolder {
    constructor(name) {
      this.id = `folder-${String(sequence++).padStart(20, '0')}`;
      this.name = name; this.folders = []; this.files = [];
      folders.set(this.id, this);
    }
    getId() { return this.id; }
    getName() { return this.name; }
    getFoldersByName(name) { return iterator(this.folders.filter((folder) => folder.name === name)); }
    createFolder(name) { const folder = new FakeFolder(name); this.folders.push(folder); return folder; }
    getFilesByName(name) { return iterator(this.files.filter((file) => file.name === name)); }
  }

  const root = new FakeFolder('Polar Penguin');
  properties.set('ROOT_FOLDER_URL', `https://drive.google.com/drive/folders/${root.getId()}`);

  context.MimeType = { GOOGLE_SHEETS: 'application/vnd.google-apps.spreadsheet' };
  context.DriveApp = {
    getFolderById: (id) => {
      if (!folders.has(id)) throw new Error('folder missing');
      return folders.get(id);
    },
    getFileById: (id) => {
      if (!files.has(id)) throw new Error('file missing');
      return files.get(id);
    }
  };
  context.SpreadsheetApp = {
    create: (name) => { const ss = new FakeSpreadsheet(name); new FakeFile(ss); return ss; },
    openById: (id) => {
      if (!spreadsheets.has(id)) throw new Error('spreadsheet missing');
      return spreadsheets.get(id);
    },
    getActiveSpreadsheet: () => null,
    newDataValidation: () => ({
      requireValueInList() { return this; }, setAllowInvalid() { return this; },
      setHelpText() { return this; }, build() { return {}; }
    })
  };
  context.PropertiesService = { getScriptProperties: () => ({
    getProperty: (key) => properties.get(key) || null,
    setProperty: (key, value) => properties.set(key, value),
    setProperties: (values) => Object.entries(values).forEach(([key, value]) => properties.set(key, value))
  }) };
  context.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) };
  context.ScriptApp = {
    getProjectTriggers: () => triggers.slice(),
    deleteTrigger: (trigger) => { triggers = triggers.filter((item) => item !== trigger); },
    newTrigger: (handler) => {
      const trigger = { getHandlerFunction: () => handler };
      return {
        timeBased() { return this; }, everyMinutes(minutes) { trigger.minutes = minutes; return this; },
        forSpreadsheet() { return this; }, onOpen() { return this; },
        create() { triggers.push(trigger); return trigger; }
      };
    }
  };
  context.writeOpLog_ = () => {};
  context.alert_ = () => {};
  context.D0_대시보드전체갱신 = () => { dashboardRefreshes++; return '✅ all'; };

  const first = context.setupSystem();
  assert.equal(properties.get('ROOT_FOLDER_ID'), root.getId());
  assert.equal(first.folders.filter((item) => item.created).length, 8);
  assert.equal(first.spreadsheets.filter((item) => item.created).length, 5);
  assert.equal(spreadsheets.size, 5);
  assert.deepEqual(triggers.map((trigger) => trigger.getHandlerFunction()).sort(), ['onOpen', 'syncAndRefresh']);

  const consoleSs = spreadsheets.get(properties.get('CONSOLE_SS_ID'));
  const configSheet = consoleSs.getSheetByName('설정');
  const configValues = configSheet.getDataRange().getValues();
  const pollingRow = configValues.findIndex((row) => row[0] === '파라미터' && row[1] === '폴링주기(분)') + 1;
  configSheet.getRange(pollingRow, 3).setValue(15);

  const orderSs = spreadsheets.get(properties.get('SPREADSHEET_ID_ORDERS'));
  const orderSheet = orderSs.getSheetByName('주문(완료)');
  orderSheet.getRange(2, 1).setValue('ORDER-KEEP');
  const masterSs = spreadsheets.get(properties.get('SPREADSHEET_ID_MASTER'));
  masterSs.getSheetByName('상품마스터').getRange(2, 1).setValue('SKU-KEEP');

  const second = context.setupSystem();
  assert.equal(second.folders.filter((item) => item.created).length, 0);
  assert.equal(second.spreadsheets.filter((item) => item.created).length, 0);
  assert.equal(spreadsheets.size, 5);
  assert.equal(configSheet.getRange(pollingRow, 3).getValue(), 15);
  assert.equal(orderSheet.getRange(2, 1).getValue(), 'ORDER-KEEP');
  assert.equal(masterSs.getSheetByName('상품마스터').getRange(2, 1).getValue(), 'SKU-KEEP');
  assert.equal(triggers.length, 2);
  assert.equal(triggers.find((trigger) => trigger.getHandlerFunction() === 'syncAndRefresh').minutes, 15);

  const oldHeaderId = properties.get('SPREADSHEET_ID_PICKING_HEADER');
  const oldHeaderFile = files.get(oldHeaderId);
  oldHeaderFile.parent.files = oldHeaderFile.parent.files.filter((file) => file !== oldHeaderFile);
  files.delete(oldHeaderId);
  spreadsheets.delete(oldHeaderId);

  const repaired = context.setupSystem();
  assert.equal(repaired.spreadsheets.filter((item) => item.created).length, 1);
  const repairedHeaderId = properties.get('SPREADSHEET_ID_PICKING_HEADER');
  assert.notEqual(repairedHeaderId, oldHeaderId);
  const repairedConfig = configSheet.getDataRange().getValues();
  const headerConfig = repairedConfig.find((row) => row[0] === '파일ID' && row[1] === '피킹헤더');
  assert.equal(headerConfig[2], repairedHeaderId);
  assert.equal(spreadsheets.size, 5);
  assert.equal(triggers.length, 2);
  assert.equal(dashboardRefreshes, 3);
});
