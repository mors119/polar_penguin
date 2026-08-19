const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');

function createContext() {
  const context = vm.createContext({ console });
  for (const file of ['S0 공통.js', 'S11 백업정리.js']) {
    vm.runInContext(fs.readFileSync(path.join(root, 'src', file), 'utf8'), context);
  }
  context.Utilities = { formatDate: (date, _tz, format) => {
    const value = new Date(date);
    if (format === 'yyyy-MM') return '2026-08';
    if (format === 'yyyy-MM-dd') return '2026-08-19';
    if (format === 'yyyy-MM-dd HHmmss') return '2026-08-19 231500';
    return '2026-08-19 23:15:00';
  } };
  context.tz_ = () => 'Asia/Seoul';
  context.writeOpLog_ = () => {};
  context.D0_대시보드전체갱신 = () => {};
  context.withLock_ = fn => fn();
  return context;
}

class FakeSheet {
  constructor(name, headers, rows = []) { this.name = name; this.values = [headers.slice(), ...rows.map(row => row.slice())]; }
  getName() { return this.name; }
  getLastRow() { return this.values.length; }
  getLastColumn() { return this.values[0].length; }
  getDataRange() { return this.getRange(1, 1, this.getLastRow(), this.getLastColumn()); }
  getRange(row, col, rowCount = 1, colCount = 1) {
    return {
      getValues: () => Array.from({ length: rowCount }, (_, r) => Array.from({ length: colCount }, (_, c) =>
        (this.values[row - 1 + r] || [])[col - 1 + c] ?? '')),
      setValues: values => {
        values.forEach((line, r) => {
          const target = row - 1 + r;
          while (this.values.length <= target) this.values.push(new Array(this.getLastColumn()).fill(''));
          line.forEach((value, c) => { this.values[target][col - 1 + c] = value; });
        });
      },
      setNumberFormat() { return this; }
    };
  }
  deleteRows(start, count) { this.values.splice(start - 1, count); }
}

function tableFor(context, role, sheet) {
  const headers = sheet.values[0], rows = sheet.values.slice(1);
  return { role, sheet, headers, rows, width: headers.length,
    headerIndex: Object.fromEntries(headers.map((header, index) => [context.normKey_(header), index])) };
}

function installOperationalTables(context) {
  const orderHeaders = ['주문번호', '품목별 주문번호', '상품품목코드', '수량', '주문상태', '확정일시', '취소일시', '피킹지시번호'];
  const sheets = {
    [context.ROLE.주문]: new FakeSheet(context.ROLE.주문, orderHeaders, [
      ['A001', 'ITEM-OLD-SHIP', 'SKU1', 1, '출고완료', '2026-07-01T00:00:00Z', '', 'PK-SHARED'],
      ['A002', 'ITEM-RECENT', 'SKU1', 1, '출고완료', '2026-08-10T00:00:00Z', '', 'PK-SHARED'],
      ['A003', 'ITEM-OLD-CANCEL', 'SKU2', 2, '취소', '', '2026-07-01T00:00:00Z', 'PK-CANCEL'],
      ['A004', 'ITEM-RESERVED', 'SKU3', 1, '예약', '2026-01-01T00:00:00Z', '', 'PK-ERROR']
    ]),
    [context.ROLE.헤더]: new FakeSheet(context.ROLE.헤더, ['피킹지시번호', '주문번호', '상태'], [
      ['PK-SHARED', 'A001', '완료'], ['PK-SHARED', 'A002', '완료'], ['PK-CANCEL', 'A003', '취소'],
      ['PK-ERROR', 'A004', '출력오류']
    ]),
    [context.ROLE.라인]: new FakeSheet(context.ROLE.라인, ['피킹지시번호', '주문번호', '라인상태'], [
      ['PK-SHARED', 'A001', '완료'], ['PK-SHARED', 'A002', '완료'], ['PK-CANCEL', 'A003', '취소'],
      ['PK-ERROR', 'A004', '미처리']
    ]),
    archive: new FakeSheet(context.CONSOLE.처리주문아카이브, Array.from(context.ORDER_ARCHIVE_HEADERS), [])
  };
  context.readTable_ = role => tableFor(context, role, sheets[role]);
  context.consoleSS_ = () => ({ getSheetByName: name => name === context.CONSOLE.처리주문아카이브 ? sheets.archive : sheets[name] });
  return sheets;
}

test('backup failure is a hard boundary and runs no destructive cleanup', () => {
  const context = createContext();
  const calls = [];
  context.param_ = key => key === '정리보존일수' ? 30 : 'ops@example.com';
  context.maintenancePreflight_ = () => true;
  context.createBackupSnapshot_ = () => { throw new Error('copy failed'); };
  context.archiveAndCleanupFinalizedOrders_ = () => { calls.push('orders'); };
  context.cleanupPickingHistory_ = () => { calls.push('picking'); };
  context.cleanupFolderByRetention_ = () => { calls.push('files'); };
  context.sendSystemNotification_ = (level, title, details) => {
    calls.push(`${level}:${title}:${details.결과}`); return { sent: true };
  };
  const result = context.runBackupAndCleanup_({ now: new Date('2026-08-19T14:15:00Z') });
  assert.equal(result.aborted, true);
  assert.equal(result.sheets.ordersDeleted, 0);
  assert.equal(result.sheets.pickingLinesDeleted, 0);
  assert.equal(result.files.successDeleted, 0);
  assert.deepEqual(calls, ['ERROR:백업 실패 - 정리 중단:백업에 실패하여 정리를 중단했습니다. 삭제된 데이터는 없습니다.']);
});

test('verified backup uses Backup month folder, source spreadsheet, and timestamped immutable name', () => {
  const context = createContext();
  const monthFolder = { getId: () => 'month-folder', getName: () => '2026-08' };
  let monthExists = false;
  const rootFolder = { getFoldersByName: () => iterator(monthExists ? [monthFolder] : []), createFolder: name => {
    assert.equal(name, '2026-08'); monthExists = true; return monthFolder;
  } };
  const copies = new Map(); let copySequence = 0;
  const source = { makeCopy: (name, folder) => {
    assert.equal(name, 'Polar Penguin Backup 2026-08-19 231500');
    assert.equal(folder, monthFolder);
    const id = `backup-file-${++copySequence}`;
    const copy = { getId: () => id, getParents: () => iterator([monthFolder]),
      getUrl: () => `https://drive.google.com/${id}` };
    copies.set(id, copy); return copy;
  } };
  const properties = new Map();
  context.DriveApp = { getFileById: id => id === 'operation-file' ? source : copies.get(id) };
  context.SpreadsheetApp = { openById: id => ({ getId: () => id, getSheetByName: () => ({}) }) };
  context.PropertiesService = { getScriptProperties: () => ({
    setProperties: values => Object.entries(values).forEach(([key, value]) => properties.set(key, value))
  }) };
  const result = context.createBackupSnapshot_({ now: new Date('2026-08-19T14:15:00Z'),
    spreadsheet: { getId: () => 'operation-file' }, backupRoot: rootFolder });
  assert.equal(result.fileId, 'backup-file-1');
  assert.equal(result.monthFolderId, 'month-folder');
  assert.equal(result.sourceSpreadsheetId, 'operation-file');
  assert.match(result.fileName, /2026-08-19 231500/);
  assert.equal(properties.get('최근백업파일ID'), 'backup-file-1');
  const repeated = context.createBackupSnapshot_({ now: new Date('2026-08-19T14:15:01Z'),
    spreadsheet: { getId: () => 'operation-file' }, backupRoot: rootFolder });
  assert.equal(repeated.fileId, 'backup-file-2', 'every maintenance run creates a new snapshot');
});

test('old final orders archive without PII, reservations/recent orders stay, and shared batch remains', () => {
  const context = createContext(), sheets = installOperationalTables(context);
  const cutoff = new Date('2026-07-20T00:00:00Z');
  const cleaned = context.archiveAndCleanupFinalizedOrders_(cutoff, new Date('2026-08-19T00:00:00Z'));
  assert.equal(cleaned.ordersDeleted, 2);
  assert.equal(cleaned.archive.inserted, 2);
  assert.deepEqual(sheets[context.ROLE.주문].values.slice(1).map(row => row[1]), ['ITEM-RECENT', 'ITEM-RESERVED']);
  assert.equal(sheets.archive.values.length, 3);
  assert.equal(sheets.archive.values.some(row => row[1] === 'ITEM-RECENT'), false);
  assert.equal(sheets.archive.values.some(row => row[1] === 'ITEM-RESERVED'), false);
  assert.equal(sheets.archive.values[0].includes('수령인 주소'), false);
  assert.equal(sheets.archive.values[0].includes('수령인 휴대전화'), false);

  const picking = context.cleanupPickingHistory_(cleaned.deletedInstructions);
  assert.equal(picking.headerRowsDeleted, 1);
  assert.equal(picking.lineRowsDeleted, 1);
  assert.deepEqual(sheets[context.ROLE.헤더].values.slice(1).map(row => row[0]), ['PK-SHARED', 'PK-SHARED', 'PK-ERROR']);
  assert.deepEqual(sheets[context.ROLE.라인].values.slice(1).map(row => row[0]), ['PK-SHARED', 'PK-SHARED', 'PK-ERROR']);

  const second = context.archiveAndCleanupFinalizedOrders_(cutoff, new Date('2026-08-20T00:00:00Z'));
  assert.equal(second.archive.inserted, 0);
  assert.equal(second.deletedInstructions['PK-CANCEL'], true,
    'an archived instruction remains retryable after active order rows were already removed');
  const retriedPicking = context.cleanupPickingHistory_(second.deletedInstructions);
  assert.equal(retriedPicking.removedInstructions['PK-CANCEL'], true,
    'a later run can finish Output cleanup even when picking rows were already removed');
  assert.equal(sheets.archive.values.filter(row => row[1] === 'ITEM-OLD-SHIP').length, 1);
});

test('retention trashes only old allowed files, removes empty child folders, and never touches Input', () => {
  const context = createContext();
  const old = new FakeFile('PK-OLD.pdf', '2026-07-01T00:00:00Z');
  const recent = new FakeFile('PK-RECENT.pdf', '2026-08-15T00:00:00Z');
  const unrelated = new FakeFile('ACTIVE.pdf', '2026-07-01T00:00:00Z');
  const dateFolder = new FakeFolder('2026-07-01', [old, recent, unrelated]);
  const oldOnly = new FakeFile('PK-ONLY.pdf', '2026-07-01T00:00:00Z');
  const emptyAfterCleanup = new FakeFolder('2026-06-01', [oldOnly]);
  const output = new FakeFolder('Output', [], [dateFolder, emptyAfterCleanup]);
  const inputOld = new FakeFile('stale-input.csv', '2026-01-01T00:00:00Z');
  const input = new FakeFolder('Input', [inputOld]);
  const result = context.cleanupFolderByRetention_(output, new Date('2026-07-20T00:00:00Z'), {
    allowedNames: { 'PK-OLD.pdf': true, 'PK-ONLY.pdf': true }
  });
  assert.equal(result.filesDeleted, 2);
  assert.equal(result.foldersDeleted, 1);
  assert.equal(old.trashed, true);
  assert.equal(recent.trashed, false);
  assert.equal(unrelated.trashed, false);
  assert.equal(inputOld.trashed, false, 'Input is never passed to retention cleanup');
  assert.equal(output.trashed, false);
});

test('successful maintenance sends one backup-link summary and Gmail failure does not undo cleanup', () => {
  const context = createContext(), sent = [];
  context.param_ = key => key === '정리보존일수' ? 30 : 'ops@example.com';
  context.maintenancePreflight_ = () => true;
  context.createBackupSnapshot_ = () => ({ fileId: 'backup-1', fileName: 'Polar Penguin Backup 2026-08-19 231500', url: 'https://drive/backup-1' });
  context.archiveAndCleanupFinalizedOrders_ = () => ({
    archive: { inserted: 4, skippedExisting: 0, failed: 0 }, ordersDeleted: 4, deletedInstructions: { 'PK-1': true }
  });
  context.cleanupPickingHistory_ = () => ({ headerRowsDeleted: 1, lineRowsDeleted: 4, removedInstructions: { 'PK-1': true } });
  context.inputFolder_ = () => ({});
  let folderCall = 0;
  context.cleanupFolderByRetention_ = () => ({ filesDeleted: ++folderCall, foldersDeleted: 0 });
  context.GmailApp = { sendEmail: (recipient, subject, body) => sent.push({ recipient, subject, body }) };
  let result = context.runBackupAndCleanup_({ now: new Date('2026-08-19T14:15:00Z') });
  assert.equal(result.emailSent, true);
  assert.equal(sent.length, 1);
  assert.match(sent[0].body, /Polar Penguin Backup 2026-08-19 231500/);
  assert.match(sent[0].body, /https:\/\/drive\/backup-1/);
  assert.match(sent[0].body, /삭제 주문 라인: 4/);
  assert.match(sent[0].body, /삭제 Output 파일: 3/);

  context.GmailApp = { sendEmail: () => { throw new Error('quota exceeded'); } };
  folderCall = 0;
  result = context.runBackupAndCleanup_({ now: new Date('2026-08-20T14:15:00Z') });
  assert.equal(result.aborted, false);
  assert.equal(result.backup.fileId, 'backup-1');
  assert.equal(result.sheets.ordersDeleted, 4);
  assert.equal(result.emailSent, false);
  assert.match(result.notificationError, /quota exceeded/);
});

test('central notification handles blank configuration and Gmail failure without throwing', () => {
  const context = createContext(), logs = [];
  context.writeOpLog_ = (...args) => logs.push(args.join('/'));
  context.param_ = () => '';
  assert.equal(context.sendSystemNotification_('ERROR', 'Input 처리 실패', {}).sent, false);
  assert.match(logs[0], /알림이메일 미설정/);

  context.param_ = () => 'ops@example.com';
  context.GmailApp = { sendEmail: () => { throw new Error('quota'); } };
  const failed = context.sendSystemNotification_('INFO', '백업 및 정리 완료', 'backup link');
  assert.equal(failed.sent, false);
  assert.match(failed.reason, /quota/);
});

function iterator(items) {
  let index = 0;
  return { hasNext: () => index < items.length, next: () => items[index++] };
}

class FakeFile {
  constructor(name, updated) { this.name = name; this.updated = new Date(updated); this.trashed = false; }
  getName() { return this.name; }
  getLastUpdated() { return this.updated; }
  setTrashed(value) { this.trashed = value; }
}

class FakeFolder {
  constructor(name, files = [], folders = []) { this.name = name; this.files = files; this.folders = folders; this.trashed = false; }
  getName() { return this.name; }
  getFiles() { return iterator(this.files.filter(file => !file.trashed)); }
  getFolders() { return iterator(this.folders.filter(folder => !folder.trashed)); }
  setTrashed(value) { this.trashed = value; }
}
