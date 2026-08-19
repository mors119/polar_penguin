const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync(path.resolve(__dirname, '..', 'src', 'S9 작업지시서.js'), 'utf8');
const context = vm.createContext({ console });
vm.runInContext(source, context);

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

  const first = context.S9_피킹PDF생성('PK-20260819-001', outputRoot);
  const second = context.S9_피킹PDF생성('PK-20260819-001', outputRoot);
  assert.equal(first.생성, true);
  assert.equal(second.생성, false);
  assert.equal(second.재사용, true);
  assert.equal(files.size, 1);
});
