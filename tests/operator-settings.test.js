const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const rootPath = path.resolve(__dirname, '..');
const context = vm.createContext({ console });
for (const file of ['S0 공통.js', 'S0 설치.js']) {
  vm.runInContext(fs.readFileSync(path.join(rootPath, 'src', file), 'utf8'), context);
}

function createSettingsFixture() {
  const rows = [
    ['구분', '키', '값', '비고'],
    ['파라미터', '폴링주기(분)', 5, 'poll'],
    ['파라미터', '지시번호접두어', 'PK', 'prefix'],
    ['파라미터', '재고경고임계치', 3, 'stock'],
    ['파라미터', '알림이메일', '', 'email'],
    ['파라미터', '정리보존일수', 30, 'retention'],
    ['파라미터', '통합Input폴더ID', 'folder-keep', 'unrelated parameter'],
    ['별칭', '상품품목코드', '품목코드,SKU', 'unrelated row']
  ];
  const writes = [];
  const sheet = {
    getLastRow: () => rows.length,
    getRange(row, col, height = 1, width = 1) {
      return {
        getValue: () => rows[row - 1][col - 1],
        getValues: () => Array.from({ length: height }, (_, r) =>
          Array.from({ length: width }, (_, c) => rows[row - 1 + r][col - 1 + c])),
        setValue(value) { rows[row - 1][col - 1] = value; writes.push([row, col, value]); return this; }
      };
    }
  };
  context.consoleSS_ = () => ({ getSheetByName: name => name === '설정' ? sheet : null });
  context.withLock_ = fn => fn();
  let cacheResets = 0;
  context.resetConfigCache_ = () => { cacheResets++; };
  return { rows, writes, cacheResets: () => cacheResets };
}

test('operator settings RPC reads the current five real parameter values', () => {
  createSettingsFixture();
  assert.deepEqual(JSON.parse(JSON.stringify(context.getOperatorSettings())), {
    '폴링주기(분)': 5,
    '지시번호접두어': 'PK',
    '재고경고임계치': 3,
    '알림이메일': '',
    '정리보존일수': 30
  });
});

test('operator settings save updates only intended cells and preserves unrelated config rows', () => {
  const fixture = createSettingsFixture();
  const unrelatedBefore = JSON.stringify(fixture.rows.slice(6));
  const result = context.saveOperatorSettings({
    '폴링주기(분)': '5', '지시번호접두어': 'PP', '재고경고임계치': '8',
    '알림이메일': 'ops@example.com', '정리보존일수': '45'
  });
  assert.deepEqual(Array.from(result.changedFields), ['지시번호접두어', '재고경고임계치', '알림이메일', '정리보존일수']);
  assert.ok(fixture.writes.every(write => write[1] === 3 && write[0] >= 2 && write[0] <= 6));
  assert.equal(JSON.stringify(fixture.rows.slice(6)), unrelatedBefore);
  assert.equal(fixture.cacheResets(), 1);
  assert.equal(result.pollingChanged, false);
});

test('operator settings validation rejects invalid values and accepts blank email', () => {
  createSettingsFixture();
  const valid = context.validateOperatorSettings_({
    '폴링주기(분)': '1', '지시번호접두어': 'PK', '재고경고임계치': '0',
    '알림이메일': '   ', '정리보존일수': '1'
  });
  assert.equal(valid['알림이메일'], '');

  const base = { '폴링주기(분)': '5', '지시번호접두어': 'PK', '재고경고임계치': '3',
    '알림이메일': '', '정리보존일수': '30' };
  for (const [key, value, message] of [
    ['폴링주기(분)', '0', /1 이상의 정수/],
    ['폴링주기(분)', '1.5', /1 이상의 정수/],
    ['지시번호접두어', '', /값을 입력/],
    ['지시번호접두어', 'ABCDEFGHIJKLM', /12자 이내/],
    ['재고경고임계치', '-1', /0 이상의 정수/],
    ['알림이메일', 'not-an-email', /올바른 이메일/],
    ['정리보존일수', '0', /1 이상의 정수/]
  ]) {
    assert.throws(() => context.validateOperatorSettings_({ ...base, [key]: value }), message);
  }
});

test('polling change returns the required repair instruction without rebuilding triggers', () => {
  const fixture = createSettingsFixture();
  const result = context.saveOperatorSettings({
    '폴링주기(분)': '10', '지시번호접두어': 'PK', '재고경고임계치': '3',
    '알림이메일': '', '정리보존일수': '30'
  });
  assert.equal(result.pollingChanged, true);
  assert.deepEqual(Array.from(result.changedFields), ['폴링주기(분)']);
  assert.equal(result.message, '폴링주기가 변경되었습니다.\n변경된 자동 확인 주기를 적용하려면\n⚙ 관리 → 시스템 설치 / 복구를 실행하세요.');
  assert.deepEqual(fixture.writes, [[2, 3, 10]]);
});
