const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const srcDir = path.join(root, 'src');

function extractGoogleScriptRunCalls(htmlContent) {
  const handlerNames = new Set(['withSuccessHandler', 'withFailureHandler', 'withUserObject']);
  const calls = [];
  
  // Find each occurrence of google.script.run
  let index = 0;
  while ((index = htmlContent.indexOf('google.script.run', index)) !== -1) {
    let pos = index + 'google.script.run'.length;
    while (pos < htmlContent.length) {
      // Skip whitespace
      while (pos < htmlContent.length && /\s/.test(htmlContent[pos])) pos++;
      if (htmlContent[pos] !== '.') break;
      pos++; // skip '.'
      
      // Read identifier
      const identStart = pos;
      while (pos < htmlContent.length && /[A-Za-z0-9_$]/.test(htmlContent[pos])) pos++;
      const ident = htmlContent.slice(identStart, pos);
      
      // Skip whitespace to '('
      while (pos < htmlContent.length && /\s/.test(htmlContent[pos])) pos++;
      if (htmlContent[pos] !== '(') break;
      
      // Find matching ')'
      let depth = 0;
      let inQuote = null;
      let parenStart = pos;
      while (pos < htmlContent.length) {
        const ch = htmlContent[pos];
        if (inQuote) {
          if (ch === '\\') { pos += 2; continue; }
          if (ch === inQuote) inQuote = null;
        } else if (ch === '"' || ch === "'" || ch === '`') {
          inQuote = ch;
        } else if (ch === '(') {
          depth++;
        } else if (ch === ')') {
          depth--;
          if (depth === 0) { pos++; break; }
        }
        pos++;
      }
      
      if (!handlerNames.has(ident)) {
        calls.push(ident);
        break; // terminal server call reached
      }
    }
    index += 'google.script.run'.length;
  }
  return calls;
}

test('HTML client-server RPC contract: all google.script.run functions exist globally in server js without underscore suffix', () => {
  const jsFiles = fs.readdirSync(srcDir).filter(f => f.endsWith('.js'));
  const htmlFiles = fs.readdirSync(srcDir).filter(f => f.endsWith('.html'));

  const context = vm.createContext({
    console,
    Logger: { log: () => {} },
    SpreadsheetApp: {},
    HtmlService: {},
    PropertiesService: {},
    ScriptApp: {},
    GmailApp: {},
    Utilities: {}
  });

  for (const jsFile of jsFiles) {
    const code = fs.readFileSync(path.join(srcDir, jsFile), 'utf8');
    vm.runInContext(code, context);
  }

  for (const htmlFile of htmlFiles) {
    const content = fs.readFileSync(path.join(srcDir, htmlFile), 'utf8');
    const calls = extractGoogleScriptRunCalls(content);

    assert.ok(calls.length > 0, `${htmlFile} should have at least one google.script.run call`);

    for (const funcName of calls) {
      assert.equal(
        funcName.endsWith('_'),
        false,
        `${htmlFile} calls private underscore function "${funcName}" via google.script.run`
      );

      assert.equal(
        typeof context[funcName],
        'function',
        `${htmlFile} calls "${funcName}" which is not defined in server source files`
      );
    }
  }
});

test('onOpen menu handlers all exist as global functions in server js', () => {
  const jsFiles = fs.readdirSync(srcDir).filter(f => f.endsWith('.js'));
  const context = vm.createContext({
    console,
    Logger: { log: () => {} },
    SpreadsheetApp: {},
    HtmlService: {},
    PropertiesService: {},
    ScriptApp: {},
    GmailApp: {},
    Utilities: {}
  });

  for (const jsFile of jsFiles) {
    const code = fs.readFileSync(path.join(srcDir, jsFile), 'utf8');
    vm.runInContext(code, context);
  }

  const expectedHandlers = [
    'processInput',
    '선택_주문취소',
    '예약상품_입고관리',
    'S9_1_작업지시서출력',
    '위치_미지정상품관리',
    'D0_대시보드전체갱신',
    '진단_시트구조',
    'setupSystem',
    '설정_열기',
    '백업_및_정리',
    '정리_로그'
  ];

  for (const handler of expectedHandlers) {
    assert.equal(
      typeof context[handler],
      'function',
      `Menu handler "${handler}" must exist as a global function`
    );
  }
});

test('public server RPCs for Reservation and Cancellation work as expected', () => {
  const context = vm.createContext({
    console,
    Logger: { log: () => {} },
    SpreadsheetApp: {},
    HtmlService: {},
    PropertiesService: {},
    ScriptApp: {},
    GmailApp: {},
    Utilities: {}
  });

  for (const jsFile of ['S0 공통.js', 'S0 설치.js', 'S3 주문확정.js', 'S7 주문취소.js', 'S8 예약관리.js']) {
    const code = fs.readFileSync(path.join(srcDir, jsFile), 'utf8');
    vm.runInContext(code, context);
  }

  assert.equal(typeof context.getReservationProductSummary, 'function');
  assert.equal(typeof context.getReservationInboundSummary, 'function');
  assert.equal(typeof context.getReservationInboundPreview, 'function');
  assert.equal(typeof context.getReservationProductPreview, 'function');
  assert.equal(typeof context.applyManualReservationInbound, 'function');
  assert.equal(typeof context.submitReservationInbound, 'function');

  assert.equal(typeof context.getCancellationContext, 'function');
  assert.equal(typeof context.executeCancellationScope, 'function');
  assert.equal(typeof context.executeCancellation, 'function');
  assert.equal(typeof context.선택_주문취소, 'function');
  assert.equal(typeof context.설정_열기, 'function');
  assert.equal(typeof context.설정_보기, 'function');
  assert.equal(typeof context.getOperatorSettings, 'function');
  assert.equal(typeof context.saveOperatorSettings, 'function');
});
