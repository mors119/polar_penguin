/**
 * ============================================================
 *  S9. 작업지시서 · PDF 생성
 * ============================================================
 *  피킹헤더와 피킹라인을 현장 작업용 지시서로 구성한다. 메뉴에서는 작업자에게
 *  미배정 슬롯을 배정해 인쇄하고, 통합 Input은 새 피킹 배치만 PDF로 보존한다.
 *
 *  PDF 파일 생성 또는 수동 인쇄 화면 준비가 성공한 시점이 fulfillment commit이다.
 *  이때 공용 finalizer가 예약재고를 소진하고 라인/헤더/주문을 자동 완료한다.
 *  이후 문제는 주문 취소 메뉴에서 주문번호 단위로 복원한다.
 * ============================================================
 */

/** 메뉴 진입점 — 출력 대화상자를 연다 */
function S9_1_작업지시서출력() {
  var html = HtmlService.createHtmlOutput(작업지시서_HTML_())
    .setWidth(900)
    .setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(html, '작업지시서 출력');
}

/**
 * 미배정 슬롯을 조회한다. (대화상자에서 호출)
 */
function S9_미배정슬롯조회() {
  var 헤더 = readTable_(ROLE.헤더);
  var H = {
    지시: col_(헤더, COL.피킹지시번호, true),
    주문: col_(헤더, COL.주문번호, true),
    슬롯: col_(헤더, COL.카트슬롯, true),
    품목수: col_(헤더, COL.품목수, true),
    총수량: col_(헤더, COL.총수량, true),
    담당: col_(헤더, COL.피킹담당자, true),
    상태: col_(헤더, COL.상태, true)
  };

  var 미배정 = 0, 배정 = {};
  헤더.rows.forEach(function (r) {
    var 상태 = toStr_(r[H.상태]);
    if (상태 === ENUM.헤더상태.완료 || 상태 === ENUM.헤더상태.취소) return;
    var 담당 = toStr_(r[H.담당]);
    if (담당) 배정[담당] = (배정[담당] || 0) + 1;
    else 미배정++;
  });

  return {
    미배정: 미배정,
    배정: Object.keys(배정).map(function (k) { return { 이름: k, 슬롯: 배정[k] }; })
  };
}

/**
 * 작업지시서를 만든다.
 * @param {String} 이름  작업자 이름
 * @param {Number} 개수  가져갈 슬롯 수 (0이면 이미 배정된 것만 재출력)
 * @param {{readOnly:Boolean=, 지시번호:String=}=} options PDF용 조회는 담당자/출력일시를 변경하지 않는다.
 * @return {Object} 작업자, 출력시각, 슬롯·품목·수량 정보 또는 오류 메시지
 * @sideEffect readOnly가 아니면 새 슬롯의 피킹담당자와 출력일시를 기록한다.
 */
function S9_지시서생성(이름, 개수, options) {
  return withLock_(function () {
    options = options || {};
    이름 = String(이름 || '').trim();
    if (!이름 && !options.readOnly) throw new Error('작업자 이름을 입력하세요.');
    if (!이름) 이름 = '자동 생성';
    개수 = Number(개수) || 0;

    var 헤더 = readTable_(ROLE.헤더);
    var H = {
      지시: col_(헤더, COL.피킹지시번호, true),
      주문: col_(헤더, COL.주문번호, true),
      슬롯: col_(헤더, COL.카트슬롯, true),
      품목수: col_(헤더, COL.품목수, true),
      총수량: col_(헤더, COL.총수량, true),
      담당: col_(헤더, COL.피킹담당자, true),
      상태: col_(헤더, COL.상태, true),
      출력일시: col_(헤더, COL.출력일시, false)
    };

    // ---------- 이미 이 사람에게 배정된 미완료 슬롯 ----------
    var 내슬롯 = [];
    헤더.rows.forEach(function (r, i) {
      var 상태 = toStr_(r[H.상태]);
      if (상태 === ENUM.헤더상태.취소) return;
      if (상태 === ENUM.헤더상태.완료 && !(options.readOnly && options.지시번호)) return;
      if (options.지시번호 && toStr_(r[H.지시]) !== String(options.지시번호)) return;
      if (options.readOnly || toStr_(r[H.담당]) === 이름) 내슬롯.push(i);
    });

    // ---------- 추가 배정 ----------
    var 신규배정 = 0;
    if (개수 > 0 && !options.readOnly) {
      var 후보 = [];
      헤더.rows.forEach(function (r, i) {
        var 상태 = toStr_(r[H.상태]);
        if (상태 === ENUM.헤더상태.완료 || 상태 === ENUM.헤더상태.취소) return;
        if (!isBlank_(r[H.담당])) return;
        후보.push({ idx: i, 슬롯: toNum_(r[H.슬롯]) });
      });
      후보.sort(function (a, b) { return a.슬롯 - b.슬롯; });   // 슬롯 번호 순

      후보.slice(0, 개수).forEach(function (c) {
        헤더.rows[c.idx][H.담당] = 이름;
        if (H.출력일시 >= 0) 헤더.rows[c.idx][H.출력일시] = new Date();
        내슬롯.push(c.idx);
        신규배정++;
      });

      if (신규배정) {
        writeColumn_(헤더.sheet, H.담당, 헤더.rows);
        if (H.출력일시 >= 0) writeColumn_(헤더.sheet, H.출력일시, 헤더.rows);
      }
    }

    if (!내슬롯.length) {
      return { 오류: '배정된 슬롯이 없습니다. 가져갈 슬롯 수를 지정하고 다시 시도하세요.' };
    }

    // ---------- 라인 조회 ----------
    var 주문 = readTable_(ROLE.주문);
    var O = {
      주문번호: col_(주문, COL.주문번호, true),
      품목별: col_(주문, COL.품목별주문번호, true)
    };
    var 품목별_주문 = {};
    주문.rows.forEach(function (r) {
      var k = toStr_(r[O.품목별]);
      if (k) 품목별_주문[k] = toStr_(r[O.주문번호]);
    });

    var 라인 = readTable_(ROLE.라인);
    var L = {
      순번: col_(라인, COL.순번, true),
      보관위치: col_(라인, COL.보관위치, true),
      상품코드: col_(라인, COL.상품코드, true),
      상품명: col_(라인, COL.상품명, true),
      옵션: col_(라인, COL.옵션, true),
      필요수량: col_(라인, COL.필요수량, true),
      품목별: col_(라인, COL.품목별주문번호, true),
      라인상태: col_(라인, COL.라인상태, true)
    };

    var 주문별라인 = {};
    라인.rows.forEach(function (r) {
      var no = 품목별_주문[toStr_(r[L.품목별])];
      if (!no) return;
      if (toStr_(r[L.라인상태]) === ENUM.라인상태.취소) return;
      (주문별라인[no] = 주문별라인[no] || []).push({
        순번: toNum_(r[L.순번]),
        위치: toStr_(r[L.보관위치]) || '(위치 미지정)',
        코드: toStr_(r[L.상품코드]),
        상품명: toStr_(r[L.상품명]),
        옵션: toStr_(r[L.옵션]) === '-' ? '' : toStr_(r[L.옵션]),
        수량: toNum_(r[L.필요수량])
      });
    });

    // ---------- 지시서 구성 ----------
    var 슬롯목록 = 내슬롯.map(function (i) {
      var r = 헤더.rows[i];
      var no = toStr_(r[H.주문]);
      var items = (주문별라인[no] || []).slice()
        .sort(function (a, b) { return a.순번 - b.순번; });
      return {
        슬롯: toNum_(r[H.슬롯]),
        주문번호: no,
        품목수: items.length,
        총수량: items.reduce(function (a, x) { return a + x.수량; }, 0),
        품목: items
      };
    }).filter(function (s) { return s.품목.length > 0; })
      .sort(function (a, b) { return a.슬롯 - b.슬롯; });

    if (!슬롯목록.length) {
      return { 오류: '출력할 품목이 없습니다. 배정된 슬롯이 모두 처리되었습니다.' };
    }

    var finalizations = [];
    if (!options.readOnly) {
      var instructions = {};
      내슬롯.forEach(function (idx) {
        var instructionNo = toStr_(헤더.rows[idx][H.지시]);
        (instructions[instructionNo] = instructions[instructionNo] || []).push(toStr_(헤더.rows[idx][H.주문]));
      });
      Object.keys(instructions).filter(String).forEach(function (instructionNo) {
        finalizations.push(finalizePickingAfterOutput_(instructionNo,
          { refresh: false, source: 'MANUAL_OUTPUT', orderNos: instructions[instructionNo] }));
      });
      try { D0_대시보드전체갱신(true); } catch (ignore) { }
    }
    writeOpLog_('S9_지시서생성', '성공', 이름 + ' / 슬롯 ' + 슬롯목록.length + '개 / 신규배정 ' + 신규배정);

    return {
      이름: 이름,
      출력시각: Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm'),
      신규배정: 신규배정,
      슬롯: 슬롯목록,
      총품목: 슬롯목록.reduce(function (a, s) { return a + s.품목수; }, 0),
      총수량: 슬롯목록.reduce(function (a, s) { return a + s.총수량; }, 0),
      출고확정: finalizations
    };
  });
}

/**
 * 신규 피킹 배치를 Output/YYYY-MM-DD에 PDF로 보존한다.
 * Output 바로 아래와 날짜 하위 폴더에서 같은 지시번호의 PDF를 먼저 찾아 재실행을 안전하게 한다.
 *
 * @param {string} 지시번호 PDF로 만들 피킹지시번호
 * @param {GoogleAppsScript.Drive.Folder=} outputRoot Output 폴더. 생략하면 설정값을 사용한다.
 * @return {Object} 생성 또는 기존 파일 재사용 결과
 * @sideEffect 신규인 경우에만 Drive에 PDF를 만들고 작업로그를 기록한다.
 */
function S9_피킹PDF생성(지시번호, outputRoot) {
  if (!지시번호) return { 생성: false, 사유: '지시번호 없음' };
  outputRoot = outputRoot || DriveApp.getFolderById(String(param_('Output폴더ID', '')));

  var fileName = String(지시번호) + '.pdf';
  if (S9_findPDF_(outputRoot, fileName)) {
    return { 생성: false, 재사용: true, 파일명: fileName,
      출고확정: finalizePickingAfterOutput_(지시번호, { refresh: false, source: 'PDF_REUSE' }) };
  }
  var dateName = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
  var dateFolder = getOrCreateSubFolder_(outputRoot, dateName);

  var data = S9_지시서생성('', 0, { readOnly: true, 지시번호: 지시번호 });
  if (data.오류) throw new Error(data.오류);
  var html = S9_PDFHTML_(data, 지시번호);
  var pdf = HtmlService.createHtmlOutput(html).getBlob().getAs(MimeType.PDF).setName(fileName);
  var file = dateFolder.createFile(pdf);
  var finalization = finalizePickingAfterOutput_(지시번호, { refresh: false, source: 'PDF_OUTPUT' });
  writeOpLog_('S9_피킹PDF생성', '성공', fileName);
  return { 생성: true, 파일ID: file.getId(), 파일명: fileName, 출고확정: finalization };
}

/** 출력오류 또는 누락 PDF를 같은 지시번호로 복구한다. 재고와 피킹행은 만들지 않는다. */
function S9_피킹PDF재시도(지시번호) {
  지시번호 = toStr_(지시번호);
  if (!지시번호) throw new Error('피킹지시번호를 입력하세요.');
  var orders = S9_ordersForInstruction_(지시번호);
  if (!orders.length) throw new Error('피킹지시번호를 찾을 수 없습니다: ' + 지시번호);
  try {
    var result = S9_피킹PDF생성(지시번호, inputFolder_('Output폴더ID'));
    D0_대시보드전체갱신(true);
    return { 메시지: (result.재사용 ? '기존 PDF를 확인했습니다.' : 'PDF를 다시 생성했습니다.') + ' ' + 지시번호 };
  } catch (e) {
    markPickingOutputState_(지시번호, ENUM.헤더상태.출력오류); throw e;
  }
}

function S9_ordersForInstruction_(instructionNo) {
  var table = readTable_(ROLE.헤더), cInstruction = col_(table, COL.피킹지시번호, true), cOrder = col_(table, COL.주문번호, true), found = {};
  table.rows.forEach(function (row) { if (toStr_(row[cInstruction]) === instructionNo) found[toStr_(row[cOrder])] = true; });
  return Object.keys(found).filter(String);
}

function S9_findPDF_(outputRoot, fileName) {
  if (outputRoot.getFilesByName(fileName).hasNext()) return true;
  var folders = outputRoot.getFolders();
  while (folders.hasNext()) {
    if (folders.next().getFilesByName(fileName).hasNext()) return true;
  }
  return false;
}

function S9_PDFHTML_(data, 지시번호) {
  var parts = ['<!doctype html><html><head><meta charset="utf-8"><style>',
    'body{font-family:sans-serif;color:#222}h1{font-size:22px}h2{background:#1F3864;color:#fff;padding:8px}',
    'table{width:100%;border-collapse:collapse;margin-bottom:18px}th,td{border:1px solid #aaa;padding:6px}',
    'th{background:#E8EDF3}.qty{text-align:center;font-weight:bold}.loc{font-weight:bold}',
    '</style></head><body><h1>피킹 작업지시서 ', S9_escapeHtml_(지시번호), '</h1>',
    '<p>생성 ', S9_escapeHtml_(data.출력시각), ' · 슬롯 ', data.슬롯.length, '개 · 총 ', data.총수량, '개</p>'];
  data.슬롯.forEach(function (slot) {
    parts.push('<h2>카트 슬롯 ', slot.슬롯, ' · ', S9_escapeHtml_(slot.주문번호), '</h2>',
      '<table><tr><th>No</th><th>보관위치</th><th>상품코드</th><th>상품명 / 옵션</th><th>수량</th></tr>');
    slot.품목.forEach(function (item) {
      parts.push('<tr><td>', item.순번, '</td><td class="loc">', S9_escapeHtml_(item.위치),
        '</td><td>', S9_escapeHtml_(item.코드), '</td><td>', S9_escapeHtml_(item.상품명),
        item.옵션 ? ' / ' + S9_escapeHtml_(item.옵션) : '', '</td><td class="qty">', item.수량,
        '</td></tr>');
    });
    parts.push('</table>');
  });
  parts.push('</body></html>');
  return parts.join('');
}

function S9_escapeHtml_(value) {
  return String(value == null ? '' : value).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

/** 대화상자 HTML */
function 작업지시서_HTML_() {
  return [
'<!DOCTYPE html><html><head><meta charset="utf-8"><style>',
'  * { box-sizing: border-box; }',
'  body { font-family: "Malgun Gothic", sans-serif; margin: 0; padding: 16px; color: #222; }',
'  h2 { margin: 0 0 4px; font-size: 18px; color: #1F3864; }',
'  .sub { color: #777; font-size: 12px; margin-bottom: 14px; }',
'  .panel { background: #F2F6FA; border: 1px solid #BFC9D4; border-radius: 6px; padding: 14px; margin-bottom: 14px; }',
'  label { font-size: 13px; font-weight: bold; margin-right: 6px; }',
'  input[type=text], input[type=number] { padding: 7px 9px; border: 1px solid #BFC9D4; border-radius: 4px; font-size: 14px; }',
'  input[type=text] { width: 150px; }',
'  input[type=number] { width: 70px; }',
'  button { padding: 8px 18px; border: 0; border-radius: 4px; font-size: 14px; cursor: pointer; }',
'  .primary { background: #2E5C8A; color: #fff; }',
'  .print { background: #1F3864; color: #fff; }',
'  button:disabled { background: #ccc; cursor: default; }',
'  .info { font-size: 12px; color: #555; margin-top: 10px; line-height: 1.6; }',
'  .err { color: #B03A2E; font-weight: bold; padding: 10px 0; }',
'  #out { margin-top: 10px; }',
'  .slot { border: 2px solid #1F3864; border-radius: 6px; margin-bottom: 16px; page-break-inside: avoid; }',
'  .slothead { background: #1F3864; color: #fff; padding: 8px 12px; font-weight: bold; font-size: 15px;',
'              display: flex; justify-content: space-between; }',
'  table { width: 100%; border-collapse: collapse; font-size: 13px; }',
'  th { background: #E8EDF3; padding: 6px; border: 1px solid #BFC9D4; font-size: 12px; }',
'  td { padding: 7px 6px; border: 1px solid #D6DCE4; }',
'  .loc { font-weight: bold; font-size: 15px; font-family: Consolas, monospace; }',
'  .qty { text-align: center; font-weight: bold; font-size: 16px; }',
'  .hdr { border-bottom: 3px solid #1F3864; padding-bottom: 8px; margin-bottom: 14px; }',
'  .hdr h1 { margin: 0; font-size: 21px; }',
'  .hdr .meta { font-size: 13px; color: #555; margin-top: 4px; }',
'  .note { font-size: 12px; color: #B03A2E; margin: 8px 0 14px; }',
'  @media print {',
'    .noprint { display: none !important; }',
'    body { padding: 0; }',
'    .slot { page-break-inside: avoid; }',
'  }',
'</style></head><body>',

'<div class="noprint">',
'  <h2>작업지시서 출력</h2>',
'  <div class="sub">이름을 넣고 가져갈 슬롯 수를 정하면, 미배정 슬롯이 순서대로 배정됩니다.</div>',
'  <div class="panel">',
'    <label>작업자 이름</label><input type="text" id="name" placeholder="예: 김서연">',
'    &nbsp;&nbsp;<label>가져갈 슬롯</label><input type="number" id="cnt" value="5" min="0" max="50">',
'    &nbsp;&nbsp;<button class="primary" id="go">불러오기</button>',
'    <div class="info" id="status">현황을 불러오는 중…</div>',
'  </div>',
'  <div class="panel">',
'    <label>PDF 조회/복구</label><input type="text" id="instruction" placeholder="예: PK-20260819-001">',
'    &nbsp;&nbsp;<button class="primary" id="retry">PDF 확인/재생성</button>',
'  </div>',
'  <div id="msg"></div>',
'</div>',
'<div id="out"></div>',

'<script>',
'function esc(s){return String(s==null?"":s).replace(/[&<>"]/g,function(c){',
'  return {"&":"&amp;","<":"&lt;",">":"&gt;","\\"":"&quot;"}[c];});}',
'',
'google.script.run.withSuccessHandler(function(d){',
'  var t = "미배정 슬롯 <b>" + d.미배정 + "개</b>";',
'  if (d.배정.length) {',
'    t += "<br>배정됨 — " + d.배정.map(function(w){return esc(w.이름)+" "+w.슬롯+"개";}).join(" · ");',
'  }',
'  document.getElementById("status").innerHTML = t;',
'}).S9_미배정슬롯조회();',
'',
'document.getElementById("go").onclick = function(){',
'  var nm = document.getElementById("name").value.trim();',
'  if(!nm){ alert("이름을 입력하세요."); return; }',
'  var n = parseInt(document.getElementById("cnt").value, 10) || 0;',
'  this.disabled = true; this.textContent = "불러오는 중…";',
'  var btn = this;',
'  document.getElementById("msg").innerHTML = "";',
'  google.script.run',
'    .withSuccessHandler(function(r){ btn.disabled=false; btn.textContent="불러오기"; render(r); })',
'    .withFailureHandler(function(e){ btn.disabled=false; btn.textContent="불러오기";',
'      document.getElementById("msg").innerHTML = "<div class=\\"err\\">"+esc(e.message)+"</div>"; })',
'    .S9_지시서생성(nm, n);',
'};',
'document.getElementById("retry").onclick = function(){',
'  var no=document.getElementById("instruction").value.trim(); if(!no){alert("피킹지시번호를 입력하세요.");return;}',
'  var btn=this; btn.disabled=true;',
'  google.script.run.withSuccessHandler(function(r){btn.disabled=false;alert(r.메시지);})',
'    .withFailureHandler(function(e){btn.disabled=false;alert(e.message);}).S9_피킹PDF재시도(no);',
'};',
'',
'function render(r){',
'  if(r && r.오류){ document.getElementById("msg").innerHTML = "<div class=\\"err\\">"+esc(r.오류)+"</div>"; return; }',
'  var h = "";',
'  h += "<div class=\\"noprint\\" style=\\"margin:10px 0\\">";',
'  h += "<button class=\\"print\\" onclick=\\"window.print()\\">🖨  인쇄하기</button>";',
'  if(r.신규배정) h += " &nbsp;<span style=\\"font-size:12px;color:#2E5C8A\\">새로 "+r.신규배정+"개 슬롯이 배정되었습니다.</span>";',
'  h += "</div>";',
'',
'  h += "<div class=\\"hdr\\"><h1>피킹 작업지시서</h1>";',
'  h += "<div class=\\"meta\\">작업자 <b>"+esc(r.이름)+"</b>";',
'  h += " &nbsp;·&nbsp; 출력 "+esc(r.출력시각);',
'  h += " &nbsp;·&nbsp; 슬롯 "+r.슬롯.length+"개 &nbsp;·&nbsp; 품목 "+r.총품목+"종 &nbsp;·&nbsp; 총 "+r.총수량+"개</div></div>";',
'  h += "<div class=\\"note\\">※ 출력 준비와 동시에 출고 처리됩니다. 문제가 있으면 주문 전체 취소로 재고를 복원하세요.</div>";',
'',
'  r.슬롯.forEach(function(s){',
'    h += "<div class=\\"slot\\"><div class=\\"slothead\\">";',
'    h += "<span>카트 슬롯 "+s.슬롯+"</span>";',
'    h += "<span style=\\"font-weight:normal;font-size:13px\\">"+esc(s.주문번호)+" &nbsp; 품목 "+s.품목수+" / 수량 "+s.총수량+"</span>";',
'    h += "</div><table><tr>";',
'    h += "<th style=\\"width:38px\\">No</th><th style=\\"width:100px\\">보관위치</th>";',
'    h += "<th style=\\"width:115px\\">상품코드</th><th>상품명 / 옵션</th>";',
'    h += "<th style=\\"width:52px\\">수량</th></tr>";',
'    s.품목.forEach(function(it){',
'      h += "<tr><td style=\\"text-align:center\\">"+it.순번+"</td>";',
'      h += "<td class=\\"loc\\">"+esc(it.위치)+"</td>";',
'      h += "<td style=\\"font-family:Consolas,monospace;font-size:12px\\">"+esc(it.코드)+"</td>";',
'      h += "<td>"+esc(it.상품명)+(it.옵션?" <span style=\\"color:#666\\">/ "+esc(it.옵션)+"</span>":"")+"</td>";',
'      h += "<td class=\\"qty\\">"+it.수량+"</td></tr>";',
'    });',
'    h += "</table></div>";',
'  });',
'',
'  document.getElementById("out").innerHTML = h;',
'}',
'</script></body></html>'
  ].join('\n');
}

/**
 * S9_2. 담당자 배정 해제 — 잘못 배정했을 때 되돌린다.
 */
function S9_2_배정해제(이름) {
  return withLock_(function () {
    if (!이름) {
      var ui = null;
      try { ui = SpreadsheetApp.getUi(); } catch (e) { }
      if (!ui) throw new Error('이름을 인자로 넘겨주세요.');
      var resp = ui.prompt('배정 해제',
        '배정을 해제할 작업자 이름을 입력하세요.\n아직 시작하지 않은(대기) 슬롯만 해제됩니다.',
        ui.ButtonSet.OK_CANCEL);
      if (resp.getSelectedButton() !== ui.Button.OK) return;
      이름 = resp.getResponseText().trim();
      if (!이름) return;
    }

    var 헤더 = readTable_(ROLE.헤더);
    var H = {
      담당: col_(헤더, COL.피킹담당자, true),
      상태: col_(헤더, COL.상태, true)
    };

    var 해제 = 0;
    헤더.rows.forEach(function (r) {
      if (toStr_(r[H.담당]) !== 이름) return;
      if (toStr_(r[H.상태]) !== ENUM.헤더상태.대기) return;   // 진행 중인 건 건드리지 않는다
      r[H.담당] = '';
      해제++;
    });

    if (해제) writeColumn_(헤더.sheet, H.담당, 헤더.rows);

    var msg = 이름 + ' 님의 대기 슬롯 ' + 해제 + '개를 해제했습니다.' +
      (해제 === 0 ? '\n(진행 중이거나 완료된 슬롯은 해제되지 않습니다)' : '');
    alert_(msg);
    writeOpLog_('S9_2_배정해제', '성공', 이름 + ' / ' + 해제 + '개');
    return 해제;
  });
}
