/**
 * ============================================================
 *  D. 대시보드  (v4.0)
 * ============================================================
 *  세 곳에 각각 대시보드 탭을 만든다. 보는 사람이 다르기 때문이다.
 *
 *   D1. 주문현황  → 주문(완료) 파일   · 관리자용
 *       주문 파이프라인 + 피킹 진행 속도 + 갱신 권고
 *   D2. 재고현황  → 상품마스터 파일   · 관리자용
 *       가용/예약/부족/예약대기 재고
 *   D3. 피킹현황  → 피킹(헤더) 파일   · 현장용
 *       슬롯별 진행, 담당자별 실적, 취소 알림
 * ============================================================
 */

/* ============================================================
 *  D1. 주문 현황  (주문(완료) 파일)
 * ============================================================ */

function D1_주문현황갱신() {
  var ss = openSS_(ROLE.주문);
  var sh = ensureDashSheet_(ss, '📊 주문현황');
  var d = collectOrderStatus_();

  sh.clear(); sh.clearFormats();
  if (sh.getMaxColumns() < 8) sh.insertColumnsAfter(sh.getMaxColumns(), 8 - sh.getMaxColumns());

  var r = 1;

  // ---- 제목 ----
  sh.getRange(r, 1, 1, 8).merge().setValue('📊  주문 현황')
    .setFontSize(17).setFontWeight('bold').setFontColor('FFFFFF')
    .setBackground(DASHCOLOR.제목).setVerticalAlignment('middle');
  sh.setRowHeight(r, 38); r++;

  sh.getRange(r, 1, 1, 8).merge()
    .setValue('갱신 ' + Utilities.formatDate(new Date(), tz_(), 'MM-dd HH:mm'))
    .setFontSize(10).setFontColor('595959');
  r += 2;

  // ---- 파이프라인 카드 ----
  var 카드 = [
    ['접수', d.접수, DASHCOLOR.대기],
    ['확정', d.확정, DASHCOLOR.진행],
    ['예약대기', d.예약대기, DASHCOLOR.경고],
    ['출고완료', d.출고완료, DASHCOLOR.완료]
  ];
  카드.forEach(function (c, i) {
    var col = 1 + i * 2;
    sh.getRange(r, col, 1, 2).merge().setValue(c[0])
      .setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('center').setBackground(c[2]);
    sh.getRange(r + 1, col, 1, 2).merge().setValue(c[1])
      .setFontSize(22).setFontWeight('bold')
      .setHorizontalAlignment('center').setBackground(DASHCOLOR.카드)
      .setVerticalAlignment('middle');
  });
  sh.setRowHeight(r, 20); sh.setRowHeight(r + 1, 42);
  r += 3;

  sh.getRange(r, 1, 1, 8).merge()
    .setValue('취소 ' + d.취소 + '건   ·   전체 주문 ' + d.전체 + '건')
    .setFontSize(10).setFontColor('595959');
  r += 2;

  // ---- 관리자 권고 ----
  sh.getRange(r, 1).setValue('오늘의 권고').setFontWeight('bold').setFontSize(12).setFontColor(DASHCOLOR.제목);
  r++;

  sh.getRange(r, 1, 1, 8).merge().setValue(d.권고.제목)
    .setFontSize(13).setFontWeight('bold')
    .setBackground(d.권고.색).setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  sh.setRowHeight(r, 34); r++;

  d.권고.내용.forEach(function (line) {
    sh.getRange(r, 1, 1, 8).merge().setValue('   ' + line).setFontSize(11);
    r++;
  });
  r++;

  // ---- 피킹 진행 상황 ----
  sh.getRange(r, 1).setValue('피킹 진행').setFontWeight('bold').setFontSize(12).setFontColor(DASHCOLOR.제목);
  r++;

  var 진행표 = [
    ['전체 라인', d.피킹.전체라인 + ' 품목'],
    ['처리 완료', d.피킹.처리라인 + ' 품목  (' + d.피킹.진행률 + '%)'],
    ['남은 작업', d.피킹.남은라인 + ' 품목'],
    ['처리 속도', d.피킹.속도 ? d.피킹.속도 + ' 품목/시간' : '측정 불가 (데이터 부족)'],
    ['예상 소요', d.피킹.예상분 !== null ? 표시시간_(d.피킹.예상분) : '—'],
    ['미착수 슬롯', d.피킹.대기슬롯 + ' 개'],
    ['진행 중 슬롯', d.피킹.진행슬롯 + ' 개']
  ];
  진행표.forEach(function (row) {
    sh.getRange(r, 1, 1, 2).merge().setValue(row[0]).setFontWeight('bold').setFontSize(10)
      .setBackground(DASHCOLOR.헤더);
    sh.getRange(r, 3, 1, 6).merge().setValue(row[1]).setFontSize(10);
    r++;
  });
  r++;

  sh.getRange(r, 1, 1, 8).merge()
    .setValue(bar_(d.피킹.진행률, 24) + '   ' + d.피킹.진행률 + '%')
    .setFontFamily('Consolas').setFontSize(11);
  r += 2;

  // ---- 예약대기 요약 ----
  sh.getRange(r, 1).setValue('예약대기 상위 품목').setFontWeight('bold').setFontSize(12).setFontColor('BF6000');
  r++;

  if (d.예약목록.length) {
    sh.getRange(r, 1, 1, 8).setValues([['상품코드', '상품명', '', '대기수량', '현재고', '부족', '상태', '']])
      .setFontWeight('bold').setFontSize(10).setBackground(DASHCOLOR.경고);
    r++;
    var pv = d.예약목록.slice(0, 10).map(function (p) {
      return [p.코드, p.상품명.substring(0, 30), '', p.수량, p.현재고,
              Math.max(p.부족, 0), p.부족 <= 0 ? '출고가능' : '입고대기', ''];
    });
    sh.getRange(r, 1, pv.length, 8).setValues(pv).setFontSize(10)
      .setBorder(true, true, true, true, true, true, DASHCOLOR.선, SpreadsheetApp.BorderStyle.SOLID);
    d.예약목록.slice(0, 10).forEach(function (p, i) {
      sh.getRange(r + i, 7).setBackground(p.부족 <= 0 ? DASHCOLOR.완료 : DASHCOLOR.위험);
    });
    r += pv.length;
  } else {
    sh.getRange(r, 1, 1, 8).merge().setValue('없음').setFontColor('808080').setFontSize(10);
    r++;
  }

  var 폭 = [110, 120, 100, 90, 80, 70, 90, 70];
  폭.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(2);
  return d;
}

/**
 * 주문 파이프라인 + 피킹 속도 + 권고 산출
 */
function collectOrderStatus_() {
  var out = {
    접수: 0, 확정: 0, 예약대기: 0, 취소: 0, 출고완료: 0, 전체: 0,
    피킹: { 전체라인: 0, 처리라인: 0, 남은라인: 0, 진행률: 0, 속도: 0, 예상분: null, 대기슬롯: 0, 진행슬롯: 0 },
    예약목록: [],
    권고: { 제목: '', 내용: [], 색: DASHCOLOR.카드 }
  };

  // ---------- 주문 상태 집계 ----------
  var 주문 = readTable_(ROLE.주문);
  var O = {
    주문번호: col_(주문, COL.주문번호, true),
    상품코드: col_(주문, COL.상품품목코드, true),
    수량: col_(주문, COL.수량, true),
    출고완료: col_(주문, COL.출고완료, true),
    주문상태: col_(주문, COL.주문상태, true)
  };

  var 본 = {}, 예약집계 = {};
  주문.rows.forEach(function (r) {
    var no = toStr_(r[O.주문번호]);
    if (!no) return;

    var 상태 = toStr_(r[O.주문상태]) || ENUM.주문상태.접수;
    var 출고 = toNum_(r[O.출고완료]) === 1;

    if (!본[no]) {
      본[no] = true;
      out.전체++;
      if (출고) out.출고완료++;
      else if (상태 === ENUM.주문상태.접수) out.접수++;
      else if (상태 === ENUM.주문상태.확정) out.확정++;
      else if (상태 === ENUM.주문상태.예약대기) out.예약대기++;
      else if (상태 === ENUM.주문상태.취소) out.취소++;
    }

    if (상태 === ENUM.주문상태.예약대기) {
      var c = toStr_(r[O.상품코드]);
      if (c) 예약집계[c] = (예약집계[c] || 0) + toNum_(r[O.수량]);
    }
  });

  // ---------- 예약대기 품목 ----------
  var 마스터 = readTable_(ROLE.마스터);
  var M = {
    코드: col_(마스터, COL.상품품목코드, true),
    상품명: col_(마스터, COL.상품명, true),
    가용: col_(마스터, COL.가용재고, true)
  };
  var 정보 = {};
  마스터.rows.forEach(function (r) {
    var c = toStr_(r[M.코드]);
    if (c) 정보[c] = { 상품명: toStr_(r[M.상품명]), 가용: toNum_(r[M.가용]) };
  });

  Object.keys(예약집계).forEach(function (c) {
    var info = 정보[c] || { 상품명: '(미등록)', 가용: 0 };
    out.예약목록.push({
      코드: c, 상품명: info.상품명, 수량: 예약집계[c],
      현재고: info.가용, 부족: 예약집계[c] - info.가용
    });
  });
  out.예약목록.sort(function (a, b) { return b.수량 - a.수량; });

  // ---------- 피킹 진행 ----------
  var 라인 = readTable_(ROLE.라인);
  var L = {
    라인상태: col_(라인, COL.라인상태, true),
    처리일시: col_(라인, COL.처리일시, true)
  };

  var 처리시각 = [];
  라인.rows.forEach(function (r) {
    var s = toStr_(r[L.라인상태]);
    if (!s) return;
    out.피킹.전체라인++;
    if (s !== ENUM.라인상태.미처리) {
      out.피킹.처리라인++;
      var t = r[L.처리일시];
      if (t instanceof Date) 처리시각.push(t.getTime());
    }
  });
  out.피킹.남은라인 = out.피킹.전체라인 - out.피킹.처리라인;
  out.피킹.진행률 = out.피킹.전체라인
    ? Math.round(out.피킹.처리라인 / out.피킹.전체라인 * 100) : 0;

  // 처리 속도: 최근 처리 이력의 시간 범위로 계산
  if (처리시각.length >= 3) {
    처리시각.sort(function (a, b) { return a - b; });
    var 최근 = 처리시각.slice(-40);              // 최근 40건
    var 경과시간 = (최근[최근.length - 1] - 최근[0]) / 3600000;   // 시간
    if (경과시간 > 0.05) {
      out.피킹.속도 = Math.round((최근.length - 1) / 경과시간);
      if (out.피킹.속도 > 0) {
        out.피킹.예상분 = Math.round(out.피킹.남은라인 / out.피킹.속도 * 60);
      }
    }
  }

  // ---------- 슬롯 상태 ----------
  var 헤더 = readTable_(ROLE.헤더);
  var H = { 상태: col_(헤더, COL.상태, true) };
  헤더.rows.forEach(function (r) {
    var s = toStr_(r[H.상태]);
    if (s === ENUM.헤더상태.대기) out.피킹.대기슬롯++;
    if (s === ENUM.헤더상태.진행) out.피킹.진행슬롯++;
  });

  // ---------- 권고 산출 ----------
  out.권고 = 권고산출_(out);
  return out;
}

/**
 * 진행 상황을 보고 관리자에게 다음 행동을 권한다.
 *
 *  판단 기준
 *   · 남은 작업이 적고 접수 주문이 쌓여 있다 → 주문 확정을 더 돌려라
 *   · 확정된 주문이 있는데 지시가 안 나갔다 → 피킹지시를 생성해라
 *   · 남은 작업이 많다 → 현재 물량부터 소화해라
 *   · 예약 재고가 들어왔다 → 확정을 다시 돌려 예약을 풀어라
 */
function 권고산출_(d) {
  var 임계 = Number(param_('추가투입임계(분)', 45));
  var 남은 = d.피킹.남은라인;
  var 예상 = d.피킹.예상분;

  // 1) 예약 해소 가능
  var 출고가능 = d.예약목록.filter(function (p) { return p.부족 <= 0; });
  if (출고가능.length) {
    return {
      제목: '⚡  예약 주문을 풀 수 있습니다',
      색: DASHCOLOR.완료,
      내용: [
        '재고가 확보된 예약 품목이 ' + 출고가능.length + '종 있습니다.',
        '「S3. 주문 확정」을 실행하면 예약대기 주문이 자동으로 확정됩니다.',
        '대상: ' + 출고가능.slice(0, 3).map(function (p) { return p.코드; }).join(', ') +
          (출고가능.length > 3 ? ' 외 ' + (출고가능.length - 3) + '종' : '')
      ]
    };
  }

  // 2) 확정됐는데 지시 미발행
  if (d.확정 > 0) {
    return {
      제목: '📋  피킹지시를 발행하세요',
      색: DASHCOLOR.진행,
      내용: [
        '확정된 주문 ' + d.확정 + '건이 지시 발행을 기다리고 있습니다.',
        '「S4. 피킹지시 생성」을 실행하면 작업 목록이 만들어집니다.'
      ]
    };
  }

  // 3) 작업 여유 있음 + 접수 주문 있음
  if (d.접수 > 0 && (남은 === 0 || (예상 !== null && 예상 <= 임계))) {
    return {
      제목: '✅  주문을 더 투입해도 됩니다',
      색: DASHCOLOR.좋음,
      내용: [
        남은 === 0
          ? '진행 중인 피킹 작업이 없습니다.'
          : '남은 작업 ' + 남은 + '품목, 예상 ' + 표시시간_(예상) + ' 내 완료 예정입니다.',
        '접수 상태 주문이 ' + d.접수 + '건 대기 중입니다.',
        '「S3. 주문 확정」 → 「S4. 피킹지시 생성」 순으로 추가 투입하세요.'
      ]
    };
  }

  // 4) 작업 과부하
  if (예상 !== null && 예상 > 임계 * 2) {
    return {
      제목: '⏳  현재 물량부터 소화하세요',
      색: DASHCOLOR.경고,
      내용: [
        '남은 작업 ' + 남은 + '품목, 예상 ' + 표시시간_(예상) + ' 소요됩니다.',
        '지금 주문을 더 넣으면 카트와 동선이 엉킬 수 있습니다.',
        d.접수 > 0 ? '대기 중인 접수 주문 ' + d.접수 + '건은 다음 차수에 투입하세요.' : ''
      ].filter(String)
    };
  }

  // 5) 진행 중
  if (남은 > 0) {
    return {
      제목: '🔄  피킹 진행 중',
      색: DASHCOLOR.진행,
      내용: [
        '남은 작업 ' + 남은 + '품목' + (예상 !== null ? ', 예상 ' + 표시시간_(예상) : '') + '.',
        '진행 중 슬롯 ' + d.피킹.진행슬롯 + '개 · 미착수 슬롯 ' + d.피킹.대기슬롯 + '개',
        d.접수 > 0 ? '접수 주문 ' + d.접수 + '건은 작업 여유가 생기면 투입하세요.' : ''
      ].filter(String)
    };
  }

  // 6) 할 일 없음
  if (d.접수 === 0 && d.확정 === 0) {
    return {
      제목: '💤  대기 중',
      색: DASHCOLOR.카드,
      내용: [
        '처리할 주문이 없습니다.',
        '쇼핑몰 주문 CSV를 올린 뒤 「S2. 주문 CSV 취입」을 실행하세요.',
        d.예약대기 > 0 ? '예약대기 주문 ' + d.예약대기 + '건은 입고를 기다리는 중입니다.' : ''
      ].filter(String)
    };
  }

  return { 제목: '🔄  진행 중', 색: DASHCOLOR.진행, 내용: ['현재 특별한 조치가 필요하지 않습니다.'] };
}

function 표시시간_(분) {
  if (분 === null || 분 === undefined) return '—';
  if (분 < 60) return 분 + '분';
  var h = Math.floor(분 / 60), m = 분 % 60;
  return h + '시간' + (m ? ' ' + m + '분' : '');
}

/* ============================================================
 *  D2. 재고 현황  (상품마스터 파일)
 * ============================================================ */

function D2_재고현황갱신() {
  var ss = openSS_(ROLE.마스터);
  var sh = ensureDashSheet_(ss, '📊 재고현황');
  var d = collectStockStatus_();

  sh.clear(); sh.clearFormats();
  if (sh.getMaxColumns() < 8) sh.insertColumnsAfter(sh.getMaxColumns(), 8 - sh.getMaxColumns());

  var r = 1;

  sh.getRange(r, 1, 1, 8).merge().setValue('📊  재고 현황')
    .setFontSize(17).setFontWeight('bold').setFontColor('FFFFFF')
    .setBackground(DASHCOLOR.제목).setVerticalAlignment('middle');
  sh.setRowHeight(r, 38); r++;

  sh.getRange(r, 1, 1, 8).merge()
    .setValue('갱신 ' + Utilities.formatDate(new Date(), tz_(), 'MM-dd HH:mm') +
              (d.최종동기화 ? '     ·     카페24 동기화 ' + d.최종동기화 : ''))
    .setFontSize(10).setFontColor('595959');
  r += 2;

  // ---- KPI ----
  var 카드 = [
    ['등록 상품', d.상품수 + '종', DASHCOLOR.헤더],
    ['가용 재고', d.총가용, DASHCOLOR.완료],
    ['예약 재고', d.총예약, DASHCOLOR.진행],
    ['품절', d.품절 + '종', DASHCOLOR.예외]
  ];
  카드.forEach(function (c, i) {
    var col = 1 + i * 2;
    sh.getRange(r, col, 1, 2).merge().setValue(c[0])
      .setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center').setBackground(c[2]);
    sh.getRange(r + 1, col, 1, 2).merge().setValue(c[1])
      .setFontSize(20).setFontWeight('bold').setHorizontalAlignment('center')
      .setBackground(DASHCOLOR.카드).setVerticalAlignment('middle');
  });
  sh.setRowHeight(r, 20); sh.setRowHeight(r + 1, 42);
  r += 3;

  sh.getRange(r, 1, 1, 8).merge()
    .setValue('총 보유 ' + (d.총가용 + d.총예약 + d.총불량) + '개  ·  불량 ' + d.총불량 +
              '개  ·  예약상품 ' + d.예약상품수 + '종  ·  보관위치 미지정 ' + d.위치없음 + '종')
    .setFontSize(10).setFontColor('595959');
  r += 2;

  // ---- 조치 필요 ----
  if (d.위치없음 > 0) {
    sh.getRange(r, 1, 1, 8).merge()
      .setValue('⚠  보관위치가 없는 상품이 ' + d.위치없음 + '종 있습니다. 피킹 시 위치가 표시되지 않습니다.')
      .setFontSize(11).setFontWeight('bold').setBackground(DASHCOLOR.경고);
    r += 2;
  }

  // ---- 재고 부족 ----
  var 임계 = Number(param_('재고경고임계치', 3));
  sh.getRange(r, 1).setValue('재고 부족  (가용 ' + 임계 + '개 이하 · 예약상품 제외)')
    .setFontWeight('bold').setFontSize(12).setFontColor('BF6000');
  r++;

  if (d.부족.length) {
    sh.getRange(r, 1, 1, 8).setValues([['상품코드', '상품명', '옵션', '위치', '가용', '예약', '총보유', '']])
      .setFontWeight('bold').setFontSize(10).setBackground(DASHCOLOR.경고);
    r++;
    var fv = d.부족.slice(0, 25).map(function (p) {
      return [p.코드, p.상품명.substring(0, 28), p.옵션, p.위치, p.가용, p.예약, p.가용 + p.예약 + p.불량, ''];
    });
    sh.getRange(r, 1, fv.length, 8).setValues(fv).setFontSize(10)
      .setBorder(true, true, true, true, true, true, DASHCOLOR.선, SpreadsheetApp.BorderStyle.SOLID);
    d.부족.slice(0, 25).forEach(function (p, i) {
      if (p.가용 <= 0) sh.getRange(r + i, 5).setBackground(DASHCOLOR.위험).setFontWeight('bold');
    });
    r += fv.length;
    if (d.부족.length > 25) {
      sh.getRange(r, 1, 1, 8).merge().setValue('… 외 ' + (d.부족.length - 25) + '종').setFontColor('808080');
      r++;
    }
  } else {
    sh.getRange(r, 1, 1, 8).merge().setValue('없음').setFontColor('808080').setFontSize(10);
    r++;
  }
  r += 2;

  // ---- 예약 재고 보유 ----
  sh.getRange(r, 1).setValue('예약 재고 보유  (확정됐지만 아직 출고 전)')
    .setFontWeight('bold').setFontSize(12).setFontColor(DASHCOLOR.제목);
  r++;

  if (d.예약보유.length) {
    sh.getRange(r, 1, 1, 8).setValues([['상품코드', '상품명', '옵션', '위치', '가용', '예약', '', '']])
      .setFontWeight('bold').setFontSize(10).setBackground(DASHCOLOR.헤더);
    r++;
    var rv = d.예약보유.slice(0, 20).map(function (p) {
      return [p.코드, p.상품명.substring(0, 28), p.옵션, p.위치, p.가용, p.예약, '', ''];
    });
    sh.getRange(r, 1, rv.length, 8).setValues(rv).setFontSize(10)
      .setBorder(true, true, true, true, true, true, DASHCOLOR.선, SpreadsheetApp.BorderStyle.SOLID);
    r += rv.length;
  } else {
    sh.getRange(r, 1, 1, 8).merge().setValue('없음').setFontColor('808080').setFontSize(10);
    r++;
  }

  var 폭 = [110, 230, 90, 85, 60, 60, 70, 60];
  폭.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(2);
  return d;
}

function collectStockStatus_() {
  var out = {
    상품수: 0, 총가용: 0, 총예약: 0, 총불량: 0, 품절: 0,
    예약상품수: 0, 위치없음: 0, 부족: [], 예약보유: [], 최종동기화: ''
  };

  var 마스터 = readTable_(ROLE.마스터);
  var M = {
    코드: col_(마스터, COL.상품품목코드, true),
    상품명: col_(마스터, COL.상품명, true),
    옵션: col_(마스터, COL.옵션명, false),
    위치: col_(마스터, COL.기본보관위치, true),
    가용: col_(마스터, COL.가용재고, true),
    예약: col_(마스터, COL.예약재고, false),
    불량: col_(마스터, COL.불량재고, false),
    상품상태: col_(마스터, COL.상품상태, false),
    예약상품: col_(마스터, COL.예약상품, false),
    동기화: col_(마스터, COL.최종동기화, false)
  };

  var 임계 = Number(param_('재고경고임계치', 3));
  var 최신 = null;

  마스터.rows.forEach(function (r) {
    var code = toStr_(r[M.코드]);
    if (!code) return;
    if (M.상품상태 >= 0 && toStr_(r[M.상품상태]) === '사용중지') return;

    out.상품수++;
    var 가용 = toNum_(r[M.가용]);
    var 예약 = M.예약 >= 0 ? toNum_(r[M.예약]) : 0;
    var 불량 = M.불량 >= 0 ? toNum_(r[M.불량]) : 0;
    var 예약상품 = M.예약상품 >= 0 && toStr_(r[M.예약상품]).toUpperCase() === 'Y';

    out.총가용 += 가용; out.총예약 += 예약; out.총불량 += 불량;
    if (예약상품) out.예약상품수++;
    if (가용 <= 0 && !예약상품) out.품절++;
    if (isBlank_(r[M.위치])) out.위치없음++;

    if (M.동기화 >= 0 && r[M.동기화] instanceof Date) {
      if (!최신 || r[M.동기화] > 최신) 최신 = r[M.동기화];
    }

    var item = {
      코드: code,
      상품명: toStr_(r[M.상품명]),
      옵션: M.옵션 >= 0 ? (toStr_(r[M.옵션]) === '-' ? '' : toStr_(r[M.옵션])) : '',
      위치: toStr_(r[M.위치]),
      가용: 가용, 예약: 예약, 불량: 불량
    };

    if (가용 <= 임계 && !예약상품) out.부족.push(item);
    if (예약 > 0) out.예약보유.push(item);
  });

  out.부족.sort(function (a, b) { return a.가용 - b.가용; });
  out.예약보유.sort(function (a, b) { return b.예약 - a.예약; });
  if (최신) out.최종동기화 = Utilities.formatDate(최신, tz_(), 'MM-dd HH:mm');

  return out;
}

/* ============================================================
 *  D3. 피킹 현황  (피킹(헤더) 파일)
 * ============================================================ */

function D3_피킹현황갱신() {
  var ss = openSS_(ROLE.헤더);
  var sh = ensureDashSheet_(ss, '📊 피킹현황');
  var d = collectPickingStatus_();

  sh.clear(); sh.clearFormats();
  if (sh.getMaxColumns() < 8) sh.insertColumnsAfter(sh.getMaxColumns(), 8 - sh.getMaxColumns());

  var r = 1;

  sh.getRange(r, 1, 1, 8).merge().setValue('📊  피킹 현황')
    .setFontSize(17).setFontWeight('bold').setFontColor('FFFFFF')
    .setBackground(DASHCOLOR.제목).setVerticalAlignment('middle');
  sh.setRowHeight(r, 38); r++;

  sh.getRange(r, 1, 1, 8).merge()
    .setValue('갱신 ' + Utilities.formatDate(new Date(), tz_(), 'MM-dd HH:mm'))
    .setFontSize(10).setFontColor('595959');
  r += 2;

  // ---- KPI ----
  var 카드 = [
    ['대기', d.kpi.대기, DASHCOLOR.대기],
    ['진행 중', d.kpi.진행, DASHCOLOR.진행],
    ['완료', d.kpi.완료, DASHCOLOR.완료],
    ['취소', d.kpi.예외, DASHCOLOR.예외]
  ];
  카드.forEach(function (c, i) {
    var col = 1 + i * 2;
    sh.getRange(r, col, 1, 2).merge().setValue(c[0])
      .setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center').setBackground(c[2]);
    sh.getRange(r + 1, col, 1, 2).merge().setValue(c[1])
      .setFontSize(22).setFontWeight('bold').setHorizontalAlignment('center')
      .setBackground(DASHCOLOR.카드).setVerticalAlignment('middle');
  });
  sh.setRowHeight(r, 20); sh.setRowHeight(r + 1, 42);
  r += 3;

  sh.getRange(r, 1, 1, 2).merge().setValue('전체 진행률').setFontWeight('bold').setFontSize(11);
  sh.getRange(r, 3, 1, 6).merge()
    .setValue(bar_(d.진행률, 20) + '   ' + d.진행률 + '%   (' + d.처리라인 + ' / ' + d.전체라인 + ' 품목)')
    .setFontFamily('Consolas').setFontSize(11);
  r += 2;

  // ---- 담당자별 실적 ----
  sh.getRange(r, 1).setValue('담당자별 실적').setFontWeight('bold').setFontSize(12).setFontColor(DASHCOLOR.제목);
  r++;

  if (d.담당자.length) {
    sh.getRange(r, 1, 1, 8).setValues([['담당자', '맡은 슬롯', '완료', '진행', '처리 품목', '진행률', '', '']])
      .setFontWeight('bold').setFontSize(10).setBackground(DASHCOLOR.헤더);
    r++;
    var wv = d.담당자.map(function (w) {
      return [w.이름, w.슬롯수, w.완료, w.진행, w.처리라인,
              bar_(w.진행률, 10) + ' ' + w.진행률 + '%', '', ''];
    });
    sh.getRange(r, 1, wv.length, 8).setValues(wv).setFontSize(10)
      .setBorder(true, true, true, true, true, true, DASHCOLOR.선, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, 6, wv.length, 1).setFontFamily('Consolas');
    r += wv.length;
  } else {
    sh.getRange(r, 1, 1, 8).merge()
      .setValue('담당자가 지정된 슬롯이 없습니다. 피킹(헤더)의 피킹담당자 칸에 이름을 적으면 집계됩니다.')
      .setFontColor('808080').setFontSize(10);
    r++;
  }
  r += 2;

  // ---- 슬롯별 현황 ----
  sh.getRange(r, 1).setValue('슬롯별 현황  (미완료만)').setFontWeight('bold').setFontSize(12).setFontColor(DASHCOLOR.제목);
  r++;

  sh.getRange(r, 1, 1, 8).setValues([['슬롯', '주문번호', '담당자', '품목', '진행', '상태', '비고', '']])
    .setFontWeight('bold').setFontSize(10).setBackground(DASHCOLOR.헤더);
  r++;

  var 미완 = d.슬롯.filter(function (s) { return s.상태 !== ENUM.헤더상태.완료; });
  if (미완.length) {
    var sv = 미완.slice(0, 30).map(function (s) {
      return [s.슬롯, s.주문번호, s.담당자 || '—', s.품목수,
              bar_(s.진행률, 8) + ' ' + s.진행률 + '%', s.상태, s.비고, ''];
    });
    sh.getRange(r, 1, sv.length, 8).setValues(sv).setFontSize(10)
      .setBorder(true, true, true, true, true, true, DASHCOLOR.선, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, 5, sv.length, 1).setFontFamily('Consolas');
    미완.slice(0, 30).forEach(function (s, i) {
      var bg = DASHCOLOR[s.상태];
      if (bg) sh.getRange(r + i, 6).setBackground(bg);
      if (s.상태 === ENUM.헤더상태.예외) sh.getRange(r + i, 1, 1, 8).setFontColor('9C0006');
    });
    r += sv.length;
  } else {
    sh.getRange(r, 1, 1, 8).merge().setValue('미완료 슬롯이 없습니다.').setFontColor('808080').setFontSize(10);
    r++;
  }
  r += 2;

  // ---- 취소 알림 ----
  sh.getRange(r, 1).setValue('⚠  취소된 주문 — 카트에서 빼서 원위치')
    .setFontWeight('bold').setFontSize(12).setFontColor('9C0006');
  r++;

  var 취소 = d.슬롯.filter(function (s) { return s.상태 === ENUM.헤더상태.예외; });
  if (취소.length) {
    var cv = 취소.map(function (s) {
      return [s.슬롯, s.주문번호, s.담당자 || '—', s.비고, '', '', '', ''];
    });
    sh.getRange(r, 1, cv.length, 8).setValues(cv).setFontSize(10)
      .setBackground(DASHCOLOR.경고)
      .setBorder(true, true, true, true, true, true, DASHCOLOR.선, SpreadsheetApp.BorderStyle.SOLID);
    r += cv.length;
  } else {
    sh.getRange(r, 1, 1, 8).merge().setValue('없음').setFontColor('808080').setFontSize(10);
  }

  var 폭 = [70, 150, 90, 60, 130, 70, 200, 60];
  폭.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(2);
  return d;
}

function collectPickingStatus_() {
  var out = {
    kpi: { 대기: 0, 진행: 0, 완료: 0, 예외: 0 },
    슬롯: [], 담당자: [], 전체라인: 0, 처리라인: 0, 진행률: 0
  };

  var 헤더 = readTable_(ROLE.헤더);
  var H = {
    지시: col_(헤더, COL.피킹지시번호, true),
    주문: col_(헤더, COL.주문번호, true),
    슬롯: col_(헤더, COL.카트슬롯, true),
    품목수: col_(헤더, COL.품목수, true),
    담당: col_(헤더, COL.피킹담당자, true),
    상태: col_(헤더, COL.상태, true)
  };

  var 라인 = readTable_(ROLE.라인);
  var L = {
    품목별: col_(라인, COL.품목별주문번호, true),
    라인상태: col_(라인, COL.라인상태, true),
    예외사유: col_(라인, COL.예외사유, true)
  };

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

  var 집계 = {};
  라인.rows.forEach(function (r) {
    var no = 품목별_주문[toStr_(r[L.품목별])];
    if (!no) return;
    if (!집계[no]) 집계[no] = { 전체: 0, 처리: 0, 예외사유: '' };
    집계[no].전체++;
    var s = toStr_(r[L.라인상태]);
    if (s && s !== ENUM.라인상태.미처리) 집계[no].처리++;
    var e = toStr_(r[L.예외사유]);
    if (e && !집계[no].예외사유) 집계[no].예외사유 = e;
  });

  var 담당맵 = {};

  헤더.rows.forEach(function (r) {
    var no = toStr_(r[H.주문]);
    if (!no) return;
    var 상태 = toStr_(r[H.상태]) || ENUM.헤더상태.대기;
    var 담당 = toStr_(r[H.담당]);
    var a = 집계[no] || { 전체: toNum_(r[H.품목수]), 처리: 0, 예외사유: '' };
    var pct = a.전체 ? Math.round(a.처리 / a.전체 * 100) : 0;

    if (out.kpi[상태] !== undefined) out.kpi[상태]++;
    out.전체라인 += a.전체;
    out.처리라인 += a.처리;

    out.슬롯.push({
      슬롯: toNum_(r[H.슬롯]), 주문번호: no, 담당자: 담당,
      품목수: a.전체, 진행률: pct, 상태: 상태,
      비고: 상태 === ENUM.헤더상태.예외 ? ('취소 · ' + (a.예외사유 || '사유 미기재')) : ''
    });

    if (담당) {
      if (!담당맵[담당]) 담당맵[담당] = { 이름: 담당, 슬롯수: 0, 완료: 0, 진행: 0, 전체라인: 0, 처리라인: 0 };
      var w = 담당맵[담당];
      w.슬롯수++;
      if (상태 === ENUM.헤더상태.완료) w.완료++;
      if (상태 === ENUM.헤더상태.진행) w.진행++;
      w.전체라인 += a.전체;
      w.처리라인 += a.처리;
    }
  });

  out.슬롯.sort(function (a, b) { return a.슬롯 - b.슬롯; });
  out.진행률 = out.전체라인 ? Math.round(out.처리라인 / out.전체라인 * 100) : 0;

  out.담당자 = Object.keys(담당맵).map(function (k) {
    var w = 담당맵[k];
    w.진행률 = w.전체라인 ? Math.round(w.처리라인 / w.전체라인 * 100) : 0;
    return w;
  }).sort(function (a, b) { return b.처리라인 - a.처리라인; });

  return out;
}

/* ============================================================
 *  메뉴용 일괄 갱신
 * ============================================================ */

function D0_대시보드전체갱신() {
  var 결과 = [];
  try { D1_주문현황갱신(); 결과.push('✅ 주문현황'); }
  catch (e) { 결과.push('❌ 주문현황: ' + e.message); }
  try { D2_재고현황갱신(); 결과.push('✅ 재고현황'); }
  catch (e) { 결과.push('❌ 재고현황: ' + e.message); }
  try { D3_피킹현황갱신(); 결과.push('✅ 피킹현황'); }
  catch (e) { 결과.push('❌ 피킹현황: ' + e.message); }

  alert_('대시보드 갱신\n\n' + 결과.join('\n'));
  return 결과.join(' / ');
}