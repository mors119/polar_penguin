/* ============================================================
 *  초기화  ⚠ 되돌릴 수 없습니다
 * ============================================================ */

/**
 * 1단계. 무엇이 지워질지 미리 확인한다. (실제로 지우지 않음)
 */
function 초기화_1_미리보기() {
  var out = ['=== 초기화하면 지워질 것 ===', ''];

  [[ROLE.주문, '주문(완료)'], [ROLE.헤더, '피킹(헤더)'], [ROLE.라인, '피킹(라인)']]
    .forEach(function (p) {
      try {
        var t = readTable_(p[0]);
        out.push('  · ' + p[1] + ' — ' + t.rows.length + '행 전부 삭제');
      } catch (e) { out.push('  · ' + p[1] + ' — ❌ ' + e.message); }
    });

  try {
    var m = readTable_(ROLE.마스터);
    var c예약 = col_(m, COL.예약재고, false);
    var 예약합 = 0;
    if (c예약 >= 0) m.rows.forEach(function (r) { 예약합 += toNum_(r[c예약]); });
    out.push('  · 상품마스터 — 행은 유지, 예약재고만 0으로 (현재 합계 ' + 예약합 + ')');
  } catch (e) { out.push('  · 상품마스터 — ❌ ' + e.message); }

  var ss = consoleSS_();
  [CONSOLE.재고이동로그, CONSOLE.작업로그].forEach(function (n) {
    var sh = ss.getSheetByName(n);
    out.push('  · ' + n + ' — ' + (sh ? Math.max(sh.getLastRow() - 1, 0) + '행 삭제' : '탭 없음'));
  });

  out.push('');
  out.push('=== 유지되는 것 ===');
  out.push('  · 상품마스터의 상품 정보와 보관위치');
  out.push('  · 설정 탭 (파일 ID · 파라미터 · 별칭)');
  out.push('  · 대시보드 탭');
  out.push('');
  out.push('실행하려면  초기화_2_전체실행("초기화")  를 편집기에서 호출하세요.');

  var msg = out.join('\n');
  Logger.log(msg);
  alert_(msg);
  return msg;
}

/**
 * 2단계. 실제 초기화.
 * @param {String} 확인문구  반드시 "초기화" 를 넘겨야 실행된다.
 */
function 초기화_2_전체실행(확인문구) {
  if (확인문구 !== '초기화') {
    throw new Error(
      '안전장치입니다. 실행하려면 인자를 넣으세요.\n\n' +
      '  초기화_2_전체실행("초기화")\n\n' +
      '먼저 초기화_1_미리보기() 로 지워질 내용을 확인하시기 바랍니다.'
    );
  }

  return withLock_(function () {
    var out = [];

    // ---------- 데이터 시트 3종: 헤더만 남기고 삭제 ----------
    [[ROLE.주문, '주문(완료)'], [ROLE.헤더, '피킹(헤더)'], [ROLE.라인, '피킹(라인)']]
      .forEach(function (p) {
        try {
          var sh = getSheet_(p[0]);
          var last = sh.getLastRow();
          if (last > 1) {
            sh.deleteRows(2, last - 1);
            out.push('✅ ' + p[1] + ' — ' + (last - 1) + '행 삭제');
          } else {
            out.push('· ' + p[1] + ' — 이미 비어 있음');
          }
        } catch (e) {
          out.push('❌ ' + p[1] + ': ' + e.message);
        }
      });

    // ---------- 상품마스터: 예약재고만 0으로 ----------
    try {
      var m = readTable_(ROLE.마스터);
      var c예약 = col_(m, COL.예약재고, false);
      if (c예약 >= 0 && m.rows.length) {
        m.sheet.getRange(2, c예약 + 1, m.rows.length, 1)
          .setValues(m.rows.map(function () { return [0]; }));
        out.push('✅ 상품마스터 — 예약재고 ' + m.rows.length + '행을 0으로 초기화');
      }
    } catch (e) {
      out.push('❌ 상품마스터: ' + e.message);
    }

    // ---------- 로그 비우기 ----------
    var ss = consoleSS_();
    [CONSOLE.재고이동로그, CONSOLE.작업로그].forEach(function (n) {
      var sh = ss.getSheetByName(n);
      if (!sh) return;
      var last = sh.getLastRow();
      if (last > 1) {
        sh.deleteRows(2, last - 1);
        out.push('✅ ' + n + ' — ' + (last - 1) + '행 삭제');
      }
    });

    // ---------- 대시보드 다시 그리기 ----------
    try { D0_대시보드전체갱신(); out.push('✅ 대시보드 갱신'); }
    catch (e) { out.push('· 대시보드 갱신 실패 (무시 가능)'); }

    var msg = '초기화 완료\n\n' + out.join('\n') +
      '\n\n다음 단계\n' +
      '  ① 카페24 재고 파일을 Input에 업로드\n' +
      '  ② 주문 파일을 같은 Input에 업로드\n' +
      '  ③ 트리거를 기다리거나 processInput() 실행';

    Logger.log(msg);
    alert_(msg);
    writeOpLog_('초기화_2_전체실행', '성공', out.join(' / '));
    return msg;
  });
}

/**
 * 상품마스터까지 전부 비운다. 재고 정보가 사라지므로 신중히.
 */
function 초기화_9_마스터포함(확인문구) {
  if (확인문구 !== '마스터까지초기화') {
    throw new Error(
      '⚠ 상품마스터의 보관위치까지 전부 사라집니다.\n\n' +
      '정말 실행하려면:\n  초기화_9_마스터포함("마스터까지초기화")\n\n' +
      '보통은 초기화_2_전체실행("초기화") 로 충분합니다.'
    );
  }

  초기화_2_전체실행('초기화');

  var sh = getSheet_(ROLE.마스터);
  var last = sh.getLastRow();
  if (last > 1) sh.deleteRows(2, last - 1);

  var msg = '상품마스터까지 초기화했습니다 (' + Math.max(last - 1, 0) + '행 삭제).\n\n' +
    '카페24 재고 파일을 Input에 업로드해 상품을 다시 등록하고,\n' +
    '보관위치를 다시 입력해야 합니다.';
  Logger.log(msg);
  alert_(msg);
  return msg;
}
