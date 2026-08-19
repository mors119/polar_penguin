/** 상품마스터의 창고 소유 필드인 기본보관위치만 관리한다. */
function 위치_미지정상품관리() {
  var html = HtmlService.createHtmlOutputFromFile('LocationManagement').setWidth(1050).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, '상품 위치 관리');
}

// HtmlService에서 호출할 공개 endpoint. 내부 함수는 테스트 가능한 business helper로 유지한다.
function getProductsForLocationManagement(query, includeAll) {
  return getProductsForLocationManagement_(query, includeAll);
}

function saveProductLocations(updates) { return saveProductLocations_(updates); }

function getProductsForLocationManagement_(query, includeAll) {
  query = toStr_(query).toLowerCase(); includeAll = !!includeAll;
  var master = readTable_(ROLE.마스터);
  var M = {
    sku: col_(master, COL.상품품목코드, true), name: col_(master, COL.상품명, true),
    option: col_(master, COL.옵션명, false), available: col_(master, COL.가용재고, true),
    reserved: col_(master, COL.예약재고, true), location: col_(master, COL.기본보관위치, true)
  };
  var result = [];
  master.rows.forEach(function (row) {
    var sku = toStr_(row[M.sku]); if (!sku) return;
    var location = toStr_(row[M.location]);
    if (!includeAll && location) return;
    var item = { sku: sku, productName: toStr_(row[M.name]), option: M.option >= 0 ? toStr_(row[M.option]) : '',
      available: toNum_(row[M.available]), reserved: toNum_(row[M.reserved]), location: location };
    var haystack = [item.sku, item.productName, item.option, item.location].join(' ').toLowerCase();
    if (!query || haystack.indexOf(query) >= 0) result.push(item);
  });
  result.sort(function (a, b) { return a.sku.localeCompare(b.sku); });
  return result;
}

function saveProductLocations_(updates) {
  return withLock_(function () {
    updates = Array.isArray(updates) ? updates : [];
    var master = readTable_(ROLE.마스터);
    var skuCol = col_(master, COL.상품품목코드, true), locationCol = col_(master, COL.기본보관위치, true);
    var index = {}, duplicateMasterSku = {};
    master.rows.forEach(function (row, i) {
      var sku = toStr_(row[skuCol]); if (!sku) return;
      if (index[sku] !== undefined) duplicateMasterSku[sku] = true;
      else index[sku] = i;
    });
    var updated = 0, failed = [], seen = {};
    updates.forEach(function (update) {
      var sku = toStr_(update && update.sku), location = toStr_(update && update.location);
      if (!sku) { failed.push({ sku: '', reason: '상품코드 없음' }); return; }
      if (seen[sku]) { failed.push({ sku: sku, reason: '중복 요청' }); return; }
      seen[sku] = true;
      if (duplicateMasterSku[sku]) { failed.push({ sku: sku, reason: '상품마스터의 상품코드가 중복됨' }); return; }
      var rowIndex = index[sku];
      if (rowIndex === undefined || toStr_(master.rows[rowIndex][skuCol]) !== sku) {
        failed.push({ sku: sku, reason: '상품을 찾을 수 없음' }); return;
      }
      if (toStr_(master.rows[rowIndex][locationCol]) === location) return;
      master.rows[rowIndex][locationCol] = location; updated++;
    });
    if (updated) writeColumn_(master.sheet, locationCol, master.rows);
    writeOpLog_('saveProductLocations_', failed.length ? '부분성공' : '성공', '갱신 ' + updated + ' / 실패 ' + failed.length);
    try { D0_대시보드전체갱신(true); } catch (ignore) { }
    return { updated: updated, failed: failed };
  });
}
