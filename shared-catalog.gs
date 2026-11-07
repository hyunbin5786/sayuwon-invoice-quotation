/**
 * 견적서&거래명세서 발행기 — 공유 품목 목록 저장소 (구글 시트 Apps Script)
 *
 * 설치
 *  1. https://script.google.com 에서 '새 프로젝트' (또는 구글 시트 → 확장 프로그램 → Apps Script)
 *  2. 기존 내용을 지우고 이 파일 내용을 통째로 붙여넣고 저장
 *  3. 위쪽 함수 목록에서 setup 을 골라 '실행' → 권한 승인 (내 드라이브에 '견적서 품목 목록' 시트가 생김)
 *  4. 배포 → 새 배포 → 유형: 웹 앱
 *       실행 사용자: 나 / 액세스 권한: 모든 사용자
 *  5. 나오는 웹 앱 URL(https://script.google.com/macros/s/…/exec)을 index.html 의 SHARED_URL 에 넣는다
 *
 * 저장 형식: 'catalog' 시트 A1 = 목록 JSON, B1 = 마지막 저장 시각
 *  { edits: { 예약상품명: { 바꾼 칸: 값 } }, added: [ { uid, src, name, kind, form, price, disc } ] }
 */

function doGet() {
  return json_({ ok: true, data: read_() });
}

// 바꾼 내용을 작업 단위로 받아 지금 저장된 목록에 합친다 (여러 사람이 동시에 고쳐도 서로 지우지 않게)
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var req = JSON.parse(e.postData.contents || "{}");
    var d = read_();
    (req.ops || []).forEach(function (o) { d = apply_(d, o); });
    write_(d);
    return json_({ ok: true, data: d });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function apply_(d, o) {
  d = d || {};
  if (o.op === "set") {
    d.edits = d.edits || {};
    d.edits[o.src] = d.edits[o.src] || {};
    d.edits[o.src][o.field] = o.value;
  } else if (o.op === "setAdded") {
    (d.added || []).forEach(function (a) { if (String(a.uid) === String(o.uid)) a[o.field] = o.value; });
  } else if (o.op === "add") {
    d.added = d.added || [];
    if (!d.added.some(function (a) { return String(a.uid) === String(o.item.uid); })) d.added.unshift(o.item);
  } else if (o.op === "delAdded") {
    d.added = (d.added || []).filter(function (a) { return String(a.uid) !== String(o.uid); });
  } else if (o.op === "reset") {
    d = {};
  } else if (o.op === "replace") {
    d = o.data || {};
  }
  return d;
}

// 시트에 붙인 스크립트면 그 시트를, script.google.com 에서 따로 만든 스크립트면
// 처음 실행할 때 '견적서 품목 목록' 시트를 내 드라이브에 만들어 쓴다
function sheet_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) {
    var props = PropertiesService.getScriptProperties();
    var id = props.getProperty("SHEET_ID");
    if (id) ss = SpreadsheetApp.openById(id);
    else {
      ss = SpreadsheetApp.create("견적서 품목 목록");
      props.setProperty("SHEET_ID", ss.getId());
    }
  }
  return ss.getSheetByName("catalog") || ss.insertSheet("catalog");
}

// 배포 전에 한 번 실행해서 권한을 승인하고 시트를 만든다
function setup() {
  var s = sheet_();
  Logger.log("저장 시트: " + s.getParent().getUrl());
}
function read_() {
  var v = sheet_().getRange("A1").getValue();
  return v ? JSON.parse(v) : {};
}
function write_(d) {
  var s = sheet_();
  s.getRange("A1").setValue(JSON.stringify(d));
  s.getRange("B1").setValue(new Date());
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
