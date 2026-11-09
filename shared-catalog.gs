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

// ?history=1 → 발급 내역(최근 것부터), 없으면 품목 목록
function doGet(e) {
  if (e && e.parameter && e.parameter.history) {
    return json_({ ok: true, history: readHistory_(Number(e.parameter.limit) || 500) });
  }
  return json_({ ok: true, data: read_(), historyApi: 2 });
}

// 바꾼 내용을 작업 단위로 받아 지금 저장된 목록에 합친다 (여러 사람이 동시에 고쳐도 서로 지우지 않게)
// history: [발급 내역] 이 오면 'history' 시트에 한 줄씩 더한다
// cancel: [{id, canceled, cancelNote}] 발급 취소 표시 (canceled 가 비어 있으면 취소 철회). 줄은 지우지 않는다
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var req = JSON.parse(e.postData.contents || "{}");
    var out = { ok: true };
    if (req.history && req.history.length) { addHistory_(req.history); out.saved = req.history.length; }
    if (req.cancel && req.cancel.length) out.canceled = cancelHistory_(req.cancel);
    if (out.saved != null || out.canceled != null) {
      if (!req.ops || !req.ops.length) return json_(out);
    }
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
// ---------- 발급 내역 ('history' 시트, 한 줄 = 한 번 발급) ----------
var HIST_COLS = ["id", "at", "kind", "parent", "doc", "client", "useDate", "issueDate", "total", "sender", "use", "files", "snapshot", "canceled", "cancelNote"];
function histSheet_() {
  var ss = sheet_().getParent();
  var s = ss.getSheetByName("history");
  if (!s) {
    s = ss.insertSheet("history");
    s.getRange(1, 1, s.getMaxRows(), HIST_COLS.length).setNumberFormat("@");   // 날짜가 시간대 바뀐 날짜값으로 바뀌지 않게 모두 글자로
    s.appendRow(HIST_COLS);
    s.setFrozenRows(1);
  } else if (s.getRange(1, HIST_COLS.length).getValue() !== HIST_COLS[HIST_COLS.length - 1]) {
    // 예전에 만든 시트: 새 칸(발급 취소) 머리글만 채운다
    s.getRange(1, 1, 1, HIST_COLS.length).setValues([HIST_COLS]);
    s.getRange(1, 1, s.getMaxRows(), HIST_COLS.length).setNumberFormat("@");
  }
  return s;
}
// 발급 취소 / 취소 철회: 그 줄의 canceled · cancelNote 칸만 바꾼다
function cancelHistory_(list) {
  var s = histSheet_();
  var n = s.getLastRow();
  if (n < 2) return 0;
  var ids = s.getRange(2, 1, n - 1, 1).getValues().map(function (r) { return String(r[0]); });
  var col = HIST_COLS.indexOf("canceled") + 1, done = 0;
  list.forEach(function (c) {
    var i = ids.indexOf(String(c && c.id));
    if (i < 0) return;
    s.getRange(i + 2, col, 1, 2).setValues([[c.canceled || "", c.cancelNote || ""]]);
    done++;
  });
  return done;
}
function addHistory_(items) {
  var s = histSheet_();
  var have = {};
  var n = s.getLastRow();
  if (n > 1) s.getRange(2, 1, n - 1, 1).getValues().forEach(function (r) { have[r[0]] = true; });
  items.forEach(function (h) {
    if (!h || !h.id || have[h.id]) return;   // 같은 발급을 두 번 보내도 한 줄만
    s.appendRow(HIST_COLS.map(function (k) {
      var v = h[k];
      return k === "snapshot" || k === "files" ? JSON.stringify(v || null) : (v == null ? "" : v);
    }));
    have[h.id] = true;
  });
}
function readHistory_(limit) {
  var s = histSheet_();
  var n = s.getLastRow();
  if (n < 2) return [];
  var from = Math.max(2, n - limit + 1);
  var rows = s.getRange(from, 1, n - from + 1, HIST_COLS.length).getValues();
  return rows.reverse().map(function (r) {
    var o = {};
    HIST_COLS.forEach(function (k, i) {
      var v = r[i];
      if (k === "snapshot" || k === "files") { try { v = JSON.parse(v); } catch (e) { v = null; } }
      if (v instanceof Date) v = v.toISOString();
      o[k] = v;
    });
    return o;
  });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
