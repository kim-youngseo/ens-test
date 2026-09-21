/**
 * Ascént 플레이북 — 글 저장소 (Google Apps Script)
 *
 * 참가자가 앱에서 남긴 「한 마디 · 질문 · 큐레이션」을 구글시트에 적재하고,
 * 모든 참가자가 같은 글을 읽도록 돌려줍니다.
 *
 * ── 설치 순서 ──────────────────────────────────────────────
 * 1. 명단을 관리하는 구글시트를 엽니다.
 * 2. 확장 프로그램 › Apps Script 를 열고, 이 파일 내용을 전부 붙여넣습니다.
 * 3. 저장 후 [배포] › [새 배포] 를 누릅니다.
 *      · 유형        : 웹 앱
 *      · 실행 사용자 : 나
 *      · 액세스 권한 : 모든 사용자          ← 이 항목이 핵심입니다
 * 4. 배포하면 나오는 https://script.google.com/macros/s/....../exec 주소를 복사합니다.
 * 5. 플레이북 로그인 화면 › 「글 저장 연결」 에 그 주소를 붙여넣습니다.
 *
 * 「글」 시트는 처음 글이 올라올 때 자동으로 만들어집니다.
 * 부적절한 글은 시트에서 해당 행을 지우면 앱에서도 사라집니다.
 * ───────────────────────────────────────────────────────────
 */

var SHEET_NAME = '글';
var HEAD = ['id', 'ts', '작성시각', 'kind', 'author', 'name', 'role', 'target', 'title', 'text', 'tone'];

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 앱이 글 목록을 읽어갈 때 */
function doGet(e) {
  try {
    var sh = sheet_();
    var values = sh.getDataRange().getValues();
    var head = values.shift() || [];
    var posts = [];
    values.forEach(function (row) {
      if (!row[0]) return;                       // 빈 행은 건너뜁니다
      var o = {};
      head.forEach(function (h, i) { o[h] = row[i]; });
      o.ts = Number(o.ts) || 0;
      o.tone = Number(o.tone) || 0;
      delete o['작성시각'];                       // 사람이 보기 위한 칼럼이라 앱에는 보내지 않습니다
      posts.push(o);
    });
    posts.sort(function (a, b) { return b.ts - a.ts; });
    return json_({ ok: true, posts: posts });
  } catch (err) {
    return json_({ ok: false, error: String(err), posts: [] });
  }
}

/** 앱이 글을 남기거나 지울 때 */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var sh = sheet_();

    if (body.action === 'add' && body.post) {
      var p = body.post;
      var ts = Number(p.ts) || Date.now();
      sh.appendRow([
        p.id, ts, new Date(ts),
        p.kind || '', p.author || '', p.name || '', p.role || '',
        p.target || '', p.title || '', p.text || '', Number(p.tone) || 0
      ]);
      return json_({ ok: true });
    }

    if (body.action === 'del' && body.id) {
      var last = sh.getLastRow();
      if (last > 1) {
        var ids = sh.getRange(2, 1, last - 1, 1).getValues();
        for (var i = 0; i < ids.length; i++) {
          if (String(ids[i][0]) === String(body.id)) {
            sh.deleteRow(i + 2);
            break;
          }
        }
      }
      return json_({ ok: true });
    }

    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
