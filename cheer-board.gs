/**
 * 뿅아리 응원 한마디 — 구글 시트 저장용 앱스스크립트
 * ------------------------------------------------------------
 * 쓰는 법은 아래 5단계예요.
 *
 *  1. 구글 드라이브에서 새 스프레드시트를 하나 만들고 이름을 붙이세요.
 *     (예: "2026 머니레터 응원")
 *  2. 그 시트에서 상단 메뉴 [확장 프로그램] > [Apps Script]를 눌러요.
 *  3. 편집기에 있던 기본 코드를 전부 지우고, 이 파일 내용을 통째로 붙여넣고 저장(Ctrl+S).
 *  4. 오른쪽 위 [배포] > [새 배포] > 톱니바퀴에서 [웹 앱] 선택
 *       - 설명       : 아무거나 (예: 응원게시판 v1)
 *       - 실행 계정  : 나
 *       - 액세스 권한: ★ 모든 사용자 ★   <- 이게 제일 중요해요!
 *     [배포]를 누르고 권한 승인 창이 뜨면 전부 허용해주세요.
 *     ("안전하지 않은 페이지로 이동" 경고가 떠도 내 스크립트니까 진행하시면 돼요)
 *  5. 마지막에 나오는 "웹 앱 URL"(.../exec 로 끝나요)을 복사해서,
 *     머니레터 HTML 안의  const CHEER_API = '';  따옴표 사이에 붙여넣으세요.
 *
 * ------------------------------------------------------------
 * 글 관리하는 법
 *  - 시트 "응원" 탭에 글이 한 줄씩 쌓여요.
 *  - 지우고 싶은 글은 F열(숨김)에 TRUE 를 넣으면 레터에서 바로 안 보여요.
 *    (행을 통째로 지우면 번호가 꼬일 수 있어서 숨김 처리를 추천드려요)
 *  - 한 사람당 한 글이고, 다시 쓰면 기존 글이 덮어써져요.
 *  - 닉네임(뿅아리1234)은 레터가 만들어 보내고, 겹치면 여기서 다른 번호로 바꿔줘요.
 *    B열 닉네임을 직접 고쳐도 되는데, 다른 사람과 겹치지 않게만 해주세요.
 *
 * 코드를 고친 뒤에는 [배포] > [배포 관리] > 연필(수정) > 버전 "새 버전" > [배포]
 * 를 해야 반영돼요. 이때 URL은 그대로라 레터는 안 고쳐도 됩니다.
 */

var SHEET_NAME = '응원';
var MAX_LEN = 300;

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['번호', '닉네임', 'uid', '내용', '작성시각', '숨김']);
    sh.setFrozenRows(1);
    sh.setColumnWidth(4, 520);
  }
  return sh;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function hidden_(v) {
  if (v === true) return true;
  var s = String(v).trim().toUpperCase();
  return s === 'TRUE' || s === 'O' || s === 'Y' || s === '숨김';
}

/* 레터가 글 목록을 읽어갈 때 */
function doGet() {
  try {
    var sh = sheet_();
    var last = sh.getLastRow();
    var items = [];
    if (last > 1) {
      var rows = sh.getRange(2, 1, last - 1, 6).getValues();
      for (var i = 0; i < rows.length; i++) {
        var r = rows[i];
        if (hidden_(r[5])) continue;
        var text = String(r[3] || '').trim();
        if (!text) continue;
        items.push({
          n: Number(r[0]) || 0,
          nick: String(r[1] || ''),
          uid: String(r[2] || ''),
          text: text,
          at: r[4] ? new Date(r[4]).toISOString() : ''
        });
      }
    }
    items.reverse(); // 최신 글이 위로
    return json_({ ok: true, items: items });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* 레터에서 글을 남길 때 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json_({ ok: false, error: 'busy' });
  }
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var uid = String(body.uid || '').slice(0, 64);
    var want = String(body.nick || '').trim().slice(0, 20);
    var text = String(body.text || '').replace(/\r/g, '').trim().slice(0, MAX_LEN);
    if (!uid || !text) return json_({ ok: false, error: 'empty' });

    var sh = sheet_();
    var last = sh.getLastRow();
    var rowIdx = 0;

    if (last > 1) {
      var uids = sh.getRange(2, 3, last - 1, 1).getValues();
      for (var i = 0; i < uids.length; i++) {
        if (String(uids[i][0]) === uid) { rowIdx = i + 2; break; }
      }
    }

    var now = new Date();
    var n, nick;

    if (rowIdx) {
      // 이미 쓴 사람 -> 내용만 갈아끼우기 (번호와 닉네임은 유지)
      n = Number(sh.getRange(rowIdx, 1).getValue()) || 0;
      nick = String(sh.getRange(rowIdx, 2).getValue() || ('뿅아리' + n));
      sh.getRange(rowIdx, 4).setValue(text);
      sh.getRange(rowIdx, 5).setValue(now);
    } else {
      // 처음 쓰는 사람 -> 순번(n)과 닉네임을 새로 부여
      n = 1;
      var used = {};
      if (last > 1) {
        var old = sh.getRange(2, 1, last - 1, 2).getValues();
        for (var k = 0; k < old.length; k++) {
          var v = Number(old[k][0]) || 0;
          if (v >= n) n = v + 1;
          used[String(old[k][1]).trim()] = true;
        }
      }
      // 레터가 보내온 번호를 그대로 쓰되, 이미 쓰는 사람이 있으면 다른 번호로
      nick = (/^뿅아리[0-9]{1,6}$/.test(want) && !used[want]) ? want : '';
      for (var g = 0; !nick && g < 400; g++) {
        var cand = '뿅아리' + (1000 + Math.floor(Math.random() * 9000));
        if (!used[cand]) nick = cand;
      }
      if (!nick) nick = '뿅아리' + (Date.now() % 1000000);
      sh.appendRow([n, nick, uid, text, now, false]);
    }

    return json_({ ok: true, n: n, nick: nick, at: now.toISOString() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}
