/**
 * 결과 수집용 Google Apps Script (선택 사항)
 *
 * 설정
 *  1. 구글 드라이브에서 스프레드시트를 새로 만든다.
 *  2. 확장 프로그램 > Apps Script 를 열고, 기본 코드를 지우고 이 파일을 통째로 붙여넣는다.
 *  3. 배포 > 새 배포 > 톱니바퀴 > 웹 앱
 *       설명        아무거나
 *       실행 계정    나
 *       액세스 권한  모든 사용자        ← 이게 핵심이다. 조원이 로그인 없이 보낼 수 있어야 한다.
 *  4. 처음 배포할 때 권한 승인 창이 뜬다. '고급' > '안전하지 않음(이동)' 을 눌러 승인한다.
 *     본인이 만든 스크립트라 뜨는 경고이며, 이 코드는 받은 자료를 시트에 적기만 한다.
 *  5. 배포가 끝나면 나오는 .../exec 로 끝나는 주소를 복사해
 *     과업 페이지 설정 화면의 '수집 주소' 칸에 붙여넣는다.
 *
 * 만들어지는 시트
 *   기록            참가자별 한 줄 요약
 *   표 D1 … 학습 로그  각 표가 시트 하나씩, 참가자들의 행이 이어 붙는다
 *
 * 주의
 *   과업 페이지는 브라우저 보안 때문에 전송 결과를 읽지 못한다. 보낸 뒤 시트에
 *   줄이 늘었는지 직접 확인해야 한다. 전송이 실패해도 참가자 컴퓨터에는
 *   결과 파일이 그대로 저장되므로 자료가 사라지지는 않는다.
 */

var SUMMARY_SHEET = '기록';

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);                       // 동시에 두 명이 끝내도 줄이 섞이지 않게 한다
  try {
    var run = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    appendSummary(ss, run);
    appendTables(ss, run);

    return ContentService.createTextOutput('ok');
  } catch (err) {
    return ContentService.createTextOutput('error: ' + err);
  } finally {
    lock.releaseLock();
  }
}

/** 브라우저로 주소를 열었을 때 살아 있는지 확인용 */
function doGet() {
  return ContentService.createTextOutput('수집 서버가 작동 중입니다.');
}

function appendSummary(ss, run) {
  var s = ss.getSheetByName(SUMMARY_SHEET);
  if (!s) {
    s = ss.insertSheet(SUMMARY_SHEET, 0);
    s.appendRow(['받은 시각', '참가자 ID', '시퀀스', '범위', '알림음 설계',
                 '조건 순서', '블록 수', '알림 수', '시작', '종료', '실험자']);
    s.setFrozenRows(1);
  }
  s.appendRow([new Date(), run.pid, run.seq,
               run.mode === 'balance' ? 'B·C 점검' : '전체',
               run.toneset === 'pitch' ? '음고형' : '방향형',
               run.conditions, run.blocks, run.trialCount,
               run.startedAt, run.finishedAt, run.experimenter]);
}

function appendTables(ss, run) {
  var csv = run.csv || {};
  Object.keys(csv).forEach(function (name) {
    var text = csv[name];
    if (!text) return;

    var rows = text.split('\n')
                   .filter(function (l) { return l.trim(); })
                   .map(parseCsvLine);
    if (rows.length < 2) return;             // 머리글만 있으면 적을 게 없다

    var header = rows[0];
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.appendRow(header);
      sh.setFrozenRows(1);
    }

    // 자리가 아니라 '이름'으로 열을 맞춘다. 먼저 도착한 자료의 열 구성이
    // 달라도 값이 엉뚱한 칸에 들어가지 않는다.
    var width = Math.max(sh.getLastColumn(), 1);
    var titles = sh.getRange(1, 1, 1, width).getValues()[0].map(function (v) { return String(v); });
    var at = {};
    titles.forEach(function (t, i) { if (t !== '') at[t] = i; });

    var grew = false;
    header.forEach(function (t) {
      if (!(t in at)) { titles.push(t); at[t] = titles.length - 1; grew = true; }
    });
    if (grew) {
      sh.getRange(1, 1, 1, titles.length).setValues([titles]);
      sh.setFrozenRows(1);
    }

    var body = rows.slice(1).map(function (r) {
      var out = [];
      for (var i = 0; i < titles.length; i++) out.push('');
      header.forEach(function (t, i) { out[at[t]] = (i < r.length) ? r[i] : ''; });
      return out;
    });
    sh.getRange(sh.getLastRow() + 1, 1, body.length, titles.length).setValues(body);
  });
}

/** 따옴표로 감싼 칸과 칸 안의 쉼표를 고려해 CSV 한 줄을 나눈다 */
function parseCsvLine(line) {
  var out = [], cur = '', inQuote = false;
  for (var i = 0; i < line.length; i++) {
    var c = line.charAt(i);
    if (inQuote) {
      if (c === '"') {
        if (line.charAt(i + 1) === '"') { cur += '"'; i++; }
        else inQuote = false;
      } else cur += c;
    } else if (c === '"') {
      inQuote = true;
    } else if (c === ',') {
      out.push(cur); cur = '';
    } else {
      cur += c;
    }
  }
  out.push(cur);
  return out;
}
