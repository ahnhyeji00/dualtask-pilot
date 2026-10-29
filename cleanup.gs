/**
 * 시트 정리 — 한 번만 누르면 되는 함수
 *
 * 쓰는 법
 *   1. 스프레드시트에서 파일 > 사본 만들기 로 백업을 먼저 뜬다. 되돌릴 수 없다.
 *   2. 확장 프로그램 > Apps Script 를 열고 이 파일을 새 파일로 붙여넣는다.
 *      (collector.gs 는 그대로 두고, + 버튼으로 스크립트를 하나 더 만든다)
 *   3. 함수 목록에서 시트정리 를 고르고 실행을 누른다.
 *   4. 끝나면 무엇을 했는지 창으로 알려준다.
 *
 * 하는 일
 *   - ZZ 로 시작하는 시험용 행을 모두 지운다
 *   - 기록 시트의 '진행중' 줄을 지운다 (블록 단위 전송의 흔적)
 *   - A10 의 두 번째 시행을 지운다 (첫 번째만 남긴다)
 *   - A20 의 A·D·E 시행 블록 번호를 5~10 으로 바꿔 B·C 와 한 사람으로 잇는다
 *   - A10·A20 의 표 D4 행을 지운다 (파생값은 다른 표에서 다시 계산하면 된다)
 *
 * 두 번 실행해도 안전하다. 이미 정리된 것은 건너뛴다.
 */

var ID_HEADERS = ['참가자ID', '참가자 ID'];

function 시트정리() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = [];

  log.push('■ 시험용 행 (ZZ…)');
  log.push(deleteWhere_(ss, function (pid) { return /^ZZ/i.test(pid); }));

  log.push('');
  log.push('■ 기록 시트의 진행중 줄');
  log.push(deleteProgressRows_(ss));

  log.push('');
  log.push('■ A10 두 번째 시행 삭제');
  log.push(keepFirstRunOnly_(ss, 'A10'));

  log.push('');
  log.push('■ A20 블록 번호 잇기 (A·D·E → 5~10)');
  log.push(renumberLaterRun_(ss, 'A20', 4));

  log.push('');
  log.push('■ A10·A20 의 표 D4 행 삭제 (파생값은 다시 계산)');
  log.push(deleteWhere_(ss, function (pid) { return pid === 'A10' || pid === 'A20'; },
                        ['표 D4 파생값']));

  SpreadsheetApp.getUi().alert('시트 정리 결과\n\n' + log.join('\n'));
}

/** 참가자ID 열의 위치. 없으면 -1 */
function idCol_(sh) {
  if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) return -1;
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  for (var i = 0; i < ID_HEADERS.length; i++) {
    var k = head.indexOf(ID_HEADERS[i]);
    if (k >= 0) return k + 1;
  }
  return -1;
}

function colOf_(sh, name) {
  var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  var k = head.indexOf(name);
  return k >= 0 ? k + 1 : -1;
}

/** 조건에 맞는 행을 모든 시트(또는 지정한 시트)에서 지운다 */
function deleteWhere_(ss, matchPid, onlySheets) {
  var out = [];
  ss.getSheets().forEach(function (sh) {
    if (onlySheets && onlySheets.indexOf(sh.getName()) < 0) return;
    var c = idCol_(sh);
    if (c < 0 || sh.getLastRow() < 2) return;
    var vals = sh.getRange(2, c, sh.getLastRow() - 1, 1).getValues();
    var n = 0;
    for (var r = vals.length - 1; r >= 0; r--) {          // 아래에서 위로 지워야 번호가 안 밀린다
      if (matchPid(String(vals[r][0]).trim())) { sh.deleteRow(r + 2); n++; }
    }
    if (n) out.push('   ' + sh.getName() + ': ' + n + '행 삭제');
  });
  return out.length ? out.join('\n') : '   지울 것 없음';
}

/** 기록 시트에서 조건 순서가 '진행중' 으로 시작하는 줄을 지운다 */
function deleteProgressRows_(ss) {
  var out = [];
  ss.getSheets().forEach(function (sh) {
    var c = colOf_(sh, '조건 순서');
    if (c < 0 || sh.getLastRow() < 2) return;
    var vals = sh.getRange(2, c, sh.getLastRow() - 1, 1).getValues();
    var n = 0;
    for (var r = vals.length - 1; r >= 0; r--) {
      if (String(vals[r][0]).indexOf('진행중') === 0) { sh.deleteRow(r + 2); n++; }
    }
    if (n) out.push('   ' + sh.getName() + ': ' + n + '행 삭제');
  });
  return out.length ? out.join('\n') : '   지울 것 없음';
}

/**
 * 한 참가자의 행을 시트 순서대로 읽어, 블록 번호가 다시 작아지는 지점에서
 * 시행을 나눈다. 블록 열이 없는 시트는 전체를 한 덩어리로 본다.
 * 반환: [[시트행번호…], [시트행번호…], …]
 */
function runsOf_(sh, pid) {
  var c = idCol_(sh);
  if (c < 0 || sh.getLastRow() < 2) return [];
  var bc = colOf_(sh, '블록');
  var last = sh.getLastRow();
  var ids = sh.getRange(2, c, last - 1, 1).getValues();
  var blocks = bc > 0 ? sh.getRange(2, bc, last - 1, 1).getValues() : null;
  var runs = [], cur = [], prev = null;
  for (var r = 0; r < ids.length; r++) {
    if (String(ids[r][0]).trim() !== pid) continue;
    var b = blocks ? Number(blocks[r][0]) : null;
    if (cur.length && b !== null && prev !== null && b <= prev) { runs.push(cur); cur = []; }
    cur.push(r + 2);
    prev = b;
  }
  if (cur.length) runs.push(cur);
  return runs;
}

/** 첫 시행만 남기고 나머지 시행의 행을 지운다 */
function keepFirstRunOnly_(ss, pid) {
  var out = [];
  ss.getSheets().forEach(function (sh) {
    var runs = runsOf_(sh, pid);
    if (runs.length < 2) return;
    var drop = [];
    for (var i = 1; i < runs.length; i++) drop = drop.concat(runs[i]);
    drop.sort(function (a, b) { return b - a; });            // 아래에서 위로
    drop.forEach(function (r) { sh.deleteRow(r); });
    out.push('   ' + sh.getName() + ': ' + drop.length + '행 삭제 (시행 ' + runs.length + '개 중 첫 것만 남김)');
  });
  return out.length ? out.join('\n') : '   이미 정리됨';
}

/** 두 번째 이후 시행의 블록 번호에 offset 을 더해 하나의 시행으로 잇는다 */
function renumberLaterRun_(ss, pid, offset) {
  var out = [];
  ss.getSheets().forEach(function (sh) {
    var bc = colOf_(sh, '블록');
    if (bc < 0) return;
    var runs = runsOf_(sh, pid);
    if (runs.length < 2) return;
    var n = 0;
    for (var i = 1; i < runs.length; i++) {
      runs[i].forEach(function (r) {
        var cell = sh.getRange(r, bc);
        var v = Number(cell.getValue());
        if (v <= offset) { cell.setValue(v + offset); n++; }   // 이미 바꾼 것은 건너뛴다
      });
    }
    if (n) out.push('   ' + sh.getName() + ': ' + n + '행의 블록 번호 +' + offset);
  });
  return out.length ? out.join('\n') : '   이미 정리됨';
}
