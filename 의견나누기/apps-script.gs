/**
 * 스포츠 도슨트 — 학생 의견 수집용 Google Apps Script
 *
 * 사용법은 같은 폴더의 "README-설정방법.md" 파일을 참고하세요.
 * 이 코드는 구글 시트에 붙어 있는 Apps Script 편집기에 그대로 붙여넣습니다.
 */

var SHEET_NAME = '의견';
var HEADERS = ['제출시각', '학번', '이름', '주제', '의견', '수업소감', '요청ID'];
var TOPICS = [
  '01 감독들은 왜 다른 옷을 입을까?',
  '02 기후변화와 스포츠의 위기',
  '03 오사카의 상징과 마라토너',
  '04 장비가 기록을 바꾼다 (기술도핑)',
  '05 축구와 인종차별',
  '06 운동과 음악의 과학',
  '07 프로야구 이전의 야구',
  '08 스포츠 음료의 시작 (게토레이)',
  '09 스포츠 기술의 혁신 (포스베리 플롭)',
  '10 e스포츠와 페이커 이상혁',
  '11 줄어드는 남녀 기록 격차',
  '12 체육 선생님이 되는 길',
  '13 스포츠에 깃든 전통의 혼 (하카)',
  '14 일본의 학교 스포츠클럽 (부카츠)',
  '15 유니폼, 하나되고 싶은 마음',
  '16 야구의 징크스 — 파울라인',
  '17 룰루레몬 창업자의 비하인드',
  '18 메이저리그 맞춤형 배트 제작',
  '19 스타디움과 파크의 차이',
  '20 팬심을 가장한 폭력',
  '21 도핑을 허용하는 대회 (인핸스드 게임)'
];

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOut({ result: 'error', message: '전달된 데이터가 없습니다.' });
    }

    var body = JSON.parse(e.postData.contents);

    var 학번 = String(body['학번'] || '').trim();
    var 이름 = String(body['이름'] || '').trim();
    var 수업소감 = String(body['수업소감'] || '').trim();
    var 항목 = body['항목'];
    var 요청ID = String(body['요청ID'] || '').trim();

    if (!/^\d{5}$/.test(학번)) {
      return jsonOut({ result: 'error', message: '학번은 숫자 5자리로 입력해 주세요.' });
    }
    if (!이름 || 이름.length > 20) {
      return jsonOut({ result: 'error', message: '필수 항목이 비어 있습니다.' });
    }
    if (!/^[a-f0-9]{32}$/.test(요청ID)) {
      return jsonOut({ result: 'error', message: '제출 요청 정보가 올바르지 않습니다.' });
    }
    if (!Array.isArray(항목) || 항목.length === 0) {
      return jsonOut({ result: 'error', message: '주제별 의견이 없습니다.' });
    }
    if (항목.length > 21) {
      return jsonOut({ result: 'error', message: '주제가 너무 많습니다.' });
    }
    if (수업소감.length > 5000) {
      return jsonOut({ result: 'error', message: '입력이 너무 깁니다.' });
    }

    // 주제 1개당 1행. 학번/이름/수업소감은 각 행에 함께 기록합니다.
    var now = new Date();
    var rows = [];
    var seen = {};

    for (var i = 0; i < 항목.length; i++) {
      var 주제 = String(항목[i]['주제'] || '').trim();
      var 의견 = String(항목[i]['의견'] || '').trim();

      if (!주제 || !의견) {
        return jsonOut({ result: 'error', message: '비어 있는 주제 또는 의견이 있습니다.' });
      }
      if (TOPICS.indexOf(주제) === -1) {
        return jsonOut({ result: 'error', message: '선택할 수 없는 주제입니다.' });
      }
      if (seen[주제]) {
        return jsonOut({ result: 'error', message: '같은 주제를 두 번 선택할 수 없습니다.' });
      }
      seen[주제] = true;
      if (의견.length > 5000) {
        return jsonOut({ result: 'error', message: '입력이 너무 깁니다.' });
      }
      rows.push([now, 학번, safeCellText(이름), 주제, safeCellText(의견), safeCellText(수업소감), 요청ID]);
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var sheet = getSheet();
      if (hasSubmissionId(sheet, 요청ID)) {
        return jsonOut({ result: 'success', saved: 0, duplicate: true });
      }
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
    } finally {
      lock.releaseLock();
    }

    return jsonOut({ result: 'success', saved: rows.length });
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    return jsonOut({ result: 'error', message: '저장 중 오류가 발생했습니다. 잠시 후 다시 시도해 주세요.' });
  }
}

function doGet() {
  return jsonOut({ result: 'ok', message: 'sports_docent opinion endpoint' });
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  sheet.hideColumns(HEADERS.length);
  return sheet;
}

function hasSubmissionId(sheet, submissionId) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  var ids = sheet.getRange(2, HEADERS.length, lastRow - 1, 1).getDisplayValues();
  return ids.some(function (row) {
    return row[0] === submissionId;
  });
}

function safeCellText(value) {
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
