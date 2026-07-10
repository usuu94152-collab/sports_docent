/**
 * 스포츠 도슨트 — 학생 의견 수집용 Google Apps Script
 *
 * 사용법은 같은 폴더의 "README-설정방법.md" 파일을 참고하세요.
 * 이 코드는 구글 시트에 붙어 있는 Apps Script 편집기에 그대로 붙여넣습니다.
 */

var SHEET_NAME = '의견';
var HEADERS = ['제출시각', '학번', '이름', '주제', '의견'];

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOut({ result: 'error', message: '전달된 데이터가 없습니다.' });
    }

    var body = JSON.parse(e.postData.contents);

    var 학번 = String(body['학번'] || '').trim();
    var 이름 = String(body['이름'] || '').trim();
    var 주제 = String(body['주제'] || '').trim();
    var 의견 = String(body['의견'] || '').trim();

    if (!학번 || !이름 || !주제 || !의견) {
      return jsonOut({ result: 'error', message: '필수 항목이 비어 있습니다.' });
    }

    // 지나치게 긴 입력 차단 (스팸 방지)
    if (의견.length > 5000) {
      return jsonOut({ result: 'error', message: '의견이 너무 깁니다.' });
    }

    var sheet = getSheet();
    sheet.appendRow([new Date(), 학번, 이름, 주제, 의견]);

    return jsonOut({ result: 'success' });
  } catch (err) {
    return jsonOut({ result: 'error', message: err.message });
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
  }
  return sheet;
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
