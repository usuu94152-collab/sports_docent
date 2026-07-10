/**
 * 스포츠 도슨트 — 학생 의견 수집용 Google Apps Script
 *
 * 사용법은 같은 폴더의 "README-설정방법.md" 파일을 참고하세요.
 * 이 코드는 구글 시트에 붙어 있는 Apps Script 편집기에 그대로 붙여넣습니다.
 */

var SHEET_NAME = '의견';
var HEADERS = ['제출시각', '학번', '이름', '주제', '의견', '수업소감'];

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

    if (!학번 || !이름 || !수업소감) {
      return jsonOut({ result: 'error', message: '필수 항목이 비어 있습니다.' });
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

    for (var i = 0; i < 항목.length; i++) {
      var 주제 = String(항목[i]['주제'] || '').trim();
      var 의견 = String(항목[i]['의견'] || '').trim();

      if (!주제 || !의견) {
        return jsonOut({ result: 'error', message: '비어 있는 주제 또는 의견이 있습니다.' });
      }
      if (의견.length > 5000) {
        return jsonOut({ result: 'error', message: '입력이 너무 깁니다.' });
      }
      rows.push([now, 학번, 이름, 주제, 의견, 수업소감]);
    }

    var sheet = getSheet();
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);

    return jsonOut({ result: 'success', saved: rows.length });
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
