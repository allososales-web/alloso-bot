/**
 * Alloso 만능 취합 봇 - Google Apps Script
 * 
 * [설치 방법]
 * 1. 구글 시트 > 확장 프로그램 > Apps Script
 * 2. 이 코드 전체 붙여넣기
 * 3. 배포 > 새 배포 > 웹앱 > 실행:나, 액세스:모든 사용자
 * 4. URL 복사 후 HTML의 APPS_SCRIPT_URL에 입력
 */

var SPREADSHEET_ID = '1XxOTaLP-UF3WD8bKZ9gzOk_kaxgm8e7XVz-2aTbQhJ8';

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  var callback = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : '';

  // action이 없으면 HTML 페이지 서빙
  if (!action) {
    return HtmlService.createHtmlOutput(getPageHtml())
      .setTitle('Alloso 취합봇')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  }

  var result;
  switch(action) {
    case 'getSettings':
      result = getSettings();
      break;
    case 'getResponses':
      result = getResponses();
      break;
    default:
      result = { error: 'Unknown action' };
  }

  var output = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + output + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var result;

    switch(action) {
      case 'saveSettings':
        result = saveSettings(data);
        break;
      case 'saveResponse':
        result = saveResponse(data);
        break;
      default:
        result = { error: 'Unknown action' };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function saveSettings(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('설정');
  if (!sheet) sheet = ss.insertSheet('설정');
  sheet.clear();
  sheet.getRange(1,1,1,7).setValues([['조사제목','조사설명','질문목록','대상매장','생성일시','전체조사목록','중복허용']]);
  sheet.getRange(2,1,1,7).setValues([[
    data.title || '',
    data.desc || '',
    JSON.stringify(data.questions || []),
    JSON.stringify(data.stores || []),
    data.createdAt || new Date().toLocaleString('ko-KR'),
    JSON.stringify(data.allSurveys || []),
    data.allowDuplicate ? 'Y' : 'N'
  ]]);
  return { success: true };
}

function getSettings() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('설정');
  if (!sheet || sheet.getLastRow() < 2) return null;
  var cols = sheet.getLastColumn();
  var row = sheet.getRange(2,1,1,cols).getValues()[0];
  var result = {
    title: row[0] || '',
    desc: row[1] || '',
    questions: JSON.parse(row[2] || '[]'),
    stores: JSON.parse(row[3] || '[]'),
    createdAt: row[4] || '',
    allSurveys: [],
    allowDuplicate: false
  };
  if (cols >= 6 && row[5]) {
    try { result.allSurveys = JSON.parse(row[5]); } catch(e) { result.allSurveys = []; }
  }
  if (cols >= 7) {
    result.allowDuplicate = (row[6] === 'Y');
  }
  return result;
}

function saveResponse(data) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('답변');
  if (!sheet) sheet = ss.insertSheet('답변');
  if (sheet.getLastRow() === 0) {
    var settings = getSettings();
    var questions = settings ? settings.questions.map(function(q){ return q.text; }) : Object.keys(data.answers || {});
    var headers = ['조사제목','매장','답변자'].concat(questions).concat(['제출시간']);
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
  }
  // 헤더에 없는 질문이 있으면 컬럼 추가
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var answerKeys = Object.keys(data.answers || {});
  answerKeys.forEach(function(k) {
    if (headers.indexOf(k) === -1) {
      headers.push(k);
      sheet.getRange(1, headers.length).setValue(k);
    }
  });
  var row = headers.map(function(h) {
    if (h === '조사제목') return data.surveyTitle || '';
    if (h === '매장') return data.store || '';
    if (h === '답변자') return data.responder || '';
    if (h === '제출시간') return data.timestamp || new Date().toLocaleString('ko-KR');
    return (data.answers && data.answers[h]) ? data.answers[h] : '';
  });
  sheet.getRange(sheet.getLastRow()+1, 1, 1, row.length).setValues([row]);
  return { success: true };
}

function getResponses() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('답변');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var rows = sheet.getRange(2,1,sheet.getLastRow()-1,sheet.getLastColumn()).getValues();
  return rows.map(function(row) {
    var obj = { answers: {} };
    headers.forEach(function(h, i) {
      if (h === '조사제목') obj.surveyTitle = row[i];
      else if (h === '매장') obj.store = row[i];
      else if (h === '답변자') obj.responder = row[i];
      else if (h === '제출시간') obj.timestamp = row[i];
      else obj.answers[h] = row[i];
    });
    return obj;
  });
}

function getPageHtml() {
  // aloso_survey_bot.html의 API 부분만 google.script.run으로 교체
  var html = HtmlService.createHtmlOutputFromFile('page').getContent();
  return html;
}
