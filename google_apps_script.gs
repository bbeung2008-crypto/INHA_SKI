/**
 * INHA ALPINE SKI TEAM - Google Form 응답 API
 * Google Form과 연결된 스프레드시트에서
 * 확장 프로그램 > Apps Script 에 붙여넣으세요.
 */
function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('설문지 응답 시트1') || ss.getSheets()[0];
  const values = sheet.getDataRange().getDisplayValues();

  if (values.length < 2) {
    return json_([]);
  }

  const headers = values[0].map((h, i) => h || ('열' + (i + 1)));
  const rows = values.slice(1)
    .filter(row => row.some(v => String(v).trim() !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i] || '');
      return obj;
    });

  return json_(rows);
}

function json_(data) {
  return ContentService
    .createTextOutput(JSON.stringify({rows:data}))
    .setMimeType(ContentService.MimeType.JSON);
}
