// 参謀ダッシュボード — Google Apps Script
// スプレッドシートに貼り付けて「ウェブアプリとしてデプロイ」してください

const SHEET_NAME = 'sanbo_data';

function doGet(e) {
  const action = e.parameter.action;

  if (action === 'load') {
    return loadData();
  }

  if (action === 'gcal') {
    const calId = e.parameter.calId || 'primary';
    return getCalendarEvents(calId);
  }

  return jsonResponse({ error: 'unknown action' });
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    saveData(data);
    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange('A1').setValue('timestamp');
    sheet.getRange('B1').setValue('state');
  }
  return sheet;
}

function saveData(data) {
  const sheet = getSheet();
  const timestamp = new Date().toISOString();
  const stateJson = JSON.stringify(data);
  // 常に2行目に上書き（履歴は1件のみ保持）
  sheet.getRange('A2').setValue(timestamp);
  sheet.getRange('B2').setValue(stateJson);
}

function loadData() {
  const sheet = getSheet();
  const val = sheet.getRange('B2').getValue();
  if (!val) {
    return jsonResponse({ state: null });
  }
  return jsonResponse({ state: val });
}

function getCalendarEvents(calId) {
  try {
    const now = new Date();
    const oneYearLater = new Date();
    oneYearLater.setFullYear(oneYearLater.getFullYear() + 1);

    const cal = CalendarApp.getCalendarById(calId) || CalendarApp.getDefaultCalendar();
    const gcalEvents = cal.getEvents(now, oneYearLater);

    const events = gcalEvents.map(ev => ({
      title: ev.getTitle(),
      date: Utilities.formatDate(ev.getStartTime(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
    }));

    return jsonResponse({ events });
  } catch (err) {
    return jsonResponse({ events: [], error: err.message });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
