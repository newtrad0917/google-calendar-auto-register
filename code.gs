/**
 * Web画面を表示するための関数です。
 * URLを開くと、Index.html が表示されます。
 */

var CUSTOMER_SPREADSHEET_ID = '';
var CUSTOMER_SHEET_NAME = '顧客DB';
var PROJECT_SPREADSHEET_ID = '';
var PROJECT_SHEET_NAME = '案件DB';


/**
 * HTML画面から受け取った予定をGoogleカレンダーへ登録します。
 */
function createCalendarEvent(data) {
  // 入力漏れを確認します。
  if (!data.title) {
    throw new Error('タイトルを入力してください。');
  }

  if (!data.startIso) {
    throw new Error('開始日時を入力してください。');
  }

  // 開始日時を、カレンダーが理解できる日時に変換します。
  const startTime = new Date(data.startIso);

  // 所要時間を分単位で受け取ります。未入力なら60分です。
  const durationMinutes = Number(data.durationMinutes || 60);

  if (durationMinutes <= 0) {
    throw new Error('所要時間は1分以上にしてください。');
  }

  // 終了日時を計算します。
  const endTime = new Date(
    startTime.getTime() + durationMinutes * 60 * 1000
  );

  // 普段使っているGoogleカレンダーを取得します。
  const calendar = CalendarApp.getDefaultCalendar();

  if (data.createMeet === true) {
    try {
      // Apps Scriptのサービスで Calendar API を有効化する必要があります。
      const meetEvent = createCalendarEventWithMeet_(
        calendar,
        data,
        startTime,
        endTime
      );

      return {
        success: true,
        message: 'カレンダーに登録しました。Google Meetも作成しました。',
        meetLink: meetEvent.meetLink || ''
      };
    } catch (error) {
      const fallbackEvent = calendar.createEvent(
        data.title,
        startTime,
        endTime,
        {
          location: data.location || '',
          description: data.memo || ''
        }
      );

      applyCalendarEventColor_(fallbackEvent, data.eventColor);

      return {
        success: true,
        message: '予定は登録しましたが、Google Meet作成は確認が必要です。',
        meetLink: '',
        meetError: error && error.message ? error.message : 'Google Meet作成に失敗しました。',
        eventId: fallbackEvent.getId()
      };
    }
  }

  // カレンダーへ予定を登録します。
  const event = calendar.createEvent(
    data.title,
    startTime,
    endTime,
    {
      location: data.location || '',
      description: data.memo || ''
    }
  );

  applyCalendarEventColor_(event, data.eventColor);

  // 登録できたことをHTML画面へ返します。
  return {
    success: true,
    message: 'カレンダーに登録しました。',
    meetLink: '',
    eventId: event.getId()
  };
}

function createCalendarEventWithMeet_(calendar, data, startTime, endTime) {
  const timeZone = Session.getScriptTimeZone();
  const colorId = getCalendarEventColorId_(data.eventColor);
  const resource = {
    summary: data.title,
    location: data.location || '',
    description: data.memo || '',
    start: {
      dateTime: startTime.toISOString(),
      timeZone: timeZone
    },
    end: {
      dateTime: endTime.toISOString(),
      timeZone: timeZone
    },
    conferenceData: {
      createRequest: {
        requestId: 'sales-ai-meet-' + new Date().getTime(),
        conferenceSolutionKey: {
          type: 'hangoutsMeet'
        }
      }
    }
  };

  if (colorId) {
    resource.colorId = colorId;
  }

  const createdEvent = Calendar.Events.insert(
    resource,
    calendar.getId(),
    {
      conferenceDataVersion: 1
    }
  );

  return {
    event: createdEvent,
    meetLink: getMeetLinkFromEvent_(createdEvent)
  };
}


function applyCalendarEventColor_(event, eventColor) {
  const calendarColor = getCalendarAppEventColor_(eventColor);

  if (!calendarColor) {
    return;
  }

  try {
    event.setColor(calendarColor);
  } catch (error) {
    // 色設定に失敗しても予定登録自体は成功させます。
  }
}

function getCalendarEventColorId_(eventColor) {
  const value = String(eventColor || '').trim().toLowerCase();
  const colorMap = {
    blue: '1',
    green: '2',
    purple: '3',
    red: '4',
    yellow: '5',
    gray: '8',
    grey: '8',
    '1': '1',
    '2': '2',
    '3': '3',
    '4': '4',
    '5': '5',
    '8': '8'
  };

  return colorMap[value] || '';
}

function getCalendarAppEventColor_(eventColor) {
  const colorId = getCalendarEventColorId_(eventColor);

  if (!colorId) {
    return '';
  }

  const eventColorMap = {
    '1': CalendarApp.EventColor.PALE_BLUE,
    '2': CalendarApp.EventColor.PALE_GREEN,
    '3': CalendarApp.EventColor.MAUVE,
    '4': CalendarApp.EventColor.PALE_RED,
    '5': CalendarApp.EventColor.YELLOW,
    '8': CalendarApp.EventColor.GRAY
  };

  return eventColorMap[colorId] || colorId;
}
function getMeetLinkFromEvent_(event) {
  if (!event) {
    return '';
  }

  if (event.hangoutLink) {
    return event.hangoutLink;
  }

  const entryPoints = event.conferenceData && event.conferenceData.entryPoints;

  if (!entryPoints || !entryPoints.length) {
    return '';
  }

  const videoEntry = entryPoints.find(function(entryPoint) {
    return entryPoint.entryPointType === 'video';
  });

  return videoEntry && videoEntry.uri ? videoEntry.uri : '';
}
/**
 * Gmail下書きを保存します。
 */
function saveGmailDraft(data) {

  if (!data.to) {
    throw new Error('宛先を入力してください。');
  }

  if (!data.subject) {
    throw new Error('件名を入力してください。');
  }

  const draftOptions = {};

  if (data.cc) {
    draftOptions.cc = data.cc;
  }

  if (draftOptions.cc) {
    GmailApp.createDraft(
      data.to,
      data.subject,
      data.body || '',
      draftOptions
    );
  } else {
    GmailApp.createDraft(
      data.to,
      data.subject,
      data.body || ''
    );
  }

  return {
    message: 'Gmail下書きを保存しました。'
  };
}
function searchContacts(query) {
var searchText = String(query || '').trim();

if (!searchText) {
throw new Error('名前または会社名を入力してください。');
}

var readMask = 'names,organizations,emailAddresses,phoneNumbers';

People.People.searchContacts({
query: '',
pageSize: 1,
readMask: readMask
});

var response = People.People.searchContacts({
query: searchText,
pageSize: 20,
readMask: readMask
});

var results = response.results || [];
var normalizedQuery = searchText.toLocaleLowerCase();

function getPrimaryValue(items, propertyName) {
if (!items || !items.length) {
return '';
}

var primaryItem = items.filter(function(item) {
  return item.metadata && item.metadata.primary;
})[0] || items[0];

return String(primaryItem[propertyName] || '');
}

return results
.map(function(result) {
var person = result.person || {};

  return {
    name: getPrimaryValue(person.names, 'displayName'),
    company: getPrimaryValue(person.organizations, 'name'),
    email: getPrimaryValue(person.emailAddresses, 'value'),
    phone: getPrimaryValue(person.phoneNumbers, 'value')
  };
})
.filter(function(contact) {
  return contact.name.toLocaleLowerCase().indexOf(normalizedQuery) !== -1 ||
    contact.company.toLocaleLowerCase().indexOf(normalizedQuery) !== -1;
});
}
function testSearchContactsAuth() {
  searchContacts('山田');
}

function generateEmailDraft(data) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');

  if (!apiKey) {
    throw new Error('OPENAI_API_KEYが設定されていません。');
  }

  const prompt = `
あなたは営業メール作成アシスタントです。
以下の情報をもとに、日本語の丁寧な営業メールを作成してください。

【会社名】
${data.company || ''}

【宛先名】
${data.name || ''}

【メール目的】
${data.purpose || ''}

【補足情報】
${data.memo || ''}

条件：
・件名と本文を作成
・本文はそのままGmail下書きに使える形
・過剰に売り込まず、自然で丁寧に
・署名は入れない

出力形式：
件名：
本文：
`;

  const payload = {
    model: 'gpt-5.5',
    input: prompt
  };

  const response = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const json = JSON.parse(response.getContentText());

  if (json.error) {
    throw new Error(json.error.message);
  }

  return json.output_text;
}

function testOpenAIConnection() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');

  if (!apiKey) {
    throw new Error('OPENAI_API_KEYが設定されていません。');
  }

  const payload = {
    model: 'gpt-4.1-mini',
    input: '「接続成功」とだけ返してください。'
  };

  const response = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const text = response.getContentText();
  Logger.log(text);

  const json = JSON.parse(text);

  if (json.error) {
    throw new Error(json.error.message);
  }

  return json.output_text;
}
function generateAiMail(data) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');

  if (!apiKey) {
    throw new Error('OPENAI_API_KEYが設定されていません。');
  }

  if (!data || !data.purpose) {
    throw new Error('メールの目的が入力されていません。');
  }

  const prompt = `
あなたは日本の営業職を支援するメール作成アシスタントです。
以下の情報をもとに、Gmail下書きにそのまま保存できる日本語メールを作成してください。

【メールの目的】
${data.purpose}

【件名欄に入っている情報】
${data.company || ''}

条件：
・丁寧だが、堅すぎない文章
・売り込み感を強くしすぎない
・営業先は病院、福祉施設、工場を想定
・件名と本文を分ける
・署名は入れない

次のJSON形式だけで返してください。
{
  "subject": "ここに件名",
  "body": "ここに本文"
}
`;

  const payload = {
    model: 'gpt-4.1-mini',
    input: prompt
  };

  const response = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const text = response.getContentText();
  Logger.log(text);

  const json = JSON.parse(text);

  if (json.error) {
    throw new Error(json.error.message);
  }

  let outputText = '';

  if (json.output_text) {
    outputText = json.output_text;
  } else if (
    json.output &&
    json.output[0] &&
    json.output[0].content &&
    json.output[0].content[0] &&
    json.output[0].content[0].text
  ) {
    outputText = json.output[0].content[0].text;
  }

  if (!outputText) {
    throw new Error('AIの本文を取得できませんでした。ログを確認してください。');
  }

  outputText = outputText
    .replace(/```json/g, '')
    .replace(/```/g, '')
    .trim();

  try {
    return JSON.parse(outputText);
  } catch (e) {
    return {
      subject: 'AI作成メール',
      body: outputText
    };
  }
}
/**
 * 今日の予定を取得します。
 * 終了時間を過ぎた予定は表示しません。
 */

function getTodayCalendarEvents() {
  const calendar = CalendarApp.getDefaultCalendar();

  const now = new Date();

  const todayStart = new Date();
  todayStart.setHours(0, 0, 0, 0);

  const todayEnd = new Date();
  todayEnd.setHours(23, 59, 59, 999);

  const events = calendar.getEvents(todayStart, todayEnd);

  return events
    .filter(function(event) {
      return event.getEndTime().getTime() > now.getTime();
    })
    .sort(function(a, b) {
      return a.getStartTime().getTime() - b.getStartTime().getTime();
    })
    .slice(0, 3)
    .map(function(event) {
      const startTime = event.getStartTime();
      const endTime = event.getEndTime();

      const isCurrent =
        startTime.getTime() <= now.getTime() &&
        endTime.getTime() > now.getTime();

      return {
        title: event.getTitle(),
        startTime: Utilities.formatDate(
          startTime,
          Session.getScriptTimeZone(),
          'HH:mm'
        ),
        endTime: Utilities.formatDate(
          endTime,
          Session.getScriptTimeZone(),
          'HH:mm'
        ),
        location: event.getLocation() || '',
        status: isCurrent ? 'current' : 'next'
      };
    });
}
function getCalendarEventsByRange(startDate, endDate) {
  const calendar = CalendarApp.getDefaultCalendar();
  const rangeStart = parseCalendarRangeDate_(startDate);
  const rangeEnd = parseCalendarRangeDate_(endDate);

  if (!rangeStart || !rangeEnd || rangeEnd.getTime() <= rangeStart.getTime()) {
    throw new Error('予定取得範囲の日付を確認してください。');
  }

  const now = new Date();
  const timeZone = Session.getScriptTimeZone();
  const events = calendar.getEvents(rangeStart, rangeEnd);

  return events
    .sort(function(a, b) {
      return a.getStartTime().getTime() - b.getStartTime().getTime();
    })
    .map(function(event) {
      return mapCalendarEventForList_(event, timeZone, now);
    });
}

function parseCalendarRangeDate_(value) {
  const text = String(value || '').trim();

  if (!text) {
    return null;
  }

  const dateMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);

  if (dateMatch) {
    return new Date(
      Number(dateMatch[1]),
      Number(dateMatch[2]) - 1,
      Number(dateMatch[3]),
      0,
      0,
      0,
      0
    );
  }

  const parsed = new Date(text);

  if (isNaN(parsed.getTime())) {
    return null;
  }

  return parsed;
}

function mapCalendarEventForList_(event, timeZone, now) {
  const startTime = event.getStartTime();
  const endTime = event.getEndTime();
  const isCurrent =
    startTime.getTime() <= now.getTime() &&
    endTime.getTime() > now.getTime();

  return {
    id: event.getId(),
    title: event.getTitle(),
    start: startTime.toISOString(),
    end: endTime.toISOString(),
    startTime: Utilities.formatDate(startTime, timeZone, 'HH:mm'),
    endTime: Utilities.formatDate(endTime, timeZone, 'HH:mm'),
    dateKey: Utilities.formatDate(startTime, timeZone, 'yyyy-MM-dd'),
    location: event.getLocation() || '',
    description: event.getDescription() || '',
    status: isCurrent ? 'current' : 'next',
    allDay: event.isAllDayEvent()
  };
}

/**
 * Tasksシートへ新しいTODOを追加します。
 */
function addTask(data) {
  if (!data) {
    throw new Error('TODOデータがありません。');
  }

  var title = String(data.title || '').trim();

  if (!title) {
    throw new Error('タスク名を入力してください。');
  }

  var now = new Date();
  var timeZone = Session.getScriptTimeZone();
  var id = 'T' + Utilities.formatDate(now, timeZone, 'yyyyMMddHHmmssSSS');
  var spreadsheet = getTaskSpreadsheet_();
  var sheet = spreadsheet.getSheetByName('Tasks');

  if (!sheet) {
    throw new Error('Tasksシートが見つかりません。');
  }

  sheet.appendRow([
    id,
    title,
    String(data.category || '').trim(),
    String(data.dueDate || '').trim(),
    String(data.priority || '').trim(),
    'Open',
    String(data.memo || '').trim(),
    now,
    '',
    'Manual'
  ]);

  return {
    message: 'TODOを追加しました。'
  };
}

/**
 * GASエディタ上で手動実行するTODO保存先の初期設定です。
 * 独立GASプロジェクトでも動くよう、営業AI秘書DBのIDを直接Script Propertiesへ保存します。
 */
function setupTaskSpreadsheet() {
  var spreadsheetId = '1mf8ewYcmH5o4K_ZSS97u038Z_PdRacUAhVcXDIDxwbk';
  var spreadsheet;

  try {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } catch (error) {
    throw new Error('営業AI秘書DBを開けませんでした。スプレッドシートIDまたは権限を確認してください。詳細: ' + error.message);
  }

  if (!spreadsheet.getSheetByName('Tasks')) {
    throw new Error('営業AI秘書DBにTasksシートが見つかりません。');
  }

  PropertiesService
    .getScriptProperties()
    .setProperty('TASK_SPREADSHEET_ID', spreadsheetId);

  Logger.log('TASK_SPREADSHEET_IDを保存しました。');

  return {
    message: 'TASK_SPREADSHEET_IDを保存しました。'
  };
}

function setupCustomerDbSheet() {
  var spreadsheet = getTaskSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(CUSTOMER_SHEET_NAME);
  var wasCreated = false;
  var insertedSample = false;
  var headers = getCustomerDbHeaders_();

  if (!sheet) {
    sheet = spreadsheet.insertSheet(CUSTOMER_SHEET_NAME);
    wasCreated = true;
  }

  var lastRow = sheet.getLastRow();
  var headerResult = ensureCustomerDbHeaders_(sheet, headers);
  var currentHeaders = getCustomerDbHeaderValues_(sheet);

  if (sheet.getLastRow() < 2 && currentHeaders.length) {
    sheet
      .getRange(2, 1, 1, currentHeaders.length)
      .setValues([buildCustomerDbSampleRowForHeaders_(currentHeaders)]);
    insertedSample = true;
  }

  styleCustomerDbSheet_(sheet, headers);

  PropertiesService
    .getScriptProperties()
    .setProperty('CUSTOMER_SPREADSHEET_ID', spreadsheet.getId());

  var result = {
    message: CUSTOMER_SHEET_NAME + 'シートの初期化が完了しました。',
    spreadsheetId: spreadsheet.getId(),
    sheetName: CUSTOMER_SHEET_NAME,
    wasCreated: wasCreated,
    wroteHeaders: headerResult.wroteHeaders,
    addedHeaders: headerResult.addedHeaders,
    renamedPhoneToOfficePhone: headerResult.renamedPhoneToOfficePhone,
    insertedSample: insertedSample,
    existingRows: lastRow
  };

  Logger.log(JSON.stringify(result, null, 2));

  return result;
}

function getCustomerDbHeaders_() {
  return [
    'id',
    'companyName',
    'corporationName',
    'contactName',
    'department',
    'industry',
    'prefecture',
    'city',
    'address',
    'officePhone',
    'mobilePhone',
    'email',
    'ccEmail',
    'googleMapUrl',
    'lastVisit',
    'nextVisit',
    'memo',
    'projectCount',
    'hotLevel',
    'aiSummary',
    'createdAt',
    'updatedAt',
    'notes'
  ];
}

function getCustomerDbSampleRow_() {
  return [
    'customer_001',
    'テスト病院',
    '',
    '山田 太郎',
    '事務長',
    '病院',
    '福岡県',
    '朝倉市',
    '福岡県朝倉市',
    '092-000-0000',
    '090-0000-0000',
    'test@example.com',
    '',
    'https://maps.google.com',
    '2026-07-01',
    '2026-07-10',
    'LED提案中',
    1,
    'WARM',
    '初回訪問予定',
    '2026-07-02',
    '2026-07-02',
    ''
  ];
}

function getCustomerDbHeaderValues_(sheet) {
  var lastColumn = sheet.getLastColumn();

  if (!lastColumn) {
    return [];
  }

  return sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map(function(header) {
      return String(header || '').trim();
    });
}

function buildCustomerDbSampleRowForHeaders_(headers) {
  var defaultHeaders = getCustomerDbHeaders_();
  var defaultRow = getCustomerDbSampleRow_();
  var sampleByHeader = {};

  defaultHeaders.forEach(function(header, index) {
    sampleByHeader[header] = defaultRow[index] || '';
  });

  return headers.map(function(header) {
    return Object.prototype.hasOwnProperty.call(sampleByHeader, header)
      ? sampleByHeader[header]
      : '';
  });
}

function ensureCustomerDbHeaders_(sheet, desiredHeaders) {
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var headerValues = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var hasExistingHeaders = headerValues.some(function(value) {
    return String(value || '').trim() !== '';
  });
  var result = {
    wroteHeaders: false,
    addedHeaders: [],
    renamedPhoneToOfficePhone: false
  };

  if (!hasExistingHeaders) {
    sheet.getRange(1, 1, 1, desiredHeaders.length).setValues([desiredHeaders]);
    result.wroteHeaders = true;
    return result;
  }

  var phoneIndex = headerValues.indexOf('phone');

  if (phoneIndex !== -1 && headerValues.indexOf('officePhone') === -1) {
    sheet.getRange(1, phoneIndex + 1).setValue('officePhone');
    headerValues[phoneIndex] = 'officePhone';
    result.renamedPhoneToOfficePhone = true;
  }

  desiredHeaders.forEach(function(header) {
    if (headerValues.indexOf(header) === -1) {
      headerValues.push(header);
      result.addedHeaders.push(header);
    }
  });

  sheet.getRange(1, 1, 1, headerValues.length).setValues([headerValues]);
  result.wroteHeaders = result.addedHeaders.length > 0 || result.renamedPhoneToOfficePhone;
  return result;
}

function styleCustomerDbSheet_(sheet, desiredHeaders) {
  var lastColumn = Math.max(sheet.getLastColumn(), desiredHeaders.length);

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, lastColumn).setFontWeight('bold');

  if (!sheet.getFilter()) {
    sheet
      .getRange(1, 1, Math.max(sheet.getLastRow(), 1), lastColumn)
      .createFilter();
  }

  sheet.setColumnWidths(1, lastColumn, 130);
  setCustomerDbColumnWidth_(sheet, 'companyName', 180);
  setCustomerDbColumnWidth_(sheet, 'corporationName', 180);
  setCustomerDbColumnWidth_(sheet, 'contactName', 140);
  setCustomerDbColumnWidth_(sheet, 'officePhone', 150);
  setCustomerDbColumnWidth_(sheet, 'mobilePhone', 150);
  setCustomerDbColumnWidth_(sheet, 'email', 190);
  setCustomerDbColumnWidth_(sheet, 'ccEmail', 190);
  setCustomerDbColumnWidth_(sheet, 'address', 220);
  setCustomerDbColumnWidth_(sheet, 'googleMapUrl', 220);
  setCustomerDbColumnWidth_(sheet, 'memo', 220);
  setCustomerDbColumnWidth_(sheet, 'aiSummary', 240);
  setCustomerDbColumnWidth_(sheet, 'notes', 220);
}

function setCustomerDbColumnWidth_(sheet, headerName, width) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var index = headers.indexOf(headerName);

  if (index !== -1) {
    sheet.setColumnWidth(index + 1, width);
  }
}

/**
 * TODO保存先スプレッドシートはScript PropertiesのTASK_SPREADSHEET_IDで管理します。
 * 将来、顧客DB、案件DB、活動履歴DB、AI提案DBも同じ方式でID管理する予定です。
 * 想定キー: CUSTOMER_SPREADSHEET_ID, PROJECT_SPREADSHEET_ID,
 * ACTIVITY_LOG_SPREADSHEET_ID, AI_PROPOSAL_SPREADSHEET_ID
 * DriveAppでファイル名検索しないことで、権限エラーや同名ファイル問題を避けます。
 */
function setupProjectDbSheet() {
  var spreadsheet = getTaskSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(PROJECT_SHEET_NAME);
  var wasCreated = false;
  var headers = getProjectDbHeaders_();

  if (!sheet) {
    sheet = spreadsheet.insertSheet(PROJECT_SHEET_NAME);
    wasCreated = true;
  }

  var lastRow = sheet.getLastRow();
  var headerResult = ensureProjectDbHeaders_(sheet, headers);

  styleProjectDbSheet_(sheet, headers);

  PropertiesService
    .getScriptProperties()
    .setProperty('PROJECT_SPREADSHEET_ID', spreadsheet.getId());

  var result = {
    message: PROJECT_SHEET_NAME + 'シートの初期化が完了しました。',
    spreadsheetId: spreadsheet.getId(),
    sheetName: PROJECT_SHEET_NAME,
    wasCreated: wasCreated,
    wroteHeaders: headerResult.wroteHeaders,
    addedHeaders: headerResult.addedHeaders,
    existingRows: lastRow
  };

  Logger.log(JSON.stringify(result, null, 2));

  return result;
}

function getProjectDbHeaders_() {
  return [
    'id',
    'customerId',
    'projectName',
    'category',
    'status',
    'priority',
    'proposalAmount',
    'costAmount',
    'profitAmount',
    'expectedOrderDate',
    'expectedWorkDate',
    'expectedPaymentDate',
    'nextAction',
    'memo',
    'createdAt',
    'updatedAt',
    'notes'
  ];
}

function getProjectDbHeaderValues_(sheet) {
  var lastColumn = sheet.getLastColumn();

  if (!lastColumn) {
    return [];
  }

  return sheet
    .getRange(1, 1, 1, lastColumn)
    .getValues()[0]
    .map(function(header) {
      return String(header || '').trim();
    });
}

function ensureProjectDbHeaders_(sheet, desiredHeaders) {
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var headerValues = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var hasExistingHeaders = headerValues.some(function(value) {
    return String(value || '').trim() !== '';
  });
  var result = {
    wroteHeaders: false,
    addedHeaders: []
  };

  if (!hasExistingHeaders) {
    sheet.getRange(1, 1, 1, desiredHeaders.length).setValues([desiredHeaders]);
    result.wroteHeaders = true;
    return result;
  }

  desiredHeaders.forEach(function(header) {
    if (headerValues.indexOf(header) === -1) {
      headerValues.push(header);
      result.addedHeaders.push(header);
    }
  });

  sheet.getRange(1, 1, 1, headerValues.length).setValues([headerValues]);
  result.wroteHeaders = result.addedHeaders.length > 0;
  return result;
}

function styleProjectDbSheet_(sheet, desiredHeaders) {
  var lastColumn = Math.max(sheet.getLastColumn(), desiredHeaders.length);

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, lastColumn).setFontWeight('bold');

  if (!sheet.getFilter()) {
    sheet
      .getRange(1, 1, Math.max(sheet.getLastRow(), 1), lastColumn)
      .createFilter();
  }

  sheet.setColumnWidths(1, lastColumn, 130);
  setProjectDbColumnWidth_(sheet, 'customerId', 140);
  setProjectDbColumnWidth_(sheet, 'projectName', 220);
  setProjectDbColumnWidth_(sheet, 'category', 130);
  setProjectDbColumnWidth_(sheet, 'status', 130);
  setProjectDbColumnWidth_(sheet, 'proposalAmount', 150);
  setProjectDbColumnWidth_(sheet, 'costAmount', 140);
  setProjectDbColumnWidth_(sheet, 'profitAmount', 140);
  setProjectDbColumnWidth_(sheet, 'expectedOrderDate', 160);
  setProjectDbColumnWidth_(sheet, 'expectedWorkDate', 160);
  setProjectDbColumnWidth_(sheet, 'expectedPaymentDate', 170);
  setProjectDbColumnWidth_(sheet, 'nextAction', 220);
  setProjectDbColumnWidth_(sheet, 'memo', 220);
  setProjectDbColumnWidth_(sheet, 'notes', 220);
}

function setProjectDbColumnWidth_(sheet, headerName, width) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var index = headers.indexOf(headerName);

  if (index !== -1) {
    sheet.setColumnWidth(index + 1, width);
  }
}

function getTaskSpreadsheet_() {
  var spreadsheetId = PropertiesService
    .getScriptProperties()
    .getProperty('TASK_SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error('TASK_SPREADSHEET_IDが未設定です。setupTaskSpreadsheet()を実行してください。');
  }

  return SpreadsheetApp.openById(spreadsheetId);
}

/**
 * 顧客DBシートからCustomer Hub用の顧客一覧を取得します。
 */
function getCustomers() {
  var sheet = getCustomerSheet_();
  var values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return [];
  }

  var headers = values[0].map(function(header) {
    return String(header || '').trim();
  });

  return values
    .slice(1)
    .map(function(row) {
      return normalizeCustomerRow_(row, headers);
    })
    .filter(function(customer) {
      return customer.id && customer.companyName;
    });
}

function getCustomer(customerId) {
  var id = String(customerId || '').trim();

  if (!id) {
    throw new Error('顧客IDがありません。');
  }

  var customer = getCustomers().filter(function(item) {
    return item.id === id;
  })[0];

  if (!customer) {
    throw new Error('顧客が見つかりません。');
  }

  return customer;
}

function addCustomer(customer) {
  var data = customer || {};
  var companyName = String(data.companyName || '').trim();

  if (!companyName) {
    throw new Error('顧客名を入力してください。');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var sheet = getCustomerSheet_();
    ensureCustomerDbHeaders_(sheet, getCustomerDbHeaders_());
    var headers = getCustomerDbHeaderValues_(sheet);
    var now = new Date();
    var id = String(data.id || '').trim() || generateCustomerId_();
    var row = buildCustomerRowForHeaders_(Object.assign({}, data, {
      id: id,
      companyName: companyName,
      createdAt: now,
      updatedAt: now
    }), headers);

    sheet.appendRow(row);

    return {
      id: id,
      message: '顧客を追加しました。',
      customer: getCustomer(id),
      customers: getCustomers()
    };
  } finally {
    lock.releaseLock();
  }
}

function updateCustomer(customer) {
  var data = customer || {};
  var id = String(data.id || '').trim();
  var companyName = String(data.companyName || '').trim();

  if (!id) {
    throw new Error('顧客IDがありません。');
  }

  if (!companyName) {
    throw new Error('顧客名を入力してください。');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var sheet = getCustomerSheet_();
    ensureCustomerDbHeaders_(sheet, getCustomerDbHeaders_());
    var headers = getCustomerDbHeaderValues_(sheet);
    var rowNumber = findCustomerRowById_(sheet, id);

    if (!rowNumber) {
      throw new Error('更新対象の顧客が見つかりません。');
    }

    var existingRow = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    var createdAtIndex = headers.indexOf('createdAt');
    var row = buildCustomerRowForHeaders_(Object.assign({}, data, {
      id: id,
      companyName: companyName,
      createdAt: createdAtIndex === -1 ? '' : existingRow[createdAtIndex],
      updatedAt: new Date()
    }), headers, existingRow);

    sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);

    return {
      id: id,
      message: '顧客情報を更新しました。',
      customer: getCustomer(id),
      customers: getCustomers()
    };
  } finally {
    lock.releaseLock();
  }
}

function deleteCustomer(customerId) {
  var id = String(customerId || '').trim();

  if (!id) {
    throw new Error('顧客IDがありません。');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var sheet = getCustomerSheet_();
    var rowNumber = findCustomerRowById_(sheet, id);

    if (!rowNumber) {
      throw new Error('削除対象の顧客が見つかりません。');
    }

    deleteCustomerRow_(sheet, rowNumber);

    return {
      id: id,
      message: '顧客を削除しました。',
      customers: getCustomers()
    };
  } finally {
    lock.releaseLock();
  }
}

function generateCustomerId_() {
  var sheet = getCustomerSheet_();
  var headers = getCustomerDbHeaderValues_(sheet);
  var idColumn = headers.indexOf('id');
  var lastRow = sheet.getLastRow();
  var maxNumber = 0;

  if (idColumn === -1 || lastRow < 2) {
    return 'C000001';
  }

  sheet
    .getRange(2, idColumn + 1, lastRow - 1, 1)
    .getValues()
    .forEach(function(row) {
      var id = String(row[0] || '').trim();
      var match = id.match(/(\d+)$/);

      if (match) {
        maxNumber = Math.max(maxNumber, Number(match[1]));
      }
    });

  return 'C' + ('000000' + (maxNumber + 1)).slice(-6);
}

function findCustomerRowById_(sheet, customerId) {
  var id = String(customerId || '').trim();
  var headers = getCustomerDbHeaderValues_(sheet);
  var idColumn = headers.indexOf('id');
  var lastRow = sheet.getLastRow();

  if (!id || idColumn === -1 || lastRow < 2) {
    return 0;
  }

  var values = sheet.getRange(2, idColumn + 1, lastRow - 1, 1).getValues();

  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === id) {
      return i + 2;
    }
  }

  return 0;
}

function buildCustomerRowForHeaders_(customer, headers, existingRow) {
  var data = customer || {};
  var existing = existingRow || [];
  var knownFields = getCustomerDbHeaders_().reduce(function(result, header) {
    result[header] = true;
    return result;
  }, {});

  return headers.map(function(header, index) {
    if (Object.prototype.hasOwnProperty.call(data, header)) {
      return data[header];
    }

    if (!knownFields[header]) {
      return existing[index] === undefined || existing[index] === null
        ? ''
        : existing[index];
    }

    return '';
  });
}

function deleteCustomerRow_(sheet, rowNumber) {
  sheet.deleteRow(rowNumber);
}

function getProjects() {
  var sheet = getProjectSheet_();
  var values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return [];
  }

  var headers = values[0].map(function(header) {
    return String(header || '').trim();
  });

  return values
    .slice(1)
    .map(function(row) {
      return normalizeProjectRow_(row, headers);
    })
    .filter(function(project) {
      return project.id && project.customerId && project.projectName;
    });
}

function getProjectsByCustomerId(customerId) {
  var id = String(customerId || '').trim();

  if (!id) {
    return [];
  }

  return getProjects().filter(function(project) {
    return project.customerId === id;
  });
}

function addProject(project) {
  var data = project || {};
  var customerId = String(data.customerId || '').trim();
  var projectName = String(data.projectName || '').trim();

  if (!customerId) {
    throw new Error('顧客IDがありません。');
  }

  if (!projectName) {
    throw new Error('案件名を入力してください。');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var sheet = getProjectSheet_();
    ensureProjectDbHeaders_(sheet, getProjectDbHeaders_());
    var headers = getProjectDbHeaderValues_(sheet);
    var now = new Date();
    var id = String(data.id || '').trim() || generateProjectId_();
    var row = buildProjectRowForHeaders_(Object.assign({}, data, {
      id: id,
      customerId: customerId,
      projectName: projectName,
      createdAt: now,
      updatedAt: now
    }), headers);

    sheet.appendRow(row);

    return {
      id: id,
      message: '案件を追加しました。',
      project: getProjectById_(id),
      projects: getProjects()
    };
  } finally {
    lock.releaseLock();
  }
}

function updateProject(project) {
  var data = project || {};
  var id = String(data.id || '').trim();
  var customerId = String(data.customerId || '').trim();
  var projectName = String(data.projectName || '').trim();

  if (!id) {
    throw new Error('案件IDがありません。');
  }

  if (!customerId) {
    throw new Error('顧客IDがありません。');
  }

  if (!projectName) {
    throw new Error('案件名を入力してください。');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var sheet = getProjectSheet_();
    ensureProjectDbHeaders_(sheet, getProjectDbHeaders_());
    var headers = getProjectDbHeaderValues_(sheet);
    var rowNumber = findProjectRowById_(sheet, id);

    if (!rowNumber) {
      throw new Error('更新対象の案件が見つかりません。');
    }

    var existingRow = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
    var createdAtIndex = headers.indexOf('createdAt');
    var row = buildProjectRowForHeaders_(Object.assign({}, data, {
      id: id,
      customerId: customerId,
      projectName: projectName,
      createdAt: createdAtIndex === -1 ? '' : existingRow[createdAtIndex],
      updatedAt: new Date()
    }), headers, existingRow);

    sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);

    return {
      id: id,
      message: '案件情報を更新しました。',
      project: getProjectById_(id),
      projects: getProjects()
    };
  } finally {
    lock.releaseLock();
  }
}

function deleteProject(projectId) {
  var id = String(projectId || '').trim();

  if (!id) {
    throw new Error('案件IDがありません。');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var sheet = getProjectSheet_();
    var rowNumber = findProjectRowById_(sheet, id);

    if (!rowNumber) {
      throw new Error('削除対象の案件が見つかりません。');
    }

    sheet.deleteRow(rowNumber);

    return {
      id: id,
      message: '案件を削除しました。',
      projects: getProjects()
    };
  } finally {
    lock.releaseLock();
  }
}

function generateProjectId_() {
  var sheet = getProjectSheet_();
  var headers = getProjectDbHeaderValues_(sheet);
  var idColumn = headers.indexOf('id');
  var lastRow = sheet.getLastRow();
  var maxNumber = 0;

  if (idColumn === -1 || lastRow < 2) {
    return 'P000001';
  }

  sheet
    .getRange(2, idColumn + 1, lastRow - 1, 1)
    .getValues()
    .forEach(function(row) {
      var id = String(row[0] || '').trim();
      var match = id.match(/(\d+)$/);

      if (match) {
        maxNumber = Math.max(maxNumber, Number(match[1]));
      }
    });

  return 'P' + ('000000' + (maxNumber + 1)).slice(-6);
}

function findProjectRowById_(sheet, projectId) {
  var id = String(projectId || '').trim();
  var headers = getProjectDbHeaderValues_(sheet);
  var idColumn = headers.indexOf('id');
  var lastRow = sheet.getLastRow();

  if (!id || idColumn === -1 || lastRow < 2) {
    return 0;
  }

  var values = sheet.getRange(2, idColumn + 1, lastRow - 1, 1).getValues();

  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === id) {
      return i + 2;
    }
  }

  return 0;
}

function buildProjectRowForHeaders_(project, headers, existingRow) {
  var data = project || {};
  var existing = existingRow || [];
  var knownFields = getProjectDbHeaders_().reduce(function(result, header) {
    result[header] = true;
    return result;
  }, {});

  return headers.map(function(header, index) {
    if (Object.prototype.hasOwnProperty.call(data, header)) {
      return data[header];
    }

    if (!knownFields[header]) {
      return existing[index] === undefined || existing[index] === null
        ? ''
        : existing[index];
    }

    return '';
  });
}

function getProjectById_(projectId) {
  var id = String(projectId || '').trim();

  return getProjects().filter(function(project) {
    return project.id === id;
  })[0] || null;
}

function normalizeProjectRow_(row, headers) {
  function value(name) {
    var index = headers.indexOf(name);
    return index === -1 ? '' : row[index];
  }

  function text(name) {
    var rawValue = value(name);

    if (rawValue === null || rawValue === undefined || rawValue === '') {
      return '';
    }

    return String(rawValue).trim();
  }

  function dateText(name) {
    var rawValue = value(name);

    if (Object.prototype.toString.call(rawValue) === '[object Date]' && !isNaN(rawValue.getTime())) {
      return Utilities.formatDate(rawValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }

    if (rawValue === null || rawValue === undefined || rawValue === '') {
      return '';
    }

    return String(rawValue).trim();
  }

  return {
    id: text('id'),
    customerId: text('customerId'),
    projectName: text('projectName'),
    category: text('category'),
    status: text('status'),
    priority: text('priority'),
    proposalAmount: text('proposalAmount'),
    costAmount: text('costAmount'),
    profitAmount: text('profitAmount'),
    expectedOrderDate: dateText('expectedOrderDate'),
    expectedWorkDate: dateText('expectedWorkDate'),
    expectedPaymentDate: dateText('expectedPaymentDate'),
    nextAction: text('nextAction'),
    memo: text('memo'),
    createdAt: dateText('createdAt'),
    updatedAt: dateText('updatedAt'),
    notes: text('notes')
  };
}

function getProjectSheet_() {
  var properties = PropertiesService.getScriptProperties();
  var spreadsheetId = PROJECT_SPREADSHEET_ID ||
    properties.getProperty('PROJECT_SPREADSHEET_ID') ||
    properties.getProperty('TASK_SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error('TASK_SPREADSHEET_IDが未設定です。setupTaskSpreadsheet()を実行してからsetupProjectDbSheet()を実行してください。');
  }

  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(PROJECT_SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(PROJECT_SHEET_NAME);
    PropertiesService
      .getScriptProperties()
      .setProperty('PROJECT_SPREADSHEET_ID', spreadsheet.getId());
  }

  if (!sheet) {
    throw new Error(PROJECT_SHEET_NAME + 'シートが見つかりません。');
  }

  ensureProjectDbHeaders_(sheet, getProjectDbHeaders_());
  styleProjectDbSheet_(sheet, getProjectDbHeaders_());

  return sheet;
}

function normalizeCustomerRow_(row, headers) {
  function value(name) {
    var index = headers.indexOf(name);
    return index === -1 ? '' : row[index];
  }

  function text(name) {
    var rawValue = value(name);

    if (rawValue === null || rawValue === undefined || rawValue === '') {
      return '';
    }

    return String(rawValue).trim();
  }

  function dateText(name) {
    var rawValue = value(name);

    if (Object.prototype.toString.call(rawValue) === '[object Date]' && !isNaN(rawValue.getTime())) {
      return Utilities.formatDate(rawValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }

    if (rawValue === null || rawValue === undefined || rawValue === '') {
      return '';
    }

    return String(rawValue).trim();
  }

  var address = text('address');
  var companyName = text('companyName');
  var googleMapUrl = text('googleMapUrl');
  var nextAction = text('nextAction') || dateText('nextVisit');
  var memo = text('memo');
  var aiSummary = text('aiSummary');
  var legacyPhone = text('phone');
  var officePhone = text('officePhone') || legacyPhone;
  var mobilePhone = text('mobilePhone');

  return {
    id: text('id'),
    companyName: companyName,
    corporationName: text('corporationName'),
    contactName: text('contactName'),
    personName: text('contactName'),
    title: '',
    department: text('department'),
    industry: text('industry'),
    prefecture: text('prefecture'),
    city: text('city'),
    address: address,
    officePhone: officePhone,
    mobilePhone: mobilePhone,
    phone: mobilePhone || officePhone || legacyPhone,
    email: text('email'),
    cc: text('ccEmail'),
    ccEmail: text('ccEmail'),
    googleMapUrl: googleMapUrl,
    memo: memo,
    salesMemo: aiSummary || memo,
    lastContactDate: dateText('lastContactDate') || dateText('lastVisit'),
    lastVisit: dateText('lastVisit'),
    nextVisit: dateText('nextVisit'),
    nextAction: nextAction,
    nextSchedule: nextAction || '未設定',
    projectCount: text('projectCount'),
    hotLevel: text('hotLevel'),
    aiSummary: aiSummary,
    notes: text('notes'),
    status: text('status'),
    createdAt: dateText('createdAt'),
    updatedAt: dateText('updatedAt')
  };
}

function getCustomerSheet_() {
  var properties = PropertiesService.getScriptProperties();
  var spreadsheetId = CUSTOMER_SPREADSHEET_ID ||
    properties.getProperty('CUSTOMER_SPREADSHEET_ID') ||
    properties.getProperty('TASK_SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error('TASK_SPREADSHEET_IDが未設定です。setupTaskSpreadsheet()を実行してからsetupCustomerDbSheet()を実行してください。');
  }

  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(CUSTOMER_SHEET_NAME);

  if (!sheet) {
    throw new Error(CUSTOMER_SHEET_NAME + 'シートが見つかりません。');
  }

  return sheet;
}

/**
 * ホーム画面のAIステータスカードに表示する軽量な状態を返します。
 * 既存機能には影響しないよう、各項目を個別に確認します。
 */
function getAiSystemStatus() {
  var items = [
    { key: 'ai', label: 'AIエンジン', state: 'waiting', text: '待機中' },
    { key: 'calendar', label: 'カレンダー', state: 'ok', text: '接続済み' },
    { key: 'todo', label: 'TODO', state: 'ok', text: '同期済み' },
    { key: 'database', label: 'Database', state: 'ok', text: '正常' },
    { key: 'gmail', label: 'Gmail', state: 'muted', text: '準備中' },
    { key: 'contacts', label: 'Contacts', state: 'muted', text: '準備中' }
  ];

  function markError(key) {
    items.forEach(function(item) {
      if (item.key === key) {
        item.state = 'error';
        item.text = '確認必要';
      }
    });
  }

  try {
    var spreadsheet = null;

    try {
      CalendarApp.getDefaultCalendar();
    } catch (error) {
      markError('calendar');
    }

    try {
      spreadsheet = getTaskSpreadsheet_();
    } catch (error) {
      markError('database');
      markError('todo');
    }

    if (spreadsheet) {
      try {
        if (!spreadsheet.getSheetByName('Tasks')) {
          throw new Error('Tasksシートが見つかりません。');
        }
      } catch (error) {
        markError('todo');
      }
    }

    return {
      ok: true,
      lastSync: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm'),
      items: items
    };
  } catch (error) {
    return {
      ok: false,
      lastSync: '取得失敗',
      items: [
        { key: 'ai', label: 'AIエンジン', state: 'waiting', text: '待機中' },
        { key: 'calendar', label: 'カレンダー', state: 'error', text: '確認必要' },
        { key: 'todo', label: 'TODO', state: 'error', text: '確認必要' },
        { key: 'database', label: 'Database', state: 'error', text: '確認必要' },
        { key: 'gmail', label: 'Gmail', state: 'muted', text: '準備中' },
        { key: 'contacts', label: 'Contacts', state: 'muted', text: '準備中' }
      ]
    };
  }
}

/**
 * Tasksシートから未完了TODOを取得します。
 */
function getTasks(category) {
  var sheet = getTaskSpreadsheet_().getSheetByName('Tasks');

  if (!sheet) {
    throw new Error('Tasksシートが見つかりません。');
  }

  var values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return [];
  }

  var headers = values[0];
  var columns = getTaskColumns_(headers);
  var filterCategory = String(category || 'すべて').trim();
  var timeZone = Session.getScriptTimeZone();

  return values
    .slice(1)
    .filter(function(row) {
      var status = String(row[columns.Status] || '').trim();
      var rowCategory = String(row[columns.Category] || '').trim();

      if (status !== 'Open') {
        return false;
      }

      if (filterCategory === '仕事' || filterCategory === 'プライベート') {
        return rowCategory === filterCategory;
      }

      return true;
    })
    .map(function(row) {
      return {
        id: String(row[columns.ID] || '').trim(),
        title: String(row[columns.Title] || '').trim(),
        category: String(row[columns.Category] || '').trim(),
        dueDate: formatTaskDate_(row[columns.DueDate], timeZone),
        priority: String(row[columns.Priority] || '').trim(),
        memo: String(row[columns.Memo] || '').trim()
      };
    })
    .sort(function(a, b) {
      return getTaskDueTime_(a.dueDate) - getTaskDueTime_(b.dueDate);
    });
}

/**
 * 指定されたTODOを完了にします。
 */
function completeTask(taskId) {
  var id = String(taskId || '').trim();

  if (!id) {
    throw new Error('TODOのIDがありません。');
  }

  var sheet = getTaskSpreadsheet_().getSheetByName('Tasks');

  if (!sheet) {
    throw new Error('Tasksシートが見つかりません。');
  }

  var values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    throw new Error('TODOが見つかりません。');
  }

  var columns = getTaskColumns_(values[0]);

  for (var index = 1; index < values.length; index++) {
    if (String(values[index][columns.ID] || '').trim() === id) {
      var rowNumber = index + 1;
      sheet.getRange(rowNumber, columns.Status + 1).setValue('Done');
      sheet.getRange(rowNumber, columns.CompletedAt + 1).setValue(new Date());

      return {
        message: 'TODOを完了しました。'
      };
    }
  }

  throw new Error('TODOが見つかりません。');
}

/**
 * 指定されたTODOを編集します。
 */
function updateTask(data) {
  if (!data) {
    throw new Error('TODOデータがありません。');
  }

  var id = String(data.id || '').trim();
  var title = String(data.title || '').trim();

  if (!id) {
    throw new Error('TODOのIDがありません。');
  }

  if (!title) {
    throw new Error('タスク名を入力してください。');
  }

  var sheet = getTaskSpreadsheet_().getSheetByName('Tasks');

  if (!sheet) {
    throw new Error('Tasksシートが見つかりません。');
  }

  var values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    throw new Error('TODOが見つかりません。');
  }

  var columns = getTaskColumns_(values[0]);

  for (var index = 1; index < values.length; index++) {
    if (String(values[index][columns.ID] || '').trim() === id) {
      var rowNumber = index + 1;

      sheet.getRange(rowNumber, columns.Title + 1).setValue(title);
      sheet.getRange(rowNumber, columns.Category + 1).setValue(String(data.category || '').trim());
      sheet.getRange(rowNumber, columns.DueDate + 1).setValue(String(data.dueDate || '').trim());
      sheet.getRange(rowNumber, columns.Priority + 1).setValue(String(data.priority || '').trim());
      sheet.getRange(rowNumber, columns.Memo + 1).setValue(String(data.memo || '').trim());

      return {
        message: 'TODOを更新しました。'
      };
    }
  }

  throw new Error('TODOが見つかりません。');
}

/**
 * Tasksシートの列番号をヘッダー名から取得します。
 */
function getTaskColumns_(headers) {
  var names = [
    'ID',
    'Title',
    'Category',
    'DueDate',
    'Priority',
    'Status',
    'Memo',
    'CreatedAt',
    'CompletedAt',
    'Source'
  ];
  var columns = {};

  names.forEach(function(name) {
    var index = headers.indexOf(name);

    if (index === -1) {
      throw new Error('Tasksシートに' + name + '列が見つかりません。');
    }

    columns[name] = index;
  });

  return columns;
}

/**
 * TODO期限を画面表示用の日付に整えます。
 */
function formatTaskDate_(value, timeZone) {
  if (!value) {
    return '';
  }

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    var pattern = value.getHours() || value.getMinutes()
      ? 'yyyy-MM-dd HH:mm'
      : 'yyyy-MM-dd';
    return Utilities.formatDate(value, timeZone, pattern);
  }

  var text = String(value).trim();
  var dateTimeMatch = text.match(/^(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2})/);

  if (dateTimeMatch) {
    return dateTimeMatch[1] + ' ' + dateTimeMatch[2];
  }

  return text;
}

/**
 * 期限順ソート用の時刻値を返します。期限なしは末尾にします。
 */
function getTaskDueTime_(dueDate) {
  if (!dueDate) {
    return 8640000000000000;
  }

  var time = new Date(dueDate).getTime();

  if (isNaN(time)) {
    return 8640000000000000;
  }

  return time;
}

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === 'todayEvents') {
    const callback = e && e.parameter && e.parameter.callback;

    try {
      const events = getTodayCalendarEvents();

      if (callback) {
        return createJsonpResponse_(callback, events);
      }

      return createJsonResponse(events);
    } catch (error) {
      const errorData = {
        error: true,
        message: error && error.message ? error.message : '予定を取得できませんでした。'
      };

      if (callback) {
        return createJsonpResponse_(callback, errorData);
      }

      return createJsonResponse(errorData);
    }
  }

  if (action === 'rangeEvents') {
    const callback = e && e.parameter && e.parameter.callback;
    const startDate = e && e.parameter && e.parameter.startDate;
    const endDate = e && e.parameter && e.parameter.endDate;

    try {
      const events = getCalendarEventsByRange(startDate, endDate);

      if (callback) {
        return createJsonpResponse_(callback, events);
      }

      return createJsonResponse(events);
    } catch (error) {
      const errorData = {
        error: true,
        message: error && error.message ? error.message : '予定を取得できませんでした。'
      };

      if (callback) {
        return createJsonpResponse_(callback, errorData);
      }

      return createJsonResponse(errorData);
    }
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('営業AI秘書')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function createJsonpResponse_(callback, data) {
  const callbackName = String(callback || '').trim();

  if (!/^[A-Za-z_$][0-9A-Za-z_$]*(\.[A-Za-z_$][0-9A-Za-z_$]*)*$/.test(callbackName)) {
    return createJsonResponse({
      error: true,
      message: 'Invalid JSONP callback.'
    });
  }

  return ContentService
    .createTextOutput(callbackName + '(' + JSON.stringify(data) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function createJsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
