/**
 * Web画面を表示するための関数です。
 * URLを開くと、Index.html が表示されます。
 */

var CUSTOMER_SPREADSHEET_ID = '';
var CUSTOMER_SHEET_NAME = '顧客DB';
var PROJECT_SPREADSHEET_ID = '';
var PROJECT_SHEET_NAME = '案件DB';
var SALES_MEMO_SPREADSHEET_ID = '';
var SALES_MEMO_SHEET_NAME = '営業メモDB';
var BUSINESS_CARD_IMPORT_SHEET_NAME = '名刺インポート';
var SETUP_WIZARD_COMPLETE_PROPERTY_KEY = 'SALES_AI_SECRETARY_SETUP_COMPLETE';


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
        meetLink: meetEvent.meetLink || '',
        eventId: meetEvent.event && meetEvent.event.iCalUID
          ? meetEvent.event.iCalUID
          : ''
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

function updateCalendarEvent(eventId, payload) {
  const id = String(eventId || '').trim();
  const data = payload || {};
  const title = String(data.title || '').trim();

  if (!id) {
    throw new Error('更新する予定IDを確認できませんでした。');
  }

  if (!title) {
    throw new Error('タイトルを入力してください。');
  }

  if (!data.startIso) {
    throw new Error('開始日時を入力してください。');
  }

  const startTime = new Date(data.startIso);

  if (isNaN(startTime.getTime())) {
    throw new Error('開始日時を確認してください。');
  }

  const durationMinutes = Number(data.durationMinutes || 60);

  if (durationMinutes <= 0) {
    throw new Error('所要時間は1分以上にしてください。');
  }

  const endTime = new Date(startTime.getTime() + durationMinutes * 60 * 1000);
  const event = CalendarApp.getEventById(id);

  if (!event) {
    throw new Error('更新対象の予定が見つかりませんでした。');
  }

  event.setTitle(title);
  event.setTime(startTime, endTime);
  event.setLocation(String(data.location || '').trim());
  event.setDescription(String(data.memo || data.description || '').trim());

  if (Object.prototype.hasOwnProperty.call(data, 'eventColor')) {
    applyCalendarEventColor_(event, data.eventColor);
  }

  return {
    success: true,
    message: '予定を更新しました。',
    eventId: event.getId(),
    event: mapCalendarEventForList_(event, getScheduleTimeZone_(), new Date())
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
  const timeZone = getScheduleTimeZone_();
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
        id: event.getId(),
        eventId: event.getId(),
        title: event.getTitle(),
        startTime: Utilities.formatDate(
          startTime,
          timeZone,
          'HH:mm'
        ),
        endTime: Utilities.formatDate(
          endTime,
          timeZone,
          'HH:mm'
        ),
        start: startTime.toISOString(),
        end: endTime.toISOString(),
        dateKey: Utilities.formatDate(startTime, timeZone, 'yyyy-MM-dd'),
        location: event.getLocation() || '',
        description: event.getDescription() || '',
        status: isCurrent ? 'current' : 'next',
        allDay: event.isAllDayEvent()
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
    eventId: event.getId(),
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

  var sheet = spreadsheet.getSheetByName('Tasks');
  var wasCreated = false;

  if (!sheet) {
    sheet = spreadsheet.insertSheet('Tasks');
    wasCreated = true;
  }

  var headerResult = ensureTaskSheetHeaders_(sheet);
  styleTaskSheet_(sheet);

  PropertiesService
    .getScriptProperties()
    .setProperty('TASK_SPREADSHEET_ID', spreadsheetId);

  Logger.log('TASK_SPREADSHEET_IDを保存しました。');

  return {
    message: 'TASK_SPREADSHEET_IDを保存しました。',
    sheetName: 'Tasks',
    wasCreated: wasCreated,
    wroteHeaders: headerResult.wroteHeaders,
    addedHeaders: headerResult.addedHeaders
  };
}

function getTaskSheetHeaders_() {
  return [
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
}

function ensureTaskSheetHeaders_(sheet) {
  var desiredHeaders = getTaskSheetHeaders_();
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var headerValues = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(header) {
    return String(header || '').trim();
  });
  var hasExistingHeaders = headerValues.some(function(value) {
    return value !== '';
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

function styleTaskSheet_(sheet) {
  var lastColumn = Math.max(sheet.getLastColumn(), getTaskSheetHeaders_().length);

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, lastColumn).setFontWeight('bold');

  if (!sheet.getFilter()) {
    sheet
      .getRange(1, 1, Math.max(sheet.getLastRow(), 1), lastColumn)
      .createFilter();
  }

  sheet.setColumnWidths(1, lastColumn, 130);
}

function runInitialSetupWizard() {
  var steps = [
    { key: 'tasks', label: 'Tasks', action: setupTaskSpreadsheet },
    { key: 'customers', label: CUSTOMER_SHEET_NAME, action: setupCustomerDbSheet },
    { key: 'projects', label: PROJECT_SHEET_NAME, action: setupProjectDbSheet },
    { key: 'salesMemos', label: SALES_MEMO_SHEET_NAME, action: setupSalesMemoDbSheet },
    { key: 'businessCardImport', label: BUSINESS_CARD_IMPORT_SHEET_NAME, action: setupBusinessCardImportSheet }
  ];
  var results = [];

  steps.forEach(function(step) {
    try {
      var detail = step.action();
      results.push({
        key: step.key,
        label: step.label,
        ok: true,
        text: '作成済み',
        detail: detail || {}
      });
    } catch (error) {
      results.push({
        key: step.key,
        label: step.label,
        ok: false,
        text: '確認が必要',
        message: error && error.message ? error.message : 'セットアップに失敗しました。'
      });
    }
  });

  return {
    ok: results.every(function(item) {
      return item.ok;
    }),
    updatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    results: results
  };
}

function getSetupWizardCompletionStatus() {
  return {
    key: SETUP_WIZARD_COMPLETE_PROPERTY_KEY,
    completed: isSetupWizardCompleted_()
  };
}

function setSetupWizardCompleted() {
  PropertiesService
    .getUserProperties()
    .setProperty(SETUP_WIZARD_COMPLETE_PROPERTY_KEY, 'true');
  PropertiesService
    .getScriptProperties()
    .setProperty(SETUP_WIZARD_COMPLETE_PROPERTY_KEY, 'true');

  return {
    key: SETUP_WIZARD_COMPLETE_PROPERTY_KEY,
    completed: true
  };
}

function isSetupWizardCompleted_() {
  var userValue = PropertiesService
    .getUserProperties()
    .getProperty(SETUP_WIZARD_COMPLETE_PROPERTY_KEY);
  var scriptValue = PropertiesService
    .getScriptProperties()
    .getProperty(SETUP_WIZARD_COMPLETE_PROPERTY_KEY);

  return userValue === 'true' || scriptValue === 'true';
}

function getAppSettingsStatus() {
  var updatedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var result = {
    appName: '営業AI秘書',
    version: 'Ver1',
    weatherRegion: '朝倉・筑後地方',
    dataStoreName: '営業AI秘書DB',
    updatedAt: updatedAt,
    dbItems: [
      buildAppSettingsDbItem_('tasks', 'Tasks', false),
      buildAppSettingsDbItem_('customers', CUSTOMER_SHEET_NAME, false),
      buildAppSettingsDbItem_('projects', PROJECT_SHEET_NAME, false),
      buildAppSettingsDbItem_('salesMemos', SALES_MEMO_SHEET_NAME, false),
      buildAppSettingsDbItem_('businessCardImport', BUSINESS_CARD_IMPORT_SHEET_NAME, false)
    ]
  };

  try {
    var spreadsheet = getTaskSpreadsheet_();

    result.dbItems = [
      buildAppSettingsDbItem_('tasks', 'Tasks', Boolean(spreadsheet.getSheetByName('Tasks'))),
      buildAppSettingsDbItem_('customers', CUSTOMER_SHEET_NAME, Boolean(spreadsheet.getSheetByName(CUSTOMER_SHEET_NAME))),
      buildAppSettingsDbItem_('projects', PROJECT_SHEET_NAME, Boolean(spreadsheet.getSheetByName(PROJECT_SHEET_NAME))),
      buildAppSettingsDbItem_('salesMemos', SALES_MEMO_SHEET_NAME, Boolean(spreadsheet.getSheetByName(SALES_MEMO_SHEET_NAME))),
      buildAppSettingsDbItem_('businessCardImport', BUSINESS_CARD_IMPORT_SHEET_NAME, Boolean(spreadsheet.getSheetByName(BUSINESS_CARD_IMPORT_SHEET_NAME)))
    ];
    result.spreadsheetName = spreadsheet.getName();
    result.spreadsheetId = spreadsheet.getId();
    result.ok = true;
  } catch (error) {
    result.ok = false;
    result.message = error && error.message ? error.message : 'DB状態を確認できませんでした。';
  }

  return result;
}

function buildAppSettingsDbItem_(key, label, exists) {
  return {
    key: key,
    label: label,
    exists: exists,
    status: exists ? 'connected' : 'missing',
    text: exists ? '接続済み' : '未作成'
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
  var headers = getCustomerDbHeaderValues_(sheet);

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
  formatCustomerPhoneColumns_(sheet, headers);
}

function setCustomerDbColumnWidth_(sheet, headerName, width) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var index = headers.indexOf(headerName);

  if (index !== -1) {
    sheet.setColumnWidth(index + 1, width);
  }
}

function formatCustomerPhoneColumns_(sheet, headers) {
  ['phone', 'officePhone', 'mobilePhone'].forEach(function(headerName) {
    var index = headers.indexOf(headerName);

    if (index !== -1) {
      sheet
        .getRange(1, index + 1, sheet.getMaxRows(), 1)
        .setNumberFormat('@');
    }
  });
}

function setupBusinessCardImportSheet() {
  var spreadsheet = getTaskSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(BUSINESS_CARD_IMPORT_SHEET_NAME);
  var wasCreated = false;
  var headers = getBusinessCardImportHeaders_();

  if (!sheet) {
    sheet = spreadsheet.insertSheet(BUSINESS_CARD_IMPORT_SHEET_NAME);
    wasCreated = true;
  }

  var lastRow = sheet.getLastRow();
  var headerResult = ensureBusinessCardImportHeaders_(sheet, headers);

  styleBusinessCardImportSheet_(sheet, headers);

  return {
    ok: true,
    message: BUSINESS_CARD_IMPORT_SHEET_NAME + 'シートを準備しました。myBridge Excelの内容を2行目以降に貼り付けてください。',
    spreadsheetId: spreadsheet.getId(),
    sheetName: BUSINESS_CARD_IMPORT_SHEET_NAME,
    wasCreated: wasCreated,
    wroteHeaders: headerResult.wroteHeaders,
    addedHeaders: headerResult.addedHeaders,
    existingRows: Math.max(lastRow - 1, 0)
  };
}

function getBusinessCardImportHeaders_() {
  return [
    '会社名',
    '名前',
    '部署',
    '役職',
    '電子メール',
    '郵便番号',
    '会社住所',
    '会社電話',
    '会社FAX',
    '携帯電話',
    '名刺登録日',
    '名刺交換日',
    '名刺帳名',
    'グループ',
    'メモ'
  ];
}

function getBusinessCardImportHeaderValues_(sheet) {
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

function ensureBusinessCardImportHeaders_(sheet, desiredHeaders) {
  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var headerValues = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(header) {
    return String(header || '').trim();
  });
  var hasExistingHeaders = headerValues.some(function(value) {
    return value !== '';
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

function styleBusinessCardImportSheet_(sheet, desiredHeaders) {
  var lastColumn = Math.max(sheet.getLastColumn(), desiredHeaders.length);
  var headers = getBusinessCardImportHeaderValues_(sheet);

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, lastColumn).setFontWeight('bold');

  if (!sheet.getFilter()) {
    sheet
      .getRange(1, 1, Math.max(sheet.getLastRow(), 1), lastColumn)
      .createFilter();
  }

  sheet.setColumnWidths(1, lastColumn, 130);
  setBusinessCardImportColumnWidth_(sheet, '会社名', 190);
  setBusinessCardImportColumnWidth_(sheet, '名前', 140);
  setBusinessCardImportColumnWidth_(sheet, '部署', 160);
  setBusinessCardImportColumnWidth_(sheet, '役職', 140);
  setBusinessCardImportColumnWidth_(sheet, '電子メール', 210);
  setBusinessCardImportColumnWidth_(sheet, '会社住所', 260);
  setBusinessCardImportColumnWidth_(sheet, '会社電話', 150);
  setBusinessCardImportColumnWidth_(sheet, '携帯電話', 150);
  setBusinessCardImportColumnWidth_(sheet, 'グループ', 170);
  setBusinessCardImportColumnWidth_(sheet, 'メモ', 260);
  formatBusinessCardPhoneColumns_(sheet, headers);
}

function setBusinessCardImportColumnWidth_(sheet, headerName, width) {
  var headers = getBusinessCardImportHeaderValues_(sheet);
  var index = headers.indexOf(headerName);

  if (index !== -1) {
    sheet.setColumnWidth(index + 1, width);
  }
}

function formatBusinessCardPhoneColumns_(sheet, headers) {
  ['会社電話', '会社FAX', '携帯電話'].forEach(function(headerName) {
    var index = headers.indexOf(headerName);

    if (index !== -1) {
      sheet
        .getRange(1, index + 1, sheet.getMaxRows(), 1)
        .setNumberFormat('@');
    }
  });
}

function previewBusinessCardImport() {
  var records = getBusinessCardImportRows_();
  var existingCustomers = getCustomers();
  var summary = {
    ok: true,
    totalRows: records.totalRows,
    readableRows: 0,
    newCount: 0,
    duplicateCount: 0,
    skipCount: 0
  };

  records.rows.forEach(function(card) {
    if (isBusinessCardImportHeaderRow_(card)) {
      summary.skipCount += 1;
      return;
    }

    if (!hasBusinessCardImportData_(card)) {
      summary.skipCount += 1;
      return;
    }

    summary.readableRows += 1;
    var customer = convertBusinessCardToCustomer_(card);

    if (!customer.companyName) {
      summary.skipCount += 1;
      return;
    }

    if (findDuplicateCustomer_(customer, existingCustomers)) {
      summary.duplicateCount += 1;
      return;
    }

    summary.newCount += 1;
  });

  summary.message = '読み込み ' + summary.readableRows + '件 / 新規 ' + summary.newCount + '件 / 重複 ' + summary.duplicateCount + '件 / スキップ ' + summary.skipCount + '件';
  return summary;
}

function importBusinessCards() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var records = getBusinessCardImportRows_();
    var customerSheet = getCustomerSheet_();
    ensureCustomerDbHeaders_(customerSheet, getCustomerDbHeaders_());
    var customerHeaders = getCustomerDbHeaderValues_(customerSheet);
    formatCustomerPhoneColumns_(customerSheet, customerHeaders);
    var existingCustomers = getCustomers();
    var now = new Date();
    var result = {
      ok: true,
      totalRows: records.totalRows,
      readableRows: 0,
      importedCount: 0,
      newCount: 0,
      duplicateCount: 0,
      skipCount: 0
    };

    records.rows.forEach(function(card) {
      if (isBusinessCardImportHeaderRow_(card)) {
        result.skipCount += 1;
        return;
      }

      if (!hasBusinessCardImportData_(card)) {
        result.skipCount += 1;
        return;
      }

      result.readableRows += 1;
      var customer = convertBusinessCardToCustomer_(card);

      if (!customer.companyName) {
        result.skipCount += 1;
        return;
      }

      if (findDuplicateCustomer_(customer, existingCustomers)) {
        result.duplicateCount += 1;
        return;
      }

      var id = generateCustomerId_();
      var row = buildCustomerRowForHeaders_(Object.assign({}, customer, {
        id: id,
        createdAt: customer.createdAt || now,
        updatedAt: now
      }), customerHeaders);

      writeCustomerRow_(customerSheet, customerSheet.getLastRow() + 1, row, customerHeaders);
      existingCustomers.push(Object.assign({}, customer, {
        id: id
      }));
      result.importedCount += 1;
      result.newCount += 1;
    });

    result.message = '名刺データを取り込みました。新規 ' + result.importedCount + '件 / 重複スキップ ' + result.duplicateCount + '件 / スキップ ' + result.skipCount + '件';
    result.customers = getCustomers();
    return result;
  } finally {
    lock.releaseLock();
  }
}

function getBusinessCardImportRows_() {
  var sheet = getBusinessCardImportSheet_();
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 1) {
    return {
      totalRows: 0,
      rows: []
    };
  }

  var headers = getBusinessCardImportHeaderValues_(sheet);
  var range = sheet.getRange(2, 1, lastRow - 1, lastColumn);
  var values = range.getValues();
  var displayValues = range.getDisplayValues();
  var rows = values
    .map(function(row, index) {
      return normalizeBusinessCardRow_(row, headers, displayValues[index]);
    });

  return {
    totalRows: rows.length,
    rows: rows
  };
}

function getBusinessCardImportSheet_() {
  var spreadsheet = getTaskSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(BUSINESS_CARD_IMPORT_SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(BUSINESS_CARD_IMPORT_SHEET_NAME);
  }

  ensureBusinessCardImportHeaders_(sheet, getBusinessCardImportHeaders_());
  styleBusinessCardImportSheet_(sheet, getBusinessCardImportHeaders_());

  return sheet;
}

function normalizeBusinessCardRow_(row, headers, displayRow) {
  function value(headerName) {
    var index = headers.indexOf(headerName);
    return index === -1 ? '' : row[index];
  }

  function displayValue(headerName) {
    var index = headers.indexOf(headerName);
    return index === -1 || !displayRow ? '' : displayRow[index];
  }

  function text(headerName) {
    var rawValue = value(headerName);

    if (rawValue === null || rawValue === undefined || rawValue === '') {
      return '';
    }

    return String(rawValue).trim();
  }

  function phoneText(headerName) {
    var displayText = normalizeCustomerPhoneText_(displayValue(headerName));

    if (displayText) {
      return displayText;
    }

    return normalizeCustomerPhoneText_(value(headerName));
  }

  return {
    companyName: text('会社名'),
    contactName: text('名前'),
    department: text('部署'),
    title: text('役職'),
    email: text('電子メール'),
    postalCode: text('郵便番号'),
    address: text('会社住所'),
    officePhone: phoneText('会社電話'),
    fax: phoneText('会社FAX'),
    mobilePhone: phoneText('携帯電話'),
    registeredDate: formatBusinessCardDate_(value('名刺登録日')),
    exchangedDate: formatBusinessCardDate_(value('名刺交換日')),
    cardBookName: text('名刺帳名'),
    group: text('グループ'),
    memo: text('メモ')
  };
}

function convertBusinessCardToCustomer_(card) {
  var data = card || {};
  var departmentParts = [data.department, data.title].filter(function(value) {
    return String(value || '').trim() !== '';
  });
  var notes = [data.group, data.cardBookName].filter(function(value) {
    return String(value || '').trim() !== '';
  }).join(' / ');

  return {
    companyName: data.companyName || '',
    corporationName: '',
    contactName: data.contactName || '',
    department: departmentParts.join(' '),
    industry: detectCustomerIndustry_(data),
    prefecture: '',
    city: '',
    address: data.address || '',
    officePhone: data.officePhone || '',
    mobilePhone: data.mobilePhone || '',
    email: data.email || '',
    ccEmail: '',
    googleMapUrl: '',
    lastVisit: data.exchangedDate || '',
    nextVisit: '',
    memo: data.memo || '',
    projectCount: 0,
    hotLevel: 'WARM',
    aiSummary: '',
    createdAt: data.registeredDate || new Date(),
    updatedAt: new Date(),
    notes: notes
  };
}

function detectCustomerIndustry_(card) {
  var data = card || {};
  var text = [
    data.companyName,
    data.group,
    data.cardBookName
  ].join(' ');

  if (text.indexOf('病院') !== -1) {
    return '病院';
  }

  if (
    text.indexOf('社会福祉法人') !== -1 ||
    text.indexOf('福祉') !== -1 ||
    text.indexOf('施設') !== -1
  ) {
    return '施設';
  }

  if (
    text.indexOf('工場') !== -1 ||
    text.indexOf('製作所') !== -1 ||
    text.indexOf('製造') !== -1
  ) {
    return '工場';
  }

  if (
    text.indexOf('株式会社') !== -1 ||
    text.indexOf('有限会社') !== -1 ||
    text.indexOf('合同会社') !== -1
  ) {
    return '企業';
  }

  return 'その他';
}

function findDuplicateCustomer_(customer, existingCustomers) {
  var data = customer || {};
  var email = normalizeBusinessCardDuplicateText_(data.email);
  var mobilePhone = normalizeBusinessCardPhone_(data.mobilePhone);
  var officePhone = normalizeBusinessCardPhone_(data.officePhone);
  var companyContact = normalizeBusinessCardDuplicateText_(data.companyName) + '|' + normalizeBusinessCardDuplicateText_(data.contactName);
  var canCheckCompanyContact = normalizeBusinessCardDuplicateText_(data.companyName) && normalizeBusinessCardDuplicateText_(data.contactName);
  var customers = Array.isArray(existingCustomers) ? existingCustomers : [];

  for (var i = 0; i < customers.length; i++) {
    var existing = customers[i] || {};

    if (email && email === normalizeBusinessCardDuplicateText_(existing.email)) {
      return existing;
    }

    if (mobilePhone && mobilePhone === normalizeBusinessCardPhone_(existing.mobilePhone)) {
      return existing;
    }

    if (officePhone && officePhone === normalizeBusinessCardPhone_(existing.officePhone)) {
      return existing;
    }

    if (
      canCheckCompanyContact &&
      companyContact === normalizeBusinessCardDuplicateText_(existing.companyName) + '|' + normalizeBusinessCardDuplicateText_(existing.contactName || existing.personName)
    ) {
      return existing;
    }
  }

  return null;
}

function hasBusinessCardImportData_(card) {
  var data = card || {};

  return [
    data.companyName,
    data.contactName,
    data.email,
    data.officePhone,
    data.mobilePhone,
    data.address,
    data.memo
  ].some(function(value) {
    return String(value || '').trim() !== '';
  });
}

function isBusinessCardImportHeaderRow_(card) {
  var data = card || {};

  return String(data.companyName || '').trim() === '会社名' &&
    String(data.contactName || '').trim() === '名前' &&
    String(data.email || '').trim() === '電子メール';
}

function formatBusinessCardDate_(value) {
  if (!value) {
    return '';
  }

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  return String(value || '').trim();
}

function normalizeBusinessCardDuplicateText_(value) {
  return String(value || '')
    .trim()
    .replace(/\s+/g, '')
    .toLowerCase();
}

function normalizeBusinessCardPhone_(value) {
  return String(value || '').replace(/[^\d]/g, '');
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
    'customerName',
    'submissionEnabled',
    'submissionHidden',
    'estimateSubmissionEnabled',
    'estimateSubmissionHidden',
    'estimateStatus',
    'estimateSubmittedDate',
    'invoiceSubmissionEnabled',
    'invoiceSubmissionHidden',
    'invoiceStatus',
    'invoiceSubmittedDate',
    'paymentConfirmedDate',
    'submissionMemo',
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
  setProjectDbColumnWidth_(sheet, 'customerName', 180);
  setProjectDbColumnWidth_(sheet, 'submissionEnabled', 150);
  setProjectDbColumnWidth_(sheet, 'submissionHidden', 150);
  setProjectDbColumnWidth_(sheet, 'estimateSubmissionEnabled', 190);
  setProjectDbColumnWidth_(sheet, 'estimateSubmissionHidden', 190);
  setProjectDbColumnWidth_(sheet, 'estimateStatus', 140);
  setProjectDbColumnWidth_(sheet, 'estimateSubmittedDate', 170);
  setProjectDbColumnWidth_(sheet, 'invoiceSubmissionEnabled', 190);
  setProjectDbColumnWidth_(sheet, 'invoiceSubmissionHidden', 190);
  setProjectDbColumnWidth_(sheet, 'invoiceStatus', 140);
  setProjectDbColumnWidth_(sheet, 'invoiceSubmittedDate', 170);
  setProjectDbColumnWidth_(sheet, 'paymentConfirmedDate', 170);
  setProjectDbColumnWidth_(sheet, 'submissionMemo', 220);
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

function setupSalesMemoDbSheet() {
  var spreadsheet = getTaskSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(SALES_MEMO_SHEET_NAME);
  var wasCreated = false;
  var headers = getSalesMemoDbHeaders_();

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SALES_MEMO_SHEET_NAME);
    wasCreated = true;
  }

  var lastRow = sheet.getLastRow();
  var headerResult = ensureSalesMemoDbHeaders_(sheet, headers);

  styleSalesMemoDbSheet_(sheet, headers);

  PropertiesService
    .getScriptProperties()
    .setProperty('SALES_MEMO_SPREADSHEET_ID', spreadsheet.getId());

  var result = {
    message: SALES_MEMO_SHEET_NAME + 'シートの初期化が完了しました。',
    spreadsheetId: spreadsheet.getId(),
    sheetName: SALES_MEMO_SHEET_NAME,
    wasCreated: wasCreated,
    wroteHeaders: headerResult.wroteHeaders,
    addedHeaders: headerResult.addedHeaders,
    existingRows: lastRow
  };

  Logger.log(JSON.stringify(result, null, 2));

  return result;
}

function addSalesMemo(memo) {
  var data = memo || {};
  var sheet = getSalesMemoSheet_();
  var headers = getSalesMemoDbHeaderValues_(sheet);
  var now = new Date();
  var id = String(data.id || '').trim() || generateSalesMemoId_();
  var row = buildSalesMemoRowForHeaders_(Object.assign({}, data, {
    id: id,
    createdAt: now,
    updatedAt: now
  }), headers);
  var todoResult = null;
  var todoError = '';
  var todoText = String(data.todoText || '').trim();

  sheet.appendRow(row);

  if (todoText) {
    try {
      todoResult = addTask({
        title: todoText,
        category: '営業',
        dueDate: String(data.nextVisitDate || '').trim(),
        priority: '中',
        memo: String(data.memo || data.nextAction || '').trim()
      });
    } catch (error) {
      todoError = error && error.message ? error.message : String(error || '');
    }
  }

  return {
    id: id,
    message: '営業メモを保存しました。',
    todoAdded: Boolean(todoResult),
    todoError: todoError
  };
}

function getSalesMemoDashboardStats() {
  var sheet = getSalesMemoSheet_();
  var values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return {
      todayCount: 0
    };
  }

  var headers = values[0].map(function(header) {
    return String(header || '').trim();
  });
  var createdAtIndex = headers.indexOf('createdAt');
  var visitDateIndex = headers.indexOf('visitDate');
  var todayKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var todayCount = values.slice(1).filter(function(row) {
    var dateValue = createdAtIndex !== -1 ? row[createdAtIndex] : '';
    var dateKey = formatSalesMemoDashboardDate_(dateValue);

    if (!dateKey && visitDateIndex !== -1) {
      dateKey = formatSalesMemoDashboardDate_(row[visitDateIndex]);
    }

    return dateKey === todayKey;
  }).length;

  return {
    todayCount: todayCount
  };
}

function formatSalesMemoDashboardDate_(value) {
  if (!value) {
    return '';
  }

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }

  var text = String(value || '').trim();
  var match = text.match(/^(\d{4}-\d{2}-\d{2})/);
  return match ? match[1] : '';
}

function generateSalesMemoId_() {
  var sheet = getSalesMemoSheet_();
  var headers = getSalesMemoDbHeaderValues_(sheet);
  var idColumn = headers.indexOf('id');
  var lastRow = sheet.getLastRow();
  var maxNumber = 0;

  if (idColumn === -1 || lastRow < 2) {
    return 'M000001';
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

  return 'M' + ('000000' + (maxNumber + 1)).slice(-6);
}

function getSalesMemoSheet_() {
  var properties = PropertiesService.getScriptProperties();
  var spreadsheetId = SALES_MEMO_SPREADSHEET_ID ||
    properties.getProperty('SALES_MEMO_SPREADSHEET_ID') ||
    properties.getProperty('TASK_SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error('TASK_SPREADSHEET_IDが未設定です。setupTaskSpreadsheet()を実行してからsetupSalesMemoDbSheet()を実行してください。');
  }

  var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var sheet = spreadsheet.getSheetByName(SALES_MEMO_SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SALES_MEMO_SHEET_NAME);
    PropertiesService
      .getScriptProperties()
      .setProperty('SALES_MEMO_SPREADSHEET_ID', spreadsheet.getId());
  }

  ensureSalesMemoDbHeaders_(sheet, getSalesMemoDbHeaders_());
  styleSalesMemoDbSheet_(sheet, getSalesMemoDbHeaders_());

  return sheet;
}

function buildSalesMemoRowForHeaders_(memo, headers, existingRow) {
  var data = memo || {};
  var existing = existingRow || [];
  var knownFields = getSalesMemoDbHeaders_().reduce(function(result, header) {
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

function getSalesMemoDbHeaders_() {
  return [
    'id',
    'customerId',
    'projectId',
    'eventTitle',
    'visitDate',
    'contactName',
    'memo',
    'nextAction',
    'nextVisitDate',
    'todoText',
    'createdAt',
    'updatedAt',
    'notes'
  ];
}

function getSalesMemoDbHeaderValues_(sheet) {
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

function ensureSalesMemoDbHeaders_(sheet, desiredHeaders) {
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

function styleSalesMemoDbSheet_(sheet, desiredHeaders) {
  var lastColumn = Math.max(sheet.getLastColumn(), desiredHeaders.length);

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, lastColumn).setFontWeight('bold');

  if (!sheet.getFilter()) {
    sheet
      .getRange(1, 1, Math.max(sheet.getLastRow(), 1), lastColumn)
      .createFilter();
  }

  sheet.setColumnWidths(1, lastColumn, 130);
  setSalesMemoDbColumnWidth_(sheet, 'eventTitle', 220);
  setSalesMemoDbColumnWidth_(sheet, 'contactName', 150);
  setSalesMemoDbColumnWidth_(sheet, 'memo', 260);
  setSalesMemoDbColumnWidth_(sheet, 'nextAction', 220);
  setSalesMemoDbColumnWidth_(sheet, 'todoText', 220);
  setSalesMemoDbColumnWidth_(sheet, 'notes', 220);
}

function setSalesMemoDbColumnWidth_(sheet, headerName, width) {
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
  var range = sheet.getDataRange();
  var values = range.getValues();
  var displayValues = range.getDisplayValues();

  if (values.length <= 1) {
    return [];
  }

  var headers = displayValues[0].map(function(header) {
    return String(header || '').trim();
  });

  return values
    .slice(1)
    .map(function(row, index) {
      return normalizeCustomerRow_(row, headers, displayValues[index + 1]);
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

function getCustomerPhoneDiagnostics() {
  var sheet = getCustomerSheet_();
  var range = sheet.getDataRange();
  var values = range.getValues();
  var displayValues = range.getDisplayValues();
  var headers = displayValues.length
    ? displayValues[0].map(function(header) {
        return String(header || '').trim();
      })
    : [];
  var phoneHeaders = ['phone', 'officePhone', 'mobilePhone'];
  var phoneIndexes = phoneHeaders.reduce(function(result, header) {
    result[header] = headers.indexOf(header);
    return result;
  }, {});

  return {
    sheetName: sheet.getName(),
    lastRow: sheet.getLastRow(),
    phoneIndexes: phoneIndexes,
    rows: values.slice(1, 11).map(function(row, index) {
      var displayRow = displayValues[index + 1] || [];
      var normalized = normalizeCustomerRow_(row, headers, displayRow);
      var item = {
        rowNumber: index + 2,
        id: normalized.id,
        companyName: normalized.companyName,
        resolved: {
          phone: normalized.phone,
          officePhone: normalized.officePhone,
          mobilePhone: normalized.mobilePhone
        },
        raw: {},
        display: {}
      };

      phoneHeaders.forEach(function(header) {
        var columnIndex = phoneIndexes[header];
        item.raw[header] = columnIndex === -1 ? '' : row[columnIndex];
        item.display[header] = columnIndex === -1 ? '' : displayRow[columnIndex];
      });

      return item;
    })
  };
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
    formatCustomerPhoneColumns_(sheet, headers);
    var now = new Date();
    var id = String(data.id || '').trim() || generateCustomerId_();
    var row = buildCustomerRowForHeaders_(Object.assign({}, data, {
      id: id,
      companyName: companyName,
      createdAt: now,
      updatedAt: now
    }), headers);

    writeCustomerRow_(sheet, sheet.getLastRow() + 1, row, headers);

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
    formatCustomerPhoneColumns_(sheet, headers);
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

    writeCustomerRow_(sheet, rowNumber, row, headers);

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
      return isCustomerPhoneHeader_(header)
        ? normalizeCustomerPhoneText_(data[header])
        : data[header];
    }

    if (!knownFields[header]) {
      return existing[index] === undefined || existing[index] === null
        ? ''
        : existing[index];
    }

    return '';
  });
}

function writeCustomerRow_(sheet, rowNumber, row, headers) {
  ensureSheetRowCapacity_(sheet, rowNumber);
  formatCustomerPhoneColumns_(sheet, headers);
  formatCustomerPhoneCells_(sheet, rowNumber, headers);
  sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
}

function ensureSheetRowCapacity_(sheet, rowNumber) {
  var maxRows = sheet.getMaxRows();

  if (rowNumber > maxRows) {
    sheet.insertRowsAfter(maxRows, rowNumber - maxRows);
  }
}

function formatCustomerPhoneCells_(sheet, rowNumber, headers) {
  ['phone', 'officePhone', 'mobilePhone'].forEach(function(headerName) {
    var index = headers.indexOf(headerName);

    if (index !== -1) {
      sheet
        .getRange(rowNumber, index + 1, 1, 1)
        .setNumberFormat('@');
    }
  });
}

function isCustomerPhoneHeader_(header) {
  return ['phone', 'officePhone', 'mobilePhone'].indexOf(header) !== -1;
}

function normalizeCustomerPhoneText_(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }

  return formatCustomerPhoneText_(String(value).trim());
}

function formatCustomerPhoneText_(phoneText) {
  var text = String(phoneText || '').trim();
  var digits = text.replace(/[^\d]/g, '');

  if (/^(070|080|090)\d{8}$/.test(digits)) {
    return digits.replace(/^(\d{3})(\d{4})(\d{4})$/, '$1-$2-$3');
  }

  // 推定補完: Sheetsで数値化済みの携帯番号は先頭0を復元できないため、
  // 10桁かつ7/8/9始まりの場合だけ表示・以後の保存用に先頭0を補い、携帯形式へ整えます。
  if (/^[789]\d{9}$/.test(digits)) {
    return ('0' + digits).replace(/^(\d{3})(\d{4})(\d{4})$/, '$1-$2-$3');
  }

  return text;
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
    customerName: text('customerName'),
    submissionEnabled: text('submissionEnabled'),
    submissionHidden: text('submissionHidden'),
    estimateSubmissionEnabled: text('estimateSubmissionEnabled'),
    estimateSubmissionHidden: text('estimateSubmissionHidden'),
    estimateStatus: text('estimateStatus'),
    estimateSubmittedDate: dateText('estimateSubmittedDate'),
    invoiceSubmissionEnabled: text('invoiceSubmissionEnabled'),
    invoiceSubmissionHidden: text('invoiceSubmissionHidden'),
    invoiceStatus: text('invoiceStatus'),
    invoiceSubmittedDate: dateText('invoiceSubmittedDate'),
    paymentConfirmedDate: dateText('paymentConfirmedDate'),
    submissionMemo: text('submissionMemo'),
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

function normalizeCustomerRow_(row, headers, displayRow) {
  function value(name) {
    var index = headers.indexOf(name);
    return index === -1 ? '' : row[index];
  }

  function displayValue(name) {
    var index = headers.indexOf(name);
    return index === -1 || !displayRow ? '' : displayRow[index];
  }

  function text(name) {
    var rawValue = value(name);

    if (rawValue === null || rawValue === undefined || rawValue === '') {
      return '';
    }

    return String(rawValue).trim();
  }

  function phoneText(name) {
    var displayText = normalizeCustomerPhoneText_(displayValue(name));

    if (displayText) {
      return displayText;
    }

    return normalizeCustomerPhoneText_(value(name));
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
  var legacyPhone = phoneText('phone');
  var officePhone = phoneText('officePhone') || legacyPhone;
  var mobilePhone = phoneText('mobilePhone');

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
      const response = createScheduleSuccessResponse_(
        events,
        events.length ? '' : '本日の予定はありません'
      );

      if (callback) {
        return createJsonpResponse_(callback, response);
      }

      return createJsonResponse(response);
    } catch (error) {
      const errorData = createScheduleErrorResponse_(error);

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
      const response = createScheduleSuccessResponse_(
        events,
        events.length ? '' : '指定期間の予定はありません'
      );

      if (callback) {
        return createJsonpResponse_(callback, response);
      }

      return createJsonResponse(response);
    } catch (error) {
      const errorData = createScheduleErrorResponse_(error);

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

function createScheduleSuccessResponse_(events, message) {
  const eventList = Array.isArray(events) ? events : [];

  return {
    ok: true,
    events: eventList,
    count: eventList.length,
    message: message || '',
    timeZone: getScheduleTimeZone_(),
    generatedAt: new Date().toISOString()
  };
}

function createScheduleErrorResponse_(error) {
  return {
    ok: false,
    error: true,
    errorCode: classifyScheduleServerError_(error),
    message: getScheduleServerErrorMessage_(error),
    timeZone: getScheduleTimeZone_(),
    generatedAt: new Date().toISOString()
  };
}

function classifyScheduleServerError_(error) {
  const message = error && error.message ? String(error.message) : '';

  if (/permission|authorization|authorize|権限|承認|認証|アクセス/i.test(message)) {
    return 'SCHEDULE_ERROR_AUTH_REQUIRED';
  }

  if (/Calendar|カレンダー/i.test(message)) {
    return 'SCHEDULE_ERROR_CALENDAR_API';
  }

  if (/範囲|日付|date/i.test(message)) {
    return 'SCHEDULE_ERROR_INVALID_RANGE';
  }

  return 'SCHEDULE_ERROR_SERVER_EXCEPTION';
}

function getScheduleServerErrorMessage_(error) {
  const errorCode = classifyScheduleServerError_(error);

  if (errorCode === 'SCHEDULE_ERROR_AUTH_REQUIRED') {
    return 'カレンダー権限の確認が必要です。';
  }

  if (errorCode === 'SCHEDULE_ERROR_CALENDAR_API') {
    return 'カレンダー予定を確認できませんでした。';
  }

  return error && error.message ? error.message : '予定を取得できませんでした。';
}

function getScheduleTimeZone_() {
  return Session.getScriptTimeZone() || 'Asia/Tokyo';
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
