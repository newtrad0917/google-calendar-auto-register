/**
 * Web画面を表示するための関数です。
 * URLを開くと、Index.html が表示されます。
 */


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

/**
 * TODO保存先スプレッドシートはScript PropertiesのTASK_SPREADSHEET_IDで管理します。
 * 将来、顧客DB、案件DB、活動履歴DB、AI提案DBも同じ方式でID管理する予定です。
 * 想定キー: CUSTOMER_SPREADSHEET_ID, PROJECT_SPREADSHEET_ID,
 * ACTIVITY_LOG_SPREADSHEET_ID, AI_PROPOSAL_SPREADSHEET_ID
 * DriveAppでファイル名検索しないことで、権限エラーや同名ファイル問題を避けます。
 */
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
