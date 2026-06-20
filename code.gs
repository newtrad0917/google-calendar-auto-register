/**
 * Web画面を表示するための関数です。
 * URLを開くと、Index.html が表示されます。
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('営業AI秘書')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

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

  // 登録できたことをHTML画面へ返します。
  return {
    message: '登録完了しました。',
  };
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

  GmailApp.createDraft(
    data.to,
    data.subject,
    data.body || ''
  );

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