/**
 * Web画面を表示するための関数です。
 * URLを開くと、Index.html が表示されます。
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('かんたん予定登録');
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
