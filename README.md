# google-calendar-auto-register
Google Calendar auto registration prototype built with GAS.
## 開発状況

### Ver1 完成

現在、以下の機能が動作します。

- 予定文の入力
- 確認画面での内容修正
- Googleカレンダーへの予定登録
- iPhone対応
- Google Apps Script（GAS）で構築

### 音声入力について

当初はWeb Speech API（SpeechRecognition）を利用した音声入力を実装していましたが、iPhone Safari と Google Apps Script Webアプリの組み合わせで動作が不安定だったため、iPhone標準キーボードの音声入力を利用する方式へ変更しました。

### 今後の開発予定

- Gmail下書き自動作成
- 音声入力強化
- Google Maps連携
- 見積作成支援
- 営業AI秘書化
