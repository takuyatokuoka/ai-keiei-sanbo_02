/**
 * email_classifier.gs
 *
 * Gmail の「要処理」ラベルが付いた未読メールを Claude API で分類し、
 * スプレッドシートへの記録と Slack 通知を行う Google Apps Script。
 *
 * 処理の流れ:
 *   1. 「要処理」ラベルの未読メールを取得する
 *   2. 本文を Claude API に送り「クレーム」「質問」「注文」「その他」に分類する
 *   3. 受信日時・送信者・件名・分類・要約を「メールログ」シートに記録する
 *   4. Slack Incoming Webhook で担当者へ通知する
 *   5. 「処理済み」ラベルを付与し「要処理」ラベルを外す
 *
 * 事前準備 (スクリプトプロパティ):
 *   CLAUDE_API_KEY    ... Claude API キー
 *   SLACK_WEBHOOK_URL ... Slack Incoming Webhook の URL
 *   SPREADSHEET_ID    ... (任意) 記録先スプレッドシートの ID。
 *                         未設定の場合はコンテナバインドのアクティブなスプレッドシートを使用する。
 */

// ===== 定数 =====

/** 処理対象の未読メールに付いているラベル名 */
var LABEL_TODO = '要処理';
/** 処理完了後に付与するラベル名 */
var LABEL_DONE = '処理済み';
/** メール記録用シート名 */
var SHEET_LOG = 'メールログ';
/** エラー記録用シート名 */
var SHEET_ERROR = 'エラーログ';

/** Claude API のエンドポイント */
var CLAUDE_API_URL = 'https://api.anthropic.com/v1/messages';
/** 使用するモデル (コスト重視のため claude-haiku の最新バージョン) */
var CLAUDE_MODEL = 'claude-haiku-4-5';
/** Claude API のバージョンヘッダー */
var ANTHROPIC_VERSION = '2023-06-01';

/** 許可する分類ラベル */
var CATEGORIES = ['クレーム', '質問', '注文', 'その他'];

/** 1 回の実行で処理するスレッドの上限 (実行時間制限対策) */
var MAX_THREADS_PER_RUN = 20;

// ===== トリガー設定 =====

/**
 * 5 分おきに processEmails を実行するトリガーを作成する。
 * 重複作成を防ぐため、既存の同名トリガーは削除してから登録する。
 * この関数は初回に手動で 1 度だけ実行すればよい。
 */
function createTrigger() {
  // 既存の processEmails 用トリガーを削除する
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processEmails') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 5 分間隔のトリガーを新規作成する
  ScriptApp.newTrigger('processEmails')
    .timeBased()
    .everyMinutes(5)
    .create();

  Logger.log('5 分おきのトリガーを作成しました。');
}

// ===== メイン処理 =====

/**
 * 「要処理」ラベルの未読メールを分類・記録・通知し、ラベルを付け替える。
 * トリガーから 5 分おきに呼び出される想定。
 */
function processEmails() {
  var props = PropertiesService.getScriptProperties();
  var claudeApiKey = props.getProperty('CLAUDE_API_KEY');
  var slackWebhookUrl = props.getProperty('SLACK_WEBHOOK_URL');

  // 必須のスクリプトプロパティが無い場合は中断する
  if (!claudeApiKey || !slackWebhookUrl) {
    logError_('processEmails', 'スクリプトプロパティ CLAUDE_API_KEY / SLACK_WEBHOOK_URL が設定されていません。', '');
    return;
  }

  var todoLabel = GmailApp.getUserLabelByName(LABEL_TODO);
  if (!todoLabel) {
    logError_('processEmails', 'ラベル「' + LABEL_TODO + '」が Gmail に存在しません。', '');
    return;
  }
  var doneLabel = getOrCreateLabel_(LABEL_DONE);

  // 「要処理」ラベルが付いた未読スレッドを検索する
  var threads = GmailApp.search('label:' + LABEL_TODO + ' is:unread', 0, MAX_THREADS_PER_RUN);
  if (threads.length === 0) {
    return; // 対象なし
  }

  for (var t = 0; t < threads.length; t++) {
    var thread = threads[t];
    try {
      processThread_(thread, claudeApiKey, slackWebhookUrl, todoLabel, doneLabel);
    } catch (err) {
      // 1 スレッドの失敗が全体を止めないよう、個別に握りつぶしてログへ記録する
      logError_('processThread_', String(err && err.message ? err.message : err), String(err && err.stack ? err.stack : ''));
    }
  }
}

/**
 * 1 スレッドを処理する。
 *
 * @param {GmailThread} thread          対象スレッド
 * @param {string}      claudeApiKey    Claude API キー
 * @param {string}      slackWebhookUrl Slack Webhook URL
 * @param {GmailLabel}  todoLabel       「要処理」ラベル
 * @param {GmailLabel}  doneLabel       「処理済み」ラベル
 */
function processThread_(thread, claudeApiKey, slackWebhookUrl, todoLabel, doneLabel) {
  var messages = thread.getMessages();
  // スレッド内の最新メッセージを分類対象とする
  var message = messages[messages.length - 1];

  var receivedAt = message.getDate();
  var sender = message.getFrom();
  var subject = message.getSubject();
  var body = message.getPlainBody();

  // Claude API で分類と要約を取得する
  var result = classifyWithClaude_(claudeApiKey, subject, body);

  // スプレッドシートの「メールログ」シートへ記録する
  logMail_(receivedAt, sender, subject, result.classification, result.summary);

  // Slack へ通知する (件名・分類・要約を含める)
  notifySlack_(slackWebhookUrl, subject, result.classification, result.summary);

  // ラベルを付け替え、未読を既読にする
  thread.addLabel(doneLabel);
  thread.removeLabel(todoLabel);
  thread.markRead();
}

// ===== Claude API 呼び出し =====

/**
 * Claude API にメール本文を送り、分類と要約を取得する。
 *
 * @param {string} apiKey  Claude API キー
 * @param {string} subject 件名
 * @param {string} body    本文 (プレーンテキスト)
 * @return {{classification: string, summary: string}} 分類結果と要約
 */
function classifyWithClaude_(apiKey, subject, body) {
  // 長すぎる本文はトークン節約のため先頭部分のみ送信する
  var trimmedBody = body.length > 6000 ? body.substring(0, 6000) : body;

  var systemPrompt =
    'あなたはカスタマーサポートのメール分類担当です。' +
    '受け取ったメールを必ず次の 4 つのいずれかに分類してください: ' +
    CATEGORIES.join(' / ') + '。' +
    '出力は JSON のみとし、前後に説明文やコードブロック記号を付けないでください。' +
    '形式: {"classification":"<分類>","summary":"<日本語で1〜2文の要約>"}';

  var userContent =
    '件名: ' + subject + '\n\n' +
    '本文:\n' + trimmedBody;

  var payload = {
    model: CLAUDE_MODEL,
    max_tokens: 512,
    system: systemPrompt,
    messages: [
      { role: 'user', content: userContent }
    ]
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': ANTHROPIC_VERSION
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(CLAUDE_API_URL, options);
  var statusCode = response.getResponseCode();
  var responseText = response.getContentText();

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('Claude API エラー (HTTP ' + statusCode + '): ' + responseText);
  }

  var json = JSON.parse(responseText);

  // content 配列から text ブロックを連結する
  var text = '';
  if (json.content && json.content.length) {
    for (var i = 0; i < json.content.length; i++) {
      if (json.content[i].type === 'text') {
        text += json.content[i].text;
      }
    }
  }
  text = text.trim();

  // 応答から JSON 部分を取り出してパースする
  var parsed = extractJson_(text);
  var classification = parsed && parsed.classification ? String(parsed.classification).trim() : '';
  var summary = parsed && parsed.summary ? String(parsed.summary).trim() : '';

  // 想定外の分類が返ってきた場合は「その他」に寄せる
  if (CATEGORIES.indexOf(classification) === -1) {
    classification = 'その他';
  }
  if (!summary) {
    summary = text || '(要約を取得できませんでした)';
  }

  return { classification: classification, summary: summary };
}

/**
 * 文字列から最初の JSON オブジェクトを抽出してパースする。
 * コードブロックや余分な前置きが付いていても取り出せるようにする。
 *
 * @param {string} text 対象文字列
 * @return {Object|null} パース結果。失敗時は null
 */
function extractJson_(text) {
  if (!text) {
    return null;
  }
  try {
    return JSON.parse(text);
  } catch (e) {
    // そのままパースできない場合は最初の { から最後の } までを取り出す
    var start = text.indexOf('{');
    var end = text.lastIndexOf('}');
    if (start !== -1 && end !== -1 && end > start) {
      try {
        return JSON.parse(text.substring(start, end + 1));
      } catch (e2) {
        return null;
      }
    }
    return null;
  }
}

// ===== Slack 通知 =====

/**
 * Slack Incoming Webhook へ通知を送る。
 *
 * @param {string} webhookUrl     Webhook URL
 * @param {string} subject        件名
 * @param {string} classification 分類
 * @param {string} summary        要約
 */
function notifySlack_(webhookUrl, subject, classification, summary) {
  var messageText =
    '*新着メールを分類しました*\n' +
    '• 件名: ' + subject + '\n' +
    '• 分類: ' + classification + '\n' +
    '• 要約: ' + summary;

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ text: messageText }),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(webhookUrl, options);
  var statusCode = response.getResponseCode();
  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('Slack 通知エラー (HTTP ' + statusCode + '): ' + response.getContentText());
  }
}

// ===== スプレッドシート操作 =====

/**
 * 記録先のスプレッドシートを取得する。
 * スクリプトプロパティ SPREADSHEET_ID があればそれを開き、無ければアクティブなものを使う。
 *
 * @return {Spreadsheet} スプレッドシート
 */
function getSpreadsheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) {
    return SpreadsheetApp.openById(id);
  }
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error('スプレッドシートを特定できません。スクリプトプロパティ SPREADSHEET_ID を設定してください。');
  }
  return active;
}

/**
 * 指定名のシートを取得する。無ければ作成し、ヘッダー行を書き込む。
 *
 * @param {string}   name    シート名
 * @param {string[]} headers ヘッダー行の内容
 * @return {Sheet} シート
 */
function getOrCreateSheet_(name, headers) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

/**
 * 「メールログ」シートに 1 行追記する。
 * 列の順序: 受信日時・送信者・件名・分類・要約
 *
 * @param {Date}   receivedAt     受信日時
 * @param {string} sender         送信者
 * @param {string} subject        件名
 * @param {string} classification 分類
 * @param {string} summary        要約
 */
function logMail_(receivedAt, sender, subject, classification, summary) {
  var sheet = getOrCreateSheet_(SHEET_LOG, ['受信日時', '送信者', '件名', '分類', '要約']);
  var timezone = Session.getScriptTimeZone();
  var formattedDate = Utilities.formatDate(receivedAt, timezone, 'yyyy-MM-dd HH:mm:ss');
  sheet.appendRow([formattedDate, sender, subject, classification, summary]);
}

/**
 * 「エラーログ」シートにエラー内容を記録する。
 * ログ記録自体の失敗でスクリプトが止まらないよう、内部で例外を握りつぶす。
 *
 * @param {string} context   発生箇所を示す文字列
 * @param {string} errMessage エラーメッセージ
 * @param {string} stack     スタックトレース (任意)
 */
function logError_(context, errMessage, stack) {
  try {
    var sheet = getOrCreateSheet_(SHEET_ERROR, ['発生日時', '発生箇所', 'エラー内容', 'スタックトレース']);
    var timezone = Session.getScriptTimeZone();
    var now = Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([now, context, errMessage, stack || '']);
  } catch (e) {
    // シートへ書けない場合は実行ログにのみ出力する
    Logger.log('エラーログの記録に失敗: ' + e + ' / 元のエラー: ' + context + ' - ' + errMessage);
  }
}

// ===== ラベル操作 =====

/**
 * 指定名の Gmail ラベルを取得する。存在しなければ作成する。
 *
 * @param {string} name ラベル名
 * @return {GmailLabel} ラベル
 */
function getOrCreateLabel_(name) {
  var label = GmailApp.getUserLabelByName(name);
  if (!label) {
    label = GmailApp.createLabel(name);
  }
  return label;
}
