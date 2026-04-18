// ============================================
// ToDo アプリ GAS コード v2（定時通知2スロット対応）
// ============================================
// 設定
var SPREADSHEET_ID = '1lIYGcsu_XWzweXHj82M4QM2pCoSLtSiFsBj5hLfzNkg';
var LINE_TOKEN = 'JessPeHRBhbEDKt3GLAD6UQeG3QuV6QBQupkPIQKuQJM93vM9Z5d6X9QCtEUo3wJy2cJxoY7fjJxBbVBGwOnltL2QKdS4zJISJ68WLl2SSyz47x9XNieCYAqJHyf/2n7kP0+98HP3/d90nP0Jtg8OgdB04t89/1O/w1cDnyilFU=';
var APP_URL = 'https://todo-app-tawny-iota-98.vercel.app';

// シート名
var SHEET_NOTIFICATIONS  = 'notifications';
var SHEET_USERS          = 'users';
var SHEET_SYNC           = 'sync';
var SHEET_LINE_MESSAGES  = 'line_messages';

// usersシート列インデックス（0始まり）
var COL_USER_ID           = 0; // A: アプリのuserId（トークン）
var COL_LINE_USER_ID      = 1; // B: LINEのuserId
var COL_CREATED_AT        = 2; // C: 登録日時
var COL_NOTIFY1_TIME      = 3; // D: 通知1の時刻 e.g. "07:00"  ← 旧notifyTime
var COL_NOTIFY1_LAST_SENT = 4; // E: 通知1の最終送信日          ← 旧lastDailyNotify
var COL_NOTIFY1_ENABLED   = 5; // F: 通知1のON/OFF (NEW)
var COL_NOTIFY2_TIME      = 6; // G: 通知2の時刻 (NEW)
var COL_NOTIFY2_ENABLED   = 7; // H: 通知2のON/OFF (NEW)
var COL_NOTIFY2_LAST_SENT = 8; // I: 通知2の最終送信日 (NEW)

// ============================================
// doGet: データ読込（load）用
// ============================================
function doGet(e) {
  try {
    var action = e.parameter.action;
    var userId = e.parameter.userId;
    if (action === 'load' && userId) return handleLoad(userId);
    return jsonResponse({ error: 'invalid request' });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ============================================
// doPost: LINE Webhook + アプリからの書き込み
// ============================================
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    // --- LINE Webhook ---
    if (body.events) {
      body.events.forEach(function (event) {
        if (event.type === 'follow') {
          handleFollow(event);
        } else if (event.type === 'message' && event.message.type === 'text') {
          handleTextMessage(event);
        } else if (event.type === 'message') {
          handleNonTextMessage(event);
        }
      });
      return ContentService.createTextOutput('OK');
    }

    // --- アプリからのリクエスト ---
    switch (body.action) {
      case 'notify':        return handleNotify(body);
      case 'cancel':        return handleCancel(body);
      case 'sync':          return handleSync(body);
      case 'saveSettings':  return handleSaveSettings(body);
      default:              return ContentService.createTextOutput('unknown action');
    }
  } catch (err) {
    return ContentService.createTextOutput('error: ' + err.message);
  }
}

// ============================================
// LINE テキストメッセージ処理
// ============================================
function handleTextMessage(event) {
  var lineUserId      = event.source.userId;
  var text            = (event.message.text || '').trim();
  var quotedMessageId = event.message.quotedMessageId || null;
  var replyToken      = event.replyToken;

  // コマンド: URL → アプリリンクをpushで送信
  if (text === 'URL') {
    var userId = getAppUserIdByLineUserId(lineUserId);
    if (userId) {
      var url = APP_URL + '?user=' + userId;
      sendLineMessage(lineUserId, '📱 アプリのURLはこちら：\n' + url);
    } else {
      replyLineMessage(replyToken, 'アカウントが見つかりませんでした。\nLINEで友だち追加してから再度お試しください。');
    }
    return;
  }

  // コマンド: 完了（引用返信）→ タスクを完了にする
  if (text === '完了' && quotedMessageId) {
    handleCompleteByQuote(lineUserId, quotedMessageId, replyToken);
    return;
  }

  // その他のメッセージ → URLと使い方を案内（ブロック解除後の最初のメッセージ対策）
  var uid = getAppUserIdByLineUserId(lineUserId);
  if (uid) {
    var appUrl = APP_URL + '?user=' + uid;
    sendLineMessage(lineUserId,
      '📱 アプリのURLはこちら：\n' + appUrl
      + '\n\n【使えるコマンド】\n'
      + '・「URL」と送信 → このリンクを再表示\n'
      + '・タスク通知に「完了」と引用返信 → タスクを完了にする'
    );
  }
}

// ============================================
// LINE 非テキストメッセージ処理（スタンプ等）
// ============================================
function handleNonTextMessage(event) {
  var lineUserId = event.source.userId;
  var uid = getAppUserIdByLineUserId(lineUserId);
  if (uid) {
    var appUrl = APP_URL + '?user=' + uid;
    sendLineMessage(lineUserId,
      '📱 アプリのURLはこちら：\n' + appUrl
      + '\n\n【使えるコマンド】\n'
      + '・「URL」と送信 → このリンクを再表示\n'
      + '・タスク通知に「完了」と引用返信 → タスクを完了にする'
    );
  }
}

// ============================================
// 引用返信「完了」でタスクを完了にする
// ============================================
function handleCompleteByQuote(lineUserId, quotedMessageId, replyToken) {
  var ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  var msgSheet = ss.getSheetByName(SHEET_LINE_MESSAGES);

  if (!msgSheet) {
    replyLineMessage(replyToken, '該当するタスクが見つかりませんでした。');
    return;
  }

  var msgData    = msgSheet.getDataRange().getValues();
  var taskId     = null;
  var userId     = null;
  var rowToDelete = -1;

  for (var i = 1; i < msgData.length; i++) {
    if (String(msgData[i][0]) === String(quotedMessageId)) {
      taskId      = msgData[i][1];
      userId      = msgData[i][2];
      rowToDelete = i + 1;
      break;
    }
  }

  var appUserId = getAppUserIdByLineUserId(lineUserId);
  if (!taskId || !appUserId || userId !== appUserId) {
    replyLineMessage(replyToken, '該当するタスクが見つかりませんでした。\n（通知から24時間以上経過している場合は無効になります）');
    return;
  }

  var syncSheet = ss.getSheetByName(SHEET_SYNC);
  if (!syncSheet) {
    replyLineMessage(replyToken, 'データが見つかりませんでした。');
    return;
  }

  var syncData = syncSheet.getDataRange().getValues();
  for (var j = 1; j < syncData.length; j++) {
    if (syncData[j][0] === appUserId) {
      var appData;
      try {
        appData = JSON.parse(syncData[j][1]);
      } catch (e) {
        replyLineMessage(replyToken, 'データの読み込みに失敗しました。');
        return;
      }

      var found    = false;
      var taskName = '';
      (appData.tabs || []).forEach(function (tab) {
        (tab.tasks || []).forEach(function (task) {
          if (task.id === taskId && !found) {
            task.done     = true;
            task.progress = 100;
            found         = true;
            taskName      = task.text;
          }
        });
      });

      if (found) {
        syncSheet.getRange(j + 1, 2).setValue(JSON.stringify(appData));
        syncSheet.getRange(j + 1, 3).setValue(new Date().toISOString());
        if (rowToDelete > 0) msgSheet.deleteRow(rowToDelete);
        replyLineMessage(replyToken, '✅ 「' + taskName + '」を完了にしました！');
      } else {
        replyLineMessage(replyToken, 'タスクが見つかりませんでした。すでに削除されているかもしれません。');
      }
      return;
    }
  }

  replyLineMessage(replyToken, 'ユーザーデータが見つかりませんでした。');
}

// ============================================
// フォロー（友だち追加）処理
// ============================================
function handleFollow(event) {
  var lineUserId = event.source.userId;
  var ss         = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet      = getOrCreateSheet(ss, SHEET_USERS,
    ['userId', 'lineUserId', 'createdAt',
     'notify1Time', 'notify1LastSent', 'notify1Enabled',
     'notify2Time', 'notify2Enabled', 'notify2LastSent']);
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][COL_LINE_USER_ID] === lineUserId) {
      sendWelcomeMessage(lineUserId, data[i][COL_USER_ID]);
      return;
    }
  }

  var token = generateToken();
  sheet.appendRow([token, lineUserId, new Date().toISOString(), '', '', false, '', false, '']);
  sendWelcomeMessage(lineUserId, token);
}

// ============================================
// 通知登録（アプリ → GAS）
// ============================================
function handleNotify(body) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, SHEET_NOTIFICATIONS,
    ['userId', 'taskId', 'taskName', 'notifyAt']);
  sheet.appendRow([body.userId, body.taskId, body.taskName, body.notifyAt]);
  return ContentService.createTextOutput('ok');
}

// ============================================
// 通知キャンセル
// ============================================
function handleCancel(body) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NOTIFICATIONS);
  if (!sheet) return ContentService.createTextOutput('ok');

  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === body.userId && data[i][1] === body.taskId) {
      sheet.deleteRow(i + 1);
    }
  }
  return ContentService.createTextOutput('ok');
}

// ============================================
// データ同期（保存）
// ============================================
function handleSync(body) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, SHEET_SYNC, ['userId', 'data', 'updatedAt']);
  var data  = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === body.userId) {
      sheet.getRange(i + 1, 2).setValue(body.data);
      sheet.getRange(i + 1, 3).setValue(body.updatedAt);
      return ContentService.createTextOutput('ok');
    }
  }
  sheet.appendRow([body.userId, body.data, body.updatedAt]);
  return ContentService.createTextOutput('ok');
}

// ============================================
// データ読込
// ============================================
function handleLoad(userId) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_SYNC);
  if (!sheet) return jsonResponse({ data: null });

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      return jsonResponse({ data: data[i][1], updatedAt: data[i][2] });
    }
  }
  return jsonResponse({ data: null });
}

// ============================================
// 設定保存（通知1・通知2）
// ============================================
function handleSaveSettings(body) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, SHEET_USERS,
    ['userId', 'lineUserId', 'createdAt',
     'notify1Time', 'notify1LastSent', 'notify1Enabled',
     'notify2Time', 'notify2Enabled', 'notify2LastSent']);
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][COL_USER_ID] === body.userId) {
      if (body.notify1Time      !== undefined) sheet.getRange(i + 1, COL_NOTIFY1_TIME      + 1).setValue(body.notify1Time);
      if (body.notify1Enabled   !== undefined) sheet.getRange(i + 1, COL_NOTIFY1_ENABLED   + 1).setValue(!!body.notify1Enabled);
      if (body.notify2Time      !== undefined) sheet.getRange(i + 1, COL_NOTIFY2_TIME      + 1).setValue(body.notify2Time);
      if (body.notify2Enabled   !== undefined) sheet.getRange(i + 1, COL_NOTIFY2_ENABLED   + 1).setValue(!!body.notify2Enabled);
      return ContentService.createTextOutput('ok');
    }
  }
  return ContentService.createTextOutput('user not found');
}

// ============================================
// 5分トリガー: タスク通知 + 定時通知 + カウントダウン
// GASのトリガーで checkAndNotify を「5分おき」に設定すること
// ============================================
function checkAndNotify() {
  // 二重実行防止
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(0); // ロック取得できなければ即スキップ
  } catch (e) {
    Logger.log('別の実行が進行中のためスキップ');
    return;
  }

  try {
    var ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
    var uSheet = ss.getSheetByName(SHEET_USERS);
    if (!uSheet) return;

    var users = uSheet.getDataRange().getValues();

    // userId → lineUserId マップ
    var userMap = {};
    for (var j = 1; j < users.length; j++) {
      userMap[users[j][COL_USER_ID]] = users[j][COL_LINE_USER_ID];
    }

    // ① タスク期限通知チェック
    checkTaskNotifications(ss, userMap);

    // ② 定時通知チェック（通知1・通知2）
    checkDailyNotify(ss, uSheet, users, userMap);

    // ③ カウントダウンリマインドチェック
    checkCountdownReminders(ss, users, userMap);

  } finally {
    lock.releaseLock();
  }
}

// ============================================
// ① タスク期限通知チェック
// ============================================
function checkTaskNotifications(ss, userMap) {
  var nSheet = ss.getSheetByName(SHEET_NOTIFICATIONS);
  if (!nSheet) return;

  var now           = new Date();
  var notifications = nSheet.getDataRange().getValues();
  var rowsToDelete  = [];

  for (var i = 1; i < notifications.length; i++) {
    var nUserId   = notifications[i][0];
    var nTaskId   = notifications[i][1];
    var nTaskName = notifications[i][2];
    var notifyAt  = new Date(notifications[i][3]);

    if (notifyAt <= now) {
      var lineUid = userMap[nUserId];
      if (lineUid) {
        var msgText = '🔔 ' + nTaskName + ' の時間です！\n\n'
          + 'このメッセージに「完了」と引用返信するとタスクを完了にできます。';
        var msgId = sendLineMessage(lineUid, msgText);
        if (msgId && nTaskId) {
          saveLineMessageId(ss, msgId, nTaskId, nUserId);
        }
      }
      rowsToDelete.push(i + 1);
    }
  }
  for (var k = rowsToDelete.length - 1; k >= 0; k--) {
    nSheet.deleteRow(rowsToDelete[k]);
  }
}

// ============================================
// ② 定時通知（通知1・通知2）
// ============================================
function checkDailyNotify(ss, uSheet, users, userMap) {
  var syncSheet = ss.getSheetByName(SHEET_SYNC);
  if (!syncSheet) return;

  var now         = new Date();
  var today       = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
  var currentTime = Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm');
  var syncData    = syncSheet.getDataRange().getValues();

  // userId → appData マップ
  var syncMap = {};
  for (var j = 1; j < syncData.length; j++) {
    syncMap[syncData[j][0]] = syncData[j][1];
  }

  for (var i = 1; i < users.length; i++) {
    var userId    = users[i][COL_USER_ID];
    var lineUid   = users[i][COL_LINE_USER_ID];
    if (!lineUid) continue;

    var rawData = syncMap[userId];
    if (!rawData) continue;

    var appData;
    try { appData = JSON.parse(rawData); } catch (e) { continue; }

    // --- 通知1 ---
    var n1Time    = String(users[i][COL_NOTIFY1_TIME]      || '');
    // 後方互換: F列が空でもD列（notify1Time）に値があれば有効扱い
    // F列が明示的に false の場合のみ無効
    var n1RawEnabled = users[i][COL_NOTIFY1_ENABLED];
    var n1Enabled = (n1RawEnabled === false || n1RawEnabled === 'FALSE') ? false : !!n1Time;
    var n1Last    = String(users[i][COL_NOTIFY1_LAST_SENT] || '');

    if (n1Enabled && n1Time && n1Last !== today) {
      if (isInWindow(currentTime, n1Time)) {
        // 先に lastSent を更新して二重送信防止
        uSheet.getRange(i + 1, COL_NOTIFY1_LAST_SENT + 1).setValue(today);
        var msg1 = buildDailyMessage(appData, today, n1Time);
        if (msg1) sendLineMessage(lineUid, msg1);
      }
    }

    // --- 通知2 ---
    var n2Time    = String(users[i][COL_NOTIFY2_TIME]      || '');
    var n2Enabled = !!users[i][COL_NOTIFY2_ENABLED];
    var n2Last    = String(users[i][COL_NOTIFY2_LAST_SENT] || '');

    if (n2Enabled && n2Time && n2Last !== today) {
      if (isInWindow(currentTime, n2Time)) {
        uSheet.getRange(i + 1, COL_NOTIFY2_LAST_SENT + 1).setValue(today);
        var msg2 = buildDailyMessage(appData, today, n2Time);
        if (msg2) sendLineMessage(lineUid, msg2);
      }
    }
  }
}

// 指定時刻の ±0〜+6分ウィンドウ内かどうか判定
// （5分トリガーに対して余裕を持たせた値）
// 指定時刻の 0〜+7分ウィンドウ内か判定（5分トリガーの遅延に余裕を持たせた値）
function isInWindow(currentTime, targetTime) {
  return currentTime >= targetTime && currentTime <= addMinutesToTime(targetTime, 7);
}

// 定時通知メッセージを組み立てる（時間帯で挨拶を変える）
function buildDailyMessage(appData, today, notifyTime) {
  var overdueTasks = [], todayTasks = [], urgentTasks = [];

  (appData.tabs || []).forEach(function (tab) {
    (tab.tasks || []).forEach(function (task) {
      if (task.done) return;
      if (task.dueDate && task.dueDate < today) {
        overdueTasks.push(task.text);
      } else if (task.dueDate === today) {
        todayTasks.push(task.text);
      } else if (task.priority === 'urgent' || task.priority === 'high') {
        urgentTasks.push(task.text);
      }
    });
  });

  if (!overdueTasks.length && !todayTasks.length && !urgentTasks.length) return null;

  // 時間帯で挨拶を変える
  var hour = notifyTime ? parseInt(notifyTime.split(':')[0]) : 8;
  var greeting;
  if (hour < 12) greeting = 'おはようございます！';
  else if (hour < 18) greeting = 'こんにちは！';
  else greeting = 'お疲れ様です！';

  var lines = ['📋 ' + greeting + '本日のToDoまとめです。'];

  if (overdueTasks.length) {
    lines.push('\n⚠️ 期限切れ（' + overdueTasks.length + '件）');
    overdueTasks.slice(0, 5).forEach(function (t) { lines.push('・' + t); });
    if (overdueTasks.length > 5) lines.push('…他' + (overdueTasks.length - 5) + '件');
  }
  if (todayTasks.length) {
    lines.push('\n📅 今日が期限（' + todayTasks.length + '件）');
    todayTasks.slice(0, 5).forEach(function (t) { lines.push('・' + t); });
    if (todayTasks.length > 5) lines.push('…他' + (todayTasks.length - 5) + '件');
  }
  if (urgentTasks.length) {
    lines.push('\n🔥 緊急・高優先度（' + urgentTasks.length + '件）');
    urgentTasks.slice(0, 5).forEach(function (t) { lines.push('・' + t); });
    if (urgentTasks.length > 5) lines.push('…他' + (urgentTasks.length - 5) + '件');
  }

  return lines.join('\n');
}

// ============================================
// ③ カウントダウンリマインド
// ============================================
function checkCountdownReminders(ss, users, userMap) {
  var syncSheet = ss.getSheetByName(SHEET_SYNC);
  if (!syncSheet) return;

  var today    = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var syncData = syncSheet.getDataRange().getValues();

  for (var j = 1; j < syncData.length; j++) {
    var userId     = syncData[j][0];
    var lineUserId = userMap[userId];
    if (!lineUserId) continue;

    var appData;
    try { appData = JSON.parse(syncData[j][1]); } catch (e) { continue; }

    var countdowns = appData.countdowns || [];
    var changed    = false;

    countdowns.forEach(function (cd) {
      if (!cd.reminder || !cd.reminder.enabled) return;

      var daysLeft = calcDaysUntil(cd.date, today);
      if (daysLeft < 0) return;
      if (cd.reminder.lastNotify === today) return;

      var shouldNotify = false;

      if (cd.reminder.dailyFromDays > 0 && daysLeft <= cd.reminder.dailyFromDays) {
        shouldNotify = true;
      }
      if (cd.reminder.milestone && [30, 20, 10].indexOf(daysLeft) !== -1) {
        shouldNotify = true;
      }
      if (cd.reminder.dailyAll) {
        shouldNotify = true;
      }

      if (shouldNotify) {
        var msg = '📅 【カウントダウン通知】\n「' + cd.name + '」まであと ' + daysLeft + '日！\n(' + cd.date + ')';
        sendLineMessage(lineUserId, msg);
        cd.reminder.lastNotify = today;
        changed = true;
      }
    });

    if (changed) {
      appData.countdowns = countdowns;
      syncSheet.getRange(j + 1, 2).setValue(JSON.stringify(appData));
      syncSheet.getRange(j + 1, 3).setValue(new Date().toISOString());
    }
  }
}

// ============================================
// ユーティリティ関数
// ============================================

function calcDaysUntil(dateStr, todayStr) {
  var target = new Date(dateStr + 'T00:00:00');
  var today  = new Date(todayStr + 'T00:00:00');
  return Math.ceil((target - today) / 86400000);
}

function generateToken() {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var token = '';
  for (var i = 0; i < 10; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return token;
}

function sendWelcomeMessage(lineUserId, userId) {
  var url     = APP_URL + '?user=' + userId;
  var message = '🎉 ToDoアプリへようこそ！\n'
    + 'このリンクからアプリを開いてください：\n'
    + url
    + '\n\n【使えるコマンド】\n'
    + '・「URL」と送信 → アプリリンクを再表示\n'
    + '・タスク通知に「完了」と引用返信 → タスクを完了にする';
  sendLineMessage(lineUserId, message);
}

function sendLineMessage(lineUserId, message) {
  try {
    var response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + LINE_TOKEN },
      payload: JSON.stringify({
        to: lineUserId,
        messages: [{ type: 'text', text: message }],
      }),
      muteHttpExceptions: true,
    });
    var result = JSON.parse(response.getContentText());
    if (result.sentMessages && result.sentMessages[0]) {
      return result.sentMessages[0].id;
    }
  } catch (e) {
    Logger.log('sendLineMessage error: ' + e.message);
  }
  return null;
}

function replyLineMessage(replyToken, message) {
  try {
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + LINE_TOKEN },
      payload: JSON.stringify({
        replyToken: replyToken,
        messages: [{ type: 'text', text: message }],
      }),
      muteHttpExceptions: true,
    });
  } catch (e) {
    Logger.log('replyLineMessage error: ' + e.message);
  }
}

function getAppUserIdByLineUserId(lineUserId) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_USERS);
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][COL_LINE_USER_ID] === lineUserId) return data[i][COL_USER_ID];
  }
  return null;
}

function saveLineMessageId(ss, messageId, taskId, userId) {
  var sheet = getOrCreateSheet(ss, SHEET_LINE_MESSAGES,
    ['messageId', 'taskId', 'userId', 'createdAt']);
  sheet.appendRow([messageId, taskId, userId, new Date().toISOString()]);
}

function addMinutesToTime(timeStr, minutes) {
  var parts = timeStr.split(':');
  var h = parseInt(parts[0]);
  var m = parseInt(parts[1]) + minutes;
  if (m >= 60) { h += Math.floor(m / 60); m = m % 60; }
  if (h >= 24) h = 23;
  return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
}

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
