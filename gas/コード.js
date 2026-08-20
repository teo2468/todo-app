// ============================================
// ToDo アプリ GAS コード v12
// 通知基盤: Discord（Vercel中継 /api/discord/send 経由）
// v12: 定時まとめの曜日指定・0件スキップ・期限超過の猶予設定に対応
// ============================================
const SPREADSHEET_ID = '1lIYGcsu_XWzweXHj82M4QM2pCoSLtSiFsBj5hLfzNkg';
const APP_URL = 'https://todo-app-tawny-iota-98.vercel.app';

const SHEET_NOTIFICATIONS = 'notifications';
const SHEET_USERS = 'users';
const SHEET_SYNC = 'sync';

// notifyAtがこれより古い予約は送信せず破棄（滞留した過去予約の一斉送信防止）
const STALE_NOTIFY_HOURS = 6;

// 期限超過アラート: 期限をこの分数以上過ぎたら1回だけ通知（「ちょうど」通知との重複を避ける猶予）
const OVERDUE_ALERT_GRACE_MIN = 30;
// 期限超過アラート: これより古い期限切れは通知せずマークのみ（導入時の一斉送信防止）
const OVERDUE_ALERT_WINDOW_HOURS = 6;

// usersシート列（lineUserId列は未使用だが列構成維持のため残す）
const COL_USER_ID = 0;
const COL_LINE_USER_ID = 1; // 未使用（読み書きしない）
const COL_CREATED_AT = 2;
const COL_NOTIFY1_TIME = 3;
const COL_NOTIFY1_LAST_SENT = 4;
const COL_NOTIFY1_ENABLED = 5;
const COL_NOTIFY2_TIME = 6;
const COL_NOTIFY2_ENABLED = 7;
const COL_NOTIFY2_LAST_SENT = 8;

const USERS_HEADERS = ['userId', 'lineUserId', 'createdAt', 'notify1Time', 'notify1LastSent', 'notify1Enabled',
  'notify2Time', 'notify2Enabled', 'notify2LastSent'];

// Embed色
const COLOR_OVERDUE = 0x992D22;
const COLOR_URGENT = 0xED4245;
const COLOR_HIGH = 0xE67E22;
const COLOR_MEDIUM = 0xFEE75C;
const COLOR_LOW = 0x95A5A6;
const COLOR_DEFAULT = 0x5865F2;

// ============================================
// Discord送信（Vercel中継）
// ============================================
// payload: { content?, embeds?, components?, channelId?, messageId? }
// channelId省略 → Script PropertiesのDISCORD_CHANNEL_IDがあればそのチャンネル宛、
// それも無ければVercel側でDISCORD_USER_ID宛のDMに解決される
// 成功時はmessageIdを返し、失敗時はログに残してnullを返す
function sendDiscordMessage(payload) {
  try {
    const props = PropertiesService.getScriptProperties();
    const secret = props.getProperty('DISCORD_RELAY_SECRET');
    if (!secret) {
      Logger.log('sendDiscordMessage: DISCORD_RELAY_SECRET が未設定');
      return null;
    }
    const channelId = props.getProperty('DISCORD_CHANNEL_ID');
    if (channelId && !payload.channelId) payload = Object.assign({ channelId: channelId }, payload);
    const res = UrlFetchApp.fetch(APP_URL + '/api/discord/send', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-relay-secret': secret },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    if (code === 200) {
      const body = JSON.parse(res.getContentText());
      if (body && body.ok) return body.messageId || null;
    }
    Logger.log('sendDiscordMessage failed: ' + code + ' ' + res.getContentText());
    return null;
  } catch (e) {
    Logger.log('sendDiscordMessage error: ' + e.message);
    return null;
  }
}

// ============================================
// 値の正規化ヘルパー
// ============================================
function getDateString(val) {
  if (val == null || val === '') return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'Asia/Tokyo', 'yyyy-MM-dd');
  const s = String(val).trim();
  if (s.match(/^\d{4}-\d{2}-\d{2}$/)) return s;
  const d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  return s;
}

function getTimeString(val) {
  if (val == null || val === '') return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'Asia/Tokyo', 'HH:mm');
  if (typeof val === 'number') {
    const totalMinutes = Math.round(val * 24 * 60);
    const h = Math.floor(totalMinutes / 60) % 24;
    const m = totalMinutes % 60;
    return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
  }
  const s = String(val).trim();
  const match = s.match(/^(\d{1,2}):(\d{2})$/);
  if (match) return String(parseInt(match[1], 10)).padStart(2, '0') + ':' + match[2];
  const match2 = s.match(/^(\d{1,2}):(\d{2}):\d{2}$/);
  if (match2) return String(parseInt(match2[1], 10)).padStart(2, '0') + ':' + match2[2];
  return s;
}

// ============================================
// Webアプリ エントリポイント
// ============================================
function doGet(e) {
  try {
    const action = e.parameter.action;
    const userId = e.parameter.userId;
    if (action === 'load' && userId) return handleLoad(userId);
    return jsonResponse({ error: 'invalid request' });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    switch (body.action) {
      case 'notify':          return handleNotify(body);
      case 'cancel':          return handleCancel(body);
      case 'sync':            return handleSync(body);
      case 'saveSettings':    return handleSaveSettings(body);
      case 'resetData':       return handleResetData(body);
      // Discord interactions（Vercel /api/interactions から。secret検証あり）
      case 'discordComplete': return handleDiscordComplete(body);
      case 'discordSnooze':   return handleDiscordSnooze(body);
      case 'discordCommand':  return handleDiscordCommand(body);
      default:                return ContentService.createTextOutput('unknown action');
    }
  } catch (err) {
    return ContentService.createTextOutput('error: ' + err.message);
  }
}

// ============================================
// action ハンドラ
// ============================================
function handleNotify(body) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEET_NOTIFICATIONS, ['userId', 'taskId', 'taskName', 'notifyAt']);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === body.userId && data[i][1] === body.taskId) sheet.deleteRow(i + 1);
  }
  // 複数タイミング対応: notifyAts配列を優先し、旧形式notifyAt単体もフォールバックで受ける
  const list = Array.isArray(body.notifyAts) ? body.notifyAts : (body.notifyAt ? [body.notifyAt] : []);
  const nowTs = Date.now();
  list.forEach((iso) => {
    // 過去時刻は登録しない（期限切れタスクの編集・チェック解除で即通知が飛ぶのを防ぐ）
    const t = new Date(iso);
    if (isNaN(t.getTime()) || t.getTime() <= nowTs) return;
    sheet.appendRow([body.userId, body.taskId, body.taskName, iso]);
  });
  return ContentService.createTextOutput('ok');
}

function handleCancel(body) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NOTIFICATIONS);
  if (!sheet) return ContentService.createTextOutput('ok');
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === body.userId && data[i][1] === body.taskId) sheet.deleteRow(i + 1);
  }
  return ContentService.createTextOutput('ok');
}

function handleSync(body) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEET_SYNC, ['userId', 'data', 'updatedAt']);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === body.userId) {
      sheet.getRange(i + 1, 2).setValue(body.data);
      sheet.getRange(i + 1, 3).setValue(body.updatedAt);
      return ContentService.createTextOutput('ok');
    }
  }
  sheet.appendRow([body.userId, body.data, body.updatedAt]);
  return ContentService.createTextOutput('ok');
}

function handleLoad(userId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SYNC);
  if (!sheet) return jsonResponse({ data: null });
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) return jsonResponse({ data: data[i][1], updatedAt: data[i][2] });
  }
  return jsonResponse({ data: null });
}

function handleSaveSettings(body) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEET_USERS, USERS_HEADERS);
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const today = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
  const currentTime = Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm');

  for (let i = 1; i < data.length; i++) {
    if (data[i][COL_USER_ID] !== body.userId) continue;
    const row = i + 1;

    // スロット単位で保存＋即時発火ガード。変更のないスロットには何もしない
    const applySlot = (newTime, newEnabled, timeCol, enabledCol, lastCol) => {
      const oldTime = getTimeString(data[i][timeCol]);
      const oldRawEnabled = data[i][enabledCol];
      const oldEnabled = !(oldRawEnabled === false || String(oldRawEnabled) === 'FALSE') && !!oldTime;

      if (newTime !== undefined) sheet.getRange(row, timeCol + 1).setNumberFormat('@').setValue(String(newTime));
      if (newEnabled !== undefined) sheet.getRange(row, enabledCol + 1).setValue(newEnabled ? 'TRUE' : 'FALSE');

      // 「時刻が変更された」または「OFF→ONに変わった」場合のみ、
      // 新しい時刻が現在時刻以前なら当日送信済み扱いにする（保存直後の即時発火防止）
      const effectiveTime = newTime !== undefined ? getTimeString(String(newTime)) : oldTime;
      const timeChanged = newTime !== undefined && getTimeString(String(newTime)) !== oldTime;
      const turnedOn = newEnabled === true && !oldEnabled;
      if ((timeChanged || turnedOn) && effectiveTime && effectiveTime <= currentTime) {
        sheet.getRange(row, lastCol + 1).setNumberFormat('@').setValue(today);
      }
    };

    applySlot(body.notify1Time, body.notify1Enabled, COL_NOTIFY1_TIME, COL_NOTIFY1_ENABLED, COL_NOTIFY1_LAST_SENT);
    applySlot(body.notify2Time, body.notify2Enabled, COL_NOTIFY2_TIME, COL_NOTIFY2_ENABLED, COL_NOTIFY2_LAST_SENT);
    return ContentService.createTextOutput('ok');
  }
  return ContentService.createTextOutput('user not found');
}

function handleResetData(body) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const userId = body.userId;
  if (!userId) return ContentService.createTextOutput('missing userId');

  const syncSheet = ss.getSheetByName(SHEET_SYNC);
  if (syncSheet) {
    const sd = syncSheet.getDataRange().getValues();
    for (let i = sd.length - 1; i >= 1; i--) if (sd[i][0] === userId) syncSheet.deleteRow(i + 1);
  }

  const nSheet = ss.getSheetByName(SHEET_NOTIFICATIONS);
  if (nSheet) {
    const nd = nSheet.getDataRange().getValues();
    for (let j = nd.length - 1; j >= 1; j--) if (nd[j][0] === userId) nSheet.deleteRow(j + 1);
  }

  const uSheet = ss.getSheetByName(SHEET_USERS);
  if (uSheet) {
    const ud = uSheet.getDataRange().getValues();
    for (let m = 1; m < ud.length; m++) {
      if (ud[m][COL_USER_ID] === userId) {
        const row = m + 1;
        uSheet.getRange(row, COL_NOTIFY1_TIME + 1).setValue('');
        uSheet.getRange(row, COL_NOTIFY1_LAST_SENT + 1).setValue('');
        uSheet.getRange(row, COL_NOTIFY1_ENABLED + 1).setValue('');
        uSheet.getRange(row, COL_NOTIFY2_TIME + 1).setValue('');
        uSheet.getRange(row, COL_NOTIFY2_ENABLED + 1).setValue('');
        uSheet.getRange(row, COL_NOTIFY2_LAST_SENT + 1).setValue('');
        break;
      }
    }
  }

  return ContentService.createTextOutput('ok');
}

// ============================================
// Discord interactions アクション
// Vercel /api/interactions からbody.secret（DISCORD_RELAY_SECRET）付きで呼ばれる
// ============================================
function checkRelaySecret(body) {
  const secret = PropertiesService.getScriptProperties().getProperty('DISCORD_RELAY_SECRET');
  return !!secret && body.secret === secret;
}

// 通知メッセージに付けるボタン行
function taskButtons(userId, taskId) {
  return [{
    type: 1,
    components: [
      { type: 2, style: 3, label: '✅ 完了',   custom_id: 'done:' + userId + ':' + taskId },
      { type: 2, style: 2, label: '+30分',    custom_id: 'snz:30:' + userId + ':' + taskId },
      { type: 2, style: 2, label: '+1時間',   custom_id: 'snz:60:' + userId + ':' + taskId },
      { type: 2, style: 2, label: '今夜21時', custom_id: 'tonight:' + userId + ':' + taskId },
    ],
  }];
}

function deleteNotificationRows(ss, userId, taskId) {
  const sheet = ss.getSheetByName(SHEET_NOTIFICATIONS);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === userId && data[i][1] === taskId) sheet.deleteRow(i + 1);
  }
}

function handleDiscordComplete(body) {
  if (!checkRelaySecret(body)) return jsonResponse({ ok: false, error: 'unauthorized' });

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch (e) { return jsonResponse({ ok: false, error: 'busy' }); }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const syncSheet = ss.getSheetByName(SHEET_SYNC);
    if (!syncSheet) return jsonResponse({ ok: false, error: 'no_data' });

    const syncData = syncSheet.getDataRange().getValues();
    for (let i = 1; i < syncData.length; i++) {
      if (syncData[i][0] !== body.userId) continue;
      let appData;
      try { appData = JSON.parse(syncData[i][1]); }
      catch (e) { return jsonResponse({ ok: false, error: 'broken_data' }); }

      const task = findTask(appData, body.taskId);
      if (!task) return jsonResponse({ ok: false, error: 'task_not_found' });

      task.done = true;
      task.progress = 100;
      task.completedAt = Date.now(); // 完了済み自動削除の起点
      if (task.repeat && task.repeat !== 'none') task.lastCompletedAt = Date.now();

      syncSheet.getRange(i + 1, 2).setValue(JSON.stringify(appData));
      syncSheet.getRange(i + 1, 3).setValue(new Date().toISOString());
      deleteNotificationRows(ss, body.userId, body.taskId);
      return jsonResponse({ ok: true, taskName: task.text });
    }
    return jsonResponse({ ok: false, error: 'user_not_found' });
  } finally {
    lock.releaseLock();
  }
}

function handleDiscordSnooze(body) {
  if (!checkRelaySecret(body)) return jsonResponse({ ok: false, error: 'unauthorized' });
  const minutes = parseInt(body.minutes, 10);
  if (!minutes || minutes <= 0) return jsonResponse({ ok: false, error: 'bad_minutes' });

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch (e) { return jsonResponse({ ok: false, error: 'busy' }); }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const appData = getAppData(ss, body.userId);
    const task = findTask(appData, body.taskId);
    if (!task) return jsonResponse({ ok: false, error: 'task_not_found' });
    if (task.done) return jsonResponse({ ok: false, error: 'already_done' });

    const notifyAt = new Date(Date.now() + minutes * 60 * 1000);
    const sheet = getOrCreateSheet(ss, SHEET_NOTIFICATIONS, ['userId', 'taskId', 'taskName', 'notifyAt']);
    deleteNotificationRows(ss, body.userId, body.taskId);
    sheet.appendRow([body.userId, body.taskId, task.text, notifyAt.toISOString()]);
    return jsonResponse({ ok: true, taskName: task.text, notifyAt: notifyAt.toISOString() });
  } finally {
    lock.releaseLock();
  }
}

function handleDiscordCommand(body) {
  if (!checkRelaySecret(body)) return jsonResponse({ ok: false, error: 'unauthorized' });

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const uSheet = ss.getSheetByName(SHEET_USERS);
  if (!uSheet) return jsonResponse({ ok: false, error: 'no_users' });
  const users = uSheet.getDataRange().getValues();
  if (users.length < 2) return jsonResponse({ ok: false, error: 'no_users' });
  const userId = users[1][COL_USER_ID]; // 実質単一ユーザー運用: 先頭行を使う

  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

  if (body.command === 'cleanup') {
    const lock = LockService.getScriptLock();
    try { lock.waitLock(10000); }
    catch (e) { return jsonResponse({ ok: false, error: 'busy' }); }
    try {
      const syncSheet = ss.getSheetByName(SHEET_SYNC);
      if (!syncSheet) return jsonResponse({ ok: false, error: 'no_data' });
      const syncData = syncSheet.getDataRange().getValues();
      for (let i = 1; i < syncData.length; i++) {
        if (syncData[i][0] !== userId) continue;
        let appData;
        try { appData = JSON.parse(syncData[i][1]); }
        catch (e) { return jsonResponse({ ok: false, error: 'broken_data' }); }
        let count = 0;
        (appData.tabs || []).forEach((tab) => {
          const keep = [];
          (tab.tasks || []).forEach((t) => {
            // 繰り返しタスクは自動リセットで復活する前提のため削除対象外
            if (t.done && (!t.repeat || t.repeat === 'none')) count++;
            else keep.push(t);
          });
          tab.tasks = keep;
        });
        if (count > 0) {
          syncSheet.getRange(i + 1, 2).setValue(JSON.stringify(appData));
          syncSheet.getRange(i + 1, 3).setValue(new Date().toISOString());
        }
        return jsonResponse({ ok: true, content: '🗑 完了済みタスクを ' + count + '件 削除しました（繰り返しタスクは残ります）', embeds: [] });
      }
      return jsonResponse({ ok: false, error: 'user_not_found' });
    } finally {
      lock.releaseLock();
    }
  }

  const appData = getAppData(ss, userId);
  if (!appData) return jsonResponse({ ok: false, error: 'no_data' });

  if (body.command === 'today') {
    const overdue = [], todayTasks = [];
    (appData.tabs || []).forEach((tab) => {
      (tab.tasks || []).forEach((task) => {
        if (task.done) return;
        if (task.dueDate && task.dueDate < today) overdue.push({ task, tabName: tab.name });
        else if (task.dueDate === today) todayTasks.push({ task, tabName: tab.name });
      });
    });
    const total = overdue.length + todayTasks.length;
    if (total === 0) {
      return jsonResponse({ ok: true, content: '✨ 今日が期限・期限切れのタスクはありません！', embeds: [] });
    }
    const sections = [];
    if (overdue.length) {
      sections.push('⚠️ **期限切れ（' + overdue.length + '件）**\n'
        + overdue.map((e) => '・' + formatTaskLine(e.task, today, { tabName: e.tabName })).join('\n'));
    }
    if (todayTasks.length) {
      sections.push('📅 **今日が期限（' + todayTasks.length + '件）**\n'
        + todayTasks.map((e) => '・' + formatTaskLine(e.task, today, { context: 'today', tabName: e.tabName })).join('\n'));
    }
    return jsonResponse({
      ok: true,
      content: '📅 今日のタスク（' + total + '件）',
      embeds: [{ description: sections.join('\n\n'), color: overdue.length ? COLOR_URGENT : COLOR_DEFAULT }],
    });
  }

  if (body.command === 'tasks') {
    const MAX_LINES = 30;
    let total = 0;
    const sections = [];
    let lines = 0;
    (appData.tabs || []).forEach((tab) => {
      const undone = (tab.tasks || []).filter((t) => !t.done);
      total += undone.length;
      if (!undone.length || lines >= MAX_LINES) return;
      const take = undone.slice(0, MAX_LINES - lines);
      lines += take.length;
      sections.push('**' + tab.name + '（' + undone.length + '件）**\n'
        + take.map((t) => '・' + formatTaskLine(t, today, {})).join('\n'));
    });
    if (total === 0) {
      return jsonResponse({ ok: true, content: '✨ 未完了のタスクはありません！', embeds: [] });
    }
    let desc = sections.join('\n\n');
    if (total > lines) desc += '\n\n…他 ' + (total - lines) + '件はアプリで確認してください';
    return jsonResponse({
      ok: true,
      content: '📋 未完了タスク（' + total + '件）',
      embeds: [{ description: desc, color: COLOR_DEFAULT }],
    });
  }

  return jsonResponse({ ok: false, error: 'unknown_command' });
}

// ============================================
// 定期トリガー本体
// ============================================
function checkAndNotify() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    Logger.log('ロック取得失敗（15秒超過）スキップ');
    return;
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const uSheet = ss.getSheetByName(SHEET_USERS);
    if (!uSheet) return;
    const users = uSheet.getDataRange().getValues();

    checkTaskNotifications(ss);
    checkDailyNotify(ss, uSheet, users);
    checkCountdownReminders(ss);
    checkOverdueAlerts(ss);
  } finally {
    lock.releaseLock();
  }
}

// ============================================
// 個別タスク通知（notificationsシート発火）
// ============================================
function checkTaskNotifications(ss) {
  const nSheet = ss.getSheetByName(SHEET_NOTIFICATIONS);
  if (!nSheet) return;

  const now = new Date();
  const today = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
  const staleCutoff = now.getTime() - STALE_NOTIFY_HOURS * 60 * 60 * 1000;
  const rows = nSheet.getDataRange().getValues();
  const rowsToDelete = [];
  const appDataCache = {};

  for (let i = 1; i < rows.length; i++) {
    const userId = rows[i][0];
    const taskId = rows[i][1];
    const taskName = rows[i][2];
    const notifyAt = new Date(rows[i][3]);

    if (isNaN(notifyAt.getTime())) { rowsToDelete.push(i + 1); continue; }
    if (notifyAt > now) continue;

    // 古すぎる予約は送信せず破棄
    if (notifyAt.getTime() < staleCutoff) { rowsToDelete.push(i + 1); continue; }

    // 送信前にsyncデータで実在・未完了を確認
    if (!(userId in appDataCache)) appDataCache[userId] = getAppData(ss, userId);
    const appData = appDataCache[userId];
    const task = findTask(appData, taskId);
    if (!task || task.done) { rowsToDelete.push(i + 1); continue; }

    const tabName = getTabNameForTask(appData, taskId);
    const msgId = sendDiscordMessage({
      content: '🔔 ' + (task.text || taskName) + ' の時間です',
      embeds: [buildTaskEmbed(task, today, tabName)],
      components: taskButtons(userId, taskId),
    });
    // 送信失敗時は行を残して次回トリガーで再試行（古くなればSTALEで破棄される）
    if (msgId) rowsToDelete.push(i + 1);
  }

  for (let k = rowsToDelete.length - 1; k >= 0; k--) nSheet.deleteRow(rowsToDelete[k]);
}

// ============================================
// 定時まとめ通知
// ============================================
function checkDailyNotify(ss, uSheet, users) {
  const syncSheet = ss.getSheetByName(SHEET_SYNC);
  if (!syncSheet) return;

  const now = new Date();
  const today = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
  const currentTime = Utilities.formatDate(now, 'Asia/Tokyo', 'HH:mm');
  const syncData = syncSheet.getDataRange().getValues();

  const syncMap = {};
  for (let j = 1; j < syncData.length; j++) syncMap[syncData[j][0]] = syncData[j][1];

  const slots = [
    { timeCol: COL_NOTIFY1_TIME, enabledCol: COL_NOTIFY1_ENABLED, lastCol: COL_NOTIFY1_LAST_SENT, daysKey: 'notify1Days' },
    { timeCol: COL_NOTIFY2_TIME, enabledCol: COL_NOTIFY2_ENABLED, lastCol: COL_NOTIFY2_LAST_SENT, daysKey: 'notify2Days' },
  ];

  // JSTの曜日（0=日〜6=土）
  const dowJst = parseInt(Utilities.formatDate(now, 'Asia/Tokyo', 'u'), 10) % 7;

  for (let i = 1; i < users.length; i++) {
    const userId = users[i][COL_USER_ID];
    const rawData = syncMap[userId];
    if (!rawData) continue;

    let appData;
    try { appData = JSON.parse(rawData); } catch (e) { continue; }

    slots.forEach((slot) => {
      const time = getTimeString(users[i][slot.timeCol]);
      const rawEnabled = users[i][slot.enabledCol];
      const enabled = (rawEnabled === false || String(rawEnabled) === 'FALSE') ? false : !!time;
      const last = getDateString(users[i][slot.lastCol]);

      // 設定時刻を過ぎている かつ 当日未送信なら送信（3分窓方式は廃止）
      if (!enabled || !time || last === today || currentTime < time) return;

      // 曜日指定: 配列があり今日が含まれない場合はスキップ（lastSentは触らない）
      const days = (appData.settings || {})[slot.daysKey];
      if (Array.isArray(days) && days.length > 0 && days.indexOf(dowJst) === -1) return;

      // fresh再読込＋送信前lastSentセットで二重送信ガード
      const fresh = getDateString(uSheet.getRange(i + 1, slot.lastCol + 1).getValue());
      if (fresh === today) return;
      uSheet.getRange(i + 1, slot.lastCol + 1).setNumberFormat('@').setValue(today);
      SpreadsheetApp.flush();

      const msg = buildDailySummary(appData, today, time);
      // 0件スキップ設定: 当日送信済み扱いのまま送らない
      if (msg.count === 0 && (appData.settings || {}).skipEmptySummary) return;
      sendDiscordMessage({ content: msg.content, embeds: [msg.embed] });
    });
  }
}

// ============================================
// カウントダウン通知
// ============================================
function checkCountdownReminders(ss) {
  const syncSheet = ss.getSheetByName(SHEET_SYNC);
  if (!syncSheet) return;

  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  const syncData = syncSheet.getDataRange().getValues();

  for (let j = 1; j < syncData.length; j++) {
    let appData;
    try { appData = JSON.parse(syncData[j][1]); } catch (e) { continue; }

    const countdowns = appData.countdowns || [];
    let changed = false;
    const bundle = [];

    countdowns.forEach((cd) => {
      if (!cd.reminder || !cd.reminder.enabled) return;
      const daysLeft = calcDaysUntil(cd.date, today);
      if (daysLeft < 0) return;
      if (cd.reminder.lastNotify === today) return;

      let shouldNotify = false;
      if (cd.reminder.dailyFromDays > 0 && daysLeft <= cd.reminder.dailyFromDays) shouldNotify = true;
      if (cd.reminder.milestone && [30, 20, 10, 7].indexOf(daysLeft) !== -1) shouldNotify = true;
      if (cd.reminder.dailyAll) shouldNotify = true;

      if (shouldNotify) {
        bundle.push({ name: cd.name, daysLeft: daysLeft, date: cd.date });
        cd.reminder.lastNotify = today;
        changed = true;
      }
    });

    if (bundle.length > 0) {
      bundle.sort((a, b) => a.daysLeft - b.daysLeft);
      const lines = bundle.map((b) =>
        '🎯 「' + b.name + '」まであと ' + b.daysLeft + '日（' + formatMonthDay(b.date) + '）');
      sendDiscordMessage({
        content: '📅 カウントダウン',
        embeds: [{ description: lines.join('\n'), color: COLOR_DEFAULT }],
      });
    }

    if (changed) {
      appData.countdowns = countdowns;
      syncSheet.getRange(j + 1, 2).setValue(JSON.stringify(appData));
      syncSheet.getRange(j + 1, 3).setValue(new Date().toISOString());
    }
  }
}

// ============================================
// 期限超過アラート
// 時刻つき期限を過ぎた未完了タスクに1回だけ通知する。
// 送信済みマークはappData内のtask.overdueAlertedに持つ（期限変更時はアプリ側でクリア）
// ============================================
function checkOverdueAlerts(ss) {
  const syncSheet = ss.getSheetByName(SHEET_SYNC);
  if (!syncSheet) return;

  const now = new Date();
  const nowTs = now.getTime();
  const today = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
  const syncData = syncSheet.getDataRange().getValues();

  for (let j = 1; j < syncData.length; j++) {
    const alertUserId = syncData[j][0];
    let appData;
    try { appData = JSON.parse(syncData[j][1]); } catch (e) { continue; }

    const settings = appData.settings || {};
    if (settings.overdueAlert === false) continue;

    let changed = false;
    (appData.tabs || []).forEach((tab) => {
      (tab.tasks || []).forEach((task) => {
        if (task.done || !task.dueDate || !task.dueTime) return;
        if (task.overdueAlerted) return;
        const dueTs = new Date(task.dueDate + 'T' + getTimeString(task.dueTime) + ':00+09:00').getTime();
        if (isNaN(dueTs)) return;
        const overdueMs = nowTs - dueTs;
        const graceMin = (typeof settings.overdueGraceMin === 'number') ? settings.overdueGraceMin : OVERDUE_ALERT_GRACE_MIN;
        if (overdueMs < graceMin * 60 * 1000) return;

        // 窓内なら送信。古すぎるものはマークのみ（一斉送信防止）
        if (overdueMs <= OVERDUE_ALERT_WINDOW_HOURS * 60 * 60 * 1000) {
          const embed = buildTaskEmbed(task, today, tab.name);
          embed.color = COLOR_OVERDUE;
          sendDiscordMessage({
            content: '⏰ 期限を過ぎています: ' + task.text,
            embeds: [embed],
            components: taskButtons(alertUserId, task.id),
          });
        }
        task.overdueAlerted = now.toISOString();
        changed = true;
      });
    });

    if (changed) {
      syncSheet.getRange(j + 1, 2).setValue(JSON.stringify(appData));
      syncSheet.getRange(j + 1, 3).setValue(new Date().toISOString());
    }
  }
}

// ============================================
// 通知カテゴリ・優先度ヘルパー
// ============================================
const DEFAULT_CATEGORIES = ['overdue', 'today', 'future', 'highPriority'];
const CATEGORY_INFO = {
  overdue:      { icon: '⚠️', label: '期限切れ' },
  today:        { icon: '📅', label: '今日が期限' },
  future:       { icon: '📅', label: '明日以降の期限' },
  highPriority: { icon: '🔥', label: '高優先度（期限なし）' },
};

function categorizeTask(task, today) {
  if (task.done) return null;
  if (task.dueDate) {
    if (task.dueDate < today) return 'overdue';
    if (task.dueDate === today) return 'today';
    return 'future';
  }
  if (task.priority === 'urgent' || task.priority === 'high') return 'highPriority';
  return null;
}

function passesPriorityFilter(task, minPriority) {
  if (!minPriority || minPriority === 'all') return true;
  const p = task.priority || 'none';
  if (minPriority === 'medium') return p === 'urgent' || p === 'high' || p === 'medium';
  if (minPriority === 'high') return p === 'urgent' || p === 'high';
  return true;
}

function getNotifiableTasksByCategory(appData, today) {
  const settings = appData.settings || {};
  const enabledCats = settings.notifyCategories || DEFAULT_CATEGORIES;
  const minPriority = settings.notifyMinPriority || 'all';
  const grouped = { overdue: [], today: [], future: [], highPriority: [] };

  (appData.tabs || []).forEach((tab) => {
    if (tab.notifyExclude) return;
    (tab.tasks || []).forEach((task) => {
      const cat = categorizeTask(task, today);
      if (!cat) return;
      if (enabledCats.indexOf(cat) === -1) return;
      if (!passesPriorityFilter(task, minPriority)) return;
      grouped[cat].push(task);
    });
  });
  return grouped;
}

// ============================================
// 整形ヘルパー
// ============================================
const WEEKDAYS_JP_GAS = ['日', '月', '火', '水', '木', '金', '土'];

function formatDateWithDay(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr + 'T00:00:00');
  if (isNaN(d.getTime())) return dateStr;
  return (d.getMonth() + 1) + '/' + d.getDate() + '(' + WEEKDAYS_JP_GAS[d.getDay()] + ')';
}

function formatMonthDay(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  if (isNaN(d.getTime())) return dateStr;
  return (d.getMonth() + 1) + '/' + d.getDate();
}

function calcDaysUntil(dateStr, todayStr) {
  const target = new Date(dateStr + 'T00:00:00');
  const today = new Date(todayStr + 'T00:00:00');
  return Math.ceil((target - today) / 86400000);
}

// dueDate＋dueTime（JST）→ unix秒。Discordの <t:unix:R> 用
function toUnixSeconds(dateStr, timeStr) {
  const d = new Date(dateStr + 'T' + getTimeString(timeStr) + ':00+09:00');
  if (isNaN(d.getTime())) return null;
  return Math.floor(d.getTime() / 1000);
}

// 期限表示: 時刻つきは <t:unix:R>（自動更新の相対表示）、日付のみはv8の日本語表記
function formatDueDisplay(task, today) {
  if (!task.dueDate) return '';
  if (task.dueTime) {
    const unix = toUnixSeconds(task.dueDate, task.dueTime);
    if (unix != null) return '<t:' + unix + ':R>';
  }
  if (task.dueDate < today) return formatDateWithDay(task.dueDate);
  if (task.dueDate === today) return '今日';
  const days = calcDaysUntil(task.dueDate, today);
  if (days === 1) return '明日';
  if (days <= 7) return 'あと' + days + '日';
  return formatDateWithDay(task.dueDate);
}

function getGreeting(hour) {
  if (hour >= 6 && hour <= 10) return 'おはようございます！';
  if (hour >= 11 && hour <= 17) return 'こんにちは！';
  if (hour >= 18 && hour <= 20) return 'こんばんは！';
  return 'お疲れ様です！';
}

function getGreetingEmoji(hour) {
  if (hour >= 6 && hour <= 17) return '🌞';
  return '🌜';
}

// タスク1行分（期限 + 優先度絵文字 + 名前 + タブ名）
// opts.context: 'today'等のカテゴリ名を渡すと「今日」プレフィックスを省略
function formatTaskLine(task, today, opts) {
  opts = opts || {};
  const parts = [];

  if (task.dueDate) {
    if (task.dueTime) {
      const unix = toUnixSeconds(task.dueDate, task.dueTime);
      parts.push(unix != null ? '<t:' + unix + ':R>' : task.dueDate + ' ' + task.dueTime);
    } else if (task.dueDate < today) {
      parts.push(formatDateWithDay(task.dueDate));
    } else if (task.dueDate === today) {
      if (opts.context !== 'today') parts.push('今日');
    } else {
      const daysAhead = calcDaysUntil(task.dueDate, today);
      if (daysAhead === 1) parts.push('明日');
      else if (daysAhead <= 7) parts.push('あと' + daysAhead + '日');
      else parts.push(formatDateWithDay(task.dueDate));
    }
  }

  if (task.priority === 'urgent') parts.push('🔴');
  else if (task.priority === 'high') parts.push('🟠');

  let name = task.text;
  if (task.progress > 0 && task.progress < 100) name += '（進捗' + task.progress + '%）';
  parts.push(name);

  if (opts.tabName) parts.push('(' + opts.tabName + ')');

  return parts.join(' ');
}

// Embed色: 期限切れ判定を優先度より優先する
function getTaskColor(task, today) {
  if (task.dueDate && task.dueDate < today) return COLOR_OVERDUE;
  switch (task.priority) {
    case 'urgent': return COLOR_URGENT;
    case 'high':   return COLOR_HIGH;
    case 'medium': return COLOR_MEDIUM;
    case 'low':    return COLOR_LOW;
    default:       return COLOR_DEFAULT;
  }
}

// 個別タスク通知のembed
function buildTaskEmbed(task, today, tabName) {
  let title = task.text;
  if (task.priority === 'urgent') title = '🔴 ' + title;
  else if (task.priority === 'high') title = '🟠 ' + title;

  const fields = [];
  const due = formatDueDisplay(task, today);
  if (due) fields.push({ name: '期限', value: due, inline: true });
  if (tabName) fields.push({ name: 'タブ', value: tabName, inline: true });
  if (task.progress > 0 && task.progress < 100) {
    fields.push({ name: '進捗', value: task.progress + '%', inline: true });
  }

  return { title: title, fields: fields, color: getTaskColor(task, today) };
}

// 定時まとめの content + embed
function buildDailySummary(appData, today, notifyTime) {
  const settings = appData.settings || {};
  const enabledCats = settings.notifyCategories || DEFAULT_CATEGORIES;
  const grouped = getNotifiableTasksByCategory(appData, today);

  const hour = notifyTime ? parseInt(notifyTime.split(':')[0], 10) : 8;
  const greeting = getGreetingEmoji(hour) + ' ' + getGreeting(hour);

  let totalCount = 0;
  enabledCats.forEach((k) => { if (grouped[k]) totalCount += grouped[k].length; });

  if (totalCount === 0) {
    return {
      count: 0,
      content: '📋 本日のまとめ（0件）',
      embed: {
        title: greeting,
        description: 'やるべきタスクはありません。\nいい調子です！',
        color: COLOR_DEFAULT,
      },
    };
  }

  const sections = [];
  DEFAULT_CATEGORIES.forEach((cat) => {
    if (enabledCats.indexOf(cat) === -1) return;
    const arr = grouped[cat] || [];
    if (!arr.length) return;
    const info = CATEGORY_INFO[cat];
    const lines = [info.icon + ' **' + info.label + '（' + arr.length + '件）**'];
    arr.forEach((t) => lines.push('・' + formatTaskLine(t, today, { context: cat })));
    sections.push(lines.join('\n'));
  });

  const overdueCount = enabledCats.indexOf('overdue') !== -1 ? grouped.overdue.length : 0;

  return {
    count: totalCount,
    content: '📋 本日のまとめ（' + totalCount + '件）',
    embed: {
      title: greeting,
      description: sections.join('\n\n'),
      color: overdueCount > 0 ? COLOR_URGENT : COLOR_DEFAULT,
    },
  };
}

// ============================================
// データ取得ヘルパー
// ============================================
function getAppData(ss, appUserId) {
  const syncSheet = ss.getSheetByName(SHEET_SYNC);
  if (!syncSheet) return null;
  const syncData = syncSheet.getDataRange().getValues();
  for (let i = 1; i < syncData.length; i++) {
    if (syncData[i][0] === appUserId) {
      try { return JSON.parse(syncData[i][1]); } catch (e) { return null; }
    }
  }
  return null;
}

function findTask(appData, taskId) {
  if (!appData) return null;
  let found = null;
  (appData.tabs || []).forEach((tab) => {
    (tab.tasks || []).forEach((task) => {
      if (task.id === taskId && !found) found = task;
    });
  });
  return found;
}

function getTabNameForTask(appData, taskId) {
  if (!appData) return null;
  let found = null;
  (appData.tabs || []).forEach((tab) => {
    (tab.tasks || []).forEach((task) => {
      if (task.id === taskId && !found) found = tab.name;
    });
  });
  return found;
}

// ============================================
// 汎用ヘルパー
// ============================================
function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
