// DISCORD_CHANNEL_ID設定後の疎通確認用。コード.jsのsendDiscordMessageを経由するので
// チャンネルIDの解決含め本番と同じ経路で送信される
function testChannelNotify() {
  const channelId = PropertiesService.getScriptProperties().getProperty('DISCORD_CHANNEL_ID');
  Logger.log('DISCORD_CHANNEL_ID: ' + (channelId || '(未設定 → DM宛になります)'));
  const msgId = sendDiscordMessage({
    content: '📢 チャンネル通知テスト',
    embeds: [{ description: 'この通知が目的のチャンネルに届いていれば設定完了です', color: 0x5865f2 }],
  });
  Logger.log(msgId ? '送信成功 messageId=' + msgId : '送信失敗（ログを確認）');
}

function testDiscordRelay() {
  const res = UrlFetchApp.fetch('https://todo-app-tawny-iota-98.vercel.app/api/discord/send', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-relay-secret': PropertiesService.getScriptProperties().getProperty('DISCORD_RELAY_SECRET') },
    payload: JSON.stringify({
      content: 'リレー疎通テスト',
      embeds: [{ title: '中継エンドポイント', description: 'GAS → Vercel → Discord', color: 0x5865f2 }]
    }),
    muteHttpExceptions: true
  });
  Logger.log(res.getResponseCode());
  Logger.log(res.getContentText());
}