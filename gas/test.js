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