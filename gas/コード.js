function testDiscordDM() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('DISCORD_BOT_TOKEN');
  const userId = props.getProperty('DISCORD_USER_ID');

  if (!token) throw new Error('トークンを取得できません。getPropertyの引数を書き換えていないか確認');
  if (!userId) throw new Error('ユーザーIDを取得できません。getPropertyの引数を書き換えていないか確認');

  const dm = JSON.parse(UrlFetchApp.fetch('https://discord.com/api/v10/users/@me/channels', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bot ' + token },
    payload: JSON.stringify({ recipient_id: userId })
  }).getContentText());

  UrlFetchApp.fetch('https://discord.com/api/v10/channels/' + dm.id + '/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bot ' + token },
    payload: JSON.stringify({ content: '疎通テスト：GAS → Discord DM' })
  });
}