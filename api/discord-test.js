module.exports = async function handler(req, res) {
  const token = process.env.DISCORD_BOT_TOKEN;
  const userId = process.env.DISCORD_USER_ID;
  const headers = {
    Authorization: `Bot ${token}`,
    'Content-Type': 'application/json',
    'User-Agent': 'DiscordBot (https://todo-app-tawny-iota-98.vercel.app, 1.0)'
  };

  const dmRes = await fetch('https://discord.com/api/v10/users/@me/channels', {
    method: 'POST',
    headers,
    body: JSON.stringify({ recipient_id: userId })
  });
  if (!dmRes.ok) {
    return res.status(500).json({ step: 'dm', status: dmRes.status, body: await dmRes.text() });
  }
  const dm = await dmRes.json();

  const msgRes = await fetch(`https://discord.com/api/v10/channels/${dm.id}/messages`, {
    method: 'POST',
    headers,
    body: JSON.stringify({ content: '疎通テスト：Vercel → Discord DM' })
  });
  return res.status(msgRes.ok ? 200 : 500).json({ step: 'message', status: msgRes.status, body: await msgRes.text() });
};
