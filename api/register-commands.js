// スラッシュコマンド登録エンドポイント（Phase 2）
// GASの registerDiscordCommands() から x-relay-secret 付きで叩く。
// Botトークンをローカルに持ち出さずにコマンド登録するための中継。
const crypto = require('crypto');

const DISCORD_API = 'https://discord.com/api/v10';

const COMMANDS = [
  { name: 'today', description: '今日が期限＋期限切れのタスクを表示', type: 1 },
  { name: 'tasks', description: '未完了タスクの一覧を表示', type: 1 },
  { name: 'cleanup', description: '完了済みタスクを一括削除', type: 1 },
];

function timingSafeEqualStr(a, b) {
  const bufA = Buffer.from(String(a), 'utf8');
  const bufB = Buffer.from(String(b), 'utf8');
  if (bufA.length !== bufB.length) return false;
  return crypto.timingSafeEqual(bufA, bufB);
}

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({ ok: false, error: 'method_not_allowed' });
  }

  const secret = process.env.DISCORD_RELAY_SECRET;
  const token = process.env.DISCORD_BOT_TOKEN;
  if (!secret || !token) {
    return res.status(500).json({ ok: false, error: 'server_misconfigured' });
  }

  const given = req.headers['x-relay-secret'];
  if (typeof given !== 'string' || !timingSafeEqualStr(given, secret)) {
    return res.status(401).json({ ok: false, error: 'unauthorized' });
  }

  try {
    const appRes = await fetch(DISCORD_API + '/applications/@me', {
      headers: { Authorization: 'Bot ' + token },
    });
    const app = await appRes.json();
    if (!appRes.ok || !app.id) {
      return res.status(502).json({ ok: false, error: 'app_lookup_failed', discord: app });
    }

    const putRes = await fetch(`${DISCORD_API}/applications/${app.id}/commands`, {
      method: 'PUT',
      headers: {
        Authorization: 'Bot ' + token,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(COMMANDS),
    });
    const result = await putRes.json();
    if (!putRes.ok) {
      return res.status(502).json({ ok: false, error: 'register_failed', discord: result });
    }

    return res.status(200).json({
      ok: true,
      registered: Array.isArray(result) ? result.map((c) => c.name) : result,
    });
  } catch (e) {
    return res.status(502).json({ ok: false, error: e.message || 'register_failed' });
  }
};
