// Discord Interactions エンドポイント（Phase 2）
// - Ed25519署名検証 + PING応答
// - ボタン（✅完了 / +30分 / +1時間 / 今夜21時）: 即ACK(type6) → GAS更新 → 元メッセージをPATCH
// - スラッシュコマンド（/today /tasks /cleanup）: 即ACK(type5) → GASから取得 → PATCH
// 必要な環境変数: DISCORD_PUBLIC_KEY, DISCORD_RELAY_SECRET
const crypto = require('crypto');

let waitUntil = null;
try {
  ({ waitUntil } = require('@vercel/functions'));
} catch (e) {
  // 依存が無い環境ではベストエフォート（応答後の処理継続が保証されない）
}

const DISCORD_API = 'https://discord.com/api/v10';
const GAS_URL = 'https://script.google.com/macros/s/AKfycbyREDlXblHd1ICrEnnrM_CsnQ61jNACJ8nXRaRrnCPfBsunsu3mgM79EOG_MQACgugCRw/exec';

function getRawBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on('data', (c) => chunks.push(c));
    req.on('end', () => resolve(Buffer.concat(chunks)));
    req.on('error', reject);
  });
}

function verifySignature(publicKeyHex, signatureHex, timestamp, rawBody) {
  try {
    const key = crypto.createPublicKey({
      key: Buffer.concat([
        Buffer.from('302a300506032b6570032100', 'hex'),
        Buffer.from(publicKeyHex, 'hex'),
      ]),
      format: 'der',
      type: 'spki',
    });
    return crypto.verify(
      null,
      Buffer.concat([Buffer.from(timestamp), rawBody]),
      key,
      Buffer.from(signatureHex, 'hex')
    );
  } catch (e) {
    return false;
  }
}

// GAS doPost を叩く。bodyに relay secret を同梱して認証する
async function callGAS(body) {
  const res = await fetch(GAS_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=UTF-8' },
    body: JSON.stringify(Object.assign({}, body, { secret: process.env.DISCORD_RELAY_SECRET })),
    redirect: 'follow',
  });
  const text = await res.text();
  try {
    return JSON.parse(text);
  } catch (e) {
    return { ok: false, error: 'gas_bad_response', raw: text.slice(0, 200) };
  }
}

// interactionトークンで元メッセージを編集（Botトークン不要）
async function editOriginal(interaction, payload) {
  const url = `${DISCORD_API}/webhooks/${interaction.application_id}/${interaction.token}/messages/@original`;
  const res = await fetch(url, {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  if (!res.ok) {
    console.error('editOriginal failed', res.status, await res.text());
  }
}

// 今日のJST 21:00までの分数（過ぎていれば明日21:00）
function minutesUntilTonight21() {
  const nowMs = Date.now();
  const jst = new Date(nowMs + 9 * 3600 * 1000);
  let target = Date.UTC(jst.getUTCFullYear(), jst.getUTCMonth(), jst.getUTCDate(), 12, 0, 0); // 21:00 JST = 12:00 UTC
  if (target <= nowMs) target += 24 * 3600 * 1000;
  return Math.round((target - nowMs) / 60000);
}

async function handleComponent(interaction) {
  const parts = String((interaction.data && interaction.data.custom_id) || '').split(':');
  const kind = parts[0];
  const origContent = (interaction.message && interaction.message.content) || '';
  const origEmbeds = (interaction.message && interaction.message.embeds) || [];
  const origComponents = (interaction.message && interaction.message.components) || [];

  if (kind === 'done') {
    const [, userId, taskId] = parts;
    const r = await callGAS({ action: 'discordComplete', userId, taskId });
    if (r && r.ok) {
      await editOriginal(interaction, {
        content: '✅ 完了: ' + (r.taskName || ''),
        embeds: [],
        components: [],
      });
    } else {
      await editOriginal(interaction, {
        content: origContent + '\n⚠️ 完了にできませんでした（' + ((r && r.error) || 'unknown') + '）',
      });
    }
    return;
  }

  if (kind === 'snz' || kind === 'tonight') {
    let userId, taskId, minutes;
    if (kind === 'snz') {
      minutes = parseInt(parts[1], 10);
      userId = parts[2];
      taskId = parts[3];
    } else {
      minutes = minutesUntilTonight21();
      userId = parts[1];
      taskId = parts[2];
    }
    const r = await callGAS({ action: 'discordSnooze', userId, taskId, minutes });
    if (r && r.ok) {
      const unix = Math.floor(new Date(r.notifyAt).getTime() / 1000);
      await editOriginal(interaction, {
        content: '⏰ 再通知します <t:' + unix + ':R>: ' + (r.taskName || ''),
        embeds: origEmbeds,
        components: origComponents, // ボタンは残す（あとで完了にできるように）
      });
    } else {
      await editOriginal(interaction, {
        content: origContent + '\n⚠️ 再通知を設定できませんでした（' + ((r && r.error) || 'unknown') + '）',
      });
    }
    return;
  }

  await editOriginal(interaction, { content: origContent + '\n⚠️ 不明な操作です' });
}

async function handleCommand(interaction) {
  const name = interaction.data && interaction.data.name;
  const r = await callGAS({ action: 'discordCommand', command: name });
  if (r && r.ok) {
    await editOriginal(interaction, {
      content: r.content || '',
      embeds: r.embeds || [],
    });
  } else {
    await editOriginal(interaction, {
      content: '⚠️ 取得に失敗しました（' + ((r && r.error) || 'unknown') + '）',
    });
  }
}

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    return res.status(405).send('Method Not Allowed');
  }

  const publicKey = process.env.DISCORD_PUBLIC_KEY;
  if (!publicKey) {
    return res.status(500).json({ error: 'DISCORD_PUBLIC_KEY not configured' });
  }

  const sig = req.headers['x-signature-ed25519'];
  const ts = req.headers['x-signature-timestamp'];
  const rawBody = await getRawBody(req);

  if (!sig || !ts || !verifySignature(publicKey, sig, ts, rawBody)) {
    return res.status(401).send('invalid request signature');
  }

  let interaction;
  try {
    interaction = JSON.parse(rawBody.toString('utf8'));
  } catch (e) {
    return res.status(400).send('bad request');
  }

  // PING → PONG
  if (interaction.type === 1) {
    return res.status(200).json({ type: 1 });
  }

  // ボタン押下: 即ACK（メッセージ編集の保留）→ 裏で処理
  if (interaction.type === 3) {
    const work = handleComponent(interaction).catch((e) => console.error('component error', e));
    if (waitUntil) waitUntil(work);
    return res.status(200).json({ type: 6 }); // DEFERRED_UPDATE_MESSAGE
  }

  // スラッシュコマンド: 即ACK（考え中表示）→ 裏で処理
  if (interaction.type === 2) {
    const work = handleCommand(interaction).catch((e) => console.error('command error', e));
    if (waitUntil) waitUntil(work);
    return res.status(200).json({ type: 5 }); // DEFERRED_CHANNEL_MESSAGE_WITH_SOURCE
  }

  return res.status(400).send('unsupported interaction type');
};

// 署名検証には生のリクエストボディが必要なため、ボディパースを無効化する
module.exports.config = { api: { bodyParser: false } };
