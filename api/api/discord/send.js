// api/discord/send.js
//
// Discord へのメッセージ送信 / 編集を中継する汎用エンドポイント。
// GAS はこのエンドポイントだけを叩く（Bot トークンは Vercel 側にのみ置く）。
//
// 認証:
//   リクエストヘッダー x-relay-secret に DISCORD_RELAY_SECRET を入れる
//
// リクエストボディ (JSON):
//   content         : string  省略可  本文
//   embeds          : array   省略可  Embed 配列
//   components      : array   省略可  ボタン等（Phase 2 用）
//   channelId       : string  省略可  省略時は DISCORD_USER_ID への DM に送る
//   messageId       : string  省略可  指定すると新規送信ではなく既存メッセージを編集
//   allowedMentions : object  省略可  省略時はメンション無効
//   （content / embeds / components のいずれか 1 つ以上が必須）
//
// レスポンス:
//   200 { ok: true,  messageId, channelId, edited }
//   4xx { ok: false, error }
//   5xx { ok: false, error, status, discord }

const crypto = require('crypto');

const DISCORD_API = 'https://discord.com/api/v10';

// userId -> DM チャンネル ID。ウォームな実行環境で使い回して往復を 1 回減らす
const dmChannelCache = {};

function timingSafeEqualStr(a, b) {
  const bufA = Buffer.from(String(a), 'utf8');
  const bufB = Buffer.from(String(b), 'utf8');
  if (bufA.length !== bufB.length) return false;
  return crypto.timingSafeEqual(bufA, bufB);
}

function sleep(ms) {
  return new Promise(function (resolve) {
    setTimeout(resolve, ms);
  });
}

async function callDiscord(path, method, body) {
  const res = await fetch(DISCORD_API + path, {
    method: method,
    headers: {
      Authorization: 'Bot ' + process.env.DISCORD_BOT_TOKEN,
      'Content-Type': 'application/json',
    },
    body: body === undefined ? undefined : JSON.stringify(body),
  });

  const text = await res.text();
  let json = null;
  if (text) {
    try {
      json = JSON.parse(text);
    } catch (e) {
      // JSON 以外が返ってきた場合は text のまま扱う
    }
  }
  return { ok: res.ok, status: res.status, json: json, text: text };
}

// 429（レート制限）のときだけ retry_after 待って 1 回だけ再送する
async function callDiscordWithRetry(path, method, body) {
  let r = await callDiscord(path, method, body);
  if (r.status === 429) {
    const retryAfter =
      r.json && typeof r.json.retry_after === 'number' ? r.json.retry_after : 1;
    await sleep(Math.min(retryAfter, 5) * 1000 + 250);
    r = await callDiscord(path, method, body);
  }
  return r;
}

async function getDmChannelId(userId) {
  if (dmChannelCache[userId]) return dmChannelCache[userId];

  const r = await callDiscordWithRetry('/users/@me/channels', 'POST', {
    recipient_id: userId,
  });
  if (!r.ok || !r.json || !r.json.id) {
    const err = new Error('dm_channel_failed');
    err.detail = { status: r.status, discord: r.json || r.text };
    throw err;
  }

  dmChannelCache[userId] = r.json.id;
  return r.json.id;
}

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({ ok: false, error: 'method_not_allowed' });
  }

  const secret = process.env.DISCORD_RELAY_SECRET;
  if (!secret || !process.env.DISCORD_BOT_TOKEN) {
    return res.status(500).json({ ok: false, error: 'server_misconfigured' });
  }

  const given = req.headers['x-relay-secret'];
  if (typeof given !== 'string' || !timingSafeEqualStr(given, secret)) {
    return res.status(401).json({ ok: false, error: 'unauthorized' });
  }

  let body = req.body;
  if (typeof body === 'string') {
    try {
      body = JSON.parse(body);
    } catch (e) {
      body = null;
    }
  }
  if (!body || typeof body !== 'object') {
    return res.status(400).json({ ok: false, error: 'invalid_json' });
  }

  const content = body.content;
  const embeds = body.embeds;
  const components = body.components;
  const channelId = body.channelId;
  const messageId = body.messageId;
  const allowedMentions = body.allowedMentions;

  if (content === undefined && embeds === undefined && components === undefined) {
    return res.status(400).json({ ok: false, error: 'empty_payload' });
  }

  const payload = {};
  if (content !== undefined) payload.content = content;
  if (embeds !== undefined) payload.embeds = embeds;
  if (components !== undefined) payload.components = components;
  payload.allowed_mentions = allowedMentions || { parse: [] };

  try {
    let targetChannelId = channelId;

    if (!targetChannelId) {
      const userId = process.env.DISCORD_USER_ID;
      if (!userId) {
        return res.status(500).json({
          ok: false,
          error: 'no_target',
          message: 'channelId 未指定かつ DISCORD_USER_ID 未設定',
        });
      }
      targetChannelId = await getDmChannelId(userId);
    }

    const path = messageId
      ? '/channels/' + targetChannelId + '/messages/' + messageId
      : '/channels/' + targetChannelId + '/messages';

    const r = await callDiscordWithRetry(path, messageId ? 'PATCH' : 'POST', payload);

    if (!r.ok) {
      return res.status(502).json({
        ok: false,
        error: 'discord_error',
        status: r.status,
        discord: r.json || r.text,
      });
    }

    return res.status(200).json({
      ok: true,
      messageId: r.json && r.json.id,
      channelId: targetChannelId,
      edited: Boolean(messageId),
    });
  } catch (e) {
    return res.status(502).json({
      ok: false,
      error: e.message || 'relay_failed',
      detail: e.detail || null,
    });
  }
};
