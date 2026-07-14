// Discord Interactions エンドポイント（Phase 2）
// Ed25519署名検証 + PING応答。コマンド/ボタンの実処理はこれから実装する。
// 必要な環境変数: DISCORD_PUBLIC_KEY（Discord Developer Portal > General Information > Public Key）
const crypto = require('crypto');

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
    // 生の32バイト公開鍵をDER(SPKI)ヘッダ付きでKeyObject化（外部ライブラリ不要）
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

  // PING → PONG（エンドポイントURL登録時の検証）
  if (interaction.type === 1) {
    return res.status(200).json({ type: 1 });
  }

  // スラッシュコマンド(2)・ボタン(3)は未実装。エフェメラルで案内だけ返す
  if (interaction.type === 2 || interaction.type === 3) {
    return res.status(200).json({
      type: 4,
      data: { content: '🚧 この操作は準備中です（Phase 2 実装中）', flags: 64 },
    });
  }

  return res.status(400).send('unsupported interaction type');
};

// 署名検証には生のリクエストボディが必要なため、ボディパースを無効化する
module.exports.config = { api: { bodyParser: false } };
