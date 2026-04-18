export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).send('Method Not Allowed');
  }

  const GAS_URL = 'https://script.google.com/macros/s/AKfycbx2I4SczhacK0icagq39dG6Wx934i7OhDrJw9duhocwJ0JBYdcQsLyJYB0h8t8JxkN38A/exec';

  const bodyStr = JSON.stringify(req.body);

  // デバッグログ（Vercelのログで確認できる）
  console.log('[webhook] received. events:', req.body?.events?.length ?? 0);

  // fetchを先に開始（まだawaitしない）
  const gasPromise = fetch(GAS_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: bodyStr,
    redirect: 'follow',
  }).then(async (r) => {
    const text = await r.text();
    console.log('[webhook] GAS responded. status:', r.status, '/ body:', text.slice(0, 200));
  }).catch(e => {
    console.error('[webhook] GAS fetch failed:', e.message);
  });

  // LINEに200を返す（タイムアウト対策）
  res.status(200).json({ status: 'ok' });

  // 関数を生かしてGASの完了を待つ
  await gasPromise;
}
