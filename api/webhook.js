export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).send('Method Not Allowed');
  }

  const GAS_URL = 'https://script.google.com/macros/s/AKfycbzCuAXEMR2e3kSppkIFS4OsnaxA9WrYVTyakGHxfmKf35r68niNj4zWQzaBjVlIgxLw/exec';

  // fetchを先に開始（まだawaitしない）
  const gasPromise = fetch(GAS_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(req.body),
    redirect: 'follow',
  }).catch(e => console.error('GAS forward error:', e));

  // LINEに200を返す（タイムアウト対策）
  res.status(200).json({ status: 'ok' });

  // 関数を生かしてGASの完了を待つ
  await gasPromise;
}
