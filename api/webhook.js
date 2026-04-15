export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).send('Method Not Allowed');
  }

  const GAS_URL = 'https://script.google.com/macros/s/AKfycbwYyGXKQXlhusXNFYYHMcxtkJQJTHRlvT4hawOsGnZ7qKsTRMV6QpCgukL4EZJrgODVkw/exec';

  // 先に200を返す（LINEのタイムアウト対策）
  res.status(200).json({ status: 'ok' });

  // 裏でGASに転送（awaitしない）
  fetch(GAS_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(req.body),
    redirect: 'follow',
  }).catch(e => console.error('GAS forward error:', e));
}
