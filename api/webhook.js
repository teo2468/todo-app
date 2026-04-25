module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).send('Method Not Allowed');
  }

  const GAS_URL = 'https://script.google.com/macros/s/AKfycbxspe5iNa8Z2rC1JbQ_AFY_BYLcConcBU3512F7lagdel_tvBnbiTOX8R5bSl_KUYWe9g/exec';

  const bodyStr = JSON.stringify(req.body);
  console.log('[webhook] received. events:', req.body?.events?.length ?? 0);

  const gasController = new AbortController();
  const gasTimeout = setTimeout(() => gasController.abort(), 4500);

  try {
    const r = await fetch(GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: bodyStr,
      redirect: 'follow',
      signal: gasController.signal,
    });
    const text = await r.text();
    console.log('[webhook] GAS status:', r.status, '/ body:', text.slice(0, 200));
  } catch (e) {
    if (e.name === 'AbortError') {
      console.log('[webhook] GAS timeout (4.5s) - proceeding');
    } else {
      console.error('[webhook] GAS error:', e.message);
    }
  } finally {
    clearTimeout(gasTimeout);
  }

  res.status(200).json({ status: 'ok' });
};
