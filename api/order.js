export default async function handler(req, res) {
  const scriptUrl = process.env.GOOGLE_SCRIPT_URL;

  if (!scriptUrl) {
    return res.status(500).json({ ok: false, reason: 'server_config' });
  }

  try {
    if (req.method === 'GET') {
      const url = new URL(scriptUrl);
      if (req.query?.action) url.searchParams.set('action', req.query.action);

      const upstream = await fetch(url.toString(), {
        method: 'GET',
        redirect: 'follow',
        cache: 'no-store'
      });

      const text = await upstream.text();
      let data;
      try { data = JSON.parse(text); }
      catch { return res.status(502).json({ ok: false, reason: 'bad_response' }); }

      return res.status(upstream.ok ? 200 : 502).json(data);
    }

    if (req.method === 'POST') {
      const upstream = await fetch(scriptUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(req.body || {}),
        redirect: 'follow'
      });

      const text = await upstream.text();
      let data;
      try { data = JSON.parse(text); }
      catch { return res.status(502).json({ ok: false, reason: 'bad_response' }); }

      return res.status(upstream.ok ? 200 : 502).json(data);
    }

    res.setHeader('Allow', 'GET, POST');
    return res.status(405).json({ ok: false, reason: 'method_not_allowed' });
  } catch (err) {
    console.error('Order proxy error:', err);
    return res.status(502).json({ ok: false, reason: 'network' });
  }
}
