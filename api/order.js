export default async function handler(req, res) {
  const scriptUrl = process.env.GOOGLE_SCRIPT_URL;

  res.setHeader('Cache-Control', 'no-store');

  if (!scriptUrl) {
    return res.status(500).json({ ok: false, reason: 'server_config' });
  }

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 18000);

  try {
    if (req.method === 'GET') {
      const url = new URL(scriptUrl);

      if (req.query?.action) {
        url.searchParams.set('action', req.query.action);
      }

      const upstream = await fetch(url.toString(), {
        method: 'GET',
        redirect: 'follow',
        cache: 'no-store',
        signal: controller.signal
      });

      const text = await upstream.text();

      let data;
      try {
        data = JSON.parse(text);
      } catch {
        return res.status(502).json({
          ok: false,
          reason: 'bad_response'
        });
      }

      return res
        .status(upstream.ok ? 200 : 502)
        .json(data);
    }

    if (req.method === 'POST') {
      const upstream = await fetch(scriptUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'text/plain;charset=utf-8'
        },
        body: JSON.stringify(req.body || {}),
        redirect: 'follow',
        signal: controller.signal
      });

      const text = await upstream.text();

      let data;
      try {
        data = JSON.parse(text);
      } catch {
        return res.status(502).json({
          ok: false,
          reason: 'bad_response'
        });
      }

      return res
        .status(upstream.ok ? 200 : 502)
        .json(data);
    }

    res.setHeader('Allow', 'GET, POST');

    return res.status(405).json({
      ok: false,
      reason: 'method_not_allowed'
    });

  } catch (err) {
    console.error('Order proxy error:', err);

    return res.status(502).json({
      ok: false,
      reason:
        err?.name === 'AbortError'
          ? 'timeout'
          : 'network'
    });

  } finally {
    clearTimeout(timer);
  }
}
