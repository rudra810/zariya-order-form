module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    res.status(405).json({ ok: false, error: 'Method not allowed.' });
    return;
  }

  const appsScriptUrl = process.env.APPS_SCRIPT_URL;
  if (!appsScriptUrl) {
    res.status(500).json({
      ok: false,
      error: 'APPS_SCRIPT_URL is missing in Vercel environment variables.',
    });
    return;
  }

  try {
    const payload =
      typeof req.body === 'string'
        ? req.body
        : JSON.stringify(req.body || {});

    const upstreamResponse = await fetch(appsScriptUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8',
      },
      body: payload,
      redirect: 'follow',
    });

    const text = await upstreamResponse.text();
    let data = null;

    if (text) {
      try {
        data = JSON.parse(text);
      } catch (error) {
        res.status(502).json({ ok: false, error: 'Invalid response from Apps Script.' });
        return;
      }
    }

    if (!upstreamResponse.ok || (data && data.ok === false)) {
      res.status(502).json({
        ok: false,
        error: (data && data.error) || 'Apps Script request failed.',
      });
      return;
    }

    res.status(200).json(data || { ok: true });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: 'Could not connect to Apps Script endpoint.',
    });
  }
};
