// Cloudflare Worker "siman-register" — the sign-up list of סימן
//
// Setup (one time, ~5 minutes):
// 1. Cloudflare → Workers & Pages → Create → Worker → name: siman-register → Deploy → Edit code → paste this file → Deploy.
// 2. Storage & Databases → KV → Create namespace: SIMAN_USERS.
// 3. The worker → Settings → Bindings → Add → KV namespace → Variable name: USERS → namespace: SIMAN_USERS.
// 4. The worker → Settings → Variables and Secrets → Add → Secret → name: ADMIN_KEY → a long password of your own.
//
// The list (opens in Excel): https://siman-register.lioralmo1.workers.dev/list?key=<ADMIN_KEY>

const ALLOWED = ['https://lial-apps.github.io'];

export default {
  async fetch(req, env) {
    const url = new URL(req.url);
    const origin = req.headers.get('Origin') || '';
    const cors = {
      'Access-Control-Allow-Origin': ALLOWED.includes(origin) ? origin : ALLOWED[0],
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
      'Vary': 'Origin'
    };
    if (req.method === 'OPTIONS') return new Response(null, { headers: cors });

    // a sign-up from the app
    if (url.pathname === '/register' && req.method === 'POST') {
      if (!ALLOWED.includes(origin)) return new Response('forbidden', { status: 403, headers: cors });
      let d;
      try { d = await req.json(); } catch { return new Response('bad request', { status: 400, headers: cors }); }
      const mail = String(d.mail || '').trim().toLowerCase().slice(0, 200);
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(mail)) return new Response('bad mail', { status: 400, headers: cors });
      const key = 'u:' + mail;
      const old = await env.USERS.get(key, 'json');
      const now = new Date().toISOString();
      const rec = {
        mail,
        name: String(d.name || '').slice(0, 60),
        news: !!d.news,
        platform: String(d.platform || '').slice(0, 20),
        lang: String(d.lang || '').slice(0, 20),
        first: old ? old.first : now,
        last: now
      };
      await env.USERS.put(key, JSON.stringify(rec));
      return new Response(JSON.stringify({ ok: true }), { headers: { ...cors, 'Content-Type': 'application/json' } });
    }

    // the whole list as a CSV file, for you only
    if (url.pathname === '/list' && req.method === 'GET') {
      if (!env.ADMIN_KEY || url.searchParams.get('key') !== env.ADMIN_KEY) return new Response('forbidden', { status: 403 });
      const rows = [['mail', 'name', 'news', 'platform', 'lang', 'first', 'last']];
      let cursor;
      do {
        const r = await env.USERS.list({ prefix: 'u:', cursor });
        for (const k of r.keys) {
          const v = await env.USERS.get(k.name, 'json');
          if (v) rows.push([v.mail, v.name, v.news ? 'yes' : 'no', v.platform, v.lang, v.first, v.last]);
        }
        cursor = r.list_complete ? null : r.cursor;
      } while (cursor);
      const csv = '\uFEFF' + rows.map(r => r.map(x => '"' + String(x ?? '').replace(/"/g, '""') + '"').join(',')).join('\r\n');
      return new Response(csv, { headers: { 'Content-Type': 'text/csv; charset=utf-8', 'Content-Disposition': 'attachment; filename="siman-users.csv"' } });
    }

    return new Response('siman-register', { status: 404 });
  }
};
