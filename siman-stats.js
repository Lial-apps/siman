// Cloudflare Worker "siman-stats" (6.2) — how many devices use סימן, and in which version. Anonymous.
//
// Once a day the app sends: a random number of the device (made on the device, not linked to the email or name),
// the app version, the kind of device (android / ios / windows / mac / other), the language, whether סימן is
// installed or opened in a browser, and whether the person signed up (yes / no). No email, name, file, or IP is kept.
// Each device is counted once a day.
//
// Setup (one time, ~5 minutes):
// 1. Cloudflare → Storage & Databases → D1 SQL Database → Create → name: siman-stats.
// 2. Workers & Pages → Create → Worker → name: siman-stats → Deploy → Edit code → paste this whole file → Deploy.
// 3. The worker → Settings → Bindings → Add → D1 database → Variable name: DB → Database: siman-stats → Deploy.
// 4. The worker → Settings → Variables and Secrets → Add → Secret → name ADMIN_KEY → the same key as siman-register.
// The address becomes https://siman-stats.siman-app.workers.dev — the app and the admin panel already point there.
// The tables are made on their own on the first request.

const ALLOWED = ['https://lial-apps.github.io', 'https://siman.siman-app.workers.dev'];
const PLATFORMS = ['android', 'ios', 'windows', 'mac', 'linux', 'other'];
const clip = (s, n) => String(s ?? '').replace(/[^\w.\-+ ]/g, '').slice(0, n);
const day = (d = new Date()) => d.toISOString().slice(0, 10);

async function tables(db) {
  await db.batch([
    db.prepare('CREATE TABLE IF NOT EXISTS pings (day TEXT NOT NULL, id TEXT NOT NULL, ver TEXT, platform TEXT, lang TEXT, installed INTEGER, reg INTEGER, PRIMARY KEY (day, id))'),
    db.prepare('CREATE TABLE IF NOT EXISTS devices (id TEXT PRIMARY KEY, first TEXT NOT NULL, last TEXT NOT NULL, ver TEXT, platform TEXT, installed INTEGER, reg INTEGER)'),
    db.prepare('CREATE INDEX IF NOT EXISTS pings_day ON pings (day)'),
  ]);
}
let ready = null;

export default {
  async fetch(req, env) {
    const url = new URL(req.url), origin = req.headers.get('Origin') || '';
    const cors = { 'Access-Control-Allow-Origin': ALLOWED.includes(origin) ? origin : ALLOWED[0], 'Access-Control-Allow-Methods': 'POST, OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type', 'Vary': 'Origin' };
    const json = (o, status = 200) => new Response(JSON.stringify(o), { status, headers: { ...cors, 'Content-Type': 'application/json', 'Cache-Control': 'no-store' } });
    if (req.method === 'OPTIONS') return new Response(null, { headers: cors });
    if (!env.DB) return json({ error: 'D1 binding DB is missing (Settings → Bindings)' }, 500);
    if (!ready) ready = tables(env.DB).catch(e => { ready = null; throw e; });
    await ready;
    let body = null; try { body = JSON.parse(await req.text()); } catch (_) {}

    // once a day from the app
    if (url.pathname === '/ping' && req.method === 'POST') {
      if (!ALLOWED.includes(origin)) return new Response('forbidden', { status: 403, headers: cors });
      const id = String(body && body.id || '');
      if (!/^[a-z0-9-]{16,40}$/i.test(id)) return json({ error: 'bad id' }, 400);
      const ver = clip(body.ver, 12), platform = PLATFORMS.includes(body.platform) ? body.platform : 'other';
      const lang = clip(body.lang, 8).toLowerCase(), installed = body.installed ? 1 : 0, reg = body.reg ? 1 : 0, d = day();
      await env.DB.batch([
        env.DB.prepare('INSERT OR IGNORE INTO pings (day, id, ver, platform, lang, installed, reg) VALUES (?, ?, ?, ?, ?, ?, ?)').bind(d, id, ver, platform, lang, installed, reg),
        env.DB.prepare('INSERT INTO devices (id, first, last, ver, platform, installed, reg) VALUES (?, ?, ?, ?, ?, ?, ?) ON CONFLICT(id) DO UPDATE SET last = excluded.last, ver = excluded.ver, platform = excluded.platform, installed = excluded.installed, reg = excluded.reg')
          .bind(id, d, d, ver, platform, installed, reg),
      ]);
      return json({ ok: true });
    }

    // the summary for the admin panel (the same ADMIN_KEY as siman-register)
    if (url.pathname === '/summary' && req.method === 'POST') {
      if (!env.ADMIN_KEY || !body || body.key !== env.ADMIN_KEY) return json({ error: 'forbidden' }, 403);
      const today = day(), d7 = day(new Date(Date.now() - 6 * 864e5)), d30 = day(new Date(Date.now() - 29 * 864e5));
      const one = async (sql, ...a) => (await env.DB.prepare(sql).bind(...a).first()) || {};
      const all = async (sql, ...a) => (await env.DB.prepare(sql).bind(...a).all()).results || [];
      const [total, newToday, act1, act7, act30, series, newSeries, vers, plats, inst, regd] = await Promise.all([
        one('SELECT COUNT(*) n FROM devices'),
        one('SELECT COUNT(*) n FROM devices WHERE first = ?', today),
        one('SELECT COUNT(*) n FROM pings WHERE day = ?', today),
        one('SELECT COUNT(DISTINCT id) n FROM pings WHERE day >= ?', d7),
        one('SELECT COUNT(DISTINCT id) n FROM pings WHERE day >= ?', d30),
        all('SELECT day, COUNT(*) n FROM pings WHERE day >= ? GROUP BY day', d30),
        all('SELECT first day, COUNT(*) n FROM devices WHERE first >= ? GROUP BY first', d30),
        // the version each device used most recently, among the devices of the last 7 days
        all('SELECT ver, COUNT(*) n FROM devices WHERE last >= ? GROUP BY ver ORDER BY n DESC', d7),
        all('SELECT platform, COUNT(*) n FROM devices WHERE last >= ? GROUP BY platform ORDER BY n DESC', d7),
        one('SELECT SUM(installed) n FROM devices WHERE last >= ?', d7),
        one('SELECT SUM(reg) n FROM devices WHERE last >= ?', d7),
      ]);
      return json({ total: total.n || 0, newToday: newToday.n || 0, active: { day: act1.n || 0, week: act7.n || 0, month: act30.n || 0 },
        daily: series, newDaily: newSeries, versions: vers, platforms: plats, installedWeek: inst.n || 0, registeredWeek: regd.n || 0, at: new Date().toISOString() });
    }
    return new Response('siman-stats', { status: 404, headers: cors });
  },
};
