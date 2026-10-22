// Cloudflare Worker "siman-register" — the sign-up list of סימן (version for app 4.7)
//
// What's new in this version:
//   • every user gets a 4-month launch period, counted from sign-up
//     (people who signed up before this version: counted from the first time they open 4.7)
//   • /status — the app checks the launch period once a day and sends how many times each feature was used
//   • /list now also shows the launch period dates, active days and feature use per user
//   • /stats — a summary: how many times each feature was used, and by how many users
//
// Deploy: Cloudflare → Workers & Pages → siman-register → Edit code → select all, paste this file → Deploy.
// The existing settings stay as they are: KV binding USERS, secret ADMIN_KEY.
//
// The list (opens in Excel):  https://siman-register.siman-app.workers.dev/list?key=<ADMIN_KEY>
// Feature summary (Excel):    https://siman-register.siman-app.workers.dev/stats?key=<ADMIN_KEY>

const ALLOWED = ['https://lial-apps.github.io'];
const TRIAL_MONTHS = 4;

function addMonths(iso, n) {
  const d = new Date(iso);
  const day = d.getUTCDate();
  d.setUTCDate(1);
  d.setUTCMonth(d.getUTCMonth() + n);
  const last = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 0)).getUTCDate();
  d.setUTCDate(Math.min(day, last));
  return d.toISOString();
}
const validMail = m => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(m);
const cleanMail = m => String(m || '').trim().toLowerCase().slice(0, 200);

async function readBody(req) {
  // the app sends text/plain (also used by sendBeacon), older versions send application/json
  try { return JSON.parse(await req.text()); } catch { return null; }
}

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
    const json = (o, status = 200) => new Response(JSON.stringify(o), { status, headers: { ...cors, 'Content-Type': 'application/json' } });
    if (req.method === 'OPTIONS') return new Response(null, { headers: cors });

    // a sign-up (or a settings change) from the app
    if (url.pathname === '/register' && req.method === 'POST') {
      if (!ALLOWED.includes(origin)) return new Response('forbidden', { status: 403, headers: cors });
      const d = await readBody(req);
      if (!d) return new Response('bad request', { status: 400, headers: cors });
      const mail = cleanMail(d.mail);
      if (!validMail(mail)) return new Response('bad mail', { status: 400, headers: cors });
      const key = 'u:' + mail;
      const old = await env.USERS.get(key, 'json');
      const now = new Date().toISOString();
      const rec = {
        ...(old || {}),
        mail,
        name: String(d.name || '').slice(0, 60),
        news: !!d.news,
        platform: String(d.platform || '').slice(0, 20),
        lang: String(d.lang || '').slice(0, 20),
        first: old ? old.first : now,
        last: now,
        trialStart: (old && old.trialStart) || now
      };
      await env.USERS.put(key, JSON.stringify(rec));
      return json({ ok: true, trialStart: rec.trialStart, trialEnd: addMonths(rec.trialStart, TRIAL_MONTHS) });
    }

    // once a day from the app: launch-period status + feature counts
    if (url.pathname === '/status' && req.method === 'POST') {
      if (!ALLOWED.includes(origin)) return new Response('forbidden', { status: 403, headers: cors });
      const d = await readBody(req);
      if (!d) return new Response('bad request', { status: 400, headers: cors });
      const mail = cleanMail(d.mail);
      if (!validMail(mail)) return new Response('bad mail', { status: 400, headers: cors });
      const key = 'u:' + mail;
      const now = new Date().toISOString();
      const day = now.slice(0, 10);
      const rec = (await env.USERS.get(key, 'json')) || { mail, first: now };
      if (!rec.trialStart) rec.trialStart = now;
      rec.last = now;
      rec.ver = String(d.ver || '').slice(0, 12);
      if (rec.lastDay !== day) { rec.days = (rec.days || 0) + 1; rec.lastDay = day; }
      const usage = rec.usage || {};
      const counts = d.counts && typeof d.counts === 'object' ? d.counts : {};
      let n = 0;
      for (const [k, v] of Object.entries(counts)) {
        if (n++ > 80) break;
        const name = String(k).replace(/[^a-z0-9_]/gi, '').slice(0, 30);
        const val = Math.max(0, Math.min(10000, parseInt(v, 10) || 0));
        if (name && val) usage[name] = (usage[name] || 0) + val;
      }
      rec.usage = usage;
      await env.USERS.put(key, JSON.stringify(rec)); // one write per user per day
      return json({ ok: true, plan: rec.plan || 'launch', trialStart: rec.trialStart, trialEnd: addMonths(rec.trialStart, TRIAL_MONTHS) });
    }

    // account deletion from the app's settings
    if (url.pathname === '/delete' && req.method === 'POST') {
      if (!ALLOWED.includes(origin)) return new Response('forbidden', { status: 403, headers: cors });
      const d = await readBody(req);
      const mail = cleanMail(d && d.mail);
      if (!validMail(mail)) return new Response('bad mail', { status: 400, headers: cors });
      await env.USERS.delete('u:' + mail);
      return json({ ok: true });
    }

    const admin = env.ADMIN_KEY && url.searchParams.get('key') === env.ADMIN_KEY;
    const csv = (rows, name) => new Response('\uFEFF' + rows.map(r => r.map(x => '"' + String(x ?? '').replace(/"/g, '""') + '"').join(',')).join('\r\n'),
      { headers: { 'Content-Type': 'text/csv; charset=utf-8', 'Content-Disposition': `attachment; filename="${name}"` } });
    async function allUsers() {
      const out = []; let cursor;
      do {
        const r = await env.USERS.list({ prefix: 'u:', cursor });
        for (const k of r.keys) { const v = await env.USERS.get(k.name, 'json'); if (v) out.push(v); }
        cursor = r.list_complete ? null : r.cursor;
      } while (cursor);
      return out;
    }
    const d10 = iso => iso ? String(iso).slice(0, 10) : '';

    // the whole list, for you only
    if (url.pathname === '/list' && req.method === 'GET') {
      if (!admin) return new Response('forbidden', { status: 403 });
      const users = await allUsers();
      const feats = [...new Set(users.flatMap(u => Object.keys(u.usage || {})))].sort();
      const rows = [['mail', 'name', 'news', 'platform', 'lang', 'first', 'last', 'trial_start', 'trial_end', 'active_days', 'app_version', 'total_actions', ...feats]];
      for (const v of users) {
        const ts = v.trialStart || '';
        const u = v.usage || {};
        const total = Object.values(u).reduce((a, b) => a + b, 0);
        rows.push([v.mail, v.name, v.news ? 'yes' : 'no', v.platform, v.lang, d10(v.first), d10(v.last), d10(ts), ts ? d10(addMonths(ts, TRIAL_MONTHS)) : '', v.days || 0, v.ver || '', total, ...feats.map(f => u[f] || 0)]);
      }
      return csv(rows, 'siman-users.csv');
    }

    // summary per feature, for you only
    if (url.pathname === '/stats' && req.method === 'GET') {
      if (!admin) return new Response('forbidden', { status: 403 });
      const users = await allUsers();
      const tot = {}, who = {};
      for (const v of users) for (const [k, n] of Object.entries(v.usage || {})) { tot[k] = (tot[k] || 0) + n; who[k] = (who[k] || 0) + 1; }
      const rows = [['feature', 'times_used', 'users']];
      Object.keys(tot).sort((a, b) => tot[b] - tot[a]).forEach(k => rows.push([k, tot[k], who[k]]));
      const active7 = users.filter(u => u.last && Date.now() - new Date(u.last) < 7 * 864e5).length;
      rows.push([], ['registered_users', users.length], ['active_last_7_days', active7]);
      return csv(rows, 'siman-stats.csv');
    }

    return new Response('siman-register', { status: 404 });
  }
};
