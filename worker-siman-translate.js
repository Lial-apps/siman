// Cloudflare Worker "siman-translate" (6.4: pass + fair daily share, see GUARD below)
// סימן — שרת תרגום (Cloudflare Worker)
// מחזיק את מפתח Gemini בסוד, כך שמשתמשי האפליקציה לא צריכים מפתח משלהם.
// הגדרות ב-Cloudflare (Settings → Variables and Secrets):
//   GEMINI_KEY  (Secret)  — מפתח Gemini מ-aistudio.google.com/apikey
//   MODEL       (Text, לא חובה) — ברירת מחדל: gemini-flash-lite-latest
// מומלץ (Settings → Bindings → Add → Rate limiter):
//   Variable name: TR_LIMIT · Namespace ID: 1001 · Limit: 8 · Period: 60
//   מגבלת הבקשות נשמרת אז ברשת של Cloudflare ולא רק בזיכרון של מופע אחד, ולכן אי אפשר לעקוף אותה.
//   בלי ה-binding השרת ממשיך לעבוד עם המגבלה הפנימית כמו קודם.
// כתובות מורשות: GitHub, קישור ההתקנה (siman.siman-app.workers.dev), ובדיקה מקומית.

const ALLOWED_ORIGINS = ['https://lial-apps.github.io', 'https://siman.siman-app.workers.dev', 'http://localhost:8765'];
const MAX_CHARS = 12000;          // עמוד ארוך במיוחד
const MAX_IMAGE = 3_000_000;      // תמונת עמוד (base64)
const PER_IP_PER_MINUTE = 8;      // הגנה בסיסית מניצול

const hits = new Map(); // מונה בקשות לכל כתובת IP (לכל מופע של השרת)

function cors(origin) {
  const ok = ALLOWED_ORIGINS.includes(origin);
  return {
    'access-control-allow-origin': ok ? origin : ALLOWED_ORIGINS[0],
    'access-control-allow-methods': 'POST, OPTIONS',
    'access-control-allow-headers': 'content-type, x-siman-pass, x-user-key',
    'access-control-max-age': '86400',
    'vary': 'origin'
  };
}
const json = (obj, status, headers) =>
  new Response(JSON.stringify(obj), { status, headers: { ...headers, 'content-type': 'application/json; charset=utf-8' } });

// ===================================================================================================
// 6.4 GUARD — who may use the Gemini key, and how much. The same block is in siman-ai, siman-voice and
// siman-translate. The "Origin" check alone can be faked by any script, so:
//  1. a pass: the app proves once that a person is using it (Cloudflare Turnstile, free and invisible), and
//     gets a pass signed with PASS_SECRET for a few hours. Every request to Gemini must carry a valid pass.
//  2. a fair daily share: requests are counted per device and per internet address (D1, free), so one heavy
//     user (or someone abusing the service) cannot use up the free Gemini quota for everybody.
//  3. a personal key: a person who added their own free Gemini key in the app's settings uses it here
//     (sent in "x-user-key", used for this request only, never stored or logged) and has no daily share.
// Setup (Settings → Variables and Secrets / Bindings):
//   PASS_SECRET       Secret — the same long random text in all three workers
//   TURNSTILE_SECRET  Secret — siman-ai only (it gives out the passes)
//   DB                D1 binding → the siman-stats database (the counters live in their own table)
//   LIMIT_DEVICE / LIMIT_IP / LIMIT_ALL   optional plain-text variables: units per day (defaults below; 0 = no limit)
// While PASS_SECRET is not set, nothing changes (the worker answers as before), so the code can be pasted first.
// ===================================================================================================
const GUARD_HOSTS = ['lial-apps.github.io', 'siman.siman-app.workers.dev'];   // where the Turnstile check may come from
const PASS_TTL = 6 * 3600;                                                     // a pass lasts 6 hours
const LIMITS = { device: 40, ip: 100, all: 0, pass: 25 };                      // per day (UTC); pass = new passes per address
const USER_KEY_RE = /^(AIza[\w-]{30,60}|AQ\.[\w.-]{20,200})$/;
const DAY_MSG = 'הגעת למכסה היומית של השירותים החכמים בסימן (הסיכום, התרגום והקריאה החכמה). הכלים שבמכשיר ממשיכים לעבוד כרגיל. אפשר להמשיך מחר, או להוסיף בהגדרות מפתח Gemini אישי בחינם ולהמשיך בלי הגבלה.';
const ALL_MSG = 'השירות החכם של סימן הגיע היום למכסה הכוללת שלו. אפשר לנסות שוב מחר, או להוסיף בהגדרות מפתח Gemini אישי בחינם ולהמשיך מיד.';
const PASS_MSG = 'לא הצלחנו לאמת את המכשיר. רעננו את סימן ונסו שוב (אם יש חוסם פרסומות, ייתכן שהוא חוסם את challenges.cloudflare.com).';
const b64u = buf => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
const sameText = (a, b) => { a = String(a || ''); b = String(b || ''); if (!a || a.length !== b.length) return false; let d = 0; for (let i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i); return d === 0; };
const hmacKeys = new Map();
async function sign(secret, text) {
  let k = hmacKeys.get(secret);
  if (!k) { k = await crypto.subtle.importKey('raw', new TextEncoder().encode(secret), { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']); hmacKeys.set(secret, k); }
  return b64u(await crypto.subtle.sign('HMAC', k, new TextEncoder().encode(text)));
}
// pass = <device>.<expires, unix seconds>.<signature>
async function makePass(env, dev) { const p = `${dev}.${Math.floor(Date.now() / 1000) + PASS_TTL}`; return { pass: `${p}.${await sign(env.PASS_SECRET, p)}`, ttl: PASS_TTL }; }
async function readPass(env, pass) {
  const m = /^([a-z0-9]{16,40})\.(\d{9,11})\.([\w-]{40,50})$/.exec(String(pass || ''));
  if (!m || +m[2] < Date.now() / 1000) return null;
  return sameText(m[3], await sign(env.PASS_SECRET, `${m[1]}.${m[2]}`)) ? m[1] : null;
}
const limitOf = (env, name) => { const v = env['LIMIT_' + name.toUpperCase()]; return v != null && v !== '' && Number.isFinite(+v) ? +v : LIMITS[name]; };
let usageReady = null;
// adds w to each counter of today and returns the new totals; null = no database (no limits)
async function addUsage(env, keys, w) {
  if (!env.DB) return null;
  if (!usageReady) usageReady = env.DB.prepare('CREATE TABLE IF NOT EXISTS ai_usage (day TEXT NOT NULL, k TEXT NOT NULL, n INTEGER NOT NULL DEFAULT 0, PRIMARY KEY (day, k))').run()
    .then(() => env.DB.prepare('DELETE FROM ai_usage WHERE day < ?').bind(new Date(Date.now() - 3 * 864e5).toISOString().slice(0, 10)).run())
    .catch(e => { usageReady = null; throw e; });
  await usageReady;
  const day = new Date().toISOString().slice(0, 10);
  const st = env.DB.prepare('INSERT INTO ai_usage (day, k, n) VALUES (?, ?, ?) ON CONFLICT (day, k) DO UPDATE SET n = n + excluded.n RETURNING n');
  const out = await env.DB.batch(keys.map(k => st.bind(day, k, w)));
  return out.map(r => (r.results && r.results[0] && r.results[0].n) || 0);
}
const clientIp = req => req.headers.get('CF-Connecting-IP') || 'unknown';
const guardFail = (status, code, message) => ({ ok: false, status, body: { error: { message, code, status, ...(code === 'DAY' || code === 'ALL' ? { quota: 'per day' } : {}) } } });
// → { ok: true, env } (env.GEMINI_KEY may be the person's own key) or { ok: false, status, body }
async function guard(req, env, weight = 1) {
  if (!env.PASS_SECRET) return { ok: true, env };                    // not set up yet: as before
  const dev = await readPass(env, req.headers.get('x-siman-pass'));
  if (!dev) return guardFail(401, 'PASS', PASS_MSG);
  const own = String(req.headers.get('x-user-key') || '').trim();
  if (USER_KEY_RE.test(own)) return { ok: true, env: { ...env, GEMINI_KEY: own }, own: true };
  let n;
  try { n = await addUsage(env, ['d:' + dev, 'i:' + clientIp(req), 'all'], weight); } catch (_) { n = null; }   // a database hiccup does not stop the app
  if (n) {
    const [d, i, all] = n, ld = limitOf(env, 'device'), li = limitOf(env, 'ip'), la = limitOf(env, 'all');
    if (la && all > la) return guardFail(429, 'ALL', ALL_MSG);
    if ((ld && d > ld) || (li && i > li)) return guardFail(429, 'DAY', DAY_MSG);
  }
  return { ok: true, env };
}
// POST /pass {t: Turnstile token, dev} → {pass, ttl}   (siman-ai only)
async function givePass(req, env) {
  if (!env.PASS_SECRET || !env.TURNSTILE_SECRET) return { status: 503, body: { error: { message: 'pass not set up', code: 'NOPASS' } } };
  let b; try { b = await req.json(); } catch (_) { b = null; }
  const dev = String(b && b.dev || ''), t = String(b && b.t || '');
  if (!/^[a-z0-9]{16,40}$/.test(dev) || !t || t.length > 2048) return { status: 400, body: { error: { message: 'bad request', code: 'PASS' } } };
  try { const n = await addUsage(env, ['p:' + clientIp(req)], 1); if (n && limitOf(env, 'pass') && n[0] > limitOf(env, 'pass')) return guardFail(429, 'PASSDAY', PASS_MSG); } catch (_) {}
  const form = new FormData(); form.append('secret', env.TURNSTILE_SECRET); form.append('response', t); form.append('remoteip', clientIp(req));
  let v = {};
  try { v = await (await fetch('https://challenges.cloudflare.com/turnstile/v0/siteverify', { method: 'POST', body: form })).json(); } catch (_) {}
  if (!v.success || !GUARD_HOSTS.includes(v.hostname)) return { status: 401, body: { error: { message: PASS_MSG, code: 'PASS', why: v['error-codes'] || [] } } };
  return { status: 200, body: await makePass(env, dev) };
}
// how many units a request costs (a careful reading of a page asks Gemini two or three times)
const WEIGHT = { scanAccurate: 3, lines: 2, table: 2, tableCheck: 1, linesCheck: 1 };
// ======================================== end of the guard ========================================

function tooManyLocal(ip) {
  const now = Date.now();
  const list = (hits.get(ip) || []).filter(t => now - t < 60000);
  list.push(now);
  hits.set(ip, list);
  if (hits.size > 5000) hits.clear();
  return list.length > PER_IP_PER_MINUTE;
}
// the shared limit (Rate limiter binding) when it is set up, otherwise the local one
async function tooMany(env, ip) {
  if (env.TR_LIMIT && typeof env.TR_LIMIT.limit === 'function') {
    try { const { success } = await env.TR_LIMIT.limit({ key: ip }); return !success; }
    catch (e) { console.log('rate limiter', String(e)); }
  }
  return tooManyLocal(ip);
}

function systemPrompt(target) {
  const t = target === 'en' ? 'English' : 'Hebrew';
  return `You are a professional document translator. Translate the page the user sends into ${t}.\n` +
    `Output format, exactly:\nLANG: <ISO 639-1 code of the source text's main language>\n---\n<the translation>\n` +
    `Keep paragraph breaks, headings, list numbering, numbers, names and line structure. ` +
    `Output only the translation, with no notes or explanations. If the text is already in ${t}, return it unchanged.`;
}

// "smart conversion": the page picture + its exact text -> a structured document (JSON)
const LAYOUT_PROMPT = `You rebuild one document page as an editable Word document.
You receive the page image and its text lines, numbered: L<n> [top-bottom|left-right] size (positions 0-1000 from the top-left) text.
Return ONLY compact JSON {"blocks":[...]} in reading order (Hebrew pages read right-to-left, top to bottom). Refer to text ONLY by line numbers; never write the text itself.
Block types:
{"type":"heading","level":1|2|3,"align":"right|center|left","lines":[n,...]}
{"type":"paragraph","align":"right|center|left|justify","lines":[n,...]}
{"type":"list","ordered":true|false,"items":[{"level":0|1|2,"lines":[n,...]}]}
{"type":"table","borders":true|false,"rows":[[{"lines":[n,...],"fill":"#RRGGBB"|null,"align":"right|center|left"}]]}
{"type":"row","columns":[{"width":0-100,"blocks":[BLOCK]}]}  (side-by-side areas without borders, e.g. a letterhead: logo | title | details)
{"type":"image","box":[left,top,right,bottom],"align":"right|center|left"}  (logos, emblems, photos, stamps, signatures; tight box, 0-1000)
Rules: use every line number exactly once. Lines of one paragraph or list item go together in one block. Numbered/bulleted items become list items.
A grid of cells becomes a table. Skip page frames and decorative rules.`;
async function layout(body, env, h) {
  const img = typeof body.image === 'string' ? body.image : '';
  if (!img || img.length > MAX_IMAGE) return json({ error: { message: 'תמונת העמוד חסרה או גדולה מדי' } }, 413, h);
  const text = typeof body.text === 'string' ? body.text.slice(0, 20000) : '';
  // if a model is busy or unavailable, try the next one
  const models = [...new Set([env.LAYOUT_MODEL || 'gemini-flash-latest', 'gemini-flash-lite-latest', 'gemini-2.5-flash'])];
  let res, last = '';
  outer: for (const model of models) {
    for (const think of [false, true]) {          // first without "thinking" (much faster)
      res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`, {
        method: 'POST',
        headers: { 'content-type': 'application/json', 'x-goog-api-key': env.GEMINI_KEY },
        body: JSON.stringify({
          systemInstruction: { parts: [{ text: LAYOUT_PROMPT }] },
          contents: [{ role: 'user', parts: [{ inline_data: { mime_type: 'image/jpeg', data: img } }, { text: 'Text lines of this page:\n' + text }] }],
          generationConfig: { temperature: 0.1, maxOutputTokens: 8192, responseMimeType: 'application/json', ...(think ? {} : { thinkingConfig: { thinkingBudget: 0 } }) }
        })
      });
      if (res.ok) break outer;
      last = await res.text();
      console.log('layout', model, think ? 'think' : 'fast', res.status, last.slice(0, 300)); // visible in Cloudflare → Observability → Logs
      if (/per ?day/i.test(last)) break outer;                       // daily quota: stop
      if (res.status === 400 && !think) continue;                    // maybe this model rejects the "no thinking" setting: retry it plainly
      if ([401, 403].includes(res.status)) break outer;              // key problem: other models won't help
      continue outer;                                                // busy / not available / bad request: next model
    }
  }
  if (!res.ok) return new Response(last, { status: res.status, headers: { ...h, 'content-type': 'application/json; charset=utf-8' } });
  return new Response(await res.text(), { status: res.status, headers: { ...h, 'content-type': 'application/json; charset=utf-8' } });
}

export default {
  async fetch(request, env) {
    const origin = request.headers.get('origin') || '';
    const h = cors(origin);
    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: h });
    if (request.method !== 'POST') return json({ error: { message: 'שרת התרגום של סימן פועל.' } }, 200, h);
    if (!ALLOWED_ORIGINS.includes(origin)) return json({ error: { message: 'Forbidden origin' } }, 403, h);
    if (!env.GEMINI_KEY) return json({ error: { message: 'השרת לא הוגדר: חסר GEMINI_KEY' } }, 500, h);

    const ip = request.headers.get('cf-connecting-ip') || 'unknown';
    if (await tooMany(env, ip)) return json({ error: { message: 'Too many requests', details: [{ retryDelay: '20s' }] } }, 429, h);

    let body;
    try { body = await request.json(); } catch { return json({ error: { message: 'Bad JSON' } }, 400, h); }
    // 6.4: a pass, and a fair daily share (or the person's own key)
    const gd = await guard(request, env, 1);
    if (!gd.ok) return json(gd.body, gd.status, h);
    env = gd.env;
    if (body.mode === 'layout') return layout(body, env, h);
    const text = typeof body.text === 'string' ? body.text : '';
    const target = body.target === 'en' ? 'en' : 'he';
    if (!text.trim()) return json({ error: { message: 'Empty text' } }, 400, h);
    if (text.length > MAX_CHARS) return json({ error: { message: 'העמוד ארוך מדי לתרגום' } }, 413, h);

    const model = env.MODEL || 'gemini-flash-lite-latest';
    const res = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(model)}:generateContent`, {
      method: 'POST',
      headers: { 'content-type': 'application/json', 'x-goog-api-key': env.GEMINI_KEY },
      body: JSON.stringify({
        systemInstruction: { parts: [{ text: systemPrompt(target) }] },
        contents: [{ role: 'user', parts: [{ text: `<page>\n${text}\n</page>` }] }],
        generationConfig: { temperature: 0.2, maxOutputTokens: 16384 }
      })
    });
    // Gemini's answer (or its error, with the same status) goes back unchanged
    return new Response(await res.text(), { status: res.status, headers: { ...h, 'content-type': 'application/json; charset=utf-8' } });
  }
};
