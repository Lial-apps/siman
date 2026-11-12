// Cloudflare Worker "siman-voice" (6.4: pass + fair daily share, see GUARD below) — התכונות החכמות של התמלול בסימן (רק טקסט, רק כשהמשתמש מבקש)
//
// התמלול עצמו נעשה בתוך המכשיר של המשתמש (siman-asr-worker.js). השמע לא מגיע לכאן ולא ל-Google —
// אין בשרת הזה בכלל דרך לקבל קובץ שמע. מגיע לכאן רק טקסט התמליל, ורק כשהמשתמש לוחץ על אחת מהפעולות
// ומאשר:
//   POST ?op=analyze    {transcript,kind}  → סיכום, נקודות, משימות, תאריכים, שמות הדוברים, הודעת המשך
//   POST ?op=ask        {transcript,question,history}  → {answer}
//   POST ?op=translate  {segments,target}  → {segments}
//   POST ?op=minutes    {transcript,meta}  → פרוטוקול / סיכום מסודר
// שום דבר לא נשמר כאן.
//
// ובנוסף: מראה לקובצי מנוע התמלול (GET /hf/... ו-/cdn/...). בחלק מהטלפונים והרשתות האתר
// huggingface.co חסום, אז האפליקציה מורידה את אותם קבצים דרך כאן. זה רק המנוע (קבצים ציבוריים),
// לא שמע ולא טקסט, ורק של המודלים של סימן (תמלול והפרדת דוברים) — זה לא פרוקסי פתוח.
//
// הקמה (פעם אחת, כ-5 דקות):
// 1. Cloudflare → Workers & Pages → Create → Worker → שם: siman-voice → Deploy → Edit code → מדביקים את הקובץ → Deploy.
// 2. ה-Worker → Settings → Variables and Secrets → Add → Type: Secret → Name: GEMINI_KEY → אותו מפתח Gemini → Deploy.
// הכתובת: https://siman-voice.siman-app.workers.dev — האפליקציה כבר מכוונת אליה.
// מי שלא רוצה שום שליחה: SMART_CLOUD = false בראש siman-voice.js, ואין צורך ב-Worker הזה בכלל.

const ALLOWED = ['https://lial-apps.github.io', 'https://siman.siman-app.workers.dev'];
const G = 'https://generativelanguage.googleapis.com';
const TEXT_MODELS = ['gemini-flash-latest', 'gemini-flash-lite-latest'];
const MAX_TRANSCRIPT = 600000;              // תווים של תמליל לניתוח ולשאלות
const LANGS = { he: 'Hebrew', en: 'English', ar: 'Arabic', ru: 'Russian', fr: 'French', es: 'Spanish', am: 'Amharic', de: 'German' };

const cors = origin => ({
  'Access-Control-Allow-Origin': origin,
  'Access-Control-Allow-Methods': 'GET, HEAD, POST, OPTIONS',
  'Access-Control-Expose-Headers': 'content-length, content-type',
  'Access-Control-Allow-Headers': 'content-type, x-siman-pass, x-user-key',
  'Access-Control-Max-Age': '86400',
  'Vary': 'Origin',
});
const json = (obj, status, origin) => new Response(JSON.stringify(obj), { status, headers: { 'content-type': 'application/json; charset=utf-8', ...cors(origin) } });
const fail = (msg, status, origin) => json({ error: { message: msg } }, status, origin);

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

export default {
  async fetch(req, env, ctx) {
    const origin = req.headers.get('Origin') || '';
    const ok = ALLOWED.includes(origin);
    if (req.method === 'OPTIONS') return new Response(null, { status: ok ? 204 : 403, headers: ok ? cors(origin) : {} });
    if (!ok) return new Response('{"error":{"message":"origin not allowed"}}', { status: 403, headers: { 'content-type': 'application/json' } });
    if (req.method === 'GET' || req.method === 'HEAD') {
      const target = mirrorTarget(new URL(req.url).pathname);
      if (target) return mirror(req, target, origin);
    }
    if (req.method !== 'POST') return fail('POST only', 405, origin);
    if (!env.GEMINI_KEY) return fail('GEMINI_KEY is not set on the worker', 500, origin);
    const op = new URL(req.url).searchParams.get('op') || '';

    let body;
    try { body = await req.json(); } catch (_) { return fail('bad JSON', 400, origin); }
    const jobs = { analyze, ask, translate, minutes };
    if (!jobs[op]) return fail('unknown op', 400, origin);
    // 6.4: a pass, and a fair daily share (or the person's own key). A long transcript translated in parts costs one unit a part.
    const gd = await guard(req, env, op === 'translate' ? Math.max(1, Math.ceil((Array.isArray(body.segments) ? body.segments.length : 0) / 120)) : 1);
    if (!gd.ok) return json(gd.body, gd.status, origin);
    env = gd.env;

    // Cloudflare מנתק בקשה ששותקת 100 שניות. סיכום ארוך יכול לקחת יותר — אז עונים מיד, שולחים רווחים
    // כדי להחזיק את הקו פתוח, ובסוף כותבים את ה-JSON (רווחים לפני JSON מותרים).
    const { readable, writable } = new TransformStream();
    const w = writable.getWriter(), enc = new TextEncoder();
    w.write(enc.encode(' '));
    const ping = setInterval(() => { w.write(enc.encode(' ')).catch(() => {}); }, 15000);
    ctx.waitUntil((async () => {
      let out;
      try { out = await jobs[op](body, env); }
      catch (e) { out = { error: { message: String(e && e.message || e), status: e && e.status || 502 } }; }
      clearInterval(ping);
      try { await w.write(enc.encode(JSON.stringify(out))); await w.close(); } catch (_) {}
    })());
    return new Response(readable, { status: 200, headers: { 'content-type': 'application/octet-stream', 'cache-control': 'no-store, no-transform', ...cors(origin) } });
  },
};

/* ---------------- mirror of the engine files ---------------- */
const MODEL_FILES = /^(itayinbar\/Ozen-v1|Xenova\/whisper-small|onnx-community\/pyannote-segmentation-3\.0|onnx-community\/wespeaker-voxceleb-resnet34-LM)\/resolve\/main\/[\w.\/-]+$/;
const ORT_FILES = /^onnxruntime-web@[\w.-]+\/dist\/ort-wasm[\w.-]*\.(wasm|mjs)$/;
function mirrorTarget(path) {
  let p;
  try { p = decodeURIComponent(path); } catch (_) { return null; }
  if (p.includes('..')) return null;
  if (p.startsWith('/hf/') && MODEL_FILES.test(p.slice(4))) return 'https://huggingface.co/' + p.slice(4);
  if (p.startsWith('/cdn/') && ORT_FILES.test(p.slice(5))) return 'https://cdn.jsdelivr.net/npm/' + p.slice(5);
  return null;
}
async function mirror(req, target, origin) {
  // kept at Cloudflare's edge for a month: after the first person, it does not even go to Hugging Face
  const r = await fetch(target, { method: req.method, redirect: 'follow', cf: { cacheEverything: true, cacheTtl: 2592000 } });
  const h = new Headers(cors(origin));
  h.set('content-type', r.headers.get('content-type') || 'application/octet-stream');
  const len = r.headers.get('content-length'); if (len) h.set('content-length', len);
  h.set('cache-control', r.ok ? 'public, max-age=2592000' : 'no-store');
  return new Response(r.ok ? r.body : null, { status: r.ok ? 200 : r.status, headers: h });
}

function err(message, status) { const e = new Error(message); e.status = status; return e; }

/* ---------------- model calls ---------------- */
async function gen(model, body, env, timeout = 280000) {
  const ctl = new AbortController();
  const t = setTimeout(() => ctl.abort(), timeout);
  try {
    const r = await fetch(`${G}/v1beta/models/${model}:generateContent`, {
      method: 'POST', headers: { 'content-type': 'application/json', 'x-goog-api-key': env.GEMINI_KEY },
      body: JSON.stringify(body), signal: ctl.signal,
    });
    const d = await r.json().catch(() => ({}));
    return { ok: r.ok, status: r.status, d };
  } catch (e) {
    return { ok: false, status: 504, d: { error: { message: ctl.signal.aborted ? 'took too long' : 'Gemini not reachable' } } };
  } finally { clearTimeout(t); }
}
// the text models, with the "thinking" setting each one takes (older: budget, newer: level)
async function genText(body, env, thinking = 'low') {
  const opts = thinking === 'none' ? [{ thinkingBudget: 0 }, { thinkingLevel: 'minimal' }, null] : [{ thinkingLevel: 'low' }, { thinkingBudget: 1024 }, null];
  let last;
  for (const model of TEXT_MODELS) {
    for (const th of opts) {
      const b = th ? { ...body, generationConfig: { ...body.generationConfig, thinkingConfig: th } } : body;
      const r = await gen(model, b, env);
      if (r.ok) { r.model = model; return r; }
      last = r;
      if (r.status === 400 && th) continue;
      break;
    }
    if (last && ![400, 404, 429, 500, 503, 504].includes(last.status)) break;
  }
  throw err((last && last.d && last.d.error && last.d.error.message) || 'Gemini error', last && last.status || 502);
}
const textOf = d => ((d.candidates && d.candidates[0] && d.candidates[0].content && d.candidates[0].content.parts) || []).map(p => p.text || '').join('');
const finish = d => d.candidates && d.candidates[0] && d.candidates[0].finishReason;
function parseJSON(s) {
  s = String(s || '').replace(/^\s*```(?:json)?/i, '').replace(/```\s*$/, '').trim();
  try { return JSON.parse(s); } catch (_) {}
  const a = s.indexOf('{'), b = s.lastIndexOf('}');
  if (a >= 0 && b > a) { try { return JSON.parse(s.slice(a, b + 1)); } catch (_) {} }
  return null;
}

/* ---------------- analyze ---------------- */
const ANALYSIS_SCHEMA = {
  type: 'OBJECT',
  properties: {
    lang: { type: 'STRING' },
    title: { type: 'STRING' },
    kind: { type: 'STRING' },
    summary: { type: 'STRING' },
    points: { type: 'ARRAY', items: { type: 'STRING' } },
    decisions: { type: 'ARRAY', items: { type: 'STRING' } },
    tasks: { type: 'ARRAY', items: { type: 'OBJECT', properties: { what: { type: 'STRING' }, who: { type: 'STRING' }, when: { type: 'STRING' }, date: { type: 'STRING' } }, required: ['what'] } },
    dates: { type: 'ARRAY', items: { type: 'OBJECT', properties: { what: { type: 'STRING' }, date: { type: 'STRING' }, time: { type: 'STRING' } }, required: ['what'] } },
    facts: { type: 'ARRAY', items: { type: 'STRING' } },
    questions: { type: 'ARRAY', items: { type: 'STRING' } },
    speakers: { type: 'ARRAY', items: { type: 'OBJECT', properties: { id: { type: 'STRING' }, name: { type: 'STRING' }, role: { type: 'STRING' } }, required: ['id'] } },
    tone: { type: 'STRING' },
    followup: { type: 'STRING' },
  },
  required: ['title', 'summary'],
};
async function analyze(b, env) {
  const tr = String(b.transcript || '').slice(0, MAX_TRANSCRIPT);
  if (!tr.trim()) throw err('transcript missing', 400);
  const today = new Date().toISOString().slice(0, 10);
  const sys = `You read the transcript of a recorded conversation (a phone call, a voice message, a meeting or a lesson) and pull out what matters.
Lines look like "[03:12] spk_1: text". Answer in the main language of the transcript (Hebrew transcript → Hebrew answer). Return only JSON:
- lang: ISO 639-1 of the transcript. title: a short specific title (max 8 words). kind: one of call / voice / meeting / lesson / interview / dictation.
- summary: 2–5 plain sentences: who talked about what and what came out of it.
- points: the key points (max 8, short). decisions: what was agreed or decided (only if really said).
- tasks: follow-up actions that were promised or asked for: what, who (a name or spk_ label), when (as said, e.g. "עד יום חמישי"), date (YYYY-MM-DD only if it can be worked out; today is ${today}).
- dates: meetings, deadlines and appointments mentioned: what, date (YYYY-MM-DD if possible), time (HH:MM if said).
- facts: numbers, amounts, addresses, phone numbers, names of places or products that someone may need later (exactly as said).
- questions: questions that stayed open.
- speakers: for every spk_ label, the person's name if it is said or clearly implied (e.g. someone is greeted by name), and a role (e.g. "לקוח", "מורה", "נציג שירות"). Leave name empty if unknown — never invent.
- tone: one short phrase about the mood of the conversation.
- followup: a short, polite follow-up message the user could send after this conversation (WhatsApp style, first person, same language).
Use only what is in the transcript. Empty arrays when there is nothing.`;
  const r = await genText({
    systemInstruction: { parts: [{ text: sys }] },
    contents: [{ role: 'user', parts: [{ text: `<transcript kind="${String(b.kind || '').slice(0, 20)}">\n${tr}\n</transcript>` }] }],
    generationConfig: { temperature: 0.2, maxOutputTokens: 8192, responseMimeType: 'application/json', responseSchema: ANALYSIS_SCHEMA },
  }, env);
  const j = parseJSON(textOf(r.d));
  if (!j) throw err('no analysis came back', 502);
  return { analysis: j, model: r.model };
}

/* ---------------- ask ---------------- */
async function ask(b, env) {
  const tr = String(b.transcript || '').slice(0, MAX_TRANSCRIPT);
  const q = String(b.question || '').slice(0, 2000);
  if (!tr.trim() || !q.trim()) throw err('transcript or question missing', 400);
  const contents = [{ role: 'user', parts: [{ text: `<transcript>\n${tr}\n</transcript>` }] }, { role: 'model', parts: [{ text: 'קראתי את התמליל. מה השאלה?' }] }];
  for (const h of (Array.isArray(b.history) ? b.history.slice(-6) : [])) {
    if (h && h.q && h.a) contents.push({ role: 'user', parts: [{ text: String(h.q).slice(0, 2000) }] }, { role: 'model', parts: [{ text: String(h.a).slice(0, 6000) }] });
  }
  contents.push({ role: 'user', parts: [{ text: q }] });
  const r = await genText({
    systemInstruction: { parts: [{ text: `You are the assistant inside סימן. The user recorded a conversation; its transcript is given with times like [03:12]. Answer only from the transcript; if it is not there, say so in one sentence. After each fact add the time it was said, like [03:12]. Answer in the language of the question. Short plain paragraphs or lines starting with "• ". No markdown headings, bold or tables. If asked to write something (an email, a message, a summary for someone), write it ready to send.` }] },
    contents,
    generationConfig: { temperature: 0.3, maxOutputTokens: 4096 },
  }, env);
  return { answer: textOf(r.d).trim() };
}

/* ---------------- translate ---------------- */
async function translate(b, env) {
  const t = String(b.target || '');
  if (!LANGS[t]) throw err('unknown target language', 400);
  const segs = (Array.isArray(b.segments) ? b.segments : []).slice(0, 3000).map(x => String(x || '').slice(0, 4000));
  if (!segs.length) throw err('nothing to translate', 400);
  const out = [];
  for (let i = 0; i < segs.length; i += 120) {
    const part = segs.slice(i, i + 120);
    const r = await genText({
      systemInstruction: { parts: [{ text: `Translate each line of a conversation transcript into ${LANGS[t]}. Input: JSON array of strings. Output: only a JSON array of the same length, one translation per line, same order. Natural spoken style, keep names and numbers. A line already in ${LANGS[t]} stays as is.${t === 'ar' ? ' Modern Standard Arabic.' : ''}` }] },
      contents: [{ role: 'user', parts: [{ text: JSON.stringify(part) }] }],
      generationConfig: { temperature: 0.2, maxOutputTokens: 32768, responseMimeType: 'application/json', responseSchema: { type: 'ARRAY', items: { type: 'STRING' } } },
    }, env, 'none');
    let a; try { a = JSON.parse(textOf(r.d)); } catch (_) { a = null; }
    if (!Array.isArray(a)) throw err('translation failed', 502);
    for (let k = 0; k < part.length; k++) out.push(String(a[k] == null ? part[k] : a[k]));
  }
  return { segments: out };
}

/* ---------------- minutes (פרוטוקול) ---------------- */
async function minutes(b, env) {
  const tr = String(b.transcript || '').slice(0, MAX_TRANSCRIPT);
  if (!tr.trim()) throw err('transcript missing', 400);
  const meta = b.meta || {};
  const r = await genText({
    systemInstruction: { parts: [{ text: `Write formal meeting minutes (פרוטוקול) from a meeting transcript, in the transcript's main language (Hebrew → formal Hebrew). Return only JSON:
{"title":"...","participants":["name or role"],"topics":[{"title":"...","discussion":"2–4 sentences, neutral, who said what in reported speech","decisions":["..."]}],"tasks":[{"what":"...","who":"...","when":"..."}],"next":"next meeting / follow-up if said, else empty"}
Use the speaker names given; for unknown speakers use their label. Only what was actually said.` }] },
    contents: [{ role: 'user', parts: [{ text: `Speakers: ${JSON.stringify(meta.speakers || {})}\nDate: ${String(meta.date || '')}\n<transcript>\n${tr}\n</transcript>` }] }],
    generationConfig: { temperature: 0.2, maxOutputTokens: 12288, responseMimeType: 'application/json' },
  }, env);
  const j = parseJSON(textOf(r.d));
  if (!j) throw err('no minutes came back', 502);
  return { minutes: j };
}
