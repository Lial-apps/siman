// Cloudflare Worker "siman-ai" — the smart features of סימן
//   translate : translation to Russian, French, Spanish and Arabic (Hebrew/English stay on siman-translate)
//   scan      : a photo / scanned page (handwritten or printed) → clean text with headings, lists and tables
//   ask       : questions about the open document, summaries, key points
//
// Setup (one time, ~5 minutes):
// 1. Cloudflare → Workers & Pages → Create → Worker → name: siman-ai → Deploy → Edit code → paste this file → Deploy.
// 2. The worker → Settings → Variables and Secrets → Add → type Secret → name GEMINI_KEY → the same Gemini key
//    that siman-translate uses → Deploy.
// The address becomes https://siman-ai.siman-app.workers.dev — the app already points there.
//
// The key never reaches the app. Only requests coming from the app's own address are answered.

// the app is served from both addresses: GitHub Pages and the Cloudflare one (siman-worker.js)
const ALLOWED = ['https://lial-apps.github.io', 'https://siman.siman-app.workers.dev'];
const MODELS = ['gemini-flash-latest', 'gemini-flash-lite-latest'];
const MAX_TEXT = 400000;            // characters of document text for "ask"
const MAX_IMAGE = 7 * 1024 * 1024;  // base64 characters of one page picture
// Cloudflare cuts a request that sends nothing back for 100 seconds (error 524). Reading a page of
// handwriting can take longer, so "scan" and "ask" answer at once and keep the line open with spaces
// until Gemini is done. The app reads the JSON at the end (JSON allows spaces before it).
const STREAM_MODES = ['scan', 'scanAccurate', 'ask', 'lines', 'linesFast', 'linesCheck', 'table', 'tableCheck'];
// the whole request has one time budget (the app gives up after 150 s): each model gets what is left
const BUDGET = { tableCheck: 100000, orient: 30000, table: 120000, linesCheck: 80000, lines: 110000, linesFast: 60000, scan: 70000, scanAccurate: 125000, ask: 85000, translate: 85000, corners: 30000 };
const MODEL_TIMEOUT = { tableCheck: 80000, orient: 20000, table: 90000, linesCheck: 70000, lines: 75000, linesFast: 45000, scan: 60000, scanAccurate: 90000, scanCheck: 45000, ask: 80000, translate: 80000, corners: 30000 }; // ms per model before trying the next one
// reading a page: "accurate" = the main model, thinking a little (best for handwriting, ~20–90 s);
// "fast" = the light model, no thinking (a few seconds, guesses more)
// lines: on the ivrit.ai Hebrew handwriting benchmark the main model reads single lines best of all models
// (median 12% of letters wrong), far ahead of the light one — and far ahead of itself on a whole page
const MODELS_FOR = { tableCheck: ['gemini-flash-latest'], orient: ['gemini-flash-lite-latest', 'gemini-flash-latest'], table: ['gemini-flash-latest', 'gemini-flash-lite-latest'], linesCheck: ['gemini-flash-latest'], lines: ['gemini-flash-latest', 'gemini-flash-lite-latest'], linesFast: ['gemini-flash-lite-latest', 'gemini-flash-latest'], scan: ['gemini-flash-lite-latest', 'gemini-flash-latest'], scanAccurate: ['gemini-flash-latest', 'gemini-flash-lite-latest'], scanCheck: ['gemini-flash-latest'], corners: ['gemini-flash-latest', 'gemini-flash-lite-latest'] };
// no "thinking" for reading and translating = a much faster answer; a little for questions
// Models differ in which setting they take (older: thinkingBudget, newer: thinkingLevel), and one that does
// not take a setting answers "invalid argument" — so each is tried in turn, and at the end none at all.
const THINKING = {
  scan: [{ thinkingBudget: 0 }, { thinkingLevel: 'minimal' }, { thinkingLevel: 'low' }, null],
  // handwriting is read better with more thinking: each step down is tried only if the model refuses the one above
  // a little thinking: much faster than "medium", and the clean picture from the app does most of the work
  // lines: the model's own default thinking, as on the benchmark; then less if the request is refused
  lines: [null],
  orient: [{ thinkingBudget: 0 }, { thinkingLevel: 'minimal' }, null],
  tableCheck: [{ thinkingLevel: 'low' }, { thinkingBudget: 2048 }, null],
  table: [{ thinkingLevel: 'low' }, { thinkingBudget: 2048 }, null],
  linesCheck: [{ thinkingLevel: 'low' }, { thinkingBudget: 2048 }, null],
  linesFast: [{ thinkingBudget: 0 }, { thinkingLevel: 'minimal' }, null],
  scanAccurate: [{ thinkingLevel: 'low' }, { thinkingBudget: 1024 }, null],
  scanCheck: [{ thinkingLevel: 'low' }, { thinkingBudget: 1024 }, null],
  ask: [{ thinkingBudget: 512 }, { thinkingLevel: 'low' }, null],
  translate: [{ thinkingBudget: 0 }, { thinkingLevel: 'minimal' }, null],
  corners: [{ thinkingLevel: 'minimal' }, { thinkingBudget: 0 }, null],
};
const LANGS = { he: 'Hebrew', en: 'English', ru: 'Russian', fr: 'French', es: 'Spanish', ar: 'Arabic' };

const PROMPTS = {
  translate: t => `You are a professional document translator. Translate the page the user sends into ${LANGS[t]}.
Output format, exactly:
LANG: <ISO 639-1 code of the source text's main language>
---
<the translation>
Keep paragraph breaks, headings, list numbering, numbers, names and line structure. When lines start with [number] markers, keep every marker and the same count of lines. Output only the translation, with no notes or explanations. If the text is already in ${LANGS[t]}, return it unchanged.${t === 'ru' ? ' Use natural modern Russian with correct cases and the formal "Вы" in letters and forms.' : ''}${t === 'ar' ? ' Use Modern Standard Arabic.' : ''}`,

  scan: hint => `You transcribe one photographed or scanned page — handwritten, printed or mixed — into clean digital text.
Return only JSON, no markdown fences:
{"lang":"<ISO 639-1 of the main language>","blocks":[ ... ]}
Each block is one of:
{"t":"h1","text":"..."}   a title
{"t":"h2","text":"..."}   a sub-heading
{"t":"p","text":"..."}    a paragraph: join the lines of one paragraph into one string
{"t":"li","text":"..."}   a bullet item (without the bullet sign)
{"t":"oli","text":"..."}  a numbered item (without its number)
{"t":"table","rows":[["cell","cell"],["cell","cell"]]}  a table, first row first, cells in reading order
Rules:
- Handwritten Hebrew is usually in Hebrew cursive script (כתב רהוט), whose letter shapes are very different from print (for example cursive א looks like a small κ, ל has a tall loop, ס and ם are round, ק and ף descend below the line). Read each letter by its cursive shape.\n- Notes are often lists of points: a circled letter or a circled star or x at the start of a line (Ⓐ ⊗ (*) (א)) is a list marker — make the line an "li" block. Lines connected by arrows going down are a sequence: write them in order joined with " → " in one block. A curly brace grouping lines, or a short word written in the margin beside a group, is kept as its own short line after that group.\n- The page may be a notebook page: ignore the printed grid, squares, ruled lines, margins, spiral holes and anything showing through from the other side.\n- Read slowly, letter by letter. Write the writer's words exactly as written. Do not correct spelling, do not complete, summarize or translate.
- Read every line of the page, top to bottom, including small, faint or crowded writing, notes in the margins and the name or date at the top. Do not skip a line because it is hard to read.
- When the letters of a word are unclear, decide between the possible readings by the letters you can see together with the sentence, the subject and the language (for example a Bible text: Бог, Авраам, Сара, אברהם, שרה). Never invent a word that the letters do not support.
- A word you are not sure of: your best reading inside [[ ]], for example [[שלום]]. Use [[?]] only when not even a guess is possible. A best guess in [[ ]] is always better than [[?]].
- Arrows drawn between words (→ ← ↓) are kept as → or ← or ↓ in the text, in the place they appear. Abbreviations, codes and English or Latin letters inside Hebrew text are kept exactly as written.
- Skip crossed-out words, page decorations and doodles. Keep numbers, dates, math and signs as written.
- Right-to-left languages (Hebrew, Arabic) in normal logical order.
- If the page is empty, return {"lang":null,"blocks":[]}.${hint ? `\nThe page is probably in: ${hint}.` : ''}`,

  ask: () => `You are the assistant inside סימן, a document app. The user opened a document; its text is given page by page, each page starting with a line like "===== עמוד 3 =====".
Answer the user's question only from the document. If the answer is not in it, say so in one short sentence.
After every fact, cite the page it came from: [עמ' 3] when you answer in Hebrew, [p. 3] in any other language.
Answer in the language the question is written in. Be clear and short: plain short paragraphs, or lines that start with "• ". No markdown headings, no bold, no tables.`,
};

// finding the sheet of paper in a photo: its four corners, so the app can cut it out and straighten it
const LINES_PROMPT = hint => `You read handwritten notes for a document app. You get several lines, numbered, in reading order (the page's first line first). Each line is cut from a photographed page and usually comes in two views of the same picture: view A is high-contrast (ink black, paper white), view B is softer and keeps the thin strokes. Use both views to decide each letter.
For every line, transcribe the writing in it exactly, letter by letter.
- Modern Hebrew handwriting is almost always Hebrew cursive script (כתב רהוט), whose letter shapes are very different from print. Read each letter by its cursive shape. Final forms ך ן ף ץ and ק go below the line, ל rises above it.
- Easily confused: ו/י/ן (stroke length), ה/ח/ת, ב/כ/נ, ס/ם, ד/ר, ג/נ, ע/צ. Look at each one again.
- Write only what the letters show. Do not correct spelling, do not complete, rephrase, translate or summarize. Do not invent words. Keep numbers, dates, slashes, English letters and abbreviations (גרשיים ״ and גרש ׳) exactly as written.
- A word you are not sure of: your best reading inside [[ ]]. Use [[?]] only when not even a guess is possible.
- Edges: letters cut by the top or bottom edge of the image belong to the neighbouring line — ignore them. If the image holds two short separate pieces of writing (for example a word written above the line), write both, the upper one first, separated by " / ".
- Skip crossed-out words, the printed grid or ruled lines of the notebook, and doodles.
- kind: "list" when the line starts with a list marker — a circled star, x, letter or number, a bullet or a dash — and then write the text WITHOUT the marker; "title" when it is clearly a heading (underlined, centered, or larger); "arrow" when the image is only an arrow; "empty" when there is no writing; otherwise "text".
Return exactly one entry per line, in the same order and with the same number n.${hint ? `\nThe notes are probably in: ${hint}.` : ''}`;
const LINES_SCHEMA = {
  type: 'OBJECT',
  properties: {
    lang: { type: 'STRING', nullable: true },
    lines: { type: 'ARRAY', items: { type: 'OBJECT', properties: { n: { type: 'INTEGER' }, kind: { type: 'STRING', enum: ['text', 'list', 'title', 'arrow', 'empty'] }, text: { type: 'STRING' } }, required: ['n', 'kind', 'text'] } },
  },
  required: ['lines'],
};

const LINES_CHECK_PROMPT = `You proofread a first reading of handwritten notes (usually Hebrew cursive, כתב רהוט) against the pictures of the lines.
You get every line of the page as a picture, numbered, and the first reading as JSON with one entry per line.
Go line by line, word by word, and look at the letters in the picture again:
- A word whose letters do not match the reading: fix it. Hebrew cursive letters are easily confused (ו/י/ן, ה/ח/ת, ב/כ/נ, ס/ם, ד/ר, ג/נ, ע/צ, ק/ף).
- A word that makes no sense in its sentence, or in the subject of the page, is probably misread. Find the reading that the letters support and that fits the rest of the page (the page has one subject; names and terms repeat). If the writer's own words are listed, prefer them when the letters fit.
- Never invent text the letters do not support, never complete, rephrase, correct spelling or translate.
- Mark with [[ ]] every word you are not fairly sure of, including words you changed and are still unsure of — the person checks every marked word, so marking an uncertain word is always better than guessing silently. Use [[?]] only when not even a guess is possible.
- Keep the same number of entries, the same n and the kind of each line (you may fix a wrong kind).
Return the full corrected reading in the same JSON format, and nothing else.`;
// the writer's own words, learned from their corrections in the app (names, places, work terms)
const vocabText = v => {
  const w = (Array.isArray(v) ? v : []).map(x => String(x).trim()).filter(x => x && x.length <= 40 && !/[<>{}]/.test(x)).slice(0, 200);
  return w.length ? `\nWords this writer often uses (use one only when the letters in the picture fit it): ${w.join(', ')}` : '';
};

const TABLE_PROMPT = hint => `You turn a photographed table (printed or handwritten, often Hebrew) into spreadsheet data.
You get the page picture, and sometimes sharper strips of the same page from top to bottom (overlapping).
- Find every table on the page. For each one return its rows from top to bottom; each row is the list of its cells in reading order of the table's language (in a Hebrew right-to-left table the rightmost column is the FIRST cell).
- Keep the header row as the first row. Every row has the same number of cells: an empty cell is "". A merged cell: its text in the first cell it covers, "" in the rest.
- Copy every cell exactly as written: numbers, decimal points, minus signs, %, currency signs, dates and units as written. Do not calculate, total, round, translate or correct anything. Do not invent rows.
- A cell you cannot read with confidence: your best reading inside [[ ]].
- title: the table's caption or the heading written right above it, if any.
- notes: other short lines written on the page outside the tables (a date line, a name, a signature line), each as written, in reading order. A blank to fill (____) stays as written.
- Use only header text that is actually written on the page. Never invent column names or rows. Text that wraps onto a second line inside one cell is ONE cell (join the parts in order, without adding characters).
- Keep the table's real structure: the number of columns and rows the ruled lines show. A column that is empty on the page stays as an empty column.
- If the page has no table but has lines of items with values (a list with amounts, a form), make a table of it with sensible columns.${hint ? `\nThe page is probably in: ${hint}.` : ''}`;
const TABLE_SCHEMA = {
  type: 'OBJECT',
  properties: {
    lang: { type: 'STRING', nullable: true },
    notes: { type: 'ARRAY', items: { type: 'STRING' } },
    tables: { type: 'ARRAY', items: { type: 'OBJECT', properties: { title: { type: 'STRING', nullable: true }, rows: { type: 'ARRAY', items: { type: 'ARRAY', items: { type: 'STRING' } } } }, required: ['rows'] } },
  },
  required: ['tables'],
};

const TABLE_CHECK_PROMPT = `You proofread spreadsheet data that was read from a photographed table, against the picture.
You get the page picture (and sometimes sharper strips of it) and the first reading as JSON.
Check cell by cell: the number of rows and columns must match the ruled lines on the page; every cell's text must match what is written in that cell, character by character (digits, commas, decimal points, signs). Fix what does not match. A cell that is empty on the page is "". Header text only as written — never invented.
In a Hebrew (right-to-left) table the rightmost column is the FIRST cell of every row.
Mark with [[ ]] every cell you are not fairly sure of — the person checks every marked cell.
Return the full corrected data in the same JSON format, and nothing else.`;

const CORNERS_PROMPT = `You find the document in a photo taken to scan it: the sheet of paper, page, notebook page, form or receipt the person is photographing.
Ignore everything around it: the table, hands, other papers, bags, plastic sleeves, shadows.
When several papers are visible, choose the one in the middle that takes the most space and has the writing.
Return its 4 corners as points [y, x], each normalized to 0-1000 over the whole image (y down from the top, x right from the left), in this order: top-left, top-right, bottom-right, bottom-left of the sheet as it appears in the photo.
Follow the real edges of the paper, also when it is tilted or photographed at an angle. A corner outside the photo goes on the photo's edge.
If no document is visible, return {"found":false,"corners":[]}.`;
const CORNERS_SCHEMA = {
  type: 'OBJECT',
  properties: {
    found: { type: 'BOOLEAN' },
    corners: { type: 'ARRAY', items: { type: 'ARRAY', items: { type: 'INTEGER' } } },
  },
  required: ['found', 'corners'],
};
const SCAN_SCHEMA = {
  type: 'OBJECT',
  properties: {
    lang: { type: 'STRING', nullable: true },
    blocks: {
      type: 'ARRAY',
      items: {
        type: 'OBJECT',
        properties: {
          t: { type: 'STRING', enum: ['h1', 'h2', 'p', 'li', 'oli', 'table'] },
          text: { type: 'STRING' },
          rows: { type: 'ARRAY', items: { type: 'ARRAY', items: { type: 'STRING' } } },
        },
        required: ['t'],
      },
    },
  },
  required: ['blocks'],
};
const cors = origin => ({
  'Access-Control-Allow-Origin': origin,
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
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

function build(body) {
  const mode = body.mode;
  if (mode === 'translate') {
    const t = String(body.target || '');
    const text = String(body.text || '');
    if (!LANGS[t]) return { err: 'unknown target language' };
    if (!text.trim() || text.length > 60000) return { err: 'text missing or too long' };
    // 6.2: the app may add the form of address (formal / masculine / feminine / neutral) and the glossary rule
    const style = String(body.style || '').slice(0, 1500).replace(/[<>]/g, '');
    return {
      systemInstruction: { parts: [{ text: PROMPTS.translate(t) + (style ? '\n' + style : '') }] },
      contents: [{ role: 'user', parts: [{ text: `<page>\n${text}\n</page>` }] }],
      generationConfig: { temperature: 0.2, maxOutputTokens: 16384 },
    };
  }
  if (mode === 'scan') {
    const image = String(body.image || '');
    if (!image || image.length > MAX_IMAGE || !/^[A-Za-z0-9+/=]+$/.test(image.slice(0, 200))) return { err: 'image missing or too large' };
    const mime = /^image\/(jpeg|png|webp)$/.test(body.mime) ? body.mime : 'image/jpeg';
    const hint = LANGS[body.hint] || '';
    // the same page in sharp strips from top to bottom (the app sends them for an accurate reading)
    const strips = (Array.isArray(body.strips) ? body.strips : []).slice(0, 4).map(String)
      .filter(s => s && s.length <= MAX_IMAGE && /^[A-Za-z0-9+/=]+$/.test(s.slice(0, 200)));
    const parts = [{ inlineData: { mimeType: mime, data: image } }];
    if (strips.length) {
      parts.push({ text: `The same page again in ${strips.length} sharper strips, from top to bottom. Neighbouring strips overlap a little, so a line can appear in two strips: write it only once. Use the whole page above for the order and the layout, and the strips to read every word.` });
      for (const s of strips) parts.push({ inlineData: { mimeType: 'image/jpeg', data: s } });
    }
    parts.push({ text: 'Transcribe this page.' });
    return {
      systemInstruction: { parts: [{ text: PROMPTS.scan(hint) }] },
      contents: [{ role: 'user', parts }],
      // the exact shape of the answer: the fast model then always returns valid JSON
      generationConfig: { temperature: 0, maxOutputTokens: 16384, responseMimeType: 'application/json', responseSchema: SCAN_SCHEMA },
    };
  }
  if (mode === 'lines') {
    const imgs = (Array.isArray(body.lines) ? body.lines : []).slice(0, 10).map(String)
      .filter(s => s && s.length <= MAX_IMAGE && /^[A-Za-z0-9+/=]+$/.test(s.slice(0, 200)));
    if (!imgs.length) return { err: 'lines missing' };
    const hint = LANGS[body.hint] || '';
    const soft = (Array.isArray(body.lines2) ? body.lines2 : []).slice(0, 10).map(String);
    const vocab = vocabText(body.vocab);
    const parts = [];
    imgs.forEach((s, i) => {
      const b = soft[i] && soft[i].length <= MAX_IMAGE && /^[A-Za-z0-9+/=]+$/.test(soft[i].slice(0, 200)) ? soft[i] : null;
      parts.push({ text: `Line ${i + 1}${b ? ', view A' : ''}:` }, { inlineData: { mimeType: 'image/jpeg', data: s } });
      if (b) parts.push({ text: `Line ${i + 1}, view B:` }, { inlineData: { mimeType: 'image/jpeg', data: b } });
    });
    parts.push({ text: `Transcribe each of the ${imgs.length} lines.${vocab}` });
    return {
      systemInstruction: { parts: [{ text: LINES_PROMPT(hint) }] },
      contents: [{ role: 'user', parts }],
      generationConfig: { temperature: 0, maxOutputTokens: 8192, responseMimeType: 'application/json', responseSchema: LINES_SCHEMA },
    };
  }
  if (mode === 'tableCheck') {
    const img = String(body.image || '');
    if (!img || img.length > MAX_IMAGE || !/^[A-Za-z0-9+/=]+$/.test(img.slice(0, 200))) return { err: 'image missing' };
    const strips = (Array.isArray(body.strips) ? body.strips : []).slice(0, 4).map(String).filter(x => x && x.length <= MAX_IMAGE);
    const draft = JSON.stringify(body.draft || {}).slice(0, 60000);
    const parts = [{ text: 'The page:' }, { inlineData: { mimeType: 'image/jpeg', data: img } }];
    strips.forEach((x, i) => parts.push({ text: `Strip ${i + 1} of ${strips.length}:` }, { inlineData: { mimeType: 'image/jpeg', data: x } }));
    parts.push({ text: 'First reading:\n' + draft + '\nCheck it against the page and return the corrected data.' });
    return {
      systemInstruction: { parts: [{ text: TABLE_CHECK_PROMPT }] },
      contents: [{ role: 'user', parts }],
      generationConfig: { temperature: 0, maxOutputTokens: 16384, responseMimeType: 'application/json', responseSchema: TABLE_SCHEMA },
    };
  }
  if (mode === 'orient') {
    const img = String(body.image || '');
    if (!img || img.length > MAX_IMAGE || !/^[A-Za-z0-9+/=]+$/.test(img.slice(0, 200))) return { err: 'image missing' };
    return {
      systemInstruction: { parts: [{ text: 'You look at a photographed page and decide how it must be rotated so that its writing reads upright (lines horizontal, letters not upside down). Answer with the clockwise rotation needed: "0" if it is already upright, "90", "180" or "270". Judge by the letters and digits, not by the shape of the paper.' }] },
      contents: [{ role: 'user', parts: [{ inlineData: { mimeType: 'image/jpeg', data: img } }, { text: 'Which clockwise rotation makes this page upright?' }] }],
      generationConfig: { temperature: 0, maxOutputTokens: 256, responseMimeType: 'application/json', responseSchema: { type: 'OBJECT', properties: { rotate: { type: 'STRING', enum: ['0', '90', '180', '270'] } }, required: ['rotate'] } },
    };
  }
  if (mode === 'table') {
    const img = String(body.image || '');
    if (!img || img.length > MAX_IMAGE || !/^[A-Za-z0-9+/=]+$/.test(img.slice(0, 200))) return { err: 'image missing' };
    const strips = (Array.isArray(body.strips) ? body.strips : []).slice(0, 4).map(String).filter(x => x && x.length <= MAX_IMAGE);
    const parts = [{ text: 'The page:' }, { inlineData: { mimeType: 'image/jpeg', data: img } }];
    strips.forEach((x, i) => parts.push({ text: `Strip ${i + 1} of ${strips.length}:` }, { inlineData: { mimeType: 'image/jpeg', data: x } }));
    parts.push({ text: 'Return every table on the page as rows of cells.' });
    return {
      systemInstruction: { parts: [{ text: TABLE_PROMPT(LANGS[body.hint] || '') }] },
      contents: [{ role: 'user', parts }],
      generationConfig: { temperature: 0, maxOutputTokens: 16384, responseMimeType: 'application/json', responseSchema: TABLE_SCHEMA },
    };
  }
  if (mode === 'linesCheck') {
    // a second look at the whole page: every line picture + the first reading, read again together.
    // The page's subject and the words that repeat on it decide between readings the letters allow.
    const imgs = (Array.isArray(body.lines) ? body.lines : []).slice(0, 30).map(String)
      .filter(s => s && s.length <= MAX_IMAGE && /^[A-Za-z0-9+/=]+$/.test(s.slice(0, 200)));
    const draft = Array.isArray(body.draft) ? body.draft.slice(0, 30) : [];
    if (!imgs.length || draft.length !== imgs.length) return { err: 'lines or draft missing' };
    const parts = [];
    imgs.forEach((s, i) => { parts.push({ text: `Line ${i + 1}:` }, { inlineData: { mimeType: 'image/jpeg', data: s } }); });
    parts.push({ text: 'First reading, one entry per line:\n' + JSON.stringify(draft.map((d, i) => ({ n: i + 1, kind: String(d && d.kind || 'text'), text: String(d && d.text || '').slice(0, 600) }))) + vocabText(body.vocab) + '\nCheck every line against its picture and return the corrected reading.' });
    return {
      systemInstruction: { parts: [{ text: LINES_CHECK_PROMPT }] },
      contents: [{ role: 'user', parts }],
      generationConfig: { temperature: 0, maxOutputTokens: 8192, responseMimeType: 'application/json', responseSchema: LINES_SCHEMA },
    };
  }
  if (mode === 'corners') {
    const image = String(body.image || '');
    if (!image || image.length > MAX_IMAGE || !/^[A-Za-z0-9+/=]+$/.test(image.slice(0, 200))) return { err: 'image missing or too large' };
    return {
      systemInstruction: { parts: [{ text: CORNERS_PROMPT }] },
      contents: [{ role: 'user', parts: [{ inlineData: { mimeType: 'image/jpeg', data: image } }, { text: 'Find the document and its 4 corners.' }] }],
      generationConfig: { temperature: 0, maxOutputTokens: 512, responseMimeType: 'application/json', responseSchema: CORNERS_SCHEMA },
    };
  }
  if (mode === 'ask') {
    const text = String(body.text || '').slice(0, MAX_TEXT);
    const q = String(body.question || '').slice(0, 2000);
    if (!text.trim() || !q.trim()) return { err: 'document text or question missing' };
    const history = Array.isArray(body.history) ? body.history.slice(-6) : [];
    const contents = [{ role: 'user', parts: [{ text: `<document>\n${text}\n</document>` }] }, { role: 'model', parts: [{ text: 'קראתי את המסמך. מה השאלה?' }] }];
    for (const h of history) {
      if (!h || !h.q || !h.a) continue;
      contents.push({ role: 'user', parts: [{ text: String(h.q).slice(0, 2000) }] }, { role: 'model', parts: [{ text: String(h.a).slice(0, 6000) }] });
    }
    contents.push({ role: 'user', parts: [{ text: q }] });
    return {
      systemInstruction: { parts: [{ text: PROMPTS.ask() }] },
      contents,
      generationConfig: { temperature: 0.3, maxOutputTokens: 4096 },
    };
  }
  return { err: 'unknown mode' };
}

export default {
  async fetch(req, env, ctx) {
    const origin = req.headers.get('Origin') || '';
    const okOrigin = ALLOWED.includes(origin);
    if (req.method === 'OPTIONS') return new Response(null, { status: okOrigin ? 204 : 403, headers: okOrigin ? cors(origin) : {} });
    if (!okOrigin) return new Response(JSON.stringify({ error: { message: 'origin not allowed' } }), { status: 403, headers: { 'content-type': 'application/json' } });
    if (req.method !== 'POST') return fail('POST only', 405, origin);
    if (new URL(req.url).pathname === '/pass') { const p = await givePass(req, env); return json(p.body, p.status, origin); }
    if (!env.GEMINI_KEY) return fail('GEMINI_KEY is not set on the worker', 500, origin);

    let body;
    try { body = await req.json(); } catch (_) { return fail('bad JSON', 400, origin); }
    const g = build(body || {});
    if (g.err) return fail(g.err, 400, origin);
    const mode = body.mode === 'scan' && body.quality !== 'fast' ? 'scanAccurate' : body.mode === 'lines' && body.quality === 'fast' ? 'linesFast' : body.mode;
    // 6.4: a pass, and a fair daily share (or the person's own key)
    const gd = await guard(req, env, WEIGHT[mode] || 1);
    if (!gd.ok) return json(gd.body, gd.status, origin);
    env = gd.env;

    const deadline = Date.now() + (BUDGET[mode] || 85000);
    if (!STREAM_MODES.includes(mode)) {
      const res = await callGemini(g, mode, env, deadline);
      return json(res.d, res.status, origin);
    }
    // answer at once, keep the line open, write the result when it is ready
    const { readable, writable } = new TransformStream();
    const w = writable.getWriter(), enc = new TextEncoder();
    w.write(enc.encode(' '));
    const ping = setInterval(() => { w.write(enc.encode(' ')).catch(() => {}); }, 15000);
    const job = (async () => {
      let res;
      const t0 = Date.now();
      try { res = await callGemini(g, mode, env, deadline); } catch (e) { res = { status: 502, d: { error: { message: String(e && e.message || e) } } }; }
      // an accurate reading is checked once more against the picture, word by word (if there is time left)
      // only when the first reading was quick and there is time left for a second look
      if (mode === 'scanAccurate' && res.status === 200 && Date.now() - t0 < 50000 && deadline - Date.now() > 40000) {
        try { res = await proofread(g, res, env, deadline); } catch (_) {}
      }
      clearInterval(ping);
      const out = res.status === 200 ? res.d : { error: { ...(res.d && res.d.error || {}), message: (res.d && res.d.error && res.d.error.message) || 'Gemini error', status: res.status } };
      try { await w.write(enc.encode(JSON.stringify(out))); await w.close(); } catch (_) {}
    })();
    ctx.waitUntil(job);
    // octet-stream: Cloudflare does not compress (and hold back) it, so the keep-alive spaces really go out
    return new Response(readable, { status: 200, headers: { 'content-type': 'application/octet-stream', 'cache-control': 'no-store, no-transform', ...cors(origin) } });
  },
};

// second look: the same pictures + the first reading. Words that were guessed wrong, or that make no sense on
// this page, are read again from the letters. If anything goes wrong, the first reading stays.
const CHECK_PROMPT = `You check and correct a draft transcription of a photographed page (often handwritten Hebrew cursive).
You get the page pictures and the draft as JSON. Go over the page line by line and compare every word of the draft with the letters in the picture.
- Fix every word whose letters in the picture do not match the draft. Hebrew cursive letters are easy to confuse (ו/י/ן, ה/ח/ת, ב/כ/נ, ס/ם, ד/ר, ג/נ): look again at each one.
- A word that makes no sense in the sentence or the subject of the page is probably misread: find the reading the letters support that fits the rest of the page (the same words often repeat across the page — use that). Never invent words the letters do not support.
- Add lines the draft missed, remove lines that are not on the page, keep the order of the page.
- Keep [[ ]] around words that stay uncertain.
Return the full corrected transcription in exactly the same JSON format, and nothing else.`;
const textOf = d => ((d && d.candidates && d.candidates[0] && d.candidates[0].content && d.candidates[0].content.parts) || []).map(p => p.text || '').join('');
async function proofread(g, first, env, deadline) {
  const draft = textOf(first.d);
  let j; try { j = JSON.parse(draft.replace(/^\s*```(?:json)?|```\s*$/g, '')); } catch (_) { return first; }
  if (!j || !Array.isArray(j.blocks) || !j.blocks.length) return first;
  const pics = g.contents[0].parts.filter(p => p.inlineData || /strips/.test(p.text || ''));
  const g2 = {
    systemInstruction: { parts: [{ text: CHECK_PROMPT }] },
    contents: [{ role: 'user', parts: [...pics, { text: 'The draft transcription:\n' + JSON.stringify(j) + '\nCheck it against the page and return the corrected JSON.' }] }],
    generationConfig: g.generationConfig,
  };
  const r = await callGemini(g2, 'scanCheck', env, deadline);
  if (r.status !== 200) return first;
  try { const k = JSON.parse(textOf(r.d).replace(/^\s*```(?:json)?|```\s*$/g, '')); if (!k || !Array.isArray(k.blocks) || k.blocks.length < Math.ceil(j.blocks.length / 2)) return first; } catch (_) { return first; }
  return r;
}

// 6.3.1: a model that hangs (it happens: "gemini-flash-latest" stopped answering at all for a while, and every
// question waited 80 seconds for nothing) must not hold the user. So:
// · HEDGE: if the first model has not answered after this many ms, the next one starts too, at the same time,
//   and the first good answer wins (the other is cancelled). Short where any good model will do (a question,
//   a translation), long where the main model reads much better (handwriting), so it normally gets its chance.
// · a model that just hung or failed is remembered for a few minutes (in this worker copy): the next requests
//   start with the other model at once, and the sick one only joins after the hedge time.
const HEDGE = { ask: 9000, translate: 8000, corners: 6000, orient: 6000, scan: 15000, linesFast: 15000, scanAccurate: 50000, lines: 45000, table: 50000 };
const SICK_MS = 4 * 60 * 1000;
const sick = new Map();   // model → time until which it is skipped first
const isSick = m => (sick.get(m) || 0) > Date.now();

// one model, with its thinking settings tried in turn (a model that does not take one answers 400)
async function runModel(model, g, mode, env, deadline, stop) {
  let last = null;
  const t0 = Date.now();
  const options = THINKING[mode] || [null];
  for (let k = 0; k < options.length; k++) {
    const thinking = options[k];
    const left = deadline - Date.now();
    if (left < 6000 || stop.signal.aborted) return last || { status: 504, d: { error: { message: 'took too long' } } };
    const body = thinking ? { ...g, generationConfig: { ...g.generationConfig, thinkingConfig: thinking } } : g;
    const ctl = new AbortController();
    const onStop = () => ctl.abort();
    stop.signal.addEventListener('abort', onStop, { once: true });
    const timer = setTimeout(() => ctl.abort(), Math.min(MODEL_TIMEOUT[mode] || 90000, left));
    let r, d;
    try {
      r = await fetch(`https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent`, {
        method: 'POST',
        headers: { 'content-type': 'application/json', 'x-goog-api-key': env.GEMINI_KEY },
        body: JSON.stringify(body),
        signal: ctl.signal,
      });
      d = await r.json().catch(() => ({}));
    } catch (e) {
      clearTimeout(timer); stop.signal.removeEventListener('abort', onStop);
      // hung, or still silent when the other model had already answered: next time it goes second
      if (!stop.signal.aborted || Date.now() - t0 > (HEDGE[mode] || 0)) sick.set(model, Date.now() + SICK_MS);
      return { status: 504, d: { error: { message: ctl.signal.aborted ? `${model} took too long` : 'Gemini not reachable' } } };
    }
    clearTimeout(timer); stop.signal.removeEventListener('abort', onStop);
    if (r.ok) { sick.delete(model); return { status: 200, d }; }
    last = { status: r.status, d };
    if (r.status === 400 && thinking) continue;   // a setting this model does not take: the next setting
    if ([429, 500, 503, 504].includes(r.status)) sick.set(model, Date.now() + SICK_MS);
    return last;
  }
  return last;
}

// the main model, then the lighter one (it has its own free quota and is faster) — see HEDGE above
async function callGemini(g, mode, env, deadline = Date.now() + 90000) {
  let models = (MODELS_FOR[mode] || MODELS).slice();
  // a model that hung a moment ago goes after the healthy one (not where the main model reads much better:
  // handwriting keeps its order and the long hedge)
  if (models.length > 1 && (HEDGE[mode] || 0) <= 15000 && isSick(models[0]) && !isSick(models[1])) models = [models[1], models[0], ...models.slice(2)];
  const hedge = HEDGE[mode];
  const stop = new AbortController();
  return await new Promise(resolve => {
    let next = 0, running = 0, done = false, last = null, timer = null;
    const finish = res => { if (done) return; done = true; clearTimeout(timer); stop.abort(); resolve(res); };
    const failOut = () => finish(last && last.d && last.d.error ? last : { status: 502, d: { error: { message: 'Gemini error' } } });
    const launch = () => {
      if (done || next >= models.length) return;
      const model = models[next++]; running++;
      // the next model joins after the hedge time if this one is still working
      clearTimeout(timer);
      if (hedge != null && next < models.length) timer = setTimeout(launch, hedge);
      runModel(model, g, mode, env, deadline, stop).then(res => {
        running--;
        if (done) return;
        if (res && res.status === 200) return finish(res);
        if (res) last = res;
        // 400 = this model refused the request itself; the other model may still take it. Other errors (403 key…): stop.
        if (res && ![400, 429, 500, 503, 504, 404].includes(res.status) && !running) return failOut();
        if (next < models.length) launch();          // failed early: the next model starts now, not after the hedge
        else if (!running) failOut();
      });
    };
    launch();
  });
}
