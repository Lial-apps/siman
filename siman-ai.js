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
const STREAM_MODES = ['scan', 'scanAccurate', 'ask'];
const MODEL_TIMEOUT = { scan: 80000, scanAccurate: 170000, ask: 80000, translate: 80000 }; // ms per model before trying the next one
// reading a page: "accurate" = the main model, thinking a little (best for handwriting, ~20–90 s);
// "fast" = the light model, no thinking (a few seconds, guesses more)
const MODELS_FOR = { scan: ['gemini-flash-lite-latest', 'gemini-flash-latest'], scanAccurate: ['gemini-flash-latest', 'gemini-flash-lite-latest'] };
// no "thinking" for reading and translating = a much faster answer; a little for questions
// Models differ in which setting they take (older: thinkingBudget, newer: thinkingLevel), and one that does
// not take a setting answers "invalid argument" — so each is tried in turn, and at the end none at all.
const THINKING = {
  scan: [{ thinkingBudget: 0 }, { thinkingLevel: 'minimal' }, { thinkingLevel: 'low' }, null],
  // handwriting is read better with more thinking: each step down is tried only if the model refuses the one above
  scanAccurate: [{ thinkingLevel: 'medium' }, { thinkingLevel: 'low' }, { thinkingBudget: 4096 }, null],
  ask: [{ thinkingBudget: 512 }, { thinkingLevel: 'low' }, null],
  translate: [{ thinkingBudget: 0 }, { thinkingLevel: 'minimal' }, null],
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
- Read slowly, letter by letter. Write the writer's words exactly as written. Do not correct spelling, do not complete, summarize or translate.
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
  'Access-Control-Allow-Headers': 'content-type',
  'Access-Control-Max-Age': '86400',
  'Vary': 'Origin',
});
const json = (obj, status, origin) => new Response(JSON.stringify(obj), { status, headers: { 'content-type': 'application/json; charset=utf-8', ...cors(origin) } });
const fail = (msg, status, origin) => json({ error: { message: msg } }, status, origin);

function build(body) {
  const mode = body.mode;
  if (mode === 'translate') {
    const t = String(body.target || '');
    const text = String(body.text || '');
    if (!LANGS[t]) return { err: 'unknown target language' };
    if (!text.trim() || text.length > 60000) return { err: 'text missing or too long' };
    return {
      systemInstruction: { parts: [{ text: PROMPTS.translate(t) }] },
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
    if (!env.GEMINI_KEY) return fail('GEMINI_KEY is not set on the worker', 500, origin);

    let body;
    try { body = await req.json(); } catch (_) { return fail('bad JSON', 400, origin); }
    const g = build(body || {});
    if (g.err) return fail(g.err, 400, origin);
    const mode = body.mode === 'scan' && body.quality !== 'fast' ? 'scanAccurate' : body.mode;

    if (!STREAM_MODES.includes(mode)) {
      const res = await callGemini(g, mode, env);
      return json(res.d, res.status, origin);
    }
    // answer at once, keep the line open, write the result when it is ready
    const { readable, writable } = new TransformStream();
    const w = writable.getWriter(), enc = new TextEncoder();
    w.write(enc.encode(' '));
    const ping = setInterval(() => { w.write(enc.encode(' ')).catch(() => {}); }, 15000);
    const job = (async () => {
      let res;
      try { res = await callGemini(g, mode, env); } catch (e) { res = { status: 502, d: { error: { message: String(e && e.message || e) } } }; }
      clearInterval(ping);
      const out = res.status === 200 ? res.d : { error: { ...(res.d && res.d.error || {}), message: (res.d && res.d.error && res.d.error.message) || 'Gemini error', status: res.status } };
      try { await w.write(enc.encode(JSON.stringify(out))); await w.close(); } catch (_) {}
    })();
    ctx.waitUntil(job);
    // octet-stream: Cloudflare does not compress (and hold back) it, so the keep-alive spaces really go out
    return new Response(readable, { status: 200, headers: { 'content-type': 'application/octet-stream', 'cache-control': 'no-store, no-transform', ...cors(origin) } });
  },
};

// the main model, then the lighter one (it has its own free quota and is faster)
async function callGemini(g, mode, env) {
  let last = null;
  for (const model of MODELS_FOR[mode] || MODELS) {
    const options = THINKING[mode] || [null];
    for (let k = 0; k < options.length; k++) {
      const thinking = options[k];
      const body = thinking ? { ...g, generationConfig: { ...g.generationConfig, thinkingConfig: thinking } } : g;
      const ctl = new AbortController();
      const timer = setTimeout(() => ctl.abort(), MODEL_TIMEOUT[mode] || 90000);
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
        clearTimeout(timer);
        last = { status: 504, d: { error: { message: ctl.signal.aborted ? `${model} took too long` : 'Gemini not reachable' } } };
        break; // next model
      }
      clearTimeout(timer);
      if (r.ok) return { status: 200, d };
      last = { status: r.status, d };
      // a setting this model does not take: the same request with the next setting
      if (r.status === 400 && thinking) continue;
      break;
    }
    // 400 here = this model refused the request itself; the other model may still take it
    if (last && ![400, 429, 500, 503, 504, 404].includes(last.status)) break;
  }
  return last && last.d && last.d.error ? last : { status: 502, d: { error: { message: 'Gemini error' } } };
}
