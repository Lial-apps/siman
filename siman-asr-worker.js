// סימן — מנוע התמלול שרץ בתוך המכשיר (Web Worker, נטען מ-siman-voice.js)
// השמע לא יוצא מהמכשיר. מה שיורד מהאינטרנט הוא רק המנוע עצמו, פעם אחת:
//   עברית: Ozen-v1 (itayinbar/Ozen-v1, Apache-2.0) — מודל עברי קטן שזוקק מ-ivrit.ai whisper-large-v3-turbo
//   שפות אחרות: Whisper small (Xenova/whisper-small)
//   מדויק (מחשב עם כרטיס מסך): ivrit-ai/whisper-large-v3-turbo בגרסה לדפדפן — המאגר שתהליך ההמרה בגיטהאב יצר
// הקבצים נשמרים במטמון של הדפדפן ("transformers-cache"), ואחר כך התמלול עובד גם בלי אינטרנט.
// הספרייה: transformers.min.js (Transformers.js 4.3.1), מונחת ליד האפליקציה.
import { env, AutoTokenizer, AutoProcessor, WhisperForConditionalGeneration, pipeline, AutoModelForAudioFrameClassification, AutoModel, AutoFeatureExtractor } from './transformers.min.js';

env.allowLocalModels = false;
env.useBrowserCache = true;

// where the engine files come from: huggingface.co, or — where it is blocked (some phones, browsers, networks) —
// the same files through סימן's own server (siman-voice worker, /hf/ and /cdn/)
const HF = 'https://huggingface.co/';
const MIRROR = 'https://siman-voice.siman-app.workers.dev/hf/';
const ORT_MIRROR = 'https://siman-voice.siman-app.workers.dev/cdn/';
const PROBE = 'itayinbar/Ozen-v1/resolve/main/config.json';
let hostReady = null;
async function reachable(url) { try { const r = await fetch(url, { cache: 'no-store' }); return r.ok; } catch (_) { return false; } }
function useMirror() {
  env.remoteHost = MIRROR;
  // the runtime's own files (WebAssembly) normally come from cdn.jsdelivr.net: the same files through the mirror
  try {
    const wasm = env.backends.onnx.wasm, p = wasm.wasmPaths, to = u => String(u).replace('https://cdn.jsdelivr.net/npm/', ORT_MIRROR);
    if (p && typeof p === 'object') { if (p.mjs) p.mjs = to(p.mjs); if (p.wasm) p.wasm = to(p.wasm); }
    else if (typeof p === 'string') wasm.wasmPaths = to(p);
  } catch (_) {}
}
function pickHost() {
  return hostReady || (hostReady = (async () => {
    try {
      const c = await caches.open('transformers-cache');
      if (await c.match(MIRROR + PROBE)) { useMirror(); return 'mirror'; }   // already downloaded that way
      if (await c.match(HF + PROBE)) return 'hf';
    } catch (_) {}
    if (await reachable(HF + PROBE)) return 'hf';
    useMirror(); return 'mirror';
  })());
}

const ENGINES = {
  // the browser files the model's author checked: fp16 encoder + fp32 decoder (an fp16 decoder does not load,
  // an int8 encoder changes the text)
  he: { id: 'itayinbar/Ozen-v1', dtype: { encoder_model: 'fp16', decoder_model_merged: 'fp32' } },
  multi: { id: 'Xenova/whisper-small', dtype: { encoder_model: 'q8', decoder_model_merged: 'q8' } },
};
const SR = 16000;
let E = null;            // the loaded engine: { key, device, run(audio, lang), dispose() }
let queue = Promise.resolve();

const progress = p => {
  if (!p || !p.file) return;
  if (p.status === 'progress') postMessage({ type: 'progress', file: p.file, loaded: p.loaded || 0, total: p.total || 0 });
  else if (p.status === 'done') postMessage({ type: 'progress', file: p.file, done: true, total: p.total || 0 });
};

async function hasGpu() {
  try { if (!self.navigator || !navigator.gpu) return false; return !!(await navigator.gpu.requestAdapter()); } catch (_) { return false; }
}

async function gpuF16() {
  try { const a = navigator.gpu && await navigator.gpu.requestAdapter(); return a ? a.features.has('shader-f16') : null; } catch (_) { return null; }
}
async function build(key, device, model) {
  if (key === 'acc') {
    // ivrit.ai whisper-large-v3-turbo for the browser (the repo the GitHub workflow made): graphics card only
    const f16 = await gpuF16();
    if (f16 === null) throw new Error('no-webgpu');
    const [tok, proc, m] = await Promise.all([
      AutoTokenizer.from_pretrained(model, { progress_callback: progress }),
      AutoProcessor.from_pretrained(model, { progress_callback: progress }),
      WhisperForConditionalGeneration.from_pretrained(model, { dtype: { encoder_model: f16 ? 'fp16' : 'q4', decoder_model_merged: 'q4' }, device: 'webgpu', progress_callback: progress }),
    ]);
    return {
      key, device: 'webgpu', model, dispose: () => m.dispose(),
      run: async (audio, lang) => {
        const { input_features } = await proc(audio);
        const ids = await m.generate({ inputs: input_features, language: lang || 'he', task: 'transcribe', max_new_tokens: 220 });
        return tok.batch_decode(ids, { skip_special_tokens: true })[0] || '';
      },
    };
  }
  const c = ENGINES[key];
  if (key === 'he') {
    const [tok, proc, model] = await Promise.all([
      AutoTokenizer.from_pretrained(c.id, { progress_callback: progress }),
      AutoProcessor.from_pretrained(c.id, { progress_callback: progress }),
      WhisperForConditionalGeneration.from_pretrained(c.id, { dtype: c.dtype, device, progress_callback: progress }),
    ]);
    // Ozen is Hebrew-only: it starts from <|startoftranscript|> alone — no language, task or
    // no-timestamps token (the library would add <|notimestamps|> on its own, so the start is given here)
    const start = (model.generation_config && model.generation_config.decoder_start_token_id) ?? 1;
    return {
      key, device,
      dispose: () => model.dispose(),
      run: async audio => {
        const { input_features } = await proc(audio);          // padded to a 30-second window
        const ids = await model.generate({ inputs: input_features, decoder_input_ids: [start], max_new_tokens: 220 });
        return tok.batch_decode(ids, { skip_special_tokens: true })[0] || '';
      },
    };
  }
  const pipe = await pipeline('automatic-speech-recognition', c.id, { dtype: c.dtype, device, progress_callback: progress });
  return {
    key, device,
    dispose: () => pipe.dispose(),
    run: async (audio, lang) => { const r = await pipe(audio, { language: lang || 'en', task: 'transcribe' }); return (r && r.text) || ''; },
  };
}

async function load(key, model) {
  if (E && E.key === key && (key !== 'acc' || E.model === model)) return E;
  const host = await pickHost();
  if (key === 'acc') {
    if (!model) throw new Error('accurate engine not installed');
    if (E) { try { await E.dispose(); } catch (_) {} E = null; }
    try { E = await build('acc', 'webgpu', model); await E.run(new Float32Array(SR), 'he'); return E; }
    catch (err) { E = null; throw new Error('acc@' + host + ': ' + ((err && err.message) || err)); }
  }
  if (E) { try { await E.dispose(); } catch (_) {} E = null; }
  // the graphics card when there is one (much faster); otherwise the processor
  const devices = key === 'he' && await hasGpu() ? ['webgpu', 'wasm'] : ['wasm'];
  const fails = [];
  for (const device of devices) {
    try {
      const e = await build(key, device);
      // some graphics cards load the model but cannot run it (no 16-bit support): one silent second shows it now
      if (device === 'webgpu') await e.run(new Float32Array(SR), 'he');
      E = e;
      return E;
    } catch (err) { fails.push(device + '@' + host + ': ' + ((err && err.message) || err)); console.warn('siman-asr:', device, err); }
  }
  throw new Error(fails.join(' | ') || 'engine did not load');
}

async function handle(m) {
  try {
    if (m.type === 'load') {
      const e = await load(m.engine, m.model);
      postMessage({ id: m.id, device: e.device, host: await pickHost() });
    } else if (/^dz-/.test(m.type)) {
      postMessage(Object.assign({ id: m.id }, await dzHandle(m)));
    } else if (m.type === 'run') {
      const key = m.engine || ((m.lang || 'he') === 'he' ? 'he' : 'multi');
      const e = await load(key, m.model);
      const t0 = performance.now();
      let text;
      try { text = await e.run(m.audio, m.lang); }
      catch (err) {
        // a graphics-card failure in the middle: once more on the processor
        if (e.device !== 'wasm' && key !== 'acc') { try { await e.dispose(); } catch (_) {} E = await build(key, 'wasm'); text = await E.run(m.audio, m.lang); }
        else throw err;
      }
      postMessage({ id: m.id, text: String(text || '').trim(), ms: Math.round(performance.now() - t0) });
    }
  } catch (err) {
    postMessage({ id: m.id, error: String((err && err.message) || err) });
  }
}

/* ---------------- who spoke when: speaker separation, also on this device ----------------
   pyannote segmentation-3.0 (MIT) finds up to 3 voices inside every 10-second window (5-second steps);
   WeSpeaker ResNet34 (voxceleb, CC-BY-4.0) gives each voice a "voiceprint"; voiceprints that are alike are
   joined into one speaker across the whole recording. About 32MB, downloaded once like the main engine. */
const DZ = { seg: 'onnx-community/pyannote-segmentation-3.0', emb: 'onnx-community/wespeaker-voxceleb-resnet34-LM' };
let D = null;          // { seg, segProc, emb, embProc }
let DS = null;         // the recording being separated: windows and voiceprints
async function dzLoad() {
  if (D) return D;
  await pickHost();
  const [seg, segProc, emb, embProc] = await Promise.all([
    AutoModelForAudioFrameClassification.from_pretrained(DZ.seg, { device: 'wasm', dtype: 'fp32', progress_callback: progress }),
    AutoProcessor.from_pretrained(DZ.seg, { progress_callback: progress }),
    AutoModel.from_pretrained(DZ.emb, { device: 'wasm', dtype: 'fp32', progress_callback: progress }),
    AutoFeatureExtractor.from_pretrained(DZ.emb, { progress_callback: progress }),
  ]);
  D = { seg, segProc, emb, embProc };
  return D;
}
const softmaxRow = (a, o, n, out) => { let m = -Infinity; for (let i = 0; i < n; i++) if (a[o + i] > m) m = a[o + i]; let s = 0; for (let i = 0; i < n; i++) { out[i] = Math.exp(a[o + i] - m); s += out[i]; } for (let i = 0; i < n; i++) out[i] /= s; };
// the 7 classes of segmentation-3.0: none, A, B, C, A+B, A+C, B+C
const HAS = [[], [0], [1], [2], [0, 1], [0, 2], [1, 2]];
async function voiceprint(audio) {
  const inp = await D.embProc(audio);
  const out = await D.emb(inp);
  const t = out.embeddings || out.last_hidden_state || out.pooler_output || Object.values(out)[0];
  const v = Float32Array.from(t.data);
  let n = 0; for (const x of v) n += x * x; n = Math.sqrt(n) || 1;
  for (let i = 0; i < v.length; i++) v[i] /= n;
  return v;
}
async function dzWindow(audio, offset, first, last) {
  const { logits } = await D.seg(await D.segProc(audio));
  const [, F, C] = logits.dims, L = logits.data;
  const fd = audio.length / F / SR;                 // seconds per frame
  const act = [new Float32Array(F), new Float32Array(F), new Float32Array(F)], lab = new Int8Array(F), p = new Float32Array(C);
  for (let f = 0; f < F; f++) {
    softmaxRow(L, f * C, C, p);
    let best = 0; for (let c = 1; c < C; c++) if (p[c] > p[best]) best = c;
    lab[f] = best;
    for (let c = 1; c < C; c++) for (const k of HAS[c]) act[k][f] += p[c];
  }
  const w = { offset, fd, F, act, local: [null, null, null], runs: [] };
  // a voiceprint for each voice heard alone for at least 0.6 s in this window
  for (let k = 0; k < 3; k++) {
    const parts = []; let n = 0;
    for (let f = 0; f < F; f++) if (lab[f] === 1 + k && act[k][f] > 0.5) { const a = Math.floor(f * fd * SR), b = Math.floor((f + 1) * fd * SR); parts.push([a, b]); n += b - a; }
    if (n >= 0.6 * SR) {
      const buf = new Float32Array(Math.min(n, 6 * SR)); let o = 0;
      for (const [a, b] of parts) { if (o >= buf.length) break; const take = Math.min(b - a, buf.length - o); buf.set(audio.subarray(a, a + take), o); o += take; }
      try { w.local[k] = { v: await voiceprint(buf), sec: n / SR }; } catch (e) { console.warn('voiceprint', e); }
    }
  }
  // what this window decides: its middle 5 seconds (the edges belong to the neighbouring windows)
  const c0 = first ? 0 : 2.5, c1 = last ? Infinity : 7.5;
  let cur = null;
  for (let f = 0; f < F; f++) {
    const t = f * fd; if (t < c0 || t >= c1) { cur = null; continue; }
    let k = -1, m = 0.5; for (let j = 0; j < 3; j++) if (act[j][f] > m) { m = act[j][f]; k = j; }
    if (cur && cur.k === k) cur.e = offset + (f + 1) * fd;
    else { cur = { k, s: offset + f * fd, e: offset + (f + 1) * fd }; w.runs.push(cur); }
  }
  w.runs = w.runs.filter(r => r.k >= 0);
  DS.windows.push(w);
}
const dot = (a, b) => { let s = 0; for (let i = 0; i < a.length; i++) s += a[i] * b[i]; return s; };
function clusterAll(K) {
  const items = [];
  DS.windows.forEach((w, wi) => w.local.forEach((l, k) => { if (l) items.push({ wi, k, v: l.v, sec: l.sec }); }));
  if (!items.length) return new Map();
  // 1. very similar voiceprints are joined first (keeps the next step small)
  let C = [];
  for (const it of items) {
    let best = null, bs = 0.72;
    for (const c of C) { const s = dot(c.n, it.v); if (s > bs) { bs = s; best = c; } }
    if (!best) { C.push({ sum: Float32Array.from(it.v, x => x * it.sec), sec: it.sec, n: it.v, items: [it] }); continue; }
    for (let i = 0; i < it.v.length; i++) best.sum[i] += it.v[i] * it.sec;
    best.sec += it.sec; best.items.push(it);
    let nn = 0; for (const x of best.sum) nn += x * x; nn = Math.sqrt(nn) || 1; best.n = best.sum.map(x => x / nn);
  }
  // 2. then the most alike groups are joined, until the number of speakers asked for — or, when not given,
  //    until the closest two are clearly different voices (pyannote's threshold for this voiceprint model)
  const T = 0.30;
  const merge = (a, b) => { for (let i = 0; i < a.sum.length; i++) a.sum[i] += b.sum[i]; a.sec += b.sec; a.items.push(...b.items); let nn = 0; for (const x of a.sum) nn += x * x; nn = Math.sqrt(nn) || 1; a.n = a.sum.map(x => x / nn); };
  // (a table of how alike every two groups are; after each join only one row is worked out again)
  const N = C.length, S = new Float32Array(N * N), alive = new Uint8Array(N).fill(1);
  for (let i = 0; i < N; i++) for (let j = i + 1; j < N; j++) S[i * N + j] = S[j * N + i] = dot(C[i].n, C[j].n);
  let left = N;
  while (left > 1) {
    let bi = -1, bj = -1, bs = -2;
    for (let i = 0; i < N; i++) if (alive[i]) for (let j = i + 1; j < N; j++) if (alive[j] && S[i * N + j] > bs) { bs = S[i * N + j]; bi = i; bj = j; }
    if (K ? left <= K : bs < T) break;
    merge(C[bi], C[bj]); alive[bj] = 0; left--;
    for (let j = 0; j < N; j++) if (alive[j] && j !== bi) S[bi * N + j] = S[j * N + bi] = dot(C[bi].n, C[j].n);
  }
  C = C.filter((_, i) => alive[i]);
  // 3. without a given number: a "speaker" with very little speech is joined to the closest one
  if (!K) {
    const total = C.reduce((a, c) => a + c.sec, 0);
    C.sort((a, b) => b.sec - a.sec);
    for (let i = C.length - 1; i > 0; i--) if (C[i].sec < total * 0.04 || C[i].sec < 4) {
      let bj = 0, bs = -2; for (let j = 0; j < i; j++) { const s = dot(C[i].n, C[j].n); if (s > bs) { bs = s; bj = j; } }
      merge(C[bj], C[i]); C.splice(i, 1);
    }
  }
  const map = new Map();
  C.forEach((c, ci) => c.items.forEach(it => map.set(it.wi + ':' + it.k, ci)));
  return map;
}
// a voice with no voiceprint in its window: the voice it overlaps in the window before (or after)
function linkByOverlap(map) {
  const W = DS.windows;
  const corr = (wa, ka, wb, kb, aFrom) => { // wa's frames from aFrom, against wb's frames from 0
    const A = W[wa].act[ka], B = W[wb].act[kb], n = Math.min(A.length - aFrom, B.length);
    let s = 0; for (let i = 0; i < n; i++) s += A[aFrom + i] * B[i]; return s / Math.max(1, n);
  };
  for (let pass = 0; pass < 2; pass++) for (let wi = 0; wi < W.length; wi++) for (let k = 0; k < 3; k++) {
    const key = wi + ':' + k; if (map.has(key)) continue;
    if (!W[wi].runs.some(r => r.k === k)) continue;
    let best = null, bs = 0.2;
    if (wi > 0) { const half = Math.floor(W[wi - 1].F / 2); for (let j = 0; j < 3; j++) { const m = map.get((wi - 1) + ':' + j); if (m == null) continue; const s = corr(wi - 1, j, wi, k, half); if (s > bs) { bs = s; best = m; } } }
    if (wi + 1 < W.length) { const half = Math.floor(W[wi].F / 2); for (let j = 0; j < 3; j++) { const m = map.get((wi + 1) + ':' + j); if (m == null) continue; const s = corr(wi, k, wi + 1, j, half); if (s > bs) { bs = s; best = m; } } }
    if (best != null) map.set(key, best);
  }
}
function dzEnd(K) {
  const map = clusterAll(K);
  linkByOverlap(map);
  let turns = [];
  DS.windows.forEach((w, wi) => w.runs.forEach(r => { const c = map.get(wi + ':' + r.k); if (c != null) turns.push({ s: r.s, e: r.e, spk: c }); }));
  turns.sort((a, b) => a.s - b.s);
  // smoothing: same speaker across a short gap is one turn; a flicker under 0.4 s goes to its neighbour
  const out = [];
  for (const t of turns) { const p = out[out.length - 1]; if (p && p.spk === t.spk && t.s - p.e < 0.5) p.e = Math.max(p.e, t.e); else out.push({ ...t }); }
  for (let i = 0; i < out.length; i++) {
    const t = out[i]; if (t.e - t.s >= 0.4 || out.length < 2) continue;
    const p = out[i - 1], n = out[i + 1];
    const to = p && n ? ((p.e - p.s) >= (n.e - n.s) ? p : n) : (p || n);
    if (to === p) p.e = Math.max(p.e, t.e); else { n.s = Math.min(n.s, t.s); }
    out.splice(i, 1); i--;
  }
  const merged = [];
  for (const t of out) { const p = merged[merged.length - 1]; if (p && p.spk === t.spk && t.s - p.e < 0.8) p.e = Math.max(p.e, t.e); else merged.push(t); }
  const order = new Map(); merged.forEach(t => { if (!order.has(t.spk)) order.set(t.spk, order.size); t.spk = order.get(t.spk); });
  const n = order.size;
  DS = null;
  return { turns: merged.map(t => ({ s: +t.s.toFixed(3), e: +t.e.toFixed(3), spk: t.spk })), speakers: n };
}
async function dzHandle(m) {
  if (m.type === 'dz-load') { await dzLoad(); return {}; }
  if (m.type === 'dz-begin') { await dzLoad(); DS = { windows: [] }; return {}; }
  if (m.type === 'dz-win') { await dzWindow(m.audio, m.offset, !!m.first, !!m.last); return {}; }
  if (m.type === 'dz-end') return dzEnd(m.speakers ? +m.speakers : 0);
  return null;
}

// one job at a time: the model is not run twice in parallel
self.onmessage = ({ data }) => { queue = queue.then(() => handle(data)); };
