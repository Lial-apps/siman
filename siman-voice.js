/* =====================================================================
   סימן — תמלול במכשיר (siman-voice.js)
   שיחות מוקלטות, שיעורים, הרצאות, הודעות קוליות והכתבה → טקסט שאפשר לערוך,
   עם נגן מסונכרן, חיפוש, דוברים, וייצוא ל-Word / PDF / כתוביות.

   פרטיות: התמלול רץ בתוך הדפדפן, במכשיר של המשתמש. השמע לא נשלח לשום שרת.
   המנוע: Ozen-v1 — מודל עברי פתוח (Apache-2.0) שזוקק מהמודל של ivrit.ai,
   רץ עם Transformers.js (WebGPU כשיש, אחרת WASM). יורד פעם אחת מ-Hugging Face
   ונשמר במכשיר; אחר כך התמלול עובד גם בלי אינטרנט.
   שפות אחרות: Whisper small, גם הוא במכשיר.
   סיכום / שאלות / תרגום / פרוטוקול: רק בלחיצה ובאישור, ורק הטקסט נשלח (SMART_CLOUD).
   נטען אחרי הסקריפט הראשי ו-siman-plus.js. מנוע התמלול: siman-asr-worker.js.
   ===================================================================== */
(() => {
'use strict';
const SMART_CLOUD = true;   // false = שום דבר לא נשלח לשרת: רק תמלול ועריכה במכשיר
const VOICE_URL = 'https://siman-voice.siman-app.workers.dev/';
const SR = 16000;                       // the engine hears 16 kHz mono
const MIN_CUT = 12, MAX_CHUNK = 27;     // a line is cut at a pause between 12 and 27 seconds
const ENGINES = {
  he: { mb: 165, probe: 'https://huggingface.co/itayinbar/Ozen-v1/resolve/main/onnx/encoder_model_fp16.onnx' },
  multi: { mb: 250, probe: 'https://huggingface.co/Xenova/whisper-small/resolve/main/onnx/encoder_model_quantized.onnx' },
};
const MEDIA_EXT = /\.(mp3|mpeg|wav|ogg|oga|opus|m4a|aac|flac|weba|webm|amr|3gp|3gpp|caf|wma|mp4|m4v|mov|mkv|aiff?)$/i;
const RTL_LANGS = ['he', 'ar', 'fa', 'ur', 'yi'];
const COLORS = ['#2E58CC', '#C2410C', '#0F766E', '#7C3AED', '#B4234A', '#0369A1', '#8A6A00', '#9D3AA8'];
const LANG_NAMES = { he: 'עברית', en: 'English', ar: 'العربية', ru: 'Русский', fr: 'Français', es: 'Español', de: 'Deutsch' };
const TR_TARGETS = { he: 'עברית', en: 'English', ar: 'العربية', ru: 'Русский', fr: 'Français', es: 'Español' };
const KINDS = {
  lesson: { label: 'שיעור או הרצאה' },
  call: { label: 'שיחה' },
  meeting: { label: 'ישיבה' },
  voice: { label: 'הודעה קולית' },
  dictation: { label: 'הכתבה למסמך' },
};
const PREF_KEY = 'siman.voice.prefs', OK_KEY = 'siman.voice.ok', CLOUD_KEY = 'siman.voice.cloud';

/* ---------------- small helpers ---------------- */
const el = (tag, cls, html) => { const e = document.createElement(tag); if (cls) e.className = cls; if (html != null) e.innerHTML = html; return e; };
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const say = m => { try { toast(m); } catch (_) { console.log(m); } };
const count = f => { try { window.simanCount && window.simanCount(f); } catch (_) {} };
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const ext = n => { const m = String(n || '').match(/\.([a-z0-9]+)$/i); return m ? m[1].toLowerCase() : ''; };
const isMedia = f => !!f && (/^(audio|video)\//.test(f.type || '') || MEDIA_EXT.test(f.name || ''));
const isVideo = f => /^video\//.test(f.type || '') || /\.(mp4|m4v|mov|mkv)$/i.test(f.name || '');
const clock = s => { if (s == null || isNaN(s)) return ''; s = Math.max(0, Math.floor(s)); const h = Math.floor(s / 3600), m = Math.floor(s % 3600 / 60), x = s % 60; return (h ? h + ':' + String(m).padStart(2, '0') : m) + ':' + String(x).padStart(2, '0'); };
const clock2 = s => { s = Math.max(0, Math.floor(s || 0)); const h = Math.floor(s / 3600), m = Math.floor(s % 3600 / 60), x = s % 60; return (h ? String(h).padStart(2, '0') + ':' : '') + String(m).padStart(2, '0') + ':' + String(x).padStart(2, '0'); };
const longDur = s => { if (!s) return ''; s = Math.round(s); if (s < 60) return s + ' שנ׳'; const m = Math.round(s / 60); return m < 60 ? m + ' דק׳' : Math.floor(m / 60) + ' שע׳ ' + (m % 60 ? (m % 60) + ' דק׳' : ''); };
const dayStr = t => { try { return new Date(t).toLocaleDateString('he-IL', { day: 'numeric', month: 'short', year: 'numeric' }); } catch (_) { return ''; } };
const CE = (() => { try { const d = document.createElement('div'); d.contentEditable = 'plaintext-only'; return d.contentEditable === 'plaintext-only' ? 'plaintext-only' : 'true'; } catch (_) { return 'true'; } })();
const isRtl = l => RTL_LANGS.includes(String(l || '').slice(0, 2));
const guessRtl = text => { const s = String(text || '').slice(0, 2000); const he = (s.match(/[\u0590-\u06FF]/g) || []).length, la = (s.match(/[A-Za-z\u0400-\u04FF]/g) || []).length; return he >= la; };
const rd = (k, d) => { try { return JSON.parse(localStorage.getItem(k) || 'null') || d; } catch (_) { return d; } };
const wr = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch (_) {} };
const sleep = ms => new Promise(r => setTimeout(r, ms));
const ico = {
  back: '<svg viewBox="0 0 24 24"><path d="M9 5l7 7-7 7"/></svg>',
  file: '<svg viewBox="0 0 24 24"><path d="M7 3h7l4 4v14H7z"/><path d="M14 3v4h4"/><path d="M10.5 17.5v-5l4-1v5"/><circle cx="9.6" cy="17.6" r="1.1"/><circle cx="13.6" cy="16.6" r="1.1"/></svg>',
  mic: '<svg viewBox="0 0 24 24"><rect x="9" y="3" width="6" height="11" rx="3"/><path d="M5.5 11a6.5 6.5 0 0013 0M12 17.5V21M8.5 21h7"/></svg>',
  wa: '<svg viewBox="0 0 24 24"><path d="M4 20l1.3-3.9A8 8 0 1112 20a8 8 0 01-4.2-1.2z"/><path d="M9 9.2c.2 2.2 2.4 4.6 5 5.4l1.2-1.2-1.7-.9-.7.7c-1-.5-1.9-1.4-2.3-2.4l.7-.7-.8-1.7z"/></svg>',
  play: '<svg viewBox="0 0 24 24"><path d="M8 5l11 7-11 7z" fill="currentColor"/></svg>',
  pause: '<svg viewBox="0 0 24 24"><path d="M7 5h3.5v14H7zM13.5 5H17v14h-3.5z" fill="currentColor" stroke="none"/></svg>',
  fwd: '<svg viewBox="0 0 24 24"><path d="M8 7a7 7 0 106 0"/><path d="M14 3v4h-4"/></svg>',
  bwd: '<svg viewBox="0 0 24 24"><path d="M16 7a7 7 0 11-6 0"/><path d="M10 3v4h4"/></svg>',
  search: '<svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="6.5"/><path d="M16 16l4.5 4.5"/></svg>',
  out: '<svg viewBox="0 0 24 24"><path d="M12 15V3M7.5 7.5L12 3l4.5 4.5"/><path d="M5 13v7h14v-7"/></svg>',
  cal: '<svg viewBox="0 0 24 24"><rect x="4" y="5" width="16" height="15" rx="2"/><path d="M4 10h16M9 3v4M15 3v4"/></svg>',
  copy: '<svg viewBox="0 0 24 24"><rect x="8" y="8" width="12" height="12" rx="2"/><path d="M16 8V5a1 1 0 00-1-1H5a1 1 0 00-1 1v10a1 1 0 001 1h3"/></svg>',
  trash: '<svg viewBox="0 0 24 24"><path d="M5 7h14M10 7V4h4v3M7 7l1 13h8l1-13"/></svg>',
  stop: '<svg viewBox="0 0 24 24"><rect x="6.5" y="6.5" width="11" height="11" rx="2" fill="currentColor"/></svg>',
  lock: '<svg viewBox="0 0 24 24" class="lk"><rect x="5" y="11" width="14" height="10" rx="2"/><path d="M8 11V8a4 4 0 018 0v3"/></svg>',
  spark: '<svg viewBox="0 0 24 24"><path d="M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8z"/><path d="M19 16l.8 2.2L22 19l-2.2.8L19 22l-.8-2.2L16 19l2.2-.8z"/></svg>',
};

/* ---------------- storage: IndexedDB (the transcripts and their audio stay on the device) ---------------- */
const DB = (() => {
  let p;
  const open = () => p || (p = new Promise((res, rej) => {
    const r = indexedDB.open('siman-voice', 1);
    r.onupgradeneeded = () => { const d = r.result; d.createObjectStore('items', { keyPath: 'id' }); d.createObjectStore('audio'); };
    r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
  }));
  const tx = async (store, mode, fn) => { const d = await open(); return new Promise((res, rej) => { const t = d.transaction(store, mode); const s = t.objectStore(store); const q = fn(s); t.oncomplete = () => res(q && q.result); t.onerror = () => rej(t.error); }); };
  return {
    put: it => tx('items', 'readwrite', s => s.put(it)),
    get: id => tx('items', 'readonly', s => s.get(id)),
    all: () => tx('items', 'readonly', s => s.getAll()),
    del: async id => { await tx('items', 'readwrite', s => s.delete(id)); await tx('audio', 'readwrite', s => s.delete(id)); },
    putAudio: (id, blob) => tx('audio', 'readwrite', s => s.put(blob, id)),
    getAudio: id => tx('audio', 'readonly', s => s.get(id)),
  };
})();

/* ---------------- smart features (text only, only when asked) ---------------- */
async function call(op, body, tries = 2) {
  let last;
  for (let k = 0; k < tries; k++) {
    try {
      const r = await fetch(VOICE_URL + '?op=' + op, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body) });
      const t = (await r.text()).trim();
      let d; try { d = JSON.parse(t); } catch (_) { throw Object.assign(new Error('תשובה לא תקינה מהשרת'), { status: r.status }); }
      if (!r.ok || d.error) throw Object.assign(new Error((d.error && d.error.message) || 'שגיאת שרת'), { status: (d.error && d.error.status) || r.status });
      return d;
    } catch (e) {
      last = e;
      if (k + 1 < tries && (!e.status || [429, 500, 502, 503, 504].includes(e.status))) { await sleep(e.status === 429 ? 6000 : 2500); continue; }
    }
  }
  throw last;
}
// the first time a smart feature is used, the person is told exactly what leaves the device
function cloudOk(what) {
  if (!SMART_CLOUD) return false;
  if (rd(CLOUD_KEY, false)) return true;
  const ok = confirm(`${what}\n\nכדי לעשות את זה, הטקסט של התמליל (בלי השמע) נשלח לעיבוד ב-Google Gemini.\nהשמע עצמו לא יוצא מהמכשיר, והתמלול נעשה במכשיר בלבד.\n\nלהמשיך?`);
  if (ok) wr(CLOUD_KEY, true);
  return ok;
}
const friendly = e => {
  const m = String(e && e.message || e);
  if (e && e.name === 'AbortError') return 'התמלול בוטל';
  if (e && e.status === 429 || /quota|exhausted|rate/i.test(m)) return 'השרת עמוס כרגע. נסו שוב בעוד כמה דקות';
  if (/could not decode|decode/i.test(m)) return 'לא הצלחתי לקרוא את השמע בקובץ. נסו קובץ MP3, M4A או WAV';
  if (/memory|allocation|Array buffer/i.test(m)) return 'אין למכשיר מספיק זיכרון להקלטה הזו. נסו במחשב, או קובץ קצר יותר';
  if (/fetch|network|Failed to load|NetworkError|ERR_/i.test(m)) return 'ההורדה הראשונה של מנוע התמלול צריכה אינטרנט. התחברו ונסו שוב (אחר כך זה עובד גם בלי)';
  if (/Worker|module/i.test(m)) return 'הדפדפן הזה לא מריץ את מנוע התמלול. נסו Chrome, Edge או Safari מעודכן';
  return 'התמלול נכשל: ' + m;
};
const abortErr = () => Object.assign(new Error('בוטל'), { name: 'AbortError' });

/* ---------------- audio: decoding to 16 kHz, and cutting at pauses ---------------- */
function resample(x, from, to) {
  if (from === to) return x;
  const ratio = from / to, n = Math.floor(x.length / ratio), y = new Float32Array(n);
  for (let i = 0; i < n; i++) {
    const a = Math.floor(i * ratio), b = Math.max(a + 1, Math.floor((i + 1) * ratio));
    let s = 0; for (let j = a; j < b; j++) s += x[j]; y[i] = s / (b - a);
  }
  return y;
}
async function decode16k(blob) {
  if (blob.size > 700 * 1024 * 1024) throw new Error('הקובץ גדול מדי לעיבוד במכשיר');
  const buf = await blob.arrayBuffer();
  const AC = window.OfflineAudioContext || window.webkitOfflineAudioContext;
  if (!AC) throw new Error('could not decode');
  let ctx = null;
  for (const r of [SR, 22050, 44100, 48000]) { try { ctx = new AC(1, r, r); break; } catch (_) {} }
  if (!ctx) throw new Error('could not decode');
  const audio = await new Promise((res, rej) => { try { const p = ctx.decodeAudioData(buf, res, rej); if (p && p.then) p.then(res, rej); } catch (e) { rej(e); } })
    .catch(() => { throw new Error('could not decode'); });
  const n = audio.length, nc = audio.numberOfChannels;
  let mono = nc === 1 ? audio.getChannelData(0).slice() : new Float32Array(n);
  if (nc > 1) { for (let c = 0; c < nc; c++) { const ch = audio.getChannelData(c); for (let i = 0; i < n; i++) mono[i] += ch[i]; } for (let i = 0; i < n; i++) mono[i] /= nc; }
  mono = resample(mono, audio.sampleRate, SR);
  // a quiet phone recording is brought up to a good level (peak ~90%)
  let peak = 0; for (let i = 0; i < mono.length; i += 7) { const v = Math.abs(mono[i]); if (v > peak) peak = v; }
  if (peak > 0.01 && peak < 0.9) { const g = Math.min(8, 0.9 / peak); for (let i = 0; i < mono.length; i++) mono[i] *= g; }
  return mono;
}
const F = 320, FPS = SR / F;            // 20 ms frames
function energies(pcm) { const nf = Math.floor(pcm.length / F), E = new Float32Array(nf); for (let f = 0; f < nf; f++) { let s = 0; const o = f * F; for (let i = 0; i < F; i++) { const v = pcm[o + i]; s += v * v; } E[f] = Math.sqrt(s / F); } return E; }
function threshold(sample) { const s = Array.from(sample).sort((a, b) => a - b); const floor = s[Math.floor(s.length * 0.1)] || 0; return Math.max(floor * 2.5, 0.006); }
// where to end the line that starts at frame s: the last pause (≥0.3 s, or ≥0.16 s) between 12 and 27 seconds
function findCut(E, s, end, thr, final) {
  const len = end - s;
  if (final && len <= MAX_CHUNK * FPS) return end;
  if (len < MIN_CUT * FPS + (final ? 0 : 10)) return 0;
  const lo = s + MIN_CUT * FPS, hi = Math.min(end - (final ? 0 : 10), s + MAX_CHUNK * FPS);
  for (const need of [15, 8]) {
    let run = 0, best = 0;
    for (let i = lo; i < hi; i++) { if (E[i] < thr) run++; else { if (run >= need) best = i - (run >> 1); run = 0; } }
    if (run >= need) best = hi - (run >> 1);
    if (best) return best;
  }
  if (len < MAX_CHUNK * FPS) return 0;
  let m = Infinity, at = s + MAX_CHUNK * FPS;
  for (let i = s + 18 * FPS; i < s + MAX_CHUNK * FPS; i++) if (E[i] < m) { m = E[i]; at = i; }
  return at;
}
// the speech inside [s,e): its start and end frame, or null when it is only silence (Whisper invents words in silence)
function speech(E, s, e, thr) {
  let a = s; while (a < e && E[a] < thr) a++;
  let b = e; while (b > a && E[b - 1] < thr) b--;
  let v = 0; for (let i = a; i < b; i++) if (E[i] > thr * 1.3) v++;
  return b - a >= 0.4 * FPS && v >= 0.25 * FPS ? [Math.max(s, a - 5), Math.min(e, b + 8)] : null;
}
function plan(pcm) {
  const E = energies(pcm), nf = E.length; if (!nf) return [];
  const thr = threshold(E.filter((_, i) => i % 5 === 0));
  const out = []; let s = 0;
  while (s < nf) {
    let e = findCut(E, s, nf, thr, true); if (e <= s) e = nf;
    const sp = speech(E, s, e, thr);
    if (sp) out.push({ s: sp[0] * F, e: sp[1] * F });
    s = e;
  }
  return out;
}
// what the engine wrote: loops of the same words (a known Whisper habit) are cut down
function clean(t, kind) {
  t = String(t || '').replace(/\s+/g, ' ').trim();
  for (let k = 0; k < 3; k++) t = t.replace(/(\S+(?:\s+\S+){0,6}?)(?:[\s,.]+\1){3,}/g, '$1');
  if (kind === 'dictation') t = t.replace(/(^|\s)(אה+|אממ+|אמ+|אהה+)(?=[\s,.]|$)/g, '$1').replace(/\s{2,}/g, ' ').trim();
  return /^[\s.,!?…-]*$/.test(t) ? '' : t;
}

/* ---------------- the engine (siman-asr-worker.js): runs in the browser, on this device ---------------- */
const ASR = (() => {
  let w = null, ready = null, cur = '', seq = 0, onProg = null, device = '';
  const wait = new Map();
  function worker() {
    if (w) return w;
    w = new Worker(new URL('siman-asr-worker.js', location.href), { type: 'module' });
    w.onmessage = e => {
      const d = e.data;
      if (d.type === 'progress') { if (onProg) onProg(d); return; }
      const p = wait.get(d.id); if (!p) return; wait.delete(d.id);
      if (d.error) p.rej(new Error(d.error)); else p.res(d);
    };
    w.onerror = e => { const err = new Error((e && e.message) || 'Worker failed'); for (const p of wait.values()) p.rej(err); wait.clear(); w = null; ready = null; cur = ''; };
    return w;
  }
  const ask = (msg, tr) => new Promise((res, rej) => { const id = ++seq; wait.set(id, { res, rej }); worker().postMessage(Object.assign({ id }, msg), tr || []); });
  const key = lang => (lang || 'he') === 'he' ? 'he' : 'multi';
  return {
    key,
    get device() { return device; },
    load(lang, progress) {
      const k = key(lang); onProg = progress || null;
      if (ready && cur === k) return ready;
      cur = k;
      ready = ask({ type: 'load', engine: k }).then(d => { device = d.device; return d; }, e => { ready = null; cur = ''; throw e; });
      return ready;
    },
    text: (pcm, lang) => ask({ type: 'run', audio: pcm, lang: lang || 'he' }, [pcm.buffer]),
    stop() { if (!w) return; w.terminate(); w = null; ready = null; cur = ''; for (const p of wait.values()) p.rej(abortErr()); wait.clear(); },
  };
})();
async function engineCached(lang) { try { const c = await caches.open('transformers-cache'); return !!(await c.match(ENGINES[ASR.key(lang)].probe)); } catch (_) { return false; } }
const mb = b => (b / 1048576).toFixed(b > 104857600 ? 0 : 1) + 'MB';
function progressSum(files) { let l = 0, t = 0; for (const k in files) { const f = files[k]; if (f.total) { t += f.total; l += f.done ? f.total : (f.loaded || 0); } } return { l, t }; }

/* ---------------- state ---------------- */
const JOB = { busy: false, queue: [], ctl: null };
let V = null;              // the open transcript
let audioEl = null, audioUrl = '';
let saveT = 0;
const prefs = Object.assign({ kind: 'lesson', lang: 'he' }, rd(PREF_KEY, {}));
const save = (now) => { if (!V) return; clearTimeout(saveT); const go = () => { V.updated = Date.now(); DB.put(stripForDb(V)).catch(e => console.error(e)); }; if (now) go(); else saveT = setTimeout(go, 700); };
const stripForDb = v => { const o = Object.assign({}, v); delete o.blob; return o; };

/* ---------------- styles ---------------- */
const CSS = `
#vx{position:fixed;inset:0;z-index:940;display:flex;flex-direction:column;background:var(--desk,#E2E7EE);color:var(--ink,#16213A);direction:rtl;font:15px/1.55 "SF-rubik","Rubik","Segoe UI",Arial,sans-serif;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}
#vx *{box-sizing:border-box}
#vx button{font:inherit;color:inherit;cursor:pointer}
#vx svg{width:20px;height:20px;fill:none;stroke:currentColor;stroke-width:1.8;stroke-linecap:round;stroke-linejoin:round;flex:none}
#vx [hidden]{display:none!important}
#vx .vx-hero .swipe{display:block;width:min(300px,70%);height:20px;margin:-6px 0 10px}
#vx :focus-visible{outline:3px solid var(--marker,#FFD43B);outline-offset:2px}
.vx-top{display:flex;align-items:center;gap:10px;padding:8px 14px;background:var(--navy,#16213A);color:#fff;min-height:54px}
.vx-back{display:flex;align-items:center;gap:4px;background:none;border:0;padding:8px 6px;border-radius:8px;color:#DCE3EF!important}
.vx-back:hover{background:rgba(255,255,255,.08)}
.vx-brand{font-size:19px;font-weight:700;background:linear-gradient(transparent 60%,var(--marker,#FFD43B) 60%,var(--marker,#FFD43B) 90%,transparent 90%);color:#fff;padding:0 3px}
.vx-brand i{font-style:normal;color:#16213A}
.vx-nav{display:flex;gap:4px;margin-inline-start:auto}
.vx-nav button{background:none;border:1px solid transparent;color:#DCE3EF!important;border-radius:999px;padding:6px 14px}
.vx-nav button[aria-current=true]{border-color:rgba(255,255,255,.35);color:#fff!important;background:rgba(255,255,255,.08)}
.vx-nav small{opacity:.7;margin-inline-start:4px}
.vx-main{flex:1;overflow:auto;position:relative}
.vx-screen{display:none}.vx-screen.on{display:block}
.vx-wrap{max-width:980px;margin:0 auto;padding:28px 20px 60px}
/* start */
.vx-hero h1{font-size:clamp(28px,5vw,44px);line-height:1.1;margin:6px 0 4px;letter-spacing:-.01em}
.vx-hero .swipe{display:block;width:min(300px,70%);height:20px;margin:-6px 0 10px}
.vx-hero p{color:var(--muted,#5D6982);max-width:60ch;margin:0 0 22px;font-size:16px}
.vx-tiles{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:22px}
.vx-tile{display:flex;flex-direction:column;align-items:flex-start;gap:6px;text-align:start;background:var(--bar,#fff);border:1px solid var(--line,#D4DBE5);border-radius:14px;padding:16px;min-height:118px;transition:border-color .15s,transform .15s}
.vx-tile:hover{border-color:var(--ink,#16213A);transform:translateY(-1px)}
.vx-tile svg{width:28px;height:28px}
.vx-tile div{display:flex;flex-direction:column;gap:3px}
.vx-tile b{font-size:16.5px}
.vx-tile span{color:var(--muted,#5D6982);font-size:13.5px}
.vx-tile.main{background:var(--navy,#16213A);color:#fff!important;border-color:var(--navy,#16213A)}
.vx-tile.main span{color:#C7D0E0}
.vx-tile.main svg{color:var(--marker,#FFD43B)}
.vx-set{background:var(--bar,#fff);border:1px solid var(--line,#D4DBE5);border-radius:14px;padding:16px;display:grid;gap:14px}
.vx-row{display:flex;flex-wrap:wrap;align-items:center;gap:8px 14px}
.vx-row>label,.vx-lbl{font-weight:600;min-width:8.5em}
.vx-seg{display:flex;flex-wrap:wrap;gap:6px}
.vx-seg button{border:1px solid var(--line,#D4DBE5);background:none;border-radius:999px;padding:6px 13px;font-size:14px}
.vx-seg button[aria-pressed=true]{background:var(--ink,#16213A);color:var(--bar,#fff)!important;border-color:var(--ink,#16213A)}
.vx-set select,.vx-set input[type=text]{font:inherit;color:inherit;background:var(--desk,#E2E7EE);border:1px solid var(--line,#D4DBE5);border-radius:9px;padding:7px 10px;min-width:12em;flex:1;max-width:520px}
.vx-hint{color:var(--muted,#5D6982);font-size:13px}
.vx-ok{display:flex;gap:9px;align-items:flex-start;font-size:13.5px;color:var(--muted,#5D6982)}
.vx-ok input{margin-top:4px;width:17px;height:17px;flex:none}
.vx-recent h2,.vx-lib h2{font-size:17px;margin:28px 0 10px}
.vx-items{display:grid;gap:8px}
.vx-item{display:flex;align-items:center;gap:12px;background:var(--bar,#fff);border:1px solid var(--line,#D4DBE5);border-radius:12px;padding:11px 14px;text-align:start;width:100%}
.vx-item:hover{border-color:var(--ink,#16213A)}
.vx-item .k{width:34px;height:34px;border-radius:50%;display:grid;place-items:center;background:var(--desk,#E2E7EE);flex:none}
.vx-item .t{flex:1;min-width:0}
.vx-item .t b{display:block;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.vx-item .t span{color:var(--muted,#5D6982);font-size:13px}
.vx-item .x{background:none;border:0;padding:6px;border-radius:8px;color:var(--muted,#5D6982)!important}
.vx-item .x:hover{color:#C2362B!important;background:var(--desk,#E2E7EE)}
.vx-empty{color:var(--muted,#5D6982);padding:18px 0}
/* work */
.vx-work{max-width:560px;margin:8vh auto 0;padding:0 20px;text-align:center}
.vx-work h2{font-size:22px;margin:0 0 4px}
.vx-work .f{color:var(--muted,#5D6982);direction:ltr;unicode-bidi:plaintext;word-break:break-all}
.vx-steps{list-style:none;padding:0;margin:26px 0;text-align:start;display:grid;gap:10px}
.vx-steps li{display:flex;align-items:center;gap:12px;padding:12px 14px;background:var(--bar,#fff);border:1px solid var(--line,#D4DBE5);border-radius:12px;color:var(--muted,#5D6982)}
.vx-steps li i{width:22px;height:22px;border-radius:50%;border:2px solid var(--line,#D4DBE5);flex:none;display:grid;place-items:center;font-style:normal;font-size:12px}
.vx-steps li.now{color:var(--ink,#16213A);border-color:var(--ink,#16213A)}
.vx-steps li.now i{border-color:var(--ink,#16213A);border-top-color:transparent;animation:vxspin .9s linear infinite}
.vx-steps li.done{color:var(--ink,#16213A)}
.vx-steps li.done i{background:var(--marker,#FFD43B);border-color:var(--marker,#FFD43B);color:#16213A}
.vx-steps li.done i::after{content:"✓"}
.vx-steps li small{margin-inline-start:auto;font-variant-numeric:tabular-nums}
@keyframes vxspin{to{transform:rotate(360deg)}}
.vx-bar{height:6px;background:var(--line,#D4DBE5);border-radius:6px;overflow:hidden;margin-top:6px;flex-basis:100%}
.vx-bar b{display:block;height:100%;width:0;background:var(--marker,#FFD43B);transition:width .2s}
.vx-btn{display:inline-flex;align-items:center;gap:7px;border:1px solid var(--line,#D4DBE5);background:var(--bar,#fff);border-radius:999px;padding:8px 16px;font-size:14.5px}
.vx-btn:hover{border-color:var(--ink,#16213A)}
.vx-btn.pri{background:var(--ink,#16213A);color:var(--bar,#fff)!important;border-color:var(--ink,#16213A)}
.vx-btn.sm{padding:5px 11px;font-size:13.5px}
.vx-btn:disabled{opacity:.5;cursor:default}
/* record */
.vx-rec{max-width:620px;margin:6vh auto 0;padding:0 20px;text-align:center}
.vx-rec .tm{font-size:clamp(48px,12vw,84px);font-weight:700;font-variant-numeric:tabular-nums;letter-spacing:-.02em;direction:ltr}
.vx-rec .st{color:var(--muted,#5D6982);min-height:1.6em}
.vx-rec canvas{width:100%;height:90px;display:block;margin:16px 0}
.vx-rec .ctl{display:flex;justify-content:center;gap:14px;align-items:center}
.vx-rec .big{width:76px;height:76px;border-radius:50%;border:0;background:#C2362B;color:#fff!important;display:grid;place-items:center;box-shadow:0 0 0 6px rgba(194,54,43,.18)}
.vx-rec .big svg{width:30px;height:30px}
.vx-rec .big.go{background:var(--navy,#16213A);box-shadow:0 0 0 6px rgba(22,33,58,.15)}
.vx-live{margin-top:22px;text-align:start;background:var(--bar,#fff);border:1px solid var(--line,#D4DBE5);border-radius:12px;padding:12px 14px;min-height:80px;max-height:30vh;overflow:auto;font-size:16px}
.vx-live:empty::before{content:attr(data-empty);color:var(--muted,#5D6982);font-size:14px}
.vx-live .int{color:var(--muted,#5D6982)}
/* document */
.vx-doc{display:grid;grid-template-columns:minmax(0,1fr) 360px;gap:0;min-height:100%}
.vx-tr{padding:20px 26px 140px;min-width:0}
.vx-side{border-inline-start:1px solid var(--line,#D4DBE5);background:var(--bar,#fff);padding:18px 18px 140px;overflow:auto;max-height:calc(100vh - 54px);min-height:calc(100vh - 54px);position:sticky;top:0}
.vx-head{display:flex;flex-wrap:wrap;align-items:flex-start;gap:10px;margin-bottom:12px}
.vx-title{font-size:clamp(21px,3vw,28px);font-weight:700;line-height:1.2;margin:0;flex:1;min-width:12em;outline:none;border-radius:6px;padding:2px 4px;margin-inline-start:-4px}
.vx-title:focus{background:var(--bar,#fff)}
.vx-meta{color:var(--muted,#5D6982);font-size:13.5px;flex-basis:100%;display:flex;flex-wrap:wrap;gap:4px 14px}
.vx-tools{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin:8px 0 14px;position:sticky;top:0;z-index:3;background:var(--desk,#E2E7EE);padding:8px 0}
.vx-find{display:flex;align-items:center;gap:6px;background:var(--bar,#fff);border:1px solid var(--line,#D4DBE5);border-radius:999px;padding:3px 12px;flex:1;min-width:150px;max-width:300px}
.vx-find input{border:0;background:none;font:inherit;color:inherit;outline:none;width:100%;padding:4px 0}
.vx-find small{color:var(--muted,#5D6982);white-space:nowrap}
.vx-spk{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:14px}
.vx-chip{display:inline-flex;align-items:center;gap:7px;border:1px solid var(--line,#D4DBE5);background:var(--bar,#fff);border-radius:999px;padding:4px 12px 4px 10px;font-size:14px}
.vx-chip i{width:11px;height:11px;border-radius:50%;flex:none}
.vx-chip small{color:var(--muted,#5D6982)}
.vx-lines{display:grid;gap:2px}
.vx-ln{display:grid;grid-template-columns:auto 1fr;gap:4px 12px;padding:9px 10px;border-radius:10px;position:relative}
.vx-ln:hover{background:rgba(127,127,127,.06)}
.vx-ln .who{display:flex;align-items:center;gap:6px;grid-column:1/-1;font-size:13.5px}
.vx-ln .who button{background:none;border:0;padding:0;font-weight:700}
.vx-ln .who .tm{color:var(--muted,#5D6982);font-variant-numeric:tabular-nums;direction:ltr;font-weight:400}
.vx-ln .who .tm:hover{text-decoration:underline}
.vx-ln .vx-sb{grid-row:2;width:3px;border-radius:3px;align-self:stretch}
.vx-ln .vx-tx{grid-row:2;font-size:16.5px;line-height:1.75;outline:none;border-radius:6px;white-space:pre-wrap;overflow-wrap:anywhere;unicode-bidi:plaintext;text-align:start}
.vx-ln .vx-tx:focus{background:var(--bar,#fff);box-shadow:0 0 0 1px var(--line,#D4DBE5)}
.vx-ln .trl{grid-row:3;grid-column:2;color:var(--muted,#5D6982);font-size:15px;unicode-bidi:plaintext;text-align:start}
.vx-ln.on .vx-tx{background:linear-gradient(transparent 58%,rgba(255,212,59,.55) 58%,rgba(255,212,59,.55) 92%,transparent 92%)}
.vx-ln .vx-tx .w.on{background:var(--marker,#FFD43B);color:#16213A;border-radius:3px}
.vx-ln .vx-tx mark{background:var(--marker,#FFD43B);color:#16213A;border-radius:3px}
.vx-ln .vx-tx mark.cur{outline:2px solid var(--ink,#16213A)}
.vx-ln.uncertain .vx-tx{}
.vx-u{border-bottom:2px dotted #C2410C}
/* side */
.vx-tabs{display:flex;gap:4px;margin-bottom:14px;border-bottom:1px solid var(--line,#D4DBE5)}
.vx-tabs button{background:none;border:0;border-bottom:3px solid transparent;padding:8px 10px;font-weight:600;color:var(--muted,#5D6982)!important}
.vx-tabs button[aria-selected=true]{color:var(--ink,#16213A)!important;border-bottom-color:var(--marker,#FFD43B)}
.vx-pane{display:none}.vx-pane.on{display:block}
.vx-card{margin-bottom:18px}
.vx-card h3{font-size:14.5px;margin:0 0 6px;display:flex;align-items:center;gap:6px}
.vx-card p{margin:0}
.vx-card ul{margin:0;padding:0 18px 0 0}
.vx-card li{margin:3px 0}
.vx-task{display:flex;gap:9px;align-items:flex-start;padding:7px 0;border-bottom:1px solid var(--line,#D4DBE5)}
.vx-task:last-child{border-bottom:0}
.vx-task input{margin-top:5px;width:16px;height:16px;flex:none}
.vx-task .d{flex:1}
.vx-task .d small{display:block;color:var(--muted,#5D6982)}
.vx-task.done .d>span{text-decoration:line-through;color:var(--muted,#5D6982)}
.vx-ic{background:none;border:0;padding:4px;border-radius:7px;color:var(--muted,#5D6982)!important}
.vx-ic:hover{background:var(--desk,#E2E7EE);color:var(--ink,#16213A)!important}
.vx-msg{background:var(--desk,#E2E7EE);border-radius:12px;padding:10px 12px;white-space:pre-wrap;unicode-bidi:plaintext;font-size:14.5px}
.vx-acts{display:flex;flex-wrap:wrap;gap:6px;margin-top:8px}
.vx-wait{display:flex;gap:10px;align-items:center;color:var(--muted,#5D6982);padding:14px 0}
.vx-wait i{width:18px;height:18px;border-radius:50%;border:2px solid var(--line,#D4DBE5);border-top-color:var(--ink,#16213A);animation:vxspin .9s linear infinite}
.vx-chat{display:grid;gap:10px;margin-bottom:12px}
.vx-q{justify-self:start;background:var(--ink,#16213A);color:var(--bar,#fff);border-radius:14px 14px 4px 14px;padding:8px 12px;max-width:90%;unicode-bidi:plaintext}
.vx-a{background:var(--desk,#E2E7EE);border-radius:14px 14px 14px 4px;padding:9px 12px;white-space:pre-wrap;unicode-bidi:plaintext}
.vx-a a{color:var(--blue,#2E58CC);cursor:pointer;direction:ltr;unicode-bidi:isolate}
.vx-ask{display:flex;gap:6px}
.vx-ask textarea{flex:1;font:inherit;color:inherit;background:var(--desk,#E2E7EE);border:1px solid var(--line,#D4DBE5);border-radius:12px;padding:8px 10px;resize:none;min-height:44px;max-height:140px}
.vx-sugg{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:10px}
/* player */
.vx-player{position:fixed;inset-inline:0;bottom:0;z-index:5;background:var(--navy,#16213A);color:#fff;padding:8px 16px calc(8px + env(safe-area-inset-bottom,0px));padding-inline-end:84px;display:grid;grid-template-columns:auto 1fr auto;gap:6px 14px;align-items:center}
.vx-player button{background:none;border:0;color:#fff!important;border-radius:999px;padding:6px;display:grid;place-items:center}
.vx-player button:hover{background:rgba(255,255,255,.1)}
.vx-player .pp{width:46px;height:46px;background:var(--marker,#FFD43B)!important;color:#16213A!important}
.vx-player .pp svg{width:22px;height:22px}
.vx-pl-ctl{display:flex;align-items:center;gap:2px}
.vx-pl-line{min-width:0}
.vx-pl-line canvas{width:100%;height:30px;display:block;cursor:pointer;touch-action:none}
.vx-pl-time{display:flex;justify-content:space-between;font-size:12px;color:#C7D0E0;font-variant-numeric:tabular-nums;direction:ltr}
.vx-pl-end{display:flex;align-items:center;gap:4px}
.vx-pl-end .spd{font-size:13px;font-weight:700;min-width:44px;border:1px solid rgba(255,255,255,.3)!important}
.vx-pl-end .fol[aria-pressed=false]{opacity:.5}
.vx-noaudio{position:fixed;inset-inline:0;bottom:0;padding:10px 16px;background:var(--bar,#fff);border-top:1px solid var(--line,#D4DBE5);color:var(--muted,#5D6982);font-size:13.5px;text-align:center}
/* menus, sheets */
.vx-menu{position:fixed;z-index:960;background:var(--bar,#fff);color:var(--ink,#16213A);border:1px solid var(--line,#D4DBE5);border-radius:12px;box-shadow:0 18px 40px -12px rgba(22,33,58,.45);padding:6px;min-width:220px;max-height:70vh;overflow:auto}
.vx-menu button{display:flex;align-items:center;gap:9px;width:100%;text-align:start;background:none;border:0;border-radius:8px;padding:9px 10px;font-size:14.5px}
.vx-menu button:hover{background:var(--desk,#E2E7EE)}
.vx-menu hr{border:0;border-top:1px solid var(--line,#D4DBE5);margin:4px 0}
.vx-menu input{font:inherit;color:inherit;width:100%;background:var(--desk,#E2E7EE);border:1px solid var(--line,#D4DBE5);border-radius:8px;padding:7px 9px;margin:4px 0 6px}
.vx-menu .cap{font-size:12.5px;color:var(--muted,#5D6982);padding:4px 10px 0}
dialog.vx-dlg{border:0;border-radius:16px;padding:0;max-width:520px;width:calc(100% - 32px);background:var(--bar,#fff);color:var(--ink,#16213A);direction:rtl}
dialog.vx-dlg::backdrop{background:rgba(13,20,36,.55)}
.vx-dlg .in{padding:22px}
.vx-dlg h2{margin:0 0 12px;font-size:20px}
.vx-dlg ol{padding:0 20px 0 0;margin:6px 0 14px}
.vx-dlg li{margin:6px 0}
.vx-dlg h3{font-size:15px;margin:14px 0 4px}
.vx-drop{position:fixed;inset:0;z-index:970;background:rgba(22,33,58,.7);color:#fff;display:none;place-items:center;font-size:24px;font-weight:700;pointer-events:none}
#vx.drag .vx-drop{display:grid}
.vx-mtabs{display:none}
.vx-priv{display:flex;gap:9px;align-items:flex-start;background:var(--bar,#fff);border:1px solid var(--line,#D4DBE5);border-inline-start:4px solid var(--marker,#FFD43B);border-radius:10px;padding:10px 14px;max-width:62ch;color:var(--ink,#16213A)!important;font-size:15px}
.vx-priv svg,.vx-hint svg.lk{width:18px;height:18px;margin-top:2px}
.vx-hint svg.lk{vertical-align:-3px;margin:0 0 0 3px}
.vx-good{color:#0F766E;font-weight:700}
.vx-wlive{text-align:start;margin:0 0 14px;max-height:34vh}
.vx-wlive p,.vx-live p{margin:0 0 6px;unicode-bidi:plaintext}
.vx-cloud{background:var(--desk,#E2E7EE);border-radius:12px;padding:12px 14px}
@media (max-width:860px){
  .vx-tiles{grid-template-columns:1fr}
  .vx-tile{min-height:0;flex-direction:row;align-items:center;gap:14px}
  .vx-tile div{display:flex;flex-direction:column}
  .vx-doc{display:block}
  .vx-tr{padding:12px 14px 150px}
  .vx-side{position:static;max-height:none;border:0;background:none;padding:12px 14px 150px;display:none}
  .vx-doc.side .vx-tr{display:none}.vx-doc.side .vx-side{display:block}
  .vx-mtabs{display:flex;gap:4px;position:sticky;top:0;z-index:4;background:var(--desk,#E2E7EE);padding:8px 14px 0;border-bottom:1px solid var(--line,#D4DBE5)}
  .vx-mtabs button{flex:1;background:none;border:0;border-bottom:3px solid transparent;padding:9px 4px;font-weight:600;color:var(--muted,#5D6982)!important}
  .vx-mtabs button[aria-selected=true]{color:var(--ink,#16213A)!important;border-bottom-color:var(--marker,#FFD43B)}
  .vx-tabs{display:none}
  .vx-tools{top:42px}
  .vx-row>label,.vx-lbl{min-width:0;flex-basis:100%}
  .vx-player{grid-template-columns:auto 1fr;padding-inline:10px}
  .vx-pl-end{grid-column:1/-1;justify-content:space-between}
  .vx-nav button{padding:6px 10px}
}
@media (prefers-reduced-motion:reduce){.vx-tile{transition:none}.vx-steps li.now i,.vx-wait i{animation:none}}
`;

/* ---------------- the shell ---------------- */
let root = null;
function ensure() {
  if (root) return root;
  const st = el('style'); st.id = 'vxCss'; st.textContent = CSS; document.head.append(st);
  root = el('div'); root.id = 'vx'; root.setAttribute('role', 'dialog'); root.setAttribute('aria-label', 'תמלול'); root.hidden = true;
  root.innerHTML = `
  <header class="vx-top">
    <button class="vx-back" data-a="close" title="חזרה לסימן">${ico.back}<span>סימן</span></button>
    <b class="vx-brand">תמלול</b>
    <nav class="vx-nav"><button data-v="new">חדש</button><button data-v="lib">הספרייה<small id="vxCnt"></small></button></nav>
  </header>
  <div class="vx-main" id="vxMain">
    <section class="vx-screen" id="vxNew"></section>
    <section class="vx-screen" id="vxWork"></section>
    <section class="vx-screen" id="vxRec"></section>
    <section class="vx-screen" id="vxDoc"></section>
    <section class="vx-screen" id="vxLib"></section>
  </div>
  <div class="vx-drop">משחררים כאן קובץ שמע או וידאו לתמלול</div>
  <input type="file" id="vxFile" accept="audio/*,video/*,.opus,.ogg,.oga,.m4a,.mp3,.wav,.aac,.flac,.amr,.3gp,.webm,.mp4,.mov,.caf,.wma" multiple hidden>`;
  document.body.append(root);
  root.addEventListener('click', e => {
    const b = e.target.closest('[data-a],[data-v]'); if (!b || !root.contains(b)) return;
    if (b.dataset.v) { show(b.dataset.v === 'lib' ? 'lib' : 'new'); return; }
    const a = b.dataset.a;
    if (a === 'close') close();
    else if (a === 'pick') pick();
    else if (a === 'rec') startRec();
    else if (a === 'wa') waHelp();
  });
  const fi = root.querySelector('#vxFile');
  fi.onchange = () => { const fs = [...fi.files]; fi.value = ''; fs.forEach(f => enqueue(f)); };
  // dropping a file on the transcription screen
  let dn = 0;
  root.addEventListener('dragenter', e => { if ([...(e.dataTransfer.types || [])].includes('Files')) { dn++; root.classList.add('drag'); } });
  root.addEventListener('dragleave', () => { if (--dn <= 0) { dn = 0; root.classList.remove('drag'); } });
  root.addEventListener('dragover', e => e.preventDefault());
  root.addEventListener('drop', e => {
    e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation(); dn = 0; root.classList.remove('drag');
    const fs = [...(e.dataTransfer.files || [])]; const media = fs.filter(isMedia);
    if (!media.length) { say('כאן מתמללים קובצי שמע ווידאו'); return; }
    media.forEach(f => enqueue(f));
  }, true);
  addEventListener('keydown', onKey);
  return root;
}
function show(name) {
  ensure();
  const map = { new: 'vxNew', work: 'vxWork', rec: 'vxRec', doc: 'vxDoc', lib: 'vxLib' };
  root.querySelectorAll('.vx-screen').forEach(s => s.classList.toggle('on', s.id === map[name]));
  root.querySelectorAll('.vx-nav button').forEach(b => b.setAttribute('aria-current', String((name === 'lib' && b.dataset.v === 'lib') || (name === 'new' && b.dataset.v === 'new'))));
  if (name !== 'doc') stopAudio();
  if (name === 'new') renderNew();
  if (name === 'lib') renderLib();
  root.querySelector('#vxMain').scrollTop = 0;
}
function open(screen) {
  if (typeof needReg === 'function' && needReg()) { try { showReg(); } catch (_) {} return false; }
  ensure(); root.hidden = false; document.body.style.overflow = 'hidden';
  show(screen || 'new'); refreshCount();
  return true;
}
function close() {
  if (REC.mr && REC.mr.state !== 'inactive') { if (!confirm('ההקלטה עוד פועלת. לעצור ולשמור אותה?')) return; stopRec(); return; }
  if (JOB.busy && !confirm('התמלול עוד בעבודה. הוא ימשיך ברקע וישמר בספרייה. לחזור לסימן?')) return;
  stopAudio(); closeMenu();
  if (root) root.hidden = true; document.body.style.overflow = '';
  if (V) save(true);
}
function onKey(e) {
  if (!root || root.hidden) return;
  const typing = e.target.isContentEditable || /INPUT|TEXTAREA|SELECT/.test(e.target.tagName);
  if (e.key === 'Escape') { if (MENU) { closeMenu(); return; } if (!typing) { close(); e.preventDefault(); } return; }
  if (!audioEl || typing) return;
  if (e.code === 'Space') { e.preventDefault(); togglePlay(); }
  else if (e.key === 'ArrowLeft') { e.preventDefault(); seek(audioEl.currentTime + 5); }
  else if (e.key === 'ArrowRight') { e.preventDefault(); seek(audioEl.currentTime - 5); }
  else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'f') { const f = root.querySelector('#vxFindIn'); if (f) { e.preventDefault(); f.focus(); } }
}
async function refreshCount() { try { const n = (await DB.all()).length; const c = root.querySelector('#vxCnt'); if (c) c.textContent = n ? ' ' + n : ''; } catch (_) {} }

/* ---------------- start screen ---------------- */
function renderNew() {
  const s = root.querySelector('#vxNew');
  const okd = !!rd(OK_KEY, false);
  if (!LANG_NAMES[prefs.lang]) prefs.lang = 'he';
  if (!KINDS[prefs.kind]) prefs.kind = 'lesson';
  s.innerHTML = `<div class="vx-wrap">
    <div class="vx-hero">
      <h1>מה מתמללים?</h1>
      <svg class="swipe" viewBox="0 0 300 26" preserveAspectRatio="none" aria-hidden="true"><path d="M6 15c46-9 98-11 150-9s92 7 138 2" style="stroke:var(--marker,#FFD43B)" stroke-width="16" stroke-linecap="round" fill="none"/></svg>
      <p>שיעור, הרצאה, שיחה חשובה או הודעה קולית, לטקסט שאפשר לערוך, לחפש בו ולשמור כקובץ Word. נגן שמסמן את השורה בזמן ההשמעה עוזר לבדוק ולתקן.</p>
      <p class="vx-priv">${ico.lock}<span><b>הכול נשאר אצלכם.</b> התמלול נעשה בתוך המכשיר הזה, והשמע לא נשלח לשום שרת.</span></p>
    </div>
    <div class="vx-tiles">
      <button class="vx-tile main" data-a="pick">${ico.file}<div><b>בחירת קובץ שמע או וידאו</b><span>MP3, M4A, WAV, הודעות OPUS של וואטסאפ, MP4 ועוד. אפשר גם לגרור לכאן</span></div></button>
      <button class="vx-tile" data-a="rec">${ico.mic}<div><b>הקלטה עכשיו</b><span>שיעור, הרצאה או ישיבה. הטקסט נכתב תוך כדי ההקלטה</span></div></button>
      <button class="vx-tile" data-a="wa">${ico.wa}<div><b>הודעה או שיחה מהטלפון</b><span>איך מעבירים הודעה קולית או שיחה מוקלטת לסימן</span></div></button>
    </div>
    <div class="vx-set">
      <div class="vx-row"><span class="vx-lbl">מה יש בהקלטה?</span><div class="vx-seg" id="vxKind">${Object.entries(KINDS).map(([k, v]) => `<button data-k="${k}" aria-pressed="${prefs.kind === k}">${v.label}</button>`).join('')}</div></div>
      <div class="vx-row"><label for="vxLang">שפת הדיבור</label><select id="vxLang">${Object.entries(LANG_NAMES).map(([k, v]) => `<option value="${k}"${prefs.lang === k ? ' selected' : ''}>${v}</option>`).join('')}</select></div>
      <div class="vx-row"><span class="vx-lbl">מנוע התמלול</span><span id="vxEng" class="vx-hint">בודק…</span></div>
      <label class="vx-ok"><input type="checkbox" id="vxOk"${okd ? ' checked' : ''}><span>אני משתתף/ת בשיחה או בשיעור שהוקלטו, או שקיבלתי הסכמה להקליט.</span></label>
    </div>
    <div class="vx-recent" id="vxRecent"></div>
  </div>`;
  s.querySelectorAll('#vxKind button').forEach(b => b.onclick = () => { prefs.kind = b.dataset.k; wr(PREF_KEY, prefs); s.querySelectorAll('#vxKind button').forEach(x => x.setAttribute('aria-pressed', String(x === b))); });
  s.querySelector('#vxLang').onchange = e => { prefs.lang = e.target.value; wr(PREF_KEY, prefs); engineRow(); };
  s.querySelector('#vxOk').onchange = e => wr(OK_KEY, e.target.checked);
  engineRow();
  DB.all().then(items => {
    const box = s.querySelector('#vxRecent'); if (!box) return;
    items.sort((a, b) => (b.updated || b.created) - (a.updated || a.created));
    if (!items.length) { box.innerHTML = ''; return; }
    box.innerHTML = `<h2>אחרונים</h2><div class="vx-items"></div>`;
    items.slice(0, 4).forEach(it => box.querySelector('.vx-items').append(itemRow(it)));
    if (items.length > 4) { const m = el('button', 'vx-btn sm', 'לכל הספרייה (' + items.length + ')'); m.style.marginTop = '10px'; m.onclick = () => show('lib'); box.append(m); }
  }).catch(() => {});
}
// is the engine already on this device? if not: what it costs, and a button to get it now (on Wi-Fi)
async function engineRow() {
  const box = root && root.querySelector('#vxEng'); if (!box) return;
  const lang = prefs.lang, E = ENGINES[ASR.key(lang)];
  const have = await engineCached(lang);
  if (have) { box.innerHTML = `<span class="vx-good">✓</span> מותקן במכשיר. עובד גם בלי אינטרנט`; return; }
  box.innerHTML = `יורד פעם אחת בשימוש הראשון (כ-${E.mb}MB, עדיף ב-Wi-Fi) <button class="vx-btn sm" id="vxGetEng">להוריד עכשיו</button>`;
  box.querySelector('#vxGetEng').onclick = async ev => {
    const b = ev.currentTarget; b.disabled = true; const files = {};
    try {
      await ASR.load(lang, p => { files[p.file] = p; const { l, t } = progressSum(files); if (t > 1e6) b.textContent = 'מוריד… ' + Math.round(l / t * 100) + '%'; });
      engineRow();
    } catch (e) { say(friendly(e)); b.disabled = false; b.textContent = 'לנסות שוב'; }
  };
}
function consentOk() {
  if (rd(OK_KEY, false)) return true;
  const c = root && root.querySelector('#vxOk');
  if (c && c.checked) { wr(OK_KEY, true); return true; }
  show('new');
  const box = root.querySelector('.vx-ok');
  if (box) { box.scrollIntoView({ block: 'center' }); box.style.color = 'var(--ink)'; box.querySelector('input').focus(); }
  say('סמנו קודם שאתם משתתפים בשיחה או שקיבלתם הסכמה להקליט');
  return false;
}
function pick() { if (!consentOk()) return; root.querySelector('#vxFile').click(); }
function waHelp() {
  const d = el('dialog', 'vx-dlg');
  d.innerHTML = `<div class="in"><h2>הודעה קולית או שיחה מוקלטת לסימן</h2>
    <h3>אנדרואיד</h3>
    <ol><li>מתקינים את סימן כאפליקציה (״התקנה״ בראש המסך או ״הוספה למסך הבית״).</li><li>בוואטסאפ: לחיצה ארוכה על ההודעה הקולית ← שיתוף <bdi>(⋮)</bdi>.</li><li>בוחרים <b>סימן</b>. התמלול מתחיל לבד.</li></ol>
    <p class="vx-hint">שיחה שהקלטתם בטלפון (בחייגן או באפליקציית הקלטה): משתפים את קובץ ההקלטה לסימן באותה דרך.</p>
    <h3>אייפון</h3>
    <ol><li>בוואטסאפ: לחיצה ארוכה על ההודעה ← העברה ← שיתוף ← <b>שמירה בקבצים</b>.</li><li>בסימן: ״בחירת קובץ שמע או וידאו״ ובוחרים אותה מ״קבצים״.</li></ol>
    <p class="vx-hint">שיחה שהוקלטה באייפון נשמרת באפליקציית הפתקים: פותחים את ההקלטה ← שיתוף ← שמירה בקבצים.</p>
    <h3>מחשב</h3>
    <ol><li>בוואטסאפ למחשב: לחיצה ימנית על ההודעה ← הורדה.</li><li>גוררים את הקובץ לחלון של סימן.</li></ol>
    <div class="vx-acts" style="justify-content:flex-end"><button class="vx-btn pri" value="ok">הבנתי</button></div></div>`;
  d.querySelector('button').onclick = () => d.close();
  d.addEventListener('close', () => d.remove());
  root.append(d); d.showModal();
}

/* ---------------- the job: engine → audio → transcription, all on this device ---------------- */
function enqueue(file, opts) {
  if (!isMedia(file)) { say('זה לא קובץ שמע או וידאו'); return; }
  if (typeof needReg === 'function' && needReg()) { try { showReg(); } catch (_) {} return; }
  if (!root || root.hidden) { if (!open(JOB.busy ? 'work' : 'new')) return; }
  JOB.queue.push({ file, opts: opts || {} });
  if (JOB.busy) { say('נוסף לתור: ' + file.name + ' (' + JOB.queue.length + ' בהמתנה)'); return; }
  run();
}
async function run() {
  while (JOB.queue.length) {
    const { file, opts } = JOB.queue.shift();
    JOB.busy = true; JOB.ctl = new AbortController();
    try { await transcribeFile(file, opts); }
    catch (e) { if (e.name !== 'AbortError') console.error(e); }
    finally { JOB.busy = false; JOB.ctl = null; }
  }
}
const STEPS = [['model', 'מכין את מנוע התמלול'], ['prep', 'קורא את השמע'], ['tx', 'מתמלל במכשיר']];
function workScreen(name) {
  show('work');
  const s = root.querySelector('#vxWork');
  s.innerHTML = `<div class="vx-work"><h2>מתמלל</h2><div class="f">${esc(name)}</div>
    <ul class="vx-steps">${STEPS.map(([k, t]) => `<li data-s="${k}"><i></i><span>${t}</span><small></small></li>`).join('')}</ul>
    <div class="vx-live vx-wlive" id="vxWorkLive" dir="auto" data-empty="הטקסט יופיע כאן כשהתמלול יתחיל"></div>
    <p class="vx-hint" id="vxWorkHint">${ico.lock} השמע לא יוצא מהמכשיר. אפשר להמשיך לעבוד בסימן בינתיים, והתמליל יחכה בספרייה.</p>
    <button class="vx-btn" id="vxCancel">ביטול</button></div>`;
  s.querySelector('#vxCancel').onclick = () => { if (JOB.ctl) JOB.ctl.abort(); ASR.stop(); };
}
function step(k, state, note) {
  const li = root && root.querySelector(`#vxWork li[data-s="${k}"]`); if (!li) return;
  li.className = state || '';
  if (note != null) li.querySelector('small').textContent = note;
}
function bar(k, frac) {
  const li = root.querySelector(`#vxWork li[data-s="${k}"]`); if (!li) return;
  let b = li.querySelector('.vx-bar'); if (!b) { li.style.flexWrap = 'wrap'; b = el('div', 'vx-bar', '<b></b>'); li.append(b); }
  b.firstChild.style.width = Math.round(Math.max(0, Math.min(1, frac)) * 100) + '%';
}
function liveLine(t) {
  const box = root && root.querySelector('#vxWorkLive'); if (!box) return;
  const p = el('p', '', esc(t)); box.append(p);
  while (box.children.length > 40) box.firstChild.remove();
  box.scrollTop = box.scrollHeight;
}
const eta = s => s < 60 ? 'פחות מדקה' : s < 3600 ? 'עוד כ-' + Math.round(s / 60) + ' דק׳' : 'עוד כ-' + (s / 3600).toFixed(1) + ' שע׳';
async function loadEngine(lang, k) {
  const files = {};
  const have = await engineCached(lang);
  step(k, 'now', have ? '' : 'הורדה חד-פעמית');
  await ASR.load(lang, p => {
    files[p.file] = p; const { l, t } = progressSum(files);
    if (t > 2e6 && !have) { bar(k, l / t); step(k, 'now', mb(l) + ' מתוך ' + mb(t)); }
  });
  step(k, 'done', ASR.device === 'webgpu' ? 'כרטיס מסך' : '');
  const b = root.querySelector(`#vxWork li[data-s="${k}"] .vx-bar`); if (b) b.remove();
}
async function transcribeFile(file, opts) {
  const kind = opts.kind || prefs.kind;
  const lang = opts.lang || prefs.lang || 'he';
  const id = opts.id || uid();
  workScreen(file.name || 'הקלטה');
  const item = {
    id, title: (file.name || 'הקלטה').replace(/\.[a-z0-9]+$/i, ''), fileName: file.name || '', created: opts.created || Date.now(), updated: Date.now(),
    kind, lang, duration: opts.duration || null, segments: [], speakers: {}, analysis: null, tr: {}, chat: [], tasksDone: {}, status: 'working', engine: 'local',
  };
  const sig = JOB.ctl.signal;
  try {
    // the audio is kept on the device first: nothing is lost if something fails
    await DB.putAudio(id, file); await DB.put(item);
    // 1. the engine (downloaded once, then from the device)
    await loadEngine(lang, 'model');
    if (sig.aborted) throw abortErr();
    // 2. the audio
    step('prep', 'now');
    const pcm = await decode16k(file);
    item.duration = pcm.length / SR;
    const parts = plan(pcm);
    step('prep', 'done', longDur(item.duration));
    if (!parts.length) throw new Error('לא נשמע דיבור בהקלטה');
    // 3. part by part; the text shows up as it is written
    step('tx', 'now', '0%');
    const total = parts.reduce((a, p) => a + p.e - p.s, 0) / SR;
    let done = 0; const t0 = Date.now();
    for (let k = 0; k < parts.length; k++) {
      if (sig.aborted) throw abortErr();
      const p = parts[k];
      const r = await ASR.text(pcm.slice(p.s, p.e), lang);
      const text = clean(r.text, kind);
      if (text) { item.segments.push({ spk: 'spk_1', start: p.s / SR, end: p.e / SR, text }); liveLine(text); }
      done += (p.e - p.s) / SR;
      const spent = (Date.now() - t0) / 1000;
      bar('tx', done / total);
      step('tx', 'now', Math.round(done / total * 100) + '%' + (k >= 1 && done < total ? ' · ' + eta(spent / done * (total - done)) : ''));
      if (k % 8 === 7) await DB.put(stripForDb(item));
    }
    item.segments = normalize(item.segments);
    if (!item.segments.length) throw new Error('לא נשמע דיבור בהקלטה');
    step('tx', 'done', clock((Date.now() - t0) / 1000));
    item.speakers = speakersOf(item.segments, item.speakers);
    item.status = 'ready';
    await DB.put(stripForDb(item));
    count('voice_transcribe'); refreshCount();
    const onWork = !root.hidden && root.querySelector('#vxWork.on');
    if (onWork && !JOB.queue.length) await openItem(id);
    else say('התמליל מוכן: ' + item.title + (root.hidden ? ' (בתמלול ← הספרייה)' : ''));
  } catch (e) {
    item.status = 'failed'; item.error = friendly(e);
    if (item.segments.length) item.segments = normalize(item.segments);
    try { await DB.put(stripForDb(item)); } catch (_) {}
    const s = root.querySelector('#vxWork');
    if (s && s.classList.contains('on')) {
      const now = s.querySelector('.vx-steps li.now'); if (now) now.className = '';
      const h = s.querySelector('#vxWorkHint'); h.textContent = item.error; h.style.color = '#C2362B';
      const c = s.querySelector('#vxCancel'); c.textContent = 'חזרה'; c.onclick = () => show('new');
      const again = el('button', 'vx-btn pri', 'לנסות שוב'); again.style.marginInlineEnd = '8px';
      again.onclick = () => retry(id); c.before(again);
    } else say(item.error);
    throw e;
  }
}
async function retry(id) {
  const it = await DB.get(id), blob = await DB.getAudio(id);
  if (!blob) { say('השמע של התמליל הזה לא נשמר'); return; }
  const f = blob instanceof File ? blob : new File([blob], (it && it.fileName) || 'הקלטה.webm', { type: blob.type });
  JOB.queue.push({ file: f, opts: { id, kind: it && it.kind, lang: it && it.lang, created: it && it.created } });
  if (!JOB.busy) run();
}
// labels in order of appearance: spk_1, spk_2…
function normalize(segs) {
  const map = {}; let n = 0;
  const out = [];
  for (const s of segs) {
    const text = String(s.text || '').replace(/\s+/g, ' ').trim(); if (!text) continue;
    const raw = String(s.spk || 'spk_1');
    if (!map[raw]) map[raw] = 'spk_' + (++n);
    out.push({ spk: map[raw], start: s.start == null ? null : +s.start, end: s.end == null ? null : +s.end, text });
  }
  for (let i = 0; i < out.length; i++) if (out[i].end == null && out[i].start != null && out[i + 1] && out[i + 1].start != null) out[i].end = out[i + 1].start;
  return out;
}
function speakersOf(segs, old) {
  const sp = {}; let i = 0;
  for (const s of segs) if (!sp[s.spk]) { const o = old && old[s.spk]; sp[s.spk] = { name: o && o.name || '', color: o && o.color || COLORS[i % COLORS.length] }; i++; }
  return sp;
}
const spkName = (it, id) => { const s = it.speakers[id]; if (s && s.name) return s.name; const n = String(id).replace(/\D/g, '') || '1'; return 'דובר ' + n; };
const plain = segs => segs.map(s => s.text).join('\n');
const forAi = it => it.segments.map(s => `${s.start != null ? '[' + clock2(s.start) + '] ' : ''}${it.speakers[s.spk] && it.speakers[s.spk].name ? it.speakers[s.spk].name + ' (' + s.spk + ')' : s.spk}: ${s.text}`).join('\n');
function applyAnalysis(it, a) {
  if (!a) return;
  it.analysis = a;
  if (a.title && !it.titleEdited) it.title = String(a.title).slice(0, 120);
  const named = [];
  for (const s of a.speakers || []) {
    const id = String(s.id || '').trim(); const nm = String(s.name || '').trim();
    if (it.speakers[id] && nm && !it.speakers[id].name && nm.length < 40 && !/^(spk|דובר|speaker)/i.test(nm)) { it.speakers[id].name = nm; it.speakers[id].auto = true; named.push(nm); }
    if (it.speakers[id] && s.role) it.speakers[id].role = String(s.role).slice(0, 40);
  }
  if (named.length) it.namedNote = 'זיהיתי שמות: ' + named.join(', ') + '. אפשר לשנות בלחיצה על השם';
}

/* ---------------- the transcript screen ---------------- */
async function openItem(id) {
  const it = await DB.get(id);
  if (!it) { say('התמליל לא נמצא'); return; }
  if (it.status === 'failed' || !it.segments || !it.segments.length) {
    if (!open('work')) return;
    workScreen(it.fileName || it.title);
    const s = root.querySelector('#vxWork'); s.querySelector('.vx-steps').remove();
    const h = s.querySelector('#vxWorkHint'); h.textContent = it.error || (it.status === 'working' ? 'התמליל עוד לא הסתיים' : 'התמלול לא הושלם'); h.style.color = '#C2362B';
    const c = s.querySelector('#vxCancel'); c.textContent = 'חזרה'; c.onclick = () => show('lib');
    const again = el('button', 'vx-btn pri', 'לתמלל שוב'); again.style.marginInlineEnd = '8px'; again.onclick = () => retry(id); c.before(again);
    return;
  }
  if (root && root.hidden) { if (!open('doc')) return; }
  else if (!root) { if (!open('doc')) return; }
  V = it; if (!V.tr) V.tr = {}; if (!V.chat) V.chat = []; if (!V.tasksDone) V.tasksDone = {};
  V.blob = await DB.getAudio(id).catch(() => null);
  renderDoc();
  show('doc');
  if (V.namedNote) { say(V.namedNote); delete V.namedNote; save(); }
}
let SHOW_TR = '', FIND = { q: '', hits: [], cur: -1 }, FOLLOW = true, activeIdx = -1;
function renderDoc() {
  const s = root.querySelector('#vxDoc');
  const rtl = isRtl(V.lang) || (!V.lang && guessRtl(plain(V.segments)));
  SHOW_TR = ''; FIND = { q: '', hits: [], cur: -1 }; activeIdx = -1;
  s.innerHTML = `
  <div class="vx-mtabs" role="tablist"><button data-m="tr" aria-selected="true">תמליל</button><button data-m="side" aria-selected="false">סיכום${SMART_CLOUD ? ' ושאלות' : ''}</button></div>
  <div class="vx-doc" id="vxDocGrid">
    <div class="vx-tr">
      <div class="vx-head">
        <h1 class="vx-title" id="vxTitle" contenteditable="${CE}" spellcheck="false">${esc(V.title)}</h1>
        <button class="vx-btn pri" id="vxExport">${ico.out}ייצוא ושיתוף</button>
        <div class="vx-meta"><span>${esc(dayStr(V.created))}</span>${V.duration ? `<span>${longDur(V.duration)}</span>` : ''}<span>${Object.keys(V.speakers).length} ${Object.keys(V.speakers).length === 1 ? 'דובר' : 'דוברים'}</span><span>${esc(KINDS[V.kind] ? KINDS[V.kind].label : '')}</span></div>
      </div>
      <div class="vx-tools">
        <label class="vx-find">${ico.search}<input id="vxFindIn" type="search" placeholder="חיפוש בתמליל" aria-label="חיפוש בתמליל"><small id="vxFindN"></small></label>
        <button class="vx-btn sm" id="vxFindPrev" hidden title="הקודם">▲</button><button class="vx-btn sm" id="vxFindNext" hidden title="הבא">▼</button>
        ${SMART_CLOUD ? '<button class="vx-btn sm" id="vxTrBtn">תרגום</button>' : ''}
      </div>
      <div class="vx-spk" id="vxSpk"></div>
      <div class="vx-lines" id="vxLines" dir="${rtl ? 'rtl' : 'ltr'}"></div>
    </div>
    <aside class="vx-side" id="vxSide">
      <div class="vx-tabs" role="tablist"><button data-t="ins" aria-selected="true">סיכום</button>${SMART_CLOUD ? '<button data-t="ask" aria-selected="false">שאלות על התמליל</button>' : ''}</div>
      <div class="vx-pane on" data-p="ins" id="vxIns"></div>
      <div class="vx-pane" data-p="ask" id="vxAsk"></div>
    </aside>
  </div>`;
  const grid = s.querySelector('#vxDocGrid');
  s.querySelectorAll('.vx-mtabs button').forEach(b => b.onclick = () => {
    s.querySelectorAll('.vx-mtabs button').forEach(x => x.setAttribute('aria-selected', String(x === b)));
    grid.classList.toggle('side', b.dataset.m === 'side'); root.querySelector('#vxMain').scrollTop = 0;
  });
  s.querySelectorAll('.vx-tabs button').forEach(b => b.onclick = () => sideTab(b.dataset.t));
  const ti = s.querySelector('#vxTitle');
  ti.oninput = () => { V.title = ti.textContent.trim() || 'תמליל'; V.titleEdited = true; save(); };
  ti.onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); ti.blur(); } };
  s.querySelector('#vxExport').onclick = e => exportMenu(e.currentTarget);
  const trb = s.querySelector('#vxTrBtn'); if (trb) trb.onclick = e => trMenu(e.currentTarget);
  const fi = s.querySelector('#vxFindIn');
  fi.oninput = () => find(fi.value);
  fi.onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); findStep(e.shiftKey ? -1 : 1); } };
  s.querySelector('#vxFindNext').onclick = () => findStep(1);
  s.querySelector('#vxFindPrev').onclick = () => findStep(-1);
  renderSpeakers(); renderLines(); renderIns(); renderAsk();
  setupPlayer(s);
}
function sideTab(t) {
  root.querySelectorAll('.vx-tabs button').forEach(x => x.setAttribute('aria-selected', String(x.dataset.t === t)));
  root.querySelectorAll('.vx-pane').forEach(p => p.classList.toggle('on', p.dataset.p === t));
  if (innerWidth <= 860) { const b = root.querySelector('.vx-mtabs [data-m="side"]'); if (b) b.click(); }
  if (t === 'ask') { const ta = root.querySelector('#vxAskIn'); if (ta) setTimeout(() => ta.focus(), 50); }
}
function renderSpeakers() {
  const box = root.querySelector('#vxSpk'); if (!box) return;
  const counts = {}; for (const s of V.segments) counts[s.spk] = (counts[s.spk] || 0) + (s.end != null && s.start != null ? s.end - s.start : s.text.length / 14);
  const total = Object.values(counts).reduce((a, b) => a + b, 0) || 1;
  box.innerHTML = '';
  Object.keys(V.speakers).forEach(id => {
    const sp = V.speakers[id];
    const b = el('button', 'vx-chip', `<i style="background:${sp.color}"></i><span>${esc(spkName(V, id))}</span>${sp.role ? `<small>${esc(sp.role)}</small>` : ''}<small>${Math.round(counts[id] / total * 100)}%</small>`);
    b.title = 'שינוי השם, איחוד עם דובר אחר';
    b.onclick = () => speakerMenu(b, id);
    box.append(b);
  });
}
function speakerMenu(anchor, id) {
  const m = menu(anchor);
  const others = Object.keys(V.speakers).filter(x => x !== id);
  m.innerHTML = `<div class="cap">השם של ${esc(spkName(V, id))}</div><input type="text" value="${esc(V.speakers[id].name || '')}" placeholder="לדוגמה: אפרת">
    <button data-x="ok">שמירת השם</button>${others.length ? '<hr><div class="cap">זה בעצם אותו אדם כמו…</div>' + others.map(o => `<button data-m="${o}"><i style="width:10px;height:10px;border-radius:50%;background:${V.speakers[o].color}"></i>${esc(spkName(V, o))}</button>`).join('') : ''}`;
  const inp = m.querySelector('input'); setTimeout(() => { inp.focus(); inp.select(); }, 30);
  const ok = () => { V.speakers[id].name = inp.value.trim(); delete V.speakers[id].auto; closeMenu(); save(); renderSpeakers(); renderLines(); };
  inp.onkeydown = e => { if (e.key === 'Enter') ok(); };
  m.querySelector('[data-x=ok]').onclick = ok;
  m.querySelectorAll('[data-m]').forEach(b => b.onclick = () => {
    const into = b.dataset.m;
    V.segments.forEach(s => { if (s.spk === id) s.spk = into; });
    delete V.speakers[id]; closeMenu(); save(); renderSpeakers(); renderLines(); drawLine();
    say('אוחדו לדובר אחד');
  });
}
function lineSpeakerMenu(anchor, i) {
  const m = menu(anchor);
  m.innerHTML = `<div class="cap">מי אמר את זה?</div>` + Object.keys(V.speakers).map(o => `<button data-s="${o}"><i style="width:10px;height:10px;border-radius:50%;background:${V.speakers[o].color}"></i>${esc(spkName(V, o))}</button>`).join('') + `<button data-s="__new">דובר חדש</button>`;
  m.querySelectorAll('[data-s]').forEach(b => b.onclick = () => {
    let to = b.dataset.s;
    if (to === '__new') { let n = Object.keys(V.speakers).length + 1; while (V.speakers['spk_' + n]) n++; to = 'spk_' + n; V.speakers[to] = { name: '', color: COLORS[(n - 1) % COLORS.length] }; }
    V.segments[i].spk = to;
    for (const k of Object.keys(V.speakers)) if (!V.segments.some(s => s.spk === k)) delete V.speakers[k];
    closeMenu(); save(); renderSpeakers(); renderLines(); drawLine();
  });
}
function lineHtml(s, i) {
  let html;
  if (FIND.q) {
    const q = FIND.q.toLowerCase(), t = s.text, lo = t.toLowerCase(); let k = 0, out = '', p;
    while ((p = lo.indexOf(q, k)) >= 0) { out += esc(t.slice(k, p)) + `<mark data-h>${esc(t.slice(p, p + q.length))}</mark>`; k = p + q.length; }
    html = out + esc(t.slice(k));
  } else html = esc(s.text);
  return html.replace(/\[\[([^\]]*)\]\]/g, (m, x) => `<span class="vx-u" title="לא בטוח שזו המילה">${x}</span>`);
}
function renderLines() {
  const box = root.querySelector('#vxLines'); if (!box) return;
  const frag = document.createDocumentFragment();
  V.segments.forEach((s, i) => {
    const sp = V.speakers[s.spk] || { color: '#888' };
    const row = el('div', 'vx-ln');
    row.dataset.i = i;
    const tr = SHOW_TR && V.tr[SHOW_TR] && V.tr[SHOW_TR][i];
    row.innerHTML = `<div class="who"><button class="nm" style="color:${sp.color}" title="שינוי הדובר של השורה">${esc(spkName(V, s.spk))}</button>${s.start != null ? `<button class="tm" title="השמעה מכאן">${clock(s.start)}</button>` : ''}</div>
      <span class="vx-sb" style="background:${sp.color}"></span><div class="vx-tx" ${FIND.q ? '' : 'contenteditable="' + CE + '"'} spellcheck="false">${lineHtml(s, i)}</div>${tr ? `<div class="trl" dir="auto">${esc(tr)}</div>` : ''}`;
    frag.append(row);
  });
  box.innerHTML = ''; box.append(frag);
  if (activeIdx >= 0) markActive(activeIdx, true);
}
// one listener for all the lines (thousands of lines stay fast)
function wireLines() {
  const box = root.querySelector('#vxLines'); if (!box || box.wired) return; box.wired = true;
  box.addEventListener('click', e => {
    const row = e.target.closest('.vx-ln'); if (!row) return; const i = +row.dataset.i;
    if (e.target.closest('.tm')) { playFrom(V.segments[i].start); return; }
    if (e.target.closest('.nm')) { lineSpeakerMenu(e.target.closest('.nm'), i); return; }
    const w = e.target.closest('.w');
    if (w && audioEl && e.altKey) { const ws = V.segments[i].words[+w.dataset.k]; if (ws && ws[1] != null) playFrom(ws[1]); }
  });
  box.addEventListener('dblclick', e => { const row = e.target.closest('.vx-ln'); if (!row || !audioEl) return; const s = V.segments[+row.dataset.i]; if (s.start != null) playFrom(s.start); });
  box.addEventListener('focusout', e => {
    const tx = e.target.closest('.vx-tx'); if (!tx || FIND.q) return; const i = +tx.closest('.vx-ln').dataset.i;
    const t = tx.innerText.replace(/\s+/g, ' ').trim();
    if (t !== V.segments[i].text) {
      if (!t) { V.segments.splice(i, 1); renderLines(); }
      else { V.segments[i].text = t; delete V.segments[i].words; for (const k in V.tr) if (V.tr[k]) V.tr[k][i] = null; }
      save();
    }
  });
  box.addEventListener('keydown', e => {
    const tx = e.target.closest('.vx-tx'); if (!tx) return;
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); splitLine(+tx.closest('.vx-ln').dataset.i, tx); }
  });
}
// Enter in the middle of a line: the rest becomes a new line (useful when two people were joined)
function splitLine(i, tx) {
  const sel = getSelection(); let pos = tx.innerText.length;
  try { const r = sel.getRangeAt(0).cloneRange(); r.selectNodeContents(tx); r.setEnd(sel.getRangeAt(0).endContainer, sel.getRangeAt(0).endOffset); pos = r.toString().length; } catch (_) {}
  const full = tx.innerText.replace(/\s+/g, ' ');
  const a = full.slice(0, pos).trim(), b = full.slice(pos).trim();
  if (!a || !b) { tx.blur(); return; }
  const s = V.segments[i];
  let mid = null;
  if (s.start != null && s.end != null) mid = s.start + (s.end - s.start) * (a.length / (a.length + b.length));
  V.segments.splice(i, 1, { spk: s.spk, start: s.start, end: mid, text: a }, { spk: s.spk, start: mid, end: s.end, text: b });
  save(); renderLines();
  const nx = root.querySelector(`.vx-ln[data-i="${i + 1}"] .vx-tx`); if (nx) nx.focus();
}
function find(q) {
  FIND.q = q.trim(); FIND.cur = -1;
  renderLines();
  FIND.hits = [...root.querySelectorAll('#vxLines mark[data-h]')];
  root.querySelector('#vxFindN').textContent = FIND.q ? (FIND.hits.length ? FIND.hits.length + ' תוצאות' : 'אין') : '';
  root.querySelector('#vxFindNext').hidden = root.querySelector('#vxFindPrev').hidden = FIND.hits.length < 2;
  if (FIND.hits.length) findStep(1);
}
function findStep(d) {
  if (!FIND.hits.length) return;
  if (FIND.hits[FIND.cur]) FIND.hits[FIND.cur].classList.remove('cur');
  FIND.cur = (FIND.cur + d + FIND.hits.length) % FIND.hits.length;
  const h = FIND.hits[FIND.cur]; h.classList.add('cur'); h.scrollIntoView({ block: 'center', behavior: 'smooth' });
  root.querySelector('#vxFindN').textContent = (FIND.cur + 1) + ' מתוך ' + FIND.hits.length;
}

/* ---------------- insights ---------------- */
function renderIns() {
  const box = root.querySelector('#vxIns'); if (!box) return;
  const a = V.analysis;
  if (!a) {
    box.innerHTML = SMART_CLOUD ? `<div class="vx-card vx-cloud"><h3>${ico.spark}סיכום חכם</h3><p class="vx-hint">סיכום, נקודות עיקריות, משימות, תאריכים ליומן והודעת המשך. רק אם תבקשו: הטקסט (בלי השמע) נשלח לעיבוד ב-Google Gemini.</p><div class="vx-acts"><button class="vx-btn pri" id="vxAnalyze">${ico.spark}לסכם את התמליל</button></div></div>
      <div class="vx-card"><h3>${ico.lock}מה נשאר במכשיר</h3><p class="vx-hint">השמע, התמלול והעריכה נעשים כאן בלבד. אפשר לייצא ל-Word, ל-PDF או לכתוביות בלי שום שליחה.</p></div>`
      : `<div class="vx-card"><h3>${ico.lock}הכול במכשיר</h3><p class="vx-hint">השמע, התמלול והעריכה נעשים כאן בלבד. הייצוא ל-Word, ל-PDF ולכתוביות בתפריט ״ייצוא ושיתוף״.</p></div>`;
    const b = box.querySelector('#vxAnalyze'); if (b) b.onclick = reanalyze;
    return;
  }
  const list = (arr, cls) => arr && arr.length ? `<ul${cls ? ` class="${cls}"` : ''}>${arr.map(x => `<li>${esc(x)}</li>`).join('')}</ul>` : '';
  const tasks = (a.tasks || []).map((t, k) => `<div class="vx-task${V.tasksDone[k] ? ' done' : ''}"><input type="checkbox" data-k="${k}"${V.tasksDone[k] ? ' checked' : ''} aria-label="בוצע"><div class="d"><span>${esc(t.what)}</span>${t.who || t.when ? `<small>${esc([nameFor(t.who), t.when].filter(Boolean).join(' · '))}</small>` : ''}</div>${t.date ? `<button class="vx-ic" data-cal="t${k}" title="הוספה ליומן">${ico.cal}</button>` : ''}</div>`).join('');
  const dates = (a.dates || []).map((d, k) => `<div class="vx-task"><div class="d"><span>${esc(d.what)}</span><small>${esc([d.date ? fmtDate(d.date) : '', d.time].filter(Boolean).join(' · '))}</small></div>${d.date ? `<button class="vx-ic" data-cal="d${k}" title="הוספה ליומן">${ico.cal}</button>` : ''}</div>`).join('');
  box.innerHTML = `
    <div class="vx-card"><h3>בקצרה</h3><p>${esc(a.summary || '')}</p>${a.tone ? `<p class="vx-hint" style="margin-top:6px">אווירה: ${esc(a.tone)}</p>` : ''}</div>
    ${a.points && a.points.length ? `<div class="vx-card"><h3>הנקודות העיקריות</h3>${list(a.points)}</div>` : ''}
    ${a.decisions && a.decisions.length ? `<div class="vx-card"><h3>מה סוכם</h3>${list(a.decisions)}</div>` : ''}
    ${tasks ? `<div class="vx-card"><h3>משימות</h3>${tasks}</div>` : ''}
    ${dates ? `<div class="vx-card"><h3>תאריכים ופגישות</h3>${dates}${(a.dates || []).filter(d => d.date).length + (a.tasks || []).filter(t => t.date).length > 1 ? `<div class="vx-acts"><button class="vx-btn sm" data-cal="all">${ico.cal}הכול ליומן</button></div>` : ''}</div>` : ''}
    ${a.facts && a.facts.length ? `<div class="vx-card"><h3>פרטים לשמור</h3>${list(a.facts)}</div>` : ''}
    ${a.questions && a.questions.length ? `<div class="vx-card"><h3>נשאר פתוח</h3>${list(a.questions)}</div>` : ''}
    ${a.followup ? `<div class="vx-card"><h3>הודעת המשך מוכנה</h3><div class="vx-msg" id="vxFollow" contenteditable="${CE}" dir="auto">${esc(a.followup)}</div>
      <div class="vx-acts"><button class="vx-btn sm" data-f="wa">${ico.wa}וואטסאפ</button><button class="vx-btn sm" data-f="mail">מייל</button><button class="vx-btn sm" data-f="copy">${ico.copy}העתקה</button></div></div>` : ''}
    <div class="vx-acts" style="margin-top:6px"><button class="vx-btn sm" id="vxMinutes">פרוטוקול ל-Word</button><button class="vx-btn sm" id="vxRe">${ico.spark}לסכם מחדש</button></div>`;
  box.querySelectorAll('.vx-task input').forEach(c => c.onchange = () => { V.tasksDone[c.dataset.k] = c.checked; c.closest('.vx-task').classList.toggle('done', c.checked); save(); });
  box.querySelectorAll('[data-cal]').forEach(b => b.onclick = () => calendar(b.dataset.cal));
  const fol = box.querySelector('#vxFollow');
  if (fol) fol.oninput = () => { V.analysis.followup = fol.innerText; save(); };
  box.querySelectorAll('[data-f]').forEach(b => b.onclick = () => {
    const t = (fol && fol.innerText) || a.followup;
    if (b.dataset.f === 'wa') window.open('https://wa.me/?text=' + encodeURIComponent(t), '_blank', 'noopener');
    else if (b.dataset.f === 'mail') location.href = 'mailto:?subject=' + encodeURIComponent(V.title) + '&body=' + encodeURIComponent(t);
    else copy(t);
  });
  box.querySelector('#vxMinutes').onclick = minutesDocx;
  box.querySelector('#vxRe').onclick = reanalyze;
}
const nameFor = who => { if (!who) return ''; const m = String(who).match(/spk_\d+/); return m && V.speakers[m[0]] ? spkName(V, m[0]) : who; };
const fmtDate = d => { try { const x = new Date(d + 'T12:00:00'); return isNaN(x) ? d : x.toLocaleDateString('he-IL', { weekday: 'short', day: 'numeric', month: 'short' }); } catch (_) { return d; } };
async function reanalyze() {
  if (!cloudOk('סיכום התמליל')) return;
  const box = root.querySelector('#vxIns');
  box.innerHTML = '<div class="vx-wait"><i></i>קורא את השיחה ומסכם…</div>';
  try {
    const r = await call('analyze', { transcript: forAi(V), kind: V.kind });
    applyAnalysis(V, r.analysis); save(true); renderSpeakers(); renderLines(); renderIns();
    if (V.namedNote) { say(V.namedNote); delete V.namedNote; }
  } catch (e) { say(friendly(e)); renderIns(); }
}

/* ---------------- ask ---------------- */
function renderAsk() {
  const box = root.querySelector('#vxAsk'); if (!box) return;
  if (!SMART_CLOUD) { box.innerHTML = ''; return; }
  box.innerHTML = `<p class="vx-hint">שאלות על התמליל. התשובות מגיעות מ-Google Gemini על סמך הטקסט בלבד (בלי השמע).</p><div class="vx-sugg">${['מה הנושאים העיקריים?', 'הכן לי סיכום ללמידה למבחן', 'מה המשימות שנאמרו?', 'הסבר במילים פשוטות את החלק המרכזי'].map(q => `<button class="vx-btn sm" data-q="${esc(q)}">${esc(q)}</button>`).join('')}</div>
    <div class="vx-chat" id="vxChat"></div>
    <div class="vx-ask"><textarea id="vxAskIn" rows="1" placeholder="שאלה על השיחה…" aria-label="שאלה על השיחה"></textarea><button class="vx-btn pri" id="vxAskGo">שאלה</button></div>`;
  const chat = box.querySelector('#vxChat');
  V.chat.forEach(h => chat.append(qa(h.q, h.a)));
  const ta = box.querySelector('#vxAskIn');
  ta.oninput = () => { ta.style.height = 'auto'; ta.style.height = Math.min(140, ta.scrollHeight) + 'px'; };
  ta.onkeydown = e => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); askGo(); } };
  box.querySelector('#vxAskGo').onclick = () => askGo();
  box.querySelectorAll('[data-q]').forEach(b => b.onclick = () => askGo(b.dataset.q));
}
function qa(q, a) {
  const w = el('div'); w.style.display = 'grid'; w.style.gap = '6px';
  w.append(el('div', 'vx-q', esc(q)));
  const an = el('div', 'vx-a'); an.setAttribute('dir', 'auto');
  if (a == null) an.innerHTML = '<div class="vx-wait" style="padding:0"><i></i>מחפש בשיחה…</div>';
  else setAnswer(an, a);
  w.append(an); return w;
}
function setAnswer(an, a) {
  an.innerHTML = esc(a).replace(/\[(\d{1,2}:\d{2}(?::\d{2})?)\]/g, (m, t) => `<a data-t="${t}" title="השמעה מהרגע הזה">${t}</a>`);
  an.querySelectorAll('a[data-t]').forEach(x => x.onclick = () => { const p = x.dataset.t.split(':').map(Number); playFrom(p.reduce((s, v) => s * 60 + v, 0)); });
  const c = el('button', 'vx-ic', ico.copy); c.title = 'העתקה'; c.style.float = 'left'; c.onclick = () => copy(a); an.prepend(c);
}
async function askGo(q) {
  const ta = root.querySelector('#vxAskIn');
  q = String(q || ta.value).trim(); if (!q) return;
  if (!cloudOk('שאלה על התמליל')) return;
  ta.value = ''; ta.style.height = 'auto';
  const chat = root.querySelector('#vxChat'); const node = qa(q, null); chat.append(node); node.scrollIntoView({ block: 'end' });
  try {
    const r = await call('ask', { transcript: forAi(V), question: q, history: V.chat.slice(-6) });
    setAnswer(node.querySelector('.vx-a'), r.answer || '');
    V.chat.push({ q, a: r.answer || '' }); save(); count('voice_ask');
  } catch (e) { node.querySelector('.vx-a').textContent = friendly(e); }
}

/* ---------------- translation ---------------- */
function trMenu(anchor) {
  const m = menu(anchor);
  m.innerHTML = `<div class="cap">תרגום מתחת לכל שורה</div>` + Object.entries(TR_TARGETS).map(([k, v]) => `<button data-l="${k}">${SHOW_TR === k ? '✓ ' : ''}${v}</button>`).join('') + (SHOW_TR ? '<hr><button data-l="">בלי תרגום</button>' : '');
  m.querySelectorAll('[data-l]').forEach(b => b.onclick = async () => {
    closeMenu(); const l = b.dataset.l;
    if (l && !(V.tr[l] && V.tr[l].every(Boolean)) && !cloudOk('תרגום התמליל')) return;
    if (!l) { SHOW_TR = ''; renderLines(); updTrBtn(); return; }
    const have = V.tr[l];
    const missing = V.segments.map((s, i) => have && have[i] ? null : i).filter(i => i != null);
    if (missing.length) {
      const btn = root.querySelector('#vxTrBtn'); btn.disabled = true; btn.textContent = 'מתרגם…';
      try {
        const r = await call('translate', { target: l, segments: missing.map(i => V.segments[i].text) });
        const arr = have || new Array(V.segments.length).fill(null);
        missing.forEach((i, k) => { arr[i] = r.segments[k]; });
        V.tr[l] = arr; save(); count('voice_translate');
      } catch (e) { say(friendly(e)); btn.disabled = false; updTrBtn(); return; }
      btn.disabled = false;
    }
    SHOW_TR = l; renderLines(); updTrBtn();
  });
}
const updTrBtn = () => { const b = root.querySelector('#vxTrBtn'); if (b) b.textContent = SHOW_TR ? 'תרגום: ' + TR_TARGETS[SHOW_TR] : 'תרגום'; };

/* ---------------- the player: synced highlight + speaker timeline ---------------- */
let raf = 0, cv = null;
function setupPlayer(s) {
  stopAudio();
  const old = root.querySelector('.vx-player,.vx-noaudio'); if (old) old.remove();
  wireLines();
  if (!V.blob) { const n = el('div', 'vx-noaudio', 'השמע של התמליל הזה לא שמור במכשיר, לכן אין השמעה'); s.append(n); return; }
  audioUrl = URL.createObjectURL(V.blob);
  audioEl = new Audio(); audioEl.preload = 'metadata'; audioEl.src = audioUrl;
  const p = el('div', 'vx-player');
  p.innerHTML = `<div class="vx-pl-ctl"><button class="bw" title="10 שניות אחורה" aria-label="10 שניות אחורה">${ico.bwd}</button><button class="pp" title="השמעה (רווח)" aria-label="השמעה">${ico.play}</button><button class="fw" title="10 שניות קדימה" aria-label="10 שניות קדימה">${ico.fwd}</button></div>
    <div class="vx-pl-line"><canvas aria-label="ציר הזמן לפי דוברים"></canvas><div class="vx-pl-time"><span class="cur">0:00</span><span class="dur">${clock(V.duration)}</span></div></div>
    <div class="vx-pl-end"><button class="spd" title="מהירות">1×</button><button class="fol" aria-pressed="true" title="גלילה אחרי ההשמעה">עוקב</button></div>`;
  s.append(p);
  cv = p.querySelector('canvas');
  p.querySelector('.pp').onclick = togglePlay;
  p.querySelector('.bw').onclick = () => seek(audioEl.currentTime - 10);
  p.querySelector('.fw').onclick = () => seek(audioEl.currentTime + 10);
  const sp = p.querySelector('.spd'), speeds = [1, 1.25, 1.5, 1.75, 2, 0.75];
  sp.onclick = () => { const k = (speeds.indexOf(audioEl.playbackRate) + 1) % speeds.length; audioEl.playbackRate = speeds[k]; sp.textContent = speeds[k] + '×'; };
  const fol = p.querySelector('.fol'); fol.setAttribute('aria-pressed', String(FOLLOW));
  fol.onclick = () => { FOLLOW = !FOLLOW; fol.setAttribute('aria-pressed', String(FOLLOW)); };
  audioEl.onplay = () => { p.querySelector('.pp').innerHTML = ico.pause; loop(); };
  audioEl.onpause = audioEl.onended = () => { p.querySelector('.pp').innerHTML = ico.play; cancelAnimationFrame(raf); tick(); };
  audioEl.onloadedmetadata = () => { if (isFinite(audioEl.duration) && audioEl.duration > 0) { if (!V.duration) { V.duration = audioEl.duration; save(); } p.querySelector('.dur').textContent = clock(audioEl.duration); } drawLine(); };
  // the timeline runs right → left, like the text
  let dragging = false;
  const at = e => { const r = cv.getBoundingClientRect(); const x = (e.touches ? e.touches[0].clientX : e.clientX); return Math.max(0, Math.min(1, (r.right - x) / r.width)) * dur(); };
  cv.onpointerdown = e => { dragging = true; cv.setPointerCapture(e.pointerId); seek(at(e)); };
  cv.onpointermove = e => { if (dragging) seek(at(e)); };
  cv.onpointerup = () => { dragging = false; };
  new ResizeObserver(drawLine).observe(cv);
  drawLine();
}
const dur = () => (audioEl && isFinite(audioEl.duration) && audioEl.duration > 0) ? audioEl.duration : (V && V.duration) || 1;
function drawLine() {
  if (!cv || !V) return;
  const r = cv.getBoundingClientRect(), dpr = devicePixelRatio || 1;
  cv.width = Math.max(1, r.width * dpr); cv.height = Math.max(1, r.height * dpr);
  const g = cv.getContext('2d'), W = cv.width, H = cv.height, D = dur();
  g.clearRect(0, 0, W, H);
  g.fillStyle = 'rgba(255,255,255,.12)'; g.fillRect(0, H * .3, W, H * .4);
  for (const s of V.segments) {
    if (s.start == null) continue;
    const a = s.start / D, b = (s.end != null ? s.end : s.start + 3) / D;
    g.fillStyle = (V.speakers[s.spk] || {}).color || '#888';
    g.fillRect(W - b * W, H * .3, Math.max(1.5 * dpr, (b - a) * W), H * .4);
  }
  if (audioEl) {
    const x = W - (audioEl.currentTime / D) * W;
    g.fillStyle = '#FFD43B'; g.fillRect(x - 1.5 * dpr, 0, 3 * dpr, H);
  }
}
function loop() { tick(); raf = requestAnimationFrame(loop); }
let lastWord = null;
function tick() {
  if (!audioEl || !V) return;
  const t = audioEl.currentTime;
  const c = root.querySelector('.vx-pl-time .cur'); if (c) c.textContent = clock(t);
  drawLine();
  // the line being said now
  const segs = V.segments; let lo = 0, hi = segs.length - 1, k = -1;
  while (lo <= hi) { const m = (lo + hi) >> 1; if (segs[m].start != null && segs[m].start <= t + 0.15) { k = m; lo = m + 1; } else hi = m - 1; }
  if (k !== activeIdx) markActive(k);
  // the word being said now
  if (lastWord) { lastWord.classList.remove('on'); lastWord = null; }
  const s = segs[k];
  if (s && s.words && !FIND.q) {
    let wi = -1; for (let j = 0; j < s.words.length; j++) { if (s.words[j][1] != null && s.words[j][1] <= t) wi = j; else if (s.words[j][1] != null) break; }
    if (wi >= 0) { const w = root.querySelector(`.vx-ln[data-i="${k}"] .w[data-k="${wi}"]`); if (w) { w.classList.add('on'); lastWord = w; } }
  }
}
function markActive(k, noScroll) {
  const prev = root.querySelector('.vx-ln.on'); if (prev) prev.classList.remove('on');
  activeIdx = k; if (k < 0) return;
  const row = root.querySelector(`.vx-ln[data-i="${k}"]`); if (!row) return;
  row.classList.add('on');
  if (FOLLOW && !noScroll && audioEl && !audioEl.paused && !(document.activeElement && document.activeElement.closest && document.activeElement.closest('.vx-ln .vx-tx'))) row.scrollIntoView({ block: 'center', behavior: 'smooth' });
}
function togglePlay() { if (!audioEl) return; if (audioEl.paused) audioEl.play().catch(() => say('לא הצלחתי להשמיע את הקובץ')); else audioEl.pause(); }
function seek(t) { if (!audioEl) return; audioEl.currentTime = Math.max(0, Math.min(dur() - 0.05, t)); tick(); }
function playFrom(t) { if (!audioEl) { say('אין השמעה לתמליל הזה'); return; } if (t == null) return; seek(Math.max(0, t - 0.3)); audioEl.play().catch(() => {}); }
function stopAudio() { cancelAnimationFrame(raf); if (audioEl) { audioEl.pause(); audioEl.src = ''; audioEl = null; } if (audioUrl) { URL.revokeObjectURL(audioUrl); audioUrl = ''; } cv = null; }

/* ---------------- recording: the text is written on the device while recording ---------------- */
const REC = { mr: null, chunks: [], stream: null, t0: 0, acc: 0, timer: 0, cutT: 0, an: null, ac: null, node: null, lock: null, mime: '', pend: [], pendLen: 0, off: 0, floor: [], segs: [], chain: null, engine: null, engineOk: false, lang: 'he', kind: 'lesson', cancel: false, dur: 0 };
async function startRec() {
  if (!consentOk()) return;
  if (!navigator.mediaDevices || !window.MediaRecorder) { say('הדפדפן הזה לא יכול להקליט. נסו Chrome, Edge או Safari מעודכן'); return; }
  try { REC.stream = await navigator.mediaDevices.getUserMedia({ audio: { echoCancellation: true, noiseSuppression: true, autoGainControl: true, channelCount: 1 } }); }
  catch (e) { say(e && e.name === 'NotAllowedError' ? 'צריך לאשר גישה למיקרופון כדי להקליט' : 'לא נמצא מיקרופון'); return; }
  const types = ['audio/webm;codecs=opus', 'audio/ogg;codecs=opus', 'audio/mp4', 'audio/webm'];
  REC.mime = types.find(t => MediaRecorder.isTypeSupported && MediaRecorder.isTypeSupported(t)) || '';
  REC.mr = new MediaRecorder(REC.stream, REC.mime ? { mimeType: REC.mime, audioBitsPerSecond: 48000 } : undefined);
  Object.assign(REC, { chunks: [], acc: 0, pend: [], pendLen: 0, off: 0, floor: [], segs: [], chain: Promise.resolve(), engineOk: false, cancel: false, lang: prefs.lang || 'he', kind: prefs.kind });
  REC.mr.ondataavailable = e => { if (e.data && e.data.size) REC.chunks.push(e.data); };
  REC.mr.onstop = finishRec;
  show('rec');
  const s = root.querySelector('#vxRec');
  s.innerHTML = `<div class="vx-rec"><div class="tm" id="vxRecT">0:00</div><div class="st" id="vxRecS">מקליט</div><canvas id="vxRecC"></canvas>
    <div class="ctl"><button class="vx-btn" id="vxRecP">השהיה</button><button class="big" id="vxRecStop" title="סיום" aria-label="סיום ההקלטה">${ico.stop}</button><button class="vx-btn" id="vxRecX">ביטול</button></div>
    <div class="vx-live" id="vxLive" dir="auto" data-empty="הטקסט נכתב כאן תוך כדי ההקלטה, בתוך המכשיר."></div>
    <p class="vx-hint" id="vxRecE"></p>
    <p class="vx-hint">${ico.lock} ההקלטה והתמלול נשארים במכשיר. המסך יישאר דלוק בזמן ההקלטה. כדאי להניח את הטלפון קרוב למי שמדבר.</p></div>`;
  s.querySelector('#vxRecStop').onclick = stopRec;
  s.querySelector('#vxRecX').onclick = () => { if (confirm('למחוק את ההקלטה?')) { REC.cancel = true; stopRec(); } };
  s.querySelector('#vxRecP').onclick = e => {
    if (REC.mr.state === 'recording') { REC.mr.pause(); REC.acc += Date.now() - REC.t0; e.target.textContent = 'המשך'; s.querySelector('#vxRecS').textContent = 'מושהה'; }
    else { REC.mr.resume(); REC.t0 = Date.now(); e.target.textContent = 'השהיה'; s.querySelector('#vxRecS').textContent = 'מקליט'; }
  };
  REC.mr.start(1000); REC.t0 = Date.now();
  REC.timer = setInterval(() => { const t = (REC.acc + (REC.mr && REC.mr.state === 'recording' ? Date.now() - REC.t0 : 0)) / 1000; const x = root.querySelector('#vxRecT'); if (x) x.textContent = clock(t); }, 250);
  try { REC.lock = await navigator.wakeLock.request('screen'); } catch (_) {}
  capture();
  // the engine starts while the recording runs (the first time it is downloaded)
  const files = {}; const note = root.querySelector('#vxRecE');
  REC.engine = ASR.load(REC.lang, p => { files[p.file] = p; const { l, t } = progressSum(files); if (t > 2e6 && note) note.textContent = 'מוריד את מנוע התמלול (פעם אחת): ' + Math.round(l / t * 100) + '%'; })
    .then(() => { REC.engineOk = true; if (note) note.textContent = ''; }, e => { REC.engineOk = false; if (note) note.textContent = 'מנוע התמלול לא נטען (' + friendly(e) + '). ההקלטה נשמרת ותתומלל אחר כך.'; });
  REC.cutT = setInterval(() => liveCut(false), 1000);
  addEventListener('beforeunload', recGuard);
}
const recGuard = e => { if (REC.mr && REC.mr.state !== 'inactive') { e.preventDefault(); e.returnValue = ''; } };
// the microphone is also read as raw sound (16 kHz) for the engine; the recording file is for listening
function capture() {
  try {
    REC.ac = new (window.AudioContext || window.webkitAudioContext)();
    if (REC.ac.resume) REC.ac.resume();
    const src = REC.ac.createMediaStreamSource(REC.stream);
    REC.an = REC.ac.createAnalyser(); REC.an.fftSize = 1024; src.connect(REC.an);
    const node = REC.ac.createScriptProcessor(4096, 1, 1), mute = REC.ac.createGain(); mute.gain.value = 0;
    node.onaudioprocess = e => {
      if (!REC.mr || REC.mr.state !== 'recording') return;
      const y = resample(e.inputBuffer.getChannelData(0), REC.ac.sampleRate, SR);
      const copy = new Float32Array(y); REC.pend.push(copy); REC.pendLen += copy.length;
    };
    src.connect(node); node.connect(mute); mute.connect(REC.ac.destination); REC.node = node;
    const c = root.querySelector('#vxRecC'), data = new Uint8Array(REC.an.fftSize), hist = [];
    const draw = () => {
      if (!REC.an || !c.isConnected) return;
      REC.an.getByteTimeDomainData(data);
      let peak = 0; for (const v of data) peak = Math.max(peak, Math.abs(v - 128) / 128);
      hist.push(REC.mr && REC.mr.state === 'recording' ? peak : 0); if (hist.length > 160) hist.shift();
      const r = c.getBoundingClientRect(), dpr = devicePixelRatio || 1; c.width = r.width * dpr; c.height = r.height * dpr;
      const g = c.getContext('2d'), W = c.width, H = c.height, bw = W / 160;
      g.fillStyle = (getComputedStyle(root).getPropertyValue('--ink') || '').trim() || '#16213A';
      hist.forEach((v, i) => { const h = Math.max(2 * dpr, Math.min(1, v * 2.2) * H); g.fillRect(W - (i + 1) * bw, (H - h) / 2, Math.max(1, bw - 2 * dpr), h); });
      requestAnimationFrame(draw);
    };
    draw();
  } catch (e) { console.warn(e); }
}
// every second: if a pause came after 12+ seconds of sound, that piece goes to the engine
function liveCut(final) {
  if (!REC.pendLen) return;
  const pcm = new Float32Array(REC.pendLen); let o = 0; for (const p of REC.pend) { pcm.set(p, o); o += p.length; }
  const E = energies(pcm);
  for (let i = 0; i < E.length; i += 5) REC.floor.push(E[i]);
  if (REC.floor.length > 4000) REC.floor.splice(0, REC.floor.length - 4000);
  const thr = threshold(REC.floor);
  let e = findCut(E, 0, E.length, thr, final);
  if (!e) { REC.pend = [pcm]; return; }
  if (final && e >= E.length) e = Math.ceil(pcm.length / F);
  const cut = Math.min(pcm.length, e * F);
  const piece = pcm.slice(0, cut), base = REC.off;
  REC.off += cut; REC.pend = cut < pcm.length ? [pcm.slice(cut)] : []; REC.pendLen = pcm.length - cut;
  const sp = speech(E, 0, Math.min(E.length, e), thr);
  if (sp) {
    const audio = piece.slice(sp[0] * F, sp[1] * F), start = (base + sp[0] * F) / SR, end = (base + sp[1] * F) / SR;
    REC.chain = REC.chain.then(async () => {
      await REC.engine; if (!REC.engineOk) return;
      const r = await ASR.text(audio, REC.lang);
      const text = clean(r.text, REC.kind);
      if (!text) return;
      REC.segs.push({ spk: 'spk_1', start, end, text });
      const box = root.querySelector('#vxLive') || root.querySelector('#vxWorkLive');
      if (box) { box.append(el('p', '', esc(text))); while (box.children.length > 30) box.firstChild.remove(); box.scrollTop = box.scrollHeight; }
    }).catch(err => console.warn(err));
  }
  if (final && REC.pendLen) liveCut(true);
}
function stopRec() {
  if (!REC.mr) return;
  REC.dur = (REC.acc + (REC.mr.state === 'recording' ? Date.now() - REC.t0 : 0)) / 1000;
  if (REC.mr.state !== 'inactive') REC.mr.stop(); else finishRec();
}
async function finishRec() {
  clearInterval(REC.timer); clearInterval(REC.cutT); removeEventListener('beforeunload', recGuard);
  try { REC.stream.getTracks().forEach(t => t.stop()); } catch (_) {}
  try { REC.node && (REC.node.onaudioprocess = null); REC.ac && REC.ac.close(); } catch (_) {} REC.an = null; REC.node = null;
  try { REC.lock && REC.lock.release(); } catch (_) {} REC.lock = null;
  const mime = (REC.mr && REC.mr.mimeType) || REC.mime || 'audio/webm';
  const chunks = REC.chunks; REC.chunks = []; REC.mr = null;
  if (REC.cancel) { REC.cancel = false; REC.pend = []; REC.pendLen = 0; show('new'); return; }
  if (!chunks.length) { say('לא הוקלט כלום'); show('new'); return; }
  const x = /mp4/.test(mime) ? 'm4a' : /ogg/.test(mime) ? 'ogg' : 'webm';
  const d = new Date(), name = `הקלטה ${d.toLocaleDateString('he-IL')} ${d.toLocaleTimeString('he-IL', { hour: '2-digit', minute: '2-digit' })}.${x}`.replace(/[\\/:]/g, '-');
  const f = new File(chunks, name, { type: mime.split(';')[0] });
  count('voice_record');
  // the engine never loaded (e.g. no internet the first time): the file is transcribed later, like any file
  if (!REC.engineOk) { JOB.queue.push({ file: f, opts: { duration: REC.dur || null, kind: REC.kind, lang: REC.lang } }); if (!JOB.busy) run(); return; }
  liveCut(true);
  show('work');
  const w = root.querySelector('#vxWork');
  w.innerHTML = `<div class="vx-work"><h2>מסיים לתמלל</h2><div class="f">${esc(name)}</div><div class="vx-wait" style="justify-content:center"><i></i>כותב את המשפטים האחרונים…</div></div>`;
  await REC.chain;
  const id = uid();
  const segs = normalize(REC.segs.sort((a, b) => a.start - b.start));
  const item = { id, title: name.replace(/\.[a-z0-9]+$/i, ''), fileName: name, created: Date.now(), updated: Date.now(), kind: REC.kind, lang: REC.lang, duration: REC.dur || null,
    segments: segs, speakers: speakersOf(segs, {}), analysis: null, tr: {}, chat: [], tasksDone: {}, status: segs.length ? 'ready' : 'failed', error: segs.length ? '' : 'לא נשמע דיבור בהקלטה', engine: 'local' };
  await DB.putAudio(id, f); await DB.put(item); refreshCount();
  await openItem(id);
}

/* ---------------- library ---------------- */
async function renderLib() {
  const s = root.querySelector('#vxLib');
  s.innerHTML = `<div class="vx-wrap vx-lib"><h2 style="margin-top:0">הספרייה</h2>
    <label class="vx-find" style="max-width:420px;margin-bottom:14px">${ico.search}<input type="search" id="vxLibQ" placeholder="חיפוש בכל התמלילים"></label>
    <div class="vx-items" id="vxLibList"><div class="vx-wait"><i></i>טוען…</div></div></div>`;
  let items = [];
  try { items = await DB.all(); } catch (_) {}
  items.sort((a, b) => (b.updated || b.created) - (a.updated || a.created));
  const list = s.querySelector('#vxLibList');
  const draw = q => {
    list.innerHTML = '';
    const ql = q.trim().toLowerCase();
    const got = !ql ? items : items.filter(it => (it.title || '').toLowerCase().includes(ql) || (it.segments || []).some(x => x.text.toLowerCase().includes(ql)) || ((it.analysis && it.analysis.summary) || '').toLowerCase().includes(ql));
    if (!got.length) { list.innerHTML = `<p class="vx-empty">${items.length ? 'לא נמצא' : 'עוד אין תמלילים. בוחרים קובץ או מקליטים, והתמלילים נשמרים כאן, במכשיר.'}</p>`; return; }
    got.forEach(it => list.append(itemRow(it, ql)));
  };
  s.querySelector('#vxLibQ').oninput = e => draw(e.target.value);
  draw('');
}
function itemRow(it, q) {
  const b = el('div', 'vx-item'); b.setAttribute('role', 'button'); b.tabIndex = 0;
  const k = it.kind === 'voice' ? ico.wa : it.kind === 'dictation' ? ico.file : ico.mic;
  let sub = [dayStr(it.created), it.duration ? longDur(it.duration) : '', it.status === 'failed' ? 'נכשל' : it.status === 'working' ? 'בעבודה…' : ''].filter(Boolean).join(' · ');
  if (q) { const hit = (it.segments || []).find(x => x.text.toLowerCase().includes(q)); if (hit) sub = '“' + hit.text.slice(0, 90) + '”'; }
  b.innerHTML = `<span class="k">${k}</span><span class="t"><b>${esc(it.title || 'תמליל')}</b><span>${esc(sub)}</span></span><button class="x" title="מחיקה" aria-label="מחיקה">${ico.trash}</button>`;
  b.onclick = e => { if (e.target.closest('.x')) return; openItem(it.id); };
  b.onkeydown = e => { if (e.key === 'Enter') openItem(it.id); };
  b.querySelector('.x').onclick = async () => { if (!confirm('למחוק את "' + (it.title || 'התמליל') + '" ואת השמע שלו מהמכשיר?')) return; await DB.del(it.id); b.remove(); refreshCount(); };
  return b;
}

/* ---------------- menus ---------------- */
let MENU = null;
function menu(anchor) {
  closeMenu();
  const m = el('div', 'vx-menu'); m.setAttribute('role', 'menu'); m.dir = 'rtl';
  root.append(m); MENU = m;
  requestAnimationFrame(() => {
    const r = anchor.getBoundingClientRect(), mw = m.offsetWidth, mh = m.offsetHeight;
    let left = Math.min(innerWidth - mw - 8, Math.max(8, r.right - mw));
    let top = r.bottom + 6; if (top + mh > innerHeight - 8) top = Math.max(8, r.top - mh - 6);
    m.style.left = left + 'px'; m.style.top = top + 'px';
  });
  setTimeout(() => addEventListener('pointerdown', outside, true), 0);
  return m;
}
const outside = e => { if (MENU && !MENU.contains(e.target)) closeMenu(); };
function closeMenu() { if (MENU) { MENU.remove(); MENU = null; } removeEventListener('pointerdown', outside, true); }

/* ---------------- export ---------------- */
function exportMenu(anchor) {
  const m = menu(anchor);
  const canShare = !!navigator.share;
  m.innerHTML = `
    <button data-x="wordEdit">${ico.file}פתיחה בעורך Word של סימן</button>
    <button data-x="pdf">${ico.file}PDF בסימן (לסימון, חתימה ושליחה)</button>
    <button data-x="docx">${ico.out}הורדת Word</button>
    ${SMART_CLOUD ? `<button data-x="minutes">${ico.spark}פרוטוקול / סיכום ל-Word</button>` : ''}
    <hr>
    <button data-x="srt">כתוביות לסרטון (SRT)</button>
    <button data-x="vtt">כתוביות לאתר (VTT)</button>
    <button data-x="txt">טקסט פשוט</button>
    <button data-x="copy">${ico.copy}העתקת התמליל</button>
    ${canShare ? `<button data-x="share">שיתוף…</button>` : ''}
    <hr>
    <button data-x="audio">הורדת השמע</button>
    <button data-x="del" style="color:#C2362B">${ico.trash}מחיקת התמליל</button>`;
  m.querySelectorAll('[data-x]').forEach(b => b.onclick = () => { closeMenu(); doExport(b.dataset.x); });
}
const safeName = t => String(t || V.title || 'תמליל').replace(/[\\/:*?"<>|\u0000-\u001f]/g, '').trim().slice(0, 90) || 'תמליל';
function download(name, data, type) {
  const blob = data instanceof Blob ? data : new Blob([data], { type });
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 30000);
}
async function copy(t) { try { await navigator.clipboard.writeText(t); say('הועתק'); } catch (_) { say('לא הצלחתי להעתיק'); } }
function textExport() {
  const lines = [V.title, dayStr(V.created) + (V.duration ? ' · ' + longDur(V.duration) : ''), ''];
  if (V.analysis && V.analysis.summary) lines.push('סיכום:', V.analysis.summary, '');
  V.segments.forEach((s, i) => {
    lines.push(`${s.start != null ? '[' + clock2(s.start) + '] ' : ''}${spkName(V, s.spk)}: ${s.text}`);
    if (SHOW_TR && V.tr[SHOW_TR] && V.tr[SHOW_TR][i]) lines.push('    ' + V.tr[SHOW_TR][i]);
  });
  return lines.join('\n').replace(/\[\[([^\]]*)\]\]/g, '$1');
}
function subs(kind) {
  const ts = (s, sep) => { s = Math.max(0, s); const h = Math.floor(s / 3600), m = Math.floor(s % 3600 / 60), x = Math.floor(s % 60), ms = Math.round((s % 1) * 1000); return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(x).padStart(2, '0')}${sep}${String(ms).padStart(3, '0')}`; };
  const cues = [];
  V.segments.forEach((s, i) => {
    if (s.start == null) return;
    const end = s.end != null ? s.end : (V.segments[i + 1] && V.segments[i + 1].start) || s.start + 4;
    const words = s.words && s.words.length ? s.words : null;
    // long lines are cut into short subtitles of up to ~7 seconds / 84 letters
    const txt = (SHOW_TR && V.tr[SHOW_TR] && V.tr[SHOW_TR][i]) || s.text.replace(/\[\[([^\]]*)\]\]/g, '$1');
    if (words && !SHOW_TR) {
      let cur = [], cs = null;
      for (const w of words) {
        if (cs == null) cs = w[1] != null ? w[1] : s.start;
        cur.push(w);
        const ce = w[2] != null ? w[2] : cs + 2;
        if (cur.map(x => x[0]).join(' ').length > 80 || ce - cs > 7) { cues.push([cs, ce, cur.map(x => x[0]).join(' ')]); cur = []; cs = null; }
      }
      if (cur.length) cues.push([cs, Math.max(cs + 0.8, cur[cur.length - 1][2] || end), cur.map(x => x[0]).join(' ')]);
    } else {
      const parts = txt.length > 84 ? txt.match(/[^.!?،,]+[.!?،,]*\s*/g) || [txt] : [txt];
      const chunks = []; let c = '';
      for (const p of parts) { if ((c + p).length > 84 && c) { chunks.push(c.trim()); c = ''; } c += p; }
      if (c.trim()) chunks.push(c.trim());
      const total = chunks.reduce((a, x) => a + x.length, 0) || 1; let t = s.start;
      for (const ch of chunks) { const d = (end - s.start) * ch.length / total; cues.push([t, t + d, ch]); t += d; }
    }
  });
  if (kind === 'vtt') return 'WEBVTT\n\n' + cues.map((c, k) => `${k + 1}\n${ts(c[0], '.')} --> ${ts(c[1], '.')}\n${c[2]}\n`).join('\n');
  return cues.map((c, k) => `${k + 1}\n${ts(c[0], ',')} --> ${ts(c[1], ',')}\n${c[2]}\n`).join('\n');
}
async function doExport(x) {
  const name = safeName();
  try {
    if (x === 'txt') { const b = new TextEncoder().encode(textExport()); if (typeof saveBytes === 'function') await saveBytes(name + '.txt', b, 'txt'); else download(name + '.txt', b, 'text/plain'); }
    else if (x === 'copy') copy(textExport());
    else if (x === 'srt') download(name + '.srt', '\ufeff' + subs('srt'), 'application/x-subrip');
    else if (x === 'vtt') download(name + '.vtt', subs('vtt'), 'text/vtt');
    else if (x === 'audio') { if (!V.blob) { say('השמע לא שמור במכשיר'); return; } download(V.fileName || name + '.' + (ext(V.blob.name) || 'webm'), V.blob); }
    else if (x === 'share') {
      const docx = await transcriptDocx();
      const f = new File([docx], name + '.docx', { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      try {
        if (navigator.canShare && navigator.canShare({ files: [f] })) await navigator.share({ files: [f], title: V.title });
        else await navigator.share({ title: V.title, text: textExport().slice(0, 60000) });
      } catch (e) { if (e.name !== 'AbortError') say('השיתוף לא הצליח'); }
    }
    else if (x === 'docx') { const b = await transcriptDocx(); if (typeof saveBytes === 'function') await saveBytes(name + '.docx', b, 'docx'); else download(name + '.docx', b); count('voice_word'); }
    else if (x === 'wordEdit') {
      const b = await transcriptDocx(); count('voice_word');
      if (typeof openWordEditor !== 'function') { download(name + '.docx', b); return; }
      close(); await openWordEditor(b, name + '.docx');
    }
    else if (x === 'pdf') {
      const b = await transcriptDocx();
      if (typeof loadEditLibs !== 'function') { say('ההמרה ל-PDF לא זמינה כאן'); return; }
      busyOn('מכין PDF…');
      await loadEditLibs('docx');
      const r = await OFFICE_EDIT.toPdf('docx', b, m => busyOn(m));
      busyOff(); close();
      const pdf = r.pdf instanceof Uint8Array ? r.pdf : new Uint8Array(r.pdf);
      await openFile(new File([pdf], name + '.pdf', { type: 'application/pdf' }));
      count('voice_pdf');
    }
    else if (x === 'minutes') minutesDocx();
    else if (x === 'del') { if (!confirm('למחוק את התמליל ואת השמע שלו מהמכשיר?')) return; await DB.del(V.id); V = null; stopAudio(); refreshCount(); show('lib'); }
  } catch (e) { busyOff(); console.error(e); say('הייצוא נכשל: ' + (e.message || e)); }
}
const busyOn = m => { try { busy(true, m); } catch (_) {} };
const busyOff = () => { try { busy(false); } catch (_) {} };
function calendar(which) {
  const a = V.analysis || {}; const evs = [];
  const add = (title, date, time, desc) => { if (!/^\d{4}-\d{2}-\d{2}$/.test(date || '')) return; evs.push({ title, date, time, desc }); };
  if (which === 'all') { (a.dates || []).forEach(d => add(d.what, d.date, d.time, V.title)); (a.tasks || []).forEach(t => add(t.what, t.date, '', [nameFor(t.who), t.when, V.title].filter(Boolean).join(' · '))); }
  else if (which[0] === 'd') { const d = (a.dates || [])[+which.slice(1)]; if (d) add(d.what, d.date, d.time, V.title); }
  else { const t = (a.tasks || [])[+which.slice(1)]; if (t) add(t.what, t.date, '', [nameFor(t.who), t.when, V.title].filter(Boolean).join(' · ')); }
  if (!evs.length) { say('אין תאריך מדויק להוסיף ליומן'); return; }
  const icsEsc = s => String(s || '').replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\n/g, '\\n');
  const stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d+/, '');
  const body = evs.map((e, k) => {
    const d = e.date.replace(/-/g, '');
    let when;
    if (/^\d{1,2}:\d{2}$/.test(e.time || '')) { const [h, mi] = e.time.split(':').map(Number); const t = String(h).padStart(2, '0') + String(mi).padStart(2, '0') + '00'; const h2 = String(Math.min(23, h + 1)).padStart(2, '0') + String(mi).padStart(2, '0') + '00'; when = `DTSTART:${d}T${t}\r\nDTEND:${d}T${h2}`; }
    else { const nx = new Date(e.date + 'T12:00:00'); nx.setDate(nx.getDate() + 1); when = `DTSTART;VALUE=DATE:${d}\r\nDTEND;VALUE=DATE:${nx.toISOString().slice(0, 10).replace(/-/g, '')}`; }
    return `BEGIN:VEVENT\r\nUID:${V.id}-${k}-${d}@siman\r\nDTSTAMP:${stamp}\r\n${when}\r\nSUMMARY:${icsEsc(e.title)}\r\nDESCRIPTION:${icsEsc(e.desc)}\r\nEND:VEVENT`;
  }).join('\r\n');
  download((evs.length === 1 ? safeName(evs[0].title) : safeName()) + '.ics', `BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Siman//Voice//HE\r\nCALSCALE:GREGORIAN\r\n${body}\r\nEND:VCALENDAR\r\n`, 'text/calendar');
  count('voice_calendar');
}

/* ---------------- Word (built here, RTL-correct: bidi paragraphs, rtl runs, complex-script font) ---------------- */
const X = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
function wrun(text, o, rtl) {
  o = o || {};
  const sz = o.size || 24;
  return `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="Arial"/>${o.b ? '<w:b/><w:bCs/>' : ''}${o.color ? `<w:color w:val="${o.color.replace('#', '')}"/>` : ''}<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>${rtl ? '<w:rtl/>' : ''}<w:lang w:val="${rtl ? 'he-IL' : 'en-US'}" w:bidi="he-IL"/></w:rPr><w:t xml:space="preserve">${X(text)}</w:t></w:r>`;
}
function para(runs, o, rtl) {
  o = o || {};
  return `<w:p><w:pPr>${o.style ? `<w:pStyle w:val="${o.style}"/>` : ''}${o.keep ? '<w:keepNext/>' : ''}<w:spacing w:before="${o.before || 0}" w:after="${o.after == null ? 120 : o.after}" w:line="${o.line || 300}" w:lineRule="auto"/>${o.ind ? (rtl ? `<w:ind w:right="${o.ind}"/>` : `<w:ind w:left="${o.ind}"/>`) : ''}${rtl ? '<w:bidi/>' : ''}${o.center ? '<w:jc w:val="center"/>' : ''}${o.border ? '<w:pBdr><w:bottom w:val="single" w:sz="8" w:space="4" w:color="FFD43B"/></w:pBdr>' : ''}</w:pPr>${runs.join('')}</w:p>`;
}
async function buildDocx(bodyXml, rtl, title) {
  if (!window.JSZip) await loadScript('jszip.min.js');
  const z = new JSZip();
  z.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`);
  z.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`);
  z.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${X(title)}</dc:title><dc:creator>סימן</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString().replace(/\.\d+/, '')}</dcterms:created></cp:coreProperties>`);
  z.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>`);
  z.file('word/settings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:defaultTabStop w:val="720"/>${rtl ? '<w:bidi/>' : ''}<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`);
  const rp = `<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="Arial"/><w:sz w:val="24"/><w:szCs w:val="24"/>${rtl ? '<w:rtl/>' : ''}<w:lang w:val="${rtl ? 'he-IL' : 'en-US'}" w:bidi="he-IL"/></w:rPr>`;
  const hs = (id, nm, sz, color) => `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${nm}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/>${rtl ? '<w:bidi/>' : ''}<w:outlineLvl w:val="${id === 'Title' ? 0 : id === 'Heading1' ? 0 : 1}"/></w:pPr><w:rPr><w:b/><w:bCs/><w:color w:val="${color}"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr></w:style>`;
  z.file('word/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault>${rp}</w:rPrDefault><w:pPrDefault><w:pPr>${rtl ? '<w:bidi/>' : ''}<w:spacing w:after="120" w:line="300" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr>${rtl ? '<w:bidi/>' : ''}</w:pPr>${rp}</w:style>${hs('Title', 'Title', 40, '16213A')}${hs('Heading1', 'heading 1', 30, '16213A')}${hs('Heading2', 'heading 2', 26, '2E58CC')}</w:styles>`);
  z.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${bodyXml}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1300" w:right="1300" w:bottom="1300" w:left="1300" w:header="708" w:footer="708" w:gutter="0"/>${rtl ? '<w:bidi/>' : ''}</w:sectPr></w:body></w:document>`);
  return z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}
const cleanU = t => String(t || '').replace(/\[\[([^\]]*)\]\]/g, '$1');
function transcriptDocx() {
  const rtl = isRtl(V.lang) || (!V.lang && guessRtl(plain(V.segments)));
  const P = (runs, o) => para(runs, o, rtl), R = (t, o) => wrun(t, o, rtl);
  const a = V.analysis;
  const out = [];
  out.push(P([R(V.title, { b: true, size: 40, color: '#16213A' })], { style: 'Title', after: 60, border: true }));
  out.push(P([R([dayStr(V.created), V.duration ? longDur(V.duration) : '', Object.keys(V.speakers).map(k => spkName(V, k)).join(', ')].filter(Boolean).join('  |  '), { size: 20, color: '#5D6982' })], { after: 240 }));
  if (a) {
    out.push(P([R('סיכום', { b: true, size: 28 })], { style: 'Heading1', keep: true }));
    out.push(P([R(a.summary || '')]));
    const sec = (h, arr) => { if (!arr || !arr.length) return; out.push(P([R(h, { b: true, size: 26, color: '#2E58CC' })], { style: 'Heading2', keep: true })); arr.forEach(x => out.push(P([R('•  ' + x)], { ind: 360, hang: 0, after: 60 }))); };
    sec('הנקודות העיקריות', a.points);
    sec('מה סוכם', a.decisions);
    sec('משימות', (a.tasks || []).map(t => t.what + (t.who || t.when ? ' (' + [nameFor(t.who), t.when].filter(Boolean).join(', ') + ')' : '')));
    sec('תאריכים', (a.dates || []).map(d => d.what + (d.date || d.time ? ' — ' + [d.date ? fmtDate(d.date) : '', d.time].filter(Boolean).join(' ') : '')));
    sec('נשאר פתוח', a.questions);
  }
  out.push(P([R('התמליל המלא', { b: true, size: 28 })], { style: 'Heading1', keep: true, before: 240 }));
  let prev = null;
  V.segments.forEach((s, i) => {
    const sp = V.speakers[s.spk] || { color: '#16213A' };
    const head = s.spk !== prev;
    if (head) out.push(P([R(spkName(V, s.spk), { b: true, color: sp.color, size: 22 }), ...(s.start != null ? [R('   ' + clock2(s.start), { color: '#8A93A6', size: 18 })] : [])], { keep: true, before: 160, after: 20 }));
    out.push(P([R(cleanU(s.text))], { after: SHOW_TR && V.tr[SHOW_TR] && V.tr[SHOW_TR][i] ? 20 : 100 }));
    if (SHOW_TR && V.tr[SHOW_TR] && V.tr[SHOW_TR][i]) out.push(para([wrun(V.tr[SHOW_TR][i], { color: '#5D6982', size: 22 }, isRtl(SHOW_TR))], { after: 100 }, isRtl(SHOW_TR)));
    prev = s.spk;
  });
  out.push(P([R('תומלל בסימן', { size: 16, color: '#8A93A6' })], { before: 360, center: true }));
  return buildDocx(out.join(''), rtl, V.title);
}
async function minutesDocx() {
  if (!cloudOk('כתיבת פרוטוקול מסודר')) return;
  const box = root.querySelector('#vxMinutes'); if (box) { box.disabled = true; box.textContent = 'כותב פרוטוקול…'; }
  try {
    const sp = {}; Object.keys(V.speakers).forEach(k => { sp[k] = spkName(V, k); });
    const r = await call('minutes', { transcript: forAi(V), meta: { speakers: sp, date: dayStr(V.created) } });
    const m = r.minutes || {};
    const rtl = isRtl(V.lang) || (!V.lang && guessRtl(plain(V.segments)));
    const P = (runs, o) => para(runs, o, rtl), R = (t, o) => wrun(t, o, rtl);
    const out = [];
    out.push(P([R('פרוטוקול', { size: 22, color: '#5D6982' })], { after: 0 }));
    out.push(P([R(m.title || V.title, { b: true, size: 40 })], { style: 'Title', border: true, after: 120 }));
    out.push(P([R('תאריך: ', { b: true }), R(dayStr(V.created))], { after: 40 }));
    if (m.participants && m.participants.length) out.push(P([R('משתתפים: ', { b: true }), R(m.participants.join(', '))], { after: 240 }));
    (m.topics || []).forEach((t, k) => {
      out.push(P([R(`${k + 1}. ${t.title || ''}`, { b: true, size: 28 })], { style: 'Heading1', keep: true }));
      if (t.discussion) out.push(P([R(t.discussion)]));
      (t.decisions || []).forEach(d => out.push(P([R('הוחלט: ', { b: true, color: '#0F766E' }), R(d)], { ind: 360, after: 60 })));
    });
    if (m.tasks && m.tasks.length) {
      out.push(P([R('משימות להמשך', { b: true, size: 28 })], { style: 'Heading1', keep: true, before: 200 }));
      m.tasks.forEach(t => out.push(P([R('•  ' + (t.what || '')), ...(t.who ? [R('  — ' + t.who, { b: true })] : []), ...(t.when ? [R(' (' + t.when + ')', { color: '#5D6982' })] : [])], { ind: 360, after: 60 })));
    }
    if (m.next) out.push(P([R('המשך: ', { b: true }), R(m.next)], { before: 200 }));
    out.push(P([R('נכתב בסימן מתוך הקלטת הישיבה', { size: 16, color: '#8A93A6' })], { before: 360, center: true }));
    const bytes = await buildDocx(out.join(''), rtl, m.title || V.title);
    count('voice_minutes');
    const name = 'פרוטוקול - ' + safeName() + '.docx';
    if (typeof openWordEditor === 'function' && confirm('הפרוטוקול מוכן. לפתוח אותו בעורך Word של סימן? (ביטול = הורדה)')) { close(); await openWordEditor(bytes, name); }
    else if (typeof saveBytes === 'function') await saveBytes(name, bytes, 'docx'); else download(name, bytes);
  } catch (e) { say(friendly(e)); }
  finally { if (box && box.isConnected) { box.disabled = false; box.textContent = 'פרוטוקול ל-Word'; } }
}

/* ---------------- hooking into סימן ---------------- */
// 1. any audio/video file opened in סימן (picker, drag, "open with", tabs) comes here
try {
  const prevOpen = openFile;
  openFile = async function (file, handle) {
    if (file && isMedia(file) && !/\.(pdf|docx?|xlsx|pptx)$/i.test(file.name || '')) {
      if (typeof needReg === 'function' && needReg()) return prevOpen(file, handle);   // sign-up first; it comes back here after
      if (!rd(OK_KEY, false)) { open('new'); pendingFile = file; say('סמנו שאתם צד לשיחה ובחרו ״המשך״'); showContinue(); return; }
      enqueue(file); return;
    }
    return prevOpen(file, handle);
  };
} catch (e) { console.warn('siman-voice: openFile hook', e); }
let pendingFile = null;
function showContinue() {
  const s = root.querySelector('#vxNew .vx-set'); if (!s || s.querySelector('#vxGo')) return;
  const b = el('button', 'vx-btn pri', 'המשך לתמלול: ' + esc(pendingFile.name)); b.id = 'vxGo'; b.style.justifySelf = 'start';
  b.onclick = () => { if (!consentOk()) return; const f = pendingFile; pendingFile = null; b.remove(); enqueue(f); };
  s.append(b);
}
// 2. the desktop file picker also offers audio and video
try {
  if (typeof OPEN_TYPES !== 'undefined' && OPEN_TYPES[0] && OPEN_TYPES[0].accept) {
    OPEN_TYPES[0].description = 'מסמכים, שמע ווידאו';
    OPEN_TYPES[0].accept['audio/*'] = ['.mp3', '.m4a', '.wav', '.ogg', '.opus', '.aac', '.flac', '.webm', '.amr'];
    OPEN_TYPES[0].accept['video/mp4'] = ['.mp4', '.mov'];
  }
} catch (_) {}
// 3. a button on the opening screen, next to the other quick actions
function addEntry() {
  const q = document.querySelector('.wl-quick');
  if (q && !q.querySelector('[data-qa="voice"]')) {
    const b = el('button', 'qa', `<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="9" y="3" width="6" height="11" rx="3"/><path d="M5.5 11a6.5 6.5 0 0013 0M12 17.5V21M8.5 21h7"/></svg>לתמלל שיחה או הקלטה<span style="font-size:11px;font-weight:700;background:var(--marker,#FFD43B);color:#16213A;border-radius:6px;padding:1px 6px;margin-inline-start:4px">חדש</span>`);
    b.dataset.qa = 'voice';
    b.onclick = () => { count('voice_open'); open('new'); };
    const tr = q.querySelector('[data-qa="tr"]');
    if (tr) tr.after(b); else q.prepend(b);
  }
  const l = document.querySelector('.wl-list');
  if (l && !l.querySelector('.vx-li')) {
    const li = el('li', 'vx-li', '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="9" y="3" width="6" height="11" rx="3"/><path d="M5.5 11a6.5 6.5 0 0013 0M12 17.5V21"/></svg>תמלול שיחות, הודעות קוליות וישיבות');
    l.append(li);
  }
}
addEntry();
setTimeout(addEntry, 1500);   // the phone layout builds its screen a moment later
// 4. ?voice opens the transcription screen; ?voice=shared = a file shared from another app (Android share sheet)
async function fromUrl() {
  const u = new URL(location.href); const v = u.searchParams.get('voice');
  if (v == null) return;
  u.searchParams.delete('voice'); history.replaceState(null, '', u.pathname + (u.search || '') + u.hash);
  if (v !== 'shared') { open('new'); return; }
  try {
    const c = await caches.open('siman-share');
    const keys = await c.keys();
    const files = [];
    for (const k of keys) {
      const r = await c.match(k); if (!r) continue;
      const blob = await r.blob();
      const name = decodeURIComponent(r.headers.get('x-name') || 'קובץ משותף');
      files.push(new File([blob], name, { type: r.headers.get('content-type') || blob.type }));
      await c.delete(k);
    }
    if (!files.length) { open('new'); return; }
    for (const f of files) await openFile(f);
  } catch (e) { console.error(e); say('לא הצלחתי לקבל את הקובץ ששותף'); }
}
setTimeout(fromUrl, 600);

window.SIMAN_VOICE = { open: () => open('new'), openFile: f => enqueue(f), library: () => open('lib') };
})();

