// סימן — מנוע התמלול שרץ בתוך המכשיר (Web Worker, נטען מ-siman-voice.js)
// השמע לא יוצא מהמכשיר. מה שיורד מהאינטרנט הוא רק המנוע עצמו, פעם אחת:
//   עברית: Ozen-v1 (itayinbar/Ozen-v1, Apache-2.0) — מודל עברי קטן שזוקק מ-ivrit.ai whisper-large-v3-turbo
//   שפות אחרות: Whisper small (Xenova/whisper-small)
// הקבצים נשמרים במטמון של הדפדפן ("transformers-cache"), ואחר כך התמלול עובד גם בלי אינטרנט.
// הספרייה: transformers.min.js (Transformers.js 4.3.1), מונחת ליד האפליקציה.
import { env, AutoTokenizer, AutoProcessor, WhisperForConditionalGeneration, pipeline } from './transformers.min.js';

env.allowLocalModels = false;
env.useBrowserCache = true;

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

async function build(key, device) {
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

async function load(key) {
  if (E && E.key === key) return E;
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
    } catch (err) { fails.push(device + ': ' + ((err && err.message) || err)); console.warn('siman-asr:', device, err); }
  }
  throw new Error(fails.join(' | ') || 'engine did not load');
}

async function handle(m) {
  try {
    if (m.type === 'load') {
      const e = await load(m.engine);
      postMessage({ id: m.id, device: e.device });
    } else if (m.type === 'run') {
      const key = (m.lang || 'he') === 'he' ? 'he' : 'multi';
      const e = await load(key);
      const t0 = performance.now();
      let text;
      try { text = await e.run(m.audio, m.lang); }
      catch (err) {
        // a graphics-card failure in the middle: once more on the processor
        if (e.device !== 'wasm') { try { await e.dispose(); } catch (_) {} E = await build(key, 'wasm'); text = await E.run(m.audio, m.lang); }
        else throw err;
      }
      postMessage({ id: m.id, text: String(text || '').trim(), ms: Math.round(performance.now() - t0) });
    }
  } catch (err) {
    postMessage({ id: m.id, error: String((err && err.message) || err) });
  }
}

// one job at a time: the model is not run twice in parallel
self.onmessage = ({ data }) => { queue = queue.then(() => handle(data)); };
