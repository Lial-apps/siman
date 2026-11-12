/* =====================================================================
   סימן 6.4 — siman-guard.js: a pass for סימן's own smart services
   The smart services (siman-ai, siman-voice, siman-translate) use סימן's Gemini key. So that only people
   using the app can reach them, and so that one heavy user cannot use up the free daily quota for everyone:
   · the first time a smart feature is used, Cloudflare Turnstile checks in the background that a person is
     using the app (no puzzle in almost all cases), and the server gives this device a signed pass for 6 hours;
   · every request to those three servers carries the pass ("x-siman-pass");
   · a person who added their own free Gemini key (settings → translation → Gemini) sends it too
     ("x-user-key"): their requests use their own quota and have no daily share. The key is used for that
     request only and is never stored on the server.
   Nothing here runs until a smart feature is used: the Turnstile script is loaded only then.
   With SITEKEY empty, this file does nothing at all.
   ===================================================================== */
(() => {
'use strict';
if (window.SIMAN_GUARD) return;
const SITEKEY = '0x4AAAAAAF95W5HhrqvA6owz';                    // Cloudflare Turnstile site key (public, not a secret)
const PASS_URL = 'https://siman-ai.siman-app.workers.dev/pass';
const HOSTS = ['siman-ai.siman-app.workers.dev', 'siman-voice.siman-app.workers.dev', 'siman-translate.siman-app.workers.dev'];
const PASS_KEY = 'siman.guard.pass', DEV_KEY = 'siman.guard.dev';
const rd = k => { try { return localStorage.getItem(k); } catch (_) { return null; } };
const wr = (k, v) => { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, v); } catch (_) {} };
const orig = window.fetch.bind(window);

function device() {
  let d = rd(DEV_KEY);
  if (!/^[a-z0-9]{16,40}$/.test(d || '')) {
    const a = new Uint8Array(16); crypto.getRandomValues(a);
    d = [...a].map(b => 'abcdefghijkmnpqrstuvwxyz23456789'[b % 32]).join(''); wr(DEV_KEY, d);
  }
  return d;
}
// the person's own Gemini key, only when they chose Gemini with their own key in the settings
function ownKey() {
  try {
    const c = JSON.parse(rd('siman.settings') || '{}');
    const k = String(c.geminiKey || '').trim();
    return c.provider === 'gemini' && /^(AIza|AQ\.)/.test(k) ? k : '';
  } catch (_) { return ''; }
}
function saved() {
  try { const p = JSON.parse(rd(PASS_KEY) || 'null'); if (p && p.pass && p.until > Date.now() + 120000) return p.pass; } catch (_) {}
  return '';
}

/* ---------- Turnstile: loaded the first time a pass is needed ---------- */
let tsLoad = null;
function loadTurnstile() {
  if (window.turnstile) return Promise.resolve();
  return tsLoad || (tsLoad = new Promise((res, rej) => {
    const s = document.createElement('script');
    s.src = 'https://challenges.cloudflare.com/turnstile/v0/api.js?render=explicit';
    s.async = true; s.onload = () => res(); s.onerror = () => { tsLoad = null; rej(new Error('turnstile blocked')); };
    document.head.append(s);
  }));
}
// a small box at the bottom of the screen; it shows only if Cloudflare asks the person to tick a box
let box = null;
function holder() {
  if (box && box.isConnected) return box;
  box = document.createElement('div');
  box.id = 'simanGuard';
  box.setAttribute('aria-live', 'polite');
  box.style.cssText = 'position:fixed;left:50%;bottom:calc(16px + env(safe-area-inset-bottom,0px));transform:translateX(-50%);z-index:2147483600;max-width:calc(100vw - 24px)';
  document.body.append(box);
  return box;
}
async function turnstileToken() {
  await loadTurnstile();
  const el = holder(); el.innerHTML = '';
  const inner = document.createElement('div'); el.append(inner);
  return await new Promise((res, rej) => {
    let id = null;
    const done = (fn, v) => { clearTimeout(t); try { window.turnstile.remove(id); } catch (_) {} el.innerHTML = ''; fn(v); };
    const t = setTimeout(() => done(rej, new Error('turnstile timeout')), 120000);
    id = window.turnstile.render(inner, {
      sitekey: SITEKEY, action: 'pass', appearance: 'interaction-only', execution: 'render',
      language: (document.documentElement.lang || 'he').slice(0, 2),
      callback: tok => done(res, tok),
      'error-callback': () => done(rej, new Error('turnstile error')),
      'unsupported-callback': () => done(rej, new Error('turnstile unsupported')),
    });
  });
}

/* ---------- the pass ---------- */
let getting = null;
function pass(force) {
  if (!force) { const p = saved(); if (p) return Promise.resolve(p); }
  return getting || (getting = (async () => {
    const t = await turnstileToken();
    const r = await orig(PASS_URL, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ t, dev: device() }) });
    const d = await r.json().catch(() => ({}));
    if (r.status === 503 && d.error && d.error.code === 'NOPASS') return '';          // the servers are not set up yet: go on without
    if (!r.ok || !d.pass) throw new Error((d.error && d.error.message) || 'pass failed');
    wr(PASS_KEY, JSON.stringify({ pass: d.pass, until: Date.now() + (d.ttl || 3600) * 1000 }));
    return d.pass;
  })().finally(() => { getting = null; }));
}

/* ---------- every request to the three servers carries the pass ---------- */
function guarded(input, init) {
  if (!SITEKEY) return false;
  let u;
  try { u = new URL(typeof input === 'string' ? input : input instanceof URL ? input.href : input.url, location.href); } catch (_) { return false; }
  if (!HOSTS.includes(u.host) || u.pathname === '/pass') return false;
  const method = String((init && init.method) || (input instanceof Request && input.method) || 'GET').toUpperCase();
  return method === 'POST';                                  // GET = the transcription engine files (no pass needed)
}
async function send(input, init, force) {
  let p = '';
  try { p = await pass(force); } catch (_) { p = ''; }     // no pass: the server answers with a clear message
  const h = new Headers((init && init.headers) || (input instanceof Request ? input.headers : undefined));
  if (p) h.set('x-siman-pass', p);
  const k = ownKey(); if (k) h.set('x-user-key', k);
  return orig(input, { ...(init || {}), headers: h });
}
window.fetch = async function (input, init) {
  if (!guarded(input, init)) return orig(input, init);
  let r = await send(input, init, false);
  if (r.status === 401) {
    // the pass ran out or the server's secret changed: one new pass, one more try
    let code = '';
    try { code = ((await r.clone().json()).error || {}).code || ''; } catch (_) {}
    if (code === 'PASS') { wr(PASS_KEY, null); r = await send(input, init, true); }
  }
  return r;
};
window.SIMAN_GUARD = { on: !!SITEKEY, pass, device, ownKey: () => !!ownKey() };
})();
