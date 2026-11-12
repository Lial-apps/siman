// סימן — service worker: works offline, picks up new versions automatically
const CACHE = 'siman-v117'; // 6.5.0 — Excel charts: the file's charts are shown (editor, marking, PDF) and new charts are added from selected cells (Insert › Chart), moved, resized, edited, deleted, saved in Excel's format; 6.4.2 — right-click paste really offers הדבק (checked in real Edge); one transcription engine (measured: two were slower); 6.4.1 — forms: lines kept in pieces are joined (whole places, full labels, signature found, no footer lines); places offered only on real forms again (5.10.2); the 5.10 converter, copy-as-text, numbers and faster transcription back; right-click paste on the home screen; 6.4 — the smart services need a pass (Cloudflare Turnstile) and keep a fair daily share per device; your own Gemini key = no limit; 6.3.8 — Word: text never runs past the last page (no page cap), Excel grows rows/columns while scrolling; 6.3.7 — printing from the browser's own menu is not blank; 6.3.6 — the home card really hides; 6.3.5 — check report (Ctrl K); 6.3.4 — the home card never freezes; 6.3.3 — closing a file from the home card asks inside the card; 6.3.2 — resume card fixed (stuck files, × to close), editors no longer break when left while opening; 6.3.1 — new file inside the Office editors, the smart server does not wait for a stuck model; 6.3 — home from anywhere, Word forms fill, autofill and my details; 6.2.1 — Word forms: values written on their blank lines; 6.2.1 — anonymous device count (siman-stats), real app version; WhatsApp on phones: link when files can't be shared; 6.2 — no sign-up for local tools, Hebrew date, business stamp, in-person signing, glossary; 6.1 — auto redaction, PDF compression, ask all open files; 6.0 — karaoke transcript, stars, chapters, study tools, cite highlight, JPG/PNG/Markdown
const ASSETS = ['./', './index.html', './manifest.webmanifest', './pdf.min.js', './pdf.worker.min.js',
  './pdf-lib.min.js', './fontkit.umd.min.js', './icon-192.png', './icon-512.png', './icon-maskable-512.png', './jszip.min.js', './pptxgen.min.js', './office.js', './office-edit.js', './word-editor.js', './excel-editor.js', './docx-preview.min.js', './html2canvas.min.js', './xlsx.mini.min.js', './font-secular-he.woff2', './font-secular-lat.woff2', './send.js', './siman-plus.js', './siman-guard.js', './siman-media.js', './ppt-editor.js', './office-ai.js', './office-translate.js', './a11y.js', './admin-panel.js', './font-rubik-400.ttf', './font-rubik-700.ttf', './siman-phone.css', './siman-phone.js', './siman-voice.js', './siman-asr-worker.js', './transformers.min.js', './siman-back.js', './siman-v7.css', './siman-v7.js', './siman-scan.js', './siman-merge.js', './siman-ux.js', './siman-nb.css', './siman-nb.js', './siman-fill.js', './siman-pro.js', './siman-convert.js', './siman-home.js', './xl-charts.js'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  // only the app's own old versions are removed: the transcription engine (transformers-cache, ~165MB) stays
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => /^siman-v\d+$/.test(k) && k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
// 5.9.6: a weak connection does not hold the app on a blank screen: after 6 seconds the saved copy is used
const fresh = req => new Promise((res, rej) => { const t = setTimeout(() => rej(new Error('slow')), 6000);
  fetch(req, { cache: 'no-cache' }).then(r => { clearTimeout(t); res(r); }, e => { clearTimeout(t); rej(e); }); });
self.addEventListener('fetch', e => {
  const req = e.request;
  // a file shared to סימן from another app (Android share sheet: WhatsApp voice messages, call recordings, PDFs):
  // it is kept for a moment in a cache and the app opens and picks it up (siman-voice.js)
  if (req.method === 'POST' && /\/share-target\/?$/.test(new URL(req.url).pathname)) {
    e.respondWith((async () => {
      try {
        const form = await req.formData();
        const files = form.getAll('files').filter(f => f && typeof f === 'object' && f.size);
        const c = await caches.open('siman-share');
        let n = 0;
        for (const f of files) {
          await c.put('./__shared/' + Date.now() + '-' + (n++), new Response(f, { headers: { 'content-type': f.type || 'application/octet-stream', 'x-name': encodeURIComponent(f.name || 'קובץ') } }));
        }
      } catch (_) {}
      return Response.redirect('./?voice=shared', 303);
    })());
    return;
  }
  if (req.method !== 'GET' || new URL(req.url).origin !== location.origin) return; // APIs go straight to the network
  // 5.9.3: video and audio go straight to the browser. A video asks for byte ranges, which the cache cannot answer
  // (Safari then shows nothing), and an old copy kept here hid the new landing-page video
  if (req.headers.has('range') || /\.(mp4|webm|mov|m4a|mp3|ogg)$/i.test(new URL(req.url).pathname)) return;
  const path = new URL(req.url).pathname;
  const isPage = req.mode === 'navigate' || /\/(index\.html)?$/.test(path);
  // the app's own code: always the newest version when online, so an upload shows up on the next refresh
  const isAppCode = /\/(office|office-edit|word-editor|excel-editor|siman-register|send|siman-plus|ppt-editor|a11y|office-ai|office-translate|siman-voice|siman-asr-worker|siman-back|siman-scan|siman-merge|siman-ux|siman-fill|siman-convert|siman-pro|siman-home|admin-panel|siman-media|siman-ai|siman-share|xl-charts)\.js$/.test(path) || /\/siman-(phone|v7|nb)\.(css|js)$/.test(path);
  if (isAppCode) {
    e.respondWith(fresh(req).then(r => { if (r.ok) { const copy = r.clone(); caches.open(CACHE).then(c => c.put(req, copy)); } return r; })
      .catch(() => caches.match(req, { ignoreSearch: true }).then(h => h || fetch(req))));
    return;
  }
  if (isPage) {
    // newest page when online, cached page when offline. Only the app itself is kept as the app page:
    // other pages (like sign.html, opened from a link) are kept under their own address
    const isApp = /\/(index\.html)?$/.test(path);
    e.respondWith(fresh(req).then(r => { if (r.ok) { const copy = r.clone(); caches.open(CACHE).then(c => c.put(isApp ? './index.html' : req, copy)); } return r; })
      .catch(() => isApp ? caches.match('./index.html') : caches.match(req)));
    return;
  }
  e.respondWith(caches.match(req, { ignoreSearch: true }).then(hit => hit || fetch(req).then(r => {
    if (r.ok) { const copy = r.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
    return r;
  })));
});
