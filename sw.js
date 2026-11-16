// סימן — service worker: works offline, picks up new versions automatically
const CACHE = 'siman-v127'; // 6.6.7 — WhatsApp on a phone: if WhatsApp does not open after the share menu, the dialog stays and offers a link (opens the WhatsApp app directly on Android), the menu again, or a download; remembered for next time; 6.6.6 — line spacing: in Word the menu says where it goes ("לכל המסמך" ticked when the cursor is not in the text, before it went silently to the first paragraph), table cells that lost their paragraph get one back; PowerPoint has a line-spacing button too (saved as PowerPoint spacing); reading: no Arabic letters inside Hebrew words; 6.6.5 — PowerPoint: a slide that is one picture (NotebookLM, Canva, exported decks) can be turned into editable text: the words are found with their places, taken out of the picture and put back as real text boxes (one slide or all, three at a time, undo restores the picture); 6.6.4 — reading a page on a phone: the screen stays on, a line cut because the app went to the background is sent again when it comes back, and a cut line is no longer called "no internet" (nor stops the other pages); 6.6.3 — full test round (desktop + phone): the command search works again after a form is opened, a left-to-right Excel sheet opens at column A (on phones it opened at M-Q), page thumbnails also prepare sideways (phone row); 6.6.2 — big files: thumbnails drawn only when seen, pages asked for 40 at a time, form fields looked for meanwhile; Excel undo copies the sheet once per edit (not twice) and is capped by size; ביטול on long jobs (OCR of all pages, compress, sensitive-data search; also phone Back); send for signature asks for the email and checks the 15MB limit before the fields are placed, and goes on after sign-up; translations and first questions already answered are reused in the session; 6.6.1 — review fixes: the page is half the size (the Alef font is a file, not written in twice), a slow network still keeps the new version, the Office assistant gets the whole request, a scan no longer marks the open file as saved, translation/OCR never write into another file, OCR survives page changes and one bad page, pages copied once per file (smaller files), less memory per page, Excel sorting keeps formulas right and scrolls faster, drafts only when something changed; 6.6.0 — מדף למסמך: printed pages read whole in seconds (auto: handwriting found by the quick reading goes on to the careful one), two pages at a time, pictures prepared three at a time; 6.5.2 — Excel: the selection is painted like Excel (grey tint over the cell colour, green frame around every area, marked headings); 6.5.1 — Excel: Ctrl+click adds areas (A + E for a chart), drag across headings selects columns/rows, a chart from several areas; 6.5.0 — Excel charts: the file's charts are shown (editor, marking, PDF) and new charts are added from selected cells (Insert › Chart), moved, resized, edited, deleted, saved in Excel's format; 6.4.2 — right-click paste really offers הדבק (checked in real Edge); one transcription engine (measured: two were slower); 6.4.1 — forms: lines kept in pieces are joined (whole places, full labels, signature found, no footer lines); places offered only on real forms again (5.10.2); the 5.10 converter, copy-as-text, numbers and faster transcription back; right-click paste on the home screen; 6.4 — the smart services need a pass (Cloudflare Turnstile) and keep a fair daily share per device; your own Gemini key = no limit; 6.3.8 — Word: text never runs past the last page (no page cap), Excel grows rows/columns while scrolling; 6.3.7 — printing from the browser's own menu is not blank; 6.3.6 — the home card really hides; 6.3.5 — check report (Ctrl K); 6.3.4 — the home card never freezes; 6.3.3 — closing a file from the home card asks inside the card; 6.3.2 — resume card fixed (stuck files, × to close), editors no longer break when left while opening; 6.3.1 — new file inside the Office editors, the smart server does not wait for a stuck model; 6.3 — home from anywhere, Word forms fill, autofill and my details; 6.2.1 — Word forms: values written on their blank lines; 6.2.1 — anonymous device count (siman-stats), real app version; WhatsApp on phones: link when files can't be shared; 6.2 — no sign-up for local tools, Hebrew date, business stamp, in-person signing, glossary; 6.1 — auto redaction, PDF compression, ask all open files; 6.0 — karaoke transcript, stars, chapters, study tools, cite highlight, JPG/PNG/Markdown
const ASSETS = ['./', './index.html', './manifest.webmanifest', './pdf.min.js', './pdf.worker.min.js',
  './pdf-lib.min.js', './fontkit.umd.min.js', './icon-192.png', './icon-512.png', './icon-maskable-512.png', './jszip.min.js', './pptxgen.min.js', './office.js', './office-edit.js', './word-editor.js', './excel-editor.js', './docx-preview.min.js', './html2canvas.min.js', './xlsx.mini.min.js', './font-secular-he.woff2', './font-secular-lat.woff2', './send.js', './siman-plus.js', './siman-guard.js', './siman-media.js', './ppt-editor.js', './office-ai.js', './office-translate.js', './a11y.js', './admin-panel.js', './font-rubik-400.ttf', './font-rubik-700.ttf', './font-alef-400.ttf', './font-alef-700.ttf', './siman-phone.css', './siman-phone.js', './siman-voice.js', './siman-asr-worker.js', './transformers.min.js', './siman-back.js', './siman-v7.css', './siman-v7.js', './siman-scan.js', './siman-merge.js', './siman-ux.js', './siman-nb.css', './siman-nb.js', './siman-fill.js', './siman-pro.js', './siman-convert.js', './siman-home.js', './xl-charts.js'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  // only the app's own old versions are removed: the transcription engine (transformers-cache, ~165MB) stays
  e.waitUntil(caches.keys().then(keys => Promise.all(keys.filter(k => /^siman-v\d+$/.test(k) && k !== CACHE).map(k => caches.delete(k))))
    .then(() => self.clients.claim()));
});
// 5.9.6: a weak connection does not hold the app on a blank screen: after 6 seconds the saved copy is used
// 6.6.1: when the network is slower than that, the download is not thrown away: it still goes into the saved copy
// (before, a slow phone downloaded the new version on every launch and never kept it, so it stayed on the old one)
const fresh = (req, key, ev) => new Promise((res, rej) => {
  const p = fetch(req, { cache: 'no-cache' });
  const t = setTimeout(() => {
    rej(new Error('slow'));
    const late = p.then(r => { if (r.ok) return caches.open(CACHE).then(c => c.put(key || req, r)); }).catch(() => {});
    if (ev) try { ev.waitUntil(late); } catch (_) {}
  }, 6000);
  p.then(r => { clearTimeout(t); res(r); }, e => { clearTimeout(t); rej(e); });
});
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
    e.respondWith(fresh(req, req, e).then(r => { if (r.ok) { const copy = r.clone(); e.waitUntil(caches.open(CACHE).then(c => c.put(req, copy))); } return r; })
      .catch(() => caches.match(req, { ignoreSearch: true }).then(h => h || fetch(req))));
    return;
  }
  if (isPage) {
    // newest page when online, cached page when offline. Only the app itself is kept as the app page:
    // other pages (like sign.html, opened from a link) are kept under their own address
    const isApp = /\/(index\.html)?$/.test(path);
    e.respondWith(fresh(req, isApp ? './index.html' : req, e).then(r => { if (r.ok) { const copy = r.clone(); e.waitUntil(caches.open(CACHE).then(c => c.put(isApp ? './index.html' : req, copy))); } return r; })
      .catch(() => isApp ? caches.match('./index.html') : caches.match(req)));
    return;
  }
  e.respondWith(caches.match(req, { ignoreSearch: true }).then(hit => hit || fetch(req).then(r => {
    if (r.ok) { const copy = r.clone(); caches.open(CACHE).then(c => c.put(req, copy)); }
    return r;
  })));
});
