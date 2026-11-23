// סימן — service worker: works offline, picks up new versions automatically
const CACHE = 'siman-v147'; // 6.9.7 — filling forms: a box with a question at its top and room under it is a place to fill (found by the picture of the page), its text breaks into lines inside the box and gets smaller when it is long; table cells are named by the heading of their column; a blank whose label ended the line above belongs to that label; the fill window shows the whole question, the page number and where the place is on the page; a list of all the places by page; on a phone a filled place opens again with a tap; 6.9.6 — the text-recognition engine (Tesseract: scanned PDF, reading print on the device) is served from the app's own site (./ocr/) instead of cdn.jsdelivr.net, kept in its own store across updates, and works with no internet after the first use; 6.9.5 — מדף למסמך: pages the smart service could not take (daily share used up, busy, no internet) wait in a queue kept on the device (pictures, text read so far, corrections) and are read by themselves when the service is back, while the app is open; the kept reading is offered when מדף למסמך is opened, until it is saved or deleted; 6.9.4 — מדף למסמך: a printed page is read on the device itself (the scanned-PDF engine) when the smart service is out of its daily share, busy or unreachable, so the work does not stop; such a page is marked and can be read again by the service later; a page the device is not sure of (handwriting) still waits for the service, without stopping the other pages; 6.9.3 — the toolbar stays on one row in a narrower window (buttons tighten, then the row scrolls sideways); the landing page in the new look; 6.9.2 — the top of the home screen fits a normal screen without scrolling: the search line sits under the opening sentence, and the headline and the open-file circle shrink on low screens; 6.9.1 — dark mode: the file icon on the home screen shows its arrow; 6.9.0 — the "ink and water" look: watercolour washes, handwritten headings (Amatic SC), Assistant text, a marker that marks each tool's words on hover, two tool columns on phones; a design layer only (siman-ink.css / siman-ink.js), no action of the app changed; 6.8.2 — a student's pages are read first without the answer key (so nothing of the key is read into the answer), then checked against it; 6.8.1 — exams and works are checked from the pictures of the pages themselves (the handwriting is read while checking, with the questions and the answer key in view), and each student's details show what was read for every question; 6.8.0 — a full check of the app: the work and class checks (their own checking mode on the server, a changed points table re-checks the students, a retry after closing the app works, pages that failed to read are named), a recording is guarded until it is saved, text under a black box no longer stays in OCR pages, a filled form keeps its text when pages change, restored drafts keep signatures, safer links and signature pictures; 6.7.2 — checking a whole class: exam + answer key → points table → student after student in the background → grades file (Word / PDF / Excel) with a breakdown per student; the work check opens on phones; 6.7.1 — work check from photos: the student's pages (also handwriting) and the exam / rubric photographed, scanned or picked, read, gathered into one PDF and checked; 'gather only' makes one file; 6.7.0 — study: work checker (feedback, fixes, grade, against a task or rubric), a quiz from a recording, course folders in the transcript library with a quiz / summary / questions on the whole folder; 6.6.13 — saving into the open file when it changed on disk since it was opened (network drive, another program): tried once more, then the save window opens with the same name instead of "השמירה נכשלה"; 6.6.12 — marking / signing view of Word files: Hebrew stays right to left (alignment read the way Word reads it in a right-to-left paragraph, Hebrew size / bold / font), floating and older-style (VML) pictures in their place, tables wider than the page fitted to it, header page number 1; 6.6.11 — old Word files (.doc) open with their formatting: styles, fonts, bold / size / colour, alignment and indents, bullets, table widths, merged cells, shading and borders, page size and margins, header / footer with the logo, and photos (also those kept inside EMF / WMF); in the page view right-to-left tables keep their column order and the page count is filled in; 6.6.10 — an old Word file (.doc) opens again (the 'what to do with it' window broke on it: toUpperCase of undefined); 6.6.9 — phone share: the link choice (keeps the Word file) is offered again (its code was never loaded on phones); a browser that says yes to a Word / Excel / PowerPoint file and then refuses it (Samsung Internet) is remembered, and such files go to WhatsApp / mail as PDF at once; 6.6.8 — the check report shows what happened in the last share attempts (menu, WhatsApp opened or not); 6.6.7 — WhatsApp on a phone: if WhatsApp does not open after the share menu, the dialog stays and offers a link (opens the WhatsApp app directly on Android), the menu again, or a download; remembered for next time; 6.6.6 — line spacing: in Word the menu says where it goes ("לכל המסמך" ticked when the cursor is not in the text, before it went silently to the first paragraph), table cells that lost their paragraph get one back; PowerPoint has a line-spacing button too (saved as PowerPoint spacing); reading: no Arabic letters inside Hebrew words; 6.6.5 — PowerPoint: a slide that is one picture (NotebookLM, Canva, exported decks) can be turned into editable text: the words are found with their places, taken out of the picture and put back as real text boxes (one slide or all, three at a time, undo restores the picture); 6.6.4 — reading a page on a phone: the screen stays on, a line cut because the app went to the background is sent again when it comes back, and a cut line is no longer called "no internet" (nor stops the other pages); 6.6.3 — full test round (desktop + phone): the command search works again after a form is opened, a left-to-right Excel sheet opens at column A (on phones it opened at M-Q), page thumbnails also prepare sideways (phone row); 6.6.2 — big files: thumbnails drawn only when seen, pages asked for 40 at a time, form fields looked for meanwhile; Excel undo copies the sheet once per edit (not twice) and is capped by size; ביטול on long jobs (OCR of all pages, compress, sensitive-data search; also phone Back); send for signature asks for the email and checks the 15MB limit before the fields are placed, and goes on after sign-up; translations and first questions already answered are reused in the session; 6.6.1 — review fixes: the page is half the size (the Alef font is a file, not written in twice), a slow network still keeps the new version, the Office assistant gets the whole request, a scan no longer marks the open file as saved, translation/OCR never write into another file, OCR survives page changes and one bad page, pages copied once per file (smaller files), less memory per page, Excel sorting keeps formulas right and scrolls faster, drafts only when something changed; 6.6.0 — מדף למסמך: printed pages read whole in seconds (auto: handwriting found by the quick reading goes on to the careful one), two pages at a time, pictures prepared three at a time; 6.5.2 — Excel: the selection is painted like Excel (grey tint over the cell colour, green frame around every area, marked headings); 6.5.1 — Excel: Ctrl+click adds areas (A + E for a chart), drag across headings selects columns/rows, a chart from several areas; 6.5.0 — Excel charts: the file's charts are shown (editor, marking, PDF) and new charts are added from selected cells (Insert › Chart), moved, resized, edited, deleted, saved in Excel's format; 6.4.2 — right-click paste really offers הדבק (checked in real Edge); one transcription engine (measured: two were slower); 6.4.1 — forms: lines kept in pieces are joined (whole places, full labels, signature found, no footer lines); places offered only on real forms again (5.10.2); the 5.10 converter, copy-as-text, numbers and faster transcription back; right-click paste on the home screen; 6.4 — the smart services need a pass (Cloudflare Turnstile) and keep a fair daily share per device; your own Gemini key = no limit; 6.3.8 — Word: text never runs past the last page (no page cap), Excel grows rows/columns while scrolling; 6.3.7 — printing from the browser's own menu is not blank; 6.3.6 — the home card really hides; 6.3.5 — check report (Ctrl K); 6.3.4 — the home card never freezes; 6.3.3 — closing a file from the home card asks inside the card; 6.3.2 — resume card fixed (stuck files, × to close), editors no longer break when left while opening; 6.3.1 — new file inside the Office editors, the smart server does not wait for a stuck model; 6.3 — home from anywhere, Word forms fill, autofill and my details; 6.2.1 — Word forms: values written on their blank lines; 6.2.1 — anonymous device count (siman-stats), real app version; WhatsApp on phones: link when files can't be shared; 6.2 — no sign-up for local tools, Hebrew date, business stamp, in-person signing, glossary; 6.1 — auto redaction, PDF compression, ask all open files; 6.0 — karaoke transcript, stars, chapters, study tools, cite highlight, JPG/PNG/Markdown
const ASSETS = ['./', './index.html', './manifest.webmanifest', './pdf.min.js', './pdf.worker.min.js',
  './pdf-lib.min.js', './fontkit.umd.min.js', './icon-192.png', './icon-512.png', './icon-maskable-512.png', './jszip.min.js', './pptxgen.min.js', './office.js', './office-edit.js', './word-editor.js', './excel-editor.js', './docx-preview.min.js', './html2canvas.min.js', './xlsx.mini.min.js', './font-secular-he.woff2', './font-secular-lat.woff2', './send.js', './siman-plus.js', './siman-guard.js', './siman-media.js', './ppt-editor.js', './office-ai.js', './office-translate.js', './a11y.js', './admin-panel.js', './font-rubik-400.ttf', './font-rubik-700.ttf', './font-alef-400.ttf', './font-alef-700.ttf', './siman-phone.css', './siman-phone.js', './siman-voice.js', './siman-asr-worker.js', './transformers.min.js', './siman-back.js', './siman-v7.css', './siman-v7.js', './siman-scan.js', './siman-merge.js', './siman-ux.js', './siman-nb.css', './siman-nb.js', './siman-ink.css', './siman-ink.js', './font-amatic-he.woff2', './font-amatic-lat.woff2', './font-amatic-cyr.woff2', './font-assistant-he.woff2', './font-assistant-lat.woff2', './siman-fill.js', './siman-pro.js', './siman-convert.js', './siman-home.js', './xl-charts.js'];

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
  // 6.9.6: the text-recognition engine (./ocr/, ~20MB with all its languages) is kept in its own store, which
  // outlives app updates: downloaded once, file by file as each is first needed, then it works with no internet
  if (/\/ocr\/[^/]+$/.test(path)) {
    e.respondWith(caches.open('siman-ocr').then(c => c.match(req, { ignoreSearch: true }).then(hit => hit || fetch(req).then(r => {
      if (r.ok) { const copy = r.clone(); e.waitUntil(c.put(req, copy).catch(() => {})); }
      return r;
    }))));
    return;
  }
  const isPage = req.mode === 'navigate' || /\/(index\.html)?$/.test(path);
  // the app's own code: always the newest version when online, so an upload shows up on the next refresh
  const isAppCode = /\/(siman-guard|office|office-edit|word-editor|excel-editor|siman-register|send|siman-plus|ppt-editor|a11y|office-ai|office-translate|siman-voice|siman-asr-worker|siman-back|siman-scan|siman-merge|siman-ux|siman-fill|siman-convert|siman-pro|siman-home|admin-panel|siman-media|siman-ai|siman-share|xl-charts)\.js$/.test(path) || /\/siman-(phone|v7|nb)\.(css|js)$/.test(path);
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
