/* סימן — זיהוי קישורים, סרטונים, קישורים לאפליקציות וקבצים מצורפים בתוך PDF
   - קישורים רגילים (Link) וכפתורי טופס עם כתובת
   - כתובות שמופיעות כטקסט בלבד (https://… / www.… / youtu.be/…)
   - סרטונים ושמע שמוטמעים בקובץ (Screen / RichMedia / Movie / Rendition)
   - קבצים מצורפים (FileAttachment ו"קבצים מצורפים" של המסמך)
   - הפניות לקבצים שנמצאים מחוץ ל-PDF (מסבירים למשתמש למה אי אפשר לפתוח)
   - אותו דבר בקובצי Word, Excel ו-PowerPoint: בעורכים ובפתיחה לסימון
   YouTube / Vimeo / Google Drive / קובצי וידאו מתנגנים בתוך סימן, בחלון נגן. */
(function () {
  'use strict';

  const rawCache = new WeakMap();   // S.bytes -> Promise<items per page>
  const scanned = new WeakMap();    // S.doc   -> list | null
  const AUDIO_EXT = /^(mp3|wav|m4a|aac|oga|ogg|opus|flac)$/i;
  const VIDEO_EXT = /^(mp4|m4v|mov|webm|ogv|avi|wmv|flv|mkv|3gp|mpg|mpeg|swf|f4v)$/i;
  const PLAYABLE = /^(mp4|m4v|mov|webm|ogv|mp3|wav|m4a|aac|oga|ogg|opus|flac)$/i;
  const MIME = { mp4: 'video/mp4', m4v: 'video/mp4', mov: 'video/quicktime', webm: 'video/webm', ogv: 'video/ogg', mp3: 'audio/mpeg', wav: 'audio/wav', m4a: 'audio/mp4', aac: 'audio/aac', ogg: 'audio/ogg', oga: 'audio/ogg', opus: 'audio/ogg', flac: 'audio/flac', pdf: 'application/pdf' };
  const BLOCKED = /^(javascript|data|vbscript|file|blob|about):/i;

  const $$ = id => document.getElementById(id);
  const say = m => { try { toast(m); } catch (_) { } };
  const extOf = n => ((n || '').match(/\.([a-z0-9]{2,5})(?:[?#].*)?$/i) || [])[1] || '';
  const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; };
  const idle = () => new Promise(r => (window.requestIdleCallback || setTimeout)(r, 30));

  /* ---------------- סוג הקישור ---------------- */
  function parseT(t) {
    if (/^\d+$/.test(t)) return +t;
    const m = String(t).match(/(?:(\d+)h)?(?:(\d+)m)?(?:(\d+)s)?/);
    return m ? (+m[1] || 0) * 3600 + (+m[2] || 0) * 60 + (+m[3] || 0) : 0;
  }
  function embedOf(u) {
    let x; try { x = new URL(u); } catch (_) { return null; }
    const h = x.hostname.replace(/^(www|m|music)\./, '');
    const t = x.searchParams.get('t') || x.searchParams.get('start') || (x.hash.match(/t=([\dhms]+)/) || [])[1];
    const st = t ? parseT(t) : 0;
    let id = null;
    if (h === 'youtu.be') id = x.pathname.slice(1).split('/')[0];
    else if (/(^|\.)youtube(-nocookie)?\.com$/.test(h)) id = x.searchParams.get('v') || (x.pathname.match(/\/(embed|shorts|live|v)\/([\w-]{6,})/) || [])[2];
    if (id) return { type: 'frame', site: 'YouTube', src: `https://www.youtube-nocookie.com/embed/${id}?autoplay=1&rel=0&playsinline=1${st ? '&start=' + st : ''}` };
    if (/(^|\.)vimeo\.com$/.test(h)) { const m = x.pathname.match(/(\d{6,})/); if (m) return { type: 'frame', site: 'Vimeo', src: `https://player.vimeo.com/video/${m[1]}?autoplay=1` }; }
    if (h === 'drive.google.com') { const m = x.pathname.match(/\/file\/d\/([\w-]+)/); const d = m ? m[1] : x.searchParams.get('id'); if (d) return { type: 'frame', site: 'Google Drive', src: `https://drive.google.com/file/d/${d}/preview` }; }
    const e = extOf(x.pathname);
    if (PLAYABLE.test(e)) return { type: AUDIO_EXT.test(e) ? 'audio' : 'video', site: 'קובץ ' + e.toUpperCase(), src: u };
    return null;
  }
  function classify(u) {
    const s = u.toLowerCase();
    if (s.startsWith('mailto:')) return { k: 'web', t: 'שליחת אימייל', ic: '✉' };
    if (/^(tel|sms):/.test(s)) return { k: 'app', t: 'חיוג / הודעה', ic: '📞' };
    const emb = embedOf(u);
    if (emb) return { k: emb.type === 'audio' ? 'audio' : (emb.site === 'Google Drive' ? 'web' : 'video'), t: emb.site, ic: emb.site === 'Google Drive' ? '🗂' : '▶', emb };
    if (/(facebook\.com\/.*video|fb\.watch|tiktok\.com|instagram\.com\/(reel|tv|p)\/|dailymotion\.com|dai\.ly|twitch\.tv|kan\.org\.il\/.*(vod|video)|mako\.co\.il\/.*vod)/.test(s)) return { k: 'video', t: 'סרטון באתר', ic: '▶' };
    if (!/^https?:/.test(s)) return { k: 'app', t: 'פתיחה באפליקציה', ic: '📱' };
    if (/(wa\.me|whatsapp\.com)/.test(s)) return { k: 'app', t: 'WhatsApp', ic: '💬' };
    if (/(zoom\.us\/(j|my|w)|teams\.microsoft|teams\.live|meet\.google)/.test(s)) return { k: 'app', t: 'פגישה מקוונת', ic: '🎥' };
    if (/(play\.google\.com\/store|apps\.apple\.com|itunes\.apple\.com|apps\.microsoft\.com|microsoft\.com\/store)/.test(s)) return { k: 'app', t: 'חנות אפליקציות', ic: '📱' };
    if (/(t\.me\/|waze\.com|maps\.app\.goo\.gl|google\.[a-z.]+\/maps|spotify\.com|kahoot|wordwall|padlet|genial\.ly|canva\.com\/design)/.test(s)) return { k: 'app', t: 'אפליקציה / שירות', ic: '📱' };
    if (/(docs\.google|forms\.gle|sites\.google)/.test(s)) return { k: 'web', t: 'Google', ic: '🔗' };
    return { k: 'web', t: 'אתר', ic: '🔗' };
  }

  /* ---------------- קריאת הקובץ הגולמי (pdf-lib) ---------------- */
  function rawPages(bytes) {
    if (!bytes || !window.PDFLib) return Promise.resolve([]);
    if (bytes.length > 150 * 1024 * 1024) return Promise.resolve([]);
    if (rawCache.has(bytes)) return rawCache.get(bytes);
    const p = (async () => {
      const L = window.PDFLib;
      const { PDFDocument, PDFName, PDFDict, PDFArray, PDFRawStream, PDFString, PDFHexString, PDFNumber, decodePDFRawStream } = L;
      const doc = await PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false, throwOnInvalidObject: false });
      const ctx = doc.context, N = k => PDFName.of(k);
      const look = x => { try { return x ? ctx.lookup(x) : undefined; } catch (_) { return undefined; } };
      const nm = x => { x = look(x); return x instanceof PDFName ? x.decodeText() : ''; };
      const txt = x => { x = look(x); return (x instanceof PDFString || x instanceof PDFHexString) ? x.decodeText() : ''; };
      const num = x => { x = look(x); return x instanceof PDFNumber ? x.asNumber() : 0; };
      const sub = (d, k) => { const r = d && look(d.get(N(k))); return r instanceof PDFDict ? r : null; };
      const bytesOf = s => { if (s instanceof PDFRawStream) return decodePDFRawStream(s).decode(); return s && s.getContents ? s.getContents() : null; };

      const fspec = v => {
        const o = look(v);
        if (o instanceof PDFString || o instanceof PDFHexString) { const n = o.decodeText(); return /^https?:\/\//i.test(n) ? { url: n } : { name: n, outside: true }; }
        if (!(o instanceof PDFDict)) return null;
        const name = txt(o.get(N('UF'))) || txt(o.get(N('F'))) || txt(o.get(N('DOS'))) || txt(o.get(N('Unix'))) || '';
        if (nm(o.get(N('FS'))) === 'URL' || /^https?:\/\//i.test(name)) return { url: name };
        const ef = sub(o, 'EF');
        const st = ef && (look(ef.get(N('F'))) || look(ef.get(N('UF'))));
        if (st && st.dict) return { name: name || 'קובץ', get: () => bytesOf(st), mime: nm(st.dict.get(N('Subtype'))) };
        return { name, outside: true };
      };
      const fromRendition = (R, depth = 0) => {
        R = look(R); if (!(R instanceof PDFDict) || depth > 4) return null;
        if (nm(R.get(N('S'))) === 'SR') {
          const arr = look(R.get(N('R')));
          if (arr instanceof PDFArray) for (let i = 0; i < arr.size(); i++) { const r = fromRendition(arr.get(i), depth + 1); if (r) return r; }
          else return fromRendition(arr, depth + 1);
          return null;
        }
        const C = sub(R, 'C'); if (!C) return null;
        const f = fspec(C.get(N('D'))); if (!f) return null;
        f.mime = txt(C.get(N('CT'))) || f.mime;
        return f;
      };
      const fromAction = (A, depth = 0) => {
        A = look(A); if (!(A instanceof PDFDict) || depth > 4) return null;
        if (nm(A.get(N('S'))) === 'Rendition') {
          const r = fromRendition(A.get(N('R'))); if (r) return r;
          const an = look(A.get(N('AN'))); if (an instanceof PDFDict && an !== A) { const m = fromScreen(an, depth + 1); if (m) return m; }
        }
        const nx = look(A.get(N('Next')));
        if (nx instanceof PDFArray) { for (let i = 0; i < nx.size(); i++) { const r = fromAction(nx.get(i), depth + 1); if (r) return r; } }
        else if (nx) return fromAction(nx, depth + 1);
        return null;
      };
      const fromScreen = (a, depth = 0) => {
        const m = fromAction(a.get(N('A')), depth); if (m) return m;
        const aa = sub(a, 'AA');
        if (aa) for (const [, v] of aa.entries()) { const r = fromAction(v, depth); if (r) return r; }
        return null;
      };
      const fromRichMedia = a => {
        const rmc = sub(a, 'RichMediaContent'); const assets = sub(rmc, 'Assets'); if (!assets) return null;
        const list = [];
        const walk = (node, d) => {
          if (!node || d > 5) return;
          const names = look(node.get(N('Names')));
          if (names instanceof PDFArray) for (let i = 0; i + 1 < names.size(); i += 2) { const f = fspec(names.get(i + 1)); if (f) { if (!f.name) f.name = txt(names.get(i)); list.push(f); } }
          const kids = look(node.get(N('Kids')));
          if (kids instanceof PDFArray) for (let i = 0; i < kids.size(); i++) { const k = look(kids.get(i)); if (k instanceof PDFDict) walk(k, d + 1); }
        };
        walk(assets, 0);
        return list.find(f => PLAYABLE.test(extOf(f.name))) || list.find(f => VIDEO_EXT.test(extOf(f.name)) || AUDIO_EXT.test(extOf(f.name))) || list[0] || null;
      };

      const out = [];
      doc.getPages().forEach((pg, pi) => {
        const arr = []; out[pi] = arr;
        let an; try { an = pg.node.Annots(); } catch (_) { }
        if (!an) return;
        for (let k = 0; k < an.size(); k++) {
          try {
            const a = look(an.get(k)); if (!(a instanceof PDFDict)) continue;
            const st = nm(a.get(N('Subtype')));
            const R = look(a.get(N('Rect'))); if (!(R instanceof PDFArray) || R.size() < 4) continue;
            const rect = [0, 1, 2, 3].map(i => num(R.get(i)));
            let m = null;
            if (st === 'Screen') m = fromScreen(a) || { media: true, outside: true };
            else if (st === 'RichMedia') m = fromRichMedia(a) || { media: true, outside: true };
            else if (st === 'Movie') { const M = sub(a, 'Movie'); m = (M && fspec(M.get(N('F')))) || { media: true, outside: true }; m.media = true; }
            else if (st === 'Sound') m = { media: true, outside: true, name: 'קטע שמע ישן' };
            else if (st === 'FileAttachment') { m = fspec(a.get(N('FS'))); if (m) m.attach = true; }
            else if (st === 'Link') { const A = look(a.get(N('A'))); if (A instanceof PDFDict && nm(A.get(N('S'))) === 'Rendition') { m = fromAction(A) || { media: true, outside: true }; m.media = true; } }
            else if (st === 'Widget') {
              const A = look(a.get(N('A')));
              if (A instanceof PDFDict && nm(A.get(N('S'))) === 'URI') { const u = txt(A.get(N('URI'))); if (u) m = { url: u }; }
            }
            if (!m) continue;
            if (m.media && m.outside && !m.name) m.name = '';
            arr.push(toItem(m, rect, st));
          } catch (e) { /* annotation פגומה — מדלגים */ }
        }
      });
      return out;
    })().catch(e => { console.warn('siman-media raw', e); return []; });
    rawCache.set(bytes, p);
    return p;
  }

  function toItem(m, rect, src) {
    if (m.url) { const c = classify(m.url); return { rect, k: c.k, url: m.url, c, src }; }
    if (m.get) {
      const e = extOf(m.name).toLowerCase();
      let mime = m.mime || MIME[e] || '';
      const isA = /^audio\//.test(mime) || AUDIO_EXT.test(e);
      const isV = /^video\//.test(mime) || VIDEO_EXT.test(e) || (m.media && !isA);
      if (!mime && isV) mime = 'video/mp4';
      return { rect, k: isA ? 'audio' : isV ? 'video' : 'file', embedded: true, name: m.name, mime, get: m.get, src };
    }
    return { rect, k: 'outside', name: m.name || '', media: !!m.media, src };
  }

  /* ---------------- כתובות שמופיעות רק כטקסט ---------------- */
  const URL_RE = /(?:https?:\/\/|www\.)[A-Za-z0-9\-._~:/?#\[\]@!$&'()*+,;=%]+|\b(?:youtu\.be|wa\.me|forms\.gle|bit\.ly|tinyurl\.com|t\.me|did\.li|meet\.google\.com|zoom\.us\/j)\/[A-Za-z0-9\-._~/?#=&%]+/g;
  function textUrls(tc) {
    const found = [];
    let run = null;
    const flush = () => {
      if (!run || !run.s) { run = null; return; }
      URL_RE.lastIndex = 0;
      let m;
      while ((m = URL_RE.exec(run.s))) {
        let u = m[0].replace(/[.,;:!?)\]'"]+$/, '');
        if (u.length < 9 || /^www\.?$/i.test(u)) continue;
        const a = m.index, b = a + u.length;
        let r = null;
        for (const g of run.segs) {
          const s0 = Math.max(a, g.start), s1 = Math.min(b, g.end); if (s1 <= s0) continue;
          const len = Math.max(1, g.end - g.start);
          const x0 = g.e + g.w * (s0 - g.start) / len, x1 = g.e + g.w * (s1 - g.start) / len;
          const rr = [Math.min(x0, x1), g.f - g.h * .25, Math.max(x0, x1), g.f + g.h * .95];
          r = r ? [Math.min(r[0], rr[0]), Math.min(r[1], rr[1]), Math.max(r[2], rr[2]), Math.max(r[3], rr[3])] : rr;
        }
        if (!r) continue;
        const href = /^[a-z]+:/i.test(u) ? u : 'https://' + u;
        const c = classify(href);
        found.push({ rect: r, k: c.k, url: href, c, src: 'text' });
      }
      run = null;
    };
    for (const it of tc.items) {
      if (typeof it.str !== 'string') continue;
      const [a, b, , d, e, f] = it.transform;
      if (!it.str || Math.abs(b) > Math.abs(a) * .2 + .01) { if (it.hasEOL) flush(); continue; }
      const h = it.height || Math.abs(d) || 10;
      if (run && Math.abs(f - run.f) < h * .35) {
        const gap = e - run.endX;
        if (gap > h * .25 || gap < -h * .6) run.s += ' ';
      } else { flush(); run = { s: '', segs: [], f, endX: e }; }
      const start = run.s.length; run.s += it.str;
      run.segs.push({ start, end: run.s.length, e, w: it.width, f, h });
      run.endX = e + it.width;
      if (it.hasEOL) flush();
    }
    flush();
    return found;
  }

  const overlap = (p, q) => !(p[2] < q[0] || q[2] < p[0] || p[3] < q[1] || q[3] < p[1]);
  const norm = r => [Math.min(r[0], r[2]), Math.min(r[1], r[3]), Math.max(r[0], r[2]), Math.max(r[1], r[3])];

  /* ---------------- כל מה שיש בעמוד ---------------- */
  function pageItems(P) {
    const doc = S.doc, bytes = S.bytes;
    if (P._sm && P._sm.doc === doc) return P._sm.items;
    const items = (async () => {
      const out = [];
      let anns = [];
      try { anns = await P.page.getAnnotations({ intent: 'display' }); } catch (_) { }
      for (const a of anns) {
        if (a.annotationType !== 2 || !a.rect) continue;
        const rect = norm(a.rect);
        if (a.url) { const c = classify(a.url); out.push({ rect, k: c.k, url: a.url, c, src: 'Link' }); }
        else if (a.attachment) { const f = a.attachment; const e = extOf(f.filename); out.push({ rect, k: AUDIO_EXT.test(e) ? 'audio' : VIDEO_EXT.test(e) ? 'video' : 'file', embedded: true, name: f.filename, mime: MIME[e.toLowerCase()] || '', get: () => f.content, src: 'Link' }); }
        else if (a.dest) out.push({ rect, k: 'dest', dest: a.dest, src: 'Link' });
        else if (a.action) out.push({ rect, k: 'dest', named: a.action, src: 'Link' });
        else if (a.unsafeUrl && !BLOCKED.test(a.unsafeUrl)) {
          if (/^[a-z][a-z0-9+.-]{1,20}:/i.test(a.unsafeUrl)) { const c = classify(a.unsafeUrl); out.push({ rect, k: c.k, url: a.unsafeUrl, c, src: 'Link' }); }
          else out.push({ rect, k: 'outside', name: a.unsafeUrl, media: VIDEO_EXT.test(extOf(a.unsafeUrl)) || AUDIO_EXT.test(extOf(a.unsafeUrl)), src: 'Link' });
        }
      }
      try { const raw = await rawPages(bytes); for (const it of (raw[P.i - 1] || [])) out.push({ ...it, rect: norm(it.rect) }); } catch (_) { }
      try {
        const tc = await P.page.getTextContent();
        for (const it of textUrls(tc)) if (!out.some(o => overlap(o.rect, it.rect))) out.push(it);
      } catch (_) { }
      return out;
    })();
    P._sm = { doc, items };
    return items;
  }

  /* ---------------- שכבת הקישורים על העמוד ---------------- */
  async function render(P, z) {
    if (!P || !P.page || !S.doc) return;
    const doc = S.doc;
    let items; try { items = await pageItems(P); } catch (_) { return; }
    if (S.doc !== doc || S.zoom !== z || !P.el || !P.el.isConnected) return;
    const vp = P.page.getViewport({ scale: z });
    let L = P.el.querySelector(':scope > .sm-links');
    if (!L) { L = el('div', 'sm-links'); P.el.append(L); }
    L.textContent = '';
    for (const it of items) {
      const r = vp.convertToViewportRectangle(it.rect);
      const x = Math.min(r[0], r[2]), y = Math.min(r[1], r[3]), w = Math.abs(r[2] - r[0]), h = Math.abs(r[3] - r[1]);
      if (w < 3 || h < 3) continue;
      const b = el('button', 'sm-l sm-' + it.k + (it.embedded ? ' sm-emb' : '') + (it.src === 'text' ? ' sm-txt' : ''));
      b.type = 'button';
      Object.assign(b.style, { left: x + 'px', top: y + 'px', width: w + 'px', height: h + 'px' });
      const label = labelOf(it);
      b.title = label; b.setAttribute('aria-label', label);
      if (it.k === 'video' || it.k === 'audio') {
        const big = it.embedded || it.src === 'Screen' || it.src === 'RichMedia' || it.src === 'Movie';
        const badge = el('span', big && w > 60 && h > 40 ? 'sm-play sm-big' : 'sm-play', '▶'); b.append(badge);
      } else if (it.k === 'file' || (it.k === 'outside')) b.append(el('span', 'sm-play sm-clip', it.k === 'outside' ? '!' : '📎'));
      b.addEventListener('click', e => { e.preventDefault(); e.stopPropagation(); act(it, e); });
      L.append(b);
    }
    if (!scanned.has(doc)) scan(doc);
  }

  function labelOf(it) {
    if (it.k === 'dest') return 'מעבר לעמוד במסמך';
    if (it.url) return (it.c ? it.c.t + ': ' : '') + it.url;
    if (it.k === 'outside') return it.media ? 'סרטון שלא נשמר בתוך הקובץ' : 'הפניה לקובץ שמחוץ ל-PDF';
    if (it.k === 'video') return 'סרטון בתוך הקובץ' + (it.name ? ': ' + it.name : '');
    if (it.k === 'audio') return 'שמע בתוך הקובץ' + (it.name ? ': ' + it.name : '');
    return 'קובץ מצורף' + (it.name ? ': ' + it.name : '');
  }

  /* ---------------- פעולה בלחיצה ---------------- */
  async function goTo(it) {
    let idx = null;
    try {
      if (it.named) {
        const cur = currentPage();
        idx = { NextPage: cur + 1, PrevPage: cur - 1, FirstPage: 0, LastPage: S.pages.length - 1 }[it.named];
      } else {
        let d = it.dest; if (typeof d === 'string') d = await S.doc.getDestination(d);
        if (Array.isArray(d)) idx = typeof d[0] === 'number' ? d[0] : await S.doc.getPageIndex(d[0]);
      }
    } catch (_) { }
    if (idx == null || !S.pages[idx]) return;
    S.pages[idx].el.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
  function currentPage() {
    const d = $$('desk'); if (!d) return 0;
    const mid = d.getBoundingClientRect().top + d.clientHeight / 2;
    let best = 0, bd = 1e9;
    S.pages.forEach((P, k) => { const r = P.el.getBoundingClientRect(); const dd = Math.abs((r.top + r.bottom) / 2 - mid); if (dd < bd) { bd = dd; best = k; } });
    return best;
  }

  function act(it, e) {
    if (it.k === 'dest') return goTo(it);
    if (it.url) {
      if (BLOCKED.test(it.url)) return;
      if (it.c && it.c.emb) return play({ title: it.c.t, emb: it.c.emb, url: it.url });
      if (e && (e.ctrlKey || e.metaKey)) return openUrl(it.url);
      return linkSheet(it);
    }
    if (it.embedded) return embeddedSheet(it);
    return outsideSheet(it);
  }

  function openUrl(u) {
    if (BLOCKED.test(u)) return;
    if (/^https?:/i.test(u)) { window.open(u, '_blank', 'noopener'); return; }
    const a = el('a'); a.href = u; a.rel = 'noopener'; document.body.append(a); a.click(); a.remove();
    setTimeout(() => { if (document.visibilityState === 'visible') say('אם שום דבר לא נפתח — האפליקציה כנראה לא מותקנת במכשיר הזה'); }, 1600);
  }
  function copy(t) {
    (navigator.clipboard ? navigator.clipboard.writeText(t) : Promise.reject()).then(() => say('הקישור הועתק'), () => say('לא הצלחתי להעתיק'));
  }
  function saveBytes(bytes, name, mime) {
    const url = URL.createObjectURL(new Blob([bytes], { type: mime || 'application/octet-stream' }));
    const a = el('a'); a.href = url; a.download = name || 'קובץ'; document.body.append(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }
  function fmtSize(n) { return n > 1048576 ? (n / 1048576).toFixed(1) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB'; }

  /* ---------------- חלונות ---------------- */
  let ov = null, ovUrl = null;
  function open(title, wide) {
    close();
    ov = el('div', 'sm-ov'); ov.setAttribute('role', 'dialog'); ov.setAttribute('aria-modal', 'true'); ov.dir = 'rtl';
    const box = el('div', 'sm-box' + (wide ? ' sm-wide' : ''));
    const head = el('div', 'sm-head');
    const h = el('div', 'sm-title', title);
    const x = el('button', 'sm-x', '✕'); x.type = 'button'; x.setAttribute('aria-label', 'סגירה'); x.onclick = close;
    head.append(h, x);
    const body = el('div', 'sm-body'), foot = el('div', 'sm-foot');
    box.append(head, body, foot); ov.append(box);
    ov.addEventListener('click', e => { if (e.target === ov) close(); });
    document.body.append(ov);
    setTimeout(() => x.focus(), 30);
    return { body, foot };
  }
  function close() {
    if (ov) { ov.remove(); ov = null; }
    if (ovUrl) { URL.revokeObjectURL(ovUrl); ovUrl = null; }
  }
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && ov) close(); });
  function btn(text, cls, fn) { const b = el('button', 'sm-btn ' + (cls || ''), text); b.type = 'button'; b.onclick = fn; return b; }
  function urlLine(u) { const d = el('div', 'sm-url', u); d.dir = 'ltr'; return d; }

  function play({ title, emb, url, blobUrl, name, bytes, mime }) {
    const { body, foot } = open(title || 'סרטון', true);
    const stage = el('div', 'sm-stage' + (emb && emb.type === 'audio' || /^audio/.test(mime || '') ? ' sm-audio' : ''));
    body.append(stage);
    if (emb && emb.type === 'frame') {
      const f = el('iframe'); f.src = emb.src;
      f.allow = 'autoplay; encrypted-media; picture-in-picture; fullscreen'; f.allowFullscreen = true;
      f.referrerPolicy = 'strict-origin-when-cross-origin'; f.title = title || 'סרטון';
      stage.append(f);
    } else {
      const isA = (emb && emb.type === 'audio') || /^audio/.test(mime || '');
      const v = el(isA ? 'audio' : 'video'); v.controls = true; v.autoplay = true; v.playsInline = true;
      v.src = blobUrl || (emb && emb.src);
      v.onerror = () => {
        stage.textContent = '';
        stage.classList.add('sm-err');
        stage.append(el('div', 'sm-msg', `הדפדפן לא יודע לנגן את סוג הקובץ הזה${name ? ' (' + (extOf(name).toUpperCase() || '?') + ')' : ''}. אפשר לשמור אותו ולפתוח בנגן של המכשיר.`));
      };
      stage.append(v);
    }
    if (bytes) foot.append(btn('שמירה למכשיר', 'sm-pri', () => saveBytes(bytes, name, mime)));
    if (url) {
      foot.append(btn(emb && emb.type === 'frame' ? 'פתיחה ב-' + emb.site : 'פתיחה בדפדפן', bytes ? '' : 'sm-pri', () => openUrl(url)));
      foot.append(btn('העתקת קישור', '', () => copy(url)));
    }
  }

  function linkSheet(it) {
    const c = it.c || classify(it.url);
    const { body, foot } = open(c.ic + ' ' + (c.k === 'app' ? c.t : c.k === 'video' ? 'סרטון' : 'קישור') );
    body.append(el('div', 'sm-msg', c.k === 'app' ? 'הקישור פותח אפליקציה או שירות חיצוני:' : c.k === 'video' ? 'הסרטון נמצא באתר חיצוני ונפתח שם:' : 'הקישור מוביל אל:'));
    body.append(urlLine(it.url));
    foot.append(btn(c.k === 'app' ? 'פתיחה באפליקציה' : 'פתיחה', 'sm-pri', () => { openUrl(it.url); close(); }));
    foot.append(btn('העתקה', '', () => copy(it.url)));
  }

  async function embeddedSheet(it) {
    let bytes;
    try { bytes = await it.get(); } catch (e) { console.warn(e); }
    if (!bytes || !bytes.length) { say('לא הצלחתי לקרוא את הקובץ שבתוך המסמך'); return; }
    const e = extOf(it.name).toLowerCase();
    const mime = it.mime || MIME[e] || '';
    if ((it.k === 'video' || it.k === 'audio') && (!e || PLAYABLE.test(e) || /^(video|audio)\//.test(mime))) {
      ovUrl = null;
      const blobUrl = URL.createObjectURL(new Blob([bytes], { type: mime || (it.k === 'audio' ? 'audio/mpeg' : 'video/mp4') }));
      play({ title: it.name || (it.k === 'audio' ? 'שמע' : 'סרטון'), blobUrl, name: it.name || ('video.' + (e || 'mp4')), bytes, mime });
      ovUrl = blobUrl;
      return;
    }
    const { body, foot } = open(it.k === 'video' ? '▶ סרטון בתוך הקובץ' : '📎 קובץ מצורף');
    const n = el('div', 'sm-file'); const nn = el('b', '', it.name || 'קובץ'); nn.dir = 'auto';
    n.append(nn, el('span', '', ' · ' + fmtSize(bytes.length)));
    body.append(n);
    if (it.k === 'video') body.append(el('div', 'sm-msg', 'סוג הסרטון הזה לא מתנגן בדפדפן. שומרים אותו ופותחים בנגן של המכשיר.'));
    if (e === 'pdf' && typeof openFile === 'function') foot.append(btn('פתיחה בסימן', 'sm-pri', async () => {
      close();
      try { if (S.dirty && !confirm('יש שינויים שלא נשמרו במסמך הנוכחי. לפתוח בכל זאת?')) return; } catch (_) { }
      openFile(new File([bytes], it.name || 'מצורף.pdf', { type: 'application/pdf' }));
    }));
    foot.append(btn('שמירה למכשיר', e === 'pdf' ? '' : 'sm-pri', () => saveBytes(bytes, it.name, mime)));
  }

  function outsideSheet(it) {
    const { body } = open(it.media ? '▶ הסרטון לא נמצא בקובץ' : '📄 קובץ חיצוני');
    if (it.name) { const n = el('div', 'sm-url', it.name); n.dir = 'auto'; body.append(n); }
    body.append(el('div', 'sm-msg', it.media
      ? 'במקום הזה במסמך היה סרטון, אבל הוא לא נשמר בתוך ה-PDF (כך קורה למשל כשמייצאים מצגת עם סרטון ל-PDF). אפשר לבקש מהשולח את הסרטון עצמו או קישור אליו.'
      : 'המסמך מפנה לקובץ ששמור במחשב של מי שיצר אותו, ולא נשלח יחד עם ה-PDF. כדי לפתוח אותו צריך לבקש מהשולח את הקובץ עצמו.'));
  }

  /* ---------------- סריקת המסמך והודעה למשתמש ---------------- */
  async function scan(doc) {
    scanned.set(doc, null);
    const old = document.querySelector('.sm-bar'); if (old) old.remove();
    const list = [];
    const n = Math.min(S.pages.length, 300);
    for (let i = 0; i < n; i++) {
      if (S.doc !== doc) return;
      const P = S.pages[i]; if (!P) break;
      let items = []; try { items = await pageItems(P); } catch (_) { }
      for (const it of items) if (it.k !== 'dest') list.push({ ...it, page: i });
      if (i % 8 === 7) await idle();
    }
    try {
      const at = await doc.getAttachments();
      if (at) for (const k of Object.keys(at)) {
        const f = at[k]; const e = extOf(f.filename);
        list.push({ k: AUDIO_EXT.test(e) ? 'audio' : VIDEO_EXT.test(e) ? 'video' : 'file', embedded: true, name: f.filename, mime: MIME[e.toLowerCase()] || '', get: () => f.content, page: -1 });
      }
    } catch (_) { }
    // קובץ Word / Excel / PowerPoint שנפתח לסימון: מה שיש בקובץ המקורי
    try {
      const src = S.office && S.office.bytes ? { bytes: S.office.bytes, kind: S.office.kind } : (window.SM_SRC && window.SM_SRC.name === S.name ? window.SM_SRC : null);
      if (src) for (const it of await officeItems(src.bytes, src.kind)) if (!(it.url && list.some(x => x.url === it.url))) list.push(it);
    } catch (_) { }
    if (S.doc !== doc) return;
    // אותה כתובת בכמה מקומות נספרת פעם אחת
    const seen = new Set(), uniq = [];
    for (const it of list) { const key = it.url || (it.name + '|' + it.page + '|' + (it.rect || '')); if (it.url && seen.has(key)) continue; seen.add(key); uniq.push(it); }
    scanned.set(doc, uniq);
    showBar(uniq);
  }

  function counts(list) {
    const c = { video: 0, link: 0, app: 0, file: 0, outside: 0 };
    for (const it of list) {
      if (it.k === 'video' || it.k === 'audio') c.video++;
      else if (it.k === 'app') c.app++;
      else if (it.k === 'file') c.file++;
      else if (it.k === 'outside') c.outside++;
      else c.link++;
    }
    return c;
  }

  function showBar(list, float) {
    const old = document.querySelector(float ? '.sm-bar.sm-float' : '.sm-bar:not(.sm-float)'); if (old) old.remove();
    if (!list.length) return;
    const desk = $$('desk'); if (!desk && !float) return;
    const c = counts(list);
    const parts = [];
    if (c.video) parts.push(`▶ ${c.video} ${c.video === 1 ? 'סרטון' : 'סרטונים'}`);
    if (c.app) parts.push(`📱 ${c.app} ${c.app === 1 ? 'קישור לאפליקציה' : 'קישורים לאפליקציות'}`);
    if (c.link) parts.push(`🔗 ${c.link} ${c.link === 1 ? 'קישור' : 'קישורים'}`);
    if (c.file) parts.push(`📎 ${c.file} ${c.file === 1 ? 'קובץ מצורף' : 'קבצים מצורפים'}`);
    if (c.outside) parts.push(`! ${c.outside} לא בתוך הקובץ`);
    const wrap = el('div', 'sm-bar' + (float ? ' sm-float' : '')); wrap.dir = 'rtl';
    const pill = el('div', 'sm-pill');
    const t = el('button', 'sm-pill-t'); t.type = 'button';
    t.append(el('span', 'sm-pill-h', float ? 'נמצא בקובץ' : 'נמצא במסמך'), el('span', 'sm-pill-p', parts.join(' · ')));
    t.onclick = () => listSheet(list);
    const x = el('button', 'sm-pill-x', '✕'); x.type = 'button'; x.setAttribute('aria-label', 'הסתרה');
    x.onclick = e => { e.stopPropagation(); pill.classList.add('sm-mini'); };
    pill.append(t, x); wrap.append(pill);
    if (float) document.body.append(wrap); else desk.prepend(wrap);
    setTimeout(() => pill.classList.add('sm-mini'), 12000);
    pill.addEventListener('click', () => { if (pill.classList.contains('sm-mini')) { pill.classList.remove('sm-mini'); listSheet(list); } });
  }

  function listSheet(list) {
    const { body } = open('מה יש בתוך הקובץ', false);
    const order = { video: 0, audio: 0, app: 1, web: 2, file: 3, outside: 4 };
    const sorted = [...list].sort((a, b) => (order[a.k] ?? 5) - (order[b.k] ?? 5) || (a.page ?? -1) - (b.page ?? -1));
    const ul = el('div', 'sm-list');
    for (const it of sorted) {
      const row = el('button', 'sm-row'); row.type = 'button';
      const ic = it.k === 'video' || it.k === 'audio' ? '▶' : it.k === 'file' ? '📎' : it.k === 'outside' ? '!' : (it.c && it.c.ic) || '🔗';
      const i = el('span', 'sm-ri sm-' + it.k, ic);
      const mid = el('span', 'sm-rm');
      let main = it.url ? (it.c ? it.c.t : 'קישור') : it.k === 'outside' ? (it.media ? 'סרטון שלא נשמר בקובץ' : 'קובץ חיצוני') : (it.name || 'קובץ');
      const sub = it.url ? it.url.replace(/^https?:\/\/(www\.)?/, '') : it.k === 'outside' ? (it.name || '') : (it.k === 'video' ? 'סרטון בתוך הקובץ' : it.k === 'audio' ? 'שמע בתוך הקובץ' : 'קובץ בתוך הקובץ');
      const a = el('b', '', main); a.dir = 'auto';
      const s = el('small', '', sub); s.dir = it.url ? 'ltr' : 'auto';
      mid.append(a, s);
      const pg = el('span', 'sm-rp', it.loc || (it.page >= 0 ? 'עמ׳ ' + (it.page + 1) : 'מצורף'));
      row.append(i, mid, pg);
      row.onclick = () => {
        close();
        const P = it.page >= 0 && S.pages && S.pages[it.page];
        if (P) P.el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        setTimeout(() => act(it), P ? 350 : 0);
      };
      ul.append(row);
    }
    body.append(ul);
  }


  /* ---------------- קובצי Word / Excel / PowerPoint ---------------- */
  const offCache = new WeakMap();
  function needZip() {
    if (window.JSZip) return Promise.resolve(window.JSZip);
    return new Promise((res, rej) => { const sc = document.createElement('script'); sc.src = 'jszip.min.js'; sc.onload = () => res(window.JSZip); sc.onerror = rej; document.head.append(sc); });
  }
  const unxml = t => t.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (m, n) => String.fromCharCode(+n)).replace(/&#x([0-9a-f]+);/gi, (m, n) => String.fromCharCode(parseInt(n, 16))).replace(/&amp;/g, '&');
  function joinPath(dir, target) {
    if (target.startsWith('/')) return target.slice(1);
    const parts = (dir ? dir.split('/') : []).concat(target.split('/'));
    const out = [];
    for (const x of parts) { if (!x || x === '.') continue; if (x === '..') out.pop(); else out.push(x); }
    return out.join('/');
  }
  function officeItems(bytes, kind) {
    if (!bytes || !bytes.length) return Promise.resolve([]);
    if (offCache.has(bytes)) return offCache.get(bytes);
    const p = (async () => {
      const JSZip = await needZip();
      const zip = await JSZip.loadAsync(bytes);
      const files = Object.keys(zip.files).filter(n => !zip.files[n].dir);
      const read = n => zip.file(n) ? zip.file(n).async('string') : Promise.resolve('');
      // שמות הגיליונות ב-Excel
      const sheetName = {};
      if (files.includes('xl/workbook.xml')) {
        const wb = await read('xl/workbook.xml'), wr = await read('xl/_rels/workbook.xml.rels');
        const rid = {};
        for (const m of wr.matchAll(/<Relationship\b[^>]*>/g)) { const id = (m[0].match(/\bId="([^"]+)"/) || [])[1], t = (m[0].match(/\bTarget="([^"]+)"/) || [])[1]; if (id && t) rid[id] = joinPath('xl', t); }
        for (const m of wb.matchAll(/<sheet\b[^>]*>/g)) { const nm = (m[0].match(/\bname="([^"]*)"/) || [])[1], id = (m[0].match(/\br:id="([^"]+)"/) || [])[1]; if (id && rid[id]) sheetName[rid[id]] = unxml(nm || ''); }
      }
      const where = part => {
        let m;
        if ((m = part.match(/^ppt\/slides\/slide(\d+)\.xml$/))) return 'שקופית ' + m[1];
        if ((m = part.match(/^ppt\/notesSlides\/notesSlide(\d+)\.xml$/))) return 'הערות שקופית ' + m[1];
        if (/^ppt\/(slideLayouts|slideMasters)\//.test(part)) return 'תבנית המצגת';
        if (/^xl\/worksheets\//.test(part)) return sheetName[part] ? 'גיליון ' + sheetName[part] : 'בגיליון';
        if (/^xl\/(sharedStrings|drawings)/.test(part)) return 'בגיליון';
        if (/^word\/header/.test(part)) return 'בכותרת העליונה';
        if (/^word\/footer/.test(part)) return 'בכותרת התחתונה';
        if (/^word\/(footnotes|endnotes)/.test(part)) return 'בהערות שוליים';
        if (/^word\//.test(part)) return 'במסמך';
        return '';
      };
      const order = s => { const m = s.match(/(\d+)/); return m ? +m[1] : 0; };
      const out = [], seenUrl = new Set(), seenFile = new Set();
      const addUrl = (u, loc) => {
        u = u.trim(); if (!u || BLOCKED.test(u) || /^#/.test(u)) return;
        if (/^www\./i.test(u)) u = 'https://' + u;
        if (!/^[a-z][a-z0-9+.-]{1,20}:/i.test(u)) { out.push({ k: 'outside', name: u, media: VIDEO_EXT.test(extOf(u)) || AUDIO_EXT.test(extOf(u)), loc, page: -1 }); return; }
        const key = u + '|' + loc; if (seenUrl.has(key)) return; seenUrl.add(key);
        const c = classify(u); out.push({ k: c.k, url: u, c, loc, page: -1 });
      };
      const addFile = (path, loc, media) => {
        if (seenFile.has(path) || !zip.file(path)) return; seenFile.add(path);
        const name = path.split('/').pop(), e = extOf(name).toLowerCase();
        if (e === 'bin') return; // אובייקט OLE פנימי
        const isA = AUDIO_EXT.test(e), isV = VIDEO_EXT.test(e) || (media && !isA);
        out.push({ k: isA ? 'audio' : isV ? 'video' : 'file', embedded: true, name, mime: MIME[e] || (isV ? 'video/mp4' : ''), get: () => zip.file(path).async('uint8array'), loc, page: -1 });
      };
      // קשרים (rels): קישורים, סרטונים, שמע, קבצים מוטמעים
      const rels = files.filter(n => /_rels\/[^/]+\.rels$/.test(n) && !/^_rels\//.test(n)).sort((a, b) => order(a) - order(b));
      for (const r of rels) {
        const dir = r.replace(/\/?_rels\/[^/]+\.rels$/, ''), part = joinPath(dir, r.split('/').pop().replace(/\.rels$/, ''));
        const loc = where(part), xml = await read(r);
        for (const m of xml.matchAll(/<Relationship\b[^>]*>/g)) {
          const tag = m[0];
          const type = (tag.match(/\bType="([^"]+)"/) || [])[1] || '', target = unxml((tag.match(/\bTarget="([^"]*)"/) || [])[1] || '');
          const ext = /TargetMode="External"/.test(tag);
          const media = /\/(video|audio|media)$/.test(type);
          if (/\/hyperlink$/.test(type)) { if (ext) addUrl(target, loc); }
          else if (media) { if (ext) addUrl(target, loc); else addFile(joinPath(dir, target), loc, true); }
          else if (/\/(package|oleObject)$/.test(type) && !ext) addFile(joinPath(dir, target), loc, false);
        }
      }
      // סרטון מקוון ב-Word ו-Excel, וכתובות שכתובות כטקסט בלבד
      const textParts = files.filter(n => /^(word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml|ppt\/slides\/slide\d+\.xml|ppt\/notesSlides\/notesSlide\d+\.xml|xl\/sharedStrings\.xml|xl\/worksheets\/sheet\d+\.xml)$/.test(n)).sort((a, b) => order(a) - order(b));
      for (const part of textParts) {
        const xml = await read(part), loc = where(part);
        for (const m of xml.matchAll(/embeddedHtml="([^"]+)"/g)) { const src = (unxml(m[1]).match(/src="([^"]+)"/) || [])[1]; if (src) addUrl(src.startsWith('//') ? 'https:' + src : src, loc); }
        const text = unxml(xml.replace(/<\/(w:p|a:p|si|w:tab|c|is|w:tc)>|<w:br\/>|<a:br\/>/g, '\n').replace(/<[^>]+>/g, ''));
        URL_RE.lastIndex = 0; let m;
        while ((m = URL_RE.exec(text))) { const u = m[0].replace(/[.,;:!?)\]'"]+$/, ''); if (u.length >= 9) addUrl(/^[a-z]+:/i.test(u) ? u : 'https://' + u, loc); }
      }
      // אותה כתובת שמופיעה גם כקישור וגם כטקסט באותו מקום נספרת פעם אחת (כבר מטופל ב-seenUrl)
      return out;
    })().catch(e => { console.warn('siman-media office', e); return []; });
    offCache.set(bytes, p);
    return p;
  }

  // בתוך עורכי Word / Excel / PowerPoint
  let floatTimer = null;
  async function office(bytes, kind) {
    const old = document.querySelector('.sm-bar.sm-float'); if (old) old.remove();
    clearInterval(floatTimer);
    let list = []; try { list = await officeItems(bytes, kind); } catch (_) { }
    if (!list.length) return;
    const api = { docx: window.WORD_EDITOR, xlsx: window.XLSX_EDITOR, pptx: window.PPTX_EDITOR }[kind];
    if (api && api.isOpen && !api.isOpen()) return;
    showBar(list, true);
    floatTimer = setInterval(() => { if (!api || !api.isOpen || api.isOpen()) return; const b = document.querySelector('.sm-bar.sm-float'); if (b) b.remove(); clearInterval(floatTimer); }, 1500);
  }

  // לחיצה על קישור בתוך העורכים: ב-Word עם Ctrl (כמו ב-Word), ב-Excel וב-PowerPoint בלחיצה רגילה
  document.addEventListener('click', e => {
    const a = e.target.closest && e.target.closest('a[href]'); if (!a) return;
    const root = a.closest('#weBody,#xed,#ped'); if (!root) return;
    if (a.closest('.sm-ov,.sm-bar')) return;
    let u = a.getAttribute('href') || '';
    if (!/^(https?:|mailto:|tel:|www\.)/i.test(u)) return;
    if (root.id === 'weBody' && !(e.ctrlKey || e.metaKey)) return;
    e.preventDefault(); e.stopPropagation();
    if (/^www\./i.test(u)) u = 'https://' + u;
    act({ url: u, c: classify(u) });
  }, true);

  /* ---------------- עיצוב ---------------- */
  const css = `
.page>.sm-links{position:absolute;inset:0;z-index:5;pointer-events:none;direction:ltr}
body:not([data-tool="select"]) .page>.sm-links{display:none}
.sm-l{position:absolute;pointer-events:auto;margin:0;padding:0;border:0;background:transparent;cursor:pointer;border-radius:3px;-webkit-tap-highlight-color:transparent}
.sm-l:hover,.sm-l:focus-visible{background:rgba(56,140,230,.16);outline:1.5px solid rgba(56,140,230,.55)}
.sm-l.sm-txt{border-bottom:1.5px dotted rgba(56,140,230,.55);border-radius:0}
.sm-play{position:absolute;top:-9px;left:-9px;width:20px;height:20px;border-radius:50%;background:var(--v7-accent,#2f7fe0);color:#fff;font:700 10px/20px Rubik,Arial,sans-serif;text-align:center;box-shadow:0 2px 6px rgba(15,33,63,.3);pointer-events:none}
.sm-play.sm-big{top:50%;left:50%;width:56px;height:56px;margin:-28px 0 0 -28px;font-size:22px;line-height:56px;background:rgba(15,26,48,.72);box-shadow:0 0 0 3px rgba(255,255,255,.85),0 8px 24px rgba(0,0,0,.35)}
.sm-l:hover .sm-play.sm-big{background:var(--v7-accent,#2f7fe0)}
.sm-play.sm-clip{background:#fff;color:#333;border:1px solid #c9d3e3;font-size:11px}
.sm-outside .sm-play{background:#e8a33a;color:#fff;border:0}
.sm-bar{position:sticky;top:6px;height:0;z-index:30;display:flex;justify-content:center;align-items:flex-start;overflow:visible;pointer-events:none}
.sm-pill{pointer-events:auto;flex:none;height:auto;box-sizing:border-box;display:flex;align-items:center;gap:2px;max-width:min(560px,calc(100% - 16px));background:var(--v7-surface,#fff);color:var(--v7-ink,#0f1a30);border:1px solid var(--v7-line,#d9e1ee);border-radius:999px;box-shadow:0 10px 30px -8px rgba(15,33,63,.35);font-family:"SF-rubik",Rubik,Arial,sans-serif;animation:smIn .35s ease;transition:max-width .3s}
.sm-pill-t{all:unset;cursor:pointer;display:flex;align-items:center;gap:10px;padding:8px 14px 8px 6px;min-width:0}
.sm-pill-h{flex:none;background:var(--v7-accent,#2f7fe0);color:var(--v7-on,#fff);font-size:12px;font-weight:700;padding:4px 10px;border-radius:999px}
.sm-pill-p{font-size:13.5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.sm-pill-x{all:unset;cursor:pointer;width:28px;height:28px;margin-inline-end:4px;border-radius:50%;text-align:center;line-height:28px;font-size:12px;color:var(--v7-muted,#5b6782)}
.sm-pill-x:hover{background:var(--v7-soft,#eef3fb)}
.sm-pill.sm-mini{cursor:pointer;max-width:120px}
.sm-pill.sm-mini .sm-pill-h,.sm-pill.sm-mini .sm-pill-x{display:none}
.sm-pill.sm-mini .sm-pill-t{padding:6px 12px;pointer-events:none}
.sm-bar.sm-float{position:fixed;top:auto;left:0;right:0;bottom:calc(46px + env(safe-area-inset-bottom,0px));height:auto;z-index:950}
@media (max-width:760px){.sm-bar.sm-float{bottom:calc(78px + env(safe-area-inset-bottom,0px))}}
@keyframes smIn{from{opacity:0;transform:translateY(-8px)}to{opacity:1;transform:none}}
.sm-ov{position:fixed;inset:0;z-index:2147483000;background:rgba(10,18,34,.55);backdrop-filter:blur(3px);display:flex;align-items:center;justify-content:center;padding:14px;font-family:"SF-rubik",Rubik,Arial,sans-serif}
.sm-box{background:var(--v7-surface,#fff);color:var(--v7-ink,#0f1a30);border-radius:20px;width:min(440px,100%);max-height:calc(100dvh - 28px);display:flex;flex-direction:column;box-shadow:0 24px 60px rgba(0,0,0,.35);overflow:hidden;animation:smIn .25s ease}
.sm-box.sm-wide{width:min(960px,100%)}
.sm-head{display:flex;align-items:center;gap:10px;padding:14px 18px 10px}
.sm-title{flex:1;font-weight:700;font-size:17px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.sm-x{all:unset;cursor:pointer;width:34px;height:34px;border-radius:50%;text-align:center;line-height:34px;color:var(--v7-muted,#5b6782)}
.sm-x:hover,.sm-x:focus-visible{background:var(--v7-soft,#eef3fb)}
.sm-body{padding:0 18px 6px;overflow:auto;display:flex;flex-direction:column;gap:10px}
.sm-foot{display:flex;flex-wrap:wrap;gap:8px;padding:12px 18px 16px}
.sm-foot:empty{display:none}
.sm-btn{all:unset;cursor:pointer;padding:10px 18px;border-radius:12px;font-size:14.5px;font-weight:600;background:var(--v7-soft,#eef3fb);color:var(--v7-ink,#0f1a30)}
.sm-btn.sm-pri{background:var(--v7-accent,#2f7fe0);color:var(--v7-on,#fff)}
.sm-btn:focus-visible{outline:2px solid var(--v7-accent,#2f7fe0);outline-offset:2px}
.sm-msg{font-size:14.5px;line-height:1.6;color:var(--v7-muted,#3e4a63)}
.sm-url{font:13.5px/1.5 ui-monospace,Consolas,monospace;background:var(--v7-bg,#f4f7fc);border:1px solid var(--v7-line,#d9e1ee);border-radius:10px;padding:9px 12px;word-break:break-all;user-select:all}
.sm-file b{font-size:15px}.sm-file span{color:var(--v7-muted,#5b6782);font-size:13px}
.sm-stage{position:relative;width:100%;aspect-ratio:16/9;max-height:calc(100dvh - 190px);background:#000;border-radius:12px;overflow:hidden}
.sm-stage iframe,.sm-stage video{position:absolute;inset:0;width:100%;height:100%;border:0;background:#000}
.sm-stage.sm-audio{aspect-ratio:auto;height:auto;background:transparent}
.sm-stage.sm-audio audio{position:static;width:100%;background:transparent}
.sm-stage.sm-err{aspect-ratio:auto;background:var(--v7-bg,#f4f7fc);padding:18px}
.sm-list{display:flex;flex-direction:column;gap:4px;padding-bottom:10px}
.sm-row{all:unset;cursor:pointer;display:flex;align-items:center;gap:12px;padding:10px;border-radius:12px}
.sm-row:hover,.sm-row:focus-visible{background:var(--v7-soft,#eef3fb)}
.sm-ri{flex:none;width:36px;height:36px;border-radius:10px;display:grid;place-items:center;background:var(--v7-bg,#f4f7fc);font-size:15px}
.sm-ri.sm-video,.sm-ri.sm-audio{background:var(--v7-accent,#2f7fe0);color:#fff}
.sm-ri.sm-outside{background:#e8a33a;color:#fff;font-weight:700}
.sm-rm{flex:1;min-width:0;display:flex;flex-direction:column}
.sm-rm b{text-align:right;font-size:14.5px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.sm-rm small{font-size:12.5px;color:var(--v7-muted,#5b6782);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;text-align:right}
.sm-rp{flex:none;font-size:12px;color:var(--v7-muted,#5b6782)}
@media (max-width:600px){.sm-ov{padding:0;align-items:flex-end}.sm-box{border-radius:20px 20px 0 0;width:100%}.sm-box.sm-wide{border-radius:0;height:100dvh;max-height:none;justify-content:center}.sm-stage{border-radius:0}.sm-pill-p{font-size:12.5px}}
@media (prefers-reduced-motion:reduce){.sm-pill,.sm-box{animation:none}}
`;
  const st = el('style'); st.id = 'siman-media-css'; st.textContent = css; document.head.append(st);

  window.SimanMedia = { render, office, officeItems, scan: () => S.doc && scan(S.doc), classify, list: () => S.doc && scanned.get(S.doc) };
})();
