/* =====================================================================
   סימן 5.10.0 — המרה מ-PDF ל-Word חדשה (siman-convert.js)
   "Word לעריכה": מסמך Word אמיתי שנבנה במכשיר עצמו, בלי לשלוח כלום החוצה.
   פסקאות שזורמות (גם מעמוד לעמוד), כותרות אמיתיות (מופיעות בחלונית הניווט של Word),
   רשימות תבליטים ומספור של Word, טבלאות עם מיזוג תאים ומילוי, תמונות במקומן,
   צבעי טקסט, מרקר, קו תחתון, הזחות, מרווחים, מספרי עמודים בכותרת/בתחתית.
   וגם חלון המרה חדש: בחירת סוג, התקדמות עם העמודים עצמם, והשוואה מקור/Word לפני השמירה.
   ===================================================================== */
(function () {
'use strict';
if (window.SIMAN_CONV) return;
const VER = '5.10.0';
const $id = id => document.getElementById(id);
const esc = t => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const RTL_CH = /[\u0590-\u05FF\uFB1D-\uFB4F\u0600-\u06FF\u0750-\u077F\uFB50-\uFDFF\uFE70-\uFEFF]/;
const LTR_CH = /[A-Za-z\u00C0-\u024F\u0370-\u03FF\u0400-\u04FF]/;
const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
const median = a => { if (!a.length) return 0; const s = a.slice().sort((x, y) => x - y); return s[s.length >> 1]; };
const pct = (a, p) => { if (!a.length) return 0; const s = a.slice().sort((x, y) => x - y); return s[clamp(Math.round((s.length - 1) * p), 0, s.length - 1)]; };
const sleep0 = () => new Promise(r => setTimeout(r, 0));
class Cancelled extends Error { constructor() { super('cancelled'); this.cancelled = true; } }

/* ---------- colours ---------- */
const hex2 = h => { const m = /^#?([0-9a-f]{6})$/i.exec(h || ''); if (!m) return null; const n = parseInt(m[1], 16); return [n >> 16 & 255, n >> 8 & 255, n & 255]; };
const rgb2hex = a => a.map(v => clamp(Math.round(v), 0, 255).toString(16).padStart(2, '0')).join('').toUpperCase();
const cdist = (a, b) => Math.abs(a[0] - b[0]) + Math.abs(a[1] - b[1]) + Math.abs(a[2] - b[2]);
// clean, "typed" colours: near-black is black, a slightly smeared red is pure red, and so on
const NICE = ['000000', 'FF0000', 'C00000', '0000FF', '0070C0', '00B0F0', '002060', '00B050', '008000', '7030A0', 'FFC000', 'ED7D31', '808080', '595959', 'FFFFFF', '1F4E79', '2E74B5', 'C55A11', '375623', '833C0B'].map(h => [h, hex2(h)]);
function niceColor(h) {
  const c = typeof h === 'string' ? hex2(h) : h; if (!c) return '000000';
  const mx = Math.max(...c), mn = Math.min(...c);
  if (mx < 70 && mx - mn < 45) return '000000';
  let best = null, bd = 1e9; for (const [k, v] of NICE) { const d = cdist(c, v); if (d < bd) { bd = d; best = k; } }
  return bd <= 60 ? best : rgb2hex(c.map(v => Math.round(v / 8) * 8));
}
const HL = { yellow: [255, 255, 0], green: [0, 255, 0], cyan: [0, 255, 255], magenta: [255, 0, 255], lightGray: [192, 192, 192], darkYellow: [128, 128, 0], red: [255, 0, 0], blue: [0, 0, 255] };
function hlName(c) { let best = null, bd = 1e9; for (const k in HL) { const d = cdist(c, HL[k]); if (d < bd) { bd = d; best = k; } } return bd <= 70 ? best : null; }

/* ---------- list markers ---------- */
// symbol fonts (Wingdings / Symbol) keep their bullets in the private area: back to the real character
const PUA = { 0xF0B7: '•', 0xF0A7: '▪', 0xF0D8: '➢', 0xF0FC: '✓', 0xF0FB: '✗', 0xF076: '❖', 0xF0A8: '◆', 0xF06E: '■', 0xF071: '❑', 0xF0E0: '➔', 0xF0E8: '➔', 0xF0F0: '⇨', 0xF0D7: '➢', 0xF0AF: '✶', 0xF09F: '•', 0xF06C: '●', 0xF0A1: '○', 0xF0B2: '☐', 0xF0FE: '☑', 0xF0D2: '❒', 0xF0E1: '←', 0xF0E2: '→' };
const RE_BUL = /^\s*([•·▪◦●■□○◆◇❖❑❒➢➤➔►▶▸‣⁃✓✔✗✘☐☑⇨→←✶\uE000-\uF8FF]|[-–—*](?=\s))\s*/;
const RE_NUM = /^\s*(\d{1,3})\s*[.)]\s+/;
const RE_NUM_RTL = /^\s*[.)](\d{1,3})\s+/;               // "1." read backwards in some Hebrew files
const RE_HEB = /^\s*([א-ט]|[יכלמנסעפצקרשת])\s*[.)׳']\s+/;
const RE_LAT = /^\s*([a-hA-H])\s*[.)]\s+/;
const HEB_ORD = 'אבגדהוזחטיכלמנסעפצקרשת';

/* =====================================================================
   1. READING — every page: lines, runs, colours; then the structure
   ===================================================================== */
async function canvasBlob(c, type, q) { return new Promise(r => c.toBlob(b => r(b), type, q)); }
async function blobCanvas(b) { const bm = await createImageBitmap(b); const c = document.createElement('canvas'); c.width = bm.width; c.height = bm.height; c.getContext('2d').drawImage(bm, 0, 0); bm.close && bm.close(); return c; }
const free = c => { if (c) { c.width = 0; c.height = 0; } };

async function collect(idx, opt, report, cancelled) {
  const built = await buildPdf();
  const d = await pdfjs.getDocument({ data: built.slice(), isEvalSupported: false }).promise;
  const pages = [];
  try {
    for (const [n, i] of idx.entries()) {
      if (cancelled()) throw new Cancelled();
      report({ phase: 'read', n, total: idx.length, page: i });
      const page = await d.getPage(i + 1); const P = { i: i + 1, page };
      const src = S.pages[i]; let ocr = !!(src && src.ocr);
      const orig = await renderTo(page, 2);
      P.rc = { vp: orig.vp, w: orig.c.width, data: orig.c.getContext('2d').getImageData(0, 0, orig.c.width, orig.c.height).data };
      let lines = ocr ? src.lines : await pageLines(P);
      const tilted = !lines && !!P.tilted;
      if ((!lines || !lines.length) && !tilted && !ocr && opt.ocr && src && src.page) { // a scanned page: read its words on the device
        report({ phase: 'ocr', n, total: idx.length, page: i });
        try { if (await ocrPage(src)) { ocr = true; lines = src.lines; } } catch (e) { console.warn('ocr', e); }
      }
      P.lines = lines || []; P.blocks = null;
      const blocks = P.lines.length ? (await pageBlocks(P)) || [] : [];
      for (const b of blocks) {
        const fn = fontNameOf(P, b.font); b.face = ocr ? 'Arial' : safeFace(fontFace(fn.name)); b.bold = fn.bold || /bold|black|heavy|semibold/i.test(fn.name) || b.lines.some(l => l.bold);
        Object.assign(b, sampleCanvas(orig.c, orig.vp, b));
        for (const l of b.lines) {
          if (!l.runs || !l.runs.length) l.runs = [{ t: l.text, it: null }];
          for (const r of l.runs) {
            const it = r.it;
            if (!it) { r.color = b.fg; r.bold = b.bold; r.face = b.face; r.size = l.size; continue; }
            const fr = fontNameOf(P, it.font); r.face = ocr ? 'Arial' : safeFace(fontFace(fr.name)) || b.face;
            r.bold = fr.bold || /bold|black|heavy|semibold/i.test(fr.name) || !!it.dbl; r.ital = /italic|oblique/i.test(fr.name);
            r.color = sampleCanvas(orig.c, orig.vp, { x0: it.x, x1: it.x + it.w, y0: it.y - it.h * .25, y1: it.y + it.h * .85 }).fg; r.size = it.h;
            r.x0 = it.x; r.x1 = it.x + it.w; r.raw = it.s;
          }
        }
      }
      const view = page.view;
      const pg = { idx: i, w: view[2] - view[0], h: view[3] - view[1], ox: view[0], oy: view[1], blocks, ocr, tilted, vp: orig.vp, W: orig.c.width, H: orig.c.height };
      pg.origBlob = await canvasBlob(orig.c, 'image/jpeg', .9);
      // a small picture for the comparison screen and the progress grid
      const k = Math.min(1, 700 / orig.c.width), sm = document.createElement('canvas'); sm.width = Math.round(orig.c.width * k); sm.height = Math.round(orig.c.height * k);
      sm.getContext('2d').drawImage(orig.c, 0, 0, sm.width, sm.height); pg.thumb = sm.toDataURL('image/jpeg', .8); free(sm);
      report({ phase: 'thumb', n, total: idx.length, page: i, thumb: pg.thumb });
      free(orig.c); P.rc = null;
      pages.push(pg);
      await sleep0();
    }
    // the same pages without their text: what stays is lines, fills, marker, pictures
    const doc = await PDFLib.PDFDocument.load(built, { ignoreEncryption: true });
    for (const pg of pages) {
      if (pg.ocr) continue; const rects = [];
      for (const b of pg.blocks) for (const l of b.lines) rects.push([l.x0 - .5, l.y0 - .5, l.x1 + .5, l.y1 + .5]);
      if (rects.length) try { removeTextInRects(PDFLib, doc, pg.idx, rects); } catch (e) { console.warn(e); }
    }
    const clean = await doc.save(); const dc = await pdfjs.getDocument({ data: clean, isEvalSupported: false }).promise;
    try {
      for (const [n, pg] of pages.entries()) {
        if (cancelled()) throw new Cancelled();
        report({ phase: 'shape', n, total: pages.length, page: pg.idx });
        const r = await renderTo(await dc.getPage(pg.idx + 1), 2); const bc = r.c, bctx = bc.getContext('2d');
        const oc = await blobCanvas(pg.origBlob);
        if (pg.ocr) for (const b of pg.blocks) for (const l of b.lines) { // a scan: paint over the words
          const [a, c] = r.vp.convertToViewportPoint(l.x0, l.y1), [e, f] = r.vp.convertToViewportPoint(l.x1, l.y0); bctx.fillStyle = b.bg; bctx.fillRect(Math.min(a, e) - 2, Math.min(c, f) - 2, Math.abs(e - a) + 4, Math.abs(f - c) + 4);
        }
        // text that is not really drawn (copies hidden behind pictures) must not come out
        const octx = oc.getContext('2d');
        if (!pg.ocr) for (const b of pg.blocks) for (const l of b.lines) {
          const [a, c] = r.vp.convertToViewportPoint(l.x0, l.y1), [e, f] = r.vp.convertToViewportPoint(l.x1, l.y0);
          const X = Math.max(0, Math.floor(Math.min(a, e))), Y = Math.max(0, Math.floor(Math.min(c, f))), W = Math.max(1, Math.min(bc.width - X, Math.ceil(Math.abs(e - a)))), H = Math.max(1, Math.min(bc.height - Y, Math.ceil(Math.abs(f - c))));
          const A = octx.getImageData(X, Y, W, H).data, B = bctx.getImageData(X, Y, W, H).data; let ch = 0;
          for (let k = 0; k < A.length; k += 4) if (Math.abs(A[k] - B[k]) + Math.abs(A[k + 1] - B[k + 1]) + Math.abs(A[k + 2] - B[k + 2]) > 60) ch++;
          l.hidden = ch < W * H * .004;
        }
        await shapes(pg, bc, oc);
        free(bc); free(oc); pg.origBlob = null;
        await sleep0();
      }
    } finally { dc.destroy(); }
  } finally { d.destroy(); }
  return pages;
}

/* ---------- connected pieces of ink (on a 4-pixel grid) ---------- */
function components(mask, W, H) {
  const G = 4, gw = Math.ceil(W / G), gh = Math.ceil(H / G), grid = new Uint16Array(gw * gh);
  for (let y = 0; y < H; y++) { const gy = ((y / G) | 0) * gw, row = y * W; for (let x = 0; x < W; x++) if (mask[row + x]) grid[gy + ((x / G) | 0)]++; }
  const seen = new Uint8Array(gw * gh), comps = [];
  for (let s = 0; s < grid.length; s++) {
    if (grid[s] < 2 || seen[s]) continue; const st = [s]; seen[s] = 1; let l = 1e9, t = 1e9, r = -1, b = -1, n = 0;
    while (st.length) { const c = st.pop(); n++; const cx = c % gw, cy = (c / gw) | 0; if (cx < l) l = cx; if (cx > r) r = cx; if (cy < t) t = cy; if (cy > b) b = cy;
      for (let dy = -2; dy <= 2; dy++) for (let dx = -2; dx <= 2; dx++) { const nx = cx + dx, ny = cy + dy; if (nx < 0 || ny < 0 || nx >= gw || ny >= gh) continue; const j = ny * gw + nx; if (grid[j] >= 2 && !seen[j]) { seen[j] = 1; st.push(j); } } }
    comps.push({ l: l * G, t: t * G, r: Math.min(W - 1, r * G + G - 1), b: Math.min(H - 1, b * G + G - 1), n });
  }
  let merged = true; while (merged) { merged = false;
    for (let a = 0; a < comps.length && !merged; a++) for (let b = a + 1; b < comps.length; b++) { const A = comps[a], B = comps[b];
      if (A.l - 10 <= B.r && B.l - 10 <= A.r && A.t - 10 <= B.b && B.t - 10 <= A.b && Math.min(A.r - A.l, A.b - A.t, B.r - B.l, B.b - B.t) > 16) {
        A.l = Math.min(A.l, B.l); A.t = Math.min(A.t, B.t); A.r = Math.max(A.r, B.r); A.b = Math.max(A.b, B.b); A.n += B.n; comps.splice(b, 1); merged = true; break; } } }
  return comps;
}

/* ---------- what the page holds besides its text ---------- */
async function shapes(pg, bc, oc) {
  const W = bc.width, H = bc.height, vp = pg.vp;
  const D = bc.getContext('2d').getImageData(0, 0, W, H).data;
  // paper colour = most common colour on the page edge
  const cnt = new Map(); const addc = (x, y) => { const i = (y * W + x) * 4; const q = (D[i] >> 4) + ',' + (D[i + 1] >> 4) + ',' + (D[i + 2] >> 4); cnt.set(q, (cnt.get(q) || 0) + 1); };
  for (let x = 0; x < W; x += 3) { addc(x, 2); addc(x, H - 3); } for (let y = 0; y < H; y += 3) { addc(2, y); addc(W - 3, y); }
  const pq = [...cnt.entries()].sort((a, b) => b[1] - a[1])[0][0].split(',').map(v => +v * 16 + 8);
  const paper = pq; pg.paper = paper;
  const px = (x, y) => { const i = (y * W + x) * 4; return [D[i], D[i + 1], D[i + 2]]; };
  const isInk = i => Math.abs(D[i] - paper[0]) + Math.abs(D[i + 1] - paper[1]) + Math.abs(D[i + 2] - paper[2]) > 60;
  const box = (x0, y0, x1, y1) => { const [a, b] = vp.convertToViewportPoint(x0, y1), [c, d] = vp.convertToViewportPoint(x1, y0); return { l: clamp(Math.floor(Math.min(a, c)), 0, W - 1), t: clamp(Math.floor(Math.min(b, d)), 0, H - 1), r: clamp(Math.ceil(Math.max(a, c)), 0, W - 1), b: clamp(Math.ceil(Math.max(b, d)), 0, H - 1) }; };
  pg.box = box; const k = W / pg.w; pg.k = k;
  const toPdfY = py => pg.oy + pg.h - py / k, toPdfX = pxx => pg.ox + pxx / k;
  const mask = new Uint8Array(W * H);
  for (let y = 0, i = 0; y < H; y++) for (let x = 0; x < W; x++, i += 4) if (isInk(i)) mask[y * W + x] = 1;
  const clear = (bx) => { for (let y = bx.t; y <= bx.b; y++) mask.fill(0, y * W + bx.l, y * W + bx.r + 1); };
  const lines = []; for (const b of pg.blocks) for (const l of b.lines) if (!l.hidden && l.text.trim()) { l.blk = b; lines.push(l); }
  pg.lines = lines;
  const lineBox = l => box(l.x0, l.y0, l.x1, l.y1);
  const linesIn = c => lines.filter(l => { const b = lineBox(l); const cx = (b.l + b.r) / 2, cy = (b.t + b.b) / 2; return cx >= c.l && cx <= c.r && cy >= c.t && cy <= c.b; });
  const colorStats = c => { let ink = 0, edge = 0; const hist = new Map();
    for (let y = c.t; y <= c.b; y += 2) for (let x = c.l; x <= c.r; x += 2) { if (!mask[y * W + x]) continue; ink++; if (x - c.l < 8 || c.r - x < 8 || y - c.t < 8 || c.b - y < 8) edge++;
      const p = px(x, y); const q = p.map(v => v >> 4).join(); const e = hist.get(q) || [0, 0, 0, 0]; e[0]++; e[1] += p[0]; e[2] += p[1]; e[3] += p[2]; hist.set(q, e); }
    const top = [...hist.values()].sort((a, b) => b[0] - a[0])[0];
    return { ink, edge, hist, top, fill: ink / Math.max(1, ((c.r - c.l + 1) / 2) * ((c.b - c.t + 1) / 2)), flat: !!top && top[0] / Math.max(1, ink) > .75, col: top ? [top[1] / top[0], top[2] / top[0], top[3] / top[0]] : paper }; };
  // --- marker (background fill behind words) and underline, run by run
  for (const l of lines) {
    for (const r of l.runs) {
      const x0 = r.x0 != null ? r.x0 : l.x0, x1 = r.x1 != null ? r.x1 : l.x1; if (x1 - x0 < 1) continue;
      const bx = box(x0, l.base - l.size * .15, x1, l.base + l.size * .7);
      const hist = new Map(); let tot = 0, ink = 0;
      for (let y = bx.t; y <= bx.b; y += 2) for (let x = bx.l; x <= bx.r; x += 2) {
        tot++; const i = (y * W + x) * 4; if (!isInk(i)) continue; ink++;
        const q = (D[i] >> 4) * 256 + (D[i + 1] >> 4) * 16 + (D[i + 2] >> 4); const e = hist.get(q) || [0, 0, 0, 0]; e[0]++; e[1] += D[i]; e[2] += D[i + 1]; e[3] += D[i + 2]; hist.set(q, e);
      }
      if (tot && ink / tot > .55) {
        const top = [...hist.values()].sort((a, b) => b[0] - a[0])[0];
        if (top[0] / tot > .45) { r.hl = [top[1] / top[0], top[2] / top[0], top[3] / top[0]]; }
      }
      // underline: a thin line right under the baseline, along most of the run
      const ub = box(x0, l.base - l.size * .34, x1, l.base - l.size * .02);
      const span = ub.r - ub.l + 1; let hit = -1;
      if (span > 4) for (let y = ub.t; y <= ub.b; y++) {
        let c = 0; for (let x = ub.l; x <= ub.r; x++) { const i = (y * W + x) * 4; if (isInk(i) && (!r.hl || cdist(px(x, y), r.hl) > 70)) c++; }
        if (c >= span * .7) {
          // a line that runs on past both ends of the word is a table / page rule, not an underline
          const run = (x0_, x1_) => { let n = 0; for (let x = Math.max(0, x0_); x <= Math.min(W - 1, x1_); x++) if (isInk((y * W + x) * 4)) n++; return n; };
          if (run(ub.l - 14, ub.l - 3) >= 9 && run(ub.r + 3, ub.r + 14) >= 9) continue;
          hit = y; break; }
      }
      if (hit >= 0) {
        let y2 = hit; while (y2 + 1 < H && y2 - hit < 6) { let c = 0; for (let x = ub.l; x <= ub.r; x += 2) if (isInk(((y2 + 1) * W + x) * 4)) c++; if (c < span * .35) break; y2++; }
        if (y2 - hit <= 5) { r.u = true; r.ub = { l: ub.l - 1, t: hit - 1, r: ub.r + 1, b: y2 + 1 }; }
      }
    }
  }
  // shaded boxes behind paragraphs (found before the marker, which looks the same up close)
  pg.boxes = [];
  for (const c of components(mask, W, H)) {
    const ins = linesIn(c); if (!ins.length) continue;
    const st = colorStats(c); if (!st.flat || st.fill < .55) continue;
    const lx0 = Math.min(...ins.map(l => lineBox(l).l)), lx1 = Math.max(...ins.map(l => lineBox(l).r)), sz = median(ins.map(l => l.size)) * k;
    if (!(c.l < lx0 - sz * .7 && c.r > lx1 + sz * .7)) continue;
    const col = rgb2hex(st.col); pg.boxes.push({ kind: 'fill', color: col });
    for (const l of ins) { l.boxFill = col; for (const r of l.runs) if (r.hl && cdist(r.hl, st.col) < 70) r.hl = null; }
  }
  for (const b of pg.boxes) { /* their area is cleared below with the words */ }
  for (const l of lines) for (const r of l.runs) { const x0 = r.x0 != null ? r.x0 : l.x0, x1 = r.x1 != null ? r.x1 : l.x1;
    if (r.ub) clear(r.ub); if (r.hl) { const g = l.size * .4; clear(box(x0 - 1, l.y0 - g, x1 + 1, l.y1 + g)); } }
  for (const c of components(mask, W, H)) { const ins = linesIn(c); if (ins.length && ins.every(l => l.boxFill) && colorStats(c).flat) clear(c); }
  // the words' own boxes never count as drawings (anti-aliasing leftovers)
  for (const l of lines) clear(box(l.x0, l.y0, l.x1, l.y1));
  // --- straight lines (rules) on the page
  const hl = [], vl = [];
  { const MIN = Math.round(18 * k / 2) * 2; const segs = [];
    for (let y = 0; y < H; y++) { let s = -1; const row = y * W;
      for (let x = 0; x <= W; x++) { const on = x < W && mask[row + x]; if (on && s < 0) s = x; else if (!on && s >= 0) { if (x - s >= MIN) segs.push({ y, x0: s, x1: x - 1 }); s = -1; } } }
    const open = []; for (const sg of segs) { // join the same line over consecutive rows
      let m = null; for (const o of open) if (o.y1 === sg.y - 1 && Math.min(o.x1, sg.x1) - Math.max(o.x0, sg.x0) > .8 * Math.min(o.x1 - o.x0, sg.x1 - sg.x0)) { m = o; break; }
      if (m) { m.y1 = sg.y; m.x0 = Math.min(m.x0, sg.x0); m.x1 = Math.max(m.x1, sg.x1); } else open.push({ y0: sg.y, y1: sg.y, x0: sg.x0, x1: sg.x1 }); }
    for (const o of open) if (o.y1 - o.y0 <= 7) hl.push(o);
  }
  { const MIN = Math.round(12 * k / 2) * 2; const segs = [];
    for (let x = 0; x < W; x++) { let s = -1;
      for (let y = 0; y <= H; y++) { const on = y < H && mask[y * W + x]; if (on && s < 0) s = y; else if (!on && s >= 0) { if (y - s >= MIN) segs.push({ x, y0: s, y1: y - 1 }); s = -1; } } }
    const open = []; for (const sg of segs) {
      let m = null; for (const o of open) if (o.x1 === sg.x - 1 && Math.min(o.y1, sg.y1) - Math.max(o.y0, sg.y0) > .8 * Math.min(o.y1 - o.y0, sg.y1 - sg.y0)) { m = o; break; }
      if (m) { m.x1 = sg.x; m.y0 = Math.min(m.y0, sg.y0); m.y1 = Math.max(m.y1, sg.y1); } else open.push({ x0: sg.x, x1: sg.x, y0: sg.y0, y1: sg.y1 }); }
    for (const o of open) if (o.x1 - o.x0 <= 7) vl.push(o);
  }
  const lineColor = (o) => { const a = [0, 0, 0]; let n = 0; for (let y = o.y0; y <= o.y1; y++) for (let x = o.x0; x <= o.x1; x += Math.max(1, (o.x1 - o.x0) >> 4)) { const p = px(x, y); a[0] += p[0]; a[1] += p[1]; a[2] += p[2]; n++; } return n ? niceColor(a.map(v => v / n)) : '000000'; };
  // --- tables: horizontal and vertical lines that meet
  const T = 6, all = [...hl.map(o => ({ o, h: 1 })), ...vl.map(o => ({ o, h: 0 }))];
  const par = all.map((_, i) => i); const find = i => par[i] === i ? i : (par[i] = find(par[i]));
  for (let a = 0; a < all.length; a++) for (let b = a + 1; b < all.length; b++) {
    const A = all[a], B = all[b]; if (A.h === B.h) continue; const h = A.h ? A.o : B.o, v = A.h ? B.o : A.o;
    if (v.x0 >= h.x0 - T && v.x1 <= h.x1 + T && h.y0 >= v.y0 - T && h.y1 <= v.y1 + T) par[find(a)] = find(b);
  }
  const groups = new Map(); all.forEach((e, i) => { const r = find(i); if (!groups.has(r)) groups.set(r, []); groups.get(r).push(e); });
  pg.tables = []; const usedRule = new Set();
  for (const g of groups.values()) {
    const H_ = g.filter(e => e.h).map(e => e.o), V_ = g.filter(e => !e.h).map(e => e.o);
    if (H_.length < 2 || V_.length < 2) continue;
    const bb = { l: Math.min(...V_.map(o => o.x0), ...H_.map(o => o.x0)), r: Math.max(...V_.map(o => o.x1), ...H_.map(o => o.x1)), t: Math.min(...H_.map(o => o.y0), ...V_.map(o => o.y0)), b: Math.max(...H_.map(o => o.y1), ...V_.map(o => o.y1)) };
    if (bb.r - bb.l < 40 * k / 2 || bb.b - bb.t < 24 * k / 2) continue;
    // pieces of line inside the table that touch no crossing line (cut by words) belong to it too
    for (const o of hl) if (!H_.includes(o) && o.x0 >= bb.l - T && o.x1 <= bb.r + T && o.y0 >= bb.t - T && o.y1 <= bb.b + T) { H_.push(o); g.push({ o, h: 1 }); }
    for (const o of vl) if (!V_.includes(o) && o.x0 >= bb.l - T && o.x1 <= bb.r + T && o.y0 >= bb.t - T && o.y1 <= bb.b + T) { V_.push(o); g.push({ o, h: 0 }); }
    const cl = (vals) => { vals.sort((a, b) => a - b); const out = []; for (const v of vals) { if (out.length && v - out[out.length - 1].v <= T) { const o = out[out.length - 1]; o.s += v; o.n++; o.v = o.s / o.n; } else out.push({ v, s: v, n: 1 }); } return out.map(o => Math.round(o.v)); };
    const xs = cl(V_.map(o => (o.x0 + o.x1) / 2)), ys = cl(H_.map(o => (o.y0 + o.y1) / 2));
    if (xs.length < 2 || ys.length < 2) continue;
    const whole = (bb.r - bb.l) * (bb.b - bb.t) > W * H * .55;
    const insideLines = lines.filter(l => { const b = box(l.x0, l.y0, l.x1, l.y1); const cx = (b.l + b.r) / 2, cy = (b.t + b.b) / 2; return cx > bb.l && cx < bb.r && cy > bb.t && cy < bb.b; });
    { // a drawing with a frame (a map, a chart) is not a table: much ink inside that is not lines or a flat fill
      const nearRule = (x, y) => H_.some(o => y >= o.y0 - 3 && y <= o.y1 + 3 && x >= o.x0 - 3 && x <= o.x1 + 3) || V_.some(o => x >= o.x0 - 3 && x <= o.x1 + 3 && y >= o.y0 - 3 && y <= o.y1 + 3);
      let ink = 0, n = 0; const hist = new Map();
      for (let y = bb.t + 4; y < bb.b - 4; y += 3) for (let x = bb.l + 4; x < bb.r - 4; x += 3) { n++; if (!mask[y * W + x] || nearRule(x, y)) continue; ink++; const p = px(x, y); const q = p.map(v => v >> 4).join(); hist.set(q, (hist.get(q) || 0) + 1); }
      const top = Math.max(0, ...hist.values());
      const textCells = insideLines.length;
      if ((ink > n * .02 && top < ink * .8) || (!textCells && ink > n * .005)) continue; // left for the picture finder
    }
    g.forEach(e => usedRule.add(e.o)); clear({ l: bb.l - 2, t: bb.t - 2, r: bb.r + 2, b: bb.b + 2 });
    if (whole && xs.length <= 2 && ys.length <= 2) continue; // a frame around the whole page: just leave it out
    // a border counts when the pieces of line along it cover most of it (words may have cut it into pieces)
    const cover = (segs, a, b) => { let t = 0; for (const [s0, s1] of segs) t += Math.max(0, Math.min(s1, b) - Math.max(s0, a)); return t; };
    const hasV = (x, y0, y1) => cover(V_.filter(o => Math.abs((o.x0 + o.x1) / 2 - x) <= T).map(o => [o.y0, o.y1]), y0, y1) >= (y1 - y0) * .6;
    const hasH = (y, x0, x1) => cover(H_.filter(o => Math.abs((o.y0 + o.y1) / 2 - y) <= T).map(o => [o.x0, o.x1]), x0, x1) >= (x1 - x0) * .6;
    const R = ys.length - 1, C = xs.length - 1;
    // cells, joined where the line between them is missing
    const owner = [...Array(R)].map(() => Array(C).fill(null)); const cells = [];
    for (let ri = 0; ri < R; ri++) {
      if (ys[ri + 1] - ys[ri] < 6) continue;
      for (let ci = 0; ci < C; ci++) {
        if (owner[ri][ci]) continue;
        let ce = ci; while (ce + 1 < C && !hasV(xs[ce + 1], ys[ri], ys[ri + 1])) ce++;
        const up = ri > 0 ? owner[ri - 1][ci] : null;
        if (up && up.c0 === ci && up.c1 === ce && !hasH(ys[ri], xs[ci], xs[ce + 1])) { up.r1 = ri; for (let q = ci; q <= ce; q++) owner[ri][q] = up; continue; }
        const cell = { r0: ri, r1: ri, c0: ci, c1: ce, lines: [] }; cells.push(cell); for (let q = ci; q <= ce; q++) owner[ri][q] = cell;
        ci = ce;
      }
    }
    for (const l of insideLines) { const b = box(l.x0, l.y0, l.x1, l.y1); const cx = (b.l + b.r) / 2, cy = (b.t + b.b) / 2;
      let ri = ys.findIndex((y, j) => j < R && cy >= ys[j] && cy < ys[j + 1]), ci = xs.findIndex((x, j) => j < C && cx >= xs[j] && cx < xs[j + 1]);
      if (ri < 0 || ci < 0) continue; const cell = owner[ri][ci]; if (cell) { cell.lines.push(l); l.used = 'table'; } }
    // fill colour of each cell (the words are already gone from this picture)
    for (const c of cells) {
      const L = xs[c.c0] + 5, Rr = xs[c.c1 + 1] - 5, Tt = ys[c.r0] + 5, B = ys[c.r1 + 1] - 5; if (Rr <= L || B <= Tt) continue;
      const hist = new Map(); let n = 0;
      for (let y = Tt; y <= B; y += 3) for (let x = L; x <= Rr; x += 3) { n++; const p = px(x, y); const q = p.map(v => v >> 4).join(); const e = hist.get(q) || [0, 0, 0, 0]; e[0]++; e[1] += p[0]; e[2] += p[1]; e[3] += p[2]; hist.set(q, e); }
      const top = [...hist.values()].sort((a, b) => b[0] - a[0])[0];
      if (top && top[0] / n > .6) { const col = [top[1] / top[0], top[2] / top[0], top[3] / top[0]]; if (cdist(col, paper) > 30) c.fill = rgb2hex(col); }
    }
    const thick = median(H_.map(o => o.y1 - o.y0 + 1));
    pg.tables.push({ xs, ys, cells, R, C, top: toPdfY(bb.t), bottom: toPdfY(bb.b), x0: toPdfX(bb.l), x1: toPdfX(bb.r), color: lineColor(H_[0]), sz: clamp(Math.round(thick / k * 8), 4, 24), k });
  }
  // free lines that are not part of a table: dividers between parts of the page
  pg.rules = [];
  for (const o of hl) { if (usedRule.has(o)) continue; const w = (o.x1 - o.x0) / k; if (w < pg.w * .25) continue;
    pg.rules.push({ top: toPdfY(o.y0), x0: toPdfX(o.x0), x1: toPdfX(o.x1), color: lineColor(o), sz: clamp(Math.round((o.y1 - o.y0 + 1) / k * 8), 4, 24) }); clear({ l: o.x0 - 1, t: o.y0 - 1, r: o.x1 + 1, b: o.y1 + 1 }); }
  for (const o of vl) if (!usedRule.has(o)) clear({ l: o.x0 - 1, t: o.y0 - 1, r: o.x1 + 1, b: o.y1 + 1 });
  // --- pictures: what ink is left, in connected pieces
  const comps = components(mask, W, H);
  pg.images = [];
  const pageArea = W * H;
  for (const c of comps) {
    const w = c.r - c.l + 1, h = c.b - c.t + 1; if (w < 10 || h < 10) continue;
    const inside = lines.filter(l => !l.used && (() => { const b = box(l.x0, l.y0, l.x1, l.y1); const cx = (b.l + b.r) / 2, cy = (b.t + b.b) / 2; return cx >= c.l && cx <= c.r && cy >= c.t && cy <= c.b; })());
    // a small mark at the start of a line = a drawn bullet
    if (w < 26 * k / 2 && h < 26 * k / 2) {
      const yc = toPdfY((c.t + c.b) / 2), xc = toPdfX((c.l + c.r) / 2);
      const L = lines.find(l => !l.used && yc > l.y0 - 2 && yc < l.y1 + 2 && (l.dir === 'R' ? xc > l.x1 - 2 && xc - l.x1 < l.size * 3 : xc < l.x0 + 2 && l.x0 - xc < l.size * 3));
      if (L) { L.vbullet = true; continue; }
      if (w < 14 || h < 14) continue;
    }
    if (w * h > pageArea * .6 && inside.reduce((s, l) => s + l.text.length, 0) > 250) continue; // a background behind the whole page
    // leftovers of words that were not taken off the picture (letters with vowel marks): the words are already in the text
    if (h < 34 * k && lines.some(l => { const b = box(l.x0 - l.size * .5, l.y0 - l.size * .5, l.x1 + l.size * .5, l.y1 + l.size * .5); return c.l < b.r && c.r > b.l && c.t < b.b && c.b > b.t; })) continue;
    const st = colorStats(c); const { ink, edge, hist, top, fill, flat } = st;
    if (inside.length && flat && fill > .55) { const col = rgb2hex(st.col); inside.forEach(l => { l.boxFill = col; for (const r of l.runs) if (r.hl && cdist(r.hl, st.col) < 70) r.hl = null; }); continue; }
    if (inside.length && fill < .25 && edge / Math.max(1, ink) > .8) { inside.forEach(l => l.boxLine = niceColor(st.col)); continue; }
    if (!inside.length && flat && fill > .9 && (w < 30 || h < 30)) continue; // a stray block of colour
    // a picture: cut from the page as it is drawn (with any small labels written on it)
    const pad = 3, X = Math.max(0, c.l - pad), Y = Math.max(0, c.t - pad), X2 = Math.min(W - 1, c.r + pad), Y2 = Math.min(H - 1, c.b + pad);
    const cut = document.createElement('canvas'); cut.width = X2 - X + 1; cut.height = Y2 - Y + 1;
    const smallText = inside.filter(l => true);
    cut.getContext('2d').drawImage(smallText.length ? oc : bc, X, Y, cut.width, cut.height, 0, 0, cut.width, cut.height);
    const photo = hist.size > 60 && cut.width * cut.height > 90000;
    const blob = await canvasBlob(cut, photo ? 'image/jpeg' : 'image/png', .9); free(cut);
    inside.forEach(l => l.used = 'image');
    pg.images.push({ top: toPdfY(Y), bottom: toPdfY(Y2), x0: toPdfX(X), x1: toPdfX(X2), wpt: (X2 - X + 1) / k, hpt: (Y2 - Y + 1) / k, blob, ext: photo ? 'jpeg' : 'png' });
  }
}

/* =====================================================================
   2. UNDERSTANDING — the document as a whole
   ===================================================================== */
function analyse(pages, opt) {
  const all = pages.flatMap(p => (p.lines || []).filter(l => !l.used));
  // body text size and column, from the whole document
  const sizes = new Map(); for (const l of all) { const s = Math.round(l.size * 2) / 2; sizes.set(s, (sizes.get(s) || 0) + l.text.length); }
  const body = [...sizes.entries()].sort((a, b) => b[1] - a[1])[0]; const bodySize = body ? body[0] : 11;
  const faces = new Map(); for (const l of all) for (const r of l.runs) faces.set(r.face || 'Arial', (faces.get(r.face || 'Arial') || 0) + (r.t || '').length);
  const face = [...faces.entries()].sort((a, b) => b[1] - a[1])[0]; const mainFace = face ? face[0] : 'Arial';
  const rtlDoc = all.filter(l => l.dir === 'R').reduce((s, l) => s + l.text.length, 0) >= all.reduce((s, l) => s + l.text.length, 0) / 2;
  // page numbers and repeating header / footer lines
  const band = p => (p.lines || []).filter(l => !l.used && (l.y0 > p.oy + p.h * .9 || l.y1 < p.oy + p.h * .1));
  const RE_PN = /^\s*(?:[-–—]\s*)?(?:עמוד\s*|page\s*|стр\.?\s*)?\d{1,4}(?:\s*(?:מתוך|of|из|\/)\s*\d{1,4})?(?:\s*[-–—])?\s*$/i;
  let pn = null; const pnHits = [];
  for (const p of pages) for (const l of band(p)) if (RE_PN.test(l.text)) pnHits.push({ p, l });
  if (pnHits.length >= Math.max(2, pages.length * .5)) {
    const l0 = pnHits[0].l, p0 = pnHits[0].p; const top = l0.y0 > p0.oy + p0.h / 2; const cx = (l0.x0 + l0.x1) / 2 - p0.ox;
    pn = { top, align: Math.abs(cx - p0.w / 2) < p0.w * .1 ? 'center' : (cx > p0.w / 2 ? 'right' : 'left'), size: l0.size, face: (l0.runs[0] || {}).face || mainFace, color: niceColor(l0.runs[0] && l0.runs[0].color), words: l0.text.replace(/\d+/g, '#').trim() };
    pnHits.forEach(h => h.l.used = 'pn');
  }
  const rep = new Map();
  for (const p of pages) for (const l of band(p)) { const key = l.text.replace(/\d+/g, '#').replace(/\s+/g, ' ').trim(); if (key.length < 3) continue; const e = rep.get(key) || { n: 0, l, p, top: l.y0 > p.oy + p.h / 2, all: [] }; e.n++; e.all.push(l); rep.set(key, e); }
  const hf = { header: [], footer: [] };
  if (pages.length >= 3) for (const e of rep.values()) if (e.n >= Math.max(3, pages.length * .6)) { e.all.forEach(l => l.used = 'hf'); (e.top ? hf.header : hf.footer).push(e.l); }
  // the text column (where paragraphs start and end)
  const long = all.filter(l => !l.used && (l.x1 - l.x0) > (pages[0].w * .45));
  const colL = long.length ? pct(long.map(l => l.x0), .1) : Math.min(...all.map(l => l.x0), pages[0].ox + 72);
  const colR = long.length ? pct(long.map(l => l.x1), .9) : Math.max(...all.map(l => l.x1), pages[0].ox + pages[0].w - 72);
  const tops = pages.map(p => { const ls = (p.lines || []).filter(l => !l.used); return ls.length ? p.oy + p.h - Math.max(...ls.map(l => l.y1)) : 72; });
  const bots = pages.map(p => { const ls = (p.lines || []).filter(l => !l.used); return ls.length ? Math.min(...ls.map(l => l.y0)) - p.oy : 72; });
  return { bodySize, mainFace, rtlDoc, pn, hf, colL, colR, top: clamp(pct(tops, .2), 18, 144), bottom: clamp(pct(bots, .2), 18, 144) };
}

/* ---------- lines that sit on one row become one line ---------- */
function rowLines(pg) {
  const ls = (pg.lines || []).filter(l => !l.used).slice().sort((a, b) => b.base - a.base);
  const rows = [];
  for (const l of ls) { const r = rows.find(r => Math.abs(r.base - l.base) < Math.max(2, Math.min(r.size, l.size) * .4) && Math.abs(r.size - l.size) < Math.max(r.size, l.size) * .35); if (r) r.parts.push(l); else rows.push({ base: l.base, size: l.size, parts: [l] }); }
  const out = [];
  for (const r of rows) {
    const rtl = r.parts.filter(p => p.dir === 'R').length >= r.parts.length / 2;
    r.parts.sort((a, b) => rtl ? b.x1 - a.x1 : a.x0 - b.x0);
    let cur = null;
    for (const p of r.parts) {
      if (!cur) { cur = clone(p); continue; }
      const gap = rtl ? cur.x0 - p.x1 : p.x0 - cur.x1;
      if (gap < Math.max(p.size, cur.size) * 5) { // a line that was cut by a wide gap: one line again
        cur.runs = cur.runs.concat([{ ...cur.runs[cur.runs.length - 1], t: ' ', x0: null, x1: null, u: false }], p.runs);
      } else { // a second piece far away on the row: a tab stop to the place where it starts
        cur.tabs = cur.tabs || []; cur.tabs.push(rtl ? p.x1 : p.x0);
        cur.runs = cur.runs.concat([{ ...cur.runs[cur.runs.length - 1], t: '\t', x0: null, x1: null, u: false }], p.runs);
      }
      cur.x0 = Math.min(cur.x0, p.x0); cur.x1 = Math.max(cur.x1, p.x1); cur.y0 = Math.min(cur.y0, p.y0); cur.y1 = Math.max(cur.y1, p.y1);
      cur.text = cur.runs.map(x => x.t).join(''); cur.vbullet = cur.vbullet || p.vbullet;
    }
    cur.dir = rtl ? 'R' : 'L'; out.push(cur);
  }
  return out;
}
const clone = l => ({ ...l, runs: l.runs.map(r => ({ ...r })) });

/* ---------- the list marker at the start of a line ---------- */
function takeMarker(L) {
  const first = L.runs[0]; if (!first) return null;
  let raw = (first.raw || '').trim(); const pua = raw && [...raw][0] && PUA[[...raw][0].codePointAt(0)];
  const t = L.runs.map(r => r.t).join('');
  let m, mk = null, len = 0;
  if ((m = RE_BUL.exec(t)) && t.length > m[0].length) { let g = m[1]; const cp = g.codePointAt(0); if (cp >= 0xE000 && cp <= 0xF8FF) g = PUA[cp] || '•'; if (g === '•' && pua) g = pua; if (/^[-–—*]$/.test(g)) g = '–'; mk = { kind: 'bullet', glyph: g }; len = m[0].length; }
  else if ((m = RE_NUM.exec(t)) || (m = RE_NUM_RTL.exec(t))) { mk = { kind: 'num', fmt: 'decimal', n: +m[1] }; len = m[0].length; }
  else if (L.dir === 'R' && (m = RE_HEB.exec(t)) && HEB_ORD.includes(m[1])) { mk = { kind: 'num', fmt: 'hebrew1', n: HEB_ORD.indexOf(m[1]) + 1 }; len = m[0].length; }
  else if (L.dir === 'L' && (m = RE_LAT.exec(t))) { const c = m[1]; mk = { kind: 'num', fmt: c === c.toLowerCase() ? 'lowerLetter' : 'upperLetter', n: c.toLowerCase().charCodeAt(0) - 96 }; len = m[0].length; }
  else if (L.vbullet) { mk = { kind: 'bullet', glyph: '•' }; len = 0; }
  if (!mk) return null;
  // the text after the marker: where it starts on the page (for the hanging indent)
  let k = len; const runs = L.runs; let i = 0;
  while (k > 0 && i < runs.length) { const n = runs[i].t.length; if (n <= k) { k -= n; i++; } else { runs[i] = { ...runs[i], t: runs[i].t.slice(k), x0: null, x1: null }; k = 0; } }
  const markerRuns = runs.slice(0, i); L.runs = runs.slice(i); if (L.runs[0]) L.runs[0].t = L.runs[0].t.replace(/^\s+/, '');
  const firstText = L.runs.find(r => r.x0 != null);
  if (markerRuns.length && firstText) mk.textEdge = L.dir === 'R' ? firstText.x1 : firstText.x0;
  else if (markerRuns.length) { const mr = markerRuns[markerRuns.length - 1]; if (mr.x0 != null) mk.textEdge = L.dir === 'R' ? mr.x0 - L.size * .35 : mr.x1 + L.size * .35; }
  mk.edge = L.dir === 'R' ? L.x1 : L.x0;
  if (L.vbullet) mk.edge += L.dir === 'R' ? L.size * 1.2 : -L.size * 1.2;
  L.text = L.runs.map(r => r.t).join('');
  return mk;
}

/* ---------- paragraphs ---------- */
function allBold(L) { const t = L.runs.filter(r => r.t.trim()); return t.length && t.every(r => r.bold); }
function mainColor(L) { const w = new Map(); for (const r of L.runs) { const c = niceColor(r.color); w.set(c, (w.get(c) || 0) + r.t.length); } return [...w.entries()].sort((a, b) => b[1] - a[1])[0][0]; }
function headish(L, A) { return L.size >= A.bodySize * 1.12 || (allBold(L) && mainColor(L) !== '000000' && L.text.length < 90); }
function firstWordW(L) { const t = L.text.trim(); const w = t.split(/\s+/)[0] || ''; return (L.x1 - L.x0) * (w.length + 1) / Math.max(1, t.length); }

function buildFlow(pages, A, opt) {
  const flow = []; // {type:'p'|'table'|'img'|'rule'|'break', ...}
  const colW = A.colR - A.colL;
  pages.forEach((pg, pi) => {
    if (pi && opt.pageBreaks) flow.push({ type: 'break', pg });
    if (pi) flow.push({ type: 'pagejoin', pg });
    const lines = rowLines(pg);
    for (const L of lines) { L.mk = takeMarker(L); L.head = headish(L, A); L.pg = pg; }
    const items = [
      ...lines.map(L => ({ type: 'line', top: L.y1, L })),
      ...(pg.tables || []).map(t => ({ type: 'table', top: t.top, t })),
      ...(pg.images || []).map(m => ({ type: 'img', top: m.top, m })),
      ...(pg.rules || []).map(r => ({ type: 'rule', top: r.top, r }))
    ].sort((a, b) => b.top - a.top);
    if (!items.length && (pg.tilted || pg.ocr === false)) flow.push({ type: 'wholepage', pg });
    let P = null;
    const close = () => { if (P) { flow.push(P); P = null; } };
    for (const it of items) {
      if (it.type !== 'line') { close(); flow.push({ ...it, pg }); continue; }
      const L = it.L;
      if (P && joins(P, L, A)) { P.lines.push(L); if (P.lines.length === 2) P.lh = P.lines[0].base - L.base; continue; }
      close(); P = { type: 'p', lines: [L], pg, mk: L.mk, head: L.head, dir: L.dir };
    }
    close();
  });
  // a paragraph cut by the end of a page goes on from where it stopped
  if (!opt.pageBreaks) {
    for (let i = 0; i < flow.length; i++) {
      if (flow[i].type !== 'pagejoin') continue;
      const prev = flow[i - 1], next = flow[i + 1];
      if (prev && next && prev.type === 'p' && next.type === 'p' && !next.mk && !prev.head && !next.head && prev.dir === next.dir && Math.abs(prev.lines[0].size - next.lines[0].size) < 1) {
        const last = prev.lines[prev.lines.length - 1], f = next.lines[0];
        if (reachesEnd(prev, last, f, A) && !/[.:!?״"']\s*$/.test(last.text) || reachesEnd(prev, last, f, A) && !/[.!?]\s*$/.test(last.text) && /^[a-zא-ת,(]/i.test(f.text.trim())) {
          prev.lines.push(...next.lines.map(l => ({ ...l, cont: true }))); flow.splice(i + 1, 1);
        }
      }
    }
  }
  return flow.filter(f => f.type !== 'pagejoin');
}
function reachesEnd(P, last, next, A) {
  const slack = P.dir === 'R' ? last.x0 - edgeEnd(P, A) : edgeEnd(P, A) - last.x1;
  return slack <= Math.max(last.size * 1.3, firstWordW(next) + last.size * 1.6);
}
function edgeEnd(P, A) { // where the lines of this paragraph end (the column, or the paragraph's own narrower edge)
  const ls = P.lines; if (ls.length >= 3) { const body = ls.slice(0, -1); return P.dir === 'R' ? Math.min(...body.map(l => l.x0)) : Math.max(...body.map(l => l.x1)); }
  return P.dir === 'R' ? A.colL : A.colR;
}
function joins(P, L, A) {
  const prev = P.lines[P.lines.length - 1];
  if (L.mk || L.dir !== P.dir || L.head !== P.head) return false;
  if (Math.abs(L.size - prev.size) > Math.max(L.size, prev.size) * .12) return false;
  if (!!L.boxFill !== !!prev.boxFill || !!L.boxLine !== !!prev.boxLine) return false;
  if (L.tabs || prev.tabs) return false;
  const gap = prev.base - L.base;
  if (P.lh) { if (Math.abs(gap - P.lh) > L.size * .3) return false; }
  else if (gap < L.size * .85 || gap > Math.max(L.size, prev.size) * 1.75) return false;
  // the new line starts where the paragraph's lines start
  const st = P.dir === 'R' ? L.x1 : L.x0, ref = P.lines.length > 1 ? (P.dir === 'R' ? P.lines[1].x1 : P.lines[1].x0) : null, first = P.dir === 'R' ? P.lines[0].x1 : P.lines[0].x0;
  const markEdge = P.mk && P.mk.textEdge != null ? P.mk.textEdge : null;
  const okStart = ref != null ? Math.abs(st - ref) < L.size * 1.2 : (Math.abs(st - first) < L.size * 4 || (markEdge != null && Math.abs(st - markEdge) < L.size * 1.5));
  const centered = isCentered(prev, A) && isCentered(L, A);
  if (!okStart && !centered) return false;
  if (centered) return true;
  return reachesEnd(P, prev, L, A);
}
function isCentered(L, A) { const c = (L.x0 + L.x1) / 2, cc = (A.colL + A.colR) / 2, w = L.x1 - L.x0; return Math.abs(c - cc) < Math.max(L.size * 1.5, (A.colR - A.colL) * .04) && w < (A.colR - A.colL) * .88 && (L.x0 - A.colL) > L.size * 2; }

/* =====================================================================
   3. WRITING — a real Word file
   ===================================================================== */
const TW = 20; // twips per point
function runXml(r, ctx, base) {
  let t = String(r.t ?? ''); if (!t) return '';
  const pr = [];
  const face = ctx.opt.font === 'orig' ? (r.face || ctx.A.mainFace) : ctx.opt.font;
  if (face !== ctx.A.mainFace || ctx.opt.font !== 'orig') pr.push(`<w:rFonts w:ascii="${esc(face)}" w:hAnsi="${esc(face)}" w:cs="${esc(face)}"/>`);
  if (r.bold) pr.push('<w:b/><w:bCs/>');
  if (r.ital) pr.push('<w:i/><w:iCs/>');
  const col = ctx.opt.colors ? niceColor(r.color) : '000000';
  if (col !== '000000') pr.push(`<w:color w:val="${col}"/>`);
  const sz = Math.round((r.size || base) * 2); if (sz !== Math.round(ctx.A.bodySize * 2)) pr.push(`<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`);
  if (ctx.opt.colors && r.hl) { const n = hlName(r.hl); if (n) pr.push(`<w:highlight w:val="${n}"/>`); else pr.push(`<w:shd w:val="clear" w:color="auto" w:fill="${rgb2hex(r.hl)}"/>`); }
  if (r.u) pr.push('<w:u w:val="single"/>');
  // Hebrew pieces and English pieces each get their own direction
  const segs = []; for (const ch of t) { const d = RTL_CH.test(ch) ? 'R' : LTR_CH.test(ch) ? 'L' : null; const l = segs[segs.length - 1]; if (l && (d === null || d === l.d)) l.t += ch; else segs.push({ d: d || (l ? l.d : ctx.dir), t: ch }); }
  return segs.map(g => {
    const p = pr.join('') + (g.d === 'R' ? '<w:rtl/>' : '');
    const body = g.t.split('\t').map(x => x ? `<w:t xml:space="preserve">${esc(x)}</w:t>` : '').join('<w:tab/>');
    return `<w:r>${p ? `<w:rPr>${p}</w:rPr>` : ''}${body}</w:r>`;
  }).join('');
}
function mergeRuns(runs) { // neighbours with the same look become one run
  const out = []; const key = r => [r.face, !!r.bold, !!r.ital, niceColor(r.color), Math.round((r.size || 0) * 2), r.hl ? rgb2hex(r.hl) : '', !!r.u].join('|');
  for (const r of runs) { if (!r.t) continue; const l = out[out.length - 1]; if (l && key(l) === key(r)) l.t += r.t; else out.push({ ...r }); }
  return out;
}
function paraXml(P, ctx) {
  const { A, opt } = ctx; const ls = P.lines; const f = ls[0];
  const rtl = P.dir === 'R'; const size = median(ls.map(l => l.size));
  // runs of all lines, joined with a space where a line wrapped
  const runs = []; ls.forEach((l, k) => { const rs = l.runs.map(r => ({ ...r })); if (k && rs.length) { const prevT = runs.length ? runs[runs.length - 1].t : ''; if (!/[\s\-־]$/.test(prevT)) rs[0].t = ' ' + rs[0].t.replace(/^\s+/, ''); } runs.push(...rs); });
  const pr = [];
  let style = '';
  if (P.head && ls.length <= 3) { const lvl = ctx.headLevel(size, P); style = `Heading${lvl}`; pr.push(`<w:pStyle w:val="${style}"/>`); }
  else if (P.mk) pr.push('<w:pStyle w:val="ListParagraph"/>');
  if (P.mk) pr.push(`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${ctx.listId(P)}"/></w:numPr>`);
  const box = f.boxFill || f.boxLine;
  if (f.boxLine) pr.push(`<w:pBdr>${['top', 'left', 'bottom', 'right'].map(s => `<w:${s} w:val="single" w:sz="6" w:space="4" w:color="${f.boxLine}"/>`).join('')}</w:pBdr>`);
  if (f.boxFill && opt.colors) pr.push(`<w:shd w:val="clear" w:color="auto" w:fill="${f.boxFill}"/>`);
  if (rtl) pr.push('<w:bidi/>');
  // space before this paragraph, and the line spacing inside it
  const mult = ls.length > 1 && P.lh ? clamp(P.lh / (size * 1.15), .9, 3) : ctx.lineMult;
  const adv = mult * size * 1.15;
  let before = 0;
  if (P.prevBottom != null) before = clamp(P.prevBottom - f.base - adv, 0, 120);
  else if (P.firstOnPage && !opt.pageBreaks) before = 0;
  else if (P.firstOnPage) before = clamp((f.pg.oy + f.pg.h - A.top) - f.y1, 0, 200);
  if (box && P.prevBox) before = 0;
  pr.push(`<w:spacing w:before="${Math.round(before * TW / 10) * 10}" w:after="0" w:line="${Math.round(240 * mult)}" w:lineRule="auto"/>`);
  // indents (start = right in Hebrew)
  const colS = rtl ? A.colR : A.colL, colE = rtl ? A.colL : A.colR, sgn = rtl ? 1 : -1;
  const startOf = l => rtl ? l.x1 : l.x0, endOf = l => rtl ? l.x0 : l.x1;
  const firstStart = startOf(f), bodyStart = ls.length > 1 ? median(ls.slice(1).map(startOf)) : firstStart;
  let ind = '', jc = '';
  const centered = ls.every(l => isCentered(l, A)) && !P.mk;
  if (centered) jc = 'center';
  else {
    let start, hang = 0, first = 0;
    if (P.mk) {
      const te = P.mk.textEdge != null ? P.mk.textEdge : (ls.length > 1 ? bodyStart : firstStart - sgn * size * 1.4);
      start = (colS - te) * sgn; hang = (P.mk.edge - te) * sgn; if (!(hang > size * .3)) hang = size * 1.2; if (hang > size * 4) hang = size * 1.5;
    } else { start = (colS - bodyStart) * sgn; first = (bodyStart - firstStart) * sgn; }
    start = start < 3 ? 0 : start;
    let end = 0; if (ls.length >= 3) { const ends = ls.slice(0, -1).map(endOf); const e = rtl ? Math.min(...ends) : Math.max(...ends); end = (e - colE) * sgn; if (end < size * 1.5) end = 0; }
    if (ls.length === 1 && !P.mk) { const e = endOf(f); const room = (e - colE) * sgn; if (start > (A.colR - A.colL) * .5 && room < size * 2) jc = 'end'; }
    const a = []; if (start > 0) a.push(`w:start="${Math.round(start * TW)}"`); if (end > 0) a.push(`w:end="${Math.round(end * TW)}"`);
    if (P.mk) a.push(`w:hanging="${Math.round(hang * TW)}"`); else if (Math.abs(first) > size * .6) a.push(first > 0 ? `w:hanging="${Math.round(first * TW)}"` : `w:firstLine="${Math.round(-first * TW)}"`);
    if (a.length) ind = `<w:ind ${a.join(' ')}/>`;
    if (!jc && ls.length > 1) { const full = ls.slice(0, -1).every(l => Math.abs(endOf(l) - (rtl ? Math.min(...ls.slice(0, -1).map(endOf)) : Math.max(...ls.slice(0, -1).map(endOf)))) < size * .5); const ragged = ls.slice(0, -1).some(l => Math.abs(endOf(l) - endOf(ls[0])) > size * .8); if (full && !ragged && ls.length > 2) jc = 'both'; }
  }
  if (f.tabs && f.tabs.length) pr.push(`<w:tabs>${f.tabs.map(x => `<w:tab w:val="start" w:pos="${Math.round((rtl ? colS - x : x - colS) * TW)}"/>`).join('')}</w:tabs>`);
  if (ind) pr.push(ind); if (jc) pr.push(`<w:jc w:val="${jc}"/>`);
  if (style) pr.push('<w:keepNext/>');
  const rx = mergeRuns(runs).map(r => runXml(r, { ...ctx, dir: P.dir }, size)).join('');
  return `<w:p><w:pPr>${pr.join('')}</w:pPr>${rx}</w:p>`;
}
function cellParas(lines, A, ctx, rtl) {
  if (!lines.length) return '<w:p><w:pPr>' + (rtl ? '<w:bidi/>' : '') + '<w:spacing w:before="0" w:after="0"/></w:pPr></w:p>';
  const ls = lines.slice().sort((a, b) => b.base - a.base); const ps = [];
  for (const l of ls) { const p = ps[ps.length - 1]; const prev = p && p.lines[p.lines.length - 1];
    if (p && Math.abs(prev.size - l.size) < l.size * .15 && prev.base - l.base < l.size * 1.8 && prev.base - l.base > l.size * .6 && !RE_BUL.test(l.text) && !RE_NUM.test(l.text)) p.lines.push(l); else ps.push({ lines: [l] }); }
  return ps.map(p => {
    const size = median(p.lines.map(l => l.size)); const runs = []; p.lines.forEach((l, k) => { const rs = l.runs.map(r => ({ ...r })); if (k && rs.length) rs[0].t = ' ' + rs[0].t.replace(/^\s+/, ''); runs.push(...rs); });
    const dir = p.lines[0].dir; const c = p.cellCenter;
    return `<w:p><w:pPr>${dir === 'R' ? '<w:bidi/>' : ''}<w:spacing w:before="0" w:after="0" w:line="${Math.round(240 * Math.min(ctx.lineMult, 1.15))}" w:lineRule="auto"/>${p.lines[0].center ? '<w:jc w:val="center"/>' : ''}</w:pPr>${mergeRuns(runs).map(r => runXml(r, { ...ctx, dir }, size)).join('')}</w:p>`;
  }).join('');
}
function tableXml(T, ctx) {
  const { A, opt } = ctx; const k = T.k; const rtl = A.rtlDoc;
  const colsPt = []; for (let c = 0; c < T.C; c++) colsPt.push((T.xs[c + 1] - T.xs[c]) / k);
  const order = [...Array(T.C).keys()]; if (rtl) order.reverse(); // bidiVisual: the first column is on the right
  const tw = colsPt.reduce((a, b) => a + b, 0);
  const brd = `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(s => `<w:${s} w:val="single" w:sz="${T.sz}" w:space="0" w:color="${opt.colors ? T.color : '000000'}"/>`).join('')}</w:tblBorders>`;
  const indPt = rtl ? A.colR - T.x1 : T.x0 - A.colL; const centered = Math.abs(((T.x0 + T.x1) / 2) - (A.colL + A.colR) / 2) < 6 && indPt > 6;
  let x = `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/>${rtl ? '<w:bidiVisual/>' : ''}<w:tblW w:w="${Math.round(tw * TW)}" w:type="dxa"/>${centered ? '<w:jc w:val="center"/>' : indPt > 3 ? `<w:tblInd w:w="${Math.round(indPt * TW)}" w:type="dxa"/>` : ''}${brd}<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="28" w:type="dxa"/><w:left w:w="85" w:type="dxa"/><w:bottom w:w="28" w:type="dxa"/><w:right w:w="85" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/></w:tblPr>`;
  x += `<w:tblGrid>${order.map(c => `<w:gridCol w:w="${Math.round(colsPt[c] * TW)}"/>`).join('')}</w:tblGrid>`;
  for (let r = 0; r < T.R; r++) {
    const hpt = (T.ys[r + 1] - T.ys[r]) / k; if (hpt < 3) continue;
    if (hpt < 14 && !T.cells.some(z => z.r0 <= r && r <= z.r1 && z.lines.length)) continue; // a sliver between two close lines
    x += `<w:tr><w:trPr><w:trHeight w:val="${Math.round(hpt * TW)}" w:hRule="atLeast"/><w:cantSplit/></w:trPr>`;
    // the cells of this row in the grid order (right to left in Hebrew)
    const seen = new Set();
    for (const c of order) {
      const cell = T.cells.find(z => z.c0 <= c && c <= z.c1 && z.r0 <= r && r <= z.r1);
      const key = cell ? cell : { c0: c, c1: c, r0: r, r1: r, lines: [] }; if (seen.has(key)) continue; seen.add(key);
      const span = key.c1 - key.c0 + 1; const wpt = colsPt.slice(key.c0, key.c1 + 1).reduce((a, b) => a + b, 0);
      const vm = key.r1 > key.r0 ? (r === key.r0 ? '<w:vMerge w:val="restart"/>' : '<w:vMerge/>') : '';
      const fill = key.fill && opt.colors ? `<w:shd w:val="clear" w:color="auto" w:fill="${key.fill}"/>` : '';
      if (r === key.r0) { const cx = (T.xs[key.c0] + T.xs[key.c1 + 1]) / 2 / k + (T.x0 - T.xs[0] / k); for (const l of key.lines) { const lc = (l.x0 + l.x1) / 2; l.center = Math.abs(lc - (T.x0 + (T.xs[key.c0] + T.xs[key.c1 + 1]) / 2 / k - T.xs[0] / k)) < l.size * 1.2 && (l.x1 - l.x0) < wpt * .8; } }
      const inner = r === key.r0 ? cellParas(key.lines, A, ctx, rtl) : `<w:p><w:pPr>${rtl ? '<w:bidi/>' : ''}<w:spacing w:before="0" w:after="0"/></w:pPr></w:p>`;
      x += `<w:tc><w:tcPr><w:tcW w:w="${Math.round(wpt * TW)}" w:type="dxa"/>${span > 1 ? `<w:gridSpan w:val="${span}"/>` : ''}${vm}${fill}<w:vAlign w:val="center"/></w:tcPr>${inner}</w:tc>`;
    }
    x += '</w:tr>';
  }
  return x + '</w:tbl>';
}
function imgXml(m, ctx, rid, id, maxW) {
  let w = m.wpt, h = m.hpt; if (w > maxW) { h *= maxW / w; w = maxW; }
  const E = 12700, cx = Math.round(w * E), cy = Math.round(h * E);
  return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="תמונה ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${id}" name="img${id}.${m.ext}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}


/* ---------- Word is strict about the order of properties: put them in the order the standard asks for ---------- */
const P_ORDER = 'pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange'.split(' ');
const R_ORDER = 'rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath'.split(' ');
function orderProps(xml) {
  const sortIn = (inner, order) => { const parts = []; const re = /<w:(\w+)\b[^>]*?(?:\/>|>[\s\S]*?<\/w:\1>)/g; let m; while ((m = re.exec(inner))) parts.push({ n: m[1], x: m[0] });
    if (parts.map(p => p.x).join('') !== inner) return inner; // anything unexpected: leave it alone
    return parts.map((p, i) => ({ ...p, i, o: order.indexOf(p.n) < 0 ? 999 : order.indexOf(p.n) })).sort((a, b) => a.o - b.o || a.i - b.i).map(p => p.x).join(''); };
  return xml.replace(/<w:rPr>((?:(?!<\/w:rPr>)[\s\S])*)<\/w:rPr>/g, (a, inner) => '<w:rPr>' + sortIn(inner, R_ORDER) + '</w:rPr>')
    .replace(/<w:pPr>((?:(?!<w:pPr>)[\s\S])*?)<\/w:pPr>/g, (a, inner) => '<w:pPr>' + sortIn(inner, P_ORDER) + '</w:pPr>');
}
async function writeDocx(pages, A, opt, name) {
  const flow = buildFlow(pages, A, opt);
  const z0 = new JSZip(), z = { file: (n, d) => z0.file(n, d, { createFolders: false }), generateAsync: o => z0.generateAsync(o) }, rels = [], stats = { p: 0, h: 0, list: 0, table: 0, img: 0, pages: pages.length };
  // the usual line spacing of the document
  const mults = flow.filter(f => f.type === 'p' && f.lh && f.lines.length > 1).map(f => clamp(f.lh / (median(f.lines.map(l => l.size)) * 1.15), .9, 3));
  const lineMult = mults.length ? Math.round(median(mults) * 20) / 20 : 1.15;
  // heading levels by size (bigger = higher)
  const hs = [...new Set(flow.filter(f => f.type === 'p' && f.head && f.lines.length <= 3).map(f => Math.round(median(f.lines.map(l => l.size)) * 2) / 2))].sort((a, b) => b - a);
  const headLevel = (size) => clamp(hs.indexOf(Math.round(size * 2) / 2) + 1, 1, 3);
  // lists: one Word list per run of items that share a marker kind and indent
  const abst = [], nums = []; const listKey = new Map(); let lastList = null;
  const listId = (P) => {
    const mk = P.mk; const kind = mk.kind === 'bullet' ? 'b:' + mk.glyph : 'n:' + mk.fmt;
    let ai = listKey.get(kind); if (ai == null) { ai = abst.length; abst.push(mk); listKey.set(kind, ai); }
    if (mk.kind === 'bullet') { if (!nums.find(n => n.ai === ai && n.bullet)) nums.push({ id: nums.length + 1, ai, bullet: true }); return nums.find(n => n.ai === ai && n.bullet).id; }
    // numbers: a new numbering when the count starts again
    if (lastList && lastList.ai === ai && mk.n === lastList.next) { lastList.next++; return lastList.id; }
    const n = { id: nums.length + 1, ai, start: mk.n, next: mk.n + 1 }; nums.push(n); lastList = n; return n.id;
  };
  const ctx = { A, opt, lineMult, headLevel, listId };
  const body = []; let imgN = 0; let prevBottom = null, prevBox = false, pageFirst = true;
  const textW = (pages[0].w - (A.colL - pages[0].ox) - (pages[0].ox + pages[0].w - A.colR));
  // sections: a page of another size or turned sideways starts a new section
  const sect = (pg, last) => { const Wt = Math.round(pg.w * TW), Ht = Math.round(pg.h * TW);
    const ml = Math.round(clamp(A.colL - pg.ox, 18, 144) * TW), mr = Math.round(clamp(pg.ox + pg.w - A.colR, 18, 144) * TW);
    const l_ = A.rtlDoc ? mr : ml, r_ = A.rtlDoc ? ml : mr; // Word's left/right margins are the page's sides
    return `<w:sectPr>${A.pn || A.hf.header.length ? '<w:headerReference w:type="default" r:id="rIdHdr"/>' : ''}${A.pn || A.hf.footer.length ? '<w:footerReference w:type="default" r:id="rIdFtr"/>' : ''}${last ? '' : '<w:type w:val="nextPage"/>'}<w:pgSz w:w="${Wt}" w:h="${Ht}"${Wt > Ht ? ' w:orient="landscape"' : ''}/><w:pgMar w:top="${Math.round(A.top * TW)}" w:right="${A.rtlDoc ? ml : mr}" w:bottom="${Math.round(A.bottom * TW)}" w:left="${A.rtlDoc ? mr : ml}" w:header="${Math.round(Math.max(12, A.top * .45) * TW)}" w:footer="${Math.round(Math.max(12, A.bottom * .45) * TW)}" w:gutter="0"/>${A.rtlDoc ? '<w:bidi/>' : ''}<w:docGrid w:linePitch="360"/></w:sectPr>`; };
  // fix the margins: Word's w:right is always the right side of the page
  const sect2 = (pg, last) => { const s = sect(pg, last); const ml = Math.round(clamp(A.colL - pg.ox, 18, 144) * TW), mr = Math.round(clamp(pg.ox + pg.w - A.colR, 18, 144) * TW);
    return s.replace(/w:right="\d+"/, `w:right="${mr}"`).replace(/w:left="\d+"/, `w:left="${ml}"`); };
  let lastPg = pages[0];
  for (let i = 0; i < flow.length; i++) {
    const f = flow[i];
    if (f.pg && f.pg !== lastPg) {
      const sizeChange = Math.abs(f.pg.w - lastPg.w) > 2 || Math.abs(f.pg.h - lastPg.h) > 2;
      if (sizeChange) body.push(`<w:p><w:pPr>${sect2(lastPg, false)}</w:pPr></w:p>`);
      lastPg = f.pg; pageFirst = true; prevBottom = null;
    }
    if (f.type === 'break') { if (body.length) body.push('<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>'); prevBottom = null; pageFirst = true; continue; }
    if (f.type === 'p') {
      f.prevBottom = prevBottom; f.firstOnPage = pageFirst; f.prevBox = prevBox;
      if (pageFirst && !opt.pageBreaks && body.length) f.prevBottom = null;
      body.push(paraXml(f, ctx)); stats.p++; if (f.head && f.lines.length <= 3) stats.h++; if (f.mk) stats.list++;
      const last = f.lines[f.lines.length - 1]; prevBottom = last.base; prevBox = !!(last.boxFill || last.boxLine);
      if (last.pg) lastPg = last.pg; // a paragraph that went on to the next page ends there
      pageFirst = false; continue;
    }
    if (f.type === 'table') {
      const before = prevBottom != null ? clamp(prevBottom - f.t.top - 4, 0, 60) : 0;
      if (before > 2) body.push(`<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${Math.round(before * TW)}" w:lineRule="exact"/></w:pPr></w:p>`);
      body.push(tableXml(f.t, ctx)); stats.table++;
      body.push(`<w:p><w:pPr>${A.rtlDoc ? '<w:bidi/>' : ''}<w:spacing w:before="0" w:after="0" w:line="120" w:lineRule="exact"/></w:pPr></w:p>`);
      prevBottom = f.t.bottom; pageFirst = false; prevBox = false; continue;
    }
    if (f.type === 'img' || f.type === 'wholepage') {
      let m = f.m;
      if (f.type === 'wholepage') continue;
      imgN++; const rid = 'rIdImg' + imgN; z.file(`word/media/image${imgN}.${m.ext}`, m.blob); rels.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image${imgN}.${m.ext}"/>`);
      const cx = (m.x0 + m.x1) / 2, cc = (A.colL + A.colR) / 2; const jc = Math.abs(cx - cc) < (A.colR - A.colL) * .12 ? 'center' : ((cx > cc) === A.rtlDoc ? '' : 'end');
      const before = prevBottom != null ? clamp(prevBottom - m.top - 6, 0, 80) : 0;
      body.push(`<w:p><w:pPr>${A.rtlDoc ? '<w:bidi/>' : ''}<w:spacing w:before="${Math.round(before * TW)}" w:after="0" w:line="240" w:lineRule="auto"/>${jc ? `<w:jc w:val="${jc}"/>` : ''}</w:pPr>${imgXml(m, ctx, rid, 1000 + imgN, textW)}</w:p>`);
      stats.img++; prevBottom = m.bottom; pageFirst = false; prevBox = false; continue;
    }
    if (f.type === 'rule') {
      const before = prevBottom != null ? clamp(prevBottom - f.r.top - 8, 0, 60) : 0;
      body.push(`<w:p><w:pPr>${A.rtlDoc ? '<w:bidi/>' : ''}<w:pBdr><w:bottom w:val="single" w:sz="${f.r.sz}" w:space="1" w:color="${opt.colors ? f.r.color : '000000'}"/></w:pBdr><w:spacing w:before="${Math.round(before * TW)}" w:after="0" w:line="120" w:lineRule="exact"/></w:pPr></w:p>`);
      prevBottom = f.r.top; pageFirst = false; continue;
    }
  }
  // pages that came out with no text at all (a drawing, a scan that was not read): the page as a picture
  for (const pg of pages) if (!(pg.lines || []).length && pg.full) { /* reserved */ }
  body.push(sect2(lastPg, true));

  // ---- numbering
  const lvlXml = (mk, i) => {
    const ind = `<w:pPr><w:ind w:start="${360 * (i + 1) + 360}" w:hanging="360"/></w:pPr>`;
    if (mk.kind === 'bullet') { const sym = /[✓✔✗✘➢➤❖◆■□☐☑❑❒➔⇨]/.test(mk.glyph); return `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${esc(mk.glyph)}"/><w:lvlJc w:val="left"/>${ind}<w:rPr><w:rFonts w:ascii="${sym ? 'Segoe UI Symbol' : 'Arial'}" w:hAnsi="${sym ? 'Segoe UI Symbol' : 'Arial'}" w:cs="Arial" w:hint="default"/><w:b w:val="0"/><w:color w:val="auto"/></w:rPr></w:lvl>`; }
    return `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${mk.fmt}"/><w:lvlText w:val="%${i + 1}."/><w:lvlJc w:val="left"/>${ind}<w:rPr><w:b w:val="0"/><w:bCs w:val="0"/><w:color w:val="auto"/></w:rPr></w:lvl>`;
  };
  const numbering = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${abst.map((mk, i) => `<w:abstractNum w:abstractNumId="${i}"><w:multiLevelType w:val="hybridMultilevel"/>${[0, 1, 2].map(l => lvlXml(mk, l)).join('')}</w:abstractNum>`).join('')}${nums.map(n => `<w:num w:numId="${n.id}"><w:abstractNumId w:val="${n.ai}"/>${n.bullet ? '' : `<w:lvlOverride w:ilvl="0"><w:startOverride w:val="${n.start}"/></w:lvlOverride>`}</w:num>`).join('')}</w:numbering>`;
  // ---- styles: body font and size from the document itself
  const face = opt.font === 'orig' ? A.mainFace : opt.font, bs = Math.round(A.bodySize * 2);
  const lang = A.rtlDoc ? '<w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="he-IL"/>' : '<w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="he-IL"/>';
  const hStyle = (n) => `<w:style w:type="paragraph" w:styleId="Heading${n}"><w:name w:val="heading ${n}"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:outlineLvl w:val="${n - 1}"/></w:pPr><w:rPr><w:b/><w:bCs/></w:rPr></w:style>`;
  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${esc(face)}" w:eastAsia="${esc(face)}" w:hAnsi="${esc(face)}" w:cs="${esc(face)}"/><w:sz w:val="${bs}"/><w:szCs w:val="${bs}"/>${lang}</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="${Math.round(240 * lineMult)}" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/>${A.rtlDoc ? '<w:pPr><w:bidi/></w:pPr>' : ''}</w:style>${hStyle(1)}${hStyle(2)}${hStyle(3)}<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="34"/><w:qFormat/></w:style><w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="39"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style></w:styles>`;
  const settings = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:zoom w:percent="100"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/><w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/></w:compat><w:themeFontLang w:val="en-US" w:bidi="he-IL"/></w:settings>`;
  // ---- header / footer with the page number in its place
  const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const pnXml = () => { const p = A.pn; const jc = p.align === 'center' ? 'center' : (p.align === 'right') === A.rtlDoc ? '' : 'end'; const sz = Math.round(p.size * 2);
    const rp = `<w:rPr>${opt.font === 'orig' ? `<w:rFonts w:ascii="${esc(p.face)}" w:hAnsi="${esc(p.face)}" w:cs="${esc(p.face)}"/>` : ''}${p.color !== '000000' && opt.colors ? `<w:color w:val="${p.color}"/>` : ''}<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr>`;
    const parts = p.words.split('#'); const out = [];
    parts.forEach((t, i) => { if (t) out.push(`<w:r>${rp}<w:t xml:space="preserve">${esc(t)}</w:t></w:r>`); if (i < parts.length - 1) out.push(`<w:r>${rp}<w:fldChar w:fldCharType="begin"/></w:r><w:r>${rp}<w:instrText xml:space="preserve"> ${i === 0 ? 'PAGE' : 'NUMPAGES'} </w:instrText></w:r><w:r>${rp}<w:fldChar w:fldCharType="separate"/></w:r><w:r>${rp}<w:t>1</w:t></w:r><w:r>${rp}<w:fldChar w:fldCharType="end"/></w:r>`); });
    return `<w:p><w:pPr>${A.rtlDoc ? '<w:bidi/>' : ''}<w:spacing w:before="0" w:after="0"/>${jc ? `<w:jc w:val="${jc}"/>` : ''}</w:pPr>${out.join('')}</w:p>`; };
  const hfPara = l => { const P = { lines: [l], dir: l.dir, pg: l.pg || pages[0] }; return `<w:p><w:pPr>${l.dir === 'R' ? '<w:bidi/>' : ''}<w:spacing w:before="0" w:after="0"/>${isCentered(l, A) ? '<w:jc w:val="center"/>' : ''}</w:pPr>${mergeRuns(l.runs).map(r => runXml(r, { ...ctx, dir: l.dir }, l.size)).join('')}</w:p>`; };
  const hdr = [...A.hf.header.map(hfPara), ...(A.pn && A.pn.top ? [pnXml()] : [])], ftr = [...A.hf.footer.map(hfPara), ...(A.pn && !A.pn.top ? [pnXml()] : [])];
  const hasH = hdr.length > 0, hasF = ftr.length > 0;
  if (hasH) z.file('word/header1.xml', orderProps(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr ${W_NS}>${hdr.join('')}</w:hdr>`));
  if (hasF) z.file('word/footer1.xml', orderProps(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr ${W_NS}>${ftr.join('')}</w:ftr>`));
  let doc = orderProps(body.join(''));
  if (!hasH) doc = doc.replace(/<w:headerReference[^>]*\/>/g, ''); if (!hasF) doc = doc.replace(/<w:footerReference[^>]*\/>/g, '');
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const title = String(name || '').replace(/\.docx$/i, '');
  z.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>${hasH ? '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' : ''}${hasF ? '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' : ''}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`);
  z.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>');
  z.file('docProps/core.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${esc(title)}</dc:title><dc:creator>סימן</dc:creator><cp:lastModifiedBy>סימן</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
  z.file('docProps/app.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>סימן ${VER}</Application></Properties>`);
  z.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>${hasH ? '<Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' : ''}${hasF ? '<Relationship Id="rIdFtr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' : ''}${rels.join('')}</Relationships>`);
  z.file('word/styles.xml', styles); z.file('word/numbering.xml', numbering); z.file('word/settings.xml', settings);
  z.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>${doc}</w:body></w:document>`);
  const bytes = await z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  return { bytes, stats };
}

/* ---------- the whole way, for code and for the window below ---------- */
async function convertWord(idx, opt, report, cancelled) {
  await loadOfficeLibs();
  const pages = await collect(idx, opt, report, cancelled);
  report({ phase: 'build' });
  const A = analyse(pages, opt);
  return { pages, A };
}
async function build(state, opt, name) {
  // the analysis is kept: changing an option only writes the file again
  const pages = state.pages.map(p => p); // flow building changes lines (markers), so work on copies
  const copy = pages.map(pg => ({ ...pg, lines: (pg.lines || []).map(l => ({ ...l, runs: l.runs.map(r => ({ ...r })) })) }));
  // table cells point at the original line objects: map them to the copies
  copy.forEach((pg, i) => { const map = new Map(); pages[i].lines && pages[i].lines.forEach((l, j) => map.set(l, copy[i].lines[j]));
    pg.tables = (pages[i].tables || []).map(t => ({ ...t, cells: t.cells.map(c => ({ ...c, lines: c.lines.map(l => map.get(l) || l) })) })); });
  const A = { ...state.A, hf: { header: state.A.hf.header.map(l => ({ ...l, runs: l.runs.map(r => ({ ...r })) })), footer: state.A.hf.footer.map(l => ({ ...l, runs: l.runs.map(r => ({ ...r })) })) } };
  return writeDocx(copy, A, opt, name);
}

/* =====================================================================
   4. THE WINDOW — choose, watch it happen, compare, keep
   ===================================================================== */
const CSS = `
.cvx{position:fixed;inset:0;z-index:2147482000;display:flex;align-items:center;justify-content:center;background:rgba(10,20,40,.46);backdrop-filter:blur(6px);-webkit-backdrop-filter:blur(6px);font-family:"SF-rubik",Rubik,"Noto Sans Arabic","Segoe UI",Arial,sans-serif;
  --c-s:var(--v7-surface,#fff);--c-i:var(--v7-ink,#13233f);--c-m:var(--v7-muted,#5d6b82);--c-l:var(--v7-line,#e2e7ef);--c-a:var(--v7-accent,#2f7de1);--c-o:var(--v7-on,#fff);--c-soft:var(--v7-soft,#eef4fd);--c-bg:var(--v7-bg,#f4f6fa)}
.cvx.hidden{display:none}
.cvx.busy{backdrop-filter:none;-webkit-backdrop-filter:none;background:rgba(10,20,40,.62)}
.cvx *{box-sizing:border-box}
.cvx-card{background:var(--c-s);color:var(--c-i);width:min(1080px,calc(100vw - 32px));height:min(760px,calc(100vh - 32px));border-radius:26px;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 30px 80px rgba(10,25,55,.35);border:1px solid var(--c-l)}
.cvx-top{display:flex;align-items:center;gap:12px;padding:18px 22px 12px;border-bottom:1px solid var(--c-l)}
.cvx-top h2{margin:0;font-size:21px;font-weight:700;letter-spacing:-.01em}
.cvx-file{color:var(--c-m);font-size:14px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;min-width:0}
.cvx-x{border:0;background:var(--c-soft);color:var(--c-i);width:40px;height:40px;border-radius:50%;font-size:20px;cursor:pointer;flex:none}
.cvx-x:focus-visible,.cvx button:focus-visible,.cvx input:focus-visible{outline:3px solid var(--c-a);outline-offset:2px}
.cvx-step{flex:1;min-height:0;display:none;flex-direction:column}
.cvx-step.on{display:flex}
.cvx-scroll{flex:1;min-height:0;overflow:auto;padding:18px 22px}
.cvx-foot{display:flex;align-items:center;gap:10px;padding:14px 22px calc(14px + env(safe-area-inset-bottom,0px));border-top:1px solid var(--c-l);flex-wrap:wrap}
.cvx-foot .grow{flex:1;min-width:120px;color:var(--c-m);font-size:14px}
.cvx-btn{border:1px solid var(--c-l);background:var(--c-s);color:var(--c-i);border-radius:14px;min-height:46px;padding:0 18px;font:inherit;font-size:15.5px;font-weight:600;cursor:pointer;display:inline-flex;align-items:center;gap:8px}
.cvx-btn.pri{background:var(--c-a);border-color:var(--c-a);color:var(--c-o)}
.cvx-btn:disabled{opacity:.5;cursor:default}
.cvx-lbl{font-size:13px;font-weight:700;color:var(--c-m);margin:0 0 8px;letter-spacing:.02em}
.cvx-chips{display:flex;gap:8px;flex-wrap:wrap;align-items:center;margin-bottom:20px}
.cvx-chip{border:1.5px solid var(--c-l);background:var(--c-s);color:var(--c-i);border-radius:999px;min-height:40px;padding:0 15px;font:inherit;font-size:14.5px;cursor:pointer}
.cvx-chip[aria-pressed=true]{border-color:var(--c-a);background:var(--c-soft);color:var(--c-a);font-weight:700}
.cvx-chip:disabled{opacity:.45;cursor:default}
.cvx-range{border:1.5px solid var(--c-l);border-radius:12px;min-height:40px;padding:0 12px;font:inherit;font-size:15px;width:170px;background:var(--c-s);color:var(--c-i)}
.cvx-grid{display:grid;grid-template-columns:repeat(5,1fr);gap:12px}
.cvx-fmt{position:relative;text-align:start;border:1.5px solid var(--c-l);background:var(--c-s);color:var(--c-i);border-radius:20px;padding:16px;cursor:pointer;font:inherit;display:flex;flex-direction:column;gap:8px;min-height:120px;transition:border-color .15s,transform .15s,box-shadow .15s}
.cvx-fmt:hover{transform:translateY(-1px);box-shadow:0 6px 18px rgba(20,40,80,.08)}
.cvx-fmt[aria-checked=true]{border-color:var(--c-a);box-shadow:0 0 0 3px color-mix(in srgb,var(--c-a) 22%,transparent)}
.cvx-fmt b{font-size:16.5px}
.cvx-fmt small{color:var(--c-m);font-size:13.5px;line-height:1.45}
.cvx-fmt .tag{display:inline-flex;align-items:center;justify-content:center;min-width:46px;height:26px;padding:0 8px;border-radius:8px;font-size:11.5px;font-weight:800;letter-spacing:.04em;color:#fff;align-self:flex-start;direction:ltr}
.tag.w{background:#2b5797}.tag.p{background:#c4472b}.tag.x{background:#1e7145}.tag.a{background:linear-gradient(135deg,#6a5cff,#2f7de1)}.tag.t{background:#6b778c}
.cvx-hero{grid-column:1/-1;flex-direction:row;align-items:stretch;gap:20px;padding:20px;min-height:0}
.cvx-hero .txt{flex:1;display:flex;flex-direction:column;gap:9px}
.cvx-hero b{font-size:20px}
.cvx-badge{position:absolute;top:14px;inset-inline-end:14px;background:var(--c-a);color:var(--c-o);font-size:12px;font-weight:700;border-radius:999px;padding:4px 10px}
.cvx-ticks{margin:2px 0 0;padding:0;list-style:none;display:grid;grid-template-columns:1fr 1fr;gap:6px 16px;font-size:14px}
.cvx-ticks li{display:flex;gap:7px;align-items:flex-start}
.cvx-ticks li::before{content:"✓";color:var(--c-a);font-weight:800}
.cvx-lock{font-size:13px;color:var(--c-m);margin-top:auto}
.cvx-art{flex:none;width:230px;display:flex;align-items:center;gap:8px}
.cvx-pg{width:96px;height:124px;border-radius:8px;background:#fff;border:1px solid #d7dde8;box-shadow:0 3px 10px rgba(0,0,0,.08);padding:10px 9px;display:flex;flex-direction:column;gap:5px;direction:rtl}
.cvx-pg i{display:block;height:5px;border-radius:3px;background:#c9d2df}
.cvx-pg.a i{outline:1px dashed #e58b8b;outline-offset:1px}
.cvx-pg.b i.h{height:7px;width:70%;background:#d23b3b}.cvx-pg.b i.h2{width:55%;background:#2f62d6}
.cvx-pg.b i.m{background:linear-gradient(90deg,#c9d2df 40%,#ffe94d 40%)}
.cvx-pg.b i.l{position:relative;margin-inline-start:9px}.cvx-pg.b i.l::after{content:"";position:absolute;inset-inline-start:-8px;top:0;width:5px;height:5px;border-radius:50%;background:#2f62d6}
.cvx-pg.b .tb{display:grid;grid-template-columns:1fr 1fr 1fr;gap:2px;margin-top:2px}.cvx-pg.b .tb span{height:9px;border:1px solid #9fb0c8}
.cvx-arrow{font-size:22px;color:var(--c-a)}
.cvx-more{margin-top:18px}
.cvx-work{flex:1;display:flex;flex-direction:column;align-items:center;gap:18px;padding:26px 22px;min-height:0}
.cvx-ring{--p:0;width:118px;height:118px;border-radius:50%;background:conic-gradient(var(--c-a) calc(var(--p)*1%),var(--c-soft) 0);display:grid;place-items:center;flex:none}
.cvx-ring span{width:94px;height:94px;border-radius:50%;background:var(--c-s);display:grid;place-items:center;font-size:24px;font-weight:700}
.cvx-wt{font-size:18px;font-weight:700;margin:0;text-align:center}
.cvx-ws{color:var(--c-m);font-size:14.5px;margin:-10px 0 0;text-align:center;min-height:20px}
.cvx-tiles{display:grid;grid-template-columns:repeat(auto-fill,minmax(78px,1fr));gap:10px;width:100%;overflow:auto;padding:4px;flex:1;min-height:0;align-content:start}
.cvx-tile{aspect-ratio:.72;border-radius:8px;background:var(--c-soft);border:1px solid var(--c-l);position:relative;overflow:hidden;background-size:cover;background-position:top center}
.cvx-tile.r::after{content:"";position:absolute;inset-inline:0;height:30%;top:-30%;background:linear-gradient(transparent,color-mix(in srgb,var(--c-a) 40%,transparent));animation:cvxscan 1.1s linear infinite}
.cvx-tile.ok::before{content:"✓";position:absolute;bottom:4px;inset-inline-start:4px;width:20px;height:20px;border-radius:50%;background:var(--c-a);color:var(--c-o);font-size:12px;font-weight:800;display:grid;place-items:center}
.cvx-tile b{position:absolute;top:3px;inset-inline-end:5px;font-size:10px;color:var(--c-m)}
@keyframes cvxscan{to{top:100%}}
@media (prefers-reduced-motion:reduce){.cvx-tile.r::after{animation:none;top:35%}.cvx-fmt{transition:none}}
.cvx-bar{display:flex;align-items:center;gap:10px;padding:12px 22px;border-bottom:1px solid var(--c-l);flex-wrap:wrap}
.cvx-stats{display:flex;gap:6px;flex-wrap:wrap;flex:1;min-width:0}
.cvx-stat{background:var(--c-soft);border-radius:10px;padding:6px 10px;font-size:13.5px}
.cvx-stat b{font-weight:800}
.cvx-seg{display:inline-flex;background:var(--c-soft);border-radius:12px;padding:3px}
.cvx-seg button{border:0;background:transparent;color:var(--c-i);font:inherit;font-size:14px;min-height:36px;padding:0 13px;border-radius:9px;cursor:pointer}
.cvx-seg button[aria-pressed=true]{background:var(--c-s);font-weight:700;box-shadow:0 1px 4px rgba(0,0,0,.12)}
.cvx-cmp{flex:1;min-height:0;display:grid;grid-template-columns:1fr 1fr;background:var(--c-bg)}
.cvx-cmp[data-v=orig]{grid-template-columns:1fr}.cvx-cmp[data-v=orig] .cvx-pw{display:none}
.cvx-cmp[data-v=word]{grid-template-columns:1fr}.cvx-cmp[data-v=word] .cvx-po{display:none}
.cvx-pane{overflow:auto;min-height:0;position:relative;padding:16px 12px 30px}
.cvx-pane+.cvx-pane{border-inline-start:1px solid var(--c-l)}
.cvx-ptag{position:sticky;top:0;z-index:2;display:inline-block;margin:-6px 0 8px;background:var(--c-s);border:1px solid var(--c-l);border-radius:999px;padding:4px 12px;font-size:12.5px;font-weight:700;color:var(--c-m)}
.cvx-po img{display:block;width:100%;max-width:640px;margin:0 auto 14px;background:#fff;box-shadow:0 2px 10px rgba(0,0,0,.12)}
.cvx-pw .doc{margin:0 auto}
.cvx-pw .doc section.cvxd{box-shadow:0 2px 10px rgba(0,0,0,.12)!important;margin:0 auto 14px!important;background:#fff!important;color:#000}
.cvx-pw .doc .cvxd-wrapper{background:transparent!important;padding:0!important}
.cvx-wait{color:var(--c-m);text-align:center;padding:40px 10px;font-size:14.5px}
.cvx-opts{position:relative}
.cvx-pop{position:absolute;top:calc(100% + 6px);inset-inline-end:0;background:var(--c-s);border:1px solid var(--c-l);border-radius:16px;box-shadow:0 14px 40px rgba(15,33,63,.22);padding:14px;width:290px;z-index:5;display:flex;flex-direction:column;gap:12px}
.cvx-pop.hidden{display:none}
.cvx-sw{display:flex;align-items:center;justify-content:space-between;gap:10px;font-size:14.5px;cursor:pointer}
.cvx-sw input{width:20px;height:20px;accent-color:var(--c-a)}
.cvx-pop select{font:inherit;border:1.5px solid var(--c-l);border-radius:10px;min-height:38px;padding:0 8px;background:var(--c-s);color:var(--c-i)}
@media (max-width:1000px){.cvx-grid{grid-template-columns:repeat(3,1fr)}}
@media (max-width:760px){
  .cvx{align-items:stretch}
  .cvx-card{width:100vw;height:100%;border-radius:0;border:0;padding-top:env(safe-area-inset-top,0px)}
  .cvx-grid{grid-template-columns:1fr 1fr}
  .cvx-hero{flex-direction:column}.cvx-art{width:auto;justify-content:center}.cvx-ticks{grid-template-columns:1fr}
  .cvx-cmp{grid-template-columns:1fr}.cvx-cmp[data-v=both] .cvx-po{display:none}
  .cvx-seg [data-v=both]{display:none}
  .cvx-foot{padding-inline-end:76px}
  .cvx-foot .cvx-btn{flex:1;justify-content:center;padding:0 10px;font-size:14.5px}.cvx-foot .grow{flex-basis:100%;font-size:13px}
  .cvx-foot .cvx-btn.pri{flex-basis:100%;order:-1}
  .cvx-bar{padding:10px 14px;gap:8px}.cvx-stats{flex-basis:100%;flex-wrap:nowrap;overflow-x:auto;scrollbar-width:none}.cvx-stat{white-space:nowrap}
  .cvx-scroll{padding:14px}.cvx-top{padding:12px 14px 10px}
}
html.cvx-on .sb-back{display:none!important}
.cvx-stat{white-space:nowrap}
`;
const FORMATS = [
  { k: 'word', tag: 'w', t: 'DOCX', b: 'Word לעריכה', s: '' },
  { k: 'exact', tag: 'w', t: 'DOCX', b: 'Word במראה זהה', s: 'כל שורה במקום המדויק שלה מעל תמונת העמוד. טוב לתיקון קטן בלי לזוז מהעימוד.' },
  { k: 'ai', tag: 'a', t: 'AI', b: 'Word בעזרת בינה מלאכותית', s: 'לעמודים מסובכים במיוחד. כל עמוד נשלח ל-Google ולוקח 10–30 שניות.' },
  { k: 'pptx', tag: 'p', t: 'PPTX', b: 'PowerPoint', s: 'כל עמוד הופך לשקף, והטקסט נפתח לעריכה.' },
  { k: 'xlsx', tag: 'x', t: 'XLSX', b: 'Excel', s: 'כל עמוד הופך לגיליון, ושורות הטבלאות מסודרות בעמודות.' },
  { k: 'text', tag: 't', t: 'TXT', b: 'Word, טקסט בלבד', s: 'רק הטקסט והטבלאות, בלי עיצוב. נוח להעתקה.' }
];
const OPT_KEY = 'siman.conv';
let ui = null, st = { fmt: 'word', scope: 'all', opt: { colors: true, pageBreaks: false, font: 'orig', ocr: true } };
try { const o = JSON.parse(localStorage.getItem(OPT_KEY) || '{}'); if (o.fmt) st.fmt = o.fmt; if (o.opt) st.opt = { ...st.opt, ...o.opt }; } catch (_) {}
const saveSt = () => { try { localStorage.setItem(OPT_KEY, JSON.stringify({ fmt: st.fmt, opt: st.opt })); } catch (_) {} };

function el(html) { const t = document.createElement('template'); t.innerHTML = html.trim(); return t.content.firstElementChild; }
function make() {
  if (ui) return ui;
  const style = document.createElement('style'); style.textContent = CSS; document.head.append(style);
  const root = el(`<div class="cvx hidden" id="cvx" role="dialog" aria-modal="true" aria-labelledby="cvxTitle" dir="rtl">
  <div class="cvx-card">
    <div class="cvx-top"><h2 id="cvxTitle">המרה לקובץ לעריכה</h2><span class="cvx-file" id="cvxFile"></span><button class="cvx-x" id="cvxX" aria-label="סגירה">✕</button></div>
    <section class="cvx-step" data-s="pick">
      <div class="cvx-scroll">
        <p class="cvx-lbl" id="cvxScopeL">אילו עמודים</p>
        <div class="cvx-chips" role="group" aria-labelledby="cvxScopeL">
          <button class="cvx-chip" data-sc="all"></button><button class="cvx-chip" data-sc="cur"></button><button class="cvx-chip" data-sc="sel">העמודים המסומנים</button><button class="cvx-chip" data-sc="range">עמודים לבחירה</button>
          <input class="cvx-range hidden" id="cvxRange" dir="ltr" inputmode="numeric" placeholder="1-3, 5" aria-label="מספרי עמודים" autocomplete="off">
        </div>
        <p class="cvx-lbl" id="cvxFmtL">לאיזה קובץ</p>
        <div class="cvx-grid" role="radiogroup" aria-labelledby="cvxFmtL" id="cvxFmts"></div>
      </div>
      <div class="cvx-foot"><span class="grow" id="cvxMsg" role="status"></span><button class="cvx-btn" id="cvxCancel">ביטול</button><button class="cvx-btn pri" id="cvxGo">המרה</button></div>
    </section>
    <section class="cvx-step" data-s="work">
      <div class="cvx-work">
        <div class="cvx-ring" id="cvxRing" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-label="התקדמות ההמרה"><span id="cvxPct">0%</span></div>
        <p class="cvx-wt" id="cvxWt">קורא את העמודים…</p><p class="cvx-ws" id="cvxWs" aria-live="polite"></p>
        <div class="cvx-tiles" id="cvxTiles"></div>
      </div>
      <div class="cvx-foot"><span class="grow">הכול קורה כאן במכשיר. הקובץ לא נשלח לשום מקום.</span><button class="cvx-btn" id="cvxStop">עצירה</button></div>
    </section>
    <section class="cvx-step" data-s="done">
      <div class="cvx-bar">
        <div class="cvx-stats" id="cvxStats"></div>
        <div class="cvx-seg" role="group" aria-label="תצוגה"><button data-v="both">זה לצד זה</button><button data-v="orig">המקור</button><button data-v="word">ה-Word</button></div>
        <div class="cvx-opts"><button class="cvx-btn" id="cvxOptB" aria-expanded="false" aria-controls="cvxPop">אפשרויות</button>
          <div class="cvx-pop hidden" id="cvxPop">
            <label class="cvx-sw">צבעים, מרקר ומילוי כמו במקור<input type="checkbox" data-o="colors"></label>
            <label class="cvx-sw">כל עמוד מתחיל בעמוד חדש<input type="checkbox" data-o="pageBreaks"></label>
            <label class="cvx-sw">גופן<select data-o="font"><option value="orig">כמו במקור</option><option value="Arial">Arial</option><option value="David">David</option><option value="Calibri">Calibri</option><option value="Times New Roman">Times New Roman</option></select></label>
          </div></div>
      </div>
      <div class="cvx-cmp" id="cvxCmp" data-v="both">
        <div class="cvx-pane cvx-po" id="cvxPO"><span class="cvx-ptag">המקור</span><div id="cvxPOi"></div></div>
        <div class="cvx-pane cvx-pw" id="cvxPW"><span class="cvx-ptag">ה-Word שיישמר</span><div class="doc" id="cvxDoc"></div></div>
      </div>
      <div class="cvx-foot"><span class="grow" id="cvxName"></span><button class="cvx-btn" id="cvxAgain">המרה אחרת</button><button class="cvx-btn hidden" id="cvxShare">שיתוף</button><button class="cvx-btn" id="cvxSave">שמירה</button><button class="cvx-btn pri" id="cvxOpen">פתיחה לעריכה בסימן</button></div>
    </section>
  </div></div>`);
  document.body.append(root);
  const q = id => root.querySelector('#' + id);
  // the format cards
  const fm = q('cvxFmts');
  for (const f of FORMATS) {
    const hero = f.k === 'word';
    const b = el(`<button class="cvx-fmt${hero ? ' cvx-hero' : ''}" role="radio" data-f="${f.k}"></button>`);
    if (hero) b.innerHTML = `<span class="cvx-badge">מומלץ</span><div class="txt"><span class="tag w">DOCX</span><b>Word לעריכה</b>
      <small>מסמך Word אמיתי, כאילו הוקלד מההתחלה: כותבים, מוחקים ומוסיפים, והטקסט זורם כמו שצריך.</small>
      <ul class="cvx-ticks"><li>פסקאות שממשיכות גם מעמוד לעמוד</li><li>כותרות שמופיעות בניווט של Word</li><li>רשימות ותבליטים של Word</li><li>טבלאות, תמונות ומספרי עמודים</li><li>צבעים, מרקר וקו תחתון כמו במקור</li><li>עמודים סרוקים: זיהוי טקסט</li></ul>
      <span class="cvx-lock">🔒 נבנה כאן במכשיר. הקובץ לא נשלח לשום מקום.</span></div>
      <div class="cvx-art" aria-hidden="true"><div class="cvx-pg a"><i></i><i style="width:80%"></i><i></i><i style="width:60%"></i><i></i><i style="width:75%"></i></div><span class="cvx-arrow">←</span><div class="cvx-pg b"><i class="h"></i><i class="h2"></i><i class="m"></i><i class="l"></i><i class="l" style="width:70%"></i><div class="tb"><span></span><span></span><span></span><span></span><span></span><span></span></div></div></div>`;
    else b.innerHTML = `<span class="tag ${f.tag}">${f.t}</span><b>${f.b}</b><small>${f.s}</small>`;
    b.onclick = () => { st.fmt = f.k; saveSt(); paintFmt(); };
    b.ondblclick = () => go();
    fm.append(b);
  }
  fm.addEventListener('keydown', e => { // arrow keys move between the cards, like radio buttons
    const list = [...fm.querySelectorAll('.cvx-fmt')]; const i = list.indexOf(document.activeElement); if (i < 0) return;
    const d = { ArrowLeft: 1, ArrowDown: 1, ArrowRight: -1, ArrowUp: -1 }[e.key]; if (!d) return; e.preventDefault();
    const n = list[(i + d + list.length) % list.length]; n.focus(); n.click(); });
  root.querySelectorAll('.cvx-chip').forEach(c => c.onclick = () => { st.scope = c.dataset.sc; paintScope(); if (st.scope === 'range') q('cvxRange').focus(); });
  q('cvxRange').addEventListener('keydown', e => { if (e.key === 'Enter') go(); });
  q('cvxX').onclick = q('cvxCancel').onclick = () => close();
  q('cvxGo').onclick = () => go();
  q('cvxStop').onclick = () => { if (job) job.stop = true; q('cvxWs').textContent = 'עוצר…'; };
  root.addEventListener('keydown', e => { if (e.key === 'Escape') { e.preventDefault(); if (!q('cvxPop').classList.contains('hidden')) { togglePop(false); return; } if (job && !job.done) { job.stop = true; return; } close(); } });
  root.addEventListener('click', e => { if (e.target === root && !(job && !job.done)) close(); });
  // results
  root.querySelectorAll('.cvx-seg button').forEach(b => b.onclick = () => setView(b.dataset.v));
  const togglePop = on => { q('cvxPop').classList.toggle('hidden', !on); q('cvxOptB').setAttribute('aria-expanded', on); };
  q('cvxOptB').onclick = () => togglePop(q('cvxPop').classList.contains('hidden'));
  document.addEventListener('pointerdown', e => { if (!q('cvxPop').classList.contains('hidden') && !e.target.closest('.cvx-opts')) togglePop(false); });
  root.querySelectorAll('[data-o]').forEach(i => i.onchange = () => { const k = i.dataset.o; st.opt[k] = i.type === 'checkbox' ? i.checked : i.value; saveSt(); rebuild(); });
  q('cvxAgain').onclick = () => show('pick');
  q('cvxSave').onclick = async () => { if (!result) return; if (await saveBytes(result.name, result.bytes, 'docx')) toast('נשמר: ' + result.name); };
  q('cvxOpen').onclick = async () => {
    if (!result) return; const f = new File([result.bytes], result.name, { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    close();
    // the person already chose to edit: answer the "what do we do with it" question for them
    const t0 = Date.now(); const tick = setInterval(() => { const d = $id('docModeDlg'); if (d && d.open) { clearInterval(tick); const b = d.querySelector('[data-mode=edit]'); if (b) b.click(); } else if (Date.now() - t0 > 20000) clearInterval(tick); }, 60);
    try { await openFile(f); } finally { setTimeout(() => clearInterval(tick), 3000); }
  };
  q('cvxShare').onclick = async () => { if (!result) return; const f = new File([result.bytes], result.name, { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    try { await navigator.share({ files: [f], title: result.name }); } catch (e) { if (!(e && e.name === 'AbortError')) toast('השיתוף לא הצליח. אפשר לשמור את הקובץ ולשלוח אותו'); } };
  // scroll the two sides together, page by page
  let lock = false; const po = q('cvxPO'), pw = q('cvxPW');
  const sync = (a, b) => () => { if (lock || q('cvxCmp').dataset.v !== 'both') return; lock = true; const r = a.scrollTop / Math.max(1, a.scrollHeight - a.clientHeight); b.scrollTop = r * (b.scrollHeight - b.clientHeight); requestAnimationFrame(() => lock = false); };
  po.addEventListener('scroll', sync(po, pw), { passive: true }); pw.addEventListener('scroll', sync(pw, po), { passive: true });
  try { new ResizeObserver(() => fitDoc()).observe(pw); } catch (_) { window.addEventListener('resize', () => fitDoc()); }
  ui = { root, q, togglePop }; return ui;
}
function paintFmt() { ui.root.querySelectorAll('.cvx-fmt').forEach(b => { const on = b.dataset.f === st.fmt; b.setAttribute('aria-checked', on); b.tabIndex = on ? 0 : -1; }); ui.q('cvxGo').textContent = st.fmt === 'word' ? 'המרה ל-Word' : 'המרה'; }
function paintScope() {
  ui.root.querySelectorAll('.cvx-chip').forEach(c => c.setAttribute('aria-pressed', c.dataset.sc === st.scope));
  ui.q('cvxRange').classList.toggle('hidden', st.scope !== 'range'); ui.q('cvxMsg').textContent = '';
}
function show(step) { ui.root.querySelectorAll('.cvx-step').forEach(s => s.classList.toggle('on', s.dataset.s === step)); ui.step = step; ui.root.classList.toggle('busy', step === 'work');
  ui.q('cvxTitle').textContent = step === 'done' ? 'ה-Word מוכן' : step === 'work' ? 'ממיר ל-Word' : 'המרה לקובץ לעריכה';
  const f = step === 'pick' ? ui.root.querySelector('.cvx-fmt[aria-checked=true]') : step === 'work' ? ui.q('cvxStop') : ui.q('cvxOpen'); setTimeout(() => f && f.focus(), 30); }
let lastFocus = null;
function open() {
  if (!S.doc) return; make(); lastFocus = document.activeElement;
  const n = S.pages.length, cur = currentPageIndex(), hasSel = S.sel && S.sel.size;
  ui.q('cvxFile').textContent = (S.name || '').replace(/\.pdf$/i, '');
  const chip = sc => ui.root.querySelector(`[data-sc=${sc}]`);
  chip('all').textContent = n === 1 ? 'כל המסמך (עמוד אחד)' : `כל המסמך (${n} עמודים)`;
  chip('cur').textContent = `העמוד הנוכחי (${cur + 1})`; chip('cur').disabled = n === 1;
  chip('sel').disabled = !hasSel; if (hasSel) chip('sel').textContent = `המסומנים (${S.sel.size})`;
  st.scope = hasSel ? 'sel' : 'all'; paintScope(); paintFmt();
  ui.root.querySelectorAll('[data-o]').forEach(i => { if (i.type === 'checkbox') i.checked = !!st.opt[i.dataset.o]; else i.value = st.opt[i.dataset.o]; });
  ui.root.classList.remove('hidden'); document.documentElement.style.overflow = 'hidden'; document.documentElement.classList.add('cvx-on'); show('pick');
}
function close() { if (!ui) return; if (job && !job.done) job.stop = true; ui.root.classList.add('hidden'); document.documentElement.style.overflow = ''; document.documentElement.classList.remove('cvx-on'); ui.togglePop(false); if (lastFocus && lastFocus.focus) try { lastFocus.focus(); } catch (_) {} }
function indices() {
  const n = S.pages.length;
  if (st.scope === 'cur') return [currentPageIndex()];
  if (st.scope === 'sel') return S.sel && S.sel.size ? [...S.sel].sort((a, b) => a - b) : null;
  if (st.scope === 'range') return parsePages(ui.q('cvxRange').value, n);
  return S.pages.map((_, i) => i);
}
const OLD = { exact: 'docx', ai: 'docxsmart', pptx: 'pptx', xlsx: 'xlsx', text: 'docxflow' };
let job = null, result = null, state = null;
async function go() {
  const idx = indices();
  if (!idx) { ui.q('cvxMsg').textContent = st.scope === 'sel' ? 'לא סומנו עמודים בחלונית "עמודים".' : 'כתבו מספרי עמודים, למשל 1-3, 5'; (st.scope === 'range' ? ui.q('cvxRange') : ui.q('cvxMsg')).focus && ui.q('cvxRange').focus(); return; }
  if (st.fmt !== 'word') { // the other kinds go the way they always went
    const sc = $id('convScope'); if (sc) { sc.value = st.scope === 'range' ? 'range' : st.scope; const r = $id('convRange'); if (r) r.value = ui.q('cvxRange').value; if (st.scope === 'sel' || st.scope === 'cur' || st.scope === 'all') sc.value = st.scope; }
    close(); convertTo(OLD[st.fmt]); return;
  }
  runWord(idx);
}
function setPct(p, label, sub) { const r = ui.q('cvxRing'); p = clamp(Math.round(p), 0, 100); r.style.setProperty('--p', p); r.setAttribute('aria-valuenow', p); ui.q('cvxPct').textContent = p + '%'; if (label != null) ui.q('cvxWt').textContent = label; if (sub != null) ui.q('cvxWs').textContent = sub; }
async function runWord(idx) {
  show('work'); result = null; state = null;
  const tiles = ui.q('cvxTiles'); tiles.innerHTML = '';
  const tl = new Map(); idx.forEach((i, n) => { const t = el(`<div class="cvx-tile" aria-hidden="true"><b>${i + 1}</b></div>`); tiles.append(t); tl.set(i, t); });
  const me = job = { stop: false, done: false }; const t0 = Date.now();
  setPct(0, 'מכין את הקובץ…', 'רגע אחד');
  const report = e => {
    if (me !== job) return; const N = e.total || idx.length; const t = tl.get(e.page);
    if (e.phase === 'read') { setPct(e.n / N * 55, 'קורא את העמודים…', `עמוד ${e.page + 1} · ${e.n + 1} מתוך ${N}`); if (t) { t.classList.add('r'); t.scrollIntoView({ block: 'nearest' }); } }
    else if (e.phase === 'ocr') setPct(e.n / N * 55, 'מזהה טקסט בעמוד סרוק…', `עמוד ${e.page + 1} · בפעם הראשונה זה לוקח קצת יותר`);
    else if (e.phase === 'thumb') { if (t) t.style.backgroundImage = `url(${e.thumb})`; }
    else if (e.phase === 'shape') { setPct(55 + e.n / N * 40, 'מזהה כותרות, רשימות, טבלאות ותמונות…', `עמוד ${e.page + 1} · ${e.n + 1} מתוך ${N}`); if (t) { t.classList.remove('r'); t.classList.add('ok'); } }
    else if (e.phase === 'build') setPct(96, 'בונה את קובץ ה-Word…', '');
  };
  try {
    state = await convertWord(idx, st.opt, report, () => me.stop);
    const part = idx.length === S.pages.length ? '' : ` - ${idx.length === 1 ? 'עמוד' : 'עמודים'} ${rangeLabel(idx)}`;
    state.name = `${(S.name || 'מסמך').replace(/\.pdf$/i, '')}${part}.docx`; state.idx = idx;
    await rebuild(true);
    me.done = true; setPct(100);
    try { if (typeof count === 'function') count('conv_word'); } catch (_) {}
    show('done');
    const secs = Math.round((Date.now() - t0) / 1000); ui.q('cvxName').textContent = `${state.name.replace(/\.docx$/i, '')} · מוכן תוך ${secs} שניות`;
  } catch (e) {
    me.done = true;
    if (e && e.cancelled) { show('pick'); ui.q('cvxMsg').textContent = 'ההמרה נעצרה.'; return; }
    console.error(e); show('pick'); ui.q('cvxMsg').textContent = 'ההמרה נכשלה: ' + (e && e.message ? e.message : String(e)) + '. שלחו לי צילום של ההודעה הזו';
  }
}
let rebuilding = null;
async function rebuild(first) {
  if (!state) return; const my = rebuilding = {};
  const r = await build(state, { ...st.opt }, state.name); if (my !== rebuilding) return;
  result = { bytes: r.bytes, name: state.name, stats: r.stats };
  const s = r.stats, chip = (n, t) => n ? `<span class="cvx-stat"><b>${n}</b> ${t}</span>` : '';
  ui.q('cvxStats').innerHTML = chip(s.pages, s.pages === 1 ? 'עמוד' : 'עמודים') + chip(s.p, 'פסקאות') + chip(s.h, 'כותרות') + chip(s.list, 'פריטי רשימה') + chip(s.table, s.table === 1 ? 'טבלה' : 'טבלאות') + chip(s.img, s.img === 1 ? 'תמונה' : 'תמונות');
  try { const f = new File([r.bytes], state.name); ui.q('cvxShare').classList.toggle('hidden', !(navigator.share && navigator.canShare && navigator.canShare({ files: [f] }))); } catch (_) {}
  if (first) { ui.q('cvxPOi').innerHTML = state.pages.map(p => `<img src="${p.thumb}" alt="עמוד ${p.idx + 1} במקור" loading="lazy">`).join(''); setView(window.matchMedia('(max-width:760px)').matches ? 'word' : 'both'); }
  await preview(r.bytes);
}
async function preview(bytes) {
  const host = ui.q('cvxDoc'); host.innerHTML = '<p class="cvx-wait">מכין תצוגה…</p>';
  try {
    if (!window.docx) await loadScript('docx-preview.min.js');
    const box = document.createElement('div');
    await window.docx.renderAsync(bytes, box, null, { className: 'cvxd', inWrapper: true, breakPages: true, ignoreLastRenderedPageBreak: true, experimental: true, renderHeaders: true, renderFooters: true, useBase64URL: true });
    box.querySelectorAll('section.cvxd').forEach(sec => sec.setAttribute('dir', 'rtl'));
    host.innerHTML = ''; host.append(box); fitDoc();
  } catch (e) { console.warn(e); host.innerHTML = '<p class="cvx-wait">אין תצוגה מקדימה במכשיר הזה, אבל הקובץ מוכן לשמירה ולפתיחה.</p>'; }
}
function fitDoc() {
  if (!ui) return; const host = ui.q('cvxDoc'), pane = ui.q('cvxPW'), box = host.firstElementChild; const sec = host.querySelector('section'); if (!sec || !box) return;
  box.style.transform = ''; const w = sec.offsetWidth || 794, room = pane.clientWidth - 24; if (room <= 0) return; const z = Math.min(1, room / w);
  box.style.width = w + 'px'; box.style.transformOrigin = '0 0'; box.style.transform = `scale(${z})`;
  host.style.direction = 'ltr'; host.style.width = Math.round(w * z) + 'px'; host.style.height = Math.round(box.offsetHeight * z) + 'px'; host.style.overflow = 'hidden';
}
function setView(v) { ui.q('cvxCmp').dataset.v = v; ui.root.querySelectorAll('.cvx-seg button').forEach(b => b.setAttribute('aria-pressed', b.dataset.v === v)); requestAnimationFrame(fitDoc); }

// the convert button (and everything that "clicks" it: the command search, the phone menu) opens the new window
function hook() { const b = $id('convBtn'); if (!b) return; b.onclick = () => { if (!S.doc || S.office) return; open(); }; }
if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', hook); else hook();

window.SIMAN_CONV = { VER, convertWord, build, collect, analyse, writeDocx, niceColor, open, close };
})();
