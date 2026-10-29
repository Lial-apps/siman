/* =====================================================================
   סימן — PowerPoint editor (ppt-editor.js)
   A .pptx opens for real editing, like in PowerPoint: slides panel, text with fonts, sizes,
   colors, alignment, right-to-left, bullets; moving and resizing shapes and pictures; text boxes,
   pictures, shapes and tables; adding, duplicating, deleting and reordering slides; slide show;
   saving as PowerPoint or PDF.
   The file is edited in place: every slide stays the original XML with only the changed parts
   rewritten, so layouts, themes, animations and notes are kept.
   Needs: JSZip, html2canvas (PDF), PptxGenJS (a new, empty presentation).
   API: window.PPTX_EDITOR = { open(opts), close(), isOpen(), toPdf(bytes) }
   ===================================================================== */
(() => {
'use strict';
const NS = {
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  pr: 'http://schemas.openxmlformats.org/package/2006/relationships',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
};
const RT = {
  slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  layout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
  master: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
  theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  notes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
};
const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
const CT_SLIDE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const PT = 12700;                       // EMU per point
const HEB = /[\u0590-\u05FF\uFB1D-\uFB4F\u0600-\u06FF\u0750-\u077F]/;
const enc = new TextEncoder(), dec = new TextDecoder('utf-8');
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
const parseXml = s => { const d = new DOMParser().parseFromString(s, 'application/xml'); if (d.getElementsByTagName('parsererror').length) throw new Error('XML פגום בקובץ'); return d; };
const ser = d => new XMLSerializer().serializeToString(d);
const kids = (el, name) => el ? [...el.childNodes].filter(n => n.nodeType === 1 && (!name || n.localName === name)) : [];
const kid = (el, name) => kids(el, name)[0] || null;
const path = (el, ...names) => { for (const n of names) { el = kid(el, n); if (!el) return null; } return el; };
const at = (el, n) => el ? el.getAttribute(n) : null;
const num = (v, d = 0) => { const n = parseFloat(v); return isFinite(n) ? n : d; };
const nextFrame = () => new Promise(r => requestAnimationFrame(() => setTimeout(r, 0)));
const aNew = (doc, name, attrs) => { const e = doc.createElementNS(NS.a, 'a:' + name); if (attrs) for (const k in attrs) e.setAttribute(k, String(attrs[k])); return e; };
const pNew = (doc, name, attrs) => { const e = doc.createElementNS(NS.p, 'p:' + name); if (attrs) for (const k in attrs) e.setAttribute(k, String(attrs[k])); return e; };

/* ------------------------------------------------------------ the package */
const isXmlPart = n => /\.(xml|rels)$/i.test(n);
async function readPkg(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const files = new Map();
  for (const name of Object.keys(zip.files)) {
    const f = zip.files[name]; if (f.dir) continue;
    files.set(name, isXmlPart(name) ? await f.async('string') : await f.async('uint8array'));
  }
  if (!files.has('ppt/presentation.xml')) throw new Error('זה לא קובץ PowerPoint תקין');
  return files;
}
const relsPathOf = part => { const i = part.lastIndexOf('/'); return part.slice(0, i + 1) + '_rels/' + part.slice(i + 1) + '.rels'; };
function resolve(part, target) {
  if (/^\//.test(target)) return target.slice(1);
  const base = part.slice(0, part.lastIndexOf('/') + 1).split('/').filter(Boolean);
  for (const seg of target.split('/')) { if (seg === '..') base.pop(); else if (seg !== '.') base.push(seg); }
  return base.join('/');
}
function relTarget(fromPart, toPart) { // a relative target from one part to another
  const a = fromPart.split('/').slice(0, -1), b = toPart.split('/');
  let i = 0; while (i < a.length && a[i] === b[i]) i++;
  return '../'.repeat(a.length - i) + b.slice(i).join('/');
}
function rels(files, part) {
  const m = new Map(); const s = files.get(relsPathOf(part)); if (!s) return m;
  const d = parseXml(s);
  for (const r of d.getElementsByTagNameNS(NS.pr, 'Relationship')) {
    const mode = at(r, 'TargetMode');
    m.set(at(r, 'Id'), { type: at(r, 'Type'), target: mode === 'External' ? at(r, 'Target') : resolve(part, at(r, 'Target')), raw: at(r, 'Target'), external: mode === 'External' });
  }
  return m;
}
function addRel(files, part, type, target) {
  const rp = relsPathOf(part);
  const d = parseXml(files.get(rp) || `${XML_HEAD}<Relationships xmlns="${NS.pr}"></Relationships>`);
  const ids = new Set([...d.getElementsByTagNameNS(NS.pr, 'Relationship')].map(r => at(r, 'Id')));
  let n = 1; while (ids.has('rId' + n)) n++;
  const r = d.createElementNS(NS.pr, 'Relationship');
  r.setAttribute('Id', 'rId' + n); r.setAttribute('Type', type); r.setAttribute('Target', relTarget(part, target));
  d.documentElement.append(r); files.set(rp, ser(d));
  return 'rId' + n;
}
function ctEnsure(files, kind, key, mime) {
  const d = parseXml(files.get('[Content_Types].xml'));
  const tag = kind === 'Default' ? 'Extension' : 'PartName';
  if ([...d.getElementsByTagNameNS(NS.ct, kind)].some(e => (at(e, tag) || '').toLowerCase() === key.toLowerCase())) return;
  const e = d.createElementNS(NS.ct, kind); e.setAttribute(tag, key); e.setAttribute('ContentType', mime);
  d.documentElement.append(e); files.set('[Content_Types].xml', ser(d));
}
function ctDrop(files, part) {
  const d = parseXml(files.get('[Content_Types].xml'));
  for (const e of [...d.getElementsByTagNameNS(NS.ct, 'Override')]) if (at(e, 'PartName') === '/' + part) e.remove();
  files.set('[Content_Types].xml', ser(d));
}

/* ------------------------------------------------------------ the presentation */
// the slides in their order, from presentation.xml
function slideList(files) {
  const d = parseXml(files.get('ppt/presentation.xml')), R = rels(files, 'ppt/presentation.xml');
  const out = [];
  for (const s of d.getElementsByTagNameNS(NS.p, 'sldId')) { const r = R.get(s.getAttributeNS(NS.r, 'id') || at(s, 'r:id')); if (r && files.has(r.target)) out.push(r.target); }
  return out;
}
function slideSize(files) {
  const d = parseXml(files.get('ppt/presentation.xml')), s = d.getElementsByTagNameNS(NS.p, 'sldSz')[0];
  return { w: num(at(s, 'cx'), 12192000) / PT, h: num(at(s, 'cy'), 6858000) / PT };
}
// writes the slide order back into presentation.xml (and its relationships) after adding, deleting or moving slides
function writeOrder(files, order) {
  const P = 'ppt/presentation.xml', d = parseXml(files.get(P));
  const rp = relsPathOf(P), rd = parseXml(files.get(rp));
  const relEls = [...rd.getElementsByTagNameNS(NS.pr, 'Relationship')];
  const byTarget = new Map();
  for (const r of relEls) if (at(r, 'Type') === RT.slide) byTarget.set(resolve(P, at(r, 'Target')), r);
  // drop relationships of slides that are gone, add for new ones
  for (const [t, r] of byTarget) if (!order.includes(t)) { r.remove(); byTarget.delete(t); }
  const ids = new Set(relEls.map(r => at(r, 'Id')));
  for (const t of order) if (!byTarget.has(t)) {
    let n = 1; while (ids.has('rId' + n)) n++; ids.add('rId' + n);
    const r = rd.createElementNS(NS.pr, 'Relationship'); r.setAttribute('Id', 'rId' + n); r.setAttribute('Type', RT.slide); r.setAttribute('Target', relTarget(P, t));
    rd.documentElement.append(r); byTarget.set(t, r);
  }
  files.set(rp, ser(rd));
  let lst = d.getElementsByTagNameNS(NS.p, 'sldIdLst')[0];
  const old = new Map();
  if (lst) { for (const s of [...lst.childNodes]) if (s.nodeType === 1) old.set(s.getAttributeNS(NS.r, 'id') || at(s, 'r:id'), at(s, 'id')); }
  else { lst = pNew(d, 'sldIdLst'); const after = d.getElementsByTagNameNS(NS.p, 'sldMasterIdLst')[0]; after ? after.after(lst) : d.documentElement.prepend(lst); }
  let max = 255; for (const v of old.values()) max = Math.max(max, +v || 0);
  while (lst.firstChild) lst.firstChild.remove();
  for (const t of order) {
    const rid = at(byTarget.get(t), 'Id');
    const e = pNew(d, 'sldId', { id: old.get(rid) || ++max }); e.setAttributeNS(NS.r, 'r:id', rid); lst.append(e);
  }
  files.set(P, ser(d));
}
function newSlidePath(files) { let n = 1; while (files.has(`ppt/slides/slide${n}.xml`)) n++; return `ppt/slides/slide${n}.xml`; }

/* ------------------------------------------------------------ theme, master, layout */
// everything a slide draws from: its layout, master and theme (parsed once)
function ctxOf(X, slidePath) {
  const files = X.files, cache = X.parts;
  const doc = p => { if (!cache.has(p)) cache.set(p, files.has(p) ? parseXml(files.get(p)) : null); return cache.get(p); };
  const sR = rels(files, slidePath);
  const layoutPath = [...sR.values()].find(r => r.type === RT.layout)?.target;
  const lR = layoutPath ? rels(files, layoutPath) : new Map();
  const masterPath = [...lR.values()].find(r => r.type === RT.master)?.target;
  const mR = masterPath ? rels(files, masterPath) : new Map();
  const themePath = [...mR.values()].find(r => r.type === RT.theme)?.target;
  const key = masterPath + '|' + themePath;
  if (!X.themes.has(key)) {
    const th = themePath && doc(themePath), m = masterPath && doc(masterPath);
    const colors = {}, fonts = { mj: { latin: 'Calibri', cs: 'Arial' }, mn: { latin: 'Calibri', cs: 'Arial' } };
    if (th) {
      const cs = th.getElementsByTagNameNS(NS.a, 'clrScheme')[0];
      for (const c of kids(cs)) { const v = kid(c, 'srgbClr') || kid(c, 'sysClr'); colors[c.localName] = v ? (at(v, 'val') === 'windowText' ? '000000' : at(v, 'val') === 'window' ? 'FFFFFF' : at(v, 'lastClr') || at(v, 'val')) : '000000'; }
      for (const [k, tag] of [['mj', 'majorFont'], ['mn', 'minorFont']]) {
        const f = th.getElementsByTagNameNS(NS.a, tag)[0];
        if (f) { fonts[k].latin = at(kid(f, 'latin'), 'typeface') || fonts[k].latin; const c = at(kid(f, 'cs'), 'typeface'); fonts[k].cs = c || fonts[k].latin; const heb = kids(f, 'font').find(x => at(x, 'script') === 'Hebr'); if (heb && !c) fonts[k].cs = at(heb, 'typeface'); }
      }
    }
    const map = {}; const cm = m && m.getElementsByTagNameNS(NS.p, 'clrMap')[0];
    for (const k of ['bg1', 'tx1', 'bg2', 'tx2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']) map[k] = cm ? at(cm, k) || k : ({ bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' })[k] || k;
    X.themes.set(key, { colors, fonts, map });
  }
  return { slidePath, layoutPath, masterPath, layout: layoutPath && doc(layoutPath), master: masterPath && doc(masterPath), sR, lR, mR, theme: X.themes.get(key), pres: doc('ppt/presentation.xml') };
}

/* ------------------------------------------------------------ colors */
function hexRgb(h) { h = String(h || '000000').replace('#', ''); if (h.length === 3) h = h.split('').map(c => c + c).join(''); const n = parseInt(h, 16) || 0; return [n >> 16 & 255, n >> 8 & 255, n & 255]; }
function rgbHsl([r, g, b]) { r /= 255; g /= 255; b /= 255; const mx = Math.max(r, g, b), mn = Math.min(r, g, b); let h = 0, s = 0; const l = (mx + mn) / 2; if (mx !== mn) { const d = mx - mn; s = l > .5 ? d / (2 - mx - mn) : d / (mx + mn); h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; h /= 6; } return [h, s, l]; }
function hslRgb([h, s, l]) { if (!s) { const v = Math.round(l * 255); return [v, v, v]; } const f = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; }; const q = l < .5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q; return [f(p, q, h + 1 / 3), f(p, q, h), f(p, q, h - 1 / 3)].map(v => Math.round(v * 255)); }
const PRST = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080' };
// a:solidFill / a:fgClr / … → CSS color; ph = the placeholder color (phClr) of a style reference
function colorOf(el, ctx, ph) {
  if (!el) return null;
  const c = kids(el).find(n => /Clr$/.test(n.localName)) || (/Clr$/.test(el.localName) ? el : null);
  if (!c) return null;
  let rgb, alpha = 1;
  const T = ctx.theme;
  if (c.localName === 'srgbClr') rgb = hexRgb(at(c, 'val'));
  else if (c.localName === 'schemeClr') { let v = at(c, 'val'); if (v === 'phClr') return ph || null; v = T.map[v] || v; rgb = hexRgb(T.colors[v] || '000000'); }
  else if (c.localName === 'sysClr') rgb = hexRgb(at(c, 'lastClr') || (at(c, 'val') === 'window' ? 'FFFFFF' : '000000'));
  else if (c.localName === 'prstClr') rgb = hexRgb(PRST[at(c, 'val')] || '000000');
  else if (c.localName === 'scrgbClr') rgb = ['r', 'g', 'b'].map(k => Math.round(num(at(c, k)) / 100000 * 255));
  else return null;
  let hsl = null;
  for (const m of kids(c)) {
    const v = num(at(m, 'val')) / 100000;
    if (m.localName === 'alpha') alpha = v;
    else if (m.localName === 'lumMod') { hsl = hsl || rgbHsl(rgb); hsl[2] *= v; rgb = hslRgb(hsl); }
    else if (m.localName === 'lumOff') { hsl = hsl || rgbHsl(rgb); hsl[2] = Math.min(1, hsl[2] + v); rgb = hslRgb(hsl); }
    else if (m.localName === 'tint') rgb = rgb.map(x => Math.round(x + (255 - x) * (1 - v)));
    else if (m.localName === 'shade') rgb = rgb.map(x => Math.round(x * v));
    hsl = null;
  }
  return alpha < 1 ? `rgba(${rgb.join(',')},${+alpha.toFixed(3)})` : `rgb(${rgb.join(',')})`;
}
const cssHex = css => { const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(css || ''); return m ? [m[1], m[2], m[3]].map(v => (+v).toString(16).padStart(2, '0')).join('').toUpperCase() : String(css || '').replace('#', '').toUpperCase(); };

/* ------------------------------------------------------------ placeholders */
const phOf = sp => { const nv = kids(sp).find(n => /^nv/.test(n.localName)); return nv ? path(nv, 'nvPr', 'ph') : null; };
const phType = ph => { const t = at(ph, 'type') || 'body'; return t === 'ctrTitle' ? 'title' : t === 'subTitle' || t === 'obj' ? 'body' : t; };
function findPh(doc, ph) {
  if (!doc || !ph) return null;
  const tree = doc.getElementsByTagNameNS(NS.p, 'spTree')[0]; if (!tree) return null;
  const all = [...tree.getElementsByTagNameNS(NS.p, 'sp')].filter(phOf);
  const idx = at(ph, 'idx'), t = at(ph, 'type') || 'body';
  return (idx != null && all.find(s => at(phOf(s), 'idx') === idx)) || all.find(s => (at(phOf(s), 'type') || 'body') === t) || all.find(s => phType(phOf(s)) === phType(ph)) || null;
}
// the shape on the layout and on the master that a placeholder inherits from
function inherit(sp, ctx) { const ph = phOf(sp); if (!ph) return []; const l = findPh(ctx.layout, ph); const m = findPh(ctx.master, l ? phOf(l) : ph); return [l, m].filter(Boolean); }
function xfrmOf(sp, ctx) {
  for (const s of [sp, ...inherit(sp, ctx)]) {
    const pr = kid(s, 'spPr') || kid(s, 'grpSpPr'); const x = pr ? kid(pr, 'xfrm') : kid(s, 'xfrm');
    if (x && kid(x, 'off') && kid(x, 'ext')) return x;
  }
  return null;
}
function box(x) { // a:xfrm → points
  const o = kid(x, 'off'), e = kid(x, 'ext');
  return { x: num(at(o, 'x')) / PT, y: num(at(o, 'y')) / PT, w: num(at(e, 'cx')) / PT, h: num(at(e, 'cy')) / PT, rot: num(at(x, 'rot')) / 60000, fh: at(x, 'flipH') === '1', fv: at(x, 'flipV') === '1' };
}

/* ------------------------------------------------------------ text styles */
// the list styles a paragraph inherits from, nearest first
function textSources(sp, ctx) {
  const out = [];
  const tb = s => kid(s, 'txBody');
  const own = tb(sp); if (own && kid(own, 'lstStyle')) out.push(kid(own, 'lstStyle'));
  const ph = phOf(sp);
  for (const s of inherit(sp, ctx)) { const t = tb(s); if (t && kid(t, 'lstStyle')) out.push(kid(t, 'lstStyle')); }
  const tx = ctx.master && ctx.master.getElementsByTagNameNS(NS.p, 'txStyles')[0];
  if (ph) { const t = phType(ph); const st = t === 'title' ? kid(tx, 'titleStyle') : ['body'].includes(t) ? kid(tx, 'bodyStyle') : kid(tx, 'otherStyle'); if (st) out.push(st); }
  else { const d = ctx.pres && ctx.pres.getElementsByTagNameNS(NS.p, 'defaultTextStyle')[0]; if (d) out.push(d); const o = kid(tx, 'otherStyle'); if (o) out.push(o); }
  return out;
}
function bodyPrOf(sp, ctx) { return [sp, ...inherit(sp, ctx)].map(s => path(s, 'txBody', 'bodyPr')).filter(Boolean); }
const lvlPr = (sources, lvl) => sources.map(s => kid(s, `lvl${lvl + 1}pPr`)).filter(Boolean);
// first value of a paragraph attribute along the chain
function pAttr(pPr, lv, name) { if (pPr && pPr.hasAttribute(name)) return at(pPr, name); for (const l of lv) if (l.hasAttribute(name)) return at(l, name); return null; }
function pKid(pPr, lv, name) { const k = pPr && kid(pPr, name); if (k) return k; for (const l of lv) { const x = kid(l, name); if (x) return x; } return null; }
function rAttr(rPr, lv, name) { if (rPr && rPr.hasAttribute(name)) return at(rPr, name); for (const l of lv) { const d = kid(l, 'defRPr'); if (d && d.hasAttribute(name)) return at(d, name); } return null; }
function rKid(rPr, lv, name) { const k = rPr && kid(rPr, name); if (k) return k; for (const l of lv) { const d = kid(l, 'defRPr'); const x = d && kid(d, name); if (x) return x; } return null; }
function fontName(tf, ctx, cs) {
  if (!tf) return null; const F = ctx.theme.fonts;
  if (tf === '+mj-lt') return F.mj.latin; if (tf === '+mn-lt') return F.mn.latin; if (tf === '+mj-cs') return F.mj.cs; if (tf === '+mn-cs') return F.mn.cs;
  if (tf === '+mj-ea') return F.mj.latin; if (tf === '+mn-ea') return F.mn.latin;
  return tf;
}
const cssFont = (...names) => [...new Set(names.filter(Boolean))].map(n => `"${n}"`).join(',') + ',"Heebo","Arial",sans-serif';

/* ------------------------------------------------------------ pictures */
const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', tif: 'image/tiff', tiff: 'image/tiff' };
function mediaUrl(X, target) {
  if (!target || !X.files.has(target)) return null;
  if (X.urls.has(target)) return X.urls.get(target);
  const ext = (target.split('.').pop() || '').toLowerCase(); const type = MIME[ext];
  if (!type) { X.urls.set(target, null); return null; }   // EMF/WMF: browsers cannot draw them
  const u = URL.createObjectURL(new Blob([X.files.get(target)], { type })); X.urls.set(target, u); return u;
}
const relsFor = (ctx, doc) => doc === ctx.layout ? [ctx.lR, ctx.layoutPath] : doc === ctx.master ? [ctx.mR, ctx.masterPath] : [ctx.sR, ctx.slidePath];
function blipUrl(X, blipFill, R) {
  const b = blipFill && kid(blipFill, 'blip'); if (!b) return null;
  const id = b.getAttributeNS(NS.r, 'embed') || at(b, 'r:embed'); const r = id && R.get(id);
  return r && !r.external ? mediaUrl(X, r.target) : null;
}

/* ------------------------------------------------------------ fills and lines */
function fillOf(pr, ctx, X, R, ph) {
  if (!pr) return null;
  for (const n of kids(pr)) {
    if (n.localName === 'noFill') return { none: true };
    if (n.localName === 'solidFill') return { color: colorOf(n, ctx, ph) };
    if (n.localName === 'gradFill') {
      const stops = kids(kid(n, 'gsLst'), 'gs').map(g => ({ pos: num(at(g, 'pos')) / 1000, c: colorOf(g, ctx, ph) })).filter(s => s.c);
      if (!stops.length) return null;
      const lin = kid(n, 'lin'); const ang = lin ? num(at(lin, 'ang')) / 60000 : 90;
      return { color: stops[0].c, grad: `linear-gradient(${ang + 90}deg,${stops.map(s => `${s.c} ${s.pos}%`).join(',')})` };
    }
    if (n.localName === 'blipFill') return { img: blipUrl(X, n, R) };
    if (n.localName === 'pattFill') return { color: colorOf(kid(n, 'fgClr'), ctx, ph) };
    if (n.localName === 'grpFill') return { group: true };
  }
  return null;
}
// a shape's style (p:style): the theme color its fill / line / text use when the shape does not set one
function styleRef(sp, name, ctx) { const st = kid(sp, 'style'); const r = st && kid(st, name); if (!r) return null; return { idx: num(at(r, 'idx')), color: colorOf(r, ctx) }; }
const DASH = { dash: '4 3', sysDash: '3 1', dot: '1 2', sysDot: '1 1', lgDash: '8 3', dashDot: '4 3 1 3', lgDashDot: '8 3 1 3' };
function lineOf(sp, ctx, ph) {
  const pr = kid(sp, 'spPr'); const ln = pr && kid(pr, 'ln');
  const ref = styleRef(sp, 'lnRef', ctx);
  let color = null, w = ln && ln.hasAttribute('w') ? num(at(ln, 'w')) / PT : (ref && ref.idx ? .75 * ref.idx : 0.75);
  if (ln && kid(ln, 'noFill')) return null;
  if (ln && kid(ln, 'solidFill')) color = colorOf(kid(ln, 'solidFill'), ctx, ph);
  else if (ln && kid(ln, 'gradFill')) { const g = kid(kid(kid(ln, 'gradFill'), 'gsLst'), 'gs'); color = colorOf(g, ctx, ph); }
  else if (ref && ref.idx) color = ref.color;
  if (!color) return null;
  const d = ln && kid(ln, 'prstDash');
  return { color, w: Math.max(.25, w), dash: d ? DASH[at(d, 'val')] || null : null, head: ln && kid(ln, 'headEnd'), tail: ln && kid(ln, 'tailEnd') };
}

/* ------------------------------------------------------------ geometry */
function geomPath(sp, w, h) {
  const pr = kid(sp, 'spPr'); const pg = pr && kid(pr, 'prstGeom'), cg = pr && kid(pr, 'custGeom');
  if (cg) {
    let d = '';
    for (const p of kids(kid(cg, 'pathLst'), 'path')) {
      const pw = num(at(p, 'w'), 0) || w * PT, ph = num(at(p, 'h'), 0) || h * PT, sx = w / pw, sy = h / ph;
      const pt = e => `${(num(at(e, 'x')) * sx).toFixed(2)} ${(num(at(e, 'y')) * sy).toFixed(2)}`;
      let cur = [0, 0];
      for (const c of kids(p)) {
        const ps = kids(c, 'pt');
        if (c.localName === 'moveTo') { d += `M${pt(ps[0])} `; cur = [num(at(ps[0], 'x')) * sx, num(at(ps[0], 'y')) * sy]; }
        else if (c.localName === 'lnTo') { d += `L${pt(ps[0])} `; cur = [num(at(ps[0], 'x')) * sx, num(at(ps[0], 'y')) * sy]; }
        else if (c.localName === 'cubicBezTo') { d += `C${ps.map(pt).join(' ')} `; const l = ps[ps.length - 1]; cur = [num(at(l, 'x')) * sx, num(at(l, 'y')) * sy]; }
        else if (c.localName === 'quadBezTo') { d += `Q${ps.map(pt).join(' ')} `; }
        else if (c.localName === 'close') d += 'Z ';
        else if (c.localName === 'arcTo') { // an approximation: a straight line to the end of the arc
          const wR = num(at(c, 'wR')) * sx, hR = num(at(c, 'hR')) * sy, st = num(at(c, 'stAng')) / 60000 * Math.PI / 180, sw = num(at(c, 'swAng')) / 60000 * Math.PI / 180;
          const cx = cur[0] - wR * Math.cos(st), cy = cur[1] - hR * Math.sin(st), ex = cx + wR * Math.cos(st + sw), ey = cy + hR * Math.sin(st + sw);
          d += `A${wR.toFixed(2)} ${hR.toFixed(2)} 0 ${Math.abs(sw) > Math.PI ? 1 : 0} ${sw > 0 ? 1 : 0} ${ex.toFixed(2)} ${ey.toFixed(2)} `; cur = [ex, ey];
        }
      }
    }
    return d || null;
  }
  const t = pg ? at(pg, 'prst') : 'rect';
  const adj = (n, def) => { const g = pg && kids(kid(pg, 'avLst'), 'gd').find(x => at(x, 'name') === n); const m = g && /val (-?\d+)/.exec(at(g, 'fmla') || ''); return m ? +m[1] / 100000 : def; };
  const s = Math.min(w, h);
  switch (t) {
    case 'ellipse': return `M0 ${h / 2}A${w / 2} ${h / 2} 0 1 0 ${w} ${h / 2}A${w / 2} ${h / 2} 0 1 0 0 ${h / 2}Z`;
    case 'roundRect': { const r = s * adj('adj', .16667); return `M${r} 0H${w - r}Q${w} 0 ${w} ${r}V${h - r}Q${w} ${h} ${w - r} ${h}H${r}Q0 ${h} 0 ${h - r}V${r}Q0 0 ${r} 0Z`; }
    case 'triangle': { const a = w * adj('adj', .5); return `M${a} 0L${w} ${h}H0Z`; }
    case 'rtTriangle': return `M0 0L${w} ${h}H0Z`;
    case 'diamond': return `M${w / 2} 0L${w} ${h / 2}L${w / 2} ${h}L0 ${h / 2}Z`;
    case 'parallelogram': { const a = s * adj('adj', .25); return `M${a} 0H${w}L${w - a} ${h}H0Z`; }
    case 'trapezoid': { const a = s * adj('adj', .25); return `M${a} 0H${w - a}L${w} ${h}H0Z`; }
    case 'hexagon': { const a = s * adj('adj', .25); return `M${a} 0H${w - a}L${w} ${h / 2}L${w - a} ${h}H${a}L0 ${h / 2}Z`; }
    case 'pentagon': return `M0 0H${w - h / 2}L${w} ${h / 2}L${w - h / 2} ${h}H0Z`;
    case 'homePlate': { const a = s * adj('adj', .5); return `M0 0H${w - a}L${w} ${h / 2}L${w - a} ${h}H0Z`; }
    case 'chevron': { const a = s * adj('adj', .5); return `M0 0H${w - a}L${w} ${h / 2}L${w - a} ${h}H0L${a} ${h / 2}Z`; }
    case 'rightArrow': { const a1 = h * adj('adj1', .5) / 2, a2 = s * adj('adj2', .5); return `M0 ${h / 2 - a1}H${w - a2}V0L${w} ${h / 2}L${w - a2} ${h}V${h / 2 + a1}H0Z`; }
    case 'leftArrow': { const a1 = h * adj('adj1', .5) / 2, a2 = s * adj('adj2', .5); return `M${w} ${h / 2 - a1}H${a2}V0L0 ${h / 2}L${a2} ${h}V${h / 2 + a1}H${w}Z`; }
    case 'upArrow': { const a1 = w * adj('adj1', .5) / 2, a2 = s * adj('adj2', .5); return `M${w / 2 - a1} ${h}V${a2}H0L${w / 2} 0L${w} ${a2}H${w / 2 + a1}V${h}Z`; }
    case 'downArrow': { const a1 = w * adj('adj1', .5) / 2, a2 = s * adj('adj2', .5); return `M${w / 2 - a1} 0V${h - a2}H0L${w / 2} ${h}L${w} ${h - a2}H${w / 2 + a1}V0Z`; }
    case 'star5': { let d = ''; for (let i = 0; i < 10; i++) { const r = i % 2 ? .38 : .5, a = -Math.PI / 2 + i * Math.PI / 5; d += `${i ? 'L' : 'M'}${(w / 2 + Math.cos(a) * w * r).toFixed(2)} ${(h / 2 + Math.sin(a) * h * r * 1.05).toFixed(2)}`; } return d + 'Z'; }
    case 'plus': { const a = s * adj('adj', .25); return `M${a} 0H${w - a}V${a}H${w}V${h - a}H${w - a}V${h}H${a}V${h - a}H0V${a}H${a}Z`; }
    case 'line': case 'straightConnector1': case 'bentConnector3': case 'curvedConnector3': return `M0 0L${w} ${h}`;
    case 'wedgeRectCallout': case 'wedgeRoundRectCallout': return `M0 0H${w}V${h * .8}H${w * .35}L${w * .2} ${h}V${h * .8}H0Z`;
    case 'cloud': case 'cloudCallout': return `M${w * .2} ${h * .35}Q${w * .2} ${h * .05} ${w * .45} ${h * .12}Q${w * .65} 0 ${w * .78} ${h * .2}Q${w} ${h * .25} ${w * .92} ${h * .55}Q${w} ${h * .9} ${w * .7} ${h * .85}Q${w * .55} ${h} ${w * .38} ${h * .88}Q${w * .1} ${h * .95} ${w * .1} ${h * .7}Q0 ${h * .5} ${w * .2} ${h * .35}Z`;
    case 'flowChartTerminator': { const r = h / 2; return `M${r} 0H${w - r}A${r} ${r} 0 0 1 ${w - r} ${h}H${r}A${r} ${r} 0 0 1 ${r} 0Z`; }
    case 'snip1Rect': { const a = s * adj('adj', .16667); return `M0 0H${w - a}L${w} ${a}V${h}H0Z`; }
    default: return `M0 0H${w}V${h}H0Z`;
  }
}
const isLine = sp => { const pg = path(sp, 'spPr', 'prstGeom'); return sp.localName === 'cxnSp' || (pg && /^(line|straightConnector1|bentConnector\d|curvedConnector\d)$/.test(at(pg, 'prst'))); };

/* ------------------------------------------------------------ text */
const ALIGN = { l: 'left', r: 'right', ctr: 'center', just: 'justify', dist: 'justify', justLow: 'justify' };
function autoNum(type, n) {
  const t = type || 'arabicPeriod';
  const alpha = (k, up) => { let s = ''; k++; while (k > 0) { k--; s = String.fromCharCode((up ? 65 : 97) + k % 26) + s; k = Math.floor(k / 26); } return s; };
  const roman = (k, up) => { const r = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]; let s = '', v = k + 1; for (const [x, c] of r) while (v >= x) { s += c; v -= x; } return up ? s.toUpperCase() : s; };
  const heb = k => 'אבגדהוזחטיכלמנסעפצקרשת'[k % 22] || String(k + 1);
  const core = /^alphaLc/.test(t) ? alpha(n, 0) : /^alphaUc/.test(t) ? alpha(n, 1) : /^romanLc/.test(t) ? roman(n, 0) : /^romanUc/.test(t) ? roman(n, 1) : /^hebrew/.test(t) ? heb(n) : String(n + 1);
  return /ParenBoth$/.test(t) ? `(${core})` : /ParenR$/.test(t) ? `${core})` : /Plain$/.test(t) ? core : `${core}.`;
}
// runs → spans; level defaults go on the paragraph, what the run itself sets goes on its span
function runCss(rPr, ctx, scale, title) {
  const st = {}; if (!rPr) return st;
  if (rPr.hasAttribute('sz')) st.fontSize = (num(at(rPr, 'sz')) / 100 * scale) + 'pt';
  if (rPr.hasAttribute('b')) st.fontWeight = at(rPr, 'b') === '1' ? '700' : '400';
  if (rPr.hasAttribute('i')) st.fontStyle = at(rPr, 'i') === '1' ? 'italic' : 'normal';
  const dec = [];
  if (rPr.hasAttribute('u')) dec.push(at(rPr, 'u') !== 'none' ? 'underline' : '');
  if (rPr.hasAttribute('strike')) dec.push(at(rPr, 'strike') !== 'noStrike' ? 'line-through' : '');
  if (rPr.hasAttribute('u') || rPr.hasAttribute('strike')) st.textDecoration = dec.filter(Boolean).join(' ') || 'none';
  const f = kid(rPr, 'solidFill'); if (f) st.color = colorOf(f, ctx);
  const lat = kid(rPr, 'latin'), cs = kid(rPr, 'cs');
  if (lat || cs) st.fontFamily = cssFont(fontName(at(cs, 'typeface'), ctx, 1), fontName(at(lat, 'typeface'), ctx));
  if (rPr.hasAttribute('baseline')) { const b = num(at(rPr, 'baseline')); if (b) { st.verticalAlign = b > 0 ? 'super' : 'sub'; st.fontSize = st.fontSize || '.7em'; } }
  const hl = kid(rPr, 'highlight'); if (hl) st.backgroundColor = colorOf(hl, ctx);
  return st;
}
function levelCss(lv, ctx, scale, title, refColor) {
  const st = {};
  const sz = rAttr(null, lv, 'sz'); st.fontSize = ((sz ? num(sz) : 1800) / 100 * scale) + 'pt';
  st.fontWeight = rAttr(null, lv, 'b') === '1' ? '700' : '400';
  st.fontStyle = rAttr(null, lv, 'i') === '1' ? 'italic' : 'normal';
  // a shape's own text color (its style) comes before the presentation's default list style
  const f = rKid(null, lv, 'solidFill'); st.color = refColor || (f && colorOf(f, ctx)) || colorOf(aColor('tx1'), ctx);
  const F = ctx.theme.fonts[title ? 'mj' : 'mn'];
  st.fontFamily = cssFont(fontName(at(rKid(null, lv, 'cs'), 'typeface'), ctx) || F.cs, fontName(at(rKid(null, lv, 'latin'), 'typeface'), ctx) || F.latin);
  return st;
}
const _cdoc = document.implementation.createDocument(NS.a, 'a:x', null);
const aColor = v => { const s = _cdoc.createElementNS(NS.a, 'a:solidFill'); const c = _cdoc.createElementNS(NS.a, 'a:schemeClr'); c.setAttribute('val', v); s.append(c); return s; };

function renderText(X, ctx, sp, txBody, host, o) {
  const src = o.sources || textSources(sp, ctx);
  const bp = o.bodyPrs || bodyPrOf(sp, ctx);
  const bAttr = n => { for (const b of bp) if (b.hasAttribute(n)) return at(b, n); return null; };
  const auto = bp.length && kid(bp[0], 'normAutofit');
  const scale = auto && auto.hasAttribute('fontScale') ? num(at(auto, 'fontScale')) / 100000 : 1;
  const lnRed = auto && auto.hasAttribute('lnSpcReduction') ? num(at(auto, 'lnSpcReduction')) / 100000 : 0;
  const ins = (n, d) => { const v = bAttr(n); return (v != null ? num(v) : d) / PT; };
  const tx = document.createElement('div'); tx.className = 'pe-tx';
  const anchor = bAttr('anchor') || 't';
  Object.assign(tx.style, { paddingLeft: ins('lIns', 91440) + 'pt', paddingRight: ins('rIns', 91440) + 'pt', paddingTop: ins('tIns', 45720) + 'pt', paddingBottom: ins('bIns', 45720) + 'pt',
    justifyContent: anchor === 'ctr' ? 'center' : anchor === 'b' ? 'flex-end' : 'flex-start' });
  if (bAttr('wrap') === 'none') tx.classList.add('nowrap');
  const vert = bAttr('vert'); if (vert && /vert/.test(vert) && vert !== 'horz') tx.classList.add('vert');
  const inner = document.createElement('div'); inner.className = 'pe-txi'; tx.append(inner);
  const ph = phOf(sp), title = ph && phType(ph) === 'title';
  const fref = styleRef(sp, 'fontRef', ctx);
  let numbering = {}, allEmpty = true;
  kids(txBody, 'p').forEach((p, pi) => {
    const pPr = kid(p, 'pPr'); const lvl = num(at(pPr, 'lvl')); const lv = lvlPr(src, lvl);
    const d = document.createElement('div'); d.className = 'pp'; d.dataset.pi = pi;
    const algn = pAttr(pPr, lv, 'algn') || 'l'; d.style.textAlign = ALIGN[algn] || 'left';
    const rtl = pAttr(pPr, lv, 'rtl') === '1'; d.dir = rtl ? 'rtl' : 'ltr';
    Object.assign(d.style, levelCss(lv, ctx, scale, title, fref && fref.color));
    const marL = num(pAttr(pPr, lv, 'marL')) / PT, ind = num(pAttr(pPr, lv, 'indent')) / PT;
    if (marL) d.style.paddingInlineStart = marL + 'pt';
    if (ind) d.style.textIndent = ind + 'pt';
    if (ind < 0) { d.classList.add('hang'); d.style.setProperty('--bw', (-ind) + 'pt'); }
    const ls = pKid(pPr, lv, 'lnSpc');
    if (ls) { const pc = path(ls, 'spcPct'), pp = path(ls, 'spcPts'); if (pc) d.style.lineHeight = ((num(at(pc, 'val')) / 100000 - lnRed) * 1.2).toFixed(3); else if (pp) d.style.lineHeight = num(at(pp, 'val')) / 100 + 'pt'; }
    else if (lnRed) d.style.lineHeight = (1.2 * (1 - lnRed)).toFixed(3);
    for (const [k, css] of [['spcBef', 'marginTop'], ['spcAft', 'marginBottom']]) { const sv = pKid(pPr, lv, k); if (!sv) continue; const pp = path(sv, 'spcPts'), pc = path(sv, 'spcPct'); if (pp) d.style[css] = num(at(pp, 'val')) / 100 + 'pt'; else if (pc) d.style[css] = (num(at(pc, 'val')) / 100000) + 'em'; }
    // bullet
    const bNone = pKid(pPr, lv, 'buNone'), bChar = pKid(pPr, lv, 'buChar'), bAuto = pKid(pPr, lv, 'buAutoNum');
    const first = [pPr, ...lv].find(x => x && (kid(x, 'buNone') || kid(x, 'buChar') || kid(x, 'buAutoNum')));
    const which = first ? (kid(first, 'buNone') ? 'none' : kid(first, 'buChar') ? 'char' : 'auto') : 'none';
    void bNone;
    if (which === 'char') d.dataset.bu = at(bChar, 'char') || '•';
    else if (which === 'auto') { const t = at(bAuto, 'type'); const k = t + ':' + lvl; numbering[k] = (numbering[k] ?? (num(at(bAuto, 'startAt'), 1) - 1)) + 1; d.dataset.bu = autoNum(t, numbering[k] - 1); }
    if (which !== 'auto') for (const k in numbering) if (k.endsWith(':' + lvl)) delete numbering[k];
    if (d.dataset.bu) { const bc = pKid(pPr, lv, 'buClr'); if (bc) d.style.setProperty('--bu', colorOf(bc, ctx)); const bf = pKid(pPr, lv, 'buFont'); if (bf && /wingdings|symbol/i.test(at(bf, 'typeface') || '')) d.dataset.bu = '•'; }
    let any = false;
    for (const n of kids(p)) {
      if (n.localName === 'r' || n.localName === 'fld') {
        const rPr = kid(n, 'rPr'); let t = (kid(n, 't') || {}).textContent || '';
        if (n.localName === 'fld' && /slidenum/.test(at(n, 'type') || '')) t = String((o.index ?? 0) + 1);
        const s = document.createElement('span'); s.textContent = t; Object.assign(s.style, runCss(rPr, ctx, scale, title));
        if (rPr && rPr.hasAttribute('sz')) s.dataset.sz = at(rPr, 'sz');
        if (n.localName === 'fld') { s.dataset.fld = at(n, 'type') || ''; s.contentEditable = 'false'; }
        if (o.rmap) o.rmap.set(s, rPr ? rPr.cloneNode(true) : null);
        d.append(s); if (t) any = true;
      } else if (n.localName === 'br') d.append(document.createElement('br'));
    }
    if (!any) { d.classList.add('empty'); const e = kid(p, 'endParaRPr'); if (e && e.hasAttribute('sz')) d.style.fontSize = (num(at(e, 'sz')) / 100 * scale) + 'pt'; if (!d.lastChild || d.lastChild.nodeName !== 'BR') d.append(document.createElement('br')); }
    else allEmpty = false;
    inner.append(d);
  });
  if (!inner.firstChild) { const d = document.createElement('div'); d.className = 'pp empty'; d.dataset.pi = 0; Object.assign(d.style, levelCss(lvlPr(src, 0), ctx, scale, title, fref && fref.color)); const a = pAttr(null, lvlPr(src, 0), 'algn'); d.style.textAlign = ALIGN[a] || 'left'; d.dir = pAttr(null, lvlPr(src, 0), 'rtl') === '1' ? 'rtl' : 'ltr'; d.append(document.createElement('br')); inner.append(d); }
  if (allEmpty) { tx.classList.add('blank'); if (ph) tx.dataset.hint = ({ title: 'לחצו כאן להוספת כותרת', body: at(ph, 'type') === 'subTitle' ? 'לחצו כאן להוספת כותרת משנה' : 'לחצו כאן להוספת טקסט' })[phType(ph)] || ''; }
  host.append(tx);
  return tx;
}

/* ------------------------------------------------------------ one slide */
// renders a slide into a box of w × h points. o.edit: the editor (shapes can be selected, text edited)
function renderSlide(X, slidePath, o = {}) {
  const ctx = ctxOf(X, slidePath);
  const doc = o.doc || parseXml(X.files.get(slidePath));
  const { w, h } = X.size;
  const el = document.createElement('div'); el.className = 'pe-slide'; Object.assign(el.style, { width: w + 'pt', height: h + 'pt' });
  // background: the slide's own, else the layout's, else the master's
  let bg = null;
  for (const d of [doc, ctx.layout, ctx.master]) {
    const b = d && d.getElementsByTagNameNS(NS.p, 'bg')[0]; if (!b) continue;
    const [R] = relsFor(ctx, d);
    const pr = kid(b, 'bgPr'); if (pr) { bg = fillOf(pr, ctx, X, R); if (bg) break; }
    const ref = kid(b, 'bgRef'); if (ref) { bg = { color: colorOf(ref, ctx) }; break; }
  }
  if (bg && bg.img) Object.assign(el.style, { backgroundImage: `url("${bg.img}")`, backgroundSize: '100% 100%' });
  else if (bg && bg.grad) el.style.backgroundImage = bg.grad;
  el.style.backgroundColor = bg && bg.color ? bg.color : '#fff';
  const showM = d => !d || at(d.documentElement, 'showMasterSp') !== '0';
  const layers = [];
  if (showM(doc) && showM(ctx.layout) && ctx.master) layers.push([ctx.master, false, false]);
  if (showM(doc) && ctx.layout) layers.push([ctx.layout, false, false]);
  layers.push([doc, !!o.edit, true]);
  for (const [d, live, own] of layers) {
    const tree = d.getElementsByTagNameNS(NS.p, 'spTree')[0]; if (!tree) continue;
    const [R] = relsFor(ctx, d);
    const layer = document.createElement('div'); layer.className = live ? 'pe-layer live' : 'pe-layer'; el.append(layer);
    for (const n of kids(tree)) {
      if (!own && n.localName === 'sp' && phOf(n)) continue;          // placeholders of the layout and master are only drawn through the slide
      drawNode(X, ctx, n, layer, null, { ...o, live, R, top: n, doc: d });
    }
  }
  return el;
}
// group transform: child coordinates → slide points
function groupMap(g, parentMap) {
  const x = path(g, 'grpSpPr', 'xfrm'); if (!x) return parentMap;
  const off = kid(x, 'off'), ext = kid(x, 'ext'), co = kid(x, 'chOff'), ce = kid(x, 'chExt');
  const ox = num(at(off, 'x')), oy = num(at(off, 'y')), ex = num(at(ext, 'cx')), ey = num(at(ext, 'cy'));
  const cx = num(at(co, 'x')), cy = num(at(co, 'y')), cex = num(at(ce, 'cx')) || ex || 1, cey = num(at(ce, 'cy')) || ey || 1;
  const sx = ex / cex || 1, sy = ey / cey || 1;
  const m = b => ({ ...b, x: (ox + (b.x * PT - cx) * sx) / PT, y: (oy + (b.y * PT - cy) * sy) / PT, w: b.w * sx, h: b.h * sy });
  return parentMap ? b => parentMap(m(b)) : m;
}
function place(div, b) {
  Object.assign(div.style, { left: b.x + 'pt', top: b.y + 'pt', width: Math.max(0, b.w) + 'pt', height: Math.max(0, b.h) + 'pt' });
  const t = []; if (b.rot) t.push(`rotate(${b.rot}deg)`); if (b.fh) t.push('scaleX(-1)'); if (b.fv) t.push('scaleY(-1)');
  if (t.length) div.style.transform = t.join(' ');
}
function drawNode(X, ctx, n, parent, map, o) {
  const ln = n.localName;
  if (ln === 'AlternateContent') { const f = kids(n).find(c => c.localName === 'Fallback') || kids(n).find(c => c.localName === 'Choice'); for (const c of kids(f)) drawNode(X, ctx, c, parent, map, o); return; }
  if (ln === 'grpSp') { const m = groupMap(n, map); for (const c of kids(n)) if (/^(sp|pic|grpSp|graphicFrame|cxnSp|AlternateContent)$/.test(c.localName)) drawNode(X, ctx, c, parent, m, o); return; }
  if (!/^(sp|pic|graphicFrame|cxnSp)$/.test(ln)) return;
  const x = ln === 'graphicFrame' ? kid(n, 'xfrm') : xfrmOf(n, ctx);
  if (!x) return;
  let b = box(x); if (map) b = { ...map(b), rot: b.rot, fh: b.fh, fv: b.fv };
  const div = document.createElement('div'); div.className = 'pe-sp'; place(div, b);
  if (o.live) { X.dom.set(div, { node: n, top: o.top }); div.dataset.k = X.keyOf(o.top); }
  parent.append(div);
  const R = o.R;
  if (ln === 'pic') {
    const bf = kid(n, 'blipFill'); const u = blipUrl(X, bf, R);
    if (u) {
      const img = document.createElement('img'); img.src = u; img.alt = ''; img.draggable = false;
      const sr = bf && kid(bf, 'srcRect');
      if (sr) { const l = num(at(sr, 'l')) / 1e5, t = num(at(sr, 't')) / 1e5, r = num(at(sr, 'r')) / 1e5, bo = num(at(sr, 'b')) / 1e5; const W = 1 - l - r || 1, H = 1 - t - bo || 1;
        Object.assign(img.style, { position: 'absolute', width: (100 / W) + '%', height: (100 / H) + '%', left: (-l / W * 100) + '%', top: (-t / H * 100) + '%', maxWidth: 'none' }); div.style.overflow = 'hidden'; }
      else img.className = 'fill';
      div.append(img);
    } else { div.classList.add('pe-missing'); div.textContent = 'תמונה'; }
    const L = lineOf(n, ctx); if (L) div.style.outline = `${L.w}pt solid ${L.color}`;
    return;
  }
  if (ln === 'graphicFrame') {
    const tbl = n.getElementsByTagNameNS(NS.a, 'tbl')[0];
    if (tbl) drawTable(X, ctx, n, tbl, div, b, o);
    else { const img = n.getElementsByTagNameNS(NS.a, 'blip')[0]; const u = img && blipUrl(X, img.parentNode, R);
      if (u) { const i = document.createElement('img'); i.src = u; i.className = 'fill'; i.alt = ''; div.append(i); }
      else { div.classList.add('pe-missing'); div.textContent = /chart/.test(n.innerHTML || ser(n)) ? 'תרשים' : 'אובייקט'; } }
    return;
  }
  // a shape: geometry, fill, line, text
  const pr = kid(n, 'spPr');
  let fill = fillOf(pr, ctx, X, R);
  if (!fill) for (const s of inherit(n, ctx)) { fill = fillOf(kid(s, 'spPr'), ctx, X, R); if (fill) break; }
  if (!fill) { const fr = styleRef(n, 'fillRef', ctx); if (fr && fr.idx) fill = { color: fr.color }; }
  const line = lineOf(n, ctx);
  const d = b.w > 0 || b.h > 0 ? geomPath(n, Math.max(b.w, .01), Math.max(b.h, .01)) : null;
  if (d && ((fill && !fill.none && (fill.color || fill.img || fill.grad)) || line)) {
    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.setAttribute('class', 'pe-geo'); svg.setAttribute('viewBox', `0 0 ${Math.max(b.w, .01)} ${Math.max(b.h, .01)}`); svg.setAttribute('preserveAspectRatio', 'none');
    const pth = document.createElementNS('http://www.w3.org/2000/svg', 'path'); pth.setAttribute('d', d);
    if (fill && fill.img) {
      const id = 'pf' + (++X.uid); const defs = document.createElementNS(svg.namespaceURI, 'defs'); const pat = document.createElementNS(svg.namespaceURI, 'pattern');
      pat.setAttribute('id', id); pat.setAttribute('patternUnits', 'objectBoundingBox'); pat.setAttribute('width', '1'); pat.setAttribute('height', '1');
      const im = document.createElementNS(svg.namespaceURI, 'image'); im.setAttribute('href', fill.img); im.setAttribute('width', b.w); im.setAttribute('height', b.h); im.setAttribute('preserveAspectRatio', 'xMidYMid slice');
      pat.append(im); defs.append(pat); svg.append(defs); pth.setAttribute('fill', `url(#${id})`);
    } else if (fill && fill.grad && !isLine(n)) {
      const id = 'pg' + (++X.uid); const defs = document.createElementNS(svg.namespaceURI, 'defs'); const lg = document.createElementNS(svg.namespaceURI, 'linearGradient'); lg.setAttribute('id', id);
      const gf = pr && kid(pr, 'gradFill'); const lin = gf && kid(gf, 'lin'); const ang = (lin ? num(at(lin, 'ang')) / 60000 : 90) * Math.PI / 180;
      lg.setAttribute('x1', .5 - Math.cos(ang) / 2); lg.setAttribute('y1', .5 - Math.sin(ang) / 2); lg.setAttribute('x2', .5 + Math.cos(ang) / 2); lg.setAttribute('y2', .5 + Math.sin(ang) / 2);
      for (const g of kids(kid(gf, 'gsLst'), 'gs')) { const s = document.createElementNS(svg.namespaceURI, 'stop'); s.setAttribute('offset', num(at(g, 'pos')) / 1000 + '%'); s.setAttribute('stop-color', colorOf(g, ctx) || '#fff'); lg.append(s); }
      defs.append(lg); svg.append(defs); pth.setAttribute('fill', `url(#${id})`);
    } else pth.setAttribute('fill', !isLine(n) && fill && !fill.none && fill.color ? fill.color : 'none');
    if (line) { pth.setAttribute('stroke', line.color); pth.setAttribute('stroke-width', line.w); pth.setAttribute('vector-effect', 'non-scaling-stroke'); if (line.dash) pth.setAttribute('stroke-dasharray', line.dash.split(' ').map(v => v * line.w).join(' '));
      if (isLine(n) && (line.tail || line.head)) {
        const mk = (end, id) => { const t = end && at(end, 'type'); if (!t || t === 'none') return null; const m = document.createElementNS(svg.namespaceURI, 'marker'); m.setAttribute('id', id); m.setAttribute('viewBox', '0 0 10 10'); m.setAttribute('refX', '8'); m.setAttribute('refY', '5'); m.setAttribute('markerWidth', '5'); m.setAttribute('markerHeight', '5'); m.setAttribute('orient', 'auto-start-reverse'); const a = document.createElementNS(svg.namespaceURI, 'path'); a.setAttribute('d', 'M0 0L10 5L0 10Z'); a.setAttribute('fill', line.color); m.append(a); return m; };
        const defs = document.createElementNS(svg.namespaceURI, 'defs'); const t1 = mk(line.tail, 'pt' + (++X.uid)), h1 = mk(line.head, 'ph' + (++X.uid));
        if (t1) { defs.append(t1); pth.setAttribute('marker-end', `url(#${t1.id})`); } if (h1) { defs.append(h1); pth.setAttribute('marker-start', `url(#${h1.id})`); }
        svg.append(defs); }
    }
    svg.append(pth); div.append(svg);
  }
  const tb = kid(n, 'txBody');
  if (tb) {
    const hasText = !!tb.getElementsByTagNameNS(NS.a, 't').length && [...tb.getElementsByTagNameNS(NS.a, 't')].some(t => t.textContent);
    if (hasText || (o.live && o.edit && phOf(n))) {
      const t = renderText(X, ctx, n, tb, div, { rmap: o.live ? X.rmap : null, index: o.index });
      if (o.live) X.dom.set(t, { node: n, top: o.top, tx: tb });
      if (b.rot && t) { /* text turns with the shape: nothing to do */ }
    }
  }
}
function drawTable(X, ctx, frame, tbl, div, b, o) {
  const t = document.createElement('table'); t.className = 'pe-tbl';
  const grid = kids(kid(tbl, 'tblGrid'), 'gridCol').map(g => num(at(g, 'w')) / PT);
  const cg = document.createElement('colgroup'); for (const w of grid) { const c = document.createElement('col'); c.style.width = w + 'pt'; cg.append(c); } t.append(cg);
  const tp = kid(tbl, 'tblPr'); const firstRow = at(tp, 'firstRow') === '1', band = at(tp, 'bandRow') === '1';
  const acc = colorOf(aColor('accent1'), ctx);
  kids(tbl, 'tr').forEach((tr, ri) => {
    const row = document.createElement('tr'); row.style.height = num(at(tr, 'h')) / PT + 'pt';
    kids(tr, 'tc').forEach((tc, ci) => {
      if (at(tc, 'hMerge') === '1' || at(tc, 'vMerge') === '1') return;
      const td = document.createElement('td');
      if (at(tc, 'gridSpan')) td.colSpan = +at(tc, 'gridSpan'); if (at(tc, 'rowSpan')) td.rowSpan = +at(tc, 'rowSpan');
      const pr = kid(tc, 'tcPr');
      let f = fillOf(pr, ctx, X, o.R);
      if (!f && at(tp, 'tableStyleId') !== undefined) { if (firstRow && ri === 0) f = { color: acc }; else if (band && ri % 2 === 1) f = { color: 'rgba(0,0,0,.05)' }; }
      if (f && f.color && !f.none) td.style.background = f.color;
      for (const [k, css] of [['lnL', 'borderLeft'], ['lnR', 'borderRight'], ['lnT', 'borderTop'], ['lnB', 'borderBottom']]) { const l = pr && kid(pr, k); if (l && !kid(l, 'noFill')) { const c = colorOf(kid(l, 'solidFill'), ctx) || '#888'; td.style[css] = `${Math.max(.5, num(at(l, 'w'), 12700) / PT)}pt solid ${c}`; } else if (!l) td.style[css] = '.75pt solid rgba(0,0,0,.25)'; }
      const tb = kid(tc, 'txBody');
      const tx = renderText(X, ctx, frame, tb || aNew(tbl.ownerDocument, 'txBody'), td, { sources: ctx.pres ? [ctx.pres.getElementsByTagNameNS(NS.p, 'defaultTextStyle')[0]].filter(Boolean) : [], bodyPrs: [], rmap: o.live ? X.rmap : null });
      Object.assign(tx.style, { position: 'relative', padding: `${num(at(pr, 'marT'), 45720) / PT}pt ${num(at(pr, 'marR'), 91440) / PT}pt ${num(at(pr, 'marB'), 45720) / PT}pt ${num(at(pr, 'marL'), 91440) / PT}pt`,
        justifyContent: at(pr, 'anchor') === 'ctr' ? 'center' : at(pr, 'anchor') === 'b' ? 'flex-end' : 'flex-start' });
      if (firstRow && ri === 0) for (const pp of tx.querySelectorAll('.pp')) { pp.style.fontWeight = pp.style.fontWeight || '700'; if (!f || f.color === acc) pp.style.color = '#fff'; }
      tx.classList.remove('blank');
      if (o.live) X.dom.set(tx, { node: frame, top: o.top, tx: tb, tc });
      row.append(td);
    });
    t.append(row);
  });
  div.append(t);
}

/* ------------------------------------------------------------ text → XML */
const RPR_ORDER = ['ln', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill', 'effectLst', 'effectDag', 'highlight', 'uLnTx', 'uLn', 'uFillTx', 'uFill', 'latin', 'ea', 'cs', 'sym', 'hlinkClick', 'hlinkMouseOver', 'rtl', 'extLst'];
const PPR_ORDER = ['lnSpc', 'spcBef', 'spcAft', 'buClrTx', 'buClr', 'buSzTx', 'buSzPct', 'buSzPts', 'buFontTx', 'buFont', 'buNone', 'buAutoNum', 'buChar', 'buBlip', 'tabLst', 'defRPr', 'extLst'];
const SPPR_ORDER = ['xfrm', 'custGeom', 'prstGeom', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill', 'ln', 'effectLst', 'effectDag', 'scene3d', 'sp3d', 'extLst'];
const FILLS = /^(noFill|solidFill|gradFill|blipFill|pattFill|grpFill)$/;
function putOrdered(parent, el, order) {
  const i = order.indexOf(el.localName);
  const after = kids(parent).find(k => { const j = order.indexOf(k.localName); return j > i || (j < 0 && i >= 0 && k.localName === 'extLst'); });
  parent.insertBefore(el, after || null);
}
const pxPt = v => { const m = /^([\d.]+)(px|pt)$/.exec(v || ''); return m ? (m[2] === 'px' ? +m[1] * .75 : +m[1]) : null; };
// the formatting of a piece of text: the nearest setting on the way up to its paragraph wins
function fmtOf(X, node, stop) {
  const f = {};
  for (let el = node.nodeType === 1 ? node : node.parentNode; el && el !== stop; el = el.parentNode) {
    if (el.nodeType !== 1) continue;
    const t = el.tagName, st = el.style;
    if (f.b === undefined) { if (t === 'B' || t === 'STRONG') f.b = true; else if (st.fontWeight) f.b = st.fontWeight === 'bold' || +st.fontWeight >= 600; }
    if (f.i === undefined) { if (t === 'I' || t === 'EM') f.i = true; else if (st.fontStyle) f.i = st.fontStyle === 'italic'; }
    const td = st.textDecorationLine || st.textDecoration;
    if (f.u === undefined) { if (t === 'U') f.u = true; else if (td) f.u = /underline/.test(td); }
    if (f.s === undefined) { if (t === 'S' || t === 'STRIKE' || t === 'DEL') f.s = true; else if (td) f.s = /line-through/.test(td); }
    if (f.sz === undefined && el.dataset && el.dataset.sz) f.sz = +el.dataset.sz;
    if (f.color === undefined && el.dataset && el.dataset.color) f.color = el.dataset.color;
    if (f.font === undefined && el.dataset && el.dataset.font) f.font = el.dataset.font;
    if (f.src === undefined && X.rmap.has(el)) f.src = X.rmap.get(el);
  }
  return f;
}
const sigOf = f => [f.b, f.i, f.u, f.s, f.sz, f.color, f.font].join('|');
function buildRPr(doc, f, base, text) {
  const r = base ? base.cloneNode(true) : aNew(doc, 'rPr', { lang: HEB.test(text || '') ? 'he-IL' : 'en-US' });
  if (r.localName !== 'rPr') { const n = aNew(doc, 'rPr'); for (const a of [...r.attributes]) n.setAttribute(a.name, a.value); for (const c of kids(r)) n.append(c.cloneNode(true)); return buildRPr(doc, f, n, text); }
  r.removeAttribute('dirty'); r.removeAttribute('err');
  if (f.b !== undefined) r.setAttribute('b', f.b ? '1' : '0');
  if (f.i !== undefined) r.setAttribute('i', f.i ? '1' : '0');
  if (f.u !== undefined) r.setAttribute('u', f.u ? 'sng' : 'none');
  if (f.s !== undefined) r.setAttribute('strike', f.s ? 'sngStrike' : 'noStrike');
  if (f.sz) r.setAttribute('sz', String(Math.round(f.sz)));
  if (f.color) { for (const k of kids(r)) if (FILLS.test(k.localName)) k.remove(); const s = aNew(doc, 'solidFill'); s.append(aNew(doc, 'srgbClr', { val: f.color })); putOrdered(r, s, RPR_ORDER); }
  if (f.font) { for (const k of kids(r)) if (k.localName === 'latin' || k.localName === 'cs') k.remove(); putOrdered(r, aNew(doc, 'latin', { typeface: f.font }), RPR_ORDER); putOrdered(r, aNew(doc, 'cs', { typeface: f.font }), RPR_ORDER); }
  if (HEB.test(text || '') && !r.hasAttribute('lang')) r.setAttribute('lang', 'he-IL');
  return r;
}
function writeTx(X, tx, txBody) {
  const doc = txBody.ownerDocument, inner = tx.querySelector('.pe-txi');
  // text typed outside any paragraph goes into one
  for (const n of [...inner.childNodes]) if (!(n.nodeType === 1 && n.classList.contains('pp'))) {
    if (n.nodeType === 3 && !n.textContent.trim()) { n.remove(); continue; }
    const prev = n.previousSibling && n.previousSibling.nodeType === 1 && n.previousSibling.classList.contains('pp') ? n.previousSibling : null;
    if (prev && n.nodeName !== 'DIV') prev.append(n); else { const d = document.createElement('div'); d.className = 'pp'; d.dataset.pi = prev ? prev.dataset.pi : 0; n.replaceWith(d); d.append(n); }
  }
  const old = kids(txBody, 'p');
  if (!kid(txBody, 'bodyPr')) txBody.prepend(aNew(doc, 'bodyPr'));
  if (!kid(txBody, 'lstStyle')) kid(txBody, 'bodyPr').after(aNew(doc, 'lstStyle'));
  for (const p of old) p.remove();
  const ALG = { left: 'l', right: 'r', center: 'ctr', justify: 'just', start: null, end: null };
  [...inner.children].forEach((div, i) => {
    const src = old[+div.dataset.pi] || old[old.length - 1] || null;
    const p = aNew(doc, 'p');
    let pPr = src && kid(src, 'pPr') ? kid(src, 'pPr').cloneNode(true) : null;
    const need = () => pPr || (pPr = aNew(doc, 'pPr'));
    if (div.dataset.algn && ALG[div.style.textAlign]) need().setAttribute('algn', ALG[div.style.textAlign]);
    if (div.dataset.rtl) need().setAttribute('rtl', div.dir === 'rtl' ? '1' : '0');
    if (div.dataset.buset) {
      const P = need(); for (const k of kids(P)) if (/^bu/.test(k.localName)) k.remove();
      if (div.dataset.buset === 'none') { putOrdered(P, aNew(doc, 'buNone'), PPR_ORDER); }
      else if (div.dataset.buset === 'num') { putOrdered(P, aNew(doc, 'buFont', { typeface: '+mj-lt' }), PPR_ORDER); putOrdered(P, aNew(doc, 'buAutoNum', { type: 'arabicPeriod' }), PPR_ORDER); if (!P.hasAttribute('marL')) { P.setAttribute('marL', '457200'); P.setAttribute('indent', '-457200'); } }
      else { putOrdered(P, aNew(doc, 'buFont', { typeface: 'Arial' }), PPR_ORDER); putOrdered(P, aNew(doc, 'buChar', { char: div.dataset.buset }), PPR_ORDER); if (!P.hasAttribute('marL')) { P.setAttribute('marL', '285750'); P.setAttribute('indent', '-285750'); } }
    }
    if (pPr) p.append(pPr);
    const firstR = src && kid(kid(src, 'r'), 'rPr');
    // the pieces of text with their formatting
    const segs = [];
    const walk = n => {
      for (const c of n.childNodes) {
        if (c.nodeType === 3) { if (c.textContent) segs.push({ t: c.textContent.replace(/\u200B/g, ''), f: fmtOf(X, c, div) }); }
        else if (c.nodeName === 'BR') { if (c !== div.lastChild) segs.push({ br: true, f: fmtOf(X, c, div) }); }   // a last <br> only keeps an empty line open
        else if (c.nodeType === 1) walk(c);
      }
    };
    walk(div);
    while (segs.length && segs[segs.length - 1].br) segs.pop();
    const merged = [];
    for (const s of segs) { const l = merged[merged.length - 1]; if (l && !l.br && !s.br && sigOf(l.f) === sigOf(s.f) && l.f.src === s.f.src) l.t += s.t; else if (s.br || s.t) merged.push({ ...s }); }
    let lastRPr = null;
    for (const s of merged) {
      const rPr = buildRPr(doc, s.f, s.f.src || firstR, s.t); lastRPr = rPr;
      if (s.br) { const b = aNew(doc, 'br'); b.append(rPr); p.append(b); continue; }
      const r = aNew(doc, 'r'); r.append(rPr); const t = aNew(doc, 't'); t.textContent = s.t; r.append(t); p.append(r);
    }
    const end = src && kid(src, 'endParaRPr') ? kid(src, 'endParaRPr').cloneNode(true) : null;
    if (end) p.append(end);
    else { const e = aNew(doc, 'endParaRPr', { lang: 'he-IL' }); const base = lastRPr || firstR; if (base) for (const a of ['sz', 'b', 'i']) if (base.hasAttribute(a)) e.setAttribute(a, at(base, a)); p.append(e); }
    txBody.append(p);
    div.dataset.pi = i;
  });
}

/* ------------------------------------------------------------ shapes in the XML */
const topBox = (X, ctx, n) => { const x = n.localName === 'graphicFrame' ? kid(n, 'xfrm') : xfrmOf(n, ctx); return x ? box(x) : { x: 0, y: 0, w: 0, h: 0, rot: 0 }; };
function ensureXfrm(n, ctx) {
  if (n.localName === 'graphicFrame') return kid(n, 'xfrm');
  const pr = kid(n, n.localName === 'grpSp' ? 'grpSpPr' : 'spPr') || (() => { const e = pNew(n.ownerDocument, n.localName === 'grpSp' ? 'grpSpPr' : 'spPr'); kids(n)[0].after(e); return e; })();
  let x = kid(pr, 'xfrm');
  if (!x) {   // a placeholder that took its place from the layout: the place becomes its own
    const b = topBox(null, ctx, n); x = aNew(n.ownerDocument, 'xfrm');
    x.append(aNew(n.ownerDocument, 'off', { x: Math.round(b.x * PT), y: Math.round(b.y * PT) }), aNew(n.ownerDocument, 'ext', { cx: Math.round(b.w * PT), cy: Math.round(b.h * PT) }));
    pr.prepend(x);
  }
  return x;
}
function setBox(n, ctx, b) {
  const x = ensureXfrm(n, ctx); if (!x) return;
  const off = kid(x, 'off'), ext = kid(x, 'ext');
  off.setAttribute('x', Math.round(b.x * PT)); off.setAttribute('y', Math.round(b.y * PT));
  ext.setAttribute('cx', Math.max(1, Math.round(b.w * PT))); ext.setAttribute('cy', Math.max(1, Math.round(b.h * PT)));
}
function nextId(doc) { let m = 1; for (const c of doc.getElementsByTagNameNS(NS.p, 'cNvPr')) m = Math.max(m, num(at(c, 'id'))); return m + 1; }
const SHAPE_NS = `xmlns:p="${NS.p}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"`;
function insertXml(doc, xml) {
  const el = doc.importNode(parseXml(xml).documentElement, true);
  const tree = doc.getElementsByTagNameNS(NS.p, 'spTree')[0];
  const ext = kids(tree, 'extLst')[0]; tree.insertBefore(el, ext || null);
  return el;
}
const xe = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
const emu = v => Math.round(v * PT);
function addTextBox(doc, b, text) {
  const id = nextId(doc), rtl = HEB.test(text) || !text;
  return insertXml(doc, `<p:sp ${SHAPE_NS}><p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${emu(b.x)}" y="${emu(b.y)}"/><a:ext cx="${emu(b.w)}" cy="${emu(b.h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" rtlCol="${rtl ? 1 : 0}"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="${rtl ? 'r' : 'l'}" rtl="${rtl ? 1 : 0}"/><a:r><a:rPr lang="${rtl ? 'he-IL' : 'en-US'}" sz="2400" dirty="0"/><a:t>${xe(text || 'טקסט')}</a:t></a:r></a:p></p:txBody></p:sp>`);
}
function addShapeXml(doc, prst, b) {
  const id = nextId(doc), line = /line/i.test(prst);
  if (line) return insertXml(doc, `<p:cxnSp ${SHAPE_NS}><p:nvCxnSpPr><p:cNvPr id="${id}" name="Connector ${id}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="${emu(b.x)}" y="${emu(b.y + b.h / 2)}"/><a:ext cx="${emu(b.w)}" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="28575"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln></p:spPr></p:cxnSp>`);
  return insertXml(doc, `<p:sp ${SHAPE_NS}><p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${emu(b.x)}" y="${emu(b.y)}"/><a:ext cx="${emu(b.w)}" cy="${emu(b.h)}"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln w="12700"><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr rtlCol="1" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr" rtl="1"/><a:endParaRPr lang="he-IL" sz="1800"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:endParaRPr></a:p></p:txBody></p:sp>`);
}
function addPictureXml(doc, rid, b, name) {
  const id = nextId(doc);
  return insertXml(doc, `<p:pic ${SHAPE_NS}><p:nvPicPr><p:cNvPr id="${id}" name="${xe(name || 'Picture ' + id)}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${emu(b.x)}" y="${emu(b.y)}"/><a:ext cx="${emu(b.w)}" cy="${emu(b.h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`);
}
function addTableXml(doc, rows, cols, b) {
  const id = nextId(doc), cw = Math.round(emu(b.w) / cols), rh = 370840;
  const cell = `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="r" rtl="1"/><a:endParaRPr lang="he-IL" sz="1800"/></a:p></a:txBody><a:tcPr/></a:tc>`;
  return insertXml(doc, `<p:graphicFrame ${SHAPE_NS}><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${emu(b.x)}" y="${emu(b.y)}"/><a:ext cx="${cw * cols}" cy="${rh * rows}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1" rtl="1"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr><a:tblGrid>${`<a:gridCol w="${cw}"/>`.repeat(cols)}</a:tblGrid>${`<a:tr h="${rh}">${cell.repeat(cols)}</a:tr>`.repeat(rows)}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`);
}
function setFill(n, color) {   // color: 'RRGGBB' or null (no fill)
  const pr = kid(n, 'spPr'); if (!pr) return;
  for (const k of kids(pr)) if (FILLS.test(k.localName)) k.remove();
  if (color === undefined) return;
  const f = color ? aNew(n.ownerDocument, 'solidFill') : aNew(n.ownerDocument, 'noFill'); if (color) f.append(aNew(n.ownerDocument, 'srgbClr', { val: color }));
  putOrdered(pr, f, SPPR_ORDER);
}
function setLine(n, color) {
  const pr = kid(n, 'spPr'); if (!pr) return;
  let ln = kid(pr, 'ln'); if (!ln) { ln = aNew(n.ownerDocument, 'ln', { w: 12700 }); putOrdered(pr, ln, SPPR_ORDER); }
  for (const k of kids(ln)) if (FILLS.test(k.localName)) k.remove();
  const f = color ? aNew(n.ownerDocument, 'solidFill') : aNew(n.ownerDocument, 'noFill'); if (color) f.append(aNew(n.ownerDocument, 'srgbClr', { val: color }));
  ln.prepend(f);
}
function setBackground(doc, color) {
  const cSld = doc.getElementsByTagNameNS(NS.p, 'cSld')[0];
  for (const b of kids(cSld, 'bg')) b.remove();
  if (!color) return;
  const bg = pNew(doc, 'bg'), pr = pNew(doc, 'bgPr'), f = aNew(doc, 'solidFill'); f.append(aNew(doc, 'srgbClr', { val: color }));
  pr.append(f, aNew(doc, 'effectLst')); bg.append(pr); cSld.prepend(bg);
}
function ensureTxBody(n) {
  let tb = kid(n, 'txBody'); if (tb) return tb;
  const d = n.ownerDocument;
  tb = parseXml(`<p:txBody ${SHAPE_NS}><a:bodyPr rtlCol="1" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr" rtl="1"/><a:endParaRPr lang="he-IL"/></a:p></p:txBody>`).documentElement;
  tb = d.importNode(tb, true); n.append(tb); return tb;
}

/* ------------------------------------------------------------ slides in the package */
const LAYOUT_HE = { title: 'שקף כותרת', obj: 'כותרת ותוכן', secHead: 'כותרת מקטע', twoObj: 'שני תכנים', twoTxTwoObj: 'השוואה', titleOnly: 'כותרת בלבד', blank: 'ריק', objTx: 'תוכן עם כיתוב', picTx: 'תמונה עם כיתוב', vertTx: 'טקסט אנכי', tx: 'כותרת וטקסט', cust: 'מותאם אישית' };
function layoutsOf(X, slidePath) {
  const ctx = ctxOf(X, slidePath); if (!ctx.masterPath) return [];
  return [...ctx.mR.values()].filter(r => r.type === RT.layout && X.files.has(r.target)).map(r => {
    const d = parseXml(X.files.get(r.target)); const t = at(d.documentElement, 'type') || 'cust';
    const nm = at(d.getElementsByTagNameNS(NS.p, 'cSld')[0], 'name') || '';
    return { path: r.target, type: t, name: LAYOUT_HE[t] && t !== 'cust' ? LAYOUT_HE[t] : nm === 'DEFAULT' ? 'שקף רגיל' : nm || LAYOUT_HE.cust };
  });
}
function newSlideXml(X, layoutPath) {
  const L = parseXml(X.files.get(layoutPath));
  const tree = L.getElementsByTagNameNS(NS.p, 'spTree')[0];
  let id = 2, sps = '';
  for (const sp of kids(tree, 'sp')) {
    const ph = phOf(sp); if (!ph) continue; const t = at(ph, 'type') || 'body';
    if (/^(dt|ftr|sldNum|hdr)$/.test(t)) continue;
    const idx = at(ph, 'idx');
    sps += `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Placeholder ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${t !== 'body' ? ` type="${t}"` : ''}${idx != null ? ` idx="${idx}"` : ''}/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="he-IL" dirty="0"/></a:p></p:txBody></p:sp>`;
    id++;
  }
  return `${XML_HEAD}<p:sld ${SHAPE_NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${sps}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
}
function insertSlide(X, at_, xml, relsXml) {
  const p = newSlidePath(X.files);
  X.files.set(p, xml); if (relsXml) X.files.set(relsPathOf(p), relsXml);
  ctEnsure(X.files, 'Override', '/' + p, CT_SLIDE);
  X.order.splice(at_, 0, p); writeOrder(X.files, X.order);
  return p;
}
function addSlide(X, after, layoutPath) {
  const p0 = newSlidePath(X.files);
  const rels = `${XML_HEAD}<Relationships xmlns="${NS.pr}"><Relationship Id="rId1" Type="${RT.layout}" Target="${relTarget(p0, layoutPath)}"/></Relationships>`;
  return insertSlide(X, after + 1, newSlideXml(X, layoutPath), rels);
}
function duplicateSlide(X, i) {
  const src = X.order[i], p0 = newSlidePath(X.files);
  const R = rels(X.files, src);
  let r = `${XML_HEAD}<Relationships xmlns="${NS.pr}">`;
  for (const [id, x] of R) { if (x.type === RT.notes) continue; r += `<Relationship Id="${id}" Type="${x.type}" Target="${x.external ? xe(x.raw) : relTarget(p0, x.target)}"${x.external ? ' TargetMode="External"' : ''}/>`; }
  r += '</Relationships>';
  return insertSlide(X, i + 1, X.files.get(src), r);
}
function deleteSlide(X, i) {
  const p = X.order[i]; X.order.splice(i, 1); writeOrder(X.files, X.order);
  X.files.delete(p); X.files.delete(relsPathOf(p)); ctDrop(X.files, p);
}
function addMedia(X, slidePath, bytes, ext, mime) {
  let n = 1; while (X.files.has(`ppt/media/image${n}.${ext}`)) n++;
  const mp = `ppt/media/image${n}.${ext}`; X.files.set(mp, bytes);
  ctEnsure(X.files, 'Default', ext, mime);
  return addRel(X.files, slidePath, RT.image, mp);
}

/* icons: the same set as the Word editor */
const ICO={
  undo:'<path d="M7 5L3 9l4 4"/><path d="M3 9h9a5 5 0 0 1 0 10H8"/>',
  redo:'<path d="M13 5l4 4-4 4"/><path d="M17 9H8a5 5 0 0 0 0 10h4"/>',
  cut:'<circle cx="6" cy="15" r="2.5"/><circle cx="14" cy="15" r="2.5"/><path d="M7.5 13L14 3M12.5 13L6 3"/>',
  copy:'<rect x="7" y="7" width="10" height="11" rx="1.5"/><path d="M4 13V4a1 1 0 0 1 1-1h8"/>',
  paste:'<rect x="4" y="4" width="12" height="14" rx="1.5"/><path d="M7.5 4V2.5h5V4"/><path d="M7 9h6M7 12h6M7 15h4"/>',
  brush:'<path d="M4 3h10v5H4z"/><path d="M14 5.5h2v4H9.5v3"/><rect x="8" y="12.5" width="3" height="5.5" rx="1"/>',
  bullets:'<circle cx="16" cy="5" r="1.2" fill="currentColor"/><circle cx="16" cy="10" r="1.2" fill="currentColor"/><circle cx="16" cy="15" r="1.2" fill="currentColor"/><path d="M3 5h10M3 10h10M3 15h10"/>',
  numbers:'<path d="M15 3.5h1.2v4M14.6 7.5h2.4M14.5 12.2c.4-.7 2.6-.8 2.4.7-.2 1-2.4 1.6-2.4 3h2.6"/><path d="M3 5h9M3 10h9M3 15h9"/>',
  outdent:'<path d="M3 4h14M9 8h8M9 12h8M3 16h14"/><path d="M3 10l3-2v4z" fill="currentColor"/>',
  indent:'<path d="M3 4h14M3 8h8M3 12h8M3 16h14"/><path d="M17 10l-3-2v4z" fill="currentColor"/>',
  aright:'<path d="M3 4h14M7 8h10M3 12h14M7 16h10"/>',
  acenter:'<path d="M3 4h14M5.5 8h9M3 12h14M5.5 16h9"/>',
  aleft:'<path d="M3 4h14M3 8h10M3 12h14M3 16h10"/>',
  ajust:'<path d="M3 4h14M3 8h14M3 12h14M3 16h14"/>',
  rtl:'<path d="M9 3h7M13 3v13M10 3v13M10 3a3 3 0 0 0 0 6"/><path d="M8 15H3m0 0l2-2m-2 2l2 2"/>',
  ltr:'<path d="M4 3h7M8 3v13M5 3v13M5 3a3 3 0 0 0 0 6"/><path d="M12 15h5m0 0l-2-2m2 2l-2 2"/>',
  spacing:'<path d="M9 4h8M9 8h8M9 12h8M9 16h8"/><path d="M4 3v14M2 5l2-2 2 2M2 15l2 2 2-2"/>',
  pilcrow:'<path d="M14 3v14M11 3v14M11 3H8.5a3.5 3.5 0 0 0 0 7H11"/>',
  find:'<circle cx="8.5" cy="8.5" r="5"/><path d="M12.5 12.5L17 17"/>',
  replace:'<path d="M3 6h9l-3-3M17 14H8l3 3"/>',
  table:'<rect x="3" y="4" width="14" height="12" rx="1"/><path d="M3 8h14M3 12h14M8 4v12M12.5 4v12"/>',
  image:'<rect x="3" y="4" width="14" height="12" rx="1.5"/><circle cx="7.5" cy="8.5" r="1.4"/><path d="M17 13l-4-4-7 7"/>',
  link:'<path d="M8.5 11.5a3 3 0 0 0 4.2 0l3-3a3 3 0 0 0-4.2-4.2l-1 1"/><path d="M11.5 8.5a3 3 0 0 0-4.2 0l-3 3a3 3 0 0 0 4.2 4.2l1-1"/>',
  pagebreak:'<path d="M5 2v5h10V2M5 18v-5h10v5"/><path d="M2 10h3M7.5 10h2M11 10h2M15 10h3"/>',
  hr:'<path d="M3 10h14"/><path d="M5 5h10M5 15h10" opacity=".4"/>',
  symbol:'<path d="M5 16h3v-2.2A6 6 0 1 1 12 13.8V16h3"/>',
  date:'<rect x="3" y="4" width="14" height="13" rx="1.5"/><path d="M3 8h14M7 2.5v3M13 2.5v3M6.5 11.5h2M11.5 11.5h2M6.5 14.5h2"/>',
  margins:'<rect x="4" y="2.5" width="12" height="15" rx="1"/><path d="M4 6h12M4 14h12M7 2.5v15M13 2.5v15" stroke-dasharray="1.5 1.5"/>',
  orient:'<rect x="2.5" y="6" width="15" height="10" rx="1"/><path d="M6 2.5h5" /><path d="M11 2.5l-1.5-1M11 2.5l-1.5 1"/>',
  size:'<path d="M5 2.5h7l3.5 3.5v11.5H5z"/><path d="M12 2.5V6h3.5"/>',
  print:'<path d="M5.5 7V3h9v4"/><rect x="3" y="7" width="14" height="7" rx="1.5"/><path d="M6 12h8v5H6z"/>',
  zin:'<circle cx="8.5" cy="8.5" r="5"/><path d="M12.5 12.5L17 17M6.5 8.5h4M8.5 6.5v4"/>',
  zout:'<circle cx="8.5" cy="8.5" r="5"/><path d="M12.5 12.5L17 17M6.5 8.5h4"/>',
  pages:'<rect x="5" y="2.5" width="10" height="6.5" rx="1"/><rect x="5" y="11" width="10" height="6.5" rx="1"/>',
  spell:'<path d="M3 12l3-8 3 8M4.2 9.5h3.6"/><path d="M11 14l2.5 2.5L18 11"/>',
  rowAbove:'<rect x="3" y="9" width="14" height="8" rx="1"/><path d="M3 13h14M10 9v8"/><path d="M10 2v5M7.5 4.5h5"/>',
  rowBelow:'<rect x="3" y="3" width="14" height="8" rx="1"/><path d="M3 7h14M10 3v8"/><path d="M10 13v5M7.5 15.5h5"/>',
  colRight:'<rect x="3" y="3" width="8" height="14" rx="1"/><path d="M7 3v14M3 10h8"/><path d="M13 10h5M15.5 7.5v5"/>',
  colLeft:'<rect x="9" y="3" width="8" height="14" rx="1"/><path d="M13 3v14M9 10h8"/><path d="M2 10h5M4.5 7.5v5"/>',
  delRow:'<rect x="3" y="6" width="14" height="8" rx="1"/><path d="M8 8.5l4 3M12 8.5l-4 3"/>',
  delCol:'<rect x="6" y="3" width="8" height="14" rx="1"/><path d="M8 8l4 4M12 8l-4 4"/>',
  delTable:'<rect x="3" y="4" width="14" height="12" rx="1"/><path d="M7 7l6 6M13 7l-6 6"/>',
  shade:'<path d="M5 11l5-7 5 7a5 5 0 0 1-10 0z"/>',
  borders:'<rect x="3" y="3" width="14" height="14"/><path d="M3 10h14M10 3v14"/>',
  noborders:'<rect x="3" y="3" width="14" height="14" stroke-dasharray="1.5 2"/><path d="M3 10h14M10 3v14" stroke-dasharray="1.5 2"/>',
  grow:'<path d="M3 16l4.5-12L12 16M4.6 12h5.8"/><path d="M14 7l2-2.5L18 7"/>',
  shrink:'<path d="M3 16l4.5-12L12 16M4.6 12h5.8"/><path d="M14 4.5l2 2.5 2-2.5"/>',
  sub:'<path d="M3 4l7 9M10 4l-7 9"/><path d="M13 12.5c.3-.9 3.1-1 2.9.8-.2 1.1-2.9 1.9-2.9 3.7h3.2"/>',
  sup:'<path d="M3 7l7 9M10 7l-7 9"/><path d="M13 3.5c.3-.9 3.1-1 2.9.8-.2 1.1-2.9 1.9-2.9 3.7h3.2"/>',
  clear:'<path d="M5 15l3.5-11h3L15 15M6.3 11h7.4"/><path d="M13 13l4 4M17 13l-4 4"/>',
  caret:'<path d="M6 8l4 4 4-4"/>',
  header:'<rect x="3.5" y="2.5" width="13" height="15" rx="1"/><path d="M6 5.5h8" stroke-width="2.4"/><path d="M6 10h8M6 13h6" opacity=".45"/>',
  footer:'<rect x="3.5" y="2.5" width="13" height="15" rx="1"/><path d="M6 14.5h8" stroke-width="2.4"/><path d="M6 6h8M6 9h6" opacity=".45"/>',
  pagenum:'<rect x="3.5" y="2.5" width="13" height="15" rx="1"/><path d="M8.6 11.2c.2-1 2.7-1.1 2.6.4-.1 1-2.6 1.6-2.6 2.9h2.8"/>',
  toc:'<path d="M3 4.5h9M5 8.5h7M5 12.5h7M3 16.5h9"/><path d="M14.5 4.5h2.5M14.5 8.5h2.5M14.5 12.5h2.5M14.5 16.5h2.5" stroke-dasharray="1 1.4"/>',
  comment:'<path d="M3.5 4.5h13v9h-7l-4 3v-3h-2z"/><path d="M10 7v4M8 9h4"/>',
  cmDel:'<path d="M3.5 4.5h13v9h-7l-4 3v-3h-2z"/><path d="M8 7.2l4 3.6M12 7.2l-4 3.6"/>',
  prev:'<path d="M3.5 4.5h13v9h-7l-4 3v-3h-2z"/><path d="M8.5 7l3 2-3 2"/>',
  next:'<path d="M3.5 4.5h13v9h-7l-4 3v-3h-2z"/><path d="M11.5 7l-3 2 3 2"/>',
  cmShow:'<path d="M3.5 4.5h13v9h-7l-4 3v-3h-2z"/><path d="M7 9h6"/>',
  count:'<path d="M4 5h4M4 9h4M4 13h4"/><path d="M11 4.5c.3-.8 2.9-.9 2.7.7-.2 1.1-2.7 1.8-2.7 3.4h3M11.2 12.2c.5-.9 3-.7 2.7.8.3.9-.3 2.2-1.7 2.1-.6 0-1.1-.2-1.3-.7"/>',
  ac:'<path d="M3 15l3.5-10h1L11 15M4.3 11.5h5.4"/><path d="M12 12.5l2.2 2.3L18 9.5"/>',
  closeX:'<rect x="3" y="3" width="14" height="14" rx="2"/><path d="M7 7l6 6M13 7l-6 6"/>',
  swap:'<path d="M5 3v14M5 3L2.5 5.5M5 3l2.5 2.5M15 17V3M15 17l-2.5-2.5M15 17l2.5-2.5"/>',
  clock:'<circle cx="10" cy="10" r="7"/><path d="M10 6v4l2.5 2"/>'
};

const BASE_CSS = `
@font-face{font-family:"David";font-weight:400;src:local("David"),url(font-david-400.ttf) format("truetype")}
@font-face{font-family:"David";font-weight:700;src:local("David Bold"),local("David-Bold"),url(font-david-700.ttf) format("truetype")}
@font-face{font-family:"FrankRuehl";font-weight:400;src:local("FrankRuehl"),url(font-frank-400.ttf) format("truetype")}
@font-face{font-family:"FrankRuehl";font-weight:700;src:local("FrankRuehl Bold"),url(font-frank-700.ttf) format("truetype")}
@font-face{font-family:"Frank Ruehl";font-weight:400;src:local("FrankRuehl"),url(font-frank-400.ttf) format("truetype")}
@font-face{font-family:"Frank Ruehl";font-weight:700;src:local("FrankRuehl Bold"),url(font-frank-700.ttf) format("truetype")}
@font-face{font-family:"Heebo";font-weight:400;src:local("Heebo"),url(font-heebo-400.ttf) format("truetype")}
@font-face{font-family:"Heebo";font-weight:700;src:local("Heebo Bold"),url(font-heebo-700.ttf) format("truetype")}
@font-face{font-family:"Rubik";font-weight:400;src:local("Rubik"),url(font-rubik-400.ttf) format("truetype")}
@font-face{font-family:"Rubik";font-weight:700;src:local("Rubik Bold"),url(font-rubik-700.ttf) format("truetype")}
@font-face{font-family:"Noto Sans Hebrew";font-weight:400;src:local("Noto Sans Hebrew"),url(font-noto-400.ttf) format("truetype")}
@font-face{font-family:"Noto Sans Hebrew";font-weight:700;src:local("Noto Sans Hebrew Bold"),url(font-noto-700.ttf) format("truetype")}
#ped{--wn:var(--navy,#16213A);--wn2:var(--navy-2,#202D4A);--wm:var(--marker,#FFD43B);--wb:var(--blue,#2E58CC);--wl:var(--line,#D4DBE5);--wt:var(--ink,#16213A);--wmu:var(--muted,#5D6982);
  position:fixed;inset:0;z-index:900;display:flex;flex-direction:column;background:#E4E8EF;color:var(--wt);direction:rtl;font:14px/1.3 "SF-heebo","Heebo",system-ui,sans-serif;-webkit-user-select:none;user-select:none}
#ped[hidden]{display:none}
#ped button{font:inherit;color:inherit}
#ped .grow{flex:1}
.pe-top{display:flex;align-items:center;gap:8px;padding:7px 12px;background:var(--wn);color:#E4EAF4;min-height:48px}
.pe-brand{display:flex;align-items:center;gap:7px;white-space:nowrap}
.pe-brand b{font-size:17px;color:#fff;background:linear-gradient(transparent 58%,rgba(255,212,59,.55) 58%);padding:0 2px}
.pe-mark{width:24px;height:24px;border-radius:5px;background:#2B579A;color:#fff;display:grid;place-items:center;font:700 13px Georgia,serif}
.pe-sub{color:#9AA7BE;font-size:13px}
.pe-namebox{direction:ltr;display:flex;align-items:center;max-width:340px;min-width:90px;flex:0 1 340px;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.14);border-radius:7px;padding:0 8px;margin-inline-start:10px}
.pe-name{flex:1;min-width:0;background:none;border:0;color:#fff;font:inherit;padding:6px 0;outline:none}
.pe-namebox:focus-within{border-color:var(--wm)}
.pe-ext{color:#9AA7BE;font-size:12px;direction:ltr}
.pe-state{font-size:12px;color:#9AA7BE;white-space:nowrap}
.pe-state.dirty{color:var(--wm)}
.pe-tb{border:1px solid rgba(255,255,255,.22);background:rgba(255,255,255,.06);color:#fff!important;border-radius:8px;padding:7px 13px;cursor:pointer;white-space:nowrap}
.pe-tb:hover{background:rgba(255,255,255,.14)}
.pe-tb.pri{background:var(--wm);border-color:var(--wm);color:var(--wn)!important;font-weight:700}
.pe-tb.pri:hover{filter:brightness(1.05)}
.pe-tb.sm{padding:4px 10px;color:var(--wt)!important;border-color:var(--wl);background:#fff}
.pe-x{width:34px;height:34px;border-radius:8px;border:0;background:none;color:#C9D2E1!important;font-size:22px;cursor:pointer}
.pe-x:hover{background:rgba(255,255,255,.12);color:#fff!important}
.pe-tabs{display:flex;gap:2px;padding:0 10px;background:var(--wn)}
#ped .pe-tabs button{border:0;background:none;color:#C9D2E1;padding:7px 14px 8px;border-radius:8px 8px 0 0;cursor:pointer;position:relative}
#ped .pe-tabs button[aria-selected=true]{background:#F7F8FB;color:var(--wn);font-weight:600}
.pe-tabs button[aria-selected=true]::after{content:"";position:absolute;inset-inline:12px;bottom:3px;height:3px;border-radius:2px;background:var(--wm)}
#ped .pe-tabs button.ctx{color:var(--wm)}
.pe-ribbon{background:#F7F8FB;border-bottom:1px solid var(--wl);overflow-x:auto;overflow-y:hidden;scrollbar-width:thin}
#ped .wpanel{display:flex;align-items:stretch;min-height:92px;padding:6px 6px 2px;width:max-content}
#ped .wpanel[hidden]{display:none}
#ped .wg{display:flex;flex-direction:column;padding:0 8px;border-inline-start:1px solid #E1E6EE}
#ped .wg:first-child{border-inline-start:0}
#ped .wg-b{flex:1;display:flex;align-items:center;gap:3px}
#ped .wg-l{font-size:11px;color:var(--wmu);text-align:center;padding:2px 0 3px;white-space:nowrap}
#ped .wg .row{display:flex;align-items:center;gap:2px}
#ped .wg-b:has(.row){flex-direction:column;align-items:stretch;justify-content:center;gap:4px}
#ped .wg .col{display:flex;flex-direction:column;gap:1px}
#ped .wb{min-width:30px;height:30px;border:1px solid transparent;border-radius:6px;background:none;cursor:pointer;display:inline-grid;place-items:center;padding:0 4px;position:relative}
#ped .wb svg{width:19px;height:19px}
#ped .wb:hover{background:#E8ECF3}
#ped .wb[aria-pressed=true],#ped .wb.on{background:#FFF3C4;border-color:#F2D46B}
#ped .wb:disabled{opacity:.35;cursor:default}
#ped .wb.big{flex-direction:column;height:auto;min-height:62px;min-width:54px;gap:3px;padding:4px 6px;font-size:12px;display:inline-flex;align-items:center;justify-content:center}
#ped .wb.big svg{width:24px;height:24px}
#ped .wb.txt{font-size:12px;padding:0 8px}
#ped .wb .gl{font:700 16px/1 Georgia,"Times New Roman",serif}
#ped .wb .gl.a{font:600 16px/1 Arial,sans-serif;margin-top:-4px}
#ped .wb .cbar{position:absolute;inset-inline:6px;bottom:4px;height:4px;border-radius:1px;background:#C00000}
#ped .wb .cbar[data-bar=hilite]{background:#FFFF00}
#ped .wb .cbar[data-bar=shade]{background:#D9E2F3}
#ped .wsplit{display:inline-flex}
#ped .wsplit .wb:first-child{border-start-end-radius:0;border-end-end-radius:0;min-width:28px}
#ped .wsplit .wb:last-child{min-width:16px;padding:0;border-start-start-radius:0;border-end-start-radius:0}
#ped .wsplit .wb:last-child svg{width:12px;height:12px}
#ped .wsel{height:28px;border:1px solid #CBD3DF;border-radius:6px;background:#fff;font:inherit;font-size:13px;padding:0 6px;color:var(--wt)}
#ped .wsel.font{width:150px}
#ped .wsel.size{width:54px;direction:ltr;text-align:center}
#ped .wsel:focus{outline:2px solid var(--wb);outline-offset:-1px}
#ped .wgal{display:flex;gap:4px;max-width:480px;overflow-x:auto;padding:2px;scrollbar-width:thin}
#ped .wgal button{flex:none;width:86px;height:58px;border:1px solid #DCE2EB;border-radius:7px;background:#fff;cursor:pointer;overflow:hidden;padding:4px 6px;text-align:start;display:flex;flex-direction:column;justify-content:space-between}
#ped .wgal button:hover{border-color:#9FB2D6}
#ped .wgal button.on{border-color:var(--wb);box-shadow:0 0 0 1px var(--wb) inset}
#ped .wgal .smp{height:30px;overflow:hidden;white-space:nowrap}
#ped .wgal .smp p{margin:0!important;padding:0!important;border:0!important;text-indent:0!important;white-space:nowrap;zoom:.62;line-height:1.2!important;min-height:0}
#ped .wgal .smp p::before{display:none!important}
#ped .wgal small{font-size:11px;color:var(--wmu);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
#ped .num{display:flex;align-items:center;gap:5px;font-size:12px;color:var(--wmu);margin:2px 0}
#ped .num input{width:58px;height:26px;border:1px solid #CBD3DF;border-radius:6px;padding:0 5px;font:inherit;color:var(--wt)}
#ped .wg-b:has(.num){flex-direction:column;align-items:stretch;justify-content:center}
#ped .zv{min-width:44px;text-align:center;font-size:13px}
.pe-scroll{flex:1;overflow:auto;position:relative;-webkit-user-select:text;user-select:text;overscroll-behavior:contain}
.pe-canvas{padding:26px 24px 60px;width:max-content;min-width:100%;box-sizing:border-box;display:flex;justify-content:center}
.pe-sheet{background:#fff;position:relative;box-sizing:border-box;box-shadow:0 1px 0 rgba(22,33,58,.05),0 12px 30px -14px rgba(22,33,58,.4);flex:none}
.pe-body{outline:none;min-height:100%;position:relative;z-index:1;color:#000;caret-color:#000;overflow-wrap:break-word;line-height:normal;text-align:start}
.pe-body:empty::before,.pe-body>p:only-child:empty::before{content:"מתחילים לכתוב כאן…";color:#A4AEC0}
.pe-body img{max-width:100%}
.pe-body a{cursor:text}
.pe-body a[href]{color:inherit}
.pe-body table{border-collapse:collapse}
.pe-body td,.pe-body th{min-width:18px}
.pe-body .pe-atom{cursor:default}
.pe-body .pe-atom img{cursor:pointer}
.pe-body .pe-pb{display:block;height:0;border-top:1px dashed #8E9AB3;margin:12px 0;position:relative;text-indent:0;clear:both}
.pe-body .pe-pb::after{content:"מעבר עמוד";position:absolute;left:50%;transform:translate(-50%,-55%);background:#fff;padding:0 8px;font:11px "SF-heebo",Heebo,sans-serif;color:#7A879E}
.pe-body .pe-tab{display:inline-block;width:36pt;white-space:pre}
.pe-body hr{border:0;border-top:1px solid #000;margin:6pt 0}
.pe-body ::selection{background:rgba(46,88,204,.24)}
.pe-marks .pe-body p::after,.pe-marks .pe-body li::after{content:"¶";color:#8E9AB3;font:400 11pt Arial,sans-serif}
.pe-marks .pe-body .pe-tab::before{content:"→";color:#8E9AB3;font:9pt Arial,sans-serif}
.pe-hdr,.pe-ftr{position:absolute;opacity:.5;overflow:hidden;min-height:22px;cursor:text;z-index:3;border-radius:3px}
.pe-hdr:hover,.pe-ftr:hover{outline:1px dashed rgba(46,88,204,.35);outline-offset:3px}
#ped .pe-hdr.ed,#ped .pe-ftr.ed{opacity:1;overflow:visible;max-height:none!important;outline:1px dashed #2E58CC;outline-offset:4px;z-index:4}
#ped .pe-hdr.ed>*,#ped .pe-ftr.ed>*{outline:none;caret-color:#000;-webkit-user-select:text;user-select:text}
#ped .pe-hdr.ed::before,#ped .pe-ftr.ed::before{position:absolute;inset-inline-start:-4px;font:600 11px "SF-heebo",Heebo,sans-serif;color:#fff;background:#2E58CC;padding:1px 7px;border-radius:3px;white-space:nowrap}
#ped .pe-hdr.ed::before{content:"כותרת עליונה";bottom:-24px}#ped .pe-ftr.ed::before{content:"כותרת תחתונה";top:-24px}
#ped.pe-hfmode .pe-body,#ped.pe-hfmode .pe-notes{opacity:.42}
#ped .pe-fld{background:#E3E7EE;border-radius:2px;padding:0 1px}
#ped .pe-cm{display:inline;width:0;font-size:0}
#ped .pe-cm[data-cs]::before{content:"";display:inline-block;width:0;height:1.05em;border-inline-start:2px solid #E3A008;vertical-align:-.15em}
::highlight(we-cmt){background-color:#FFF1C2}::highlight(we-cmt-on){background-color:#FFD76A}
#ped .pe-canvas{gap:18px}
#ped .pe-cpane{position:relative;width:270px;flex:none;display:none;align-self:flex-start;-webkit-user-select:none;user-select:none}
#ped.pe-hascm .pe-cpane{display:block}
#ped .cmc{position:absolute;inset-inline:0;background:#fff;border:1px solid var(--wl);border-inline-start:3px solid #E3A008;border-radius:8px;padding:8px 10px;box-shadow:0 4px 14px -8px rgba(22,33,58,.35);font-size:13px;cursor:pointer;transition:top .15s}
#ped .cmc.on{box-shadow:0 0 0 2px #E3A008,0 8px 22px -10px rgba(22,33,58,.45)}
#ped .cmc.done{opacity:.62;border-inline-start-color:#7A879E}
#ped .cmd{font-size:11px;color:#2E7D32;margin-bottom:3px}
#ped .cmh{display:flex;align-items:center;gap:6px;margin-bottom:4px}#ped .cmh b{font-size:12.5px}#ped .cmh small{color:var(--wmu);font-size:11px;margin-inline-start:auto}
#ped .av{display:inline-grid;place-items:center;width:22px;height:22px;border-radius:50%;background:#2E58CC;color:#fff;font:700 10px/1 sans-serif;font-style:normal;flex:none}
#ped .cmx{white-space:normal;overflow-wrap:anywhere;line-height:1.4;color:#16213A}
#ped .cmr{border-top:1px solid #EEF1F5;margin-top:7px;padding-top:7px}
#ped .cmt{width:100%;box-sizing:border-box;border:1px solid #CBD3DF;border-radius:6px;padding:5px 7px;font:inherit;resize:vertical;min-height:44px;-webkit-user-select:text;user-select:text}
#ped .cmt:focus{outline:2px solid var(--wb);outline-offset:-1px}
#ped .cma{display:flex;gap:4px;justify-content:flex-end;margin-top:6px;flex-wrap:wrap}
#ped .cma .wb.txt{height:26px;padding:0 8px;font-size:12px;width:auto;border-radius:6px}
#ped .cma .wb.pri{background:var(--wb);color:#fff}
#ped .pe-toc{display:flex;align-items:baseline}#ped .pe-toc>span:first-child{flex:none;max-width:85%}
#ped .pe-toc .pe-lead{flex:1;border-bottom:1.5px dotted #8E9AB3;margin:0 4px;width:auto;min-width:20px;color:transparent;height:.6em}
#ped .pe-toc:hover{background:#F1F5FC;cursor:pointer}
#ped .pe-wc{border:0;background:none;color:inherit;font:inherit;padding:0 4px;cursor:pointer;border-radius:4px}#ped .pe-wc:hover{background:#E4E8EF}
#ped table.wc{border-collapse:collapse;width:100%;margin:6px 0 4px;font-size:13.5px}#ped table.wc td,#ped table.wc th{padding:5px 8px;border-bottom:1px solid #EEF1F5;text-align:start}#ped table.wc th{color:var(--wmu);font-weight:600;font-size:12px}
@media (max-width:900px){#ped .pe-canvas{flex-direction:column;align-items:center}#ped .pe-cpane{width:min(100%,520px);align-self:center}#ped .cmc{position:relative;top:auto!important;margin-bottom:8px}#ped .pe-cpane{height:auto!important}}
.pe-hdr header,.pe-ftr footer{margin:0!important;min-height:0!important}
.pe-notes{border-top:1px solid #D6DCE6;margin-top:18px;padding-top:6px;font-size:10pt;opacity:.8}
.pe-notes:empty{display:none}
.pe-guides{position:absolute;inset:0;pointer-events:none;z-index:2}
.pe-guides i{position:absolute;inset-inline:0;border-top:1px dashed rgba(46,88,204,.35)}
.pe-guides i b{position:absolute;inset-inline-end:calc(100% + 8px);top:-9px;font:600 11px "SF-heebo",Heebo,sans-serif;color:#7A879E;white-space:nowrap}
.pe-noguides .pe-guides{display:none}
.pe-imgbox{position:absolute;z-index:3;outline:2px solid var(--wb);pointer-events:none}
.pe-imgbox i{position:absolute;width:10px;height:10px;background:#fff;border:2px solid var(--wb);border-radius:2px;pointer-events:auto}
.pe-imgbox i[data-h=nw]{top:-6px;left:-6px;cursor:nwse-resize}.pe-imgbox i[data-h=ne]{top:-6px;right:-6px;cursor:nesw-resize}
.pe-imgbox i[data-h=sw]{bottom:-6px;left:-6px;cursor:nesw-resize}.pe-imgbox i[data-h=se]{bottom:-6px;right:-6px;cursor:nwse-resize}
.pe-status{display:flex;align-items:center;gap:14px;padding:3px 12px;background:#F7F8FB;border-top:1px solid var(--wl);font-size:12px;color:var(--wmu);min-height:28px}
#ped .pe-status .wb{min-width:22px;height:22px}
.pe-status input[type=range]{width:110px;accent-color:var(--wb)}
.pe-hint{color:#2E58CC}
.pe-find{position:absolute;top:150px;inset-inline-start:24px;z-index:20;background:#fff;border:1px solid var(--wl);border-radius:10px;box-shadow:0 12px 30px -10px rgba(22,33,58,.35);padding:8px;display:flex;flex-direction:column;gap:6px;width:330px}
.pe-find[hidden]{display:none}
.pe-find .fr{display:flex;align-items:center;gap:4px}
.pe-find input{flex:1;min-width:0;height:30px;border:1px solid #CBD3DF;border-radius:6px;padding:0 8px;font:inherit}
.pe-find .fn{font-size:12px;color:var(--wmu);min-width:46px;text-align:center}
.pe-find:not(.rep-on) .rep{display:none}
.pe-pop{position:fixed;z-index:30;max-height:calc(100vh - 16px);overflow:auto;box-sizing:border-box;background:#fff;border:1px solid var(--wl);border-radius:10px;box-shadow:0 14px 34px -12px rgba(22,33,58,.4);padding:8px;max-width:340px}
.pe-pop[hidden]{display:none}
.pe-pop h4{margin:2px 2px 6px;font-size:12px;color:var(--wmu);font-weight:600}
.pe-pop .pal{display:grid;grid-template-columns:repeat(10,22px);gap:4px}
.pe-pop .pal button{width:22px;height:22px;border-radius:4px;border:1px solid rgba(0,0,0,.15);cursor:pointer;padding:0}
.pe-pop .pal button:hover{transform:scale(1.15)}
.pe-pop .mi{display:flex;align-items:center;gap:8px;width:100%;border:0;background:none;padding:7px 10px;border-radius:7px;cursor:pointer;text-align:start;white-space:nowrap}
.pe-pop .mi:hover,.pe-pop .mi.on{background:#EEF2F8}
.pe-pop .mi small{color:var(--wmu);margin-inline-start:auto;font-size:12px;padding-inline-start:18px}
.pe-pop .sep{height:1px;background:var(--wl);margin:5px 0}
.pe-pop .tgrid{display:grid;grid-template-columns:repeat(10,18px);gap:3px;direction:rtl}
.pe-pop .tgrid i{width:18px;height:18px;border:1px solid #C6CFDD;border-radius:2px;cursor:pointer;display:block}
.pe-pop .tgrid i.on{background:#FFF0B3;border-color:#E3B92B}
.pe-pop .tlab{text-align:center;font-size:12px;color:var(--wmu);margin-top:6px}
.pe-pop .syms{display:grid;grid-template-columns:repeat(8,32px);gap:3px}
.pe-pop .syms button{height:32px;border:1px solid #E1E6EE;border-radius:6px;background:#fff;cursor:pointer;font-size:16px}
.pe-pop .syms button:hover{background:#FFF3C4}
.pe-pop .lnk{border:0;background:none;color:var(--wb);cursor:pointer;padding:6px 2px;font-size:13px}
.pe-dlg{position:absolute;inset:0;z-index:40;background:rgba(22,33,58,.35);display:grid;place-items:center}
.pe-dlg[hidden]{display:none}
.pe-dlg-in{background:#fff;border-radius:14px;padding:18px 20px;width:min(420px,92vw);box-shadow:0 24px 60px -20px rgba(0,0,0,.45);-webkit-user-select:text;user-select:text}
.pe-dlg-in h3{margin:0 0 10px;font-size:17px}
.pe-dlg-in label{display:block;font-size:13px;color:var(--wmu);margin:8px 0 3px}
.pe-dlg-in input{width:100%;box-sizing:border-box;height:36px;border:1px solid #CBD3DF;border-radius:8px;padding:0 10px;font:inherit}
.pe-dlg-in p{color:var(--wmu);font-size:13px;line-height:1.5;margin:4px 0 0}
.pe-dlg-in .acts{display:flex;gap:8px;justify-content:flex-end;margin-top:16px}
.pe-dlg-in .acts button{border-radius:8px;padding:8px 16px;border:1px solid var(--wl);background:#fff;cursor:pointer}
.pe-dlg-in .acts .pri{background:var(--wn);color:#fff!important;border-color:var(--wn)}
.pe-dlg-in .acts .warn{background:#C0392B;color:#fff!important;border-color:#C0392B}
.pe-load{position:absolute;inset:0;display:grid;place-items:center;color:var(--wmu);font-size:15px;z-index:5;background:#E4E8EF}
@media (max-width:760px){.pe-sub,.pe-state,.wg-l{display:none}.pe-tb{padding:6px 9px;font-size:13px}#weDesk{display:none}.pe-namebox{margin-inline-start:4px}.pe-canvas{padding:12px 8px 40px}.wpanel{min-height:76px}.pe-find{top:auto;bottom:40px;inset-inline:8px;width:auto}}
@media print{
  html>body>*:not(#ped){display:none!important}
  html>body>#ped{position:static!important;display:block!important;background:#fff!important;inset:auto!important}
  #ped>*:not(.pe-scroll){display:none!important}
  .pe-scroll{overflow:visible}.pe-canvas{padding:0;display:block}
  .pe-sheet{box-shadow:none!important;padding:0!important;width:auto!important;min-height:0!important;zoom:1!important}
  .pe-hdr,.pe-ftr,.pe-guides,.pe-imgbox{display:none!important}
  .pe-body .pe-pb{break-after:page;border:0;margin:0}.pe-body .pe-pb::after{display:none}
  .pe-marks .pe-body p::after{display:none}
  .pe-body{caret-color:transparent!important}
}
#ped .wb{color:#0F1A30}#ped .wg-l{color:#3E4A63!important;font-weight:500}#ped .wb:disabled{opacity:.45}@media (forced-colors:active){#ped .wb svg{forced-color-adjust:auto}}`;

/* ------------------------------------------------------------ the screen */
const ic = n => `<svg viewBox="0 0 20 20" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${ICO[n] || ''}</svg>`;
const B = (cmd, icon, title, extra = '') => `<button type="button" class="wb" data-cmd="${cmd}" title="${title}" aria-label="${title}" ${extra}>${icon}</button>`;
const BL = (cmd, icon, label, title, extra = '') => `<button type="button" class="wb big" data-cmd="${cmd}" title="${title || label}" ${extra}>${ic(icon)}<span>${label}</span></button>`;
const G = (label, inner, cls = '') => `<div class="wg ${cls}"><div class="wg-b">${inner}</div><div class="wg-l">${label}</div></div>`;
const FONTS = ['Arial', 'David', 'FrankRuehl', 'Heebo', 'Rubik', 'Noto Sans Hebrew', 'Miriam', 'Narkisim', 'Tahoma', 'Segoe UI', 'Calibri', 'Times New Roman', 'Georgia', 'Verdana'];
const SIZES = [8, 10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 96];
const COLORS = ['000000', '404040', '7F7F7F', 'BFBFBF', 'FFFFFF', 'C00000', 'FF0000', 'FFC000', 'FFFF00', '92D050', '00B050', '00B0F0', '0070C0', '002060', '7030A0', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', '5B9BD5', '70AD47', '16213A', 'FFD43B'];
const SHAPES = [['rect', 'מלבן'], ['roundRect', 'מלבן מעוגל'], ['ellipse', 'אליפסה'], ['triangle', 'משולש'], ['diamond', 'מעוין'], ['rightArrow', 'חץ ימינה'], ['leftArrow', 'חץ שמאלה'], ['star5', 'כוכב'], ['wedgeRoundRectCallout', 'בועת דיבור'], ['line', 'קו']];
function ribbonHtml() {
  return `
<div class="wpanel" data-panel="home">
 ${G('ביטול', B('undo', ic('undo'), 'ביטול (Ctrl+Z)') + B('redo', ic('redo'), 'חזרה (Ctrl+Y)'), 'tight')}
 ${G('שקפים', BL('newSlide', 'slideNew', 'שקף חדש', 'שקף חדש (Ctrl+M)', 'data-pop="layouts"') + `<div class="col">${B('dupSlide', ic('copy'), 'שכפול השקף')}${B('delSlide', ic('delTable'), 'מחיקת השקף')}</div>`)}
 ${G('גופן', `<div class="row"><select class="wsel font" data-in="font" title="גופן" aria-label="גופן"></select><input class="wsel size" data-in="size" list="peSizes" inputmode="decimal" title="גודל גופן" aria-label="גודל גופן"><datalist id="peSizes">${SIZES.map(s => `<option value="${s}">`).join('')}</datalist>${B('grow', ic('grow'), 'הגדלת הגופן')}${B('shrink', ic('shrink'), 'הקטנת הגופן')}</div>
   <div class="row">${B('bold', '<b class="gl">B</b>', 'מודגש (Ctrl+B)')}${B('italic', '<i class="gl" style="font-family:Georgia,serif">I</i>', 'נטוי (Ctrl+I)')}${B('underline', '<u class="gl">U</u>', 'קו תחתי (Ctrl+U)')}${B('strike', '<s class="gl">ab</s>', 'קו חוצה')}
   <span class="wsplit">${B('color', '<span class="gl a">A</span><i class="cbar" data-bar="color"></i>', 'צבע גופן')}${B('colorMenu', ic('caret'), 'בחירת צבע גופן', 'data-pop="color"')}</span></div>`)}
 ${G('פסקה', `<div class="row">${B('bullets', ic('bullets'), 'תבליטים')}${B('numbers', ic('numbers'), 'מספור')}${B('rtl', ic('rtl'), 'מימין לשמאל')}${B('ltr', ic('ltr'), 'משמאל לימין')}</div>
   <div class="row">${B('right', ic('aright'), 'יישור לימין')}${B('center', ic('acenter'), 'מרכוז')}${B('left', ic('aleft'), 'יישור לשמאל')}${B('justify', ic('ajust'), 'יישור לשני הצדדים')}</div>`)}
 ${G('סידור', `<div class="col">${B('front', ic('front'), 'הבאה לקדמה')}${B('back', ic('back'), 'שליחה לרקע')}</div>` + B('delShape', ic('trash'), 'מחיקת הצורה (Delete)') + B('dupShape', ic('copy'), 'שכפול הצורה (Ctrl+D)'))}
 ${G('הצגה', BL('playStart', 'play', 'הצגה', 'הצגת המצגת מההתחלה (F5)'))}
</div>
<div class="wpanel" data-panel="insert" hidden>
 ${G('שקפים', BL('newSlide', 'slideNew', 'שקף חדש', 'שקף חדש', 'data-pop="layouts"'))}
 ${G('טקסט', BL('textbox', 'textbox', 'תיבת טקסט', 'הוספת תיבת טקסט'))}
 ${G('איורים', BL('image', 'image', 'תמונה', 'הוספת תמונה מהמכשיר') + BL('shapeMenu', 'shapes', 'צורות', 'הוספת צורה', 'data-pop="shapes"'))}
 ${G('טבלאות', BL('tableMenu', 'table', 'טבלה', 'הוספת טבלה', 'data-pop="table"'))}
</div>
<div class="wpanel" data-panel="design" hidden>
 ${G('צורה', BL('fillMenu', 'shade', 'מילוי', 'צבע המילוי של הצורה', 'data-pop="fill"') + BL('lineMenu', 'borders', 'קו מתאר', 'צבע הקו של הצורה', 'data-pop="line"'))}
 ${G('שקף', BL('bgMenu', 'bg', 'רקע', 'צבע הרקע של השקף', 'data-pop="bg"'))}
</div>
<div class="wpanel" data-panel="show" hidden>
 ${G('הצגת שקפים', BL('playStart', 'play', 'מההתחלה', 'מהשקף הראשון (F5)') + BL('playHere', 'playHere', 'מהשקף הזה', 'מהשקף הנוכחי (Shift+F5)'))}
 ${G('זום', B('zout', ic('zout'), 'הקטנה') + '<span class="zv" id="peZoomV">100%</span>' + B('zin', ic('zin'), 'הגדלה') + '<button type="button" class="wb txt" data-cmd="zfit">התאמה לחלון</button>')}
</div>`;
}
function shellHtml() {
  return `
<div class="pe-top">
  <div class="pe-brand"><span class="pe-mark" aria-hidden="true">P</span><b>סימן</b><span class="pe-sub">עורך PowerPoint</span></div>
  <label class="pe-namebox"><input class="pe-name" id="peName" aria-label="שם הקובץ" dir="auto" spellcheck="false"><span class="pe-ext">.pptx</span></label>
  <span class="pe-state" id="peState" role="status"></span>
  <span class="grow"></span>
  <button type="button" class="pe-tb" id="pePlay" title="הצגת המצגת (F5)">▶ הצגה</button>
  <button type="button" class="pe-tb" id="peDesk" title="פתיחת המצגת לסימון, חתימה ושיתוף">לחתימה ולסימון</button>
  <button type="button" class="pe-tb" id="pePdf" title="שמירת המצגת כקובץ PDF">שמירה כ-PDF</button>
  <button type="button" class="pe-tb pri" id="peSave" title="שמירה כקובץ PowerPoint (Ctrl+S)">שמירה</button>
  <button type="button" class="pe-x" id="peClose" aria-label="סגירת העורך" title="סגירה">×</button>
</div>
<div class="pe-tabs" role="tablist">
  <button type="button" role="tab" data-tab="home" aria-selected="true">בית</button>
  <button type="button" role="tab" data-tab="insert" aria-selected="false">הוספה</button>
  <button type="button" role="tab" data-tab="design" aria-selected="false">עיצוב</button>
  <button type="button" role="tab" data-tab="show" aria-selected="false">הצגת שקפים</button>
</div>
<div class="pe-ribbon" id="peRibbon">${ribbonHtml()}</div>
<div class="pe-main">
  <div class="pe-slides" id="peSlides" role="listbox" aria-label="השקפים"></div>
  <div class="pe-stage" id="peStage"><div class="pe-zoom" id="peZoom"></div></div>
</div>
<div class="pe-status"><span id="peCount"></span><span class="grow"></span><span class="pe-hint" id="peHint"></span>
  <button type="button" class="wb" data-cmd="zout" aria-label="הקטנה">−</button><input type="range" id="peZoomR" min="20" max="250" step="5" value="100" aria-label="זום"><button type="button" class="wb" data-cmd="zin" aria-label="הגדלה">+</button><span id="peZoomS">100%</span></div>
<div class="pe-pop" id="pePop" hidden></div>
<div class="pe-dlg" id="peDlg" hidden><div class="pe-dlg-in" role="dialog" aria-modal="true"></div></div>
<input type="file" id="peImgFile" accept="image/*" hidden>`;
}
const CSS2 = `
#ped .pe-mark{background:#C43E1C}
#ped .pe-main{flex:1;display:flex;min-height:0}
#ped .pe-slides{width:190px;flex:none;overflow-y:auto;background:#EEF1F6;border-inline-end:1px solid var(--wl);padding:10px 8px 40px;display:flex;flex-direction:column;gap:10px}
#ped .pe-th{position:relative;display:flex;gap:6px;align-items:flex-start;cursor:pointer;border-radius:8px;padding:4px}
#ped .pe-th .n{width:16px;flex:none;font-size:11px;color:var(--wmu);text-align:center;padding-top:2px}
#ped .pe-th .bx{flex:1;position:relative;overflow:hidden;border-radius:3px;background:#fff;box-shadow:0 0 0 1px #CBD2DE;direction:ltr}
#ped .pe-th.on .bx{box-shadow:0 0 0 2.5px #C43E1C}
#ped .pe-th:hover .bx{box-shadow:0 0 0 2px #E59E86}
#ped .pe-th.on:hover .bx{box-shadow:0 0 0 2.5px #C43E1C}
#ped .pe-th .bx>.pe-slide{position:absolute;top:0;left:0;transform-origin:0 0;pointer-events:none}
#ped .pe-th.drop-before::before{content:"";position:absolute;inset-inline:0;top:-6px;height:3px;background:#C43E1C;border-radius:2px}
#ped .pe-th.drop-after::after{content:"";position:absolute;inset-inline:0;bottom:-6px;height:3px;background:#C43E1C;border-radius:2px}
#ped .pe-addsl{border:1px dashed #A9B3C4;background:none;border-radius:8px;padding:8px;cursor:pointer;color:var(--wmu);font-size:13px}
#ped .pe-addsl:hover{border-color:#C43E1C;color:#C43E1C}
#ped .pe-stage{flex:1;overflow:auto;position:relative;direction:ltr;background:#E4E8EF;display:grid;place-items:center;padding:24px}
#ped .pe-zoom{position:relative;box-shadow:0 10px 28px -12px rgba(22,33,58,.45);background:#fff}
#ped .pe-zoom>.pe-slide{position:absolute;top:0;left:0;transform-origin:0 0}
.pe-slide{position:relative;overflow:hidden;background:#fff;direction:ltr;font-family:"Heebo",Arial,sans-serif;color:#000;-webkit-user-select:none;user-select:none}
.pe-slide .pe-layer{position:absolute;inset:0}
.pe-slide .pe-sp{position:absolute;box-sizing:border-box}
.pe-slide .pe-geo{position:absolute;inset:0;width:100%;height:100%;overflow:visible}
.pe-slide img.fill{position:absolute;inset:0;width:100%;height:100%;object-fit:fill}
.pe-slide .pe-tx{position:absolute;inset:0;display:flex;flex-direction:column;box-sizing:border-box;overflow:visible;line-height:1.2}
.pe-slide .pe-tx.nowrap .pp{white-space:nowrap}
.pe-slide .pe-tx.vert{writing-mode:vertical-rl}
.pe-slide .pe-txi{outline:none;min-height:1em;word-wrap:break-word;overflow-wrap:break-word}
.pe-slide .pp{margin:0;white-space:pre-wrap}
.pe-slide .pp[data-bu]:not(.empty)::before{content:attr(data-bu);display:inline-block;min-width:1.1em;text-indent:0;color:var(--bu,inherit);margin-inline-end:.35em}
.pe-slide .pp.hang[data-bu]:not(.empty)::before{width:var(--bw);min-width:0;margin-inline-end:0}
.pe-slide .pe-tbl{border-collapse:collapse;width:100%;height:100%;table-layout:fixed}
.pe-slide .pe-tbl td{position:relative;vertical-align:top;padding:0}
.pe-slide .pe-tbl td .pe-tx{position:relative;inset:auto;min-height:100%}
.pe-slide .pe-missing{display:grid;place-items:center;background:repeating-linear-gradient(45deg,#F1F3F7,#F1F3F7 6px,#E6EAF1 6px,#E6EAF1 12px);color:#7A869C;font:12pt "Heebo",Arial;border:1px dashed #A9B3C4}
#ped .pe-stage .pe-layer.live .pe-sp{cursor:move}
#ped .pe-stage .pe-tx.blank::before{content:attr(data-hint);position:absolute;inset:0;display:flex;align-items:center;justify-content:center;color:#8C96A8;font-size:20pt;text-align:center;border:1.5px dashed #B7C0CF;pointer-events:none;padding:4pt;direction:rtl}
#ped .pe-stage .pe-tx.blank .pe-txi{opacity:0}
#ped .pe-stage .pe-tx.editing{outline:1.5px dashed #2E58CC;outline-offset:1px;cursor:text}
#ped .pe-stage .pe-tx.editing .pe-txi{opacity:1;cursor:text;-webkit-user-select:text;user-select:text}
#ped .pe-stage .pe-tx.editing.blank::before{display:none}
#ped .pe-sel{position:absolute;border:calc(1.5px * var(--hz,1)) solid #2E58CC;pointer-events:none;z-index:50;box-sizing:border-box}
#ped .pe-sel i{position:absolute;width:9px;height:9px;transform:scale(var(--hz,1));background:#fff;border:1.5px solid #2E58CC;border-radius:2px;pointer-events:auto;box-sizing:border-box}
#ped .pe-sel i[data-h=nw]{left:-5px;top:-5px;cursor:nwse-resize}#ped .pe-sel i[data-h=n]{left:calc(50% - 4.5px);top:-5px;cursor:ns-resize}#ped .pe-sel i[data-h=ne]{right:-5px;top:-5px;cursor:nesw-resize}
#ped .pe-sel i[data-h=e]{right:-5px;top:calc(50% - 4.5px);cursor:ew-resize}#ped .pe-sel i[data-h=se]{right:-5px;bottom:-5px;cursor:nwse-resize}#ped .pe-sel i[data-h=s]{left:calc(50% - 4.5px);bottom:-5px;cursor:ns-resize}
#ped .pe-sel i[data-h=sw]{left:-5px;bottom:-5px;cursor:nesw-resize}#ped .pe-sel i[data-h=w]{left:-5px;top:calc(50% - 4.5px);cursor:ew-resize}
#ped .pe-sel.editing{border-style:dashed;border-color:#7F98D9}
#ped .pe-sel.editing i{display:none}
#ped .pe-pop .sw{display:grid;grid-template-columns:repeat(6,24px);gap:5px;padding:4px}
#ped .pe-pop .sw button{width:24px;height:24px;border-radius:5px;border:1px solid rgba(0,0,0,.15);cursor:pointer}
#ped .pe-pop .pl{display:flex;flex-direction:column;min-width:200px}
#ped .pe-pop .pl button{text-align:start;border:0;background:none;padding:8px 12px;border-radius:6px;cursor:pointer;font-size:13.5px}
#ped .pe-pop .pl button:hover{background:#EEF1F6}
#ped .pe-pop .tg{display:grid;grid-template-columns:repeat(8,18px);gap:3px;padding:4px}
#ped .pe-pop .tg i{width:18px;height:18px;border:1px solid #C7D0DE;border-radius:2px;background:#fff;cursor:pointer}
#ped .pe-pop .tg i.on{background:#FCE3D9;border-color:#C43E1C}
#ped .pe-pop .tgl{font-size:12px;color:var(--wmu);text-align:center;padding:4px}
.pe-play{position:fixed;inset:0;z-index:2000;background:#000;display:grid;place-items:center;cursor:none;direction:ltr}
.pe-play .pe-slide{transform-origin:center center;position:absolute}
.pe-play .bar{position:fixed;bottom:14px;left:50%;transform:translateX(-50%);color:#fff;font:13px "Heebo",Arial;opacity:0;transition:opacity .2s;background:rgba(0,0,0,.5);padding:6px 14px;border-radius:999px;direction:rtl}
.pe-play:hover .bar{opacity:.85}
.pe-play.cur{cursor:default}
@media (max-width:760px){#ped .pe-main{flex-direction:column-reverse}#ped .pe-slides{width:auto;height:96px;flex-direction:row;overflow-x:auto;overflow-y:hidden;border-inline-end:0;border-top:1px solid var(--wl);padding:8px}#ped .pe-th{width:130px;flex:none}#ped .pe-addsl{flex:none;width:60px}#ped .pe-stage{padding:10px}#ped .pe-sub,#ped .pe-state{display:none}#ped #peDesk,#ped #pePlay{display:none}}
`;
const ICO2 = {
  slideNew: '<rect x="3" y="4" width="14" height="10" rx="1.5"/><path d="M10 16.5v-3M8.5 15h3"/>',
  front: '<rect x="3" y="3" width="9" height="9" rx="1"/><rect x="8" y="8" width="9" height="9" rx="1" fill="currentColor" fill-opacity=".25"/>',
  back: '<rect x="8" y="8" width="9" height="9" rx="1"/><rect x="3" y="3" width="9" height="9" rx="1" fill="currentColor" fill-opacity=".25"/>',
  trash: '<path d="M4 6h12M8 6V4h4v2M6 6l1 11h6l1-11"/>',
  play: '<path d="M6 4l10 6-10 6z" fill="currentColor"/>',
  playHere: '<rect x="2.5" y="4" width="15" height="10" rx="1.5"/><path d="M8.5 6.8l4 2.2-4 2.2z" fill="currentColor"/>',
  textbox: '<rect x="3" y="4" width="14" height="12" rx="1" stroke-dasharray="2 1.6"/><path d="M7 8h6M10 8v5"/>',
  shapes: '<circle cx="7" cy="7" r="4"/><rect x="9" y="9" width="8" height="8" rx="1"/>',
  bg: '<rect x="3" y="4" width="14" height="12" rx="1.5"/><path d="M3 13l4-4 3 3 3-3 4 4"/>',
};
Object.assign(ICO, ICO2);

/* ------------------------------------------------------------ the editor */
let X = null;
const q = id => X.root.querySelector('#' + id);
const KEYS = new WeakMap();
const keyOfNode = n => { if (!KEYS.has(n)) KEYS.set(n, 'k' + (++keyOfNode.n)); return KEYS.get(n); }; keyOfNode.n = 0;
function model(files) {
  return { files, order: slideList(files), size: slideSize(files), parts: new Map(), themes: new Map(), urls: new Map(), dom: new WeakMap(), rmap: new WeakMap(), uid: 0, keyOf: keyOfNode };
}
const curPath = () => X.order[X.cur];
const ctxCur = () => ctxOf(X, curPath());
function say(t) { const h = q('peHint'); h.textContent = t || ''; clearTimeout(say.t); if (t) say.t = setTimeout(() => { if (X) h.textContent = ''; }, 4000); }
function markDirty() {
  X.dirty = X.hi !== X.cleanHi;
  q('peState').textContent = X.dirty ? 'שינויים שלא נשמרו' : (X.savedOnce ? 'נשמר' : '');
  q('peCount').textContent = `שקף ${X.cur + 1} מתוך ${X.order.length}`;
  q('peRibbon').querySelector('[data-cmd=undo]').disabled = X.hi <= 0;
  q('peRibbon').querySelector('[data-cmd=redo]').disabled = X.hi >= X.hist.length - 1;
  X.root.dispatchEvent(new Event('input', { bubbles: true }));   // the app's tab shows the unsaved-changes dot
}
function pushHist() {
  X.hist = X.hist.slice(0, X.hi + 1); X.hist.push({ files: new Map(X.files), order: [...X.order], cur: X.cur }); X.hi++;
  if (X.hist.length > 80) { X.hist.shift(); X.hi--; X.cleanHi--; }
}
function commit() { X.files.set(curPath(), ser(X.doc)); pushHist(); markDirty(); thumbSoon(X.cur); }
function restore(s) {
  X.files = new Map(s.files); X.order = [...s.order]; X.cur = Math.min(s.cur, X.order.length - 1); X.parts.delete('ppt/presentation.xml');
  X.doc = parseXml(X.files.get(curPath())); X.sel = null; X.editing = null; renderStage(); renderThumbs(); markDirty();
}
function undo() { commitText(); if (X.hi <= 0) return; X.hi--; restore(X.hist[X.hi]); }
function redo() { commitText(); if (X.hi >= X.hist.length - 1) return; X.hi++; restore(X.hist[X.hi]); }

/* ---------- the stage */
function fitZoom() {
  const st = q('peStage'), wpx = X.size.w * 4 / 3, hpx = X.size.h * 4 / 3;
  const z = Math.max(.15, Math.min((st.clientWidth - 48) / wpx, (st.clientHeight - 48) / hpx));
  setZoom(z, true);
}
function setZoom(z, fit) {
  X.zoom = Math.max(.15, Math.min(3, z)); X.fit = !!fit;
  const zw = q('peZoom'); Object.assign(zw.style, { width: X.size.w * X.zoom + 'pt', height: X.size.h * X.zoom + 'pt' });
  const s = zw.querySelector('.pe-slide'); if (s) { s.style.transform = `scale(${X.zoom})`; s.style.setProperty('--hz', 1 / X.zoom); }
  const p = Math.round(X.zoom * 100) + '%'; q('peZoomS').textContent = p; const v = q('peZoomV'); if (v) v.textContent = p; q('peZoomR').value = Math.round(X.zoom * 100);
}
function renderStage() {
  X.dom = new WeakMap(); X.rmap = new WeakMap();
  const zw = q('peZoom'); zw.innerHTML = '';
  if (!X.doc) return;
  const el = renderSlide(X, curPath(), { edit: true, doc: X.doc, index: X.cur });
  el.style.transform = `scale(${X.zoom})`; el.style.setProperty('--hz', 1 / X.zoom); zw.append(el); X.slideEl = el;
  const sb = document.createElement('div'); sb.className = 'pe-sel'; sb.hidden = true; sb.innerHTML = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'].map(h => `<i data-h="${h}"></i>`).join('');
  el.append(sb); X.selBox = sb;
  if (X.sel && !findTop(X.sel)) X.sel = null;
  drawSel(); updateUI();
}
function findTop(key) { for (const e of X.slideEl.querySelectorAll(`.pe-layer.live [data-k="${key}"]`)) { const m = X.dom.get(e); if (m) return m.top; } return null; }
const shapeEls = key => [...X.slideEl.querySelectorAll(`.pe-layer.live [data-k="${key}"]`)];
function drawSel() {
  const sb = X.selBox; if (!sb) return;
  const top = X.sel && findTop(X.sel);
  if (!top) { sb.hidden = true; return; }
  const b = topBox(X, ctxCur(), top);
  Object.assign(sb.style, { left: b.x + 'pt', top: b.y + 'pt', width: Math.max(b.w, 4) + 'pt', height: Math.max(b.h, 4) + 'pt', transform: b.rot ? `rotate(${b.rot}deg)` : '', translate: '' });
  sb.classList.toggle('editing', !!X.editing);
  sb.hidden = false;
}
function select(key) { if (X.editing) exitEdit(); X.sel = key; drawSel(); updateUI(); }

/* ---------- text editing */
function enterEdit(key, pt) {
  const top = findTop(key); if (!top) return;
  let els = shapeEls(key);
  let tx = null;
  if (pt) { const hit = document.elementFromPoint(pt.x, pt.y); tx = hit && hit.closest('.pe-tx'); if (tx && !els.some(e => e.contains(tx))) tx = null; }
  if (!tx) for (const e of els) { tx = e.querySelector('.pe-tx'); if (tx) break; }
  if (!tx) {
    if (top.localName !== 'sp') return;
    ensureTxBody(top); X.files.set(curPath(), ser(X.doc)); renderStage(); X.sel = key; drawSel();
    els = shapeEls(key); for (const e of els) { tx = e.querySelector('.pe-tx'); if (tx) break; } if (!tx) return; pt = null;
  }
  const m = X.dom.get(tx); if (!m) return;
  X.editing = { tx, txBody: m.tx || ensureTxBody(m.node), key };
  if (!m.tx && m.tc) { X.editing.txBody = kid(m.tc, 'txBody') || (() => { const t = aNew(m.tc.ownerDocument, 'txBody'); t.append(aNew(m.tc.ownerDocument, 'bodyPr'), aNew(m.tc.ownerDocument, 'lstStyle'), aNew(m.tc.ownerDocument, 'p')); m.tc.prepend(t); return t; })(); }
  const wasBlank = tx.classList.contains('blank');
  tx.classList.add('editing'); tx.classList.remove('blank');
  const inner = tx.querySelector('.pe-txi'); inner.contentEditable = 'true'; inner.spellcheck = true; inner.focus({ preventScroll: true });
  const sel = getSelection(); sel.removeAllRanges();
  let r = null;
  if (pt && !wasBlank && document.caretRangeFromPoint) { r = document.caretRangeFromPoint(pt.x, pt.y); if (r && !inner.contains(r.startContainer)) r = null; }
  if (!r) { r = document.createRange(); if (wasBlank) { const p = inner.querySelector('.pp') || inner; r.selectNodeContents(p); r.collapse(true); } else { r.selectNodeContents(inner); r.collapse(false); } }
  sel.addRange(r);
  drawSel(); updateUI();
}
function commitText() {
  if (!X || !X.editing || !X.txDirty) return;
  clearTimeout(X.txT); X.txDirty = false;
  try { writeTx(X, X.editing.tx, X.editing.txBody); commit(); } catch (e) { console.error('text', e); }
}
function exitEdit() {
  if (!X.editing) return;
  commitText();
  const { tx } = X.editing; X.editing = null;
  const inner = tx.querySelector('.pe-txi'); inner.contentEditable = 'false'; tx.classList.remove('editing');
  getSelection().removeAllRanges();
  if (!inner.textContent.replace(/\u200B/g, '').trim()) renderStage(); else drawSel();
  updateUI();
}

/* ---------- formatting */
// what a command works on: the selected text while typing, or all the text of the selected shape
function scope() {
  if (X.editing) return { txs: [X.editing.tx], range: getSelection().rangeCount ? getSelection().getRangeAt(0) : null };
  if (!X.sel) return null;
  const txs = shapeEls(X.sel).flatMap(e => [...e.querySelectorAll('.pe-tx')]);
  return txs.length ? { txs, range: null } : null;
}
const parasIn = (sc) => { const all = sc.txs.flatMap(t => [...t.querySelectorAll('.pp')]); if (!sc.range) return all; return all.filter(p => sc.range.intersectsNode(p)); };
function stripProp(root, kind) {
  for (const el of root.querySelectorAll ? root.querySelectorAll('*') : []) {
    if (kind === 'b' && /^(B|STRONG)$/.test(el.tagName)) { el.replaceWith(...el.childNodes); continue; }
    if (kind === 'i' && /^(I|EM)$/.test(el.tagName)) { el.replaceWith(...el.childNodes); continue; }
    if (kind === 'u' && el.tagName === 'U') { el.replaceWith(...el.childNodes); continue; }
    if (kind === 'b') el.style.fontWeight = ''; else if (kind === 'i') el.style.fontStyle = '';
    else if (kind === 'u' || kind === 's') el.style.textDecoration = '';
    else if (kind === 'sz') { el.style.fontSize = ''; delete el.dataset.sz; }
    else if (kind === 'color') { el.style.color = ''; delete el.dataset.color; }
    else if (kind === 'font') { el.style.fontFamily = ''; delete el.dataset.font; }
  }
}
function wrapStyle(kind, css, data) {
  const sc = scope(); if (!sc) { say('בוחרים צורה עם טקסט, או לוחצים בתוך הטקסט'); return false; }
  const mk = () => { const s = document.createElement('span'); Object.assign(s.style, css); Object.assign(s.dataset, data || {}); return s; };
  if (sc.range && sc.range.collapsed) {   // the next typed letters get the format
    const s = mk(); s.textContent = '\u200B'; sc.range.insertNode(s); const r = document.createRange(); r.setStart(s.firstChild, 1); r.collapse(true); getSelection().removeAllRanges(); getSelection().addRange(r); X.txDirty = true; return true;
  }
  const made = [];
  for (const p of parasIn(sc)) {
    const r = document.createRange();
    if (sc.range) { r.selectNodeContents(p); if (sc.range.compareBoundaryPoints(Range.START_TO_START, r) > 0) r.setStart(sc.range.startContainer, sc.range.startOffset); if (sc.range.compareBoundaryPoints(Range.END_TO_END, r) < 0) r.setEnd(sc.range.endContainer, sc.range.endOffset); }
    else r.selectNodeContents(p);
    if (r.collapsed) continue;
    const frag = r.extractContents(); const tmp = document.createElement('div'); tmp.append(frag);
    const trail = tmp.lastChild && tmp.lastChild.nodeName === 'BR' && !sc.range ? tmp.lastChild : null; if (trail) trail.remove();
    stripProp(tmp, kind); const s = mk(); s.append(...tmp.childNodes); r.insertNode(s); if (trail) p.append(trail); made.push(s);
  }
  if (sc.range && made.length) { const r = document.createRange(); r.setStartBefore(made[0]); r.setEndAfter(made[made.length - 1]); getSelection().removeAllRanges(); getSelection().addRange(r); }
  return finish(sc);
}
function finish(sc) {
  if (X.editing) { X.txDirty = true; commitText(); return true; }
  for (const tx of sc.txs) { const m = X.dom.get(tx); if (!m) continue; const tb = m.tx || (m.tc && kid(m.tc, 'txBody')); if (tb) writeTx(X, tx, tb); }
  commit(); renderStage(); return true;
}
function toggle(kind) {
  const sc = scope(); if (!sc) { say('בוחרים צורה עם טקסט, או לוחצים בתוך הטקסט'); return; }
  const probe = sc.range ? (sc.range.startContainer.nodeType === 1 ? sc.range.startContainer : sc.range.startContainer.parentNode) : sc.txs[0].querySelector('span') || sc.txs[0].querySelector('.pp');
  const cs = probe && getComputedStyle(probe);
  const on = !cs ? true : kind === 'b' ? !(+cs.fontWeight >= 600) : kind === 'i' ? cs.fontStyle !== 'italic' : kind === 'u' ? !/underline/.test(cs.textDecorationLine) : !/line-through/.test(cs.textDecorationLine);
  const css = kind === 'b' ? { fontWeight: on ? '700' : '400' } : kind === 'i' ? { fontStyle: on ? 'italic' : 'normal' } : { textDecoration: on ? (kind === 'u' ? 'underline' : 'line-through') : 'none' };
  wrapStyle(kind, css);
}
function curSizePt() { const sc = scope(); if (!sc) return 18; const n = sc.range ? (sc.range.startContainer.nodeType === 1 ? sc.range.startContainer : sc.range.startContainer.parentNode) : (sc.txs[0].querySelector('span') || sc.txs[0].querySelector('.pp')); return n ? Math.round(parseFloat(getComputedStyle(n).fontSize) * .75) : 18; }
const setSize = pt => { pt = Math.max(1, Math.min(400, +pt || 0)); if (!pt) return; wrapStyle('sz', { fontSize: pt + 'pt' }, { sz: String(Math.round(pt * 100)) }); };
const setColor = hex => { X.lastColor = hex; q('peRibbon').querySelector('[data-bar=color]').style.background = '#' + hex; wrapStyle('color', { color: '#' + hex }, { color: hex }); };
const setFont = f => wrapStyle('font', { fontFamily: cssFont(f) }, { font: f });
function paraSet(fn) {
  const sc = scope(); if (!sc) { say('בוחרים צורה עם טקסט, או לוחצים בתוך הטקסט'); return; }
  const ps = parasIn(sc); if (!ps.length) return;
  ps.forEach(fn); if (X.editing) { X.txDirty = true; commitText(); } else finish(sc);
}
const align = a => paraSet(p => { p.style.textAlign = a; p.dataset.algn = '1'; });
const setDir = d => paraSet(p => { p.dir = d; p.dataset.rtl = '1'; p.style.textAlign = d === 'rtl' ? 'right' : 'left'; p.dataset.algn = '1'; });
function bullets(kind) {
  const sc = scope(); if (!sc) return; const ps = parasIn(sc);
  const all = ps.length && ps.every(p => kind === 'num' ? /^\d/.test(p.dataset.bu || '') : p.dataset.bu && !/^\d/.test(p.dataset.bu));
  let n = 0;
  paraSet(p => { if (all) { p.dataset.buset = 'none'; delete p.dataset.bu; } else { p.dataset.buset = kind === 'num' ? 'num' : '•'; p.dataset.bu = kind === 'num' ? `${++n}.` : '•'; } });
}

/* ---------- shapes */
function insertCentered(fn, w, h) {
  if (X.editing) exitEdit();
  const W = X.size.w, H = X.size.h; let b = { x: (W - w) / 2, y: (H - h) / 2, w, h };
  const tree = X.doc.getElementsByTagNameNS(NS.p, 'spTree')[0];
  for (let k = 0; k < 8 && kids(tree).some(n => { if (!/^(sp|pic|graphicFrame|cxnSp)$/.test(n.localName)) return false; const o = topBox(X, ctxCur(), n); return Math.abs(o.x - b.x) < 4 && Math.abs(o.y - b.y) < 4; }); k++) b = { ...b, x: b.x + 18, y: b.y + 18 };
  const n = fn(b); commit(); renderStage(); X.sel = keyOfNode(n); drawSel(); updateUI(); return n;
}
function deleteSel() { const top = X.sel && findTop(X.sel); if (!top) return; top.remove(); X.sel = null; commit(); renderStage(); }
function duplicateSel() {
  const top = X.sel && findTop(X.sel); if (!top) return;
  const c = top.cloneNode(true); let id = nextId(X.doc); for (const e of c.getElementsByTagNameNS(NS.p, 'cNvPr')) e.setAttribute('id', id++);
  top.after(c); const b = topBox(X, ctxCur(), c); setBox(c, ctxCur(), { ...b, x: b.x + 12, y: b.y + 12 });
  commit(); renderStage(); X.sel = keyOfNode(c); drawSel();
}
function reorder(front) {
  const top = X.sel && findTop(X.sel); if (!top) return; const tree = top.parentNode;
  if (front) { const ext = kids(tree, 'extLst')[0]; tree.insertBefore(top, ext || null); }
  else { const first = kids(tree).find(k => !/^(nvGrpSpPr|grpSpPr)$/.test(k.localName) && k !== top); if (first) tree.insertBefore(top, first); }
  commit(); renderStage();
}
async function insertImage(file) {
  if (!file || !/^image\//.test(file.type)) return;
  let ext = (file.type.split('/')[1] || 'png').replace('jpeg', 'jpg').replace('svg+xml', 'svg');
  let bytes = new Uint8Array(await file.arrayBuffer()), mime = file.type;
  if (!/^(png|jpg|gif|bmp|svg)$/.test(ext)) {   // webp, heic…: as PNG, which PowerPoint always shows
    const bmp = await createImageBitmap(file); const c = document.createElement('canvas'); c.width = bmp.width; c.height = bmp.height; c.getContext('2d').drawImage(bmp, 0, 0);
    bytes = new Uint8Array(await (await new Promise(r => c.toBlob(r, 'image/png'))).arrayBuffer()); ext = 'png'; mime = 'image/png';
  }
  let w = 300, h = 200;
  try { const bmp = await createImageBitmap(new Blob([bytes], { type: mime })); w = bmp.width; h = bmp.height; } catch (_) {}
  const k = Math.min(1, X.size.w * .6 / (w * .75), X.size.h * .6 / (h * .75)); w = w * .75 * k; h = h * .75 * k;
  const rid = addMedia(X, curPath(), bytes, ext, mime);
  X.parts.clear();
  insertCentered(b => addPictureXml(X.doc, rid, b, file.name), w, h);
}

/* ---------- slides */
let thumbT = 0; const thumbDirty = new Set();
function thumbSoon(i) { thumbDirty.add(i); clearTimeout(thumbT); thumbT = setTimeout(() => { if (X) for (const k of thumbDirty) drawThumb(k); thumbDirty.clear(); }, 400); }
function drawThumb(i) {
  if (!X) return;
  const th = q('peSlides').querySelector(`.pe-th[data-i="${i}"] .bx`); if (!th || !X.order[i]) return;
  const w = th.clientWidth || 150; th.style.height = (w * X.size.h / X.size.w) + 'px';
  th.innerHTML = '';
  try { const el = renderSlide(X, X.order[i], { index: i }); el.style.transform = `scale(${w / (X.size.w * 4 / 3)})`; th.append(el); } catch (e) { console.warn('thumb', e); }
}
function renderThumbs() {
  const box = q('peSlides'); box.innerHTML = '';
  X.order.forEach((p, i) => {
    const d = document.createElement('div'); d.className = 'pe-th' + (i === X.cur ? ' on' : ''); d.dataset.i = i; d.draggable = true; d.setAttribute('role', 'option'); d.setAttribute('aria-selected', i === X.cur); d.tabIndex = i === X.cur ? 0 : -1;
    d.innerHTML = `<span class="n">${i + 1}</span><div class="bx"></div>`; box.append(d);
  });
  const add = document.createElement('button'); add.type = 'button'; add.className = 'pe-addsl'; add.textContent = '+ שקף חדש'; add.dataset.pop = 'layouts'; box.append(add);
  requestAnimationFrame(() => { if (X) X.order.forEach((_, i) => drawThumb(i)); });
}
function goSlide(i) {
  if (!X || i < 0 || i >= X.order.length) return;
  if (X.editing) exitEdit();
  X.cur = i; X.doc = parseXml(X.files.get(curPath())); X.sel = null; renderStage();
  q('peSlides').querySelectorAll('.pe-th').forEach(t => { const on = +t.dataset.i === i; t.classList.toggle('on', on); t.setAttribute('aria-selected', on); t.tabIndex = on ? 0 : -1; if (on) t.scrollIntoView({ block: 'nearest', inline: 'nearest' }); });
  markDirty();
}
function slideOp(fn) { if (X.editing) exitEdit(); fn(); X.parts.delete('ppt/presentation.xml'); pushHist(); renderThumbs(); goSlide(Math.min(X.cur, X.order.length - 1)); }
function newSlide(layoutPath) {
  slideOp(() => {
    const L = layoutPath || (layoutsOf(X, curPath()).find(l => l.type === 'obj') || {}).path || [...ctxCur().sR.values()].find(r => r.type === RT.layout).target;
    const p = addSlide(X, X.cur, L); X.cur = X.order.indexOf(p);
    // a layout with no placeholders: a title box to start from
    const d = parseXml(X.files.get(p)); if (!d.getElementsByTagNameNS(NS.p, 'sp').length) { addTextBox(d, { x: X.size.w * .06, y: X.size.h * .06, w: X.size.w * .88, h: 60 }, 'כותרת'); X.files.set(p, ser(d)); }
  });
}
function dupSlide() { slideOp(() => { const p = duplicateSlide(X, X.cur); X.cur = X.order.indexOf(p); }); }
function delSlide() { if (X.order.length < 2) { say('במצגת חייב להיות לפחות שקף אחד'); return; } slideOp(() => deleteSlide(X, X.cur)); }
function moveSlide(from, to) { if (from === to) return; slideOp(() => { const [p] = X.order.splice(from, 1); X.order.splice(to, 0, p); writeOrder(X.files, X.order); X.cur = to; }); }

/* ---------- popups */
function pop(kind, anchor) {
  const P = q('pePop'); P.innerHTML = '';
  const sw = (fn, none) => `<div class="sw">${COLORS.map(c => `<button type="button" data-c="${c}" title="#${c}" style="background:#${c}"></button>`).join('')}</div>${none ? `<div class="pl"><button type="button" data-c="">${none}</button></div>` : ''}`;
  if (kind === 'color') P.innerHTML = sw();
  else if (kind === 'fill') P.innerHTML = sw(0, 'ללא מילוי');
  else if (kind === 'line') P.innerHTML = sw(0, 'ללא קו');
  else if (kind === 'bg') P.innerHTML = sw(0, 'הרקע של התבנית');
  else if (kind === 'shapes') P.innerHTML = `<div class="pl">${SHAPES.map(([k, l]) => `<button type="button" data-s="${k}">${l}</button>`).join('')}</div>`;
  else if (kind === 'layouts') P.innerHTML = `<div class="pl">${layoutsOf(X, curPath()).map(l => `<button type="button" data-l="${esc(l.path)}">${esc(l.name)}</button>`).join('') || '<button type="button" data-l="">שקף ריק</button>'}</div>`;
  else if (kind === 'table') P.innerHTML = `<div class="tg">${Array.from({ length: 48 }, (_, i) => `<i data-r="${Math.floor(i / 8) + 1}" data-c="${i % 8 + 1}"></i>`).join('')}</div><div class="tgl">טבלה</div>`;
  P.hidden = false; P.dataset.kind = kind;
  const r = anchor.getBoundingClientRect(), R = X.root.getBoundingClientRect();
  P.style.top = (r.bottom - R.top + 4) + 'px'; P.style.left = Math.max(6, Math.min(R.width - P.offsetWidth - 6, r.right - R.left - P.offsetWidth)) + 'px';
}
function popPick(e) {
  const P = q('pePop'), k = P.dataset.kind; const t = e.target;
  if (k === 'table') { const c = t.closest('i'); if (!c) return; P.hidden = true; const rows = +c.dataset.r, cols = +c.dataset.c; insertCentered(b => addTableXml(X.doc, rows, cols, b), Math.min(X.size.w * .8, cols * 110), rows * 29.2); return; }
  const b = t.closest('button'); if (!b) return; P.hidden = true;
  if (k === 'color') setColor(b.dataset.c || '000000');
  else if (k === 'fill' || k === 'line') { const top = X.sel && findTop(X.sel); if (!top || !kid(top, 'spPr')) { say('בוחרים קודם צורה'); return; } (k === 'fill' ? setFill : setLine)(top, b.dataset.c || null); commit(); renderStage(); }
  else if (k === 'bg') { setBackground(X.doc, b.dataset.c || null); commit(); renderStage(); }
  else if (k === 'shapes') insertCentered(bx => addShapeXml(X.doc, b.dataset.s, bx), 180, b.dataset.s === 'line' ? 2 : 120);
  else if (k === 'layouts') newSlide(b.dataset.l || null);
}
function tableHover(e) { const c = e.target.closest && e.target.closest('.tg i'); if (!c) return; const r = +c.dataset.r, cc = +c.dataset.c; for (const i of q('pePop').querySelectorAll('.tg i')) i.classList.toggle('on', +i.dataset.r <= r && +i.dataset.c <= cc); q('pePop').querySelector('.tgl').textContent = `${cc} × ${r}`; }

/* ---------- the ribbon state */
function updateUI() {
  if (!X) return;
  const sc = scope(), on = (cmd, v) => { const b = q('peRibbon').querySelectorAll(`[data-cmd=${cmd}]`); b.forEach(x => x.setAttribute('aria-pressed', v ? 'true' : 'false')); };
  let n = null;
  if (sc) n = sc.range ? (sc.range.startContainer.nodeType === 1 ? sc.range.startContainer : sc.range.startContainer.parentNode) : (sc.txs[0].querySelector('span') || sc.txs[0].querySelector('.pp'));
  const cs = n && getComputedStyle(n);
  on('bold', cs && +cs.fontWeight >= 600); on('italic', cs && cs.fontStyle === 'italic'); on('underline', cs && /underline/.test(cs.textDecorationLine)); on('strike', cs && /line-through/.test(cs.textDecorationLine));
  const p = n && n.closest && n.closest('.pp');
  const ta = p ? getComputedStyle(p).textAlign : '';
  on('right', ta === 'right'); on('center', ta === 'center'); on('left', ta === 'left'); on('justify', ta === 'justify'); on('rtl', p && p.dir === 'rtl'); on('ltr', p && p.dir === 'ltr');
  const sz = q('peRibbon').querySelector('[data-in=size]'); if (document.activeElement !== sz) sz.value = cs ? Math.round(parseFloat(cs.fontSize) * .75 * 2) / 2 : '';
  const fs = q('peRibbon').querySelector('[data-in=font]');
  if (cs && document.activeElement !== fs) { const fam = (cs.fontFamily.split(',')[0] || '').replace(/["']/g, '').trim(); if (![...fs.options].some(o => o.value === fam) && fam) fs.add(new Option(fam, fam)); fs.value = fam; }
  const has = !!(X.sel && findTop(X.sel));
  for (const c of ['delShape', 'dupShape', 'front', 'back']) q('peRibbon').querySelectorAll(`[data-cmd=${c}]`).forEach(b => b.disabled = !has);
}

/* ---------- pointer on the stage */
function onStageDown(e) {
  if (e.button > 0) return;
  const h = e.target.closest('.pe-sel i');
  if (X.editing && X.editing.tx.contains(e.target)) return;      // inside the text being typed
  if (X.editing) exitEdit();
  if (h && X.sel) { startResize(e, h.dataset.h); return; }
  const el = e.target.closest('.pe-layer.live .pe-sp');
  if (!el) { select(null); return; }
  const key = el.dataset.k, was = X.sel === key;
  if (!was) select(key);
  const top = findTop(key); const b0 = topBox(X, ctxCur(), top);
  const z = X.zoom * 4 / 3, x0 = e.clientX, y0 = e.clientY; let moved = false;
  const els = shapeEls(key);
  const mv = ev => {
    const dx = (ev.clientX - x0) / z, dy = (ev.clientY - y0) / z;
    if (!moved && Math.hypot(ev.clientX - x0, ev.clientY - y0) < 4) return; moved = true;
    for (const x of [...els, X.selBox]) x.style.translate = `${dx}pt ${dy}pt`;
  };
  const up = ev => {
    removeEventListener('pointermove', mv); removeEventListener('pointerup', up);
    if (!moved) { if (was && top.localName !== 'pic') enterEdit(key, { x: ev.clientX, y: ev.clientY }); return; }
    const dx = (ev.clientX - x0) / z, dy = (ev.clientY - y0) / z;
    setBox(top, ctxCur(), { ...b0, x: b0.x + dx, y: b0.y + dy }); commit(); renderStage(); X.sel = key; drawSel();
  };
  addEventListener('pointermove', mv); addEventListener('pointerup', up);
}
function startResize(e, hd) {
  e.preventDefault();
  const top = findTop(X.sel); const b0 = topBox(X, ctxCur(), top); const z = X.zoom * 4 / 3, x0 = e.clientX, y0 = e.clientY;
  const keep = top.localName === 'pic' && hd.length === 2;
  const els = shapeEls(X.sel); const single = els.length === 1 ? els[0] : null;
  let nb = b0;
  const mv = ev => {
    let dx = (ev.clientX - x0) / z, dy = (ev.clientY - y0) / z; let { x, y, w, h } = b0;
    if (/e/.test(hd)) w = b0.w + dx; if (/w/.test(hd)) { w = b0.w - dx; x = b0.x + dx; }
    if (/s/.test(hd)) h = b0.h + dy; if (/n/.test(hd)) { h = b0.h - dy; y = b0.y + dy; }
    if (keep || (ev.shiftKey && hd.length === 2)) { const r = b0.h ? b0.w / b0.h : 1; if (Math.abs(w - b0.w) > Math.abs(h - b0.h) * r) { const nh = w / r; if (/n/.test(hd)) y = b0.y + b0.h - nh; h = nh; } else { const nw = h * r; if (/w/.test(hd)) x = b0.x + b0.w - nw; w = nw; } }
    if (w < 4) { if (/w/.test(hd)) x -= 4 - w; w = 4; } if (h < 2 && top.localName !== 'cxnSp') { if (/n/.test(hd)) y -= 2 - h; h = 2; }
    nb = { ...b0, x, y, w, h };
    Object.assign(X.selBox.style, { left: x + 'pt', top: y + 'pt', width: w + 'pt', height: h + 'pt' });
    if (single) { Object.assign(single.style, { left: x + 'pt', top: y + 'pt', width: w + 'pt', height: h + 'pt' }); const g = single.querySelector('.pe-geo'); if (g) g.setAttribute('viewBox', `0 0 ${Math.max(w, .01)} ${Math.max(h, .01)}`); }
  };
  const up = () => { removeEventListener('pointermove', mv); removeEventListener('pointerup', up); if (nb === b0) return; const k = X.sel; setBox(top, ctxCur(), nb); commit(); renderStage(); X.sel = k; drawSel(); };
  addEventListener('pointermove', mv); addEventListener('pointerup', up);
}

/* ---------- slide show */
function present(from) {
  if (X.editing) exitEdit();
  let i = from || 0;
  const ov = document.createElement('div'); ov.className = 'pe-play'; ov.tabIndex = 0;
  const bar = document.createElement('div'); bar.className = 'bar';
  document.body.append(ov);
  let hideT = 0;
  const show = () => {
    for (const s of ov.querySelectorAll('.pe-slide')) s.remove();
    const el = renderSlide(X, X.order[i], { index: i });
    const k = Math.min(innerWidth / (X.size.w * 4 / 3), innerHeight / (X.size.h * 4 / 3));
    el.style.transform = `scale(${k})`; el.style.left = `calc(50% - ${X.size.w / 2}pt)`; el.style.top = `calc(50% - ${X.size.h / 2}pt)`;
    ov.prepend(el); bar.textContent = `${i + 1} / ${X.order.length} · חצים למעבר · Esc ליציאה`;
  };
  ov.append(bar); show();
  const end = () => { removeEventListener('keydown', key, true); removeEventListener('resize', show); ov.remove(); if (document.fullscreenElement) document.exitFullscreen().catch(() => {}); if (X) goSlide(i); };
  const go = d => { const n = i + d; if (n < 0) return; if (n >= X.order.length) { end(); return; } i = n; show(); };
  const key = e => {
    e.stopPropagation();
    if (e.key === 'Escape') { e.preventDefault(); end(); }
    else if (/^(ArrowLeft|ArrowDown|PageDown| |Enter|n)$/.test(e.key)) { e.preventDefault(); go(1); }
    else if (/^(ArrowRight|ArrowUp|PageUp|Backspace|p)$/.test(e.key)) { e.preventDefault(); go(-1); }
    else if (e.key === 'Home') { i = 0; show(); } else if (e.key === 'End') { i = X.order.length - 1; show(); }
  };
  ov.addEventListener('click', e => go(e.clientX < innerWidth / 2 ? 1 : -1));
  ov.addEventListener('mousemove', () => { ov.classList.add('cur'); clearTimeout(hideT); hideT = setTimeout(() => ov.classList.remove('cur'), 1500); });
  addEventListener('keydown', key, true); addEventListener('resize', show);
  if (ov.requestFullscreen) ov.requestFullscreen().catch(() => {});
  ov.focus();
}

/* ---------- saving */
async function buildPptx(m) {
  m = m || X; const zip = new JSZip();
  zip.file('[Content_Types].xml', m.files.get('[Content_Types].xml'));
  for (const [n, v] of m.files) if (n !== '[Content_Types].xml') zip.file(n, v);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType: PPTX_MIME });
}
// every slide as a picture on a PDF page of the slide's size
async function slidesPdf(m, progress) {
  if (!window.html2canvas) throw new Error('ספריית הציור לא נטענה');
  const { PDFDocument } = PDFLib; const pdf = await PDFDocument.create();
  const host = document.createElement('div'); host.style.cssText = 'position:fixed;left:-100000px;top:0;pointer-events:none';
  document.body.append(host);
  try {
    for (let i = 0; i < m.order.length; i++) {
      progress && progress(`מכין את שקף ${i + 1} מתוך ${m.order.length}…`);
      const el = renderSlide(m, m.order[i], { index: i }); host.innerHTML = ''; host.append(el);
      await Promise.all([...el.querySelectorAll('img')].map(im => im.decode ? im.decode().catch(() => {}) : null)); if (document.fonts) await document.fonts.ready;
      const c = await html2canvas(el, { scale: 2, backgroundColor: '#ffffff', logging: false, useCORS: true });
      const img = await pdf.embedJpg(await (await new Promise(r => c.toBlob(r, 'image/jpeg', .92))).arrayBuffer());
      const pg = pdf.addPage([m.size.w, m.size.h]); pg.drawImage(img, { x: 0, y: 0, width: m.size.w, height: m.size.h });
    }
  } finally { host.remove(); }
  return pdf.save();
}
const fileName = () => ((q('peName').value || '').trim().replace(/[\\/:*?"<>|]+/g, '').replace(/\.pptx$/i, '') || 'מצגת') + '.pptx';
async function doSave() {
  if (!X || X.saving) return; commitText(); X.saving = true;
  try { X.opts.busy(true, 'שומר את המצגת…'); const b = await buildPptx(); X.opts.busy(false);
    if (await X.opts.save(fileName(), b, 'pptx')) { X.cleanHi = X.hi; X.savedOnce = true; markDirty(); X.opts.toast('נשמר: ' + fileName()); } }
  catch (e) { X.opts.busy(false); if (!(e && e.name === 'AbortError')) { console.error(e); X.opts.toast('השמירה נכשלה: ' + (e.message || e)); } }
  finally { if (X) X.saving = false; }
}
async function doPdf() {
  if (!X || X.saving) return; commitText(); X.saving = true;
  try { const b = await slidesPdf(X, m => X.opts.busy(true, m)); X.opts.busy(false); const n = fileName().replace(/\.pptx$/, '.pdf'); if (await X.opts.save(n, b, 'pdf')) X.opts.toast('נשמר: ' + n); }
  catch (e) { X.opts.busy(false); if (!(e && e.name === 'AbortError')) { console.error(e); X.opts.toast('יצירת ה-PDF נכשלה: ' + (e.message || e)); } }
  finally { if (X) X.saving = false; }
}
async function doDesk() {
  if (!X || X.saving) return; commitText(); X.saving = true;
  try { const b = await slidesPdf(X, m => X.opts.busy(true, m)); X.opts.busy(false); const f = new File([b], fileName().replace(/\.pptx$/, '.pdf'), { type: 'application/pdf' }); const o = X.opts; await forceClose(); await o.toDesk(f); }
  catch (e) { console.error(e); if (X) { X.opts.busy(false); X.saving = false; X.opts.toast('לא הצלחתי לפתוח לחתימה: ' + (e.message || e)); } }
}
function ask(title, text, buttons) {
  return new Promise(res => {
    const d = q('peDlg'), inn = d.firstElementChild;
    inn.innerHTML = `<h3>${title}</h3><p>${text}</p><div class="acts">${buttons.map(b => `<button type="button" class="${b.cls || ''}" data-v="${b.v}">${b.l}</button>`).join('')}</div>`;
    d.hidden = false; const fin = v => { d.hidden = true; res(v); };
    inn.querySelectorAll('[data-v]').forEach(b => b.onclick = () => fin(b.dataset.v === 'x' ? null : b.dataset.v));
    d.onclick = e => { if (e.target === d) fin(null); };
    (inn.querySelector('.pri') || inn.querySelector('button')).focus();
  });
}
async function doClose() {
  if (!X) return; commitText();
  if (X.dirty) { const a = await ask('לשמור את השינויים?', `יש שינויים ב-${esc(fileName())} שעוד לא נשמרו.`, [{ v: 'x', l: 'ביטול' }, { v: 'drop', l: 'סגירה בלי לשמור', cls: 'warn' }, { v: 'save', l: 'שמירה', cls: 'pri' }]);
    if (!a) return; if (a === 'save') { await doSave(); if (X && X.dirty) return; } }
  const o = X.opts; await forceClose(); o.onClose && o.onClose();
}
async function forceClose() {
  if (!X) return; const e = X;
  removeEventListener('keydown', e.onKey, true); removeEventListener('paste', e.onPaste, true); removeEventListener('copy', e.onCopy, true); removeEventListener('resize', e.onResize); removeEventListener('beforeunload', e.onUnload);
  document.removeEventListener('selectionchange', e.onSel);
  e.root.remove(); for (const u of e.urls.values()) if (u) URL.revokeObjectURL(u); X = null;
}

/* ---------- wiring */
function wire() {
  const R = q('peRibbon');
  X.root.querySelectorAll('.pe-tabs [data-tab]').forEach(t => t.onclick = () => { X.root.querySelectorAll('.pe-tabs [data-tab]').forEach(x => x.setAttribute('aria-selected', x === t)); R.querySelectorAll('.wpanel').forEach(p => p.hidden = p.dataset.panel !== t.dataset.tab); });
  const fs = R.querySelector('[data-in=font]'); for (const f of FONTS) fs.add(new Option(f, f));
  fs.onchange = () => setFont(fs.value);
  const sz = R.querySelector('[data-in=size]'); sz.onchange = () => setSize(sz.value); sz.onkeydown = e => { if (e.key === 'Enter') { e.preventDefault(); setSize(sz.value); } };
  // keep the text selection when a ribbon button is pressed
  X.root.addEventListener('mousedown', e => { if (X && e.target.closest('.wb,.pe-pop button,.pe-pop i') && X.editing) e.preventDefault(); });
  const run = { undo, redo, bold: () => toggle('b'), italic: () => toggle('i'), underline: () => toggle('u'), strike: () => toggle('s'),
    grow: () => setSize(curSizePt() + (curSizePt() < 20 ? 2 : 4)), shrink: () => setSize(Math.max(6, curSizePt() - (curSizePt() <= 20 ? 2 : 4))),
    color: () => setColor(X.lastColor || 'C00000'), bullets: () => bullets('char'), numbers: () => bullets('num'), rtl: () => setDir('rtl'), ltr: () => setDir('ltr'),
    right: () => align('right'), center: () => align('center'), left: () => align('left'), justify: () => align('justify'),
    front: () => reorder(true), back: () => reorder(false), delShape: deleteSel, dupShape: duplicateSel,
    textbox: () => { const n = insertCentered(b => addTextBox(X.doc, b, 'טקסט'), X.size.w * .5, 50); enterEdit(keyOfNode(n)); document.execCommand('selectAll'); },
    image: () => q('peImgFile').click(), dupSlide, delSlide, playStart: () => present(0), playHere: () => present(X.cur),
    zin: () => setZoom(X.zoom * 1.15), zout: () => setZoom(X.zoom / 1.15), zfit: fitZoom };
  X.root.addEventListener('click', e => {
    if (!X) return;   // the editor closed during this click
    const b = e.target.closest('[data-cmd],[data-pop]'); if (!b || !X.root.contains(b)) { if (!e.target.closest('.pe-pop')) q('pePop').hidden = true; return; }
    if (b.closest('.pe-pop')) return;
    if (b.dataset.pop) { e.stopPropagation(); const P = q('pePop'); if (!P.hidden && P.dataset.kind === b.dataset.pop) { P.hidden = true; return; } pop(b.dataset.pop, b); return; }
    q('pePop').hidden = true; const f = run[b.dataset.cmd]; if (f) f();
  });
  q('pePop').addEventListener('click', popPick); q('pePop').addEventListener('mouseover', tableHover);
  q('peImgFile').onchange = e => { const f = e.target.files[0]; e.target.value = ''; insertImage(f); };
  q('peSave').onclick = doSave; q('pePdf').onclick = doPdf; q('peDesk').onclick = doDesk; q('peClose').onclick = doClose; q('pePlay').onclick = () => present(X.cur);
  q('peZoomR').oninput = e => setZoom(+e.target.value / 100);
  const st = q('peStage');
  st.addEventListener('pointerdown', onStageDown);
  st.addEventListener('dblclick', e => { const el = e.target.closest('.pe-layer.live .pe-sp'); if (el && !X.editing) enterEdit(el.dataset.k, { x: e.clientX, y: e.clientY }); });
  st.addEventListener('input', e => { if (X.editing && X.editing.tx.contains(e.target)) { X.txDirty = true; clearTimeout(X.txT); X.txT = setTimeout(commitText, 700); } });
  // slides panel
  const sl = q('peSlides');
  sl.addEventListener('click', e => { const t = e.target.closest('.pe-th'); if (t) goSlide(+t.dataset.i); });
  sl.addEventListener('contextmenu', e => { const t = e.target.closest('.pe-th'); if (!t) return; e.preventDefault(); goSlide(+t.dataset.i);
    ask('שקף ' + (+t.dataset.i + 1), 'מה לעשות עם השקף?', [{ v: 'x', l: 'ביטול' }, { v: 'del', l: 'מחיקה', cls: 'warn' }, { v: 'dup', l: 'שכפול', cls: 'pri' }]).then(a => { if (a === 'del') delSlide(); else if (a === 'dup') dupSlide(); }); });
  let dragI = null;
  sl.addEventListener('dragstart', e => { const t = e.target.closest('.pe-th'); if (!t) return; dragI = +t.dataset.i; e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', 'slide'); });
  sl.addEventListener('dragover', e => { if (dragI === null) return; e.preventDefault(); e.stopPropagation(); const t = e.target.closest('.pe-th'); sl.querySelectorAll('.drop-before,.drop-after').forEach(x => x.classList.remove('drop-before', 'drop-after')); if (t) { const r = t.getBoundingClientRect(); const after = innerWidth <= 760 ? e.clientX < r.left + r.width / 2 : e.clientY > r.top + r.height / 2; t.classList.add(after ? 'drop-after' : 'drop-before'); } });
  sl.addEventListener('drop', e => { if (dragI === null) return; e.preventDefault(); e.stopPropagation(); const t = e.target.closest('.pe-th'); const from = dragI; dragI = null; if (!t) return; let to = +t.dataset.i + (t.classList.contains('drop-after') ? 1 : 0); if (to > from) to--; moveSlide(from, to); });
  sl.addEventListener('dragend', () => { dragI = null; sl.querySelectorAll('.drop-before,.drop-after').forEach(x => x.classList.remove('drop-before', 'drop-after')); });
  sl.addEventListener('keydown', e => { if (e.target.closest('.pe-th')) { if (e.key === 'ArrowDown' || e.key === 'ArrowLeft') { e.preventDefault(); goSlide(X.cur + 1); q('peSlides').querySelector('.pe-th.on')?.focus(); } else if (e.key === 'ArrowUp' || e.key === 'ArrowRight') { e.preventDefault(); goSlide(X.cur - 1); q('peSlides').querySelector('.pe-th.on')?.focus(); } else if (e.key === 'Delete') { e.preventDefault(); delSlide(); } } });
  X.onSel = () => { if (X && X.editing) updateUI(); };
  document.addEventListener('selectionchange', X.onSel);
  X.onResize = () => { if (X && X.fit) fitZoom(); };
  addEventListener('resize', X.onResize);
  X.onUnload = e => { if (X && X.dirty) { e.preventDefault(); e.returnValue = ''; } };
  addEventListener('beforeunload', X.onUnload);
  X.onKey = e => {
    if (!X || document.querySelector('.pe-play')) return;
    if (!X.root.contains(e.target) && e.target !== document.body) return;
    if (!q('peDlg').hidden) { if (e.key === 'Escape') { e.preventDefault(); q('peDlg').hidden = true; } return; }
    const c = e.ctrlKey || e.metaKey, k = e.key.toLowerCase(), typing = !!X.editing, inField = e.target.closest('input,select,textarea');
    if (c && k === 's') { e.preventDefault(); doSave(); return; }
    if (inField) return;
    if (e.key === 'F5') { e.preventDefault(); present(e.shiftKey ? X.cur : 0); return; }
    if (c && (k === 'z' || k === 'y')) { e.preventDefault(); (k === 'z' && !e.shiftKey ? undo : redo)(); return; }
    if (c && k === 'm') { e.preventDefault(); newSlide(); return; }
    if (typing) {
      if (c && k === 'b') { e.preventDefault(); toggle('b'); } else if (c && k === 'i') { e.preventDefault(); toggle('i'); } else if (c && k === 'u') { e.preventDefault(); toggle('u'); }
      else if (e.key === 'Escape') { e.preventDefault(); const key = X.editing.key; exitEdit(); X.sel = key; drawSel(); }
      return;
    }
    if (e.key === 'Escape') { q('pePop').hidden = true; select(null); return; }
    if (X.sel) {
      if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); deleteSel(); return; }
      if (c && k === 'd') { e.preventDefault(); duplicateSel(); return; }
      if (/^Arrow/.test(e.key)) { e.preventDefault(); const top = findTop(X.sel), b = topBox(X, ctxCur(), top), s = e.shiftKey ? 10 : 1; const dx = e.key === 'ArrowLeft' ? -s : e.key === 'ArrowRight' ? s : 0, dy = e.key === 'ArrowUp' ? -s : e.key === 'ArrowDown' ? s : 0; const key = X.sel; setBox(top, ctxCur(), { ...b, x: b.x + dx, y: b.y + dy }); commit(); renderStage(); X.sel = key; drawSel(); return; }
      if (e.key === 'Enter' || e.key === 'F2') { e.preventDefault(); enterEdit(X.sel); return; }
      if (e.key.length === 1 && !c && !e.altKey) { const key = X.sel; enterEdit(key); return; }   // typing replaces nothing: the letter goes into the text
    } else if (e.key === 'PageDown') { e.preventDefault(); goSlide(X.cur + 1); } else if (e.key === 'PageUp') { e.preventDefault(); goSlide(X.cur - 1); }
  };
  addEventListener('keydown', X.onKey, true);
  X.onPaste = e => {
    if (!X || !X.root.contains(e.target) && e.target !== document.body) return;
    if (e.target.closest && e.target.closest('input,textarea')) return;
    const dt = e.clipboardData; if (!dt) return;
    if (X.editing) { e.preventDefault(); const t = dt.getData('text/plain'); if (t) document.execCommand('insertText', false, t.replace(/\r\n/g, '\n')); return; }
    const img = [...dt.files].find(f => /^image\//.test(f.type)); if (img) { e.preventDefault(); insertImage(img); return; }
    if (X.clip && dt.getData('text/plain') === X.clip.text) { e.preventDefault(); pasteShape(); }
  };
  addEventListener('paste', X.onPaste, true);
  X.onCopy = e => {
    if (!X || X.editing || !X.sel || e.target.closest && e.target.closest('input,textarea')) return;
    const top = findTop(X.sel); if (!top) return; e.preventDefault();
    const text = top.textContent ? [...top.getElementsByTagNameNS(NS.a, 't')].map(t => t.textContent).join(' ') || '​' : '​';
    X.clip = { xml: ser(top), from: curPath(), text }; e.clipboardData.setData('text/plain', text);
    if (e.type === 'cut') deleteSel();
  };
  addEventListener('copy', X.onCopy, true); addEventListener('cut', X.onCopy, true);
}
function pasteShape() {
  const c = X.clip; if (!c) return;
  let el = X.doc.importNode(parseXml(c.xml.replace(/^<([\w:]+)/, `<$1 ${SHAPE_NS}`)).documentElement, true);
  if (c.from !== curPath()) {   // pictures: the new slide needs its own link to the image
    const R = rels(X.files, c.from);
    for (const b of el.getElementsByTagNameNS(NS.a, 'blip')) { const id = b.getAttributeNS(NS.r, 'embed'); const r = id && R.get(id); if (r && !r.external) b.setAttributeNS(NS.r, 'r:embed', addRel(X.files, curPath(), RT.image, r.target)); }
  }
  let id = nextId(X.doc); for (const e of el.getElementsByTagNameNS(NS.p, 'cNvPr')) e.setAttribute('id', id++);
  const tree = X.doc.getElementsByTagNameNS(NS.p, 'spTree')[0]; tree.insertBefore(el, kids(tree, 'extLst')[0] || null);
  if (c.from === curPath()) { const b = topBox(X, ctxCur(), el); setBox(el, ctxCur(), { ...b, x: b.x + 12, y: b.y + 12 }); }
  commit(); renderStage(); X.sel = keyOfNode(el); drawSel();
}

/* ---------- open */
function injectCss() { if (document.getElementById('peCss')) return; const s = document.createElement('style'); s.id = 'peCss'; s.textContent = BASE_CSS + CSS2; document.head.append(s); }
async function blankPptx() {
  if (!window.PptxGenJS) throw new Error('ספריית המצגות לא נטענה');
  const p = new PptxGenJS(); p.layout = 'LAYOUT_WIDE'; p.rtlMode = true;
  const s = p.addSlide();
  s.addText('כותרת המצגת', { x: .8, y: 2.3, w: 11.73, h: 1.4, fontSize: 44, bold: true, align: 'center', valign: 'middle', rtlMode: true, fontFace: 'Arial', color: '16213A' });
  s.addText('כותרת משנה', { x: .8, y: 3.8, w: 11.73, h: .8, fontSize: 24, align: 'center', valign: 'middle', rtlMode: true, fontFace: 'Arial', color: '5D6982' });
  return new Uint8Array(await p.write({ outputType: 'arraybuffer' }));
}
async function open(opts) {
  if (X) await forceClose();
  if (!window.JSZip) throw new Error('ספריות העריכה לא נטענו');
  injectCss();
  const bytes = opts.bytes || await blankPptx();
  const files = await readPkg(bytes);
  const root = document.createElement('div'); root.id = 'ped'; root.innerHTML = shellHtml(); document.body.append(root);
  X = Object.assign(model(files), { opts, root, cur: 0, zoom: 1, fit: true, hist: [], hi: -1, cleanHi: 0, sel: null, editing: null, dirty: false });
  if (!X.order.length) { await forceClose(); throw new Error('לא נמצאו שקפים במצגת'); }
  q('peName').value = (opts.name || 'מצגת חדשה.pptx').replace(/\.pptx?$/i, '');
  wire();
  pushHist(); X.cleanHi = X.hi;
  X.doc = parseXml(X.files.get(curPath()));
  renderThumbs(); await nextFrame(); fitZoom(); renderStage(); markDirty();
  if (opts.restored) { X.cleanHi = -1; markDirty(); }
  if (opts.isNew) q('peName').select();
  say('לחיצה על צורה בוחרת אותה, לחיצה נוספת (או לחיצה כפולה) נכנסת לטקסט');
  return true;
}
// a presentation → PDF, without opening the editor (for "לסימון וחתימה" and the assistant)
async function toPdf(bytes, progress) { const m = model(await readPkg(bytes)); try { return await slidesPdf(m, progress); } finally { for (const u of m.urls.values()) if (u) URL.revokeObjectURL(u); } }
async function textOf(bytes) { const m = bytes ? model(await readPkg(bytes)) : X; return m.order.map((p, i) => `===== שקף ${i + 1} =====\n` + [...parseXml(m.files.get(p)).getElementsByTagNameNS(NS.a, 'p')].map(pp => [...pp.getElementsByTagNameNS(NS.a, 't')].map(t => t.textContent).join('')).filter(Boolean).join('\n')).join('\n\n'); }

/* ---------- the assistant: describes the presentation and makes the changes it is asked for ---------- */
const aiHex = c => { if (c === null || c === false || c === 'none') return null; c = String(c || '').trim().replace(/^#/, ''); if (/^[0-9a-f]{3}$/i.test(c)) c = c.replace(/./g, x => x + x); return /^[0-9a-f]{6}$/i.test(c) ? c.toUpperCase() : undefined; };
const AI_LAYOUT = { title: 'title', cover: 'title', content: 'obj', obj: 'obj', bullets: 'obj', section: 'secHead', titleonly: 'titleOnly', blank: 'blank', two: 'twoObj', twocontent: 'twoObj', comparison: 'twoTxTwoObj', picture: 'picTx' };
const aiShapes = doc => { const tree = doc.getElementsByTagNameNS(NS.p, 'spTree')[0]; return kids(tree).filter(n => /^(sp|pic|graphicFrame|grpSp|cxnSp)$/.test(n.localName)); };
function aiKind(n) {
  if (n.localName === 'pic') return 'תמונה'; if (n.localName === 'graphicFrame') return n.getElementsByTagNameNS(NS.a, 'tbl').length ? 'טבלה' : 'אובייקט';
  if (n.localName === 'grpSp') return 'קבוצה'; if (n.localName === 'cxnSp') return 'קו';
  const ph = phOf(n); if (ph) { const t = at(ph, 'type') || 'body'; return { title: 'כותרת', ctrTitle: 'כותרת', subTitle: 'כותרת משנה', body: 'תוכן', obj: 'תוכן', dt: 'תאריך', ftr: 'כותרת תחתונה', sldNum: 'מספר שקף', pic: 'מקום לתמונה' }[t] || 'תוכן'; }
  const nv = kid(kid(n, 'nvSpPr'), 'cNvSpPr'); if (nv && at(nv, 'txBox') === '1') return 'תיבת טקסט';
  const g = n.getElementsByTagNameNS(NS.a, 'prstGeom')[0]; return 'צורה' + (g ? ' ' + at(g, 'prst') : '');
}
const aiParas = n => [...n.getElementsByTagNameNS(NS.a, 'p')].map(p => [...p.getElementsByTagNameNS(NS.a, 't')].map(t => t.textContent).join(''));
const aiDocOf = i => i === X.cur ? X.doc : parseXml(X.files.get(X.order[i]));
function aiContext() {
  if (!X) return null; if (X.editing) exitEdit();
  const out = []; let total = 0;
  X.order.forEach((p, i) => {
    if (total > 50000) return;
    const doc = aiDocOf(i), ctx = ctxOf(X, p);
    const bg = doc.getElementsByTagNameNS(NS.p, 'bg')[0], bgc = bg && bg.getElementsByTagNameNS(NS.a, 'srgbClr')[0];
    let s = `שקף ${i + 1}${i === X.cur ? ' (פתוח עכשיו)' : ''}${bgc ? `, רקע #${at(bgc, 'val')}` : ''}:`;
    aiShapes(doc).forEach((n, k) => {
      const b = topBox(X, ctx, n), pc = v => Math.round(v);
      const r = n.getElementsByTagNameNS(NS.a, 'rPr')[0], sz = r && at(r, 'sz') ? ' ' + (+at(r, 'sz') / 100) + 'pt' : '';
      const t = aiParas(n).filter(Boolean).join(' / ').slice(0, 300);
      s += `\n  צורה ${k}: ${aiKind(n)}${sz} [x${pc(b.x / X.size.w * 100)}% y${pc(b.y / X.size.h * 100)}% w${pc(b.w / X.size.w * 100)}% h${pc(b.h / X.size.h * 100)}%]${t ? ' "' + t + '"' : ' (ריק)'}`;
    });
    out.push(s); total += s.length;
  });
  const lays = layoutsOf(X, X.order[X.cur]).map(l => l.name).filter((v, i, a) => a.indexOf(v) === i);
  return { kind: 'pptx', name: X.opts.name || '', text: `סוג הקובץ: מצגת PowerPoint. ${X.order.length} שקפים, גודל השקף ${Math.round(X.size.w)}×${Math.round(X.size.h)} נקודות (מיקומים באחוזים מהשקף).\nפריסות זמינות: ${lays.join(', ')}.\n${out.join('\n')}` };
}
function aiSlides(v) {
  const n = X.order.length;
  if (v === undefined || v === null || v === 'current' || v === 'this') return [X.cur];
  if (v === 'all') return X.order.map((_, i) => i);
  if (v === 'last') return [n - 1];
  const list = Array.isArray(v) ? v : [v]; const out = [];
  for (const x of list) { const m = /^(\d+)\s*[-–]\s*(\d+)$/.exec(String(x)); if (m) { for (let i = +m[1]; i <= +m[2]; i++) if (i >= 1 && i <= n) out.push(i - 1); } else if (+x >= 1 && +x <= n) out.push(+x - 1); }
  return [...new Set(out)];
}
function aiTargetsIn(doc, sel) {
  const all = aiShapes(doc);
  if (sel === undefined || sel === null || sel === 'all') return sel === 'all' ? all.filter(n => n.localName === 'sp') : all;
  const list = Array.isArray(sel) ? sel : [sel]; const out = [];
  for (const s of list) {
    if (/^\d+$/.test(String(s))) { const n = all[+s]; if (n) out.push(n); continue; }
    const want = String(s).toLowerCase();
    for (const n of all) { const ph = n.localName === 'sp' && phOf(n); const t = ph ? at(ph, 'type') || 'body' : n.localName === 'sp' ? 'text' : n.localName;
      if ((want === 'title' && /^(title|ctrTitle)$/.test(t)) || (want === 'subtitle' && t === 'subTitle') || (want === 'body' && /^(body|obj)$/.test(t) && ph && !/^(dt|ftr|sldNum)$/.test(t)) || (want === 'text' && n.localName === 'sp') || (want === 'picture' && n.localName === 'pic') || (want === 'table' && n.localName === 'graphicFrame')) out.push(n); }
  }
  return [...new Set(out)];
}
const HEBX = /[\u0590-\u05FF\u0600-\u06FF]/;
function aiSetText(n, paras) {
  if (n.localName !== 'sp') return false;
  const doc = n.ownerDocument; const tb = ensureTxBody(n);
  const ph = phOf(n), isTitle = !!ph && /^(title|ctrTitle|subTitle)$/.test(at(ph, 'type') || '');
  const old = kids(tb, 'p'); const p0 = old[0];
  const pPr0 = p0 && kid(p0, 'pPr'); const r0 = p0 && p0.getElementsByTagNameNS(NS.a, 'rPr')[0]; const e0 = p0 && kid(p0, 'endParaRPr');
  for (const p of old) p.remove();
  const list = (Array.isArray(paras) ? paras : String(paras ?? '').split('\n')).map(x => typeof x === 'object' && x ? x : { text: String(x ?? '') });
  if (!list.length) list.push({ text: '' });
  for (const it of list) {
    const t = String(it.text ?? ''), rtl = HEBX.test(t) || (!/[A-Za-z]/.test(t) && true);
    const p = aNew(doc, 'p'); const pPr = pPr0 ? pPr0.cloneNode(true) : aNew(doc, 'pPr');
    pPr.setAttribute('rtl', rtl ? '1' : '0'); if (+it.level > 0) pPr.setAttribute('lvl', String(Math.min(8, +it.level)));
    // Hebrew body text starts on the right, like in PowerPoint (titles keep the layout's alignment)
    if (!pPr.hasAttribute('algn') && !isTitle) pPr.setAttribute('algn', rtl ? 'r' : 'l');
    if (pPr.attributes.length || pPr.childNodes.length) p.append(pPr);
    if (t) { const r = aNew(doc, 'r'); const rp = r0 ? r0.cloneNode(true) : aNew(doc, 'rPr'); rp.setAttribute('lang', rtl ? 'he-IL' : 'en-US'); rp.removeAttribute('dirty'); r.append(rp); const tt = aNew(doc, 't'); tt.textContent = t; r.append(tt); p.append(r); }
    const e = e0 ? e0.cloneNode(true) : aNew(doc, 'endParaRPr', { lang: rtl ? 'he-IL' : 'en-US' }); p.append(e);
    tb.append(p);
  }
  return true;
}
function aiFormat(n, st) {
  if (!st || typeof st !== 'object') return false;
  const doc = n.ownerDocument; let ok = false;
  const f = {}; if (st.bold !== undefined) f.b = !!st.bold; if (st.italic !== undefined) f.i = !!st.italic; if (st.underline !== undefined) f.u = !!st.underline; if (st.strike !== undefined) f.s = !!st.strike;
  if (+st.size > 0) f.sz = Math.round(Math.min(400, +st.size) * 100); const c = aiHex(st.color); if (c) f.color = c; if (st.font) f.font = String(st.font);
  if (Object.keys(f).length) for (const r of [...n.getElementsByTagNameNS(NS.a, 'r'), ...n.getElementsByTagNameNS(NS.a, 'fld')]) {
    const old = kid(r, 'rPr'); const t = (kid(r, 't') || {}).textContent || '';
    const rp = buildRPr(doc, f, old, t); if (old) old.replaceWith(rp); else r.prepend(rp); ok = true;
    if (old === null && rp.parentNode !== r) r.prepend(rp);
  }
  if (Object.keys(f).length) for (const e of n.getElementsByTagNameNS(NS.a, 'endParaRPr')) { const rp = buildRPr(doc, f, e, ''); const ne = aNew(doc, 'endParaRPr'); for (const a of [...rp.attributes]) ne.setAttribute(a.name, a.value); for (const k of kids(rp)) ne.append(k); e.replaceWith(ne); }
  const AL = { right: 'r', left: 'l', center: 'ctr', justify: 'just', ימין: 'r', שמאל: 'l', מרכז: 'ctr' };
  if (st.align && AL[st.align]) for (const p of n.getElementsByTagNameNS(NS.a, 'p')) { let pPr = kid(p, 'pPr'); if (!pPr) { pPr = aNew(doc, 'pPr'); p.prepend(pPr); } pPr.setAttribute('algn', AL[st.align]); ok = true; }
  const fl = aiHex(st.fill); if (fl !== undefined && n.localName === 'sp') { setFill(n, fl); ok = true; }
  const ln = aiHex(st.line); if (ln !== undefined && n.localName === 'sp') { setLine(n, ln); ok = true; }
  return ok || Object.keys(f).length > 0;
}
const pctBox = (a, def) => { const v = (k, d) => (a[k] !== undefined && isFinite(+a[k]) ? +a[k] : d); return { x: v('x', def.x) / 100 * X.size.w, y: v('y', def.y) / 100 * X.size.h, w: v('w', def.w) / 100 * X.size.w, h: v('h', def.h) / 100 * X.size.h }; };
function aiAddSlide(a, done) {
  const lays = layoutsOf(X, X.order[X.cur]);
  const want = AI_LAYOUT[String(a.layout || (a.bullets || a.body ? 'content' : a.subtitle ? 'title' : 'content')).toLowerCase().replace(/[\s_-]/g, '')] || 'obj';
  const L = (lays.find(l => l.type === want) || lays.find(l => l.type === 'obj') || lays[0] || {}).path || [...ctxOf(X, X.order[X.cur]).sR.values()].find(r => r.type === RT.layout).target;
  const after = a.after === 'start' ? -1 : a.after !== undefined && a.after !== 'end' && isFinite(+a.after) ? Math.max(-1, Math.min(X.order.length - 1, +a.after - 1)) : a.position === 'next' ? X.cur : X.order.length - 1;
  const p = addSlide(X, after, L); const i = X.order.indexOf(p);
  const doc = parseXml(X.files.get(p));
  const T = aiTargetsIn(doc, 'title')[0], S = aiTargetsIn(doc, 'subtitle')[0], B = aiTargetsIn(doc, 'body')[0];
  if (a.title !== undefined) { if (T) aiSetText(T, String(a.title)); else aiSetText(addTextBox(doc, { x: X.size.w * .06, y: X.size.h * .06, w: X.size.w * .88, h: 60 }, String(a.title)), String(a.title)); }
  const body = a.bullets || a.body || a.subtitle;
  if (body !== undefined) { const tgt = (a.subtitle !== undefined && S) || B || S; if (tgt) aiSetText(tgt, Array.isArray(body) ? body : String(body).split('\n')); else addTextBox(doc, { x: X.size.w * .08, y: X.size.h * .28, w: X.size.w * .84, h: X.size.h * .6 }, (Array.isArray(body) ? body : [body]).join('\n')); }
  if (a.subtitle !== undefined && S && (a.bullets || a.body)) aiSetText(S, String(a.subtitle));
  const bgc = aiHex(a.background); if (bgc) setBackground(doc, bgc);
  if (a.style) for (const n of aiTargetsIn(doc, 'all')) aiFormat(n, a.style);
  X.files.set(p, ser(doc)); done.push(i);
  return i;
}
// a table's columns mirrored: the grid and every row, a merged cell (gridSpan + hMerge) moving as one
function aiTableMirror(tbl) {
  const kidsA = (el, n) => el ? [...el.childNodes].filter(x => x.nodeType === 1 && x.namespaceURI === NS.a && x.localName === n) : [];
  const grid = kidsA(tbl, 'tblGrid')[0], cols = kidsA(grid, 'gridCol'); if (cols.length < 2) return false;
  const gx = kidsA(grid, 'extLst')[0] || null; for (let i = cols.length - 1; i >= 0; i--) grid.insertBefore(cols[i], gx);
  for (const tr of kidsA(tbl, 'tr')) {
    const groups = [];
    for (const tc of kidsA(tr, 'tc')) { if (tc.getAttribute('hMerge') === '1' && groups.length) groups[groups.length - 1].push(tc); else groups.push([tc]); }
    const tx = kidsA(tr, 'extLst')[0] || null; for (let g = groups.length - 1; g >= 0; g--) for (const tc of groups[g]) tr.insertBefore(tc, tx);
  }
  return true;
}
function aiApply(actions) {
  if (!X) throw new Error('העורך סגור'); if (X.editing) exitEdit();
  const mark = { hi: X.hi, cur: X.cur }; let done = 0, skipped = 0; let goTo = null;
  X.files.set(X.order[X.cur], ser(X.doc));
  const touched = new Map(); // slide path -> doc
  const docFor = i => { const p = X.order[i]; if (!touched.has(p)) touched.set(p, parseXml(X.files.get(p))); return touched.get(p); };
  const flush = () => { for (const [p, d] of touched) if (X.files.has(p)) X.files.set(p, ser(d)); touched.clear(); };
  for (const a of (Array.isArray(actions) ? actions : []).slice(0, 80)) {
    if (!a || typeof a !== 'object') continue; const op = String(a.op || a.action || '').toLowerCase();
    try {
      if (op === 'addslide' || op === 'newslide') { flush(); const made = []; goTo = aiAddSlide(a, made); done++; }
      else if (op === 'deleteslide') { flush(); const sl = aiSlides(a.slide).sort((x, y) => y - x); for (const i of sl) if (X.order.length > 1) deleteSlide(X, i); sl.length ? done++ : skipped++; }
      else if (op === 'duplicateslide') { flush(); const sl = aiSlides(a.slide); if (sl[0] !== undefined) { const p = duplicateSlide(X, sl[0]); goTo = X.order.indexOf(p); done++; } else skipped++; }
      else if (op === 'moveslide') { flush(); const f = +a.from - 1, t = Math.max(0, Math.min(X.order.length - 1, +a.to - 1)); if (X.order[f]) { const [p] = X.order.splice(f, 1); X.order.splice(t, 0, p); writeOrder(X.files, X.order); done++; } else skipped++; }
      else if (op === 'text' || op === 'settext') { let k = 0; for (const i of aiSlides(a.slide)) { const d = docFor(i); let ts = aiTargetsIn(d, a.shape ?? 'title'); if (!ts.length && (a.shape === 'title' || a.shape === undefined)) ts = [addTextBox(d, { x: X.size.w * .06, y: X.size.h * .06, w: X.size.w * .88, h: 60 }, '')]; for (const n of ts.slice(0, 1)) if (aiSetText(n, a.paragraphs || a.bullets || a.text)) k++; } k ? done++ : skipped++; }
      else if (op === 'format' || op === 'style') { let k = 0; for (const i of aiSlides(a.slide)) { const d = docFor(i); for (const n of aiTargetsIn(d, a.shape ?? 'all')) if (aiFormat(n, a.style || a.format)) k++; } k ? done++ : skipped++; }
      else if (op === 'background') { const c = aiHex(a.color); if (c === undefined) { skipped++; continue; } for (const i of aiSlides(a.slide)) setBackground(docFor(i), c); done++; }
      else if (op === 'addtext' || op === 'textbox') { for (const i of aiSlides(a.slide)) { const d = docFor(i); const n = addTextBox(d, pctBox(a, { x: 10, y: 40, w: 80, h: 15 }), String(a.text || '')); aiSetText(n, String(a.text || '').split('\n')); if (a.style) aiFormat(n, a.style); } done++; }
      else if (op === 'addshape' || op === 'shape') { const prst = /^[a-zA-Z0-9]+$/.test(a.shape || '') ? a.shape : 'rect';
        for (const i of aiSlides(a.slide)) { const d = docFor(i); const n = addShapeXml(d, prst, pctBox(a, { x: 30, y: 35, w: 40, h: 30 })); const c = aiHex(a.color || a.fill); if (c !== undefined) setFill(n, c);
          if (a.text) { aiSetText(n, String(a.text).split('\n')); aiFormat(n, { color: a.textColor || 'FFFFFF', align: 'center', ...(a.style || {}) }); } } done++; }
      else if (op === 'move' || op === 'resize' || op === 'position') { let k = 0; for (const i of aiSlides(a.slide)) { const d = docFor(i); const ctx = ctxOf(X, X.order[i]);
          for (const n of aiTargetsIn(d, a.shape)) { const b = topBox(X, ctx, n); const cur = { x: b.x / X.size.w * 100, y: b.y / X.size.h * 100, w: b.w / X.size.w * 100, h: b.h / X.size.h * 100 }; setBox(n, ctx, pctBox(a, cur)); k++; } } k ? done++ : skipped++; }
      else if (op === 'delete' || op === 'deleteshape' || op === 'remove') { let k = 0; for (const i of aiSlides(a.slide)) for (const n of aiTargetsIn(docFor(i), a.shape)) { n.remove(); k++; } k ? done++ : skipped++; }
      else if (op === 'reversecols' || op === 'reversecolumns' || op === 'mirrorcols' || op === 'flipcols') { let k = 0;
        for (const i of aiSlides(a.slide ?? 'all')) for (const t of docFor(i).getElementsByTagNameNS(NS.a, 'tbl')) if (aiTableMirror(t)) k++; k ? done++ : skipped++; }
      else if (op === 'goto') { const s = aiSlides(a.slide)[0]; if (s !== undefined) goTo = s; }
      else skipped++;
    } catch (e) { console.warn('assistant action', a, e); skipped++; }
  }
  flush();
  X.parts.clear(); X.parts.delete('ppt/presentation.xml');
  if (goTo !== null) X.cur = Math.max(0, Math.min(X.order.length - 1, goTo)); else X.cur = Math.min(X.cur, X.order.length - 1);
  X.doc = parseXml(X.files.get(X.order[X.cur])); X.sel = null;
  pushHist(); renderThumbs(); renderStage(); markDirty();
  q('peSlides').querySelectorAll('.pe-th').forEach(t => t.classList.toggle('on', +t.dataset.i === X.cur));
  return { done, skipped, mark };
}
function aiRevert(mark) {
  if (!X || !mark || !X.hist[mark.hi]) return false; if (X.editing) exitEdit();
  X.hi = mark.hi; restore(X.hist[mark.hi]); renderThumbs(); goSlide(Math.min(mark.cur, X.order.length - 1)); return true;
}
async function shareFile(fmt) {
  if (!X) throw new Error('אין מצגת פתוחה'); commitText();
  if (fmt !== 'pdf') return new File([await buildPptx()], fileName(), { type: PPTX_MIME });
  try { const b = await slidesPdf(X, m => X.opts.busy(true, m)); return new File([b], fileName().replace(/\.pptx$/i, '') + '.pdf', { type: 'application/pdf' }); } finally { X && X.opts.busy(false); }
}
window.PPTX_EDITOR = { open, close: forceClose, isOpen: () => !!X, toPdf, textOf, share: { file: shareFile, name: () => X ? fileName() : '' },
  assist: { context: aiContext, apply: aiApply, revert: aiRevert, root: () => X && X.root },
  _test: { get X() { return X; }, buildPptx: () => buildPptx(), goSlide, select, enterEdit, exitEdit, commitText, newSlide, dupSlide, delSlide, moveSlide, setSize, setColor, setFont, toggle, align, setDir, bullets, deleteSel, duplicateSel, insertImage, present, undo, redo, keyOf: keyOfNode, findTop, slidesPdf: () => slidesPdf(X), setFill: c => { const t = findTop(X.sel); setFill(t, c); commit(); renderStage(); } }
};
})();
