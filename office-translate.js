/* =====================================================================
   סימן — office-translate.js (5.2.0)
   Translating a Word, Excel or PowerPoint file from inside its editor.
   The translation is written into the file itself, so the design stays as it was:
   fonts, sizes, colors, tables, pictures, slides and cells. Each paragraph / cell / text
   box keeps its own formatting; a paragraph that mixes formats (a bold word in a line)
   keeps the formats on the matching translated words.
   Translating into Hebrew or Arabic turns the paragraphs right-to-left (and back for
   English, Russian, French and Spanish).
   The translated file opens in a new tab next to the original, which is not changed.
   Uses the app's translation service (callTranslate in index.html): siman-translate for
   Hebrew/English, siman-ai for the other languages.
   API: window.OFFICE_TR = { open(kind), translateBytes(kind, bytes, target, progress) }
   ===================================================================== */
(() => {
'use strict';
const ED = {
  docx: { id: 'wed', api: 'WORD_EDITOR', p: 'we', build: 'buildDocx', open: 'openWordEditor', app: 'Word' },
  xlsx: { id: 'xed', api: 'XLSX_EDITOR', p: 'xe', build: 'buildXlsx', open: 'openExcelEditor', app: 'Excel' },
  pptx: { id: 'ped', api: 'PPTX_EDITOR', p: 'pe', build: 'buildPptx', open: 'openPptEditor', app: 'PowerPoint' },
};
const LANGS = [['he', 'עברית'], ['en', 'אנגלית'], ['ru', 'רוסית'], ['ar', 'ערבית'], ['fr', 'צרפתית'], ['es', 'ספרדית']];
const SUF = { he: 'עברית', en: 'English', ru: 'Русский', ar: 'العربية', fr: 'Français', es: 'Español' };
const RTL = { he: 1, ar: 1 };
const LETTER = /[A-Za-z\u00C0-\u024F\u0590-\u05FF\u0400-\u04FF\u0600-\u06FF]/;
const RTL_CH = /[\u0590-\u05FF\u0600-\u06FF]/g, LTR_CH = /[A-Za-z\u00C0-\u024F\u0400-\u04FF]/g;
const isRtlText = t => ((t.match(RTL_CH) || []).length) > ((t.match(LTR_CH) || []).length);
const BR = ' ⟦br⟧ ';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const say = m => { if (typeof toast === 'function') toast(m); else console.log(m); };
const wait = (on, m) => { if (typeof busy === 'function') busy(on, m); };
const liveKind = () => Object.keys(ED).find(k => document.getElementById(ED[k].id));

/* ------------------------------------------------------------ reading the text out of the file */
// the element order Word expects inside a paragraph's / run's properties
const PPR = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];
const RPR = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'];
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const kids = (el, name) => [...el.childNodes].filter(n => n.nodeType === 1 && n.nodeName === name);
const kid = (el, name) => kids(el, name)[0] || null;
const up = (el, name, stop) => { for (let n = el.parentNode; n && n !== stop; n = n.parentNode) if (n.nodeName === name) return n; return null; };
const fmtKey = r => { if (!r) return ''; const p = kid(r, r.nodeName === 'a:r' ? 'a:rPr' : r.nodeName === 'w:r' ? 'w:rPr' : 'rPr'); return p ? new XMLSerializer().serializeToString(p).replace(/\s(w:|)(lang|eastAsia|bidi|dirty|err|noProof)="[^"]*"/g, '') : ''; };
// one translatable unit = one paragraph / one cell string: its text pieces, grouped by format
function units(doc, kind) {
  const out = [];
  const make = (root, ts) => {
    ts = ts.filter(t => t.textContent);
    if (!ts.length) return;
    const segs = [];
    for (const t of ts) {
      const run = t.parentNode, key = fmtKey(run);
      const last = segs[segs.length - 1];
      if (last && last.key === key) { last.ts.push(t); last.text += t.textContent; }
      else segs.push({ key, ts: [t], text: t.textContent });
    }
    const text = segs.map(s => s.text).join('');
    if (!LETTER.test(text)) return;
    out.push({ root, segs, text, rtl: isRtlText(text) });
  };
  if (kind === 'docx') {
    for (const p of [...doc.getElementsByTagName('w:p')]) {
      const ts = [...p.getElementsByTagName('w:t')].filter(t => up(t, 'w:p') === p && !up(t, 'w:del', p) && t.parentNode.nodeName === 'w:r');
      make(p, ts);
    }
  } else if (kind === 'pptx') {
    for (const p of [...doc.getElementsByTagName('a:p')]) {
      const ts = [...p.getElementsByTagName('a:t')].filter(t => up(t, 'a:p') === p && t.parentNode.nodeName === 'a:r');
      make(p, ts);
    }
  } else {
    for (const si of [...doc.getElementsByTagName('si'), ...doc.getElementsByTagName('is')]) {
      const ts = [...si.getElementsByTagName('t')].filter(t => !up(t, 'rPh', si));
      make(si, ts);
    }
  }
  return out;
}

/* ------------------------------------------------------------ the translation service */
// a unit with several formats goes with ⟦n⟧ marks, so each format lands on its own translated words
const encode = u => (u.segs.length > 1 ? u.segs.map((s, i) => `⟦${i + 1}⟧${s.text}`).join('') : u.text).replace(/\r?\n/g, BR).replace(/\s+/g, ' ').trim();
function decode(u, t) {
  t = String(t || '').replace(/\s*⟦br⟧\s*/g, '\n');
  if (u.segs.length < 2) return [t.replace(/⟦\d+⟧/g, '')];
  const parts = t.split(/⟦(\d+)⟧/), res = u.segs.map(() => '');
  let ok = 0;
  if (parts[0].trim()) res[0] += parts[0];
  for (let i = 1; i < parts.length; i += 2) { const n = +parts[i] - 1; if (n >= 0 && n < res.length) { res[n] += parts[i + 1] || ''; ok++; } else res[0] += parts[i + 1] || ''; }
  if (!ok) return [t.replace(/⟦\d+⟧/g, '')];
  return res;
}
async function translateTexts(list, target, progress) {
  if (typeof callTranslate !== 'function') throw new Error('שירות התרגום לא נטען');
  const uniq = [...new Set(list)], done = new Map();
  const chunks = []; let cur = [], len = 0;
  for (const s of uniq) { if (cur.length && (len + s.length > 3500 || cur.length >= 60)) { chunks.push(cur); cur = []; len = 0; } cur.push(s); len += s.length + 8; }
  if (cur.length) chunks.push(cur);
  let lang = null;
  for (let c = 0; c < chunks.length; c++) {
    progress && progress(`מתרגם… ${Math.round(c / chunks.length * 100)}%`);
    const ch = chunks[c];
    const text = 'Translate each numbered line separately. Keep every [number] marker at the start of its line and keep the same count of lines. Keep every ⟦number⟧ mark, placed before the translated words it belongs to, and keep every ⟦br⟧ mark.\n' + ch.map((s, k) => `[${k + 1}] ${s}`).join('\n');
    const r = await callTranslate(text, target);
    if (r.error) throw new Error(typeof errText === 'function' ? errText(r) : String(r.error));
    const p = typeof parseTr === 'function' ? parseTr(r.raw) : { text: r.raw, lang: null };
    if (!lang && p.lang) lang = p.lang;
    const got = {}; let last = null;
    for (const line of p.text.split('\n')) {
      const m = /^\s*\[(\d+)\]\s?(.*)$/.exec(line);
      if (m) { last = +m[1]; got[last] = m[2].trim(); } else if (last && line.trim()) got[last] += ' ' + line.trim();
    }
    ch.forEach((s, k) => { if (got[k + 1]) done.set(s, got[k + 1]); });
  }
  return { map: done, lang };
}

/* ------------------------------------------------------------ writing it back */
function setOrdered(parent, el, order, prefix) {
  const name = el.nodeName.replace(prefix, ''), at = order.indexOf(name);
  const old = kid(parent, el.nodeName); if (old) old.remove();
  const before = [...parent.childNodes].find(n => n.nodeType === 1 && order.indexOf(n.nodeName.replace(prefix, '')) > at);
  parent.insertBefore(el, before || null);
}
function ensure(doc, parent, name, order, prefix) {
  let p = kid(parent, name);
  if (!p) { p = doc.createElementNS(W, name); setOrdered(parent, p, order, prefix); }
  return p;
}
function wordDir(doc, u, rtl) {
  const p = u.root;
  if (u.rtl === rtl) return;
  const pPr = kid(p, 'w:pPr') || (() => { const x = doc.createElementNS(W, 'w:pPr'); p.insertBefore(x, p.firstChild); return x; })();
  if (rtl) setOrdered(pPr, doc.createElementNS(W, 'w:bidi'), PPR, 'w:');
  else { const b = kid(pPr, 'w:bidi'); if (b) b.remove(); }
  for (const s of u.segs) for (const t of s.ts) {
    const r = t.parentNode; if (r.nodeName !== 'w:r') continue;
    let rPr = kid(r, 'w:rPr'); if (!rPr) { rPr = doc.createElementNS(W, 'w:rPr'); r.insertBefore(rPr, r.firstChild); }
    if (rtl) setOrdered(rPr, doc.createElementNS(W, 'w:rtl'), RPR, 'w:'); else { const x = kid(rPr, 'w:rtl'); if (x) x.remove(); }
  }
}
function slideDir(doc, u, rtl) {
  if (u.rtl === rtl) return;
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  let pPr = kid(u.root, 'a:pPr'); if (!pPr) { pPr = doc.createElementNS(A, 'a:pPr'); u.root.insertBefore(pPr, u.root.firstChild); }
  if (rtl) pPr.setAttribute('rtl', '1'); else pPr.removeAttribute('rtl');
  const a = pPr.getAttribute('algn');
  if (a === 'l') pPr.setAttribute('algn', 'r'); else if (a === 'r') pPr.setAttribute('algn', 'l'); // centered / inherited alignment stays
}
function writeUnit(doc, kind, u, t, rtl) {
  const parts = decode(u, t);
  u.segs.forEach((s, i) => {
    let txt = parts.length === 1 ? (i === 0 ? parts[0] : '') : parts[i] || '';
    if (parts.length > 1 && i < u.segs.length - 1 && txt && !/\s$/.test(txt) && (parts[i + 1] || '') && !/^\s/.test(parts[i + 1])) txt += ' ';
    s.ts.forEach((el, j) => {
      el.textContent = j === 0 ? txt : '';
      if (kind !== 'pptx' && j === 0) el.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
    });
  });
  if (kind === 'docx') wordDir(doc, u, rtl);
  else if (kind === 'pptx') slideDir(doc, u, rtl);
}

/* ------------------------------------------------------------ the whole file */
const PARTS = {
  docx: n => /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(n),
  pptx: n => /^ppt\/slides\/slide\d+\.xml$/.test(n),
  xlsx: n => n === 'xl/sharedStrings.xml' || /^xl\/worksheets\/sheet\d+\.xml$/.test(n),
};
async function translateBytes(kind, bytes, target, progress) {
  if (!window.JSZip) throw new Error('ספריית הקבצים לא נטענה');
  const zip = await JSZip.loadAsync(bytes);
  const names = Object.keys(zip.files).filter(PARTS[kind]).sort((a, b) => a.localeCompare(b, 'en', { numeric: true }));
  const docs = [], all = [];
  for (const n of names) {
    const x = await zip.file(n).async('string');
    if (kind === 'xlsx' && /worksheets/.test(n) && !/<is>/.test(x)) { docs.push({ n, doc: null, x }); continue; }
    const doc = new DOMParser().parseFromString(x, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) continue;
    const us = units(doc, kind); docs.push({ n, doc, us }); all.push(...us);
  }
  if (!all.length) throw new Error('לא נמצא בקובץ טקסט לתרגום');
  const rtl = !!RTL[target];
  const { map, lang } = await translateTexts(all.map(encode), target, progress);
  if (!map.size) throw new Error('התרגום לא החזיר תוצאה. נסו שוב');
  let n = 0;
  for (const d of docs) {
    if (!d.doc) continue;
    for (const u of d.us) { const t = map.get(encode(u)); if (t) { writeUnit(d.doc, kind, u, t, rtl); n++; } }
    let x = new XMLSerializer().serializeToString(d.doc.documentElement);
    zip.file(d.n, XML_HEAD + x);
  }
  // a sheet turns right-to-left for Hebrew and Arabic, and back for the other languages
  if (kind === 'xlsx') {
    const srcRtl = isRtlText(all.map(u => u.text).join(' '));
    if (srcRtl !== rtl) for (const d of docs) {
      if (!/worksheets/.test(d.n)) continue;
      let x = d.doc ? await zip.file(d.n).async('string') : d.x;
      x = x.replace(/<sheetView\b([^>]*?)(\/?)>/g, (m, a, sl) => `<sheetView${a.replace(/\srightToLeft="[^"]*"/, '')}${rtl ? ' rightToLeft="1"' : ''}${sl}>`);
      zip.file(d.n, x);
    }
  }
  const out = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  return { bytes: out, count: n, total: all.length, lang };
}

/* ------------------------------------------------------------ in the editor */
const CSS = `
.otr-dlg{position:fixed;inset:0;z-index:2147483000;display:grid;place-items:center;background:rgba(16,22,38,.45);direction:rtl;font-family:"SF-heebo",Heebo,Arial,sans-serif}
.otr-box{background:#fff;color:#1B1C1F;border-radius:14px;box-shadow:0 20px 50px -12px rgba(0,0,0,.45);width:min(420px,92vw);padding:20px 20px 16px}
.otr-box h3{margin:0 0 4px;font-size:18px}
.otr-box p{margin:0 0 14px;font-size:13.5px;color:#5A6273;line-height:1.5}
.otr-langs{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:14px}
.otr-langs button{font:inherit;font-size:14.5px;padding:10px 6px;border:1px solid #CBD3DF;border-radius:10px;background:#F6F8FB;color:#1B1C1F;cursor:pointer}
.otr-langs button:hover,.otr-langs button:focus-visible{border-color:#E0A100;background:#FFF6D6;outline:none}
.otr-acts{display:flex;justify-content:flex-start}
.otr-acts button{font:inherit;font-size:14px;padding:7px 14px;border-radius:8px;border:1px solid #CBD3DF;background:#fff;cursor:pointer;color:#1B1C1F}
@media (max-width:720px){#weTr,#xeTr,#peTr{display:none}}
`;
function injectCss() { if (document.getElementById('otrCss')) return; const s = document.createElement('style'); s.id = 'otrCss'; s.textContent = CSS; document.head.append(s); }
function pickLang(app) {
  injectCss();
  return new Promise(res => {
    const d = document.createElement('div'); d.className = 'otr-dlg'; d.setAttribute('role', 'dialog'); d.setAttribute('aria-modal', 'true'); d.setAttribute('aria-labelledby', 'otrT');
    d.innerHTML = `<div class="otr-box"><h3 id="otrT">תרגום הקובץ</h3><p>לאיזו שפה לתרגם? התרגום נכתב בתוך קובץ ה-${app} עצמו ושומר על העיצוב. הקובץ המתורגם נפתח בלשונית חדשה, והמקור לא משתנה.</p>
      <div class="otr-langs">${LANGS.map(([c, l]) => `<button type="button" data-l="${c}">${l}</button>`).join('')}</div>
      <div class="otr-acts"><button type="button" data-x>ביטול</button></div></div>`;
    const fin = v => { d.remove(); removeEventListener('keydown', key, true); res(v); };
    const key = e => { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); fin(null); } };
    addEventListener('keydown', key, true);
    d.onclick = e => { const b = e.target.closest('[data-l]'); if (b) return fin(b.dataset.l); if (e.target === d || e.target.closest('[data-x]')) fin(null); };
    document.body.append(d); d.querySelector('[data-l]').focus();
  });
}
let running = false;
async function open(kind) {
  kind = kind || liveKind(); const e = ED[kind]; if (!e || running) return;
  const A = window[e.api]; if (!A || !A._test || typeof A._test[e.build] !== 'function') { say('התרגום לא זמין בעורך הזה'); return; }
  const target = await pickLang(e.app); if (!target) return;
  const inp = document.getElementById(e.p + 'Name');
  const base = ((inp && inp.value) || 'קובץ').trim().replace(/\.(docx|xlsx|pptx)$/i, '');
  running = true;
  try {
    wait(true, 'מכין את הקובץ לתרגום…');
    if (kind === 'pptx' && typeof A._test.commitText === 'function') A._test.commitText();
    const r = await A._test[e.build]();
    const src = r && r.bytes ? r.bytes : r;
    const res = await translateBytes(kind, src, target, m => wait(true, m));
    wait(false);
    if (res.lang === target) say('הקובץ כבר כתוב ב' + LANGS.find(x => x[0] === target)[1]);
    const name = `${base} - ${SUF[target]}.${kind}`;
    const fn = window[e.open];
    if (typeof fn !== 'function') throw new Error('לא הצלחתי לפתוח את הקובץ המתורגם');
    await fn(res.bytes, name, false);
    say(`התרגום מוכן ונפתח בלשונית חדשה (${res.count} קטעים). לחצו "שמירה" כדי לשמור אותו`);
  } catch (err) { console.error(err); wait(false); say('התרגום נכשל: ' + (err.message || err)); }
  finally { running = false; wait(false); }
}
// a "תרגום" button in each editor's title bar, next to "שמירה כ-PDF"
function addButton(k, root) {
  const e = ED[k]; if (root.querySelector('#' + e.p + 'Tr')) return;
  const pdf = root.querySelector('#' + e.p + 'Pdf'); if (!pdf) return;
  injectCss();
  const b = document.createElement('button'); b.type = 'button'; b.id = e.p + 'Tr'; b.className = pdf.className.replace(/\bpri\b/, '').trim();
  b.textContent = 'תרגום'; b.title = 'תרגום הקובץ לעברית, אנגלית, רוסית, ערבית, צרפתית או ספרדית, בלי לשנות את העיצוב';
  b.onclick = () => open(k);
  pdf.before(b);
}
new MutationObserver(ms => {
  for (const m of ms) for (const n of m.addedNodes) if (n.nodeType === 1) { const k = Object.keys(ED).find(x => ED[x].id === n.id); if (k) setTimeout(() => addButton(k, n), 0); }
}).observe(document.body, { childList: true });
for (const k of Object.keys(ED)) { const r = document.getElementById(ED[k].id); if (r) addButton(k, r); }
window.OFFICE_TR = { open, translateBytes, _units: units, _decode: decode };
})();
