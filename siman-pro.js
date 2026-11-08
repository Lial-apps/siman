/* =====================================================================
   סימן 6.1 — siman-pro.js
   1. השחרה אוטומטית: מוצא בקובץ מספרי זהות (עם ספרת ביקורת), טלפונים, מיילים, כרטיסי אשראי
      ומספרי IBAN, מראה רשימה לבדיקה, ומשחיר את מה שסומן. ההשחרה היא אותה השחרה של כלי "השחרה":
      הטקסט שמתחת נמחק מהקובץ באמת כששומרים. הכול במכשיר.
   2. דחיסת PDF: כל עמוד נכתב מחדש כתמונה קטנה יותר (כולל הסימונים והחתימות), לשליחה במייל
      ובטפסים מקוונים שמגבילים גודל. הכול במכשיר.
   נטען אחרי הסקריפט הראשי ומשתמש ב: S, drawAnns, snapshot, buildPdf, saveBytes, busy, toast, pdfjs, PDFLib
   ===================================================================== */
(() => {
'use strict';
if (window.SIMAN_PRO) return;
const $ = id => document.getElementById(id);
const esc = t => String(t ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
const h = html => { const t = document.createElement('template'); t.innerHTML = html.trim(); return t.content.firstElementChild; };

document.head.append(h(`<style>
.pro-dlg .sig-in{max-width:560px}
.pro-list{max-height:min(52vh,420px);overflow:auto;border:1px solid var(--line);border-radius:12px;margin:10px 0;background:var(--bar)}
.pro-grp{padding:8px 12px;border-bottom:1px solid var(--line)}
.pro-grp:last-child{border-bottom:0}
.pro-grp>label{display:flex;gap:8px;align-items:center;font-weight:700;font-size:14px;margin-bottom:4px}
.pro-grp ul{list-style:none;margin:0;padding:0 26px 0 0;display:grid;gap:2px}
.pro-grp li label{display:flex;gap:8px;align-items:center;font-size:13.5px;cursor:pointer}
.pro-grp li code{direction:ltr;unicode-bidi:isolate;font-family:inherit;background:var(--desk);border-radius:6px;padding:0 6px}
.pro-grp li small{color:var(--muted)}
.pro-note{font-size:13px;color:var(--muted);margin:4px 0 0}
.pro-opts{display:grid;gap:8px;margin:12px 0}
.pro-opt{display:flex;gap:10px;align-items:flex-start;border:1px solid var(--line);border-radius:12px;padding:10px 12px;cursor:pointer;background:var(--bar)}
.pro-opt:has(input:checked){border-color:var(--ink);box-shadow:0 0 0 1px var(--ink)}
.pro-opt b{display:block}.pro-opt small{color:var(--muted)}
</style>`));

/* ---------------- 1. sensitive information ---------------- */
const KINDS = { id: 'מספרי זהות', phone: 'טלפונים', mail: 'כתובות מייל', card: 'כרטיסי אשראי', iban: 'חשבונות בנק (IBAN)' };
const digits = s => s.replace(/\D/g, '');
// Israeli ID: 9 digits (8 are padded with a 0) and a check digit
function validId(d) {
  if (d.length < 8 || d.length > 9) return false;
  d = d.padStart(9, '0'); if (/^0+$/.test(d) || /^(\d)\1+$/.test(d)) return false;
  let s = 0; for (let i = 0; i < 9; i++) { let x = +d[i] * ((i % 2) + 1); if (x > 9) x -= 9; s += x; }
  return s % 10 === 0;
}
function luhn(d) {
  let s = 0, alt = false;
  for (let i = d.length - 1; i >= 0; i--) { let x = +d[i]; if (alt) { x *= 2; if (x > 9) x -= 9; } s += x; alt = !alt; }
  return s % 10 === 0;
}
// the order matters: a longer number (a card) is taken before the shorter ones inside it
const RULES = [
  ['mail', /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g],
  ['iban', /\bIL\d{2}(?:[ -]?\d{4}){4}[ -]?\d{3}\b/gi],
  ['card', /(?<![\d-])(?:\d[ -]?){12,18}\d(?![\d-])/g, m => { const d = digits(m); return d.length >= 13 && d.length <= 19 && luhn(d) && !/^(\d)\1+$/.test(d); }],
  ['phone', /(?<![\d-])(?:\+?972[- ]?|0)(?:5\d|7\d|[2-489])[- ]?\d{3}[- ]?\d{4}(?![\d-])/g],
  // a PDF that keeps Hebrew backwards may keep the digits backwards too: both ways are checked
  ['id', /(?<![\d-])\d{8,9}(?![\d-])|(?<![\d-])\d{3}-?\d{5}-?\d(?![\d-])/g, m => { const d = digits(m); return validId(d) || validId([...d].reverse().join('')); }],
];
const RTL = /[֐-ࣿ]/;
// where letters j0..j1 sit inside a piece of text, as a share of its width: measured, since letters differ in width
const MC = document.createElement('canvas').getContext('2d');
function span(it, j0, j1) {
  MC.font = `100px ${it.fam}`;
  const W = MC.measureText(it.s).width;
  if (!(W > 0)) return [j0 / it.s.length, (j1 + 1) / it.s.length];
  return [MC.measureText(it.s.slice(0, j0)).width / W, Math.min(1, MC.measureText(it.s.slice(0, j1 + 1)).width / W)];
}
// the words of one page, joined into lines, with each letter knowing which word piece it came from
async function pageHits(P) {
  const tc = await P.page.getTextContent();
  const items = tc.items.filter(it => it.str && it.str.trim() && Math.abs(it.transform[1]) < 0.01 && Math.abs(it.transform[2]) < 0.01)
    .map(it => ({ s: it.str, x: it.transform[4], y: it.transform[5], w: it.width, h: Math.abs(it.transform[3]) || it.height || 10, fam: (tc.styles && tc.styles[it.fontName] && tc.styles[it.fontName].fontFamily) || 'sans-serif' }));
  const lines = [];
  for (const it of items.sort((a, b) => b.y - a.y || a.x - b.x)) {
    const L = lines.find(l => Math.abs(l.y - it.y) < Math.max(l.h, it.h) * 0.45);
    if (L) L.items.push(it); else lines.push({ y: it.y, h: it.h, items: [it] });
  }
  const hits = [];
  for (const L of lines) {
    // left to right on the page: numbers, mails and phones are written left to right also inside Hebrew
    L.items.sort((a, b) => a.x - b.x);
    let str = ''; const map = [];
    L.items.forEach((it, k) => {
      if (k) { const p = L.items[k - 1]; const gap = it.x - (p.x + p.w); if (gap > it.h * 0.25) { str += ' '; map.push(null); } }
      for (let j = 0; j < it.s.length; j++) { str += it.s[j]; map.push([k, j]); }
    });
    const used = [];
    for (const [kind, re, ok] of RULES) {
      re.lastIndex = 0; let m;
      while ((m = re.exec(str))) {
        const a = m.index, b = a + m[0].length;
        if (used.some(([x, y]) => a < y && b > x)) continue;
        if (ok && !ok(m[0])) continue;
        used.push([a, b]);
        // the piece of each word item that the match covers
        const per = new Map();
        for (let c = a; c < b; c++) { const mp = map[c]; if (!mp) continue; const r = per.get(mp[0]) || [mp[1], mp[1]]; r[0] = Math.min(r[0], mp[1]); r[1] = Math.max(r[1], mp[1]); per.set(mp[0], r); }
        const rects = [];
        for (const [k, [j0, j1]] of per) {
          const it = L.items[k];
          let [f0, f1] = span(it, j0, j1);
          if (RTL.test(it.s)) { const t = f0; f0 = 1 - f1; f1 = 1 - t; }   // a Hebrew line runs right to left
          // a little extra on each side: better to cover a hair too much than to leave a digit showing
          const pad = it.h * 0.18;
          rects.push({ x1: it.x + it.w * f0 - pad, x2: it.x + it.w * f1 + pad, y1: it.y - it.h * 0.28, y2: it.y + it.h * 0.95 });
        }
        // a number found backwards is shown the right way round in the list
        const shown = kind === 'id' && !validId(digits(m[0])) ? [...m[0].trim()].reverse().join('') : m[0].trim();
        if (rects.length) hits.push({ kind, text: shown, page: P.i, rects });
      }
    }
  }
  return hits;
}
const mask = (kind, t) => kind === 'mail' ? t.replace(/^(.).*?(@)/, '$1•••$2') : t.replace(/\d(?=[\d\s-]{3})/g, '•');
async function autoRedact() {
  if (!S.doc) { toast('פתחו קובץ PDF קודם'); return; }
  if (S.office) { toast('השחרה אוטומטית עובדת על PDF. שמרו את הקובץ כ-PDF ופתחו אותו שוב'); return; }
  busy(true, 'מחפש מידע רגיש…');
  let hits = [];
  try {
    for (const P of S.pages) { $('busyMsg').textContent = `מחפש מידע רגיש… עמוד ${P.i} מתוך ${S.pages.length}`; hits.push(...await pageHits(P)); }
  } catch (e) { console.error(e); busy(false); toast('החיפוש נכשל: ' + (e.message || e)); return; }
  busy(false);
  // what is already blacked out is not offered again
  const covered = r => S.anns.some(a => a.type === 'redact' && Math.min(a.x1, a.x2) <= r.x1 + 2 && Math.max(a.x1, a.x2) >= r.x2 - 2 && Math.min(a.y1, a.y2) <= r.y1 + 2 && Math.max(a.y1, a.y2) >= r.y2 - 2);
  hits = hits.filter((x, i) => !x.rects.every(r => covered({ x1: r.x1, x2: r.x2, y1: r.y1, y2: r.y2 })));
  if (!hits.length) {
    const any = S.pages.some(P => P.ocr) || (await Promise.all(S.pages.slice(0, 3).map(P => P.page.getTextContent().then(t => t.items.length)))).some(Boolean);
    toast(any ? 'לא נמצאו מספרי זהות, טלפונים, מיילים או כרטיסי אשראי בקובץ' : 'לא נמצא טקסט בקובץ. אם הוא סרוק: "עמודים" ← "זיהוי טקסט (OCR)" קודם');
    return;
  }
  const by = {}; hits.forEach((x, i) => { (by[x.kind] = by[x.kind] || []).push(i); });
  const dlg = h(`<dialog class="sig pro-dlg" aria-labelledby="proRedT"><div class="sig-in">
    <h2 id="proRedT">נמצא מידע רגיש (${hits.length})</h2>
    <p class="pro-note">מה שמסומן יושחר. הטקסט שמתחת יימחק מהקובץ באמת כששומרים. אפשר לבטל אחר כך עם Ctrl+Z.</p>
    <div class="pro-list">${Object.keys(KINDS).filter(k => by[k]).map(k => `<div class="pro-grp"><label><input type="checkbox" data-g="${k}" checked> ${KINDS[k]} (${by[k].length})</label><ul>${by[k].map(i => `<li><label><input type="checkbox" data-i="${i}" data-k="${k}" checked><code>${esc(mask(k, hits[i].text))}</code><small>עמוד ${hits[i].page}</small></label></li>`).join('')}</ul></div>`).join('')}</div>
    <div class="sig-row"><span class="grow"></span><button class="btn" data-a="no">ביטול</button><button class="btn primary" data-a="go">להשחיר את המסומנים</button></div></div></dialog>`);
  document.body.append(dlg);
  dlg.querySelectorAll('[data-g]').forEach(g => g.onchange = () => dlg.querySelectorAll(`[data-k="${g.dataset.g}"]`).forEach(c => { c.checked = g.checked; }));
  dlg.querySelectorAll('[data-i]').forEach(c => c.onchange = () => { const all = [...dlg.querySelectorAll(`[data-k="${c.dataset.k}"]`)]; const g = dlg.querySelector(`[data-g="${c.dataset.k}"]`); g.checked = all.every(x => x.checked); g.indeterminate = !g.checked && all.some(x => x.checked); });
  dlg.addEventListener('close', () => dlg.remove());
  dlg.querySelector('[data-a=no]').onclick = () => dlg.close();
  dlg.querySelector('[data-a=go]').onclick = () => {
    const pick = [...dlg.querySelectorAll('[data-i]:checked')].map(c => hits[+c.dataset.i]);
    dlg.close();
    if (!pick.length) return;
    snapshot();
    const pages = new Set();
    for (const x of pick) for (const r of x.rects) { S.anns.push({ id: S.nextId++, type: 'redact', page: x.page - 1, x1: r.x1, y1: r.y1, x2: r.x2, y2: r.y2 }); pages.add(x.page); }
    for (const p of pages) drawAnns(S.pages[p - 1]);
    const first = S.pages[Math.min(...pages) - 1]; if (first) first.el.scrollIntoView({ block: 'start', behavior: 'smooth' });
    toast(`הושחרו ${pick.length} פריטים. בדקו את הדפים ושמרו כדי למחוק אותם מהקובץ`);
  };
  dlg.showModal();
}

/* ---------------- 2. a smaller PDF ---------------- */
const LEVELS = {
  mail: { l: 'מאוזן', d: 'מתאים למייל ולטפסים באתרים. הטקסט נשאר קריא וחד.', dpi: 150, q: 0.72 },
  small: { l: 'הכי קטן', d: 'לקבצים גדולים במיוחד. מתאים לקריאה על המסך.', dpi: 105, q: 0.55 },
};
const kb = n => n > 1048576 ? (n / 1048576).toFixed(1).replace('.0', '') + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB';
function compressDialog() {
  if (!S.doc) { toast('פתחו קובץ PDF קודם'); return; }
  const size = S.bytes ? S.bytes.length || S.bytes.byteLength : 0;
  const dlg = h(`<dialog class="sig pro-dlg" aria-labelledby="proCmpT"><div class="sig-in">
    <h2 id="proCmpT">דחיסה: PDF קטן יותר</h2>
    <p class="pro-note">${size ? `הגודל עכשיו: <b>${kb(size)}</b>. ` : ''}כל עמוד נשמר מחדש כתמונה קטנה יותר, עם הסימונים והחתימות. נוצר עותק חדש, והקובץ המקורי לא משתנה.</p>
    <div class="pro-opts">${Object.entries(LEVELS).map(([k, v], i) => `<label class="pro-opt"><input type="radio" name="proLvl" value="${k}"${i ? '' : ' checked'}><span><b>${v.l}</b><small>${v.d}</small></span></label>`).join('')}</div>
    <p class="pro-note">שימו לב: אחרי הדחיסה אי אפשר לסמן ולהעתיק טקסט מהקובץ (אפשר להחזיר עם "זיהוי טקסט (OCR)").</p>
    <div class="sig-row"><span class="grow"></span><button class="btn" data-a="no">ביטול</button><button class="btn primary" data-a="go">דחיסה ושמירה</button></div></div></dialog>`);
  document.body.append(dlg);
  dlg.addEventListener('close', () => dlg.remove());
  dlg.querySelector('[data-a=no]').onclick = () => dlg.close();
  dlg.querySelector('[data-a=go]').onclick = () => { const k = dlg.querySelector('input[name=proLvl]:checked').value; dlg.close(); compress(LEVELS[k]); };
  dlg.showModal();
}
async function compress(lv) {
  busy(true, 'מכין את הקובץ…');
  let doc = null;
  try {
    const built = await buildPdf();
    const before = built.length || built.byteLength;
    doc = await pdfjs.getDocument({ data: built.slice ? built.slice() : built, isEvalSupported: false }).promise;
    const out = await PDFLib.PDFDocument.create();
    for (let i = 1; i <= doc.numPages; i++) {
      $('busyMsg').textContent = `דוחס עמוד ${i} מתוך ${doc.numPages}…`;
      const page = await doc.getPage(i), v1 = page.getViewport({ scale: 1 });
      // no side over 3000 pixels: phones run out of memory beyond that
      const sc = Math.min(lv.dpi / 72, 3000 / Math.max(v1.width, v1.height)), vp = page.getViewport({ scale: sc });
      const c = document.createElement('canvas'); c.width = Math.ceil(vp.width); c.height = Math.ceil(vp.height);
      const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
      await page.render({ canvasContext: g, viewport: vp }).promise;
      const blob = await new Promise(r => c.toBlob(r, 'image/jpeg', lv.q)); c.width = c.height = 0;
      const img = await out.embedJpg(new Uint8Array(await blob.arrayBuffer()));
      const pg = out.addPage([v1.width, v1.height]); pg.drawImage(img, { x: 0, y: 0, width: v1.width, height: v1.height });
      page.cleanup();
    }
    out.setTitle(S.name.replace(/\.pdf$/i, '')); out.setProducer('סימן'); out.setCreator('סימן');
    const bytes = await out.save({ useObjectStreams: true });
    busy(false);
    if (bytes.length >= before * 0.95) { toast(`הקובץ כבר קטן (${kb(before)}). דחיסה לא תקטין אותו עוד`); return; }
    const name = S.name.replace(/\.pdf$/i, '') + ' - מכווץ.pdf';
    const saved = await saveBytes(name, bytes, 'pdf');
    // the size the user knows is the file as it was opened (when nothing was added since)
    const orig = (S.bytes && (S.bytes.length || S.bytes.byteLength)) || 0, from = S.anns.length || !orig ? before : orig;
    if (saved) toast(`נשמר: ${name} · ${kb(from)} ← ${kb(bytes.length)} (קטן ב-${Math.round((1 - bytes.length / from) * 100)}%)`);
  } catch (e) { console.error(e); busy(false); toast('הדחיסה נכשלה: ' + (e.message || e)); }
  finally { busy(false); try { doc && doc.destroy(); } catch (_) {} }
}

window.SIMAN_PRO = { autoRedact, compressDialog, _test: { pageHits, validId, luhn, compress, LEVELS } };
})();
