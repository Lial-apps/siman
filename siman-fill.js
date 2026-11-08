/* =====================================================================
   סימן 5.9.6 — siman-fill.js
   מילוי וחתימה: כשנפתח טופס, סימן מוצא בעצמו את המקומות למילוי
   · קווים ריקים (____ / .......), גם בטופס סרוק — לפי הציור של הדף
   · תאים ריקים בטבלה (מסגרת עם קווים מכל הצדדים ובלי טקסט בפנים)
   · תיבות סימון ☐ (מצוירות או כתו)
   · "חתימה:", "תאריך:", "שם:" — לפי הכיתוב שליד הקו
   לחיצה על מקום מסומן: תיבת סימון ← V (לחיצה נוספת מורידה). חתימה ← החתימה השמורה נכנסת בדיוק על הקו.
   טקסט / תאריך ← נפתחת חלונית מילוי שגדלה ככל שכותבים, והטקסט על הדף גדל איתה (קודם לרוחב, אחר כך שורות).
   "הבא" עובר למקום הבא. הסימונים הצהובים לא נשמרים בקובץ ולא מודפסים.
   משתמש ב: S, toPdf, toCss, snapshot, drawAnns, addImage, startPlace, openSig, loadSigs, setDirty, cssFont, toast, pdfjs
   ===================================================================== */
(() => {
'use strict';
if (window.SIMAN_FILL || typeof S === 'undefined') return;

const LANG = () => (document.documentElement.lang || 'he').slice(0, 2);
const T = {
  he: { found: n => `מצאתי ${n} מקומות למילוי. לוחצים על מקום מסומן כדי למלא`, next: 'הבא', done: 'סיום', left: n => `נשארו ${n}`, all: 'הכול מולא ✓', hide: 'הסתרת הסימונים', show: n => `מקומות למילוי (${n})`,
    k: { text: 'מילוי', name: 'שם', date: 'תאריך', sig: 'חתימה', id: 'מספר זהות', phone: 'טלפון', cell: 'תא בטבלה', check: 'סימון' },
    today: 'היום', asV: 'סימון V', asSig: 'חתימה', asText: 'טקסט', del: 'מחיקה', ph: 'כותבים כאן…', otherSig: 'חתימה אחרת', tip: 'Enter: הבא · Shift+Enter: שורה חדשה' },
  en: { found: n => `Found ${n} places to fill. Tap a marked place to fill it`, next: 'Next', done: 'Done', left: n => `${n} left`, all: 'All filled ✓', hide: 'Hide marks', show: n => `Places to fill (${n})`,
    k: { text: 'Fill in', name: 'Name', date: 'Date', sig: 'Signature', id: 'ID number', phone: 'Phone', cell: 'Table cell', check: 'Check' },
    today: 'Today', asV: 'Check ✓', asSig: 'Signature', asText: 'Text', del: 'Delete', ph: 'Type here…', otherSig: 'Another signature', tip: 'Enter: next · Shift+Enter: new line' },
  ar: { found: n => `وجدت ${n} أماكن للتعبئة. اضغطوا على مكان مُعلَّم لتعبئته`, next: 'التالي', done: 'إنهاء', left: n => `تبقّى ${n}`, all: 'تمت التعبئة ✓', hide: 'إخفاء العلامات', show: n => `أماكن للتعبئة (${n})`,
    k: { text: 'تعبئة', name: 'الاسم', date: 'التاريخ', sig: 'التوقيع', id: 'رقم الهوية', phone: 'الهاتف', cell: 'خلية', check: 'علامة' },
    today: 'اليوم', asV: 'علامة ✓', asSig: 'توقيع', asText: 'نص', del: 'حذف', ph: 'اكتبوا هنا…', otherSig: 'توقيع آخر', tip: 'Enter: التالي · Shift+Enter: سطر جديد' },
  ru: { found: n => `Найдено мест для заполнения: ${n}. Нажмите на отмеченное место`, next: 'Далее', done: 'Готово', left: n => `Осталось ${n}`, all: 'Всё заполнено ✓', hide: 'Скрыть метки', show: n => `Места для заполнения (${n})`,
    k: { text: 'Заполнить', name: 'Имя', date: 'Дата', sig: 'Подпись', id: 'Номер ID', phone: 'Телефон', cell: 'Ячейка', check: 'Отметка' },
    today: 'Сегодня', asV: 'Галочка ✓', asSig: 'Подпись', asText: 'Текст', del: 'Удалить', ph: 'Пишите здесь…', otherSig: 'Другая подпись', tip: 'Enter: далее · Shift+Enter: новая строка' }
};
const L = () => T[LANG()] || T.he;
// 5.9.6 (second round): what the window says about suggestions, checks and the end of the form
const TX = {
  he: { prev: 'הקודם', save: 'הכול מולא ✓ · שמירה', badId: 'מספר הזהות לא תקין (ספרת ביקורת). כדאי לבדוק', badPhone: 'מספר הטלפון נראה קצר או ארוך מדי', badMail: 'כתובת המייל לא נראית תקינה', addr: 'כתובת', email: 'מייל', tipK: 'Enter / Tab: הבא · Shift+Tab: הקודם · Shift+Enter: שורה חדשה' },
  en: { prev: 'Back', save: 'All filled ✓ · Save', badId: 'This ID number does not check out. Worth a look', badPhone: 'The phone number looks too short or too long', badMail: 'The e-mail address does not look right', addr: 'Address', email: 'E-mail', tipK: 'Enter / Tab: next · Shift+Tab: back · Shift+Enter: new line' },
  ar: { prev: 'السابق', save: 'تمت التعبئة ✓ · حفظ', badId: 'رقم الهوية غير صحيح. يستحسن التحقق', badPhone: 'رقم الهاتف يبدو قصيرًا أو طويلًا جدًا', badMail: 'عنوان البريد لا يبدو صحيحًا', addr: 'العنوان', email: 'البريد', tipK: 'Enter / Tab: التالي · Shift+Tab: السابق · Shift+Enter: سطر جديد' },
  ru: { prev: 'Назад', save: 'Всё заполнено ✓ · Сохранить', badId: 'Номер ID не проходит проверку. Стоит проверить', badPhone: 'Номер телефона слишком короткий или длинный', badMail: 'Адрес почты выглядит неверно', addr: 'Адрес', email: 'Почта', tipK: 'Enter / Tab: далее · Shift+Tab: назад · Shift+Enter: новая строка' }
};
const LX = () => TX[LANG()] || TX.he;
const kindName = t => { const k = t.kind === 'cell' ? (t.sub && t.sub !== 'text' ? t.sub : 'cell') : t.kind; return L().k[k] || LX()[k] || L().k.text; };
// what was typed before, kept only on this device: the next form suggests it (name, ID, phone, address, e-mail)
const MEM_KEY = 'siman.fill.mem', MEM_KINDS = ['name', 'id', 'phone', 'addr', 'email'];
const memGet = () => { try { return JSON.parse(localStorage.getItem(MEM_KEY) || '{}') || {}; } catch (_) { return {}; } };
function memPut(kind, v) {
  v = String(v || '').trim(); if (!v || v.length > 80 || !MEM_KINDS.includes(kind)) return;
  const m = memGet(); m[kind] = [v, ...(m[kind] || []).filter(x => x !== v)].slice(0, 3);
  try { localStorage.setItem(MEM_KEY, JSON.stringify(m)); } catch (_) {}
}
// Israeli ID: 9 digits with a check digit
const idOk = v => { const d = String(v).replace(/\D/g, ''); if (!d || d.length > 9) return false; const s = d.padStart(9, '0'); let t = 0; for (let i = 0; i < 9; i++) { let x = +s[i] * (i % 2 + 1); if (x > 9) x -= 9; t += x; } return t % 10 === 0; };
function warnFor(kind, v) {
  v = String(v || '').trim(); if (!v) return '';
  if (kind === 'id' && /^\d[\d\s-]*$/.test(v) && v.replace(/\D/g, '').length >= 5 && !idOk(v)) return LX().badId;
  if (kind === 'phone' && /\d/.test(v)) { const n = v.replace(/\D/g, '').length; if (n < 9 || n > 13) return LX().badPhone; }
  if (kind === 'email' && v.length > 3 && !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(v)) return LX().badMail;
  return '';
}
const RTL_CH = /[֐-ࣿיִ-﷿ﹰ-﻿]/, LTR_CH = /[A-Za-zÀ-ɏЀ-ӿ]/;
const sleep = ms => new Promise(r => setTimeout(r, ms));
const clamp = (v, a, b) => Math.max(a, Math.min(b, v));

/* ---------------------------------------------------------------- עיצוב */
const css = document.createElement('style');
css.id = 'simanFillCss';
css.textContent = `
.fzl{position:absolute;inset:0;pointer-events:none;z-index:3}
.fz{position:absolute;pointer-events:auto;cursor:pointer;border-radius:3px;background:rgba(255,216,77,.30);outline:1.5px dashed #C99A00;outline-offset:0;transition:background .15s}
.fz::after{content:"";position:absolute;inset:-8px -4px}
.fz:hover,.fz:focus-visible{background:rgba(255,216,77,.55);outline:2px solid #6B2BD9}
.fz.check{border-radius:2px}
.fz.sig{background:rgba(107,43,217,.12);outline-color:#6B2BD9}
.fz.cur{background:rgba(107,43,217,.18);outline:2px solid #6B2BD9}
.fz.done{background:transparent;outline-color:transparent}
.fz.done:not(.check){pointer-events:none}
.fz.done:hover{outline:1px dashed rgba(107,43,217,.5)}
.fz .fzk{position:absolute;bottom:100%;inset-inline-end:0;margin-bottom:1px;font:600 10px/1.2 "Varela Round",Rubik,Arial,sans-serif;color:#4A1FA0;background:#FFF5C7;border-radius:4px 4px 0 0;padding:1px 5px;white-space:nowrap;opacity:0;pointer-events:none;direction:rtl}
.fz:hover .fzk,.fz.cur .fzk{opacity:1}
body.fz-off .fzl{display:none}
body[data-tool="pen"] .fzl,body[data-tool="highlight"] .fzl,body[data-tool="eraser"] .fzl,body[data-tool="redact"] .fzl,body[data-tool="snip"] .fzl,body[data-tool="edit"] .fzl,body[data-tool="place"] .fzl{display:none}
@media print{.fzl,.fzp,.fzbar{display:none!important}}
.fzbar{position:fixed;z-index:60;inset-inline-start:50%;transform:translateX(50%);bottom:calc(18px + env(safe-area-inset-bottom,0px));display:flex;gap:6px;align-items:center;background:#fff;color:#2B1660;border:2px solid #2B1660;border-radius:999px;box-shadow:3px 3px 0 #2B1660;padding:4px 6px 4px 6px;font:600 14px "Varela Round",Rubik,Arial,sans-serif;direction:rtl}
[dir=ltr] .fzbar{transform:translateX(-50%)}
.fzbar[hidden]{display:none}
.fzbar button{font:inherit;border:0;background:none;color:inherit;cursor:pointer;min-height:38px;padding:0 12px;border-radius:999px;white-space:nowrap}
@media (max-width:720px){.fzbar{bottom:calc(84px + env(safe-area-inset-bottom,0px));font-size:13.5px}}
.fzbar .go{background:#FFD84D;border:1.5px solid #2B1660}
.fzbar .x{width:34px;padding:0;font-size:18px;opacity:.7}
.fzp{position:fixed;z-index:61;inset-inline:0;bottom:0;margin:0 auto;max-width:620px;background:#fff;color:#2B1660;border:2px solid #2B1660;border-bottom:0;border-radius:18px 18px 0 0;box-shadow:0 -6px 30px rgba(43,22,96,.25);padding:10px 12px calc(10px + env(safe-area-inset-bottom,0px));font-family:"Varela Round",Rubik,Arial,sans-serif;direction:rtl}
.fzp[hidden]{display:none}
.fzp{z-index:2147483000}
body.fz-filling .a11y-fab,body.fz-filling .sb-back{visibility:hidden}
.fzp .hd{display:flex;align-items:center;gap:8px;margin-bottom:6px;font-size:14px}
.fzp .hd b{font-size:15px}
.fzp .hd .cnt{margin-inline-start:auto;color:#6B5A92;font-size:12.5px}
.fzp .lens{display:flex;justify-content:center;overflow:hidden;margin:0 0 8px;border:1.5px solid #CFC4EA;border-radius:12px;background:#fff;padding:4px}
.fzp .lens[hidden]{display:none}
.fzp .warn{margin-top:6px;font-size:13px;color:#9A3412;background:#FFF4E5;border-radius:8px;padding:5px 9px}
.fzp .warn[hidden],.fzp .mem[hidden]{display:none}
.fzp .mem{display:flex;gap:6px;margin-top:8px;overflow-x:auto;scrollbar-width:none}
.fzp .mem button{flex:none;max-width:70%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font:inherit;font-size:13.5px;min-height:36px;padding:0 12px;border-radius:999px;border:1.5px dashed #6B2BD9;background:#F4EEFF;color:#2B1660;cursor:pointer;unicode-bidi:plaintext}
.fzp .mem button::before{content:"↺ ";opacity:.6}
.fzbar .go.fin{background:#2ED3A0}
[data-theme=dark] .fzp .mem button{background:#241D48;color:#EDE9F7;border-color:#C9B6FA}
[data-theme=dark] .fzp .warn{background:#3A2414;color:#FFD3A8}
.fzp .lens canvas{display:block;max-width:100%;height:auto!important}
.fzp textarea{display:block;width:100%;box-sizing:border-box;min-height:46px;max-height:40vh;resize:none;overflow-y:auto;border:2px solid #6B2BD9;border-radius:12px;padding:10px 12px;font:16px/1.45 "AlefEmbed",Rubik,Arial,sans-serif;color:#1C2333;background:#FBFAFF;unicode-bidi:plaintext;text-align:start;transition:height .08s}
.fzp textarea:focus{outline:3px solid rgba(107,43,217,.25);outline-offset:0}
.fzp .row{display:flex;flex-wrap:wrap;gap:6px;margin-top:8px;align-items:center}
.fzp .row button{font:inherit;font-size:13.5px;min-height:40px;padding:0 12px;border-radius:12px;border:1.5px solid #CFC4EA;background:#fff;color:#2B1660;cursor:pointer}
.fzp .row button.main{background:#6B2BD9;border-color:#6B2BD9;color:#fff;font-weight:700}
.fzp .row button.sun{background:#FFD84D;border-color:#2B1660}
.fzp .row .sp{flex:1}
.fzp .tip{font-size:11.5px;color:#7D6FA3;margin-top:6px}
@media (pointer:coarse){.fzp .tip{display:none}}
[data-theme=dark] .fzp,[data-theme=dark] .fzbar{background:#1B1638;color:#EDE9F7;border-color:#C9B6FA}
[data-theme=dark] .fzp textarea{background:#120E28;color:#F4F1FF;border-color:#C9B6FA}
[data-theme=dark] .fzp .row button{background:#241D48;color:#EDE9F7;border-color:#4B3F86}
[data-theme=dark] .fzp .row button.main{background:#8B5CF6;border-color:#8B5CF6}
[data-theme=dark] .fzp .row button.sun,[data-theme=dark] .fzbar .go{background:#F5CD3D;color:#2B1660}
[data-theme=dark] .fzp .hd .cnt,[data-theme=dark] .fzp .tip{color:#B9AEE0}
[data-theme=dark] .fzp .lens{border-color:#4B3F86}
`;
document.head.append(css);

/* ---------------------------------------------------------------- זיהוי */
// what a label next to a blank asks for
// Hebrew in many PDFs is stored backwards ("ךיראת"): the label is read both ways
function kindOf(t) {
  const a = kindOf1(t); if (a !== "text" || !RTL_CH.test(String(t || ""))) return a;
  const b = kindOf1([...String(t)].reverse().join("")); return b || a;
}
function kindOf1(t) {
  t = String(t || '').replace(/[|_=~<>"'״׳`.,;()\[\]{}:：*]/g, ' ').replace(/\s+/g, ' ').trim();
  if (!t) return null;
  if (/חתימ|ח\s?ת\s?י\s?מ|signature|sign here|توقيع|подпис/i.test(t)) return 'sig';
  if (/תאריך|date|تاريخ|дата/i.test(t)) return 'date';
  if (/(^|\s)ת\s?ז(\s|$)|תעודת זהות|מס(פר)?\s*זהות|\bid\b|identity|هوية|паспорт/i.test(t)) return 'id';
  if (/מייל|דוא\s?ל|e-?mail|بريد|почт/i.test(t)) return 'email';
  if (/טלפון|נייד|phone|mobile|هاتف|телефон/i.test(t)) return 'phone';
  if (/כתובת|רחוב|address|عنوان|адрес/i.test(t)) return 'addr';
  if (/(^|\s)שם(\s|$)|משפחה|name|اسم|имя|фамил/i.test(t)) return 'name';
  return 'text';
}
const BLANK = /_{3,}|\.{6,}|…{3,}|-{6,}|‗{3,}/g;
const BOXCH = /[☐□▢◻❏❑⬜]/g;

async function textOf(page, vp) {
  const tc = await page.getTextContent();
  const out = [];
  for (const it of tc.items) {
    const s = String(it.str || ''); if (!s.trim()) continue;
    const tx = it.transform, fh = Math.hypot(tx[2], tx[3]) || 10;
    const p1 = vp.convertToViewportPoint(tx[4], tx[5]), p2 = vp.convertToViewportPoint(tx[4] + (it.width || fh * s.length * .5), tx[5] + fh);
    out.push({ s, fh, x0: Math.min(p1[0], p2[0]), x1: Math.max(p1[0], p2[0]), y0: Math.min(p1[1], p2[1]), y1: Math.max(p1[1], p2[1]), rtl: it.dir === 'rtl' || RTL_CH.test(s) });
  }
  return out;
}
// the words next to a blank on the same line (Hebrew: to its right, otherwise to its left), and their size
function labelFor(items, b, pageRtl) {
  const midY = (b.y0 + b.y1) / 2, cand = items.filter(t => !t.blank && t.y0 < b.y1 + 2 && t.y1 > b.y0 - 6 && Math.abs((t.y0 + t.y1) / 2 - midY) < Math.max(8, t.fh));
  const right = cand.filter(t => t.x0 >= b.x1 - 3).sort((a, c) => a.x0 - c.x0)[0];
  const left = cand.filter(t => t.x1 <= b.x0 + 3).sort((a, c) => c.x1 - a.x1)[0];
  const near = t => t && Math.abs((t === right ? t.x0 - b.x1 : b.x0 - t.x1)) < 160;
  const first = pageRtl ? (near(right) ? right : near(left) ? left : null) : (near(left) ? left : near(right) ? right : null);
  if (first) return { text: first.s, fh: first.fh };
  // the label is written above the line ("שם מלא של התלמיד/ה:" and the line under it)
  const above = items.filter(t => !t.blank && t.y1 <= b.y0 + 4 && b.y0 - t.y1 < Math.max(14, t.fh * 1.6) && t.x0 < b.x1 && t.x1 > b.x0)
    .sort((a, c) => c.y1 - a.y1)[0];
  return above ? { text: above.s, fh: above.fh, above: true } : null;
}

// the page drawn small: ink as 1/0
async function inkOf(page) {
  const v1 = page.getViewport({ scale: 1 }), k = Math.min(2, 1500 / Math.max(v1.width, v1.height));
  const vp = page.getViewport({ scale: k });
  const c = document.createElement('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
  const g = c.getContext('2d', { willReadFrequently: true }); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
  await page.render({ canvasContext: g, viewport: vp, annotationMode: 0 }).promise;
  const W = c.width, H = c.height, d = g.getImageData(0, 0, W, H).data, dark = new Uint8Array(W * H);
  for (let i = 0, j = 0; i < dark.length; i++, j += 4) dark[i] = (d[j] * .3 + d[j + 1] * .59 + d[j + 2] * .11) < 165 ? 1 : 0;
  const out = { W, H, k, dark, g };
  c.width = c.height = 0; // the picture itself is not needed any more
  return out;
}
const inkIn = (I, x0, y0, x1, y1) => {
  x0 = Math.max(0, Math.round(x0)); y0 = Math.max(0, Math.round(y0)); x1 = Math.min(I.W - 1, Math.round(x1)); y1 = Math.min(I.H - 1, Math.round(y1));
  if (x1 < x0 || y1 < y0) return 0; let n = 0;
  for (let y = y0; y <= y1; y++) { const r = y * I.W; for (let x = x0; x <= x1; x++) n += I.dark[r + x]; }
  return n / ((x1 - x0 + 1) * (y1 - y0 + 1));
};
// horizontal strokes (blank lines and the rules of tables), each split where a vertical rule crosses it
function strokes(I) {
  const { W, H, dark } = I, minL = Math.round(W * .035), gap = Math.max(2, Math.round(W * .003)), runs = [];
  for (let y = 0; y < H; y++) { const row = y * W; let x = 0;
    while (x < W) { if (!dark[row + x]) { x++; continue; }
      let e = x, x1 = x, miss = 0; while (e < W && miss <= gap) { if (dark[row + e]) { x1 = e; miss = 0; } else miss++; e++; }
      if (x1 - x >= minL) { let f = 0; for (let k = x; k <= x1; k++) f += dark[row + k]; if (f / (x1 - x + 1) > .72) runs.push({ x0: x, x1, y }); }
      x = e; } }
  const lines = [];
  for (const r of runs) { const l = lines.find(l => r.y - l.y1 <= 1 && r.x0 < l.x1 && r.x1 > l.x0);
    if (l) { l.y1 = r.y; l.x0 = Math.min(l.x0, r.x0); l.x1 = Math.max(l.x1, r.x1); } else lines.push({ x0: r.x0, x1: r.x1, y0: r.y, y1: r.y }); }
  // thick = a filled bar or a picture, not a line
  return lines.filter(l => l.y1 - l.y0 <= Math.max(4, H * .004));
}
// how far a vertical rule goes up from (x, y)
function vUp(I, x, y) {
  let best = 0;
  for (let dx = -1; dx <= 1; dx++) { const xx = x + dx; if (xx < 0 || xx >= I.W) continue;
    let h = 0, miss = 0, yy = y - 1; while (yy >= 0 && miss <= 1) { if (I.dark[yy * I.W + xx]) { h = y - yy; miss = 0; } else miss++; yy--; }
    best = Math.max(best, h); }
  return best;
}
// how far a vertical rule goes down from (x, y): the top edge of a table or of a frame
function vDown(I, x, y) {
  let best = 0;
  for (let dx = -1; dx <= 1; dx++) { const xx = x + dx; if (xx < 0 || xx >= I.W) continue;
    let h = 0, miss = 0, yy = y + 1; while (yy < I.H && miss <= 1 && yy - y < I.H * .05) { if (I.dark[yy * I.W + xx]) { h = yy - y; miss = 0; } else miss++; yy++; }
    best = Math.max(best, h); }
  return best;
}
// the rule above a table cell: the first row, going up, that is mostly ink across the cell
function cellTop(I, x0, x1, y, hMax) {
  for (let d = 3; d <= hMax; d++) if (inkIn(I, x0 + 3, y - d, x1 - 3, y - d) > .6) return y - d;
  return y - hMax;
}
function pieces(I, ln) {
  const y = ln.y0, minV = Math.round(I.H * .012), cuts = [];
  for (let x = ln.x0; x <= ln.x1; x++) { const h = vUp(I, x, y); if (h >= minV) { if (cuts.length && x - cuts[cuts.length - 1].x <= 3) { const c = cuts[cuts.length - 1]; if (h > c.h) { c.h = h; } continue; } cuts.push({ x, h }); } }
  if (!cuts.length) return [{ x0: ln.x0, x1: ln.x1, cell: false }];
  const pts = [{ x: ln.x0, h: 0 }, ...cuts, { x: ln.x1, h: 0 }], out = [];
  for (let i = 0; i < pts.length - 1; i++) { const a = pts[i], b = pts[i + 1]; if (b.x - a.x < 6) continue;
    out.push({ x0: a.x, x1: b.x, cell: a.h > 0 && b.h > 0, h: Math.min(a.h || 0, b.h || 0) || Math.max(a.h, b.h) }); }
  return out;
}

async function detectPage(P, widgets) {
  const page = P.page, vp1 = page.getViewport({ scale: 1 }), W1 = vp1.width, H1 = vp1.height;
  let items = []; try { items = await textOf(page, vp1); } catch (_) {}
  const rtlN = items.filter(t => t.rtl).reduce((n, t) => n + t.s.length, 0), allN = items.reduce((n, t) => n + t.s.length, 0) || 1;
  const pageRtl = items.length ? rtlN / allN > .25 : /^(he|ar)/.test(LANG());
  const out = [];
  const push = t => {
    t.w = t.x1 - t.x0; t.h = t.y1 - t.y0; if (t.w < 6 || t.h < 4) return;
    if (widgets.some(r => t.x0 < r[2] && t.x1 > r[0] && t.y0 < r[3] && t.y1 > r[1])) return;  // a real form field is already there
    const dup = out.find(o => o.x0 < t.x1 - 3 && o.x1 > t.x0 + 3 && Math.abs(o.y1 - t.y1) < Math.max(6, t.h * .6));
    if (dup) return; out.push(t);
  };
  // 1. blanks written as characters: "שם: ________", "........", "☐ אני מסכים/ה"
  for (const it of items) {
    const s = it.s, n = s.length, span = it.x1 - it.x0;
    for (const m of [...s.matchAll(new RegExp(BLANK.source, "g"))]) {
      const a = m.index / n, b = (m.index + m[0].length) / n;
      const x0 = it.rtl ? it.x1 - span * b : it.x0 + span * a, x1 = it.rtl ? it.x1 - span * a : it.x0 + span * b;
      const rest = (s.slice(0, m.index) + ' ' + s.slice(m.index + m[0].length)).replace(BLANK, ' ').trim();
      const lb = rest.length > 1 ? { text: rest, fh: it.fh } : null;
      const bl = { x0, x1, y0: it.y1 - Math.max(14, it.fh * 1.5), y1: it.y1 + it.fh * .15 };
      const lab = lb || labelFor(items, bl, pageRtl);
      push({ ...bl, kind: kindOf(lab && lab.text) || 'text', fh: lab ? lab.fh : it.fh, label: lab ? lab.text : '', rtl: pageRtl });
      if (m[0].length === n) it.blank = true;
    }
    // check boxes written as a character
    for (const m of [...s.matchAll(new RegExp(BOXCH.source, "g"))]) { const a = m.index / n, b = (m.index + 1) / n;
      const x0 = it.rtl ? it.x1 - span * b : it.x0 + span * a, x1 = it.rtl ? it.x1 - span * a : it.x0 + span * b, sz = Math.max(x1 - x0, it.fh * .8);
      push({ x0: (x0 + x1) / 2 - sz / 2, x1: (x0 + x1) / 2 + sz / 2, y0: it.y1 - it.fh * .85, y1: it.y1 - it.fh * .85 + sz, kind: 'check', label: '', rtl: pageRtl }); }
  }
  // 2. lines, table cells and check boxes drawn on the page (and every scanned form)
  let I = null; try { I = await inkOf(page); } catch (e) { console.warn('fill: page picture', e); }
  if (I) {
    const k = I.k, toV = v => v / k;
    const textBox = (x0, y0, x1, y1) => items.some(t => !t.blank && t.x0 < x1 - 2 && t.x1 > x0 + 2 && t.y0 < y1 - 1 && t.y1 > y0 + 1);
    for (const ln of strokes(I)) {
      for (const pc of pieces(I, ln)) {
        const len = pc.x1 - pc.x0; if (len < I.W * .035) continue;
        if (pc.cell) {
          // a table cell: the rule above it is where the vertical rules start
          const top = cellTop(I, pc.x0, pc.x1, ln.y0, pc.h), hh = ln.y0 - top; if (hh < I.H * .014 || hh > I.H * .14) continue;
          const ix0 = pc.x0 + 3, ix1 = pc.x1 - 3, iy0 = ln.y0 - hh + 3, iy1 = ln.y0 - 3;
          if (inkIn(I, ix0, iy0, ix1, iy1) > .006) continue;   // written in already
          const X0 = toV(pc.x0), X1 = toV(pc.x1), Y0 = toV(ln.y0 - hh), Y1 = toV(ln.y0);
          if (textBox(X0, Y0, X1, Y1)) continue;
          if (len < I.W * .06 && Math.abs(len - hh) < hh * .3) { push({ x0: X0 + 1, x1: X1 - 1, y0: Y0 + 1, y1: Y1 - 1, kind: 'check', label: '', rtl: pageRtl }); continue; }
          const lab = labelFor(items, { x0: X0, x1: X1, y0: Y0, y1: Y1 }, pageRtl);
          push({ x0: X0 + 1, x1: X1 - 1, y0: Y0 + 1, y1: Y1 - 1, kind: 'cell', sub: lab ? kindOf(lab.text) : null, label: lab ? lab.text : '', rtl: pageRtl, fh: lab ? lab.fh : 0 });
        } else {
          if (len > I.W * .8) continue;                         // a rule across the page
          const dn = Math.round(I.H * .012);
          if (vDown(I, pc.x0, ln.y1) >= dn || vDown(I, pc.x1, ln.y1) >= dn) continue; // the top edge of a table or a frame
          // room above the line, up to the ink above it
          let up = 0; const maxUp = Math.round(I.H * .03);
          for (let d = 3; d <= maxUp; d++) { if (inkIn(I, pc.x0 + 2, ln.y0 - d, pc.x1 - 2, ln.y0 - d) > .02) break; up = d; }
          if (up < I.H * .009) continue;                        // letters sit on it: underlined text
          if (inkIn(I, pc.x0, ln.y1 + 2, pc.x1, ln.y1 + Math.round(I.H * .006)) > .25) continue; // part of a block of text
          const X0 = toV(pc.x0), X1 = toV(pc.x1), Y1 = toV(ln.y0), Y0 = toV(ln.y0 - up);
          if (textBox(X0, Y0 + 2, X1, Y1 - 1)) continue;
          const lab = labelFor(items, { x0: X0, x1: X1, y0: Y0, y1: Y1 }, pageRtl);
          push({ x0: X0, x1: X1, y0: Y0, y1: Y1 + 1, kind: kindOf(lab && lab.text) || 'text', label: lab ? lab.text : '', rtl: pageRtl, fh: lab ? lab.fh : 0 });
        }
      }
    }
    // drawn check boxes (the detector of send.js, when it is loaded)
    if (window.SIMAN_SEND && SIMAN_SEND.findBoxes) {
      try { const bw = { getImageData: () => { const d = new Uint8ClampedArray(I.W * I.H * 4); for (let i = 0, j = 0; i < I.dark.length; i++, j += 4) { const v = I.dark[i] ? 0 : 255; d[j] = d[j + 1] = d[j + 2] = v; d[j + 3] = 255; } return { data: d }; } };
        for (const b of SIMAN_SEND.findBoxes(bw, I.W, I.H)) if (!textBox(toV(b.x), toV(b.y), toV(b.x + b.w), toV(b.y + b.h))) push({ x0: toV(b.x), x1: toV(b.x + b.w), y0: toV(b.y), y1: toV(b.y + b.h), kind: 'check', label: '', rtl: pageRtl }); } catch (e) { console.warn('fill: boxes', e); }
    }
    I.dark = null; I.g = null;
  }
  // 3. a label with nothing after it ("חתימה:" at the end of a line): the space beside it
  for (const it of items) {
    const s = it.s.trim(); if (s.length > 30 || !/[:：]\s*$/.test(s)) continue;
    const kd = kindOf(s); if (!['sig', 'date', 'name', 'id', 'phone', 'email'].includes(kd)) continue;
    const wd = W1 * (kd === 'sig' ? .24 : .2);
    const x0 = it.rtl ? Math.max(8, it.x0 - 6 - wd) : it.x1 + 6, x1 = it.rtl ? it.x0 - 6 : Math.min(W1 - 8, it.x1 + 6 + wd);
    if (items.some(t => t !== it && t.x0 < x1 && t.x1 > x0 && t.y0 < it.y1 && t.y1 > it.y0)) continue;  // something is written there
    // the label of a line under it (label above, line below): that line is the place, not the space beside the label
    if (out.some(o => o.kind !== 'check' && o.y1 > it.y1 && o.y0 - it.y1 < Math.max(16, it.fh * 2.5) && o.x0 < it.x1 + 4 && o.x1 > it.x0 - 4)) continue;
    push({ x0, x1, y0: it.y0 - 2, y1: it.y1 + 2, kind: kd, label: s, rtl: it.rtl, fh: it.fh });
  }
  out.forEach(t => { t.page = P.i - 1; });
  // how far a box may grow on its line: up to the next words or the next place on the same line (then it wraps)
  out.forEach(t => {
    const y0 = t.kind === 'cell' ? t.y0 : t.y1 - 16, y1 = t.y1;
    const obst = [...items.filter(it => !it.blank).map(it => [it.x0, it.x1, it.y0, it.y1]), ...out.filter(o => o !== t).map(o => [o.x0, o.x1, o.y0, o.y1])]
      .filter(b => b[2] < y1 - 1 && b[3] > y0 + 1);
    t.limL = obst.filter(b => b[1] <= t.x0 + 2).reduce((m, b) => Math.max(m, b[1] + 4), 14);
    t.limR = obst.filter(b => b[0] >= t.x1 - 2).reduce((m, b) => Math.min(m, b[0] - 4), W1 - 14);
    if (t.kind === 'cell') { t.limL = t.x0; t.limR = t.x1; }   // a table cell grows down, not into the next cell
  });
  // reading order: top to bottom, and on each line from the start of the line
  return out.filter(t => t.x0 >= 0 && t.x1 <= W1 + 2 && t.y0 >= -2 && t.y1 <= H1 + 2)
    .sort((a, b) => Math.abs(a.y1 - b.y1) < 6 ? (pageRtl ? b.x1 - a.x1 : a.x0 - b.x0) : a.y1 - b.y1);
}

/* ---------------------------------------------------------------- המצב */
const F = { list: [], cur: null, run: 0, busy: false, doc: null };
// each open file has its own places (switching between files in tabs brings back that file's places and what was filled)
const FOUND = new WeakMap();
let ensureT = 0;
function ensure() {
  if (!S.doc || F.doc === S.doc || S.office) return true;
  clearTimeout(ensureT);
  ensureT = setTimeout(() => {
    if (!S.doc || F.doc === S.doc) return;
    if (F.cur) { pnl.hidden = true; F.cur = null; document.body.classList.remove('fz-filling'); }
    if (FOUND.has(S.doc)) { F.doc = S.doc; F.list = FOUND.get(S.doc); paintAll(); bar(); }
    else run().catch(e => console.warn('fill', e));
  }, 0);
  return false;
}
const annOf = t => t.ann && S.anns.find(a => a.id === t.ann);
const open = () => F.list.filter(t => !annOf(t));

function layerOf(P) {
  if (!P.fzl || !P.fzl.isConnected) { P.fzl = document.createElement('div'); P.fzl.className = 'fzl'; P.ann.append(P.fzl); }
  return P.fzl;
}
function paint(P) {
  if (!P || !P.ann) return;
  if (!ensure()) { if (P.fzl) P.fzl.innerHTML = ''; return; }
  const ly = layerOf(P); ly.innerHTML = '';
  const vp1 = P.page.getViewport({ scale: 1 });
  for (const t of F.list) { if (t.page !== P.i - 1) continue;
    const d = document.createElement('div'); d.className = 'fz ' + t.kind + (annOf(t) ? ' done' : '') + (F.cur === t ? ' cur' : '');
    d.tabIndex = annOf(t) ? -1 : 0; d.setAttribute('role', 'button');
    const nm = kindName(t);
    d.setAttribute('aria-label', nm + (t.label ? ': ' + t.label : ''));
    Object.assign(d.style, { left: t.x0 / vp1.width * 100 + '%', top: t.y0 / vp1.height * 100 + '%', width: (t.x1 - t.x0) / vp1.width * 100 + '%', height: (t.y1 - t.y0) / vp1.height * 100 + '%' });
    const k = document.createElement('span'); k.className = 'fzk'; k.textContent = nm; d.append(k);
    d.addEventListener('pointerdown', e => e.stopPropagation());
    d.addEventListener('click', e => { e.preventDefault(); e.stopPropagation(); act(t); });
    d.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); act(t); } });
    ly.append(d);
  }
}
const paintAll = () => S.pages.forEach(paint);

/* ---------------------------------------------------------------- מילוי */
// the text box sits on the line: the size of the words beside it, its start at the start of the line
function sizeFor(t) {
  const room = t.kind === 'cell' ? (t.y1 - t.y0) * .62 : (t.y1 - t.y0) * .75;
  const base = t.fh ? t.fh * .95 : 11;
  return Math.round(clamp(Math.min(base, room), 7, 14) * 2) / 2;
}
const meas = document.createElement('canvas').getContext('2d');
function natural(a) {
  meas.font = `${a.bold ? 700 : 400} ${a.size}px ${typeof cssFont === 'function' ? cssFont(a.font) : 'Arial'}`;
  return String(a.text || '').split('\n').reduce((m, l) => Math.max(m, meas.measureText(l).width), 0) + a.size * .4 + 2;
}
// grow: wider as you type (up to the page margin), then more lines
function fit(a) {
  const f = a.fz; if (!f) return; const P = S.pages[a.page]; if (!P) return;
  const vp1 = P.page.getViewport({ scale: 1 });
  const fs = [...String(a.text || '')].find(c => RTL_CH.test(c) || LTR_CH.test(c));
  const rtl = fs ? RTL_CH.test(fs) : f.rtl;
  const max = rtl ? f.r - (f.limL != null ? f.limL : 14) : (f.limR != null ? f.limR : vp1.width - 14) - f.l;
  const w = clamp(natural(a), Math.max(24, f.r - f.l), Math.max(40, max));
  const left = rtl ? f.r - w : f.l;
  const pt = vp1.convertToPdfPoint(left, f.top);
  a.w = w; a.x = pt[0]; a.y = pt[1];
  const el = P.tl.querySelector(`.tw[data-id="${a.id}"]`);
  if (el) { const [l, tp] = toCss(P, a.x, a.y); el.style.left = l + 'px'; el.style.top = tp + 'px'; const tx = el.querySelector('.tx'); if (tx) { tx.style.width = a.w * S.zoom + 'px'; tx.style.resize = 'none'; tx.dir = rtl ? 'rtl' : 'ltr'; tx.style.textAlign = rtl ? 'right' : 'left'; } }
}
function textAt(t, text) {
  const P = S.pages[t.page], vp1 = P.page.getViewport({ scale: 1 }), size = sizeFor(t);
  // the first line's letters stand on the line (or in the middle of a table cell)
  const top = t.kind === 'cell' ? t.y0 + Math.max(1, ((t.y1 - t.y0) - size * 1.35) / 2) : t.y1 - 1.5 - size * 1.08;
  const pt = vp1.convertToPdfPoint(t.rtl ? t.x1 - 30 : t.x0, top);
  t.hist = S.hist.length; snapshot();
  const a = { id: S.nextId++, type: 'text', page: t.page, x: pt[0], y: pt[1], w: Math.max(24, t.x1 - t.x0), size, color: (S.color && S.color.text) || '#1C2333', text: text || '', font: S.textFont, bold: false,
    fz: { l: t.x0, r: t.x1, top, rtl: t.rtl, limL: t.limL, limR: t.limR } };
  S.anns.push(a); t.ann = a.id; fit(a); drawAnns(P);
  return a;
}
function checkSrc() {
  const c = document.createElement('canvas'), n = c.width = c.height = 96, g = c.getContext('2d');
  g.strokeStyle = '#14213D'; g.lineWidth = n * .12; g.lineCap = 'round'; g.lineJoin = 'round';
  g.beginPath(); g.moveTo(n * .18, n * .54); g.lineTo(n * .42, n * .78); g.lineTo(n * .84, n * .22); g.stroke();
  return c.toDataURL('image/png');
}
let CHECK = null;
function imgAt(t, src, ratio, sig) {
  const P = S.pages[t.page]; let w, cx, cy;
  if (t.kind === 'check') { w = Math.max(8, Math.min(t.x1 - t.x0, t.y1 - t.y0) * 1.15); cx = (t.x0 + t.x1) / 2; cy = (t.y0 + t.y1) / 2; }
  else { // a signature stands on its line, a little over it, like a pen would
    const hMax = Math.max(26, (t.y1 - t.y0) * 2.2); w = Math.min(t.x1 - t.x0, 170); if (w * ratio > hMax) w = hMax / ratio;
    const h = w * ratio; cx = t.rtl ? t.x1 - w / 2 - 2 : t.x0 + w / 2 + 2; cy = t.y1 - h / 2 + h * .12;
    if (t.kind === 'cell') cy = (t.y0 + t.y1) / 2;
  }
  // addImage works in screen pixels at the current zoom
  const sc = S.zoom; addImage(P, src, ratio, w, cx * sc, cy * sc, sig);
  t.ann = S.nextId - 1; if (typeof pickAnn === 'function') pickAnn(null);
}
function sigInto(t) {
  const sigs = typeof loadSigs === 'function' ? loadSigs() : [];
  if (sigs.length) { if (t.ann) dropAnn(t); imgAt(t, sigs[0].src, sigs[0].ratio, true); return true; }
  F.sigFor = t; openSig(); return false;
}
function dropAnn(t) { const a = annOf(t); t.ann = null; if (a && typeof removeAnn === 'function') removeAnn(a.id); }

function act(t) {
  if (F.doc !== S.doc || !F.list.includes(t)) return;   // a place of a file that is no longer the one on screen
  if (F.cur && F.cur !== t) closePanel(true);
  if (t.kind === 'check') {
    if (annOf(t)) dropAnn(t); else imgAt(t, CHECK || (CHECK = checkSrc()), 1, false);
    paint(S.pages[t.page]); bar(); return;
  }
  if (t.kind === 'sig' && !annOf(t)) { if (sigInto(t)) { paint(S.pages[t.page]); bar(); openPanel(t); } return; }
  openPanel(t);
}

/* ---------------------------------------------------------------- חלונית המילוי */
const pnl = document.createElement('div'); pnl.className = 'fzp'; pnl.hidden = true; pnl.setAttribute('role', 'dialog');
pnl.innerHTML = `<div class="hd"><b class="ttl"></b><span class="lab"></span><span class="cnt"></span></div>
<div class="lens" aria-hidden="true"><canvas></canvas></div>
<textarea rows="1" dir="auto" enterkeyhint="next" autocomplete="on"></textarea>
<div class="warn" role="status" hidden></div>
<div class="mem" hidden></div>
<div class="row"><span class="chips"></span><span class="sp"></span><button type="button" data-a="del"></button><button type="button" data-a="prev"></button><button type="button" class="main" data-a="next"></button><button type="button" data-a="done"></button></div>
<div class="tip"></div>`;
document.body.append(pnl);
const ta = pnl.querySelector('textarea');
const grow = () => { ta.style.height = 'auto'; ta.style.height = Math.min(ta.scrollHeight + 4, innerHeight * .4) + 'px'; };

/* 5.9.6: the magnifier — on a phone the keyboard and this window cover the page, so the window itself shows
   the piece of the form being filled (the line and the words beside it), enlarged, with the text written into it live */
const lensBox = pnl.querySelector('.lens'), lensC = lensBox.querySelector('canvas');
const LZ = { t: null, img: null, crop: null, s: 1, job: 0 };
async function lensOpen(t) {
  const my = ++LZ.job; LZ.t = t; LZ.img = null;
  const P = S.pages[t.page]; if (!P) { lensBox.hidden = true; return; }
  const vp1 = P.page.getViewport({ scale: 1 }), W = vp1.width;
  // the field and the words beside it (Hebrew: on its right), at most ~300pt wide
  const side = 210, maxW = 380;
  let x0 = t.rtl ? t.x0 - 8 : t.x0 - side, x1 = t.rtl ? t.x1 + side : t.x1 + 8;
  x0 = Math.max(0, x0); x1 = Math.min(W, x1);
  if (x1 - x0 > maxW) { if (t.rtl) x0 = x1 - maxW; else x1 = x0 + maxW; }
  const y0 = Math.max(0, t.y0 - (t.kind === 'cell' ? 4 : 10)), y1 = Math.min(vp1.height, t.y1 + 8);
  const cssW = Math.max(200, Math.min(620, innerWidth) - 28);
  const s = Math.min(3.2, cssW / (x1 - x0)), dpr = Math.min(2, window.devicePixelRatio || 1), k = s * dpr;
  const c = document.createElement('canvas'); c.width = Math.round((x1 - x0) * k); c.height = Math.round((y1 - y0) * k);
  const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
  lensBox.hidden = false;
  lensC.width = c.width; lensC.height = c.height; lensC.style.width = (x1 - x0) * s + 'px'; lensC.style.height = (y1 - y0) * s + 'px';
  try { await P.page.render({ canvasContext: g, viewport: P.page.getViewport({ scale: k }), transform: [1, 0, 0, 1, -x0 * k, -y0 * k], annotationMode: 0 }).promise; }
  catch (e) { if (my === LZ.job) lensBox.hidden = true; return; }
  if (my !== LZ.job) return;
  LZ.img = c; LZ.crop = { x0, y0, x1, y1 }; LZ.s = k;
  lensDraw();
}
function lensDraw() {
  const t = LZ.t, c = LZ.img; if (!t || !c || lensBox.hidden) return;
  const g = lensC.getContext('2d'), k = LZ.s, { x0, y0 } = LZ.crop, X = v => (v - x0) * k, Y = v => (v - y0) * k;
  g.clearRect(0, 0, lensC.width, lensC.height); g.drawImage(c, 0, 0);
  const a = annOf(t);
  // the field itself, lightly marked
  g.fillStyle = 'rgba(255,216,77,.28)'; g.fillRect(X(t.x0), Y(t.y0), (t.x1 - t.x0) * k, (t.y1 - t.y0) * k);
  if (!a || a.type !== 'text' || !String(a.text || '').trim()) return;
  const size = a.size, f = a.fz || {};
  const fs = [...a.text].find(ch => RTL_CH.test(ch) || LTR_CH.test(ch)), rtl = fs ? RTL_CH.test(fs) : f.rtl;
  g.font = `${a.bold ? 700 : 400} ${size * k}px ${typeof cssFont === 'function' ? cssFont(a.font) : 'Arial'}`;
  g.fillStyle = a.color || '#1C2333'; g.textBaseline = 'alphabetic'; g.direction = rtl ? 'rtl' : 'ltr'; g.textAlign = rtl ? 'right' : 'left';
  const base = (f.top != null ? f.top : t.y1 - size * 1.08) + size * 1.08 - .5;
  const lines = String(a.text).split('\n');
  lines.forEach((ln, i) => g.fillText(ln, rtl ? X(f.r != null ? f.r : t.x1) : X(f.l != null ? f.l : t.x0), Y(base + i * size * 1.35)));
}
function openPanel(t) {
  F.cur = t; const l = L(), P = S.pages[t.page];
  let a = annOf(t);
  const isText = !(a && a.type === 'img');
  if (isText && !a) a = textAt(t, '');
  const nm = kindName(t);
  pnl.querySelector('.ttl').textContent = nm;
  pnl.querySelector('.lab').textContent = t.label && t.label.length < 40 ? '· ' + t.label.replace(BLANK, '').replace(/[:：]\s*$/, '') : '';
  pnl.querySelector('[data-a=next]').textContent = l.next + ' ›';
  pnl.querySelector('[data-a=done]').textContent = l.done;
  pnl.querySelector('[data-a=del]').textContent = l.del;
  { const an = annOf(t); pnl.querySelector('[data-a=del]').hidden = !an || (an.type === 'text' && !String(an.text || '').trim()); }
  pnl.querySelector('.tip').textContent = LX().tipK;
  { const pv = pnl.querySelector('[data-a=prev]'); pv.textContent = '‹ ' + LX().prev; pv.hidden = F.list.indexOf(t) <= 0; }
  const chips = pnl.querySelector('.chips'); chips.innerHTML = '';
  const chip = (txt, fn, cls) => { const b = document.createElement('button'); b.type = 'button'; b.textContent = txt; if (cls) b.className = cls; b.onclick = fn; chips.append(b); };
  const kd = t.kind === 'cell' ? t.sub : t.kind;
  if (kd === 'date') chip(l.today + ': ' + new Date().toLocaleDateString('he-IL'), () => { ta.value = new Date().toLocaleDateString('he-IL'); sync(); ta.focus(); }, 'sun');
  // 6.2: the Hebrew date too (school and community forms often ask for it)
  { const hd = kd === 'date' && window.SIMAN_HEBDATE ? SIMAN_HEBDATE() : ''; if (hd) chip('תאריך עברי: ' + hd, () => { ta.value = hd; sync(); ta.focus(); }); }
  if (isText && (t.kind === 'sig' || t.kind === 'cell' || kd === 'sig')) chip('✍ ' + l.asSig, () => { dropAnn(t); if (sigInto(t)) openPanel(t); paint(P); bar(); });
  if (!isText) chip('✍ ' + l.otherSig, () => { dropAnn(t); F.sigFor = t; closePanel(true); openSig(); });
  if (isText && (t.kind === 'cell' || (t.x1 - t.x0) < 60)) chip(l.asV, () => { dropAnn(t); imgAt(t, CHECK || (CHECK = checkSrc()), 1, false); paint(P); bar(); closePanel(true); next(); });
  if (!isText) chip(l.asText, () => { dropAnn(t); openPanel(t); paint(P); });
  ta.hidden = !isText;
  if (isText) { ta.value = a.text || ''; ta.placeholder = l.ph; ta.dir = 'auto'; ta.inputMode = kd === 'id' ? 'numeric' : kd === 'phone' ? 'tel' : kd === 'email' ? 'email' : 'text'; }
  memShow(isText ? kd : null);
  checkNow();
  count(); pnl.hidden = false; document.body.classList.add('fz-filling'); paint(P); place(); bar();
  if (isText) { ta.focus({ preventScroll: true }); grow(); try { ta.setSelectionRange(ta.value.length, ta.value.length); } catch (_) {} }
  lensOpen(t);
  setTimeout(() => show(t), 30);
}
function sync() {
  const t = F.cur, a = t && annOf(t); if (!a || a.type !== 'text') return;
  a.text = ta.value; setDirty(true); fit(a);
  const tx = (S.pages[a.page] ? S.pages[a.page].tl : document).querySelector(`.tw[data-id="${a.id}"] .tx`); if (tx && tx.textContent !== a.text) tx.textContent = a.text;
  grow(); pnl.querySelector('[data-a=del]').hidden = !a.text; lensDraw(); checkNow(); memShow(kindOfCur());
}
const kindOfCur = () => { const t = F.cur; return t ? (t.kind === 'cell' ? t.sub : t.kind) : null; };
// a gentle note under the box (an ID whose check digit is wrong, a short phone number); it never blocks
function checkNow() { const w = pnl.querySelector('.warn'), m = warnFor(kindOfCur(), ta.value); w.textContent = m; w.hidden = !m || ta.hidden; }
// what was typed in this kind of place before (only on this device): one tap fills it
function memShow(kd) {
  const box = pnl.querySelector('.mem'); box.innerHTML = '';
  const list = kd && MEM_KINDS.includes(kd) ? (memGet()[kd] || []).filter(v => v !== ta.value.trim()).slice(0, 3) : [];
  box.hidden = !list.length;
  for (const v of list) { const b = document.createElement('button'); b.type = 'button'; b.textContent = v; b.title = v;
    b.onclick = () => { ta.value = v; sync(); ta.focus(); }; box.append(b); }
}
// the value is remembered when the place is left
function remember(t) { const a = t && annOf(t); if (!a || a.type !== 'text') return; const kd = t.kind === 'cell' ? t.sub : t.kind; memPut(kd, a.text); }
ta.addEventListener('input', sync);
ta.addEventListener('keydown', e => {
  if (e.key === 'Enter' && !e.shiftKey && !e.isComposing) { e.preventDefault(); next(); }
  else if (e.key === 'Tab' && !e.isComposing) { e.preventDefault(); if (e.shiftKey) prev(); else next(); }
  else if (e.key === 'Escape') { e.preventDefault(); closePanel(); }
});
function prev() {
  const cur = F.cur, i = cur ? F.list.indexOf(cur) : -1; if (i <= 0) return;
  closePanel(true); act(F.list[i - 1]);
}
function closePanel(keepCur) {
  const t = F.cur; remember(t); pnl.hidden = true; F.cur = null; document.body.classList.remove('fz-filling');
  if (t) { const a = annOf(t); if (a && a.type === 'text' && !String(a.text || '').trim()) { const k = S.anns.indexOf(a); if (k >= 0) { S.anns.splice(k, 1); if (S.hist.length === t.hist + 1) S.hist.pop(); } t.ann = null; drawAnns(S.pages[t.page]); } paint(S.pages[t.page]); }
  if (!keepCur && document.activeElement === ta) ta.blur();
  bar();
}
function next() {
  const cur = F.cur, list = F.list, i = cur ? list.indexOf(cur) : -1;
  closePanel(true);
  const rest = [...list.slice(i + 1), ...list.slice(0, Math.max(0, i))].filter(t => !annOf(t));
  if (!rest.length) { pnl.hidden = true; toast(L().all); ta.blur(); bar(); const g = bar.el.querySelector('.go'); if (g && !bar.el.hidden) g.focus({ preventScroll: true }); return; }
  act(rest[0]);
}
pnl.querySelector('[data-a=next]').onclick = () => next();
pnl.querySelector('[data-a=prev]').onclick = () => prev();
pnl.querySelector('[data-a=done]').onclick = () => closePanel();
pnl.querySelector('[data-a=del]').onclick = () => { const t = F.cur; if (!t) return; ta.value = ''; dropAnn(t); closePanel(); };
function count() { const n = open().length; pnl.querySelector('.cnt').textContent = n ? L().left(n) : L().all; }
function show(t) {
  const P = S.pages[t.page]; if (!P) return; const d = $('desk'); if (!d) return;
  const vp1 = P.page.getViewport({ scale: 1 }), k = S.zoom;
  const pr = P.el.getBoundingClientRect(), dr = d.getBoundingClientRect();
  const y = pr.top + (t.y0 + t.y1) / 2 * k, x = pr.left + (t.x0 + t.x1) / 2 * k;
  const bottom = pnl.hidden ? dr.bottom : Math.min(dr.bottom, pnl.getBoundingClientRect().top);
  const want = dr.top + (bottom - dr.top) * .42;
  if (y < dr.top + 40 || y > bottom - 30) d.scrollBy({ top: y - want, behavior: 'smooth' });
  if (x < dr.left + 10 || x > dr.right - 10) d.scrollBy({ left: x - (dr.left + dr.right) / 2, behavior: 'smooth' });
  void vp1;
}
// above the phone's keyboard
function place() {
  const vv = window.visualViewport; if (!vv) return;
  const off = Math.max(0, innerHeight - vv.height - vv.offsetTop);
  pnl.style.bottom = off + 'px'; bar.el.style.bottom = '';
}
if (window.visualViewport) { visualViewport.addEventListener('resize', place); visualViewport.addEventListener('scroll', place); }

/* ---------------------------------------------------------------- הפס "מקומות למילוי" */
function bar() {
  const n = open().length, b = bar.el;
  b.hidden = !F.list.length || !pnl.hidden || !S.doc || !!S.office || F.doc !== S.doc;
  b.querySelector('.go').textContent = n ? L().show(n) + ' ›' : LX().save;
  b.querySelector('.go').classList.toggle('fin', !n);
  b.querySelector('.x').title = L().hide; b.querySelector('.x').setAttribute('aria-label', L().hide);
  b.querySelector('.x').textContent = document.body.classList.contains('fz-off') ? '👁' : '×';
}
bar.el = document.createElement('div'); bar.el.className = 'fzbar'; bar.el.hidden = true;
bar.el.innerHTML = '<button type="button" class="go"></button><button type="button" class="x"></button>';
document.body.append(bar.el);
bar.el.querySelector('.go').onclick = () => {
  // everything is filled: the same button saves the form
  if (!open().length) { const sb = document.getElementById('saveBtn'); if (sb && !sb.disabled) sb.click(); return; }
  document.body.classList.remove('fz-off'); if (typeof setTool === 'function' && S.tool !== 'select' && S.tool !== 'text') setTool('select'); const r = open(); if (r.length) act(r[0]); };
bar.el.querySelector('.x').onclick = () => { document.body.classList.toggle('fz-off'); bar(); };

/* ---------------------------------------------------------------- חיבור לאפליקציה */
async function run() {
  const my = ++F.run; F.list = []; F.cur = null; pnl.hidden = true; document.body.classList.remove("fz-filling");
  F.doc = S.doc; if (S.doc) FOUND.set(S.doc, F.list); bar();
  S.pages.forEach(P => { if (P.fzl) P.fzl.innerHTML = ''; });
  if (!S.doc || S.office || !S.pages.length) return;
  await sleep(400);                                      // the first page is drawn first
  if (my !== F.run) return;
  try { if (typeof loadSend === 'function') await loadSend(); } catch (_) {}
  const pages = S.pages.slice(0, 30);
  for (const P of pages) {
    if (my !== F.run) return;
    let widgets = [];
    try { const vp1 = P.page.getViewport({ scale: 1 }); widgets = (await P.page.getAnnotations()).filter(a => a.subtype === 'Widget' && !a.hidden).map(a => { const r = vp1.convertToViewportRectangle(a.rect); return [Math.min(r[0], r[2]), Math.min(r[1], r[3]), Math.max(r[0], r[2]), Math.max(r[1], r[3])]; }); } catch (_) {}
    let found = [];
    try { found = await detectPage(P, widgets); } catch (e) { console.warn('fill: page', P.i, e); }
    if (my !== F.run) return;
    // a page full of "lines" is a lined page or a table of data, not a form
    if (found.length > 70) found = found.filter(t => t.kind !== 'text' || t.label);
    F.list.push(...found.slice(0, 90)); paint(P); bar();
    await sleep(0);
  }
  const n = F.list.length;
  if (n && !S.hasForm) toast(L().found(n));
}

// the marks follow the page: drawn again whenever the page's notes are drawn (zoom, undo)
const _draw = window.drawAnns;
if (typeof _draw === 'function') window.drawAnns = function (P) { _draw.apply(this, arguments); try { for (const a of S.anns) if (a.fz && a.type === 'text' && a.page === P.i - 1) fit(a); paint(P); } catch (_) {} };
// a signature made now for a marked place goes straight onto its line
const _place = window.startPlace;
if (typeof _place === 'function') window.startPlace = function (src, ratio, wPt, sig) {
  const t = F.sigFor; F.sigFor = null;
  if (t && sig && S.pages[t.page]) { imgAt(t, src, ratio, true); paint(S.pages[t.page]); bar(); return; }
  return _place.apply(this, arguments);
};
const _undo = window.undo;
if (typeof _undo === 'function') window.undo = function () { const r = _undo.apply(this, arguments); if (F.cur && !annOf(F.cur)) { pnl.hidden = true; F.cur = null; document.body.classList.remove("fz-filling"); } paintAll(); bar(); return r; };
// typing straight on the page grows the box too
document.addEventListener('input', e => {
  const tw = e.target.closest && e.target.closest('.tw'); if (!tw) return;
  const a = S.anns.find(x => x.id === +tw.dataset.id); if (a && a.fz) { a.text = e.target.innerText.replace(/\n$/, ''); fit(a); if (F.cur && F.cur.ann === a.id && ta.value !== a.text) { ta.value = a.text; grow(); } }
}, true);
document.addEventListener('keydown', e => { if (e.key === 'Escape' && !pnl.hidden && document.activeElement !== ta) closePanel(); });
// the signature window closed without signing: the marked place is not kept for a later signature
{ const sd = document.getElementById('sigDlg'); if (sd) sd.addEventListener('close', () => setTimeout(() => { F.sigFor = null; }, 0)); }

window.SIMAN_FILL = { run, list: () => F.list, next, close: closePanel, detectPage, _t: { inkOf, strokes, pieces, cellTop, inkIn } };
})();
