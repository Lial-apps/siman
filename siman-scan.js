/* =====================================================================
   סימן 5.4 — siman-scan.js
   מדף למסמך: סורק מצלמה מלא, בסגנון אפליקציות הסריקה המובילות
   · מצלמה חיה עם זיהוי הדף בזמן אמת (במכשיר, בלי לשלוח לשום שרת)
   · צילום אוטומטי כשהדף יציב, פנס, כמה עמודים ברצף
   · מצבים: מסמך, ספר (שני עמודים בצילום אחד), תעודה (שני צדדים בעמוד אחד), לוח, קוד QR
   · יישור וחיתוך הדף, סיבוב, מסננים (משופר, ללא צללים, אפור, שחור-לבן), מחיקת עט כחול/אדום
   · סידור, מחיקה והוספה של עמודים
   · שמירה: PDF, PDF עם טקסט לחיפוש, Word לעריכה, תמונות JPG, שיתוף, או פתיחה בסימן לחתימה וסימון
   נטען אחרי siman-plus.js ומשתמש ב: openFile, runOCR, S, toast, needReg, showReg, PDFLib, window.SIMAN_SCAN
   ===================================================================== */
(() => {
'use strict';
if (window.SIMAN_CAM) return;

const h = html => { const t = document.createElement('template'); t.innerHTML = html.trim(); return t.content.firstElementChild; };
const esc = t => String(t ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
const sleep = ms => new Promise(r => setTimeout(r, ms));
const FULL = [[0, 0], [1, 0], [1, 1], [0, 1]];
const isFull = q => !q || q.every((p, i) => Math.abs(p[0] - FULL[i][0]) < 0.004 && Math.abs(p[1] - FULL[i][1]) < 0.004);
const quadArea = q => Math.abs(q.reduce((a, p, i) => { const n = q[(i + 1) % 4]; return a + p[0] * n[1] - n[0] * p[1]; }, 0)) / 2;
const LS = { get: k => { try { return localStorage.getItem(k); } catch (_) { return null; } }, set: (k, v) => { try { localStorage.setItem(k, v); } catch (_) {} } };
let toastT = 0;
const say = m => {
  const r = document.querySelector('.scn');
  if (r && !r.hidden) { const t = r.querySelector('.scn-toast'); t.textContent = m; t.hidden = false; clearTimeout(toastT); toastT = setTimeout(() => { t.hidden = true; }, 3800); return; }
  if (typeof toast === 'function') toast(m);
};

const SVG = (p, w = 1.9) => `<svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="${w}" stroke-linecap="round" stroke-linejoin="round">${p}</svg>`;
const I = {
  x: SVG('<path d="M6 6l12 12M18 6L6 18"/>', 2.2),
  torch: SVG('<path d="M8 2h8l-1 6H9z"/><path d="M9 8l1 14h4l1-14"/><path d="M12 13v3"/>'),
  auto: SVG('<path d="M4 8V5a1 1 0 011-1h3M16 4h3a1 1 0 011 1v3M20 16v3a1 1 0 01-1 1h-3M8 20H5a1 1 0 01-1-1v-3"/><path d="M9 15l3-7 3 7M10.2 12.5h3.6"/>'),
  gallery: SVG('<rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="9" cy="10" r="1.6"/><path d="M21 16l-5-5-8 8"/>'),
  rot: SVG('<path d="M20 11a8 8 0 10-2.3 5.7"/><path d="M20 4v7h-7"/>'),
  crop: SVG('<path d="M6 2v14a2 2 0 002 2h14"/><path d="M18 22V8a2 2 0 00-2-2H2"/>'),
  trash: SVG('<path d="M4 7h16M9 7V4h6v3M6 7l1 13h10l1-13"/>'),
  erase: SVG('<path d="M16 3l5 5-11 11H5v-5z"/><path d="M13 6l5 5"/><path d="M3 21h18"/>'),
  plus: SVG('<path d="M12 5v14M5 12h14"/>', 2.2),
  pdf: SVG('<path d="M7 3h7l4 4v14H7z"/><path d="M14 3v4h4"/><path d="M9.5 13h5M9.5 16.5h5"/>'),
  word: SVG('<path d="M7 3h7l4 4v14H7z"/><path d="M14 3v4h4"/><path d="M9 11l1.2 6 1.8-4.5 1.8 4.5L15 11"/>'),
  xls: SVG('<path d="M7 3h7l4 4v14H7z"/><path d="M14 3v4h4"/><path d="M9.5 11.5l4 5M13.5 11.5l-4 5"/>'),
  pen: SVG('<path d="M4 20l4-1L19 8l-3-3L5 16z"/><path d="M14 7l3 3"/>'),
  search: SVG('<circle cx="11" cy="11" r="6"/><path d="M20 20l-4.2-4.2"/>'),
  img: SVG('<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 16l5-5 4 4 3-3 6 6"/>'),
  share: SVG('<circle cx="18" cy="5" r="2.5"/><circle cx="6" cy="12" r="2.5"/><circle cx="18" cy="19" r="2.5"/><path d="M8.3 10.8l7.4-4.3M8.3 13.2l7.4 4.3"/>'),
  next: SVG('<path d="M15 6l-6 6 6 6"/>', 2.2),
  prev: SVG('<path d="M9 6l6 6-6 6"/>', 2.2),
  ok: SVG('<path d="M5 12.5l4.5 4.5L19 7"/>', 2.4),
};

/* ---------------------------------------------------------------- עיצוב */
document.head.append(h(`<style id="simanCamCss">
.scn{position:fixed;inset:0;z-index:2147483100;display:flex;flex-direction:column;background:#0A1222;color:#EEF3FB;font-family:inherit;direction:rtl;text-align:right;-webkit-user-select:none;user-select:none;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}
.scn[hidden]{display:none}
.scn *{box-sizing:border-box}
.scn [hidden]:not(.scn-stack){display:none!important}
.scn button{font:inherit;color:inherit;cursor:pointer}
.scn button:focus-visible,.scn input:focus-visible,.scn select:focus-visible{outline:2px solid var(--v7-accent,#4C9BF5);outline-offset:2px}
.scn svg{width:22px;height:22px;flex:none}
.scn-view{flex:1;display:flex;flex-direction:column;min-height:0}
.scn-view[hidden]{display:none}
.scn-top{flex:none;display:flex;align-items:center;gap:8px;padding:10px 12px;min-height:56px}
.scn-top h2{margin:0;font-size:17px;font-weight:700;flex:1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.scn-ib{width:42px;height:42px;border-radius:12px;border:0;background:rgba(255,255,255,.08);display:grid;place-items:center}
.scn-ib:hover{background:rgba(255,255,255,.15)}
.scn-ib[aria-pressed="true"]{background:var(--v7-accent,#4C9BF5);color:#fff}
.scn-ib[hidden]{display:none}
.scn-pill{height:36px;padding:0 12px;border-radius:999px;border:1px solid rgba(255,255,255,.22);background:transparent;display:inline-flex;align-items:center;gap:6px;font-size:13.5px;white-space:nowrap}
.scn-pill svg{width:18px;height:18px}
.scn-pill[aria-pressed="true"]{background:#fff;color:#0A1222;border-color:#fff}
.scn-stage{flex:1;min-height:0;position:relative;display:grid;place-items:center;overflow:hidden}
.scn-vf{position:relative;border-radius:14px;overflow:hidden;background:#000}
.scn-vf video{display:block;width:100%;height:100%;object-fit:fill}
.scn-vf svg.ov{position:absolute;inset:0;width:100%;height:100%}
.scn-vf .ov polygon{fill:color-mix(in srgb,var(--v7-accent,#4C9BF5) 22%,transparent);stroke:var(--v7-accent,#4C9BF5);stroke-width:.6;vector-effect:non-scaling-stroke;stroke-width:3px;stroke-linejoin:round;transition:all .12s linear}
.scn-vf .ov polygon.ok{stroke:#3DDC97;fill:rgba(61,220,151,.18)}
.scn-vf .grid{position:absolute;inset:0;pointer-events:none;background-image:linear-gradient(rgba(255,255,255,.14) 1px,transparent 1px),linear-gradient(90deg,rgba(255,255,255,.14) 1px,transparent 1px);background-size:33.34% 33.34%;background-position:-1px -1px}
.scn-vf .idbox{position:absolute;inset:22% 8%;border:2px dashed rgba(255,255,255,.75);border-radius:14px;pointer-events:none}
.scn-vf .bookline{position:absolute;top:6%;bottom:6%;left:50%;border-left:2px dashed rgba(255,255,255,.6);pointer-events:none}
.scn-vf .qrbox{position:absolute;width:56%;aspect-ratio:1;left:22%;top:50%;transform:translateY(-50%);border:3px solid #fff;border-radius:18px;box-shadow:0 0 0 2000px rgba(0,0,0,.35);pointer-events:none}
.scn-flash{position:absolute;inset:0;background:#fff;opacity:0;pointer-events:none}
.scn-flash.go{animation:scnFlash .35s ease-out}
@keyframes scnFlash{0%{opacity:.85}100%{opacity:0}}
.scn-hint{position:absolute;top:12px;left:50%;transform:translateX(-50%);padding:7px 14px;border-radius:999px;background:rgba(10,18,34,.72);font-size:14px;white-space:nowrap;pointer-events:none;max-width:92%;overflow:hidden;text-overflow:ellipsis}
.scn-nocam{max-width:420px;padding:24px;text-align:center;line-height:1.6}
.scn-nocam h3{margin:0 0 6px;font-size:18px}
.scn-nocam p{margin:0 0 16px;color:#B8C4D8;font-size:14.5px}
.scn-nocam .row{display:flex;gap:10px;justify-content:center;flex-wrap:wrap}
.scn-modes{flex:none;display:flex;gap:4px;padding:10px 8px 2px;overflow-x:auto;scrollbar-width:none}
.scn-modes button{border:0;background:none;padding:7px 13px;border-radius:999px;font-size:14.5px;color:#9FB0C9;white-space:nowrap}
.scn-modes button[aria-pressed="true"]{color:#0A1222;background:#FFD43B;font-weight:700}
.scn-bottom{flex:none;display:grid;grid-template-columns:1fr auto 1fr;align-items:center;padding:12px 18px 18px}
.scn-shutter{width:76px;height:76px;border-radius:50%;border:4px solid #fff;background:transparent;display:grid;place-items:center;position:relative;padding:0}
.scn-shutter i{display:block;width:60px;height:60px;border-radius:50%;background:#fff;transition:transform .1s}
.scn-shutter:active i{transform:scale(.9)}
.scn-shutter:disabled{opacity:.4}
.scn-shutter.rec i{width:30px;height:30px;border-radius:8px;background:#E5484D}
.scn-mini{position:absolute;bottom:10px;inset-inline-start:10px;z-index:3;background:rgba(10,18,34,.8);border:1px solid rgba(255,255,255,.35);border-radius:10px;padding:4px;pointer-events:none}
.scn-mini[hidden]{display:none}
.scn-mini canvas{display:block}
.scn-shutter svg.ring{position:absolute;inset:-6px;width:88px;height:88px;transform:rotate(-90deg);pointer-events:none}
.scn-shutter .ring circle{fill:none;stroke:#3DDC97;stroke-width:4;stroke-dasharray:264;stroke-dashoffset:264;transition:stroke-dashoffset .12s linear}
.scn-stack{justify-self:start;position:relative;width:58px;height:58px;border:0;padding:0;background:none}
.scn-stack img{width:100%;height:100%;object-fit:cover;border-radius:10px;border:2px solid #fff;display:block;background:#fff}
.scn-stack b{position:absolute;top:-8px;inset-inline-end:-8px;min-width:24px;height:24px;padding:0 6px;border-radius:12px;background:#FFD43B;color:#0A1222;font-size:13px;display:grid;place-items:center}
.scn-stack[hidden]{visibility:hidden;display:block}
.scn-gal{justify-self:end}
.scn-done{justify-self:start;height:44px;padding:0 18px;border-radius:12px;border:0;background:#FFD43B;color:#0A1222;font-weight:700;font-size:15px}
.scn-body{flex:1;min-height:0;overflow:auto;padding:4px 14px 14px}
.scn-name{display:flex;align-items:center;gap:8px;margin:2px 0 12px}
.scn-name input{flex:1;min-width:0;height:42px;border-radius:10px;border:1px solid rgba(255,255,255,.2);background:rgba(255,255,255,.06);color:#fff;padding:0 12px;font:inherit;font-size:15px;direction:rtl}
.scn-grid{list-style:none;margin:0;padding:0;display:grid;grid-template-columns:repeat(auto-fill,minmax(136px,1fr));gap:14px}
.scn-grid li{position:relative;background:rgba(255,255,255,.05);border-radius:12px;padding:8px;display:flex;flex-direction:column;gap:6px}
.scn-grid .pv{border:0;padding:0;background:#fff;border-radius:6px;overflow:hidden;aspect-ratio:3/4;display:grid;place-items:center}
.scn-grid .pv img{max-width:100%;max-height:100%;display:block}
.scn-grid .row{display:flex;align-items:center;justify-content:space-between;gap:4px;font-size:13px;color:#B8C4D8}
.scn-grid .row button{width:30px;height:30px;border-radius:8px;border:0;background:rgba(255,255,255,.08);display:grid;place-items:center;padding:0}
.scn-grid .row button svg{width:17px;height:17px}
.scn-grid .row button:disabled{opacity:.3}
.scn-grid li.add{border:2px dashed rgba(255,255,255,.25);background:none;align-items:center;justify-content:center;min-height:200px;cursor:pointer;color:#B8C4D8;font-size:14px;text-align:center}
.scn-sec{margin:18px 0 0}
.scn-sec h3{margin:0 0 8px;font-size:14px;color:#9FB0C9;font-weight:600}
.scn-chips{display:flex;gap:8px;flex-wrap:wrap}
.scn-opts{display:flex;gap:16px;flex-wrap:wrap;align-items:center;font-size:14px;color:#D5DEEC}
.scn-opts label{display:flex;align-items:center;gap:6px}
.scn-opts select,.scn-opts input[type=text]{height:36px;border-radius:9px;border:1px solid rgba(255,255,255,.2);background:#14203A;color:#fff;font:inherit;font-size:14px;padding:0 8px}
.scn-opts input[type=text]{width:150px}
.scn-acts{flex:none;display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:8px;padding:12px 14px 14px;border-top:1px solid rgba(255,255,255,.1);background:#0D172B}
.scn-act{min-height:52px;border-radius:12px;border:1px solid rgba(255,255,255,.16);background:rgba(255,255,255,.05);display:flex;align-items:center;gap:10px;padding:8px 12px;text-align:right}
.scn-act b{display:block;font-size:14.5px}
.scn-act span{display:block;font-size:12px;color:#9FB0C9;line-height:1.3}
.scn-act.main{background:#FFD43B;border-color:#FFD43B;color:#0A1222}
.scn-act.main span{color:#3B3413}
.scn-edit{flex:1;min-height:0;display:grid;place-items:center;padding:8px 12px;position:relative}
.scn-edit .canvas{position:relative;max-width:100%;max-height:100%}
.scn-edit img.big{display:block;max-width:100%;max-height:100%;background:#fff;border-radius:4px;box-shadow:0 6px 30px rgba(0,0,0,.5)}
.scn-crop{position:relative;touch-action:none}
.scn-crop img{display:block;max-width:100%;max-height:100%;pointer-events:none}
.scn-crop svg{position:absolute;inset:0;width:100%;height:100%;pointer-events:none}
.scn-crop .shade{fill:rgba(0,0,0,.5)}
.scn-crop .edge{fill:none;stroke:#FFD43B;stroke-width:3px;vector-effect:non-scaling-stroke}
.scn-hd{position:absolute;width:34px;height:34px;margin:-17px 0 0 -17px;border-radius:50%;border:3px solid #FFD43B;background:rgba(255,212,59,.25);touch-action:none;padding:0}
.scn-hd.on{background:#FFD43B}
.scn-tools{flex:none;display:flex;gap:6px;justify-content:center;flex-wrap:wrap;padding:8px 10px}
.scn-tool{min-width:64px;height:58px;border:0;border-radius:12px;background:rgba(255,255,255,.06);display:flex;flex-direction:column;align-items:center;justify-content:center;gap:3px;font-size:12px;padding:4px 8px}
.scn-tool[aria-pressed="true"]{background:#FFD43B;color:#0A1222}
.scn-filters{flex:none;display:flex;gap:8px;overflow-x:auto;padding:4px 12px 12px;scrollbar-width:none}
.scn-filters>:first-child,.scn-modes>:first-child{margin-inline-start:auto}.scn-filters>:last-child,.scn-modes>:last-child{margin-inline-end:auto}
.scn-filters button{flex:none;border:2px solid transparent;border-radius:10px;background:rgba(255,255,255,.06);padding:4px 4px 6px;width:76px;display:flex;flex-direction:column;align-items:center;gap:4px;font-size:12px}
.scn-filters button img{width:64px;height:64px;object-fit:cover;border-radius:6px;background:#fff}
.scn-filters button[aria-pressed="true"]{border-color:#FFD43B}
.scn-busy{position:absolute;inset:0;z-index:5;background:rgba(6,11,22,.78);display:grid;place-items:center;text-align:center;font-size:15.5px;padding:24px}
.scn-busy[hidden]{display:none}
.scn-busy i{display:block;width:42px;height:42px;margin:0 auto 14px;border-radius:50%;border:4px solid rgba(255,255,255,.2);border-top-color:#FFD43B;animation:scnSpin 1s linear infinite}
@keyframes scnSpin{to{transform:rotate(360deg)}}
.scn-sheet{position:absolute;inset-inline:12px;bottom:12px;z-index:4;background:#fff;color:#0A1222;border-radius:16px;padding:16px;box-shadow:0 10px 40px rgba(0,0,0,.5)}
.scn-sheet[hidden]{display:none}
.scn-sheet h3{margin:0 0 6px;font-size:16px}
.scn-sheet .val{direction:ltr;text-align:left;unicode-bidi:isolate;word-break:break-all;background:#F1F4FA;border-radius:10px;padding:10px;font-size:14px;margin:0 0 12px;max-height:30vh;overflow:auto;-webkit-user-select:text;user-select:text}
.scn-sheet .row{display:flex;gap:8px;flex-wrap:wrap}
.scn-sheet button{height:40px;padding:0 14px;border-radius:10px;border:1px solid #CBD5E6;background:#fff;color:#0A1222}
.scn-sheet button.main{background:#0A1222;color:#fff;border-color:#0A1222}
.scn-btn{height:42px;padding:0 16px;border-radius:12px;border:1px solid rgba(255,255,255,.25);background:rgba(255,255,255,.06);display:inline-flex;align-items:center;gap:8px;font-size:14.5px}
.scn-btn.main{background:#FFD43B;border-color:#FFD43B;color:#0A1222;font-weight:700}
.scn-toast{position:absolute;left:50%;bottom:24px;transform:translateX(-50%);z-index:6;max-width:90%;padding:10px 16px;border-radius:12px;background:#fff;color:#0A1222;font-size:14.5px;box-shadow:0 6px 24px rgba(0,0,0,.4);text-align:center}
.scn-toast[hidden]{display:none}
.scn-vf .ov polygon.other{fill:rgba(0,0,0,.35);stroke:rgba(255,255,255,.6);stroke-dasharray:6 5}
.scn-note{margin:10px 0 0;font-size:12.5px;color:#8FA1BC;line-height:1.5}
@media (max-width:560px){.scn-grid{grid-template-columns:repeat(2,1fr)}.scn-acts{grid-template-columns:1fr 1fr}.scn-act span{display:none}.scn-act{min-height:48px}}
@media (prefers-reduced-motion:reduce){.scn-flash.go{animation:none}.scn-vf .ov polygon{transition:none}}
</style>`));

/* ---------------------------------------------------------------- המסך */
const MODES = [['doc', 'מסמך'], ['book', 'ספר'], ['mega', 'מגה סריקה'], ['id', 'תעודה'], ['board', 'לוח'], ['qr', 'קוד QR']];
const FILTERS = [['magic', 'משופר'], ['clean', 'ללא צללים'], ['orig', 'מקורי'], ['gray', 'אפור'], ['bw', 'שחור-לבן'], ['board', 'חזק']];
const root = h(`<div class="scn" hidden role="dialog" aria-modal="true" aria-label="מדף למסמך: סריקה">
  <section class="scn-view" data-v="cam">
    <div class="scn-top">
      <button class="scn-ib" type="button" data-a="close" aria-label="סגירת הסורק">${I.x}</button>
      <h2>מדף למסמך</h2>
      <button class="scn-pill" type="button" data-a="swap" hidden><svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 8h14l-4-4M20 16H6l4 4"/></svg><span>העמוד השני</span></button>
      <button class="scn-pill" type="button" data-a="pages" hidden>${I.crop}<span>עמוד אחד</span></button>
      <button class="scn-pill" type="button" data-a="auto" aria-pressed="true">${I.auto}<span>אוטומטי</span></button>
      <button class="scn-ib" type="button" data-a="torch" aria-pressed="false" aria-label="פנס" hidden>${I.torch}</button>
    </div>
    <div class="scn-stage">
      <div class="scn-vf" hidden><video playsinline muted autoplay></video><div class="grid"></div><div class="guide"></div>
        <svg class="ov" viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true"><polygon class="other" points=""/><polygon class="main" points=""/></svg><div class="scn-flash"></div></div>
      <div class="scn-nocam" hidden></div>
      <p class="scn-hint" role="status" aria-live="polite"></p>
      <div class="scn-mini" hidden aria-hidden="true"><canvas width="150" height="150"></canvas></div>
      <div class="scn-sheet" hidden></div>
    </div>
    <div class="scn-modes" role="group" aria-label="סוג הסריקה">${MODES.map(([k, l]) => `<button type="button" data-m="${k}" aria-pressed="${k === 'doc'}">${l}</button>`).join('')}</div>
    <div class="scn-bottom">
      <button class="scn-stack" type="button" data-a="review" hidden aria-label="מעבר לעמודים שנסרקו"><img alt=""><b>0</b></button>
      <button class="scn-shutter" type="button" data-a="shoot" aria-label="צילום">
        <svg class="ring" viewBox="0 0 88 88" aria-hidden="true"><circle cx="44" cy="44" r="42"/></svg><i></i></button>
      <button class="scn-ib scn-gal" type="button" data-a="gallery" aria-label="ייבוא תמונות מהגלריה">${I.gallery}</button>
    </div>
  </section>

  <section class="scn-view" data-v="review" hidden>
    <div class="scn-top">
      <button class="scn-ib" type="button" data-a="close" aria-label="סגירת הסורק">${I.x}</button>
      <h2 data-r="title">העמודים שנסרקו</h2>
      <button class="scn-pill" type="button" data-a="more">${I.plus}<span>עוד עמודים</span></button>
    </div>
    <div class="scn-body">
      <label class="scn-name"><span>שם הקובץ</span><input type="text" data-r="name" maxlength="80" aria-label="שם הקובץ"></label>
      <ul class="scn-grid" data-r="grid"></ul>
      <div class="scn-sec"><h3>סגנון לכל העמודים</h3><div class="scn-chips" data-r="allf">${FILTERS.map(([k, l]) => `<button class="scn-pill" type="button" data-f="${k}" aria-pressed="false">${l}</button>`).join('')}
        <button class="scn-pill" type="button" data-a="inkall" aria-pressed="false">${I.erase}<span>מחיקת עט כחול ואדום</span></button></div></div>
      <div class="scn-sec"><h3>הגדרות שמירה</h3><div class="scn-opts">
        <label>גודל עמוד <select data-o="size"><option value="auto">כמו הדף שצולם</option><option value="a4">A4</option><option value="letter">Letter</option></select></label>
        <label>איכות <select data-o="q"><option value="high">גבוהה</option><option value="mid">רגילה</option><option value="small">קובץ קטן לשליחה</option></select></label>
        <label>סימן מים <input type="text" data-o="wm" placeholder="ללא" maxlength="40"></label>
      </div><p class="scn-note">הסריקה, היישור והמסננים נעשים במכשיר שלכם. רק ״Word לעריכה״ שולח את התמונות לזיהוי הטקסט.</p></div>
    </div>
    <div class="scn-acts">
      <button class="scn-act main" type="button" data-x="pdf">${I.pdf}<div><b>שמירה כ־PDF</b><span>קובץ סרוק נקי, מוכן לשליחה</span></div></button>
      <button class="scn-act" type="button" data-x="word">${I.word}<div><b>Word לעריכה</b><span>הטקסט נקרא ונבנה מסמך ערוך</span></div></button>
      <button class="scn-act" type="button" data-x="excel">${I.xls}<div><b>טבלה ל־Excel</b><span>הטבלה שבדף, לעריכה כגיליון</span></div></button>
      <button class="scn-act" type="button" data-x="open">${I.pen}<div><b>פתיחה בסימן</b><span>לחתימה, סימון, מילוי ועריכה</span></div></button>
      <button class="scn-act" type="button" data-x="ocr">${I.search}<div><b>PDF עם טקסט</b><span>אפשר לחפש ולהעתיק ממנו</span></div></button>
      <button class="scn-act" type="button" data-x="jpg">${I.img}<div><b>תמונות JPG</b><span>כל עמוד כתמונה</span></div></button>
      <button class="scn-act" type="button" data-x="share" hidden>${I.share}<div><b>שיתוף</b><span>וואטסאפ, מייל ועוד</span></div></button>
    </div>
  </section>

  <section class="scn-view" data-v="edit" hidden>
    <div class="scn-top">
      <button class="scn-pill" type="button" data-a="editdone" aria-label="סיום, חזרה לכל העמודים">${I.ok}<span>סיום</span></button>
      <h2 data-e="title">עמוד 1</h2>
      <button class="scn-ib" type="button" data-a="pprev" aria-label="העמוד הקודם">${I.prev}</button>
      <button class="scn-ib" type="button" data-a="pnext" aria-label="העמוד הבא">${I.next}</button>
    </div>
    <div class="scn-edit" data-e="stage"></div>
    <div class="scn-filters" data-e="filters" role="group" aria-label="מסנן לעמוד"></div>
    <div class="scn-tools" data-e="tools">
      <button class="scn-tool" type="button" data-t="crop">${I.crop}<span>חיתוך</span></button>
      <button class="scn-tool" type="button" data-t="rot">${I.rot}<span>סיבוב</span></button>
      <button class="scn-tool" type="button" data-t="ink" aria-pressed="false">${I.erase}<span>מחיקת עט</span></button>
      <button class="scn-tool" type="button" data-t="all">${I.ok}<span>לכל העמודים</span></button>
      <button class="scn-tool" type="button" data-t="del">${I.trash}<span>מחיקה</span></button>
    </div>
  </section>
  <div class="scn-busy" hidden><div><i></i><span></span></div></div>
  <p class="scn-toast" role="status" hidden></p>
</div>`);
document.body.append(root);
const galIn = h('<input type="file" accept="image/*" multiple hidden>');
const camIn = h('<input type="file" accept="image/*" capture="environment" hidden>');
root.append(galIn, camIn);
const $r = s => root.querySelector(s);
const V = { cam: $r('[data-v=cam]'), review: $r('[data-v=review]'), edit: $r('[data-v=edit]') };

/* ---------------------------------------------------------------- תמונות */
async function decodeImg(src) {
  try { return await createImageBitmap(src, { imageOrientation: 'from-image' }); }
  catch (_) {
    return await new Promise((res, rej) => { const u = URL.createObjectURL(src), im = new Image();
      im.onload = () => { URL.revokeObjectURL(u); res(im); }; im.onerror = e => { URL.revokeObjectURL(u); rej(e); }; im.src = u; });
  }
}
const dimsOf = b => [b.width || b.naturalWidth, b.height || b.naturalHeight];
function canvasOf(w, hh) { const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(w)); c.height = Math.max(1, Math.round(hh)); return c; }
function scaledCanvas(img, max) {
  const [w, hh] = dimsOf(img), k = Math.min(1, max / Math.max(w, hh)), c = canvasOf(w * k, hh * k);
  const x = c.getContext('2d'); x.imageSmoothingQuality = 'high'; x.drawImage(img, 0, 0, c.width, c.height); return c;
}
const toBlob = (c, q = 0.9, type = 'image/jpeg') => new Promise((res, rej) => c.toBlob(b => b ? res(b) : rej(new Error('התמונה לא נשמרה')), type, q));
const free = c => { if (c) { c.width = c.height = 0; } };

/* ---- זיהוי הדף בתמונה, במכשיר: הדף בהיר מהרקע. מוצאים את האזור הבהיר הגדול ואת ארבע פינותיו ---- */
function boxBlur(a, w, hh, r) {
  const t = new Float32Array(a.length);
  for (let y = 0; y < hh; y++) { let s = 0; const o = y * w;
    for (let x = -r; x <= r; x++) s += a[o + Math.min(w - 1, Math.max(0, x))];
    for (let x = 0; x < w; x++) { t[o + x] = s / (2 * r + 1); s += a[o + Math.min(w - 1, x + r + 1)] - a[o + Math.max(0, x - r)]; } }
  for (let x = 0; x < w; x++) { let s = 0;
    for (let y = -r; y <= r; y++) s += t[Math.min(hh - 1, Math.max(0, y)) * w + x];
    for (let y = 0; y < hh; y++) { a[y * w + x] = s / (2 * r + 1); s += t[Math.min(hh - 1, y + r + 1) * w + x] - t[Math.max(0, y - r) * w + x]; } }
}
function findSheet(data, w, hh) {
  const n = w * hh, g = new Float32Array(n);
  for (let i = 0; i < n; i++) g[i] = data[i * 4] * 0.3 + data[i * 4 + 1] * 0.59 + data[i * 4 + 2] * 0.11;
  boxBlur(g, w, hh, 2);
  // Otsu: הסף שמפריד הכי טוב בין הבהיר לכהה
  const hist = new Float64Array(256); for (let i = 0; i < n; i++) hist[g[i] | 0]++;
  let sum = 0; for (let t = 0; t < 256; t++) sum += t * hist[t];
  let wB = 0, sB = 0, best = -1, th = 128, m0 = 0, m1 = 0;
  for (let t = 0; t < 256; t++) {
    wB += hist[t]; if (!wB) continue; const wF = n - wB; if (!wF) break;
    sB += t * hist[t]; const mB = sB / wB, mF = (sum - sB) / wF, v = wB * wF * (mB - mF) * (mB - mF);
    if (v > best) { best = v; th = t; m0 = mB; m1 = mF; }
  }
  if (m1 - m0 < 32) return null;
  const mask = new Uint8Array(n); for (let i = 0; i < n; i++) mask[i] = g[i] > th ? 1 : 0;
  // האזור הבהיר הגדול ביותר
  const seen = new Uint8Array(n), stack = new Int32Array(n);
  let bestC = null;
  for (let s = 0; s < n; s++) {
    if (!mask[s] || seen[s]) continue;
    let sp = 0, area = 0, mnS = 1e9, mxS = -1e9, mnD = 1e9, mxD = -1e9, pS = 0, PS = 0, pD = 0, PD = 0;
    stack[sp++] = s; seen[s] = 1;
    while (sp) {
      const p = stack[--sp], x = p % w, y = (p / w) | 0; area++;
      const su = x + y, di = x - y;
      if (su < mnS) { mnS = su; pS = p; } if (su > mxS) { mxS = su; PS = p; }
      if (di < mnD) { mnD = di; pD = p; } if (di > mxD) { mxD = di; PD = p; }
      if (x > 0 && mask[p - 1] && !seen[p - 1]) { seen[p - 1] = 1; stack[sp++] = p - 1; }
      if (x < w - 1 && mask[p + 1] && !seen[p + 1]) { seen[p + 1] = 1; stack[sp++] = p + 1; }
      if (y > 0 && mask[p - w] && !seen[p - w]) { seen[p - w] = 1; stack[sp++] = p - w; }
      if (y < hh - 1 && mask[p + w] && !seen[p + w]) { seen[p + w] = 1; stack[sp++] = p + w; }
    }
    if (!bestC || area > bestC.area) bestC = { area, tl: pS, br: PS, bl: pD, tr: PD };
  }
  if (!bestC) return null;
  const frac = bestC.area / n;
  if (frac < 0.1 || frac > 0.97) return null;
  const pt = p => [((p % w) + 0.5) / w, (((p / w) | 0) + 0.5) / hh];
  const q = [pt(bestC.tl), pt(bestC.tr), pt(bestC.br), pt(bestC.bl)];
  const qa = quadArea(q);
  if (qa < 0.08 || qa / frac < 0.8 || qa / frac > 1.18) return null;   // לא דומה למלבן
  // ארבע הפינות על שולי התמונה: זה הרקע, לא דף
  const onEdge = q.filter(([x, y]) => x < 0.02 || x > 0.98 || y < 0.02 || y > 0.98).length;
  if (onEdge >= 4) return null;
  // מעט פנימה, כדי שקצה השולחן לא ייכנס לדף
  const cx = q.reduce((a, p) => a + p[0], 0) / 4, cy = q.reduce((a, p) => a + p[1], 0) / 4;
  return q.map(([x, y]) => [Math.max(0, Math.min(1, cx + (x - cx) * 0.996)), Math.max(0, Math.min(1, cy + (y - cy) * 0.996))]);
}
function sheetIn(img) {
  const c = scaledCanvas(img, 320), x = c.getContext('2d', { willReadFrequently: true });
  const q = findSheet(x.getImageData(0, 0, c.width, c.height).data, c.width, c.height); free(c); return q;
}

/* ---- שני עמודים בדף שזוהה (מחברת או ספר פתוחים)? ----
   הספירלה או התפר: עמודה צרה באמצע עם הרבה נקודות כהות. בלי ספירלה: דף רחב בהרבה מגובהו */
function spreadOf(c, q) {
  const [W0, H0] = dimsOf(c), P = q.map(([x, y]) => [x * W0, y * H0]), dist = (a, b) => Math.hypot(a[0] - b[0], a[1] - b[1]);
  const aspect = (dist(P[0], P[1]) + dist(P[3], P[2])) / Math.max(1, dist(P[0], P[3]) + dist(P[1], P[2]));
  if (aspect < 0.95) return null;                       // דף לאורך: עמוד אחד
  const w = warp(c, q, 520), W = w.width, H = w.height, d = w.getContext('2d', { willReadFrequently: true }).getImageData(0, 0, W, H).data; free(w);
  // כתב וספירלה = נקודות כהות בהרבה מהסביבה הקרובה שלהן. צל של יד או גוף כהה בהדרגה, ולכן לא נספר
  const g = new Float32Array(W * H);
  for (let i = 0, k = 0; k < g.length; i += 4, k++) g[k] = Math.min(d[i], d[i + 1], d[i + 2]);
  const bg = Float32Array.from(g); boxBlur(bg, W, H, Math.max(4, Math.round(W / 40)));
  const ink = new Uint8Array(W * H);
  for (let k = 0; k < g.length; k++) if (bg[k] - g[k] > Math.max(18, bg[k] * 0.1)) ink[k] = 1;
  const col = new Float32Array(W);
  for (let y = Math.round(H * 0.04); y < H * 0.96; y++) for (let x = 0; x < W; x++) col[x] += ink[y * W + x];
  for (let x = 0; x < W; x++) col[x] /= H;
  const sm = x => (col[Math.max(0, x - 1)] + col[x] + col[Math.min(W - 1, x + 1)]) / 3;
  const sorted = [...col].sort((a, b) => a - b), med = sorted[W >> 1];
  let best = -1, bx = 0;
  for (let x = Math.round(W * 0.3); x <= W * 0.7; x++) { const v = sm(x); if (v > best) { best = v; bx = x; } }
  const spiral = best > 0.12 && best > med * 4 + 0.05;
  let t = null;
  if (spiral) t = bx / W;                                // ספירלה או תפר
  else if (aspect > 1.25) t = 0.5;                       // רחב: כנראה שני עמודים
  if (t == null) return null;
  const gap = 0.035, cnt = { left: 0, right: 0 };
  for (let y = Math.round(H * 0.05); y < H * 0.95; y++) for (let x = Math.round(W * 0.05); x < W * 0.95; x++) {
    const u = x / W; if (Math.abs(u - t) < gap + 0.03) continue;
    if (ink[y * W + x]) cnt[u > t ? 'right' : 'left']++;
  }
  const lerp = (a, b, k) => [a[0] + (b[0] - a[0]) * k, a[1] + (b[1] - a[1]) * k];
  const [TL, TR, BR, BL] = q;
  return { t, ink: cnt, spiral,
    left: [TL, lerp(TL, TR, t - gap), lerp(BL, BR, t - gap), BL],
    right: [lerp(TL, TR, t + gap), TR, BR, lerp(BL, BR, t + gap)] };
}

/* ---- מגה סריקה: מציאת התזוזה בין שתי תמונות (phase correlation) ---- */
function fft1(re, im, inv) {
  const n = re.length;
  for (let i = 1, j = 0; i < n; i++) { let bit = n >> 1; for (; j & bit; bit >>= 1) j ^= bit; j ^= bit; if (i < j) { let t = re[i]; re[i] = re[j]; re[j] = t; t = im[i]; im[i] = im[j]; im[j] = t; } }
  for (let len = 2; len <= n; len <<= 1) {
    const ang = 2 * Math.PI / len * (inv ? 1 : -1), wr = Math.cos(ang), wi = Math.sin(ang);
    for (let i = 0; i < n; i += len) {
      let cr = 1, ci = 0;
      for (let k = 0; k < len / 2; k++) {
        const a = i + k, b = a + len / 2, xr = re[b] * cr - im[b] * ci, xi = re[b] * ci + im[b] * cr;
        re[b] = re[a] - xr; im[b] = im[a] - xi; re[a] += xr; im[a] += xi;
        const t = cr * wr - ci * wi; ci = cr * wi + ci * wr; cr = t;
      }
    }
  }
  if (inv) for (let i = 0; i < n; i++) { re[i] /= n; im[i] /= n; }
}
function fft2(re, im, N, inv) {
  const r = new Float64Array(N), i2 = new Float64Array(N);
  for (let y = 0; y < N; y++) { const o = y * N; for (let x = 0; x < N; x++) { r[x] = re[o + x]; i2[x] = im[o + x]; } fft1(r, i2, inv); for (let x = 0; x < N; x++) { re[o + x] = r[x]; im[o + x] = i2[x]; } }
  for (let x = 0; x < N; x++) { for (let y = 0; y < N; y++) { r[y] = re[y * N + x]; i2[y] = im[y * N + x]; } fft1(r, i2, inv); for (let y = 0; y < N; y++) { re[y * N + x] = r[y]; im[y * N + x] = i2[y]; } }
}
// spectrum of a grey N×N picture (with a soft window, so the picture edges do not look like a shift)
function spectrum(g, N) {
  const re = new Float64Array(N * N), im = new Float64Array(N * N);
  let m = 0; for (let i = 0; i < N * N; i++) m += g[i]; m /= N * N;
  for (let y = 0; y < N; y++) { const wy = 0.5 - 0.5 * Math.cos(2 * Math.PI * y / (N - 1));
    for (let x = 0; x < N; x++) re[y * N + x] = (g[y * N + x] - m) * wy * (0.5 - 0.5 * Math.cos(2 * Math.PI * x / (N - 1))); }
  fft2(re, im, N, false);
  return { re, im };
}
// how far picture B moved relative to picture A, in pixels of the N×N grid, and how sure
function phaseShift(A, B, N) {
  const re = new Float64Array(N * N), im = new Float64Array(N * N);
  for (let i = 0; i < N * N; i++) {
    const r = A.re[i] * B.re[i] + A.im[i] * B.im[i], q = A.im[i] * B.re[i] - A.re[i] * B.im[i], m = Math.hypot(r, q) || 1;
    re[i] = r / m; im[i] = q / m;
  }
  fft2(re, im, N, true);
  let best = -1e9, bi = 0, sum = 0;
  for (let i = 0; i < N * N; i++) { sum += Math.abs(re[i]); if (re[i] > best) { best = re[i]; bi = i; } }
  let dx = bi % N, dy = (bi / N) | 0;
  if (dx > N / 2) dx -= N; if (dy > N / 2) dy -= N;
  return { dx, dy, conf: best / (sum / (N * N)) / (N * N) * 100, peak: best };
}

/* ---- יישור: מסגרת של ארבע פינות → דף ישר (מיפוי ריבוע→מרובע של Heckbert) ---- */
function warp(img, quad, max) {
  const [W0, H0] = dimsOf(img);
  if (isFull(quad)) return scaledCanvas(img, max);
  const q = quad.map(([x, y]) => [x * W0, y * H0]);
  const dist = (a, b) => Math.hypot(a[0] - b[0], a[1] - b[1]);
  let W = Math.max(dist(q[0], q[1]), dist(q[3], q[2])), H = Math.max(dist(q[0], q[3]), dist(q[1], q[2]));
  const k = Math.min(1, max / Math.max(W, H)); W = Math.max(8, Math.round(W * k)); H = Math.max(8, Math.round(H * k));
  // המקור מוקטן קודם, כדי שלא ייקרא יותר ממה שצריך
  const sk = Math.min(1, Math.max(W, H) * 1.35 / Math.max(W0, H0));
  const src = sk < 1 ? scaledCanvas(img, Math.max(W0, H0) * sk) : scaledCanvas(img, Math.max(W0, H0));
  const sw = src.width, sh = src.height, fx0 = sw / W0, fy0 = sh / H0;
  const S = src.getContext('2d', { willReadFrequently: true }).getImageData(0, 0, sw, sh).data; free(src);
  const [[x0, y0], [x1, y1], [x2, y2], [x3, y3]] = q.map(([x, y]) => [x * fx0, y * fy0]);
  const dx1 = x1 - x2, dx2 = x3 - x2, dx3 = x0 - x1 + x2 - x3, dy1 = y1 - y2, dy2 = y3 - y2, dy3 = y0 - y1 + y2 - y3;
  let a, b, c, d, e, f, g = 0, hh = 0;
  if (Math.abs(dx3) < 1e-9 && Math.abs(dy3) < 1e-9) { a = x1 - x0; b = x2 - x1; c = x0; d = y1 - y0; e = y2 - y1; f = y0; }
  else { const det = dx1 * dy2 - dx2 * dy1; g = (dx3 * dy2 - dx2 * dy3) / det; hh = (dx1 * dy3 - dx3 * dy1) / det;
    a = x1 - x0 + g * x1; b = x3 - x0 + hh * x3; c = x0; d = y1 - y0 + g * y1; e = y3 - y0 + hh * y3; f = y0; }
  const out = canvasOf(W, H), ox = out.getContext('2d'), im = ox.createImageData(W, H), D = im.data;
  for (let Y = 0; Y < H; Y++) {
    const v = (Y + 0.5) / H;
    for (let X = 0; X < W; X++) {
      const u = (X + 0.5) / W, z = g * u + hh * v + 1, px = (a * u + b * v + c) / z - 0.5, py = (d * u + e * v + f) / z - 0.5;
      const ix = Math.max(0, Math.min(sw - 2, px | 0)), iy = Math.max(0, Math.min(sh - 2, py | 0)), fx = Math.max(0, Math.min(1, px - ix)), fy = Math.max(0, Math.min(1, py - iy));
      const i00 = (iy * sw + ix) * 4, i10 = i00 + 4, i01 = i00 + sw * 4, i11 = i01 + 4, o = (Y * W + X) * 4;
      for (let ch = 0; ch < 3; ch++) D[o + ch] = (S[i00 + ch] * (1 - fx) + S[i10 + ch] * fx) * (1 - fy) + (S[i01 + ch] * (1 - fx) + S[i11 + ch] * fx) * fy;
      D[o + 3] = 255;
    }
  }
  ox.putImageData(im, 0, 0);
  return out;
}
function rotated(c, rot) {
  rot = ((rot % 360) + 360) % 360; if (!rot) return c;
  const o = rot % 180 ? canvasOf(c.height, c.width) : canvasOf(c.width, c.height), x = o.getContext('2d');
  x.translate(o.width / 2, o.height / 2); x.rotate(rot * Math.PI / 180); x.drawImage(c, -c.width / 2, -c.height / 2);
  free(c); return o;
}

/* ---- מסננים: מוצאים את צבע הנייר בכל מקום בדף (כולל צללים) ומחלקים בו — הנייר יוצא לבן והכתב חד ---- */
function paperMap(c) {
  const W = c.width, H = c.height, k = Math.max(6, Math.round(Math.max(W, H) / 72));
  const sw = Math.max(4, Math.round(W / k)), sh = Math.max(4, Math.round(H / k));
  const s = canvasOf(sw, sh), sx = s.getContext('2d', { willReadFrequently: true });
  sx.imageSmoothingQuality = 'high'; sx.drawImage(c, 0, 0, sw, sh);
  const im = sx.getImageData(0, 0, sw, sh), d = im.data, n = sw * sh;
  // הרחבה של הבהיר (מוחקת את הכתב), ואחר כך טשטוש
  for (let pass = 0; pass < 2; pass++) {
    const src = new Uint8ClampedArray(d);
    for (let y = 0; y < sh; y++) for (let x = 0; x < sw; x++) {
      const o = (y * sw + x) * 4; let r = 0, g = 0, b = 0;
      for (let yy = Math.max(0, y - 1); yy <= Math.min(sh - 1, y + 1); yy++) for (let xx = Math.max(0, x - 1); xx <= Math.min(sw - 1, x + 1); xx++) {
        const p = (yy * sw + xx) * 4, l = src[p] + src[p + 1] + src[p + 2];
        if (l > r + g + b) { r = src[p]; g = src[p + 1]; b = src[p + 2]; }
      }
      d[o] = r; d[o + 1] = g; d[o + 2] = b;
    }
  }
  for (let ch = 0; ch < 3; ch++) { const a = new Float32Array(n); for (let i = 0; i < n; i++) a[i] = d[i * 4 + ch]; boxBlur(a, sw, sh, 1); for (let i = 0; i < n; i++) d[i * 4 + ch] = a[i]; }
  sx.putImageData(im, 0, 0);
  const big = canvasOf(W, H), bx = big.getContext('2d', { willReadFrequently: true });
  bx.imageSmoothingQuality = 'high'; bx.drawImage(s, 0, 0, W, H); free(s);
  const out = bx.getImageData(0, 0, W, H).data; free(big); return out;
}
function applyFilter(c, f, ink) {
  if (f === 'orig' && !ink) return c;
  const W = c.width, H = c.height, x = c.getContext('2d', { willReadFrequently: true }), im = x.getImageData(0, 0, W, H), d = im.data, n = W * H;
  const bg = f === 'orig' ? null : paperMap(c);
  // הנייר יוצא לבן לגמרי: מחלקים בצבע הנייר, מעט מעל הלבן, ונחתך
  if (bg) for (let i = 0; i < d.length; i += 4) {
    d[i] = Math.min(255, d[i] * 268 / Math.max(40, bg[i])); d[i + 1] = Math.min(255, d[i + 1] * 268 / Math.max(40, bg[i + 1])); d[i + 2] = Math.min(255, d[i + 2] * 268 / Math.max(40, bg[i + 2]));
  }
  if (ink) {
    // עט כחול או אדום: צבע רווי. הכתב המודפס (שחור, אפור) נשאר. גם השוליים הרכים של הקו נמחקים
    const strong = new Float32Array(n), weak = new Uint8Array(n);
    for (let p = 0, i = 0; p < n; p++, i += 4) {
      const r = d[i], g = d[i + 1], b = d[i + 2], sat = Math.max(r, g, b) - Math.min(r, g, b);
      const blue = b > r + 12 && b >= g - 10, red = r > g + 22 && r > b + 14;
      if (sat > 42 && ((b > r + 22 && b >= g - 6) || (r > g + 40 && r > b + 28))) strong[p] = 1;
      if ((blue || red) && sat > 12) weak[p] = 1;
    }
    boxBlur(strong, W, H, 2);
    for (let p = 0, i = 0; p < n; p++, i += 4) {
      const L = d[i] * 0.3 + d[i + 1] * 0.59 + d[i + 2] * 0.11;
      if (strong[p] > 0.02 && (weak[p] || L > 140)) d[i] = d[i + 1] = d[i + 2] = 255;
    }
  }
  if (f !== 'orig') {
    const dark = { magic: 1.32, clean: 1.12, gray: 1.4, bw: 1, board: 1.6 }[f] || 1, sat = { magic: 1.28, board: 1.5 }[f] || 1;
    for (let i = 0; i < d.length; i += 4) {
      let r = d[i], g = d[i + 1], b = d[i + 2];
      if (f === 'gray' || f === 'bw') {
        let L = r * 0.3 + g * 0.59 + b * 0.11;
        if (f === 'gray') L = 255 - (255 - L) * dark;
        else L = L > 212 ? 255 : L < 150 ? 0 : (L - 150) * 4.1;
        d[i] = d[i + 1] = d[i + 2] = L;
      } else {
        r = 255 - (255 - r) * dark; g = 255 - (255 - g) * dark; b = 255 - (255 - b) * dark;
        if (sat !== 1) { const L = r * 0.3 + g * 0.59 + b * 0.11; r = L + (r - L) * sat; g = L + (g - L) * sat; b = L + (b - L) * sat; }
        d[i] = r; d[i + 1] = g; d[i + 2] = b;
      }
    }
  }
  x.putImageData(im, 0, 0);
  return c;
}

/* ---- עמוד: התמונה המקורית + המסגרת + סיבוב + מסנן. הכל נשמר, ואפשר לשנות כל דבר אחר כך ---- */
let pid = 0;
function newPage(o) { return Object.assign({ id: ++pid, quad: FULL.map(p => [...p]), rot: 0, filter: CAM.defFilter, ink: false, thumb: '', ver: 0 }, o); }
async function render(p, max, { filter = true } = {}) {
  const img = await decodeImg(p.src);
  let c = warp(img, p.quad, max);
  if (img.close) img.close();
  c = rotated(c, p.rot);
  if (filter) applyFilter(c, p.filter, p.ink);
  return c;
}
async function refreshThumb(p) {
  const v = ++p.ver, c = await render(p, 520);
  if (v !== p.ver) { free(c); return; }
  p.thumb = c.toDataURL('image/jpeg', 0.82); p.tw = c.width; p.th = c.height; free(c);
}

/* ---------------------------------------------------------------- הסורק */
const CAM = {
  pages: [], mode: 'doc', auto: true, defFilter: 'magic', stream: null, track: null, torch: false,
  quad: null, stable: 0, lastShot: 0, lastShotQuad: null, timer: 0, shooting: false, idFront: null, view: 'cam', cur: 0, cropping: false,
  qrOn: false, qrDet: null, name: '', opts: { size: 'auto', q: 'high', wm: '' },
  video: $r('video'), vf: $r('.scn-vf'), poly: $r('.ov polygon.main'), poly2: $r('.ov polygon.other'), pick: 'auto', spread: null, hint: $r('.scn-hint'), ring: $r('.ring circle'),

  open() {
    if (typeof needReg === 'function' && needReg()) { if (typeof showReg === 'function') showReg(); return; }
    this.prevFocus = document.activeElement;
    root.hidden = false; document.documentElement.style.overflow = 'hidden';
    if (!this.pages.length) this.name = defaultName();
    this.setMode(this.mode === 'qr' ? 'doc' : this.mode);
    this.show('cam');
    setTimeout(() => $r('[data-v=cam] [data-a=close]').focus(), 50);
  },
  close(force) {
    if (!force && this.pages.length && !confirm('לסגור את הסורק? העמודים שסרקתם לא יישמרו')) return;
    this.stopCam(); root.hidden = true; document.documentElement.style.overflow = '';
    this.pages = []; this.idFront = null; this.cropping = false; this.updateStack();
    try { this.prevFocus && this.prevFocus.focus(); } catch (_) {}
  },
  show(v) {
    this.view = v;
    for (const k in V) V[k].hidden = k !== v;
    if (v === 'cam') this.startCam(); else this.stopCam();
    if (v === 'review') this.renderReview();
    if (v === 'edit') this.renderEdit();
  },
  busy(msg) { const b = $r('.scn-busy'); b.hidden = !msg; if (msg) b.querySelector('span').textContent = msg; },

  /* ---- מצלמה ---- */
  async startCam() {
    if (this.stream) { this.loop(); return; }
    const nc = $r('.scn-nocam');
    nc.hidden = true;
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) return this.noCam('הדפדפן הזה לא נותן גישה ישירה למצלמה.');
    this.setHint('פותח את המצלמה…');
    try {
      this.stream = await navigator.mediaDevices.getUserMedia({ audio: false, video: { facingMode: { ideal: 'environment' }, width: { ideal: 3840 }, height: { ideal: 2160 } } });
    } catch (e) {
      try { this.stream = await navigator.mediaDevices.getUserMedia({ audio: false, video: { facingMode: { ideal: 'environment' } } }); }
      catch (e2) { return this.noCam(/NotAllowed|Permission|Security/i.test(e2.name + e2.message) ? 'לא ניתנה הרשאה למצלמה. אפשר לאשר אותה בהגדרות הדפדפן, או לצלם בדרך אחרת:' : 'לא נמצאה מצלמה במכשיר הזה.'); }
    }
    if (root.hidden || this.view !== 'cam') { this.stopCam(); return; }
    this.track = this.stream.getVideoTracks()[0];
    const v = this.video; v.srcObject = this.stream;
    await new Promise(r => { if (v.readyState >= 1) r(); else v.onloadedmetadata = () => r(); });
    try { await v.play(); } catch (_) {}
    this.vf.hidden = false; this.fit(); setTimeout(() => this.fit(), 120);
    let caps = {}; try { caps = this.track.getCapabilities ? this.track.getCapabilities() : {}; } catch (_) {}
    $r('[data-a=torch]').hidden = !caps.torch;
    // מיקוד רציף, אם יש
    try { if (caps.focusMode && caps.focusMode.includes('continuous')) await this.track.applyConstraints({ advanced: [{ focusMode: 'continuous' }] }); } catch (_) {}
    this.setHint(this.modeHint());
    this.loop();
  },
  stopCam() {
    clearTimeout(this.timer); this.timer = 0;
    if (this.mega) this.megaCancel();
    if (this.stream) { this.stream.getTracks().forEach(t => t.stop()); this.stream = null; this.track = null; this.torch = false; $r('[data-a=torch]').setAttribute('aria-pressed', 'false'); }
    this.video.srcObject = null; this.quad = null; this.stable = 0; this.ic = null; this.spread = null; this.pick = 'auto'; this.drawQuad();
  },
  noCam(msg) {
    this.vf.hidden = true; this.setHint('');
    const nc = $r('.scn-nocam'); nc.hidden = false;
    nc.innerHTML = `<h3>המצלמה לא זמינה</h3><p>${esc(msg)}</p><div class="row">
      <button class="scn-btn main" type="button" data-n="cam">${I.auto} צילום במצלמה של המכשיר</button>
      <button class="scn-btn" type="button" data-n="gal">${I.gallery} בחירת תמונות</button></div>`;
    nc.querySelector('[data-n=cam]').onclick = () => camIn.click();
    nc.querySelector('[data-n=gal]').onclick = () => galIn.click();
  },
  fit() {
    const v = this.video, st = $r('.scn-stage'); if (!v.videoWidth) return;
    const ar = v.videoWidth / v.videoHeight, W = st.clientWidth - 16, H = st.clientHeight - 16;
    let w = W, hh = W / ar; if (hh > H) { hh = H; w = H * ar; }
    this.vf.style.width = Math.round(w) + 'px'; this.vf.style.height = Math.round(hh) + 'px';
  },
  async toggleTorch() {
    if (!this.track) return;
    this.torch = !this.torch;
    try { await this.track.applyConstraints({ advanced: [{ torch: this.torch }] }); } catch (_) { this.torch = false; }
    $r('[data-a=torch]').setAttribute('aria-pressed', String(this.torch));
  },
  setHint(t) { this.hint.textContent = t; this.hint.hidden = !t; },
  modeHint() {
    return { doc: 'כוונו את המצלמה אל הדף', book: 'ספר פתוח: שני העמודים בתוך המסגרת, התפר באמצע', id: this.idFront ? 'עכשיו הצד האחורי של התעודה' : 'הצד הקדמי של התעודה, בתוך המסגרת', board: 'כוונו אל הלוח', qr: 'כוונו אל הקוד', mega: this.mega && this.mega.on ? 'הזיזו לאט מעל כל הדף, באותו גובה' : 'דף גדול: לחצו על הכפתור, והזיזו את הטלפון לאט מעל כל הדף' }[this.mode];
  },
  setMode(m) {
    this.mode = m;
    root.querySelectorAll('.scn-modes [data-m]').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.m === m)));
    const gd = $r('.scn-vf .guide');
    gd.innerHTML = m === 'id' ? '<div class="idbox"></div>' : m === 'book' ? '<div class="bookline"></div>' : m === 'qr' ? '<div class="qrbox"></div>' : '';
    $r('[data-a=shoot]').style.visibility = m === 'qr' ? 'hidden' : ''; $r('[data-a=auto]').hidden = m === 'qr' || m === 'id' || m === 'mega';
    if (m !== 'mega' && this.mega) this.megaCancel();
    if (m !== 'id' && this.idFront) { free(this.idFront); this.idFront = null; }
    this.quad = null; this.stable = 0; this.drawQuad(); this.sheet(null);
    this.setHint(this.modeHint());
    if (m === 'qr') this.prepQr();
    this.updateStack();
  },
  drawQuad() {
    const q = this.quad;
    this.poly.setAttribute('points', q && this.mode !== 'qr' ? q.map(([x, y]) => `${x * 100},${y * 100}`).join(' ') : '');
    const o = this.spread && q && this.mode !== 'qr' && this.pick !== 'both' ? this.spread.other : null;
    this.poly2.setAttribute('points', o ? o.map(([x, y]) => `${x * 100},${y * 100}`).join(' ') : '');
    const pb = $r('[data-a=pages]'); pb.hidden = !this.spread; $r('[data-a=swap]').hidden = !this.spread || this.pick === 'both'; if (this.spread) pb.querySelector('span').textContent = this.pick === 'both' ? (this.spread.spiral ? 'שני העמודים' : 'כל הדף') : 'עמוד אחד';
    this.poly.classList.toggle('ok', this.stable >= 3);
    const need = 8, p = this.auto && this.mode !== 'id' && q ? Math.min(1, this.stable / need) : 0;
    this.ring.style.strokeDashoffset = String(264 * (1 - p));
  },
  loop() {
    clearTimeout(this.timer);
    if (!this.stream || this.view !== 'cam' || root.hidden) return;
    const v = this.video;
    if (v.videoWidth && !this.shooting && !$r('.scn-sheet:not([hidden])')) {
      try {
        if (this.mode === 'qr') this.scanQr().catch(() => {});
        else if (this.mode === 'mega') { if (this.mega && this.mega.on) this.megaStep(); }
        else {
          const k = 320 / Math.max(v.videoWidth, v.videoHeight), c = this.small || (this.small = document.createElement('canvas'));
          c.width = Math.round(v.videoWidth * k); c.height = Math.round(v.videoHeight * k);
          const x = c.getContext('2d', { willReadFrequently: true }); x.drawImage(v, 0, 0, c.width, c.height);
          let q = findSheet(x.getImageData(0, 0, c.width, c.height).data, c.width, c.height);
          if (q && (this.mode === 'doc' || this.mode === 'board')) {
            // לבדיקת הכתב בכל עמוד: תמונה גדולה יותר מזו של זיהוי הדף, כי כתב קטן נעלם בהקטנה
            const k2 = 900 / Math.max(v.videoWidth, v.videoHeight), m = this.mid || (this.mid = document.createElement('canvas'));
            m.width = Math.round(v.videoWidth * k2); m.height = Math.round(v.videoHeight * k2);
            m.getContext('2d', { willReadFrequently: true }).drawImage(v, 0, 0, m.width, m.height);
            q = this.choosePage(m, q);
          } else q = this.choosePage(c, q);
          this.track2(q);
        }
      } catch (e) { console.warn(e); }
    }
    this.timer = setTimeout(() => this.loop(), this.mode === 'qr' ? 220 : this.mode === 'mega' ? 160 : 110);
  },
  // מעקב אחרי הדף בין תמונות: החלקה, יציבות, וצילום אוטומטי
  track2(q) {
    if (!q) { this.miss = (this.miss || 0) + 1; if (this.miss > 4) { this.quad = null; this.stable = 0; } this.drawQuad();
      if (!this.quad) this.setHint(this.modeHint()); return; }
    this.miss = 0;
    const prev = this.quad;
    const move = prev ? Math.max(...q.map((p, i) => Math.hypot(p[0] - prev[i][0], p[1] - prev[i][1]))) : 1;
    this.quad = prev && move < 0.08 ? q.map((p, i) => [prev[i][0] * 0.45 + p[0] * 0.55, prev[i][1] * 0.45 + p[1] * 0.55]) : q;
    this.stable = move < 0.014 ? this.stable + 1 : 0;
    this.drawQuad();
    const same = this.lastShotQuad && Math.max(...q.map((p, i) => Math.hypot(p[0] - this.lastShotQuad[i][0], p[1] - this.lastShotQuad[i][1]))) < 0.03;
    if (this.mode === 'id') { this.setHint(this.modeHint()); return; }
    if (quadArea(q) < 0.07) { this.stable = 0; this.drawQuad(); this.setHint('קרבו את המצלמה לדף, כדי שהכתב ייקרא'); return; }
    if (this.spread && this.pick !== 'both' && !this.spreadTold) { this.setHint('נסרק העמוד המסומן · נגיעה בשני מחליפה'); this.spreadTold = Date.now(); return; }
    if (this.spreadTold && Date.now() - this.spreadTold < 2200) return;
    if (!this.auto) { this.setHint('הדף זוהה. לוחצים על הכפתור לצילום'); return; }
    if (same && Date.now() - this.lastShot < 6000) { this.setHint('צולם ✓ החליפו דף או הפכו עמוד'); return; }
    this.setHint(this.stable >= 3 ? 'החזיקו יציב…' : 'הדף זוהה');
    if (this.stable >= 8 && Date.now() - this.lastShot > 1500) this.shoot();
  },
  // צילום אמיתי במצלמה, ברזולוציה המלאה שלה. המסגרת שזוהתה בווידאו מועברת לתמונה:
  // הווידאו הוא בדרך כלל חיתוך מהמרכז של אותה תמונה (16:9 מתוך 4:3)
  async photo(W, H, q) {
    if (!this.track || !('ImageCapture' in window) || this.noPhoto) return null;
    try {
      const ic = this.ic || (this.ic = new ImageCapture(this.track));
      const pb = await Promise.race([ic.takePhoto(), sleep(5000).then(() => null)]);
      if (!pb) { this.noPhoto = true; return null; }
      const bmp = await decodeImg(pb), [pw, ph] = dimsOf(bmp);
      if ((pw > ph) !== (W > H)) { if (bmp.close) bmp.close(); this.noPhoto = true; return null; }   // מסובבת אחרת מהווידאו: לא בטוח
      const av = W / H, ap = pw / ph;
      const map = ([x, y]) => av > ap ? [x, 0.5 + (y - 0.5) * ap / av] : [0.5 + (x - 0.5) * av / ap, y];
      const c = scaledCanvas(bmp, 4000); if (bmp.close) bmp.close();
      return { c, q: q ? q.map(map) : null };
    } catch (e) { console.warn('takePhoto', e); this.noPhoto = true; return null; }
  },
  // מחברת או ספר פתוחים: הדף שזוהה הוא שני עמודים. מוצאים את הספירלה או את האמצע,
  // ובוחרים את העמוד שיש בו כתב (או את מה שהמשתמש בחר בנגיעה)
  choosePage(c, q) {
    this.spread = null;
    if (!q || (this.mode !== 'doc' && this.mode !== 'board')) return q;
    const s = spreadOf(c, q);
    if (!s) return q;
    // העמוד עם יותר כתב. כשההפרש קטן נשאר העמוד הקודם, כדי שהמסגרת לא תקפוץ
    const L = s.ink.left, Rr = s.ink.right, clear = Math.max(L, Rr) > 60 && Math.max(L, Rr) > Math.min(L, Rr) * 1.3;
    let side = this.pick === 'left' || this.pick === 'right' ? this.pick : clear ? (L > Rr ? 'left' : 'right') : (this.autoSide || (L > Rr ? 'left' : 'right'));
    if (this.pick === 'auto') this.autoSide = side;
    const qq = side === 'right' ? s.right : s.left;
    this.spread = { t: s.t, full: q, other: side === 'right' ? s.left : s.right, side };
    return this.pick === 'both' ? q : qq;
  },
  tapPick(e) {
    if (!this.spread || this.mode === 'book') return;
    const r = this.vf.getBoundingClientRect(), x = (e.clientX - r.left) / r.width, y = (e.clientY - r.top) / r.height;
    const f = this.spread.full, t = this.spread.t, top = f[0][0] + (f[1][0] - f[0][0]) * t, bot = f[3][0] + (f[2][0] - f[3][0]) * t;
    const mid = top + (bot - top) * Math.max(0, Math.min(1, (y - Math.min(f[0][1], f[1][1])) / Math.max(0.01, Math.max(f[2][1], f[3][1]) - Math.min(f[0][1], f[1][1]))));
    this.pick = x > mid ? 'right' : 'left'; this.quad = null; this.stable = 0; this.spreadTold = 0;
    this.setHint(this.pick === 'right' ? 'נבחר העמוד הימני' : 'נבחר העמוד השמאלי');
  },
  /* ---- מגה סריקה: דף גדול מ-A4 (שרטוט, פוסטר, מפה) נסרק בתנועה רציפה ----
     כל כמה עשיריות שנייה נמדדת התזוזה של התמונה מהחלק האחרון שנשמר (phase correlation, במכשיר).
     כשהטלפון זז כרבע מסך נשמר חלק חדש במקומו, ובסוף כל החלקים מחוברים לתמונה אחת גדולה */
  megaFrame(full) {
    const v = this.video, N = 256;
    const c = this.megaC || (this.megaC = document.createElement('canvas')); c.width = c.height = N;
    const x = c.getContext('2d', { willReadFrequently: true }); x.drawImage(v, 0, 0, N, N);
    const d = x.getImageData(0, 0, N, N).data, g = new Float64Array(N * N);
    for (let i = 0; i < N * N; i++) g[i] = d[i * 4] * 0.3 + d[i * 4 + 1] * 0.59 + d[i * 4 + 2] * 0.11;
    const out = { S: spectrum(g, N) };
    if (full) {
      const k = Math.min(1, 1600 / Math.max(v.videoWidth, v.videoHeight));
      out.c = canvasOf(v.videoWidth * k, v.videoHeight * k); out.c.getContext('2d').drawImage(v, 0, 0, out.c.width, out.c.height);
      out.t = scaledCanvas(out.c, 120);
    }
    return out;
  },
  megaStart() {
    const f = this.megaFrame(true);
    this.mega = { on: true, tiles: [], ref: { S: f.S, x: 0, y: 0 }, tw: f.c.width, th: f.c.height, cur: [0, 0], lost: 0, map: canvasOf(4096, 4096) };
    this.megaPlace(f, 0, 0);
    $r('[data-a=shoot]').classList.add('rec'); $r('[data-a=shoot]').setAttribute('aria-label', 'סיום מגה סריקה');
    $r('.scn-mini').hidden = false; this.megaMini(); this.setHint(this.modeHint());
    try { navigator.vibrate && navigator.vibrate(20); } catch (_) {}
  },
  megaStep() {
    const M = this.mega, N = 256, f = this.megaFrame(false), r = phaseShift(M.ref.S, f.S, N);
    if (r.peak < 0.1) { if (++M.lost > 2) this.setHint('לאט יותר: חזרו מעט אל החלק הקודם'); return; }
    M.lost = 0;
    const x = M.ref.x + r.dx * M.tw / N, y = M.ref.y + r.dy * M.th / N; M.cur = [x, y];
    if (Math.max(Math.abs(r.dx), Math.abs(r.dy)) > N * 0.26 && M.tiles.length < 80) {
      const g = this.megaFrame(true), at = this.megaPlace(g, x, y);
      M.ref = { S: g.S, x: at[0], y: at[1] }; M.cur = at;
      try { navigator.vibrate && navigator.vibrate(8); } catch (_) {}
      this.setHint(`נסרקו ${M.tiles.length} חלקים · ממשיכים לזוז, ולוחצים לסיום`);
    }
    this.megaMini();
  },
  // חלק חדש נבדק מול כל מה שכבר נסרק באותו מקום (לא רק מול החלק הקודם), ומתוקן,
  // כך שהסטייה לא מצטברת כשעוברים לשורה הבאה של הדף
  megaPlace(f, x, y) {
    const M = this.mega, N = 256, mx = c => 2048 + c * N / M.tw, my = c => 2048 + c * N / M.th;
    if (M.tiles.length) {
      const rx = Math.round(mx(x)), ry = Math.round(my(y));
      if (rx >= 0 && ry >= 0 && rx + N <= 4096 && ry + N <= 4096) {
        const c = this.megaR || (this.megaR = canvasOf(N, N)), cx = c.getContext('2d', { willReadFrequently: true });
        cx.fillStyle = '#808080'; cx.fillRect(0, 0, N, N); cx.drawImage(M.map, rx, ry, N, N, 0, 0, N, N);
        const d = cx.getImageData(0, 0, N, N).data, g = new Float64Array(N * N);
        for (let i = 0; i < N * N; i++) g[i] = d[i * 4] * 0.3 + d[i * 4 + 1] * 0.59 + d[i * 4 + 2] * 0.11;
        const r = phaseShift(spectrum(g, N), f.S, N);
        if (r.peak > 0.08 && Math.abs(r.dx) < N * 0.12 && Math.abs(r.dy) < N * 0.12) { x += r.dx * M.tw / N; y += r.dy * M.th / N; }
      }
    }
    M.tiles.push({ c: f.c, t: f.t, x, y });
    M.map.getContext('2d').drawImage(f.c, mx(x), my(y), N, N);
    return [x, y];
  },
  // מפה קטנה בפינה: מה כבר נסרק, ואיפה המצלמה עכשיו
  megaMini() {
    const M = this.mega; if (!M) return;
    const cv = $r('.scn-mini canvas'), x = cv.getContext('2d');
    let x0 = Math.min(...M.tiles.map(t => t.x), M.cur[0]), y0 = Math.min(...M.tiles.map(t => t.y), M.cur[1]);
    let x1 = Math.max(...M.tiles.map(t => t.x), M.cur[0]) + M.tw, y1 = Math.max(...M.tiles.map(t => t.y), M.cur[1]) + M.th;
    const k = Math.min(150 / (x1 - x0), 150 / (y1 - y0));
    cv.width = Math.max(20, Math.round((x1 - x0) * k)); cv.height = Math.max(20, Math.round((y1 - y0) * k));
    for (const t of M.tiles) x.drawImage(t.t, (t.x - x0) * k, (t.y - y0) * k, M.tw * k, M.th * k);
    x.strokeStyle = '#FFD43B'; x.lineWidth = 2; x.strokeRect((M.cur[0] - x0) * k + 1, (M.cur[1] - y0) * k + 1, M.tw * k - 2, M.th * k - 2);
  },
  megaCancel() {
    if (!this.mega) return;
    this.mega.tiles.forEach(t => { free(t.c); free(t.t); }); free(this.mega.map); this.mega = null;
    $r('[data-a=shoot]').classList.remove('rec'); $r('[data-a=shoot]').setAttribute('aria-label', 'צילום'); $r('.scn-mini').hidden = true;
  },
  async megaStop() {
    const M = this.mega; if (!M) return;
    M.on = false;
    // חלק אחרון מהמקום שבו המצלמה נמצאת עכשיו, אם זזה מאז
    if (Math.hypot(M.cur[0] - M.ref.x, M.cur[1] - M.ref.y) > M.tw * 0.08) this.megaPlace(this.megaFrame(true), M.cur[0], M.cur[1]);
    this.busy(`מחבר ${M.tiles.length} חלקים לתמונה אחת…`);
    await sleep(30);
    try {
      const x0 = Math.min(...M.tiles.map(t => t.x)), y0 = Math.min(...M.tiles.map(t => t.y));
      const W = Math.max(...M.tiles.map(t => t.x)) + M.tw - x0, H = Math.max(...M.tiles.map(t => t.y)) + M.th - y0;
      const k = Math.min(1, 9000 / Math.max(W, H)), out = canvasOf(W * k, H * k), ox = out.getContext('2d');
      ox.fillStyle = '#fff'; ox.fillRect(0, 0, out.width, out.height); ox.imageSmoothingQuality = 'high';
      // כל חלק מצויר עם שוליים רכים, כדי שהחיבור בין החלקים לא ייראה
      const fe = Math.round(Math.min(M.tw, M.th) * 0.06 * k);
      for (const [i, t] of M.tiles.entries()) {
        const w = Math.round(M.tw * k), h = Math.round(M.th * k), tc = canvasOf(w, h), tx = tc.getContext('2d');
        tx.drawImage(t.c, 0, 0, w, h);
        if (i && fe > 2) {
          tx.globalCompositeOperation = 'destination-in';
          const gx = tx.createLinearGradient(0, 0, w, 0); gx.addColorStop(0, 'rgba(0,0,0,0)'); gx.addColorStop(fe / w, '#000'); gx.addColorStop(1 - fe / w, '#000'); gx.addColorStop(1, 'rgba(0,0,0,0)');
          tx.fillStyle = gx; tx.fillRect(0, 0, w, h);
          const gy = tx.createLinearGradient(0, 0, 0, h); gy.addColorStop(0, 'rgba(0,0,0,0)'); gy.addColorStop(fe / h, '#000'); gy.addColorStop(1 - fe / h, '#000'); gy.addColorStop(1, 'rgba(0,0,0,0)');
          tx.fillStyle = gy; tx.fillRect(0, 0, w, h);
        }
        ox.drawImage(tc, Math.round((t.x - x0) * k), Math.round((t.y - y0) * k)); free(tc);
      }
      const q = sheetIn(out), blob = await toBlob(out, 0.92); free(out);
      this.megaCancel();
      const p = newPage({ src: blob, quad: q || FULL.map(p => [...p]), auto: !!q, filter: this.defFilter });
      this.pages.push(p); await refreshThumb(p); this.updateStack();
      this.busy(null); this.cur = this.pages.length - 1; this.show('review');
      say(q ? 'הדף הגדול חובר. אפשר ליישר את המסגרת בעריכת העמוד' : 'הדף הגדול חובר. אפשר לחתוך אותו בעריכת העמוד');
    } catch (e) { console.error(e); this.busy(null); this.megaCancel(); say('החיבור נכשל: ' + (e.message || e)); }
  },
  async shoot() {
    if (this.mode === 'mega') { if (!this.stream) return; return this.mega && this.mega.on ? this.megaStop() : this.megaStart(); }
    if (this.shooting || !this.stream) return;
    const v = this.video; if (!v.videoWidth) return;
    this.shooting = true;
    try {
      const W = v.videoWidth, H = v.videoHeight;
      const fl = $r('.scn-flash'); fl.classList.remove('go'); void fl.offsetWidth; fl.classList.add('go');
      try { navigator.vibrate && navigator.vibrate(25); } catch (_) {}
      let q = this.quad ? this.quad.map(p => [...p]) : null;
      const sp = this.spread && this.pick === 'both' && this.mode !== 'book' ? this.spread : null;
      // ״שני העמודים״: במחברת עם ספירלה — שני עמודים נפרדים. בדף רחב בלי ספירלה — הדף כולו
      if (sp) q = sp.full.map(p => [...p]);
      const both = sp && sp.spiral ? sp : null;
      // תמונה אמיתית של המצלמה (חדה פי כמה מהווידאו), ואם אין — התמונה מהווידאו
      let c = await this.photo(W, H, q);
      if (c && c.q !== undefined) { q = c.q; c = c.c; }
      if (!c) { const k = Math.min(1, 4000 / Math.max(W, H)); c = canvasOf(W * k, H * k); c.getContext('2d').drawImage(v, 0, 0, c.width, c.height); }
      // בתעודה: אם לא זוהתה, המסגרת המסומנת על המסך
      if (!q && this.mode === 'id') q = [[0.08, 0.22], [0.92, 0.22], [0.92, 0.78], [0.08, 0.78]];
      this.lastShot = Date.now(); this.lastShotQuad = q; this.stable = 0;
      const blob = await toBlob(c, 0.93); free(c);
      await this.addShot(blob, q, false, both ? 'book' : null, both ? both.t : 0.5);
    } catch (e) { console.error(e); this.setHint('הצילום נכשל, נסו שוב'); }
    finally { this.shooting = false; }
  },
  async addShot(blob, q, fromGallery, forceMode, splitT = 0.5) {
    const m = forceMode || (this.mode === 'qr' ? 'doc' : this.mode);
    if (m === 'book') {
      // ספר: הדף כולו מיושר, ונחתך לשני עמודים. בספר בעברית העמוד הימני קודם
      const img = await decodeImg(blob), c = warp(img, q || FULL, 3600); if (img.close) img.close();
      const half = Math.round(c.width * Math.max(0.3, Math.min(0.7, splitT))), mk = async (x0, w) => { const o = canvasOf(w, c.height); o.getContext('2d').drawImage(c, x0, 0, w, c.height, 0, 0, w, c.height); const b = await toBlob(o, 0.92); free(o); return b; };
      const right = await mk(half, c.width - half), left = await mk(0, half); free(c);
      for (const b of this.rtlBook ? [right, left] : [left, right]) { const p = newPage({ src: b }); this.pages.push(p); await refreshThumb(p); }
    } else if (m === 'id') {
      const img = await decodeImg(blob), c = warp(img, q || FULL, 1400); if (img.close) img.close();
      // כרטיס לרוחב
      const card = c.height > c.width ? rotated(c, 90) : c;
      if (!this.idFront) { this.idFront = card; this.setHint(this.modeHint()); this.sheet('id1'); return; }
      const front = this.idFront; this.idFront = null;
      await this.addIdPage(front, card);
    } else {
      const p = newPage({ src: blob, quad: q || FULL.map(p => [...p]), auto: !!q, filter: m === 'board' ? 'board' : this.defFilter });
      this.pages.push(p); await refreshThumb(p);
    }
    this.updateStack();
    if (!fromGallery) this.setHint(`צולם ✓ עמוד ${this.pages.length}`);
  },
  // שני צדי התעודה על עמוד A4 אחד, בגודל האמיתי (85.6×54 מ״מ), מוכן להדפסה
  async addIdPage(front, back) {
    const A4w = 1654, A4h = 2339, mm = A4w / 210, cw = Math.round(85.6 * mm), ch = Math.round(54 * mm);
    const pg = canvasOf(A4w, A4h), x = pg.getContext('2d');
    x.fillStyle = '#fff'; x.fillRect(0, 0, A4w, A4h); x.imageSmoothingQuality = 'high';
    const left = Math.round((A4w - cw) / 2), top1 = Math.round(A4h * 0.14), top2 = top1 + ch + Math.round(30 * mm);
    x.drawImage(front, left, top1, cw, ch); if (back) x.drawImage(back, left, top2, cw, ch);
    x.strokeStyle = '#D0D6E0'; x.lineWidth = 2; x.strokeRect(left, top1, cw, ch); if (back) x.strokeRect(left, top2, cw, ch);
    free(front); free(back);
    const p = newPage({ src: await toBlob(pg, 0.93), filter: 'orig' }); free(pg);
    this.pages.push(p); await refreshThumb(p);
    this.updateStack(); this.setHint(`התעודה נוספה ✓ עמוד ${this.pages.length}`);
  },
  updateStack() {
    const s = $r('[data-a=review]'), last = this.pages[this.pages.length - 1];
    s.hidden = !last; if (last) { s.querySelector('img').src = last.thumb || ''; s.querySelector('b').textContent = this.pages.length; s.setAttribute('aria-label', `מעבר לעמודים שנסרקו (${this.pages.length})`); }
  },
  // חלונית קטנה מעל המצלמה
  sheet(kind, data) {
    const el = $r('.scn-sheet');
    if (!kind) { el.hidden = true; el.innerHTML = ''; return; }
    el.hidden = false;
    if (kind === 'id1') {
      el.innerHTML = `<h3>הצד הקדמי נשמר</h3><p style="margin:0 0 12px;font-size:14px">הפכו את התעודה וצלמו את הצד האחורי. שני הצדדים יופיעו על עמוד אחד, בגודל האמיתי, מוכנים להדפסה.</p>
        <div class="row"><button class="main" type="button" data-s="ok">צילום הצד האחורי</button><button type="button" data-s="one">צד אחד מספיק</button><button type="button" data-s="redo">צילום מחדש</button></div>`;
      el.querySelector('[data-s=ok]').onclick = () => this.sheet(null);
      el.querySelector('[data-s=one]').onclick = async () => { this.sheet(null); const f = this.idFront; this.idFront = null; await this.addIdPage(f, null); this.setHint(this.modeHint()); };
      el.querySelector('[data-s=redo]').onclick = () => { free(this.idFront); this.idFront = null; this.sheet(null); this.setHint(this.modeHint()); };
    } else if (kind === 'qr') {
      const url = /^https?:\/\/\S+$/i.test(data.trim()) ? data.trim() : '';
      el.innerHTML = `<h3>${url ? 'נמצא קישור' : 'נמצא טקסט'}</h3><p class="val"></p><div class="row">${url ? '<button class="main" type="button" data-s="go">פתיחת הקישור</button>' : ''}<button type="button" data-s="copy">העתקה</button><button type="button" data-s="again">סריקת קוד נוסף</button></div>`;
      el.querySelector('.val').textContent = data;
      if (url) el.querySelector('[data-s=go]').onclick = () => { window.open(url, '_blank', 'noopener,noreferrer'); };
      el.querySelector('[data-s=copy]').onclick = async e => { try { await navigator.clipboard.writeText(data); e.target.textContent = 'הועתק ✓'; } catch (_) { e.target.textContent = 'ההעתקה נחסמה'; } };
      el.querySelector('[data-s=again]').onclick = () => { this.qrLast = data; this.qrLastAt = Date.now(); this.sheet(null); };
    }
  },

  /* ---- קוד QR ---- */
  async prepQr() {
    if (this.qrDet || window.jsQR) return;
    try {
      if ('BarcodeDetector' in window) { const f = await BarcodeDetector.getSupportedFormats(); if (f.includes('qr_code')) { this.qrDet = new BarcodeDetector({ formats: ['qr_code'] }); return; } }
    } catch (_) {}
    try { await loadJs('https://cdn.jsdelivr.net/npm/jsqr@1.4.0/dist/jsQR.js'); } catch (_) { this.setHint('קורא הקודים לא נטען. בדקו חיבור לאינטרנט'); }
  },
  async scanQr() {
    const v = this.video; let txt = '';
    if (this.qrDet) { const r = await this.qrDet.detect(v); txt = r && r[0] && r[0].rawValue || ''; }
    else if (window.jsQR) {
      const k = Math.min(1, 900 / Math.max(v.videoWidth, v.videoHeight)), c = this.qrC || (this.qrC = document.createElement('canvas'));
      c.width = Math.round(v.videoWidth * k); c.height = Math.round(v.videoHeight * k);
      const x = c.getContext('2d', { willReadFrequently: true }); x.drawImage(v, 0, 0, c.width, c.height);
      const r = window.jsQR(x.getImageData(0, 0, c.width, c.height).data, c.width, c.height, { inversionAttempts: 'dontInvert' }); txt = r && r.data || '';
    }
    if (txt && !(txt === this.qrLast && Date.now() - this.qrLastAt < 4000) && this.mode === 'qr') {
      try { navigator.vibrate && navigator.vibrate(30); } catch (_) {}
      this.sheet('qr', txt);
    }
  },

  /* ---- ייבוא מהגלריה ---- */
  async fromFiles(files) {
    files = [...files].filter(f => /^image\//.test(f.type) || /\.(jpe?g|png|webp|heic|heif)$/i.test(f.name));
    if (!files.length) return;
    const was = this.mode; if (this.mode === 'qr' || this.mode === 'id') this.mode = 'doc';
    const bad = [];
    for (const [i, f] of files.entries()) {
      this.busy(files.length > 1 ? `מכין את התמונות… ${i + 1} מתוך ${files.length}` : 'מכין את התמונה…');
      try {
        const img = await decodeImg(f), c = scaledCanvas(img, 4000); if (img.close) img.close();
        let q = sheetIn(c), sp = null;
        // מחברת פתוחה בתמונה: העמוד שיש בו יותר כתב
        if (q && this.mode === 'doc') { const sm = scaledCanvas(c, 900); sp = spreadOf(sm, q); free(sm); if (sp) q = sp.ink.left > sp.ink.right ? sp.left : sp.right; }
        const b = await toBlob(c, 0.93); free(c);
        await this.addShot(b, q, true);
      } catch (e) { console.warn(e); bad.push(f.name + (/heic|heif/i.test(f.type + f.name) ? ' (HEIC: שמרו כ-JPG)' : '')); }
    }
    this.mode = was; this.busy(null);
    if (bad.length) say('לא הצלחתי לפתוח: ' + bad.join(', '));
    if (this.pages.length) this.show('review');
  },

  /* ---- כל העמודים ---- */
  renderReview() {
    const n = this.pages.length;
    $r('[data-r=title]').textContent = n === 1 ? 'עמוד אחד' : `${n} עמודים`;
    const nm = $r('[data-r=name]'); nm.value = this.name; nm.oninput = () => { this.name = nm.value; };
    const g = $r('[data-r=grid]');
    g.innerHTML = this.pages.map((p, i) => `<li><button class="pv" type="button" data-ed="${i}" aria-label="עריכת עמוד ${i + 1}"><img alt="" src="${p.thumb}"></button>
      <div class="row"><button type="button" data-mv="${i}:-1" aria-label="להקדים את עמוד ${i + 1}" ${i ? '' : 'disabled'}>${I.prev}</button><span>עמוד ${i + 1}</span>
      <button type="button" data-mv="${i}:1" aria-label="לדחות את עמוד ${i + 1}" ${i < n - 1 ? '' : 'disabled'}>${I.next}</button><button type="button" data-del="${i}" aria-label="מחיקת עמוד ${i + 1}">${I.trash}</button></div></li>`).join('')
      + `<li class="add" role="button" tabindex="0" data-add="1">${I.plus}<span>הוספת עמודים</span></li>`;
    g.querySelectorAll('[data-ed]').forEach(b => b.onclick = () => { this.cur = +b.dataset.ed; this.show('edit'); });
    g.querySelectorAll('[data-mv]').forEach(b => b.onclick = () => { const [i, d] = b.dataset.mv.split(':').map(Number), j = i + d; [this.pages[i], this.pages[j]] = [this.pages[j], this.pages[i]]; this.renderReview(); });
    g.querySelectorAll('[data-del]').forEach(b => b.onclick = () => { this.pages.splice(+b.dataset.del, 1); this.updateStack(); if (!this.pages.length) this.show('cam'); else this.renderReview(); });
    const add = g.querySelector('[data-add]'); add.onclick = () => this.show('cam'); add.onkeydown = e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); this.show('cam'); } };
    const allF = this.pages.every(p => p.filter === this.pages[0].filter) ? this.pages[0].filter : '';
    root.querySelectorAll('[data-r=allf] [data-f]').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.f === allF)));
    $r('[data-a=inkall]').setAttribute('aria-pressed', String(this.pages.every(p => p.ink)));
    root.querySelectorAll('[data-o]').forEach(el => { el.value = this.opts[el.dataset.o]; el.onchange = el.oninput = () => { this.opts[el.dataset.o] = el.value; LS.set('siman.scan.opts', JSON.stringify(this.opts)); }; });
    const sh = $r('[data-x=share]'); sh.hidden = !(navigator.canShare && navigator.canShare({ files: [new File([new Uint8Array(4)], 'a.pdf', { type: 'application/pdf' })] }));
  },
  async setAll(change) {
    this.busy('מעדכן את העמודים…');
    try { for (const p of this.pages) { Object.assign(p, change); await refreshThumb(p); } } finally { this.busy(null); }
    if (change.filter && change.filter !== 'board') { this.defFilter = change.filter; LS.set('siman.scan.filter', change.filter); }
    this.renderReview(); this.updateStack();
  },

  /* ---- עריכת עמוד ---- */
  async renderEdit() {
    const p = this.pages[this.cur]; if (!p) return this.show('review');
    $r('[data-e=title]').textContent = `עמוד ${this.cur + 1} מתוך ${this.pages.length}`;
    $r('[data-a=pprev]').disabled = this.cur === 0; $r('[data-a=pnext]').disabled = this.cur >= this.pages.length - 1;
    root.querySelector('[data-t=ink]').setAttribute('aria-pressed', String(!!p.ink));
    root.querySelector('[data-t=crop]').setAttribute('aria-pressed', String(this.cropping));
    const stg = $r('[data-e=stage]');
    if (this.cropping) return this.renderCrop(p, stg);
    stg.innerHTML = `<img class="big" alt="עמוד ${this.cur + 1}" src="${p.thumb}">`;
    $r('[data-e=tools]').hidden = false;
    // תצוגה חדה יותר מהתמונה הקטנה
    const v = p.ver; render(p, 1400).then(c => { if (v === p.ver && this.view === 'edit' && !this.cropping && this.pages[this.cur] === p) { const im = stg.querySelector('img.big'); if (im) im.src = c.toDataURL('image/jpeg', 0.85); } free(c); });
    // דוגמאות של כל מסנן על העמוד הזה
    const fl = $r('[data-e=filters]'); fl.hidden = false;
    fl.innerHTML = FILTERS.map(([k, l]) => `<button type="button" data-ef="${k}" aria-pressed="${p.filter === k}"><img alt=""><span>${l}</span></button>`).join('');
    fl.querySelectorAll('[data-ef]').forEach(b => b.onclick = async () => { p.filter = b.dataset.ef; this.busy('מחיל…'); await refreshThumb(p); this.busy(null); this.renderEdit(); });
    const base = await (async () => { const img = await decodeImg(p.src); const c = rotated(warp(img, p.quad, 160), p.rot); if (img.close) img.close(); return c; })();
    for (const b of fl.querySelectorAll('[data-ef]')) {
      const c = canvasOf(base.width, base.height); c.getContext('2d').drawImage(base, 0, 0);
      applyFilter(c, b.dataset.ef, p.ink); b.querySelector('img').src = c.toDataURL('image/jpeg', 0.7); free(c);
    }
    free(base);
  },
  renderCrop(p, stg) {
    $r('[data-e=filters]').hidden = true;
    const url = URL.createObjectURL(p.src);
    stg.innerHTML = `<div class="scn-crop"><img alt="התמונה המקורית" src="${url}" draggable="false"><svg viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true"><path class="shade" fill-rule="evenodd"/><polygon class="edge"/></svg>
      ${['ימין למעלה', 'שמאל למעלה', 'שמאל למטה', 'ימין למטה'].map((l, k) => `<button type="button" class="scn-hd" data-k="${[1, 0, 3, 2][k]}" aria-label="פינה: ${l}. אפשר להזיז בחצים"></button>`).join('')}</div>
      <div style="position:absolute;bottom:8px;inset-inline:0;display:flex;gap:8px;justify-content:center;flex-wrap:wrap">
      <button class="scn-btn" type="button" data-c="detect">זיהוי הדף</button><button class="scn-btn" type="button" data-c="full">כל התמונה</button><button class="scn-btn main" type="button" data-c="ok">${I.ok} אישור</button></div>`;
    const wrap = stg.querySelector('.scn-crop'), im = wrap.querySelector('img'), shade = wrap.querySelector('.shade'), poly = wrap.querySelector('.edge'), hs = [...wrap.querySelectorAll('.scn-hd')];
    // הכפתורים בתבנית אינם לפי סדר הפינות: מסדרים לפי data-k
    hs.sort((a, b) => a.dataset.k - b.dataset.k);
    let q = p.quad.map(x => [...x]);
    im.onload = () => {
      URL.revokeObjectURL(url);
      const box = stg.getBoundingClientRect(), r = im.naturalWidth / im.naturalHeight; let w = box.width - 24, hh = w / r; const maxH = box.height - 70;
      if (hh > maxH) { hh = maxH; w = hh * r; }
      wrap.style.width = w + 'px'; wrap.style.height = hh + 'px'; im.style.width = '100%'; im.style.height = '100%';
    };
    const draw = () => {
      const pts = q.map(([x, y]) => `${x * 100},${y * 100}`);
      poly.setAttribute('points', pts.join(' ')); shade.setAttribute('d', `M0,0H100V100H0Z M${pts.join(' L')}Z`);
      hs.forEach((hd, k) => { hd.style.left = q[k][0] * 100 + '%'; hd.style.top = q[k][1] * 100 + '%'; });
    };
    draw();
    hs.forEach((hd, k) => {
      hd.onpointerdown = e => { e.preventDefault(); hd.setPointerCapture(e.pointerId); hd.classList.add('on'); };
      hd.onpointermove = e => { if (!hd.hasPointerCapture(e.pointerId)) return; const r = wrap.getBoundingClientRect();
        q[k] = [Math.max(0, Math.min(1, (e.clientX - r.left) / r.width)), Math.max(0, Math.min(1, (e.clientY - r.top) / r.height))]; draw(); };
      hd.onpointerup = hd.onpointercancel = e => { hd.classList.remove('on'); try { hd.releasePointerCapture(e.pointerId); } catch (_) {} };
      hd.onkeydown = e => { const st = e.shiftKey ? 0.05 : 0.01, mv = { ArrowLeft: [-st, 0], ArrowRight: [st, 0], ArrowUp: [0, -st], ArrowDown: [0, st] }[e.key];
        if (!mv) return; e.preventDefault(); q[k] = [Math.max(0, Math.min(1, q[k][0] + mv[0])), Math.max(0, Math.min(1, q[k][1] + mv[1]))]; draw(); };
    });
    stg.querySelector('[data-c=full]').onclick = () => { q = FULL.map(x => [...x]); draw(); };
    stg.querySelector('[data-c=detect]').onclick = async e => { const b = e.currentTarget; b.textContent = 'מזהה…';
      const img = await decodeImg(p.src), f = sheetIn(img); if (img.close) img.close();
      if (f) { q = f; draw(); b.textContent = 'זוהה ✓'; } else b.textContent = 'לא זוהה, גררו את הפינות'; };
    stg.querySelector('[data-c=ok]').onclick = async () => {
      if (quadArea(q) < 0.02) { say('המסגרת קטנה מדי'); return; }
      p.quad = q; this.cropping = false; this.busy('מיישר את הדף…'); await refreshThumb(p); this.busy(null); this.updateStack(); this.renderEdit();
    };
  },
  async tool(t) {
    const p = this.pages[this.cur]; if (!p) return;
    if (t === 'crop') { this.cropping = !this.cropping; return this.renderEdit(); }
    if (t === 'rot') { p.rot = (p.rot + 90) % 360; this.busy('מסובב…'); await refreshThumb(p); this.busy(null); this.updateStack(); return this.renderEdit(); }
    if (t === 'ink') { p.ink = !p.ink; this.busy(p.ink ? 'מוחק את סימני העט…' : 'מחזיר…'); await refreshThumb(p); this.busy(null); return this.renderEdit(); }
    if (t === 'all') { await this.setAll({ filter: p.filter, ink: p.ink }); this.show('edit'); say('הסגנון הוחל על כל העמודים'); return; }
    if (t === 'del') { this.pages.splice(this.cur, 1); this.updateStack(); if (!this.pages.length) return this.show('cam'); this.cur = Math.min(this.cur, this.pages.length - 1); return this.renderEdit(); }
  },

  /* ---- שמירה ---- */
  fileBase() { return (this.name || '').replace(/[\\/:*?"<>|]+/g, ' ').trim() || defaultName(); },
  quality() { return { high: [2600, 0.88], mid: [2000, 0.8], small: [1500, 0.66] }[this.opts.q] || [2000, 0.8]; },
  async pageJpeg(p, max, q) {
    const c = await render(p, max);
    if (this.opts.wm.trim()) watermark(c, this.opts.wm.trim());
    const b = await toBlob(c, q); const w = c.width, hh = c.height; free(c);
    return { bytes: new Uint8Array(await b.arrayBuffer()), blob: b, w, h: hh };
  },
  async buildPdf() {
    if (!window.PDFLib) throw new Error('רכיב ה-PDF לא נטען');
    const { PDFDocument } = PDFLib, pdf = await PDFDocument.create(), [max, qq] = this.quality();
    const sizes = { a4: [595.28, 841.89], letter: [612, 792] };
    for (const [i, p] of this.pages.entries()) {
      this.busy(`בונה את הקובץ… עמוד ${i + 1} מתוך ${this.pages.length}`);
      const j = await this.pageJpeg(p, max, qq), img = await pdf.embedJpg(j.bytes), land = j.w > j.h * 1.05;
      let pw, ph;
      if (sizes[this.opts.size]) { [pw, ph] = sizes[this.opts.size]; if (land) [pw, ph] = [ph, pw]; }
      else if (land) { ph = 595.28; pw = ph * j.w / j.h; } else { pw = 595.28; ph = pw * j.h / j.w; }
      const k = Math.min(pw / j.w, ph / j.h), dw = j.w * k, dh = j.h * k;
      pdf.addPage([pw, ph]).drawImage(img, { x: (pw - dw) / 2, y: (ph - dh) / 2, width: dw, height: dh });
    }
    pdf.setTitle(this.fileBase()); pdf.setCreator('סימן'); pdf.setProducer('סימן');
    return new File([await pdf.save()], this.fileBase() + '.pdf', { type: 'application/pdf' });
  },
  async act(x) {
    if (!this.pages.length) return;
    try {
      if (x === 'pdf') { const f = await this.buildPdf(); this.busy(null); download(f); say('הקובץ נשמר בתיקיית ההורדות'); }
      else if (x === 'share') { const f = await this.buildPdf(); this.busy(null); try { await navigator.share({ files: [f], title: f.name }); } catch (e) { if (e.name !== 'AbortError') download(f); } }
      else if (x === 'jpg') await this.saveJpgs();
      else if (x === 'open' || x === 'ocr') await this.toEditor(x === 'ocr');
      else if (x === 'word') await this.toWord();
      else if (x === 'excel') await this.toExcel();
    } catch (e) { console.error(e); say('השמירה נכשלה: ' + (e.message || e)); }
    finally { this.busy(null); }
  },
  async saveJpgs() {
    const [max, qq] = this.quality(), base = this.fileBase(), files = [];
    for (const [i, p] of this.pages.entries()) {
      this.busy(`מכין תמונה ${i + 1} מתוך ${this.pages.length}…`);
      const j = await this.pageJpeg(p, max, qq);
      files.push(new File([j.blob], this.pages.length > 1 ? `${base} - ${i + 1}.jpg` : `${base}.jpg`, { type: 'image/jpeg' }));
    }
    this.busy(null);
    if (navigator.canShare && navigator.canShare({ files }) && /Android|iPhone|iPad/i.test(navigator.userAgent)) {
      try { await navigator.share({ files, title: base }); return; } catch (e) { if (e.name === 'AbortError') return; }
    }
    if (files.length === 1) { download(files[0]); return; }
    this.busy('אורז את התמונות…');
    if (!window.JSZip) await loadJs('jszip.min.js');
    const z = new JSZip(); files.forEach(f => z.file(f.name, f));
    download(new File([await z.generateAsync({ type: 'blob' })], base + '.zip', { type: 'application/zip' }));
  },
  async toEditor(ocr) {
    if (typeof openFile !== 'function') throw new Error('העורך לא נטען');
    const f = await this.buildPdf(); this.busy(null);
    if (typeof S !== 'undefined' && S.dirty && !confirm('בקובץ הפתוח יש שינויים שלא נשמרו. לפתוח את הסריקה במקומו?')) return;
    try { if (typeof S !== 'undefined') S.dirty = false; } catch (_) {}
    this.close(true);
    await openFile(f);
    try { if (typeof setDirty === 'function') setDirty(true); } catch (_) {}
    if (ocr) { if (typeof runOCR === 'function' && typeof S !== 'undefined' && S.pages) await runOCR(S.pages.map((_, i) => i)); }
    else say('הסריקה פתוחה: אפשר לחתום, לסמן ולמלא. לשמירה, כפתור השמירה למעלה');
  },
  /* דף שצולם לרוחב או הפוך (טבלה רחבה, מחברת שסובבה): קודם בודקים במכשיר אם השורות אופקיות.
     רק כשלא ברור — שואלים את הקורא איך לסובב (בקשה קטנה ומהירה). הסיבוב נשמר בעמוד, ורואים אותו גם בתמונה הקטנה */
  async orient(p) {
    if (p.orientDone) return;
    p.orientDone = true;
    try {
      const c = await render(Object.assign({}, p, { filter: 'orig', ink: false }), 700);
      const dir = textDirection(c);
      if (dir === 'h') { free(c); return; }
      const u = c.toDataURL('image/jpeg', 0.8); free(c);
      const AI = window.SIMAN_AI; if (!AI) return;
      const raw = await AI({ mode: 'orient', image: u.split(',')[1], mime: 'image/jpeg' }, { timeout: 40000 });
      const r = parseInt(JSON.parse(String(raw).replace(/^\s*```(?:json)?|```\s*$/g, '')).rotate, 10);
      if ([90, 180, 270].includes(r)) { p.rot = (p.rot + r) % 360; await refreshThumb(p); this.updateStack(); }
    } catch (e) { console.warn('orient', e); }
  },
  // טבלה בדף → גיליון Excel: כל טבלה בגיליון משלה, תאים לא ודאיים ברשימה נפרדת לבדיקה, ופתיחה בעורך הגיליונות של סימן
  async toExcel() {
    const AI = window.SIMAN_AI;
    if (!AI) throw new Error('קריאת הטקסט לא נטענה');
    const tables = [], t0 = Date.now(), strip = t => String(t).replace(/^\s*```(?:json)?|```\s*$/g, '');
    for (const [i, p] of this.pages.entries()) {
      const label = this.pages.length > 1 ? `עמוד ${i + 1} מתוך ${this.pages.length}: ` : '';
      this.busy(label + 'בודק שהדף ישר…');
      await this.orient(p);
      this.busy(label + 'מכין את התמונה…');
      const pic = await tablePic(p);
      const clock = setInterval(() => this.busy(`${label}קורא את הטבלה… ${Math.round((Date.now() - t0) / 1000)} שנ׳`), 1000);
      let j;
      try { j = JSON.parse(strip(await AI({ mode: 'table', image: pic.b64, strips: pic.strips, mime: 'image/jpeg' }, { timeout: 150000, onWait: s => this.busy(`השירות עמוס, ממשיך בעוד ${s} שניות…`) }))); }
      finally { clearInterval(clock); }
      for (const t of (j && j.tables) || []) if (Array.isArray(t.rows) && t.rows.some(r => Array.isArray(r) && r.some(c => String(c || '').trim()))) tables.push({ title: t.title || '', rows: t.rows, page: i + 1 });
    }
    if (!tables.length) throw new Error('לא נמצאה טבלה בדף');
    this.busy('בונה את הגיליון…');
    if (!window.XLSX || !XLSX.utils) await loadJs('xlsx.mini.min.js');
    const wb = XLSX.utils.book_new(), unsure = [['גיליון', 'תא', 'מה נקרא']];
    let he = 0, lat = 0;
    tables.forEach((t, ti) => {
      const width = Math.max(...t.rows.map(r => (r || []).length));
      const rows = t.rows.map(r => { const a = (r || []).map(c => String(c ?? '')); while (a.length < width) a.push(''); return a; });
      const top = t.title ? [[t.title], []] : [], aoa = [...top], name = (tables.length > 1 ? `טבלה ${ti + 1}` : 'טבלה').slice(0, 31);
      rows.forEach((r, ri) => aoa.push(r.map((c, ci) => {
        if (/[\u0590-\u05FF]/.test(c)) he++; else if (/[A-Za-z]/.test(c)) lat++;
        // טקסט שנשבר לשתי שורות בתוך התא מחובר לשורה אחת, כמו שהוא נכתב
        c = c.replace(/\s*\n\s*/g, (m, off, str) => /[,.;:\/\-]$/.test(str.slice(0, off)) || /^[,.;:\/\-]/.test(str.slice(off + m.length)) ? '' : ' ');
        if (/\[\[/.test(c)) { c = c.replace(/\[\[|\]\]/g, ''); unsure.push([name, XLSX.utils.encode_cell({ r: top.length + ri, c: ci }), c]); }
        return cellValue(c);
      })));
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws['!cols'] = [...Array(width).keys()].map(ci => ({ wch: Math.min(48, Math.max(8, ...rows.map(r => String(r[ci]).length + 2))) }));
      XLSX.utils.book_append_sheet(wb, ws, name);
    });
    if (unsure.length > 1) { const ws = XLSX.utils.aoa_to_sheet(unsure); ws['!cols'] = [{ wch: 12 }, { wch: 8 }, { wch: 40 }]; XLSX.utils.book_append_sheet(wb, ws, 'לבדיקה'); }
    // טבלה בעברית (יותר תאים עם עברית מאשר עם אנגלית; מספרים לא נספרים) → גיליון מימין לשמאל,
    // כך שהעמודה הימנית בדף (למשל מק״ט) היא גם העמודה הימנית בקובץ
    if (he && he >= lat) { wb.Workbook = { Views: [{ RTL: true }] }; wb.SheetNames.forEach(n => { wb.Sheets[n]['!views'] = [{ RTL: true }]; }); }
    const bytes = new Uint8Array(XLSX.write(wb, { bookType: 'xlsx', type: 'array' })), name = this.fileBase() + '.xlsx';
    this.busy(null); this.close(true);
    if (typeof openExcelEditor === 'function') await openExcelEditor(bytes, name, false);
    else download(new File([bytes], name, { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    say(unsure.length > 1 ? `הטבלה מוכנה. ${unsure.length - 1} תאים לבדיקה ברשימה ״לבדיקה״` : 'הטבלה מוכנה לעריכה');
  },
  async toWord() {
    const SC = window.SIMAN_SCAN;
    if (!SC || !SC.openReader) throw new Error('קריאת הטקסט לא נטענה');
    const pics = [];
    for (const [i, p] of this.pages.entries()) {
      this.busy(`בודק שעמוד ${i + 1} ישר…`);
      await this.orient(p);
      this.busy(`מכין עמוד ${i + 1} מתוך ${this.pages.length} לקריאה…`);
      pics.push(await readingPic(p, SC.quality === 'fast'));
    }
    const name = this.fileBase();
    this.busy(null); this.close(true);
    SC.openReader(); SC.name = name; SC.run(pics);
  },
};

/* ---- כיוון השורות בדף: 'h' (אופקיות, הדף ישר או הפוך), 'v' (אנכיות: הדף צולם לרוחב), '?' ----
   שורות כתב יוצרות פרופיל "מפוספס" לאורך הכיוון שלהן: הרבה דיו בשורה, רווח, שוב שורה */
function textDirection(c) {
  const W = c.width, H = c.height, d = c.getContext('2d', { willReadFrequently: true }).getImageData(0, 0, W, H).data;
  const g = new Float32Array(W * H);
  for (let i = 0, k = 0; k < g.length; i += 4, k++) g[k] = Math.min(d[i], d[i + 1], d[i + 2]);
  const bg = Float32Array.from(g); boxBlur(bg, W, H, 10);
  const rows = new Float64Array(H), cols = new Float64Array(W);
  for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) { const k = y * W + x; if (bg[k] - g[k] > Math.max(25, bg[k] * 0.15)) { rows[y]++; cols[x]++; } }
  const stripe = p => { const n = p.length; let m = 0; for (const v of p) m += v; m /= n; if (!m) return 0; let e = 0;
    for (let i = 0; i < n; i++) { let s = 0, c2 = 0; for (let k = i - 15; k <= i + 15; k++) if (k >= 0 && k < n) { s += p[k]; c2++; } e += (p[i] - s / c2) ** 2; }
    return e / n / (m * m); };
  const vr = stripe(rows), vc = stripe(cols);
  return vr > vc * 1.8 ? 'h' : vc > vr * 1.8 ? 'v' : '?';
}

/* ---- תמונה לקריאת טבלה: הדף בצבע נקי (קווי הטבלה נשמרים!), ועוד שלוש רצועות חדות מלמעלה למטה ---- */
async function tablePic(p) {
  const c = await render(Object.assign({}, p, { filter: p.filter === 'orig' || p.filter === 'bw' ? 'clean' : p.filter, ink: false }), 2600);
  const u = c.toDataURL('image/jpeg', 0.88), strips = [], n = c.height > c.width * 1.1 ? 3 : 2, ov = 0.06;
  for (let s = 0; s < n; s++) {
    const y0 = Math.max(0, c.height * (s / n - ov)), y1 = Math.min(c.height, c.height * ((s + 1) / n + ov)), k = Math.min(1, 2000 / c.width);
    const o = canvasOf(c.width * k, (y1 - y0) * k); o.getContext('2d').drawImage(c, 0, y0, c.width, y1 - y0, 0, 0, o.width, o.height);
    strips.push(o.toDataURL('image/jpeg', 0.86).split(',')[1]); free(o);
  }
  free(c);
  return { b64: u.split(',')[1], strips };
}
// תא: מספר רגיל נשמר כמספר (כדי שאפשר יהיה לחשב איתו), כל השאר כטקסט כמו שנכתב
function cellValue(c) {
  const t = String(c).trim(), n = t.replace(/,/g, '');
  if (/^-?\d+(\.\d+)?$/.test(n) && !(/^-?0\d/.test(n)) && n.length < 16) return Number(n);
  return t;
}

/* ---- תמונה לקריאת הטקסט: לא מה שיפה לעין, אלא מה שהכי קל לקרוא ----
   · הנייר לבן לגמרי, גם בצללים
   · כל עט (כחול, שחור, אדום) נעשה כהה, והקווים והמשבצות המודפסים של המחברת (בהירים) נעלמים:
     לוקחים מכל נקודה את הערוץ הכהה ביותר שלה — דיו כחול כהה בכל הערוצים חוץ מהכחול, משבצת תכלת בהירה ברובם
   · רק האזור שיש בו כתב נשלח, בלי החלק הריק של הדף — כך כל אות גדולה וחדה יותר
   · ועוד רצועות חדות של אותו אזור, מלמעלה למטה, כדי לקרוא כל מילה */
async function readingPic(p, fast) {
  const img = await decodeImg(p.src);
  // מסגרת מעט רחבה יותר: עדיף פס של שולחן בשוליים (נחתך בהמשך) מאשר מילה שנחתכה
  const cx = p.quad.reduce((a, q) => a + q[0], 0) / 4, cy = p.quad.reduce((a, q) => a + q[1], 0) / 4;
  const wide = isFull(p.quad) ? p.quad : p.quad.map(([qx, qy]) => [Math.max(0, Math.min(1, cx + (qx - cx) * 1.05)), Math.max(0, Math.min(1, cy + (qy - cy) * 1.04))]);
  let c = warp(img, wide, 3600); if (img.close) img.close();
  c = rotated(c, p.rot);
  const W = c.width, H = c.height, x = c.getContext('2d', { willReadFrequently: true }), im = x.getImageData(0, 0, W, H), d = im.data;
  const bg = paperMap(c);
  const g = new Uint8ClampedArray(W * H), hist = new Uint32Array(256);
  const ix0 = W * 0.08, ix1 = W * 0.92, iy0 = H * 0.06, iy1 = H * 0.94;
  let inner = 0;
  for (let i = 0, k = 0; i < d.length; i += 4, k++) {
    const r = d[i] * 255 / Math.max(40, bg[i]), gg = d[i + 1] * 255 / Math.max(40, bg[i + 1]), b = d[i + 2] * 255 / Math.max(40, bg[i + 2]);
    g[k] = Math.min(255, r, gg, b);
    const xx = k % W, yy = (k / W) | 0;
    if (xx > ix0 && xx < ix1 && yy > iy0 && yy < iy1) { hist[g[k]]++; inner++; }
  }
  // הגבולות נקבעים לפי פנים הדף (בלי קצה השולחן והספירלה):
  // lo = הדיו הכהה ביותר → שחור. hi = הקווים או המשבצות המודפסים (או הנייר) → לבן.
  // כך גם כתב דק או רחוק נשאר כהה וקריא, והמשבצות נעלמות
  const pct = f => { let a = 0; for (let v = 0; v < 256; v++) { a += hist[v]; if (a >= inner * f) return v; } return 255; };
  let lo = Math.min(pct(0.003), 170), hi = Math.min(242, pct(0.1) - 6);
  if (hi < lo + 28) hi = Math.min(250, lo + 28);
  const lut = new Uint8ClampedArray(256);
  for (let v = 0; v < 256; v++) { const t = Math.max(0, Math.min(1, (v - lo) / (hi - lo))); lut[v] = 255 * Math.pow(t, 1.6); }
  // מבט שני ורך של אותה שורה: מתיחה ישרה בלי הכהיה, כך שקווים דקים ופרטים קטנים של האותיות נשמרים
  const softLut = new Uint8ClampedArray(256);
  for (let v = 0; v < 256; v++) softLut[v] = 255 * Math.max(0, Math.min(1, (v - lo * 0.6) / (250 - lo * 0.6)));
  const gs = new Uint8ClampedArray(g.length);
  for (let k = 0; k < g.length; k++) { gs[k] = softLut[g[k]]; g[k] = lut[g[k]]; }
  // אם המחיקה של העט הופעלה, העמוד נקרא כמו שהוא מוצג
  if (p.ink) { applyFilter(c, 'clean', true); const dd = x.getImageData(0, 0, W, H).data; for (let i = 0, k = 0; k < g.length; i += 4, k++) if (dd[i] === 255 && dd[i + 1] === 255 && dd[i + 2] === 255) g[k] = gs[k] = 255; }
  // האזור שיש בו כתב: שורות ועמודות עם מספיק נקודות כהות
  // כתב = כמעט שחור אחרי ההכנה. נקודות המשבצות נשארות אפורות בהירות ולא נספרות
  const INK = 100, rows = new Uint32Array(H), cols = new Uint32Array(W);
  for (let y = 0; y < H; y++) for (let xx = 0; xx < W; xx++) if (g[y * W + xx] < INK) { rows[y]++; cols[xx]++; }
  const edge = (arr, len, min) => { let a = 0, b = len - 1; while (a < len && arr[a] < min) a++; while (b > a && arr[b] < min) b--; return [a, b]; };
  // שוליים של 2% מכל צד לא נבדקים: שם יושבים צל, ספירלה וקצה הדף
  // הספירלה של המחברת והצל בקצה לא נחשבים כתב: השורות נספרות בלי 6% מכל צד
  const mx = Math.round(W * 0.02), my = Math.round(H * 0.02), side = Math.round(W * 0.06);
  // קו שוליים לאורך כל הדף (עמודה כהה ברוב הגובה) או קו לרוחב כל הדף אינם כתב
  const vRule = new Uint8Array(W); for (let xx = 0; xx < W; xx++) vRule[xx] = cols[xx] > H * 0.4 ? 1 : 0;
  for (let xx = 1; xx < W - 1; xx++) if ((vRule[xx - 1] & 1) || (vRule[xx + 1] & 1)) vRule[xx] |= 2;
  rows.fill(0);
  for (let y = my; y < H - my; y++) for (let xx = side; xx < W - side; xx++) if (!vRule[xx] && g[y * W + xx] < INK) rows[y]++;
  const hRule = new Uint8Array(H); for (let y = 0; y < H; y++) hRule[y] = rows[y] > W * 0.35 ? 1 : 0;
  for (let y = 0; y < H; y++) if (hRule[y]) rows[y] = 0;
  // ממוצע על כמה שורות: נקודה בודדת או לכלוך לא נחשבים שורת כתב
  { const win = Math.max(3, Math.round(H * 0.008)), t = Float32Array.from(rows);
    for (let y = 0; y < H; y++) { let a = 0, n = 0; for (let k = y - win; k <= y + win; k++) if (k >= 0 && k < H) { a += t[k]; n++; } rows[y] = a / n; } }
  cols.fill(0);
  for (let y = my; y < H - my; y++) if (!hRule[y]) for (let xx = 0; xx < W; xx++) if (!(vRule[xx] & 1) && g[y * W + xx] < INK) cols[xx]++;
  for (let xx = 0; xx < mx; xx++) cols[xx] = cols[W - 1 - xx] = 0;
  let [y0, y1] = edge(rows, H, Math.max(3, W * 0.008)), [x0, x1] = edge(cols, W, Math.max(4, H * 0.004));
  if (y1 - y0 < H * 0.05 || x1 - x0 < W * 0.1) { y0 = 0; y1 = H - 1; x0 = 0; x1 = W - 1; }
  const pad = Math.round(Math.max(W, H) * 0.025);
  x0 = Math.max(0, x0 - pad); y0 = Math.max(0, y0 - pad); x1 = Math.min(W - 1, x1 + pad); y1 = Math.min(H - 1, y1 + pad);
  const cw = x1 - x0 + 1, ch = y1 - y0 + 1;
  const gray = canvasOf(cw, ch), gx = gray.getContext('2d'), gi = gx.createImageData(cw, ch), gd = gi.data;
  for (let y = 0; y < ch; y++) for (let xx = 0; xx < cw; xx++) { const v = g[(y + y0) * W + xx + x0], o = (y * cw + xx) * 4; gd[o] = gd[o + 1] = gd[o + 2] = v; gd[o + 3] = 255; }
  gx.putImageData(gi, 0, 0); free(c);
  // אזור כתב קטן (צולם מרחוק) מוגדל, כדי שהקורא יראה כל אות בגודל מספיק
  const jpeg = (src, sx, sy, sw, sh, max) => { const k = Math.min(2, max / Math.max(sw, sh)), o = canvasOf(sw * k, sh * k), ox = o.getContext('2d');
    ox.imageSmoothingQuality = 'high'; ox.drawImage(src, sx, sy, sw, sh, 0, 0, o.width, o.height); const u = o.toDataURL('image/jpeg', 0.9); const r = { u, w: o.width, h: o.height }; free(o); return r; };
  const page = jpeg(gray, 0, 0, cw, ch, fast ? 1600 : 2000), strips = [];
  // רצועות רק כשיש הרבה שורות: בכמה שורות הדף עצמו כבר חד מספיק, ותמונות נוספות רק מאטות
  if (!fast && ch > cw * 0.55) {
    // רצועות: בערך שש שורות כתב בכל אחת, עד ארבע
    const n = Math.max(2, Math.min(3, Math.round(ch / cw * 2.5))), ov = 0.08;
    for (let s = 0; s < n; s++) {
      const sy0 = Math.max(0, ch * (s / n - ov)), sy1 = Math.min(ch, ch * ((s + 1) / n + ov));
      strips.push(jpeg(gray, 0, sy0, cw, sy1 - sy0, 1600).u.split(',')[1]);
    }
  }
  free(gray);
  const sp = splitLines(g, gs, W, H, x0, y0, cw, ch);
  return { url: page.u, b64: page.u.split(',')[1], w: page.w, h: page.h, strips, lines: sp.map(l => l.a), lines2: sp.map(l => l.b) };
}
window.SIMAN_CAM = CAM;

/* ---- חלוקת הדף לשורות ----
   מודלי הקריאה קוראים כתב רהוט עברי טוב בהרבה כשמקבלים שורה אחת בכל תמונה, ולא דף שלם
   (במדד כתב היד העברי של ivrit.ai: טעות של כ-12% מהאותיות בשורות, לעומת יותר מפי שלושה בדף שלם).
   לכן הדף מיושר (תיקון הטיה קלה), נחתך לשורות לפי פיזור הדיו, וכל שורה נשלחת כתמונה נפרדת. */
function splitLines(g, gs, W, H, X0, Y0, cw, ch) {
  const INK = 110;
  // הטיה: הזווית (עד 6 מעלות) שבה שורות הדיו הכי "חדות" בהטלה האופקית
  const sk = Math.min(1, 700 / cw), pts = [];
  const sx = Math.max(1, Math.round(1 / sk));
  for (let y = 0; y < ch; y += sx) for (let x = 0; x < cw; x += sx) if (g[(y + Y0) * W + x + X0] < INK) pts.push(x, y);
  if (pts.length < 40) return [];
  let bestA = 0, bestV = -1;
  for (let a = -6; a <= 6.001; a += 0.5) {
    const t = Math.tan(a * Math.PI / 180), bins = new Float32Array(Math.ceil(ch / sx) + Math.ceil(cw * 0.12 / sx) * 2 + 4), off = Math.ceil(cw * 0.12 / sx) + 2;
    for (let i = 0; i < pts.length; i += 2) { const b = Math.round((pts[i + 1] - (pts[i] - cw / 2) * t) / sx) + off; if (b >= 0 && b < bins.length) bins[b]++; }
    let v = 0; for (const c of bins) v += c * c;
    if (v > bestV) { bestV = v; bestA = a; }
  }
  // הדף המיושר (אפור), בגודל מלא
  const straight = arr => {
    const src = canvasOf(cw, ch), sxc = src.getContext('2d'), si = sxc.createImageData(cw, ch), sd = si.data;
    for (let y = 0; y < ch; y++) for (let x = 0; x < cw; x++) { const v = arr[(y + Y0) * W + x + X0], o = (y * cw + x) * 4; sd[o] = sd[o + 1] = sd[o + 2] = v; sd[o + 3] = 255; }
    sxc.putImageData(si, 0, 0);
    if (Math.abs(bestA) < 0.5) return src;
    const o = canvasOf(cw, ch), px = o.getContext('2d'); px.fillStyle = '#fff'; px.fillRect(0, 0, cw, ch);
    px.translate(cw / 2, ch / 2); px.rotate(-bestA * Math.PI / 180); px.drawImage(src, -cw / 2, -ch / 2); free(src);
    return o;
  };
  const pg = straight(g), pg2 = straight(gs);
  const p2ctx = pg2.getContext('2d', { willReadFrequently: true }), p2im = p2ctx.getImageData(0, 0, cw, ch), d2 = p2im.data;
  const pctx = pg.getContext('2d', { willReadFrequently: true }), pim = pctx.getImageData(0, 0, cw, ch), d = pim.data;
  // קו ארוך ודק (קו תחתון, קצה שולחן, צל בשוליים) אינו כתב: כל רכיב דיו מחובר כזה נמחק,
  // כדי שלא ייחשב שורה ולא יחבר שורות
  {
    const n = cw * ch, lab = new Int32Array(n).fill(-1), stack = new Int32Array(n);
    for (let s0 = 0; s0 < n; s0++) {
      if (lab[s0] !== -1 || d[s0 * 4] >= INK) continue;
      let sp = 0, x0 = cw, x1 = -1, y0 = ch, y1 = -1; const pix = [];
      stack[sp++] = s0; lab[s0] = s0;
      while (sp) {
        const p = stack[--sp], x = p % cw, y = (p / cw) | 0; pix.push(p);
        if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y;
        for (const q of [x > 0 ? p - 1 : -1, x < cw - 1 ? p + 1 : -1, y > 0 ? p - cw : -1, y < ch - 1 ? p + cw : -1])
          if (q >= 0 && lab[q] === -1 && d[q * 4] < INK) { lab[q] = s0; stack[sp++] = q; }
      }
      const bw = x1 - x0 + 1, bh = y1 - y0 + 1;
      const rule = (bw > cw * 0.2 && bh < bw * 0.15) || ((y0 === 0 || y1 === ch - 1) && bw > cw * 0.12) || (bh > ch * 0.5 && bw < bh * 0.1);
      if (rule) for (const p of pix) { const o = p * 4; d[o] = d[o + 1] = d[o + 2] = 255; d2[o] = d2[o + 1] = d2[o + 2] = 255; }
    }
  }
  pctx.putImageData(pim, 0, 0); p2ctx.putImageData(p2im, 0, 0);
  const ink = (x, y) => d[(y * cw + x) * 4] < INK;
  // קווים אנכיים לכל אורך הדף (שוליים, ספירלה) לא נספרים
  const colN = new Uint32Array(cw);
  for (let y = 0; y < ch; y++) for (let x = 0; x < cw; x++) if (ink(x, y)) colN[x]++;
  const skip = new Uint8Array(cw); for (let x = 0; x < cw; x++) if (colN[x] > ch * 0.35) { skip[x] = 1; if (x) skip[x - 1] = 1; if (x < cw - 1) skip[x + 1] = 1; }
  const rows = new Float32Array(ch);
  const sideX = Math.round(cw * 0.05);
  for (let y = 0; y < ch; y++) { let n = 0; for (let x = sideX; x < cw - sideX; x++) if (!skip[x] && ink(x, y)) n++; rows[y] = n; }
  // פס כהה לרוחב כל הדף (קצה השולחן, צל) אינו שורת כתב
  for (let y = 0; y < ch; y++) if (rows[y] > cw * 0.3) { rows[y] = 0; for (let x = 0; x < cw; x++) if (ink(x, y)) d[(y * cw + x) * 4] = 255; }
  // גובה גוף האות: הרצועות הרציפות של דיו בפרופיל
  const sm0 = new Float32Array(ch), w0 = Math.max(1, Math.round(ch * 0.003));
  for (let y = 0; y < ch; y++) { let a = 0, n = 0; for (let k = y - w0; k <= y + w0; k++) if (k >= 0 && k < ch) { a += rows[k]; n++; } sm0[y] = a / n; }
  const pos = [...sm0].filter(v => v > 0).sort((a, b) => a - b);
  if (!pos.length) { free(pg); free(pg2); return []; }
  const top95 = pos[Math.floor(pos.length * 0.95)], hs = [];
  for (let y = 0, st = -1; y <= ch; y++) { const on = y < ch && sm0[y] > top95 * 0.15; if (on && st < 0) st = y; if (!on && st >= 0) { hs.push(y - st); st = -1; } }
  hs.sort((a, b) => a - b);
  const med = Math.max(10, hs.length ? hs[hs.length >> 1] : 20);
  // שורות = פסגות בפרופיל המוחלק; הגבול בין שתי שורות = הנקודה הכי ריקה ביניהן.
  // כך גם אותיות שיורדות מהשורה (ק ף ן) או עולות ממנה (ל) לא מחברות שתי שורות לאחת
  const r = Math.max(2, Math.round(med * 0.35)), sm = new Float32Array(ch);
  for (let y = 0; y < ch; y++) { let a = 0, n = 0; for (let k = y - r; k <= y + r; k++) if (k >= 0 && k < ch) { a += rows[k]; n++; } sm[y] = a / n; }
  let mx = 0; for (const v of sm) if (v > mx) mx = v;
  let peaks = [];
  for (let y = 1; y < ch - 1; y++) if (sm[y] >= sm[y - 1] && sm[y] > sm[y + 1] && sm[y] > mx * 0.12) peaks.push(y);
  // שתי פסגות קרובות מדי הן שורה אחת
  const pk = [];
  for (const y of peaks) { const l = pk[pk.length - 1]; if (l != null && y - l < med * 0.9) { if (sm[y] > sm[l]) pk[pk.length - 1] = y; } else pk.push(y); }
  peaks = pk;
  if (!peaks.length) { free(pg); free(pg2); return []; }
  if (peaks.length > 45) { free(pg); free(pg2); return []; }   // דף מודפס צפוף: נקרא כדף שלם
  const cuts = [];
  for (let i = 0; i < peaks.length - 1; i++) { let by = peaks[i], bv = 1e9; for (let y = peaks[i]; y <= peaks[i + 1]; y++) if (sm[y] < bv) { bv = sm[y]; by = y; } cuts.push(by); }
  // כל שורה: עד הגבול עם השכנה, אבל לא יותר משורה וחצי סביב מרכזה (כדי שלא תימשך על כל החלק הריק של הדף)
  const reach = Math.round(med * 1.25);
  let segs = peaks.map((p, i) => [Math.max(i ? cuts[i - 1] : 0, p - reach), Math.min(i < cuts.length ? cuts[i] : ch - 1, p + reach)]);
  // קטע עם מעט מאוד דיו (קו תחתון, זנב, נקודה) מצטרף לשכן הקרוב
  const mass = sg => { let m = 0; for (let y = sg[0]; y <= sg[1]; y++) m += rows[y]; return m; };
  const ms = segs.map(mass), mm = [...ms].sort((a, b) => a - b)[ms.length >> 1] || 1;
  for (let i = segs.length - 1; i >= 0 && segs.length > 1; i--) {
    if (mass(segs[i]) >= mm * 0.18) continue;
    const j = i === 0 ? 1 : i === segs.length - 1 ? i - 1 : (segs[i][0] - segs[i - 1][1] <= segs[i + 1][0] - segs[i][1] ? i - 1 : i + 1);
    segs[j] = [Math.min(segs[j][0], segs[i][0]), Math.max(segs[j][1], segs[i][1])]; segs.splice(i, 1);
  }
  // רוחב: רק העמודות שיש בהן דיו (בלי הספירלה וקווי השוליים)
  const side = Math.round(cw * 0.05), out = [];
  segs.forEach(sg => {
    const t0 = Math.max(0, sg[0] - Math.round(med * 0.12)), b0 = Math.min(ch - 1, sg[1] + Math.round(med * 0.12));
    let lx = cw, rx = -1;
    for (let y = sg[0]; y <= sg[1]; y++) for (let x = side; x < cw - side; x++) if (!skip[x] && ink(x, y)) { if (x < lx) lx = x; if (x > rx) rx = x; }
    if (rx < 0) return;
    const padX = Math.round(med * 0.8); lx = Math.max(0, lx - padX); rx = Math.min(cw - 1, rx + padX);
    const w = rx - lx + 1, hh = b0 - t0 + 1;
    // גובה שורה כ-160 נקודות: חד מספיק לכל אות, בלי תמונה ענקית
    const k = Math.max(0.5, Math.min(2.5, 160 / hh, 2200 / w));
    const o = canvasOf(w * k, hh * k), ox = o.getContext('2d'); ox.fillStyle = '#fff'; ox.fillRect(0, 0, o.width, o.height);
    ox.imageSmoothingQuality = 'high'; ox.drawImage(pg, lx, t0, w, hh, 0, 0, o.width, o.height);
    const a = o.toDataURL('image/jpeg', 0.9).split(',')[1];
    ox.drawImage(pg2, lx, t0, w, hh, 0, 0, o.width, o.height);
    out.push({ a, b: o.toDataURL('image/jpeg', 0.88).split(',')[1] }); free(o);
  });
  free(pg); free(pg2);
  return out;
}
CAM._readingPic = readingPic; CAM._find = findSheet; CAM._spread = spreadOf;
if (FILTERS.some(f => f[0] === LS.get('siman.scan.filter'))) CAM.defFilter = LS.get('siman.scan.filter');
try { const o = JSON.parse(LS.get('siman.scan.opts') || 'null'); if (o) Object.assign(CAM.opts, { size: o.size || 'auto', q: o.q || 'high', wm: '' }); } catch (_) {}

function defaultName() {
  const d = new Date(), z = n => String(n).padStart(2, '0');
  return `סריקה ${z(d.getDate())}.${z(d.getMonth() + 1)}.${d.getFullYear()} ${z(d.getHours())}-${z(d.getMinutes())}`;
}
function download(file) {
  const a = document.createElement('a'); a.href = URL.createObjectURL(file); a.download = file.name;
  document.body.append(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 5000);
}
function loadJs(src) {
  if (typeof loadScript === 'function') return loadScript(src);
  return new Promise((res, rej) => { const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => rej(new Error('הטעינה נכשלה')); document.head.append(s); });
}
function watermark(c, text) {
  const x = c.getContext('2d'), fs = Math.round(Math.max(c.width, c.height) / 22);
  x.save(); x.globalAlpha = 0.13; x.fillStyle = '#1B2A4A'; x.font = `700 ${fs}px Rubik, Heebo, Arial, sans-serif`;
  x.textAlign = 'center'; x.textBaseline = 'middle'; x.direction = 'rtl';
  x.translate(c.width / 2, c.height / 2); x.rotate(-Math.atan2(c.height, c.width));
  const step = fs * 4.2, span = Math.hypot(c.width, c.height);
  for (let y = -span / 2; y <= span / 2; y += step) x.fillText(text, 0, y);
  x.restore();
}

/* ---------------------------------------------------------------- אירועים */
root.addEventListener('click', e => {
  const b = e.target.closest('[data-a],[data-m],[data-x],[data-t],[data-f]'); if (!b || !root.contains(b)) return;
  if (b.dataset.m) return CAM.setMode(b.dataset.m);
  if (b.dataset.x) return CAM.act(b.dataset.x);
  if (b.dataset.t) return CAM.tool(b.dataset.t);
  if (b.dataset.f) return CAM.setAll({ filter: b.dataset.f });
  switch (b.dataset.a) {
    case 'close': return CAM.close();
    case 'auto': CAM.auto = !CAM.auto; b.setAttribute('aria-pressed', String(CAM.auto)); b.querySelector('span').textContent = CAM.auto ? 'אוטומטי' : 'ידני'; CAM.stable = 0; return CAM.drawQuad();
    case 'torch': return CAM.toggleTorch();
    case 'swap': if (CAM.spread) { CAM.pick = CAM.spread.side === 'right' ? 'left' : 'right'; CAM.quad = null; CAM.stable = 0; CAM.lastShot = Date.now(); CAM.setHint(CAM.pick === 'right' ? 'נבחר העמוד הימני' : 'נבחר העמוד השמאלי'); } return;
    case 'pages': CAM.pick = CAM.pick === 'both' ? 'auto' : 'both'; CAM.quad = null; CAM.stable = 0; CAM.spreadTold = 0; return CAM.drawQuad();
    case 'shoot': return CAM.shoot();
    case 'gallery': return galIn.click();
    case 'review': return CAM.show('review');
    case 'more': return CAM.show('cam');
    case 'inkall': return CAM.setAll({ ink: !CAM.pages.every(p => p.ink) });
    case 'editdone': CAM.cropping = false; return CAM.show('review');
    case 'pprev': if (CAM.cur > 0) { CAM.cur--; CAM.cropping = false; CAM.renderEdit(); } return;
    case 'pnext': if (CAM.cur < CAM.pages.length - 1) { CAM.cur++; CAM.cropping = false; CAM.renderEdit(); } return;
  }
});
$r('.scn-vf').addEventListener('pointerup', e => { if (e.pointerType === 'mouse' && e.button !== 0) return; CAM.tapPick(e); });
galIn.onchange = e => { const f = [...e.target.files]; e.target.value = ''; CAM.fromFiles(f); };
camIn.onchange = e => { const f = [...e.target.files]; e.target.value = ''; CAM.fromFiles(f); };
addEventListener('resize', () => { if (!root.hidden && CAM.view === 'cam') CAM.fit(); });
document.addEventListener('visibilitychange', () => {
  if (root.hidden || CAM.view !== 'cam') return;
  if (document.hidden) CAM.stopCam(); else CAM.startCam();
});
root.addEventListener('keydown', e => {
  if (e.key === 'Escape') { e.stopPropagation(); if (CAM.view === 'edit') { CAM.cropping = false; CAM.show('review'); } else if (CAM.view === 'review') CAM.show('cam'); else CAM.close(); }
  else if ((e.key === ' ' || e.key === 'Enter') && CAM.view === 'cam' && e.target === root) { e.preventDefault(); CAM.shoot(); }
});
// ספר בעברית: העמוד הימני קודם (ברירת מחדל). בספר בשפה משמאל לימין אפשר להחליף בלחיצה ארוכה על ״ספר״
CAM.rtlBook = true;
const bookBtn = $r('[data-m=book]'); let lp = 0;
bookBtn.addEventListener('pointerdown', () => { lp = setTimeout(() => { CAM.rtlBook = !CAM.rtlBook; CAM.setMode('book'); CAM.setHint(CAM.rtlBook ? 'ספר מימין לשמאל: העמוד הימני קודם' : 'ספר משמאל לימין: העמוד השמאלי קודם'); }, 650); });
['pointerup', 'pointerleave', 'pointercancel'].forEach(t => bookBtn.addEventListener(t, () => clearTimeout(lp)));

/* ---------------------------------------------------------------- חיבור לאפליקציה
   כל הכניסות ל״מדף למסמך״ (מסך הבית, צילום דף, חיפוש פעולה) פותחות עכשיו את הסורק.
   ״מדף למסמך״ על קובץ פתוח (SIMAN_SCAN.open('doc')) נשאר כמו שהיה */
function hook() {
  const SC = window.SIMAN_SCAN;
  if (!SC) return false;
  if (SC.openReader) return true;
  const orig = SC.open.bind(SC);
  SC.openReader = orig;
  SC.open = function (src) { if (src === 'doc') return orig(src); CAM.open(); };
  return true;
}
if (!hook()) { let n = 0; const t = setInterval(() => { if (hook() || ++n > 40) clearInterval(t); }, 250); }
})();
