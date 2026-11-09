/* =====================================================================
   סימן 6.3.1 — siman-home.js: מסך הבית מכל מקום, וחזרה בדיוק לאותה נקודה; "חדש" בתוך עורכי Office
   · כפתור "בית" בסרגל העליון (וגם לחיצה על הלוגו "סימן", גם בעורכי Word / Excel / PowerPoint,
     ו-Alt+Home במחשב, ו-Ctrl K ← "מסך הבית").
   · שום דבר לא נסגר: הקובץ שבחזית נשמר בצד כמו במעבר בין לשוניות — העמוד, המקום בדף, הסימונים,
     הביטול, התרגום ושינויים שלא נשמרו. עורך Word / Excel / מצגת נשמר כקובץ שלו.
   · במסך הבית מופיע "ממשיכים מאיפה שעצרת": הקובץ שעזבת (ועוד קבצים פתוחים) בלחיצה אחת חזרה.
   · הקלטה שרצה ממשיכה ברקע. פס קטן למעלה ("● מקליט 3:12 · חזרה") מחזיר אליה מכל מקום.
   · כפתור / החלקת "חזרה" במסך הבית מחזירים לקובץ שעזבת.
   ===================================================================== */
(() => {
'use strict';
if (window.SIMAN_HOME) return;
const $ = id => document.getElementById(id);
const h = html => { const t = document.createElement('template'); t.innerHTML = html.trim(); return t.content.firstElementChild; };
const esc = t => String(t ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
const T = () => window.SIMAN_TABS;
const clock = s => { s = Math.max(0, Math.floor(s || 0)); const m = Math.floor(s / 60), x = s % 60; return (m >= 60 ? Math.floor(m / 60) + ':' + String(m % 60).padStart(2, '0') : m) + ':' + String(x).padStart(2, '0'); };
const HOUSE = '<svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 11l9-7 9 7"/><path d="M5 10v10h5v-6h4v6h5V10"/></svg>';

document.head.append(h(`<style>
.sh-home{display:none;align-items:center;gap:6px;font:inherit;font-size:14px;font-weight:700;cursor:pointer;border-radius:999px;padding:6px 12px;border:2px solid currentColor;background:transparent;color:inherit;margin-inline-start:10px}
.sh-home svg{width:18px;height:18px}
/* on the home screen the bars of the file that was put aside are not shown */
html.v7:not(.v7-doc) .fzbar,html.v7:not(.v7-doc) .fzundo,html.v7:not(.v7-doc) .fzp{display:none!important}
html.v7-doc .sh-home{display:inline-flex}
header.bar .brand b{cursor:pointer}
.we-brand,.xe-brand,.pe-brand{cursor:pointer}
.sh-edhome{display:inline-flex;align-items:center;gap:5px;font:inherit;font-size:13px;font-weight:700;background:rgba(255,255,255,.12);color:#fff;border:1px solid rgba(255,255,255,.35);border-radius:999px;padding:5px 11px;cursor:pointer;margin-inline-end:6px}
.sh-edhome svg{width:16px;height:16px}
.sh-ednew{display:inline-flex;align-items:center;gap:4px;font:inherit;font-size:13px;font-weight:700;background:#fff;color:#1B1C1F;border:1px solid rgba(255,255,255,.6);border-radius:999px;padding:5px 11px;cursor:pointer;margin-inline-end:6px}
.sh-ednew svg{width:15px;height:15px}
.we-top .sh-ednew,.xe-top .sh-ednew,.pe-top .sh-ednew,.sh-ednew span{color:#2B1660!important;background:#fff!important}
.sh-newmenu{position:fixed;z-index:2147483000;direction:rtl;text-align:start;background:#fff;color:#16213A;border:2px solid #2B1660;border-radius:14px;box-shadow:4px 4px 0 #2B1660;padding:8px;display:grid;gap:2px;font:500 14px/1.3 "SF-rubik","Rubik","Segoe UI",Arial,sans-serif}
.sh-newmenu .sh-nm-h{font-size:12px;font-weight:700;opacity:.65;padding:2px 8px 6px}
.sh-newmenu button{display:flex;align-items:center;gap:10px;font:inherit;text-align:start;background:transparent;border:0;border-radius:9px;padding:9px 10px;cursor:pointer;color:inherit}
.sh-newmenu button:hover,.sh-newmenu button:focus-visible{background:#F1ECFF;outline:none}
.sh-newmenu button i{font-style:normal;font-size:18px;width:22px;text-align:center}
.sh-newmenu hr{border:0;border-top:1px solid #E3E6EE;margin:4px 6px}
.sh-newmenu small{font-size:11.5px;opacity:.6;padding:4px 8px 2px}
/* the editors' top bar: on a window that is not wide, the last buttons (× among them) ran off the screen.
   Now they go to a second line instead */
.we-top,.xe-top,.pe-top{flex-wrap:wrap;row-gap:6px}
.sh-resume{position:relative;z-index:3;max-width:860px;margin:14px auto 4px;padding:14px 16px;border-radius:18px;background:var(--bar,#fff);color:var(--ink,#16213A);border:2px solid var(--ink,#2B1660);box-shadow:4px 4px 0 var(--ink,#2B1660);display:grid;gap:10px;direction:rtl;text-align:start}
.sh-resume h3{margin:0;font-size:15px;display:flex;align-items:center;gap:8px}
.sh-resume .sh-main{display:flex;align-items:center;gap:12px;width:100%;font:inherit;text-align:start;cursor:pointer;border:0;border-radius:14px;padding:12px 14px;background:var(--marker,#FFD84D);color:#1B1C1F}
.sh-resume .sh-main b{font-size:16px;display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.sh-resume .sh-main small{font-size:13px;opacity:.8}
.sh-resume .sh-main .k{flex:none;font-size:11px;font-weight:800;border-radius:6px;padding:4px 6px;background:#1B1C1F;color:#fff}
.sh-resume .sh-main .go{margin-inline-start:auto;font-weight:800;font-size:15px;white-space:nowrap}
.sh-resume .sh-more{display:flex;flex-wrap:wrap;gap:6px}
.sh-resume .sh-more button{font:inherit;font-size:13px;border:1px solid var(--line,#D4DBE5);background:var(--desk,#F4F6FA);color:inherit;border-radius:999px;padding:5px 12px;cursor:pointer;max-width:240px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.sh-resume .sh-more button:hover,.sh-resume .sh-main:hover{filter:brightness(.97)}
.sh-live{position:fixed;z-index:2147482990;top:calc(8px + env(safe-area-inset-top,0px));left:50%;transform:translateX(-50%);display:none;align-items:center;gap:8px;font:600 14px/1 "SF-rubik","Rubik","Segoe UI",Arial,sans-serif;direction:rtl;background:#B3232A;color:#fff;border:0;border-radius:999px;padding:9px 16px;box-shadow:0 6px 20px rgba(0,0,0,.25);cursor:pointer}
.sh-live.on{display:inline-flex}
.sh-live i{width:9px;height:9px;border-radius:50%;background:#fff;animation:shBlink 1s steps(2) infinite}
.sh-live.job{background:#2E58CC}.sh-live.job i{animation:none;border:2px solid #fff;background:transparent;width:8px;height:8px;border-top-color:transparent;animation:shSpin .9s linear infinite}
@keyframes shBlink{50%{opacity:.2}}@keyframes shSpin{to{transform:rotate(360deg)}}
@media (max-width:720px){.sh-home span{display:none}.sh-home{padding:6px;margin-inline-start:6px}.sh-resume{margin:10px 10px 4px}.sh-edhome span{display:none}.sh-ednew span{display:none}.sh-ednew{padding:5px 8px}}
@media (prefers-reduced-motion:reduce){.sh-live i{animation:none}}
</style>`));

/* ---------------- going home, and back ---------------- */
let lastPage = 0;
async function go() {
  const tabs = T();
  // a file or an editor in front: put aside like a tab switch
  if (tabs && tabs.TB.cur) {
    try { lastPage = typeof S !== 'undefined' && S && S.doc && typeof currentPageIndex === 'function' ? currentPageIndex() + 1 : 0; } catch (_) { lastPage = 0; }
    if (window.SIMAN_FILL && SIMAN_FILL.close) try { SIMAN_FILL.close(); } catch (_) {}
    document.querySelectorAll('dialog[open]').forEach(d => { try { d.close(); } catch (_) {} });
    const ok = await tabs.goHome(); if (!ok) return false;
  }
  // the transcription screen: put aside (a recording or a transcript in work goes on)
  const v = window.SIMAN_VOICE && SIMAN_VOICE.state && SIMAN_VOICE.state();
  if (v && v.open) SIMAN_VOICE.park();
  refresh(true);
  try { ($('desk') || document.scrollingElement).scrollTop = 0; } catch (_) {}
  return true;
}
async function back(t) {
  const tabs = T(); if (!tabs) return false;
  const ok = await tabs.backFromHome(t);
  refresh(); return ok;
}
const canBack = () => { const tabs = T(); return !!(tabs && !tabs.TB.cur && tabs.homeFrom && tabs.homeFrom()); };

/* ---------------- the buttons ---------------- */
function addHeaderButton() {
  const bar = document.querySelector('header.bar'), brand = bar && bar.querySelector('.brand');
  if (!brand || bar.querySelector('.sh-home')) return;
  const b = h(`<button type="button" class="sh-home" title="מסך הבית (Alt+Home). הקובץ נשאר פתוח, וחוזרים אליו בדיוק לאותו מקום">${HOUSE}<span>בית</span></button>`);
  b.onclick = () => go();
  brand.after(b);
  const logo = brand.querySelector('b'); if (logo) { logo.title = 'מסך הבית'; logo.addEventListener('click', () => { if (T() && T().TB.cur) go(); }); }
}
// the editors (Word / Excel / PowerPoint): the logo goes home, and a "בית" button next to ×
document.addEventListener('click', e => {
  const b = e.target.closest && e.target.closest('.we-brand,.xe-brand,.pe-brand,.sh-edhome'); if (!b) return;
  e.preventDefault(); go();
}, true);
// (next to the logo, which always shows: at the end of the bar the buttons may run off a narrow window)
function addEditorButtons() {
  for (const sel of ['.we-brand', '.xe-brand', '.pe-brand']) {
    const br = document.querySelector(sel); if (!br || (br.nextElementSibling && br.nextElementSibling.classList.contains('sh-edhome'))) continue;
    const kind = sel === '.we-brand' ? 'docx' : sel === '.xe-brand' ? 'xlsx' : 'pptx';
    const home = h(`<button type="button" class="sh-edhome" title="מסך הבית. הקובץ נשאר פתוח, וחוזרים אליו מאותו מקום">${HOUSE}<span>בית</span></button>`);
    const nw = h(`<button type="button" class="sh-ednew" data-kind="${kind}" aria-haspopup="menu" title="קובץ חדש (Alt+N). הקובץ הזה נשאר פתוח בלשונית שלו">${PLUS}<span>חדש</span></button>`);
    br.after(home, nw);
  }
}

/* 6.3.1: a new clean file from inside the editor, as in Office (File → New): it opens in a tab of its own,
   and the file you were in stays open in its tab, exactly where you left it. */
const PLUS = '<svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>';
const NEW_KINDS = {
  docx: { ico: '📝', label: 'מסמך Word ריק', base: 'מסמך חדש', ext: 'docx', fn: () => window.openWordEditor || (typeof openWordEditor === 'function' && openWordEditor) },
  xlsx: { ico: '📊', label: 'גיליון Excel ריק', base: 'גיליון חדש', ext: 'xlsx', fn: () => window.openExcelEditor || (typeof openExcelEditor === 'function' && openExcelEditor) },
  pptx: { ico: '📽️', label: 'מצגת PowerPoint ריקה', base: 'מצגת חדשה', ext: 'pptx', fn: () => window.openPptEditor || (typeof openPptEditor === 'function' && openPptEditor) },
};
// "מסמך חדש.docx", then "מסמך חדש 2.docx"… so two new files do not look the same in the tabs or when saved
function freeName(k) {
  const K = NEW_KINDS[k], tabs = T(), taken = new Set();
  try { for (const t of tabs.TB.list) taken.add(String(tabs.tabInfo(t).name || '').replace(/^•\s*/, '')); } catch (_) {}
  let n = 1, name = `${K.base}.${K.ext}`;
  while (taken.has(name)) name = `${K.base} ${++n}.${K.ext}`;
  return name;
}
function newFile(k) {
  const K = NEW_KINDS[k], open = K && K.fn();
  if (!open) return;
  closeNewMenu();
  open(null, freeName(k), true);
}
let newMenu = null;
function closeNewMenu() { if (newMenu) { newMenu.remove(); newMenu = null; } }
function openNewMenu(btn) {
  if (newMenu) { closeNewMenu(); return; }
  const cur = btn.dataset.kind || 'docx';
  // the kind you are in comes first: one tap for "another one like this"
  const order = [cur, ...Object.keys(NEW_KINDS).filter(x => x !== cur)];
  newMenu = h(`<div class="sh-newmenu" role="menu" aria-label="קובץ חדש">
    <div class="sh-nm-h">קובץ חדש בלשונית נוספת</div>
    ${order.map(k => `<button type="button" role="menuitem" data-new="${k}"><i>${NEW_KINDS[k].ico}</i>${NEW_KINDS[k].label}</button>`).join('')}
    <hr><button type="button" role="menuitem" data-new="open"><i>📂</i>פתיחת קובץ מהמכשיר…</button>
    <small>הקובץ שפתוח עכשיו נשאר בלשונית שלו</small></div>`);
  document.body.append(newMenu);
  const r = btn.getBoundingClientRect(), W = Math.min(280, innerWidth - 16);
  newMenu.style.width = W + 'px';
  newMenu.style.top = Math.round(r.bottom + 6) + 'px';
  newMenu.style.left = Math.round(Math.max(8, Math.min(innerWidth - W - 8, r.right - W))) + 'px';
  newMenu.onclick = e => {
    const b = e.target.closest('[data-new]'); if (!b) return;
    if (b.dataset.new === 'open') { closeNewMenu(); if (typeof pickFile === 'function') pickFile(); else if ($('file')) $('file').click(); return; }
    newFile(b.dataset.new);
  };
  newMenu.addEventListener('keydown', e => {
    const items = [...newMenu.querySelectorAll('[data-new]')], i = items.indexOf(document.activeElement);
    if (e.key === 'Escape') { closeNewMenu(); btn.focus(); }
    else if (e.key === 'ArrowDown') { e.preventDefault(); items[(i + 1) % items.length].focus(); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); items[(i - 1 + items.length) % items.length].focus(); }
  });
  newMenu.querySelector('[data-new]').focus();
}
document.addEventListener('click', e => {
  const b = e.target.closest && e.target.closest('.sh-ednew');
  if (b) { e.preventDefault(); e.stopPropagation(); openNewMenu(b); return; }
  if (newMenu && !newMenu.contains(e.target)) closeNewMenu();
}, true);
addEventListener('resize', closeNewMenu);
// Alt+N in an editor: the menu (Ctrl+N belongs to the browser: it opens a new window and cannot be taken)
addEventListener('keydown', e => {
  if (!(e.altKey && !e.ctrlKey && !e.metaKey && (e.code === 'KeyN'))) return;
  const b = document.querySelector('.sh-ednew'); if (!b || !b.offsetParent) return;
  e.preventDefault(); openNewMenu(b);
}, true);

/* ---------------- "continue where you stopped" on the home screen ---------------- */
const card = h('<section class="sh-resume" aria-label="ממשיכים מאיפה שעצרת" hidden></section>');
const live = h('<button type="button" class="sh-live" aria-live="polite"><i></i><span></span></button>');
document.body.append(live);
live.onclick = () => { if (window.SIMAN_VOICE) SIMAN_VOICE.unpark(); refresh(); };
function homeShown() { const e = $('empty'); return !!e && !e.classList.contains('hidden') && getComputedStyle(e).display !== 'none'; }
let lastSig = '';
function refresh(force) {
  addHeaderButton(); addEditorButtons();
  const tabs = T(), v = window.SIMAN_VOICE && SIMAN_VOICE.state ? SIMAN_VOICE.state() : null;
  // the bar of a recording / a transcript in work, while its screen is put aside
  const showLive = v && v.parked && (v.rec || v.job);
  live.classList.toggle('on', !!showLive); live.classList.toggle('job', !!(showLive && !v.rec));
  if (showLive) live.querySelector('span').textContent = v.rec ? `${v.paused ? 'הקלטה מושהית' : 'מקליט'} ${clock(v.recT)} · חזרה` : 'התמלול בעבודה · חזרה';
  // the card on the home screen
  const list = tabs ? tabs.TB.list : [];
  const on = homeShown() && (list.length || (v && v.parked));
  const home = document.querySelector('#empty .v7-home') || $('empty');
  if (!on || !home) { card.hidden = true; return; }
  if (card.parentNode !== home) home.prepend(card);
  const from = tabs && tabs.homeFrom ? tabs.homeFrom() : null;
  const main = from || list[list.length - 1];
  const sig = [list.map(t => t.id + (tabs.tabInfo(t).name)).join(','), main && main.id, v && v.parked, lastPage].join('|');
  if (!force && sig === lastSig && !card.hidden) return;
  lastSig = sig; card.hidden = false;
  const info = t => tabs.tabInfo(t);
  const kind = t => { const n = info(t).name || ''; return (/\.(\w+)$/.exec(n) || ['', 'pdf'])[1].toUpperCase(); };
  card.innerHTML = `<h3>${HOUSE.replace('<svg', '<svg width="18" height="18"')}ממשיכים מאיפה שעצרת</h3>
    ${main ? `<button type="button" class="sh-main" data-t="${main.id}"><span class="k">${esc(kind(main))}</span><span style="min-width:0"><b>${esc(info(main).name)}</b><small>${main === from && lastPage ? 'עמוד ' + lastPage + ' · ' : ''}${info(main).dirty ? 'יש שינויים שעוד לא נשמרו · ' : ''}בדיוק איפה שהפסקת</small></span><span class="go">חזרה ›</span></button>` : ''}
    ${v && v.parked ? `<button type="button" class="sh-main" data-v="1" style="background:#FDE2E2"><span class="k" style="background:#B3232A">🎙</span><span><b>${v.rec ? 'הקלטה פעילה' : v.job ? 'תמלול בעבודה' : 'תמלול'}</b><small>${v.rec ? clock(v.recT) + ' · ממשיכה ברקע' : esc(v.title || '')}</small></span><span class="go">חזרה ›</span></button>` : ''}
    ${list.length > 1 ? `<div class="sh-more">${list.filter(t => t !== main).map(t => `<button type="button" data-t="${t.id}" title="${esc(info(t).name)}">${esc(kind(t))} · ${esc(info(t).name)}</button>`).join('')}</div>` : ''}`;
  card.querySelectorAll('[data-t]').forEach(b => b.onclick = () => { const t = list.find(x => String(x.id) === b.dataset.t); if (t) back(t); });
  const vb = card.querySelector('[data-v]'); if (vb) vb.onclick = () => { SIMAN_VOICE.unpark(); refresh(); };
}
setInterval(() => { if (!document.hidden) refresh(); }, 1000);
addEventListener('siman:voice', () => refresh(true));
new MutationObserver(() => refresh()).observe(document.documentElement, { attributes: true, attributeFilter: ['class'] });

/* ---------------- keyboard ---------------- */
addEventListener('keydown', e => {
  if (e.altKey && (e.key === 'Home' || e.code === 'Home')) { e.preventDefault(); go(); }
});

// "הפרטים שלי" in the settings too, so they can be prepared before any form
{ const note = document.querySelector('#setDlg .fmt-note');
  if (note) { const b = h('<button type="button" class="pro-link" style="display:block;margin:6px 0">👤 הפרטים שלי ושל המשפחה (למילוי אוטומטי של טפסים)</button>');
    b.onclick = () => { const d = $('setDlg'); if (d && d.open) d.close(); if (window.SIMAN_FILL && SIMAN_FILL.profile) SIMAN_FILL.profile(); }; note.after(b); } }

window.SIMAN_HOME = { go, back, canBack, refresh };
refresh();
})();
