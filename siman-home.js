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
/* 6.3.6: "display:grid" above beat the browser's own rule for [hidden], so a hidden card stayed on the screen */
.sh-resume[hidden]{display:none!important}
.sh-resume h3{margin:0;font-size:15px;display:flex;align-items:center;gap:8px}
.sh-resume .sh-main{display:flex;align-items:center;gap:12px;width:100%;font:inherit;text-align:start;cursor:pointer;border:0;border-radius:14px;padding:12px 14px;background:var(--marker,#FFD84D);color:#1B1C1F}
.sh-resume .sh-main b{font-size:16px;display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.sh-resume .sh-main small{font-size:13px;opacity:.8}
.sh-resume .sh-main .k{flex:none;font-size:11px;font-weight:800;border-radius:6px;padding:4px 6px;background:#1B1C1F;color:#fff}
.sh-resume .sh-main .go{margin-inline-start:auto;font-weight:800;font-size:15px;white-space:nowrap}
.sh-resume .sh-more{display:flex;flex-wrap:wrap;gap:6px}
.sh-resume .sh-row{display:flex;align-items:stretch;gap:8px}
.sh-resume .sh-row .sh-main{flex:1;min-width:0}
.sh-resume .sh-x{flex:none;font:inherit;font-size:22px;line-height:1;width:44px;border-radius:14px;border:2px solid var(--ink,#2B1660);background:var(--bar,#fff);color:var(--ink,#2B1660);cursor:pointer}
.sh-resume .sh-x:hover{background:#FDE2E2;color:#B3232A;border-color:#B3232A}
.sh-resume .sh-chip{display:inline-flex;align-items:center;border:1px solid var(--line,#D4DBE5);background:var(--desk,#F4F6FA);border-radius:999px;max-width:280px}
.sh-resume .sh-chip>button[data-t]{border:0!important;background:transparent!important}
.sh-resume .sh-more .sh-x.s{flex:none;display:grid;place-items:center;width:26px;height:26px;padding:0;font-size:17px;line-height:1;overflow:visible;text-overflow:clip;border:0;border-radius:50%;background:transparent;margin-inline-end:3px;opacity:.7}
.sh-resume .sh-more .sh-x.s:hover{opacity:1}
.sh-resume .sh-ask{border:2px solid #B3232A;background:#FDE2E2;color:#1B1C1F;border-radius:14px;padding:10px 12px;display:grid;gap:8px}
.sh-resume .sh-ask p{margin:0;font-size:14px;line-height:1.4}
.sh-resume .sh-ask b{font-weight:800;overflow-wrap:anywhere}
.sh-resume .sh-ask div{display:flex;gap:8px;flex-wrap:wrap}
.sh-resume .sh-ask button{font:inherit;font-size:14px;font-weight:700;border-radius:999px;padding:7px 14px;cursor:pointer;border:2px solid var(--ink,#2B1660);background:#fff;color:var(--ink,#2B1660)}
.sh-resume .sh-ask .sh-ask-yes{background:#B3232A;border-color:#B3232A;color:#fff}
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
  let ok = false;
  try { ok = await tabs.backFromHome(t); }
  catch (e) { console.error(e); if (typeof toast === 'function') toast('לא הצלחתי לחזור לקובץ: ' + (e && e.message || e)); }
  refresh(true); return ok;
}
// × on the card: the file is closed. With unsaved changes the card itself asks (not the browser's own
// window: on some browsers it did not close the file after "אישור")
let asking = null;
async function closeFromCard(t, sure) {
  const tabs = T(); if (!tabs) return;
  let info = {}; try { info = tabs.tabInfo(t); } catch (_) {}
  if (info.dirty && !sure) { asking = t.id; askClose(t, info); note('שאלה הוצגה בכרטיס · בתוך המסמך=' + !!card.querySelector('.sh-ask')); return; }
  asking = null;
  let ok = false;
  try { ok = await tabs.closeTab(t, { force: true }); } catch (e) { console.error(e); note('שגיאה בסגירה: ' + (e && e.message)); }
  note('סגירה: ' + ok + ' · עדיין ברשימה=' + tabs.TB.list.includes(t));
  if (ok === false && typeof toast === 'function') toast('רגע, פעולה עדיין רצה. נסו שוב בעוד רגע');
  refresh(true);
}
function askClose(t, info) {
  const box = h(`<div class="sh-ask" role="alertdialog" aria-label="סגירת קובץ">
    <p>ב"<b>${esc(info.name)}</b>" יש שינויים שעוד לא נשמרו. לסגור בלי לשמור?</p>
    <div><button type="button" class="sh-ask-no">ביטול</button><button type="button" class="sh-ask-yes">סגירה בלי לשמור</button></div></div>`);
  const row = card.querySelector(`[data-x="${t.id}"]`);
  const at = row && (row.closest('.sh-row') || row.closest('.sh-chip'));
  card.querySelectorAll('.sh-ask').forEach(x => x.remove());
  if (at) at.after(box); else card.append(box);
  box.querySelector('.sh-ask-yes').onclick = () => closeFromCard(t, true);
  box.querySelector('.sh-ask-no').onclick = () => { asking = null; refresh(true); };
  box.querySelector('.sh-ask-no').focus();
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
  const b = e.target.closest && e.target.closest('.we-brand,.xe-brand,.pe-brand,.sh-edhome'); if (!b || b.classList.contains('ed-ai')) return;
  e.preventDefault(); go();
}, true);
// (next to the logo, which always shows: at the end of the bar the buttons may run off a narrow window)
function addEditorButtons() {
  for (const sel of ['.we-brand', '.xe-brand', '.pe-brand']) {
    const br = document.querySelector(sel); if (!br || br.parentElement.querySelector(':scope>.sh-ednew')) continue;
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
// one listener for the whole card: the file is looked up in the tabs as they are NOW (not as they were when
// the card was drawn). A file that is no longer open: the card is simply drawn again without it.
card.addEventListener('click', e => {
  const b = e.target.closest && e.target.closest('[data-t],[data-x]'); if (!b || !card.contains(b)) return;
  const tabs = T(), id = b.dataset.x || b.dataset.t;
  const t = tabs && tabs.TB.list.find(x => String(x.id) === id);
  note(`הכרטיס קיבל ${b.dataset.x ? '×' : 'חזרה'} של #${id} · ${t ? 'הקובץ נמצא' : 'הקובץ לא ברשימה'}`);
  if (!t) { asking = null; lastSig = ''; refresh(true); return; }
  if (b.dataset.x) closeFromCard(t); else back(t);
});
const live = h('<button type="button" class="sh-live" aria-live="polite"><i></i><span></span></button>');
document.body.append(live);
live.onclick = () => { if (window.SIMAN_VOICE) SIMAN_VOICE.unpark(); refresh(); };
function homeShown() { const e = $('empty'); return !!e && !e.classList.contains('hidden') && getComputedStyle(e).display !== 'none'; }
let lastSig = '';
// 6.3.4: the card must never freeze. Before, one error inside refresh (every second) left the card as it was:
// a file already closed stayed on it, and its × / "חזרה" found nothing to act on
function refresh(force) {
  try { refresh1(force); }
  catch (e) {
    console.error('siman-home', e);
    try { const tabs = T(); if (!tabs || !tabs.TB.list.length || !homeShown()) card.hidden = true; } catch (_) { card.hidden = true; }
  }
}
const safeInfo = t => { try { return T().tabInfo(t) || {}; } catch (_) { return { name: 'קובץ', dirty: false }; } };
function refresh1(force) {
  try { addHeaderButton(); addEditorButtons(); } catch (e) { console.error(e); }
  const tabs = T();
  let v = null; try { v = window.SIMAN_VOICE && SIMAN_VOICE.state ? SIMAN_VOICE.state() : null; } catch (_) {}
  // the bar of a recording / a transcript in work, while its screen is put aside
  const showLive = v && v.parked && (v.rec || v.job);
  live.classList.toggle('on', !!showLive); live.classList.toggle('job', !!(showLive && !v.rec));
  if (showLive) live.querySelector('span').textContent = v.rec ? `${v.paused ? 'הקלטה מושהית' : 'מקליט'} ${clock(v.recT)} · חזרה` : 'התמלול בעבודה · חזרה';
  // the card on the home screen
  try { if (tabs && tabs.dropDead && homeShown()) tabs.dropDead(); } catch (e) { console.error(e); }
  const list = tabs ? tabs.TB.list : [];
  const on = homeShown() && (list.length || (v && v.parked));
  const home = document.querySelector('#empty .v7-home') || $('empty');
  if (!on || !home) { card.hidden = true; return; }
  if (card.parentNode !== home) home.prepend(card);
  const from = tabs && tabs.homeFrom ? tabs.homeFrom() : null;
  const main = from || list[list.length - 1];
  const sig = [list.map(t => t.id + (safeInfo(t).name)).join(','), main && main.id, v && v.parked, lastPage].join('|');
  if (asking != null && list.some(t => t.id === asking) && !card.hidden) return;   // the card is asking: leave it
  asking = null;
  if (!force && sig === lastSig && !card.hidden) return;
  lastSig = sig; card.hidden = false;
  const info = safeInfo;
  const kind = t => { const n = info(t).name || ''; return (/\.(\w+)$/.exec(n) || ['', 'pdf'])[1].toUpperCase(); };
  const X = (t, small) => `<button type="button" class="sh-x${small ? ' s' : ''}" data-x="${t.id}" title="סגירת ${esc(info(t).name)}" aria-label="סגירת ${esc(info(t).name)}">×</button>`;
  card.innerHTML = `<h3>${HOUSE.replace('<svg', '<svg width="18" height="18"')}ממשיכים מאיפה שעצרת</h3>
    ${main ? `<div class="sh-row"><button type="button" class="sh-main" data-t="${main.id}"><span class="k">${esc(kind(main))}</span><span style="min-width:0"><b>${esc(info(main).name)}</b><small>${main === from && lastPage ? 'עמוד ' + lastPage + ' · ' : ''}${info(main).dirty ? 'יש שינויים שעוד לא נשמרו · ' : ''}בדיוק איפה שהפסקת</small></span><span class="go">חזרה ›</span></button>${X(main)}</div>` : ''}
    ${v && v.parked ? `<button type="button" class="sh-main" data-v="1" style="background:#FDE2E2"><span class="k" style="background:#B3232A">🎙</span><span><b>${v.rec ? 'הקלטה פעילה' : v.job ? 'תמלול בעבודה' : 'תמלול'}</b><small>${v.rec ? clock(v.recT) + ' · ממשיכה ברקע' : esc(v.title || '')}</small></span><span class="go">חזרה ›</span></button>` : ''}
    ${list.length > 1 ? `<div class="sh-more">${list.filter(t => t !== main).map(t => `<span class="sh-chip"><button type="button" data-t="${t.id}" title="${esc(info(t).name)}">${esc(kind(t))} · ${esc(info(t).name)}</button>${X(t, 1)}</span>`).join('')}</div>` : ''}`;
  const vb = card.querySelector('[data-v]'); if (vb) vb.onclick = () => { SIMAN_VOICE.unpark(); refresh(); };
}
setInterval(() => { if (!document.hidden) refresh(); }, 1000);
addEventListener('siman:voice', () => refresh(true));
new MutationObserver(() => refresh()).observe(document.documentElement, { attributes: true, attributeFilter: ['class'] });
// an editor that opens gets its "בית" and "+ חדש" at once (not on the next tick of the timer)
new MutationObserver(() => addEditorButtons()).observe(document.body, { childList: true });

/* ---------------- keyboard ---------------- */
addEventListener('keydown', e => {
  if (e.altKey && (e.key === 'Home' || e.code === 'Home')) { e.preventDefault(); go(); }
});

/* ---------------- 6.4.1: paste a file with the mouse, anywhere on the home screen ----------------
   Ctrl+V already opened a copied file anywhere on the home screen, but the browser offers "הדבק" in the right-click
   menu only over a box you can type in — on the home screen that was only the "גוררים לכאן" box. Now a right-click
   on any empty spot of the home screen puts a small invisible paste box under the mouse, so the menu offers
   "הדבק", and the pasted file opens. (Not on a phone's long press: a box to type in would open the keyboard.) */
const catcher = h('<div class="sh-paste" contenteditable="true" spellcheck="false" aria-hidden="true" tabindex="-1"></div>');
document.body.append(catcher);
document.head.append(h(`<style>.sh-paste{position:fixed;display:none;width:64px;height:64px;margin:-32px 0 0 -32px;opacity:0;z-index:2147483000;caret-color:transparent;color:transparent;outline:0;overflow:hidden;white-space:nowrap;font-size:1px}.sh-paste.on{display:block}</style>`));
const onHome = () => homeShown() && !(typeof S !== 'undefined' && S && S.doc);
const catcherOff = () => { catcher.classList.remove('on'); catcher.textContent = ''; };
const catchAt = e => {
  if (!onHome()) return;
  const t = e.target;
  if (t === catcher || (t.closest && t.closest('input,textarea,select,[contenteditable="true"],a[href],dialog,.sh-resume,.nb-field'))) return;
  catcher.style.left = e.clientX + 'px'; catcher.style.top = e.clientY + 'px';
  catcher.classList.add('on'); catcher.focus({ preventScroll: true });
  // the press itself would move the focus to the page right after this (and the box closed on blur, so Edge's menu
  // showed no "הדבק" — seen in a real Edge window): the press's own focus move is cancelled on this empty spot
  if (e.preventDefault) e.preventDefault();
  catchAt.at = Date.now();
};
addEventListener('mousedown', e => { if (e.button === 2) catchAt(e); else if (e.target !== catcher) catcherOff(); }, true);
// if the focus still moves away during the right-click itself, it is brought back before the menu is built
addEventListener('contextmenu', () => { if (catcher.classList.contains('on') && document.activeElement !== catcher) catcher.focus({ preventScroll: true }); }, true);
// a blur right at the press (the menu opening) does not close the box; a later one (the menu closed, a click
// elsewhere) does
catcher.addEventListener('blur', () => { if (Date.now() - (catchAt.at || 0) < 1500) return; setTimeout(catcherOff, 0); });
addEventListener('keydown', e => { if (e.key === 'Escape' && catcher.classList.contains('on')) catcherOff(); }, true);
catcher.addEventListener('keydown', e => { if (!(e.ctrlKey || e.metaKey)) { e.preventDefault(); catcherOff(); } });
catcher.addEventListener('beforeinput', e => { if (e.inputType !== 'insertFromPaste') e.preventDefault(); });
catcher.addEventListener('paste', e => {
  e.preventDefault(); e.stopPropagation();
  const f = e.clipboardData && e.clipboardData.files && e.clipboardData.files[0];
  catcherOff(); catcher.blur();
  if (f) { const open = window.openFile || (typeof openFile === 'function' ? openFile : null); if (open) open(f.name ? f : new File([f], 'הדבקה' + ({ 'image/png': '.png', 'image/jpeg': '.jpg', 'application/pdf': '.pdf' }[f.type] || ''), { type: f.type })); }
  else if (typeof toast === 'function') toast('אין קובץ בהעתקה. מעתיקים קובץ (בסייר הקבצים: לחיצה ימנית ← העתקה) ומדביקים כאן');
});

// "הפרטים שלי" in the settings too, so they can be prepared before any form
{ const note = document.querySelector('#setDlg .fmt-note');
  if (note) { const b = h('<button type="button" class="pro-link" style="display:block;margin:6px 0">👤 הפרטים שלי ושל המשפחה (למילוי אוטומטי של טפסים)</button>');
    b.onclick = () => { const d = $('setDlg'); if (d && d.open) d.close(); if (window.SIMAN_FILL && SIMAN_FILL.profile) SIMAN_FILL.profile(); }; note.after(b); } }

/* ---------------- 6.3.5: a check report, without the browser's developer tools ----------------
   Ctrl K → "דוח בדיקה": what the home card shows, the open files as the app sees them, what lies over the ×,
   and the last clicks (where they landed and whether the card received them). Nothing is sent anywhere. */
const CLICKS = [];
const desc = el => { if (!el || el.nodeType !== 1) return String(el && el.nodeName); const c = typeof el.className === 'string' ? el.className.trim().split(/\s+/).slice(0, 2).join('.') : ''; return el.tagName.toLowerCase() + (el.id ? '#' + el.id : '') + (c ? '.' + c : ''); };
for (const ev of ['pointerdown', 'click']) addEventListener(ev, e => {
  const t = e.target, inCard = !!(t && t.closest && t.closest('.sh-resume'));
  CLICKS.push(`${new Date().toLocaleTimeString('he-IL')} ${ev} ${Math.round(e.clientX)},${Math.round(e.clientY)} → ${desc(t)}${inCard ? ' [בכרטיס]' : ''}`);
  if (CLICKS.length > 14) CLICKS.shift();
}, true);
const NOTES = [];
const note = s => { NOTES.push(new Date().toLocaleTimeString('he-IL') + ' ' + s); if (NOTES.length > 14) NOTES.shift(); };
function reportText() {
  const tabs = T(), out = [];
  out.push('גרסה: ' + ((document.querySelector('.brand .ver') || {}).textContent || '?') + ' · ' + navigator.userAgent.replace(/^.*?\) /, ''));
  out.push('חלון: ' + innerWidth + '×' + innerHeight + (matchMedia('(display-mode: standalone)').matches ? ' · מותקנת' : ' · דפדפן'));
  try {
    out.push('קבצים פתוחים: ' + (tabs ? tabs.TB.list.length : 'אין') + (tabs && tabs.TB.cur ? ' · בחזית: ' + tabs.TB.cur.id : ' · בחזית: אין'));
    if (tabs) for (const t of tabs.TB.list) { const i = safeInfo(t); out.push(`  #${t.id} ${t.ed || 'desk'} ${i.name} · שינויים=${!!i.dirty} · פתוח=${!!(t.liveEl && document.body.contains(t.liveEl))} · שמור בצד=${!!t.data}`); }
  } catch (e) { out.push('שגיאה בקבצים: ' + e.message); }
  const cards = document.querySelectorAll('.sh-resume');
  out.push('כרטיסים: ' + cards.length + ' · מוסתר=' + card.hidden + ' · במסמך=' + document.body.contains(card) + ' · שאלה פתוחה=' + !!card.querySelector('.sh-ask') + ' · asking=' + asking);
  for (const x of card.querySelectorAll('[data-x],[data-t]')) {
    const r = x.getBoundingClientRect(), top = r.width ? document.elementFromPoint(r.x + r.width / 2, r.y + r.height / 2) : null;
    out.push(`  ${x.dataset.x ? '×' : 'חזרה'} של #${x.dataset.x || x.dataset.t} · מקום ${Math.round(r.x)},${Math.round(r.y)} ${Math.round(r.width)}×${Math.round(r.height)} · מעליו: ${desc(top)}${top && (top === x || x.contains(top)) ? ' (הכפתור עצמו)' : ' ⚠'}`);
  }
  try { out.push('עסוק: ' + !$('busy').classList.contains('hidden') + ' · מסך הבית מוצג: ' + homeShown()); } catch (_) {}
  out.push('לחיצות אחרונות:'); for (const c of CLICKS) out.push('  ' + c);
  out.push('פעולות הכרטיס:'); for (const n of NOTES) out.push('  ' + n);
  return out.join('\n');
}
function report() {
  const text = reportText();
  const d = h(`<dialog class="sh-rep" dir="rtl"><h3>דוח בדיקה</h3><p>מצלמים את החלון הזה ושולחים. לא נשלח שום דבר לשום מקום.</p><pre></pre><div><button type="button" data-copy>העתקה</button><button type="button" data-close>סגירה</button></div></dialog>`);
  d.querySelector('pre').textContent = text;
  d.querySelector('[data-close]').onclick = () => { d.close(); d.remove(); };
  d.querySelector('[data-copy]').onclick = async () => { try { await navigator.clipboard.writeText(text); d.querySelector('[data-copy]').textContent = 'הועתק ✓'; } catch (_) {} };
  d.addEventListener('close', () => d.remove());
  document.body.append(d); d.showModal();
}
document.head.append(h(`<style>.sh-rep{max-width:min(860px,94vw);width:100%;border:2px solid #2B1660;border-radius:16px;padding:16px 18px;font:14px/1.45 "SF-rubik","Rubik","Segoe UI",Arial,sans-serif;color:#16213A}
.sh-rep h3{margin:0 0 4px}.sh-rep p{margin:0 0 8px;opacity:.75}
.sh-rep pre{unicode-bidi:plaintext;white-space:pre-wrap;word-break:break-word;font:12px/1.5 Consolas,monospace;background:#F4F6FA;border:1px solid #D4DBE5;border-radius:10px;padding:10px;max-height:60vh;overflow:auto;margin:0 0 10px}
.sh-rep pre{direction:rtl;text-align:right}
.sh-rep div{display:flex;gap:8px}.sh-rep button{font:inherit;font-weight:700;border:2px solid #2B1660;background:#fff;border-radius:999px;padding:6px 16px;cursor:pointer}</style>`));

window.SIMAN_HOME = { go, back, canBack, refresh, report, reportText };
refresh();
})();
