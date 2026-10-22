/* סימן — תפריט נגישות (a11y.js)
   One file for every page of the app (index, sign, about, privacy, terms, install, accessibility).
   Settings are saved in this browser (localStorage 'siman.a11y') and apply on every page.
   Documents themselves (PDF pages, Word/Excel/PowerPoint editing surfaces) are never resized,
   so marking and signing positions stay exact. */
(function () {
  'use strict';
  if (window.__simanA11y) return; window.__simanA11y = true;

  const KEY = 'siman.a11y';
  const SIZES = [1, 1.15, 1.3, 1.5];
  const DEF = { size: 0, contrast: false, gray: false, links: false, font: false, spacing: false, motion: false, cursor: false, focus: false };
  let st = Object.assign({}, DEF);
  try { Object.assign(st, JSON.parse(localStorage.getItem(KEY) || '{}')); } catch (_) {}
  const save = () => { try { localStorage.setItem(KEY, JSON.stringify(st)); } catch (_) {} };

  // ---------- styles ----------
  const css = `
.a11y-panel{--a-bg:#FFFFFF;--a-ink:#16213A;--a-mute:#5D6982;--a-line:#D4DBE5;--a-tile:#F2F4F8;--a-on:#16213A;--a-on-ink:#FFD43B;--a-acc:#2E58CC}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]) .a11y-panel{--a-bg:#1B2436;--a-ink:#E4EAF4;--a-mute:#9AA6BB;--a-line:#2D3950;--a-tile:#232E44;--a-on:#FFD43B;--a-on-ink:#16213A;--a-acc:#86A6FF}}
:root[data-theme="dark"] .a11y-panel{--a-bg:#1B2436;--a-ink:#E4EAF4;--a-mute:#9AA6BB;--a-line:#2D3950;--a-tile:#232E44;--a-on:#FFD43B;--a-on-ink:#16213A;--a-acc:#86A6FF}
.a11y-panel{position:fixed;z-index:2147483000;width:min(340px,calc(100vw - 20px));max-height:calc(100vh - 90px);overflow:auto;direction:rtl;text-align:right;
  background:var(--a-bg);color:var(--a-ink);border:1px solid var(--a-line);border-radius:16px;box-shadow:0 18px 48px -12px rgba(10,16,30,.45);
  padding:16px;font:15px/1.45 "Segoe UI",Arial,sans-serif;box-sizing:border-box}
.a11y-panel[hidden]{display:none}
.a11y-panel *{box-sizing:border-box;font-family:inherit}
.a11y-head{display:flex;align-items:center;justify-content:space-between;gap:10px;margin:0 0 12px}
.a11y-head h2{margin:0;font-size:18px;font-weight:700;color:var(--a-ink)}
.a11y-x{border:0;background:transparent;color:var(--a-ink);width:34px;height:34px;border-radius:9px;cursor:pointer;font-size:20px;line-height:1}
.a11y-x:hover{background:var(--a-tile)}
.a11y-size{display:flex;align-items:center;gap:8px;background:var(--a-tile);border-radius:12px;padding:8px 10px;margin:0 0 10px}
.a11y-size span{flex:1;font-weight:600}
.a11y-size output{min-width:44px;text-align:center;font-weight:700}
.a11y-size button{width:38px;height:38px;border-radius:10px;border:1px solid var(--a-line);background:var(--a-bg);color:var(--a-ink);font-size:20px;font-weight:700;cursor:pointer}
.a11y-size button:disabled{opacity:.4;cursor:default}
.a11y-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.a11y-t{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:6px;min-height:78px;padding:10px 6px;border-radius:12px;
  border:2px solid transparent;background:var(--a-tile);color:var(--a-ink);cursor:pointer;font-size:14px;font-weight:600;text-align:center}
.a11y-t svg{width:24px;height:24px}
.a11y-t[aria-pressed="true"]{background:var(--a-on);color:var(--a-on-ink);border-color:var(--a-on)}
.a11y-t:hover{border-color:var(--a-acc)}
.a11y-foot{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-top:12px;padding-top:12px;border-top:1px solid var(--a-line)}
.a11y-reset{border:1px solid var(--a-line);background:transparent;color:var(--a-ink);border-radius:10px;padding:8px 12px;font-weight:600;cursor:pointer}
.a11y-foot a{color:var(--a-acc);font-weight:600}
.a11y-panel button:focus-visible,.a11y-panel a:focus-visible,.a11y-fab:focus-visible{outline:3px solid var(--a-acc);outline-offset:2px}
.a11y-fab{position:fixed;z-index:2147482999;left:16px;bottom:16px;width:52px;height:52px;border-radius:50%;border:0;cursor:pointer;
  background:#16213A;color:#FFD43B;box-shadow:0 8px 22px -6px rgba(10,16,30,.5);display:flex;align-items:center;justify-content:center;
  margin-bottom:env(safe-area-inset-bottom,0px)}
.a11y-fab svg{width:28px;height:28px}
@media print{.a11y-panel,.a11y-fab{display:none!important}}

/* ---- the modes ---- */
html.a11y-contrast.a11y-contrast:root{--navy:#000;--navy-2:#000;--ink:#000;--muted:#1d1d1d;--line:#000;--bar:#fff;--desk:#fff;--chrome-text:#fff;--chrome-mute:#fff;--blue:#0033B3}
html.a11y-contrast body :is(button,input,select,textarea,.btn,.icon-btn):not(.a11y-t):not(.a11y-fab){border:2px solid currentColor!important}
html.a11y-gray{filter:grayscale(1)}
html.a11y-links a:not(.a11y-panel a){text-decoration:underline!important;text-underline-offset:3px;outline:2px dashed currentColor;outline-offset:2px}
html.a11y-font body :not(#pages):not(#pages *):not(#wed *):not(#xed *):not([contenteditable] *):not(svg):not(svg *){font-family:Arial,"Segoe UI",sans-serif!important;letter-spacing:0!important}
html.a11y-spacing body :is(p,li,label,summary,td,th,dd,dt,h1,h2,h3,button):not(#pages *):not(#wed *):not(#xed *){line-height:1.8!important;letter-spacing:.04em!important;word-spacing:.12em!important}
html.a11y-motion *,html.a11y-motion *::before,html.a11y-motion *::after{animation:none!important;transition:none!important;scroll-behavior:auto!important}
html.a11y-cursor,html.a11y-cursor *{cursor:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='40' height='40' viewBox='0 0 24 24'%3E%3Cpath d='M5 2l14 10.5-6.4 1.4L16 21.5l-3 1.3-3.4-7.7L5 19z' fill='%23000' stroke='%23fff' stroke-width='1.4' stroke-linejoin='round'/%3E%3C/svg%3E") 6 3,auto!important}
html.a11y-focus *:focus{outline:4px solid #FFB000!important;outline-offset:2px!important;box-shadow:0 0 0 7px #16213A!important}
`;
  const styleEl = document.createElement('style'); styleEl.id = 'a11y-style'; styleEl.textContent = css;
  (document.head || document.documentElement).appendChild(styleEl);

  // ---------- text size (UI only, documents untouched) ----------
  const SKIP = '#pages,#wed,#xed,[contenteditable="true"],[contenteditable=""],canvas,svg,script,style,noscript,.a11y-panel,.a11y-fab,.a11y-skip';
  const base = new WeakMap();
  function baseOf(el) {
    if (base.has(el)) return base.get(el);
    let b = parseFloat(getComputedStyle(el).fontSize) || 16;
    const f = SIZES[st.size] || 1;
    if (f !== 1 && !el.style.fontSize) {
      // added while the page is enlarged: an inherited size is already scaled, an own size is not
      const p = el.parentElement;
      if (p && base.has(p) && Math.abs(parseFloat(getComputedStyle(p).fontSize) - b) < 0.2) b = base.get(p);
    }
    base.set(el, b); return b;
  }
  function scaleTree(root) {
    const f = SIZES[st.size] || 1;
    const els = root.nodeType === 1 ? [root, ...root.querySelectorAll('*')] : [];
    for (const el of els) {
      if (el.closest(SKIP)) continue;
      if (f === 1) { if (el.dataset.a11yFs) { el.style.fontSize = el.dataset.a11yFs === '-' ? '' : el.dataset.a11yFs; delete el.dataset.a11yFs; } continue; }
      const b = baseOf(el);
      if (!el.dataset.a11yFs) el.dataset.a11yFs = el.style.fontSize || '-';
      el.style.fontSize = (b * f).toFixed(2) + 'px';
    }
  }
  let mo = null;
  function applySize() {
    if (!document.body) return;
    if (st.size === 0) { if (mo) { mo.disconnect(); mo = null; } scaleTree(document.body); return; }
    // measure every element at its own size first, then enlarge
    if (mo) mo.disconnect();
    const prev = [...document.body.querySelectorAll('[data-a11y-fs]')];
    prev.forEach(el => { el.style.fontSize = el.dataset.a11yFs === '-' ? '' : el.dataset.a11yFs; delete el.dataset.a11yFs; });
    [document.body, ...document.body.querySelectorAll('*')].forEach(el => { if (!el.closest(SKIP)) { base.delete(el); baseOf(el); } });
    scaleTree(document.body);
    mo = new MutationObserver(list => {
      for (const m of list) for (const n of m.addedNodes) if (n.nodeType === 1 && !n.closest(SKIP)) scaleTree(n);
    });
    mo.observe(document.body, { childList: true, subtree: true });
  }

  // ---------- apply everything ----------
  const MODES = ['contrast', 'gray', 'links', 'font', 'spacing', 'motion', 'cursor', 'focus'];
  function apply() {
    const h = document.documentElement;
    MODES.forEach(m => h.classList.toggle('a11y-' + m, !!st[m]));
    if (st.motion) document.querySelectorAll('video').forEach(v => { try { v.pause(); } catch (_) {} });
    applySize();
    syncPanel();
  }

  // ---------- the panel ----------
  const ic = {
    a11y: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="4.6" r="1.8" fill="currentColor" stroke="none"/><path d="M5 8.5l7 1.5 7-1.5M12 10v4.5M12 14.5l-3 6M12 14.5l3 6"/></svg>',
    contrast: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" aria-hidden="true"><circle cx="12" cy="12" r="8.5"/><path d="M12 3.5a8.5 8.5 0 0 1 0 17z" fill="currentColor"/></svg>',
    gray: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" aria-hidden="true"><path d="M12 3.5c3 4 5.5 6.8 5.5 10a5.5 5.5 0 0 1-11 0c0-3.2 2.5-6 5.5-10z"/><path d="M12 9v10" /></svg>',
    links: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" aria-hidden="true"><path d="M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1"/><path d="M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1"/></svg>',
    font: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" aria-hidden="true"><path d="M4 19L9 5l5 14M6 14h6M15 19h5M17.5 11v8"/></svg>',
    spacing: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" aria-hidden="true"><path d="M9 6h11M9 12h11M9 18h11M4 4v16M2.5 6.5L4 4l1.5 2.5M2.5 17.5L4 20l1.5-2.5"/></svg>',
    motion: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" aria-hidden="true"><circle cx="12" cy="12" r="8.5"/><path d="M10 8.5v7M14 8.5v7" stroke-linecap="round"/></svg>',
    cursor: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round" aria-hidden="true"><path d="M6 3l12 9-5.5 1.2L15.5 20l-2.6 1.1-3-6.8L6 18z"/></svg>',
    focus: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" aria-hidden="true"><rect x="6" y="8" width="12" height="8" rx="2"/><path d="M3 6V3h3M18 3h3v3M21 18v3h-3M6 21H3v-3"/></svg>'
  };
  const LABELS = {
    contrast: 'ניגודיות גבוהה', gray: 'גווני אפור', links: 'הדגשת קישורים', font: 'גופן קריא',
    spacing: 'ריווח טקסט', motion: 'עצירת אנימציות', cursor: 'סמן גדול', focus: 'הדגשת מיקוד'
  };
  let panel = null, opener = null;
  function build() {
    panel = document.createElement('div');
    panel.className = 'a11y-panel'; panel.id = 'a11yPanel'; panel.hidden = true;
    panel.setAttribute('role', 'dialog'); panel.setAttribute('aria-modal', 'false'); panel.setAttribute('aria-labelledby', 'a11yTtl');
    panel.innerHTML = `
      <div class="a11y-head"><h2 id="a11yTtl">תפריט נגישות</h2><button type="button" class="a11y-x" aria-label="סגירת תפריט הנגישות">✕</button></div>
      <div class="a11y-size" role="group" aria-label="גודל טקסט">
        <span>גודל טקסט</span>
        <button type="button" data-sz="-1" aria-label="הקטנת טקסט">−</button>
        <output aria-live="polite">100%</output>
        <button type="button" data-sz="1" aria-label="הגדלת טקסט">+</button>
      </div>
      <div class="a11y-grid">${MODES.map(m => `<button type="button" class="a11y-t" data-m="${m}" aria-pressed="false">${ic[m]}<span>${LABELS[m]}</span></button>`).join('')}</div>
      <div class="a11y-foot"><button type="button" class="a11y-reset">איפוס הגדרות</button><a href="accessibility.html" target="_blank" rel="noopener">הצהרת נגישות</a></div>`;
    document.body.appendChild(panel);
    panel.querySelector('.a11y-x').addEventListener('click', close);
    panel.querySelectorAll('[data-sz]').forEach(b => b.addEventListener('click', () => {
      st.size = Math.max(0, Math.min(SIZES.length - 1, st.size + (+b.dataset.sz))); save(); apply();
    }));
    panel.querySelectorAll('[data-m]').forEach(b => b.addEventListener('click', () => { const m = b.dataset.m; st[m] = !st[m]; save(); apply(); }));
    panel.querySelector('.a11y-reset').addEventListener('click', () => { st = Object.assign({}, DEF); save(); apply(); });
    panel.addEventListener('keydown', e => { if (e.key === 'Escape') { e.stopPropagation(); close(); } }, true);
    document.addEventListener('pointerdown', e => {
      if (!panel.hidden && !panel.contains(e.target) && !(opener && opener.contains(e.target))) close(false);
    }, true);
    window.addEventListener('resize', () => { if (!panel.hidden) place(); });
  }
  function syncPanel() {
    if (!panel) return;
    panel.querySelector('output').textContent = Math.round((SIZES[st.size] || 1) * 100) + '%';
    panel.querySelector('[data-sz="-1"]').disabled = st.size === 0;
    panel.querySelector('[data-sz="1"]').disabled = st.size === SIZES.length - 1;
    panel.querySelectorAll('[data-m]').forEach(b => b.setAttribute('aria-pressed', st[b.dataset.m] ? 'true' : 'false'));
  }
  function place() {
    const r = opener ? opener.getBoundingClientRect() : null;
    const w = panel.offsetWidth, vw = innerWidth, vh = innerHeight;
    if (r && r.bottom < vh / 2) {               // opened from the header: drop down under the button
      panel.style.top = (r.bottom + 8) + 'px'; panel.style.bottom = 'auto';
      panel.style.left = Math.max(10, Math.min(vw - w - 10, r.left + r.width / 2 - w / 2)) + 'px';
    } else {                                     // opened from the corner button: open above it
      panel.style.top = 'auto'; panel.style.bottom = (r ? vh - r.top + 10 : 78) + 'px';
      panel.style.left = Math.max(10, r ? r.left : 16) + 'px';
    }
  }
  function open(from) {
    if (!panel) build();
    opener = from || opener; syncPanel(); panel.hidden = false; place();
    if (opener) opener.setAttribute('aria-expanded', 'true');
    panel.querySelector('[data-sz="1"]').focus();
  }
  function close(refocus = true) {
    if (!panel || panel.hidden) return;
    panel.hidden = true;
    if (opener) { opener.setAttribute('aria-expanded', 'false'); if (refocus) opener.focus(); }
  }
  const toggle = from => (panel && !panel.hidden ? close() : open(from));
  window.simanA11y = { open, close, toggle };

  function wire() {
    let btn = document.getElementById('a11yBtn');
    if (!btn) { // pages without a header button get a round corner button
      btn = document.createElement('button');
      btn.type = 'button'; btn.className = 'a11y-fab'; btn.id = 'a11yBtn';
      btn.setAttribute('aria-label', 'תפריט נגישות'); btn.title = 'נגישות';
      btn.innerHTML = ic.a11y; document.body.appendChild(btn);
    } else if (!btn.innerHTML.trim()) btn.innerHTML = ic.a11y;
    btn.setAttribute('aria-haspopup', 'dialog'); btn.setAttribute('aria-controls', 'a11yPanel'); btn.setAttribute('aria-expanded', 'false');
    btn.addEventListener('click', e => { e.preventDefault(); e.stopPropagation(); toggle(btn); });
    apply();
  }
  // apply the visual modes immediately (no flash), the rest once the page is ready
  MODES.forEach(m => document.documentElement.classList.toggle('a11y-' + m, !!st[m]));
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', wire); else wire();
  // Word/Excel editors and dialogs that are built later get the enlarged size automatically (MutationObserver above)
})();
