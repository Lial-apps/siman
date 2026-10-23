/* סימן — "לפי משימה" shell (6.0)
   Rearranges the existing interface without replacing any feature:
   buttons are MOVED (not copied), so every existing handler keeps working.
   - top bar: name + file, "ask Siman" bar, save & share
   - bottom dock: five work modes (mark, sign, form, translate, pages);
     above it only the tools of the current mode
   - opening screen: one question, a big drop area and task tiles */
(function () {
  'use strict';
  const $ = id => document.getElementById(id);
  const bar = document.querySelector('header.bar');
  if (!bar || !$('tools')) return;
  document.body.classList.add('s6');

  const svg = d => `<svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${d}</svg>`;
  const I = {
    hl: '<path d="M9 15l-4 4h6l1-1"/><path d="M9 15l7-10 4 3-7 10z"/><path d="M8 13l4 3"/>',
    sign: '<path d="M3 17c3-1 4-9 6-9s0 9 3 9 3-5 5-5 1 4 4 4"/><path d="M3 21h18"/>',
    form: '<rect x="4" y="3" width="16" height="18" rx="2"/><path d="M8 8h8M8 12h8M8 16h4"/>',
    tr: '<path d="M4 5h9M8.5 3v2c0 4-2 7-5 9"/><path d="M6 9c1 2 3 4 6 5"/><path d="M13 21l4-9 4 9M14.5 18h5"/>',
    pages: '<rect x="7" y="3" width="12" height="15" rx="2"/><path d="M5 7v12a2 2 0 002 2h9"/>',
    up: '<path d="M12 16V4M7 9l5-5 5 5"/><path d="M4 16v3a2 2 0 002 2h12a2 2 0 002-2v-3"/>',
    dots: '<circle cx="5" cy="12" r="1.4" fill="currentColor"/><circle cx="12" cy="12" r="1.4" fill="currentColor"/><circle cx="19" cy="12" r="1.4" fill="currentColor"/>',
    link: '<path d="M10 14a4 4 0 006 0l3-3a4 4 0 00-6-6l-1 1"/><path d="M14 10a4 4 0 00-6 0l-3 3a4 4 0 006 6l1-1"/>',
    spark: '<path d="M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8z"/>',
    read: '<path d="M4 9v6h4l5 4V5L8 9z"/><path d="M16 9a4 4 0 010 6M18.5 6.5a8 8 0 010 11"/>',
    panel: '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M15 4v16"/>',
    merge: '<path d="M7 3h7l4 4v14H7z"/><path d="M14 3v4h4M12.5 11v6M9.5 14h6"/>',
    split: '<path d="M4 4h7v9H4zM13 11h7v9h-7z"/>',
    conv: '<path d="M7 3h7l4 4v14H7z"/><path d="M14 3v4h4M10 13h5M10 17h5"/>',
    print: '<path d="M7 9V3h10v6"/><rect x="3" y="9" width="18" height="8" rx="2"/><path d="M7 14h10v7H7z"/>',
    open: '<path d="M3 7a2 2 0 012-2h4l2 2h8a2 2 0 012 2v8a2 2 0 01-2 2H5a2 2 0 01-2-2z"/>',
    ocr: '<path d="M4 8V5a1 1 0 011-1h3M16 4h3a1 1 0 011 1v3M20 16v3a1 1 0 01-1 1h-3M8 20H5a1 1 0 01-1-1v-3M8 10h8M8 14h5"/>',
    plus: '<path d="M12 5v14M5 12h14"/>'
  };
  const el = (html) => { const t = document.createElement('template'); t.innerHTML = html.trim(); return t.content.firstElementChild; };
  const click = id => { const b = $(id); if (b && !b.disabled) b.click(); };
  const hidden = n => !n || n.classList.contains('hidden');

  /* ================================================================ top bar */
  const brand = bar.querySelector('.brand');
  const groups = [...bar.querySelectorAll(':scope > .group')];
  const start = el('<div class="sh-start"></div>');
  const mid = el('<div class="sh-mid"></div>');
  const end = el('<div class="sh-end"></div>');
  const stash = el('<div class="sh-stash" hidden></div>'); // still in the page, still clickable
  bar.prepend(start); start.after(mid); mid.after(end); end.after(stash);
  start.append(brand);

  const cmd = $('cmdBtn');
  if (cmd) {
    mid.append(cmd);
    const lbl = cmd.querySelector('.lbl');
    if (lbl) lbl.textContent = 'בקשו מסימן… למשל: תרגום לאנגלית, פיצול עמודים';
  }

  // the "more" menu: things used once in a while
  const moreBtn = el(`<button class="icon-btn sh-more" id="shMoreBtn" type="button" aria-haspopup="menu" aria-expanded="false" title="עוד פעולות" aria-label="עוד פעולות">${svg(I.dots)}</button>`);
  const moreMenu = el('<div class="sh-menu hidden" id="shMoreMenu" role="menu"></div>');
  const moreItems = [['openBtn', I.open, 'פתיחת קובץ אחר'], ['printBtn', I.print, 'הדפסה'], ['convBtn', I.conv, 'המרה ל-Word, Excel או PowerPoint']];
  for (const [id, ic, text] of moreItems) {
    const b = $(id); if (!b) continue;
    b.classList.add('sh-mi'); b.setAttribute('role', 'menuitem');
    b.innerHTML = svg(ic) + '<span>' + text + '</span>';
    moreMenu.append(b);
  }
  const moreWrap = el('<div class="sh-more-wrap"></div>');
  moreWrap.append(moreBtn, moreMenu);
  const toggleMore = open => { moreMenu.classList.toggle('hidden', !open); moreBtn.setAttribute('aria-expanded', open); };
  moreBtn.addEventListener('click', e => { e.stopPropagation(); toggleMore(moreMenu.classList.contains('hidden')); });
  moreMenu.addEventListener('click', () => toggleMore(false));
  document.addEventListener('click', e => { if (!moreWrap.contains(e.target)) toggleMore(false); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') toggleMore(false); });

  const undo = $('undo');
  if (undo) { undo.innerHTML = svg('<path d="M9 14L4 9l5-5"/><path d="M4 9h11a5 5 0 010 10h-3"/>'); end.append(undo); }
  for (const id of ['aiBtn', 'admBtn', 'setBtn', 'installBtn']) { const b = $(id); if (b) end.append(b); }
  end.append(moreWrap);
  const wordEdit = $('wordEditBtn'); if (wordEdit) end.append(wordEdit);
  const split = bar.querySelector('.split-btn'); if (split) end.append(split);
  const share = $('shareBtn');
  if (share) {
    end.append(share);
    const t = [...share.childNodes].find(n => n.nodeType === 3 && n.textContent.trim());
    if (t) t.textContent = 'שיתוף'; else if (!share.querySelector('span')) share.append('שיתוף');
  }
  for (const id of ['pgBtn', 'trBtn']) { const b = $(id); if (b) stash.append(b); }

  /* ================================================================ bottom dock */
  const MODES = [
    { k: 'mark', ic: I.hl, l: 'סימון וכתיבה', tools: ['select', 'highlight', 'pen', 'text', 'edit', 'eraser', 'image', 'snip', 'redact'] },
    { k: 'sign', ic: I.sign, l: 'חתימה', tools: ['sign', 'text', 'select'], px: [['sendSign', I.link, 'שליחה לחתימה']] },
    { k: 'form', ic: I.form, l: 'טופס', tools: ['select', 'text', 'sign'], px: [['ocrPx', I.ocr, 'זיהוי טקסט']], hint: 'לוחצים על שדה בטופס וכותבים. אין שדות? בוחרים "טקסט" וכותבים על הקו.' },
    { k: 'tr', ic: I.tr, l: 'תרגום', tools: [], px: [['trPx', I.panel, 'חלונית התרגום'], ['aiPx', I.spark, 'שאלות וסיכום'], ['readPx', I.read, 'הקראה']] },
    { k: 'pages', ic: I.pages, l: 'עמודים', tools: [], px: [['pgPx', I.panel, 'חלונית העמודים'], ['mergePx', I.merge, 'הוספת PDF'], ['splitPx', I.split, 'פיצול לקבצים'], ['convPx', I.conv, 'המרה'], ['printPx', I.print, 'הדפסה']] }
  ];
  const PX = {
    sendSign: () => { if (typeof sendLink === 'function') sendLink('sign'); },
    ocrPx: () => { const b = $('ocrBtn'); if (b) { b.classList.contains('hidden') ? click('ocrAllBtn') : b.click(); } },
    trPx: () => { if (hidden($('tpanel'))) click('trBtn'); },
    aiPx: () => click('aiBtn'),
    readPx: () => click('aiRead'),
    pgPx: () => { if (hidden($('pgpanel'))) click('pgBtn'); },
    mergePx: () => click('mergeBtn'),
    splitPx: () => click('splitAllBtn'),
    convPx: () => click('convBtn'),
    printPx: () => click('printBtn')
  };

  const wrap = el('<div class="sh-dockwrap" id="shDock"></div>');
  const opts = $('opts');
  const tray = el('<div class="sh-tray" role="toolbar" aria-label="הכלים של המצב"></div>');
  const hint = el('<p class="sh-hint hidden"></p>');
  const pxBox = el('<div class="sh-px"></div>');
  const tools = $('tools');
  tray.append(tools, pxBox);
  const dock = el('<nav class="sh-dock" role="tablist" aria-label="מצבי עבודה"></nav>');
  for (const m of MODES) {
    dock.append(el(`<button type="button" class="sh-mode" role="tab" data-mode="${m.k}" aria-selected="false">${svg(m.ic)}<span>${m.l}</span><i aria-hidden="true"></i></button>`));
  }
  if (opts) wrap.append(opts);
  wrap.append(hint, tray, dock);
  document.body.append(wrap);

  // zoom + page number: a small floating corner
  const zoom = bar.querySelector('.group.zoom');
  const corner = el('<div class="sh-corner"></div>');
  if (zoom) corner.append(zoom);
  const pageInfo = $('pageInfo'); if (pageInfo) corner.prepend(pageInfo);
  document.body.append(corner);
  groups.forEach(g => { if (!g.children.length) g.remove(); });

  let mode = 'mark';
  function toolBtn(k) { return tools.querySelector(k === 'sign' || k === 'image' ? `.tool[data-action="${k}"]` : `.tool[data-tool="${k}"]`); }
  function setMode(k, opts2 = {}) {
    const m = MODES.find(x => x.k === k) || MODES[0];
    const prev = mode; mode = m.k;
    document.body.dataset.shmode = m.k;
    dock.querySelectorAll('.sh-mode').forEach(b => b.setAttribute('aria-selected', b.dataset.mode === m.k));
    tools.querySelectorAll('.tool').forEach(b => {
      const key = b.dataset.tool || b.dataset.action;
      b.hidden = !m.tools.includes(key);
      b.style.order = m.tools.indexOf(key);
    });
    tools.hidden = !m.tools.length;
    pxBox.innerHTML = '';
    for (const [id, ic, l] of (m.px || [])) {
      const b = el(`<button type="button" class="sh-tool" data-px="${id}">${svg(ic)}<span>${l}</span></button>`);
      b.onclick = PX[id];
      pxBox.append(b);
    }
    pxBox.hidden = !(m.px && m.px.length);
    hint.textContent = m.hint || ''; hint.classList.toggle('hidden', !m.hint);
    if (opts2.quiet) return;
    // the panels that belong to a mode open and close with it
    if (prev === 'tr' && m.k !== 'tr' && !hidden($('tpanel'))) click('tpClose');
    if (prev === 'pages' && m.k !== 'pages' && !hidden($('pgpanel'))) click('pgClose');
    if (m.k === 'tr') PX.trPx();
    if (m.k === 'pages') PX.pgPx();
    if (typeof setTool === 'function' && prev !== m.k) {
      if (m.k === 'mark') { const h = toolBtn('highlight'); if (h) h.click(); }
      else if (m.k === 'form' || m.k === 'sign') { const s = toolBtn('select'); if (s) s.click(); }
    }
    place();
  }
  dock.addEventListener('click', e => { const b = e.target.closest('.sh-mode'); if (b) setMode(b.dataset.mode); });
  dock.addEventListener('keydown', e => {
    if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return;
    const bs = [...dock.querySelectorAll('.sh-mode')]; const i = bs.indexOf(document.activeElement); if (i < 0) return;
    const n = bs[(i + (e.key === 'ArrowLeft' ? 1 : -1) + bs.length) % bs.length]; n.focus(); n.click(); e.preventDefault();
  });
  window.simanMode = setMode;
  setMode('mark', { quiet: true });

  // the dock sits in the middle of the document area, also when side panels are open
  const desk = $('desk');
  function place() {
    if (!desk) return;
    const r = desk.getBoundingClientRect();
    wrap.style.left = r.left + 'px'; wrap.style.width = r.width + 'px';
    corner.style.setProperty('--sh-dl', r.left + 'px');
    corner.style.setProperty('--sh-dr', (innerWidth - r.right) + 'px');
    corner.style.setProperty('--sh-dt', r.top + 'px');
  }
  addEventListener('resize', place);
  if (desk && 'ResizeObserver' in window) new ResizeObserver(place).observe(desk);

  // is a document open on the desk? (not the opening screen, not the full Word/Excel editors)
  function sync() {
    const docOpen = hidden($('empty')) && hidden($('officeView')) && !document.body.classList.contains('ed-live') && !!document.querySelector('#pages .page, #pages canvas');
    document.body.classList.toggle('sh-doc', docOpen);
    // a panel closed by its own × brings the dock back to marking
    if (mode === 'tr' && hidden($('tpanel'))) setMode('mark', { quiet: true });
    if (mode === 'pages' && hidden($('pgpanel'))) setMode('mark', { quiet: true });
    if (!hidden($('tpanel')) && mode !== 'tr' && mode !== 'pages') setMode('tr', { quiet: true });
    if (!hidden($('pgpanel')) && mode !== 'pages' && mode !== 'tr') setMode('pages', { quiet: true });
    // a tool chosen from the keyboard or the command bar moves the dock to its mode
    const t = document.body.dataset.tool;
    if (docOpen && t && t !== 'select' && t !== 'place') {
      const cur = MODES.find(x => x.k === mode);
      if (!cur.tools.includes(t)) setMode('mark', { quiet: true });
    }
    place();
  }
  // the options pill shows only when it has something in it
  function optsCheck() {
    if (!opts) return;
    const any = [...opts.children].some(c => !c.classList.contains('hidden') && !c.hidden && getComputedStyle(c).display !== 'none' && (c.textContent.trim() || c.querySelector('input,select,button,span')));
    opts.classList.toggle('sh-empty', !any);
  }
  if (opts) { new MutationObserver(optsCheck).observe(opts, { subtree: true, childList: true, attributes: true, attributeFilter: ['class', 'hidden'], characterData: true }); optsCheck(); }
  const mo = new MutationObserver(sync);
  for (const n of [$('empty'), $('officeView'), $('tpanel'), $('pgpanel'), $('pages')].filter(Boolean)) mo.observe(n, { attributes: true, attributeFilter: ['class'], childList: n.id === 'pages' });
  mo.observe(document.body, { attributes: true, attributeFilter: ['class', 'data-tool'] });
  sync();

  /* ================================================================ opening screen */
  const welcome = document.querySelector('.welcome');
  if (welcome) {
    welcome.classList.add('sh-home');
    const h1 = welcome.querySelector('h1');
    if (h1) h1.innerHTML = 'מה עושים היום <mark>עם הקובץ?</mark>';
    const sub = welcome.querySelector('.wl-sub');
    const cta = welcome.querySelector('.wl-cta');
    const open2 = $('openBtn2');
    const drop = el(`<div class="sh-drop" role="group" aria-label="פתיחת קובץ">
      <span class="sh-drop-ic">${svg(I.up)}</span>
      <div class="sh-drop-t"><b>גוררים לכאן קובץ, או בוחרים מהמכשיר</b>
      <span>הקובץ נשאר אצלכם במכשיר. אפשר גם לצלם דף עם הטלפון.</span>
      <span class="sh-fmts" dir="ltr"><i>PDF</i><i>Word</i><i>Excel</i><i>PowerPoint</i></span></div></div>`);
    if (cta) cta.replaceWith(drop); else if (h1) h1.after(drop);
    if (open2) { open2.textContent = 'בחירת קובץ'; drop.append(open2); }
    drop.addEventListener('click', e => { if (e.target === open2) return; if (open2) open2.click(); });
    if (sub) sub.remove();

    const quick = welcome.querySelector('.wl-quick');
    if (quick) {
      const tasks = el('<div class="sh-tasks" role="group" aria-label="מה עושים"></div>');
      const more = el('<div class="sh-small" role="group" aria-label="עוד"></div>');
      const DESC = {
        sign: ['לחתום', 'חתימה, ראשי תיבות ושם בכמה לחיצות'],
        mark: ['לסמן ולכתוב', 'מרקר, עט, הערות ותיקון הטקסט'],
        tr: ['לתרגם', 'לעברית, אנגלית, רוסית, ערבית ועוד'],
        pages: ['לסדר עמודים', 'לאחד, לפצל, לסובב ולמחוק'],
        send: ['לשלוח לחתימה', 'קישור לכל חותם, או לקבוצה שלמה']
      };
      const markQ = el(`<button class="qa" type="button" data-qa="mark">${svg(I.hl)}לסמן ולכתוב</button>`);
      markQ.onclick = () => {
        if (typeof needReg === 'function' && needReg()) { if (typeof showReg === 'function') showReg(); return; }
        try { afterOpen = () => setMode('mark'); } catch (_) {}
        if (open2) open2.click();
      };
      const byQa = k => k === 'mark' ? markQ : quick.querySelector(`.qa[data-qa="${k}"]`);
      for (const k of ['sign', 'mark', 'tr', 'pages', 'send']) {
        const b = byQa(k); if (!b) continue;
        const ic = b.querySelector('svg');
        b.textContent = '';
        b.classList.add('sh-task', 'sh-t-' + k);
        b.append(el('<span class="sh-ti"></span>'));
        if (ic) b.firstChild.append(ic);
        b.append(el(`<b>${DESC[k][0]}</b>`), el(`<small>${DESC[k][1]}</small>`));
        tasks.append(b);
      }
      [...quick.querySelectorAll('.qa')].forEach(b => { if (!b.classList.contains('sh-task')) more.append(b); });
      quick.replaceWith(tasks); tasks.after(more);
      // after a file opens from a task tile, the dock starts in the matching mode
      const startMode = { sign: 'sign', tr: 'tr', pages: 'pages' };
      tasks.addEventListener('click', e => { const b = e.target.closest('.qa'); if (b && startMode[b.dataset.qa]) pendingMode = startMode[b.dataset.qa]; }, true);
    }
    const list = welcome.querySelector('.wl-list'); if (list) list.classList.add('hidden');
  }
  let pendingMode = null;
  if (typeof openFile === 'function') {
    const _o = openFile;
    openFile = function () {
      const r = _o.apply(this, arguments);
      const go = () => { if (pendingMode) { const k = pendingMode; pendingMode = null; setMode(k, { quiet: k === 'sign' }); } };
      if (r && typeof r.then === 'function') r.then(go, () => { pendingMode = null; }); else go();
      return r;
    };
  }
  // any open from the top bar starts in marking
  const ob = $('openBtn'); if (ob) ob.addEventListener('click', () => { pendingMode = 'mark'; });
})();
