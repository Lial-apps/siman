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

  /* ================================================================ opening screen (6.1: "night desk") */
  let pendingMode = null;
  const welcome = document.querySelector('.welcome');
  const open2 = $('openBtn2');
  const homeOn = () => !hidden($('empty'));
  const guard = () => { if (typeof needReg === 'function' && needReg()) { if (typeof showReg === 'function') showReg(); return false; } return true; };
  if (welcome) {
    welcome.classList.add('sh-home');
    const txt = welcome.querySelector('.wl-text') || welcome;
    const hi = $('wlHi'), trial = $('wlTrial'), ios = $('iosHint');
    const h1 = welcome.querySelector('h1');
    if (h1) h1.innerHTML = 'מה נעשה <mark>עם הקובץ?</mark>';
    const quick = welcome.querySelector('.wl-quick');
    const findQa = re => quick && [...quick.querySelectorAll('.qa')].find(b => re.test(b.textContent));
    const scanQ = findQa(/כתב יד/);

    const hero = el('<section class="sh-hero"></section>');
    const heroTxt = el('<div class="sh-hero-txt"></div>');
    const lead = el('<p class="sh-lead">לסמן, לחתום, לתרגם ולעצב. PDF, Word, Excel ו-PowerPoint, במחשב ובטלפון. הקובץ נשאר אצלכם.</p>');
    const acts = el('<div class="sh-acts"></div>');
    if (open2) { open2.className = 'sh-open'; open2.innerHTML = svg(I.up) + '<span>פתיחת קובץ</span>'; acts.append(open2); }
    const scanB = el(`<button type="button" class="sh-round" aria-label="סריקת דף">${svg('<path d="M4 8V5a1 1 0 011-1h3M16 4h3a1 1 0 011 1v3M20 16v3a1 1 0 01-1 1h-3M8 20H5a1 1 0 01-1-1v-3"/><circle cx="12" cy="12" r="3.2"/>')}<small>סריקה</small></button>`);
    scanB.onclick = () => { if (scanQ) scanQ.click(); else if (open2) open2.click(); };
    const askB = el(`<button type="button" class="sh-round sh-orb" aria-label="לבקש מסימן">${svg(I.spark)}<small>לבקש</small></button>`);
    askB.onclick = () => click('cmdBtn');
    acts.append(scanB, askB);
    const drop = el('<p class="sh-droptxt">או גוררים קובץ לכל מקום במסך</p>');
    heroTxt.append(...[hi, h1, lead, acts, drop, trial].filter(Boolean));
    const stage = el('<div class="sh-stage" aria-live="polite"></div>');
    hero.append(heroTxt, stage);

    // the five tasks
    const tasks = el('<div class="sh-tasks" role="group" aria-label="מה עושים"></div>');
    const more = el('<div class="sh-small" role="group" aria-label="עוד"></div>');
    const DESC = {
      sign: ['לחתום', 'חתימה ושם בכמה לחיצות'],
      mark: ['לסמן ולכתוב', 'מרקר, עט, הערות ותיקונים'],
      tr: ['לתרגם', 'לעברית, אנגלית, רוסית, ערבית'],
      pages: ['לסדר עמודים', 'לאחד, לפצל, לסובב'],
      send: ['לשלוח לחתימה', 'קישור לחותם או לקבוצה']
    };
    if (quick) {
      const markQ = el(`<button class="qa" type="button" data-qa="mark">${svg(I.hl)}לסמן ולכתוב</button>`);
      markQ.onclick = () => { if (!guard()) return; try { afterOpen = () => setMode('mark'); } catch (_) {} if (open2) open2.click(); };
      for (const k of ['sign', 'mark', 'tr', 'pages', 'send']) {
        const b = k === 'mark' ? markQ : quick.querySelector(`.qa[data-qa="${k}"]`); if (!b) continue;
        const ic = b.querySelector('svg');
        b.textContent = ''; b.classList.add('sh-task', 'sh-t-' + k);
        b.append(el('<span class="sh-ti"></span>')); if (ic) b.firstChild.append(ic);
        b.append(el(`<b>${DESC[k][0]}</b>`), el(`<small>${DESC[k][1]}</small>`));
        tasks.append(b);
      }
      [...quick.querySelectorAll('.qa')].forEach(b => { if (!b.classList.contains('sh-task')) more.append(b); });
      const startMode = { sign: 'sign', tr: 'tr', pages: 'pages' };
      tasks.addEventListener('click', e => { const b = e.target.closest('.qa'); if (b && startMode[b.dataset.qa]) pendingMode = startMode[b.dataset.qa]; }, true);
      quick.remove();
    }
    // build the new screen and drop the old pieces
    const cta = welcome.querySelector('.wl-cta'); if (cta) cta.remove();
    const sub = welcome.querySelector('.wl-sub'); if (sub) sub.remove();
    const art = welcome.querySelector('.wl-art'); if (art) art.remove();
    const swipe = welcome.querySelector('.swipe'); if (swipe) swipe.remove();
    const list = welcome.querySelector('.wl-list'); if (list) list.classList.add('hidden');
    welcome.prepend(hero); hero.after(tasks); tasks.after(more);
    if (ios) more.after(ios);
    if (txt !== welcome && !txt.children.length) txt.remove(); else if (txt !== welcome) txt.classList.add('hidden');

    /* ---------- recent files: kept only on this device (IndexedDB), never uploaded ---------- */
    const RK = 'siman-recent', MAX = 6, MAXB = 30 * 1024 * 1024;
    const idb = () => new Promise(res => { try { const r = indexedDB.open(RK, 1); r.onupgradeneeded = () => r.result.createObjectStore('f', { keyPath: 'id' }); r.onsuccess = () => res(r.result); r.onerror = () => res(null); } catch (_) { res(null); } });
    const tx = async (mode, fn) => { const db = await idb(); if (!db) return null; return new Promise(res => { try { const t = db.transaction('f', mode); const st = t.objectStore('f'); const out = fn(st); t.oncomplete = () => { res(out && out.result !== undefined ? out.result : out); db.close(); }; t.onerror = () => { res(null); db.close(); }; } catch (_) { res(null); } }); };
    const allRecent = async () => ((await tx('readonly', st => st.getAll())) || []).sort((a, b) => b.t - a.t);
    const kindOf = n => { const e = (String(n).split('.').pop() || '').toLowerCase(); return /^docx?$/.test(e) ? 'word' : /^xlsx?$/.test(e) ? 'excel' : /^pptx?$/.test(e) ? 'ppt' : e === 'pdf' ? 'pdf' : 'file'; };
    async function pdfThumb(file) {
      try {
        if (typeof pdfjsLib === 'undefined') return '';
        const doc = await pdfjsLib.getDocument({ data: new Uint8Array(await file.arrayBuffer()) }).promise;
        const pg = await doc.getPage(1); const v0 = pg.getViewport({ scale: 1 }); const vp = pg.getViewport({ scale: 260 / v0.width });
        const c = document.createElement('canvas'); c.width = vp.width; c.height = vp.height;
        await pg.render({ canvasContext: c.getContext('2d'), viewport: vp }).promise;
        const url = c.toDataURL('image/jpeg', .72); doc.destroy(); return url;
      } catch (_) { return ''; }
    }
    async function remember(file) {
      if (!file || !file.name || !file.size || file.size > MAXB) return;
      const id = file.name + '|' + file.size;
      const old = ((await allRecent()) || []).find(r => r.id === id);
      const rec = { id, name: file.name, type: file.type || '', size: file.size, t: Date.now(), blob: file, thumb: old ? old.thumb : '' };
      await tx('readwrite', st => st.put(rec));
      const all = await allRecent(); for (const r of all.slice(MAX)) await tx('readwrite', st => st.delete(r.id));
      if (!rec.thumb && kindOf(file.name) === 'pdf') { setTimeout(async () => { rec.thumb = await pdfThumb(file); if (rec.thumb) await tx('readwrite', st => st.put(rec)); }, 1800); }
    }
    const ago = t => { const m = Math.round((Date.now() - t) / 60000); if (m < 1) return 'עכשיו'; if (m < 60) return `לפני ${m} דק׳`; const h = Math.round(m / 60); if (h < 24) return h === 1 ? 'לפני שעה' : `לפני ${h} שעות`; const d = Math.round(h / 24); return d === 1 ? 'אתמול' : `לפני ${d} ימים`; };
    const LBL = { word: 'Word', excel: 'Excel', ppt: 'PowerPoint', pdf: 'PDF', file: 'קובץ' };
    const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
    const prettyName = n => n.replace(/\.[^.]+$/, '');
    function miniPaper(k) {
      const lines = k === 'excel' ? '<i class="g"></i><i class="g"></i><i class="g"></i><i class="g"></i>' : k === 'ppt' ? '<i class="t"></i><i></i><i class="s"></i>' : '<i class="t"></i><i></i><i></i><i class="s"></i><i></i>';
      return `<span class="sh-mp sh-mp-${k}">${lines}</span>`;
    }
    async function renderStage() {
      const recs = (await allRecent()) || [];
      stage.classList.toggle('has-recent', !!recs.length);
      if (!recs.length) {
        stage.innerHTML = `<div class="sh-art" aria-hidden="true"><div class="sh-paper p2"></div><div class="sh-paper p1">
          <i class="t"></i><i></i><i class="m"></i><i></i><i></i><i class="s"></i>
          <svg class="sig" viewBox="0 0 200 70"><path d="M10 50c14-4 18-38 28-38s-4 40 8 40 10-26 20-26-2 22 8 22 12-30 22-30 0 30 12 30c10 0 14-14 24-14s10 10 20 10 14-6 28-8"/></svg>
          <span class="mk"></span></div></div>`;
        return;
      }
      const [top, ...rest] = recs;
      const k = kindOf(top.name);
      stage.innerHTML = `<p class="sh-st-h">ממשיכים מאיפה שעצרת</p>
        <div class="sh-rc-main" role="group" aria-label="${esc(top.name)}">
          <div class="sh-rc-pv">${top.thumb ? `<img src="${top.thumb}" alt="">` : miniPaper(k)}</div>
          <div class="sh-rc-t"><span class="sh-kind k-${k}">${LBL[k]}</span><b title="${esc(top.name)}">${esc(prettyName(top.name))}</b><small>${ago(top.t)}</small>
          <button type="button" class="sh-rc-open" data-id="${esc(top.id)}">לפתוח שוב</button></div>
          <button type="button" class="sh-rc-x" data-del="${esc(top.id)}" aria-label="הסרה מהרשימה" title="הסרה מהרשימה">×</button>
        </div>
        ${rest.length ? `<div class="sh-rc-list">${rest.slice(0, 4).map(r => { const kk = kindOf(r.name); return `<button type="button" class="sh-rc" data-id="${esc(r.id)}" title="${esc(r.name)}"><span class="sh-kind k-${kk}">${LBL[kk]}</span><b>${esc(prettyName(r.name))}</b><small>${ago(r.t)}</small></button>`; }).join('')}</div>` : ''}
        <button type="button" class="sh-rc-clear">ניקוי הרשימה</button>
        <p class="sh-rc-note">הרשימה נשמרת רק במכשיר הזה</p>`;
    }
    stage.addEventListener('click', async e => {
      const del = e.target.closest('[data-del]');
      if (del) { await tx('readwrite', st => st.delete(del.dataset.del)); renderStage(); return; }
      if (e.target.closest('.sh-rc-clear')) { await tx('readwrite', st => st.clear()); renderStage(); return; }
      const b = e.target.closest('[data-id]'); if (!b) return;
      if (!guard()) return;
      const r = ((await allRecent()) || []).find(x => x.id === b.dataset.id); if (!r) { renderStage(); return; }
      const f = new File([r.blob], r.name, { type: r.type, lastModified: r.t });
      if (typeof openFile === 'function') openFile(f);
    });
    window.simanRecent = { remember, renderStage };
    renderStage();
    // back on the opening screen: refresh the list
    new MutationObserver(() => { if (homeOn()) renderStage(); }).observe($('empty'), { attributes: true, attributeFilter: ['class'] });
  }

  function homeSync() { document.body.classList.toggle('sh-homeon', homeOn() && !document.body.classList.contains('ed-live')); }
  new MutationObserver(homeSync).observe($('empty'), { attributes: true, attributeFilter: ['class'] });
  new MutationObserver(homeSync).observe(document.body, { attributes: true, attributeFilter: ['class'] });
  homeSync();

  if (typeof openFile === 'function') {
    const _o = openFile;
    openFile = function (file) {
      try { if (file && window.simanRecent) window.simanRecent.remember(file); } catch (_) {}
      const r = _o.apply(this, arguments);
      const go = () => { if (pendingMode) { const k = pendingMode; pendingMode = null; setMode(k, { quiet: k === 'sign' }); } };
      if (r && typeof r.then === 'function') r.then(go, () => { pendingMode = null; }); else go();
      return r;
    };
  }
  const ob = $('openBtn'); if (ob) ob.addEventListener('click', () => { pendingMode = 'mark'; });

  /* ================================================================ selection bubble */
  const bub = el(`<div class="sh-bub" role="toolbar" aria-label="פעולות על הטקסט המסומן" hidden>
    <button type="button" data-b="hl">${svg(I.hl)}<span>מרקר</span></button>
    <button type="button" data-b="tr">${svg(I.tr)}<span>תרגום</span></button>
    <button type="button" data-b="ex">${svg(I.spark)}<span>להסביר</span></button>
    <button type="button" data-b="rd">${svg(I.read)}<span>הקראה</span></button>
    <button type="button" data-b="cp">${svg('<rect x="8" y="8" width="12" height="12" rx="2"/><path d="M16 8V5a1 1 0 00-1-1H5a1 1 0 00-1 1v10a1 1 0 001 1h3"/>')}<span>העתקה</span></button>
  </div>`);
  document.body.append(bub);
  let selRange = null, selText = '';
  const pagesEl = $('pages');
  function hideBub() { bub.hidden = true; selRange = null; }
  function placeBub() {
    const s = getSelection();
    if (!s || s.isCollapsed || !s.rangeCount || !pagesEl) return hideBub();
    const r = s.getRangeAt(0);
    const node = r.commonAncestorContainer.nodeType === 1 ? r.commonAncestorContainer : r.commonAncestorContainer.parentElement;
    if (!node || !pagesEl.contains(node)) return hideBub();
    const t = s.toString().replace(/\s+/g, ' ').trim(); if (t.length < 2) return hideBub();
    selRange = r.cloneRange(); selText = t;
    const rects = [...r.getClientRects()].filter(x => x.width > 1 && x.height > 1); if (!rects.length) return hideBub();
    const top = Math.min(...rects.map(x => x.top)), bottom = Math.max(...rects.map(x => x.bottom));
    const left = Math.min(...rects.map(x => x.left)), right = Math.max(...rects.map(x => x.right));
    bub.hidden = false;
    const bw = bub.offsetWidth, bh = bub.offsetHeight;
    let y = top - bh - 12; if (y < 70) y = bottom + 12;
    const x = Math.max(8, Math.min(innerWidth - bw - 8, (left + right) / 2 - bw / 2));
    bub.style.left = x + 'px'; bub.style.top = y + 'px';
  }
  let selT = 0;
  document.addEventListener('selectionchange', () => { clearTimeout(selT); selT = setTimeout(placeBub, 220); });
  if (desk) desk.addEventListener('scroll', () => { if (!bub.hidden) placeBub(); }, { passive: true });
  bub.addEventListener('pointerdown', e => e.preventDefault()); // keep the selection
  function markSelection() {
    if (!selRange || typeof S === 'undefined' || typeof toPdf !== 'function' || typeof drawAnns !== 'function') return false;
    const rects = [...selRange.getClientRects()].filter(x => x.width > 1 && x.height > 1);
    const adds = [];
    for (const P of S.pages || []) {
      if (!P || !P.ann) continue;
      const pr = P.ann.getBoundingClientRect();
      for (const q of rects) {
        const cx = (q.left + q.right) / 2, cy = (q.top + q.bottom) / 2;
        if (cx < pr.left || cx > pr.right || cy < pr.top || cy > pr.bottom) continue;
        const [x1, y1] = toPdf(P, q.left - pr.left, q.top - pr.top - 1), [x2, y2] = toPdf(P, q.right - pr.left, q.bottom - pr.top + 1);
        adds.push({ P, a: { type: 'hl', page: P.i - 1, x1, y1, x2, y2, color: S.color && S.color.highlight } });
      }
    }
    if (!adds.length) return false;
    if (typeof snapshot === 'function') snapshot();
    for (const { a } of adds) S.anns.push(Object.assign({ id: S.nextId++ }, a));
    new Set(adds.map(x => x.P)).forEach(P => drawAnns(P));
    return true;
  }
  function askAI(q) {
    const panel = $('aiPanel'), inp = $('aiQ'), send = $('aiSend');
    if (!inp || !send) { if (typeof toast === 'function') toast('העוזר לא זמין כרגע'); return; }
    if (panel && panel.classList.contains('hidden')) click('aiBtn');
    inp.value = q; inp.dispatchEvent(new Event('input', { bubbles: true }));
    setTimeout(() => send.click(), 80);
  }
  bub.addEventListener('click', async e => {
    const b = e.target.closest('[data-b]'); if (!b || !selText) return;
    const t = selText.slice(0, 2500), k = b.dataset.b;
    if (k === 'hl') { if (!markSelection() && typeof toast === 'function') toast('אי אפשר לסמן כאן. נסו את כלי המרקר'); }
    else if (k === 'tr') askAI(`תרגם את הקטע הזה ${/[\u0590-\u05FF]/.test(t) ? 'לאנגלית' : 'לעברית'}, בלי הסברים:\n«${t}»`);
    else if (k === 'ex') askAI(`הסבר בפשטות ובקצרה מה כתוב כאן:\n«${t}»`);
    else if (k === 'rd') { try { speechSynthesis.cancel(); const u = new SpeechSynthesisUtterance(t); u.lang = /[\u0590-\u05FF]/.test(t) ? 'he-IL' : /[\u0600-\u06FF]/.test(t) ? 'ar' : /[\u0400-\u04FF]/.test(t) ? 'ru-RU' : 'en-US'; speechSynthesis.speak(u); } catch (_) {} }
    else if (k === 'cp') { try { await navigator.clipboard.writeText(t); if (typeof toast === 'function') toast('הועתק'); } catch (_) { document.execCommand && document.execCommand('copy'); } }
    if (k !== 'rd') { getSelection().removeAllRanges(); hideBub(); }
  });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && !bub.hidden) hideBub(); });
})();
