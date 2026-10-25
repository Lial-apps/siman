/* סימן — phone helpers (5.0.5)
   1. "שמירה" in the Word / Excel / PowerPoint editors asks how to save
   2. after a translation into the document: the panel closes, the translated document shows, and a notice says so
   Word (or Excel / PowerPoint) or PDF, plus opening the file for signing. It calls the editors' own
   buttons, so saving works exactly as before. On a computer nothing changes. */
(function () {
  'use strict';
  const KINDS = {
    weSave: { pdf: 'wePdf', desk: 'weDesk', name: 'Word', tag: 'DOCX', c: '#2B579A' },
    xeSave: { pdf: 'xePdf', desk: 'xeDesk', name: 'Excel', tag: 'XLSX', c: '#217346' },
    peSave: { pdf: 'pePdf', desk: 'peDesk', name: 'PowerPoint', tag: 'PPTX', c: '#C43E1C' }
  };
  const phone = () => matchMedia('(max-width:720px)').matches;
  const menu = document.createElement('div');
  menu.className = 'sp-save'; menu.hidden = true; menu.setAttribute('role', 'menu');
  document.body.append(menu);
  const run = id => { const b = document.getElementById(id); if (b && typeof b.onclick === 'function') b.onclick.call(b, new MouseEvent('click')); };
  function close() { menu.hidden = true; }
  function open(btn, k) {
    menu.innerHTML = `
      <button type="button" role="menuitem" data-go="save"><span class="ic" style="background:${k.c}">${k.tag}</span><span>שמירה כקובץ ${k.name}<small>לעריכה בהמשך</small></span></button>
      <button type="button" role="menuitem" data-go="pdf"><span class="ic" style="background:#D93F2E">PDF</span><span>שמירה כקובץ PDF<small>לשליחה ולהדפסה</small></span></button>
      <button type="button" role="menuitem" data-go="tr"><span class="ic" style="background:#E0A100;color:#1B1C1F">א→A</span><span>תרגום הקובץ<small>לעברית, אנגלית, רוסית, ערבית, צרפתית או ספרדית</small></span></button>
      <hr>
      <button type="button" role="menuitem" class="sec" data-go="desk">פתיחה לחתימה וסימון</button>`;
    menu.dataset.save = btn.id; menu.dataset.pdf = k.pdf; menu.dataset.desk = k.desk;
    menu.hidden = false;
    const r = btn.getBoundingClientRect(), w = menu.offsetWidth, h = menu.offsetHeight;
    let x = Math.min(innerWidth - w - 8, Math.max(8, r.left + r.width / 2 - w / 2));
    let y = r.bottom + 8; if (y + h > innerHeight - 8) y = Math.max(8, r.top - h - 8);
    menu.style.left = x + 'px'; menu.style.top = y + 'px';
    const first = menu.querySelector('button'); if (first) first.focus({ preventScroll: true });
  }
  document.addEventListener('click', e => {
    const btn = e.target.closest && e.target.closest('#weSave,#xeSave,#peSave');
    if (btn && phone() && !btn.disabled) {
      e.preventDefault(); e.stopImmediatePropagation();
      if (!menu.hidden && menu.dataset.save === btn.id) { close(); return; }
      open(btn, KINDS[btn.id]); return;
    }
    if (!menu.hidden && !menu.contains(e.target)) close();
  }, true);
  menu.addEventListener('click', e => {
    const b = e.target.closest('[data-go]'); if (!b) return;
    const go = b.dataset.go; close();
    if (go === 'tr') { if (window.OFFICE_TR) window.OFFICE_TR.open({ weSave: 'docx', xeSave: 'xlsx', peSave: 'pptx' }[menu.dataset.save]); return; }
    run(go === 'save' ? menu.dataset.save : go === 'pdf' ? menu.dataset.pdf : menu.dataset.desk);
  });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && !menu.hidden) close(); });
  addEventListener('resize', close);
})();

/* ---------- 2. translation into the document, on a phone ----------
   The translation panel covers the screen, so the translated document stayed hidden behind it.
   When the translation is done: close the panel, show the document from its first page,
   and say it in a notice that screen readers announce (role="status") and that stays long enough to read. */
(function () {
  'use strict';
  const st = document.getElementById('tpStatus');
  if (!st) return;
  const phone = () => matchMedia('(max-width:720px)').matches;
  const calm = () => matchMedia('(prefers-reduced-motion: reduce)').matches;
  const note = document.createElement('div');
  note.className = 'sp-note'; note.setAttribute('role', 'status'); note.setAttribute('aria-live', 'polite'); note.hidden = true;
  note.innerHTML = '<p class="sp-note-t"></p><button type="button" class="sp-note-x" aria-label="סגירת ההודעה">×</button>';
  document.body.append(note);
  let hideT = 0;
  const hide = () => { note.hidden = true; clearTimeout(hideT); };
  note.querySelector('.sp-note-x').addEventListener('click', hide);
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && !note.hidden) hide(); });
  function show(msg) {
    note.querySelector('.sp-note-t').textContent = msg;
    note.hidden = false;
    clearTimeout(hideT); hideT = setTimeout(hide, 12000);
  }
  let last = '';
  new MutationObserver(() => {
    const t = st.textContent || '';
    if (!/^התרגום הוצב בתוך המסמך/.test(t)) { last = ''; return; }
    if (t === last || !phone()) return;
    last = t;
    setTimeout(() => {
      const panel = document.getElementById('tpanel');
      if (panel && !panel.classList.contains('hidden')) { const x = document.getElementById('tpClose'); if (x) x.click(); }
      const desk = document.getElementById('desk');
      if (desk) {
        desk.scrollTo({ top: 0, behavior: calm() ? 'auto' : 'smooth' });
        if (!desk.hasAttribute('tabindex')) desk.setAttribute('tabindex', '-1');
        desk.focus({ preventScroll: true }); // the panel that had the focus is gone
      }
      const m = t.match(/\((\d+) קטעים ב-(\d+) עמודים\)/);
      show(`התרגום מוכן${m ? ` (${m[2] === '1' ? 'עמוד אחד' : m[2] + ' עמודים'})` : ''}. לתיקון משפט: "עריכה". כדי לשמור את הקובץ המתורגם: "שמירה".`);
    }, 400);
  }).observe(st, { childList: true, characterData: true, subtree: true });
})();
