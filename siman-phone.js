/* סימן — phone: "שמירה" in the Word / Excel / PowerPoint editors asks how to save (5.0.3)
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
    run(go === 'save' ? menu.dataset.save : go === 'pdf' ? menu.dataset.pdf : menu.dataset.desk);
  });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && !menu.hidden) close(); });
  addEventListener('resize', close);
})();
