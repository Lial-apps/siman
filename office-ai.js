/* =====================================================================
   סימן — office-ai.js
   The Word, Excel and PowerPoint editors, one step closer to Office:
   · the Office look: the app's own color (Word blue, Excel green, PowerPoint orange), a "קובץ" menu
     and quick-access buttons (save, undo, redo) in the title bar
   · "ספרו לי מה לעשות": a box in the title bar that finds any ribbon command by name, or passes the
     request to the assistant ("הכן לי כותרת לקובץ", "עצב את הטבלה")
   · speaking instead of typing (the microphone), in the box and in the assistant panel
   · the assistant's actions: it reads the file, answers with a list of changes, and the editor makes
     them. One request = one step, so "ביטול" takes the whole request back.
   API: window.OFFICE_AI = { run(question, opts), revert(result), kind(), available() }
   ===================================================================== */
(() => {
'use strict';
const ED = {
  docx: { id: 'wed', api: 'WORD_EDITOR', p: 'we', app: 'Word', ribbon: 'weRibbon', cls: 'wb' },
  xlsx: { id: 'xed', api: 'XLSX_EDITOR', p: 'xe', app: 'Excel', ribbon: 'xeRibbon', cls: 'xb' },
  pptx: { id: 'ped', api: 'PPTX_EDITOR', p: 'pe', app: 'PowerPoint', ribbon: 'peRibbon', cls: 'wb' },
};
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
const h = html => { const t = document.createElement('template'); t.innerHTML = html.trim(); return t.content.firstElementChild; };
const liveKind = () => Object.keys(ED).find(k => document.getElementById(ED[k].id));
const apiOf = k => k && window[ED[k].api];
const toast = m => { if (typeof window.toast === 'function') window.toast(m); else console.log(m); };

/* ------------------------------------------------------------ the Office look */
const CSS = `
#wed.oa{--oa:#2B579A;--oa2:#1F4B8E;--oaL:#DCE6F4;--oaB:#9DB5DB}
#xed.oa{--oa:#217346;--oa2:#185C37;--oaL:#DDEEE3;--oaB:#93C3A4}
#ped.oa{--oa:#B7472A;--oa2:#963A22;--oaL:#F7E1DA;--oaB:#E0A795}
.oa .we-top,.oa .xe-top,.oa .pe-top,.oa .we-tabs,.oa .xe-tabs,.oa .pe-tabs{background:var(--oa)!important}
.oa .we-top,.oa .xe-top,.oa .pe-top{min-height:44px!important;padding-top:5px!important;padding-bottom:5px!important}
.oa .we-tabs button,.oa .xe-tabs button,.oa .pe-tabs button{color:rgba(255,255,255,.9)!important}
.oa .we-tabs button:hover,.oa .xe-tabs button:hover,.oa .pe-tabs button:hover{background:rgba(255,255,255,.12)!important}
.oa .we-tabs button[aria-selected=true],.oa .xe-tabs button[aria-selected=true],.oa .pe-tabs button[aria-selected=true]{color:var(--oa)!important}
.oa .we-tabs button[aria-selected=true]::after,.oa .xe-tabs button[aria-selected=true]::after,.oa .pe-tabs button[aria-selected=true]::after{background:var(--oa)!important}
.oa .we-tabs button.ctx,.oa .pe-tabs button.ctx{color:#FFE9A8!important}
.oa .we-brand b,.oa .xe-brand b,.oa .pe-brand b{background:none!important;font-weight:600!important}
.oa .we-mark,.oa .xe-mark,.oa .pe-mark{background:#fff!important;color:var(--oa)!important}
.oa .we-namebox,.oa .xe-namebox,.oa .pe-namebox{background:rgba(255,255,255,.14)!important;border-color:rgba(255,255,255,.28)!important}
.oa .we-namebox:focus-within,.oa .xe-namebox:focus-within,.oa .pe-namebox:focus-within{border-color:#fff!important}
.oa .we-ext,.oa .xe-ext,.oa .pe-ext{color:rgba(255,255,255,.75)!important}
.oa .we-tb,.oa .xe-tb,.oa .pe-tb{border-color:rgba(255,255,255,.35)!important;background:rgba(255,255,255,.1)!important}
.oa .we-tb.pri,.oa .xe-tb.pri,.oa .pe-tb.pri{background:#fff!important;border-color:#fff!important;color:var(--oa)!important}
.oa .we-x,.oa .xe-x,.oa .pe-x{color:#fff!important}
.oa .we-state.dirty,.oa .xe-state.dirty,.oa .pe-state.dirty{color:#FFE9A8!important}
#wed.oa .wb[aria-pressed=true],#wed.oa .wb.on,#ped.oa .wb[aria-pressed=true],#ped.oa .wb.on,#xed.oa .xb[aria-pressed=true]{background:var(--oaL)!important;border-color:var(--oaB)!important}
.oa .we-status input[type=range],.oa .pe-status input[type=range]{accent-color:var(--oa)!important}
.oa .btn.ai{border-color:rgba(255,255,255,.35)!important;color:#fff!important}
/* the File tab */
.oa .oa-file{color:#fff!important;font-weight:600;background:rgba(0,0,0,.12)!important;border-radius:8px 8px 0 0}
.oa .oa-file[aria-expanded=true]{background:#fff!important;color:var(--oa)!important}
.oa-menu{position:fixed;z-index:1200;min-width:240px;background:#fff;border:1px solid #D5DBE5;border-radius:12px;box-shadow:0 18px 44px -14px rgba(20,30,50,.45);padding:6px;direction:rtl;font:14px/1.3 Heebo,Rubik,system-ui,sans-serif;color:#1E2633}
.oa-menu[hidden]{display:none}
.oa-menu button{display:flex;align-items:center;gap:10px;width:100%;border:0;background:none;padding:9px 12px;border-radius:8px;cursor:pointer;text-align:start;font:inherit;color:inherit}
.oa-menu button:hover,.oa-menu button.on{background:#F0F3F8}
.oa-menu button kbd{margin-inline-start:auto;font:12px system-ui;color:#6B7689;direction:ltr}
.oa-menu .sep{height:1px;background:#E6EAF0;margin:5px 4px}
.oa-menu svg{width:18px;height:18px;flex:none}
/* quick access */
.oa-qat{display:flex;gap:2px;margin-inline:2px}
.oa-qat button{width:30px;height:30px;border:0;border-radius:6px;background:none;color:#fff;cursor:pointer;display:grid;place-items:center}
.oa-qat button:hover{background:rgba(255,255,255,.16)}
.oa-qat button:disabled{opacity:.45;cursor:default}
.oa-qat svg{width:17px;height:17px}
/* tell me */
.oa-tell{position:relative;display:flex;align-items:center;flex:0 1 380px;min-width:150px;height:32px;margin-inline:8px;background:rgba(255,255,255,.16);border:1px solid rgba(255,255,255,.3);border-radius:7px;color:#fff}
.oa-tell:focus-within{background:#fff;color:#1E2633;border-color:#fff}
.oa-tell>svg{width:16px;height:16px;flex:none;margin-inline:9px 2px;opacity:.9}
.oa-tell input{flex:1;min-width:0;height:100%;border:0;background:none;color:inherit;font:inherit;font-size:14px;outline:none;padding:0 4px}
.oa-tell input::placeholder{color:rgba(255,255,255,.85)}
.oa-tell:focus-within input::placeholder{color:#7A8597}
.oa-mic{width:30px;height:30px;flex:none;border:0;border-radius:6px;background:none;color:inherit;cursor:pointer;display:grid;place-items:center;margin-inline-end:1px}
.oa-mic svg{width:17px;height:17px}
.oa-mic:hover{background:rgba(127,127,127,.18)}
.oa-mic.rec{color:#fff!important;background:#D93025!important;animation:oaPulse 1.2s ease-in-out infinite}
@keyframes oaPulse{50%{box-shadow:0 0 0 5px rgba(217,48,37,.25)}}
@media (prefers-reduced-motion:reduce){.oa-mic.rec{animation:none}}
.oa-drop{position:fixed;z-index:1200;width:min(440px,calc(100vw - 16px));max-height:min(420px,70vh);overflow:auto;background:#fff;border:1px solid #D5DBE5;border-radius:12px;box-shadow:0 18px 44px -14px rgba(20,30,50,.45);padding:6px;direction:rtl;font:14px/1.35 Heebo,Rubik,system-ui,sans-serif;color:#1E2633}
.oa-drop[hidden]{display:none}
.oa-drop h5{margin:6px 10px 4px;font-size:12px;font-weight:600;color:#6B7689}
.oa-drop button{display:flex;align-items:center;gap:10px;width:100%;border:0;background:none;padding:8px 10px;border-radius:8px;cursor:pointer;text-align:start;font:inherit;color:inherit}
.oa-drop button.on,.oa-drop button:hover{background:#EEF2F8}
.oa-drop button .ico{width:26px;height:26px;border-radius:6px;display:grid;place-items:center;background:#F0F3F8;flex:none;color:#3A4558}
.oa-drop button .ico svg{width:16px;height:16px}
.oa-drop button small{display:block;color:#6B7689;font-size:12px}
.oa-drop button.ask .ico{background:linear-gradient(135deg,#FFE47A,#FFC93C);color:#3B2A00}
.oa-drop .tab{margin-inline-start:auto;font-size:11.5px;color:#6B7689;white-space:nowrap}
@media (max-width:860px){.oa-qat{display:none}
  .oa .we-top,.oa .xe-top,.oa .pe-top{position:relative}
  .oa-tell{flex:0 0 34px;min-width:34px;width:34px;margin-inline:2px;cursor:pointer;justify-content:center}
  .oa-tell>svg{margin:0}
  .oa-tell input{width:0;flex:0 0 0;padding:0}
  .oa-tell .oa-mic{display:none}
  .oa-tell:focus-within{position:absolute;inset-inline:8px;top:6px;bottom:6px;height:auto;width:auto;z-index:6;justify-content:flex-start;cursor:text}
  .oa-tell:focus-within>svg{margin-inline:9px 2px}
  .oa-tell:focus-within input{flex:1;width:auto;padding:0 4px}
  .oa-tell:focus-within .oa-mic{display:grid}}
/* the assistant panel: speaking, and the "undo" of a request */
#aiPanel .oa-mic{color:#3A4558;height:38px;width:38px;border:1px solid #D5DBE5;background:#fff}
.ai-res{display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-top:6px;font-size:12.5px;color:#3A4558}
.ai-res .done{display:inline-flex;align-items:center;gap:5px;background:#E6F4EA;color:#17603A;border-radius:999px;padding:3px 10px;font-weight:600}
.ai-res .done.none{background:#F1F3F6;color:#555F70}
.ai-res button{border:1px solid #CBD3DF;background:#fff;border-radius:999px;padding:3px 12px;cursor:pointer;font:inherit;color:#1E2633}
.ai-res button:hover{background:#F0F3F8}
.ai-res button:disabled{opacity:.5;cursor:default}
`;
function injectCss() { if (document.getElementById('oaCss')) return; const s = document.createElement('style'); s.id = 'oaCss'; s.textContent = CSS; document.head.append(s); }

const I = {
  search: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" aria-hidden="true"><circle cx="8.5" cy="8.5" r="5"/><path d="M12.5 12.5L17 17"/></svg>',
  mic: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="7" y="2.5" width="6" height="10" rx="3"/><path d="M4.5 9.5a5.5 5.5 0 0 0 11 0M10 15v2.5"/></svg>',
  spark: '<svg viewBox="0 0 20 20" fill="currentColor" aria-hidden="true"><path d="M10 1.8l1.7 4.7 4.7 1.7-4.7 1.7L10 14.6l-1.7-4.7-4.7-1.7 4.7-1.7z"/><path d="M15.5 12.5l.8 2.2 2.2.8-2.2.8-.8 2.2-.8-2.2-2.2-.8 2.2-.8z" opacity=".75"/></svg>',
  cmd: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" aria-hidden="true"><rect x="3" y="3" width="14" height="14" rx="3"/><path d="M7 10h6M10 7v6"/></svg>',
  save: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round" aria-hidden="true"><path d="M4 3h9.5L17 6.5V17H4z"/><path d="M7 3v4h6V3M7 17v-5h6v5"/></svg>',
  undo: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M7 5L3.5 8.5 7 12"/><path d="M3.5 8.5H12a4.5 4.5 0 0 1 0 9H9"/></svg>',
  redo: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M13 5l3.5 3.5L13 12"/><path d="M16.5 8.5H8a4.5 4.5 0 0 0 0 9h3"/></svg>',
  pdf: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round" aria-hidden="true"><path d="M5 2.5h7l3.5 3.5v11.5H5z"/><path d="M12 2.5V6h3.5M7.5 13h5M7.5 10h5"/></svg>',
  print: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linejoin="round" aria-hidden="true"><path d="M5.5 7V3h9v4"/><rect x="3" y="7" width="14" height="7" rx="1.5"/><path d="M6 12h8v5H6z"/></svg>',
  sign: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 15c2.5-4 4-9 6-9 1.5 0-1 7 1 7s2-3 3.5-3c1 0 .5 2 2.5 2"/><path d="M3 17.5h14"/></svg>',
  close: '<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" aria-hidden="true"><path d="M5 5l10 10M15 5L5 15"/></svg>',
  play: '<svg viewBox="0 0 20 20" fill="currentColor" aria-hidden="true"><path d="M6 4l10 6-10 6z"/></svg>',
};

/* ------------------------------------------------------------ voice */
const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
const VOICE = { rec: null, btn: null };
function listen(btn, onText, onFinal) {
  if (!SR) { toast('הדפדפן הזה לא תומך בהכתבה. אפשר להשתמש ב-Chrome או ב-Edge'); return; }
  if (VOICE.rec) { const same = VOICE.btn === btn; try { VOICE.rec.stop(); } catch (_) {} if (same) return; }
  const rec = new SR();
  const docLang = (() => { const k = liveKind(); const t = k && document.getElementById(ED[k].id).textContent || ''; return /[\u0590-\u05FF]/.test(t.slice(0, 4000)) || !t ? 'he-IL' : (navigator.language || 'he-IL'); })();
  rec.lang = docLang; rec.interimResults = true; rec.continuous = false; rec.maxAlternatives = 1;
  let finalText = '';
  rec.onresult = e => {
    let interim = ''; for (let i = e.resultIndex; i < e.results.length; i++) { const r = e.results[i]; if (r.isFinal) finalText += r[0].transcript; else interim += r[0].transcript; }
    onText((finalText + interim).trim());
  };
  rec.onerror = e => { if (e.error === 'not-allowed' || e.error === 'service-not-allowed') toast('צריך לאשר לסימן גישה למיקרופון (בסמל המנעול שליד הכתובת)'); else if (e.error === 'no-speech') toast('לא שמעתי כלום. נסו שוב'); else if (e.error === 'network') toast('ההכתבה צריכה חיבור לאינטרנט'); };
  rec.onend = () => { btn.classList.remove('rec'); btn.setAttribute('aria-pressed', 'false'); VOICE.rec = null; VOICE.btn = null; if (finalText.trim()) onFinal(finalText.trim()); };
  try { rec.start(); } catch (_) { return; }
  VOICE.rec = rec; VOICE.btn = btn; btn.classList.add('rec'); btn.setAttribute('aria-pressed', 'true');
}
const micBtn = label => { const b = h(`<button type="button" class="oa-mic" title="${label}" aria-label="${label}" aria-pressed="false">${I.mic}</button>`); if (!SR) b.hidden = true; return b; };

/* ------------------------------------------------------------ the ribbon's commands, by name */
function commands(root, k) {
  const out = [], seen = new Set();
  const tabName = {}; root.querySelectorAll(`.${ED[k].p}-tabs [data-tab]`).forEach(b => tabName[b.dataset.tab] = b.textContent.trim());
  for (const b of root.querySelectorAll(`#${ED[k].ribbon} button[data-cmd], #${ED[k].ribbon} select[data-in], #${ED[k].ribbon} input[data-in]`)) {
    const panel = b.closest('[data-panel]'); const tab = panel ? panel.dataset.panel : null;
    let label = (b.querySelector('span') && b.classList.contains('big') ? b.querySelector('span').textContent : '') || b.getAttribute('aria-label') || b.title || b.textContent;
    label = String(label || '').replace(/\s+/g, ' ').trim(); if (!label || label.length < 2) continue;
    const title = (b.title || '').replace(/\s+/g, ' ').trim();
    const key = label + '|' + (b.dataset.cmd || b.dataset.in);
    if (seen.has(key)) continue; seen.add(key);
    out.push({ el: b, label: label.replace(/\s*\(.*?\)\s*$/, ''), title: title !== label ? title : '', tab, tabLabel: tabName[tab] || '', kbd: (/\(([^)]*)\)\s*$/.exec(b.title || '') || [])[1] || '' });
  }
  for (const [t, l] of Object.entries(tabName)) out.push({ tabOnly: t, label: 'לשונית ' + l, tab: t, tabLabel: l });
  return out;
}
const norm = s => String(s || '').toLowerCase().replace(/[״"׳'.,:;!?()]/g, '').replace(/\s+/g, ' ').trim();
const SYN = [['כותרת', 'heading title'], ['מודגש', 'בולט עבה bold'], ['נטוי', 'italic'], ['קו תחתי', 'underline'], ['צבע', 'color'], ['טבלה', 'table'], ['תמונה', 'image picture'], ['מספור', 'numbers'], ['תבליטים', 'bullets רשימה'], ['שמירה', 'save'], ['הדפסה', 'print'], ['חיפוש', 'find search'], ['החלפה', 'replace'], ['סכום', 'sum'], ['מיון', 'sort'], ['סינון', 'filter'], ['שקף', 'slide'], ['מיזוג', 'merge'], ['גבולות', 'borders מסגרת'], ['יישור', 'align'], ['זום', 'zoom הגדלה']];
function score(c, q) {
  const L = norm(c.label + ' ' + c.title + ' ' + c.tabLabel); const w = norm(q).split(' ').filter(Boolean); if (!w.length) return 0;
  let s = 0;
  for (const x of w) {
    if (L.includes(x)) s += norm(c.label).startsWith(x) ? 3 : 2;
    else { const syn = SYN.find(([he, en]) => (he.includes(x) || en.split(' ').some(e => e.startsWith(x))) && L.includes(he)); if (syn) s += 1; else return 0; }
  }
  return s;
}
function runCommand(root, k, c) {
  const tabBtn = c.tab && root.querySelector(`.${ED[k].p}-tabs [data-tab="${c.tab}"]`);
  if (tabBtn && !tabBtn.hidden && tabBtn.getAttribute('aria-selected') !== 'true') tabBtn.click();
  if (c.tabOnly) return;
  const el = c.el; if (!el || !el.isConnected) return;
  setTimeout(() => {
    if (el.tagName === 'SELECT' || el.tagName === 'INPUT') { el.focus(); if (el.showPicker) try { el.showPicker(); } catch (_) {} return; }
    el.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    el.click();
  }, 40);
}

/* ------------------------------------------------------------ the title bar: File, quick access, tell me */
const UI = { root: null, k: null, drop: null, menu: null, items: [], at: 0, input: null };
function attach(root) {
  const k = Object.keys(ED).find(x => ED[x].id === root.id); if (!k || root.dataset.oa) return;
  root.dataset.oa = '1'; root.classList.add('oa'); injectCss();
  const P = ED[k].p, top = root.querySelector(`.${P}-top`), tabs = root.querySelector(`.${P}-tabs`); if (!top || !tabs) return;
  const q = id => root.querySelector('#' + P + id);
  // the File tab
  const file = h(`<button type="button" class="oa-file" aria-haspopup="menu" aria-expanded="false">קובץ</button>`);
  tabs.prepend(file);
  file.addEventListener('mousedown', e => e.preventDefault());
  file.onclick = e => { e.stopPropagation(); UI.menu && !UI.menu.hidden ? closeMenu() : openMenu(root, k, file); };
  // quick access: save, undo, redo
  const qat = h(`<div class="oa-qat" role="toolbar" aria-label="גישה מהירה"><button type="button" data-a="save" title="שמירה (Ctrl+S)" aria-label="שמירה">${I.save}</button><button type="button" data-a="undo" title="ביטול (Ctrl+Z)" aria-label="ביטול">${I.undo}</button><button type="button" data-a="redo" title="חזרה (Ctrl+Y)" aria-label="חזרה">${I.redo}</button></div>`);
  const brand = top.querySelector(`.${P}-brand`); (brand || top.firstChild).after(qat);
  qat.addEventListener('mousedown', e => e.preventDefault());
  qat.onclick = e => { const b = e.target.closest('[data-a]'); if (!b) return; if (b.dataset.a === 'save') { const s = q('Save'); if (s) s.click(); return; }
    const r = root.querySelector(`#${ED[k].ribbon} [data-cmd="${b.dataset.a}"]`); if (r && !r.disabled) r.click(); };
  // undo/redo state follows the ribbon's own buttons
  const syncQat = () => { for (const a of ['undo', 'redo']) { const r = root.querySelector(`#${ED[k].ribbon} [data-cmd="${a}"]`); const b = qat.querySelector(`[data-a="${a}"]`); if (r && b) b.disabled = !!r.disabled; } };
  new MutationObserver(syncQat).observe(root.querySelector('#' + ED[k].ribbon), { subtree: true, attributes: true, attributeFilter: ['disabled'] }); syncQat();
  // tell me
  const tell = h(`<div class="oa-tell" role="search">${I.search}<input type="text" placeholder="חיפוש פקודה, או בקשה מהעוזר (Alt+Q)" aria-label="חיפוש פקודה או בקשה מהעוזר" autocomplete="off" spellcheck="false" dir="auto" aria-autocomplete="list" aria-expanded="false"></div>`);
  const mic = micBtn('הכתבה: אמרו לעוזר מה לעשות'); tell.append(mic);
  const name = top.querySelector(`.${P}-namebox`); (name || qat).after(tell);
  const input = tell.querySelector('input');
  tell.addEventListener('mousedown', e => { if (e.target !== input && !e.target.closest('.oa-mic')) { e.preventDefault(); input.focus(); } });
  input.addEventListener('input', () => showDrop(root, k, input));
  input.addEventListener('focus', () => showDrop(root, k, input));
  input.addEventListener('blur', () => setTimeout(() => { if (document.activeElement !== input) hideDrop(); }, 150));
  mic.addEventListener('mousedown', e => e.preventDefault());
  mic.onclick = () => listen(mic, t => { input.value = t; showDrop(root, k, input); }, t => { input.value = t; askAssistant(t); input.value = ''; hideDrop(); });
  root._oa = { input, k };
}
function openMenu(root, k, anchor) {
  closeMenu(); const P = ED[k].p, q = id => root.querySelector('#' + P + id);
  const items = [
    ['save', I.save, 'שמירה', 'Ctrl+S', () => q('Save') && q('Save').click()],
    ['pdf', I.pdf, 'שמירה כ-PDF', '', () => q('Pdf') && q('Pdf').click()],
    ['print', I.print, 'הדפסה', 'Ctrl+P', () => { const b = root.querySelector(`#${ED[k].ribbon} [data-cmd="print"]`); if (b) b.click(); else if (q('Pdf')) q('Pdf').click(); }],
    k === 'pptx' ? ['play', I.play, 'הצגת המצגת', 'F5', () => q('Play') && q('Play').click()] : null,
    '-',
    ['desk', I.sign, 'לסימון, חתימה ושליחה', '', () => q('Desk') && q('Desk').click()],
    ['ai', I.spark, 'העוזר: עיצוב ועריכה לפי בקשה', '', () => { if (root._oa) root._oa.input.focus(); }],
    '-',
    ['close', I.close, 'סגירת הקובץ', '', () => q('Close') && q('Close').click()],
  ].filter(Boolean);
  const m = h(`<div class="oa-menu" role="menu" aria-label="קובץ"></div>`);
  for (const it of items) {
    if (it === '-') { m.append(h('<div class="sep" role="separator"></div>')); continue; }
    const b = h(`<button type="button" role="menuitem">${it[1]}<span>${esc(it[2])}</span>${it[3] ? `<kbd>${it[3]}</kbd>` : ''}</button>`);
    b.onclick = () => { closeMenu(); it[4](); }; m.append(b);
  }
  document.body.append(m); UI.menu = m; anchor.setAttribute('aria-expanded', 'true'); UI.menuFor = anchor;
  const r = anchor.getBoundingClientRect(); m.style.top = (r.bottom + 2) + 'px';
  m.style.left = Math.max(6, Math.min(r.right - m.offsetWidth, innerWidth - m.offsetWidth - 6)) + 'px';
  m.querySelector('button').focus();
}
function closeMenu() { if (UI.menu) { UI.menu.remove(); UI.menu = null; } if (UI.menuFor) { UI.menuFor.setAttribute('aria-expanded', 'false'); UI.menuFor = null; } }
function showDrop(root, k, input) {
  const v = input.value.trim();
  if (!UI.drop) { UI.drop = h('<div class="oa-drop" role="listbox" aria-label="פקודות והצעות" hidden></div>'); document.body.append(UI.drop);
    UI.drop.addEventListener('mousedown', e => e.preventDefault());
    UI.drop.addEventListener('click', e => { const b = e.target.closest('button[data-i]'); if (b) pick(+b.dataset.i); }); }
  UI.root = root; UI.k = k; UI.input = input;
  const d = UI.drop; d.innerHTML = ''; UI.items = [];
  const add = (it, html) => { const i = UI.items.length; UI.items.push(it); const b = h(`<button type="button" role="option" data-i="${i}" class="${it.ask ? 'ask' : ''}">${html}</button>`); d.append(b); };
  if (v) {
    add({ ask: v }, `<span class="ico">${I.spark}</span><span><b>בקשה מהעוזר</b><small dir="auto">${esc(v)}</small></span>`);
    const cmds = commands(root, k).map(c => ({ c, s: score(c, v) })).filter(x => x.s > 0).sort((a, b) => b.s - a.s).slice(0, 7);
    if (cmds.length) { d.append(h('<h5>פקודות</h5>')); for (const { c } of cmds) add({ cmd: c }, `<span class="ico">${I.cmd}</span><span>${esc(c.label)}${c.title ? `<small>${esc(c.title.slice(0, 80))}</small>` : ''}</span><span class="tab">${esc(c.tabLabel)}</span>`); }
    // a short request that names a command is most likely the command; anything longer is a request
    UI.at = cmds.length && v.split(/\s+/).length <= 2 && cmds[0].s >= 3 ? 1 : 0;
  } else {
    d.append(h('<h5>אפשר לבקש מהעוזר, למשל</h5>'));
    for (const s of SUGGEST[k]) add({ ask: s }, `<span class="ico">${I.spark}</span><span dir="auto">${esc(s)}</span>`);
    UI.at = -1;
  }
  mark();
  const r = input.closest('.oa-tell').getBoundingClientRect();
  d.hidden = false; d.style.top = (r.bottom + 4) + 'px';
  d.style.left = Math.max(8, Math.min(r.right - d.offsetWidth, innerWidth - d.offsetWidth - 8)) + 'px';
  input.setAttribute('aria-expanded', 'true');
}
function mark() { if (!UI.drop) return; UI.drop.querySelectorAll('button[data-i]').forEach(b => { const on = +b.dataset.i === UI.at; b.classList.toggle('on', on); b.setAttribute('aria-selected', on); if (on) b.scrollIntoView({ block: 'nearest' }); }); }
function hideDrop() { if (UI.drop) UI.drop.hidden = true; if (UI.input) UI.input.setAttribute('aria-expanded', 'false'); }
function pick(i) {
  const it = UI.items[i]; if (!it) return; const input = UI.input;
  hideDrop();
  if (it.cmd) { input.value = ''; input.blur(); runCommand(UI.root, UI.k, it.cmd); return; }
  if (it.ask) { input.value = ''; input.blur(); askAssistant(it.ask); }
}
const SUGGEST = {
  docx: ['הכן לי כותרת לקובץ', 'עצב את כל הכותרות בכחול כהה ובגופן Heebo', 'הוסף בסוף טבלה מסודרת עם שורת כותרת צבועה', 'סדר את המסמך: יישור לשני הצדדים ורווח שורה 1.5', 'הוסף מספרי עמודים בכותרת התחתונה'],
  xlsx: ['עצב את השורה הראשונה ככותרת טבלה', 'הוסף שורת סיכום עם סכום לכל עמודה', 'הפוך את הנתונים לטבלה מעוצבת עם פסים', 'הקפא את השורה העליונה והתאם את רוחב העמודות', 'עצב את עמודת המחירים כשקלים'],
  pptx: ['הכן שקף פתיחה עם כותרת למצגת', 'הוסף שקף סיכום עם 3 נקודות עיקריות', 'שנה את הרקע של כל השקפים לכחול כהה ואת הטקסט ללבן', 'הגדל את הכותרות ועשה אותן מודגשות', 'הוסף שקף תוכן עניינים לפי השקפים'],
};
function askAssistant(text) {
  text = String(text || '').trim(); if (!text) return;
  if (typeof window.SIMAN_ASK === 'function') window.SIMAN_ASK(text);
  else run(text).then(r => toast(r.reply || (r.done ? 'בוצע' : 'לא בוצע שינוי'))).catch(e => toast(e.message || String(e)));
}
// keys inside the box are the box's: the editors listen on the whole window
for (const ev of ['keydown', 'keyup', 'keypress', 'paste', 'copy', 'cut']) addEventListener(ev, e => {
  const box = e.target && e.target.closest && e.target.closest('.oa-tell,.oa-menu');
  if (!box) {
    if (ev === 'keydown' && e.altKey && !e.ctrlKey && !e.metaKey && e.code === 'KeyQ') { const k = liveKind(); const r = k && document.getElementById(ED[k].id); if (r && r._oa) { e.preventDefault(); e.stopImmediatePropagation(); r._oa.input.focus(); r._oa.input.select(); } }
    return;
  }
  if (box.classList.contains('oa-menu')) {
    e.stopImmediatePropagation(); if (ev !== 'keydown') return;
    const bs = [...box.querySelectorAll('button')], i = bs.indexOf(document.activeElement);
    if (e.key === 'ArrowDown') { e.preventDefault(); bs[(i + 1) % bs.length].focus(); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); bs[(i - 1 + bs.length) % bs.length].focus(); }
    else if (e.key === 'Escape' || e.key === 'Tab') { e.preventDefault(); const a = UI.menuFor; closeMenu(); if (a) a.focus(); }
    return;
  }
  e.stopImmediatePropagation();
  if (ev !== 'keydown' || e.isComposing) return;
  const n = UI.items.length;
  if (e.key === 'ArrowDown') { e.preventDefault(); if (UI.drop && UI.drop.hidden) showDrop(UI.root, UI.k, e.target); UI.at = (UI.at + 1) % Math.max(1, n); mark(); }
  else if (e.key === 'ArrowUp') { e.preventDefault(); UI.at = (UI.at - 1 + n) % Math.max(1, n); mark(); }
  else if (e.key === 'Enter') { e.preventDefault(); const v = e.target.value.trim(); if (UI.at >= 0 && UI.items[UI.at]) pick(UI.at); else if (v) { e.target.value = ''; hideDrop(); e.target.blur(); askAssistant(v); } }
  else if (e.key === 'Escape') { e.preventDefault(); e.target.value = ''; hideDrop(); e.target.blur(); }
}, true);
addEventListener('pointerdown', e => { if (UI.menu && !UI.menu.contains(e.target) && e.target !== UI.menuFor) closeMenu(); }, true);
addEventListener('resize', () => { hideDrop(); closeMenu(); });

/* ------------------------------------------------------------ the assistant's actions */
const COMMON = `You are the editing assistant inside the Siman app's ${'{APP}'} editor. The user wrote (or said) a request about the open file.
The file is described in the document text above. Decide what to change and answer with actions the editor will carry out.
Return ONLY one JSON object, without markdown fences, without page references, exactly in this shape:
{"reply":"one or two short sentences in the user's language saying what you did, or the answer if the user only asked a question","actions":[ ... ]}
Rules:
- When the user asks for a change (create, add, write, design, format, color, fix, organize, translate, delete…), DO it with actions. Never just explain how to do it.
- When the user only asks a question about the file, answer in "reply" and return "actions":[].
- New text goes in the language of the file (Hebrew unless the file is in another language), written well and fitting the content.
- Colors are 6-digit hex without # (e.g. "1F3864"). Sizes are in points.
- Keep everything the user did not ask to change. Do not delete content unless asked.
- At most 40 actions. Use only the actions listed below; unknown fields are ignored.
- If the request cannot be done with these actions, say so briefly in "reply" and return the closest useful actions or none.
`;
const SPEC = {
  docx: `Blocks of the document are numbered #0, #1… in the description; refer to them with "ref":N or "refs":[N,…].
"refs" may also be one of: "all", "headings", "body" (non-heading paragraphs), "tables", "selection" (where the user's cursor is), "first", "last", or a range "3-7".
Actions:
{"op":"insert","at":"start"|"end"|"cursor"} or {"op":"insert","ref":N,"position":"before"|"after"} — plus "blocks":[BLOCK,…]
  BLOCK = {"type":"title"|"subtitle"|"h1"|"h2"|"h3"|"quote"|"p"|"bullet"|"number"|"hr"|"pagebreak","text":"…","style":FMT}
        | {"type":"table","rows":[["header 1","header 2"],["…","…"]],"header":true,"headerColor":"1F3864","banded":true}
{"op":"format","refs":…,"style":FMT,"find":"optional exact words — only these words get the look"}
  FMT = {"bold":true|false,"italic":…,"underline":…,"strike":…,"font":"David"|"Arial"|"Heebo"|"Rubik"|"Frank Ruehl"|"Calibri"|…,"size":14,"color":"RRGGBB","highlight":"RRGGBB"|null,
         "align":"right"|"center"|"left"|"justify","dir":"rtl"|"ltr","lineHeight":1.5,"spaceBefore":6,"spaceAfter":6,"indent":1.0 (cm),"fill":"RRGGBB" (table cells)}
{"op":"style","refs":…,"name":"title"|"subtitle"|"h1"|"h2"|"h3"|"quote"|"normal"}
{"op":"list","refs":…,"kind":"bullet"|"number"|"none"}
{"op":"text","ref":N,"text":"the new full text of that block"}
{"op":"delete","refs":…}
{"op":"replace","find":"…","with":"…","all":true}
{"op":"table","refs":[N]|"tables","style":"header"|"banded"|"grid"|"none","headerColor":"RRGGBB","bandColor":"RRGGBB","headerText":"FFFFFF","width":"full"}
{"op":"header"|"footer","text":"…","align":"center"|"right"|"left"}
{"op":"pageNumbers","where":"footer"|"header","format":"xofy"|"page","align":"center"}
{"op":"page","orientation":"portrait"|"landscape","margins":"normal"|"narrow"|"moderate"|"wide"}
{"op":"toc"}  (a table of contents from the headings, at the start)
Examples: "הכן לי כותרת לקובץ" → {"op":"insert","at":"start","blocks":[{"type":"title","text":"<a short fitting title from the content>"}]}
"הדגש את המילה חשוב" → {"op":"format","refs":"all","find":"חשוב","style":{"bold":true}}`,
  xlsx: `Cells use A1 notation. A "range" is like "A1:D1", "B:B", "3:3", "A1", "used" (the whole used range) or "selection".
Actions (add "sheet":"name" to work on another sheet):
{"op":"set","cells":{"A1":"text","B2":42,"C2":"=SUM(B2:B9)"}}  or  {"op":"set","range":"A1","values":[["row1 col1","row1 col2"],["…","…"]]}  (a block starting at that cell)
{"op":"format","range":"A1:D1","style":{"bold":true,"italic":…,"underline":…,"size":12,"font":"Arial","color":"FFFFFF","fill":"1F4E78","align":"center"|"right"|"left","valign":"middle"|"top"|"bottom","wrap":true,"border":"all"|"outside"|"thick"|"bottom"|"none","numFormat":"currency"|"number"|"thousands"|"integer"|"percent"|"date"|"time"|"text"|"<Excel format code>"}}
{"op":"table","range":"A1:D9","headerColor":"1F4E78","headerText":"FFFFFF","banded":true,"bandColor":"DDEBF7"}  (header row + borders + bands + fitted widths)
{"op":"merge","range":"A1:D1"} / {"op":"unmerge","range":"A1:D1"}
{"op":"insertRows","at":5,"count":1} / {"op":"deleteRows","at":5,"count":1} / {"op":"insertCols","at":"C","count":1} / {"op":"deleteCols","at":"C","count":1}
{"op":"colWidth","cols":"A:D","width":"auto"|15} / {"op":"rowHeight","rows":"1","height":28}
{"op":"freeze","rows":1,"cols":0}
{"op":"sort","range":"A1:D9","by":"B","desc":false,"header":true}
{"op":"clear","range":"A1:B2","formats":false}
{"op":"direction","rtl":true}
{"op":"addSheet","name":"…"} (stays on the current sheet: give later actions for the new sheet "sheet":"<its name>") / {"op":"renameSheet","name":"…"}
Formulas are written in English Excel syntax (SUM, AVERAGE, IF, VLOOKUP, XLOOKUP, COUNTIF…), starting with "=". Inserted rows shift the rows below them: plan cell addresses accordingly.
Example: "הוסף שורת סיכום" (data in A1:C9) → [{"op":"set","cells":{"A10":"סה\\"כ","B10":"=SUM(B2:B9)","C10":"=SUM(C2:C9)"}},{"op":"format","range":"A10:C10","style":{"bold":true,"border":"top"}}]`,
  pptx: `Slides are numbered from 1. Inside a slide, shapes are numbered from 0 as in the description ("צורה N"). "shape" may also be "title", "subtitle", "body", "text" (all text shapes), "picture", "table" or "all".
"slide" may be a number, a list [1,3], a range "2-5", "all", "current" or "last". Positions and sizes (x,y,w,h) are percents of the slide.
Actions:
{"op":"addSlide","layout":"title"|"content"|"section"|"titleOnly"|"blank"|"two","title":"…","subtitle":"…","bullets":["…","…"],"after":N|"end"|"start","background":"RRGGBB","style":FMT}
{"op":"text","slide":N,"shape":"title"|N,"text":"…"}  or  "paragraphs":["line 1","line 2"] (each line a bullet/paragraph)
{"op":"format","slide":N|"all","shape":"title"|"body"|"all"|N,"style":FMT}
  FMT = {"bold":…,"italic":…,"underline":…,"size":40,"font":"Heebo"|"Arial"|…,"color":"RRGGBB","align":"right"|"center"|"left"|"justify","fill":"RRGGBB"|null (shape background),"line":"RRGGBB"|null}
{"op":"background","slide":N|"all","color":"RRGGBB"}
{"op":"addText","slide":N,"text":"…","x":10,"y":70,"w":80,"h":12,"style":FMT}
{"op":"addShape","slide":N,"shape":"rect"|"roundRect"|"ellipse"|"triangle"|"rightArrow"|"leftArrow"|"star5"|"chevron","x":…,"y":…,"w":…,"h":…,"color":"RRGGBB","text":"…","textColor":"FFFFFF"}
{"op":"move","slide":N,"shape":N,"x":…,"y":…,"w":…,"h":…}
{"op":"delete","slide":N,"shape":N} / {"op":"deleteSlide","slide":N} / {"op":"duplicateSlide","slide":N} / {"op":"moveSlide","from":N,"to":M} / {"op":"goto","slide":N}
Hebrew slides read right to left: align Hebrew text "right" unless centered.
Example: "הכן שקף פתיחה" → {"op":"addSlide","layout":"title","after":"start","title":"<title>","subtitle":"<subtitle>"}`,
};
function parseReply(raw) {
  let t = String(raw || '').trim().replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '');
  const a = t.indexOf('{'), b = t.lastIndexOf('}');
  if (a < 0 || b <= a) return null;
  let s = t.slice(a, b + 1);
  try { return JSON.parse(s); } catch (_) {}
  // page marks the server may add, and trailing commas
  s = s.replace(/\s*\[(?:עמ['׳]?|עמוד|p\.|pp\.)\s*[\d\s,–-]+\]/g, '').replace(/,\s*([}\]])/g, '$1');
  try { return JSON.parse(s); } catch (_) { return null; }
}
const cleanReply = t => String(t || '').replace(/\s*\[(?:עמ['׳]?|עמוד|עמודים|p\.|pp\.|page)\s*[\d\s,–-]+\]/gi, '').trim();
async function run(question, opts = {}) {
  const k = liveKind(); const A = apiOf(k);
  if (!k || !A || !A.assist) throw new Error('העוזר עובד בתוך עורך ה-Word, ה-Excel או ה-PowerPoint');
  if (typeof window.SIMAN_AI !== 'function') throw new Error('השירות החכם לא נטען. רעננו את הדף');
  const ctx = A.assist.context(); if (!ctx) throw new Error('העורך סגור');
  const prompt = COMMON.replace('{APP}', ED[k].app) + '\n' + SPEC[k] + `\n\nThe file name: "${ctx.name}".\nThe user's request: """${question}"""`;
  const raw = await window.SIMAN_AI({ mode: 'ask', text: ctx.text, question: prompt, history: opts.history || [] }, { onWait: opts.onWait, signal: opts.signal });
  const obj = parseReply(raw);
  if (!obj || typeof obj !== 'object') return { kind: k, reply: cleanReply(raw), done: 0, skipped: 0, raw };
  const acts = Array.isArray(obj.actions) ? obj.actions : obj.action ? [obj.action] : [];
  if (!liveKind() || liveKind() !== k) throw new Error('הקובץ נסגר לפני שהעוזר סיים');
  let res = { done: 0, skipped: 0, mark: null };
  if (acts.length) res = apiOf(k).assist.apply(acts);
  return { kind: k, reply: cleanReply(obj.reply || obj.answer || ''), done: res.done, skipped: res.skipped, mark: res.mark, asked: acts.length };
}
function revert(r) {
  if (!r || !r.mark) return false; const k = liveKind(); if (k !== r.kind) return false;
  const A = apiOf(k); return !!(A && A.assist && A.assist.revert(r.mark));
}

/* ------------------------------------------------------------ the assistant panel: a microphone */
function micInPanel() {
  const ask = document.querySelector('#aiPanel .ai-ask'); if (!ask || ask.querySelector('.oa-mic') || !SR) return;
  const q = document.getElementById('aiQ'), send = document.getElementById('aiSend');
  const b = micBtn('הכתבה: לדבר במקום להקליד'); ask.insertBefore(b, send);
  b.addEventListener('mousedown', e => e.preventDefault());
  b.onclick = () => { if (q.disabled) return; listen(b, t => { q.value = t; q.dispatchEvent(new Event('input')); }, t => { q.value = t; send.click(); }); };
}

/* ------------------------------------------------------------ the editors come and go */
const watch = new MutationObserver(ms => { for (const m of ms) for (const n of m.addedNodes) if (n.nodeType === 1 && Object.values(ED).some(e => e.id === n.id)) setTimeout(() => attach(n), 0);
  for (const m of ms) for (const n of m.removedNodes) if (n.nodeType === 1 && Object.values(ED).some(e => e.id === n.id)) { hideDrop(); closeMenu(); if (VOICE.rec) try { VOICE.rec.stop(); } catch (_) {} } });
watch.observe(document.body, { childList: true });
for (const e of Object.values(ED)) { const r = document.getElementById(e.id); if (r) attach(r); }
micInPanel(); setTimeout(micInPanel, 1500);

window.OFFICE_AI = { run, revert, kind: liveKind, available: () => !!(liveKind() && apiOf(liveKind()) && apiOf(liveKind()).assist), parseReply, _ui: UI };
})();
