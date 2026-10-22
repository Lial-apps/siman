/* =====================================================================
   סימן 4.6 — siman-plus.js
   · several files open together (tabs) + actions on all of them
   · the assistant: questions about the document, summaries, key points
   · read aloud
   · smart scan: handwriting / scanned pages → a printed Word or PDF file
   · OCR in more languages (Russian, Arabic, French, Spanish)
   Loaded after the main script and uses its globals (S, openFile, buildPdf, pageText…).
   ===================================================================== */
(() => {
'use strict';
const AI = typeof AI_URL !== 'undefined' ? AI_URL : 'https://siman-ai.siman-app.workers.dev/';
const esc = t => String(t ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
const isBusy = () => !$('busy').classList.contains('hidden');
const h = html => { const t = document.createElement('template'); t.innerHTML = html.trim(); return t.content.firstElementChild; };
const ICON = {
  spark: '<svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8z"/><path d="M19 15l.8 2.2L22 18l-2.2.8L19 21l-.8-2.2L16 18l2.2-.8z"/></svg>',
  speak: '<svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 9.5h3.5L12 5v14l-4.5-4.5H4z"/><path d="M15.5 9a4 4 0 010 6M18 6.5a7.5 7.5 0 010 11"/></svg>',
  scan: '<svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 8V5a1 1 0 011-1h3M16 4h3a1 1 0 011 1v3M20 16v3a1 1 0 01-1 1h-3M8 20H5a1 1 0 01-1-1v-3"/><path d="M7.5 13c1.5-2.5 2.5-3 3-1.5s1.5 1.5 2.5-.5 2-2 3.5.5"/></svg>',
  plus: '<svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>',
};

/* ---------------------------------------------------------------- styles */
document.head.append(h(`<style id="simanPlusCss">
.tabs{flex:none;display:flex;align-items:stretch;gap:2px;min-height:40px;padding:0 10px;background:var(--bar);border-bottom:1px solid var(--line);overflow-x:auto;scrollbar-width:thin}
.tabs.hidden{display:none}
.tab{position:relative;display:flex;align-items:center;gap:8px;max-width:240px;min-width:0;padding:0 8px 0 6px;border:0;background:none;color:var(--muted);font:inherit;font-size:13.5px;cursor:pointer;border-radius:8px 8px 0 0}
.tab:hover{color:var(--ink);background:color-mix(in srgb,var(--ink) 5%,transparent)}
.tab:focus-visible{outline:2px solid var(--blue);outline-offset:-2px}
.tab .kd{flex:none;font-size:9.5px;font-weight:700;letter-spacing:.02em;padding:2px 4px;border-radius:4px;border:1px solid currentColor;opacity:.7}
.tab .nm{position:relative;z-index:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;padding:2px 1px}
.tab[aria-selected="true"]{color:var(--ink);font-weight:600}
.tab[aria-selected="true"] .nm::after{content:"";position:absolute;z-index:-1;inset-inline:-2px;top:55%;bottom:1px;background:var(--marker);border-radius:2px;transform:skewX(-12deg)}
.tab[aria-selected="true"]::before{content:"";position:absolute;inset-inline:6px;bottom:-1px;height:2px;background:var(--ink)}
.tab .dot{width:7px;height:7px;border-radius:50%;background:var(--ink);flex:none}
.tab .x{flex:none;width:22px;height:22px;border-radius:6px;display:grid;place-items:center;font-size:16px;line-height:1;color:var(--muted)}
.tab .x:hover{background:color-mix(in srgb,var(--ink) 12%,transparent);color:var(--ink)}
.tab.drop-before{box-shadow:inset 2px 0 0 var(--blue)}
.tab-add,.tab-all{flex:none;align-self:center;display:inline-flex;align-items:center;gap:6px;height:30px;padding:0 10px;margin-inline-start:4px;border:1px dashed var(--line);border-radius:8px;background:none;color:var(--muted);font:inherit;font-size:13px;cursor:pointer}
.tab-add svg{width:15px;height:15px}
.tab-add:hover,.tab-all:hover{color:var(--ink);border-color:var(--ink)}
.tab-all{border-style:solid;margin-inline-start:auto}
/* above the full Word / Excel editors */
body.ed-live .tabs{position:fixed;top:0;left:0;right:0;z-index:950;height:40px;box-shadow:0 1px 0 var(--line)}
body.ed-live #wed,body.ed-live #xed,body.ed-live #ped{top:40px!important}
body.ed-live #aiPanel{position:fixed;z-index:960;top:40px;bottom:0;inset-inline-end:0;width:min(400px,100vw);box-shadow:0 0 40px -12px rgba(0,0,0,.45);background:var(--bar)}
@media (max-width:720px){body.ed-live #aiPanel{top:auto;height:72vh;width:100vw;border-radius:16px 16px 0 0}}
.ed-ai{display:inline-flex;align-items:center;gap:6px}
.ed-ai svg{width:16px;height:16px;color:var(--marker,#FFD43B)}
.tab.ed .kd{color:#fff;border-color:transparent;opacity:1}
.tab.ed.docx .kd{background:#2B579A}
.tab.ed.xlsx .kd{background:#217346}
.tab.ed.pptx .kd{background:#C43E1C}
.tab-menu{position:fixed;z-index:60;min-width:230px;background:var(--bar);border:1px solid var(--line);border-radius:12px;box-shadow:0 18px 40px -16px rgba(0,0,0,.45);padding:6px}
.tab-menu button{display:block;width:100%;text-align:start;border:0;background:none;font:inherit;font-size:14px;color:var(--ink);padding:9px 12px;border-radius:8px;cursor:pointer}
.tab-menu button:hover,.tab-menu button:focus-visible{background:color-mix(in srgb,var(--marker) 45%,var(--bar));outline:none}
.tab-menu small{display:block;color:var(--muted);font-size:12px}
/* toolbar button of the assistant */
.bar .btn.ai{gap:6px;display:inline-flex;align-items:center}
.bar .btn.ai svg{width:16px;height:16px;color:var(--marker)}
/* the assistant panel */
.aip .ai-tools{display:flex;gap:8px;flex-wrap:wrap;padding:10px 14px;border-bottom:1px solid var(--line)}
.aip .ai-tools .btn{display:inline-flex;align-items:center;gap:6px}
.aip .ai-tools svg{width:16px;height:16px}
.ai-log{flex:1;overflow:auto;padding:14px;display:flex;flex-direction:column;gap:12px;min-height:0}
.ai-hello{color:var(--muted);font-size:14px;line-height:1.6}
.ai-hello b{color:var(--ink)}
.ai-chips{display:flex;flex-wrap:wrap;gap:6px}
.ai-chip{border:1px solid var(--line);background:var(--bar);color:var(--ink);font:inherit;font-size:13px;padding:6px 11px;border-radius:999px;cursor:pointer}
.ai-chip:hover{border-color:var(--ink)}
.ai-q{align-self:flex-start;max-width:88%;background:color-mix(in srgb,var(--marker) 55%,var(--bar));color:var(--marker-ink,#1B1C1F);padding:8px 12px;border-radius:14px 14px 14px 4px;font-size:14px;white-space:pre-wrap}
.ai-a{align-self:stretch;font-size:14.5px;line-height:1.65}
.ai-a p{margin:0 0 6px}
.ai-a .cite{border:0;background:color-mix(in srgb,var(--blue) 12%,transparent);color:var(--blue);font:inherit;font-size:12px;font-weight:600;padding:0 6px;border-radius:6px;cursor:pointer;margin-inline:2px}
.ai-a .cite:hover{background:color-mix(in srgb,var(--blue) 22%,transparent)}
.ai-acts{display:flex;gap:12px;margin-top:2px}
.ai-acts button{border:0;background:none;color:var(--muted);font:inherit;font-size:12.5px;cursor:pointer;padding:2px 0}
.ai-acts button:hover{color:var(--ink);text-decoration:underline}
.ai-wait{color:var(--muted);font-size:13.5px;display:flex;align-items:center;gap:8px}
.ai-wait::before{content:"";width:10px;height:10px;border-radius:50%;background:var(--marker);animation:aiPulse 1s ease-in-out infinite alternate}
@keyframes aiPulse{to{transform:scale(.5);opacity:.5}}
.ai-err{color:var(--danger,#C2362B);font-size:13.5px}
.ai-err .btn{margin-top:8px}
.ai-ask{display:flex;gap:8px;align-items:flex-end;padding:10px 14px;border-top:1px solid var(--line)}
.ai-ask textarea{flex:1;resize:none;min-height:40px;max-height:140px;border:1px solid var(--line);border-radius:12px;padding:9px 12px;font:inherit;font-size:14.5px;background:var(--bar);color:var(--ink);direction:rtl;unicode-bidi:plaintext}
.ai-ask textarea:focus{outline:2px solid var(--marker);border-color:var(--ink)}
.ai-note{margin:0;padding:0 14px 10px;font-size:11.5px;color:var(--muted)}
/* read aloud */
.ra{position:fixed;z-index:980;inset-inline:0;bottom:calc(16px + env(safe-area-inset-bottom,0px));margin:auto;width:min(640px,calc(100vw - 20px));display:flex;align-items:center;gap:6px;padding:8px 10px;background:var(--navy);color:var(--chrome-text,#F2F2EC);border-radius:16px;box-shadow:0 20px 44px -18px rgba(0,0,0,.6)}
.ra.hidden{display:none}
.ra button,.ra select{flex:none;height:34px;min-width:34px;border:0;border-radius:9px;background:rgba(255,255,255,.08);color:inherit;font:inherit;font-size:14px;cursor:pointer}
.ra button:hover{background:rgba(255,255,255,.16)}
.ra .play{background:var(--marker);color:#1B1C1F;min-width:42px;font-size:16px}
.ra select{padding:0 6px}
.ra .now{flex:1;min-width:0;font-size:13px;line-height:1.35;max-height:2.7em;overflow:hidden;opacity:.9;padding-inline:6px;unicode-bidi:plaintext}
.ra .pg{flex:none;font-size:12px;opacity:.65}
@media (max-width:560px){.ra .now{display:none}.ra{justify-content:center}}
/* smart scan */
dialog.scan{border:0;padding:0;width:min(1100px,calc(100vw - 16px));max-width:none;height:min(860px,calc(100vh - 16px));max-height:none;border-radius:18px;background:var(--bar);color:var(--ink);box-shadow:0 30px 80px -24px rgba(0,0,0,.6);overflow:hidden}
dialog.scan[open]{display:flex;flex-direction:column}
dialog.scan::backdrop{background:color-mix(in srgb,var(--navy) 60%,transparent);backdrop-filter:blur(3px)}
.sc-head{display:flex;align-items:center;gap:10px;padding:14px 18px;border-bottom:1px solid var(--line)}
.sc-head h2{margin:0;margin-inline-end:auto;font-size:18px}
.sc-body{flex:1;overflow:auto;padding:22px;min-height:0}
.sc-pick h3{font-family:"SF-secular","SF-rubik",sans-serif;font-weight:400;font-size:clamp(28px,4vw,40px);line-height:1.1;margin:4px 0 8px;max-width:20ch}
.sc-pick p.sub{color:var(--muted);max-width:58ch;margin:0 0 20px;line-height:1.6}
.sc-srcs{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px;max-width:880px}
.sc-src{display:flex;flex-direction:column;align-items:flex-start;gap:8px;text-align:start;padding:16px;border:1px solid var(--line);border-radius:14px;background:var(--bar);color:var(--ink);font:inherit;cursor:pointer}
.sc-src:hover{border-color:var(--ink)}
.sc-src:focus-visible{outline:3px solid var(--marker);outline-offset:2px}
.sc-src b{font-size:15.5px}
.sc-src span{color:var(--muted);font-size:13px;line-height:1.45}
.sc-src svg{width:26px;height:26px}
.sc-opts{display:flex;flex-wrap:wrap;gap:16px;align-items:center;margin:20px 0 0;font-size:14px}
.sc-opts select{font:inherit;padding:6px 8px;border-radius:8px;border:1px solid var(--line);background:var(--bar);color:var(--ink)}
.sc-fine{color:var(--muted);font-size:12.5px;margin-top:18px;max-width:70ch;line-height:1.55}
.sc-prog{list-style:none;margin:0;padding:0;display:grid;grid-template-columns:repeat(auto-fill,minmax(130px,1fr));gap:12px}
.sc-prog li{border:1px solid var(--line);border-radius:12px;padding:6px;font-size:12.5px;color:var(--muted);text-align:center}
.sc-prog img{display:block;width:100%;height:150px;object-fit:contain;background:#fff;border-radius:8px;margin-bottom:6px}
.sc-prog li.ok{color:var(--ink)}
.sc-prog li.bad{color:var(--danger,#C2362B)}
.sc-prog li.now{border-color:var(--ink)}
.sc-rev{display:flex;flex-direction:column;gap:26px}
.sc-pg{display:grid;grid-template-columns:minmax(0,.85fr) minmax(0,1.15fr);gap:18px;align-items:start}
.sc-pg>figure{margin:0;position:sticky;top:0}
.sc-pg img{display:block;width:100%;max-height:72vh;object-fit:contain;background:#fff;border:1px solid var(--line);border-radius:10px;cursor:zoom-in}
.sc-pg figcaption{font-size:12px;color:var(--muted);margin-top:6px}
.sc-again{border:0;background:none;color:var(--blue);font:inherit;font-size:12px;cursor:pointer;text-decoration:underline;padding:0}
.sc-again:disabled{color:var(--muted);text-decoration:none;cursor:default}
.sc-ed{min-height:200px;background:#fff;color:#1B1C1F;border:1px solid var(--line);border-radius:10px;padding:26px 30px;font-family:"SF-rubik",Arial,sans-serif;font-size:15px;line-height:1.6;box-shadow:var(--paper-shadow)}
.sc-ed:focus{outline:2px solid var(--marker)}
.sc-ed h1{font-size:23px;margin:0 0 10px}
.sc-ed h2{font-size:18px;margin:14px 0 6px}
.sc-ed p{margin:0 0 9px}
.sc-ed ul,.sc-ed ol{margin:0 0 9px;padding-inline-start:24px}
.sc-ed table{border-collapse:collapse;width:100%;margin:0 0 10px}
.sc-ed td{border:1px solid #9a9a92;padding:4px 7px;vertical-align:top}
.sc-ed mark{background:#FFE58A;color:inherit;border-radius:3px;padding:0 2px;box-shadow:inset 0 -2px 0 #E0A100}
.sc-ed .empty{color:#8a8a82}
.sc-foot{display:flex;flex-wrap:wrap;gap:8px;align-items:center;padding:12px 18px;border-top:1px solid var(--line)}
.sc-foot label{display:flex;align-items:center;gap:8px;font-size:13.5px;color:var(--muted);margin-inline-end:auto}
.sc-foot input{font:inherit;font-size:14px;padding:6px 9px;border:1px solid var(--line);border-radius:8px;width:min(260px,52vw);background:var(--bar);color:var(--ink)}
.sc-foot .hint{font-size:13px;color:var(--muted)}
.sc-foot .hint button{border:0;background:none;color:var(--blue);font:inherit;cursor:pointer;text-decoration:underline}
.sc-zoom{position:fixed;inset:0;z-index:5;background:rgba(0,0,0,.85);display:grid;place-items:center;cursor:zoom-out}
.sc-zoom img{max-width:96vw;max-height:96vh}
@media (max-width:760px){.sc-pg{grid-template-columns:1fr}.sc-pg>figure{position:static}.sc-pg img{max-height:38vh}.sc-ed{padding:16px}.sc-body{padding:14px}}
.ocr-lang{font:inherit;font-size:13px;padding:4px 6px;border-radius:8px;border:1px solid var(--line);background:var(--bar);color:var(--ink)}
.pg-tools .ocr-lang{margin:2px 0 6px;width:100%}
@media (prefers-reduced-motion:reduce){.ai-wait::before{animation:none}}
</style>`));

/* =====================================================================
   1. TABS — several files open together
   ===================================================================== */
const KEEP = ['bytes', 'name', 'doc', 'pages', 'zoom', 'anns', 'hist', 'redo', 'dirty', 'nextId', 'handle', 'office', 'hasForm', 'sel'];
const MAX_TABS = 12;
const TB = { list: [], cur: null, prev: null, seq: 0, depth: 0, drag: null };
const tabData = t => t === TB.cur && !t.ed ? S : (t.data || {});
// Word / Excel files open in the full editors are tabs too ("editor tabs", t.ed = 'docx' | 'xlsx').
// The editors are one-at-a-time modules, so a tab that is not in front keeps its file as bytes ("parked")
// and is opened again when it comes back.
// the three editors: Word, Excel and PowerPoint
const ED = {
  docx: { api: 'WORD_EDITOR', st: 'E', id: 'wed', name: '#weName', kind: 'WORD', build: 'buildDocx', lib: 'docx', file: 'word-editor.js' },
  xlsx: { api: 'XLSX_EDITOR', st: 'X', id: 'xed', name: '#xeName', kind: 'EXCEL', build: 'buildXlsx', lib: 'xlsx', file: 'excel-editor.js' },
  pptx: { api: 'PPTX_EDITOR', st: 'X', id: 'ped', name: '#peName', kind: 'POWERPOINT', build: 'buildPptx', lib: 'pptx', file: 'ppt-editor.js' },
};
const ED_IDS = Object.values(ED).map(e => e.id), ED_SEL = ED_IDS.map(i => '#' + i).join(',');
const edKindOfId = id => Object.keys(ED).find(k => ED[k].id === id);
const edApi = k => window[ED[k].api];
const edState = k => { const A = edApi(k); return A && A._test ? A._test[ED[k].st] : null; };
const edLive = t => !!(t && t.ed && t.liveEl && document.body.contains(t.liveEl));
function tabInfo(t) {
  if (t.ed) {
    const live = edLive(t), st = live && edState(t.ed);
    const inp = live && t.liveEl.querySelector(ED[t.ed].name);
    const base = ((inp && inp.value) || (t.data && t.data.name) || 'מסמך').replace(/\.(docx|xlsx|pptx)$/i, '');
    return { name: base + '.' + t.ed, dirty: live ? !!(st && (st.dirty || st.snapT)) : !!(t.data && t.data.dirty), kind: ED[t.ed].kind, edit: true };
  }
  const d = tabData(t);
  return { name: (d.office ? d.office.name : d.name) || 'מסמך', dirty: !!d.dirty, kind: d.office ? d.office.kind.toUpperCase() : 'PDF' };
}
const tabName = t => tabInfo(t).name;
const deskTabs = () => TB.list.filter(t => !t.ed);
async function park(t) {
  if (!edLive(t)) return;
  const A = edApi(t.ed), st = edState(t.ed); if (!A || !st) return;
  busy(true, 'שומר את הקובץ בצד…');
  try {
    const r = await A._test[ED[t.ed].build]();
    const info = tabInfo(t);
    t.data = { bytes: r && r.bytes ? r.bytes : r, name: info.name, dirty: info.dirty, opts: st.opts };
    t.parking = true;
    await A.close();
  } finally { t.parking = false; t.liveEl = null; busy(false); }
}
async function unpark(t) {
  const d = t.data; if (!d) return;
  TB.opening = t; TB.cur = t;
  try {
    await loadEditLibs(ED[t.ed].lib); if (!edApi(t.ed)) await loadScript(ED[t.ed].file);
    await edApi(t.ed).open({ ...d.opts, bytes: d.bytes, name: d.name, isNew: false, restored: d.dirty });
  } catch (e) { console.error(e); toast('לא הצלחתי לפתוח שוב את הקובץ: ' + (e.message || e)); }
  finally { if (TB.opening === t) TB.opening = null; }
}
// leave the tab in front / bring a tab to the front (desk tabs and editor tabs alike)
async function leave() {
  const c = TB.cur; if (!c) return;
  if (c.ed) { if (edLive(c)) await park(c); } else stash();
  TB.cur = null;
}
async function enter(t) { if (t.ed) await unpark(t); else restore(t); renderTabs(); AS.onTab(); }

function observe() {
  if (io) io.disconnect();
  io = new IntersectionObserver(es => es.forEach(e => { if (e.isIntersecting) renderPage(S.pages[+e.target.dataset.i]); }), { root: $('desk'), rootMargin: '800px 0px' });
  if (far) far.disconnect();
  far = new IntersectionObserver(es => es.forEach(e => { if (!e.isIntersecting && S.pages.length > 6) releasePage(S.pages[+e.target.dataset.i]); }), { root: $('desk'), rootMargin: '3000px 0px' });
}
function stash() {
  const t = TB.cur; if (!t || t.ed || !S.doc) return;
  RA.stop();
  const d = {};
  for (const k of KEEP) d[k] = S[k];
  d.top = $('desk').scrollTop; d.left = $('desk').scrollLeft;
  if (io) io.disconnect(); if (far) far.disconnect();
  for (const P of S.pages) releasePage(P);          // the page pictures are drawn again when the tab comes back
  const fr = document.createDocumentFragment();
  while ($('pages').firstChild) fr.append($('pages').firstChild);
  d.frag = fr;
  d.tr = { pages: TR.pages, src: TR.src, langFromApi: TR.langFromApi, target: TR.target, lang: $('tpLang').textContent,
    foot: !$('tpFoot').classList.contains('hidden'), fmt: !$('tpFmt').classList.contains('hidden'), status: $('tpStatus').textContent };
  const rf = document.createDocumentFragment();
  while ($('tpResults').firstChild) rf.append($('tpResults').firstChild);
  d.trFrag = rf;
  t.data = d;
  pickAnn(null);
  Object.assign(S, { bytes: null, doc: null, pages: [], anns: [], hist: [], redo: [], dirty: false, handle: null, office: null, sel: new Set(), hasForm: false, name: '' });
  TR.stop = true; TR.pages = [];
}
function applyUI() {
  setDirty(S.dirty);
  $('fname').textContent = dispName(S.office ? S.office.name : S.name);
  $('empty').classList.add('hidden');
  ['saveBtn', 'saveAsBtn', 'saveMenuBtn', 'trBtn', 'pgBtn', 'printBtn', 'convBtn', 'shareBtn', 'driveMenuBtn', 'icloudMenuBtn'].forEach(id => { const b = $(id); if (b) b.disabled = false; });
  $('convBtn').disabled = !!S.office;
  const ed = S.office && (S.office.kind === 'docx' || S.office.kind === 'xlsx');
  $('wordEditBtn').classList.toggle('hidden', !ed);
  if (ed) { $('wordEditBtn').textContent = S.office.kind === 'xlsx' ? 'עריכת Excel' : 'עריכת Word'; }
  setTool(S.tool);
}
function restore(t) {
  const d = t.data; t.data = null; TB.cur = t;
  showOfficeView(false);
  for (const k of KEEP) S[k] = d[k];
  $('pages').append(d.frag);
  observe();
  layout();
  $('desk').scrollTop = d.top; $('desk').scrollLeft = d.left;
  TR.pages = d.tr.pages || []; TR.src = d.tr.src; TR.langFromApi = d.tr.langFromApi; TR.stop = true;
  $('tpResults').append(d.trFrag);
  $('tpLang').textContent = d.tr.lang; $('tpFoot').classList.toggle('hidden', !d.tr.foot); $('tpFmt').classList.toggle('hidden', !d.tr.fmt);
  $('tpStatus').textContent = d.tr.status; $('tpStatus').classList.remove('err');
  setTarget(d.tr.target || 'he');
  applyUI();
  updatePageInfo();
  if (!$('pgpanel').classList.contains('hidden')) { renderThumbs(); updateSel(); }
  AS.onTab();
}
function toEmpty() {
  if (io) io.disconnect(); if (far) far.disconnect();
  $('pages').innerHTML = '';
  Object.assign(S, { bytes: null, doc: null, pages: [], anns: [], hist: [], redo: [], dirty: false, handle: null, office: null, sel: new Set(), hasForm: false, name: '' });
  pickAnn(null); TR.reset();
  setDirty(false); document.title = 'סימן';
  $('fname').textContent = '';
  $('empty').classList.remove('hidden');
  ['saveBtn', 'saveAsBtn', 'saveMenuBtn', 'trBtn', 'pgBtn', 'printBtn', 'convBtn', 'shareBtn', 'driveMenuBtn', 'icloudMenuBtn'].forEach(id => { const b = $(id); if (b) b.disabled = true; });
  $('wordEditBtn').classList.add('hidden');
  $('pgpanel').classList.add('hidden'); $('tpanel').classList.add('hidden');
  updatePageInfo(); AS.onTab();
}

// every way of opening a file comes through here: a file that opens while another is open gets its own tab
const openOne = openFile;
openFile = async function (file, handle) {
  if (!file) return;
  if (TB.depth === 0 && TB.cur && TB.cur.ed && !(typeof needReg === 'function' && needReg())) {
    if (TB.list.length >= MAX_TABS) { toast(`אפשר לפתוח עד ${MAX_TABS} קבצים יחד. סגרו אחד מהם כדי לפתוח עוד`); return; }
    if (edLive(TB.cur)) { await park(TB.cur); TB.prev = TB.cur; } else TB.prev = null;
    TB.cur = null;
  } else if (TB.depth === 0 && S.doc && TB.cur && !(typeof needReg === 'function' && needReg())) {
    if (TB.list.length >= MAX_TABS) { toast(`אפשר לפתוח עד ${MAX_TABS} קבצים יחד. סגרו אחד מהם כדי לפתוח עוד`); return; }
    if (isBusy()) return;
    TB.prev = TB.cur; stash(); TB.cur = null;
  }
  TB.depth++;
  try { await openOne(file, handle); }
  finally { TB.depth--; }
  if (TB.depth === 0) settle();
};
function settle() {
  if (S.doc && !TB.cur) {
    const t = { id: ++TB.seq, data: null };
    const at = TB.prev ? TB.list.indexOf(TB.prev) + 1 : TB.list.length;
    TB.list.splice(at || TB.list.length, 0, t);
    TB.cur = t; TB.prev = null; AS.onTab();
  } else if (!S.doc && !TB.cur && TB.prev && TB.list.includes(TB.prev)) {      // nothing came up (cancelled or failed)
    const p = TB.prev; TB.prev = null; enter(p);
  } else TB.prev = null;
  renderTabs();
}
let switching = false;
async function switchTo(t) {
  if (t === TB.cur || switching) return;
  if (isBusy() || TR.running) { toast('רגע, פעולה עדיין רצה'); return; }
  switching = true;
  try { await leave(); await enter(t); } finally { switching = false; renderTabs(); }
}
async function closeTab(t) {
  if (isBusy() || switching) return;
  const info = tabInfo(t);
  if (info.dirty && !confirm(`ב"${info.name}" יש שינויים שלא נשמרו. לסגור בכל זאת?`)) return;
  const i = TB.list.indexOf(t); TB.list.splice(i, 1);
  AS.forget(t);
  if (t.ed) { if (edLive(t)) { t.closing = true; try { await edApi(t.ed).close(); } catch (_) {} } }
  else if (t !== TB.cur) { try { t.data.doc && t.data.doc.destroy(); } catch (_) {} }
  if (t === TB.cur) {
    if (!t.ed) {
      RA.stop();
      try { S.doc && S.doc.destroy(); } catch (_) {}
      if (io) io.disconnect(); if (far) far.disconnect();
      $('pages').innerHTML = '';
      Object.assign(S, { doc: null, dirty: false });
    }
    TB.cur = null;
    const next = TB.list[Math.min(i, TB.list.length - 1)];
    if (next) await enter(next); else toEmpty();
  }
  renderTabs();
}
async function closeAll() {
  const dirty = TB.list.filter(t => tabInfo(t).dirty).length;
  if (dirty && !confirm(dirty === 1 ? 'בקובץ אחד יש שינויים שלא נשמרו. לסגור הכל בכל זאת?' : `ב-${dirty} קבצים יש שינויים שלא נשמרו. לסגור הכל בכל זאת?`)) return;
  const list = TB.list; TB.list = [];
  for (const t of list) {
    if (t.ed) { if (edLive(t)) { t.closing = true; try { await edApi(t.ed).close(); } catch (_) {} } }
    else if (t !== TB.cur) { try { t.data.doc && t.data.doc.destroy(); } catch (_) {} }
  }
  RA.stop();
  if (TB.cur && !TB.cur.ed) { try { S.doc && S.doc.destroy(); } catch (_) {} }
  TB.cur = null; AS.forgetAll();
  toEmpty(); renderTabs();
}

/* ---- the full Word / Excel editors as tabs ---- */
function edButton(el) {
  if (el.querySelector('.ed-ai')) return;
  const close = el.querySelector('#weClose,#xeClose,#peClose'); if (!close) return;
  const peer = [...close.parentElement.querySelectorAll('button')].find(b => b !== close && !b.classList.contains('primary'));
  const b = document.createElement('button');
  b.type = 'button'; b.className = (peer ? peer.className + ' ' : '') + 'ed-ai';
  b.title = 'העוזר: שאלות על הקובץ, סיכום והקראה';
  b.innerHTML = ICON.spark + '<span>עוזר</span>';
  b.onclick = () => toggleAI();
  close.after(b);
}
function edAdded(kind, el) {
  setTimeout(() => edButton(el), 0);
  if (TB.opening && TB.opening.ed === kind) { const t = TB.opening; TB.opening = null; t.liveEl = el; t.data = null; TB.cur = t; renderTabs(); return; }
  const same = TB.list.find(x => x.ed === kind && x.liveEl && !document.body.contains(x.liveEl) && x === TB.cur);
  if (same) { same.liveEl = el; renderTabs(); return; }   // the editor reopened itself (restoring a draft)
  if (TB.cur && !TB.cur.ed) { TB.prev = TB.cur; stash(); }
  const t = { id: ++TB.seq, ed: kind, liveEl: el, data: null };
  setTimeout(() => AS.onTab(), 0);
  const at = TB.prev && TB.list.includes(TB.prev) ? TB.list.indexOf(TB.prev) + 1 : TB.cur ? TB.list.indexOf(TB.cur) + 1 : TB.list.length;
  TB.list.splice(at || TB.list.length, 0, t);
  TB.cur = t; TB.prev = null;
  if (TB.list.length > MAX_TABS) toast(`יש ${TB.list.length} קבצים פתוחים. כדאי לסגור כמה`);
  renderTabs();
}
function edRemoved(kind, el) {
  const t = TB.list.find(x => x.liveEl === el);
  if (!t || t.parking || t.closing) { renderTabs(); return; }
  setTimeout(async () => {
    const id = ED[kind].id, now = document.getElementById(id);
    if (now && TB.cur === t && !TB.list.some(x => x !== t && x.liveEl === now)) { t.liveEl = now; renderTabs(); return; }
    if (edLive(t) || t.data || !TB.list.includes(t)) { renderTabs(); return; }
    // closed in the editor itself (×), or sent to the desk for signing
    const i = TB.list.indexOf(t); TB.list.splice(i, 1); AS.forget(t);
    if (TB.cur === t) {
      TB.cur = null;
      if (!S.doc && TB.depth === 0) { const n = TB.list[Math.min(i, TB.list.length - 1)]; if (n) await enter(n); else toEmpty(); }
    }
    renderTabs();
  }, 450);
}
new MutationObserver(ms => {
  for (const m of ms) {
    for (const n of m.addedNodes) if (n.nodeType === 1 && ED_IDS.includes(n.id)) edAdded(edKindOfId(n.id), n);
    for (const n of m.removedNodes) if (n.nodeType === 1 && ED_IDS.includes(n.id)) edRemoved(edKindOfId(n.id), n);
  }
}).observe(document.body, { childList: true });
// a new editor while another one is in front: the one in front is put aside first, not closed
for (const fn of ['openWordEditor', 'openExcelEditor', 'openPptEditor']) {
  const one = window[fn];
  if (typeof one !== 'function') continue;
  window[fn] = async function (...args) {
    if (TB.cur && TB.cur.ed && edLive(TB.cur)) { await park(TB.cur); TB.prev = TB.cur; TB.cur = null; }
    return one.apply(this, args);
  };
}
// the unsaved-changes dot of an editor tab follows the typing
let edT = 0;
document.addEventListener('input', e => { if (e.target.closest && e.target.closest(ED_SEL)) { clearTimeout(edT); edT = setTimeout(renderTabs, 600); } }, true);

const strip = h('<div class="tabs hidden" id="tabs" role="tablist" aria-label="הקבצים הפתוחים"></div>');
document.querySelector('.work').before(strip);
let tabsRaf = 0;
function renderTabs() {
  cancelAnimationFrame(tabsRaf);
  tabsRaf = requestAnimationFrame(() => {
    strip.classList.toggle('hidden', !TB.list.length);
    document.body.classList.toggle('ed-live', TB.list.some(edLive));
    strip.innerHTML = '';
    for (const t of TB.list) {
      const on = t === TB.cur, d = tabInfo(t);
      const b = h(`<div class="tab${d.edit ? ' ed ' + t.ed : ''}" role="tab" tabindex="${on ? 0 : -1}" aria-selected="${on}" draggable="true"><span class="kd">${d.kind}</span><span class="nm"></span>${d.dirty ? '<span class="dot" title="יש שינויים שלא נשמרו"></span>' : ''}<span class="x" role="button" aria-label="סגירת הקובץ" title="סגירה">×</span></div>`);
      b.querySelector('.nm').textContent = dispName(d.name);
      b.title = d.name + (d.edit ? ' (עריכה)' : '');
      b.onclick = e => { if (e.target.closest('.x')) { closeTab(t); return; } switchTo(t); };
      b.onauxclick = e => { if (e.button === 1) { e.preventDefault(); closeTab(t); } };
      b.onkeydown = e => {
        const i = TB.list.indexOf(t);
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); switchTo(t); }
        else if (e.key === 'ArrowLeft' || e.key === 'ArrowRight') { e.preventDefault(); const n = TB.list[(i + (e.key === 'ArrowLeft' ? 1 : -1) + TB.list.length) % TB.list.length]; switchTo(n); setTimeout(() => strip.querySelector('[aria-selected="true"]')?.focus(), 30); }
        else if (e.key === 'Delete') { e.preventDefault(); closeTab(t); }
      };
      b.ondragstart = e => { TB.drag = t; e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', 'siman-tab'); };
      b.ondragend = () => { TB.drag = null; strip.querySelectorAll('.drop-before').forEach(x => x.classList.remove('drop-before')); };
      b.ondragover = e => { if (!TB.drag) return; e.preventDefault(); strip.querySelectorAll('.drop-before').forEach(x => x.classList.remove('drop-before')); b.classList.add('drop-before'); };
      b.ondrop = e => { if (!TB.drag) return; e.preventDefault(); const from = TB.list.indexOf(TB.drag); TB.list.splice(from, 1); TB.list.splice(TB.list.indexOf(t), 0, TB.drag); TB.drag = null; renderTabs(); };
      strip.append(b);
    }
    const add = h(`<button class="tab-add" type="button" title="פתיחת עוד קבצים (אפשר לבחור כמה יחד)">${ICON.plus}<span>עוד קובץ</span></button>`);
    add.onclick = () => pickFile();
    strip.append(add);
    if (TB.list.length > 1) {
      const all = h(`<button class="tab-all" type="button" aria-haspopup="menu">כל הקבצים (${TB.list.length}) ▾</button>`);
      all.onclick = e => { e.stopPropagation(); allMenu(all); };
      strip.append(all);
    }
    const cur = strip.querySelector('[aria-selected="true"]'); if (cur) cur.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  });
}
// the name or the unsaved-changes dot of the open tab can change anywhere
const setDirtyOne = setDirty;
setDirty = function (v) { setDirtyOne(v); if (TB.list.length) renderTabs(); };
// tabs dragged inside the strip are not files
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(ev => addEventListener(ev, e => { if (TB.drag) e.stopPropagation(); }, true));

function allMenu(anchor) {
  document.querySelector('.tab-menu')?.remove();
  const m = h(`<div class="tab-menu" role="menu">
    <button role="menuitem" data-a="merge">איחוד לקובץ PDF אחד<small>לפי סדר הלשוניות, עם כל הסימונים והחתימות${TB.list.some(t => t.ed) ? '. בלי הקבצים שבעורך' : ''}</small></button>
    <button role="menuitem" data-a="zip">שמירת כל הקבצים<small>כל קובץ כ-PDF, יחד בקובץ ZIP אחד${TB.list.some(t => t.ed) ? '. בלי הקבצים שבעורך' : ''}</small></button>
    <button role="menuitem" data-a="close">סגירת כל הקבצים</button></div>`);
  document.body.append(m);
  const r = anchor.getBoundingClientRect();
  m.style.top = (r.bottom + 6) + 'px';
  m.style.left = Math.max(8, Math.min(innerWidth - m.offsetWidth - 8, r.left + r.width - m.offsetWidth)) + 'px';
  const close = () => { m.remove(); removeEventListener('pointerdown', out, true); };
  const out = e => { if (!m.contains(e.target)) close(); };
  setTimeout(() => addEventListener('pointerdown', out, true), 0);
  m.onkeydown = e => { if (e.key === 'Escape') { close(); anchor.focus(); } };
  m.querySelectorAll('button').forEach(b => b.onclick = () => { close(); ({ merge: mergeAll, zip: zipAll, close: closeAll })[b.dataset.a](); });
  m.querySelector('button').focus();
}
// a PDF of every tab, with its marks: each tab is brought up in turn and built
async function eachTabPdf(label) {
  const back = TB.cur, out = [];
  if (back && back.ed) await leave();
  for (const t of deskTabs()) {
    if (t !== TB.cur) { stash(); restore(t); }
    busy(true, `${label}: ${tabName(t)}…`);
    const bytes = await buildPdf();
    const base = (S.office ? S.office.base || S.office.name : S.name).replace(/\.(pdf|docx|xlsx|pptx|doc)$/i, '');
    out.push({ name: base + '.pdf', bytes: bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes) });
  }
  if (TB.cur !== back && back) { if (back.ed) { stash(); TB.cur = null; } else stash(); await enter(back); }
  renderTabs();
  return out;
}
async function mergeAll() {
  if (deskTabs().length < 2) { toast('האיחוד עובד על קבצים שפתוחים לסימון (לא על קבצים בעורך Word / Excel)'); return; }
  try {
    const parts = await eachTabPdf('מאחד');
    busy(true, 'מחבר את הקבצים…');
    const out = await PDFLib.PDFDocument.create();
    for (const p of parts) {
      const d = await PDFLib.PDFDocument.load(p.bytes, { ignoreEncryption: true });
      for (const pg of await out.copyPages(d, d.getPageIndices())) out.addPage(pg);
    }
    const bytes = await out.save();
    busy(false);
    await openFile(new File([bytes], 'קבצים מאוחדים.pdf', { type: 'application/pdf' }));
    toast(`${parts.length} קבצים אוחדו לקובץ אחד (${S.pages.length} עמודים). הקבצים המקוריים לא השתנו`);
  } catch (e) { console.error(e); toast('האיחוד נכשל: ' + (e.message || e)); }
  finally { busy(false); }
}
async function zipAll() {
  if (!deskTabs().length) return;
  try {
    const parts = await eachTabPdf('מכין');
    const used = new Set(), enc = [];
    for (const p of parts) { let n = p.name, k = 2; while (used.has(n)) n = p.name.replace(/\.pdf$/i, '') + ` (${k++}).pdf`; used.add(n); enc.push({ name: n, data: p.bytes }); }
    const z = zip(enc);
    busy(false);
    if (await saveBytes('קבצים מסימן.zip', z, 'zip')) toast(`${parts.length} קבצים נשמרו בקובץ ZIP`);
  } catch (e) { console.error(e); toast('השמירה נכשלה: ' + (e.message || e)); }
  finally { busy(false); }
}

// choosing, dropping or launching several files at once
async function openMany(items) {
  if (!items.length) return;
  if (typeof needReg === 'function' && needReg()) { const it = items[0]; openFile(it.f || await it.h.getFile(), it.h || null); return; }
  const room = MAX_TABS - TB.list.length;
  if (items.length > room) { toast(`נפתחים ${room} הקבצים הראשונים (עד ${MAX_TABS} יחד)`); items = items.slice(0, room); }
  for (const it of items) { try { await openFile(it.f || await it.h.getFile(), it.h || null); } catch (e) { console.error(e); } }
}
pickFile = async function () {
  if (window.showOpenFilePicker) {
    try { const hs = await showOpenFilePicker({ types: OPEN_TYPES, multiple: true }); await openMany(hs.map(x => ({ h: x }))); }
    catch (e) { if (e && e.name !== 'AbortError') $('file').click(); }
  } else $('file').click();
};
$('file').multiple = true;
$('file').onchange = e => { const fs = [...e.target.files]; e.target.value = ''; openMany(fs.map(f => ({ f }))); };
addEventListener('drop', async e => {
  if (TB.drag) return;
  const fs = e.dataTransfer && e.dataTransfer.files;
  const onEditor = ED_IDS.some(i => document.getElementById(i));
  const docs = fs && fs.length && [...fs].every(f => /\.(pdf|docx?|xlsx|pptx)$/i.test(f.name));
  if (!fs || !fs.length || (fs.length < 2 && !(onEditor && docs))) return;   // one file on the desk: the usual way
  if (onEditor && !docs) return;                                            // pictures dropped into the editor stay there
  e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
  dragN = 0; document.body.classList.remove('dragging');
  const files = [...fs], items = [...(e.dataTransfer.items || [])];
  // the handles must be asked for now: the dropped items are gone once the event is over
  const hs = files.map((_, k) => { try { const it = items[k]; return it && it.getAsFileSystemHandle ? it.getAsFileSystemHandle().catch(() => null) : null; } catch (_) { return null; } });
  const got = await Promise.all(hs.map(p => p || Promise.resolve(null)));
  openMany(files.map((f, k) => ({ f, h: got[k] && got[k].kind === 'file' ? got[k] : null })));
}, true);
window.SIMAN_TABS = { TB, switchTo, closeTab, mergeAll, zipAll, openMany, tabInfo };

/* =====================================================================
   2. AI — the smart server (siman-ai)
   ===================================================================== */
async function ai(payload, { onWait, signal, timeout = 300000 } = {}) {
  for (let attempt = 0; attempt < 4; attempt++) {
    let r, d, status;
    const ctl = new AbortController(), timer = setTimeout(() => ctl.abort(), timeout);
    if (signal) signal.addEventListener('abort', () => ctl.abort(), { once: true });
    try {
      r = await fetch(AI, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(payload), signal: ctl.signal });
      // the server keeps the line open with spaces while Gemini works; the JSON comes at the end
      const raw = await r.text();
      try { d = JSON.parse(raw.trim() || '{}'); } catch (_) { d = {}; }
    } catch (e) {
      clearTimeout(timer);
      if (signal && signal.aborted) throw e;
      if (e && e.name === 'AbortError') throw new Error('הבקשה לקחה יותר מדי זמן. נסו שוב, או צלמו את הדף קרוב יותר ובאור טוב');
      throw new Error(navigator.onLine === false ? 'אין חיבור לאינטרנט' : 'השירות החכם לא זמין כרגע. נסו שוב בעוד רגע');
    }
    clearTimeout(timer);
    status = r.status;
    if (r.ok && d.error) { status = +d.error.status || 502; r = { ok: false, status }; }
    if (r.ok) {
      const c = (d.candidates || [])[0];
      const out = c && c.content && (c.content.parts || []).filter(x => x.text && !x.thought).map(x => x.text).join('');
      if (out) return out;
      if (c && /SAFETY|PROHIBITED|RECITATION/.test(c.finishReason || '')) throw new Error('השירות לא הסכים לעבד את התוכן הזה');
      throw new Error('התקבלה תשובה ריקה. נסו שוב');
    }
    const all = JSON.stringify(d), msg = (d.error && d.error.message) || ('HTTP ' + r.status);
    if (r.status === 429 && /per ?day|PerDay/i.test(all)) throw new Error('נגמרה המכסה היומית החינמית. היא מתאפסת מחר בבוקר');
    if (r.status === 403) throw new Error('השרת החכם חוסם את הכתובת הזו של האפליקציה');
    if (r.status === 404) throw new Error('השרת החכם (siman-ai) עוד לא הוקם');
    if (r.status === 504 || r.status === 524) throw new Error('הקריאה לקחה יותר מדי זמן. נסו שוב בעוד רגע');
    if ([429, 500, 502, 503].includes(r.status)) {
      const wait = 8 * (attempt + 1);
      for (let k = wait; k > 0; k--) { onWait && onWait(k); await sleep(1000); if (signal && signal.aborted) throw new DOMException('stop', 'AbortError'); }
      continue;
    }
    throw new Error('שגיאה מהשרת: ' + msg);
  }
  throw new Error('השירות עמוס כרגע. נסו שוב בעוד דקה');
}

/* =====================================================================
   3. THE ASSISTANT — questions, summaries, key points
   ===================================================================== */
const textCache = new WeakMap();
const onEditor = () => !!(TB.cur && TB.cur.ed && edLive(TB.cur));
const hasDoc = () => !!S.doc || onEditor();
const docTitle = () => onEditor() ? tabInfo(TB.cur).name : (S.office ? S.office.name : S.name);
// the text of a file open in the Word / Excel editor
function editorText() {
  const k = TB.cur.ed, st = edState(k); if (!st) return '';
  if (k === 'docx') {
    const t = (st.main || st.body || TB.cur.liveEl.querySelector('#weBody')).innerText || '';
    return `===== עמוד 1 =====\n${t.replace(/\n{3,}/g, '\n\n').slice(0, 380000)}`;
  }
  const out = []; let total = 0;
  for (const sh of (st.st && st.st.sheets) || []) {
    if (sh.hidden) continue;
    const rows = new Map();
    for (const [kk, cell] of sh.cells) {
      const r = Math.floor(kk / 16384), c = kk % 16384;
      let v = cell && cell.v != null && cell.v !== '' ? cell.v : cell && cell.f ? '=' + cell.f : '';
      if (v === '' || v == null) continue;
      if (!rows.has(r)) rows.set(r, []);
      rows.get(r)[c] = String(v);
    }
    const lines = [...rows.keys()].sort((a, b) => a - b).map(r => `שורה ${r + 1}: ` + Array.from(rows.get(r), x => x ?? '').join('\t'));
    const part = `===== גיליון: ${sh.name || ''} =====\n${lines.join('\n')}`;
    out.push(part); total += part.length; if (total > 380000) break;
  }
  return out.join('\n\n');
}
async function docText() {
  if (onEditor() && TB.cur.ed === 'pptx') return (await PPTX_EDITOR.textOf()).slice(0, 380000);
  if (onEditor()) return editorText();
  if (!S.doc) return '';
  const sig = S.pages.map(P => P.ocr ? 1 : 0).join('') + '|' + S.pages.length;
  const c = textCache.get(S.doc);
  if (c && c.sig === sig) return c.text;
  const parts = [];
  let total = 0;
  for (let i = 0; i < S.pages.length && total < 380000; i++) {
    let t = '';
    try { t = await pageText(i); } catch (_) {}
    if (S.office && S.office.kind === 'docx' && i > 0 && !t) break;
    const part = `===== עמוד ${i + 1} =====\n${t || ''}`;
    parts.push(part); total += part.length;
  }
  const text = parts.join('\n\n');
  textCache.set(S.doc, { sig, text });
  return text;
}
const hasRealText = t => t.replace(/=====.*?=====/g, '').replace(/\s/g, '').length > 30;

const panel = h(`<aside class="tpanel aip hidden" id="aiPanel" aria-label="העוזר">
  <div class="tp-head"><h2>העוזר של סימן</h2><button class="icon-btn" id="aiClose" title="סגירה" aria-label="סגירת העוזר">×</button></div>
  <div class="ai-tools">
    <button class="btn" id="aiRead" type="button">${ICON.speak}<span>הקראה בקול</span></button>
    <button class="btn" id="aiScan" type="button">${ICON.scan}<span>כתב יד ל-Word / PDF</span></button>
  </div>
  <div class="ai-log" id="aiLog" aria-live="polite"></div>
  <div class="ai-ask"><textarea id="aiQ" rows="1" placeholder="שאלו כל דבר על המסמך…" aria-label="שאלה על המסמך"></textarea><button class="btn primary" id="aiSend" type="button">שליחה</button></div>
  <p class="ai-note">התשובות נכתבות לפי הטקסט שבמסמך, בעזרת Google Gemini. כדאי לבדוק פרטים חשובים מול המקור.</p>
</aside>`);
$('tpanel').after(panel);
const CHIPS = [
  ['סיכום קצר', 'סכם את המסמך בקצרה, ב-3 עד 5 משפטים.'],
  ['הנקודות העיקריות', 'מה הנקודות העיקריות במסמך? כתוב אותן כרשימה קצרה.'],
  ['תאריכים, סכומים ושמות', 'רכז את כל התאריכים, הסכומים, השמות ומספרי המסמכים שמופיעים במסמך, כרשימה.'],
  ['הסבר בשפה פשוטה', 'הסבר את המסמך בשפה פשוטה, כמו לאדם שאינו מכיר את התחום.'],
  ['מה צריך לעשות?', 'מה המסמך מבקש ממני לעשות, למלא או לחתום, ועד מתי?'],
];
const AS = {
  hist: new Map(),           // tab id -> [{q,a}] or {err}
  key() { return TB.cur ? TB.cur.id : 0; },
  list() { const k = this.key(); if (!this.hist.has(k)) this.hist.set(k, []); return this.hist.get(k); },
  forget(t) { this.hist.delete(t.id); },
  forgetAll() { this.hist.clear(); },
  onTab() { if (!panel.classList.contains('hidden')) renderLog(); },
  busy: false,
};
const hideAI = () => { panel.classList.add('hidden'); setTimeout(fitWidth, 0); };
function showAI() {
  $('tpanel').classList.add('hidden');
  panel.classList.remove('hidden');
  renderLog(); setTimeout(fitWidth, 0);
  setTimeout(() => { if (matchMedia('(pointer:fine)').matches) $('aiQ').focus(); }, 60);
}
const toggleAI = () => panel.classList.contains('hidden') ? showAI() : hideAI();
$('aiClose').onclick = hideAI;
// the editors listen to the keyboard and to paste on the whole window: inside the assistant, the keys are the assistant's
for (const ev of ['keydown', 'keyup', 'keypress', 'paste', 'copy', 'cut']) addEventListener(ev, e => { if (e.target && e.target.closest && e.target.closest('#aiPanel')) { e.stopImmediatePropagation();
    // stopping here also stops the textarea's own key handler, so Enter (send) and Esc are handled here
    if (ev === 'keydown' && e.key === 'Escape') hideAI();
    if (ev === 'keydown' && e.key === 'Enter' && !e.shiftKey && !e.isComposing && e.target.id === 'aiQ') { e.preventDefault(); $('aiSend').click(); }
  } }, true);
$('trBtn').addEventListener('click', () => { if (!$('tpanel').classList.contains('hidden')) panel.classList.add('hidden'); });

function citeHtml(text) {
  // [עמ' 3] / [עמ׳ 3-4] / [p. 3] / [pp. 3, 5] → buttons that jump to the page
  return esc(text).replace(/\[(?:עמ['׳]?|עמוד|עמודים|p\.|pp\.|page|стр\.|ص)\s*([\d\s,–-]+)\]/gi, (m, nums) => {
    const first = parseInt(nums, 10);
    return Number.isFinite(first) ? `<button class="cite" type="button" data-p="${first}" title="מעבר לעמוד ${first}">${esc(m.slice(1, -1))}</button>` : m;
  });
}
function answerEl(a, q, plain) {
  const el = h('<div class="ai-a"></div>');
  // in the Word / Excel editor there are no page numbers to jump to: the page marks are left out
  const strip = l => l.replace(/\s*\[(?:עמ['׳]?|עמוד|עמודים|p\.|pp\.|page)\s*[\d\s,–-]+\]/gi, '');
  el.innerHTML = a.split(/\n+/).filter(l => l.trim()).map(l => `<p dir="auto">${plain ? esc(strip(l)) : citeHtml(l)}</p>`).join('');
  const acts = h('<div class="ai-acts"><button type="button" data-a="copy">העתקה</button><button type="button" data-a="read">הקראה</button></div>');
  acts.querySelector('[data-a=copy]').onclick = async () => { try { await navigator.clipboard.writeText(a.replace(/\s*\[(?:עמ['׳]?|p\.)\s*[\d,\s–-]+\]/g, '')); toast('התשובה הועתקה'); } catch (_) { toast('ההעתקה לא הצליחה'); } };
  acts.querySelector('[data-a=read]').onclick = () => RA.speakText(a.replace(/\[[^\]]*\]/g, ''));
  el.append(acts);
  el.addEventListener('click', e => { const b = e.target.closest('.cite'); if (b) goPage(+b.dataset.p); });
  return el;
}
function goPage(n) {
  const P = S.pages[n - 1]; if (!P) return;
  P.el.scrollIntoView({ behavior: matchMedia('(prefers-reduced-motion:reduce)').matches ? 'auto' : 'smooth', block: 'start' });
  if (matchMedia('(max-width:720px)').matches) hideAI();
}
function renderLog() {
  const log = $('aiLog'); log.innerHTML = '';
  if (!hasDoc()) {
    log.append(h(`<div class="ai-hello"><b>פתחו קובץ ואשאל, אסכם ואקריא אותו.</b><br>בלי קובץ פתוח אפשר כבר עכשיו להפוך דף בכתב יד או סריקה למסמך מודפס.</div>`));
    $('aiQ').disabled = $('aiSend').disabled = true; $('aiRead').disabled = true;
    return;
  }
  $('aiQ').disabled = $('aiSend').disabled = AS.busy; $('aiRead').disabled = false;
  const items = AS.list();
  if (!items.length) {
    log.append(h(`<div class="ai-hello">שאלו כל שאלה על <b></b>, או בחרו:</div>`));
    log.lastChild.querySelector('b').textContent = dispName(docTitle());
    const chips = h('<div class="ai-chips"></div>');
    for (const [l, q] of CHIPS) { const b = h(`<button class="ai-chip" type="button">${esc(l)}</button>`); b.onclick = () => askQ(q, l); chips.append(b); }
    log.append(chips);
  }
  for (const it of items) {
    const q = h('<div class="ai-q" dir="auto"></div>'); q.textContent = it.label || it.q; log.append(q);
    if (it.a) log.append(answerEl(it.a, it.q, it.plain));
    else if (it.err) { const e = h('<div class="ai-err"></div>'); e.textContent = it.err; if (it.scanned) e.append(scannedHelp()); log.append(e); }
    else log.append(h(`<div class="ai-wait" id="aiWait">${it.wait || 'קורא את המסמך וכותב תשובה…'}</div>`));
  }
  if (items.length) {
    const chips = h('<div class="ai-chips"></div>');
    for (const [l, q] of CHIPS.filter(c => !items.some(i => i.q === c[1])).slice(0, 3)) { const b = h(`<button class="ai-chip" type="button">${esc(l)}</button>`); b.onclick = () => askQ(q, l); chips.append(b); }
    if (!AS.busy && chips.childElementCount) log.append(chips);
  }
  log.scrollTop = log.scrollHeight;
}
function scannedHelp() {
  const w = h('<div><button class="btn" type="button">זיהוי טקסט בכל העמודים</button> <button class="btn" type="button">כתב יד / סריקה חכמה</button></div>');
  w.children[0].onclick = async () => { await runOCR(S.pages.map((_, i) => i)); textCache.delete(S.doc); };
  w.children[1].onclick = () => SC.open('doc');
  return w;
}
async function askQ(q, label) {
  q = String(q || '').trim();
  if (!q || AS.busy || !hasDoc()) return;
  if (panel.classList.contains('hidden')) showAI();
  const list = AS.list(), tabKey = AS.key();
  const it = { q, label }; list.push(it);
  AS.busy = true; renderLog();
  try {
    const text = await docText();
    it.plain = onEditor();
    const empty = it.plain ? !text.replace(/=====.*?=====/g, '').replace(/שורה \d+:/g, '').trim() : !hasRealText(text);
    if (empty) { if (it.plain) { it.err = 'הקובץ ריק, אין עדיין על מה לשאול'; return; } it.err = 'לא מצאתי טקסט במסמך. כנראה שהוא סרוק או מצולם — אפשר לזהות בו את הטקסט קודם:'; it.scanned = true; return; }
    const history = list.filter(x => x.a && x !== it).slice(-4).map(x => ({ q: x.q, a: x.a }));
    it.a = await ai({ mode: 'ask', text, question: q, history }, { onWait: s => { it.wait = `השירות עמוס, מנסה שוב בעוד ${s} שניות…`; const w = document.getElementById('aiWait'); if (w && AS.key() === tabKey) w.textContent = it.wait; } });
  } catch (e) { it.err = e.message || String(e); }
  finally { AS.busy = false; if (AS.key() === tabKey) renderLog(); }
}
$('aiSend').onclick = () => { const q = $('aiQ').value; $('aiQ').value = ''; $('aiQ').style.height = ''; askQ(q); };
$('aiQ').addEventListener('input', e => { const t = e.target; t.style.height = 'auto'; t.style.height = Math.min(140, t.scrollHeight) + 'px'; });
$('aiRead').onclick = () => RA.start(currentPageIndex());
$('aiScan').onclick = () => SC.open();

// the assistant's button in the toolbar
const aiBtn = h(`<button class="btn ai" id="aiBtn" type="button" title="העוזר: שאלות על המסמך, סיכום, הקראה וכתב יד ל-Word">${ICON.spark}<span>עוזר</span></button>`);
$('trBtn').after(aiBtn);
aiBtn.onclick = toggleAI;

/* =====================================================================
   4. READ ALOUD
   ===================================================================== */
const BCP = { he: 'he-IL', en: 'en-US', ru: 'ru-RU', ar: 'ar-SA', fr: 'fr-FR', es: 'es-ES', de: 'de-DE', it: 'it-IT', pt: 'pt-PT', uk: 'uk-UA' };
const bar = h(`<div class="ra hidden" role="region" aria-label="הקראה">
  <button class="play" type="button" data-a="play" aria-label="השהיה">❚❚</button>
  <button type="button" data-a="prev" aria-label="המשפט הקודם" title="המשפט הקודם">↺</button>
  <button type="button" data-a="next" aria-label="המשפט הבא" title="המשפט הבא">↻</button>
  <span class="now" dir="auto"></span><span class="pg"></span>
  <select aria-label="מהירות"><option value=".8">0.8×</option><option value="1" selected>1×</option><option value="1.2">1.2×</option><option value="1.5">1.5×</option></select>
  <button type="button" data-a="stop" aria-label="סיום ההקראה" title="סיום">✕</button></div>`);
document.body.append(bar);
const RA = {
  q: [], at: 0, on: false, paused: false, tok: 0, rate: 1, lang: 'he', voice: null, warned: {},
  supported: () => 'speechSynthesis' in window && 'SpeechSynthesisUtterance' in window,
  pick(lang) {
    const vs = speechSynthesis.getVoices(), code = (BCP[lang] || lang).toLowerCase(), short = code.slice(0, 2);
    const ok = vs.filter(v => (v.lang || '').toLowerCase().replace('_', '-').startsWith(short) || (short === 'he' && /^iw/i.test(v.lang)));
    return ok.find(v => (v.lang || '').toLowerCase().replace('_', '-') === code && v.localService) || ok.find(v => (v.lang || '').toLowerCase().replace('_', '-') === code) || ok[0] || null;
  },
  split(text, page) {
    const out = [];
    for (const para of String(text || '').split(/\n+/)) {
      const p = para.replace(/\s+/g, ' ').trim(); if (!p) continue;
      for (let s of p.replace(/([.!?؟…:;])\s+/g, '$1\u0001').split('\u0001')) {
        s = s.trim(); if (!s || !/[\p{L}\p{N}]/u.test(s)) continue;
        while (s.length > 220) { let cut = s.lastIndexOf(', ', 200); if (cut < 60) cut = s.lastIndexOf(' ', 200); if (cut < 60) cut = 200; out.push({ t: s.slice(0, cut + 1).trim(), page }); s = s.slice(cut + 1).trim(); }
        if (s) out.push({ t: s, page });
      }
    }
    return out;
  },
  async start(from) {
    if (!this.supported()) { toast('הדפדפן הזה לא תומך בהקראה'); return; }
    if (onEditor()) { const t = (TB.cur.ed === 'pptx' ? await PPTX_EDITOR.textOf() : editorText()).replace(/=====.*?=====/g, '').replace(/^שורה \d+: /gm, '').replace(/\t/g, ', '); if (t.trim()) this.speakText(t); else toast('אין טקסט להקראה'); return; }
    if (!S.doc) { toast('פתחו קובץ כדי להקריא אותו'); return; }
    this.stop();
    busy(true, 'מכין את ההקראה…');
    const q = [];
    try { for (let i = Math.max(0, from || 0); i < S.pages.length; i++) q.push(...this.split(await pageText(i), i)); }
    finally { busy(false); }
    if (!q.length) { toast('לא נמצא טקסט להקראה. אם המסמך סרוק: "עמודים" ← "זיהוי טקסט (OCR)"'); return; }
    this.run(q, detectLocal(q.slice(0, 40).map(x => x.t).join(' ')) || 'he');
  },
  speakText(text) {
    if (!this.supported()) { toast('הדפדפן הזה לא תומך בהקראה'); return; }
    this.stop(); const q = this.split(text, -1); if (q.length) this.run(q, detectLocal(text) || 'he');
  },
  run(q, lang) {
    this.q = q; this.at = 0; this.on = true; this.paused = false; this.lang = lang;
    this.voice = this.pick(lang);
    if (!this.voice && !this.warned[lang]) {
      this.warned[lang] = 1;
      toast(`אין במכשיר הזה קול ב${LANG_HE[lang] || 'שפה הזו'}. אפשר להוסיף קול בהגדרות המערכת (שפה ודיבור / Speech)`);
    }
    bar.classList.remove('hidden'); this.ui(); this.next(0);
  },
  next(step) {
    if (!this.on) return;
    this.at = Math.max(0, this.at + step);
    if (this.at >= this.q.length) { this.stop(); toast('ההקראה הסתיימה'); return; }
    const it = this.q[this.at], tok = ++this.tok;
    speechSynthesis.cancel();
    const u = new SpeechSynthesisUtterance(it.t);
    u.lang = BCP[this.lang] || this.lang; if (this.voice) u.voice = this.voice; u.rate = this.rate;
    u.onend = () => { if (tok === this.tok && this.on && !this.paused) this.next(1); };
    u.onerror = e => { if (tok !== this.tok || e.error === 'interrupted' || e.error === 'canceled') return; if (this.on && !this.paused) this.next(1); };
    this.ui();
    if (it.page >= 0 && S.pages[it.page]) {
      const r = S.pages[it.page].el.getBoundingClientRect(), d = $('desk').getBoundingClientRect();
      if (r.bottom < d.top + 40 || r.top > d.bottom - 80) S.pages[it.page].el.scrollIntoView({ block: 'start' });
    }
    setTimeout(() => { if (tok === this.tok) speechSynthesis.speak(u); }, 40);
  },
  toggle() {
    if (!this.on) return;
    if (this.paused) { this.paused = false; this.next(0); }
    else { this.paused = true; this.tok++; speechSynthesis.cancel(); }
    this.ui();
  },
  stop() {
    if (!this.on && bar.classList.contains('hidden')) return;
    this.on = false; this.paused = false; this.tok++;
    try { speechSynthesis.cancel(); } catch (_) {}
    bar.classList.add('hidden');
  },
  ui() {
    const it = this.q[this.at];
    bar.querySelector('.now').textContent = it ? it.t : '';
    bar.querySelector('.pg').textContent = it && it.page >= 0 ? `עמ' ${it.page + 1}` : '';
    const p = bar.querySelector('.play'); p.textContent = this.paused ? '▶' : '❚❚'; p.setAttribute('aria-label', this.paused ? 'המשך' : 'השהיה');
  },
};
if (RA.supported()) speechSynthesis.addEventListener && speechSynthesis.addEventListener('voiceschanged', () => { if (RA.on && !RA.voice) RA.voice = RA.pick(RA.lang); });
bar.querySelector('[data-a=play]').onclick = () => RA.toggle();
bar.querySelector('[data-a=next]').onclick = () => { RA.paused = false; RA.next(1); };
bar.querySelector('[data-a=prev]').onclick = () => { RA.paused = false; RA.next(-1); };
bar.querySelector('[data-a=stop]').onclick = () => RA.stop();
bar.querySelector('select').onchange = e => { RA.rate = +e.target.value; if (RA.on && !RA.paused) RA.next(0); };
addEventListener('pagehide', () => RA.stop());
window.SIMAN_READ = RA;

/* =====================================================================
   5. SMART SCAN — handwriting and scanned pages → printed Word / PDF
   ===================================================================== */
const SC_LANGS = [['', 'זיהוי אוטומטי'], ['he', 'עברית'], ['en', 'אנגלית'], ['ru', 'רוסית'], ['ar', 'ערבית'], ['fr', 'צרפתית'], ['es', 'ספרדית']];
const dlg = h(`<dialog class="scan" id="scanDlg" aria-label="כתב יד וסריקות למסמך מודפס">
  <div class="sc-head"><h2>כתב יד וסריקות ← מסמך מודפס</h2><button class="icon-btn" type="button" data-a="close" aria-label="סגירה">×</button></div>
  <div class="sc-body"></div>
  <div class="sc-foot hidden"></div>
</dialog>`);
document.body.append(dlg);
const camIn = h('<input type="file" accept="image/*" capture="environment" hidden>');
const imgIn = h('<input type="file" accept="image/*" multiple hidden>');
const pdfIn = h('<input type="file" accept="application/pdf,.pdf" hidden>');
dlg.append(camIn, imgIn, pdfIn);
const SC = {
  pages: [], stop: false, running: false, hint: '', name: '', quality: 'accurate',
  body: () => dlg.querySelector('.sc-body'), foot: () => dlg.querySelector('.sc-foot'),
  // a modal dialog covers the app's toasts and busy screen, so messages are shown inside it
  note(msg) {
    const f = this.foot(), hint = !f.classList.contains('hidden') && f.querySelector('.hint');
    if (hint) { hint.textContent = msg; return; }
    let n = dlg.querySelector('.sc-note');
    if (!n) { n = h('<p class="sc-note" role="status"></p>'); this.body().prepend(n); }
    n.textContent = msg;
  },
  wait(msg) { this.foot().classList.add('hidden'); this.body().innerHTML = `<p class="ai-wait" style="font-size:15px">${esc(msg)}</p>`; },
  open(src) {
    if (typeof needReg === 'function' && needReg()) { showReg(); return; }
    this.pick();
    if (!dlg.open) dlg.showModal();
    if (src === 'doc' && S.doc) this.fromDoc();
  },
  close() {
    if (this.running) { this.stop = true; }
    else if (this.pages.some(p => p.blocks) && !confirm('לסגור? הטקסט שזוהה לא יישמר')) return;
    dlg.close();
  },
  pick() {
    this.pages = []; this.stop = false; this.running = false;
    const docN = S.doc ? S.pages.length : 0;
    const b = this.body(); this.foot().classList.add('hidden');
    b.innerHTML = `<div class="sc-pick">
      <h3>מצלמים דף — מקבלים מסמך מודפס</h3>
      <p class="sub">כתב יד, צילום של מסמך או PDF סרוק: סימן קורא את הכתוב, גם בכתב יד, ובונה ממנו קובץ Word או PDF עם כותרות, רשימות וטבלאות. לפני השמירה רואים את התמונה ואת הטקסט זה לצד זה ומתקנים. מילים שהזיהוי לא בטוח בהן מסומנות בצהוב.</p>
      <div class="sc-srcs">
        <button class="sc-src" type="button" data-s="cam">${ICON.scan}<b>צילום במצלמה</b><span>בטלפון: נפתחת המצלמה. מצלמים ישר מלמעלה, באור טוב</span></button>
        <button class="sc-src" type="button" data-s="img"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"><rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="9" cy="10" r="1.6"/><path d="M21 16l-5-5-8 8"/></svg><b>בחירת תמונות</b><span>אפשר כמה תמונות יחד. כל תמונה היא עמוד</span></button>
        <button class="sc-src" type="button" data-s="pdf"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"><path d="M7 3h7l4 4v14H7z"/><path d="M14 3v4h4"/></svg><b>PDF סרוק</b><span>קובץ שנסרק או צולם, עד 30 עמודים</span></button>
        ${docN ? `<button class="sc-src" type="button" data-s="doc"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"><rect x="5" y="3" width="14" height="18" rx="1.5"/><path d="M9 8h6M9 12h6M9 16h3"/></svg><b>הקובץ הפתוח</b><span></span></button>` : ''}
      </div>
      <div class="sc-opts"><label>שפת הכתב <select data-o="hint">${SC_LANGS.map(([v, l]) => `<option value="${v}">${l}</option>`).join('')}</select></label>
        <label>אופן הקריאה <select data-o="quality"><option value="accurate">מדויק: הכי טוב לכתב יד (חצי דקה עד שתי דקות לעמוד)</option><option value="fast">מהיר: כמה שניות לעמוד, מנחש יותר</option></select></label></div>
      <p class="sc-fine">התמונות נשלחות לזיהוי אצל Google Gemini ולא נשמרות אצלנו. בשירות החינמי Google עשויה להשתמש בהן לשיפור השירות שלה, לכן לא כדאי לסרוק כך מסמכים רגישים במיוחד (תעודות זהות, מסמכים רפואיים). כל עמוד לוקח כמה שניות.</p>
    </div>`;
    if (docN) b.querySelector('[data-s=doc] span').textContent = `${dispName(S.office ? S.office.name : S.name)} · ${docN} ${docN === 1 ? 'עמוד' : 'עמודים'}`;
    const sel = b.querySelector('[data-o=hint]'); sel.value = this.hint; sel.onchange = () => { this.hint = sel.value; };
    const q = b.querySelector('[data-o=quality]'); q.value = this.quality; q.onchange = () => { this.quality = q.value; };
    b.querySelector('[data-s=cam]').onclick = () => camIn.click();
    b.querySelector('[data-s=img]').onclick = () => imgIn.click();
    b.querySelector('[data-s=pdf]').onclick = () => pdfIn.click();
    const d = b.querySelector('[data-s=doc]'); if (d) d.onclick = () => this.fromDoc();
  },
  async fromImages(files) {
    if (!files.length) return;
    this.name = files.length === 1 ? files[0].name.replace(/\.[a-z0-9]+$/i, '') : 'כתב יד';
    const pics = [];
    for (const f of files.slice(0, 30)) {
      try { pics.push(await imgToJpeg(f)); }
      catch (e) { this.note(`לא הצלחתי לקרוא את התמונה ${f.name}${/heic|heif/i.test(f.type + f.name) ? ' (HEIC). צלמו ישר מהמצלמה או שמרו אותה כ-JPG' : ''}`); }
    }
    if (pics.length) this.run(pics);
  },
  async fromPdf(file) {
    if (!file) return;
    this.name = file.name.replace(/\.pdf$/i, '');
    this.wait('קורא את ה-PDF…');
    try {
      const d = await pdfjs.getDocument({ data: new Uint8Array(await file.arrayBuffer()), isEvalSupported: false }).promise;
      const pics = [];
      for (let i = 1; i <= Math.min(30, d.numPages); i++) { this.wait(`מכין עמוד ${i} מתוך ${Math.min(30, d.numPages)}…`); pics.push(await pageToJpeg(await d.getPage(i))); }
      d.destroy(); this.run(pics, d.numPages > 30 ? 'נסרקים 30 העמודים הראשונים. ' : '');
    } catch (e) { console.error(e); this.pick(); this.note(e && e.name === 'PasswordException' ? 'הקובץ מוגן בסיסמה' : 'לא הצלחתי לפתוח את ה-PDF'); }
  },
  async fromDoc() {
    if (!S.doc) return;
    this.name = (S.office ? S.office.base || S.office.name : S.name).replace(/\.(pdf|docx|xlsx|pptx)$/i, '');
    try {
      const pics = [], n = Math.min(30, S.pages.length);
      for (const P of S.pages.slice(0, 30)) { this.wait(`מכין עמוד ${P.i} מתוך ${n}…`); pics.push(await pageToJpeg(P.page)); }
      this.run(pics, S.pages.length > 30 ? 'נסרקים 30 העמודים הראשונים. ' : '');
    } catch (e) { console.error(e); this.pick(); this.note('הכנת העמודים נכשלה'); }
  },
  async run(pics, pre = '') {
    this.pages = pics.map(p => ({ ...p, blocks: null, err: null }));
    this.running = true; this.stop = false;
    const b = this.body();
    b.innerHTML = `<p class="sc-status" role="status" style="margin:0 0 14px;font-size:15px"></p><ul class="sc-prog">${this.pages.map((p, i) => `<li><img alt="עמוד ${i + 1}" src="${p.url}"><span>ממתין</span></li>`).join('')}</ul>`;
    const f = this.foot(); f.classList.remove('hidden');
    f.innerHTML = '<span class="hint" style="margin-inline-end:auto"></span><button class="btn" type="button" data-a="stop">עצירה</button>';
    f.querySelector('[data-a=stop]').onclick = () => { this.stop = true; };
    const st = b.querySelector('.sc-status'), lis = [...b.querySelectorAll('.sc-prog li')];
    for (let i = 0; i < this.pages.length; i++) {
      if (this.stop) break;
      const p = this.pages[i], li = lis[i];
      li.className = 'now'; li.querySelector('span').textContent = 'קורא…';
      const t0 = Date.now(), line = `${i ? '' : pre}קורא את עמוד ${i + 1} מתוך ${this.pages.length}`;
      st.textContent = line + '…';
      const clock = setInterval(() => { const sec = Math.round((Date.now() - t0) / 1000); li.querySelector('span').textContent = `קורא… ${sec} שנ׳`; st.textContent = `${line} · ${sec} שניות${sec > 45 ? ' (כתב יד צפוף לוקח לפעמים עד שתי דקות)' : ''}`; }, 1000);
      try {
        const raw = await ai({ mode: 'scan', image: p.b64, mime: 'image/jpeg', hint: this.hint, quality: this.quality }, { onWait: s => { st.textContent = `השירות עמוס, ממשיך בעוד ${s} שניות…`; } }).finally(() => clearInterval(clock));
        p.blocks = parseBlocks(raw); p.lang = p.blocks.lang;
        li.className = 'ok'; li.querySelector('span').textContent = p.blocks.length ? `${p.blocks.length} קטעים` : 'עמוד ריק';
      } catch (e) {
        p.err = e.message || String(e); li.className = 'bad'; li.querySelector('span').textContent = 'לא נקרא';
        if (/מכסה|לא הוקם|חוסם|אין חיבור/.test(p.err)) { st.textContent = p.err; this.stop = true; break; }
      }
      if (i < this.pages.length - 1 && !this.stop) await sleep(1200);   // the free service allows only a few pages a minute
    }
    this.running = false;
    const done = this.pages.filter(p => p.blocks);
    if (!done.length) {
      st.textContent = (this.pages.find(p => p.err) || {}).err || 'לא זוהה טקסט';
      f.innerHTML = '<span class="hint" style="margin-inline-end:auto"></span><button class="btn" type="button" data-a="again">התחלה מחדש</button>';
      f.querySelector('[data-a=again]').onclick = () => this.pick();
      return;
    }
    this.review();
  },
  review() {
    const b = this.body();
    b.innerHTML = '<div class="sc-rev"></div>';
    const rev = b.firstChild;
    this.pages.forEach((p, i) => {
      if (!p.blocks && !p.err) return;
      const row = h(`<section class="sc-pg" aria-label="עמוד ${i + 1}"><figure><img alt="המקור, עמוד ${i + 1}" src="${p.url}"><figcaption>המקור · עמוד ${i + 1}${p.err ? ' · לא נקרא: ' + esc(p.err) : ''} · <button type="button" class="sc-again">לקרוא את העמוד שוב במצב מדויק</button></figcaption></figure><div class="sc-ed" contenteditable="true" spellcheck="true"></div></section>`);
      const ed = row.querySelector('.sc-ed');
      ed.innerHTML = p.blocks && p.blocks.length ? blocksHtml(p.blocks) : '<p class="empty">לא זוהה טקסט בעמוד הזה. אפשר להקליד כאן.</p>';
      ed.dir = p.blocks && p.blocks.length ? (blocksRtl(p.blocks) ? 'rtl' : 'ltr') : 'rtl';
      ed.addEventListener('focus', () => { try { document.execCommand('defaultParagraphSeparator', false, 'p'); } catch (_) {} });
      ed.addEventListener('input', () => { ed.querySelectorAll('mark').forEach(m => { if (!m.textContent.trim()) m.remove(); }); this.count(); });
      ed.addEventListener('click', e => { const m = e.target.closest('mark'); if (m && !m.dataset.seen) { m.dataset.seen = 1; } });
      row.querySelector('img').onclick = () => { const z = h(`<div class="sc-zoom"><img alt="" src="${p.url}"></div>`); z.onclick = () => z.remove(); dlg.append(z); };
      p.ed = ed;
      row.querySelector('.sc-again').onclick = () => this.again(i, row);
      rev.append(row);
    });
    const f = this.foot(); f.classList.remove('hidden');
    f.innerHTML = `<label>שם הקובץ <input type="text" data-o="name" dir="auto"></label>
      <span class="hint" data-o="unsure"></span>
      <button class="btn" type="button" data-a="copy">העתקת הטקסט</button>
      <button class="btn" type="button" data-a="edit" title="פתיחה בעורך ה-Word של סימן">עריכה בעורך Word</button>
      <button class="btn" type="button" data-a="pdf">שמירה כ-PDF</button>
      <button class="btn primary" type="button" data-a="docx">שמירה כ-Word</button>`;
    const nm = f.querySelector('[data-o=name]'); nm.value = this.name || 'סריקה';
    nm.oninput = () => { this.name = nm.value; };
    f.querySelector('[data-a=docx]').onclick = () => this.save('docx');
    f.querySelector('[data-a=pdf]').onclick = () => this.save('pdf');
    f.querySelector('[data-a=edit]').onclick = () => this.save('edit');
    f.querySelector('[data-a=copy]').onclick = async () => { try { await navigator.clipboard.writeText(this.collect().map(pg => pg.map(blockText).join('\n')).join('\n\n')); toast('הטקסט הועתק'); } catch (_) { toast('ההעתקה לא הצליחה'); } };
    this.count();
  },
  async again(i, row) {
    const p = this.pages[i], btn = row.querySelector('.sc-again'), ed = row.querySelector('.sc-ed');
    if (this.running) return;
    this.running = true; btn.disabled = true;
    const t0 = Date.now(), clock = setInterval(() => { btn.textContent = `קורא שוב… ${Math.round((Date.now() - t0) / 1000)} שנ׳`; }, 1000);
    try {
      const raw = await ai({ mode: 'scan', image: p.b64, mime: 'image/jpeg', hint: this.hint, quality: 'accurate' });
      const bl = parseBlocks(raw);
      if (bl.length) { p.blocks = bl; ed.innerHTML = blocksHtml(bl); ed.dir = blocksRtl(bl) ? 'rtl' : 'ltr'; this.count(); btn.textContent = 'נקרא שוב ✓'; }
      else btn.textContent = 'לא זוהה טקסט';
    } catch (e) { btn.textContent = 'נכשל: ' + (e.message || e); btn.disabled = false; }
    finally { clearInterval(clock); this.running = false; }
  },
  count() {
    const n = this.pages.reduce((t, p) => t + (p.ed ? p.ed.querySelectorAll('mark').length : 0), 0);
    const el = this.foot().querySelector('[data-o=unsure]'); if (!el) return;
    el.innerHTML = n ? `${n} ${n === 1 ? 'מילה מסומנת' : 'מילים מסומנות'} לבדיקה · <button type="button">למילה הבאה</button>` : 'אין מילים מסומנות. אפשר לשמור';
    const btn = el.querySelector('button');
    if (btn) btn.onclick = () => {
      const ms = this.pages.flatMap(p => p.ed ? [...p.ed.querySelectorAll('mark')] : []);
      const m = ms.find(x => !x.dataset.seen) || ms[0]; if (!m) return;
      m.dataset.seen = 1; m.scrollIntoView({ block: 'center', behavior: 'smooth' });
      const r = document.createRange(); r.selectNodeContents(m); const s = getSelection(); s.removeAllRanges(); s.addRange(r); m.closest('.sc-ed').focus();
    };
  },
  collect() { return this.pages.filter(p => p.ed).map(p => domBlocks(p.ed)).filter(pg => pg.length); },
  async save(kind) {
    const pages = this.collect();
    if (!pages.length) { toast('אין טקסט לשמירה'); return; }
    const base = cleanBase(this.name) || 'סריקה';
    try {
      if (kind === 'pdf') {
        this.note('בונה PDF…'); const bytes = await blocksPdf(pages);
        if (await saveBytes(base + '.pdf', bytes, 'pdf')) this.note('נשמר כ-PDF ✓'); else this.count();
        return;
      }
      const bytes = blocksDocx(pages);
      if (kind === 'edit') { dlg.close(); await openWordEditor(bytes, base + '.docx'); return; }
      if (await saveBytes(base + '.docx', bytes, 'docx')) this.note('נשמר כקובץ Word ✓'); else this.count();
    } catch (e) { console.error(e); this.note('השמירה נכשלה: ' + (e.message || e)); }
  },
};
dlg.querySelector('[data-a=close]').onclick = () => SC.close();
dlg.addEventListener('cancel', e => { if (SC.running || SC.pages.some(p => p.blocks)) { e.preventDefault(); SC.close(); } });
camIn.onchange = e => { const f = [...e.target.files]; e.target.value = ''; SC.fromImages(f); };
imgIn.onchange = e => { const f = [...e.target.files]; e.target.value = ''; SC.fromImages(f); };
pdfIn.onchange = e => { const f = e.target.files[0]; e.target.value = ''; SC.fromPdf(f); };
window.SIMAN_SCAN = SC;

// pictures for the service: long side ~2000px, JPEG
function canvasJpeg(draw, w, hgt) {
  const s = Math.min(1, (SC.quality === 'fast' ? 1700 : 2400) / Math.max(w, hgt));
  const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(w * s)); c.height = Math.max(1, Math.round(hgt * s));
  const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height);
  return Promise.resolve(draw(x, c.width, c.height)).then(() => {
    const url = c.toDataURL('image/jpeg', 0.86);
    return { url, b64: url.slice(url.indexOf(',') + 1), w: c.width, h: c.height };
  });
}
async function imgToJpeg(file) {
  let bmp;
  try { bmp = await createImageBitmap(file, { imageOrientation: 'from-image' }); }
  catch (_) { bmp = await new Promise((res, rej) => { const im = new Image(); im.onload = () => res(im); im.onerror = rej; im.src = URL.createObjectURL(file); }); }
  const w = bmp.width || bmp.naturalWidth, hh = bmp.height || bmp.naturalHeight;
  return canvasJpeg((x, cw, ch) => x.drawImage(bmp, 0, 0, cw, ch), w, hh);
}
async function pageToJpeg(page) {
  const v1 = page.getViewport({ scale: 1 }), sc = Math.min(4, (SC.quality === 'fast' ? 1700 : 2400) / Math.max(v1.width, v1.height));
  const vp = page.getViewport({ scale: sc });
  return canvasJpeg(async x => { await page.render({ canvasContext: x, viewport: vp }).promise; }, vp.width, vp.height);
}

// the service's JSON → clean blocks
function parseBlocks(raw) {
  let j = null;
  const s = String(raw || '').replace(/^\s*```(?:json)?\s*/i, '').replace(/\s*```\s*$/, '');
  try { j = JSON.parse(s); } catch (_) { const m = /\{[\s\S]*\}/.exec(s); if (m) { try { j = JSON.parse(m[0]); } catch (_) {} } }
  if (!j) j = salvage(s);
  if (!j) { console.warn('scan: unreadable answer', raw); throw new Error('התשובה מהשירות לא הייתה תקינה. נסו שוב'); }
  const arr = Array.isArray(j) ? j : Array.isArray(j.blocks) ? j.blocks : [];
  const out = [];
  for (const b of arr) {
    if (!b || typeof b !== 'object') continue;
    const t = ['h1', 'h2', 'p', 'li', 'oli', 'table'].includes(b.t) ? b.t : 'p';
    if (t === 'table') {
      const rows = (Array.isArray(b.rows) ? b.rows : []).filter(Array.isArray).map(r => r.map(c => String(c ?? '')));
      if (rows.length) out.push({ t, rows });
    } else {
      const text = String(b.text ?? '').replace(/\s+/g, ' ').trim();
      if (text) out.push({ t, text });
    }
  }
  out.lang = typeof j.lang === 'string' ? j.lang.slice(0, 2).toLowerCase() : null;
  return out;
}
// a cut or broken answer: keep every block that is whole; plain text becomes paragraphs
function salvage(s) {
  const found = [];
  for (const m of s.matchAll(/\{[^{}\[\]]*\}|\{[^{}]*"rows"\s*:\s*\[[\s\S]*?\]\s*\]\s*\}/g)) {
    try { const o = JSON.parse(m[0]); if (o && (o.text || o.rows)) found.push({ at: m.index, o }); } catch (_) {}
  }
  if (found.length) { console.warn('scan: answer was broken, kept', found.length, 'blocks'); return { blocks: found.sort((a, b) => a.at - b.at).map(x => x.o) }; }
  if (!/^\s*[\[{]/.test(s) && s.trim()) return { blocks: s.split(/\n+/).map(t => t.trim()).filter(Boolean).map(text => ({ t: 'p', text })) };
  return null;
}
const unsureHtml = t => esc(t).replace(/\[\[(.*?)\]\]/g, (m, w) => w.trim() === '?' ? '<mark title="לא הצלחתי לקרוא">[?]</mark>' : `<mark title="לא בטוח, כדאי לבדוק">${w}</mark>`);
const dirOf = t => paraDir(String(t || ''), 'R') === 'R' ? 'rtl' : 'ltr';
function blocksHtml(blocks) {
  let html = '', list = null;
  const endList = () => { if (list) { html += `</${list}>`; list = null; } };
  for (const b of blocks) {
    if (b.t === 'li' || b.t === 'oli') { const tag = b.t === 'li' ? 'ul' : 'ol'; if (list !== tag) { endList(); html += `<${tag}>`; list = tag; } html += `<li dir="${dirOf(b.text)}">${unsureHtml(b.text)}</li>`; continue; }
    endList();
    if (b.t === 'table') html += `<table>${b.rows.map(r => `<tr>${r.map(c => `<td dir="${dirOf(c)}">${unsureHtml(c)}</td>`).join('')}</tr>`).join('')}</table>`;
    else { const tag = b.t === 'h1' ? 'h1' : b.t === 'h2' ? 'h2' : 'p'; html += `<${tag} dir="${dirOf(b.text)}">${unsureHtml(b.text)}</${tag}>`; }
  }
  endList();
  return html;
}
const blocksRtl = blocks => paraDir(blocks.map(b => b.text || (b.rows || []).flat().join(' ')).join(' '), 'R') === 'R';
const clean = t => String(t || '').replace(/\u00A0/g, ' ').replace(/[ \t]+/g, ' ').replace(/ *\n */g, '\n').trim();
// the edited page back to blocks
function domBlocks(ed) {
  const out = [];
  const visit = el => {
    for (const n of [...el.childNodes]) {
      if (n.nodeType === 3) { const t = clean(n.textContent); if (t) out.push({ t: 'p', text: t }); continue; }
      if (n.nodeType !== 1 || n.classList.contains('empty')) continue;
      const tag = n.tagName;
      if (tag === 'H1' || tag === 'H2') { const t = clean(n.innerText); if (t) out.push({ t: tag.toLowerCase(), text: t }); }
      else if (tag === 'UL' || tag === 'OL') { for (const li of n.querySelectorAll(':scope > li')) { const t = clean(li.innerText); if (t) out.push({ t: tag === 'UL' ? 'li' : 'oli', text: t }); } }
      else if (tag === 'TABLE') { const rows = [...n.querySelectorAll('tr')].map(tr => [...tr.children].map(td => clean(td.innerText))); if (rows.length) out.push({ t: 'table', rows }); }
      else if (tag === 'BR') continue;
      else if (tag === 'DIV' && n.querySelector('p,h1,h2,ul,ol,table,div')) visit(n);
      else { for (const part of clean(n.innerText).split('\n')) if (part.trim()) out.push({ t: 'p', text: part.trim() }); }
    }
  };
  visit(ed);
  return out;
}
const blockText = b => b.t === 'table' ? b.rows.map(r => r.join('\t')).join('\n') : (b.t === 'li' ? '• ' : '') + b.text;

/* ---------- Word (.docx) from blocks — RTL aware ---------- */
function blocksDocx(pages) {
  const X = t => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const LANGTAG = { ru: 'ru-RU', fr: 'fr-FR', es: 'es-ES', en: 'en-US', he: 'en-US', ar: 'en-US' };
  const all = pages.flat().map(b => b.text || (b.rows || []).flat().join(' ')).join(' ');
  const docRtl = paraDir(all, 'R') === 'R';
  const lang = detectLocal(all) || (docRtl ? 'he' : 'en');
  const run = (t, rtl, extra = '') => `<w:r><w:rPr>${extra}${rtl ? '<w:rtl/>' : ''}<w:lang w:val="${LANGTAG[lang] || 'en-US'}" w:bidi="${lang === 'ar' ? 'ar-SA' : 'he-IL'}"/></w:rPr><w:t xml:space="preserve">${X(t)}</w:t></w:r>`;
  const para = (t, { style, num, keep } = {}) => {
    const rtl = paraDir(t, docRtl ? 'R' : 'L') === 'R';
    return `<w:p><w:pPr>${style ? `<w:pStyle w:val="${style}"/>` : ''}${keep ? '<w:keepNext/>' : ''}${num ? `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${num}"/></w:numPr>` : ''}${rtl ? '<w:bidi/>' : ''}</w:pPr>${run(t, rtl)}</w:p>`;
  };
  let nextNum = 2; const nums = [];
  const body = [];
  pages.forEach((blocks, pi) => {
    if (pi) body.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
    let lastOl = false, olId = 0;
    for (const b of blocks) {
      if (b.t !== 'oli') lastOl = false;
      if (b.t === 'h1') body.push(para(b.text, { style: 'Heading1', keep: true }));
      else if (b.t === 'h2') body.push(para(b.text, { style: 'Heading2', keep: true }));
      else if (b.t === 'li') body.push(para(b.text, { style: 'ListParagraph', num: 1 }));
      else if (b.t === 'oli') { if (!lastOl) { olId = nextNum++; nums.push(olId); lastOl = true; } body.push(para(b.text, { style: 'ListParagraph', num: olId })); }
      else if (b.t === 'table') {
        const n = Math.max(...b.rows.map(r => r.length)), cw = Math.floor(9638 / n);
        const tRtl = paraDir(b.rows.flat().join(' '), docRtl ? 'R' : 'L') === 'R';
        const bd = s => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="808080"/>`;
        body.push(`<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/>${tRtl ? '<w:bidiVisual/>' : ''}<w:tblW w:w="5000" w:type="pct"/><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(bd).join('')}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>${('<w:gridCol w:w="' + cw + '"/>').repeat(n)}</w:tblGrid>${b.rows.map((r, ri) => `<w:tr>${Array.from({ length: n }, (_, k) => r[k] ?? '').map(c => `<w:tc><w:tcPr><w:tcW w:w="${cw}" w:type="dxa"/></w:tcPr>${c ? (ri === 0 ? para(c, { style: 'TableHead' }) : para(c, { style: 'TableText' })) : '<w:p><w:pPr><w:pStyle w:val="TableText"/></w:pPr></w:p>'}</w:tc>`).join('')}</w:tr>`).join('')}</w:tbl>`);
        body.push('<w:p/>');
      } else body.push(para(b.text));
    }
  });
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  const lvl = (fmt, txt) => `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${txt}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:start="720" w:hanging="360"/></w:pPr>${fmt === 'bullet' ? '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/></w:rPr>' : ''}</w:lvl>`;
  const numbering = `${head}<w:numbering ${NS}><w:abstractNum w:abstractNumId="0">${lvl('bullet', '•')}</w:abstractNum><w:abstractNum w:abstractNumId="1">${lvl('decimal', '%1.')}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>${nums.map(n => `<w:num w:numId="${n}"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>`).join('')}</w:numbering>`;
  const font = lang === 'ar' ? 'Arial' : 'Arial';
  const styles = `${head}<w:styles ${NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}" w:eastAsia="${font}"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="${LANGTAG[lang] || 'en-US'}" w:bidi="${lang === 'ar' ? 'ar-SA' : 'he-IL'}"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="140" w:line="300" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="160"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:bCs/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="60"/><w:contextualSpacing/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="TableText"><w:name w:val="Table Text"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="40" w:after="40" w:line="260" w:lineRule="auto"/></w:pPr><w:rPr><w:sz w:val="21"/><w:szCs w:val="21"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="TableHead"><w:name w:val="Table Head"/><w:basedOn w:val="TableText"/><w:rPr><w:b/><w:bCs/></w:rPr></w:style>
<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblCellMar><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>`;
  const enc = new TextEncoder();
  const doc = `${head}<w:document ${NS} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${body.join('')}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1247" w:right="1247" w:bottom="1247" w:left="1247" w:header="708" w:footer="708" w:gutter="0"/>${docRtl ? '<w:bidi/>' : ''}</w:sectPr></w:body></w:document>`;
  const files = [
    ['[Content_Types].xml', `${head}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`],
    ['_rels/.rels', `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`],
    ['docProps/core.xml', `${head}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>סימן</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString().slice(0, 19)}Z</dcterms:created></cp:coreProperties>`],
    ['word/_rels/document.xml.rels', `${head}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`],
    ['word/document.xml', doc], ['word/styles.xml', styles], ['word/numbering.xml', numbering],
  ];
  return zip(files.map(([name, s]) => ({ name, data: enc.encode(s) })));
}

/* ---------- a printed PDF from blocks (real text: searchable and copyable) ---------- */
async function blocksPdf(pages) {
  const { PDFDocument, rgb } = PDFLib;
  const doc = await PDFDocument.create(); doc.registerFontkit(fontkit);
  const reg = await doc.embedFont(await fontBytes('rubik', false), { subset: true });
  const bold = await doc.embedFont(await fontBytes('rubik', true), { subset: true });
  const W = 595.28, H = 841.89, M = 56, maxW = W - 2 * M, ink = rgb(.1, .11, .13), rule = rgb(.55, .56, .55);
  const all = pages.flat().map(b => b.text || (b.rows || []).flat().join(' ')).join(' ');
  const docDir = paraDir(all, 'R');
  let pg, y;
  const newPage = () => { pg = doc.addPage([W, H]); y = H - M; };
  // arrows and check marks are in none of the app's fonts: they are drawn as lines
  const SYM = /([→←↑↓⇒⇐➔➜►◄✓✔])/;
  const norm = t => String(t).replace(/[✗✘❌]/g, '×').replace(/[●○■□▲▼★◆◇▪]/g, '•').replace(/[⇒➔➜►]/g, '→').replace(/[⇐◄]/g, '←');
  const drawSym = (c, x, yy, sz, w) => {
    const th = Math.max(.7, sz * .07), mid = yy + sz * .32, a = sz * .22, L = (x1, y1, x2, y2) => pg.drawLine({ start: { x: x1, y: y1 }, end: { x: x2, y: y2 }, thickness: th, color: ink });
    if (c === '→' || c === '←') {
      const l = x + w * .12, r = x + w * .88; L(l, mid, r, mid);
      if (c === '→') { L(r, mid, r - a, mid + a * .8); L(r, mid, r - a, mid - a * .8); } else { L(l, mid, l + a, mid + a * .8); L(l, mid, l + a, mid - a * .8); }
    } else if (c === '↓' || c === '↑') {
      const cx = x + w / 2, top = yy + sz * .78, bot = yy - sz * .08; L(cx, top, cx, bot);
      if (c === '↓') { L(cx, bot, cx - a * .8, bot + a); L(cx, bot, cx + a * .8, bot + a); } else { L(cx, top, cx - a * .8, top - a); L(cx, top, cx + a * .8, top - a); }
    } else { L(x + w * .18, mid, x + w * .4, yy + sz * .05); L(x + w * .4, yy + sz * .05, x + w * .85, yy + sz * .72); } // ✓
  };
  const lineAt = (line, x0, w, yy, f, sz, dir) => {
    const { base, runs } = bidiRuns(norm(line), dir);
    const parts = [];
    for (const r of runs) {
      if (!SYM.test(r.text)) { const s = runForPdf(r); parts.push({ s, w: f.widthOfTextAtSize(s, sz) }); continue; }
      const segs = r.text.split(SYM).filter(Boolean).map(t => SYM.test(t) && t.length === 1 ? { sym: t, w: sz * .95 } : (() => { const s = runForPdf({ dir: r.dir, text: t }); return { s, w: f.widthOfTextAtSize(s, sz) }; })());
      parts.push(...(r.dir === 'R' ? segs.reverse() : segs));   // a right-to-left run is laid out from its right end
    }
    const tot = parts.reduce((t, p) => t + p.w, 0);
    let x = base === 'R' ? x0 + w - tot : x0;
    for (const p of parts) { if (p.sym) drawSym(p.sym, x, yy, sz, p.w); else if (p.s) pg.drawText(p.s, { x, y: yy, size: sz, font: f, color: ink }); x += p.w; }
  };
  const para = (text, { f = reg, sz = 12, lh = 1.5, before = 0, after = 7, mark = null } = {}) => {
    const dir = paraDir(text, docDir), ind = mark ? 20 : 0;
    const lines = wrapLines(norm(text), f, sz, maxW - ind);
    y -= before;
    lines.forEach((ln, j) => {
      const L = sz * lh;
      if (y - L < M) newPage();
      const yy = y - f.heightAtSize(sz, { descender: false }) - (L - sz) / 2;
      lineAt(ln, dir === 'R' ? M : M + ind, maxW - ind, yy, f, sz, dir);
      if (j === 0 && mark) lineAt(mark, dir === 'R' ? W - M - ind + 4 : M, ind - 4, yy, f, sz, dir);
      y -= L;
    });
    y -= after;
  };
  const table = rows => {
    const n = Math.max(...rows.map(r => r.length)), cw = maxW / n, sz = 10, L = sz * 1.4, pad = 5;
    const tDir = paraDir(rows.flat().join(' '), docDir);
    y -= 4;
    rows.forEach((r, ri) => {
      const f = ri === 0 ? bold : reg;
      const cells = Array.from({ length: n }, (_, k) => wrapLines(String(r[k] ?? ''), f, sz, cw - 2 * pad));
      const rh = Math.max(1, ...cells.map(c => c.length)) * L + 2 * pad;
      if (y - rh < M) newPage();
      cells.forEach((c, k) => {
        const col = tDir === 'R' ? n - 1 - k : k, x0 = M + col * cw;
        pg.drawRectangle({ x: x0, y: y - rh, width: cw, height: rh, borderColor: rule, borderWidth: .6 });
        c.forEach((ln, j) => lineAt(ln, x0 + pad, cw - 2 * pad, y - pad - f.heightAtSize(sz, { descender: false }) - j * L - (L - sz) / 2, f, sz, paraDir(ln, tDir)));
      });
      y -= rh;
    });
    y -= 10;
  };
  pages.forEach(blocks => {
    newPage();
    let k = 0;
    for (const b of blocks) {
      if (b.t !== 'oli') k = 0;
      if (b.t === 'h1') para(b.text, { f: bold, sz: 18, lh: 1.35, before: 4, after: 10 });
      else if (b.t === 'h2') para(b.text, { f: bold, sz: 14, lh: 1.4, before: 6, after: 6 });
      else if (b.t === 'li') para(b.text, { mark: '•', after: 3 });
      else if (b.t === 'oli') para(b.text, { mark: `${++k}.`, after: 3 });
      else if (b.t === 'table') table(b.rows);
      else para(b.text);
    }
  });
  return doc.save();
}
window.SIMAN_SCAN_EXPORT = { blocksDocx, blocksPdf, parseBlocks, blocksHtml, domBlocks };

/* =====================================================================
   6. OCR IN MORE LANGUAGES
   ===================================================================== */
const OCR_SETS = { he: ['heb', 'eng'], ru: ['rus', 'eng'], ar: ['ara', 'eng'], fr: ['fra', 'eng'], es: ['spa', 'eng'], heru: ['heb', 'rus', 'eng'] };
let ocrPref = 'he';
try { ocrPref = localStorage.getItem('siman.ocrlang') || 'he'; if (!OCR_SETS[ocrPref]) ocrPref = 'he'; } catch (_) {}
getTess = async function () {
  if (!window.Tesseract) await loadScript(TESS_URL);
  if (tessWorker && tessWorker._siman !== ocrPref) { try { await tessWorker.terminate(); } catch (_) {} tessWorker = null; }
  if (!tessWorker) { tessWorker = await Tesseract.createWorker(OCR_SETS[ocrPref], 1, { logger: m => tessLog && tessLog(m) }); tessWorker._siman = ocrPref; }
  return tessWorker;
};
const ocrSel = () => h(`<select class="ocr-lang" aria-label="שפת הטקסט הסרוק" title="שפת הטקסט הסרוק"><option value="he">עברית + אנגלית</option><option value="ru">רוסית + אנגלית</option><option value="heru">עברית + רוסית + אנגלית</option><option value="ar">ערבית + אנגלית</option><option value="fr">צרפתית + אנגלית</option><option value="es">ספרדית + אנגלית</option></select>`);
const sels = [ocrSel(), ocrSel()];
$('ocrBtn').after(sels[0]); $('ocrAllBtn').after(sels[1]);
sels.forEach(s => { s.value = ocrPref; s.onchange = () => { ocrPref = s.value; sels.forEach(x => x.value = ocrPref); try { localStorage.setItem('siman.ocrlang', ocrPref); } catch (_) {} }; });
// the small one in the options bar shows together with the OCR button
const syncOcr = () => sels[0].classList.toggle('hidden', $('ocrBtn').classList.contains('hidden'));
new MutationObserver(syncOcr).observe($('ocrBtn'), { attributes: true, attributeFilter: ['class'] }); syncOcr();

/* =====================================================================
   7. ENTRY POINTS — command bar, first screen
   ===================================================================== */
const more = [
  { l: 'שאלה על המסמך', k: 'שאל שאלה שאלות תשאל עוזר בינה מלאכותית ai צאט צ׳אט chat ask question', doc: 1, run: showAI },
  { l: 'סיכום המסמך', k: 'סכם סיכום תקציר תמצית קצר נקודות עיקריות summary summarize', doc: 1, run: () => askQ(CHIPS[0][1], CHIPS[0][0]) },
  { l: 'הנקודות העיקריות במסמך', k: 'נקודות עיקריות עיקר חשוב רשימה key points', doc: 1, run: () => askQ(CHIPS[1][1], CHIPS[1][0]) },
  { l: 'הקראה בקול', k: 'הקרא הקראה קרא תקריא בקול שמע שמיעה דיבור קול read aloud speak tts', doc: 1, run: () => RA.start(currentPageIndex()) },
  { l: 'כתב יד או סריקה ל-Word / PDF', k: 'כתב יד כתוב ביד סריקה סרוק צילום מצלמה תמונה הקלדה מודפס הפוך המר וורד word pdf handwriting scan photo camera', run: () => SC.open() },
  { l: 'תרגום לרוסית', k: 'תרגם תרגום רוסית רוסי translate russian русский', doc: 1, run: () => trQuick('ru') },
  { l: 'תרגום לערבית', k: 'תרגם תרגום ערבית ערבי translate arabic عربي', doc: 1, run: () => trQuick('ar') },
  { l: 'תרגום לצרפתית', k: 'תרגם תרגום צרפתית צרפתי translate french français', doc: 1, run: () => trQuick('fr') },
  { l: 'תרגום לספרדית', k: 'תרגם תרגום ספרדית ספרדי translate spanish español', doc: 1, run: () => trQuick('es') },
  { l: 'פתיחת כמה קבצים יחד', k: 'כמה קבצים מספר קבצים במקביל לשוניות טאבים יחד פתח multiple tabs', run: () => pickFile() },
  { l: 'איחוד כל הקבצים הפתוחים', k: 'אחד איחוד מזג מיזוג חבר כל הקבצים הפתוחים merge all tabs', doc: 1, run: () => TB.list.length > 1 ? mergeAll() : toast('פתחו עוד קובץ כדי לאחד') },
  { l: 'שמירת כל הקבצים הפתוחים', k: 'שמור שמירה כל הקבצים zip כולם save all', doc: 1, run: () => zipAll() },
];
more.forEach(c => { c.hay = normHe(c.l + ' ' + c.k); c.toks = c.hay.split(/\s+/).filter(t => t.length >= 3); });
CMDS.splice(2, 0, ...more);

// first screen: two more quick actions + the updated list
const quick = document.querySelector('.wl-quick');
if (quick) {
  const scanQ = h(`<button class="qa" type="button">${ICON.scan}כתב יד ל-Word</button>`);
  const askQa = h(`<button class="qa" type="button">${ICON.spark}לשאול את המסמך</button>`);
  scanQ.onclick = () => SC.open();
  askQa.onclick = () => { if (typeof needReg === 'function' && needReg()) { showReg(); return; } afterOpen = showAI; $('openBtn2').click(); };
  const ghost = quick.querySelector('.qa.ghost');
  quick.insertBefore(scanQ, ghost); quick.insertBefore(askQa, ghost);
}
document.querySelectorAll('.wl-list li').forEach(li => { if (/תרגום לעברית ולאנגלית/.test(li.textContent)) li.lastChild.textContent = 'תרגום ל-6 שפות, כולל רוסית וערבית'; });
const wl = document.querySelector('.wl-list');
if (wl) {
  wl.append(h(`<li>${ICON.scan}כתב יד וסריקות למסמך מודפס</li>`));
  wl.append(h(`<li>${ICON.spark}שאלות, סיכום והקראה של המסמך</li>`));
}
document.addEventListener('keydown', e => { if (e.key === 'Escape' && RA.on && !document.querySelector('dialog[open]')) RA.stop(); });
})();