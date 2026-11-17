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
  crop: '<svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 2v14a2 2 0 002 2h14"/><path d="M18 22V8a2 2 0 00-2-2H2"/></svg>',
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
.tab .aiw{color:#E8A400;font-size:12px;line-height:1;animation:aiwP 1.2s ease-in-out infinite}
.tab .aiw.ready{color:#1E9E5A;animation:none}
@keyframes aiwP{50%{opacity:.35}}
@media (prefers-reduced-motion:reduce){.tab .aiw{animation:none}}
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
.ai-multi{display:flex;gap:9px;align-items:flex-start;border:1px dashed var(--blue);border-radius:12px;padding:8px 10px;cursor:pointer;font-size:13.5px}
.ai-multi:has(input:checked){border-style:solid;background:color-mix(in srgb,var(--blue) 8%,transparent)}
.ai-multi input{margin-top:3px;width:16px;height:16px;flex:none}
.ai-multi small{display:block;color:var(--muted);font-size:12px}
.ai-chip.follow{border-style:dashed;border-color:var(--blue);color:var(--blue);text-align:start}
.ai-chip.follow::before{content:"↳ ";opacity:.7}
.ai-chips.study{align-items:center;padding-top:6px;border-top:1px dashed var(--line)}
.ai-st-l{font-size:12.5px;color:var(--muted);margin-inline-end:2px}
.ai-export{align-self:flex-start;border:0;background:none;color:var(--muted);font:inherit;font-size:12.5px;text-decoration:underline;cursor:pointer;padding:0}
.ai-study{display:grid;gap:8px}
.ai-st-h{font-size:12.5px;color:var(--muted);font-weight:600}
.ai-st-h b{color:var(--ink)}
.ai-fc{display:grid;gap:6px;text-align:start;font:inherit;color:var(--ink);background:var(--bar);border:1px solid var(--line);border-radius:12px;padding:10px 12px;cursor:pointer;box-shadow:0 2px 0 var(--line)}
.ai-fc .q{font-weight:700}
.ai-fc .an{display:none;border-top:1px dashed var(--line);padding-top:6px}
.ai-fc[aria-expanded=true] .an{display:block}
.ai-fc[aria-expanded=true]{border-color:var(--ink)}
.ai-qz{border:1px solid var(--line);border-radius:12px;padding:10px 12px;display:grid;gap:6px;background:var(--bar)}
.ai-qz .q{margin:0;font-weight:700}
.ai-qz button{text-align:start;font:inherit;font-size:14px;color:var(--ink);background:none;border:1px solid var(--line);border-radius:9px;padding:7px 10px;cursor:pointer}
.ai-qz button:not(:disabled):hover{border-color:var(--ink)}
.ai-qz button.right{background:color-mix(in srgb,#1F9D55 18%,transparent);border-color:#1F9D55;font-weight:700}
.ai-qz button.wrong{background:color-mix(in srgb,#C2362B 15%,transparent);border-color:#C2362B;text-decoration:line-through}
.ai-qz button:disabled{cursor:default;opacity:1}
.ai-qz .why{margin:0;font-size:13.5px;color:var(--muted)}
.ck-photo{width:100%;justify-content:center;margin-top:8px;white-space:normal;height:auto;min-height:40px;padding:8px 12px}
.ckd .ck-sec{margin:0 0 22px}
.ckd .sc-trayhead small{font-weight:400;color:var(--muted);font-size:14px}
.ckd .ck-txt{width:100%;min-height:70px;resize:vertical;font:inherit;font-size:14.5px;border:1px solid var(--line);border-radius:12px;padding:9px 12px;background:var(--bar);color:var(--ink);direction:rtl;unicode-bidi:plaintext;box-sizing:border-box}
.ckd .ck-opts{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:10px 14px;align-items:end}
.ckd .ck-opts label{display:grid;gap:4px;font-size:13px;color:var(--muted)}
.ckd .ck-opts label.ck{display:flex;align-items:center;gap:7px;color:var(--ink);font-size:14px}
.ckd .ck-opts input[type=text],.ckd .ck-opts select{font:inherit;font-size:14px;color:var(--ink);background:var(--bar);border:1px solid var(--line);border-radius:10px;padding:8px 10px}
.ck-class{display:flex;align-items:center;gap:10px;flex-wrap:wrap;border:1px dashed var(--line);border-radius:12px;padding:10px 12px;margin:0 0 18px;background:color-mix(in srgb,var(--marker) 22%,transparent)}
.ck-class span{flex:1;min-width:180px;font-size:14px}
.grd .gr-steps{margin:0 0 14px;color:var(--muted);font-size:13.5px}.grd .gr-steps b{color:var(--ink)}
.grd .gr-help{margin:4px 0 10px;color:var(--muted);font-size:13.5px;line-height:1.5}
.grd .gr-nopic{height:150px;display:grid;place-items:center;background:var(--bar);border-radius:8px;margin-bottom:6px;color:#1F9D55;font-weight:700}
.grd .gr-tbl{display:grid;gap:6px;margin:0 0 10px}
.grd .gr-row{display:grid;grid-template-columns:64px 76px 1fr 34px;gap:6px;align-items:start}
.grd .gr-row.gr-h span{font-size:12.5px;color:var(--muted);font-weight:600}
.grd .gr-row input,.grd .gr-row textarea{font:inherit;font-size:14px;color:var(--ink);background:var(--bar);border:1px solid var(--line);border-radius:9px;padding:6px 8px;min-width:0;box-sizing:border-box;width:100%}
.grd .gr-row textarea{resize:vertical;unicode-bidi:plaintext}
.grd .gr-del{border:0;background:none;color:var(--muted);font-size:20px;cursor:pointer;padding:4px;line-height:1}
.grd .gr-sum{font-weight:700;margin-inline-start:8px}
.grd .gr-cur{border:2px solid var(--ink);border-radius:14px;padding:12px;margin:0 0 18px}
.grd .gr-cur h3,.grd .gr-list h3{margin:0 0 8px;font-size:18px}
.grd .gr-name{display:grid;gap:4px;font-size:13px;color:var(--muted);max-width:420px}
.grd .gr-name input{font:inherit;font-size:16px;color:var(--ink);background:var(--bar);border:1px solid var(--line);border-radius:10px;padding:8px 10px}
.grd .gr-list ol{margin:0;padding:0;list-style:none;display:grid;gap:6px}
.grd .gr-list li{border:1px solid var(--line);border-radius:12px;padding:8px 10px;background:var(--bar)}
.grd .gr-li{display:flex;align-items:center;gap:10px}
.grd .gr-nm{flex:1;text-align:start;background:none;border:0;font:inherit;font-weight:700;color:var(--ink);cursor:pointer;padding:4px 0}
.grd .gr-nm:disabled{cursor:default;color:var(--ink)}
.grd .gr-g{font-size:20px;min-width:44px;text-align:center}
.grd .gr-st{font-size:13px;color:var(--muted)}.grd .gr-st.bad,.grd .gr-unc{color:#C2362B}
.grd .gr-unc{font-size:13px;margin:6px 0 0}
.grd .gr-det{display:none;margin-top:8px}.grd li.open .gr-det{display:block}
.grd .gr-det p{margin:0 0 6px;font-size:14px}
.grd table{border-collapse:collapse;width:100%;font-size:13.5px}
.grd th,.grd td{border:1px solid var(--line);padding:5px 7px;text-align:start;vertical-align:top}
.grd th{background:color-mix(in srgb,var(--marker) 30%,transparent);font-weight:700}
.grd td input{width:56px;font:inherit;font-size:13.5px;border:1px solid var(--line);border-radius:6px;padding:2px 4px;background:var(--bar);color:var(--ink)}
.grd .gr-sumbox{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin:0 0 14px}
.grd .gr-sumbox div{border:1px solid var(--line);border-radius:12px;padding:10px;text-align:center;display:grid;gap:2px}
.grd .gr-sumbox b{font-size:24px}.grd .gr-sumbox span{font-size:12.5px;color:var(--muted)}
.grd .gr-final{display:block;overflow:auto;max-width:100%}
.grd .gr-final th small{color:var(--muted);font-weight:400}
.grd .gr-final tr.avg td{font-weight:700;background:var(--bar)}
@media (max-width:560px){.grd .gr-row{grid-template-columns:52px 64px 1fr 28px}.grd .gr-sumbox b{font-size:19px}}
.ai-ck-form{border:1px solid var(--line);border-radius:14px;padding:12px;display:grid;gap:9px;background:var(--bar)}
.ai-ck-form h4{margin:0;font-size:15px}
.ai-ck-form p{margin:0;font-size:13px;color:var(--muted)}
.ai-ck-form label{display:grid;gap:4px;font-size:13px;color:var(--muted)}
.ai-ck-form label.ck{display:flex;align-items:center;gap:7px;color:var(--ink)}
.ai-ck-form select,.ai-ck-form textarea{font:inherit;font-size:14px;color:var(--ink);background:var(--paper,#fff);border:1px solid var(--line);border-radius:10px;padding:7px 9px}
.ai-ck-form textarea{min-height:80px;resize:vertical;direction:rtl;unicode-bidi:plaintext}
.ai-ck-form .acts{display:flex;gap:8px;justify-content:flex-end}
.ai-ck-sum{border-radius:12px;padding:10px 12px;background:color-mix(in srgb,var(--marker) 35%,transparent);font-weight:600;margin:0}
.ai-ck-sec{display:grid;gap:5px}
.ai-ck-sec h5{margin:4px 0 0;font-size:13.5px}
.ai-ck-sec p{margin:0;padding:7px 10px;border-radius:10px;border:1px solid var(--line);background:var(--bar);font-size:14px}
.ai-ck-sec.good p{border-inline-start:4px solid #1F9D55}
.ai-ck-sec.fix p{border-inline-start:4px solid #C2362B}
.ai-ck-sec.tip p{border-inline-start:4px solid #D99A00}
.ai-ck-grade{display:flex;align-items:center;gap:10px;border:2px solid var(--ink);border-radius:12px;padding:8px 12px}
.ai-ck-grade b{font-size:24px;line-height:1}
.ai-ck-grade span{font-size:13.5px;color:var(--muted)}
.ai-map details{margin-inline-start:4px}
.ai-map summary{cursor:pointer;font-weight:700;padding:3px 0}
.ai-map .kids{border-inline-start:2px solid color-mix(in srgb,var(--marker) 70%,var(--line));margin-inline-start:6px;padding-inline-start:12px;display:grid;gap:2px}
.ai-map .leaf{padding:2px 0;position:relative}
.ai-map .leaf::before{content:"";position:absolute;inset-inline-start:-12px;top:50%;width:9px;border-top:2px solid color-mix(in srgb,var(--marker) 70%,var(--line))}
.ai-map>details>summary{font-size:15.5px;background:color-mix(in srgb,var(--marker) 45%,transparent);border-radius:8px;padding:4px 8px;display:inline-block}
.page .ai-src{position:absolute;z-index:3;pointer-events:none;border-radius:4px;background:rgba(255,212,59,.42);box-shadow:0 0 0 2px rgba(224,161,0,.9),0 0 14px rgba(255,212,59,.8);animation:aiSrc .5s ease;transition:opacity 1.2s}
.page .ai-src.out{opacity:0}
@keyframes aiSrc{from{transform:scale(1.2)}}
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
.sc-tray{list-style:none;margin:0 0 14px;padding:0;display:grid;grid-template-columns:repeat(auto-fill,minmax(120px,1fr));gap:12px}
.sc-tray li{position:relative;border:1px solid var(--line);border-radius:12px;padding:6px;font-size:12.5px;color:var(--muted);text-align:center;background:var(--bar)}
.sc-tray img{display:block;width:100%;height:150px;object-fit:contain;background:#fff;border-radius:8px;margin-bottom:6px}
.sc-tray .fr{position:absolute;bottom:34px;inset-inline-start:10px;display:inline-flex;align-items:center;gap:4px;height:30px;padding:0 10px;border-radius:999px;border:0;background:rgba(20,24,32,.78);color:#fff;font:inherit;font-size:12.5px;cursor:pointer}
.sc-tray .fr svg{width:15px;height:15px}
.sc-frame{display:flex;flex-direction:column;align-items:center;gap:10px}
.sc-fhelp{margin:0;color:var(--muted);text-align:center}
.sc-fwrap{position:relative;display:inline-block;max-width:100%;touch-action:none;user-select:none;-webkit-user-select:none}
.sc-fwrap img{display:block;max-width:100%;max-height:62vh;border-radius:6px;pointer-events:none}
.sc-fwrap svg{position:absolute;inset:0;width:100%;height:100%;pointer-events:none}
.sc-fwrap .shade{fill:rgba(10,14,22,.55)}
.sc-fwrap .edge{fill:none;stroke:#FFD93D;stroke-width:.6;vector-effect:non-scaling-stroke;stroke-width:3}
.sc-hd{position:absolute;width:34px;height:34px;margin:-17px 0 0 -17px;border-radius:50%;border:3px solid #FFD93D;background:rgba(255,217,61,.25);cursor:grab;touch-action:none;padding:0}
.sc-hd::after{content:"";position:absolute;inset:-12px;border-radius:50%}
.sc-hd.on,.sc-hd:focus-visible{background:rgba(255,217,61,.6);outline:2px solid #fff}
.sc-tray .x{position:absolute;top:10px;inset-inline-end:10px;width:30px;height:30px;border-radius:50%;border:0;background:rgba(20,24,32,.78);color:#fff;font-size:18px;line-height:1;cursor:pointer}
.sc-tray li.add{display:flex;flex-direction:column;align-items:center;justify-content:center;gap:6px;min-height:178px;border-style:dashed;cursor:pointer;color:var(--ink);font:inherit;font-weight:600}
.sc-tray li.add svg{width:30px;height:30px}
.sc-trayhead{display:flex;align-items:baseline;gap:10px;flex-wrap:wrap;margin:0 0 12px}
.sc-trayhead h3{margin:0;font-size:20px}
.sc-trayhead p{margin:0;color:var(--muted)}
.sc-pg .sc-del{border:0;background:none;color:#B42318;font:inherit;text-decoration:underline;cursor:pointer;padding:0}
.sc-saved{flex-basis:100%;display:flex;flex-direction:column;gap:4px;background:#E7F5EC;border:1px solid #9BD3B1;border-radius:12px;padding:10px 12px;color:#14532D}
.sc-saved span{font-size:13px;color:#2F5D43;word-break:break-all}
.sc-saved-a{display:flex;gap:8px;flex-wrap:wrap;margin-top:4px}
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
/* 6.3.2: an editor takes a moment to open (it builds the page, then waits a frame to fit the zoom). "בית" or
   "+ חדש" pressed in that moment closed it half-built: it crashed, and its tab stayed on the home screen as a
   "back" that did nothing. Now each editor's open / close wait for an open that is still running. */
const ED_BUSY = {};
for (const k of Object.keys(ED)) {
  const name = ED[k].api;
  const wrap = A => {
    if (!A || A.__simanSafe) return A;
    const o = A.open, c = A.close;
    A.open = async function (opts) {
      const prev = ED_BUSY[k]; if (prev) { try { await prev; } catch (_) {} }
      const p = o.call(this, opts); ED_BUSY[k] = p;
      try { return await p; } finally { if (ED_BUSY[k] === p) ED_BUSY[k] = null; }
    };
    A.close = async function (...a) { const p = ED_BUSY[k]; if (p) { try { await p; } catch (_) {} } return c.apply(this, a); };
    A.__simanSafe = true; return A;
  };
  let val = window[name];
  if (val) window[name] = wrap(val);
  else Object.defineProperty(window, name, { configurable: true, enumerable: true, get: () => val, set: v => { val = wrap(v); } });
}
const edSettled = async k => { const p = ED_BUSY[k]; if (p) { try { await p; } catch (_) {} } };
const edKindOfId = id => Object.keys(ED).find(k => ED[k].id === id);
const edApi = k => window[ED[k].api];
const edState = k => { const A = edApi(k); return A && A._test ? A._test[ED[k].st] : null; };
const edLive = t => !!(t && t.ed && t.liveEl && document.body.contains(t.liveEl));
// 6.3.4: never throws — one odd tab (a file kept in an unusual state) froze the tab bar and the home card
function tabInfo(t) {
  try { return tabInfo1(t); } catch (e) { console.error('tabInfo', e); return { name: (t && t.data && t.data.name) || 'מסמך', dirty: !!(t && t.data && t.data.dirty), kind: t && t.ed ? ED[t.ed].kind : 'PDF', edit: !!(t && t.ed) }; }
}
function tabInfo1(t) {
  if (t.ed) {
    const live = edLive(t), st = live && edState(t.ed);
    const inp = live && t.liveEl.querySelector(ED[t.ed].name);
    const base = ((inp && inp.value) || (t.data && t.data.name) || 'מסמך').replace(/\.(docx|xlsx|pptx)$/i, '');
    return { name: base + '.' + t.ed, dirty: live ? !!(st && (st.dirty || st.snapT)) : !!(t.data && t.data.dirty), kind: ED[t.ed].kind, edit: true };
  }
  const d = tabData(t);
  return { name: (d.office ? d.office.name : d.name) || 'מסמך', dirty: !!d.dirty, kind: d.office ? String(d.office.kind || 'office').toUpperCase() : 'PDF' };
}
const tabName = t => tabInfo(t).name;
const deskTabs = () => TB.list.filter(t => !t.ed);
async function park(t) {
  await edSettled(t.ed);
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
async function enter(t) { if (t.ed) { await unpark(t); await new Promise(r => setTimeout(r, 0)); aiFlush(t); } else restore(t); renderTabs(); AS.onTab(); }

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
  switching = true; switchingAt = Date.now();
  try { await leave(); await enter(t); } finally { switching = false; renderTabs(); }
}
// o.force: the question about unsaved changes was already asked (the home card asks inside the app)
async function closeTab(t, o = {}) {
  if (switching && Date.now() - switchingAt > 15000) switching = false;
  if (isBusy() || switching) return false;
  const info = tabInfo(t);
  if (info.dirty && !o.force && !confirm(`ב"${info.name}" יש שינויים שלא נשמרו. לסגור בכל זאת?`)) return false;
  const i = TB.list.indexOf(t); if (i < 0) { renderTabs(); return true; }
  TB.list.splice(i, 1);
  if (homeFrom === t) homeFrom = null;
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
  return true;
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
  // (not the "בית" / "+ חדש" buttons of siman-home: their class carries their action)
  const peer = [...close.parentElement.querySelectorAll('button')].find(b => b !== close && !b.classList.contains('primary') && !/\bsh-ed/.test(b.className));
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
      const b = h(`<div class="tab${d.edit ? ' ed ' + t.ed : ''}" role="tab" tabindex="${on ? 0 : -1}" aria-selected="${on}" draggable="true"><span class="kd">${d.kind}</span><span class="nm"></span>${AS.busySet.has(t.id) ? '<span class="aiw" title="העוזר עובד על הקובץ הזה">✦</span>' : t.aiQueue && t.aiQueue.some(x => x.plan && x.plan.actions.length) ? '<span class="aiw ready" title="העוזר סיים. השינויים ייכנסו כשתחזרו לקובץ">✦</span>' : ''}${d.dirty ? '<span class="dot" title="יש שינויים שלא נשמרו"></span>' : ''}<span class="x" role="button" aria-label="סגירת הקובץ" title="סגירה">×</span></div>`);
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
/* 6.3: the home screen, without closing anything. The file in front is put aside exactly like when switching tabs
   (its pages, place on the page, marks, undo, translation and unsaved changes stay with its tab; an editor is
   kept as its file), and the home screen shows. "Back" returns to the same file at the same place. */
let homeFrom = null;
async function goHome() {
  if (!TB.cur) return false;
  if (isBusy() || TR.running) { toast('רגע, פעולה עדיין רצה'); return false; }
  const t = TB.cur;
  await leave(); toEmpty(); homeFrom = t; renderTabs();
  return true;
}
// 6.3.2: an editor tab that is neither open nor kept aside (its editor closed, or setting it aside failed) cannot
// come back: it is taken off the list instead of staying on the home screen as a "back" that does nothing
const deadTab = t => !!(t && t.ed && !edLive(t) && !t.data && TB.opening !== t && !t.parking);
function dropDead() {
  const dead = TB.list.filter(deadTab); if (!dead.length) return 0;
  for (const t of dead) { TB.list.splice(TB.list.indexOf(t), 1); AS.forget(t); if (TB.cur === t) TB.cur = null; if (homeFrom === t) homeFrom = null; }
  renderTabs(); return dead.length;
}
let switchingAt = 0;
async function backFromHome(t) {
  if (dropDead() && t && !TB.list.includes(t)) { toast('הקובץ הזה כבר לא פתוח'); return false; }
  t = t || homeFrom; if (!t || !TB.list.includes(t)) t = TB.list[TB.list.length - 1];
  if (!t) return false;
  homeFrom = null;
  if (switching && Date.now() - switchingAt > 15000) switching = false;     // a switch that never finished
  // the home screen shows, but this tab is still marked as the one in front: bring it up as it is
  if (t === TB.cur) {
    if (t.ed && edLive(t)) { renderTabs(); return true; }
    if (!t.ed && S.doc) { applyUI(); renderTabs(); AS.onTab(); return true; }
    TB.cur = null;
  }
  await switchTo(t);
  return TB.cur === t;
}
window.SIMAN_TABS = { TB, switchTo, closeTab, mergeAll, zipAll, openMany, tabInfo, goHome, backFromHome, dropDead, homeFrom: () => (homeFrom && TB.list.includes(homeFrom) ? homeFrom : null) };

/* =====================================================================
   2. AI — the smart server (siman-ai)
   ===================================================================== */
// 6.6.2: answers already given in this session (see askQ), keyed by a short fingerprint of the document text
const ASK_CACHE = new Map();
const textSig = t => { let h = 2166136261; for (let i = 0; i < t.length; i++) { h ^= t.charCodeAt(i); h = Math.imul(h, 16777619); } return t.length + ':' + (h >>> 0).toString(36); };
async function ai(payload, { onWait, signal, timeout = 300000, once = false } = {}) {
  // 6.2: the smart services run on our server: they ask for an email first (the tools on the device do not)
  if (window.cloudGate && !cloudGate('ai')) throw new Error(window.CLOUD_REG_MSG || 'צריך להירשם עם מייל');
  // 6.6.1: once = a quick helper request (the page edges): one try, no waiting, it is only a help
  const tries = once ? 1 : 4;
  let bgRetries = 0;
  for (let attempt = 0; attempt < tries; attempt++) {
    // 5.9.6: with no internet at all there is nothing to wait for (before: three retries, 30 seconds of "working")
    if (navigator.onLine === false) throw new Error('אין חיבור לאינטרנט');
    let r, d, status;
    const ctl = new AbortController(), timer = setTimeout(() => ctl.abort(), timeout);
    if (signal) signal.addEventListener('abort', () => ctl.abort(), { once: true });
    // 6.6.4: a phone stops the network of a page whose screen went off or that went to the background
    // (a page of handwriting takes a minute or two). Such a drop is not "no internet": the request waits
    // until the app is back on screen and is sent again, without counting as a failed try.
    let hid = document.hidden; const onVis = () => { if (document.hidden) hid = true; };
    document.addEventListener('visibilitychange', onVis);
    try {
      r = await fetch(AI, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(payload), signal: ctl.signal });
      // the server keeps the line open with spaces while Gemini works; the JSON comes at the end
      const raw = await r.text();
      try { d = JSON.parse(raw.trim() || '{}'); } catch (_) { d = {}; }
    } catch (e) {
      clearTimeout(timer);
      document.removeEventListener('visibilitychange', onVis);
      if (signal && signal.aborted) throw e;
      if (hid && bgRetries < 3 && !once) {
        bgRetries++; attempt--;
        if (document.hidden) await new Promise(res => { const f = () => { if (!document.hidden) { document.removeEventListener('visibilitychange', f); res(); } }; document.addEventListener('visibilitychange', f); });
        if (signal && signal.aborted) throw new DOMException('stop', 'AbortError');
        onWait && onWait(1); await sleep(800);
        continue;
      }
      if (e && e.name === 'AbortError') throw new Error('הבקשה לקחה יותר מדי זמן. נסו שוב, או צלמו את הדף קרוב יותר ובאור טוב');
      // a phone's connection drops for a moment (switching between Wi-Fi and mobile data, a long wait for a page):
      // the request is sent again a few times before giving up
      if (attempt < tries - 1) {
        for (let k = 5 * (attempt + 1); k > 0; k--) { onWait && onWait(k); await sleep(1000); if (signal && signal.aborted) throw new DOMException('stop', 'AbortError'); }
        continue;
      }
      // 6.6.4: the internet is there but the line to the service broke: said so (before: "אין חיבור", which also
      // stopped every other page as if the phone were offline)
      throw new Error(navigator.onLine === false ? 'אין חיבור לאינטרנט' : 'החיבור לשירות החכם נקטע באמצע. נסו שוב, ועדיף להשאיר את המסך דלוק עד שהקריאה מסתיימת');
    }
    clearTimeout(timer); document.removeEventListener('visibilitychange', onVis);
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
    // 6.6.1: the server's own words for its daily share (they also say a personal key lifts the limit)
    if (d.error && /^(DAY|ALL)$/.test(d.error.code || '') && d.error.message) throw new Error(d.error.message);
    if (r.status === 429 && /per ?day|PerDay/i.test(all)) throw new Error('נגמרה המכסה היומית החינמית. היא מתאפסת מחר בבוקר');
    if (r.status === 403) throw new Error('השרת החכם חוסם את הכתובת הזו של האפליקציה');
    if (r.status === 404) throw new Error('השרת החכם (siman-ai) עוד לא הוקם');
    // 6.3.1: one more try at once: the server now remembers a model that hung and starts with the other one
    if ((r.status === 504 || r.status === 524) && attempt === 0 && !once) { onWait && onWait(1); continue; }
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
    <button class="btn" id="aiScan" type="button">${ICON.scan}<span>מדף למסמך</span></button>
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
// 6.0: study tools (shown under the questions, on any PDF or document)
const STUDY_CHIPS = [['🗂 כרטיסיות לחזרה', 'cards'], ['✅ בוחן אמריקאי', 'quiz'], ['🧠 מפת חשיבה', 'map'], ['📝 בדיקת עבודה', 'check']];
const FOLLOW_TAIL = '\n\n(After the answer, add one last line that starts with "❓" and holds 3 short follow-up questions the user may want to ask next about this document, separated by " | ", in the language of the question.)';
function splitFollow(a) {
  const m = /^\s*❓\s*(.+)$/m.exec(a || '');
  if (!m) return { a, follow: [] };
  return { a: a.replace(m[0], '').trim(), follow: m[1].split(/\s*\|\s*/).map(x => x.replace(/^[\d.)\s-]+/, '').trim()).filter(x => x && x.length < 140).slice(0, 3) };
}
const ED_CHIPS = {
  docx: [['כותרת לקובץ', 'הכן לי כותרת מתאימה לקובץ, בתחילת המסמך.'], ['עיצוב הכותרות', 'עצב את כל הכותרות במסמך בצורה אחידה ומקצועית: מודגש, צבע כחול כהה וגודל מתאים.'], ['סידור המסמך', 'סדר את עיצוב המסמך: יישור לשני הצדדים, רווח שורה 1.5 ורווח קטן אחרי כל פסקה.'], ['טבלה מעוצבת', 'עצב את הטבלאות במסמך: שורת כותרת צבועה עם טקסט לבן מודגש ופסים בשורות.'], ['מספרי עמודים', 'הוסף מספרי עמודים במרכז הכותרת התחתונה, בתבנית עמוד X מתוך Y.'], ['סיכום קצר', 'סכם את המסמך בקצרה, ב-3 עד 5 משפטים. אל תשנה את הקובץ.']],
  xlsx: [['כותרת לטבלה', 'עצב את השורה הראשונה ככותרת: מודגש, רקע כחול כהה, טקסט לבן וממורכז, והקפא אותה.'], ['טבלה מעוצבת', 'הפוך את הנתונים לטבלה מעוצבת: שורת כותרת צבועה, גבולות ופסים, והתאם את רוחב העמודות.'], ['שורת סיכום', 'הוסף מתחת לנתונים שורת סה"כ עם סכום לכל עמודה מספרית, מודגשת.'], ['עיצוב מספרים', 'עצב את העמודות המספריות: מחירים וסכומים כשקלים, כמויות עם מפריד אלפים.'], ['מה בגיליון?', 'הסבר בקצרה מה יש בגיליון ומה הנתונים החשובים. אל תשנה את הקובץ.']],
  pptx: [['שקף פתיחה', 'הכן שקף פתיחה בתחילת המצגת, עם כותרת וכותרת משנה מתאימות לתוכן.'], ['שקף סיכום', 'הוסף בסוף המצגת שקף סיכום עם 3 עד 5 נקודות עיקריות מהמצגת.'], ['עיצוב אחיד', 'עצב את המצגת בצורה אחידה: כל הכותרות מודגשות באותו גופן ובאותו צבע.'], ['רקע לשקפים', 'שנה את הרקע של כל השקפים לכחול כהה, ואת כל הטקסט ללבן.'], ['תוכן עניינים', 'הוסף אחרי השקף הראשון שקף תוכן עניינים עם הכותרות של השקפים.']],
};
const edKind = () => onEditor() ? TB.cur.ed : null;
/* 6.1: one question about all the PDF files open in tabs (compare two contracts, find a detail in one of five files) */
const pdfTabs = () => TB.list.filter(t => !t.ed && tabData(t).doc && tabData(t).pages && tabData(t).pages.length);
const MULTI = { on: false };
const MULTI_TAIL = '\n\n(Several documents are given, each starting with a line like "##### קובץ 2: name #####". Use all of them; when they differ, say how. Cite every fact as [קובץ 2, עמ\' 3] — the file number and the page — instead of a page alone.)';
async function pagesText(pages, cap) {
  const parts = []; let total = 0;
  for (let i = 0; i < pages.length && total < cap; i++) {
    let t = '';
    try {
      const P = pages[i], lines = await pageLines(P);
      if (lines) { const out = []; let prev = null; for (const L of lines) { if (prev && prev.y - L.y > Math.max(prev.h, L.h) * 1.9) out.push(''); out.push(L.text); prev = L; } t = out.join('\n'); }
      else if (P.tilted) { for (const it of P.tilted) { t += it.str; if (it.hasEOL) t += '\n'; } }
    } catch (_) {}
    const part = `===== עמוד ${i + 1} =====\n${t.trim()}`; parts.push(part); total += part.length;
  }
  return parts.join('\n\n');
}
async function allText() {
  const tabs = pdfTabs(), cap = Math.floor(380000 / Math.max(1, tabs.length));
  const out = [];
  for (const [k, t] of tabs.entries()) {
    const body = t === TB.cur ? (await docText()).slice(0, cap) : await pagesText(tabData(t).pages, cap);
    out.push(`##### קובץ ${k + 1}: ${dispName(tabName(t))} #####\n${body}`);
  }
  return { text: out.join('\n\n'), tabs };
}
async function citeFile(f, p) {
  const t = MULTI.tabs && MULTI.tabs[f - 1];
  if (!t || !TB.list.includes(t)) { toast('הקובץ הזה כבר לא פתוח'); return; }
  if (TB.cur !== t) { await switchTo(t); await new Promise(r => setTimeout(r, 400)); }
  citeJump(p, MULTI.line || '');
}
const chipsNow = () => { const k = edKind(); return k && ED_CHIPS[k] && window.OFFICE_AI ? ED_CHIPS[k] : CHIPS; };
const AS = {
  hist: new Map(),           // tab id -> [{q,a}] or {err}
  key() { return TB.cur ? TB.cur.id : 0; },
  list() { const k = this.key(); if (!this.hist.has(k)) this.hist.set(k, []); return this.hist.get(k); },
  forget(t) { this.hist.delete(t.id); },
  forgetAll() { this.hist.clear(); },
  onTab() { if (!panel.classList.contains('hidden')) renderLog(); },
  busySet: new Set(),        // tabs the assistant is working on (each file has its own)
  isBusy() { return this.busySet.has(this.key()); },
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
  // 6.1: [קובץ 2, עמ' 3] / [file 2, p. 3] → a button that opens that file at that page
  return esc(text).replace(/\[(?:קובץ|file)\s*(\d+)\s*[,،]\s*(?:עמ['׳]?|עמוד|p\.|page)\s*(\d+)[^\]]*\]/gi, (m, f, p) => `<button class="cite" type="button" data-f="${f}" data-p="${p}" title="פתיחת קובץ ${f} בעמוד ${p}">${esc(m.slice(1, -1))}</button>`).replace(/\[(?:עמ['׳]?|עמוד|עמודים|p\.|pp\.|page|стр\.|ص)\s*([\d\s,–-]+)\]/gi, (m, nums) => {
    const first = parseInt(nums, 10);
    return Number.isFinite(first) ? `<button class="cite" type="button" data-p="${first}" title="מעבר לעמוד ${first}">${esc(m.slice(1, -1))}</button>` : m;
  });
}
function answerEl(a, q, plain, it) {
  const el = h('<div class="ai-a"></div>');
  // in the Word / Excel editor there are no page numbers to jump to: the page marks are left out
  const strip = l => l.replace(/\s*\[(?:עמ['׳]?|עמוד|עמודים|p\.|pp\.|page)\s*[\d\s,–-]+\]/gi, '');
  el.innerHTML = a.split(/\n+/).filter(l => l.trim()).map(l => `<p dir="auto">${plain ? esc(strip(l)) : citeHtml(l)}</p>`).join('');
  const acts = h('<div class="ai-acts"><button type="button" data-a="copy">העתקה</button><button type="button" data-a="read">הקראה</button></div>');
  acts.querySelector('[data-a=copy]').onclick = async () => { try { await navigator.clipboard.writeText(a.replace(/\s*\[(?:עמ['׳]?|p\.)\s*[\d,\s–-]+\]/g, '')); toast('התשובה הועתקה'); } catch (_) { toast('ההעתקה לא הצליחה'); } };
  acts.querySelector('[data-a=read]').onclick = () => RA.speakText(a.replace(/\[[^\]]*\]/g, ''));
  el.append(acts);
  if (it && it.res && (it.res.asked || it.res.done)) el.append(resultEl(it));
  el.addEventListener('click', e => {
    const b = e.target.closest('.cite'); if (!b) return;
    const line = (b.closest('p,.an,.leaf,summary') || el).textContent;
    if (b.dataset.f) { MULTI.tabs = it && it.tabs; MULTI.line = line; citeFile(+b.dataset.f, +b.dataset.p); return; }
    citeJump(+b.dataset.p, line);
  });
  if (it && it.kind) { el.innerHTML = ''; el.append(studyEl(it, a)); el.append(acts); }
  return el;
}
/* 6.0: a citation jumps to its page AND lights up the sentence it came from, so the answer can be checked at a glance */
const normHe = s => String(s || '').normalize('NFKD').replace(/[֑-ׇ]/g, '').replace(/ך/g, 'כ').replace(/ם/g, 'מ').replace(/ן/g, 'נ').replace(/ף/g, 'פ').replace(/ץ/g, 'צ').toLowerCase();
const toks = s => (normHe(s).match(/[\p{L}\p{N}]{2,}/gu) || []).map(w => w.length > 3 && /^[והבלמשכ]/.test(w) && /[א-ת]/.test(w) ? w.slice(1) : w);
async function citeJump(n, line) {
  goPage(n);
  const P = S.pages[n - 1]; if (!P) return;
  const want = new Set(toks(String(line || '').replace(/\[[^\]]*\]/g, ''))); if (want.size < 2) return;
  // the words on the page appear once the page is drawn
  let spans = [];
  for (let k = 0; k < 30; k++) { spans = [...(P.txl ? P.txl.querySelectorAll('span') : [])].filter(s => s.textContent.trim()); if (spans.length) break; await new Promise(r => setTimeout(r, 100)); }
  if (!spans.length) return;
  const score = spans.map(sp => {
    const ws = toks(sp.textContent); if (!ws.length) return 0;
    // some PDFs keep Hebrew backwards: each word is also tried reversed
    let hit = 0; for (const w of ws) if (want.has(w) || want.has([...w].reverse().join(''))) hit++;
    return hit < 2 && ws.length > 2 ? 0 : hit / Math.sqrt(ws.length + 1);
  });
  let best = 0; score.forEach((v, i) => { if (v > score[best]) best = i; });
  if (!score[best]) return;
  const pick = []; for (let i = Math.max(0, best - 6); i <= Math.min(spans.length - 1, best + 6); i++) if (score[i] >= score[best] * 0.45) pick.push(spans[i]);
  P.el.querySelectorAll('.ai-src').forEach(x => x.remove());
  const base = P.el.getBoundingClientRect();
  for (const sp of pick) {
    const r = sp.getBoundingClientRect(); if (!r.width) continue;
    const box = document.createElement('div'); box.className = 'ai-src';
    Object.assign(box.style, { left: (r.left - base.left - 2) + 'px', top: (r.top - base.top - 2) + 'px', width: (r.width + 4) + 'px', height: (r.height + 4) + 'px' });
    P.el.append(box);
  }
  const first = P.el.querySelector('.ai-src');
  if (first) setTimeout(() => first.scrollIntoView({ behavior: matchMedia('(prefers-reduced-motion:reduce)').matches ? 'auto' : 'smooth', block: 'center' }), 250);
  setTimeout(() => P.el.querySelectorAll('.ai-src').forEach(x => x.classList.add('out')), 5000);
  setTimeout(() => P.el.querySelectorAll('.ai-src').forEach(x => x.remove()), 6200);
}
/* 6.0: study tools — flashcards, a quiz with a score, and a mind map — all from the open document */
const STUDY = {
  cards: 'Make 10 to 15 study flashcards from the most important content of the document. Output ONLY lines, one per card, exactly in this format:\nQ: question || A: short answer [page cite]\nWrite in the language of the document. No other text.',
  quiz: 'Make a quiz of 6 multiple-choice questions that check understanding of the document. Output ONLY blocks separated by an empty line, each exactly in this format:\nQ: question\nA) option\nB) option\nC) option\nD) option\nANSWER: the correct letter | one short sentence why [page cite]\nWrite in the language of the document. No other text.',
  // 6.7: checking a submitted work (an assignment, an exercise, an essay) — feedback the way a good teacher gives it
  check: o => `You are a fair, encouraging and exact teacher. The document is a work someone submitted: an assignment, homework, exercises, an essay or a project. The writer is ${o.level}.
${o.task ? 'The task, the questions or the rubric are given at the end of the document, after the line "##### המשימה / המחוון #####". That part is not the work: check the work against it, question by question or criterion by criterion.' : 'No task was given: work out from the document what it tries to do, and check it on its own terms.'}
Check correctness (for math, science and exercises redo every step yourself and find the exact step that goes wrong), completeness, reasoning, structure and language.
Output ONLY lines, each starting with one of these labels, in the language of the work:
SUMMARY: one or two sentences with the overall impression
GOOD: one strength [page cite]   (2 to 4 such lines)
FIX: one mistake, gap or missing part, and exactly how to correct it [page cite]   (as many as needed, the most important first; write the correct answer or step)
TIP: one way to make the work better   (2 to 4 such lines)
${o.grade ? 'GRADE: a number from 0 to 100 | one sentence why' + (o.task ? ', by the rubric' : '') : ''}
No other text, no markdown.`,
  map: 'Make a mind map of the document as an indented outline. The first line is the main subject. Every other line starts with "- " and is indented with 2 spaces per level (up to 3 levels under the subject). Short phrases, not sentences. Write in the language of the document. Output ONLY the outline.',
};
function studyEl(it, a) {
  const w = h('<div class="ai-study"></div>');
  const cite = s => citeHtml(s);
  if (it.kind === 'cards') {
    const cards = a.split(/\n+/).map(l => /(?:Q|ש)\s*[:：]\s*(.+?)\s*\|\|\s*(?:A|ת)\s*[:：]\s*(.+)/i.exec(l)).filter(Boolean);
    if (!cards.length) { w.innerHTML = a.split(/\n+/).map(l => `<p dir="auto">${cite(l)}</p>`).join(''); return w; }
    w.innerHTML = `<div class="ai-st-h">${cards.length} כרטיסיות · לוחצים על כרטיס כדי לראות את התשובה</div>` + cards.map(m => `<button type="button" class="ai-fc" aria-expanded="false"><span class="q" dir="auto">${esc(m[1])}</span><span class="an" dir="auto">${cite(m[2])}</span></button>`).join('');
    w.addEventListener('click', e => { if (e.target.closest('.cite')) return; const c = e.target.closest('.ai-fc'); if (c) c.setAttribute('aria-expanded', String(c.getAttribute('aria-expanded') !== 'true')); });
    return w;
  }
  if (it.kind === 'quiz') {
    const qs = [];
    for (const block of a.split(/\n\s*\n/)) {
      const q = /(?:Q|ש)\s*[:：]\s*(.+)/i.exec(block); if (!q) continue;
      const opts = [...block.matchAll(/^\s*([A-Dא-ד])[).]\s*(.+)$/gm)].map(m => ({ k: m[1], t: m[2] }));
      const ans = /(?:ANSWER|תשובה)\s*[:：]\s*([A-Dא-ד])\s*\|?\s*(.*)/i.exec(block);
      if (opts.length >= 2 && ans) qs.push({ q: q[1].trim(), opts, right: ans[1].toUpperCase(), why: ans[2] || '' });
    }
    if (!qs.length) { w.innerHTML = a.split(/\n+/).map(l => `<p dir="auto">${cite(l)}</p>`).join(''); return w; }
    let got = 0, done = 0;
    w.innerHTML = `<div class="ai-st-h">בוחן · ${qs.length} שאלות <b class="ai-score"></b></div>` + qs.map((x, i) => `<div class="ai-qz" data-i="${i}"><p class="q" dir="auto">${i + 1}. ${esc(x.q)}</p>${x.opts.map(o => `<button type="button" data-k="${esc(o.k)}" dir="auto">${esc(o.t)}</button>`).join('')}<p class="why" dir="auto" hidden></p></div>`).join('');
    // the answers are kept on the item, so the quiz stays answered when the panel is drawn again
    it.picks = it.picks || {};
    const pick = (b, keep) => {
      const box = b.closest('.ai-qz'); if (box.dataset.done) return; box.dataset.done = 1;
      if (keep) it.picks[box.dataset.i] = b.dataset.k;
      const x = qs[+box.dataset.i], ok = b.dataset.k.toUpperCase() === x.right; done++; if (ok) got++;
      box.querySelectorAll('button').forEach(o => { o.disabled = true; if (o.dataset.k.toUpperCase() === x.right) o.classList.add('right'); });
      if (!ok) b.classList.add('wrong');
      const why = box.querySelector('.why'); why.hidden = false; why.innerHTML = (ok ? '✓ נכון. ' : '✗ לא נכון. ') + cite(x.why);
      w.querySelector('.ai-score').textContent = `${got} מתוך ${done}` + (done === qs.length ? (got === qs.length ? ' · מושלם!' : got >= qs.length * 0.7 ? ' · יפה מאוד' : ' · כדאי לעבור שוב על החומר') : '');
    };
    w.addEventListener('click', e => { const b = e.target.closest('.ai-qz button'); if (b) pick(b, true); });
    for (const [i, k] of Object.entries(it.picks)) { const b = [...w.querySelectorAll(`.ai-qz[data-i="${i}"] button`)].find(o => o.dataset.k === k); if (b) pick(b, false); }
    return w;
  }
  if (it.kind === 'check') {
    const L = { SUMMARY: [], GOOD: [], FIX: [], TIP: [], GRADE: [] };
    for (const l of a.split(/\n+/)) { const m = /^\s*[-•*]?\s*(SUMMARY|GOOD|FIX|TIP|GRADE|סיכום|טוב|לתקן|טיפ|ציון)\s*[:：]\s*(.+)$/i.exec(l); if (!m) continue;
      const k = { 'סיכום': 'SUMMARY', 'טוב': 'GOOD', 'לתקן': 'FIX', 'טיפ': 'TIP', 'ציון': 'GRADE' }[m[1]] || m[1].toUpperCase(); L[k].push(m[2].trim()); }
    if (!L.GOOD.length && !L.FIX.length && !L.SUMMARY.length) { w.innerHTML = a.split(/\n+/).map(l => `<p dir="auto">${cite(l)}</p>`).join(''); return w; }
    const sec = (cls, title, list) => list.length ? `<div class="ai-ck-sec ${cls}"><h5>${title}</h5>${list.map(x => `<p dir="auto">${cite(x)}</p>`).join('')}</div>` : '';
    let grade = '';
    if (L.GRADE.length) { const [n, why] = L.GRADE[0].split(/\s*\|\s*/); const num = parseInt(n, 10);
      grade = `<div class="ai-ck-grade"><b dir="ltr">${Number.isFinite(num) ? num : esc(n)}</b><span dir="auto">ציון משוער מתוך 100${why ? ' · ' + esc(why) : ''}</span></div>`; }
    w.innerHTML = `<div class="ai-st-h">בדיקת העבודה${it.taskName ? ' · מול ' + esc(it.taskName) : ''}</div>`
      + (L.SUMMARY.length ? `<p class="ai-ck-sum" dir="auto">${cite(L.SUMMARY.join(' '))}</p>` : '') + grade
      + sec('fix', `✗ מה לתקן (${L.FIX.length})`, L.FIX) + sec('good', '✓ מה טוב', L.GOOD) + sec('tip', '💡 איך לשפר', L.TIP);
    return w;
  }
  // the mind map: an outline that opens and closes, branch by branch
  const lines = a.split(/\n/).filter(l => l.trim());
  if (!lines.length) return w;
  const root = { t: lines[0].replace(/^[\s\-•*#]+/, '').trim(), kids: [] }, stack = [{ lvl: -1, node: root }];
  for (const l of lines.slice(1)) {
    const lvl = Math.floor((l.match(/^\s*/)[0].replace(/\t/g, '  ').length) / 2);
    const node = { t: l.replace(/^[\s\-•*]+/, '').trim(), kids: [] }; if (!node.t) continue;
    while (stack.length > 1 && stack[stack.length - 1].lvl >= lvl) stack.pop();
    stack[stack.length - 1].node.kids.push(node); stack.push({ lvl, node });
  }
  const draw = (n, d) => n.kids.length ? `<details${d < 2 ? ' open' : ''}><summary dir="auto">${cite(n.t)}</summary><div class="kids">${n.kids.map(k => draw(k, d + 1)).join('')}</div></details>` : `<div class="leaf" dir="auto">${cite(n.t)}</div>`;
  w.innerHTML = `<div class="ai-st-h">מפת חשיבה · לוחצים על ענף כדי לפתוח או לסגור</div><div class="ai-map">${draw(root, 0)}</div>`;
  w.addEventListener('click', e => { if (e.target.closest('.cite')) e.preventDefault(); }); // a page mark inside a branch title jumps, without folding the branch
  return w;
}
// 6.0: the whole conversation about the file, as a text file
async function exportChat() {
  const items = AS.list().filter(x => x.a);
  if (!items.length) return;
  const t = [dispName(docTitle()), ''].concat(...items.map(x => ['שאלה: ' + (x.label || x.q), x.a.replace(/^❓.*$/gm, '').trim(), ''])).join('\n');
  const name = dispName(docTitle()).replace(/\.[a-z0-9]+$/i, '') + ' - שאלות ותשובות.txt';
  try { if (typeof saveBytes === 'function') { if (await saveBytes(name, new TextEncoder().encode('﻿' + t), 'txt')) toast('נשמר: ' + name); } else { await navigator.clipboard.writeText(t); toast('השיחה הועתקה'); } }
  catch (e) { toast('השמירה לא הצליחה'); }
}
// what the assistant changed in the file, and taking it back
function resultEl(it) {
  const r = it.res, w = h('<div class="ai-res"></div>');
  const n = r.done || 0;
  const tag = h(`<span class="done${n ? '' : ' none'}"></span>`);
  tag.textContent = n ? (n === 1 ? '✓ שינוי אחד בוצע בקובץ' : `✓ ${n} שינויים בוצעו בקובץ`) : 'לא בוצע שינוי בקובץ';
  w.append(tag);
  if (n && r.mark) {
    const b = h('<button type="button">ביטול השינוי</button>');
    if (it.undone) { b.textContent = 'השינוי בוטל'; b.disabled = true; }
    b.onclick = () => {
      if (AS.key() !== it.tab || !window.OFFICE_AI || !OFFICE_AI.revert(r)) { toast('אפשר לבטל רק כשהקובץ הזה פתוח. אפשר גם Ctrl+Z בעורך'); return; }
      it.undone = true; b.textContent = 'השינוי בוטל'; b.disabled = true;
      // later requests were made on top of this one: they are taken back with it
      for (const x of AS.list()) if (x !== it && x.res && x.res.mark && x.res.mark.hi >= r.mark.hi && x.tab === it.tab) x.undone = true;
      renderLog();
    };
    w.append(b);
  }
  if (r.skipped && n) { const s = h('<span></span>'); s.textContent = `(${r.skipped} לא בוצעו)`; w.append(s); }
  return w;
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
    { const b = h('<button class="btn ck-photo" type="button">📝 בדיקת עבודה מצילומים: דפי תלמיד, מבחן ומחוון</button>'); b.onclick = () => CK.open(); log.append(b); const c = h('<button class="btn ck-photo" type="button">🏫 בדיקת מבחנים לכיתה: תלמיד אחרי תלמיד, וקובץ ציונים</button>'); c.onclick = () => GR.open(); log.append(c); }
    $('aiQ').disabled = $('aiSend').disabled = true; $('aiRead').disabled = true;
    return;
  }
  $('aiQ').disabled = $('aiSend').disabled = AS.isBusy(); $('aiRead').disabled = false;
  $('aiQ').placeholder = edKind() && window.OFFICE_AI ? 'מה לעשות בקובץ?' : 'שאלו כל דבר על המסמך…';
  // 6.1: two PDF files or more open: one switch asks about all of them together
  const nPdf = !edKind() ? pdfTabs().length : 0;
  if (nPdf > 1) {
    const sw = h(`<label class="ai-multi"><input type="checkbox"${MULTI.on ? ' checked' : ''}><span><b>לשאול על כל ${nPdf} הקבצים הפתוחים</b><small>להשוואה בין חוזים, הצעות מחיר או גרסאות, ולמציאת פרט באחד מהם</small></span></label>`);
    sw.querySelector('input').onchange = e => { MULTI.on = e.target.checked; $('aiQ').placeholder = MULTI.on ? `שאלו על כל ${nPdf} הקבצים… (למשל: מה ההבדלים ביניהם?)` : 'שאלו כל דבר על המסמך…'; };
    log.append(sw);
    if (MULTI.on) $('aiQ').placeholder = `שאלו על כל ${nPdf} הקבצים… (למשל: מה ההבדלים ביניהם?)`;
  }
  const items = AS.list();
  if (!items.length) {
    if (edKind() && window.OFFICE_AI) log.append(h(`<div class="ai-hello">כתבו או אמרו מה לעשות ב-<b></b>, ואני אבצע: כותרת, עיצוב, טבלאות, תוכן חדש ועוד. אפשר גם לשאול שאלות על הקובץ.</div>`));
    else log.append(h(`<div class="ai-hello">שאלו כל שאלה על <b></b>, או בחרו:</div>`));
    log.lastChild.querySelector('b').textContent = dispName(docTitle());
    const chips = h('<div class="ai-chips"></div>');
    for (const [l, q] of chipsNow()) { const b = h(`<button class="ai-chip" type="button">${esc(l)}</button>`); b.onclick = () => askQ(q, l); chips.append(b); }
    log.append(chips);
    if (!edKind()) log.append(studyChips());
  }
  for (const it of items) {
    const q = h('<div class="ai-q" dir="auto"></div>'); q.textContent = it.label || it.q; log.append(q);
    if (it.a) log.append(answerEl(it.a, it.q, it.plain, it));
    else if (it.err) { const e = h('<div class="ai-err"></div>'); e.textContent = it.err; if (it.scanned) e.append(scannedHelp()); log.append(e); }
    else log.append(h(`<div class="ai-wait" id="aiWait">${it.wait || 'קורא את המסמך וכותב תשובה…'}</div>`));
  }
  if (items.length) {
    const chips = h('<div class="ai-chips"></div>');
    // 6.0: the follow-up questions the assistant suggested after its last answer come first
    const last = items[items.length - 1];
    for (const f of (last && last.follow) || []) { const b = h(`<button class="ai-chip follow" type="button" dir="auto"></button>`); b.textContent = f; b.onclick = () => askQ(f); chips.append(b); }
    for (const [l, q] of chipsNow().filter(c => !items.some(i => i.q === c[1] && i.a)).slice(0, last && last.follow && last.follow.length ? 1 : 3)) { const b = h(`<button class="ai-chip" type="button">${esc(l)}</button>`); b.onclick = () => askQ(q, l); chips.append(b); }
    if (!AS.isBusy() && chips.childElementCount) log.append(chips);
    if (!AS.isBusy() && !edKind()) {
      log.append(studyChips());
      if (items.some(x => x.a)) { const x = h('<button class="ai-export" type="button">שמירת כל השיחה כקובץ טקסט</button>'); x.onclick = exportChat; log.append(x); }
    }
  }
  log.scrollTop = log.scrollHeight;
}
function studyChips() {
  const w = h('<div class="ai-chips study"><span class="ai-st-l">ללמוד מהמסמך:</span></div>');
  for (const [l, k] of STUDY_CHIPS) { const b = h(`<button class="ai-chip" type="button">${esc(l)}</button>`); b.onclick = () => k === 'check' ? checkForm() : askQ(STUDY[k], l.replace(/^\S+\s/, ''), k); w.append(b); }
  return w;
}
/* 6.7: checking a work — the open file is the work; the task or rubric can come from another open file or be pasted */
function checkForm() {
  const log = $('aiLog'); const old = log.querySelector('.ai-ck-form'); if (old) { old.scrollIntoView({ block: 'nearest' }); return; }
  const others = pdfTabs().filter(t => t !== TB.cur);
  const f = h(`<div class="ai-ck-form"><h4>📝 בדיקת עבודה</h4><p>הקובץ הפתוח הוא העבודה. מקבלים משוב: מה לתקן ואיך, מה טוב, ואיך לשפר.</p><button class="btn ck-photo" type="button" data-a="photo">📷 העבודה או המבחן על נייר? צילום וסריקה של דפים (גם כתב יד)</button>
    <label>מה המשימה, השאלות או המחוון?<select data-k="task"><option value="">בלי, לבדוק לפי מה שכתוב בעבודה</option>${others.map((t, i) => `<option value="t${i}">הקובץ "${esc(dispName(tabName(t)))}"</option>`).join('')}<option value="paste">להדביק או לכתוב את המשימה</option></select></label>
    <textarea data-k="paste" hidden placeholder="מדביקים כאן את השאלות, ההנחיות או המחוון…" aria-label="המשימה או המחוון"></textarea>
    <label>מי כתב את העבודה?<select data-k="level"><option value="a school student">תלמיד או תלמידה בבית ספר</option><option value="a university or college student" selected>סטודנט או סטודנטית</option><option value="a professional (a work report or document)">עבודה מקצועית</option></select></label>
    <label class="ck"><input type="checkbox" data-k="grade" checked> לתת גם ציון משוער מתוך 100</label>
    <div class="acts"><button class="btn" type="button" data-a="x">ביטול</button><button class="btn primary" type="button" data-a="go">לבדוק את העבודה</button></div></div>`);
  const sel = f.querySelector('[data-k=task]'), ta = f.querySelector('[data-k=paste]');
  sel.onchange = () => { ta.hidden = sel.value !== 'paste'; if (!ta.hidden) ta.focus(); };
  f.querySelector('[data-a=x]').onclick = () => f.remove();
  f.querySelector('[data-a=photo]').onclick = () => { f.remove(); CK.open(); };
  f.querySelector('[data-a=go]').onclick = async () => {
    let extra = '', taskName = '';
    if (sel.value === 'paste') { extra = ta.value.trim(); taskName = 'המשימה שהודבקה'; if (!extra) { toast('מדביקים את המשימה, או בוחרים "בלי"'); ta.focus(); return; } }
    else if (sel.value) { const t = others[+sel.value.slice(1)]; if (!t || !TB.list.includes(t)) { toast('הקובץ הזה כבר לא פתוח'); return; }
      f.querySelector('[data-a=go]').disabled = true; extra = (await pagesText(tabData(t).pages, 120000)).replace(/=====\s*עמוד\s*(\d+)\s*=====/g, '(עמוד $1 של המשימה)'); taskName = dispName(tabName(t));
      if (!hasRealText(extra)) { f.querySelector('[data-a=go]').disabled = false; toast('לא מצאתי טקסט בקובץ המשימה. אפשר להדביק אותה'); return; } }
    const o = { level: f.querySelector('[data-k=level]').value, grade: f.querySelector('[data-k=grade]').checked, task: !!extra };
    f.remove();
    askQ(STUDY.check(o), 'בדיקת עבודה' + (taskName ? ' מול ' + taskName : ''), 'check', { extra, taskName });
  };
  log.append(f); f.scrollIntoView({ block: 'nearest' });
}
function scannedHelp() {
  const w = h('<div><button class="btn" type="button">זיהוי טקסט בכל העמודים</button> <button class="btn" type="button">מדף למסמך</button></div>');
  w.children[0].onclick = async () => { await runOCR(S.pages.map((_, i) => i)); textCache.delete(S.doc); };
  w.children[1].onclick = () => SC.open('doc');
  return w;
}
async function askQ(q, label, kind, opt) {
  q = String(q || '').trim();
  const tabKey = AS.key();
  if (!q || AS.busySet.has(tabKey) || !hasDoc()) return;
  if (panel.classList.contains('hidden')) showAI();
  const list = AS.list();
  const it = { q, label }; if (kind) { it.kind = kind; it.wait = kind === 'quiz' ? 'מכין בוחן מהמסמך…' : kind === 'map' ? 'מצייר מפת חשיבה…' : kind === 'check' ? 'קורא את העבודה ובודק אותה…' : 'מכין כרטיסיות…'; } list.push(it);
  if (opt && opt.taskName) it.taskName = opt.taskName;
  AS.busySet.add(tabKey); renderLog(); renderTabs();
  try {
    if (onEditor() && window.OFFICE_AI && OFFICE_AI.available()) {
      it.plain = true; it.tab = tabKey; it.wait = 'קורא את הקובץ ומכין את השינויים…'; renderLog();
      const history = list.filter(x => x.a && x !== it).slice(-3).map(x => ({ q: x.q, a: x.a }));
      const onWait = s => { it.wait = `השירות עמוס, מנסה שוב בעוד ${s} שניות…`; const w = document.getElementById('aiWait'); if (w && AS.key() === tabKey) w.textContent = it.wait; };
      if (!OFFICE_AI.plan) { const r = await OFFICE_AI.run(q, { history, onWait }); it.res = r; it.a = r.reply || (r.done ? 'בוצע.' : 'לא הצלחתי לבצע את זה. נסו לנסח את הבקשה אחרת, או לבקש דבר אחד בכל פעם.'); return; }
      const plan = await OFFICE_AI.plan(q, { history, onWait });
      it.plan = plan;
      const t = TB.list.find(x => x.id === tabKey);
      if (!t) { it.err = 'הקובץ נסגר לפני שהעוזר סיים'; return; }
      // the file is still in front: the changes are made now. Otherwise they wait for the file to come back
      if (TB.cur === t && edLive(t)) { aiFinish(it); return; }
      (t.aiQueue = t.aiQueue || []).push(it);
      it.a = plan.reply || 'התשובה מוכנה.';
      if (plan.actions.length) {
        it.queued = true;
        toast(`העוזר סיים ב"${dispName(tabName(t))}". השינויים ייכנסו לקובץ כשתחזרו אליו`);
      }
      return;
    }
    it.plain = onEditor();
    // 6.1: all the open PDF files together
    const multi = !it.plain && MULTI.on && pdfTabs().length > 1;
    let text;
    if (multi) { it.wait = 'קורא את כל הקבצים הפתוחים…'; renderLog(); const r = await allText(); text = r.text; it.tabs = r.tabs; it.label = (it.label || q) + ` · ${r.tabs.length} קבצים`; }
    else text = opt && opt.text ? opt.text : await docText();   // 6.7.1: pages read from photos come with their text
    const empty = it.plain ? !text.replace(/=====.*?=====/g, '').replace(/שורה \d+:/g, '').trim() : !hasRealText(text);
    if (empty) { if (it.plain) { it.err = 'הקובץ ריק, אין עדיין על מה לשאול'; return; } it.err = 'לא מצאתי טקסט במסמך. כנראה שהוא סרוק או מצולם — אפשר לזהות בו את הטקסט קודם:'; it.scanned = true; return; }
    // 6.7: the task or rubric of a work being checked goes after the work, clearly apart from it
    if (opt && opt.extra) text = text + '\n\n##### המשימה / המחוון #####\n' + String(opt.extra).slice(0, 120000);
    const history = list.filter(x => x.a && x !== it).slice(-4).map(x => ({ q: x.q, a: x.a }));
    const followOk = !kind && !it.plain;
    // 6.6.2: the same first question on the same text (a summary asked again, after clearing the chat or in another
    // tab of the same file) is answered from this session's memory: at once, and without using the daily share
    const ck = !history.length ? `${multi ? 'M' : ''}${followOk ? 'F' : ''}${q}\u0001${textSig(text)}` : null;
    if (ck && ASK_CACHE.has(ck)) it.a = ASK_CACHE.get(ck);
    else {
      it.a = await ai({ mode: 'ask', text, question: (multi ? q + MULTI_TAIL : q) + (followOk ? FOLLOW_TAIL : ''), history: history.map(x => ({ q: x.q, a: x.a.replace(/^❓.*$/gm, '').trim() })) }, { onWait: s => { it.wait = `השירות עמוס, מנסה שוב בעוד ${s} שניות…`; const w = document.getElementById('aiWait'); if (w && AS.key() === tabKey) w.textContent = it.wait; } });
      if (ck && it.a) { ASK_CACHE.set(ck, it.a); if (ASK_CACHE.size > 60) ASK_CACHE.delete(ASK_CACHE.keys().next().value); }
    }
    if (followOk) { const f = splitFollow(it.a); it.a = f.a; it.follow = f.follow; }
  } catch (e) { it.err = e.message || String(e); }
  finally { AS.busySet.delete(tabKey); renderTabs(); if (AS.key() === tabKey) renderLog(); }
}
// the assistant's changes go into the file in front
function aiFinish(it) {
  try {
    const r = OFFICE_AI.applyPlan(it.plan);
    it.res = r; it.queued = false;
    it.a = r.reply || (r.done ? 'בוצע.' : 'לא הצלחתי לבצע את זה. נסו לנסח את הבקשה אחרת, או לבקש דבר אחד בכל פעם.');
  } catch (e) { it.err = e.message || String(e); it.a = null; }
}
// a file comes back to the front: what the assistant finished for it meanwhile is put in now, in order
function aiFlush(t) {
  if (!t || !t.aiQueue || !t.aiQueue.length || TB.cur !== t || !edLive(t)) return;
  const q = t.aiQueue; t.aiQueue = [];
  let n = 0;
  for (const it of q) { if (!it.plan.actions.length) continue; aiFinish(it); if (it.res && it.res.done) n++; }
  if (n) toast(n === 1 ? 'השינויים שהעוזר הכין נכנסו לקובץ' : `${n} בקשות של העוזר נכנסו לקובץ`);
  renderTabs(); if (!panel.classList.contains('hidden')) renderLog();
}
window.SIMAN_AI = ai;
window.SIMAN_ASK = (q, label) => { if (panel.classList.contains('hidden')) showAI(); askQ(q, label); };
{ const s = document.createElement('script'); s.src = 'office-ai.js'; s.defer = true; document.head.append(s); }
{ const s = document.createElement('script'); s.src = 'office-translate.js'; s.defer = true; document.head.append(s); }
$('aiSend').onclick = () => { const q = $('aiQ').value; $('aiQ').value = ''; $('aiQ').style.height = ''; askQ(q); };
$('aiQ').addEventListener('input', e => { const t = e.target; t.style.height = 'auto'; t.style.height = Math.min(140, t.scrollHeight) + 'px'; });
$('aiRead').onclick = () => RA.start(currentPageIndex());
$('aiScan').onclick = () => SC.open();

// the assistant's button in the toolbar
const aiBtn = h(`<button class="btn ai" id="aiBtn" type="button" title="העוזר: שאלות על המסמך, סיכום, הקראה והפיכת דף למסמך">${ICON.spark}<span>עוזר</span></button>`);
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
const dlg = h(`<dialog class="scan" id="scanDlg" aria-label="מדף למסמך">
  <div class="sc-head"><h2>מדף למסמך ← קובץ ערוך</h2><button class="icon-btn" type="button" data-a="close" aria-label="סגירה">×</button></div>
  <div class="sc-body"></div>
  <div class="sc-foot hidden"></div>
</dialog>`);
document.body.append(dlg);
const camIn = h('<input type="file" accept="image/*" capture="environment" hidden>');
const imgIn = h('<input type="file" accept="image/*" multiple hidden>');
const pdfIn = h('<input type="file" accept="application/pdf,.pdf" hidden>');
dlg.append(camIn, imgIn, pdfIn);
const SC = {
  pages: [], tray: [], stop: false, running: false, hint: '', name: '', quality: (() => { try { const q = localStorage.getItem('siman.scan.q'); return ['auto', 'print', 'accurate', 'fast'].includes(q) ? q : 'auto'; } catch (_) { return 'auto'; } })(),
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
    // 5.9.6: closing while a page is being read stops the request itself (before, the window waited for it)
    if (this.running) { this.stop = true; if (this.ctl) this.ctl.abort(); }
    else if ((this.pages.some(p => p.blocks) || this.tray.length) && !confirm('לסגור? הטקסט שזוהה לא יישמר')) return;
    dlg.close();
    if (this.onFail && !this.pages.some(p => p.blocks)) { const f = this.onFail; this.onFail = this.onOk = null; f(); }
  },
  pick() {
    this.pages = []; this.tray = []; this.stop = false; this.running = false; this.name = '';
    const docN = S.doc ? S.pages.length : 0;
    const b = this.body(); this.foot().classList.add('hidden');
    b.innerHTML = `<div class="sc-pick">
      <h3>מצלמים דף, מקבלים מסמך מוכן</h3>
      <p class="sub">דף מודפס, כתב יד, צילום של מסמך או PDF סרוק: סימן קורא את הכתוב (דף מודפס בכמה שניות, גם כתב יד) ובונה ממנו קובץ Word או PDF עם כותרות, רשימות וטבלאות. לפני השמירה רואים את התמונה ואת הטקסט זה לצד זה ומתקנים. מילים שהזיהוי לא בטוח בהן מסומנות בצהוב.</p>
      <div class="sc-srcs">
        <button class="sc-src" type="button" data-s="cam">${ICON.scan}<b>צילום במצלמה</b><span>בטלפון: נפתחת המצלמה. מצלמים ישר מלמעלה, באור טוב. אפשר לצלם כמה עמודים לקובץ אחד</span></button>
        <button class="sc-src" type="button" data-s="img"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"><rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="9" cy="10" r="1.6"/><path d="M21 16l-5-5-8 8"/></svg><b>בחירת תמונות</b><span>אפשר כמה תמונות יחד. כל תמונה היא עמוד</span></button>
        <button class="sc-src" type="button" data-s="pdf"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"><path d="M7 3h7l4 4v14H7z"/><path d="M14 3v4h4"/></svg><b>PDF סרוק</b><span>קובץ שנסרק או צולם, עד 30 עמודים</span></button>
        ${docN ? `<button class="sc-src" type="button" data-s="doc"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"><rect x="5" y="3" width="14" height="18" rx="1.5"/><path d="M9 8h6M9 12h6M9 16h3"/></svg><b>הקובץ הפתוח</b><span></span></button>` : ''}
      </div>
      <div class="sc-opts"><label>שפת הכתב <select data-o="hint">${SC_LANGS.map(([v, l]) => `<option value="${v}">${l}</option>`).join('')}</select></label>
        <label>סוג הדף <select data-o="quality"><option value="auto">זיהוי אוטומטי (מומלץ): מודפס בכמה שניות, כתב יד בזהירות</option><option value="print">מודפס או מוקלד: הכי מהיר (ספר, מכתב, טופס, צילום מסך)</option><option value="accurate">כתב יד: הכי מדויק (חצי דקה עד שתי דקות לעמוד)</option><option value="fast">כתב יד מהיר: כמה שניות לעמוד, מנחש יותר</option></select></label></div>
      <p class="sc-fine">התמונות נשלחות לזיהוי אצל Google Gemini ולא נשמרות אצלנו. בשירות החינמי Google עשויה להשתמש בהן לשיפור השירות שלה, לכן לא כדאי לסרוק כך מסמכים רגישים במיוחד (תעודות זהות, מסמכים רפואיים). דף מודפס נקרא בכמה שניות; כתב יד לוקח יותר.</p>
    </div>`;
    if (docN) b.querySelector('[data-s=doc] span').textContent = `${dispName(S.office ? S.office.name : S.name)} · ${docN} ${docN === 1 ? 'עמוד' : 'עמודים'}`;
    const sel = b.querySelector('[data-o=hint]'); sel.value = this.hint; sel.onchange = () => { this.hint = sel.value; };
    const q = b.querySelector('[data-o=quality]'); q.value = this.quality; q.onchange = () => { this.quality = q.value; try { localStorage.setItem('siman.scan.q', q.value); } catch (_) {} };
    b.querySelector('[data-s=cam]').onclick = () => camIn.click();
    b.querySelector('[data-s=img]').onclick = () => imgIn.click();
    b.querySelector('[data-s=pdf]').onclick = () => pdfIn.click();
    const d = b.querySelector('[data-s=doc]'); if (d) d.onclick = () => this.fromDoc();
  },
  // the pages wait in a tray: more can be photographed or added, and then all are read into one file
  room() { return Math.max(0, 30 - this.pages.length - this.tray.length); },
  async fromImages(files, fromCamera = false) {
    if (!files.length) return;
    if (!this.name && !this.pages.length && !this.tray.length) this.name = files.length === 1 && !/^(image|img|photo|pxl|dsc|תמונה)[-_ ]?\d*/i.test(files[0].name) ? files[0].name.replace(/\.[a-z0-9]+$/i, '') : 'כתב יד';
    const bad = [], take = files.slice(0, this.room());
    // 6.6: several pictures are prepared three at a time (each waits for its page-edge check), in their order
    const got = new Array(take.length).fill(null); let nx = 0, ready = 0;
    this.wait(take.length > 1 ? `מזהה את הדפים ומכין אותם… (0 מתוך ${take.length})` : 'מזהה את הדף ומכין אותו…');
    await Promise.all(Array.from({ length: Math.min(3, take.length) }, async () => {
      while (nx < take.length) {
        const i = nx++, f = take[i];
        try { got[i] = await imgToJpeg(f); }
        catch (e) { bad.push(`${f.name}${/heic|heif/i.test(f.type + f.name) ? ' (HEIC: צלמו ישר מהמצלמה או שמרו כ-JPG)' : ''}`); }
        ready++; if (take.length > 1) this.wait(`מזהה את הדפים ומכין אותם… (${ready} מתוך ${take.length})`);
      }
    }));
    for (const it of got) if (it) this.tray.push(it);
    // a photo just taken: its frame is shown at once, to confirm or fix
    if (fromCamera && take.length === 1 && this.tray.length && !bad.length) { this.frame(this.tray.length - 1); return; }
    this.showTray();
    if (files.length > take.length) this.note('אפשר עד 30 עמודים בקובץ אחד');
    if (bad.length) this.note('לא הצלחתי לקרוא: ' + bad.join(', '));
  },
  // the frame of one photo: four corners to drag onto the edges of the sheet
  frame(i) {
    const it = this.tray[i]; if (!it) return this.showTray();
    const b = this.body(), f = this.foot();
    let q = (it.quad || FULL).map(p => [...p]);
    b.innerHTML = `<div class="sc-frame"><p class="sc-fhelp">${it.auto ? 'זה הדף שזיהיתי. אם צריך, גררו את הפינות לקצוות הדף.' : 'גררו את ארבע הפינות לקצוות הדף, כדי שרק הדף ייקרא.'}</p>
      <div class="sc-fwrap"><img alt="התמונה שצולמה" src="${it.view}" draggable="false">
      <svg viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true"><path class="shade" fill-rule="evenodd"/><polygon class="edge"/></svg>
      ${[0, 1, 2, 3].map(k => `<button type="button" class="sc-hd" data-k="${k}" aria-label="${['פינה ימנית עליונה', 'פינה שמאלית עליונה', 'פינה שמאלית תחתונה', 'פינה ימנית תחתונה'][k === 0 ? 1 : k === 1 ? 0 : k === 2 ? 3 : 2]}"></button>`).join('')}</div></div>`;
    const wrap = b.querySelector('.sc-fwrap'), path = b.querySelector('.shade'), poly = b.querySelector('.edge'), hs = [...b.querySelectorAll('.sc-hd')];
    const draw = () => {
      const pts = q.map(([x, y]) => `${x * 100},${y * 100}`);
      poly.setAttribute('points', pts.join(' '));
      path.setAttribute('d', `M0,0H100V100H0Z M${pts.join(' L')}Z`);
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
    f.classList.remove('hidden');
    f.innerHTML = `<span class="hint" style="margin-inline-end:auto"></span><button class="btn" type="button" data-a="full">כל התמונה</button>${it.auto ? '<button class="btn" type="button" data-a="auto">הדף שזוהה</button>' : '<button class="btn" type="button" data-a="detect">זיהוי הדף</button>'}<button class="btn primary" type="button" data-a="ok">אישור</button>`;
    f.querySelector('[data-a=full]').onclick = () => { q = FULL.map(p => [...p]); draw(); };
    const au = f.querySelector('[data-a=auto]'); if (au) au.onclick = () => { q = it.auto.map(p => [...p]); draw(); };
    const de = f.querySelector('[data-a=detect]');
    if (de) de.onclick = async () => { de.disabled = true; de.textContent = 'מזהה…'; const a = await detectSheet(it);
      if (a) { it.auto = a; q = a.map(p => [...p]); draw(); de.textContent = 'זוהה ✓'; } else { de.textContent = 'לא זוהה. גררו ידנית'; } };
    f.querySelector('[data-a=ok]').onclick = async () => {
      if (quadArea(q) < 0.03) { this.note('המסגרת קטנה מדי'); return; }
      it.quad = q; this.wait('מיישר את הדף…'); try { await cutSheet(it); } catch (e) { console.warn(e); this.note('לא הצלחתי לחתוך. נשארה התמונה המלאה'); }
      this.showTray();
    };
  },
  showTray() {
    const had = this.pages.filter(p => p.blocks || p.err).length;
    if (!this.tray.length && !had) { this.pick(); return; }
    const b = this.body(), f = this.foot(), n = this.tray.length;
    const pg = k => k === 1 ? 'עמוד אחד' : `${k} עמודים`;
    b.innerHTML = `<div class="sc-trayhead"><h3>${n ? pg(n) + (had ? ' להוספה' : ' לקריאה') : 'הוספת עמודים'}</h3><p>${had ? `בקובץ כבר יש ${pg(had)}. העמודים החדשים יתווספו אחריהם. ` : ''}הכל נשמר כקובץ אחד, לפי הסדר.</p></div>
      <ul class="sc-tray">${this.tray.map((p, i) => `<li><img alt="עמוד ${had + i + 1}" src="${p.url}">עמוד ${had + i + 1}<button type="button" class="x" data-del="${i}" aria-label="הסרת עמוד ${had + i + 1}">×</button>${p.srcBlob ? `<button type="button" class="fr" data-frame="${i}" aria-label="מסגרת לעמוד ${had + i + 1}">${ICON.crop} מסגרת</button>` : ''}</li>`).join('')}
        ${this.room() ? `<li class="add" role="button" tabindex="0" data-add="cam">${ICON.scan}צילום עמוד${n || had ? ' נוסף' : ''}</li><li class="add" role="button" tabindex="0" data-add="img"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"><rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="9" cy="10" r="1.6"/><path d="M21 16l-5-5-8 8"/></svg>בחירת תמונות</li><li class="add" role="button" tabindex="0" data-add="pdf"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"><path d="M7 3h7l4 4v14H7z"/><path d="M14 3v4h4"/></svg>PDF סרוק</li>` : ''}
      </ul>`;
    b.querySelectorAll('[data-del]').forEach(x => x.onclick = () => { this.tray.splice(+x.dataset.del, 1); this.showTray(); });
    b.querySelectorAll('[data-frame]').forEach(x => x.onclick = () => this.frame(+x.dataset.frame));
    b.querySelectorAll('[data-add]').forEach(x => { const go = () => ({ cam: camIn, img: imgIn, pdf: pdfIn })[x.dataset.add].click(); x.onclick = go; x.onkeydown = e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); go(); } }; });
    f.classList.remove('hidden');
    f.innerHTML = `<span class="hint" style="margin-inline-end:auto"></span>${had ? '<button class="btn" type="button" data-a="back">חזרה לטקסט</button>' : '<button class="btn" type="button" data-a="restart">ביטול</button>'}<button class="btn primary" type="button" data-a="read" ${n ? '' : 'disabled'}>${n > 1 ? `קריאת ${n} העמודים` : 'קריאת העמוד'}</button>`;
    const back = f.querySelector('[data-a=back]'); if (back) back.onclick = () => { this.tray = []; this.review(); };
    const rs = f.querySelector('[data-a=restart]'); if (rs) rs.onclick = () => this.pick();
    f.querySelector('[data-a=read]').onclick = () => { const pics = this.tray; this.tray = []; this.run(pics, '', had > 0); };
  },
  async fromPdf(file) {
    if (!file) return;
    if (!this.name && !this.pages.length && !this.tray.length) this.name = file.name.replace(/\.pdf$/i, '');
    this.wait('קורא את ה-PDF…');
    try {
      const d = await pdfjs.getDocument({ data: new Uint8Array(await file.arrayBuffer()), isEvalSupported: false }).promise;
      const n = Math.min(this.room(), d.numPages);
      for (let i = 1; i <= n; i++) { this.wait(`מכין עמוד ${i} מתוך ${n}…`); this.tray.push(await pageToJpeg(await d.getPage(i))); }
      d.destroy(); this.showTray();
      if (d.numPages > n) this.note(`נוספו ${n} עמודים: אפשר עד 30 עמודים בקובץ אחד`);
    } catch (e) { console.error(e); this.showTray(); this.note(e && e.name === 'PasswordException' ? 'הקובץ מוגן בסיסמה' : 'לא הצלחתי לפתוח את ה-PDF'); }
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
  async run(pics, pre = '', append = false) {
    this.keepEdits();
    const fresh = pics.map(p => ({ ...p, blocks: null, err: null }));
    if (append) this.pages = this.pages.filter(p => p.blocks || p.err);
    const start = append ? this.pages.length : 0;
    this.pages = append ? this.pages.concat(fresh) : fresh;
    this.running = true; this.stop = false; this.ctl = new AbortController();
    // 6.6.4: the screen stays on while the pages are read (when it went off, the phone cut the line to the service)
    let wake = null; try { if (navigator.wakeLock) wake = await navigator.wakeLock.request('screen'); } catch (_) {}
    const b = this.body();
    b.innerHTML = `<p class="sc-status" role="status" style="margin:0 0 14px;font-size:15px"></p><ul class="sc-prog">${fresh.map((p, i) => `<li><img alt="עמוד ${start + i + 1}" src="${p.url}"><span>ממתין</span></li>`).join('')}</ul>`;
    const f = this.foot(); f.classList.remove('hidden');
    f.innerHTML = '<span class="hint" style="margin-inline-end:auto"></span><button class="btn" type="button" data-a="stop">עצירה</button>';
    f.querySelector('[data-a=stop]').onclick = () => { this.stop = true; this.ctl.abort(); };
    const st = b.querySelector('.sc-status'), lis = [...b.querySelectorAll('.sc-prog li')];
    // 6.6: printed pages are read two at a time (each is one quick request); handwriting stays one page at a time
    const par = (this.quality === 'auto' || this.quality === 'print') && fresh.length > 1 ? 2 : 1;
    let nextI = 0, t00 = Date.now();
    const one = async i => {
      const p = fresh[i], li = lis[i], sp = li.querySelector('span');
      li.className = 'now'; sp.textContent = 'קורא…';
      const t0 = Date.now(), lineOf = () => `${pre}קורא ${fresh.length > 1 ? `את העמודים: ${fresh.filter(x => x.blocks || x.err).length} מתוך ${fresh.length} מוכנים` : 'את העמוד'}`;
      let line = lineOf(); st.textContent = line + '…';
      let hand = false;
      const clock = setInterval(() => { const sec = Math.round((Date.now() - t0) / 1000); if (!/שורה|בודק/.test(sp.textContent)) sp.textContent = `${hand ? 'כתב יד: קורא בזהירות…' : 'קורא…'} ${sec} שנ׳`; line = lineOf(); st.textContent = `${line} · ${Math.round((Date.now() - t00) / 1000)} שניות${hand && sec > 45 ? ' (דף כתב יד צפוף לוקח לפעמים עד שתי דקות)' : ''}`; }, 1000);
      try {
        const raw = await readPage(p, this.hint, this.quality, { signal: this.ctl.signal, onHand: () => { hand = true; sp.textContent = 'כתב יד: קורא בזהירות…'; }, onWait: s => { st.textContent = `השירות עמוס, ממשיך בעוד ${s} שניות…`; }, onLines: (d, n) => { sp.textContent = `שורה ${d} מתוך ${n}`; }, onCheck: () => { sp.textContent = 'בודק את כל הדף שוב…'; } }).finally(() => clearInterval(clock));
        p.blocks = parseBlocks(raw); p.lang = p.blocks.lang;
        li.className = 'ok'; sp.textContent = p.blocks.length ? `${p.blocks.length} קטעים · ${Math.round((Date.now() - t0) / 1000)} שנ׳` : 'עמוד ריק';
      } catch (e) {
        if (this.stop && e && e.name === 'AbortError') { li.className = ''; sp.textContent = 'נעצר'; return false; }
        p.err = e.message || String(e); li.className = 'bad'; sp.textContent = 'לא נקרא';
        if (/מכסה|לא הוקם|חוסם|אין חיבור/.test(p.err)) { st.textContent = p.err; this.stop = true; if (this.ctl) this.ctl.abort(); return false; }
      }
      return true;
    };
    const lane = async () => {
      while (nextI < fresh.length && !this.stop) {
        const i = nextI++;
        if (!(await one(i))) break;
        if (nextI < fresh.length && !this.stop && par === 1) await sleep(1200);   // the free service allows only a few pages a minute
      }
    };
    await Promise.all(Array.from({ length: par }, lane));
    this.running = false;
    try { if (wake) await wake.release(); } catch (_) {}
    const done = this.pages.filter(p => p.blocks);
    if (!done.length && !append) {
      st.textContent = (this.pages.find(p => p.err) || {}).err || (this.stop ? 'הקריאה נעצרה' : 'לא זוהה טקסט');
      f.innerHTML = '<span class="hint" style="margin-inline-end:auto"></span>' + (this.onFail ? '<button class="btn primary" type="button" data-a="back">חזרה לעמודים שצילמתי</button> ' : '') + '<button class="btn" type="button" data-a="again">התחלה מחדש</button>';
      f.querySelector('[data-a=again]').onclick = () => { this.onFail = this.onOk = null; this.pick(); };
      const bk = f.querySelector('[data-a=back]');
      if (bk) bk.onclick = () => { const fn = this.onFail; this.onFail = this.onOk = null; dlg.close(); fn(); };
      return;
    }
    if (this.onOk) { const fn = this.onOk; this.onFail = this.onOk = null; try { fn(); } catch (_) {} }
    this.review();
  },
  keepEdits() { for (const p of this.pages) if (p.ed && p.ed.isConnected) p.html = p.ed.innerHTML; },
  review() {
    this.keepEdits();
    const b = this.body();
    b.innerHTML = '<div class="sc-rev"></div>';
    const rev = b.firstChild;
    this.pages.forEach((p, i) => {
      if (!p.blocks && !p.err) return;
      const row = h(`<section class="sc-pg" aria-label="עמוד ${i + 1}"><figure><img alt="המקור, עמוד ${i + 1}" src="${p.url}"><figcaption>המקור · עמוד ${i + 1}${p.err ? ' · לא נקרא: ' + esc(p.err) : ''} · <button type="button" class="sc-again">לקרוא את העמוד שוב במצב מדויק</button></figcaption></figure><div class="sc-ed" contenteditable="true" spellcheck="true"></div></section>`);
      const ed = row.querySelector('.sc-ed');
      ed.innerHTML = p.html != null ? p.html : p.blocks && p.blocks.length ? blocksHtml(p.blocks) : '<p class="empty">לא זוהה טקסט בעמוד הזה. אפשר להקליד כאן.</p>';
      ed.dir = p.blocks && p.blocks.length ? (blocksRtl(p.blocks) ? 'rtl' : 'ltr') : 'rtl';
      ed.addEventListener('focus', () => { try { document.execCommand('defaultParagraphSeparator', false, 'p'); } catch (_) {} });
      ed.addEventListener('input', () => { ed.querySelectorAll('mark').forEach(m => { if (!m.textContent.trim()) m.remove(); }); this.count(); });
      ed.addEventListener('click', e => { const m = e.target.closest('mark'); if (m && !m.dataset.seen) { m.dataset.seen = 1; } });
      row.querySelector('img').onclick = () => { const z = h(`<div class="sc-zoom"><img alt="" src="${p.url}"></div>`); z.onclick = () => z.remove(); dlg.append(z); };
      p.ed = ed;
      row.querySelector('.sc-again').onclick = () => this.again(i, row);
      if (this.pages.filter(x => x.blocks || x.err).length > 1) {
        const del = h('<button type="button" class="sc-del">הסרת העמוד</button>');
        del.onclick = () => { if (!confirm(`להסיר את עמוד ${i + 1} מהקובץ?`)) return; this.keepEdits(); this.pages.splice(i, 1); this.review(); };
        row.querySelector('figcaption').append(' · ', del);
      }
      rev.append(row);
    });
    const f = this.foot(); f.classList.remove('hidden');
    f.innerHTML = `<label>שם הקובץ <input type="text" data-o="name" dir="auto"></label>
      <span class="hint" data-o="unsure"></span>
      <button class="btn" type="button" data-a="more" title="צילום או הוספה של עוד עמודים לאותו קובץ">+ הוספת עמודים</button>
      <button class="btn" type="button" data-a="copy">העתקת הטקסט</button>
      <button class="btn" type="button" data-a="edit" title="פתיחה בעורך ה-Word של סימן">עריכה בעורך Word</button>
      <button class="btn" type="button" data-a="pdf">שמירה כ-PDF</button>
      <button class="btn primary" type="button" data-a="docx">שמירה כ-Word</button>`;
    const nm = f.querySelector('[data-o=name]'); nm.value = this.name || 'סריקה';
    nm.oninput = () => { this.name = nm.value; };
    f.querySelector('[data-a=docx]').onclick = () => this.save('docx');
    f.querySelector('[data-a=pdf]').onclick = () => this.save('pdf');
    f.querySelector('[data-a=edit]').onclick = () => this.save('edit');
    f.querySelector('[data-a=more]').onclick = () => { if (!this.room()) { this.note('אפשר עד 30 עמודים בקובץ אחד'); return; } this.keepEdits(); this.tray = []; this.showTray(); };
    f.querySelector('[data-a=copy]').onclick = async () => { try { await navigator.clipboard.writeText(this.collect().map(pg => pg.map(blockText).join('\n')).join('\n\n')); toast('הטקסט הועתק'); } catch (_) { toast('ההעתקה לא הצליחה'); } };
    this.count();
  },
  async again(i, row) {
    const p = this.pages[i], btn = row.querySelector('.sc-again'), ed = row.querySelector('.sc-ed');
    if (this.running) return;
    this.running = true; btn.disabled = true;
    const t0 = Date.now(), clock = setInterval(() => { btn.textContent = `קורא שוב… ${Math.round((Date.now() - t0) / 1000)} שנ׳`; }, 1000);
    try {
      const raw = await readPage(p, this.hint, 'accurate', {});
      const bl = parseBlocks(raw);
      if (bl.length) { p.blocks = bl; p.html = null; ed.innerHTML = blocksHtml(bl); ed.dir = blocksRtl(bl) ? 'rtl' : 'ltr'; this.count(); btn.textContent = 'נקרא שוב ✓'; }
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
  // after saving: where the file went, and two ways out if the phone did not keep it
  saved(name, bytes, k, n) {
    const f = this.foot(); f.querySelectorAll('.sc-saved').forEach(x => x.remove());
    const box = h(`<div class="sc-saved" role="status"><b></b><span></span><div class="sc-saved-a"><button class="btn" type="button" data-s="open">פתיחת הקובץ עכשיו</button></div></div>`);
    box.querySelector('b').textContent = savedNote(k, n);
    box.querySelector('span').textContent = `שם הקובץ: ${name}`;
    const file = () => new File([bytes], name, { type: k === 'PDF' ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    box.querySelector('[data-s=open]').onclick = () => { dlg.close(); openFile(file()); };
    if (ANDROID && navigator.canShare && navigator.canShare({ files: [file()] })) {
      const sh = h('<button class="btn" type="button">שליחה / שמירה ב-Drive</button>');
      sh.onclick = async () => { try { await navigator.share({ files: [file()], title: name }); } catch (_) {} };
      box.querySelector('.sc-saved-a').append(sh);
    }
    f.prepend(box); this.count();
  },
  collect() { return this.pages.filter(p => p.ed).map(p => domBlocks(p.ed)).filter(pg => pg.length); },
  async save(kind) {
    const pages = this.collect();
    // מילים שהמשתמש תיקן נלמדות לקריאות הבאות
    try { this.pages.forEach(p => { if (p.ed && p.blocks && p.blocks.length) VOCAB.learn(p.blocks.map(blockText).join(' '), domBlocks(p.ed).map(blockText).join(' ')); }); } catch (_) {}
    if (!pages.length) { toast('אין טקסט לשמירה'); return; }
    // a new name each time when the name is the default one: an old broken file of the same name is never opened by mistake
    let base = cleanBase(this.name) || 'סריקה';
    if (/^(סריקה|כתב יד)$/.test(base)) { const d = new Date(), z = n => String(n).padStart(2, '0'); base += ` ${z(d.getDate())}-${z(d.getMonth() + 1)} ${z(d.getHours())}-${z(d.getMinutes())}`; }
    try {
      if (kind === 'pdf') {
        this.note('בונה PDF…'); const bytes = await blocksPdf(pages);
        if (await saveBytes(base + '.pdf', bytes, 'pdf')) this.saved(base + '.pdf', bytes, 'PDF', pages.length); else this.count();
        return;
      }
      const bytes = blocksDocx(pages);
      if (kind === 'edit') { dlg.close(); await openWordEditor(bytes, base + '.docx'); return; }
      if (await saveBytes(base + '.docx', bytes, 'docx')) this.saved(base + '.docx', bytes, 'Word', pages.length); else this.count();
    } catch (e) { console.error(e); this.note('השמירה נכשלה: ' + (e.message || e)); }
  },
};
const ANDROID = /Android/i.test(navigator.userAgent);
const savedNote = (k, n) => `נשמר כקובץ ${k} אחד${n > 1 ? ` עם ${n} עמודים` : ''} ✓${ANDROID ? ' · בתיקיית "הורדות" (Downloads)' : ''}`;
dlg.querySelector('[data-a=close]').onclick = () => SC.close();
dlg.addEventListener('cancel', e => { if (SC.running || SC.pages.some(p => p.blocks) || SC.tray.length) { e.preventDefault(); SC.close(); } });
camIn.onchange = e => { const f = [...e.target.files]; e.target.value = ''; SC.fromImages(f, true); };
imgIn.onchange = e => { const f = [...e.target.files]; e.target.value = ''; SC.fromImages(f); };
pdfIn.onchange = e => { const f = e.target.files[0]; e.target.value = ''; SC.fromPdf(f); };
window.SIMAN_SCAN = SC;
window.SIMAN_AI = ai;

/* =====================================================================
   6.7.1: CHECKING A WORK FROM PHOTOS — the student's pages (also handwriting) and the exam or task
   are photographed, scanned or picked as files; the pages are read, gathered into ONE PDF that opens
   in Siman, and the work is checked there. "Gather only" makes the one file without a check.
   ===================================================================== */
const ckDlg = h(`<dialog class="scan ckd" id="ckDlg" aria-label="בדיקת עבודה">
  <div class="sc-head"><h2>📝 בדיקת עבודה מצילומים וקבצים</h2><button class="icon-btn" type="button" data-a="close" aria-label="סגירה">×</button></div>
  <div class="sc-body"></div><div class="sc-foot"></div></dialog>`);
document.body.append(ckDlg);
const ckCam = h('<input type="file" accept="image/*" capture="environment" hidden>');
const ckImg = h('<input type="file" accept="image/*" multiple hidden>');
const ckFile = h('<input type="file" accept="application/pdf,.pdf,.docx,image/*" multiple hidden>');
ckDlg.append(ckCam, ckImg, ckFile);
const CK_MAX = 30;
const CK = {
  work: [], task: [], taskText: '', target: 'work', busy: false, stop: false, level: 'a school student', grade: true, order: 'task', who: '',
  body: () => ckDlg.querySelector('.sc-body'), foot: () => ckDlg.querySelector('.sc-foot'),
  open() {
    if (typeof needReg === 'function' && needReg()) { showReg(); return; }
    try {
      if (!ckDlg.open) { this.work = []; this.task = []; this.taskText = ''; this.who = '';
        // 6.7.2: some phone browsers refuse a modal right after another window closed: then it opens as a plain one
        try { ckDlg.showModal(); } catch (_) { ckDlg.setAttribute('open', ''); } }
      this.draw();
    } catch (e) { console.error(e); toast('לא הצלחתי לפתוח את בדיקת העבודה: ' + (e.message || e)); }
  },
  close() {
    if (this.busy) { this.stop = true; if (this.ctl) this.ctl.abort(); return; }
    if ((this.work.length || this.task.length) && !confirm('לסגור? הדפים שצולמו לא יישמרו')) return;
    ckDlg.close();
  },
  note(m) { const n = this.foot().querySelector('.hint'); if (n) n.textContent = m; else toast(m); },
  trayHtml(list, part) {
    const add = list.length < CK_MAX ? `<li class="add" role="button" tabindex="0" data-add="cam" data-p="${part}">${ICON.scan}צילום עמוד</li>
      <li class="add" role="button" tabindex="0" data-add="img" data-p="${part}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"><rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="9" cy="10" r="1.6"/><path d="M21 16l-5-5-8 8"/></svg>תמונות או סריקות</li>
      <li class="add" role="button" tabindex="0" data-add="file" data-p="${part}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"><path d="M7 3h7l4 4v14H7z"/><path d="M14 3v4h4"/></svg>קבצים (PDF, Word)</li>
      ${S.doc && part === 'work' && !list.some(x => x.fromDoc) ? `<li class="add" role="button" tabindex="0" data-add="doc" data-p="${part}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"><rect x="5" y="3" width="14" height="18" rx="1.5"/><path d="M9 8h6M9 12h6M9 16h3"/></svg>הקובץ הפתוח</li>` : ''}` : '';
    return `<ul class="sc-tray">${list.map((p, i) => `<li><img alt="${part === 'work' ? 'עבודה' : 'משימה'}, עמוד ${i + 1}" src="${p.url}">${esc(p.label || 'עמוד ' + (i + 1))}${p.text ? ' · יש טקסט' : ''}<button type="button" class="x" data-del="${i}" data-p="${part}" aria-label="הסרת עמוד ${i + 1}">×</button>${p.srcBlob ? `<button type="button" class="fr" data-frame="${i}" data-p="${part}">${ICON.crop} מסגרת</button>` : ''}</li>`).join('')}${add}</ul>`;
  },
  draw() {
    const b = this.body(), f = this.foot();
    b.innerHTML = `<div class="ck-class"><span>🏫 <b>מבחנים של כיתה שלמה?</b> מעלים מבחן ומחוון, ואז תלמיד אחרי תלמיד, ובסוף מקבלים קובץ עם כל הציונים והפירוט.</span><button class="btn" type="button" data-a="class">בדיקת מבחנים לכיתה</button></div>
      <div class="ck-sec"><div class="sc-trayhead"><h3>1. העבודה</h3><p>הדפים של התלמיד או הסטודנט: מצלמים דף אחרי דף (גם כתב יד), סורקים, או בוחרים קבצים.</p></div>${this.trayHtml(this.work, 'work')}</div>
      <div class="ck-sec"><div class="sc-trayhead"><h3>2. המבחן, המשימה או המחוון <small>(לא חובה)</small></h3><p>מצלמים את דף השאלות או המחוון, או מדביקים אותם בתיבה.</p></div>${this.trayHtml(this.task, 'task')}
        <textarea class="ck-txt" placeholder="או מדביקים או כותבים כאן את השאלות, ההנחיות או המחוון…" aria-label="המשימה בטקסט"></textarea></div>
      <div class="ck-sec ck-opts"><label>שם התלמיד או העבודה <input type="text" data-o="who" placeholder="לא חובה, נכנס לשם הקובץ" dir="auto"></label>
        <label>מי כתב את העבודה? <select data-o="level"><option value="a school student">תלמיד או תלמידה בבית ספר</option><option value="a university or college student">סטודנט או סטודנטית</option><option value="a professional (a work report or document)">עבודה מקצועית</option></select></label>
        <label>סדר הדפים בקובץ <select data-o="order"><option value="task">המשימה ואחריה העבודה</option><option value="work">העבודה ואחריה המשימה</option><option value="only">העבודה בלבד</option></select></label>
        <label class="ck"><input type="checkbox" data-o="grade"> ציון משוער מתוך 100</label>
        <label>כתב היד והדפים <select data-o="quality"><option value="auto">זיהוי אוטומטי (מומלץ)</option><option value="accurate">כתב יד: הכי מדויק (עד שתי דקות לעמוד)</option><option value="print">מודפס בלבד: הכי מהיר</option></select></label></div>
      <p class="sc-fine">דפים מצולמים נקראים אצל Google Gemini (כתב יד לוקח חצי דקה עד שתי דקות לעמוד) ולא נשמרים אצלנו. כל הדפים מתאחדים לקובץ PDF אחד שנפתח בסימן, ואפשר לסמן ולחתום עליו ולשלוח אותו.</p>`;
    b.querySelector('[data-a=class]').onclick = () => { if ((this.work.length || this.task.length) && !confirm('לעבור לבדיקת כיתה? הדפים שצולמו כאן לא יישמרו')) return; this.work = []; this.task = []; ckDlg.close(); GR.open(); };
    const ta = b.querySelector('.ck-txt'); ta.value = this.taskText; ta.oninput = () => { this.taskText = ta.value; this.foot().querySelector('[data-a=check]') && this.btns(); };
    const who = b.querySelector('[data-o=who]'); who.value = this.who; who.oninput = () => { this.who = who.value; };
    const lv = b.querySelector('[data-o=level]'); lv.value = this.level; lv.onchange = () => { this.level = lv.value; };
    const od = b.querySelector('[data-o=order]'); od.value = this.order; od.onchange = () => { this.order = od.value; };
    const gr = b.querySelector('[data-o=grade]'); gr.checked = this.grade; gr.onchange = () => { this.grade = gr.checked; };
    const qu = b.querySelector('[data-o=quality]'); qu.value = ['auto', 'accurate', 'print'].includes(SC.quality) ? SC.quality : 'auto'; qu.onchange = () => { SC.quality = qu.value; };
    b.querySelectorAll('[data-del]').forEach(x => x.onclick = () => { this[x.dataset.p].splice(+x.dataset.del, 1); this.draw(); });
    b.querySelectorAll('[data-frame]').forEach(x => x.onclick = () => this.frame(x.dataset.p, +x.dataset.frame));
    b.querySelectorAll('[data-add]').forEach(x => { const go = () => { this.target = x.dataset.p; if (x.dataset.add === 'doc') return this.fromDoc(); ({ cam: ckCam, img: ckImg, file: ckFile })[x.dataset.add].click(); };
      x.onclick = go; x.onkeydown = e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); go(); } }; });
    f.classList.remove('hidden');
    f.innerHTML = '<span class="hint" style="margin-inline-end:auto"></span><button class="btn" type="button" data-a="merge">רק לרכז לקובץ אחד</button><button class="btn primary" type="button" data-a="check">בדיקת העבודה</button>';
    f.querySelector('[data-a=merge]').onclick = () => this.go(false);
    f.querySelector('[data-a=check]').onclick = () => this.go(true);
    this.btns();
  },
  btns() { const f = this.foot(), n = this.work.length; const c = f.querySelector('[data-a=check]'), m = f.querySelector('[data-a=merge]'); if (!c) return;
    c.disabled = !n; m.disabled = !n && !this.task.length;
    const t = this.task.length + (this.taskText.trim() ? 1 : 0); f.querySelector('.hint').textContent = n ? `עבודה: ${n} ${n === 1 ? 'עמוד' : 'עמודים'}${t ? ' · יש משימה' : ''}` : 'מוסיפים את דפי העבודה כדי להתחיל'; },
  wait(msg) { this.body().innerHTML = `<p class="ai-wait" style="font-size:15px">${esc(msg)}</p>`; this.foot().innerHTML = '<span class="hint" style="margin-inline-end:auto"></span>'; },
  room(part) { return Math.max(0, CK_MAX - this[part].length); },
  async addImages(files, cam) {
    const part = this.target, list = this[part], take = files.slice(0, this.room(part)), bad = [];
    for (const [i, f] of take.entries()) {
      this.wait(take.length > 1 ? `מכין את הדפים… (${i + 1} מתוך ${take.length})` : 'מזהה את הדף ומכין אותו…');
      try { const it = await imgToJpeg(f); it.label = ''; list.push(it); } catch (e) { bad.push(f.name); }
    }
    if (cam && take.length === 1 && !bad.length) { this.frame(part, list.length - 1); return; }
    this.draw(); if (bad.length) this.note('לא הצלחתי לקרוא: ' + bad.join(', '));
  },
  async addFiles(files) {
    const part = this.target, list = this[part], bad = [];
    for (const f of files) {
      if (!this.room(part)) { this.note(`אפשר עד ${CK_MAX} עמודים בכל חלק`); break; }
      try {
        if (/^image\//.test(f.type)) { this.wait('מכין את התמונה…'); const it = await imgToJpeg(f); it.label = f.name; list.push(it); continue; }
        let bytes = new Uint8Array(await f.arrayBuffer()), wordText = null;
        if (/\.docx$/i.test(f.name)) {
          this.wait(`מכין את "${f.name}"…`); await loadEditLibs('docx');
          try { const z = await JSZip.loadAsync(bytes); const x = await z.file('word/document.xml').async('string');
            wordText = x.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => (p.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || []).map(t => t.replace(/<[^>]+>/g, '')).join('') + '\n').replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').trim(); } catch (_) {}
          const r = await OFFICE_EDIT.toPdf('docx', bytes, m => this.wait(m)); bytes = r.pdf instanceof Uint8Array ? r.pdf : new Uint8Array(r.pdf);
        }
        const d = await pdfjs.getDocument({ data: bytes, isEvalSupported: false }).promise;
        const n = Math.min(this.room(part), d.numPages);
        for (let i = 1; i <= n; i++) {
          this.wait(`מכין את "${f.name}": עמוד ${i} מתוך ${n}…`);
          const pg = await d.getPage(i);
          let text = '';
          if (wordText != null) text = i === 1 ? wordText : '(המשך)';
          else { try { const tc = await pg.getTextContent(); text = tc.items.map(t => t.str + (t.hasEOL ? '\n' : '')).join(''); } catch (_) {} }
          const it = await pageToJpeg(pg); it.label = `${f.name.replace(/\.[a-z0-9]+$/i, '')}${n > 1 ? ' · ' + i : ''}`;
          if (text.replace(/\s/g, '').length > 30) it.text = text.trim();
          list.push(it);
        }
        d.destroy();
      } catch (e) { console.warn(e); bad.push(f.name); }
    }
    this.draw(); if (bad.length) this.note('לא הצלחתי לפתוח: ' + bad.join(', '));
  },
  async fromDoc() {
    if (!S.doc) return; const list = this.work, n = Math.min(this.room('work'), S.pages.length);
    for (let i = 0; i < n; i++) {
      this.wait(`מכין את הקובץ הפתוח: עמוד ${i + 1} מתוך ${n}…`);
      const P = S.pages[i]; const it = await pageToJpeg(P.page); let t = ''; try { t = await pageText(i); } catch (_) {}
      it.label = 'הקובץ הפתוח · ' + (i + 1); it.fromDoc = true; if (t && t.replace(/\s/g, '').length > 30) it.text = t; list.push(it);
    }
    this.draw();
  },
  // the frame of a photo (four corners) — the same tool as in the scanner
  frame(part, i) {
    const it = this[part][i]; if (!it) return this.draw();
    SC.tray = [it]; const back = () => { SC.tray = []; this.draw(); };
    // the scanner draws its frame inside its own dialog: here it is drawn in ours
    const keep = SC.body, keepF = SC.foot, keepShow = SC.showTray;
    SC.body = () => this.body(); SC.foot = () => this.foot();
    SC.showTray = () => { SC.body = keep; SC.foot = keepF; SC.showTray = keepShow; back(); };
    SC.frame(0);
  },
  // reads every page that has no text yet; returns false when stopped
  async readAll(pages) {
    const todo = pages.filter(p => !p.text && !p.read);
    if (!todo.length) return true;
    this.busy = true; this.stop = false; this.ctl = new AbortController();
    let wake = null; try { if (navigator.wakeLock) wake = await navigator.wakeLock.request('screen'); } catch (_) {}
    const b = this.body();
    b.innerHTML = `<p class="sc-status" role="status" style="margin:0 0 14px;font-size:15px">קורא את הדפים…</p><ul class="sc-prog">${todo.map((p, i) => `<li><img alt="עמוד ${i + 1}" src="${p.url}"><span>ממתין</span></li>`).join('')}</ul>`;
    this.foot().innerHTML = '<span class="hint" style="margin-inline-end:auto"></span><button class="btn" type="button" data-a="stop">עצירה</button>';
    this.foot().querySelector('[data-a=stop]').onclick = () => { this.stop = true; this.ctl.abort(); };
    const st = b.querySelector('.sc-status'), lis = [...b.querySelectorAll('.sc-prog li')], t00 = Date.now();
    let ok = true;
    for (const [i, p] of todo.entries()) {
      if (this.stop) { ok = false; break; }
      const li = lis[i], sp = li.querySelector('span'); li.className = 'now'; sp.textContent = 'קורא…';
      const t0 = Date.now(); let hand = false;
      const clock = setInterval(() => { sp.textContent = `${hand ? 'כתב יד: קורא בזהירות…' : 'קורא…'} ${Math.round((Date.now() - t0) / 1000)} שנ׳`; st.textContent = `קורא את הדפים: ${i} מתוך ${todo.length} מוכנים · ${Math.round((Date.now() - t00) / 1000)} שניות`; }, 1000);
      try {
        const raw = await readPage(p, SC.hint, SC.quality === 'fast' ? 'auto' : SC.quality, { signal: this.ctl.signal, onHand: () => { hand = true; }, onWait: s => { st.textContent = `השירות עמוס, ממשיך בעוד ${s} שניות…`; }, onLines: (d, n) => { sp.textContent = `שורה ${d} מתוך ${n}`; } });
        const bl = parseBlocks(raw); p.text = bl.map(blockText).join('\n'); p.read = true;
        li.className = 'ok'; sp.textContent = p.text ? `נקרא · ${Math.round((Date.now() - t0) / 1000)} שנ׳` : 'עמוד ריק';
      } catch (e) {
        if (this.stop) { ok = false; break; }
        p.read = true; p.text = ''; li.className = 'bad'; sp.textContent = 'לא נקרא';
        if (/מכסה|לא הוקם|חוסם|אין חיבור/.test(e.message || '')) { st.textContent = e.message; ok = false; break; }
      } finally { clearInterval(clock); }
      if (i < todo.length - 1 && hand) await sleep(1200);
    }
    try { if (wake) await wake.release(); } catch (_) {}
    this.busy = false;
    return ok;
  },
  // the pages in their order in the one file, with where each part starts
  layout() {
    const task = this.order === 'only' ? [] : this.task;
    const pages = this.order === 'work' ? [...this.work, ...task] : [...task, ...this.work];
    const workStart = this.order === 'work' ? 1 : task.length + 1;
    return { pages, task, workStart };
  },
  async buildPdf(pages) {
    const pdf = await PDFLib.PDFDocument.create();
    for (const p of pages) {
      const img = await pdf.embedJpg(Uint8Array.from(atob(p.b64), c => c.charCodeAt(0)));
      const W = 595.28, H = Math.min(1400, W * p.h / p.w);
      const pg = pdf.addPage([W, H]); pg.drawImage(img, { x: 0, y: 0, width: W, height: H });
    }
    return pdf.save();
  },
  async go(check) {
    if (this.busy) return;
    if (check && !this.work.length) { this.note('מוסיפים קודם את דפי העבודה'); return; }
    const all = [...this.work, ...this.task];
    if (!all.length) { this.note('אין עדיין דפים'); return; }
    // for a check every page is read first (the task too); gathering only keeps the pages as they are
    if (check && !(await this.readAll(all))) { this.draw(); this.note(this.stop ? 'הקריאה נעצרה' : 'לא הצלחתי לקרוא את כל הדפים'); return; }
    this.wait('מאחד את הדפים לקובץ אחד…');
    const { pages, task, workStart } = this.layout();
    let bytes; try { bytes = await this.buildPdf(pages); } catch (e) { console.error(e); this.draw(); this.note('איחוד הדפים נכשל: ' + (e.message || e)); return; }
    const d = new Date(), z = n => String(n).padStart(2, '0');
    const name = cleanBase(`${check ? 'בדיקת עבודה' : 'קובץ מאוחד'}${this.who.trim() ? ' - ' + this.who.trim() : ''} ${z(d.getDate())}-${z(d.getMonth() + 1)} ${z(d.getHours())}-${z(d.getMinutes())}`) + '.pdf';
    // the text of every page of the one file, for the check and for every later question about it
    const full = pages.map((p, i) => `===== עמוד ${i + 1} =====\n${p.text || ''}`).join('\n\n');
    const workText = this.work.map((p, i) => `===== עמוד ${workStart + i} =====\n${p.text || ''}`).join('\n\n');
    const pasted = this.taskText.trim();
    const taskText = [task.length ? task.map((p, i) => `(עמוד ${(this.order === 'work' ? this.work.length : 0) + i + 1} בקובץ)\n${p.text || ''}`).join('\n\n') : this.order === 'only' && this.task.length ? this.task.map(p => p.text || '').join('\n\n') : '', pasted].filter(Boolean).join('\n\n');
    const o = { level: this.level, grade: this.grade, task: !!taskText.trim() };
    this.work = []; this.task = []; this.taskText = '';
    ckDlg.close();
    await openFile(new File([bytes], name, { type: 'application/pdf' }));
    // the pages are pictures: the text read from them is kept for this file, so the assistant can answer about it
    if (S.doc) textCache.set(S.doc, { sig: S.pages.map(P => P.ocr ? 1 : 0).join('') + '|' + S.pages.length, text: full });
    if (!check) { toast(`הקובץ המאוחד נפתח: ${pages.length} עמודים`); return; }
    await sleep(300);
    askQ(STUDY.check(o), 'בדיקת עבודה' + (o.task ? ' מול המשימה' : ''), 'check', { text: workText, extra: taskText, taskName: o.task ? 'המשימה' : '' });
  },
};
ckDlg.querySelector('[data-a=close]').onclick = () => CK.close();
ckDlg.addEventListener('cancel', e => { if (CK.busy || CK.work.length || CK.task.length) { e.preventDefault(); CK.close(); } });
ckCam.onchange = e => { const f = [...e.target.files]; e.target.value = ''; CK.addImages(f, true); };
ckImg.onchange = e => { const f = [...e.target.files]; e.target.value = ''; CK.addImages(f); };
ckFile.onchange = e => { const f = [...e.target.files]; e.target.value = ''; CK.addFiles(f); };
window.SIMAN_CHECK = CK;

/* =====================================================================
   6.7.2: CHECKING A WHOLE CLASS — like a chat with a teacher's assistant:
   1. the exam and its answer key (photographed, scanned, files or typed) → a points table the teacher approves
   2. student after student: name + photos of the pages → "next student" (the checking runs in the background,
      so the next one can be photographed at once)
   3. "that was the last one" → one file with every grade, the class summary and, for every student,
      what earned points, what lost points and why. Word, PDF or Excel.
   The work in progress is kept on the device (without the photos), so closing the app loses nothing.
   ===================================================================== */
const grDlg = h(`<dialog class="scan ckd grd" id="grDlg" aria-label="בדיקת מבחנים לכיתה">
  <div class="sc-head"><h2>🏫 בדיקת מבחנים לכיתה</h2><button class="icon-btn" type="button" data-a="close" aria-label="סגירה">×</button></div>
  <div class="sc-body"></div><div class="sc-foot"></div></dialog>`);
document.body.append(grDlg);
const grCam = h('<input type="file" accept="image/*" capture="environment" hidden>');
const grImg = h('<input type="file" accept="image/*" multiple hidden>');
const grFile = h('<input type="file" accept="application/pdf,.pdf,.docx,image/*" multiple hidden>');
grDlg.append(grCam, grImg, grFile);
const GR_KEY = 'siman.grade.v1';
const jsonOf = a => { const s = String(a || ''); const i = s.indexOf('{'), j = s.lastIndexOf('}'); if (i < 0 || j <= i) return null; try { return JSON.parse(s.slice(i, j + 1)); } catch (_) { return null; } };
const GR_SCHEME = total => `From the exam and the answer key make the points table for checking it.
Use the points written in the exam or in the key. If there are none, split ${total} points between the questions by their weight and difficulty. A question with parts (א, ב, ג or a, b, c) becomes one row per part, like "2א".
Output ONLY JSON, in the language of the exam:
{"title":"the exam's name","total":${total},"questions":[{"q":"1","max":10,"topic":"what the question asks, in a few words","expect":"the correct answer or what must appear in it, taken from the key"}]}`;
const GR_GRADE = `Check the student's answers (the last part of the document) against the exam, the answer key and the points table.
For EVERY row of the points table give the points earned (from 0 to the row's max; halves are allowed), what earned the points, and what lost points and why.
Be exact and fair, and consistent: the same mistake loses the same points for every student. Partial credit for a right method with a small slip. A question with no answer gets 0 with "לא נענתה". If a word could not be read, say so instead of guessing against the student.
Output ONLY JSON, in the language of the exam:
{"questions":[{"q":"1","max":10,"got":7,"plus":"what was right","minus":"what was wrong and why the points were taken (empty if full marks)"}],"summary":"two sentences for the teacher about this student","unclear":"anything that was hard to read, or empty"}`;
const GR = {
  s: null, part: 'exam', queue: [], working: false, draft: { name: '', pages: [] },
  body: () => grDlg.querySelector('.sc-body'), foot: () => grDlg.querySelector('.sc-foot'),
  fresh() { return { id: Date.now(), name: '', level: 'a school student', total: 100, exam: { pages: [], text: '' }, key: { pages: [], text: '' }, scheme: null, students: [], step: 'setup' }; },
  keep() { try { const s = this.s; if (!s) return; const strip = p => ({ text: p.text || '', read: !!p.read });
    localStorage.setItem(GR_KEY, JSON.stringify({ ...s, exam: { text: s.exam.text, pages: s.exam.pages.map(strip) }, key: { text: s.key.text, pages: s.key.pages.map(strip) },
      students: s.students.map(st => ({ ...st, pages: [], status: st.status === 'done' || st.status === 'err' ? st.status : 'lost' })) })); } catch (_) {} },
  open() {
    if (typeof needReg === 'function' && needReg()) { showReg(); return; }
    if (!this.s) { let old = null; try { old = JSON.parse(localStorage.getItem(GR_KEY) || 'null'); } catch (_) {}
      if (old && old.students && (old.students.length || old.scheme) && confirm(`להמשיך את הבדיקה של "${old.name || 'המבחן'}"${old.students.length ? ` (${old.students.length} תלמידים)` : ''}?\n\nביטול = להתחיל בדיקה חדשה`)) { this.s = old; }
      else this.s = this.fresh(); }
    if (!grDlg.open) { try { grDlg.showModal(); } catch (_) { grDlg.setAttribute('open', ''); } }
    this.draw();
  },
  close() { grDlg.close(); if (this.working) toast('הבדיקה ממשיכה ברקע. פותחים שוב את "בדיקת מבחנים לכיתה" כדי לראות'); },
  draw() { const st = this.s.step; if (st === 'setup') this.setup(); else if (st === 'scheme') this.schemeView(); else if (st === 'students') this.studentsView(); else this.report(); },
  wait(msg) { this.body().innerHTML = `<p class="ai-wait" style="font-size:15px">${esc(msg)}</p>`; this.foot().innerHTML = '<span class="hint" style="margin-inline-end:auto"></span>'; },
  note(m) { const n = this.foot().querySelector('.hint'); if (n) n.textContent = m; else toast(m); },
  tray(list, part) {
    const add = list.length < 30 ? `<li class="add" role="button" tabindex="0" data-add="cam" data-p="${part}">${ICON.scan}צילום עמוד</li><li class="add" role="button" tabindex="0" data-add="img" data-p="${part}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"><rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="9" cy="10" r="1.6"/><path d="M21 16l-5-5-8 8"/></svg>תמונות</li><li class="add" role="button" tabindex="0" data-add="file" data-p="${part}"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"><path d="M7 3h7l4 4v14H7z"/><path d="M14 3v4h4"/></svg>קבצים</li>` : '';
    return `<ul class="sc-tray">${list.map((p, i) => `<li>${p.url ? `<img alt="עמוד ${i + 1}" src="${p.url}">` : '<div class="gr-nopic">✓ נקרא</div>'}עמוד ${i + 1}<button type="button" class="x" data-del="${i}" data-p="${part}" aria-label="הסרת עמוד ${i + 1}">×</button></li>`).join('')}${add}</ul>`;
  },
  wire(b) {
    b.querySelectorAll('[data-del]').forEach(x => x.onclick = () => { this.pagesOf(x.dataset.p).splice(+x.dataset.del, 1); this.draw(); });
    b.querySelectorAll('[data-add]').forEach(x => { const go = () => { this.part = x.dataset.p; ({ cam: grCam, img: grImg, file: grFile })[x.dataset.add].click(); };
      x.onclick = go; x.onkeydown = e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); go(); } }; });
  },
  pagesOf(part) { return part === 'exam' ? this.s.exam.pages : part === 'key' ? this.s.key.pages : this.draft.pages; },
  // photos and files become pages (the same tools as one work check)
  async add(files, cam) {
    const list = this.pagesOf(this.part), back = () => this.draw();
    CK.target = '__gr'; CK.__gr = list;
    const keepDraw = CK.draw, keepWait = CK.wait, keepBody = CK.body, keepFoot = CK.foot;
    CK.draw = back; CK.wait = m => this.wait(m); CK.body = () => this.body(); CK.foot = () => this.foot();
    try { if (cam !== undefined) await CK.addImages(files, false); else await CK.addFiles(files); }
    finally { CK.draw = keepDraw; CK.wait = keepWait; CK.body = keepBody; CK.foot = keepFoot; delete CK.__gr; CK.target = 'work'; }
    this.draw();
  },
  setup() {
    const s = this.s, b = this.body(), f = this.foot();
    b.innerHTML = `<p class="gr-steps"><b>1. המבחן והמחוון</b> › 2. חלוקת הנקודות › 3. התלמידים › 4. קובץ הציונים</p>
      <div class="ck-sec ck-opts"><label>שם המבחן והכיתה <input type="text" data-o="name" placeholder="למשל: מבחן בחשבון – ח'2" dir="auto"></label>
        <label>הכיתה <select data-o="level"><option value="a school student">בית ספר</option><option value="a university or college student">סטודנטים</option></select></label>
        <label>סך הנקודות <input type="number" data-o="total" min="1" max="1000" dir="ltr"></label></div>
      <div class="ck-sec"><div class="sc-trayhead"><h3>המבחן</h3><p>דפי השאלות: מצלמים, סורקים או בוחרים קובץ. אפשר גם להקליד או להדביק.</p></div>${this.tray(s.exam.pages, 'exam')}<textarea class="ck-txt" data-t="exam" placeholder="או מדביקים כאן את שאלות המבחן…"></textarea></div>
      <div class="ck-sec"><div class="sc-trayhead"><h3>המחוון או הפתרון</h3><p>הפתרון הנכון והניקוד. בלי מחוון הבדיקה נעשית לפי מה שנכון, אבל עם מחוון היא מדויקת ועקבית.</p></div>${this.tray(s.key.pages, 'key')}<textarea class="ck-txt" data-t="key" placeholder="או מדביקים כאן את הפתרון והניקוד…"></textarea></div>`;
    const nm = b.querySelector('[data-o=name]'); nm.value = s.name; nm.oninput = () => { s.name = nm.value; this.keep(); };
    const lv = b.querySelector('[data-o=level]'); lv.value = s.level; lv.onchange = () => { s.level = lv.value; this.keep(); };
    const tt = b.querySelector('[data-o=total]'); tt.value = s.total; tt.oninput = () => { s.total = Math.max(1, +tt.value || 100); this.keep(); };
    b.querySelectorAll('[data-t]').forEach(t => { t.value = s[t.dataset.t].text; t.oninput = () => { s[t.dataset.t].text = t.value; this.keep(); }; });
    this.wire(b);
    f.innerHTML = `<span class="hint" style="margin-inline-end:auto"></span>${s.scheme ? '<button class="btn" type="button" data-a="skip">לטבלת הנקודות הקיימת</button>' : ''}<button class="btn primary" type="button" data-a="go">המשך: חלוקת הנקודות</button>`;
    const sk = f.querySelector('[data-a=skip]'); if (sk) sk.onclick = () => { s.step = 'scheme'; this.draw(); };
    f.querySelector('[data-a=go]').onclick = () => this.makeScheme();
  },
  textOf(part) { const x = this.s[part]; return [x.pages.map(p => p.text || '').join('\n\n'), x.text].filter(t => t && t.trim()).join('\n\n').trim(); },
  async readPages(pages, onStep) {
    for (const [i, p] of pages.entries()) {
      if (p.text || p.read) continue;
      onStep && onStep(i, pages.length);
      const raw = await readPage(p, SC.hint, ['auto', 'accurate', 'print'].includes(SC.quality) ? SC.quality : 'auto', {});
      p.text = parseBlocks(raw).map(blockText).join('\n'); p.read = true;
    }
  },
  async makeScheme() {
    const s = this.s;
    if (!s.exam.pages.length && !s.exam.text.trim()) { this.note('מוסיפים קודם את המבחן'); return; }
    try {
      await this.readPages([...s.exam.pages, ...s.key.pages], (i, n) => this.wait(`קורא את דפי המבחן והמחוון… (${i + 1} מתוך ${n})`));
      this.wait('מכין את טבלת הנקודות מהמבחן ומהמחוון…');
      const ex = this.textOf('exam'), key = this.textOf('key');
      if (!ex.replace(/\s/g, '')) throw new Error('לא הצלחתי לקרוא את המבחן. נסו לצלם שוב באור טוב, או להדביק את השאלות');
      const a = await ai({ mode: 'ask', text: `===== המבחן =====\n${ex}\n\n===== המחוון / הפתרון =====\n${key || '(אין מחוון: לבדוק לפי מה שנכון)'}`, question: GR_SCHEME(s.total), history: [] });
      const j = jsonOf(a); if (!j || !Array.isArray(j.questions) || !j.questions.length) throw new Error('לא הצלחתי לבנות את טבלת הנקודות. נסו שוב');
      s.scheme = j.questions.map(q => ({ q: String(q.q ?? '').slice(0, 12), max: Math.max(0, +q.max || 0), topic: String(q.topic || '').slice(0, 200), expect: String(q.expect || '').slice(0, 600) }));
      if (!s.name && j.title) s.name = String(j.title).slice(0, 80);
      s.step = 'scheme'; this.keep(); this.draw();
    } catch (e) { this.setup(); this.note(e.message || String(e)); }
  },
  schemeView() {
    const s = this.s, b = this.body(), f = this.foot();
    const sum = () => s.scheme.reduce((t, q) => t + (+q.max || 0), 0);
    b.innerHTML = `<p class="gr-steps">1. המבחן והמחוון › <b>2. חלוקת הנקודות</b> › 3. התלמידים › 4. קובץ הציונים</p>
      <p class="gr-help">זו החלוקה שמצאתי במבחן ובמחוון. אפשר לתקן ניקוד או את מה שנדרש בכל שאלה. ככה ייבדק כל תלמיד, באותה מידה.</p>
      <div class="gr-tbl" role="table"><div class="gr-row gr-h" role="row"><span>שאלה</span><span>ניקוד</span><span>מה נדרש</span><span></span></div>
      ${s.scheme.map((q, i) => `<div class="gr-row" role="row" data-i="${i}"><input data-k="q" value="${esc(q.q)}" aria-label="מספר השאלה" dir="auto"><input data-k="max" type="number" min="0" step="0.5" value="${q.max}" aria-label="ניקוד" dir="ltr"><textarea data-k="expect" rows="2" aria-label="מה נדרש" dir="auto">${esc(q.expect || q.topic)}</textarea><button type="button" class="gr-del" data-del="${i}" aria-label="מחיקת השורה">×</button></div>`).join('')}</div>
      <button class="btn" type="button" data-a="row">+ שאלה</button> <span class="gr-sum"></span>`;
    const upd = () => { const n = sum(); b.querySelector('.gr-sum').textContent = `סך הכול: ${n} נקודות${n !== s.total ? ` (המבחן מתוך ${s.total})` : ''}`; };
    b.querySelectorAll('.gr-row[data-i]').forEach(r => { const q = s.scheme[+r.dataset.i];
      r.querySelectorAll('[data-k]').forEach(inp => inp.oninput = () => { q[inp.dataset.k] = inp.dataset.k === 'max' ? Math.max(0, +inp.value || 0) : inp.value; upd(); this.keep(); }); });
    b.querySelectorAll('[data-del]').forEach(x => x.onclick = () => { s.scheme.splice(+x.dataset.del, 1); this.keep(); this.draw(); });
    b.querySelector('[data-a=row]').onclick = () => { s.scheme.push({ q: String(s.scheme.length + 1), max: 0, topic: '', expect: '' }); this.keep(); this.draw(); };
    upd();
    f.innerHTML = '<span class="hint" style="margin-inline-end:auto"></span><button class="btn" type="button" data-a="back">חזרה למבחן</button><button class="btn primary" type="button" data-a="go">התחלת בדיקת התלמידים</button>';
    f.querySelector('[data-a=back]').onclick = () => { s.step = 'setup'; this.draw(); };
    f.querySelector('[data-a=go]').onclick = () => { if (!s.scheme.length) { this.note('אין שאלות בטבלה'); return; } s.step = 'students'; this.keep(); this.draw(); };
  },
  studentsView() {
    const s = this.s, b = this.body(), f = this.foot(), d = this.draft;
    const n = s.students.length, done = s.students.filter(x => x.status === 'done').length;
    b.innerHTML = `<p class="gr-steps">1. המבחן והמחוון › 2. חלוקת הנקודות › <b>3. התלמידים</b> › 4. קובץ הציונים</p>
      <div class="gr-cur"><h3>תלמיד ${n + 1}</h3>
        <label class="gr-name">שם התלמיד או התלמידה <input type="text" data-o="sname" placeholder="למשל: דנה לוי" dir="auto"></label>
        <p class="gr-help">מצלמים את כל דפי התשובות של התלמיד, דף אחרי דף. אחר כך לוחצים "לבדוק ולעבור לתלמיד הבא": הבדיקה רצה ברקע, ואפשר לצלם מיד את התלמיד הבא.</p>
        ${this.tray(d.pages, 'draft')}</div>
      <div class="gr-list"><h3>התלמידים שנבדקו${n ? ` (${done} מתוך ${n} מוכנים)` : ''}</h3>${n ? '' : '<p class="gr-help">עוד אין. התלמיד הראשון יופיע כאן.</p>'}<ol>${s.students.map((st, i) => this.rowHtml(st, i)).join('')}</ol></div>`;
    const nm = b.querySelector('[data-o=sname]'); nm.value = d.name; nm.oninput = () => { d.name = nm.value; };
    this.wire(b);
    b.querySelectorAll('[data-again]').forEach(x => x.onclick = () => { const st = s.students[+x.dataset.again]; if (!st.pages.length && !st.text) { toast('הצילומים של התלמיד הזה כבר לא שמורים. מצלמים אותו שוב כתלמיד חדש'); return; } st.status = 'wait'; st.err = ''; this.queue.push(st); this.pump(); this.draw(); });
    b.querySelectorAll('[data-open]').forEach(x => x.onclick = () => { const li = x.closest('li'); li.classList.toggle('open'); });
    b.querySelectorAll('[data-pts]').forEach(inp => inp.onchange = () => { const [i, k] = inp.dataset.pts.split(':').map(Number); const q = s.students[i].res.questions[k]; q.got = Math.max(0, Math.min(+q.max || 0, +inp.value || 0)); this.total(s.students[i]); this.keep(); this.draw(); });
    b.querySelectorAll('[data-rm]').forEach(x => x.onclick = () => { const i = +x.dataset.rm; if (!confirm(`להסיר את "${s.students[i].name}" מהרשימה?`)) return; s.students.splice(i, 1); this.keep(); this.draw(); });
    f.innerHTML = `<span class="hint" style="margin-inline-end:auto"></span><button class="btn" type="button" data-a="scheme">טבלת הנקודות</button><button class="btn" type="button" data-a="next" ${d.pages.length ? '' : 'disabled'}>לבדוק ולעבור לתלמיד הבא</button><button class="btn primary" type="button" data-a="last" ${n || d.pages.length ? '' : 'disabled'}>זה היה האחרון: קובץ ציונים</button>`;
    f.querySelector('[data-a=scheme]').onclick = () => { s.step = 'scheme'; this.draw(); };
    f.querySelector('[data-a=next]').onclick = () => this.next();
    f.querySelector('[data-a=last]').onclick = () => { if (d.pages.length) this.next(); s.step = 'report'; this.keep(); this.draw(); };
    this.note(this.working ? `בודק ברקע… ${this.queue.length ? `(${this.queue.length} ממתינים)` : ''}` : d.pages.length ? `${d.pages.length} עמודים לתלמיד הזה` : 'מצלמים את דפי התלמיד');
  },
  rowHtml(st, i) {
    const r = st.res, stt = { wait: 'ממתין בתור', read: 'קורא את הדפים…', check: 'בודק…', done: '', err: 'הבדיקה נכשלה', lost: 'לא הסתיים (האפליקציה נסגרה)' }[st.status] || '';
    const g = st.status === 'done' ? `<b class="gr-g">${fmtN(st.grade)}</b>` : `<span class="gr-st${st.status === 'err' || st.status === 'lost' ? ' bad' : ''}">${stt}</span>`;
    const det = r ? `<div class="gr-det">${r.summary ? `<p>${esc(r.summary)}</p>` : ''}${r.unclear ? `<p class="gr-unc">לא היה ברור לקריאה: ${esc(r.unclear)}</p>` : ''}
      <table><thead><tr><th>שאלה</th><th>נקודות</th><th>מה היה נכון</th><th>למה ירדו נקודות</th></tr></thead><tbody>${r.questions.map((q, k) => `<tr><td>${esc(q.q)}</td><td dir="ltr"><input type="number" min="0" max="${q.max}" step="0.5" value="${q.got}" data-pts="${i}:${k}" aria-label="נקודות בשאלה ${esc(q.q)}"> / ${q.max}</td><td>${esc(q.plus || '')}</td><td>${esc(q.minus || '')}</td></tr>`).join('')}</tbody></table></div>` : '';
    return `<li class="${st.status}"><div class="gr-li"><button type="button" class="gr-nm" data-open="${i}" ${r ? '' : 'disabled'}>${esc(st.name || 'תלמיד ' + (i + 1))}</button>${g}${st.status === 'err' || st.status === 'lost' ? `<button type="button" class="btn" data-again="${i}">לנסות שוב</button>` : ''}<button type="button" class="gr-del" data-rm="${i}" aria-label="הסרה">×</button></div>${st.err ? `<p class="gr-unc">${esc(st.err)}</p>` : ''}${det}</li>`;
  },
  total(st) { const qs = st.res.questions; const got = qs.reduce((t, q) => t + (+q.got || 0), 0), max = qs.reduce((t, q) => t + (+q.max || 0), 0) || 1;
    st.points = got; st.max = max; st.grade = Math.round(got / max * this.s.total * 10) / 10; },
  next() {
    const d = this.draft, s = this.s; if (!d.pages.length) { this.note('מצלמים קודם את דפי התלמיד'); return; }
    const st = { name: d.name.trim() || `תלמיד ${s.students.length + 1}`, pages: d.pages, status: 'wait', res: null };
    s.students.push(st); this.queue.push(st); this.draft = { name: '', pages: [] };
    this.keep(); this.pump(); this.draw();
    toast(`"${st.name}" נשלח לבדיקה. אפשר לצלם את התלמיד הבא`);
    const nm = this.body().querySelector('[data-o=sname]'); if (nm && matchMedia('(pointer:fine)').matches) nm.focus();
  },
  // the students are checked one after another, in the background
  async pump() {
    if (this.working) return; this.working = true;
    let wake = null; try { if (navigator.wakeLock) wake = await navigator.wakeLock.request('screen'); } catch (_) {}
    try {
      while (this.queue.length) {
        const st = this.queue.shift(), s = this.s, redraw = () => { if (grDlg.open && s.step === 'students') this.draw(); };
        try {
          st.status = 'read'; redraw();
          await this.readPages(st.pages);
          st.text = st.pages.map((p, i) => `===== עמוד ${i + 1} =====\n${p.text || ''}`).join('\n\n');
          if (!st.pages.some(p => (p.text || '').replace(/\s/g, '').length > 3)) throw new Error('לא הצלחתי לקרוא את הדפים. נסו לצלם שוב, ישר ובאור טוב');
          st.status = 'check'; redraw();
          const table = s.scheme.map(q => `${q.q} | ${q.max} נקודות | ${q.expect || q.topic}`).join('\n');
          const a = await ai({ mode: 'ask', text: `===== המבחן =====\n${this.textOf('exam')}\n\n===== המחוון / הפתרון =====\n${this.textOf('key') || '(אין מחוון)'}\n\n===== טבלת הנקודות (שאלה | ניקוד מרבי | מה נדרש) =====\n${table}\n\n===== התשובות של ${st.name} =====\n${st.text}`, question: GR_GRADE + `\nThe writer is ${s.level}.`, history: [] });
          const j = jsonOf(a); if (!j || !Array.isArray(j.questions)) throw new Error('התשובה מהשירות לא הייתה תקינה');
          // every row of the table, in its order, even if the answer skipped one
          const by = new Map(j.questions.map(q => [String(q.q).trim(), q]));
          st.res = { summary: String(j.summary || ''), unclear: String(j.unclear || ''), questions: s.scheme.map((q, k) => { const g = by.get(String(q.q).trim()) || j.questions[k] || {}; return { q: q.q, max: +q.max || 0, got: Math.max(0, Math.min(+q.max || 0, +g.got || 0)), plus: String(g.plus || ''), minus: String(g.minus || (g.got == null ? 'לא נבדקה' : '')) }; }) };
          this.total(st); st.status = 'done'; st.pages = st.pages.map(p => ({ text: p.text, read: true }));
        } catch (e) { st.status = 'err'; st.err = e.message || String(e); if (/מכסה|אין חיבור|חוסם|לא הוקם/.test(st.err)) { for (const x of this.queue) { x.status = 'err'; x.err = st.err; } this.queue = []; } }
        this.keep(); redraw();
      }
    } finally { this.working = false; try { if (wake) await wake.release(); } catch (_) {} if (grDlg.open && this.s.step === 'students') this.draw(); if (grDlg.open && this.s.step === 'report') this.draw(); }
  },
  stats() {
    const s = this.s, ok = s.students.filter(x => x.status === 'done');
    const avg = ok.length ? ok.reduce((t, x) => t + x.grade, 0) / ok.length : 0;
    const perQ = s.scheme.map((q, k) => { const v = ok.map(x => +x.res.questions[k].got || 0); return { q: q.q, max: +q.max || 0, avg: v.length ? v.reduce((a, b) => a + b, 0) / v.length : 0 }; });
    return { ok, avg, hi: ok.length ? Math.max(...ok.map(x => x.grade)) : 0, lo: ok.length ? Math.min(...ok.map(x => x.grade)) : 0, perQ };
  },
  report() {
    const s = this.s, b = this.body(), f = this.foot(), S2 = this.stats();
    const left = s.students.filter(x => x.status !== 'done');
    b.innerHTML = `<p class="gr-steps">1. המבחן והמחוון › 2. חלוקת הנקודות › 3. התלמידים › <b>4. קובץ הציונים</b></p>
      ${this.working ? `<p class="ai-wait">עוד בודק ${s.students.filter(x => ['wait', 'read', 'check'].includes(x.status)).length} תלמידים… הקובץ יהיה מוכן כשהם יסתיימו.</p>` : ''}
      ${left.length && !this.working ? `<p class="gr-unc">${left.length} תלמידים לא נבדקו (${left.map(x => esc(x.name)).join(', ')}). אפשר לחזור לרשימה ולנסות שוב, או להפיק את הקובץ בלעדיהם.</p>` : ''}
      <div class="gr-sumbox"><div><b>${S2.ok.length}</b><span>תלמידים</span></div><div><b>${fmtN(S2.avg)}</b><span>ממוצע</span></div><div><b>${fmtN(S2.hi)}</b><span>הכי גבוה</span></div><div><b>${fmtN(S2.lo)}</b><span>הכי נמוך</span></div></div>
      <table class="gr-final"><thead><tr><th>תלמיד</th><th>ציון</th>${s.scheme.map(q => `<th>${esc(q.q)}<small>/${q.max}</small></th>`).join('')}</tr></thead><tbody>${S2.ok.map(x => `<tr><td>${esc(x.name)}</td><td><b>${fmtN(x.grade)}</b></td>${x.res.questions.map(q => `<td>${fmtN(q.got)}</td>`).join('')}</tr>`).join('')}
      <tr class="avg"><td>ממוצע</td><td>${fmtN(S2.avg)}</td>${S2.perQ.map(q => `<td>${fmtN(q.avg)}</td>`).join('')}</tr></tbody></table>
      <p class="gr-help">בקובץ: טבלת הציונים, סיכום הכיתה (כולל השאלות הקשות), ולכל תלמיד עמוד עם הניקוד בכל שאלה, מה היה נכון ולמה ירדו נקודות.</p>`;
    f.innerHTML = `<span class="hint" style="margin-inline-end:auto"></span><button class="btn" type="button" data-a="back">חזרה לתלמידים</button><button class="btn" type="button" data-a="xlsx" ${S2.ok.length ? '' : 'disabled'}>Excel</button><button class="btn" type="button" data-a="pdf" ${S2.ok.length ? '' : 'disabled'}>PDF</button><button class="btn primary" type="button" data-a="docx" ${S2.ok.length ? '' : 'disabled'}>שמירה כ-Word</button>`;
    f.querySelector('[data-a=back]').onclick = () => { s.step = 'students'; this.keep(); this.draw(); };
    f.querySelectorAll('[data-a=docx],[data-a=pdf],[data-a=xlsx]').forEach(x => x.onclick = () => this.save(x.dataset.a));
    f.insertAdjacentHTML('afterbegin', '');
    const nw = h('<button class="btn" type="button">בדיקה חדשה</button>'); nw.onclick = () => { if (!confirm('להתחיל בדיקה חדשה? הציונים של הבדיקה הזו יימחקו מהמכשיר (קובץ שכבר נשמר נשאר)')) return; localStorage.removeItem(GR_KEY); this.s = this.fresh(); this.queue = []; this.draft = { name: '', pages: [] }; this.draw(); };
    f.querySelector('[data-a=back]').before(nw);
  },
  blocks() {
    const s = this.s, S2 = this.stats(), d = new Date().toLocaleDateString('he-IL');
    const pages = [];
    const hard = [...S2.perQ].filter(q => q.max).sort((a, b) => a.avg / a.max - b.avg / b.max).slice(0, 3);
    // numbers stand in their own table cells and are written with words ("מתוך", "אחוז"): mixed with Hebrew,
    // signs like % / ( ) come out in the wrong place in some programs
    const pct = q => q.max ? Math.round(q.avg / q.max * 100) : 0;
    pages.push([{ t: 'h1', text: `${s.name || 'מבחן'} – ציונים` },
      { t: 'table', rows: [['תאריך', 'תלמידים', 'ממוצע', 'הכי גבוה', 'הכי נמוך'], [d, String(S2.ok.length), fmtN(S2.avg), fmtN(S2.hi), fmtN(S2.lo)]] },
      { t: 'h2', text: 'הציונים' },
      { t: 'table', rows: [['מספר', 'תלמיד', 'ציון', 'נקודות'], ...S2.ok.map((x, i) => [String(i + 1), x.name, fmtN(x.grade), `${fmtN(x.points)} מתוך ${fmtN(x.max)}`])] },
      { t: 'h2', text: 'לפי שאלה' }, { t: 'table', rows: [['שאלה', 'ניקוד', 'ממוצע הכיתה', 'הצלחה'], ...S2.perQ.map(q => [q.q, fmtN(q.max), fmtN(q.avg), q.max ? `${pct(q)} אחוז` : '–'])] },
      ...(hard.length ? [{ t: 'h2', text: 'השאלות הקשות ביותר לכיתה' }, ...hard.map(q => ({ t: 'li', text: `שאלה ${q.q}: הצליחו בה ${pct(q)} אחוז מהכיתה` }))] : [])]);
    for (const x of S2.ok) {
      const pg = [{ t: 'h1', text: `${x.name} – ${fmtN(x.grade)}` }, { t: 'p', text: `${fmtN(x.points)} נקודות מתוך ${fmtN(x.max)}` }, ...(x.res.summary ? [{ t: 'p', text: x.res.summary }] : []),
        { t: 'table', rows: [['שאלה', 'נקודות', 'מה היה נכון', 'למה ירדו נקודות'], ...x.res.questions.map(q => [q.q, `${fmtN(q.got)} מתוך ${fmtN(q.max)}`, q.plus || '–', q.got >= q.max ? '–' : (q.minus || '–')])] }];
      if (x.res.unclear) pg.push({ t: 'p', text: 'לא היה ברור לקריאה: ' + x.res.unclear });
      pages.push(pg);
    }
    return pages;
  },
  async save(kind) {
    const s = this.s, base = cleanBase(`${s.name || 'מבחן'} - ציונים`) || 'ציונים';
    try {
      if (kind === 'xlsx') {
        await loadEditLibs('xlsx'); const S2 = this.stats();
        const head = ['תלמיד', 'ציון', 'נקודות', 'מתוך', ...s.scheme.map(q => `שאלה ${q.q} (${q.max})`)];
        const rows = S2.ok.map(x => [x.name, x.grade, x.points, x.max, ...x.res.questions.map(q => q.got)]);
        rows.push(['ממוצע', +S2.avg.toFixed(2), '', '', ...S2.perQ.map(q => +q.avg.toFixed(2))]);
        const det = [['תלמיד', 'שאלה', 'נקודות', 'מתוך', 'מה היה נכון', 'למה ירדו נקודות']];
        for (const x of S2.ok) for (const q of x.res.questions) det.push([x.name, q.q, q.got, q.max, q.plus, q.got >= q.max ? '' : q.minus]);
        const wb = XLSX.utils.book_new(); const a = XLSX.utils.aoa_to_sheet([head, ...rows]), bsh = XLSX.utils.aoa_to_sheet(det);
        a['!cols'] = [{ wch: 22 }, { wch: 8 }, { wch: 8 }, { wch: 7 }, ...s.scheme.map(() => ({ wch: 11 }))]; bsh['!cols'] = [{ wch: 20 }, { wch: 7 }, { wch: 7 }, { wch: 6 }, { wch: 45 }, { wch: 55 }];
        XLSX.utils.book_append_sheet(wb, a, 'ציונים'); XLSX.utils.book_append_sheet(wb, bsh, 'פירוט');
        wb.Workbook = { Views: [{ RTL: true }] };
        const bytes = new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }));
        if (await saveBytes(base + '.xlsx', bytes, 'xlsx')) toast('נשמר: ' + base + '.xlsx');
        return;
      }
      const pages = this.blocks();
      if (kind === 'pdf') { this.note('בונה PDF…'); const bytes = await blocksPdf(pages); if (await saveBytes(base + '.pdf', bytes, 'pdf')) { toast('נשמר: ' + base + '.pdf'); const o = h('<button class="btn" type="button">פתיחה בסימן</button>'); o.onclick = () => { grDlg.close(); openFile(new File([bytes], base + '.pdf', { type: 'application/pdf' })); }; this.foot().prepend(o); } return; }
      const bytes = blocksDocx(pages); if (await saveBytes(base + '.docx', bytes, 'docx')) toast('נשמר: ' + base + '.docx');
    } catch (e) { console.error(e); this.note('השמירה נכשלה: ' + (e.message || e)); }
  },
};
const fmtN = v => { const n = +v || 0; return Number.isInteger(n) ? String(n) : n.toFixed(1).replace(/\.0$/, ''); };
grDlg.querySelector('[data-a=close]').onclick = () => GR.close();
grDlg.addEventListener('cancel', e => { e.preventDefault(); GR.close(); });
grCam.onchange = e => { const f = [...e.target.files]; e.target.value = ''; GR.add(f, true); };
grImg.onchange = e => { const f = [...e.target.files]; e.target.value = ''; GR.add(f, false); };
grFile.onchange = e => { const f = [...e.target.files]; e.target.value = ''; GR.add(f); };
window.SIMAN_GRADE = GR;

/* The writer's own words: names, places and work terms the person typed when fixing a reading.
   Kept on this device only, and sent with the next readings as a hint (used only when the letters fit). */
const VOCAB = {
  key: 'siman.scan.vocab',
  load() { try { return JSON.parse(localStorage.getItem(this.key) || '{}') || {}; } catch (_) { return {}; } },
  top(n = 150) { const v = this.load(); return Object.keys(v).sort((a, b) => v[b] - v[a]).slice(0, n); },
  words: t => String(t || '').replace(/\[\[|\]\]/g, ' ').split(/[\s,.;:!?()"״׳'\/\\\-–—]+/).map(w => w.trim()).filter(w => w.length >= 2 && /[\u0590-\u05FFA-Za-z]/.test(w) && w !== '?'),
  learn(before, after) {
    const had = new Set(this.words(before)), v = this.load(); let n = 0;
    for (const w of new Set(this.words(after))) if (!had.has(w)) { v[w] = (v[w] || 0) + 1; n++; }
    if (!n) return;
    const keep = Object.keys(v).sort((a, b) => v[b] - v[a]).slice(0, 400), out = {};
    for (const k of keep) out[k] = v[k];
    try { localStorage.setItem(this.key, JSON.stringify(out)); } catch (_) {}
  },
};
window.SIMAN_VOCAB = VOCAB;

/* Reading one page.
   Pages prepared by the scanner (siman-scan.js) come cut into lines: each line is read as its own picture,
   in a few requests of up to 6 lines, two at a time. On the ivrit.ai Hebrew handwriting benchmark this is
   where the reading models are by far the most accurate. Other pages (a PDF, an old reader photo) are read
   as one page, as before. The answer has the same shape either way: {"lang","blocks"}. */
/* 6.6: a printed page (a book, a letter, a form, a screenshot) is read whole, in one quick request, in a few
   seconds — the careful line-by-line reading below was made for handwriting and took minutes on print.
   "auto" (the default): the quick reading also says whether the page is really handwritten; only such a page
   goes on to the careful handwriting reading. An older server that does not know "print" answers "unknown
   mode": then the reading goes on as before. */
let PRINT_OK = true;
async function readPrinted(p, hint, { onWait, signal } = {}) {
  if (!PRINT_OK) return null;
  try {
    const raw = await ai({ mode: 'print', image: p.b64, strips: p.strips || [], mime: 'image/jpeg', hint }, { timeout: 100000, onWait, signal });
    let hand = false;
    try { hand = !!JSON.parse(String(raw).replace(/^\s*```(?:json)?|```\s*$/g, '')).hand; } catch (_) {}
    return { raw, hand };
  } catch (e) {
    if (/unknown mode/i.test(e && e.message || '')) { PRINT_OK = false; return null; }
    throw e;
  }
}
async function readPage(p, hint, quality, opts = {}) {
  if (quality === 'auto' || quality === 'print') {
    const r = await readPrinted(p, hint, opts);
    if (r && (quality === 'print' || !r.hand)) { p.kind = 'print'; return r.raw; }
    if (!r && quality === 'print') return ai({ mode: 'scan', image: p.b64, strips: p.strips || [], mime: 'image/jpeg', hint, quality: 'fast' }, { timeout: 150000, onWait: opts.onWait, signal: opts.signal });
    p.kind = 'hand'; opts.onHand && opts.onHand();
    quality = 'accurate';
  }
  return readHand(p, hint, quality, opts);
}
async function readHand(p, hint, quality, { onWait, onLines, onCheck, signal } = {}) {
  if (!p.lines || !p.lines.length) return ai({ mode: 'scan', image: p.b64, strips: p.strips || [], mime: 'image/jpeg', hint, quality }, { timeout: 150000, onWait, signal });
  // 5.9.6: no internet / the daily limit: stop at once, instead of rows of [?] and more requests that will fail too
  const FATAL = /מכסה|אין חיבור|חוסם|לא הוקם/;
  const L = p.lines, B = 6, got = new Array(L.length).fill(null), starts = [];
  for (let i = 0; i < L.length; i += B) starts.push(i);
  let next = 0, done = 0, firstErr = null, lang = null;
  const strip = t => String(t).replace(/^\s*```(?:json)?|```\s*$/g, '');
  const worker = async () => {
    while (next < starts.length) {
      const s0 = starts[next++], part = L.slice(s0, s0 + B);
      try {
        if (signal && signal.aborted) throw new DOMException('stop', 'AbortError');
        const raw = await ai({ mode: 'lines', lines: part, lines2: (p.lines2 || []).slice(s0, s0 + B), vocab: VOCAB.top(), hint, quality }, { timeout: 130000, onWait, signal });
        const j = JSON.parse(strip(raw));
        if (j.lang && !lang) lang = j.lang;
        (j.lines || []).forEach((r, k) => { const n = Number.isInteger(r.n) && r.n >= 1 && r.n <= part.length ? r.n - 1 : k; if (n < part.length && !got[s0 + n]) got[s0 + n] = r; });
      } catch (e) {
        if ((e && e.name === 'AbortError') || FATAL.test(e && e.message || '')) { next = starts.length; firstErr = e; fatal = e; return; }
        firstErr = firstErr || e;
      }
      done += part.length; onLines && onLines(Math.min(done, L.length), L.length);
    }
  };
  let fatal = null;
  await Promise.all([worker(), worker()]);
  if (fatal) throw fatal;
  if (!got.some(Boolean)) throw firstErr || new Error('לא זוהה טקסט');
  // מבט שני על כל הדף יחד: כל השורות והקריאה הראשונה. מתקן מילים לפי נושא הדף ומסמן בצהוב כל מילה לא בטוחה
  if (quality !== 'fast') {
    onCheck && onCheck();
    const views = p.lines2 && p.lines2.length === L.length ? p.lines2 : L;
    for (let s0 = 0; s0 < L.length; s0 += 15) {
      const idx = [...Array(Math.min(15, L.length - s0)).keys()].map(k => s0 + k);
      if (!idx.some(i => got[i])) continue;
      try {
        if (signal && signal.aborted) break;
        const raw = await ai({ mode: 'linesCheck', lines: idx.map(i => views[i]), draft: idx.map(i => got[i] || { kind: 'text', text: '[[?]]' }), vocab: VOCAB.top(), hint }, { timeout: 100000, onWait, signal });
        const j = JSON.parse(strip(raw));
        const fixed = j.lines || [];
        if (fixed.length === idx.length) fixed.forEach((r, k) => { if (r && typeof r.text === 'string') got[idx[k]] = r; });
      } catch (e) { console.warn('second look', e); }   // the first reading stays
    }
  }
  // שורות → קטעים. שורה שהיא רק חץ מחברת את השורה שלפניה לזו שאחריה (רצף של שלבים)
  const blocks = []; let joinNext = false;
  got.forEach((r, i) => {
    if (!r) { blocks.push({ t: 'p', text: '[[?]]' }); joinNext = false; return; }
    const text = String(r.text || '').replace(/\s+/g, ' ').trim();
    if (r.kind === 'empty' || !text || /^[-–—_.·\s]+$/.test(text)) return;
    if (r.kind === 'arrow' || /^[↓↙↘⬇←→⇣⟵⟶\s]+$/.test(text)) { if (blocks.length) joinNext = true; return; }
    if (joinNext && blocks.length) { blocks[blocks.length - 1].text += ' ← ' + text; joinNext = false; return; }
    joinNext = false;
    // 5.9.6: a line that continues the sentence of the line before is the same paragraph
    // (before, every written line became a paragraph of its own and had to be joined by hand in Word)
    const prev = blocks[blocks.length - 1], t = r.kind === 'list' ? 'li' : r.kind === 'title' && blocks.length < 2 ? 'h2' : 'p';
    if (t === 'p' && prev && prev.t === 'p' && prev.text !== '[[?]]' && r.kind !== 'title' && !/[.!?:;׃…"״)\]]$/.test(prev.text) && !/^[-•*–\d]+[.)]?\s/.test(text) && prev.text.length > 25) { prev.text += ' ' + text; return; }
    blocks.push({ t, text });
  });
  return JSON.stringify({ lang: lang || hint || 'he', blocks });
}

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
// The reader sees one picture as a fixed number of dots, however big the photo is: a whole page of small
// handwriting comes out blurred to it. So beside the whole page, the page also goes in strips from top to
// bottom (overlapping a little), each at full sharpness. The reader uses the page for the layout and the
// strips for the words. Darker ink, whiter paper: the 1% darkest and lightest tones become black and white.
function levels(x, cw, ch) {
  try {
    const img = x.getImageData(0, 0, cw, ch), d = img.data, hist = new Uint32Array(256), step = Math.max(1, Math.floor(d.length / 4 / 200000));
    for (let i = 0; i < d.length; i += 4 * step) hist[(d[i] * 3 + d[i + 1] * 6 + d[i + 2]) / 10 | 0]++;
    let tot = 0; for (const v of hist) tot += v;
    let lo = 0, hi = 255, acc = 0;
    for (let t = 0; t < 256; t++) { acc += hist[t]; if (acc > tot * 0.01) { lo = t; break; } }
    acc = 0; for (let t = 255; t >= 0; t--) { acc += hist[t]; if (acc > tot * 0.01) { hi = t; break; } }
    if (hi - lo < 60 || (lo < 8 && hi > 247)) return;
    const lut = new Uint8ClampedArray(256); for (let t = 0; t < 256; t++) lut[t] = (t - lo) * 255 / (hi - lo);
    for (let i = 0; i < d.length; i += 4) { d[i] = lut[d[i]]; d[i + 1] = lut[d[i + 1]]; d[i + 2] = lut[d[i + 2]]; }
    x.putImageData(img, 0, 0);
  } catch (_) {}
}
async function stripsOf(src, w, hh) {
  if (SC.quality === 'fast' || Math.max(w, hh) < 1400) return [];
  const n = hh > w * 1.15 ? 3 : 2, ov = 0.07, out = [];
  for (let i = 0; i < n; i++) {
    const y0 = Math.max(0, hh * (i / n - ov)), y1 = Math.min(hh, hh * ((i + 1) / n + ov));
    const st = await canvasJpeg((x, cw, ch) => { x.drawImage(src, 0, y0, w, y1 - y0, 0, 0, cw, ch); levels(x, cw, ch); }, w, y1 - y0, 2000);
    out.push(st.b64);
  }
  return out;
}
/* ---- framing a photographed page: find the sheet, let the person fix the frame, cut it out straight ---- */
async function decodeImg(src) {
  try { return await createImageBitmap(src, { imageOrientation: 'from-image' }); }
  catch (_) { return await new Promise((res, rej) => { const im = new Image(); im.onload = () => res(im); im.onerror = rej; im.src = URL.createObjectURL(src); }); }
}
function scaled(img, w, hh, max) {
  const k = Math.min(1, max / Math.max(w, hh)), c = document.createElement('canvas');
  c.width = Math.max(1, Math.round(w * k)); c.height = Math.max(1, Math.round(hh * k));
  c.getContext('2d').drawImage(img, 0, 0, c.width, c.height); return c;
}
const jpegBlob = (c, q = 0.92) => new Promise(r => c.toBlob(r, 'image/jpeg', q));
const FULL = [[0, 0], [1, 0], [1, 1], [0, 1]];
// a photo, kept once (up to 3200 dots) so the frame can be changed later without losing sharpness
async function prepPhoto(file) {
  const bmp = await decodeImg(file), w = bmp.width || bmp.naturalWidth, hh = bmp.height || bmp.naturalHeight;
  const src = scaled(bmp, w, hh, 3200), view = scaled(bmp, w, hh, 1100);
  if (bmp.close) bmp.close();
  const vu = view.toDataURL('image/jpeg', 0.8);
  const it = { srcBlob: await jpegBlob(src), sw: src.width, sh: src.height, view: vu, viewB64: vu.slice(vu.indexOf(',') + 1), quad: null, auto: null };
  src.width = src.height = 0;
  return it;
}
// four corners in order: top-left, top-right, bottom-right, bottom-left
function orderQuad(q) {
  const cx = q.reduce((a, p) => a + p[0], 0) / 4, cy = q.reduce((a, p) => a + p[1], 0) / 4;
  const s = [...q].sort((a, b) => Math.atan2(a[1] - cy, a[0] - cx) - Math.atan2(b[1] - cy, b[0] - cx)); // clockwise on screen
  let k = 0; for (let i = 1; i < 4; i++) if (s[i][0] + s[i][1] < s[k][0] + s[k][1]) k = i;
  return [0, 1, 2, 3].map(i => s[(k + i) % 4]);
}
const quadArea = q => Math.abs(q.reduce((a, p, i) => { const n = q[(i + 1) % 4]; return a + p[0] * n[1] - n[0] * p[1]; }, 0)) / 2;
// the smart service points at the sheet's corners; an older server, or no answer, leaves the whole photo
async function detectSheet(it) {
  try {
    const raw = await ai({ mode: 'corners', image: it.viewB64 }, { timeout: 20000, once: true });
    const j = JSON.parse(String(raw).replace(/^\s*```(?:json)?|```\s*$/g, ''));
    if (!j || !j.found || !Array.isArray(j.corners) || j.corners.length !== 4) return null;
    const q = orderQuad(j.corners.map(p => [Math.max(0, Math.min(1, (+p[1] || 0) / 1000)), Math.max(0, Math.min(1, (+p[0] || 0) / 1000))]));
    const a = quadArea(q);
    return a > 0.12 && a < 0.985 ? q : null;
  } catch (_) { return null; }
}
// the frame → a straight, flat page (a perspective cut), sharp enough for the reader
async function cutSheet(it) {
  const bmp = await decodeImg(it.srcBlob), W0 = it.sw, H0 = it.sh;
  const q = (it.quad || FULL).map(([x, y]) => [x * W0, y * H0]);
  const dist = (a, b) => Math.hypot(a[0] - b[0], a[1] - b[1]);
  let W = Math.max(dist(q[0], q[1]), dist(q[3], q[2])), H = Math.max(dist(q[0], q[3]), dist(q[1], q[2]));
  const k = Math.min(1, 2800 / Math.max(W, H)); W = Math.max(8, Math.round(W * k)); H = Math.max(8, Math.round(H * k));
  const out = document.createElement('canvas');
  if (!it.quad || it.quad.every((p, i) => Math.abs(p[0] - FULL[i][0]) < 0.003 && Math.abs(p[1] - FULL[i][1]) < 0.003)) {
    out.width = W; out.height = H; out.getContext('2d').drawImage(bmp, 0, 0, W, H);
  } else {
    const sc = document.createElement('canvas'); sc.width = W0; sc.height = H0;
    const sx = sc.getContext('2d', { willReadFrequently: true }); sx.drawImage(bmp, 0, 0, W0, H0);
    const S = sx.getImageData(0, 0, W0, H0).data; sc.width = sc.height = 0;
    // unit square → the four corners (Heckbert's square-to-quad mapping)
    const [[x0, y0], [x1, y1], [x2, y2], [x3, y3]] = q;
    const dx1 = x1 - x2, dx2 = x3 - x2, dx3 = x0 - x1 + x2 - x3, dy1 = y1 - y2, dy2 = y3 - y2, dy3 = y0 - y1 + y2 - y3;
    let a, b, c, d, e, f, g = 0, hh = 0;
    if (Math.abs(dx3) < 1e-9 && Math.abs(dy3) < 1e-9) { a = x1 - x0; b = x2 - x1; c = x0; d = y1 - y0; e = y2 - y1; f = y0; }
    else { const det = dx1 * dy2 - dx2 * dy1; g = (dx3 * dy2 - dx2 * dy3) / det; hh = (dx1 * dy3 - dx3 * dy1) / det;
      a = x1 - x0 + g * x1; b = x3 - x0 + hh * x3; c = x0; d = y1 - y0 + g * y1; e = y3 - y0 + hh * y3; f = y0; }
    out.width = W; out.height = H;
    const ox = out.getContext('2d'), img = ox.createImageData(W, H), D = img.data;
    for (let Y = 0; Y < H; Y++) {
      const v = (Y + 0.5) / H;
      for (let X = 0; X < W; X++) {
        const u = (X + 0.5) / W, z = g * u + hh * v + 1, px = (a * u + b * v + c) / z - 0.5, py = (d * u + e * v + f) / z - 0.5;
        const ix = Math.max(0, Math.min(W0 - 2, px | 0)), iy = Math.max(0, Math.min(H0 - 2, py | 0)), fx = Math.max(0, Math.min(1, px - ix)), fy = Math.max(0, Math.min(1, py - iy));
        const i00 = (iy * W0 + ix) * 4, i10 = i00 + 4, i01 = i00 + W0 * 4, i11 = i01 + 4, o = (Y * W + X) * 4;
        for (let ch = 0; ch < 3; ch++) D[o + ch] = (S[i00 + ch] * (1 - fx) + S[i10 + ch] * fx) * (1 - fy) + (S[i01 + ch] * (1 - fx) + S[i11 + ch] * fx) * fy;
        D[o + 3] = 255;
      }
    }
    ox.putImageData(img, 0, 0);
  }
  if (bmp.close) bmp.close();
  const page = await canvasJpeg((x, cw, ch) => { x.drawImage(out, 0, 0, cw, ch); levels(x, cw, ch); }, W, H);
  page.strips = await stripsOf(out, W, H);
  out.width = out.height = 0;
  Object.assign(it, { url: page.url, b64: page.b64, w: page.w, h: page.h, strips: page.strips });
  return it;
}
async function imgToJpeg(file) { const it = await prepPhoto(file); it.quad = it.auto = await detectSheet(it); return cutSheet(it); }
async function pageToJpeg(page) {
  const v1 = page.getViewport({ scale: 1 }), fast = SC.quality === 'fast';
  // drawn once, big, and the page and its strips are taken from that drawing
  const vp = page.getViewport({ scale: Math.min(5, (fast ? 1700 : 3600) / Math.max(v1.width, v1.height)) });
  const big = document.createElement('canvas'); big.width = Math.round(vp.width); big.height = Math.round(vp.height);
  const bx = big.getContext('2d'); bx.fillStyle = '#fff'; bx.fillRect(0, 0, big.width, big.height);
  await page.render({ canvasContext: bx, viewport: vp }).promise;
  const out = await canvasJpeg(x => { x.drawImage(big, 0, 0, x.canvas.width, x.canvas.height); }, big.width, big.height);
  out.strips = await stripsOf(big, big.width, big.height);
  big.width = big.height = 0;
  return out;
}

// the service's JSON → clean blocks
// circled letters and numbers (Ⓐ ① ❶) are in no font the files use: they become (A) (1)
const unCircle = t => String(t).replace(/[\u24B6-\u24CF]/g, c => `(${String.fromCharCode(c.charCodeAt(0) - 0x24B6 + 65)})`)
  .replace(/[\u24D0-\u24E9]/g, c => `(${String.fromCharCode(c.charCodeAt(0) - 0x24D0 + 97)})`)
  .replace(/[\u2460-\u2473]/g, c => `(${c.charCodeAt(0) - 0x2460 + 1})`)
  .replace(/[\u2776-\u277F]/g, c => `(${c.charCodeAt(0) - 0x2776 + 1})`)
  .replace(/[\u2780-\u2789]/g, c => `(${c.charCodeAt(0) - 0x2780 + 1})`);
// an item that already starts with the label the writer wrote (А. / 1) / (ב) / iii.) keeps it, and gets no added number
const LABEL = /^\s*(?:\(?(?:[A-Za-z\u0400-\u04FF\u05D0-\u05EA]|\d{1,3}|[ivxIVX]{1,4})[.):]|\((?:[A-Za-z\u0400-\u04FF\u05D0-\u05EA]|\d{1,3})\))\s*/;
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
      const rows = (Array.isArray(b.rows) ? b.rows : []).filter(Array.isArray).map(r => r.map(c => unCircle(String(c ?? ''))));
      if (rows.length) out.push({ t, rows });
    } else {
      const text = unCircle(String(b.text ?? '')).replace(/\s+/g, ' ').trim();
      if (text) out.push({ t: t === 'oli' && LABEL.test(text) ? 'p' : t, text });
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
  let lastOl = false, olId = 0;
  pages.forEach((blocks, pi) => {
    if (pi) body.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
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
  const setOf = f => { try { return new Set(f.getCharacterSet()); } catch (_) { return null; } };
  const sets = new Map([[reg, setOf(reg)], [bold, setOf(bold)]]);
  const fits = (f, t) => { const st = sets.get(f); return !st || [...t].every(c => /\s/.test(c) || st.has(c.codePointAt(0))); };
  const fontFor = (f, t) => f === bold && !fits(bold, t) ? reg : f;
  const keepable = (f, t) => { const st = sets.get(f); return !st ? t : [...t].filter(c => /\s/.test(c) || SYM.test(c) || st.has(c.codePointAt(0))).join(''); };
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
    const { base, runs } = bidiRuns(keepable(f, norm(line)), dir);
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
    f = fontFor(f, norm(text));
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
      const f = ri === 0 && r.every(c => fits(bold, norm(String(c ?? '')))) ? bold : reg;
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
  let k = 0;   // a numbered list that goes on to the next page keeps counting
  pages.forEach(blocks => {
    newPage();
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
  // 6.6.1: the worker is kept as a promise at once (two quick calls made two engines before), and the first
  // download shows its progress again
  if (tessWorker && tessWorker._siman !== ocrPref) { const old = tessWorker; tessWorker = null; try { (await old).terminate(); } catch (_) {} }
  if (!tessWorker) {
    const lang = ocrPref;
    tessWorker = Tesseract.createWorker(OCR_SETS[lang], 1, { logger: m => {
      if (tessLog && m.status === 'recognizing text') tessLog(m);
      else if (/load|download|init/i.test(m.status || '') && m.progress != null) { const b = $('busyMsg'); if (b && !$('busy').classList.contains('hidden')) b.textContent = `מוריד את מנוע זיהוי הטקסט (פעם אחת)… ${Math.round(m.progress * 100)}%`; }
    } }).catch(e => { tessWorker = null; throw e; });
    tessWorker._siman = lang;
  }
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
  { l: 'מדף למסמך: צילום או סריקה ל-Word / PDF', k: 'כתב יד כתוב ביד סריקה סרוק צילום מצלמה תמונה הקלדה מודפס הפוך המר וורד word pdf handwriting scan photo camera', run: () => SC.open() },
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
  const scanQ = h(`<button class="qa" type="button">${ICON.scan}מדף למסמך</button>`);
  const askQa = h(`<button class="qa" type="button">${ICON.spark}לשאול את המסמך</button>`);
  scanQ.onclick = () => SC.open();
  askQa.onclick = () => { if (typeof needReg === 'function' && needReg()) { showReg(); return; } afterOpen = showAI; $('openBtn2').click(); };
  const ghost = quick.querySelector('.qa.ghost');
  quick.insertBefore(scanQ, ghost); quick.insertBefore(askQa, ghost);
}
document.querySelectorAll('.wl-list li').forEach(li => { if (/תרגום לעברית ולאנגלית/.test(li.textContent)) li.lastChild.textContent = 'תרגום ל-6 שפות, כולל רוסית וערבית'; });
const wl = document.querySelector('.wl-list');
if (wl) {
  wl.append(h(`<li>${ICON.scan}מדף למסמך: צילום או סריקה לקובץ ערוך</li>`));
  wl.append(h(`<li>${ICON.spark}שאלות, סיכום והקראה של המסמך</li>`));
}
document.addEventListener('keydown', e => { if (e.key === 'Escape' && RA.on && !document.querySelector('dialog[open]')) RA.stop(); });
})();
