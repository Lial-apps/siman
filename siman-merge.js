/* סימן — "לאחד ולפצל" (5.8.2): מסך אחד לאיחוד קבצים ולפיצול עמודים, בטלפון ובמחשב.
   בטלפון נפתח כמסך מלא; במחשב כחלון. כל פעולה משתמשת בפונקציות הקיימות של האפליקציה
   (rebuild, openSplit, __exportPages, buildPdf, saveBytes, zip), כך שהתוצאה זהה לזו של חלונית "עמודים".
   Needs: index.html (S, rebuild, openSplit, __exportPages, buildPdf, saveBytes, zip, busy, toast), pdf.js, pdf-lib. */
(function(){
'use strict';
const $=id=>document.getElementById(id);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[c]);
const T={
  he:{dir:'rtl',title:'לאחד ולפצל',merge:'איחוד',split:'פיצול',close:'סגירה',
    mergeSub:'הקבצים מתאחדים לפי הסדר שכאן. גוררים את החצים כדי לשנות סדר.',cur:'המסמך הפתוח',pages:n=>`${n} עמודים`,page1:'עמוד אחד',
    add:'הוספת קבצים',addSub:'PDF אחד או כמה',up:'הזזה למעלה',down:'הזזה למטה',remove:'הסרה',mergeGo:'איחוד לקובץ אחד',mergeNote:'אחרי האיחוד המסמך נפתח כאן; שומרים אותו בכפתור "שמירה".',
    needMore:'מוסיפים לפחות קובץ אחד כדי לאחד',merged:n=>`אוחדו ${n} קבצים למסמך אחד`,bad:'אחד הקבצים לא נפתח. ייתכן שהוא פגום או מוגן.',
    splitSub:'בוחרים איך לחלק את המסמך לקבצים. כל קובץ נשמר כ-PDF.',
    modes:{all:'כל עמוד לקובץ נפרד',each:'העמודים שסימנתי, כל אחד לקובץ',one:'העמודים שסימנתי, יחד בקובץ אחד',ranges:'לפי טווחי עמודים',chunk:'כל כמה עמודים לקובץ'},
    rangesHint:'למשל: 1-3, 4, 5-8 — כל טווח יישמר כקובץ נפרד',chunkLbl:'עמודים בכל קובץ',selAll:'סימון הכל',selNone:'ניקוי',selected:n=>`${n} מסומנים`,
    splitGo:'המשך',pickPages:'מסמנים עמודים בלחיצה על העמוד',badRange:'לא הבנתי את הטווחים. כותבים למשל: 1-3, 5',ready:n=>`${n} קבצים יישמרו`,readyOne:'קובץ אחד יישמר',
    saved:n=>`נשמרו ${n} קבצים`,savedZip:n=>`${n} הקבצים נשמרו בקובץ ZIP`,fail:'הפיצול נכשל: ',noDoc:'קודם פותחים מסמך PDF',pageN:n=>`עמוד ${n}`},
  ar:{dir:'rtl',title:'دمج وتقسيم',merge:'دمج',split:'تقسيم',close:'إغلاق',
    mergeSub:'تُدمج الملفات بهذا الترتيب. استخدم الأسهم لتغيير الترتيب.',cur:'المستند المفتوح',pages:n=>`${n} صفحات`,page1:'صفحة واحدة',
    add:'إضافة ملفات',addSub:'ملف PDF واحد أو أكثر',up:'نقل لأعلى',down:'نقل لأسفل',remove:'إزالة',mergeGo:'دمج في ملف واحد',mergeNote:'بعد الدمج يُفتح المستند هنا؛ احفظه بزر "حفظ".',
    needMore:'أضف ملفًا واحدًا على الأقل للدمج',merged:n=>`تم دمج ${n} ملفات في مستند واحد`,bad:'تعذّر فتح أحد الملفات. قد يكون تالفًا أو محميًا.',
    splitSub:'اختر كيفية تقسيم المستند إلى ملفات. يُحفظ كل ملف بصيغة PDF.',
    modes:{all:'كل صفحة في ملف منفصل',each:'الصفحات المحددة، كل واحدة في ملف',one:'الصفحات المحددة معًا في ملف واحد',ranges:'حسب نطاقات الصفحات',chunk:'كل عدد من الصفحات في ملف'},
    rangesHint:'مثلًا: 1-3, 4, 5-8 — يُحفظ كل نطاق كملف منفصل',chunkLbl:'صفحات في كل ملف',selAll:'تحديد الكل',selNone:'مسح',selected:n=>`${n} محددة`,
    splitGo:'متابعة',pickPages:'اضغط على صفحة لتحديدها',badRange:'لم أفهم النطاقات. اكتب مثلًا: 1-3, 5',ready:n=>`سيُحفظ ${n} ملفات`,readyOne:'سيُحفظ ملف واحد',
    saved:n=>`تم حفظ ${n} ملفات`,savedZip:n=>`تم حفظ ${n} ملفات في ملف ZIP`,fail:'فشل التقسيم: ',noDoc:'افتح مستند PDF أولًا',pageN:n=>`صفحة ${n}`},
  ru:{dir:'ltr',title:'Объединить и разделить',merge:'Объединить',split:'Разделить',close:'Закрыть',
    mergeSub:'Файлы объединяются в этом порядке. Стрелками можно менять порядок.',cur:'Открытый документ',pages:n=>`${n} стр.`,page1:'1 стр.',
    add:'Добавить файлы',addSub:'Один или несколько PDF',up:'Вверх',down:'Вниз',remove:'Убрать',mergeGo:'Объединить в один файл',mergeNote:'После объединения документ откроется здесь; сохраните его кнопкой «Сохранить».',
    needMore:'Добавьте хотя бы один файл',merged:n=>`Объединено файлов: ${n}`,bad:'Один из файлов не открылся. Возможно, он повреждён или защищён.',
    splitSub:'Выберите, как разделить документ на файлы. Каждый файл сохраняется как PDF.',
    modes:{all:'Каждая страница — отдельный файл',each:'Выбранные страницы, каждая в файл',one:'Выбранные страницы вместе в один файл',ranges:'По диапазонам страниц',chunk:'По N страниц в файл'},
    rangesHint:'Например: 1-3, 4, 5-8 — каждый диапазон станет отдельным файлом',chunkLbl:'страниц в файле',selAll:'Выбрать все',selNone:'Снять',selected:n=>`Выбрано: ${n}`,
    splitGo:'Продолжить',pickPages:'Нажмите на страницу, чтобы выбрать её',badRange:'Не понял диапазоны. Напишите, например: 1-3, 5',ready:n=>`Будет сохранено файлов: ${n}`,readyOne:'Будет сохранён один файл',
    saved:n=>`Сохранено файлов: ${n}`,savedZip:n=>`${n} файлов сохранены в ZIP`,fail:'Не удалось разделить: ',noDoc:'Сначала откройте PDF',pageN:n=>`Стр. ${n}`},
  en:{dir:'ltr',title:'Merge and split',merge:'Merge',split:'Split',close:'Close',
    mergeSub:'Files are merged in this order. Use the arrows to reorder.',cur:'Open document',pages:n=>`${n} pages`,page1:'1 page',
    add:'Add files',addSub:'One or more PDFs',up:'Move up',down:'Move down',remove:'Remove',mergeGo:'Merge into one file',mergeNote:'After merging, the document opens here; save it with the Save button.',
    needMore:'Add at least one file to merge',merged:n=>`${n} files merged into one document`,bad:'One of the files could not be opened. It may be damaged or protected.',
    splitSub:'Choose how to divide the document into files. Each file is saved as a PDF.',
    modes:{all:'Every page as a separate file',each:'Selected pages, each as a file',one:'Selected pages together in one file',ranges:'By page ranges',chunk:'Every N pages as a file'},
    rangesHint:'For example: 1-3, 4, 5-8 — each range becomes a separate file',chunkLbl:'pages per file',selAll:'Select all',selNone:'Clear',selected:n=>`${n} selected`,
    splitGo:'Continue',pickPages:'Tap a page to select it',badRange:'Could not read the ranges. Write, for example: 1-3, 5',ready:n=>`${n} files will be saved`,readyOne:'One file will be saved',
    saved:n=>`${n} files saved`,savedZip:n=>`${n} files saved in a ZIP`,fail:'Split failed: ',noDoc:'Open a PDF document first',pageN:n=>`Page ${n}`}
};
const L=()=>T[document.documentElement.getAttribute('lang')]||T.he;
const I={
  x:'<svg viewBox="0 0 24 24" width="22" height="22" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M6 6l12 12M18 6L6 18"/></svg>',
  up:'<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 19V5M6 11l6-6 6 6"/></svg>',
  down:'<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 5v14M6 13l6 6 6-6"/></svg>',
  plus:'<svg viewBox="0 0 24 24" width="24" height="24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>',
  merge:'<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M8 3h6l4 4v10H8z"/><path d="M5 7v14h10"/></svg>',
  split:'<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 4h7v9H4zM13 11h7v9h-7z"/><path d="M15 4h5v4M9 20H4v-4"/></svg>',
  check:'<svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>'
};
const CSS=`
dialog.ms{border:0;padding:0;width:min(820px,94vw);max-height:min(92vh,920px);border-radius:22px;background:var(--bar,#fff);color:var(--ink,#1C2333);box-shadow:0 30px 60px -20px rgba(13,20,36,.45);overflow:hidden}
dialog.ms::backdrop{background:rgba(13,20,36,.42);backdrop-filter:blur(3px)}
.ms-in{display:flex;flex-direction:column;height:min(92vh,920px);min-height:0}
.ms-head{display:flex;align-items:center;gap:12px;padding:16px 18px 10px;border-bottom:1px solid var(--line,#D4DBE5)}
.ms-head h2{margin:0;font-size:21px;font-weight:800;flex:1}
.ms-x{width:42px;height:42px;border-radius:50%;border:1px solid var(--line,#D4DBE5);background:transparent;color:var(--ink,#1C2333);display:grid;place-items:center;cursor:pointer}
.ms-x:hover{background:var(--desk,#EEF3FB)}
.ms-tabs{display:flex;gap:6px;padding:12px 18px 0}
.ms-tab{flex:1;display:flex;align-items:center;justify-content:center;gap:8px;min-height:48px;border-radius:14px;border:1.5px solid var(--line,#D4DBE5);background:transparent;color:var(--ink,#1C2333);font:inherit;font-size:16px;font-weight:700;cursor:pointer}
.ms-tab[aria-selected="true"]{background:var(--blue,#2B59C3);border-color:var(--blue,#2B59C3);color:var(--marker-ink,#fff)}
.ms-body{flex:1;min-height:0;overflow:auto;padding:14px 18px;display:flex;flex-direction:column;gap:14px;-webkit-overflow-scrolling:touch}
.ms-sub{margin:0;color:var(--muted,#5D6678);font-size:14px;line-height:1.5}
.ms-list{display:flex;flex-direction:column;gap:10px}
.ms-file{display:flex;align-items:center;gap:12px;padding:10px 12px;border:1px solid var(--line,#D4DBE5);border-radius:16px;background:var(--desk,#F3F6FA)}
.ms-file.cur{border-color:var(--blue,#2B59C3);background:color-mix(in srgb,var(--blue,#2B59C3) 8%,var(--bar,#fff))}
.ms-th{width:46px;height:60px;flex:none;border-radius:6px;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.25);overflow:hidden;display:grid;place-items:center}
.ms-th canvas{max-width:100%;max-height:100%;display:block}
.ms-fi{flex:1;min-width:0}
.ms-fi b{display:block;font-size:15px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;direction:ltr;text-align:start;unicode-bidi:plaintext}
.ms-fi small{color:var(--muted,#5D6678);font-size:13px}
.ms-fi .tag{display:inline-block;margin-inline-start:6px;padding:1px 8px;border-radius:999px;background:var(--blue,#2B59C3);color:var(--marker-ink,#fff);font-size:11px;font-weight:700}
.ms-ord{display:flex;flex-direction:column;gap:4px}
.ms-ord button,.ms-rm{width:36px;height:30px;border-radius:9px;border:1px solid var(--line,#D4DBE5);background:var(--bar,#fff);color:var(--ink,#1C2333);display:grid;place-items:center;cursor:pointer;font:inherit}
.ms-ord button:disabled{opacity:.35;cursor:default}
.ms-rm{height:36px;color:var(--danger,#C2362B);font-size:18px}
.ms-num{width:28px;height:28px;border-radius:50%;background:var(--ink,#1C2333);color:var(--bar,#fff);display:grid;place-items:center;font-size:13px;font-weight:700;flex:none}
.ms-add{display:flex;align-items:center;gap:12px;padding:14px;border:2px dashed var(--blue,#2B59C3);border-radius:16px;background:transparent;color:var(--blue,#2B59C3);cursor:pointer;font:inherit;text-align:start;width:100%}
.ms-add .pl{width:40px;height:40px;border-radius:12px;display:grid;place-items:center;background:color-mix(in srgb,var(--blue,#2B59C3) 12%,transparent)}
.ms-add b{display:block;font-size:16px}.ms-add small{color:var(--muted,#5D6678)}
.ms-foot{display:flex;align-items:center;gap:10px;padding:12px 18px calc(14px + env(safe-area-inset-bottom,0px));border-top:1px solid var(--line,#D4DBE5);background:var(--bar,#fff)}
.ms-foot .st{flex:1;font-size:14px;color:var(--muted,#5D6678)}
.ms-pri{min-height:48px;padding:0 22px;border-radius:14px;border:0;background:var(--blue,#2B59C3);color:var(--marker-ink,#fff);font:inherit;font-size:16px;font-weight:700;cursor:pointer}
.ms-pri:disabled{opacity:.45;cursor:default}
.ms-modes{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:8px}
.ms-mode{display:flex;align-items:center;gap:10px;min-height:50px;padding:8px 12px;border-radius:14px;border:1.5px solid var(--line,#D4DBE5);background:transparent;color:var(--ink,#1C2333);font:inherit;font-size:14.5px;font-weight:600;cursor:pointer;text-align:start}
.ms-mode i{width:20px;height:20px;border-radius:50%;border:2px solid var(--line,#D4DBE5);flex:none;display:grid;place-items:center;color:transparent}
.ms-mode[aria-checked="true"]{border-color:var(--blue,#2B59C3);background:color-mix(in srgb,var(--blue,#2B59C3) 9%,transparent)}
.ms-mode[aria-checked="true"] i{background:var(--blue,#2B59C3);border-color:var(--blue,#2B59C3);color:#fff}
.ms-field{display:flex;flex-direction:column;gap:6px}
.ms-field[hidden],.ms-selrow[hidden],.ms-sub[hidden],.ms-grid[hidden]{display:none}
.ms-field label{font-size:14px;font-weight:600}
.ms-field input{min-height:46px;padding:0 12px;border-radius:12px;border:1.5px solid var(--line,#D4DBE5);background:var(--bar,#fff);color:var(--ink,#1C2333);font:inherit;font-size:16px}
.ms-field small{color:var(--muted,#5D6678)}
.ms-selrow{display:flex;align-items:center;gap:14px;font-size:14px}
.ms-selrow button{border:0;background:none;color:var(--blue,#2B59C3);font:inherit;font-weight:700;cursor:pointer;padding:6px 0}
.ms-selrow span{color:var(--muted,#5D6678);margin-inline-start:auto}
.ms-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(104px,1fr));gap:10px}
.ms-pg{position:relative;border:2px solid var(--line,#D4DBE5);border-radius:12px;padding:6px 6px 4px;background:var(--desk,#F3F6FA);cursor:pointer;display:flex;flex-direction:column;align-items:center;gap:4px;font:inherit;color:var(--ink,#1C2333)}
.ms-pg .c{width:100%;aspect-ratio:3/4;background:#fff;border-radius:6px;box-shadow:0 1px 3px rgba(0,0,0,.25);display:grid;place-items:center;overflow:hidden}
.ms-pg canvas{max-width:100%;max-height:100%;display:block}
.ms-pg small{font-size:12px;font-weight:600}
.ms-pg .ck{position:absolute;top:10px;inset-inline-end:10px;width:24px;height:24px;border-radius:7px;border:2px solid var(--line,#D4DBE5);background:var(--bar,#fff);display:grid;place-items:center;color:transparent}
.ms-pg[aria-checked="true"]{border-color:var(--blue,#2B59C3)}
.ms-pg[aria-checked="true"] .ck{background:var(--blue,#2B59C3);border-color:var(--blue,#2B59C3);color:#fff}
.ms-pg.dim{opacity:.55}
@media (max-width:720px){
  dialog.ms{width:100vw;max-width:100vw;height:100dvh;max-height:100dvh;border-radius:0;margin:0}
  .ms-in{height:100dvh}
  .ms-head{padding:calc(12px + env(safe-area-inset-top,0px)) 14px 8px}
  .ms-head h2{font-size:19px}
  .ms-tabs{padding:10px 14px 0}
  .ms-body{padding:12px 14px}
  .ms-grid{grid-template-columns:repeat(3,minmax(0,1fr))}
  .ms-modes{grid-template-columns:1fr}
  .ms-foot{flex-wrap:wrap}.ms-foot .ms-pri{flex:1 1 100%}
}
@media (hover:hover){.ms-pg:hover{border-color:var(--blue,#2B59C3)}}
`;
let st=null; // {tab, extra:[{name,bytes,count,doc}], order:[...indexes into files], sel:Set, mode, ranges, chunk}
let dlg=null;
function T2(){return L();}
function build(){
  if(dlg)return;
  const s=document.createElement('style');s.textContent=CSS;document.head.append(s);
  dlg=document.createElement('dialog');dlg.className='ms';dlg.id='msDlg';
  dlg.innerHTML=`<div class="ms-in"><div class="ms-head"><h2 id="msTitle"></h2><button class="ms-x" id="msX" type="button">${I.x}</button></div>
    <div class="ms-tabs" role="tablist"><button class="ms-tab" role="tab" data-tab="merge" type="button">${I.merge}<span></span></button><button class="ms-tab" role="tab" data-tab="split" type="button">${I.split}<span></span></button></div>
    <div class="ms-body" id="msBody"></div>
    <div class="ms-foot"><span class="st" id="msSt"></span><button class="ms-pri" id="msGo" type="button"></button></div></div>
    <input type="file" id="msFile" accept="application/pdf,.pdf" multiple hidden>`;
  document.body.append(dlg);
  dlg.querySelector('#msX').onclick=close;
  dlg.addEventListener('click',e=>{if(e.target===dlg)close();});
  dlg.addEventListener('cancel',e=>{e.preventDefault();close();});
  dlg.querySelectorAll('.ms-tab').forEach(b=>b.onclick=()=>{st.tab=b.dataset.tab;render();});
  dlg.querySelector('#msFile').onchange=async e=>{const files=[...e.target.files];e.target.value='';if(!files.length)return;await addFiles(files);};
  dlg.querySelector('#msGo').onclick=()=>st.tab==='merge'?doMerge():doSplit();
}
function close(){if(!dlg||!dlg.open)return;dlg.close();
  if(st){for(const x of st.extra)try{x.doc&&x.doc.destroy();}catch(_){}
    // the page selection carries over to the "pages" panel
    if(typeof S!=='undefined'&&S.sel){S.sel=new Set(st.sel);try{updateSel();}catch(_){}}}
}
async function open(tab){
  if(typeof S==='undefined'||!S.doc||!S.pages||!S.pages.length){try{toast(T2().noDoc);}catch(_){}return;}
  build();
  st={tab:tab||'merge',extra:[],order:[0],sel:new Set(S.sel||[]),mode:'all',ranges:'',chunk:2};
  dlg.dir=T2().dir;render();dlg.showModal();
}
function render(){
  const L=T2();dlg.dir=L.dir;dlg.querySelector('#msTitle').textContent=L.title;
  dlg.querySelectorAll('.ms-tab').forEach(b=>{b.setAttribute('aria-selected',b.dataset.tab===st.tab);b.querySelector('span').textContent=b.dataset.tab==='merge'?L.merge:L.split;});
  const body=dlg.querySelector('#msBody');body.scrollTop=0;
  if(st.tab==='merge')renderMerge(body);else renderSplit(body);
}
/* ------------------------------------------------------------ merge */
const files=()=>[{cur:true,name:S.name,count:S.pages.length},...st.extra];
function renderMerge(body){
  const L=T2(),fs=files();
  body.innerHTML=`<p class="ms-sub">${esc(L.mergeSub)}</p><div class="ms-list" id="msList"></div>
    <button class="ms-add" type="button" id="msAdd"><span class="pl">${I.plus}</span><span><b>${esc(L.add)}</b><small>${esc(L.addSub)}</small></span></button>
    <p class="ms-sub">${esc(L.mergeNote)}</p>`;
  const list=body.querySelector('#msList');
  st.order.forEach((fi,k)=>{const f=fs[fi];const el=document.createElement('div');el.className='ms-file'+(f.cur?' cur':'');
    el.innerHTML=`<span class="ms-num">${k+1}</span><div class="ms-th" data-th="${fi}"></div><div class="ms-fi"><b title="${esc(f.name)}">${esc(f.name)}</b><small>${f.count===1?esc(L.page1):esc(L.pages(f.count))}${f.cur?`<span class="tag">${esc(L.cur)}</span>`:''}</small></div>
      <div class="ms-ord"><button type="button" data-mv="-1" title="${esc(L.up)}" aria-label="${esc(L.up)}" ${k?'':'disabled'}>${I.up}</button><button type="button" data-mv="1" title="${esc(L.down)}" aria-label="${esc(L.down)}" ${k<st.order.length-1?'':'disabled'}>${I.down}</button></div>
      ${f.cur?'':`<button type="button" class="ms-rm" title="${esc(L.remove)}" aria-label="${esc(L.remove)}">×</button>`}`;
    el.querySelectorAll('[data-mv]').forEach(b=>b.onclick=()=>{const j=k+ +b.dataset.mv;if(j<0||j>=st.order.length)return;[st.order[k],st.order[j]]=[st.order[j],st.order[k]];render();});
    const rm=el.querySelector('.ms-rm');if(rm)rm.onclick=()=>{const x=st.extra[fi-1];try{x.doc&&x.doc.destroy();}catch(_){}st.extra.splice(fi-1,1);st.order=st.order.filter(o=>o!==fi).map(o=>o>fi?o-1:o);render();};
    list.append(el);
    thumb(el.querySelector('.ms-th'),f.cur?S.pages[0].page:null,f.cur?null:f);});
  body.querySelector('#msAdd').onclick=()=>dlg.querySelector('#msFile').click();
  const go=dlg.querySelector('#msGo');go.textContent=L.mergeGo;go.disabled=st.extra.length<1;
  const n=files().reduce((t,f)=>t+f.count,0);dlg.querySelector('#msSt').textContent=st.extra.length?L.pages(n):L.needMore;
}
async function thumb(slot,page,f){
  try{
    let pg=page;
    if(!pg&&f){if(!f.doc){const lib=window.pdfjsLib||window['pdfjs-dist/build/pdf'];f.doc=await lib.getDocument({data:f.bytes.slice(0),isEvalSupported:false}).promise;}pg=await f.doc.getPage(1);}
    if(!pg||!slot.isConnected)return;
    const v1=pg.getViewport({scale:1});const dpr=Math.min(2,devicePixelRatio||1);const sc=Math.min(60/v1.width,60/v1.height);
    const v=pg.getViewport({scale:sc*dpr});const c=document.createElement('canvas');c.width=Math.ceil(v.width);c.height=Math.ceil(v.height);
    c.style.width=Math.ceil(v.width/dpr)+'px';c.style.height=Math.ceil(v.height/dpr)+'px';
    await pg.render({canvasContext:c.getContext('2d'),viewport:v}).promise;if(slot.isConnected){slot.innerHTML='';slot.append(c);}
  }catch(_){}
}
async function addFiles(list){
  const L=T2();busy(true,'…');
  try{
    for(const f of list){const bytes=new Uint8Array(await f.arrayBuffer());let count=0;
      try{const d=await PDFLib.PDFDocument.load(bytes,{ignoreEncryption:true});count=d.getPageCount();}catch(_){toast(L.bad);continue;}
      st.extra.push({name:f.name,bytes,count});st.order.push(st.extra.length);}
  }finally{busy(false);}
  render();
}
async function doMerge(){
  const L=T2();if(!st.extra.length)return;
  if(S.hasForm&&!confirm('הוספת קובץ תבטל את שדות הטופס בקובץ. להמשיך?'))return;
  const fs=files();const order=[];
  for(const fi of st.order){const f=fs[fi];for(let k=0;k<f.count;k++)order.push({src:fi,i:k,rot:0});}
  const extra=st.extra.map(x=>x.bytes);const n=fs.length;
  close();
  await rebuild(order,extra);
  toast(L.merged(n));
}
/* ------------------------------------------------------------ split */
function renderSplit(body){
  const L=T2(),n=S.pages.length;
  const modes=['all','each','one','ranges','chunk'];
  body.innerHTML=`<p class="ms-sub">${esc(L.splitSub)}</p>
    <div class="ms-modes" role="radiogroup">${modes.map(m=>`<button type="button" class="ms-mode" role="radio" data-m="${m}" aria-checked="${st.mode===m}"><i>${I.check}</i><span>${esc(L.modes[m])}</span></button>`).join('')}</div>
    <div class="ms-field" id="msRanges" hidden><label for="msRangesIn">${esc(L.modes.ranges)}</label><input id="msRangesIn" inputmode="numeric" dir="ltr" placeholder="1-3, 4, 5-8" value="${esc(st.ranges)}"><small>${esc(L.rangesHint)}</small></div>
    <div class="ms-field" id="msChunk" hidden><label for="msChunkIn">${esc(L.chunkLbl)}</label><input id="msChunkIn" type="number" min="1" max="${n}" dir="ltr" value="${st.chunk}"></div>
    <div class="ms-selrow" id="msSelRow"><button type="button" id="msAll">${esc(L.selAll)}</button><button type="button" id="msNone">${esc(L.selNone)}</button><span id="msCount"></span></div>
    <p class="ms-sub" id="msPick">${esc(L.pickPages)}</p>
    <div class="ms-grid" id="msGrid" role="group"></div>`;
  body.querySelectorAll('.ms-mode').forEach(b=>b.onclick=()=>{st.mode=b.dataset.m;body.querySelectorAll('.ms-mode').forEach(x=>x.setAttribute('aria-checked',x===b));syncSplit();});
  body.querySelector('#msRangesIn').oninput=e=>{st.ranges=e.target.value;syncSplit();};
  body.querySelector('#msChunkIn').oninput=e=>{st.chunk=Math.max(1,+e.target.value||1);syncSplit();};
  body.querySelector('#msAll').onclick=()=>{st.sel=new Set(S.pages.map((_,i)=>i));paintSel();syncSplit();};
  body.querySelector('#msNone').onclick=()=>{st.sel=new Set();paintSel();syncSplit();};
  const grid=body.querySelector('#msGrid');
  S.pages.forEach((P,i)=>{const b=document.createElement('button');b.type='button';b.className='ms-pg';b.dataset.i=i;b.setAttribute('role','checkbox');
    b.innerHTML=`<span class="c"></span><small>${esc(L.pageN(i+1))}</small><span class="ck">${I.check}</span>`;
    b.onclick=()=>{if(st.sel.has(i))st.sel.delete(i);else st.sel.add(i);paintSel();syncSplit();};grid.append(b);});
  paintSel();syncSplit();
  (async()=>{for(let i=0;i<n;i++){const slot=grid.querySelector(`.ms-pg[data-i="${i}"] .c`);if(!slot||!slot.isConnected)return;
    const P=S.pages[i];try{const v1=P.page.getViewport({scale:1});const dpr=Math.min(2,devicePixelRatio||1);const sc=Math.min(120/v1.width,160/v1.height);const v=P.page.getViewport({scale:sc*dpr});
      const c=document.createElement('canvas');c.width=Math.ceil(v.width);c.height=Math.ceil(v.height);c.style.width=Math.ceil(v.width/dpr)+'px';c.style.height=Math.ceil(v.height/dpr)+'px';
      await P.page.render({canvasContext:c.getContext('2d'),viewport:v}).promise;if(slot.isConnected){slot.innerHTML='';slot.append(c);}}catch(_){}}})();
}
function paintSel(){const L=T2();dlg.querySelectorAll('.ms-pg').forEach(b=>b.setAttribute('aria-checked',st.sel.has(+b.dataset.i)));const c=dlg.querySelector('#msCount');if(c)c.textContent=st.sel.size?L.selected(st.sel.size):'';}
function parseRanges(s,n){const out=[];for(const part of String(s).split(/[,;،]+/)){const p=part.trim();if(!p)continue;const m=/^(\d+)\s*[-–—]\s*(\d+)$/.exec(p)||/^(\d+)$/.exec(p);if(!m)return null;
  let a=+m[1],b=m[2]!==undefined?+m[2]:a;if(a>b)[a,b]=[b,a];if(a<1||b>n)return null;const g=[];for(let i=a;i<=b;i++)g.push(i-1);out.push(g);}return out.length?out:null;}
function groups(){const n=S.pages.length,sel=[...st.sel].sort((a,b)=>a-b);
  switch(st.mode){
    case 'all':return S.pages.map((_,i)=>[i]);
    case 'each':return sel.map(i=>[i]);
    case 'one':return sel.length?[sel]:[];
    case 'ranges':return parseRanges(st.ranges,n)||[];
    case 'chunk':{const k=Math.max(1,st.chunk|0),out=[];for(let i=0;i<n;i+=k)out.push(S.pages.slice(i,i+k).map((_,j)=>i+j));return out;}
  }return [];}
function syncSplit(){
  const L=T2();const usesSel=st.mode==='each'||st.mode==='one';
  dlg.querySelector('#msRanges').hidden=st.mode!=='ranges';dlg.querySelector('#msChunk').hidden=st.mode!=='chunk';
  dlg.querySelector('#msSelRow').hidden=!usesSel;dlg.querySelector('#msPick').hidden=!usesSel;
  const gs=groups();const inG=new Set(gs.flat());
  dlg.querySelectorAll('.ms-pg').forEach(b=>b.classList.toggle('dim',usesSel?false:!inG.has(+b.dataset.i)));
  const bad=st.mode==='ranges'&&st.ranges.trim()&&!parseRanges(st.ranges,S.pages.length);
  dlg.querySelector('#msSt').textContent=bad?L.badRange:gs.length===1?L.readyOne:gs.length?L.ready(gs.length):(usesSel?L.pickPages:'');
  const go=dlg.querySelector('#msGo');go.textContent=L.splitGo;go.disabled=!gs.length||bad;
}
async function doSplit(){
  const L=T2();const gs=groups();if(!gs.length)return;
  if(st.mode==='all'||st.mode==='each'){const ix=gs.map(g=>g[0]);close();openSplit(ix);return;}            // names per page, as before
  if(gs.length===1){const g=gs[0];close();await window.__exportPages(g,'one');return;}
  close();await exportGroups(gs);
}
// several files, each with a group of pages (ranges / every N pages)
async function exportGroups(gs){
  const L=T2();busy(true,'…');
  try{
    const built=await buildPdf();const src=await PDFLib.PDFDocument.load(built);const base=S.name.replace(/\.pdf$/i,'');const out=[];
    for(const g of gs){const d=await PDFLib.PDFDocument.create();(await d.copyPages(src,g)).forEach(p=>d.addPage(p));out.push({name:`${base} - ${rangeLabel(g).replace(/,/g,'_')}.pdf`,data:await d.save()});}
    busy(false);
    if(window.showDirectoryPicker){let dir;try{dir=await showDirectoryPicker({mode:'readwrite'});}catch(e){if(e&&e.name==='AbortError')return;dir=null;}
      if(dir){busy(true,'…');for(const f of out){const h=await dir.getFileHandle(f.name,{create:true});await writeHandle(h,f.data);}busy(false);toast(L.saved(out.length));return;}}
    if(await saveBytes(`${base} - ${out.length}.zip`,zip(out),'zip'))toast(L.savedZip(out.length));
  }catch(e){console.error(e);toast(L.fail+(e.message||e));}
  finally{busy(false);}
}
/* ------------------------------------------------------------ hooks into the app */
window.openMergeSplit=open;
function hook(){
  // the first-screen tile and the command "לאחד ולפצל" open this screen once the file is loaded
  if(typeof QA!=='undefined')QA.pages=()=>open('merge');
  // a prominent button at the top of the "pages" panel
  const tools=document.querySelector('#pgpanel .pg-tools');
  if(tools&&!$('msBtn')){const b=document.createElement('button');b.className='act ms-act';b.id='msBtn';b.innerHTML=`${I.merge}<span>${esc(T2().title)}</span>`;b.onclick=()=>open('merge');tools.prepend(b);
    const s=document.createElement('style');s.textContent=`#msBtn{background:var(--blue,#2B59C3);color:var(--marker-ink,#fff);border-color:transparent;font-weight:700;margin-bottom:6px}#msBtn svg{stroke:currentColor}@media (max-width:720px){.pgpanel .pg-tools{padding-bottom:12px}.pgpanel .thumbs{padding-bottom:calc(96px + env(safe-area-inset-bottom,0px))}}`;document.head.append(s);}
  // the merge tile when a document is already open: go straight to the screen
  document.addEventListener('click',e=>{const t=e.target.closest&&e.target.closest('[data-v7="merge"]');if(!t||typeof S==='undefined'||!S.doc)return;e.stopImmediatePropagation();e.preventDefault();open('merge');},true);
  // keep the button's caption in the interface language
  new MutationObserver(()=>{const b=$('msBtn');if(b)b.querySelector('span').textContent=T2().title;}).observe(document.documentElement,{attributes:true,attributeFilter:['lang']});
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',hook);else hook();
})();
