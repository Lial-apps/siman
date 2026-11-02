/* סימן — שדרוגי חוויית משתמש (5.9.0).
   1. שמירה אוטומטית של טיוטה ל-PDF (הקובץ + הסימונים, במכשיר בלבד) ו"להמשיך מאיפה שהפסקת" בפתיחה.
   2. "קבצים אחרונים" במסך הפתיחה (נשמרים במכשיר, לא עולים לשום מקום).
   3. טלפון: הכפתורים הצפים לא מכסים תוכן, ההודעות מעל שורת הקבוצות, מספר הגרסה מוסתר, תפריט "עוד" בסרגל.
   4. Word בטלפון: "תצוגת טלפון" — הטקסט נשבר לרוחב המסך (רק לקריאה ולכתיבה; הקובץ עצמו לא משתנה).
   5. מדף למסמך במחשב: אזור גרירה גדול, סוגי סריקה ברורים, בצבעי המראה.
   6. PowerPoint בטלפון: ההסבר לא נשבר מילה-מילה.
   Needs: index.html (S, openFile, setDirty, histStr, histParse, drawAnns, toast, redoAct). All storage is IndexedDB, on this device. */
(function(){
'use strict';
const $=id=>document.getElementById(id);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[c]);
const narrow=()=>matchMedia('(max-width:720px)').matches;
const lang=()=>document.documentElement.getAttribute('lang')||'he';
const T={
  he:{resume:'להמשיך מאיפה שהפסקת?',resumeSub:(n,t)=>`${n} · נשמר ${t}`,go:'המשך',drop:'לא, להתחיל מחדש',recent:'קבצים אחרונים',clear:'ניקוי',remove:'הסרה מהרשימה',
    ago:{now:'לפני רגע',min:n=>`לפני ${n} דק׳`,hr:n=>`לפני ${n} שע׳`,day:n=>`לפני ${n} ימים`},more:'עוד',reflowOn:'תצוגת טלפון',reflowOff:'תצוגת דף',reflowTip:'תצוגת טלפון: הטקסט נשבר לרוחב המסך. הקובץ עצמו לא משתנה',
    restored:'הטיוטה שוחזרה. השינויים האחרונים שלך כאן',redo:'בצע שוב (Ctrl+Y)',savedLocal:'הטיוטה נשמרת אוטומטית במכשיר'},
  ar:{resume:'المتابعة من حيث توقفت؟',resumeSub:(n,t)=>`${n} · حُفظ ${t}`,go:'متابعة',drop:'لا، البدء من جديد',recent:'الملفات الأخيرة',clear:'مسح',remove:'إزالة من القائمة',
    ago:{now:'قبل لحظة',min:n=>`قبل ${n} د`,hr:n=>`قبل ${n} س`,day:n=>`قبل ${n} أيام`},more:'المزيد',reflowOn:'عرض الهاتف',reflowOff:'عرض الصفحة',reflowTip:'عرض الهاتف: يلتف النص على عرض الشاشة. الملف نفسه لا يتغير',
    restored:'تمت استعادة المسودة',redo:'إعادة (Ctrl+Y)',savedLocal:'تُحفظ المسودة تلقائيًا على الجهاز'},
  ru:{resume:'Продолжить с того места?',resumeSub:(n,t)=>`${n} · сохранено ${t}`,go:'Продолжить',drop:'Нет, начать заново',recent:'Недавние файлы',clear:'Очистить',remove:'Убрать из списка',
    ago:{now:'только что',min:n=>`${n} мин назад`,hr:n=>`${n} ч назад`,day:n=>`${n} дн назад`},more:'Ещё',reflowOn:'Вид для телефона',reflowOff:'Вид страницы',reflowTip:'Вид для телефона: текст переносится по ширине экрана. Сам файл не меняется',
    restored:'Черновик восстановлен',redo:'Повторить (Ctrl+Y)',savedLocal:'Черновик сохраняется на устройстве автоматически'},
  en:{resume:'Pick up where you left off?',resumeSub:(n,t)=>`${n} · saved ${t}`,go:'Continue',drop:'No, start fresh',recent:'Recent files',clear:'Clear',remove:'Remove from list',
    ago:{now:'just now',min:n=>`${n} min ago`,hr:n=>`${n} h ago`,day:n=>`${n} d ago`},more:'More',reflowOn:'Phone view',reflowOff:'Page view',reflowTip:'Phone view: text wraps to the screen width. The file itself does not change',
    restored:'Draft restored',redo:'Redo (Ctrl+Y)',savedLocal:'The draft is saved on this device automatically'}
};
const L=()=>T[lang()]||T.he;
const ago=t=>{const L0=L().ago,d=Date.now()-t;if(d<60e3)return L0.now;if(d<3600e3)return L0.min(Math.round(d/60e3));if(d<86400e3)return L0.hr(Math.round(d/3600e3));return L0.day(Math.round(d/86400e3));};
const ICON={
  pdf:'<span class="ux-ic pdf">PDF</span>',docx:'<span class="ux-ic docx">W</span>',xlsx:'<span class="ux-ic xlsx">X</span>',pptx:'<span class="ux-ic pptx">P</span>',doc:'<span class="ux-ic docx">W</span>'
};
const kindOf=n=>(String(n).match(/\.(pdf|docx|xlsx|pptx|doc)$/i)||[,'pdf'])[1].toLowerCase();

/* ------------------------------------------------------------ storage (IndexedDB, this device only) */
const DB={
  open(){if(this.p)return this.p;return this.p=new Promise((res,rej)=>{const r=indexedDB.open('siman-ux',1);
    r.onupgradeneeded=()=>{const d=r.result;if(!d.objectStoreNames.contains('kv'))d.createObjectStore('kv');};
    r.onsuccess=()=>res(r.result);r.onerror=()=>rej(r.error);});},
  async get(k){try{const d=await this.open();return await new Promise((res,rej)=>{const q=d.transaction('kv').objectStore('kv').get(k);q.onsuccess=()=>res(q.result);q.onerror=()=>rej(q.error);});}catch(_){return undefined;}},
  async set(k,v){try{const d=await this.open();await new Promise((res,rej)=>{const t=d.transaction('kv','readwrite');t.objectStore('kv').put(v,k);t.oncomplete=res;t.onerror=()=>rej(t.error);});return true;}catch(_){return false;}},
  async del(k){try{const d=await this.open();await new Promise((res,rej)=>{const t=d.transaction('kv','readwrite');t.objectStore('kv').delete(k);t.oncomplete=res;t.onerror=()=>rej(t.error);});}catch(_){}}
};
const MAX_FILE=20*1024*1024,MAX_TOTAL=80*1024*1024,MAX_N=8;

/* ------------------------------------------------------------ CSS */
const css=document.createElement('style');css.textContent=`
.ux-ic{display:inline-grid;place-items:center;width:38px;height:38px;border-radius:11px;font-weight:800;font-size:12px;color:#fff;flex:none;letter-spacing:.02em}
.ux-ic.pdf{background:#C2362B}.ux-ic.docx{background:#2B579A}.ux-ic.xlsx{background:#1E7145}.ux-ic.pptx{background:#C43E1C}
/* resume bar */
.ux-resume{display:flex;align-items:center;gap:12px;margin:10px auto 0;max-width:980px;padding:10px 14px;border-radius:16px;background:var(--v7-soft,#DCE6FA);color:var(--v7-ink,#16213A);border:1px solid var(--v7-line,#D6E0F0)}
.ux-resume .t{flex:1;min-width:0}.ux-resume b{display:block;font-size:15px}.ux-resume small{display:block;color:var(--v7-muted,#4B5B77);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.ux-resume button{min-height:40px;padding:0 14px;border-radius:12px;border:1px solid var(--v7-line,#D6E0F0);background:var(--v7-surface,#fff);color:var(--v7-ink,#16213A);font:inherit;font-weight:700;cursor:pointer}
.ux-resume button.pri{background:var(--v7-accent,#1D4ED8);color:var(--v7-on,#fff);border-color:transparent}
@media (max-width:720px){.ux-resume{margin:10px 12px 0;flex-wrap:wrap}.ux-resume .t{flex:1 1 100%}.ux-resume button{flex:1}}
/* recent files on the first screen */
.ux-recent{margin:18px 0 6px}.ux-recent .hd{display:flex;align-items:center;gap:10px;margin:0 0 8px}.ux-recent h2{margin:0;font-size:16px;font-weight:800;flex:1}
.ux-recent .hd button{border:0;background:none;color:var(--v7-muted,#4B5B77);font:inherit;font-size:13px;cursor:pointer;padding:4px 6px}
.ux-recent .ls{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:8px}
.ux-rf{position:relative;display:flex;align-items:center;gap:10px;padding:8px 10px;border-radius:14px;border:1px solid var(--v7-line,#D6E0F0);background:var(--v7-surface,#fff);color:var(--v7-ink,#16213A);cursor:pointer;text-align:start;font:inherit;min-height:56px}
.ux-rf:hover{border-color:var(--v7-dash,#8FAEEB)}.ux-rf:focus-visible{outline:3px solid var(--v7-accent,#1D4ED8);outline-offset:2px}
.ux-rf .n{flex:1;min-width:0}.ux-rf b{display:block;font-size:14px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;unicode-bidi:plaintext;text-align:start}.ux-rf small{color:var(--v7-muted,#4B5B77);font-size:12px}
.ux-rf .x{width:30px;height:30px;border-radius:50%;border:0;background:transparent;color:var(--v7-muted,#4B5B77);font-size:18px;cursor:pointer;flex:none;display:grid;place-items:center}
.ux-rf .x:hover{background:var(--v7-soft,#DCE6FA)}
@media (max-width:720px){.ux-recent .ls{grid-template-columns:1fr 1fr}.ux-rf{padding:8px}.ux-rf .x{display:none}}
/* redo button */
#redoBtn{font-size:inherit}
/* phone: floating buttons out of the way, toasts above the dock, version hidden */
@media (max-width:720px){
  .sb-back,.a11y-fab{width:44px!important;height:44px!important;opacity:.92}
  .sb-back svg{width:22px!important;height:22px!important}
  .brand .ver{display:none!important}
  html.v7-doc header.bar .group{gap:4px!important}
  html.v7-doc #redoBtn{display:none!important}
  html.v7-doc #toast,html.v7.v7-doc #toast{bottom:calc(84px + env(safe-area-inset-bottom,0px))!important}
  #aiPanel .ai-ask{padding-inline-end:64px!important;padding-bottom:12px}
  #xed .xe-bottom,#ped .pe-status,#wed .we-status{padding-inline-end:60px!important}
  .pgpanel .pg-tools{padding-inline-end:58px}
  /* PowerPoint: the help line in the status bar wraps as a sentence, not word by word */
  #ped .pe-status{flex-wrap:wrap;row-gap:2px}#ped .pe-hint{flex:1 1 100%;white-space:normal;line-height:1.35;min-width:0}
  /* editors: a tighter top on a small screen; the file strip hides while one file is open in an editor */
  body.ed-live.ux-onefile #tabs{display:none!important}
  html body.ed-live.ux-onefile #wed,html body.ed-live.ux-onefile #xed,html body.ed-live.ux-onefile #ped,html body.ed-live.ux-onefile #aiPanel{top:0!important}
  #wed .we-top,#xed .xe-top,#ped .pe-top{min-height:42px;padding:4px 10px}
  #wed .wpanel,#xed .xpanel,#ped .ppanel{min-height:76px;padding:4px 6px 2px}
}
/* Word on a phone: the text wraps to the screen (the document's page size is untouched) */
#wed.ux-reflow #weSheet{width:100%!important;max-width:100%!important;box-sizing:border-box!important;min-height:60vh!important;padding:14px 12px 40px!important;box-shadow:none!important;zoom:1!important;align-self:stretch}
#wed.ux-reflow #weHdr,#wed.ux-reflow #weFtr,#wed.ux-reflow #weGuides{display:none!important}
#wed.ux-reflow .we-canvas{padding:0!important;width:100%!important;min-width:0!important;max-width:100%!important;display:block!important}
#wed.ux-reflow .we-scroll{overflow-x:hidden!important}
#wed.ux-reflow #weBody{max-width:100%!important}
#wed.ux-reflow #weBody img,#wed.ux-reflow #weBody table{max-width:100%!important}
#wed .ux-reflow-btn{display:none;min-height:30px;padding:0 10px;border-radius:9px;border:1px solid var(--wl,#D4DBE5);background:var(--bar,#fff);color:inherit;font:inherit;font-size:12px;font-weight:700;cursor:pointer}
@media (max-width:760px){#wed .ux-reflow-btn{display:inline-flex;align-items:center}#wed.ux-reflow #weZoom,#wed.ux-reflow .we-status .wb,#wed.ux-reflow #weZoomS{display:none}}
/* the "more" menu in the phone bar */
.ux-more{display:none}
@media (max-width:720px){
  html.v7-doc .ux-more{display:inline-flex;align-items:center;justify-content:center;width:44px;height:44px;border-radius:12px;border:1px solid var(--v7-line,#D6E0F0);background:var(--v7-surface,#fff);color:var(--v7-ink,#16213A);font-size:22px;line-height:1;cursor:pointer;flex:none}
  .ux-more-pop{position:fixed;z-index:2147482600;inset-inline-end:10px;top:calc(60px + env(safe-area-inset-top,0px));width:min(300px,calc(100vw - 20px));padding:6px;border-radius:16px;background:var(--v7-surface,#fff);color:var(--v7-ink,#16213A);border:1px solid var(--v7-line,#D6E0F0);box-shadow:0 18px 40px -14px rgba(0,0,0,.45)}
  .ux-more-pop button{display:flex;align-items:center;gap:12px;width:100%;min-height:46px;padding:0 12px;border:0;border-radius:11px;background:transparent;color:inherit;font:inherit;font-size:15px;text-align:start;cursor:pointer}
  .ux-more-pop button:active{background:var(--v7-soft,#DCE6FA)}.ux-more-pop button svg{width:20px;height:20px;flex:none}
}
/* scanning on a computer: a real drop zone and clear scan types */
@media (min-width:721px){
  .scn-desk{width:min(760px,86%)!important;gap:18px!important}
  .scn-drop{padding:56px 28px!important;border-width:2.5px!important}
  .scn-drop b{font-size:22px!important}.scn-drop span{font-size:15.5px!important}.scn-drop svg{width:56px!important;height:56px!important}
  .scn-modes{justify-content:center!important;gap:8px!important;padding:12px 8px 6px!important}
  .scn-modes button{font-size:16px!important;padding:9px 18px!important;border:1px solid rgba(255,255,255,.18)!important}
}
`;document.head.append(css);

/* ------------------------------------------------------------ 1. autosave + resume (PDF) */
const DRAFT='draft:pdf';let saveT=null,suppress=false;
async function saveDraft(){
  if(typeof S==='undefined'||!S.doc||!S.bytes||S.office||!S.dirty)return;
  if(S.bytes.length>MAX_FILE)return;
  let anns='';try{anns=histStr();}catch(_){return;}
  // the big strings (signature / image data) travel with the draft
  const big={};try{for(const m of anns.matchAll(/"(\u0001b\d+)"/g)){const v=BIG.get(m[1]);if(v)big[m[1]]=v;}}catch(_){}
  await DB.set(DRAFT,{name:S.name,bytes:S.bytes,anns,big,time:Date.now()});
}
function schedule(){clearTimeout(saveT);saveT=setTimeout(saveDraft,2500);}
function hookDirty(){
  if(typeof setDirty!=='function')return;const orig=setDirty;
  window.setDirty=function(v){orig(v);if(suppress)return;if(v)schedule();else{clearTimeout(saveT);DB.del(DRAFT);}};
  document.addEventListener('visibilitychange',()=>{if(document.visibilityState==='hidden'&&typeof S!=='undefined'&&S.dirty){clearTimeout(saveT);saveDraft();}});
  addEventListener('pagehide',()=>{if(typeof S!=='undefined'&&S.dirty)saveDraft();});
}
async function offerResume(){
  const d=await DB.get(DRAFT);if(!d||!d.bytes||typeof S==='undefined'||S.doc)return;
  const home=document.querySelector('.v7-home')||$('empty');if(!home)return;
  const Lx=L();const bar=document.createElement('div');bar.className='ux-resume';bar.dir=Lx===T.he||Lx===T.ar?'rtl':'ltr';
  bar.innerHTML=`<div class="t"><b>${esc(Lx.resume)}</b><small>${esc(Lx.resumeSub(d.name,ago(d.time)))}</small></div><button type="button" class="pri" data-go>${esc(Lx.go)}</button><button type="button" data-drop>${esc(Lx.drop)}</button>`;
  bar.querySelector('[data-drop]').onclick=()=>{DB.del(DRAFT);bar.remove();};
  bar.querySelector('[data-go]').onclick=async()=>{bar.remove();await resume(d);};
  home.prepend(bar);
}
async function resume(d){
  try{
    const f=new File([d.bytes],d.name,{type:'application/pdf'});
    suppress=true;await openFile(f);
    if(S.name!==d.name){suppress=false;return;}
    try{for(const k in d.big||{}){BIG.set(k,d.big[k]);BIG_ID.set(d.big[k],k);}}catch(_){}
    S.anns=histParse(d.anns||'[]');let mx=0;for(const a of S.anns)if(a.id>mx)mx=a.id;S.nextId=Math.max(S.nextId||1,mx+1);
    S.pages.forEach(P=>{try{drawAnns(P);}catch(_){}});
    suppress=false;setDirty(true);toast(L().restored);
  }catch(e){suppress=false;console.error(e);}
}

/* ------------------------------------------------------------ 2. recent files */
const RECENT='recent';
async function remember(file){
  if(!file||!file.size||file.size>MAX_FILE)return;
  const list=(await DB.get(RECENT))||[];
  let bytes;try{bytes=new Uint8Array(await file.arrayBuffer());}catch(_){return;}
  const rest=list.filter(x=>x.name!==file.name);rest.unshift({name:file.name,size:file.size,time:Date.now(),bytes});
  let total=0;const keep=[];for(const x of rest){if(keep.length>=MAX_N)break;total+=x.size;if(total>MAX_TOTAL)break;keep.push(x);}
  await DB.set(RECENT,keep);renderRecent();
}
function hookOpen(){
  if(typeof openFile!=='function')return;const orig=openFile;
  window.openFile=async function(file,handle){const r=await orig(file,handle);try{if(file&&typeof S!=='undefined'&&(S.name===file.name||/\.(docx|xlsx|pptx|doc)$/i.test(file.name)))remember(file);}catch(_){}return r;};
}
async function renderRecent(){
  const home=document.querySelector('.v7-home');if(!home)return;
  let box=home.querySelector('.ux-recent');const list=(await DB.get(RECENT))||[];
  if(!list.length){if(box)box.remove();return;}
  const Lx=L();if(!box){box=document.createElement('section');box.className='ux-recent';const anchor=home.querySelector('.v7-new');if(anchor)home.insertBefore(box,anchor);else home.append(box);}
  box.innerHTML=`<div class="hd"><h2>${esc(Lx.recent)}</h2><button type="button" data-clear>${esc(Lx.clear)}</button></div><div class="ls">${list.map((x,i)=>`<button type="button" class="ux-rf" data-i="${i}">${ICON[kindOf(x.name)]||ICON.pdf}<span class="n"><b>${esc(x.name)}</b><small>${esc(ago(x.time))} · ${Math.max(1,Math.round(x.size/1024))}KB</small></span><span class="x" role="button" data-x="${i}" title="${esc(Lx.remove)}" aria-label="${esc(Lx.remove)}">×</span></button>`).join('')}</div>`;
  box.querySelector('[data-clear]').onclick=async()=>{await DB.del(RECENT);box.remove();};
  box.querySelectorAll('[data-x]').forEach(x=>x.onclick=async e=>{e.stopPropagation();const i=+x.dataset.x;list.splice(i,1);await DB.set(RECENT,list);renderRecent();});
  box.querySelectorAll('.ux-rf').forEach(b=>b.onclick=async()=>{const x=list[+b.dataset.i];if(!x)return;
    const type=kindOf(x.name)==='pdf'?'application/pdf':'';await openFile(new File([x.bytes],x.name,type?{type}:{}));});
}
// the first screen is rebuilt when the language changes: draw the list again each time
function watchHome(){
  let last=null;const tick=()=>{const h=document.querySelector('.v7-home');if(h&&h!==last){last=h;renderRecent();offerResume();}};
  new MutationObserver(tick).observe(document.body,{childList:true,subtree:true});tick();
}

/* ------------------------------------------------------------ 3. redo button next to undo */
function redoButton(){
  const u=$('undo');if(!u||$('redoBtn'))return;
  const b=document.createElement('button');b.className=u.className;b.id='redoBtn';b.title=L().redo;b.setAttribute('aria-label',L().redo);b.textContent='↷';
  b.onclick=()=>{if(typeof redoAct==='function')redoAct();};u.after(b);
  const sync=()=>{b.disabled=!(typeof S!=='undefined'&&S.redo&&S.redo.length);};setInterval(sync,500);sync();
}

/* ------------------------------------------------------------ 4. the "more" menu in the phone bar */
const MORE_IDS=['trBtn','aiBtn','openBtn','printBtn','convBtn','cmdBtn','v7LangBtn','v7ColorBtn','setBtn','installBtn','admBtn'];
function moreMenu(){
  const grp=$('setBtn')&&$('setBtn').parentElement;if(!grp||$('uxMore'))return;
  const b=document.createElement('button');b.className='ux-more';b.id='uxMore';b.type='button';b.textContent='⋯';b.title=L().more;b.setAttribute('aria-label',L().more);b.setAttribute('aria-haspopup','menu');
  const anchor=$('shareBtn')||$('aiBtn');if(anchor)anchor.after(b);else grp.append(b);
  const st=document.createElement('style');st.textContent=`@media (max-width:720px){html.v7-doc ${MORE_IDS.map(i=>'#'+i).join(',html.v7-doc ')}{display:none!important}}`;document.head.append(st);
  let pop=null;const closePop=()=>{if(pop){pop.remove();pop=null;b.setAttribute('aria-expanded','false');}};
  b.onclick=e=>{e.stopPropagation();if(pop){closePop();return;}
    pop=document.createElement('div');pop.className='ux-more-pop';pop.setAttribute('role','menu');pop.dir=document.documentElement.dir||'rtl';
    const items=MORE_IDS.map(id=>$(id)).filter(x=>x&&!x.classList.contains('hidden')&&!x.disabled&&!x.hidden);
    for(const x of items){const it=document.createElement('button');it.type='button';it.setAttribute('role','menuitem');
      const ic=x.querySelector('svg');const cl=x.cloneNode(true);cl.querySelectorAll('kbd,svg').forEach(k=>k.remove());const txt=(cl.textContent||'').replace(/\s+/g,' ').trim();it.innerHTML=(ic?ic.outerHTML:'')+'<span>'+esc(txt||x.getAttribute('aria-label')||x.title||'')+'</span>';
      it.onclick=()=>{closePop();x.click();};pop.append(it);}
    document.body.append(pop);b.setAttribute('aria-expanded','true');
    setTimeout(()=>document.addEventListener('click',closePop,{once:true}),0);};
  addEventListener('keydown',e=>{if(e.key==='Escape')closePop();});
}

/* ------------------------------------------------------------ 5. Word on a phone: phone view by default */
function wordReflow(){
  const obs=new MutationObserver(()=>{const wed=$('wed');if(!wed||wed.querySelector('.ux-reflow-btn'))return;const st=wed.querySelector('.we-status');if(!st)return;
    const b=document.createElement('button');b.type='button';b.className='ux-reflow-btn';b.title=L().reflowTip;
    const paint=()=>{b.textContent=wed.classList.contains('ux-reflow')?L().reflowOff:L().reflowOn;b.setAttribute('aria-pressed',wed.classList.contains('ux-reflow'));};
    b.onclick=()=>{wed.classList.toggle('ux-reflow');try{localStorage.setItem('siman.ux.reflow',wed.classList.contains('ux-reflow')?'1':'0');}catch(_){}paint();};
    let on='1';try{on=localStorage.getItem('siman.ux.reflow')||'1';}catch(_){}
    if(narrow()&&on==='1')wed.classList.add('ux-reflow');paint();st.append(b);});
  obs.observe(document.body,{childList:true,subtree:true});
}

/* ------------------------------------------------------------ go */
function oneFile(){const f=()=>{const t=$('tabs');document.body.classList.toggle('ux-onefile',!!t&&t.querySelectorAll('.tab').length<2);};
  new MutationObserver(f).observe(document.body,{childList:true,subtree:true,attributes:true,attributeFilter:['class']});f();}
function init(){oneFile();hookDirty();hookOpen();watchHome();redoButton();moreMenu();wordReflow();
  new MutationObserver(()=>{const r=$('redoBtn');if(r){r.title=L().redo;r.setAttribute('aria-label',L().redo);}const m=$('uxMore');if(m){m.title=L().more;m.setAttribute('aria-label',L().more);}}).observe(document.documentElement,{attributes:true,attributeFilter:['lang']});}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
