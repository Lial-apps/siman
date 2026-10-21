/* סימן — שליחה לחתימה ושיתוף בקישור.
   השולח מסמן איפה כל אחד חותם, מקבל קישור לכל חותם ושולח אותו בוואטסאפ או במייל.
   החותם פותח את הקישור בטלפון או במחשב, בלי הרשמה ובלי התקנה, וחותם.
   המסמך החתום נבנה בדפדפן: המקור + החתימות + עמוד תיעוד (מי חתם, מתי, מאיזה מכשיר).
   Needs: pdf.js (pdfjsLib), pdf-lib (PDFLib). */
(function(){
'use strict';
const API='https://siman-share.siman-app.workers.dev/docs';
const HOST=API.replace(/\/docs$/,'');
const LIST_KEY='siman.sent';
const COLORS=['#2E58CC','#D9480F','#1F8A4C','#7A3EB1','#C2255C','#0B7285','#8F5B00','#495057','#5F3DC4','#2B8A3E'];
const TYPES={sig:{l:'חתימה',w:.26,h:.06},initials:{l:'ראשי תיבות',w:.12,h:.05},name:{l:'שם מלא',w:.24,h:.035},date:{l:'תאריך',w:.16,h:.035},text:{l:'טקסט',w:.26,h:.035}};
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[c]);
const pdfjs=()=>{const p=window.pdfjsLib||window['pdfjs-dist/build/pdf'];if(p&&!p.GlobalWorkerOptions.workerSrc)p.GlobalWorkerOptions.workerSrc='pdf.worker.min.js';return p;};
const fmtDate=d=>{const x=new Date(d);return isNaN(x)?'':x.toLocaleString('he-IL',{day:'numeric',month:'numeric',year:'numeric',hour:'2-digit',minute:'2-digit'});};
const today=()=>new Date().toLocaleDateString('he-IL',{day:'2-digit',month:'2-digit',year:'numeric'});

/* ------------------------------------------------------------ server */
async function api(path,opt={}){
  let r;try{r=await fetch(API+path,opt);}catch(_){throw new Error('אין חיבור לאינטרנט, או שהשירות לא זמין כרגע');}
  if(!r.ok){let m='';try{m=(await r.json()).error;}catch(_){}throw new Error(m||('שגיאה '+r.status));}
  return /json/.test(r.headers.get('Content-Type')||'')?r.json():new Uint8Array(await r.arrayBuffer());
}
const qs=a=>a.k?`?k=${encodeURIComponent(a.k)}`:a.t?`?t=${encodeURIComponent(a.t)}`:`?v=${encodeURIComponent(a.v)}`;
const getMeta=(id,a)=>api(`/${id}${qs(a)}`);
const getPdf=(id,a)=>api(`/${id}/pdf${qs(a)}`);
const signLink=(id,t)=>`${HOST}/s/${id}/${t}`;
const viewLink=(id,v)=>`${HOST}/v/${id}/${v}`;
const ownerLink=(id,k)=>`${HOST}/o/${id}#k=${k}`;

/* ------------------------------------------------------------ the signed document */
// draws a field's content into a picture with exactly the field's proportions
async function fieldPng(f,val,wPt,hPt){
  const S=4,c=document.createElement('canvas');c.width=Math.max(8,Math.round(wPt*S));c.height=Math.max(8,Math.round(hPt*S));const g=c.getContext('2d');
  if(val.png){const img=new Image();img.src=val.png;await img.decode();const k=Math.min(c.width/img.width,c.height/img.height);const w=img.width*k,h=img.height*k;
    g.drawImage(img,(c.width-w)/2,(c.height-h)/2,w,h);}
  else{const t=String(val.text||'');let fs=c.height*.72;g.font=`${fs}px "Heebo","Arial",sans-serif`;while(fs>6&&g.measureText(t).width>c.width*.98){fs-=1;g.font=`${fs}px "Heebo","Arial",sans-serif`;}
    g.fillStyle='#111';g.textBaseline='middle';g.direction=/[\u0590-\u05FF]/.test(t)?'rtl':'ltr';g.textAlign=g.direction==='rtl'?'right':'left';g.fillText(t,g.direction==='rtl'?c.width-2:2,c.height/2);}
  return new Uint8Array(await (await new Promise(r=>c.toBlob(r,'image/png'))).arrayBuffer());
}
function auditCanvas(meta){
  const W=1190,H=1684,c=document.createElement('canvas');c.width=W;c.height=H;const g=c.getContext('2d');g.fillStyle='#fff';g.fillRect(0,0,W,H);
  g.direction='rtl';g.textAlign='right';const R=W-90;let y=110;
  const line=(t,sz,col,bold,dy)=>{g.font=`${bold?'700 ':''}${sz}px "Heebo","Arial",sans-serif`;g.fillStyle=col||'#16213A';g.fillText(t,R,y);y+=dy||sz*1.6;};
  g.fillStyle='#16213A';g.fillRect(0,0,W,24);g.fillStyle='#FFD43B';g.fillRect(R-120,52,120,8);
  line('תעודת חתימה',46,'#16213A',true,80);
  line(meta.title,30,'#16213A',true,52);
  line(`מזהה המסמך: ${meta.id}`,22,'#5D6982');
  line(`נשלח על ידי: ${meta.from&&meta.from.name||'—'}${meta.from&&meta.from.mail?` (${meta.from.mail})`:''}`,22,'#5D6982');
  line(`נוצר: ${fmtDate(meta.created)}`,22,'#5D6982');
  g.font='18px monospace';g.fillStyle='#5D6982';g.direction='ltr';g.textAlign='right';g.fillText('SHA-256: '+meta.hash,R,y);y+=36;g.direction='rtl';g.textAlign='right';
  g.fillStyle='#D4DBE5';g.fillRect(90,y,W-180,2);y+=50;
  for(const s of meta.signers){
    const st=s.status==='signed'?'✓ חתם/ה':s.status==='declined'?'✗ סירב/ה לחתום':'ממתין/ה לחתימה';
    line(`${s.name} — ${st}`,28,s.status==='signed'?'#1F8A4C':s.status==='declined'?'#C2362B':'#8F5B00',true,44);
    if(s.at)line(`מועד: ${fmtDate(s.at)}${s.tz?` (${s.tz})`:''}`,20,'#5D6982',false,32);
    if(s.ip)line(`כתובת IP: ${s.ip}${s.country?` · ${s.country}`:''}`,20,'#5D6982',false,32);
    if(s.ua)line(`מכשיר: ${devName(s.ua)}`,20,'#5D6982',false,32);
    if(s.reason)line(`סיבה: ${s.reason}`,20,'#5D6982',false,32);
    y+=18;if(y>H-260)break;
  }
  y=H-200;g.fillStyle='#D4DBE5';g.fillRect(90,y,W-180,2);y+=50;
  line('המסמך נחתם באמצעות סימן, עורך PDF בעברית · lial-apps.github.io/siman',20,'#5D6982',false,32);
  line('כל חותם אישר שהוא מסכים לחתום על המסמך באופן אלקטרוני. הפרטים בעמוד זה נרשמו בזמן החתימה.',18,'#7A879E',false,30);
  return c;
}
function devName(ua){ua=String(ua||'');const os=/iPhone/.test(ua)?'iPhone':/iPad/.test(ua)?'iPad':/Android/.test(ua)?'Android':/Windows/.test(ua)?'Windows':/Mac OS/.test(ua)?'Mac':/Linux/.test(ua)?'Linux':'';
  const br=/Edg\//.test(ua)?'Edge':/SamsungBrowser/.test(ua)?'Samsung Internet':/Chrome\//.test(ua)?'Chrome':/Firefox\//.test(ua)?'Firefox':/Safari\//.test(ua)?'Safari':'';return [br,os].filter(Boolean).join(' ב-')||'לא ידוע';}
async function buildSigned(meta,pdfBytes){
  const{PDFDocument,degrees}=PDFLib;const doc=await PDFDocument.load(pdfBytes,{ignoreEncryption:true});const pages=doc.getPages();
  for(const s of meta.signers){if(s.status!=='signed'||!s.values)continue;
    for(const f of meta.fields.filter(x=>x.signer===s.n)){const val=s.values[f.id];const pg=pages[f.page];if(!val||!pg||!f.r)continue;
      const rot=((pg.getRotation().angle%360)+360)%360;const[x1,y1,x2,y2]=f.r;const minX=Math.min(x1,x2),maxX=Math.max(x1,x2),minY=Math.min(y1,y2),maxY=Math.max(y1,y2);const Wr=maxX-minX,Hr=maxY-minY;
      const side=rot===90||rot===270;const img=await doc.embedPng(await fieldPng(f,val,side?Hr:Wr,side?Wr:Hr));
      const o=rot===90?{x:maxX,y:minY,width:Hr,height:Wr,rotate:degrees(90)}:rot===180?{x:maxX,y:maxY,width:Wr,height:Hr,rotate:degrees(180)}:rot===270?{x:minX,y:maxY,width:Hr,height:Wr,rotate:degrees(270)}:{x:minX,y:minY,width:Wr,height:Hr};
      pg.drawImage(img,o);}}
  const c=auditCanvas(meta);const png=await doc.embedPng(new Uint8Array(await (await new Promise(r=>c.toBlob(r,'image/png'))).arrayBuffer()));
  const p=doc.addPage([595.28,841.89]);p.drawImage(png,{x:0,y:0,width:595.28,height:841.89});
  doc.setProducer('Siman · lial-apps.github.io/siman');doc.setKeywords(['siman',meta.id]);doc.setModificationDate(new Date());
  return doc.save();
}

/* ------------------------------------------------------------ the list of documents I sent */
const listGet=()=>{try{return JSON.parse(localStorage.getItem(LIST_KEY)||'[]');}catch(_){return[];}};
const listSet=l=>{try{localStorage.setItem(LIST_KEY,JSON.stringify(l.slice(0,200)));}catch(_){}};

/* ------------------------------------------------------------ the screen */
const CSS=`
#snd{--n:#16213A;--m:#FFD43B;--b:#2E58CC;--l:#D4DBE5;--t:#16213A;--mu:#5D6982;position:fixed;inset:0;z-index:950;display:flex;flex-direction:column;background:#E4E8EF;color:var(--t);direction:rtl;font:14px/1.4 "SF-heebo","Heebo",system-ui,sans-serif}
#snd *{box-sizing:border-box}
#snd .s-top{display:flex;align-items:center;gap:10px;padding:8px 14px;background:var(--n);color:#E4EAF4;min-height:52px}
#snd .s-top b{font-size:16px}#snd .s-top .grow{flex:1}
#snd .s-x{background:none;border:0;color:#C9D3E6;font-size:26px;line-height:1;cursor:pointer;padding:2px 8px;border-radius:8px}#snd .s-x:hover{background:rgba(255,255,255,.1)}
#snd .s-btn{border:1px solid var(--l);background:#fff;color:var(--t);border-radius:9px;padding:8px 14px;font:inherit;cursor:pointer;display:inline-flex;align-items:center;gap:6px;white-space:nowrap}
#snd .s-btn:hover{background:#F3F5F9}#snd .s-btn.pri{background:var(--m);border-color:var(--m);color:var(--n);font-weight:700}#snd .s-btn.pri:hover{filter:brightness(.96)}
#snd .s-btn.dark{background:var(--n);border-color:var(--n);color:#fff}#snd .s-btn:disabled{opacity:.45;cursor:default}
#snd .s-btn.sm{padding:5px 10px;font-size:13px;border-radius:7px}#snd .s-btn.wa{background:#25D366;border-color:#25D366;color:#073b1c;font-weight:600}
#snd .s-main{flex:1;display:flex;min-height:0}
#snd .s-side{width:300px;flex:none;background:#fff;border-inline-start:1px solid var(--l);overflow:auto;padding:16px;display:flex;flex-direction:column;gap:14px}
#snd .s-side h3{margin:0;font-size:13px;color:var(--mu);font-weight:600}
#snd .s-pages{flex:1;overflow:auto;padding:24px 16px 80px;display:flex;flex-direction:column;align-items:center;gap:18px}
#snd .s-page{position:relative;background:#fff;box-shadow:0 12px 30px -14px rgba(22,33,58,.45);flex:none}
#snd .s-page canvas{display:block;width:100%;height:100%}
#snd .s-lay{position:absolute;inset:0}#snd.arm .s-lay{cursor:crosshair}
#snd .s-pn{position:absolute;bottom:-20px;inset-inline-start:0;font-size:11px;color:var(--mu)}
#snd .s-f{position:absolute;border:2px solid;border-radius:5px;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:600;cursor:move;user-select:none;touch-action:none;overflow:hidden}
#snd .s-f.on{box-shadow:0 0 0 3px rgba(46,88,204,.25)}#snd .s-f span{pointer-events:none;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;padding:0 4px}
#snd .s-f .del{position:absolute;top:-1px;inset-inline-end:-1px;width:18px;height:18px;border:0;border-radius:0 3px 0 6px;color:#fff;font-size:13px;line-height:18px;cursor:pointer;padding:0;display:none}
#snd .s-f.on .del{display:block}#snd .s-f .rs{position:absolute;bottom:-2px;inset-inline-start:-2px;width:12px;height:12px;cursor:nesw-resize;display:none}#snd .s-f.on .rs{display:block}
#snd .sg{display:flex;align-items:center;gap:8px;padding:6px 8px;border:1px solid var(--l);border-radius:9px;cursor:pointer}
#snd .sg.on{border-color:var(--b);box-shadow:0 0 0 2px rgba(46,88,204,.15)}#snd .sg i{width:12px;height:12px;border-radius:50%;flex:none}
#snd .sg input{flex:1;min-width:0;border:0;font:inherit;background:none;outline:none;padding:4px 0}
#snd .sg button{border:0;background:none;color:#8E9AB3;cursor:pointer;font-size:16px;padding:0 4px}#snd .sg small{color:var(--mu);font-size:11px;white-space:nowrap}
#snd .tools{display:grid;grid-template-columns:1fr 1fr;gap:6px}
#snd .tool{border:1px dashed #9AA6B8;background:#F7F8FB;border-radius:8px;padding:9px 6px;cursor:pointer;font:inherit;font-size:13px;color:var(--t)}
#snd .tool.on{border-style:solid;border-color:var(--b);background:#EAF0FD;color:var(--b);font-weight:600}
#snd textarea,#snd .s-in{width:100%;border:1px solid var(--l);border-radius:8px;padding:8px 10px;font:inherit;resize:vertical}
#snd .ck{display:flex;gap:8px;align-items:center;font-size:13px}
#snd .hint{font-size:12px;color:var(--mu);line-height:1.5;margin:0}
#snd .s-wide{flex:1;overflow:auto;padding:28px 16px 60px}#snd .s-card{max-width:680px;margin:0 auto 14px;background:#fff;border-radius:14px;padding:18px 20px;box-shadow:0 8px 24px -14px rgba(22,33,58,.35)}
#snd .s-card h2{margin:0 0 6px;font-size:20px}#snd .s-card h4{margin:0 0 8px;font-size:15px;display:flex;align-items:center;gap:8px}
#snd .lnk{display:flex;gap:6px;flex-wrap:wrap;align-items:center;margin-top:8px}#snd .lnk input{flex:1;min-width:180px;direction:ltr;font:12.5px monospace;border:1px solid var(--l);border-radius:8px;padding:8px;background:#F7F8FB}
#snd .chip{display:inline-flex;align-items:center;gap:5px;border-radius:20px;padding:3px 10px;font-size:12.5px;background:#F1F3F7;margin:3px 0 0 4px}
#snd .chip.ok{background:#E3F4EA;color:#1F6B3B}#snd .chip.no{background:#FBE7E5;color:#A02B22}#snd .chip.wait{background:#FFF4D6;color:#7A5300}
#snd .bar{height:6px;border-radius:4px;background:#EEF1F5;overflow:hidden;margin:8px 0}#snd .bar i{display:block;height:100%;background:#1F8A4C}
#snd .empty{text-align:center;color:var(--mu);padding:40px 10px}
#snd .s-toast{position:fixed;left:50%;transform:translateX(-50%);bottom:calc(22px + env(safe-area-inset-bottom));z-index:960;background:#16213A;color:#fff;border-inline-start:4px solid #FFD43B;padding:11px 18px;border-radius:10px;box-shadow:0 12px 30px -10px rgba(0,0,0,.5);max-width:92vw;font-size:14px;transition:opacity .2s;opacity:0;pointer-events:none}
#snd .s-toast.on{opacity:1}#snd .s-toast.err{border-color:#E5484D}
#snd .tool.combo{grid-column:1/-1;border-style:solid;background:#FFF8DB;border-color:#E9C46A}
#snd .s-load{position:absolute;inset:0;display:grid;place-items:center;background:rgba(228,232,239,.85);font-size:15px;z-index:5}
@media (max-width:760px){#snd .s-main{flex-direction:column-reverse}#snd .s-side{width:auto;max-height:46vh;border:0;border-top:1px solid var(--l)}#snd .s-pages{padding:14px 8px 40px}}
`;
let U=null; // the open screen
function shell(title,actions){
  if(!document.getElementById('sndCss')){const s=document.createElement('style');s.id='sndCss';s.textContent=CSS;document.head.append(s);}
  close();const root=document.createElement('div');root.id='snd';root.setAttribute('role','dialog');root.setAttribute('aria-label',title);
  root.innerHTML=`<div class="s-top"><b>${esc(title)}</b><span class="grow"></span>${actions||''}<button class="s-x" data-x aria-label="סגירה">×</button></div><div class="s-body" style="flex:1;display:flex;flex-direction:column;min-height:0;position:relative"></div>`;
  document.body.append(root);U={root,body:root.querySelector('.s-body')};
  root.querySelector('[data-x]').onclick=()=>{if(U&&U.dirty&&!confirm('לצאת בלי לשלוח? מה שסימנת לא יישמר.'))return;close();};
  U.onKey=e=>{if(!U)return;e.stopImmediatePropagation();
    if(e.key==='Escape'&&!U.dirty){close();return;}
    if(U.onDel)U.onDel(e);};
  window.addEventListener('keydown',U.onKey,true);
  return U;
}
function close(){if(!U)return;window.removeEventListener('keydown',U.onKey,true);if(U.cleanup)U.cleanup();U.root.remove();U=null;}
function say(msg,err){
  if(!U){alert(msg);return;}let t=U.root.querySelector('.s-toast');if(!t){t=document.createElement('div');t.className='s-toast';t.setAttribute('role','status');U.root.append(t);}
  t.textContent=msg;t.classList.toggle('err',!!err);t.classList.add('on');clearTimeout(t._h);t._h=setTimeout(()=>t.classList.remove('on'),err?6000:3500);
}
const busy=(on,msg)=>{if(!U)return;let l=U.body.querySelector('.s-load');if(on){if(!l){l=document.createElement('div');l.className='s-load';U.body.append(l);}l.textContent=msg||'רגע…';}else if(l)l.remove();};

/* ---------- step 1: who signs and where */
async function open(opts){
  const{bytes,name}=opts;const title=String(name||'מסמך').replace(/\.pdf$/i,'');
  const u=shell('שליחה לחתימה',`<button class="s-btn pri" data-go>יצירת קישורים לחתימה</button>`);
  u.opts=opts;u.dirty=true;u.signers=[{name:''}];u.cur=0;u.fields=[];u.tools=new Set(['sig']);u.sel=null;
  if(opts.selfName)u.signers=[{name:''}];
  u.body.innerHTML=`<div class="s-main"><div class="s-pages" tabindex="0" aria-label="עמודי המסמך"></div><aside class="s-side">
    <label><h3>שם המסמך</h3><input class="s-in" data-title value="${esc(title)}" maxlength="120" dir="auto"></label>
    <div><h3 style="margin-bottom:8px">מי חותם?</h3><div data-signers style="display:flex;flex-direction:column;gap:6px"></div>
      <button class="s-btn sm" data-add style="margin-top:8px">+ הוספת חותם</button>
      <label class="ck" style="margin-top:10px"><input type="checkbox" data-seq> לחתום לפי הסדר (כל אחד אחרי הקודם)</label></div>
    <div><h3 style="margin-bottom:8px">מה להוסיף בשביל <span data-who></span>?</h3><div class="tools"><button class="tool combo" data-combo>חתימה + שם מלא + תאריך</button>${Object.entries(TYPES).map(([k,v])=>`<button class="tool${k==='sig'?' on':''}" data-tool="${k}" aria-pressed="${k==='sig'}">${v.l}</button>`).join('')}</div>
      <p class="hint" style="margin-top:8px">אפשר לבחור כמה שדות יחד. לוחצים על המסמך, והשדות שנבחרו נכנסים אחד מתחת לשני. אפשר לגרור, להגדיל ולמחוק.</p>
      <button class="s-btn sm" data-auto style="margin-top:8px">הצבה מהירה: חתימה, שם ותאריך בסוף המסמך לכל אחד</button></div>
    <label><h3>הודעה לחותמים (לא חובה)</h3><textarea data-note rows="2" maxlength="500" dir="auto" placeholder="למשל: נא לחתום עד יום חמישי. תודה!"></textarea></label>
    <p class="hint">המסמך נשמר בשרת של סימן 30 יום ונמחק לבד. אפשר למחוק אותו קודם, מ"המסמכים ששלחתי".</p>
  </aside></div>`;
  const P=u.body.querySelector('.s-pages');
  busy(true,'טוען את המסמך…');
  try{
    const pdf=await pdfjs().getDocument({data:bytes.slice(0)}).promise;u.pdf=pdf;u.pages=[];
    const avail=Math.min(900,P.clientWidth-32);
    for(let i=1;i<=pdf.numPages;i++){const pg=await pdf.getPage(i);const v1=pg.getViewport({scale:1});const sc=avail/v1.width;const vp=pg.getViewport({scale:sc});
      const box=document.createElement('div');box.className='s-page';box.style.width=vp.width+'px';box.style.height=vp.height+'px';box.dataset.p=i-1;
      const cv=document.createElement('canvas');cv.dir='ltr';cv.style.direction='ltr';const dpr=Math.min(2,devicePixelRatio||1);cv.width=vp.width*dpr;cv.height=vp.height*dpr;box.append(cv);
      const lay=document.createElement('div');lay.className='s-lay';box.append(lay);const pn=document.createElement('span');pn.className='s-pn';pn.textContent=`עמוד ${i} מתוך ${pdf.numPages}`;box.append(pn);P.append(box);
      u.pages.push({vp,box,lay,page:pg});
      pg.render({canvasContext:cv.getContext('2d'),viewport:pg.getViewport({scale:sc*dpr})});
    }
  }catch(e){busy(false);u.dirty=false;close();throw new Error('לא הצלחתי לקרוא את המסמך: '+(e.message||e));}
  busy(false);
  const side=u.body.querySelector('.s-side');
  const drawSigners=()=>{
    const box=side.querySelector('[data-signers]');
    box.innerHTML=u.signers.map((s,i)=>`<div class="sg${i===u.cur?' on':''}" data-i="${i}"><i style="background:${COLORS[i%COLORS.length]}"></i><input value="${esc(s.name)}" placeholder="השם של חותם ${i+1}" maxlength="60" dir="auto" data-name="${i}"><small>${u.fields.filter(f=>f.signer===i).length} שדות</small>${u.signers.length>1?`<button data-del="${i}" title="הסרה" aria-label="הסרת החותם">×</button>`:''}</div>`).join('');
    side.querySelector('[data-who]').textContent=u.signers[u.cur].name||`חותם ${u.cur+1}`;side.querySelector('[data-who]').style.color=COLORS[u.cur%COLORS.length];
  };
  const fieldEl=f=>{const pgx=u.pages[f.page];const el=document.createElement('div');el.className='s-f';const col=COLORS[f.signer%COLORS.length];
    el.style.cssText=`left:${f.x*100}%;top:${f.y*100}%;width:${f.w*100}%;height:${f.h*100}%;border-color:${col};background:${col}1F;color:${col}`;
    el.innerHTML=`<span>${esc(TYPES[f.type].l)} · ${esc(u.signers[f.signer].name||'חותם '+(f.signer+1))}</span><button class="del" style="background:${col}" aria-label="מחיקה">×</button><i class="rs" style="background:${col}"></i>`;
    f.el=el;pgx.lay.append(el);return el;};
  const redraw=()=>{for(const p of u.pages)p.lay.innerHTML='';for(const f of u.fields){fieldEl(f);if(f===u.sel)f.el.classList.add('on');}drawSigners();};
  const select=f=>{u.sel=f;for(const x of u.fields)x.el&&x.el.classList.toggle('on',x===f);};
  const add=(page,x,y,type,signer)=>{const T=TYPES[type];const pgx=u.pages[page];const ar=pgx.vp.width/pgx.vp.height;
    const w=T.w,h=T.h*ar;const f={page,type,signer,x:Math.max(0,Math.min(1-w,x-w/2)),y:Math.max(0,Math.min(1-h,y-h/2)),w,h};u.fields.push(f);fieldEl(f);select(f);drawSigners();return f;};
  // several fields at once: the first where the click was, the rest one under the other
  const ORDER=['sig','initials','name','date','text'];
  const placeGroup=(page,x,y,types,signer)=>{types=ORDER.filter(t=>types.includes(t));let top=null;
    for(const t of types){const f=add(page,x,top===null?y:0,t,signer);if(top!==null){f.y=Math.min(1-f.h,top);f.el.style.top=f.y*100+'%';}top=f.y+f.h+.006;}};
  side.addEventListener('click',e=>{
    const hit=q=>{const x=e.target.closest(q);return x&&side.contains(x)?x:null;};
    const paint=()=>side.querySelectorAll('[data-tool]').forEach(b=>{const on=u.tools.has(b.dataset.tool);b.classList.toggle('on',on);b.setAttribute('aria-pressed',on);});
    if(hit('[data-combo]')){u.tools=new Set(['sig','name','date']);paint();say('נבחרו חתימה, שם מלא ותאריך. לוחצים על המסמך במקום החתימה');return;}
    const t=hit('[data-tool]');if(t){const k=t.dataset.tool;if(u.tools.has(k)){if(u.tools.size>1)u.tools.delete(k);}else u.tools.add(k);paint();return;}
    const d=hit('[data-del]');if(d){const i=+d.dataset.del;u.fields=u.fields.filter(f=>f.signer!==i);u.fields.forEach(f=>{if(f.signer>i)f.signer--;});u.signers.splice(i,1);u.cur=Math.min(u.cur,u.signers.length-1);redraw();return;}
    const r=hit('.sg');if(r){const i=+r.dataset.i;if(i!==u.cur){u.cur=i;side.querySelectorAll('.sg').forEach(x=>x.classList.toggle('on',+x.dataset.i===i));
      const w=side.querySelector('[data-who]');w.textContent=u.signers[i].name||`חותם ${i+1}`;w.style.color=COLORS[i%COLORS.length];}
      if(e.target.tagName!=='INPUT')side.querySelector(`[data-name="${i}"]`).focus();return;}
    if(hit('[data-add]')){if(u.signers.length>=10)return;u.signers.push({name:''});u.cur=u.signers.length-1;drawSigners();side.querySelector(`[data-name="${u.cur}"]`).focus();return;}
    if(hit('[data-auto]')){const last=u.pages.length-1;const n=u.signers.length;
      u.signers.forEach((s,i)=>{if(u.fields.some(f=>f.signer===i&&f.type==='sig'))return;const col=n<=2?(i===0?.75:.25):(1-(i%3+.5)/3);const row=Math.floor(i/(n<=2?2:3));
        placeGroup(last,col,.78-row*.13,['sig','name','date'],i);});say('השדות נוספו בעמוד האחרון. אפשר לגרור אותם למקום המדויק');return;}
  });
  side.addEventListener('focusin',e=>{const n=e.target.dataset&&e.target.dataset.name;if(n===undefined||+n===u.cur)return;u.cur=+n;
    side.querySelectorAll('.sg').forEach(x=>x.classList.toggle('on',+x.dataset.i===u.cur));const w=side.querySelector('[data-who]');w.textContent=u.signers[u.cur].name||`חותם ${u.cur+1}`;w.style.color=COLORS[u.cur%COLORS.length];});
  side.addEventListener('input',e=>{const n=e.target.dataset.name;if(n!==undefined){u.signers[+n].name=e.target.value;side.querySelector('[data-who]').textContent=e.target.value||`חותם ${+n+1}`;
    for(const f of u.fields)if(f.signer===+n&&f.el)f.el.querySelector('span').textContent=`${TYPES[f.type].l} · ${e.target.value||'חותם '+(+n+1)}`;}});
  u.root.classList.add('arm');
  // place, move and resize fields
  P.addEventListener('pointerdown',e=>{
    const del=e.target.closest('.del');if(del){const f=u.fields.find(x=>x.el&&x.el.contains(del));u.fields=u.fields.filter(x=>x!==f);f.el.remove();u.sel=null;drawSigners();return;}
    const fe=e.target.closest('.s-f');const box=e.target.closest('.s-page');if(box&&!P.contains(box))return;if(!box)return;const pgx=u.pages[+box.dataset.p];const R=pgx.box.getBoundingClientRect();
    if(fe){const f=u.fields.find(x=>x.el===fe);select(f);e.preventDefault();const rs=e.target.classList.contains('rs');const x0=e.clientX,y0=e.clientY,o={...f};fe.setPointerCapture(e.pointerId);
      const mv=ev=>{const dx=(ev.clientX-x0)/R.width,dy=(ev.clientY-y0)/R.height;
        if(rs){const nw=Math.max(.04,Math.min(1-o.x+ (o.w), o.w-dx));f.x=Math.max(0,o.x+o.w-nw);f.w=o.x+o.w-f.x;f.h=Math.max(.015,Math.min(1-o.y,o.h+dy));}
        else{f.x=Math.max(0,Math.min(1-f.w,o.x+dx));f.y=Math.max(0,Math.min(1-f.h,o.y+dy));}
        Object.assign(fe.style,{left:f.x*100+'%',top:f.y*100+'%',width:f.w*100+'%',height:f.h*100+'%'});};
      const up=()=>{fe.removeEventListener('pointermove',mv);fe.removeEventListener('pointerup',up);};fe.addEventListener('pointermove',mv);fe.addEventListener('pointerup',up);return;}
    placeGroup(+box.dataset.p,(e.clientX-R.left)/R.width,(e.clientY-R.top)/R.height,[...u.tools],u.cur);
  });
  u.onDel=e=>{if((e.key==='Delete'||e.key==='Backspace')&&u.sel&&!/INPUT|TEXTAREA/.test(document.activeElement.tagName)){e.preventDefault();u.sel.el.remove();u.fields=u.fields.filter(x=>x!==u.sel);u.sel=null;drawSigners();}};
  u.cleanup=()=>{if(u.pdf)u.pdf.destroy();};
  drawSigners();
  u.root.querySelector('[data-go]').onclick=async()=>{
    const names=u.signers.map((s,i)=>s.name.trim()||null);
    const miss=names.findIndex(n=>!n);if(miss>=0){u.cur=miss;drawSigners();const inp=side.querySelector(`[data-name="${miss}"]`);inp.focus();say('צריך לכתוב את השם של כל חותם',1);return;}
    const noF=u.signers.findIndex((s,i)=>!u.fields.some(f=>f.signer===i));if(noF>=0){u.cur=noF;drawSigners();say(`עוד לא סימנת איפה ${names[noF]} חותם/ת`,1);return;}
    const fields=u.fields.map(f=>{const vp=u.pages[f.page].vp;const W=vp.width,H=vp.height;const a=vp.convertToPdfPoint(f.x*W,f.y*H),b=vp.convertToPdfPoint((f.x+f.w)*W,(f.y+f.h)*H);
      return{signer:f.signer,page:f.page,type:f.type,x:f.x,y:f.y,w:f.w,h:f.h,r:[a[0],a[1],b[0],b[1]].map(v=>+v.toFixed(2)),req:true};});
    await upload(u,{mode:'sign',title:side.querySelector('[data-title]').value.trim()||title,note:side.querySelector('[data-note]').value.trim(),order:side.querySelector('[data-seq]').checked?'seq':'any',
      signers:names.map(n=>({name:n})),fields,pages:u.pages.length});
  };
}
async function upload(u,meta){
  const o=u.opts;busy(true,'מעלה את המסמך…');
  try{
    const fd=new FormData();fd.append('meta',JSON.stringify({...meta,from:{name:o.sender&&o.sender.name||'',mail:o.sender&&o.sender.mail||''}}));
    fd.append('pdf',new Blob([o.bytes],{type:'application/pdf'}),'doc.pdf');
    const r=await api('',{method:'POST',body:fd});
    const rec={id:r.id,k:r.owner,v:r.view,title:meta.title,mode:meta.mode,created:new Date().toISOString(),expires:r.expires,signers:(r.signers||[]).map(s=>({n:s.n,name:s.name,t:s.token})),note:meta.note||''};
    listSet([rec,...listGet()]);u.dirty=false;busy(false);links(rec,o);
  }catch(e){busy(false);say('השליחה נכשלה: '+(e.message||e),1);}
}

/* ---------- step 2: the links */
function msgFor(rec,s,o){const from=o&&o.sender&&o.sender.name?o.sender.name:'';
  return rec.mode==='view'?`${from?from+' שיתף/ה איתך':'שותף איתך'} את "${rec.title}" בסימן:\n${viewLink(rec.id,rec.v)}`
    :`שלום ${s.name}, ${from?from+' שלח/ה לך':'נשלח אליך'} את "${rec.title}" לחתימה.\nפותחים את הקישור וחותמים, גם מהטלפון. לא צריך להירשם:\n${signLink(rec.id,s.t)}${rec.note?`\n\n${rec.note}`:''}`;}
function links(rec,o){
  o=o||{};const u=shell(rec.mode==='view'?'קישור לשיתוף':'הקישורים לחתימה',`<button class="s-btn" data-list>המסמכים ששלחתי</button>`);
  u.root.querySelector('[data-list]').onclick=()=>list(o);
  const row=(s,link,msg)=>`<div class="s-card" style="padding:14px 18px"><h4>${s?`<i style="width:12px;height:12px;border-radius:50%;background:${COLORS[s.n%COLORS.length]}"></i>${esc(s.name)}`:'הקישור למסמך'}</h4>
    <div class="lnk"><input readonly value="${esc(link)}" aria-label="הקישור"><button class="s-btn sm wa" data-wa="${esc(msg)}">וואטסאפ</button><button class="s-btn sm" data-copy="${esc(link)}">העתקה</button>${navigator.share?`<button class="s-btn sm" data-share="${esc(msg)}">שיתוף</button>`:''}</div></div>`;
  u.body.innerHTML=`<div class="s-wide"><div class="s-card"><h2>✓ ${rec.mode==='view'?'הקישור מוכן':'המסמך מוכן לחתימה'}</h2>
    <p class="hint">${rec.mode==='view'?'כל מי שמקבל את הקישור יכול לראות ולהוריד את המסמך, בלי להירשם.':rec.signers.length>1?'לכל חותם יש קישור משלו. שלחו לכל אחד את הקישור שלו.':'שלחו את הקישור לחותם. הוא יחתום מהטלפון או מהמחשב, בלי להירשם.'}
    ${rec.mode==='sign'?'<br>מה שנחתם מופיע ב"המסמכים ששלחתי", ומשם מורידים את המסמך החתום עם עמוד התיעוד.':''}<br>המסמך יימחק מהשרת ב-${fmtDate(rec.expires).split(',')[0]}.</p></div>
    ${rec.mode==='view'?row(null,viewLink(rec.id,rec.v),msgFor(rec,null,o)):rec.signers.map(s=>row(s,signLink(rec.id,s.t),msgFor(rec,s,o))).join('')}</div>`;
  u.body.addEventListener('click',async e=>{
    if(!u.body.contains(e.target))return;
    const w=e.target.closest('[data-wa]');if(w){window.open('https://wa.me/?text='+encodeURIComponent(w.dataset.wa),'_blank','noopener');return;}
    const c=e.target.closest('[data-copy]');if(c){try{await navigator.clipboard.writeText(c.dataset.copy);c.textContent='הועתק ✓';setTimeout(()=>c.textContent='העתקה',1600);}catch(_){c.previousElementSibling&&c.closest('.lnk').querySelector('input').select();}return;}
    const sh=e.target.closest('[data-share]');if(sh){try{await navigator.share({text:sh.dataset.share});}catch(_){}}
  });
}

/* ---------- the documents I sent */
async function list(o){
  o=o||{};const u=shell('המסמכים ששלחתי');const l=listGet();
  if(!l.length){u.body.innerHTML=`<div class="s-wide"><div class="s-card empty"><h2>עוד לא שלחת מסמכים</h2><p>פותחים מסמך, לוחצים "שיתוף" ובוחרים "שליחה לחתימה בקישור".</p></div></div>`;return;}
  u.body.innerHTML=`<div class="s-wide">${l.map(r=>`<div class="s-card" data-id="${r.id}"><h4 style="justify-content:space-between"><span dir="auto">${esc(r.title)}</span><small style="color:#5D6982;font-weight:400">${fmtDate(r.created)}</small></h4><div data-st><span class="chip">בודק…</span></div></div>`).join('')}</div>`;
  const draw=async r=>{const card=u.body.querySelector(`[data-id="${r.id}"]`);if(!card)return;const st=card.querySelector('[data-st]');
    let m;try{m=await getMeta(r.id,{k:r.k});}catch(e){st.innerHTML=`<p class="hint">${esc(e.message)}</p><div class="lnk"><button class="s-btn sm" data-forget="${r.id}">הסרה מהרשימה</button></div>`;return;}
    r.meta=m;const done=m.signers.filter(s=>s.status==='signed').length,all=m.signers.length;
    st.innerHTML=(m.mode==='sign'?`<div class="bar"><i style="width:${all?done/all*100:0}%"></i></div>${m.signers.map(s=>`<span class="chip ${s.status==='signed'?'ok':s.status==='declined'?'no':'wait'}">${s.status==='signed'?'✓':s.status==='declined'?'✗':'⏳'} ${esc(s.name)}${s.at?` · ${fmtDate(s.at)}`:''}</span>`).join('')}
        ${m.signers.filter(s=>s.reason).map(s=>`<p class="hint">${esc(s.name)} כתב/ה: ${esc(s.reason)}</p>`).join('')}`:'<span class="chip">קישור לצפייה</span>')+
      `<div class="lnk">${m.mode==='sign'?`<button class="s-btn sm ${done===all?'pri':''}" data-dl="${r.id}" ${done?'':'disabled'}>${done===all?'הורדת המסמך החתום':'הורדה עם החתימות עד עכשיו'}</button>`:''}
       <a class="s-btn sm" href="${ownerLink(r.id,r.k)}" target="_blank" rel="noopener">צפייה במסמך</a><button class="s-btn sm" data-links="${r.id}">הקישורים</button><button class="s-btn sm" data-rm="${r.id}">מחיקה מהשרת</button></div>`;};
  l.forEach(draw);
  u.body.addEventListener('click',async e=>{
    const el=e.target.closest('[data-dl],[data-links],[data-rm],[data-forget]');if(!el||!u.body.contains(el))return;const id=el.dataset;const r=l.find(x=>x.id===(id.dl||id.links||id.rm||id.forget));if(!r)return;
    if(id.links)return links(r,o);
    if(id.forget){listSet(listGet().filter(x=>x.id!==r.id));return list(o);}
    if(id.rm){if(!confirm(`למחוק את "${r.title}" מהשרת? הקישורים יפסיקו לעבוד.`))return;try{await api(`/${r.id}?k=${encodeURIComponent(r.k)}`,{method:'DELETE'});}catch(err){if(!/לא נמצא/.test(err.message)){say(err.message,1);return;}}
      listSet(listGet().filter(x=>x.id!==r.id));return list(o);}
    if(id.dl){busy(true,'בונה את המסמך החתום…');try{const pdf=await getPdf(r.id,{k:r.k});const out=await buildSigned(r.meta,pdf);busy(false);
      const name=`${r.title} - חתום.pdf`;if(o.save)await o.save(name,out,'pdf');else{const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([out],{type:'application/pdf'}));a.download=name;a.click();}}
      catch(err){busy(false);say('לא הצלחתי לבנות את המסמך: '+(err.message||err),1);}}
  });
}

/* ---------- a link for viewing only */
async function openView(opts){
  const u=shell('קישור לשיתוף');u.opts=opts;
  u.body.innerHTML='<div class="s-wide"></div>';
  await upload(u,{mode:'view',title:String(opts.name||'מסמך').replace(/\.pdf$/i,''),signers:[],fields:[],pages:1});
}

window.SIMAN_SEND={open,openView,list,links,close,buildSigned,getMeta,getPdf,api,signLink,viewLink,ownerLink,HOST,devName,COLORS,TYPES,isOpen:()=>!!U};
})();
