/* סימן — תרשימים בקובצי Excel.
   קורא את התרשימים שבקובץ (עמודות, פסים, קו, שטח, עוגה, טבעת ופיזור), מצייר אותם כמו ב-Excel
   מתוך הנתונים שבתאים, וכותב תרשימים חדשים או ערוכים בפורמט ש-Excel פותח.
   Used by excel-editor.js (the editor) and office-edit.js (a sheet opened for marking / saved as PDF). */
(function(){
'use strict';
const NC='http://schemas.openxmlformats.org/drawingml/2006/chart',NA='http://schemas.openxmlformats.org/drawingml/2006/main',
  NR='http://schemas.openxmlformats.org/officeDocument/2006/relationships',NX='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  NREL='http://schemas.openxmlformats.org/package/2006/relationships';
const RT_CHART='http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const RT_DRAW='http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
const CT_CHART='application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
const CT_DRAW='application/vnd.openxmlformats-officedocument.drawing+xml';
const EMU=9525;
const DEF_ACCENT=['4472C4','ED7D31','A5A5A5','FFC000','5B9BD5','70AD47'];
const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'})[c]);
const xml=s=>{const d=new DOMParser().parseFromString(s,'application/xml');return d.getElementsByTagName('parsererror').length?null:d;};
const ch=(el,ns,name)=>el?[...el.childNodes].find(n=>n.nodeType===1&&n.namespaceURI===ns&&n.localName===name)||null:null;
const chs=(el,ns,name)=>el?[...el.childNodes].filter(n=>n.nodeType===1&&n.namespaceURI===ns&&(!name||n.localName===name)):[];
const c1=(el,name)=>ch(el,NC,name);
const val=(el,name,d)=>{const e=c1(el,name);return e&&e.getAttribute('val')!==null?e.getAttribute('val'):d;};
const on=(el,name,d)=>{const v=val(el,name,null);return v===null?d:!/^(0|false)$/.test(v);};

/* ---------------- paths and relationships (works on any reader: get(path) -> text|null) */
function relsPath(part){const i=part.lastIndexOf('/');return part.slice(0,i+1)+'_rels/'+part.slice(i+1)+'.rels';}
function resolve(part,t){if(/^\//.test(t))return t.slice(1);const b=part.slice(0,part.lastIndexOf('/')+1).split('/').filter(Boolean);for(const s of t.split('/')){if(s==='..')b.pop();else if(s&&s!=='.')b.push(s);}return b.join('/');}
function rels(get,part){const s=get(relsPath(part));const d=s&&xml(s);if(!d)return[];return[...d.getElementsByTagNameNS(NREL,'Relationship')].map(e=>({id:e.getAttribute('Id'),type:e.getAttribute('Type'),target:e.getAttribute('Target')}));}

/* ---------------- colours */
function accents(get){
  const out=DEF_ACCENT.slice();const s=get('xl/theme/theme1.xml');const d=s&&xml(s);if(!d)return out;
  for(let i=1;i<=6;i++){const e=d.getElementsByTagNameNS(NA,'accent'+i)[0];if(!e)continue;const c=e.getElementsByTagNameNS(NA,'srgbClr')[0]||e.getElementsByTagNameNS(NA,'sysClr')[0];
    const v=c&&(c.getAttribute('val')&&c.localName==='srgbClr'?c.getAttribute('val'):c.getAttribute('lastClr'));if(v&&/^[0-9a-f]{6}$/i.test(v))out[i-1]=v.toUpperCase();}
  return out;
}
const hx=h=>[0,2,4].map(i=>parseInt(h.slice(i,i+2),16));
const toHex=a=>a.map(x=>Math.max(0,Math.min(255,Math.round(x))).toString(16).padStart(2,'0')).join('').toUpperCase();
function lum(h,mod,off){const [r,g,b]=hx(h).map(x=>x/255);const mx=Math.max(r,g,b),mn=Math.min(r,g,b);let H=0,S=0,L=(mx+mn)/2;const d=mx-mn;
  if(d){S=L>.5?d/(2-mx-mn):d/(mx+mn);H=mx===r?(g-b)/d+(g<b?6:0):mx===g?(b-r)/d+2:(r-g)/d+4;H/=6;}
  L=Math.max(0,Math.min(1,L*mod+off));const q=L<.5?L*(1+S):L+S-L*S,p=2*L-q;
  const f=t=>{t<0&&(t+=1);t>1&&(t-=1);return t<1/6?p+(q-p)*6*t:t<.5?q:t<2/3?p+(q-p)*(2/3-t)*6:p;};
  return S?toHex([f(H+1/3),f(H),f(H-1/3)].map(x=>x*255)):toHex([L*255,L*255,L*255]);}
function colorIn(el,acc){
  if(!el)return null;const sf=chs(el,NA,'solidFill')[0];if(!sf)return null;
  const c=chs(sf,NA)[0];if(!c)return null;let h=null;
  if(c.localName==='srgbClr')h=c.getAttribute('val');
  else if(c.localName==='schemeClr'){const v=c.getAttribute('val');const m=/^accent(\d)$/.exec(v);h=m?acc[+m[1]-1]:v==='tx1'||v==='dk1'?'000000':v==='bg1'||v==='lt1'?'FFFFFF':null;}
  else if(c.localName==='sysClr')h=c.getAttribute('lastClr');
  if(!h)return null;
  const mod=chs(c,NA,'lumMod')[0],off=chs(c,NA,'lumOff')[0],sh=chs(c,NA,'shade')[0],ti=chs(c,NA,'tint')[0];
  if(mod||off)h=lum(h,mod?+mod.getAttribute('val')/1e5:1,off?+off.getAttribute('val')/1e5:0);
  if(sh)h=toHex(hx(h).map(x=>x*(+sh.getAttribute('val')/1e5)));
  if(ti)h=toHex(hx(h).map(x=>x+(255-x)*(1-(+ti.getAttribute('val')/1e5))));
  return h.toUpperCase();
}
// the colour Excel gives series number i when the file does not say
function autoColor(i,acc){const base=acc[i%6];const round=Math.floor(i/6);if(!round)return base;return lum(base,round%2?.6:.8,round%2?0:.2);}

/* ---------------- reading a chart part into a model */
const KIND={barChart:'bar',bar3DChart:'bar',lineChart:'line',line3DChart:'line',stockChart:'line',radarChart:'line',areaChart:'area',area3DChart:'area',
  pieChart:'pie',pie3DChart:'pie',ofPieChart:'pie',doughnutChart:'doughnut',scatterChart:'scatter',bubbleChart:'scatter'};
function richText(el){if(!el)return'';return[...el.getElementsByTagNameNS(NA,'p')].map(p=>[...p.getElementsByTagNameNS(NA,'t')].map(t=>t.textContent).join('')).join(' ').trim();}
function dataRef(el){
  // <c:strRef>/<c:numRef> with a formula and a cache, or literal values
  if(!el)return null;const ref=c1(el,'numRef')||c1(el,'strRef')||c1(el,'multiLvlStrRef');const lit=c1(el,'numLit')||c1(el,'strLit');
  const cacheEl=ref?(c1(ref,'numCache')||c1(ref,'strCache')||c1(ref,'multiLvlStrCache')):lit;
  const cache=[];let fmt=null;
  if(cacheEl){const n=+val(cacheEl,'ptCount',0)||0;fmt=(c1(cacheEl,'formatCode')||{}).textContent||null;
    const lvl=c1(cacheEl,'lvl')||cacheEl;for(const p of chs(lvl,NC,'pt')){const i=+p.getAttribute('idx');const v=(c1(p,'v')||{}).textContent;if(i>=0&&i<Math.max(n,1e5))cache[i]=v;}
    for(let i=0;i<n;i++)if(cache[i]===undefined)cache[i]=null;}
  return{f:ref?(c1(ref,'f')||{}).textContent||null:null,cache,fmt,num:!!(ref&&ref.localName==='numRef'||lit&&lit.localName==='numLit')};
}
function parse(text,acc){
  const d=xml(text);if(!d)return null;const cs=d.documentElement,chart=c1(cs,'chart');if(!chart)return null;
  const pa=c1(chart,'plotArea');if(!pa)return null;
  const groups=chs(pa,NC).filter(e=>KIND[e.localName]);if(!groups.length)return null;
  const m={type:KIND[groups[0].localName],dir:'col',grouping:'clustered',series:[],title:null,legend:null,labels:false,pct:false,catName:false,catRev:false,valPos:'l',
    valFmt:null,grid:false,catHide:false,valHide:false,markers:false,smooth:false,hole:50,gap:150,overlap:0,vary:false,min:null,max:null};
  const g0=groups[0];
  if(m.type==='bar'){m.dir=val(g0,'barDir','col');m.grouping=val(g0,'grouping','clustered');m.gap=+val(g0,'gapWidth',150);m.overlap=+val(g0,'overlap',0);}
  else if(m.type==='line'||m.type==='area')m.grouping=val(g0,'grouping','standard');
  if(m.type==='doughnut')m.hole=+val(g0,'holeSize',50);
  m.vary=on(g0,'varyColors',false);
  let k=0;
  for(const g of groups){const kind=KIND[g.localName];const gl=c1(g,'dLbls');
    const gLbl=gl?{v:on(gl,'showVal',false),p:on(gl,'showPercent',false),c:on(gl,'showCatName',false)}:null;
    const gMark=g.localName==='lineChart'?on(g,'marker',true):false;
    for(const s of chs(g,NC,'ser')){
      const idx=+val(s,'idx',k);const sp=c1(s,'spPr');
      const line=sp&&chs(sp,NA,'ln')[0];
      const color=kind==='line'||kind==='scatter'?(colorIn(line,acc)||colorIn(sp,acc)):(colorIn(sp,acc)||colorIn(line,acc));
      const tx=c1(s,'tx');let name=null,nameRef=null;
      if(tx){const r=c1(tx,'strRef');if(r){nameRef=(c1(r,'f')||{}).textContent||null;const c=c1(r,'strCache');const p=c&&c1(c,'pt');name=p?(c1(p,'v')||{}).textContent:null;}else name=(c1(tx,'v')||{}).textContent||null;}
      const sl=c1(s,'dLbls');const lbl=sl&&!on(sl,'delete',false)?{v:on(sl,'showVal',false),p:on(sl,'showPercent',false),c:on(sl,'showCatName',false)}:(sl?{v:false,p:false,c:false}:gLbl);
      const mk=c1(s,'marker');const sym=mk?val(mk,'symbol',null):null;
      const pts={};for(const dp of chs(s,NC,'dPt')){const i=+val(dp,'idx',-1);const c=colorIn(c1(dp,'spPr'),acc);if(i>=0&&c)pts[i]=c;}
      m.series.push({kind,idx,name,nameRef,color:color||null,cat:dataRef(c1(s,'cat')||c1(s,'xVal')),val:dataRef(c1(s,'val')||c1(s,'yVal')),
        labels:lbl?lbl.v:false,pct:lbl?lbl.p:false,catName:lbl?lbl.c:false,marker:kind==='line'||kind==='scatter'?(sym?sym!=='none':gMark):false,
        smooth:on(s,'smooth',false),pts,dir:kind==='bar'?val(g,'barDir','col'):null});
      k++;
    }
  }
  m.series.forEach((s,i)=>{if(!s.color)s.color=autoColor(s.idx>=0?s.idx:i,acc);});
  m.labels=m.series.some(s=>s.labels);m.pct=m.series.some(s=>s.pct);m.catName=m.series.some(s=>s.catName);m.markers=m.series.some(s=>s.marker);
  // title: written text, or the series name when the title is "automatic"
  const t=c1(chart,'title');
  if(t){const tx=c1(t,'tx');let txt=tx?richText(c1(tx,'rich')):'';if(!txt&&tx){const r=c1(tx,'strRef');const c=r&&c1(r,'strCache');const p=c&&c1(c,'pt');txt=p?(c1(p,'v')||{}).textContent:'';}
    m.title=txt||(m.series.length===1?m.series[0].name||'':(m.type==='pie'||m.type==='doughnut')&&m.series[0]?m.series[0].name||'':'כותרת תרשים');m.titleAuto=!txt;}
  else if(!on(chart,'autoTitleDeleted',false)&&m.series.length===1&&(m.type==='pie'||m.type==='doughnut'))m.title=m.series[0].name||null;
  const lg=c1(chart,'legend');if(lg)m.legend=val(lg,'legendPos','r');
  const cat=c1(pa,'catAx')||c1(pa,'dateAx')||(m.type==='scatter'?chs(pa,NC,'valAx')[0]:null);
  const va=m.type==='scatter'?chs(pa,NC,'valAx')[1]:c1(pa,'valAx');
  if(cat){const sc=c1(cat,'scaling');m.catRev=val(sc,'orientation','minMax')==='maxMin';m.catHide=on(cat,'delete',false);const nf=c1(cat,'numFmt');m.catFmt=nf?nf.getAttribute('formatCode'):null;m.catLinked=nf?nf.getAttribute('sourceLinked')!=='0':true;}
  if(va){const sc=c1(va,'scaling');m.valHide=on(va,'delete',false);m.valPos=val(va,'axPos','l');m.grid=!!c1(va,'majorGridlines');m.valRev=val(sc,'orientation','minMax')==='maxMin';
    const mn=c1(sc,'min'),mx=c1(sc,'max');if(mn)m.min=+mn.getAttribute('val');if(mx)m.max=+mx.getAttribute('val');const nf=c1(va,'numFmt');m.valFmt=nf?nf.getAttribute('formatCode'):null;}
  else m.grid=m.type!=='pie'&&m.type!=='doughnut';
  return m;
}

/* the charts on one sheet: [{part, rid, anchor:{from,to} or {from, ext} or {pos, ext}, model}] plus the drawing they sit in */
function readSheet(get,sheetPath,sheetXmlText,acc){
  const out={drawPath:null,drawRid:null,charts:[]};
  const sd=xml(sheetXmlText||get(sheetPath)||'');if(!sd)return out;
  const de=[...sd.documentElement.childNodes].find(n=>n.nodeType===1&&n.localName==='drawing');if(!de)return out;
  const rid=de.getAttributeNS(NR,'id');const rl=rels(get,sheetPath).find(r=>r.id===rid&&r.type===RT_DRAW);if(!rl)return out;
  const dp=resolve(sheetPath,rl.target);const ds=get(dp);const dd=ds&&xml(ds);if(!dd)return out;
  out.drawPath=dp;out.drawRid=rid;out.drawDoc=dd;
  const drl=rels(get,dp);
  const pos=el=>{if(!el)return null;const g=n=>{const e=chs(el,NX,n)[0];return e?+e.textContent||0:0;};return{c:g('col'),co:g('colOff')/EMU,r:g('row'),ro:g('rowOff')/EMU};};
  chs(dd.documentElement,NX).forEach((an,ai)=>{
    if(!/Anchor$/.test(an.localName))return;
    const gf=an.getElementsByTagNameNS(NX,'graphicFrame')[0];if(!gf)return;
    // a frame inside a "newer charts" block (waterfall, funnel…) is skipped: only the classic chart is drawn
    if(gf.parentNode!==an)return;
    const cEl=gf.getElementsByTagNameNS(NC,'chart')[0];if(!cEl)return;
    const cid=cEl.getAttributeNS(NR,'id');const cr=drl.find(r=>r.id===cid&&r.type===RT_CHART);if(!cr)return;
    const part=resolve(dp,cr.target);const txt=get(part);if(!txt)return;
    let model=null;try{model=parse(txt,acc);}catch(e){console.warn('chart',part,e);}
    if(!model)return;
    const ext=chs(an,NX,'ext')[0];const ps=chs(an,NX,'pos')[0];
    out.charts.push({part,rid:cid,ai,el:an,model,from:pos(chs(an,NX,'from')[0]),to:pos(chs(an,NX,'to')[0]),
      ext:ext?{w:(+ext.getAttribute('cx')||0)/EMU,h:(+ext.getAttribute('cy')||0)/EMU}:null,abs:ps?{x:(+ps.getAttribute('x')||0)/EMU,y:(+ps.getAttribute('y')||0)/EMU}:null,
      name:(gf.getElementsByTagNameNS(NX,'cNvPr')[0]||{getAttribute:()=>''}).getAttribute('name')||''});
  });
  return out;
}

/* ---------------- numbers on the chart */
function general(v){if(!isFinite(v))return'';if(Number.isInteger(v))return String(v);const a=Math.abs(v);if(a>=1e11||a<1e-9&&a>0)return v.toExponential(4).replace(/\.?0+e/,'E').replace('e','E');return String(+v.toPrecision(10)).replace(/(\.\d{0,})$/,m=>m);}
function fmtNum(v,code){
  if(v===null||v===undefined||v==='')return'';if(typeof v!=='number'){const n=+v;if(v===''||isNaN(n))return String(v);v=n;}
  if(!code||/^general$/i.test(code))return general(+v.toPrecision(12));
  // Excel's built-in short date: written as m/d/yyyy, shown the local way (day first in Hebrew)
  if(/^m+\/d+\/y+$/i.test(code))code='dd/mm/yyyy';
  try{if(window.XLSX&&XLSX.SSF)return XLSX.SSF.format(code,v);}catch(_){}
  return general(v);
}
// a reference "'Sheet name'!$A$1:$A$9" → {sheet, r1,c1,r2,c2}
const colIdx=s=>{let n=0;for(const ch of s.toUpperCase())n=n*26+(ch.charCodeAt(0)-64);return n-1;};
function splitRef(f){
  if(!f)return null;f=String(f).trim().replace(/^\(|\)$/g,'');if(f.includes(','))return null;
  const m=/^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(f);if(!m)return null;
  const sheet=m[1]!==undefined?m[1].replace(/''/g,"'"):m[2];const r1=+m[4]-1,c1=colIdx(m[3]),r2=m[6]?+m[6]-1:r1,c2=m[5]?colIdx(m[5]):c1;
  return{sheet,r1:Math.min(r1,r2),c1:Math.min(c1,c2),r2:Math.max(r1,r2),c2:Math.max(c1,c2)};
}
const colName=c=>{let s='';c++;while(c>0){const m=(c-1)%26;s=String.fromCharCode(65+m)+s;c=Math.floor((c-1)/26);}return s;};
function refText(sheet,r1,c1,r2,c2){const nm=/^[A-Za-z_][\w.]*$/.test(sheet)?sheet:`'${String(sheet).replace(/'/g,"''")}'`;const a=`$${colName(c1)}$${r1+1}`;return nm+'!'+a+(r2!==r1||c2!==c1?`:$${colName(c2)}$${r2+1}`:'');}

/* the numbers and names a model shows: live values from the cells when a reader is given (cells(ref) -> [{v,w}]), else the file's cache */
function data(m,cells){
  const get=(r,isNum,fmt)=>{
    if(!r)return null;const live=r.f&&cells?cells(r.f):null;
    if(live)return live.map(p=>({v:p?p.v:null,w:p?(p.w!==undefined&&p.w!==null?String(p.w):p.v===null||p.v===undefined?'':typeof p.v==='number'?fmtNum(p.v,fmt):String(p.v)):''}));
    return r.cache.map(v=>{if(v===null||v===undefined)return{v:null,w:''};const n=+v;return r.num&&v!==''&&!isNaN(n)?{v:n,w:fmtNum(n,r.fmt||fmt)}:{v,w:String(v)};});
  };
  const series=m.series.map(s=>{
    const vals=(get(s.val,true,m.valFmt)||[]).map(p=>typeof p.v==='number'&&isFinite(p.v)?p.v:(p.v!==null&&p.v!==''&&!isNaN(+p.v)?+p.v:null));
    const cats=get(s.cat,false,m.catFmt);
    let name=s.name;if(s.nameRef&&cells){const n=cells(s.nameRef);if(n&&n[0])name=n[0].w||String(n[0].v??'');}
    return{...s,name:name||'',vals,cats:cats?cats.map(p=>p.w):null,xs:cats?cats.map(p=>typeof p.v==='number'?p.v:null):null};
  });
  return series;
}

/* ---------------- drawing (SVG) */
const FONT="Calibri, 'Segoe UI', Arial, sans-serif";
const tw=(s,fs)=>{let w=0;for(const ch of String(s))w+=/[֐-׿]/.test(ch)?.56:/[0-9]/.test(ch)?.52:/[.,:/ ]/.test(ch)?.28:/[A-Z]/.test(ch)?.62:.5;return w*fs;};
function nice(lo,hi,count){
  if(!isFinite(lo)||!isFinite(hi)){lo=0;hi=1;}
  if(lo===hi){if(hi===0)hi=1;else if(hi>0)lo=0;else hi=0;}
  const span=hi-lo,raw=span/Math.max(1,count),mag=10**Math.floor(Math.log10(raw)),f=raw/mag;
  const step=(f<=1?1:f<=2?2:f<=2.5?2.5:f<=5?5:10)*mag;
  return{lo:Math.floor(lo/step+1e-9)*step,hi:Math.ceil(hi/step-1e-9)*step,step};
}
const T=(x,y,s,o={})=>`<text x="${x.toFixed(1)}" y="${y.toFixed(1)}" font-size="${o.fs||12}" fill="#${o.c||'595959'}"${o.a?` text-anchor="${o.a}"`:''}${o.b?' font-weight="bold"':''}${o.rot?` transform="rotate(${o.rot} ${x.toFixed(1)} ${y.toFixed(1)})"`:''}>${esc(s)}</text>`;
function svg(m,W,H,cells,opt={}){
  W=Math.max(60,W);H=Math.max(40,H);
  const S=data(m,cells);const pie=m.type==='pie'||m.type==='doughnut';
  let o=`<svg xmlns="http://www.w3.org/2000/svg" width="${opt.pxW||W}" height="${opt.pxH||H}" viewBox="0 0 ${W} ${H}" font-family="${FONT}" direction="ltr" style="direction:ltr">`;
  o+=`<rect x="0.5" y="0.5" width="${W-1}" height="${H-1}" fill="#fff" stroke="#D9D9D9"/>`;
  let top=8,bottom=H-8,left=10,right=W-10;
  if(m.title){const fs=Math.min(18.6,Math.max(11,H/14));o+=T(W/2,top+fs,m.title.length>80?m.title.slice(0,78)+'…':m.title,{fs,a:'middle'});top+=fs+12;}
  // legend
  const legItems=pie?(S[0]&&S[0].cats?S[0].cats.map((c,i)=>({n:c,c:S[0].pts[i]||(m.vary||pie?autoColor(i,opt.acc||DEF_ACCENT):S[0].color)})):[]):S.map(s=>({n:s.name||'',c:s.color,line:s.kind==='line'||s.kind==='scatter'}));
  if(m.legend&&legItems.length){
    const fs=12;const lp=m.legend;
    if(lp==='b'||lp==='t'){
      const rows=[];let row=[],w=0;for(const it of legItems){const iw=tw(it.n,fs)+26;if(row.length&&w+iw>right-left){rows.push(row);row=[];w=0;}row.push(it);w+=iw;}if(row.length)rows.push(row);
      const lh=rows.length*(fs+6);let y=lp==='b'?bottom-lh+fs:top+fs;
      for(const r of rows){let x=W/2-r.reduce((a,it)=>a+tw(it.n,fs)+26,0)/2;
        for(const it of r){o+=it.line?`<line x1="${x}" y1="${y-4}" x2="${x+16}" y2="${y-4}" stroke="#${it.c}" stroke-width="2.5"/>`:`<rect x="${x+3}" y="${y-9}" width="9" height="9" fill="#${it.c}"/>`;
          o+=T(x+20,y,it.n,{fs});x+=tw(it.n,fs)+26;}y+=fs+6;}
      if(lp==='b')bottom-=lh+6;else top+=lh+6;
    }else{
      const lw=Math.min(W*.35,Math.max(...legItems.map(it=>tw(it.n,fs)))+30);const lh=legItems.length*(fs+6);let y=(top+bottom)/2-lh/2+fs;
      const x0=lp==='l'?left:right-lw;
      for(const it of legItems.slice(0,Math.floor((bottom-top)/(fs+6)))){o+=it.line?`<line x1="${x0}" y1="${y-4}" x2="${x0+16}" y2="${y-4}" stroke="#${it.c}" stroke-width="2.5"/>`:`<rect x="${x0+3}" y="${y-9}" width="9" height="9" fill="#${it.c}"/>`;
        o+=T(x0+20,y,it.n.length>40?it.n.slice(0,38)+'…':it.n,{fs});y+=fs+6;}
      if(lp==='l')left+=lw+6;else right-=lw+6;
    }
  }
  if(!S.length||!S.some(s=>s.vals.some(v=>v!==null))){o+=T(W/2,(top+bottom)/2,'אין נתונים לתרשים',{fs:12,a:'middle',c:'7F7F7F'});return o+'</svg>';}
  if(pie)o+=pieSvg(m,S[0],left,top,right,bottom,opt);
  else o+=axesSvg(m,S,left,top,right,bottom);
  return o+'</svg>';
}
function pieSvg(m,s,L,Tp,R,B,opt){
  const acc=opt.acc||DEF_ACCENT;const vals=s.vals.map(v=>v&&v>0?v:0);const sum=vals.reduce((a,b)=>a+b,0);if(!sum)return'';
  const cx=(L+R)/2,cy=(Tp+B)/2,r=Math.max(10,Math.min(R-L,B-Tp)/2-6),ri=m.type==='doughnut'?r*Math.max(.1,Math.min(.9,(m.hole||50)/100)):0;
  let a=-Math.PI/2,o='';const lbl=[];
  vals.forEach((v,i)=>{if(!v)return;const da=v/sum*Math.PI*2,b=a+da,col=s.pts[i]||autoColor(i,acc);
    const P=(ang,rr)=>`${(cx+rr*Math.cos(ang)).toFixed(2)} ${(cy+rr*Math.sin(ang)).toFixed(2)}`;const big=da>Math.PI?1:0;
    if(da>=Math.PI*2-1e-6)o+=ri?`<circle cx="${cx}" cy="${cy}" r="${(r+ri)/2}" fill="none" stroke="#${col}" stroke-width="${r-ri}"/>`:`<circle cx="${cx}" cy="${cy}" r="${r}" fill="#${col}"/>`;
    else o+=ri?`<path d="M${P(a,r)} A${r} ${r} 0 ${big} 1 ${P(b,r)} L${P(b,ri)} A${ri} ${ri} 0 ${big} 0 ${P(a,ri)} Z" fill="#${col}" stroke="#fff" stroke-width="1"/>`
      :`<path d="M${cx} ${cy} L${P(a,r)} A${r} ${r} 0 ${big} 1 ${P(b,r)} Z" fill="#${col}" stroke="#fff" stroke-width="1"/>`;
    if(m.labels||m.pct||m.catName){const mid=(a+b)/2,rr=ri?(r+ri)/2:r*.65;const parts=[];if(m.catName&&s.cats)parts.push(s.cats[i]);if(m.labels)parts.push(fmtNum(v,m.valFmt));if(m.pct)parts.push(Math.round(v/sum*100)+'%');
      lbl.push(T(cx+rr*Math.cos(mid),cy+rr*Math.sin(mid)+4,parts.join('; '),{fs:11,a:'middle',c:'FFFFFF',b:true}));}
    a=b;});
  return o+lbl.join('');
}
function axesSvg(m,S,L,Tp,R,B){
  const horiz=m.type==='bar'&&m.dir==='bar';
  const n=Math.max(...S.map(s=>Math.max(s.vals.length,s.cats?s.cats.length:0)));
  const cats=((S.find(s=>s.cats&&s.cats.length)||{}).cats||Array.from({length:n},(_,i)=>String(i+1))).map(c=>{c=String(c??'');return c.length>18?c.slice(0,17)+'…':c;});
  const stacked=/stacked/i.test(m.grouping),pct=m.grouping==='percentStacked';
  // value range
  let lo=Infinity,hi=-Infinity;
  if(stacked){for(let i=0;i<n;i++){let p=0,q=0,t=0;for(const s of S){const v=s.vals[i];if(v==null)continue;t+=Math.abs(v);if(v>=0)p+=v;else q+=v;}
    if(pct){p=t?p/t*100:0;q=t?q/t*100:0;}lo=Math.min(lo,q);hi=Math.max(hi,p);}}
  else for(const s of S)for(const v of s.vals)if(v!=null){lo=Math.min(lo,v);hi=Math.max(hi,v);}
  if(m.type!=='line'&&m.type!=='scatter'||lo>=0&&lo<hi*.84)lo=Math.min(0,lo);if(hi<0)hi=0;
  if(m.type==='line'&&lo>0&&lo>=hi*.84){} // Excel keeps a floating minimum when the values are close together
  const fsA=11;
  const span=horiz?(R-L):(B-Tp);
  let sc=nice(m.min!=null?m.min:lo,m.max!=null?m.max:hi,Math.max(2,Math.min(10,Math.floor(span/(horiz?70:34)))));
  if(m.min!=null)sc.lo=m.min;if(m.max!=null)sc.hi=m.max;
  const ticks=[];for(let v=sc.lo;v<=sc.hi+sc.step*1e-6;v+=sc.step)ticks.push(+v.toPrecision(12));
  const vf=v=>pct?Math.round(v)+'%':fmtNum(v,m.valFmt);
  // room for the labels
  const vw=m.valHide?0:Math.max(...ticks.map(t=>tw(vf(t),fsA)))+8;
  let x0=L,x1=R,y0=Tp+4,y1=B;
  const catW0=Math.max(...cats.map(c=>tw(c,fsA)));
  if(!horiz){
    if(!m.valHide){if(m.valPos==='r')x1-=vw;else x0+=vw;}
    const cw=(x1-x0)/Math.max(1,n);
    let rot=0,step=1,ch=m.catHide?0:fsA+8;
    if(!m.catHide&&catW0>cw-4){rot=-45;ch=Math.min((y1-y0)*.4,catW0*.72+fsA);step=Math.max(1,Math.ceil((fsA*1.25)/cw));}
    y1-=ch;
    var cat={cw,rot,step};
  }else{
    const ch=(y1-y0)/Math.max(1,n);
    if(!m.catHide){const w=Math.min((x1-x0)*.4,catW0+8);if(m.valPos==='r'||m.catRev)x1-=w;else x0+=w;}
    if(!m.valHide)y1-=fsA+8;
    var cat={ch,step:Math.max(1,Math.ceil((fsA+2)/ch))};
  }
  let o='';
  const vpos=v=>horiz?x0+(v-sc.lo)/(sc.hi-sc.lo||1)*(x1-x0):y1-(v-sc.lo)/(sc.hi-sc.lo||1)*(y1-y0);
  // grid and value labels
  for(const t of ticks){const p=vpos(t);
    if(m.grid)o+=horiz?`<line x1="${p.toFixed(1)}" y1="${y0}" x2="${p.toFixed(1)}" y2="${y1}" stroke="#D9D9D9" stroke-width="1"/>`:`<line x1="${x0}" y1="${p.toFixed(1)}" x2="${x1}" y2="${p.toFixed(1)}" stroke="#D9D9D9" stroke-width="1"/>`;
    if(!m.valHide)o+=horiz?T(p,y1+fsA+4,vf(t),{fs:fsA,a:'middle'}):(m.valPos==='r'?T(x1+5,p+4,vf(t),{fs:fsA}):T(x0-5,p+4,vf(t),{fs:fsA,a:'end'}));}
  const base=vpos(Math.max(sc.lo,Math.min(sc.hi,0)));
  o+=horiz?`<line x1="${base.toFixed(1)}" y1="${y0}" x2="${base.toFixed(1)}" y2="${y1}" stroke="#BFBFBF"/>`:`<line x1="${x0}" y1="${base.toFixed(1)}" x2="${x1}" y2="${base.toFixed(1)}" stroke="#BFBFBF"/>`;
  // category positions: horizontal charts run bottom-up like in Excel; a reversed axis turns them around
  const ci=i=>horiz?(m.catRev?i:n-1-i):(m.catRev?n-1-i:i);
  const cpos=i=>horiz?y0+(ci(i)+.5)*cat.ch:x0+(ci(i)+.5)*cat.cw;
  if(!m.catHide)for(let i=0;i<n;i+=cat.step){const c=cats[i]??'';
    if(horiz)o+=m.valPos==='r'||m.catRev?T(x1+5,cpos(i)+4,c,{fs:fsA}):T(x0-5,cpos(i)+4,c,{fs:fsA,a:'end'});
    else o+=cat.rot?T(cpos(i)+3,y1+8,c,{fs:fsA,a:'end',rot:cat.rot}):T(cpos(i),y1+fsA+4,c,{fs:fsA,a:'middle'});}
  // bars
  const bars=S.filter(s=>s.kind==='bar');
  const lbls=[];
  if(bars.length){
    const slot=horiz?cat.ch:cat.cw,grp=slot/(1+(m.gap||150)/100);const nb=stacked?1:bars.length;const bw=grp/nb;
    for(let i=0;i<n;i++){let p=0,q=0,t=0;if(pct)for(const s of bars){const v=s.vals[i];if(v!=null)t+=Math.abs(v);}
      bars.forEach((s,k)=>{let v=s.vals[i];if(v==null)return;if(pct)v=t?v/t*100:0;
        let a,b;if(stacked){if(v>=0){a=p;b=p+v;p=b;}else{a=q;b=q+v;q=b;}}else{a=0;b=v;}
        const off=stacked?-grp/2:-grp/2+k*bw;const c=cpos(i)+off;const pa=vpos(Math.max(sc.lo,Math.min(sc.hi,a))),pb=vpos(Math.max(sc.lo,Math.min(sc.hi,b)));
        const col=s.pts[i]||(m.vary&&bars.length===1?autoColor(i,DEF_ACCENT):s.color);
        if(horiz)o+=`<rect x="${Math.min(pa,pb).toFixed(1)}" y="${c.toFixed(1)}" width="${Math.abs(pb-pa).toFixed(1)}" height="${bw.toFixed(1)}" fill="#${col}"/>`;
        else o+=`<rect x="${c.toFixed(1)}" y="${Math.min(pa,pb).toFixed(1)}" width="${bw.toFixed(1)}" height="${Math.abs(pb-pa).toFixed(1)}" fill="#${col}"/>`;
        if(s.labels)lbls.push(horiz?T(Math.max(pa,pb)+4,c+bw/2+4,fmtNum(s.vals[i],m.valFmt),{fs:10,c:'404040'}):T(c+bw/2,Math.min(pa,pb)-4,fmtNum(s.vals[i],m.valFmt),{fs:10,a:'middle',c:'404040'}));
      });}
  }
  // areas, then lines on top
  const stackBase=new Array(n).fill(0);
  for(const s of S.filter(s=>s.kind==='area')){
    const pts=[];for(let i=0;i<n;i++){const v=s.vals[i]??0;const b=stacked?stackBase[i]:0;const t=b+v;pts.push([cpos(i),vpos(t),vpos(b)]);if(stacked)stackBase[i]=t;}
    if(!pts.length)continue;
    const d='M'+pts.map(p=>p[0].toFixed(1)+' '+p[1].toFixed(1)).join(' L')+' L'+pts.slice().reverse().map(p=>p[0].toFixed(1)+' '+p[2].toFixed(1)).join(' L')+' Z';
    o+=`<path d="${d}" fill="#${s.color}" fill-opacity="${stacked?1:.85}"/>`;
    if(s.labels)s.vals.forEach((v,i)=>{if(v!=null)lbls.push(T(cpos(i),pts[i][1]-5,fmtNum(v,m.valFmt),{fs:10,a:'middle',c:'404040'}));});
  }
  const lineBase=new Array(n).fill(0);
  for(const s of S.filter(s=>s.kind==='line'||s.kind==='scatter')){
    let d='',pen=false;const pts=[];
    for(let i=0;i<n;i++){let v=s.vals[i];if(v==null){pen=false;continue;}if(stacked){v+=lineBase[i];lineBase[i]=v;}
      const x=cpos(i),y=vpos(Math.max(sc.lo,Math.min(sc.hi,v)));d+=(pen?' L':' M')+x.toFixed(1)+' '+y.toFixed(1);pen=true;pts.push([x,y,s.vals[i]]);}
    if(s.kind!=='scatter'||s.marker===false||true)o+=d?`<path d="${d.trim()}" fill="none" stroke="#${s.color}" stroke-width="2.25" stroke-linejoin="round" stroke-linecap="round"/>`:'';
    if(s.marker)for(const p of pts)o+=`<circle cx="${p[0].toFixed(1)}" cy="${p[1].toFixed(1)}" r="3" fill="#${s.color}" stroke="#fff" stroke-width=".8"/>`;
    if(s.labels)for(const p of pts)lbls.push(T(p[0],p[1]-6,fmtNum(p[2],m.valFmt),{fs:10,a:'middle',c:'404040'}));
  }
  return o+lbls.join('');
}

/* ---------------- writing a chart part (new charts, or a chart that was changed in Siman) */
function chartXml(m,cells,opt={}){
  const S=data(m,cells);const pie=m.type==='pie'||m.type==='doughnut';
  const tag={bar:'barChart',line:'lineChart',area:'areaChart',pie:'pieChart',doughnut:'doughnutChart',scatter:'lineChart'}[m.type]||'barChart';
  const grouping=m.type==='bar'?(m.grouping||'clustered'):m.grouping&&m.grouping!=='clustered'?m.grouping:'standard';
  const cache=(arr,num,fmt)=>`${num?`<c:formatCode>${esc(fmt||'General')}</c:formatCode>`:''}<c:ptCount val="${arr.length}"/>${arr.map((v,i)=>v===null||v===undefined||v===''?'':`<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join('')}`;
  const lbl=on=>on?`<c:dLbls><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:showLegendKey val="0"/><c:showVal val="${m.labels?1:0}"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="${pie&&m.pct?1:0}"/><c:showBubbleSize val="0"/>${pie?'':''}</c:dLbls>`:'';
  const ser=S.map((s,i)=>{
    const raw=m.series[i];
    const tx=raw.nameRef?`<c:tx><c:strRef><c:f>${esc(raw.nameRef)}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${esc(s.name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`:s.name?`<c:tx><c:v>${esc(s.name)}</c:v></c:tx>`:'';
    const col=s.color||autoColor(i,opt.acc||DEF_ACCENT);
    const sp=m.type==='line'||m.type==='scatter'?`<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${col}"/></a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="${m.markers?'circle':'none'}"/>${m.markers?`<c:size val="5"/><c:spPr><a:solidFill><a:srgbClr val="${col}"/></a:solidFill></c:spPr>`:''}</c:marker>`
      :pie?`<c:spPr><a:ln w="12700"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr>`
      :`<c:spPr><a:solidFill><a:srgbClr val="${col}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>${m.type==='bar'?'<c:invertIfNegative val="0"/>':''}`;
    let cat='';
    if(raw.cat&&(raw.cat.f||s.cats)){const nums=s.xs&&s.xs.every((x,k)=>x!==null||!s.cats[k]);
      const f=raw.cat.f?`<c:f>${esc(raw.cat.f)}</c:f>`:'';
      cat=nums&&raw.cat.f?`<c:cat><c:numRef>${f}<c:numCache>${cache(s.xs,true,raw.cat.fmt||m.catFmt)}</c:numCache></c:numRef></c:cat>`
        :raw.cat.f?`<c:cat><c:strRef>${f}<c:strCache>${cache(s.cats||[],false)}</c:strCache></c:strRef></c:cat>`
        :`<c:cat><c:strLit>${cache(s.cats||[],false)}</c:strLit></c:cat>`;}
    const vals=raw.val&&raw.val.f?`<c:val><c:numRef><c:f>${esc(raw.val.f)}</c:f><c:numCache>${cache(s.vals,true,'General')}</c:numCache></c:numRef></c:val>`
      :`<c:val><c:numLit>${cache(s.vals,true,'General')}</c:numLit></c:val>`;
    return`<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${sp}${cat}${vals}${m.type==='line'||m.type==='scatter'?'<c:smooth val="0"/>':''}</c:ser>`;
  }).join('');
  const ax='<c:axId val="500000001"/><c:axId val="500000002"/>';
  let group;
  if(m.type==='bar')group=`<c:barChart><c:barDir val="${m.dir==='bar'?'bar':'col'}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${ser}${lbl(m.labels)}<c:gapWidth val="${m.gap||150}"/>${/stacked/i.test(grouping)?'<c:overlap val="100"/>':''}${ax}</c:barChart>`;
  else if(m.type==='area')group=`<c:areaChart><c:grouping val="${grouping}"/><c:varyColors val="0"/>${ser}${lbl(m.labels)}${ax}</c:areaChart>`;
  else if(m.type==='pie')group=`<c:pieChart><c:varyColors val="1"/>${ser}${lbl(m.labels||m.pct)}<c:firstSliceAng val="0"/></c:pieChart>`;
  else if(m.type==='doughnut')group=`<c:doughnutChart><c:varyColors val="1"/>${ser}${lbl(m.labels||m.pct)}<c:firstSliceAng val="0"/><c:holeSize val="${m.hole||50}"/></c:doughnutChart>`;
  else group=`<c:lineChart><c:grouping val="${grouping}"/><c:varyColors val="0"/>${ser}${lbl(m.labels)}<c:marker val="1"/>${ax}</c:lineChart>`;
  const txt='<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="he-IL"/></a:p></c:txPr>';
  const line='<c:spPr><a:noFill/><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr>';
  const horiz=m.type==='bar'&&m.dir==='bar';
  const axes=pie?'':`<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="${m.catRev?'maxMin':'minMax'}"/></c:scaling><c:delete val="${m.catHide?1:0}"/><c:axPos val="${horiz?(m.valPos==='r'?'r':'l'):'b'}"/>${m.catFmt?`<c:numFmt formatCode="${esc(m.catFmt)}" sourceLinked="1"/>`:''}<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${line}${txt}<c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`
    +`<c:valAx><c:axId val="500000002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${m.valHide?1:0}"/><c:axPos val="${horiz?'b':(m.valPos==='r'?'r':'l')}"/>${m.grid!==false?'<c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>':''}<c:numFmt formatCode="${esc(m.valFmt||'General')}" sourceLinked="${m.valFmt?0:1}"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${txt}<c:crossAx val="500000001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
  const title=m.title?`<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:defRPr></a:pPr><a:r><a:rPr lang="he-IL" sz="1400" b="0"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:rPr><a:t>${esc(m.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`:'';
  const legend=m.legend?`<c:legend><c:legendPos val="${m.legend}"/><c:overlay val="0"/>${txt}</c:legend>`:'';
  return`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<c:chartSpace xmlns:c="${NC}" xmlns:a="${NA}" xmlns:r="${NR}"><c:date1904 val="0"/><c:lang val="he-IL"/><c:roundedCorners val="0"/>`
    +`<c:chart>${title}<c:autoTitleDeleted val="${m.title?0:1}"/><c:plotArea><c:layout/>${group}${axes}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>`
    +`<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr>`
    +`<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="he-IL"/></a:p></c:txPr><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>`;
}
// the anchor that places a chart over the cells
function anchorXml(from,to,rid,id,name){
  const P=(t,p)=>`<xdr:${t}><xdr:col>${p.c}</xdr:col><xdr:colOff>${Math.round(p.co*EMU)}</xdr:colOff><xdr:row>${p.r}</xdr:row><xdr:rowOff>${Math.round(p.ro*EMU)}</xdr:rowOff></xdr:${t}>`;
  return`<xdr:twoCellAnchor xmlns:xdr="${NX}" xmlns:a="${NA}">${P('from',from)}${P('to',to)}<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${esc(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${NC}"><c:chart xmlns:c="${NC}" xmlns:r="${NR}" r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
}

window.XL_CHARTS={NS:{c:NC,a:NA,r:NR,xdr:NX,rel:NREL},RT_CHART,RT_DRAW,CT_CHART,CT_DRAW,EMU,
  accents,parse,readSheet,svg,chartXml,anchorXml,data,splitRef,refText,fmtNum,autoColor,relsPath,resolve,rels};
})();
