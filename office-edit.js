/* סימן — Word / Excel files on the editing desk.
   toPdf():  renders a .docx (docx-preview) or .xlsx (SheetJS) into page images and a PDF,
             and remembers where every page came from in the original file.
   exportDocx() / exportXlsx(): write signatures, pictures, text boxes and drawings back into
             the original file, so it still opens and edits in Word / Excel.
   Needs JSZip, html2canvas, PDFLib, and docx (docx-preview) or XLSX (SheetJS). */
(function(){
'use strict';
const EMU=9525;                 // 1 CSS px (96 dpi) in EMU
const NS={
  w:'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r:'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp:'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a:'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic:'http://schemas.openxmlformats.org/drawingml/2006/picture',
  wps:'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  mc:'http://schemas.openxmlformats.org/markup-compatibility/2006',
  xdr:'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
  rel:'http://schemas.openxmlformats.org/package/2006/relationships',
  ct:'http://schemas.openxmlformats.org/package/2006/content-types'
};
const REL_IMG='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const REL_DRAW='http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
const XML_HEAD='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'})[c]);
const HEB=/[\u0590-\u05FF\u0600-\u06FF]/;
const FONT_NAME={alef:'Alef',heebo:'Heebo',rubik:'Rubik',noto:'Noto Sans Hebrew',david:'David',frank:'Frank Ruehl'};
const hex=c=>{c=String(c||'#000').trim();if(/^#[0-9a-f]{3}$/i.test(c))c='#'+c.slice(1).split('').map(x=>x+x).join('');
  const m=c.match(/^#([0-9a-f]{6})/i);if(m)return m[1].toUpperCase();
  const r=c.match(/rgba?\((\d+)\D+(\d+)\D+(\d+)/);return r?[r[1],r[2],r[3]].map(v=>(+v).toString(16).padStart(2,'0')).join('').toUpperCase():'000000';};
const nextFrame=()=>new Promise(r=>requestAnimationFrame(()=>setTimeout(r,0)));

/* ------------------------------------------------------------ rendering */
function makeHost(width){
  const h=document.createElement('div');
  h.setAttribute('aria-hidden','true');
  Object.assign(h.style,{position:'absolute',left:'0',top:'0',width:width?width+'px':'auto',zIndex:'-1',background:'#fff',color:'#000',pointerEvents:'none',direction:'ltr'});
  document.body.append(h);return h;
}
async function snap(el,scale,w,h){
  return html2canvas(el,{scale,backgroundColor:'#ffffff',logging:false,useCORS:true,
    width:w||el.offsetWidth,height:h||el.scrollHeight,windowWidth:Math.max(document.documentElement.clientWidth,(w||el.offsetWidth)+20),
    windowHeight:Math.max(document.documentElement.clientHeight,h||el.scrollHeight),scrollX:0,scrollY:-window.scrollY});
}
const jpeg=c=>new Promise(r=>c.toBlob(b=>b.arrayBuffer().then(x=>r(new Uint8Array(x))),'image/jpeg',.9));

async function pagesToPdf(pages){
  const {PDFDocument}=PDFLib;const doc=await PDFDocument.create();
  for(const p of pages){
    const img=await doc.embedJpg(await jpeg(p.canvas));
    const W=p.w*.75,H=p.h*.75;const pg=doc.addPage([W,H]);
    pg.drawImage(img,{x:0,y:0,width:W,height:H});
  }
  return doc.save();
}

/* ---------- Word ---------- */
// the top and bottom of every line of text, picture and short table row, in px from the top of the section
/* 6.2.1: the blanks of a paragraph — runs of underscores or dots, the lines of a form ("שם: ______").
   Their place on the shown page is kept, in the order they appear in the text, so a value written on one
   of them can be put back into the Word file on that very line (and Word lays it out itself). */
const BLANK_RE=/_{3,}|\.{5,}|…{2,}/g;
function blanksOf(p,R){
  const walker=document.createTreeWalker(p,NodeFilter.SHOW_TEXT);const nodes=[];let all='',n;
  while((n=walker.nextNode())){nodes.push([n,all.length]);all+=n.data;}
  const out=[];BLANK_RE.lastIndex=0;let m,prevEnd=0;
  // where a range of the paragraph's text sits on the page
  const rectOf=(s,e)=>{let x0=1e9,y0=1e9,x1=-1e9,y1=-1e9;
    for(const[node,o]of nodes){const a=Math.max(s,o),b=Math.min(e,o+node.data.length);if(a>=b)continue;
      const rg=document.createRange();rg.setStart(node,a-o);rg.setEnd(node,b-o);
      for(const r of rg.getClientRects()){if(!r.width)continue;x0=Math.min(x0,r.left-R.left);x1=Math.max(x1,r.right-R.left);y0=Math.min(y0,r.top-R.top);y1=Math.max(y1,r.bottom-R.top);}}
    return x1>x0?{x0,x1,y0,y1}:null;};
  while((m=BLANK_RE.exec(all))){
    const s=m.index,e=s+m[0].length,r=rectOf(s,e);
    // 6.3: the words before the blank on its line ("שם התלמיד/ה:") — what it asks for, and where they stand
    const raw=all.slice(prevEnd,s),lab=raw.replace(/\s+/g,' ').trim().slice(-40),ls=prevEnd+raw.search(/\S|$/);
    const lr=lab?rectOf(ls,s):null;
    out.push(r?{...r,lab,lr:lr&&r&&Math.abs(lr.y1-r.y1)<(r.y1-r.y0)?lr:null}:null);prevEnd=e;
  }
  return out;
}
function lineBoxes(sec,R,H){
  const out=[],rg=document.createRange(),tw=document.createTreeWalker(sec,NodeFilter.SHOW_TEXT);
  for(let t=tw.nextNode();t;t=tw.nextNode()){
    if(!t.nodeValue.trim())continue;
    rg.selectNodeContents(t);
    for(const r of rg.getClientRects())if(r.height>0&&r.width>0)out.push([r.top-R.top,r.bottom-R.top]);
  }
  for(const el of sec.querySelectorAll('img,svg,canvas,tr')){
    const r=el.getBoundingClientRect();if(r.height>0&&r.height<H*.5)out.push([r.top-R.top,r.bottom-R.top]);
  }
  return out;
}
/* 6.6.11: the copy that is drawn into pages. docx-preview does not know right-to-left tables (bidiVisual) and leaves
   simple fields empty, so in the copy: a right-to-left table is turned around (its cells in the visual order, a plain
   left-to-right table that looks the same), simple fields keep their shown text, and the total of pages is marked to
   be filled in once the pages are counted. The file itself is not changed. */
const NUMP='⁡';
function flipRtlTables(d,textW){
  const W=NS.w,kid=(el,n)=>{for(const c of el.children)if(c.namespaceURI===W&&c.localName===n)return c;return null;};
  const kidsOf=(el,n)=>[...el.children].filter(c=>c.namespaceURI===W&&c.localName===n);
  const rename=(el,n)=>{const e=d.createElementNS(W,'w:'+n);for(const a of el.attributes)e.setAttributeNS(a.namespaceURI,a.name,a.value);while(el.firstChild)e.appendChild(el.firstChild);el.replaceWith(e);};
  const swapLR=box=>{if(!box)return;for(const c of [...box.children]){if(c.localName==='left'||c.localName==='start')rename(c,'right');else if(c.localName==='right'||c.localName==='end')rename(c,'left');}};
  const val=el=>el&&(el.getAttributeNS(W,'val')||el.getAttribute('w:val'));
  const num=el=>el?+(el.getAttributeNS(W,'w')||el.getAttribute('w:w')||0):0;
  for(const tbl of [...d.getElementsByTagNameNS(W,'tbl')]){
    const pr=kid(tbl,'tblPr');const bv=pr&&kid(pr,'bidiVisual');if(!bv||/^(0|false|off)$/.test(val(bv)||''))continue;
    bv.remove();
    const grid=kid(tbl,'tblGrid');let gw=0;if(grid){const cols=kidsOf(grid,'gridCol');cols.reverse().forEach(c=>{grid.appendChild(c);gw+=num(c);});}
    swapLR(kid(pr,'tblBorders'));swapLR(kid(pr,'tblCellMar'));
    const jc=kid(pr,'jc');
    if(jc){const v=val(jc);const nv=v==='left'||v==='start'?'right':v==='right'||v==='end'?'left':v;jc.setAttributeNS(W,'w:val',nv);}
    else if(textW){const ind=kid(pr,'tblInd'),tw=kid(pr,'tblW');const wd=tw&&(tw.getAttributeNS(W,'type')||tw.getAttribute('w:type'))==='dxa'?num(tw):gw;
      if(wd){const x=num(ind);let ni=kid(pr,'tblInd');if(!ni){ni=d.createElementNS(W,'w:tblInd');pr.appendChild(ni);}ni.setAttributeNS(W,'w:w',String(Math.round(textW-x-wd)));ni.setAttributeNS(W,'w:type','dxa');}}
    for(const tr of kidsOf(tbl,'tr')){
      const tcs=kidsOf(tr,'tc');tcs.reverse().forEach(tc=>{tr.appendChild(tc);const p=kid(tc,'tcPr');if(p){swapLR(kid(p,'tcBorders'));swapLR(kid(p,'tcMar'));}});
      const tp=kid(tr,'trPr');if(tp){const gb=kid(tp,'gridBefore'),ga=kid(tp,'gridAfter');if(gb)rename(gb,'gridAfter');if(ga)rename(ga,'gridBefore');}
    }
  }
}
async function docxForPreview(bytes){
  let zip;try{zip=await JSZip.loadAsync(bytes);}catch(_){return bytes;}
  const parts=Object.keys(zip.files).filter(n=>/^word\/(document|header\d*|footer\d*)\.xml$/.test(n));
  let changed=false,textW=null;
  const dx=zip.file('word/document.xml')?await zip.file('word/document.xml').async('string'):'';
  const pw=/<w:pgSz\b[^>]*\bw:w="(\d+)"/.exec(dx),ml=/<w:pgMar\b[^>]*\bw:left="(-?\d+)"/.exec(dx),mr=/<w:pgMar\b[^>]*\bw:right="(-?\d+)"/.exec(dx);
  if(pw)textW=+pw[1]-(ml?+ml[1]:1440)-(mr?+mr[1]:1440);
  for(const n of parts){
    let x=await zip.file(n).async('string');const before=x;
    // the total of pages: its shown text is marked (filled in later); simple fields keep their runs
    x=x.replace(/(NUMPAGES\b[^<]*<\/w:instrText>[\s\S]*?<w:fldChar\b[^>]*w:fldCharType="separate"[^>]*\/>[\s\S]*?<w:t(?:\s[^>]*)?>)([^<]*)(<\/w:t>)/g,(m,a,t,b)=>a+NUMP+t+NUMP+b);
    x=x.replace(/<w:fldSimple\b[^>]*w:instr="\s*NUMPAGES[^"]*"[^>]*>([\s\S]*?)<\/w:fldSimple>/g,(m,inner)=>inner.replace(/(<w:t(?:\s[^>]*)?>)([^<]*)(<\/w:t>)/,(m2,a,t,b)=>a+NUMP+t+NUMP+b));
    x=x.replace(/<w:fldSimple\b[^>]*\/>/g,'').replace(/<w:fldSimple\b[^>]*>/g,'').replace(/<\/w:fldSimple>/g,'');
    if(/bidiVisual/.test(x)){try{const d=new DOMParser().parseFromString(x,'application/xml');if(!d.getElementsByTagName('parsererror').length){flipRtlTables(d,textW);x=new XMLSerializer().serializeToString(d);}}catch(e){console.warn('rtl tables',e);}}
    if(x!==before){zip.file(n,x);changed=true;}
  }
  return changed?await zip.generateAsync({type:'uint8array'}):bytes;
}
async function docxToPages(bytes,progress){
  const host=makeHost(0);
  try{
    try{bytes=await docxForPreview(bytes);}catch(e){console.warn('preview copy',e);}
    await docx.renderAsync(bytes,host,host,{className:'dx',inWrapper:false,ignoreWidth:false,ignoreHeight:false,ignoreFonts:false,
      breakPages:true,ignoreLastRenderedPageBreak:false,experimental:true,trimXmlDeclaration:true,useBase64URL:true,
      renderHeaders:true,renderFooters:true,renderFootnotes:true,renderEndnotes:true});
    const fix=document.createElement('style');
    fix.textContent='section.dx{margin:0!important;box-shadow:none!important;overflow:visible!important}section.dx+section.dx{margin-top:0!important}';
    host.append(fix);
    if(document.fonts&&document.fonts.ready)await document.fonts.ready;
    await Promise.all([...host.querySelectorAll('img')].map(im=>im.complete?0:new Promise(r=>{im.onload=im.onerror=r;})));
    await nextFrame();
    const secs=[...host.querySelectorAll('section.dx')];
    if(!secs.length)throw new Error('לא נמצא תוכן במסמך');
    const out=[],paras=[],plan=[];
    let n=0;
    for(let si=0;si<secs.length;si++){
      const sec=secs[si],R=sec.getBoundingClientRect();
      const cs=getComputedStyle(sec);
      const W=Math.round(sec.offsetWidth);
      let H=parseFloat(cs.minHeight)||Math.round(W*1.4142);
      const padT=parseFloat(cs.paddingTop)||0,padB=parseFloat(cs.paddingBottom)||0;
      const total=Math.max(sec.scrollHeight,H);
      // body-level paragraphs (for anchoring marks back into the file)
      const art=sec.querySelector(':scope>article')||sec;
      for(const p of art.children)if(p.tagName==='P')paras.push({sec:si,top:p.getBoundingClientRect().top-R.top,blanks:blanksOf(p,R)});
      // safe places to cut a page that runs longer than one sheet
      const cuts=[...art.querySelectorAll('p,tr,li,img,h1,h2,h3,h4,h5,h6')].map(e=>e.getBoundingClientRect().bottom-R.top).filter(b=>b>0).sort((a,b)=>a-b);
      // every line of text and every picture: a page is never cut through one of them (a cut through a line
      // left the bottoms of its letters at the foot of the page, like a dotted line between the pages)
      const boxes=lineBoxes(sec,R,H);
      const lineEnds=[...new Set(boxes.map(b=>Math.round(b[1]*2)/2))].sort((a,b)=>a-b);
      const free=y=>!boxes.some(([a,b])=>a<y-.5&&b>y+.5);
      const slices=[];
      if(total<=H*1.02)slices.push([0,H,0]);
      else{
        let s=0,first=true;
        while(s<total-padB-2){
          const room=H-(first?0:padT)-padB;
          let e=s+room;
          if(e<total-padB){
            const ok=cuts.filter(b=>b<=e&&b>s+room*.45&&free(b));
            if(ok.length)e=ok[ok.length-1];
            else{const ln=lineEnds.filter(b=>b<=e&&b>s+room*.3&&free(b));
              if(ln.length)e=ln[ln.length-1];
              else{let y=Math.floor(e);while(y>s+room*.3&&!free(y))y--;if(y>s+room*.3)e=y;}}
          }
          else e=total;
          slices.push([s,e,first?0:padT]);s=e;first=false;
        }
      }
      plan.push({sec,si,W,H,total,slices});n+=slices.length;
    }
    // the total of pages is known now: fill it in where the file shows it (NUMPAGES), then draw
    {const tw=document.createTreeWalker(host,NodeFilter.SHOW_TEXT);const hits=[];while(tw.nextNode())if(tw.currentNode.nodeValue.includes(NUMP))hits.push(tw.currentNode);
      for(const t of hits)t.nodeValue=t.nodeValue.replace(new RegExp(NUMP+'[^'+NUMP+']*'+NUMP,'g'),String(n)).split(NUMP).join('');
      if(hits.length)await nextFrame();}
    for(const{sec,si,W,H,total,slices}of plan){
      const scale=Math.min(2,30000/total);
      progress&&progress(`מכין את העמודים… ${si+1}/${secs.length}`);
      const full=await snap(sec,scale,W,total);
      for(const[s,e,pad]of slices){
        const c=document.createElement('canvas');c.width=Math.round(W*2);c.height=Math.round(H*2);
        const g=c.getContext('2d');g.fillStyle='#fff';g.fillRect(0,0,c.width,c.height);
        const k=full.width/W;
        g.drawImage(full,0,Math.round(s*k),full.width,Math.round((e-s)*k),0,Math.round(pad*2),c.width,Math.round((e-s)*2));
        out.push({canvas:c,w:W,h:H,src:{sec:si,top:s,pad}});
      }
    }
    return{pages:out,map:{kind:'docx',paras,pages:out.map(p=>({w:p.w,h:p.h,...p.src}))}};
  }finally{host.remove();}
}

/* ---------- Excel ---------- */
async function sheetFiles(zip){
  const wb=await zip.file('xl/workbook.xml').async('string');
  const rels=await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const rel={};rels.replace(/<Relationship\b[^>]*>/g,t=>{const id=(t.match(/\bId="([^"]+)"/)||[])[1],tg=(t.match(/\bTarget="([^"]+)"/)||[])[1];if(id&&tg)rel[id]=tg;return t;});
  const out=[];wb.replace(/<sheet\b[^>]*>/g,t=>{
    const name=(t.match(/\bname="([^"]*)"/)||[])[1],id=(t.match(/\br:id="([^"]+)"/)||t.match(/:id="([^"]+)"/)||[])[1];
    let tg=rel[id]||'';tg=tg.startsWith('/')?tg.slice(1):'xl/'+tg.replace(/^\.\//,'');
    out.push({name:name&&name.replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'"),path:tg});return t;});
  return out;
}
async function xlsxToPages(bytes,progress){
  const zip=await JSZip.loadAsync(bytes);
  const files=await sheetFiles(zip);
  const wb=XLSX.read(bytes,{type:'array',cellStyles:true,cellNF:true});
  const host=makeHost(0);
  const pages=[],mp=[];
  // charts (xl-charts.js): the parts they need, read once
  const XC=window.XL_CHARTS,T=new Map();
  if(XC)for(const p of Object.keys(zip.files))if(/^xl\/(drawings|charts|theme|worksheets\/_rels)\/.*\.(xml|rels)$/.test(p))T.set(p,await zip.file(p).async('string'));
  const getT=p=>T.has(p)?T.get(p):null;const acc=XC?XC.accents(getT):null;
  const cellsOf=f=>{const R=XC.splitRef(f);if(!R)return null;const w2=wb.Sheets[R.sheet];if(!w2)return null;if((R.r2-R.r1+1)*(R.c2-R.c1+1)>20000)return null;const out=[];
    for(let r=R.r1;r<=R.r2;r++)for(let c=R.c1;c<=R.c2;c++){const cell=w2[XLSX.utils.encode_cell({r,c})];if(!cell){out.push({v:null,w:''});continue;}
      const date=cell.t==='n'&&cell.z&&/^m+\/d+\/y+$/i.test(String(cell.z));out.push({v:cell.t==='e'?null:cell.v,w:date?XC.fmtNum(cell.v,'m/d/yyyy'):(cell.w!==undefined?cell.w:null)});}
    return out;};
  try{
    for(let si=0;si<wb.SheetNames.length;si++){
      const meta=(wb.Workbook&&wb.Workbook.Sheets||[])[si]||{};if(meta.Hidden)continue;
      const ws=wb.Sheets[wb.SheetNames[si]];if(!ws||!ws['!ref'])continue;
      const f=files[si]||{};let rtl=false,charts=[];
      if(f.path&&zip.file(f.path)){const x=await zip.file(f.path).async('string');rtl=/<sheetView\b[^>]*rightToLeft="(1|true)"/.test(x);
        if(XC&&/<(\w+:)?drawing\b/.test(x))try{charts=XC.readSheet(getT,f.path,x,acc).charts.filter(c=>c.from);}catch(e){console.warn('charts',e);}}
      const rg=XLSX.utils.decode_range(ws['!ref']);
      let chR=0,chC=0;for(const ch of charts){const t=ch.to||{c:ch.from.c+8,r:ch.from.r+15};chR=Math.max(chR,t.r);chC=Math.max(chC,t.c);}
      const R1=Math.min(Math.max(rg.e.r,chR),rg.s.r+4999),C1=Math.min(Math.max(rg.e.c,chC),rg.s.c+79);
      // trim empty trailing rows / columns
      let lastR=0,lastC=0;
      for(const k in ws){if(k[0]==='!')continue;const a=XLSX.utils.decode_cell(k),c=ws[k];if(c&&(c.v!==undefined&&c.v!==''||c.s&&c.s.fgColor&&c.s.patternType==='solid')){if(a.r>lastR)lastR=a.r;if(a.c>lastC)lastC=a.c;}}
      for(const m of ws['!merges']||[]){if(m.e.r>lastR)lastR=m.e.r;if(m.e.c>lastC)lastC=m.e.c;}
      // the page also takes in the charts that sit next to the table
      for(const ch of charts){const t=ch.to||{c:ch.from.c+8,r:ch.from.r+15};lastR=Math.max(lastR,t.r);lastC=Math.max(lastC,t.c);}
      const rN=Math.min(R1,lastR),cN=Math.min(C1,lastC);
      const cols=ws['!cols']||[],rows=ws['!rows']||[];
      const cw=[],rh=[];
      for(let c=0;c<=cN;c++){const d=cols[c]||{};cw.push(d.hidden?0:Math.round(d.wpx||(d.wch?d.wch*7+5:d.width?d.width*7+5:64)));}
      for(let r=0;r<=rN;r++){const d=rows[r]||{};rh.push(d.hidden?0:Math.round(d.hpx||(d.hpt?d.hpt*4/3:20)));}
      const colL=[0];cw.forEach(w=>colL.push(colL[colL.length-1]+w));
      const rowT=[0];rh.forEach(h=>rowT.push(rowT[rowT.length-1]+h));
      const tableW=colL[colL.length-1];
      let pw=794,ph=1123;const M=38;
      if(tableW>pw-2*M){pw=1123;ph=794;}
      const scale=Math.min(1,(pw-2*M)/tableW);
      const usableH=(ph-2*M)/scale;
      // merges
      const cover=new Map();
      for(const m of ws['!merges']||[]){for(let r=m.s.r;r<=m.e.r;r++)for(let c=m.s.c;c<=m.e.c;c++)cover.set(r+','+c,m);}
      let r0=0;
      while(r0<=rN){
        let r1=r0,h=0;
        while(r1<=rN&&(h+rh[r1]<=usableH||r1===r0)){h+=rh[r1];r1++;}
        // a chart is not cut between two pages when it fits on one
        for(const ch of charts){const t=ch.to;if(t&&ch.from.r>r0&&ch.from.r<r1&&t.r>=r1&&rowT[Math.min(t.r,rN+1)]-rowT[ch.from.r]<=usableH)r1=ch.from.r;}
        progress&&progress(`מכין את הגיליון "${wb.SheetNames[si]}"…`);
        // one table per page
        const t=document.createElement('table');t.dir=rtl?'rtl':'ltr';
        Object.assign(t.style,{borderCollapse:'collapse',tableLayout:'fixed',width:tableW+'px',font:'14.5px Calibri, Arial, "SF-heebo", sans-serif',background:'#fff'});
        const cg=document.createElement('colgroup');cw.forEach(w=>{const c=document.createElement('col');c.style.width=w+'px';cg.append(c);});t.append(cg);
        for(let r=r0;r<r1;r++){
          const tr=document.createElement('tr');tr.style.height=rh[r]+'px';
          for(let c=0;c<=cN;c++){
            const m=cover.get(r+','+c);
            if(m){const top=Math.max(m.s.r,r0);if(!(r===top&&c===m.s.c))continue;}
            const td=document.createElement('td');
            if(m){const top=Math.max(m.s.r,r0);td.rowSpan=Math.min(m.e.r,r1-1)-top+1;td.colSpan=m.e.c-m.s.c+1;}
            const ref=XLSX.utils.encode_cell(m?{r:m.s.r,c:m.s.c}:{r,c});const cell=ws[ref];
            Object.assign(td.style,{border:'1px solid #D9D9D9',padding:'0 3px',overflow:'hidden',whiteSpace:'nowrap',height:rh[r]+'px',maxHeight:rh[r]+'px',lineHeight:'1.2',verticalAlign:'bottom'});
            if(rh[r]===0||cw[c]===0){td.style.padding='0';td.style.border='0';}
            if(cell){
              const txt=cell.w!==undefined?cell.w:(cell.v===undefined||cell.v===null?'':String(cell.v));
              td.textContent=(m&&r!==m.s.r)?'':txt;
              if(cell.t==='n'||cell.t==='d')td.style.textAlign=rtl?'left':'right';
              else{td.dir='auto';td.style.textAlign='start';}
              const s=cell.s;
              if(s&&s.patternType==='solid'&&s.fgColor&&s.fgColor.rgb){const rgb=String(s.fgColor.rgb).slice(-6);if(/^[0-9a-f]{6}$/i.test(rgb)&&rgb.toUpperCase()!=='FFFFFF')td.style.background='#'+rgb;}
            }
            tr.append(td);
          }
          t.append(tr);
        }
        host.innerHTML='';host.style.width=tableW+'px';host.append(t);
        await nextFrame();
        const cv=await snap(t,2*scale,tableW,t.offsetHeight);
        const c=document.createElement('canvas');c.width=pw*2;c.height=ph*2;
        const g=c.getContext('2d');g.fillStyle='#fff';g.fillRect(0,0,c.width,c.height);
        const left=rtl?pw-M-tableW*scale:M;
        g.drawImage(cv,Math.round(left*2),M*2,Math.round(tableW*scale*2),Math.round(t.offsetHeight*scale*2));
        for(const ch of charts){
          if(ch.from.r<r0||ch.from.r>=r1)continue;
          try{
            const X0=(colL[Math.min(ch.from.c,cN+1)]||0)+ch.from.co,Y0=rowT[ch.from.r]+ch.from.ro;
            let w,hh;if(ch.to){w=(colL[Math.min(ch.to.c,cN+1)]||0)+ch.to.co-X0;hh=(rowT[Math.min(ch.to.r,rN+1)]||0)+ch.to.ro-Y0;}else if(ch.ext){w=ch.ext.w;hh=ch.ext.h;}else continue;
            if(w<20||hh<20)continue;
            const sx=rtl?left+(tableW-X0-w)*scale:left+X0*scale,sy=M+(Y0-rowT[r0])*scale;
            const svg=XC.svg(ch.model,w,hh,cellsOf,{pxW:Math.round(w*scale*2),pxH:Math.round(hh*scale*2),acc});
            const url=URL.createObjectURL(new Blob([svg],{type:'image/svg+xml'}));
            try{const im=new Image();im.src=url;await im.decode();g.drawImage(im,Math.round(sx*2),Math.round(sy*2),Math.round(w*scale*2),Math.round(hh*scale*2));}finally{URL.revokeObjectURL(url);}
          }catch(e){console.warn('chart',e);}
        }
        pages.push({canvas:c,w:pw,h:ph});
        mp.push({w:pw,h:ph,sheet:si,path:f.path,r0,left,top:M,scale,tableW,rtl,colL,rowT:rowT.slice(0)});
        r0=r1;
      }
    }
  }finally{host.remove();}
  if(!pages.length)throw new Error('לא נמצא תוכן בגיליונות');
  return{pages,map:{kind:'xlsx',pages:mp}};
}

async function toPdf(kind,bytes,progress){
  const r=kind==='docx'?await docxToPages(bytes,progress):await xlsxToPages(bytes,progress);
  progress&&progress('בונה את העמודים…');
  const pdf=await pagesToPdf(r.pages);
  return{pdf,map:r.map};
}

/* ------------------------------------------------------------ package helpers */
const parse=s=>new DOMParser().parseFromString(s,'application/xml');
const ser=d=>XML_HEAD+new XMLSerializer().serializeToString(d.documentElement);
async function ensurePng(zip){
  const p='[Content_Types].xml';let ct=await zip.file(p).async('string');
  if(!/Extension="png"/i.test(ct))ct=ct.replace(/<Types\b[^>]*>/,m=>m+'<Default Extension="png" ContentType="image/png"/>');
  if(!/Extension="jpeg"/i.test(ct))ct=ct.replace(/<Types\b[^>]*>/,m=>m+'<Default Extension="jpeg" ContentType="image/jpeg"/>');
  zip.file(p,ct);
}
async function addOverride(zip,part,type){
  const p='[Content_Types].xml';let ct=await zip.file(p).async('string');
  if(!ct.includes(`PartName="/${part}"`))ct=ct.replace('</Types>',`<Override PartName="/${part}" ContentType="${type}"/></Types>`);
  zip.file(p,ct);
}
async function relsAdd(zip,relsPath,type,target){
  let x=zip.file(relsPath)?await zip.file(relsPath).async('string'):`${XML_HEAD}<Relationships xmlns="${NS.rel}"></Relationships>`;
  const ids=[...x.matchAll(/\bId="rId(\d+)s?"/g)].map(m=>+m[1]);let n=Math.max(0,...ids)+1;
  const id='rId'+n+'s';
  x=x.replace('</Relationships>',`<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`);
  zip.file(relsPath,x);return id;
}
function freeName(zip,dir,base,ext){let i=1;while(zip.file(`${dir}/${base}${i}.${ext}`))i++;return`${dir}/${base}${i}.${ext}`;}
const lines=t=>String(t||'').replace(/\r/g,'').split('\n');

/* ------------------------------------------------------------ Word export */
function wRunProps(it){
  const f=FONT_NAME[it.font]||'Arial',sz=Math.round((it.size||12)*2);
  return`<w:rPr><w:rFonts w:ascii="${esc(f)}" w:hAnsi="${esc(f)}" w:cs="${esc(f)}"/>${it.bold?'<w:b/><w:bCs/>':''}<w:color w:val="${hex(it.color)}"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`;
}
function wTextParas(it){
  return lines(it.text).map(l=>{const rtl=HEB.test(l)||(!l.trim()&&HEB.test(it.text));
    return`<w:p><w:pPr>${rtl?'<w:bidi/>':''}<w:spacing w:before="0" w:after="0" w:line="${Math.round(240*1.25)}" w:lineRule="auto"/><w:jc w:val="${rtl?'right':'left'}"/></w:pPr><w:r>${wRunProps(it)}${rtl?'<w:rtl/>':''}</w:rPr><w:t xml:space="preserve">${esc(l)}</w:t></w:r></w:p>`;}).join('');
}
function anchorOpen(id,x,y,cx,cy,z){
  return`<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${z}" behindDoc="0" locked="0" layoutInCell="0" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>${Math.round(x)}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>${Math.round(y)}</wp:posOffset></wp:positionV><wp:extent cx="${Math.round(cx)}" cy="${Math.round(cy)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${id}" name="סימן ${id}"/><wp:cNvGraphicFramePr/>`;
}
function wPic(id,rid,x,y,cx,cy,z){
  return`<w:r><w:drawing>${anchorOpen(id,x,y,cx,cy,z)}<a:graphic><a:graphicData uri="${NS.pic}"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="סימן ${id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round(cx)}" cy="${Math.round(cy)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
}
function wTextBox(id,it,x,y,cx,cy,z){
  return`<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>${anchorOpen(id,x,y,cx,cy,z)}<a:graphic><a:graphicData uri="${NS.wps}"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round(cx)}" cy="${Math.round(cy)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent>${wTextParas(it)}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"><a:spAutoFit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>`;
}
const FRAG_NS=`xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}" xmlns:wps="${NS.wps}" xmlns:mc="${NS.mc}"`;

async function exportDocx(src,map,items){
  const zip=await JSZip.loadAsync(src);
  const docPath='word/document.xml';
  const dom=parse(await zip.file(docPath).async('string'));
  if(dom.getElementsByTagName('parsererror').length)throw new Error('המסמך המקורי לא נקרא');
  const root=dom.documentElement;
  // declare the prefixes we use on the root, so Word sees them where it expects
  for(const[p,u]of Object.entries({r:NS.r,wp:NS.wp,a:NS.a,pic:NS.pic,wps:NS.wps,mc:NS.mc}))
    if(!root.getAttribute('xmlns:'+p))root.setAttributeNS('http://www.w3.org/2000/xmlns/','xmlns:'+p,u);
  const body=dom.getElementsByTagNameNS(NS.w,'body')[0];
  const bodyP=[...body.children].filter(e=>e.namespaceURI===NS.w&&e.localName==='p');
  if(!bodyP.length){const p=dom.createElementNS(NS.w,'w:p');const sp=[...body.children].find(e=>e.localName==='sectPr');body.insertBefore(p,sp||null);bodyP.push(p);}
  // rendered paragraph list ↔ file paragraph list (proportional if the counts differ)
  const R=map.paras,same=R.length===bodyP.length;
  const xmlIndex=i=>same?i:Math.min(bodyP.length-1,Math.round(i*(bodyP.length-1)/Math.max(1,R.length-1)));
  let id=Math.max(9000,...[...dom.getElementsByTagNameNS(NS.wp,'docPr')].map(e=>+e.getAttribute('id')||0))+1;
  let z=251700000;
  await ensurePng(zip);
  const relsPath='word/_rels/document.xml.rels';
  let placed=0,inline=0;
  for(const it of items){
    const pm=map.pages[it.page];if(!pm)continue;
    const yy=pm.top+(it.y-pm.pad);       // y inside the rendered section
    // 6.2.1: on a blank line of the form → written into the text there, so Word places it itself
    if(same&&(await intoBlank(zip,relsPath,R,bodyP,pm,it,yy,id))){id++;placed++;inline++;continue;}
    // last body paragraph at or above the mark, in the same section
    let pi=-1;for(let i=0;i<R.length;i++){if(R[i].sec!==pm.sec)continue;if(R[i].top<=yy+1)pi=i;else if(pi>=0)break;}
    if(pi<0)pi=R.findIndex(p=>p.sec===pm.sec);
    if(pi<0)pi=Math.max(0,R.length-1);
    const target=bodyP[xmlIndex(pi)]||bodyP[bodyP.length-1];
    const offY=(yy-(R[pi]?R[pi].top:0))*EMU,x=it.x*EMU,cx=Math.max(1,it.w)*EMU,cy=Math.max(1,it.h)*EMU;
    let xml;
    if(it.type==='text')xml=wTextBox(id,it,x,offY,cx,cy,z);
    else{const ext=it.ext||'png';const name=freeName(zip,'word/media','siman',ext);zip.file(name,it.bytes);
      const rid=await relsAdd(zip,relsPath,REL_IMG,name.replace('word/',''));xml=wPic(id,rid,x,offY,cx,cy,z);}
    const frag=parse(`<x ${FRAG_NS}>${xml}</x>`).documentElement.firstElementChild;
    target.appendChild(dom.importNode(frag,true));
    id++;z+=2;placed++;
  }
  zip.file(docPath,ser(dom));
  return{bytes:await zip.generateAsync({type:'uint8array',compression:'DEFLATE'}),placed,inline,exact:same};
}
/* 6.2.1: a value or a signature that sits on a blank line ("שם: ______") goes into the Word text in place of
   the underscores: text as an underlined run in the same font, a signature as a picture inside the line.
   (Floating boxes placed by their spot on the shown page landed elsewhere in Word: Word lays the line out with
   its own fonts and tab widths, so the blank is not where it was on the screen.) */
const W_NS='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
function findBlank(R,pm,it,yy){
  const ix0=it.x,ix1=it.x+Math.max(1,it.w),iy0=yy,iy1=yy+Math.max(1,it.h);
  let best=null,score=0;
  R.forEach((p,pi)=>{if(p.sec!==pm.sec||!p.blanks)return;
    p.blanks.forEach((b,k)=>{if(!b)return;
      const ox=Math.min(ix1,b.x1)-Math.max(ix0,b.x0);if(ox<=0)return;
      // the value sits on the line: its bottom near the line, or the line inside it (a signature goes a little above)
      const lineH=b.y1-b.y0,near=iy1>=b.y0-lineH*.6&&iy0<=b.y1+lineH*.4;if(!near)return;
      const s=ox/Math.min(ix1-ix0,b.x1-b.x0)-Math.abs((iy0+iy1)/2-(b.y0+b.y1)/2)/(lineH*4+1);
      if(s>score){score=s;best={pi,k};}});});
  return score>.25?best:null;
}
async function intoBlank(zip,relsPath,R,bodyP,pm,it,yy,id){
  const hit=findBlank(R,pm,it,yy);if(!hit)return false;
  const para=bodyP[hit.pi];if(!para)return false;
  // the text of the paragraph's own runs, in order (not inside text boxes or drawings)
  const inside=t=>{for(let n=t.parentNode;n&&n!==para;n=n.parentNode)if(/^(txbxContent|drawing|pict|AlternateContent)$/.test(n.localName))return true;return false;};
  const ts=[...para.getElementsByTagNameNS(W_NS,'t')].filter(t=>!inside(t));
  let all='';const at=[];for(const t of ts){at.push(all.length);all+=t.textContent;}
  // blanks already filled in this paragraph are gone from the text: the wanted one moved up by that many
  const used=para.__simanUsed||(para.__simanUsed=new Set());if(used.has(hit.k))return false;
  const want=hit.k-[...used].filter(u=>u<hit.k).length;
  const re=new RegExp(BLANK_RE.source,'g');let m,k=-1;while((m=re.exec(all))){if(++k===want)break;}
  if(!m||k!==want)return false;
  used.add(hit.k);
  const s=m.index,e=s+m[0].length;
  const a=ts.findIndex((t,i)=>s>=at[i]&&s<at[i]+t.textContent.length),b=ts.findIndex((t,i)=>e-1>=at[i]&&e-1<at[i]+t.textContent.length);
  if(a<0||b<0)return false;
  const doc=para.ownerDocument,tA=ts[a],runA=tA.parentNode,txtA=tA.textContent,txtB=ts[b].textContent;
  const before=txtA.slice(0,s-at[a]),after=txtB.slice(e-at[b]);
  const rPrA=[...runA.childNodes].find(c=>c.localName==='rPr');
  const mk=tag=>doc.createElementNS(W_NS,'w:'+tag);
  const newRun=()=>{const r=mk('r');if(rPrA)r.appendChild(rPrA.cloneNode(true));return r;};
  // the value's run
  const vr=newRun();
  if(it.type==='text'){
    let rPr=[...vr.childNodes].find(c=>c.localName==='rPr');if(!rPr){rPr=mk('rPr');vr.insertBefore(rPr,vr.firstChild);}
    for(const c of [...rPr.childNodes])if(c.localName==='u'||c.localName==='rtl')rPr.removeChild(c);
    const u=mk('u');u.setAttributeNS(W_NS,'w:val','single');rPr.appendChild(u);
    if(HEB.test(it.text))rPr.appendChild(mk('rtl'));
    lines(it.text).forEach((l,i)=>{if(i)vr.appendChild(mk('br'));const t=mk('t');t.setAttribute('xml:space','preserve');t.textContent=' '+l.trim()+' ';vr.appendChild(t);});
  }else{
    // a signature (or a check mark): a picture tied to this spot in the text, so it moves with the line in Word,
    // lying over the line, which stays. (A picture inside the line was cut to the height of the text when Word
    // saves as PDF.) It sits on the line, a little above it, and runs along the blank.
    const ext=it.ext||'png',name=freeName(zip,'word/media','siman',ext);zip.file(name,it.bytes);
    const rid=await relsAdd(zip,relsPath,REL_IMG,name.replace('word/',''));
    const bl=R[hit.pi].blanks[hit.k],blankW=Math.max(1,bl.x1-bl.x0);
    const sc=Math.min(1,blankW*1.15/Math.max(1,it.w),48/Math.max(1,it.h)),cx=Math.round(Math.max(1,it.w)*sc*EMU),cy=Math.round(Math.max(1,it.h)*sc*EMU);
    const szEl=rPrA&&[...rPrA.childNodes].find(c=>c.localName==='sz'),fontPt=szEl?(+szEl.getAttributeNS(W_NS,'val')||24)/2:12;
    const voff=Math.round(-(cy-fontPt*12700));
    const xml=`<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${251700000+id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="character"><wp:posOffset>${-cx}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="line"><wp:posOffset>${voff}</wp:posOffset></wp:positionV><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="${id}" name="סימן ${id}"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${NS.pic}"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="סימן ${id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
    const imported=doc.importNode(parse(`<x ${FRAG_NS}>${xml}</x>`).documentElement.firstElementChild,true);
    if(rPrA)imported.insertBefore(rPrA.cloneNode(true),imported.firstChild);
    // the run is split just before the blank: [before] [the picture] [the blank and what follows in that run]
    tA.textContent=before;tA.setAttribute('xml:space','preserve');
    const rest=newRun(),rt=mk('t');rt.setAttribute('xml:space','preserve');rt.textContent=txtA.slice(s-at[a]);rest.appendChild(rt);
    runA.parentNode.insertBefore(imported,runA.nextSibling);
    imported.parentNode.insertBefore(rest,imported.nextSibling);
    used.delete(hit.k); // the blank is still there (the picture lies over it): the count of blanks does not change
    return true;
  }
  return place(vr);
  function place(run){
    // the blank's underscores leave; what was before and after them in the same runs stays
    tA.textContent=before;tA.setAttribute('xml:space','preserve');
    for(let i=a+1;i<b;i++)ts[i].textContent='';
    let tail=null;
    if(a===b){if(after){tail=newRun();const t=mk('t');t.setAttribute('xml:space','preserve');t.textContent=after;tail.appendChild(t);}}
    else{ts[b].textContent=after;ts[b].setAttribute('xml:space','preserve');}
    runA.parentNode.insertBefore(run,runA.nextSibling);
    if(tail)run.parentNode.insertBefore(tail,run.nextSibling);
    return true;
  }
}

/* ------------------------------------------------------------ Excel export */
function cellAt(pm,px,py,w){
  // page px → logical sheet position (column A's edge is on the right in RTL sheets)
  const lx=pm.rtl?(pm.left+pm.tableW*pm.scale-(px+w))/pm.scale:(px-pm.left)/pm.scale;
  const ly=pm.rowT[pm.r0]+(py-pm.top)/pm.scale;
  const L=pm.colL,T=pm.rowT;
  let c=0;while(c<L.length-2&&L[c+1]<=lx)c++;
  let r=0;while(r<T.length-2&&T[r+1]<=ly)r++;
  return{c:Math.max(0,c),cOff:Math.max(0,lx-L[c]),r:Math.max(0,r),rOff:Math.max(0,ly-T[r])};
}
function xFrom(p,cx,cy){return`<xdr:from><xdr:col>${p.c}</xdr:col><xdr:colOff>${Math.round(p.cOff*EMU)}</xdr:colOff><xdr:row>${p.r}</xdr:row><xdr:rowOff>${Math.round(p.rOff*EMU)}</xdr:rowOff></xdr:from><xdr:ext cx="${Math.round(cx)}" cy="${Math.round(cy)}"/>`;}
function xPic(id,rid,p,cx,cy){return`<xdr:oneCellAnchor>${xFrom(p,cx,cy)}<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="סימן ${id}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round(cx)}" cy="${Math.round(cy)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`;}
function xText(id,it,p,cx,cy){
  const f=esc(FONT_NAME[it.font]||'Arial'),sz=Math.round((it.size||12)*100);
  const ps=lines(it.text).map(l=>{const rtl=HEB.test(l);return`<a:p><a:pPr algn="${rtl?'r':'l'}"${rtl?' rtl="1"':''}/><a:r><a:rPr lang="${rtl?'he-IL':'en-US'}" sz="${sz}"${it.bold?' b="1"':''}><a:solidFill><a:srgbClr val="${hex(it.color)}"/></a:solidFill><a:latin typeface="${f}"/><a:cs typeface="${f}"/></a:rPr><a:t>${esc(l)}</a:t></a:r></a:p>`;}).join('');
  return`<xdr:oneCellAnchor>${xFrom(p,cx,cy)}<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="סימן ${id}"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${Math.round(cx)}" cy="${Math.round(cy)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></xdr:spPr><xdr:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="t"><a:spAutoFit/></a:bodyPr><a:lstStyle/>${ps}</xdr:txBody></xdr:sp><xdr:clientData/></xdr:oneCellAnchor>`;
}
const SHEET_AFTER_DRAWING=/<(legacyDrawing|legacyDrawingHF|drawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b|<\/worksheet>/;
async function exportXlsx(src,map,items){
  const zip=await JSZip.loadAsync(src);
  await ensurePng(zip);
  const bySheet=new Map();
  for(const it of items){const pm=map.pages[it.page];if(!pm||!pm.path)continue;if(!bySheet.has(pm.path))bySheet.set(pm.path,[]);bySheet.get(pm.path).push([it,pm]);}
  let placed=0;
  for(const[sheetPath,list]of bySheet){
    let sx=await zip.file(sheetPath).async('string');
    const dir=sheetPath.replace(/[^/]+$/,''),fname=sheetPath.split('/').pop();
    const sRels=`${dir}_rels/${fname}.rels`;
    let drawPath=null;
    const dm=sx.match(/<drawing\b[^>]*\br:id="([^"]+)"/);
    if(dm&&zip.file(sRels)){const rx=await zip.file(sRels).async('string');const t=(rx.match(new RegExp(`<Relationship\\b[^>]*Id="${dm[1]}"[^>]*>`))||[''])[0];const tg=(t.match(/Target="([^"]+)"/)||[])[1];
      if(tg)drawPath=tg.startsWith('/')?tg.slice(1):new URL(tg,'http://x/'+dir).pathname.slice(1);}
    if(!drawPath||!zip.file(drawPath)){
      drawPath=freeName(zip,'xl/drawings','drawing',"xml");
      zip.file(drawPath,`${XML_HEAD}<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}" xmlns:r="${NS.r}"></xdr:wsDr>`);
      await addOverride(zip,drawPath,'application/vnd.openxmlformats-officedocument.drawing+xml');
      const rid=await relsAdd(zip,sRels,REL_DRAW,'../drawings/'+drawPath.split('/').pop());
      if(!/xmlns:r=/.test(sx.slice(0,sx.indexOf('>',sx.indexOf('<worksheet')))))sx=sx.replace(/<worksheet\b/,`<worksheet xmlns:r="${NS.r}"`);
      const m=sx.match(SHEET_AFTER_DRAWING);sx=sx.slice(0,m.index)+`<drawing r:id="${rid}"/>`+sx.slice(m.index);
      zip.file(sheetPath,sx);
    }
    let dx=await zip.file(drawPath).async('string');
    if(!/xmlns:a=/.test(dx))dx=dx.replace(/<xdr:wsDr\b/,`<xdr:wsDr xmlns:a="${NS.a}"`);
    if(!/xmlns:r=/.test(dx))dx=dx.replace(/<xdr:wsDr\b/,`<xdr:wsDr xmlns:r="${NS.r}"`);
    let id=Math.max(9000,...[...dx.matchAll(/<xdr:cNvPr\b[^>]*\bid="(\d+)"/g)].map(m=>+m[1]))+1;
    const dRels=drawPath.replace(/[^/]+$/,'')+'_rels/'+drawPath.split('/').pop()+'.rels';
    let add='';
    for(const[it,pm]of list){
      const p=cellAt(pm,it.x,it.y,it.w);
      const cx=Math.max(1,it.w)/pm.scale*EMU,cy=Math.max(1,it.h)/pm.scale*EMU;
      if(it.type==='text')add+=xText(id,{...it,size:(it.size||12)/pm.scale},p,cx,cy);
      else{const ext=it.ext||'png';const name=freeName(zip,'xl/media','siman',ext);zip.file(name,it.bytes);
        const rid=await relsAdd(zip,dRels,REL_IMG,'../media/'+name.split('/').pop());add+=xPic(id,rid,p,cx,cy);}
      id++;placed++;
    }
    dx=dx.replace(/<\/xdr:wsDr>\s*$/,add+'</xdr:wsDr>');
    zip.file(drawPath,dx);
  }
  return{bytes:await zip.generateAsync({type:'uint8array',compression:'DEFLATE'}),placed,exact:true};
}


/* ------------------------------------------------------------ old Word (.doc, 97–2003) */
// Reads the text and tables of a binary .doc and builds a clean .docx from them.
// (the text-only reader: used when the formatted reader below cannot read a file)
function docText(bytes){
  if(!window.XLSX||!XLSX.CFB)throw new Error('רכיב הקריאה לא נטען');
  let cfb;try{cfb=XLSX.CFB.read(bytes,{type:'array'});}catch(e){throw new Error('הקובץ לא נראה כמו מסמך Word');}
  const wd=XLSX.CFB.find(cfb,'WordDocument');if(!wd)throw new Error('הקובץ לא נראה כמו מסמך Word');
  const W=Uint8Array.from(wd.content),dv=new DataView(W.buffer);
  if(dv.getUint16(0,true)!==0xA5EC)throw new Error('הקובץ לא נראה כמו מסמך Word');
  const nFib=dv.getUint16(2,true),flags=dv.getUint16(0x0A,true);
  if(nFib<101)throw new Error('זה קובץ Word 6/95 ישן מאוד. פתחו ושמרו אותו כ-DOCX ב-Word');
  if(flags&0x0100)throw new Error('הקובץ מוגן בסיסמה. הסירו את הסיסמה ב-Word ונסו שוב');
  const tbl=XLSX.CFB.find(cfb,(flags&0x0200)?'1Table':'0Table');if(!tbl)throw new Error('חסר חלק בקובץ');
  const T=Uint8Array.from(tbl.content),tv=new DataView(T.buffer);
  const ccpText=dv.getInt32(76,true),fcClx=dv.getUint32(418,true),lcbClx=dv.getUint32(422,true);
  let p=fcClx;const end=fcClx+lcbClx;
  while(p<end&&T[p]===1){p+=3+tv.getUint16(p+1,true);}       // skip Prc
  if(T[p]!==2)throw new Error('לא הצלחתי לקרוא את הטקסט בקובץ');
  const lcb=tv.getUint32(p+1,true),plc=p+5,n=(lcb-4)/12;
  const cp=i=>tv.getUint32(plc+i*4,true);
  const dec1252=new TextDecoder('windows-1252'),dec16=new TextDecoder('utf-16le');
  let out='';
  for(let i=0;i<n&&out.length<ccpText;i++){
    const a=cp(i),b=cp(i+1),pcd=plc+(n+1)*4+i*8;let fc=tv.getUint32(pcd+2,true);
    const len=Math.min(b-a,ccpText-out.length);if(len<=0)continue;
    if(fc&0x40000000){fc=(fc&~0x40000000)/2;out+=dec1252.decode(W.subarray(fc,fc+len));}
    else out+=dec16.decode(W.subarray(fc,fc+len*2));
  }
  // field codes: keep only the shown result
  let clean='';const st=[];
  for(const ch of out){
    if(ch==='\x13'){st.push(0);continue;}
    if(ch==='\x14'){if(st.length)st[st.length-1]=1;continue;}
    if(ch==='\x15'){st.pop();continue;}
    if(st.some(v=>v===0))continue;
    if('\x01\x02\x05\x08\x1f'.includes(ch))continue;
    clean+=ch==='\x1e'?'-':ch;
  }
  return clean;
}
function docBlocks(t){
  const blocks=[];let row=[],table=null,cols=0,cur='',prev='';
  const endTable=()=>{if(table){blocks.push({t:'table',rows:table});table=null;cols=0;}};
  for(const ch of t){
    if(ch==='\x07'){
      const rowEnd=prev==='\x07'&&cur===''&&(cols?row.length>=cols:true)&&row.length>0;
      if(rowEnd){if(!table)table=[];if(!cols)cols=row.length;table.push(row);row=[];}
      else{row.push((row.pending||[]).concat(cur));row.pending=null;}
      cur='';prev=ch;continue;
    }
    if(ch==='\r'){
      if(row.length||(table&&row.pending)){(row.pending=row.pending||[]).push(cur);}  // a paragraph inside a cell
      else{endTable();blocks.push({t:'p',text:cur});}
      cur='';prev=ch;continue;
    }
    cur+=ch;prev=ch;
  }
  if(row.length){if(!table)table=[];table.push(row);}
  endTable();
  if(cur)blocks.push({t:'p',text:cur});
  return blocks;
}
function wPara(text,opts){
  const rtl=HEB.test(text)||(!text.trim()&&opts.rtl);
  const runs=[];
  String(text).split(/(\x0b|\x0c|\t)/).forEach(seg=>{
    if(seg==='\x0b')runs.push('<w:r><w:br/></w:r>');
    else if(seg==='\x0c')runs.push('<w:r><w:br w:type="page"/></w:r>');
    else if(seg==='\t')runs.push('<w:r><w:tab/></w:r>');
    else if(seg)runs.push(`<w:r>${rtl?'<w:rPr><w:rtl/></w:rPr>':''}<w:t xml:space="preserve">${esc(seg)}</w:t></w:r>`);
  });
  return`<w:p><w:pPr>${rtl?'<w:bidi/>':''}<w:jc w:val="${rtl?'right':'left'}"/></w:pPr>${runs.join('')}</w:p>`;
}
/* ------------------------------------------------------------ 6.6.11: old Word (.doc) WITH its formatting
   Reads the formatting kept in the binary file (as described in Microsoft's [MS-DOC]): styles, fonts, bold / italic /
   underline / size / colour, alignment, indents and spacing, bullets and numbering, table column widths, merged cells,
   cell shading and borders, row heights, page size and margins, the page header and footer, and their pictures
   (PNG / JPEG / BMP; drawings kept as EMF / WMF are left out). If any of it fails, the text-only reader is used. */
const DOC_ICO=[null,'000000','0000FF','00FFFF','00FF00','FF00FF','FF0000','FFFF00','FFFFFF','000080','008080','008000','800080','800000','808000','808080','C0C0C0'];
const DOC_HL=[null,'black','blue','cyan','green','magenta','red','yellow','white','darkBlue','darkCyan','darkGreen','darkMagenta','darkRed','darkYellow','darkGray','lightGray'];
// bullets drawn with the Symbol / Wingdings fonts, as ordinary characters that show everywhere
const DOC_SYM={0xB7:'•',0xA7:'▪',0x6E:'■',0xFC:'✓',0xD8:'➢',0x76:'❖',0xA8:'□',0x71:'❑',0x6C:'●',0x75:'◆',0xFE:'☑',0x6F:'○',0xE0:'→',0xE8:'➔',0x2D:'–',0x9F:'•',0x77:'⬥',0x73:'♦',0xF0:'⇨',0xDE:'➤'};
const DOC_PCT={0:0,1:1,2:.05,3:.1,4:.2,5:.25,6:.3,7:.4,8:.5,9:.6,10:.7,11:.75,12:.8,13:.9,35:.025,36:.075,37:.125,38:.15,39:.175,40:.225,41:.275,42:.325,43:.35,44:.375,45:.425,46:.45,47:.475,48:.525,49:.55,50:.575,51:.625,52:.65,53:.675,54:.725,55:.775,56:.825,57:.85,58:.875,59:.925,60:.95,61:.975,62:.97};
const DOC_UL={1:'single',2:'words',3:'double',4:'dotted',6:'thick',7:'dash',9:'dotDash',10:'dotDotDash',11:'wave',20:'dottedHeavy',23:'dashedHeavy',25:'dashDotHeavy',26:'dashDotDotHeavy',27:'wavyHeavy',39:'dashLong',43:'wavyDouble'};
const DOC_BRC={1:'single',2:'thick',3:'double',5:'single',6:'dotted',7:'dashed',8:'dotDash',9:'dotDotDash',10:'triple',11:'thinThickSmallGap',12:'thickThinSmallGap',13:'thinThickThinSmallGap',20:'wave',21:'doubleWave',22:'dashSmallGap',23:'dashDotStroked',24:'threeDEmboss',25:'threeDEngrave'};
const DOC_NFC={0:'decimal',1:'upperRoman',2:'lowerRoman',3:'upperLetter',4:'lowerLetter',5:'ordinal',6:'cardinalText',7:'ordinalText',22:'decimalZero',23:'bullet',45:'hebrew1',46:'arabicAlpha',47:'hebrew2',48:'arabicAbjad',255:'none'};
const hex2=n=>n.toString(16).padStart(2,'0').toUpperCase();
function docRead(bytes){
  if(!window.XLSX||!XLSX.CFB)throw new Error('רכיב הקריאה לא נטען');
  const cfb=XLSX.CFB.read(bytes,{type:'array'});
  const wd=XLSX.CFB.find(cfb,'WordDocument');if(!wd)throw new Error('not doc');
  const W=Uint8Array.from(wd.content),wv=new DataView(W.buffer);
  if(wv.getUint16(0,true)!==0xA5EC)throw new Error('not doc');
  const flags=wv.getUint16(0x0A,true);if(wv.getUint16(2,true)<101||(flags&0x0100))throw new Error('old or protected');
  const tb=XLSX.CFB.find(cfb,(flags&0x0200)?'1Table':'0Table');if(!tb)throw new Error('no table stream');
  const T=Uint8Array.from(tb.content),tv=new DataView(T.buffer);
  const ds=XLSX.CFB.find(cfb,'Data'),D=ds?Uint8Array.from(ds.content):new Uint8Array(0),dv=new DataView(D.buffer);
  // the FIB: fixed part, then the counts, then the (offset, size) pairs
  const csw=wv.getUint16(32,true),lwAt=34+csw*2+2,cslw=wv.getUint16(34+csw*2,true),fcAt=lwAt+cslw*4+2;
  const cbFc=wv.getUint16(lwAt+cslw*4,true);
  const lw=i=>wv.getInt32(lwAt+i*4,true);
  const fcl=i=>i<cbFc?[wv.getUint32(fcAt+i*8,true),wv.getUint32(fcAt+i*8+4,true)]:[0,0];
  const ccpText=lw(3),ccpFtn=lw(4),ccpHdd=lw(5);
  const u16=(a,o)=>a[o]|a[o+1]<<8,i16=(a,o)=>(u16(a,o)<<16)>>16,u32=(a,o)=>(a[o]|a[o+1]<<8|a[o+2]<<16|a[o+3]<<24)>>>0,i32=(a,o)=>a[o]|a[o+1]<<8|a[o+2]<<16|a[o+3]<<24;
  const dec16=new TextDecoder('utf-16le'),dec1252=new TextDecoder('windows-1252');
  // ---- sprms: a 2-byte code, then an operand whose size the code says
  function sprms(g){
    const out=[];let i=0;
    while(i+2<=g.length){
      const op=u16(g,i);i+=2;const spra=op>>13;let n;
      if(spra===6){
        if(op===0xD608||op===0xD606)n=u16(g,i)+1;
        else if(op===0xC615&&g[i]===255){const cd=g[i+1];n=3+cd*4+g[i+2+cd*4]*3;}
        else n=1+g[i];
      }else n=[1,1,2,4,2,2,0,3][spra];
      if(i+n>g.length)break;out.push([op,g.subarray(i,i+n)]);i+=n;
    }
    return out;
  }
  // ---- pieces of text (CP → place in the WordDocument stream), and their own formatting (Prc)
  const [fcClx,lcbClx]=fcl(33);let p=fcClx;const prc=[];
  while(p<fcClx+lcbClx&&T[p]===1){const cb=tv.getInt16(p+1,true);prc.push(T.subarray(p+3,p+3+cb));p+=3+cb;}
  if(T[p]!==2)throw new Error('no piece table');
  const lcbPcd=tv.getUint32(p+1,true),plc=p+5,npc=(lcbPcd-4)/12,pieces=[];
  for(let i=0;i<npc;i++){
    const cp0=tv.getUint32(plc+i*4,true),cp1=tv.getUint32(plc+i*4+4,true),pcd=plc+(npc+1)*4+i*8;
    let fc=tv.getUint32(pcd+2,true);const comp=!!(fc&0x40000000);if(comp)fc=(fc&~0x40000000)/2;
    const prm=tv.getUint16(pcd+6,true);
    pieces.push({cp0,cp1,fc,comp,g:(prm&1)?(prc[prm>>1]||null):null});
  }
  const pieceAt=cp=>{let lo=0,hi=pieces.length-1;while(lo<=hi){const m=(lo+hi)>>1,q=pieces[m];if(cp<q.cp0)hi=m-1;else if(cp>=q.cp1)lo=m+1;else return q;}return null;};
  const charAt=cp=>{const q=pieceAt(cp);if(!q)return['',0,null];const k=cp-q.cp0;
    if(q.comp){const fc=q.fc+k;return[dec1252.decode(W.subarray(fc,fc+1)),fc,q];}
    const fc=q.fc+k*2;return[String.fromCharCode(u16(W,fc)),fc,q];};
  // ---- formatting tables (FKP pages): character runs and paragraphs, found by their place in the stream
  function bte(i){const [fc,lcb]=fcl(i);const n=(lcb-4)/8,out=[];for(let k=0;k<n;k++)out.push(tv.getUint32(fc+(n+1)*4+k*4,true)&0x3FFFFF);return out;}
  const chpRuns=[],papRuns=[];
  for(const pn of bte(12)){const pg=W.subarray(pn*512,pn*512+512),n=pg[511];
    for(let k=0;k<n;k++){const b=pg[(n+1)*4+k];chpRuns.push([u32(pg,k*4),u32(pg,k*4+4),b?pg.subarray(b*2+1,b*2+1+pg[b*2]):new Uint8Array(0)]);}}
  for(const pn of bte(13)){const pg=W.subarray(pn*512,pn*512+512),n=pg[511];
    for(let k=0;k<n;k++){const b=pg[(n+1)*4+k*13];let istd=0,g=new Uint8Array(0);
      if(b){const at=b*2;let cb=pg[at],g0,len;if(cb===0){cb=pg[at+1];g0=at+2;len=cb*2;}else{g0=at+1;len=cb*2-1;}istd=u16(pg,g0);g=pg.subarray(g0+2,g0+len);}
      papRuns.push([u32(pg,k*4),u32(pg,k*4+4),g,istd]);}}
  const byFc=(arr,fc)=>{let lo=0,hi=arr.length-1;while(lo<=hi){const m=(lo+hi)>>1,r=arr[m];if(fc<r[0])hi=m-1;else if(fc>=r[1])lo=m+1;else return r;}return null;};
  // ---- fonts
  const fonts=[];{const [fc,lcb]=fcl(15);if(lcb){let q=fc;let n=u16(T,q);q+=2;if(n===0xFFFF){n=u16(T,q);q+=2;}const extra=u16(T,q);q+=2;
    for(let i=0;i<n&&q<fc+lcb;i++){const cb=T[q],s=q+40;let e=s;while(e+1<q+1+cb&&(T[e]|T[e+1]))e+=2;fonts.push(dec16.decode(T.subarray(s,e)));q+=1+cb+extra;}}}
  // ---- colours, borders, shading
  const cv=(a,o)=>a[o+3]===0xFF?null:hex2(a[o])+hex2(a[o+1])+hex2(a[o+2]);
  const brc80=(a,o)=>{const v=u32(a,o);if(!v||v===0xFFFFFFFF||!a[o+1]||a[o+1]===0xFF)return v===0xFFFFFFFF?undefined:null;return{val:DOC_BRC[a[o+1]]||'single',sz:Math.max(2,a[o]),color:DOC_ICO[a[o+2]]||'auto'};};
  const brc=(a,o)=>{const v=u32(a,o+4);if(v===0xFFFFFFFF)return undefined;if(!a[o+5]||a[o+5]===0xFF)return null;return{val:DOC_BRC[a[o+5]]||'single',sz:Math.max(2,a[o+4]),color:cv(a,o)||'auto'};};
  const mix=(f,b,pc)=>{const F=f?[0,2,4].map(i=>parseInt(f.substr(i,2),16)):[0,0,0],B=b?[0,2,4].map(i=>parseInt(b.substr(i,2),16)):[255,255,255];return F.map((x,i)=>hex2(Math.round(x*pc+B[i]*(1-pc)))).join('');};
  const shdFill=(fore,back,ipat)=>{if(ipat===0xFFFF)return undefined;const pc=DOC_PCT[ipat]!=null?DOC_PCT[ipat]:.25;if(!pc)return back||null;if(pc===1)return fore||'000000';return mix(fore,back,pc);};
  const shd=(a,o)=>shdFill(cv(a,o),cv(a,o+4),u16(a,o+8));
  const shd80=(a,o)=>{const w=u16(a,o);if(w===0xFFFF)return undefined;return shdFill(DOC_ICO[w&0x1F],DOC_ICO[(w>>5)&0x1F],w>>10);};
  const tog=(v,base)=>v===0?false:v===1?true:v===0x80?!!base:v===0x81?!base:!!v;
  // ---- character formatting
  let styleChp=null;
  function applyChp(c,g,base){
    for(const [op,a] of sprms(g))switch(op){
      case 0x4A30:Object.assign(c,styleChp(u16(a,0)));break;
      case 0x0835:c.b=tog(a[0],base.b);break;case 0x085C:c.bCs=tog(a[0],base.bCs);break;
      case 0x0836:c.i=tog(a[0],base.i);break;case 0x085D:c.iCs=tog(a[0],base.iCs);break;
      case 0x0837:c.strike=tog(a[0],base.strike);break;case 0x2A53:c.dstrike=tog(a[0],base.dstrike);break;
      case 0x083A:c.smallCaps=tog(a[0],base.smallCaps);break;case 0x083B:c.caps=tog(a[0],base.caps);break;
      case 0x083C:c.vanish=tog(a[0],base.vanish);break;case 0x0800:c.del=!!a[0];break;
      case 0x2A3E:c.u=DOC_UL[a[0]]||(a[0]?'single':null);break;
      case 0x2A42:c.color=DOC_ICO[a[0]&0x1F]||null;break;case 0x6870:c.color=cv(a,0);break;
      case 0x4A43:c.sz=u16(a,0);break;case 0x4A61:c.szCs=u16(a,0);break;
      case 0x2A48:c.va=a[0]===1?'superscript':a[0]===2?'subscript':null;break;
      case 0x4A4F:c.f0=u16(a,0);break;case 0x4A50:c.f1=u16(a,0);break;case 0x4A51:c.f2=u16(a,0);break;case 0x4A5E:c.fBi=u16(a,0);break;
      case 0x085A:c.rtl=tog(a[0],base.rtl);break;
      case 0x2A0C:c.hl=DOC_HL[a[0]]||null;break;
      case 0x4866:c.shd=shd80(a,0);break;case 0xCA71:c.shd=shd(a,1);break;
      case 0x0855:c.spec=!!a[0];break;case 0x6A03:c.pic=u32(a,0);break;case 0x0806:c.data=!!a[0];break;case 0x0856:c.obj=!!a[0];break;
      case 0x6A09:c.sym={ftc:u16(a,0),ch:u16(a,2)};break;
    }
    return c;
  }
  // ---- paragraph formatting (and the table row formatting, kept on the row's end mark)
  function applyPap(pp,g){
    for(const [op,a] of sprms(g))switch(op){
      case 0x2403:if(pp.jcL==null)pp.jc80=a[0];break;case 0x2461:pp.jcL=a[0];break;
      case 0x2441:pp.bidi=!!a[0];break;
      case 0x840F:if(pp.indS==null)pp.l80=i16(a,0);break;case 0x845E:pp.indS=i16(a,0);break;
      case 0x840E:if(pp.indE==null)pp.r80=i16(a,0);break;case 0x845D:pp.indE=i16(a,0);break;
      case 0x8411:case 0x8460:pp.first=i16(a,0);break;
      case 0xA413:pp.before=u16(a,0);break;case 0xA414:pp.after=u16(a,0);break;
      case 0x6412:pp.line=[i16(a,0),i16(a,2)];break;
      case 0x2416:pp.inTable=!!a[0];break;case 0x2417:pp.ttp=!!a[0];break;
      case 0x6649:pp.itap=i32(a,0);break;case 0x244B:pp.innerCell=!!a[0];break;case 0x244C:pp.innerTtp=!!a[0];break;
      case 0x460B:pp.ilfo=i16(a,0);break;case 0x260A:pp.ilvl=a[0];break;
      case 0x2407:pp.pageBreak=!!a[0];break;case 0x2405:pp.keepLines=!!a[0];break;case 0x2406:pp.keepNext=!!a[0];break;
      case 0x442D:pp.shd=shd80(a,0);break;case 0xC64D:pp.shd=shd(a,1);break;
      case 0x6424:pp.bTop=brc80(a,0);break;case 0x6425:pp.bLeft=brc80(a,0);break;case 0x6426:pp.bBottom=brc80(a,0);break;case 0x6427:pp.bRight=brc80(a,0);break;
      case 0xC64E:pp.bTop=brc(a,1);break;case 0xC64F:pp.bLeft=brc(a,1);break;case 0xC650:pp.bBottom=brc(a,1);break;case 0xC651:pp.bRight=brc(a,1);break;
      case 0x6646:{const at=u32(a,0);if(at+2<=D.length){const cb=u16(D,at);applyPap(pp,D.subarray(at+2,at+2+cb));}break;}
      default:if((op>>10&7)===5)(pp.tap=pp.tap||[]).push([op,a]); // a table sprm
    }
    return pp;
  }
  // ---- styles
  const styles=[];let defF=null,defFo=null;
  {const [fc,lcb]=fcl(1);if(lcb){const cbStshi=u16(T,fc),cstd=u16(T,fc+2),cbBase=u16(T,fc+4);
    if(cbStshi>=18){defF=u16(T,fc+14);defFo=u16(T,fc+18);}
    let q=fc+2+cbStshi;
    for(let i=0;i<cstd&&q<fc+lcb;i++){const cb=u16(T,q);q+=2;
      if(cb){const s=q,stk=u16(T,s+2)&0xF,base=u16(T,s+2)>>4,cupx=u16(T,s+4)&0xF;
        let r=s+cbBase;const cch=u16(T,r);r+=2+cch*2+2;const upx=[];
        for(let k=0;k<cupx;k++){if((r-s)&1)r++;const n=u16(T,r);upx.push(T.subarray(r+2,r+2+n));r+=2+n;}
        styles[i]={stk,base,upx};}
      q+=cb;}}}
  const sCache=new Map();
  function styleOf(istd,depth){
    if(sCache.has(istd))return sCache.get(istd);
    const s=styles[istd];let r={pap:{},chp:{}};
    if(s&&(depth||0)<12){
      const b=s.base!==0xFFF&&s.base!==istd?styleOf(s.base,(depth||0)+1):{pap:{},chp:{}};
      r={pap:{...b.pap},chp:{...b.chp}};
      if(s.stk===1){if(s.upx[0]&&s.upx[0].length>=2)applyPap(r.pap,s.upx[0].subarray(2));if(s.upx[1])applyChp(r.chp,s.upx[1],b.chp);}
      else if(s.stk===2&&s.upx[0])applyChp(r.chp,s.upx[0],b.chp);
      delete r.pap.tap;
    }
    sCache.set(istd,r);return r;
  }
  styleChp=istd=>styleOf(istd).chp;
  // ---- lists
  const lists={},lfos=[];
  {const [fc,lcb]=fcl(73);if(lcb){const n=i16(T,fc);let q=fc+2;const lst=[];
    for(let i=0;i<n;i++){lst.push({id:i32(T,q),simple:!!(T[q+26]&1)});q+=28;}
    for(const L of lst){L.lvls=[];for(let k=0;k<(L.simple?1:9);k++){
      const start=i32(T,q),nfc=T[q+4],jc=T[q+5]&3,follow=T[q+15],cbC=T[q+24],cbP=T[q+25];q+=28;
      const pp=applyPap({},T.subarray(q,q+cbP));q+=cbP;const ch=applyChp({},T.subarray(q,q+cbC),{});q+=cbC;
      const cch=u16(T,q);const xst=dec16.decode(T.subarray(q+2,q+2+cch*2));q+=2+cch*2;
      L.lvls.push({start,nfc,jc,follow,pp,ch,xst});}
      lists[L.id]=L;}}}
  {const [fc,lcb]=fcl(74);if(lcb){const n=i32(T,fc);for(let i=0;i<n;i++)lfos.push(i32(T,fc+4+i*16));}}
  // ---- pictures: the shared picture store, and drawn shapes (spid → picture)
  const blips=[],shapePib={};
  function blipAt(a,o){ // an OfficeArt picture record → {ext,bytes}
    const vi=u16(a,o),type=u16(a,o+2),len=u32(a,o+4),inst=vi>>4,d=o+8;
    const two=inst&1;let ext=null,skip=0;
    if(type===0xF01E){ext='png';skip=17+(two?16:0);}else if(type===0xF01D||type===0xF02A){ext='jpeg';skip=17+(two?16:0);}
    else if(type===0xF01F){ext='bmp';skip=17+(two?16:0);}
    else if(type===0xF01A||type===0xF01B){ // EMF / WMF: often a photo wrapped in a drawing; the photo is taken out later (docImages)
      const m=d+16+(two?16:0);return{ext:type===0xF01A?'emf':'wmf',meta:true,deflated:a[m+32]===0,bytes:a.slice(m+34,d+len)};}
    else return null;
    let b=a.slice(d+skip,d+len);
    if(ext==='bmp')b=docDibToBmp(b);
    return{ext,bytes:b};
  }
  function fbse(a,o,len){ // o = start of the FBSE data
    const cbName=a[o+33];
    if(len>36+cbName)return blipAt(a,o+36+cbName);
    const fo=u32(a,o+28);if(fo&&fo!==0xFFFFFFFF&&fo+8<=W.length)return blipAt(W,fo);
    return null;
  }
  function walk(a,s,e,cb){let q=s;while(q+8<=e){const vi=u16(a,q),type=u16(a,q+2),len=u32(a,q+4);if(len>e-q-8)break;cb(type,vi,q+8,len);if((vi&0xF)===0xF)walk(a,q+8,q+8+len,cb);q+=8+len;}}
  // OfficeArtContent: the picture store (one container), then for each story a one-byte label and its drawing
  {const [fc,lcb]=fcl(50);if(lcb>8){const end=fc+lcb;let curSp=null;
    const see=(type,vi,d,len)=>{
      if(type===0xF007)blips.push(fbse(T,d,len));
      else if(type===0xF00A)curSp=u32(T,d);
      else if(type===0xF00B&&curSp!=null){const n=vi>>4;for(let k=0;k<n;k++){if((u16(T,d+k*6)&0x3FFF)===0x0104)shapePib[curSp]=u32(T,d+k*6+2);}}
    };
    let q=fc;const len0=u32(T,q+4);walk(T,q,Math.min(end,q+8+len0),see);q+=8+len0;
    while(q+9<=end){q+=1;const len=u32(T,q+4);if(u16(T,q+2)!==0xF002||len>end-q-8)break;walk(T,q,q+8+len,see);q+=8+len;}
  }}
  function spa(i,base){ // drawn shapes of a story: cp → {pib,w,h}
    const [fc,lcb]=fcl(i);const out=new Map();if(!lcb)return out;const n=(lcb-4)/30;
    for(let k=0;k<n;k++){const cp=tv.getInt32(fc+k*4,true),o=fc+(n+1)*4+k*26;const id=tv.getInt32(o,true);
      out.set(base+cp,{pib:shapePib[id],w:tv.getInt32(o+12,true)-tv.getInt32(o+4,true),h:tv.getInt32(o+16,true)-tv.getInt32(o+8,true)});}
    return out;
  }
  const hdrBase=ccpText+ccpFtn;
  const shapesMain=spa(40,0),shapesHdr=spa(41,hdrBase);
  function inlinePic(at){ // a picture inside the text: PICF in the Data stream, then its picture record
    if(at+68>D.length)return null;const lcb=u32(D,at),cbH=u16(D,at+4),mm=u16(D,at+6);
    const w=u16(D,at+28)*(u16(D,at+32)||1000)/1000,h=u16(D,at+30)*(u16(D,at+34)||1000)/1000;
    let s=at+cbH;if(mm===0x66)s+=1+D[s];let img=null;
    walk(D,s,Math.min(at+lcb,D.length),(type,vi,d,len)=>{if(img)return;if(type===0xF007)img=fbse(D,d,len);else if(type>=0xF018&&type<=0xF117&&type!==0xF00B&&type!==0xF00A&&type!==0xF122)img=blipAt(D,d-8);});
    return img?{...img,w,h}:null;
  }
  // ---- read a stretch of text into paragraphs of runs
  let fldN=0;
  function readRange(a,b,shapes){
    const paras=[];let runs=[],cur=null,field=[];
    const flush=()=>{if(cur){runs.push(cur);cur=null;}};
    for(let cp=a;cp<b;cp++){
      const [ch,fc,pc]=charAt(cp);if(!pc)break;
      const cr=byFc(chpRuns,fc);const g=cr?cr[2]:null;
      // fields: the code is dropped and the shown result kept; page numbers stay live fields (PAGE / NUMPAGES)
      if(ch==='\x13'){field.push({s:0,code:''});flush();continue;}
      if(ch==='\x14'){const f=field[field.length-1];if(f){f.s=1;const m=/^\s*(PAGE|NUMPAGES|SECTIONPAGES)\b/i.exec(f.code);if(m)f.fld={instr:m[1].toUpperCase(),id:++fldN};}flush();continue;}
      if(ch==='\x15'){field.pop();flush();continue;}
      const inCode=field.some(v=>v.s===0);
      if(inCode&&field.length&&field[field.length-1].s===0)field[field.length-1].code+=ch;
      const fld=(field.length&&field[field.length-1].fld)||null;
      if(ch==='\r'||ch==='\x07'){flush();if(!inCode||ch==='\x07'){const pr=byFc(papRuns,fc);paras.push({runs,end:ch,papG:pr?pr[2]:null,istd:pr?pr[3]:0,pcG:pc.g,markG:g});runs=[];}continue;}
      if(inCode)continue;
      if(ch==='\x01'||ch==='\x08'){flush();runs.push({ch,g,pcG:pc.g,cp});continue;}
      if('\x02\x03\x04\x05\x1f'.includes(ch))continue;
      const t=ch==='\x1e'?'-':ch;
      if(cur&&cur.g===g&&cur.pcG===pc.g&&cur.fld===fld&&!'\t\x0b\x0c'.includes(t)&&!'\t\x0b\x0c'.includes(cur.t.slice(-1)))cur.t+=t;
      else{flush();cur={t,g,pcG:pc.g,fld};}
    }
    flush();if(runs.length)paras.push({runs,end:'\r',papG:null,istd:0});
    return paras.map(P=>finishPara(P,shapes));
  }
  function finishPara(P,shapes){
    const st=styleOf(P.istd);const pp=applyPap({...st.pap},P.papG||new Uint8Array(0));if(P.pcG)applyPap(pp,P.pcG);
    pp.istd=P.istd;
    const runs=[];
    for(const r of P.runs){
      const c=applyChp({...st.chp},r.g||new Uint8Array(0),st.chp);if(r.pcG)applyChp(c,r.pcG,st.chp);
      if(c.del||c.vanish)continue;
      if(r.ch==='\x01'){const im=c.pic!=null&&!c.obj?inlinePic(c.pic):null;if(im)runs.push({img:im,c});continue;}
      if(r.ch==='\x08'){const s=shapes&&shapes.get(r.cp);if(s&&s.pib&&blips[s.pib-1])runs.push({img:{...blips[s.pib-1],w:s.w,h:s.h},c});continue;}
      let t=r.t;
      if(c.sym){const m=DOC_SYM[c.sym.ch&0xFF];t=m||String.fromCharCode(c.sym.ch);}
      else{const fn=(fonts[c.f0!=null?c.f0:defF]||'');if(/symbol|wingdings/i.test(fn))t=[...t].map(x=>{const k=x.charCodeAt(0);return DOC_SYM[k&0xFF]||(k>=0xF000?String.fromCharCode(k&0xFF):x);}).join('');}
      runs.push({t,c,fld:r.fld});
    }
    const mc=applyChp({...st.chp},P.markG||new Uint8Array(0),st.chp);
    return{runs,pp,end:P.end,mark:mc};
  }
  // ---- section: page size, margins, header distance
  const sec={w:12240,h:15840,l:1800,r:1800,t:1440,b:1440,hd:720,fd:720,title:false,bidi:false};
  {const [fc,lcb]=fcl(6);if(lcb){const n=(lcb-4)/16;const o=fc+(n+1)*4+(n-1)*12;const at=tv.getUint32(o+2,true);
    if(at!==0xFFFFFFFF&&at+2<=W.length){const cb=wv.getInt16(at,true);
      for(const [op,a] of sprms(W.subarray(at+2,at+2+cb)))switch(op){
        case 0xB01F:sec.w=u16(a,0);break;case 0xB020:sec.h=u16(a,0);break;case 0xB021:sec.l=u16(a,0);break;case 0xB022:sec.r=u16(a,0);break;
        case 0x9023:sec.t=Math.abs(i16(a,0));break;case 0x9024:sec.b=Math.abs(i16(a,0));break;case 0xB017:sec.hd=u16(a,0);break;case 0xB018:sec.fd=u16(a,0);break;
        case 0x300A:sec.title=!!a[0];break;case 0x3228:sec.bidi=!!a[0];break;case 0x301D:sec.land=a[0]===2;break;}}}}
  // ---- headers and footers (the stories of the first section; the last paragraph of each is a closing mark)
  const hf={};
  {const [fc,lcb]=fcl(11);if(lcb&&ccpHdd>0){const n=lcb/4-1;const cpAt=i=>tv.getInt32(fc+i*4,true);
    const names=['evenH','oddH','evenF','oddF','firstH','firstF'];
    for(let k=0;k<6;k++){const i=6+k;if(i+1>n)break;const a=cpAt(i),b=cpAt(i+1);if(b>a){const ps=readRange(hdrBase+a,hdrBase+b,shapesHdr);if(ps.length>1&&!ps[ps.length-1].runs.length)ps.pop();hf[names[k]]=ps;}}}}
  const body=readRange(0,ccpText,shapesMain);
  return{body,hf,sec,fonts,lists,lfos,defF,defFo,fontOf:i=>fonts[i]};
}
// paragraphs → blocks: paragraphs and tables (a row ends on a mark that carries the row's own formatting)
function docModelBlocks(paras){
  const blocks=[];let tbl=null,row=[],cell=[];
  const endTable=()=>{if(tbl){if(cell.length||row.length){if(cell.length)row.push(cell);tbl.rows.push({cells:row,tap:null});}blocks.push(tbl);tbl=null;row=[];cell=[];}};
  for(const p of paras){
    const pp=p.pp;
    if(pp.inTable||(pp.itap||0)>0){
      if(!tbl)tbl={t:'table',rows:[]};
      if(pp.ttp&&!(pp.itap>1)){if(cell.length){row.push(cell);cell=[];}tbl.rows.push({cells:row,tap:pp.tap||[],pp});row=[];continue;}
      if(pp.innerTtp)continue;
      cell.push(p);
      if(p.end==='\x07'&&!(pp.itap>1)){row.push(cell);cell=[];}
      continue;
    }
    endTable();blocks.push({t:'p',p});
  }
  endTable();
  return blocks;
}
function docModelXml(M,ctx){
  const font=i=>i==null?null:M.fontOf(i)||null;
  const bXml=(k,b)=>b?`<w:${k} w:val="${b.val}" w:sz="${b.sz}" w:space="0" w:color="${b.color}"/>`:b===null?`<w:${k} w:val="nil"/>`:'';
  function rPr(c,t){
    const o=[];let f0=font(c.f0!=null?c.f0:M.defF);const fb=font(c.fBi!=null?c.fBi:(c.f2!=null?c.f2:null))||f0;
    const isHeb=t&&HEB.test(t);
    // Hebrew is drawn with the right-to-left ("cs") font, size and bold. Viewers that read only the Latin ones
    // (the page view of סימן) get the same values, so a run that is Hebrew only looks the same everywhere
    const he=isHeb&&!/[A-Za-z]/.test(t);
    if(he&&fb)f0=fb;
    if(f0||fb)o.push(`<w:rFonts${f0?` w:ascii="${esc(f0)}" w:hAnsi="${esc(f0)}"`:''}${fb?` w:cs="${esc(fb)}"`:''}/>`);
    const bC=c.bCs!=null?c.bCs:c.b,iC=c.iCs!=null?c.iCs:c.i;
    if(he?bC:c.b)o.push('<w:b/>');if(bC)o.push('<w:bCs/>');
    if(he?iC:c.i)o.push('<w:i/>');if(iC)o.push('<w:iCs/>');
    if(c.caps)o.push('<w:caps/>');if(c.smallCaps)o.push('<w:smallCaps/>');if(c.strike)o.push('<w:strike/>');if(c.dstrike)o.push('<w:dstrike/>');
    if(c.color)o.push(`<w:color w:val="${c.color}"/>`);
    const szc=c.szCs||c.sz||20,sz=he?szc:(c.sz||20);o.push(`<w:sz w:val="${sz}"/><w:szCs w:val="${szc}"/>`);
    if(c.hl)o.push(`<w:highlight w:val="${c.hl}"/>`);
    if(c.u)o.push(`<w:u w:val="${c.u}"/>`);
    if(c.shd)o.push(`<w:shd w:val="clear" w:color="auto" w:fill="${c.shd}"/>`);
    if(c.va)o.push(`<w:vertAlign w:val="${c.va}"/>`);
    if(c.rtl||isHeb)o.push('<w:rtl/>');
    return o.length?`<w:rPr>${o.join('')}</w:rPr>`:'';
  }
  const fldSeen=new Set();
  function run(r){
    if(r.img)return ctx.img(r.img);
    // a page field in its long form: Word and the editor keep it live, and a viewer that does not count pages shows the kept number
    if(r.fld){if(fldSeen.has(r.fld.id))return'';fldSeen.add(r.fld.id);const rp=rPr(r.c,r.t);
      return`<w:r>${rp}<w:fldChar w:fldCharType="begin"/></w:r><w:r>${rp}<w:instrText xml:space="preserve"> ${r.fld.instr} </w:instrText></w:r><w:r>${rp}<w:fldChar w:fldCharType="separate"/></w:r><w:r>${rp}<w:t xml:space="preserve">${esc(r.t)}</w:t></w:r><w:r>${rp}<w:fldChar w:fldCharType="end"/></w:r>`;}
    return String(r.t).split(/(\t|\x0b|\x0c)/).map(s=>s==='\t'?'<w:r><w:tab/></w:r>':s==='\x0b'?'<w:r><w:br/></w:r>':s==='\x0c'?'<w:r><w:br w:type="page"/></w:r>':s?`<w:r>${rPr(r.c,s)}<w:t xml:space="preserve">${esc(s)}</w:t></w:r>`:'').join('');
  }
  function pPr(p){
    const pp=p.pp,o=[];const txt=p.runs.map(r=>r.t||'').join('');
    const bidi=pp.bidi!=null?pp.bidi:HEB.test(txt);
    if(pp.keepNext)o.push('<w:keepNext/>');if(pp.keepLines)o.push('<w:keepLines/>');if(pp.pageBreak)o.push('<w:pageBreakBefore/>');
    if(pp.ilfo>0&&pp.ilfo<=M.lfos.length&&ctx.numOk(pp.ilfo))o.push(`<w:numPr><w:ilvl w:val="${pp.ilvl||0}"/><w:numId w:val="${pp.ilfo}"/></w:numPr>`);
    if(pp.bTop||pp.bLeft||pp.bBottom||pp.bRight)o.push(`<w:pBdr>${bXml('top',pp.bTop)}${bXml('left',pp.bLeft)}${bXml('bottom',pp.bBottom)}${bXml('right',pp.bRight)}</w:pBdr>`);
    if(pp.shd)o.push(`<w:shd w:val="clear" w:color="auto" w:fill="${pp.shd}"/>`);
    if(bidi)o.push('<w:bidi/>');
    let sp='';if(pp.before!=null)sp+=` w:before="${pp.before}"`;if(pp.after!=null)sp+=` w:after="${pp.after}"`;
    if(pp.line){const[v,m]=pp.line;sp+=m?` w:line="${v}" w:lineRule="auto"`:v<0?` w:line="${-v}" w:lineRule="exact"`:` w:line="${v}" w:lineRule="atLeast"`;}
    if(sp)o.push(`<w:spacing${sp}/>`);
    // the newer indents are logical (start / end); the older "80" ones are physical, so in a right-to-left paragraph right = start
    const s=pp.indS!=null?pp.indS:(bidi?pp.r80:pp.l80),e=pp.indE!=null?pp.indE:(bidi?pp.l80:pp.r80);
    if(s!=null||e!=null||pp.first!=null)o.push(`<w:ind${s!=null?` w:left="${s}"`:''}${e!=null?` w:right="${e}"`:''}${pp.first!=null?(pp.first<0?` w:hanging="${-pp.first}"`:` w:firstLine="${pp.first}"`):''}/>`);
    // alignment: the logical value (start / end) is written as left / right, which Word reads logically in a right-to-left paragraph
    let jc=null;if(pp.jcL!=null)jc=pp.jcL;else if(pp.jc80!=null)jc=bidi&&(pp.jc80===0||pp.jc80===2)?2-pp.jc80:pp.jc80;
    if(jc!=null)o.push(`<w:jc w:val="${['left','center','right','both','distribute'][jc]||'left'}"/>`);
    const mr=rPr(p.mark,txt);if(mr)o.push(mr);
    return`<w:pPr>${o.join('')}</w:pPr>`;
  }
  const para=p=>`<w:p>${pPr(p)}${p.runs.map(run).join('')}</w:p>`;
  function table(tb){
    // every row's cell edges; the table's grid is all the edges together
    const rows=tb.rows.map(r=>{const t={centers:null,tcs:[],shd:[],bTbl:null,setBrc:[],h:null,bidi:false,header:false,cant:false,gap:null,vm:{},va:{}};
      for(const [op,a] of (r.tap||[])){
        switch(op){
          case 0xD608:{const n=a[2];t.centers=[];for(let k=0;k<=n;k++)t.centers.push((a[3+k*2]|a[4+k*2]<<8)<<16>>16);
            let o=3+(n+1)*2;for(let k=0;k<n&&o+20<=a.length;k++,o+=20){const gf=a[o]|a[o+1]<<8;t.tcs[k]={gf,top:docBrc80(a,o+4),left:docBrc80(a,o+8),bottom:docBrc80(a,o+12),right:docBrc80(a,o+16)};}break;}
          case 0xD609:for(let k=0;k*2+1<a[0];k++)t.shd[k]=docShd80(a,1+k*2);break;
          case 0xD612:for(let k=0;k*10+9<a[0];k++)t.shd[k]=docShd(a,1+k*10);break;
          case 0xD616:for(let k=0;k*10+9<a[0];k++)t.shd[22+k]=docShd(a,1+k*10);break;
          case 0xD60C:for(let k=0;k*10+9<a[0];k++)t.shd[44+k]=docShd(a,1+k*10);break;
          case 0xD605:t.bTbl=[0,1,2,3,4,5].map(k=>docBrc80(a,1+k*4));break;
          case 0xD613:t.bTbl=[0,1,2,3,4,5].map(k=>docBrc(a,1+k*8));break;
          case 0xD620:t.setBrc.push([a[1],a[2],a[3],docBrc80(a,4)]);break;
          case 0xD62F:t.setBrc.push([a[1],a[2],a[3],docBrc(a,4)]);break;
          case 0x9407:t.h=(a[0]|a[1]<<8)<<16>>16;break;
          case 0x560B:t.bidi=!!(a[0]|a[1]);break;case 0x3404:t.header=!!a[0];break;case 0x3403:t.cant=!!a[0];break;
          case 0x9602:t.gap=(a[0]|a[1]<<8)<<16>>16;break;
          case 0x548A:t.jc=a[0];break; // the table placed at the start / centre / end (the older 0x5400 is not what Word follows)
          case 0xD634:if(a[3]&2||a[3]&8)t.pad=(a[5]|a[6]<<8);break; // cell padding for the whole table (left / right)
          case 0xD62B:t.vm[a[1]]=a[2];break;
          case 0xD62C:for(let k=a[1];k<a[2];k++)t.va[k]=a[3];break;
        }
      }
      return{...r,t};});
    const edges=new Set();for(const r of rows)if(r.t.centers)r.t.centers.forEach(x=>edges.add(x));
    let grid=[...edges].sort((x,y)=>x-y);
    if(grid.length<2){const n=Math.max(1,...rows.map(r=>r.cells.length));grid=[];for(let k=0;k<=n;k++)grid.push(Math.round(k*9000/n));}
    // right-to-left only when the table itself says so: a Hebrew table can be laid out left to right
    const bidi=rows.some(r=>r.t.bidi);
    // the table starts where its first edge is, plus the half-gap when the file has one; the text inside keeps its padding
    const gap=rows.find(r=>r.t.gap!=null),padR=rows.find(r=>r.t.pad!=null);const g=gap?gap.t.gap:0,cm=padR?padR.t.pad:gap?gap.t.gap:108;
    const tjc=(rows.find(r=>r.t.jc!=null)||{t:{}}).t.jc;
    const tb0=(rows.find(r=>r.t.bTbl)||{t:{}}).t.bTbl;
    let x=`<w:tbl><w:tblPr>${bidi?'<w:bidiVisual/>':''}<w:tblW w:w="${grid[grid.length-1]-grid[0]}" w:type="dxa"/>${tjc===1?'<w:jc w:val="center"/>':tjc===2?'<w:jc w:val="right"/>':`<w:tblInd w:w="${grid[0]+g}" w:type="dxa"/>`}`;
    if(tb0)x+=`<w:tblBorders>${['top','left','bottom','right','insideH','insideV'].map((k,i)=>bXml(k,tb0[i])).join('')}</w:tblBorders>`;
    x+=`<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="${cm}" w:type="dxa"/><w:right w:w="${cm}" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>`;
    for(let k=0;k+1<grid.length;k++)x+=`<w:gridCol w:w="${grid[k+1]-grid[k]}"/>`;
    x+='</w:tblGrid>';
    for(const r of rows){
      const t=r.t,cs=t.centers||grid.slice(0,r.cells.length+1);
      x+=`<w:tr><w:trPr>${t.cant?'<w:cantSplit/>':''}${t.h?`<w:trHeight w:val="${Math.abs(t.h)}" w:hRule="${t.h<0?'exact':'atLeast'}"/>`:''}${t.header?'<w:tblHeader/>':''}</w:trPr>`;
      const before=grid.indexOf(cs[0]);
      if(before>0)x=x.replace(/<w:trPr>([^]*?)<\/w:trPr>$/,(m0,m1)=>`<w:trPr><w:gridBefore w:val="${before}"/>${m1}</w:trPr>`);
      for(let k=0;k<r.cells.length;k++){
        const tc=t.tcs[k]||{gf:0};if(tc.gf&2)continue; // merged into the cell before it
        let k2=k+1;while(k2<r.cells.length&&t.tcs[k2]&&(t.tcs[k2].gf&2))k2++;
        const a0=cs[k],a1=cs[Math.min(k2,cs.length-1)];
        const span=Math.max(1,grid.indexOf(a1)-grid.indexOf(a0));
        const vm=t.vm[k]!=null?t.vm[k]:(tc.gf&0x40)?3:(tc.gf&0x20)?1:0;
        const va=t.va[k]!=null?t.va[k]:(tc.gf>>7)&3;
        const bs={top:tc.top,left:tc.left,bottom:tc.bottom,right:tc.right};
        for(const [f,l,which,b] of t.setBrc)if(k>=f&&k<l){if(which&1)bs.top=b;if(which&2)bs.left=b;if(which&4)bs.bottom=b;if(which&8)bs.right=b;}
        const bx=['top','left','bottom','right'].map(s=>bs[s]?bXml(s,bs[s]):'').join('');
        const fill=t.shd[k];
        x+=`<w:tc><w:tcPr><w:tcW w:w="${Math.max(0,a1-a0)}" w:type="dxa"/>${span>1?`<w:gridSpan w:val="${span}"/>`:''}${vm===3?'<w:vMerge w:val="restart"/>':vm===1?'<w:vMerge/>':''}${bx?`<w:tcBorders>${bx}</w:tcBorders>`:''}${fill?`<w:shd w:val="clear" w:color="auto" w:fill="${fill}"/>`:''}${va?`<w:vAlign w:val="${va===1?'center':'bottom'}"/>`:''}</w:tcPr>`;
        const ps=r.cells.slice(k,k2).flat();
        x+=ps.length?ps.map(para).join(''):'<w:p/>';
        x+='</w:tc>';
      }
      x+='</w:tr>';
    }
    return x+'</w:tbl>';
  }
  return function blocksXml(blocks){let out='';for(let i=0;i<blocks.length;i++){const b=blocks[i];out+=b.t==='p'?para(b.p):table(b);if(b.t==='table'&&(i===blocks.length-1||blocks[i+1].t==='table'))out+='<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/></w:pPr></w:p>';}return out;};
}
// a bare DIB (bitmap header + colours + pixels) → a .bmp file
function docDibToBmp(b){
  const u=o=>(b[o]|b[o+1]<<8|b[o+2]<<16|b[o+3]<<24)>>>0;
  const hs=u(0),bc=b[14]|b[15]<<8,comp=u(16),used=u(32);
  const pal=(used||(bc<=8?1<<bc:0))*4+(comp===3&&hs===40?12:0);
  const out=new Uint8Array(14+b.length);out[0]=0x42;out[1]=0x4D;const v=new DataView(out.buffer);v.setUint32(2,out.length,true);v.setUint32(10,14+hs+pal,true);out.set(b,14);return out;
}
// the photo inside an EMF / WMF drawing (Word 97-2003 keeps pasted photos this way). null = a real drawing, left out
async function docMetaImage(im){
  let e=im.bytes;
  if(im.deflated){try{e=new Uint8Array(await new Response(new Blob([e]).stream().pipeThrough(new DecompressionStream('deflate'))).arrayBuffer());}catch(_){return null;}}
  const u32=o=>(e[o]|e[o+1]<<8|e[o+2]<<16|e[o+3]<<24)>>>0,u16=o=>e[o]|e[o+1]<<8;
  const fromBmi=(bmi,bits)=>{const c=(bmi[16]|bmi[17]<<8|bmi[18]<<16|bmi[19]<<24)>>>0;
    if(c===4)return{ext:'jpeg',bytes:bits.slice()};if(c===5)return{ext:'png',bytes:bits.slice()};
    const dib=new Uint8Array(bmi.length+bits.length);dib.set(bmi);dib.set(bits,bmi.length);return{ext:'bmp',bytes:docDibToBmp(dib)};};
  let best=null;
  if(im.ext==='emf'){
    for(let q=0;q+8<=e.length;){const t=u32(q),sz=u32(q+4);if(sz<8||q+sz>e.length)break;
      if(t===0x51||t===0x50){const ob=u32(q+48),cb=u32(q+52),obt=u32(q+56),cbt=u32(q+60);
        if(cb&&cbt&&q+obt+cbt<=e.length&&(!best||cbt>best.n))best={n:cbt,img:fromBmi(e.subarray(q+ob,q+ob+cb),e.subarray(q+obt,q+obt+cbt))};}
      q+=sz;}
  }else{
    let q=u32(0)===0x9AC6CDD7?22:0;q+=u16(q+2)*2; // the header
    while(q+6<=e.length){const sz=u32(q)*2,fn=u16(q+4);if(sz<6||q+sz>e.length)break;
      const at=fn===0x0F43?q+28:fn===0x0B41?q+26:fn===0x0940?q+22:0;
      if(at&&at+40<q+sz&&(!best||sz>best.n)){const dib=e.subarray(at,q+sz),c=u32(at+16);
        best={n:sz,img:c===4||c===5?null:{ext:'bmp',bytes:docDibToBmp(dib)}};}
      q+=sz;}
  }
  return best&&best.img;
}
// pictures made ready for the new file: drawings → their photo, big bitmaps → JPEG (a .bmp photo is many MB)
async function docImages(M){
  const all=[];const scan=ps=>{for(const p of ps||[])for(const r of p.runs)if(r.img)all.push(r);};
  scan(M.body);for(const k in M.hf)scan(M.hf[k]);
  for(const r of all){
    let im=r.img;
    if(im.meta){const got=await docMetaImage(im);if(!got){r.img=null;r.t='';continue;}im={...im,...got,meta:false};}
    if(im.ext==='bmp'&&typeof createImageBitmap==='function'){
      try{const bm=await createImageBitmap(new Blob([im.bytes],{type:'image/bmp'}));const cv=document.createElement('canvas');cv.width=bm.width;cv.height=bm.height;
        const g=cv.getContext('2d');g.fillStyle='#fff';g.fillRect(0,0,cv.width,cv.height);g.drawImage(bm,0,0);
        const big=im.bytes.length>300000;const bl=await new Promise(res=>cv.toBlob(res,big?'image/jpeg':'image/png',.9));
        if(bl)im={...im,ext:big?'jpeg':'png',bytes:new Uint8Array(await bl.arrayBuffer())};}catch(_){}
    }
    r.img=im;
  }
}
// border / shading readers shared with the table code (same rules as inside docRead)
function docCv(a,o){return a[o+3]===0xFF?null:[a[o],a[o+1],a[o+2]].map(hex2).join('');}
function docBrc80(a,o){if(o+4>a.length)return undefined;const v=(a[o]|a[o+1]<<8|a[o+2]<<16|a[o+3]<<24)>>>0;if(v===0xFFFFFFFF)return undefined;if(!v||!a[o+1]||a[o+1]===0xFF)return undefined;return{val:DOC_BRC[a[o+1]]||'single',sz:Math.max(2,a[o]),color:DOC_ICO[a[o+2]]||'auto'};}
function docBrc(a,o){if(o+8>a.length)return undefined;if(a[o+4]===0xFF&&a[o+5]===0xFF)return undefined;if(!a[o+5]||a[o+5]===0xFF)return null;return{val:DOC_BRC[a[o+5]]||'single',sz:Math.max(2,a[o+4]),color:docCv(a,o)||'auto'};}
function docFill(fore,back,ipat){if(ipat===0xFFFF)return null;const pc=DOC_PCT[ipat]!=null?DOC_PCT[ipat]:.25;if(!pc)return back||null;if(pc===1)return fore||'000000';
  const F=fore?[0,2,4].map(i=>parseInt(fore.substr(i,2),16)):[0,0,0],B=back?[0,2,4].map(i=>parseInt(back.substr(i,2),16)):[255,255,255];return F.map((x,i)=>hex2(Math.round(x*pc+B[i]*(1-pc)))).join('');}
function docShd(a,o){if(o+10>a.length)return null;return docFill(docCv(a,o),docCv(a,o+4),a[o+8]|a[o+9]<<8);}
function docShd80(a,o){if(o+2>a.length)return null;const w=a[o]|a[o+1]<<8;if(w===0xFFFF)return null;return docFill(DOC_ICO[w&0x1F],DOC_ICO[(w>>5)&0x1F],w>>10);}
async function docToDocxFormatted(bytes){
  const M=docRead(bytes);
  await docImages(M);
  if(!M.body.some(p=>p.runs.some(r=>(r.t||'').trim()||r.img)))throw new Error('empty');
  const zip=new JSZip();const media=[];const rels={doc:[],h:[],f:[]};let imgN=0,docPr=1;
  const mkCtx=key=>({
    img:im=>{const name=`image${++imgN}.${im.ext==='jpeg'?'jpeg':im.ext}`;zip.file('word/media/'+name,im.bytes);media.push(im.ext);
      const rid='rIdImg'+imgN;rels[key].push(`<Relationship Id="${rid}" Type="${REL_IMG}" Target="media/${name}"/>`);
      const cx=Math.max(1,Math.round((im.w||1440)*635)),cy=Math.max(1,Math.round((im.h||1440)*635)),id=docPr++;
      return`<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${id}" name="Picture ${id}"/><a:graphic xmlns:a="${NS.a}"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="${NS.pic}"><pic:nvPicPr><pic:cNvPr id="${id}" name="${name}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;},
    numOk:ilfo=>!!numDefs[ilfo]
  });
  // numbering: one abstract list per list, one number per list use
  const numDefs={};let abs='',nums='';const absDone={};
  M.lfos.forEach((lsid,i)=>{const L=M.lists[lsid];if(!L)return;numDefs[i+1]=1;
    if(!absDone[lsid]){absDone[lsid]=Object.keys(absDone).length+1;const aid=absDone[lsid];
      abs+=`<w:abstractNum w:abstractNumId="${aid}"><w:multiLevelType w:val="${L.simple?'singleLevel':'multilevel'}"/>`;
      L.lvls.forEach((v,k)=>{
        let txt='';for(const ch of v.xst){const c=ch.charCodeAt(0);txt+=c<9?'%'+(c+1):ch;}
        const fnt=M.fontOf(v.ch.f0!=null?v.ch.f0:-1)||'';
        if(v.nfc===23){txt=[...txt].map(x=>{const c=x.charCodeAt(0);return DOC_SYM[c&0xFF]&&(c>=0xF000||/symbol|wingdings/i.test(fnt))?DOC_SYM[c&0xFF]:x;}).join('');}
        const s=v.pp.indS!=null?v.pp.indS:v.pp.l80,f=v.pp.first;
        const rp=[];if(v.nfc===23)rp.push('<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>');if(v.ch.b)rp.push('<w:b/>');if(v.ch.color)rp.push(`<w:color w:val="${v.ch.color}"/>`);if(v.ch.sz)rp.push(`<w:sz w:val="${v.ch.sz}"/><w:szCs w:val="${v.ch.sz}"/>`);
        abs+=`<w:lvl w:ilvl="${k}"><w:start w:val="${v.start||1}"/><w:numFmt w:val="${DOC_NFC[v.nfc]||'decimal'}"/><w:suff w:val="${['tab','space','nothing'][v.follow]||'tab'}"/><w:lvlText w:val="${esc(txt)}"/><w:lvlJc w:val="${['left','center','right'][v.jc]||'left'}"/>${s!=null||f!=null?`<w:pPr><w:ind${s!=null?` w:left="${s}"`:''}${f!=null?(f<0?` w:hanging="${-f}"`:` w:firstLine="${f}"`):''}/></w:pPr>`:''}${rp.length?`<w:rPr>${rp.join('')}</w:rPr>`:''}</w:lvl>`;
      });
      abs+='</w:abstractNum>';}
    nums+=`<w:num w:numId="${i+1}"><w:abstractNumId w:val="${absDone[lsid]}"/></w:num>`;
  });
  const hasNum=!!nums;
  const ns=`xmlns:w="${NS.w}" xmlns:r="${NS.r}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}"`;
  const bodyXml=docModelXml(M,mkCtx('doc'))(docModelBlocks(M.body));
  const S=M.sec,heb=S.bidi||M.body.some(p=>p.pp.bidi);
  let hRef='',ct='';
  const part=(kind,key,ps,type)=>{
    if(!ps||!ps.length)return;
    const blocksXml=docModelXml(M,mkCtx(kind==='header'?'h':'f'))(docModelBlocks(ps))||'<w:p/>';
    const n=(part.n=(part.n||0)+1),file=`${kind}${n}.xml`,rid='rIdHF'+n,tag=kind==='header'?'hdr':'ftr';
    let x=blocksXml;if(!/<w:p[ >][^]*<\/w:p>\s*$/.test(x))x+='<w:p/>';
    zip.file('word/'+file,`${XML_HEAD}<w:${tag} ${ns}>${x}</w:${tag}>`);
    const own=rels[kind==='header'?'h':'f'].splice(0);
    if(own.length)zip.file(`word/_rels/${file}.rels`,`${XML_HEAD}<Relationships xmlns="${NS.rel}">${own.join('')}</Relationships>`);
    rels.doc.push(`<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}" Target="${file}"/>`);
    ct+=`<Override PartName="/word/${file}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/>`;
    hRef+=`<w:${kind}Reference w:type="${type}" r:id="${rid}"/>`;
  };
  part('header','h',M.hf.oddH,'default');if(S.title)part('header','h',M.hf.firstH,'first');
  part('footer','f',M.hf.oddF,'default');if(S.title)part('footer','f',M.hf.firstF,'first');
  if(hasNum){zip.file('word/numbering.xml',`${XML_HEAD}<w:numbering xmlns:w="${NS.w}">${abs}${nums}</w:numbering>`);
    rels.doc.push(`<Relationship Id="rIdNum" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`);
    ct+='<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>';}
  rels.doc.push(`<Relationship Id="rIdSt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`);
  const exts=[...new Set(media)].map(e=>`<Default Extension="${e}" ContentType="image/${e==='bmp'?'bmp':e}"/>`).join('');
  zip.file('[Content_Types].xml',`${XML_HEAD}<Types xmlns="${NS.ct}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${exts}<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>${ct}</Types>`);
  zip.file('_rels/.rels',`${XML_HEAD}<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file('word/_rels/document.xml.rels',`${XML_HEAD}<Relationships xmlns="${NS.rel}">${rels.doc.join('')}</Relationships>`);
  const f0=M.fontOf(M.defF)||'Times New Roman',fo=M.fontOf(M.defFo)||f0;
  zip.file('word/styles.xml',`${XML_HEAD}<w:styles xmlns:w="${NS.w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${esc(f0)}" w:hAnsi="${esc(fo)}" w:cs="${esc(f0)}" w:eastAsia="${esc(f0)}"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="en-US" w:bidi="he-IL"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>`);
  const pw=S.land&&S.w<S.h?S.h:S.w,ph=S.land&&S.w<S.h?S.w:S.h;
  zip.file('word/document.xml',`${XML_HEAD}<w:document ${ns}><w:body>${bodyXml}<w:sectPr>${hRef}<w:pgSz w:w="${pw}" w:h="${ph}"${S.land?' w:orient="landscape"':''}/><w:pgMar w:top="${S.t}" w:right="${S.r}" w:bottom="${S.b}" w:left="${S.l}" w:header="${S.hd}" w:footer="${S.fd}" w:gutter="0"/>${S.title?'<w:titlePg/>':''}${heb?'<w:bidi/>':''}</w:sectPr></w:body></w:document>`);
  const tables=(bodyXml.match(/<w:tbl>/g)||[]).length,paras=(bodyXml.match(/<w:p>/g)||[]).length;
  return{bytes:await zip.generateAsync({type:'uint8array',compression:'DEFLATE'}),paras,tables,formatted:true};
}
async function docToDocx(bytes){
  try{return await docToDocxFormatted(bytes);}
  catch(e){console.warn('doc with formatting failed, text only',e);}
  return docToDocxPlain(bytes);
}
async function docToDocxPlain(bytes){
  const blocks=docBlocks(docText(bytes));
  const heb=blocks.some(b=>b.t==='p'?HEB.test(b.text):b.rows.some(r=>r.some(c=>c.some(x=>HEB.test(x)))));
  let body='',paras=0,tables=0;
  for(const b of blocks){
    if(b.t==='p'){body+=wPara(b.text,{rtl:heb});paras++;continue;}
    tables++;const n=Math.max(...b.rows.map(r=>r.length)),w=Math.floor(5000/n);
    const bd=k=>`<w:${k} w:val="single" w:sz="4" w:space="0" w:color="808080"/>`;
    body+=`<w:tbl><w:tblPr>${heb?'<w:bidiVisual/>':''}<w:tblW w:w="5000" w:type="pct"/><w:tblBorders>${['top','left','bottom','right','insideH','insideV'].map(bd).join('')}</w:tblBorders><w:tblLayout w:type="autofit"/></w:tblPr><w:tblGrid>${'<w:gridCol/>'.repeat(n)}</w:tblGrid>`;
    for(const r of b.rows){body+='<w:tr>';for(let i=0;i<n;i++){const c=r[i]||[''];body+=`<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="pct"/></w:tcPr>${c.map(x=>wPara(x,{rtl:heb})).join('')}</w:tc>`;}body+='</w:tr>';}
    body+='</w:tbl>';
  }
  body+=wPara('',{rtl:heb});
  if(!paras&&!tables)throw new Error('לא נמצא טקסט בקובץ');
  const zip=new JSZip();
  zip.file('[Content_Types].xml',`${XML_HEAD}<Types xmlns="${NS.ct}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`);
  zip.file('_rels/.rels',`${XML_HEAD}<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
  zip.file('word/_rels/document.xml.rels',`${XML_HEAD}<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  const font=heb?'Arial':'Calibri';
  zip.file('word/styles.xml',`${XML_HEAD}<w:styles xmlns:w="${NS.w}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="Arial" w:eastAsia="${font}"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:bidi="he-IL"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`);
  zip.file('word/document.xml',`${XML_HEAD}<w:document xmlns:w="${NS.w}" xmlns:r="${NS.r}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>${heb?'<w:bidi/>':''}</w:sectPr></w:body></w:document>`);
  return{bytes:await zip.generateAsync({type:'uint8array',compression:'DEFLATE'}),paras,tables};
}

window.OFFICE_EDIT={toPdf,exportDocx,exportXlsx,docToDocx};
})();
