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
async function docxToPages(bytes,progress){
  const host=makeHost(0);
  try{
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
    const out=[],paras=[];
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
      for(const p of art.children)if(p.tagName==='P')paras.push({sec:si,top:p.getBoundingClientRect().top-R.top});
      // safe places to cut a page that runs longer than one sheet
      const cuts=[...art.querySelectorAll('p,tr,li,img,h1,h2,h3,h4,h5,h6')].map(e=>e.getBoundingClientRect().bottom-R.top).filter(b=>b>0).sort((a,b)=>a-b);
      const slices=[];
      if(total<=H*1.02)slices.push([0,H,0]);
      else{
        let s=0,first=true;
        while(s<total-padB-2){
          const room=H-(first?0:padT)-padB;
          let e=s+room;
          if(e<total-padB){const ok=cuts.filter(b=>b<=e&&b>s+room*.45);if(ok.length)e=ok[ok.length-1];}
          else e=total;
          slices.push([s,e,first?0:padT]);s=e;first=false;
        }
      }
      const scale=Math.min(2,30000/total);
      progress&&progress(`מכין את העמודים… ${si+1}/${secs.length}`);
      const full=await snap(sec,scale,W,total);
      for(const[s,e,pad]of slices){
        const c=document.createElement('canvas');c.width=Math.round(W*2);c.height=Math.round(H*2);
        const g=c.getContext('2d');g.fillStyle='#fff';g.fillRect(0,0,c.width,c.height);
        const k=full.width/W;
        g.drawImage(full,0,Math.round(s*k),full.width,Math.round((e-s)*k),0,Math.round(pad*2),c.width,Math.round((e-s)*2));
        out.push({canvas:c,w:W,h:H,src:{sec:si,top:s,pad}});n++;
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
  try{
    for(let si=0;si<wb.SheetNames.length;si++){
      const meta=(wb.Workbook&&wb.Workbook.Sheets||[])[si]||{};if(meta.Hidden)continue;
      const ws=wb.Sheets[wb.SheetNames[si]];if(!ws||!ws['!ref'])continue;
      const f=files[si]||{};let rtl=false;
      if(f.path&&zip.file(f.path)){const x=await zip.file(f.path).async('string');rtl=/<sheetView\b[^>]*rightToLeft="(1|true)"/.test(x);}
      const rg=XLSX.utils.decode_range(ws['!ref']);
      const R1=Math.min(rg.e.r,rg.s.r+4999),C1=Math.min(rg.e.c,rg.s.c+79);
      // trim empty trailing rows / columns
      let lastR=0,lastC=0;
      for(const k in ws){if(k[0]==='!')continue;const a=XLSX.utils.decode_cell(k),c=ws[k];if(c&&(c.v!==undefined&&c.v!==''||c.s&&c.s.fgColor&&c.s.patternType==='solid')){if(a.r>lastR)lastR=a.r;if(a.c>lastC)lastC=a.c;}}
      for(const m of ws['!merges']||[]){if(m.e.r>lastR)lastR=m.e.r;if(m.e.c>lastC)lastC=m.e.c;}
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
  let placed=0;
  for(const it of items){
    const pm=map.pages[it.page];if(!pm)continue;
    const yy=pm.top+(it.y-pm.pad);       // y inside the rendered section
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
  return{bytes:await zip.generateAsync({type:'uint8array',compression:'DEFLATE'}),placed,exact:same};
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
// Formatting, pictures, headers and footers of the old file are not carried over.
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
async function docToDocx(bytes){
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
