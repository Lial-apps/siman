/* סימן — עורך Word.
   מסמך Word נפתח לעריכה אמיתית, כמו ב-Word: גופנים, גודל, הדגשה, קו תחתי, צבעים, יישור,
   רשימות, סגנונות, טבלאות, תמונות, חיפוש והחלפה, פריסת עמוד ועוד.
   השמירה כותבת קובץ DOCX אמיתי: פסקאות שלא נגעו בהן נשמרות מילה במילה מהקובץ המקורי
   (כולל הערות שוליים, שדות ותמונות צפות), ורק מה ששונה נכתב מחדש.
   Needs: JSZip, docx-preview (window.docx). */
(function(){
'use strict';
const P='wx';
const NS={
  w:'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r:'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp:'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a:'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic:'http://schemas.openxmlformats.org/drawingml/2006/picture',
  mc:'http://schemas.openxmlformats.org/markup-compatibility/2006',
  rel:'http://schemas.openxmlformats.org/package/2006/relationships',
  ct:'http://schemas.openxmlformats.org/package/2006/content-types'
};
const W=NS.w;
const RT={
  doc:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  styles:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
  numbering:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering',
  settings:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',
  image:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  link:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  theme:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'
};
const OX='application/vnd.openxmlformats-officedocument.';
const DOCX_MIME=OX+'wordprocessingml.document';
const XML_HEAD='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const HEB=/[\u0590-\u05FF\uFB1D-\uFB4F\u0600-\u06FF\u0750-\u077F]/;
const HEBG=/[\u0590-\u05FF\uFB1D-\uFB4F]/g, LATG=/[A-Za-z\u00C0-\u024F]/g;
const enc=new TextEncoder(),dec=new TextDecoder('utf-8');
const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'})[c]);
const xmlText=s=>String(s).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'})[c]).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,'');
const parseXml=s=>{const d=new DOMParser().parseFromString(s,'application/xml');if(d.getElementsByTagName('parsererror').length)throw new Error('XML פגום בקובץ');return d;};
const ser=n=>new XMLSerializer().serializeToString(n);
const kids=(el,...names)=>el?[...el.childNodes].filter(n=>n.nodeType===1&&n.namespaceURI===W&&(!names.length||names.includes(n.localName))):[];
const kid=(el,name)=>kids(el,name)[0]||null;
const wAttr=(el,a)=>el?(el.getAttributeNS(W,a)??el.getAttribute('w:'+a)):null;
const wSet=(el,a,v)=>el.setAttributeNS(W,'w:'+a,String(v));
const wNew=(doc,name,attrs)=>{const e=doc.createElementNS(W,'w:'+name);if(attrs)for(const k in attrs)wSet(e,k,attrs[k]);return e;};
const onOff=el=>{if(!el)return null;const v=wAttr(el,'val');return !(v==='0'||v==='false'||v==='off');};
const twip=px=>Math.round(px*15);              // CSS px → twips (1/20 pt)
const px=tw=>(+tw||0)/15;
const EMU=9525;
const hex=c=>{c=String(c||'').trim();if(!c||c==='transparent')return null;
  const r=c.match(/rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+))?/);
  if(r){if(r[4]!==undefined&&+r[4]===0)return null;return [r[1],r[2],r[3]].map(v=>(+v).toString(16).padStart(2,'0')).join('').toUpperCase();}
  const m=c.match(/^#([0-9a-f]{6})$/i);if(m)return m[1].toUpperCase();
  const s=c.match(/^#([0-9a-f]{3})$/i);if(s)return s[1].split('').map(x=>x+x).join('').toUpperCase();return null;};
const nextFrame=()=>new Promise(r=>requestAnimationFrame(()=>setTimeout(r,0)));
const HL={FFFF00:'yellow','00FF00':'green','00FFFF':'cyan',FF00FF:'magenta','0000FF':'blue',FF0000:'red','000080':'darkBlue','008080':'darkCyan','008000':'darkGreen','800080':'darkMagenta','800000':'darkRed','808000':'darkYellow','808080':'darkGray',C0C0C0:'lightGray','000000':'black'};
const HL_REV={};for(const k in HL)HL_REV[HL[k]]=k;

/* ------------------------------------------------------------ package */
async function readZip(bytes){
  const zip=await JSZip.loadAsync(bytes);const files=new Map();
  for(const [p,f] of Object.entries(zip.files))if(!f.dir)files.set(p,await f.async('uint8array'));
  return files;
}
const txt=(files,p)=>files.has(p)?dec.decode(files.get(p)):null;
const put=(files,p,s)=>files.set(p,typeof s==='string'?enc.encode(s):s);
function relsPathOf(part){const i=part.lastIndexOf('/');return part.slice(0,i+1)+'_rels/'+part.slice(i+1)+'.rels';}
function resolveTarget(part,target){
  if(/^\//.test(target))return target.slice(1);
  const base=part.slice(0,part.lastIndexOf('/')+1).split('/').filter(Boolean);
  for(const seg of target.split('/')){if(seg==='..')base.pop();else if(seg&&seg!=='.')base.push(seg);}
  return base.join('/');
}
function mainPart(files){
  const r=txt(files,'_rels/.rels');
  if(r){const d=parseXml(r);for(const e of d.getElementsByTagNameNS(NS.rel,'Relationship'))
    if(e.getAttribute('Type')===RT.doc)return resolveTarget('',e.getAttribute('Target'));}
  return 'word/document.xml';
}
function ctEnsureDefault(files,ext,mime){
  let s=txt(files,'[Content_Types].xml');if(!s)return;
  if(new RegExp(`Extension="${ext}"`,'i').test(s))return;
  s=s.replace(/<Types[^>]*>/,m=>m+`<Default Extension="${ext}" ContentType="${mime}"/>`);put(files,'[Content_Types].xml',s);
}
function ctEnsureOverride(files,part,mime){
  let s=txt(files,'[Content_Types].xml');if(!s)return;
  if(s.includes(`PartName="/${part}"`))return;
  s=s.replace('</Types>',`<Override PartName="/${part}" ContentType="${mime}"/></Types>`);put(files,'[Content_Types].xml',s);
}

function relsOf(files,part){
  const s=txt(files,relsPathOf(part));if(!s)return[];
  return [...parseXml(s).getElementsByTagNameNS(NS.rel,'Relationship')].map(e=>({id:e.getAttribute('Id'),type:e.getAttribute('Type')||'',target:e.getAttribute('Target')||'',mode:e.getAttribute('TargetMode')}));
}
function dropRel(files,part,id){const rp=relsPathOf(part);const s=txt(files,rp);if(s)put(files,rp,s.replace(new RegExp(`<Relationship\\b[^>]*\\bId="${id}"[^>]*/>`),''));}
function ctDropOverride(files,part){const s=txt(files,'[Content_Types].xml');if(s)put(files,'[Content_Types].xml',s.replace(new RegExp(`<Override\\b[^>]*PartName="/${part.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}"[^>]*/>`),''));}
const W14='http://schemas.microsoft.com/office/word/2010/wordml',W15='http://schemas.microsoft.com/office/word/2012/wordml';
const RT_CM='http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
const RT_CMX='http://schemas.microsoft.com/office/2011/relationships/commentsExtended';
const RT_HF=k=>`http://schemas.openxmlformats.org/officeDocument/2006/relationships/${k}`;
// PAGE / NUMPAGES fields become one marked link, so the editor can show and keep them
function markFields(doc){
  const mk=(kind,rPr)=>{const h=wNew(doc,'hyperlink',{anchor:'__f'+kind});const r=wNew(doc,'r');if(rPr)r.appendChild(rPr.cloneNode(true));const t=wNew(doc,'t');t.textContent=kind==='PAGE'?'1':'#';r.appendChild(t);h.appendChild(r);return h;};
  const kindOf=s=>{const m=/^\s*(PAGE|NUMPAGES|SECTIONPAGES)\b/i.exec(s||'');return m?(m[1].toUpperCase()==='PAGE'?'PAGE':'NUMPAGES'):null;};
  for(const f of [...doc.getElementsByTagNameNS(W,'fldSimple')]){const k=kindOf(wAttr(f,'instr'));if(!k)continue;const r=f.getElementsByTagNameNS(W,'r')[0];f.parentNode.replaceChild(mk(k,r&&kid(r,'rPr')),f);}
  for(const p of [...doc.getElementsByTagNameNS(W,'p')]){
    const rs=kids(p,'r');
    for(let i=0;i<rs.length;i++){
      const fc=kid(rs[i],'fldChar');if(!fc||wAttr(fc,'fldCharType')!=='begin')continue;
      let depth=0,j=i,instr='',sep=-1;
      for(;j<rs.length;j++){const c=kid(rs[j],'fldChar');const t=c&&wAttr(c,'fldCharType');
        if(t==='begin')depth++;else if(t==='end'){depth--;if(!depth)break;}else if(t==='separate'&&depth===1)sep=j;
        if(depth===1&&sep<0)for(const it of rs[j].getElementsByTagNameNS(W,'instrText'))instr+=it.textContent;}
      if(j>=rs.length)break;const k=kindOf(instr);if(!k){i=j;continue;}
      const res=sep>=0?rs.slice(sep+1,j).find(r=>kid(r,'t')):null;
      const h=mk(k,(res&&kid(res,'rPr'))||kid(rs[i],'rPr'));p.insertBefore(h,rs[i]);
      for(let x=i;x<=j;x++)p.removeChild(rs[x]);i=j;
    }
  }
}
function readComments(files,docPart){
  const out=new Map();const rl=relsOf(files,docPart);
  const cr=rl.find(r=>r.type===RT_CM);if(!cr)return{list:out,path:null};
  const path=resolveTarget(docPart,cr.target);const s=txt(files,path);if(!s)return{list:out,path};
  const d=parseXml(s);const byPara=new Map();
  for(const c of kids(d.documentElement,'comment')){
    const ps=kids(c,'p');const text=ps.map(p=>textOf(p)).join('\n');const last=ps[ps.length-1];
    const pid=last?(last.getAttributeNS(W14,'paraId')||last.getAttribute('w14:paraId')):null;
    const it={id:+wAttr(c,'id'),author:wAttr(c,'author')||'',initials:wAttr(c,'initials')||'',date:wAttr(c,'date')||'',text,parent:null,done:false,pid};
    out.set(it.id,it);if(pid)byPara.set(pid,it);
  }
  const xr=rl.find(r=>r.type===RT_CMX);const xs=xr&&txt(files,resolveTarget(docPart,xr.target));
  if(xs)for(const e of parseXml(xs).getElementsByTagNameNS(W15,'commentEx')){
    const it=byPara.get(e.getAttributeNS(W15,'paraId'));if(!it)continue;
    const par=byPara.get(e.getAttributeNS(W15,'paraIdParent'));if(par)it.parent=par.id;it.done=/^(1|true)$/.test(e.getAttributeNS(W15,'done')||'');
  }
  return{list:out,path};
}

/* ---------- a brand-new document ---------- */
function blankDocx(){
  const files=new Map();
  put(files,'[Content_Types].xml',XML_HEAD+'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="'+OX+'wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="'+OX+'wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="'+OX+'wordprocessingml.settings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="'+OX+'extended-properties+xml"/></Types>');
  put(files,'_rels/.rels',XML_HEAD+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="'+RT.doc+'" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>');
  const now=new Date().toISOString().replace(/\.\d+Z$/,'Z');
  put(files,'docProps/core.xml',XML_HEAD+`<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>סימן</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`);
  put(files,'docProps/app.xml',XML_HEAD+'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Siman</Application></Properties>');
  put(files,'word/_rels/document.xml.rels',XML_HEAD+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="'+RT.styles+'" Target="styles.xml"/><Relationship Id="rId2" Type="'+RT.settings+'" Target="settings.xml"/></Relationships>');
  put(files,'word/settings.xml',XML_HEAD+'<w:settings xmlns:w="'+W+'"><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat><w:themeFontLang w:val="en-US" w:bidi="he-IL"/></w:settings>');
  const F='<w:rFonts w:ascii="David" w:hAnsi="David" w:cs="David" w:eastAsia="David"/>';
  put(files,'word/styles.xml',XML_HEAD+'<w:styles xmlns:w="'+W+'"><w:docDefaults><w:rPrDefault><w:rPr>'+F+'<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="he-IL"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:bidi/><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'+
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/><w:pPr><w:bidi/></w:pPr><w:rPr><w:rtl/></w:rPr></w:style>'+
    '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>'+
    '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'+
    '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>'+
    '</w:styles>');
  put(files,'word/document.xml',XML_HEAD+'<w:document xmlns:w="'+W+'" xmlns:r="'+NS.r+'" xmlns:wp="'+NS.wp+'"><w:body><w:p><w:pPr><w:bidi/></w:pPr></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/><w:bidi/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>');
  return files;
}

/* ---------- styles: read, add the ones the ribbon offers ---------- */
const classKey=id=>`${P}_${String(id).replace(/[ .]+/g,'-').replace(/[&]+/g,'and').toLowerCase()}`;
const BUILTIN=[
  {name:'title',label:'כותרת',id:'Title',ppr:'<w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:contextualSpacing/>',rpr:'<w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/>'},
  {name:'subtitle',label:'כותרת משנה',id:'Subtitle',ppr:'<w:spacing w:after="160"/>',rpr:'<w:color w:val="595959"/><w:spacing w:val="15"/><w:sz w:val="28"/><w:szCs w:val="28"/>'},
  {name:'heading 1',label:'כותרת 1',id:'Heading1',ppr:'<w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="80"/><w:outlineLvl w:val="0"/>',rpr:'<w:b/><w:bCs/><w:color w:val="1F3864"/><w:sz w:val="36"/><w:szCs w:val="36"/>'},
  {name:'heading 2',label:'כותרת 2',id:'Heading2',ppr:'<w:keepNext/><w:keepLines/><w:spacing w:before="160" w:after="80"/><w:outlineLvl w:val="1"/>',rpr:'<w:b/><w:bCs/><w:color w:val="2F5496"/><w:sz w:val="30"/><w:szCs w:val="30"/>'},
  {name:'heading 3',label:'כותרת 3',id:'Heading3',ppr:'<w:keepNext/><w:keepLines/><w:spacing w:before="160" w:after="80"/><w:outlineLvl w:val="2"/>',rpr:'<w:b/><w:bCs/><w:color w:val="2F5496"/><w:sz w:val="26"/><w:szCs w:val="26"/>'},
  {name:'quote',label:'ציטוט',id:'Quote',ppr:'<w:spacing w:before="160" w:after="160"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/>',rpr:'<w:i/><w:iCs/><w:color w:val="404040"/>'},
  {name:'list paragraph',label:'',id:'ListParagraph',ppr:'<w:ind w:left="720"/><w:contextualSpacing/>',rpr:''}
];
function readStyles(sdoc){
  const out=[];const root=sdoc.documentElement;
  for(const s of kids(root,'style')){
    const name=(wAttr(kid(s,'name'),'val')||'').toLowerCase();
    out.push({el:s,id:wAttr(s,'styleId'),type:wAttr(s,'type'),name,label:wAttr(kid(s,'name'),'val')||'',
      basedOn:wAttr(kid(s,'basedOn'),'val'),def:/^(1|true|on)$/.test(wAttr(s,'default')||''),
      q:!!kid(s,'qFormat'),hidden:!!kid(s,'semiHidden')||!!kid(s,'hidden')});
  }
  return out;
}
function ensureStyles(files){
  let s=txt(files,'word/styles.xml');
  if(!s){s=XML_HEAD+`<w:styles xmlns:w="${W}"></w:styles>`;put(files,'word/styles.xml',s);
    ctEnsureOverride(files,'word/styles.xml',OX+'wordprocessingml.styles+xml');addRel(files,'word/document.xml',RT.styles,'styles.xml');}
  const d=parseXml(s);const have=readStyles(d);
  const normal=have.find(x=>x.type==='paragraph'&&x.def)||have.find(x=>x.name==='normal');
  const ids=new Set(have.map(x=>x.id));let add='';
  for(const b of BUILTIN){
    if(have.some(x=>x.type==='paragraph'&&x.name===b.name))continue;
    let id=b.id;while(ids.has(id))id+='_';ids.add(id);
    add+=`<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${b.name}"/>${normal?`<w:basedOn w:val="${normal.id}"/><w:next w:val="${normal.id}"/>`:''}<w:uiPriority w:val="9"/><w:qFormat/><w:pPr>${b.ppr}</w:pPr><w:rPr>${b.rpr}</w:rPr></w:style>`;
  }
  if(add){s=s.replace(/<\/w:styles>\s*$/,add+'</w:styles>');put(files,'word/styles.xml',s);}
}
function addRel(files,part,type,target,mode){
  const rp=relsPathOf(part);let s=txt(files,rp)||(XML_HEAD+'<Relationships xmlns="'+NS.rel+'"></Relationships>');
  const used=new Set([...s.matchAll(/Id="([^"]+)"/g)].map(m=>m[1]));let n=1,id;do id='rIdSm'+(n++);while(used.has(id));
  s=s.replace('</Relationships>',`<Relationship Id="${id}" Type="${type}" Target="${esc(target)}"${mode?` TargetMode="${mode}"`:''}/></Relationships>`);
  put(files,rp,s);return id;
}

/* ---------- numbering: a bullet list and a numbered list for the ribbon ---------- */
function ensureNumbering(files,docPart){
  let s=txt(files,'word/numbering.xml');
  if(!s){s=XML_HEAD+`<w:numbering xmlns:w="${W}"></w:numbering>`;
    ctEnsureOverride(files,'word/numbering.xml',OX+'wordprocessingml.numbering+xml');
    const rels=txt(files,relsPathOf(docPart))||'';if(!rels.includes(RT.numbering))addRel(files,docPart,RT.numbering,'numbering.xml');}
  const absIds=[...s.matchAll(/<w:abstractNum\b[^>]*w:abstractNumId="(\d+)"/g)].map(m=>+m[1]);
  const numIds=[...s.matchAll(/<w:num\b[^>]*w:numId="(\d+)"/g)].map(m=>+m[1]);
  const a1=Math.max(0,...absIds)+1,a2=a1+1,n1=Math.max(0,...numIds)+1,n2=n1+1;
  const bul=['•','◦','▪'],fmt=['decimal','lowerLetter','lowerRoman'];
  const lvl=(i,kind)=>`<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${kind==='b'?'bullet':fmt[i%3]}"/><w:lvlText w:val="${kind==='b'?bul[i%3]:`%${i+1}.`}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720*(i+1)}" w:hanging="360"/></w:pPr>${kind==='b'?'<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:hint="default"/></w:rPr>':''}</w:lvl>`;
  const abs=(id,kind)=>`<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>${[0,1,2,3,4,5,6,7,8].map(i=>lvl(i,kind)).join('')}</w:abstractNum>`;
  const absXml=abs(a1,'b')+abs(a2,'n');
  const i=s.search(/<w:num\b/);
  s=i>=0?s.slice(0,i)+absXml+s.slice(i):s.replace(/<\/w:numbering>\s*$/,absXml+'</w:numbering>');
  s=s.replace(/<\/w:numbering>\s*$/,`<w:num w:numId="${n1}"><w:abstractNumId w:val="${a1}"/></w:num><w:num w:numId="${n2}"><w:abstractNumId w:val="${a2}"/></w:num></w:numbering>`);
  put(files,'word/numbering.xml',s);
  return{bullet:n1,decimal:n2,decAbs:a2};
}

/* ---------- the working copy of document.xml ---------- */
function isAtomChild(n){
  if(n.nodeType!==1)return false;
  if(n.namespaceURI===NS.mc&&n.localName==='AlternateContent')return true;
  if(n.namespaceURI!==W)return false;
  if(['drawing','pict','object','footnoteReference','endnoteReference'].includes(n.localName))return true;
  return n.localName==='br'&&wAttr(n,'type')==='page';
}
function insideObject(r){
  for(let p=r.parentNode;p&&p.nodeType===1;p=p.parentNode){
    if(p.namespaceURI===NS.mc&&p.localName==='AlternateContent')return true;
    if(p.namespaceURI===W&&['drawing','pict','object','txbxContent','footnote','endnote'].includes(p.localName))return true;
    if(p.namespaceURI===W&&p.localName==='body')return false;
  }
  return false;
}
// content controls and custom-XML wrappers around paragraphs/tables/rows/cells are opened up,
// so every paragraph and table sits directly in its container (as it does on screen)
function unwrap(container){
  let again=true;
  while(again){again=false;
    for(const c of [...container.childNodes]){
      if(c.nodeType!==1||c.namespaceURI!==W)continue;
      if(c.localName==='sdt'){const sc=kid(c,'sdtContent');while(sc&&sc.firstChild)container.insertBefore(sc.firstChild,c);container.removeChild(c);again=true;}
      else if(c.localName==='customXml'){for(const g of [...c.childNodes])if(!(g.nodeType===1&&g.localName==='customXmlPr'))container.insertBefore(g,c);container.removeChild(c);again=true;}
    }
  }
  for(const c of kids(container,'tbl')){unwrap(c);for(const tr of kids(c,'tr')){unwrap(tr);for(const tc of kids(tr,'tc'))unwrap(tc);}}
}
// a run holding a picture or a page break together with text is split, so the text stays editable
function splitAtomRuns(doc){
  for(const r of [...doc.getElementsByTagNameNS(W,'r')]){
    if(insideObject(r))continue;
    const rPr=kid(r,'rPr');const content=[...r.childNodes].filter(n=>n.nodeType===1&&n!==rPr);
    if(content.length<2||!content.some(isAtomChild))continue;
    for(const c of content){const nr=wNew(doc,'r');if(rPr)nr.appendChild(rPr.cloneNode(true));nr.appendChild(c);r.parentNode.insertBefore(nr,r);}
    r.parentNode.removeChild(r);
  }
}
function textOf(el){let t='';for(const x of el.getElementsByTagNameNS(W,'t'))t+=x.textContent;return t;}
function hebrewMajority(doc){const t=textOf(doc.documentElement);return (t.match(HEBG)||[]).length>=(t.match(LATG)||[]).length;}

/* paragraph direction from the style chain */
function styleBidiMap(sdoc){
  const map={};const st={};let defId=null,defBidi=false;
  const ppd=sdoc.getElementsByTagNameNS(W,'pPrDefault')[0];
  if(ppd){const b=kid(kid(ppd,'pPr'),'bidi');if(b)defBidi=onOff(b);}
  for(const s of kids(sdoc.documentElement,'style')){if(wAttr(s,'type')!=='paragraph')continue;const id=wAttr(s,'styleId');st[id]=s;if(/^(1|true|on)$/.test(wAttr(s,'default')||''))defId=id;}
  const get=(id,seen=new Set())=>{
    if(id in map)return map[id];if(!id||!st[id]||seen.has(id))return defBidi;seen.add(id);
    const b=kid(kid(st[id],'pPr'),'bidi');const v=b?onOff(b):get(wAttr(kid(st[id],'basedOn'),'val'),seen);return(map[id]=v);
  };
  return{of:id=>get(id||defId),defId};
}
function paraBidi(p,sb){
  const pPr=kid(p,'pPr');const b=kid(pPr,'bidi');if(b)return onOff(b);
  return sb.of(wAttr(kid(pPr,'pStyle'),'val'));
}
// for the screen only: docx-preview reads jc literally, but in a right-to-left paragraph Word reads
// "left" as the start (right side) and "right" as the end
function swapJc(pPr){const jc=kid(pPr,'jc');if(!jc)return;const v=wAttr(jc,'val');if(v==='left')wSet(jc,'val','right');else if(v==='right')wSet(jc,'val','left');}
// for the screen only: Hebrew letters use the complex-script size, bold, italic and font
function csRun(rPr){
  if(!rPr)return;const d=rPr.ownerDocument;
  const pairs=[['szCs','sz'],['bCs','b'],['iCs','i']];
  for(const [cs,l] of pairs){const c=kid(rPr,cs);if(!c)continue;let e=kid(rPr,l);if(!e){e=wNew(d,l);rPr.insertBefore(e,c);}
    const v=wAttr(c,'val');if(v===null)e.removeAttributeNS(W,'val');else wSet(e,'val',v);}
  const f=kid(rPr,'rFonts');if(f){const cs=wAttr(f,'cs'),th=wAttr(f,'cstheme');
    if(cs){wSet(f,'ascii',cs);wSet(f,'hAnsi',cs);f.removeAttributeNS(W,'asciiTheme');f.removeAttributeNS(W,'hAnsiTheme');}
    else if(th){wSet(f,'asciiTheme',th);wSet(f,'hAnsiTheme',th);f.removeAttributeNS(W,'ascii');}}
}

/* everything the editor needs from a .docx */
async function prepare(bytes){
  const files=bytes?await readZip(bytes):blankDocx();
  const docPart=mainPart(files);
  if(!files.has(docPart))throw new Error('לא נמצא תוכן Word בקובץ');
  ensureStyles(files);
  const num=ensureNumbering(files,docPart);
  const docStr=txt(files,docPart);
  const norm=parseXml(docStr);const body=norm.getElementsByTagNameNS(W,'body')[0];
  if(!body)throw new Error('לא נמצא גוף המסמך');
  unwrap(body);splitAtomRuns(norm);
  const sdoc=parseXml(txt(files,'word/styles.xml'));
  const sb=styleBidiMap(sdoc);const heb=hebrewMajority(norm);
  // the copy that is drawn on screen
  const view=norm.cloneNode(true);
  const nr=[...norm.getElementsByTagNameNS(W,'r')],vr=[...view.getElementsByTagNameNS(W,'r')];
  const atoms=[];
  nr.forEach((r,i)=>{
    if(insideObject(r)||(r.parentNode&&r.parentNode.localName==='hyperlink'))return;
    const c=[...r.childNodes].filter(n=>n.nodeType===1&&!(n.namespaceURI===W&&n.localName==='rPr'));
    if(c.length!==1||!isAtomChild(c[0]))return;
    const pb=c[0].localName==='br';const k=atoms.length;
    atoms.push({xml:r,pb,draw:!pb&&!!r.getElementsByTagNameNS(NS.wp,'extent').length});
    const v=vr[i],h=wNew(view,'hyperlink',{anchor:`__${pb?'p':'a'}${k}`});v.parentNode.insertBefore(h,v);h.appendChild(v);
  });
  // comments: a marker where each one starts and ends, so it shows on screen and survives editing
  for(const [tag,k] of [['commentRangeStart','cs'],['commentRangeEnd','ce']])for(const m of [...view.getElementsByTagNameNS(W,tag)]){
    const par=m.parentNode;if(!par||par.namespaceURI!==W||par.localName!=='p')continue;
    const h=wNew(view,'hyperlink',{anchor:`__${k}${wAttr(m,'id')}`});const r=wNew(view,'r');const t=wNew(view,'t');t.textContent='\u200B';r.appendChild(t);h.appendChild(r);par.insertBefore(h,m);}
  for(const p of view.getElementsByTagNameNS(W,'p'))if(paraBidi(p,sb))swapJc(kid(p,'pPr'));
  for(const r of view.getElementsByTagNameNS(W,'r'))if(HEB.test(textOf(r)))csRun(kid(r,'rPr'));
  const vs=sdoc.cloneNode(true);const vsb=styleBidiMap(vs);
  for(const s of kids(vs.documentElement,'style')){
    if(wAttr(s,'type')==='paragraph'&&vsb.of(wAttr(s,'styleId')))swapJc(kid(s,'pPr'));
    if(heb)csRun(kid(s,'rPr'));
  }
  if(heb){const d=vs.getElementsByTagNameNS(W,'rPrDefault')[0];if(d)csRun(kid(d,'rPr'));}
  const vfiles=new Map(files);put(vfiles,docPart,ser(view));put(vfiles,'word/styles.xml',ser(vs));
  // headers and footers: which part is shown (first-page header when the document has one), and its page-number fields
  const sect0=kids(body,'sectPr').pop()||null;const titlePg=!!kid(sect0,'titlePg');const hf={};const drl=relsOf(files,docPart);
  for(const kind of ['header','footer']){
    const refs=kids(sect0,kind+'Reference');const ref=(titlePg&&refs.find(r=>wAttr(r,'type')==='first'))||refs.find(r=>(wAttr(r,'type')||'default')==='default')||null;
    let path=null;if(ref){const rl=drl.find(x=>x.id===(ref.getAttributeNS(NS.r,'id')||ref.getAttribute('r:id')));if(rl)path=resolveTarget(docPart,rl.target);}
    hf[kind]={type:ref?(wAttr(ref,'type')||'default'):'default',path:path&&files.has(path)?path:null};
    if(hf[kind].path){try{const hd=parseXml(txt(files,path));markFields(hd);put(vfiles,path,ser(hd));}catch(_){}}
  }
  const cm=readComments(files,docPart);
  const vz=new JSZip();for(const [p,b] of vfiles)vz.file(p,b);
  const viewBytes=await vz.generateAsync({type:'uint8array'});
  // page geometry (last section)
  const sect=kids(body,'sectPr').pop()||null;
  const pg=kid(sect,'pgSz'),mg=kid(sect,'pgMar');
  const page={w:+wAttr(pg,'w')||11906,h:+wAttr(pg,'h')||16838,landscape:wAttr(pg,'orient')==='landscape',
    top:+wAttr(mg,'top')||1440,bottom:+wAttr(mg,'bottom')||1440,left:+(wAttr(mg,'left')??1440),right:+(wAttr(mg,'right')??1440),
    header:+wAttr(mg,'header')||708,footer:+wAttr(mg,'footer')||708};
  // default font and size
  const def=defaultsOf(files,sdoc,heb);
  // theme fonts for Hebrew (docx-preview knows only the Latin ones)
  let themeCs={};const th=[...files.keys()].find(k=>/^word\/theme\/theme\d*\.xml$/.test(k));
  if(th){const t=txt(files,th);const mj=t.match(/<a:majorFont>[\s\S]*?<a:cs typeface="([^"]*)"/),mn=t.match(/<a:minorFont>[\s\S]*?<a:cs typeface="([^"]*)"/);
    themeCs={major:mj&&mj[1],minor:mn&&mn[1]};}
  return{files,docPart,docStr,norm,body,sdoc,styles:readStyles(sdoc),atoms,viewBytes,page,sect,num,heb,def,themeCs,sb,hf,cm};
}
function defaultsOf(files,sdoc,heb){
  let font=null,size=null;
  const d=sdoc.getElementsByTagNameNS(W,'rPrDefault')[0];const rPr=d&&kid(d,'rPr');
  const normal=kids(sdoc.documentElement,'style').find(s=>wAttr(s,'type')==='paragraph'&&/^(1|true|on)$/.test(wAttr(s,'default')||''));
  const look=r=>{if(!r)return;const f=kid(r,'rFonts');if(f){const v=heb?(wAttr(f,'cs')||wAttr(f,'ascii')):(wAttr(f,'ascii')||wAttr(f,'cs'));if(v)font=v;}
    const z=kid(r,heb?'szCs':'sz')||kid(r,'sz');if(z)size=(+wAttr(z,'val'))/2;};
  look(rPr);look(normal&&kid(normal,'rPr'));
  return{font:font||(heb?'David':'Calibri'),size:size||(rPr?10:12)};
}

/* ------------------------------------------------------------ the editor screen */
const ICO={
  undo:'<path d="M7 5L3 9l4 4"/><path d="M3 9h9a5 5 0 0 1 0 10H8"/>',
  redo:'<path d="M13 5l4 4-4 4"/><path d="M17 9H8a5 5 0 0 0 0 10h4"/>',
  cut:'<circle cx="6" cy="15" r="2.5"/><circle cx="14" cy="15" r="2.5"/><path d="M7.5 13L14 3M12.5 13L6 3"/>',
  copy:'<rect x="7" y="7" width="10" height="11" rx="1.5"/><path d="M4 13V4a1 1 0 0 1 1-1h8"/>',
  paste:'<rect x="4" y="4" width="12" height="14" rx="1.5"/><path d="M7.5 4V2.5h5V4"/><path d="M7 9h6M7 12h6M7 15h4"/>',
  brush:'<path d="M4 3h10v5H4z"/><path d="M14 5.5h2v4H9.5v3"/><rect x="8" y="12.5" width="3" height="5.5" rx="1"/>',
  bullets:'<circle cx="16" cy="5" r="1.2" fill="currentColor"/><circle cx="16" cy="10" r="1.2" fill="currentColor"/><circle cx="16" cy="15" r="1.2" fill="currentColor"/><path d="M3 5h10M3 10h10M3 15h10"/>',
  numbers:'<path d="M15 3.5h1.2v4M14.6 7.5h2.4M14.5 12.2c.4-.7 2.6-.8 2.4.7-.2 1-2.4 1.6-2.4 3h2.6"/><path d="M3 5h9M3 10h9M3 15h9"/>',
  outdent:'<path d="M3 4h14M9 8h8M9 12h8M3 16h14"/><path d="M3 10l3-2v4z" fill="currentColor"/>',
  indent:'<path d="M3 4h14M3 8h8M3 12h8M3 16h14"/><path d="M17 10l-3-2v4z" fill="currentColor"/>',
  aright:'<path d="M3 4h14M7 8h10M3 12h14M7 16h10"/>',
  acenter:'<path d="M3 4h14M5.5 8h9M3 12h14M5.5 16h9"/>',
  aleft:'<path d="M3 4h14M3 8h10M3 12h14M3 16h10"/>',
  ajust:'<path d="M3 4h14M3 8h14M3 12h14M3 16h14"/>',
  rtl:'<path d="M9 3h7M13 3v13M10 3v13M10 3a3 3 0 0 0 0 6"/><path d="M8 15H3m0 0l2-2m-2 2l2 2"/>',
  ltr:'<path d="M4 3h7M8 3v13M5 3v13M5 3a3 3 0 0 0 0 6"/><path d="M12 15h5m0 0l-2-2m2 2l-2 2"/>',
  spacing:'<path d="M9 4h8M9 8h8M9 12h8M9 16h8"/><path d="M4 3v14M2 5l2-2 2 2M2 15l2 2 2-2"/>',
  pilcrow:'<path d="M14 3v14M11 3v14M11 3H8.5a3.5 3.5 0 0 0 0 7H11"/>',
  find:'<circle cx="8.5" cy="8.5" r="5"/><path d="M12.5 12.5L17 17"/>',
  replace:'<path d="M3 6h9l-3-3M17 14H8l3 3"/>',
  table:'<rect x="3" y="4" width="14" height="12" rx="1"/><path d="M3 8h14M3 12h14M8 4v12M12.5 4v12"/>',
  image:'<rect x="3" y="4" width="14" height="12" rx="1.5"/><circle cx="7.5" cy="8.5" r="1.4"/><path d="M17 13l-4-4-7 7"/>',
  link:'<path d="M8.5 11.5a3 3 0 0 0 4.2 0l3-3a3 3 0 0 0-4.2-4.2l-1 1"/><path d="M11.5 8.5a3 3 0 0 0-4.2 0l-3 3a3 3 0 0 0 4.2 4.2l1-1"/>',
  pagebreak:'<path d="M5 2v5h10V2M5 18v-5h10v5"/><path d="M2 10h3M7.5 10h2M11 10h2M15 10h3"/>',
  hr:'<path d="M3 10h14"/><path d="M5 5h10M5 15h10" opacity=".4"/>',
  symbol:'<path d="M5 16h3v-2.2A6 6 0 1 1 12 13.8V16h3"/>',
  date:'<rect x="3" y="4" width="14" height="13" rx="1.5"/><path d="M3 8h14M7 2.5v3M13 2.5v3M6.5 11.5h2M11.5 11.5h2M6.5 14.5h2"/>',
  margins:'<rect x="4" y="2.5" width="12" height="15" rx="1"/><path d="M4 6h12M4 14h12M7 2.5v15M13 2.5v15" stroke-dasharray="1.5 1.5"/>',
  orient:'<rect x="2.5" y="6" width="15" height="10" rx="1"/><path d="M6 2.5h5" /><path d="M11 2.5l-1.5-1M11 2.5l-1.5 1"/>',
  size:'<path d="M5 2.5h7l3.5 3.5v11.5H5z"/><path d="M12 2.5V6h3.5"/>',
  print:'<path d="M5.5 7V3h9v4"/><rect x="3" y="7" width="14" height="7" rx="1.5"/><path d="M6 12h8v5H6z"/>',
  zin:'<circle cx="8.5" cy="8.5" r="5"/><path d="M12.5 12.5L17 17M6.5 8.5h4M8.5 6.5v4"/>',
  zout:'<circle cx="8.5" cy="8.5" r="5"/><path d="M12.5 12.5L17 17M6.5 8.5h4"/>',
  pages:'<rect x="5" y="2.5" width="10" height="6.5" rx="1"/><rect x="5" y="11" width="10" height="6.5" rx="1"/>',
  spell:'<path d="M3 12l3-8 3 8M4.2 9.5h3.6"/><path d="M11 14l2.5 2.5L18 11"/>',
  rowAbove:'<rect x="3" y="9" width="14" height="8" rx="1"/><path d="M3 13h14M10 9v8"/><path d="M10 2v5M7.5 4.5h5"/>',
  rowBelow:'<rect x="3" y="3" width="14" height="8" rx="1"/><path d="M3 7h14M10 3v8"/><path d="M10 13v5M7.5 15.5h5"/>',
  colRight:'<rect x="3" y="3" width="8" height="14" rx="1"/><path d="M7 3v14M3 10h8"/><path d="M13 10h5M15.5 7.5v5"/>',
  colLeft:'<rect x="9" y="3" width="8" height="14" rx="1"/><path d="M13 3v14M9 10h8"/><path d="M2 10h5M4.5 7.5v5"/>',
  delRow:'<rect x="3" y="6" width="14" height="8" rx="1"/><path d="M8 8.5l4 3M12 8.5l-4 3"/>',
  delCol:'<rect x="6" y="3" width="8" height="14" rx="1"/><path d="M8 8l4 4M12 8l-4 4"/>',
  delTable:'<rect x="3" y="4" width="14" height="12" rx="1"/><path d="M7 7l6 6M13 7l-6 6"/>',
  shade:'<path d="M5 11l5-7 5 7a5 5 0 0 1-10 0z"/>',
  borders:'<rect x="3" y="3" width="14" height="14"/><path d="M3 10h14M10 3v14"/>',
  noborders:'<rect x="3" y="3" width="14" height="14" stroke-dasharray="1.5 2"/><path d="M3 10h14M10 3v14" stroke-dasharray="1.5 2"/>',
  grow:'<path d="M3 16l4.5-12L12 16M4.6 12h5.8"/><path d="M14 7l2-2.5L18 7"/>',
  shrink:'<path d="M3 16l4.5-12L12 16M4.6 12h5.8"/><path d="M14 4.5l2 2.5 2-2.5"/>',
  sub:'<path d="M3 4l7 9M10 4l-7 9"/><path d="M13 12.5c.3-.9 3.1-1 2.9.8-.2 1.1-2.9 1.9-2.9 3.7h3.2"/>',
  sup:'<path d="M3 7l7 9M10 7l-7 9"/><path d="M13 3.5c.3-.9 3.1-1 2.9.8-.2 1.1-2.9 1.9-2.9 3.7h3.2"/>',
  clear:'<path d="M5 15l3.5-11h3L15 15M6.3 11h7.4"/><path d="M13 13l4 4M17 13l-4 4"/>',
  caret:'<path d="M6 8l4 4 4-4"/>',
  header:'<rect x="3.5" y="2.5" width="13" height="15" rx="1"/><path d="M6 5.5h8" stroke-width="2.4"/><path d="M6 10h8M6 13h6" opacity=".45"/>',
  footer:'<rect x="3.5" y="2.5" width="13" height="15" rx="1"/><path d="M6 14.5h8" stroke-width="2.4"/><path d="M6 6h8M6 9h6" opacity=".45"/>',
  pagenum:'<rect x="3.5" y="2.5" width="13" height="15" rx="1"/><path d="M8.6 11.2c.2-1 2.7-1.1 2.6.4-.1 1-2.6 1.6-2.6 2.9h2.8"/>',
  toc:'<path d="M3 4.5h9M5 8.5h7M5 12.5h7M3 16.5h9"/><path d="M14.5 4.5h2.5M14.5 8.5h2.5M14.5 12.5h2.5M14.5 16.5h2.5" stroke-dasharray="1 1.4"/>',
  comment:'<path d="M3.5 4.5h13v9h-7l-4 3v-3h-2z"/><path d="M10 7v4M8 9h4"/>',
  cmDel:'<path d="M3.5 4.5h13v9h-7l-4 3v-3h-2z"/><path d="M8 7.2l4 3.6M12 7.2l-4 3.6"/>',
  prev:'<path d="M3.5 4.5h13v9h-7l-4 3v-3h-2z"/><path d="M8.5 7l3 2-3 2"/>',
  next:'<path d="M3.5 4.5h13v9h-7l-4 3v-3h-2z"/><path d="M11.5 7l-3 2 3 2"/>',
  cmShow:'<path d="M3.5 4.5h13v9h-7l-4 3v-3h-2z"/><path d="M7 9h6"/>',
  count:'<path d="M4 5h4M4 9h4M4 13h4"/><path d="M11 4.5c.3-.8 2.9-.9 2.7.7-.2 1.1-2.7 1.8-2.7 3.4h3M11.2 12.2c.5-.9 3-.7 2.7.8.3.9-.3 2.2-1.7 2.1-.6 0-1.1-.2-1.3-.7"/>',
  ac:'<path d="M3 15l3.5-10h1L11 15M4.3 11.5h5.4"/><path d="M12 12.5l2.2 2.3L18 9.5"/>',
  closeX:'<rect x="3" y="3" width="14" height="14" rx="2"/><path d="M7 7l6 6M13 7l-6 6"/>',
  swap:'<path d="M5 3v14M5 3L2.5 5.5M5 3l2.5 2.5M15 17V3M15 17l-2.5-2.5M15 17l2.5-2.5"/>',
  clock:'<circle cx="10" cy="10" r="7"/><path d="M10 6v4l2.5 2"/>'
};
const ic=n=>`<svg viewBox="0 0 20 20" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">${ICO[n]||''}</svg>`;
const B=(cmd,icon,title,extra='')=>`<button type="button" class="wb" data-cmd="${cmd}" title="${title}" aria-label="${title}" ${extra}>${icon}</button>`;
const BL=(cmd,icon,label,title,extra='')=>`<button type="button" class="wb big" data-cmd="${cmd}" title="${title||label}" ${extra}>${ic(icon)}<span>${label}</span></button>`;
const G=(label,inner,cls='')=>`<div class="wg ${cls}"><div class="wg-b">${inner}</div><div class="wg-l">${label}</div></div>`;
const FONTS_HE=['David','Arial','FrankRuehl','Miriam','Narkisim','Guttman Yad','Heebo','Rubik','Noto Sans Hebrew','Tahoma','Times New Roman'];
const FONTS_LAT=['Calibri','Cambria','Segoe UI','Georgia','Verdana','Courier New','Consolas'];
const SIZES=[8,9,10,10.5,11,12,14,16,18,20,22,24,26,28,36,48,72];
const COLORS=['000000','404040','7F7F7F','BFBFBF','FFFFFF','C00000','FF0000','FFC000','FFFF00','92D050','00B050','00B0F0','0070C0','002060','7030A0','2E58CC','1F3864','833C0B','385623','7F6000'];
const HLS=['FFFF00','00FF00','00FFFF','FF00FF','0000FF','FF0000','000080','008080','008000','800080','800000','808000','808080','C0C0C0','000000'];
const SYMBOLS=['₪','€','$','£','©','®','™','§','°','±','×','÷','≠','≤','≥','…','–','—','״','׳','•','✓','✗','★','→','←','↑','↓','½','¼','¾','∞','√','π','α','β','Ω','µ'];

function ribbonHtml(){
  return `
<div class="wpanel" data-panel="home">
 ${G('ביטול',B('undo',ic('undo'),'ביטול (Ctrl+Z)')+B('redo',ic('redo'),'חזרה (Ctrl+Y)'),'tight')}
 ${G('לוח',BL('paste','paste','הדבקה','הדבקה (Ctrl+V)')+`<div class="col">${B('cut',ic('cut'),'גזירה (Ctrl+X)')}${B('copy',ic('copy'),'העתקה (Ctrl+C)')}${B('painter',ic('brush'),'מברשת עיצוב: מעתיקה את העיצוב לטקסט הבא שתסמנו','data-toggle')}</div>`)}
 ${G('גופן',`<div class="row"><select class="wsel font" data-in="font" title="גופן" aria-label="גופן"></select><input class="wsel size" data-in="size" list="weSizes" inputmode="decimal" title="גודל גופן" aria-label="גודל גופן"><datalist id="weSizes">${SIZES.map(s=>`<option value="${s}">`).join('')}</datalist>${B('grow',ic('grow'),'הגדלת הגופן (Ctrl+Shift+>)')}${B('shrink',ic('shrink'),'הקטנת הגופן (Ctrl+Shift+<)')}</div>
   <div class="row">${B('bold','<b class="gl">B</b>','מודגש (Ctrl+B)')}${B('italic','<i class="gl" style="font-family:Georgia,serif">I</i>','נטוי (Ctrl+I)')}${B('underline','<u class="gl">U</u>','קו תחתי (Ctrl+U)')}${B('strike','<s class="gl">ab</s>','קו חוצה')}${B('sub',ic('sub'),'כתב תחתי')}${B('sup',ic('sup'),'כתב עילי')}
   <span class="wsplit">${B('color','<span class="gl a">A</span><i class="cbar" data-bar="color"></i>','צבע גופן')}${B('colorMenu',ic('caret'),'בחירת צבע גופן','data-pop="color"')}</span>
   <span class="wsplit">${B('hilite','<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round" aria-hidden="true"><path d="M5 12l6-8 4 3-6 8H6z"/><path d="M5 12l-1.5 3H7"/></svg><i class="cbar" data-bar="hilite"></i>','צבע הדגשה (מרקר)')}${B('hiliteMenu',ic('caret'),'בחירת צבע הדגשה','data-pop="hilite"')}</span>
   ${B('clear',ic('clear'),'ניקוי העיצוב (Ctrl+Space)')}</div>`)}
 ${G('פסקה',`<div class="row">${B('bullets',ic('bullets'),'תבליטים')}${B('numbers',ic('numbers'),'מספור')}${B('outdent',ic('outdent'),'הקטנת הזחה')}${B('indent',ic('indent'),'הגדלת הזחה')}${B('rtl',ic('rtl'),'פסקה מימין לשמאל')}${B('ltr',ic('ltr'),'פסקה משמאל לימין')}${B('marks',ic('pilcrow'),'הצגת סימני עיצוב','data-toggle')}</div>
   <div class="row">${B('right',ic('aright'),'יישור לימין (Ctrl+R)')}${B('center',ic('acenter'),'מרכוז (Ctrl+E)')}${B('left',ic('aleft'),'יישור לשמאל (Ctrl+L)')}${B('justify',ic('ajust'),'יישור לשני הצדדים (Ctrl+J)')}${B('spacingMenu',ic('spacing'),'ריווח בין שורות ופסקאות','data-pop="spacing"')}</div>`)}
 ${G('סגנונות','<div class="wgal" id="weGal"></div>','gal')}
 ${G('עריכה',BL('find','find','חיפוש','חיפוש (Ctrl+F)')+BL('replace','replace','החלפה','חיפוש והחלפה (Ctrl+H)'))}
</div>
<div class="wpanel" data-panel="insert" hidden>
 ${G('טבלאות',BL('tableMenu','table','טבלה','הוספת טבלה','data-pop="table"'))}
 ${G('איורים',BL('image','image','תמונה','הוספת תמונה מהמכשיר'))}
 ${G('קישורים',BL('link','link','קישור','הוספת קישור (Ctrl+K)'))}
 ${G('עמודים',BL('pagebreak','pagebreak','מעבר עמוד','מעבר עמוד (Ctrl+Enter)')+BL('hr','hr','קו מפריד','קו מפריד לרוחב העמוד'))}
 ${G('כותרת עליונה ותחתונה',BL('hdr','header','כותרת עליונה','עריכת הכותרת העליונה (או לחיצה כפולה על החלק העליון של העמוד)')+BL('ftr','footer','כותרת תחתונה','עריכת הכותרת התחתונה (או לחיצה כפולה על החלק התחתון של העמוד)')+BL('pnMenu','pagenum','מספר עמוד','הוספת מספרי עמודים','data-pop="pagenum"'))}
 ${G('תוכן',BL('tocMenu','toc','תוכן עניינים','תוכן עניינים לפי הכותרות במסמך','data-pop="toc"')+BL('cmNew','comment','הערה','הוספת הערה לטקסט המסומן (Ctrl+Alt+M)'))}
 ${G('סמלים',BL('symbolMenu','symbol','סמל','הוספת סמל','data-pop="symbol"')+BL('date','date','תאריך','הוספת התאריך של היום'))}
</div>
<div class="wpanel" data-panel="review" hidden>
 ${G('הגהה',BL('spell','spell','בדיקת איות','סימון שגיאות כתיב','data-toggle aria-pressed="true"')+BL('acToggle','ac','תיקון אוטומטי','תיקונים בזמן ההקלדה: גרשיים ״ וגרש ׳ בעברית, מקף –, אות גדולה בתחילת משפט באנגלית, © ½ ועוד','data-toggle aria-pressed="true"')+BL('wcount','count','ספירת מילים','מילים, תווים, פסקאות ועמודים'))}
 ${G('הערות',BL('cmNew','comment','הערה חדשה','הוספת הערה לטקסט המסומן (Ctrl+Alt+M)')+BL('cmDel','cmDel','מחיקה','מחיקת ההערה הנוכחית')+`<div class="col">${B('cmPrev',ic('prev'),'ההערה הקודמת')}${B('cmNext',ic('next'),'ההערה הבאה')}</div>`+BL('cmShow','cmShow','הצגת הערות','הצגה או הסתרה של ההערות בצד העמוד','data-toggle aria-pressed="true"'))}
</div>
<div class="wpanel" data-panel="hf" hidden>
 ${G('הוספה',BL('pnMenu','pagenum','מספר עמוד','הוספת מספרי עמודים','data-pop="pagenum"')+BL('pnXofY','pagenum','עמוד X מתוך Y','"עמוד 1 מתוך 5" במקום הסמן')+BL('date','date','תאריך','הוספת התאריך של היום'))}
 ${G('ניווט','<button type="button" class="wb big" data-cmd="hfGo" id="weHfGo" title="מעבר בין הכותרת העליונה לתחתונה">'+ic('swap')+'<span>לכותרת התחתונה</span></button>')}
 ${G('סגירה',BL('hfClose','closeX','סגירת הכותרות','חזרה לגוף המסמך (Esc)'))}
</div>
<div class="wpanel" data-panel="layout" hidden>
 ${G('הגדרת עמוד',BL('marginsMenu','margins','שוליים','שולי העמוד','data-pop="margins"')+BL('orientMenu','orient','כיוון','לאורך או לרוחב','data-pop="orient"')+BL('sizeMenu','size','גודל','גודל הנייר','data-pop="paper"'))}
 ${G('הזחה',`<label class="num">מימין<input type="number" step="0.25" min="0" data-in="indR" dir="ltr"> ס״מ</label><label class="num">משמאל<input type="number" step="0.25" min="0" data-in="indL" dir="ltr"> ס״מ</label>`)}
 ${G('ריווח',`<label class="num">לפני<input type="number" step="6" min="0" data-in="spB" dir="ltr"> נק׳</label><label class="num">אחרי<input type="number" step="6" min="0" data-in="spA" dir="ltr"> נק׳</label>`)}
 ${G('כיוון הטקסט',BL('rtl','rtl','מימין לשמאל','פסקה מימין לשמאל')+BL('ltr','ltr','משמאל לימין','פסקה משמאל לימין'))}
</div>
<div class="wpanel" data-panel="view" hidden>
 ${G('זום',B('zout',ic('zout'),'הקטנה')+'<span class="zv" id="weZoomV">100%</span>'+B('zin',ic('zin'),'הגדלה')+'<button type="button" class="wb txt" data-cmd="z100">100%</button><button type="button" class="wb txt" data-cmd="zfit">רוחב העמוד</button>')}
 ${G('הצגה',BL('guides','pages','גבולות עמוד','הצגת המקום שבו כל עמוד נגמר','data-toggle aria-pressed="true"')+BL('marks','pilcrow','סימני עיצוב','הצגת סימני פסקה','data-toggle')+BL('spell','spell','בדיקת איות','סימון שגיאות כתיב','data-toggle aria-pressed="true"'))}
 ${G('הדפסה',BL('print','print','הדפסה','הדפסה (Ctrl+P)'))}
</div>
<div class="wpanel" data-panel="table" hidden>
 ${G('שורות ועמודות',B('rowAbove',ic('rowAbove'),'הוספת שורה מעל')+B('rowBelow',ic('rowBelow'),'הוספת שורה מתחת')+B('colRight',ic('colRight'),'הוספת עמודה מימין')+B('colLeft',ic('colLeft'),'הוספת עמודה משמאל'))}
 ${G('מחיקה',B('delRow',ic('delRow'),'מחיקת השורה')+B('delCol',ic('delCol'),'מחיקת העמודה')+B('delTable',ic('delTable'),'מחיקת הטבלה'))}
 ${G('עיצוב',`<span class="wsplit">${B('shade',ic('shade')+'<i class="cbar" data-bar="shade"></i>','הצללת התא')}${B('shadeMenu',ic('caret'),'בחירת צבע הצללה','data-pop="shade"')}</span>`+B('borders',ic('borders'),'גבולות לכל התאים')+B('noborders',ic('noborders'),'בלי גבולות'))}
</div>`;
}
function shellHtml(){
  return `
<div class="we-top">
  <div class="we-brand"><span class="we-mark" aria-hidden="true">W</span><b>סימן</b><span class="we-sub">עורך Word</span></div>
  <label class="we-namebox"><input class="we-name" id="weName" aria-label="שם הקובץ" dir="auto" spellcheck="false"><span class="we-ext">.docx</span></label>
  <span class="we-state" id="weState" role="status"></span>
  <span class="grow"></span>
  <button type="button" class="we-tb" id="weDesk" title="פתיחת המסמך לסימון, חתימה ושיתוף">לחתימה ולסימון</button>
  <button type="button" class="we-tb" id="wePdf" title="שמירת המסמך כקובץ PDF">שמירה כ-PDF</button>
  <button type="button" class="we-tb pri" id="weSave" title="שמירה כקובץ Word (Ctrl+S)">שמירה</button>
  <button type="button" class="we-x" id="weClose" aria-label="סגירת העורך" title="סגירה">×</button>
</div>
<div class="we-tabs" role="tablist">
  <button type="button" role="tab" data-tab="home" aria-selected="true">בית</button>
  <button type="button" role="tab" data-tab="insert" aria-selected="false">הוספה</button>
  <button type="button" role="tab" data-tab="layout" aria-selected="false">פריסה</button>
  <button type="button" role="tab" data-tab="review" aria-selected="false">סקירה</button>
  <button type="button" role="tab" data-tab="view" aria-selected="false">תצוגה</button>
  <button type="button" role="tab" data-tab="table" aria-selected="false" class="ctx" hidden>טבלה</button>
  <button type="button" role="tab" data-tab="hf" aria-selected="false" class="ctx" hidden>כותרת עליונה ותחתונה</button>
</div>
<div class="we-ribbon" id="weRibbon">${ribbonHtml()}</div>
<div class="we-scroll" id="weScroll"><div class="we-canvas" id="weCanvas">
  <div class="we-sheet" id="weSheet">
    <div class="we-hdr" id="weHdr" title="לחיצה כפולה לעריכת הכותרת העליונה"></div>
    <div class="we-body ${P}" id="weBody" contenteditable="true" spellcheck="true" lang="he" role="textbox" aria-multiline="true" aria-label="תוכן המסמך"></div>
    <div class="we-notes ${P}" id="weNotes" aria-label="הערות שוליים"></div>
    <div class="we-ftr" id="weFtr" title="לחיצה כפולה לעריכת הכותרת התחתונה"></div>
    <div class="we-guides" id="weGuides" aria-hidden="true"></div>
    <div class="we-imgbox" id="weImgBox" hidden><i data-h="nw"></i><i data-h="ne"></i><i data-h="sw"></i><i data-h="se"></i></div>
  </div>
  <div class="we-cpane" id="weCPane" aria-label="הערות"></div>
</div></div>
<div class="we-status"><span id="wePages"></span><button type="button" class="we-wc" id="weWords" title="ספירת מילים"></button><span class="grow"></span><span class="we-hint" id="weHint"></span>
  <button type="button" class="wb" data-cmd="zout" aria-label="הקטנה">−</button><input type="range" id="weZoom" min="40" max="200" step="10" value="100" aria-label="זום"><button type="button" class="wb" data-cmd="zin" aria-label="הגדלה">+</button><span id="weZoomS">100%</span></div>
<div class="we-find" id="weFind" hidden role="dialog" aria-label="חיפוש והחלפה">
  <div class="fr"><input id="weFindIn" placeholder="חיפוש במסמך" aria-label="חיפוש" dir="auto"><span id="weFindN" class="fn"></span>
    <button type="button" class="wb" id="weFindPrev" title="הקודם" aria-label="הקודם">▲</button><button type="button" class="wb" id="weFindNext" title="הבא" aria-label="הבא">▼</button><button type="button" class="wb" id="weFindX" aria-label="סגירה">×</button></div>
  <div class="fr rep"><input id="weRepIn" placeholder="להחליף ב…" aria-label="החלפה ב" dir="auto"><button type="button" class="we-tb sm" id="weRep1">החלפה</button><button type="button" class="we-tb sm" id="weRepAll">החלפת הכל</button></div>
</div>
<div class="we-pop" id="wePop" hidden></div>
<div class="we-dlg" id="weDlg" hidden><div class="we-dlg-in" role="dialog" aria-modal="true"></div></div>
<input type="file" id="weImgFile" accept="image/*" hidden>
<div id="weStyles" hidden></div>`;
}
const CSS=`
@font-face{font-family:"David";font-weight:400;src:local("David"),url(font-david-400.ttf) format("truetype")}
@font-face{font-family:"David";font-weight:700;src:local("David Bold"),local("David-Bold"),url(font-david-700.ttf) format("truetype")}
@font-face{font-family:"FrankRuehl";font-weight:400;src:local("FrankRuehl"),url(font-frank-400.ttf) format("truetype")}
@font-face{font-family:"FrankRuehl";font-weight:700;src:local("FrankRuehl Bold"),url(font-frank-700.ttf) format("truetype")}
@font-face{font-family:"Frank Ruehl";font-weight:400;src:local("FrankRuehl"),url(font-frank-400.ttf) format("truetype")}
@font-face{font-family:"Frank Ruehl";font-weight:700;src:local("FrankRuehl Bold"),url(font-frank-700.ttf) format("truetype")}
@font-face{font-family:"Heebo";font-weight:400;src:local("Heebo"),url(font-heebo-400.ttf) format("truetype")}
@font-face{font-family:"Heebo";font-weight:700;src:local("Heebo Bold"),url(font-heebo-700.ttf) format("truetype")}
@font-face{font-family:"Rubik";font-weight:400;src:local("Rubik"),url(font-rubik-400.ttf) format("truetype")}
@font-face{font-family:"Rubik";font-weight:700;src:local("Rubik Bold"),url(font-rubik-700.ttf) format("truetype")}
@font-face{font-family:"Noto Sans Hebrew";font-weight:400;src:local("Noto Sans Hebrew"),url(font-noto-400.ttf) format("truetype")}
@font-face{font-family:"Noto Sans Hebrew";font-weight:700;src:local("Noto Sans Hebrew Bold"),url(font-noto-700.ttf) format("truetype")}
#wed{--wn:var(--navy,#16213A);--wn2:var(--navy-2,#202D4A);--wm:var(--marker,#FFD43B);--wb:var(--blue,#2E58CC);--wl:var(--line,#D4DBE5);--wt:var(--ink,#16213A);--wmu:var(--muted,#5D6982);
  position:fixed;inset:0;z-index:900;display:flex;flex-direction:column;background:#E4E8EF;color:var(--wt);direction:rtl;font:14px/1.3 "SF-heebo","Heebo",system-ui,sans-serif;-webkit-user-select:none;user-select:none}
#wed[hidden]{display:none}
#wed button{font:inherit;color:inherit}
#wed .grow{flex:1}
.we-top{display:flex;align-items:center;gap:8px;padding:7px 12px;background:var(--wn);color:#E4EAF4;min-height:48px}
.we-brand{display:flex;align-items:center;gap:7px;white-space:nowrap}
.we-brand b{font-size:17px;color:#fff;background:linear-gradient(transparent 58%,rgba(255,212,59,.55) 58%);padding:0 2px}
.we-mark{width:24px;height:24px;border-radius:5px;background:#2B579A;color:#fff;display:grid;place-items:center;font:700 13px Georgia,serif}
.we-sub{color:#9AA7BE;font-size:13px}
.we-namebox{direction:ltr;display:flex;align-items:center;max-width:340px;min-width:90px;flex:0 1 340px;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.14);border-radius:7px;padding:0 8px;margin-inline-start:10px}
.we-name{flex:1;min-width:0;background:none;border:0;color:#fff;font:inherit;padding:6px 0;outline:none}
.we-namebox:focus-within{border-color:var(--wm)}
.we-ext{color:#9AA7BE;font-size:12px;direction:ltr}
.we-state{font-size:12px;color:#9AA7BE;white-space:nowrap}
.we-state.dirty{color:var(--wm)}
.we-tb{border:1px solid rgba(255,255,255,.22);background:rgba(255,255,255,.06);color:#fff!important;border-radius:8px;padding:7px 13px;cursor:pointer;white-space:nowrap}
.we-tb:hover{background:rgba(255,255,255,.14)}
.we-tb.pri{background:var(--wm);border-color:var(--wm);color:var(--wn)!important;font-weight:700}
.we-tb.pri:hover{filter:brightness(1.05)}
.we-tb.sm{padding:4px 10px;color:var(--wt)!important;border-color:var(--wl);background:#fff}
.we-x{width:34px;height:34px;border-radius:8px;border:0;background:none;color:#C9D2E1!important;font-size:22px;cursor:pointer}
.we-x:hover{background:rgba(255,255,255,.12);color:#fff!important}
.we-tabs{display:flex;gap:2px;padding:0 10px;background:var(--wn)}
#wed .we-tabs button{border:0;background:none;color:#C9D2E1;padding:7px 14px 8px;border-radius:8px 8px 0 0;cursor:pointer;position:relative}
#wed .we-tabs button[aria-selected=true]{background:#F7F8FB;color:var(--wn);font-weight:600}
.we-tabs button[aria-selected=true]::after{content:"";position:absolute;inset-inline:12px;bottom:3px;height:3px;border-radius:2px;background:var(--wm)}
#wed .we-tabs button.ctx{color:var(--wm)}
.we-ribbon{background:#F7F8FB;border-bottom:1px solid var(--wl);overflow-x:auto;overflow-y:hidden;scrollbar-width:thin}
#wed .wpanel{display:flex;align-items:stretch;min-height:92px;padding:6px 6px 2px;width:max-content}
#wed .wpanel[hidden]{display:none}
#wed .wg{display:flex;flex-direction:column;padding:0 8px;border-inline-start:1px solid #E1E6EE}
#wed .wg:first-child{border-inline-start:0}
#wed .wg-b{flex:1;display:flex;align-items:center;gap:3px}
#wed .wg-l{font-size:11px;color:var(--wmu);text-align:center;padding:2px 0 3px;white-space:nowrap}
#wed .wg .row{display:flex;align-items:center;gap:2px}
#wed .wg-b:has(.row){flex-direction:column;align-items:stretch;justify-content:center;gap:4px}
#wed .wg .col{display:flex;flex-direction:column;gap:1px}
#wed .wb{min-width:30px;height:30px;border:1px solid transparent;border-radius:6px;background:none;cursor:pointer;display:inline-grid;place-items:center;padding:0 4px;position:relative}
#wed .wb svg{width:19px;height:19px}
#wed .wb:hover{background:#E8ECF3}
#wed .wb[aria-pressed=true],#wed .wb.on{background:#FFF3C4;border-color:#F2D46B}
#wed .wb:disabled{opacity:.35;cursor:default}
#wed .wb.big{flex-direction:column;height:auto;min-height:62px;min-width:54px;gap:3px;padding:4px 6px;font-size:12px;display:inline-flex;align-items:center;justify-content:center}
#wed .wb.big svg{width:24px;height:24px}
#wed .wb.txt{font-size:12px;padding:0 8px}
#wed .wb .gl{font:700 16px/1 Georgia,"Times New Roman",serif}
#wed .wb .gl.a{font:600 16px/1 Arial,sans-serif;margin-top:-4px}
#wed .wb .cbar{position:absolute;inset-inline:6px;bottom:4px;height:4px;border-radius:1px;background:#C00000}
#wed .wb .cbar[data-bar=hilite]{background:#FFFF00}
#wed .wb .cbar[data-bar=shade]{background:#D9E2F3}
#wed .wsplit{display:inline-flex}
#wed .wsplit .wb:first-child{border-start-end-radius:0;border-end-end-radius:0;min-width:28px}
#wed .wsplit .wb:last-child{min-width:16px;padding:0;border-start-start-radius:0;border-end-start-radius:0}
#wed .wsplit .wb:last-child svg{width:12px;height:12px}
#wed .wsel{height:28px;border:1px solid #CBD3DF;border-radius:6px;background:#fff;font:inherit;font-size:13px;padding:0 6px;color:var(--wt)}
#wed .wsel.font{width:150px}
#wed .wsel.size{width:54px;direction:ltr;text-align:center}
#wed .wsel:focus{outline:2px solid var(--wb);outline-offset:-1px}
#wed .wgal{display:flex;gap:4px;max-width:480px;overflow-x:auto;padding:2px;scrollbar-width:thin}
#wed .wgal button{flex:none;width:86px;height:58px;border:1px solid #DCE2EB;border-radius:7px;background:#fff;cursor:pointer;overflow:hidden;padding:4px 6px;text-align:start;display:flex;flex-direction:column;justify-content:space-between}
#wed .wgal button:hover{border-color:#9FB2D6}
#wed .wgal button.on{border-color:var(--wb);box-shadow:0 0 0 1px var(--wb) inset}
#wed .wgal .smp{height:30px;overflow:hidden;white-space:nowrap}
#wed .wgal .smp p{margin:0!important;padding:0!important;border:0!important;text-indent:0!important;white-space:nowrap;zoom:.62;line-height:1.2!important;min-height:0}
#wed .wgal .smp p::before{display:none!important}
#wed .wgal small{font-size:11px;color:var(--wmu);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
#wed .num{display:flex;align-items:center;gap:5px;font-size:12px;color:var(--wmu);margin:2px 0}
#wed .num input{width:58px;height:26px;border:1px solid #CBD3DF;border-radius:6px;padding:0 5px;font:inherit;color:var(--wt)}
#wed .wg-b:has(.num){flex-direction:column;align-items:stretch;justify-content:center}
#wed .zv{min-width:44px;text-align:center;font-size:13px}
.we-scroll{flex:1;overflow:auto;position:relative;-webkit-user-select:text;user-select:text;overscroll-behavior:contain}
.we-canvas{padding:26px 24px 60px;width:max-content;min-width:100%;box-sizing:border-box;display:flex;justify-content:center}
.we-sheet{background:#fff;position:relative;box-sizing:border-box;box-shadow:0 1px 0 rgba(22,33,58,.05),0 12px 30px -14px rgba(22,33,58,.4);flex:none}
.we-body{outline:none;min-height:100%;position:relative;z-index:1;color:#000;caret-color:#000;overflow-wrap:break-word;line-height:normal;text-align:start}
.we-body:empty::before,.we-body>p:only-child:empty::before{content:"מתחילים לכתוב כאן…";color:#A4AEC0}
.we-body img{max-width:100%}
.we-body a{cursor:text}
.we-body a[href]{color:inherit}
.we-body table{border-collapse:collapse}
.we-body td,.we-body th{min-width:18px}
.we-body .we-atom{cursor:default}
.we-body .we-atom img{cursor:pointer}
.we-body .we-pb{display:block;height:0;border-top:1px dashed #8E9AB3;margin:12px 0;position:relative;text-indent:0;clear:both}
.we-body .we-pb::after{content:"מעבר עמוד";position:absolute;left:50%;transform:translate(-50%,-55%);background:#fff;padding:0 8px;font:11px "SF-heebo",Heebo,sans-serif;color:#7A879E}
.we-body .we-tab{display:inline-block;width:36pt;white-space:pre}
.we-body hr{border:0;border-top:1px solid #000;margin:6pt 0}
.we-body ::selection{background:rgba(46,88,204,.24)}
.we-marks .we-body p::after,.we-marks .we-body li::after{content:"¶";color:#8E9AB3;font:400 11pt Arial,sans-serif}
.we-marks .we-body .we-tab::before{content:"→";color:#8E9AB3;font:9pt Arial,sans-serif}
.we-hdr,.we-ftr{position:absolute;opacity:.5;overflow:hidden;min-height:22px;cursor:text;z-index:3;border-radius:3px}
.we-hdr:hover,.we-ftr:hover{outline:1px dashed rgba(46,88,204,.35);outline-offset:3px}
#wed .we-hdr.ed,#wed .we-ftr.ed{opacity:1;overflow:visible;max-height:none!important;outline:1px dashed #2E58CC;outline-offset:4px;z-index:4}
#wed .we-hdr.ed>*,#wed .we-ftr.ed>*{outline:none;caret-color:#000;-webkit-user-select:text;user-select:text}
#wed .we-hdr.ed::before,#wed .we-ftr.ed::before{position:absolute;inset-inline-start:-4px;font:600 11px "SF-heebo",Heebo,sans-serif;color:#fff;background:#2E58CC;padding:1px 7px;border-radius:3px;white-space:nowrap}
#wed .we-hdr.ed::before{content:"כותרת עליונה";bottom:-24px}#wed .we-ftr.ed::before{content:"כותרת תחתונה";top:-24px}
#wed.we-hfmode .we-body,#wed.we-hfmode .we-notes{opacity:.42}
#wed .we-fld{background:#E3E7EE;border-radius:2px;padding:0 1px}
#wed .we-cm{display:inline;width:0;font-size:0}
#wed .we-cm[data-cs]::before{content:"";display:inline-block;width:0;height:1.05em;border-inline-start:2px solid #E3A008;vertical-align:-.15em}
::highlight(we-cmt){background-color:#FFF1C2}::highlight(we-cmt-on){background-color:#FFD76A}
#wed .we-canvas{gap:18px}
#wed .we-cpane{position:relative;width:270px;flex:none;display:none;align-self:flex-start;-webkit-user-select:none;user-select:none}
#wed.we-hascm .we-cpane{display:block}
#wed .cmc{position:absolute;inset-inline:0;background:#fff;border:1px solid var(--wl);border-inline-start:3px solid #E3A008;border-radius:8px;padding:8px 10px;box-shadow:0 4px 14px -8px rgba(22,33,58,.35);font-size:13px;cursor:pointer;transition:top .15s}
#wed .cmc.on{box-shadow:0 0 0 2px #E3A008,0 8px 22px -10px rgba(22,33,58,.45)}
#wed .cmc.done{opacity:.62;border-inline-start-color:#7A879E}
#wed .cmd{font-size:11px;color:#2E7D32;margin-bottom:3px}
#wed .cmh{display:flex;align-items:center;gap:6px;margin-bottom:4px}#wed .cmh b{font-size:12.5px}#wed .cmh small{color:var(--wmu);font-size:11px;margin-inline-start:auto}
#wed .av{display:inline-grid;place-items:center;width:22px;height:22px;border-radius:50%;background:#2E58CC;color:#fff;font:700 10px/1 sans-serif;font-style:normal;flex:none}
#wed .cmx{white-space:normal;overflow-wrap:anywhere;line-height:1.4;color:#16213A}
#wed .cmr{border-top:1px solid #EEF1F5;margin-top:7px;padding-top:7px}
#wed .cmt{width:100%;box-sizing:border-box;border:1px solid #CBD3DF;border-radius:6px;padding:5px 7px;font:inherit;resize:vertical;min-height:44px;-webkit-user-select:text;user-select:text}
#wed .cmt:focus{outline:2px solid var(--wb);outline-offset:-1px}
#wed .cma{display:flex;gap:4px;justify-content:flex-end;margin-top:6px;flex-wrap:wrap}
#wed .cma .wb.txt{height:26px;padding:0 8px;font-size:12px;width:auto;border-radius:6px}
#wed .cma .wb.pri{background:var(--wb);color:#fff}
#wed .we-toc{display:flex;align-items:baseline}#wed .we-toc>span:first-child{flex:none;max-width:85%}
#wed .we-toc .we-lead{flex:1;border-bottom:1.5px dotted #8E9AB3;margin:0 4px;width:auto;min-width:20px;color:transparent;height:.6em}
#wed .we-toc:hover{background:#F1F5FC;cursor:pointer}
#wed .we-wc{border:0;background:none;color:inherit;font:inherit;padding:0 4px;cursor:pointer;border-radius:4px}#wed .we-wc:hover{background:#E4E8EF}
#wed table.wc{border-collapse:collapse;width:100%;margin:6px 0 4px;font-size:13.5px}#wed table.wc td,#wed table.wc th{padding:5px 8px;border-bottom:1px solid #EEF1F5;text-align:start}#wed table.wc th{color:var(--wmu);font-weight:600;font-size:12px}
@media (max-width:900px){#wed .we-canvas{flex-direction:column;align-items:center}#wed .we-cpane{width:min(100%,520px);align-self:center}#wed .cmc{position:relative;top:auto!important;margin-bottom:8px}#wed .we-cpane{height:auto!important}}
.we-hdr header,.we-ftr footer{margin:0!important;min-height:0!important}
.we-notes{border-top:1px solid #D6DCE6;margin-top:18px;padding-top:6px;font-size:10pt;opacity:.8}
.we-notes:empty{display:none}
.we-guides{position:absolute;inset:0;pointer-events:none;z-index:2}
.we-guides i{position:absolute;inset-inline:0;border-top:1px dashed rgba(46,88,204,.35)}
.we-guides i b{position:absolute;inset-inline-end:calc(100% + 8px);top:-9px;font:600 11px "SF-heebo",Heebo,sans-serif;color:#7A879E;white-space:nowrap}
.we-noguides .we-guides{display:none}
.we-imgbox{position:absolute;z-index:3;outline:2px solid var(--wb);pointer-events:none}
.we-imgbox i{position:absolute;width:10px;height:10px;background:#fff;border:2px solid var(--wb);border-radius:2px;pointer-events:auto}
.we-imgbox i[data-h=nw]{top:-6px;left:-6px;cursor:nwse-resize}.we-imgbox i[data-h=ne]{top:-6px;right:-6px;cursor:nesw-resize}
.we-imgbox i[data-h=sw]{bottom:-6px;left:-6px;cursor:nesw-resize}.we-imgbox i[data-h=se]{bottom:-6px;right:-6px;cursor:nwse-resize}
.we-status{display:flex;align-items:center;gap:14px;padding:3px 12px;background:#F7F8FB;border-top:1px solid var(--wl);font-size:12px;color:var(--wmu);min-height:28px}
#wed .we-status .wb{min-width:22px;height:22px}
.we-status input[type=range]{width:110px;accent-color:var(--wb)}
.we-hint{color:#2E58CC}
.we-find{position:absolute;top:150px;inset-inline-start:24px;z-index:20;background:#fff;border:1px solid var(--wl);border-radius:10px;box-shadow:0 12px 30px -10px rgba(22,33,58,.35);padding:8px;display:flex;flex-direction:column;gap:6px;width:330px}
.we-find[hidden]{display:none}
.we-find .fr{display:flex;align-items:center;gap:4px}
.we-find input{flex:1;min-width:0;height:30px;border:1px solid #CBD3DF;border-radius:6px;padding:0 8px;font:inherit}
.we-find .fn{font-size:12px;color:var(--wmu);min-width:46px;text-align:center}
.we-find:not(.rep-on) .rep{display:none}
.we-pop{position:fixed;z-index:30;max-height:calc(100vh - 16px);overflow:auto;box-sizing:border-box;background:#fff;border:1px solid var(--wl);border-radius:10px;box-shadow:0 14px 34px -12px rgba(22,33,58,.4);padding:8px;max-width:340px}
.we-pop[hidden]{display:none}
.we-pop h4{margin:2px 2px 6px;font-size:12px;color:var(--wmu);font-weight:600}
.we-pop .pal{display:grid;grid-template-columns:repeat(10,22px);gap:4px}
.we-pop .pal button{width:22px;height:22px;border-radius:4px;border:1px solid rgba(0,0,0,.15);cursor:pointer;padding:0}
.we-pop .pal button:hover{transform:scale(1.15)}
.we-pop .mi{display:flex;align-items:center;gap:8px;width:100%;border:0;background:none;padding:7px 10px;border-radius:7px;cursor:pointer;text-align:start;white-space:nowrap}
.we-pop .mi:hover,.we-pop .mi.on{background:#EEF2F8}
.we-pop .mi small{color:var(--wmu);margin-inline-start:auto;font-size:12px;padding-inline-start:18px}
.we-pop .sep{height:1px;background:var(--wl);margin:5px 0}
.we-pop .tgrid{display:grid;grid-template-columns:repeat(10,18px);gap:3px;direction:rtl}
.we-pop .tgrid i{width:18px;height:18px;border:1px solid #C6CFDD;border-radius:2px;cursor:pointer;display:block}
.we-pop .tgrid i.on{background:#FFF0B3;border-color:#E3B92B}
.we-pop .tlab{text-align:center;font-size:12px;color:var(--wmu);margin-top:6px}
.we-pop .syms{display:grid;grid-template-columns:repeat(8,32px);gap:3px}
.we-pop .syms button{height:32px;border:1px solid #E1E6EE;border-radius:6px;background:#fff;cursor:pointer;font-size:16px}
.we-pop .syms button:hover{background:#FFF3C4}
.we-pop .lnk{border:0;background:none;color:var(--wb);cursor:pointer;padding:6px 2px;font-size:13px}
.we-dlg{position:absolute;inset:0;z-index:40;background:rgba(22,33,58,.35);display:grid;place-items:center}
.we-dlg[hidden]{display:none}
.we-dlg-in{background:#fff;border-radius:14px;padding:18px 20px;width:min(420px,92vw);box-shadow:0 24px 60px -20px rgba(0,0,0,.45);-webkit-user-select:text;user-select:text}
.we-dlg-in h3{margin:0 0 10px;font-size:17px}
.we-dlg-in label{display:block;font-size:13px;color:var(--wmu);margin:8px 0 3px}
.we-dlg-in input{width:100%;box-sizing:border-box;height:36px;border:1px solid #CBD3DF;border-radius:8px;padding:0 10px;font:inherit}
.we-dlg-in p{color:var(--wmu);font-size:13px;line-height:1.5;margin:4px 0 0}
.we-dlg-in .acts{display:flex;gap:8px;justify-content:flex-end;margin-top:16px}
.we-dlg-in .acts button{border-radius:8px;padding:8px 16px;border:1px solid var(--wl);background:#fff;cursor:pointer}
.we-dlg-in .acts .pri{background:var(--wn);color:#fff!important;border-color:var(--wn)}
.we-dlg-in .acts .warn{background:#C0392B;color:#fff!important;border-color:#C0392B}
.we-load{position:absolute;inset:0;display:grid;place-items:center;color:var(--wmu);font-size:15px;z-index:5;background:#E4E8EF}
@media (max-width:760px){.we-sub,.we-state,.wg-l{display:none}.we-tb{padding:6px 9px;font-size:13px}#weDesk{display:none}.we-namebox{margin-inline-start:4px}.we-canvas{padding:12px 8px 40px}.wpanel{min-height:76px}.we-find{top:auto;bottom:40px;inset-inline:8px;width:auto}}
@media print{
  html>body>*:not(#wed){display:none!important}
  html>body>#wed{position:static!important;display:block!important;background:#fff!important;inset:auto!important}
  #wed>*:not(.we-scroll){display:none!important}
  .we-scroll{overflow:visible}.we-canvas{padding:0;display:block}
  .we-sheet{box-shadow:none!important;padding:0!important;width:auto!important;min-height:0!important;zoom:1!important}
  .we-hdr,.we-ftr,.we-guides,.we-imgbox{display:none!important}
  .we-body .we-pb{break-after:page;border:0;margin:0}.we-body .we-pb::after{display:none}
  .we-marks .we-body p::after{display:none}
  .we-body{caret-color:transparent!important}
}
#wed .wb{color:#0F1A30}#wed .wg-l{color:#3E4A63!important;font-weight:500}#wed .wb:disabled{opacity:.45}@media (forced-colors:active){#wed .wb svg{forced-color-adjust:auto}}`;

/* ------------------------------------------------------------ runtime */
let E=null;
const ZW=/\u200B/g;
const q=id=>E.root.querySelector('#'+id);
const PARA=/^(P|H[1-6]|LI|PRE|BLOCKQUOTE)$/;
const CELL=/^(TD|TH)$/;
const GENERIC=/^(serif|sans-serif|monospace|cursive|fantasy|system-ui|ui-[a-z-]+|-apple-system|inherit|initial)$/i;
const firstFamily=f=>{for(const x of String(f||'').split(',')){const n=x.trim().replace(/^["']|["']$/g,'');if(n&&!GENERIC.test(n))return n;}return null;};
const cssFont=n=>`"${String(n).replace(/"/g,'')}", "SF-david", "David", serif`;
const isTok=el=>el.nodeType===1&&(el.tagName==='BR'||el.tagName==='IMG'||el.tagName==='HR'||el.dataset.atom!==undefined||el.classList.contains('we-pb')||el.classList.contains('we-tab')||el.classList.contains(P+'-tab-stop')||el.getAttribute('contenteditable')==='false');
const isBlk=el=>el.nodeType===1&&(PARA.test(el.tagName)||CELL.test(el.tagName));
function tokOf(n){let t=null;for(let x=n.nodeType===1?n:n.parentNode;x&&x!==E.body;x=x.parentNode)if(isTok(x))t=x;return t;}
function blockOf(n){for(;n&&n!==E.body;n=n.parentNode)if(n.nodeType===1&&PARA.test(n.tagName))return n;return null;}
function curRange(){const s=getSelection();if(!s||!s.rangeCount)return null;const r=s.getRangeAt(0);return E.body.contains(r.commonAncestorContainer)?r:null;}
function liveRange(){const r=curRange();if(r)return r;if(E.lastRange){E.body.focus({preventScroll:true});const s=getSelection();s.removeAllRanges();s.addRange(E.lastRange);return E.lastRange;}return null;}
// the selection as character positions, so it survives any rebuilding of the paragraph
function offsetOf(node,off){
  const t=tokOf(node);if(t){off=[...t.parentNode.childNodes].indexOf(t);node=t.parentNode;}
  let acc=0,res=null;
  const target=node.nodeType===3?node:(node.childNodes[off]||null),end=node.nodeType===3||node.childNodes[off]?null:node;
  const rec=n=>{for(const c of n.childNodes){
    if(c===target){res=c.nodeType===3?acc+off:acc;return true;}
    if(c.nodeType===3)acc+=c.length;
    else if(c.nodeType===1){if(isTok(c))acc+=1;else{const b=isBlk(c);if(b)acc+=1;if(rec(c))return true;if(c===end){res=acc;return true;}}}
  }return false;};
  rec(E.body);return res===null?acc:res;
}
function pointAt(pos){
  let acc=0,found=null;
  const rec=n=>{const ch=n.childNodes;for(let i=0;i<ch.length;i++){const c=ch[i];
    if(c.nodeType===3){if(pos>=acc&&pos<=acc+c.length){found=[c,pos-acc];return true;}acc+=c.length;}
    else if(c.nodeType===1){
      if(isTok(c)){if(pos===acc){found=[n,i];return true;}acc+=1;}
      else{const b=isBlk(c);if(b)acc+=1;if(rec(c))return true;if(b&&pos===acc){found=[c,c.childNodes.length];return true;}}
    }}return false;};
  rec(E.body);return found||[E.body,E.body.childNodes.length];
}
function getSel(){const r=curRange();if(!r)return E.lastSel||null;return{a:offsetOf(r.startContainer,r.startOffset),b:offsetOf(r.endContainer,r.endOffset)};}
function setSel(s){if(!s)return;const r=document.createRange();const[a,ao]=pointAt(s.a),[b,bo]=pointAt(s.b);
  try{r.setStart(a,ao);r.setEnd(b,bo);}catch(_){return;}
  const sel=getSelection();sel.removeAllRanges();sel.addRange(r);E.lastRange=r.cloneRange();}

/* ---------- history ---------- */
function snap(){clearTimeout(E.snapT);E.snapT=0;const html=E.body.innerHTML;const cur=E.hist[E.hi];
  if(cur&&cur.html===html){cur.sel=getSel();return false;}
  E.hist.splice(E.hi+1);E.hist.push({html,sel:getSel()});if(E.hist.length>80)E.hist.shift();E.hi=E.hist.length-1;return true;}
const snapLater=()=>{clearTimeout(E.snapT);E.snapT=setTimeout(()=>{if(E){snap();markDirty();}},450);};
function undo(){if(E.snapT)snap();if(E.hi<=0){hint('אין עוד מה לבטל');return;}E.hi--;restoreSnap(E.hist[E.hi]);}
function redo(){if(E.hi>=E.hist.length-1)return;E.hi++;restoreSnap(E.hist[E.hi]);}
function restoreSnap(s){E.body.innerHTML=s.html;hideImgBox();setSel(s.sel);after();if(E.cm&&E.cm.size)cmRender();}
function mutate(fn){if(E.snapT)snap();const r=fn();snap();after();return r;}
function after(){markDirty();refreshCounters();schedule();updateUI();}
const cmSig=()=>JSON.stringify([...E.cm.values()].filter(c=>!c.fresh&&(c.keep||cmLive(c))).map(c=>[c.id,c.text,c.done,c.parent]));
function markDirty(){const d=E.main.innerHTML!==E.cleanHtml||E.pageDirty||hfChanged()||cmSig()!==E.cmClean;if(d!==E.dirty){E.dirty=d;}
  const s=q('weState');s.textContent=E.dirty?'שינויים שלא נשמרו':(E.savedOnce?'נשמר':'');s.classList.toggle('dirty',E.dirty);}
let schT=0;function schedule(){clearTimeout(schT);schT=setTimeout(()=>{if(!E)return;if(!E.hf&&!E.composing&&liftBlocks(E.main))snap();guides();words();posImgBox();cmRender();},250);}
function hint(m){const h=q('weHint');h.textContent=m;clearTimeout(E.hintT);E.hintT=setTimeout(()=>{if(E)h.textContent='';},3000);}

/* ---------- character formatting ---------- */
function runProps(el,block){
  const cs=getComputedStyle(el);
  let u=false,s=false,us='solid',bg=null,va=null,base=null;
  for(let n=el;n&&n.nodeType===1;n=n.parentElement){
    const c=n===el?cs:getComputedStyle(n);const d=c.textDecorationLine||'';
    if(d.includes('underline')){u=true;us=c.textDecorationStyle||'solid';}
    if(d.includes('line-through'))s=true;
    if(n!==block&&!bg)bg=hex(c.backgroundColor);
    if(!va&&(c.verticalAlign==='super'||c.verticalAlign==='sub')){va=c.verticalAlign;
      base=n.dataset&&n.dataset.fsb?+n.dataset.fsb:parseFloat(getComputedStyle(n.parentElement||n).fontSize)*.75;}
    if(n===block||n===E.body)break;
  }
  const fs=Math.round(parseFloat(cs.fontSize)*.75*2)/2;
  return{ff:firstFamily(cs.fontFamily),fs:va?Math.round(base*2)/2:fs,b:parseInt(cs.fontWeight,10)>=600,i:cs.fontStyle!=='normal',u,us,s,color:hex(cs.color)||'000000',bg,va};
}
function propsCss(p){
  let c='';if(p.ff)c+=`font-family:${cssFont(p.ff)};`;
  c+=`font-size:${p.va?Math.round(p.fs*.66*10)/10:p.fs}pt;font-weight:${p.b?700:400};font-style:${p.i?'italic':'normal'};`;
  const d=[p.u&&'underline',p.s&&'line-through'].filter(Boolean);if(d.length)c+=`text-decoration-line:${d.join(' ')};`;
  if(p.u&&p.us&&p.us!=='solid')c+=`text-decoration-style:${p.us};`;
  c+=`color:#${p.color};`;if(p.bg)c+=`background-color:#${p.bg};`;if(p.va)c+=`vertical-align:${p.va};`;
  return c;
}
function makeSpan(p,text){const s=document.createElement('span');s.setAttribute('style',propsCss(p));if(p.va)s.dataset.fsb=p.fs;if(text!=null)s.textContent=text;return s;}
const FMT=/^(SPAN|FONT|B|STRONG|I|EM|U|S|STRIKE|DEL|INS|SUB|SUP|MARK|SMALL|BIG|CODE|LABEL|ABBR|CITE|Q|TT|VAR|KBD|SAMP|BDI|BDO|TIME|DATA|O:P)$/;
// one flat row of spans per paragraph, each with its full look written on it
function flatten(container,block){
  const items=[];
  const walk=n=>{for(const c of [...n.childNodes]){
    if(c.nodeType===3){if(c.data)items.push({t:c.data,p:runProps(c.parentElement,block)});}
    else if(c.nodeType===1){
      if(isTok(c))items.push({keep:c});
      else if(c.tagName==='A'){flatten(c,block);items.push({keep:c});}
      else if(FMT.test(c.tagName))walk(c);
      else items.push({keep:c});
    }}};
  walk(container);
  container.textContent='';
  let last=null,lastCss=null;
  for(const it of items){
    if(it.keep){container.append(it.keep);last=null;continue;}
    const css=propsCss(it.p);
    if(last&&css===lastCss)last.firstChild.appendData(it.t);
    else{last=makeSpan(it.p,it.t);lastCss=css;container.append(last);}
  }
}
function mergeSpans(container){
  for(const s of [...container.querySelectorAll('span')]){
    const n=s.nextSibling;
    if(n&&n.nodeType===1&&n.tagName==='SPAN'&&!isTok(n)&&!isTok(s)&&s.parentNode===n.parentNode&&s.getAttribute('style')===n.getAttribute('style')&&s.dataset.fsb===n.dataset.fsb&&s.attributes.length===n.attributes.length){
      while(n.firstChild)s.append(n.firstChild);n.remove();}
  }
  container.normalize();
}
function selBlocks(r){
  r=r||curRange();if(!r)return[];
  const leafs=[...E.body.querySelectorAll('p,h1,h2,h3,h4,h5,h6,li,pre,blockquote')].filter(b=>!b.querySelector('p,h1,h2,h3,h4,h5,h6,li,table'));
  const out=leafs.filter(b=>r.intersectsNode(b));
  if(out.length>1&&!r.collapsed&&r.endOffset===0){const l=out[out.length-1];if(l.contains(r.endContainer)&&!l.contains(r.startContainer))out.pop();}
  return out;
}
function wordAround(r){
  const t=r.startContainer;if(t.nodeType!==3)return null;const s=t.data,o=r.startOffset;const L=/[\p{L}\p{N}\p{M}'"״׳-]/u;
  if(!(o>0&&o<s.length&&L.test(s[o-1])&&L.test(s[o])))return null;
  let a=o,b=o;while(a>0&&L.test(s[a-1]))a--;while(b<s.length&&L.test(s[b]))b++;
  const w=document.createRange();w.setStart(t,a);w.setEnd(t,b);return w;
}
// apply(span, all): changes the look of the selected text. test(span) tells whether the whole selection already has it
function inline(apply,test){
  const r=liveRange();if(!r)return;
  if(r.collapsed){
    const w=wordAround(r);
    if(w){const keep=getSel();const s=getSelection();s.removeAllRanges();s.addRange(w);inline(apply,test);setSel(keep);return;}
    return typing(apply,test);
  }
  mutate(()=>{
    const sel=getSel();const blocks=selBlocks(r);
    if(!blocks.length)return;
    blocks.forEach(b=>flatten(b,b));
    setSel(sel);const rr=curRange();if(!rr)return;
    let sc=rr.startContainer,so=rr.startOffset,ec=rr.endContainer,eo=rr.endOffset;
    if(ec.nodeType===3&&eo>0&&eo<ec.length)ec.splitText(eo);
    if(sc.nodeType===3&&so>0&&so<sc.length){const n=sc.splitText(so);if(ec===sc){ec=n;eo-=so;}sc=n;so=0;}
    rr.setStart(sc,so);rr.setEnd(ec,eo);
    const texts=[];
    for(const b of blocks){const tw=document.createTreeWalker(b,NodeFilter.SHOW_TEXT);let t;
      while((t=tw.nextNode()))if(t.length&&!tokOf(t)&&rr.comparePoint(t,0)>=0&&rr.comparePoint(t,t.length)<=0)texts.push(t);}
    const spans=texts.map(isolate).filter(Boolean);
    if(!spans.length)return;
    const all=test?spans.every(test):false;
    spans.forEach(s=>apply(s,all));
    blocks.forEach(mergeSpans);
    setSel(sel);
  });
}
function isolate(t){
  const p=t.parentNode;
  if(p.nodeType===1&&p.tagName==='SPAN'&&!isTok(p)){
    if(p.childNodes.length===1)return p;
    for(const c of [...p.childNodes]){const ns=p.cloneNode(false);p.parentNode.insertBefore(ns,p);ns.append(c);}
    p.remove();return t.parentNode;
  }
  const b=blockOf(t)||p;const s=makeSpan(runProps(p,b));p.insertBefore(s,t);s.append(t);return s;
}
// no selection: the next letters typed get the new look (like Word)
function typing(apply,test){
  const r=liveRange();if(!r)return;
  let el=r.startContainer.nodeType===3?r.startContainer.parentElement:r.startContainer;
  if(isTok(el))el=el.parentElement;const b=blockOf(el);if(!b)return;
  if(E.snapT)snap();
  let span=el.tagName==='SPAN'&&el.textContent.replace(ZW,'')===''&&el.dataset.typing?el:null;
  if(!span){span=makeSpan(runProps(el,b),'\u200B');span.dataset.typing='1';r.insertNode(span);}
  apply(span,test?test(span):false);
  const t=span.firstChild;const nr=document.createRange();nr.setStart(t,t.length);nr.collapse(true);
  const s=getSelection();s.removeAllRanges();s.addRange(nr);E.lastRange=nr.cloneRange();
  updateUI();
}
const deco=s=>new Set((s.style.textDecorationLine||getComputedStyle(s).textDecorationLine||'').split(/\s+/).filter(x=>x&&x!=='none'));
function setDeco(s,tok,on){const d=deco(s);on?d.add(tok):d.delete(tok);s.style.textDecorationLine=d.size?[...d].join(' '):'none';}
const spanPt=s=>s.dataset.fsb?+s.dataset.fsb:Math.round(parseFloat(getComputedStyle(s).fontSize)*.75*2)/2;
function setPt(s,pt){if(s.dataset.fsb){s.dataset.fsb=pt;s.style.fontSize=Math.round(pt*.66*10)/10+'pt';}else s.style.fontSize=pt+'pt';}
function setVa(s,v){const base=spanPt(s);if(v){s.dataset.fsb=base;s.style.verticalAlign=v;s.style.fontSize=Math.round(base*.66*10)/10+'pt';}
  else{delete s.dataset.fsb;s.style.verticalAlign='';s.style.fontSize=base+'pt';}}
const CH={
  bold:[(s,a)=>s.style.fontWeight=a?'400':'700',s=>parseInt(getComputedStyle(s).fontWeight,10)>=600],
  italic:[(s,a)=>s.style.fontStyle=a?'normal':'italic',s=>getComputedStyle(s).fontStyle!=='normal'],
  underline:[(s,a)=>setDeco(s,'underline',!a),s=>deco(s).has('underline')],
  strike:[(s,a)=>setDeco(s,'line-through',!a),s=>deco(s).has('line-through')],
  sup:[(s,a)=>setVa(s,a?null:'super'),s=>s.style.verticalAlign==='super'],
  sub:[(s,a)=>setVa(s,a?null:'sub'),s=>s.style.verticalAlign==='sub']
};
function fmt(name){const c=CH[name];if(c)inline(c[0],c[1]);}
function setFont(n){inline(s=>s.style.fontFamily=cssFont(n));E.lastFont=n;}
function setSize(pt){pt=Math.max(1,Math.min(1638,Math.round(pt*2)/2));inline(s=>setPt(s,pt));}
function growFont(dir,step){inline(s=>{const c=spanPt(s);let n;
  if(step)n=c+dir;else if(dir>0)n=SIZES.find(x=>x>c)||c+10;else n=[...SIZES].reverse().find(x=>x<c)||Math.max(1,c-1);setPt(s,n);});}
function setColor(h){E.color=h;q('weRibbon').querySelector('[data-bar=color]').style.background=h?'#'+h:'#000';inline(s=>s.style.color=h?'#'+h:'#000000');}
function setHilite(h){E.hl=h;q('weRibbon').querySelector('[data-bar=hilite]').style.background=h?'#'+h:'transparent';inline(s=>s.style.backgroundColor=h?'#'+h:'');}
function clearFmt(){inline(s=>{s.removeAttribute('style');delete s.dataset.fsb;});}

/* ---------- paragraphs ---------- */
function ensureBlocks(){
  // text typed straight into the page gets a paragraph of its own
  const b=E.body;let changed=false,sel=null;
  for(const n of [...b.childNodes]){
    if(n.nodeType===3&&!n.data.trim()&&!n.data.includes('\u00a0')){if(n.data==='')n.remove();continue;}
    if(n.nodeType===3||(n.nodeType===1&&!PARA.test(n.tagName)&&!/^(TABLE|UL|OL|HR|DIV)$/.test(n.tagName))){
      if(!sel)sel=getSel();const p=document.createElement('p');p.style.direction=E.st.heb?'rtl':'ltr';b.insertBefore(p,n);
      while(p.nextSibling&&(p.nextSibling.nodeType===3||(!PARA.test(p.nextSibling.tagName)&&!/^(TABLE|UL|OL|HR|DIV)$/.test(p.nextSibling.tagName))))p.append(p.nextSibling);
      changed=true;}
  }
  if(!b.querySelector('p,h1,h2,h3,h4,h5,h6,li,table')){const p=document.createElement('p');p.style.direction=E.st.heb?'rtl':'ltr';p.append(document.createElement('br'));b.append(p);changed=true;}
  if(changed&&sel)setSel(sel);
  return changed;
}
function blocksOr(){let bs=selBlocks();if(!bs.length&&ensureBlocks())bs=selBlocks();if(!bs.length){const r=liveRange();const b=r&&blockOf(r.startContainer);if(b)bs=[b];}return bs;}
const visAlign=cs=>{const t=cs.textAlign,rtl=cs.direction==='rtl';
  if(t==='start'||t==='-webkit-auto'||t==='auto')return rtl?'right':'left';if(t==='end')return rtl?'left':'right';
  if(t==='-webkit-right')return'right';if(t==='-webkit-left')return'left';if(t==='-webkit-center')return'center';return t;};
function align(v){mutate(()=>blocksOr().forEach(b=>b.style.textAlign=v));}
function setDir(d){mutate(()=>blocksOr().forEach(b=>{const cs=getComputedStyle(b),va=visAlign(cs),old=cs.direction;
  b.style.direction=d;if(va===(old==='rtl'?'right':'left'))b.style.textAlign='start';else if(va===(old==='rtl'?'left':'right'))b.style.textAlign='end';}));}
function lineSpacing(v){mutate(()=>blocksOr().forEach(b=>b.style.lineHeight=String(v)));}
function paraSpace(which,pt){mutate(()=>blocksOr().forEach(b=>b.style[which==='before'?'marginTop':'marginBottom']=pt+'pt'));}
function indentPhys(side,cm){mutate(()=>blocksOr().forEach(b=>b.style[side==='r'?'marginRight':'marginLeft']=(Math.max(0,cm)*37.7953).toFixed(1)+'px'));}
const numRe=new RegExp(`^${P}-num-(\\d+)-(\\d+)$`);
function numOf(b){for(const c of b.classList){const m=numRe.exec(c);if(m)return{id:+m[1],lvl:+m[2],cls:c};}return null;}
function styleCls(b){for(const c of b.classList)if(c.startsWith(P+'_'))return c;return null;}
const lpCls=()=>{const s=E.st.styles.find(x=>x.type==='paragraph'&&x.name==='list paragraph');return s?classKey(s.id):null;};
function indent(dir){mutate(()=>blocksOr().forEach(b=>{
  const n=numOf(b);
  if(n){b.classList.remove(n.cls);b.classList.add(`${P}-num-${n.id}-${Math.max(0,Math.min(8,n.lvl+dir))}`);return;}
  const cs=getComputedStyle(b);const cur=parseFloat(cs.direction==='rtl'?cs.marginRight:cs.marginLeft)||0;
  b.style.marginInlineStart=Math.max(0,Math.round((cur+dir*48)/48)*48)+'px';
}));}
function toggleList(kind){mutate(()=>{
  const id=kind==='b'?E.st.num.bullet:E.st.num.decimal;const bs=blocksOr();const all=bs.length&&bs.every(b=>{const n=numOf(b);return n&&n.id===id;});const lp=lpCls();
  bs.forEach(b=>{const n=numOf(b);if(n)b.classList.remove(n.cls);
    if(all){if(lp)b.classList.remove(lp);(b.style.removeProperty('counter-set'),delete b.dataset.ns);}
    else{b.classList.add(`${P}-num-${id}-${n?n.lvl:0}`);if(lp&&!styleCls(b))b.classList.add(lp);}});
});}
function endList(b){const n=numOf(b);if(!n)return false;b.classList.remove(n.cls);const lp=lpCls();if(lp)b.classList.remove(lp);(b.style.removeProperty('counter-set'),delete b.dataset.ns);return true;}
// a new numbered list starts again from 1
function refreshCounters(){
  if(!E)return;const id=E.st.num.decimal;
  for(const b of E.body.querySelectorAll(`[class*="${P}-num-${id}-"]`)){
    const n=numOf(b);if(!n)continue;let p=b.previousElementSibling;const pn=p&&numOf(p);
    const fresh=!(pn&&pn.id===id);const want=fresh?`${P}-num-${id}-${n.lvl} 0${n.lvl<8?` ${P}-num-${id}-${n.lvl+1} 0`:''}`:'';
    if((b.dataset.ns||'')!==want){if(want){b.style.setProperty('counter-set',want);b.dataset.ns=want;}else{b.style.removeProperty('counter-set');delete b.dataset.ns;}}
  }
}
function galItems(){
  const ps=E.st.styles.filter(s=>s.type==='paragraph');const out=[];
  const add=(s,label,normal)=>{if(s&&!out.some(o=>o.id===s.id))out.push({id:s.id,label:label||s.label,cls:classKey(s.id),normal:!!normal});};
  add(ps.find(s=>s.def)||ps.find(s=>s.name==='normal'),'רגיל',true);
  if(!out.length)out.push({id:null,label:'רגיל',cls:null,normal:true});
  BUILTIN.filter(b=>b.label).forEach(b=>add(ps.find(s=>s.name===b.name),b.label));
  ps.filter(s=>s.q&&!s.hidden&&!/^(normal|list paragraph|caption|no spacing|footnote text|endnote text|header|footer)$/.test(s.name)).forEach(s=>add(s));
  return out.slice(0,16);
}
function buildGallery(){
  const g=q('weGal');g.innerHTML='';
  for(const it of galItems()){
    const b=document.createElement('button');b.type='button';b.dataset.style=it.id||'';b.title=it.label;
    const smp=document.createElement('span');smp.className='smp '+P;const p=document.createElement('p');if(!it.normal&&it.cls)p.className=it.cls;
    const sp=document.createElement('span');sp.textContent=it.label.length>14?it.label.slice(0,14)+'…':it.label;p.append(sp);smp.append(p);
    smp.style.fontFamily=E.body.style.fontFamily;smp.style.fontSize=E.body.style.fontSize;
    const sm=document.createElement('small');sm.textContent=it.label;b.append(smp,sm);
    b.addEventListener('mousedown',e=>e.preventDefault());b.onclick=()=>applyStyle(it);g.append(b);
  }
}
function applyStyle(it){mutate(()=>{
  for(let b of blocksOr()){
    if(b.tagName!=='P'){const p=document.createElement('p');for(const a of b.attributes)p.setAttribute(a.name,a.value);while(b.firstChild)p.append(b.firstChild);b.replaceWith(p);b=p;}
    for(const c of [...b.classList])if(c.startsWith(P+'_'))b.classList.remove(c);
    if(!it.normal&&it.cls)b.classList.add(it.cls);
    for(const s of b.querySelectorAll('span,font,b,strong,i,em')){for(const pr of ['font-size','font-family','color','font-weight','font-style'])s.style.removeProperty(pr);delete s.dataset.fsb;}
    for(const t of b.querySelectorAll('b,strong,i,em,font'))t.replaceWith(...t.childNodes);
  }
});}
function isHeading(b){const c=styleCls(b);if(!c)return /^H\d$/.test(b.tagName);
  const s=E.st.styles.find(x=>classKey(x.id)===c);return !!s&&/^(heading \d|title|subtitle)$/.test(s.name);}

/* ---------- insertions ---------- */
function curBlock(){const r=liveRange();return r?blockOf(r.startContainer):null;}
const contentW=()=>px(E.page.w)-px(E.page.left)-px(E.page.right);
function newPara(dir){const p=document.createElement('p');p.style.direction=dir||(E.st.heb?'rtl':'ltr');p.append(document.createElement('br'));return p;}
function caretTo(node,off){const r=document.createRange();r.setStart(node,off||0);r.collapse(true);const s=getSelection();s.removeAllRanges();s.addRange(r);E.lastRange=r.cloneRange();}
function caretIn(block){const tw=document.createTreeWalker(block,NodeFilter.SHOW_TEXT);const t=tw.nextNode();if(t)caretTo(t,0);else caretTo(block,0);}
// splits the paragraph at the caret: [before, after]
function splitAtCaret(){
  const r=liveRange();const b=r&&blockOf(r.startContainer);if(!b)return null;
  if(!r.collapsed)r.deleteContents();
  const tail=document.createRange();tail.setStart(r.startContainer,r.startOffset);tail.setEnd(b,b.childNodes.length);
  const frag=tail.extractContents();const b2=b.cloneNode(false);b2.append(frag);
  b.after(b2);
  for(const x of [b,b2])if(!x.textContent.replace(ZW,'')&&!x.querySelector('img,[data-atom],.we-pb,br'))x.append(document.createElement('br'));
  return[b,b2];
}
function insertBlockAfterCaret(el){
  const parts=splitAtCaret();
  if(!parts){E.body.append(el);return;}
  const[a,b]=parts;
  const aEmpty=!a.textContent.replace(ZW,'').trim()&&!a.querySelector('img,[data-atom]');
  a.after(el);
  if(aEmpty&&a.previousElementSibling)a.remove();
  return b;
}
function insertTable(rows,cols){mutate(()=>{
  const b=curBlock();const dir=b?getComputedStyle(b).direction:(E.st.heb?'rtl':'ltr');
  const w=Math.floor(contentW()),cw=Math.floor(w/cols);
  const t=document.createElement('table');t.style.cssText=`width:${w}px;border-collapse:collapse;direction:${dir}`;
  const tb=document.createElement('tbody');t.append(tb);
  for(let i=0;i<rows;i++){const tr=document.createElement('tr');for(let j=0;j<cols;j++){const td=document.createElement('td');
    td.style.cssText=`border:0.5pt solid #000;width:${cw}px;padding:0 5.4pt;vertical-align:top`;td.append(newPara(dir));tr.append(td);}tb.append(tr);}
  const after=insertBlockAfterCaret(t);
  if(!after||after===t)t.after(newPara(dir));
  caretIn(t.querySelector('p'));
});}
function insertHr(){mutate(()=>{const hr=document.createElement('hr');const nb=insertBlockAfterCaret(hr);if(nb)caretIn(nb);});}
function insertPageBreak(){mutate(()=>{
  const b=curBlock();const p=document.createElement('p');p.style.direction=b?getComputedStyle(b).direction:'rtl';
  const a=document.createElement('a');a.className='we-pb';a.contentEditable='false';p.append(a);
  const nb=insertBlockAfterCaret(p);if(nb)caretIn(nb);
});}
function insertText(t){const r=liveRange();if(!r)return;document.execCommand('insertText',false,t);}
async function insertImageFile(file){
  if(!file||!/^image\//.test(file.type))return;
  const url=URL.createObjectURL(file);E.blobs.add(url);
  const img=new Image();img.src=url;try{await img.decode();}catch(_){E.opts.toast('לא הצלחתי לקרוא את התמונה');return;}
  let w=img.naturalWidth,h=img.naturalHeight;const mw=contentW();if(w>mw){h=h*mw/w;w=mw;}
  mutate(()=>{
    const el=document.createElement('img');el.src=url;el.alt='';el.dataset.new='1';el.style.width=Math.round(w)+'px';el.style.height=Math.round(h)+'px';
    const r=liveRange();
    if(r&&blockOf(r.startContainer)){r.deleteContents();r.insertNode(el);caretTo(el.parentNode,[...el.parentNode.childNodes].indexOf(el)+1);}
    else{const p=newPara();p.textContent='';p.append(el);E.body.append(p);}
  });
}

/* ---------- tables ---------- */
const cellOf=()=>{const r=liveRange();if(!r)return null;let n=r.startContainer;n=n.nodeType===3?n.parentElement:n;const c=n&&n.closest('td,th');return c&&E.body.contains(c)?c:null;};
function emptyCellLike(td){const n=td.cloneNode(false);n.removeAttribute('rowspan');const p=td.querySelector('p');const np=p?p.cloneNode(false):document.createElement('p');np.append(document.createElement('br'));n.append(np);return n;}
function tableOp(op){
  const td=cellOf();if(!td)return;const tr=td.parentElement,tbl=td.closest('table');
  const rows=[...tbl.rows].filter(r=>r.closest('table')===tbl);const ci=td.cellIndex;const rtl=getComputedStyle(tbl).direction==='rtl';
  mutate(()=>{
    tbl.querySelectorAll(':scope>colgroup').forEach(c=>c.remove());
    if(op==='rowAbove'||op==='rowBelow'){const nr=tr.cloneNode(false);for(const c of tr.cells)nr.append(emptyCellLike(c));op==='rowAbove'?tr.before(nr):tr.after(nr);caretIn(nr.cells[Math.min(ci,nr.cells.length-1)]);}
    else if(op==='colRight'||op==='colLeft'){const before=(op==='colRight')===rtl;
      for(const row of rows){const ref=row.cells[Math.min(ci,row.cells.length-1)];if(!ref)continue;const n=emptyCellLike(ref);n.removeAttribute('colspan');before?ref.before(n):ref.after(n);}
      const w=tbl.offsetWidth;const cnt=Math.max(...rows.map(r=>r.cells.length));for(const row of rows)for(const c of row.cells)c.style.width=Math.floor(w/cnt)+'px';}
    else if(op==='delRow'){const nx=tr.nextElementSibling||tr.previousElementSibling;tr.remove();if(!tbl.rows.length)removeTable(tbl);else if(nx)caretIn(nx.cells[0]);}
    else if(op==='delCol'){for(const row of rows){const c=row.cells[Math.min(ci,row.cells.length-1)];if(c){const s=+c.getAttribute('colspan')||1;if(s>1)c.setAttribute('colspan',s-1);else c.remove();}}
      for(const row of [...tbl.rows])if(!row.cells.length)row.remove();if(!tbl.rows.length||!tbl.querySelector('td,th'))removeTable(tbl);}
    else if(op==='delTable')removeTable(tbl);
    else if(op==='borders'||op==='noborders'){for(const c of tbl.querySelectorAll('td,th'))if(c.closest('table')===tbl)c.style.border=op==='borders'?'0.5pt solid #000':'none';}
  });
}
function removeTable(tbl){const p=newPara(getComputedStyle(tbl).direction);tbl.replaceWith(p);caretIn(p);}
function selCells(){const r=liveRange();if(!r)return[];const c=cellOf();const tbl=c&&c.closest('table');if(!tbl)return[];
  const cs=[...tbl.querySelectorAll('td,th')].filter(x=>x.closest('table')===tbl&&r.intersectsNode(x));return cs.length?cs:[c];}
function shade(h){E.shade=h;q('weRibbon').querySelector('[data-bar=shade]').style.background=h?'#'+h:'transparent';
  const cs=selCells();if(!cs.length)return;mutate(()=>cs.forEach(c=>c.style.backgroundColor=h?'#'+h:''));}
function nextCell(back){
  const td=cellOf();if(!td)return false;const tbl=td.closest('table');
  const cells=[...tbl.querySelectorAll('td,th')].filter(x=>x.closest('table')===tbl);let i=cells.indexOf(td)+(back?-1:1);
  if(i<0)return true;
  if(i>=cells.length){mutate(()=>{const tr=td.parentElement;const nr=tr.cloneNode(false);for(const c of tr.cells)nr.append(emptyCellLike(c));tr.after(nr);caretIn(nr.cells[0]);});return true;}
  const c=cells[i];const r=document.createRange();r.selectNodeContents(c.querySelector('p')||c);const s=getSelection();s.removeAllRanges();s.addRange(r);return true;
}

/* ---------- pictures: select and resize ---------- */
function offIn(el,anc){let x=0,y=0;for(let n=el;n&&n!==anc;n=n.offsetParent){x+=n.offsetLeft;y+=n.offsetTop;if(!n.offsetParent)break;}return{x,y};}
function selectImg(img){E.selImg=img;posImgBox();}
function hideImgBox(){E.selImg=null;q('weImgBox').hidden=true;}
function posImgBox(){const img=E.selImg,box=q('weImgBox');if(!img||!E.body.contains(img)){box.hidden=true;E.selImg=null;return;}
  const o=offIn(img,q('weSheet'));Object.assign(box.style,{left:o.x+'px',top:o.y+'px',width:img.offsetWidth+'px',height:img.offsetHeight+'px'});box.hidden=false;}
function startResize(e){
  const img=E.selImg;if(!img)return;e.preventDefault();e.stopPropagation();
  const h=e.target.dataset.h,x0=e.clientX,w0=img.offsetWidth,h0=img.offsetHeight,ratio=h0/w0||1;if(E.snapT)snap();
  const right=h==='ne'||h==='se';const mw=contentW();
  const mv=ev=>{const dx=(ev.clientX-x0)/E.zoom;let w=Math.max(16,Math.min(mw,w0+(right?dx:-dx)));const hh=w*ratio;
    img.style.width=Math.round(w)+'px';img.style.height=Math.round(hh)+'px';img.style.maxWidth='none';img.dataset.rs='1';
    const p=img.parentElement;if(p&&p.tagName==='DIV'&&p.closest('[data-atom]')){p.style.width=Math.round(w)+'px';p.style.height=Math.round(hh)+'px';}
    posImgBox();};
  const up=()=>{removeEventListener('pointermove',mv);removeEventListener('pointerup',up);snap();after();};
  addEventListener('pointermove',mv);addEventListener('pointerup',up);
}
function deleteImg(){const img=E.selImg;if(!img)return;mutate(()=>{const a=img.closest('[data-atom]');(a||img).remove();hideImgBox();});}

/* ---------- links ---------- */
function linkDialog(){
  const r=liveRange();if(!r)return;const n=r.startContainer.nodeType===3?r.startContainer.parentElement:r.startContainer;
  const a=n&&n.closest('a[href]');const sel=getSel();
  dialog(`<h3>${a?'עריכת קישור':'הוספת קישור'}</h3>
    <label>הטקסט שיוצג</label><input id="weLT" dir="auto" value="${esc(a?a.textContent:r.toString())}">
    <label>כתובת</label><input id="weLU" dir="ltr" placeholder="https://" value="${esc(a?a.getAttribute('href'):'')}">
    <div class="acts">${a?'<button type="button" data-a="rm">הסרת הקישור</button>':''}<span class="grow"></span><button type="button" data-a="x">ביטול</button><button type="button" class="pri" data-a="ok">אישור</button></div>`,
    (act,root)=>{
      if(act==='x')return true;setSel(sel);
      if(act==='rm'){mutate(()=>{a.replaceWith(...a.childNodes);});return true;}
      let url=root.querySelector('#weLU').value.trim();const text=root.querySelector('#weLT').value;
      if(!url)return false;if(!/^(https?:|mailto:|#|tel:)/i.test(url))url=(/@/.test(url)&&!/\//.test(url)?'mailto:':'https://')+url;
      mutate(()=>{
        if(a){a.setAttribute('href',url);if(text&&text!==a.textContent){a.textContent='';const sp=document.createElement('span');sp.textContent=text;a.append(sp);}return;}
        const rr=curRange()||liveRange();const el=document.createElement('a');el.href=url;el.style.color='#0563C1';el.style.textDecorationLine='underline';
        if(rr.collapsed||text!==rr.toString()){rr.deleteContents();const sp=document.createElement('span');sp.textContent=text||url;el.append(sp);rr.insertNode(el);}
        else{el.append(rr.extractContents());rr.insertNode(el);}
        caretTo(el.parentNode,[...el.parentNode.childNodes].indexOf(el)+1);
      });return true;
    },'#weLU');
}
function dialog(html,onAct,focusSel){
  const d=q('weDlg'),inn=d.firstElementChild;inn.innerHTML=html;d.hidden=false;E.dlgOpen=true;
  const close=()=>{d.hidden=true;E.dlgOpen=false;inn.innerHTML='';};
  inn.onclick=e=>{const b=e.target.closest('[data-a]');if(!b)return;const r=onAct(b.dataset.a,inn);if(r!==false)close();};
  inn.onkeydown=e=>{if(e.key==='Escape'){e.preventDefault();onAct('x',inn);close();}else if(e.key==='Enter'&&e.target.tagName==='INPUT'){e.preventDefault();const ok=inn.querySelector('[data-a=ok]');if(ok)ok.click();}};
  setTimeout(()=>{const f=focusSel&&inn.querySelector(focusSel)||inn.querySelector('input,button.pri');if(f){f.focus();f.select&&f.select();}},30);
  return close;
}
function ask(title,text,btns){return new Promise(res=>{
  dialog(`<h3>${title}</h3><p>${text}</p><div class="acts">${btns.map(b=>`<button type="button" data-a="${b.v}" class="${b.cls||''}">${b.l}</button>`).join('')}</div>`,a=>{res(a==='x'?null:a);return true;});});}

/* ---------- find and replace ---------- */
function findAll(qs){
  const out=[];if(!qs)return out;const needle=qs.toLowerCase();
  const leafs=[...E.body.querySelectorAll('p,h1,h2,h3,h4,h5,h6,li,pre,blockquote,td,th')].filter(b=>!b.querySelector('p,h1,h2,h3,h4,h5,h6,li,td,th'));
  for(const b of leafs){
    const nodes=[];let s='';const tw=document.createTreeWalker(b,NodeFilter.SHOW_TEXT);let t;
    while((t=tw.nextNode())){if(tokOf(t))continue;nodes.push([t,s.length]);s+=t.data;}
    const low=s.toLowerCase();let i=low.indexOf(needle);
    const at=k=>{for(let j=nodes.length-1;j>=0;j--)if(nodes[j][1]<=k)return[nodes[j][0],k-nodes[j][1]];return null;};
    while(i>=0){const a=at(i),z=at(i+needle.length);if(a&&z){const r=document.createRange();r.setStart(a[0],a[1]);
        const ze=z[1]===0&&z[0]!==a[0]?(()=>{const k=nodes.findIndex(n=>n[0]===z[0]);return k>0?[nodes[k-1][0],nodes[k-1][0].length]:z;})():z;r.setEnd(ze[0],ze[1]);out.push(r);}
      i=low.indexOf(needle,i+Math.max(1,needle.length));}
  }
  return out;
}
function paintFind(){
  if(!(window.CSS&&window.CSS.highlights&&window.Highlight))return;
  const F=E.find;window.CSS.highlights.delete('we-find');window.CSS.highlights.delete('we-find-cur');
  if(!F||!F.list.length)return;window.CSS.highlights.set('we-find',new Highlight(...F.list));if(F.list[F.at])window.CSS.highlights.set('we-find-cur',new Highlight(F.list[F.at]));
}
function runFind(keepAt){
  const qs=q('weFindIn').value;const list=findAll(qs);const at=keepAt&&E.find?Math.min(E.find.at,list.length-1):0;
  E.find={q:qs,list,at:Math.max(0,at)};q('weFindN').textContent=qs?(list.length?`${E.find.at+1}/${list.length}`:'0'):'';paintFind();
}
function gotoFind(d){
  if(!E.find||E.find.q!==q('weFindIn').value)runFind();const F=E.find;if(!F.list.length)return;
  F.at=(F.at+d+F.list.length)%F.list.length;const r=F.list[F.at];
  const s=getSelection();s.removeAllRanges();s.addRange(r);E.lastRange=r.cloneRange();
  const rc=r.getBoundingClientRect(),sc=q('weScroll'),sr=sc.getBoundingClientRect();
  if(rc.top<sr.top+60||rc.bottom>sr.bottom-40)sc.scrollTop+=rc.top-sr.top-sr.height/3;
  q('weFindN').textContent=`${F.at+1}/${F.list.length}`;paintFind();
}
function replaceOne(){
  if(!E.find||!E.find.list.length){runFind();return;}const F=E.find;const r=F.list[F.at];const rep=q('weRepIn').value;
  mutate(()=>{replaceRange(r,rep);});runFind(true);if(E.find.list.length){E.find.at=Math.max(0,E.find.at-1);gotoFind(1);}
}
function replaceRange(r,rep){const sc=r.startContainer,so=r.startOffset;r.deleteContents();if(sc.nodeType===3)sc.insertData(so,rep);else r.insertNode(document.createTextNode(rep));}
function replaceAll(){
  runFind();const F=E.find;if(!F.list.length){hint('לא נמצא');return;}const rep=q('weRepIn').value;const n=F.list.length;
  mutate(()=>{for(const r of [...F.list].reverse())replaceRange(r,rep);});runFind();E.opts.toast(`הוחלפו ${n} מופעים`);
}
function openFind(rep){const f=q('weFind');f.hidden=false;f.classList.toggle('rep-on',!!rep);const r=curRange();
  if(r&&!r.collapsed&&r.toString().length<80)q('weFindIn').value=r.toString();
  const i=q('weFindIn');i.focus();i.select();runFind();}
function closeFind(){q('weFind').hidden=true;E.find=null;paintFind();if(window.CSS&&window.CSS.highlights){window.CSS.highlights.delete('we-find');window.CSS.highlights.delete('we-find-cur');window.CSS.highlights.delete('we-cmt');window.CSS.highlights.delete('we-cmt-on');}
  if(E.lastRange){E.body.focus({preventScroll:true});const s=getSelection();s.removeAllRanges();s.addRange(E.lastRange);}}

/* ---------- pop-up menus ---------- */
function pop(anchor,html,onPick){
  const p=q('wePop');p.innerHTML=html;p.hidden=false;E.popFor=anchor;
  const a=anchor.getBoundingClientRect();const pw=p.offsetWidth,ph=p.offsetHeight;
  let x=a.right-pw;if(x<6)x=6;if(x+pw>innerWidth-6)x=innerWidth-6-pw;let y=a.bottom+4;if(y+ph>innerHeight-6)y=Math.max(6,a.top-ph-4);
  p.style.left=x+'px';p.style.top=y+'px';
  p.onmousedown=e=>{if(!e.target.closest('input'))e.preventDefault();};
  p.onclick=e=>{const b=e.target.closest('[data-v]');if(!b)return;const keep=onPick(b.dataset.v,b);if(!keep)closePop();};
  return p;
}
function closePop(){const p=q('wePop');if(p.hidden)return;p.hidden=true;p.innerHTML='';E.popFor=null;}
const pal=(list,none)=>`<div class="pal">${list.map(c=>`<button type="button" data-v="${c}" title="#${c}" style="background:#${c}"></button>`).join('')}</div>${none?`<button type="button" class="mi" data-v="none">${none}</button>`:''}`;
function openPop(kind,btn){
  if(E.popFor===btn){closePop();return;}
  if(kind==='pagenum'){const opts=[['header','right','בראש העמוד, בצד ימין'],['header','center','בראש העמוד, במרכז'],['header','left','בראש העמוד, בצד שמאל'],['footer','right','בתחתית העמוד, בצד ימין'],['footer','center','בתחתית העמוד, במרכז'],['footer','left','בתחתית העמוד, בצד שמאל']];
    pop(btn,`<h4>מספרי עמודים</h4>${opts.map(o=>`<button type="button" class="mi" data-v="${o[0]}:${o[1]}">${o[2]}</button>`).join('')}<div class="sep"></div><button type="button" class="mi" data-v="footer:center:xofy">"עמוד 1 מתוך 5" בתחתית העמוד</button><button type="button" class="mi" data-v="remove">הסרת מספרי העמודים</button>`,v=>{
      if(v==='remove')return removePageNumbers();const[w,a,f]=v.split(':');pageNumber(w,a,f);});return;}
  if(kind==='toc'){const has=!!tocGroup();pop(btn,`<button type="button" class="mi" data-v="ins">${has?'עדכון תוכן העניינים':'הוספת תוכן עניינים'}</button>${has?'<button type="button" class="mi" data-v="del">הסרת תוכן העניינים</button>':''}<div class="sep"></div><div style="font-size:11.5px;color:#7A879E;padding:2px 10px 4px;max-width:240px;white-space:normal">נבנה מהפסקאות בסגנון "כותרת 1", "כותרת 2" ו"כותרת 3". Ctrl+לחיצה על שורה עוברת לכותרת.</div>`,v=>{if(v==='ins')tocInsert();else tocRemove();});return;}
  if(kind==='color'){const p=pop(btn,`<h4>צבע גופן</h4>${pal(COLORS,'אוטומטי (שחור)')}<div class="sep"></div><label class="mi">צבע אחר… <input type="color" value="#2E58CC" style="margin-inline-start:auto"></label>`,v=>setColor(v==='none'?null:v));
    p.querySelector('input[type=color]').onchange=e=>{closePop();setColor(e.target.value.slice(1).toUpperCase());};}
  else if(kind==='hilite')pop(btn,`<h4>צבע הדגשה</h4>${pal(HLS,'ללא צבע')}`,v=>setHilite(v==='none'?null:v));
  else if(kind==='shade')pop(btn,`<h4>הצללת התא</h4>${pal(['D9E2F3','DEEAF6','E2EFD9','FFF2CC','FBE4D5','EDEDED','B4C6E7','C5E0B3','FFE599','F7CAAC','BFBFBF','FFFFFF','8EAADB','A8D08D','FFD966','F4B083','A5A5A5','2E58CC','FFD43B','16213A'],'ללא הצללה')}`,v=>shade(v==='none'?null:v));
  else if(kind==='spacing'){
    const b=curBlock();const cs=b&&getComputedStyle(b);const cur=b?(b.style.lineHeight||''):'';
    const mb=cs?parseFloat(cs.marginBottom):0,mt=cs?parseFloat(cs.marginTop):0;
    pop(btn,`<h4>ריווח בין שורות</h4>${['1','1.15','1.5','2','2.5','3'].map(v=>`<button type="button" class="mi ${cur===v?'on':''}" data-v="lh:${v}">${v==='1'?'1.0':v}</button>`).join('')}<div class="sep"></div>
      <button type="button" class="mi" data-v="${mt>0?'sb:0':'sb:12'}">${mt>0?'הסרת הרווח לפני הפסקה':'הוספת רווח לפני הפסקה'}</button>
      <button type="button" class="mi" data-v="${mb>0?'sa:0':'sa:8'}">${mb>0?'הסרת הרווח אחרי הפסקה':'הוספת רווח אחרי הפסקה'}</button>`,v=>{
      const[k,x]=v.split(':');if(k==='lh')lineSpacing(+x);else paraSpace(k==='sb'?'before':'after',+x);});}
  else if(kind==='table'){
    const p=pop(btn,`<h4>הוספת טבלה</h4><div class="tgrid">${Array.from({length:80},(_,i)=>`<i data-v="${Math.floor(i/10)+1}x${i%10+1}"></i>`).join('')}</div><div class="tlab" id="weTLab">בחרו גודל</div>`,v=>{const[r,c]=v.split('x').map(Number);insertTable(r,c);});
    p.querySelector('.tgrid').onmouseover=e=>{const v=e.target.dataset.v;if(!v)return;const[r,c]=v.split('x').map(Number);
      p.querySelectorAll('.tgrid i').forEach(i=>{const[a,b]=i.dataset.v.split('x').map(Number);i.classList.toggle('on',a<=r&&b<=c);});p.querySelector('#weTLab').textContent=`${c} עמודות × ${r} שורות`;};}
  else if(kind==='symbol')pop(btn,`<h4>סמלים</h4><div class="syms">${SYMBOLS.map(s=>`<button type="button" data-v="${s}">${s}</button>`).join('')}</div>`,v=>insertText(v));
  else if(kind==='margins'){
    const M=[['normal','רגילים','2.54 ס״מ מכל צד',[1440,1440,1440,1440]],['narrow','צרים','1.27 ס״מ מכל צד',[720,720,720,720]],['moderate','בינוניים','2.54 למעלה ולמטה, 1.91 בצדדים',[1440,1080,1440,1080]],['wide','רחבים','2.54 למעלה ולמטה, 5.08 בצדדים',[1440,2880,1440,2880]]];
    pop(btn,M.map(m=>`<button type="button" class="mi" data-v="${m[0]}"><b>${m[1]}</b><small>${m[2]}</small></button>`).join('')+'<div class="sep"></div><button type="button" class="mi" data-v="custom">שוליים מותאמים…</button>',v=>{
      if(v==='custom'){marginsDialog();return;}const m=M.find(x=>x[0]===v)[3];setPage({top:m[0],right:m[1],bottom:m[2],left:m[3]});});}
  else if(kind==='orient')pop(btn,`<button type="button" class="mi ${!E.page.landscape?'on':''}" data-v="p">לאורך</button><button type="button" class="mi ${E.page.landscape?'on':''}" data-v="l">לרוחב</button>`,v=>{
    const l=v==='l';if(l===E.page.landscape)return;setPage({landscape:l,w:E.page.h,h:E.page.w});});
  else if(kind==='paper'){const S=[['A4','21 × 29.7 ס״מ',11906,16838],['Letter','21.59 × 27.94 ס״מ',12240,15840],['Legal','21.59 × 35.56 ס״מ',12240,20160],['A5','14.8 × 21 ס״מ',8391,11906],['A3','29.7 × 42 ס״מ',16838,23811]];
    pop(btn,S.map(s=>`<button type="button" class="mi" data-v="${s[0]}"><b>${s[0]}</b><small>${s[1]}</small></button>`).join(''),v=>{const s=S.find(x=>x[0]===v);
      const w=E.page.landscape?s[3]:s[2],h=E.page.landscape?s[2]:s[3];setPage({w,h});});}
}
function marginsDialog(){
  const cm=t=>(t/567).toFixed(2);const pg=E.page;
  dialog(`<h3>שוליים מותאמים (ס״מ)</h3>
    <label>למעלה</label><input id="weM1" type="number" step="0.1" min="0" dir="ltr" value="${cm(pg.top)}">
    <label>למטה</label><input id="weM2" type="number" step="0.1" min="0" dir="ltr" value="${cm(pg.bottom)}">
    <label>ימין</label><input id="weM3" type="number" step="0.1" min="0" dir="ltr" value="${cm(pg.right)}">
    <label>שמאל</label><input id="weM4" type="number" step="0.1" min="0" dir="ltr" value="${cm(pg.left)}">
    <div class="acts"><button type="button" data-a="x">ביטול</button><button type="button" class="pri" data-a="ok">אישור</button></div>`,(a,root)=>{
      if(a!=='ok')return true;const v=id=>Math.max(0,Math.round((+root.querySelector(id).value||0)*567));
      setPage({top:v('#weM1'),bottom:v('#weM2'),right:v('#weM3'),left:v('#weM4')});return true;});
}
function setPage(ch){if(E.snapT)snap();Object.assign(E.page,ch);E.pageDirty=true;applyPage();after();}
function applyPage(){
  const pg=E.page,sh=q('weSheet');
  Object.assign(sh.style,{width:px(pg.w)+'px',minHeight:px(pg.h)+'px',paddingTop:px(pg.top)+'px',paddingBottom:px(pg.bottom)+'px',paddingLeft:px(pg.left)+'px',paddingRight:px(pg.right)+'px'});
  Object.assign(q('weHdr').style,{top:px(pg.header)+'px',left:px(pg.left)+'px',right:px(pg.right)+'px',maxHeight:Math.max(0,px(pg.top)-px(pg.header))+'px'});
  Object.assign(q('weFtr').style,{bottom:px(pg.footer)+'px',left:px(pg.left)+'px',right:px(pg.right)+'px',maxHeight:Math.max(0,px(pg.bottom)-px(pg.footer))+'px'});
}

/* ---------- where each page ends ---------- */
function guides(){
  const g=q('weGuides'),b=E.main,pg=E.page;g.innerHTML='';
  const H=px(pg.h)-px(pg.top)-px(pg.bottom);if(H<50)return;
  let start=0,pages=1;const marks=[];
  for(const el of b.children){
    const t=el.offsetTop,h=el.offsetHeight;
    let guard=0;
    while(t+h>start+H+.5&&guard++<500){
      const lh=parseFloat(getComputedStyle(el).lineHeight)||20;
      const whole=t>start+1&&h<=H&&(el.tagName!=='P'||h<lh*3.2);
      start=whole?t:start+H;marks.push(start);pages++;
    }
    const pb=el.querySelector('.we-pb')||(el.classList.contains('we-pb')?el:null);
    if(pb){const y=offIn(pb,b).y+pb.offsetHeight;start=y;marks.push(y);pages++;}
  }
  E.pageMarks=marks;E.pageCount=pages;
  for(const f of E.root.querySelectorAll('.we-fld[data-f=NUMPAGES]'))if(f.textContent!==String(pages))f.textContent=pages;
  const top=b.offsetTop;
  marks.forEach((y,i)=>{const m=document.createElement('i');m.style.top=(top+y)+'px';const l=document.createElement('b');l.textContent='עמוד '+(i+2);m.append(l);g.append(m);});
  pageInfo();
}
function pageInfo(){
  if(!E)return;const r=curRange();let cur=1;
  if(r&&E.pageMarks&&!E.hf){const n=r.startContainer.nodeType===3?r.startContainer.parentElement:r.startContainer;if(n&&n!==E.main){const y=offIn(n,E.main).y;cur=1+E.pageMarks.filter(m=>m<=y+1).length;}}
  q('wePages').textContent=`עמוד ${cur} מתוך ${E.pageCount||1}`;
}
function words(){E.wcTotal=countText(E.main.textContent).w;wordsSel();}
function wordsSel(){if(!E)return;const r=curRange();const n=E.wcTotal||0;
  q('weWords').textContent=r&&!r.collapsed&&!E.hf?`${countText(r.toString()).w.toLocaleString('he-IL')} מתוך ${n.toLocaleString('he-IL')} מילים`:`${n.toLocaleString('he-IL')} מילים`;}
function setZoom(z){z=Math.max(.3,Math.min(3,Math.round(z*100)/100));E.zoom=z;q('weSheet').style.zoom=z;const t=Math.round(z*100)+'%';q('weZoomV').textContent=t;q('weZoomS').textContent=t;q('weZoom').value=Math.round(z*100);posImgBox();}
function fitZoom(){const sc=q('weScroll');const w=px(E.page.w)+(innerWidth<760?16:48);setZoom(Math.min(1.6,(sc.clientWidth)/w));}

/* ---------- ribbon state ---------- */
let uiT=0;
function updateUI(){cancelAnimationFrame(uiT);uiT=requestAnimationFrame(updateNow);}
function updateNow(){
  if(!E)return;const r=curRange();const rb=q('weRibbon');
  const press=(cmd,on)=>rb.querySelectorAll(`[data-cmd="${cmd}"]`).forEach(b=>b.setAttribute('aria-pressed',on?'true':'false'));
  rb.querySelector('[data-cmd=undo]').disabled=E.hi<=0&&!E.snapT;rb.querySelector('[data-cmd=redo]').disabled=E.hi>=E.hist.length-1;
  const tabT=E.root.querySelector('.we-tabs [data-tab=table]');
  if(!r)return;
  let el=r.startContainer.nodeType===3?r.startContainer.parentElement:r.startContainer;if(!el||el===E.body){el=E.body.querySelector('p,h1,h2,h3,li')||E.body;}
  const blk=blockOf(el);const cs=getComputedStyle(el);
  const ff=firstFamily(cs.fontFamily)||E.st.def.font;const fsel=rb.querySelector('[data-in=font]');
  if(fsel.value!==ff){if(![...fsel.options].some(o=>o.value===ff)){const o=new Option(ff,ff);fsel.querySelector('optgroup')?fsel.querySelector('optgroup').prepend(o):fsel.prepend(o);}fsel.value=ff;}
  const sz=rb.querySelector('[data-in=size]');if(document.activeElement!==sz){const pt=el.dataset&&el.dataset.fsb?+el.dataset.fsb:Math.round(parseFloat(cs.fontSize)*.75*2)/2;sz.value=pt;}
  const p=blk?runProps(el,blk):null;
  press('bold',p&&p.b);press('italic',p&&p.i);press('underline',p&&p.u);press('strike',p&&p.s);press('sup',p&&p.va==='super');press('sub',p&&p.va==='sub');
  if(blk){const bc=getComputedStyle(blk);const va=visAlign(bc);['right','center','left','justify'].forEach(a=>press(a,va===a));
    press('rtl',bc.direction==='rtl');press('ltr',bc.direction!=='rtl');
    const n=numOf(blk);press('bullets',n&&n.id===E.st.num.bullet||(blk.tagName==='LI'&&blk.parentElement.tagName==='UL'));press('numbers',n&&n.id===E.st.num.decimal||(blk.tagName==='LI'&&blk.parentElement.tagName==='OL'));
    const sc=styleCls(blk);q('weGal').querySelectorAll('button').forEach(b=>{const it=galItems().find(x=>(x.id||'')===b.dataset.style);b.classList.toggle('on',!!it&&(it.normal?!sc:it.cls===sc));});
    const L=rb.querySelector('[data-panel=layout]');const set=(k,v)=>{const i=L.querySelector(`[data-in=${k}]`);if(document.activeElement!==i)i.value=v;};
    set('indR',(parseFloat(bc.marginRight)/37.7953).toFixed(2).replace(/\.?0+$/,''));set('indL',(parseFloat(bc.marginLeft)/37.7953).toFixed(2).replace(/\.?0+$/,''));
    set('spB',Math.round(parseFloat(bc.marginTop)*.75));set('spA',Math.round(parseFloat(bc.marginBottom)*.75));
  }
  const inCell=!!(el.closest&&el.closest('td,th')&&E.body.contains(el.closest('td,th')));tabT.hidden=!inCell;
  if(!inCell&&tabT.getAttribute('aria-selected')==='true')showTab('home');
  pageInfo();wordsSel();
}
function showTab(t){
  E.root.querySelectorAll('.we-tabs [data-tab]').forEach(b=>b.setAttribute('aria-selected',b.dataset.tab===t?'true':'false'));
  E.root.querySelectorAll('.wpanel').forEach(p=>p.hidden=p.dataset.panel!==t);
}
function fillFonts(){
  const s=q('weRibbon').querySelector('[data-in=font]');s.innerHTML='';
  const used=new Set();const docF=[];
  const ft=txt(E.st.files,'word/fontTable.xml');if(ft)for(const m of ft.matchAll(/<w:font\b[^>]*w:name="([^"]+)"/g))docF.push(m[1]);
  docF.unshift(E.st.def.font);
  const grp=(label,list)=>{const g=document.createElement('optgroup');g.label=label;for(const f of list){if(used.has(f)||/^(Symbol|Wingdings|Times New Roman \(.*|MS Mincho|SimSun|Mangal)$/.test(f))continue;used.add(f);const o=new Option(f,f);o.style.fontFamily=cssFont(f);g.append(o);}if(g.children.length)s.append(g);};
  grp('גופני המסמך',docF.slice(0,10));grp('גופנים בעברית',FONTS_HE);grp('גופנים באנגלית',FONTS_LAT);
}

/* ---------- commands ---------- */
const CMD={
  undo,redo,
  cut:()=>{E.body.focus();document.execCommand('cut');},copy:()=>{E.body.focus();document.execCommand('copy');},
  paste:async()=>{
    if(!navigator.clipboard||!navigator.clipboard.read){hint('הדבקה: Ctrl+V');return;}
    try{const items=await navigator.clipboard.read();liveRange();
      for(const it of items){
        const img=it.types.find(t=>t.startsWith('image/'));if(img){await insertImageFile(new File([await it.getType(img)],'image',{type:img}));return;}
        if(it.types.includes('text/html')){pasteHtml(await (await it.getType('text/html')).text());return;}
        if(it.types.includes('text/plain')){insertText(await (await it.getType('text/plain')).text());return;}
      }}catch(_){hint('הדבקה: Ctrl+V');}
  },
  painter:()=>{if(E.painter){E.painter=null;updatePainter();return;}const r=liveRange();if(!r)return;const el=r.startContainer.nodeType===3?r.startContainer.parentElement:r.startContainer;const b=blockOf(el);if(!b)return;E.painter=runProps(el,b);updatePainter();hint('סמנו את הטקסט שיקבל את העיצוב');},
  bold:()=>fmt('bold'),italic:()=>fmt('italic'),underline:()=>fmt('underline'),strike:()=>fmt('strike'),sub:()=>fmt('sub'),sup:()=>fmt('sup'),
  grow:()=>growFont(1),shrink:()=>growFont(-1),clear:clearFmt,
  color:()=>setColor(E.color===undefined?'C00000':E.color),hilite:()=>setHilite(E.hl===undefined?'FFFF00':E.hl),
  bullets:()=>toggleList('b'),numbers:()=>toggleList('n'),indent:()=>indent(1),outdent:()=>indent(-1),
  rtl:()=>setDir('rtl'),ltr:()=>setDir('ltr'),right:()=>align('right'),center:()=>align('center'),left:()=>align('left'),justify:()=>align('justify'),
  marks:()=>{E.root.classList.toggle('we-marks');const on=E.root.classList.contains('we-marks');E.root.querySelectorAll('[data-cmd=marks]').forEach(b=>b.setAttribute('aria-pressed',on));},
  guides:()=>{E.root.classList.toggle('we-noguides');E.root.querySelectorAll('[data-cmd=guides]').forEach(b=>b.setAttribute('aria-pressed',!E.root.classList.contains('we-noguides')));},
  spell:()=>{const on=E.main.spellcheck=!E.main.spellcheck;E.body.spellcheck=on;E.root.querySelectorAll('[data-cmd=spell]').forEach(b=>b.setAttribute('aria-pressed',on));E.body.blur();E.body.focus();},
  find:()=>openFind(false),replace:()=>openFind(true),
  image:()=>q('weImgFile').click(),link:linkDialog,pagebreak:insertPageBreak,hr:insertHr,
  date:()=>insertText(new Date().toLocaleDateString('he-IL',{day:'numeric',month:'numeric',year:'numeric'})),
  rowAbove:()=>tableOp('rowAbove'),rowBelow:()=>tableOp('rowBelow'),colRight:()=>tableOp('colRight'),colLeft:()=>tableOp('colLeft'),
  delRow:()=>tableOp('delRow'),delCol:()=>tableOp('delCol'),delTable:()=>tableOp('delTable'),borders:()=>tableOp('borders'),noborders:()=>tableOp('noborders'),
  shade:()=>shade(E.shade===undefined?'D9E2F3':E.shade),
  zin:()=>setZoom(E.zoom+.1),zout:()=>setZoom(E.zoom-.1),z100:()=>setZoom(1),zfit:fitZoom,
  print:printDoc,
  hdr:()=>enterHF('header'),ftr:()=>enterHF('footer'),hfClose:()=>exitHF(),hfGo:()=>enterHF(E.hf&&E.hf.kind==='header'?'footer':'header'),
  pnXofY:()=>{if(!E.hf)enterHF('footer');hfField('PAGE');const r=liveRange();if(!r)return;mutate(()=>{const rr=liveRange();const t=document.createTextNode(' מתוך ');rr.insertNode(t);rr.setStartAfter(t);rr.collapse(true);
    const tp=document.createElement('template');tp.innerHTML=fld('NUMPAGES');const n=tp.content.firstChild;rr.insertNode(n);caretTo(n.parentNode,[...n.parentNode.childNodes].indexOf(n)+1);
    const f=E.body.querySelectorAll('.we-fld[data-f=PAGE]');const last=f[f.length-1];if(last&&!/עמוד\s*$/.test((last.previousSibling&&last.previousSibling.textContent)||''))last.before('עמוד ');});},
  cmNew,cmDel:()=>{const id=cmCurrent();if(id==null||!E.cm.get(id))return hint('אין הערה במקום הזה');cmRemove(id);},cmPrev:()=>cmGo(-1),cmNext:()=>cmGo(1),
  cmShow:()=>{E.cmHide=!E.cmHide;E.root.querySelectorAll('[data-cmd=cmShow]').forEach(b=>b.setAttribute('aria-pressed',!E.cmHide));cmRender();},
  wcount:wordCount,
  acToggle:()=>{acOn=!acOn;try{localStorage.setItem('siman-autocorrect',acOn?'1':'0');}catch(_){}E.root.querySelectorAll('[data-cmd=acToggle]').forEach(b=>b.setAttribute('aria-pressed',acOn));hint(acOn?'תיקון אוטומטי פעיל':'תיקון אוטומטי כבוי');}
};
function updatePainter(){E.root.querySelectorAll('[data-cmd=painter]').forEach(b=>b.setAttribute('aria-pressed',!!E.painter));}

// printing happens in a separate frame that holds only the document, so nothing else on the page can hide it
async function printFrame(html,title){
  const old=document.getElementById('smPrintFrame');if(old)old.remove();
  const fr=document.createElement('iframe');fr.id='smPrintFrame';fr.setAttribute('aria-hidden','true');fr.title=title||'';
  fr.style.cssText='position:fixed;right:0;bottom:0;width:1px;height:1px;border:0;opacity:0;pointer-events:none';
  document.body.append(fr);const d=fr.contentDocument;d.open();d.write(html);d.close();
  try{await Promise.race([d.fonts&&d.fonts.ready,new Promise(r=>setTimeout(r,2500))]);}catch(_){}
  await Promise.all([...d.images].map(i=>i.complete?0:new Promise(r=>{i.onload=i.onerror=r;setTimeout(r,3000);})));
  await new Promise(r=>setTimeout(r,60));
  const w=fr.contentWindow;w.addEventListener('afterprint',()=>setTimeout(()=>fr.remove(),300));
  if(window.__smPrintHook)window.__smPrintHook(w);else{w.focus();w.print();}
}
async function printDoc(){
  if(E.hf)exitHF(true);
  const pg=E.page;hideImgBox();const sel=getSelection();if(sel)sel.removeAllRanges();E.body.blur();
  const cm=t=>(t/567).toFixed(2)+'cm';
  const fonts=CSS.slice(0,CSS.indexOf('#wed{'));
  const styles=[...q('weStyles').querySelectorAll('style')].map(x=>x.textContent).join('\n');
  const notes=q('weNotes').innerHTML;
  const html=`<!doctype html><html lang="he" dir="rtl"><head><meta charset="utf-8"><base href="${esc(location.href)}"><title>${esc(fileName())}</title><style>${fonts}
${styles}
@page{margin:${cm(pg.top)} ${cm(pg.right)} ${cm(pg.bottom)} ${cm(pg.left)};${pg.w>pg.h?'size:landscape;':''}}
html,body{margin:0;padding:0;background:#fff}*{-webkit-print-color-adjust:exact;print-color-adjust:exact}
.doc{color:#000;overflow-wrap:break-word;line-height:normal;text-align:start}
.doc img{max-width:100%}.doc table{border-collapse:collapse}
.doc .we-pb{display:block;height:0;border:0;margin:0;break-after:page}
.doc .we-tab{display:inline-block;width:36pt;white-space:pre}
.doc hr{border:0;border-top:1px solid #000;margin:6pt 0}
.notes{border-top:1px solid #999;margin-top:18px;padding-top:6px;font-size:10pt}
</style></head><body><div class="doc ${P}" style="${esc(E.body.getAttribute('style')||'')}">${E.body.innerHTML}</div>${notes?`<div class="notes ${P}">${notes}</div>`:''}</body></html>`;
  await printFrame(html,fileName());
}
function sanitize(html){
  const d=new DOMParser().parseFromString(html,'text/html');
  d.querySelectorAll('script,style,meta,link,title,xml,iframe,object,embed,form,input,button,select,textarea,svg,canvas,video,audio,noscript').forEach(n=>n.remove());
  const tw=d.createTreeWalker(d.body,NodeFilter.SHOW_COMMENT);const cm=[];while(tw.nextNode())cm.push(tw.currentNode);cm.forEach(n=>n.remove());
  for(const el of [...d.body.querySelectorAll('*')]){
    if(/^O:P$/i.test(el.tagName)||el.tagName.includes(':')){el.replaceWith(...el.childNodes);continue;}
    for(const a of [...el.attributes]){const n=a.name.toLowerCase();
      if(!['style','href','src','colspan','rowspan','dir','width','height'].includes(n))el.removeAttribute(a.name);
      else if(n==='href'&&/^\s*javascript:/i.test(a.value))el.removeAttribute(a.name);}
    if(el.style){for(const pr of ['position','left','top','right','bottom','z-index','float','display','visibility','opacity','transform','content'])el.style.removeProperty(pr);
      [...el.style].filter(x=>x.startsWith('mso-')).forEach(x=>el.style.removeProperty(x));}
    if(el.tagName==='IMG'&&!/^(data:image\/|blob:|https?:)/i.test(el.getAttribute('src')||''))el.remove();
  }
  return d.body.innerHTML;
}
function pasteHtml(html){const clean=sanitize(html);if(!clean.trim())return;liveRange();document.execCommand('insertHTML',false,clean);}

/* ---------- events ---------- */
function wire(){
  const R=E.root,rb=q('weRibbon');
  R.querySelector('.we-tabs').onclick=e=>{const b=e.target.closest('[data-tab]');if(b)showTab(b.dataset.tab);};
  R.querySelector('.we-tabs').onmousedown=e=>e.preventDefault();
  R.addEventListener('mousedown',e=>{const b=e.target.closest('button.wb,.we-status .wb');if(b)e.preventDefault();});
  R.addEventListener('click',e=>{
    const b=e.target.closest('[data-cmd]');if(!b||b.disabled)return;
    if(b.dataset.pop){openPop(b.dataset.pop,b);return;}
    closePop();const f=CMD[b.dataset.cmd];if(f){if(!/^(z|print|guides|marks|spell|find|replace|hf|hdr|ftr|cm|wcount|ac|toc)/.test(b.dataset.cmd))liveRange();f();}
  });
  const fsel=rb.querySelector('[data-in=font]');
  fsel.onchange=()=>{const v=fsel.value;liveRange();setFont(v);E.body.focus();};
  const sz=rb.querySelector('[data-in=size]');
  const applySz=()=>{const v=parseFloat(String(sz.value).replace(',','.'));if(v>0){liveRange();setSize(v);}};
  sz.onchange=applySz;sz.onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();applySz();E.body.focus();}};
  const L=rb.querySelector('[data-panel=layout]');
  L.querySelectorAll('input[data-in]').forEach(i=>i.onchange=()=>{const v=parseFloat(i.value)||0;liveRange();
    if(i.dataset.in==='indR')indentPhys('r',v);else if(i.dataset.in==='indL')indentPhys('l',v);else paraSpace(i.dataset.in==='spB'?'before':'after',v);});
  q('weZoom').oninput=e=>setZoom(+e.target.value/100);
  q('weImgFile').onchange=async e=>{const f=e.target.files[0];e.target.value='';liveRange();await insertImageFile(f);};
  q('weImgBox').addEventListener('pointerdown',startResize);
  // find
  q('weFindIn').oninput=()=>runFind();q('weFindIn').onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();gotoFind(e.shiftKey?-1:1);}else if(e.key==='Escape')closeFind();};
  q('weRepIn').onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();replaceOne();}else if(e.key==='Escape')closeFind();};
  q('weFindNext').onclick=()=>gotoFind(1);q('weFindPrev').onclick=()=>gotoFind(-1);q('weFindX').onclick=closeFind;q('weRep1').onclick=replaceOne;q('weRepAll').onclick=replaceAll;
  // top bar
  q('weSave').onclick=()=>doSave();q('wePdf').onclick=()=>doPdf();q('weDesk').onclick=()=>doDesk();q('weClose').onclick=()=>doClose();
  bindEditable(E.body);
  q('weScroll').addEventListener('wheel',e=>{if(e.ctrlKey||e.metaKey){e.preventDefault();setZoom(E.zoom+(e.deltaY<0?.1:-.1));}},{passive:false});
  // headers and footers, comments, table of contents
  for(const k of ['header','footer'])hfBox(k).addEventListener('dblclick',e=>{e.preventDefault();enterHF(k);});
  E.main.addEventListener('dblclick',e=>{if(E.hf){e.preventDefault();exitHF();}});
  E.main.addEventListener('click',e=>{const t=e.target.closest('.we-toc');if(t&&(e.ctrlKey||e.metaKey)){e.preventDefault();tocGo(t);}});
  const cp=q('weCPane');
  cp.addEventListener('click',e=>{const b=e.target.closest('[data-ca]');if(b){cmAction(b.dataset.ca,+b.dataset.id);return;}
    if(e.target.closest('textarea'))return;const c=e.target.closest('.cmc');if(c)cmSelect(+c.dataset.id,false);});
  cp.addEventListener('dblclick',e=>{const x=e.target.closest('.cmx');if(!x)return;const r=x.closest('.cmr')||x.closest('.cmc');const c=E.cm.get(+r.dataset.id);if(c&&c.author===cmAuthor())cmAction('edit',c.id);});
  cp.addEventListener('keydown',e=>{const ta=e.target.closest('textarea');if(!ta)return;
    if(e.key==='Enter'&&(e.ctrlKey||e.metaKey)){e.preventDefault();cmAction('post',+ta.dataset.id);}else if(e.key==='Escape'){e.preventDefault();cmAction('cancel',+ta.dataset.id);}});
  cp.addEventListener('input',e=>{const ta=e.target.closest('textarea');if(ta){const c=E.cm.get(+ta.dataset.id);if(c)c.text=ta.value;}});
  q('weWords').onclick=()=>wordCount();
  E.onSel=()=>{if(!E)return;const r=curRange();if(r){E.lastRange=r.cloneRange();E.lastSel=null;}updateUI();
    if(r&&!E.hf&&E.cm.size){const id=cmCurrent();if(id!==E.cmOn&&!q('weCPane').querySelector('textarea')){E.cmOn=id;q('weCPane').querySelectorAll('.cmc').forEach(c=>c.classList.toggle('on',+c.dataset.id===id));cmPaint();}}};
  document.addEventListener('selectionchange',E.onSel);
  E.onKey=onKey;window.addEventListener('keydown',E.onKey,true);
  E.onPaste=e=>{if(!E)return;if(!E.root.contains(e.target)){e.stopImmediatePropagation();return;}
    if(!E.body.contains(e.target))return;e.stopImmediatePropagation();e.preventDefault();const dt=e.clipboardData;if(!dt)return;
    const f=[...dt.files].find(x=>/^image\//.test(x.type));if(f){insertImageFile(f);return;}
    const h=dt.getData('text/html');if(h){pasteHtml(h);return;}
    const t=dt.getData('text/plain');if(t)insertText(t);};
  window.addEventListener('paste',E.onPaste,true);
  E.onCopy=e=>{if(E&&!E.root.contains(e.target))e.stopImmediatePropagation();};window.addEventListener('copy',E.onCopy,true);window.addEventListener('cut',E.onCopy,true);
  ['keydown','keyup','keypress','paste','copy','cut','input','change'].forEach(t=>R.addEventListener(t,e=>e.stopPropagation()));
  E.onDrag=e=>{if(!E)return;e.stopImmediatePropagation();if(e.dataTransfer&&[...e.dataTransfer.types].includes('Files'))e.preventDefault();};
  E.onDrop=e=>{if(!E)return;e.stopImmediatePropagation();const dt=e.dataTransfer;if(!dt||![...dt.types].includes('Files'))return;e.preventDefault();
    const imgs=[...dt.files].filter(f=>/^image\//.test(f.type));if(!imgs.length){E.opts.toast('אפשר לגרור לכאן תמונות');return;}
    const cr=document.caretRangeFromPoint&&document.caretRangeFromPoint(e.clientX,e.clientY);if(cr&&body.contains(cr.startContainer)){const s=getSelection();s.removeAllRanges();s.addRange(cr);E.lastRange=cr.cloneRange();}
    (async()=>{for(const f of imgs)await insertImageFile(f);})();};
  window.addEventListener('dragover',E.onDrag,true);window.addEventListener('drop',E.onDrop,true);
  E.onUnload=e=>{if(E&&E.dirty){e.preventDefault();e.returnValue='';}};window.addEventListener('beforeunload',E.onUnload);
  E.onDown=e=>{if(!E)return;const p=q('wePop');if(!p.hidden&&!p.contains(e.target)&&!(E.popFor&&E.popFor.contains(e.target)))closePop();};
  document.addEventListener('pointerdown',E.onDown,true);
  E.onResize=()=>{schedule();};window.addEventListener('resize',E.onResize);
  q('weScroll').addEventListener('scroll',()=>{if(!q('wePop').hidden)closePop();},{passive:true});
}
function bindEditable(body){
  body.addEventListener('contextmenu',ctxMenu);
  body.addEventListener('input',e=>{
    E.lastAC=null;
    if(e.inputType==='insertText'&&(e.data===' '||e.data==='\u00a0')&&autoList()){snap();after();return;}
    if(e.inputType==='insertText'&&e.data&&e.data.length===1)autoCorrect(e.data);
    if(e.inputType==='insertParagraph'){const b=curBlock();const pv=b&&b.previousElementSibling;
      if(b&&pv&&isHeading(pv)&&!b.textContent.replace(ZW,'').trim()){for(const c of [...b.classList])if(c.startsWith(P+'_'))b.classList.remove(c);if(/^H\d$/.test(b.tagName)){const p=document.createElement('p');p.innerHTML='<br>';b.replaceWith(p);caretTo(p,0);}else{b.innerHTML='<br>';caretTo(b,0);}}}
    if(!body.firstElementChild||[...body.childNodes].some(n=>n.nodeType===3&&n.data.trim()))ensureBlocks();
    snapLater();markDirtySoon();refreshCounters();schedule();updateUI();
  });
  body.addEventListener('beforeinput',e=>{
    const t=e.inputType;
    if(t==='historyUndo'){e.preventDefault();if(!(E.lastAC&&undoAC()))undo();}else if(t==='historyRedo'){e.preventDefault();redo();}
    else if(t==='formatBold'||t==='formatItalic'||t==='formatUnderline'||t==='formatStrikeThrough'){e.preventDefault();fmt({formatBold:'bold',formatItalic:'italic',formatUnderline:'underline',formatStrikeThrough:'strike'}[t]);}
    else if(E.selImg&&(t==='deleteContentBackward'||t==='deleteContentForward')){e.preventDefault();deleteImg();}
    else if(E.snapT===0&&/^(insert|delete)/.test(t)&&E.hist[E.hi]&&E.hist[E.hi].html!==body.innerHTML)snap();
  });
  body.addEventListener('click',e=>{
    const img=e.target.closest('img');if(img&&body.contains(img)){selectImg(img);return;}
    hideImgBox();
    const a=e.target.closest('a[href]');if(a&&(e.ctrlKey||e.metaKey)){const h=a.getAttribute('href');if(/^(https?:|mailto:)/i.test(h))window.open(h,'_blank','noopener');}
  });
  body.addEventListener('mouseup',()=>{if(E.painter){const r=curRange();if(r&&!r.collapsed){const P0=E.painter;E.painter=null;updatePainter();
    inline(s=>{s.setAttribute('style',propsCss(P0));if(P0.va)s.dataset.fsb=P0.fs;else delete s.dataset.fsb;});}}});
  body.addEventListener('dragstart',e=>{if(e.target.tagName==='IMG')e.preventDefault();});
}
let mdT=0;function markDirtySoon(){clearTimeout(mdT);mdT=setTimeout(()=>{if(E)markDirty();},300);}
function onKey(e){
  if(!E||(e.target.closest&&e.target.closest('#shareDlg')))return;
  const t=e.target,inBody=E.body.contains(t)||t===E.body,inRoot=E.root.contains(t);
  if(!inRoot)e.stopImmediatePropagation();
  if(E.dlgOpen){if(!inRoot)e.preventDefault();return;}
  const mod=(e.ctrlKey||e.metaKey)&&!e.altKey,c=e.code;
  const run=(f,...a)=>{e.preventDefault();f(...a);};
  if(e.key==='Escape'){if(!q('wePop').hidden)return run(closePop);if(E.hf&&(inBody||t===document.body))return run(exitHF);if(!q('weFind').hidden)return run(closeFind);if(E.selImg)return run(hideImgBox);if(E.painter){E.painter=null;updatePainter();return;}return;}
  if((e.ctrlKey||e.metaKey)&&e.altKey&&c==='KeyM')return run(cmNew);
  if(mod){
    if(c==='KeyS')return run(doSave);
    if(c==='KeyP')return run(printDoc);
    if(c==='KeyF')return run(openFind,false);
    if(c==='KeyH')return run(openFind,true);
    if(!inBody&&!(t===document.body))return;
    if(c==='KeyZ'&&!e.shiftKey&&E.lastAC){e.preventDefault();if(!undoAC())undo();return;}
    if(c==='KeyZ')return run(e.shiftKey?redo:undo);
    if(c==='KeyY')return run(redo);
    if(c==='KeyB')return run(fmt,'bold');
    if(c==='KeyI')return run(fmt,'italic');
    if(c==='KeyU')return run(fmt,'underline');
    if(c==='KeyK')return run(linkDialog);
    if(c==='KeyE')return run(align,'center');
    if(c==='KeyR')return run(align,'right');
    if(c==='KeyL')return run(align,'left');
    if(c==='KeyJ')return run(align,'justify');
    if(c==='Space')return run(clearFmt);
    if(c==='Enter')return run(insertPageBreak);
    if(c==='Period'&&e.shiftKey)return run(growFont,1);
    if(c==='Comma'&&e.shiftKey)return run(growFont,-1);
    if(c==='BracketRight')return run(growFont,1,true);
    if(c==='BracketLeft')return run(growFont,-1,true);
    if(c==='Equal')return run(fmt,e.shiftKey?'sup':'sub');
    return;
  }
  if((e.ctrlKey||e.metaKey)&&e.altKey&&/^Digit[0-3]$/.test(c)&&inBody){const n=+c.slice(5);const it=galItems();
    const pick=n===0?it.find(x=>x.normal):it.find(x=>x.label==='כותרת '+n);if(pick)run(applyStyle,pick);return;}
  if(!inBody)return;
  if(E.selImg&&(e.key==='Delete'||e.key==='Backspace'))return run(deleteImg);
  if(e.key==='Tab'){
    if(cellOf()){e.preventDefault();nextCell(e.shiftKey);return;}
    const b=curBlock();const r=curRange();
    if(b&&numOf(b)&&r&&r.collapsed&&offsetOf(r.startContainer,r.startOffset)===offsetOf(b,0)+0){return run(indent,e.shiftKey?-1:1);}
    if(e.shiftKey)return run(indent,-1);
    e.preventDefault();mutate(()=>{const rr=liveRange();if(!rr)return;rr.deleteContents();const sp=document.createElement('span');sp.className='we-tab '+P+'-tab-stop';sp.contentEditable='false';sp.textContent='\t';rr.insertNode(sp);caretTo(sp.parentNode,[...sp.parentNode.childNodes].indexOf(sp)+1);});return;
  }
  if(e.key==='Enter'&&!e.shiftKey){const b=curBlock();if(b&&numOf(b)&&!b.textContent.replace(ZW,'').trim()){e.preventDefault();mutate(()=>endList(b));return;}}
  if(e.key==='Backspace'){const b=curBlock();const r=curRange();if(b&&numOf(b)&&r&&r.collapsed){const st=document.createRange();st.setStart(b,0);st.setEnd(r.startContainer,r.startOffset);
    if(!st.toString().replace(ZW,'')){e.preventDefault();mutate(()=>endList(b));return;}}}
}

/* ------------------------------------------------------------ writing the .docx */
function sigOf(el){
  const cs=getComputedStyle(el);const rtl=cs.direction==='rtl';const r1=v=>Math.round((parseFloat(v)||0)*10)/10;
  return{tag:el.tagName,cls:[...el.classList].filter(c=>c.startsWith(P+'_')||numRe.test(c)).sort().join(' '),dir:cs.direction,ta:visAlign(cs),
    ms:r1(rtl?cs.marginRight:cs.marginLeft),me:r1(rtl?cs.marginLeft:cs.marginRight),ti:r1(cs.textIndent),mt:r1(cs.marginTop),mb:r1(cs.marginBottom),lh:el.style.lineHeight||cs.lineHeight};
}
const PPR_ORDER=['pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr','suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens','kinsoku','wordWrap','overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi','adjustRightInd','snapToGrid','spacing','ind','contextualSpacing','mirrorIndents','suppressOverlap','jc','textDirection','textAlignment','textboxTightWrap','outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange'];
const SECT_ORDER=['headerReference','footerReference','footnotePr','endnotePr','type','pgSz','pgMar','paperSrc','pgBorders','lnNumType','pgNumType','cols','formProt','vAlign','noEndnote','titlePg','textDirection','bidi','rtlGutter','docGrid','printerSettings','sectPrChange'];
function setChild(parent,e,order){
  for(const x of kids(parent,e.localName))parent.removeChild(x);
  const i=order.indexOf(e.localName);const nx=kids(parent).find(c=>order.indexOf(c.localName)>i);
  parent.insertBefore(e,nx||null);return e;
}
const jcFor=(va,rtl)=>va==='center'?'center':va==='justify'?'both':va==='right'?(rtl?'left':'right'):va==='left'?(rtl?'right':'left'):null;
function styleIdOf(el){
  const c=styleCls(el);if(c){const s=E.st.styles.find(x=>classKey(x.id)===c);if(s)return s.id;}
  const h=/^H([1-6])$/.exec(el.tagName);if(h){const s=E.st.styles.find(x=>x.type==='paragraph'&&x.name==='heading '+h[1]);if(s)return s.id;}
  return null;
}
function numFor(el,ctx){
  const n=numOf(el);if(!n)return null;
  if(n.id!==E.st.num.decimal)return{id:n.id,lvl:n.lvl};
  const start=!!el.dataset.ns;
  if(start||ctx.curDec===null){ctx.curDec=ctx.decUsed?ctx.newDec():E.st.num.decimal;ctx.decUsed=true;}
  return{id:ctx.curDec,lvl:n.lvl};
}
function patchPPr(pPr,el,old,now,ctx,opts){
  const d=pPr.ownerDocument,rtl=now.dir==='rtl',fresh=!old;const ch=k=>fresh||old[k]!==now[k];
  const set=(name,attrs)=>setChild(pPr,wNew(d,name,attrs),PPR_ORDER);
  const del=name=>{for(const e of kids(pPr,name))pPr.removeChild(e);};
  if(opts.list||ch('cls')||ch('tag')){
    const sid=styleIdOf(el);if(sid)set('pStyle',{val:sid});else del('pStyle');
    const n=opts.list||numFor(el,ctx);
    if(n){const np=wNew(d,'numPr');np.append(wNew(d,'ilvl',{val:n.lvl}),wNew(d,'numId',{val:n.id}));setChild(pPr,np,PPR_ORDER);}else del('numPr');
  }
  if(fresh){if(rtl)set('bidi');else if(E.st.heb)set('bidi',{val:'0'});}
  else if(ch('dir'))rtl?set('bidi'):set('bidi',{val:'0'});
  if(fresh?!!el.style.textAlign:(ch('ta')||ch('dir'))){const j=jcFor(now.ta,rtl);if(j)set('jc',{val:j});else del('jc');}
  const st=el.style;
  if(fresh?(!numOf(el)&&!opts.list&&!!(st.marginLeft||st.marginRight||st.marginInlineStart||st.marginInlineEnd||st.textIndent)):(ch('ms')||ch('me')||ch('ti'))){
    const a={left:twip(now.ms),right:twip(now.me)};if(now.ti>0)a.firstLine=twip(now.ti);else if(now.ti<0)a.hanging=twip(-now.ti);set('ind',a);}
  if(fresh?!!(st.marginTop||st.marginBottom||st.lineHeight):(ch('mt')||ch('mb')||ch('lh'))){
    const old0=kid(pPr,'spacing');const sp=wNew(d,'spacing');
    if(old0)for(const at of [...old0.attributes])if(!/(before|after)Autospacing/.test(at.localName))sp.setAttributeNode(at.cloneNode());
    wSet(sp,'before',twip(now.mt));wSet(sp,'after',twip(now.mb));
    if(fresh?!!st.lineHeight:ch('lh')){const lh=st.lineHeight||'';
      if(/^[\d.]+$/.test(lh)){wSet(sp,'line',Math.round(+lh*240));wSet(sp,'lineRule','auto');}
      else if(/px|pt/.test(lh)){wSet(sp,'line',twip(/pt/.test(lh)?parseFloat(lh)/.75:parseFloat(lh)));wSet(sp,'lineRule','exact');}
      else if(/px/.test(now.lh)&&!/normal/.test(now.lh)&&lh){wSet(sp,'line',twip(parseFloat(now.lh)));wSet(sp,'lineRule','exact');}
      else{sp.removeAttributeNS(W,'line');sp.removeAttributeNS(W,'lineRule');}}
    setChild(pPr,sp,PPR_ORDER);
  }
  if(fresh&&!kid(pPr,'rPr')){const r=wNew(d,'rPr');const hp=Math.round(parseFloat(getComputedStyle(el).fontSize)*.75*2);
    r.append(wNew(d,'sz',{val:hp}),wNew(d,'szCs',{val:hp}));if(rtl)r.append(wNew(d,'rtl'));setChild(pPr,r,PPR_ORDER);}
}
const U_MAP={solid:'single',double:'double',dotted:'dotted',dashed:'dash',wavy:'wave'};
function rprXml(p,rtl,styled){
  let x='';if(p.ff){const f=esc(p.ff);x+=`<w:rFonts w:ascii="${f}" w:hAnsi="${f}" w:cs="${f}" w:eastAsia="${f}"/>`;}
  if(p.b)x+='<w:b/><w:bCs/>';else if(styled)x+='<w:b w:val="0"/><w:bCs w:val="0"/>';
  if(p.i)x+='<w:i/><w:iCs/>';else if(styled)x+='<w:i w:val="0"/><w:iCs w:val="0"/>';
  if(p.s)x+='<w:strike/>';
  if(p.color!=='000000'||styled)x+=`<w:color w:val="${p.color}"/>`;
  const hp=Math.max(2,Math.round(p.fs*2));x+=`<w:sz w:val="${hp}"/><w:szCs w:val="${hp}"/>`;
  if(p.bg&&HL[p.bg])x+=`<w:highlight w:val="${HL[p.bg]}"/>`;
  if(p.u)x+=`<w:u w:val="${U_MAP[p.us]||'single'}"/>`;
  if(p.bg&&!HL[p.bg])x+=`<w:shd w:val="clear" w:color="auto" w:fill="${p.bg}"/>`;
  if(p.va)x+=`<w:vertAlign w:val="${p.va==='super'?'superscript':'subscript'}"/>`;
  if(rtl)x+='<w:rtl/><w:lang w:bidi="he-IL"/>';
  return`<w:rPr>${x}</w:rPr>`;
}
const LATRUN=/[A-Za-z\u00C0-\u024F](?:[A-Za-z\u00C0-\u024F0-9 .,'’:\/@&+_#%=-]*[A-Za-z\u00C0-\u024F0-9])?/g;
function scriptSegs(t,rtlPara){
  if(!/[A-Za-z\u00C0-\u024F]/.test(t))return[[t,rtlPara||HEB.test(t)]];
  const out=[];let last=0,m;LATRUN.lastIndex=0;
  while((m=LATRUN.exec(t))){if(m.index>last){const s=t.slice(last,m.index);out.push([s,rtlPara||HEB.test(s)]);}out.push([m[0],false]);last=m.index+m[0].length;}
  if(last<t.length){const s=t.slice(last);out.push([s,rtlPara||HEB.test(s)]);}
  return out;
}
function textXml(t){
  return t.split(/(\t|\n)/).map(s=>s==='\t'?'<w:tab/>':s==='\n'?'<w:br/>':s?`<w:t xml:space="preserve">${xmlText(s)}</w:t>`:'').join('');
}
function trailingBr(br,block){
  for(let n=br;n&&n!==block;n=n.parentNode){
    for(let s=n.nextSibling;s;s=s.nextSibling){
      if(s.nodeType===3&&s.data.replace(ZW,''))return false;
      if(s.nodeType===1&&(isTok(s)||s.textContent.replace(ZW,'')||s.querySelector('img,br')))return false;
    }
  }
  return true;
}
function propsAt(el,block,ctx){let p=ctx.props.get(el);if(!p){p=runProps(el,block);ctx.props.set(el,p);}return p;}
function runsOf(nodes,block,ctx){
  const out=[];const bcs=getComputedStyle(block);const rtlPara=bcs.direction==='rtl';
  const styled=!!styleCls(block)||/^H\d$/.test(block.tagName);
  const push=(rpr,text)=>{const l=out[out.length-1];if(l&&l.rpr===rpr&&l.text!=null)l.text+=text;else out.push({rpr,text});};
  const walk=list=>{for(const c of list){
    if(c.nodeType===3){
      let t=c.data.replace(ZW,'');if(!t)continue;
      const par=c.parentElement||block;const pre=/^pre|break-spaces/.test(getComputedStyle(par).whiteSpace);
      if(!pre)t=t.replace(/[\t\n\r ]+/g,' ');
      const p=propsAt(par,block,ctx);
      for(const [s,r] of scriptSegs(t,rtlPara))push(rprXml(p,r,styled),s);
      continue;
    }
    if(c.nodeType!==1)continue;
    if(c.dataset.atom!==undefined){out.push({raw:atomXml(c,ctx)});continue;}
    if(c.classList.contains('we-pb')){out.push({raw:'<w:r><w:br w:type="page"/></w:r>'});continue;}
    if(c.classList.contains('we-tab')||c.classList.contains(P+'-tab-stop')){const p=propsAt(c.parentElement||block,block,ctx);out.push({raw:`<w:r>${rprXml(p,rtlPara,styled)}<w:tab/></w:r>`});continue;}
    if(c.tagName==='BR'){if(!trailingBr(c,block))out.push({raw:'<w:r><w:br/></w:r>'});continue;}
    if(c.tagName==='IMG'){out.push({raw:imageXml(c,ctx)});continue;}
    if(c.tagName==='HR'||c.tagName==='STYLE'||c.tagName==='SCRIPT')continue;
    if(c.tagName==='A'&&c.getAttribute('href')){
      const inner=runsOf([...c.childNodes],block,ctx);if(!inner)continue;const h=c.getAttribute('href');
      if(h.startsWith('#'))out.push({raw:`<w:hyperlink w:anchor="${esc(h.slice(1))}" w:history="1">${inner}</w:hyperlink>`});
      else out.push({raw:`<w:hyperlink r:id="${ctx.rel(RT.link,h,'External')}" w:history="1">${inner}</w:hyperlink>`});
      continue;
    }
    if(c.classList.contains('we-cm')){const id=c.dataset.cs||c.dataset.ce;const cm=E.cm&&E.cm.get(+id);if(!cm||cm.fresh)continue;
      out.push({raw:c.dataset.cs?`<w:commentRangeStart w:id="${id}"/>`:`<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`});continue;}
    if(c.classList.contains('we-fld')){const p={...propsAt(c,block,ctx),bg:null};const f=c.dataset.f==='NUMPAGES'?'NUMPAGES':'PAGE';
      out.push({raw:`<w:fldSimple w:instr=" ${f} \\* MERGEFORMAT "><w:r>${rprXml(p,rtlPara,styled)}<w:t>${xmlText(c.textContent||'1')}</w:t></w:r></w:fldSimple>`});continue;}
    if(c.getAttribute('contenteditable')==='false')continue;
    walk([...c.childNodes]);
  }};
  walk(nodes);
  return out.map(o=>o.raw!=null?o.raw:`<w:r>${o.rpr}${textXml(o.text)}</w:r>`).join('');
}
function atomXml(a,ctx){
  const at=E.st.atoms[+a.dataset.atom];if(!at)return'';
  let x=at.xml;const img=a.querySelector('img');const again=ctx.atomUsed.has(+a.dataset.atom);ctx.atomUsed.add(+a.dataset.atom);
  if((img&&img.dataset.rs&&at.draw)||again){
    x=x.cloneNode(true);
    if(img&&img.dataset.rs&&at.draw){const ex=x.getElementsByTagNameNS(NS.wp,'extent')[0];const cx=Math.round(img.offsetWidth*EMU),cy=Math.round(img.offsetHeight*EMU);
      if(ex){const ox=+ex.getAttribute('cx')||cx,oy=+ex.getAttribute('cy')||cy;ex.setAttribute('cx',cx);ex.setAttribute('cy',cy);
        for(const e of x.getElementsByTagNameNS(NS.a,'ext'))if(e.parentNode&&e.parentNode.localName==='xfrm'){const ex2=+e.getAttribute('cx'),ey2=+e.getAttribute('cy');e.setAttribute('cx',Math.round(ex2*cx/ox));e.setAttribute('cy',Math.round(ey2*cy/oy));}}}
    if(again)for(const dp of x.getElementsByTagNameNS(NS.wp,'docPr'))dp.setAttribute('id',ctx.nextDocPr++);
  }
  return ctx.clean(ser(x));
}
function imageXml(img,ctx){
  const d=ctx.imgs.get(img);if(!d){ctx.lost++;return'';}
  const w=img.offsetWidth||parseFloat(img.style.width)||d.w||100,h=img.offsetHeight||parseFloat(img.style.height)||d.h||100;
  const cx=Math.round(w*EMU),cy=Math.round(h*EMU);
  let n=1,path;do path=`media/smImg${n++}.${d.ext}`;while(ctx.files.has('word/'+path));
  put(ctx.files,'word/'+path,d.bytes);ctEnsureDefault(ctx.files,d.ext,d.ext==='jpeg'?'image/jpeg':'image/'+d.ext);
  const rid=ctx.rel(RT.image,path);const id=ctx.nextDocPr++;
  return`<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="תמונה ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${NS.a}" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="${NS.a}"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="${NS.pic}"><pic:nvPicPr><pic:cNvPr id="${id}" name="image${id}.${d.ext}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
}
async function imgBytes(img){
  try{
    const res=await fetch(img.src);const blob=await res.blob();const t=blob.type;
    if(/^image\/(png|jpeg|gif|bmp)$/.test(t))return{bytes:new Uint8Array(await blob.arrayBuffer()),ext:t.split('/')[1],w:img.naturalWidth,h:img.naturalHeight};
    const bm=await createImageBitmap(blob);const c=document.createElement('canvas');c.width=bm.width;c.height=bm.height;c.getContext('2d').drawImage(bm,0,0);
    const png=await new Promise(r=>c.toBlob(r,'image/png'));return{bytes:new Uint8Array(await png.arrayBuffer()),ext:'png',w:bm.width,h:bm.height};
  }catch(_){return null;}
}
function paraXml(el,ctx,opts={}){
  const key=el.dataset.src,ent=key?E.map.get(key):null;
  if(ent&&ent.kind==='p'&&!opts.list&&!ctx.used.has(key)&&el.outerHTML===ent.html){
    ctx.used.add(key);let xe=ent.xml;const sp=kid(kid(xe,'pPr'),'sectPr');
    if(sp&&ctx.last.get(key)!==el){xe=xe.cloneNode(true);const pp=kid(xe,'pPr');pp.removeChild(kid(pp,'sectPr'));}
    return ctx.clean(ser(xe));
  }
  let pPr,old=null;
  if(ent&&ent.kind==='p'){const t=kid(ent.xml,'pPr');pPr=t?t.cloneNode(true):wNew(E.st.norm,'pPr');old=ent.sig;
    if(ctx.last.get(key)!==el){const sp=kid(pPr,'sectPr');if(sp)pPr.removeChild(sp);}}
  else pPr=wNew(E.st.norm,'pPr');
  patchPPr(pPr,el,old,sigOf(el),ctx,opts);
  let pre='',post='';
  if(el.dataset&&el.dataset.toc==='e'){
    if(el.classList.contains('we-toc')){const tabs=wNew(E.st.norm,'tabs');tabs.appendChild(wNew(E.st.norm,'tab',{val:'right',leader:'dot',pos:Math.round(twip(contentW()))}));setChild(pPr,tabs,PPR_ORDER);}
    const pv=el.previousElementSibling,nx=el.nextElementSibling;
    if(!(pv&&pv.dataset&&pv.dataset.toc==='e'))pre='<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>';
    if(!(nx&&nx.dataset&&nx.dataset.toc==='e'))post='<w:r><w:fldChar w:fldCharType="end"/></w:r>';
  }
  const runs=runsOf([...el.childNodes].filter(n=>!(n.nodeType===1&&/^(UL|OL|TABLE)$/.test(n.tagName))),el,ctx);
  return`<w:p>${pPr.childNodes.length?ctx.clean(ser(pPr)):''}${pre}${runs}${post}</w:p>`;
}
function hrXml(el){const rtl=getComputedStyle(el.parentElement||E.body).direction==='rtl';
  return`<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr>${rtl?'<w:bidi/>':''}<w:spacing w:before="0" w:after="120"/></w:pPr></w:p>`;}
function listXml(list,ctx,lvl){
  let x='';const ordered=list.tagName==='OL';const id=ordered?(ctx.decUsed?ctx.newDec():(ctx.decUsed=true,E.st.num.decimal)):E.st.num.bullet;
  for(const li of list.children){
    if(li.tagName!=='LI'){x+=exportContainer({childNodes:[li]},ctx);continue;}
    const blocks=[...li.children].filter(c=>PARA.test(c.tagName)&&c.tagName!=='LI');
    if(blocks.length)blocks.forEach((b,i)=>{x+=paraXml(b,ctx,i===0?{list:{id,lvl}}:{});});
    else x+=paraXml(li,ctx,{list:{id,lvl}});
    for(const sub of li.children)if(/^(UL|OL)$/.test(sub.tagName))x+=listXml(sub,ctx,Math.min(8,lvl+1));
  }
  return x;
}
const hasBlockKid=n=>[...n.children].some(c=>/^(P|H[1-6]|LI|UL|OL|TABLE|DIV|BLOCKQUOTE|PRE|HR)$/.test(c.tagName));
/* A paragraph that holds whole paragraphs inside it (left by pasting, by the assistant or by a converted file)
   looks like separate lines on screen, but Word keeps only one paragraph: the lines ran together into one block
   and the page count stopped growing. Such a paragraph is split into real paragraphs, each keeping the
   outer paragraph's alignment, direction and look where it has none of its own. */
const BLOCK_TAGS=/^(P|H[1-6]|DIV|UL|OL|TABLE|BLOCKQUOTE|PRE|HR|SECTION|ARTICLE|FIGURE|CENTER)$/;
const BLOCK_SEL='p,h1,h2,h3,h4,h5,h6,div,ul,ol,table,blockquote,pre,hr,section,article,figure,center';
const INHERIT=['textAlign','direction','fontFamily','fontSize','color','lineHeight','marginTop','marginBottom','textIndent'];
function liftBlocks(scope){
  if(!scope||!scope.querySelector(':is(p,h1,h2,h3,h4,h5,h6) :is('+BLOCK_SEL+')'))return 0;
  const sel=getSelection(),r=sel&&sel.rangeCount&&scope.contains(sel.anchorNode)?sel.getRangeAt(0):null;
  const keep=r?[r.startContainer,r.startOffset,r.endContainer,r.endOffset]:null;
  const n=liftBlocksIn(scope);
  if(keep&&keep[0].isConnected&&keep[2].isConnected){try{const nr=document.createRange();nr.setStart(keep[0],keep[1]);nr.setEnd(keep[2],keep[3]);sel.removeAllRanges();sel.addRange(nr);}catch(_){}}
  return n;
}
function liftBlocksIn(scope){
  let n=0;
  const isEmpty=el=>!el.textContent.replace(ZW,'').trim()&&!el.querySelector('img,[data-atom],.we-pb,.we-fld');
  for(const p of [...scope.querySelectorAll('p,h1,h2,h3,h4,h5,h6')]){
    if(!p.isConnected||!p.querySelector(BLOCK_SEL))continue;
    if(p.closest('[data-atom],[contenteditable="false"]'))continue;
    // an inline wrapper (span, b…) around a block: its look goes onto the pieces, then it is taken away
    let guard=0;
    for(let w=[...p.querySelectorAll(BLOCK_SEL)].map(b=>b.parentElement).find(x=>x!==p&&!BLOCK_TAGS.test(x.tagName));w&&guard++<400;
        w=[...p.querySelectorAll(BLOCK_SEL)].map(b=>b.parentElement).find(x=>x!==p&&!BLOCK_TAGS.test(x.tagName))){
      const parts=[];
      for(const c of [...w.childNodes]){
        if(c.nodeType===1&&BLOCK_TAGS.test(c.tagName)){if(w.style.cssText)c.style.cssText=w.style.cssText+';'+c.style.cssText;parts.push(c);}
        else{const sh=w.cloneNode(false);sh.removeAttribute('id');sh.append(c);parts.push(sh);}
      }
      w.replaceWith(...parts);
    }
    const cs=p.style,out=[];let cur=null,first=true;
    const shells=new Set();
    const shell=()=>{const c=p.cloneNode(false);c.removeAttribute('id');if(!first)c.removeAttribute('data-src');first=false;shells.add(c);return c;};
    for(const c of [...p.childNodes]){
      if(c.nodeType===1&&BLOCK_TAGS.test(c.tagName)){
        cur=null;
        if(/^(P|H[1-6]|DIV|BLOCKQUOTE|PRE|CENTER)$/.test(c.tagName)){
          for(const k of INHERIT)if(cs[k]&&!c.style[k])c.style[k]=cs[k];
          if(!c.getAttribute('dir')&&p.getAttribute('dir'))c.setAttribute('dir',p.getAttribute('dir'));
          if(!c.className&&p.className&&c.tagName==='P')c.className=p.className;
        }
        out.push(c);
      }else{if(!cur){cur=shell();out.push(cur);}cur.append(c);}
    }
    const keep=out.filter(x=>!(shells.has(x)&&isEmpty(x)));
    p.replaceWith(...(keep.length?keep:out));
    n++;
  }
  return n;
}
function exportContainer(container,ctx){
  let x='';let inl=[];
  const host=container.nodeType===1?container:E.body;
  const flush=()=>{if(!inl.length)return;const nodes=inl;inl=[];
    if(!nodes.some(n=>n.nodeType===1||n.data.replace(ZW,'').trim()))return;
    const rtl=getComputedStyle(host).direction==='rtl';
    x+=`<w:p><w:pPr>${rtl?'<w:bidi/>':''}</w:pPr>${runsOf(nodes,host,ctx)}</w:p>`;};
  for(const n of [...container.childNodes]){
    if(n.nodeType===3){if(n.data.replace(ZW,'').trim()||inl.length)inl.push(n);continue;}
    if(n.nodeType!==1)continue;const tg=n.tagName;
    if(tg==='TABLE'){flush();x+=tableXml(n,ctx);}
    else if(tg==='UL'||tg==='OL'){flush();x+=listXml(n,ctx,0);}
    else if(tg==='HR'){flush();x+=hrXml(n);}
    else if(tg==='LI'){flush();x+=paraXml(n,ctx,{list:{id:E.st.num.bullet,lvl:0}});}
    else if(PARA.test(tg)||/^(DIV|SECTION|ARTICLE|HEADER|FOOTER|FIGURE|MAIN|ASIDE|NAV|CENTER|FIGCAPTION|DL|DT|DD)$/.test(tg)){flush();x+=(tg!=='P'&&hasBlockKid(n))?exportContainer(n,ctx):paraXml(n,ctx);}
    else if(tg==='TBODY'||tg==='THEAD'||tg==='TFOOT'||tg==='TR'){flush();}
    else inl.push(n);
  }
  flush();return x;
}
function cellXml(td,ctx,tcPr){
  let inner=exportContainer(td,ctx);
  if(!inner||!/<\/w:p>$|<w:p\/>$/.test(inner))inner+=`<w:p>${getComputedStyle(td).direction==='rtl'?'<w:pPr><w:bidi/></w:pPr>':''}</w:p>`;
  return`<w:tc>${tcPr||''}${inner}</w:tc>`;
}
const isCont=tc=>{const vm=kid(kid(tc,'tcPr'),'vMerge');return !!vm&&wAttr(vm,'val')!=='restart';};
function tableXml(el,ctx){
  const key=el.dataset.src,ent=key?E.map.get(key):null;
  if(ent&&ent.kind==='tbl'&&!ctx.used.has(key)&&el.outerHTML===ent.html){ctx.used.add(key);return ctx.clean(ser(ent.xml));}
  const rows=[...el.rows].filter(r=>r.closest('table')===el);
  if(ent&&ent.kind==='tbl'&&ent.ok&&!ctx.used.has(key)&&rows.length===ent.rows.length&&rows.every((r,i)=>r.cells.length===ent.rows[i].shown.length)){
    ctx.used.add(key);let x='<w:tbl>';
    for(const c of kids(ent.xml))if(c.localName==='tblPr'||c.localName==='tblGrid')x+=ctx.clean(ser(c));
    ent.rows.forEach((R,ri)=>{x+='<w:tr>';let si=0;
      for(const c of kids(R.tr)){
        if(c.localName!=='tc'){if(c.localName==='trPr'||c.localName==='tblPrEx')x+=ctx.clean(ser(c));continue;}
        if(isCont(c)){x+=ctx.clean(ser(c));if(R.all)si++;continue;}
        const td=rows[ri].cells[si++];let tcPr=kid(c,'tcPr');
        const bg0=ent.bg.get(td.dataset.cell),bg=hex(td.style.backgroundColor);
        if(td.dataset.cell&&bg0!==undefined&&bg!==bg0){tcPr=tcPr?tcPr.cloneNode(true):wNew(E.st.norm,'tcPr');for(const s of kids(tcPr,'shd'))tcPr.removeChild(s);
          if(bg){const s=wNew(E.st.norm,'shd',{val:'clear',color:'auto',fill:bg});const after=kids(tcPr).find(k=>['noWrap','tcMar','textDirection','tcFitText','vAlign','hideMark'].includes(k.localName));tcPr.insertBefore(s,after||null);}}
        x+=cellXml(td,ctx,tcPr?ctx.clean(ser(tcPr)):'');
      }
      x+='</w:tr>';});
    return x+'</w:tbl>';
  }
  // a new table, or one whose rows and columns changed
  const occ=[];let ncols=0;
  rows.forEach((tr,r)=>{occ[r]=occ[r]||[];let c=0;
    for(const td of tr.cells){while(occ[r][c])c++;const cs=Math.max(1,+td.getAttribute('colspan')||1),rs=Math.max(1,+td.getAttribute('rowspan')||1);
      for(let rr=r;rr<Math.min(rows.length,r+rs);rr++){occ[rr]=occ[rr]||[];for(let cc=c;cc<c+cs;cc++)occ[rr][cc]={td,c0:c,cs,rs,first:rr===r};}c+=cs;}
    ncols=Math.max(ncols,occ[r].length);});
  const tot=el.offsetWidth||contentW();const colW=new Array(ncols).fill(0);
  for(const row of occ)row.forEach((o,c)=>{if(o&&o.cs===1&&!colW[c])colW[c]=o.td.offsetWidth;});
  const known=colW.reduce((a,b)=>a+b,0),miss=colW.filter(w=>!w).length;colW.forEach((w,i)=>{if(!w)colW[i]=miss?Math.max(20,(tot-known)/miss):tot/ncols;});
  const rtl=getComputedStyle(el).direction==='rtl';
  const first=el.querySelector('td,th');const fb=first?getComputedStyle(first):null;const hasB=fb&&parseFloat(fb.borderTopWidth)>0&&fb.borderTopStyle!=='none';
  const B=v=>['top','left','bottom','right','insideH','insideV'].map(s=>`<w:${s} w:val="${v}" w:sz="4" w:space="0" w:color="auto"/>`).join('');
  let tblPr=ent&&ent.kind==='tbl'?kid(ent.xml,'tblPr'):null;
  let x='<w:tbl>'+(tblPr?ctx.clean(ser(tblPr)):`<w:tblPr>${rtl?'<w:bidiVisual/>':''}<w:tblW w:w="${twip(tot)}" w:type="dxa"/><w:tblBorders>${B(hasB?'single':'nil')}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr>`);
  x+='<w:tblGrid>'+colW.map(w=>`<w:gridCol w:w="${twip(w)}"/>`).join('')+'</w:tblGrid>';
  occ.forEach((row,r)=>{x+='<w:tr>';
    for(let c=0;c<ncols;c++){const o=row[c];
      if(!o){x+=`<w:tc><w:tcPr><w:tcW w:w="${twip(colW[c])}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`;continue;}
      if(o.c0!==c)continue;
      const w=colW.slice(c,c+o.cs).reduce((a,b)=>a+b,0);const td=o.td,cs=getComputedStyle(td);
      const tb=parseFloat(cs.borderTopWidth)>0&&cs.borderTopStyle!=='none';
      let pr=`<w:tcW w:w="${twip(w)}" w:type="dxa"/>${o.cs>1?`<w:gridSpan w:val="${o.cs}"/>`:''}`;
      if(o.rs>1||!o.first)pr+=o.first?'<w:vMerge w:val="restart"/>':'<w:vMerge/>';
      if(o.first){if(tb!==hasB)pr+=`<w:tcBorders>${['top','left','bottom','right'].map(s=>`<w:${s} w:val="${tb?'single':'nil'}" w:sz="4" w:space="0" w:color="auto"/>`).join('')}</w:tcBorders>`;
        const bg=hex(cs.backgroundColor);if(bg)pr+=`<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>`;
        const va=cs.verticalAlign;if(va==='middle')pr+='<w:vAlign w:val="center"/>';else if(va==='bottom')pr+='<w:vAlign w:val="bottom"/>';
        x+=cellXml(td,ctx,`<w:tcPr>${pr}</w:tcPr>`);}
      else x+=`<w:tc><w:tcPr>${pr}</w:tcPr><w:p/></w:tc>`;
    }
    x+='</w:tr>';});
  return x+'</w:tbl>';
}
function patchSect(sect){
  const d=sect.ownerDocument,pg=E.page;
  let sz=kid(sect,'pgSz')||setChild(sect,wNew(d,'pgSz'),SECT_ORDER);wSet(sz,'w',pg.w);wSet(sz,'h',pg.h);
  if(pg.landscape)wSet(sz,'orient','landscape');else sz.removeAttributeNS(W,'orient');
  let mg=kid(sect,'pgMar');if(!mg){mg=setChild(sect,wNew(d,'pgMar',{header:708,footer:708,gutter:0}),SECT_ORDER);}
  for(const k of ['top','right','bottom','left'])wSet(mg,k,pg[k]);
}
async function buildDocx(){
  if(E.hf)exitHF(true);
  if(E.snapT)snap();
  if(liftBlocks(E.main)){snap();guides();}
  const st=E.st;const files=new Map(st.files);
  let head=st.docStr.slice(0,st.docStr.search(/<w:body[\s>]/));
  const rootTag=head.match(/<w:document\b[^>]*>/);
  if(rootTag){let t=rootTag[0];for(const [p,u] of [['w',W],['r',NS.r],['wp',NS.wp]])if(!new RegExp(`xmlns:${p}=`).test(t))t=t.replace(/>$/,` xmlns:${p}="${u}">`).replace(/\/>$/,'>');head=head.replace(rootTag[0],t);}
  const decl=new Map();for(const m of (head.match(/<w:document\b[^>]*>/)||[''])[0].matchAll(/xmlns:(\w+)="([^"]+)"/g))decl.set(m[1],m[2]);
  const clean=s=>s.replace(/ xmlns:(\w+)="([^"]+)"/g,(m,p,u)=>decl.get(p)===u?'':m);
  const numMax=Math.max(0,...[...(txt(files,'word/numbering.xml')||'').matchAll(/<w:num\b[^>]*w:numId="(\d+)"/g)].map(m=>+m[1]));
  const ctx={files,used:new Set(),last:new Map(),imgs:new Map(),props:new Map(),atomUsed:new Set(),lost:0,clean,
    nextDocPr:Math.max(0,...[...st.docStr.matchAll(/<wp:docPr\b[^>]*\bid="(\d+)"/g)].map(m=>+m[1]))+1,
    numMax,extra:[],decUsed:false,curDec:null,
    rel:(type,target,mode)=>addRel(files,st.docPart,type,target,mode),
    newDec(){const id=++this.numMax;this.extra.push(`<w:num w:numId="${id}"><w:abstractNumId w:val="${st.num.decAbs}"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>`);return id;}};
  for(const el of E.main.querySelectorAll('[data-src]'))ctx.last.set(el.dataset.src,el);
  for(const img of E.main.querySelectorAll('img'))if(!img.closest('[data-atom]'))ctx.imgs.set(img,await imgBytes(img));
  let bodyXml=exportContainer(E.main,ctx);
  const newRefs=[];
  for(const k of ['header','footer'])if(hfHtml(k)!==E.hfOrig[k]){const r=await hfXml(k,files,st.docPart,st);if(r)newRefs.push(r);}
  const liveNow=[...E.cm.values()].filter(c=>cmLive(c)).map(c=>c.id).sort().join(',');
  if(E.cmDirty||liveNow!==E.cmOrigLive)cmXml(files,st.docPart);
  if(!bodyXml)bodyXml=`<w:p><w:pPr>${st.heb?'<w:bidi/>':''}</w:pPr></w:p>`;
  let sect=st.sect?st.sect.cloneNode(true):parseXml(`<w:sectPr xmlns:w="${W}"><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>`).documentElement;
  if(E.pageDirty)patchSect(sect);
  for(const r of newRefs.reverse()){const e=sect.ownerDocument.createElementNS(W,`w:${r.kind}Reference`);wSet(e,'type',r.type);e.setAttributeNS(NS.r,'r:id',r.rid);sect.insertBefore(e,sect.firstChild);}
  put(files,st.docPart,head+'<w:body>'+bodyXml+clean(ser(sect))+'</w:body></w:document>');
  if(ctx.extra.length){const n=txt(files,'word/numbering.xml');put(files,'word/numbering.xml',n.replace(/<\/w:numbering>\s*$/,ctx.extra.join('')+'</w:numbering>'));}
  const z=new JSZip();for(const [p,b] of files)z.file(p,b);
  const bytes=await z.generateAsync({type:'uint8array',compression:'DEFLATE',compressionOptions:{level:6}});
  return{bytes,lost:ctx.lost};
}


/* ---------- headers and footers: double-click to edit, like in Word ---------- */
const hfBox=k=>q(k==='header'?'weHdr':'weFtr');
const hfHtml=k=>{const el=hfBox(k).firstElementChild;return el?el.innerHTML:'';};
const hfChanged=()=>['header','footer'].some(k=>hfHtml(k)!==(E.hfClean||{})[k]);
function hfEl(k){
  const box=hfBox(k);let el=box.firstElementChild;
  if(!el){el=document.createElement(k);box.append(el);}
  if(!el.querySelector('p,table')){el.innerHTML='';const p=newPara();if(k==='footer')p.style.textAlign='center';el.append(p);}
  return el;
}
function enterHF(k){
  if(!E)return;if(E.hf){if(E.hf.kind===k)return;exitHF(true);}
  if(E.snapT)snap();hideImgBox();
  const el=hfEl(k);
  E.hf={kind:k,el,saved:{hist:E.hist,hi:E.hi,lastRange:E.lastRange,lastSel:E.lastSel}};
  E.body=el;E.hist=[];E.hi=-1;E.lastRange=null;E.lastSel=null;
  el.contentEditable='true';el.spellcheck=E.main.spellcheck;el.lang=E.main.lang;
  if(!el.dataset.bound){bindEditable(el);el.dataset.bound='1';}
  E.root.classList.add('we-hfmode');hfBox(k).classList.add('ed');
  const tab=E.root.querySelector('.we-tabs [data-tab=hf]');tab.hidden=false;showTab('hf');
  q('weHfGo').querySelector('span').textContent=k==='header'?'לכותרת התחתונה':'לכותרת העליונה';
  el.focus({preventScroll:true});const last=[...el.querySelectorAll('p')].pop()||el;caretTo(last,last.childNodes.length&&last.lastChild.tagName==='BR'?last.childNodes.length-1:last.childNodes.length);
  snap();updateUI();hint(k==='header'?'עריכת הכותרת העליונה. Esc או לחיצה כפולה על המסמך לחזרה':'עריכת הכותרת התחתונה. Esc או לחיצה כפולה על המסמך לחזרה');
}
function exitHF(noFocus){
  const h=E&&E.hf;if(!h)return;if(E.snapT)snap();
  h.el.contentEditable='false';hfBox(h.kind).classList.remove('ed');E.root.classList.remove('we-hfmode');
  E.body=E.main;E.hist=h.saved.hist;E.hi=h.saved.hi;E.lastRange=h.saved.lastRange;E.lastSel=h.saved.lastSel;E.hf=null;
  const tab=E.root.querySelector('.we-tabs [data-tab=hf]');tab.hidden=true;if(tab.getAttribute('aria-selected')==='true')showTab('home');
  markDirty();updateUI();
  if(!noFocus){E.main.focus({preventScroll:true});if(E.lastRange){const s=getSelection();s.removeAllRanges();s.addRange(E.lastRange);}}
}
const fld=(f,style)=>`<span class="we-fld" data-f="${f}" contenteditable="false"${style?` style="${style}"`:''}>${f==='PAGE'?'1':(E.pageCount||1)}</span>`;
// a page number in the header or the footer, at the side the person chose
function pageNumber(where,al,fmt){
  enterHF(where);const el=E.body;
  mutate(()=>{
    for(const f of [...el.querySelectorAll('.we-fld')]){const p=f.closest('p');f.remove();if(p&&!p.textContent.replace(ZW,'').replace(/עמוד|מתוך|Page|of|[\s\-–|]/g,'')&&!p.querySelector('img,.we-fld'))p.remove();}
    const p=newPara();p.innerHTML='';p.style.textAlign=al;
    p.innerHTML=fmt==='xofy'?`עמוד ${fld('PAGE')} מתוך ${fld('NUMPAGES')}`:fld('PAGE');
    const empty=[...el.children].filter(x=>x.tagName==='P'&&!x.textContent.replace(ZW,'').trim()&&!x.querySelector('img,.we-fld,[data-atom]'));
    if(empty.length===[...el.children].length){el.innerHTML='';el.append(p);}else if(where==='header')el.prepend(p);else el.append(p);
    caretTo(p,p.childNodes.length);
  });
}
function removePageNumbers(){
  if(E.hf)exitHF(true);let n=0;
  for(const k of ['header','footer']){const el=hfBox(k).firstElementChild;if(!el)continue;
    for(const f of [...el.querySelectorAll('.we-fld')]){const p=f.closest('p');f.remove();n++;if(p&&!p.textContent.replace(ZW,'').replace(/עמוד|מתוך|Page|of|[\s\-–|]/g,'')&&!p.querySelector('img,.we-fld'))p.remove();}}
  markDirty();hint(n?'מספרי העמודים הוסרו':'אין מספרי עמודים במסמך');
}
function hfField(f){if(!E.hf)return;mutate(()=>{const r=liveRange();if(!r)return;r.deleteContents();const t=document.createElement('template');t.innerHTML=fld(f);const n=t.content.firstChild;r.insertNode(n);caretTo(n.parentNode,[...n.parentNode.childNodes].indexOf(n)+1);});}
async function hfXml(k,files,docPart,st){
  const el=hfBox(k).firstElementChild;const tag=k==='header'?'w:hdr':'w:ftr';let path=st.hf[k].path,ref=null;
  if(!path){let n=1;do path=`word/${k}${n++}.xml`;while(files.has(path));
    ctEnsureOverride(files,path,OX+`wordprocessingml.${k}+xml`);ref={kind:k,type:st.hf[k].type,rid:addRel(files,docPart,RT_HF(k),path.replace(/^word\//,''))};}
  const hctx={files,used:new Set(),last:new Map(),imgs:new Map(),props:new Map(),atomUsed:new Set(),lost:0,clean:x=>x,numMax:0,extra:[],decUsed:false,
    nextDocPr:5000+(k==='footer'?500:0),rel:(type,target,mode)=>addRel(files,path,type,target,mode),newDec(){return E.st.num.decimal;}};
  if(el)for(const img of el.querySelectorAll('img'))if(!img.closest('[data-atom]'))hctx.imgs.set(img,await imgBytes(img));
  const inner=(el&&exportContainer(el,hctx))||`<w:p><w:pPr>${st.heb?'<w:bidi/>':''}</w:pPr></w:p>`;
  put(files,path,XML_HEAD+`<${tag} xmlns:w="${W}" xmlns:r="${NS.r}" xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}">${inner}</${tag}>`);
  return ref;
}

/* ---------- comments in the margin ---------- */
function cmMarker(k,id){const s=document.createElement('span');s.className='we-cm';s.dataset[k]=id;s.contentEditable='false';return s;}
const cmFind=(k,id)=>E.main.querySelector(`.we-cm[data-${k}="${id}"]`);
const cmLive=c=>!!cmFind('cs',c.id)||!!cmFind('ce',c.id);
const cmAuthor=()=>{try{const a=E.opts.author&&E.opts.author();if(a)return String(a);}catch(_){}return 'אני';};
const initials=n=>String(n||'?').trim().split(/\s+/).map(w=>w[0]).join('').slice(0,2).toUpperCase();
function cmRange(c){const a=cmFind('cs',c.id),b=cmFind('ce',c.id);if(!a&&!b)return null;const r=document.createRange();
  if(a)r.setStartAfter(a);else r.setStartBefore(b);if(b)r.setEndBefore(b);else r.setEndAfter(a);return r;}
function cmPaint(){
  if(!window.CSS||!window.CSS.highlights||!window.Highlight)return;
  const all=[],on=[];for(const c of E.cm.values()){if(c.done)continue;const r=cmRange(c);if(!r)continue;(c.id===E.cmOn?on:all).push(r);}
  window.CSS.highlights.set('we-cmt',new Highlight(...all));window.CSS.highlights.set('we-cmt-on',new Highlight(...on));
}
function cmList(){return [...E.cm.values()].filter(c=>!c.parent&&cmLive(c)).map(c=>{const m=cmFind('cs',c.id)||cmFind('ce',c.id);return{c,y:offIn(m,q('weSheet')).y};}).sort((a,b)=>a.y-b.y);}
function cmRender(){
  if(!E)return;const pane=q('weCPane');
  for(const ta of pane.querySelectorAll('textarea')){const c=E.cm.get(+ta.dataset.id);if(c)c.text=ta.value;}const list=cmList();
  E.root.classList.toggle('we-hascm',list.length>0&&!E.cmHide);
  if(!list.length||E.cmHide){pane.innerHTML='';cmPaint();return;}
  const fmtD=d=>{const t=Date.parse(d);return isNaN(t)?'':new Date(t).toLocaleString('he-IL',{day:'numeric',month:'numeric',year:'2-digit',hour:'2-digit',minute:'2-digit'});};
  const one=(c,reply)=>`<div class="cmh"><i class="av">${esc(c.initials||initials(c.author))}</i><b>${esc(c.author||'ללא שם')}</b><small>${fmtD(c.date)}</small></div>`+
    (c.editing?`<textarea class="cmt" data-id="${c.id}" rows="2" dir="auto" placeholder="${reply?'כתבו תשובה…':'כתבו הערה…'}">${esc(c.text)}</textarea><div class="cma"><button type="button" class="wb txt" data-ca="cancel" data-id="${c.id}">ביטול</button><button type="button" class="wb txt pri" data-ca="post" data-id="${c.id}">פרסום</button></div>`
      :`<div class="cmx" dir="auto">${esc(c.text).replace(/\n/g,'<br>')}</div>`);
  pane.innerHTML=list.map(({c})=>{const reps=[...E.cm.values()].filter(x=>x.parent===c.id&&cmLive(x));
    return`<div class="cmc${c.id===E.cmOn?' on':''}${c.done?' done':''}" data-id="${c.id}">${c.done?'<div class="cmd">✓ סומנה כטופלה</div>':''}${one(c)}${reps.map(r=>`<div class="cmr" data-id="${r.id}">${one(r,true)}</div>`).join('')}
      ${c.editing||reps.some(r=>r.editing)?'':`<div class="cma"><button type="button" class="wb txt" data-ca="reply" data-id="${c.id}">תשובה</button><button type="button" class="wb txt" data-ca="done" data-id="${c.id}">${c.done?'פתיחה מחדש':'סימון כטופל'}</button><button type="button" class="wb txt" data-ca="del" data-id="${c.id}">מחיקה</button></div>`}</div>`;}).join('');
  let y=0;const cards=[...pane.children];
  cards.forEach((card,i)=>{const want=list[i].y*(E.zoom||1)-6;const top=Math.max(want,y);card.style.top=top+'px';y=top+card.offsetHeight+8;});
  pane.style.height=y+'px';cmPaint();
  const ta=pane.querySelector('textarea');if(ta&&E.cmFocus){E.cmFocus=false;ta.focus();ta.setSelectionRange(ta.value.length,ta.value.length);}
}
function cmNew(){
  if(E.hf)exitHF();const r=liveRange();if(!r||!E.main.contains(r.commonAncestorContainer))return hint('מקמו את הסמן במסמך');
  if(r.collapsed){const w=wordAround(r);if(w){r.setStart(w.startContainer,w.startOffset);r.setEnd(w.endContainer,w.endOffset);}}
  const id=Math.max(0,...[...E.cm.keys()])+1;
  mutate(()=>{const end=r.cloneRange();end.collapse(false);end.insertNode(cmMarker('ce',id));const st=r.cloneRange();st.collapse(true);st.insertNode(cmMarker('cs',id));
    const s=getSelection();s.removeAllRanges();});
  const a=cmAuthor();E.cm.set(id,{id,author:a,initials:initials(a),date:new Date().toISOString().replace(/\.\d+Z$/,'Z'),text:'',parent:null,done:false,editing:true,fresh:true});
  E.cmOn=id;E.cmHide=false;E.cmDirty=true;E.cmFocus=true;cmRender();markDirty();
}
function cmReply(pid){
  const p=E.cm.get(pid);if(!p)return;const id=Math.max(0,...[...E.cm.keys()])+1;const a=cmFind('cs',pid),b=cmFind('ce',pid);
  if(a)a.after(cmMarker('cs',id));if(b)b.after(cmMarker('ce',id));if(E.snapT)snap();snap();
  const au=cmAuthor();E.cm.set(id,{id,author:au,initials:initials(au),date:new Date().toISOString().replace(/\.\d+Z$/,'Z'),text:'',parent:pid,done:false,editing:true,fresh:true});
  E.cmOn=pid;E.cmDirty=true;E.cmFocus=true;cmRender();markDirty();
}
function cmRemove(id){
  const kill=[id,...[...E.cm.values()].filter(x=>x.parent===id).map(x=>x.id)];
  mutate(()=>{for(const k of kill)for(const m of E.main.querySelectorAll(`.we-cm[data-cs="${k}"],.we-cm[data-ce="${k}"]`))m.remove();});
  for(const k of kill)E.cm.delete(k);if(E.cmOn===id)E.cmOn=null;E.cmDirty=true;cmRender();markDirty();
}
function cmAction(a,id){
  const c=E.cm.get(id);if(!c)return;
  if(a==='post'){const ta=q('weCPane').querySelector(`textarea[data-id="${id}"]`);const t=(ta?ta.value:'').trim();if(!t){cmAction('cancel',id);return;}c.text=t;c.editing=false;c.fresh=false;c.date=new Date().toISOString().replace(/\.\d+Z$/,'Z');E.cmDirty=true;}
  else if(a==='cancel'){if(c.fresh){for(const m of E.main.querySelectorAll(`.we-cm[data-cs="${id}"],.we-cm[data-ce="${id}"]`))m.remove();E.cm.delete(id);snap();}else c.editing=false;}
  else if(a==='reply'){cmReply(id);return;}
  else if(a==='done'){c.done=!c.done;E.cmDirty=true;}
  else if(a==='del'){cmRemove(id);return;}
  else if(a==='edit'){c.editing=true;E.cmFocus=true;}
  cmRender();markDirty();
}
function cmGo(dir){
  const list=cmList();if(!list.length)return hint('אין הערות במסמך');
  let i=list.findIndex(x=>x.c.id===E.cmOn);i=i<0?(dir>0?0:list.length-1):(i+dir+list.length)%list.length;
  cmSelect(list[i].c.id,true);
}
function cmSelect(id,scroll){
  E.cmOn=id;cmRender();const c=E.cm.get(id);const r=c&&cmRange(c);if(!r)return;
  if(scroll){const m=cmFind('cs',id)||cmFind('ce',id);m.parentElement.scrollIntoView({block:'center',behavior:'smooth'});}
}
function cmCurrent(){const r=curRange();if(!r)return E.cmOn;for(const c of E.cm.values()){if(c.parent)continue;const x=cmRange(c);if(x&&x.comparePoint(r.startContainer,r.startOffset)===0)return c.id;}return E.cmOn;}
function cmXml(files,docPart){
  const live=[...E.cm.values()].filter(c=>cmLive(c)&&!c.fresh||c.keep).sort((a,b)=>a.id-b.id);
  const pidOf=c=>(0x10000000+c.id*7919).toString(16).toUpperCase().padStart(8,'0').slice(-8);
  const rl=relsOf(files,docPart);let r1=rl.find(r=>r.type===RT_CM);let path=r1?resolveTarget(docPart,r1.target):'word/comments.xml';
  if(!live.length&&!r1)return;
  const para=(t,i,last,c)=>`<w:p${last?` w14:paraId="${pidOf(c)}" w14:textId="77777777"`:''}><w:pPr>${HEB.test(t)||(!t&&E.st.heb)?'<w:bidi/>':''}</w:pPr>${i===0?'<w:r><w:annotationRef/></w:r>':''}${t?`<w:r><w:rPr>${HEB.test(t)?'<w:rtl/>':''}</w:rPr>${textXml(t)}</w:r>`:''}</w:p>`;
  const body=live.map(c=>{const lines=(c.text||'').split('\n');return`<w:comment w:id="${c.id}" w:author="${esc(c.author||'')}" w:date="${esc(c.date||new Date().toISOString().replace(/\.\d+Z$/,'Z'))}" w:initials="${esc(c.initials||initials(c.author))}">${lines.map((t,i)=>para(t,i,i===lines.length-1,c)).join('')}</w:comment>`;}).join('');
  const NSX=`xmlns:w="${W}" xmlns:w14="${W14}" xmlns:w15="${W15}" xmlns:mc="${NS.mc}" mc:Ignorable="w14 w15"`;
  put(files,path,XML_HEAD+`<w:comments ${NSX}>${body}</w:comments>`);
  if(!r1){ctEnsureOverride(files,path,OX+'wordprocessingml.comments+xml');addRel(files,docPart,RT_CM,'comments.xml');}
  const r2=rl.find(r=>r.type===RT_CMX);const xpath=r2?resolveTarget(docPart,r2.target):'word/commentsExtended.xml';
  put(files,xpath,XML_HEAD+`<w15:commentsEx ${NSX}>${live.map(c=>{const par=c.parent&&E.cm.get(c.parent);return`<w15:commentEx w15:paraId="${pidOf(c)}"${par?` w15:paraIdParent="${pidOf(par)}"`:''} w15:done="${c.done?1:0}"/>`;}).join('')}</w15:commentsEx>`);
  if(!r2){ctEnsureOverride(files,xpath,OX+'wordprocessingml.commentsExtended+xml');addRel(files,docPart,RT_CMX,xpath.replace(/^word\//,''));}
  // newer Word keeps extra lists keyed to the old paragraph ids; they are rebuilt by Word when missing
  for(const r of rl)if(/relationships\/(commentsIds|commentsExtensible)$/.test(r.type)){const pth=resolveTarget(docPart,r.target);files.delete(pth);ctDropOverride(files,pth);dropRel(files,docPart,r.id);}
}

/* ---------- table of contents ---------- */
function headingLevel(b){
  if(/^H[1-6]$/.test(b.tagName))return +b.tagName[1];
  const c=styleCls(b);if(!c)return 0;const s=E.st.styles.find(x=>classKey(x.id)===c);if(!s)return 0;
  const m=/^heading (\d)$/.exec(s.name);if(m)return +m[1];
  let ol=null;if(s.el){const o=s.el.getElementsByTagNameNS(W,'outlineLvl')[0];if(o)ol=+wAttr(o,'val');}return ol!==null&&ol<9?ol+1:0;
}
const isTocP=b=>!!b&&b.nodeType===1&&(b.dataset.toc==='e'||b.dataset.toc==='t'||/(^|\s)wx_toc\d/.test(b.className)||/(^|\s)wx_tocheading/.test(b.className));
function tocItems(){
  guides();const out=[];const marks=E.pageMarks||[];
  for(const b of E.main.querySelectorAll('p,h1,h2,h3,h4,h5,h6')){if(isTocP(b)||b.closest('td,th'))continue;const lv=headingLevel(b);if(!lv||lv>3)continue;
    const t=b.textContent.replace(ZW,'').replace(/\s+/g,' ').trim();if(!t)continue;
    const y=offIn(b,E.main).y;out.push({t,lv,page:1+marks.filter(m=>m<=y+1).length});}
  return out;
}
function tocBlock(items,dir){
  const fr=document.createDocumentFragment();
  const h=document.createElement('p');h.dataset.toc='t';h.style.cssText=`direction:${dir};font-size:16pt;font-weight:700;color:#1F3864;margin:12pt 0 6pt`;h.textContent='תוכן עניינים';fr.append(h);
  if(!items.length){const p=document.createElement('p');p.dataset.toc='e';p.dataset.lvl='1';p.style.direction=dir;p.textContent='אין כותרות במסמך. החילו "כותרת 1" או "כותרת 2" על הכותרות ועדכנו.';fr.append(p);return fr;}
  for(const it of items){
    const p=document.createElement('p');p.dataset.toc='e';p.dataset.lvl=it.lv;p.className='we-toc';
    p.style.cssText=`direction:${dir};margin:0 0 3pt;margin-inline-start:${(it.lv-1)*14}pt;${it.lv===1?'font-weight:700;':''}`;
    const tx=document.createElement('span');tx.textContent=it.t;
    const tab=document.createElement('span');tab.className='we-tab '+P+'-tab-stop we-lead';tab.contentEditable='false';tab.textContent='\t';
    const pg=document.createElement('span');pg.textContent=it.page;
    p.append(tx,tab,pg);fr.append(p);
  }
  return fr;
}
function tocGroup(){
  const first=[...E.main.children].find(isTocP);if(!first)return null;const g=[first];let n=first.nextElementSibling;while(n&&isTocP(n)){g.push(n);n=n.nextElementSibling;}return g;
}
function tocInsert(){
  if(E.hf)exitHF();
  const dir=E.st.heb?'rtl':'ltr';const g=tocGroup();
  mutate(()=>{const items=tocItems();const blk=tocBlock(items,dir);
    if(g){g[0].before(blk);g.forEach(x=>x.remove());hint('תוכן העניינים עודכן');}
    else{const r=liveRange();let b=r&&E.main.contains(r.startContainer)?blockOf(r.startContainer):null;
      const atStart=b&&(()=>{const x=document.createRange();x.setStart(b,0);x.setEnd(r.startContainer,r.startOffset);return !x.toString().replace(ZW,'').trim();})();
      if(!b||atStart){let top=b;while(top&&top.parentNode!==E.main)top=top.parentNode;E.main.insertBefore(blk,top||E.main.firstChild);}
      else{const holder=document.createElement('p');const nb=insertBlockAfterCaret(holder);holder.replaceWith(blk);if(nb)caretIn(nb);}
      hint('תוכן העניינים נוסף. אחרי שינוי בכותרות: הוספה ← תוכן עניינים ← עדכון');}});
}
function tocRemove(){const g=tocGroup();if(!g)return hint('אין תוכן עניינים במסמך');mutate(()=>g.forEach(x=>x.remove()));}
function tocGo(p){
  const t=p.querySelector('span')?p.querySelector('span').textContent.trim():p.textContent.replace(/\t.*$/,'').trim();
  const h=[...E.main.querySelectorAll('p,h1,h2,h3,h4,h5,h6')].find(b=>!isTocP(b)&&headingLevel(b)&&b.textContent.replace(ZW,'').replace(/\s+/g,' ').trim()===t);
  if(h){h.scrollIntoView({block:'start',behavior:'smooth'});caretIn(h);}
}

/* ---------- automatic corrections while typing, like in Word ---------- */
const AC_WORDS={teh:'the',adn:'and',recieve:'receive',wich:'which',thier:'their',alot:'a lot',becuase:'because',dont:"don’t",cant:"can’t",wont:"won’t",didnt:"didn’t",doesnt:"doesn’t",isnt:"isn’t",im:"I’m",ive:"I’ve",
  definately:'definitely',seperate:'separate',occured:'occurred',untill:'until',tommorow:'tomorrow',goverment:'government',freind:'friend',beleive:'believe',accross:'across',wierd:'weird'};
const AC_SYM={'(c)':'©','(r)':'®','(tm)':'™','...':'…','->':'→','<-':'←','=>':'⇒','<=>':'⇔','1/2':'½','1/4':'¼','3/4':'¾','+-':'±'};
let acOn=true;try{acOn=localStorage.getItem('siman-autocorrect')!=='0';}catch(_){}
function autoCorrect(ch){
  if(!acOn)return false;const s=getSelection();if(!s||!s.rangeCount)return false;const r=s.getRangeAt(0);if(!r.collapsed)return false;
  const n=r.startContainer,o=r.startOffset;if(n.nodeType!==3||o<1||n.data[o-1]!==ch)return false;
  const t=n.data,bf=t.slice(0,o);const HE=/[\u05D0-\u05EA]/,LA=/[A-Za-z]/;let at=-1,len=0,to='';
  if(HE.test(ch)&&o>=3&&bf[o-2]==='"'&&HE.test(bf[o-3])){at=o-2;len=1;to='״';}
  else if(ch==="'"&&o>=2&&HE.test(bf[o-2])){at=o-1;len=1;to='׳';}
  else if(ch==='"'&&o>=2&&LA.test(bf[o-2])){at=o-1;len=1;to='”';}
  else if(ch==='"'&&(o===1||/[\s(\[]/.test(bf[o-2]))&&LA.test(blockOf(n)?blockOf(n).textContent.slice(-40):'')&&!HE.test(blockOf(n)?blockOf(n).textContent.slice(-40):'')){at=o-1;len=1;to='“';}
  else if(ch==="'"&&o>=2&&LA.test(bf[o-2])){at=o-1;len=1;to='’';}
  else if(/[\s.,;:!?)]/.test(ch)||ch==='\u00a0'){
    const m=/(^|[\s(\[“"])(\S+?)([\s.,;:!?)\u00a0])$/.exec(bf);
    if(ch===' '&&/ --$/.test(bf.slice(0,-1))){at=o-3;len=2;to='–';}
    else if(m){const w=m[2],ws=o-1-w.length;const lw=w.toLowerCase();
      if(AC_SYM[lw]){at=ws;len=w.length;to=AC_SYM[lw];}
      else if(/^[a-z]+$/.test(w)&&AC_WORDS[w]){at=ws;len=w.length;to=AC_WORDS[w];}
      else if(w==='i'){at=ws;len=1;to='I';}
      else if(/^[a-z][a-z’']*$/.test(w)){
        const blk=blockOf(n);if(blk){const pre=document.createRange();pre.setStart(blk,0);pre.setEnd(n,ws);const before=pre.toString().replace(ZW,'');
          if(/(^|[.!?]\s+)$/.test(before)&&!/\b(e\.g|i\.e|etc|vs)\.\s+$/i.test(before)&&!HE.test(before.slice(-3))){at=ws;len=1;to=w[0].toUpperCase();}}}
    }
  }
  if(at<0)return false;
  const old=t.substr(at,len);n.data=t.slice(0,at)+to+t.slice(at+len);const c=o-len+to.length;
  const rr=document.createRange();rr.setStart(n,c);rr.collapse(true);s.removeAllRanges();s.addRange(rr);E.lastRange=rr.cloneRange();
  E.lastAC={n,at,old,to,c};return true;
}
function undoAC(){const a=E.lastAC;E.lastAC=null;if(!a||!a.n.isConnected||a.n.data.substr(a.at,a.to.length)!==a.to)return false;
  a.n.data=a.n.data.slice(0,a.at)+a.old+a.n.data.slice(a.at+a.to.length);const c=a.c-a.to.length+a.old.length;
  const r=document.createRange();r.setStart(a.n,Math.min(c,a.n.length));r.collapse(true);const s=getSelection();s.removeAllRanges();s.addRange(r);snapLater();markDirtySoon();hint('התיקון האוטומטי בוטל');return true;}

/* ---------- word count ---------- */
function countText(t){t=t.replace(ZW,'');return{w:(t.match(/[\p{L}\p{N}][\p{L}\p{N}\p{M}'"״׳.-]*/gu)||[]).length,c:t.replace(/\s/g,'').length,cs:t.replace(/[\r\n]/g,'').length};}
function wordCount(){
  const r=curRange();const sel=r&&!r.collapsed?countText(r.toString()):null;const all=countText(E.main.innerText||E.main.textContent);
  const paras=[...E.main.querySelectorAll('p,h1,h2,h3,h4,h5,h6,li')].filter(p=>p.textContent.replace(ZW,'').trim()).length;
  const n=v=>v.toLocaleString('he-IL');
  const row=(l,k)=>`<tr><td>${l}</td>${sel?`<td dir="ltr">${n(sel[k])}</td>`:''}<td dir="ltr">${n(all[k])}</td></tr>`;
  dialog(`<h3>ספירת מילים</h3><table class="wc"><thead><tr><th></th>${sel?'<th>בבחירה</th>':''}<th>במסמך</th></tr></thead><tbody>
    <tr><td>עמודים</td>${sel?'<td></td>':''}<td dir="ltr">${n(E.pageCount||1)}</td></tr>${row('מילים','w')}${row('תווים (בלי רווחים)','c')}${row('תווים (כולל רווחים)','cs')}
    <tr><td>פסקאות</td>${sel?'<td></td>':''}<td dir="ltr">${n(paras)}</td></tr></tbody></table><div class="acts"><button type="button" class="pri" data-a="ok">סגירה</button></div>`,()=>true);
}

/* ---------- extras: right-click menu, automatic lists, drafts ---------- */
function ctxMenu(e){
  if(e.shiftKey)return; // Shift + right click keeps the browser's menu (spelling suggestions)
  const r0=curRange();
  if(!r0||r0.collapsed){const cr=document.caretRangeFromPoint&&document.caretRangeFromPoint(e.clientX,e.clientY);if(cr&&E.body.contains(cr.startContainer)){const s=getSelection();s.removeAllRanges();s.addRange(cr);E.lastRange=cr.cloneRange();}}
  e.preventDefault();
  const inCell=!!cellOf(),r=curRange(),hasSel=r&&!r.collapsed;
  const n=r&&(r.startContainer.nodeType===3?r.startContainer.parentElement:r.startContainer);const link=n&&n.closest&&n.closest('a[href]');
  const items=[['cut','גזירה','Ctrl+X',!hasSel],['copy','העתקה','Ctrl+C',!hasSel],['paste','הדבקה','Ctrl+V'],'-',
    ['bold','מודגש','Ctrl+B'],['italic','נטוי','Ctrl+I'],['underline','קו תחתי','Ctrl+U'],['clear','ניקוי העיצוב','Ctrl+Space'],'-',
    ['link',link?'עריכת הקישור…':'הוספת קישור…','Ctrl+K'],['cmNew','הערה חדשה','Ctrl+Alt+M'],['bullets','תבליטים'],['numbers','מספור'],'-',['right','יישור לימין'],['center','מרכוז'],['left','יישור לשמאל']];
  if(inCell)items.push('-',['rowAbove','הוספת שורה מעל'],['rowBelow','הוספת שורה מתחת'],['colRight','הוספת עמודה מימין'],['colLeft','הוספת עמודה משמאל'],['delRow','מחיקת השורה'],['delCol','מחיקת העמודה'],['delTable','מחיקת הטבלה']);
  const html=items.map(it=>it==='-'?'<div class="sep"></div>':`<button type="button" class="mi" data-v="${it[0]}"${it[3]?' disabled style="opacity:.4"':''}>${it[1]}${it[2]?`<small dir="ltr">${it[2]}</small>`:''}</button>`).join('')+'<div class="sep"></div><div style="font-size:11px;color:#7A879E;padding:2px 10px 4px">Shift + לחיצה ימנית: תיקוני איות</div>';
  const anchor={getBoundingClientRect:()=>({left:e.clientX,right:e.clientX,top:e.clientY,bottom:e.clientY}),contains:()=>false};
  const pp=pop(anchor,html,v=>{const f=CMD[v];if(f){liveRange();f();}});
  pp.style.left=Math.max(6,Math.min(e.clientX,innerWidth-pp.offsetWidth-6))+'px';pp.style.top=Math.max(6,Math.min(e.clientY,innerHeight-pp.offsetHeight-6))+'px';
}
// typing "1. " or "- " at the start of a paragraph starts a list, like in Word
function autoList(){
  const r=curRange();if(!r||!r.collapsed)return false;const b=blockOf(r.startContainer);if(!b||numOf(b)||b.closest('td,th'))return false;
  const pre=document.createRange();pre.setStart(b,0);pre.setEnd(r.startContainer,r.startOffset);const t=pre.toString().replace(ZW,'');
  const m=/^\s*(?:(\d{1,3})[.)]|([-*•]))[\s\u00a0]$/.exec(t);if(!m)return false;
  if(E.snapT)snap();
  pre.deleteContents();toggleList(m[2]?'b':'n');
  return true;
}
const DRAFT={db:null,
  open(){return this.db||(this.db=new Promise((res,rej)=>{const r=indexedDB.open('siman-drafts',1);r.onupgradeneeded=()=>r.result.createObjectStore('d');r.onsuccess=()=>res(r.result);r.onerror=()=>rej(r.error);}));},
  async put(k,v){try{const db=await this.open();await new Promise(r=>{const t=db.transaction('d','readwrite');t.objectStore('d').put(v,k);t.oncomplete=t.onerror=r;});}catch(_){}},
  async get(k){try{const db=await this.open();return await new Promise(r=>{const q2=db.transaction('d').objectStore('d').get(k);q2.onsuccess=()=>r(q2.result||null);q2.onerror=()=>r(null);});}catch(_){return null;}},
  async del(k){try{const db=await this.open();await new Promise(r=>{const t=db.transaction('d','readwrite');t.objectStore('d').delete(k);t.oncomplete=t.onerror=r;});}catch(_){}}
};
const draftKey=()=>'w:'+(E.opts.name||'new');
function startDrafts(){E.draftT=setInterval(async()=>{if(!E||!E.dirty||E.saving)return;try{const{bytes}=await buildDocx();if(E)await DRAFT.put(draftKey(),{bytes,time:Date.now(),name:fileName()});}catch(_){}},40000);}

/* ------------------------------------------------------------ opening */
function mapContainer(domC,xmlC,prefix){
  const d=[...domC.children].filter(e=>e.tagName==='P'||e.tagName==='TABLE');
  const x=kids(xmlC,'p','tbl');
  if(d.length!==x.length){E.mapMiss++;return;}
  d.forEach((el,i)=>{const key=prefix+i,xe=x[i];
    if(el.tagName==='P'&&xe.localName==='p'){el.dataset.src=key;E.map.set(key,{kind:'p',xml:xe});}
    else if(el.tagName==='TABLE'&&xe.localName==='tbl'){el.dataset.src=key;const ent={kind:'tbl',xml:xe,ok:false,bg:new Map()};E.map.set(key,ent);mapTable(el,xe,key,ent);}
    else E.mapMiss++;});
}
function mapTable(el,xe,key,ent){
  const dr=[...el.rows].filter(r=>r.closest('table')===el),xr=kids(xe,'tr');
  if(dr.length!==xr.length)return;
  const shown=[],all=[];
  for(let i=0;i<xr.length;i++){const tcs=kids(xr[i],'tc'),nc=tcs.filter(tc=>!isCont(tc)),n=dr[i].cells.length;
    if(n===tcs.length){shown.push(tcs);all.push(true);}else if(n===nc.length){shown.push(nc);all.push(false);}else return;}
  ent.ok=true;ent.rows=xr.map((tr,i)=>({tr,shown:shown[i],all:all[i]}));
  dr.forEach((row,r)=>[...row.cells].forEach((td,c)=>{const ck=`${key}.${r}.${c}`;td.dataset.cell=ck;ent.bg.set(ck,hex(td.style.backgroundColor));mapContainer(td,shown[r][c],ck+'.');}));
}
async function mount(){
  const st=E.st,body=E.body,sheet=q('weSheet');
  E.page={...st.page};applyPage();
  body.style.fontFamily=cssFont(st.def.font);body.style.fontSize=st.def.size+'pt';
  for(const b of [q('weHdr'),q('weFtr')])Object.assign(b.style,{fontFamily:cssFont(st.def.font),fontSize:st.def.size+'pt',color:'#000',lineHeight:'normal',textAlign:'start'});
  if(st.themeCs.minor)body.style.setProperty('--docx-minorBidi-font',`"${st.themeCs.minor}"`);
  if(st.themeCs.major)body.style.setProperty('--docx-majorBidi-font',`"${st.themeCs.major}"`);
  const host=document.createElement('div');host.style.cssText='position:absolute;visibility:hidden;left:0;top:0;pointer-events:none';sheet.append(host);
  await docx.renderAsync(st.viewBytes,host,q('weStyles'),{className:P,inWrapper:false,breakPages:false,ignoreLastRenderedPageBreak:true,experimental:true,useBase64URL:false,
    renderHeaders:true,renderFooters:true,renderFootnotes:true,renderEndnotes:true,renderChanges:false,renderComments:false,trimXmlDeclaration:true,ignoreWidth:false,ignoreHeight:false,ignoreFonts:false});
  // paragraph styles also reach text typed straight into the paragraph (docx-preview styles only its spans)
  for(const el of q('weStyles').querySelectorAll('style'))el.textContent=el.textContent.replace(/(^|\n|\})\s*(p\.[\w-]+) span(\s*\{)/g,'$1$2, $2 span$3');
  let hdr=null,ftr=null;const notes=q('weNotes');
  for(const s of [...host.children]){
    if(s.tagName!=='SECTION'){continue;}
    for(const c of [...s.children]){
      if(c.tagName==='HEADER'){if(!hdr)hdr=c;}else if(c.tagName==='FOOTER'){if(!ftr)ftr=c;}
      else if(c.tagName==='ARTICLE'){while(c.firstChild)body.append(c.firstChild);}
      else notes.append(c);
    }
  }
  host.remove();
  if(hdr)q('weHdr').append(hdr);if(ftr)q('weFtr').append(ftr);
  for(const a of body.querySelectorAll('a[href^="#__"]')){const m=/^#__([ap])(\d+)$/.exec(a.getAttribute('href'));if(!m)continue;
    a.removeAttribute('href');a.contentEditable='false';a.dataset.atom=m[2];a.classList.add(m[1]==='p'?'we-pb':'we-atom');
    if(m[1]==='p')a.textContent='';a.querySelectorAll('img').forEach(i=>i.draggable=false);}
  for(const a of body.querySelectorAll('a[href^="#__c"]')){const m=/^#__(cs|ce)(\d+)$/.exec(a.getAttribute('href'));if(!m)continue;a.replaceWith(cmMarker(m[1],m[2]));}
  for(const s of body.querySelectorAll(`span.${P}_commentreference,span.${P}-commentreference`))if(!s.textContent.replace(ZW,''))s.remove();
  for(const box of [q('weHdr'),q('weFtr')])for(const a of box.querySelectorAll('a[href^="#__f"]')){const f=a.getAttribute('href').slice(4);const sp=document.createElement('span');sp.className='we-fld';sp.dataset.f=f;sp.contentEditable='false';
    const inner=a.querySelector('span');if(inner&&inner.getAttribute('style'))sp.setAttribute('style',inner.getAttribute('style'));sp.textContent=f==='PAGE'?'1':'#';a.replaceWith(sp);}
  for(const box of [q('weHdr'),q('weFtr')])for(const a of box.querySelectorAll('a[href^="#__"]'))a.replaceWith(...a.childNodes);
  E.hfOrig={header:hfHtml('header'),footer:hfHtml('footer')};E.hfClean={...E.hfOrig};
  E.cm=E.st.cm.list;E.cmDirty=false;for(const c of E.cm.values())if(!cmLive(c))c.keep=true;
  E.cmOrigLive=[...E.cm.values()].filter(c=>cmLive(c)).map(c=>c.id).sort().join(',');E.cmClean=cmSig();
  for(const p of body.querySelectorAll('p'))if(!p.childNodes.length)p.append(document.createElement('br'));
  for(const t of body.querySelectorAll(`span.${P}-tab-stop`))t.contentEditable='false';
  E.map=new Map();E.mapMiss=0;mapContainer(body,st.body,'');
  await nextFrame();
  for(const [k,ent] of E.map){const el=body.querySelector(`[data-src="${k}"]`);if(!el)continue;ent.html=el.outerHTML;if(ent.kind==='p')ent.sig=sigOf(el);}
  ensureBlocks();
  // the blank document starts as a paragraph the editor owns
  E.cleanHtml=body.innerHTML;
}
function injectCss(){if(document.getElementById('weCss'))return;const s=document.createElement('style');s.id='weCss';
  s.textContent=CSS+'\n::highlight(we-find){background:#FFE58F}\n::highlight(we-find-cur){background:#FFB300;color:#000}';document.head.append(s);}
async function open(opts){
  if(E)await forceClose();
  if(!window.JSZip||!window.docx)throw new Error('ספריות העריכה לא נטענו');
  injectCss();
  const root=document.createElement('div');root.id='wed';root.innerHTML=shellHtml();document.body.append(root);
  const name=(opts.name||'מסמך חדש.docx').replace(/\.(docx?|dotx)$/i,'');
  E={opts,root,body:null,hist:[],hi:-1,zoom:1,dirty:false,blobs:new Set(),map:new Map(),name};
  E.body=q('weBody');E.main=E.body;q('weName').value=name;
  const load=document.createElement('div');load.className='we-load';load.textContent='פותח את המסמך לעריכה…';q('weScroll').append(load);
  try{E.st=await prepare(opts.bytes||null);await mount();}
  catch(e){root.remove();E=null;throw e;}
  load.remove();
  fillFonts();buildGallery();wire();
  if(E.st.heb===false){E.body.lang='en';}
  snap();updateUI();
  if(q('weScroll').clientWidth<px(E.page.w)+48)fitZoom();
  await nextFrame();guides();words();cmRender();
  E.body.focus({preventScroll:true});caretIn(E.body.querySelector('p,h1,h2,h3,li,td')||E.body);updateUI();
  if(opts.isNew)q('weName').select();
  if(opts.restored){E.dirty=true;E.cleanHtml='';markDirty();}
  else if(window.indexedDB){const d=await DRAFT.get(draftKey());if(d&&d.bytes&&E){
    const when=new Date(d.time).toLocaleString('he-IL',{dateStyle:'short',timeStyle:'short'});
    const a=await ask('נמצאה עבודה שלא נשמרה',`המסמך ${esc(d.name||'')} נערך כאן ב-${when} ולא נשמר. לפתוח את הגרסה הזו?`,[{v:'x',l:'לא, לפתוח את הקובץ'},{v:'ok',l:'שחזור',cls:'pri'}]);
    if(a==='ok'){const o={...opts,bytes:d.bytes,restored:true,isNew:false};await forceClose();return open(o);}
    await DRAFT.del(draftKey());}}
  startDrafts();
  return true;
}
const fileName=()=>((q('weName').value||'').trim().replace(/[\\/:*?"<>|]+/g,'').replace(/\.docx$/i,'')||'מסמך')+'.docx';
async function build(msg){E.opts.busy(true,msg||'שומר את המסמך…');try{await nextFrame();return await buildDocx();}finally{E.opts.busy(false);}}
async function doSave(){
  if(!E||E.saving)return;E.saving=true;
  try{const{bytes,lost}=await build();const ok=await E.opts.save(fileName(),bytes,'docx');
    if(ok){E.cleanHtml=E.main.innerHTML;E.hfClean={header:hfHtml('header'),footer:hfHtml('footer')};E.cmClean=cmSig();E.pageDirty=false;E.savedOnce=true;markDirty();DRAFT.del(draftKey());E.opts.toast('נשמר: '+fileName()+(lost?`. ${lost} תמונות לא נשמרו`:''));}}
  catch(e){if(!(e&&e.name==='AbortError')){console.error(e);E.opts.toast('השמירה נכשלה: '+(e.message||e));}}
  finally{if(E)E.saving=false;}
}
async function doPdf(){
  if(!E||E.saving)return;E.saving=true;
  try{const{bytes}=await build('מכין PDF…');E.opts.busy(true,'מכין PDF…');const pdf=await E.opts.toPdf(bytes);E.opts.busy(false);
    const n=fileName().replace(/\.docx$/,'.pdf');if(await E.opts.save(n,pdf,'pdf'))E.opts.toast('נשמר: '+n);}
  catch(e){E.opts.busy(false);if(!(e&&e.name==='AbortError')){console.error(e);E.opts.toast('יצירת ה-PDF נכשלה: '+(e.message||e));}}
  finally{if(E)E.saving=false;}
}
async function doDesk(){
  if(!E||E.saving)return;E.saving=true;
  try{const{bytes}=await build('מכין את המסמך לחתימה…');const f=new File([bytes],fileName(),{type:DOCX_MIME});const o=E.opts;DRAFT.del(draftKey());await forceClose();await o.toDesk(f);}
  catch(e){console.error(e);if(E){E.saving=false;E.opts.toast('לא הצלחתי לפתוח לחתימה: '+(e.message||e));}}
}
async function doClose(){
  if(!E)return;if(E.snapT){snap();markDirty();}
  if(E.dirty){const a=await ask('לשמור את השינויים?',`יש שינויים ב-${esc(fileName())} שעוד לא נשמרו.`,[{v:'x',l:'ביטול'},{v:'drop',l:'סגירה בלי לשמור',cls:'warn'},{v:'save',l:'שמירה',cls:'pri'}]);
    if(!a)return;if(a==='save'){await doSave();if(E&&E.dirty)return;}}
  DRAFT.del(draftKey());
  const o=E.opts;await forceClose();o.onClose&&o.onClose();
}
async function forceClose(){
  if(!E)return;const e=E;clearInterval(e.draftT);
  document.removeEventListener('selectionchange',e.onSel);window.removeEventListener('keydown',e.onKey,true);window.removeEventListener('paste',e.onPaste,true);
  window.removeEventListener('copy',e.onCopy,true);window.removeEventListener('cut',e.onCopy,true);window.removeEventListener('dragover',e.onDrag,true);window.removeEventListener('drop',e.onDrop,true);
  window.removeEventListener('beforeunload',e.onUnload);document.removeEventListener('pointerdown',e.onDown,true);window.removeEventListener('resize',e.onResize);
  clearTimeout(e.snapT);if(window.CSS&&window.CSS.highlights){window.CSS.highlights.delete('we-find');window.CSS.highlights.delete('we-find-cur');}
  e.root.remove();for(const u of e.blobs)URL.revokeObjectURL(u);E=null;
}
/* ---------- the assistant: describes the document and makes the changes it is asked for ---------- */
// The assistant gets the document as numbered blocks (#0, #1…) and answers with a list of actions.
// All the actions of one request are one step: "ביטול" (or Ctrl+Z) takes the whole request back.
const AI_STY={title:'title',subtitle:'subtitle',h1:'heading 1',h2:'heading 2',h3:'heading 3',quote:'quote',heading1:'heading 1',heading2:'heading 2',heading3:'heading 3'};
const aiHex=c=>{if(c===null||c===false||c==='none')return null;c=String(c||'').trim().replace(/^#/,'');if(/^[0-9a-f]{3}$/i.test(c))c=c.replace(/./g,x=>x+x);return /^[0-9a-f]{6}$/i.test(c)?c.toUpperCase():undefined;};
function aiTopBlocks(){
  const out=[];
  const walk=el=>{for(const c of el.children){
    if(c.tagName==='TABLE'){out.push(c);continue;}
    if(c.tagName==='HR'){out.push(c);continue;}
    if(PARA.test(c.tagName)&&!c.querySelector('p,h1,h2,h3,h4,h5,h6,li,table')){out.push(c);continue;}
    if(c.querySelector&&c.querySelector('p,h1,h2,h3,h4,h5,h6,li,table'))walk(c);
  }};
  walk(E.main);return out;
}
function aiStyleName(b){const c=styleCls(b);if(!c)return /^H(\d)$/.test(b.tagName)?'כותרת '+b.tagName[1]:'רגיל';
  const s=E.st.styles.find(x=>classKey(x.id)===c);if(!s)return 'רגיל';const bi=BUILTIN.find(x=>x.name===s.name);return bi&&bi.label?bi.label:(s.label||s.name||'רגיל');}
const AL_HE={right:'ימין',left:'שמאל',center:'מרכז',justify:'מיושר'};
function aiDescribe(b){
  if(b.tagName==='HR')return '[קו מפריד]';
  if(b.tagName==='TABLE'){const rows=[...b.rows].filter(r=>r.closest('table')===b);const cols=Math.max(0,...rows.map(r=>r.cells.length));
    const first=rows[0]&&rows[0].cells[0];const bg=first?hex(getComputedStyle(first).backgroundColor):null;
    const lines=rows.slice(0,12).map((r,i)=>`  שורה ${i+1}: `+[...r.cells].map(c=>c.textContent.replace(ZW,'').replace(/\s+/g,' ').trim().slice(0,40)).join(' | '));
    return `[טבלה ${rows.length}×${cols}${bg?', רקע כותרת #'+bg:''}]\n${lines.join('\n')}${rows.length>12?`\n  …ועוד ${rows.length-12} שורות`:''}`;}
  const cs=getComputedStyle(b);const sp=b.querySelector('span')||b;const rs=getComputedStyle(sp);
  const n=numOf(b);const list=n?(n.id===E.st.num.bullet?'תבליט':'מספור'):'';
  const meta=[aiStyleName(b),AL_HE[visAlign(cs)]||'',firstFamily(rs.fontFamily)+' '+Math.round(parseFloat(rs.fontSize)*.75*2)/2,parseInt(rs.fontWeight,10)>=600?'מודגש':'',hex(rs.color)&&hex(rs.color)!=='000000'?'#'+hex(rs.color):'',list].filter(Boolean).join(' | ');
  const t=b.textContent.replace(ZW,'').replace(/\s+/g,' ').trim();
  return `[${meta}] ${t?t.slice(0,400)+(t.length>400?'…':''):'(שורה ריקה)'}`;
}
function aiContext(){
  if(!E)return null;if(E.hf)exitHF(true);if(E.snapT)snap();
  const bl=aiTopBlocks();E.aiRefs=bl;
  const r=E.lastRange&&E.main.contains(E.lastRange.commonAncestorContainer)?E.lastRange:null;
  let sel='';if(r){const ids=bl.map((b,i)=>r.intersectsNode(b)?i:-1).filter(i=>i>=0);
    const t=r.toString().replace(ZW,'').trim();
    sel=ids.length?`הסמן נמצא ב-#${ids[0]}${ids.length>1?` עד #${ids[ids.length-1]}`:''}${t?`, טקסט מסומן: "${t.slice(0,300)}"`:''}`:'';}
  const lines=[];let total=0,cut=0;
  for(let i=0;i<bl.length;i++){const l=`#${i} ${aiDescribe(bl[i])}`;if(total+l.length>60000){cut=bl.length-i;break;}lines.push(l);total+=l.length;}
  const pg=E.page;const land=pg.w>pg.h;
  const hdr=hfHtml('header')?hfBox('header').textContent.replace(ZW,'').trim():'',ftr=hfHtml('footer')?hfBox('footer').textContent.replace(ZW,'').trim():'';
  return {kind:'docx',name:E.opts.name||'',text:
`סוג הקובץ: מסמך Word. כיוון ברירת המחדל: ${E.st.heb?'מימין לשמאל (עברית)':'משמאל לימין'}. גופן ברירת מחדל: ${E.st.def.font} ${E.st.def.size}.
עמוד: ${land?'לרוחב':'לאורך'}, שוליים ${(pg.top/567).toFixed(1)} ס"מ. כותרת עליונה: ${hdr?'"'+hdr.slice(0,120)+'"':'אין'}. כותרת תחתונה: ${ftr?'"'+ftr.slice(0,120)+'"':'אין'}.
${sel||'אין סמן במסמך.'}
סגנונות זמינים: ${galItems().map(g=>g.label).join(', ')}.
בלוקים במסמך (${bl.length}):
${lines.join('\n')}${cut?`\n…ועוד ${cut} בלוקים שלא נכללו`:''}`};
}
// which blocks an action points at
function aiTargets(a){
  const bl=E.aiRefs&&E.aiRefs.length?E.aiRefs:aiTopBlocks();
  const live=b=>b&&b.isConnected&&E.main.contains(b);
  let t=a.refs!==undefined?a.refs:a.ref!==undefined?a.ref:a.target;
  if(t===undefined||t===null)t='selection';
  const pick=x=>{
    if(typeof x==='number'||/^#?\d+$/.test(String(x))){const b=bl[+String(x).replace('#','')];return live(b)?[b]:[];}
    const m=/^#?(\d+)\s*[-–]\s*#?(\d+)$/.exec(String(x));if(m){const o=[];for(let i=+m[1];i<=+m[2];i++)if(live(bl[i]))o.push(bl[i]);return o;}
    switch(String(x)){
      case 'all':return aiTopBlocks();
      case 'headings':return aiTopBlocks().filter(b=>b.tagName!=='TABLE'&&b.tagName!=='HR'&&isHeading(b));
      case 'body':case 'paragraphs':return aiTopBlocks().filter(b=>b.tagName!=='TABLE'&&b.tagName!=='HR'&&!isHeading(b));
      case 'tables':return [...E.main.querySelectorAll('table')].filter(t=>!t.parentElement.closest('table'));
      case 'first':{const b=aiTopBlocks().find(x=>x.tagName!=='HR'&&x.textContent.replace(ZW,'').trim());return b?[b]:[];}
      case 'last':{const b=aiTopBlocks().reverse().find(x=>x.tagName!=='HR'&&x.textContent.replace(ZW,'').trim());return b?[b]:[];}
      case 'selection':default:{const r=E.lastRange;if(!r||!E.main.contains(r.commonAncestorContainer))return[];
        return aiTopBlocks().filter(b=>r.intersectsNode(b));}
    }};
  const out=[];for(const x of Array.isArray(t)?t:[t])for(const b of pick(x))if(!out.includes(b))out.push(b);
  return out;
}
// the paragraphs inside a block (a table: the paragraphs of its cells)
const aiParas=b=>b.tagName==='TABLE'?[...b.querySelectorAll('p,h1,h2,h3,h4,h5,h6,li')].filter(p=>!p.querySelector('p,table')):b.tagName==='HR'?[]:[b];
// spans holding the given words (or all the text), each on its own
function aiSpans(b,find){
  flatten(b,b);
  const out=[];const tw=document.createTreeWalker(b,NodeFilter.SHOW_TEXT);const texts=[];let t;while((t=tw.nextNode()))if(t.length&&!tokOf(t))texts.push(t);
  if(!find){for(const x of texts){const s=isolate(x);if(s)out.push(s);}return out;}
  const f=String(find);
  for(let x of texts){let i;while(x&&(i=x.data.indexOf(f))>=0){if(i>0)x=x.splitText(i);const rest=x.length>f.length?x.splitText(f.length):null;const s=isolate(x);if(s)out.push(s);x=rest;}}
  return out;
}
function aiRunStyle(s,st){
  if(st.bold!==undefined)s.style.fontWeight=st.bold?'700':'400';
  if(st.italic!==undefined)s.style.fontStyle=st.italic?'italic':'normal';
  if(st.underline!==undefined)setDeco(s,'underline',!!st.underline);
  if(st.strike!==undefined)setDeco(s,'line-through',!!st.strike);
  if(st.font)s.style.fontFamily=cssFont(st.font);
  if(+st.size>0)setPt(s,Math.max(1,Math.min(400,Math.round(+st.size*2)/2)));
  const c=aiHex(st.color);if(c!==undefined)s.style.color=c?'#'+c:'#000000';
  if(st.highlight!==undefined){const h=aiHex(st.highlight);if(h!==undefined)s.style.backgroundColor=h?'#'+h:'';}
}
const RUN_KEYS=['bold','italic','underline','strike','font','size','color','highlight'];
function aiBlockStyle(b,st){
  if(st.align){const v={right:'right',left:'left',center:'center',justify:'justify',start:'start',end:'end',ימין:'right',שמאל:'left',מרכז:'center'}[st.align];if(v)b.style.textAlign=v;}
  if(st.dir==='rtl'||st.dir==='ltr')b.style.direction=st.dir;
  if(+st.lineHeight>0)b.style.lineHeight=String(Math.min(4,+st.lineHeight));
  if(st.spaceBefore!==undefined&&+st.spaceBefore>=0)b.style.marginTop=(+st.spaceBefore)+'pt';
  if(st.spaceAfter!==undefined&&+st.spaceAfter>=0)b.style.marginBottom=(+st.spaceAfter)+'pt';
  if(st.indent!==undefined&&+st.indent>=0)b.style.marginInlineStart=(+st.indent*37.7953).toFixed(1)+'px';
}
function aiFormat(b,st,find){
  if(!st||typeof st!=='object')return;
  if(b.tagName==='TABLE'){
    const sh=aiHex(st.fill!==undefined?st.fill:st.shading);
    if(sh!==undefined)for(const td of b.querySelectorAll('td,th'))if(td.closest('table')===b)td.style.backgroundColor=sh?'#'+sh:'';
  }
  for(const p of aiParas(b)){
    if(!find)aiBlockStyle(p,st);
    if(RUN_KEYS.some(k=>st[k]!==undefined)){
      if(!p.textContent.replace(ZW,'')){aiRunStyle(p,st);continue;}
      for(const s of aiSpans(p,find))aiRunStyle(s,st);mergeSpans(p);
    }
  }
}
function aiStyleOf(name){
  name=String(name||'').toLowerCase().trim();const key=AI_STY[name.replace(/\s+/g,'')]||name;
  if(/^(normal|רגיל|p|paragraph)$/.test(key))return galItems().find(g=>g.normal)||{normal:true};
  const s=E.st.styles.find(x=>x.type==='paragraph'&&x.name===key)||E.st.styles.find(x=>x.type==='paragraph'&&(x.label===name||x.label.toLowerCase()===name));
  if(s)return{id:s.id,cls:classKey(s.id),normal:false};
  const g=galItems().find(x=>x.label===name);return g||null;
}
function aiSetStyle(b,name){
  const it=aiStyleOf(name);if(!it||b.tagName==='TABLE'||b.tagName==='HR')return b;
  if(b.tagName!=='P'){const p=document.createElement('p');for(const a of b.attributes)p.setAttribute(a.name,a.value);while(b.firstChild)p.append(b.firstChild);b.replaceWith(p);
    const i=E.aiRefs?E.aiRefs.indexOf(b):-1;if(i>=0)E.aiRefs[i]=p;b=p;}
  for(const c of [...b.classList])if(c.startsWith(P+'_'))b.classList.remove(c);
  if(!it.normal&&it.cls)b.classList.add(it.cls);
  for(const s of b.querySelectorAll('span,font,b,strong,i,em')){for(const pr of ['font-size','font-family','color','font-weight','font-style'])s.style.removeProperty(pr);delete s.dataset.fsb;}
  for(const t of b.querySelectorAll('b,strong,i,em,font'))t.replaceWith(...t.childNodes);
  for(const pr of ['font-size','font-family','color','font-weight','font-style'])b.style.removeProperty(pr);
  return b;
}
function aiList(b,kind){
  for(const p of aiParas(b)){const n=numOf(p);if(n)p.classList.remove(n.cls);const lp=lpCls();
    if(kind==='none'||!kind){if(lp)p.classList.remove(lp);p.style.removeProperty('counter-set');delete p.dataset.ns;continue;}
    const id=/^(number|numbers|decimal|מספור|ממוספר)$/.test(kind)?E.st.num.decimal:E.st.num.bullet;
    p.classList.add(`${P}-num-${id}-${n?n.lvl:0}`);if(lp&&!styleCls(p))p.classList.add(lp);}
}
const aiDir=t=>/[\u0590-\u05FF\u0600-\u06FF]/.test(t||'')?'rtl':/[A-Za-z]/.test(t||'')?'ltr':(E.st.heb?'rtl':'ltr');
function aiPara(text,dir){const p=document.createElement('p');p.style.direction=dir||aiDir(text);
  const lines=String(text??'').split('\n');lines.forEach((l,i)=>{if(i)p.append(document.createElement('br'));if(l)p.append(document.createTextNode(l));});
  if(!p.textContent)p.append(document.createElement('br'));return p;}
function aiTable(bk){
  let rows=Array.isArray(bk.rows)?bk.rows.filter(Array.isArray):[];
  if(!rows.length){const r=Math.max(1,Math.min(60,+bk.rowsCount||+bk.r||3)),c=Math.max(1,Math.min(20,+bk.cols||+bk.c||3));rows=Array.from({length:r},()=>Array(c).fill(''));}
  rows=rows.slice(0,200);const cols=Math.min(20,Math.max(...rows.map(r=>r.length)));
  const dir=aiDir(rows.flat().join(' '));const w=Math.floor(contentW()),cw=Math.floor(w/cols);
  const t=document.createElement('table');t.style.cssText=`width:${w}px;border-collapse:collapse;direction:${dir}`;
  const tb=document.createElement('tbody');t.append(tb);
  const head=bk.header!==false;const hc=aiHex(bk.headerColor)||'1F3864';const band=bk.banded?aiHex(bk.bandColor)||'EEF2F8':null;
  rows.forEach((r,ri)=>{const tr=document.createElement('tr');for(let j=0;j<cols;j++){const td=document.createElement('td');
    td.style.cssText=`border:0.5pt solid #000;width:${cw}px;padding:0 5.4pt;vertical-align:top`;
    const p=aiPara(r[j]==null?'':String(r[j]),dir);
    if(head&&ri===0&&p.textContent){td.style.backgroundColor='#'+hc;const s=makeSpan({...runProps(E.main,E.main),b:true,color:aiHex(bk.headerText)||'FFFFFF',bg:null},p.textContent);p.textContent='';p.append(s);p.style.textAlign='center';}
    else if(head&&ri===0)td.style.backgroundColor='#'+hc;
    else if(band&&ri%2===0)td.style.backgroundColor='#'+band;
    td.append(p);tr.append(td);}tb.append(tr);});
  return t;
}
function aiBuild(bk){
  if(typeof bk==='string')bk={type:'p',text:bk};if(!bk||typeof bk!=='object')return null;
  const type=String(bk.type||bk.kind||'p').toLowerCase();
  if(type==='table')return aiTable(bk);
  if(type==='hr'||type==='line')return document.createElement('hr');
  if(type==='pagebreak'||type==='page-break'){const p=document.createElement('p');p.style.direction=E.st.heb?'rtl':'ltr';const a=document.createElement('a');a.className='we-pb';a.contentEditable='false';p.append(a);return p;}
  let p=aiPara(bk.text,bk.dir);
  const sty=bk.style&&typeof bk.style==='string'?bk.style:AI_STY[type]?type:null;
  if(sty)p=aiSetStyle(p,sty);
  if(/^(bullet|bullets|number|numbers|li)$/.test(type))aiList(p,type==='li'?'bullet':type);
  const st=bk.style&&typeof bk.style==='object'?bk.style:bk.format&&typeof bk.format==='object'?bk.format:null;
  if(st)aiFormat(p,st);
  return p;
}
function aiAnchor(a){
  const where=String(a.at||a.position||a.pos||'').toLowerCase();
  const bl=E.aiRefs||[];const ref=a.ref!==undefined?bl[+String(a.ref).replace('#','')]:null;
  if(ref&&ref.isConnected)return {node:ref,before:/before|לפני|above/.test(where)};
  if(/start|top|begin|התחלה/.test(where)){const f=E.main.firstElementChild;return {node:f,before:true};}
  if(/cursor|caret|here|selection/.test(where)){const r=E.lastRange;let b=r&&E.main.contains(r.startContainer)?blockOf(r.startContainer):null;while(b&&b.parentNode!==E.main&&b.parentNode)b=b.parentNode.closest('table')&&b.parentNode.closest('table').parentNode===E.main?b.parentNode.closest('table'):b.parentNode;if(b&&b.parentNode===E.main)return {node:b,before:false};}
  return {node:E.main.lastElementChild,before:false};
}
function aiInsert(a){
  const blocks=(Array.isArray(a.blocks)?a.blocks:a.block?[a.block]:a.text!==undefined?[{type:a.type||'p',text:a.text,style:a.style}]:[]).map(aiBuild).filter(Boolean);
  if(!blocks.length)return 0;
  const {node,before}=aiAnchor(a);
  // a lone empty paragraph is replaced, not pushed down
  const emptyDoc=E.main.children.length===1&&node&&node.tagName==='P'&&!node.textContent.replace(ZW,'').trim()&&!node.querySelector('img,[data-atom],.we-pb');
  if(!node)E.main.append(...blocks);
  else if(emptyDoc){node.replaceWith(...blocks);}
  else if(before)node.before(...blocks);else node.after(...blocks);
  const last=blocks[blocks.length-1];if(last.tagName==='TABLE'&&!last.nextElementSibling)last.after(newPara(getComputedStyle(last).direction));
  return blocks.length;
}
function aiSetText(b,text){
  if(b.tagName==='TABLE'||b.tagName==='HR')return;
  flatten(b,b);const sp=b.querySelector('span');const p=sp?runProps(sp,b):null;
  const keep=[...b.childNodes].filter(n=>n.nodeType===1&&n.dataset&&n.dataset.atom!==undefined);
  b.textContent='';
  const lines=String(text??'').split('\n');
  lines.forEach((l,i)=>{if(i)b.append(document.createElement('br'));if(!l)return;
    if(p&&sp&&sp.getAttribute('style'))b.append(makeSpan(p,l));else b.append(document.createTextNode(l));});
  for(const k of keep)b.append(k);
  if(!b.textContent&&!b.querySelector('img,[data-atom]'))b.append(document.createElement('br'));
  b.style.direction=aiDir(text);
}
function aiTableStyle(t,a){
  const rows=[...t.rows].filter(r=>r.closest('table')===t);if(!rows.length)return;
  const kind=String(a.style||a.preset||'header').toLowerCase();
  const hc=aiHex(a.headerColor)||'1F3864',bc=aiHex(a.bandColor)||'EEF2F8',ht=aiHex(a.headerText)||'FFFFFF';
  const border=aiHex(a.borderColor)||'000000';
  rows.forEach((r,ri)=>{for(const td of r.cells){
    if(kind!=='plain'){td.style.border=kind==='none'?'none':`0.5pt solid #${border}`;}
    if(ri===0&&kind!=='none'&&kind!=='plain'&&kind!=='grid'){td.style.backgroundColor='#'+hc;
      for(const p of aiParas(td)){for(const s of aiSpans(p))aiRunStyle(s,{bold:true,color:ht});mergeSpans(p);p.style.textAlign='center';}}
    else if(/band|striped|zebra|פסים/.test(kind)&&ri%2===0)td.style.backgroundColor='#'+bc;
    else if(ri>0&&a.clearFill!==false&&/band|striped|zebra|header|פסים/.test(kind))td.style.backgroundColor='';
  }});
  if(a.width==='full'||a.fit){t.style.width=Math.floor(contentW())+'px';}
}
function aiHF(k,a){
  const el=hfEl(k);const p=newPara(aiDir(a.text));p.innerHTML='';
  if(a.pageNumber){p.innerHTML=(a.text?esc(a.text)+' ':'')+(a.pageNumber==='xofy'?`עמוד ${fld('PAGE')} מתוך ${fld('NUMPAGES')}`:fld('PAGE'));}
  else p.textContent=a.text||'';
  if(!p.childNodes.length)p.append(document.createElement('br'));
  p.style.textAlign={right:'right',left:'left',center:'center'}[a.align]||'center';
  if(a.replace!==false)el.innerHTML='';el.append(p);
}
const MARG={normal:{top:1440,bottom:1440,left:1440,right:1440},narrow:{top:720,bottom:720,left:720,right:720},wide:{top:1440,bottom:1440,left:2880,right:2880},moderate:{top:1440,bottom:1440,left:1080,right:1080}};
function aiApply(actions){
  if(!E)throw new Error('העורך סגור');if(E.hf)exitHF(true);if(E.snapT)snap();
  const mark={hi:E.hi,page:{...E.page},pageDirty:E.pageDirty,hf:{header:hfBox('header').innerHTML,footer:hfBox('footer').innerHTML}};
  const done=[],skipped=[];
  for(const a of (Array.isArray(actions)?actions:[]).slice(0,80)){
    if(!a||typeof a!=='object')continue;const op=String(a.op||a.action||'').toLowerCase();
    try{
      if(op==='insert'||op==='add'){const n=aiInsert(a);if(n)done.push(op);else skipped.push(op);}
      else if(op==='format'||op==='style'&&typeof a.style==='object'){const ts=aiTargets(a);ts.forEach(b=>aiFormat(b,a.style||a.format,a.find||a.text));ts.length?done.push(op):skipped.push(op);}
      else if(op==='style'||op==='setstyle'||op==='heading'){const ts=aiTargets(a);ts.forEach(b=>aiSetStyle(b,a.name||a.style||a.level&&'h'+a.level));ts.length?done.push(op):skipped.push(op);}
      else if(op==='list'){const ts=aiTargets(a);ts.forEach(b=>aiList(b,a.kind||a.type||'bullet'));ts.length?done.push(op):skipped.push(op);}
      else if(op==='text'||op==='settext'||op==='rewrite'){const ts=aiTargets(a);if(ts[0]){aiSetText(ts[0],a.text);done.push(op);}else skipped.push(op);}
      else if(op==='delete'||op==='remove'){const ts=aiTargets(a);ts.forEach(b=>b.remove());if(!E.main.children.length)E.main.append(newPara());ts.length?done.push(op):skipped.push(op);}
      else if(op==='replace'){const f=String(a.find||'');if(!f){skipped.push(op);continue;}let n=0;
        const tw=document.createTreeWalker(E.main,NodeFilter.SHOW_TEXT);const ts=[];let t;while((t=tw.nextNode()))if(t.data.includes(f)&&!tokOf(t))ts.push(t);
        for(const x of ts){const parts=x.data.split(f);n+=parts.length-1;x.data=parts.join(String(a.with??a.replace??''));if(a.all===false&&n)break;}
        n?done.push(op):skipped.push(op);}
      else if(op==='table'||op==='tablestyle'){const ts=aiTargets({...a,refs:a.refs??a.ref??'tables'}).filter(b=>b.tagName==='TABLE');ts.forEach(t=>aiTableStyle(t,a));ts.length?done.push(op):skipped.push(op);}
      else if(op==='page'){const ch={};
        if(a.orientation){const land=/land|רוחב/.test(a.orientation);const w=Math.min(E.page.w,E.page.h),h=Math.max(E.page.w,E.page.h);Object.assign(ch,land?{w:h,h:w}:{w,h});}
        if(a.margins&&MARG[a.margins])Object.assign(ch,MARG[a.margins]);
        else if(+a.marginCm>0){const v=Math.round(+a.marginCm*567);Object.assign(ch,{top:v,bottom:v,left:v,right:v});}
        if(Object.keys(ch).length){Object.assign(E.page,ch);E.pageDirty=true;applyPage();done.push(op);}else skipped.push(op);}
      else if(op==='header'||op==='footer'){aiHF(op,a);done.push(op);}
      else if(op==='pagenumbers'||op==='pagenumber'){aiHF(a.where==='header'?'header':'footer',{text:a.text||'',pageNumber:a.format==='page'?'page':'xofy',align:a.align||'center',replace:a.replace});done.push(op);}
      else if(op==='toc'){const first=E.main.firstElementChild;const blk=tocBlock(tocItems(),E.st.heb?'rtl':'ltr');const g=tocGroup();
        if(g){g[0].before(blk);g.forEach(x=>x.remove());}else{const at=a.ref!==undefined&&E.aiRefs&&E.aiRefs[+a.ref]?E.aiRefs[+a.ref]:null;if(at)at.after(blk);else E.main.insertBefore(blk,first);}done.push(op);}
      else skipped.push(op||'?');
    }catch(e){console.warn('assistant action',a,e);skipped.push(op);}
  }
  E.aiRefs=null;
  snap();after();buildGallery();refreshCounters();
  E.aiMark=mark;
  return {done:done.length,skipped:skipped.length,mark};
}
function aiRevert(mark){
  if(!E||!mark)return false;if(E.hf)exitHF(true);if(E.snapT)snap();
  while(E.hi>mark.hi&&E.hi>0){E.hi--;}restoreSnap(E.hist[E.hi]);
  E.page=mark.page;E.pageDirty=mark.pageDirty;applyPage();
  hfBox('header').innerHTML=mark.hf.header;hfBox('footer').innerHTML=mark.hf.footer;
  after();return true;
}
async function shareFile(fmt){
  if(!E)throw new Error('אין מסמך פתוח');
  const{bytes}=await buildDocx();
  if(fmt!=='pdf')return new File([bytes],fileName(),{type:DOCX_MIME});
  try{const pdf=await E.opts.toPdf(bytes);return new File([pdf],fileName().replace(/\.docx$/i,'')+'.pdf',{type:'application/pdf'});}finally{E&&E.opts.busy(false);}
}
window.WORD_EDITOR={open,close:forceClose,isOpen:()=>!!E,share:{file:shareFile,name:()=>E?fileName():''},
  assist:{context:aiContext,apply:aiApply,revert:aiRevert,ribbon:()=>E&&E.root.querySelector('.we-ribbon'),root:()=>E&&E.root},
  _test:{get E(){return E;},enterHF,exitHF,pageNumber,removePageNumbers,cmNew,cmAction,cmRemove,tocInsert,tocRemove,wordCount,autoCorrect,countText,buildDocx:()=>buildDocx(),inline,fmt,setFont,setSize,setColor,setHilite,align,setDir,toggleList,indent,applyStyle,galItems,insertTable,tableOp,insertPageBreak,lineSpacing,setPage,undo,redo,getSel,setSel,snap,clearFmt,insertText,findAll,replaceAll:()=>replaceAll(),runFind}};
})();
