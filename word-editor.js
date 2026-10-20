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
  for(const p of view.getElementsByTagNameNS(W,'p'))if(paraBidi(p,sb))swapJc(kid(p,'pPr'));
  for(const r of view.getElementsByTagNameNS(W,'r'))if(HEB.test(textOf(r)))csRun(kid(r,'rPr'));
  const vs=sdoc.cloneNode(true);const vsb=styleBidiMap(vs);
  for(const s of kids(vs.documentElement,'style')){
    if(wAttr(s,'type')==='paragraph'&&vsb.of(wAttr(s,'styleId')))swapJc(kid(s,'pPr'));
    if(heb)csRun(kid(s,'rPr'));
  }
  if(heb){const d=vs.getElementsByTagNameNS(W,'rPrDefault')[0];if(d)csRun(kid(d,'rPr'));}
  const vfiles=new Map(files);put(vfiles,docPart,ser(view));put(vfiles,'word/styles.xml',ser(vs));
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
  return{files,docPart,docStr,norm,body,sdoc,styles:readStyles(sdoc),atoms,viewBytes,page,sect,num,heb,def,themeCs,sb};
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
  caret:'<path d="M6 8l4 4 4-4"/>'
};
const ic=n=>`<svg viewBox="0 0 20 20" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">${ICO[n]||''}</svg>`;
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
 ${G('סמלים',BL('symbolMenu','symbol','סמל','הוספת סמל','data-pop="symbol"')+BL('date','date','תאריך','הוספת התאריך של היום'))}
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
  <button type="button" role="tab" data-tab="view" aria-selected="false">תצוגה</button>
  <button type="button" role="tab" data-tab="table" aria-selected="false" class="ctx" hidden>טבלה</button>
</div>
<div class="we-ribbon" id="weRibbon">${ribbonHtml()}</div>
<div class="we-scroll" id="weScroll"><div class="we-canvas" id="weCanvas">
  <div class="we-sheet" id="weSheet">
    <div class="we-hdr" id="weHdr" aria-hidden="true"></div>
    <div class="we-body ${P}" id="weBody" contenteditable="true" spellcheck="true" lang="he" role="textbox" aria-multiline="true" aria-label="תוכן המסמך"></div>
    <div class="we-notes ${P}" id="weNotes" aria-label="הערות שוליים"></div>
    <div class="we-ftr" id="weFtr" aria-hidden="true"></div>
    <div class="we-guides" id="weGuides" aria-hidden="true"></div>
    <div class="we-imgbox" id="weImgBox" hidden><i data-h="nw"></i><i data-h="ne"></i><i data-h="sw"></i><i data-h="se"></i></div>
  </div>
</div></div>
<div class="we-status"><span id="wePages"></span><span id="weWords"></span><span class="grow"></span><span class="we-hint" id="weHint"></span>
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
.we-hdr,.we-ftr{position:absolute;opacity:.5;pointer-events:none;overflow:hidden}
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
.we-pop{position:fixed;z-index:30;background:#fff;border:1px solid var(--wl);border-radius:10px;box-shadow:0 14px 34px -12px rgba(22,33,58,.4);padding:8px;max-width:340px}
.we-pop[hidden]{display:none}
.we-pop h4{margin:2px 2px 6px;font-size:12px;color:var(--wmu);font-weight:600}
.we-pop .pal{display:grid;grid-template-columns:repeat(10,22px);gap:4px}
.we-pop .pal button{width:22px;height:22px;border-radius:4px;border:1px solid rgba(0,0,0,.15);cursor:pointer;padding:0}
.we-pop .pal button:hover{transform:scale(1.15)}
.we-pop .mi{display:flex;align-items:center;gap:8px;width:100%;border:0;background:none;padding:7px 10px;border-radius:7px;cursor:pointer;text-align:start;white-space:nowrap}
.we-pop .mi:hover,.we-pop .mi.on{background:#EEF2F8}
.we-pop .mi small{color:var(--wmu);margin-inline-start:auto;font-size:12px}
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
}`;

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
function restoreSnap(s){E.body.innerHTML=s.html;hideImgBox();setSel(s.sel);after();}
function mutate(fn){if(E.snapT)snap();const r=fn();snap();after();return r;}
function after(){markDirty();refreshCounters();schedule();updateUI();}
function markDirty(){const d=E.body.innerHTML!==E.cleanHtml||E.pageDirty;if(d!==E.dirty){E.dirty=d;}
  const s=q('weState');s.textContent=E.dirty?'שינויים שלא נשמרו':(E.savedOnce?'נשמר':'');s.classList.toggle('dirty',E.dirty);}
let schT=0;function schedule(){clearTimeout(schT);schT=setTimeout(()=>{if(!E)return;guides();words();posImgBox();},250);}
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
  if(!(window.CSS&&CSS.highlights&&window.Highlight))return;
  const F=E.find;CSS.highlights.delete('we-find');CSS.highlights.delete('we-find-cur');
  if(!F||!F.list.length)return;CSS.highlights.set('we-find',new Highlight(...F.list));if(F.list[F.at])CSS.highlights.set('we-find-cur',new Highlight(F.list[F.at]));
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
function closeFind(){q('weFind').hidden=true;E.find=null;paintFind();if(window.CSS&&CSS.highlights){CSS.highlights.delete('we-find');CSS.highlights.delete('we-find-cur');}
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
  const g=q('weGuides'),b=E.body,pg=E.page;g.innerHTML='';
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
  const top=b.offsetTop;
  marks.forEach((y,i)=>{const m=document.createElement('i');m.style.top=(top+y)+'px';const l=document.createElement('b');l.textContent='עמוד '+(i+2);m.append(l);g.append(m);});
  pageInfo();
}
function pageInfo(){
  if(!E)return;const r=curRange();let cur=1;
  if(r&&E.pageMarks){const n=r.startContainer.nodeType===3?r.startContainer.parentElement:r.startContainer;if(n&&n!==E.body){const y=offIn(n,E.body).y;cur=1+E.pageMarks.filter(m=>m<=y+1).length;}}
  q('wePages').textContent=`עמוד ${cur} מתוך ${E.pageCount||1}`;
}
function words(){const t=E.body.textContent.replace(ZW,'');const n=(t.match(/[\p{L}\p{N}][\p{L}\p{N}\p{M}'"״׳.-]*/gu)||[]).length;q('weWords').textContent=`${n.toLocaleString('he-IL')} מילים`;}
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
  if(!r){tabT.hidden=true;return;}
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
  pageInfo();
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
  spell:()=>{const on=E.body.spellcheck=!E.body.spellcheck;E.root.querySelectorAll('[data-cmd=spell]').forEach(b=>b.setAttribute('aria-pressed',on));E.body.blur();E.body.focus();},
  find:()=>openFind(false),replace:()=>openFind(true),
  image:()=>q('weImgFile').click(),link:linkDialog,pagebreak:insertPageBreak,hr:insertHr,
  date:()=>insertText(new Date().toLocaleDateString('he-IL',{day:'numeric',month:'numeric',year:'numeric'})),
  rowAbove:()=>tableOp('rowAbove'),rowBelow:()=>tableOp('rowBelow'),colRight:()=>tableOp('colRight'),colLeft:()=>tableOp('colLeft'),
  delRow:()=>tableOp('delRow'),delCol:()=>tableOp('delCol'),delTable:()=>tableOp('delTable'),borders:()=>tableOp('borders'),noborders:()=>tableOp('noborders'),
  shade:()=>shade(E.shade===undefined?'D9E2F3':E.shade),
  zin:()=>setZoom(E.zoom+.1),zout:()=>setZoom(E.zoom-.1),z100:()=>setZoom(1),zfit:fitZoom,
  print:printDoc
};
function updatePainter(){E.root.querySelectorAll('[data-cmd=painter]').forEach(b=>b.setAttribute('aria-pressed',!!E.painter));}
function printDoc(){
  const pg=E.page;let st=document.getElementById('wePrintPage');if(!st){st=document.createElement('style');st.id='wePrintPage';document.head.append(st);}
  const cm=t=>(t/567).toFixed(2)+'cm';st.textContent=`@page{size:${cm(pg.w)} ${cm(pg.h)};margin:${cm(pg.top)} ${cm(pg.right)} ${cm(pg.bottom)} ${cm(pg.left)}}`;
  hideImgBox();const sel=getSelection();if(sel)sel.removeAllRanges();E.body.blur();setTimeout(()=>print(),50);
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
    closePop();const f=CMD[b.dataset.cmd];if(f){if(!/^(z|print|guides|marks|spell|find|replace)/.test(b.dataset.cmd))liveRange();f();}
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
  const body=E.body;
  body.addEventListener('input',e=>{
    if(e.inputType==='insertParagraph'){const b=curBlock();const pv=b&&b.previousElementSibling;
      if(b&&pv&&isHeading(pv)&&!b.textContent.replace(ZW,'').trim()){for(const c of [...b.classList])if(c.startsWith(P+'_'))b.classList.remove(c);if(/^H\d$/.test(b.tagName)){const p=document.createElement('p');p.innerHTML='<br>';b.replaceWith(p);caretTo(p,0);}else{b.innerHTML='<br>';caretTo(b,0);}}}
    if(!body.firstElementChild||[...body.childNodes].some(n=>n.nodeType===3&&n.data.trim()))ensureBlocks();
    snapLater();markDirtySoon();refreshCounters();schedule();updateUI();
  });
  body.addEventListener('beforeinput',e=>{
    const t=e.inputType;
    if(t==='historyUndo'){e.preventDefault();undo();}else if(t==='historyRedo'){e.preventDefault();redo();}
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
  E.onSel=()=>{if(!E)return;const r=curRange();if(r){E.lastRange=r.cloneRange();E.lastSel=null;}updateUI();};
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
let mdT=0;function markDirtySoon(){clearTimeout(mdT);mdT=setTimeout(()=>{if(E)markDirty();},300);}
function onKey(e){
  if(!E)return;
  const t=e.target,inBody=E.body.contains(t)||t===E.body,inRoot=E.root.contains(t);
  if(!inRoot)e.stopImmediatePropagation();
  if(E.dlgOpen){if(!inRoot)e.preventDefault();return;}
  const mod=(e.ctrlKey||e.metaKey)&&!e.altKey,c=e.code;
  const run=(f,...a)=>{e.preventDefault();f(...a);};
  if(e.key==='Escape'){if(!q('wePop').hidden)return run(closePop);if(!q('weFind').hidden)return run(closeFind);if(E.selImg)return run(hideImgBox);if(E.painter){E.painter=null;updatePainter();return;}return;}
  if(mod){
    if(c==='KeyS')return run(doSave);
    if(c==='KeyP')return run(printDoc);
    if(c==='KeyF')return run(openFind,false);
    if(c==='KeyH')return run(openFind,true);
    if(!inBody&&!(t===document.body))return;
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
  const runs=runsOf([...el.childNodes].filter(n=>!(n.nodeType===1&&/^(UL|OL|TABLE)$/.test(n.tagName))),el,ctx);
  return`<w:p>${pPr.childNodes.length?ctx.clean(ser(pPr)):''}${runs}</w:p>`;
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
  if(E.snapT)snap();
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
  for(const el of E.body.querySelectorAll('[data-src]'))ctx.last.set(el.dataset.src,el);
  for(const img of E.body.querySelectorAll('img'))if(!img.closest('[data-atom]'))ctx.imgs.set(img,await imgBytes(img));
  let bodyXml=exportContainer(E.body,ctx);
  if(!bodyXml)bodyXml=`<w:p><w:pPr>${st.heb?'<w:bidi/>':''}</w:pPr></w:p>`;
  let sect=st.sect?st.sect.cloneNode(true):parseXml(`<w:sectPr xmlns:w="${W}"><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>`).documentElement;
  if(E.pageDirty)patchSect(sect);
  put(files,st.docPart,head+'<w:body>'+bodyXml+clean(ser(sect))+'</w:body></w:document>');
  if(ctx.extra.length){const n=txt(files,'word/numbering.xml');put(files,'word/numbering.xml',n.replace(/<\/w:numbering>\s*$/,ctx.extra.join('')+'</w:numbering>'));}
  const z=new JSZip();for(const [p,b] of files)z.file(p,b);
  const bytes=await z.generateAsync({type:'uint8array',compression:'DEFLATE',compressionOptions:{level:6}});
  return{bytes,lost:ctx.lost};
}

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
  E.body=q('weBody');q('weName').value=name;
  const load=document.createElement('div');load.className='we-load';load.textContent='פותח את המסמך לעריכה…';q('weScroll').append(load);
  try{E.st=await prepare(opts.bytes||null);await mount();}
  catch(e){root.remove();E=null;throw e;}
  load.remove();
  fillFonts();buildGallery();wire();
  if(E.st.heb===false){E.body.lang='en';}
  snap();updateUI();
  if(q('weScroll').clientWidth<px(E.page.w)+48)fitZoom();
  await nextFrame();guides();words();
  E.body.focus({preventScroll:true});caretIn(E.body.querySelector('p,h1,h2,h3,li,td')||E.body);updateUI();
  if(opts.isNew)q('weName').select();
  return true;
}
const fileName=()=>((q('weName').value||'').trim().replace(/[\\/:*?"<>|]+/g,'').replace(/\.docx$/i,'')||'מסמך')+'.docx';
async function build(msg){E.opts.busy(true,msg||'שומר את המסמך…');try{await nextFrame();return await buildDocx();}finally{E.opts.busy(false);}}
async function doSave(){
  if(!E||E.saving)return;E.saving=true;
  try{const{bytes,lost}=await build();const ok=await E.opts.save(fileName(),bytes,'docx');
    if(ok){E.cleanHtml=E.body.innerHTML;E.pageDirty=false;E.savedOnce=true;markDirty();E.opts.toast('נשמר: '+fileName()+(lost?`. ${lost} תמונות לא נשמרו`:''));}}
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
  try{const{bytes}=await build('מכין את המסמך לחתימה…');const f=new File([bytes],fileName(),{type:DOCX_MIME});const o=E.opts;await forceClose();await o.toDesk(f);}
  catch(e){console.error(e);if(E){E.saving=false;E.opts.toast('לא הצלחתי לפתוח לחתימה: '+(e.message||e));}}
}
async function doClose(){
  if(!E)return;if(E.snapT){snap();markDirty();}
  if(E.dirty){const a=await ask('לשמור את השינויים?',`יש שינויים ב-${esc(fileName())} שעוד לא נשמרו.`,[{v:'x',l:'ביטול'},{v:'drop',l:'סגירה בלי לשמור',cls:'warn'},{v:'save',l:'שמירה',cls:'pri'}]);
    if(!a)return;if(a==='save'){await doSave();if(E&&E.dirty)return;}}
  const o=E.opts;await forceClose();o.onClose&&o.onClose();
}
async function forceClose(){
  if(!E)return;const e=E;
  document.removeEventListener('selectionchange',e.onSel);window.removeEventListener('keydown',e.onKey,true);window.removeEventListener('paste',e.onPaste,true);
  window.removeEventListener('copy',e.onCopy,true);window.removeEventListener('cut',e.onCopy,true);window.removeEventListener('dragover',e.onDrag,true);window.removeEventListener('drop',e.onDrop,true);
  window.removeEventListener('beforeunload',e.onUnload);document.removeEventListener('pointerdown',e.onDown,true);window.removeEventListener('resize',e.onResize);
  clearTimeout(e.snapT);if(window.CSS&&CSS.highlights){CSS.highlights.delete('we-find');CSS.highlights.delete('we-find-cur');}
  e.root.remove();for(const u of e.blobs)URL.revokeObjectURL(u);E=null;
}
window.WORD_EDITOR={open,close:forceClose,isOpen:()=>!!E,
  _test:{get E(){return E;},buildDocx:()=>buildDocx(),inline,fmt,setFont,setSize,setColor,setHilite,align,setDir,toggleList,indent,applyStyle,galItems,insertTable,tableOp,insertPageBreak,lineSpacing,setPage,undo,redo,getSel,setSel,snap,clearFmt,insertText,findAll,replaceAll:()=>replaceAll(),runFind}};
})();
