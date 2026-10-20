/* סימן — עורך Excel.
   קובץ Excel נפתח לעריכה כמו ב-Excel: ערכים ונוסחאות, גופנים, צבעים, גבולות, יישור, מיזוג תאים,
   עיצוב מספרים, הוספה ומחיקה של שורות ועמודות, וגיליונות.
   השמירה כותבת את אותו קובץ XLSX: מה שלא שונה נשאר כמו במקור (עיצובים, תמונות, תרשימים, הגדרות),
   ורק התאים שנערכו נכתבים מחדש. Needs: JSZip, SheetJS (XLSX.SSF) for number formats. */
(function(){
'use strict';
const NS={s:'http://schemas.openxmlformats.org/spreadsheetml/2006/main',r:'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel:'http://schemas.openxmlformats.org/package/2006/relationships',a:'http://schemas.openxmlformats.org/drawingml/2006/main'};
const S=NS.s;
const OX='application/vnd.openxmlformats-officedocument.';
const XLSX_MIME=OX+'spreadsheetml.sheet';
const RT_DOC='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const RT_SHEET='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
const RT_SST='http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
const RT_STY='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const XML_HEAD='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
const enc=new TextEncoder(),dec=new TextDecoder('utf-8');
const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'})[c]);
const parseXml=s=>{const d=new DOMParser().parseFromString(s,'application/xml');if(d.getElementsByTagName('parsererror').length)throw new Error('XML פגום בקובץ');return d;};
const ser=n=>{const s=new XMLSerializer().serializeToString(n);return n.nodeType===9&&!/^<\?xml/.test(s)?XML_HEAD+s:s;};
const kids=(el,...n)=>el?[...el.childNodes].filter(x=>x.nodeType===1&&x.namespaceURI===S&&(!n.length||n.includes(x.localName))):[];
const kid=(el,n)=>kids(el,n)[0]||null;
const sNew=(doc,name,attrs)=>{const e=doc.createElementNS(S,name);if(attrs)for(const k in attrs)if(attrs[k]!==undefined&&attrs[k]!==null)e.setAttribute(k,String(attrs[k]));return e;};
const HEB=/[\u0590-\u05FF\u0600-\u06FF]/;
const txt=(f,p)=>f.has(p)?dec.decode(f.get(p)):null;
const put=(f,p,s)=>f.set(p,typeof s==='string'?enc.encode(s):s);
const nextFrame=()=>new Promise(r=>requestAnimationFrame(()=>setTimeout(r,0)));
/* A1 references */
const colName=c=>{let s='';c++;while(c>0){const m=(c-1)%26;s=String.fromCharCode(65+m)+s;c=Math.floor((c-1)/26);}return s;};
const colIdx=s=>{let n=0;for(const ch of s.toUpperCase())n=n*26+(ch.charCodeAt(0)-64);return n-1;};
const refOf=(r,c)=>colName(c)+(r+1);
const parseRef=s=>{const m=/^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(String(s).trim());return m?{r:+m[2]-1,c:colIdx(m[1])}:null;};
const parseRange=s=>{const [a,b]=String(s).split(':');const A=parseRef(a);if(!A)return null;const B=b?parseRef(b):A;if(!B)return null;
  return{r1:Math.min(A.r,B.r),c1:Math.min(A.c,B.c),r2:Math.max(A.r,B.r),c2:Math.max(A.c,B.c)};};
const rangeStr=g=>g.r1===g.r2&&g.c1===g.c2?refOf(g.r1,g.c1):refOf(g.r1,g.c1)+':'+refOf(g.r2,g.c2);
const key=(r,c)=>r*16384+c;

/* ------------------------------------------------------------ package */
async function readZip(bytes){const z=await JSZip.loadAsync(bytes);const f=new Map();for(const [p,x] of Object.entries(z.files))if(!x.dir)f.set(p,await x.async('uint8array'));return f;}
function relsPathOf(part){const i=part.lastIndexOf('/');return part.slice(0,i+1)+'_rels/'+part.slice(i+1)+'.rels';}
function resolveTarget(part,t){if(/^\//.test(t))return t.slice(1);const b=part.slice(0,part.lastIndexOf('/')+1).split('/').filter(Boolean);for(const s of t.split('/')){if(s==='..')b.pop();else if(s&&s!=='.')b.push(s);}return b.join('/');}
function rels(files,part){const s=txt(files,relsPathOf(part));if(!s)return[];return[...parseXml(s).getElementsByTagNameNS(NS.rel,'Relationship')].map(e=>({id:e.getAttribute('Id'),type:e.getAttribute('Type'),target:e.getAttribute('Target'),mode:e.getAttribute('TargetMode')}));}
function addRel(files,part,type,target){
  const rp=relsPathOf(part);let s=txt(files,rp)||(XML_HEAD+`<Relationships xmlns="${NS.rel}"></Relationships>`);
  const used=new Set([...s.matchAll(/Id="([^"]+)"/g)].map(m=>m[1]));let n=1,id;do id='rIdSm'+(n++);while(used.has(id));
  s=s.replace('</Relationships>',`<Relationship Id="${id}" Type="${type}" Target="${esc(target)}"/></Relationships>`);put(files,rp,s);return id;
}
function ctOverride(files,part,mime){let s=txt(files,'[Content_Types].xml');if(!s||s.includes(`PartName="/${part}"`))return;s=s.replace('</Types>',`<Override PartName="/${part}" ContentType="${mime}"/></Types>`);put(files,'[Content_Types].xml',s);}
function ctRemove(files,part){let s=txt(files,'[Content_Types].xml');if(!s)return;s=s.replace(new RegExp(`<Override PartName="/${part.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}"[^>]*/>`),'');put(files,'[Content_Types].xml',s);}
function blankXlsx(){
  const f=new Map();
  put(f,'[Content_Types].xml',XML_HEAD+`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="${OX}spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="${OX}spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="${OX}spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="${OX}spreadsheetml.sharedStrings+xml"/><Override PartName="/docProps/app.xml" ContentType="${OX}extended-properties+xml"/></Types>`);
  put(f,'_rels/.rels',XML_HEAD+`<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="${RT_DOC}" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`);
  put(f,'docProps/app.xml',XML_HEAD+'<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Siman</Application></Properties>');
  put(f,'xl/workbook.xml',XML_HEAD+`<workbook xmlns="${S}" xmlns:r="${NS.r}"><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="16000" windowHeight="9000" rightToLeft="1"/></bookViews><sheets><sheet name="גיליון1" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029"/></workbook>`);
  put(f,'xl/_rels/workbook.xml.rels',XML_HEAD+`<Relationships xmlns="${NS.rel}"><Relationship Id="rId1" Type="${RT_SHEET}" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${RT_STY}" Target="styles.xml"/><Relationship Id="rId3" Type="${RT_SST}" Target="sharedStrings.xml"/></Relationships>`);
  put(f,'xl/worksheets/sheet1.xml',XML_HEAD+`<worksheet xmlns="${S}" xmlns:r="${NS.r}"><sheetViews><sheetView rightToLeft="1" tabSelected="1" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`);
  put(f,'xl/sharedStrings.xml',XML_HEAD+`<sst xmlns="${S}" count="0" uniqueCount="0"/>`);
  put(f,'xl/styles.xml',XML_HEAD+`<styleSheet xmlns="${S}"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="Arial"/><family val="2"/><scheme val="minor"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`);
  return f;
}

/* ------------------------------------------------------------ colors */
const IDX=['000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF','000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF','800000','008000','000080','808000','800080','008080','C0C0C0','808080','9999FF','993366','FFFFCC','CCFFFF','660066','FF8080','0066CC','CCCCFF','000080','FF00FF','FFFF00','00FFFF','800080','800000','008080','0000FF','00CCFF','CCFFFF','CCFFCC','FFFF99','99CCFF','FF99CC','CC99FF','FFCC99','3366FF','33CCCC','99CC00','FFCC00','FF9900','FF6600','666699','969696','003366','339966','003300','333300','993300','993366','333399','333333'];
function themeColors(files){
  const p=[...files.keys()].find(k=>/^xl\/theme\/theme\d*\.xml$/.test(k));const out=['FFFFFF','000000','E7E6E6','44546A','4472C4','ED7D31','A5A5A5','FFC000','5B9BD5','70AD47','0563C1','954F72'];
  if(!p)return out;const d=parseXml(txt(files,p));const cs=d.getElementsByTagNameNS(NS.a,'clrScheme')[0];if(!cs)return out;
  const get=n=>{const e=cs.getElementsByTagNameNS(NS.a,n)[0];if(!e)return null;const c=e.firstElementChild;if(!c)return null;return(c.getAttribute('lastClr')||c.getAttribute('val')||'').toUpperCase()||null;};
  ['lt1','dk1','lt2','dk2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'].forEach((n,i)=>{const v=get(n);if(v&&/^[0-9A-F]{6}$/.test(v))out[i]=v;});
  return out;
}
function tintHex(h,t){if(!t)return h;const v=[0,2,4].map(i=>parseInt(h.slice(i,i+2),16)).map(c=>Math.round(t<0?c*(1+t):c+(255-c)*t));return v.map(x=>Math.max(0,Math.min(255,x)).toString(16).padStart(2,'0')).join('').toUpperCase();}
function colorOf(el,theme,def){
  if(!el)return def;if(el.getAttribute('auto')==='1')return def;
  const rgb=el.getAttribute('rgb');let h=null;
  if(rgb)h=rgb.slice(-6).toUpperCase();
  else if(el.hasAttribute('theme'))h=theme[+el.getAttribute('theme')]||def;
  else if(el.hasAttribute('indexed')){const i=+el.getAttribute('indexed');h=i===64?def:i===65?'FFFFFF':IDX[i]||def;}
  if(!h)return def;return tintHex(h,+el.getAttribute('tint')||0);
}

/* ------------------------------------------------------------ styles */
const BUILTIN_FMT={0:'General',1:'0',2:'0.00',3:'#,##0',4:'#,##0.00',9:'0%',10:'0.00%',11:'0.00E+00',12:'# ?/?',13:'# ??/??',14:'dd/mm/yyyy',15:'d-mmm-yy',16:'d-mmm',17:'mmm-yy',18:'h:mm AM/PM',19:'h:mm:ss AM/PM',20:'h:mm',21:'h:mm:ss',22:'dd/mm/yyyy h:mm',37:'#,##0 ;(#,##0)',38:'#,##0 ;[Red](#,##0)',39:'#,##0.00;(#,##0.00)',40:'#,##0.00;[Red](#,##0.00)',45:'mm:ss',46:'[h]:mm:ss',47:'mmss.0',48:'##0.0E+0',49:'@'};
function readStyles(files,path,theme){
  let s=path&&txt(files,path);
  if(!s){path='xl/styles.xml';s=txt(blankXlsx(),'xl/styles.xml');put(files,path,s);}
  const doc=parseXml(s),root=doc.documentElement;
  const ensure=(n,before)=>{let e=kid(root,n);if(!e){e=sNew(doc,n,{count:0});const b=before.map(x=>kid(root,x)).find(Boolean);root.insertBefore(e,b||null);}return e;};
  const fontsEl=ensure('fonts',['fills','borders','cellStyleXfs','cellXfs']);const fillsEl=ensure('fills',['borders','cellStyleXfs','cellXfs']);
  const bordersEl=ensure('borders',['cellStyleXfs','cellXfs']);const xfsEl=ensure('cellXfs',['cellStyles','dxfs','tableStyles','colors','extLst']);
  if(!kids(fontsEl,'font').length)fontsEl.append(parseXml(`<font xmlns="${S}"><sz val="11"/><name val="Arial"/></font>`).documentElement);
  if(!kids(fillsEl,'fill').length){fillsEl.append(parseXml(`<fill xmlns="${S}"><patternFill patternType="none"/></fill>`).documentElement);fillsEl.append(parseXml(`<fill xmlns="${S}"><patternFill patternType="gray125"/></fill>`).documentElement);}
  if(!kids(bordersEl,'border').length)bordersEl.append(parseXml(`<border xmlns="${S}"><left/><right/><top/><bottom/><diagonal/></border>`).documentElement);
  if(!kids(xfsEl,'xf').length)xfsEl.append(sNew(doc,'xf',{numFmtId:0,fontId:0,fillId:0,borderId:0,xfId:0}));
  const numFmts={};for(const n of kids(kid(root,'numFmts'),'numFmt'))numFmts[n.getAttribute('numFmtId')]=n.getAttribute('formatCode');
  return{path,doc,root,numFmts,fontsEl,fillsEl,bordersEl,xfsEl,theme,cache:new Map(),desc:new Map()};
}
const onv=e=>!!e&&!/^(0|false)$/.test(e.getAttribute('val')||'');
function fontDesc(st,f){
  const v=(n,a='val')=>{const e=kid(f,n);return e?e.getAttribute(a):null;};
  const u=kid(f,'u');
  return{b:onv(kid(f,'b')),i:onv(kid(f,'i')),strike:onv(kid(f,'strike')),u:u?(u.getAttribute('val')||'single'):null,
    sz:+v('sz')||11,color:colorOf(kid(f,'color'),st.theme,'000000'),name:v('name')||'Arial',va:v('vertAlign')};
}
function descOf(st,i){
  i=+i||0;if(st.desc.has(i))return st.desc.get(i);
  const xfs=kids(st.xfsEl,'xf');const xf=xfs[i]||xfs[0];
  const font=kids(st.fontsEl,'font')[+xf.getAttribute('fontId')||0];const fill=kids(st.fillsEl,'fill')[+xf.getAttribute('fillId')||0];
  const border=kids(st.bordersEl,'border')[+xf.getAttribute('borderId')||0];
  let bg=null;const pf=fill&&kid(fill,'patternFill');
  if(pf){const pt=pf.getAttribute('patternType');if(pt&&pt!=='none'&&pt!=='gray125')bg=colorOf(kid(pf,'fgColor'),st.theme,pt==='solid'?null:'D9D9D9')||colorOf(kid(pf,'bgColor'),st.theme,null);}
  else if(fill&&kid(fill,'gradientFill')){const s=fill.getElementsByTagNameNS(S,'color')[0];bg=colorOf(s,st.theme,null);}
  const bd={};for(const side of ['left','right','top','bottom']){const e=kid(border,side)||kid(border,side==='left'?'start':side==='right'?'end':side);
    const sty=e&&e.getAttribute('style');if(sty&&sty!=='none')bd[side]={style:sty,color:colorOf(kid(e,'color'),st.theme,'000000')};}
  const al=kid(xf,'alignment');const nid=+xf.getAttribute('numFmtId')||0;
  const d={font:font?fontDesc(st,font):{sz:11,name:'Arial',color:'000000'},bg,bd,
    h:al&&al.getAttribute('horizontal')||null,v:al&&al.getAttribute('vertical')||null,wrap:!!al&&/^(1|true)$/.test(al.getAttribute('wrapText')||''),
    indent:al?+al.getAttribute('indent')||0:0,rot:al?+al.getAttribute('textRotation')||0:0,
    nid,fmt:st.numFmts[nid]??BUILTIN_FMT[nid]??'General'};
  st.desc.set(i,d);return d;
}
const FONT_ORDER=['b','i','strike','condense','extend','outline','shadow','u','vertAlign','sz','color','name','family','charset','scheme'];
function newXf(st,base,ch){
  base=+base||0;const k=base+'|'+JSON.stringify(ch);if(st.cache.has(k))return st.cache.get(k);
  const d=st.doc;const xfs=kids(st.xfsEl,'xf');const xf=(xfs[base]||xfs[0]).cloneNode(true);const cur=descOf(st,base);
  if(ch.font){
    const old=kids(st.fontsEl,'font')[+xf.getAttribute('fontId')||0];const f=sNew(d,'font');const m={...cur.font,...ch.font};
    const add=(n,a)=>f.append(sNew(d,n,a));
    if(m.b)add('b');if(m.i)add('i');if(m.strike)add('strike');if(m.u)add('u',m.u==='single'?null:{val:m.u});if(m.va)add('vertAlign',{val:m.va});
    add('sz',{val:m.sz});add('color',{rgb:'FF'+m.color});add('name',{val:m.name});
    if(old){for(const n of ['family','charset'])if(kid(old,n))f.append(kid(old,n).cloneNode(true));if(kid(old,'scheme')&&m.name===cur.font.name)f.append(kid(old,'scheme').cloneNode(true));}
    st.fontsEl.append(f);xf.setAttribute('fontId',kids(st.fontsEl,'font').length-1);xf.setAttribute('applyFont','1');
  }
  if(ch.bg!==undefined){
    if(ch.bg===null)xf.setAttribute('fillId','0');
    else{const f=parseXml(`<fill xmlns="${S}"><patternFill patternType="solid"><fgColor rgb="FF${ch.bg}"/><bgColor indexed="64"/></patternFill></fill>`).documentElement;
      st.fillsEl.append(d.importNode(f,true));xf.setAttribute('fillId',kids(st.fillsEl,'fill').length-1);}
    xf.setAttribute('applyFill','1');
  }
  if(ch.bd){
    const m={...cur.bd};for(const s in ch.bd){if(ch.bd[s])m[s]=ch.bd[s];else delete m[s];}
    const b=sNew(d,'border');for(const s of ['left','right','top','bottom','diagonal']){const e=sNew(d,s);if(m[s]){e.setAttribute('style',m[s].style);e.append(sNew(d,'color',{rgb:'FF'+(m[s].color||'000000')}));}b.append(e);}
    st.bordersEl.append(b);xf.setAttribute('borderId',kids(st.bordersEl,'border').length-1);xf.setAttribute('applyBorder','1');
  }
  if(ch.al){
    let al=kid(xf,'alignment');if(!al){al=sNew(d,'alignment');xf.insertBefore(al,xf.firstChild);}
    for(const [a,v] of Object.entries(ch.al)){const n={h:'horizontal',v:'vertical',wrap:'wrapText'}[a]||a;if(v===null||v===false||v===undefined)al.removeAttribute(n);else al.setAttribute(n,v===true?'1':String(v));}
    if(!al.attributes.length)xf.removeChild(al);xf.setAttribute('applyAlignment','1');
  }
  if(ch.fmt!==undefined){xf.setAttribute('numFmtId',fmtId(st,ch.fmt));xf.setAttribute('applyNumberFormat','1');}
  st.xfsEl.append(xf);const id=kids(st.xfsEl,'xf').length-1;st.cache.set(k,id);return id;
}
function fmtId(st,code){
  for(const [id,c] of Object.entries(BUILTIN_FMT))if(c===code&&+id!==14&&+id!==22)return +id;
  if(code==='dd/mm/yyyy')return 14;
  for(const [id,c] of Object.entries(st.numFmts))if(c===code)return +id;
  let el=kid(st.root,'numFmts');if(!el){el=sNew(st.doc,'numFmts',{count:0});st.root.insertBefore(el,st.root.firstElementChild);}
  const id=Math.max(163,...Object.keys(st.numFmts).map(Number))+1;el.append(sNew(st.doc,'numFmt',{numFmtId:id,formatCode:code}));st.numFmts[id]=code;return id;
}

/* ------------------------------------------------------------ workbook and sheets */
function readSst(files,path){
  const out=[];if(!path||!files.has(path))return{path:path||'xl/sharedStrings.xml',list:out,doc:null};
  const d=parseXml(txt(files,path));
  for(const si of kids(d.documentElement,'si')){let t='';for(const x of si.getElementsByTagNameNS(S,'t'))if(!(x.parentNode&&x.parentNode.localName==='rPh'))t+=x.textContent;out.push(t);}
  return{path,list:out,doc:d};
}
function readSheet(files,path,sst){
  const doc=parseXml(txt(files,path)),root=doc.documentElement;
  const sv=kid(kid(root,'sheetViews'),'sheetView');
  const fp=kid(root,'sheetFormatPr');
  const defH=fp&&fp.getAttribute('defaultRowHeight')?+fp.getAttribute('defaultRowHeight'):15;
  const defW=fp&&fp.getAttribute('defaultColWidth')?+fp.getAttribute('defaultColWidth'):(fp&&fp.getAttribute('baseColWidth')?+fp.getAttribute('baseColWidth')+.71:8.43);
  const cols=new Map();
  for(const c of kids(kid(root,'cols'),'col')){const a=+c.getAttribute('min')-1,b=+c.getAttribute('max')-1;
    for(let i=a;i<=Math.min(b,a+16383)&&i<=16383;i++){cols.set(i,{w:c.getAttribute('width')?+c.getAttribute('width'):null,hidden:c.getAttribute('hidden')==='1',style:c.getAttribute('style'),custom:c.getAttribute('customWidth')});if(b>a+300&&i>a+300)break;}}
  const rows=new Map(),cells=new Map();const shared={};
  let rr=-1;
  for(const row of kids(kid(root,'sheetData'),'row')){
    rr=row.getAttribute('r')?+row.getAttribute('r')-1:rr+1;
    const info={attrs:{}};for(const a of row.attributes)if(!['r','spans'].includes(a.name))info.attrs[a.name]=a.value;
    info.ht=row.getAttribute('ht')?+row.getAttribute('ht'):null;info.hidden=row.getAttribute('hidden')==='1';rows.set(rr,info);
    let cc=-1;
    for(const c of kids(row,'c')){
      const ref=c.getAttribute('r');const p=ref?parseRef(ref):null;cc=p?p.c:cc+1;const r=p?p.r:rr;
      const t=c.getAttribute('t')||'n';const vEl=kid(c,'v');const vt=vEl?vEl.textContent:null;const fEl=kid(c,'f');
      const cell={s:+c.getAttribute('s')||0,el:c,f:null,v:null,t:'n'};
      if(t==='s'){cell.sidx=+vt;cell.v=sst.list[+vt]??'';cell.t='s';}
      else if(t==='inlineStr'){let x='';for(const e of c.getElementsByTagNameNS(S,'t'))x+=e.textContent;cell.v=x;cell.t='s';}
      else if(t==='str'){cell.v=vt??'';cell.t='s';}
      else if(t==='b'){cell.v=vt==='1';cell.t='b';}
      else if(t==='e'){cell.v=vt||'#N/A';cell.t='e';}
      else if(t==='d'){cell.v=vt;cell.t='s';}
      else{cell.v=vt===null||vt===''?null:+vt;cell.t='n';}
      if(fEl){const ft=fEl.getAttribute('t');
        if(ft==='shared'){const si=fEl.getAttribute('si');if(fEl.textContent){shared[si]={f:fEl.textContent,r,c:cc};cell.f=fEl.textContent;}else cell.shared=si;}
        else if(ft!=='dataTable'){cell.f=fEl.textContent;if(ft==='array')cell.array=fEl.getAttribute('ref')||refOf(r,cc);}}
      cells.set(key(r,cc),Object.assign(cell,{r,c:cc}));
    }
  }
  for(const cell of cells.values())if(cell.shared!==undefined){const m=shared[cell.shared];if(m)cell.f=moveFormula(m.f,cell.r-m.r,cell.c-m.c);delete cell.shared;cell.wasShared=true;}
  const merges=kids(kid(root,'mergeCells'),'mergeCell').map(m=>parseRange(m.getAttribute('ref'))).filter(Boolean);
  let maxR=0,maxC=0;for(const c of cells.values()){if(c.r>maxR)maxR=c.r;if(c.c>maxC)maxC=c.c;}
  for(const m of merges){maxR=Math.max(maxR,m.r2);maxC=Math.max(maxC,m.c2);}
  const pane=kid(sv,'pane');
  const pm=kid(root,'pageMargins'),ps=kid(root,'pageSetup'),po=kid(root,'printOptions'),spr=kid(kid(root,'sheetPr'),'pageSetUpPr');
  const at=(e,a,d)=>e&&e.getAttribute(a)!==null&&e.getAttribute(a)!==''?e.getAttribute(a):d;
  const page={l:+at(pm,'left',.7),r:+at(pm,'right',.7),t:+at(pm,'top',.75),b:+at(pm,'bottom',.75),h:+at(pm,'header',.3),f:+at(pm,'footer',.3),
    orient:at(ps,'orientation','portrait'),paper:+at(ps,'paperSize',9),scale:+at(ps,'scale',100),fit:at(spr,'fitToPage','0')==='1'||at(spr,'fitToPage','0')==='true',
    fw:+at(ps,'fitToWidth',1),fh:+at(ps,'fitToHeight',1),grid:/^(1|true)$/.test(at(po,'gridLines','')),hc:/^(1|true)$/.test(at(po,'horizontalCentered',''))};
  return{path,doc,root,page,headings:!sv||sv.getAttribute('showRowColHeaders')!=='0',formulas:!!sv&&sv.getAttribute('showFormulas')==='1',
    rtl:!!sv&&sv.getAttribute('rightToLeft')==='1',grid:!sv||sv.getAttribute('showGridLines')!=='0',
    defH,defW,cols,rows,cells,merges,maxR,maxC,
    freeze:pane&&pane.getAttribute('state')==='frozen'?{r:+pane.getAttribute('ySplit')||0,c:+pane.getAttribute('xSplit')||0}:null};
}
async function prepare(bytes){
  const files=bytes?await readZip(bytes):blankXlsx();
  let wbPart='xl/workbook.xml';for(const r of rels(files,''))if(r.type===RT_DOC)wbPart=resolveTarget('',r.target);
  if(!files.has(wbPart))throw new Error('לא נמצא תוכן Excel בקובץ');
  const wbDoc=parseXml(txt(files,wbPart));const wr=rels(files,wbPart);
  const theme=themeColors(files);
  const stR=wr.find(r=>r.type===RT_STY),ssR=wr.find(r=>r.type===RT_SST);
  const st=readStyles(files,stR&&resolveTarget(wbPart,stR.target),theme);
  if(!stR){addRel(files,wbPart,RT_STY,'styles.xml');ctOverride(files,'xl/styles.xml',OX+'spreadsheetml.styles+xml');}
  const sst=readSst(files,ssR&&resolveTarget(wbPart,ssR.target));sst.hadPart=!!ssR;
  const wp=kid(wbDoc.documentElement,'workbookPr');const d1904=!!wp&&/^(1|true)$/.test(wp.getAttribute('date1904')||'');
  const sheets=[];let wi=-1;
  const dns=[...wbDoc.getElementsByTagNameNS(S,'definedName')];
  for(const s of kids(kid(wbDoc.documentElement,'sheets'),'sheet')){wi++;
    const rid=s.getAttributeNS(NS.r,'id');const rel=wr.find(r=>r.id===rid);if(!rel||rel.type!==RT_SHEET)continue;
    const path=resolveTarget(wbPart,rel.target);if(!files.has(path))continue;
    const sh=readSheet(files,path,sst);sh.name=s.getAttribute('name');sh.wbIndex=wi;
    const pa=dns.find(d=>d.getAttribute('name')==='_xlnm.Print_Area'&&+d.getAttribute('localSheetId')===wi);
    if(pa){const m=/!?\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?/.exec(pa.textContent.split(',')[0].split('!').pop());if(m)sh.printArea=parseRange(m[0].replace(/[$!]/g,''));}sh.hidden=s.getAttribute('state')==='hidden'||s.getAttribute('state')==='veryHidden';sh.el=s;sheets.push(sh);
  }
  if(!sheets.length)throw new Error('לא נמצאו גיליונות בקובץ');
  let heb=false;for(const sh of sheets){if(sh.rtl){heb=true;break;}for(const c of sh.cells.values())if(typeof c.v==='string'&&HEB.test(c.v)){heb=true;break;}if(heb)break;}
  const def=descOf(st,0).font;
  return{files,wbPart,wbDoc,st,sst,sheets,d1904,heb,def};
}

/* ------------------------------------------------------------ formulas: references */
// edits each A1 reference outside "string literals"; fn(sheet, r1,c1,abs, r2,c2,abs2) returns the new text
const REF_RE=/(?<![A-Za-z0-9_.\u0590-\u05FF$])(?:('(?:[^']|'')+'|[A-Za-z_\u0590-\u05FF][A-Za-z0-9_.\u0590-\u05FF]*)!)?(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?::(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7}))?(?![A-Za-z0-9_(!])/g;
function mapRefs(f,fn){
  return String(f).split(/("(?:[^"]|"")*")/).map((part,i)=>i%2?part:part.replace(REF_RE,(m,sh,ac,col,ar,row,ac2,col2,ar2,row2)=>{
    const a={sheet:sh?sh.replace(/^'|'$/g,'').replace(/''/g,"'"):null,c:colIdx(col),r:+row-1,ac:!!ac,ar:!!ar};
    const b=col2?{c:colIdx(col2),r:+row2-1,ac:!!ac2,ar:!!ar2}:null;
    const out=fn(a,b);if(out===undefined)return m;if(out===null)return(sh?sh+'!':'')+'#REF!';
    const w=x=>(x.ac?'$':'')+colName(x.c)+(x.ar?'$':'')+(x.r+1);
    return(sh?sh+'!':'')+w(out[0])+(out[1]?':'+w(out[1]):'');
  })).join('');
}
function moveFormula(f,dr,dc){
  return mapRefs(f,(a,b)=>{const mv=x=>({...x,r:x.ar?x.r:x.r+dr,c:x.ac?x.c:x.c+dc});const A=mv(a),B=b&&mv(b);
    if(A.r<0||A.c<0||(B&&(B.r<0||B.c<0)))return null;return[A,B];});
}
// rows/columns inserted (n>0) or deleted (n<0) at index `at` of sheet `name`
function shiftFormula(f,name,own,axis,at,n){
  return mapRefs(f,(a,b)=>{
    if(a.sheet?a.sheet!==name:!own)return undefined;
    const k=axis==='r'?'r':'c';const A={...a},B=b?{...b}:null;
    if(n>0){if(A[k]>=at)A[k]+=n;if(B&&B[k]>=at)B[k]+=n;return[A,B];}
    const d=-n,end=at+d;
    if(!B){if(A[k]>=end)A[k]-=d;else if(A[k]>=at)return null;return[A,null];}
    let s=A[k],e=B[k];s=s>=end?s-d:(s>=at?at:s);e=e>=end?e-d:(e>=at?at-1:e);if(e<s)return null;A[k]=s;B[k]=e;return[A,B];
  });
}
function shiftRangeStr(sq,axis,at,n){
  return String(sq).split(/\s+/).map(x=>{const g=parseRange(x);if(!g)return x;const k1=axis==='r'?'r1':'c1',k2=axis==='r'?'r2':'c2';
    if(n>0){if(g[k1]>=at)g[k1]+=n;if(g[k2]>=at)g[k2]+=n;}
    else{const d=-n,end=at+d;g[k1]=g[k1]>=end?g[k1]-d:(g[k1]>=at?at:g[k1]);g[k2]=g[k2]>=end?g[k2]-d:(g[k2]>=at?at-1:g[k2]);if(g[k2]<g[k1])return'';}
    return rangeStr(g);}).filter(Boolean).join(' ');
}

/* ------------------------------------------------------------ formulas: calculation */
class XErr{constructor(e){this.e=e;}toString(){return this.e;}}
const ERR=e=>new XErr(e);const isErr=v=>v instanceof XErr;
class NoFn extends Error{}
function tokenize(s){
  const out=[];let i=0;
  const re={ws:/^\s+/,str:/^"((?:[^"]|"")*)"/,num:/^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/,err:/^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A)/i,
    fn:/^([A-Za-z_][A-Za-z0-9_.]*)\s*\(/,
    ref:/^(?:('(?:[^']|'')+'|[A-Za-z_\u0590-\u05FF][A-Za-z0-9_.\u0590-\u05FF]*)!)?\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::\$?([A-Za-z]{1,3})\$?(\d{1,7}))?(?![A-Za-z0-9_(])/,
    col:/^(?:('(?:[^']|'')+'|[A-Za-z_\u0590-\u05FF][A-Za-z0-9_.\u0590-\u05FF]*)!)?\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})(?![A-Za-z0-9_(])/,
    bool:/^(TRUE|FALSE)(?![A-Za-z0-9_(])/i,name:/^[A-Za-z_\u0590-\u05FF][A-Za-z0-9_.\u0590-\u05FF]*/,op:/^(<=|>=|<>|[-+*/^&=<>%])/};
  while(i<s.length){
    const rest=s.slice(i);let m;
    if((m=re.ws.exec(rest))){i+=m[0].length;continue;}
    if((m=re.str.exec(rest))){out.push({t:'str',v:m[1].replace(/""/g,'"')});i+=m[0].length;continue;}
    if((m=re.num.exec(rest))){out.push({t:'num',v:+m[0]});i+=m[0].length;continue;}
    if((m=re.err.exec(rest))){out.push({t:'err',v:m[0].toUpperCase()});i+=m[0].length;continue;}
    if((m=re.fn.exec(rest))){out.push({t:'fn',v:m[1].toUpperCase().replace(/^_XLFN\./,'')});i+=m[0].length;continue;}
    if((m=re.ref.exec(rest))){const sh=m[1]?m[1].replace(/^'|'$/g,'').replace(/''/g,"'"):null;
      out.push({t:'ref',sheet:sh,c1:colIdx(m[2]),r1:+m[3]-1,c2:m[4]?colIdx(m[4]):colIdx(m[2]),r2:m[5]?+m[5]-1:+m[3]-1,range:!!m[4]});i+=m[0].length;continue;}
    if((m=re.col.exec(rest))){const sh=m[1]?m[1].replace(/^'|'$/g,'').replace(/''/g,"'"):null;
      out.push({t:'ref',sheet:sh,c1:colIdx(m[2]),c2:colIdx(m[3]),r1:0,r2:-1,range:true});i+=m[0].length;continue;}
    if((m=re.bool.exec(rest))){out.push({t:'bool',v:m[1].toUpperCase()==='TRUE'});i+=m[0].length;continue;}
    if((m=re.op.exec(rest))){out.push({t:'op',v:m[1]});i+=m[1].length;continue;}
    const ch=s[i];if(ch==='('||ch===')'||ch===','||ch===';'){out.push({t:ch===';'?',':ch});i++;continue;}
    if((m=re.name.exec(rest))){out.push({t:'name',v:m[0]});i+=m[0].length;continue;}
    throw new Error('bad');
  }
  return out;
}
function parseFormula(src){
  const tk=tokenize(src);let p=0;const peek=()=>tk[p],next=()=>tk[p++];
  const isOp=(...o)=>peek()&&peek().t==='op'&&o.includes(peek().v);
  const bin=(sub,ops)=>()=>{let l=sub();while(isOp(...ops)){const o=next().v;const r=sub();l={op:o,l,r};}return l;};
  const primary=()=>{const t=next();if(!t)throw new Error('end');
    if(t.t==='num'||t.t==='str'||t.t==='bool')return{lit:t.v};
    if(t.t==='err')return{lit:ERR(t.v)};
    if(t.t==='ref')return{ref:t};
    if(t.t==='name')return{lit:ERR('#NAME?')};
    if(t.t==='fn'){const args=[];if(peek()&&peek().t===')'){next();return{fn:t.v,args};}
      for(;;){if(peek()&&(peek().t===','||peek().t===')'))args.push({lit:null});else args.push(cmp());const n=next();if(!n)throw new Error('paren');if(n.t===')')break;if(n.t!==',')throw new Error('sep');}
      return{fn:t.v,args};}
    if(t.t==='('){const e=cmp();if(!next()||tk[p-1].t!==')')throw new Error('paren');return e;}
    throw new Error('tok');};
  const postfix=()=>{let e=primary();while(isOp('%')){next();e={op:'%',l:e};}return e;};
  const unary=()=>{if(isOp('-','+')){const o=next().v;return{op:'u'+o,l:unary()};}return postfix();};
  const pow=bin(unary,['^']),mul=bin(pow,['*','/']),add=bin(mul,['+','-']),cat=bin(add,['&']),cmp=bin(cat,['=','<>','<','>','<=','>=']);
  const e=cmp();if(p<tk.length)throw new Error('extra');return e;
}
const toNum=v=>{if(isErr(v))throw v;if(v===null||v===undefined||v==='')return 0;if(typeof v==='boolean')return v?1:0;if(typeof v==='number')return v;
  const n=+String(v).replace(/,/g,'');if(isNaN(n))throw ERR('#VALUE!');return n;};
const toStr=v=>{if(isErr(v))throw v;if(v===null||v===undefined)return'';if(typeof v==='boolean')return v?'TRUE':'FALSE';if(typeof v==='number')return fmtGeneral(v);return String(v);};
const toBool=v=>{if(isErr(v))throw v;if(typeof v==='string'){if(/^true$/i.test(v))return true;if(/^false$/i.test(v))return false;throw ERR('#VALUE!');}return !!toNum(v);};
function fmtGeneral(n){if(!isFinite(n))return'#NUM!';if(Number.isInteger(n))return String(n);const a=Math.abs(n);if(a>=1e11||a<1e-9)return n.toExponential(5).replace(/\.?0+e/,'E').replace('e','E');return String(+n.toPrecision(11));}
function critFn(c){
  if(typeof c==='number')return v=>typeof v==='number'&&v===c;
  const s=String(c);const m=/^(<=|>=|<>|=|<|>)?(.*)$/.exec(s);const op=m[1]||'=';const rhs=m[2];const n=rhs!==''&&!isNaN(+rhs)?+rhs:null;
  if(n!==null)return v=>{if(typeof v!=='number'){if(op==='<>')return true;return false;}return op==='='?v===n:op==='<>'?v!==n:op==='<'?v<n:op==='>'?v>n:op==='<='?v<=n:v>=n;};
  const rx=new RegExp('^'+rhs.replace(/[.+^${}()|[\]\\]/g,'\\$&').replace(/\*/g,'.*').replace(/\?/g,'.')+'$','i');
  return v=>{const t=v===null||v===undefined?'':String(v);return op==='<>'?!rx.test(t):op==='='?rx.test(t):(op==='<'?t<rhs:op==='>'?t>rhs:op==='<='?t<=rhs:t>=rhs);};
}
const serialNow=()=>{const d=new Date();return d.getTime()/864e5+25569-d.getTimezoneOffset()/1440;};
const dateFromSerial=s=>new Date(Math.round((s-25569)*864e5));
function makeCalc(X){
  const memo=new Map(),busy=new Set();
  const sheetIdx=name=>{if(!name)return null;const i=X.st.sheets.findIndex(s=>s.name.toLowerCase()===name.toLowerCase());return i;};
  function cellVal(si,r,c){
    const sh=X.st.sheets[si];if(!sh)return ERR('#REF!');const cell=sh.cells.get(key(r,c));if(!cell)return null;
    if(!cell.f||cell.array)return cell.t==='e'?ERR(cell.v):cell.v;
    const k=si+':'+key(r,c);if(memo.has(k))return memo.get(k);
    if(busy.has(k))return 0;busy.add(k);
    let v;try{v=evalCell(si,cell);}finally{busy.delete(k);}
    memo.set(k,v);return v;
  }
  function evalCell(si,cell){
    try{if(!cell.ast)cell.ast=parseFormula(cell.f);}catch(e){return cell.v!==undefined&&cell.v!==null?cell.v:ERR('#NAME?');}
    try{let v=ev(cell.ast,si);if(v&&v.rng){v=v.rng[0]&&v.rng[0][0]!==undefined?v.rng[0][0]:null;}cell.calcOk=true;return v===null?0:v;}
    catch(e){if(isErr(e)){cell.calcOk=true;return e;}if(e instanceof NoFn){cell.calcOk=false;return cell.t==='e'?ERR(cell.v):cell.v;}cell.calcOk=false;return cell.v;}
  }
  function rangeVals(si,t){
    const s=t.sheet?sheetIdx(t.sheet):si;if(s===null||s<0)throw ERR('#REF!');const sh=X.st.sheets[s];
    const r2=t.r2<0?sh.maxR:t.r2;const rows=[];
    for(let r=t.r1;r<=Math.min(r2,t.r1+100000);r++){const row=[];for(let c=t.c1;c<=t.c2;c++)row.push(cellVal(s,r,c));rows.push(row);}
    return rows;
  }
  function ev(n,si){
    if('lit' in n)return n.lit;
    if(n.ref){const t=n.ref;if(t.range)return{rng:rangeVals(si,t)};const s=t.sheet?sheetIdx(t.sheet):si;if(s===null||s<0)throw ERR('#REF!');const v=cellVal(s,t.r1,t.c1);if(isErr(v))throw v;return v;}
    if(n.op){
      if(n.op==='u-')return -toNum(val(n.l,si));if(n.op==='u+')return toNum(val(n.l,si));if(n.op==='%')return toNum(val(n.l,si))/100;
      const a=val(n.l,si),b=val(n.r,si);
      switch(n.op){
        case'+':return toNum(a)+toNum(b);case'-':return toNum(a)-toNum(b);case'*':return toNum(a)*toNum(b);
        case'/':{const d=toNum(b);if(d===0)throw ERR('#DIV/0!');return toNum(a)/d;}
        case'^':return Math.pow(toNum(a),toNum(b));case'&':return toStr(a)+toStr(b);
        default:{const c=cmpv(a,b);return n.op==='='?c===0:n.op==='<>'?c!==0:n.op==='<'?c<0:n.op==='>'?c>0:n.op==='<='?c<=0:c>=0;}
      }
    }
    if(n.fn)return callFn(n.fn,n.args,si);
    throw ERR('#VALUE!');
  }
  function val(n,si){const v=ev(n,si);if(v&&v.rng){const x=v.rng[0]?v.rng[0][0]:null;if(isErr(x))throw x;return x;}if(isErr(v))throw v;return v;}
  function cmpv(a,b){if(isErr(a))throw a;if(isErr(b))throw b;
    if(a===null)a=typeof b==='string'?'':0;if(b===null)b=typeof a==='string'?'':0;
    if(typeof a==='number'&&typeof b==='number')return a-b;if(typeof a==='string'&&typeof b==='string'){const x=a.toLowerCase(),y=b.toLowerCase();return x<y?-1:x>y?1:0;}
    if(typeof a==='boolean'&&typeof b==='boolean')return(a?1:0)-(b?1:0);const ra=typeof a==='number'?0:typeof a==='string'?1:2,rb=typeof b==='number'?0:typeof b==='string'?1:2;return ra-rb;}
  const flat=v=>v&&v.rng?v.rng.flat():[v];
  function nums(args,si,countAll){const out=[];for(const a of args){const v=ev(a,si);
      if(v&&v.rng){for(const x of v.rng.flat()){if(isErr(x))throw x;if(typeof x==='number')out.push(x);}}
      else{if(isErr(v))throw v;if(v===null)continue;out.push(toNum(v));}}return out;}
  function callFn(name,args,si){
    const A=i=>args[i]?val(args[i],si):null,R=i=>{const v=ev(args[i],si);return v&&v.rng?v.rng:[[v]];};
    switch(name){
      case'SUM':return nums(args,si).reduce((a,b)=>a+b,0);
      case'PRODUCT':return nums(args,si).reduce((a,b)=>a*b,1);
      case'AVERAGE':{const n=nums(args,si);if(!n.length)throw ERR('#DIV/0!');return n.reduce((a,b)=>a+b,0)/n.length;}
      case'MIN':{const n=nums(args,si);return n.length?Math.min(...n):0;}
      case'MAX':{const n=nums(args,si);return n.length?Math.max(...n):0;}
      case'MEDIAN':{const n=nums(args,si).sort((a,b)=>a-b);if(!n.length)throw ERR('#NUM!');const m=n.length>>1;return n.length%2?n[m]:(n[m-1]+n[m])/2;}
      case'COUNT':{let k=0;for(const a of args){const v=ev(a,si);for(const x of flat(v))if(typeof x==='number')k++;}return k;}
      case'COUNTA':{let k=0;for(const a of args){const v=ev(a,si);for(const x of flat(v))if(x!==null&&x!=='')k++;}return k;}
      case'COUNTBLANK':{let k=0;for(const x of R(0).flat())if(x===null||x==='')k++;return k;}
      case'IF':{const c=toBool(A(0));return c?(args[1]?val(args[1],si):true):(args[2]?val(args[2],si):false);}
      case'IFERROR':try{return val(args[0],si);}catch(e){if(isErr(e))return A(1);throw e;}
      case'IFNA':try{return val(args[0],si);}catch(e){if(isErr(e)&&e.e==='#N/A')return A(1);throw e;}
      case'AND':{let r=true;for(const a of args)for(const x of flat(ev(a,si)))if(x!==null&&!toBool(x))r=false;return r;}
      case'OR':{let r=false;for(const a of args)for(const x of flat(ev(a,si)))if(x!==null&&toBool(x))r=true;return r;}
      case'NOT':return !toBool(A(0));
      case'ROUND':{const p=Math.pow(10,toNum(A(1)));return Math.round(toNum(A(0))*p+(toNum(A(0))>=0?1e-9:-1e-9))/p;}
      case'ROUNDUP':{const p=Math.pow(10,toNum(A(1))),x=toNum(A(0));return(x<0?-Math.ceil(-x*p-1e-9):Math.ceil(x*p-1e-9))/p;}
      case'ROUNDDOWN':{const p=Math.pow(10,toNum(A(1))),x=toNum(A(0));return Math.trunc(x*p+1e-9*Math.sign(x))/p;}
      case'INT':return Math.floor(toNum(A(0)));case'ABS':return Math.abs(toNum(A(0)));
      case'MOD':{const d=toNum(A(1));if(d===0)throw ERR('#DIV/0!');const x=toNum(A(0));return x-d*Math.floor(x/d);}
      case'POWER':return Math.pow(toNum(A(0)),toNum(A(1)));
      case'SQRT':{const x=toNum(A(0));if(x<0)throw ERR('#NUM!');return Math.sqrt(x);}
      case'CONCATENATE':case'CONCAT':{let s='';for(const a of args)for(const x of flat(ev(a,si)))s+=toStr(x);return s;}
      case'TEXTJOIN':{const d=toStr(A(0)),ig=toBool(A(1));const parts=[];for(const a of args.slice(2))for(const x of flat(ev(a,si))){const t=toStr(x);if(!(ig&&t===''))parts.push(t);}return parts.join(d);}
      case'LEN':return toStr(A(0)).length;
      case'LEFT':return toStr(A(0)).slice(0,args[1]?toNum(A(1)):1);
      case'RIGHT':{const s=toStr(A(0)),n=args[1]?toNum(A(1)):1;return n?s.slice(-n):'';}
      case'MID':return toStr(A(0)).substr(toNum(A(1))-1,toNum(A(2)));
      case'UPPER':return toStr(A(0)).toUpperCase();case'LOWER':return toStr(A(0)).toLowerCase();
      case'TRIM':return toStr(A(0)).trim().replace(/ +/g,' ');
      case'VALUE':return toNum(A(0));
      case'TEXT':return fmtValue(toNum(A(0)),toStr(A(1)));
      case'TODAY':return Math.floor(serialNow());case'NOW':return serialNow();
      case'YEAR':return dateFromSerial(toNum(A(0))).getUTCFullYear();case'MONTH':return dateFromSerial(toNum(A(0))).getUTCMonth()+1;case'DAY':return dateFromSerial(toNum(A(0))).getUTCDate();
      case'DATE':return Date.UTC(toNum(A(0)),toNum(A(1))-1,toNum(A(2)))/864e5+25569;
      case'ISBLANK':{const v=args[0]?ev(args[0],si):null;const x=v&&v.rng?v.rng[0][0]:v;return x===null||x==='';}
      case'ISNUMBER':try{return typeof val(args[0],si)==='number';}catch(_){return false;}
      case'ISTEXT':try{return typeof val(args[0],si)==='string';}catch(_){return false;}
      case'ISERROR':try{val(args[0],si);return false;}catch(e){if(isErr(e))return true;throw e;}
      case'SUMIF':case'COUNTIF':case'AVERAGEIF':{
        const rng=R(0).flat(),f=critFn(A(1));const sr=name==='SUMIF'||name==='AVERAGEIF'?(args[2]?R(2).flat():rng):null;let s=0,k=0;
        rng.forEach((v,i)=>{if(f(v)){k++;if(sr){const x=sr[i];if(typeof x==='number')s+=x;}}});
        if(name==='COUNTIF')return k;if(name==='SUMIF')return s;if(!k)throw ERR('#DIV/0!');return s/k;}
      case'SUMIFS':case'COUNTIFS':{
        const off=name==='SUMIFS'?1:0;const sr=off?R(0).flat():null;let s=0,k=0;const pairs=[];
        for(let i=off;i<args.length;i+=2)pairs.push([R(i).flat(),critFn(A(i+1))]);
        const n=pairs[0]?pairs[0][0].length:0;for(let i=0;i<n;i++)if(pairs.every(([r,f])=>f(r[i]))){k++;if(sr&&typeof sr[i]==='number')s+=sr[i];}
        return off?s:k;}
      case'VLOOKUP':case'HLOOKUP':{
        const q=A(0),tbl=R(1),idx=toNum(A(2))-1,approx=args[3]?toBool(A(3)):true;
        const lines=name==='VLOOKUP'?tbl:tbl[0].map((_,c)=>tbl.map(r=>r[c]));
        let hit=-1;
        if(!approx){hit=lines.findIndex(l=>cmpv(l[0],q)===0);}
        else for(let i=0;i<lines.length;i++){const v=lines[i][0];if(v===null)continue;if(cmpv(v,q)<=0)hit=i;else break;}
        if(hit<0)throw ERR('#N/A');const v=lines[hit][idx];if(v===undefined)throw ERR('#REF!');return v;}
      case'INDEX':{const t=R(0);const r=toNum(A(1)),c=args[2]?toNum(A(2)):1;const row=t.length===1&&!args[2]?t[0]:null;
        if(row)return row[r-1]??ERR('#REF!');const v=(t[r-1]||[])[c-1];if(v===undefined)throw ERR('#REF!');return v;}
      case'MATCH':{const q=A(0),arr=R(1).flat(),type=args[2]?toNum(A(2)):1;
        if(type===0){const i=arr.findIndex(v=>cmpv(v,q)===0);if(i<0)throw ERR('#N/A');return i+1;}
        let hit=-1;for(let i=0;i<arr.length;i++){if(arr[i]===null)continue;const c=cmpv(arr[i],q);if(type>0?c<=0:c>=0)hit=i;else break;}if(hit<0)throw ERR('#N/A');return hit+1;}
      default:throw new NoFn(name);
    }
  }
  return{value:cellVal,reset:()=>memo.clear()};
}

/* ------------------------------------------------------------ number formats */
function fmtValue(v,code){
  if(typeof v!=='number')return v;
  if(!code||code==='General')return fmtGeneral(v);
  try{const S2=window.XLSX&&XLSX.SSF;if(S2)return S2.format(code,v);}catch(_){}
  return fmtGeneral(v);
}
const isDateFmt=c=>!!c&&/(^|[^"\\])(d|m|y|h|s)/i.test(c.replace(/"[^"]*"|\[[^\]]*\]/g,''))&&!/^[#0.,%E+-]+$/.test(c);
const FMTS=[['General','כללי'],['#,##0.00','מספר'],['[$₪-40D] #,##0.00','מטבע ₪'],['0.00%','אחוזים'],['dd/mm/yyyy','תאריך'],['h:mm','שעה'],['@','טקסט']];

/* ------------------------------------------------------------ the editor screen */
const ICO={
  undo:'<path d="M7 5L3 9l4 4"/><path d="M3 9h9a5 5 0 0 1 0 10H8"/>',redo:'<path d="M13 5l4 4-4 4"/><path d="M17 9H8a5 5 0 0 0 0 10h4"/>',
  cut:'<circle cx="6" cy="15" r="2.5"/><circle cx="14" cy="15" r="2.5"/><path d="M7.5 13L14 3M12.5 13L6 3"/>',
  copy:'<rect x="7" y="7" width="10" height="11" rx="1.5"/><path d="M4 13V4a1 1 0 0 1 1-1h8"/>',
  paste:'<rect x="4" y="4" width="12" height="14" rx="1.5"/><path d="M7.5 4V2.5h5V4"/><path d="M7 9h6M7 12h6M7 15h4"/>',
  aright:'<path d="M3 4h14M7 8h10M3 12h14M7 16h10"/>',acenter:'<path d="M3 4h14M5.5 8h9M3 12h14M5.5 16h9"/>',aleft:'<path d="M3 4h14M3 8h10M3 12h14M3 16h10"/>',
  vtop:'<path d="M3 3h14"/><path d="M10 6v11M7 9l3-3 3 3"/>',vmid:'<path d="M3 10h14"/><path d="M10 2v5M8 5l2 2 2-2M10 18v-5M8 15l2-2 2 2"/>',vbot:'<path d="M3 17h14"/><path d="M10 3v11M7 11l3 3 3-3"/>',
  wrap:'<path d="M3 5h14M3 10h11a3 3 0 0 1 0 6h-4"/><path d="M12 14l-2 2 2 2M3 15h4"/>',
  merge:'<rect x="2.5" y="4" width="15" height="12" rx="1"/><path d="M6 10h8M6 10l2-2M6 10l2 2M14 10l-2-2M14 10l-2 2"/>',
  borders:'<rect x="3" y="3" width="14" height="14"/><path d="M3 10h14M10 3v14"/>',
  rowIns:'<rect x="3" y="9" width="14" height="8" rx="1"/><path d="M10 2v5M7.5 4.5h5"/>',colIns:'<rect x="3" y="3" width="8" height="14" rx="1"/><path d="M13 10h5M15.5 7.5v5"/>',
  rowDel:'<rect x="3" y="6" width="14" height="8" rx="1"/><path d="M8 8.5l4 3M12 8.5l-4 3"/>',colDel:'<rect x="6" y="3" width="8" height="14" rx="1"/><path d="M8 8l4 4M12 8l-4 4"/>',
  sum:'<path d="M15 4H5l6 6-6 6h10"/>',find:'<circle cx="8.5" cy="8.5" r="5"/><path d="M12.5 12.5L17 17"/>',
  clear:'<path d="M5 15l3.5-11h3L15 15M6.3 11h7.4"/><path d="M13 13l4 4M17 13l-4 4"/>',
  decP:'<path d="M2 15h2M6 11a2 2 0 0 1 4 0v2a2 2 0 0 1-4 0zM12 11a2 2 0 0 1 4 0v2a2 2 0 0 1-4 0z"/><path d="M13 5h5M16 3l2 2-2 2"/>',
  decM:'<path d="M2 15h2M6 11a2 2 0 0 1 4 0v2a2 2 0 0 1-4 0z"/><path d="M18 5h-5M15 3l-2 2 2 2"/>',
  grow:'<path d="M3 16l4.5-12L12 16M4.6 12h5.8"/><path d="M14 7l2-2.5L18 7"/>',shrink:'<path d="M3 16l4.5-12L12 16M4.6 12h5.8"/><path d="M14 4.5l2 2.5 2-2.5"/>',
  zin:'<circle cx="8.5" cy="8.5" r="5"/><path d="M12.5 12.5L17 17M6.5 8.5h4M8.5 6.5v4"/>',zout:'<circle cx="8.5" cy="8.5" r="5"/><path d="M12.5 12.5L17 17M6.5 8.5h4"/>',
  grid:'<rect x="3" y="3" width="14" height="14"/><path d="M3 7.7h14M3 12.3h14M7.7 3v14M12.3 3v14"/>',sortAZ:'<path d="M4 4h5M4 9h4M4 14h3"/><path d="M14 3v13M11 13l3 3 3-3"/>',sortZA:'<path d="M4 4h3M4 9h4M4 14h5"/><path d="M14 3v13M11 13l3 3 3-3"/>',
  margins:'<rect x="4" y="2.5" width="12" height="15" rx="1"/><path d="M4 6h12M4 14h12M7 2.5v15M13 2.5v15" stroke-dasharray="1.5 1.5"/>',
  orient:'<rect x="2.5" y="6" width="15" height="10" rx="1"/><path d="M6 2.5h5M11 2.5l-1.5-1M11 2.5l-1.5 1"/>',paper:'<path d="M5 2.5h7l3.5 3.5v11.5H5z"/><path d="M12 2.5V6h3.5"/>',
  fit:'<rect x="3" y="3" width="14" height="14" rx="1"/><path d="M7 7l6 6M13 7v6H7"/>',area:'<rect x="3" y="4" width="14" height="12" rx="1" stroke-dasharray="2 1.5"/><path d="M7 9h6M7 12h4"/>',
  areaX:'<rect x="3" y="4" width="14" height="12" rx="1" stroke-dasharray="2 1.5"/><path d="M8 8l4 4M12 8l-4 4"/>',hcenter:'<path d="M3 4h14M5.5 8h9M3 12h14M5.5 16h9"/><path d="M10 2v16" stroke-dasharray="1 1.5"/>',
  print:'<path d="M5.5 7V3h9v4"/><rect x="3" y="7" width="14" height="7" rx="1.5"/><path d="M6 12h8v5H6z"/>',heads:'<rect x="3" y="3" width="14" height="14" rx="1"/><path d="M3 7h14M7 3v14"/><path d="M10 5h1M14 5h1M4.5 10v1M4.5 14v1"/>',
  fx:'<path d="M9 4c-2 0-2.5 1.5-2.8 3L5 16c-.3 1.5-1 2-2 2M3.5 9h6"/><path d="M11 9l5 6M16 9l-5 6"/>',pbreak:'<rect x="3" y="3" width="14" height="14" rx="1"/><path d="M10 3v14M3 10h14" stroke-dasharray="2 1.5"/>',
  freeze:'<rect x="3" y="3" width="14" height="14" rx="1"/><path d="M3 7.5h14M7.5 3v14" stroke-width="2.2"/>',
  caret:'<path d="M6 8l4 4 4-4"/>',addSheet:'<path d="M10 4v12M4 10h12"/>'
};
const ic=n=>`<svg viewBox="0 0 20 20" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">${ICO[n]||''}</svg>`;
const B=(cmd,icon,title,extra='')=>`<button type="button" class="xb" data-cmd="${cmd}" title="${title}" aria-label="${title}" ${extra}>${icon}</button>`;
const BL=(cmd,icon,label,title,extra='')=>`<button type="button" class="xb big" data-cmd="${cmd}" title="${title||label}" ${extra}>${ic(icon)}<span>${label}</span></button>`;
const G=(label,inner)=>`<div class="xg"><div class="xg-b">${inner}</div><div class="xg-l">${label}</div></div>`;
const FONTS=['Arial','David','Calibri','Tahoma','Times New Roman','FrankRuehl','Miriam','Narkisim','Heebo','Rubik','Segoe UI','Courier New','Verdana'];
const SIZES=[8,9,10,11,12,14,16,18,20,22,24,26,28,36,48,72];
const COLORS=['000000','404040','7F7F7F','BFBFBF','FFFFFF','C00000','FF0000','FFC000','FFFF00','92D050','00B050','00B0F0','0070C0','002060','7030A0','2E58CC','1F3864','833C0B','385623','7F6000'];
const FILLS=['FFFF00','FFF2CC','FCE4D6','E2EFDA','DDEBF7','D9E1F2','EDEDED','F8CBAD','C6E0B4','BDD7EE','FFD966','F4B084','A9D08E','9BC2E6','8EA9DB','FFC000','FF0000','00B050','2E58CC','16213A'];
function ribbonHtml(){return `
<div class="xpanel" data-panel="home">
 ${G('ביטול',B('undo',ic('undo'),'ביטול (Ctrl+Z)')+B('redo',ic('redo'),'חזרה (Ctrl+Y)'))}
 ${G('לוח',BL('paste','paste','הדבקה','הדבקה (Ctrl+V)')+`<div class="col">${B('cut',ic('cut'),'גזירה (Ctrl+X)')}${B('copy',ic('copy'),'העתקה (Ctrl+C)')}</div>`)}
 ${G('גופן',`<div class="row"><select class="xsel font" data-in="font" aria-label="גופן" title="גופן">${FONTS.map(f=>`<option value="${f}" style="font-family:'${f}'">${f}</option>`).join('')}</select><input class="xsel size" data-in="size" list="xeSizes" inputmode="decimal" aria-label="גודל גופן" title="גודל גופן"><datalist id="xeSizes">${SIZES.map(s=>`<option value="${s}">`).join('')}</datalist>${B('grow',ic('grow'),'הגדלת הגופן')}${B('shrink',ic('shrink'),'הקטנת הגופן')}</div>
  <div class="row">${B('bold','<b class="gl">B</b>','מודגש (Ctrl+B)')}${B('italic','<i class="gl" style="font-family:Georgia,serif">I</i>','נטוי (Ctrl+I)')}${B('underline','<u class="gl">U</u>','קו תחתי (Ctrl+U)')}${B('strike','<s class="gl">ab</s>','קו חוצה')}
  <span class="xsplit">${B('borders',ic('borders'),'גבולות לכל התאים')}${B('bordersMenu',ic('caret'),'עוד גבולות','data-pop="borders"')}</span>
  <span class="xsplit">${B('fill','<svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round" aria-hidden="true"><path d="M4 10l6-6 6 6-6 6z"/><path d="M16 12s2 2.3 2 3.5a2 2 0 0 1-4 0c0-1.2 2-3.5 2-3.5z"/></svg><i class="cbar" data-bar="fill"></i>','צבע מילוי')}${B('fillMenu',ic('caret'),'בחירת צבע מילוי','data-pop="fill"')}</span>
  <span class="xsplit">${B('color','<span class="gl a">A</span><i class="cbar" data-bar="color"></i>','צבע גופן')}${B('colorMenu',ic('caret'),'בחירת צבע גופן','data-pop="color"')}</span></div>`)}
 ${G('יישור',`<div class="row">${B('vtop',ic('vtop'),'יישור למעלה')}${B('vmid',ic('vmid'),'יישור לאמצע')}${B('vbot',ic('vbot'),'יישור למטה')}${B('wrap',ic('wrap'),'גלישת טקסט')}</div>
  <div class="row">${B('right',ic('aright'),'יישור לימין')}${B('center',ic('acenter'),'מרכוז')}${B('left',ic('aleft'),'יישור לשמאל')}${B('merge',ic('merge'),'מיזוג ומרכוז')}</div>`)}
 ${G('מספר',`<div class="row"><select class="xsel fmt" data-in="fmt" aria-label="עיצוב מספר" title="עיצוב מספר">${FMTS.map(f=>`<option value="${esc(f[0])}">${f[1]}</option>`).join('')}<option value="__custom" hidden>מותאם</option></select></div>
  <div class="row">${B('pct','<b class="gl sm">%</b>','סגנון אחוזים')}${B('thou','<b class="gl sm">000</b>','מפריד אלפים')}${B('decP',ic('decP'),'הוספת ספרה אחרי הנקודה')}${B('decM',ic('decM'),'הסרת ספרה אחרי הנקודה')}</div>`)}
 ${G('תאים',`<div class="row">${B('rowIns',ic('rowIns'),'הוספת שורה מעל')}${B('colIns',ic('colIns'),'הוספת עמודה לפני')}</div><div class="row">${B('rowDel',ic('rowDel'),'מחיקת שורות')}${B('colDel',ic('colDel'),'מחיקת עמודות')}</div>`)}
 ${G('עריכה',BL('sum','sum','סכום','סכום אוטומטי (Alt+=)')+`<div class="col">${B('sortAZ',ic('sortAZ'),'מיון מהקטן לגדול')}${B('sortZA',ic('sortZA'),'מיון מהגדול לקטן')}${B('clear',ic('clear'),'ניקוי העיצוב')}</div>`+BL('find','find','חיפוש','חיפוש (Ctrl+F)'))}
</div>
<div class="xpanel" data-panel="layout" hidden>
 ${G('הגדרת עמוד',BL('marginsMenu','margins','שוליים','שולי העמוד','data-pop="margins"')+BL('orientMenu','orient','כיוון','לאורך או לרוחב','data-pop="orient"')+BL('paperMenu','paper','גודל','גודל הנייר','data-pop="paper"'))}
 ${G('התאמה לעמוד',BL('fitMenu','fit','התאמה','כמה עמודים יודפסו','data-pop="fit"')+'<div class="col" style="justify-content:center"><label class="num">קנה מידה<input type="number" min="10" max="400" step="5" data-in="scale" dir="ltr"> %</label></div>')}
 ${G('אזור הדפסה',BL('paSet','area','הגדרה','הדפסה רק של התאים המסומנים')+BL('paClear','areaX','ניקוי','ביטול אזור ההדפסה'))}
 ${G('אפשרויות',BL('printGrid','grid','קווי רשת בהדפסה','הדפסת קווי הרשת','data-toggle')+BL('centerH','hcenter','מרכוז בדף','מרכוז הטבלה לרוחב הדף','data-toggle'))}
 ${G('הדפסה',BL('print','print','הדפסה','הדפסה (Ctrl+P)'))}
</div>
<div class="xpanel" data-panel="view" hidden>
 ${G('זום',B('zout',ic('zout'),'הקטנה')+'<span class="zv" id="xeZoomV">100%</span>'+B('zin',ic('zin'),'הגדלה')+'<button type="button" class="xb txt" data-cmd="z100">100%</button><button type="button" class="xb txt" data-cmd="zsel">לפי הבחירה</button>')}
 ${G('הצגה',BL('gridlines','grid','קווי רשת','הצגת קווי הרשת','data-toggle aria-pressed="true"')+BL('headings','heads','כותרות','מספרי השורות ואותיות העמודות','data-toggle aria-pressed="true"')+BL('formulas','fx','נוסחאות','הצגת הנוסחאות במקום התוצאות','data-toggle')+BL('breaks','pbreak','גבולות עמוד','הצגת המקום שבו כל עמוד מודפס נגמר','data-toggle'))}
 ${G('חלון',BL('freezeMenu','freeze','הקפאה','הקפאת שורות ועמודות כך שיישארו במקום בגלילה','data-pop="freeze"'))}
 ${G('הדפסה',BL('print','print','הדפסה','הדפסה (Ctrl+P)'))}
</div>`;}
function shellHtml(){return `
<div class="xe-top">
  <div class="xe-brand"><span class="xe-mark" aria-hidden="true">X</span><b>סימן</b><span class="xe-sub">עורך Excel</span></div>
  <label class="xe-namebox"><input class="xe-name" id="xeName" aria-label="שם הקובץ" dir="auto" spellcheck="false"><span class="xe-ext">.xlsx</span></label>
  <span class="xe-state" id="xeState" role="status"></span><span class="grow"></span>
  <button type="button" class="xe-tb" id="xeDesk" title="פתיחת הגיליון לסימון, חתימה ושיתוף">לחתימה ולסימון</button>
  <button type="button" class="xe-tb" id="xePdf" title="שמירת הגיליון כקובץ PDF">שמירה כ-PDF</button>
  <button type="button" class="xe-tb pri" id="xeSave" title="שמירה כקובץ Excel (Ctrl+S)">שמירה</button>
  <button type="button" class="xe-x" id="xeClose" aria-label="סגירת העורך" title="סגירה">×</button>
</div>
<div class="xe-tabs" role="tablist"><button type="button" role="tab" data-tab="home" aria-selected="true">בית</button><button type="button" role="tab" data-tab="layout" aria-selected="false">פריסה</button><button type="button" role="tab" data-tab="view" aria-selected="false">תצוגה</button></div>
<div class="xe-ribbon" id="xeRibbon">${ribbonHtml()}</div>
<div class="xe-fbar"><input id="xeRef" class="xe-ref" aria-label="כתובת התא" dir="ltr" spellcheck="false"><span class="xe-fx">fx</span><input id="xeFx" class="xe-fxin" aria-label="תוכן התא" dir="auto" spellcheck="false" autocomplete="off"></div>
<div class="xe-wrap" id="xeWrap" tabindex="0" aria-label="תאי הגיליון"><table class="xe-grid" id="xeGrid"></table><textarea class="xe-edit" id="xeEdit" hidden spellcheck="false" aria-label="עריכת תא"></textarea></div>
<div class="xe-bottom"><div class="xe-sheets" id="xeSheets"></div><span class="grow"></span><span class="xe-stat" id="xeStat"></span>
  <button type="button" class="xb" data-cmd="zout" aria-label="הקטנה">−</button><input type="range" id="xeZoom" min="50" max="200" step="10" value="100" aria-label="זום"><button type="button" class="xb" data-cmd="zin" aria-label="הגדלה">+</button></div>
<div class="xe-find" id="xeFind" hidden><input id="xeFindIn" placeholder="חיפוש בגיליון" dir="auto" aria-label="חיפוש"><button type="button" class="xb" id="xeFindNext" aria-label="הבא">▼</button><button type="button" class="xb" id="xeFindX" aria-label="סגירה">×</button></div>
<div class="xe-pop" id="xePop" hidden></div>
<div id="xePrint"></div>
<div class="xe-dlg" id="xeDlg" hidden><div class="xe-dlg-in" role="dialog" aria-modal="true"></div></div>`;}
const CSS=`
#xed{--xn:var(--navy,#16213A);--xm:var(--marker,#FFD43B);--xg:#217346;--xl:var(--line,#D4DBE5);--xt:var(--ink,#16213A);--xmu:var(--muted,#5D6982);
  position:fixed;inset:0;z-index:900;display:flex;flex-direction:column;background:#fff;color:var(--xt);direction:rtl;font:14px/1.3 "SF-heebo","Heebo",system-ui,sans-serif;-webkit-user-select:none;user-select:none}
#xed button{font:inherit;color:inherit}#xed .grow{flex:1}
#xed .xe-top{display:flex;align-items:center;gap:8px;padding:7px 12px;background:var(--xn);color:#E4EAF4;min-height:48px}
#xed .xe-brand{display:flex;align-items:center;gap:7px;white-space:nowrap}
#xed .xe-brand b{font-size:17px;color:#fff;background:linear-gradient(transparent 58%,rgba(255,212,59,.55) 58%);padding:0 2px}
#xed .xe-mark{width:24px;height:24px;border-radius:5px;background:var(--xg);color:#fff;display:grid;place-items:center;font:700 13px Arial,sans-serif}
#xed .xe-sub{color:#9AA7BE;font-size:13px}
#xed .xe-namebox{direction:ltr;display:flex;align-items:center;max-width:340px;min-width:90px;flex:0 1 340px;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.14);border-radius:7px;padding:0 8px;margin-inline-start:10px}
#xed .xe-name{flex:1;min-width:0;background:none;border:0;color:#fff;font:inherit;padding:6px 0;outline:none}
#xed .xe-namebox:focus-within{border-color:var(--xm)}#xed .xe-ext{color:#9AA7BE;font-size:12px}
#xed .xe-state{font-size:12px;color:#9AA7BE;white-space:nowrap}#xed .xe-state.dirty{color:var(--xm)}
#xed .xe-tb{border:1px solid rgba(255,255,255,.22);background:rgba(255,255,255,.06);color:#fff;border-radius:8px;padding:7px 13px;cursor:pointer;white-space:nowrap}
#xed .xe-tb:hover{background:rgba(255,255,255,.14)}#xed .xe-tb.pri{background:var(--xm);border-color:var(--xm);color:var(--xn);font-weight:700}
#xed .xe-x{width:34px;height:34px;border-radius:8px;border:0;background:none;color:#C9D2E1;font-size:22px;cursor:pointer}#xed .xe-x:hover{background:rgba(255,255,255,.12);color:#fff}
#xed .xe-tabs{display:flex;gap:2px;padding:0 10px;background:var(--xn)}
#xed .xe-tabs button{border:0;background:none;color:#C9D2E1;padding:7px 14px 8px;border-radius:8px 8px 0 0;cursor:pointer;position:relative}
#xed .xe-tabs button[aria-selected=true]{background:#F7F8FB;color:var(--xn);font-weight:600}
#xed .xe-tabs button[aria-selected=true]::after{content:"";position:absolute;inset-inline:12px;bottom:3px;height:3px;border-radius:2px;background:var(--xm)}
#xed .xe-ribbon{background:#F7F8FB;border-bottom:1px solid var(--xl);overflow-x:auto;overflow-y:hidden;scrollbar-width:thin}
#xed .xpanel{display:flex;align-items:stretch;min-height:86px;padding:6px 6px 2px;width:max-content}#xed .xpanel[hidden]{display:none}
#xed .xg{display:flex;flex-direction:column;padding:0 8px;border-inline-start:1px solid #E1E6EE}#xed .xg:first-child{border-inline-start:0}
#xed .xg-b{flex:1;display:flex;align-items:center;gap:3px}#xed .xg-b:has(.row){flex-direction:column;align-items:stretch;justify-content:center;gap:4px}
#xed .xg-l{font-size:11px;color:var(--xmu);text-align:center;padding:2px 0 3px;white-space:nowrap}
#xed .xg .row{display:flex;align-items:center;gap:2px}#xed .xg .col{display:flex;flex-direction:column;gap:1px}
#xed .xb{min-width:30px;height:30px;border:1px solid transparent;border-radius:6px;background:none;cursor:pointer;display:inline-grid;place-items:center;padding:0 4px;position:relative}
#xed .xb svg{width:19px;height:19px}#xed .xb:hover{background:#E8ECF3}
#xed .xb[aria-pressed=true]{background:#FFF3C4;border-color:#F2D46B}#xed .xb:disabled{opacity:.35;cursor:default}
#xed .xb.big{flex-direction:column;height:auto;min-height:60px;min-width:52px;gap:3px;padding:4px 6px;font-size:12px;display:inline-flex;align-items:center;justify-content:center}
#xed .xb.big svg{width:24px;height:24px}#xed .xb.txt{font-size:12px;padding:0 8px}
#xed .xb .gl{font:700 16px/1 Georgia,serif}#xed .xb .gl.a{font:600 16px/1 Arial,sans-serif;margin-top:-4px}#xed .xb .gl.sm{font:700 12px/1 Arial,sans-serif}
#xed .xb .cbar{position:absolute;inset-inline:6px;bottom:4px;height:4px;border-radius:1px;background:#C00000}#xed .xb .cbar[data-bar=fill]{background:#FFFF00}
#xed .xsplit{display:inline-flex}#xed .xsplit .xb:first-child{border-start-end-radius:0;border-end-end-radius:0;min-width:28px}
#xed .xsplit .xb:last-child{min-width:16px;padding:0}#xed .xsplit .xb:last-child svg{width:12px;height:12px}
#xed .xsel{height:28px;border:1px solid #CBD3DF;border-radius:6px;background:#fff;font:inherit;font-size:13px;padding:0 6px;color:var(--xt)}
#xed .xsel.font{width:140px}#xed .xsel.size{width:52px;direction:ltr;text-align:center}#xed .xsel.fmt{width:150px}
#xed .zv{min-width:44px;text-align:center;font-size:13px}
#xed .xe-fbar{display:flex;align-items:center;gap:6px;padding:4px 8px;border-bottom:1px solid var(--xl);background:#fff}
#xed .xe-ref{width:90px;height:26px;border:1px solid #CBD3DF;border-radius:5px;padding:0 6px;font:13px Consolas,monospace;text-align:center}
#xed .xe-fx{font:italic 600 14px Georgia,serif;color:var(--xmu);padding:0 4px;border-inline:1px solid #E1E6EE}
#xed .xe-fxin{flex:1;height:26px;border:1px solid transparent;border-radius:5px;padding:0 6px;font:14px Arial,"Heebo",sans-serif;outline:none;-webkit-user-select:text;user-select:text}
#xed .xe-fxin:focus{border-color:var(--xg)}
#xed .xe-wrap{flex:1;overflow:auto;position:relative;outline:none;background:#fff;overscroll-behavior:contain}
#xed .xe-grid{border-collapse:collapse;table-layout:fixed;font:11pt Arial,sans-serif;color:#000;width:max-content}
#xed .xe-grid td{height:20px;padding:0 3px;border:1px solid #E1E5EA;white-space:nowrap;overflow:visible;vertical-align:bottom;position:relative;cursor:cell;line-height:1.2}
#xed .xe-grid.nogrid td{border-color:transparent}
#xed .xe-grid td>div{overflow:visible;white-space:nowrap}
#xed .xe-grid td.w>div{white-space:pre-wrap;overflow-wrap:anywhere;overflow:hidden}
#xed .xe-grid td.v{background:#fff}
#xed .xe-grid th{position:sticky;background:#F3F5F8;color:#5D6982;font:11px Arial,sans-serif;border:1px solid #D9DEE6;z-index:2;cursor:default;padding:0}
#xed .xe-grid thead th{top:0;height:22px}#xed .xe-grid tbody th{inset-inline-start:0;min-width:44px;text-align:center}
#xed .xe-grid thead th.corner{inset-inline-start:0;z-index:3}
#xed .xe-grid th.on{background:#D7E6DC;color:#14532D;font-weight:700}
#xed .xe-grid th .rz{position:absolute;top:0;bottom:0;inset-inline-end:-3px;width:6px;cursor:col-resize;z-index:4}
#xed .xe-grid tr:not(.fzr)>td.xin:not(.fz){background-color:rgba(33,115,70,.12)!important}
#xed .xe-grid td.xact{outline:2px solid var(--xg);outline-offset:-2px;z-index:1}
#xed .xe-grid tr.sp td{border:0;padding:0;height:auto}
#xed .xe-edit{position:absolute;z-index:5;margin:0;border:2px solid var(--xg);padding:1px 3px;font:11pt Arial,sans-serif;resize:none;overflow:hidden;background:#fff;box-sizing:border-box;outline:none;-webkit-user-select:text;user-select:text;white-space:pre}
#xed .xe-bottom{display:flex;align-items:center;gap:10px;padding:2px 8px;background:#F7F8FB;border-top:1px solid var(--xl);min-height:32px;font-size:12px;color:var(--xmu)}
#xed .xe-sheets{display:flex;gap:2px;overflow-x:auto;max-width:60vw}
#xed .xe-sheets button{border:1px solid transparent;border-radius:0 0 6px 6px;background:none;padding:4px 12px;cursor:pointer;white-space:nowrap;font-size:13px}
#xed .xe-sheets button.on{background:#fff;border-color:var(--xl);border-top-color:#fff;color:var(--xg);font-weight:700;box-shadow:inset 0 -2px 0 var(--xg)}
#xed .xe-sheets .add{font-size:16px;padding:2px 10px}
#xed .xe-stat{white-space:nowrap}
#xed .xe-bottom input[type=range]{width:100px;accent-color:var(--xg)}
#xed .xe-find{position:absolute;top:150px;inset-inline-start:24px;z-index:20;background:#fff;border:1px solid var(--xl);border-radius:10px;box-shadow:0 12px 30px -10px rgba(22,33,58,.35);padding:6px;display:flex;gap:4px}
#xed .xe-find[hidden]{display:none}#xed .xe-find input{width:220px;height:30px;border:1px solid #CBD3DF;border-radius:6px;padding:0 8px;font:inherit}
#xed .xe-pop{position:fixed;z-index:30;background:#fff;border:1px solid var(--xl);border-radius:10px;box-shadow:0 14px 34px -12px rgba(22,33,58,.4);padding:8px;max-width:340px}
#xed .xe-pop[hidden]{display:none}#xed .xe-pop h4{margin:2px 2px 6px;font-size:12px;color:var(--xmu);font-weight:600}
#xed .xe-pop .pal{display:grid;grid-template-columns:repeat(10,22px);gap:4px}
#xed .xe-pop .pal button{width:22px;height:22px;border-radius:4px;border:1px solid rgba(0,0,0,.15);cursor:pointer;padding:0}
#xed .xe-pop .mi{display:flex;align-items:center;gap:8px;width:100%;border:0;background:none;padding:7px 10px;border-radius:7px;cursor:pointer;text-align:start;white-space:nowrap}
#xed .xe-pop .mi:hover{background:#EEF2F8}#xed .xe-pop .sep{height:1px;background:var(--xl);margin:5px 0}
#xed .xe-dlg{position:absolute;inset:0;z-index:40;background:rgba(22,33,58,.35);display:grid;place-items:center}#xed .xe-dlg[hidden]{display:none}
#xed .xe-dlg-in{background:#fff;border-radius:14px;padding:18px 20px;width:min(420px,92vw);box-shadow:0 24px 60px -20px rgba(0,0,0,.45)}
#xed .xe-dlg-in h3{margin:0 0 10px;font-size:17px}#xed .xe-dlg-in p{color:var(--xmu);font-size:13px;line-height:1.5;margin:4px 0 0}
#xed .xe-dlg-in .acts{display:flex;gap:8px;justify-content:flex-end;margin-top:16px}
#xed .xe-dlg-in .acts button{border-radius:8px;padding:8px 16px;border:1px solid var(--xl);background:#fff;cursor:pointer}
#xed .xe-dlg-in .acts .pri{background:var(--xn);color:#fff;border-color:var(--xn)}#xed .xe-dlg-in .acts .warn{background:#C0392B;color:#fff;border-color:#C0392B}
#xed .xg-b:has(> .num){flex-direction:column;align-items:stretch;justify-content:center}
#xed .num{display:flex;align-items:center;gap:5px;font-size:12px;color:var(--xmu)}#xed .num input{width:56px;height:26px;border:1px solid #CBD3DF;border-radius:6px;padding:0 5px;font:inherit;color:var(--xt)}
#xed .xe-grid td.pbc{border-inline-end:1.5px dashed #2E58CC!important}#xed .xe-grid td.pbr{border-bottom:1.5px dashed #2E58CC!important}
#xed .xe-grid td.fz{background-color:#fff}#xed .xe-grid tr.fzr td.xin,#xed .xe-grid td.fz.xin{background-image:linear-gradient(rgba(33,115,70,.14),rgba(33,115,70,.14))}#xed .xe-grid tr.fzr td,#xed .xe-grid tr.fzr th{background-color:#fff}
#xed .xe-grid tr.fzr th{background:#F3F5F8}#xed .xe-grid td.fzl{box-shadow:inset -1px 0 0 #8E9AB3}#xed .xe-grid tr.fzlast td{border-bottom:1.5px solid #8E9AB3}
#xed .xe-grid td.fzedge{border-inline-end:1.5px solid #8E9AB3}
#xed #xePrint{display:none}
@media print{html>body>*:not(#xed){display:none!important}html>body>#xed{position:static!important;display:block!important;background:#fff!important;inset:auto!important}#xed>*:not(#xePrint){display:none!important}#xed #xePrint{display:block!important}
  #xePrint table{border-collapse:collapse;table-layout:fixed;font:11pt Arial,sans-serif;color:#000}#xePrint td{padding:0 3px;white-space:nowrap;overflow:hidden;vertical-align:bottom;line-height:1.2}
  #xePrint td.w{white-space:pre-wrap}#xePrint.pg td{border:1px solid #BFBFBF}#xePrint td>div{overflow:hidden}}
#xed .xe-load{position:absolute;inset:0;display:grid;place-items:center;color:var(--xmu);font-size:15px;background:#fff;z-index:6}
@media (max-width:760px){#xed .xe-sub,#xed .xe-state,#xed .xg-l{display:none}#xed .xe-tb{padding:6px 9px;font-size:13px}#xed #xeDesk{display:none}}`;

/* ------------------------------------------------------------ runtime */
let X=null;
const q=id=>X.root.querySelector('#'+id);
const cur=()=>X.st.sheets[X.si];
const colPx=(sh,c)=>{const ci=sh.cols.get(c);if(ci&&ci.hidden)return 0;const w=ci&&ci.w!=null?ci.w:sh.defW;return Math.max(4,Math.round(w*7+5));};
const rowPx=(sh,r)=>{const ri=sh.rows.get(r);if(ri&&ri.hidden)return 0;const h=ri&&ri.ht?ri.ht:sh.defH;return Math.max(4,Math.round(h*4/3));};
function geo(sh){
  const g=sh.g||(sh.g={});
  g.rows=Math.min(1048576,Math.max(g.rows||0,sh.maxR+40,60));g.cols=Math.min(16384,Math.max(g.cols||0,sh.maxC+6,26));
  const top=new Float64Array(g.rows+1);for(let r=0;r<g.rows;r++)top[r+1]=top[r]+rowPx(sh,r);g.top=top;
  g.cw=[];for(let c=0;c<g.cols;c++)g.cw.push(colPx(sh,c));return g;
}
const cellAt=(sh,r,c)=>sh.cells.get(key(r,c));
const valueOf=(si,r,c)=>{const v=X.calc.value(si,r,c);return v;};
function display(si,cell,r,c){
  if(!cell)return{t:'',num:false};
  if(cell.f&&X.st.sheets[si].formulas)return{t:'='+cell.f,num:false};
  let v=cell.f&&!cell.array?valueOf(si,r,c):(cell.t==='e'?ERR(cell.v):cell.v);
  if(v===null||v===undefined)return{t:'',num:false};
  if(isErr(v))return{t:v.e,num:false,err:true};
  if(typeof v==='boolean')return{t:v?'TRUE':'FALSE',num:false,bool:true};
  const d=descOf(X.st.st,cell.s);
  if(typeof v==='number'){let t=fmtValue(v,d.fmt==='@'?'General':d.fmt);return{t:String(t),num:true};}
  return{t:String(v),num:false};
}
const BSTY={thin:'1px solid',medium:'2px solid',thick:'3px solid',dashed:'1px dashed',dotted:'1px dotted',double:'3px double',hair:'1px dotted',mediumDashed:'2px dashed',dashDot:'1px dashed',mediumDashDot:'2px dashed',dashDotDot:'1px dotted',mediumDashDotDot:'2px dotted',slantDashDot:'2px dashed'};
function cssOf(s){
  const st=X.st.st;const c=st.css||(st.css=new Map());if(c.has(s))return c.get(s);
  const d=descOf(st,s),f=d.font;let td='';
  td+=`font-family:'${f.name.replace(/'/g,'')}',Arial,sans-serif;font-size:${f.sz}pt;`;if(f.b)td+='font-weight:700;';if(f.i)td+='font-style:italic;';
  const dec=[f.u&&'underline',f.strike&&'line-through'].filter(Boolean);if(dec.length)td+=`text-decoration:${dec.join(' ')}${f.u==='double'?' double':''};`;
  if(f.color&&f.color!=='000000')td+=`color:#${f.color};`;if(d.bg)td+=`background:#${d.bg};`;
  // Excel draws "left" and "right" borders mirrored in a right-to-left sheet
  for(const [side,cssSide] of [['top','top'],['bottom','bottom'],['left','inline-start'],['right','inline-end']]){const b=d.bd[side];if(b)td+=`border-${cssSide}:${BSTY[b.style]||'1px solid'} #${b.color};`;}
  td+=`vertical-align:${d.v==='top'?'top':d.v==='center'?'middle':'bottom'};`;
  const o={td,h:d.h,wrap:d.wrap||d.h==='justify',indent:d.indent,bg:!!d.bg};c.set(s,o);return o;
}
const escH=s=>String(s).replace(/[&<>]/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;'})[m]);
// the HTML for rows ra..rb; o: {lead, headH, left, fc, fr, brk, print}
function rowsHtml(sh,g,ra,rb,vis,o){
  let h='';const covered=new Map();
  for(const m of sh.merges)if(m.r2>=ra&&m.r1<=rb){const a=Math.max(m.r1,ra);for(let r=a;r<=Math.min(m.r2,rb);r++)for(let c=m.c1;c<=m.c2;c++)covered.set(key(r,c),r===a&&c===m.c1?m:true);}
  const lastVisFz=o.fc?vis.filter(c=>c<o.fc).pop():-1;
  for(let r=ra;r<=rb;r++){
    const rh=g.top[r+1]-g.top[r];if(!rh)continue;const fzr=r<o.fr;
    const tops=fzr?`position:sticky;top:${o.headH+g.top[r]}px;z-index:2;`:'';
    h+=`<tr style="height:${rh}px"${fzr?` class="fzr${r===o.fr-1?' fzlast':''}"`:''}>`;
    if(o.lead)h+=`<th data-row="${r}"${fzr?` style="${tops}z-index:4"`:''}>${r+1}</th>`;
    for(const c of vis){
      const cv=covered.get(key(r,c));if(cv===true)continue;
      let span='',mr=r,mc=c,lastC=c;
      if(cv){mr=cv.r1;mc=cv.c1;lastC=cv.c2;let rs=0;for(let x=Math.max(cv.r1,ra);x<=Math.min(cv.r2,rb);x++)if(rowPx(sh,x))rs++;const cs=vis.filter(x=>x>=cv.c1&&x<=cv.c2).length;span=`${rs>1?` rowspan="${rs}"`:''}${cs>1?` colspan="${cs}"`:''}`;}
      const cls=[];let st='';
      if(c<o.fc){cls.push('fz');st+=`position:sticky;inset-inline-start:${o.lead+o.left.get(c)}px;z-index:${fzr?3:1};`;if(c===lastVisFz)cls.push('fzedge');if(fzr)st+=`top:${o.headH+g.top[r]}px;`;}
      else if(fzr)st+=tops;
      if(o.brk){if(o.brk.cols.has(lastC))cls.push('pbc');if(o.brk.rows.has(cv?Math.min(cv.r2,rb):r))cls.push('pbr');}
      const da=o.print?'':` data-r="${r}" data-c="${c}"${cv?` data-mr="${mr}" data-mc="${mc}"`:''}`;
      const cell=cellAt(sh,mr,mc);
      if(!cell){h+=`<td${da}${span}${cls.length?` class="${cls.join(' ')}"`:''}${st?` style="${st}"`:''}></td>`;continue;}
      const ds=display(X.st.sheets.indexOf(sh),cell,mr,mc),css=cssOf(cell.s);
      const al=css.h&&css.h!=='general'?({left:'left',right:'right',center:'center',centerContinuous:'center',fill:'start',justify:'justify',distributed:'center'}[css.h]||'start'):(ds.num?(sh.rtl?'left':'right'):ds.bool||ds.err?'center':'start');
      const pad=css.indent?`padding-inline-start:${css.indent*9}px;`:'';
      if(ds.t||css.bg)cls.push('v');if(css.wrap)cls.push('w');
      h+=`<td${da}${span}${cls.length?` class="${cls.join(' ')}"`:''} style="${css.td}${st}"><div dir="${ds.num?'ltr':'auto'}" style="text-align:${al};${pad}">${escH(ds.t)}</div></td>`;
    }
    h+='</tr>';
  }
  return h;
}
const PAPER={1:['Letter',8.5,11],5:['Legal',8.5,14],8:['A3',11.69,16.54],9:['A4',8.27,11.69],11:['A5',5.83,8.27]};
function printArea(sh){return sh.printArea||{r1:0,c1:0,r2:Math.max(0,sh.maxR),c2:Math.max(0,sh.maxC)};}
function pageBreaks(sh){
  const g=geo(sh),p=sh.page;let [,w,h]=PAPER[p.paper]||PAPER[9];if(p.orient==='landscape')[w,h]=[h,w];
  const pw=(w-p.l-p.r)*96,ph=(h-p.t-p.b)*96;const A=printArea(sh);
  let tw=0;for(let c=A.c1;c<=A.c2;c++)tw+=c<g.cols?g.cw[c]:colPx(sh,c);const th=(A.r2+1<=g.rows?g.top[A.r2+1]:g.top[g.rows])-g.top[Math.min(A.r1,g.rows)];
  let scale=p.fit?Math.min(1,p.fw?pw/Math.max(1,tw):1,p.fh?ph/Math.max(1,th):1):(p.scale||100)/100;
  const ew=pw/scale,eh=ph/scale;const cols=new Set(),rows=new Set();
  let acc=0;for(let c=A.c1;c<=A.c2;c++){const x=colPx(sh,c);if(acc+x>ew&&acc>0){cols.add(c-1);acc=0;}acc+=x;}cols.add(A.c2);
  acc=0;for(let r=A.r1;r<=A.r2;r++){const x=rowPx(sh,r);if(acc+x>eh&&acc>0){rows.add(r-1);acc=0;}acc+=x;}rows.add(A.r2);
  return{cols,rows,scale,pages:cols.size*rows.size,w,h};
}
function render(){
  const sh=cur(),g=geo(sh),wrap=q('xeWrap'),z=X.zoom;
  const lead=sh.headings!==false?46:0,headH=lead?22:0;
  const fr=sh.freeze?Math.min(sh.freeze.r,g.rows):0,fc=sh.freeze?sh.freeze.c:0;
  const pos=wrap.scrollTop/z+g.top[fr];
  let lo=fr,hi=g.rows;while(lo<hi){const m=(lo+hi)>>1;if(g.top[m+1]<=pos)lo=m+1;else hi=m;}const r0=Math.max(fr,lo-8);
  let r1=r0;const vh=wrap.clientHeight/z;while(r1<g.rows&&g.top[r1]<pos+vh+200)r1++;r1=Math.min(g.rows-1,r1+8);
  X.win={r0,r1};
  const vis=[];for(let c=0;c<g.cols;c++)if(g.cw[c]>4||!(sh.cols.get(c)||{}).hidden)vis.push(c);
  const left=new Map();let acc=0;for(const c of vis){if(c>=fc)break;left.set(c,acc);acc+=g.cw[c];}
  const brk=sh.breaks?pageBreaks(sh):null;X.brk=brk;
  const o={lead,headH,left,fc,fr,brk};
  let h=`<colgroup>${lead?'<col style="width:46px">':''}${vis.map(c=>`<col style="width:${g.cw[c]}px">`).join('')}</colgroup>`;
  if(lead)h+=`<thead><tr><th class="corner" data-all="1"></th>${vis.map(c=>`<th data-col="${c}"${c<fc?` style="inset-inline-start:${lead+left.get(c)}px;z-index:5"`:''}>${colName(c)}<i class="rz" data-rz="${c}"></i></th>`).join('')}</tr></thead>`;
  h+='<tbody>';
  if(fr)h+=rowsHtml(sh,g,0,fr-1,vis,o);
  const gap=g.top[r0]-g.top[fr];if(gap>0)h+=`<tr class="sp"><td colspan="${vis.length+1}" style="height:${gap}px"></td></tr>`;
  h+=rowsHtml(sh,g,r0,r1,vis,o);
  const rest=g.top[g.rows]-g.top[r1+1];if(rest>0)h+=`<tr class="sp"><td colspan="${vis.length+1}" style="height:${rest}px"></td></tr>`;
  h+='</tbody>';
  const grid=q('xeGrid');grid.innerHTML=h;grid.style.direction=sh.rtl?'rtl':'ltr';grid.classList.toggle('nogrid',!sh.grid);grid.style.zoom=z;
  paintSel();
}
let rT=0;const renderSoon=()=>{cancelAnimationFrame(rT);rT=requestAnimationFrame(()=>{if(X)render();});};
function paintSel(){
  const s=X.sel,a=X.act;const grid=q('xeGrid');
  for(const td of grid.querySelectorAll('td[data-r]')){
    const r=+(td.dataset.mr??td.dataset.r),c=+(td.dataset.mc??td.dataset.c);
    td.classList.toggle('xin',r>=s.r1&&r<=s.r2&&c>=s.c1&&c<=s.c2&&!(r===a.r&&c===a.c));
    td.classList.toggle('xact',r===a.r&&c===a.c);
  }
  for(const th of grid.querySelectorAll('th[data-col]')){const c=+th.dataset.col;th.classList.toggle('on',c>=s.c1&&c<=s.c2);}
  for(const th of grid.querySelectorAll('th[data-row]')){const r=+th.dataset.row;th.classList.toggle('on',r>=s.r1&&r<=s.r2);}
  updateUI();
}
function tdOf(r,c){return q('xeGrid').querySelector(`td[data-r="${r}"][data-c="${c}"]`)||q('xeGrid').querySelector(`td[data-mr="${r}"][data-mc="${c}"]`);}
function mergeAt(sh,r,c){return sh.merges.find(m=>r>=m.r1&&r<=m.r2&&c>=m.c1&&c<=m.c2)||null;}
function select(r,c,extend){
  const sh=cur();r=Math.max(0,Math.min(1048575,r));c=Math.max(0,Math.min(16383,c));
  const g=sh.g;if(g&&(r>=g.rows-5||c>=g.cols-3)){g.rows=Math.max(g.rows,r+40);g.cols=Math.max(g.cols,c+6);}
  if(extend){const a=X.anchor||X.act;X.sel={r1:Math.min(a.r,r),c1:Math.min(a.c,c),r2:Math.max(a.r,r),c2:Math.max(a.c,c)};X.cursor={r,c};}
  else{const m=mergeAt(sh,r,c);if(m){r=m.r1;c=m.c1;X.sel={...m};}else X.sel={r1:r,c1:c,r2:r,c2:c};X.act={r,c};X.anchor={r,c};X.cursor={r,c};}
  // a selection always covers whole merged blocks
  for(let i=0;i<3;i++)for(const m of sh.merges)if(m.r2>=X.sel.r1&&m.r1<=X.sel.r2&&m.c2>=X.sel.c1&&m.c1<=X.sel.c2){X.sel.r1=Math.min(X.sel.r1,m.r1);X.sel.c1=Math.min(X.sel.c1,m.c1);X.sel.r2=Math.max(X.sel.r2,m.r2);X.sel.c2=Math.max(X.sel.c2,m.c2);}
  scrollToCell(X.cursor.r,X.cursor.c);paintSel();
}
function scrollToCell(r,c){
  const sh=cur(),g=geo(sh),wrap=q('xeWrap'),z=X.zoom;
  const fr=sh.freeze?sh.freeze.r:0,fc=sh.freeze?sh.freeze.c:0;
  const head=((sh.headings!==false?22:0)+g.top[Math.min(fr,g.rows)])*z;
  if(r>=fr){const y=(g.top[r]-g.top[Math.min(fr,g.rows)])*z,h=(g.top[r+1]-g.top[r])*z;
    if(y<wrap.scrollTop)wrap.scrollTop=y;else if(y+h+head>wrap.scrollTop+wrap.clientHeight)wrap.scrollTop=y+h+head-wrap.clientHeight+4;}
  render();const td=c>=fc?tdOf(r,c):null;if(td){const tr=td.getBoundingClientRect(),wr=wrap.getBoundingClientRect();
    let lead=(sh.headings!==false?46:0)*z;for(let x=0;x<fc;x++)lead+=colPx(sh,x)*z;
    if(sh.rtl){if(tr.right>wr.right-lead)wrap.scrollLeft+=tr.right-(wr.right-lead);else if(tr.left<wr.left)wrap.scrollLeft-=wr.left-tr.left+(wrap.offsetWidth-wrap.clientWidth);}
    else{if(tr.left<wr.left+lead)wrap.scrollLeft-=wr.left+lead-tr.left;else if(tr.right>wr.right)wrap.scrollLeft+=tr.right-wr.right+(wrap.offsetWidth-wrap.clientWidth);}}
}

/* ---------- history ---------- */
const cloneCell=c=>{const n={...c};delete n.ast;return n;};
function snapSheets(all){
  const list=all?X.st.sheets.map((_,i)=>i):[X.si];
  return list.map(i=>{const s=X.st.sheets[i];return{i,cells:new Map([...s.cells].map(([k,v])=>[k,cloneCell(v)])),rows:new Map([...s.rows].map(([k,v])=>[k,{...v,attrs:{...v.attrs}}])),
    cols:new Map([...s.cols].map(([k,v])=>[k,{...v}])),merges:s.merges.map(m=>({...m})),maxR:s.maxR,maxC:s.maxC,dirty:s.dirty,doc:all?ser(s.doc):null};});
}
function restore(snap){for(const s of snap){const sh=X.st.sheets[s.i];Object.assign(sh,{cells:s.cells,rows:s.rows,cols:s.cols,merges:s.merges,maxR:s.maxR,maxC:s.maxC,dirty:true});if(s.doc){sh.doc=parseXml(s.doc);sh.root=sh.doc.documentElement;}}}
function change(fn,all){
  const before=snapSheets(all);fn();const after=snapSheets(all);
  X.hist.splice(X.hi+1);X.hist.push({si:X.si,before,after});if(X.hist.length>60)X.hist.shift();X.hi=X.hist.length-1;
  for(const s of after)X.st.sheets[s.i].dirty=true;
  X.calc.reset();X.dirty=true;markDirty();render();
}
function undo(){if(X.hi<0)return hint('אין עוד מה לבטל');const h=X.hist[X.hi--];restore(h.before.map(s=>({...s,cells:new Map([...s.cells].map(([k,v])=>[k,cloneCell(v)]))})));X.si=h.si;X.calc.reset();X.dirty=true;markDirty();buildTabs();render();}
function redo(){if(X.hi>=X.hist.length-1)return;const h=X.hist[++X.hi];restore(h.after.map(s=>({...s,cells:new Map([...s.cells].map(([k,v])=>[k,cloneCell(v)]))})));X.si=h.si;X.calc.reset();X.dirty=true;markDirty();buildTabs();render();}
function markDirty(){const s=q('xeState');s.textContent=X.dirty?'שינויים שלא נשמרו':(X.savedOnce?'נשמר':'');s.classList.toggle('dirty',X.dirty);}
function hint(m){X.opts.toast(m);}

/* ---------- values ---------- */
function parseInput(s,curFmt){
  if(s===''||s===null)return{empty:true};
  if(s[0]==='='&&s.length>1)return{f:s.slice(1)};
  if(s[0]==="'")return{v:s.slice(1),t:'s'};
  const t=s.trim();const gen=!curFmt||curFmt==='General';
  if(/^(true|false)$/i.test(t))return{v:/^true$/i.test(t),t:'b'};
  let m;
  if((m=/^([-+]?)([₪$€])?\s?((?:\d{1,3}(?:,\d{3})+|\d+)?(?:\.\d+)?)\s?([₪$€])?$/.exec(t))&&/\d/.test(m[3])&&(m[2]||m[4]||/^[-+]?(\d|\.\d)/.test(t))){
    const n=+(m[1]+m[3].replace(/,/g,''));const cur=m[2]||m[4];
    return{v:n,t:'n',fmt:gen?(cur?(cur==='₪'?'[$₪-40D] #,##0.00':`"${cur}"#,##0.00`):(/,/.test(m[3])?(/\./.test(m[3])?'#,##0.00':'#,##0'):undefined)):undefined};}
  if((m=/^([-+]?(?:\d+\.?\d*|\.\d+))\s?%$/.exec(t)))return{v:+m[1]/100,t:'n',fmt:gen?(/\./.test(m[1])?'0.00%':'0%'):undefined};
  if((m=/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2}|\d{4})$/.exec(t))){let y=+m[3];if(y<100)y+=y<30?2000:1900;const d=+m[1],mo=+m[2];
    if(mo>=1&&mo<=12&&d>=1&&d<=31)return{v:Date.UTC(y,mo-1,d)/864e5+25569,t:'n',fmt:gen||!isDateFmt(curFmt)?'dd/mm/yyyy':undefined};}
  if((m=/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(t)))return{v:(+m[1]*3600+ +m[2]*60+(+m[3]||0))/86400,t:'n',fmt:gen?(m[3]?'h:mm:ss':'h:mm'):undefined};
  return{v:s,t:'s'};
}
function rawText(cell){
  if(!cell)return'';if(cell.f)return'='+cell.f;
  if(cell.v===null||cell.v===undefined)return'';if(typeof cell.v==='boolean')return cell.v?'TRUE':'FALSE';
  if(cell.t==='n'&&typeof cell.v==='number'){const d=descOf(X.st.st,cell.s);if(isDateFmt(d.fmt))return fmtValue(cell.v,/[hs]/i.test(d.fmt.replace(/"[^"]*"/g,''))&&!/[dy]/i.test(d.fmt)?d.fmt:(/[hs]/i.test(d.fmt)?'dd/mm/yyyy h:mm':'dd/mm/yyyy'));
    if(/%/.test(d.fmt))return fmtGeneral(+(cell.v*100).toPrecision(12))+'%';return fmtGeneral(cell.v);}
  return String(cell.v);
}
function putValue(sh,r,c,s){
  let cell=cellAt(sh,r,c);const fmt=cell?descOf(X.st.st,cell.s).fmt:'General';const p=parseInput(s,fmt);
  if(p.empty){if(!cell)return;cell.v=null;cell.f=null;cell.t='n';cell.mod=true;delete cell.ast;if(!cell.s)sh.cells.delete(key(r,c));return;}
  if(!cell){cell={r,c,s:(sh.cols.get(c)||{}).style?+(sh.cols.get(c)).style:0,t:'n',v:null,f:null};sh.cells.set(key(r,c),cell);}
  cell.mod=true;delete cell.ast;delete cell.array;delete cell.sidx;
  if(p.f!==undefined){cell.f=p.f.replace(/^([A-Za-z.]+)\(/,(m,n)=>n.toUpperCase()+'(');cell.v=null;cell.t='n';X.formulaTouched=true;}
  else{cell.f=null;cell.v=p.v;cell.t=p.t;if(p.fmt){cell.s=newXf(X.st.st,cell.s,{fmt:p.fmt});X.styleDirty=true;}}
  if(r>sh.maxR)sh.maxR=r;if(c>sh.maxC)sh.maxC=c;
}
function setInput(r,c,s){change(()=>putValue(cur(),r,c,s));}
function clearCells(){change(()=>{const sh=cur(),s=X.sel;for(const cell of [...sh.cells.values()])if(cell.r>=s.r1&&cell.r<=s.r2&&cell.c>=s.c1&&cell.c<=s.c2){cell.v=null;cell.f=null;cell.t='n';cell.mod=true;delete cell.ast;if(!cell.s)sh.cells.delete(key(cell.r,cell.c));}});}

/* ---------- formatting ---------- */
function eachSelCell(fn,create){
  const sh=cur(),s=X.sel;const n=(s.r2-s.r1+1)*(s.c2-s.c1+1);
  if(create&&n<=60000){for(let r=s.r1;r<=s.r2;r++)for(let c=s.c1;c<=s.c2;c++){let cell=cellAt(sh,r,c);if(!cell){cell={r,c,s:0,t:'n',v:null,f:null,mod:true};sh.cells.set(key(r,c),cell);}fn(cell);}}
  else for(const cell of sh.cells.values())if(cell.r>=s.r1&&cell.r<=s.r2&&cell.c>=s.c1&&cell.c<=s.c2)fn(cell);
}
function style(ch){change(()=>{eachSelCell(cell=>{cell.s=newXf(X.st.st,cell.s,ch);cell.mod=true;},true);X.styleDirty=true;X.st.st.css=null;});}
const actDesc=()=>{const c=cellAt(cur(),X.act.r,X.act.c);return descOf(X.st.st,c?c.s:0);};
function toggleFont(k){const d=actDesc();const on=k==='u'?!!d.font.u:!!d.font[k];style({font:{[k]:on?(k==='u'?null:false):(k==='u'?'single':true)}});}
function borders(kind){
  const sh=cur(),s=X.sel;const T={style:'thin',color:'000000'},K={style:'medium',color:'000000'};
  change(()=>{eachSelCell(cell=>{
    const top=cell.r===s.r1,bot=cell.r===s.r2,lft=cell.c===s.c1,rgt=cell.c===s.c2;let bd=null;
    if(kind==='all')bd={top:T,bottom:T,left:T,right:T};
    else if(kind==='none')bd={top:null,bottom:null,left:null,right:null};
    else if(kind==='outside'||kind==='thick'){const B2=kind==='thick'?K:T;bd={};if(top)bd.top=B2;if(bot)bd.bottom=B2;if(lft)bd.left=B2;if(rgt)bd.right=B2;if(!Object.keys(bd).length)return;}
    else if(kind==='bottom'){if(!bot)return;bd={bottom:T};}else if(kind==='top'){if(!top)return;bd={top:T};}
    else if(kind==='dbottom'){if(!bot)return;bd={bottom:{style:'double',color:'000000'}};}
    cell.s=newXf(X.st.st,cell.s,{bd});cell.mod=true;},true);X.styleDirty=true;X.st.st.css=null;});
}
function setFmt(code){style({fmt:code});}
function decimals(d){
  const f=actDesc().fmt;let code=f;
  if(f==='General'||!/0/.test(f)){const c=cellAt(cur(),X.act.r,X.act.c);const v=c&&typeof c.v==='number'?c.v:0;const dec=(String(v).split('.')[1]||'').length;code=dec>0?'0.'+'0'.repeat(Math.max(0,dec+d)):'0'+(d>0?'.0':'');}
  else code=f.split(';').map(part=>part.replace(/0(\.0+)?(?![0#])/,(m,dp)=>{const n=(dp?dp.length-1:0)+d;return n>0?'0.'+'0'.repeat(n):'0';})).join(';');
  setFmt(code.replace(/\.$/,''));
}
function toggleMerge(){
  const sh=cur(),s=X.sel;const m=mergeAt(sh,X.act.r,X.act.c);
  if(m&&m.r1===s.r1&&m.c1===s.c1&&m.r2===s.r2&&m.c2===s.c2){change(()=>{sh.merges=sh.merges.filter(x=>x!==m);});render();return;}
  if(s.r1===s.r2&&s.c1===s.c2)return;
  let lost=0;for(const c of sh.cells.values())if(c.r>=s.r1&&c.r<=s.r2&&c.c>=s.c1&&c.c<=s.c2&&!(c.r===s.r1&&c.c===s.c1)&&(c.v!==null||c.f))lost++;
  const go=()=>change(()=>{
    for(const c of [...sh.cells.values()])if(c.r>=s.r1&&c.r<=s.r2&&c.c>=s.c1&&c.c<=s.c2&&!(c.r===s.r1&&c.c===s.c1)){c.v=null;c.f=null;c.mod=true;}
    sh.merges=sh.merges.filter(x=>!(x.r2>=s.r1&&x.r1<=s.r2&&x.c2>=s.c1&&x.c1<=s.c2));sh.merges.push({...s});
    let tl=cellAt(sh,s.r1,s.c1);if(!tl){tl={r:s.r1,c:s.c1,s:0,t:'n',v:null,f:null};sh.cells.set(key(s.r1,s.c1),tl);}
    tl.s=newXf(X.st.st,tl.s,{al:{h:'center',v:'center'}});tl.mod=true;X.styleDirty=true;X.st.st.css=null;});
  if(lost)ask('מיזוג תאים','במיזוג נשאר רק הערך של התא הראשון. שאר הערכים יימחקו.',[{v:'x',l:'ביטול'},{v:'ok',l:'מיזוג',cls:'pri'}]).then(a=>{if(a==='ok')go();});else go();
}

/* ---------- rows and columns ---------- */
function structural(axis,at,n){
  const sh=cur(),name=sh.name;
  change(()=>{
    const nm=new Map();
    for(const c of sh.cells.values()){const p=axis==='r'?c.r:c.c;
      if(n<0&&p>=at&&p<at-n)continue;
      if(p>=at){if(axis==='r')c.r+=n;else c.c+=n;c.mod=true;}
      nm.set(key(c.r,c.c),c);}
    sh.cells=nm;
    for(const s2 of X.st.sheets)for(const c of s2.cells.values())if(c.f){const nf=shiftFormula(c.f,name,s2===sh,axis,at,n);if(nf!==c.f){c.f=nf;delete c.ast;c.mod=true;s2.dirty=true;}}
    const shiftMap=m=>{const o=new Map();for(const [k,v] of m){if(n<0&&k>=at&&k<at-n)continue;o.set(k>=at?k+n:k,v);}return o;};
    if(axis==='r')sh.rows=shiftMap(sh.rows);else sh.cols=shiftMap(sh.cols);
    sh.merges=sh.merges.map(m=>parseRange(shiftRangeStr(rangeStr(m),axis,at,n))).filter(m=>m&&(m.r1!==m.r2||m.c1!==m.c2));
    const R=sh.root;
    for(const tag of ['conditionalFormatting','dataValidation']){for(const e of [...R.getElementsByTagNameNS(S,tag)]){const v=shiftRangeStr(e.getAttribute('sqref')||'',axis,at,n);if(v)e.setAttribute('sqref',v);else e.parentNode.removeChild(e);}}
    for(const e of R.getElementsByTagNameNS(S,'formula'))e.textContent=shiftFormula(e.textContent,name,true,axis,at,n);
    for(const tag of ['hyperlink','autoFilter']){for(const e of [...R.getElementsByTagNameNS(S,tag)]){const v=shiftRangeStr(e.getAttribute('ref')||'',axis,at,n);if(v)e.setAttribute('ref',v);else e.parentNode.removeChild(e);}}
    for(const dn of X.st.wbDoc.getElementsByTagNameNS(S,'definedName')){const nf=shiftFormula(dn.textContent,name,false,axis,at,n);if(nf!==dn.textContent){dn.textContent=nf;X.wbDirty=true;}}
    if(axis==='r')sh.maxR=Math.max(0,sh.maxR+n);else sh.maxC=Math.max(0,sh.maxC+n);
    X.structChanged=true;if(sh.g){sh.g.rows=0;sh.g.cols=0;}
  },true);
  select(X.act.r,X.act.c);
}
const insRows=()=>structural('r',X.sel.r1,X.sel.r2-X.sel.r1+1);
const insCols=()=>structural('c',X.sel.c1,X.sel.c2-X.sel.c1+1);
const delRows=()=>structural('r',X.sel.r1,-(X.sel.r2-X.sel.r1+1));
const delCols=()=>structural('c',X.sel.c1,-(X.sel.c2-X.sel.c1+1));
function sortSel(desc){
  const sh=cur();let s={...X.sel};
  if(s.r1===s.r2){ // the block around the active cell
    let a=X.act.r,b=X.act.r;const has=r=>{const c=cellAt(sh,r,X.act.c);return c&&(c.v!==null||c.f);};
    while(a>0&&has(a-1))a--;while(has(b+1))b++;
    s={r1:a,r2:b,c1:0,c2:sh.maxC};
    const h=cellAt(sh,a,X.act.c),n=cellAt(sh,a+1,X.act.c);if(h&&typeof h.v==='string'&&n&&typeof n.v==='number')s.r1=a+1;
  }
  if(s.r2<=s.r1)return hint('סמנו כמה שורות למיון');
  if(sh.merges.some(m=>m.r2>=s.r1&&m.r1<=s.r2&&m.c2>=s.c1&&m.c1<=s.c2))return hint('אי אפשר למיין תאים ממוזגים');
  change(()=>{
    const rows=[];for(let r=s.r1;r<=s.r2;r++){const cells=[];for(let c=s.c1;c<=s.c2;c++){const x=cellAt(sh,r,c);if(x)cells.push(x);sh.cells.delete(key(r,c));}rows.push({cells,k:valueOf(X.si,r,X.act.c)});}
    X.calc.reset();
    const rank=v=>v===null||v===undefined||v===''?2:typeof v==='number'?0:1;
    rows.sort((x,y)=>{const a=x.k,b=y.k;const ra=rank(a),rb=rank(b);if(ra!==rb)return ra-rb;if(ra===2)return 0;const d=ra===0?a-b:String(a).localeCompare(String(b),'he');return desc?-d:d;});
    rows.forEach((row,i)=>{for(const c of row.cells){c.r=s.r1+i;c.mod=true;delete c.ast;sh.cells.set(key(c.r,c.c),c);}});
  });
}
function autoSum(){
  const sh=cur(),{r,c}=X.act;const num=(rr,cc)=>{const x=cellAt(sh,rr,cc);return x&&(typeof x.v==='number'||x.f)&&!(x.f&&/^SUM\(/i.test(x.f)&&false);};
  if(X.sel.r2>X.sel.r1||X.sel.c2>X.sel.c1){const s=X.sel;change(()=>{for(let cc=s.c1;cc<=s.c2;cc++)putValue(sh,s.r2+1,cc,`=SUM(${refOf(s.r1,cc)}:${refOf(s.r2,cc)})`);});select(s.r2+1,s.c1);return;}
  let a=r-1;while(a>=0&&num(a,c))a--;
  if(a<r-1){setInput(r,c,`=SUM(${refOf(a+1,c)}:${refOf(r-1,c)})`);select(r+1,c);return;}
  let b=c-1;while(b>=0&&num(r,b))b--;
  if(b<c-1){setInput(r,c,`=SUM(${refOf(r,b+1)}:${refOf(r,c-1)})`);select(r,c+1);return;}
  startEdit('=SUM()',true);const e=q('xeEdit');e.selectionStart=e.selectionEnd=5;
}

/* ---------- clipboard ---------- */
function selTsv(){
  const s=X.sel;const lines=[];
  for(let r=s.r1;r<=s.r2;r++){const row=[];for(let c=s.c1;c<=s.c2;c++){const t=display(X.si,cellAt(cur(),r,c),r,c).t;row.push(/[\t\n"]/.test(t)?'"'+t.replace(/"/g,'""')+'"':t);}lines.push(row.join('\t'));}
  return lines.join('\r\n');
}
function doCopy(cut){
  const s={...X.sel},sh=cur();const cells=[];
  for(const c of sh.cells.values())if(c.r>=s.r1&&c.r<=s.r2&&c.c>=s.c1&&c.c<=s.c2)cells.push(cloneCell(c));
  const text=selTsv();X.clip={si:X.si,sel:s,cells,text,cut,merges:sh.merges.filter(m=>m.r1>=s.r1&&m.r2<=s.r2&&m.c1>=s.c1&&m.c2<=s.c2).map(m=>({...m}))};
  return text;
}
function parseTsv(t){
  const rows=[];let row=[],cell='',i=0,q2=false;t=t.replace(/\r\n?/g,'\n');
  while(i<t.length){const ch=t[i];
    if(q2){if(ch==='"'){if(t[i+1]==='"'){cell+='"';i++;}else q2=false;}else cell+=ch;}
    else if(ch==='"'&&cell==='')q2=true;else if(ch==='\t'){row.push(cell);cell='';}else if(ch==='\n'){row.push(cell);rows.push(row);row=[];cell='';}else cell+=ch;
    i++;}
  if(cell!==''||row.length)row.push(cell),rows.push(row);
  return rows;
}
function doPaste(text){
  const sh=cur(),{r:r0,c:c0}=X.act;const cl=X.clip;
  if(cl&&cl.text===text){
    const dr=r0-cl.sel.r1,dc=c0-cl.sel.c1;const src=X.st.sheets[cl.si];
    change(()=>{
      if(cl.cut)for(const c of cl.cells){const k=key(c.r,c.c);if(src.cells.get(k)){src.cells.delete(k);}src.dirty=true;}
      for(let r=cl.sel.r1;r<=cl.sel.r2;r++)for(let c=cl.sel.c1;c<=cl.sel.c2;c++)sh.cells.delete(key(r+dr,c+dc));
      for(const c0_ of cl.cells){const n=cloneCell(c0_);n.r+=dr;n.c+=dc;n.mod=true;delete n.el;delete n.sidx;
        if(n.f&&!cl.cut)n.f=moveFormula(n.f,dr,dc);sh.cells.set(key(n.r,n.c),n);sh.maxR=Math.max(sh.maxR,n.r);sh.maxC=Math.max(sh.maxC,n.c);}
      for(const m of cl.merges){const nm={r1:m.r1+dr,r2:m.r2+dr,c1:m.c1+dc,c2:m.c2+dc};sh.merges=sh.merges.filter(x=>!(x.r2>=nm.r1&&x.r1<=nm.r2&&x.c2>=nm.c1&&x.c1<=nm.c2));sh.merges.push(nm);}
      if(cl.cut&&src!==sh)src.merges=src.merges.filter(m=>!cl.merges.some(x=>x.r1===m.r1&&x.c1===m.c1));
    },cl.si!==X.si);
    if(cl.cut)X.clip=null;
    X.sel={r1:r0,c1:c0,r2:r0+cl.sel.r2-cl.sel.r1,c2:c0+cl.sel.c2-cl.sel.c1};paintSel();return;
  }
  const rows=parseTsv(text);if(!rows.length)return;
  change(()=>{rows.forEach((row,i)=>row.forEach((v,j)=>putValue(sh,r0+i,c0+j,v)));});
  X.sel={r1:r0,c1:c0,r2:r0+rows.length-1,c2:c0+Math.max(...rows.map(r=>r.length))-1};paintSel();
}

/* ---------- editing a cell ---------- */
function startEdit(text,typed){
  const{r,c}=X.act;const td=tdOf(r,c);if(!td){scrollToCell(r,c);}
  const t2=tdOf(r,c);if(!t2)return;
  const rc=t2.getBoundingClientRect(),e=q('xeEdit');const cell=cellAt(cur(),r,c);const d=descOf(X.st.st,cell?cell.s:0);
  X.editing={r,c,typed:!!typed,orig:rawText(cell)};
  Object.assign(e.style,{position:'fixed',left:rc.left+'px',top:rc.top+'px',minWidth:Math.max(rc.width,60)+'px',width:Math.max(rc.width,60)+'px',height:Math.max(rc.height,22)+'px',
    fontFamily:`'${d.font.name}',Arial,sans-serif`,fontSize:(d.font.sz*X.zoom)+'pt',fontWeight:d.font.b?'700':'400',direction:cur().rtl?'rtl':'ltr',textAlign:'start'});
  e.hidden=false;e.value=text===undefined?X.editing.orig:text;e.focus();const n=e.value.length;e.setSelectionRange(n,n);
  q('xeFx').value=e.value;growEdit();
}
function growEdit(){const e=q('xeEdit');if(e.hidden)return;e.style.width='1px';e.style.width=Math.max(parseFloat(e.style.minWidth),e.scrollWidth+8)+'px';e.style.height='auto';e.style.height=Math.max(22,e.scrollHeight)+'px';}
function commitEdit(move){
  const ed=X.editing;if(!ed)return false;const e=q('xeEdit');const v=e.hidden?q('xeFx').value:e.value;
  X.editing=null;e.hidden=true;
  if(v!==ed.orig)setInput(ed.r,ed.c,v);
  q('xeWrap').focus({preventScroll:true});
  if(move)moveAct(...move);else paintSel();
  return true;
}
function cancelEdit(){if(!X.editing)return;X.editing=null;q('xeEdit').hidden=true;q('xeWrap').focus({preventScroll:true});paintSel();}
function moveAct(dr,dc,extend){
  const sh=cur();
  if(extend){const cu=X.cursor||X.act;select(cu.r+dr,cu.c+dc,true);return;}
  let{r,c}=X.act;const m=mergeAt(sh,r,c);
  if(m){if(dr>0)r=m.r2;if(dc>0)c=m.c2;}
  select(r+dr,c+dc);
}
function jump(dr,dc,extend){
  const sh=cur();const cu=extend?(X.cursor||X.act):X.act;let r=cu.r,c=cu.c;
  const has=(a,b)=>{const x=cellAt(sh,a,b);return!!x&&(x.v!==null&&x.v!==''||!!x.f);};
  const lim=dr?Math.max(sh.maxR,r):Math.max(sh.maxC,c);
  const step=()=>{r+=dr;c+=dc;};
  if(has(r,c)&&has(r+dr,c+dc)){while(r+dr>=0&&c+dc>=0&&has(r+dr,c+dc))step();}
  else{step();while(r>=0&&c>=0&&(dr?r:c)<=lim&&!has(r,c))step();if((dr?r:c)>lim||r<0||c<0){r=Math.max(0,Math.min(r,dr?lim:r));c=Math.max(0,Math.min(c,dc?lim:c));}}
  select(Math.max(0,r),Math.max(0,c),extend);
}

/* ---------- find ---------- */
function findNext(){
  const qs=q('xeFindIn').value.toLowerCase();if(!qs)return;const sh=cur();
  const list=[...sh.cells.values()].sort((a,b)=>a.r-b.r||a.c-b.c);const{r,c}=X.act;
  const hit=list.find(x=>(x.r>r||(x.r===r&&x.c>c))&&display(X.si,x,x.r,x.c).t.toLowerCase().includes(qs))||list.find(x=>display(X.si,x,x.r,x.c).t.toLowerCase().includes(qs));
  if(hit)select(hit.r,hit.c);else hint('לא נמצא');
}

/* ---------- pop-up menus and dialogs ---------- */
function pop(anchor,html,onPick){
  const p=q('xePop');p.innerHTML=html;p.hidden=false;X.popFor=anchor;const a=anchor.getBoundingClientRect();
  let x=a.right-p.offsetWidth;if(x<6)x=6;if(x+p.offsetWidth>innerWidth-6)x=innerWidth-6-p.offsetWidth;let y=a.bottom+4;if(y+p.offsetHeight>innerHeight-6)y=Math.max(6,a.top-p.offsetHeight-4);
  p.style.left=x+'px';p.style.top=y+'px';p.onmousedown=e=>{if(!e.target.closest('input'))e.preventDefault();};
  p.onclick=e=>{const b=e.target.closest('[data-v]');if(!b)return;closePop();onPick(b.dataset.v);};return p;
}
function closePop(){const p=q('xePop');if(p.hidden)return;p.hidden=true;p.innerHTML='';X.popFor=null;}
const pal=(list,none)=>`<div class="pal">${list.map(c=>`<button type="button" data-v="${c}" title="#${c}" style="background:#${c}"></button>`).join('')}</div>${none?`<div class="sep"></div><button type="button" class="mi" data-v="none">${none}</button>`:''}`;
function setColor(h){X.color=h;q('xeRibbon').querySelector('[data-bar=color]').style.background='#'+(h||'000000');style({font:{color:h||'000000'}});}
function setFill(h){X.fill=h;q('xeRibbon').querySelector('[data-bar=fill]').style.background=h?'#'+h:'transparent';style({bg:h});}
function openPop(kind,btn){
  if(X.popFor===btn){closePop();return;}
  if(kind==='color')pop(btn,`<h4>צבע גופן</h4>${pal(COLORS,'אוטומטי (שחור)')}`,v=>setColor(v==='none'?null:v));
  else if(kind==='fill')pop(btn,`<h4>צבע מילוי</h4>${pal(FILLS,'ללא מילוי')}`,v=>setFill(v==='none'?null:v));
  else if(kind==='margins'){const M=[['normal','רגילים','1.91 למעלה ולמטה, 1.78 בצדדים',.75,.7],['wide','רחבים','2.54 מכל צד',1,1],['narrow','צרים','1.91 למעלה ולמטה, 0.64 בצדדים',.75,.25]];
    pop(btn,M.map(m=>`<button type="button" class="mi" data-v="${m[0]}"><b>${m[1]}</b><small style="margin-inline-start:auto;color:#5D6982">${m[2]} ס״מ</small></button>`).join('')+'<div class="sep"></div><button type="button" class="mi" data-v="custom">שוליים מותאמים…</button>',v=>{
      if(v==='custom')return marginsDialog();const m=M.find(x=>x[0]===v);setPage({t:m[3],b:m[3],l:m[4],r:m[4]});});}
  else if(kind==='orient'){const l=cur().page.orient==='landscape';pop(btn,`<button type="button" class="mi"${!l?' style="background:#EEF2F8"':''} data-v="portrait">לאורך</button><button type="button" class="mi"${l?' style="background:#EEF2F8"':''} data-v="landscape">לרוחב</button>`,v=>setPage({orient:v}));}
  else if(kind==='paper')pop(btn,[9,1,5,11,8].map(k=>`<button type="button" class="mi" data-v="${k}"><b>${PAPER[k][0]}</b><small style="margin-inline-start:auto;color:#5D6982">${(PAPER[k][1]*2.54).toFixed(1)} × ${(PAPER[k][2]*2.54).toFixed(1)} ס״מ</small></button>`).join(''),v=>setPage({paper:+v}));
  else if(kind==='fit')pop(btn,[['none','בלי התאמה','מודפס בגודל הרגיל'],['w','כל העמודות בעמוד אחד','לרוחב עמוד אחד'],['all','כל הגיליון בעמוד אחד','מוקטן כדי להיכנס לעמוד']].map(x=>`<button type="button" class="mi" data-v="${x[0]}"><b>${x[1]}</b><small style="margin-inline-start:auto;color:#5D6982">${x[2]}</small></button>`).join(''),v=>{
    if(v==='none')setPage({fit:false,scale:100});else if(v==='w')setPage({fit:true,fw:1,fh:0});else setPage({fit:true,fw:1,fh:1});});
  else if(kind==='freeze')pop(btn,[['top','הקפאת השורה העליונה'],['first','הקפאת העמודה הראשונה'],['here','הקפאה עד התא הנבחר'],['none','ביטול ההקפאה']].map(x=>`<button type="button" class="mi" data-v="${x[0]}">${x[1]}</button>`).join(''),v=>{
    const sh=cur();sh.freeze=v==='top'?{r:1,c:0}:v==='first'?{r:0,c:1}:v==='here'?{r:X.act.r,c:X.act.c}:null;if(sh.freeze&&!sh.freeze.r&&!sh.freeze.c)sh.freeze=null;
    sh.viewDirty=true;touch();q('xeWrap').scrollTop=0;render();});
  else if(kind==='borders')pop(btn,[['all','כל הגבולות'],['outside','גבולות חיצוניים'],['thick','גבול חיצוני עבה'],['top','גבול עליון'],['bottom','גבול תחתון'],['dbottom','גבול תחתון כפול'],['none','בלי גבולות']].map(([v,l])=>`<button type="button" class="mi" data-v="${v}">${l}</button>`).join(''),v=>borders(v));
}
function marginsDialog(){
  const p=cur().page;const cm=v=>(v*2.54).toFixed(2);
  dialog(`<h3>שוליים מותאמים (ס״מ)</h3><div style="display:grid;grid-template-columns:auto 1fr;gap:8px 12px;align-items:center;margin-top:8px">
    ${[['t','למעלה'],['b','למטה'],['r','ימין'],['l','שמאל']].map(([k,l])=>`<label for="xeM${k}" style="font-size:13px;color:#5D6982">${l}</label><input id="xeM${k}" type="number" step="0.1" min="0" dir="ltr" value="${cm(p[k])}" style="height:34px;border:1px solid #CBD3DF;border-radius:8px;padding:0 10px;font:inherit">`).join('')}</div>
    <div class="acts"><button type="button" data-a="x">ביטול</button><button type="button" class="pri" data-a="ok">אישור</button></div>`,a=>{});
  const inn=q('xeDlg').firstElementChild;const vals={};
  inn.querySelectorAll('input').forEach(i=>i.oninput=()=>{vals[i.id.slice(3)]=Math.max(0,(+i.value||0)/2.54);});
  inn.onclick=e=>{const b=e.target.closest('[data-a]');if(!b)return;const ok=b.dataset.a==='ok';q('xeDlg').hidden=true;X.dlgOpen=false;inn.innerHTML='';if(ok&&Object.keys(vals).length)setPage(vals);q('xeWrap').focus();};
}
function dialog(html,onAct){
  const d=q('xeDlg'),inn=d.firstElementChild;inn.innerHTML=html;d.hidden=false;X.dlgOpen=true;
  const close=()=>{d.hidden=true;X.dlgOpen=false;inn.innerHTML='';q('xeWrap').focus({preventScroll:true});};
  inn.onclick=e=>{const b=e.target.closest('[data-a]');if(!b)return;close();onAct(b.dataset.a);};
  inn.onkeydown=e=>{if(e.key==='Escape'){e.preventDefault();close();onAct('x');}};
  setTimeout(()=>{const f=inn.querySelector('button.pri')||inn.querySelector('button');if(f)f.focus();},30);
}
function ask(title,text,btns){return new Promise(res=>dialog(`<h3>${title}</h3><p>${text}</p><div class="acts">${btns.map(b=>`<button type="button" data-a="${b.v}" class="${b.cls||''}">${b.l}</button>`).join('')}</div>`,a=>res(a==='x'?null:a)));}

/* ---------- ribbon state ---------- */
function updateUI(){
  if(!X)return;const sh=cur(),s=X.sel,a=X.act;const cell=cellAt(sh,a.r,a.c);const d=descOf(X.st.st,cell?cell.s:0);const rb=q('xeRibbon');
  q('xeRef').value=s.r1===s.r2&&s.c1===s.c2||mergeAt(sh,a.r,a.c)&&!(s.r2>mergeAt(sh,a.r,a.c).r2||s.c2>mergeAt(sh,a.r,a.c).c2)?refOf(a.r,a.c):rangeStr(s);
  if(!X.editing&&document.activeElement!==q('xeFx'))q('xeFx').value=rawText(cell);
  const press=(cmd,on)=>rb.querySelectorAll(`[data-cmd="${cmd}"]`).forEach(b=>b.setAttribute('aria-pressed',on?'true':'false'));
  press('bold',d.font.b);press('italic',d.font.i);press('underline',!!d.font.u);press('strike',d.font.strike);
  press('right',d.h==='right');press('center',d.h==='center'||d.h==='centerContinuous');press('left',d.h==='left');
  press('vtop',d.v==='top');press('vmid',d.v==='center');press('vbot',!d.v||d.v==='bottom');press('wrap',d.wrap);press('merge',!!mergeAt(sh,a.r,a.c));
  const f=rb.querySelector('[data-in=font]');if(![...f.options].some(o=>o.value===d.font.name))f.append(new Option(d.font.name,d.font.name));f.value=d.font.name;
  const z=rb.querySelector('[data-in=size]');if(document.activeElement!==z)z.value=d.font.sz;
  const fm=rb.querySelector('[data-in=fmt]');const known=FMTS.find(x=>x[0]===d.fmt);fm.value=known?d.fmt:'__custom';
  press('gridlines',sh.grid);press('headings',sh.headings!==false);press('formulas',sh.formulas);press('breaks',sh.breaks);press('printGrid',sh.page.grid);press('centerH',sh.page.hc);
  const sc=rb.querySelector('[data-in=scale]');if(sc&&document.activeElement!==sc)sc.value=sh.page.fit?'':Math.round(sh.page.scale||100);
  rb.querySelector('[data-cmd=paClear]').disabled=!sh.printArea;
  rb.querySelector('[data-cmd=undo]').disabled=X.hi<0;rb.querySelector('[data-cmd=redo]').disabled=X.hi>=X.hist.length-1;
  // sum, average and count of the selection
  let n=0,sum=0,cnt=0;if((s.r2-s.r1+1)*(s.c2-s.c1+1)>1){for(const c of sh.cells.values())if(c.r>=s.r1&&c.r<=s.r2&&c.c>=s.c1&&c.c<=s.c2){const v=c.f?valueOf(X.si,c.r,c.c):c.v;if(v!==null&&v!==''&&v!==undefined)cnt++;if(typeof v==='number'){n++;sum+=v;}}}
  const pages=sh.breaks&&X.brk?`עמודים להדפסה: ${X.brk.pages}   `:'';
  q('xeStat').textContent=pages+(n>1?`ממוצע: ${fmtGeneral(+(sum/n).toFixed(10))}   ספירה: ${cnt}   סכום: ${fmtGeneral(+sum.toFixed(10))}`:(cnt>1?`ספירה: ${cnt}`:''));
}
function buildTabs(){
  const box=q('xeSheets');box.innerHTML='';
  X.st.sheets.forEach((sh,i)=>{if(sh.hidden)return;const b=document.createElement('button');b.type='button';b.textContent=sh.name;b.className=i===X.si?'on':'';b.onclick=()=>switchSheet(i);box.append(b);});
  const add=document.createElement('button');add.type='button';add.className='add';add.title='גיליון חדש';add.setAttribute('aria-label','גיליון חדש');add.textContent='+';add.onclick=addSheet;box.append(add);
}
function switchSheet(i){commitEdit();X.si=i;X.sel={r1:0,c1:0,r2:0,c2:0};X.act={r:0,c:0};q('xeWrap').scrollTop=0;q('xeWrap').scrollLeft=0;buildTabs();render();q('xeWrap').focus({preventScroll:true});}
function addSheet(){
  const st=X.st;let n=st.sheets.length+1;const base=st.heb?'גיליון':'Sheet';while(st.sheets.some(s=>s.name===base+n))n++;
  let k=1;while(st.files.has(`xl/worksheets/sheet${k}.xml`))k++;const path=`xl/worksheets/sheet${k}.xml`;
  put(st.files,path,XML_HEAD+`<worksheet xmlns="${S}" xmlns:r="${NS.r}"><sheetViews><sheetView ${st.heb?'rightToLeft="1" ':''}workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`);
  ctOverride(st.files,path,OX+'spreadsheetml.worksheet+xml');
  const rid=addRel(st.files,st.wbPart,RT_SHEET,path.replace(/^xl\//,''));
  const sheetsEl=kid(st.wbDoc.documentElement,'sheets');const ids=kids(sheetsEl,'sheet').map(s=>+s.getAttribute('sheetId')||0);
  const el=sNew(st.wbDoc,'sheet',{name:base+n,sheetId:Math.max(0,...ids)+1});el.setAttributeNS(NS.r,'r:id',rid);sheetsEl.append(el);X.wbDirty=true;
  const sh=readSheet(st.files,path,st.sst);sh.name=base+n;sh.el=el;sh.wbIndex=kids(sheetsEl,'sheet').length-1;sh.dirty=true;st.sheets.push(sh);X.dirty=true;markDirty();switchSheet(st.sheets.length-1);
}

/* ---------- commands ---------- */
const CMD={
  undo,redo,
  copy:()=>{const t=doCopy(false);navigator.clipboard&&navigator.clipboard.writeText(t).catch(()=>{});hint('הועתק');},
  cut:()=>{const t=doCopy(true);navigator.clipboard&&navigator.clipboard.writeText(t).catch(()=>{});},
  paste:async()=>{try{const t=await navigator.clipboard.readText();if(t)doPaste(t);}catch(_){hint('הדבקה: Ctrl+V');}},
  bold:()=>toggleFont('b'),italic:()=>toggleFont('i'),underline:()=>toggleFont('u'),strike:()=>toggleFont('strike'),
  grow:()=>{const z=actDesc().font.sz;style({font:{sz:SIZES.find(x=>x>z)||z+2}});},shrink:()=>{const z=actDesc().font.sz;style({font:{sz:[...SIZES].reverse().find(x=>x<z)||Math.max(1,z-1)}});},
  borders:()=>borders('all'),fill:()=>setFill(X.fill===undefined?'FFFF00':X.fill),color:()=>setColor(X.color===undefined?'C00000':X.color),
  vtop:()=>style({al:{v:'top'}}),vmid:()=>style({al:{v:'center'}}),vbot:()=>style({al:{v:null}}),
  wrap:()=>style({al:{wrap:!actDesc().wrap}}),
  right:()=>style({al:{h:actDesc().h==='right'?null:'right'}}),center:()=>style({al:{h:actDesc().h==='center'?null:'center'}}),left:()=>style({al:{h:actDesc().h==='left'?null:'left'}}),
  merge:toggleMerge,
  pct:()=>setFmt('0%'),thou:()=>setFmt('#,##0'),decP:()=>decimals(1),decM:()=>decimals(-1),
  rowIns:insRows,colIns:insCols,rowDel:delRows,colDel:delCols,
  sum:autoSum,sortAZ:()=>sortSel(false),sortZA:()=>sortSel(true),
  clear:()=>{change(()=>{eachSelCell(c=>{c.s=0;c.mod=true;});X.st.st.css=null;});},
  find:()=>{q('xeFind').hidden=false;q('xeFindIn').focus();q('xeFindIn').select();},
  zin:()=>setZoom(X.zoom+.1),zout:()=>setZoom(X.zoom-.1),z100:()=>setZoom(1),
  zsel:()=>{const sh=cur(),g=geo(sh),s=X.sel,w=q('xeWrap');let tw=0;for(let c=s.c1;c<=s.c2;c++)tw+=g.cw[c]||colPx(sh,c);const th=g.top[Math.min(s.r2+1,g.rows)]-g.top[s.r1];
    setZoom(Math.min((w.clientWidth-60)/Math.max(40,tw),(w.clientHeight-40)/Math.max(20,th)));select(s.r1,s.c1);},
  headings:()=>{const sh=cur();sh.headings=sh.headings===false;sh.viewDirty=true;touch();render();},
  formulas:()=>{const sh=cur();sh.formulas=!sh.formulas;sh.viewDirty=true;touch();render();},
  breaks:()=>{const sh=cur();sh.breaks=!sh.breaks;render();},
  printGrid:()=>{const sh=cur();sh.page.grid=!sh.page.grid;sh.pageDirty=true;touch();updateUI();},
  centerH:()=>{const sh=cur();sh.page.hc=!sh.page.hc;sh.pageDirty=true;touch();updateUI();},
  paSet:()=>{const sh=cur();const s=X.sel;if(s.r1===s.r2&&s.c1===s.c2)return hint('סמנו קודם את התאים שיודפסו');sh.printArea={...s};sh.paDirty=true;sh.breaks=true;touch();render();hint('אזור ההדפסה: '+rangeStr(s));},
  paClear:()=>{const sh=cur();if(!sh.printArea)return;sh.printArea=null;sh.paDirty=true;touch();render();},
  print:printSheet,
  gridlines:()=>{const sh=cur();sh.grid=!sh.grid;sh.gridChanged=true;sh.dirty=true;X.dirty=true;markDirty();render();X.root.querySelectorAll('[data-cmd=gridlines]').forEach(b=>b.setAttribute('aria-pressed',sh.grid));}
};
function touch(){const sh=cur();sh.dirty=true;X.dirty=true;markDirty();}
function setPage(ch){const sh=cur();Object.assign(sh.page,ch);sh.pageDirty=true;touch();if(!sh.breaks)sh.breaks=true;render();}
function printSheet(){
  const sh=cur(),g=geo(sh),A=printArea(sh),b=pageBreaks(sh),p=sh.page;
  const vis=[];for(let c=A.c1;c<=A.c2;c++)if(colPx(sh,c)>4)vis.push(c);
  const o={lead:0,headH:0,left:new Map(),fc:0,fr:0,brk:null,print:true};
  const box=q('xePrint');box.className=p.grid?'pg':'';
  box.innerHTML=`<table style="direction:${sh.rtl?'rtl':'ltr'};zoom:${b.scale};${p.hc?'margin:0 auto':''}"><colgroup>${vis.map(c=>`<col style="width:${colPx(sh,c)}px">`).join('')}</colgroup><tbody>${rowsHtml(sh,g,A.r1,A.r2,vis,o)}</tbody></table>`;
  let st=document.getElementById('xePrintPage');if(!st){st=document.createElement('style');st.id='xePrintPage';document.head.append(st);}
  st.textContent=`@page{size:${b.w}in ${b.h}in;margin:${p.t}in ${p.r}in ${p.b}in ${p.l}in}`;
  setTimeout(()=>{print();setTimeout(()=>{if(X)q('xePrint').innerHTML='';},500);},50);
}
function setZoom(z){X.zoom=Math.max(.5,Math.min(2,Math.round(z*10)/10));q('xeZoomV').textContent=Math.round(X.zoom*100)+'%';q('xeZoom').value=Math.round(X.zoom*100);render();}

/* ---------- events ---------- */
function cellFromEvent(e){const td=e.target.closest&&e.target.closest('td[data-r]');if(!td)return null;return{r:+(td.dataset.mr??td.dataset.r),c:+(td.dataset.mc??td.dataset.c)};}
function wire(){
  const R=X.root,wrap=q('xeWrap'),rb=q('xeRibbon');
  R.querySelector('.xe-tabs').onclick=e=>{const b=e.target.closest('[data-tab]');if(!b)return;R.querySelectorAll('.xe-tabs [data-tab]').forEach(x=>x.setAttribute('aria-selected',x===b?'true':'false'));R.querySelectorAll('.xpanel').forEach(p=>p.hidden=p.dataset.panel!==b.dataset.tab);};
  R.addEventListener('mousedown',e=>{if(e.target.closest('button.xb'))e.preventDefault();});
  R.addEventListener('click',e=>{const b=e.target.closest('[data-cmd]');if(!b||b.disabled)return;if(b.dataset.pop){openPop(b.dataset.pop,b);return;}closePop();commitEdit();const f=CMD[b.dataset.cmd];if(f)f();});
  rb.querySelector('[data-in=font]').onchange=e=>{style({font:{name:e.target.value}});wrap.focus();};
  const sz=rb.querySelector('[data-in=size]');const apply=()=>{const v=parseFloat(sz.value);if(v>0)style({font:{sz:v}});};
  sz.onchange=apply;sz.onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();apply();wrap.focus();}};
  const sc=rb.querySelector('[data-in=scale]');sc.onchange=()=>{const v=Math.max(10,Math.min(400,+sc.value||100));setPage({fit:false,scale:v});};
  sc.onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();sc.onchange();wrap.focus();}};
  rb.querySelector('[data-in=fmt]').onchange=e=>{if(e.target.value!=='__custom')setFmt(e.target.value);wrap.focus();};
  // grid
  let drag=false;
  wrap.addEventListener('mousedown',e=>{
    if(e.button!==0)return;const rz=e.target.closest('[data-rz]');if(rz){startResize(e,+rz.dataset.rz);return;}
    const th=e.target.closest('th');
    if(th){e.preventDefault();commitEdit();const sh=cur(),g=geo(sh);
      if(th.dataset.all){X.act={r:0,c:0};X.anchor={r:0,c:0};X.sel={r1:0,c1:0,r2:Math.max(sh.maxR,0),c2:Math.max(sh.maxC,0)};}
      else if(th.dataset.col!==undefined){const c=+th.dataset.col;if(e.shiftKey){X.sel={r1:0,r2:Math.max(sh.maxR,g.rows-1),c1:Math.min(X.anchor.c,c),c2:Math.max(X.anchor.c,c)};}else{X.act={r:0,c};X.anchor={r:0,c};X.sel={r1:0,r2:Math.max(sh.maxR,0),c1:c,c2:c};}}
      else if(th.dataset.row!==undefined){const r=+th.dataset.row;if(e.shiftKey){X.sel={c1:0,c2:Math.max(sh.maxC,0),r1:Math.min(X.anchor.r,r),r2:Math.max(X.anchor.r,r)};}else{X.act={r,c:0};X.anchor={r,c:0};X.sel={r1:r,r2:r,c1:0,c2:Math.max(sh.maxC,0)};}}
      wrap.focus({preventScroll:true});paintSel();return;}
    const p=cellFromEvent(e);if(!p)return;e.preventDefault();
    if(X.editing){if(X.editing.r===p.r&&X.editing.c===p.c)return;commitEdit();}
    wrap.focus({preventScroll:true});if(e.shiftKey)select(p.r,p.c,true);else select(p.r,p.c);drag=true;
  });
  wrap.addEventListener('mouseover',e=>{if(!drag||!(e.buttons&1)){drag=false;return;}const p=cellFromEvent(e);if(p)select(p.r,p.c,true);});
  X.onUp=()=>{drag=false;};window.addEventListener('mouseup',X.onUp);
  wrap.addEventListener('dblclick',e=>{const p=cellFromEvent(e);if(p){select(p.r,p.c);startEdit();}});
  wrap.addEventListener('scroll',()=>{if(X.editing)commitEdit();closePop();renderSoon();},{passive:true});
  const ed=q('xeEdit');ed.addEventListener('input',()=>{q('xeFx').value=ed.value;growEdit();});
  const fx=q('xeFx');
  fx.addEventListener('focus',()=>{if(!X.editing)X.editing={r:X.act.r,c:X.act.c,typed:false,orig:rawText(cellAt(cur(),X.act.r,X.act.c)),fx:true};});
  fx.addEventListener('input',()=>{const e2=q('xeEdit');if(!e2.hidden)e2.value=fx.value;});
  fx.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();const e2=q('xeEdit');e2.hidden=true;commitEdit([1,0]);}else if(e.key==='Escape'){e.preventDefault();fx.value=X.editing?X.editing.orig:'';cancelEdit();}});
  fx.addEventListener('blur',()=>{if(X&&X.editing&&X.editing.fx&&q('xeEdit').hidden){const v=fx.value;const ed2=X.editing;X.editing=null;if(v!==ed2.orig)setInput(ed2.r,ed2.c,v);}});
  const ref=q('xeRef');ref.addEventListener('keydown',e=>{if(e.key!=='Enter')return;e.preventDefault();const g=parseRange(ref.value.replace(/\s/g,''));
    if(!g){hint('כתובת לא תקינה, למשל B7 או A1:C5');return;}select(g.r1,g.c1);if(g.r2>g.r1||g.c2>g.c1){X.sel=g;paintSel();}wrap.focus();});
  ref.addEventListener('focus',()=>ref.select());
  q('xeFindIn').addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();findNext();}else if(e.key==='Escape'){q('xeFind').hidden=true;wrap.focus();}});
  q('xeFindNext').onclick=findNext;q('xeFindX').onclick=()=>{q('xeFind').hidden=true;wrap.focus();};
  q('xeZoom').oninput=e=>setZoom(+e.target.value/100);
  q('xeSave').onclick=()=>doSave();q('xePdf').onclick=()=>doPdf();q('xeDesk').onclick=()=>doDesk();q('xeClose').onclick=()=>doClose();
  X.onKey=onKey;window.addEventListener('keydown',X.onKey,true);
  X.onCopy=e=>{if(!X)return;if(!R.contains(e.target)&&e.target!==document.body){e.stopImmediatePropagation();return;}
    if(X.editing||['INPUT','TEXTAREA'].includes(e.target.tagName))return;e.stopImmediatePropagation();e.preventDefault();e.clipboardData.setData('text/plain',doCopy(e.type==='cut'));};
  X.onPaste=e=>{if(!X)return;if(['INPUT','TEXTAREA'].includes(e.target.tagName)&&R.contains(e.target))return;e.stopImmediatePropagation();
    if(X.dlgOpen)return;e.preventDefault();const t=e.clipboardData&&e.clipboardData.getData('text/plain');if(t)doPaste(t);};
  window.addEventListener('copy',X.onCopy,true);window.addEventListener('cut',X.onCopy,true);window.addEventListener('paste',X.onPaste,true);
  ['keydown','keyup','keypress','input','change'].forEach(t=>R.addEventListener(t,e=>e.stopPropagation()));
  X.onDrag=e=>{if(!X)return;e.stopImmediatePropagation();e.preventDefault();};window.addEventListener('dragover',X.onDrag,true);window.addEventListener('drop',X.onDrag,true);
  X.onUnload=e=>{if(X&&X.dirty){e.preventDefault();e.returnValue='';}};window.addEventListener('beforeunload',X.onUnload);
  X.onDown=e=>{if(!X)return;const p=q('xePop');if(!p.hidden&&!p.contains(e.target)&&!(X.popFor&&X.popFor.contains(e.target)))closePop();};document.addEventListener('pointerdown',X.onDown,true);
  X.onResize=()=>renderSoon();window.addEventListener('resize',X.onResize);
}
function startResize(e,c){
  e.preventDefault();e.stopPropagation();const sh=cur(),x0=e.clientX,w0=colPx(sh,c),rtl=sh.rtl;
  const mv=ev=>{const dx=(ev.clientX-x0)/X.zoom*(rtl?-1:1);const w=Math.max(8,w0+dx);const ci=sh.cols.get(c)||{};sh.cols.set(c,{...ci,w:Math.max(0,(w-5)/7),custom:'1',hidden:false});render();};
  const up=()=>{removeEventListener('mousemove',mv);removeEventListener('mouseup',up);sh.dirty=true;X.dirty=true;markDirty();};
  addEventListener('mousemove',mv);addEventListener('mouseup',up);
}
function onKey(e){
  if(!X)return;const t=e.target,inRoot=X.root.contains(t);
  if(!inRoot)e.stopImmediatePropagation();
  if(X.dlgOpen){if(!inRoot)e.preventDefault();return;}
  const mod=(e.ctrlKey||e.metaKey)&&!e.altKey,c=e.code,k=e.key;const run=(f,...a)=>{e.preventDefault();f(...a);};
  if(mod&&c==='KeyS')return run(doSave);
  if(mod&&c==='KeyF')return run(CMD.find);
  if(mod&&c==='KeyP')return run(printSheet);
  if(k==='Escape'&&!q('xePop').hidden)return run(closePop);
  const rtl=cur().rtl;
  const editor=t===q('xeEdit');
  if(editor){
    if(k==='Enter'&&e.altKey){e.preventDefault();const el=q('xeEdit');const p=el.selectionStart;el.value=el.value.slice(0,p)+'\n'+el.value.slice(el.selectionEnd);el.selectionStart=el.selectionEnd=p+1;growEdit();q('xeFx').value=el.value;return;}
    if(k==='Enter')return run(commitEdit,[e.shiftKey?-1:1,0]);
    if(k==='Tab')return run(commitEdit,[0,e.shiftKey?-1:1]);
    if(k==='Escape')return run(cancelEdit);
    if(X.editing&&X.editing.typed&&/^Arrow/.test(k)){const d={ArrowUp:[-1,0],ArrowDown:[1,0],ArrowLeft:[0,rtl?1:-1],ArrowRight:[0,rtl?-1:1]}[k];return run(commitEdit,d);}
    return;
  }
  if(t!==q('xeWrap'))return;
  if(mod){
    if(c==='KeyZ')return run(e.shiftKey?redo:undo);if(c==='KeyY')return run(redo);
    if(c==='KeyB')return run(CMD.bold);if(c==='KeyI')return run(CMD.italic);if(c==='KeyU')return run(CMD.underline);
    if(c==='KeyA'){e.preventDefault();const sh=cur();X.sel={r1:0,c1:0,r2:Math.max(0,sh.maxR),c2:Math.max(0,sh.maxC)};paintSel();return;}
    if(c==='Home')return run(select,0,0,e.shiftKey);
    if(/^Arrow/.test(k)){const d={ArrowUp:[-1,0],ArrowDown:[1,0],ArrowLeft:[0,rtl?1:-1],ArrowRight:[0,rtl?-1:1]}[k];return run(jump,d[0],d[1],e.shiftKey);}
    if(c==='KeyC'||c==='KeyX'||c==='KeyV')return;
    return;
  }
  if(e.altKey&&(c==='Equal'))return run(autoSum);
  const arrows={ArrowUp:[-1,0],ArrowDown:[1,0],ArrowLeft:[0,rtl?1:-1],ArrowRight:[0,rtl?-1:1]};
  if(arrows[k])return run(moveAct,...arrows[k],e.shiftKey);
  if(k==='Enter')return run(moveAct,e.shiftKey?-1:1,0);
  if(k==='Tab')return run(moveAct,0,e.shiftKey?-1:1);
  if(k==='Home')return run(select,X.act.r,0,e.shiftKey);
  if(k==='PageDown'||k==='PageUp'){const n=Math.max(5,Math.floor(q('xeWrap').clientHeight/(20*X.zoom))-2);return run(moveAct,k==='PageDown'?n:-n,0,e.shiftKey);}
  if(k==='Delete'||k==='Backspace'){if(k==='Backspace'){e.preventDefault();startEdit('',true);return;}return run(clearCells);}
  if(k==='F2')return run(startEdit);
  if(k.length===1&&!e.ctrlKey&&!e.metaKey){startEdit('',true);return;}
}

/* ------------------------------------------------------------ saving */
const SHEET_ORDER=['sheetPr','dimension','sheetViews','sheetFormatPr','cols','sheetData','sheetCalcPr','sheetProtection','protectedRanges','scenarios','autoFilter','sortState','dataConsolidate','customSheetViews','mergeCells','phoneticPr','conditionalFormatting','dataValidations','hyperlinks','printOptions','pageMargins','pageSetup','headerFooter','rowBreaks','colBreaks','customProperties','cellWatches','ignoredErrors','smartTags','drawing','legacyDrawing','legacyDrawingHF','drawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst'];
function insertOrdered(root,el){const i=SHEET_ORDER.indexOf(el.localName);const nx=kids(root).find(c=>SHEET_ORDER.indexOf(c.localName)>i);root.insertBefore(el,nx||null);}
const SHEET_AFTER_MERGE=['phoneticPr','conditionalFormatting','dataValidations','hyperlinks','printOptions','pageMargins','pageSetup','headerFooter','rowBreaks','colBreaks','customProperties','cellWatches','ignoredErrors','smartTags','drawing','legacyDrawing','legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst'];
function numStr(v){if(!isFinite(v))return'0';return String(v);}
function sheetXml(si,ctx){
  const sh=X.st.sheets[si];const d=sh.doc.cloneNode(true),root=d.documentElement;
  const sd=kid(root,'sheetData')||root.appendChild(sNew(d,'sheetData'));while(sd.firstChild)sd.removeChild(sd.firstChild);
  const pfx=sd.prefix?sd.prefix+':':'';
  for(const c of kids(root,'cols'))root.removeChild(c);
  if(sh.cols.size){const ce=sNew(d,'cols');const list=[...sh.cols.entries()].sort((a,b)=>a[0]-b[0]);
    let i=0;while(i<list.length){const [c0,v]=list[i];let j=i;while(j+1<list.length&&list[j+1][0]===list[j][0]+1&&JSON.stringify(list[j+1][1])===JSON.stringify(v))j++;
      const a={min:c0+1,max:list[j][0]+1,width:(v.w!=null?+v.w.toFixed(4):+sh.defW.toFixed(4))};if(v.style)a.style=v.style;if(v.hidden)a.hidden=1;if(v.w!=null)a.customWidth=1;ce.append(sNew(d,'col',a));i=j+1;}
    root.insertBefore(ce,sd);}
  for(const m of kids(root,'mergeCells'))root.removeChild(m);
  if(sh.merges.length){const me=sNew(d,'mergeCells',{count:sh.merges.length});for(const m of sh.merges)me.append(sNew(d,'mergeCell',{ref:rangeStr(m)}));
    const after=kids(root).find(e=>SHEET_AFTER_MERGE.includes(e.localName));root.insertBefore(me,after||null);}
  const dim=kid(root,'dimension');if(dim)dim.setAttribute('ref',sh.cells.size?`A1:${refOf(sh.maxR,sh.maxC)}`:'A1');
  if(sh.gridChanged||sh.viewDirty){
    let svs=kid(root,'sheetViews');if(!svs){svs=sNew(d,'sheetViews');insertOrdered(root,svs);}
    let sv=kid(svs,'sheetView');if(!sv){sv=sNew(d,'sheetView',{workbookViewId:0});svs.append(sv);}
    if(sh.grid)sv.removeAttribute('showGridLines');else sv.setAttribute('showGridLines','0');
    if(sh.headings!==false)sv.removeAttribute('showRowColHeaders');else sv.setAttribute('showRowColHeaders','0');
    if(sh.formulas)sv.setAttribute('showFormulas','1');else sv.removeAttribute('showFormulas');
    if(sh.viewDirty){for(const e of kids(sv,'pane','selection'))sv.removeChild(e);
      if(sh.freeze){const f=sh.freeze;const pn=sNew(d,'pane',{xSplit:f.c||null,ySplit:f.r||null,topLeftCell:refOf(f.r,f.c),activePane:f.r&&f.c?'bottomRight':f.r?'bottomLeft':'topRight',state:'frozen'});sv.insertBefore(pn,sv.firstChild);}}
  }
  if(sh.pageDirty){
    const p=sh.page;
    let po=kid(root,'printOptions');if(p.grid||p.hc){if(!po){po=sNew(d,'printOptions');insertOrdered(root,po);}
      if(p.grid)po.setAttribute('gridLines','1');else po.removeAttribute('gridLines');if(p.hc)po.setAttribute('horizontalCentered','1');else po.removeAttribute('horizontalCentered');}
    else if(po){po.removeAttribute('gridLines');po.removeAttribute('horizontalCentered');if(!po.attributes.length)root.removeChild(po);}
    let pm=kid(root,'pageMargins');if(!pm){pm=sNew(d,'pageMargins');insertOrdered(root,pm);}
    for(const [k,a] of [['l','left'],['r','right'],['t','top'],['b','bottom'],['h','header'],['f','footer']])pm.setAttribute(a,+(+p[k]).toFixed(4));
    let ps=kid(root,'pageSetup');if(!ps){ps=sNew(d,'pageSetup');insertOrdered(root,ps);}
    ps.setAttribute('paperSize',p.paper);ps.setAttribute('orientation',p.orient);
    if(p.fit){ps.setAttribute('fitToWidth',p.fw);ps.setAttribute('fitToHeight',p.fh);ps.removeAttribute('scale');}
    else{ps.setAttribute('scale',Math.round(p.scale||100));ps.removeAttribute('fitToWidth');ps.removeAttribute('fitToHeight');}
    let spr=kid(root,'sheetPr');if(!spr){spr=sNew(d,'sheetPr');root.insertBefore(spr,root.firstElementChild);}
    let pp=kid(spr,'pageSetUpPr');if(p.fit){if(!pp){pp=sNew(d,'pageSetUpPr');spr.append(pp);}pp.setAttribute('fitToPage','1');}else if(pp)pp.removeAttribute('fitToPage');
  }
  let s=ser(d);if(!/^<\?xml/.test(s))s=XML_HEAD+s;
  // rows
  const byRow=new Map();for(const c of sh.cells.values()){if(!byRow.has(c.r))byRow.set(c.r,[]);byRow.get(c.r).push(c);}
  const rowsIdx=[...new Set([...byRow.keys(),...sh.rows.keys()])].sort((a,b)=>a-b);
  let x='';
  for(const r of rowsIdx){
    const info=sh.rows.get(r);let at=`r="${r+1}"`;if(info)for(const [k,v] of Object.entries(info.attrs))at+=` ${k}="${esc(v)}"`;
    const cells=(byRow.get(r)||[]).sort((a,b)=>a.c-b.c);
    if(!cells.length&&!info)continue;
    x+=`<${pfx}row ${at}>${cells.map(c=>cellXml(c,si,pfx,ctx)).join('')}</${pfx}row>`;
  }
  return s.replace(/<((?:\w+:)?)sheetData\s*\/>/,(m,p)=>`<${p}sheetData>${x}</${p}sheetData>`);
}
function cellXml(c,si,pfx,ctx){
  if(!c.mod&&c.el&&!c.wasShared&&!(c.el.getElementsByTagNameNS(S,'f')[0]&&c.el.getElementsByTagNameNS(S,'f')[0].getAttribute('t')==='shared'))
    return ser(c.el).replace(/ xmlns(:\w+)?="[^"]*"/g,'');
  const r=refOf(c.r,c.c);let a=`r="${r}"`;if(c.s)a+=` s="${c.s}"`;
  if(c.f){
    let v=X.calc.value(si,c.r,c.c);if(v===undefined)v=c.v;
    const f=`<${pfx}f${c.array?` t="array" ref="${c.array}"`:''}>${esc(c.f)}</${pfx}f>`;
    if(v===null||v===undefined)return`<${pfx}c ${a}>${f}</${pfx}c>`;
    if(isErr(v))return`<${pfx}c ${a} t="e">${f}<${pfx}v>${esc(v.e)}</${pfx}v></${pfx}c>`;
    if(typeof v==='boolean')return`<${pfx}c ${a} t="b">${f}<${pfx}v>${v?1:0}</${pfx}v></${pfx}c>`;
    if(typeof v==='number')return`<${pfx}c ${a}>${f}<${pfx}v>${numStr(v)}</${pfx}v></${pfx}c>`;
    return`<${pfx}c ${a} t="str">${f}<${pfx}v>${esc(v)}</${pfx}v></${pfx}c>`;
  }
  if(c.v===null||c.v===undefined||c.v==='')return`<${pfx}c ${a}/>`;
  if(c.t==='b')return`<${pfx}c ${a} t="b"><${pfx}v>${c.v?1:0}</${pfx}v></${pfx}c>`;
  if(c.t==='e')return`<${pfx}c ${a} t="e"><${pfx}v>${esc(c.v)}</${pfx}v></${pfx}c>`;
  if(c.t==='n'&&typeof c.v==='number')return`<${pfx}c ${a}><${pfx}v>${numStr(c.v)}</${pfx}v></${pfx}c>`;
  const str=String(c.v);let idx=c.sidx!==undefined&&X.st.sst.list[c.sidx]===str?c.sidx:ctx.addStr(str);
  return`<${pfx}c ${a} t="s"><${pfx}v>${idx}</${pfx}v></${pfx}c>`;
}
async function buildXlsx(){
  if(X.editing)commitEdit();
  const st=X.st,files=new Map(st.files);const sst=st.sst;
  const newStrs=[];const strIdx=new Map();
  const ctx={addStr:s=>{if(strIdx.has(s))return strIdx.get(s);const i=sst.list.length+newStrs.length;newStrs.push(s);strIdx.set(s,i);return i;}};
  X.calc.reset();
  // formulas whose result changed get the new result written, so every program shows it
  st.sheets.forEach((sh,i)=>{for(const c of sh.cells.values()){if(!c.f||c.array)continue;const v=X.calc.value(i,c.r,c.c);
    if(c.calcOk===false)continue;const old=c.t==='e'?ERR(c.v):c.v;const same=isErr(v)?isErr(old)&&old.e===v.e:(typeof v==='number'&&typeof old==='number'?Math.abs(v-old)<=1e-9*Math.max(1,Math.abs(v)):v===old);
    if(!same){c.mod=true;sh.dirty=true;}}});
  st.sheets.forEach((sh,i)=>{if(sh.dirty)put(files,sh.path,sheetXml(i,ctx));});
  if(newStrs.length){
    let s=sst.doc?ser(sst.doc):XML_HEAD+`<sst xmlns="${S}" count="0" uniqueCount="0"></sst>`;
    const add=newStrs.map(t=>`<si><t xml:space="preserve">${esc(t).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,'')}</t></si>`).join('');
    s=/<sst\b[^>]*\/>\s*$/.test(s)?s.replace(/<sst\b([^>]*)\/>\s*$/,`<sst$1>${add}</sst>`):s.replace(/<\/sst>\s*$/,add+'</sst>');
    const n=sst.list.length+newStrs.length;s=s.replace(/uniqueCount="\d+"/,`uniqueCount="${n}"`);if(!/uniqueCount=/.test(s))s=s.replace(/<sst\b/,`<sst uniqueCount="${n}"`);
    put(files,sst.path,s);
    if(!sst.hadPart){ctOverride(files,sst.path,OX+'spreadsheetml.sharedStrings+xml');if(!rels(files,st.wbPart).some(r=>r.type===RT_SST))addRel(files,st.wbPart,RT_SST,sst.path.replace(/^xl\//,''));}
  }
  if(X.styleDirty){const S2=st.st;for(const [el,n] of [[S2.fontsEl,'font'],[S2.fillsEl,'fill'],[S2.bordersEl,'border'],[S2.xfsEl,'xf']])el.setAttribute('count',kids(el,n).length);
    const nf=kid(S2.root,'numFmts');if(nf)nf.setAttribute('count',kids(nf,'numFmt').length);put(files,S2.path,ser(S2.doc));}
  for(const sh of st.sheets)if(sh.paDirty){
    const root=st.wbDoc.documentElement;let dn=kid(root,'definedNames');
    const old=[...st.wbDoc.getElementsByTagNameNS(S,'definedName')].find(e=>e.getAttribute('name')==='_xlnm.Print_Area'&&+e.getAttribute('localSheetId')===sh.wbIndex);
    if(old)old.parentNode.removeChild(old);
    if(sh.printArea){if(!dn){dn=sNew(st.wbDoc,'definedNames');const after=kids(root).find(e=>['calcPr','oleSize','customWorkbookViews','pivotCaches','smartTagPr','smartTagTypes','webPublishing','fileRecoveryPr','webPublishObjects','extLst'].includes(e.localName));root.insertBefore(dn,after||null);}
      const a=sh.printArea,nm=/^[A-Za-z_][\w.]*$/.test(sh.name)?sh.name:`'${sh.name.replace(/'/g,"''")}'`;
      const e=sNew(st.wbDoc,'definedName',{name:'_xlnm.Print_Area',localSheetId:sh.wbIndex});e.textContent=`${nm}!$${colName(a.c1)}$${a.r1+1}:$${colName(a.c2)}$${a.r2+1}`;dn.append(e);}
    else if(dn&&!kids(dn,'definedName').length)root.removeChild(dn);
    X.wbDirty=true;}
  const fchg=X.formulaTouched||X.structChanged;
  if(fchg||X.wbDirty){
    const root=st.wbDoc.documentElement;if(fchg){let cp=kid(root,'calcPr');if(!cp){cp=sNew(st.wbDoc,'calcPr',{calcId:191029});const after=kids(root).find(e=>['oleSize','customWorkbookViews','pivotCaches','smartTagPr','smartTagTypes','webPublishing','fileRecoveryPr','webPublishObjects','extLst'].includes(e.localName));root.insertBefore(cp,after||null);}cp.setAttribute('fullCalcOnLoad','1');}
    put(files,st.wbPart,ser(st.wbDoc));
    // Excel rebuilds the calculation chain by itself
    const cc=rels(files,st.wbPart).find(r=>/calcChain$/.test(r.type));
    if(cc){const p=resolveTarget(st.wbPart,cc.target);files.delete(p);ctRemove(files,p);const rp=relsPathOf(st.wbPart);put(files,rp,txt(files,rp).replace(new RegExp(`<Relationship[^>]*Id="${cc.id}"[^>]*/>`),''));}
  }
  const z=new JSZip();for(const [p,b] of files)z.file(p,b);
  return z.generateAsync({type:'uint8array',compression:'DEFLATE',compressionOptions:{level:6}});
}

/* ------------------------------------------------------------ opening and closing */
function injectCss(){if(document.getElementById('xeCss'))return;const s=document.createElement('style');s.id='xeCss';s.textContent=CSS;document.head.append(s);}
async function open(opts){
  if(X)await forceClose();
  if(!window.JSZip)throw new Error('ספריות העריכה לא נטענו');
  injectCss();
  const root=document.createElement('div');root.id='xed';root.innerHTML=shellHtml();document.body.append(root);
  X={opts,root,si:0,sel:{r1:0,c1:0,r2:0,c2:0},act:{r:0,c:0},anchor:{r:0,c:0},cursor:{r:0,c:0},hist:[],hi:-1,zoom:1,dirty:false};
  q('xeName').value=(opts.name||'גיליון חדש.xlsx').replace(/\.(xlsx|xlsm|xls)$/i,'');
  const load=document.createElement('div');load.className='xe-load';load.textContent='פותח את הגיליון לעריכה…';q('xeWrap').append(load);
  try{X.st=await prepare(opts.bytes||null);}catch(e){root.remove();X=null;throw e;}
  load.remove();X.calc=makeCalc(X);
  X.si=Math.max(0,X.st.sheets.findIndex(s=>!s.hidden));
  wire();buildTabs();render();q('xeWrap').focus({preventScroll:true});select(0,0);
  return true;
}
const fileName=()=>((q('xeName').value||'').trim().replace(/[\\/:*?"<>|]+/g,'').replace(/\.xlsx$/i,'')||'גיליון')+'.xlsx';
async function build(msg){X.opts.busy(true,msg||'שומר את הגיליון…');try{await nextFrame();return await buildXlsx();}finally{X.opts.busy(false);}}
async function doSave(){
  if(!X||X.saving)return;X.saving=true;
  try{const b=await build();if(await X.opts.save(fileName(),b,'xlsx')){X.dirty=false;X.savedOnce=true;markDirty();X.opts.toast('נשמר: '+fileName());}}
  catch(e){if(!(e&&e.name==='AbortError')){console.error(e);X.opts.toast('השמירה נכשלה: '+(e.message||e));}}finally{if(X)X.saving=false;}
}
async function doPdf(){
  if(!X||X.saving)return;X.saving=true;
  try{const b=await build('מכין PDF…');X.opts.busy(true,'מכין PDF…');const pdf=await X.opts.toPdf(b);X.opts.busy(false);const n=fileName().replace(/\.xlsx$/,'.pdf');if(await X.opts.save(n,pdf,'pdf'))X.opts.toast('נשמר: '+n);}
  catch(e){X.opts.busy(false);if(!(e&&e.name==='AbortError')){console.error(e);X.opts.toast('יצירת ה-PDF נכשלה: '+(e.message||e));}}finally{if(X)X.saving=false;}
}
async function doDesk(){
  if(!X||X.saving)return;X.saving=true;
  try{const b=await build('מכין את הגיליון לחתימה…');const f=new File([b],fileName(),{type:XLSX_MIME});const o=X.opts;await forceClose();await o.toDesk(f);}
  catch(e){console.error(e);if(X){X.saving=false;X.opts.toast('לא הצלחתי לפתוח לחתימה: '+(e.message||e));}}
}
async function doClose(){
  if(!X)return;commitEdit();
  if(X.dirty){const a=await ask('לשמור את השינויים?',`יש שינויים ב-${esc(fileName())} שעוד לא נשמרו.`,[{v:'x',l:'ביטול'},{v:'drop',l:'סגירה בלי לשמור',cls:'warn'},{v:'save',l:'שמירה',cls:'pri'}]);
    if(!a)return;if(a==='save'){await doSave();if(X&&X.dirty)return;}}
  const o=X.opts;await forceClose();o.onClose&&o.onClose();
}
async function forceClose(){
  if(!X)return;const x=X;
  window.removeEventListener('keydown',x.onKey,true);window.removeEventListener('copy',x.onCopy,true);window.removeEventListener('cut',x.onCopy,true);window.removeEventListener('paste',x.onPaste,true);
  window.removeEventListener('dragover',x.onDrag,true);window.removeEventListener('drop',x.onDrag,true);window.removeEventListener('beforeunload',x.onUnload);
  document.removeEventListener('pointerdown',x.onDown,true);window.removeEventListener('resize',x.onResize);window.removeEventListener('mouseup',x.onUp);
  x.root.remove();X=null;
}
window.XLSX_EDITOR={open,close:forceClose,isOpen:()=>!!X,
  _test:{get X(){return X;},buildXlsx:()=>buildXlsx(),setInput,select,style,borders,toggleMerge,insRows,insCols,delRows,delCols,sortSel,autoSum,undo,redo,doPaste,doCopy,display:(r,c)=>display(X.si,cellAt(cur(),r,c),r,c).t,parseFormula,moveFormula,shiftFormula}};
})();
