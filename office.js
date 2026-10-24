/* סימן — Office files (docx / xlsx / pptx): inspect, repair, extract.
   An Office file is a ZIP of XML parts. Most "can't open" errors come from
   one damaged part, a broken archive, or references to parts that are gone.
   Needs: JSZip (window.JSZip), PptxGenJS (window.PptxGenJS) for slide rebuild. */
'use strict';
const OFFICE = (() => {
  const enc = new TextEncoder(), dec = new TextDecoder('utf-8', { fatal: false });
  const KIND = {
    pptx: { name: 'מצגת PowerPoint', main: 'ppt/presentation.xml', mime: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' },
    docx: { name: 'מסמך Word', main: 'word/document.xml', mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
    xlsx: { name: 'גיליון Excel', main: 'xl/workbook.xml', mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }
  };
  const OX = 'application/vnd.openxmlformats-officedocument.';
  // content types for parts, by path pattern
  const CT = [
    [/^ppt\/presentation\.xml$/, OX + 'presentationml.presentation.main+xml'],
    [/^ppt\/slides\/slide\d+\.xml$/, OX + 'presentationml.slide+xml'],
    [/^ppt\/slideLayouts\/slideLayout\d+\.xml$/, OX + 'presentationml.slideLayout+xml'],
    [/^ppt\/slideMasters\/slideMaster\d+\.xml$/, OX + 'presentationml.slideMaster+xml'],
    [/^ppt\/notesSlides\/notesSlide\d+\.xml$/, OX + 'presentationml.notesSlide+xml'],
    [/^ppt\/notesMasters\/notesMaster\d+\.xml$/, OX + 'presentationml.notesMaster+xml'],
    [/^ppt\/handoutMasters\/handoutMaster\d+\.xml$/, OX + 'presentationml.handoutMaster+xml'],
    [/^ppt\/presProps\.xml$/, OX + 'presentationml.presProps+xml'],
    [/^ppt\/viewProps\.xml$/, OX + 'presentationml.viewProps+xml'],
    [/^ppt\/tableStyles\.xml$/, OX + 'presentationml.tableStyles+xml'],
    [/^ppt\/commentAuthors\.xml$/, OX + 'presentationml.commentAuthors+xml'],
    [/^ppt\/comments\/comment\d+\.xml$/, OX + 'presentationml.comments+xml'],
    [/^word\/document\.xml$/, OX + 'wordprocessingml.document.main+xml'],
    [/^word\/styles\.xml$/, OX + 'wordprocessingml.styles+xml'],
    [/^word\/settings\.xml$/, OX + 'wordprocessingml.settings+xml'],
    [/^word\/webSettings\.xml$/, OX + 'wordprocessingml.webSettings+xml'],
    [/^word\/fontTable\.xml$/, OX + 'wordprocessingml.fontTable+xml'],
    [/^word\/numbering\.xml$/, OX + 'wordprocessingml.numbering+xml'],
    [/^word\/footnotes\.xml$/, OX + 'wordprocessingml.footnotes+xml'],
    [/^word\/endnotes\.xml$/, OX + 'wordprocessingml.endnotes+xml'],
    [/^word\/comments\.xml$/, OX + 'wordprocessingml.comments+xml'],
    [/^word\/header\d+\.xml$/, OX + 'wordprocessingml.header+xml'],
    [/^word\/footer\d+\.xml$/, OX + 'wordprocessingml.footer+xml'],
    [/^xl\/workbook\.xml$/, OX + 'spreadsheetml.sheet.main+xml'],
    [/^xl\/worksheets\/sheet\d+\.xml$/, OX + 'spreadsheetml.worksheet+xml'],
    [/^xl\/sharedStrings\.xml$/, OX + 'spreadsheetml.sharedStrings+xml'],
    [/^xl\/styles\.xml$/, OX + 'spreadsheetml.styles+xml'],
    [/^xl\/calcChain\.xml$/, OX + 'spreadsheetml.calcChain+xml'],
    [/^xl\/tables\/table\d+\.xml$/, OX + 'spreadsheetml.table+xml'],
    [/^xl\/comments\d+\.xml$/, OX + 'spreadsheetml.comments+xml'],
    [/^xl\/drawings\/drawing\d+\.xml$/, OX + 'drawing+xml'],
    [/\/theme\/theme\d+\.xml$/, OX + 'theme+xml'],
    [/\/charts\/chart\d+\.xml$/, OX + 'drawingml.chart+xml'],
    [/^docProps\/core\.xml$/, 'application/vnd.openxmlformats-package.core-properties+xml'],
    [/^docProps\/app\.xml$/, OX + 'extended-properties+xml'],
    [/^docProps\/custom\.xml$/, OX + 'custom-properties+xml']
  ];
  const DEFAULTS = { rels: 'application/vnd.openxmlformats-package.relationships+xml', xml: 'application/xml', png: 'image/png', jpeg: 'image/jpeg', jpg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', tif: 'image/tiff', tiff: 'image/tiff', svg: 'image/svg+xml', emf: 'image/x-emf', wmf: 'image/x-wmf', wdp: 'image/vnd.ms-photo', mp4: 'video/mp4', m4a: 'audio/mp4', mp3: 'audio/mpeg', wav: 'audio/wav', bin: 'application/vnd.openxmlformats-officedocument.oleObject', xlsx: OX + 'spreadsheetml.sheet', vml: OX + 'vmlDrawing', fntdata: 'application/x-fontdata', odttf: 'application/vnd.openxmlformats-officedocument.obfuscatedFont' };
  // parts that can be dropped without losing the document
  const OPTIONAL = /^(docProps\/|customXml\/|.*\/printerSettings\/|.*\/comments|ppt\/commentAuthors|ppt\/notesSlides\/|ppt\/notesMasters\/|ppt\/handoutMasters\/|xl\/calcChain\.xml|ppt\/viewProps\.xml|ppt\/presProps\.xml|word\/webSettings\.xml|.*\/thumbnail\.|docProps\/thumbnail)/;
  const PLACEHOLDER_PNG = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg=='), c => c.charCodeAt(0));

  // friendly names for parts in messages
  const nice = n => { let m;
    if ((m = /^ppt\/slides\/slide(\d+)\.xml$/.exec(n))) return `שקף ${m[1]}`;
    if ((m = /^xl\/worksheets\/sheet(\d+)\.xml$/.exec(n))) return `גיליון ${m[1]}`;
    if (n === 'word/document.xml') return 'גוף המסמך';
    if (/\/media\//.test(n)) return `התמונה ${n.split('/').pop()}`;
    if (/notesSlide/.test(n)) return 'הערות הדובר';
    if (/^docProps\//.test(n)) return 'פרטי הקובץ';
    return n; };
  const extOf = n => (n.split('.').pop() || '').toLowerCase();
  const dirOf = p => p.includes('/') ? p.slice(0, p.lastIndexOf('/') + 1) : '';
  function resolve(base, target) { // target of a relationship, relative to its source folder
    if (target.startsWith('/')) return target.slice(1);
    const parts = (base + target).split('/'), out = [];
    for (const s of parts) { if (s === '..') out.pop(); else if (s !== '.' && s !== '') out.push(s); }
    return out.join('/');
  }
  const relsOf = part => dirOf(part) + '_rels/' + part.slice(dirOf(part).length) + '.rels';
  const sourceOfRels = r => { const m = /^(.*?)_rels\/(.+)\.rels$/.exec(r); return m ? m[1] + m[2] : null; };
  function concat(chunks) { const n = chunks.reduce((t, c) => t + c.length, 0), o = new Uint8Array(n); let p = 0; for (const c of chunks) { o.set(c, p); p += c.length; } return o; }

  async function inflateRaw(data) {
    const ds = new DecompressionStream('deflate-raw'); const w = ds.writable.getWriter();
    w.write(data).catch(() => {}); w.close().catch(() => {});
    const r = ds.readable.getReader(), chunks = [];
    try { for (;;) { const { done, value } = await r.read(); if (done) break; chunks.push(value); } }
    catch (e) { if (!chunks.length) throw e; } // truncated data: keep what came out
    return concat(chunks);
  }
  const findSig = (b, s, from) => { for (let i = from; i < b.length - 3; i++) if (b[i] === 0x50 && b[i + 1] === 0x4b && b[i + 2] === s[0] && b[i + 3] === s[1]) return i; return -1; };
  // read the archive entry by entry, without the (possibly missing) central directory
  async function salvage(bytes, log) {
    const files = new Map(); const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let p = findSig(bytes, [3, 4], 0);
    while (p >= 0 && p + 30 <= bytes.length) {
      const flags = dv.getUint16(p + 6, true), method = dv.getUint16(p + 8, true);
      const csize = dv.getUint32(p + 18, true), nlen = dv.getUint16(p + 26, true), xlen = dv.getUint16(p + 28, true);
      const name = dec.decode(bytes.subarray(p + 30, p + 30 + nlen)); const start = p + 30 + nlen + xlen;
      let end;
      if (csize > 0 && !(flags & 8) && start + csize <= bytes.length) end = start + csize;
      else { const cands = [findSig(bytes, [3, 4], start), findSig(bytes, [1, 2], start), findSig(bytes, [7, 8], start)].filter(x => x >= 0); end = cands.length ? Math.min(...cands) : bytes.length; }
      if (name && !name.endsWith('/')) {
        try { const raw = bytes.subarray(start, end); files.set(name, method === 8 ? await inflateRaw(raw) : raw.slice()); }
        catch (e) { log.push({ lvl: 'warn', msg: `${nice(name)}: לא ניתן לשחזר` }); }
      }
      p = findSig(bytes, [3, 4], Math.max(end, p + 4));
    }
    return files;
  }
  async function readZip(bytes, log) {
    try {
      const z = await JSZip.loadAsync(bytes); const files = new Map();
      for (const f of Object.values(z.files)) {
        if (f.dir) continue;
        try { files.set(f.name, await f.async('uint8array')); }
        catch (e) { log.push({ lvl: 'warn', msg: `${nice(f.name)}: פגום ולא ניתן לקריאה` }); }
      }
      return { files, salvaged: false };
    } catch (e) {
      log.push({ lvl: 'err', msg: 'מבנה הקובץ שבור (לרוב: הקובץ לא ירד או לא נשמר עד הסוף). משחזר את החלקים אחד-אחד.' });
      if (typeof window.fileSig === 'function') log.push({ lvl: 'warn', msg: 'מה יש בקובץ: ' + window.fileSig(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)) });
      const files = await salvage(bytes, log);
      return { files, salvaged: true };
    }
  }

  const parseXml = t => { const d = new DOMParser().parseFromString(t, 'application/xml'); return d.getElementsByTagName('parsererror').length ? null : d; };
  // try to make broken XML well-formed again
  function fixXml(t) {
    const steps = [
      s => s.replace(/^\uFEFF/, '').replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]/g, ''),
      s => s.replace(/&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/g, '&amp;'),
      s => closeTags(s)
    ];
    let cur = t;
    for (const f of steps) { cur = f(cur); if (parseXml(cur)) return cur; }
    return null;
  }
  function closeTags(s) { // a part cut off in the middle: drop the broken tail, close what is open
    const lastGt = s.lastIndexOf('>'); if (lastGt < 0) return s; s = s.slice(0, lastGt + 1);
    const stack = []; const re = /<(\/?)([A-Za-z_][\w:.-]*)[^>]*?(\/?)>|<\?[^>]*\?>|<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>/g; let m;
    while ((m = re.exec(s))) { if (!m[2]) continue; if (m[1]) { if (stack[stack.length - 1] === m[2]) stack.pop(); } else if (!m[3]) stack.push(m[2]); }
    return s + stack.reverse().map(n => `</${n}>`).join('');
  }
  function kindOf(name, files) {
    const e = extOf(name); if (KIND[e]) return e;
    if (files.has('ppt/presentation.xml')) return 'pptx'; if (files.has('word/document.xml')) return 'docx'; if (files.has('xl/workbook.xml')) return 'xlsx';
    return null;
  }

  /* ---------------- inspect + repair ---------------- */
  async function repair(bytes, fileName) {
    const log = []; const { files, salvaged } = await readZip(bytes, log);
    const kind = kindOf(fileName, files);
    const res = { kind, log, files, changed: salvaged, fatal: false, slidesRemoved: [] };
    if (!kind) { log.push({ lvl: 'err', msg: 'זה לא נראה כמו קובץ Word, Excel או PowerPoint.' }); res.fatal = true; return res; }
    if (!files.size) { log.push({ lvl: 'err', msg: 'לא נמצא בקובץ שום חלק שאפשר לשחזר.' }); res.fatal = true; return res; }
    const removed = new Set();
    const removePart = (p, why) => { if (files.delete(p)) { removed.add(p); res.changed = true; if (why) log.push({ lvl: 'fix', msg: why }); } const r = relsOf(p); if (files.delete(r)) removed.add(r); };

    // 1. every XML part must be well-formed
    for (const [name, data] of [...files]) {
      if (!/\.(xml|rels)$/i.test(name)) continue;
      const text = dec.decode(data);
      if (parseXml(text)) continue;
      const fixed = fixXml(text);
      if (fixed) { files.set(name, enc.encode(fixed)); res.changed = true; log.push({ lvl: 'fix', msg: `${nice(name)}: תוקנו תווים פגומים` }); continue; }
      if (name === KIND[kind].main) { log.push({ lvl: 'err', msg: `החלק הראשי של הקובץ (${name}) פגום מכדי לתקן` }); res.fatal = true; continue; }
      if (/^ppt\/slides\/slide\d+\.xml$/.test(name)) { const n = +name.match(/slide(\d+)\.xml/)[1]; res.slidesRemoved.push(n); removePart(name, `שקף ${n} פגום מכדי לתקן והוסר`); continue; }
      if (/\.rels$/.test(name)) { files.delete(name); res.changed = true; log.push({ lvl: 'fix', msg: `קובץ קישורים פגום הוסר: ${name}` }); continue; }
      removePart(name, OPTIONAL.test(name) ? `${nice(name)}: פגום והוסר (לא משפיע על התוכן)` : `${nice(name)}: פגום והוסר`);
    }
    // 2. relationships that point to parts that are gone
    for (const [rname, data] of [...files]) {
      if (!rname.endsWith('.rels')) continue;
      const doc = parseXml(dec.decode(data)); if (!doc) continue;
      const src = sourceOfRels(rname) ?? ''; const base = dirOf(src);
      let dirty = false;
      for (const rel of [...doc.getElementsByTagName('Relationship')]) {
        if (rel.getAttribute('TargetMode') === 'External') continue;
        const tgt = resolve(rname.startsWith('_rels/') ? '' : base, rel.getAttribute('Target') || '');
        if (files.has(tgt)) continue;
        const type = rel.getAttribute('Type') || '';
        if (/\/image$/.test(type) && !removed.has(tgt)) { files.set(tgt, PLACEHOLDER_PNG); res.changed = true; log.push({ lvl: 'fix', msg: `${nice(tgt)} חסרה, והוחלפה בתמונה ריקה` }); continue; }
        rel.parentNode.removeChild(rel); dirty = true;
        if (!removed.has(tgt)) log.push({ lvl: 'fix', msg: `${nice(tgt)} לא נמצא בקובץ, והקישור אליו הוסר` });
      }
      if (dirty) { files.set(rname, enc.encode(new XMLSerializer().serializeToString(doc))); res.changed = true; }
    }
    // 3. presentation: slide list must match the slides that exist
    if (kind === 'pptx' && files.has('ppt/presentation.xml')) {
      const rels = parseXml(dec.decode(files.get('ppt/_rels/presentation.xml.rels') || new Uint8Array()));
      const pres = parseXml(dec.decode(files.get('ppt/presentation.xml')));
      if (pres) {
        const ok = new Set(rels ? [...rels.getElementsByTagName('Relationship')].map(r => r.getAttribute('Id')) : []);
        let dirty = false;
        for (const tag of ['p:sldId', 'p:notesMasterId', 'p:handoutMasterId']) for (const el of [...pres.getElementsByTagName(tag)]) {
          const rid = el.getAttribute('r:id'); if (rid && !ok.has(rid)) { el.parentNode.removeChild(el); dirty = true; }
        }
        if (dirty) { files.set('ppt/presentation.xml', enc.encode(new XMLSerializer().serializeToString(pres))); res.changed = true; log.push({ lvl: 'fix', msg: 'רשימת השקפים עודכנה לפי השקפים שקיימים' }); }
        if (!pres.getElementsByTagName('p:sldId').length) log.push({ lvl: 'warn', msg: 'במצגת לא נשארו שקפים תקינים' });
      }
    }
    // 4. content types: rebuild when missing/broken, otherwise complete it
    let ctDoc = files.has('[Content_Types].xml') ? parseXml(dec.decode(files.get('[Content_Types].xml'))) : null;
    if (!ctDoc) { log.push({ lvl: 'fix', msg: 'רשימת סוגי החלקים ([Content_Types].xml) נבנתה מחדש' }); res.changed = true; ctDoc = parseXml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'); }
    const types = ctDoc.documentElement, NS = types.namespaceURI;
    const defaults = new Set([...types.getElementsByTagName('Default')].map(d => (d.getAttribute('Extension') || '').toLowerCase()));
    const overrides = new Map([...types.getElementsByTagName('Override')].map(o => [(o.getAttribute('PartName') || '').replace(/^\//, ''), o]));
    let ctDirty = false;
    for (const [p, o] of overrides) if (!files.has(p)) { types.removeChild(o); overrides.delete(p); ctDirty = true; }
    for (const p of files.keys()) {
      if (p === '[Content_Types].xml' || overrides.has(p)) continue;
      const hit = CT.find(([re]) => re.test(p));
      if (hit) { const o = ctDoc.createElementNS(NS, 'Override'); o.setAttribute('PartName', '/' + p); o.setAttribute('ContentType', hit[1]); types.appendChild(o); overrides.set(p, o); ctDirty = true; continue; }
      const e = extOf(p);
      if (!defaults.has(e) && DEFAULTS[e]) { const d = ctDoc.createElementNS(NS, 'Default'); d.setAttribute('Extension', e); d.setAttribute('ContentType', DEFAULTS[e]); types.insertBefore(d, types.firstChild); defaults.add(e); ctDirty = true; }
    }
    if (ctDirty) { res.changed = true; if (!log.some(l => l.msg.includes('Content_Types'))) log.push({ lvl: 'fix', msg: 'רשימת סוגי החלקים הושלמה' }); }
    files.set('[Content_Types].xml', enc.encode(new XMLSerializer().serializeToString(ctDoc)));
    // 5. package relationships
    if (!files.has('_rels/.rels')) {
      files.set('_rels/.rels', enc.encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="${KIND[kind].main}"/></Relationships>`));
      res.changed = true; log.push({ lvl: 'fix', msg: 'קובץ הקישורים הראשי נבנה מחדש' });
    }
    if (!res.changed && !log.length) log.push({ lvl: 'ok', msg: 'מבנה הקובץ תקין. לא נמצאו חלקים פגומים.' });
    return res;
  }
  async function build(res) { // a fresh, clean archive; [Content_Types].xml first
    const z = new JSZip();
    z.file('[Content_Types].xml', res.files.get('[Content_Types].xml'));
    for (const [n, d] of res.files) if (n !== '[Content_Types].xml') z.file(n, d);
    return z.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 }, mimeType: KIND[res.kind].mime });
  }

  /* ---------------- content: what is inside, for preview and extraction ---------------- */
  const txt = el => el ? el.textContent : '';
  const byNum = (a, b) => (+a.match(/(\d+)\.xml$/)[1]) - (+b.match(/(\d+)\.xml$/)[1]);
  function relMap(files, part) {
    const d = parseXml(dec.decode(files.get(relsOf(part)) || new Uint8Array())); const m = {};
    if (d) for (const r of d.getElementsByTagName('Relationship')) m[r.getAttribute('Id')] = { t: resolve(dirOf(part), r.getAttribute('Target') || ''), type: r.getAttribute('Type') || '', ext: r.getAttribute('TargetMode') === 'External' };
    return m;
  }
  function slideOrder(files) {
    const pres = parseXml(dec.decode(files.get('ppt/presentation.xml') || new Uint8Array()));
    const rm = relMap(files, 'ppt/presentation.xml'); const order = [];
    if (pres) for (const s of pres.getElementsByTagName('p:sldId')) { const r = rm[s.getAttribute('r:id')]; if (r && files.has(r.t)) order.push(r.t); }
    const rest = [...files.keys()].filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n) && !order.includes(n)).sort(byNum);
    return order.concat(rest);
  }
  // text from raw XML even when it can't be parsed (last resort)
  const rawRuns = (s, tag) => [...s.matchAll(new RegExp(`<${tag}(?:\\s[^>]*)?>([^<]*)</${tag}>`, 'g'))].map(m => m[1].replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&'));
  function content(res) {
    const f = res.files, out = { kind: res.kind, items: [] };
    if (res.kind === 'pptx') {
      for (const p of slideOrder(f)) {
        const s = dec.decode(f.get(p)); const d = parseXml(s); const rm = relMap(f, p); const paras = [], images = [];
        if (d) {
          for (const sp of d.getElementsByTagName('p:sp')) {
            const ph = sp.getElementsByTagName('p:ph')[0]; const isTitle = ph && /title|ctrTitle/.test(ph.getAttribute('type') || '');
            for (const para of sp.getElementsByTagName('a:p')) { const t = [...para.getElementsByTagName('a:t')].map(txt).join(''); if (t.trim()) paras.push({ t, title: !!isTitle }); }
          }
          for (const g of d.getElementsByTagName('a:graphicFrame')) for (const para of g.getElementsByTagName('a:p')) { const t = [...para.getElementsByTagName('a:t')].map(txt).join(''); if (t.trim()) paras.push({ t }); }
          for (const b of d.getElementsByTagName('a:blip')) { const r = rm[b.getAttribute('r:embed')]; if (r && f.has(r.t) && /\.(png|jpe?g|gif|bmp)$/i.test(r.t)) images.push(r.t); }
        } else rawRuns(s, 'a:t').forEach(t => t.trim() && paras.push({ t }));
        out.items.push({ part: p, paras, images });
      }
    } else if (res.kind === 'docx') {
      const s = dec.decode(f.get('word/document.xml') || new Uint8Array()); const d = parseXml(s); const paras = [];
      if (d) for (const p of d.getElementsByTagName('w:p')) { const t = [...p.getElementsByTagName('w:t')].map(txt).join(''); const st = p.getElementsByTagName('w:pStyle')[0]; paras.push({ t, title: !!(st && /Heading|Title|כותרת/i.test(st.getAttribute('w:val') || '')) }); }
      else rawRuns(s, 'w:t').forEach(t => paras.push({ t }));
      const images = [...f.keys()].filter(n => /^word\/media\/.*\.(png|jpe?g|gif)$/i.test(n));
      out.items.push({ part: 'word/document.xml', paras, images });
    } else if (res.kind === 'xlsx') {
      const ss = []; const sd = parseXml(dec.decode(f.get('xl/sharedStrings.xml') || new Uint8Array()));
      if (sd) for (const si of sd.getElementsByTagName('si')) ss.push([...si.getElementsByTagName('t')].map(txt).join(''));
      const wb = parseXml(dec.decode(f.get('xl/workbook.xml') || new Uint8Array())); const rm = relMap(f, 'xl/workbook.xml'); const sheets = [];
      if (wb) for (const sh of wb.getElementsByTagName('sheet')) { const r = rm[sh.getAttribute('r:id')]; if (r && f.has(r.t)) sheets.push({ name: sh.getAttribute('name'), part: r.t }); }
      for (const n of [...f.keys()].filter(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort(byNum)) if (!sheets.some(s => s.part === n)) sheets.push({ name: n.match(/(sheet\d+)/)[1], part: n });
      for (const sh of sheets) {
        const d = parseXml(dec.decode(f.get(sh.part))); const rows = [];
        if (d) for (const row of d.getElementsByTagName('row')) {
          const cells = [];
          for (const c of row.getElementsByTagName('c')) {
            const ref = c.getAttribute('r') || ''; const col = ref.replace(/\d+/g, ''); let ci = 0; for (const ch of col) ci = ci * 26 + ch.charCodeAt(0) - 64; ci = Math.max(ci - 1, cells.length);
            const t = c.getAttribute('t'); const v = txt(c.getElementsByTagName('v')[0]);
            cells[ci] = t === 's' ? (ss[+v] ?? '') : t === 'inlineStr' ? txt(c.getElementsByTagName('is')[0]) : v;
          }
          rows.push(cells);
        }
        out.items.push({ part: sh.part, name: sh.name, rows });
      }
    }
    return out;
  }

  /* ---------------- extraction into a brand-new file ---------------- */
  const isHeb = t => /[\u0590-\u05FF]/.test(t);
  async function extractToNew(res) {
    const c = content(res), f = res.files;
    if (res.kind === 'pptx') {
      const pptx = new PptxGenJS(); pptx.layout = 'LAYOUT_WIDE'; pptx.rtlMode = true;
      for (const it of c.items) {
        const s = pptx.addSlide();
        const title = it.paras.find(p => p.title) || it.paras[0]; const body = it.paras.filter(p => p !== title);
        if (title) s.addText(title.t, { x: 0.5, y: 0.35, w: 12.3, h: 0.9, fontSize: 30, bold: true, align: isHeb(title.t) ? 'right' : 'left', rtlMode: isHeb(title.t), fontFace: 'Arial' });
        const imgs = it.images.slice(0, 2); const bodyW = imgs.length ? 7.6 : 12.3;
        if (body.length) s.addText(body.map(p => ({ text: p.t, options: { breakLine: true, rtlMode: isHeb(p.t), align: isHeb(p.t) ? 'right' : 'left' } })), { x: imgs.length ? 5.2 : 0.5, y: 1.4, w: bodyW, h: 5.6, fontSize: 18, valign: 'top', fontFace: 'Arial' });
        imgs.forEach((p, k) => { const e = extOf(p); const b64 = btoa(Array.from(f.get(p), ch => String.fromCharCode(ch)).join('')); s.addImage({ data: `image/${e === 'jpg' ? 'jpeg' : e};base64,${b64}`, x: 0.5, y: 1.4 + k * 2.9, w: 4.4, h: 2.7, sizing: { type: 'contain', w: 4.4, h: 2.7 } }); });
      }
      if (!c.items.length) pptx.addSlide().addText('לא נמצא תוכן לשחזור', { x: 0.5, y: 3, w: 12.3, h: 1, align: 'center', fontSize: 24 });
      return { bytes: new Uint8Array(await pptx.write({ outputType: 'arraybuffer' })), ext: 'pptx', mime: KIND.pptx.mime };
    }
    if (res.kind === 'docx') {
      const paras = c.items[0] ? c.items[0].paras : [];
      const esc = t => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      const body = paras.map(p => { const h = isHeb(p.t); return `<w:p><w:pPr>${h ? '<w:bidi/>' : ''}${p.title ? '<w:spacing w:before="240" w:after="120"/>' : ''}</w:pPr>${p.t ? `<w:r><w:rPr>${p.title ? '<w:b/><w:bCs/><w:sz w:val="32"/><w:szCs w:val="32"/>' : ''}${h ? '<w:rtl/>' : ''}</w:rPr><w:t xml:space="preserve">${esc(p.t)}</w:t></w:r>` : ''}</w:p>`; }).join('');
      const z = new JSZip();
      z.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="' + OX + 'wordprocessingml.document.main+xml"/></Types>');
      z.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
      z.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/><w:bidi/></w:sectPr></w:body></w:document>`);
      return { bytes: await z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), ext: 'docx', mime: KIND.docx.mime };
    }
    if (res.kind === 'xlsx') {
      const esc = t => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      const colName = i => { let s = ''; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
      const sheets = c.items.length ? c.items : [{ name: 'Sheet1', rows: [] }];
      const z = new JSZip();
      z.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="' + OX + 'spreadsheetml.sheet.main+xml"/>' + sheets.map((_, k) => `<Override PartName="/xl/worksheets/sheet${k + 1}.xml" ContentType="${OX}spreadsheetml.worksheet+xml"/>`).join('') + '</Types>');
      z.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
      const used = new Set();
      const names = sheets.map((s, k) => { let n = (s.name || `Sheet${k + 1}`).replace(/[\\/?*[\]:]/g, ' ').slice(0, 31) || `Sheet${k + 1}`; while (used.has(n)) n = n.slice(0, 28) + k; used.add(n); return n; });
      z.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${names.map((n, k) => `<sheet name="${esc(n)}" sheetId="${k + 1}" r:id="rId${k + 1}"/>`).join('')}</sheets></workbook>`);
      z.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, k) => `<Relationship Id="rId${k + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${k + 1}.xml"/>`).join('')}</Relationships>`);
      sheets.forEach((s, k) => {
        const rtl = s.rows.some(r => r.some(v => v && isHeb(String(v))));
        const rows = s.rows.map((r, ri) => `<row r="${ri + 1}">${r.map((v, ci) => v === undefined || v === '' ? '' : (/^-?\d+(\.\d+)?$/.test(v) ? `<c r="${colName(ci)}${ri + 1}"><v>${v}</v></c>` : `<c r="${colName(ci)}${ri + 1}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`)).join('')}</row>`).join('');
        z.file(`xl/worksheets/sheet${k + 1}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"${rtl ? ' rightToLeft="1"' : ''}/></sheetViews><sheetData>${rows}</sheetData></worksheet>`);
      });
      return { bytes: await z.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), ext: 'xlsx', mime: KIND.xlsx.mime };
    }
    throw new Error('סוג קובץ לא נתמך');
  }
  return { KIND, repair, build, content, extractToNew, isOffice: n => /\.(docx|xlsx|pptx)$/i.test(n) };
})();
window.OFFICE = OFFICE;
