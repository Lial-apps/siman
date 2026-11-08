// Cloudflare Worker "siman-share" (5.9.4, fix 5.9.6) — documents sent from סימן for signing or viewing by a link.
//
// Paste this whole file: Cloudflare → Workers & Pages → siman-share → Edit code → select all → paste → Deploy.
// Binding (already set up): Settings → Bindings → KV namespace, Variable name DOCS → namespace SIMAN_DOCS.
//
// Links people open in a browser (WhatsApp, mail…) — no Origin needed, the secret token in the link is the key:
//   /s/<id>/<token>  a named signer      /g/<id>/<token>  one link for a whole group (everyone writes their name)
//   /v/<id>/<token>  view only           /o/<id>#k=<key>  the sender's own page
// They show sign.html of the app, served from this address together with the files it needs.
//
// The app's API (/docs/…): creating, deleting and closing need a request from the app itself.
// Everything is deleted automatically after 30 days.
// Keys: m:<id> details · f:<id> the PDF · o:<id> the original Word/Excel/PowerPoint of a view link · s:<id>:<n> what signer n did · r:<id>:<rid> one group signature
//       rl:<ip>:<day> documents sent today

const SITE = 'https://siman.siman-app.workers.dev';                       // the app (hides GitHub) — for redirects only
// 5.9.6: the signing page and its files are fetched straight from GitHub Pages.
// A worker can't fetch another worker on the same workers.dev account (Cloudflare error 1042),
// so fetching from SITE failed and every /s /v /o /g link answered "דף החתימה לא זמין כרגע".
const PAGES = 'https://lial-apps.github.io/siman';
const ALLOWED = ['https://lial-apps.github.io', 'https://siman.siman-app.workers.dev'];
const TTL = 30 * 24 * 3600;
const MAX_PDF = 15 * 1024 * 1024, MAX_PER_DAY = 40, MAX_SIGNERS = 10, MAX_FIELDS = 200, MAX_IMG = 300 * 1024;
const MAX_GROUP = 2000;                                                   // signatures on one group link
const TYPES = ['sig', 'initials', 'name', 'date', 'text', 'check'];
const NOT_FOUND = 'המסמך לא נמצא. ייתכן שהשולח מחק אותו, או שעברו 30 יום מאז שנשלח.';

const rnd = n => { const a = new Uint8Array(n); crypto.getRandomValues(a); return [...a].map(b => 'abcdefghijkmnpqrstuvwxyz23456789'[b % 32]).join(''); };
const sha = async s => [...new Uint8Array(await crypto.subtle.digest('SHA-256', typeof s === 'string' ? new TextEncoder().encode(s) : s))].map(b => b.toString(16).padStart(2, '0')).join('');
const clip = (s, n) => String(s ?? '').slice(0, n);
const num = (v, lo, hi) => { v = +v; return Number.isFinite(v) ? Math.min(hi, Math.max(lo, v)) : lo; };
const same = (a, b) => { a = String(a || ''); b = String(b || ''); if (!a || a.length !== b.length) return false; let d = 0; for (let i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i); return d === 0; };

// what a signer filled in: texts and small signature pictures only
function cleanValues(v, allowed) {
  const out = {}; let size = 0;
  for (const [id, x] of Object.entries(v && typeof v === 'object' ? v : {})) {
    if (!allowed.has(id) || !x || typeof x !== 'object') continue;
    if (typeof x.png === 'string' && /^data:image\/(png|jpeg);base64,/.test(x.png)) {
      if (x.png.length > MAX_IMG) throw new Error('החתימה גדולה מדי. אפשר לחתום שוב, קצת יותר קטן.');
      out[id] = { png: x.png }; size += x.png.length;
    } else if (x.text != null) out[id] = { text: clip(x.text, 200) };
    if (size > MAX_IMG * 6) throw new Error('יותר מדי תמונות חתימה במסמך אחד.');
  }
  return out;
}
const who = (req, tz) => ({ ip: req.headers.get('CF-Connecting-IP') || '', ua: clip(req.headers.get('User-Agent'), 300), country: (req.cf && req.cf.country) || '', tz: clip(tz, 60) });

export default {
  async fetch(req, env) {
    const url = new URL(req.url);
    const parts = url.pathname.split('/').filter(Boolean);
    const origin = req.headers.get('Origin') || '';
    const okOrigin = ALLOWED.includes(origin) || origin === url.origin;
    const cors = {
      'Access-Control-Allow-Origin': okOrigin ? origin : ALLOWED[1],
      'Access-Control-Allow-Methods': 'GET, POST, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
      'Access-Control-Max-Age': '86400', 'Vary': 'Origin'
    };
    const json = (o, status = 200) => new Response(JSON.stringify(o), { status, headers: { ...cors, 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' } });
    const fail = (msg, status = 400) => json({ error: msg }, status);
    if (req.method === 'OPTIONS') return new Response(null, { status: okOrigin ? 204 : 403, headers: cors });

    // ---------------- pages people open from a link
    if (parts[0] !== 'docs') {
      if (req.method !== 'GET' && req.method !== 'HEAD') return new Response('Method not allowed', { status: 405 });
      if (['s', 'g', 'v', 'o'].includes(parts[0]) && parts[1]) return signPage(req);
      if (parts[0] === 'go' || !parts.length) return Response.redirect(SITE + '/?ref=sign', 302);
      if (parts.length === 1 && /^[\w.-]+\.(js|css|png|jpe?g|svg|ico|woff2?|ttf|json|webmanifest|mjs|wasm)$/i.test(parts[0])) return asset(parts[0]);
      return new Response('לא נמצא', { status: 404, headers: { 'Content-Type': 'text/plain; charset=utf-8' } });
    }

    if (!env.DOCS) return fail('the worker has no DOCS binding', 500);
    const id = parts[1], action = parts[2] || '';
    const ip = req.headers.get('CF-Connecting-IP') || '';
    try {
      // ---------------- a new document (from the app only)
      if (!id && req.method === 'POST') {
        if (!ALLOWED.includes(origin)) return fail('forbidden', 403);
        const day = new Date().toISOString().slice(0, 10), rk = `rl:${await sha(ip)}:${day}`;
        const used = +(await env.DOCS.get(rk) || 0);
        if (used >= MAX_PER_DAY) return fail('יותר מדי מסמכים נשלחו היום מהחיבור הזה. אפשר לנסות שוב מחר.', 429);
        const fd = await req.formData(); let meta;
        try { meta = JSON.parse(fd.get('meta') || '{}'); } catch { return fail('bad meta'); }
        const file = fd.get('pdf');
        if (!file || typeof file === 'string') return fail('no pdf');
        const pdf = await file.arrayBuffer();
        if (pdf.byteLength > MAX_PDF) return fail('הקובץ גדול מ-15MB');
        if (new TextDecoder().decode(new Uint8Array(pdf, 0, 5)) !== '%PDF-') return fail('הקובץ אינו PDF');
        const mode = ['view', 'open'].includes(meta.mode) ? meta.mode : 'sign';
        // 5.9.6: a view link may also keep the original Word / Excel / PowerPoint file, so the one who opens it can choose
        let orig = null;
        const of = fd.get('orig');
        if (mode === 'view' && of && typeof of !== 'string') {
          const ext = (String(of.name || '').match(/\.(docx|xlsx|pptx)$/i) || [])[1];
          const ob = await of.arrayBuffer();
          if (!ext) return fail('סוג הקובץ המקורי לא נתמך');
          if (ob.byteLength > MAX_PDF) return fail('הקובץ המקורי גדול מ-15MB');
          if (new TextDecoder().decode(new Uint8Array(ob, 0, 2)) !== 'PK') return fail('הקובץ המקורי פגום');
          orig = { bytes: ob, name: clip(of.name, 140), ext: ext.toLowerCase(), size: ob.byteLength };
        }
        let signers = (Array.isArray(meta.signers) ? meta.signers : []).slice(0, MAX_SIGNERS);
        if (mode === 'open') signers = [{ name: 'חותם' }];
        if (mode === 'view') signers = [];
        if (mode === 'sign' && !signers.length) return fail('אין חותמים');
        signers = signers.map((s, i) => ({ n: i, name: clip(s && s.name, 60).trim() || `חותם ${i + 1}`, token: rnd(20) }));
        const fields = mode === 'view' ? [] : (Array.isArray(meta.fields) ? meta.fields : []).slice(0, MAX_FIELDS)
          .filter(f => f && TYPES.includes(f.type) && Number.isInteger(+f.signer) && +f.signer >= 0 && +f.signer < signers.length)
          .map((f, i) => ({ id: 'f' + i, signer: +f.signer, page: num(f.page, 0, 9999) | 0, type: f.type,
            x: num(f.x, 0, 1), y: num(f.y, 0, 1), w: num(f.w, 0, 1), h: num(f.h, 0, 1),
            r: Array.isArray(f.r) ? f.r.slice(0, 4).map(v => num(v, -1e5, 1e5)) : null,
            req: f.type !== 'check' && f.req !== false, label: f.type === 'text' ? clip(f.label, 40) : '' }));
        if (mode !== 'view' && !fields.length) return fail('לא סומן איפה חותמים');
        const newId = rnd(12), owner = rnd(24), view = rnd(16), now = Date.now();
        const doc = {
          v: 2, id: newId, mode, title: clip(meta.title, 120) || 'מסמך', note: clip(meta.note, 500),
          from: { name: clip(meta.from && meta.from.name, 80), mail: clip(meta.from && meta.from.mail, 120) },
          order: meta.order === 'seq' ? 'seq' : 'any', pages: num(meta.pages, 1, 9999) | 0,
          created: new Date(now).toISOString(), expires: new Date(now + TTL * 1000).toISOString(),
          hash: await sha(pdf), ownerHash: await sha(owner), view, signers, fields,
          open: mode === 'open' ? { ask: clip(meta.ask, 60), closed: false } : null,
          orig: orig ? { name: orig.name, ext: orig.ext, size: orig.size } : null
        };
        await env.DOCS.put('f:' + newId, pdf, { expirationTtl: TTL });
        if (orig) await env.DOCS.put('o:' + newId, orig.bytes, { expirationTtl: TTL });
        await env.DOCS.put('m:' + newId, JSON.stringify(doc), { expirationTtl: TTL });
        await env.DOCS.put(rk, String(used + 1), { expirationTtl: 2 * 24 * 3600 });
        return json({ id: newId, owner, view, expires: doc.expires, signers: signers.map(s => ({ n: s.n, name: s.name, token: s.token })) });
      }
      if (!id || !/^[a-z0-9]{6,40}$/.test(id)) return fail('not found', 404);

      const raw = await env.DOCS.get('m:' + id);
      if (!raw) return fail(NOT_FOUND, 404);
      const doc = JSON.parse(raw);
      if (doc.mode === 'open' && !doc.open) doc.open = { ask: '', closed: false };
      const q = url.searchParams;
      const isOwner = q.get('k') ? same(await sha(q.get('k')), doc.ownerHash) : false;
      const signer = q.get('t') && doc.mode === 'sign' ? doc.signers.find(s => same(s.token, q.get('t'))) : null;
      const group = q.get('g') && doc.mode === 'open' && same(doc.signers[0].token, q.get('g'));
      const viewer = q.get('v') ? same(doc.view, q.get('v')) : false;
      const ttlLeft = () => Math.max(60, Math.round((Date.parse(doc.expires) - Date.now()) / 1000));

      // what every signer did, each in its own key (two people signing at once never overwrite each other)
      const states = async () => {
        if (doc.mode !== 'sign') return [];
        return Promise.all(doc.signers.map(async s => { const x = await env.DOCS.get(`s:${id}:${s.n}`); return x ? JSON.parse(x) : null; }));
      };
      const responses = async (withValues) => {
        const out = []; let cursor;
        do {
          const l = await env.DOCS.list({ prefix: `r:${id}:`, cursor });
          for (const k of l.keys) {
            if (withValues) { const x = await env.DOCS.get(k.name); if (x) out.push(JSON.parse(x)); }
            else if (k.metadata) out.push(k.metadata);
          }
          cursor = l.list_complete ? null : l.cursor;
        } while (cursor && out.length < MAX_GROUP);
        return out.sort((a, b) => String(a.at).localeCompare(String(b.at)));
      };
      // the details as the one asking may see them
      const view = async (me, extra) => {
        const st = await states();
        const showValues = isOwner || !!signer;              // a viewer and the group never see other people's signatures
        const out = {
          id: doc.id, mode: doc.mode, title: doc.title, note: doc.note, from: doc.from, order: doc.order, pages: doc.pages,
          created: doc.created, expires: doc.expires, hash: doc.hash, owner: isOwner, me,
          orig: doc.orig ? { name: doc.orig.name, ext: doc.orig.ext, size: doc.orig.size } : null,
          fields: doc.fields,
          signers: doc.mode === 'open' ? [{ n: 0, name: 'חותם', status: 'waiting' }] : doc.signers.map((s, i) => {
            const x = st[i] || {};
            const o = { n: s.n, name: s.name, status: x.status || 'waiting' };
            if (x.at) o.at = x.at;
            if (isOwner || (signer && signer.n === s.n)) Object.assign(o, { ip: x.ip, ua: x.ua, country: x.country, tz: x.tz, reason: x.reason });
            if (showValues && x.status === 'signed') o.values = x.values;
            return o;
          })
        };
        if (doc.mode === 'open') { out.open = { ask: doc.open.ask, closed: !!doc.open.closed }; if (isOwner) out.responses = await responses(false); }
        return Object.assign(out, extra || {});
      };

      // ---------------- reading
      if (req.method === 'GET' && !action) {
        if (q.get('g') && q.get('r') && group) {            // one person's own signed copy of a group link
          const x = await env.DOCS.get(`r:${id}:${q.get('r')}`);
          if (!x) return fail('החתימה לא נמצאה', 404);
          const r = JSON.parse(x);
          const base = await view(0);
          base.signers = [{ n: 0, name: r.label || r.name, status: 'signed', at: r.at, values: r.values, ip: r.ip, ua: r.ua, country: r.country, tz: r.tz }];
          return json(base);
        }
        if (!(isOwner || signer || group || viewer)) return fail('הקישור לא תקין', 403);
        return json(await view(signer ? signer.n : group ? 0 : null));
      }
      if (req.method === 'GET' && action === 'pdf') {
        if (!(isOwner || signer || group || viewer)) return fail('הקישור לא תקין', 403);
        const pdf = await env.DOCS.get('f:' + id, 'arrayBuffer');
        if (!pdf) return fail(NOT_FOUND, 404);
        const name = encodeURIComponent(doc.title.replace(/[\\/:*?"<>|]+/g, ' ') + '.pdf');
        return new Response(pdf, { headers: { ...cors, 'Content-Type': 'application/pdf', 'Content-Disposition': `inline; filename*=UTF-8''${name}`, 'Cache-Control': 'private, no-store' } });
      }
      if (req.method === 'GET' && action === 'orig') {        // the original Word / Excel / PowerPoint file of a view link
        if (!(isOwner || viewer)) return fail('הקישור לא תקין', 403);
        const ob = doc.orig ? await env.DOCS.get('o:' + id, 'arrayBuffer') : null;
        if (!ob) return fail(NOT_FOUND, 404);
        const types = { docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' };
        return new Response(ob, { headers: { ...cors, 'Content-Type': types[doc.orig.ext] || 'application/octet-stream', 'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(doc.orig.name)}`, 'Cache-Control': 'private, no-store' } });
      }
      if (req.method === 'GET' && action === 'responses') {
        if (!isOwner) return fail('forbidden', 403);
        return json({ responses: await responses(true) });
      }

      // ---------------- changes: from the app or from the signing page only
      if (!okOrigin) return fail('forbidden', 403);
      let body = {};
      if (req.method === 'POST') { try { body = await req.json(); } catch { body = {}; } }

      if (req.method === 'POST' && action === 'sign' && group) {
        if (doc.open.closed) return fail('הקישור נסגר, ואי אפשר לחתום עליו יותר.', 409);
        if (!body.agree) return fail('צריך לאשר את החתימה');
        const name = clip(body.name, 80).trim();
        if (name.length < 2) return fail('צריך לכתוב את השם המלא');
        const extra = clip(body.extra, 120).trim();
        if (doc.open.ask && !extra) return fail(`צריך למלא: ${doc.open.ask}`);
        const values = cleanValues(body.values, new Set(doc.fields.map(f => f.id)));
        const missing = doc.fields.filter(f => f.req && f.type !== 'name' && !values[f.id]);
        if (missing.length) return fail('יש שדות שעוד לא מולאו');
        for (const f of doc.fields) if (f.type === 'name') values[f.id] = { text: name };
        const count = (await env.DOCS.list({ prefix: `r:${id}:`, limit: 1000 })).keys.length;
        if (count >= MAX_GROUP) return fail('הגעתם למספר החתימות המרבי בקישור הזה', 429);
        const rid = rnd(14), at = new Date().toISOString();
        const label = extra ? `${name} · ${extra}` : name;
        const rec = { rid, name, extra, label, at, values, ...who(req, body.tz) };
        await env.DOCS.put(`r:${id}:${rid}`, JSON.stringify(rec), { expirationTtl: ttlLeft(), metadata: { rid, name: clip(name, 80), extra: clip(extra, 120), at } });
        const out = await view(0, { rid });
        out.signers = [{ n: 0, name: label, status: 'signed', at, values }];
        return json(out);
      }
      if (req.method === 'POST' && (action === 'sign' || action === 'decline') && signer) {
        const st = await states();
        const mine = st[signer.n];
        if (mine && mine.status === 'signed') return fail('כבר חתמת על המסמך הזה', 409);
        if (mine && mine.status === 'declined') return fail('כבר סירבת לחתום על המסמך הזה', 409);
        if (doc.order === 'seq' && doc.signers.some((s, i) => i < signer.n && !(st[i] && st[i].status === 'signed'))) return fail('עוד לא תורך לחתום', 409);
        let rec;
        if (action === 'sign') {
          if (!body.agree) return fail('צריך לאשר את החתימה');
          const own = doc.fields.filter(f => f.signer === signer.n);
          const values = cleanValues(body.values, new Set(own.map(f => f.id)));
          if (own.some(f => f.req && !values[f.id])) return fail('יש שדות שעוד לא מולאו');
          rec = { status: 'signed', at: new Date().toISOString(), values, ...who(req, body.tz) };
        } else rec = { status: 'declined', at: new Date().toISOString(), reason: clip(body.reason, 500), ...who(req, body.tz) };
        await env.DOCS.put(`s:${id}:${signer.n}`, JSON.stringify(rec), { expirationTtl: ttlLeft() });
        return json(await view(signer.n));
      }
      if (req.method === 'POST' && action === 'close' && isOwner) {
        if (doc.mode !== 'open') return fail('רק קישור לקבוצה נסגר');
        doc.open.closed = body.closed !== false;
        await env.DOCS.put('m:' + id, JSON.stringify(doc), { expirationTtl: ttlLeft() });
        return json({ ok: true, closed: doc.open.closed });
      }
      if (req.method === 'DELETE' && !action && isOwner) {
        const keys = ['m:' + id, 'f:' + id, 'o:' + id, ...doc.signers.map(s => `s:${id}:${s.n}`)];
        let cursor;
        do { const l = await env.DOCS.list({ prefix: `r:${id}:`, cursor }); keys.push(...l.keys.map(k => k.name)); cursor = l.list_complete ? null : l.cursor; } while (cursor);
        for (let i = 0; i < keys.length; i += 50) await Promise.all(keys.slice(i, i + 50).map(k => env.DOCS.delete(k)));
        return json({ ok: true });
      }
      return fail('הקישור לא תקין', 403);
    } catch (e) {
      return fail(e && e.message ? e.message : 'שגיאה בשרת', 500);
    }
  }
};

// the signing page of the app, from this address; files it asks for are fetched from here too
async function signPage(req) {
  const r = await fetch(PAGES + '/sign.html', { cf: { cacheEverything: true, cacheTtlByStatus: { '200-299': 300, '300-599': 0 } } });
  if (!r.ok) return new Response('דף החתימה לא זמין כרגע. נסו שוב בעוד דקה.', { status: 502, headers: { 'Content-Type': 'text/plain; charset=utf-8' } });
  let html = await r.text();
  html = html.replace(/<head([^>]*)>/i, '<head$1><base href="/">');
  // 6.2: the preview WhatsApp / mail show for the link matches what the link is. A view link is not "a document
  // to sign": it gets its own title, text and picture. (Link previews read these tags without running the page.)
  const kind = new URL(req.url).pathname.split('/')[1];
  const OG = {
    v: ['שיתפו איתך מסמך', 'פותחים בלחיצה וקוראים, גם מהטלפון. בלי להירשם ובלי להתקין.', 'og-view.jpg'],
    g: ['יש לך מסמך לחתימה', 'כותבים את השם וחותמים, גם מהטלפון. בלי להירשם ובלי להתקין.', 'og-sign.png'],
    o: ['המסמך ששלחת · מי כבר חתם', 'רואים מי פתח ומי חתם, ומורידים את המסמך החתום.', 'og-sign.png'],
  }[kind];
  if (OG) {
    const [t, d, img] = OG;
    html = html.replace(/(<meta property="og:title" content=")[^"]*/, '$1' + t)
      .replace(/(<meta property="og:description" content=")[^"]*/, '$1' + d)
      .replace(/(<meta property="og:image" content="[^"]*\/)og-sign\.png/, '$1' + img)
      .replace(/(<meta name="description" content=")[^"]*/, '$1' + d)
      .replace(/<title>[^<]*<\/title>/, `<title>${t} · סימן</title>`);
  }
  return new Response(req.method === 'HEAD' ? null : html, { headers: {
    'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store',
    'Referrer-Policy': 'no-referrer', 'X-Robots-Tag': 'noindex', 'X-Content-Type-Options': 'nosniff' } });
}
async function asset(name) {
  const r = await fetch(PAGES + '/' + name, { cf: { cacheEverything: true, cacheTtlByStatus: { '200-299': 3600, '300-599': 0 } } });
  if (!r.ok) return new Response('לא נמצא', { status: 404 });
  const h = new Headers();
  for (const k of ['Content-Type', 'Content-Length', 'ETag', 'Last-Modified']) { const v = r.headers.get(k); if (v) h.set(k, v); }
  h.set('Cache-Control', 'public, max-age=3600');
  return new Response(r.body, { headers: h });
}
