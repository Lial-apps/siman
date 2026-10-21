// Cloudflare Worker "siman-share" — documents sent from סימן for signing or viewing by a link
//
// Setup (one time, ~5 minutes):
// 1. Cloudflare → Storage & Databases → KV → Create namespace: SIMAN_DOCS.
// 2. Workers & Pages → Create → Worker → name: siman-share → Deploy → Edit code → paste this file → Deploy.
// 3. The worker → Settings → Bindings → Add → KV namespace → Variable name: DOCS → namespace: SIMAN_DOCS.
//
// Every document and its signatures are deleted automatically after 30 days.
// Keys:  m:<id>  the document's details      f:<id>  the original PDF
//        s:<id>:<n>  what signer n signed      rl:<ip>:<day>  how many documents this address sent today

const ALLOWED = ['https://lial-apps.github.io'];
const TTL = 30 * 24 * 3600;           // seconds
const MAX_PDF = 15 * 1024 * 1024;     // bytes
const MAX_PER_DAY = 40;               // new documents per address per day
const MAX_SIGNERS = 10, MAX_FIELDS = 200, MAX_IMG = 300 * 1024;

const rnd = n => { const a = new Uint8Array(n); crypto.getRandomValues(a); return [...a].map(b => 'abcdefghijkmnpqrstuvwxyz23456789'[b % 32]).join(''); };
const sha = async s => [...new Uint8Array(await crypto.subtle.digest('SHA-256', typeof s === 'string' ? new TextEncoder().encode(s) : s))].map(b => b.toString(16).padStart(2, '0')).join('');
const clip = (s, n) => String(s ?? '').slice(0, n);

export default {
  async fetch(req, env) {
    const url = new URL(req.url);
    const origin = req.headers.get('Origin') || '';
    const cors = {
      'Access-Control-Allow-Origin': ALLOWED.includes(origin) ? origin : ALLOWED[0],
      'Access-Control-Allow-Methods': 'GET, POST, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
      'Access-Control-Max-Age': '86400',
      'Vary': 'Origin'
    };
    const json = (o, status = 200) => new Response(JSON.stringify(o), { status, headers: { ...cors, 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' } });
    const fail = (msg, status = 400) => json({ error: msg }, status);
    if (req.method === 'OPTIONS') return new Response(null, { headers: cors });
    if (!ALLOWED.includes(origin)) return fail('forbidden', 403);
    if (!env.DOCS) return fail('the worker has no DOCS binding', 500);

    const parts = url.pathname.split('/').filter(Boolean); // ['docs', id, action]
    if (parts[0] !== 'docs') return fail('not found', 404);
    const ip = req.headers.get('CF-Connecting-IP') || '';

    try {
      // ---------- a new document
      if (parts.length === 1 && req.method === 'POST') {
        const day = new Date().toISOString().slice(0, 10);
        const rk = `rl:${await sha(ip)}:${day}`;
        const used = +(await env.DOCS.get(rk) || 0);
        if (used >= MAX_PER_DAY) return fail('יותר מדי מסמכים נשלחו היום מהחיבור הזה. אפשר לנסות שוב מחר.', 429);
        const fd = await req.formData();
        const file = fd.get('pdf'); let meta;
        try { meta = JSON.parse(fd.get('meta') || '{}'); } catch { return fail('bad meta'); }
        if (!file || typeof file === 'string') return fail('no pdf');
        const pdf = await file.arrayBuffer();
        if (pdf.byteLength > MAX_PDF) return fail('הקובץ גדול מ-15MB');
        if (new TextDecoder().decode(new Uint8Array(pdf, 0, 5)) !== '%PDF-') return fail('not a pdf');
        const mode = meta.mode === 'view' ? 'view' : 'sign';
        const signers = (Array.isArray(meta.signers) ? meta.signers : []).slice(0, MAX_SIGNERS).map((s, i) => ({ n: i, name: clip(s.name, 60) || `חותם ${i + 1}`, token: rnd(20) }));
        if (mode === 'sign' && !signers.length) return fail('no signers');
        const fields = (Array.isArray(meta.fields) ? meta.fields : []).slice(0, MAX_FIELDS).map((f, i) => ({
          id: 'f' + i, signer: Math.max(0, Math.min(signers.length - 1, +f.signer || 0)), page: Math.max(0, +f.page || 0),
          type: ['sig', 'initials', 'name', 'date', 'text'].includes(f.type) ? f.type : 'sig',
          x: +(+f.x).toFixed(5), y: +(+f.y).toFixed(5), w: +(+f.w).toFixed(5), h: +(+f.h).toFixed(5), label: clip(f.label, 60), req: f.req !== false,
          r: Array.isArray(f.r) && f.r.length === 4 && f.r.every(v => isFinite(+v)) ? f.r.map(v => +(+v).toFixed(2)) : null
        })).filter(f => [f.x, f.y, f.w, f.h].every(v => isFinite(v) && v >= -0.01 && v <= 1.01));
        if (mode === 'sign') for (const s of signers) if (!fields.some(f => f.signer === s.n)) return fail(`אין מקום לחתימה של ${s.name}`);
        const id = rnd(12), owner = rnd(24), view = rnd(20);
        const rec = {
          v: 1, id, mode, title: clip(meta.title, 120) || 'מסמך', note: clip(meta.note, 500),
          from: { name: clip(meta.from && meta.from.name, 60), mail: clip(meta.from && meta.from.mail, 120) },
          order: meta.order === 'seq' ? 'seq' : 'any', pages: Math.max(1, +meta.pages || 1),
          created: new Date().toISOString(), expires: new Date(Date.now() + TTL * 1000).toISOString(),
          hash: await sha(pdf), size: pdf.byteLength, signers, fields, ownerHash: await sha(owner), view
        };
        await env.DOCS.put('f:' + id, pdf, { expirationTtl: TTL });
        await env.DOCS.put('m:' + id, JSON.stringify(rec), { expirationTtl: TTL });
        await env.DOCS.put(rk, String(used + 1), { expirationTtl: 2 * 86400 });
        return json({ id, owner, view, signers: signers.map(s => ({ n: s.n, name: s.name, token: s.token })), expires: rec.expires });
      }

      const id = parts[1];
      if (!/^[a-z0-9]{12}$/.test(id || '')) return fail('not found', 404);
      const rec = await env.DOCS.get('m:' + id, 'json');
      if (!rec) return fail('המסמך לא נמצא. ייתכן שעברו 30 יום והוא נמחק, או שהשולח מחק אותו.', 404);

      // who is asking: the sender (k), a signer (t) or someone with the viewing link (v)
      const k = url.searchParams.get('k'), t = url.searchParams.get('t'), v = url.searchParams.get('v');
      const isOwner = !!k && (await sha(k)) === rec.ownerHash;
      const me = t ? rec.signers.find(s => s.token === t) : null;
      const viewer = !!v && v === rec.view;
      if (!isOwner && !me && !viewer) return fail('הקישור לא תקין', 403);

      const signed = async () => {
        const out = {};
        await Promise.all(rec.signers.map(async s => { const x = await env.DOCS.get(`s:${id}:${s.n}`, 'json'); if (x) out[s.n] = x; }));
        return out;
      };
      const pub = async () => {
        const sg = await signed();
        return {
          id, mode: rec.mode, title: rec.title, note: rec.note, from: rec.from, order: rec.order, pages: rec.pages,
          created: rec.created, expires: rec.expires, hash: rec.hash, fields: rec.fields,
          signers: rec.signers.map(s => {
            const x = sg[s.n];
            return { n: s.n, name: s.name, status: x ? x.status : 'waiting', at: x ? x.at : null, values: x && x.status === 'signed' ? x.values : null,
              ...(x ? { ip: isOwner || me ? x.ip : undefined, ua: x.ua, country: x.country, tz: x.tz, reason: x.reason } : {}) };
          }),
          me: me ? me.n : null, owner: isOwner
        };
      };

      if (parts.length === 2 && req.method === 'GET') return json(await pub());

      if (parts[2] === 'pdf' && req.method === 'GET') {
        const pdf = await env.DOCS.get('f:' + id, 'arrayBuffer');
        if (!pdf) return fail('not found', 404);
        return new Response(pdf, { headers: { ...cors, 'Content-Type': 'application/pdf', 'Cache-Control': 'private, max-age=3600' } });
      }

      // a signer signs, or declines
      if ((parts[2] === 'sign' || parts[2] === 'decline') && req.method === 'POST') {
        if (!me || rec.mode !== 'sign') return fail('הקישור לא מאפשר חתימה', 403);
        const sg = await signed();
        if (sg[me.n]) return fail(sg[me.n].status === 'signed' ? 'כבר חתמת על המסמך הזה' : 'כבר סירבת לחתום על המסמך הזה', 409);
        if (rec.order === 'seq') { const before = rec.signers.filter(s => s.n < me.n && !(sg[s.n] && sg[s.n].status === 'signed')); if (before.length) return fail(`עוד לא תורך. קודם צריך לחתום: ${before.map(s => s.name).join(', ')}`, 409); }
        let d; try { d = await req.json(); } catch { return fail('bad request'); }
        const base = { at: new Date().toISOString(), ip, ua: clip(req.headers.get('User-Agent'), 200), country: (req.cf && req.cf.country) || '', tz: clip(d.tz, 60) };
        if (parts[2] === 'decline') {
          await env.DOCS.put(`s:${id}:${me.n}`, JSON.stringify({ ...base, status: 'declined', reason: clip(d.reason, 300) }), { expirationTtl: TTL });
          return json(await pub());
        }
        if (!d.agree) return fail('צריך לאשר את ההסכמה לחתימה אלקטרונית');
        const values = {};
        for (const f of rec.fields.filter(f => f.signer === me.n)) {
          const x = d.values && d.values[f.id];
          if (!x) { if (f.req) return fail('יש עוד שדות למלא'); continue; }
          if (f.type === 'sig' || f.type === 'initials') {
            const png = String(x.png || '');
            if (!/^data:image\/png;base64,[A-Za-z0-9+/=]+$/.test(png) || png.length > MAX_IMG * 1.4) return fail('חתימה לא תקינה');
            values[f.id] = { png };
          } else values[f.id] = { text: clip(x.text, 300), png: /^data:image\/png;base64,/.test(String(x.png || '')) && String(x.png).length < MAX_IMG * 1.4 ? x.png : undefined };
        }
        await env.DOCS.put(`s:${id}:${me.n}`, JSON.stringify({ ...base, status: 'signed', values }), { expirationTtl: TTL });
        return json(await pub());
      }

      // the sender deletes the document
      if (parts.length === 2 && req.method === 'DELETE') {
        if (!isOwner) return fail('forbidden', 403);
        await Promise.all([env.DOCS.delete('m:' + id), env.DOCS.delete('f:' + id), ...rec.signers.map(s => env.DOCS.delete(`s:${id}:${s.n}`))]);
        return json({ ok: true });
      }
      return fail('not found', 404);
    } catch (e) {
      return fail('server error: ' + (e && e.message || e), 500);
    }
  }
};
