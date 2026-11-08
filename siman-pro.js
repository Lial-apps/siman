/* =====================================================================
   סימן 6.1 — siman-pro.js
   1. השחרה אוטומטית: מוצא בקובץ מספרי זהות (עם ספרת ביקורת), טלפונים, מיילים, כרטיסי אשראי
      ומספרי IBAN, מראה רשימה לבדיקה, ומשחיר את מה שסומן. ההשחרה היא אותה השחרה של כלי "השחרה":
      הטקסט שמתחת נמחק מהקובץ באמת כששומרים. הכול במכשיר.
   2. דחיסת PDF: כל עמוד נכתב מחדש כתמונה קטנה יותר (כולל הסימונים והחתימות), לשליחה במייל
      ובטפסים מקוונים שמגבילים גודל. הכול במכשיר.
   נטען אחרי הסקריפט הראשי ומשתמש ב: S, drawAnns, snapshot, buildPdf, saveBytes, busy, toast, pdfjs, PDFLib
   ===================================================================== */
(() => {
'use strict';
if (window.SIMAN_PRO) return;
const $ = id => document.getElementById(id);
const esc = t => String(t ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);
const h = html => { const t = document.createElement('template'); t.innerHTML = html.trim(); return t.content.firstElementChild; };

document.head.append(h(`<style>
.pro-dlg .sig-in{max-width:560px}
.pro-list{max-height:min(52vh,420px);overflow:auto;border:1px solid var(--line);border-radius:12px;margin:10px 0;background:var(--bar)}
.pro-grp{padding:8px 12px;border-bottom:1px solid var(--line)}
.pro-grp:last-child{border-bottom:0}
.pro-grp>label{display:flex;gap:8px;align-items:center;font-weight:700;font-size:14px;margin-bottom:4px}
.pro-grp ul{list-style:none;margin:0;padding:0 26px 0 0;display:grid;gap:2px}
.pro-grp li label{display:flex;gap:8px;align-items:center;font-size:13.5px;cursor:pointer}
.pro-grp li code{direction:ltr;unicode-bidi:isolate;font-family:inherit;background:var(--desk);border-radius:6px;padding:0 6px}
.pro-grp li small{color:var(--muted)}
.pro-note{font-size:13px;color:var(--muted);margin:4px 0 0}
.pro-opts{display:grid;gap:8px;margin:12px 0}
.pro-opt{display:flex;gap:10px;align-items:flex-start;border:1px solid var(--line);border-radius:12px;padding:10px 12px;cursor:pointer;background:var(--bar)}
.pro-opt:has(input:checked){border-color:var(--ink);box-shadow:0 0 0 1px var(--ink)}
.pro-opt b{display:block}.pro-opt small{color:var(--muted)}
.pro-st-grid{display:grid;grid-template-columns:1fr 210px;gap:16px;align-items:start;margin-top:8px}
.pro-st-f{display:grid;gap:8px}
.pro-st-f .field input{width:100%}
.pro-st-seg{display:flex;flex-wrap:wrap;gap:6px}
.pro-st-seg button{font:inherit;font-size:13px;border:1px solid var(--line);background:var(--bar);color:var(--ink);border-radius:999px;padding:5px 11px;cursor:pointer;display:inline-flex;align-items:center;gap:6px}
.pro-st-seg button[aria-pressed=true]{border-color:var(--ink);box-shadow:0 0 0 1px var(--ink);font-weight:700}
.pro-st-seg i{width:11px;height:11px;border-radius:50%}
.pro-st-prev{background:#fff;border:1px dashed var(--line);border-radius:12px;padding:10px;display:grid;place-items:center;min-height:200px}
.pro-st-prev img{max-width:100%;max-height:220px}
.wa-choice{display:grid;gap:8px;margin-top:4px;text-align:start}
.wa-choice b{font-size:14px;line-height:1.4}
.wa-choice .btn{justify-content:flex-start;width:100%}
.wa-choice .wa-tip{color:var(--muted);font-size:12.5px;line-height:1.5}
.tp-gl{justify-self:start;margin:2px 0 6px}
.gl-rows{display:grid;gap:6px;margin:10px 0;max-height:48vh;overflow:auto}
.gl-row{display:grid;grid-template-columns:1fr auto 1fr auto;gap:6px;align-items:center}
.gl-row input{font:inherit;padding:7px 9px;border:1px solid var(--line);border-radius:8px;background:var(--bar);color:var(--ink);min-width:0}
.gl-row span{color:var(--muted)}
.pro-pill{font:inherit;font-size:12px;border:1px solid var(--line);background:var(--bar);color:var(--muted);border-radius:999px;padding:2px 10px;cursor:pointer;margin-inline-start:8px;white-space:nowrap;transition:background .2s,color .2s}
.pro-pill:hover{color:var(--ink);border-color:var(--ink)}
.pro-pill.out{background:color-mix(in srgb,var(--blue,#2E58CC) 14%,var(--bar));color:var(--blue,#2E58CC);border-color:var(--blue,#2E58CC);font-weight:700}
.pro-priv h3{font-size:15px;margin:14px 0 6px}
.pro-loc{columns:2;column-gap:20px;margin:0;padding-inline-start:18px;font-size:14px}
.pro-out{width:100%;border-collapse:collapse;font-size:13.5px}
.pro-out th,.pro-out td{border-bottom:1px solid var(--line);padding:6px 4px;text-align:start;vertical-align:top}
.pro-out th{color:var(--muted);font-weight:600}
.pro-link{border:0;background:none;color:var(--blue,#2E58CC);font:inherit;font-size:13.5px;text-decoration:underline;cursor:pointer;padding:0;justify-self:start}
@media (max-width:560px){.pro-loc{columns:1}}
body.pro-pick .page{cursor:crosshair;outline:3px dashed var(--marker);outline-offset:2px}
.pro-kiosk{width:100vw;height:100dvh;max-width:none;max-height:none;margin:0;border:0;padding:0;background:var(--bar,#fff);color:var(--ink)}
.pro-kiosk::backdrop{background:#000}
.pk-in{max-width:760px;margin:0 auto;padding:18px 18px calc(18px + env(safe-area-inset-bottom,0px));display:grid;gap:10px;min-height:100%;align-content:start}
.pk-in h2{margin:0;font-size:24px}
.pk-doc{margin:0;color:var(--muted)}
.pk-page{max-height:34dvh;overflow:auto;border:1px solid var(--line);border-radius:12px;background:var(--desk)}
.pk-pw{position:relative}.pk-pw img{display:block;width:100%}
.pk-pw i{position:absolute;transform:translate(-50%,-50%);background:var(--marker,#FFD84D);color:#1B1C1F;font-style:normal;font-weight:700;font-size:13px;padding:2px 9px;border-radius:999px;box-shadow:0 0 0 3px rgba(0,0,0,.25)}
.pk-in .field input{width:100%;font-size:18px;padding:10px 12px}
.pk-padw{position:relative;height:min(34dvh,260px);border:2px dashed var(--ink);border-radius:14px;background:#fff;touch-action:none}
.pk-pad{width:100%;height:100%;display:block;touch-action:none;cursor:crosshair}
.pk-ph{position:absolute;inset:auto 0 12px;text-align:center;color:#8A93A6;pointer-events:none}
.pk-ok{font-size:16px}
.pk-row{display:flex;gap:8px;align-items:center}.pk-row .grow{flex:1}
.pk-st{margin:0;min-height:1.2em;color:var(--muted);font-size:14px}
.pk-thanks{display:grid;gap:10px;justify-items:center;text-align:center;padding:18vh 10px}
.pk-thanks b{font-size:28px}
@media (max-width:560px){.pro-st-grid{grid-template-columns:1fr}.pro-st-prev{min-height:150px}.pro-st-prev img{max-height:160px}}
</style>`));

/* ---------------- 0. the Hebrew date (6.2) ----------------
   "ט״ז בחשוון תשפ״ז" — the calendar comes from the browser itself (Intl, Hebrew calendar), the letters are counted here */
function gem(n) {
  const U = ['', 'א', 'ב', 'ג', 'ד', 'ה', 'ו', 'ז', 'ח', 'ט'], T = ['', 'י', 'כ', 'ל', 'מ', 'נ', 'ס', 'ע', 'פ', 'צ'], H = ['', 'ק', 'ר', 'ש', 'ת', 'תק', 'תר', 'תש', 'תת', 'תתק'];
  let s = H[Math.floor(n / 100)]; n %= 100;
  s += n === 15 ? 'טו' : n === 16 ? 'טז' : T[Math.floor(n / 10)] + U[n % 10];
  return s.length > 1 ? s.slice(0, -1) + '״' + s.slice(-1) : s + '׳';
}
function hebDate(d = new Date()) {
  try {
    const p = new Intl.DateTimeFormat('en-u-ca-hebrew', { day: 'numeric', month: 'numeric', year: 'numeric' }).formatToParts(d);
    const day = +p.find(x => x.type === 'day').value, year = +p.find(x => x.type === 'year').value;
    const month = new Intl.DateTimeFormat('he-u-ca-hebrew', { month: 'long' }).format(d).replace(/^ב/, '');
    if (!day || !year || !month || /\d/.test(month)) return '';
    return `${gem(day)} ב${month} ${gem(year % 1000)}`;
  } catch (_) { return ''; }
}
window.SIMAN_HEBDATE = hebDate;

/* ---------------- 0b. a business stamp (6.2) ----------------
   round or rectangular, like a rubber stamp: the business name, its number (ח.פ / ע.מ / ע.ר) and one more line.
   It is drawn here and placed on the page like a signature. Kept on this device for next time. */
const STAMP_KEY = 'siman.stamp';
const INK = { blue: '#1D3FB0', red: '#B3232A', black: '#1B1C1F', purple: '#5B2BB8' };
function drawStamp(o) {
  const ink = INK[o.color] || INK.blue, font = w => `${w} %spx "SF-rubik","Rubik","Segoe UI",Arial,sans-serif`;
  const c = document.createElement('canvas'), g = c.getContext('2d');
  const fit = (t, max, size, w) => { let s = size; do { g.font = font(w).replace('%s', s); s -= 2; } while (g.measureText(t).width > max && s > 14); return g.font; };
  if (o.shape === 'round') {
    const W = 640, R = 300, C = W / 2; c.width = c.height = W;
    g.strokeStyle = g.fillStyle = ink; g.lineWidth = 12; g.beginPath(); g.arc(C, C, R, 0, Math.PI * 2); g.stroke();
    g.lineWidth = 5; g.beginPath(); g.arc(C, C, R - 92, 0, Math.PI * 2); g.stroke();
    // text along an arc, read right to left: the top line runs over the top, the bottom line under it
    const arc = (t, r, top) => {
      if (!t) return;
      let size = 64; g.font = font(700).replace('%s', size);
      const room = Math.PI * (top ? 0.95 : 0.7) * r;
      while (g.measureText(t).width + t.length * 4 > room && size > 26) { size -= 2; g.font = font(700).replace('%s', size); }
      // letters are laid one by one from right to left: numbers and English inside are written left to right,
      // so each such run is turned around first ("03-1234567" stays readable)
      const vis = t.replace(/[A-Za-z0-9](?:[A-Za-z0-9.\-\/:+]*[A-Za-z0-9])?/g, m => [...m].reverse().join(''));
      const chars = [...vis], ws = chars.map(ch => g.measureText(ch).width + 4), span = ws.reduce((a, b) => a + b, 0) / r;
      let a = top ? -Math.PI / 2 + span / 2 : Math.PI / 2 - span / 2;
      g.textAlign = 'center'; g.textBaseline = 'middle';
      chars.forEach((ch, k) => {
        const step = ws[k] / r, mid = top ? a - step / 2 : a + step / 2;
        g.save(); g.translate(C + Math.cos(mid) * r, C + Math.sin(mid) * r); g.rotate(top ? mid + Math.PI / 2 : mid - Math.PI / 2); g.fillText(ch, 0, 0); g.restore();
        a = top ? a - step : a + step;
      });
    };
    arc(o.name, R - 48, true); arc(o.line2, R - 48, false);
    g.textAlign = 'center'; g.textBaseline = 'middle';
    if (o.num) { g.font = fit(o.num, 2 * (R - 110), 58, 700); g.fillText(o.num, C, C + (o.center ? -26 : 0)); }
    if (o.center) { g.font = fit(o.center, 2 * (R - 120), 40, 400); g.fillText(o.center, C, C + (o.num ? 40 : 0)); }
  } else {
    const W = 900, lines = [o.name, o.num, o.line2, o.center].filter(Boolean), H = 70 + lines.length * 74;
    c.width = W; c.height = H; g.strokeStyle = g.fillStyle = ink;
    g.lineWidth = 10; g.strokeRect(8, 8, W - 16, H - 16); g.lineWidth = 4; g.strokeRect(26, 26, W - 52, H - 52);
    g.textAlign = 'center'; g.textBaseline = 'middle'; g.direction = 'rtl';
    lines.forEach((t, k) => { g.font = fit(t, W - 110, k ? 46 : 60, k ? 400 : 700); g.fillText(t, W / 2, 35 + 74 * k + 37); });
  }
  // a rubber stamp is never perfectly even: a light, fixed grain
  const d = g.getImageData(0, 0, c.width, c.height), px = d.data; let seed = 7;
  for (let i = 3; i < px.length; i += 4) if (px[i]) { seed = (seed * 16807) % 2147483647; if (seed % 100 < 9) px[i] = px[i] * 0.55; }
  g.putImageData(d, 0, 0);
  return c;
}
function stampDialog() {
  if (!S.doc) { toast('פתחו קובץ קודם, ואז מוסיפים את החותמת'); return; }
  let o; try { o = JSON.parse(localStorage.getItem(STAMP_KEY) || 'null'); } catch (_) {}
  o = Object.assign({ shape: 'round', color: 'blue', name: '', num: '', line2: '', center: '' }, o || {});
  const dlg = h(`<dialog class="sig pro-dlg pro-stamp" aria-labelledby="proStT"><div class="sig-in">
    <h2 id="proStT">חותמת עסק</h2>
    <div class="pro-st-grid"><div class="pro-st-f">
      <label class="field">שם העסק או העמותה<input data-k="name" maxlength="40" placeholder="למשל: כהן ובניו בע״מ"></label>
      <label class="field">מספר (ח.פ / ע.מ / ע.ר)<input data-k="num" maxlength="30" placeholder="ח.פ 512345678"></label>
      <label class="field">שורה נוספת (לא חובה)<input data-k="line2" maxlength="40" placeholder="למשל: תל אביב · 03-1234567"></label>
      <label class="field">באמצע (לא חובה)<input data-k="center" maxlength="24" placeholder="למשל: נאמן למקור"></label>
      <div class="pro-st-seg" data-g="shape"><button type="button" data-v="round">עגולה</button><button type="button" data-v="rect">מלבנית</button></div>
      <div class="pro-st-seg" data-g="color">${Object.entries({ blue: 'כחול', red: 'אדום', purple: 'סגול', black: 'שחור' }).map(([k, v]) => `<button type="button" data-v="${k}"><i style="background:${INK[k]}"></i>${v}</button>`).join('')}</div>
    </div><div class="pro-st-prev"><img alt="תצוגה מקדימה של החותמת"></div></div>
    <p class="pro-note">החותמת נשמרת במכשיר הזה בלבד, לפעם הבאה. אחרי "הצבה" לוחצים על הדף במקום שבו לשים אותה.</p>
    <div class="sig-row"><span class="grow"></span><button class="btn" data-a="no">ביטול</button><button class="btn primary" data-a="go">הצבה על הדף</button></div></div></dialog>`);
  document.body.append(dlg);
  const img = dlg.querySelector('img');
  const redraw = () => {
    dlg.querySelectorAll('.pro-st-seg').forEach(s => s.querySelectorAll('button').forEach(b => b.setAttribute('aria-pressed', String(o[s.dataset.g] === b.dataset.v))));
    const c = drawStamp(o.name || o.num ? o : { ...o, name: 'שם העסק', num: 'ח.פ 000000000' }); img.src = c.toDataURL('image/png');
  };
  dlg.querySelectorAll('input[data-k]').forEach(i => { i.value = o[i.dataset.k] || ''; i.oninput = () => { o[i.dataset.k] = i.value.trim(); redraw(); }; });
  dlg.querySelectorAll('.pro-st-seg button').forEach(b => b.onclick = () => { o[b.closest('.pro-st-seg').dataset.g] = b.dataset.v; redraw(); });
  dlg.addEventListener('close', () => dlg.remove());
  dlg.querySelector('[data-a=no]').onclick = () => dlg.close();
  dlg.querySelector('[data-a=go]').onclick = () => {
    if (!o.name && !o.num) { dlg.querySelector('input[data-k=name]').focus(); return; }
    try { localStorage.setItem(STAMP_KEY, JSON.stringify(o)); } catch (_) {}
    const c = drawStamp(o); dlg.close();
    startPlace(c.toDataURL('image/png'), c.height / c.width, o.shape === 'round' ? 105 : 170, false);
    toast('לוחצים על הדף במקום שבו לשים את החותמת');
  };
  redraw(); dlg.showModal();
}

/* ---------------- 0c. signing face to face (6.2) ----------------
   The person who holds the document taps where the client signs, and hands over the phone or tablet.
   The client sees the page, writes their name, ticks "I read it" and signs on a big pad.
   Name, date and time are written under the signature. Nothing is sent anywhere. */
function inPerson() {
  if (!S.doc) { toast('פתחו קובץ קודם'); return; }
  if (typeof setTool === 'function') setTool('select');
  toast('לחצו על הדף במקום שבו החותם יחתום (Esc לביטול)');
  document.body.classList.add('pro-pick');
  const done = () => { removeEventListener('pointerdown', down, true); removeEventListener('keydown', key, true); document.body.classList.remove('pro-pick'); };
  const down = e => {
    const pg = e.target.closest && e.target.closest('.page'); if (!pg) return;
    e.preventDefault(); e.stopPropagation(); done();
    const P = S.pages[+pg.dataset.i]; const [cx, cy] = localXY(P, e);
    kiosk(P, cx, cy);
  };
  const key = e => { if (e.key === 'Escape') { done(); toast('בוטל'); } };
  addEventListener('pointerdown', down, true); addEventListener('keydown', key, true);
}
function kiosk(P, cx, cy) {
  let pageImg = ''; try { pageImg = P.canvas.toDataURL('image/jpeg', 0.8); } catch (_) {}
  const W = P.el.clientWidth || 1, Hh = P.el.clientHeight || 1;
  const dlg = h(`<dialog class="pro-kiosk" aria-labelledby="proKT"><div class="pk-in">
    <h2 id="proKT">חתימה על המסמך</h2><p class="pk-doc"></p>
    <div class="pk-page">${pageImg ? `<div class="pk-pw"><img alt="העמוד שעליו חותמים"><i style="left:${cx / W * 100}%;top:${cy / Hh * 100}%">כאן</i></div>` : ''}</div>
    <label class="field">השם המלא<input class="pk-name" autocomplete="off" maxlength="50"></label>
    <div class="pk-padw"><canvas class="pk-pad" aria-label="משטח חתימה"></canvas><span class="pk-ph">חותמים כאן באצבע או בעט</span></div>
    <label class="chk pk-ok"><input type="checkbox"> קראתי את המסמך ואני מאשר/ת את החתימה</label>
    <div class="pk-row"><button class="btn" data-a="clr" type="button">ניקוי</button><span class="grow"></span><button class="btn" data-a="no" type="button">ביטול</button><button class="btn primary big" data-a="go" type="button" disabled>אני חותם/ת</button></div>
    <p class="pk-st" role="status"></p>
  </div></dialog>`);
  document.body.append(dlg);
  dlg.querySelector('.pk-doc').textContent = (S.name || 'מסמך') + ' · עמוד ' + P.i;
  if (pageImg) dlg.querySelector('.pk-page img').src = pageImg;
  const cv = dlg.querySelector('.pk-pad'), g = cv.getContext('2d'), ok = dlg.querySelector('.pk-ok input'), go = dlg.querySelector('[data-a=go]'), nm = dlg.querySelector('.pk-name');
  let ink = false, drawing = false, last = null;
  const size = () => { const r = cv.getBoundingClientRect(), d = devicePixelRatio || 1; cv.width = r.width * d; cv.height = r.height * d; g.lineCap = g.lineJoin = 'round'; g.lineWidth = 3.2 * d; g.strokeStyle = '#14215C'; ink = false; dlg.querySelector('.pk-ph').hidden = false; upd(); };
  const at = e => { const r = cv.getBoundingClientRect(), d = devicePixelRatio || 1; return [(e.clientX - r.left) * d, (e.clientY - r.top) * d]; };
  const upd = () => { go.disabled = !(ink && ok.checked && nm.value.trim()); dlg.querySelector('.pk-st').textContent = !nm.value.trim() ? 'כתבו את השם המלא' : !ink ? 'חתמו במשטח' : !ok.checked ? 'סמנו שקראתם את המסמך' : ''; };
  cv.onpointerdown = e => { drawing = true; last = at(e); try { cv.setPointerCapture(e.pointerId); } catch (_) {} g.beginPath(); g.arc(last[0], last[1], g.lineWidth / 2, 0, 7); g.fillStyle = g.strokeStyle; g.fill(); };
  cv.onpointermove = e => { if (!drawing) return; for (const ev of e.getCoalescedEvents ? e.getCoalescedEvents() : [e]) { const p = at(ev); g.beginPath(); g.moveTo(last[0], last[1]); g.lineTo(p[0], p[1]); g.stroke(); last = p; } if (!ink) { ink = true; dlg.querySelector('.pk-ph').hidden = true; upd(); } };
  cv.onpointerup = cv.onpointercancel = () => { drawing = false; };
  ok.onchange = upd; nm.oninput = upd;
  dlg.querySelector('[data-a=clr]').onclick = () => { g.clearRect(0, 0, cv.width, cv.height); size(); };
  dlg.querySelector('[data-a=no]').onclick = () => dlg.close();
  dlg.addEventListener('cancel', e => e.preventDefault());   // the client cannot close it by mistake with Esc
  dlg.addEventListener('close', () => dlg.remove());
  go.onclick = () => {
    // the signature, trimmed to its ink, with the name and the date and time under it
    const d = g.getImageData(0, 0, cv.width, cv.height).data; let x0 = cv.width, y0 = cv.height, x1 = 0, y1 = 0;
    for (let y = 0; y < cv.height; y += 2) for (let x = 0; x < cv.width; x += 2) if (d[(y * cv.width + x) * 4 + 3] > 20) { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; }
    const pad = 10, sw = Math.max(1, x1 - x0 + pad * 2), sh = Math.max(1, y1 - y0 + pad * 2);
    const now = new Date(), line = `${nm.value.trim()} · ${now.toLocaleDateString('he-IL')} ${now.toLocaleTimeString('he-IL', { hour: '2-digit', minute: '2-digit' })}`;
    const out = document.createElement('canvas'), o = out.getContext('2d');
    const fs = Math.max(22, Math.round(sw * 0.07)); o.font = `${fs}px "SF-rubik","Rubik",Arial,sans-serif`;
    out.width = Math.max(sw, o.measureText(line).width + 20); out.height = sh + fs * 1.6;
    o.drawImage(cv, x0 - pad, y0 - pad, sw, sh, (out.width - sw) / 2, 0, sw, sh);
    o.font = `${fs}px "SF-rubik","Rubik",Arial,sans-serif`; o.fillStyle = '#14215C'; o.textAlign = 'center'; o.direction = 'rtl'; o.fillText(line, out.width / 2, sh + fs * 1.15);
    addImage(P, out.toDataURL('image/png'), out.height / out.width, 150, cx, cy, true);
    dlg.querySelector('.pk-in').innerHTML = `<div class="pk-thanks"><b>תודה, ${esc(nm.value.trim())}!</b><p>החתימה נכנסה למסמך. אפשר להחזיר את המכשיר.</p><button class="btn primary big" type="button">סיום</button></div>`;
    dlg.querySelector('.pk-thanks button').onclick = () => { dlg.close(); toast('החתימה נוספה. לא לשכוח לשמור'); };
  };
  dlg.showModal(); size(); setTimeout(() => nm.focus(), 60);
}

/* ---------------- 0d. what leaves the device (6.2) ----------------
   A small pill says "🔒 במכשיר". While something is really being sent to a server (the assistant, translation,
   a transcript summary, sending a document) it turns into "☁ נשלח ל-…" — watched on the requests themselves,
   so it can never claim "on the device" while something is on its way out. A tap opens the full list. */
const CLOUD = [
  [/siman-ai\.|siman-voice\./, 'Google Gemini'],
  [/siman-translate\./, 'שירות התרגום'],
  [/siman-share\./, 'שרת הקישורים של סימן'],
];
const pill = h('<button type="button" class="pro-pill" title="מה יוצא מהמכשיר? לחצו לפרטים"><span>🔒 במכשיר</span></button>');
let live = 0;
const setPill = to => { pill.classList.toggle('out', !!to); pill.firstChild.textContent = to ? '☁ נשלח ל-' + to + '…' : '🔒 במכשיר'; };
{ const pi = document.getElementById('pageInfo'); (pi ? pi.parentNode : document.body).insertBefore(pill, pi ? pi.nextSibling : null); }
pill.onclick = () => privacyDialog();
{ const f0 = window.fetch;
  window.fetch = function (input, init) {
    const url = typeof input === 'string' ? input : (input && input.url) || '';
    // (downloading the transcription engine through the voice server sends nothing about the user: not counted)
    const hit = CLOUD.find(([re]) => re.test(url)) && !/\/(register|status|delete|hf)\b/.test(url);
    if (!hit) return f0.apply(this, arguments);
    const to = CLOUD.find(([re]) => re.test(url))[1];
    live++; setPill(to);
    const end = () => { if (--live <= 0) { live = 0; setPill(''); } };
    return f0.apply(this, arguments).then(r => { end(); return r; }, e => { end(); throw e; });
  }; }
const LOCAL = ['פתיחה, סימון, עט ומרקר', 'עריכת טקסט קיים', 'מילוי טפסים וזיהוי השדות', 'חתימה, חותמת עסק, חתימה פנים מול פנים', 'השחרה (גם אוטומטית)', 'דחיסת PDF', 'המרה ל-Word, Excel, PowerPoint, תמונות, Markdown וטקסט', 'איחוד ופיצול', 'זיהוי טקסט (OCR)', 'תמלול השמע', 'הקראה בקול', 'עורכי Word, Excel ומצגות'];
const OUT = [['העוזר החכם: שאלות, סיכום, כרטיסיות, בוחן, מפת חשיבה, עריכה בעזרת AI', 'הטקסט של המסמך והשאלה', 'Google Gemini'],
  ['תרגום', 'הטקסט של המסמך', 'Google Gemini או Anthropic Claude'],
  ['מדף למסמך (כתב יד) ו-Word חכם', 'תמונת העמוד', 'Google Gemini'],
  ['סיכום, שאלות, פרקים ותרגום של תמליל', 'הטקסט של התמליל, בלי השמע', 'Google Gemini'],
  ['שליחה לחתימה וקישור לצפייה', 'המסמך, לכן נשמר 30 יום', 'שרת הקישורים של סימן (Cloudflare)']];
function privacyDialog() {
  const dlg = h(`<dialog class="sig pro-dlg pro-priv" aria-labelledby="proPrT"><div class="sig-in">
    <h2 id="proPrT">מה יוצא מהמכשיר?</h2>
    <h3>🔒 נשאר במכשיר, תמיד</h3><ul class="pro-loc">${LOCAL.map(x => `<li>${x}</li>`).join('')}</ul>
    <h3>☁ נשלח רק כשלוחצים על הפעולה</h3>
    <table class="pro-out"><thead><tr><th>פעולה</th><th>מה נשלח</th><th>לאן</th></tr></thead><tbody>${OUT.map(r => `<tr>${r.map(c => `<td>${c}</td>`).join('')}</tr>`).join('')}</tbody></table>
    <p class="pro-note">כשמשהו נשלח ממש עכשיו, התג "🔒 במכשיר" מתחלף ל"☁ נשלח ל-…" עד שהתשובה חוזרת. שום דבר לא נשמר אצלנו אחרי התשובה, חוץ ממסמך ששולחים בקישור. <a href="privacy.html" target="_blank" rel="noopener">מדיניות הפרטיות המלאה</a></p>
    <div class="sig-row"><span class="grow"></span><button class="btn primary" data-a="ok">הבנתי</button></div></div></dialog>`);
  document.body.append(dlg); dlg.addEventListener('close', () => dlg.remove());
  dlg.querySelector('[data-a=ok]').onclick = () => dlg.close(); dlg.showModal();
}
{ const note = document.querySelector('#setDlg .fmt-note'); if (note) { const b = h('<button type="button" class="pro-link">🔒 מה יוצא מהמכשיר?</button>'); b.onclick = privacyDialog; note.after(b); } }

/* ---------------- 1. sensitive information ---------------- */
const KINDS = { id: 'מספרי זהות', phone: 'טלפונים', mail: 'כתובות מייל', card: 'כרטיסי אשראי', iban: 'חשבונות בנק (IBAN)' };
const digits = s => s.replace(/\D/g, '');
// Israeli ID: 9 digits (8 are padded with a 0) and a check digit
function validId(d) {
  if (d.length < 8 || d.length > 9) return false;
  d = d.padStart(9, '0'); if (/^0+$/.test(d) || /^(\d)\1+$/.test(d)) return false;
  let s = 0; for (let i = 0; i < 9; i++) { let x = +d[i] * ((i % 2) + 1); if (x > 9) x -= 9; s += x; }
  return s % 10 === 0;
}
function luhn(d) {
  let s = 0, alt = false;
  for (let i = d.length - 1; i >= 0; i--) { let x = +d[i]; if (alt) { x *= 2; if (x > 9) x -= 9; } s += x; alt = !alt; }
  return s % 10 === 0;
}
// the order matters: a longer number (a card) is taken before the shorter ones inside it
const RULES = [
  ['mail', /[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g],
  ['iban', /\bIL\d{2}(?:[ -]?\d{4}){4}[ -]?\d{3}\b/gi],
  ['card', /(?<![\d-])(?:\d[ -]?){12,18}\d(?![\d-])/g, m => { const d = digits(m); return d.length >= 13 && d.length <= 19 && luhn(d) && !/^(\d)\1+$/.test(d); }],
  ['phone', /(?<![\d-])(?:\+?972[- ]?|0)(?:5\d|7\d|[2-489])[- ]?\d{3}[- ]?\d{4}(?![\d-])/g],
  // a PDF that keeps Hebrew backwards may keep the digits backwards too: both ways are checked
  ['id', /(?<![\d-])\d{8,9}(?![\d-])|(?<![\d-])\d{3}-?\d{5}-?\d(?![\d-])/g, m => { const d = digits(m); return validId(d) || validId([...d].reverse().join('')); }],
];
const RTL = /[֐-ࣿ]/;
// where letters j0..j1 sit inside a piece of text, as a share of its width: measured, since letters differ in width
const MC = document.createElement('canvas').getContext('2d');
function span(it, j0, j1) {
  MC.font = `100px ${it.fam}`;
  const W = MC.measureText(it.s).width;
  if (!(W > 0)) return [j0 / it.s.length, (j1 + 1) / it.s.length];
  return [MC.measureText(it.s.slice(0, j0)).width / W, Math.min(1, MC.measureText(it.s.slice(0, j1 + 1)).width / W)];
}
// the words of one page, joined into lines, with each letter knowing which word piece it came from
async function pageHits(P) {
  const tc = await P.page.getTextContent();
  const items = tc.items.filter(it => it.str && it.str.trim() && Math.abs(it.transform[1]) < 0.01 && Math.abs(it.transform[2]) < 0.01)
    .map(it => ({ s: it.str, x: it.transform[4], y: it.transform[5], w: it.width, h: Math.abs(it.transform[3]) || it.height || 10, fam: (tc.styles && tc.styles[it.fontName] && tc.styles[it.fontName].fontFamily) || 'sans-serif' }));
  const lines = [];
  for (const it of items.sort((a, b) => b.y - a.y || a.x - b.x)) {
    const L = lines.find(l => Math.abs(l.y - it.y) < Math.max(l.h, it.h) * 0.45);
    if (L) L.items.push(it); else lines.push({ y: it.y, h: it.h, items: [it] });
  }
  const hits = [];
  for (const L of lines) {
    // left to right on the page: numbers, mails and phones are written left to right also inside Hebrew
    L.items.sort((a, b) => a.x - b.x);
    let str = ''; const map = [];
    L.items.forEach((it, k) => {
      if (k) { const p = L.items[k - 1]; const gap = it.x - (p.x + p.w); if (gap > it.h * 0.25) { str += ' '; map.push(null); } }
      for (let j = 0; j < it.s.length; j++) { str += it.s[j]; map.push([k, j]); }
    });
    const used = [];
    for (const [kind, re, ok] of RULES) {
      re.lastIndex = 0; let m;
      while ((m = re.exec(str))) {
        const a = m.index, b = a + m[0].length;
        if (used.some(([x, y]) => a < y && b > x)) continue;
        if (ok && !ok(m[0])) continue;
        used.push([a, b]);
        // the piece of each word item that the match covers
        const per = new Map();
        for (let c = a; c < b; c++) { const mp = map[c]; if (!mp) continue; const r = per.get(mp[0]) || [mp[1], mp[1]]; r[0] = Math.min(r[0], mp[1]); r[1] = Math.max(r[1], mp[1]); per.set(mp[0], r); }
        const rects = [];
        for (const [k, [j0, j1]] of per) {
          const it = L.items[k];
          let [f0, f1] = span(it, j0, j1);
          if (RTL.test(it.s)) { const t = f0; f0 = 1 - f1; f1 = 1 - t; }   // a Hebrew line runs right to left
          // a little extra on each side: better to cover a hair too much than to leave a digit showing
          const pad = it.h * 0.18;
          rects.push({ x1: it.x + it.w * f0 - pad, x2: it.x + it.w * f1 + pad, y1: it.y - it.h * 0.28, y2: it.y + it.h * 0.95 });
        }
        // a number found backwards is shown the right way round in the list
        const shown = kind === 'id' && !validId(digits(m[0])) ? [...m[0].trim()].reverse().join('') : m[0].trim();
        if (rects.length) hits.push({ kind, text: shown, page: P.i, rects });
      }
    }
  }
  return hits;
}
const mask = (kind, t) => kind === 'mail' ? t.replace(/^(.).*?(@)/, '$1•••$2') : t.replace(/\d(?=[\d\s-]{3})/g, '•');
async function autoRedact() {
  if (!S.doc) { toast('פתחו קובץ PDF קודם'); return; }
  if (S.office) { toast('השחרה אוטומטית עובדת על PDF. שמרו את הקובץ כ-PDF ופתחו אותו שוב'); return; }
  busy(true, 'מחפש מידע רגיש…');
  let hits = [];
  try {
    for (const P of S.pages) { $('busyMsg').textContent = `מחפש מידע רגיש… עמוד ${P.i} מתוך ${S.pages.length}`; hits.push(...await pageHits(P)); }
  } catch (e) { console.error(e); busy(false); toast('החיפוש נכשל: ' + (e.message || e)); return; }
  busy(false);
  // what is already blacked out is not offered again
  const covered = r => S.anns.some(a => a.type === 'redact' && Math.min(a.x1, a.x2) <= r.x1 + 2 && Math.max(a.x1, a.x2) >= r.x2 - 2 && Math.min(a.y1, a.y2) <= r.y1 + 2 && Math.max(a.y1, a.y2) >= r.y2 - 2);
  hits = hits.filter((x, i) => !x.rects.every(r => covered({ x1: r.x1, x2: r.x2, y1: r.y1, y2: r.y2 })));
  if (!hits.length) {
    const any = S.pages.some(P => P.ocr) || (await Promise.all(S.pages.slice(0, 3).map(P => P.page.getTextContent().then(t => t.items.length)))).some(Boolean);
    toast(any ? 'לא נמצאו מספרי זהות, טלפונים, מיילים או כרטיסי אשראי בקובץ' : 'לא נמצא טקסט בקובץ. אם הוא סרוק: "עמודים" ← "זיהוי טקסט (OCR)" קודם');
    return;
  }
  const by = {}; hits.forEach((x, i) => { (by[x.kind] = by[x.kind] || []).push(i); });
  const dlg = h(`<dialog class="sig pro-dlg" aria-labelledby="proRedT"><div class="sig-in">
    <h2 id="proRedT">נמצא מידע רגיש (${hits.length})</h2>
    <p class="pro-note">מה שמסומן יושחר. הטקסט שמתחת יימחק מהקובץ באמת כששומרים. אפשר לבטל אחר כך עם Ctrl+Z.</p>
    <div class="pro-list">${Object.keys(KINDS).filter(k => by[k]).map(k => `<div class="pro-grp"><label><input type="checkbox" data-g="${k}" checked> ${KINDS[k]} (${by[k].length})</label><ul>${by[k].map(i => `<li><label><input type="checkbox" data-i="${i}" data-k="${k}" checked><code>${esc(mask(k, hits[i].text))}</code><small>עמוד ${hits[i].page}</small></label></li>`).join('')}</ul></div>`).join('')}</div>
    <div class="sig-row"><span class="grow"></span><button class="btn" data-a="no">ביטול</button><button class="btn primary" data-a="go">להשחיר את המסומנים</button></div></div></dialog>`);
  document.body.append(dlg);
  dlg.querySelectorAll('[data-g]').forEach(g => g.onchange = () => dlg.querySelectorAll(`[data-k="${g.dataset.g}"]`).forEach(c => { c.checked = g.checked; }));
  dlg.querySelectorAll('[data-i]').forEach(c => c.onchange = () => { const all = [...dlg.querySelectorAll(`[data-k="${c.dataset.k}"]`)]; const g = dlg.querySelector(`[data-g="${c.dataset.k}"]`); g.checked = all.every(x => x.checked); g.indeterminate = !g.checked && all.some(x => x.checked); });
  dlg.addEventListener('close', () => dlg.remove());
  dlg.querySelector('[data-a=no]').onclick = () => dlg.close();
  dlg.querySelector('[data-a=go]').onclick = () => {
    const pick = [...dlg.querySelectorAll('[data-i]:checked')].map(c => hits[+c.dataset.i]);
    dlg.close();
    if (!pick.length) return;
    snapshot();
    const pages = new Set();
    for (const x of pick) for (const r of x.rects) { S.anns.push({ id: S.nextId++, type: 'redact', page: x.page - 1, x1: r.x1, y1: r.y1, x2: r.x2, y2: r.y2 }); pages.add(x.page); }
    for (const p of pages) drawAnns(S.pages[p - 1]);
    const first = S.pages[Math.min(...pages) - 1]; if (first) first.el.scrollIntoView({ block: 'start', behavior: 'smooth' });
    toast(`הושחרו ${pick.length} פריטים. בדקו את הדפים ושמרו כדי למחוק אותם מהקובץ`);
  };
  dlg.showModal();
}

/* ---------------- 2. a smaller PDF ---------------- */
const LEVELS = {
  mail: { l: 'מאוזן', d: 'מתאים למייל ולטפסים באתרים. הטקסט נשאר קריא וחד.', dpi: 150, q: 0.72 },
  small: { l: 'הכי קטן', d: 'לקבצים גדולים במיוחד. מתאים לקריאה על המסך.', dpi: 105, q: 0.55 },
};
const kb = n => n > 1048576 ? (n / 1048576).toFixed(1).replace('.0', '') + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB';
function compressDialog() {
  if (!S.doc) { toast('פתחו קובץ PDF קודם'); return; }
  const size = S.bytes ? S.bytes.length || S.bytes.byteLength : 0;
  const dlg = h(`<dialog class="sig pro-dlg" aria-labelledby="proCmpT"><div class="sig-in">
    <h2 id="proCmpT">דחיסה: PDF קטן יותר</h2>
    <p class="pro-note">${size ? `הגודל עכשיו: <b>${kb(size)}</b>. ` : ''}כל עמוד נשמר מחדש כתמונה קטנה יותר, עם הסימונים והחתימות. נוצר עותק חדש, והקובץ המקורי לא משתנה.</p>
    <div class="pro-opts">${Object.entries(LEVELS).map(([k, v], i) => `<label class="pro-opt"><input type="radio" name="proLvl" value="${k}"${i ? '' : ' checked'}><span><b>${v.l}</b><small>${v.d}</small></span></label>`).join('')}</div>
    <p class="pro-note">שימו לב: אחרי הדחיסה אי אפשר לסמן ולהעתיק טקסט מהקובץ (אפשר להחזיר עם "זיהוי טקסט (OCR)").</p>
    <div class="sig-row"><span class="grow"></span><button class="btn" data-a="no">ביטול</button><button class="btn primary" data-a="go">דחיסה ושמירה</button></div></div></dialog>`);
  document.body.append(dlg);
  dlg.addEventListener('close', () => dlg.remove());
  dlg.querySelector('[data-a=no]').onclick = () => dlg.close();
  dlg.querySelector('[data-a=go]').onclick = () => { const k = dlg.querySelector('input[name=proLvl]:checked').value; dlg.close(); compress(LEVELS[k]); };
  dlg.showModal();
}
async function compress(lv) {
  busy(true, 'מכין את הקובץ…');
  let doc = null;
  try {
    const built = await buildPdf();
    const before = built.length || built.byteLength;
    doc = await pdfjs.getDocument({ data: built.slice ? built.slice() : built, isEvalSupported: false }).promise;
    const out = await PDFLib.PDFDocument.create();
    for (let i = 1; i <= doc.numPages; i++) {
      $('busyMsg').textContent = `דוחס עמוד ${i} מתוך ${doc.numPages}…`;
      const page = await doc.getPage(i), v1 = page.getViewport({ scale: 1 });
      // no side over 3000 pixels: phones run out of memory beyond that
      const sc = Math.min(lv.dpi / 72, 3000 / Math.max(v1.width, v1.height)), vp = page.getViewport({ scale: sc });
      const c = document.createElement('canvas'); c.width = Math.ceil(vp.width); c.height = Math.ceil(vp.height);
      const g = c.getContext('2d'); g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height);
      await page.render({ canvasContext: g, viewport: vp }).promise;
      const blob = await new Promise(r => c.toBlob(r, 'image/jpeg', lv.q)); c.width = c.height = 0;
      const img = await out.embedJpg(new Uint8Array(await blob.arrayBuffer()));
      const pg = out.addPage([v1.width, v1.height]); pg.drawImage(img, { x: 0, y: 0, width: v1.width, height: v1.height });
      page.cleanup();
    }
    out.setTitle(S.name.replace(/\.pdf$/i, '')); out.setProducer('סימן'); out.setCreator('סימן');
    const bytes = await out.save({ useObjectStreams: true });
    busy(false);
    if (bytes.length >= before * 0.95) { toast(`הקובץ כבר קטן (${kb(before)}). דחיסה לא תקטין אותו עוד`); return; }
    const name = S.name.replace(/\.pdf$/i, '') + ' - מכווץ.pdf';
    const saved = await saveBytes(name, bytes, 'pdf');
    // the size the user knows is the file as it was opened (when nothing was added since)
    const orig = (S.bytes && (S.bytes.length || S.bytes.byteLength)) || 0, from = S.anns.length || !orig ? before : orig;
    if (saved) toast(`נשמר: ${name} · ${kb(from)} ← ${kb(bytes.length)} (קטן ב-${Math.round((1 - bytes.length / from) * 100)}%)`);
  } catch (e) { console.error(e); busy(false); toast('הדחיסה נכשלה: ' + (e.message || e)); }
  finally { busy(false); try { doc && doc.destroy(); } catch (_) {} }
}

// 6.2: a "חותמת" button in the signature window
{ const row = document.querySelector('#sigDlg .sig-row'), go = document.getElementById('sigUse');
  if (row && go) { const b = h('<button class="btn" id="stampBtn" type="button" title="חותמת עגולה או מלבנית עם שם העסק ומספר ח.פ">🔵 חותמת עסק</button>'); b.onclick = () => { document.getElementById('sigDlg').close(); stampDialog(); }; row.querySelector('.grow').after(b);
    const k = h('<button class="btn" id="kioskBtn" type="button" title="מעבירים את המכשיר לחותם: הוא רואה את העמוד, כותב שם וחותם">🤝 החותם לידי</button>'); k.onclick = () => { document.getElementById('sigDlg').close(); inPerson(); }; b.after(k); } }

window.SIMAN_PRO = { autoRedact, compressDialog, stampDialog, inPerson, _test: { pageHits, validId, luhn, compress, LEVELS } };
})();
