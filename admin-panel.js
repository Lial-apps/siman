/* =====================================================================
   סימן 4.9 — לוח ניהול בתוך האפליקציה
   מופיע רק כשהמייל בהגדרות הוא מייל של מנהל.
   הנתונים עצמם מוגנים במפתח הניהול (ADMIN_KEY) בשרת: מזינים אותו פעם אחת
   במכשיר, והוא נשמר רק במכשיר הזה.
   ===================================================================== */
(function () {
  const ADMIN_EMAILS = ['lioralmo1@gmail.com'];
  const API = 'https://siman-register.siman-app.workers.dev/admin';
  const KEY_STORE = 'siman.adminKey';
  const REFRESH_MS = 60000;

  const el = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const getKey = () => { try { return localStorage.getItem(KEY_STORE) || ''; } catch (_) { return ''; } };
  const setKey = (k) => { try { k ? localStorage.setItem(KEY_STORE, k) : localStorage.removeItem(KEY_STORE); } catch (_) {} };
  const userMail = () => { try { return String((typeof USER !== 'undefined' && USER && USER.mail) || '').trim().toLowerCase(); } catch (_) { return ''; } };
  const isAdmin = () => ADMIN_EMAILS.includes(userMail());

  const FEAT = {
    translate: 'תרגום', translate_open: 'פתיחת תרגום', ocr: 'זיהוי טקסט', ocr_all: 'זיהוי טקסט לכל המסמך',
    convert_open: 'פתיחת המרה', word_edit: 'עריכת וורד', print: 'הדפסה', save: 'שמירה', save_as: 'שמירה בשם',
    save_drive: 'שמירה ל־Drive', save_icloud: 'שמירה ל־iCloud', share_open: 'פתיחת שיתוף', pages: 'ניהול עמודים',
    merge: 'מיזוג', split: 'פיצול', save_pages: 'שמירת עמודים נבחרים', paste: 'הדבקה', command_bar: 'שורת פקודות',
    signature: 'חתימה', ai_open: 'פתיחת העוזר', ai_ask: 'שאלה למסמך', read_aloud: 'הקראה', handwriting: 'כתב יד',
    install: 'התקנה', auto_names: 'שמות אוטומטיים', link_sign: 'קישור לחתימה', link_view: 'קישור לצפייה'
  };
  const featName = (k) => FEAT[k] || k.replace(/^open_/, 'פתיחת קובץ ').replace(/^tool_/, 'כלי: ').replace(/^share_/, 'שיתוף: ')
    .replace(/^convert_/, 'המרה ל־').replace(/^lang_/, 'שפה: ').replace(/^quick_/, 'פעולה מהירה: ');

  /* ---------- עיצוב ---------- */
  const css = `
  #admPanel{--a-ink:#16213e;--a-ink2:#2a3760;--a-mk:#ffd93d;--a-mks:#fff4b8;--a-bg:#f6f4ee;--a-card:#fff;--a-tx:#1c2238;--a-mu:#6b7190;--a-ln:#e4e1d6;--a-bad:#d64545;
    position:fixed;inset:0;z-index:2147483000;background:var(--a-bg);color:var(--a-tx);direction:rtl;text-align:right;
    font-family:"Heebo","Assistant","Rubik",Arial,sans-serif;font-size:15px;overflow:auto;
    padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}
  @media (prefers-color-scheme:dark){#admPanel{--a-ink:#0e1424;--a-ink2:#5b6db0;--a-mks:#3a3418;--a-bg:#11141d;--a-card:#1a1f2c;--a-tx:#e8eaf3;--a-mu:#9aa0bb;--a-ln:#2a3042}}
  #admPanel *{box-sizing:border-box}
  #admPanel .a-top{background:var(--a-ink);color:#fff;padding:12px 16px;display:flex;align-items:center;gap:10px;flex-wrap:wrap;position:sticky;top:0;z-index:2}
  #admPanel .a-top h2{margin:0;font-size:19px;font-weight:800}
  #admPanel .a-top mark{background:linear-gradient(transparent 55%,var(--a-mk) 55%);color:inherit;padding:0 3px}
  #admPanel .a-sp{flex:1}
  #admPanel .a-live{font-size:13px;opacity:.85;display:flex;align-items:center;gap:6px}
  #admPanel .a-dot{width:9px;height:9px;border-radius:50%;background:#777}
  #admPanel .a-dot.on{background:#3ddc84}
  #admPanel button{font:inherit;border:0;border-radius:10px;padding:7px 13px;cursor:pointer;background:var(--a-mk);color:#16213e;font-weight:700}
  #admPanel button.a-gh{background:transparent;color:inherit;border:1px solid currentColor}
  #admPanel .a-main{max-width:1200px;margin:0 auto;padding:16px}
  #admPanel .a-card{background:var(--a-card);border:1px solid var(--a-ln);border-radius:14px;padding:15px}
  #admPanel .a-kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin-bottom:12px}
  #admPanel .a-n{font-size:30px;font-weight:800;line-height:1.1}
  #admPanel .a-l{color:var(--a-mu);font-size:13px}
  #admPanel .a-hl{background:var(--a-mks);border-color:var(--a-mk)}
  #admPanel .a-g2{display:grid;grid-template-columns:2fr 1fr 1fr;gap:10px;margin-bottom:12px}
  @media (max-width:900px){#admPanel .a-g2{grid-template-columns:1fr}}
  #admPanel h3{margin:0 0 10px;font-size:15px}
  #admPanel .a-chart{display:flex;align-items:flex-end;gap:3px;height:150px;border-bottom:1px solid var(--a-ln);padding-top:18px}
  #admPanel .a-bar{flex:1;background:var(--a-ink2);border-radius:4px 4px 0 0;min-height:2px;position:relative}
  #admPanel .a-bar.t{background:var(--a-mk)}
  #admPanel .a-bar span{position:absolute;top:-17px;inset-inline:0;text-align:center;font-size:11px;color:var(--a-mu)}
  #admPanel .a-axis{display:flex;justify-content:space-between;color:var(--a-mu);font-size:12px;margin-top:4px}
  #admPanel .a-hb{margin:7px 0}
  #admPanel .a-hb .a-t{display:flex;justify-content:space-between;gap:8px;font-size:14px}
  #admPanel .a-tr{height:8px;background:var(--a-ln);border-radius:5px;overflow:hidden;margin-top:3px}
  #admPanel .a-fl{height:100%;background:var(--a-mk);border-radius:5px}
  #admPanel .a-tools{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin-bottom:10px}
  #admPanel input{font:inherit;padding:9px 12px;border-radius:10px;border:1px solid var(--a-ln);background:var(--a-bg);color:var(--a-tx)}
  #admPanel .a-tools input{flex:1;min-width:180px}
  #admPanel .a-wrap{overflow-x:auto;border:1px solid var(--a-ln);border-radius:12px}
  #admPanel table{border-collapse:collapse;width:100%;font-size:14px;background:var(--a-card)}
  #admPanel th,#admPanel td{padding:8px 10px;border-bottom:1px solid var(--a-ln);text-align:right;white-space:nowrap}
  #admPanel th{background:var(--a-card);font-weight:700}
  #admPanel th[data-s]{cursor:pointer;user-select:none}
  #admPanel tbody tr:hover{background:var(--a-mks)}
  #admPanel tbody tr.a-new td:first-child{box-shadow:inset -4px 0 0 var(--a-mk)}
  #admPanel tbody tr.a-row{cursor:pointer}
  #admPanel .a-mu{color:var(--a-mu)}
  #admPanel .a-ltr{direction:ltr;unicode-bidi:isolate}
  #admPanel .a-badge{display:inline-block;background:var(--a-mk);color:#16213e;border-radius:6px;padding:0 6px;font-size:12px;font-weight:700;margin-inline-start:6px}
  #admPanel .a-err{color:var(--a-bad);margin-top:10px;white-space:pre-line}
  #admPanel .a-login{max-width:430px;margin:50px auto}
  #admPanel .a-login input{width:100%;direction:ltr;margin-top:6px}
  #admPanel .a-sec{margin-top:12px}
  #admPanel .a-det td{white-space:normal;background:var(--a-bg)}
  #admPanel .a-chips span{display:inline-block;border:1px solid var(--a-ln);border-radius:8px;padding:2px 8px;margin:3px;font-size:13px;background:var(--a-card)}
  #admBtn.hidden,#admSetBtn.hidden{display:none!important}
  #admSetBtn{margin-top:10px}
  `;
  const st = document.createElement('style'); st.textContent = css; document.head.appendChild(st);

  /* ---------- כפתורי כניסה ---------- */
  const icon = '<svg viewBox="0 0 24 24" aria-hidden="true" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v18h18"/><rect x="7" y="11" width="3" height="6"/><rect x="12" y="7" width="3" height="10"/><rect x="17" y="13" width="3" height="4"/></svg>';
  const setBtn = el('setBtn');
  const topBtn = document.createElement('button');
  topBtn.type = 'button'; topBtn.id = 'admBtn'; topBtn.className = 'icon-btn hidden';
  topBtn.title = 'לוח ניהול'; topBtn.setAttribute('aria-label', 'לוח ניהול'); topBtn.innerHTML = icon;
  if (setBtn && setBtn.parentNode) setBtn.parentNode.insertBefore(topBtn, setBtn);

  const delBox = el('setDelete') && el('setDelete').closest('.set-del');
  const setPanelBtn = document.createElement('button');
  setPanelBtn.type = 'button'; setPanelBtn.id = 'admSetBtn'; setPanelBtn.className = 'btn hidden';
  setPanelBtn.textContent = 'לוח ניהול (מנהל)';
  if (delBox && delBox.parentNode) delBox.parentNode.insertBefore(setPanelBtn, delBox);

  function syncButtons() {
    const a = isAdmin();
    topBtn.classList.toggle('hidden', !a);
    setPanelBtn.classList.toggle('hidden', !a);
  }
  syncButtons();
  setInterval(syncButtons, 2000);
  document.addEventListener('click', () => setTimeout(syncButtons, 300), true);
  topBtn.onclick = open;
  setPanelBtn.onclick = open;

  /* ---------- הלוח ---------- */
  let panel = null, data = null, timer = null, q = '', sortK = 'first', sortD = -1, openRow = '';

  function open() {
    if (!isAdmin()) return;
    if (!panel) {
      panel = document.createElement('div');
      panel.id = 'admPanel'; panel.setAttribute('role', 'dialog'); panel.setAttribute('aria-label', 'לוח ניהול');
      panel.innerHTML = `
        <div class="a-top">
          <h2><mark>סימן</mark> · לוח ניהול</h2><div class="a-sp"></div>
          <div class="a-live"><span class="a-dot" id="admDot"></span><span id="admUpd"></span></div>
          <button type="button" class="a-gh" id="admRef">רענון</button>
          <button type="button" class="a-gh" id="admClose">סגירה</button>
        </div>
        <div class="a-main" id="admBody"></div>`;
      document.body.appendChild(panel);
      el('admClose').onclick = close;
      el('admRef').onclick = () => load();
      panel.addEventListener('keydown', (e) => { if (e.key === 'Escape') close(); });
    }
    panel.style.display = '';
    document.documentElement.style.overflow = 'hidden';
    if (getKey()) { if (data) render(); load(); } else showLogin('');
    clearInterval(timer);
    timer = setInterval(() => { if (!document.hidden && getKey()) load(true); }, REFRESH_MS);
  }
  function close() {
    if (panel) panel.style.display = 'none';
    document.documentElement.style.overflow = '';
    clearInterval(timer);
  }

  function showLogin(err) {
    el('admDot').classList.remove('on'); el('admUpd').textContent = '';
    el('admBody').innerHTML = `
      <div class="a-card a-login">
        <h3>חיבור ראשון במכשיר הזה</h3>
        <p class="a-mu" style="margin:0 0 8px;font-size:14px;line-height:1.6">מזינים פעם אחת את מפתח הניהול (ADMIN_KEY) שהגדרת ב־Cloudflare. מכאן והלאה הלוח ייפתח ישר.</p>
        <input type="password" id="admKey" autocomplete="current-password" placeholder="ADMIN_KEY">
        <div style="margin-top:12px"><button type="button" id="admGo">חיבור</button></div>
        <div class="a-err">${esc(err)}</div>
      </div>`;
    const go = () => { const k = el('admKey').value.trim(); if (!k) return; setKey(k); load(); };
    el('admGo').onclick = go;
    el('admKey').addEventListener('keydown', (e) => { if (e.key === 'Enter') go(); });
    setTimeout(() => el('admKey') && el('admKey').focus(), 50);
  }

  async function load(silent) {
    const key = getKey(); if (!key) return showLogin('');
    if (!silent) el('admUpd').textContent = 'טוען…';
    try {
      const r = await fetch(API, { method: 'POST', headers: { 'Content-Type': 'text/plain' }, body: JSON.stringify({ key }), cache: 'no-store' });
      if (r.status === 403) { setKey(''); data = null; return showLogin('המפתח לא נכון. נסו שוב.'); }
      if (!r.ok) throw new Error('http ' + r.status);
      data = await r.json();
      el('admDot').classList.add('on');
      el('admUpd').textContent = 'עודכן ' + new Date().toLocaleTimeString('he-IL', { hour: '2-digit', minute: '2-digit' });
      render();
    } catch (e) {
      el('admDot').classList.remove('on');
      if (!data) {
        el('admBody').innerHTML = `<div class="a-card a-login"><h3>אין חיבור לשרת</h3>
          <p class="a-mu" style="line-height:1.6">בדקו שיש אינטרנט ושה־worker ‏siman-register עודכן לגרסה 4.9 (עם הכתובת ‎/admin).</p>
          <button type="button" id="admRetry">לנסות שוב</button> <button type="button" class="a-gh" id="admNewKey" style="color:var(--a-tx)">להזין מפתח אחר</button></div>`;
        el('admRetry').onclick = () => load();
        el('admNewKey').onclick = () => { setKey(''); showLogin(''); };
      } else el('admUpd').textContent = 'החיבור נכשל · מנסה שוב';
    }
  }

  const DAY = 864e5;
  const dk = (d) => d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate();
  const D = (v) => { const d = v ? new Date(v) : null; return d && !isNaN(d) ? d : null; };
  const fd = (v) => { const d = D(v); return d ? d.toLocaleDateString('he-IL', { day: '2-digit', month: '2-digit', year: '2-digit' }) : ''; };
  const fdt = (v) => { const d = D(v); return d ? d.toLocaleString('he-IL', { day: '2-digit', month: '2-digit', hour: '2-digit', minute: '2-digit' }) : ''; };
  const ago = (v) => {
    const d = D(v); if (!d) return '';
    const m = Math.round((Date.now() - d) / 60000);
    if (m < 2) return 'עכשיו'; if (m < 60) return `לפני ${m} דק׳`;
    const h = Math.round(m / 60); if (h < 24) return `לפני ${h} שע׳`;
    const dd = Math.round(h / 24); return dd === 1 ? 'אתמול' : `לפני ${dd} ימים`;
  };

  function bars(map, total) {
    const ent = Object.entries(map).sort((a, b) => b[1] - a[1]).slice(0, 8);
    if (!ent.length) return '<p class="a-mu">אין נתונים עדיין.</p>';
    return ent.map(([k, n]) => { const p = Math.round(n / Math.max(1, total) * 100);
      return `<div class="a-hb"><div class="a-t"><bdi>${esc(k)}</bdi><span>${n} · ${p}%</span></div><div class="a-tr"><div class="a-fl" style="width:${p}%"></div></div></div>`; }).join('');
  }

  function render() {
    if (!data || !panel) return;
    const users = data.users || [], now = Date.now(), today = dk(new Date());
    const since = (f, n) => users.filter((u) => { const d = D(u[f]); return d && now - d <= n * DAY; }).length;
    const plat = {}, ver = {};
    users.forEach((u) => { const p = u.platform || 'לא ידוע'; plat[p] = (plat[p] || 0) + 1; const v = u.ver || 'לא ידוע'; ver[v] = (ver[v] || 0) + 1; });

    const counts = {}; users.forEach((u) => { const d = D(u.first); if (d) counts[dk(d)] = (counts[dk(d)] || 0) + 1; });
    const days = []; for (let i = 29; i >= 0; i--) { const d = new Date(now - i * DAY); days.push({ d, n: counts[dk(d)] || 0 }); }
    const max = Math.max(1, ...days.map((x) => x.n));
    const chart = days.map((x, i) => `<div class="a-bar${i === 29 ? ' t' : ''}" style="height:${x.n / max * 100}%" title="${x.d.toLocaleDateString('he-IL')}: ${x.n}">${x.n ? `<span>${x.n}</span>` : ''}</div>`).join('');

    const feats = (data.stats && data.stats.features) || [];
    const featRows = feats.slice(0, 40).map(([k, t, w]) => `<tr><td>${esc(featName(k))}</td><td>${t}</td><td>${w}</td></tr>`).join('')
      || '<tr><td colspan="3" class="a-mu">עוד אין נתוני שימוש.</td></tr>';

    el('admBody').innerHTML = `
      <div class="a-kpis">
        <div class="a-card a-hl"><div class="a-n">${users.length}</div><div class="a-l">משתמשים רשומים</div></div>
        <div class="a-card"><div class="a-n">${users.filter((u) => D(u.first) && dk(D(u.first)) === today).length}</div><div class="a-l">נרשמו היום</div></div>
        <div class="a-card"><div class="a-n">${since('first', 7)}</div><div class="a-l">נרשמו ב־7 ימים</div></div>
        <div class="a-card"><div class="a-n">${since('last', 1)}</div><div class="a-l">פעילים היום</div></div>
        <div class="a-card"><div class="a-n">${since('last', 7)}</div><div class="a-l">פעילים ב־7 ימים</div></div>
        <div class="a-card"><div class="a-n">${users.filter((u) => u.news).length}</div><div class="a-l">מסכימים לקבל עדכונים</div></div>
      </div>
      <div class="a-g2">
        <div class="a-card"><h3>הרשמות ב־30 הימים האחרונים</h3><div class="a-chart">${chart}</div>
          <div class="a-axis"><span>${days[0].d.toLocaleDateString('he-IL', { day: 'numeric', month: 'numeric' })}</span><span>היום</span></div></div>
        <div class="a-card"><h3>לפי פלטפורמה</h3>${bars(plat, users.length)}</div>
        <div class="a-card"><h3>לפי גרסת אפליקציה</h3>${bars(ver, users.length)}</div>
      </div>
      <div class="a-card">
        <h3>כל המשתמשים</h3>
        <div class="a-tools">
          <input type="search" id="admQ" placeholder="חיפוש לפי מייל, שם או פלטפורמה" value="${esc(q)}">
          <span class="a-mu" id="admCount"></span>
          <button type="button" class="a-gh" id="admCsv" style="color:var(--a-tx)">הורדת CSV</button>
        </div>
        <div class="a-wrap"><table><thead><tr>
          <th>#</th><th data-s="mail">מייל</th><th data-s="name">שם</th><th data-s="platform">פלטפורמה</th>
          <th data-s="first">נרשם</th><th data-s="last">פעיל לאחרונה</th><th data-s="days">ימי פעילות</th>
          <th data-s="total">פעולות</th><th data-s="ver">גרסה</th><th data-s="news">עדכונים</th>
        </tr></thead><tbody id="admRows"></tbody></table></div>
        <p class="a-mu" style="font-size:13px;margin:8px 0 0">לוחצים על שורה כדי לראות באילו יכולות המשתמש השתמש.</p>
      </div>
      <div class="a-card a-sec"><h3>שימוש ביכולות</h3>
        <div class="a-wrap"><table><thead><tr><th>יכולת</th><th>פעמים</th><th>משתמשים</th></tr></thead><tbody>${featRows}</tbody></table></div></div>
      <div class="a-sec a-mu" style="font-size:13px">מחובר כמנהל: <bdi class="a-ltr">${esc(userMail())}</bdi> ·
        <a href="#" id="admForget" style="color:inherit">ניתוק המפתח מהמכשיר הזה</a></div>`;

    el('admQ').addEventListener('input', (e) => { q = e.target.value; rows(); });
    el('admCsv').onclick = csv;
    el('admForget').onclick = (e) => { e.preventDefault(); setKey(''); data = null; showLogin(''); };
    panel.querySelectorAll('th[data-s]').forEach((th) => th.onclick = () => {
      const k = th.dataset.s; if (sortK === k) sortD *= -1; else { sortK = k; sortD = -1; } rows();
    });
    rows();
  }

  function list() {
    const s = q.trim().toLowerCase();
    const users = (data.users || []).filter((u) => !s || [u.mail, u.name, u.platform, u.lang, u.ver].some((v) => String(v || '').toLowerCase().includes(s)));
    return users.sort((a, b) => {
      let x = a[sortK], y = b[sortK];
      if (sortK === 'first' || sortK === 'last') { x = D(x) ? +D(x) : 0; y = D(y) ? +D(y) : 0; }
      if (typeof x === 'number' || typeof y === 'number' || typeof x === 'boolean') return ((+x || 0) - (+y || 0)) * sortD;
      return String(x || '').localeCompare(String(y || ''), 'he') * sortD;
    });
  }

  function rows() {
    const L = list(), now = Date.now();
    el('admCount').textContent = q ? `${L.length} מתוך ${(data.users || []).length}` : `${L.length} משתמשים`;
    panel.querySelectorAll('th[data-s]').forEach((th) => {
      th.textContent = th.textContent.replace(/ [▲▼]$/, '') + (th.dataset.s === sortK ? (sortD > 0 ? ' ▲' : ' ▼') : '');
    });
    el('admRows').innerHTML = L.map((u, i) => {
      const nw = D(u.first) && now - D(u.first) < DAY;
      const row = `<tr class="a-row${nw ? ' a-new' : ''}" data-m="${esc(u.mail)}">
        <td class="a-mu">${i + 1}</td>
        <td class="a-ltr">${esc(u.mail)}${nw ? '<span class="a-badge">חדש</span>' : ''}</td>
        <td>${esc(u.name)}</td><td><bdi>${esc(u.platform)}</bdi></td>
        <td>${fd(u.first)}</td><td title="${fdt(u.last)}">${ago(u.last)}</td>
        <td>${u.days || 0}</td><td>${u.total || 0}</td><td class="a-ltr">${esc(u.ver)}</td><td>${u.news ? 'כן' : '–'}</td></tr>`;
      if (openRow !== u.mail) return row;
      const us = Object.entries(u.usage || {}).sort((a, b) => b[1] - a[1]);
      const chips = us.length ? us.map(([k, n]) => `<span>${esc(featName(k))}: ${n}</span>`).join('') : '<span>עוד לא נשלחו נתוני שימוש</span>';
      return row + `<tr class="a-det"><td></td><td colspan="9">
        <div class="a-mu" style="font-size:13px;margin-bottom:4px">שפה: ${esc(u.lang || '–')} · נרשם: ${fdt(u.first)} · פעיל לאחרונה: ${fdt(u.last)}${u.plan ? ' · מסלול: ' + esc(u.plan) : ''}</div>
        <div class="a-chips">${chips}</div></td></tr>`;
    }).join('') || '<tr><td colspan="10" class="a-mu">אין משתמשים</td></tr>';
    el('admRows').querySelectorAll('tr.a-row').forEach((tr) => tr.onclick = () => { openRow = openRow === tr.dataset.m ? '' : tr.dataset.m; rows(); });
  }

  function csv() {
    const H = ['mail', 'name', 'platform', 'lang', 'first', 'last', 'days', 'total', 'ver', 'news'];
    const e = (v) => '"' + String(v ?? '').replace(/"/g, '""') + '"';
    const text = '\uFEFF' + [H.join(','), ...list().map((u) => H.map((h) => e(h === 'news' ? (u.news ? 'yes' : 'no') : u[h])).join(','))].join('\r\n');
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([text], { type: 'text/csv;charset=utf-8' }));
    a.download = 'siman-users-' + new Date().toISOString().slice(0, 10) + '.csv';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 1500);
  }

  window.simanOpenAdmin = open;
})();
