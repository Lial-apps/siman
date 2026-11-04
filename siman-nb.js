/* סימן 5.9.5 — עיצוב "מחברת" (ראו siman-nb.css).
   הקובץ לא מחליף שום פעולה של האפליקציה. הוא רק מוסיף למסך הפתיחה:
   1. אזור גרירה עם כפתור "פותחים קובץ" (לוחץ על אותה פעולה כמו האריח הקודם, data-v7="open").
   2. הדבקת קובץ: Ctrl V במחשב, כפתור "הדבקה" בטלפון. כשהדפדפן לא נותן לקרוא את הלוח מופיעה מסגרת ללחיצה ארוכה.
   3. אייקון לכל מדבקה, והדגשה צהובה של "נסמן" בכותרת.
   4. חצים לגלילת "ממשיכים מאיפה שעצרנו" בטלפון.
   מסך הפתיחה נבנה מחדש כשמחליפים שפה, ולכן הכול נבדק שוב בכל שינוי. */
(function () {
  'use strict';
  var root = document.documentElement;
  root.classList.add('nb');

  var STR = {
    he: { open: 'פותחים קובץ ✦', paste: 'הדבקה', drag: 'או גוררים קובץ לכאן', ctrlv: 'או מדביקים עם', narrow: 'מהטלפון, מהדרייב או מהמייל', types: 'PDF · Word · Excel · PowerPoint', hot: 'משחררים כאן ✦',
      field: 'לחיצה ארוכה כאן ← <b>הדבקה</b>', fieldTip: 'לוחצים לחיצה ארוכה על המסגרת ובוחרים "הדבקה"', noFile: 'בלוח יש טקסט ולא קובץ. מעתיקים את הקובץ עצמו ואז לוחצים "הדבקה"', empty: 'אין כרגע קובץ בלוח. מעתיקים קובץ ואז לוחצים "הדבקה"', pasted: 'קובץ מודבק', prev: 'הקבצים הקודמים', next: 'הקבצים הבאים' },
    ar: { open: 'فتح ملف ✦', paste: 'لصق', drag: 'أو اسحب ملفًا إلى هنا', ctrlv: 'أو الصق بـ', narrow: 'من الهاتف أو Drive أو البريد', types: 'PDF · Word · Excel · PowerPoint', hot: 'أفلته هنا ✦',
      field: 'اضغط مطولًا هنا ← <b>لصق</b>', fieldTip: 'اضغط مطولًا على الإطار واختر "لصق"', noFile: 'في الحافظة نص وليس ملفًا. انسخ الملف نفسه ثم اضغط "لصق"', empty: 'لا يوجد ملف في الحافظة. انسخ ملفًا ثم اضغط "لصق"', pasted: 'ملف ملصق', prev: 'الملفات السابقة', next: 'الملفات التالية' },
    ru: { open: 'Открыть файл ✦', paste: 'Вставить', drag: 'или перетащите файл сюда', ctrlv: 'или вставьте', narrow: 'С телефона, из Диска или почты', types: 'PDF · Word · Excel · PowerPoint', hot: 'Отпустите здесь ✦',
      field: 'Нажмите и удерживайте здесь → <b>Вставить</b>', fieldTip: 'Нажмите и удерживайте рамку и выберите «Вставить»', noFile: 'В буфере текст, а не файл. Скопируйте сам файл и нажмите «Вставить»', empty: 'В буфере нет файла. Скопируйте файл и нажмите «Вставить»', pasted: 'Вставленный файл', prev: 'Предыдущие файлы', next: 'Следующие файлы' },
    en: { open: 'Open a file ✦', paste: 'Paste', drag: 'or drop a file here', ctrlv: 'or paste with', narrow: 'From your phone, Drive or email', types: 'PDF · Word · Excel · PowerPoint', hot: 'Drop it here ✦',
      field: 'Long-press here → <b>Paste</b>', fieldTip: 'Long-press the box and choose “Paste”', noFile: 'The clipboard has text, not a file. Copy the file itself, then tap “Paste”', empty: 'There is no file on the clipboard. Copy a file, then tap “Paste”', pasted: 'Pasted file', prev: 'Previous files', next: 'Next files' }
  };
  var EMOJI = { translate: '🌍', scan: '📷', transcribe: '🎙️', sign: '✍️', word: '📝', excel: '📊', slides: '📽️', conv: '🔁', ask: '💬', send: '📤', merge: '✂️', record: '⏺️' };
  var EXT = { 'application/pdf': '.pdf', 'image/png': '.png', 'image/jpeg': '.jpg', 'image/webp': '.webp',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': '.docx', 'application/msword': '.doc',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': '.xlsx', 'application/vnd.openxmlformats-officedocument.presentationml.presentation': '.pptx' };
  var ICON_FILE = '<svg viewBox="0 0 40 40" aria-hidden="true"><path d="M10 4h14l8 8v24H10z" fill="#fff" stroke="currentColor" stroke-width="2.5" stroke-linejoin="round"/><path d="M24 4v8h8" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linejoin="round"/><path d="M21 17v11M16.5 23.5L21 28l4.5-4.5" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/></svg>';
  var ICON_PASTE = '<svg viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><rect x="6" y="4" width="12" height="17" rx="2"/><path d="M9 4V3h6v1M9 10h6M9 14h6M9 18h3"/></svg>';

  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function langOf(home) { var l = home && home.getAttribute('lang'); return STR[l] ? l : 'he'; }
  function say(t) { if (typeof window.toast === 'function') window.toast(t); }
  function homeShown() { var h = document.querySelector('.v7-home'); return !!(h && h.offsetParent) && !(typeof S !== 'undefined' && S && S.doc); }
  function openAny(file) {
    var f = window.openFile || (typeof openFile === 'function' ? openFile : null);
    if (f) f(file);
  }

  /* ---------------------------------------------------------------- 1-3. מסך הפתיחה */
  function decorate(home) {
    if (home.getAttribute('data-nb')) return;
    home.setAttribute('data-nb', '1');
    var L = STR[langOf(home)];

    var h1 = home.querySelector('h1');
    if (h1 && langOf(home) === 'he' && h1.textContent.trim() === 'מה נסמן היום?') h1.innerHTML = 'מה <span class="nb-tilt">נסמן</span><br>היום?';

    var drop = document.createElement('div');
    drop.className = 'nb-drop';
    // data-v7="open" on the zone itself: a click anywhere in it runs the app's own "open file" (siman-v7.js listens on the home screen)
    drop.setAttribute('data-v7', 'open');
    drop.innerHTML = '<div class="nb-btns"><button type="button" class="nb-open">' + esc(L.open) + '</button>' +
      '<button type="button" class="nb-paste" data-nb-paste>' + ICON_PASTE + esc(L.paste) + '</button></div>' +
      '<div class="nb-dt">' + ICON_FILE + '<span><b class="nb-idle"><span class="nb-only-wide">' + esc(L.drag) + '</span><span class="nb-only-narrow">' + esc(L.narrow) + '</span></b>' +
      '<b class="nb-hot">' + esc(L.hot) + '</b><small class="nb-only-wide">' + esc(L.ctrlv) + ' <kbd>Ctrl V</kbd></small><small>' + esc(L.types) + '</small></span></div>';
    var field = document.createElement('div');
    field.className = 'nb-field';
    field.setAttribute('contenteditable', 'true');
    field.setAttribute('inputmode', 'none');
    field.setAttribute('aria-label', L.paste);
    field.innerHTML = L.field;
    var sub = home.querySelector('.v7-sub');
    if (sub) { sub.after(drop); drop.after(field); } else { home.prepend(field); home.prepend(drop); }

    drop.querySelector('[data-nb-paste]').addEventListener('click', function (e) { e.stopPropagation(); pasteFromClipboard(home, field); });
    field.addEventListener('click', function (e) { e.stopPropagation(); });
    field.addEventListener('beforeinput', function (e) { if (e.inputType !== 'insertFromPaste') e.preventDefault(); });

    home.querySelectorAll('.v7-bento [data-v7]').forEach(function (t) {
      var k = t.getAttribute('data-v7');
      if (!EMOJI[k] || t.querySelector('.nb-e')) return;
      var s = document.createElement('span');
      s.className = 'nb-e'; s.setAttribute('aria-hidden', 'true'); s.textContent = EMOJI[k];
      t.prepend(s);
      // the hand-drawn circle that draws itself around the sticker on hover (siman-nb.css, computers only)
      t.insertAdjacentHTML('beforeend', '<svg class="nb-ring" viewBox="0 0 200 160" preserveAspectRatio="none" aria-hidden="true"><path d="M110 6 C 170 4, 198 40, 194 84 C 190 130, 140 156, 92 154 C 40 152, 4 122, 6 78 C 8 34, 50 8, 120 12"/></svg>');
    });
  }

  /* ---------------------------------------------------------------- 2. הדבקת קובץ */
  async function pasteFromClipboard(home, field) {
    var L = STR[langOf(home)];
    try {
      var items = await navigator.clipboard.read();
      for (var i = 0; i < items.length; i++) {
        var t = items[i].types.filter(function (x) { return x.indexOf('text/') !== 0; })[0];
        if (t) { var blob = await items[i].getType(t); openAny(new File([blob], L.pasted + (EXT[t] || ''), { type: t })); return; }
      }
      say(items.length ? L.noFile : L.empty);
    } catch (_) {
      // the browser did not let us read the clipboard: a long-press on this box opens the phone's own "Paste" menu
      field.classList.add('show'); field.focus(); say(L.fieldTip);
    }
  }
  document.addEventListener('paste', function (e) {
    if (!homeShown()) return;
    var t = e.target;
    if (t && t !== document.querySelector('.nb-field') && (t.isContentEditable || /INPUT|TEXTAREA/.test(t.tagName))) return;
    var files = e.clipboardData && e.clipboardData.files;
    var f = files && files[0];
    var fld = document.querySelector('.nb-field');
    if (f) { e.preventDefault(); if (fld) fld.classList.remove('show'); openAny(f.name ? f : new File([f], STR[langOf(document.querySelector('.v7-home'))].pasted + (EXT[f.type] || ''), { type: f.type })); return; }
    if (t === fld) { e.preventDefault(); say(STR[langOf(document.querySelector('.v7-home'))].noFile); }
  });

  /* ---------------------------------------------------------------- 4. חצים לקבצים האחרונים (בטלפון) */
  function decorateRecent(home) {
    var box = home.querySelector('.ux-recent');
    if (!box || box.querySelector('.nb-nav')) return;
    var hd = box.querySelector('.hd'), row = box.querySelector('.ls');
    if (!hd || !row) return;
    var L = STR[langOf(home)], rtl = home.getAttribute('dir') !== 'ltr';
    var nav = document.createElement('span');
    nav.className = 'nb-nav';
    nav.innerHTML = '<button type="button" data-d="-1" aria-label="' + esc(L.prev) + '">' + (rtl ? '→' : '←') + '</button><button type="button" data-d="1" aria-label="' + esc(L.next) + '">' + (rtl ? '←' : '→') + '</button>';
    hd.insertBefore(nav, hd.querySelector('[data-clear]'));
    var bs = nav.querySelectorAll('button');
    var step = function () { var c = row.querySelector('.ux-rf'); return c ? c.offsetWidth + 16 : 200; };
    var sync = function () { var x = Math.abs(row.scrollLeft), max = row.scrollWidth - row.clientWidth; bs[0].disabled = x < 24; bs[1].disabled = x > max - 24; };
    bs.forEach(function (b) { b.addEventListener('click', function () { var d = +b.getAttribute('data-d'); row.scrollBy({ left: d * step() * (rtl ? -1 : 1) }); }); });
    row.addEventListener('scroll', sync, { passive: true });
    addEventListener('resize', sync);
    sync();
  }

  /* ---------------------------------------------------------------- 5. הסרגל העליון בשתי שורות
     שורה עליונה: הלוגו מימין, ומשמאל חיפוש, שיתוף, שפה, צבע והגדרות. שורה שנייה: "קלמר" אחד עם הכלים ואחריהם הפעולות.
     הכפתורים עוברים כמו שהם (אותו id ואותן פעולות). כפתורים שקבצים אחרים מוסיפים "ליד setBtn" או "אחרי trBtn" נכנסים
     גם הם למקום הנכון, כי השכן שלהם עבר איתם. */
  var TOP_IDS = ['cmdBtn', 'shareBtn', 'uxMore', 'v7LangBtn', 'v7ColorBtn', 'admBtn', 'setBtn', 'installBtn'];
  function buildBar() {
    var bar = document.querySelector('header.bar');
    if (!bar || bar.querySelector('.nb-case')) return !!bar;
    var tools = document.getElementById('tools'), undo = document.getElementById('undo');
    var acts = undo && undo.closest('.group');
    if (!tools || !acts) return false;
    var top = document.createElement('div');
    top.className = 'group nb-top';
    var kase = document.createElement('div');
    kase.className = 'nb-case';
    var sep = document.createElement('span');
    sep.className = 'nb-sep'; sep.setAttribute('aria-hidden', 'true');
    bar.insertBefore(top, tools);
    bar.insertBefore(kase, tools);
    var zoom = bar.querySelector('.group.zoom');
    if (zoom) top.appendChild(zoom);
    TOP_IDS.forEach(function (id) { var b = document.getElementById(id); if (b) top.appendChild(b); });
    kase.appendChild(tools); kase.appendChild(sep); kase.appendChild(acts);
    return true;
  }
  // the language and colour buttons (siman-v7.js) and the phone "more" button (siman-ux.js) are added when those files start,
  // before this one: wait for them, but never longer than a few seconds
  (function waitBar(n) {
    var ready = document.getElementById('v7ColorBtn') || n > 40;
    if (ready && (buildBar() || n > 60)) return;
    setTimeout(function () { waitBar(n + 1); }, 100);
  })(0);

  var queued = false;
  function tick() {
    queued = false;
    var home = document.querySelector('.v7-home');
    if (!home) return;
    decorate(home);
    decorateRecent(home);
  }
  new MutationObserver(function () { if (!queued) { queued = true; requestAnimationFrame(tick); } }).observe(document.body, { childList: true, subtree: true });
  tick();
})();
