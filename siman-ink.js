/* סימן 6.9.0 — עיצוב "דיו ומים" (ראו siman-ink.css).
   הקובץ לא מחליף שום פעולה של האפליקציה. הוא רק:
   1. מסמן את <html> במילה ink, כדי ש-siman-ink.css יפעל.
   2. עוטף את שם הכלי ואת התיאור שלו במסך הפתיחה ב-<span class="ink-m">, כדי שמשיכת המרקר תעבור על המלל עצמו
      (ולא על כל רוחב השורה) כשעוברים מעל הכלי. הטקסט עצמו לא משתנה.
   מסך הפתיחה נבנה מחדש כשמחליפים שפה, ולכן הכול נבדק שוב בכל שינוי (כמו ב-siman-nb.js). */
(function () {
  'use strict';
  document.documentElement.classList.add('ink');

  function wrap(el) {
    if (!el || el.querySelector('.ink-m') || !el.firstChild) return;
    var s = document.createElement('span');
    s.className = 'ink-m';
    while (el.firstChild) s.appendChild(el.firstChild);
    el.appendChild(s);
  }
  var queued = false;
  function tick() {
    queued = false;
    var home = document.querySelector('.v7-home');
    if (!home) return;
    home.querySelectorAll('.v7-bento .v7-t:not(.v7-open)').forEach(function (t) {
      var b = t.querySelector('b');
      // the small tiles carry only a name; their description lives in aria-label ("name. description"): show it too, so every tool has words to mark
      if (b && !t.querySelector('small')) {
        var a = t.getAttribute('aria-label') || '', i = a.indexOf('. ');
        if (i > 0) { var sm = document.createElement('small'); sm.className = 'ink-sub'; sm.textContent = a.slice(i + 2); b.after(sm); }
      }
      t.querySelectorAll('b, small').forEach(wrap);
    });
  }
  new MutationObserver(function () { if (!queued) { queued = true; requestAnimationFrame(tick); } }).observe(document.body, { childList: true, subtree: true, characterData: true });
  tick();
})();
