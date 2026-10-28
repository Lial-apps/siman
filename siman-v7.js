/* סימן — עיצוב 7: מסך "לפני ← אחרי", כחול שמיים, בחירת צבע ושפת ממשק.

   איך מחליפים בין העיצובים
   - DEFAULT_DESIGN למטה קובע מה כולם רואים: 'classic' = העיצוב הקודם, 'new' = העיצוב החדש.
   - כל אחד יכול לנסות במכשיר שלו: פותחים את האפליקציה עם ?design=new או ?design=classic בסוף הכתובת.
   - בעיצוב החדש יש בחלון הצבעים קישור "חזרה לעיצוב הקודם".
   כשהעיצוב הקודם פעיל, הקובץ הזה לא נוגע בשום דבר.

   הקובץ לא מזיז אף כפתור קיים: הוא מוסיף מסך פתיחה, שני כפתורים בסרגל וסרגל קבוצות בטלפון,
   וכל פעולה שלו לוחצת על הכפתור המקורי של האפליקציה. */
(function () {
  'use strict';
  var DEFAULT_DESIGN = 'classic';

  var LS = {
    get: function (k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
    set: function (k, v) { try { localStorage.setItem(k, v); } catch (e) {} }
  };
  try {
    var q = new URLSearchParams(location.search).get('design');
    if (q === 'new' || q === 'classic') LS.set('siman.design', q);
  } catch (e) {}
  var design = LS.get('siman.design') || DEFAULT_DESIGN;
  if (design !== 'new') return;

  var root = document.documentElement;
  root.classList.add('v7');
  var $ = function (id) { return document.getElementById(id); };
  var NS = 'http://www.w3.org/2000/svg';
  var svg = function (d, w) { return '<svg viewBox="0 0 24 24" width="' + (w || 22) + '" height="' + (w || 22) + '" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">' + d + '</svg>'; };
  var I = {
    up: '<path d="M12 15V4"/><path d="M7.5 8.5L12 4l4.5 4.5"/><path d="M4 15v3a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-3"/>',
    file: '<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/>',
    cam: '<path d="M4 8h3l2-3h6l2 3h3v11H4z"/><circle cx="12" cy="13" r="3.5"/>',
    mic: '<rect x="9" y="3" width="6" height="11" rx="3"/><path d="M5 11a7 7 0 0 0 14 0"/><path d="M12 18v3"/>',
    search: '<circle cx="11" cy="11" r="6.5"/><path d="M16 16l4 4"/>',
    lock: '<rect x="5" y="11" width="14" height="9" rx="2"/><path d="M8 11V8a4 4 0 0 1 8 0v3"/>',
    globe: '<circle cx="12" cy="12" r="9"/><path d="M3 12h18"/><path d="M12 3c2.5 2.6 3.8 5.6 3.8 9s-1.3 6.4-3.8 9c-2.5-2.6-3.8-5.6-3.8-9S9.5 5.6 12 3z"/>',
    pal: '<path d="M12 3a9 9 0 1 0 0 18c1.1 0 1.8-.8 1.8-1.8 0-.5-.2-.9-.5-1.2-.3-.3-.5-.7-.5-1.2 0-1 .8-1.8 1.8-1.8H17a4 4 0 0 0 4-4c0-4.4-4-8-9-8z"/><circle cx="7.5" cy="11" r="1.2"/><circle cx="10.5" cy="7.5" r="1.2"/><circle cx="15" cy="7.5" r="1.2"/>',
    check: '<path d="M5 12.5l4.5 4.5L19 7.5"/>',
    arrow: '<path d="M19 12H5"/><path d="M10 7l-5 5 5 5"/>',
    select: '<path d="M6 3l13 7-5.5 1.8L11.5 18z"/>',
    mark: '<path d="M9.5 14.5l7-8.5 3 3-8.5 7z"/><path d="M9.5 14.5L7 18h3.5l.5-1.5"/><path d="M4 21h16"/>',
    text: '<path d="M5 5h14"/><path d="M12 5v14"/><path d="M9 19h6"/>',
    sign: '<path d="M3 16c3-6 5-7 6-3s2 4 4 0 3-3 4 0 2 2 4 0"/><path d="M3 20.5h18"/>'
  };

  /* ================================================================ צבעים */
  var PAL = [
    { id: 'sky', bg: '#EEF3FB', surface: '#FFFFFF', ink: '#0F213F', muted: '#4B5B77', line: '#D6E0F0', accent: '#1D4ED8', on: '#FFFFFF', dash: '#8FAEEB', soft: '#DCE6FA', art: '#1D4ED8', dark: false },
    { id: 'sea', bg: '#EEF6F4', surface: '#FFFFFF', ink: '#12302C', muted: '#4E6B66', line: '#D3E6E1', accent: '#0F766E', on: '#FFFFFF', dash: '#7FB8AF', soft: '#D6F0EB', art: '#0F766E', dark: false },
    { id: 'lavender', bg: '#F5F3FA', surface: '#FFFFFF', ink: '#221E3D', muted: '#5E5878', line: '#E2DDF0', accent: '#5A48D6', on: '#FFFFFF', dash: '#B1A6EC', soft: '#E6E1FA', art: '#5A48D6', dark: false },
    { id: 'sunset', bg: '#FFF5EF', surface: '#FFFFFF', ink: '#3A1E14', muted: '#7A5446', line: '#F3DDD0', accent: '#C2410C', on: '#FFFFFF', dash: '#EBA27E', soft: '#FBE1D3', art: '#C2410C', dark: false },
    { id: 'raspberry', bg: '#FCF2F6', surface: '#FFFFFF', ink: '#3B1028', muted: '#7A4A62', line: '#F1D9E4', accent: '#BE185D', on: '#FFFFFF', dash: '#E79BBE', soft: '#F8DCE8', art: '#BE185D', dark: false },
    { id: 'forest', bg: '#F2F6EE', surface: '#FFFFFF', ink: '#1C2B16', muted: '#55654B', line: '#DCE6D2', accent: '#3F6212', on: '#FFFFFF', dash: '#A3BE84', soft: '#E1EBD5', art: '#3F6212', dark: false },
    { id: 'night', bg: '#14161C', surface: '#1E2129', ink: '#F1F2F5', muted: '#A9AFBC', line: '#2E323C', accent: '#7DD3C0', on: '#0B1F1B', dash: '#4C7F75', soft: '#253A37', art: '#2C8C79', dark: true }
  ];
  var colorId = LS.get('siman.v7.color');
  if (!PAL.some(function (p) { return p.id === colorId; })) {
    colorId = (window.matchMedia && matchMedia('(prefers-color-scheme: dark)').matches) ? 'night' : 'sky';
  }
  function applyColor(id) {
    var p = PAL.filter(function (x) { return x.id === id; })[0] || PAL[0];
    colorId = p.id;
    ['bg', 'surface', 'ink', 'muted', 'line', 'accent', 'on', 'dash', 'soft', 'art'].forEach(function (k) { root.style.setProperty('--v7-' + k, p[k]); });
    root.setAttribute('data-theme', p.dark ? 'dark' : 'light');
    var m = document.querySelector('meta[name="theme-color"]');
    if (m) m.setAttribute('content', p.surface);
    renderPops();
  }

  /* ================================================================ שפות */
  var T = {
    he: { dir: 'rtl', native: 'עברית', title: 'כל מסמך. כל מה שצריך לעשות איתו.', sub: 'לערוך, לתרגם, לתמלל, לסרוק, לעצב ולשתף. באפליקציה אחת.', search: 'מה תרצו לעשות? תרגום, תמלול, עריכת מצגת', searchLabel: 'חיפוש פעולה', drop: 'גררו קובץ לכאן', dropSub: 'PDF, Word, Excel, PowerPoint, תמונה או הקלטה. הקובץ נשאר אצלכם במכשיר.', choose: 'בחירת קובץ', openFile: 'פתיחת קובץ', scan: 'צילום דף', record: 'הקלטה', fav: 'הכי אוהבים', more: 'ועוד הרבה', newLbl: 'ליצור חדש:', newDoc: 'מסמך Word', newXl: 'גיליון Excel', newPpt: 'מצגת PowerPoint', priv: 'הקבצים נשארים אצלכם במכשיר', colorBtn: 'בחירת צבע', langBtn: 'שפת הממשק', yourColor: 'הצבע שלכם', colorNote: 'הבחירה נשמרת במכשיר, ואפשר לשנות בכל רגע.', classic: 'חזרה לעיצוב הקודם', langTitle: 'שפת הממשק', langNote: 'חלק מהחלונות עדיין מופיעים בעברית.',
      pal: ['כחול שמיים', 'ים', 'לבנדר', 'שקיעה', 'פטל', 'יער', 'לילה'], dock: ['בחירה', 'סימון', 'טקסט', 'מסמך', 'חתימה'], dockLabel: 'קבוצות כלים',
      cards: { translate: ['לתרגם מסמך', 'מכל שפה, באותו עיצוב בדיוק'], transcribe: ['לתמלל הקלטה', 'שיעור, שיחה או הודעה קולית לטקסט'], word: ['לערוך Word', 'כמו במחשב, גם בטלפון'], scan: ['מדף למסמך', 'מצלמים דף, מקבלים מסמך'], slides: ['לעצב ולתקן מצגת', 'גם מצגת שלא נפתחת'], ask: ['לשאול את המסמך', 'תשובה מיד, בלי לקרוא הכול'], excel: ['לסדר Excel', 'עריכה מלאה בכמה לחיצות'], conv: ['להמיר ל־Word', 'PDF שנפתח לעריכה רגילה'], send: ['לשלוח לחתימה', 'לאדם אחד או לכמה אנשים'], merge: ['לאחד ולפצל', 'הרבה קבצים לאחד, או ההפך'], sign: ['למלא ולחתום', 'טופס מוכן תוך שתי דקות'] } },
    ar: { dir: 'rtl', native: 'العربية', title: 'كل مستند. كل ما تحتاج فعله به.', sub: 'حرّر، ترجم، فرّغ التسجيلات، امسح ضوئيًا، صمّم وشارك. في تطبيق واحد.', search: 'ماذا تريد أن تفعل؟ ترجمة، تفريغ، تحرير عرض', searchLabel: 'البحث عن إجراء', drop: 'اسحب ملفًا إلى هنا', dropSub: 'PDF وWord وExcel وPowerPoint، صورة أو تسجيل. يبقى الملف على جهازك.', choose: 'اختيار ملف', openFile: 'فتح ملف', scan: 'تصوير صفحة', record: 'تسجيل', fav: 'الأكثر استخدامًا', more: 'والمزيد', newLbl: 'إنشاء جديد:', newDoc: 'مستند Word', newXl: 'جدول Excel', newPpt: 'عرض PowerPoint', priv: 'تبقى ملفاتك على جهازك', colorBtn: 'اختيار اللون', langBtn: 'لغة الواجهة', yourColor: 'لونك', colorNote: 'يُحفظ اختيارك على الجهاز ويمكنك تغييره في أي وقت.', classic: 'العودة إلى التصميم السابق', langTitle: 'لغة الواجهة', langNote: 'بعض النوافذ ما زالت تظهر بالعبرية.',
      pal: ['أزرق سماوي', 'بحر', 'لافندر', 'غروب', 'توت', 'غابة', 'ليل'], dock: ['اختيار', 'تمييز', 'نص', 'مستند', 'توقيع'], dockLabel: 'مجموعات الأدوات',
      cards: { translate: ['ترجمة مستند', 'من أي لغة، بنفس التصميم تمامًا'], transcribe: ['تفريغ تسجيل', 'درس أو مكالمة أو رسالة صوتية إلى نص'], word: ['تحرير Word', 'كما على الحاسوب، وعلى الهاتف أيضًا'], scan: ['من ورقة إلى مستند', 'صوّر صفحة واحصل على مستند'], slides: ['تصميم وإصلاح عرض', 'حتى عرض لا يفتح'], ask: ['اسأل المستند', 'إجابة فورية دون قراءة كل شيء'], excel: ['ترتيب Excel', 'تحرير كامل ببضع نقرات'], conv: ['تحويل إلى Word', 'PDF يُفتح للتحرير العادي'], send: ['إرسال للتوقيع', 'لشخص واحد أو لعدة أشخاص'], merge: ['دمج وتقسيم', 'عدة ملفات في ملف واحد، أو العكس'], sign: ['تعبئة وتوقيع', 'نموذج جاهز خلال دقيقتين'] } },
    ru: { dir: 'ltr', native: 'Русский', title: 'Любой документ. Всё, что с ним нужно сделать.', sub: 'Редактируйте, переводите, расшифровывайте, сканируйте, оформляйте и делитесь. В одном приложении.', search: 'Что нужно сделать? Перевод, расшифровка, презентация', searchLabel: 'Поиск действия', drop: 'Перетащите файл сюда', dropSub: 'PDF, Word, Excel, PowerPoint, фото или запись. Файл остаётся на вашем устройстве.', choose: 'Выбрать файл', openFile: 'Открыть файл', scan: 'Снять страницу', record: 'Запись', fav: 'Самое популярное', more: 'И многое другое', newLbl: 'Создать:', newDoc: 'Документ Word', newXl: 'Таблицу Excel', newPpt: 'Презентацию PowerPoint', priv: 'Файлы остаются на вашем устройстве', colorBtn: 'Выбор цвета', langBtn: 'Язык интерфейса', yourColor: 'Ваш цвет', colorNote: 'Выбор сохраняется на устройстве, его можно изменить в любой момент.', classic: 'Вернуть прежний дизайн', langTitle: 'Язык интерфейса', langNote: 'Некоторые окна пока на иврите.',
      pal: ['Небесный', 'Море', 'Лаванда', 'Закат', 'Малина', 'Лес', 'Ночь'], dock: ['Выбор', 'Пометки', 'Текст', 'Документ', 'Подпись'], dockLabel: 'Группы инструментов',
      cards: { translate: ['Перевести документ', 'С любого языка, без потери оформления'], transcribe: ['Расшифровать запись', 'Урок, звонок или голосовое — в текст'], word: ['Редактировать Word', 'Как на компьютере, даже на телефоне'], scan: ['Из листа в документ', 'Снимите лист — получите документ'], slides: ['Оформить презентацию', 'Даже ту, что не открывается'], ask: ['Спросить документ', 'Ответ сразу, без чтения целиком'], excel: ['Работать с Excel', 'Полное редактирование в пару касаний'], conv: ['Преобразовать в Word', 'PDF для обычной правки'], send: ['Отправить на подпись', 'Одному человеку или нескольким'], merge: ['Объединить и разделить', 'Много файлов в один или наоборот'], sign: ['Заполнить и подписать', 'Форма готова за две минуты'] } },
    en: { dir: 'ltr', native: 'English', title: 'Any document. Everything you need to do with it.', sub: 'Edit, translate, transcribe, scan, design and share. All in one app.', search: 'What do you want to do? Translate, transcribe, edit slides', searchLabel: 'Search actions', drop: 'Drop a file here', dropSub: 'PDF, Word, Excel, PowerPoint, a photo or a recording. Your file stays on your device.', choose: 'Choose file', openFile: 'Open file', scan: 'Scan a page', record: 'Record', fav: 'Most loved', more: 'And much more', newLbl: 'Create new:', newDoc: 'Word document', newXl: 'Excel sheet', newPpt: 'PowerPoint deck', priv: 'Your files stay on your device', colorBtn: 'Choose color', langBtn: 'Interface language', yourColor: 'Your color', colorNote: 'Your choice is saved on this device, and you can change it anytime.', classic: 'Back to the previous design', langTitle: 'Interface language', langNote: 'Some windows are still in Hebrew.',
      pal: ['Sky', 'Sea', 'Lavender', 'Sunset', 'Raspberry', 'Forest', 'Night'], dock: ['Select', 'Mark', 'Text', 'Document', 'Sign'], dockLabel: 'Tool groups',
      cards: { translate: ['Translate a document', 'From any language, same exact layout'], transcribe: ['Transcribe a recording', 'A lesson, call or voice note into text'], word: ['Edit Word files', 'Like on a computer, even on your phone'], scan: ['Page to document', 'Snap a page, get a clean document'], slides: ['Design and fix slides', 'Even a deck that won’t open'], ask: ['Ask the document', 'Instant answers, no need to read it all'], excel: ['Work in Excel', 'Full editing in a few taps'], conv: ['Convert to Word', 'A PDF opened for normal editing'], send: ['Send for signature', 'To one person or several'], merge: ['Merge and split', 'Many files into one, or the reverse'], sign: ['Fill and sign', 'A form done in two minutes'] } }
  };
  var LANGS = ['he', 'ar', 'ru', 'en'];
  var lang = LS.get('siman.v7.lang');
  if (LANGS.indexOf(lang) < 0) {
    var nl = String((navigator.language || 'he')).slice(0, 2).toLowerCase();
    lang = (nl === 'ar' || nl === 'ru') ? nl : 'he';
  }

  /* מילון לטקסטים של האפליקציה (הסרגל העליון, הכלים ושורת האפשרויות): עברית → [ערבית, רוסית, אנגלית] */
  var DICT_ROWS = [
    ['בחירה', 'اختيار', 'Выбор', 'Select'], ['טקסט', 'نص', 'Текст', 'Text'], ['עריכה', 'تحرير', 'Правка', 'Edit'], ['מרקר', 'تمييز', 'Маркер', 'Highlight'],
    ['עט', 'قلم', 'Ручка', 'Pen'], ['מחק', 'ممحاة', 'Ластик', 'Eraser'], ['השחרה', 'تعتيم', 'Скрытие', 'Redact'], ['העתקה', 'نسخ', 'Копия', 'Copy'],
    ['חתימה', 'توقيع', 'Подпись', 'Sign'], ['תמונה', 'صورة', 'Фото', 'Image'],
    ['בחירה וגלילה (V)', 'اختيار وتمرير (V)', 'Выбор и прокрутка (V)', 'Select and scroll (V)'],
    ['הוספת טקסט (T)', 'إضافة نص (T)', 'Добавить текст (T)', 'Add text (T)'],
    ['עריכת טקסט קיים בקובץ (R)', 'تحرير النص الموجود في الملف (R)', 'Править текст в файле (R)', 'Edit existing text (R)'],
    ['מרקר (H)', 'قلم تمييز (H)', 'Маркер (H)', 'Highlighter (H)'], ['עט (P)', 'قلم (P)', 'Ручка (P)', 'Pen (P)'], ['מחק (E)', 'ممحاة (E)', 'Ластик (E)', 'Eraser (E)'],
    ['השחרה: מחיקה אמיתית של מידע רגיש (D)', 'تعتيم: حذف حقيقي للمعلومات الحساسة (D)', 'Скрытие: настоящее удаление конфиденциальных данных (D)', 'Redact: truly removes sensitive information (D)'],
    ['העתקת אזור כתמונה (C)', 'نسخ منطقة كصورة (C)', 'Копировать область как изображение (C)', 'Copy an area as an image (C)'],
    ['הוספת חתימה', 'إضافة توقيع', 'Добавить подпись', 'Add signature'], ['הוספת תמונה', 'إضافة صورة', 'Добавить изображение', 'Add image'],
    ['כלים', 'أدوات', 'Инструменты', 'Tools'], ['עוזר', 'المساعد', 'Помощник', 'Assistant'], ['הקטנה', 'تصغير', 'Уменьшить', 'Zoom out'], ['הגדלה', 'تكبير', 'Увеличить', 'Zoom in'],
    ['רוחב', 'العرض', 'По ширине', 'Fit'], ['התאמה לרוחב', 'ملاءمة للعرض', 'По ширине страницы', 'Fit to width'],
    ['ביטול (Ctrl+Z)', 'تراجع (Ctrl+Z)', 'Отменить (Ctrl+Z)', 'Undo (Ctrl+Z)'], ['ביטול', 'تراجع', 'Отменить', 'Undo'],
    ['עמודים', 'الصفحات', 'Страницы', 'Pages'], ['ניהול עמודים', 'إدارة الصفحات', 'Управление страницами', 'Manage pages'],
    ['תרגום', 'ترجمة', 'Перевод', 'Translate'], ['תרגום המסמך', 'ترجمة المستند', 'Перевести документ', 'Translate the document'],
    ['פתיחה', 'فتح', 'Открыть', 'Open'], ['הדפסה', 'طباعة', 'Печать', 'Print'], ['הדפסה (Ctrl+P)', 'طباعة (Ctrl+P)', 'Печать (Ctrl+P)', 'Print (Ctrl+P)'],
    ['המרה', 'تحويل', 'Конвертировать', 'Convert'],
    ['המרה ל-PowerPoint, Word או Excel לעריכה', 'تحويل إلى PowerPoint أو Word أو Excel للتحرير', 'Преобразовать в PowerPoint, Word или Excel для правки', 'Convert to PowerPoint, Word or Excel for editing'],
    ['עריכת Word', 'تحرير Word', 'Правка Word', 'Edit as Word'], ['שמירה', 'حفظ', 'Сохранить', 'Save'], ['שמירה (Ctrl+S)', 'حفظ (Ctrl+S)', 'Сохранить (Ctrl+S)', 'Save (Ctrl+S)'],
    ['עוד אפשרויות שמירה', 'خيارات حفظ أخرى', 'Другие варианты сохранения', 'More save options'], ['שמירה בשם…', 'حفظ باسم…', 'Сохранить как…', 'Save as…'],
    ['שמירה ב-Google Drive', 'حفظ في Google Drive', 'Сохранить в Google Drive', 'Save to Google Drive'], ['שמירה ב-iCloud Drive', 'حفظ في iCloud Drive', 'Сохранить в iCloud Drive', 'Save to iCloud Drive'],
    ['מה לעשות?', 'ماذا تريد أن تفعل؟', 'Что сделать?', 'What to do?'], ['מה לעשות? (Ctrl+K)', 'ماذا تريد أن تفعل؟ (Ctrl+K)', 'Что сделать? (Ctrl+K)', 'What to do? (Ctrl+K)'],
    ['מה לעשות', 'ماذا تريد أن تفعل', 'Что сделать', 'What to do'], ['שיתוף', 'مشاركة', 'Поделиться', 'Share'],
    ['שליחה במייל, בוואטסאפ, ב-AirDrop או בשיתוף מהיר', 'إرسال بالبريد أو واتساب أو AirDrop أو المشاركة السريعة', 'Отправить по почте, в WhatsApp, через AirDrop или быструю отправку', 'Send by email, WhatsApp, AirDrop or Quick Share'],
    ['הגדרות: השם והמייל שלך', 'الإعدادات: اسمك وبريدك', 'Настройки: ваше имя и почта', 'Settings: your name and email'], ['הגדרות', 'الإعدادات', 'Настройки', 'Settings'],
    ['התקנה', 'تثبيت', 'Установить', 'Install'], ['התקנת האפליקציה במכשיר', 'تثبيت التطبيق على الجهاز', 'Установить приложение на устройство', 'Install the app on this device'],
    ['מה לעשות? למשל: לתרגם לאנגלית, לחתום, לפצל עמודים', 'ماذا تريد أن تفعل؟ مثلًا: ترجمة إلى الإنجليزية، توقيع، تقسيم الصفحات', 'Что сделать? Например: перевести на английский, подписать, разделить страницы', 'What to do? For example: translate to English, sign, split pages'],
    ['בחרו כלי כדי לסמן על המסמך', 'اختر أداة للكتابة على المستند', 'Выберите инструмент, чтобы отмечать в документе', 'Choose a tool to mark up the document'],
    ['גודל', 'الحجم', 'Размер', 'Size'], ['גודל גופן', 'حجم الخط', 'Размер шрифта', 'Font size'], ['עובי קו', 'سُمك الخط', 'Толщина линии', 'Line width'],
    ['גופן', 'الخط', 'Шрифт', 'Font'], ['מודגש', 'عريض', 'Жирный', 'Bold'], ['פסקה שלמה', 'فقرة كاملة', 'Весь абзац', 'Whole paragraph'],
    ['זיהוי טקסט (OCR)', 'التعرف على النص (OCR)', 'Распознать текст (OCR)', 'Recognize text (OCR)'], ['הדבקה', 'لصق', 'Вставить', 'Paste'], ['מחיקה', 'حذف', 'Удалить', 'Delete'],
    ['לחצו על שורה כדי לערוך אותה. Esc לסיום', 'انقر على سطر لتحريره. Esc للإنهاء', 'Нажмите на строку, чтобы изменить её. Esc — готово', 'Click a line to edit it. Esc to finish'],
    ['גררו מלבן על מידע רגיש. הטקסט שמתחת יימחק מהקובץ לגמרי ויוחלף בפס שחור', 'اسحب مستطيلًا فوق المعلومات الحساسة. سيُحذف النص تحته من الملف نهائيًا ويُستبدل بشريط أسود', 'Выделите рамкой конфиденциальные данные. Текст под ней будет полностью удалён из файла и заменён чёрной полосой', 'Drag a box over sensitive information. The text underneath is removed from the file and replaced with a black bar'],
    ['גררו מלבן סביב האזור שרוצים להעתיק. אחר כך מדביקים ב-Word, בוואטסאפ או בתוך המסמך', 'اسحب مستطيلًا حول المنطقة المراد نسخها، ثم الصقها في Word أو واتساب أو داخل المستند', 'Выделите рамкой нужную область, затем вставьте её в Word, WhatsApp или в документ', 'Drag a box around the area to copy, then paste it into Word, WhatsApp or the document'],
    ['לחצו על הדף כדי לכתוב', 'انقر على الصفحة للكتابة', 'Нажмите на страницу, чтобы писать', 'Click the page to type'],
    ['גררו מלבן על מה שרוצים לסמן', 'اسحب مستطيلًا فوق ما تريد تمييزه', 'Выделите рамкой то, что нужно отметить', 'Drag a box over what you want to highlight'],
    ['כתבו או ציירו חופשי', 'اكتب أو ارسم بحرية', 'Пишите или рисуйте от руки', 'Write or draw freely'],
    ['לוחצים על סימון כדי למחוק אותו. גוררים מלבן על הדף כדי למחוק מה שכתוב שם', 'انقر على علامة لحذفها، أو اسحب مستطيلًا لحذف ما كُتب هناك', 'Нажмите на пометку, чтобы удалить её, или выделите рамкой область для стирания', 'Tap a mark to delete it, or drag a box to erase what’s written there'],
    ['לחצו על הדף במקום שבו לשים את החתימה', 'انقر على الصفحة حيث تريد وضع التوقيع', 'Нажмите на страницу там, где поставить подпись', 'Click the page where the signature should go'],
    ['לחצו על הדף במקום שבו לשים את התמונה', 'انقر على الصفحة حيث تريد وضع الصورة', 'Нажмите на страницу там, где разместить изображение', 'Click the page where the image should go']
  ];
  var DICT = {};
  DICT_ROWS.forEach(function (r) { DICT[r[0]] = { ar: r[1], ru: r[2], en: r[3] }; });

  /* מתרגם עדין: זוכר את הטקסט העברי המקורי ומחזיר אותו כשחוזרים לעברית */
  var origText = new WeakMap(), setText = new WeakMap(), origAttr = new WeakMap();
  var ATTRS = ['title', 'aria-label', 'placeholder'];
  function trFor(he) { var d = DICT[he]; return d && lang !== 'he' ? d[lang] : null; }
  function doText(n, force) {
    var v = n.nodeValue;
    if (!force && setText.get(n) === v) return;
    var key = v.trim();
    if (!key) return;
    if (!/[\u0590-\u05FF]/.test(key) && !origText.has(n)) return;
    var he = /[\u0590-\u05FF]/.test(key) ? key : origText.get(n);
    if (/[\u0590-\u05FF]/.test(key)) origText.set(n, key);
    var out = lang === 'he' ? he : (trFor(he) || he);
    if (out === key) return;
    var nv = v.replace(key, out);
    setText.set(n, nv);
    n.nodeValue = nv;
  }
  function doAttrs(el, force) {
    var rec = origAttr.get(el) || {};
    ATTRS.forEach(function (a) {
      if (!el.hasAttribute(a)) return;
      var v = el.getAttribute(a);
      if (!force && rec[a] && rec[a].set === v) return;
      var he = /[\u0590-\u05FF]/.test(v) ? v : (rec[a] && rec[a].he);
      if (!he) return;
      var out = lang === 'he' ? he : (trFor(he) || he);
      rec[a] = { he: he, set: out };
      if (out !== v) el.setAttribute(a, out);
    });
    origAttr.set(el, rec);
  }
  function translateTree(rootEl, force) {
    if (!rootEl) return;
    doAttrs(rootEl, force);
    var w = document.createTreeWalker(rootEl, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
    var n;
    while ((n = w.nextNode())) {
      if (n.nodeType === 3) { if (n.parentNode && !/^(SCRIPT|STYLE)$/.test(n.parentNode.nodeName)) doText(n, force); }
      else doAttrs(n, force);
    }
  }
  var TR_ROOTS = function () { return [document.querySelector('header.bar'), $('opts'), $('cmdDlg')]; };
  var trBusy = false;
  var trObs = new MutationObserver(function (list) {
    if (trBusy) return;
    trBusy = true;
    try {
      list.forEach(function (m) {
        if (m.type === 'characterData') doText(m.target);
        else if (m.type === 'attributes') doAttrs(m.target);
        else m.addedNodes.forEach(function (a) { if (a.nodeType === 3) doText(a); else if (a.nodeType === 1) translateTree(a); });
      });
    } finally { trBusy = false; }
  });

  /* ================================================================ כלים קטנים */
  var h = function (html) { var t = document.createElement('template'); t.innerHTML = html.trim(); return t.content.firstElementChild; };
  var esc = function (s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  function qa(id) { var b = document.querySelector('.qa[data-qa="' + id + '"]'); if (b) { b.click(); return true; } return false; }
  function qaText(re) {
    var bs = document.querySelectorAll('.wl-quick .qa');
    for (var i = 0; i < bs.length; i++) if (re.test(bs[i].textContent)) { bs[i].click(); return true; }
    return false;
  }
  function openFile() { var b = $('openBtn2') || $('openBtn'); if (b) b.click(); }
  var ACT = {
    translate: function () { qa('tr') || openFile(); },
    transcribe: function () { qa('voice') || openFile(); },
    word: openFile, slides: openFile, excel: openFile,
    scan: function () { qaText(/מדף למסמך/) || openFile(); },
    ask: function () { qaText(/לשאול את המסמך/) || openFile(); },
    conv: function () { qa('conv') || openFile(); },
    send: function () { qa('send') || openFile(); },
    merge: function () { qa('pages') || openFile(); },
    sign: function () { qa('sign') || openFile(); },
    record: function () { qa('voice-rec') || qa('voice') || openFile(); },
    newdoc: function () { qa('newdoc'); }, newxl: function () { qa('newxl'); }, newppt: function () { qa('newppt'); }
  };

  /* ================================================================ איורי "לפני ← אחרי" */
  var LN = function (w) { return '<i class="v7-ln"' + (w ? ' style="width:' + w + '%"' : '') + '></i>'; };
  var PG = function (inner, on, extra) { return '<span class="v7-pg' + (on ? ' on' : '') + '"' + (extra ? ' ' + extra : '') + '>' + inner + '</span>'; };
  var ARROW = '<svg class="v7-arrow" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">' + I.arrow + '</svg>';
  var ART = {
    translate: [PG('<b dir="auto" style="font-size:8px;line-height:1" data-art="from">Договор</b>' + LN() + LN() + LN(60)), PG('<b dir="auto" style="font-size:10px;line-height:1" data-art="to">חוזה</b>' + LN() + LN() + LN(60), 1)],
    transcribe: ['<svg width="50" height="40" viewBox="0 0 50 40" fill="none" stroke="#0F213F" stroke-width="3" stroke-linecap="round" aria-hidden="true"><path d="M4 16v8M11 10v20M18 4v32M25 12v16M32 7v26M39 14v12M46 18v4"/></svg>', PG(LN() + LN() + LN(80) + LN() + LN(45), 1)],
    word: [PG(LN() + LN() + LN() + LN(50)), PG('<i style="display:flex;gap:2px"><i style="width:7px;height:6px;border-radius:1px;background:#CBD5EA"></i><i style="width:7px;height:6px;border-radius:1px;background:var(--v7-art)"></i><i style="width:7px;height:6px;border-radius:1px;background:#CBD5EA"></i></i><i style="height:5px;width:75%;border-radius:2px;background:#0F213F"></i>' + LN() + LN(65), 1)],
    scan: ['<span style="width:46px;height:60px;background:#F4EEDC;border-radius:4px;border:1px solid #D8CDAE;transform:rotate(-9deg);display:flex;align-items:center;justify-content:center"><svg width="34" height="40" viewBox="0 0 34 40" fill="none" stroke="#8C7F5E" stroke-width="1.6" stroke-linecap="round" aria-hidden="true"><path d="M3 8c4-3 7 3 11 0s7 3 11 0 5 2 6 1"/><path d="M3 18c4-3 7 3 11 0s7 3 11 0"/><path d="M3 28c4-3 7 3 11 0s5 2 8 0"/></svg></span>', PG('<i style="height:4px;width:70%;border-radius:2px;background:#0F213F"></i>' + LN() + LN() + LN() + LN(60), 1)],
    slides: ['<span style="width:60px;height:40px;border-radius:5px;border:1.5px dashed #8B97AD;display:flex;align-items:center;justify-content:center"><svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#8B97AD" stroke-width="2.4" stroke-linecap="round" aria-hidden="true"><path d="M6 6l12 12M18 6L6 18"/></svg></span>', '<span style="width:60px;height:40px;border-radius:5px;background:#fff;border:1.5px solid var(--v7-art);padding:5px;box-sizing:border-box;display:flex;gap:4px"><span style="flex:1;display:flex;flex-direction:column;gap:3px"><i style="height:4px;border-radius:2px;background:#0F213F"></i>' + LN() + LN(70) + '</span><i style="width:18px;border-radius:3px;background:#FFD7C2"></i></span>'],
    ask: [PG(LN() + LN() + LN() + LN() + LN(50)), '<span style="padding:8px 10px;border-radius:12px 12px 12px 3px;background:var(--v7-art);color:#fff;font-size:12px;font-weight:600;line-height:1.2;display:flex;flex-direction:column"><span style="font-size:9px;font-weight:400">?</span><span>12.11</span></span>'],
    excel: [PG(LN(80) + LN(40) + LN(90) + LN(55)), '<span style="width:50px;height:46px;background:#fff;border-radius:5px;border:1.5px solid var(--v7-art);box-sizing:border-box;display:grid;grid-template-columns:repeat(3,1fr);overflow:hidden">' + new Array(4).join('<i style="background:color-mix(in srgb,var(--v7-art) 25%,#fff);border-left:1px solid #fff"></i>') + new Array(7).join('<i style="border-left:1px solid #E2E8F3;border-top:1px solid #E2E8F3"></i>') + '</span>'],
    conv: [PG('<b style="font-size:8px;line-height:1">PDF</b>' + LN() + LN() + LN(60)), PG('<b style="font-size:8px;line-height:1;color:var(--v7-art)">W</b>' + LN() + LN() + LN(60), 1)],
    send: [PG(LN() + LN(70) + '<i style="height:10px;border-bottom:1.5px dashed #8B97AD"></i>', 0, 'style="justify-content:flex-end"'), '<span style="position:relative;width:56px;height:64px"><i style="position:absolute;top:0;left:0;width:40px;height:52px;background:#fff;border-radius:5px;border:1px solid #D9DEEA"></i><i style="position:absolute;top:5px;left:7px;width:40px;height:52px;background:#fff;border-radius:5px;border:1px solid #D9DEEA"></i><i style="position:absolute;top:10px;left:14px;width:40px;height:52px;background:#fff;border-radius:5px;border:1.5px solid var(--v7-art)"></i><i style="position:absolute;bottom:-4px;right:-8px;width:24px;height:24px;border-radius:12px;background:var(--v7-art);color:#fff;display:grid;place-items:center">' + svg(I.check, 14).replace('stroke-width="1.8"', 'stroke-width="3"') + '</i></span>'],
    merge: ['<span style="display:flex;gap:3px"><i style="width:18px;height:26px;background:#fff;border-radius:3px;border:1px solid #D9DEEA"></i><i style="width:18px;height:26px;background:#fff;border-radius:3px;border:1px solid #D9DEEA"></i><i style="width:18px;height:26px;background:#fff;border-radius:3px;border:1px solid #D9DEEA"></i></span>', PG(LN() + LN() + LN() + LN() + LN(50), 1)],
    sign: [PG(LN() + '<i style="height:12px;border-bottom:1.5px solid #8B97AD"></i>', 0, 'style="justify-content:flex-end"'), PG(LN() + '<svg width="32" height="14" viewBox="0 0 32 14" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" aria-hidden="true" style="color:var(--v7-art)"><path d="M2 10c3-8 5-8 6-2s3 5 5-1 4-3 5 1 5 1 12-3"/></svg>', 1, 'style="justify-content:flex-end;gap:2px"')]
  };
  var TINT = { translate: '#D9ECFF', ask: '#D9ECFF', transcribe: '#D6F5E8', scan: '#D6F5E8', word: '#EDE6FF', slides: '#EDE6FF', excel: '#EDE6FF', conv: '#EDE6FF', send: '#FFE1D6', sign: '#FFE1D6', merge: '#FFF1C2' };
  var FAV = ['translate', 'transcribe'];
  var REST = ['word', 'scan', 'slides', 'ask', 'excel', 'conv', 'send', 'merge', 'sign'];

  /* ================================================================ מסך הפתיחה */
  var home = null;
  function card(k) {
    var L = T[lang].cards[k];
    return '<button type="button" class="v7-card" data-v7="' + k + '"><span class="v7-art" style="background:' + TINT[k] + '"><span>' + ART[k][0] + ARROW + ART[k][1] + '</span></span><b>' + esc(L[0]) + '</b><small>' + esc(L[1]) + '</small></button>';
  }
  function buildHome() {
    var empty = $('empty');
    if (!empty) return;
    var L = T[lang];
    var html = '<section class="v7-home" dir="' + L.dir + '" lang="' + lang + '">' +
      '<h1>' + esc(L.title) + '</h1><p class="v7-sub">' + esc(L.sub) + '</p>' +
      '<label class="v7-search">' + svg(I.search, 20) + '<span class="sr-only" style="position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0)">' + esc(L.searchLabel) + '</span><input type="search" id="v7Search" autocomplete="off" placeholder="' + esc(L.search) + '"></label>' +
      '<div class="v7-drop"><span class="v7-ic">' + svg(I.up, 28) + '</span><span class="v7-dt"><b>' + esc(L.drop) + '</b><span>' + esc(L.dropSub) + '</span></span>' +
      '<span class="v7-acts"><button type="button" class="v7-b pri" data-v7="open">' + '<span class="v7-only-narrow">' + svg(I.file, 24) + '</span>' + '<span class="v7-only-narrow">' + esc(L.openFile) + '</span><span class="v7-only-wide">' + esc(L.choose) + '</span></button>' +
      '<button type="button" class="v7-b" data-v7="scan">' + svg(I.cam, 22) + '<span>' + esc(L.scan) + '</span></button>' +
      '<button type="button" class="v7-b" data-v7="record">' + svg(I.mic, 22) + '<span>' + esc(L.record) + '</span></button></span></div>' +
      '<h2 class="v7-h2">' + esc(L.fav) + '</h2><div class="v7-grid fav">' + FAV.map(card).join('') + '</div>' +
      '<h2 class="v7-h2">' + esc(L.more) + '</h2><div class="v7-grid">' + REST.map(card).join('') + '</div>' +
      '<div class="v7-new"><span>' + esc(L.newLbl) + '</span><button type="button" class="v7-b" data-v7="newdoc">' + esc(L.newDoc) + '</button><button type="button" class="v7-b" data-v7="newxl">' + esc(L.newXl) + '</button><button type="button" class="v7-b" data-v7="newppt">' + esc(L.newPpt) + '</button></div>' +
      '<p class="v7-private">' + svg(I.lock, 18) + esc(L.priv) + '</p></section>';
    var n = h(html);
    var AW = { he: ['Договор', 'חוזה'], ar: ['Договор', 'عقد'], ru: ['Contract', 'Договор'], en: ['Договор', 'Contract'] }[lang];
    n.querySelectorAll('[data-art="from"]').forEach(function (x) { x.textContent = AW[0]; });
    n.querySelectorAll('[data-art="to"]').forEach(function (x) { x.textContent = AW[1]; });
    if (home && home.parentNode) home.parentNode.replaceChild(n, home); else empty.appendChild(n);
    home = n;
    home.addEventListener('click', function (e) {
      var b = e.target.closest('[data-v7]');
      if (!b) return;
      var k = b.getAttribute('data-v7');
      if (k === 'open') openFile(); else if (ACT[k]) ACT[k]();
    });
    var s = $('v7Search');
    if (s) {
      var go = function () {
        var v = s.value;
        var cb = $('cmdBtn');
        if (!cb) return;
        cb.click();
        setTimeout(function () {
          var ci = $('cmdIn');
          if (!ci) return;
          ci.value = v; ci.dispatchEvent(new Event('input', { bubbles: true })); ci.focus();
          s.value = '';
        }, 30);
      };
      s.addEventListener('input', function () { if (s.value.trim()) go(); });
      s.addEventListener('keydown', function (e) { if (e.key === 'Enter') { e.preventDefault(); go(); } });
    }
  }

  /* ================================================================ כפתורי צבע ושפה + החלונות שלהם */
  var popColor = null, popLang = null, btnColor = null, btnLang = null;
  function addHeaderButtons() {
    var set = $('setBtn');
    var host = set ? set.parentNode : document.querySelector('header.bar .group:last-of-type');
    if (!host) return;
    btnLang = h('<button type="button" class="icon-btn v7-hbtn" id="v7LangBtn" aria-haspopup="dialog" aria-expanded="false">' + svg(I.globe, 20) + '</button>');
    btnColor = h('<button type="button" class="icon-btn v7-hbtn" id="v7ColorBtn" aria-haspopup="dialog" aria-expanded="false">' + svg(I.pal, 20) + '</button>');
    if (set) { host.insertBefore(btnLang, set); host.insertBefore(btnColor, set); } else { host.appendChild(btnLang); host.appendChild(btnColor); }
    btnLang.addEventListener('click', function (e) { e.stopPropagation(); toggle(popLang, btnLang, popColor, btnColor); });
    btnColor.addEventListener('click', function (e) { e.stopPropagation(); toggle(popColor, btnColor, popLang, btnLang); });
    popColor = h('<div class="v7-pop" role="dialog" hidden></div>');
    popLang = h('<div class="v7-pop" role="dialog" hidden></div>');
    document.body.appendChild(popColor); document.body.appendChild(popLang);
    [popColor, popLang].forEach(function (p) { p.addEventListener('click', function (e) { e.stopPropagation(); }); });
    document.addEventListener('click', closePops);
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape') closePops(); });
    window.addEventListener('resize', closePops);
  }
  function closePops() {
    [[popColor, btnColor], [popLang, btnLang]].forEach(function (x) { if (x[0] && !x[0].hidden) { x[0].hidden = true; x[1].setAttribute('aria-expanded', 'false'); } });
  }
  function toggle(pop, btn, other, otherBtn) {
    if (other) { other.hidden = true; otherBtn.setAttribute('aria-expanded', 'false'); }
    if (!pop.hidden) { pop.hidden = true; btn.setAttribute('aria-expanded', 'false'); return; }
    pop.hidden = false; btn.setAttribute('aria-expanded', 'true');
    var r = btn.getBoundingClientRect(), w = pop.offsetWidth;
    pop.style.top = Math.round(r.bottom + 8) + 'px';
    pop.style.left = Math.round(Math.min(Math.max(12, r.left + r.width / 2 - w / 2), innerWidth - w - 12)) + 'px';
    var f = pop.querySelector('[aria-pressed="true"]') || pop.querySelector('button');
    if (f) f.focus();
  }
  function renderPops() {
    if (!popColor) return;
    var L = T[lang];
    btnColor.setAttribute('aria-label', L.colorBtn); btnColor.title = L.colorBtn;
    btnLang.setAttribute('aria-label', L.langBtn); btnLang.title = L.langBtn;
    popColor.setAttribute('aria-label', L.yourColor); popLang.setAttribute('aria-label', L.langTitle);
    popColor.dir = L.dir; popLang.dir = L.dir;
    popColor.innerHTML = '<h3>' + esc(L.yourColor) + '</h3><div class="v7-sw">' + PAL.map(function (p, i) {
      return '<button type="button" data-c="' + p.id + '" aria-pressed="' + (p.id === colorId) + '"><i style="background:' + p.accent + ';color:' + p.on + '">' + svg(I.check, 20) + '</i><span>' + esc(L.pal[i]) + '</span></button>';
    }).join('') + '</div><p class="v7-note">' + esc(L.colorNote) + '</p><button type="button" class="v7-link" data-classic="1">' + esc(L.classic) + '</button>';
    popLang.innerHTML = '<h3>' + esc(L.langTitle) + '</h3><div class="v7-langs">' + LANGS.map(function (c) {
      return '<button type="button" data-l="' + c + '" lang="' + c + '" dir="' + T[c].dir + '" aria-pressed="' + (c === lang) + '"><span>' + esc(T[c].native) + '</span>' + svg(I.check, 20) + '</button>';
    }).join('') + '</div>' + (lang !== 'he' ? '<p class="v7-note">' + esc(L.langNote) + '</p>' : '');
    popColor.onclick = function (e) {
      e.stopPropagation();
      var c = e.target.closest('[data-c]');
      if (c) { LS.set('siman.v7.color', c.getAttribute('data-c')); applyColor(c.getAttribute('data-c')); popColor.hidden = false; return; }
      if (e.target.closest('[data-classic]')) { LS.set('siman.design', 'classic'); location.reload(); }
    };
    popLang.onclick = function (e) {
      e.stopPropagation();
      var b = e.target.closest('[data-l]');
      if (b) { LS.set('siman.v7.lang', b.getAttribute('data-l')); applyLang(b.getAttribute('data-l')); closePops(); }
    };
  }

  /* ================================================================ טלפון: סרגל קבוצות */
  var dock = null;
  var GRP_OF = { select: 'select', highlight: 'mark', pen: 'mark', eraser: 'mark', text: 'text', edit: 'text', redact: 'doc', snip: 'doc' };
  var DOCK = [['select', I.select], ['mark', I.mark], ['text', I.text], ['doc', I.file], ['sign', I.sign]];
  function buildDock() {
    var L = T[lang];
    var n = h('<nav class="v7-dock" aria-label="' + esc(L.dockLabel) + '" dir="' + L.dir + '">' + DOCK.map(function (g, i) {
      return '<button type="button" data-g="' + g[0] + '" aria-pressed="false">' + svg(g[1], 22) + '<span>' + esc(L.dock[i]) + '</span></button>';
    }).join('') + '</nav>');
    if (dock && dock.parentNode) dock.parentNode.replaceChild(n, dock); else document.body.appendChild(n);
    dock = n;
    dock.addEventListener('click', function (e) {
      var b = e.target.closest('[data-g]');
      if (!b) return;
      var g = b.getAttribute('data-g');
      var cur = root.getAttribute('data-v7grp');
      if (g === 'select') { var s = document.querySelector('.tool[data-tool="select"]'); if (s) s.click(); setGroup('select'); return; }
      if (g === 'sign') { var sg = document.querySelector('.tool[data-action="sign"]'); setGroup('sign'); if (sg) sg.click(); return; }
      setGroup(cur === g ? 'select' : g);
    });
    setGroup(root.getAttribute('data-v7grp') || GRP_OF[document.body.dataset.tool] || 'select');
  }
  function setGroup(g) {
    root.setAttribute('data-v7grp', g);
    if (dock) dock.querySelectorAll('[data-g]').forEach(function (b) { b.setAttribute('aria-pressed', String(b.getAttribute('data-g') === g)); });
  }

  /* ================================================================ שפה: החלה */
  function applyLang(code) {
    lang = LANGS.indexOf(code) >= 0 ? code : 'he';
    var L = T[lang];
    root.setAttribute('lang', lang);
    [document.querySelector('header.bar'), $('opts'), $('cmdDlg')].forEach(function (el) { if (el) el.dir = L.dir; });
    buildHome(); buildDock(); renderPops();
    trBusy = true;
    try { TR_ROOTS().forEach(function (el) { translateTree(el, true); }); } finally { trBusy = false; }
  }

  /* ================================================================ הפעלה */
  function start() {
    if (!document.querySelector('header.bar') || !$('empty')) { root.classList.remove('v7'); return; }
    addHeaderButtons();
    applyColor(colorId);
    applyLang(lang);
    // האם פתוח מסמך? (מסך הפתיחה מוסתר)
    var empty = $('empty');
    var syncDoc = function () { root.classList.toggle('v7-doc', empty.classList.contains('hidden') || getComputedStyle(empty).display === 'none'); };
    syncDoc();
    new MutationObserver(syncDoc).observe(empty, { attributes: true, attributeFilter: ['class', 'style', 'hidden'] });
    // הקבוצה בטלפון עוקבת אחרי הכלי שנבחר
    new MutationObserver(function () {
      var g = GRP_OF[document.body.dataset.tool];
      if (g && g !== 'select') setGroup(g);
    }).observe(document.body, { attributes: true, attributeFilter: ['data-tool'] });
    TR_ROOTS().forEach(function (el) {
      if (el) trObs.observe(el, { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: ATTRS });
    });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
})();
