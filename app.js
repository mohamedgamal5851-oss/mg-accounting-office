/* Accounting office app: UI, db sync, forms, reports, PDF/Excel. Logic lives in core.js (OFCore). */
(function () {
  'use strict';
  const C = window.OFCore;

  /* ---------------- labels ---------------- */
  const KIND = { fee: 'مطالبة أتعاب', expense: 'مصروفات نيابة عن العميل', payment: 'تحصيل', discount: 'خصم / تسوية' };
  const KIND_PILL = { fee: 'acc', expense: 'warn', payment: 'ok', discount: 'mute' };
  const FREQ = { monthly: 'شهري', quarterly: 'ربع سنوي', semiannual: 'نصف سنوي', annual: 'سنوي', once: 'مرة واحدة' };
  const FREQ_PER_YEAR = { monthly: 12, quarterly: 4, semiannual: 2, annual: 1, once: 0 };
  const OB_FREQ = { monthly: 'شهري', quarterly: 'ربع سنوي', annual: 'سنوي' };
  const TSTATUS = { pending: 'لم يبدأ', progress: 'جارٍ العمل', waiting: 'بانتظار مستندات العميل', done: 'تم' };
  const TSTATUS_PILL = { pending: 'mute', progress: 'info', waiting: 'warn', done: 'ok' };
  const METHODS = ['نقدًا', 'تحويل بنكي', 'شيك', 'محفظة إلكترونية / إنستاباي'];
  const LEGAL = ['شركة مساهمة', 'شركة ذات مسؤولية محدودة', 'شركة الشخص الواحد', 'شركة توصية بسيطة', 'شركة تضامن', 'منشأة فردية', 'شخص طبيعي (مهنة حرة)', 'فرع شركة أجنبية', 'جمعية / مؤسسة أهلية', 'أخرى'];
  const DOC_NAMES = ['السجل التجاري', 'البطاقة الضريبية', 'شهادة التسجيل في القيمة المضافة', 'ملف التأمينات الاجتماعية', 'رخصة مزاولة النشاط', 'عقد التأسيس', 'توكيل رسمي', 'شهادة الهيئة العامة للاستثمار', 'التوقيع الإلكتروني (Token)'];
  const TAX_LAW = { '91-2005': 'قانون 91 لسنة 2005', '6-2025': 'قانون 6 لسنة 2025 (التيسيرات)' };
  const lawOf = (c) => (c && c.taxLaw === '6-2025' ? '6-2025' : '91-2005');
  const freqFor = (o, c) => C.effFreq(o, c);
  const MONTHS = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'];

  const DEFAULT_SETTINGS = {
    officeName: 'مكتب المحاسبة', officeSub: 'محاسبون ومراجعون قانونيون', address: '', phone: '', email: '', taxNo: '', footer: '',
    currency: 'ج.م', alertDays: 7, docAlertDays: 30,
    services: ['مراجعة القوائم المالية', 'مسك الدفاتر', 'الإقرارات الضريبية', 'التأمينات الاجتماعية', 'تأسيس الشركات', 'استشارات ضريبية ومالية', 'الفحص الضريبي'],
    obligations: [
      { id: 'vat', name: 'إقرار القيمة المضافة', freq: 'monthly', freqLaw6: 'quarterly', dueOffset: 1, dueDay: 0, active: true },
      { id: 'payroll', name: 'ضريبة كسب العمل (المرتبات)', freq: 'monthly', dueOffset: 1, dueDay: 15, active: true },
      { id: 'insurance', name: 'التأمينات الاجتماعية الشهرية', freq: 'monthly', dueOffset: 1, dueDay: 15, active: true },
      { id: 'wht', name: 'الخصم والإضافة (نموذج 41)', freq: 'quarterly', dueOffset: 1, dueDay: 0, active: true },
      { id: 'corp', name: 'إقرار ضريبة الدخل السنوي – أشخاص اعتبارية', freq: 'annual', law: '91-2005', entity: 'legal', dueOffset: 4, dueDay: 0, active: true },
      { id: 'indiv', name: 'إقرار ضريبة الدخل السنوي – أشخاص طبيعيون', freq: 'annual', law: '91-2005', entity: 'natural', dueOffset: 3, dueDay: 0, active: true },
      { id: 'simple25', name: 'إقرار الدخل المبسط (نموذج 25) – قانون 6 لسنة 2025', freq: 'annual', law: '6-2025', entity: 'all', dueOffset: 3, dueDay: 0, active: true },
      { id: 'audit', name: 'القوائم المالية وتقرير مراقب الحسابات', freq: 'annual', dueOffset: 3, dueDay: 0, active: true }
    ]
  };

  /* ---------------- state ---------------- */
  const S = {
    raw: { clients: [], staff: [], contracts: [], transactions: [], tasks: [], certificates: [], vatReturns: [], vatParties: [], settings: JSON.parse(JSON.stringify(DEFAULT_SETTINGS)) },
    settingsSaved: false, loaded: {}, M: null, db: null, user: null, downloads: null, canWrite: null, isOwner: false, dbMissing: false, slow: false,
    ui: { cQ: '', cStatus: 'active', cStaff: '', cBal: '', cLaw: '', ceQ: '', fKind: '', fClient: '', fFrom: '', fTo: '', fQ: '', tClient: '', tStaff: '', tOb: '', tQ: '',
      kQ: '', kActive: 'active', stFrom: '', stTo: '', rClient: '', rFrom: '', rTo: '', rNonZero: '1', rDocs: 'alert', rStatus: '' }
  };
  const COLS = ['clients', 'staff', 'contracts', 'transactions', 'tasks', 'certificates', 'vatReturns', 'vatParties'];

  const $ = (s, r) => (r || document).querySelector(s);
  const $$ = (s, r) => Array.from((r || document).querySelectorAll(s));
  const esc = (v) => String(v == null ? '' : v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
  const uid = () => (crypto.randomUUID ? crypto.randomUUID() : 'id-' + Date.now().toString(36) + Math.random().toString(36).slice(2));
  const money = (n) => Number(n || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const cur = () => (S.raw.settings.currency || 'ج.م');
  const moneyC = (n) => money(n) + ' ' + cur();
  const int = (n) => Number(n || 0).toLocaleString('en-US');
  const monthLabel = (ym) => { if (!ym) return ''; const [y, m] = ym.split('-').map(Number); return MONTHS[m - 1] + ' ' + y; };
  const periodLabel = (t) => t.period ? monthLabel(t.period) : '';
  const clean = (o) => JSON.parse(JSON.stringify(o));
  const clientName = (id) => { const c = S.M && S.M.clientById.get(id); return c ? c.name : '—'; };
  const staffName = (id) => { const s = S.M && S.M.staffById.get(id); return s ? s.name : ''; };
  const settings = () => S.raw.settings;

  function rebuild() { S.M = C.build(S.raw); }
  let pending = false;
  function scheduleRender() {
    if (pending) return; pending = true;
    requestAnimationFrame(() => { pending = false; rebuild(); renderNav(); renderPage(); });
  }
  const allLoaded = () => COLS.every(c => S.loaded[c]) && S.loaded.settings;
  const writable = () => !!S.db && S.canWrite !== false;

  function options(list, sel, allLabel) {
    return (allLabel != null ? '<option value="">' + esc(allLabel) + '</option>' : '') + list.map(v => {
      const [val, lab] = Array.isArray(v) ? v : [v, v];
      return '<option value="' + esc(val) + '"' + (String(val) === String(sel) ? ' selected' : '') + '>' + esc(lab) + '</option>';
    }).join('');
  }
  const clientOptions = (includeInactive) => S.M.clients.filter(c => includeInactive || c.status !== 'inactive').map(c => [c.id, (c.code ? c.code + ' · ' : '') + c.name]);
  const staffOptions = () => S.M.staff.filter(s => s.active !== false).map(s => [s.id, s.name]);
  const amountCell = (n) => '<td class="n">' + (n ? money(n) : '') + '</td>';
  const balCell = (n) => '<td class="n ' + (n > 0.004 ? 'pos' : n < -0.004 ? 'neg' : '') + '">' + money(n) + '</td>';

  /* ---------------- shell ---------------- */
  const ICONS = {
    dashboard: '<path d="M4 4h7v7H4zM13 4h7v4h-7zM13 10h7v10h-7zM4 13h7v7H4z"/>',
    clients: '<circle cx="9" cy="8" r="3.2"/><path d="M3 20c.6-3.4 3-5.4 6-5.4s5.4 2 6 5.4M16 4.6a3.2 3.2 0 0 1 0 6.3M18 14.8c1.7.7 2.8 2.5 3 5.2"/>',
    finance: '<rect x="3" y="6" width="18" height="12" rx="2"/><circle cx="12" cy="12" r="2.6"/><path d="M6 9v6M18 9v6"/>',
    tasks: '<rect x="4" y="4" width="16" height="17" rx="2"/><path d="M8 2v4M16 2v4M4 10h16M8 14l2 2 4-4"/>',
    contracts: '<path d="M6 3h9l4 4v14H6z"/><path d="M9 12h7M9 16h5M14 3v5h5"/>',
    vat: '<path d="M7 3h10l3 3v15H4V3h3"/><path d="M8 8h8M8 12h3M13 12h3M8 16h3M13 16h3"/>',
    certs: '<path d="M5 3h14v18l-3-2-2 2-2-2-2 2-2-2-3 2z"/><path d="M9 8h6M9 12h6M9 16h3"/>',
    reports: '<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>',
    backup: '<path d="M12 3v12M7 10l5 5 5-5"/><path d="M4 17v3h16v-3"/>',
    'vat-list': '<path d="M8 6h12M8 12h12M8 18h12"/><circle cx="4" cy="6" r="1"/><circle cx="4" cy="12" r="1"/><circle cx="4" cy="18" r="1"/>',
    'vat-rep': '<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>',
    'vat-parties': '<circle cx="9" cy="8" r="3.2"/><path d="M3 20c.6-3.4 3-5.4 6-5.4s5.4 2 6 5.4M15 11h6M15 15h6"/>',
    f41: '<path d="M6 3h9l4 4v14H6z"/><path d="M9 13h6M9 17h4"/>',
    apps: '<rect x="4" y="4" width="6" height="6" rx="1.2"/><rect x="14" y="4" width="6" height="6" rx="1.2"/><rect x="4" y="14" width="6" height="6" rx="1.2"/><rect x="14" y="14" width="6" height="6" rx="1.2"/>',
    setup: '<circle cx="12" cy="12" r="3"/><path d="M12 2v3M12 19v3M4.2 4.2l2.1 2.1M17.7 17.7l2.1 2.1M2 12h3M19 12h3M4.2 19.8l2.1-2.1M17.7 6.3l2.1-2.1"/>'
  };
  /* the program is split into modules (like Odoo apps): each has its own menu; the clients file is shared */
  const MODULES = [
    { id: 'office', name: 'حسابات وإجراءات الشركات', desc: 'العملاء والأتعاب والتحصيل، الإجراءات والمواعيد، العقود والشهادات والتقارير', icon: 'clients', home: 'dashboard',
      nav: [['dashboard', 'لوحة التحكم'], ['clients', 'العملاء'], ['finance', 'الحركات المالية'], ['tasks', 'الإجراءات والمواعيد'], ['contracts', 'العقود والأتعاب'], ['certs', 'الشهادات'], ['reports', 'التقارير'], ['backup', 'النسخ الاحتياطي'], ['setup', 'الإعدادات']] },
    { id: 'tax', name: 'الإقرارات الضريبية', desc: 'إقرارات القيمة المضافة: تحليل فواتير البيع والشراء، ملفات الرفع على المنظومة، ومطابقة الإقرار بعد تقديمه', icon: 'vat', home: 'vat',
      nav: [['vat', 'لوحة الإقرارات'], ['vat-list', 'إقرارات القيمة المضافة'], ['clients', 'بيانات العملاء'], ['backup', 'النسخ الاحتياطي'], ['vat-rep', 'تقارير الإقرارات'], ['vat-parties', 'الموردون والعملاء'], ['f41', 'نموذج 41 (قريبًا)']] }
  ];
  const NAV = MODULES[0].nav;
  const SHARED = ['clients', 'client', 'backup'];
  let lastMod = (() => { try { return sessionStorage.getItem('of-mod') || 'office'; } catch (e) { return 'office'; } })();
  const moduleOf = (page) => page === 'home' ? null : page === 'vat' ? MODULES[1] : SHARED.includes(page) ? (MODULES.find(m => m.id === lastMod) || MODULES[0]) : MODULES[0];
  function noteModule(page) {
    if (page === 'home' || SHARED.includes(page)) return;
    lastMod = moduleOf(page).id; try { sessionStorage.setItem('of-mod', lastMod); } catch (e) {}
  }
  const svgIcon = (k) => '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">' + ICONS[k] + '</svg>';

  let pinned = true;
  try { pinned = localStorage.getItem('of-pinned') !== '0'; } catch (e) {}
  function renderShell() {
    const L = window.OF_LOGO;
    if (L) {
      const img = $('#logo-img'); img.src = L.light; img.style.display = 'block';
      if (!document.querySelector('link[rel=icon]')) { const l = document.createElement('link'); l.rel = 'icon'; l.href = L.icon; document.head.appendChild(l); }
    }
    // wide screens: the menu is a bar across the top; narrow screens: it opens from the «القائمة» button
    $('#app').classList.remove('floating'); $('#side').classList.remove('open');
    $('#side-foot').innerHTML = '<button type="button" class="pin-btn" id="pin-btn">' + (pinned ? '⇥ إخفاء القائمة' : '📌 تثبيت القائمة') + '</button><span id="as-of"></span>';
    $('#pin-btn').onclick = () => { pinned = !pinned; try { localStorage.setItem('of-pinned', pinned ? '1' : '0'); } catch (e) {} renderShell(); };
    $('#side-scrim').onclick = () => $('#side').classList.remove('open');
    // the hidden menu opens from the screen edge only when the pointer rests there (0.6 s), not when it passes by
    let hoverOpened = false, hoverTimer = 0;
    const edge = $('#side-edge');
    edge.onmouseenter = () => { if (pinned || true) return; clearTimeout(hoverTimer); hoverTimer = setTimeout(() => { hoverOpened = true; $('#side').classList.add('open'); }, 600); };
    edge.onmouseleave = () => clearTimeout(hoverTimer);
    $('#side').onmouseleave = () => { if (!pinned && hoverOpened) { hoverOpened = false; $('#side').classList.remove('open'); } };
  }
  /* top bar: a highlight slides from the previous page's tab to the new one, and the page enters from that side */
  let navIdx = -1;
  function placeNavIndicator() {
    const side = $('#side'), nav = $('#nav');
    let ind = $('#nav-ind');
    if (!ind) { ind = document.createElement('span'); ind.id = 'nav-ind'; side.appendChild(ind); }
    const links = Array.from(nav.querySelectorAll('a')), on = nav.querySelector('a.on'), i = links.indexOf(on);
    const m = $('#main'); m.style.setProperty('--nav-dir', navIdx < 0 || i < 0 || i === navIdx ? 0 : (i > navIdx ? -1 : 1));
    navIdx = i;
    if (!on) { ind.style.opacity = '0'; return; }
    const place = () => {
      const sr = side.getBoundingClientRect(), r = on.getBoundingClientRect();
      ind.style.width = r.width + 'px'; ind.style.height = r.height + 'px';
      ind.style.transform = 'translate(' + (r.left - sr.left) + 'px,' + (r.top - sr.top) + 'px)';
      ind.style.opacity = '1';
    };
    // keep the active tab in view when the bar scrolls
    if (on.offsetLeft < nav.scrollLeft || on.offsetLeft + on.offsetWidth > nav.scrollLeft + nav.clientWidth) on.scrollIntoView({ block: 'nearest', inline: 'center' });
    place();
    if (!ind.classList.contains('ready')) requestAnimationFrame(() => ind.classList.add('ready'));
    nav.onscroll = () => { ind.classList.remove('ready'); place(); requestAnimationFrame(() => ind.classList.add('ready')); };
  }
  window.addEventListener('resize', () => { const ind = $('#nav-ind'); if (ind) { ind.classList.remove('ready'); placeNavIndicator(); } });
  function renderNav() {
    const cp = currentPage(), page = cp.page; noteModule(page); const mod = moduleOf(page);
    const A = S.M && S.M.alerts;
    const badge = { tasks: A ? A.overdue.length : 0 };
    const on = (k) => page === 'vat' ? (k === 'vat-rep' ? /^rep/.test(cp.sub) : k === 'vat-parties' ? /^parties/.test(cp.sub) : k === 'vat-list' ? /^(list|ret)/.test(cp.sub) : k === 'vat' && !cp.sub) : (page === k || (page === 'client' && k === 'clients'));
    const items = mod ? mod.nav : MODULES.map(m => [m.home, m.name, m.icon]);
    $('#nav').innerHTML = '<a href="#home" class="apps' + (page === 'home' ? ' on' : '') + '">' + svgIcon('apps') + '<span>' + (mod ? 'جميع الوحدات' : 'الوحدات') + '</span></a>' +
      (mod ? '<div class="nav-mod">' + esc(mod.name) + '</div>' : '') +
      items.map(([k, label, ic]) => k === 'f41' ? '<a class="soon" aria-disabled="true">' + svgIcon(k) + '<span>' + esc(label) + '</span></a>' :
        '<a href="#' + k + '" class="' + (mod && on(k) ? 'on' : '') + '">' + svgIcon(ic || k) + '<span>' + esc(label) + '</span>' +
        (badge[k] ? '<span class="badge" title="إجراءات متأخرة">' + badge[k] + '</span>' : '') + '</a>').join('');
    placeNavIndicator();
    const build = (document.querySelector('meta[name=of-build]') || {}).content || '';
    $('#as-of').textContent = 'اليوم: ' + C.todayISO() + (build ? ' · الإصدار ' + build : '');
    $('.logo span').textContent = mod ? mod.name : 'إدارة المكتب';
    $('.logo').title = 'اليوم: ' + C.todayISO() + (build ? ' · الإصدار ' + build : '');
    $('#logo-name').textContent = settings().officeName || 'مكتب المحاسبة';
    document.title = (settings().officeName || 'مكتب المحاسبة') + ' – إدارة المكتب';
  }

  const PAGES = ['home', 'backup', 'dashboard', 'clients', 'client', 'finance', 'tasks', 'vat', 'contracts', 'certs', 'reports', 'setup'];
  function currentPage() {
    // hashes: #page, #page-sub, or #client-<tab>-<client id> (the id may itself contain hyphens)
    const parts = decodeURIComponent((location.hash || '#home').slice(1)).split('-');
    const page = PAGES.includes(parts[0]) ? parts[0] : 'home';
    if (page === 'client') return { page, sub: parts.slice(2).join('-'), extra: parts[1] || 'statement' };
    return { page, sub: parts.slice(1).join('-'), extra: '' };
  }
  function route() {
    $('#side').classList.remove('open'); rebuild(); renderNav(); renderPage(); $('#main').scrollTop = 0;
    const m = $('#main'); m.classList.remove('page-leave'); m.classList.remove('page-enter'); void m.offsetWidth; m.classList.add('page-enter');
  }
  /* leaving a page: the work area recedes, then the next page rises in (skipped when the system asks for reduced motion) */
  function leaveTo(target, card) {
    const reduce = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    if (reduce) return go(target);
    if (card) card.classList.add('mod-picked');
    $('#main').classList.add('page-leave');
    setTimeout(() => go(target), card ? 320 : 230);
  }
  const go = (h) => { if (location.hash === '#' + h) route(); else location.hash = h; };

  function topBar(title, sub, actions) {
    const home = currentPage().page === 'home';
    return '<div class="top"><button type="button" class="btn menu-btn" data-act="menu">☰ القائمة</button><div>' +
      (home ? '<h1>' + esc(title) + '</h1>' : '<h1><button type="button" class="to-mods" data-act="to-mods" title="العودة إلى الوحدات"><span class="to-mods-ic">' + svgIcon('apps') + '</span>' + esc(title) + '</button></h1>') +
      (sub ? '<div class="sub">' + sub + '</div>' : '') + '</div><div class="grow"></div>' + (actions || '') + '</div>' +
      (S.db && S.canWrite === false ? '<div class="ro-banner">يمكنك عرض البيانات فقط، والتعديل متاح لأصحاب صلاحية التعديل.</div>' : '');
  }
  const actBtn = (act, label, cls, extra) => writable() ? '<button type="button" class="btn ' + (cls || '') + '" data-act="' + act + '"' + (extra || '') + '>' + esc(label) + '</button>' : '';
  const expBtns = (key) => S.downloads ? '<button type="button" class="btn" data-act="pdf" data-key="' + key + '">⤓ PDF</button><button type="button" class="btn" data-act="xlsx" data-key="' + key + '">⤓ Excel</button>' : '';

  let rendering = false;
  function renderPage() {
    // a focused input that loses focus while the page is replaced fires «change»; its re-render waits for this one
    if (rendering) { scheduleRender(); return; }
    rendering = true;
    try { renderPageNow(); } finally { rendering = false; }
  }
  function renderPageNow() {
    const { page, sub, extra } = currentPage();
    const main = $('#main');
    if (S.dbMissing) { main.innerHTML = topBar('إدارة المكتب') + '<div class="content"><div class="card"><div class="hero-empty"><h2>تخزين البيانات غير متاح في هذا العرض</h2><p>افتح الصفحة على claude.ai بعد تسجيل الدخول.</p></div></div></div>'; return; }
    if (!allLoaded()) { main.innerHTML = '<div class="loading">' + (S.slow ? 'يستغرق تحميل البيانات وقتًا أطول من المعتاد، وستظهر فور استجابة الاتصال.' : 'جارٍ تحميل البيانات…') + '</div>'; return; }
    const focus = captureFocus();
    let h = '';
    if (page === 'home') h = pageHome();
    else if (page === 'backup') h = pageBackup();
    else if (page === 'dashboard') h = pageDashboard();
    else if (page === 'clients') h = pageClients();
    else if (page === 'client') h = pageClient(sub, extra || 'statement');
    else if (page === 'finance') h = pageFinance();
    else if (page === 'tasks') h = pageTasks(sub || 'open');
    else if (page === 'vat') h = window.OFVat ? window.OFVat.page(sub) : '';
    else if (page === 'contracts') h = pageContracts();
    else if (page === 'certs') h = pageCerts();
    else if (page === 'reports') h = pageReports(sub || 'balances');
    else if (page === 'setup') h = pageSettings();
    main.innerHTML = h;
    restoreFocus(focus);
    if (page === 'setup') { bindSettings(); bindCertSettings(); }
    if (page === 'client' && extra === 'access') bindAccess();
    if (page === 'vat' && window.OFVat) window.OFVat.bind(sub);
    if (page === 'backup') { const ib = $('#imp-backup2'); if (ib) ib.addEventListener('change', () => { if (ib.files[0]) importBackup(ib.files[0]); ib.value = ''; }); }
  }
  function captureFocus() {
    const a = document.activeElement;
    if (!a || !a.dataset || !a.dataset.ui) return null;
    return { key: a.dataset.ui, s: a.selectionStart, e: a.selectionEnd };
  }
  function restoreFocus(f) {
    if (!f) return;
    const el = document.querySelector('[data-ui="' + f.key + '"]'); if (!el) return;
    el.focus(); try { if (f.s != null) el.setSelectionRange(f.s, f.e); } catch (e) {}
  }
  // filters: any element with data-ui writes to S.ui and re-renders
  document.addEventListener('input', (ev) => { const el = ev.target; if (el.dataset && el.dataset.ui && el.dataset.live != null) { S.ui[el.dataset.ui] = el.value; renderPage(); } });
  document.addEventListener('change', (ev) => {
    const el = ev.target;
    if (el.dataset && el.dataset.ui && el.dataset.live == null) { S.ui[el.dataset.ui] = el.type === 'checkbox' ? (el.checked ? '1' : '') : el.value; renderPage(); return; }
    if (el.dataset && el.dataset.taskStatus) return setTaskStatus(el.dataset.taskStatus, el.value);
  });

  /* ---------------- dashboard ---------------- */
  function lastBackup() { try { return localStorage.getItem('of-last-backup') || ''; } catch (e) { return ''; } }
  function pageBackup() {
    const local = !!window.OF_LOCAL, last = lastBackup();
    let h = topBar('النسخ الاحتياطي والاسترجاع', local ? 'هذه النسخة تعمل على هذا الجهاز' : 'هذه النسخة منشورة على claude.ai') + '<div class="content">';
    h += '<section class="card"><div class="body">' +
      '<p style="margin:0 0 6px">النسخة الاحتياطية ملف واحد بصيغة JSON يضم جميع البيانات: العملاء، والحركات المالية، والإجراءات، والعقود، والشهادات، وإقرارات القيمة المضافة بفواتيرها وملفات الإقرارات المقدَّمة، وبيانات الدخول مشفرة بكلمة المرور الرئيسية.</p>' +
      '<p style="margin:0;color:var(--muted)">لنقل البيانات بين النسخة المنشورة على claude.ai والنسخة المثبتة على جهاز Mac: نزّل نسخة احتياطية من إحداهما، ثم استرجعها في الأخرى. السجلات التي تحمل المعرّف نفسه تُستبدل، وتبقى بقية البيانات كما هي.</p>' +
      '<div class="toolbar" style="margin-top:14px">' + (S.downloads ? '<button type="button" class="btn primary" data-act="backup">⤓ تنزيل نسخة احتياطية كاملة</button>' : '<span class="hint">التنزيل غير متاح في هذا العرض.</span>') +
      (S.isOwner ? '<label class="btn" for="imp-backup2">⤒ استرجاع نسخة احتياطية</label><input type="file" id="imp-backup2" accept=".json,application/json" hidden>' : '') + '</div>' +
      '<p class="hint" style="margin:10px 0 0">' + (last ? 'آخر نسخة احتياطية نُزّلت من هذا المتصفح: ' + esc(last.replace('T', ' ').slice(0, 16)) : 'لم تُنزَّل نسخة احتياطية من هذا المتصفح بعد.') + '</p>' +
      (local ? '<div class="note" style="margin-top:12px">بيانات هذه النسخة محفوظة داخل متصفح هذا الجهاز فقط؛ مسح بيانات المتصفح يمسحها. احرص على تنزيل نسخة احتياطية بصفة دورية.</div>' : '') +
      '</div></section></div>';
    return h;
  }
  function pageHome() {
    const M = S.M, A = M.alerts;
    const vat = window.OFVat ? window.OFVat.homeStats() : null;
    const stat = { office: int(M.active.length) + ' عميل نشط' + (A.overdue.length ? ' · ' + int(A.overdue.length) + ' إجراء متأخر' : ''),
      tax: vat ? int(vat.open) + ' إقرار تحت الإعداد' + (vat.diff ? ' · ' + int(vat.diff) + ' به فروق' : '') : '' };
    let h = topBar(settings().officeName || 'مكتب المحاسبة', 'اختر الوحدة') + '<div class="content"><div class="mods">';
    for (const [i, m] of MODULES.entries()) h += '<a class="mod" style="--i:' + i + '" href="#' + m.home + '" data-act="open-mod" data-to="' + m.home + '"><span class="mod-ic">' + svgIcon(m.icon) + '</span><b>' + esc(m.name) + '</b><small>' + esc(m.desc) + '</small><span class="mod-stat">' + esc(stat[m.id] || '') + '</span></a>';
    return h + '</div></div>';
  }
  function pageDashboard() {
    const M = S.M, A = M.alerts, T = M.totals;
    const actions = actBtn('add-tx', '+ تسجيل تحصيل', 'primary', ' data-kind="payment"') + actBtn('add-client', '+ عميل جديد');
    let h = topBar('لوحة التحكم', esc(settings().officeName) + ' · ' + esc(monthLabel(M.ym)), actions) + '<div class="content">';
    if (!M.clients.length) {
      h += '<div class="card"><div class="hero-empty"><h2>ابدأ بإضافة العملاء</h2><p>أضف عملاء المكتب واحدًا واحدًا، أو استوردهم مرة واحدة من ملف Excel من صفحة الإعدادات. بعدها سجّل عقود الأتعاب والالتزامات الدورية لكل عميل، وسيتولى البرنامج إصدار المطالبات وتوليد مواعيد الإقرارات.</p><div class="toolbar">' +
        actBtn('add-client', '+ إضافة عميل', 'primary') + '<a class="btn" href="#setup">استيراد من Excel</a></div></div></div></div>';
      return h;
    }
    h += '<div class="tiles">' +
      '<a class="tile c1" href="#clients"><span>العملاء النشطون</span><b>' + int(M.active.length) + '</b><small>من إجمالي ' + int(M.clients.length) + ' عميل</small></a>' +
      '<a class="tile c2" href="#reports-balances"><span>إجمالي المستحق على العملاء</span><b>' + money(T.due) + '</b><small>' + esc(cur()) + (T.advance ? ' · دفعات مقدمة ' + money(T.advance) : '') + '</small></a>' +
      '<a class="tile c4" href="#finance"><span>مطالبات الأتعاب هذا الشهر</span><b>' + money(T.billedMonth) + '</b><small>' + esc(cur()) + '</small></a>' +
      '<a class="tile c3" href="#reports-collections"><span>التحصيلات هذا الشهر</span><b>' + money(T.collectedMonth) + '</b><small>' + esc(cur()) + '</small></a>' +
      '<a class="tile c5" href="#tasks-overdue"><span>إجراءات متأخرة</span><b>' + int(A.overdue.length) + '</b><small>' + int(A.soon.length) + ' إجراء يستحق خلال ' + int(settings().alertDays || 7) + ' أيام</small></a>' +
      '</div>';
    const al = [];
    if (A.overdue.length) al.push(['crit', A.overdue.length, 'إجراءات تجاوزت موعدها', 'إقرارات ومهام لم تُنجز حتى الآن', '#tasks-overdue']);
    if (A.docsExpired.length) al.push(['crit', A.docsExpired.length, 'مستندات منتهية', 'سجل تجاري أو بطاقة ضريبية أو غيرها', '#reports-docs']);
    if (A.late90.length) al.push(['crit', A.late90.length, 'عملاء عليهم مستحقات أكثر من 90 يومًا', 'راجع أعمار الديون', '#reports-balances']);
    if (A.soon.length) al.push(['warn', A.soon.length, 'إجراءات تستحق قريبًا', 'خلال ' + (settings().alertDays || 7) + ' أيام', '#tasks-soon']);
    if (A.docsSoon.length) al.push(['warn', A.docsSoon.length, 'مستندات تنتهي قريبًا', 'خلال ' + (settings().docAlertDays || 30) + ' يومًا', '#reports-docs']);
    if (A.unbilled.length) al.push(['info', A.unbilled.length, 'مطالبات أتعاب لم تصدر لهذا الشهر', 'اضغط لإصدارها من العقود', 'act:gen-fees']);
    if (A.missingTasks.length) al.push(['info', A.missingTasks.length, 'إجراءات دورية لم تُولّد للفترة السابقة', 'اضغط لتوليدها من الالتزامات', 'act:gen-tasks']);
    h += '<section><div class="alerts">' + (al.length ? al.map(a => {
      const attr = a[4].startsWith('act:') ? ' data-act="' + a[4].slice(4) + '" role="button" tabindex="0"' : ' href="' + a[4] + '"';
      return '<a class="alert ' + a[0] + '"' + attr + '><b>' + int(a[1]) + '</b><div><strong>' + esc(a[2]) + '</strong><small>' + esc(a[3]) + '</small></div></a>';
    }).join('') : '<div class="alert ok"><b>✓</b><div><strong>لا توجد ملاحظات</strong><small>كل الإجراءات والمستندات في مواعيدها</small></div></div>') + '</div></section>';

    // top balances + trend
    const top = M.clients.map(c => [c, M.acct.get(c.id).balance]).filter(x => x[1] > 0).sort((a, b) => b[1] - a[1]).slice(0, 8);
    const maxB = top.length ? top[0][1] : 1;
    const months = []; for (let i = 5; i >= 0; i--) months.push(C.addMonths(M.ym, -i));
    const tr = months.map(ym => { let f = 0, p = 0; for (const t of M.txs) if (t.date.slice(0, 7) === ym) { if (t.kind === 'fee') f += Number(t.amount) || 0; if (t.kind === 'payment') p += Number(t.amount) || 0; } return { ym, f, p }; });
    const maxT = Math.max(1, ...tr.map(x => Math.max(x.f, x.p)));
    h += '<div class="grid2"><section class="card"><header><h2>أعلى الأرصدة المستحقة</h2><span class="hint">' + esc(cur()) + '</span><span class="grow"></span><a href="#reports-balances" class="btn sm">كل الأرصدة</a></header><div class="body">' +
      (top.length ? top.map(([c, b]) => '<div class="chart-row"><a class="lbl" href="#client-statement-' + c.id + '" title="' + esc(c.name) + '">' + esc(c.name) + '</a><div class="bar"><i style="width:' + (b / maxB * 100).toFixed(1) + '%;background:var(--crit);opacity:.75"></i></div><div class="v">' + money(b) + '</div></div>').join('') : '<div class="empty">لا توجد أرصدة مستحقة</div>') +
      '</div></section><section class="card"><header><h2>المطالبات والتحصيلات – آخر 6 أشهر</h2></header><div class="body"><div class="trend">' +
      tr.map(x => '<div class="col" title="' + esc(monthLabel(x.ym)) + ': مطالبات ' + money(x.f) + ' · تحصيلات ' + money(x.p) + '"><i style="height:' + (x.f / maxT * 100).toFixed(1) + '%;background:var(--accent)"></i><i style="height:' + (x.p / maxT * 100).toFixed(1) + '%;background:var(--ok)"></i></div>').join('') +
      '</div><div class="trend-labels">' + tr.map(x => '<span>' + esc(MONTHS[Number(x.ym.slice(5)) - 1]) + '</span>').join('') + '</div>' +
      '<div class="legend"><span><i style="background:var(--accent)"></i>مطالبات الأتعاب</span><span><i style="background:var(--ok)"></i>التحصيلات</span></div></div></section></div>';

    // upcoming tasks
    const limit = C.addDays(M.today, 14);
    const up = M.tasks.filter(t => t.status !== 'done' && t.dueDate && t.dueDate <= limit).slice(0, 15);
    h += '<section class="card"><header><h2>الإجراءات المستحقة خلال 14 يومًا والمتأخرة</h2><span class="grow"></span><a class="btn sm" href="#tasks-open">كل الإجراءات</a></header>' + taskTable(up, true) + '</section>';
    return h + '</div>';
  }

  /* ---------------- clients ---------------- */
  function filteredClients() {
    const u = S.ui, M = S.M, q = u.cQ.trim().toLowerCase();
    return M.clients.filter(c => {
      if (u.cStatus === 'active' && c.status === 'inactive') return false;
      if (u.cStatus === 'inactive' && c.status !== 'inactive') return false;
      if (u.cStaff && c.assignedTo !== u.cStaff) return false;
      if (u.cLaw && lawOf(c) !== u.cLaw) return false;
      const b = M.acct.get(c.id).balance;
      if (u.cBal === 'due' && !(b > 0.004)) return false;
      if (u.cBal === 'adv' && !(b < -0.004)) return false;
      if (q && ![c.code, c.name, c.taxRegNo, c.phone, c.contactName, c.activity].some(v => String(v || '').toLowerCase().includes(q))) return false;
      return true;
    });
  }
  function pageClients() {
    const M = S.M, u = S.ui, list = filteredClients();
    let h = topBar('العملاء', int(list.length) + ' عميل', actBtn('add-client', '+ عميل جديد', 'primary') + expBtns('clients')) + '<div class="content">';
    h += '<div class="toolbar"><input type="search" placeholder="بحث بالاسم أو الكود أو رقم التسجيل الضريبي…" data-ui="cQ" data-live value="' + esc(u.cQ) + '">' +
      '<select data-ui="cStatus">' + options([['active', 'العملاء النشطون'], ['inactive', 'العملاء المتوقفون'], ['all', 'كل العملاء']], u.cStatus) + '</select>' +
      '<select data-ui="cStaff">' + options(staffOptions(), u.cStaff, 'جميع الموظفين') + '</select>' +
      '<select data-ui="cLaw">' + options(Object.entries(TAX_LAW), u.cLaw, 'كل الأنظمة الضريبية') + '</select>' +
      '<select data-ui="cBal">' + options([['due', 'عليهم مستحقات'], ['adv', 'لهم دفعات مقدمة']], u.cBal, 'كل الأرصدة') + '</select></div>';
    h += '<div class="card tablewrap"><table><thead><tr><th>الكود</th><th>العميل</th><th>الموظف المسؤول</th><th>الهاتف</th><th>الالتزامات</th><th class="n">الرصيد (' + esc(cur()) + ')</th><th>آخر تحصيل</th><th>أقدم مستحق</th></tr></thead><tbody>';
    if (!list.length) h += '<tr><td colspan="8" class="empty">لا يوجد عملاء مطابقون</td></tr>';
    let tot = 0;
    for (const c of list) {
      const a = M.acct.get(c.id); tot += a.balance;
      const obs = C.effectiveObligations(M, c).map(id => M.obById.get(id)).filter(Boolean);
      h += '<tr class="click" data-client="' + c.id + '"><td class="code">' + esc(c.code) + '</td><td><div class="stack"><strong>' + esc(c.name) + '</strong><span class="sub">' + esc(c.legalForm || '') + (c.activity ? ' · ' + esc(c.activity) : '') + '</span></div>' + (lawOf(c) === '6-2025' ? ' <span class="pill info">قانون 6</span>' : '') + (c.status === 'inactive' ? ' <span class="pill mute">متوقف</span>' : '') + '</td>' +
        '<td>' + esc(staffName(c.assignedTo)) + '</td><td class="num">' + esc(c.phone || '') + '</td><td><span class="sub">' + (obs.length ? obs.length + ' التزام' : '—') + '</span></td>' + balCell(a.balance) +
        '<td class="num">' + esc(a.lastPayment || '—') + '</td><td>' + (a.aging.oldest && a.balance > 0 ? '<span class="pill ' + (a.aging.oldestDays > 90 ? 'crit' : a.aging.oldestDays > 30 ? 'warn' : 'mute') + '">' + int(a.aging.oldestDays) + ' يوم</span>' : '') + '</td></tr>';
    }
    h += '</tbody>' + (list.length ? '<tfoot><tr><td colspan="5">الإجمالي</td>' + balCell(tot) + '<td colspan="2"></td></tr></tfoot>' : '') + '</table></div>';
    return h + '</div>';
  }

  function pageClient(id, tab) {
    const M = S.M, c = M.clientById.get(id);
    if (!c) return topBar('العميل') + '<div class="content"><div class="card"><div class="empty">العميل غير موجود. <a href="#clients">العودة إلى العملاء</a></div></div></div>';
    const a = M.acct.get(c.id);
    const actions = actBtn('add-tx', '+ تحصيل', 'primary', ' data-kind="payment" data-client="' + c.id + '"') + actBtn('add-tx', '+ مطالبة أتعاب', '', ' data-kind="fee" data-client="' + c.id + '"') +
      actBtn('add-tx', '+ مصروف', '', ' data-kind="expense" data-client="' + c.id + '"') + actBtn('edit-client', 'تعديل البيانات', '', ' data-id="' + c.id + '"');
    let h = topBar(c.name, '<a href="#clients">العملاء</a> · <span class="code">' + esc(c.code) + '</span>' + (c.status === 'inactive' ? ' · متوقف' : ''), actions) + '<div class="content">';
    h += '<section class="card"><div class="body client-head"><div class="kv">' +
      '<span>الشكل القانوني</span><span>' + esc(c.legalForm || '—') + '</span><span>النشاط</span><span>' + esc(c.activity || '—') + '</span>' +
      '<span>القانون الضريبي</span><span>' + esc(TAX_LAW[lawOf(c)]) + '</span><span>رقم التسجيل الضريبي</span><span class="num">' + esc(c.taxRegNo || '—') + '</span><span>المأمورية</span><span>' + esc(c.taxOffice || '—') + '</span>' +
      '<span>الموظف المسؤول</span><span>' + esc(staffName(c.assignedTo) || '—') + '</span></div><div class="kv">' +
      '<span>المسؤول لدى العميل</span><span>' + esc(c.contactName || '—') + '</span><span>الهاتف</span><span class="num">' + esc(c.phone || '—') + '</span>' +
      '<span>البريد الإلكتروني</span><span>' + esc(c.email || '—') + '</span><span>آخر تحصيل</span><span class="num">' + esc(a.lastPayment || '—') + '</span></div>' +
      '<div class="bal"><small>الرصيد الحالي</small><b class="' + (a.balance > 0.004 ? 'pos' : a.balance < -0.004 ? 'neg' : '') + '">' + money(Math.abs(a.balance)) + ' ' + esc(cur()) + '</b><small>' + (a.balance > 0.004 ? 'مستحق على العميل' : a.balance < -0.004 ? 'دفعة مقدمة للعميل' : 'لا يوجد رصيد') + '</small>' +
      (a.balance > 0.004 ? '<div style="margin-top:8px"><button type="button" class="btn sm" data-act="reminder" data-id="' + c.id + '">رسالة تذكير بالمستحقات</button></div>' : '') + '</div></div></section>';
    const tasksOpen = M.tasks.filter(t => t.client === c.id && t.status !== 'done').length;
    const tabs = [['statement', 'كشف الحساب'], ['tasks', 'الإجراءات', tasksOpen], ['contracts', 'العقود والأتعاب'], ['info', 'البيانات والمستندات'], ['access', 'بيانات المنظومات']];
    h += '<div class="tabs">' + tabs.map(([k, l, n]) => '<a href="#client-' + k + '-' + c.id + '" class="' + (tab === k ? 'on' : '') + '">' + esc(l) + (n ? '<span class="count">' + n + '</span>' : '') + '</a>').join('') + '</div>';
    if (tab === 'statement') h += clientStatementTab(c);
    else if (tab === 'tasks') {
      const list = M.tasks.filter(t => t.client === c.id).sort((x, y) => (x.status === 'done') - (y.status === 'done') || (x.status === 'done' ? (y.dueDate || '').localeCompare(x.dueDate || '') : (x.dueDate || '').localeCompare(y.dueDate || '')));
      h += '<section class="card"><header><h2>إجراءات العميل</h2><span class="grow"></span>' + actBtn('add-task', '+ إجراء', '', ' data-client="' + c.id + '"') + '</header>' + taskTable(list, false) + '</section>';
    } else if (tab === 'contracts') {
      const list = M.contracts.filter(k => k.client === c.id);
      h += '<section class="card"><header><h2>عقود الأتعاب</h2><span class="grow"></span>' + actBtn('add-contract', '+ عقد', '', ' data-client="' + c.id + '"') + '</header>' + contractTable(list, false) + '</section>';
    } else if (tab === 'access') h += accessTab(c);
    else h += clientInfoTab(c);
    return h + '</div>';
  }

  function clientStatementTab(c) {
    const u = S.ui, st = C.statement(S.M, c.id, u.stFrom || '', u.stTo || '');
    let h = '<div class="toolbar"><label class="inl">من <input type="date" data-ui="stFrom" value="' + esc(u.stFrom) + '"></label><label class="inl">إلى <input type="date" data-ui="stTo" value="' + esc(u.stTo) + '"></label>' +
      (u.stFrom || u.stTo ? '<button type="button" class="btn sm" data-act="clear-st">كل الفترات</button>' : '') + '<span class="grow"></span>' +
      (S.downloads ? '<button type="button" class="btn primary" data-act="pdf-statement" data-id="' + c.id + '">⤓ كشف الحساب PDF</button><button type="button" class="btn" data-act="xlsx-statement" data-id="' + c.id + '">⤓ Excel</button>' : '') + '</div>';
    h += '<div class="card tablewrap"><table><thead><tr><th>التاريخ</th><th>الرقم</th><th>نوع الحركة</th><th>البيان</th><th class="n">مدين</th><th class="n">دائن</th><th class="n">الرصيد</th><th></th></tr></thead><tbody>';
    h += '<tr><td class="num">' + esc(st.from || c.openingDate || '') + '</td><td></td><td colspan="2"><strong>' + (st.from ? 'رصيد أول المدة' : 'الرصيد الافتتاحي') + '</strong></td><td></td><td></td>' + balCell(st.opening) + '<td></td></tr>';
    for (const l of st.lines) {
      const t = l.t;
      h += '<tr><td class="num">' + esc(t.date) + '</td><td class="num">' + esc(t.number || '') + '</td><td><span class="pill ' + KIND_PILL[t.kind] + '">' + esc(KIND[t.kind]) + '</span></td><td>' + esc(t.description || '') + (t.method ? '<div class="sub">' + esc(t.method) + (t.ref ? ' · ' + esc(t.ref) : '') + '</div>' : '') + '</td>' +
        amountCell(l.debit) + amountCell(l.credit) + balCell(l.balance) + '<td style="white-space:nowrap">' + (t.kind === 'payment' && S.downloads ? '<button type="button" class="btn sm" data-act="receipt" data-id="' + t.id + '">إيصال</button> ' : '') + actBtn('edit-tx', 'تعديل', 'sm', ' data-id="' + t.id + '"') + '</td></tr>';
    }
    if (!st.lines.length) h += '<tr><td colspan="8" class="empty">لا توجد حركات في هذه الفترة</td></tr>';
    h += '</tbody><tfoot><tr><td colspan="4">الإجمالي ورصيد آخر المدة</td>' + amountCell(st.debit) + amountCell(st.credit) + balCell(st.closing) + '<td></td></tr></tfoot></table></div>';
    return h;
  }

  function clientInfoTab(c) {
    const M = S.M;
    const obs = C.effectiveObligations(M, c).map(id => M.obById.get(id)).filter(Boolean);
    let h = '<div class="grid2e"><section class="card"><header><h2>البيانات الأساسية</h2></header><div class="body"><div class="kv">' +
      [['الكود', c.code], ['الاسم', c.name], ['الشكل القانوني', c.legalForm], ['النشاط', c.activity], ['رقم التسجيل الضريبي', c.taxRegNo], ['المأمورية', c.taxOffice],
        ['القانون الضريبي', TAX_LAW[lawOf(c)]], ['مسجل في القيمة المضافة', c.vatRegistered || (c.obligations || []).includes('vat') ? 'نعم' : 'لا'], ['مسجل في كسب العمل', c.payrollRegistered || (c.obligations || []).includes('payroll') ? 'نعم' : 'لا'], ['رقم التأمينات', c.insuranceNo], ['نهاية السنة المالية', MONTHS[(Number(c.fyEnd) || 12) - 1]], ['العنوان', c.address],
        ['بداية التعامل', c.startDate], ['الرصيد الافتتاحي', money(c.openingBalance) + (c.openingDate ? ' في ' + c.openingDate : '')], ['ملاحظات', c.notes]]
        .map(([k, v]) => '<span>' + esc(k) + '</span><span>' + esc(v || '—') + '</span>').join('') + '</div></div></section>';
    h += '<section class="card"><header><h2>الالتزامات الدورية</h2></header><div class="body">' +
      (obs.length ? '<div class="checks">' + obs.map(o => '<span class="pill acc">' + esc(o.name) + ' · ' + esc(OB_FREQ[freqFor(o, c)]) + '</span>').join('') + '</div>' : '<div class="empty">لم تُحدد التزامات دورية. عدّل بيانات العميل لاختيارها.</div>') +
      '</div><header style="border-top:1px solid var(--line)"><h2>المستندات الرسمية</h2></header><div class="tablewrap"><table><thead><tr><th>المستند</th><th>الرقم</th><th>تاريخ الانتهاء</th><th>الحالة</th></tr></thead><tbody>' +
      ((c.docs || []).length ? c.docs.map(d => { const s = docState(d); return '<tr><td>' + esc(d.name) + '</td><td class="num">' + esc(d.number || '') + '</td><td class="num">' + esc(d.expiry || '—') + '</td><td>' + s + '</td></tr>'; }).join('') : '<tr><td colspan="4" class="empty">لا توجد مستندات مسجلة</td></tr>') +
      '</tbody></table></div></section></div>';
    if (writable()) h += '<div class="toolbar"><span class="grow"></span><button type="button" class="btn danger" data-act="del-client" data-id="' + c.id + '">حذف العميل</button></div>';
    return h;
  }
  function docState(d) {
    if (!d.expiry) return '<span class="pill mute">بدون تاريخ انتهاء</span>';
    const days = C.daysBetween(S.M.today, d.expiry);
    if (days < 0) return '<span class="pill crit">منتهٍ منذ ' + int(-days) + ' يوم</span>';
    if (days <= Number(settings().docAlertDays || 30)) return '<span class="pill warn">ينتهي خلال ' + int(days) + ' يوم</span>';
    return '<span class="pill ok">ساري</span>';
  }

  /* ---------------- finance ---------------- */
  function filteredTx() {
    const u = S.ui, q = u.fQ.trim().toLowerCase();
    return S.M.txs.filter(t => {
      if (u.fKind && t.kind !== u.fKind) return false;
      if (u.fClient && t.client !== u.fClient) return false;
      if (u.fFrom && t.date < u.fFrom) return false;
      if (u.fTo && t.date > u.fTo) return false;
      if (q && ![t.description, t.ref, t.number, clientName(t.client), t.method].some(v => String(v || '').toLowerCase().includes(q))) return false;
      return true;
    }).slice().reverse();
  }
  function pageFinance() {
    const u = S.ui, list = filteredTx();
    const actions = actBtn('add-tx', '+ تحصيل', 'primary', ' data-kind="payment"') + actBtn('add-tx', '+ مطالبة أتعاب', '', ' data-kind="fee"') + actBtn('add-tx', '+ مصروف', '', ' data-kind="expense"') +
      actBtn('add-tx', '+ خصم / تسوية', '', ' data-kind="discount"') + actBtn('gen-fees', 'إصدار مطالبات الشهر من العقود');
    let h = topBar('الحركات المالية', 'مطالبات الأتعاب والمصروفات والتحصيلات', actions) + '<div class="content">';
    h += '<div class="toolbar"><input type="search" placeholder="بحث في البيان أو الرقم أو المرجع…" data-ui="fQ" data-live value="' + esc(u.fQ) + '">' +
      '<select data-ui="fKind">' + options(Object.entries(KIND), u.fKind, 'كل الحركات') + '</select>' +
      '<select data-ui="fClient">' + options(clientOptions(true), u.fClient, 'كل العملاء') + '</select>' +
      '<label class="inl">من <input type="date" data-ui="fFrom" value="' + esc(u.fFrom) + '"></label><label class="inl">إلى <input type="date" data-ui="fTo" value="' + esc(u.fTo) + '"></label><span class="grow"></span>' + expBtns('finance') + '</div>';
    let dr = 0, cr = 0;
    h += '<div class="card tablewrap"><table><thead><tr><th>التاريخ</th><th>الرقم</th><th>نوع الحركة</th><th>العميل</th><th>البيان</th><th class="n">مدين</th><th class="n">دائن</th><th></th></tr></thead><tbody>';
    for (const t of list) {
      const a = Number(t.amount) || 0, s = C.sign(t.kind); if (s > 0) dr += a; else cr += a;
      h += '<tr><td class="num">' + esc(t.date) + '</td><td class="num">' + esc(t.number || '') + '</td><td><span class="pill ' + KIND_PILL[t.kind] + '">' + esc(KIND[t.kind]) + '</span></td>' +
        '<td><a href="#client-statement-' + t.client + '">' + esc(clientName(t.client)) + '</a></td><td>' + esc(t.description || '') + (t.method ? '<div class="sub">' + esc(t.method) + (t.ref ? ' · ' + esc(t.ref) : '') + '</div>' : '') + '</td>' +
        amountCell(s > 0 ? a : 0) + amountCell(s < 0 ? a : 0) + '<td style="white-space:nowrap">' + (t.kind === 'payment' && S.downloads ? '<button type="button" class="btn sm" data-act="receipt" data-id="' + t.id + '">إيصال</button> ' : '') + actBtn('edit-tx', 'تعديل', 'sm', ' data-id="' + t.id + '"') + '</td></tr>';
    }
    if (!list.length) h += '<tr><td colspan="8" class="empty">لا توجد حركات مطابقة</td></tr>';
    h += '</tbody>' + (list.length ? '<tfoot><tr><td colspan="5">الإجمالي (' + int(list.length) + ' حركة)</td>' + amountCell(dr) + amountCell(cr) + '<td></td></tr></tfoot>' : '') + '</table></div>';
    return h + '</div>';
  }

  /* ---------------- tasks ---------------- */
  function taskTable(list, showClient) {
    let h = '<div class="tablewrap"><table><thead><tr><th>الموعد</th>' + (showClient !== false ? '<th>العميل</th>' : '') + '<th>الإجراء</th><th>الفترة</th><th>المسؤول</th><th>الحالة</th><th>ملاحظات</th></tr></thead><tbody>';
    if (!list.length) h += '<tr><td colspan="7" class="empty">لا توجد إجراءات</td></tr>';
    for (const t of list) {
      const rowCls = t.state === 'overdue' ? 'row-crit' : t.state === 'soon' ? 'row-warn' : '';
      const due = '<td class="num"><div class="stack"><span>' + esc(t.dueDate || '—') + '</span>' + (t.state === 'overdue' ? '<span class="sub" style="color:var(--crit)">متأخر ' + int(t.daysLate) + ' يوم</span>' : t.state === 'done' && t.doneDate ? '<span class="sub">أُنجز ' + esc(t.doneDate) + '</span>' : '') + '</div></td>';
      const status = writable() ? '<select data-task-status="' + t.id + '" aria-label="الحالة">' + options(Object.entries(TSTATUS), t.status || 'pending') + '</select>' : '<span class="pill ' + TSTATUS_PILL[t.status || 'pending'] + '">' + esc(TSTATUS[t.status || 'pending']) + '</span>';
      h += '<tr class="' + rowCls + '">' + due + (showClient !== false ? '<td><a href="#client-tasks-' + t.client + '">' + esc(clientName(t.client)) + '</a></td>' : '') +
        '<td><a href="#" data-act="edit-task" data-id="' + t.id + '">' + esc(t.title) + '</a></td><td>' + esc(periodLabel(t)) + '</td><td>' + esc(staffName(t.assignedTo)) + '</td><td>' + status + '</td><td><span class="sub">' + esc([t.ref, t.notes].filter(Boolean).join(' · ')) + '</span></td></tr>';
    }
    return h + '</tbody></table></div>';
  }
  function pageTasks(tab) {
    const M = S.M, u = S.ui, q = u.tQ.trim().toLowerCase();
    const base = M.tasks.filter(t => (!u.tClient || t.client === u.tClient) && (!u.tStaff || t.assignedTo === u.tStaff) && (!u.tOb || t.obligation === u.tOb) &&
      (!q || [t.title, t.notes, t.ref, clientName(t.client)].some(v => String(v || '').toLowerCase().includes(q))));
    const groups = {
      overdue: base.filter(t => t.state === 'overdue'), soon: base.filter(t => t.state === 'soon'), open: base.filter(t => t.status !== 'done'),
      waiting: base.filter(t => t.status === 'waiting'), done: base.filter(t => t.status === 'done').sort((a, b) => (b.doneDate || b.dueDate || '').localeCompare(a.doneDate || a.dueDate || '')), all: base
    };
    const tabs = [['overdue', 'متأخرة'], ['soon', 'تستحق قريبًا'], ['open', 'كل المفتوحة'], ['waiting', 'بانتظار العميل'], ['done', 'المنجزة'], ['all', 'الكل']];
    const actions = actBtn('add-task', '+ إجراء', 'primary') + actBtn('gen-tasks', 'توليد إجراءات الفترة من الالتزامات');
    let h = topBar('الإجراءات والمواعيد', 'الإقرارات الدورية ومهام العملاء', actions) + '<div class="content">';
    h += '<div class="toolbar"><input type="search" placeholder="بحث…" data-ui="tQ" data-live value="' + esc(u.tQ) + '">' +
      '<select data-ui="tClient">' + options(clientOptions(true), u.tClient, 'كل العملاء') + '</select>' +
      '<select data-ui="tStaff">' + options(staffOptions(), u.tStaff, 'جميع الموظفين') + '</select>' +
      '<select data-ui="tOb">' + options(M.obligations.map(o => [o.id, o.name]), u.tOb, 'كل أنواع الإجراءات') + '</select><span class="grow"></span>' + expBtns('tasks') + '</div>';
    h += '<div class="tabs">' + tabs.map(([k, l]) => '<a href="#tasks-' + k + '" class="' + (tab === k ? 'on' : '') + '">' + esc(l) + (['overdue', 'soon', 'waiting'].includes(k) && groups[k].length ? '<span class="count">' + groups[k].length + '</span>' : '') + '</a>').join('') + '</div>';
    S.taskView = groups[tab] || groups.open;
    h += '<div class="card">' + taskTable(S.taskView, true) + '</div>';
    return h + '</div>';
  }
  async function setTaskStatus(id, status) {
    if (!writable()) return;
    const raw = S.raw.tasks.find(t => t.id === id); if (!raw) return;
    const patch = { status, doneDate: status === 'done' ? (raw.doneDate || C.todayISO()) : '' };
    if (await write(() => S.db.doc('tasks/' + id).update(patch))) toast('تم الحفظ');
  }

  /* ---------------- contracts ---------------- */
  function contractTable(list, showClient) {
    let h = '<div class="tablewrap"><table><thead><tr>' + (showClient ? '<th>العميل</th>' : '') + '<th>الخدمة</th><th class="n">قيمة الأتعاب</th><th>الدورية</th><th>من</th><th>إلى</th><th class="n">القيمة السنوية</th><th>الحالة</th><th></th></tr></thead><tbody>';
    if (!list.length) h += '<tr><td colspan="9" class="empty">لا توجد عقود</td></tr>';
    let tot = 0;
    for (const k of list) {
      const yearly = (Number(k.amount) || 0) * (FREQ_PER_YEAR[k.freq] || 0);
      const on = k.active !== false && (!k.endDate || k.endDate >= S.M.today);
      if (on) tot += yearly;
      h += '<tr>' + (showClient ? '<td><a href="#client-contracts-' + k.client + '">' + esc(clientName(k.client)) + '</a></td>' : '') + '<td>' + esc(k.service) + (k.notes ? '<div class="sub">' + esc(k.notes) + '</div>' : '') + '</td>' + amountCell(k.amount) +
        '<td>' + esc(FREQ[k.freq] || '') + '</td><td class="num">' + esc(k.startDate || '') + '</td><td class="num">' + esc(k.endDate || '—') + '</td>' + amountCell(yearly) +
        '<td>' + (on ? '<span class="pill ok">ساري</span>' : '<span class="pill mute">منتهٍ / موقوف</span>') + '</td><td>' + actBtn('edit-contract', 'تعديل', 'sm', ' data-id="' + k.id + '"') + '</td></tr>';
    }
    h += '</tbody>' + (list.length ? '<tfoot><tr><td colspan="' + (showClient ? 6 : 5) + '">إجمالي القيمة السنوية للعقود السارية</td>' + amountCell(tot) + '<td colspan="2"></td></tr></tfoot>' : '') + '</table></div>';
    return h;
  }
  function filteredContracts() {
    const u = S.ui, q = u.kQ.trim().toLowerCase();
    return S.M.contracts.filter(k => {
      const on = k.active !== false && (!k.endDate || k.endDate >= S.M.today);
      if (u.kActive === 'active' && !on) return false;
      if (u.kActive === 'ended' && on) return false;
      if (q && ![k.service, k.notes, clientName(k.client)].some(v => String(v || '').toLowerCase().includes(q))) return false;
      return true;
    }).sort((a, b) => clientName(a.client).localeCompare(clientName(b.client), 'ar'));
  }
  function pageContracts() {
    const u = S.ui, list = filteredContracts();
    const yearly = list.filter(k => k.active !== false && (!k.endDate || k.endDate >= S.M.today)).reduce((s, k) => s + (Number(k.amount) || 0) * (FREQ_PER_YEAR[k.freq] || 0), 0);
    let h = topBar('العقود والأتعاب', 'متوسط الأتعاب الشهرية المتوقعة: ' + esc(moneyC(yearly / 12)), actBtn('add-contract', '+ عقد أتعاب', 'primary') + actBtn('gen-fees', 'إصدار مطالبات الشهر')) + '<div class="content">';
    h += '<div class="toolbar"><input type="search" placeholder="بحث بالعميل أو الخدمة…" data-ui="kQ" data-live value="' + esc(u.kQ) + '"><select data-ui="kActive">' + options([['active', 'العقود السارية'], ['ended', 'المنتهية والموقوفة'], ['all', 'كل العقود']], u.kActive) + '</select><span class="grow"></span>' + expBtns('contracts') + '</div>';
    h += '<div class="card">' + contractTable(list, true) + '</div>';
    h += '<div class="note info">يُصدر البرنامج مطالبة الأتعاب في أول كل فترة حسب دورية العقد، بدءًا من شهر بداية العقد. استخدم «إصدار مطالبات الشهر» لمراجعة المطالبات المستحقة وإصدارها دفعة واحدة.</div>';
    return h + '</div>';
  }

  /* ---------------- modal + toast ---------------- */
  function modal(title, body, footer, wide) {
    const root = $('#modal-root');
    root.innerHTML = '<div class="modal-bg"><div class="modal' + (wide ? ' wide' : '') + '" role="dialog" aria-modal="true" aria-labelledby="mtitle"><header><h2 id="mtitle">' + esc(title) + '</h2><button type="button" class="btn sm" data-close aria-label="إغلاق">✕</button></header><div class="body">' + body + '</div>' + (footer ? '<footer>' + footer + '</footer>' : '') + '</div></div>';
    const bg = root.firstChild;
    bg.addEventListener('click', e => { if (e.target === bg || e.target.closest('[data-close]')) closeModal(); });
    const first = bg.querySelector('input:not([type=checkbox]),select,textarea'); if (first) setTimeout(() => first.focus(), 0);
    return bg;
  }
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && $('#modal-root').innerHTML) closeModal(); });
  function closeModal() { $('#modal-root').innerHTML = ''; }
  let toastTimer;
  function toast(msg, bad) {
    let el = $('#toast'); if (!el) { el = document.createElement('div'); el.id = 'toast'; document.body.appendChild(el); }
    el.className = 'toast' + (bad ? ' bad' : ''); el.textContent = msg; el.hidden = false;
    clearTimeout(toastTimer); toastTimer = setTimeout(() => { el.hidden = true; }, bad ? 6000 : 2500);
  }
  async function write(fn) {
    try { await fn(); return true; } catch (e) {
      if (e && e.code === 'invalid_argument') { S.canWrite = false; scheduleRender(); }
      toast((e && e.message) || 'تعذّر الحفظ. تحقّق من الاتصال وحاول مرة أخرى.', true); return false;
    }
  }
  function confirmBox(title, text, okLabel, fn) {
    const bg = modal(title, '<p style="margin:0">' + esc(text) + '</p>', '<button type="button" class="btn" data-close>إلغاء</button><button type="button" class="btn primary" id="cf-ok">' + esc(okLabel) + '</button>');
    $('#cf-ok', bg).addEventListener('click', fn);
  }
  const fld = (id, label, input, full) => '<div class="field' + (full ? ' full' : '') + '"><label for="' + id + '">' + esc(label) + '</label>' + input + '</div>';
  const inp = (id, val, type, extra) => '<input type="' + (type || 'text') + '" id="' + id + '" value="' + esc(val == null ? '' : val) + '"' + (extra || '') + '>';
  const val = (id, bg) => { const el = $('#' + id, bg); return el ? el.value.trim() : ''; };

  /* ---------------- client form ---------------- */
  function openClientForm(c) {
    const isNew = !c; const M = S.M;
    c = c ? clean(c) : { id: uid(), code: C.nextClientCode(M.clients), name: '', legalForm: '', activity: '', taxRegNo: '', taxOffice: '', vatRegistered: false, payrollRegistered: false, taxLaw: '91-2005', insuranceNo: '', fyEnd: 12,
      contactName: '', phone: '', email: '', address: '', assignedTo: '', status: 'active', startDate: C.todayISO(), openingBalance: 0, openingDate: C.todayISO(), obligations: [], docs: [], notes: '' };
    const docs = (c.docs || []).slice();
    const body = '<form id="cl-form" class="form"><h3>البيانات الأساسية</h3>' +
      fld('c-code', 'الكود', inp('c-code', c.code)) + fld('c-name', 'اسم العميل', inp('c-name', c.name, 'text', ' required')) +
      fld('c-legal', 'الشكل القانوني', '<select id="c-legal">' + options(LEGAL, c.legalForm, '—') + '</select>') + fld('c-act', 'النشاط', inp('c-act', c.activity)) +
      fld('c-status', 'الحالة', '<select id="c-status">' + options([['active', 'نشط'], ['inactive', 'متوقف']], c.status) + '</select>') +
      fld('c-staff', 'الموظف المسؤول بالمكتب', '<select id="c-staff">' + options(staffOptions(), c.assignedTo, '—') + '</select>') +
      '<h3>البيانات الضريبية والتأمينية</h3>' +
      fld('c-law', 'القانون الضريبي', '<select id="c-law">' + options([['91-2005', 'قانون 91 لسنة 2005'], ['6-2025', 'قانون 6 لسنة 2025 (التيسيرات)']], lawOf(c)) + '</select>') +
      fld('c-tax', 'رقم التسجيل الضريبي', inp('c-tax', c.taxRegNo)) + fld('c-office', 'المأمورية', inp('c-office', c.taxOffice)) +
      fld('c-ins', 'رقم التأمينات', inp('c-ins', c.insuranceNo)) + fld('c-fy', 'نهاية السنة المالية', '<select id="c-fy">' + options(MONTHS.map((m, i) => [i + 1, m]), c.fyEnd || 12) + '</select>') +
      '<div class="field full"><label>التسجيل</label><div class="checks"><label><input type="checkbox" id="c-vat"' + (c.vatRegistered || (c.obligations || []).includes('vat') ? ' checked' : '') + '> مسجل في ضريبة القيمة المضافة</label>' +
      '<label><input type="checkbox" id="c-pay"' + (c.payrollRegistered || (c.obligations || []).includes('payroll') ? ' checked' : '') + '> مسجل في ضريبة كسب العمل (المرتبات)</label></div></div>' +
      '<div class="field full"><label>الالتزامات الدورية (تُولَّد منها الإجراءات والمواعيد تلقائيًا)</label><div class="checks">' +
      M.obligations.filter(o => o.active !== false || (c.obligations || []).includes(o.id)).map(o => '<label data-oblabel="' + esc(o.id) + '"><input type="checkbox" name="c-ob" value="' + esc(o.id) + '"' + ((c.obligations || []).includes(o.id) ? ' checked' : '') + '>' + esc(o.name) + ' <small data-ofreq="' + esc(o.id) + '">(' + esc(OB_FREQ[freqFor(o, c)]) + ')</small></label>').join('') + '</div><small id="c-law-note" style="color:var(--muted)"></small></div>' +
      '<h3>بيانات الاتصال</h3>' +
      fld('c-contact', 'المسؤول لدى العميل', inp('c-contact', c.contactName)) + fld('c-phone', 'الهاتف', inp('c-phone', c.phone, 'tel')) +
      fld('c-email', 'البريد الإلكتروني', inp('c-email', c.email, 'email')) + fld('c-addr', 'العنوان', inp('c-addr', c.address)) +
      '<h3>الحساب</h3>' +
      fld('c-start', 'بداية التعامل', inp('c-start', c.startDate, 'date')) + fld('c-open', 'الرصيد الافتتاحي (موجب = مستحق على العميل)', inp('c-open', c.openingBalance, 'number', ' step="0.01"')) +
      fld('c-opend', 'تاريخ الرصيد الافتتاحي', inp('c-opend', c.openingDate, 'date')) +
      '<h3>المستندات الرسمية وتواريخ انتهائها</h3><div class="field full" id="c-docs"></div>' +
      '<div class="field full"><button type="button" class="btn sm" id="c-add-doc" style="align-self:flex-start">+ إضافة مستند</button></div>' +
      fld('c-notes', 'ملاحظات', '<textarea id="c-notes">' + esc(c.notes || '') + '</textarea>', true) +
      '<datalist id="doc-names">' + DOC_NAMES.map(n => '<option value="' + esc(n) + '">').join('') + '</datalist>' +
      '<div class="field full err" id="c-err"></div></form>';
    const foot = '<button type="button" class="btn" data-close>إلغاء</button><button type="submit" form="cl-form" class="btn primary">حفظ</button>';
    const bg = modal(isNew ? 'عميل جديد' : 'تعديل · ' + c.name, body, foot, true);
    function drawDocs() {
      $('#c-docs', bg).innerHTML = docs.length ? docs.map((d, i) => '<div class="doc-row" data-i="' + i + '"><input type="text" list="doc-names" placeholder="اسم المستند" data-k="name" value="' + esc(d.name) + '"><input type="text" placeholder="الرقم" data-k="number" value="' + esc(d.number || '') + '">' +
        '<input type="date" title="تاريخ الإصدار" data-k="issueDate" value="' + esc(d.issueDate || '') + '"><input type="date" title="تاريخ الانتهاء" data-k="expiry" value="' + esc(d.expiry || '') + '"><button type="button" class="btn sm danger" data-rm="' + i + '">حذف</button></div>').join('') +
        '<small style="color:var(--muted)">الأعمدة: اسم المستند · الرقم · تاريخ الإصدار · تاريخ الانتهاء</small>' : '<small style="color:var(--muted)">لا توجد مستندات. أضف السجل التجاري والبطاقة الضريبية ليتنبه البرنامج قبل انتهائهما.</small>';
      $$('[data-k]', bg).forEach(el => el.addEventListener('input', () => { docs[Number(el.closest('[data-i]').dataset.i)][el.dataset.k] = el.value; }));
      $$('[data-rm]', bg).forEach(b => b.addEventListener('click', () => { docs.splice(Number(b.dataset.rm), 1); drawDocs(); }));
    }
    drawDocs();
    // registration ticks and the matching obligations stay in step; the tax law changes the shown VAT frequency
    const obBox = (id) => $('input[name=c-ob][value="' + id + '"]', bg);
    const link = (regId, obId) => {
      const r = $('#' + regId, bg), o = obBox(obId); if (!r || !o) return;
      r.addEventListener('change', () => { o.checked = r.checked; });
      o.addEventListener('change', () => { r.checked = o.checked; });
    };
    link('c-vat', 'vat'); link('c-pay', 'payroll');
    // annual income-tax returns form one group: when the law or legal form rules one out, the one that applies takes its place
    const INCOME = ['corp', 'indiv', 'simple25'];
    const syncLaw = () => {
      const tmp = { taxLaw: val('c-law', bg), legalForm: val('c-legal', bg) };
      $$('[data-ofreq]', bg).forEach(el => { const o = M.obById.get(el.dataset.ofreq); if (o) el.textContent = '(' + OB_FREQ[freqFor(o, tmp)] + ')'; });
      const hadIncome = INCOME.some(id => { const b = obBox(id); return b && b.checked; });
      $$('[data-oblabel]', bg).forEach(lb => {
        const o = M.obById.get(lb.dataset.oblabel), box = $('input', lb); if (!o) return;
        const ok = C.obligationAllowed(o, tmp);
        box.disabled = !ok; if (!ok) box.checked = false;
        lb.style.opacity = ok ? '' : '.45'; lb.title = ok ? '' : 'لا ينطبق على هذا العميل حسب القانون الضريبي أو الشكل القانوني';
      });
      if (hadIncome && !INCOME.some(id => { const b = obBox(id); return b && b.checked; })) {
        const pick = INCOME.map(id => obBox(id)).find(b => b && !b.disabled); if (pick) pick.checked = true;
      }
      ['vat', 'payroll'].forEach(id => { const b = obBox(id), r = $(id === 'vat' ? '#c-vat' : '#c-pay', bg); if (b && r) r.checked = b.checked; });
      $('#c-law-note', bg).textContent = tmp.taxLaw === '6-2025' ? 'عملاء قانون 6 لسنة 2025: إقرار القيمة المضافة ربع سنوي، وإقرار الدخل المبسط (نموذج 25) بدلًا من إقرار الدخل العادي. يمكن تعديل ذلك من الإعدادات.' : '';
    };
    $('#c-law', bg).addEventListener('change', syncLaw); $('#c-legal', bg).addEventListener('change', syncLaw); syncLaw();
    $('#c-add-doc', bg).addEventListener('click', () => { docs.push({ id: uid(), name: docs.length ? '' : 'السجل التجاري', number: '', issueDate: '', expiry: '' }); drawDocs(); });
    $('#cl-form', bg).addEventListener('submit', async (ev) => {
      ev.preventDefault();
      const n = Object.assign({}, c, {
        code: val('c-code', bg), name: val('c-name', bg), legalForm: val('c-legal', bg), activity: val('c-act', bg), status: val('c-status', bg), assignedTo: val('c-staff', bg),
        taxRegNo: val('c-tax', bg), taxOffice: val('c-office', bg), insuranceNo: val('c-ins', bg), fyEnd: Number(val('c-fy', bg)) || 12, vatRegistered: $('#c-vat', bg).checked, payrollRegistered: $('#c-pay', bg).checked, taxLaw: val('c-law', bg),
        obligations: $$('input[name=c-ob]:checked', bg).map(x => x.value),
        contactName: val('c-contact', bg), phone: val('c-phone', bg), email: val('c-email', bg), address: val('c-addr', bg),
        startDate: val('c-start', bg), openingBalance: C.round2(val('c-open', bg)), openingDate: val('c-opend', bg), notes: val('c-notes', bg),
        docs: docs.filter(d => (d.name || '').trim()).map(d => Object.assign({ id: d.id || uid() }, d, { name: d.name.trim() }))
      });
      const err = $('#c-err', bg);
      if (!n.name) return (err.textContent = 'اسم العميل مطلوب.');
      if (n.code && S.M.clients.some(x => x.id !== n.id && String(x.code).trim() === n.code)) return (err.textContent = 'الكود مستخدم لعميل آخر.');
      if (await write(() => S.db.doc('clients/' + n.id).set(clean(n)))) { closeModal(); toast('تم الحفظ'); if (isNew) go('client-info-' + n.id); }
    });
  }
  function deleteClient(id) {
    const M = S.M, c = M.clientById.get(id); if (!c) return;
    const used = M.txs.some(t => t.client === id) || M.tasks.some(t => t.client === id) || M.contracts.some(k => k.client === id);
    if (used) return modal('لا يمكن حذف العميل', '<p style="margin:0">هذا العميل مرتبط بحركات مالية أو إجراءات أو عقود. اجعل حالته «متوقف» بدلًا من الحذف حتى يبقى سجله محفوظًا.</p>', '<button type="button" class="btn primary" data-close>حسنًا</button>');
    confirmBox('حذف العميل', 'هل تريد حذف «' + c.name + '» نهائيًا؟', 'حذف', async () => { if (await write(() => S.db.doc('clients/' + id).delete())) { closeModal(); toast('تم الحذف'); go('clients'); } });
  }

  /* ---------------- transaction form ---------------- */
  function openTxForm(t, preset) {
    const isNew = !t; const M = S.M; preset = preset || {};
    t = t ? clean(t) : { id: uid(), kind: preset.kind || 'payment', client: preset.client || '', date: C.todayISO(), amount: '', description: '', service: '', method: METHODS[0], ref: '' };
    const body = '<form id="tx-form" class="form">' +
      '<div class="field full"><label>نوع الحركة</label><div class="checks">' + Object.entries(KIND).map(([k, l]) => '<label><input type="radio" name="tx-kind" value="' + k + '"' + (t.kind === k ? ' checked' : '') + '>' + esc(l) + '</label>').join('') + '</div></div>' +
      fld('tx-client', 'العميل', '<select id="tx-client" required>' + options(clientOptions(!isNew), t.client, 'اختر العميل…') + '</select>', true) +
      fld('tx-date', 'التاريخ', inp('tx-date', t.date, 'date', ' required')) + fld('tx-amount', 'المبلغ (' + cur() + ')', inp('tx-amount', t.amount, 'number', ' step="0.01" min="0" required')) +
      '<div class="field" data-for="fee">' + '<label for="tx-svc">الخدمة</label><select id="tx-svc">' + options(settings().services || [], t.service, '—') + '</select></div>' +
      '<div class="field" data-for="payment"><label for="tx-method">طريقة السداد</label><select id="tx-method">' + options(METHODS, t.method) + '</select></div>' +
      '<div class="field" data-for="payment"><label for="tx-ref">رقم الشيك / التحويل</label>' + inp('tx-ref', t.ref) + '</div>' +
      fld('tx-desc', 'البيان', inp('tx-desc', t.description, 'text', ' placeholder="مثال: أتعاب مسك الدفاتر عن شهر أكتوبر"'), true) +
      (isNew ? '' : '<div class="field full"><small style="color:var(--muted)">رقم المستند: ' + esc(t.number || '—') + (t.period ? ' · فترة العقد: ' + esc(monthLabel(t.period)) : '') + '</small></div>') +
      '<div class="field full err" id="tx-err"></div></form>';
    const foot = (isNew ? '' : '<button type="button" class="btn danger" id="tx-del">حذف</button><span style="flex:1"></span>') + '<button type="button" class="btn" data-close>إلغاء</button>' +
      (isNew ? '<button type="button" class="btn" id="tx-save-new">حفظ وإضافة أخرى</button>' : '') + '<button type="submit" form="tx-form" class="btn primary">حفظ</button>';
    const bg = modal(isNew ? 'حركة جديدة' : 'تعديل حركة', body, foot);
    const syncKind = () => { const k = $('input[name=tx-kind]:checked', bg).value; $$('[data-for]', bg).forEach(el => { el.style.display = el.dataset.for === k ? '' : 'none'; }); };
    $$('input[name=tx-kind]', bg).forEach(r => r.addEventListener('change', syncKind)); syncKind();
    $('#tx-svc', bg).addEventListener('change', () => { const d = $('#tx-desc', bg); if (!d.value && $('#tx-svc', bg).value) d.value = 'أتعاب ' + $('#tx-svc', bg).value; });
    async function save(again) {
      const kind = $('input[name=tx-kind]:checked', bg).value;
      const n = Object.assign({}, t, { kind, client: val('tx-client', bg), date: val('tx-date', bg), amount: C.round2(val('tx-amount', bg)), description: val('tx-desc', bg),
        service: kind === 'fee' ? val('tx-svc', bg) : '', method: kind === 'payment' ? val('tx-method', bg) : '', ref: kind === 'payment' ? val('tx-ref', bg) : '' });
      const err = $('#tx-err', bg);
      if (!n.client) return (err.textContent = 'اختر العميل.');
      if (!n.date) return (err.textContent = 'أدخل التاريخ.');
      if (!(n.amount > 0)) return (err.textContent = 'أدخل مبلغًا أكبر من صفر.');
      if (!n.description) n.description = kind === 'fee' && n.service ? 'أتعاب ' + n.service : KIND[kind];
      if (isNew || n.kind !== t.kind) { n.number = (kind === 'fee' || kind === 'payment') ? C.nextNumber(S.raw.transactions, kind) : ''; }
      if (isNew) { n.seq = C.nextSeq(S.raw.transactions); n.createdAt = new Date().toISOString(); }
      if (await write(() => S.db.doc('transactions/' + n.id).set(clean(n)))) {
        toast('تم الحفظ' + (n.number ? ' · رقم ' + n.number : ''));
        if (again) openTxForm(null, { kind, client: n.client }); else {
          closeModal();
          if (isNew && kind === 'payment' && S.downloads) offerReceipt(n);
        }
      }
    }
    $('#tx-form', bg).addEventListener('submit', (ev) => { ev.preventDefault(); save(false); });
    const sn = $('#tx-save-new', bg); if (sn) sn.addEventListener('click', () => save(true));
    const del = $('#tx-del', bg); if (del) del.addEventListener('click', () => confirmBox('حذف الحركة', 'هل تريد حذف هذه الحركة؟ سيتغير رصيد العميل تبعًا لذلك.', 'حذف', async () => { if (await write(() => S.db.doc('transactions/' + t.id).delete())) { closeModal(); toast('تم الحذف'); } }));
  }
  function offerReceipt(tx) {
    const bg = modal('تم تسجيل التحصيل', '<p style="margin:0">تم تسجيل تحصيل ' + esc(moneyC(tx.amount)) + ' من ' + esc(clientName(tx.client)) + ' برقم إيصال ' + esc(tx.number) + '.</p>', '<button type="button" class="btn" data-close>إغلاق</button><button type="button" class="btn primary" id="rc-go">⤓ تنزيل الإيصال PDF</button>');
    $('#rc-go', bg).addEventListener('click', () => { closeModal(); pdfReceipt(tx); });
  }

  /* ---------------- task form ---------------- */
  function openTaskForm(t, preset) {
    const isNew = !t; const M = S.M; preset = preset || {};
    t = t ? clean(t) : { id: uid(), client: preset.client || '', obligation: '', title: '', period: C.addMonths(M.ym, -1), dueDate: '', assignedTo: '', status: 'pending', doneDate: '', ref: '', notes: '' };
    const body = '<form id="tk-form" class="form">' +
      fld('tk-client', 'العميل', '<select id="tk-client" required>' + options(clientOptions(!isNew), t.client, 'اختر العميل…') + '</select>', true) +
      fld('tk-ob', 'نوع الإجراء', '<select id="tk-ob">' + options(M.obligations.map(o => [o.id, o.name]), t.obligation, 'إجراء آخر (يُكتب يدويًا)') + '</select>') +
      fld('tk-period', 'الفترة', inp('tk-period', t.period, 'month')) +
      fld('tk-title', 'وصف الإجراء', inp('tk-title', t.title, 'text', ' placeholder="مثال: تجديد السجل التجاري"'), true) +
      fld('tk-due', 'الموعد النهائي', inp('tk-due', t.dueDate, 'date')) + fld('tk-staff', 'المسؤول', '<select id="tk-staff">' + options(staffOptions(), t.assignedTo, '—') + '</select>') +
      fld('tk-status', 'الحالة', '<select id="tk-status">' + options(Object.entries(TSTATUS), t.status) + '</select>') + fld('tk-done', 'تاريخ الإنجاز', inp('tk-done', t.doneDate, 'date')) +
      fld('tk-ref', 'رقم الإقرار / المرجع', inp('tk-ref', t.ref)) + fld('tk-notes', 'ملاحظات', '<textarea id="tk-notes">' + esc(t.notes || '') + '</textarea>', true) +
      '<div class="field full err" id="tk-err"></div></form>';
    const foot = (isNew ? '' : '<button type="button" class="btn danger" id="tk-del">حذف</button><span style="flex:1"></span>') + '<button type="button" class="btn" data-close>إلغاء</button><button type="submit" form="tk-form" class="btn primary">حفظ</button>';
    const bg = modal(isNew ? 'إجراء جديد' : 'تعديل إجراء', body, foot);
    const autoFill = () => {
      const ob = M.obById.get(val('tk-ob', bg)), per = val('tk-period', bg);
      if (!ob) return;
      if (per) $('#tk-due', bg).value = C.dueDate(ob, per);
      $('#tk-title', bg).value = ob.name + (per ? ' – ' + monthLabel(per) : '');
      const c = M.clientById.get(val('tk-client', bg)); if (c && !val('tk-staff', bg)) $('#tk-staff', bg).value = c.assignedTo || '';
    };
    $('#tk-ob', bg).addEventListener('change', autoFill); $('#tk-period', bg).addEventListener('change', autoFill);
    $('#tk-client', bg).addEventListener('change', () => { const c = M.clientById.get(val('tk-client', bg)); if (c && !val('tk-staff', bg)) $('#tk-staff', bg).value = c.assignedTo || ''; });
    $('#tk-status', bg).addEventListener('change', () => { if (val('tk-status', bg) === 'done' && !val('tk-done', bg)) $('#tk-done', bg).value = C.todayISO(); });
    $('#tk-form', bg).addEventListener('submit', async (ev) => {
      ev.preventDefault();
      const n = Object.assign({}, t, { client: val('tk-client', bg), obligation: val('tk-ob', bg), period: val('tk-period', bg), title: val('tk-title', bg), dueDate: val('tk-due', bg),
        assignedTo: val('tk-staff', bg), status: val('tk-status', bg), doneDate: val('tk-status', bg) === 'done' ? (val('tk-done', bg) || C.todayISO()) : '', ref: val('tk-ref', bg), notes: val('tk-notes', bg) });
      const err = $('#tk-err', bg);
      if (!n.client) return (err.textContent = 'اختر العميل.');
      if (!n.title) return (err.textContent = 'اكتب وصف الإجراء.');
      if (!n.dueDate) return (err.textContent = 'أدخل الموعد النهائي.');
      if (await write(() => S.db.doc('tasks/' + n.id).set(clean(n)))) { closeModal(); toast('تم الحفظ'); }
    });
    const del = $('#tk-del', bg); if (del) del.addEventListener('click', () => confirmBox('حذف الإجراء', 'هل تريد حذف هذا الإجراء؟', 'حذف', async () => { if (await write(() => S.db.doc('tasks/' + t.id).delete())) { closeModal(); toast('تم الحذف'); } }));
  }

  /* ---------------- contract form ---------------- */
  function openContractForm(k, preset) {
    const isNew = !k; preset = preset || {};
    k = k ? clean(k) : { id: uid(), client: preset.client || '', service: (settings().services || [])[0] || '', amount: '', freq: 'monthly', startDate: C.todayISO().slice(0, 8) + '01', endDate: '', active: true, notes: '' };
    const body = '<form id="k-form" class="form">' +
      fld('k-client', 'العميل', '<select id="k-client" required>' + options(clientOptions(!isNew), k.client, 'اختر العميل…') + '</select>', true) +
      fld('k-svc', 'الخدمة', inp('k-svc', k.service, 'text', ' list="svc-list" required') + '<datalist id="svc-list">' + (settings().services || []).map(s => '<option value="' + esc(s) + '">').join('') + '</datalist>') +
      fld('k-amount', 'قيمة الأتعاب لكل فترة (' + cur() + ')', inp('k-amount', k.amount, 'number', ' step="0.01" min="0" required')) +
      fld('k-freq', 'الدورية', '<select id="k-freq">' + options(Object.entries(FREQ), k.freq) + '</select>') +
      fld('k-start', 'تاريخ البداية', inp('k-start', k.startDate, 'date', ' required')) + fld('k-end', 'تاريخ الانتهاء (اختياري)', inp('k-end', k.endDate, 'date')) +
      fld('k-active', 'الحالة', '<select id="k-active">' + options([['1', 'ساري'], ['0', 'موقوف']], k.active === false ? '0' : '1') + '</select>') +
      fld('k-notes', 'ملاحظات', '<textarea id="k-notes">' + esc(k.notes || '') + '</textarea>', true) + '<div class="field full err" id="k-err"></div></form>';
    const foot = (isNew ? '' : '<button type="button" class="btn danger" id="k-del">حذف</button><span style="flex:1"></span>') + '<button type="button" class="btn" data-close>إلغاء</button><button type="submit" form="k-form" class="btn primary">حفظ</button>';
    const bg = modal(isNew ? 'عقد أتعاب جديد' : 'تعديل عقد أتعاب', body, foot);
    $('#k-form', bg).addEventListener('submit', async (ev) => {
      ev.preventDefault();
      const n = Object.assign({}, k, { client: val('k-client', bg), service: val('k-svc', bg), amount: C.round2(val('k-amount', bg)), freq: val('k-freq', bg), startDate: val('k-start', bg), endDate: val('k-end', bg), active: val('k-active', bg) === '1', notes: val('k-notes', bg) });
      const err = $('#k-err', bg);
      if (!n.client) return (err.textContent = 'اختر العميل.');
      if (!n.service) return (err.textContent = 'اكتب الخدمة.');
      if (!(n.amount > 0)) return (err.textContent = 'أدخل قيمة الأتعاب.');
      if (n.endDate && n.endDate < n.startDate) return (err.textContent = 'تاريخ الانتهاء يسبق تاريخ البداية.');
      if (await write(() => S.db.doc('contracts/' + n.id).set(clean(n)))) { closeModal(); toast('تم الحفظ'); }
    });
    const del = $('#k-del', bg); if (del) del.addEventListener('click', () => confirmBox('حذف العقد', 'هل تريد حذف هذا العقد؟ لن تُحذف المطالبات التي صدرت منه.', 'حذف', async () => { if (await write(() => S.db.doc('contracts/' + k.id).delete())) { closeModal(); toast('تم الحذف'); } }));
  }

  /* ---------------- batch generation ---------------- */
  function openGenFees() {
    let ym = S.M.ym;
    const bg = modal('إصدار مطالبات الأتعاب من العقود', '<div class="toolbar"><label class="inl">شهر المطالبة <input type="month" id="gf-month" value="' + ym + '"></label><label class="inl">تاريخ المطالبة <input type="date" id="gf-date" value="' + ym + '-01"></label></div><div id="gf-list"></div>',
      '<button type="button" class="btn" data-close>إلغاء</button><button type="button" class="btn primary" id="gf-go">إصدار المحدد</button>', true);
    let due = [];
    function draw() {
      due = C.feesDue(S.M, ym);
      $('#gf-list', bg).innerHTML = due.length ? '<div class="picker">' + due.map((d, i) => '<label><input type="checkbox" checked data-i="' + i + '"><span><strong>' + esc(d.client.name) + '</strong> <small>· ' + esc(d.contract.service) + ' · ' + esc(FREQ[d.contract.freq]) + '</small></span><span class="num">' + money(d.amount) + '</span><span></span></label>').join('') + '</div>' +
        '<div class="note info" style="margin-top:10px">الإجمالي: ' + esc(moneyC(due.reduce((s, d) => s + d.amount, 0))) + ' · ' + due.length + ' مطالبة. المطالبات التي صدرت من قبل لنفس العقد والشهر لا تظهر هنا.</div>' : '<div class="empty">لا توجد مطالبات مستحقة لهذا الشهر، أو صدرت كلها بالفعل.</div>';
    }
    draw();
    $('#gf-month', bg).addEventListener('change', (e) => { ym = e.target.value || S.M.ym; $('#gf-date', bg).value = ym + '-01'; draw(); });
    $('#gf-go', bg).addEventListener('click', async () => {
      const pick = $$('[data-i]:checked', bg).map(x => due[Number(x.dataset.i)]);
      if (!pick.length) return toast('لم تحدد أي مطالبة.', true);
      const date = $('#gf-date', bg).value || ym + '-01';
      let num = C.nextNumber(S.raw.transactions, 'fee'), seq = C.nextSeq(S.raw.transactions);
      $('#gf-go', bg).disabled = true;
      const ok = await write(async () => {
        for (const d of pick) {
          const n = { id: uid(), kind: 'fee', client: d.client.id, date, amount: d.amount, service: d.contract.service, contract: d.contract.id, period: ym,
            description: 'أتعاب ' + d.contract.service + ' عن ' + (d.contract.freq === 'monthly' ? 'شهر ' : 'الفترة من ') + monthLabel(ym), number: num++, seq: seq++, createdAt: new Date().toISOString() };
          await S.db.doc('transactions/' + n.id).set(n);
        }
      });
      if (ok) { closeModal(); toast('تم إصدار ' + pick.length + ' مطالبة'); }
    });
  }
  function openGenTasks() {
    let ym = C.addMonths(S.M.ym, -1);
    const bg = modal('توليد الإجراءات الدورية', '<div class="toolbar"><label class="inl">الفترة (الشهر الذي يُقدَّم عنه الإقرار) <input type="month" id="gt-month" value="' + ym + '"></label></div><div id="gt-list"></div>',
      '<button type="button" class="btn" data-close>إلغاء</button><button type="button" class="btn primary" id="gt-go">توليد المحدد</button>', true);
    let list = [];
    function draw() {
      list = C.tasksToGenerate(S.M, ym);
      $('#gt-list', bg).innerHTML = list.length ? '<div class="picker">' + list.map((d, i) => '<label><input type="checkbox" checked data-i="' + i + '"><span><strong>' + esc(d.client.name) + '</strong> <small>· ' + esc(d.obligation.name) + ' (' + esc(OB_FREQ[freqFor(d.obligation, d.client)]) + ')</small></span><span class="num">' + esc(d.dueDate) + '</span><span><small>' + esc(staffName(d.client.assignedTo)) + '</small></span></label>').join('') + '</div>' +
        '<div class="note info" style="margin-top:10px">يُحسب الموعد النهائي من إعدادات كل التزام (صفحة الإعدادات). الإجراءات الربع سنوية تظهر في شهور نهاية الربع، والسنوية في شهر نهاية السنة المالية للعميل.</div>' : '<div class="empty">لا توجد إجراءات جديدة لهذه الفترة. تأكد من اختيار الالتزامات الدورية في بيانات كل عميل.</div>';
    }
    draw();
    $('#gt-month', bg).addEventListener('change', (e) => { ym = e.target.value || ym; draw(); });
    $('#gt-go', bg).addEventListener('click', async () => {
      const pick = $$('[data-i]:checked', bg).map(x => list[Number(x.dataset.i)]);
      if (!pick.length) return toast('لم تحدد أي إجراء.', true);
      $('#gt-go', bg).disabled = true;
      const ok = await write(async () => {
        for (const d of pick) {
          const fq = freqFor(d.obligation, d.client); const per = fq === 'annual' ? 'السنة المالية المنتهية ' + monthLabel(ym) : fq === 'quarterly' ? 'الربع المنتهي ' + monthLabel(ym) : monthLabel(ym);
          const n = { id: uid(), client: d.client.id, obligation: d.obligation.id, period: ym, title: d.obligation.name + ' – ' + per, dueDate: d.dueDate, assignedTo: d.client.assignedTo || '', status: 'pending', doneDate: '', ref: '', notes: '' };
          await S.db.doc('tasks/' + n.id).set(n);
        }
      });
      if (ok) { closeModal(); toast('تم توليد ' + pick.length + ' إجراء'); }
    });
  }

  function openReminder(id) {
    const c = S.M.clientById.get(id), a = S.M.acct.get(id), s = settings();
    const text = 'السادة / ' + c.name + '\nتحية طيبة وبعد،\nنحيطكم علمًا بأن رصيد حسابكم لدى ' + (s.officeName || 'المكتب') + ' حتى تاريخ ' + S.M.today + ' هو ' + moneyC(a.balance) + ' مستحق السداد.\nنرجو التكرم بالسداد في أقرب وقت، ويسعدنا إرسال كشف الحساب التفصيلي عند الطلب.\nمع خالص التحية،\n' + (s.officeName || '') + (s.phone ? '\n' + s.phone : '');
    const phone = String(c.phone || '').replace(/\D/g, '').replace(/^0/, '20');
    const bg = modal('رسالة تذكير بالمستحقات', '<textarea id="rm-text" style="min-height:200px">' + esc(text) + '</textarea>',
      '<button type="button" class="btn" data-close>إغلاق</button><button type="button" class="btn" id="rm-copy">نسخ النص</button>' + (phone ? '<a class="btn primary" target="_blank" rel="noopener" id="rm-wa" href="#">فتح في واتساب</a>' : ''));
    $('#rm-copy', bg).addEventListener('click', async () => { try { await navigator.clipboard.writeText($('#rm-text', bg).value); toast('تم نسخ النص'); } catch (e) { $('#rm-text', bg).select(); toast('حدد النص وانسخه يدويًا', true); } });
    const wa = $('#rm-wa', bg); if (wa) { const upd = () => { wa.href = 'https://wa.me/' + phone + '?text=' + encodeURIComponent($('#rm-text', bg).value); }; upd(); $('#rm-text', bg).addEventListener('input', upd); }
  }

  /* ---------------- reports ---------------- */
  const REPORTS = [['balances', 'أرصدة العملاء وأعمار الديون'], ['statement', 'كشف حساب عميل'], ['collections', 'التحصيلات خلال فترة'], ['revenue', 'الإيرادات حسب الخدمة والعميل'],
    ['taskrep', 'تقرير الإجراءات'], ['docs', 'المستندات المنتهية وقريبة الانتهاء'], ['staff', 'عبء العمل لكل موظف']];
  const col = (h, n) => ({ h, n: !!n });
  const rangeCtx = (from, to) => (from || to) ? 'الفترة: ' + (from || 'البداية') + ' ← ' + (to || 'اليوم') : 'كل الفترات';

  function reportData(key) {
    const M = S.M, u = S.ui;
    if (key === 'clients') {
      const list = filteredClients();
      return { title: 'قائمة العملاء', ctx: int(list.length) + ' عميل', cols: [col('الكود'), col('العميل'), col('الشكل القانوني'), col('النشاط'), col('القانون الضريبي'), col('رقم التسجيل الضريبي'), col('المأمورية'), col('المسؤول'), col('الهاتف'), col('البريد الإلكتروني'), col('الحالة'), col('الرصيد', 1)],
        rows: list.map(c => [c.code, c.name, c.legalForm, c.activity, TAX_LAW[lawOf(c)], c.taxRegNo, c.taxOffice, staffName(c.assignedTo), c.phone, c.email, c.status === 'inactive' ? 'متوقف' : 'نشط', M.acct.get(c.id).balance]),
        foot: ['الإجمالي', '', '', '', '', '', '', '', '', '', '', list.reduce((s, c) => s + M.acct.get(c.id).balance, 0)], landscape: true };
    }
    if (key === 'finance') {
      const list = filteredTx();
      let dr = 0, cr = 0;
      const rows = list.map(t => { const a = Number(t.amount) || 0, s = C.sign(t.kind); if (s > 0) dr += a; else cr += a; return [t.date, t.number || '', KIND[t.kind], clientName(t.client), t.description || '', t.method ? t.method + (t.ref ? ' · ' + t.ref : '') : '', s > 0 ? a : '', s < 0 ? a : '']; });
      return { title: 'الحركات المالية', ctx: rangeCtx(u.fFrom, u.fTo) + (u.fKind ? ' · ' + KIND[u.fKind] : '') + (u.fClient ? ' · ' + clientName(u.fClient) : ''),
        cols: [col('التاريخ'), col('الرقم'), col('نوع الحركة'), col('العميل'), col('البيان'), col('طريقة السداد'), col('مدين', 1), col('دائن', 1)], rows, foot: ['الإجمالي', '', '', '', '', '', dr, cr], landscape: true };
    }
    if (key === 'tasks') {
      const list = S.taskView || M.tasks;
      return { title: 'الإجراءات والمواعيد', ctx: int(list.length) + ' إجراء', cols: [col('الموعد'), col('العميل'), col('الإجراء'), col('الفترة'), col('المسؤول'), col('الحالة'), col('أيام التأخير'), col('تاريخ الإنجاز'), col('المرجع')],
        rows: list.map(t => [t.dueDate, clientName(t.client), t.title, periodLabel(t), staffName(t.assignedTo), TSTATUS[t.status || 'pending'], t.daysLate || '', t.doneDate || '', t.ref || '']), landscape: true };
    }
    if (key === 'contracts') {
      const list = filteredContracts();
      let tot = 0;
      const rows = list.map(k => { const y = (Number(k.amount) || 0) * (FREQ_PER_YEAR[k.freq] || 0); const on = k.active !== false && (!k.endDate || k.endDate >= M.today); if (on) tot += y; return [clientName(k.client), k.service, Number(k.amount) || 0, FREQ[k.freq], k.startDate, k.endDate || '', y, on ? 'ساري' : 'منتهٍ / موقوف']; });
      return { title: 'عقود الأتعاب', ctx: int(list.length) + ' عقد', cols: [col('العميل'), col('الخدمة'), col('قيمة الأتعاب', 1), col('الدورية'), col('من'), col('إلى'), col('القيمة السنوية', 1), col('الحالة')], rows, foot: ['إجمالي العقود السارية', '', '', '', '', '', tot, ''], landscape: true };
    }
    if (key === 'certs') return certReport();
    if (key === 'balances') {
      const list = M.clients.filter(c => !u.rNonZero || Math.abs(M.acct.get(c.id).balance) > 0.004);
      const rows = list.map(c => { const a = M.acct.get(c.id), g = a.aging; return [c.code, c.name, staffName(c.assignedTo), a.opening, a.debit, a.credit, a.balance, g.b0, g.b30, g.b60, g.b90, a.lastPayment || '']; });
      const foot = ['الإجمالي', '', '', 0, 0, 0, 0, 0, 0, 0, 0, ''];
      for (const r of rows) for (let i = 3; i <= 10; i++) foot[i] += r[i];
      return { title: 'أرصدة العملاء وأعمار الديون', ctx: 'حتى ' + M.today + ' · ' + cur(), cols: [col('الكود'), col('العميل'), col('المسؤول'), col('الرصيد الافتتاحي', 1), col('المدين', 1), col('الدائن', 1), col('الرصيد', 1), col('حتى 30 يومًا', 1), col('31–60 يومًا', 1), col('61–90 يومًا', 1), col('أكثر من 90 يومًا', 1), col('آخر تحصيل')], rows, foot, landscape: true };
    }
    if (key === 'collections') {
      const list = M.txs.filter(t => t.kind === 'payment' && (!u.rFrom || t.date >= u.rFrom) && (!u.rTo || t.date <= u.rTo) && (!u.rClient || t.client === u.rClient));
      const byMethod = new Map(); let tot = 0;
      const rows = list.map(t => { const a = Number(t.amount) || 0; tot += a; byMethod.set(t.method || '—', (byMethod.get(t.method || '—') || 0) + a); return [t.date, t.number || '', clientName(t.client), t.method || '', t.ref || '', t.description || '', a]; });
      return { title: 'التحصيلات خلال فترة', ctx: rangeCtx(u.rFrom, u.rTo) + (u.rClient ? ' · ' + clientName(u.rClient) : ''), cols: [col('التاريخ'), col('رقم الإيصال'), col('العميل'), col('طريقة السداد'), col('المرجع'), col('البيان'), col('المبلغ', 1)], rows, foot: ['الإجمالي', '', '', '', '', '', tot],
        extra: { title: 'حسب طريقة السداد', cols: [col('طريقة السداد'), col('المبلغ', 1)], rows: Array.from(byMethod.entries()), foot: ['الإجمالي', tot] } };
    }
    if (key === 'revenue') {
      const r = C.revenue(M, u.rFrom, u.rTo);
      const per = new Map();
      for (const t of M.txs) {
        if ((u.rFrom && t.date < u.rFrom) || (u.rTo && t.date > u.rTo)) continue;
        const p = per.get(t.client) || { fee: 0, expense: 0, payment: 0, discount: 0 }; p[t.kind] += Number(t.amount) || 0; per.set(t.client, p);
      }
      const rows = Array.from(per.entries()).map(([id, p]) => [clientName(id), p.fee, p.expense, p.discount, p.payment]).sort((a, b) => b[1] - a[1]);
      const foot = ['الإجمالي', r.fees, r.expenses, r.discounts, r.payments];
      const svc = Array.from(r.byService.entries()).sort((a, b) => b[1] - a[1]);
      return { title: 'الإيرادات حسب العميل', ctx: rangeCtx(u.rFrom, u.rTo), cols: [col('العميل'), col('مطالبات الأتعاب', 1), col('مصروفات نيابة عن العميل', 1), col('خصومات وتسويات', 1), col('التحصيلات', 1)], rows, foot,
        extra: { title: 'مطالبات الأتعاب حسب الخدمة', cols: [col('الخدمة'), col('المبلغ', 1), col('النسبة')], rows: svc.map(([s, v]) => [s, v, r.fees ? (v / r.fees * 100).toFixed(1) + '%' : '']), foot: ['الإجمالي', r.fees, '100%'] } };
    }
    if (key === 'taskrep') {
      const list = M.tasks.filter(t => (!u.rFrom || (t.dueDate || '') >= u.rFrom) && (!u.rTo || (t.dueDate || '') <= u.rTo) && (!u.rClient || t.client === u.rClient) &&
        (!u.rStatus || (u.rStatus === 'overdue' ? t.state === 'overdue' : u.rStatus === 'open' ? t.status !== 'done' : t.status === u.rStatus)));
      return { title: 'تقرير الإجراءات', ctx: 'الموعد النهائي: ' + rangeCtx(u.rFrom, u.rTo).replace('الفترة: ', '') + (u.rClient ? ' · ' + clientName(u.rClient) : ''),
        cols: [col('الموعد'), col('العميل'), col('الإجراء'), col('المسؤول'), col('الحالة'), col('أيام التأخير'), col('تاريخ الإنجاز'), col('المرجع')],
        rows: list.map(t => [t.dueDate, clientName(t.client), t.title, staffName(t.assignedTo), t.state === 'overdue' ? 'متأخر' : TSTATUS[t.status || 'pending'], t.daysLate || '', t.doneDate || '', t.ref || '']), landscape: true };
    }
    if (key === 'docs') {
      const list = M.docs.filter(d => d.client.status !== 'inactive' && (u.rDocs === 'all' || d.state !== 'ok'));
      return { title: 'المستندات الرسمية للعملاء', ctx: u.rDocs === 'all' ? 'كل المستندات المسجل لها تاريخ انتهاء' : 'المنتهية والتي تنتهي خلال ' + (settings().docAlertDays || 30) + ' يومًا',
        cols: [col('العميل'), col('المستند'), col('الرقم'), col('تاريخ الانتهاء'), col('الحالة')],
        rows: list.map(d => [d.client.name, d.doc.name, d.doc.number || '', d.doc.expiry, d.state === 'expired' ? 'منتهٍ منذ ' + (-d.days) + ' يوم' : d.state === 'soon' ? 'ينتهي خلال ' + d.days + ' يوم' : 'ساري']) };
    }
    if (key === 'staff') {
      const mb = C.monthBounds(M.ym);
      const rows = M.staff.map(s => {
        const cl = M.active.filter(c => c.assignedTo === s.id).length;
        const ts = M.tasks.filter(t => t.assignedTo === s.id);
        return [s.name, s.title || '', cl, ts.filter(t => t.status !== 'done').length, ts.filter(t => t.state === 'overdue').length, ts.filter(t => t.status === 'done' && t.doneDate >= mb.start && t.doneDate <= mb.end).length];
      });
      const un = M.active.filter(c => !c.assignedTo).length;
      if (un) rows.push(['بدون موظف مسؤول', '', un, M.tasks.filter(t => !t.assignedTo && t.status !== 'done').length, M.tasks.filter(t => !t.assignedTo && t.state === 'overdue').length, '']);
      return { title: 'عبء العمل لكل موظف', ctx: monthLabel(M.ym), cols: [col('الموظف'), col('الوظيفة'), col('العملاء النشطون'), col('إجراءات مفتوحة'), col('إجراءات متأخرة'), col('أُنجز هذا الشهر')], rows };
    }
    return null;
  }
  function cellHtml(v, c) { return '<td' + (c.n ? ' class="n"' : '') + '>' + (c.n ? (v === '' || v == null ? '' : money(v)) : esc(v == null ? '' : v)) + '</td>'; }
  function tableHtml(rep) {
    return '<table><thead><tr>' + rep.cols.map(c => '<th' + (c.n ? ' class="n"' : '') + '>' + esc(c.h) + '</th>').join('') + '</tr></thead><tbody>' +
      (rep.rows.length ? rep.rows.map(r => '<tr>' + r.map((v, i) => cellHtml(v, rep.cols[i])).join('') + '</tr>').join('') : '<tr><td colspan="' + rep.cols.length + '" class="empty none">لا توجد بيانات</td></tr>') + '</tbody>' +
      (rep.foot && rep.rows.length ? '<tfoot><tr>' + rep.foot.map((v, i) => cellHtml(v, rep.cols[i])).join('') + '</tr></tfoot>' : '') + '</table>';
  }

  function pageReports(key) {
    const u = S.ui;
    let h = topBar('التقارير', 'كل تقرير قابل للتنزيل PDF للطباعة أو Excel') + '<div class="content">';
    h += '<div class="tabs">' + REPORTS.map(([k, l]) => '<a href="#reports-' + k + '" class="' + (key === k ? 'on' : '') + '">' + esc(l) + '</a>').join('') + '</div>';
    const range = '<label class="inl">من <input type="date" data-ui="rFrom" value="' + esc(u.rFrom) + '"></label><label class="inl">إلى <input type="date" data-ui="rTo" value="' + esc(u.rTo) + '"></label>';
    const clientSel = (all) => '<select data-ui="rClient">' + options(clientOptions(true), u.rClient, all ? 'كل العملاء' : 'اختر العميل…') + '</select>';
    if (key === 'statement') {
      h += '<div class="toolbar">' + clientSel(false) + range + '<span class="grow"></span>' + (u.rClient && S.downloads ? '<button type="button" class="btn primary" data-act="pdf-statement" data-id="' + esc(u.rClient) + '" data-range="r">⤓ PDF</button><button type="button" class="btn" data-act="xlsx-statement" data-id="' + esc(u.rClient) + '" data-range="r">⤓ Excel</button>' : '') + '</div>';
      if (!u.rClient || !S.M.clientById.get(u.rClient)) return h + '<div class="card"><div class="empty">اختر العميل لعرض كشف حسابه.</div></div></div>';
      const st = C.statement(S.M, u.rClient, u.rFrom, u.rTo);
      return h + '<div class="card tablewrap">' + tableHtml(statementRep(st)) + '</div></div>';
    }
    let filters = '';
    if (key === 'balances') filters = '<label class="inl"><input type="checkbox" data-ui="rNonZero"' + (u.rNonZero ? ' checked' : '') + '> العملاء الذين لهم رصيد فقط</label>';
    if (key === 'collections') filters = clientSel(true) + range;
    if (key === 'revenue') filters = range;
    if (key === 'taskrep') filters = clientSel(true) + '<select data-ui="rStatus">' + options([['open', 'المفتوحة'], ['overdue', 'المتأخرة'], ['done', 'المنجزة'], ['waiting', 'بانتظار العميل']], u.rStatus, 'جميع الحالات') + '</select>' + range;
    if (key === 'docs') filters = '<select data-ui="rDocs">' + options([['alert', 'المنتهية وقريبة الانتهاء'], ['all', 'كل المستندات']], u.rDocs) + '</select>';
    const rep = reportData(key);
    h += '<div class="toolbar">' + filters + '<span class="grow"></span>' + expBtns(key) + '</div>';
    h += '<section class="card"><header><h2>' + esc(rep.title) + '</h2><span class="hint">' + esc(rep.ctx || '') + '</span></header><div class="tablewrap">' + tableHtml(rep) + '</div></section>';
    if (rep.extra) h += '<section class="card"><header><h2>' + esc(rep.extra.title) + '</h2></header><div class="tablewrap">' + tableHtml(rep.extra) + '</div></section>';
    return h + '</div>';
  }
  function statementRep(st) {
    const rows = [[st.from || (st.client.openingDate || ''), '', st.from ? 'رصيد أول المدة' : 'الرصيد الافتتاحي', '', '', st.opening]];
    for (const l of st.lines) rows.push([l.t.date, l.t.number || '', KIND[l.t.kind] + (l.t.description ? ' – ' + l.t.description : '') + (l.t.method ? ' (' + l.t.method + (l.t.ref ? ' ' + l.t.ref : '') + ')' : ''), l.debit || '', l.credit || '', l.balance]);
    return { title: 'كشف حساب', cols: [col('التاريخ'), col('الرقم'), col('البيان'), col('مدين', 1), col('دائن', 1), col('الرصيد', 1)], rows, foot: ['الإجمالي ورصيد آخر المدة', '', '', st.debit, st.credit, st.closing] };
  }

  /* ---------------- PDF ---------------- */
  function pdfHead(title, ctx) {
    const s = settings();
    return '<header class="pdf-head"><div class="brand-wrap">' + (window.OF_LOGO ? '<img class="brand-logo" src="' + window.OF_LOGO.dark + '" alt="">' : '') + '<div><div class="brand">' + esc(s.officeName || 'مكتب المحاسبة') + '</div><div class="sub">' + esc(s.officeSub || '') + '</div><div class="sub">' + esc([s.address, s.phone, s.email].filter(Boolean).join(' · ')) + '</div>' + (s.taxNo ? '<div class="sub">رقم التسجيل الضريبي: ' + esc(s.taxNo) + '</div>' : '') + '</div></div>' +
      '<div class="ttl"><h1>' + esc(title) + '</h1>' + (ctx ? '<div class="ctx">' + esc(ctx) + '</div>' : '') + '<div class="ctx">تاريخ الإصدار: ' + esc(C.todayISO()) + '</div></div></header>';
  }
  /* PDF pages are drawn by the page itself instead of html2canvas: html2canvas places Arabic text word by word and
     drops the spaces between words and the order of brackets. Here the browser lays the document out, then each line
     of text is drawn whole with the canvas set to the element's direction, so shaping, spacing and bidi match the screen. */
  function cssColor(c) { return c && c !== 'transparent' && !/rgba\(\s*0,\s*0,\s*0,\s*0\s*\)/.test(c) ? c : null; }
  function pageBreaks(doc, pageH) {
    const top0 = doc.getBoundingClientRect().top, H = doc.scrollHeight;
    const blocks = Array.from(doc.querySelectorAll('tr, .box, .summary, .pdf-head, h2, p, .foot-note, .kv-row')).map(el => { const r = el.getBoundingClientRect(); return { t: r.top - top0, b: r.bottom - top0, h2: el.tagName === 'H2' }; });
    const cuts = []; let start = 0;
    while (H - start > pageH + 1) {
      let end = start + pageH;
      for (const k of blocks) if (k.t < end && k.b > end && k.t > start + pageH * 0.3) end = Math.min(end, k.t);
      // a heading never ends a page
      for (const k of blocks) if (k.h2 && k.t >= end - 40 && k.t < end && k.t > start + pageH * 0.3) end = k.t;
      cuts.push([start, end]); start = end;
    }
    cuts.push([start, H]);
    return cuts;
  }
  async function drawDocPage(doc, from, to, scale) {
    const o = doc.getBoundingClientRect(), W = o.width;
    const cv = document.createElement('canvas'); cv.width = Math.ceil(W * scale); cv.height = Math.ceil((to - from) * scale);
    const g = cv.getContext('2d'); g.scale(scale, scale); g.fillStyle = '#fff'; g.fillRect(0, 0, W, to - from); g.translate(-o.left, -o.top - from);
    const inPage = (r) => r.bottom > o.top + from && r.top < o.top + to;
    const els = [doc].concat(Array.from(doc.querySelectorAll('*')));
    for (const el of els) {
      const r = el.getBoundingClientRect(); if (!r.width || !inPage(r)) continue;
      const cs = getComputedStyle(el); if (cs.visibility === 'hidden' || cs.display === 'none') continue;
      const bg = cssColor(cs.backgroundColor);
      if (bg) { g.fillStyle = bg; const rad = parseFloat(cs.borderTopLeftRadius) || 0; g.beginPath(); if (rad && g.roundRect) g.roundRect(r.left, r.top, r.width, r.height, rad); else g.rect(r.left, r.top, r.width, r.height); g.fill(); }
      const sides = [['Top', r.left, r.top, r.width, 0], ['Bottom', r.left, r.bottom, r.width, 0], ['Left', r.left, r.top, 0, r.height], ['Right', r.right, r.top, 0, r.height]];
      const rad = parseFloat(cs.borderTopLeftRadius) || 0, bw = parseFloat(cs.borderTopWidth) || 0;
      if (rad && bw && cs.borderTopStyle !== 'none' && cs.borderTopWidth === cs.borderBottomWidth && cs.borderLeftWidth === cs.borderTopWidth && g.roundRect) {
        g.strokeStyle = cs.borderTopColor; g.lineWidth = bw; g.beginPath(); g.roundRect(r.left + bw / 2, r.top + bw / 2, r.width - bw, r.height - bw, rad); g.stroke();
      } else for (const [side, x, y, w, h] of sides) {
        const wd = parseFloat(cs['border' + side + 'Width']) || 0; if (!wd || cs['border' + side + 'Style'] === 'none') continue;
        const col = cssColor(cs['border' + side + 'Color']); if (!col) continue;
        g.fillStyle = col; g.fillRect(side === 'Right' ? x - wd : x, side === 'Bottom' ? y - wd : y, w || wd, h || wd);
      }
      if (el.tagName === 'IMG' && el.complete && el.naturalWidth) { try { g.drawImage(el, r.left, r.top, r.width, r.height); } catch (e) {} }
    }
    // text: one fillText per line of each text node
    const walker = document.createTreeWalker(doc, NodeFilter.SHOW_TEXT), range = document.createRange();
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const txt = n.nodeValue; if (!txt || !txt.trim()) continue;
      const pe = n.parentElement, cs = getComputedStyle(pe); if (cs.visibility === 'hidden') continue;
      const lines = [];
      const re = /\S+/g; let m;
      while ((m = re.exec(txt))) {
        range.setStart(n, m.index); range.setEnd(n, m.index + m[0].length);
        const rr = range.getClientRects(); if (!rr.length) continue;
        const rc = rr[0], mid = (rc.top + rc.bottom) / 2;
        let line = lines.find(l => Math.abs(l.mid - mid) < rc.height * 0.45);
        if (!line) { line = { mid, top: rc.top, bottom: rc.bottom, left: rc.left, right: rc.right, s: m.index, e: m.index + m[0].length }; lines.push(line); }
        line.left = Math.min(line.left, rc.left); line.right = Math.max(line.right, rc.right); line.top = Math.min(line.top, rc.top); line.bottom = Math.max(line.bottom, rc.bottom);
        line.s = Math.min(line.s, m.index); line.e = Math.max(line.e, m.index + m[0].length);
      }
      g.fillStyle = cs.color; g.font = cs.fontStyle + ' ' + cs.fontWeight + ' ' + cs.fontSize + ' ' + cs.fontFamily; g.textBaseline = 'middle';
      const rtl = cs.direction === 'rtl';
      try { g.direction = cs.direction; } catch (e) {}
      for (const l of lines) {
        if (l.bottom < o.top + from || l.top > o.top + to) continue;
        const t = txt.slice(l.s, l.e).replace(/\s+/g, ' ');
        g.textAlign = rtl ? 'right' : 'left';
        g.fillText(t, rtl ? l.right : l.left, (l.top + l.bottom) / 2 + 0.5);
      }
    }
    return cv;
  }
  async function makePdf(filename, inner, landscape) {
    if (!window.html2pdf) return toast('يتطلب تنزيل PDF اتصالًا بالإنترنت عند فتح الصفحة. اتصل ثم أعد تحميل الصفحة.', true);
    toast('جارٍ إعداد ملف PDF…');
    const holder = document.createElement('div'); holder.style.cssText = 'position:fixed;left:-20000px;top:0;';
    const doc = document.createElement('div'); doc.className = 'pdf-doc'; doc.dir = 'rtl'; doc.lang = 'ar';
    const innerW = landscape ? 281 : 194, innerH = (landscape ? 210 : 297) - 24;
    doc.style.width = innerW + 'mm';
    doc.innerHTML = inner + (settings().footer ? '<div class="foot-note">' + esc(settings().footer) + '</div>' : '');
    holder.appendChild(doc); document.body.appendChild(holder);
    try {
      if (document.fonts && document.fonts.ready) await document.fonts.ready;
      await Promise.all(Array.from(doc.querySelectorAll('img')).map(i => i.complete ? null : new Promise(r => { i.onload = i.onerror = r; })));
      const pxPerMm = doc.getBoundingClientRect().width / innerW;
      const pages = [];
      for (const [a, b] of pageBreaks(doc, innerH * pxPerMm)) pages.push(await drawDocPage(doc, a, b, 2));
      const blob = await html2pdf().set({ margin: [10, 8, 14, 8], image: { type: 'jpeg', quality: 0.95 }, jsPDF: { unit: 'mm', format: 'a4', orientation: landscape ? 'landscape' : 'portrait' }, pagebreak: { mode: [] } })
        .from(pages[0], 'canvas').toPdf().get('pdf').then(p => {
          for (let i = 1; i < pages.length; i++) {
            p.addPage(); const c = pages[i];
            p.addImage(c.toDataURL('image/jpeg', 0.95), 'JPEG', 8, 10, innerW, innerW * c.height / c.width);
          }
          const n = p.internal.getNumberOfPages(), w = p.internal.pageSize.getWidth(), hgt = p.internal.pageSize.getHeight();
          for (let i = 1; i <= n; i++) { p.setPage(i); p.setFontSize(8); p.setTextColor(120); p.text(i + ' / ' + n, w / 2, hgt - 6, { align: 'center' }); p.text(C.todayISO(), 8, hgt - 6); }
          return p.output('blob');
        });
      await saveFile(filename, blob);
    } catch (e) { toast((e && e.message) || 'تعذّر إعداد الملف.', true); }
    finally { holder.remove(); }
  }
  window.OF_PDF_TEST = async (inner, landscape) => { // used by tests only: returns page images
    const holder = document.createElement('div'); holder.style.cssText = 'position:fixed;left:-20000px;top:0;';
    const doc = document.createElement('div'); doc.className = 'pdf-doc'; doc.dir = 'rtl'; const innerW = landscape ? 281 : 194; doc.style.width = innerW + 'mm'; doc.innerHTML = inner;
    holder.appendChild(doc); document.body.appendChild(holder); await document.fonts.ready;
    const pxPerMm = doc.getBoundingClientRect().width / innerW, out = [];
    for (const [a, b] of pageBreaks(doc, ((landscape ? 210 : 297) - 24) * pxPerMm)) out.push((await drawDocPage(doc, a, b, 2)).toDataURL('image/png'));
    holder.remove(); return out;
  };
  function pdfReport(key) {
    const rep = reportData(key); if (!rep) return;
    let inner = pdfHead(rep.title, rep.ctx) + tableHtml(rep);
    if (rep.extra) inner += '<h2>' + esc(rep.extra.title) + '</h2>' + tableHtml(rep.extra);
    makePdf(fileName(rep.title) + '.pdf', inner, rep.landscape);
  }
  function pdfStatement(id, useReportRange) {
    const u = S.ui, from = useReportRange ? u.rFrom : u.stFrom, to = useReportRange ? u.rTo : u.stTo;
    const st = C.statement(S.M, id, from, to), c = st.client;
    const box = '<div class="box"><span><b>العميل:</b> ' + esc(c.name) + '</span><span><b>الكود:</b> ' + esc(c.code || '') + '</span><span><b>رقم التسجيل الضريبي:</b> ' + esc(c.taxRegNo || '—') + '</span><span><b>العنوان:</b> ' + esc(c.address || '—') + '</span></div>';
    const sum = '<div class="summary"><div><small>' + (st.from ? 'رصيد أول المدة' : 'الرصيد الافتتاحي') + '</small><b>' + money(st.opening) + '</b></div><div><small>إجمالي المدين</small><b>' + money(st.debit) + '</b></div><div><small>إجمالي الدائن</small><b>' + money(st.credit) + '</b></div><div><small>الرصيد ' + (st.closing > 0 ? 'المستحق' : st.closing < 0 ? '(دفعة مقدمة)' : '') + '</small><b>' + money(Math.abs(st.closing)) + ' ' + esc(cur()) + '</b></div></div>';
    const inner = pdfHead('كشف حساب عميل', rangeCtx(from, to)) + box + sum + tableHtml(statementRep(st)) +
      '<p class="foot-note">نرجو مطابقة الرصيد وإخطارنا بأي ملاحظات خلال 15 يومًا من تاريخ الكشف، وإلا اعتُبر الرصيد صحيحًا.</p>';
    makePdf(fileName('كشف حساب ' + c.name) + '.pdf', inner, false);
  }
  function pdfReceipt(t) {
    const c = S.M.clientById.get(t.client) || { name: '' };
    const inner = pdfHead('إيصال استلام نقدية', 'رقم الإيصال: ' + (t.number || '')) +
      '<div class="receipt">استلمنا من السادة / <b>' + esc(c.name) + '</b><br>مبلغًا وقدره: <span class="amt">' + money(t.amount) + ' ' + esc(cur()) + '</span><br>' +
      'وذلك ' + esc(t.description || 'سدادًا من الحساب') + '<br>طريقة السداد: ' + esc(t.method || '—') + (t.ref ? ' · رقم: ' + esc(t.ref) : '') + '<br>التاريخ: <span class="num">' + esc(t.date) + '</span><br>' +
      'الرصيد الحالي للحساب: ' + money(S.M.acct.get(t.client) ? S.M.acct.get(t.client).balance : 0) + ' ' + esc(cur()) + '</div>' +
      '<div class="sign"><div>توقيع المستلم</div><div>ختم المكتب</div></div>';
    makePdf(fileName('إيصال ' + (t.number || '') + ' ' + c.name) + '.pdf', inner, false);
  }
  const fileName = (s) => String(s).replace(/[\\/:*?"<>|]/g, ' ').trim();

  /* ---------------- Excel ---------------- */
  async function saveFile(filename, data) {
    if (!S.downloads) return toast('تنزيل الملفات غير متاح في هذا العرض.', true);
    try { await S.downloads.save({ filename, data }); } catch (e) { if (e && e.code !== 'declined') toast(e.message || 'تعذّر التنزيل.', true); }
  }
  function aoa(rep) {
    const rows = [rep.cols.map(c => c.h)].concat(rep.rows.map(r => r.map(v => v === '' ? null : v)));
    if (rep.foot && rep.rows.length) rows.push(rep.foot);
    return rows;
  }
  function xlsxBook(filename, sheets) {
    if (!window.XLSX) return toast('يتطلب تنزيل Excel اتصالًا بالإنترنت عند فتح الصفحة. اتصل ثم أعد تحميل الصفحة.', true);
    const wb = XLSX.utils.book_new();
    for (const s of sheets) {
      // office name and report title above the table (Excel files from here cannot carry the logo image)
      const head = s.noHead ? [] : [[settings().officeName || 'مكتب المحاسبة'], [s.name + (s.ctx ? ' — ' + s.ctx : '')], ['تاريخ الإصدار: ' + C.todayISO()], []];
      const ws = XLSX.utils.aoa_to_sheet(head.concat(s.rows));
      ws['!cols'] = s.rows[0].map((_, i) => ({ wch: Math.min(50, Math.max(10, ...s.rows.map(r => String(r[i] == null ? '' : r[i]).length + 2))) }));
      ws['!views'] = [{ RTL: true }];
      XLSX.utils.book_append_sheet(wb, ws, s.name.slice(0, 31).replace(/[\\/?*[\]:]/g, ' '));
    }
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    saveFile(filename, new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  }
  function xlsxReport(key) {
    const rep = reportData(key); if (!rep) return;
    const sheets = [{ name: rep.title, ctx: rep.ctx, rows: aoa(rep) }];
    if (rep.extra) sheets.push({ name: rep.extra.title, ctx: rep.ctx, rows: aoa(rep.extra) });
    xlsxBook(fileName(rep.title) + '.xlsx', sheets);
  }
  function xlsxStatement(id, useReportRange) {
    const u = S.ui, st = C.statement(S.M, id, useReportRange ? u.rFrom : u.stFrom, useReportRange ? u.rTo : u.stTo);
    xlsxBook(fileName('كشف حساب ' + st.client.name) + '.xlsx', [{ name: 'كشف حساب', ctx: st.client.name + ' · ' + rangeCtx(st.from, st.to), rows: aoa(statementRep(st)) }]);
  }

  /* ---------------- credentials vault ----------------
     Portal logins are stored encrypted. A random data key (AES-GCM 256) encrypts each client's record; the data key
     itself is stored wrapped with a key derived from the office master password (PBKDF2-SHA256). Nothing readable is
     saved: without the master password the records cannot be opened, and changing the password rewraps only the key.
     db rules keep both paths away from view-only people. */
  const CRED_SECTIONS = [
    { k: 'eta', t: 'بوابة مصلحة الضرائب (الإقرارات الإلكترونية)', f: [['user', 'اسم المستخدم'], ['pass', 'كلمة المرور', 1], ['email', 'البريد الإلكتروني المسجل'], ['phone', 'رقم الهاتف المسجل']] },
    { k: 'einv', t: 'منظومة الفاتورة الإلكترونية', f: [['user', 'اسم مستخدم البوابة'], ['pass', 'كلمة المرور', 1], ['clientId', 'Client ID'], ['secret1', 'Client Secret 1', 1], ['secret2', 'Client Secret 2', 1], ['activity', 'كود النشاط'], ['since', 'تاريخ الانضمام للمنظومة']] },
    { k: 'ereceipt', t: 'منظومة الإيصال الإلكتروني', f: [['user', 'اسم المستخدم'], ['pass', 'كلمة المرور', 1], ['clientId', 'Client ID'], ['secret', 'Client Secret', 1], ['pos', 'بيانات أجهزة نقاط البيع (POS)']] },
    { k: 'token', t: 'التوقيع الإلكتروني (Token)', f: [['issuer', 'الجهة المصدرة'], ['serial', 'رقم الشهادة / التوكن'], ['pin', 'الرقم السري (PIN)', 1], ['expiry', 'تاريخ الانتهاء'], ['holder', 'مكان التوكن (مع العميل / بالمكتب)']] },
    { k: 'insurance', t: 'بوابة التأمينات الاجتماعية', f: [['user', 'اسم المستخدم'], ['pass', 'كلمة المرور', 1], ['email', 'البريد الإلكتروني المسجل']] }
  ];
  const V = { doc: null, docLoaded: false, creds: new Map(), credsLoaded: false, key: null, plain: new Map(), shown: new Set(), busy: false };
  const hasCrypto = () => !!(window.crypto && crypto.subtle);
  const b64 = (buf) => { const a = new Uint8Array(buf); let s = ''; for (let i = 0; i < a.length; i++) s += String.fromCharCode(a[i]); return btoa(s); };
  const unb64 = (s) => { const b = atob(s), a = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) a[i] = b.charCodeAt(i); return a; };
  async function kekFrom(pass, salt, iterations) {
    const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(pass), 'PBKDF2', false, ['deriveKey']);
    return crypto.subtle.deriveKey({ name: 'PBKDF2', salt, iterations, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
  }
  async function seal(key, bytes) { const iv = crypto.getRandomValues(new Uint8Array(12)); const ct = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, bytes); return { iv: b64(iv), data: b64(ct) }; }
  async function unseal(key, box) { return new Uint8Array(await crypto.subtle.decrypt({ name: 'AES-GCM', iv: unb64(box.iv) }, key, unb64(box.data))); }
  async function wrapVault(pass, rawKey) {
    const salt = crypto.getRandomValues(new Uint8Array(16)), iterations = 310000;
    const kek = await kekFrom(pass, salt, iterations);
    return { salt: b64(salt), iterations, wrapped: await seal(kek, rawKey), updatedAt: new Date().toISOString() };
  }
  async function unlockVault(pass) {
    const d = V.doc; const kek = await kekFrom(pass, unb64(d.salt), d.iterations || 310000);
    const raw = await unseal(kek, d.wrapped); // throws on a wrong password
    return crypto.subtle.importKey('raw', raw, { name: 'AES-GCM' }, true, ['encrypt', 'decrypt']);
  }
  async function decryptAll() {
    if (!V.key) return;
    const out = new Map();
    for (const [id, d] of V.creds) { try { out.set(id, JSON.parse(new TextDecoder().decode(await unseal(V.key, d)))); } catch (e) { out.set(id, { _error: true }); } }
    V.plain = out; scheduleRender();
  }
  function lockVault() { V.key = null; V.plain = new Map(); V.shown.clear(); scheduleRender(); }

  function accessTab(c) {
    if (!hasCrypto()) return '<div class="card"><div class="empty">هذا المتصفح لا يدعم التشفير، لذا لا يمكن فتح بيانات الدخول هنا.</div></div>';
    if (!V.docLoaded || !V.credsLoaded) return '<div class="card"><div class="empty">جارٍ التحميل…</div></div>';
    if (!V.doc) {
      if (!writable()) return '<div class="card"><div class="empty">لم يُفعَّل حفظ بيانات الدخول بعد، أو ليست لديك صلاحية الاطلاع عليها.</div></div>';
      return '<section class="card"><header><h2>تفعيل خزنة بيانات الدخول</h2></header><div class="body"><form id="vault-setup" class="form">' +
        '<p class="field full" style="margin:0">تُحفظ أسماء المستخدمين وكلمات المرور الخاصة بالعملاء مشفرة بكلمة مرور رئيسية للمكتب، ولا يستطيع أحد قراءتها بدون هذه الكلمة. اختر كلمة قوية (12 حرفًا على الأقل) وشاركها فقط مع من يحتاجها من فريقك.</p>' +
        '<div class="note field full"><b>مهم:</b> إذا نُسيت كلمة المرور الرئيسية فلا توجد أي طريقة لاسترجاع البيانات المحفوظة. احتفظ بها في مكان آمن.</div>' +
        fld('vs-p1', 'كلمة المرور الرئيسية', '<input type="password" id="vs-p1" autocomplete="new-password" required>') + fld('vs-p2', 'تأكيد كلمة المرور', '<input type="password" id="vs-p2" autocomplete="new-password" required>') +
        '<div class="field full err" id="vs-err"></div><div class="field full"><button type="submit" class="btn primary" style="align-self:flex-start">تفعيل الخزنة</button></div></form></div></section>';
    }
    if (!V.key) {
      return '<section class="card"><header><h2>🔒 بيانات الدخول مقفلة</h2></header><div class="body"><form id="vault-unlock" class="toolbar">' +
        '<input type="password" id="vu-pass" placeholder="كلمة المرور الرئيسية للمكتب" autocomplete="current-password" style="min-width:260px"><button type="submit" class="btn primary">فتح</button><span class="err" id="vu-err"></span></form>' +
        '<p style="color:var(--muted);margin:10px 0 0;font-size:13px">تبقى مفتوحة حتى تضغط «قفل» أو تغلق الصفحة.</p></div></section>';
    }
    const d = V.plain.get(c.id) || {};
    if (d._error) return '<div class="card"><div class="empty">تعذّر فك تشفير بيانات هذا العميل.</div></div>';
    const raw = V.creds.get(c.id);
    let h = '<div class="toolbar"><span class="hint" style="color:var(--muted);font-size:12.5px">' + (raw && raw.updatedAt ? 'آخر تعديل: ' + esc(raw.updatedAt.slice(0, 16).replace('T', ' ')) : 'لا توجد بيانات محفوظة لهذا العميل بعد') + '</span><span class="grow"></span>' +
      actBtn('cred-edit', 'تعديل بيانات الدخول', 'primary', ' data-id="' + c.id + '"') + '<button type="button" class="btn" data-act="vault-lock">🔒 قفل</button></div>';
    h += '<div class="grid2e">';
    for (const sec of CRED_SECTIONS) {
      const v = d[sec.k] || {};
      const filled = sec.f.some(([k]) => v[k]);
      h += '<section class="card"><header><h2>' + esc(sec.t) + '</h2></header><div class="body">' + (filled ? '<div class="kv">' + sec.f.filter(([k]) => v[k]).map(([k, l, secret]) => credRow(c.id + '.' + sec.k + '.' + k, l, v[k], secret)).join('') + '</div>' : '<span style="color:var(--muted)">لم تُسجل بيانات</span>') +
        (sec.k === 'token' && v.expiry ? '<p style="font-size:12px;color:var(--muted);margin:8px 0 0">لتصلك تنبيهات قبل انتهاء التوكن، سجّل تاريخ انتهائه أيضًا في «المستندات الرسمية» للعميل.</p>' : '') + '</div></section>';
    }
    const extra = (d.extra || []).filter(x => x.label || x.value);
    h += '<section class="card"><header><h2>منظومات وبيانات أخرى</h2></header><div class="body">' + (extra.length || d.notes ? '<div class="kv">' + extra.map((x, i) => credRow(c.id + '.x.' + i, x.label || '—', x.value, x.secret)).join('') + (d.notes ? '<span>ملاحظات</span><span style="white-space:pre-wrap">' + esc(d.notes) + '</span>' : '') + '</div>' : '<span style="color:var(--muted)">لا يوجد</span>') + '</div></section>';
    return h + '</div>';
  }
  function credRow(key, label, value, secret) {
    const show = !secret || V.shown.has(key);
    return '<span>' + esc(label) + '</span><span style="display:flex;gap:6px;align-items:center;flex-wrap:wrap"><span class="code" style="direction:ltr;unicode-bidi:isolate">' + (show ? esc(value) : '••••••••') + '</span>' +
      (secret ? '<button type="button" class="btn sm" data-act="cred-show" data-key="' + esc(key) + '">' + (show ? 'إخفاء' : 'إظهار') + '</button>' : '') +
      '<button type="button" class="btn sm" data-act="cred-copy" data-key="' + esc(key) + '">نسخ</button></span>';
  }
  function credValue(key) {
    const [cid, sec, k] = key.split('.'); const d = V.plain.get(cid) || {};
    if (sec === 'x') { const x = (d.extra || [])[Number(k)]; return x ? x.value : ''; }
    return (d[sec] || {})[k] || '';
  }
  function bindAccess() {
    const vs = $('#vault-setup');
    if (vs) vs.addEventListener('submit', async (ev) => {
      ev.preventDefault();
      const p1 = $('#vs-p1').value, p2 = $('#vs-p2').value, err = $('#vs-err');
      if (p1.length < 12) return (err.textContent = 'كلمة المرور يجب ألا تقل عن 12 حرفًا.');
      if (p1 !== p2) return (err.textContent = 'كلمتا المرور غير متطابقتين.');
      if (V.busy) return; V.busy = true;
      try {
        const key = await crypto.subtle.generateKey({ name: 'AES-GCM', length: 256 }, true, ['encrypt', 'decrypt']);
        const raw = new Uint8Array(await crypto.subtle.exportKey('raw', key));
        const doc = await wrapVault(p1, raw);
        if (await write(() => S.db.doc('vault/main').set(doc))) { V.key = key; toast('تم تفعيل الخزنة'); }
      } finally { V.busy = false; }
    });
    const vu = $('#vault-unlock');
    if (vu) vu.addEventListener('submit', async (ev) => {
      ev.preventDefault();
      const pass = $('#vu-pass').value; if (!pass || V.busy) return;
      V.busy = true; $('#vu-err').textContent = 'جارٍ الفتح…';
      try { V.key = await unlockVault(pass); await decryptAll(); }
      catch (e) { const er = $('#vu-err'); if (er) er.textContent = 'كلمة المرور غير صحيحة.'; }
      finally { V.busy = false; }
    });
  }
  function openCredForm(cid) {
    const c = S.M.clientById.get(cid); if (!c || !V.key) return;
    const d = clean(V.plain.get(cid) || {}); const extra = (d.extra || []).slice();
    let body = '<form id="cr-form" class="form">';
    for (const sec of CRED_SECTIONS) {
      body += '<h3>' + esc(sec.t) + '</h3>';
      for (const [k, l, secret] of sec.f) body += fld('cr-' + sec.k + '-' + k, l, '<input type="' + (k === 'expiry' || k === 'since' ? 'date' : 'text') + '" id="cr-' + sec.k + '-' + k + '" value="' + esc((d[sec.k] || {})[k] || '') + '" autocomplete="off" spellcheck="false" dir="ltr"' + (secret ? ' data-secret' : '') + '>');
    }
    body += '<h3>منظومات وبيانات أخرى</h3><div class="field full" id="cr-extra"></div><div class="field full"><button type="button" class="btn sm" id="cr-add" style="align-self:flex-start">+ إضافة بيان</button></div>' +
      fld('cr-notes', 'ملاحظات', '<textarea id="cr-notes">' + esc(d.notes || '') + '</textarea>', true) + '</form>';
    const bg = modal('بيانات الدخول · ' + c.name, body, '<button type="button" class="btn" data-close>إلغاء</button><button type="submit" form="cr-form" class="btn primary">حفظ مشفرًا</button>', true);
    const drawExtra = () => {
      $('#cr-extra', bg).innerHTML = extra.map((x, i) => '<div class="doc-row" data-i="' + i + '" style="grid-template-columns:1.2fr 1.6fr auto auto"><input type="text" placeholder="البيان (مثال: بوابة الهيئة العامة للاستثمار)" data-k="label" value="' + esc(x.label || '') + '">' +
        '<input type="text" dir="ltr" placeholder="القيمة" data-k="value" value="' + esc(x.value || '') + '" autocomplete="off"><label style="display:flex;gap:4px;align-items:center;font-size:12px"><input type="checkbox" data-k="secret"' + (x.secret ? ' checked' : '') + '> سري</label><button type="button" class="btn sm danger" data-rm="' + i + '">حذف</button></div>').join('');
      $$('#cr-extra [data-k]', bg).forEach(el => el.addEventListener(el.type === 'checkbox' ? 'change' : 'input', () => { extra[Number(el.closest('[data-i]').dataset.i)][el.dataset.k] = el.type === 'checkbox' ? el.checked : el.value; }));
      $$('#cr-extra [data-rm]', bg).forEach(b => b.addEventListener('click', () => { extra.splice(Number(b.dataset.rm), 1); drawExtra(); }));
    };
    drawExtra();
    $('#cr-add', bg).addEventListener('click', () => { extra.push({ label: '', value: '', secret: false }); drawExtra(); });
    $('#cr-form', bg).addEventListener('submit', async (ev) => {
      ev.preventDefault();
      const n = {};
      for (const sec of CRED_SECTIONS) { const o = {}; for (const [k] of sec.f) { const v = $('#cr-' + sec.k + '-' + k, bg).value.trim(); if (v) o[k] = v; } if (Object.keys(o).length) n[sec.k] = o; }
      n.extra = extra.filter(x => (x.label || '').trim() || (x.value || '').trim()).map(x => ({ label: (x.label || '').trim(), value: (x.value || '').trim(), secret: !!x.secret }));
      n.notes = $('#cr-notes', bg).value.trim();
      const box = await seal(V.key, new TextEncoder().encode(JSON.stringify(n)));
      if (await write(() => S.db.doc('credentials/' + cid).set(Object.assign(box, { updatedAt: new Date().toISOString() })))) { V.plain.set(cid, n); closeModal(); toast('تم الحفظ مشفرًا'); }
    });
  }
  function openVaultPassword() {
    if (!V.key) return toast('افتح الخزنة أولًا من صفحة أي عميل (تبويب بيانات المنظومات).', true);
    const bg = modal('تغيير كلمة المرور الرئيسية', '<form id="vp-form" class="form">' + fld('vp-1', 'كلمة المرور الجديدة', '<input type="password" id="vp-1" autocomplete="new-password">') + fld('vp-2', 'تأكيدها', '<input type="password" id="vp-2" autocomplete="new-password">') + '<div class="field full err" id="vp-err"></div></form>',
      '<button type="button" class="btn" data-close>إلغاء</button><button type="submit" form="vp-form" class="btn primary">تغيير</button>');
    $('#vp-form', bg).addEventListener('submit', async (ev) => {
      ev.preventDefault();
      const p1 = $('#vp-1', bg).value, p2 = $('#vp-2', bg).value, err = $('#vp-err', bg);
      if (p1.length < 12) return (err.textContent = 'كلمة المرور يجب ألا تقل عن 12 حرفًا.');
      if (p1 !== p2) return (err.textContent = 'كلمتا المرور غير متطابقتين.');
      const raw = new Uint8Array(await crypto.subtle.exportKey('raw', V.key));
      const doc = await wrapVault(p1, raw);
      if (await write(() => S.db.doc('vault/main').set(doc))) { closeModal(); toast('تم تغيير كلمة المرور الرئيسية'); }
    });
  }
  function vaultSettingsSection() {
    if (!V.doc) return '';
    return '<section class="card"><header><h2>خزنة بيانات الدخول</h2></header><div class="body"><div class="toolbar"><span>' + (V.key ? 'مفتوحة الآن' : 'مقفلة') + ' · ' + int(V.creds.size) + ' عميل له بيانات محفوظة</span><span class="grow"></span>' +
      (writable() && V.key ? '<button type="button" class="btn" data-act="vault-pass">تغيير كلمة المرور الرئيسية</button>' : '') + (V.key ? '<button type="button" class="btn" data-act="vault-lock">🔒 قفل</button>' : '') + '</div></div></section>';
  }
  function subscribeVault(db) {
    db.doc('vault/main').onSnapshot(snap => { V.doc = snap.exists ? clean(snap.data()) : null; V.docLoaded = true; scheduleRender(); }, () => { V.docLoaded = true; scheduleRender(); });
    db.collection('credentials').onSnapshot(snap => {
      V.creds = new Map(snap.docs.map(d => [d.id, clean(d.data())])); V.credsLoaded = true;
      if (V.key) decryptAll(); else scheduleRender();
    }, () => { V.credsLoaded = true; scheduleRender(); });
  }

  /* ---------------- income certificates (شهادات الدخل) ----------------
     Same wording and layout as the office's existing certificate: logo top-left, green corner shapes,
     bold body lines, signature block, contact footer. Figures are written in words automatically (editable). */
  const CERT_DEFAULTS = {
    officeLine: 'مكتب محاسب قانوني محمد جمال – س.م.م 33395', signerTitle: 'محاسب قانوني', signerName: 'محمد جمال سليم',
    phones: '+2 01061947848   +2 01111869998', email: 'Mohamed.gamal5851@gmail.com',
    addressEn: '11th District - Building 64, in front of the City of Culture and Science, 6th of October City, Giza', color: '#00be62',
    capOffice: 'محمد جمال', registryNo: '33395'
  };
  const certSettings = () => Object.assign({}, CERT_DEFAULTS, settings().cert || {});

  /* Arabic amount in words, office style: «فقط ثمانية آلاف جنيها مصريا لاغير». */
  const W_ONES = ['', 'واحد', 'اثنان', 'ثلاثة', 'أربعة', 'خمسة', 'ستة', 'سبعة', 'ثمانية', 'تسعة', 'عشرة', 'أحد عشر', 'اثنا عشر', 'ثلاثة عشر', 'أربعة عشر', 'خمسة عشر', 'ستة عشر', 'سبعة عشر', 'ثمانية عشر', 'تسعة عشر'];
  const W_TENS = ['', '', 'عشرون', 'ثلاثون', 'أربعون', 'خمسون', 'ستون', 'سبعون', 'ثمانون', 'تسعون'];
  const W_HUND = ['', 'مائة', 'مائتان', 'ثلاثمائة', 'أربعمائة', 'خمسمائة', 'ستمائة', 'سبعمائة', 'ثمانمائة', 'تسعمائة'];
  function words999(n) {
    const h = Math.floor(n / 100), r = n % 100, parts = [];
    if (h) parts.push(W_HUND[h]);
    if (r) { if (r < 20) parts.push(W_ONES[r]); else { const u = r % 10, t = Math.floor(r / 10); parts.push(u ? W_ONES[u] + ' و' + W_TENS[t] : W_TENS[t]); } }
    return parts.join(' و');
  }
  function scale(n, one, two, plural, single) {
    if (n === 1) return one; if (n === 2) return two;
    if (n >= 3 && n <= 10) return words999(n) + ' ' + plural;
    return words999(n) + ' ' + single;
  }
  function numberWords(n) {
    n = Math.floor(Math.abs(Number(n) || 0));
    if (!n) return 'صفر';
    const parts = [];
    const bil = Math.floor(n / 1e9), mil = Math.floor(n / 1e6) % 1000, th = Math.floor(n / 1000) % 1000, rest = n % 1000;
    if (bil) parts.push(scale(bil, 'مليار', 'ملياران', 'مليارات', 'مليار'));
    if (mil) parts.push(scale(mil, 'مليون', 'مليونان', 'ملايين', 'مليون'));
    if (th) parts.push(scale(th, 'ألف', 'ألفان', 'آلاف', 'ألف'));
    if (rest) parts.push(words999(rest));
    return parts.join(' و');
  }
  function amountWords(v) {
    const n = Math.round((Number(v) || 0) * 100), pounds = Math.floor(n / 100), pt = n % 100;
    return 'فقط ' + numberWords(pounds) + ' جنيها مصريا' + (pt ? ' و' + numberWords(pt) + ' قرشا' : '') + ' لاغير';
  }
  const arDate = (iso) => { if (!iso) return ''; const [y, m, d] = iso.split('-').map(Number); return d + ' ' + MONTHS[m - 1] + ' ' + y; };
  const numFmt = (v) => { const n = Number(v) || 0; return Number.isInteger(n) ? String(n) : n.toFixed(2); };

  const certTemplate = () => certSettings().template || 'classic';
  const isCapital = (c) => c && c.type === 'capital';
  function certBody(c) { if (isCapital(c)) return certBodyClassic(c); const t = certTemplate(); return t === 'modern' ? certBodyModern(c) : t === 'original' ? certBodyOriginal(c) : certBodyClassic(c); }
  const digits = (t) => esc(t).replace(/(\d[\d,./-]*)/g, '<span class="cn">$1</span>');
  const startText = (v) => (/^\d{4}-\d{2}-\d{2}$/.test(v || '') ? arDate(v) : (v || ''));

  /* Modern layout: letterhead, reference line, details table, income boxes, signature and stamp, contact band. */
  function certBodyModern(c) {
    const s = certSettings(), f = c.gender === 'f';
    const row = (k, v) => v ? '<tr><th>' + k + '</th><td>' + v + '</td></tr>' : '';
    const to = (c.submitTo || '').trim();
    return '<div class="cert-page cm" style="--cg:' + esc(s.color) + '">' +
      (window.OF_LOGO ? '<img class="cm-mark" src="' + window.OF_LOGO.icon + '" alt="">' : '') +
      '<header class="cm-head">' + (window.OF_LOGO ? '<img class="cm-logo" src="' + window.OF_LOGO.dark + '" alt="">' : '<span></span>') +
      '<div class="cm-office"><b>' + digits(s.officeLine) + '</b><span>' + esc(s.signerTitle) + ' · ' + esc(s.signerName) + '</span></div></header>' +
      '<div class="cm-rule"></div>' +
      '<div class="cm-ref"><span></span><span>التاريخ: <b>' + digits(arDate(c.date)) + '</b></span></div>' +
      '<div class="cm-title"><h1>شهادة دخل</h1><small>INCOME CERTIFICATE</small></div>' +
      (to ? '<p class="cm-to">السادة / <b>' + esc(to) + '</b> &nbsp;المحترمين،&nbsp; <span class="cm-greet">تحية طيبة وبعد،</span></p>' : '') +
      '<p class="cm-lead">نشهد نحن / <b>' + digits(s.officeLine) + '</b> ' + (f ? 'بأن السيدة' : 'بأن السيد') + ' المبينة بياناته أدناه:</p>'.replace('المبينة بياناته', f ? 'المبينة بياناتها' : 'المبينة بياناته') +
      '<table class="cm-table">' + row('الاسم', '<b>' + esc(c.name) + '</b>') + row('الرقم القومي', '<span class="cn">' + esc(c.nationalId) + '</span>') + row('محل الإقامة', c.address ? digits(c.address) : '') +
      row('الوظيفة', esc(c.job)) + row(f ? 'تعمل بها منذ' : 'يعمل بها منذ', c.startDate ? digits(startText(c.startDate)) : '') + '</table>' +
      '<p class="cm-lead">' + (f ? 'وأن صافي دخلها' : 'وأن صافي دخله') + ' كالتالي:</p>' +
      '<div class="cm-income"><div><small>صافي الدخل الشهري</small><b><span class="cn">' + esc(money(c.monthly)) + '</span> جنيه</b><em>' + esc(c.monthlyWords || amountWords(c.monthly)) + '</em></div>' +
      '<div><small>صافي الدخل السنوي</small><b><span class="cn">' + esc(money(c.annual)) + '</span> جنيه</b><em>' + esc(c.annualWords || amountWords(c.annual)) + '</em></div></div>' +
      '<p class="cm-body">وذلك وفقًا لما تقدم إلينا من مستندات. وهذه شهادة منا بذلك، حُررت بناءً على ' + (f ? 'طلبها' : 'طلبه') + ' لتقديمها إلى ' + (to ? esc(to) : 'من يهمه الأمر') + '، دون أدنى مسؤولية على المكتب.</p>' +
      '<p class="cm-close">وتفضلوا بقبول فائق الاحترام والتقدير،</p>' +
      '<div class="cm-sign"><div class="cm-stamp">ختم المكتب</div><div class="cm-signer"><small>' + esc(s.signerTitle) + '</small><b>' + esc(s.signerName) + '</b><i></i><small>التوقيع</small></div></div>' +
      '<footer class="cm-foot"><span><i>' + CERT_ICON.phone + '</i>' + esc(s.phones) + '</span><span><i>' + CERT_ICON.mail + '</i>' + esc(s.email) + '</span><span class="w"><i>' + CERT_ICON.web + '</i>' + esc(s.addressEn) + '</span></footer>' +
      '</div>';
  }

  /* The certificate wording (ديباجة) exactly as the office writes it; shared by the classic and original layouts. */
  function certWording(c) {
    const s = certSettings(), f = c.gender === 'f';
    const to = (c.submitTo || '').trim();
    const L = (label, value) => '<p><span class="lb">' + label + '</span> ' + value + '</p>';
    const num = (v) => '<span class="cn">' + esc(v) + '</span>';
    return L('نشهد نحن /', digits(s.officeLine)) +
      L(f ? 'بأن السيدة /' : 'بأن السيد /', esc(c.name)) +
      L('رقم قومي:', num(c.nationalId)) +
      (c.address ? L(f ? 'المقيمة في :' : 'المقيم في :', digits(c.address)) : '') +
      L(f ? 'تعمل بوظيفة:' : 'يعمل بوظيفة:', esc(c.job) + (c.startDate ? ' منذ ' + digits(startText(c.startDate)) : '')) +
      L(f ? 'وأن صافى دخلها الشهري:' : 'وأن صافى دخله الشهري:', num(numFmt(c.monthly)) + ' جنيه (' + esc(c.monthlyWords || amountWords(c.monthly)) + ')') +
      L(f ? 'و أن صافي دخلها السنوي :' : 'و أن صافي دخله السنوي :', num(numFmt(c.annual)) + ' جنيه ( ' + esc(c.annualWords || amountWords(c.annual)) + ') وذلك وفقا لما تقدم الينا من مستندات') +
      '<p>وهذه شهاده منا بذلك تم تحريرها بناء علي ' + (f ? 'طلبها' : 'طلبه') + ' لتقديمها الي ' + (to ? esc(to) : 'من يهمه الامر') + ' وذلك دون ادني مسئولية علينا.</p>';
  }

  /* Classic layout: framed page, centred letterhead, Ruqaa title, Naskh text, the original wording unchanged. */
  function certBodyClassic(c) {
    const s = certSettings(), spec = classicSpec(c);
    const orn = '<svg class="cc-orn" viewBox="0 0 200 12" aria-hidden="true"><path d="M0 6 H84"/><path d="M116 6 H200"/><path d="M100 0 L106 6 L100 12 L94 6 Z" class="f"/><circle cx="88" cy="6" r="1.6" class="f"/><circle cx="112" cy="6" r="1.6" class="f"/></svg>';
    const corner = '<svg viewBox="0 0 40 40" aria-hidden="true"><path d="M2 38 V2 H38" /><path d="M8 38 V8 H38" /><path d="M14 14 l5 0 l-5 5 z" class="f"/></svg>';
    return '<div class="cert-page cc" style="--cg:' + esc(s.color) + '">' +
      '<div class="cc-frame"></div>' +
      ['tr', 'tl', 'br', 'bl'].map(k => '<i class="cc-corner ' + k + '">' + corner + '</i>').join('') +
      (window.OF_LOGO ? '<img class="cc-mark" src="' + window.OF_LOGO.icon + '" alt="">' : '') +
      '<header class="cc-head">' + (window.OF_LOGO ? '<img class="cc-logo" src="' + window.OF_LOGO.dark + '" alt="">' : '') + '</header>' +
      '<div class="cc-title">' + orn + '<h1>' + esc(spec.title) + '</h1>' + orn + '</div>' +
      '<div class="cc-text">' + spec.paras.map(t => '<p>' + digits(t) + '</p>').join('') + (spec.close ? '<p class="cc-close">' + esc(spec.close) + '</p>' : '') +
      '<div class="cc-sign"><div class="cc-date">' + (spec.noDate ? '' : 'تحريرا في ' + digits(arDate(c.date))) + '</div><div class="cc-signer">' + spec.signer.map((t, i) => i === spec.signer.length - 1 ? '<b>' + esc(t) + '</b>' : '<div>' + esc(t) + '</div>').join('') + '<i></i></div></div></div>' +
      '<footer class="cc-foot">' + orn + '<div class="cc-contacts"><span><i>' + CERT_ICON.phone + '</i>' + esc(s.phones) + '</span><span><i>' + CERT_ICON.mail + '</i>' + esc(s.email) + '</span></div>' +
      '<div class="cc-contacts"><span><i>' + CERT_ICON.web + '</i>' + esc(s.addressEn) + '</span></div></footer>' +
      '</div>';
  }

  /* Original layout: the office's first certificate, reproduced as-is. */
  function certBodyOriginal(c) {
    const s = certSettings();
    return '<div class="cert-page" style="--cg:' + esc(s.color) + '">' +
      '<svg class="cert-shape" viewBox="0 0 827 1170" preserveAspectRatio="none" aria-hidden="true"><path d="M664 22 C 680 95, 760 185, 805 248 L 805 22 Z"/><path d="M10 1055 C 30 1075, 65 1110, 68 1146 L 815 1146 L 815 1162 L 10 1162 Z"/></svg>' +
      (window.OF_LOGO ? '<img class="cert-logo" src="' + window.OF_LOGO.dark + '" alt="">' : '') +
      '<h1 class="cert-title">شهادة دخل</h1>' +
      '<div class="cert-text">' + certWording(c) +
      '<p class="cert-close">وتفضلوا بقبول فائق الاحترام والتقدير،</p>' +
      '<div class="cert-sign"><div class="cert-date">تحريرا في ' + digits(arDate(c.date)) + '</div><div class="cert-signer"><div>' + esc(s.signerTitle) + '</div><div>' + esc(s.signerName) + '</div></div></div>' +
      '</div>' +
      '<div class="cert-foot">' +
      '<div><i>' + CERT_ICON.phone + '</i><span>' + esc(s.phones) + '</span></div>' +
      '<div><i>' + CERT_ICON.mail + '</i><span>' + esc(s.email) + '</span></div>' +
      '<div><i>' + CERT_ICON.web + '</i><span>' + esc(s.addressEn) + '</span></div></div>' +
      '</div>';
  }
  const CERT_ICON = {
    phone: '<svg viewBox="0 0 24 24"><path d="M8 4.5l2 3.5-1.6 1.4a9 9 0 0 0 6.2 6.2l1.4-1.6 3.5 2-.8 2.6c-.3.8-1.1 1.2-1.9 1A15 15 0 0 1 4.4 7.2c-.2-.8.2-1.6 1-1.9z"/></svg>',
    mail: '<svg viewBox="0 0 24 24"><rect x="4" y="6.5" width="16" height="11" rx="1.5"/><path d="M4.5 7.5l7.5 6 7.5-6"/></svg>',
    web: '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="8"/><path d="M4 12h16M12 4c2.5 2.5 2.5 13.5 0 16M12 4c-2.5 2.5-2.5 13.5 0 16"/></svg>'
  };

  /* Long addresses or entity names can push the signature into the footer: shrink the text a little until it fits. */
  function fitCert(page) {
    if (!page) return;
    const cls = page.classList.contains('cm') ? 'cm' : page.classList.contains('cc') ? 'cc' : '';
    if (!cls) return;
    const sign = page.querySelector('.' + cls + '-sign'), foot = page.querySelector('.' + cls + '-foot');
    let fs = parseFloat(getComputedStyle(page).fontSize) * 0.75;
    while (fs > 9 && sign.getBoundingClientRect().bottom > foot.getBoundingClientRect().top - 4) { fs -= 0.25; page.style.fontSize = fs + 'pt'; }
  }
  /* The classic certificate is drawn straight onto a canvas for the PDF. html2canvas places Arabic text word by word
     and loses the right-to-left order of slashes, colons and brackets; the browser's own canvas text drawing with
     direction "rtl" shapes and orders each line correctly. Layout mirrors the .cc styles used for the on-screen preview. */
  const loadImg = (src) => new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = src; });
  const svgUri = (inner, color) => 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="' + color + '" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">' + inner + '</svg>');
  const ICON_PATHS = {
    phone: '<path d="M8 4.5l2 3.5-1.6 1.4a9 9 0 0 0 6.2 6.2l1.4-1.6 3.5 2-.8 2.6c-.3.8-1.1 1.2-1.9 1A15 15 0 0 1 4.4 7.2c-.2-.8.2-1.6 1-1.9z"/>',
    mail: '<rect x="4" y="6.5" width="16" height="11" rx="1.5"/><path d="M4.5 7.5l7.5 6 7.5-6"/>',
    web: '<circle cx="12" cy="12" r="8"/><path d="M4 12h16M12 4c2.5 2.5 2.5 13.5 0 16M12 4c-2.5 2.5-2.5 13.5 0 16"/>'
  };
  function classicParagraphs(c) {
    const s = certSettings(), f = c.gender === 'f', to = (c.submitTo || '').trim();
    return [
      'نشهد نحن / ' + s.officeLine,
      (f ? 'بأن السيدة / ' : 'بأن السيد / ') + c.name,
      'رقم قومي: ' + c.nationalId,
      c.address ? (f ? 'المقيمة في : ' : 'المقيم في : ') + c.address : null,
      (f ? 'تعمل بوظيفة: ' : 'يعمل بوظيفة: ') + c.job + (c.startDate ? ' منذ ' + startText(c.startDate) : ''),
      (f ? 'وأن صافى دخلها الشهري: ' : 'وأن صافى دخله الشهري: ') + numFmt(c.monthly) + ' جنيه (' + (c.monthlyWords || amountWords(c.monthly)) + ')',
      (f ? 'و أن صافي دخلها السنوي : ' : 'و أن صافي دخله السنوي : ') + numFmt(c.annual) + ' جنيه ( ' + (c.annualWords || amountWords(c.annual)) + ') وذلك وفقا لما تقدم الينا من مستندات',
      'وهذه شهاده منا بذلك تم تحريرها بناء علي ' + (f ? 'طلبها' : 'طلبه') + ' لتقديمها الي ' + (to || 'من يهمه الامر') + ' وذلك دون ادني مسئولية علينا.'
    ].filter(Boolean);
  }
  const CHAMBERS = ['غرفة صناعة تكنولوجيا المعلومات والاتصالات', 'الغرفة التجارية بالجيزة', 'الغرفة التجارية بالقاهرة', 'غرفة الصناعات الغذائية', 'غرفة الصناعات الهندسية', 'غرفة صناعة مواد البناء'];
  const chamberList = () => { const l = certSettings().chambers; return Array.isArray(l) && l.length ? l : CHAMBERS; };
  const thousands = (v) => { const n = Number(v) || 0; return n.toLocaleString('en-US', { maximumFractionDigits: 2 }); };
  const dmy = (iso) => (/^\d{4}-\d{2}-\d{2}$/.test(iso || '') ? iso.split('-').reverse().join('-') : (iso || ''));
  function capitalParagraphs(c) {
    const s = certSettings();
    return [
      'السادة ' + (c.chamber || ''),
      'يشهد مكتب محاسب قانوني : ' + (s.capOffice || '') + ' مقيد بالسجل العام للمحاسبين و المراجعين تحت رقم ' + (s.registryNo || ''),
      'بأن شركة : ' + (c.company || '') + (c.hq ? ' مقرها ' + c.hq : ''),
      'و مقيدة تحت سجل رقم ' + (c.crNo || '') + (c.crOffice ? ' مكتب ' + c.crOffice : ''),
      'بأن رأس مالها المستثمر هو ' + thousands(c.capital) + ' جنيه ( ' + (c.capitalWords || amountWords(c.capital)) + ' )',
      'من واقع أخر ميزانية عن العام المالي المنتهى فى ' + '\u200E' + dmy(c.fyEnd) + '\u200E',  // LRM keeps 31-12-2025 in day-month-year order inside Arabic text
    ];
  }
  /* What the classic layout prints for each kind of certificate. */
  function classicSpec(c) {
    const s = certSettings();
    if (isCapital(c)) return { title: 'شهادة', paras: capitalParagraphs(c), close: 'وهذه شهادة منا بذلك', noDate: true, signer: ['يعتمد ،،', s.signerTitle || 'محاسب قانوني', s.signerName || ''] };
    return { title: 'شهادة دخل', paras: classicParagraphs(c), close: 'وتفضلوا بقبول فائق الاحترام والتقدير،', signer: [s.signerTitle || '', s.signerName || ''] };
  }
  async function drawClassicCert(c) {
    const spec = classicSpec(c);
    const s = certSettings(), color = s.color || CERT_DEFAULTS.color;
    const W = 1654, H = Math.floor(W * 297 / 210), K = W / 210, PT = 0.3528 * K; // px per mm, px per pt
    const cv = document.createElement('canvas'); cv.width = W; cv.height = H;
    const g = cv.getContext('2d');
    g.fillStyle = '#fff'; g.fillRect(0, 0, W, H);
    const mm = (v) => v * K;
    // frame
    g.strokeStyle = '#1c1c1c'; g.lineWidth = mm(0.9); g.strokeRect(mm(9.45), mm(9.45), W - mm(18.9), H - mm(18.9));
    g.lineWidth = mm(0.3); g.strokeRect(mm(10.95), mm(10.95), W - mm(21.9), H - mm(21.9));
    // corners
    const corner = (x, y, sx, sy) => {
      g.save(); g.translate(x, y); g.scale(sx * mm(13) / 40, sy * mm(13) / 40);
      g.strokeStyle = color; g.lineWidth = 1.6; g.beginPath(); g.moveTo(2, 38); g.lineTo(2, 2); g.lineTo(38, 2); g.moveTo(8, 38); g.lineTo(8, 8); g.lineTo(38, 8); g.stroke();
      g.fillStyle = color; g.beginPath(); g.moveTo(14, 14); g.lineTo(19, 14); g.lineTo(14, 19); g.closePath(); g.fill(); g.restore();
    };
    corner(mm(12.5), mm(12.5), 1, 1); corner(W - mm(12.5), mm(12.5), -1, 1); corner(mm(12.5), H - mm(12.5), 1, -1); corner(W - mm(12.5), H - mm(12.5), -1, -1);
    // watermark and logo
    if (window.OF_LOGO) {
      const [icon, logo] = await Promise.all([loadImg(window.OF_LOGO.icon), loadImg(window.OF_LOGO.dark)]);
      g.globalAlpha = 0.04; g.drawImage(icon, mm(47.5), mm(100), mm(115), mm(115)); g.globalAlpha = 1;
      const lw = mm(54), lh = lw * logo.height / logo.width; g.drawImage(logo, (W - lw) / 2, mm(20), lw, lh);
    }
    // ornament: line — dot — diamond — dot — line
    const ornament = (cx, cy, width) => {
      const u = width / 200;
      g.strokeStyle = '#1c1c1c'; g.lineWidth = Math.max(1, mm(0.25)); g.beginPath(); g.moveTo(cx - 100 * u, cy); g.lineTo(cx - 16 * u, cy); g.moveTo(cx + 16 * u, cy); g.lineTo(cx + 100 * u, cy); g.stroke();
      g.fillStyle = color; g.beginPath(); g.moveTo(cx, cy - 6 * u); g.lineTo(cx + 6 * u, cy); g.lineTo(cx, cy + 6 * u); g.lineTo(cx - 6 * u, cy); g.closePath(); g.fill();
      [-12, 12].forEach(d => { g.beginPath(); g.arc(cx + d * u, cy, 1.6 * u, 0, Math.PI * 2); g.fill(); });
    };
    g.direction = 'rtl'; g.textBaseline = 'middle'; g.fillStyle = '#1a1a1a';
    // title
    g.font = '700 ' + (30 * PT) + 'px CertSerif'; g.textAlign = 'center';
    const ty = mm(57) + 30 * PT * 0.62;
    g.fillText(spec.title, W / 2, ty);
    const tw = g.measureText(spec.title).width, ow = mm(34);
    ornament(W / 2 + tw / 2 + mm(5) + ow / 2, ty, ow); ornament(W / 2 - tw / 2 - mm(5) - ow / 2, ty, ow);
    // body, shrinking the type a little if long content would reach the footer
    const paras = spec.paras;
    const right = W - mm(25), boxW = W - mm(50), top = mm(80), limit = H - mm(15) - mm(17);
    let fs = 15, lines, lh;
    const wrap = (text) => {
      const words = text.split(/\s+/).filter(Boolean), out = []; let cur = '';
      for (const w of words) { const t = cur ? cur + ' ' + w : w; if (g.measureText(t).width <= boxW || !cur) cur = t; else { out.push(cur); cur = w; } }
      if (cur) out.push(cur); return out;
    };
    for (;;) {
      g.font = '700 ' + (fs * PT) + 'px CertSerif'; lh = fs * PT * 2.05;
      lines = []; paras.forEach(p => lines.push(...wrap(p)));
      const bodyEnd = top + lines.length * lh + (spec.close ? mm(3) + lh : 0) + mm(4) + fs * PT * 1.7 * spec.signer.length + mm(9);
      if (bodyEnd <= limit || fs <= 10) break;
      fs -= 0.25;
    }
    g.textAlign = 'right'; g.fillStyle = '#1a1a1a';
    let y = top + lh / 2;
    for (const l of lines) { g.fillText(l, right, y); y += lh; }
    y -= lh / 2;
    if (spec.close) { y += mm(3); g.textAlign = 'center'; g.fillText(spec.close, mm(25) + boxW / 2, y + lh / 2); y += lh; }
    y += mm(4);
    const sl = fs * PT * 1.7;
    if (!spec.noDate) { g.textAlign = 'right'; g.fillText('تحريرا في ' + arDate(c.date), right, y + sl / 2); }
    const sx = mm(25) + mm(4) + mm(27.5);
    g.textAlign = 'center';
    spec.signer.forEach((t, i) => { const last = i === spec.signer.length - 1; g.font = '700 ' + (fs * (last ? 1.033 : 1) * PT) + 'px CertSerif'; g.fillText(t, sx, y + sl * (i + 0.5)); });
    const ly = y + sl * spec.signer.length + mm(9);
    g.strokeStyle = '#777'; g.lineWidth = Math.max(1, mm(0.3)); g.setLineDash([mm(0.4), mm(0.6)]);
    g.beginPath(); g.moveTo(sx - mm(24), ly); g.lineTo(sx + mm(24), ly); g.stroke(); g.setLineDash([]);
    // footer: ornament, phones + email, address (left to right)
    g.direction = 'ltr';
    const fpx = 8.8 * PT, line2 = H - mm(15) - fpx * 0.7, line1 = line2 - fpx * 1.35 - mm(1.2);
    ornament(W / 2, line1 - fpx * 0.7 - mm(2.2), mm(60));
    g.font = '400 ' + fpx + 'px Roboto, Arial, sans-serif'; g.fillStyle = '#444'; g.textAlign = 'left';
    const icons = await Promise.all(['phone', 'mail', 'web'].map(k => loadImg(svgUri(ICON_PATHS[k], color))));
    const isz = mm(3.4), igap = mm(1.8), sep = mm(8);
    const drawRow = (items, cy) => {
      const widths = items.map(([, t]) => isz + igap + g.measureText(t).width);
      let x = (W - (widths.reduce((a, b) => a + b, 0) + sep * (items.length - 1))) / 2;
      items.forEach(([ic, t], i) => { g.drawImage(ic, x, cy - isz / 2, isz, isz); g.fillText(t, x + isz + igap, cy); x += widths[i] + sep; });
    };
    drawRow([[icons[0], s.phones || ''], [icons[1], s.email || '']].filter(r => r[1]), line1);
    if (s.addressEn) drawRow([[icons[2], s.addressEn]], line2);
    return cv;
  }

  async function pdfCertificate(c) {
    if (!window.html2pdf) return toast('يتطلب تنزيل PDF اتصالًا بالإنترنت عند فتح الصفحة. اتصل ثم أعد تحميل الصفحة.', true);
    toast('جارٍ إعداد الشهادة…');
    if (document.fonts) {
      // make sure the certificate fonts are really loaded before anything is drawn
      await Promise.all(['400 15pt CertSerif', '700 15pt CertSerif', '400 10pt Roboto', '400 12pt "IBM Plex Sans Arabic"', '700 12pt "IBM Plex Sans Arabic"'].map(f => document.fonts.load(f, 'شهادة دخل 123 abc').catch(() => null)));
      await document.fonts.ready;
    }
    const opts = { margin: 0, image: { type: 'jpeg', quality: 0.97 }, jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }, pagebreak: { mode: [] } };
    const name = fileName((isCapital(c) ? 'شهادة رأس مال ' + (c.company || '') : 'شهادة دخل ' + c.name) + ' ' + (c.date || '')) + '.pdf';
    if (isCapital(c) || certTemplate() === 'classic') {
      try { const cv = await drawClassicCert(c); await saveFile(name, await html2pdf().set(opts).from(cv, 'canvas').outputPdf('blob')); }
      catch (e) { toast((e && e.message) || 'تعذّر إعداد الشهادة.', true); }
      return;
    }
    const holder = document.createElement('div'); holder.style.cssText = 'position:fixed;left:-20000px;top:0;';
    holder.innerHTML = certBody(c); document.body.appendChild(holder);
    const page = holder.firstChild;
    try {
      fitCert(page);
      const blob = await html2pdf().set(Object.assign(opts, { html2canvas: { scale: 2.5, backgroundColor: '#ffffff', useCORS: true, scrollX: 0, scrollY: 0 } })).from(page).outputPdf('blob');
      await saveFile(name, blob);
    } catch (e) { toast((e && e.message) || 'تعذّر إعداد الشهادة.', true); }
    finally { holder.remove(); }
  }

  function filteredCerts() {
    const q = (S.ui.ceQ || '').trim().toLowerCase();
    return (S.raw.certificates || []).filter(c => !q || [c.name, c.nationalId, c.job, c.address, c.company, c.chamber, c.crNo, c.submitTo].some(v => String(v || '').toLowerCase().includes(q)))
      .sort((a, b) => (b.date || '').localeCompare(a.date || '') || String(b.createdAt || '').localeCompare(String(a.createdAt || '')));
  }
  const certKind = (c) => (isCapital(c) ? 'رأس مال (للغرف)' : 'شهادة دخل');
  const certWho = (c) => (isCapital(c) ? c.company : c.name) || '';
  const certTo = (c) => (isCapital(c) ? c.chamber : c.submitTo) || '';
  function pageCerts() {
    const list = filteredCerts();
    const fees = list.reduce((s, c) => s + (Number(c.fee) || 0), 0), unpaid = list.filter(c => Number(c.fee) > 0 && !c.paid).reduce((s, c) => s + (Number(c.fee) || 0), 0);
    const actions = actBtn('add-cert', '+ شهادة دخل', 'primary') + actBtn('add-capcert', '+ شهادة رأس مال للغرف', 'primary') + expBtns('certs');
    let h = topBar('الشهادات', int(list.length) + ' شهادة · أتعاب ' + esc(moneyC(fees)) + (unpaid ? ' · غير محصل ' + esc(moneyC(unpaid)) : ''), actions) + '<div class="content">';
    h += '<div class="toolbar"><input type="search" placeholder="بحث بالاسم أو الشركة أو الرقم القومي أو الغرفة…" data-ui="ceQ" data-live value="' + esc(S.ui.ceQ || '') + '"></div>';
    h += '<div class="card tablewrap"><table><thead><tr><th>التاريخ</th><th>النوع</th><th>الاسم / الشركة</th><th>مقدمة إلى</th><th class="n">المبلغ</th><th class="n">الأتعاب</th><th></th></tr></thead><tbody>';
    if (!list.length) h += '<tr><td colspan="7" class="empty">لا توجد شهادات بعد.</td></tr>';
    for (const c of list) {
      const cap = isCapital(c);
      h += '<tr><td class="num">' + esc(c.date || '') + '</td><td><span class="pill ' + (cap ? 'info' : 'acc') + '">' + esc(certKind(c)) + '</span></td>' +
        '<td><div class="stack"><strong>' + esc(certWho(c)) + '</strong><span class="sub">' + esc(cap ? (c.crNo ? 'سجل رقم ' + c.crNo : '') : (c.job || '')) + '</span>' + (c.client ? '<a class="sub" href="#client-statement-' + esc(c.client) + '">' + esc(clientName(c.client)) + '</a>' : '') + '</div></td>' +
        '<td>' + esc(certTo(c) || (cap ? '' : 'من يهمه الأمر')) + '</td>' +
        '<td class="n">' + (cap ? money(c.capital) + '<div class="sub">رأس المال المستثمر</div>' : money(c.monthly) + '<div class="sub">الدخل الشهري</div>') + '</td>' +
        '<td class="n">' + (Number(c.fee) ? money(c.fee) + ' ' + (c.paid ? '<span class="pill ok">محصل</span>' : '<span class="pill warn">غير محصل</span>') : '') + '</td>' +
        '<td style="white-space:nowrap">' + (S.downloads ? '<button type="button" class="btn sm primary" data-act="cert-pdf" data-id="' + c.id + '">⤓ PDF</button> ' : '') + '<button type="button" class="btn sm" data-act="cert-preview" data-id="' + c.id + '">معاينة</button> ' +
        actBtn('edit-cert', 'تعديل', 'sm', ' data-id="' + c.id + '"') + ' ' + actBtn('copy-cert', 'نسخة جديدة', 'sm', ' data-id="' + c.id + '"') + '</td></tr>';
    }
    return h + '</tbody></table></div></div>';
  }
  function certReport() {
    const list = filteredCerts();
    return { title: 'سجل الشهادات', ctx: int(list.length) + ' شهادة', landscape: true,
      cols: [col('التاريخ'), col('النوع'), col('الاسم / الشركة'), col('الرقم القومي / رقم السجل'), col('مقدمة إلى'), col('الوظيفة / المقر'), col('المبلغ', 1), col('الأتعاب', 1), col('التحصيل')],
      rows: list.map(c => [c.date, certKind(c), certWho(c), isCapital(c) ? (c.crNo || '') : (c.nationalId || ''), certTo(c), isCapital(c) ? (c.hq || '') : (c.job || ''), Number(isCapital(c) ? c.capital : c.monthly) || 0, Number(c.fee) || '', Number(c.fee) ? (c.paid ? 'محصل' : 'غير محصل') : '']),
      foot: ['الإجمالي', '', '', '', '', '', '', list.reduce((s, c) => s + (Number(c.fee) || 0), 0), ''] };
  }
  function previewCert(c) {
    const bg = modal('معاينة · ' + certKind(c) + ' · ' + certWho(c), '<div class="cert-preview"><div class="cert-scale">' + certBody(c) + '</div></div>',
      '<button type="button" class="btn" data-close>إغلاق</button>' + (S.downloads ? '<button type="button" class="btn primary" id="cp-pdf">⤓ تنزيل PDF</button>' : ''), true);
    fitCert($('.cert-page', bg));
    const b = $('#cp-pdf', bg); if (b) b.addEventListener('click', () => pdfCertificate(c));
  }

  function openCertForm(cert, copyFrom) {
    const isNew = !cert;
    const base = cert ? clean(cert) : Object.assign({ gender: 'm', name: '', nationalId: '', address: '', job: '', startDate: '', submitTo: '', monthly: '', annual: '', monthlyWords: '', annualWords: '', client: '', fee: '', paid: false, notes: '' },
      copyFrom ? clean(copyFrom) : {}, { id: uid(), date: C.todayISO(), createdAt: new Date().toISOString() });
    const c = base;
    const body = '<form id="ce-form" class="form">' +
      fld('ce-client', 'ربط بعميل (اختياري)', '<select id="ce-client">' + options(clientOptions(true), c.client, '—') + '</select>') +
      fld('ce-gender', 'السيد / السيدة', '<select id="ce-gender">' + options([['m', 'السيد'], ['f', 'السيدة']], c.gender) + '</select>') +
      fld('ce-name', 'الاسم', inp('ce-name', c.name, 'text', ' required'), true) +
      fld('ce-nid', 'الرقم القومي (14 رقمًا)', inp('ce-nid', c.nationalId, 'text', ' inputmode="numeric" maxlength="14" dir="ltr" required')) +
      fld('ce-job', 'الوظيفة', inp('ce-job', c.job, 'text', ' required')) +
      fld('ce-start', 'تاريخ بداية العمل (اختياري)', inp('ce-start', c.startDate, 'date')) +
      fld('ce-addr', 'العنوان (اختياري – لا يظهر في الشهادة إذا تُرك فارغًا)', inp('ce-addr', c.address), true) +
      fld('ce-to', 'جهة التقديم (اختياري – وإلا تُكتب «من يهمه الأمر»)', inp('ce-to', c.submitTo, 'text', ' list="ce-to-list" placeholder="مثال: سفارة ألمانيا بالقاهرة"') +
        '<datalist id="ce-to-list">' + Array.from(new Set((S.raw.certificates || []).map(x => x.submitTo).filter(Boolean))).map(v => '<option value="' + esc(v) + '">').join('') + '</datalist>', true) +
      fld('ce-month', 'صافي الدخل الشهري (جنيه)', inp('ce-month', c.monthly, 'number', ' min="0" step="0.01" required')) +
      fld('ce-year', 'صافي الدخل السنوي (جنيه)', inp('ce-year', c.annual, 'number', ' min="0" step="0.01" required')) +
      fld('ce-mw', 'الدخل الشهري بالحروف', inp('ce-mw', c.monthlyWords), true) +
      fld('ce-yw', 'الدخل السنوي بالحروف', inp('ce-yw', c.annualWords), true) +
      fld('ce-date', 'تاريخ التحرير', inp('ce-date', c.date, 'date', ' required')) +
      fld('ce-fee', 'أتعاب الشهادة (اختياري)', inp('ce-fee', c.fee, 'number', ' min="0" step="0.01"')) +
      '<label class="field" style="flex-direction:row;gap:8px;align-items:center;align-self:end"><input type="checkbox" id="ce-paid"' + (c.paid ? ' checked' : '') + '> تم تحصيل الأتعاب</label>' +
      fld('ce-notes', 'ملاحظات داخلية (لا تظهر في الشهادة)', inp('ce-notes', c.notes), true) +
      '<div class="field full err" id="ce-err"></div></form>';
    const foot = (isNew ? '' : '<button type="button" class="btn danger" id="ce-del">حذف</button><span style="flex:1"></span>') + '<button type="button" class="btn" data-close>إلغاء</button>' +
      '<button type="button" class="btn" id="ce-prev">معاينة</button><button type="submit" form="ce-form" class="btn primary">حفظ' + (S.downloads ? ' وتنزيل PDF' : '') + '</button>';
    const bg = modal(isNew ? 'شهادة دخل جديدة' : 'تعديل شهادة دخل · ' + c.name, body, foot, true);
    // words follow the figures until the user types their own wording; annual follows monthly × 12 the same way
    let yearTouched = !!(c.annual && Number(c.annual) !== Number(c.monthly) * 12), mwTouched = !!c.monthlyWords && c.monthlyWords !== amountWords(c.monthly), ywTouched = !!c.annualWords && c.annualWords !== amountWords(c.annual);
    const sync = () => {
      const m = $('#ce-month', bg).value;
      if (!yearTouched && m !== '') $('#ce-year', bg).value = C.round2(Number(m) * 12);
      if (!mwTouched) $('#ce-mw', bg).value = m !== '' ? amountWords(m) : '';
      if (!ywTouched) $('#ce-yw', bg).value = $('#ce-year', bg).value !== '' ? amountWords($('#ce-year', bg).value) : '';
    };
    $('#ce-month', bg).addEventListener('input', sync);
    $('#ce-year', bg).addEventListener('input', () => { yearTouched = true; sync(); });
    $('#ce-mw', bg).addEventListener('input', () => { mwTouched = true; });
    $('#ce-yw', bg).addEventListener('input', () => { ywTouched = true; });
    $('#ce-client', bg).addEventListener('change', () => {
      const cl = S.M.clientById.get(val('ce-client', bg)); if (!cl) return;
      if (!val('ce-name', bg)) $('#ce-name', bg).value = cl.name || '';
      if (!val('ce-addr', bg)) $('#ce-addr', bg).value = cl.address || '';
      if (!val('ce-job', bg) && cl.activity) $('#ce-job', bg).value = cl.activity;
    });
    sync();
    const read = () => Object.assign({}, c, {
      client: val('ce-client', bg), gender: val('ce-gender', bg), name: val('ce-name', bg), nationalId: val('ce-nid', bg).replace(/\s/g, ''), job: val('ce-job', bg), address: val('ce-addr', bg), startDate: val('ce-start', bg), submitTo: val('ce-to', bg),
      monthly: C.round2(val('ce-month', bg)), annual: C.round2(val('ce-year', bg)), monthlyWords: val('ce-mw', bg), annualWords: val('ce-yw', bg),
      date: val('ce-date', bg), fee: val('ce-fee', bg) === '' ? '' : C.round2(val('ce-fee', bg)), paid: $('#ce-paid', bg).checked, notes: val('ce-notes', bg)
    });
    $('#ce-prev', bg).addEventListener('click', () => {
      const n = read(); const w = document.createElement('div'); w.className = 'modal-bg'; w.style.zIndex = 70;
      w.innerHTML = '<div class="modal wide"><header><h2>معاينة</h2><button type="button" class="btn sm" data-x>✕</button></header><div class="body"><div class="cert-preview"><div class="cert-scale">' + certBody(n) + '</div></div></div></div>';
      w.addEventListener('click', e => { if (e.target === w || e.target.closest('[data-x]')) w.remove(); }); document.body.appendChild(w); fitCert($('.cert-page', w));
    });
    $('#ce-form', bg).addEventListener('submit', async (ev) => {
      ev.preventDefault();
      const n = read(), err = $('#ce-err', bg);
      if (!n.name) return (err.textContent = 'اكتب الاسم.');
      if (!/^\d{14}$/.test(n.nationalId)) return (err.textContent = 'الرقم القومي يجب أن يكون 14 رقمًا.');
      if (!n.job) return (err.textContent = 'اكتب الوظيفة.');
      if (!(n.monthly > 0) || !(n.annual > 0)) return (err.textContent = 'أدخل صافي الدخل الشهري والسنوي.');
      if (!n.date) return (err.textContent = 'أدخل تاريخ التحرير.');
      if (!n.monthlyWords) n.monthlyWords = amountWords(n.monthly);
      if (!n.annualWords) n.annualWords = amountWords(n.annual);
      if (await write(() => S.db.doc('certificates/' + n.id).set(clean(n)))) { closeModal(); toast('تم حفظ الشهادة'); if (S.downloads) pdfCertificate(n); }
    });
    const del = $('#ce-del', bg); if (del) del.addEventListener('click', () => confirmBox('حذف الشهادة', 'هل تريد حذف هذه الشهادة من السجل؟', 'حذف', async () => { if (await write(() => S.db.doc('certificates/' + c.id).delete())) { closeModal(); toast('تم الحذف'); } }));
  }

  /* Invested-capital certificate for the chambers (غرف): company data, capital from the latest balance sheet, chamber name. */
  function openCapitalForm(cert, copyFrom) {
    const isNew = !cert;
    const lastYear = (Number(C.todayISO().slice(0, 4)) - 1) + '-12-31';
    const c = cert ? clean(cert) : Object.assign({ chamber: chamberList()[0], company: '', hq: '', crNo: '', crOffice: '', capital: '', capitalWords: '', fyEnd: lastYear, client: '', fee: '', paid: false, notes: '' },
      copyFrom ? clean(copyFrom) : {}, { id: uid(), type: 'capital', date: C.todayISO(), createdAt: new Date().toISOString() });
    const used = Array.from(new Set(chamberList().concat((S.raw.certificates || []).map(x => x.chamber).filter(Boolean))));
    const body = '<form id="cc-form" class="form">' +
      fld('cc-chamber', 'الغرفة المقدمة إليها', inp('cc-chamber', c.chamber, 'text', ' list="cc-ch-list" required') + '<datalist id="cc-ch-list">' + used.map(v => '<option value="' + esc(v) + '">').join('') + '</datalist>', true) +
      fld('cc-client', 'ربط بعميل (يملأ بيانات الشركة)', '<select id="cc-client">' + options(clientOptions(true), c.client, '—') + '</select>') +
      fld('cc-company', 'اسم الشركة', inp('cc-company', c.company, 'text', ' required')) +
      fld('cc-hq', 'المقر', inp('cc-hq', c.hq), true) +
      fld('cc-cr', 'رقم السجل التجاري', inp('cc-cr', c.crNo, 'text', ' required')) +
      fld('cc-cro', 'مكتب السجل (يُكتب بعد كلمة «مكتب»)', inp('cc-cro', c.crOffice, 'text', ' list="cc-cro-list" placeholder="مثال: سجل تجاري الجيزة"') + '<datalist id="cc-cro-list">' + Array.from(new Set((S.raw.certificates || []).map(x => x.crOffice).filter(Boolean))).map(v => '<option value="' + esc(v) + '">').join('') + '</datalist>') +
      fld('cc-fy', 'نهاية العام المالي لآخر ميزانية', inp('cc-fy', c.fyEnd, 'date', ' required')) +
      fld('cc-cap', 'رأس المال المستثمر (جنيه)', inp('cc-cap', c.capital, 'number', ' min="0" step="0.01" required')) +
      fld('cc-words', 'رأس المال بالحروف', inp('cc-words', c.capitalWords), true) +
      fld('cc-date', 'تاريخ الإصدار (للسجل فقط، لا يُطبع)', inp('cc-date', c.date, 'date', ' required')) +
      fld('cc-fee', 'أتعاب الشهادة (اختياري)', inp('cc-fee', c.fee, 'number', ' min="0" step="0.01"')) +
      '<label class="field" style="flex-direction:row;gap:8px;align-items:center;align-self:end"><input type="checkbox" id="cc-paid"' + (c.paid ? ' checked' : '') + '> تم تحصيل الأتعاب</label>' +
      fld('cc-notes', 'ملاحظات داخلية (لا تظهر في الشهادة)', inp('cc-notes', c.notes), true) +
      '<div class="field full err" id="cc-err"></div></form>';
    const foot = (isNew ? '' : '<button type="button" class="btn danger" id="cc-del">حذف</button><span style="flex:1"></span>') + '<button type="button" class="btn" data-close>إلغاء</button>' +
      '<button type="button" class="btn" id="cc-prev">معاينة</button><button type="submit" form="cc-form" class="btn primary">حفظ' + (S.downloads ? ' وتنزيل PDF' : '') + '</button>';
    const bg = modal(isNew ? 'شهادة رأس مال مستثمر للغرف' : 'تعديل شهادة رأس مال · ' + c.company, body, foot, true);
    let wordsTouched = !!c.capitalWords && c.capitalWords !== amountWords(c.capital);
    const sync = () => { if (!wordsTouched) { const v = $('#cc-cap', bg).value; $('#cc-words', bg).value = v !== '' ? amountWords(v) : ''; } };
    $('#cc-cap', bg).addEventListener('input', sync);
    $('#cc-words', bg).addEventListener('input', () => { wordsTouched = true; });
    $('#cc-client', bg).addEventListener('change', () => {
      const cl = S.M.clientById.get(val('cc-client', bg)); if (!cl) return;
      $('#cc-company', bg).value = cl.name || '';
      if (cl.address) $('#cc-hq', bg).value = cl.address;
      const cr = (cl.docs || []).find(d => /السجل التجاري/.test(d.name || '') && d.number); if (cr) $('#cc-cr', bg).value = cr.number;
    });
    sync();
    const read = () => Object.assign({}, c, {
      chamber: val('cc-chamber', bg), client: val('cc-client', bg), company: val('cc-company', bg), hq: val('cc-hq', bg), crNo: val('cc-cr', bg), crOffice: val('cc-cro', bg), fyEnd: val('cc-fy', bg),
      capital: C.round2(val('cc-cap', bg)), capitalWords: val('cc-words', bg), date: val('cc-date', bg), fee: val('cc-fee', bg) === '' ? '' : C.round2(val('cc-fee', bg)), paid: $('#cc-paid', bg).checked, notes: val('cc-notes', bg)
    });
    $('#cc-prev', bg).addEventListener('click', () => {
      const w = document.createElement('div'); w.className = 'modal-bg'; w.style.zIndex = 70;
      w.innerHTML = '<div class="modal wide"><header><h2>معاينة</h2><button type="button" class="btn sm" data-x>✕</button></header><div class="body"><div class="cert-preview"><div class="cert-scale">' + certBody(read()) + '</div></div></div></div>';
      w.addEventListener('click', e => { if (e.target === w || e.target.closest('[data-x]')) w.remove(); }); document.body.appendChild(w); fitCert($('.cert-page', w));
    });
    $('#cc-form', bg).addEventListener('submit', async (ev) => {
      ev.preventDefault();
      const n = read(), err = $('#cc-err', bg);
      if (!n.chamber) return (err.textContent = 'اكتب اسم الغرفة.');
      if (!n.company) return (err.textContent = 'اكتب اسم الشركة.');
      if (!n.crNo) return (err.textContent = 'اكتب رقم السجل التجاري.');
      if (!(n.capital > 0)) return (err.textContent = 'أدخل رأس المال المستثمر.');
      if (!n.fyEnd || !n.date) return (err.textContent = 'أدخل نهاية العام المالي وتاريخ التحرير.');
      if (!n.capitalWords) n.capitalWords = amountWords(n.capital);
      n.name = n.company;
      // a chamber typed for the first time joins the saved list in settings
      if (writable() && !chamberList().includes(n.chamber)) saveSettings({ cert: Object.assign({}, settings().cert || {}, { chambers: chamberList().concat([n.chamber]) }) });
      if (await write(() => S.db.doc('certificates/' + n.id).set(clean(n)))) { closeModal(); toast('تم حفظ الشهادة'); if (S.downloads) pdfCertificate(n); }
    });
    const del = $('#cc-del', bg); if (del) del.addEventListener('click', () => confirmBox('حذف الشهادة', 'هل تريد حذف هذه الشهادة من السجل؟', 'حذف', async () => { if (await write(() => S.db.doc('certificates/' + c.id).delete())) { closeModal(); toast('تم الحذف'); } }));
  }

  function certSettingsSection() {
    const s = certSettings(), dis = writable() ? '' : ' disabled';
    return '<section class="card"><header><h2>بيانات الشهادات</h2><span class="hint">تظهر في شهادات الدخل وشهادات رأس المال</span></header><div class="body"><form id="set-cert" class="form">' +
      fld('sc-line', 'سطر المكتب (نشهد نحن /)', inp('sc-line', s.officeLine, 'text', dis), true) +
      fld('sc-title', 'صفة الموقّع', inp('sc-title', s.signerTitle, 'text', dis)) + fld('sc-name', 'اسم الموقّع', inp('sc-name', s.signerName, 'text', dis)) +
      fld('sc-phones', 'أرقام الهاتف (التذييل)', inp('sc-phones', s.phones, 'text', ' dir="ltr"' + dis)) + fld('sc-email', 'البريد الإلكتروني (التذييل)', inp('sc-email', s.email, 'text', ' dir="ltr"' + dis)) +
      fld('sc-addr', 'العنوان بالإنجليزية (التذييل)', inp('sc-addr', s.addressEn, 'text', ' dir="ltr"' + dis), true) +
      fld('sc-capoffice', 'اسم المكتب في شهادة رأس المال (يشهد مكتب محاسب قانوني :)', inp('sc-capoffice', s.capOffice, 'text', dis)) +
      fld('sc-reg', 'رقم القيد بالسجل العام للمحاسبين والمراجعين', inp('sc-reg', s.registryNo, 'text', dis)) +
      fld('sc-chambers', 'الغرف (غرفة في كل سطر – الأولى تظهر تلقائيًا في شهادة رأس المال)', '<textarea id="sc-chambers" style="min-height:150px"' + dis + '>' + esc(chamberList().join('\n')) + '</textarea>', true) +
      fld('sc-tpl', 'تصميم الشهادة', '<select id="sc-tpl"' + dis + '>' + options([['classic', 'الكلاسيكي (إطار وزخارف بنفس الديباجة)'], ['original', 'الأصلي (مطابق للشهادة القديمة)'], ['modern', 'الحديث (جدول بيانات ومربعات الدخل)']], s.template || 'classic') + '</select>') +
      fld('sc-color', 'اللون المميز', '<input type="color" id="sc-color" value="' + esc(s.color) + '"' + dis + '>') +
      (writable() ? '<div class="field full"><button type="submit" class="btn primary" style="align-self:flex-start">حفظ بيانات الشهادة</button></div>' : '') + '</form></div></section>';
  }
  function bindCertSettings() {
    const f = $('#set-cert'); if (!f || !writable()) return;
    f.addEventListener('submit', (ev) => {
      ev.preventDefault();
      saveSettings({ cert: { officeLine: val('sc-line'), signerTitle: val('sc-title'), signerName: val('sc-name'), phones: val('sc-phones'), email: val('sc-email'), addressEn: val('sc-addr'), capOffice: val('sc-capoffice'), registryNo: val('sc-reg'), chambers: $('#sc-chambers').value.split('\n').map(x => x.trim()).filter(Boolean), color: val('sc-color') || CERT_DEFAULTS.color, template: val('sc-tpl') || 'classic' } });
    });
  }

  /* ---------------- settings ---------------- */
  function pageSettings() {
    const s = settings(), M = S.M, w = writable();
    const dis = w ? '' : ' disabled';
    let h = topBar('الإعدادات', 'بيانات المكتب والموظفين والخدمات والالتزامات الدورية') + '<div class="content">';
    if (!S.settingsSaved && w) h += '<div class="note">هذه هي الإعدادات الافتراضية ولم تُحفظ بعد. راجع بيانات المكتب ومواعيد الالتزامات، ثم اضغط «حفظ» في كل قسم.</div>';
    h += '<section class="card"><header><h2>بيانات المكتب</h2><span class="hint">تظهر في رأس كشوف الحساب والإيصالات والتقارير</span></header><div class="body"><form id="set-office" class="form">' +
      fld('s-name', 'اسم المكتب', inp('s-name', s.officeName, 'text', dis)) + fld('s-sub', 'الوصف المهني', inp('s-sub', s.officeSub, 'text', dis)) +
      fld('s-addr', 'العنوان', inp('s-addr', s.address, 'text', dis)) + fld('s-phone', 'الهاتف', inp('s-phone', s.phone, 'text', dis)) +
      fld('s-email', 'البريد الإلكتروني', inp('s-email', s.email, 'text', dis)) + fld('s-tax', 'رقم التسجيل الضريبي للمكتب', inp('s-tax', s.taxNo, 'text', dis)) +
      fld('s-cur', 'العملة', inp('s-cur', s.currency, 'text', dis)) + fld('s-ad', 'التنبيه قبل موعد الإجراء (أيام)', inp('s-ad', s.alertDays, 'number', ' min="1"' + dis)) +
      fld('s-dd', 'التنبيه قبل انتهاء المستندات (أيام)', inp('s-dd', s.docAlertDays, 'number', ' min="1"' + dis)) +
      fld('s-foot', 'ملاحظة أسفل المستندات المطبوعة', inp('s-foot', s.footer, 'text', dis), true) +
      (w ? '<div class="field full"><button type="submit" class="btn primary" style="align-self:flex-start">حفظ بيانات المكتب</button></div>' : '') + '</form></div></section>';

    h += '<section class="card"><header><h2>موظفو المكتب</h2><span class="grow"></span>' + actBtn('add-staff', '+ موظف') + '</header><div class="tablewrap"><table><thead><tr><th>الاسم</th><th>الوظيفة</th><th>الهاتف</th><th>العملاء المسؤول عنهم</th><th>الحالة</th><th></th></tr></thead><tbody>' +
      (M.staff.length ? M.staff.map(p => '<tr><td>' + esc(p.name) + '</td><td>' + esc(p.title || '') + '</td><td class="num">' + esc(p.phone || '') + '</td><td>' + int(M.active.filter(c => c.assignedTo === p.id).length) + '</td><td>' + (p.active === false ? '<span class="pill mute">غير نشط</span>' : '<span class="pill ok">نشط</span>') + '</td><td>' + actBtn('edit-staff', 'تعديل', 'sm', ' data-id="' + p.id + '"') + '</td></tr>').join('') : '<tr><td colspan="6" class="empty">أضف موظفي المكتب لتوزيع العملاء والإجراءات عليهم.</td></tr>') +
      '</tbody></table></div></section>';

    h += '<section class="card"><header><h2>الخدمات</h2><span class="hint">خدمة في كل سطر؛ تظهر في العقود ومطالبات الأتعاب</span></header><div class="body"><form id="set-svc"><textarea id="s-svc" style="min-height:150px"' + dis + '>' + esc((s.services || []).join('\n')) + '</textarea>' +
      (w ? '<div style="margin-top:10px"><button type="submit" class="btn primary">حفظ الخدمات</button></div>' : '') + '</form></div></section>';

    h += '<section class="card"><header><h2>الالتزامات الدورية ومواعيدها</h2><span class="hint">منها تُولَّد إجراءات العملاء ومواعيدها النهائية</span></header><div class="body">' +
      '<div class="note" style="margin-bottom:12px">المواعيد الافتراضية مبدئية، فراجعها واعتمدها وفق المواعيد القانونية السارية قبل الاعتماد عليها. «بعد نهاية الفترة بـ» عدد الشهور بعد شهر نهاية الفترة، و«اليوم» = 0 يعني آخر يوم في الشهر.</div>' +
      '<div class="tablewrap"><table><thead><tr><th>الالتزام</th><th>الدورية</th><th>لعملاء قانون 6</th><th>القانون</th><th>نوع الممول</th><th>بعد نهاية الفترة بـ (شهر)</th><th>اليوم</th><th>مثال: فترة ' + esc(monthLabel(C.addMonths(M.ym, -1))) + '</th><th>مفعّل</th><th></th></tr></thead><tbody id="ob-rows"></tbody></table></div>' +
      (w ? '<div class="toolbar" style="margin-top:10px"><button type="button" class="btn" id="ob-add">+ التزام</button><span class="grow"></span><button type="button" class="btn primary" id="ob-save">حفظ الالتزامات</button></div>' : '') + '</div></section>';

    h += certSettingsSection();
    h += vaultSettingsSection();
    h += '<section class="card"><header><h2>استيراد العملاء من Excel</h2></header><div class="body"><p style="margin-top:0;color:var(--muted)">نزّل النموذج واملأه، ثم ارفعه. العميل الذي له نفس الكود يُحدَّث، والجديد يُضاف.</p><div class="toolbar">' +
      (S.downloads ? '<button type="button" class="btn" data-act="client-template">⤓ نموذج استيراد العملاء</button>' : '') +
      (w ? '<label class="btn primary" for="imp-clients">رفع ملف Excel</label><input type="file" id="imp-clients" accept=".xlsx,.xls,.csv" hidden>' : '') + '</div></div></section>';

    h += '<section class="card"><header><h2>النسخ الاحتياطي</h2></header><div class="body"><div class="kv" style="margin-bottom:12px">' +
      [['العملاء', M.clients.length], ['الحركات المالية', M.txs.length], ['الإجراءات', M.tasks.length], ['العقود', M.contracts.length], ['الموظفون', M.staff.length]].map(([k, v]) => '<span>' + esc(k) + '</span><span class="num">' + int(v) + '</span>').join('') +
      '</div><div class="toolbar">' + (S.downloads ? '<button type="button" class="btn" data-act="backup">⤓ تنزيل نسخة احتياطية كاملة (.json)</button>' : '') +
      (S.isOwner ? '<label class="btn" for="imp-backup">استيراد نسخة احتياطية</label><input type="file" id="imp-backup" accept=".json,application/json" hidden>' : '') + '</div></div></section>';
    return h + '</div>';
  }

  function bindSettings() {
    if (!writable()) { drawObligations(false); return; }
    $('#set-office').addEventListener('submit', (ev) => {
      ev.preventDefault();
      saveSettings({ officeName: val('s-name'), officeSub: val('s-sub'), address: val('s-addr'), phone: val('s-phone'), email: val('s-email'), taxNo: val('s-tax'),
        currency: val('s-cur') || 'ج.م', alertDays: Number(val('s-ad')) || 7, docAlertDays: Number(val('s-dd')) || 30, footer: val('s-foot') });
    });
    $('#set-svc').addEventListener('submit', (ev) => { ev.preventDefault(); saveSettings({ services: $('#s-svc').value.split('\n').map(x => x.trim()).filter(Boolean) }); });
    drawObligations(true);
    const ic = $('#imp-clients'); if (ic) ic.addEventListener('change', () => { if (ic.files[0]) importClients(ic.files[0]); ic.value = ''; });
    const ib = $('#imp-backup'); if (ib) ib.addEventListener('change', () => { if (ib.files[0]) importBackup(ib.files[0]); ib.value = ''; });
  }
  let obDraft = null;
  function drawObligations(w) {
    obDraft = clean(settings().obligations || []);
    const per = C.addMonths(S.M.ym, -1);
    const draw = () => {
      $('#ob-rows').innerHTML = obDraft.map((o, i) => w ?
        '<tr data-i="' + i + '"><td><input type="text" data-k="name" value="' + esc(o.name) + '" style="width:100%;min-width:300px"></td><td><select data-k="freq">' + options(Object.entries(OB_FREQ), o.freq) + '</select></td>' +
        '<td><select data-k="freqLaw6">' + options(Object.entries(OB_FREQ), o.freqLaw6 != null ? o.freqLaw6 : (o.id === 'vat' ? 'quarterly' : ''), 'نفس الدورية') + '</select></td>' +
        '<td><select data-k="law">' + options(Object.entries(OB_LAW), C.obLaw(o)) + '</select></td><td><select data-k="entity">' + options(Object.entries(OB_ENTITY), C.obEntity(o)) + '</select></td>' +
        '<td><input type="number" min="0" max="12" data-k="dueOffset" value="' + esc(o.dueOffset) + '" style="width:80px"></td><td><input type="number" min="0" max="31" data-k="dueDay" value="' + esc(o.dueDay) + '" style="width:80px"></td>' +
        '<td class="num">' + esc(C.dueDate(o, o.freq === 'annual' ? per.slice(0, 4) + '-12' : o.freq === 'quarterly' ? qEnd(per) : per)) + '</td><td><input type="checkbox" data-k="active"' + (o.active !== false ? ' checked' : '') + '></td>' +
        '<td>' + (S.M.clients.some(c => (c.obligations || []).includes(o.id)) ? '' : '<button type="button" class="btn sm danger" data-rm="' + i + '">حذف</button>') + '</td></tr>' :
        '<tr><td>' + esc(o.name) + '</td><td>' + esc(OB_FREQ[o.freq]) + '</td><td>' + esc(OB_FREQ[C.effFreq(o, { taxLaw: '6-2025' })]) + '</td><td>' + esc(OB_LAW[C.obLaw(o)]) + '</td><td>' + esc(OB_ENTITY[C.obEntity(o)]) + '</td><td>' + esc(o.dueOffset) + '</td><td>' + esc(o.dueDay || 'آخر الشهر') + '</td><td class="num">' + esc(C.dueDate(o, per)) + '</td><td>' + (o.active !== false ? 'نعم' : 'لا') + '</td><td></td></tr>').join('');
      if (!w) return;
      $$('#ob-rows [data-k]').forEach(el => el.addEventListener(el.type === 'text' ? 'input' : 'change', () => {
        const o = obDraft[Number(el.closest('tr').dataset.i)], k = el.dataset.k;
        o[k] = el.type === 'checkbox' ? el.checked : (k === 'dueOffset' || k === 'dueDay') ? Number(el.value) || 0 : el.value;
        if (k !== 'name') draw();
      }));
      $$('#ob-rows [data-rm]').forEach(b => b.addEventListener('click', () => { obDraft.splice(Number(b.dataset.rm), 1); draw(); }));
    };
    draw();
    if (!w) return;
    $('#ob-add').addEventListener('click', () => { obDraft.push({ id: 'ob-' + uid().slice(0, 8), name: 'التزام جديد', freq: 'monthly', dueOffset: 1, dueDay: 0, active: true }); draw(); });
    $('#ob-save').addEventListener('click', () => {
      if (obDraft.some(o => !String(o.name || '').trim())) return toast('اكتب اسم كل التزام.', true);
      saveSettings({ obligations: obDraft.map(o => Object.assign({}, o, { name: o.name.trim() })) });
    });
  }
  const OB_LAW = { all: 'كل العملاء', '91-2005': 'قانون 91 فقط', '6-2025': 'قانون 6 فقط' };
  const OB_ENTITY = { all: 'الكل', legal: 'أشخاص اعتبارية', natural: 'أفراد ومنشآت فردية' };
  const qEnd = (ym) => { const [y, m] = ym.split('-').map(Number); const q = Math.ceil(m / 3) * 3; return y + '-' + String(q).padStart(2, '0'); };
  async function saveSettings(patch) {
    const n = Object.assign({}, S.raw.settings, patch, { knownDefaults: DEFAULT_SETTINGS.obligations.map(o => o.id) });
    if (await write(() => S.db.doc('meta/settings').set(clean(n)))) toast('تم الحفظ');
  }

  function openStaffForm(p) {
    const isNew = !p;
    p = p ? clean(p) : { id: uid(), name: '', title: '', phone: '', email: '', active: true };
    const body = '<form id="sf-form" class="form">' + fld('sf-name', 'الاسم', inp('sf-name', p.name, 'text', ' required')) + fld('sf-title', 'الوظيفة', inp('sf-title', p.title, 'text', ' placeholder="محاسب / مراجع / مسؤول ضرائب"')) +
      fld('sf-phone', 'الهاتف', inp('sf-phone', p.phone, 'tel')) + fld('sf-email', 'البريد الإلكتروني', inp('sf-email', p.email, 'email')) +
      fld('sf-active', 'الحالة', '<select id="sf-active">' + options([['1', 'نشط'], ['0', 'غير نشط']], p.active === false ? '0' : '1') + '</select>') + '<div class="field full err" id="sf-err"></div></form>';
    const used = S.M.clients.some(c => c.assignedTo === p.id) || S.M.tasks.some(t => t.assignedTo === p.id);
    const foot = (!isNew && !used ? '<button type="button" class="btn danger" id="sf-del">حذف</button><span style="flex:1"></span>' : '') + '<button type="button" class="btn" data-close>إلغاء</button><button type="submit" form="sf-form" class="btn primary">حفظ</button>';
    const bg = modal(isNew ? 'موظف جديد' : 'تعديل · ' + p.name, body, foot);
    $('#sf-form', bg).addEventListener('submit', async (ev) => {
      ev.preventDefault();
      const n = Object.assign({}, p, { name: val('sf-name', bg), title: val('sf-title', bg), phone: val('sf-phone', bg), email: val('sf-email', bg), active: val('sf-active', bg) === '1' });
      if (!n.name) return ($('#sf-err', bg).textContent = 'الاسم مطلوب.');
      if (await write(() => S.db.doc('staff/' + n.id).set(clean(n)))) { closeModal(); toast('تم الحفظ'); }
    });
    const del = $('#sf-del', bg); if (del) del.addEventListener('click', async () => { if (await write(() => S.db.doc('staff/' + p.id).delete())) { closeModal(); toast('تم الحذف'); } });
  }

  /* ---------------- import / backup ---------------- */
  const CLIENT_COLS = [['code', 'الكود'], ['name', 'اسم العميل'], ['legalForm', 'الشكل القانوني'], ['activity', 'النشاط'], ['taxRegNo', 'رقم التسجيل الضريبي'], ['taxOffice', 'المأمورية'],
    ['taxLaw', 'القانون الضريبي (91 أو 6)'], ['vatRegistered', 'مسجل في القيمة المضافة (نعم/لا)'], ['payrollRegistered', 'مسجل في كسب العمل (نعم/لا)'], ['insuranceNo', 'رقم التأمينات'], ['contactName', 'المسؤول لدى العميل'], ['phone', 'الهاتف'], ['email', 'البريد الإلكتروني'], ['address', 'العنوان'],
    ['openingBalance', 'الرصيد الافتتاحي'], ['openingDate', 'تاريخ الرصيد الافتتاحي'], ['startDate', 'بداية التعامل'], ['notes', 'ملاحظات']];
  function clientTemplate() {
    xlsxBook('نموذج استيراد العملاء.xlsx', [
      { name: 'العملاء', noHead: true, rows: [CLIENT_COLS.map(c => c[1]), ['C-001', 'شركة المثال للتجارة', 'شركة ذات مسؤولية محدودة', 'تجارة عامة', '123-456-789', 'مأمورية الشركات المساهمة', 91, 'نعم', 'لا', '', 'أ. أحمد', '01000000000', '', 'الجيزة', 15000, '2026-01-01', '2024-03-01', '']] },
      { name: 'تعليمات', noHead: true, rows: [['التعليمات'], ['الكود فريد لكل عميل؛ إذا وُجد عميل بنفس الكود تُحدَّث بياناته.'], ['التواريخ بصيغة YYYY-MM-DD.'], ['الرصيد الافتتاحي موجب إذا كان مستحقًا على العميل، وسالب إذا كان دفعة مقدمة منه.'], ['الشكل القانوني من القائمة: ' + LEGAL.join('، ')], ['القانون الضريبي: اكتب 91 لقانون 91 لسنة 2005 أو 6 لقانون 6 لسنة 2025 (التيسيرات).'], ['التسجيل في القيمة المضافة وكسب العمل: نعم أو لا، ويُضاف الالتزام الدوري المقابل تلقائيًا.']] }
    ]);
  }
  function excelDate(v) {
    if (v == null || v === '') return '';
    if (typeof v === 'number') { const d = new Date(Math.round((v - 25569) * 86400000)); return d.toISOString().slice(0, 10); }
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    const s = String(v).trim(); const m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/.exec(s); if (m) return m[1] + '-' + m[2].padStart(2, '0') + '-' + m[3].padStart(2, '0');
    const m2 = /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/.exec(s); if (m2) return m2[3] + '-' + m2[2].padStart(2, '0') + '-' + m2[1].padStart(2, '0');
    return '';
  }
  function importClients(file) {
    if (!window.XLSX) return toast('يتطلب الاستيراد اتصالًا بالإنترنت عند فتح الصفحة.', true);
    const rd = new FileReader();
    rd.onload = () => {
      let rows;
      try { const wb = XLSX.read(new Uint8Array(rd.result), { type: 'array' }); rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: '' }); } catch (e) { return toast('تعذّرت قراءة الملف.', true); }
      const head = (rows[0] || []).map(x => String(x).trim());
      const idx = {}; for (const [k, l] of CLIENT_COLS) idx[k] = head.indexOf(l);
      if (idx.name < 0) return toast('لم أجد عمود «اسم العميل» في الصف الأول.', true);
      const byCode = new Map(S.M.clients.filter(c => c.code).map(c => [String(c.code).trim(), c]));
      const out = [], errs = []; let nextCode = C.nextClientCode(S.M.clients); const seen = new Set();
      rows.slice(1).forEach((r, i) => {
        const get = (k) => idx[k] >= 0 ? r[idx[k]] : '';
        const name = String(get('name') || '').trim(); if (!name) return;
        let code = String(get('code') || '').trim();
        if (code && seen.has(code)) { errs.push('الصف ' + (i + 2) + ': الكود ' + code + ' مكرر في الملف'); return; }
        const ex = code ? byCode.get(code) : null;
        if (!code) { code = nextCode; nextCode = C.nextClientCode([{ code }].concat(S.M.clients, out)); }
        seen.add(code);
        const base = ex ? clean(ex) : { id: uid(), status: 'active', fyEnd: 12, obligations: [], docs: [], vatRegistered: false, assignedTo: '' };
        const n = Object.assign(base, { code, name });
        for (const [k] of CLIENT_COLS) {
          if (k === 'code' || k === 'name' || idx[k] < 0) continue;
          const v = get(k);
          if (k === 'openingBalance') { if (v !== '') n.openingBalance = C.round2(v); }
          else if (k === 'taxLaw') { if (v !== '') n.taxLaw = /(^|\D)6(\D|$)|2025|تيسير|مبسط/.test(String(v)) ? '6-2025' : '91-2005'; }
          else if (k === 'vatRegistered' || k === 'payrollRegistered') {
            if (v === '') continue;
            const yes = /^(نعم|yes|y|1|true|✓)$/i.test(String(v).trim()), ob = k === 'vatRegistered' ? 'vat' : 'payroll';
            n[k] = yes; n.obligations = (n.obligations || []).filter(x => x !== ob).concat(yes ? [ob] : []);
          }
          else if (k === 'openingDate' || k === 'startDate') { const d = excelDate(v); if (d) n[k] = d; }
          else if (v !== '') n[k] = String(v).trim();
        }
        out.push(n);
      });
      const upd = out.filter(n => S.M.clientById.has(n.id)).length;
      const bg = modal('استيراد العملاء', '<p style="margin:0">سيتم إضافة <b>' + (out.length - upd) + '</b> عميل جديد وتحديث <b>' + upd + '</b> عميل موجود.</p>' + (errs.length ? '<div class="note">' + errs.map(esc).join('<br>') + '</div>' : ''),
        '<button type="button" class="btn" data-close>إلغاء</button><button type="button" class="btn primary" id="ic-go"' + (out.length ? '' : ' disabled') + '>استيراد</button>');
      $('#ic-go', bg).addEventListener('click', async () => {
        $('#ic-go', bg).disabled = true;
        const ok = await write(async () => { for (const n of out) await S.db.doc('clients/' + n.id).set(clean(n)); });
        if (ok) { closeModal(); toast('تم استيراد ' + out.length + ' عميل'); }
      });
    };
    rd.readAsArrayBuffer(file);
  }
  async function backup() {
    const data = { app: 'accounting-office', version: 1, exportedAt: new Date().toISOString(), settings: S.raw.settings };
    for (const c of COLS) data[c] = S.raw[c];
    // portal logins travel encrypted; the file is useless for them without the master password
    if (V.doc) data.vault = V.doc;
    data.credentials = Array.from(V.creds.entries()).map(([id, d]) => Object.assign({ id }, d));
    if (window.OFVat) { try { data.vatInvoices = await window.OFVat.allChunks(); data.vatFiles = await window.OFVat.allFiles(); } catch (e) { return toast('تعذّر تجميع فواتير الإقرارات للنسخة الاحتياطية. حاول مرة أخرى.', true); } }
    data.source = window.OF_LOCAL ? 'device' : 'claude.ai';
    await saveFile('نسخة احتياطية - ' + fileName(settings().officeName || 'المكتب') + ' - ' + C.todayISO() + '.json', new Blob([JSON.stringify(data)], { type: 'application/json' }));
    try { localStorage.setItem('of-last-backup', new Date().toISOString()); } catch (e) {}
    if (currentPage().page === 'backup') renderPage();
  }
  function importBackup(file) {
    const rd = new FileReader();
    rd.onload = () => {
      let d; try { d = JSON.parse(rd.result); } catch (e) { return toast('الملف ليس JSON صالحًا.', true); }
      if (!d || d.app !== 'accounting-office') return toast('هذا الملف ليس نسخة احتياطية من هذا البرنامج.', true);
      const n = COLS.reduce((s, c) => s + (Array.isArray(d[c]) ? d[c].length : 0), 0);
      const cnt = (k) => Array.isArray(d[k]) ? d[k].length : 0;
      const total = n + cnt('vatInvoices') + cnt('vatFiles') + cnt('credentials') + (d.settings ? 1 : 0) + (d.vault ? 1 : 0);
      const bg = modal('استرجاع نسخة احتياطية', '<p style="margin:0 0 8px">' + (d.exportedAt ? 'نسخة بتاريخ ' + esc(String(d.exportedAt).replace('T', ' ').slice(0, 16)) + (d.source ? ' من ' + (d.source === 'device' ? 'نسخة الجهاز' : 'claude.ai') : '') + '. ' : '') + 'تحتوي على:</p>' +
        '<div class="kv"><span>العملاء</span><span>' + int(cnt('clients')) + '</span><span>الحركات المالية</span><span>' + int(cnt('transactions')) + '</span><span>الإجراءات</span><span>' + int(cnt('tasks')) + '</span><span>العقود</span><span>' + int(cnt('contracts')) + '</span><span>الشهادات</span><span>' + int(cnt('certificates')) + '</span><span>إقرارات القيمة المضافة</span><span>' + int(cnt('vatReturns')) + '</span></div>' +
        '<p style="margin:10px 0 0">السجلات التي تحمل المعرّف نفسه تُستبدل، وتبقى بقية البيانات كما هي.</p><p class="hint" id="ib-prog" style="margin:6px 0 0"></p>',
        '<button type="button" class="btn" data-close>إلغاء</button><button type="button" class="btn primary" id="ib-go">بدء الاسترجاع</button>');
      $('#ib-go', bg).addEventListener('click', async () => {
        $('#ib-go', bg).disabled = true;
        let done = 0; const tick = () => { done++; if (done % 10 === 0 || done === total) { const p = $('#ib-prog', bg); if (p) p.textContent = 'جارٍ الاسترجاع… ' + int(done) + ' من ' + int(total); } };
        const set0 = S.db.doc.bind(S.db);
        const db = { doc: (path) => { const r = set0(path); return { set: (v) => r.set(v).then(x => { tick(); return x; }) }; } };
        const ok = await write(async () => {
          for (const c of COLS) for (const r of d[c] || []) { if (!r || !r.id) continue; const { id, ...rest } = r; await db.doc(c + '/' + id).set(clean(rest)); }
          if (d.settings) await db.doc('meta/settings').set(clean(d.settings));
          if (d.vault && d.vault.wrapped) await db.doc('vault/main').set(clean(d.vault));
          for (const r of d.vatFiles || []) { if (!r || !r.id) continue; const { id, ...rest } = r; await db.doc('vatFiles/' + id).set(clean(rest)); }
          for (const r of d.vatInvoices || []) { if (!r || !r.id) continue; const { id, ...rest } = r; await db.doc('vatInvoices/' + id).set(clean(rest)); }
          for (const r of d.credentials || []) { if (!r || !r.id || !r.data) continue; const { id, ...rest } = r; await db.doc('credentials/' + id).set(clean(rest)); }
        });
        if (ok) { if (window.OFVat && window.OFVat.resetCache) window.OFVat.resetCache(); closeModal(); toast('اكتمل استرجاع النسخة الاحتياطية'); scheduleRender(); }
      });
    };
    rd.readAsText(file);
  }

  /* ---------------- click delegation ---------------- */
  document.addEventListener('click', (ev) => {
    const el = ev.target.closest('[data-act],[data-client]');
    if (!el) return;
    if (el.closest('.modal') && !el.dataset.act) return;
    if (el.dataset.act) {
      const a = el.dataset.act, id = el.dataset.id;
      if (el.tagName === 'A') ev.preventDefault();
      const tx = () => S.raw.transactions.find(t => t.id === id);
      switch (a) {
        case 'menu': return $('#side').classList.toggle('open');
        case 'add-client': return openClientForm(null);
        case 'edit-client': return openClientForm(S.raw.clients.find(c => c.id === id));
        case 'del-client': return deleteClient(id);
        case 'add-tx': return openTxForm(null, { kind: el.dataset.kind, client: el.dataset.client });
        case 'edit-tx': return openTxForm(tx());
        case 'receipt': return pdfReceipt(tx());
        case 'add-task': return openTaskForm(null, { client: el.dataset.client });
        case 'edit-task': return openTaskForm(S.raw.tasks.find(t => t.id === id));
        case 'add-contract': return openContractForm(null, { client: el.dataset.client });
        case 'edit-contract': return openContractForm(S.raw.contracts.find(k => k.id === id));
        case 'add-staff': return openStaffForm(null);
        case 'edit-staff': return openStaffForm(S.raw.staff.find(p => p.id === id));
        case 'gen-fees': return writable() ? openGenFees() : null;
        case 'gen-tasks': return writable() ? openGenTasks() : null;
        case 'reminder': return openReminder(id);
        case 'clear-st': S.ui.stFrom = ''; S.ui.stTo = ''; return renderPage();
        case 'pdf': return pdfReport(el.dataset.key);
        case 'xlsx': return xlsxReport(el.dataset.key);
        case 'pdf-statement': return pdfStatement(id, el.dataset.range === 'r');
        case 'xlsx-statement': return xlsxStatement(id, el.dataset.range === 'r');
        case 'client-template': return clientTemplate();
        case 'backup': return backup();
        case 'add-cert': return openCertForm(null);
        case 'add-capcert': return openCapitalForm(null);
        case 'edit-cert': { const c = (S.raw.certificates || []).find(x => x.id === id); return c && c.type === 'capital' ? openCapitalForm(c) : openCertForm(c); }
        case 'copy-cert': { const c = (S.raw.certificates || []).find(x => x.id === id); return c && c.type === 'capital' ? openCapitalForm(null, c) : openCertForm(null, c); }
        case 'cert-pdf': return pdfCertificate((S.raw.certificates || []).find(c => c.id === id));
        case 'cert-preview': return previewCert((S.raw.certificates || []).find(c => c.id === id));
        case 'vault-lock': return lockVault();
        case 'vault-pass': return openVaultPassword();
        case 'cred-edit': return writable() ? openCredForm(id) : null;
        case 'cred-show': { const k = el.dataset.key; if (V.shown.has(k)) V.shown.delete(k); else V.shown.add(k); return renderPage(); }
        default: if (window.OFVat && a.indexOf('vat-') === 0) return window.OFVat.act(a, el);
          break;
        case 'to-mods': return leaveTo('home');
        case 'open-mod': return leaveTo(el.dataset.to, el);
        case 'cred-copy': return navigator.clipboard.writeText(credValue(el.dataset.key)).then(() => toast('تم النسخ'), () => toast('تعذّر النسخ', true));
      }
      return;
    }
    if (el.dataset.client && el.tagName === 'TR') go('client-statement-' + el.dataset.client);
  });
  document.addEventListener('keydown', (ev) => { if (ev.key === 'Enter' && ev.target.matches('.alert[data-act]')) ev.target.click(); });

  /* ---------------- shared with modules (vat.js) ---------------- */
  window.OFApp = { S, C, $, $$, esc, uid, money, moneyC, int, cur, monthLabel, clean, clientName, settings, options, clientOptions, topBar, actBtn, modal, closeModal, toast, write,
    confirmBox, fld, inp, val, writable, go, renderPage, scheduleRender, makePdf, pdfHead, tableHtml, xlsxBook, saveFile, fileName, col, lawOf, MONTHS, TAX_LAW, svgIcon };

  /* ---------------- boot ---------------- */
  renderShell();
  route();
  window.addEventListener('hashchange', route);

  (async function boot() {
    const [db, user, downloads] = await Promise.all([
      window.claude ? window.claude.use('db').catch(() => null) : null,
      window.claude ? window.claude.use('user').catch(() => null) : null,
      window.claude ? window.claude.use('downloads').catch(() => null) : null
    ]);
    S.db = db; S.user = user; S.downloads = downloads;
    if (user) {
      try { S.isOwner = !!(await user.isOwner()); } catch (e) {}
      try { S.canWrite = await user.can('data.write'); } catch (e) { S.canWrite = null; }
    }
    if (!db) { S.dbMissing = true; scheduleRender(); return; }
    const timer = setTimeout(() => { S.slow = true; scheduleRender(); }, 10000);
    let firstDone = 0;
    const done = () => { firstDone++; if (firstDone >= COLS.length + 1) { clearTimeout(timer); S.slow = false; } };
    for (const c of COLS) {
      let got = false;
      db.collection(c).onSnapshot(snap => {
        S.raw[c] = snap.docs.map(d => Object.assign({}, d.data(), { id: d.id }));
        S.loaded[c] = true;
        if (!got) { got = true; done(); }
        scheduleRender();
      }, err => { S.loaded[c] = true; if (!got) { got = true; done(); } toast(err && err.message ? err.message : 'تعذّر تحميل البيانات.', true); scheduleRender(); });
    }
    subscribeVault(db);
    let gotS = false;
    db.doc('meta/settings').onSnapshot(snap => {
      if (snap.exists) {
        const st = Object.assign(JSON.parse(JSON.stringify(DEFAULT_SETTINGS)), clean(snap.data()));
        // built-in obligations added in later versions appear once in saved settings; ones the user deleted stay deleted
        const known = st.knownDefaults || ['vat', 'payroll', 'insurance', 'wht', 'corp', 'indiv', 'audit'];
        for (const d of DEFAULT_SETTINGS.obligations) if (!known.includes(d.id) && !st.obligations.some(o => o.id === d.id)) st.obligations.push(clean(d));
        S.raw.settings = st; S.settingsSaved = true;
      }
      S.loaded.settings = true;
      if (!gotS) { gotS = true; done(); }
      scheduleRender();
    }, () => { S.loaded.settings = true; if (!gotS) { gotS = true; done(); } scheduleRender(); });
  })();
})();
