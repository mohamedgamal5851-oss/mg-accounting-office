/* Tax returns module (الإقرارات الضريبية): VAT returns (Form 10).
   Data: vatReturns/<id> (header, stored summary, submitted-return check), vatInvoices/<id>-<k> (the invoices, in chunks),
   vatParties/<clientId> (what the office set per supplier / customer: address, file no., item type, expense flag). */
(function () {
  'use strict';
  const A = window.OFApp, V = window.OFVatCore;
  if (!A || !V) return;
  const { S, C, $, $$, esc, uid, money, int, clean, toast, write, modal, closeModal, confirmBox, fld, inp, val, options, topBar, actBtn } = A;

  const VS = { cache: new Map(), timers: new Map(), saving: 0, err: '', sel: new Set(), selRet: '', undo: new Map(), dueSel: new Set(), memo: { of: null, res: new Map() } };
  const STATUS = { draft: ['مسودة', 'mute'], ready: ['جاهز للتقديم', 'info'], submitted: ['مقدَّم', 'acc'], matched: ['مقدَّم ومطابق', 'ok'], accepted: ['مطابق بفرق مقبول', 'ok'], diff: ['مقدَّم وبه فروق', 'crit'] };
  const FREQ = { monthly: 'شهري', quarterly: 'ربع سنوي' };
  const S_DT = [[1, 'فاتورة'], [2, 'إشعار إضافة'], [3, 'إشعار خصم (إشعار دائن)'], [5, 'مستهلك نهائي محلي'], [6, 'جهة حكومية'], [7, 'مستهلك نهائي أجنبي (طبيعي)'], [8, 'مستهلك نهائي أجنبي (اعتباري)']];
  const P_DT = [[1, 'فاتورة'], [2, 'إشعار إضافة'], [3, 'إشعار خصم (إشعار دائن)'], [4, 'إذن إفراج جمركي']];
  const S_O = [[1, 'محلي'], [2, 'صادرات'], [5, 'آلات ومعدات'], [6, 'أجزاء آلات وقطع غيار'], [7, 'إعفاءات']];
  const P_O = [[3, 'سلع'], [4, 'خدمات'], [5, 'آلات ومعدات'], [6, 'أجزاء آلات'], [7, 'إعفاءات']];
  const RATES = [[14, '14%'], [5, '5%'], [0, 'صفر / معفى']];
  const dtLabel = (side, dt) => ((side === 'S' ? S_DT : P_DT).find(x => x[0] === dt) || [0, String(dt)])[1];
  const isConsumer = (inv) => inv.side === 'S' && (inv.dt === 5 || inv.dt === 7 || inv.dt === 8);
  const ui = S.ui;
  Object.assign(ui, { dYear: String(new Date().getFullYear()), dClient: '', dStaff: '', vClient: '', vStatus: '', vYear: '', vQ: '', vF: '', vMode: 'group', vSame: '1', rbClient: '', rbSide: 'S', rbFrom: '', rbTo: '', rsYear: '', pClient: '', pQ: '', rmClient: '', rmYear: String(new Date().getFullYear()), duClient: '', duStaff: '', duAge: '', stmtClient: '' });

  /* ---------------- data helpers ---------------- */
  const rets = () => S.raw.vatReturns || [];
  const retById = (id) => rets().find(r => r.id === id);
  const client = (id) => S.M && S.M.clientById.get(id);
  const periodLabel = (r) => r.freq === 'quarterly' ? 'الربع المنتهي ' + A.monthLabel(r.period) : A.monthLabel(r.period);
  const partyDoc = (cid) => (S.raw.vatParties || []).find(p => p.id === cid) || { id: cid, map: {} };
  const memFor = (cid) => { const m = partyDoc(cid).map || {}; return (side, key) => m[side + ':' + key] || null; };
  const statusOf = (r) => r.sub ? (r.sub.ok ? 'matched' : r.sub.accepted ? 'accepted' : 'diff') : (r.status || 'draft');
  const pill = (k) => '<span class="pill ' + STATUS[k][1] + '">' + esc(STATUS[k][0]) + '</span>';

  function prevReturn(r) {
    return rets().filter(x => x.client === r.client && x.period < r.period && x.id !== r.id)
      .sort((a, b) => b.period.localeCompare(a.period) || (a.kind === 'amended' ? -1 : 1) || String(b.updatedAt || '').localeCompare(String(a.updatedAt || '')))[0] || null;
  }
  /* result of a return: the submitted figure when there is one, otherwise the system's */
  function resultOf(r, depth) {
    // remembered until the returns change: dashboards ask for the same results many times per screen
    if (VS.memo.of !== S.raw.vatReturns) { VS.memo.of = S.raw.vatReturns; VS.memo.res = new Map(); }
    if (VS.memo.res.has(r.id)) return VS.memo.res.get(r.id);
    const v = resultOfNow(r, depth); VS.memo.res.set(r.id, v); return v;
  }
  function resultOfNow(r, depth) {
    const sub = r.sub && r.sub.parsed && r.sub.parsed.summary;
    if (sub && sub.res != null && r.sub.src === 'manual') return sub.res;
    const due = r.sum ? r.sum.due : 0;
    return V.r2(due - prevCredit(r, (depth || 0) + 1));
  }
  function prevCredit(r, depth) {
    if (r.prevMode === 'manual' || (depth || 0) > 36) return Number(r.prevCredit) || 0;
    const p = prevReturn(r);
    if (!p) return Number(r.prevCredit) || 0;
    const res = resultOf(p, depth);
    return res < 0 ? -res : 0;
  }

  /* invoice numbers keep digits only; the original stays in noOrig */
  function fixNo(inv) {
    const c = V.normNo(String(inv.no || '').trim());
    if (c !== inv.no) { if (!inv.noOrig) inv.noOrig = inv.no; inv.no = c; return true; }
    return false;
  }
  /* ---------------- loading and saving invoices ---------------- */
  const exists = (snap) => typeof snap.exists === 'function' ? snap.exists() : !!snap.exists;
  async function loadRet(r) {
    if (VS.cache.has(r.id)) return VS.cache.get(r.id);
    const entry = { rows: [], loaded: false, loading: true }; VS.cache.set(r.id, entry);
    try {
      for (let k = 0; k < (r.chunks || 0); k++) {
        const snap = await S.db.doc('vatInvoices/' + r.id + '-' + k).get();
        // data from the database comes back frozen; the invoices are edited in place, so keep an editable copy
        if (exists(snap)) entry.rows.push(...JSON.parse(JSON.stringify((snap.data() || {}).rows || [])));
      }
      entry.loaded = true;
      let fixed = 0; for (const inv of entry.rows) if (fixNo(inv)) fixed++;
      if (fixed && A.writable()) saveSoon(r.id);
    } catch (e) { entry.error = (e && e.message) || 'تعذّر تحميل الفواتير.'; VS.cache.delete(r.id); toast(entry.error, true); }
    entry.loading = false; A.scheduleRender();
    return entry;
  }
  function chunkRows(rows) {
    const out = []; let cur = [], size = 0;
    for (const r of rows) { const n = JSON.stringify(r).length; if (cur.length && size + n > 300000) { out.push(cur); cur = []; size = 0; } cur.push(r); size += n; }
    if (cur.length) out.push(cur);
    return out;
  }
  function summaryOf(r, rows) {
    const calc = V.compute(rows, prevCredit(r));
    const rr = Object.assign({}, r, { defAddr: partyDoc(r.client).defAddr });
    let crit = 0; for (const inv of rows) if (V.issues(inv, rr).some(i => i[0] === 'crit')) crit++;
    const outP = rows.filter(i => !i.x && i.date && (i.date < r.from || i.date > r.to)).length, prior = rows.filter(i => !i.x && i.prior).length;
    return { out: calc.out, inp: calc.inp, due: calc.due, sales: calc.salesTotal.total, inputs: calc.inputsTotal.total, crit, outP, prior, br: V.breakdown(rows),
      nS: rows.filter(i => i.side === 'S').length, nP: rows.filter(i => i.side === 'P').length };
  }
  /* undo: every change to a return's invoices first keeps a copy of how they were (last 30 steps) */
  function snap(id, label) {
    const e = VS.cache.get(id); if (!e) return;
    const st = VS.undo.get(id) || []; st.push({ rows: JSON.stringify(e.rows), label });
    if (st.length > 30) st.shift(); VS.undo.set(id, st);
  }
  function undo(id) {
    const st = VS.undo.get(id), e = VS.cache.get(id); if (!st || !st.length || !e) return toast('لا توجد خطوة للتراجع عنها.');
    const last = st.pop(); e.rows = JSON.parse(last.rows); VS.sel.clear();
    saveSoon(id); A.renderPage(); toast('تم التراجع عن: ' + last.label);
  }
  function saveSoon(id) {
    clearTimeout(VS.timers.get(id));
    VS.timers.set(id, setTimeout(() => saveNow(id), 700));
    setSaveState('جارٍ الحفظ…');
  }
  async function saveNow(id) {
    VS.timers.delete(id);
    const r = retById(id), e = VS.cache.get(id); if (!r || !e || !e.loaded) return;
    VS.saving++; setSaveState('جارٍ الحفظ…');
    const parts = chunkRows(e.rows);
    const ok = await write(async () => {
      for (let k = 0; k < parts.length; k++) await S.db.doc('vatInvoices/' + id + '-' + k).set({ ret: id, k, rows: clean(parts[k]) });
      for (let k = parts.length; k < (r.chunks || 0); k++) await S.db.doc('vatInvoices/' + id + '-' + k).delete();
      const patch = { chunks: parts.length, sum: summaryOf(r, e.rows), keys: e.rows.filter(i => i.side === 'P' && !i.x).map(V.invKey), updatedAt: new Date().toISOString() };
      if (r.sub && r.sub.parsed) { const rec = V.reconcile(V.compute(e.rows, prevCredit(r)), r.sub.parsed, r); patch.sub = Object.assign({}, r.sub, { ok: rec.ok }); }
      await S.db.doc('vatReturns/' + id).update(clean(patch));
    });
    VS.saving--; setSaveState(ok ? 'تم الحفظ' : 'لم يُحفظ');
  }
  function setSaveState(t) { VS.saveText = t; const el = $('#vat-save'); if (el) el.textContent = t; }
  async function saveParties(cid, patch) {
    const d = partyDoc(cid), map = Object.assign({}, d.map || {});
    for (const [k, v] of Object.entries(patch)) map[k] = Object.assign({}, map[k] || {}, v);
    await write(() => S.db.doc('vatParties/' + cid).set(clean({ client: cid, map, defAddr: d.defAddr || '' })));
  }
  async function allFiles() {
    const out = [];
    for (const r of rets()) for (let k = 0; k < ((r.sub && r.sub.fileChunks) || 0); k++) {
      const snap = await S.db.doc('vatFiles/' + r.id + '-' + k).get();
      if (exists(snap)) out.push(Object.assign({ id: r.id + '-' + k }, snap.data()));
    }
    return out;
  }
  async function allChunks() {
    const out = [];
    for (const r of rets()) for (let k = 0; k < (r.chunks || 0); k++) {
      const snap = await S.db.doc('vatInvoices/' + r.id + '-' + k).get();
      if (exists(snap)) out.push(Object.assign({ id: r.id + '-' + k }, snap.data()));
    }
    return out;
  }

  /* ---------------- routing ---------------- */
  function parse(sub) {
    const p = String(sub || '').split('-');
    if (p[0] === 'ret') return { view: 'ret', tab: p[1] || 'sum', id: p.slice(2).join('-') };
    if (p[0] === 'rep') return { view: 'rep', key: p[1] || 'months' };
    if (p[0] === 'parties') return { view: 'parties' };
    if (p[0] === 'list') return { view: 'list' };
    if (p[0] === 'due') return { view: 'due', client: p.slice(1).join('-') };
    if (p[0] === 'hist') return { view: 'hist' };
    return { view: 'dash' };
  }
  function page(sub) {
    const r = parse(sub);
    if (!S.db) return '';
    if (r.view === 'ret') return pageRet(r.id, r.tab);
    if (r.view === 'rep') return pageReports(r.key);
    if (r.view === 'parties') return pageParties();
    if (r.view === 'list') return pageList();
    if (r.view === 'due') return pageDue(r.client);
    if (r.view === 'hist') return pageHist();
    return pageDash();
  }
  function bind(sub) {
    const r = parse(sub);
    if (r.view === 'ret') bindRet(r.id, r.tab);
    if (r.view === 'parties') bindParties();
    if (r.view === 'due' && !r.client) bindDue();
    if (r.view === 'hist') bindHist();
  }

  /* ---------------- dashboard ---------------- */
  const vatOb = () => (S.M.obligations || []).find(o => o.id === 'vat') || { id: 'vat', dueOffset: 1, dueDay: 0 };
  const isQ = (c) => A.lawOf(c) === '6-2025';
  const resOf = (r) => V.r2((r.sum ? r.sum.due || 0 : 0) - prevCredit(r));
  const done = (r) => !!r && !!r.sub;
  function vatClients() {
    return S.M.clients.filter(c => c.status !== 'inactive' && (c.vatRegistered || (c.obligations || []).includes('vat') || rets().some(r => r.client === c.id)) &&
      (!ui.dClient || c.id === ui.dClient) && (!ui.dStaff || c.assignedTo === ui.dStaff));
  }
  /* periods a client owes: monthly, or quarter ends under Law 6; from its first return (or start date) on */
  function periodsOf(c, toYM) {
    const first = [rets().filter(r => r.client === c.id).map(r => r.period).sort()[0], (c.startDate || '').slice(0, 7)].filter(Boolean).sort()[0] || toYM;
    const out = []; let p = first;
    while (p <= toYM && out.length < 120) { const m = Number(p.slice(5)); if (!isQ(c) || m % 3 === 0) out.push(p); p = C.addMonths(p, 1); }
    return out;
  }
  function svgBars(labels, series, h) {
    // grouped bars, one group per label; series: [{name, color, values}]
    const W = 640, H = h || 200, pad = 34, n = labels.length, g = (W - pad) / n, bw = Math.max(3, (g - 8) / series.length);
    const max = Math.max(1, ...series.flatMap(x => x.values.map(v => Math.abs(v))));
    let o = '<svg viewBox="0 0 ' + W + ' ' + (H + 24) + '" class="vchart" role="img"><line x1="' + pad + '" y1="' + H + '" x2="' + W + '" y2="' + H + '" class="ax"/>';
    for (let k = 0; k <= 4; k++) { const y = H - (H - 10) * k / 4; o += '<line x1="' + pad + '" y1="' + y + '" x2="' + W + '" y2="' + y + '" class="gl"/><text x="' + (pad - 4) + '" y="' + (y + 4) + '" class="tk" text-anchor="end">' + esc(short(max * k / 4)) + '</text>'; }
    labels.forEach((l, i) => {
      const x0 = W - (i + 1) * g + 4; // right to left
      series.forEach((sr, j) => { const v = Math.max(0, sr.values[i] || 0), hh = (H - 10) * v / max; o += '<rect x="' + (x0 + j * bw) + '" y="' + (H - hh) + '" width="' + (bw - 1) + '" height="' + hh + '" rx="2" fill="' + sr.color + '"><title>' + esc(sr.name + ' · ' + l + ': ' + money(sr.values[i] || 0)) + '</title></rect>'; });
      o += '<text x="' + (x0 + (g - 8) / 2) + '" y="' + (H + 16) + '" class="tk" text-anchor="middle">' + esc(l) + '</text>';
    });
    return o + '</svg>' + '<div class="vlegend">' + series.map(x => '<span><i style="background:' + x.color + '"></i>' + esc(x.name) + '</span>').join('') + '</div>';
  }
  function svgLine(labels, values, color) {
    const W = 640, H = 170, pad = 34, n = labels.length, step = (W - pad - 10) / Math.max(1, n - 1);
    const max = Math.max(1, ...values.map(v => v || 0));
    // months without a return have no point (null), so the line does not dive to zero
    const pts = values.map((v, i) => v == null ? null : [W - 10 - i * step, H - (H - 10) * v / max]);
    let o = '<svg viewBox="0 0 ' + W + ' ' + (H + 24) + '" class="vchart" role="img"><line x1="' + pad + '" y1="' + H + '" x2="' + W + '" y2="' + H + '" class="ax"/>';
    for (let k = 0; k <= 4; k++) { const y = H - (H - 10) * k / 4; o += '<line x1="' + pad + '" y1="' + y + '" x2="' + W + '" y2="' + y + '" class="gl"/><text x="' + (pad - 4) + '" y="' + (y + 4) + '" class="tk" text-anchor="end">' + esc(short(max * k / 4)) + '</text>'; }
    o += '<polyline fill="none" stroke="' + color + '" stroke-width="2.5" points="' + pts.filter(Boolean).map(p => p.join(',')).join(' ') + '"/>';
    labels.forEach((l, i) => { if (!pts[i]) o += '<text x="' + (W - 10 - i * step) + '" y="' + (H + 16) + '" class="tk" text-anchor="middle">' + esc(l) + '</text>'; });
    pts.forEach((p, i) => { if (!p) return; o += '<circle cx="' + p[0] + '" cy="' + p[1] + '" r="3.5" fill="' + color + '"><title>' + esc(labels[i] + ': ' + money(values[i] || 0)) + '</title></circle><text x="' + p[0] + '" y="' + (H + 16) + '" class="tk" text-anchor="middle">' + esc(labels[i]) + '</text>'; });
    return o + '</svg>';
  }
  const short = (v) => { const a = Math.abs(v); return a >= 1e6 ? (v / 1e6).toFixed(1) + 'م' : a >= 1e3 ? Math.round(v / 1e3) + 'ك' : String(Math.round(v)); };
  const MSHORT = ['ينا', 'فبر', 'مار', 'أبر', 'ماي', 'يون', 'يول', 'أغس', 'سبت', 'أكت', 'نوف', 'ديس'];

  function pageDash() {
    const y = ui.dYear || String(new Date().getFullYear()), today = C.todayISO(), ob = vatOb();
    const cl = vatClients(), ids = new Set(cl.map(c => c.id));
    const all = rets().filter(r => ids.has(r.client));
    const years = Array.from(new Set(rets().map(r => r.period.slice(0, 4)).concat([String(new Date().getFullYear())]))).sort().reverse();
    let h = topBar('لوحة الإقرارات', 'متابعة إقرارات القيمة المضافة لجميع الشركات', actBtn('vat-new', '+ إقرار جديد', 'primary') + '<a class="btn" href="#vat-list">كل الإقرارات</a>') + '<div class="content">';
    h += '<div class="toolbar"><select data-ui="dYear">' + options(years, y) + '</select><select data-ui="dClient">' + options(A.clientOptions(true), ui.dClient, 'جميع الشركات') + '</select>' +
      '<select data-ui="dStaff">' + options(S.M.staff.filter(x => x.active !== false).map(x => [x.id, x.name]), ui.dStaff, 'جميع الموظفين') + '</select>' +
      (ui.dClient || ui.dStaff ? '<button type="button" class="btn sm" data-act="vat-dash-clear">إلغاء التصفية</button>' : '') + '</div>';
    if (!cl.length) return h + '<div class="card"><div class="hero-empty"><h2>لا توجد شركات مسجلة في ضريبة القيمة المضافة</h2><p>فعّل خيار «مسجل في ضريبة القيمة المضافة» في بيانات الشركة من قسم حسابات وإجراءات الشركات، أو أنشئ أول إقرار من «+ إقرار جديد».</p></div></div></div>';

    // 1) this cycle: the period whose deadline falls this month
    const cycle = C.addMonths(S.M.ym, -1);
    const dueRows = cl.map(c => { const per = isQ(c) ? (Number(cycle.slice(5)) % 3 === 0 ? cycle : null) : cycle; return per ? { c, per, due: C.dueDate(ob, per), r: retFor(c.id, per) } : null; }).filter(Boolean);
    const dDone = dueRows.filter(x => done(x.r)).length, deadline = dueRows.length ? dueRows[0].due : '';
    const daysLeft = deadline ? C.daysBetween(today, deadline) : 0;
    // overdue: every owed period whose deadline passed without a submitted return
    const overdue = [];
    for (const c of cl) for (const per of periodsOf(c, C.addMonths(S.M.ym, -1))) { const due = C.dueDate(ob, per); if (due < today && !done(retFor(c.id, per))) overdue.push({ c, per, due }); }
    h += '<section class="card"><div class="body dl-strip">' +
      '<div><small>إقرارات ' + esc(A.monthLabel(cycle)) + '</small><b>' + int(dueRows.length) + '</b><span>مطلوبة</span></div>' +
      '<div><small>تم تقديمها</small><b class="match-ok">' + int(dDone) + '</b><span>من ' + int(dueRows.length) + '</span></div>' +
      '<div><small>المتبقي</small><b' + (dueRows.length - dDone ? ' class="match-bad"' : '') + '>' + int(dueRows.length - dDone) + '</b><span>إقرار</span></div>' +
      '<div><small>آخر ميعاد</small><b>' + esc(deadline || '—') + '</b><span>' + (deadline ? (daysLeft > 0 ? 'متبقٍّ ' + int(daysLeft) + ' يوم' : daysLeft === 0 ? 'اليوم' : 'مضى ' + int(-daysLeft) + ' يوم') : '') + '</span></div>' +
      '<div class="dl-bar"><i style="width:' + (dueRows.length ? Math.round(dDone * 100 / dueRows.length) : 0) + '%"></i></div></div>';
    if (dueRows.length - dDone) h += '<div class="dl-left">' + dueRows.filter(x => !done(x.r)).map(x => x.r ? '<a class="pill ' + STATUS[statusOf(x.r)][1] + '" href="#vat-ret-sum-' + x.r.id + '">' + esc(x.c.name) + ' · ' + esc(STATUS[statusOf(x.r)][0]) + '</a>' : '<a class="pill mute" href="#" data-act="vat-new" data-client="' + x.c.id + '" data-period="' + x.per + '">' + esc(x.c.name) + ' · لم يُعَدّ</a>').join('') + '</div>';
    h += '</section>';
    if (overdue.length) h += '<div class="note" style="background:var(--crit-soft)"><b style="color:var(--crit)">' + int(overdue.length) + ' إقرار تجاوز موعده ولم يُقدَّم:</b> ' +
      overdue.slice(0, 12).map(o => { const r = retFor(o.c.id, o.per); return r ? '<a href="#vat-ret-sum-' + r.id + '">' + esc(o.c.name) + ' – ' + esc(A.monthLabel(o.per)) + '</a>' : '<a href="#" data-act="vat-new" data-client="' + o.c.id + '" data-period="' + o.per + '">' + esc(o.c.name) + ' – ' + esc(A.monthLabel(o.per)) + '</a>'; }).join('، ') + (overdue.length > 12 ? ' …' : '') + '</div>';

    // 2) tiles: latest return of each company
    const latest = cl.map(c => all.filter(r => r.client === c.id).sort((a, b) => b.period.localeCompare(a.period))[0]).filter(Boolean);
    const cred = latest.reduce((a, r) => a + Math.max(0, -resOf(r)), 0), unpaid = dueList({ client: ui.dClient, staff: ui.dStaff });
    const yr = all.filter(r => r.period.startsWith(y));
    const crit = all.reduce((a, r) => a + ((r.sum && r.sum.crit) || 0), 0);
    h += '<div class="tiles">' +
      '<a class="tile c2" href="#vat-due"><span>ضريبة مستحقة لم تُسدَّد</span><b>' + money(unpaid.reduce((a, x) => a + x.left, 0)) + '</b><small>' + int(unpaid.length) + ' إقرار' + (unpaid.some(x => x.late > 0) ? ' · ' + int(unpaid.filter(x => x.late > 0).length) + ' متأخر السداد' : '') + '</small></a>' +
      '<div class="tile c3"><span>أرصدة دائنة مرحّلة</span><b>' + money(cred) + '</b><small>' + int(latest.filter(r => resOf(r) < 0).length) + ' شركة</small></div>' +
      '<a class="tile c4" href="#vat-list"><span>مقدَّمة ومطابقة · ' + esc(y) + '</span><b>' + int(yr.filter(r => /^(matched|accepted)$/.test(statusOf(r))).length) + '</b><small>' + int(yr.filter(r => statusOf(r) === 'diff').length) + ' فيها فروق</small></a>' +
      '<div class="tile c5"><span>فواتير عليها ملاحظات لازمة</span><b>' + int(crit) + '</b><small>في ' + int(all.filter(r => r.sum && r.sum.crit).length) + ' إقرار</small></div></div>';

    // 3) companies × months
    h += '<section class="card"><header><h2>حالة الإقرارات بالشهور · ' + esc(y) + '</h2><span class="grow"></span><span class="vlegend">' +
      [['matched', 'مطابق'], ['diff', 'فروق'], ['submitted', 'مقدَّم'], ['ready', 'جاهز'], ['draft', 'مسودة'], ['late', 'متأخر لم يُعَدّ']].map(([k, l]) => '<span><i class="hm-' + k + '"></i>' + l + '</span>').join('') + '<span><i class="hm-dot"></i>متأخر السداد</span></span></header><div class="tablewrap"><table class="heat"><thead><tr><th>الشركة</th><th class="hm-freq-h">دورية الإقرار</th>' + MSHORT.map(m => '<th>' + m + '</th>').join('') + '</tr></thead><tbody>';
    for (const c of cl) {
      h += '<tr><td><a href="#vat-list" data-act="vat-pick" data-client="' + c.id + '" title="' + esc(c.name) + '">' + esc(c.name) + '</a></td><td class="hm-freq">' + (isQ(c) ? 'ربع سنوي' : 'شهري') + '</td>';
      for (let m = 1; m <= 12; m++) {
        const per = y + '-' + String(m).padStart(2, '0');
        if (isQ(c) && m % 3) { h += '<td class="hm-na"></td>'; continue; }
        const r = retFor(c.id, per);
        if (r) { const st = statusOf(r), res = resOf(r), pi = res > 0 ? payInfo(r) : null, owe = pi && pi.left > 0 && pi.late > 0;
          h += '<td class="hm-' + st + '"><a href="#vat-ret-sum-' + r.id + '" title="' + esc(STATUS[st][0] + ' · ' + (res > 0 ? 'مستحق ' + money(res) + ' · ' + PAYST[pi.st][0] : res < 0 ? 'دائن ' + money(-res) : 'صفر')) + '">' + (res > 0 ? short(res) : res < 0 ? '−' + short(-res) : '0') + (owe ? '<i class="hm-dot"></i>' : pi && pi.st === 'paid' ? '<i class="hm-paid">✓</i>' : '') + '</a></td>'; continue; }
        const late = C.dueDate(ob, per) < today && periodsOf(c, per).includes(per);
        h += '<td class="' + (late ? 'hm-late' : '') + '">' + (late && A.writable() ? '<a href="#" data-act="vat-new" data-client="' + c.id + '" data-period="' + per + '" title="لم يُعَدّ · اضغط لإنشائه">!</a>' : '') + '</td>';
      }
      h += '</tr>';
    }
    h += '</tbody></table></div><div class="hint" style="padding:8px 16px">الرقم في الخانة: المستحق السداد، والقيمة السالبة رصيد دائن (ك = ألف). اضغط على الخانة لفتح الإقرار.</div></section>';

    // 4) needs action
    const act = [];
    for (const r of all.filter(r => statusOf(r) === 'diff')) act.push(['crit', 'فروق بعد المطابقة', r]);
    for (const r of all.filter(r => statusOf(r) === 'ready')) act.push(['info', 'جاهز ولم يُقدَّم بعد', r]);
    for (const x of unpaid.filter(x => x.late > 0)) act.push([x.late > 60 ? 'crit' : 'warn', 'ضريبة ' + money(x.left) + ' لم تُسدَّد · ' + lateText(x), x.r]);
    for (const r of all.filter(r => r.sum && (r.sum.prior || r.sum.outP))) act.push(['warn', (r.sum.prior ? int(r.sum.prior) + ' فاتورة شراء مدرجة في إقرار آخر' : '') + (r.sum.prior && r.sum.outP ? ' · ' : '') + (r.sum.outP ? int(r.sum.outP) + ' مستند تاريخه خارج الفترة' : ''), r]);
    for (const c of cl) {
      const list = all.filter(r => r.client === c.id).sort((a, b) => b.period.localeCompare(a.period));
      let n = 0; for (const r of list) { if (resOf(r) < 0) n++; else break; }
      if (n >= 3) act.push(['warn', 'رصيد دائن متراكم منذ ' + n + ' إقرارات (' + money(-resOf(list[0])) + ') · قد يستدعي تقديم طلب رد', list[0]]);
    }
    h += '<section class="card"><header><h2>بنود تتطلب إجراءً</h2><span class="hint">' + int(act.length) + '</span></header>' +
      (act.length ? '<div class="tablewrap"><table><tbody>' + act.slice(0, 30).map(([k, t, r]) => '<tr><td><span class="iss ' + k + '">' + esc(t) + '</span></td><td><a href="#vat-ret-sum-' + r.id + '">' + esc(A.clientName(r.client)) + '</a></td><td>' + esc(periodLabel(r)) + '</td><td>' + pill(statusOf(r)) + '</td></tr>').join('') + '</tbody></table></div>' : '<div class="empty">لا توجد بنود معلّقة.</div>') + '</section>';

    // 5) charts
    const months = Array.from({ length: 12 }, (_, i) => y + '-' + String(i + 1).padStart(2, '0'));
    const sumBy = (f) => months.map(per => cl.reduce((a, c) => { const r = retFor(c.id, per); return a + (r ? f(r) : 0); }, 0));
    h += '<div class="grid2e">';
    h += '<section class="card"><header><h2>ضريبة المبيعات والمدخلات بالشهور</h2><span class="hint">' + esc(ui.dClient ? A.clientName(ui.dClient) : 'جميع الشركات') + '</span></header><div class="body">' +
      svgBars(MSHORT, [{ name: 'ضريبة المبيعات', color: 'var(--accent)', values: sumBy(r => (r.sum && r.sum.out) || 0) }, { name: 'ضريبة المدخلات', color: 'var(--muted)', values: sumBy(r => (r.sum && r.sum.inp) || 0) }]) + '</div></section>';
    h += '<section class="card"><header><h2>الرصيد الدائن آخر كل شهر</h2><span class="hint">' + esc(ui.dClient ? A.clientName(ui.dClient) : 'جميع الشركات') + '</span></header><div class="body">' +
      svgLine(MSHORT, months.map(per => { const rs = cl.map(c => retFor(c.id, per)).filter(Boolean); return rs.length ? rs.reduce((a, r) => a + Math.max(0, -resOf(r)), 0) : null; }), 'var(--ok)') + '</div></section>';
    h += '</div>';
    const top = cl.map(c => ({ c, v: all.filter(r => r.client === c.id && r.period.startsWith(y)).reduce((a, r) => a + Math.max(0, resOf(r)), 0) })).filter(x => x.v > 0).sort((a, b) => b.v - a.v).slice(0, 5);
    const tmax = Math.max(1, ...top.map(x => x.v));
    h += '<section class="card"><header><h2>أكبر 5 شركات في الضريبة المستحقة السداد · ' + esc(y) + '</h2></header><div class="body">' +
      (top.length ? top.map(x => '<div class="hbar"><span>' + esc(x.c.name) + '</span><div><i style="width:' + Math.round(x.v * 100 / tmax) + '%"></i></div><b>' + money(x.v) + '</b></div>').join('') : '<div class="empty">لا توجد ضريبة مستحقة السداد في هذه السنة.</div>') + '</div></section>';
    return h + '</div>';
  }

  /* ---------------- tax due for payment (سداد الضريبة) ----------------
     Each return may carry pay: [{id, date, amount, method, ref, note, at}]. What a client owes for a period is the result of
     the return that counts for it (the latest amended one), less every payment recorded on any return of that period. */
  const PAYST = { unpaid: ['مستحق ولم يُسدَّد', 'crit'], partial: ['مسدد جزئيًا', 'warn'], paid: ['مسدد', 'ok'], none: ['لا يوجد مستحق', 'mute'], off: ['قبل بداية المتابعة', 'mute'] };
  const PAY_METHODS = ['تحويل بنكي', 'سداد إلكتروني', 'شيك', 'نقدًا بالمأمورية', 'خصم من رصيد دائن'];
  const AGE = [['soon', 'لم يحل موعده'], ['a30', 'متأخر حتى 30 يومًا'], ['a60', 'متأخر 31 – 60 يومًا'], ['a90', 'متأخر 61 – 90 يومًا'], ['a90p', 'متأخر أكثر من 90 يومًا']];
  const ageOf = (late) => late <= 0 ? 'soon' : late <= 30 ? 'a30' : late <= 60 ? 'a60' : late <= 90 ? 'a90' : 'a90p';
  const payPill = (k) => '<span class="pill ' + PAYST[k][1] + '">' + esc(PAYST[k][0]) + '</span>';
  const payCfg = () => { const s = A.settings(); return { from: s.vatPayFrom || '', rate: Number(s.vatLateRate) || 0, part: s.vatLatePart !== false }; };
  const canPay = () => !!S.db && S.canWrite !== false && (!A.level || A.level('vat') === 'edit');
  const LRM = '‎';
  const numIn = (v) => C.round2(String(v == null ? '' : v).replace(/[^\d.\-]/g, ''));
  const grp = (n) => { const s = String(C.round2(n)); const [i, f] = s.split('.'); return i.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (f ? '.' + f : ''); };
  function paysOf(cid, period) {
    return rets().filter(r => r.client === cid && r.period === period)
      .flatMap(r => (r.pay || []).map(p => Object.assign({ ret: r.id }, p))).sort((a, b) => String(a.date).localeCompare(String(b.date)));
  }
  function payInfo(r) {
    const cfg = payCfg(), today = C.todayISO();
    const amt = V.r2(Math.max(0, resultOf(r)));
    const pays = paysOf(r.client, r.period), paid = V.r2(pays.reduce((a, p) => a + (Number(p.amount) || 0), 0));
    const due = C.dueDate(vatOb(), r.period);
    const off = !!cfg.from && r.period < cfg.from && !pays.length;
    let left = V.r2(Math.max(0, amt - paid));
    const st = amt <= 0 && !paid ? 'none' : off ? 'off' : left <= 0.004 ? 'paid' : paid > 0 ? 'partial' : 'unpaid';
    if (st === 'off' || st === 'none') left = 0;
    const late = left > 0 ? Math.max(0, C.daysBetween(due, today)) : 0;
    const months = late > 0 ? (cfg.part ? Math.ceil(late / 30) : Math.floor(late / 30)) : 0;
    const extra = cfg.rate && left > 0 ? V.r2(left * cfg.rate / 100 * months) : 0;
    return { r, amt, paid, pays, left, due, late, toGo: C.daysBetween(today, due), st, months, extra, age: ageOf(late) };
  }
  /* every period still owed, latest deadline missed first */
  function dueList(opt) {
    opt = opt || {};
    const out = [];
    for (const c of S.M.clients) {
      if ((opt.client && c.id !== opt.client) || (opt.staff && c.assignedTo !== opt.staff)) continue;
      for (const per of new Set(rets().filter(r => r.client === c.id).map(r => r.period))) {
        const p = payInfo(retFor(c.id, per));
        if (p.left > 0) out.push(Object.assign({ c, key: c.id + '|' + per }, p));
      }
    }
    return out.sort((a, b) => b.late - a.late || a.due.localeCompare(b.due) || b.left - a.left);
  }
  function dueSummary() {
    if (!S.M) return { total: 0, lateTotal: 0, n: 0, nLate: 0, extra: 0, rows: [] };
    const all = dueList(), late = all.filter(x => x.late > 0);
    return { total: all.reduce((a, x) => a + x.left, 0), lateTotal: late.reduce((a, x) => a + x.left, 0), extra: all.reduce((a, x) => a + x.extra, 0),
      n: all.length, nLate: late.length, clients: new Set(all.map(x => x.c.id)).size, rows: all };
  }
  const lateText = (p) => p.late > 0 ? 'متأخر ' + int(p.late) + ' يوم' : p.toGo === 0 ? 'آخر موعد اليوم' : 'متبقٍّ ' + int(p.toGo) + ' يوم';

  function pageDue(cid) {
    if (cid) return pageTaxStatement(cid);
    const cfg = payCfg(), all = dueList({ client: ui.duClient, staff: ui.duStaff });
    const list = all.filter(x => !ui.duAge || (ui.duAge === 'late' ? x.late > 0 : x.age === ui.duAge));
    const sum = (arr, f) => arr.reduce((a, x) => a + (f ? f(x) : x.left), 0);
    const late = all.filter(x => x.late > 0), soon = all.filter(x => x.late <= 0);
    let h = topBar('الضريبة المستحقة السداد', 'ضريبة القيمة المضافة المستحقة من واقع الإقرارات، وما سُدِّد منها وما تأخر',
      '<button type="button" class="btn" data-act="vat-pay-cfg">⚙ إعدادات المتابعة</button>' +
      (S.downloads && all.length ? '<button type="button" class="btn" data-act="vat-rep-pdf" data-key="due">⤓ PDF</button><button type="button" class="btn" data-act="vat-rep-xlsx" data-key="due">⤓ Excel</button>' : '')) + '<div class="content">';
    h += '<div class="tiles">' +
      '<a class="tile c2" href="#" data-act="vat-due-age" data-v=""><span>إجمالي غير المسدد</span><b>' + money(sum(all)) + '</b><small>' + int(all.length) + ' إقرار · ' + int(new Set(all.map(x => x.c.id)).size) + ' شركة</small></a>' +
      '<a class="tile c5" href="#" data-act="vat-due-age" data-v="late"><span>تجاوز موعد السداد</span><b class="' + (late.length ? 'match-bad' : '') + '">' + money(sum(late)) + '</b><small>' + int(late.length) + ' إقرار</small></a>' +
      '<a class="tile c4" href="#" data-act="vat-due-age" data-v="soon"><span>لم يحل موعده بعد</span><b>' + money(sum(soon)) + '</b><small>' + int(soon.length) + ' إقرار</small></a>' +
      (cfg.rate ? '<div class="tile c3"><span>ضريبة إضافية تقديرية</span><b>' + money(sum(all, x => x.extra)) + '</b><small>' + esc(String(cfg.rate)) + '% عن كل شهر تأخير</small></div>' : '') + '</div>';
    // aging: how long the unpaid tax has been late
    h += '<section class="card"><header><h2>أعمار الضريبة غير المسددة</h2><span class="hint">اضغط على الفئة لتصفية الجدول</span></header><div class="body"><div class="aging">' +
      AGE.map(([k, l]) => { const g = all.filter(x => x.age === k); return '<a href="#" data-act="vat-due-age" data-v="' + k + '" class="age-' + k + (ui.duAge === k ? ' on' : '') + '"><small>' + esc(l) + '</small><b>' + money(sum(g)) + '</b><span>' + int(g.length) + ' إقرار</span></a>'; }).join('') + '</div></div></section>';
    const oldOnes = all.filter(x => x.late > 60);
    if (!cfg.from && oldOnes.length) h += '<div class="note">فيه ' + int(oldOnes.length) + ' إقرار متأخر أكثر من شهرين. لو الإقرارات القديمة اتسددت قبل استخدام البرنامج، حدد «بداية متابعة السداد» من ⚙ إعدادات المتابعة فتختفي من هنا، أو حددها من الجدول واضغط «تسجيل المحدد كمسدد».</div>';
    h += '<div class="toolbar"><select data-ui="duClient">' + options(A.clientOptions(true), ui.duClient, 'جميع الشركات') + '</select>' +
      '<select data-ui="duStaff">' + options(S.M.staff.filter(x => x.active !== false).map(x => [x.id, x.name]), ui.duStaff, 'جميع الموظفين') + '</select>' +
      '<select data-ui="duAge">' + options([['late', 'المتأخر فقط']].concat(AGE), ui.duAge, 'كل الأعمار') + '</select>' +
      (ui.duClient || ui.duStaff || ui.duAge ? '<button type="button" class="btn sm" data-act="vat-due-clear">إلغاء التصفية</button>' : '') +
      '<span class="grow"></span>' + (canPay() && VS.dueSel.size ? '<span class="hint">' + int(VS.dueSel.size) + ' محدد</span><button type="button" class="btn primary" data-act="vat-pay-bulk">تسجيل المحدد كمسدد</button>' : '') + '</div>';
    const showX = !!cfg.rate;
    h += '<div class="card"><div class="tablewrap"><table class="vat-tbl due-tbl"><thead><tr>' + (canPay() ? '<th><input type="checkbox" data-due-all' + (list.length && list.every(x => VS.dueSel.has(x.key)) ? ' checked' : '') + ' title="تحديد الكل"></th>' : '') +
      '<th>الشركة</th><th>الفترة</th><th>حالة الإقرار</th><th>آخر موعد للسداد</th><th>التأخير</th><th class="n">المستحق</th><th class="n">المسدد</th><th class="n">المتبقي</th>' + (showX ? '<th class="n">ض. إضافية تقديرية</th>' : '') + '<th></th></tr></thead><tbody>';
    if (!list.length) h += '<tr><td colspan="' + (showX ? 11 : 10) + '" class="empty">' + (all.length ? 'لا يوجد مستحق بهذه الشروط.' : 'لا توجد ضريبة مستحقة غير مسددة. 👌') + '</td></tr>';
    for (const x of list) {
      h += '<tr class="' + (x.late > 60 ? 'dl-crit' : x.late > 0 ? 'dl-warn' : '') + '">' + (canPay() ? '<td><input type="checkbox" data-due-sel="' + esc(x.key) + '"' + (VS.dueSel.has(x.key) ? ' checked' : '') + '></td>' : '') +
        '<td><a href="#vat-due-' + x.c.id + '" title="كشف الحساب الضريبي">' + esc(x.c.name) + '</a></td><td><a href="#vat-ret-sum-' + x.r.id + '">' + esc(periodLabel(x.r)) + '</a></td><td>' + pill(statusOf(x.r)) + '</td>' +
        '<td class="num">' + esc(x.due) + '</td><td><span class="iss ' + (x.late > 60 ? 'crit' : x.late > 0 ? 'warn' : 'info') + '">' + esc(lateText(x)) + '</span></td>' +
        '<td class="n">' + money(x.amt) + '</td><td class="n">' + (x.paid ? money(x.paid) : '') + '</td><td class="n pos"><b>' + money(x.left) + '</b></td>' + (showX ? '<td class="n">' + (x.extra ? money(x.extra) : '') + '</td>' : '') +
        '<td class="acts">' + (canPay() ? '<button type="button" class="btn sm primary" data-act="vat-pay" data-id="' + x.r.id + '">تسجيل سداد</button>' : '') + '<button type="button" class="btn sm" data-act="vat-pay-remind" data-client="' + x.c.id + '">تذكير</button></td></tr>';
    }
    h += '</tbody>' + (list.length ? '<tfoot><tr>' + (canPay() ? '<td></td>' : '') + '<td colspan="5">الإجمالي</td><td class="n">' + money(sum(list, x => x.amt)) + '</td><td class="n">' + money(sum(list, x => x.paid)) + '</td><td class="n">' + money(sum(list)) + '</td>' + (showX ? '<td class="n">' + money(sum(list, x => x.extra)) + '</td>' : '') + '<td></td></tr></tfoot>' : '') + '</table></div></div>';
    if (showX) h += '<div class="hint" style="padding:6px 4px">الضريبة الإضافية تقديرية: ' + esc(String(cfg.rate)) + '% من المتبقي عن كل شهر تأخير' + (cfg.part ? ' (جزء الشهر يُحسب شهرًا)' : '') + '، وتُراجَع على القانون الساري قبل الاعتماد عليها.</div>';
    return h + '</div>';
  }
  function bindDue() {
    $$('[data-due-sel]').forEach(el => el.addEventListener('change', () => { if (el.checked) VS.dueSel.add(el.dataset.dueSel); else VS.dueSel.delete(el.dataset.dueSel); A.renderPage(); }));
    const all = $('[data-due-all]');
    if (all) all.addEventListener('change', () => { for (const el of $$('[data-due-sel]')) { if (all.checked) VS.dueSel.add(el.dataset.dueSel); else VS.dueSel.delete(el.dataset.dueSel); } A.renderPage(); });
  }

  /* the client's tax account: what each return made due, and every payment against it */
  function statementRows(cid) {
    const pers = Array.from(new Set(rets().filter(r => r.client === cid).map(r => r.period))).sort();
    const rows = []; let bal = 0, dr = 0, cr = 0;
    for (const per of pers) {
      const p = payInfo(retFor(cid, per));
      if (p.st === 'none' && !p.pays.length) continue;
      if (p.st === 'off') continue;
      bal += p.amt; dr += p.amt;
      rows.push({ p, date: p.due, text: 'ضريبة مستحقة · إقرار ' + periodLabel(p.r) + (p.r.kind === 'amended' ? ' (معدل)' : ''), dr: p.amt, cr: 0, bal: V.r2(bal) });
      for (const x of p.pays) { bal -= Number(x.amount) || 0; cr += Number(x.amount) || 0; rows.push({ p, pay: x, date: x.date, text: 'سداد · ' + (x.method || '') + (x.ref ? ' · مرجع ' + x.ref : '') + (x.note ? ' · ' + x.note : ''), dr: 0, cr: Number(x.amount) || 0, bal: V.r2(bal) }); }
    }
    return { rows, dr: V.r2(dr), cr: V.r2(cr), bal: V.r2(bal) };
  }
  function pageTaxStatement(cid) {
    const c = client(cid);
    if (!c) return topBar('شركة غير موجودة') + '<div class="content"><div class="card"><div class="empty"><a href="#vat-due">العودة</a></div></div></div>';
    const st = statementRows(cid), open = dueList({ client: cid });
    ui.stmtClient = cid;
    let h = topBar('كشف الحساب الضريبي · ' + c.name, 'ضريبة القيمة المضافة المستحقة من الإقرارات والمسدد منها',
      '<button type="button" class="btn" data-act="vat-pay-remind" data-client="' + cid + '"' + (open.length ? '' : ' disabled') + '>تذكير بالسداد</button>' +
      (S.downloads ? '<button type="button" class="btn" data-act="vat-rep-pdf" data-key="stmt">⤓ PDF</button><button type="button" class="btn" data-act="vat-rep-xlsx" data-key="stmt">⤓ Excel</button>' : '')) + '<div class="content">';
    h += '<div class="tiles"><div class="tile c1"><span>إجمالي المستحق من الإقرارات</span><b>' + money(st.dr) + '</b></div><div class="tile c3"><span>إجمالي المسدد</span><b>' + money(st.cr) + '</b></div>' +
      '<div class="tile ' + (st.bal > 0 ? 'c2' : 'c4') + '"><span>' + (st.bal > 0 ? 'الرصيد المستحق' : 'الرصيد') + '</span><b>' + money(Math.abs(st.bal)) + '</b><small>' + (open.some(x => x.late > 0) ? 'منه متأخر ' + money(open.filter(x => x.late > 0).reduce((a, x) => a + x.left, 0)) : '') + '</small></div></div>';
    h += '<div class="card"><div class="tablewrap"><table class="vat-tbl"><thead><tr><th>التاريخ</th><th>البيان</th><th class="n">مدين (مستحق)</th><th class="n">دائن (مسدد)</th><th class="n">الرصيد</th><th>الحالة</th><th></th></tr></thead><tbody>';
    if (!st.rows.length) h += '<tr><td colspan="7" class="empty">لا توجد ضريبة مستحقة على هذه الشركة' + (payCfg().from ? ' منذ بداية متابعة السداد' : '') + '.</td></tr>';
    for (const x of st.rows) {
      h += '<tr' + (x.pay ? ' class="sub-row"' : '') + '><td class="num">' + esc(x.date) + '</td><td>' + (x.pay ? esc(x.text) : '<a href="#vat-ret-sum-' + x.p.r.id + '">' + esc(x.text) + '</a>') + '</td><td class="n">' + (x.dr ? money(x.dr) : '') + '</td><td class="n">' + (x.cr ? money(x.cr) : '') + '</td><td class="n">' + money(x.bal) + '</td>' +
        '<td>' + (x.pay ? '' : payPill(x.p.st) + (x.p.late > 0 ? ' <span class="iss ' + (x.p.late > 60 ? 'crit' : 'warn') + '">' + esc(lateText(x.p)) + '</span>' : '')) + '</td>' +
        '<td class="acts">' + (!x.pay && x.p.left > 0 && canPay() ? '<button type="button" class="btn sm primary" data-act="vat-pay" data-id="' + x.p.r.id + '">تسجيل سداد</button>' : '') + '</td></tr>';
    }
    return h + '</tbody>' + (st.rows.length ? '<tfoot><tr><td colspan="2">الإجمالي</td><td class="n">' + money(st.dr) + '</td><td class="n">' + money(st.cr) + '</td><td class="n">' + money(st.bal) + '</td><td colspan="2"></td></tr></tfoot>' : '') + '</table></div></div></div>';
  }
  function dueRep() {
    const cfg = payCfg(), col = A.col, list = dueList({ client: ui.duClient, staff: ui.duStaff }).filter(x => !ui.duAge || (ui.duAge === 'late' ? x.late > 0 : x.age === ui.duAge));
    const rows = list.map(x => [x.c.name, periodLabel(x.r), STATUS[statusOf(x.r)][0], x.due, x.late > 0 ? x.late : '', x.amt, x.paid || '', x.left].concat(cfg.rate ? [x.extra || ''] : []));
    const s = (i) => rows.reduce((a, r) => a + (Number(r[i]) || 0), 0);
    return { title: 'الضريبة المستحقة السداد', ctx: 'حتى ' + C.todayISO() + (ui.duClient ? ' · ' + A.clientName(ui.duClient) : '') + (ui.duAge ? ' · ' + (ui.duAge === 'late' ? 'المتأخر فقط' : AGE.find(a => a[0] === ui.duAge)[1]) : ''), landscape: true,
      cols: [col('الشركة'), col('الفترة'), col('حالة الإقرار'), col('آخر موعد للسداد'), col('أيام التأخير'), col('المستحق', 1), col('المسدد', 1), col('المتبقي', 1)].concat(cfg.rate ? [col('ض. إضافية تقديرية', 1)] : []),
      rows, foot: ['الإجمالي', '', '', '', '', s(5), s(6), s(7)].concat(cfg.rate ? [s(8)] : []) };
  }
  function stmtRep() {
    const cid = ui.stmtClient, st = statementRows(cid), col = A.col;
    return { title: 'كشف الحساب الضريبي – القيمة المضافة', ctx: A.clientName(cid) + ' · حتى ' + C.todayISO(),
      cols: [col('التاريخ'), col('البيان'), col('مدين (مستحق)', 1), col('دائن (مسدد)', 1), col('الرصيد', 1)],
      rows: st.rows.map(x => [x.date, x.text, x.dr || '', x.cr || '', x.bal]), foot: ['الإجمالي', '', st.dr, st.cr, st.bal] };
  }

  function openPay(id) {
    if (!canPay()) return;
    const r0 = retById(id); if (!r0) return;
    const r = retFor(r0.client, r0.period) || r0, p = payInfo(r), c = client(r.client) || {};
    const body = '<form id="vp-form" class="form"><div class="field full"><div class="note info"><b>' + esc(c.name || '') + '</b> · ' + esc(periodLabel(r)) + '<br>المستحق ' + money(p.amt) + ' · المسدد ' + money(p.paid) + ' · <b>المتبقي ' + money(p.left) + '</b> · آخر موعد ' + esc(p.due) + '</div></div>' +
      fld('vp-date', 'تاريخ السداد', inp('vp-date', C.todayISO(), 'date', ' required')) +
      fld('vp-amt', 'المبلغ المسدد', '<input type="text" inputmode="decimal" id="vp-amt" value="' + esc(grp(p.left)) + '" required>') +
      fld('vp-method', 'طريقة السداد', '<select id="vp-method">' + options(PAY_METHODS, PAY_METHODS[0]) + '</select>') +
      fld('vp-ref', 'رقم المرجع / إيصال السداد', inp('vp-ref', '')) + fld('vp-note', 'ملاحظة', inp('vp-note', ''), true) + '<div class="field full err" id="vp-err"></div></form>';
    const bg = modal('تسجيل سداد ضريبة القيمة المضافة', body, '<button type="button" class="btn" data-close>إلغاء</button><button type="submit" form="vp-form" class="btn primary">حفظ السداد</button>');
    const amt = $('#vp-amt', bg);
    amt.addEventListener('blur', () => { const n = numIn(amt.value); if (n) amt.value = grp(n); });
    amt.addEventListener('focus', () => amt.select());
    $('#vp-form', bg).addEventListener('submit', async (ev) => {
      ev.preventDefault();
      const n = numIn(amt.value), err = $('#vp-err', bg);
      if (!(n > 0)) return (err.textContent = 'اكتب المبلغ المسدد.');
      if (!val('vp-date', bg)) return (err.textContent = 'اختر تاريخ السداد.');
      if (n > p.left + 0.004 && p.left > 0 && !bg.dataset.over) { bg.dataset.over = '1'; return (err.textContent = 'المبلغ أكبر من المتبقي (' + money(p.left) + '). اضغط «حفظ السداد» مرة أخرى لو ده مقصود.'); }
      const pay = { id: uid(), date: val('vp-date', bg), amount: n, method: val('vp-method', bg), ref: val('vp-ref', bg), note: val('vp-note', bg), at: new Date().toISOString() };
      if (await write(() => S.db.doc('vatReturns/' + r.id).update({ pay: clean((r.pay || []).concat([pay])) }))) { closeModal(); toast('تم تسجيل السداد'); }
    });
  }
  function delPay(retId, pid) {
    const r = retById(retId); if (!r || !canPay()) return;
    const x = (r.pay || []).find(p => p.id === pid); if (!x) return;
    confirmBox('حذف السداد', 'حذف سداد ' + money(x.amount) + ' بتاريخ ' + x.date + '؟', 'حذف', async () => {
      if (await write(() => S.db.doc('vatReturns/' + r.id).update({ pay: clean((r.pay || []).filter(p => p.id !== pid)) }))) { closeModal(); toast('تم الحذف'); }
    });
  }
  function openBulkPay() {
    if (!canPay()) return;
    const sel = dueList().filter(x => VS.dueSel.has(x.key));
    if (!sel.length) return toast('حدد إقرارًا واحدًا على الأقل.');
    const total = sel.reduce((a, x) => a + x.left, 0);
    const body = '<form id="vb-form" class="form"><div class="field full"><div class="note info">سيُسجَّل سداد <b>المتبقي بالكامل</b> لـ ' + int(sel.length) + ' إقرار، إجمالي ' + money(total) + ' ' + esc(A.cur()) + '.</div></div>' +
      fld('vb-when', 'تاريخ السداد', '<select id="vb-when">' + options([['due', 'آخر موعد لكل إقرار (للإقرارات القديمة)'], ['date', 'تاريخ واحد للكل']], 'due') + '</select>') +
      fld('vb-date', 'التاريخ', inp('vb-date', C.todayISO(), 'date')) +
      fld('vb-method', 'طريقة السداد', '<select id="vb-method">' + options(PAY_METHODS, PAY_METHODS[0]) + '</select>') + fld('vb-ref', 'مرجع / ملاحظة', inp('vb-ref', '')) + '</form>';
    const bg = modal('تسجيل المحدد كمسدد', body, '<button type="button" class="btn" data-close>إلغاء</button><button type="submit" form="vb-form" class="btn primary">تسجيل ' + int(sel.length) + ' سداد</button>');
    const sync = () => { $('#vb-date', bg).closest('.field').style.display = val('vb-when', bg) === 'date' ? '' : 'none'; };
    $('#vb-when', bg).addEventListener('change', sync); sync();
    $('#vb-form', bg).addEventListener('submit', async (ev) => {
      ev.preventDefault();
      const when = val('vb-when', bg), d = val('vb-date', bg), method = val('vb-method', bg), ref = val('vb-ref', bg);
      const ok = await write(async () => {
        for (const x of sel) {
          const r = retById(x.r.id) || x.r;
          await S.db.doc('vatReturns/' + r.id).update({ pay: clean((r.pay || []).concat([{ id: uid(), date: when === 'due' ? x.due : d, amount: x.left, method, ref, note: '', at: new Date().toISOString() }])) });
        }
      });
      if (ok) { VS.dueSel.clear(); closeModal(); toast('تم تسجيل ' + int(sel.length) + ' سداد'); }
    });
  }
  function openPayCfg() {
    const cfg = payCfg(), w = canPay(), dis = w ? '' : ' disabled';
    const body = '<form id="vc-form" class="form">' +
      fld('vc-from', 'بداية متابعة السداد (الإقرارات قبل هذا الشهر لا تظهر كمستحقة)', inp('vc-from', cfg.from, 'month', dis)) +
      fld('vc-rate', 'نسبة الضريبة الإضافية التقديرية عن كل شهر تأخير (%) — اتركها فارغة لإخفائها', inp('vc-rate', cfg.rate || '', 'number', ' step="0.01" min="0"' + dis)) +
      '<div class="field full"><label class="inl"><input type="checkbox" id="vc-part"' + (cfg.part ? ' checked' : '') + dis + '> جزء الشهر يُحسب شهرًا كاملًا</label></div>' +
      '<div class="field full"><div class="note">الضريبة الإضافية في البرنامج تقدير للمتابعة فقط، فراجع النسبة وطريقة الحساب على القانون الساري.</div></div></form>';
    const bg = modal('إعدادات متابعة السداد', body, '<button type="button" class="btn" data-close>إغلاق</button>' + (w ? '<button type="submit" form="vc-form" class="btn primary">حفظ</button>' : ''));
    $('#vc-form', bg).addEventListener('submit', async (ev) => {
      ev.preventDefault();
      await A.saveSettings({ vatPayFrom: val('vc-from', bg), vatLateRate: Number(val('vc-rate', bg)) || 0, vatLatePart: $('#vc-part', bg).checked });
      closeModal();
    });
  }
  function remindText(cid) {
    const c = client(cid) || {}, s = A.settings(), list = dueList({ client: cid }).sort((a, b) => a.r.period.localeCompare(b.r.period));
    const total = list.reduce((a, x) => a + x.left, 0);
    return { c, list, total, text: 'السادة / ' + (c.name || '') + '\nتحية طيبة وبعد،\nنحيطكم علمًا بأن ضريبة القيمة المضافة المستحقة السداد على شركتكم حتى تاريخ ' + C.todayISO() + ' كالتالي:\n' +
      list.map(x => '• ' + periodLabel(x.r) + ': ' + money(x.left) + ' ' + A.cur() + ' – آخر موعد ' + x.due + (x.late > 0 ? '، متأخر ' + x.late + ' يوم' : '')).join('\n') +
      '\nالإجمالي: ' + money(total) + ' ' + A.cur() + '\nنرجو التكرم بالسداد في أقرب وقت تجنبًا لمقابل التأخير، وإرسال ما يفيد السداد للمكتب.\nمع خالص التحية،\n' + (s.officeName || '') + (s.phone ? '\n' + s.phone : '') };
  }
  function openRemind(cid) {
    const t = remindText(cid);
    if (!t.list.length) return toast('لا توجد ضريبة مستحقة على هذه الشركة.');
    const phone = String(t.c.phone || '').replace(/\D/g, '').replace(/^0/, '20');
    const bg = modal('تذكير بسداد ضريبة القيمة المضافة', '<textarea id="vr-text" style="min-height:240px">' + esc(t.text) + '</textarea>',
      '<button type="button" class="btn" data-close>إغلاق</button><button type="button" class="btn" id="vr-copy">نسخ النص</button>' + (S.downloads ? '<button type="button" class="btn" id="vr-pdf">⤓ خطاب PDF</button>' : '') +
      (phone ? '<a class="btn primary" target="_blank" rel="noopener" id="vr-wa" href="#">فتح في واتساب</a>' : ''));
    $('#vr-copy', bg).addEventListener('click', async () => { try { await navigator.clipboard.writeText($('#vr-text', bg).value); toast('تم نسخ النص'); } catch (e) { $('#vr-text', bg).select(); toast('حدد النص وانسخه يدويًا', true); } });
    const wa = $('#vr-wa', bg); if (wa) { const upd = () => { wa.href = 'https://wa.me/' + phone + '?text=' + encodeURIComponent($('#vr-text', bg).value); }; upd(); $('#vr-text', bg).addEventListener('input', upd); }
    const pdf = $('#vr-pdf', bg); if (pdf) pdf.addEventListener('click', () => remindPdf(cid));
  }
  function remindPdf(cid) {
    const t = remindText(cid), s = A.settings();
    const rep = { cols: [A.col('الفترة'), A.col('آخر موعد للسداد'), A.col('أيام التأخير'), A.col('المستحق', 1), A.col('المسدد', 1), A.col('المتبقي', 1)],
      rows: t.list.map(x => [periodLabel(x.r), x.due, x.late > 0 ? x.late : '', x.amt, x.paid || '', x.left]), foot: ['الإجمالي', '', '', t.list.reduce((a, x) => a + x.amt, 0), t.list.reduce((a, x) => a + x.paid, 0), t.total] };
    const inner = A.pdfHead('تذكير بسداد ضريبة القيمة المضافة', t.c.name || '') +
      '<p>السادة / <b>' + esc(t.c.name || '') + '</b></p><p>تحية طيبة وبعد،</p><p>نحيطكم علمًا بأن ضريبة القيمة المضافة المستحقة السداد على شركتكم حتى تاريخ ' + LRM + esc(C.todayISO()) + LRM + ' بيانها كالتالي:</p>' +
      A.tableHtml(rep) + '<p>نرجو التكرم بالسداد في أقرب وقت تجنبًا لمقابل التأخير، وموافاة المكتب بما يفيد السداد.</p><p>وتفضلوا بقبول فائق الاحترام،</p>' +
      '<div class="sign"><div>' + esc(s.officeName || '') + '</div><div>محاسب قانوني</div></div>';
    A.makePdf(A.fileName('تذكير سداد ضريبة ' + (t.c.name || '')) + '.pdf', inner, false);
  }
  /* the payment box on a return's summary tab */
  function payCard(r0) {
    const r = retFor(r0.client, r0.period) || r0, p = payInfo(r);
    if (p.st === 'none' && !p.pays.length) return '';
    let h = '<section class="card pay-card"><header><h2>سداد الضريبة</h2>' + payPill(p.st) + (p.late > 0 ? ' <span class="iss ' + (p.late > 60 ? 'crit' : 'warn') + '">' + esc(lateText(p)) + '</span>' : '') + '<span class="grow"></span>' +
      (p.left > 0 && canPay() ? '<button type="button" class="btn sm primary" data-act="vat-pay" data-id="' + r.id + '">+ تسجيل سداد</button>' : '') +
      '<a class="btn sm" href="#vat-due-' + r.client + '">كشف الحساب الضريبي</a></header><div class="body">';
    if (r.id !== r0.id) h += '<div class="note info" style="margin-bottom:10px">السداد يُتابَع على الإقرار المعدل لنفس الفترة، لأنه المعتمد.</div>';
    if (p.st === 'off') h += '<div class="note info" style="margin-bottom:10px">الفترة قبل «بداية متابعة السداد» في الإعدادات، فلا تظهر في المستحقات.</div>';
    h += '<div class="pay-sum"><div><small>المستحق السداد</small><b>' + money(p.amt) + '</b></div><div><small>المسدد</small><b>' + money(p.paid) + '</b></div><div><small>المتبقي</small><b class="' + (p.left > 0 ? 'match-bad' : 'match-ok') + '">' + money(p.left) + '</b></div>' +
      '<div><small>آخر موعد للسداد</small><b class="num">' + esc(p.due) + '</b></div>' + (p.extra ? '<div><small>ض. إضافية تقديرية</small><b>' + money(p.extra) + '</b></div>' : '') + '</div>';
    if (p.pays.length) h += '<div class="tablewrap" style="margin-top:12px"><table><thead><tr><th>تاريخ السداد</th><th class="n">المبلغ</th><th>الطريقة</th><th>المرجع</th><th>ملاحظة</th><th></th></tr></thead><tbody>' +
      p.pays.map(x => '<tr><td class="num">' + esc(x.date) + '</td><td class="n">' + money(x.amount) + '</td><td>' + esc(x.method || '') + '</td><td class="num">' + esc(x.ref || '') + '</td><td>' + esc(x.note || '') + '</td><td>' +
        (canPay() ? '<button type="button" class="btn sm danger" data-act="vat-pay-del" data-id="' + x.ret + '" data-pid="' + esc(x.id) + '">حذف</button>' : '') + '</td></tr>').join('') + '</tbody></table></div>';
    return h + '</div></section>';
  }

  /* ---------------- old returns (الإقرارات القديمة) ----------------
     Returns that were already submitted on the tax returns system before the office used this program: the e-invoice
     portal's sheets are split by invoice date into the periods, the submitted figures are typed, pasted, or read from the
     returns' PDFs, each period is matched, and the returns are saved as submitted (locked when they match). */
  const H = () => VS.hist || (VS.hist = { client: '', freq: 'monthly', from: '', to: '', invs: [], files: [], figs: {}, pdf: {}, opening: '', accept: {}, paid: true, lock: true, busy: false });
  const hNum = (v) => v === '' || v == null ? null : C.round2(String(v).replace(/[^\d.\-]/g, ''));
  function histPeriods(h) {
    if (!h.client || !/^\d{4}-\d{2}$/.test(h.from) || !/^\d{4}-\d{2}$/.test(h.to) || h.from > h.to) return [];
    const out = []; let p = h.from;
    while (p <= h.to && out.length < 72) { const per = h.freq === 'quarterly' ? V.quarterEnd(p) : p; if (!out.includes(per)) out.push(per); p = C.addMonths(p, 1); }
    return out;
  }
  const histPer = (h, date) => { const ym = String(date || '').slice(0, 7); if (!ym || ym < h.from || ym > h.to) return null; return h.freq === 'quarterly' ? V.quarterEnd(ym) : ym; };
  function histModel(h) {
    const pers = histPeriods(h), by = new Map(pers.map(p => [p, []]));
    let outside = 0;
    for (const inv of h.invs) { const p = histPer(h, inv.date); if (p && by.has(p)) by.get(p).push(inv); else outside++; }
    const before = pers.length ? rets().filter(r => r.client === h.client && r.period < pers[0]).sort((a, b) => b.period.localeCompare(a.period))[0] : null;
    const f0 = h.figs[pers[0]] || {};
    // the credit carried into the first period: the return before it, or the opening credit (by default the first return's own figure)
    let carry = before ? Math.max(0, -resultOf(retFor(h.client, before.period))) : hNum(h.opening) != null ? hNum(h.opening) : hNum(f0.prev) || 0;
    const rows = pers.map(per => {
      const ex = retFor(h.client, per);
      if (ex) {
        const res = resultOf(ex); const x = { per, ex, invs: by.get(per), prevIn: carry, res };
        carry = res < 0 ? -res : 0; return x;
      }
      const invs = by.get(per), calc = V.compute(invs, carry), f = h.figs[per] || {};
      const has = ['out', 'inp', 'res'].some(k => hNum(f[k]) != null);
      let theirs = null;
      if (has) {
        const out = hNum(f.out) || 0, inp = hNum(f.inp) || 0, prev = hNum(f.prev) != null ? hNum(f.prev) : calc.prev, due = V.r2(out - inp);
        theirs = { out, inp, due, prev, res: hNum(f.res) != null ? hNum(f.res) : V.r2(due - prev) };
      }
      const d = theirs ? { out: V.r2(theirs.out - calc.out), inp: V.r2(theirs.inp - calc.inp), prev: V.r2(theirs.prev - calc.prev), res: V.r2(theirs.res - calc.res) } : null;
      const ok = !!d && Object.values(d).every(v => Math.abs(v) < 0.015);
      const res = theirs ? theirs.res : calc.res;
      carry = res < 0 ? -res : 0;
      return { per, invs, calc, f, theirs, d, ok, res, nS: invs.filter(i => i.side === 'S').length, nP: invs.filter(i => i.side === 'P').length };
    });
    rows.forEach((x, i) => { x.hint = histHint(rows, i); });
    return { pers, rows, outside, before };
  }
  function histHint(rows, i) {
    const x = rows[i]; if (!x.d || x.ok) return '';
    const d = x.d, nx = rows[i + 1] && rows[i + 1].d, pv = rows[i - 1] && rows[i - 1].d, cancels = (a, b) => b != null && Math.abs(a + b) < 1;
    if (Object.values(d).every(v => Math.abs(v) < 1)) return 'فرق كسور';
    const out = [];
    if (Math.abs(d.prev) >= 1) out.push('الرصيد السابق مختلف: راجع الفترة اللي قبلها أو الرصيد الافتتاحي');
    if (d.out >= 1) out.push('المقدَّم فيه ضريبة مبيعات أكتر بـ ' + money(d.out) + ': غالبًا مبيعات مستهلك نهائي أو خارج البوابة');
    if (d.out <= -1) out.push('الشيت فيه ضريبة مبيعات أكتر بـ ' + money(-d.out) + ': فواتير اتلغت أو اتسجلت بعد التقديم');
    if (d.inp <= -1) out.push(cancels(d.inp, nx && nx.inp) ? 'مدخلات ' + money(-d.inp) + ' اتخصمت في إقرار الفترة اللي بعدها' : 'مدخلات في الشيت بـ ' + money(-d.inp) + ' ماتخصمتش في الإقرار');
    if (d.inp >= 1) out.push(cancels(d.inp, pv && pv.inp) ? 'مدخلات ' + money(d.inp) + ' من الفترة اللي قبلها اتخصمت هنا' : 'المقدَّم فيه مدخلات أكتر بـ ' + money(d.inp) + ': فواتير ورقية أو جمركية مش في البوابة');
    return out.join(' · ');
  }
  function histSetClient(cid) {
    const h = H(), c = client(cid);
    h.client = cid; h.invs = []; h.files = []; h.figs = {}; h.pdf = {}; h.opening = ''; h.accept = {};
    if (!c) return;
    h.freq = A.lawOf(c) === '6-2025' ? 'quarterly' : 'monthly';
    const first = rets().filter(r => r.client === cid).map(r => r.period).sort()[0];
    h.to = first ? C.addMonths(first, -1) : payCfg().from ? C.addMonths(payCfg().from, -1) : C.addMonths(S.M.ym, -1);
    h.from = h.to.slice(0, 4) + '-01';
  }

  function pageHist() {
    const h = H(), c = client(h.client), w = A.writable();
    let o = topBar('الإقرارات القديمة', 'إضافة إقرارات اتقدمت قبل استخدام البرنامج: من شيتات بوابة الفاتورة الإلكترونية، ومطابقتها مع الإقرارات المقدَّمة', '<a class="btn" href="#vat-list">رجوع للإقرارات</a>') + '<div class="content hist">';
    if (!w) return o + '<div class="card"><div class="empty">هذه الصفحة تحتاج صلاحية التعديل على الإقرارات.</div></div></div>';
    const step = (n, t, hint, body) => '<section class="card hstep"><header><span class="hnum">' + n + '</span><h2>' + esc(t) + '</h2>' + (hint ? '<span class="hint">' + hint + '</span>' : '') + '</header><div class="body">' + body + '</div></section>';
    const vatCl = S.M.clients.filter(x => x.status !== 'inactive').map(x => [x.id, (x.code ? x.code + ' · ' : '') + x.name]);
    o += step('١', 'الشركة والفترة', 'كل فترة ليها إقرار في البرنامج هتظهر «موجود» ومش هتتغير', '<div class="form hform">' +
      fld('h-client', 'الشركة', '<select id="h-client" data-h="client">' + options(vatCl, h.client, 'اختر الشركة…') + '</select>') +
      fld('h-freq', 'دورية الإقرارات في الفترة دي', '<select id="h-freq" data-h="freq"' + (c ? '' : ' disabled') + '>' + options([['monthly', 'شهري'], ['quarterly', 'ربع سنوي']], h.freq) + '</select>') +
      fld('h-from', 'من شهر', inp('h-from', h.from, 'month', ' data-h="from"' + (c ? '' : ' disabled'))) + fld('h-to', 'إلى شهر', inp('h-to', h.to, 'month', ' data-h="to"' + (c ? '' : ' disabled'))) + '</div>');
    const m = histModel(h);
    if (!c || !m.pers.length) return o + '</div>';
    const nIn = h.invs.length - m.outside;
    o += step('٢', 'شيتات بوابة الفاتورة الإلكترونية', 'ارفع ملف الفواتير المرسلة والمستلمة (ممكن أكتر من ملف، ولكل الشهور مرة واحدة)',
      '<div class="toolbar"><label class="btn primary" for="h-files">⤒ رفع الشيتات</label><input type="file" id="h-files" multiple accept=".xlsx,.xls,.csv" hidden>' +
      (h.files.length ? '<button type="button" class="btn sm" data-act="vat-hist-clear">مسح الشيتات</button>' : '') + '<span class="grow"></span>' +
      (h.files.length ? '<span class="hint">' + h.files.map(f => esc(f.name) + ' (' + int(f.n) + ')').join(' · ') + '</span>' : '') + '</div>' +
      (h.files.length ? '<div class="note info" style="margin-top:10px">' + int(nIn) + ' مستند داخل الفترة اتوزعوا على ' + int(m.rows.filter(x => !x.ex && x.invs.length).length) + ' ' + (h.freq === 'quarterly' ? 'ربع' : 'شهر') + ' حسب تاريخ الإصدار' + (m.outside ? ' · ' + int(m.outside) + ' مستند تاريخه خارج الفترة (مش هيتضاف)' : '') + '</div>' : ''));

    // 3: submitted figures and the match, one row per period
    const z = (v) => v == null || v === '' ? '' : String(v);
    const cell = (per, k, v, ph, type) => '<input id="hf-' + per + '-' + k + '" data-hf="' + per + '|' + k + '" value="' + esc(z(v)) + '"' + (type === 'date' ? ' type="date"' : ' inputmode="decimal"') + (ph != null && ph !== '' ? ' placeholder="' + esc(ph) + '"' : '') + ' class="hin' + (type ? ' h' + type : '') + '">';
    let t = '<table class="vat-tbl grouped hist-tbl"><thead><tr class="grp"><th></th><th colspan="3" class="g">حسب الشيتات</th><th colspan="6" class="g">الإقرار المقدَّم</th><th colspan="2"></th></tr><tr>' +
      '<th>الفترة</th><th class="n">ض. المبيعات</th><th class="n">ض. المدخلات</th><th class="n">الصافي</th>' +
      '<th>ض. المبيعات</th><th>ض. المدخلات</th><th>الرصيد السابق</th><th title="موجب = مستحق السداد، سالب = رصيد دائن">الصافي ±</th><th>رقم الإقرار</th><th>تاريخ التقديم</th><th>النتيجة</th><th>فرق مقبول</th></tr></thead><tbody>';
    for (const x of m.rows) {
      const lbl = esc(h.freq === 'quarterly' ? 'الربع المنتهي ' + A.monthLabel(x.per) : A.monthLabel(x.per));
      if (x.ex) {
        t += '<tr class="hx"><td><a href="#vat-ret-sum-' + x.ex.id + '">' + lbl + '</a></td><td colspan="9"><span class="pill mute">موجود في البرنامج</span> ' + pill(statusOf(x.ex)) + ' <span class="hint">الصافي ' + money(x.res) + (x.invs.length ? ' · ' + int(x.invs.length) + ' مستند من الشيت لن يُضاف هنا' : '') + '</span></td><td colspan="2"></td></tr>';
        continue;
      }
      const sres = x.calc.res, f = x.f;
      const resCell = !x.theirs ? '<span class="hint">—</span>' : x.ok ? '<span class="match-ok">✓ مطابق</span>' :
        '<span class="match-bad">✗ فرق ' + money(x.d.res) + '</span>' + (x.hint ? '<div class="hhint">' + esc(x.hint) + '</div>' : '');
      t += '<tr class="' + (x.theirs ? (x.ok ? 'h-ok' : h.accept[x.per] ? 'h-acc' : 'h-bad') : '') + '"><td>' + lbl + (h.pdf[x.per] ? ' <span class="iss info" title="' + esc(h.pdf[x.per].name) + '">PDF</span>' : '') + '<div class="hhint">' + (x.invs.length ? int(x.nS) + ' بيع · ' + int(x.nP) + ' شراء' : 'لا توجد مستندات') + '</div></td>' +
        '<td class="n">' + money(x.calc.out) + '</td><td class="n">' + money(x.calc.inp) + '</td><td class="n">' + money(sres) + '</td>' +
        '<td>' + cell(x.per, 'out', f.out) + '</td><td>' + cell(x.per, 'inp', f.inp) + '</td><td>' + cell(x.per, 'prev', f.prev, x.calc.prev ? grp(x.calc.prev) : '0') + '</td>' +
        '<td>' + cell(x.per, 'res', f.res, x.theirs ? grp(V.r2(x.theirs.due - x.theirs.prev)) : '') + '</td><td>' + cell(x.per, 'no', f.no, '', 'no') + '</td><td>' + cell(x.per, 'date', f.date, '', 'date') + '</td>' +
        '<td class="hres">' + resCell + '</td><td>' + (x.theirs && !x.ok ? '<input type="checkbox" data-hacc="' + x.per + '"' + (h.accept[x.per] ? ' checked' : '') + ' title="اعتماد الفرق كمقبول">' : '') + '</td></tr>';
    }
    t += '</tbody></table>';
    const showOpen = !m.before;
    o += step('٣', 'أرقام الإقرارات المقدَّمة والمطابقة', 'اكتبها من الإقرار، أو الصقها من Excel، أو ارفع ملفات الإقرارات PDF',
      '<div class="toolbar"><label class="btn" for="h-pdfs">⤒ رفع ملفات الإقرارات PDF</label><input type="file" id="h-pdfs" multiple accept=".pdf" hidden><button type="button" class="btn" data-act="vat-hist-paste">📋 لصق من Excel</button>' +
      '<button type="button" class="btn" data-act="vat-hist-copy" title="للفترات الفاضية: اكتب أرقام النظام في خانات المقدَّم">نسخ أرقام النظام للخانات الفاضية</button><span class="grow"></span>' +
      (showOpen ? '<label class="inl">الرصيد الدائن قبل أول فترة <input id="h-open" data-hopen value="' + esc(z(h.opening)) + '" inputmode="decimal" class="hin" placeholder="' + esc(grp(hNum((h.figs[m.pers[0]] || {}).prev) || 0)) + '"></label>' : '<span class="hint">الرصيد السابق لأول فترة من إقرار ' + esc(periodLabel(m.before)) + '</span>') + '</div>' +
      '<div class="tablewrap" style="margin-top:10px">' + t + '</div>');

    const todo = m.rows.filter(x => !x.ex && (x.invs.length || x.theirs));
    const nOk = todo.filter(x => x.ok).length, nAcc = todo.filter(x => x.theirs && !x.ok && h.accept[x.per]).length, nBad = todo.filter(x => x.theirs && !x.ok && !h.accept[x.per]).length, nNo = todo.filter(x => !x.theirs).length;
    o += step('٤', 'الحفظ', '', '<div class="hsum"><span class="match-ok">✓ ' + int(nOk) + ' مطابق</span>' + (nAcc ? '<span>' + int(nAcc) + ' فرق مقبول</span>' : '') + (nBad ? '<span class="match-bad">✗ ' + int(nBad) + ' فيه فرق</span>' : '') + (nNo ? '<span class="hint">' + int(nNo) + ' من غير أرقام مقدَّمة</span>' : '') + '</div>' +
      '<div class="hopts"><label class="inl"><input type="checkbox" data-hopt="paid"' + (h.paid ? ' checked' : '') + '> اعتبار الضريبة المستحقة في الفترات دي مسددة (اتسددت قبل استخدام البرنامج)</label>' +
      '<label class="inl"><input type="checkbox" data-hopt="lock"' + (h.lock ? ' checked' : '') + '> قفل الإقرارات المطابقة والمقبولة من التعديل</label></div>' +
      '<div class="toolbar"><button type="button" class="btn primary" data-act="vat-hist-save"' + (todo.length && !h.busy ? '' : ' disabled') + '>' + (h.busy ? 'جارٍ الحفظ…' : 'حفظ ' + int(todo.length) + ' إقرار') + '</button>' +
      '<span class="hint">الإقرارات بتتحفظ «مقدَّمة» برقمها وتاريخها، والرصيد الدائن بيترحّل من الأرقام المقدَّمة.</span></div>');
    return o + '</div>';
  }
  function bindHist() {
    const h = H();
    $$('[data-h]').forEach(el => el.addEventListener('change', () => {
      const k = el.dataset.h;
      if (k === 'client') histSetClient(el.value); else h[k] = el.value;
      A.renderPage();
    }));
    const later = (fn) => () => { fn(); setTimeout(() => { const a = document.activeElement; VS.histFocus = a && a.id && a.closest('.hist') ? a.id : ''; A.renderPage(); }, 0); };
    $$('[data-hf]').forEach(el => el.addEventListener('change', later(() => {
      const [per, k] = el.dataset.hf.split('|'); h.figs[per] = Object.assign({}, h.figs[per], { [k]: el.value.trim() });
    })));
    const op = $('[data-hopen]'); if (op) op.addEventListener('change', later(() => { h.opening = op.value.trim(); }));
    $$('[data-hacc]').forEach(el => el.addEventListener('change', () => { h.accept[el.dataset.hacc] = el.checked; A.renderPage(); }));
    $$('[data-hopt]').forEach(el => el.addEventListener('change', () => { h[el.dataset.hopt] = el.checked; }));
    const fi = $('#h-files'); if (fi) fi.addEventListener('change', async () => { const fs = Array.from(fi.files); fi.value = ''; for (const f of fs) await histSheet(f); A.renderPage(); });
    const pi = $('#h-pdfs'); if (pi) pi.addEventListener('change', async () => { const fs = Array.from(pi.files); pi.value = ''; await histPdfs(fs); A.renderPage(); });
    if (VS.histFocus) { const el = document.getElementById(VS.histFocus); VS.histFocus = ''; if (el) { el.focus(); try { el.select(); } catch (e) {} } }
  }
  async function histSheet(file) {
    const h = H();
    let book; try { book = await readBook(file); } catch (e) { return toast((e && e.message) || 'تعذّر قراءة الملف.', true); }
    const res = V.readWorkbook(book, memFor(h.client));
    if (!res) return toast('«' + file.name + '»: صيغة غير معروفة. ارفع ملف الفواتير من بوابة الفاتورة الإلكترونية.', true);
    const have = new Set(h.invs.map(V.invKey)); let n = 0;
    for (const inv of res.invoices) { const k = V.invKey(inv); if (have.has(k)) continue; have.add(k); h.invs.push(inv); n++; }
    h.files.push({ name: file.name, n });
    toast('«' + file.name + '»: ' + int(n) + ' مستند');
  }
  async function readPdfFigures(file) {
    const lib = await pdfLib(), bytes = new Uint8Array(await file.arrayBuffer());
    const doc = await lib.getDocument({ data: bytes.slice() }).promise, pages = [];
    for (let p = 1; p <= doc.numPages; p++) { const tc = await (await doc.getPage(p)).getTextContent(); pages.push(tc.items.map(i => i.str)); }
    return { parsed: V.parseReturnPdf(pages), bytes };
  }
  async function histPdfs(files) {
    const h = H(), pers = histPeriods(h); let got = 0; const miss = [];
    toast('جارٍ قراءة ' + int(files.length) + ' ملف…');
    for (const f of files) {
      try {
        const { parsed, bytes } = await readPdfFigures(f), s = parsed.summary;
        const ym = String(parsed.to || '').slice(0, 7), per = ym ? (h.freq === 'quarterly' ? V.quarterEnd(ym) : ym) : '';
        if (!s || !pers.includes(per)) { miss.push(f.name); continue; }
        const z = (v) => v == null ? '' : String(v);
        h.figs[per] = { out: z(s.out), inp: z(s.inp), prev: z(s.prev), res: z(s.res), no: parsed.formNo || '', date: parsed.sentDate || '' };
        h.pdf[per] = { name: f.name, bytes }; got++;
      } catch (e) { miss.push(f.name); }
    }
    toast('اتقرأ ' + int(got) + ' إقرار' + (miss.length ? ' · تعذّر ربط: ' + miss.join('، ') : ''), !!miss.length);
  }
  function parseMonthCell(v) {
    const s = V.toLatin(String(v == null ? '' : v)).trim(), p2 = (x) => String(x).padStart(2, '0');
    let m = s.match(/^(\d{4})[-\/.](\d{1,2})(?:[-\/.]\d{1,2})?$/); if (m) return m[1] + '-' + p2(m[2]);
    m = s.match(/^\d{1,2}[-\/.](\d{1,2})[-\/.](\d{4})$/); if (m) return m[2] + '-' + p2(m[1]);
    m = s.match(/^(\d{1,2})[-\/.](\d{4})$/); if (m) return m[2] + '-' + p2(m[1]);
    const mi = A.MONTHS.findIndex(x => s.includes(x)), y = s.match(/\d{4}/);
    return mi >= 0 && y ? y[0] + '-' + p2(mi + 1) : null;
  }
  function openHistPaste() {
    const h = H();
    const body = '<p style="margin:0 0 8px">انسخ من Excel صف لكل إقرار بالترتيب ده (أي عمود فاضي سيبه فاضي):</p><div class="note info" style="margin-bottom:8px">الشهر · ض. المبيعات · ض. المدخلات · الرصيد السابق · الصافي · رقم الإقرار · تاريخ التقديم</div>' +
      '<textarea id="hp-text" style="min-height:200px;direction:ltr;text-align:left" placeholder="2026-01	15000	9000	0	6000	123456	2026-02-20"></textarea><div class="err" id="hp-err"></div>';
    const bg = modal('لصق أرقام الإقرارات من Excel', body, '<button type="button" class="btn" data-close>إلغاء</button><button type="button" class="btn primary" id="hp-go">إدخال</button>', true);
    $('#hp-go', bg).addEventListener('click', () => {
      const pers = histPeriods(h); let n = 0; const bad = [];
      for (const line of $('#hp-text', bg).value.split(/\r?\n/)) {
        if (!line.trim()) continue;
        const c = line.split('\t'), ym = parseMonthCell(c[0]), per = ym ? (h.freq === 'quarterly' ? V.quarterEnd(ym) : ym) : '';
        if (!pers.includes(per)) { bad.push(c[0]); continue; }
        const nm = (v) => v == null || String(v).trim() === '' ? '' : String(V.num(v));
        const date = V.toLatin(String(c[6] || '').trim()), dm = date.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})$/);
        h.figs[per] = { out: nm(c[1]), inp: nm(c[2]), prev: nm(c[3]), res: nm(c[4]), no: String(c[5] || '').trim(), date: dm ? dm[3] + '-' + dm[2].padStart(2, '0') + '-' + dm[1].padStart(2, '0') : date.slice(0, 10) };
        n++;
      }
      if (!n) return ($('#hp-err', bg).textContent = 'مفيش صفوف اتقرت. أول عمود لازم يكون الشهر (مثلاً 2026-01 أو 01/2026) وداخل الفترة.');
      closeModal(); A.renderPage(); toast('اتدخل ' + int(n) + ' إقرار' + (bad.length ? ' · اتجاهل: ' + bad.slice(0, 5).join('، ') : ''), !!bad.length);
    });
  }
  function histCopySystem() {
    const h = H(), m = histModel(h); let n = 0;
    for (const x of m.rows) if (!x.ex && !x.theirs && x.invs.length) { h.figs[x.per] = Object.assign({}, h.figs[x.per], { out: String(x.calc.out), inp: String(x.calc.inp), prev: String(x.calc.prev), res: String(x.calc.res) }); n++; }
    A.renderPage(); toast(n ? 'اتكتبت أرقام النظام في ' + int(n) + ' فترة · راجعها على الإقرارات المقدَّمة' : 'مفيش فترات فاضية فيها مستندات.');
  }
  async function saveHist() {
    const h = H(), m = histModel(h), c = client(h.client);
    if (!c || h.busy || !A.writable()) return;
    const todo = m.rows.filter(x => !x.ex && (x.invs.length || x.theirs));
    if (!todo.length) return toast('لا توجد فترات جديدة للحفظ.');
    h.busy = true; A.renderPage();
    let n = 0;
    const ok = await write(async () => {
      for (const x of todo) {
        const b = V.periodBounds(x.per, h.freq), id = uid(), now = new Date().toISOString();
        const rows = x.invs.map(inv => { const i = clean(inv); i.id = uid(); fixNo(i); return i; });
        const r = { id, client: c.id, period: x.per, freq: h.freq, from: b.from, to: b.to, kind: 'original', status: 'submitted', hist: true, prevMode: 'auto',
          prevCredit: x === todo[0] && !m.before ? x.calc.prev : 0, chunks: 0, keys: [], createdAt: now, updatedAt: now };
        const parts = chunkRows(rows);
        for (let k = 0; k < parts.length; k++) await S.db.doc('vatInvoices/' + id + '-' + k).set({ ret: id, k, rows: clean(parts[k]) });
        r.chunks = parts.length; r.sum = summaryOf(r, rows); r.keys = rows.filter(i => i.side === 'P' && !i.x).map(V.invKey);
        if (x.theirs) {
          r.sub = { src: 'manual', formNo: x.f.no || '', sentDate: x.f.date || '', kind: 'original', from: b.from, to: b.to, parsed: { summary: x.theirs, counts: null, numbers: [] }, ok: x.ok, at: now };
          if (!x.ok && h.accept[x.per]) r.sub.accepted = { note: x.hint || '', at: now };
          r.locked = !!h.lock && (x.ok || !!h.accept[x.per]);
        }
        if (h.paid && x.res > 0) r.pay = [{ id: uid(), date: C.dueDate(vatOb(), x.per), amount: x.res, method: '', ref: '', note: 'إقرار قديم – اتسدد قبل استخدام البرنامج', at: now }];
        await S.db.doc('vatReturns/' + id).set(clean(r));
        const pdf = h.pdf[x.per];
        if (pdf && r.sub) { const fc = await saveFileBytes(r, pdf.bytes, pdf.name, 0); if (fc) await S.db.doc('vatReturns/' + id).update({ sub: clean(Object.assign({}, r.sub, { viaPdf: true, file: pdf.name, fileChunks: fc, fileSize: pdf.bytes.length })) }); }
        VS.cache.set(id, { rows, loaded: true });
        const t = x.theirs && (S.raw.tasks || []).find(t => t.client === c.id && t.obligation === 'vat' && t.period === x.per && t.status !== 'done');
        if (t) await S.db.doc('tasks/' + t.id).update({ status: 'done', doneDate: x.f.date || C.todayISO(), ref: x.f.no || t.ref || '' });
        n++; toast('جارٍ الحفظ… ' + int(n) + ' / ' + int(todo.length));
      }
    });
    h.busy = false;
    if (ok) { toast('تم حفظ ' + int(n) + ' إقرار'); histSetClient(c.id); ui.rmClient = c.id; ui.rmYear = todo[0].per.slice(0, 4); A.go('vat-rep-months'); }
    else A.renderPage();
  }

  /* ---------------- list ---------------- */
  function pageList() {
    const all = rets(), years = Array.from(new Set(all.map(r => r.period.slice(0, 4)))).sort().reverse();
    const list = all.filter(r => (!ui.vClient || r.client === ui.vClient) && (!ui.vStatus || statusOf(r) === ui.vStatus) && (!ui.vYear || r.period.startsWith(ui.vYear)))
      .sort((a, b) => b.period.localeCompare(a.period) || A.clientName(a.client).localeCompare(A.clientName(b.client), 'ar'));
    const cnt = (k) => all.filter(r => statusOf(r) === k).length;
    let h = topBar('إقرارات القيمة المضافة', 'نموذج 10 · من فواتير البيع والشراء إلى ملفات الرفع ثم مطابقة الإقرار المقدَّم', (A.writable() ? '<a class="btn" href="#vat-hist">⤒ إقرارات قديمة</a>' : '') + actBtn('vat-new', '+ إقرار جديد', 'primary')) + '<div class="content">';
    h += '<div class="tiles">' +
      '<a class="tile c5" href="#" data-act="vat-filter" data-v="draft"><span>مسودات</span><b>' + int(cnt('draft')) + '</b><small>تحت الإعداد</small></a>' +
      '<a class="tile c4" href="#" data-act="vat-filter" data-v="ready"><span>جاهزة للتقديم</span><b>' + int(cnt('ready')) + '</b><small>نُزّلت ملفات الرفع</small></a>' +
      '<a class="tile c3" href="#" data-act="vat-filter" data-v="matched"><span>مقدَّمة ومطابقة</span><b>' + int(cnt('matched') + cnt('accepted')) + '</b><small>طابقت الإقرار المقدَّم</small></a>' +
      '<a class="tile c2" href="#" data-act="vat-filter" data-v="diff"><span>مقدَّمة وبها فروق</span><b>' + int(cnt('diff')) + '</b><small>راجع المطابقة</small></a></div>';
    h += '<div class="toolbar"><select data-ui="vClient">' + options(A.clientOptions(true), ui.vClient, 'جميع الشركات') + '</select>' +
      '<select data-ui="vStatus">' + options(Object.entries(STATUS).map(([k, v]) => [k, v[0]]), ui.vStatus, 'جميع الحالات') + '</select>' +
      '<select data-ui="vYear">' + options(years, ui.vYear, 'جميع السنوات') + '</select></div>';
    h += '<div class="card"><div class="tablewrap"><table class="vat-tbl"><thead><tr><th>الشركة</th><th>الفترة</th><th>الدورية</th><th class="n">المبيعات</th><th class="n">ضريبة المبيعات</th><th class="n">ضريبة المدخلات</th><th class="n">الرصيد الدائن السابق</th><th class="n">مستحق السداد</th><th class="n">رصيد دائن مرحّل</th><th>الحالة</th><th>رقم الإقرار</th><th>تاريخ التقديم</th></tr></thead><tbody>';
    if (!list.length) h += '<tr><td colspan="12" class="empty">' + (all.length ? 'لا توجد إقرارات بهذه الشروط' : 'لا توجد إقرارات بعد. ابدأ بـ «+ إقرار جديد» واختر الشركة والفترة.') + '</td></tr>';
    for (const r of list) {
      const s = r.sum || {}, prev = prevCredit(r), res = V.r2((s.due || 0) - prev), st = statusOf(r);
      h += '<tr><td><a href="#vat-ret-sum-' + r.id + '">' + esc(A.clientName(r.client)) + '</a>' + (r.kind === 'amended' ? ' <span class="pill warn">معدل</span>' : '') + '</td><td>' + esc(periodLabel(r)) + '</td><td>' + esc(FREQ[r.freq] || '') + '</td>' +
        '<td class="n">' + money(s.sales) + '</td><td class="n">' + money(s.out) + '</td><td class="n">' + money(s.inp) + '</td><td class="n">' + money(prev) + '</td>' +
        '<td class="n pos">' + (res > 0 ? money(res) + (() => { const pi = payInfo(retFor(r.client, r.period) || r); return pi.st === 'paid' ? ' <span class="iss ok" title="مسدد">✓ مسدد</span>' : pi.st === 'partial' ? ' <span class="iss warn">جزئي</span>' : pi.st === 'unpaid' && pi.late > 0 ? ' <span class="iss crit">متأخر</span>' : ''; })() : '') + '</td><td class="n neg">' + (res < 0 ? money(-res) : '') + '</td><td>' + pill(st) + (s.crit ? ' <span class="iss crit">' + int(s.crit) + ' ملاحظة</span>' : '') + '</td><td class="num">' + (r.sub && r.sub.formNo ? (r.sub.fileChunks ? '<a href="#" data-act="vat-dl-sub" data-id="' + r.id + '" title="تنزيل الإقرار المقدَّم PDF">' + esc(r.sub.formNo) + ' ⤓</a>' : '<a href="#vat-ret-match-' + r.id + '" title="لم يُحفظ ملف هذا الإقرار؛ اضغط لرفعه">' + esc(r.sub.formNo) + ' <span class="iss warn">ارفع الملف</span></a>') : '') + '</td><td class="num">' + esc(r.sub ? r.sub.sentDate || '' : '') + '</td></tr>';
    }
    return h + '</tbody></table></div></div></div>';
  }

  function openNew(preset) {
    preset = preset || {};
    const cl = S.M.clients.filter(c => c.status !== 'inactive');
    const vatFirst = cl.filter(c => c.vatRegistered || (c.obligations || []).includes('vat'));
    const opts = (vatFirst.length ? vatFirst : cl).map(c => [c.id, (c.code ? c.code + ' · ' : '') + c.name]);
    const last = C.addMonths(S.M.ym, -1);
    const body = '<form id="vn-form" class="form">' +
      fld('vn-client', 'الشركة', '<select id="vn-client" required>' + options(opts, preset.client || ui.vClient, 'اختر الشركة…') + '</select>', true) +
      fld('vn-period', 'الشهر (للربع سنوي: أي شهر في الربع)', inp('vn-period', preset.period || last, 'month')) +
      fld('vn-kind', 'نوع الإقرار', '<select id="vn-kind">' + options([['original', 'أصلي'], ['amended', 'معدل']], 'original') + '</select>') +
      '<div class="field full"><div class="note info" id="vn-note"></div></div>' +
      fld('vn-prev', 'الرصيد الدائن السابق (يُستخدم فقط إذا لم يوجد إقرار سابق مسجل للشركة)', inp('vn-prev', '', 'number', ' step="0.01" min="0"')) +
      '<div class="field full err" id="vn-err"></div></form>';
    const bg = modal('إقرار قيمة مضافة جديد', body, '<button type="button" class="btn" data-close>إلغاء</button><button type="submit" form="vn-form" class="btn primary">إنشاء</button>');
    const info = () => {
      const c = client(val('vn-client', bg)), per = val('vn-period', bg) || last;
      if (!c) { $('#vn-note', bg).textContent = 'اختر الشركة.'; return null; }
      const freq = A.lawOf(c) === '6-2025' ? 'quarterly' : 'monthly', period = freq === 'quarterly' ? V.quarterEnd(per) : per, b = V.periodBounds(period, freq);
      const draft = { client: c.id, period }, p = prevReturn(draft);
      $('#vn-note', bg).textContent = A.TAX_LAW[A.lawOf(c)] + ' · إقرار ' + FREQ[freq] + ' · من ' + b.from + ' إلى ' + b.to + ' · ' +
        (p ? 'الرصيد الدائن السابق يُحسب تلقائيًا من إقرار ' + periodLabel(p) : 'لا يوجد إقرار سابق مسجل لهذه الشركة؛ أدخل الرصيد الدائن السابق إن وُجد.');
      $('#vn-prev', bg).closest('.field').style.display = p ? 'none' : '';
      return { c, freq, period, b, p };
    };
    $('#vn-client', bg).addEventListener('change', info); $('#vn-period', bg).addEventListener('change', info); info();
    $('#vn-form', bg).addEventListener('submit', async (ev) => {
      ev.preventDefault();
      const x = info(), err = $('#vn-err', bg); if (!x) return (err.textContent = 'اختر الشركة.');
      const kind = val('vn-kind', bg);
      const dup = rets().find(r => r.client === x.c.id && r.period === x.period && r.kind === kind && kind === 'original');
      if (dup) return (err.textContent = 'يوجد إقرار أصلي للشركة والفترة نفسيهما. افتحه من القائمة، أو اختر «معدل».');
      const n = { id: uid(), client: x.c.id, period: x.period, freq: x.freq, from: x.b.from, to: x.b.to, kind, status: 'draft', prevMode: 'auto', prevCredit: x.p ? 0 : C.round2(val('vn-prev', bg)),
        chunks: 0, sum: { out: 0, inp: 0, due: 0, sales: 0, inputs: 0, crit: 0, nS: 0, nP: 0 }, keys: [], createdAt: new Date().toISOString(), updatedAt: new Date().toISOString() };
      if (await write(() => S.db.doc('vatReturns/' + n.id).set(n))) { VS.cache.set(n.id, { rows: [], loaded: true }); closeModal(); A.go('vat-ret-sum-' + n.id); }
    });
  }

  /* ---------------- return workspace ---------------- */
  const TABS = [['sum', 'ملخص الإقرار'], ['S', 'المبيعات'], ['P', 'المشتريات والمصروفات'], ['files', 'ملفات الرفع'], ['match', 'المطابقة بعد التقديم']];
  function pageRet(id, tab) {
    const r = retById(id);
    if (!r) return topBar('إقرار غير موجود') + '<div class="content"><div class="card"><div class="empty">الإقرار غير موجود أو حُذف. <a href="#vat-list">العودة إلى القائمة</a></div></div></div>';
    const e = VS.cache.get(id);
    if (!e) { loadRet(r); return topBar(A.clientName(r.client)) + '<div class="loading">جارٍ تحميل فواتير الإقرار…</div>'; }
    if (!e.loaded) return topBar(A.clientName(r.client)) + '<div class="loading">جارٍ تحميل فواتير الإقرار…</div>';
    const rows = e.rows, c = client(r.client) || {};
    const calc = V.compute(rows, prevCredit(r));
    const iss = new Map(rows.map(inv => [inv.id, V.issues(inv, Object.assign({}, r, { defAddr: partyDoc(r.client).defAddr }))]));
    const nCrit = (side) => rows.filter(i => i.side === side && iss.get(i.id).some(x => x[0] === 'crit')).length;
    const st = VS.undo.get(id) || [];
    const actions = (A.writable() && st.length ? '<button type="button" class="btn" data-act="vat-undo" title="' + esc(st[st.length - 1].label) + '">↶ تراجع</button>' : '') +
      (A.writable() ? '<button type="button" class="btn primary" data-act="vat-import">⤒ استيراد ملف Excel</button><input type="file" id="vat-file" accept=".xlsx,.xls,.csv" hidden>' : '') +
      (S.downloads ? '<button type="button" class="btn" data-act="vat-pdf-sum" data-id="' + r.id + '">⤓ ملخص PDF</button>' : '') +
      (r.sub && r.sub.fileChunks ? '<button type="button" class="btn" data-act="vat-view-sub">📄 الإقرار المقدَّم</button>' : '') + actBtn('vat-del-ret', 'حذف الإقرار', 'danger', ' data-id="' + r.id + '"');
    let h = topBar(c.name || '—', esc(periodLabel(r)) + ' · ' + esc(FREQ[r.freq]) + ' (' + esc(r.from) + ' ← ' + esc(r.to) + ') · ' + (r.kind === 'amended' ? 'معدل' : 'أصلي') + ' · ' + pill(statusOf(r)) + (r.locked ? ' <span class="pill mute">🔒 مقفول</span>' : '') + ' <span class="savestate" id="vat-save">' + esc(VS.saveText || '') + '</span>', actions) + '<div class="content' + (r.locked ? ' ret-locked' : '') + '">';
    if (r.locked) h += '<div class="note info">🔒 الإقرار مقفول بعد المطابقة، فالفواتير والأرقام للعرض فقط.' + (A.writable() ? ' <a href="#" data-act="vat-lock" data-id="' + r.id + '" data-v="0">فتح القفل للتعديل</a>' : '') + '</div>';
    h += '<div class="tabs">' + TABS.map(([k, l]) => {
      const n = k === 'S' || k === 'P' ? rows.filter(i => i.side === k).length : 0, bad = k === 'S' || k === 'P' ? nCrit(k) : 0;
      return '<a href="#vat-ret-' + k + '-' + r.id + '" class="' + (tab === k ? 'on' : '') + '">' + esc(l) + (n ? ' <small>(' + int(n) + ')</small>' : '') + (bad ? '<span class="count">' + bad + '</span>' : '') + '</a>';
    }).join('') + '</div>';
    if (tab === 'S' || tab === 'P') h += invoicesTab(r, rows, tab, iss);
    else if (tab === 'files') h += filesTab(r, rows, iss);
    else if (tab === 'match') h += matchTab(r, rows, calc);
    else h += summaryTab(r, rows, calc, nCrit('S') + nCrit('P'));
    return h + '</div>';
  }

  function form10(calc) {
    let h = '<section class="card"><header><h2>المبيعات</h2><span class="hint">ضريبة القيمة المضافة (1)</span></header><div class="tablewrap"><table class="vat-form10"><thead><tr><th>فئة الضريبة</th><th>قيمة السلع</th><th>قيمة الخدمات</th><th>تسويات (+)</th><th>تسويات (−)</th><th>إجمالي القيمة</th><th>الضريبة</th></tr></thead><tbody>';
    for (const [k, l] of V.RATE_ROWS) { const x = calc.sales[k]; h += '<tr><td>' + l + '</td><td class="n">' + money(x.g) + '</td><td class="n">' + money(x.s) + '</td><td class="n">' + money(x.ap) + '</td><td class="n">' + money(x.am) + '</td><td class="n">' + money(x.total) + '</td><td class="n">' + money(x.tax) + '</td></tr>'; }
    const st = calc.salesTotal;
    h += '</tbody><tfoot><tr><td>الإجمالي</td><td class="n">' + money(st.g) + '</td><td class="n">' + money(st.s) + '</td><td class="n">' + money(st.ap) + '</td><td class="n">' + money(st.am) + '</td><td class="n">' + money(st.total) + '</td><td class="n">' + money(st.tax) + '</td></tr></tfoot></table></div></section>';
    h += '<section class="card"><header><h2>المدخلات</h2><span class="hint">ضريبة المدخلات (2) · المصروفات لا تدخل هنا</span></header><div class="tablewrap"><table class="vat-form10"><thead><tr><th>البيان</th><th>محلي</th><th>مستورد</th><th>تسويات (+)</th><th>تسويات (−)</th><th>إجمالي القيمة</th><th>الضريبة</th></tr></thead><tbody>';
    for (const [k, l] of V.IN_ROWS) { const x = calc.inputs[k]; h += '<tr><td>' + esc(l) + '</td><td class="n">' + money(x.loc) + '</td><td class="n">' + money(x.imp) + '</td><td class="n">' + money(x.ap) + '</td><td class="n">' + money(x.am) + '</td><td class="n">' + money(x.total) + '</td><td class="n">' + money(x.tax) + '</td></tr>'; }
    const it = calc.inputsTotal;
    h += '</tbody><tfoot><tr><td>الإجمالي</td><td class="n">' + money(it.loc) + '</td><td class="n">' + money(it.imp) + '</td><td class="n">' + money(it.ap) + '</td><td class="n">' + money(it.am) + '</td><td class="n">' + money(it.total) + '</td><td class="n">' + money(it.tax) + '</td></tr></tfoot></table></div></section>';
    return h;
  }
  /* the return by kind. Saved with its summary; returns saved before this was added get it once their invoices are read */
  const BRQ = { q: [], busy: 0, seen: new Set() };
  function brOf(r) {
    if (r.sum && r.sum.br) return r.sum.br;
    const e = VS.cache.get(r.id);
    if (e && e.loaded) return V.breakdown(e.rows);
    const m = r.sub && r.sub.src === 'manual' && r.sub.parsed && r.sub.parsed.summary && r.sub.parsed.summary.br;
    if (m) return m;
    if (!(r.chunks > 0)) return null;           // no invoices: nothing to split
    if (!BRQ.seen.has(r.id)) { BRQ.seen.add(r.id); BRQ.q.push(r); pumpBr(); }
    return undefined;                          // being read
  }
  function pumpBr() {
    while (BRQ.busy < 3 && BRQ.q.length) {
      const r = BRQ.q.shift(); BRQ.busy++;
      Promise.resolve(loadRet(r)).then(async (e) => {
        const cur = retById(r.id);
        if (e && e.loaded && cur && cur.sum && !cur.sum.br && A.writable()) {
          try { await S.db.doc('vatReturns/' + r.id).update({ sum: clean(Object.assign({}, cur.sum, { br: V.breakdown(e.rows) })) }); } catch (err) { /* shown from memory until the return is saved */ }
        }
      }).catch(() => {}).then(() => { BRQ.busy--; pumpBr(); A.scheduleRender(); });
    }
  }
  function brTables(b) {
    const cell = (v) => '<td class="n">' + (v ? money(v) : '—') + '</td>';
    const sK = V.BR_S.filter(([k]) => k !== 's0' || b.s0);
    const pct = (v, t) => t > 0 && v > 0 ? Math.round(v * 1000 / t) / 10 : 0;
    const mix = b.sAll > 0 ? '<div class="br-mix" title="توزيع المبيعات">' + sK.map(([k, l], i) => pct(b[k], b.sAll) ? '<i class="br-' + k + '" style="width:' + pct(b[k], b.sAll) + '%" title="' + esc(l + ' ' + pct(b[k], b.sAll) + '%') + '"></i>' : '').join('') + '</div>' : '';
    return '<div class="br-wrap"><div><table class="br-tbl"><thead><tr><th>المبيعات</th><th class="n">القيمة</th><th class="n">النسبة</th></tr></thead><tbody>' +
      sK.map(([k, l]) => '<tr><td><i class="br-key br-' + k + '"></i>' + esc(l) + '</td>' + cell(b[k]) + '<td class="n">' + (pct(b[k], b.sAll) ? pct(b[k], b.sAll) + '%' : '') + '</td></tr>').join('') +
      '</tbody><tfoot><tr><td>إجمالي المبيعات</td>' + cell(b.sAll) + '<td></td></tr><tr><td>ضريبة المبيعات</td>' + cell(b.sVat) + '<td></td></tr>' + (b.sTbl ? '<tr><td>ضريبة الجدول</td>' + cell(b.sTbl) + '<td></td></tr>' : '') + '</tfoot></table>' + mix + '</div>' +
      '<div><table class="br-tbl"><thead><tr><th>المشتريات</th><th class="n">القيمة</th><th class="n">النسبة</th></tr></thead><tbody>' +
      V.BR_P.map(([k, l]) => '<tr><td>' + esc(l) + '</td>' + cell(b[k]) + '<td class="n">' + (pct(b[k], b.pAll) ? pct(b[k], b.pAll) + '%' : '') + '</td></tr>').join('') +
      '</tbody><tfoot><tr><td>إجمالي المشتريات</td>' + cell(b.pAll) + '<td></td></tr><tr><td>ضريبة المدخلات</td>' + cell(b.pVat) + '<td></td></tr></tfoot></table></div></div>';
  }
  function summaryTab(r, rows, calc, crit) {
    const prevAuto = r.prevMode !== 'manual', p = prevReturn(r);
    let h = '';
    if (!rows.length) h += '<div class="note info">الإقرار فارغ. استورد ملف التعاملات (Excel الفواتير المرسلة والمستلمة من منظومة الفاتورة الإلكترونية) من زر «استيراد ملف Excel»، أو أضف الفواتير يدويًا من تبويب المبيعات أو المشتريات.</div>';
    if (crit) h += '<div class="note" style="background:var(--crit-soft)"><b style="color:var(--crit)">' + int(crit) + ' فاتورة عليها ملاحظات يجب تصحيحها قبل إعداد ملفات الرفع.</b> افتح تبويب المبيعات أو المشتريات واختر «عليها ملاحظات لازمة».</div>';
    h += '<div class="tiles">' +
      '<div class="tile c1"><span>ضريبة المبيعات (1)</span><b>' + money(calc.out) + '</b><small>وعاء ' + money(calc.salesTotal.total) + '</small></div>' +
      '<div class="tile c4"><span>ضريبة المدخلات (2)</span><b>' + money(calc.inp) + '</b><small>وعاء ' + money(calc.inputsTotal.total) + '</small></div>' +
      '<div class="tile c5"><span>الضريبة المستحقة (1) − (2)</span><b>' + money(calc.due) + '</b><small>قبل الرصيد السابق</small></div>' +
      '<div class="tile c3"><span>الرصيد الدائن السابق</span><b>' + money(calc.prev) + '</b><small>' + (prevAuto ? (p ? 'تلقائي من ' + esc(periodLabel(p)) : 'رصيد افتتاحي') : 'مُدخل يدويًا') + '</small></div>' +
      '<div class="tile ' + (calc.res > 0 ? 'c2' : 'c3') + '"><span>' + (calc.res > 0 ? 'مستحق السداد' : 'رصيد دائن يُرحَّل') + '</span><b>' + money(Math.abs(calc.res)) + '</b><small>' + esc(A.cur()) + '</small></div></div>';
    if (A.writable()) h += '<div class="toolbar"><label class="inl">الرصيد الدائن السابق <input type="number" step="0.01" min="0" id="vat-prev" value="' + esc(calc.prev) + '" style="width:140px"></label>' +
      (prevAuto ? '<span class="hint">' + (p ? 'يتغير تلقائيًا مع الإقرار السابق، وإذا أدخلت رقمًا هنا يُثبَّت.' : 'لا يوجد إقرار سابق مسجل؛ أدخل الرصيد الافتتاحي.') + '</span>' : '<button type="button" class="btn sm" data-act="vat-prev-auto" data-id="' + r.id + '">إعادته إلى الحساب التلقائي</button>') + '</div>';
    h += payCard(r);
    if (rows.length) h += '<section class="card"><header><h2>المبيعات والمشتريات حسب النوع</h2><span class="hint">قيم صافية بعد إشعارات الإضافة والخصم · المصروفات خارج الإقرار</span></header><div class="body">' + brTables(V.breakdown(rows)) + '</div></section>';
    h += form10(calc);
    const k = calc.counts;
    h += '<section class="card"><header><h2>المستندات</h2></header><div class="body"><div class="kv">' +
      '<span>فواتير ضريبية (مبيعات)</span><span>' + int(k.sInv) + (calc.sFirst ? ' · من ' + esc(calc.sFirst) + ' إلى ' + esc(calc.sLast) : '') + '</span>' +
      '<span>مستندات مستهلك نهائي وأخرى</span><span>' + int(k.sOther) + '</span><span>إشعارات إضافة / خصم (مبيعات)</span><span>' + int(k.sDebit) + ' / ' + int(k.sCredit) + '</span>' +
      '<span>فواتير الشراء</span><span>' + int(k.pInv) + '</span><span>أذون الإفراج</span><span>' + int(k.pRelease) + '</span>' +
      '<span>إشعارات إضافة / خصم (مشتريات)</span><span>' + int(k.pDebit) + ' / ' + int(k.pCredit) + '</span>' +
      '<span>فواتير مصروفات (خارج الإقرار)</span><span>' + int(k.expOnly || 0) + '</span><span>مستبعدة</span><span>' + int(k.excluded) + '</span></div></div></section>';
    return h;
  }

  function lineSel(inv) {
    // one select for the whole invoice: shows the shared value or «مختلط»
    const key = inv.side === 'S' ? 'n' : 'o';
    const vals = Array.from(new Set((inv.lines || []).map(l => l[key])));
    const list = inv.side === 'S' ? [['G', 'سلعة'], ['S', 'خدمة']] : P_O;
    if (!A.writable()) return esc((list.find(x => String(x[0]) === String(vals[0])) || ['', vals.length > 1 ? 'مختلط' : ''])[1]);
    return '<select data-vat-set="' + key + '" data-inv="' + inv.id + '">' + (vals.length > 1 ? '<option value="">مختلط</option>' : '') + options(list, vals.length === 1 ? vals[0] : '') + '</select>';
  }
  function useSel(inv) {
    const ex = (inv.lines || []).map(l => !!l.e), all = ex.every(Boolean), none = !ex.some(Boolean);
    if (!A.writable()) return all ? 'مصروفات' : none ? 'مشتريات' : 'مختلط';
    return '<select data-vat-set="e" data-inv="' + inv.id + '">' + (!all && !none ? '<option value="">مختلط</option>' : '') + options([['0', 'مشتريات'], ['1', 'مصروفات']], all ? '1' : none ? '0' : '') + '</select>';
  }
  function invoicesTab(r, rows, side, iss) {
    const q = ui.vQ.trim().toLowerCase();
    let list = rows.filter(i => i.side === side);
    if (ui.vF === 'issues') list = list.filter(i => iss.get(i.id).some(x => x[0] === 'crit'));
    else if (ui.vF === 'warn') list = list.filter(i => iss.get(i.id).length);
    else if (ui.vF === 'x') list = list.filter(i => i.x);
    else if (ui.vF === 'exp') list = list.filter(i => (i.lines || []).some(l => l.e));
    else if (ui.vF === 'out') list = list.filter(i => i.date && (i.date < r.from || i.date > r.to));
    if (q) list = list.filter(i => [i.no, i.name, i.reg, i.nid].some(v => String(v || '').toLowerCase().includes(q)));
    list.sort((a, b) => (a.date || '').localeCompare(b.date || '') || String(a.no).localeCompare(String(b.no), 'en', { numeric: true }));
    const filters = [['', 'الكل'], ['issues', 'عليها ملاحظات لازمة'], ['warn', 'أي ملاحظة'], ['out', 'خارج الفترة'], ['x', 'مستبعدة']].concat(side === 'P' ? [['exp', 'مصروفات']] : []);
    let h = '<div class="toolbar"><input type="search" placeholder="بحث بالرقم أو الاسم أو التسجيل…" data-ui="vQ" data-live value="' + esc(ui.vQ) + '"><select data-ui="vF">' + options(filters, ui.vF) + '</select>' +
      (A.writable() ? '<label class="inl"><input type="checkbox" data-ui="vSame"' + (ui.vSame ? ' checked' : '') + '> تطبيق التصنيف على جميع فواتير ' + (side === 'P' ? 'المورد' : 'العميل') + ' نفسه' + '</label>' : '') + '<span class="grow"></span>' +
      (side === 'S' ? actBtn('vat-multi', '+ مبيعات مستهلك نهائي') + actBtn('vat-add-inv', '+ فاتورة بيع', 'primary', ' data-side="S" data-dt="1"') : actBtn('vat-add-inv', '+ فاتورة شراء', 'primary', ' data-side="P" data-dt="1"')) + '</div>';
    if (side === 'P') h += '<div class="note info">حدّد لكل فاتورة: «مشتريات» تدخل الإقرار وتُخصم ضريبتها، أو «مصروفات» تبقى في دفتر المشتريات وتُستخدم لاحقًا في نموذج 41، ولا تدخل الإقرار ولا ملف الرفع. إذا تضمنت الفاتورة النوعين فافتحها وحدّد نوع كل بند.</div>';
    if (VS.selRet !== r.id) { VS.sel.clear(); VS.selRet = r.id; }
    for (const id of Array.from(VS.sel)) if (!rows.some(i => i.id === id && i.side === side)) VS.sel.delete(id);
    const W = A.writable(), allSel = list.length && list.every(i => VS.sel.has(i.id));
    const ck = (id, on) => W ? '<td style="width:28px"><input type="checkbox" data-vat-sel="' + id + '"' + (on ? ' checked' : '') + ' aria-label="تحديد"></td>' : '';
    if (W && VS.sel.size) {
      const nSel = VS.sel.size;
      h += '<div class="bulk"><b>' + int(nSel) + ' محدد</b>' +
        '<button type="button" class="btn sm" data-act="vat-bulk" data-op="x1">استبعاد</button><button type="button" class="btn sm" data-act="vat-bulk" data-op="x0">إرجاع للإقرار</button>' +
        (side === 'P' ? '<button type="button" class="btn sm" data-act="vat-bulk" data-op="e0">مشتريات</button><button type="button" class="btn sm" data-act="vat-bulk" data-op="e1">مصروفات</button>' +
          '<select data-vat-bulk="o"><option value="">نوع السلعة…</option>' + options(P_O, '') + '</select>' :
          '<select data-vat-bulk="n"><option value="">سلعة / خدمة…</option>' + options([['G', 'سلعة'], ['S', 'خدمة']], '') + '</select><select data-vat-bulk="o"><option value="">نوع السلعة…</option>' + options(S_O, '') + '</select>') +
        '<select data-vat-bulk="rate"><option value="">نسبة الضريبة…</option>' + options(RATES, '') + '</select>' +
        '<button type="button" class="btn sm danger" data-act="vat-bulk" data-op="del">حذف</button>' +
        '<span class="grow"></span><button type="button" class="btn sm" data-act="vat-bulk" data-op="clear">إلغاء التحديد</button></div>';
    }
    const head = (W ? '<th><input type="checkbox" data-vat-selall="' + side + '"' + (allSel ? ' checked' : '') + ' aria-label="تحديد الكل"></th>' : '') + (side === 'S' ? '<th>النوع</th><th>الرقم</th><th>التاريخ</th><th>العميل</th><th>التسجيل / القومي</th><th>سلعة / خدمة</th><th class="n">الصافي</th><th class="n">الضريبة</th><th class="n">الإجمالي</th><th>ملاحظات</th><th></th>'
      : '<th>النوع</th><th>الرقم</th><th>التاريخ</th><th>المورد</th><th>التسجيل</th><th>نوع السلعة</th><th>الاستخدام</th><th class="n">الصافي</th><th class="n">الضريبة</th><th class="n">الإجمالي</th><th class="n">خصم وإضافة</th><th>ملاحظات</th><th></th>');
    h += '<div class="card"><div class="tablewrap"><table class="vat-tbl"><thead><tr>' + head + '</tr></thead><tbody>';
    if (!list.length) h += '<tr><td colspan="13" class="empty">لا توجد فواتير' + (ui.vF || q ? ' بهذه الشروط' : '') + '</td></tr>';
    let tn = 0, tv = 0, tw = 0;
    const shown = list.slice(0, 600);
    for (const inv of list) { const t = V.invTotals(inv); if (!inv.x) { const sg = inv.dt === 3 ? -1 : 1; tn += sg * t.net; tv += sg * t.vat; tw += Number(inv.wht) || 0; } }
    for (const inv of shown) {
      const t = V.invTotals(inv), notes = (inv.noOrig ? '<span class="iss info">الرقم الأصلي: ' + esc(inv.noOrig) + '</span>' : '') + iss.get(inv.id).map(x => '<span class="iss ' + x[0] + '">' + esc(x[1]) + '</span>').join('') + (inv.x ? '<span class="iss warn">مستبعدة' + (inv.xr ? ': ' + esc(inv.xr) : '') + '</span>' : '');
      const btns = '<button type="button" class="btn sm" data-act="vat-edit-inv" data-id="' + inv.id + '">' + (A.writable() ? 'تعديل' : 'عرض') + '</button>' + (A.writable() ? ' <button type="button" class="btn sm" data-act="vat-toggle-x" data-id="' + inv.id + '">' + (inv.x ? 'إرجاع' : 'استبعاد') + '</button>' : '');
      h += '<tr class="' + (inv.x ? 'x' : '') + '">' + ck(inv.id, VS.sel.has(inv.id)) + '<td>' + esc(dtLabel(side, inv.dt)) + '</td><td class="num">' + esc(inv.no) + '</td><td class="num">' + esc(inv.date) + '</td><td class="wrap">' + esc(inv.name) + '</td><td class="num">' + esc(isConsumer(inv) ? inv.nid || '' : inv.reg || '') + '</td>' +
        '<td>' + lineSel(inv) + '</td>' + (side === 'P' ? '<td>' + useSel(inv) + '</td>' : '') +
        '<td class="n">' + money(t.net) + '</td><td class="n">' + money(t.vat) + '</td><td class="n">' + money(t.total) + '</td>' + (side === 'P' ? '<td class="n">' + (inv.wht ? money(inv.wht) : '') + '</td>' : '') +
        '<td class="wrap">' + notes + '</td><td style="white-space:nowrap">' + btns + '</td></tr>';
    }
    h += '</tbody>' + (list.length ? '<tfoot><tr><td colspan="' + ((side === 'P' ? 7 : 6) + (W ? 1 : 0)) + '">الإجمالي (' + int(list.length) + ' مستند، دون المستبعد، والإشعار الدائن بالسالب)</td><td class="n">' + money(tn) + '</td><td class="n">' + money(tv) + '</td><td class="n">' + money(tn + tv) + '</td>' + (side === 'P' ? '<td class="n">' + money(tw) + '</td>' : '') + '<td colspan="2"></td></tr></tfoot>' : '') + '</table></div>' +
      (list.length > shown.length ? '<div class="note info">يظهر أول ' + shown.length + ' مستند. استخدم البحث للوصول إلى البقية.</div>' : '') + '</div>';
    return h;
  }

  function filesTab(r, rows, iss) {
    const crit = rows.filter(i => iss.get(i.id).some(x => x[0] === 'crit'));
    const sh = (side) => V.uploadSheet(rows, side, ui.vMode, '');
    const sS = sh('S'), sP = sh('P');
    let h = '<div class="card"><div class="body"><div class="toolbar">' +
      '<label class="inl">طريقة الصفوف <select data-ui="vMode">' + options([['group', 'صف لكل فاتورة (حسب نوع السلعة والنسبة)'], ['items', 'صف لكل صنف كما في الفاتورة']], ui.vMode) + '</select></label></div>' +
      '<p class="hint" style="margin:10px 0 0">تُعَدّ الملفات وفق نماذج المصلحة تمامًا: اسم الورقة نفسه، وصف العناوين، وصف الأكواد مخفيًا، والبيانات بدءًا من الصف الثالث. «فئة الضريبة» 14 أو 5، وتُترك فارغة للصفر والإعفاء. الإشعار الدائن يُكتب «إشعار خصم» بنوع بيان «تسويات» وقيم موجبة. المصروفات والمستبعد لا تدخل الملف.</p></div></div>';
    if (crit.length) h += '<div class="note" style="background:var(--crit-soft)"><b style="color:var(--crit)">' + int(crit.length) + ' فاتورة عليها ملاحظات لازمة</b> (رقم تسجيل غير صحيح، رقم قومي ناقص…). سترفض المنظومة هذه الصفوف إذا رُفعت على حالها. صحّحها من تبويب المبيعات أو المشتريات.</div>';
    const card = (side, sh) => {
      const tot = sh.rows.reduce((a, x) => [a[0] + Number(x[21] || 0) * (x[0] === 3 ? -1 : 1), a[1] + Number(x[22] || 0) * (x[0] === 3 ? -1 : 1)], [0, 0]);
      let c = '<section class="card"><header><h2>' + (side === 'S' ? 'شيت المبيعات' : 'شيت المشتريات') + '</h2><span class="hint">' + int(sh.rows.length) + ' صف · صافي ' + money(tot[0]) + ' · ضريبة ' + money(tot[1]) + '</span><span class="grow"></span>' +
        (S.downloads && sh.rows.length ? '<button type="button" class="btn primary" data-act="vat-dl" data-side="' + side + '" data-fmt="xlsx" data-id="' + r.id + '">⤓ Excel</button><button type="button" class="btn" data-act="vat-dl" data-side="' + side + '" data-fmt="csv" data-id="' + r.id + '">⤓ CSV</button>' : '') + '</header>';
      if (!sh.rows.length) return c + '<div class="empty">لا توجد ' + (side === 'S' ? 'مبيعات' : 'مشتريات') + ' للرفع.</div></section>';
      c += '<div class="tablewrap"><table class="vat-tbl"><thead><tr>' + sh.head.map(x => '<th>' + esc(x) + '</th>').join('') + '</tr></thead><tbody>';
      for (const x of sh.rows.slice(0, 25)) c += '<tr>' + x.map((v, i) => '<td' + (typeof v === 'number' && i > 15 ? ' class="n"' : '') + '>' + esc(v) + '</td>').join('') + '</tr>';
      return c + '</tbody></table></div>' + (sh.rows.length > 25 ? '<div class="hint" style="padding:8px 16px">معاينة أول 25 صفًا من ' + int(sh.rows.length) + '.</div>' : '') + '</section>';
    };
    return h + card('S', sS) + card('P', sP);
  }

  function matchTab(r, rows, calc) {
    let h = '<div class="note info">بعد تقديم الإقرار على المنظومة، نزّل الإقرار بصيغة PDF وارفعه هنا. يقرأ البرنامج منه الملخص والأرقام وعدد الفواتير ويطابقها مع الفواتير المسجلة. إذا تعذّرت قراءة الملف فأدخل أرقام الملخص يدويًا.</div>';
    h += '<div class="toolbar">' + (A.writable() && !r.locked ? '<button type="button" class="btn primary" data-act="vat-upload-pdf">⤒ رفع الإقرار المقدَّم (PDF)</button><input type="file" id="vat-pdf" accept=".pdf" hidden><button type="button" class="btn" data-act="vat-manual-sub" data-id="' + r.id + '">إدخال الملخص يدويًا</button>' : '') +
      (r.sub && A.writable() && !r.locked ? '<span class="grow"></span><button type="button" class="btn danger" data-act="vat-clear-sub" data-id="' + r.id + '">حذف بيانات التقديم</button>' : '') + '</div>';
    if (!r.sub) return h + '<div class="card"><div class="empty">لم يُسجَّل تقديم الإقرار بعد.</div></div>';
    const s = r.sub, rec = V.reconcile(calc, s.parsed || {}, r);
    if (A.writable()) h += '<div class="toolbar">' + (!rec.ok && !s.accepted ? '<button type="button" class="btn" data-act="vat-accept" data-id="' + r.id + '" data-v="1">اعتماد الفرق كمقبول</button>' : '') +
      (s.accepted ? '<span class="pill ok">فرق مقبول' + (s.accepted.note ? ': ' + esc(s.accepted.note) : '') + '</span><button type="button" class="btn sm" data-act="vat-accept" data-id="' + r.id + '" data-v="0">إلغاء الاعتماد</button>' : '') +
      '<span class="grow"></span><button type="button" class="btn" data-act="vat-lock" data-id="' + r.id + '" data-v="' + (r.locked ? '0' : '1') + '">' + (r.locked ? '🔓 فتح القفل' : '🔒 قفل الإقرار') + '</button></div>';
    h += '<section class="card"><header><h2>' + (rec.ok ? '<span class="match-ok">✓ الإقرار المقدَّم مطابق</span>' : '<span class="match-bad">✗ يوجد اختلاف</span>') + '</h2><span class="grow"></span><span class="hint">' + (s.src === 'pdf' || s.viaPdf ? 'من ملف ' + esc(s.file || 'PDF') + (s.fileChunks ? ' · <a href="#" data-act="vat-view-sub">عرض الملف</a>' : '') : 'إدخال يدوي') + '</span></header><div class="body"><div class="kv">' +
      '<span>رقم الإقرار</span><span class="num">' + esc(s.formNo || '—') + '</span><span>تاريخ التقديم</span><span class="num">' + esc(s.sentDate || '—') + '</span>' +
      '<span>نوع الإقرار</span><span>' + (s.kind === 'amended' ? 'معدل' : s.kind === 'original' ? 'أصلي' : '—') + '</span>' +
      '<span>الفترة في الإقرار</span><span class="num">' + esc(s.from ? s.from + ' ← ' + s.to : '—') + (rec.periodOk ? '' : ' <span class="iss crit">لا تطابق فترة هذا الإقرار (' + esc(r.from) + ' ← ' + esc(r.to) + ')</span>') + '</span></div></div>';
    h += '<div class="tablewrap"><table><thead><tr><th>البند</th><th class="n">حسب النظام</th><th class="n">في الإقرار المقدَّم</th><th class="n">الفرق</th><th>النتيجة</th></tr></thead><tbody>';
    for (const x of rec.rows) {
      const f = x.isCount ? int : money;
      h += '<tr class="' + (x.ok ? '' : 'row-crit') + '"><td>' + esc(x.label) + '</td><td class="n">' + f(x.mine) + '</td><td class="n">' + (x.theirs == null ? 'غير موجود' : f(x.theirs)) + '</td><td class="n">' + (x.theirs == null ? '' : f(V.r2(x.mine - x.theirs))) + '</td><td>' + (x.ok ? '<span class="match-ok">✓</span>' : '<span class="match-bad">✗</span>') + '</td></tr>';
    }
    return h + '</tbody></table></div></section>';
  }

  /* ---------------- binding (inputs that are not plain filters) ---------------- */
  function bindRet(id, tab) {
    const r = retById(id), e = VS.cache.get(id); if (!r || !e || !e.loaded || r.locked) return;
    const f = $('#vat-file'); if (f) f.addEventListener('change', () => { if (f.files[0]) importFile(r, f.files[0]); f.value = ''; });
    const pf = $('#vat-pdf'); if (pf) pf.addEventListener('change', () => { if (pf.files[0]) readSubmittedPdf(r, pf.files[0]); pf.value = ''; });
    const prev = $('#vat-prev');
    if (prev) prev.addEventListener('change', async () => {
      if (await write(() => S.db.doc('vatReturns/' + id).update({ prevMode: 'manual', prevCredit: C.round2(prev.value) }))) { toast('تم تثبيت الرصيد الدائن السابق'); saveSoon(id); }
    });
    $$('[data-vat-set]').forEach(sel => sel.addEventListener('change', () => setInvoiceKind(r, sel.dataset.inv, sel.dataset.vatSet, sel.value)));
    $$('[data-vat-sel]').forEach(cb => cb.addEventListener('change', () => { if (cb.checked) VS.sel.add(cb.dataset.vatSel); else VS.sel.delete(cb.dataset.vatSel); A.renderPage(); }));
    $$('[data-vat-selall]').forEach(cb => cb.addEventListener('change', () => {
      const ids = $$('[data-vat-sel]').map(x => x.dataset.vatSel);
      for (const id of ids) { if (cb.checked) VS.sel.add(id); else VS.sel.delete(id); }
      A.renderPage();
    }));
    $$('[data-vat-bulk]').forEach(sel => sel.addEventListener('change', () => { if (sel.value !== '') bulk(r, sel.dataset.vatBulk, sel.value); }));
  }
  function setInvoiceKind(r, invId, key, value) {
    if (value === '') return;
    const e = VS.cache.get(r.id), inv = e.rows.find(i => i.id === invId); if (!inv) return;
    const v = key === 'o' ? Number(value) : key === 'e' ? value === '1' : value;
    const party = inv.reg || inv.name;
    const targets = ui.vSame ? e.rows.filter(i => i.side === inv.side && (i.reg || i.name) === party) : [inv];
    snap(r.id, key === 'e' ? 'تغيير مشتريات / مصروفات' : 'تغيير التصنيف');
    for (const t of targets) for (const l of t.lines) l[key] = v;
    saveSoon(r.id);
    saveParties(r.client, { [inv.side + ':' + party]: { [key]: v, name: inv.name } });
    if (targets.length > 1) toast('طُبِّق على ' + targets.length + ' فاتورة للـ' + (inv.side === 'P' ? 'مورد' : 'عميل') + ' نفسه');
    A.renderPage();
  }

  /* ---------------- bulk edit of the selected invoices ---------------- */
  function bulk(r, op, value) {
    const e = VS.cache.get(r.id); if (!e) return;
    const pick = e.rows.filter(i => VS.sel.has(i.id)); if (!pick.length) return;
    const done = (msg) => { saveSoon(r.id); A.renderPage(); toast(msg + ' (' + pick.length + ' مستند)'); };
    if (op === 'clear') { VS.sel.clear(); return A.renderPage(); }
    if (op !== 'del') snap(r.id, 'تعديل مجمع على ' + pick.length + ' مستند');
    if (op === 'x1' || op === 'x0') { for (const i of pick) { i.x = op === 'x1'; if (i.x) i.xr = 'استبعاد يدوي'; else delete i.xr; } return done(op === 'x1' ? 'تم الاستبعاد' : 'أُعيدت إلى الإقرار'); }
    if (op === 'e0' || op === 'e1') {
      const mem = {};
      for (const i of pick) { for (const l of i.lines) l.e = op === 'e1'; mem[i.side + ':' + (i.reg || i.name)] = { e: op === 'e1', name: i.name }; }
      saveParties(r.client, mem);
      return done(op === 'e1' ? 'صُنّفت مصروفات' : 'صُنّفت مشتريات');
    }
    if (op === 'o' || op === 'n' || op === 'rate') {
      const v = op === 'n' ? value : Number(value);
      for (const i of pick) for (const l of i.lines) {
        l[op] = v;
        if (op === 'rate') l.vat = V.r5((Number(l.net) || 0) * v / 100);
      }
      if (op !== 'rate') { const mem = {}; for (const i of pick) mem[i.side + ':' + (i.reg || i.name)] = { [op]: v, name: i.name }; saveParties(r.client, mem); }
      return done(op === 'rate' ? 'عُدّلت النسبة وأُعيد حساب الضريبة' : 'عُدّل التصنيف');
    }
    if (op === 'del') return confirmBox('حذف المستندات المحددة', 'سيُحذف ' + pick.length + ' مستند من الإقرار.', 'حذف', () => {
      snap(r.id, 'حذف ' + pick.length + ' مستند'); const ids = new Set(pick.map(i => i.id)); e.rows = e.rows.filter(i => !ids.has(i.id)); VS.sel.clear(); closeModal(); saveSoon(r.id); A.renderPage(); toast('تم الحذف');
    });
  }

  /* ---------------- import ---------------- */
  function readBook(file) {
    return new Promise((res, rej) => {
      if (!window.XLSX) return rej(new Error('تتطلب قراءة ملفات Excel اتصالًا بالإنترنت عند فتح الصفحة. اتصل بالإنترنت ثم أعد تحميل الصفحة.'));
      const rd = new FileReader();
      rd.onload = () => {
        try {
          const wb = XLSX.read(new Uint8Array(rd.result), { type: 'array', cellDates: true });
          const book = {}; for (const n of wb.SheetNames) book[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' });
          res(book);
        } catch (e) { rej(e); }
      };
      rd.onerror = () => rej(rd.error); rd.readAsArrayBuffer(file);
    });
  }
  async function importFile(r, file) {
    let book; try { book = await readBook(file); } catch (e) { return toast((e && e.message) || 'تعذّر قراءة الملف.', true); }
    const res = V.readWorkbook(book, memFor(r.client));
    if (!res) return toast('صيغة الملف غير معروفة. الملفات المقبولة: ملف التعاملات من منظومة الفاتورة الإلكترونية (الفواتير المرسلة والمستلمة)، أو نموذج رفع المبيعات أو المشتريات الصادر عن المصلحة.', true);
    const e = VS.cache.get(r.id), have = new Set(e.rows.map(V.invKey));
    const claimed = new Map(); for (const o of rets()) if (o.id !== r.id && o.client === r.client) for (const k of o.keys || []) claimed.set(k, o);
    // nothing is excluded automatically: invoices outside the period or already in another return are added and flagged
    const g = { inS: [], inP: [], out: [], dup: [], prior: [] };
    for (const inv of res.invoices) {
      inv.id = uid(); fixNo(inv);
      const k = V.invKey(inv);
      if (have.has(k)) { g.dup.push(inv); continue; }
      have.add(k);
      if (inv.side === 'P' && claimed.has(k)) { inv.prior = 'إقرار ' + periodLabel(claimed.get(k)); g.prior.push(inv); }
      if (inv.date && (inv.date < r.from || inv.date > r.to)) g.out.push(inv);
      (inv.side === 'S' ? g.inS : g.inP).push(inv);
    }
    const sum = (l) => l.reduce((a, i) => { const t = V.invTotals(i); return [a[0] + t.net, a[1] + t.vat]; }, [0, 0]);
    const line = (label, l) => { const t = sum(l); return '<tr><td>' + label + '</td><td class="n">' + int(l.length) + '</td><td class="n">' + money(t[0]) + '</td><td class="n">' + money(t[1]) + '</td></tr>'; };
    const body = '<p style="margin:0 0 8px">' + (res.kind === 'eta' ? 'ملف التعاملات من منظومة الفاتورة الإلكترونية.' : 'نموذج رفع ' + (res.invoices[0] && res.invoices[0].side === 'P' ? 'المشتريات' : 'المبيعات') + ' الصادر عن المصلحة.') + ' ' + esc(res.notes.join(' ')) + '</p>' +
      '<table><thead><tr><th></th><th class="n">العدد</th><th class="n">الصافي</th><th class="n">الضريبة</th></tr></thead><tbody>' +
      line('مبيعات', g.inS) + line('مشتريات', g.inP) + line('مسجلة بالفعل في الإقرار (تُتجاهل)', g.dup) + '</tbody></table>' +
      (g.out.length ? '<div class="note" style="margin-top:10px">' + int(g.out.length) + ' مستند تاريخه خارج فترة الإقرار (' + esc(r.from) + ' ← ' + esc(r.to) + '): ' + g.out.slice(0, 15).map(i => esc(i.no) + ' (' + esc(i.date) + ')').join('، ') + (g.out.length > 15 ? ' …' : '') + '. ستُضاف مع ملاحظة عليها، ويمكنك استبعاد ما تريد يدويًا.</div>' : '') +
      (g.prior.length ? '<div class="note" style="margin-top:10px">' + int(g.prior.length) + ' فاتورة شراء مدرجة أيضًا في إقرار آخر للشركة، وستُضاف مع ملاحظة عليها.</div>' : '') +
      (res.kind === 'eta' && !g.inS.length ? '<div class="note info" style="margin-top:10px">لا توجد مبيعات في الملف. إذا كانت الشركة تبيع للمستهلك النهائي دون فاتورة إلكترونية، فأضف هذه المبيعات من تبويب المبيعات «+ مبيعات مستهلك نهائي».</div>' : '');
    const bg = modal('استيراد الفواتير', body, '<button type="button" class="btn" data-close>إلغاء</button><button type="button" class="btn primary" id="vi-go">إضافة للإقرار</button>', true);
    $('#vi-go', bg).addEventListener('click', () => {
      const add = g.inS.concat(g.inP);
      if (!add.length) { closeModal(); return toast('لا توجد فواتير جديدة.'); }
      snap(r.id, 'استيراد ' + add.length + ' مستند'); e.rows.push(...add); closeModal(); saveSoon(r.id); A.renderPage();
      toast('أُضيف ' + add.length + ' مستند');
    });
  }

  /* ---------------- invoice editor ---------------- */
  function openInvoice(r, inv, preset) {
    const e = VS.cache.get(r.id), isNew = !inv, ro = !A.writable() || !!r.locked;
    const side = inv ? inv.side : preset.side;
    inv = inv ? clean(inv) : { id: uid(), side, src: 'manual', dt: Number(preset.dt) || 1, no: '', date: r.to, name: Number(preset.dt) === 5 ? 'مستهلك نهائي' : '', reg: '', file: '', addr: '', nid: '', phone: '', lines: [] };
    if (!inv.lines.length) inv.lines.push(V.defaultLine(side, { d: side === 'S' ? 'مبيعات' : '', net: 0, vat: 0, rate: 14, o: side === 'S' ? 1 : 3 }));
    const lines = inv.lines;
    const body = '<form id="vf-form" class="form">' +
      fld('vf-dt', 'نوع المستند', '<select id="vf-dt">' + options(side === 'S' ? S_DT : P_DT, inv.dt) + '</select>') + fld('vf-no', 'رقم المستند', inp('vf-no', inv.no)) +
      fld('vf-date', 'التاريخ', inp('vf-date', inv.date, 'date')) + fld('vf-name', side === 'S' ? 'اسم العميل' : 'اسم المورد', inp('vf-name', inv.name)) +
      '<div class="field" data-reg>' + '<label for="vf-reg">رقم التسجيل الضريبي (9 أرقام)</label>' + inp('vf-reg', inv.reg, 'text', ' inputmode="numeric" maxlength="9"') + '</div>' +
      '<div class="field" data-reg><label for="vf-file">رقم الملف الضريبي</label>' + inp('vf-file', inv.file) + '</div>' +
      fld('vf-addr', 'العنوان (اختياري)', inp('vf-addr', inv.addr)) + '<div class="field" data-nid><label for="vf-nid" id="vf-nid-l">الرقم القومي / رقم جواز السفر</label>' + inp('vf-nid', inv.nid) + '</div>' +
      fld('vf-phone', 'رقم الموبايل', inp('vf-phone', inv.phone)) +
      '<div class="field full"><label>البنود</label><div class="tablewrap"><table class="lines-ed"><thead><tr><th>الوصف</th>' + (side === 'S' ? '<th>سلعة/خدمة</th>' : '') + '<th>نوع السلعة</th>' + (side === 'P' ? '<th>مصروف؟</th>' : '') + '<th>النسبة</th><th>الصافي</th><th>الضريبة</th><th></th></tr></thead><tbody id="vf-lines"></tbody></table></div>' +
      (ro ? '' : '<button type="button" class="btn sm" id="vf-add" style="align-self:flex-start;margin-top:6px">+ بند</button>') + '</div>' +
      '<div class="field full"><label class="inl"><input type="checkbox" id="vf-x"' + (inv.x ? ' checked' : '') + '> استبعاد من الإقرار</label></div>' +
      (inv.wht ? '<div class="field full"><span class="hint">خصم وإضافة على الفاتورة: ' + money(inv.wht) + ' (يُستخدم في نموذج 41)</span></div>' : '') +
      '<div class="field full err" id="vf-err"></div></form>';
    const foot = (isNew || ro ? '' : '<button type="button" class="btn danger" id="vf-del">حذف</button><span style="flex:1"></span>') + '<button type="button" class="btn" data-close>' + (ro ? 'إغلاق' : 'إلغاء') + '</button>' + (ro ? '' : '<button type="submit" form="vf-form" class="btn primary">حفظ</button>');
    const bg = modal((isNew ? 'إضافة ' : '') + dtLabel(side, inv.dt) + (inv.no ? ' ' + inv.no : ''), body, foot, true);
    const sync = () => {
      const dt = Number(val('vf-dt', bg)), cons = side === 'S' && (dt === 5 || dt === 7 || dt === 8);
      $$('[data-reg]', bg).forEach(x => { x.style.display = cons ? 'none' : ''; });
      $('[data-nid]', bg).style.display = side === 'S' && dt !== 1 && dt !== 6 ? '' : 'none';
      $('#vf-nid-l', bg).textContent = dt === 7 ? 'رقم جواز السفر (إلزامي من 5,000)' : 'الرقم القومي (إلزامي إذا تجاوز المستند 150,000)';
    };
    function draw() {
      // leaving a focused field fires its «change» (which redraws) before the rows are replaced
      const a = document.activeElement; if (a && bg.contains(a) && a.closest('#vf-lines')) a.blur();
      $('#vf-lines', bg).innerHTML = lines.map((l, i) => '<tr data-i="' + i + '"><td class="d"><input data-k="d" value="' + esc(l.d) + '"></td>' +
        (side === 'S' ? '<td><select data-k="n">' + options([['G', 'سلعة'], ['S', 'خدمة']], l.n) + '</select></td>' : '') +
        '<td><select data-k="o">' + options(side === 'S' ? S_O : P_O, l.o) + '</select></td>' + (side === 'P' ? '<td style="text-align:center"><input type="checkbox" data-k="e"' + (l.e ? ' checked' : '') + '></td>' : '') +
        '<td><select data-k="rate">' + options(RATES.concat(RATES.some(x => x[0] === Number(l.rate)) ? [] : [[l.rate, l.rate + '%']]), Number(l.rate)) + '</select></td>' +
        '<td><input type="number" step="0.00001" data-k="net" value="' + esc(l.net) + '"></td><td><input type="number" step="0.00001" data-k="vat" value="' + esc(l.vat) + '"></td>' +
        '<td>' + (ro || lines.length < 2 ? '' : '<button type="button" class="btn sm danger" data-rm="' + i + '">✕</button>') + '</td></tr>').join('');
      $$('#vf-lines [data-k]', bg).forEach(el => el.addEventListener('change', () => {
        const l = lines[Number(el.closest('[data-i]').dataset.i)], k = el.dataset.k;
        if (k === 'e') l.e = el.checked; else if (k === 'o' || k === 'rate') l[k] = Number(el.value); else if (k === 'net' || k === 'vat') l[k] = Number(el.value) || 0; else l[k] = el.value;
        if (k === 'net' || k === 'rate') { l.vat = V.r5(l.net * (Number(l.rate) || 0) / 100); l.p = l.net; l.q = 1; }
        if (k === 'rate' && !Number(l.rate) && side === 'P') l.o = 7;
        draw();
      }));
      $$('#vf-lines [data-rm]', bg).forEach(b => b.addEventListener('click', () => { lines.splice(Number(b.dataset.rm), 1); draw(); }));
      if (ro) $$('input,select', bg).forEach(x => { x.disabled = true; });
    }
    draw(); sync();
    $('#vf-dt', bg).addEventListener('change', sync);
    const add = $('#vf-add', bg); if (add) add.addEventListener('click', () => { lines.push(V.defaultLine(side, { d: '', net: 0, vat: 0, rate: 14, o: side === 'S' ? 1 : 3 })); draw(); });
    $('#vf-form', bg).addEventListener('submit', (ev) => {
      ev.preventDefault();
      const dt = Number(val('vf-dt', bg)), cons = side === 'S' && (dt === 5 || dt === 7 || dt === 8);
      const n = Object.assign(inv, { dt, no: V.normNo(val('vf-no', bg)), date: val('vf-date', bg), name: val('vf-name', bg), reg: cons ? '' : val('vf-reg', bg).replace(/\D/g, ''), file: cons ? '' : val('vf-file', bg),
        addr: val('vf-addr', bg), nid: val('vf-nid', bg).replace(/\s/g, ''), phone: val('vf-phone', bg), x: $('#vf-x', bg).checked, lines: lines.filter(l => Number(l.net) || l.d) });
      if (!n.x) delete n.xr;
      const err = $('#vf-err', bg);
      if (!n.no) return (err.textContent = 'أدخل رقم المستند.');
      if (!n.date) return (err.textContent = 'أدخل التاريخ.');
      if (!n.lines.some(l => Number(l.net) > 0)) return (err.textContent = 'أدخل قيمة بند واحد على الأقل.');
      if (n.lines.some(l => Number(l.net) < 0 || Number(l.vat) < 0)) return (err.textContent = 'لا تُقبل القيم السالبة؛ يُسجَّل المرتجع «إشعار خصم» بقيم موجبة.');
      const i = e.rows.findIndex(x => x.id === n.id);
      snap(r.id, i >= 0 ? 'تعديل المستند ' + (inv.no || '') : 'إضافة مستند');
      if (i >= 0) e.rows[i] = n; else e.rows.push(n);
      const key = n.reg || n.name;
      if (key && !cons) saveParties(r.client, { [side + ':' + key]: { name: n.name, addr: n.addr, file: n.file } });
      closeModal(); saveSoon(r.id); A.renderPage(); toast('تم الحفظ');
    });
    const del = $('#vf-del', bg);
    if (del) del.addEventListener('click', () => confirmBox('حذف المستند', 'هل تريد حذف هذا المستند من الإقرار؟', 'حذف', () => { snap(r.id, 'حذف المستند ' + (inv.no || '')); e.rows = e.rows.filter(x => x.id !== inv.id); closeModal(); saveSoon(r.id); A.renderPage(); }));
  }

  /* ---------------- several sales documents typed (or pasted from Excel) at once ---------------- */
  function openMultiSales(r) {
    const e = VS.cache.get(r.id); if (!e) return;
    const blank = () => ({ dt: 5, no: '', date: r.to, name: 'مستهلك نهائي', reg: '', nid: '', n: 'G', rate: 14, net: '', vat: '', vatSet: false });
    const rows = Array.from({ length: 8 }, blank);
    const before = e.rows.filter(i => i.side === 'S' && i.dt === 5 && !i.x).reduce((a, i) => a + V.invTotals(i).net, 0);
    const body = '<div class="toolbar" style="margin-bottom:8px"><label class="inl">إجمالي مبيعات المستهلك النهائي من الدفاتر (صافي) <input type="number" step="0.01" min="0" id="ms-target" value="' + esc(r.consumerTarget || '') + '" style="width:170px"></label>' +
        (before ? '<span class="hint">المسجل سابقًا في الإقرار: ' + money(before) + '</span>' : '') + '<span class="grow"></span><b id="ms-left"></b></div>' +
      '<p class="hint" style="margin:0 0 8px">مبيعات للمستهلك النهائي المحلي (نوع المستند 5) دون رقم تسجيل. أدخل كل فاتورة أو إيصال بيع في سطر مستقل. تُحسب الضريبة من النسبة ويمكن تعديلها. ويمكنك النسخ من Excel واللصق في أي خانة، والأعمدة بالترتيب: رقم المستند، التاريخ، الصافي، الرقم القومي، الاسم.</p>' +
      '<div class="tablewrap" style="max-height:52vh"><table class="lines-ed"><thead><tr><th>رقم المستند</th><th>التاريخ</th><th>الاسم</th><th>الرقم القومي</th><th>سلعة / خدمة</th><th>النسبة</th><th>الصافي</th><th>الضريبة</th><th></th></tr></thead><tbody id="ms-rows"></tbody></table></div>' +
      '<div class="toolbar" style="margin-top:8px"><button type="button" class="btn sm" id="ms-add">+ 5 أسطر</button><span class="grow"></span><span id="ms-tot" class="hint"></span></div><div class="err" id="ms-err"></div>';
    const bg = modal('مبيعات المستهلك النهائي – ' + periodLabel(r), body, '<button type="button" class="btn" data-close>إلغاء</button><button type="button" class="btn primary" id="ms-go">إضافة للإقرار</button>', true);
    const idOk = (x) => !(x.dt === 5 && Number(x.net) + (Number(x.vat) || 0) > V.ID_LIMIT && !x.nid);
    function tot() {
      const f = rows.filter(x => Number(x.net) > 0), bad = f.filter(x => !idOk(x)).length;
      const n = f.reduce((a, x) => a + (x.dt === 3 ? -1 : 1) * Number(x.net), 0), v = f.reduce((a, x) => a + (x.dt === 3 ? -1 : 1) * (Number(x.vat) || 0), 0);
      const tg = Number($('#ms-target', bg).value) || 0, left = V.r2(tg - before - f.reduce((a, x) => a + Number(x.net), 0));
      $('#ms-left', bg).innerHTML = tg ? (Math.abs(left) < 0.005 ? '<span class="match-ok">✓ مطابق للإجمالي</span>' : left > 0 ? 'المتبقي: ' + money(left) : '<span class="match-bad">زيادة على الإجمالي: ' + money(-left) + '</span>') : '';
      $('#ms-tot', bg).innerHTML = int(f.length) + ' مستند · صافي ' + money(n) + ' · ضريبة ' + money(v) + (bad ? ' · <b style="color:var(--crit)">' + bad + ' مستند يتجاوز 150,000 دون رقم قومي</b>' : '');
    }
    function draw() {
      const a = document.activeElement; if (a && bg.contains(a) && a.closest('#ms-rows')) a.blur();
      $('#ms-rows', bg).innerHTML = rows.map((x, i) => '<tr data-i="' + i + '"' + (Number(x.net) > 0 && !idOk(x) ? ' class="row-crit"' : '') + '>' +
        '<td><input data-k="no" value="' + esc(x.no) + '" style="min-width:90px"></td>' +
        '<td><input type="date" data-k="date" value="' + esc(x.date) + '"></td><td><input data-k="name" value="' + esc(x.name) + '" style="min-width:120px"></td>' +
        '<td><input data-k="id" value="' + esc(x.nid) + '" inputmode="numeric" placeholder="إلزامي إذا تجاوز 150,000" style="min-width:130px"></td>' +
        '<td><select data-k="n" style="min-width:75px">' + options([['G', 'سلعة'], ['S', 'خدمة']], x.n) + '</select></td><td><select data-k="rate" style="min-width:110px">' + options(RATES, x.rate) + '</select></td>' +
        '<td><input type="number" step="0.01" min="0" data-k="net" value="' + esc(x.net) + '" style="min-width:100px"></td><td><input type="number" step="0.01" min="0" data-k="vat" value="' + esc(x.vat) + '" style="min-width:90px"></td>' +
        '<td><button type="button" class="btn sm danger" data-rm="' + i + '">✕</button></td></tr>').join('');
      $$('#ms-rows [data-k]', bg).forEach(el => el.addEventListener('change', () => {
        const x = rows[Number(el.closest('[data-i]').dataset.i)], k = el.dataset.k, v = el.value.trim();
        if (k === 'dt' || k === 'rate') x[k] = Number(v);
        else if (k === 'id') { if (x.dt === 5) x.nid = v.replace(/\s/g, ''); else x.reg = v.replace(/\D/g, ''); }
        else if (k === 'vat') { x.vat = v; x.vatSet = v !== ''; }
        else x[k] = v;
        if (k === 'dt' && x.dt !== 5 && x.name === 'مستهلك نهائي') x.name = '';
        if ((k === 'net' || k === 'rate') && !x.vatSet) x.vat = x.net === '' ? '' : String(V.r2(Number(x.net) * x.rate / 100));
        draw(); tot();
      }));
      $$('#ms-rows [data-rm]', bg).forEach(b => b.addEventListener('click', () => { rows.splice(Number(b.dataset.rm), 1); if (!rows.length) rows.push(blank()); draw(); tot(); }));
    }
    // pasting several lines from Excel fills the rows from the cell where it is pasted
    $('#ms-rows', bg).addEventListener('paste', (ev) => {
      const t = (ev.clipboardData || window.clipboardData).getData('text');
      if (!/[\t\n]/.test(t.trim())) return;
      ev.preventDefault();
      const start = Number((ev.target.closest('[data-i]') || { dataset: { i: 0 } }).dataset.i);
      t.replace(/\r/g, '').split('\n').filter(l => l.trim()).forEach((l, j) => {
        const c = l.split('\t').map(v => v.trim()), i = start + j;
        while (rows.length <= i) rows.push(blank());
        const x = rows[i];
        if (c[0]) x.no = c[0];
        if (c[1]) x.date = V.isoDate(c[1]) || x.date;
        if (c[2]) { x.net = String(V.num(c[2])); x.vatSet = false; x.vat = String(V.r2(V.num(c[2]) * x.rate / 100)); }
        if (c[3]) x.nid = c[3].replace(/\s/g, '');
        if (c[4]) x.name = c[4];
      });
      draw(); tot();
    });
    $('#ms-target', bg).addEventListener('input', tot);
    $('#ms-target', bg).addEventListener('change', () => { const v = C.round2($('#ms-target', bg).value); if (v !== (r.consumerTarget || 0)) write(() => S.db.doc('vatReturns/' + r.id).update({ consumerTarget: v })); });
    $('#ms-add', bg).addEventListener('click', () => { for (let i = 0; i < 5; i++) rows.push(blank()); draw(); });
    $('#ms-go', bg).addEventListener('click', () => {
      const a = document.activeElement; if (a && bg.contains(a)) a.blur();
      const f = rows.filter(x => Number(x.net) > 0), err = $('#ms-err', bg);
      if (!f.length) return (err.textContent = 'أدخل قيمة مستند واحد على الأقل.');
      const noNo = f.filter(x => !x.no).length; if (noNo) return (err.textContent = noNo + ' مستند بلا رقم. أدخل رقم كل مستند.');
      const noDate = f.filter(x => !x.date).length; if (noDate) return (err.textContent = noDate + ' مستند بلا تاريخ.');
      const have = new Set(e.rows.map(V.invKey)), add = []; let dup = 0;
      for (const x of f) {
        const inv = { id: uid(), side: 'S', src: 'manual', dt: x.dt, no: V.normNo(x.no), date: x.date, name: x.name || (x.dt === 5 ? 'مستهلك نهائي' : ''), reg: x.dt === 5 ? '' : x.reg, nid: x.dt === 5 ? x.nid : '', file: '', addr: '', phone: '',
          lines: [V.defaultLine('S', { d: x.n === 'S' ? 'خدمات' : 'مبيعات', q: 1, p: Number(x.net), net: Number(x.net), vat: Number(x.vat) || 0, rate: x.rate, n: x.n, o: x.rate ? 1 : 7 })] };
        if (have.has(V.invKey(inv))) { dup++; continue; }
        have.add(V.invKey(inv)); add.push(inv);
      }
      if (!add.length) return (err.textContent = 'جميع هذه المستندات مسجلة بالفعل في الإقرار.');
      snap(r.id, 'إدخال ' + add.length + ' مستند بيع'); e.rows.push(...add);
      closeModal(); saveSoon(r.id); A.renderPage();
      toast('أُضيف ' + add.length + ' مستند' + (dup ? ' · تم تجاهل ' + dup + ' مستندًا مكررًا' : ''));
    });
    draw(); tot();
  }

  /* ---------------- upload files ---------------- */
  function downloadSheet(r, side, fmt) {
    const e = VS.cache.get(r.id); if (!e) return;
    const sh = V.uploadSheet(e.rows, side, ui.vMode, '');
    if (!sh.rows.length) return toast('لا توجد صفوف للرفع.', true);
    const base = A.fileName((side === 'S' ? 'مبيعات' : 'مشتريات') + ' - ' + A.clientName(r.client) + ' - ' + r.period);
    const aoa = [sh.head, sh.codes].concat(sh.rows);
    if (fmt === 'csv') {
      const q = (v) => { const s = v == null ? '' : String(v); return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
      A.saveFile(base + '.csv', new Blob(['﻿' + aoa.map(row => row.map(q).join(',')).join('\r\n')], { type: 'text/csv;charset=utf-8' }));
    } else {
      if (!window.XLSX) return toast('يتطلب تنزيل ملفات Excel اتصالًا بالإنترنت عند فتح الصفحة.', true);
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws['!rows'] = [{ hpt: 30 }, { hidden: true }];
      ws['!cols'] = sh.head.map((h, i) => ({ wch: [4, 11, 7, 3, 9, 8, 6, 11].includes(i) ? 22 : 14 }));
      // text columns stay text (registration numbers, invoice numbers and dates keep their exact form)
      const wb = XLSX.utils.book_new(); wb.Workbook = { Views: [{ RTL: true }] };
      XLSX.utils.book_append_sheet(wb, ws, sh.name);
      const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      A.saveFile(base + '.xlsx', new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    }
    if ((r.status || 'draft') === 'draft') write(() => S.db.doc('vatReturns/' + r.id).update({ status: 'ready' }));
  }

  /* ---------------- submitted return ---------------- */
  function loadScript(src) {
    return new Promise((res, rej) => { const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => rej(new Error('تعذّر تحميل قارئ ملفات PDF. تحقق من الاتصال بالإنترنت.')); document.head.appendChild(s); });
  }
  async function pdfLib() {
    if (window.pdfjsLib) return window.pdfjsLib;
    // the worker script runs on the page itself (no separate worker), so it works inside the artifact frame too
    await loadScript('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js');
    await loadScript('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js');
    window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
    return window.pdfjsLib;
  }
  async function readSubmittedPdf(r, file) {
    toast('جارٍ قراءة الإقرار…');
    let parsed, bytes;
    try {
      const lib = await pdfLib();
      bytes = new Uint8Array(await file.arrayBuffer());
      const doc = await lib.getDocument({ data: bytes.slice() }).promise;
      const pages = [];
      for (let p = 1; p <= doc.numPages; p++) { const tc = await (await doc.getPage(p)).getTextContent(); pages.push(tc.items.map(i => i.str)); }
      parsed = V.parseReturnPdf(pages);
    } catch (e) { return toast((e && e.message) || 'تعذّر قراءة الملف.', true); }
    if (!parsed.summary && !parsed.formNo) return toast('لم يُعثر على ملخص الضريبة في الملف. تأكد أنه إقرار القيمة المضافة (نموذج 10) الصادر من المنظومة، أو أدخل الملخص يدويًا.', true);
    const keep = Array.from(new Set(parsed.numbers.filter(x => x && Math.abs(x) >= 0.01).map(x => V.r2(x)))).slice(0, 400);
    const fileChunks = await saveFileBytes(r, bytes, file.name, r.sub && r.sub.fileChunks);
    await saveSub(r, { src: 'pdf', file: file.name, fileChunks, fileSize: bytes.length, formNo: parsed.formNo, kind: parsed.kind, sentDate: parsed.sentDate, from: parsed.from, to: parsed.to,
      parsed: { summary: parsed.summary, counts: parsed.counts, numbers: keep, from: parsed.from, to: parsed.to } });
  }
  /* the submitted return's PDF is kept in the database as base64 pieces: vatFiles/<returnId>-<k> */
  const B64 = 240000;
  function toB64(bytes) { let s = ''; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000)); return btoa(s); }
  function fromB64(b64) { const s = atob(b64), out = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i); return out; }
  async function saveFileBytes(r, bytes, name, oldChunks) {
    const b64 = toB64(bytes), n = Math.ceil(b64.length / B64);
    const ok = await write(async () => {
      for (let k = 0; k < n; k++) await S.db.doc('vatFiles/' + r.id + '-' + k).set({ ret: r.id, k, name, data: b64.slice(k * B64, (k + 1) * B64) });
      for (let k = n; k < (oldChunks || 0); k++) await S.db.doc('vatFiles/' + r.id + '-' + k).delete();
    });
    return ok ? n : 0;
  }
  async function deleteFile(r) { for (let k = 0; k < ((r.sub && r.sub.fileChunks) || 0); k++) await S.db.doc('vatFiles/' + r.id + '-' + k).delete(); }
  async function loadFile(r) {
    let b64 = '';
    for (let k = 0; k < r.sub.fileChunks; k++) { const snap = await S.db.doc('vatFiles/' + r.id + '-' + k).get(); if (!exists(snap)) throw new Error('ملف الإقرار غير مكتمل. أعد رفعه من تبويب المطابقة.'); b64 += snap.data().data; }
    return fromB64(b64);
  }
  async function downloadSubmitted(r) {
    if (!r || !r.sub || !r.sub.fileChunks) return toast('لم يُحفظ ملف لهذا الإقرار. ارفعه من تبويب «المطابقة بعد التقديم».', true);
    toast('جارٍ تجهيز الملف…');
    try {
      const bytes = await loadFile(r);
      await A.saveFile(r.sub.file || (A.fileName('إقرار ' + A.clientName(r.client) + ' ' + r.period) + '.pdf'), new Blob([bytes], { type: 'application/pdf' }));
    } catch (e) { toast((e && e.message) || 'تعذّر تنزيل الملف.', true); }
  }
  async function viewSubmitted(r) {
    if (!r.sub || !r.sub.fileChunks) return toast('لم يُحفظ ملف للإقرار المقدَّم. ارفعه من تبويب «المطابقة بعد التقديم».', true);
    const bg = modal('الإقرار المقدَّم' + (r.sub.formNo ? ' رقم ' + r.sub.formNo : '') + (r.sub.sentDate ? ' · تاريخ التقديم ' + r.sub.sentDate : ''),
      '<div id="sp-pages" class="sp-pages"><div class="loading">جارٍ تحميل الإقرار…</div></div>',
      '<button type="button" class="btn" data-close>إغلاق</button>' + (S.downloads ? '<button type="button" class="btn primary" id="sp-dl">⤓ تنزيل PDF</button>' : ''), true);
    bg.querySelector('.modal').classList.add('pdfv');
    let bytes;
    try { bytes = await loadFile(r); } catch (e) { $('#sp-pages', bg).innerHTML = '<div class="empty">' + esc(e.message) + '</div>'; return; }
    const name = r.sub.file || (A.fileName('إقرار ' + A.clientName(r.client) + ' ' + r.period) + '.pdf');
    const dl = $('#sp-dl', bg); if (dl) dl.addEventListener('click', () => A.saveFile(name, new Blob([bytes], { type: 'application/pdf' })));
    try {
      const lib = await pdfLib(), CDN = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/';
      // character maps and standard fonts so the Arabic text of the portal's PDF draws correctly
      const doc = await lib.getDocument({ data: bytes.slice(), cMapUrl: CDN + 'cmaps/', cMapPacked: true, standardFontDataUrl: CDN + 'standard_fonts/', useSystemFonts: true }).promise, box = $('#sp-pages', bg);
      box.innerHTML = '';
      const w = Math.min(900, box.clientWidth || 860);
      for (let p = 1; p <= doc.numPages; p++) {
        const page = await doc.getPage(p), v1 = page.getViewport({ scale: 1 }), scale = w / v1.width, vp = page.getViewport({ scale: scale * 2 });
        const cv = document.createElement('canvas'); cv.width = vp.width; cv.height = vp.height; cv.style.width = '100%';
        box.appendChild(cv);
        await page.render({ canvasContext: cv.getContext('2d'), viewport: vp }).promise;
      }
    } catch (e) { $('#sp-pages', bg).innerHTML = '<div class="empty">تعذّر عرض الملف هنا' + (S.downloads ? '، ويمكنك تنزيله.' : '.') + '</div>'; }
  }
  async function saveSub(r, sub) {
    const e = VS.cache.get(r.id);
    const rec = V.reconcile(V.compute(e.rows, prevCredit(r)), sub.parsed, r);
    sub.ok = rec.ok; sub.at = new Date().toISOString();
    if (await write(() => S.db.doc('vatReturns/' + r.id).update(clean({ sub, status: 'submitted' })))) {
      toast(rec.ok ? 'الإقرار المقدَّم مطابق' : 'تم الحفظ: يوجد اختلاف، راجع الجدول', !rec.ok);
      // the matching VAT task in «الإجراءات والمواعيد» is marked done with the return number
      const t = (S.raw.tasks || []).find(t => t.client === r.client && t.obligation === 'vat' && t.period === r.period && t.status !== 'done');
      if (t) write(() => S.db.doc('tasks/' + t.id).update({ status: 'done', doneDate: sub.sentDate || C.todayISO(), ref: sub.formNo || t.ref || '' }));
    }
  }
  function openManualSub(r) {
    const s = (r.sub && r.sub.parsed && r.sub.parsed.summary) || {};
    const f = (id, label, v) => fld(id, label, inp(id, v == null ? '' : v, 'number', ' step="0.01"'));
    const body = '<form id="ms-form" class="form">' + fld('ms-no', 'رقم الإقرار', inp('ms-no', r.sub ? r.sub.formNo : '')) + fld('ms-date', 'تاريخ التقديم', inp('ms-date', r.sub ? r.sub.sentDate : C.todayISO(), 'date')) +
      f('ms-out', 'ضريبة القيمة المضافة (1)', s.out) + f('ms-inp', 'ضريبة المدخلات (2)', s.inp) + f('ms-due', 'الضريبة المستحقة (1) − (2)', s.due) + f('ms-prev', 'الرصيد الدائن السابق', s.prev) +
      f('ms-res', 'الصافي (موجب = مدين مستحق السداد، سالب = دائن)', s.res) + '</form>';
    const bg = modal('بيانات الإقرار المقدَّم', body, '<button type="button" class="btn" data-close>إلغاء</button><button type="submit" form="ms-form" class="btn primary">حفظ ومطابقة</button>');
    $('#ms-form', bg).addEventListener('submit', async (ev) => {
      ev.preventDefault();
      const n = (id) => C.round2(val(id, bg));
      await saveSub(r, { src: 'manual', formNo: val('ms-no', bg), sentDate: val('ms-date', bg), kind: r.kind, from: r.from, to: r.to,
        parsed: { summary: { out: n('ms-out'), inp: n('ms-inp'), due: n('ms-due'), prev: n('ms-prev'), res: n('ms-res') }, counts: null, numbers: [] } });
      closeModal();
    });
  }

  /* ---------------- summary PDF ---------------- */
  function pdfSummary(r) {
    const e = VS.cache.get(r.id); if (!e) return;
    const c = client(r.client) || {}, calc = V.compute(e.rows, prevCredit(r));
    const box = '<div class="box"><span><b>الشركة:</b> ' + esc(c.name || '') + '</span><span><b>رقم التسجيل:</b> ' + esc(c.taxRegNo || '—') + '</span><span><b>المأمورية:</b> ' + esc(c.taxOffice || '—') + '</span><span><b>الفترة:</b> ' + esc(r.from) + ' ← ' + esc(r.to) + ' (' + esc(FREQ[r.freq]) + ')</span>' +
      (r.sub ? '<span><b>رقم الإقرار:</b> ' + esc(r.sub.formNo || '—') + '</span><span><b>تاريخ التقديم:</b> ' + esc(r.sub.sentDate || '—') + '</span>' : '') + '</div>';
    const sumBox = '<div class="summary"><div><small>ضريبة المبيعات (1)</small><b>' + money(calc.out) + '</b></div><div><small>ضريبة المدخلات (2)</small><b>' + money(calc.inp) + '</b></div><div><small>المستحقة (1)−(2)</small><b>' + money(calc.due) + '</b></div><div><small>الرصيد الدائن السابق</small><b>' + money(calc.prev) + '</b></div><div><small>' + (calc.res > 0 ? 'مستحق السداد' : 'رصيد دائن مرحّل') + '</small><b>' + money(Math.abs(calc.res)) + '</b></div></div>';
    const strip = (html) => html.replace(/<section class="card"><header><h2>/g, '<h2>').replace(/<\/h2><span class="hint">[^<]*<\/span><\/header><div class="tablewrap">/g, '</h2>').replace(/<\/table><\/div><\/section>/g, '</table>');
    const inner = A.pdfHead('إقرار ضريبة القيمة المضافة (نموذج 10)', periodLabel(r) + (r.kind === 'amended' ? ' · معدل' : '')) + box + sumBox + '<h2>المبيعات والمشتريات حسب النوع</h2>' + brTables(V.breakdown(e.rows)) + strip(form10(calc)) +
      '<p class="foot-note">عدد فواتير المبيعات ' + int(calc.counts.sInv + calc.counts.sOther) + ' · فواتير الشراء ' + int(calc.counts.pInv) + ' · أذون الإفراج ' + int(calc.counts.pRelease) + '</p>';
    A.makePdf(A.fileName('إقرار قيمة مضافة ' + (c.name || '') + ' ' + r.period) + '.pdf', inner, true);
  }

  /* ---------------- reports ---------------- */
  const REPS = [['months', 'ملخص الإقرارات بالشهور'], ['summary', 'جميع الإقرارات'], ['due', 'الضريبة المستحقة السداد'], ['credit', 'الأرصدة الدائنة والمستحق'], ['books', 'دفتر المبيعات والمشتريات']];
  /* the return used for a client and period: the latest amended one when it exists */
  function retFor(cid, period) {
    return rets().filter(r => r.client === cid && r.period === period).sort((a, b) => (a.kind === 'amended' ? -1 : 1) - (b.kind === 'amended' ? -1 : 1) || String(b.updatedAt || '').localeCompare(String(a.updatedAt || '')))[0] || null;
  }
  const firstPer = (cid) => rets().filter(r => r.client === cid).map(r => r.period).sort()[0] || '9999';
  const SHORT = { draft: 'مسودة', ready: 'جاهز', submitted: 'مقدَّم', matched: '✓ مطابق', accepted: '✓ فرق مقبول', diff: '✗ فروق' };
  /* rows of returns with sales and purchases split by kind. list: [{ lead: [label cells], r: return or null, note }] */
  const BRL = Object.fromEntries(V.BR_S.concat(V.BR_P));
  function brReport(leadCols, list) {
    const col = A.col; let pending = 0;
    const recs = list.map(x => {
      if (!x.r) return x;
      const s = x.r.sum || {}, prev = prevCredit(x.r), res = V.r2((s.due || 0) - prev), b = brOf(x.r);
      if (b === undefined) pending++;
      return Object.assign({ s, prev, res, b: b || null }, x);
    });
    const any0 = recs.some(x => x.b && x.b.s0), anyT = recs.some(x => x.b && x.b.sTbl);
    const sK = V.BR_S.filter(([k]) => k !== 's0' || any0);
    const cols = leadCols.concat(sK.map(([, l]) => col(l, 1)), [col('ض. المبيعات', 1)], anyT ? [col('ض. الجدول', 1)] : [], V.BR_P.map(([, l]) => col(l, 1)),
      [col('ض. المدخلات', 1), col('المستحقة', 1), col('الرصيد السابق', 1), col('مستحق السداد', 1), col('رصيد دائن مرحّل', 1), col('الحالة'), col('التقديم')]);
    const groups = [['', leadCols.length], ['المبيعات', sK.length + 1 + (anyT ? 1 : 0)], ['المشتريات', V.BR_P.length + 1], ['نتيجة الإقرار', 4], ['', 2]];
    const nNum = cols.length - leadCols.length - 2, tot = new Array(nNum).fill(0);
    const rows = recs.map(x => {
      if (!x.r) return x.lead.concat(new Array(nNum).fill(''), [x.note || '', '']);
      const b = x.b, s = x.s, v = (k) => b ? b[k] || 0 : '';
      const nums = sK.map(([k]) => v(k)).concat([s.out || 0], anyT ? [v('sTbl')] : [], V.BR_P.map(([k]) => v(k)), [s.inp || 0, s.due || 0, x.prev, x.res > 0 ? x.res : '', x.res < 0 ? -x.res : '']);
      nums.forEach((n, i) => { tot[i] += Number(n) || 0; });
      return x.lead.concat(nums, [SHORT[statusOf(x.r)], x.r.sub ? [x.r.sub.formNo, x.r.sub.sentDate].filter(Boolean).join(' · ') : '']);
    });
    // previous credit and the credit carried forward do not add up across periods
    tot[nNum - 4] = ''; tot[nNum - 1] = '';
    const foot = ['الإجمالي'].concat(new Array(leadCols.length - 1).fill(''), tot.map(n => n === '' ? '' : V.r2(n)), ['', '']);
    return { cols, groups, rows, foot, pending };
  }
  function repData(key) {
    const col = A.col;
    if (key === 'due') return dueRep();
    if (key === 'stmt') return stmtRep();
    if (key === 'months') {
      const y = ui.rmYear || String(new Date().getFullYear());
      if (!ui.rmClient) {
        const cl = S.M.clients.filter(c => c.status !== 'inactive' && (c.vatRegistered || (c.obligations || []).includes('vat') || rets().some(r => r.client === c.id)));
        const rows = cl.map(c => {
          const q = A.lawOf(c) === '6-2025';
          return [c.name, q ? 'ربع سنوي' : 'شهري'].concat(A.MONTHS.map((m, i) => {
            const per = y + '-' + String(i + 1).padStart(2, '0');
            if (q && (i + 1) % 3) return '';
            const r = retFor(c.id, per); if (!r) return per < S.M.ym && per > firstPer(c.id) ? '— لم يُعَدّ' : '';
            const res = V.r2((r.sum ? r.sum.due : 0) - prevCredit(r));
            return SHORT[statusOf(r)] + (res > 0 ? ' · مستحق ' + money(res) : res < 0 ? ' · دائن ' + money(-res) : '');
          }));
        });
        return { title: 'حالة إقرارات القيمة المضافة بالشهور', ctx: 'سنة ' + y + ' · ' + int(rows.length) + ' شركة', landscape: true, cols: [col('الشركة'), col('الدورية')].concat(A.MONTHS.map(m => col(m))), rows };
      }
      const c = client(ui.rmClient) || {}, q = A.lawOf(c) === '6-2025', list = [];
      A.MONTHS.forEach((m, i) => {
        if (q && (i + 1) % 3) return;
        const per = y + '-' + String(i + 1).padStart(2, '0'), r = retFor(ui.rmClient, per);
        list.push({ lead: [(q ? 'الربع المنتهي ' + m : m) + (r && r.kind === 'amended' ? ' (معدل)' : '')], r, note: !r && per < S.M.ym && per > firstPer(ui.rmClient) ? 'لم يُعَدّ' : '' });
      });
      const b = brReport([col(q ? 'الربع' : 'الشهر')], list);
      return Object.assign(b, { title: 'ملخص إقرارات القيمة المضافة بالشهور', ctx: (c.name || '') + ' · سنة ' + y + (q ? ' · ربع سنوي' : ' · شهري'), landscape: true });
    }
    if (key === 'credit') {
      const rows = [];
      for (const c of S.M.clients) {
        const list = rets().filter(r => r.client === c.id).sort((a, b) => b.period.localeCompare(a.period));
        if (!list.length) continue;
        const last = list[0], res = V.r2((last.sum ? last.sum.due : 0) - prevCredit(last));
        const lastSub = list.find(r => r.sub && r.sub.sentDate);
        rows.push([c.name, periodLabel(last), STATUS[statusOf(last)][0], res > 0 ? res : '', res < 0 ? -res : '', list.filter(r => r.sub).length, list.filter(r => statusOf(r) === 'diff').length, lastSub ? lastSub.sub.sentDate + ' (' + periodLabel(lastSub) + ')' : '']);
      }
      return { title: 'الأرصدة الدائنة والمستحق حسب آخر إقرار', ctx: int(rows.length) + ' شركة', cols: [col('الشركة'), col('آخر إقرار'), col('الحالة'), col('مستحق السداد', 1), col('رصيد دائن مرحّل', 1), col('إقرارات مقدمة'), col('بها فروق'), col('آخر تاريخ تقديم')], rows,
        foot: ['الإجمالي', '', '', rows.reduce((s, r) => s + (Number(r[3]) || 0), 0), rows.reduce((s, r) => s + (Number(r[4]) || 0), 0), '', '', ''] };
    }
    if (key === 'books') {
      const side = ui.rbSide === 'P' ? 'P' : 'S';
      const cols = [col('التاريخ'), col('الإقرار'), col('نوع المستند'), col('الرقم'), col(side === 'S' ? 'العميل' : 'المورد'), col('التسجيل / القومي'), col(side === 'S' ? 'التصنيف' : 'نوع السلعة')]
        .concat(side === 'P' ? [col('الاستخدام'), col('التصنيف')] : []).concat([col('الصافي', 1), col('الضريبة', 1), col('الإجمالي', 1)]).concat(side === 'P' ? [col('خصم وإضافة', 1)] : []);
      const out = { title: side === 'S' ? 'دفتر المبيعات' : 'دفتر المشتريات والمصروفات', ctx: (A.clientName(ui.rbClient) || '') + ' · ' + (ui.rbFrom || 'البداية') + ' ← ' + (ui.rbTo || 'اليوم'), cols, rows: [], landscape: true, pending: 0 };
      if (!ui.rbClient) { out.needClient = true; return out; }
      const list = rets().filter(r => r.client === ui.rbClient && (!ui.rbFrom || r.to >= ui.rbFrom) && (!ui.rbTo || r.from <= ui.rbTo)).sort((a, b) => a.period.localeCompare(b.period));
      const t = [0, 0, 0, 0], byKind = new Map();
      const addKind = (k, net, vat) => { const x = byKind.get(k) || [0, 0]; x[0] += net; x[1] += vat; byKind.set(k, x); };
      for (const r of list) {
        const e = VS.cache.get(r.id); if (!e || !e.loaded) { out.pending++; loadRet(r); continue; }
        for (const inv of e.rows.filter(i => i.side === side && !i.x && (!ui.rbFrom || i.date >= ui.rbFrom) && (!ui.rbTo || i.date <= ui.rbTo)).sort((a, b) => a.date.localeCompare(b.date))) {
          const sg = inv.dt === 3 ? -1 : 1, tt = V.invTotals(inv);
          const cat = (ls) => Array.from(new Set(ls.map(l => BRL[V.lineKind(inv, l)]))).join(' + ');
          for (const l of inv.lines) addKind(side === 'P' && l.e ? 'exp' : V.lineKind(inv, l), sg * (Number(l.net) || 0), sg * (Number(l.vat) || 0));
          const kinds = side === 'S' ? cat(inv.lines) : Array.from(new Set(inv.lines.map(l => (P_O.find(x => x[0] === l.o) || ['', ''])[1]))).join(' + ');
          const use = side === 'P' ? (inv.lines.every(l => l.e) ? 'مصروفات' : inv.lines.some(l => l.e) ? 'مختلط' : 'مشتريات') : null;
          const row = [inv.date, periodLabel(r), dtLabel(side, inv.dt), inv.no, inv.name, isConsumer(inv) ? inv.nid || '' : inv.reg || '', kinds].concat(side === 'P' ? [use, cat(inv.lines.filter(l => !l.e)) || '—'] : []).concat([sg * tt.net, sg * tt.vat, sg * tt.total]).concat(side === 'P' ? [Number(inv.wht) || ''] : []);
          out.rows.push(row); t[0] += sg * tt.net; t[1] += sg * tt.vat; t[2] += sg * tt.total; t[3] += Number(inv.wht) || 0;
        }
      }
      out.foot = ['الإجمالي', '', '', '', '', '', ''].concat(side === 'P' ? ['', ''] : []).concat([t[0], t[1], t[2]]).concat(side === 'P' ? [t[3]] : []);
      const kinds = (side === 'S' ? V.BR_S : V.BR_P.concat([['exp', 'مصروفات (خارج الإقرار)']])).filter(([k]) => byKind.has(k));
      out.extra = { title: 'الإجمالي حسب التصنيف', cols: [col('التصنيف'), col('الصافي', 1), col('الضريبة', 1), col('الإجمالي', 1)],
        rows: kinds.map(([k, l]) => { const x = byKind.get(k); return [l, V.r2(x[0]), V.r2(x[1]), V.r2(x[0] + x[1])]; }), foot: ['الإجمالي', V.r2(t[0]), V.r2(t[1]), V.r2(t[2])] };
      return out;
    }
    const list = rets().filter(r => (!ui.rsYear || r.period.startsWith(ui.rsYear)) && (!ui.rbClient || r.client === ui.rbClient)).sort((a, b) => a.period.localeCompare(b.period) || A.clientName(a.client).localeCompare(A.clientName(b.client), 'ar'));
    const b = brReport([col('الشركة'), col('الفترة')], list.map(r => ({ lead: [A.clientName(r.client), periodLabel(r) + (r.kind === 'amended' ? ' (معدل)' : '')], r })));
    return Object.assign(b, { title: 'ملخص إقرارات القيمة المضافة', ctx: (ui.rsYear || 'جميع السنوات') + (ui.rbClient ? ' · ' + A.clientName(ui.rbClient) : ''), landscape: true });
  }
  function pageReports(key) {
    if (!REPS.some(x => x[0] === key)) key = 'months';
    let h = topBar('تقارير الإقرارات', 'يمكن تنزيل كل تقرير بصيغة PDF أو Excel') + '<div class="content">';
    h += '<div class="tabs">' + REPS.map(([k, l]) => '<a href="#vat-rep-' + k + '" class="' + (key === k ? 'on' : '') + '">' + esc(l) + '</a>').join('') + '</div>';
    const years = Array.from(new Set(rets().map(r => r.period.slice(0, 4)))).sort().reverse();
    const rep = repData(key);
    let f = '';
    if (key === 'months') {
      const ys = Array.from(new Set(years.concat([String(new Date().getFullYear())]))).sort().reverse();
      f = '<select data-ui="rmClient">' + options(A.clientOptions(true), ui.rmClient, 'جميع الشركات (الحالة الشهرية)') + '</select><select data-ui="rmYear">' + options(ys, ui.rmYear) + '</select>';
    }
    if (key === 'due') f = '<select data-ui="duClient">' + options(A.clientOptions(true), ui.duClient, 'جميع الشركات') + '</select><select data-ui="duAge">' + options([['late', 'المتأخر فقط']].concat(AGE), ui.duAge, 'كل الأعمار') + '</select>';
    if (key === 'summary') f = '<select data-ui="rsYear">' + options(years, ui.rsYear, 'جميع السنوات') + '</select><select data-ui="rbClient">' + options(A.clientOptions(true), ui.rbClient, 'جميع الشركات') + '</select>';
    if (key === 'books') f = '<select data-ui="rbClient">' + options(A.clientOptions(true), ui.rbClient, 'اختر الشركة…') + '</select><select data-ui="rbSide">' + options([['S', 'دفتر المبيعات'], ['P', 'دفتر المشتريات والمصروفات']], ui.rbSide) + '</select>' +
      '<label class="inl">من <input type="date" data-ui="rbFrom" value="' + esc(ui.rbFrom) + '"></label><label class="inl">إلى <input type="date" data-ui="rbTo" value="' + esc(ui.rbTo) + '"></label>';
    const canExp = S.downloads && !rep.needClient && !(rep.pending && key === 'books');
    h += '<div class="toolbar">' + f + '<span class="grow"></span>' + (canExp ? '<button type="button" class="btn" data-act="vat-rep-pdf" data-key="' + key + '">⤓ PDF</button><button type="button" class="btn" data-act="vat-rep-xlsx" data-key="' + key + '">⤓ Excel</button>' : '') + '</div>';
    if (rep.needClient) return h + '<div class="card"><div class="empty">اختر الشركة لعرض الدفتر.</div></div></div>';
    if (rep.pending && key === 'books') return h + '<div class="loading">جارٍ تحميل فواتير ' + int(rep.pending) + ' إقرار…</div></div>';
    if (rep.pending) h += '<div class="note info">جارٍ تجهيز تقسيم المبيعات والمشتريات لـ ' + int(rep.pending) + ' إقرار محفوظ قبل إضافة التقسيم… (مرة واحدة فقط)</div>';
    return h + '<section class="card"><header><h2>' + esc(rep.title) + '</h2><span class="hint">' + esc(rep.ctx || '') + '</span></header><div class="tablewrap">' + A.tableHtml(rep) + '</div></section>' +
      (rep.extra ? '<section class="card"><header><h2>' + esc(rep.extra.title) + '</h2></header><div class="tablewrap">' + A.tableHtml(rep.extra) + '</div></section>' : '') + '</div>';
  }

  /* ---------------- suppliers / customers memory ---------------- */
  function pageParties() {
    let h = topBar('الموردون والعملاء', 'البيانات المحفوظة لكل مورد أو عميل: العنوان ورقم الملف والتصنيف والمصروفات، وتُطبَّق تلقائيًا عند الاستيراد التالي') + '<div class="content">';
    h += '<div class="toolbar"><select data-ui="pClient">' + options(A.clientOptions(true), ui.pClient, 'اختر الشركة…') + '</select><input type="search" placeholder="بحث…" data-ui="pQ" data-live value="' + esc(ui.pQ) + '"></div>';
    if (!ui.pClient) return h + '<div class="card"><div class="empty">اختر الشركة.</div></div></div>';
    const d = partyDoc(ui.pClient), q = ui.pQ.trim().toLowerCase();
    const list = Object.entries(d.map || {}).filter(([k, v]) => !q || (k + ' ' + (v.name || '') + ' ' + (v.addr || '')).toLowerCase().includes(q)).sort((a, b) => String(a[1].name || '').localeCompare(String(b[1].name || ''), 'ar'));
    h += '<div class="card"><div class="tablewrap"><table class="vat-tbl"><thead><tr><th>الجهة</th><th>الاسم</th><th>رقم التسجيل</th><th>العنوان</th><th>رقم الملف</th><th>التصنيف</th><th>الاستخدام</th></tr></thead><tbody>';
    if (!list.length) h += '<tr><td colspan="7" class="empty">لا يوجد شيء محفوظ لهذه الشركة بعد.</td></tr>';
    for (const [k, v] of list) {
      const [side, id] = [k.slice(0, 1), k.slice(2)];
      const dis = A.writable() ? '' : ' disabled';
      h += '<tr data-party="' + esc(k) + '"><td>' + (side === 'P' ? 'مورد' : 'عميل') + '</td><td class="wrap">' + esc(v.name || '') + '</td><td class="num">' + esc(/^\d{9}$/.test(id) ? id : '') + '</td>' +
        '<td><input data-pk="addr" value="' + esc(v.addr || '') + '"' + dis + '></td><td><input data-pk="file" value="' + esc(v.file || '') + '" style="width:110px"' + dis + '></td>' +
        '<td>' + (side === 'P' ? '<select data-pk="o"' + dis + '>' + options(P_O, v.o || '', 'تلقائي') + '</select>' : '<select data-pk="n"' + dis + '>' + options([['G', 'سلعة'], ['S', 'خدمة']], v.n || '', 'تلقائي') + '</select>') + '</td>' +
        '<td>' + (side === 'P' ? '<select data-pk="e"' + dis + '>' + options([['0', 'مشتريات'], ['1', 'مصروفات']], v.e ? '1' : '0') + '</select>' : '') + '</td></tr>';
    }
    return h + '</tbody></table></div></div></div>';
  }
  function bindParties() {
    $$('[data-party] [data-pk]').forEach(el => el.addEventListener('change', () => {
      const k = el.closest('[data-party]').dataset.party, pk = el.dataset.pk;
      const v = pk === 'o' ? (Number(el.value) || null) : pk === 'e' ? el.value === '1' : pk === 'n' ? (el.value || null) : el.value.trim();
      saveParties(ui.pClient, { [k]: { [pk]: v } }).then(() => toast('تم الحفظ'));
    }));
  }

  /* ---------------- actions ---------------- */
  function cur() { const p = parse(decodeURIComponent(location.hash.slice(1)).split('-').slice(1).join('-')); return p.view === 'ret' ? retById(p.id) : null; }
  const EDIT_ACTS = ['vat-import', 'vat-add-inv', 'vat-bulk', 'vat-multi', 'vat-toggle-x', 'vat-undo', 'vat-prev-auto', 'vat-upload-pdf', 'vat-manual-sub', 'vat-clear-sub', 'vat-del-ret', 'vat-accept'];
  function act(a, el) {
    const r = cur();
    if (r && r.locked && EDIT_ACTS.includes(a)) return toast('الإقرار مقفول بعد المطابقة. افتح القفل الأول.', true);
    switch (a) {
      case 'vat-new': return A.writable() ? openNew({ client: el.dataset.client, period: el.dataset.period }) : null;
      case 'vat-pick': ui.vClient = el.dataset.client; return A.go('vat-list');
      case 'vat-dash-clear': ui.dClient = ''; ui.dStaff = ''; return A.renderPage();
      case 'vat-hist-clear': { const h = H(); h.invs = []; h.files = []; return A.renderPage(); }
      case 'vat-hist-paste': return openHistPaste();
      case 'vat-hist-copy': return histCopySystem();
      case 'vat-hist-save': return saveHist();
      case 'vat-lock': { const x = retById(el.dataset.id); return x && A.writable() ? write(() => S.db.doc('vatReturns/' + x.id).update({ locked: el.dataset.v === '1' })).then(ok => ok && toast(el.dataset.v === '1' ? 'اتقفل الإقرار' : 'اتفتح القفل')) : null; }
      case 'vat-accept': {
        const x = retById(el.dataset.id); if (!x || !x.sub || !A.writable()) return;
        if (el.dataset.v === '0') return write(() => S.db.doc('vatReturns/' + x.id).update({ sub: clean(Object.assign({}, x.sub, { accepted: null })) }));
        const bg = modal('اعتماد الفرق كمقبول', fld('ac-note', 'سبب الفرق (يظهر في التقارير)', inp('ac-note', ''), true), '<button type="button" class="btn" data-close>إلغاء</button><button type="button" class="btn primary" id="ac-go">اعتماد</button>');
        $('#ac-go', bg).addEventListener('click', async () => { if (await write(() => S.db.doc('vatReturns/' + x.id).update({ sub: clean(Object.assign({}, x.sub, { accepted: { note: val('ac-note', bg), at: new Date().toISOString() } })) }))) closeModal(); });
        return;
      }
      case 'vat-pay': return openPay(el.dataset.id);
      case 'vat-pay-del': return delPay(el.dataset.id, el.dataset.pid);
      case 'vat-pay-bulk': return openBulkPay();
      case 'vat-pay-cfg': return openPayCfg();
      case 'vat-pay-remind': return openRemind(el.dataset.client);
      case 'vat-due-age': ui.duAge = ui.duAge === el.dataset.v ? '' : el.dataset.v; return location.hash.startsWith('#vat-due') ? A.renderPage() : A.go('vat-due');
      case 'vat-due-clear': ui.duClient = ''; ui.duStaff = ''; ui.duAge = ''; return A.renderPage();
      case 'vat-bulk': return r ? bulk(r, el.dataset.op) : null;
      case 'vat-list-status': ui.vStatus = el.dataset.v; return A.go('vat-list');
      case 'vat-book': ui.rbSide = el.dataset.v; return A.go('vat-rep-books');
      case 'vat-multi': return r && A.writable() ? openMultiSales(r) : null;
      case 'vat-undo': return r ? undo(r.id) : null;
      case 'vat-filter': ui.vStatus = ui.vStatus === el.dataset.v ? '' : el.dataset.v; return A.renderPage();
      case 'vat-import': { const f = $('#vat-file'); return f ? f.click() : null; }
      case 'vat-upload-pdf': { const f = $('#vat-pdf'); return f ? f.click() : null; }
      case 'vat-add-inv': return r ? openInvoice(r, null, { side: el.dataset.side, dt: el.dataset.dt }) : null;
      case 'vat-edit-inv': { const e = r && VS.cache.get(r.id); const inv = e && e.rows.find(i => i.id === el.dataset.id); return inv ? openInvoice(r, inv) : null; }
      case 'vat-toggle-x': {
        const e = r && VS.cache.get(r.id); const inv = e && e.rows.find(i => i.id === el.dataset.id); if (!inv) return;
        snap(r.id, (inv.x ? 'إرجاع ' : 'استبعاد ') + (inv.no || '')); inv.x = !inv.x; if (inv.x) inv.xr = 'استبعاد يدوي'; else delete inv.xr; saveSoon(r.id); return A.renderPage();
      }
      case 'vat-dl': return r ? downloadSheet(r, el.dataset.side, el.dataset.fmt) : null;
      case 'vat-pdf-sum': return r ? pdfSummary(r) : null;
      case 'vat-view-sub': return r ? viewSubmitted(r) : null;
      case 'vat-dl-sub': return downloadSubmitted(retById(el.dataset.id));
      case 'vat-manual-sub': return r ? openManualSub(r) : null;
      case 'vat-clear-sub': return r ? confirmBox('حذف بيانات التقديم', 'ستُحذف بيانات الإقرار المقدَّم ونتيجة المطابقة، وتبقى الفواتير كما هي.', 'حذف', async () => {
        if (await write(async () => { await deleteFile(r); await S.db.doc('vatReturns/' + r.id).update({ sub: null, status: 'ready' }); })) { closeModal(); toast('تم الحذف'); }
      }) : null;
      case 'vat-prev-auto': return r ? write(() => S.db.doc('vatReturns/' + r.id).update({ prevMode: 'auto' })).then(ok => { if (ok) saveSoon(r.id); }) : null;
      case 'vat-del-ret': {
        const x = retById(el.dataset.id); if (!x) return;
        return confirmBox('حذف الإقرار', 'سيُحذف إقرار ' + A.clientName(x.client) + ' عن ' + periodLabel(x) + ' بجميع فواتيره. الحذف نهائي.', 'حذف', async () => {
          const ok = await write(async () => { for (let k = 0; k < (x.chunks || 0); k++) await S.db.doc('vatInvoices/' + x.id + '-' + k).delete(); await deleteFile(x); await S.db.doc('vatReturns/' + x.id).delete(); });
          if (ok) { VS.cache.delete(x.id); closeModal(); A.go('vat-list'); toast('تم حذف الإقرار'); }
        });
      }
      case 'vat-rep-pdf': { const rep = repData(el.dataset.key); return A.makePdf(A.fileName(rep.title + ' ' + (rep.ctx || '')) + '.pdf', A.pdfHead(rep.title, rep.ctx) + A.tableHtml(rep) + (rep.extra ? '<h2>' + esc(rep.extra.title) + '</h2>' + A.tableHtml(rep.extra) : ''), rep.landscape); }
      case 'vat-rep-xlsx': {
        const rep = repData(el.dataset.key);
        const aoa = (x) => { const rows = (x.groups ? [x.groups.flatMap(([l, n]) => [l || null].concat(new Array(n - 1).fill(null)))] : []).concat([x.cols.map(c => c.h)], x.rows.map(y => y.map(v => v === '' ? null : v))); if (x.foot && x.rows.length) rows.push(x.foot); return rows; };
        return A.xlsxBook(A.fileName(rep.title) + '.xlsx', [{ name: rep.title, ctx: rep.ctx, rows: aoa(rep) }].concat(rep.extra ? [{ name: rep.extra.title, ctx: rep.ctx, rows: aoa(rep.extra) }] : []));
      }
    }
  }

  /* top bar menus and the «where am I» trail for this section */
  function navMenu(k, W) {
    const L = (label, href) => ({ label, href }), A = (label, act, data) => W ? { label, act, data: data || {} } : null;
    if (k === 'vat') return [L('لوحة الإقرارات', '#vat'), A('+ إقرار جديد', 'vat-new')];
    if (k === 'vat-list') return [L('جميع الإقرارات', '#vat-list'), W ? L('⤒ إقرارات قديمة من شيت البوابة', '#vat-hist') : null, { label: 'المسودات', nav: 1, act: 'vat-list-status', data: { v: 'draft' } }, { label: 'جاهزة للتقديم', nav: 1, act: 'vat-list-status', data: { v: 'ready' } },
      { label: 'مقدَّمة ومطابقة', nav: 1, act: 'vat-list-status', data: { v: 'matched' } }, { label: 'مقدَّمة وبها فروق', nav: 1, act: 'vat-list-status', data: { v: 'diff' } }, A('+ إقرار جديد', 'vat-new')];
    if (k === 'vat-rep') return REPS.map(([r, l]) => L(l, '#vat-rep-' + r)).concat([{ label: 'دفتر المبيعات', nav: 1, act: 'vat-book', data: { v: 'S' } }, { label: 'دفتر المشتريات والمصروفات', nav: 1, act: 'vat-book', data: { v: 'P' } }]).filter(x => x.href !== '#vat-rep-books');
    if (k === 'vat-parties') return [L('الموردون والعملاء', '#vat-parties')];
    if (k === 'vat-due') return [L('كل المستحقات', '#vat-due'), { label: 'المتأخر السداد فقط', nav: 1, act: 'vat-due-age', data: { v: 'late' } }, L('تقرير الضريبة المستحقة', '#vat-rep-due'), A('⚙ إعدادات المتابعة', 'vat-pay-cfg')];
    return null;
  }
  function crumbs(sub) {
    const p = parse(sub);
    if (p.view === 'dash') return [];
    if (p.view === 'list') return [['إقرارات القيمة المضافة', '#vat-list']];
    if (p.view === 'hist') return [['إقرارات القيمة المضافة', '#vat-list'], ['الإقرارات القديمة', '']];
    if (p.view === 'rep') return [['تقارير الإقرارات', '#vat-rep'], [(REPS.find(x => x[0] === p.key) || REPS[0])[1], '']];
    if (p.view === 'parties') return [['الموردون والعملاء', '']];
    if (p.view === 'due') return [['الضريبة المستحقة السداد', p.client ? '#vat-due' : '']].concat(p.client ? [['كشف حساب ' + A.clientName(p.client), '']] : []);
    const r = retById(p.id);
    return [['إقرارات القيمة المضافة', '#vat-list']].concat(r ? [[A.clientName(r.client) + ' – ' + periodLabel(r), '#vat-ret-sum-' + r.id], [(TABS.find(t => t[0] === p.tab) || TABS[0])[1], '']] : []);
  }
  function homeStats() {
    if (!S.M) return { open: 0, diff: 0, due: 0, late: 0 }; const all = rets(), d = dueSummary();
    return { open: all.filter(r => !r.sub).length, diff: all.filter(r => statusOf(r) === 'diff').length, due: d.total, late: d.nLate }; }

  document.addEventListener('keydown', (ev) => {
    if (!(ev.ctrlKey || ev.metaKey) || ev.key.toLowerCase() !== 'z' || ev.shiftKey) return;
    if (ev.target.closest && ev.target.closest('input,textarea,select,.modal')) return;
    const r = cur(); if (!r || !A.writable() || r.locked) return;
    ev.preventDefault(); undo(r.id);
  });
  window.OFVat = { page, bind, act, allChunks, allFiles, homeStats, navMenu, crumbs, dueSummary, payInfo, resetCache: () => { VS.cache.clear(); VS.undo.clear(); VS.sel.clear(); } };
})();
