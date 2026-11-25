/* Accounting office core logic: pure functions over plain data. No DOM, no db.
   Client balances are derived from the transaction log: fee claims and expenses paid on the client's behalf
   are debits; collections and discounts are credits. */
(function (root) {
  'use strict';

  const DEBIT = { fee: 1, expense: 1 };
  const CREDIT = { payment: 1, discount: 1 };
  const KINDS = ['fee', 'expense', 'payment', 'discount'];

  const pad2 = (n) => String(n).padStart(2, '0');
  function todayISO() {
    const d = new Date();
    return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
  }
  function lastDay(y, m) { return new Date(y, m, 0).getDate(); }
  function addMonths(ym, n) {
    let [y, m] = ym.split('-').map(Number);
    m += n; while (m < 1) { m += 12; y--; } while (m > 12) { m -= 12; y++; }
    return y + '-' + pad2(m);
  }
  function monthDiff(a, b) { // months from a to b (YYYY-MM)
    const [ya, ma] = a.split('-').map(Number), [yb, mb] = b.split('-').map(Number);
    return (yb - ya) * 12 + (mb - ma);
  }
  function monthBounds(ym) {
    const [y, m] = ym.split('-').map(Number);
    return { start: ym + '-01', end: ym + '-' + pad2(lastDay(y, m)) };
  }
  function daysBetween(a, b) { return Math.round((Date.parse(b) - Date.parse(a)) / 86400000); }
  function addDays(iso, n) {
    const d = new Date(Date.parse(iso) + n * 86400000);
    return d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate());
  }
  const round2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
  const sign = (kind) => (DEBIT[kind] ? 1 : CREDIT[kind] ? -1 : 0);

  function sortTx(list) {
    return list.slice().sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : (a.seq || 0) - (b.seq || 0)));
  }

  /* Due date of an obligation for the period ending in month `periodYM`.
     dueOffset = months after the period's month; dueDay 0 = last day of that month. */
  function dueDate(ob, periodYM) {
    const ym = addMonths(periodYM, Number(ob.dueOffset || 0));
    const [y, m] = ym.split('-').map(Number);
    const ld = lastDay(y, m);
    const d = Number(ob.dueDay || 0);
    return ym + '-' + pad2(d <= 0 || d > ld ? ld : d);
  }

  /* Clients under Law 6 of 2025 (simplified regime) may follow a different frequency for an obligation,
     e.g. a quarterly VAT return instead of monthly. Older saved settings have no freqLaw6, so VAT defaults to quarterly. */
  const isLaw6 = (client) => !!client && client.taxLaw === '6-2025';
  function effFreq(ob, client) {
    if (isLaw6(client)) {
      const f = ob.freqLaw6 != null ? ob.freqLaw6 : (ob.id === 'vat' ? 'quarterly' : '');
      if (f) return f;
    }
    return ob.freq;
  }
  /* Which clients an obligation is for: by tax law (Law 91 of 2005 / Law 6 of 2025) and by kind of taxpayer
     (legal persons vs natural persons and sole proprietorships). Older saved settings lack these fields,
     so the built-in annual returns fall back to their defaults by id. */
  const NATURAL_FORMS = ['منشأة فردية', 'شخص طبيعي (مهنة حرة)'];
  const OB_LAW_DEFAULT = { corp: '91-2005', indiv: '91-2005', simple25: '6-2025' };
  const OB_ENTITY_DEFAULT = { corp: 'legal', indiv: 'natural' };
  const obLaw = (ob) => ob.law || OB_LAW_DEFAULT[ob.id] || 'all';
  const obEntity = (ob) => ob.entity || OB_ENTITY_DEFAULT[ob.id] || 'all';
  function entityOf(client) {
    const f = (client && client.legalForm) || '';
    if (!f || f === 'أخرى') return '';
    return NATURAL_FORMS.includes(f) ? 'natural' : 'legal';
  }
  function obligationAllowed(ob, client) {
    const law = obLaw(ob);
    if (law !== 'all' && law !== (isLaw6(client) ? '6-2025' : '91-2005')) return false;
    const ent = obEntity(ob), ce = entityOf(client);
    if (ent !== 'all' && ce && ent !== ce) return false;
    return true;
  }
  /* Does an obligation have a period ending in `ym` for this client? */
  function obligationApplies(ob, client, ym) {
    const m = Number(ym.slice(5, 7)), f = effFreq(ob, client);
    if (f === 'monthly') return true;
    if (f === 'quarterly') return m % 3 === 0;
    if (f === 'annual') return m === Number(client.fyEnd || 12);
    return false;
  }

  /* Does a contract bill in month `ym`? */
  function contractBills(c, ym) {
    if (!c.active && c.active !== undefined) return false;
    const start = (c.startDate || '').slice(0, 7);
    if (!start || ym < start) return false;
    if (c.endDate && ym > c.endDate.slice(0, 7)) return false;
    const d = monthDiff(start, ym);
    if (c.freq === 'monthly') return true;
    if (c.freq === 'quarterly') return d % 3 === 0;
    if (c.freq === 'semiannual') return d % 6 === 0;
    if (c.freq === 'annual') return d % 12 === 0;
    if (c.freq === 'once') return d === 0;
    return false;
  }

  /* FIFO aging: credits settle the oldest debits first; what is left is aged by its date. */
  function aging(opening, openingDate, txs, today) {
    const debits = [];
    let credit = 0;
    if (opening > 0) debits.push({ date: openingDate || '1900-01-01', left: opening });
    else if (opening < 0) credit += -opening;
    for (const t of txs) {
      const s = sign(t.kind), a = Number(t.amount) || 0;
      if (s > 0) debits.push({ date: t.date, left: a });
      else if (s < 0) credit += a;
    }
    for (const d of debits) { const use = Math.min(d.left, credit); d.left -= use; credit -= use; if (credit <= 0) credit = 0; }
    const b = { b0: 0, b30: 0, b60: 0, b90: 0, oldest: null, advance: round2(credit) };
    for (const d of debits) {
      if (d.left <= 0.004) continue;
      const days = Math.max(0, daysBetween(d.date, today));
      if (days <= 30) b.b0 += d.left; else if (days <= 60) b.b30 += d.left; else if (days <= 90) b.b60 += d.left; else b.b90 += d.left;
      if (!b.oldest || d.date < b.oldest) b.oldest = d.date;
    }
    b.b0 = round2(b.b0); b.b30 = round2(b.b30); b.b60 = round2(b.b60); b.b90 = round2(b.b90);
    b.total = round2(b.b0 + b.b30 + b.b60 + b.b90 - b.advance);
    b.oldestDays = b.oldest ? daysBetween(b.oldest, today) : 0;
    return b;
  }

  function build(raw, today) {
    today = today || todayISO();
    const settings = raw.settings || {};
    const clients = (raw.clients || []).slice().sort((a, b) => String(a.code || '').localeCompare(String(b.code || ''), 'ar', { numeric: true }));
    const staff = raw.staff || [];
    const contracts = raw.contracts || [];
    const txs = sortTx(raw.transactions || []);
    const tasks = (raw.tasks || []).slice().sort((a, b) => (a.dueDate || '').localeCompare(b.dueDate || ''));
    const obligations = settings.obligations || [];

    const clientById = new Map(clients.map(c => [c.id, c]));
    const staffById = new Map(staff.map(s => [s.id, s]));
    const obById = new Map(obligations.map(o => [o.id, o]));
    const txByClient = new Map(clients.map(c => [c.id, []]));
    for (const t of txs) { if (!txByClient.has(t.client)) txByClient.set(t.client, []); txByClient.get(t.client).push(t); }

    const ym = today.slice(0, 7);
    const mb = monthBounds(ym);
    const acct = new Map();
    let totalDue = 0, totalAdvance = 0;
    for (const c of clients) {
      const list = txByClient.get(c.id) || [];
      const opening = Number(c.openingBalance) || 0;
      let debit = 0, credit = 0, lastPayment = null, billedMonth = 0, collectedMonth = 0;
      for (const t of list) {
        const a = Number(t.amount) || 0;
        if (sign(t.kind) > 0) debit += a; else credit += a;
        if (t.kind === 'payment') { if (!lastPayment || t.date > lastPayment) lastPayment = t.date; if (t.date >= mb.start && t.date <= mb.end) collectedMonth += a; }
        if (t.kind === 'fee' && t.date >= mb.start && t.date <= mb.end) billedMonth += a;
      }
      const balance = round2(opening + debit - credit);
      const ag = aging(opening, c.openingDate, list, today);
      if (balance > 0) totalDue += balance; else totalAdvance += -balance;
      acct.set(c.id, { opening, debit: round2(debit), credit: round2(credit), balance, lastPayment, aging: ag, billedMonth, collectedMonth, count: list.length });
    }

    // month figures
    let billedMonth = 0, collectedMonth = 0, expensesMonth = 0;
    for (const t of txs) {
      if (t.date < mb.start || t.date > mb.end) continue;
      const a = Number(t.amount) || 0;
      if (t.kind === 'fee') billedMonth += a;
      if (t.kind === 'payment') collectedMonth += a;
      if (t.kind === 'expense') expensesMonth += a;
    }

    // tasks
    const soonLimit = addDays(today, Number(settings.alertDays || 7));
    const taskState = (t) => (t.status === 'done' ? 'done' : !t.dueDate ? 'open' : t.dueDate < today ? 'overdue' : t.dueDate <= soonLimit ? 'soon' : 'open');
    const taskView = tasks.map(t => Object.assign({}, t, { state: taskState(t), daysLate: t.status !== 'done' && t.dueDate && t.dueDate < today ? daysBetween(t.dueDate, today) : 0 }));

    // documents
    const docLimit = addDays(today, Number(settings.docAlertDays || 30));
    const docs = [];
    for (const c of clients) for (const d of c.docs || []) {
      if (!d.expiry) continue;
      const st = d.expiry < today ? 'expired' : d.expiry <= docLimit ? 'soon' : 'ok';
      docs.push({ client: c, doc: d, state: st, days: daysBetween(today, d.expiry) });
    }
    docs.sort((a, b) => a.doc.expiry.localeCompare(b.doc.expiry));

    const active = clients.filter(c => c.status !== 'inactive');
    const model = {
      today, ym, settings, clients, active, staff, contracts, txs, tasks: taskView, obligations,
      clientById, staffById, obById, txByClient, acct, docs,
      totals: { due: round2(totalDue), advance: round2(totalAdvance), billedMonth: round2(billedMonth), collectedMonth: round2(collectedMonth), expensesMonth: round2(expensesMonth) }
    };
    model.alerts = alerts(model);
    return model;
  }

  function alerts(M) {
    const overdue = M.tasks.filter(t => t.state === 'overdue');
    const soon = M.tasks.filter(t => t.state === 'soon');
    const docsExpired = M.docs.filter(d => d.state === 'expired' && M.clientById.get(d.client.id).status !== 'inactive');
    const docsSoon = M.docs.filter(d => d.state === 'soon' && d.client.status !== 'inactive');
    const late90 = M.clients.filter(c => { const a = M.acct.get(c.id); return a && a.aging.b90 > 0; });
    const unbilled = feesDue(M, M.ym);
    const missingTasks = tasksToGenerate(M, M.ym === M.today.slice(0, 7) ? addMonths(M.ym, -1) : M.ym);
    return { overdue, soon, docsExpired, docsSoon, late90, unbilled, missingTasks };
  }

  /* Fee claims that the active contracts call for in `ym` and that have not been recorded yet. */
  function feesDue(M, ym) {
    const done = new Set(M.txs.filter(t => t.kind === 'fee' && t.contract && t.period).map(t => t.contract + '|' + t.period));
    const out = [];
    for (const c of M.contracts) {
      const cl = M.clientById.get(c.client);
      if (!cl || cl.status === 'inactive') continue;
      if (c.active === false) continue;
      if (!contractBills(c, ym)) continue;
      if (done.has(c.id + '|' + ym)) continue;
      out.push({ contract: c, client: cl, period: ym, amount: Number(c.amount) || 0 });
    }
    return out;
  }

  /* A client's obligations as they apply now: an annual income-tax return that the client's law or legal form
     rules out is replaced by the one that does apply (e.g. Law 6 clients file the simplified return, form 25). */
  const INCOME_GROUP = ['corp', 'indiv', 'simple25'];
  function effectiveObligations(M, c) {
    const out = [];
    for (const id of c.obligations || []) {
      const ob = M.obById.get(id);
      if (ob && INCOME_GROUP.includes(id) && !obligationAllowed(ob, c)) {
        const alt = INCOME_GROUP.map(x => M.obById.get(x)).find(o => o && o.active !== false && obligationAllowed(o, c));
        if (alt && !out.includes(alt.id)) out.push(alt.id);
        continue;
      }
      if (!out.includes(id)) out.push(id);
    }
    return out;
  }

  /* Procedures that the clients' obligations call for, for the period ending in `ym`, not yet created. */
  function tasksToGenerate(M, ym) {
    const have = new Set(M.tasks.filter(t => t.obligation && t.period).map(t => t.client + '|' + t.obligation + '|' + t.period));
    const out = [];
    for (const c of M.clients) {
      if (c.status === 'inactive') continue;
      if (c.startDate && c.startDate.slice(0, 7) > ym) continue;
      for (const oid of effectiveObligations(M, c)) {
        const ob = M.obById.get(oid);
        if (!ob || ob.active === false) continue;
        if (!obligationAllowed(ob, c)) continue;
        if (!obligationApplies(ob, c, ym)) continue;
        if (have.has(c.id + '|' + ob.id + '|' + ym)) continue;
        out.push({ client: c, obligation: ob, period: ym, dueDate: dueDate(ob, ym) });
      }
    }
    return out;
  }

  /* Account statement for a client between two dates (inclusive). */
  function statement(M, clientId, from, to) {
    const c = M.clientById.get(clientId);
    const list = M.txByClient.get(clientId) || [];
    let opening = Number(c && c.openingBalance) || 0;
    const rows = [];
    let debit = 0, credit = 0;
    for (const t of list) {
      const a = Number(t.amount) || 0, s = sign(t.kind);
      if (from && t.date < from) { opening += s * a; continue; }
      if (to && t.date > to) continue;
      if (s > 0) debit += a; else credit += a;
      rows.push(t);
    }
    opening = round2(opening);
    let bal = opening;
    const lines = rows.map(t => { const a = Number(t.amount) || 0, s = sign(t.kind); bal = round2(bal + s * a); return { t, debit: s > 0 ? a : 0, credit: s < 0 ? a : 0, balance: bal }; });
    return { client: c, from, to, opening, lines, debit: round2(debit), credit: round2(credit), closing: bal };
  }

  function nextNumber(txs, kind) {
    let n = 0;
    for (const t of txs) if (t.kind === kind) n = Math.max(n, Number(t.number) || 0);
    return n + 1;
  }
  function nextSeq(txs) { return txs.reduce((n, t) => Math.max(n, t.seq || 0), 0) + 1; }
  function nextClientCode(clients) {
    let max = 0;
    for (const c of clients) { const m = /(\d+)\s*$/.exec(c.code || ''); if (m) max = Math.max(max, Number(m[1])); }
    return 'C-' + String(max + 1).padStart(3, '0');
  }

  /* Income by service and month for a date range. */
  function revenue(M, from, to) {
    const byService = new Map(), byClient = new Map();
    let fees = 0, payments = 0, expenses = 0, discounts = 0;
    const contractById = new Map(M.contracts.map(c => [c.id, c]));
    for (const t of M.txs) {
      if (from && t.date < from) continue;
      if (to && t.date > to) continue;
      const a = Number(t.amount) || 0;
      if (t.kind === 'fee') {
        fees += a;
        const svc = t.service || (t.contract && contractById.get(t.contract) && contractById.get(t.contract).service) || '—';
        byService.set(svc, (byService.get(svc) || 0) + a);
        byClient.set(t.client, (byClient.get(t.client) || 0) + a);
      } else if (t.kind === 'payment') payments += a;
      else if (t.kind === 'expense') expenses += a;
      else if (t.kind === 'discount') discounts += a;
    }
    return { fees: round2(fees), payments: round2(payments), expenses: round2(expenses), discounts: round2(discounts), byService, byClient };
  }

  const api = { effectiveObligations, effFreq, isLaw6, obLaw, obEntity, entityOf, obligationAllowed, NATURAL_FORMS, build, statement, aging, feesDue, tasksToGenerate, dueDate, obligationApplies, contractBills, revenue, nextNumber, nextSeq, nextClientCode,
    todayISO, addMonths, monthDiff, monthBounds, daysBetween, addDays, round2, sign, KINDS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.OFCore = api;
})(typeof window !== 'undefined' ? window : globalThis);
