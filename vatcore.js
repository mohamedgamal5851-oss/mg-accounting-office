/* VAT return core (Form 10, Egyptian Tax Authority). Pure functions, no DOM, no db.
   An invoice keeps the codes of the ETA upload templates so the upload sheets are a straight projection:
     side 'S' (sales) / 'P' (purchases)
     dt   نوع المستند   sales 1 فاتورة · 2 إشعار إضافة · 3 إشعار خصم · 5 مستهلك نهائي محلي · 6 جهة حكومية · 7/8 مستهلك أجنبي
                        purchases 1 فاتورة · 2 إشعار إضافة · 3 إشعار خصم · 4 إذن إفراج
   each line:  n  (sales only) 'G' سلعة / 'S' خدمة
               o  نوع السلعة   sales 1 محلي · 2 صادرات · 5 آلات ومعدات · 6 أجزاء آلات · 7 إعفاءات
                               purchases 3 سلع · 4 خدمات · 5 آلات ومعدات · 6 أجزاء آلات · 7 إعفاءات
   Amounts are kept as given (up to 5 decimals); the return rounds only its totals, as the portal does. */
(function (root) {
  'use strict';

  const r2 = (n) => Math.round((Number(n) || 0) * 100) / 100;
  const r5 = (n) => Math.round((Number(n) || 0) * 100000) / 100000;
  const num = (v) => {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return v;
    const s = toLatin(String(v)).replace(/[,\s]/g, '').replace('٫', '.');
    const neg = /-$/.test(s) || /^-/.test(s) || /^\(.*\)$/.test(s);
    const n = parseFloat(s.replace(/[()\-]/g, ''));
    return isNaN(n) ? 0 : (neg ? -n : n);
  };
  function toLatin(s) {
    return String(s).replace(/[٠-٩]/g, d => String(d.charCodeAt(0) - 0x660)).replace(/[۰-۹]/g, d => String(d.charCodeAt(0) - 0x6F0)).replace(/٫/g, '.').replace(/٬/g, ',');
  }
  const pad2 = (n) => String(n).padStart(2, '0');
  const norm = (s) => String(s == null ? '' : s).normalize('NFKC').replace(/[ً-ْـ]/g, '').replace(/[إأآ]/g, 'ا').replace(/ى/g, 'ي').replace(/ة/g, 'ه').replace(/\*/g, '').replace(/\s+/g, ' ').trim();
  const digitsOnly = (s) => toLatin(String(s == null ? '' : s)).replace(/\D/g, '');

  /* dates: ISO from 'dd/mm/yyyy', 'dd.mm.yyyy', 'yyyy-mm-dd', Excel serial or Date */
  function isoDate(v) {
    if (v == null || v === '') return '';
    if (v instanceof Date && !isNaN(v)) return v.getFullYear() + '-' + pad2(v.getMonth() + 1) + '-' + pad2(v.getDate());
    if (typeof v === 'number' && v > 20000 && v < 80000) { const d = new Date(Math.round((v - 25569) * 86400000)); return d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate()); }
    const s = toLatin(String(v)).trim();
    let m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/); if (m) return m[1] + '-' + pad2(m[2]) + '-' + pad2(m[3]);
    m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/); if (m) return m[3] + '-' + pad2(m[2]) + '-' + pad2(m[1]);
    return '';
  }
  const dotDate = (iso) => iso ? iso.split('-').reverse().join('.') : '';

  /* periods: a return covers one month, or a quarter (Law 6 of 2025) ending in `period` (YYYY-MM) */
  function periodBounds(period, freq) {
    const [y, m] = period.split('-').map(Number);
    const startM = freq === 'quarterly' ? m - 2 : m;
    const sy = startM < 1 ? y - 1 : y, sm = startM < 1 ? startM + 12 : startM;
    const last = new Date(y, m, 0).getDate();
    return { from: sy + '-' + pad2(sm) + '-01', to: period + '-' + pad2(last) };
  }
  function quarterEnd(ym) { const [y, m] = ym.split('-').map(Number); return y + '-' + pad2(Math.ceil(m / 3) * 3); }

  /* ---------- classification ---------- */
  const SERVICE_RX = /خدم|service|fee|عمول|commission|اشتراك|subscri|شحن|shipping|deliver|توصيل|نقل|ايجار|إيجار|rent|صيان|mainten|استشار|consult|اعلان|إعلان|advert|marketing|تسويق|اتعاب|أتعاب|تدريب|training|حجز|booking|license|ترخيص|hosting|استضاف|برمج|software/i;
  const isServiceText = (d) => SERVICE_RX.test(String(d || ''));
  const VALID_REG = (r) => /^\d{9}$/.test(r);
  function snapRate(net, vat) {
    if (!net || Math.abs(vat) < 0.005) return 0;
    const r = Math.abs(vat / net) * 100;
    for (const k of [14, 5, 10, 8, 1]) if (Math.abs(r - k) < 0.6) return k;
    return Math.round(r * 100) / 100;
  }
  /* default codes for a fresh line. memory = what the office set before for this party (o for purchases, n for sales) */
  function defaultLine(side, line, mem) {
    const l = Object.assign({ d: '', code: '', unit: '', q: 1, p: 0, net: 0, vat: 0 }, line);
    l.net = r5(l.net); l.vat = r5(l.vat);
    if (l.rate == null) l.rate = snapRate(l.net, l.vat);
    if (side === 'P') {
      if (l.o == null) l.o = mem && mem.o ? mem.o : (l.rate === 0 ? 7 : isServiceText(l.d) ? 4 : 3);
      // e = expense: kept in the purchases book (and for Form 41 later) but outside the VAT return and the upload sheet
      if (l.e == null) l.e = !!(mem && mem.e);
    } else {
      if (l.n == null) l.n = mem && mem.n ? mem.n : (isServiceText(l.d) ? 'S' : 'G');
      if (l.o == null) l.o = l.rate === 0 ? 7 : 1;
    }
    return l;
  }
  const invKey = (inv) => inv.side + '|' + (inv.reg || inv.name || '') + '|' + String(inv.no || '').trim() + '|' + (inv.dt || 1);
  const isExp = (l) => !!(l && l.e);
  const vatLines = (inv) => (inv.lines || []).filter(l => !isExp(l));
  /* an invoice the return uses: not excluded, and (for purchases) at least one item that is not an expense */
  const inReturn = (inv) => !inv.x && vatLines(inv).length > 0;
  function invTotals(inv) {
    let net = 0, vat = 0;
    for (const l of inv.lines || []) { net += Number(l.net) || 0; vat += Number(l.vat) || 0; }
    return { net: r5(net), vat: r5(vat), total: r5(net + vat) };
  }

  /* ---------- reading spreadsheets (array-of-arrays) ---------- */
  function headerIndex(row, wanted) {
    const h = row.map(norm);
    const out = {};
    for (const [k, names] of Object.entries(wanted)) {
      out[k] = -1;
      for (const n of names) { const i = h.indexOf(norm(n)); if (i >= 0) { out[k] = i; break; } }
    }
    return out;
  }
  const ETA_INV = {
    serial: ['مسلسل'], type: ['نوع المستند'], status: ['حالة الفاتورة', 'الحالة'], no: ['رقم الفاتورة', 'الرقم الداخلى', 'الرقم الداخلي'],
    seller: ['اسم البائع'], sellerReg: ['رقم التسجيل الضريبي للبائع'], buyer: ['اسم المشتري'], buyerReg: ['رقم التسجيل الضريبي للمشتري'],
    issued: ['تاريخ الاصدار', 'تاريخ الإصدار'], net: ['ثمن الفاتورة'], vat: ['ضريبة القيمة المضافة'], wht: ['ضريبة الخصم'], other: ['ضرائب ورسوم أخرى'], tbl: ['ضريبة الجدول'], total: ['اجمالي الفاتورة', 'إجمالي الفاتورة']
  };
  const ETA_LINE = {
    serial: ['مسلسل'], type: ['نوع المستند'], status: ['الحالة'], no: ['الرقم الداخلى', 'الرقم الداخلي'], issued: ['تاريخ الإصدار', 'تاريخ الاصدار'],
    desc: ['الوصف'], code: ['كود الصنف'], unit: ['الوحدة'], q: ['الكمية'], p: ['سعر الوحدة'], net: ['إجمالي البند (قبل الضريبة)'], vat: ['ضريبة القيمة المضافة'], tbl: ['ضريبة الجدول']
  };
  const etaDocType = (t, side, reg) => {
    const s = norm(t);
    if (/دائن|خصم/.test(s)) return 3;
    if (/مدين|اضاف/.test(s)) return 2;
    return side === 'S' && !VALID_REG(reg) ? 5 : 1;
  };
  const etaValid = (st) => { const s = norm(st); return !s || /صحيح|valid|مقبول/i.test(s); };

  /* workbook = { sheetName: aoa }. Returns { kind, invoices, skipped, notes } or null when not recognized */
  function readWorkbook(book, memFor) {
    const names = Object.keys(book);
    const find = (rx) => names.find(n => rx.test(norm(n)));
    const sent = find(/^فواتير مرسله/), recv = find(/^فواتير مستلمه/);
    if (sent || recv) return readEta(book, sent, recv, find(/بنود الفواتير المرسله/), find(/بنود الفواتير المستلمه/), memFor);
    for (const n of names) {
      const aoa = book[n]; if (!aoa || !aoa.length) continue;
      const codes = (aoa[1] || []).map(v => String(v || '').trim());
      if (codes.includes('DOC_TYP_S') || codes.includes('DOC_TYP_P')) return readTemplate(aoa, codes.includes('DOC_TYP_S') ? 'S' : 'P', memFor);
      const h = (aoa[0] || []).map(norm);
      if (h.includes(norm('نوع المستند *')) && h.includes(norm('رقم الفاتورة *'))) return readTemplate(aoa, h.includes(norm('اسم المورد *')) ? 'P' : 'S', memFor);
    }
    return null;
  }

  function readEta(book, sentName, recvName, sentLinesName, recvLinesName, memFor) {
    const invoices = [], notes = []; let skipped = 0;
    function side(name, linesName, s) {
      if (!name) return;
      const aoa = book[name]; if (!aoa || !aoa.length) return;
      const H = headerIndex(aoa[0], ETA_INV);
      // lines grouped by internal number + issue date, consumed in sheet order
      const groups = new Map();
      if (linesName && book[linesName] && book[linesName].length) {
        const la = book[linesName], L = headerIndex(la[0], ETA_LINE);
        for (const r of la.slice(1)) {
          if (typeof r[L.serial] !== 'number' || !etaValid(r[L.status])) continue;
          const k = String(r[L.no] == null ? '' : r[L.no]).trim() + '|' + isoDate(r[L.issued]);
          if (!groups.has(k)) groups.set(k, []);
          groups.get(k).push({ d: String(r[L.desc] || ''), code: String(r[L.code] || ''), unit: String(r[L.unit] || ''), q: num(r[L.q]) || 1, p: num(r[L.p]), net: num(r[L.net]), vat: num(r[L.vat]), tbl: num(r[L.tbl]) });
        }
      }
      for (const r of aoa.slice(1)) {
        if (typeof r[H.serial] !== 'number') continue;
        if (!etaValid(r[H.status])) { skipped++; continue; }
        const no = String(r[H.no] == null ? '' : r[H.no]).trim();
        const reg = digitsOnly(s === 'P' ? r[H.sellerReg] : r[H.buyerReg]);
        const name = String((s === 'P' ? r[H.seller] : r[H.buyer]) || '').trim();
        const date = isoDate(r[H.issued]);
        const net = num(r[H.net]), vat = num(r[H.vat]);
        const inv = { side: s, src: 'eta', dt: etaDocType(r[H.type], s, reg), no, date, name, reg: VALID_REG(reg) ? reg : '', nid: !VALID_REG(reg) && reg.length === 14 ? reg : '', file: '', addr: '', phone: '',
          wht: num(r[H.wht]), tbl: num(r[H.tbl]), lines: [] };
        const mem = memFor ? memFor(s, inv.reg || inv.name) : null;
        if (mem && mem.addr) inv.addr = mem.addr;
        if (mem && mem.file) inv.file = mem.file;
        // take this invoice's lines when they add up to its value; otherwise one summary line
        const g = groups.get(no + '|' + date);
        let lines = null;
        if (g && g.length) {
          let acc = 0, take = 0;
          while (take < g.length && Math.abs(acc - net) > 0.05) { acc += g[take].net; take++; }
          if (Math.abs(acc - net) <= 0.05) { lines = g.splice(0, take); }
        }
        if (!lines) lines = [{ d: name ? 'فاتورة ' + (s === 'P' ? 'مشتريات من ' : 'مبيعات إلى ') + name : 'فاتورة', q: 1, p: net, net, vat }];
        inv.lines = lines.map(l => defaultLine(s, l, mem));
        invoices.push(inv);
      }
    }
    side(sentName, sentLinesName, 'S');
    side(recvName, recvLinesName, 'P');
    if (skipped) notes.push('تم تجاهل ' + skipped + ' مستندًا حالته غير «صحيحة» (ملغاة أو مرفوضة).');
    return { kind: 'eta', invoices, skipped, notes };
  }

  /* the ETA upload templates themselves (filled by hand or by another program): one row per item */
  function readTemplate(aoa, s, memFor) {
    const invoices = [], byKey = new Map(), notes = [];
    for (const r of aoa.slice(2)) {
      if (!r || r.every(v => v == null || v === '')) continue;
      const dt = Number(num(r[0])) || 1, no = String(r[3] == null ? '' : r[3]).trim();
      if (!no) continue;
      const reg = digitsOnly(r[5]), date = isoDate(r[10]);
      const k = dt + '|' + no + '|' + reg + '|' + date;
      let inv = byKey.get(k);
      if (!inv) {
        inv = { side: s, src: 'sheet', dt, no, date, name: String(r[4] || '').trim(), reg: VALID_REG(reg) ? reg : '', file: String(r[6] == null ? '' : r[6]), addr: String(r[7] || ''), nid: String(r[8] == null ? '' : r[8]), phone: String(r[9] == null ? '' : r[9]), lines: [] };
        byKey.set(k, inv); invoices.push(inv);
      }
      const N = Number(num(r[13])), O = Number(num(r[14]));
      const line = { d: String(r[11] || ''), code: String(r[12] == null ? '' : r[12]), unit: String(r[15] || ''), p: num(r[16]), q: num(r[18]) || 1, net: num(r[21]) || num(r[19]) - num(r[20]), vat: num(r[22]), rate: r[17] === '' || r[17] == null ? null : num(r[17]) };
      if (s === 'S') { if (N === 3 || N === 4) line.n = N === 4 ? 'S' : 'G'; if (O) line.o = O; }
      else if (O) line.o = O;
      inv.lines.push(defaultLine(s, line, memFor ? memFor(s, inv.reg || inv.name) : null));
    }
    return { kind: 'template', invoices, skipped: 0, notes };
  }

  /* ---------- checks ---------- */
  const ID_LIMIT = 150000;
  const V_TOTAL = (inv) => invTotals(inv).total;
  const NO_OK = /^[A-Za-z0-9؀-ۿ \-/\\]+$/;
  /* invoice numbers keep digits only (letters, «-», «/» and other marks are removed, office rule); a number with no digits stays as it was */
  const normNo = (no) => { const d = digitsOnly(no); return d || cleanNo(no); };
  const cleanNo = (no) => String(no || '').replace(/[^A-Za-z0-9؀-ۿ \-/\\]/g, '-');
  function issues(inv, ret) {
    const out = [];
    if (inv.x) return out;
    if (inv.side === 'P' && !inReturn(inv)) return out; // expenses only: not uploaded
    if (!inv.no) out.push(['crit', 'رقم المستند ناقص']);
    else if (!digitsOnly(inv.no)) out.push(['crit', 'رقم المستند لا يحتوي على أرقام']);
    if (!inv.date) out.push(['crit', 'التاريخ ناقص']);
    else if (ret && (inv.date < ret.from || inv.date > ret.to)) out.push(['warn', 'التاريخ خارج فترة الإقرار']);
    if (inv.prior) out.push(['warn', 'مدرجة أيضًا في ' + inv.prior]);
    if (!inv.name) out.push(['crit', 'اسم ' + (inv.side === 'P' ? 'المورد' : 'العميل') + ' ناقص']);
    if (inv.side === 'S' && inv.dt === 5 && inv.reg) out.push(['info', 'مستهلك نهائي: رقم التسجيل لا يُكتب في ملف الرفع']);
    if (inv.dt === 1 && !VALID_REG(inv.reg)) out.push(['crit', 'رقم التسجيل الضريبي يجب أن يكون 9 أرقام مع نوع المستند «فاتورة»']);
    const t = invTotals(inv);
    // national ID of the buyer: required when a final-consumer document exceeds 150,000 (Minister of Finance decree 324 of 2022)
    if (inv.side === 'S' && inv.dt === 5 && V_TOTAL(inv) > ID_LIMIT && !inv.nid) out.push(['crit', 'مستهلك نهائي بمبلغ يتجاوز 150,000: الرقم القومي إلزامي']);
    if (inv.side === 'S' && inv.dt === 7 && t.net >= 5000 && !inv.nid) out.push(['crit', 'مستهلك أجنبي بمبلغ 5,000 أو أكثر: رقم جواز السفر إلزامي']);
    if (!(inv.lines || []).some(l => Number(l.net) > 0)) out.push(['crit', 'لا توجد قيمة']);
    if (inv.side === 'P' && (inv.lines || []).some(l => l.o === 6)) out.push(['info', 'نوع السلعة 6 «أجزاء آلات» يُستخدم مع ممولي الجدول والقيمة المضافة فقط (ملاحظة شيت المشتريات)']);
    if (inv.side === 'P' && inv.dt === 4 && (inv.lines || []).some(l => l.o === 4)) out.push(['warn', 'إذن الإفراج الجمركي للسلع المستوردة؛ راجع تصنيف البند «خدمات»']);
    if ((inv.lines || []).some(l => Number(l.net) < 0 || Number(l.vat) < 0)) out.push(['crit', 'قيمة سالبة: المنظومة لا تقبل قيمًا سالبة، والمرتجعات تُسجَّل «إشعار خصم» بقيم موجبة']);
    for (const l of vatLines(inv)) {
      if (Number(l.net) > 0 && Number(l.vat) === 0 && l.o !== 7 && !(inv.side === 'S' && l.o === 2)) { out.push(['warn', 'بند بدون ضريبة ولم يُصنَّف إعفاء أو تصدير']); break; }
    }
    return out;
  }

  /* ---------- the return (Form 10) ---------- */
  const RATE_ROWS = [['5', '%5'], ['14', '%14'], ['0', 'صفر'], ['ex', 'إعفاء']];
  const IN_ROWS = [[3, 'سلع'], [4, 'خدمات'], [5, 'الآلات والمعدات'], [6, 'أجزاء الآلات والمعدات وقطع الغيار'], [7, 'إعفاء']];
  function salesRow(l) { if (l.o === 7) return 'ex'; if (l.o === 2 || !Number(l.rate)) return '0'; return Number(l.rate) === 5 ? '5' : '14'; }

  function compute(invoices, prevCredit) {
    const sales = {}, inputs = {};
    for (const [k] of RATE_ROWS) sales[k] = { g: 0, s: 0, ap: 0, am: 0, total: 0, tax: 0 };
    for (const [k] of IN_ROWS) inputs[k] = { loc: 0, imp: 0, ap: 0, am: 0, total: 0, tax: 0 };
    const counts = { sInv: 0, sDebit: 0, sCredit: 0, sOther: 0, pInv: 0, pRelease: 0, pDebit: 0, pCredit: 0, excluded: 0 };
    let sFirst = null, sLast = null;
    for (const inv of invoices) {
      if (inv.x) { counts.excluded++; continue; }
      if (!inReturn(inv)) { counts.expOnly = (counts.expOnly || 0) + 1; continue; }
      if (inv.side === 'S') {
        if (inv.dt === 2) counts.sDebit++; else if (inv.dt === 3) counts.sCredit++; else if (inv.dt === 1 || inv.dt === 6) counts.sInv++; else counts.sOther++;
        if (inv.dt !== 2 && inv.dt !== 3) { if (sFirst == null || inv.no < sFirst) sFirst = inv.no; if (sLast == null || inv.no > sLast) sLast = inv.no; }
        for (const l of inv.lines) {
          const row = sales[salesRow(l)], net = Number(l.net) || 0, vat = Number(l.vat) || 0;
          if (inv.dt === 2) { row.ap += net; row.tax += vat; }
          else if (inv.dt === 3) { row.am += net; row.tax -= vat; }
          else { if (l.n === 'S') row.s += net; else row.g += net; row.tax += vat; }
        }
      } else {
        if (inv.dt === 4) counts.pRelease++; else if (inv.dt === 2) counts.pDebit++; else if (inv.dt === 3) counts.pCredit++; else counts.pInv++;
        for (const l of vatLines(inv)) {
          const row = inputs[IN_ROWS.some(r => r[0] === l.o) ? l.o : 3], net = Number(l.net) || 0, vat = Number(l.vat) || 0;
          if (inv.dt === 2) { row.ap += net; row.tax += vat; }
          else if (inv.dt === 3) { row.am += net; row.tax -= vat; }
          else { if (inv.dt === 4) row.imp += net; else row.loc += net; row.tax += vat; }
        }
      }
    }
    const tot = (o, keys) => { const t = {}; for (const k of keys) t[k] = 0; for (const r of Object.values(o)) for (const k of keys) t[k] += r[k]; return t; };
    for (const r of Object.values(sales)) { r.total = r.g + r.s + r.ap - r.am; for (const k in r) r[k] = r2(r[k]); }
    for (const r of Object.values(inputs)) { r.total = r.loc + r.imp + r.ap - r.am; for (const k in r) r[k] = r2(r[k]); }
    const sT = tot(sales, ['g', 's', 'ap', 'am', 'total', 'tax']), iT = tot(inputs, ['loc', 'imp', 'ap', 'am', 'total', 'tax']);
    for (const k in sT) sT[k] = r2(sT[k]); for (const k in iT) iT[k] = r2(iT[k]);
    const out = sT.tax, inp = iT.tax, due = r2(out - inp), prev = r2(prevCredit || 0), res = r2(due - prev);
    return { sales, salesTotal: sT, inputs, inputsTotal: iT, out, inp, due, prev, res, payable: res > 0 ? res : 0, credit: res < 0 ? -res : 0, counts, sFirst, sLast };
  }

  /* ---------- upload sheets ---------- */
  const S_HEAD = ['نوع المستند *', 'نوع الضريبة *', 'نوع سلع الجدول *', 'رقم الفاتورة *', 'اسم العميل *', 'رقم التسجيل الضريبي للعميل *', 'رقم الملف الضريبي للعميل', 'العنوان *', 'الرقم القومي / رقم جواز السفر', 'رقم الموبيل', 'تاريخ الفاتورة *', 'إسم المنتج *', 'كود المنتج', 'نوع البيان *', 'نوع السلعة *', 'وحدة قياس المنتج', 'سعر الوحدة *', 'فئة الضريبة *', 'كمية المنتج *', 'المبلغ الإجمالي *', 'قيمة الخصم', 'المبلغ الصافي *', 'قيمة الضريبة *', 'الإجمالي *'];
  const S_CODES = ['DOC_TYP_S', 'TAX_TYP', 'TYP_SCH', 'INVNUM', 'CUST_NAME', 'CUST_REG_NO', 'CUST_FILE_NO', 'ADDRESS', 'ID', 'PHONE', 'DOC_TIN_NUM', 'PROD_NAME', 'PROD_CODE', 'STAT_TYPE_S', 'COMM_TYPE_S', 'PMU', 'UNITP', 'TAX_CAT', 'QOP', 'TAMMOUNT', 'DISC', 'NET_AMM', 'TAX_VAL', 'TOTAL'];
  const P_HEAD = S_HEAD.map((h, i) => i === 4 ? 'اسم المورد *' : i === 5 ? 'رقم التسجيل الضريبي للمورد' : i === 6 ? 'رقم الملف الضريبي للمورد' : h);
  const P_CODES = S_CODES.map(c => c === 'DOC_TYP_S' ? 'DOC_TYP_P' : c === 'STAT_TYPE_S' ? 'STAT_TYPE_P' : c === 'COMM_TYPE_S' ? 'COMM_TYPE_P' : c);

  /* فئة الضريبة: 14 or 5; left empty for zero-rated and exempt items (office rule).
     credit notes (إشعار دائن) are doc type 3 «إشعار خصم» with statement type 5 «تسويات»; amounts stay positive, discount is 0.
     mode 'group': one row per invoice for each (type, rate) · mode 'items': one row per item */
  function uploadRows(invoices, side, mode, defAddr) {
    const rows = [];
    for (const inv of invoices) {
      if (inv.side !== side || !inReturn(inv)) continue;
      let parts;
      if (mode === 'items') parts = vatLines(inv).filter(l => Number(l.net) > 0).map(l => Object.assign({ cnt: 1 }, l));
      else {
        const m = new Map();
        for (const l of vatLines(inv)) {
          if (!(Number(l.net) > 0)) continue;
          const k = (l.n || '') + '|' + l.o + '|' + l.rate;
          if (!m.has(k)) m.set(k, { d: l.d, n: l.n, o: l.o, rate: l.rate, net: 0, vat: 0, cnt: 0, code: '', unit: '' });
          const g = m.get(k); g.net += Number(l.net); g.vat += Number(l.vat); g.cnt++;
        }
        parts = Array.from(m.values()).map(g => Object.assign(g, { net: r5(g.net), vat: r5(g.vat), q: 1, p: r5(g.net), d: (String(g.d || '').slice(0, 90) + (g.cnt > 1 ? ' وأصناف أخرى' : '')).trim() }));
      }
      for (const l of parts) {
        const adj = inv.dt === 2 || inv.dt === 3;
        const N = side === 'S' ? (adj ? 5 : l.n === 'S' ? 4 : 3) : (adj ? 5 : inv.dt === 4 ? 2 : 1);
        let q = Number(l.q) || 1, p = Number(l.p) || 0;
        // template notes: unit price and quantity accept no zero or negative values (price at least 0.01)
        if (!(p >= 0.01) || !(q > 0) || Math.abs(q * p - l.net) > 0.05) { q = 1; p = l.net; }
        const consumer = side === 'S' && inv.dt >= 5 && inv.dt !== 6;
        rows.push([inv.dt, 1, 0, normNo(inv.no), inv.name || '', consumer ? '' : (inv.reg || ''), consumer ? '' : (inv.file || ''), inv.addr || defAddr || '', inv.nid || '', inv.phone || '', dotDate(inv.date),
          String(l.d || (side === 'S' ? 'مبيعات' : 'مشتريات')).slice(0, 100) || '—', l.code || '', N, l.o, l.unit || '', r5(p), Number(l.rate) ? Number(l.rate) : '', r5(q), r5(l.net), 0, r5(l.net), r5(l.vat), r5(Number(l.net) + Number(l.vat))]);
      }
    }
    return rows;
  }
  function uploadSheet(invoices, side, mode, defAddr) {
    return { name: side === 'S' ? 'مستندات المبيعات' : 'Purchase Final', head: side === 'S' ? S_HEAD : P_HEAD, codes: side === 'S' ? S_CODES : P_CODES, rows: uploadRows(invoices, side, mode, defAddr) };
  }

  /* ---------- the submitted return (PDF from the portal) ----------
     pages: array of arrays of text items in pdf.js order. Numbers are Arabic-Indic; a trailing "-" item marks a negative. */
  function parseReturnPdf(pages) {
    const out = { formNo: '', kind: '', from: '', to: '', sentDate: '', summary: null, numbers: [], counts: null, raw: '' };
    const isNum = (s) => /^-?[0-9٠-٩][0-9٠-٩٫.,]*-?$/.test(s);
    let best = null;
    pages.forEach((items, pi) => {
      const its = items.map(s => String(s || '').trim()).filter(Boolean);
      const N = its.map(s => norm(s));
      const joined = toLatin(its.join(' '));
      out.raw += joined + '\n';
      if (!out.formNo) { const m = joined.match(/\b(\d{11})\b/); if (m) out.formNo = m[1]; }
      if (!out.kind) { if (N.some(s => /^اصيل$|^اصلي$/.test(s))) out.kind = 'original'; else if (N.some(s => /^معدل$/.test(s))) out.kind = 'amended'; }
      const ds = joined.match(/\b(20\d\d)\/(\d\d)\/(\d\d)\b/g) || [];
      if (ds.length >= 2 && !out.from) { out.from = ds[0].replace(/\//g, '-'); out.to = ds[1].replace(/\//g, '-'); }
      for (let i = 0; i < its.length; i++) {
        if (isNum(its[i])) {
          let v = num(its[i]);
          if (its[i + 1] === '-' || /-$/.test(its[i])) v = -Math.abs(v);
          out.numbers.push(v);
        }
      }
      // tax summary: the numbers after the last «ملخص» label, read right to left: output, inputs, due, previous credit, result
      for (let i = 0; i < N.length; i++) {
        if (!/ملخص/.test(N[i])) continue;
        let j = i + 1, nums = [];
        while (j < its.length && nums.length < 6) {
          if (isNum(its[j]) && !/\($/.test(its[j - 1] || '')) { let v = num(its[j]); if (its[j + 1] === '-') { v = -Math.abs(v); j++; } nums.push(v); }
          else if (nums.length && !/^[()\-+]$/.test(its[j])) break;
          j++;
        }
        if (nums.length >= 5) {
          const r = nums.slice(0, 5).reverse(); // [out, inp, due, prev, res]
          const s = { out: r[0], inp: r[1], due: r[2], prev: r[3], res: r[4], page: pi + 1 };
          if (!best || (Math.abs(s.out) + Math.abs(s.inp) + Math.abs(s.prev)) > (Math.abs(best.out) + Math.abs(best.inp) + Math.abs(best.prev))) best = s;
        }
      }
      // counts block: «أصدرنا المستندات التالية»
      const k = N.findIndex(s => /اصدرنا/.test(s));
      if (k >= 0) {
        const nums = [];
        for (let j = k + 1; j < its.length && nums.length < 14; j++) if (/^[0-9٠-٩]+$/.test(its[j])) nums.push(Number(toLatin(its[j])));
        if (nums.length >= 14) out.counts = { sInv: nums[2], sDebit: nums[5], sCredit: nums[8], sOther: nums[11], pInv: nums[12], pRelease: nums[13] };
      }
    });
    // the send date is the first date written with Arabic digits after the period dates
    const dates = (out.raw.replace(/\s*\/\s*/g, '/').match(/\b20\d\d\/\d{1,2}\/\d{1,2}\b/g) || []).map(d => { const [y, m, dd] = d.split('/'); return y + '-' + pad2(m) + '-' + pad2(dd); });
    out.sentDate = dates.find(d => d !== out.from && d !== out.to) || '';
    out.summary = best;
    return out;
  }

  /* compare the system's return with the submitted one */
  function reconcile(calc, sub, ret) {
    const has = (v) => sub.numbers && sub.numbers.some(x => Math.abs(Math.abs(x) - Math.abs(v)) < 0.015);
    const rows = [];
    const add = (label, mine, theirs, strict) => {
      const t = theirs == null ? (has(mine) ? mine : null) : theirs;
      rows.push({ label, mine, theirs: t, ok: t != null && Math.abs(r2(mine) - r2(t)) < 0.015, strict: !!strict });
    };
    const S = sub.summary || {};
    add('ضريبة القيمة المضافة على المبيعات (1)', calc.out, S.out, 1);
    add('ضريبة المدخلات (2)', calc.inp, S.inp, 1);
    add('الضريبة المستحقة (1) − (2)', calc.due, S.due, 1);
    add('الرصيد الدائن السابق', calc.prev, S.prev, 1);
    add(calc.res >= 0 ? 'الصافي المستحق السداد' : 'الرصيد الدائن المرحّل', calc.res, S.res, 1);
    add('إجمالي وعاء المبيعات', calc.salesTotal.total);
    add('إجمالي قيمة المدخلات', calc.inputsTotal.total);
    for (const [k, l] of RATE_ROWS) if (calc.sales[k].total) { add('مبيعات ' + l + ' – القيمة', calc.sales[k].total); add('مبيعات ' + l + ' – الضريبة', calc.sales[k].tax); }
    for (const [k, l] of IN_ROWS) if (calc.inputs[k].total) { add('مدخلات ' + l + ' – القيمة', calc.inputs[k].total); add('مدخلات ' + l + ' – الضريبة', calc.inputs[k].tax); }
    if (sub.counts) {
      const c = calc.counts;
      rows.push({ label: 'عدد فواتير المبيعات', mine: c.sInv + c.sOther, theirs: sub.counts.sInv + sub.counts.sOther, ok: c.sInv + c.sOther === sub.counts.sInv + sub.counts.sOther, isCount: 1 });
      rows.push({ label: 'عدد فواتير الشراء', mine: c.pInv, theirs: sub.counts.pInv, ok: c.pInv === sub.counts.pInv, isCount: 1 });
      if (c.pRelease || sub.counts.pRelease) rows.push({ label: 'عدد أذون الإفراج', mine: c.pRelease, theirs: sub.counts.pRelease, ok: c.pRelease === sub.counts.pRelease, isCount: 1 });
    }
    const periodOk = !sub.from || !ret || (sub.from === ret.from && sub.to === ret.to);
    return { rows, periodOk, ok: periodOk && rows.every(r => r.ok) };
  }

  const API = { normNo, ID_LIMIT, isExp, vatLines, inReturn, r2, r5, num, toLatin, isoDate, dotDate, periodBounds, quarterEnd, snapRate, defaultLine, invKey, invTotals, isServiceText, readWorkbook, issues, cleanNo,
    RATE_ROWS, IN_ROWS, compute, uploadSheet, S_HEAD, P_HEAD, S_CODES, P_CODES, parseReturnPdf, reconcile, VALID_REG, norm };
  if (typeof module === 'object' && module.exports) module.exports = API; else root.OFVatCore = API;
})(typeof window !== 'undefined' ? window : this);
