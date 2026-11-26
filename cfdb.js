/* Cloudflare mode: when the program is opened from its own website (Cloudflare Pages), data lives in the office's
   database on Cloudflare and every employee signs in with a username and password the manager creates.
   Keeps the same small db API the app uses everywhere (collection/doc, onSnapshot, get/set/update/delete),
   and checks the server every few seconds so changes by colleagues appear without a refresh. */
(function () {
  'use strict';
  const host = location.hostname;
  const forced = /[?&]cf=1\b/.test(location.search);
  const local = location.protocol === 'file:' || /[?&]local=1\b/.test(location.search);
  let framed = false; try { framed = window.top !== window.self; } catch (e) { framed = true; }   // claude.ai shows the program inside a frame
  if (window.claude || local || (framed && !forced) || (!forced && (!/^https?:$/.test(location.protocol) || /claude|anthropic|usercontent|localhost|127\.0\.0\.1/i.test(host)))) return;

  const HEAVY = ['vatInvoices', 'vatFiles'];
  const POLL_MS = 5000, OVERLAP_MS = 30000;
  const mem = new Map(), colSubs = new Map(), docSubs = new Map();
  let me = null, since = 0, pollTimer = null, onAuth = null;

  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const copy = (v) => v == null ? v : JSON.parse(JSON.stringify(v));
  const deepFreeze = (o) => { if (o && typeof o === 'object' && !Object.isFrozen(o)) { Object.freeze(o); for (const k of Object.keys(o)) deepFreeze(o[k]); } return o; };
  const frozen = (v) => deepFreeze(copy(v));
  const parentOf = (path) => path.split('/').slice(0, -1).join('/');

  async function api(path, opts) {
    opts = opts || {};
    const init = { method: opts.method || 'GET', credentials: 'same-origin', headers: {} };
    if (opts.body !== undefined) { init.headers['content-type'] = 'application/json'; init.body = JSON.stringify(opts.body); }
    let r;
    try { r = await fetch('/api/' + path, init); } catch (e) { throw new Error('تعذّر الاتصال بالخادم. تأكد من الإنترنت.'); }
    let j = null; try { j = await r.json(); } catch (e) {}
    if (r.status === 401 && !opts.noAuthRedirect) { me = null; stopPoll(); showLogin(false); }
    if (!r.ok) { const err = new Error((j && j.error) || ('خطأ ' + r.status)); err.status = r.status; throw err; }
    return j;
  }

  /* ---------- login screen ---------- */
  function overlay() {
    let o = document.getElementById('cf-login');
    if (!o) {
      o = document.createElement('div'); o.id = 'cf-login'; o.dir = 'rtl';
      o.style.cssText = 'position:fixed;inset:0;z-index:9999;display:flex;align-items:center;justify-content:center;background:#f5f5f4;font-family:inherit;padding:16px';
      document.body.appendChild(o);
    }
    return o;
  }
  function showLogin(setup) {
    const o = overlay();
    o.innerHTML = '<form id="cf-form" style="background:#fff;border:1px solid #e5e5e5;border-radius:14px;padding:28px;width:100%;max-width:380px;box-shadow:0 10px 30px rgba(0,0,0,.06)">' +
      '<h1 style="margin:0 0 4px;font-size:22px">مكتب المحاسبة</h1>' +
      '<p style="margin:0 0 18px;color:#666;font-size:14px">' + (setup ? 'أول تشغيل: أنشئ حساب مدير البرنامج. بعد كده تضيف حسابات الموظفين من داخل البرنامج.' : 'سجّل الدخول بحسابك') + '</p>' +
      (setup ? '<label style="display:block;font-size:13px;margin-bottom:4px">الاسم</label><input name="name" autocomplete="name" style="width:100%;box-sizing:border-box;padding:10px;border:1px solid #ccc;border-radius:8px;margin-bottom:12px;font:inherit">' : '') +
      '<label style="display:block;font-size:13px;margin-bottom:4px">اسم المستخدم</label><input name="username" dir="ltr" autocomplete="username" autocapitalize="none" required style="width:100%;box-sizing:border-box;padding:10px;border:1px solid #ccc;border-radius:8px;margin-bottom:12px;font:inherit;text-align:right">' +
      '<label style="display:block;font-size:13px;margin-bottom:4px">كلمة المرور</label><input name="password" type="password" dir="ltr" autocomplete="' + (setup ? 'new-password' : 'current-password') + '" required style="width:100%;box-sizing:border-box;padding:10px;border:1px solid #ccc;border-radius:8px;margin-bottom:16px;font:inherit;text-align:right">' +
      '<div id="cf-err" role="alert" style="color:#b91c1c;font-size:13px;min-height:18px;margin-bottom:8px"></div>' +
      '<button type="submit" style="width:100%;padding:11px;border:0;border-radius:8px;background:#1c1c1c;color:#fff;font:inherit;font-weight:600;cursor:pointer">' + (setup ? 'إنشاء حساب المدير' : 'دخول') + '</button></form>';
    o.style.display = 'flex';
    const f = o.querySelector('#cf-form');
    setTimeout(() => f.querySelector('input').focus(), 0);
    f.addEventListener('submit', async (ev) => {
      ev.preventDefault();
      const fd = new FormData(f), btn = f.querySelector('button'), err = f.querySelector('#cf-err');
      btn.disabled = true; err.textContent = '';
      try {
        const r = await api(setup ? 'setup' : 'login', { method: 'POST', body: { name: fd.get('name') || '', username: fd.get('username'), password: fd.get('password') }, noAuthRedirect: true });
        signedIn(r.user);
      } catch (e) { err.textContent = e.message; btn.disabled = false; }
    });
  }
  function signedIn(user) {
    const o = document.getElementById('cf-login'); if (o) o.style.display = 'none';
    if (onAuth && !onAuth.done) { me = user; onAuth.done = true; onAuth.resolve(); return; }
    location.reload();   // signed in again after the session ended (role or person may differ): start clean
  }

  /* ---------- data ---------- */
  function colSnap(col, filters) {
    const docs = [];
    mem.forEach((v, k) => { if (parentOf(k) === col && (!filters || filters.every(([f, op, val]) => op !== '==' || v[f] === val))) docs.push({ id: k.slice(col.length + 1), exists: true, data: () => frozen(v) }); });
    return { docs, size: docs.length, empty: !docs.length };
  }
  const docSnap = (path, v) => ({ id: path.split('/').pop(), exists: v != null, data: () => frozen(v) });
  function notify(paths) {
    const cols = new Set(paths.map(parentOf));
    cols.forEach(col => (colSubs.get(col) || new Set()).forEach(fn => { try { fn(); } catch (e) { console.error(e); } }));
    paths.forEach(p => (docSubs.get(p) || new Set()).forEach(fn => { try { fn(); } catch (e) { console.error(e); } }));
  }
  function apply(list) {
    const changed = [];
    for (const { p, d } of list) {
      if (HEAVY.includes(parentOf(p))) continue;
      const before = mem.has(p) ? JSON.stringify(mem.get(p)) : null, after = d == null ? null : JSON.stringify(d);
      if (before === after) continue;
      if (d == null) mem.delete(p); else mem.set(p, d);
      changed.push(p);
    }
    if (changed.length) notify(changed);
  }
  async function sync(full) {
    const r = await api('sync?since=' + (full ? 0 : Math.max(0, since - OVERLAP_MS)));
    if (full) mem.clear();
    apply(r.docs);
    since = Math.max(since, r.now || 0);
  }
  function stopPoll() { clearTimeout(pollTimer); pollTimer = null; }
  function schedulePoll() {
    stopPoll();
    pollTimer = setTimeout(async () => { if (me && document.visibilityState === 'visible') { try { await sync(false); } catch (e) {} } schedulePoll(); }, POLL_MS);
  }
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && me) sync(false).catch(() => {}); });

  const isHeavy = (path) => HEAVY.includes(parentOf(path));
  const enc = encodeURIComponent;
  function docRef(path) {
    return {
      id: path.split('/').pop(), path,
      async get() {
        if (!isHeavy(path)) return docSnap(path, mem.get(path));
        const r = await api('doc?p=' + enc(path)); return docSnap(path, r.exists ? r.data : null);
      },
      async set(data) { const r = await api('doc?p=' + enc(path), { method: 'PUT', body: { data: copy(data) } }); if (!isHeavy(path)) { mem.set(path, r.data); notify([path]); } },
      async update(patch) { const r = await api('doc?p=' + enc(path), { method: 'PATCH', body: { data: copy(patch) } }); if (!isHeavy(path)) { mem.set(path, r.data); notify([path]); } },
      async delete() { await api('doc?p=' + enc(path), { method: 'DELETE' }); if (mem.has(path)) { mem.delete(path); notify([path]); } },
      onSnapshot(fn) {
        const w = () => fn(docSnap(path, mem.get(path)));
        if (!docSubs.has(path)) docSubs.set(path, new Set()); docSubs.get(path).add(w); setTimeout(w, 0);
        return () => docSubs.get(path).delete(w);
      }
    };
  }
  function colRef(col, filters) {
    filters = filters || [];
    return {
      doc: (id) => docRef(col + '/' + id),
      where: (f, op, v) => colRef(col, filters.concat([[f, op, v]])),
      get: async () => colSnap(col, filters),
      onSnapshot(fn) {
        const w = () => fn(colSnap(col, filters));
        if (!colSubs.has(col)) colSubs.set(col, new Set()); colSubs.get(col).add(w); setTimeout(w, 0);
        return () => colSubs.get(col).delete(w);
      }
    };
  }

  const ready = (async () => {
    await new Promise(res => document.readyState === 'loading' ? document.addEventListener('DOMContentLoaded', res, { once: true }) : res());
    let st;
    try { st = await api('status', { noAuthRedirect: true }); }
    catch (e) {
      const o = overlay(); o.innerHTML = '<div style="max-width:420px;text-align:center"><h2>تعذّر تشغيل البرنامج</h2><p style="color:#666">' + esc(e.message) + '</p></div>';
      return new Promise(() => {});
    }
    if (st.user) me = st.user;
    else await new Promise(resolve => { onAuth = { resolve, done: false }; showLogin(st.needsSetup); });
    if (!onAuth) onAuth = { done: true };
    await sync(true);
    schedulePoll();
  })();

  const db = { collection: (p) => colRef(p), doc: (p) => docRef(p) };
  const downloads = {
    async save({ filename, data }) {
      const blob = data instanceof Blob ? data : new Blob([data]);
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename;
      document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
    }
  };
  const user = {
    isOwner: async () => !!me && me.role === 'admin',
    canEdit: async () => !!me && me.role === 'admin',
    can: async (n) => n === 'data.write' ? !!me && me.role !== 'viewer' : null,
    id: async () => me ? me.username : '',
    me: async () => me
  };
  window.OF_CF = {
    me: () => me,
    async logout() { try { await api('logout', { method: 'POST', noAuthRedirect: true }); } catch (e) {} location.reload(); },
    changePassword: (current, password) => api('password', { method: 'POST', body: { current, password } }),
    listUsers: () => api('users').then(r => r.users),
    saveUser: (u) => api('users', { method: 'POST', body: u }),
    deleteUser: (username) => api('users?u=' + enc(username), { method: 'DELETE' })
  };
  window.claude = { use: async (name) => { await ready; return name === 'db' ? db : name === 'downloads' ? downloads : name === 'user' ? user : null; } };
})();
