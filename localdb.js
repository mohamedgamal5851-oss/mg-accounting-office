/* Local mode: when the program is opened from a folder on the computer (file://), there is no claude.ai storage.
   This keeps the same small db API the app uses (collection/doc, onSnapshot, get/set/update/delete) on top of the
   browser's IndexedDB, so all data stays on this computer, in this browser. Downloads save straight to the Downloads folder. */
(function () {
  'use strict';
  const local = location.protocol === 'file:' || /[?&]local=1\b/.test(location.search);
  if (!local || window.claude) return;

  const DB_NAME = 'accounting-office-local', STORE = 'docs';
  const mem = new Map();               // path -> data
  const colSubs = new Map();           // collection path -> Set(fn)
  const docSubs = new Map();           // doc path -> Set(fn)
  let idb = null;

  function openIdb() {
    return new Promise((res) => {
      if (!window.indexedDB) return res(null);
      const rq = indexedDB.open(DB_NAME, 1);
      rq.onupgradeneeded = () => rq.result.createObjectStore(STORE);
      rq.onsuccess = () => res(rq.result);
      rq.onerror = () => res(null);
    });
  }
  function loadAll() {
    return new Promise((res) => {
      if (!idb) {
        try { const raw = localStorage.getItem(DB_NAME); if (raw) for (const [k, v] of Object.entries(JSON.parse(raw))) mem.set(k, v); } catch (e) {}
        return res();
      }
      const tx = idb.transaction(STORE, 'readonly'), st = tx.objectStore(STORE), rq = st.openCursor();
      rq.onsuccess = () => { const c = rq.result; if (c) { mem.set(c.key, c.value); c.continue(); } else res(); };
      rq.onerror = () => res();
    });
  }
  function persist(path, data) {
    if (!idb) { try { const o = {}; mem.forEach((v, k) => { o[k] = v; }); localStorage.setItem(DB_NAME, JSON.stringify(o)); } catch (e) { return Promise.reject(new Error('مساحة التخزين المحلي امتلأت.')); } return Promise.resolve(); }
    return new Promise((res, rej) => {
      const tx = idb.transaction(STORE, 'readwrite'), st = tx.objectStore(STORE);
      if (data === undefined) st.delete(path); else st.put(data, path);
      tx.oncomplete = () => res(); tx.onerror = () => rej(tx.error || new Error('تعذّر الحفظ على هذا الجهاز.'));
    });
  }
  const parentOf = (path) => path.split('/').slice(0, -1).join('/');
  const copy = (v) => v == null ? v : JSON.parse(JSON.stringify(v));
  // records read back are frozen, as they are on claude.ai, so the program behaves the same in both places
  const deepFreeze = (o) => { if (o && typeof o === 'object' && !Object.isFrozen(o)) { Object.freeze(o); for (const k of Object.keys(o)) deepFreeze(o[k]); } return o; };
  const frozen = (v) => deepFreeze(copy(v));
  function colSnap(col) {
    const docs = [];
    mem.forEach((v, k) => { if (parentOf(k) === col) docs.push({ id: k.slice(col.length + 1), data: () => frozen(v), exists: true }); });
    return { docs, size: docs.length, empty: !docs.length };
  }
  const docSnap = (path) => ({ id: path.split('/').pop(), exists: mem.has(path), data: () => frozen(mem.get(path)) });
  function notify(path) {
    const col = parentOf(path);
    (colSubs.get(col) || new Set()).forEach(fn => { try { fn(colSnap(col)); } catch (e) { console.error(e); } });
    (docSubs.get(path) || new Set()).forEach(fn => { try { fn(docSnap(path)); } catch (e) { console.error(e); } });
  }
  async function put(path, data) { mem.set(path, copy(data)); await persist(path, mem.get(path)); notify(path); }
  async function del(path) { mem.delete(path); await persist(path, undefined); notify(path); }

  function docRef(path) {
    return {
      id: path.split('/').pop(), path,
      get: async () => docSnap(path),
      set: (data) => put(path, data),
      update: (patch) => { if (!mem.has(path)) return Promise.reject(new Error('السجل غير موجود.')); return put(path, Object.assign({}, mem.get(path), copy(patch))); },
      delete: () => del(path),
      onSnapshot(fn) { if (!docSubs.has(path)) docSubs.set(path, new Set()); docSubs.get(path).add(fn); setTimeout(() => fn(docSnap(path)), 0); return () => docSubs.get(path).delete(fn); }
    };
  }
  function colRef(col, filters) {
    filters = filters || [];
    const snap = () => { const s = colSnap(col); s.docs = s.docs.filter(d => filters.every(([f, op, v]) => op === '==' ? (d.data() || {})[f] === v : true)); return s; };
    return {
      doc: (id) => docRef(col + '/' + id),
      where: (f, op, v) => colRef(col, filters.concat([[f, op, v]])),
      get: async () => snap(),
      onSnapshot(fn) { if (!colSubs.has(col)) colSubs.set(col, new Set()); const w = () => fn(snap()); colSubs.get(col).add(w); setTimeout(w, 0); return () => colSubs.get(col).delete(w); }
    };
  }
  const ready = (async () => { idb = await openIdb(); await loadAll(); })();
  const db = { collection: (p) => colRef(p), doc: (p) => docRef(p) };
  const downloads = {
    async save({ filename, data }) {
      const blob = data instanceof Blob ? data : new Blob([data]);
      const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = filename;
      document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
    }
  };
  const user = { isOwner: async () => true, can: async () => true, id: async () => 'local' };
  window.OF_LOCAL = true;
  window.claude = { use: async (name) => { await ready; return name === 'db' ? db : name === 'downloads' ? downloads : name === 'user' ? user : null; } };
})();
