/* Cloudflare Pages Function: the office program's server.
   Needs one D1 database bound to the Pages project under the variable name DB; tables are created on first use.
   Stores every record as one JSON row keyed by its path (collection/id), the same shape the program uses everywhere,
   plus user accounts (PBKDF2 password hashes) and login sessions (random token in an HttpOnly cookie).
   Roles: admin (everything, users), staff (read and write data), viewer (read only). */

const COOKIE = 'of_s';
const SESSION_DAYS = 30;
const PBKDF2_ITER = 100000;
const HEAVY = ['vatInvoices', 'vatFiles'];          // read one record at a time, never in the bulk sync
const ADMIN_ONLY_WRITE = [/^vault\//, /^meta\/users/];
const ADMIN_ONLY_DELETE = [/^clients\//];
const PATH_RE = /^[A-Za-z0-9_]{1,40}\/[^\/\s]{1,200}$/;

let schemaReady = false;
async function ensureSchema(db) {
  if (schemaReady) return;
  await db.batch([
    db.prepare('CREATE TABLE IF NOT EXISTS docs (path TEXT PRIMARY KEY, col TEXT NOT NULL, data TEXT, deleted INTEGER NOT NULL DEFAULT 0, updated INTEGER NOT NULL, by TEXT)'),
    db.prepare('CREATE INDEX IF NOT EXISTS docs_col_upd ON docs (col, updated)'),
    db.prepare('CREATE INDEX IF NOT EXISTS docs_upd ON docs (updated)'),
    db.prepare('CREATE TABLE IF NOT EXISTS users (username TEXT PRIMARY KEY, name TEXT, role TEXT NOT NULL, salt TEXT NOT NULL, hash TEXT NOT NULL, active INTEGER NOT NULL DEFAULT 1, fails INTEGER NOT NULL DEFAULT 0, locked_until INTEGER NOT NULL DEFAULT 0, created TEXT, last_login TEXT)'),
    db.prepare('CREATE TABLE IF NOT EXISTS sessions (token TEXT PRIMARY KEY, username TEXT NOT NULL, expires INTEGER NOT NULL)')
  ]);
  schemaReady = true;
}

const json = (obj, status = 200, headers = {}) => new Response(JSON.stringify(obj), { status, headers: Object.assign({ 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store' }, headers) });
const fail = (msg, status = 400) => json({ error: msg }, status);
const hex = (buf) => Array.from(new Uint8Array(buf), b => b.toString(16).padStart(2, '0')).join('');
const randHex = (n) => hex(crypto.getRandomValues(new Uint8Array(n)));

async function hashPassword(password, saltHex) {
  const salt = new Uint8Array(saltHex.match(/../g).map(h => parseInt(h, 16)));
  const key = await crypto.subtle.importKey('raw', new TextEncoder().encode(password), 'PBKDF2', false, ['deriveBits']);
  return hex(await crypto.subtle.deriveBits({ name: 'PBKDF2', hash: 'SHA-256', salt, iterations: PBKDF2_ITER }, key, 256));
}
function sameHex(a, b) { if (a.length !== b.length) return false; let d = 0; for (let i = 0; i < a.length; i++) d |= a.charCodeAt(i) ^ b.charCodeAt(i); return d === 0; }
const cleanUser = (s) => String(s || '').trim().toLowerCase();
const validUser = (s) => /^[a-z0-9._-]{3,32}$/.test(s);
const ROLES = ['admin', 'staff', 'viewer'];
const pubUser = (u) => ({ username: u.username, name: u.name || u.username, role: u.role, active: !!u.active, created: u.created || '', lastLogin: u.last_login || '' });

function cookieOf(req) {
  const m = (req.headers.get('cookie') || '').match(new RegExp('(?:^|;\\s*)' + COOKIE + '=([a-f0-9]{64})'));
  return m ? m[1] : null;
}
function sessionCookie(req, token, maxAge) {
  const secure = new URL(req.url).protocol === 'https:' ? '; Secure' : '';
  return COOKIE + '=' + token + '; Path=/; HttpOnly; SameSite=Strict; Max-Age=' + maxAge + secure;
}
async function currentUser(db, req) {
  const t = cookieOf(req); if (!t) return null;
  const row = await db.prepare('SELECT u.* FROM sessions s JOIN users u ON u.username = s.username WHERE s.token = ? AND s.expires > ?').bind(t, Date.now()).first();
  return row && row.active ? row : null;
}
async function newSession(db, req, username) {
  const token = randHex(32);
  await db.prepare('INSERT INTO sessions (token, username, expires) VALUES (?, ?, ?)').bind(token, username, Date.now() + SESSION_DAYS * 864e5).run();
  await db.prepare('DELETE FROM sessions WHERE expires < ?').bind(Date.now()).run();
  return sessionCookie(req, token, SESSION_DAYS * 86400);
}

// strictly increasing change stamp, so clients can ask "what changed since X" without missing writes
async function stamp(db) {
  const r = await db.prepare('SELECT MAX(updated) AS m FROM docs').first();
  return Math.max(Date.now(), ((r && r.m) || 0) + 1);
}

async function readBody(req) {
  if (!(req.headers.get('content-type') || '').includes('application/json')) throw new Error('bad content type');
  return req.json();
}

export async function onRequest(ctx) {
  const { request: req, env } = ctx;
  const db = env.DB;
  if (!db) return fail('لم تُربط قاعدة البيانات بعد: أضف ربط D1 باسم DB في إعدادات المشروع على Cloudflare ثم أعد النشر.', 500);
  try {
    await ensureSchema(db);
    const url = new URL(req.url);
    const route = url.pathname.replace(/^\/api\/?/, '').replace(/\/$/, '');
    const method = req.method;

    /* ---------- accounts ---------- */
    if (route === 'status' && method === 'GET') {
      const n = await db.prepare('SELECT COUNT(*) AS n FROM users').first();
      const u = await currentUser(db, req);
      return json({ needsSetup: !(n && n.n), user: u ? pubUser(u) : null });
    }
    if (route === 'setup' && method === 'POST') {
      const n = await db.prepare('SELECT COUNT(*) AS n FROM users').first();
      if (n && n.n) return fail('تم إعداد البرنامج من قبل.', 403);
      const b = await readBody(req), username = cleanUser(b.username), password = String(b.password || '');
      if (!validUser(username)) return fail('اسم المستخدم من 3 إلى 32 حرفًا إنجليزيًا أو رقمًا.');
      if (password.length < 8) return fail('كلمة المرور 8 أحرف على الأقل.');
      const salt = randHex(16);
      await db.prepare('INSERT INTO users (username, name, role, salt, hash, active, created, last_login) VALUES (?, ?, ?, ?, ?, 1, ?, ?)').bind(username, String(b.name || '').trim() || username, 'admin', salt, await hashPassword(password, salt), new Date().toISOString(), new Date().toISOString()).run();
      const u = await db.prepare('SELECT * FROM users WHERE username = ?').bind(username).first();
      return json({ user: pubUser(u) }, 200, { 'set-cookie': await newSession(db, req, username) });
    }
    if (route === 'login' && method === 'POST') {
      const b = await readBody(req), username = cleanUser(b.username), password = String(b.password || '');
      const u = await db.prepare('SELECT * FROM users WHERE username = ?').bind(username).first();
      const bad = () => fail('اسم المستخدم أو كلمة المرور غير صحيحة.', 401);
      if (!u || !u.active) { await hashPassword(password, '00'.repeat(16)); return bad(); }
      if (u.locked_until > Date.now()) return fail('تم إيقاف الدخول مؤقتًا بعد محاولات خاطئة متكررة. حاول بعد ' + Math.ceil((u.locked_until - Date.now()) / 60000) + ' دقيقة.', 429);
      if (!sameHex(await hashPassword(password, u.salt), u.hash)) {
        const fails = u.fails + 1;
        await db.prepare('UPDATE users SET fails = ?, locked_until = ? WHERE username = ?').bind(fails >= 8 ? 0 : fails, fails >= 8 ? Date.now() + 15 * 60000 : 0, username).run();
        return bad();
      }
      await db.prepare('UPDATE users SET fails = 0, locked_until = 0, last_login = ? WHERE username = ?').bind(new Date().toISOString(), username).run();
      return json({ user: pubUser(u) }, 200, { 'set-cookie': await newSession(db, req, username) });
    }
    if (route === 'logout' && method === 'POST') {
      const t = cookieOf(req); if (t) await db.prepare('DELETE FROM sessions WHERE token = ?').bind(t).run();
      return json({ ok: true }, 200, { 'set-cookie': sessionCookie(req, '', 0) });
    }

    const me = await currentUser(db, req);
    if (!me) return fail('انتهت الجلسة، سجّل الدخول من جديد.', 401);
    const isAdmin = me.role === 'admin', canWrite = me.role !== 'viewer';

    if (route === 'password' && method === 'POST') {
      const b = await readBody(req);
      if (!sameHex(await hashPassword(String(b.current || ''), me.salt), me.hash)) return fail('كلمة المرور الحالية غير صحيحة.', 403);
      if (String(b.password || '').length < 8) return fail('كلمة المرور الجديدة 8 أحرف على الأقل.');
      const salt = randHex(16);
      await db.prepare('UPDATE users SET salt = ?, hash = ? WHERE username = ?').bind(salt, await hashPassword(String(b.password), salt), me.username).run();
      await db.prepare('DELETE FROM sessions WHERE username = ? AND token <> ?').bind(me.username, cookieOf(req)).run();
      return json({ ok: true });
    }

    if (route === 'users') {
      if (!isAdmin) return fail('إدارة المستخدمين لمدير البرنامج فقط.', 403);
      if (method === 'GET') {
        const r = await db.prepare('SELECT * FROM users ORDER BY created').all();
        return json({ users: r.results.map(pubUser) });
      }
      if (method === 'POST') {
        const b = await readBody(req), username = cleanUser(b.username);
        if (!validUser(username)) return fail('اسم المستخدم من 3 إلى 32 حرفًا إنجليزيًا أو رقمًا (مسموح . _ -).');
        const role = ROLES.includes(b.role) ? b.role : 'staff';
        const old = await db.prepare('SELECT * FROM users WHERE username = ?').bind(username).first();
        const active = b.active === false ? 0 : 1;
        if (old && old.username === me.username && (role !== 'admin' || !active)) return fail('لا يمكنك سحب صلاحية المدير من حسابك أو إيقافه.');
        const pw = b.password == null ? '' : String(b.password);
        if (!old && pw.length < 8) return fail('كلمة المرور 8 أحرف على الأقل.');
        if (pw && pw.length < 8) return fail('كلمة المرور 8 أحرف على الأقل.');
        const name = String(b.name || '').trim() || username;
        if (old) {
          await db.prepare('UPDATE users SET name = ?, role = ?, active = ? WHERE username = ?').bind(name, role, active, username).run();
          if (pw) { const salt = randHex(16); await db.prepare('UPDATE users SET salt = ?, hash = ?, fails = 0, locked_until = 0 WHERE username = ?').bind(salt, await hashPassword(pw, salt), username).run(); }
          if (pw || !active || role !== old.role) await db.prepare('DELETE FROM sessions WHERE username = ?').bind(username).run();
        } else {
          const salt = randHex(16);
          await db.prepare('INSERT INTO users (username, name, role, salt, hash, active, created) VALUES (?, ?, ?, ?, ?, ?, ?)').bind(username, name, role, salt, await hashPassword(pw, salt), active, new Date().toISOString()).run();
        }
        return json({ ok: true });
      }
      if (method === 'DELETE') {
        const username = cleanUser(url.searchParams.get('u'));
        if (username === me.username) return fail('لا يمكنك حذف حسابك.');
        await db.batch([db.prepare('DELETE FROM sessions WHERE username = ?').bind(username), db.prepare('DELETE FROM users WHERE username = ?').bind(username)]);
        return json({ ok: true });
      }
    }

    /* ---------- data ---------- */
    if (route === 'sync' && method === 'GET') {
      const since = Number(url.searchParams.get('since')) || 0;
      const r = since
        ? await db.prepare('SELECT path, col, data, deleted, updated FROM docs WHERE updated > ? AND col NOT IN (?, ?) ORDER BY updated').bind(since, HEAVY[0], HEAVY[1]).all()
        : await db.prepare('SELECT path, col, data, deleted, updated FROM docs WHERE deleted = 0 AND col NOT IN (?, ?)').bind(HEAVY[0], HEAVY[1]).all();
      const m = await db.prepare('SELECT MAX(updated) AS m FROM docs').first();
      return json({ now: (m && m.m) || 0, docs: r.results.map(x => ({ p: x.path, d: x.deleted ? null : JSON.parse(x.data) })) });
    }
    if (route === 'doc') {
      const path = url.searchParams.get('p') || '';
      if (!PATH_RE.test(path)) return fail('مسار غير صالح.');
      const col = path.split('/')[0];
      if (method === 'GET') {
        const x = await db.prepare('SELECT data FROM docs WHERE path = ? AND deleted = 0').bind(path).first();
        return json({ exists: !!x, data: x ? JSON.parse(x.data) : null });
      }
      if (!canWrite) return fail('ليست لديك صلاحية التعديل.', 403);
      if (!isAdmin && ADMIN_ONLY_WRITE.some(r => r.test(path))) return fail('هذا الإجراء لمدير البرنامج فقط.', 403);
      if (method === 'PUT' || method === 'PATCH') {
        const b = await readBody(req);
        if (!b || typeof b.data !== 'object' || Array.isArray(b.data) || b.data === null) return fail('بيانات غير صالحة.');
        let data = b.data;
        if (method === 'PATCH') {
          const x = await db.prepare('SELECT data FROM docs WHERE path = ? AND deleted = 0').bind(path).first();
          if (!x) return fail('السجل غير موجود.', 404);
          data = Object.assign(JSON.parse(x.data), b.data);
        }
        const text = JSON.stringify(data);
        if (text.length > 1800000) return fail('السجل أكبر من الحد المسموح.', 413);
        const t = await stamp(db);
        await db.prepare('INSERT INTO docs (path, col, data, deleted, updated, by) VALUES (?, ?, ?, 0, ?, ?) ON CONFLICT(path) DO UPDATE SET data = excluded.data, deleted = 0, updated = excluded.updated, by = excluded.by').bind(path, col, text, t, me.username).run();
        return json({ ok: true, updated: t, data });
      }
      if (method === 'DELETE') {
        if (!isAdmin && ADMIN_ONLY_DELETE.some(r => r.test(path))) return fail('الحذف هنا لمدير البرنامج فقط.', 403);
        const t = await stamp(db);
        await db.prepare('UPDATE docs SET data = NULL, deleted = 1, updated = ?, by = ? WHERE path = ?').bind(t, me.username, path).run();
        return json({ ok: true, updated: t });
      }
    }
    return fail('غير موجود.', 404);
  } catch (e) {
    return fail('خطأ في الخادم: ' + (e && e.message ? e.message : String(e)), 500);
  }
}
