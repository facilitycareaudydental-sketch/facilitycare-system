import { authenticate, hashPassword, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination, getSearchParam } from '../utils/pagination.js';

export async function handleUsers(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'users', 'read')) return forbidden(origin);

  const url = new URL(request.url);
  const path = url.pathname.replace('/api/users', '');
  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'GET' && path === '') return listUsers(request, env, origin);
  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'users', 'admin')) return forbidden(origin);
    return importUsers(request, env, origin);
  }
  if (request.method === 'POST' && path === '') return createUser(request, env, user, origin);
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') return getUser(id, env, origin);
    if (request.method === 'PUT') return updateUser(id, request, env, user, origin);
    if (request.method === 'DELETE') return deleteUser(id, env, user, origin);
  }
  return error('Not found', 404, origin);
}

async function listUsers(request, env, origin) {
  const { page, limit, offset } = getPagination(request.url);
  const search = getSearchParam(request.url, 'search');

  let query, countQuery;
  if (search) {
    query = `SELECT id, username, email, full_name, role, is_active, created_at FROM users 
             WHERE full_name LIKE ? OR username LIKE ? OR email LIKE ? 
             ORDER BY full_name LIMIT ? OFFSET ?`;
    countQuery = `SELECT COUNT(*) as total FROM users WHERE full_name LIKE ? OR username LIKE ? OR email LIKE ?`;
    const s = `%${search}%`;
    const [rows, count] = await Promise.all([
      env.DB.prepare(query).bind(s, s, s, limit, offset).all(),
      env.DB.prepare(countQuery).bind(s, s, s).first()
    ]);
    return paginated(rows.results, count.total, page, limit, origin);
  }

  query = `SELECT id, username, email, full_name, role, is_active, created_at FROM users ORDER BY full_name LIMIT ? OFFSET ?`;
  countQuery = `SELECT COUNT(*) as total FROM users`;
  const [rows, count] = await Promise.all([
    env.DB.prepare(query).bind(limit, offset).all(),
    env.DB.prepare(countQuery).first()
  ]);
  return paginated(rows.results, count.total, page, limit, origin);
}

async function getUser(id, env, origin) {
  const row = await env.DB.prepare(
    'SELECT id, username, email, full_name, role, is_active, created_at, updated_at FROM users WHERE id = ?'
  ).bind(id).first();
  if (!row) return notFound(origin);
  return ok(row, 200, origin);
}

async function createUser(request, env, currentUser, origin) {
  if (!hasPermission(currentUser, 'users', 'admin')) return forbidden(origin);
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }

  const { username, email, full_name, password, role } = body;
  if (!username || !email || !full_name || !password) return error('username, email, full_name, password required', 400, origin);

  const existing = await env.DB.prepare('SELECT id FROM users WHERE username = ? OR email = ?').bind(username, email).first();
  if (existing) return error('Username or email already exists', 409, origin);

  const hash = await hashPassword(password);
  const result = await env.DB.prepare(
    'INSERT INTO users (username, email, password_hash, full_name, role) VALUES (?, ?, ?, ?, ?)'
  ).bind(username, email, hash, full_name, role || 'viewer').run();

  return ok({ id: result.meta.last_row_id, message: 'User created' }, 201, origin);
}

async function updateUser(id, request, env, currentUser, origin) {
  if (!hasPermission(currentUser, 'users', 'admin') && currentUser.id != id) return forbidden(origin);
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }

  const existing = await env.DB.prepare('SELECT id FROM users WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);

  const { email, full_name, role, is_active, password } = body;
  let updates = [];
  let values = [];

  if (email) { updates.push('email = ?'); values.push(email); }
  if (full_name) { updates.push('full_name = ?'); values.push(full_name); }
  if (role && hasPermission(currentUser, 'users', 'admin')) { updates.push('role = ?'); values.push(role); }
  if (is_active !== undefined && hasPermission(currentUser, 'users', 'admin')) {
    updates.push('is_active = ?'); values.push(is_active ? 1 : 0);
  }
  if (password && password.length >= 6) {
    const hash = await hashPassword(password);
    updates.push('password_hash = ?'); values.push(hash);
  }
  updates.push("updated_at = datetime('now')");
  values.push(id);

  await env.DB.prepare(`UPDATE users SET ${updates.join(', ')} WHERE id = ?`).bind(...values).run();
  return ok({ message: 'User updated' }, 200, origin);
}

async function deleteUser(id, env, currentUser, origin) {
  if (!hasPermission(currentUser, 'users', 'admin')) return forbidden(origin);
  if (currentUser.id == id) return error('Cannot delete yourself', 400, origin);
  const existing = await env.DB.prepare('SELECT id FROM users WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);
  await env.DB.prepare('DELETE FROM users WHERE id = ?').bind(id).run();
  return ok({ message: 'User deleted' }, 200, origin);
}

async function importUsers(request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
  if (body.length === 0) return ok({ message: 'No data to import' }, 200, origin);

  const stmts = [];
  for (const item of body) {
    if (!item.username || !item.email || !item.password || !item.full_name) continue;
    
    // Hash password sequentially
    const hash = await hashPassword(item.password);

    stmts.push(
      env.DB.prepare(
        `INSERT INTO users (username, email, password_hash, full_name, role) 
         VALUES (?, ?, ?, ?, ?)
         ON CONFLICT(username) DO UPDATE SET
           email = excluded.email,
           full_name = excluded.full_name,
           role = excluded.role,
           password_hash = excluded.password_hash`
      ).bind(
        item.username,
        item.email,
        hash,
        item.full_name,
        item.role || 'viewer'
      )
    );
  }

  try {
    if (stmts.length > 0) await env.DB.batch(stmts);
    return ok({ message: `Berhasil mengimport ${stmts.length} user` }, 200, origin);
  } catch (err) {
    return error('Gagal import data: ' + err.message, 500, origin);
  }
}
