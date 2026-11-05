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
  if (request.method === 'DELETE' && path === '/bulk') {
    if (!hasPermission(user, 'users', 'admin')) return forbidden(origin);
    return bulkDeleteUsers(request, env, user, origin);
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
  const cleanUsername = username ? String(username).trim() : null;
  const cleanEmail = email ? String(email).trim().toLowerCase() : null;
  const cleanFullName = full_name ? String(full_name).trim() : null;
  const cleanPassword = password ? String(password).trim() : null;

  if (!cleanUsername || !cleanEmail || !cleanFullName || !cleanPassword) {
    return error('username, email, full_name, password required', 400, origin);
  }
  if (cleanPassword.length < 6) {
    return error('Password minimal harus 6 karakter', 400, origin);
  }

  const existing = await env.DB.prepare(
    'SELECT id, username, email FROM users WHERE LOWER(username) = LOWER(?) OR LOWER(email) = LOWER(?)'
  ).bind(cleanUsername, cleanEmail).first();
  if (existing) {
    if (existing.email && existing.email.toLowerCase() === cleanEmail) {
      return error(`Email '${cleanEmail}' sudah digunakan oleh user lain.`, 409, origin);
    }
    return error(`Username '${cleanUsername}' sudah digunakan oleh user lain.`, 409, origin);
  }

  const hash = await hashPassword(cleanPassword);
  try {
    const result = await env.DB.prepare(
      'INSERT INTO users (username, email, password_hash, full_name, role) VALUES (?, ?, ?, ?, ?)'
    ).bind(cleanUsername, cleanEmail, hash, cleanFullName, role || 'viewer').run();

    return ok({ id: result.meta.last_row_id, message: 'User created' }, 201, origin);
  } catch (err) {
    const msg = err?.message || String(err);
    if (msg.includes('UNIQUE constraint failed: users.email')) {
      return error('Email sudah digunakan oleh user lain.', 409, origin);
    }
    if (msg.includes('UNIQUE constraint failed: users.username')) {
      return error('Username sudah digunakan oleh user lain.', 409, origin);
    }
    return error('Gagal membuat user: ' + msg, 500, origin);
  }
}

async function updateUser(id, request, env, currentUser, origin) {
  if (!hasPermission(currentUser, 'users', 'admin') && currentUser.id != id) return forbidden(origin);
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }

  const existing = await env.DB.prepare(
    'SELECT id, username, email, full_name, role, is_active FROM users WHERE id = ?'
  ).bind(id).first();
  if (!existing) return notFound(origin);

  const { username, email, full_name, role, is_active, password } = body;
  let updates = [];
  let values = [];

  const cleanEmail = (email !== undefined && email !== null) ? String(email).trim().toLowerCase() : null;
  const cleanUsername = (username !== undefined && username !== null) ? String(username).trim() : null;
  const cleanFullName = (full_name !== undefined && full_name !== null) ? String(full_name).trim() : null;

  // 1. Email: Only check uniqueness and update if email actually changed!
  if (cleanEmail && cleanEmail !== (existing.email || '').toLowerCase()) {
    const emailConflict = await env.DB.prepare(
      'SELECT id, username, full_name FROM users WHERE LOWER(email) = LOWER(?) AND id != ?'
    ).bind(cleanEmail, id).first();
    if (emailConflict) {
      return error(`Email '${cleanEmail}' sudah digunakan oleh user lain (${emailConflict.username} - ${emailConflict.full_name}). Gunakan email yang berbeda.`, 409, origin);
    }
    updates.push('email = ?');
    values.push(cleanEmail);
  }

  // 2. Username: Only check uniqueness and update if username actually changed!
  if (cleanUsername && cleanUsername !== existing.username) {
    if (!hasPermission(currentUser, 'users', 'admin')) {
      return forbidden(origin);
    }
    const usernameConflict = await env.DB.prepare(
      'SELECT id, username, full_name FROM users WHERE LOWER(username) = LOWER(?) AND id != ?'
    ).bind(cleanUsername, id).first();
    if (usernameConflict) {
      return error(`Username '${cleanUsername}' sudah digunakan oleh user lain (${usernameConflict.full_name}). Gunakan username yang berbeda.`, 409, origin);
    }
    updates.push('username = ?');
    values.push(cleanUsername);
  }

  // 3. Full Name: Only update if changed
  if (cleanFullName && cleanFullName !== existing.full_name) {
    updates.push('full_name = ?');
    values.push(cleanFullName);
  }

  // 4. Role: Only update if changed and user is admin
  if (role && hasPermission(currentUser, 'users', 'admin') && role !== existing.role) {
    updates.push('role = ?');
    values.push(role);
  }

  // 5. Active Status: Only update if changed and user is admin
  if (is_active !== undefined && is_active !== null && hasPermission(currentUser, 'users', 'admin')) {
    const targetActive = (is_active === 1 || is_active === true || is_active === '1') ? 1 : 0;
    if (currentUser.id == id && targetActive === 0) {
      return error('Tidak dapat menonaktifkan akun sendiri', 400, origin);
    }
    if (targetActive !== existing.is_active) {
      updates.push('is_active = ?');
      values.push(targetActive);
    }
  }

  // 6. Password update: Only if provided, non-empty, and at least 6 characters
  if (password && typeof password === 'string') {
    const trimmedPw = password.trim();
    if (trimmedPw.length > 0) {
      if (trimmedPw.length < 6) {
        return error('Password minimal harus 6 karakter', 400, origin);
      }
      const hash = await hashPassword(trimmedPw);
      updates.push('password_hash = ?');
      values.push(hash);
    }
  }

  // If nothing changed, return success directly without unnecessary DB write
  if (updates.length === 0) {
    return ok({ message: 'Tidak ada perubahan data' }, 200, origin);
  }

  updates.push("updated_at = datetime('now')");
  values.push(id);

  try {
    await env.DB.prepare(`UPDATE users SET ${updates.join(', ')} WHERE id = ?`).bind(...values).run();
    return ok({ message: 'User updated' }, 200, origin);
  } catch (err) {
    const msg = err?.message || String(err);
    if (msg.includes('UNIQUE constraint failed: users.email')) {
      return error('Email sudah digunakan oleh user lain. Gunakan email yang berbeda.', 409, origin);
    }
    if (msg.includes('UNIQUE constraint failed: users.username')) {
      return error('Username sudah digunakan oleh user lain. Gunakan username yang berbeda.', 409, origin);
    }
    return error(`Gagal memperbarui user: ${msg}`, 500, origin);
  }
}

async function deleteUser(id, env, currentUser, origin) {
  if (!hasPermission(currentUser, 'users', 'admin')) return forbidden(origin);
  if (currentUser.id == id) return error('Cannot delete yourself', 400, origin);
  const existing = await env.DB.prepare('SELECT id FROM users WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);
  await env.DB.prepare('DELETE FROM users WHERE id = ?').bind(id).run();
  return ok({ message: 'User deleted' }, 200, origin);
}

async function bulkDeleteUsers(request, env, currentUser, origin) {
  if (!hasPermission(currentUser, 'users', 'admin')) return forbidden(origin);
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const { ids } = body;
  if (!Array.isArray(ids) || ids.length === 0) return error('ids array required', 400, origin);

  // Validate all are numbers and exclude current logged-in user
  const validIds = ids
    .filter(id => Number.isInteger(Number(id)))
    .map(Number)
    .filter(id => id !== Number(currentUser.id));

  if (validIds.length === 0) return error('Tidak ada user valid untuk dihapus (akun sendiri tidak dapat dihapus)', 400, origin);

  const placeholders = validIds.map(() => '?').join(',');
  try {
    const result = await env.DB.prepare(
      `DELETE FROM users WHERE id IN (${placeholders})`
    ).bind(...validIds).run();
    return ok({ message: `Berhasil menghapus ${result.meta.changes} user`, deleted: result.meta.changes }, 200, origin);
  } catch (err) {
    return error('Gagal hapus user: ' + err.message, 500, origin);
  }
}

async function importUsers(request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
  if (body.length === 0) return ok({ message: 'No data to import' }, 200, origin);

  const stmts = [];
  for (const item of body) {
    const username = item.username ? String(item.username).trim() : null;
    const email = item.email ? String(item.email).trim().toLowerCase() : null;
    const full_name = item.full_name ? String(item.full_name).trim() : null;
    const password = item.password ? String(item.password).trim() : null;
    if (!username || !email || !password || !full_name) continue;
    
    // Hash password sequentially
    const hash = await hashPassword(password);

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
        username,
        email,
        hash,
        full_name,
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
