import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination, getSearchParam } from '../utils/pagination.js';

export async function handleBranches(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace('/api/branches', '');
  const idMatch = path.match(/^\/(\d+)$/);

  // GET all branches is public (needed for forms)
  if (request.method === 'GET' && path === '') return listBranches(request, env, origin);

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);

  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') return getBranch(id, env, origin);
    if (request.method === 'PUT') {
      if (!hasPermission(user, 'branches', 'write')) return forbidden(origin);
      return updateBranch(id, request, env, origin);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'branches', 'delete')) return forbidden(origin);
      return deleteBranch(id, env, origin);
    }
  }
  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'branches', 'write')) return forbidden(origin);
    return importBranches(request, env, origin);
  }
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'branches', 'write')) return forbidden(origin);
    return createBranch(request, env, origin);
  }
  return error('Not found', 404, origin);
}

async function listBranches(request, env, origin) {
  const { page, limit, offset } = getPagination(request.url);
  const search = getSearchParam(request.url, 'search');
  const all = new URL(request.url).searchParams.get('all');

  if (all === '1') {
    const rows = await env.DB.prepare('SELECT * FROM branches WHERE is_active = 1 ORDER BY code').all();
    return ok(rows.results, 200, origin);
  }

  let query, countQuery, rows, count;
  if (search) {
    const s = `%${search}%`;
    query = 'SELECT * FROM branches WHERE (name LIKE ? OR code LIKE ? OR full_name LIKE ?) ORDER BY code LIMIT ? OFFSET ?';
    countQuery = 'SELECT COUNT(*) as total FROM branches WHERE name LIKE ? OR code LIKE ? OR full_name LIKE ?';
    [rows, count] = await Promise.all([
      env.DB.prepare(query).bind(s, s, s, limit, offset).all(),
      env.DB.prepare(countQuery).bind(s, s, s).first()
    ]);
  } else {
    [rows, count] = await Promise.all([
      env.DB.prepare('SELECT * FROM branches ORDER BY code LIMIT ? OFFSET ?').bind(limit, offset).all(),
      env.DB.prepare('SELECT COUNT(*) as total FROM branches').first()
    ]);
  }
  return paginated(rows.results, count.total, page, limit, origin);
}

async function getBranch(id, env, origin) {
  const row = await env.DB.prepare('SELECT * FROM branches WHERE id = ?').bind(id).first();
  if (!row) return notFound(origin);
  return ok(row, 200, origin);
}

async function createBranch(request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const { city, full_name } = body;
  let { code, name } = body;
  if (!full_name) return error('full_name required', 400, origin);
  if (!code) code = `AUTO-${Math.floor(Math.random() * 100000)}`;
  if (!name) name = full_name;
  
  const result = await env.DB.prepare(
    'INSERT INTO branches (code, name, full_name, city) VALUES (?, ?, ?, ?)'
  ).bind(code, name, full_name, city || null).run();
  return ok({ id: result.meta.last_row_id }, 201, origin);
}

async function updateBranch(id, request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const existing = await env.DB.prepare('SELECT id FROM branches WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);
  const { name, full_name, city, is_active } = body;
  await env.DB.prepare(
    "UPDATE branches SET name = COALESCE(?, name), full_name = COALESCE(?, full_name), city = COALESCE(?, city), is_active = COALESCE(?, is_active), updated_at = datetime('now') WHERE id = ?"
  ).bind(name || null, full_name || null, city || null, is_active !== undefined ? (is_active ? 1 : 0) : null, id).run();
  return ok({ message: 'Branch updated' }, 200, origin);
}

async function deleteBranch(id, env, origin) {
  const existing = await env.DB.prepare('SELECT id FROM branches WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);
  await env.DB.prepare("UPDATE branches SET is_active = 0, updated_at = datetime('now') WHERE id = ?").bind(id).run();
  return ok({ message: 'Branch deactivated' }, 200, origin);
}

async function importBranches(request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
  if (body.length === 0) return ok({ message: 'No data to import' }, 200, origin);

  const stmts = [];
  for (const item of body) {
    if (!item.code || !item.name) continue;
    stmts.push(
      env.DB.prepare(
        `INSERT INTO branches (code, name, full_name, city, is_active) 
         VALUES (?, ?, ?, ?, 1)
         ON CONFLICT(code) DO UPDATE SET 
           name = excluded.name, 
           full_name = excluded.full_name, 
           city = excluded.city,
           is_active = 1,
           updated_at = datetime('now')`
      ).bind(item.code, item.name, item.full_name || item.name, item.city || null)
    );
  }

  try {
    if (stmts.length > 0) await env.DB.batch(stmts);
    return ok({ message: `Berhasil mengimport ${stmts.length} cabang` }, 200, origin);
  } catch (err) {
    return error('Gagal import data: ' + err.message, 500, origin);
  }
}
