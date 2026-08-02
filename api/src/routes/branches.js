import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination, getSearchParam } from '../utils/pagination.js';
import { buildOutboxQuery } from '../utils/sync_engine.js';
import { mapDBToPayload } from '../utils/sync_mapper.js';

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
    const rows = await env.DB.prepare('SELECT * FROM branches WHERE is_active = 1 AND deleted_at IS NULL ORDER BY code').all();
    return ok(rows.results, 200, origin);
  }

  let query, countQuery, rows, count;
  if (search) {
    const s = `%${search}%`;
    query = 'SELECT * FROM branches WHERE (name LIKE ? OR code LIKE ? OR full_name LIKE ?) AND deleted_at IS NULL ORDER BY code LIMIT ? OFFSET ?';
    countQuery = 'SELECT COUNT(*) as total FROM branches WHERE (name LIKE ? OR code LIKE ? OR full_name LIKE ?) AND deleted_at IS NULL';
    [rows, count] = await Promise.all([
      env.DB.prepare(query).bind(s, s, s, limit, offset).all(),
      env.DB.prepare(countQuery).bind(s, s, s).first()
    ]);
  } else {
    [rows, count] = await Promise.all([
      env.DB.prepare('SELECT * FROM branches WHERE deleted_at IS NULL ORDER BY code LIMIT ? OFFSET ?').bind(limit, offset).all(),
      env.DB.prepare('SELECT COUNT(*) as total FROM branches WHERE deleted_at IS NULL').first()
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
    'INSERT INTO branches (code, name, full_name, city, row_version, last_sync_source) VALUES (?, ?, ?, ?, 1, "FCMS")'
  ).bind(code, name, full_name, city || null).run();
  
  const newId = result.meta.last_row_id;
  const newRecord = await env.DB.prepare('SELECT * FROM branches WHERE id = ?').bind(newId).first();
  if (newRecord) {
    const payload = mapDBToPayload('Data Cabang', newRecord);
    await buildOutboxQuery(env, 'Data Cabang', newId, 'INSERT', payload).run();
  }

  return ok({ id: newId }, 201, origin);
}

async function updateBranch(id, request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const existing = await env.DB.prepare('SELECT id FROM branches WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);
  const { name, full_name, city, is_active } = body;
  
  const updatedRecord = { ...existing, name: name || existing.name, full_name: full_name || existing.full_name, city: city || existing.city, is_active: is_active !== undefined ? (is_active ? 1 : 0) : existing.is_active };
  const payload = mapDBToPayload('Data Cabang', updatedRecord);

  await env.DB.batch([
    env.DB.prepare(
      "UPDATE branches SET name = COALESCE(?, name), full_name = COALESCE(?, full_name), city = COALESCE(?, city), is_active = COALESCE(?, is_active), updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ?"
    ).bind(name || null, full_name || null, city || null, is_active !== undefined ? (is_active ? 1 : 0) : null, id),
    buildOutboxQuery(env, 'Data Cabang', id, 'UPDATE', payload)
  ]);
  
  return ok({ message: 'Branch updated' }, 200, origin);
}

async function deleteBranch(id, env, origin) {
  const existing = await env.DB.prepare('SELECT * FROM branches WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);
  
  const deletedRecord = { ...existing, is_active: 0, deleted_at: new Date().toISOString() };
  const payload = mapDBToPayload('Data Cabang', deletedRecord);

  await env.DB.batch([
    env.DB.prepare("UPDATE branches SET is_active = 0, deleted_at = datetime('now'), updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ?").bind(id),
    buildOutboxQuery(env, 'Data Cabang', id, 'DELETE', payload)
  ]);
  
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
