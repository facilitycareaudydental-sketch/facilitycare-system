import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination } from '../utils/pagination.js';
import { runSync } from '../utils/calendar.js';

export async function handleRelievers(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'relievers', 'read')) return forbidden(origin);

  const url = new URL(request.url);
  const path = url.pathname.replace('/api/relievers', '');
  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'GET' && path === '') return list(request, env, origin);
  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'relievers', 'write')) return forbidden(origin);
    return importRelievers(request, env, origin);
  }
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'relievers', 'write')) return forbidden(origin);
    return create(request, env, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') return getOne(id, env, origin);
    if (request.method === 'PUT') {
      if (!hasPermission(user, 'relievers', 'write')) return forbidden(origin);
      return update(id, request, env, origin);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'relievers', 'delete')) return forbidden(origin);
      return remove(id, env, origin);
    }
  }
  return error('Not found', 404, origin);
}

async function list(request, env, origin) {
  const { page, limit, offset } = getPagination(request.url);
  const url = new URL(request.url);
  const branch_id = url.searchParams.get('branch_id') || '';
  const period = url.searchParams.get('period') || '';
  const status = url.searchParams.get('status') || '';
  const reliever_name = url.searchParams.get('reliever_name') || '';
  const search = url.searchParams.get('search') || '';
  const month = url.searchParams.get('month') || '';

  let conditions = [];
  let values = [];
  if (branch_id) { conditions.push('r.branch_id = ?'); values.push(branch_id); }
  if (period) { conditions.push('r.period = ?'); values.push(period); }
  if (status) { conditions.push('r.status = ?'); values.push(status); }
  if (reliever_name) { conditions.push('r.reliever_name LIKE ?'); values.push(`%${reliever_name}%`); }
  if (search) { conditions.push('(r.reliever_name LIKE ? OR r.original_fc_name LIKE ? OR r.reason LIKE ?)'); values.push(`%${search}%`, `%${search}%`, `%${search}%`); }
  if (month) { conditions.push("strftime('%Y-%m', r.backup_date) = ?"); values.push(month); }

  const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';

  const [countResult, rows] = await Promise.all([
    env.DB.prepare(`SELECT COUNT(*) as total FROM relievers r ${where}`).bind(...values).first(),
    env.DB.prepare(
      `SELECT r.*, b.full_name as branch_name FROM relievers r
       LEFT JOIN branches b ON r.branch_id = b.id
       ${where} ORDER BY r.period DESC, r.backup_date DESC LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all()
  ]);

  return paginated(rows.results, countResult.total, page, limit, origin);
}

async function getOne(id, env, origin) {
  const row = await env.DB.prepare(
    'SELECT r.*, b.full_name as branch_name FROM relievers r LEFT JOIN branches b ON r.branch_id = b.id WHERE r.id = ?'
  ).bind(id).first();
  if (!row) return notFound(origin);
  return ok(row, 200, origin);
}

async function create(request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const { branch_id, original_fc_name, period, reliever_name, backup_date, completion_date, reason, shift, status } = body;
  if (!reliever_name || !backup_date) return error('reliever_name and backup_date required', 400, origin);

  const result = await env.DB.prepare(
    'INSERT INTO relievers (branch_id, original_fc_name, period, reliever_name, backup_date, completion_date, reason, shift, status) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)'
  ).bind(branch_id || null, original_fc_name || null, period || null, reliever_name,
    backup_date, completion_date || null, reason || null, shift || null, status !== null && status !== undefined && status !== '' ? status : '').run();

  const newId = result.meta.last_row_id;
  try {
    await runSync(env.DB, 'relievers', newId, {
      reliever_name, backup_date, status: status !== null && status !== undefined && status !== '' ? status : '', branch_id, original_fc_name, reason, shift
    });
  } catch (e) { console.error('sync error', e.message); }

  return ok({ id: newId }, 201, origin);
}

async function update(id, request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const existing = await env.DB.prepare('SELECT id FROM relievers WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);

  const { branch_id, original_fc_name, period, reliever_name, backup_date, completion_date, reason, shift, status } = body;
  await env.DB.prepare(
    `UPDATE relievers SET branch_id = COALESCE(?, branch_id), original_fc_name = COALESCE(?, original_fc_name),
     period = COALESCE(?, period), reliever_name = COALESCE(?, reliever_name), backup_date = COALESCE(?, backup_date),
     completion_date = COALESCE(?, completion_date), reason = COALESCE(?, reason),
     shift = COALESCE(?, shift), status = COALESCE(?, status), updated_at = datetime('now') WHERE id = ?`
  ).bind(branch_id || null, original_fc_name || null, period || null, reliever_name || null,
    backup_date || null, completion_date || null, reason || null, shift || null, status !== null && status !== undefined && status !== '' ? status : '', id).run();

  const updated = await env.DB.prepare('SELECT * FROM relievers WHERE id = ?').bind(id).first();
  if (updated) {
    try { await runSync(env.DB, 'relievers', id, updated); } catch (e) { console.error('sync error', e.message); }
  }

  return ok({ message: 'Updated' }, 200, origin);
}

async function remove(id, env, origin) {
  const existing = await env.DB.prepare('SELECT id FROM relievers WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);
  await env.DB.prepare('DELETE FROM relievers WHERE id = ?').bind(id).run();
  try { await runSync(env.DB, 'relievers', id, null); } catch (e) { /* non-fatal */ }
  return ok({ message: 'Deleted' }, 200, origin);
}

async function importRelievers(request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
  if (body.length === 0) return ok({ message: 'No data to import' }, 200, origin);

  let processed = 0;
  try {
    for (const item of body) {
      if (!item.reliever_name || !item.backup_date) continue;
      
      const existing = await env.DB.prepare('SELECT id FROM relievers WHERE (branch_id = ? OR branch_id IS NULL) AND reliever_name = ? AND backup_date = ?').bind(item.branch_id || null, item.reliever_name, item.backup_date).first();

      if (existing) {
        await env.DB.prepare(
          `UPDATE relievers SET original_fc_name = ?, period = ?, completion_date = ?, reason = ?, shift = ?, status = ? WHERE id = ?`
        ).bind(
          item.original_fc_name || null,
          item.period || null,
          item.completion_date || null,
          item.reason || null,
          item.shift || null,
          item.status !== null && item.status !== undefined && item.status !== '' ? item.status : '',
          existing.id
        ).run();
      } else {
        await env.DB.prepare(
          `INSERT INTO relievers (branch_id, original_fc_name, period, reliever_name, backup_date, completion_date, reason, shift, status) 
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(
          item.branch_id || null,
          item.original_fc_name || null,
          item.period || null,
          item.reliever_name,
          item.backup_date,
          item.completion_date || null,
          item.reason || null,
          item.shift || null,
          item.status !== null && item.status !== undefined && item.status !== '' ? item.status : ''
        ).run();
      }
      processed++;
    }
    return ok({ message: `Berhasil memperbarui/mengimport ${processed} data` }, 200, origin);
  } catch (err) {
    return error('Gagal import data: ' + err.message, 500, origin);
  }
}
