import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination } from '../utils/pagination.js';
import { runSync } from '../utils/calendar.js';

export async function handleSchedule(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'schedule', 'read')) return forbidden(origin);

  const url = new URL(request.url);
  const path = url.pathname.replace('/api/schedule', '');
  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'GET' && path === '') return listSchedule(request, env, origin);
  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'schedule', 'write')) return forbidden(origin);
    return importSchedule(request, env, origin);
  }
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'schedule', 'write')) return forbidden(origin);
    return createSchedule(request, env, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') return getSchedule(id, env, origin);
    if (request.method === 'PUT') {
      if (!hasPermission(user, 'schedule', 'write')) return forbidden(origin);
      return updateSchedule(id, request, env, origin);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'schedule', 'delete')) return forbidden(origin);
      return deleteSchedule(id, env, origin);
    }
  }
  return error('Not found', 404, origin);
}

async function listSchedule(request, env, origin) {
  const { page, limit, offset } = getPagination(request.url);
  const url = new URL(request.url);
  const branch_id = url.searchParams.get('branch_id') || '';
  const activity_type = url.searchParams.get('activity_type') || '';
  const period = url.searchParams.get('period') || '';
  const status = url.searchParams.get('status') || '';
  const pic = url.searchParams.get('pic') || '';

  let conditions = [];
  let values = [];
  if (branch_id) { conditions.push('s.branch_id = ?'); values.push(branch_id); }
  if (activity_type) { conditions.push('s.activity_type = ?'); values.push(activity_type); }
  if (period) { conditions.push('s.period = ?'); values.push(period); }
  if (status) { conditions.push('s.status = ?'); values.push(status); }
  if (pic) { conditions.push('s.pic = ?'); values.push(pic); }

  const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';

  const [countResult, rows] = await Promise.all([
    env.DB.prepare(`SELECT COUNT(*) as total FROM activity_schedule s ${where}`).bind(...values).first(),
    env.DB.prepare(
      `SELECT s.*, b.full_name as branch_name FROM activity_schedule s
       LEFT JOIN branches b ON s.branch_id = b.id
       ${where} ORDER BY s.target_date ASC, s.id DESC LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all()
  ]);

  return paginated(rows.results, countResult.total, page, limit, origin);
}

async function getSchedule(id, env, origin) {
  const row = await env.DB.prepare(
    'SELECT s.*, b.full_name as branch_name FROM activity_schedule s LEFT JOIN branches b ON s.branch_id = b.id WHERE s.id = ?'
  ).bind(id).first();
  if (!row) return notFound(origin);
  return ok(row, 200, origin);
}

async function createSchedule(request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const { branch_id, activity_type, period, pic, opening_date, target_date, completion_date, status, notes } = body;
  if (!activity_type || !period) return error('activity_type and period required', 400, origin);

  const result = await env.DB.prepare(
    'INSERT INTO activity_schedule (branch_id, activity_type, period, pic, opening_date, target_date, completion_date, status, notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)'
  ).bind(branch_id || null, activity_type, period, pic || null, opening_date || null,
    target_date || null, completion_date || null, status !== null && status !== undefined && status !== '' ? status : '', notes || null).run();

  const newId = result.meta.last_row_id;
  try {
    await runSync(env.DB, 'schedule', newId, {
      activity_type, period, pic, target_date, status: status !== null && status !== undefined && status !== '' ? status : '', branch_id, notes
    });
  } catch (e) { console.error('sync error', e.message); }

  return ok({ id: newId }, 201, origin);
}

async function updateSchedule(id, request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const existing = await env.DB.prepare('SELECT id FROM activity_schedule WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);

  const { branch_id, activity_type, period, pic, opening_date, target_date, completion_date, status, notes } = body;
  await env.DB.prepare(
    `UPDATE activity_schedule SET
      branch_id = COALESCE(?, branch_id), activity_type = COALESCE(?, activity_type),
      period = COALESCE(?, period), pic = COALESCE(?, pic),
      opening_date = COALESCE(?, opening_date), target_date = COALESCE(?, target_date),
      completion_date = COALESCE(?, completion_date), status = COALESCE(?, status),
      notes = COALESCE(?, notes), updated_at = datetime('now')
     WHERE id = ?`
  ).bind(branch_id || null, activity_type || null, period || null, pic || null,
    opening_date || null, target_date || null, completion_date || null,
    status !== null && status !== undefined && status !== '' ? status : '', notes || null, id).run();

  const updated = await env.DB.prepare('SELECT * FROM activity_schedule WHERE id = ?').bind(id).first();
  if (updated) {
    try { await runSync(env.DB, 'schedule', id, updated); } catch (e) { console.error('sync error', e.message); }
  }

  return ok({ message: 'Schedule updated' }, 200, origin);
}

async function deleteSchedule(id, env, origin) {
  const existing = await env.DB.prepare('SELECT id FROM activity_schedule WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);
  await env.DB.prepare('DELETE FROM activity_schedule WHERE id = ?').bind(id).run();
  try { await runSync(env.DB, 'schedule', id, null); } catch (e) { /* non-fatal */ }
  return ok({ message: 'Schedule deleted' }, 200, origin);
}

async function importSchedule(request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
  if (body.length === 0) return ok({ message: 'No data to import' }, 200, origin);

  const existing = await env.DB.prepare('SELECT id, activity_type, period, branch_id FROM activity_schedule').all();
  const existingMap = new Map();
  (existing.results || []).forEach(s => {
    if (s.activity_type && s.period) {
      existingMap.set(s.activity_type.toLowerCase().trim() + '_' + s.period.toLowerCase().trim() + '_' + s.branch_id, s.id);
    }
  });

  const stmts = [];
  let inserted = 0;
  let updated = 0;

  for (const item of body) {
    if (!item.activity_type || !item.period) continue;
    
    const key = item.activity_type.toLowerCase().trim() + '_' + item.period.toLowerCase().trim() + '_' + (item.branch_id || 'null');
    const status = item.status !== null && item.status !== undefined && item.status !== '' ? item.status : '';

    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      stmts.push(
        env.DB.prepare(
          `UPDATE activity_schedule SET pic = ?, opening_date = ?, target_date = ?, completion_date = ?, status = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(
          item.pic || null,
          item.opening_date || null,
          item.target_date || null,
          item.completion_date || null,
          status,
          item.notes || null,
          id
        )
      );
      updated++;
    } else {
      stmts.push(
        env.DB.prepare(
          `INSERT INTO activity_schedule (branch_id, activity_type, period, pic, opening_date, target_date, completion_date, status, notes) 
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(
          item.branch_id || null,
          item.activity_type,
          item.period,
          item.pic || null,
          item.opening_date || null,
          item.target_date || null,
          item.completion_date || null,
          status,
          item.notes || null
        )
      );
      inserted++;
    }
  }

  try {
    if (stmts.length > 0) await env.DB.batch(stmts);
    return ok({ message: `Berhasil mengimport data (Baru: ${inserted}, Update: ${updated})` }, 200, origin);
  } catch (err) {
    return error('Gagal import data: ' + err.message, 500, origin);
  }
}
