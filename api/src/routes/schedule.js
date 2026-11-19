import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination, resolveSortOrder } from '../utils/pagination.js';
import { requireExportAccess, getExportPagination } from '../utils/export.js';
import { runSync } from '../utils/calendar.js';
import { buildOutboxQuery, buildConditionalOutboxQuery, buildOutboxFromLastInsertQuery } from '../utils/sync_engine.js';
import { mapDBToPayload } from '../utils/sync_mapper.js';
import { buildAuditStatement, buildCreateAuditStatement, buildCreateAuditFromOutboxStatement, executeAtomicImportBatch } from '../utils/audit.js';

const SCHEDULE_SORT_MAP = Object.freeze({
  target_date_asc: 's.target_date ASC, s.id DESC',
  opening_date_desc: "CASE WHEN s.opening_date IS NULL OR s.opening_date = '' OR s.opening_date = '-' THEN 1 ELSE 0 END, substr(s.opening_date, 7, 4) DESC, substr(s.opening_date, 4, 2) DESC, substr(s.opening_date, 1, 2) DESC, s.opening_date DESC, s.id DESC",
  date_desc: 's.target_date DESC, s.id DESC',
});

export async function handleSchedule(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace('/api/schedule', '');

  if (request.method === 'GET' && path === '/export') {
    const { response } = await requireExportAccess(request, env, origin, 'schedule');
    if (response) return response;
    return listSchedule(request, env, origin, { isExport: true });
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'schedule', 'read')) return forbidden(origin);

  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'GET' && path === '') return listSchedule(request, env, origin);
  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'schedule', 'import')) return forbidden(origin);
    return importSchedule(request, env, origin, user);
  }
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'schedule', 'write')) return forbidden(origin);
    return createSchedule(request, env, origin, user);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') return getSchedule(id, env, origin);
    if (request.method === 'PUT' || request.method === 'PATCH') {
      if (!hasPermission(user, 'schedule', 'write')) return forbidden(origin);
      return updateSchedule(id, request, env, origin, user);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'schedule', 'delete')) return forbidden(origin);
      return deleteSchedule(id, env, origin, user);
    }
  }
  return error('Not found', 404, origin);
}

async function listSchedule(request, env, origin, { isExport = false } = {}) {
  const { page, limit, offset } = isExport ? getExportPagination(request.url) : getPagination(request.url);
  const url = new URL(request.url);
  const branch_id = (url.searchParams.get('branch_id') || '').trim();
  const activity_type = (url.searchParams.get('activity_type') || '').trim();
  const period = (url.searchParams.get('period') || '').trim();
  const status = (url.searchParams.get('status') || '').trim();
  const pic = (url.searchParams.get('pic') || '').trim();
  const month = (url.searchParams.get('month') || '').trim();
  const search = (url.searchParams.get('search') || '').trim();
  const sort = url.searchParams.get('sort') || 'target_date_asc';

  const conditions = ['s.deleted_at IS NULL'];
  const values = [];
  if (branch_id) { conditions.push('s.branch_id = ?'); values.push(branch_id); }
  if (activity_type) {
    if (activity_type === 'GCDC') {
      conditions.push("s.activity_type IN ('General Cleaning', 'Deep Cleaning')");
    } else {
      conditions.push('s.activity_type = ?');
      values.push(activity_type);
    }
  }
  if (status) { conditions.push('s.status = ?'); values.push(status); }
  if (pic) { conditions.push('s.pic = ?'); values.push(pic); }
  if (search) {
    conditions.push('(b.full_name LIKE ? OR s.activity_type LIKE ? OR s.pic LIKE ? OR s.notes LIKE ?)');
    values.push(`%${search}%`, `%${search}%`, `%${search}%`, `%${search}%`);
  }

  if (month) {
    conditions.push("strftime('%Y-%m', COALESCE(s.completion_date, s.target_date)) = ?");
    values.push(month);
    // Ignore period filter if month is explicitly requested to prevent conflicts
  } else if (period) {
    conditions.push('s.period = ?');
    values.push(period);
  }

  const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';
  const countJoin = search ? 'LEFT JOIN branches b ON s.branch_id = b.id' : '';
  const orderBy = resolveSortOrder(sort, SCHEDULE_SORT_MAP, 'target_date_asc');

  const [countResult, rows] = await Promise.all([
    env.DB.prepare(`SELECT COUNT(*) as total FROM activity_schedule s ${countJoin} ${where}`).bind(...values).first(),
    env.DB.prepare(
      `SELECT s.*, b.full_name as branch_name FROM activity_schedule s
       LEFT JOIN branches b ON s.branch_id = b.id
       ${where} ORDER BY ${orderBy} LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all()
  ]);

  return paginated(rows.results || [], countResult?.total || 0, page, limit, origin, env?.ENVIRONMENT);
}

async function getSchedule(id, env, origin) {
  const row = await env.DB.prepare(
    'SELECT s.*, b.full_name as branch_name FROM activity_schedule s LEFT JOIN branches b ON s.branch_id = b.id WHERE s.id = ? AND s.deleted_at IS NULL'
  ).bind(id).first();
  if (!row) return notFound(origin);
  return ok(row, 200, origin);
}

async function validateScheduleBranch(env, branch_id) {
  if (branch_id === undefined || branch_id === null || branch_id === '') return { valid: true, branchId: null };
  const parsed = Number(branch_id);
  if (!Number.isInteger(parsed) || parsed <= 0) return { valid: false, error: 'branch_id tidak valid' };
  const row = await env.DB.prepare('SELECT id FROM branches WHERE id = ? AND deleted_at IS NULL').bind(parsed).first();
  if (!row) return { valid: false, error: 'Cabang (branch_id) tidak ditemukan atau sudah dinonaktifkan' };
  return { valid: true, branchId: parsed };
}

async function createSchedule(request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
  const { branch_id, activity_type, period, pic, opening_date, target_date, completion_date, status, notes } = body;
  if (!activity_type || !period) return error('activity_type and period required', 400, origin);

  const branchCheck = await validateScheduleBranch(env, branch_id);
  if (!branchCheck.valid) return error(branchCheck.error, 400, origin);

  const normalizedStatus = status !== null && status !== undefined && status !== '' ? status : '';
  const newSnapshot = {
    branch_id: branchCheck.branchId,
    activity_type,
    period,
    pic: pic || null,
    opening_date: opening_date || null,
    target_date: target_date || null,
    completion_date: completion_date || null,
    status: normalizedStatus,
    notes: notes || null,
    last_sync_source: 'FCMS',
    row_version: 1
  };
  const payload = mapDBToPayload('Time Line', {
    id: '',
    ...newSnapshot
  });
  const corrId = crypto.randomUUID();

  const batchRes = await env.DB.batch([
    env.DB.prepare(
      "INSERT INTO activity_schedule (branch_id, activity_type, period, pic, opening_date, target_date, completion_date, status, notes, row_version, last_sync_source) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1, 'FCMS')"
    ).bind(branchCheck.branchId, activity_type, period, pic || null, opening_date || null,
      target_date || null, completion_date || null, normalizedStatus, notes || null),
    buildOutboxFromLastInsertQuery(env, 'Time Line', payload, corrId),
    buildCreateAuditFromOutboxStatement(env, user, 'schedule', corrId, newSnapshot)
  ]);

  const newId = batchRes?.[0]?.meta?.last_row_id ?? 1;
  try {
    await runSync(env.DB, 'schedule', newId, {
      activity_type, period, pic, target_date, status: normalizedStatus, branch_id: branchCheck.branchId, notes
    });
  } catch (e) { console.error('sync error', e.message); }

  return ok({ id: newId }, 201, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function updateSchedule(id, request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
  // OCC: Strict row_version enforcement
  if (body.row_version === undefined || body.row_version === null) {
    return error('row_version is required for optimistic concurrency control', 400, origin);
  }
  const clientVersion = Number(body.row_version);
  if (!Number.isInteger(clientVersion) || clientVersion < 0) {
    return error('Invalid row_version format: must be a non-negative integer', 400, origin);
  }

  const existing = await env.DB.prepare('SELECT * FROM activity_schedule WHERE id = ? AND deleted_at IS NULL').bind(id).first();
  if (!existing) return notFound(origin);

  const isExistingNull = existing.row_version === null || existing.row_version === undefined;
  if (isExistingNull) {
    if (clientVersion !== 1 && clientVersion !== 0) {
      return error('Conflict: Record has been modified by another user or session', 409, origin);
    }
  } else {
    if (existing.row_version !== clientVersion) {
      return error('Conflict: Record has been modified by another user or session', 409, origin);
    }
  }

  if (body.branch_id !== undefined && body.branch_id !== null && body.branch_id !== '') {
    const branchCheck = await validateScheduleBranch(env, body.branch_id);
    if (!branchCheck.valid) return error(branchCheck.error, 400, origin);
  }
  const branch_id = body.branch_id !== undefined ? (body.branch_id ? Number(body.branch_id) : null) : existing.branch_id;
  const activity_type = body.activity_type !== undefined ? body.activity_type : existing.activity_type;
  const period = body.period !== undefined ? body.period : existing.period;
  const pic = body.pic !== undefined ? (body.pic || null) : existing.pic;
  const opening_date = body.opening_date !== undefined ? (body.opening_date || null) : existing.opening_date;
  const target_date = body.target_date !== undefined ? (body.target_date || null) : existing.target_date;
  const completion_date = body.completion_date !== undefined ? (body.completion_date || null) : existing.completion_date;
  const status = body.status !== undefined ? (body.status !== null && body.status !== '' ? body.status : '') : existing.status;
  const notes = body.notes !== undefined ? (body.notes || null) : existing.notes;

  const baseVersion = isExistingNull ? clientVersion : existing.row_version;
  const nextRowVersion = baseVersion + 1;
  const updatedRecord = {
    ...existing,
    branch_id,
    activity_type,
    period,
    pic,
    opening_date,
    target_date,
    completion_date,
    status,
    notes,
    row_version: nextRowVersion,
    last_sync_source: 'FCMS'
  };

  const updateStmt = env.DB.prepare(
    `UPDATE activity_schedule SET
      branch_id = ?,
      activity_type = ?,
      period = ?,
      pic = ?,
      opening_date = ?,
      target_date = ?,
      completion_date = ?,
      status = ?,
      notes = ?,
      updated_at = datetime('now'),
      row_version = COALESCE(row_version, ?) + 1,
      last_sync_source = 'FCMS'
     WHERE id = ? AND (row_version = ? OR (row_version IS NULL AND (? = 1 OR ? = 0)))`
  ).bind(
    branch_id,
    activity_type,
    period,
    pic,
    opening_date,
    target_date,
    completion_date,
    status,
    notes,
    clientVersion,
    id,
    clientVersion,
    clientVersion,
    clientVersion
  );

  const batchRes = await env.DB.batch([
    updateStmt,
    buildAuditStatement(env, user, 'UPDATE', 'schedule', id, existing, updatedRecord, null, 'SUCCESS', true),
    buildConditionalOutboxQuery(env, 'Time Line', id, 'UPDATE', mapDBToPayload('Time Line', updatedRecord))
  ]);

  const changes = batchRes?.[0]?.meta?.changes ?? batchRes?.[0]?.changes;
  if (changes === 0) {
    return error('Conflict: Record has been modified by another user or session', 409, origin);
  }

  const updated = await env.DB.prepare('SELECT * FROM activity_schedule WHERE id = ?').bind(id).first();
  if (updated) {
    try { await runSync(env.DB, 'schedule', id, updated); } catch (e) { console.error('sync error', e.message); }
  }

  return ok({ message: 'Schedule updated', row_version: nextRowVersion }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function deleteSchedule(id, env, origin, user = null) {
  const existing = await env.DB.prepare('SELECT * FROM activity_schedule WHERE id = ? AND deleted_at IS NULL').bind(id).first();
  if (!existing) return notFound(origin);
  
  const deletedRecord = { ...existing, deleted_at: new Date().toISOString() };
  const { buildAuditStatement } = await import('../utils/audit.js');
  const auditStmt = buildAuditStatement(env, user, 'DELETE', 'schedule', id, existing, null, null, 'SUCCESS');

  await env.DB.batch([
    env.DB.prepare(`UPDATE activity_schedule SET deleted_at = datetime('now'), updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ?`).bind(id),
    buildOutboxQuery(env, 'Time Line', id, 'DELETE', mapDBToPayload('Time Line', deletedRecord)),
    auditStmt
  ]);
  
  try { await runSync(env.DB, 'schedule', id, null); } catch (e) { /* non-fatal */ }
  return ok({ message: 'Schedule deleted' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
}

export async function importSchedule(request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
  if (body.length === 0) return ok({ message: 'No data to import' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });

  const existing = await env.DB.prepare('SELECT * FROM activity_schedule').all();
  const existingMap = new Map();
  (existing.results || []).forEach(s => {
    if (s.activity_type && s.period) {
      existingMap.set(s.activity_type.toLowerCase().trim() + '_' + s.period.toLowerCase().trim() + '_' + String(s.branch_id || 'null'), s);
    }
  });

  const stmts = [];
  let inserted = 0;
  let updated = 0;

  for (const item of body) {
    if (!item.activity_type || !item.period) continue;
    
    const key = item.activity_type.toLowerCase().trim() + '_' + item.period.toLowerCase().trim() + '_' + String(item.branch_id || 'null');
    const status = item.status !== null && item.status !== undefined && item.status !== '' ? item.status : '';
    const nextFields = {
      branch_id: item.branch_id || null,
      activity_type: item.activity_type,
      period: item.period,
      pic: item.pic || null,
      opening_date: item.opening_date || null,
      target_date: item.target_date || null,
      completion_date: item.completion_date || null,
      status,
      notes: item.notes || null
    };

    if (existingMap.has(key)) {
      const oldRow = existingMap.get(key);
      const afterRow = { ...oldRow, ...nextFields };
      stmts.push(
        env.DB.prepare(
          `UPDATE activity_schedule SET pic = ?, opening_date = ?, target_date = ?, completion_date = ?, status = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(
          nextFields.pic,
          nextFields.opening_date,
          nextFields.target_date,
          nextFields.completion_date,
          nextFields.status,
          nextFields.notes,
          oldRow.id
        ),
        buildAuditStatement(env, user, 'UPDATE', 'schedule', oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true)
      );
      existingMap.set(key, afterRow);
      updated++;
    } else {
      stmts.push(
        env.DB.prepare(
          `INSERT INTO activity_schedule (branch_id, activity_type, period, pic, opening_date, target_date, completion_date, status, notes) 
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(
          nextFields.branch_id,
          nextFields.activity_type,
          nextFields.period,
          nextFields.pic,
          nextFields.opening_date,
          nextFields.target_date,
          nextFields.completion_date,
          nextFields.status,
          nextFields.notes
        ),
        buildCreateAuditStatement(env, user, 'schedule', 'new', nextFields, { source: 'import' })
      );
      inserted++;
    }
  }

  try {
    if (stmts.length > 0) await executeAtomicImportBatch(env, stmts);
    return ok({ message: `Berhasil mengimport data (Baru: ${inserted}, Update: ${updated})` }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
  } catch (err) {
    console.error('Import error:', err); return error('Gagal mengimport data', 500, origin);
  }
}
