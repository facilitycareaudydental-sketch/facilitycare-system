import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination, resolveSortOrder } from '../utils/pagination.js';
import { requireExportAccess, getExportPagination } from '../utils/export.js';
import { buildOutboxQuery, buildConditionalOutboxQuery, buildOutboxFromLastInsertQuery } from '../utils/sync_engine.js';
import { mapDBToPayload } from '../utils/sync_mapper.js';
import { hasTableColumn } from '../utils/schemaCompat.js';
import { buildAuditStatement, buildCreateAuditStatement, buildCreateAuditFromOutboxStatement, executeAtomicImportBatch } from '../utils/audit.js';

const ONE_ON_ONE_SORT_MAP = Object.freeze({
  date_desc: 'o.meeting_date DESC, o.id DESC',
  date_asc: 'o.meeting_date ASC, o.id ASC',
});

export async function handleOneOnOne(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace('/api/one-on-one', '');

  if (request.method === 'GET' && path === '/export') {
    const { response } = await requireExportAccess(request, env, origin, 'one_on_one');
    if (response) return response;
    return list(request, env, origin, { isExport: true });
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'one_on_one', 'read')) return forbidden(origin);

  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'GET' && path === '') return list(request, env, origin);
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'one_on_one', 'write')) return forbidden(origin);
    return create(request, env, origin, user);
  }
  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'one_on_one', 'import')) return forbidden(origin);
    let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
    if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
    if (body.length === 0) return ok({ message: 'No data to import' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });

    const existing = await env.DB.prepare('SELECT * FROM one_on_one').all();
    const existingMap = new Map();
    (existing.results || []).forEach(r => {
      if (r.meeting_date && r.employee_name && r.branch_id) {
        existingMap.set(r.meeting_date + '_' + r.employee_name.toLowerCase().trim() + '_' + r.branch_id, r);
      }
    });

    const stmts = [];
    let inserted = 0; let updated = 0;
    for (const item of body) {
      if (!item.meeting_date || !item.employee_name || !item.branch_id) continue;
      const key = item.meeting_date + '_' + item.employee_name.toLowerCase().trim() + '_' + item.branch_id;
      const status = item.status || 'Open';
      const dayCount = item.completion_date ? Math.max(0, Math.floor((new Date(item.completion_date) - new Date(item.meeting_date)) / 86400000)) : 0;
      const newSnapshot = {
        meeting_date: item.meeting_date,
        branch_id: item.branch_id,
        employee_name: item.employee_name,
        pic: item.pic || null,
        problem: item.problem || null,
        solution: item.solution || null,
        status,
        completion_date: item.completion_date || null,
        day_count: dayCount,
        document_link: item.document_link || null
      };

      if (existingMap.has(key)) {
        const oldRow = existingMap.get(key);
        stmts.push(
          env.DB.prepare(`UPDATE one_on_one SET pic=?, problem=?, solution=?, status=?, completion_date=?, day_count=?, document_link=?, updated_at=datetime('now') WHERE id=?`)
            .bind(item.pic || null, item.problem || null, item.solution || null, status, item.completion_date || null, dayCount, item.document_link || null, oldRow.id),
          buildAuditStatement(env, user, 'UPDATE', 'one_on_one', oldRow.id, oldRow, { ...oldRow, ...newSnapshot }, { source: 'import' }, 'SUCCESS', true)
        );
        existingMap.set(key, { ...oldRow, ...newSnapshot });
        updated++;
      } else {
        stmts.push(
          env.DB.prepare(`INSERT INTO one_on_one (meeting_date, branch_id, employee_name, pic, problem, solution, status, completion_date, day_count, document_link) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`)
            .bind(item.meeting_date, item.branch_id, item.employee_name, item.pic || null, item.problem || null, item.solution || null, status, item.completion_date || null, dayCount, item.document_link || null),
          buildCreateAuditStatement(env, user, 'one_on_one', 'new', newSnapshot, { source: 'import' })
        );
        inserted++;
      }
    }
    try {
      if (stmts.length > 0) await executeAtomicImportBatch(env, stmts);
      return ok({ message: 'Import sukses', inserted, updated, total: body.length }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
    } catch (e) {
      console.error('Import error:', e); return error('Gagal mengimport data', 500, origin);
    }
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') return getOne(id, env, origin);
    if (request.method === 'PUT' || request.method === 'PATCH') {
      if (!hasPermission(user, 'one_on_one', 'write')) return forbidden(origin);
      return update(id, request, env, origin, user);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'one_on_one', 'delete')) return forbidden(origin);
      return remove(id, env, origin, user);
    }
  }
  return error('Not found', 404, origin);
}

async function list(request, env, origin, { isExport = false } = {}) {
  const { page, limit, offset } = isExport ? getExportPagination(request.url) : getPagination(request.url);
  const url = new URL(request.url);
  const branch_id = (url.searchParams.get('branch_id') || '').trim();
  const status = (url.searchParams.get('status') || '').trim();
  const dashFilter = (url.searchParams.get('dash_filter') || '').trim();
  const search = (url.searchParams.get('search') || '').trim();
  const sort = url.searchParams.get('sort') || 'date_desc';

  const conditions = ['o.deleted_at IS NULL'];
  const values = [];
  if (branch_id) { conditions.push('o.branch_id = ?'); values.push(branch_id); }
  if (status) {
    conditions.push('LOWER(o.status) = LOWER(?)');
    values.push(status);
  } else if (dashFilter === 'pending') {
    conditions.push("LOWER(o.status) = 'pending'");
  }
  if (search) {
    conditions.push('(o.employee_name LIKE ? OR o.problem LIKE ? OR b.full_name LIKE ?)');
    values.push(`%${search}%`, `%${search}%`, `%${search}%`);
  }

  const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';
  const countJoin = search ? 'LEFT JOIN branches b ON o.branch_id = b.id' : '';
  const orderBy = resolveSortOrder(sort, ONE_ON_ONE_SORT_MAP, 'date_desc');

  const [countResult, rows] = await Promise.all([
    env.DB.prepare(`SELECT COUNT(*) as total FROM one_on_one o ${countJoin} ${where}`).bind(...values).first(),
    env.DB.prepare(
      `SELECT o.*, b.full_name as branch_name FROM one_on_one o
       LEFT JOIN branches b ON o.branch_id = b.id
       ${where} ORDER BY ${orderBy} LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all()
  ]);

  return paginated(rows.results || [], countResult?.total || 0, page, limit, origin, env?.ENVIRONMENT);
}

async function getOne(id, env, origin) {
  const row = await env.DB.prepare(
    'SELECT o.*, b.full_name as branch_name FROM one_on_one o LEFT JOIN branches b ON o.branch_id = b.id WHERE o.id = ?'
  ).bind(id).first();
  if (!row) return notFound(origin);
  return ok(row, 200, origin);
}

async function validateOneOnOneBranch(env, branch_id) {
  if (branch_id === undefined || branch_id === null || branch_id === '') return { valid: true, branchId: null };
  const parsed = Number(branch_id);
  if (!Number.isInteger(parsed) || parsed <= 0) return { valid: false, error: 'branch_id tidak valid' };
  const row = await env.DB.prepare('SELECT id FROM branches WHERE id = ? AND deleted_at IS NULL').bind(parsed).first();
  if (!row) return { valid: false, error: 'Cabang (branch_id) tidak ditemukan atau sudah dinonaktifkan' };
  return { valid: true, branchId: parsed };
}

async function resolveOneOnOneEmployeeId(env, rawEmpId, existingEmpId = null, hasExisting = false) {
  if (rawEmpId === undefined) {
    return { valid: true, employeeId: hasExisting ? (existingEmpId ?? null) : null };
  }
  if (rawEmpId === null || rawEmpId === '') {
    return { valid: true, employeeId: null };
  }
  const parsed = Number(rawEmpId);
  if (!Number.isInteger(parsed) || parsed <= 0) {
    return { valid: false, error: 'Invalid employee_id' };
  }
  const empRow = await env.DB.prepare('SELECT id FROM employees WHERE id = ? AND deleted_at IS NULL').bind(parsed).first();
  if (!empRow) {
    const eCount = await env.DB.prepare('SELECT COUNT(*) as c FROM employees').first();
    if (eCount && typeof eCount.c === 'number' && eCount.c > 0) {
      return { valid: false, error: 'Karyawan (employee_id) tidak ditemukan atau sudah dinonaktifkan' };
    }
  }
  return { valid: true, employeeId: parsed };
}

async function create(request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
  const { meeting_date, branch_id, employee_id, employee_name, pic, problem, solution, status, completion_date, document_link } = body;
  if (!meeting_date || !employee_name || !problem) return error('meeting_date, employee_name, problem required', 400, origin);

  const [branchCheck, empCheck, hasEmpIdCol] = await Promise.all([
    validateOneOnOneBranch(env, branch_id),
    resolveOneOnOneEmployeeId(env, employee_id, null, false),
    hasTableColumn(env.DB, 'one_on_one', 'employee_id')
  ]);
  if (!branchCheck.valid) return error(branchCheck.error, 400, origin);
  if (!empCheck.valid) return error(empCheck.error, 400, origin);

  let day_count = null;
  if (completion_date && meeting_date) {
    day_count = Math.floor((new Date(completion_date) - new Date(meeting_date)) / (1000 * 60 * 60 * 24));
  }

  const normalizedStatus = status !== null && status !== undefined && status !== '' ? status : '';
  const payloadData = {
    id: '',
    meeting_date,
    branch_id: branch_id || null,
    employee_name,
    pic: pic || null,
    problem,
    solution: solution || null,
    status: normalizedStatus,
    completion_date: completion_date || null,
    day_count,
    document_link: document_link || null,
    last_sync_source: 'FCMS',
    row_version: 1
  };
  if (hasEmpIdCol) payloadData.employee_id = empCheck.employeeId;
  const payload = mapDBToPayload('One on One', payloadData);
  const { id: _ignoredId, ...newSnapshot } = payloadData;

  const insertStmt = hasEmpIdCol
    ? env.DB.prepare(
        "INSERT INTO one_on_one (meeting_date, branch_id, employee_id, employee_name, pic, problem, solution, status, completion_date, day_count, document_link, row_version, last_sync_source) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, 'FCMS')"
      ).bind(meeting_date, branch_id || null, empCheck.employeeId, employee_name, pic || null, problem,
        solution || null, normalizedStatus, completion_date || null, day_count, document_link || null)
    : env.DB.prepare(
        "INSERT INTO one_on_one (meeting_date, branch_id, employee_name, pic, problem, solution, status, completion_date, day_count, document_link, row_version, last_sync_source) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, 'FCMS')"
      ).bind(meeting_date, branch_id || null, employee_name, pic || null, problem,
        solution || null, normalizedStatus, completion_date || null, day_count, document_link || null);

  const corrId = crypto.randomUUID();
  const batchRes = await env.DB.batch([
    insertStmt,
    buildOutboxFromLastInsertQuery(env, 'One on One', payload, corrId),
    buildCreateAuditFromOutboxStatement(env, user, 'one_on_one', corrId, newSnapshot)
  ]);

  const newId = batchRes?.[0]?.meta?.last_row_id ?? 1;

  return ok({ id: newId }, 201, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function update(id, request, env, origin, user = null) {
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

  const existing = await env.DB.prepare('SELECT * FROM one_on_one WHERE id = ?').bind(id).first();
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

  const { meeting_date, branch_id, employee_id, employee_name, pic, problem, solution, status, completion_date, document_link } = body;
  const [branchCheck, empCheck, hasEmpIdCol] = await Promise.all([
    validateOneOnOneBranch(env, branch_id),
    resolveOneOnOneEmployeeId(env, employee_id, existing.employee_id, true),
    hasTableColumn(env.DB, 'one_on_one', 'employee_id')
  ]);
  if (!branchCheck.valid) return error(branchCheck.error, 400, origin);
  if (!empCheck.valid) return error(empCheck.error, 400, origin);

  const rd = meeting_date || existing.meeting_date;
  const cd = completion_date !== undefined ? completion_date : existing.completion_date;
  let day_count = existing.day_count;
  if (rd && cd) day_count = Math.floor((new Date(cd) - new Date(rd)) / (1000 * 60 * 60 * 24));

  const baseVersion = isExistingNull ? clientVersion : existing.row_version;
  const nextRowVersion = baseVersion + 1;
  const updatedRecord = {
    ...existing,
    meeting_date: meeting_date || existing.meeting_date,
    branch_id: branch_id || existing.branch_id,
    employee_name: employee_name || existing.employee_name,
    pic: pic || existing.pic,
    problem: problem || existing.problem,
    solution: solution || existing.solution,
    status: status !== null && status !== undefined && status !== '' ? status : existing.status,
    completion_date: completion_date || existing.completion_date,
    day_count,
    document_link: document_link || existing.document_link,
    row_version: nextRowVersion
  };
  if (hasEmpIdCol) updatedRecord.employee_id = empCheck.employeeId;

  const updateStmt = hasEmpIdCol
    ? env.DB.prepare(
        `UPDATE one_on_one SET meeting_date = COALESCE(?, meeting_date), branch_id = COALESCE(?, branch_id),
         employee_id = ?, employee_name = COALESCE(?, employee_name), pic = COALESCE(?, pic), problem = COALESCE(?, problem),
         solution = COALESCE(?, solution), status = COALESCE(?, status), completion_date = COALESCE(?, completion_date),
         day_count = ?, document_link = COALESCE(?, document_link), updated_at = datetime('now'), row_version = COALESCE(row_version, ?) + 1, last_sync_source = 'FCMS'
         WHERE id = ? AND (row_version = ? OR (row_version IS NULL AND (? = 1 OR ? = 0)))`
      ).bind(meeting_date || null, branch_id || null, empCheck.employeeId, employee_name || null, pic || null, problem || null,
        solution || null, status !== null && status !== undefined && status !== '' ? status : '', completion_date || null, day_count, document_link || null,
        clientVersion, id, clientVersion, clientVersion, clientVersion)
    : env.DB.prepare(
        `UPDATE one_on_one SET meeting_date = COALESCE(?, meeting_date), branch_id = COALESCE(?, branch_id),
         employee_name = COALESCE(?, employee_name), pic = COALESCE(?, pic), problem = COALESCE(?, problem),
         solution = COALESCE(?, solution), status = COALESCE(?, status), completion_date = COALESCE(?, completion_date),
         day_count = ?, document_link = COALESCE(?, document_link), updated_at = datetime('now'), row_version = COALESCE(row_version, ?) + 1, last_sync_source = 'FCMS'
         WHERE id = ? AND (row_version = ? OR (row_version IS NULL AND (? = 1 OR ? = 0)))`
      ).bind(meeting_date || null, branch_id || null, employee_name || null, pic || null, problem || null,
        solution || null, status !== null && status !== undefined && status !== '' ? status : '', completion_date || null, day_count, document_link || null,
        clientVersion, id, clientVersion, clientVersion, clientVersion);

  const batchRes = await env.DB.batch([
    updateStmt,
    buildAuditStatement(env, user, 'UPDATE', 'one_on_one', id, existing, updatedRecord, null, 'SUCCESS', true),
    buildConditionalOutboxQuery(env, 'One on One', id, 'UPDATE', mapDBToPayload('One on One', { ...updatedRecord, last_sync_source: 'FCMS' }))
  ]);

  const changes = batchRes?.[0]?.meta?.changes ?? batchRes?.[0]?.changes;
  if (changes === 0) {
    return error('Conflict: Record has been modified by another user or session', 409, origin);
  }

  return ok({ message: 'Updated', row_version: nextRowVersion }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function remove(id, env, origin, user = null) {
  const existing = await env.DB.prepare('SELECT * FROM one_on_one WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);
  
  const deletedRecord = { ...existing, deleted_at: new Date().toISOString() };
  const { buildAuditStatement } = await import('../utils/audit.js');
  const auditStmt = buildAuditStatement(env, user, 'DELETE', 'one_on_one', id, existing, null, null, 'SUCCESS');

  await env.DB.batch([
    env.DB.prepare(`UPDATE one_on_one SET deleted_at = datetime('now'), updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ?`).bind(id),
    buildOutboxQuery(env, 'One on One', id, 'DELETE', mapDBToPayload('One on One', deletedRecord)),
    auditStmt
  ]);
  
  return ok({ message: 'Deleted' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
}
