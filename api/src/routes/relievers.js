import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination, resolveSortOrder } from '../utils/pagination.js';
import { requireExportAccess, getExportPagination } from '../utils/export.js';
import { runSync } from '../utils/calendar.js';
import { buildOutboxQuery, buildConditionalOutboxQuery, buildOutboxFromLastInsertQuery } from '../utils/sync_engine.js';
import { mapDBToPayload } from '../utils/sync_mapper.js';
import { buildAuditStatement, buildCreateAuditStatement, buildCreateAuditFromOutboxStatement, executeAtomicImportBatch } from '../utils/audit.js';
import { executeJournaledImport, extractImportRequestMeta } from '../utils/import_journal.js';

const RELIEVER_SORT_MAP = Object.freeze({
  period_asc_date_desc: 'r.period ASC, r.backup_date DESC, r.id DESC',
  backup_date_desc: 'r.backup_date DESC, r.id DESC',
  backup_date_asc: 'r.backup_date ASC, r.id ASC',
});

export async function handleRelievers(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace('/api/relievers', '');

  if (request.method === 'GET' && path === '/export') {
    const { response } = await requireExportAccess(request, env, origin, 'relievers');
    if (response) return response;
    return list(request, env, origin, { isExport: true });
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'relievers', 'read')) return forbidden(origin);

  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'GET' && path === '') return list(request, env, origin);
  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'relievers', 'import')) return forbidden(origin);
    return importRelievers(request, env, origin, user);
  }
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'relievers', 'write')) return forbidden(origin);
    return create(request, env, origin, user);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') return getOne(id, env, origin);
    if (request.method === 'PUT' || request.method === 'PATCH') {
      if (!hasPermission(user, 'relievers', 'write')) return forbidden(origin);
      return update(id, request, env, origin, user);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'relievers', 'delete')) return forbidden(origin);
      return remove(id, env, origin, user);
    }
  }
  return error('Not found', 404, origin);
}

async function list(request, env, origin, { isExport = false } = {}) {
  const { page, limit, offset } = isExport ? getExportPagination(request.url) : getPagination(request.url);
  const url = new URL(request.url);
  const branch_id = (url.searchParams.get('branch_id') || '').trim();
  const period = (url.searchParams.get('period') || '').trim();
  const status = (url.searchParams.get('status') || '').trim();
  const reliever_name = (url.searchParams.get('reliever_name') || '').trim();
  const search = (url.searchParams.get('search') || '').trim();
  const month = (url.searchParams.get('month') || '').trim();
  const sort = url.searchParams.get('sort') || 'period_asc_date_desc';

  const conditions = ['r.deleted_at IS NULL'];
  const values = [];
  if (branch_id) { conditions.push('r.branch_id = ?'); values.push(branch_id); }
  if (period) { conditions.push('r.period = ?'); values.push(period); }
  if (status) { conditions.push('r.status = ?'); values.push(status); }
  if (reliever_name) { conditions.push('r.reliever_name LIKE ?'); values.push(`%${reliever_name}%`); }
  if (search) {
    conditions.push('(r.reliever_name LIKE ? OR r.original_fc_name LIKE ? OR r.reason LIKE ? OR b.full_name LIKE ?)');
    values.push(`%${search}%`, `%${search}%`, `%${search}%`, `%${search}%`);
  }
  if (month) { conditions.push("strftime('%Y-%m', r.backup_date) = ?"); values.push(month); }

  const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';
  const countJoin = search ? 'LEFT JOIN branches b ON r.branch_id = b.id' : '';
  const orderBy = resolveSortOrder(sort, RELIEVER_SORT_MAP, 'period_asc_date_desc');

  const [countResult, rows] = await Promise.all([
    env.DB.prepare(`SELECT COUNT(*) as total FROM relievers r ${countJoin} ${where}`).bind(...values).first(),
    env.DB.prepare(
      `SELECT r.*, b.full_name as branch_name,
       COALESCE(r.status, 'Pending') as status
       FROM relievers r
       LEFT JOIN branches b ON r.branch_id = b.id
       ${where} ORDER BY ${orderBy} LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all()
  ]);

  return paginated(rows.results || [], countResult?.total || 0, page, limit, origin, env?.ENVIRONMENT);
}

async function getOne(id, env, origin) {
  const row = await env.DB.prepare(
    'SELECT r.*, b.full_name as branch_name FROM relievers r LEFT JOIN branches b ON r.branch_id = b.id WHERE r.id = ?'
  ).bind(id).first();
  if (!row) return notFound(origin);
  return ok(row, 200, origin);
}

async function validateRelieverBranch(env, branch_id) {
  if (branch_id === undefined || branch_id === null || branch_id === '') return { valid: true, branchId: null };
  const parsed = Number(branch_id);
  if (!Number.isInteger(parsed) || parsed <= 0) return { valid: false, error: 'branch_id tidak valid' };
  const row = await env.DB.prepare('SELECT id FROM branches WHERE id = ? AND deleted_at IS NULL').bind(parsed).first();
  if (!row) return { valid: false, error: 'Cabang (branch_id) tidak ditemukan atau sudah dinonaktifkan' };
  return { valid: true, branchId: parsed };
}

async function create(request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
  const { branch_id, original_fc_name, period, reliever_name, backup_date, completion_date, reason, shift, status } = body;
  if (!reliever_name || !backup_date) return error('reliever_name and backup_date required', 400, origin);

  const branchCheck = await validateRelieverBranch(env, branch_id);
  if (!branchCheck.valid) return error(branchCheck.error, 400, origin);

  const normalizedStatus = status !== null && status !== undefined && status !== '' ? status : '';
  const newSnapshot = {
    branch_id: branch_id || null,
    original_fc_name: original_fc_name || null,
    period: period || null,
    reliever_name,
    backup_date,
    completion_date: completion_date || null,
    reason: reason || null,
    shift: shift || null,
    status: normalizedStatus,
    last_sync_source: 'FCMS',
    row_version: 1
  };
  const payload = mapDBToPayload('Jadwal Reliefer', {
    id: '',
    ...newSnapshot
  });
  const corrId = crypto.randomUUID();

  const batchRes = await env.DB.batch([
    env.DB.prepare(
      "INSERT INTO relievers (branch_id, original_fc_name, period, reliever_name, backup_date, completion_date, reason, shift, status, row_version, last_sync_source) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1, 'FCMS')"
    ).bind(branch_id || null, original_fc_name || null, period || null, reliever_name,
      backup_date, completion_date || null, reason || null, shift || null, normalizedStatus),
    buildOutboxFromLastInsertQuery(env, 'Jadwal Reliefer', payload, corrId),
    buildCreateAuditFromOutboxStatement(env, user, 'relievers', corrId, newSnapshot)
  ]);

  const newId = batchRes?.[0]?.meta?.last_row_id ?? 1;
  try {
    await runSync(env.DB, 'relievers', newId, {
      reliever_name, backup_date, status: normalizedStatus, branch_id, original_fc_name, reason, shift
    });
  } catch (e) { console.error('sync error', e.message); }

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

  const existing = await env.DB.prepare('SELECT * FROM relievers WHERE id = ?').bind(id).first();
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

  const { branch_id, original_fc_name, period, reliever_name, backup_date, completion_date, reason, shift, status } = body;
  const branchCheck = await validateRelieverBranch(env, branch_id);
  if (!branchCheck.valid) return error(branchCheck.error, 400, origin);
  
  const baseVersion = isExistingNull ? clientVersion : existing.row_version;
  const nextRowVersion = baseVersion + 1;
  const updatedRecord = { ...existing, branch_id: branch_id || existing.branch_id, original_fc_name: original_fc_name || existing.original_fc_name, period: period || existing.period, reliever_name: reliever_name || existing.reliever_name, backup_date: backup_date || existing.backup_date, completion_date: completion_date || existing.completion_date, reason: reason || existing.reason, shift: shift || existing.shift, status: status !== null && status !== undefined && status !== '' ? status : existing.status, row_version: nextRowVersion };
  
  const updateStmt = env.DB.prepare(
    `UPDATE relievers SET branch_id = COALESCE(?, branch_id), original_fc_name = COALESCE(?, original_fc_name),
     period = COALESCE(?, period), reliever_name = COALESCE(?, reliever_name), backup_date = COALESCE(?, backup_date),
     completion_date = COALESCE(?, completion_date), reason = COALESCE(?, reason),
     shift = COALESCE(?, shift), status = COALESCE(?, status), updated_at = datetime('now'), row_version = COALESCE(row_version, ?) + 1, last_sync_source = 'FCMS'
     WHERE id = ? AND (row_version = ? OR (row_version IS NULL AND (? = 1 OR ? = 0)))`
  ).bind(branch_id || null, original_fc_name || null, period || null, reliever_name || null,
    backup_date || null, completion_date || null, reason || null, shift || null, status !== null && status !== undefined && status !== '' ? status : '',
    clientVersion, id, clientVersion, clientVersion, clientVersion);

  const batchRes = await env.DB.batch([
    updateStmt,
    buildAuditStatement(env, user, 'UPDATE', 'relievers', id, existing, updatedRecord, null, 'SUCCESS', true),
    buildConditionalOutboxQuery(env, 'Jadwal Reliefer', id, 'UPDATE', mapDBToPayload('Jadwal Reliefer', { ...updatedRecord, last_sync_source: 'FCMS' }))
  ]);

  const changes = batchRes?.[0]?.meta?.changes ?? batchRes?.[0]?.changes;
  if (changes === 0) {
    return error('Conflict: Record has been modified by another user or session', 409, origin);
  }

  const updated = await env.DB.prepare('SELECT * FROM relievers WHERE id = ?').bind(id).first();
  if (updated) {
    try { await runSync(env.DB, 'relievers', id, updated); } catch (e) { console.error('sync error', e.message); }
  }

  return ok({ message: 'Updated', row_version: nextRowVersion }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function remove(id, env, origin, user = null) {
  const existing = await env.DB.prepare('SELECT * FROM relievers WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);
  
  const deletedRecord = { ...existing, deleted_at: new Date().toISOString() };
  const { buildAuditStatement } = await import('../utils/audit.js');
  const auditStmt = buildAuditStatement(env, user, 'DELETE', 'relievers', id, existing, null, null, 'SUCCESS');

  await env.DB.batch([
    env.DB.prepare(`UPDATE relievers SET deleted_at = datetime('now'), updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ?`).bind(id),
    buildOutboxQuery(env, 'Jadwal Reliefer', id, 'DELETE', mapDBToPayload('Jadwal Reliefer', deletedRecord)),
    auditStmt
  ]);
  
  try { await runSync(env.DB, 'relievers', id, null); } catch (e) { /* non-fatal */ }
  return ok({ message: 'Deleted' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function importRelievers(request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
  if (body.length === 0) return ok({ message: 'No data to import' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });

  const importMeta = extractImportRequestMeta(request);
  try {
    return await executeJournaledImport(env, user, origin, {
      module: 'relievers_route',
      rows: body,
      onDuplicate: 'upsert',
      importMeta,
      buildPlan: async ({ startRowIndex, initialCounts }) => {
        const existingAll = await env.DB.prepare('SELECT * FROM relievers').all();
        const existingRows = existingAll.results || [];
        const findExisting = (branchId, relieverName, backupDate) => {
          const bVal = branchId || null;
          return existingRows.find(r =>
            ((bVal !== null && r.branch_id === bVal) || r.branch_id === null) &&
            r.reliever_name === relieverName &&
            r.backup_date === backupDate
          );
        };

        const seenNewKeys = new Set();
        const insertGroups = [];
        let inserted = Number(initialCounts?.inserted || 0);
        let updated = Number(initialCounts?.updated || 0);
        let skipped = Number(initialCounts?.skipped || 0);

        for (let idx = startRowIndex; idx < body.length; idx++) {
          const item = body[idx];
          if (!item || typeof item !== 'object' || !item.reliever_name || !item.backup_date) {
            skipped++;
            continue;
          }
          const statusVal = item.status !== null && item.status !== undefined && item.status !== '' ? item.status : '';
          const existing = findExisting(item.branch_id, item.reliever_name, item.backup_date);
          const nextFields = {
            branch_id: item.branch_id || null,
            original_fc_name: item.original_fc_name || null,
            period: item.period || null,
            reliever_name: item.reliever_name,
            backup_date: item.backup_date,
            completion_date: item.completion_date || null,
            reason: item.reason || null,
            shift: item.shift || null,
            status: statusVal
          };
          const newKey = `${nextFields.branch_id || 'null'}|${nextFields.reliever_name}|${nextFields.backup_date}`;

          if (existing) {
            const afterRow = { ...existing, ...nextFields };
            updated++;
            insertGroups.push({
              rowIndex: idx,
              counts: { inserted, updated, skipped },
              stmts: [
                env.DB.prepare(
                  `UPDATE relievers SET original_fc_name = ?, period = ?, completion_date = ?, reason = ?, shift = ?, status = ? WHERE id = ?`
                ).bind(
                  nextFields.original_fc_name,
                  nextFields.period,
                  nextFields.completion_date,
                  nextFields.reason,
                  nextFields.shift,
                  nextFields.status,
                  existing.id
                ),
                buildAuditStatement(env, user, 'UPDATE', 'relievers', existing.id, existing, afterRow, { source: 'import' }, 'SUCCESS', true)
              ]
            });
            Object.assign(existing, nextFields);
          } else if (seenNewKeys.has(newKey)) {
            skipped++;
          } else {
            seenNewKeys.add(newKey);
            inserted++;
            insertGroups.push({
              rowIndex: idx,
              counts: { inserted, updated, skipped },
              stmts: [
                env.DB.prepare(
                  `INSERT INTO relievers (branch_id, original_fc_name, period, reliever_name, backup_date, completion_date, reason, shift, status) 
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
                ).bind(
                  nextFields.branch_id,
                  nextFields.original_fc_name,
                  nextFields.period,
                  nextFields.reliever_name,
                  nextFields.backup_date,
                  nextFields.completion_date,
                  nextFields.reason,
                  nextFields.shift,
                  nextFields.status
                ),
                buildCreateAuditStatement(env, user, 'relievers', 'new', nextFields, { source: 'import' })
              ]
            });
          }
        }

        return { insertGroups, finalCounts: { inserted, updated, skipped } };
      },
      buildFinalResult: (counts) => {
        const processed = Number(counts.inserted || 0) + Number(counts.updated || 0);
        return {
          message: `Berhasil memperbarui/mengimport ${processed} data`,
          inserted: counts.inserted,
          updated: counts.updated,
          skipped: counts.skipped,
        };
      },
    });
  } catch (err) {
    console.error('Import error:', err); return error('Gagal mengimport data', 500, origin);
  }
}
