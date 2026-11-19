import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination } from '../utils/pagination.js';
import { getExportPagination } from '../utils/export.js';
import {
  buildAuditStatement,
  buildCreateAuditStatement,
  buildSoftDeleteRecordAuditStatement
} from '../utils/audit.js';
import { executeJournaledImport, extractImportRequestMeta } from '../utils/import_journal.js';

async function softDeleteReport(env, user, tableName, moduleName, id, origin) {
  const existing = await env.DB.prepare(`SELECT * FROM ${tableName} WHERE id = ? AND deleted_at IS NULL`).bind(id).first();
  if (!existing) return notFound(origin);

  const deletedAtIso = new Date().toISOString();
  const delStmt = env.DB.prepare(
    `UPDATE ${tableName}
     SET deleted_at = ?, updated_at = ?, row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS'
     WHERE id = ? AND deleted_at IS NULL`
  ).bind(deletedAtIso, deletedAtIso, id);

  const auditStmt = buildSoftDeleteRecordAuditStatement(
    env, user, moduleName, tableName, 'id', 'deleted_at', deletedAtIso, existing
  );
  const batchRes = await env.DB.batch([delStmt, auditStmt]);
  const changes = Number(batchRes?.[0]?.meta?.changes ?? batchRes?.[0]?.changes ?? 0);
  if (changes === 0) return notFound(origin);

  return ok({ message: 'Deleted' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function atomicCreateReport(env, user, moduleName, insertStmt, newData, origin) {
  const auditStmt = buildCreateAuditStatement(env, user, moduleName, 'new', newData);
  const batchRes = await env.DB.batch([insertStmt, auditStmt]);
  return ok({ id: batchRes?.[0]?.meta?.last_row_id }, 201, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function atomicUpdateReport(env, user, moduleName, updateStmt, existing, newData, origin) {
  const auditStmt = buildAuditStatement(env, user, 'UPDATE', moduleName, existing.id, existing, newData, null, 'SUCCESS', true);
  const batchRes = await env.DB.batch([updateStmt, auditStmt]);
  const changes = Number(batchRes?.[0]?.meta?.changes ?? batchRes?.[0]?.changes ?? 0);
  if (changes === 0) return notFound(origin);
  return ok({ message: 'Updated' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
}

// Handles: /api/reports/inspection, /api/reports/cleaning, /api/reports/fogging, /api/reports/basecamp
export async function handleReports(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace('/api/reports', '');

  // Public supply requests
  if (path.startsWith('/supply')) return handleSupply(request, env, origin, path.replace('/supply', ''));

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'reports', 'read')) return forbidden(origin);

  if (path.startsWith('/inspection')) return handleInspection(request, env, user, origin, path.replace('/inspection', ''));
  if (path.startsWith('/cleaning')) return handleCleaning(request, env, user, origin, path.replace('/cleaning', ''));
  if (path.startsWith('/fogging')) return handleFogging(request, env, user, origin, path.replace('/fogging', ''));
  if (path.startsWith('/basecamp')) return handleBasecamp(request, env, user, origin, path.replace('/basecamp', ''));

  return error('Not found', 404, origin);
}

// ---- Generic CRUD helper ----
async function crudList(request, env, origin, table, joinClause = '', extraConditions = [], { useExport = false } = {}) {
  const { page, limit, offset } = useExport ? getExportPagination(request.url) : getPagination(request.url);
  const url = new URL(request.url);
  const branch_id = url.searchParams.get('branch_id') || '';
  const period = url.searchParams.get('period') || '';
  const status = url.searchParams.get('status') || '';
  const year = url.searchParams.get('year') || '';
  const search = url.searchParams.get('search') || '';

  const month = url.searchParams.get('month') || '';
  const pic = url.searchParams.get('pic') || '';
  const activity_type = url.searchParams.get('activity_type') || '';

  const dateField = table === 'basecamp_reports' ? 'info_date' : 
                    table === 'inspection_reports' ? 'inspection_date' : 'activity_date';

  let conditions = ['t.deleted_at IS NULL', ...extraConditions];
  let values = [];
  if (branch_id) { conditions.push('t.branch_id = ?'); values.push(branch_id); }
  if (period) { conditions.push('t.period = ?'); values.push(period); }
  if (status) { conditions.push('t.status = ?'); values.push(status); }
  if (pic && table === 'basecamp_reports') { conditions.push('t.pic = ?'); values.push(pic); }
  if (activity_type && (table === 'cleaning_reports' || table === 'fogging_reports')) {
    conditions.push('t.activity_type = ?');
    values.push(activity_type);
  }
  if (year) {
    conditions.push(`strftime('%Y', t.${dateField}) = ?`);
    values.push(year);
  }
  if (month) {
    conditions.push(`strftime('%m', t.${dateField}) = ?`);
    values.push(month);
  }
  if (search) {
    if (table === 'cleaning_reports' || table === 'fogging_reports') {
      conditions.push('(b.full_name LIKE ? OR t.activity_type LIKE ?)');
      values.push(`%${search}%`, `%${search}%`);
    } else {
      conditions.push('(b.full_name LIKE ?)');
      values.push(`%${search}%`);
    }
  }

  const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';
  const countJoin = (search || joinClause) ? `LEFT JOIN branches b ON t.branch_id = b.id ${joinClause}` : '';

  const [countResult, rows] = await Promise.all([
    env.DB.prepare(`SELECT COUNT(*) as total FROM ${table} t ${countJoin} ${where}`).bind(...values).first(),
    env.DB.prepare(
      `SELECT t.*, b.full_name as branch_name FROM ${table} t
       LEFT JOIN branches b ON t.branch_id = b.id
       ${where} ORDER BY t.${dateField} DESC, t.id DESC LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all()
  ]);

  return paginated(rows.results, countResult.total, page, limit, origin);
}

// ---- INSPECTION ----
async function handleInspection(request, env, user, origin, path) {
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === 'GET' && path === '/export') {
    if (!hasPermission(user, 'reports', 'export')) return forbidden(origin);
    return crudList(request, env, origin, 'inspection_reports', '', [], { useExport: true });
  }
  if (request.method === 'GET' && path === '') return crudList(request, env, origin, 'inspection_reports');
  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'reports', 'import')) return forbidden(origin);
    return importInspection(request, env, user, origin);
  }
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'reports', 'write')) return forbidden(origin);
    return createInspection(request, env, user, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') {
      const row = await env.DB.prepare('SELECT t.*, b.full_name as branch_name FROM inspection_reports t LEFT JOIN branches b ON t.branch_id = b.id WHERE t.id = ? AND t.deleted_at IS NULL').bind(id).first();
      return row ? ok(row, 200, origin) : notFound(origin);
    }
    if (request.method === 'PUT' || request.method === 'PATCH') {
      if (!hasPermission(user, 'reports', 'write')) return forbidden(origin);
      return updateInspection(id, request, env, user, origin);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'reports', 'delete')) return forbidden(origin);
      return softDeleteReport(env, user, 'inspection_reports', 'reports/inspection', id, origin);
    }
  }
  return error('Not found', 404, origin);
}

async function validateBranchExists(env, branch_id) {
  if (branch_id === undefined || branch_id === null || branch_id === '') return { valid: true, branchId: null };
  const parsed = Number(branch_id);
  if (!Number.isInteger(parsed) || parsed <= 0) {
    return { valid: false, error: 'branch_id tidak valid' };
  }
  const branch = await env.DB.prepare('SELECT id FROM branches WHERE id = ? AND deleted_at IS NULL').bind(parsed).first();
  if (!branch) {
    return { valid: false, error: 'Cabang (branch_id) tidak ditemukan atau sudah dinonaktifkan' };
  }
  return { valid: true, branchId: parsed };
}

function validateScoreField(rawScore, fieldName) {
  if (rawScore === undefined || rawScore === null || rawScore === '') return { valid: true, value: null };
  const num = Number(rawScore);
  if (isNaN(num) || num < 0 || num > 100) {
    return { valid: false, error: `${fieldName} harus berupa angka antara 0 dan 100` };
  }
  return { valid: true, value: num };
}

async function createInspection(request, env, user, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
  const { branch_id, period, inspection_date, status, fc_score, spv_score, document_link, notes } = body;
  if (!period || !inspection_date) return error('period and inspection_date required', 400, origin);

  const branchCheck = await validateBranchExists(env, branch_id);
  if (!branchCheck.valid) return error(branchCheck.error, 400, origin);

  const fcCheck = validateScoreField(fc_score, 'fc_score');
  if (!fcCheck.valid) return error(fcCheck.error, 400, origin);
  const spvCheck = validateScoreField(spv_score, 'spv_score');
  if (!spvCheck.valid) return error(spvCheck.error, 400, origin);

  const finalStatus = status !== null && status !== undefined && status !== '' ? status : '';
  const insertStmt = env.DB.prepare(
    'INSERT INTO inspection_reports (branch_id, period, inspection_date, status, fc_score, spv_score, document_link, notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?)'
  ).bind(branchCheck.branchId, period, inspection_date, finalStatus, fcCheck.value, spvCheck.value, document_link || null, notes || null);

  return atomicCreateReport(env, user, 'reports/inspection', insertStmt, {
    branch_id: branchCheck.branchId,
    period,
    inspection_date,
    status: finalStatus,
    fc_score: fcCheck.value,
    spv_score: spvCheck.value,
    document_link: document_link || null,
    notes: notes || null
  }, origin);
}

async function executeJournaledReportImport(request, env, user, origin, spec) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
  if (body.length === 0) return ok({ message: 'No data to import' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });

  const importMeta = extractImportRequestMeta(request);
  try {
    return await executeJournaledImport(env, user, origin, {
      module: spec.module,
      rows: body,
      onDuplicate: 'upsert',
      importMeta,
      buildPlan: async ({ startRowIndex, initialCounts }) => {
        const existing = await env.DB.prepare(spec.selectSql).all();
        const existingMap = new Map();
        (existing.results || []).forEach(r => {
          const k = spec.getExistingKey(r);
          if (k) existingMap.set(k, r);
        });

        const seenNewKeys = new Set();
        const insertGroups = [];
        let inserted = Number(initialCounts?.inserted || 0);
        let updated = Number(initialCounts?.updated || 0);
        let skipped = Number(initialCounts?.skipped || 0);

        for (let idx = startRowIndex; idx < body.length; idx++) {
          const item = body[idx];
          const key = item && typeof item === 'object' ? spec.getItemKey(item) : null;
          if (!key) { skipped++; continue; }

          if (existingMap.has(key)) {
            const oldRow = existingMap.get(key);
            const { stmt, afterRow } = spec.buildUpdate(env, item, oldRow);
            updated++;
            insertGroups.push({
              rowIndex: idx,
              counts: { inserted, updated, skipped },
              stmts: [
                stmt,
                buildAuditStatement(env, user, 'UPDATE', spec.auditModule || spec.module, oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true),
              ],
            });
            existingMap.set(key, afterRow);
          } else if (seenNewKeys.has(key)) {
            skipped++;
          } else {
            seenNewKeys.add(key);
            const { stmt, snapshot } = spec.buildInsert(env, item);
            inserted++;
            insertGroups.push({
              rowIndex: idx,
              counts: { inserted, updated, skipped },
              stmts: [
                stmt,
                buildCreateAuditStatement(env, user, spec.auditModule || spec.module, 'new', snapshot, { source: 'import' }),
              ],
            });
          }
        }

        return { insertGroups, finalCounts: { inserted, updated, skipped } };
      },
      buildFinalResult: (counts) => ({
        message: 'Berhasil mengimport data',
        inserted: counts.inserted,
        updated: counts.updated,
        failed: 0,
        total: body.length,
      }),
    });
  } catch (err) {
    console.error('Import error:', err);
    return error('Gagal mengimport data', 500, origin);
  }
}

async function importInspection(request, env, user, origin) {
  return executeJournaledReportImport(request, env, user, origin, {
    module: 'reports_inspection_route',
    auditModule: 'reports/inspection',
    selectSql: 'SELECT * FROM inspection_reports WHERE deleted_at IS NULL',
    getExistingKey: (s) => (s.branch_id && s.period && s.inspection_date)
      ? `${s.branch_id}_${s.period.toLowerCase().trim()}_${String(s.inspection_date).slice(0, 4)}`
      : null,
    getItemKey: (item) => (item.branch_id && item.period && item.inspection_date)
      ? `${item.branch_id}_${String(item.period).toLowerCase().trim()}_${String(item.inspection_date).slice(0, 4)}`
      : null,
    buildUpdate: (envDb, item, oldRow) => {
      const status = item.status !== null && item.status !== undefined && item.status !== '' ? item.status : '';
      const afterRow = {
        ...oldRow,
        inspection_date: item.inspection_date,
        fc_score: item.fc_score,
        spv_score: item.spv_score,
        status,
        document_link: item.document_link || null,
        notes: item.notes || null,
      };
      return {
        afterRow,
        stmt: envDb.DB.prepare(
          `UPDATE inspection_reports SET inspection_date = ?, fc_score = ?, spv_score = ?, status = ?, document_link = ?, notes = ?, updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ? AND deleted_at IS NULL`
        ).bind(item.inspection_date, item.fc_score, item.spv_score, status, item.document_link || null, item.notes || null, oldRow.id),
      };
    },
    buildInsert: (envDb, item) => {
      const status = item.status !== null && item.status !== undefined && item.status !== '' ? item.status : '';
      const snapshot = {
        branch_id: item.branch_id,
        period: item.period,
        inspection_date: item.inspection_date,
        fc_score: item.fc_score,
        spv_score: item.spv_score,
        status,
        document_link: item.document_link || null,
        notes: item.notes || null,
      };
      return {
        snapshot,
        stmt: envDb.DB.prepare(
          `INSERT INTO inspection_reports (branch_id, period, inspection_date, fc_score, spv_score, status, document_link, notes) 
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(snapshot.branch_id, snapshot.period, snapshot.inspection_date, snapshot.fc_score, snapshot.spv_score, snapshot.status, snapshot.document_link, snapshot.notes),
      };
    },
  });
}

async function updateInspection(id, request, env, user, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
  const existing = await env.DB.prepare('SELECT * FROM inspection_reports WHERE id = ? AND deleted_at IS NULL').bind(id).first();
  if (!existing) return notFound(origin);
  const { branch_id, period, inspection_date, status, fc_score, spv_score, document_link, notes } = body;

  const branchCheck = await validateBranchExists(env, branch_id);
  if (!branchCheck.valid) return error(branchCheck.error, 400, origin);
  const fcCheck = validateScoreField(fc_score, 'fc_score');
  if (!fcCheck.valid) return error(fcCheck.error, 400, origin);
  const spvCheck = validateScoreField(spv_score, 'spv_score');
  if (!spvCheck.valid) return error(spvCheck.error, 400, origin);

  const nextStatus = status !== null && status !== undefined && status !== '' ? status : null;
  const updateStmt = env.DB.prepare(
    `UPDATE inspection_reports SET branch_id = COALESCE(?, branch_id), period = COALESCE(?, period),
     inspection_date = COALESCE(?, inspection_date), status = COALESCE(?, status),
     fc_score = COALESCE(?, fc_score), spv_score = COALESCE(?, spv_score),
     document_link = COALESCE(?, document_link), notes = COALESCE(?, notes),
     updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS'
     WHERE id = ? AND deleted_at IS NULL`
  ).bind(branchCheck.branchId, period || null, inspection_date || null, nextStatus,
    fcCheck.value, spvCheck.value, document_link || null, notes || null, id);

  const newData = {
    ...existing,
    branch_id: branchCheck.branchId ?? existing.branch_id,
    period: period || existing.period,
    inspection_date: inspection_date || existing.inspection_date,
    status: nextStatus ?? existing.status,
    fc_score: fcCheck.value ?? existing.fc_score,
    spv_score: spvCheck.value ?? existing.spv_score,
    document_link: document_link || existing.document_link,
    notes: notes || existing.notes
  };
  return atomicUpdateReport(env, user, 'reports/inspection', updateStmt, existing, newData, origin);
}

// ---- CLEANING ----
async function handleCleaning(request, env, user, origin, path) {
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === 'GET' && path === '/export') {
    if (!hasPermission(user, 'reports', 'export')) return forbidden(origin);
    return crudList(request, env, origin, 'cleaning_reports', '', [], { useExport: true });
  }
  if (request.method === 'GET' && path === '') return crudList(request, env, origin, 'cleaning_reports');
  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'reports', 'import')) return forbidden(origin);
    return importCleaning(request, env, user, origin);
  }
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'reports', 'write')) return forbidden(origin);
    let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
    if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
    const { branch_id, activity_type, period, activity_date, status, document_link, notes } = body;
    if (!activity_type || !period || !activity_date) return error('activity_type, period, activity_date required', 400, origin);
    const branchCheck = await validateBranchExists(env, branch_id);
    if (!branchCheck.valid) return error(branchCheck.error, 400, origin);
    const finalStatus = status !== null && status !== undefined && status !== '' ? status : '';
    const insertStmt = env.DB.prepare(
      'INSERT INTO cleaning_reports (branch_id, activity_type, period, activity_date, status, document_link, notes) VALUES (?, ?, ?, ?, ?, ?, ?)'
    ).bind(branchCheck.branchId, activity_type, period, activity_date, finalStatus, document_link || null, notes || null);
    return atomicCreateReport(env, user, 'reports/cleaning', insertStmt, {
      branch_id: branchCheck.branchId,
      activity_type,
      period,
      activity_date,
      status: finalStatus,
      document_link: document_link || null,
      notes: notes || null
    }, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') {
      const row = await env.DB.prepare('SELECT t.*, b.full_name as branch_name FROM cleaning_reports t LEFT JOIN branches b ON t.branch_id = b.id WHERE t.id = ? AND t.deleted_at IS NULL').bind(id).first();
      return row ? ok(row, 200, origin) : notFound(origin);
    }
    if (request.method === 'PUT' || request.method === 'PATCH') {
      if (!hasPermission(user, 'reports', 'write')) return forbidden(origin);
      let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
      if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
      const existing = await env.DB.prepare('SELECT * FROM cleaning_reports WHERE id = ? AND deleted_at IS NULL').bind(id).first();
      if (!existing) return notFound(origin);
      const { branch_id, activity_type, period, activity_date, status, document_link, notes } = body;
      const branchCheck = await validateBranchExists(env, branch_id);
      if (!branchCheck.valid) return error(branchCheck.error, 400, origin);
      const nextStatus = status !== null && status !== undefined && status !== '' ? status : null;
      const updateStmt = env.DB.prepare(
        `UPDATE cleaning_reports SET branch_id = COALESCE(?, branch_id), activity_type = COALESCE(?, activity_type), period = COALESCE(?, period), activity_date = COALESCE(?, activity_date), status = COALESCE(?, status), document_link = COALESCE(?, document_link), notes = COALESCE(?, notes), updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ? AND deleted_at IS NULL`
      ).bind(branchCheck.branchId, activity_type || null, period || null, activity_date || null, nextStatus, document_link || null, notes || null, id);
      const newData = {
        ...existing,
        branch_id: branchCheck.branchId ?? existing.branch_id,
        activity_type: activity_type || existing.activity_type,
        period: period || existing.period,
        activity_date: activity_date || existing.activity_date,
        status: nextStatus ?? existing.status,
        document_link: document_link || existing.document_link,
        notes: notes || existing.notes
      };
      return atomicUpdateReport(env, user, 'reports/cleaning', updateStmt, existing, newData, origin);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'reports', 'delete')) return forbidden(origin);
      return softDeleteReport(env, user, 'cleaning_reports', 'reports/cleaning', id, origin);
    }
  }
  return error('Not found', 404, origin);
}

async function importCleaning(request, env, user, origin) {
  return executeJournaledReportImport(request, env, user, origin, {
    module: 'reports_cleaning_route',
    auditModule: 'reports/cleaning',
    selectSql: 'SELECT * FROM cleaning_reports WHERE deleted_at IS NULL',
    getExistingKey: (s) => (s.branch_id && s.activity_type && s.period && s.activity_date)
      ? `${s.branch_id}_${s.activity_type.toLowerCase().trim()}_${s.period.toLowerCase().trim()}_${String(s.activity_date).slice(0, 4)}`
      : null,
    getItemKey: (item) => (item.branch_id && item.activity_type && item.period && item.activity_date)
      ? `${item.branch_id}_${String(item.activity_type).toLowerCase().trim()}_${String(item.period).toLowerCase().trim()}_${String(item.activity_date).slice(0, 4)}`
      : null,
    buildUpdate: (envDb, item, oldRow) => {
      const status = item.status !== null && item.status !== undefined && item.status !== '' ? item.status : '';
      const afterRow = {
        ...oldRow,
        activity_date: item.activity_date,
        status,
        document_link: item.document_link || null,
        notes: item.notes || null,
      };
      return {
        afterRow,
        stmt: envDb.DB.prepare(
          `UPDATE cleaning_reports SET activity_date = ?, status = ?, document_link = ?, notes = ?, updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ? AND deleted_at IS NULL`
        ).bind(item.activity_date, status, item.document_link || null, item.notes || null, oldRow.id),
      };
    },
    buildInsert: (envDb, item) => {
      const status = item.status !== null && item.status !== undefined && item.status !== '' ? item.status : '';
      const snapshot = {
        branch_id: item.branch_id,
        activity_type: item.activity_type,
        period: item.period,
        activity_date: item.activity_date,
        status,
        document_link: item.document_link || null,
        notes: item.notes || null,
      };
      return {
        snapshot,
        stmt: envDb.DB.prepare(
          `INSERT INTO cleaning_reports (branch_id, activity_type, period, activity_date, status, document_link, notes) 
           VALUES (?, ?, ?, ?, ?, ?, ?)`
        ).bind(snapshot.branch_id, snapshot.activity_type, snapshot.period, snapshot.activity_date, snapshot.status, snapshot.document_link, snapshot.notes),
      };
    },
  });
}

// ---- FOGGING ----
async function handleFogging(request, env, user, origin, path) {
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === 'GET' && path === '/export') {
    if (!hasPermission(user, 'reports', 'export')) return forbidden(origin);
    return crudList(request, env, origin, 'fogging_reports', '', [], { useExport: true });
  }
  if (request.method === 'GET' && path === '') return crudList(request, env, origin, 'fogging_reports');
  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'reports', 'import')) return forbidden(origin);
    return importFogging(request, env, user, origin);
  }
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'reports', 'write')) return forbidden(origin);
    let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
    if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
    const { branch_id, activity_type, period, activity_date, status, document_link, notes } = body;
    if (!period) return error('period required', 400, origin);
    const branchCheck = await validateBranchExists(env, branch_id);
    if (!branchCheck.valid) return error(branchCheck.error, 400, origin);
    const finalStatus = status !== null && status !== undefined && status !== '' ? status : '';
    const insertStmt = env.DB.prepare(
      'INSERT INTO fogging_reports (branch_id, activity_type, period, activity_date, status, document_link, notes) VALUES (?, ?, ?, ?, ?, ?, ?)'
    ).bind(branchCheck.branchId, activity_type || 'Fogging', period, activity_date || null, finalStatus, document_link || null, notes || null);
    return atomicCreateReport(env, user, 'reports/fogging', insertStmt, {
      branch_id: branchCheck.branchId,
      activity_type: activity_type || 'Fogging',
      period,
      activity_date: activity_date || null,
      status: finalStatus,
      document_link: document_link || null,
      notes: notes || null
    }, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') {
      const row = await env.DB.prepare('SELECT t.*, b.full_name as branch_name FROM fogging_reports t LEFT JOIN branches b ON t.branch_id = b.id WHERE t.id = ? AND t.deleted_at IS NULL').bind(id).first();
      return row ? ok(row, 200, origin) : notFound(origin);
    }
    if (request.method === 'PUT' || request.method === 'PATCH') {
      if (!hasPermission(user, 'reports', 'write')) return forbidden(origin);
      let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
      if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
      const existing = await env.DB.prepare('SELECT * FROM fogging_reports WHERE id = ? AND deleted_at IS NULL').bind(id).first();
      if (!existing) return notFound(origin);
      const { branch_id, activity_type, period, activity_date, status, document_link, notes } = body;
      const branchCheck = await validateBranchExists(env, branch_id);
      if (!branchCheck.valid) return error(branchCheck.error, 400, origin);
      const nextStatus = status !== null && status !== undefined && status !== '' ? status : null;
      const updateStmt = env.DB.prepare(
        `UPDATE fogging_reports SET branch_id = COALESCE(?, branch_id), activity_type = COALESCE(?, activity_type), period = COALESCE(?, period), activity_date = COALESCE(?, activity_date), status = COALESCE(?, status), document_link = COALESCE(?, document_link), notes = COALESCE(?, notes), updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ? AND deleted_at IS NULL`
      ).bind(branchCheck.branchId, activity_type || null, period || null, activity_date || null, nextStatus, document_link || null, notes || null, id);
      const newData = {
        ...existing,
        branch_id: branchCheck.branchId ?? existing.branch_id,
        activity_type: activity_type || existing.activity_type,
        period: period || existing.period,
        activity_date: activity_date || existing.activity_date,
        status: nextStatus ?? existing.status,
        document_link: document_link || existing.document_link,
        notes: notes || existing.notes
      };
      return atomicUpdateReport(env, user, 'reports/fogging', updateStmt, existing, newData, origin);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'reports', 'delete')) return forbidden(origin);
      return softDeleteReport(env, user, 'fogging_reports', 'reports/fogging', id, origin);
    }
  }
  return error('Not found', 404, origin);
}

async function importFogging(request, env, user, origin) {
  return executeJournaledReportImport(request, env, user, origin, {
    module: 'reports_fogging_route',
    auditModule: 'reports/fogging',
    selectSql: 'SELECT * FROM fogging_reports WHERE deleted_at IS NULL',
    getExistingKey: (s) => (s.branch_id && s.period && s.activity_date)
      ? `${s.branch_id}_${s.period.toLowerCase().trim()}_${String(s.activity_date).slice(0, 4)}`
      : null,
    getItemKey: (item) => (item.branch_id && item.period && item.activity_date)
      ? `${item.branch_id}_${String(item.period).toLowerCase().trim()}_${String(item.activity_date).slice(0, 4)}`
      : null,
    buildUpdate: (envDb, item, oldRow) => {
      const status = item.status !== null && item.status !== undefined && item.status !== '' ? item.status : '';
      const afterRow = {
        ...oldRow,
        activity_date: item.activity_date,
        status,
        document_link: item.document_link || null,
        notes: item.notes || null,
      };
      return {
        afterRow,
        stmt: envDb.DB.prepare(
          `UPDATE fogging_reports SET activity_date = ?, status = ?, document_link = ?, notes = ?, updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ? AND deleted_at IS NULL`
        ).bind(item.activity_date, status, item.document_link || null, item.notes || null, oldRow.id),
      };
    },
    buildInsert: (envDb, item) => {
      const status = item.status !== null && item.status !== undefined && item.status !== '' ? item.status : '';
      const snapshot = {
        branch_id: item.branch_id,
        activity_type: item.activity_type || 'Fogging',
        period: item.period,
        activity_date: item.activity_date,
        status,
        document_link: item.document_link || null,
        notes: item.notes || null,
      };
      return {
        snapshot,
        stmt: envDb.DB.prepare(
          `INSERT INTO fogging_reports (branch_id, activity_type, period, activity_date, status, document_link, notes) 
           VALUES (?, ?, ?, ?, ?, ?, ?)`
        ).bind(snapshot.branch_id, snapshot.activity_type, snapshot.period, snapshot.activity_date, snapshot.status, snapshot.document_link, snapshot.notes),
      };
    },
  });
}

// ---- BASECAMP ----
async function listBasecamp(request, env, origin, { useExport = false } = {}) {
  const { page, limit, offset } = useExport ? getExportPagination(request.url) : getPagination(request.url);
  const url = new URL(request.url);
  const branch_id = url.searchParams.get('branch_id') || '';
  const status = url.searchParams.get('status') || '';
  const pic = url.searchParams.get('pic') || '';
  const period = url.searchParams.get('period') || '';
  const month = url.searchParams.get('month') || '';
  const year = url.searchParams.get('year') || '';
  const search = url.searchParams.get('search') || '';
  let conditions = ['t.deleted_at IS NULL']; let values = [];
  if (branch_id) { conditions.push('t.branch_id = ?'); values.push(branch_id); }
  if (status) { conditions.push('t.status = ?'); values.push(status); }
  if (pic) { conditions.push('t.pic = ?'); values.push(pic); }
  if (period) {
    const quarterMonths = {
      Q1: ['01', '02', '03'],
      Q2: ['04', '05', '06'],
      Q3: ['07', '08', '09'],
      Q4: ['10', '11', '12']
    }[String(period).trim().toUpperCase()];
    if (quarterMonths) {
      conditions.push("strftime('%m', t.info_date) IN (?, ?, ?)");
      values.push(...quarterMonths);
    } else {
      return error('Periode tidak valid (gunakan Q1, Q2, Q3, atau Q4)', 400, origin);
    }
  }
  if (month) { conditions.push("strftime('%m', t.info_date) = ?"); values.push(month); }
  if (year) { conditions.push("strftime('%Y', t.info_date) = ?"); values.push(year); }
  if (search) { conditions.push('(t.problem LIKE ? OR t.pic LIKE ? OR b.full_name LIKE ?)'); values.push(`%${search}%`, `%${search}%`, `%${search}%`); }
  const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';
  const countJoin = search ? 'LEFT JOIN branches b ON t.branch_id = b.id' : '';
  const [countResult, rows] = await Promise.all([
    env.DB.prepare(`SELECT COUNT(*) as total FROM basecamp_reports t ${countJoin} ${where}`).bind(...values).first(),
    env.DB.prepare(`SELECT t.*, b.full_name as branch_name FROM basecamp_reports t LEFT JOIN branches b ON t.branch_id = b.id ${where} ORDER BY t.info_date DESC, t.id DESC LIMIT ? OFFSET ?`).bind(...values, limit, offset).all()
  ]);
  return paginated(rows.results || [], countResult?.total || 0, page, limit, origin, env?.ENVIRONMENT);
}

async function handleBasecamp(request, env, user, origin, path) {
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === 'GET' && path === '/export') {
    if (!hasPermission(user, 'reports', 'export')) return forbidden(origin);
    return listBasecamp(request, env, origin, { useExport: true });
  }
  if (request.method === 'GET' && path === '') {
    return listBasecamp(request, env, origin);
  }
  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'reports', 'import')) return forbidden(origin);
    return importBasecamp(request, env, user, origin);
  }
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'reports', 'write')) return forbidden(origin);
    let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
    if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
    const { branch_id, problem, pic, info_date, done_date, status, notes } = body;
    if (!problem || !info_date) return error('problem and info_date required', 400, origin);
    if (done_date && String(done_date).trim() && String(done_date).trim() < String(info_date).trim()) {
      return error('done_date tidak boleh lebih awal dari info_date', 400, origin);
    }
    const branchCheck = await validateBranchExists(env, branch_id);
    if (!branchCheck.valid) return error(branchCheck.error, 400, origin);
    const finalStatus = status !== null && status !== undefined && status !== '' ? status : '';
    const insertStmt = env.DB.prepare(
      'INSERT INTO basecamp_reports (branch_id, problem, pic, info_date, done_date, status, notes) VALUES (?, ?, ?, ?, ?, ?, ?)'
    ).bind(branchCheck.branchId, problem, pic || null, info_date, done_date || null, finalStatus, notes || null);
    return atomicCreateReport(env, user, 'reports/basecamp', insertStmt, {
      branch_id: branchCheck.branchId,
      problem,
      pic: pic || null,
      info_date,
      done_date: done_date || null,
      status: finalStatus,
      notes: notes || null
    }, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') {
      const row = await env.DB.prepare('SELECT t.*, b.full_name as branch_name FROM basecamp_reports t LEFT JOIN branches b ON t.branch_id = b.id WHERE t.id = ? AND t.deleted_at IS NULL').bind(id).first();
      return row ? ok(row, 200, origin) : notFound(origin);
    }
    if (request.method === 'PUT' || request.method === 'PATCH') {
      if (!hasPermission(user, 'reports', 'write')) return forbidden(origin);
      let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
      if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
      const existing = await env.DB.prepare('SELECT * FROM basecamp_reports WHERE id = ? AND deleted_at IS NULL').bind(id).first();
      if (!existing) return notFound(origin);
      const { branch_id, problem, pic, info_date, done_date, status, notes } = body;
      const branchCheck = await validateBranchExists(env, branch_id);
      if (!branchCheck.valid) return error(branchCheck.error, 400, origin);
      const effInfoDate = info_date || existing.info_date;
      const effDoneDate = Object.prototype.hasOwnProperty.call(body, 'done_date')
        ? (done_date && String(done_date).trim() ? String(done_date).trim() : null)
        : existing.done_date;
      if (effInfoDate && effDoneDate && effDoneDate < effInfoDate) {
        return error('done_date tidak boleh lebih awal dari info_date', 400, origin);
      }
      const nextStatus = status !== null && status !== undefined && status !== '' ? status : null;
      const updateStmt = env.DB.prepare(
        `UPDATE basecamp_reports SET branch_id = COALESCE(?, branch_id), problem = COALESCE(?, problem), pic = COALESCE(?, pic), info_date = COALESCE(?, info_date), done_date = ?, status = COALESCE(?, status), notes = COALESCE(?, notes), updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ? AND deleted_at IS NULL`
      ).bind(branchCheck.branchId, problem || null, pic || null, info_date || null, effDoneDate, nextStatus, notes || null, id);
      const newData = {
        ...existing,
        branch_id: branchCheck.branchId ?? existing.branch_id,
        problem: problem || existing.problem,
        pic: pic || existing.pic,
        info_date: effInfoDate,
        done_date: effDoneDate,
        status: nextStatus ?? existing.status,
        notes: notes || existing.notes
      };
      return atomicUpdateReport(env, user, 'reports/basecamp', updateStmt, existing, newData, origin);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'reports', 'delete')) return forbidden(origin);
      return softDeleteReport(env, user, 'basecamp_reports', 'reports/basecamp', id, origin);
    }
  }
  return error('Not found', 404, origin);
}

async function importBasecamp(request, env, user, origin) {
  return executeJournaledReportImport(request, env, user, origin, {
    module: 'reports_basecamp_route',
    auditModule: 'reports/basecamp',
    selectSql: 'SELECT * FROM basecamp_reports WHERE deleted_at IS NULL',
    getExistingKey: (s) => (s.branch_id && s.info_date && s.problem)
      ? `${s.branch_id}_${s.info_date}_${s.problem.toLowerCase().trim()}`
      : null,
    getItemKey: (item) => (item.branch_id && item.info_date && item.problem)
      ? `${item.branch_id}_${item.info_date}_${String(item.problem).toLowerCase().trim()}`
      : null,
    buildUpdate: (envDb, item, oldRow) => {
      const status = item.status !== null && item.status !== undefined && item.status !== '' ? item.status : 'Open';
      const afterRow = {
        ...oldRow,
        pic: item.pic || null,
        done_date: item.done_date || null,
        status,
        notes: item.notes || null,
      };
      return {
        afterRow,
        stmt: envDb.DB.prepare(
          `UPDATE basecamp_reports SET pic = ?, done_date = ?, status = ?, notes = ?, updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ? AND deleted_at IS NULL`
        ).bind(item.pic || null, item.done_date || null, status, item.notes || null, oldRow.id),
      };
    },
    buildInsert: (envDb, item) => {
      const status = item.status !== null && item.status !== undefined && item.status !== '' ? item.status : 'Open';
      const snapshot = {
        branch_id: item.branch_id,
        problem: item.problem,
        pic: item.pic || null,
        info_date: item.info_date,
        done_date: item.done_date || null,
        status,
        notes: item.notes || null,
      };
      return {
        snapshot,
        stmt: envDb.DB.prepare(
          `INSERT INTO basecamp_reports (branch_id, problem, pic, info_date, done_date, status, notes) 
           VALUES (?, ?, ?, ?, ?, ?, ?)`
        ).bind(snapshot.branch_id, snapshot.problem, snapshot.pic, snapshot.info_date, snapshot.done_date, snapshot.status, snapshot.notes),
      };
    },
  });
}

// ---- SUPPLY REQUESTS (public POST, auth required for GET/PUT/DELETE) ----
const supplyRateLimits = new Map();
const SUPPLY_RATE_WINDOW_MS = 5 * 60 * 1000; // 5 minutes cooldown window
const SUPPLY_RATE_MAX_REQUESTS = 10; // Max 10 submissions per 5 minutes per IP

function getClientIp(request, env) {
  const cfIp = request.headers.get('CF-Connecting-IP');
  if (cfIp) return cfIp;
  const isDeployedEnv = env?.ENVIRONMENT === 'production' || env?.ENVIRONMENT === 'staging';
  if (isDeployedEnv) return 'untrusted-edge';
  return request.headers.get('x-real-ip')
    || request.headers.get('x-forwarded-for')?.split(',')[0]?.trim()
    || '127.0.0.1';
}

function isSupplyRateLimited(ip) {
  const now = Date.now();
  if (supplyRateLimits.size > 2000) {
    for (const [key, timestamps] of supplyRateLimits.entries()) {
      const active = timestamps.filter(t => now - t < SUPPLY_RATE_WINDOW_MS);
      if (active.length === 0) supplyRateLimits.delete(key);
      else supplyRateLimits.set(key, active);
    }
  }
  const timestamps = supplyRateLimits.get(ip) || [];
  const active = timestamps.filter(t => now - t < SUPPLY_RATE_WINDOW_MS);
  if (active.length >= SUPPLY_RATE_MAX_REQUESTS) {
    return true;
  }
  active.push(now);
  supplyRateLimits.set(ip, active);
  return false;
}

export function _resetSupplyRateLimitForTest() {
  supplyRateLimits.clear();
}

function parseAndSanitizeItemList(raw, maxItems = 50, maxItemLen = 100) {
  if (raw === undefined || raw === null || raw === '') return [];
  let items = [];
  if (Array.isArray(raw)) {
    items = raw;
  } else if (typeof raw === 'string') {
    const trimmed = raw.trim();
    if (!trimmed) return [];
    try {
      const parsed = JSON.parse(trimmed);
      if (Array.isArray(parsed)) {
        items = parsed;
      } else {
        items = trimmed.split(',');
      }
    } catch {
      items = trimmed.split(',');
    }
  } else {
    throw new Error('Format daftar item tidak valid');
  }

  if (items.length > maxItems) {
    throw new Error(`Jumlah item melebihi batas maksimal (${maxItems} item)`);
  }

  const cleaned = [];
  for (const item of items) {
    if (item === null || item === undefined) continue;
    const str = String(item).trim();
    if (!str) continue;
    if (str.length > maxItemLen) {
      throw new Error(`Nama item melebihi batas ${maxItemLen} karakter`);
    }
    cleaned.push(str);
  }
  return cleaned;
}

async function listSupply(request, env, origin, { useExport = false } = {}) {
  const { page, limit, offset } = useExport ? getExportPagination(request.url) : getPagination(request.url);
  const url = new URL(request.url);
  const status = url.searchParams.get('status') || '';
  const branch_id = url.searchParams.get('branch_id') || '';
  const search = url.searchParams.get('search') || '';
  const conditions = ['r.deleted_at IS NULL'];
  const values = [];
  if (status) { conditions.push('r.status = ?'); values.push(status); }
  if (branch_id) { conditions.push('r.branch_id = ?'); values.push(branch_id); }
  if (search) {
    conditions.push('(r.submitter_name LIKE ? OR r.branch_name LIKE ? OR b.full_name LIKE ? OR r.tools_items LIKE ? OR r.chemical_items LIKE ?)');
    values.push(`%${search}%`, `%${search}%`, `%${search}%`, `%${search}%`, `%${search}%`);
  }
  const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';
  const countJoin = search ? 'LEFT JOIN branches b ON r.branch_id = b.id' : '';
  const [countResult, rows] = await Promise.all([
    env.DB.prepare(`SELECT COUNT(*) as total FROM supply_requests r ${countJoin} ${where}`).bind(...values).first(),
    env.DB.prepare(`SELECT r.*, b.full_name as branch_name_ref FROM supply_requests r LEFT JOIN branches b ON r.branch_id = b.id ${where} ORDER BY r.submitted_at DESC, r.id DESC LIMIT ? OFFSET ?`).bind(...values, limit, offset).all()
  ]);
  return paginated(rows.results || [], countResult?.total || 0, page, limit, origin, env?.ENVIRONMENT);
}

async function handleSupply(request, env, origin, path) {
  const idMatch = path.match(/^\/(\d+)$/);

  // Public: POST /api/reports/supply (Authenticated callers must hold reports:write permission)
  if (request.method === 'POST' && path === '') {
    let supplyCaller = null;
    if (request.headers.get('Authorization')) {
      const caller = await authenticate(request, env);
      if (!caller) return unauthorized(origin);
      if (!hasPermission(caller, 'reports', 'write')) return forbidden(origin);
      supplyCaller = caller;
    }
    // 1. In-isolate IP Cooldown / Rate Limiting (Evaluated before parsing large body or touching D1; global multi-isolate enforcement handled by Cloudflare WAF)
    const clientIp = getClientIp(request, env);
    if (isSupplyRateLimited(clientIp)) {
      return error('Terlalu banyak permintaan pengajuan. Silakan coba lagi beberapa saat lagi.', 429, origin);
    }

    // 2. Request body size check (max 64KB)
    const contentLength = request.headers.get('content-length');
    if (contentLength && parseInt(contentLength, 10) > 65536) {
      return error('Ukuran permintaan terlalu besar (maksimal 64KB)', 413, origin);
    }

    // 3. JSON parsing
    let body;
    try {
      body = await request.json();
    } catch {
      return error('Format JSON tidak valid', 400, origin);
    }

    if (!body || typeof body !== 'object' || Array.isArray(body)) {
      return error('Payload body harus berupa objek JSON', 400, origin);
    }

    // 4. Strict Field Validation & Sanitization
    // Submitter name: required, non-empty string, max 100 characters
    const rawSubmitter = body.submitter_name;
    if (!rawSubmitter || typeof rawSubmitter !== 'string' || !rawSubmitter.trim()) {
      return error('Nama pengirim wajib diisi', 400, origin);
    }
    const submitter_name = rawSubmitter.trim();
    if (submitter_name.length > 100) {
      return error('Nama pengirim maksimal 100 karakter', 400, origin);
    }

    // Branch ID: optional, must be positive integer and exist in branches if provided
    let branch_id = null;
    if (body.branch_id !== undefined && body.branch_id !== null && body.branch_id !== '') {
      const branchCheck = await validateBranchExists(env, body.branch_id);
      if (!branchCheck.valid) {
        return error(branchCheck.error, 400, origin);
      }
      branch_id = branchCheck.branchId;
    }

    // Branch name: optional string, max 100 characters
    let branch_name = null;
    if (body.branch_name !== undefined && body.branch_name !== null && body.branch_name !== '') {
      if (typeof body.branch_name !== 'string') {
        return error('Nama Cabang harus berupa string', 400, origin);
      }
      const cleanBranch = body.branch_name.trim();
      if (cleanBranch.length > 100) {
        return error('Nama Cabang maksimal 100 karakter', 400, origin);
      }
      branch_name = cleanBranch || null;
    }

    // Tools & Chemical Items: sanitized list, max 50 items each, max 100 chars per item
    let tools_items = [];
    let chemical_items = [];
    try {
      tools_items = parseAndSanitizeItemList(body.tools_items);
    } catch (e) {
      return error(e.message, 400, origin);
    }
    try {
      chemical_items = parseAndSanitizeItemList(body.chemical_items);
    } catch (e) {
      return error(e.message, 400, origin);
    }

    // Quantities: optional, max 50 characters
    let tools_quantity = null;
    if (body.tools_quantity !== undefined && body.tools_quantity !== null && body.tools_quantity !== '') {
      const qStr = String(body.tools_quantity).trim();
      if (qStr.length > 50) return error('Jumlah alat maksimal 50 karakter', 400, origin);
      tools_quantity = qStr || null;
    }

    let chemical_quantity = null;
    if (body.chemical_quantity !== undefined && body.chemical_quantity !== null && body.chemical_quantity !== '') {
      const qStr = String(body.chemical_quantity).trim();
      if (qStr.length > 50) return error('Jumlah chemical maksimal 50 karakter', 400, origin);
      chemical_quantity = qStr || null;
    }

    // Additional notes: optional, max 1000 characters
    let additional_notes = null;
    if (body.additional_notes !== undefined && body.additional_notes !== null && body.additional_notes !== '') {
      if (typeof body.additional_notes !== 'string') {
        return error('Catatan harus berupa teks', 400, origin);
      }
      const nStr = body.additional_notes.trim();
      if (nStr.length > 1000) return error('Catatan maksimal 1000 karakter', 400, origin);
      additional_notes = nStr || null;
    }

    // 5. Atomic D1 insertion + audit log
    const toolsJson = JSON.stringify(tools_items);
    const chemicalJson = JSON.stringify(chemical_items);
    const auditPrincipal = supplyCaller || { id: null, username: submitter_name, role: 'public' };
    const insertStmt = env.DB.prepare(
      'INSERT INTO supply_requests (submitter_name, branch_id, branch_name, tools_items, tools_quantity, chemical_items, chemical_quantity, additional_notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?)'
    ).bind(
      submitter_name,
      branch_id,
      branch_name,
      toolsJson,
      tools_quantity,
      chemicalJson,
      chemical_quantity,
      additional_notes
    );
    const auditStmt = buildCreateAuditStatement(env, auditPrincipal, 'reports/supply', 'new', {
      submitter_name,
      branch_id,
      branch_name,
      tools_items,
      tools_quantity,
      chemical_items,
      chemical_quantity,
      additional_notes
    });
    const batchRes = await env.DB.batch([insertStmt, auditStmt]);

    return ok(
      { id: batchRes?.[0]?.meta?.last_row_id, message: 'Request submitted successfully' },
      201,
      origin,
      undefined,
      { 'X-Audit-Batched': 'true' }
    );
  }

  // Auth required for other operations
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'reports', 'read')) return forbidden(origin);

  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'reports', 'import')) return forbidden(origin);
    return importSupply(request, env, user, origin);
  }

  if (request.method === 'GET' && path === '/export') {
    if (!hasPermission(user, 'reports', 'export')) return forbidden(origin);
    return listSupply(request, env, origin, { useExport: true });
  }

  if (request.method === 'GET' && path === '') {
    return listSupply(request, env, origin);
  }

  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') {
      const row = await env.DB.prepare('SELECT r.*, b.full_name as branch_name_ref FROM supply_requests r LEFT JOIN branches b ON r.branch_id = b.id WHERE r.id = ? AND r.deleted_at IS NULL').bind(id).first();
      return row ? ok(row, 200, origin) : notFound(origin);
    }
    if (request.method === 'PUT' || request.method === 'PATCH') {
      if (!hasPermission(user, 'reports', 'write')) return forbidden(origin);
      let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
      if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
      const existing = await env.DB.prepare('SELECT * FROM supply_requests WHERE id = ? AND deleted_at IS NULL').bind(id).first();
      if (!existing) return notFound(origin);
      const { status, processed_by, submitter_name, branch_id, branch_name, tools_items, tools_quantity, chemical_items, chemical_quantity, additional_notes } = body;
      const branchCheck = await validateBranchExists(env, branch_id);
      if (!branchCheck.valid) return error(branchCheck.error, 400, origin);
      
      let t_items = null;
      if (tools_items !== undefined && tools_items !== null) {
        try { t_items = parseAndSanitizeItemList(tools_items); } catch (e) { return error(e.message, 400, origin); }
      }
      let c_items = null;
      if (chemical_items !== undefined && chemical_items !== null) {
        try { c_items = parseAndSanitizeItemList(chemical_items); } catch (e) { return error(e.message, 400, origin); }
      }

      const finalStatus = status !== null && status !== undefined && status !== '' ? status : '';
      const tItemsStr = t_items ? JSON.stringify(t_items) : null;
      const cItemsStr = c_items ? JSON.stringify(c_items) : null;
      const updateStmt = env.DB.prepare(
        `UPDATE supply_requests SET status = COALESCE(?, status), processed_by = COALESCE(?, processed_by), submitter_name = COALESCE(?, submitter_name), branch_id = COALESCE(?, branch_id), branch_name = COALESCE(?, branch_name), tools_items = COALESCE(?, tools_items), tools_quantity = COALESCE(?, tools_quantity), chemical_items = COALESCE(?, chemical_items), chemical_quantity = COALESCE(?, chemical_quantity), additional_notes = COALESCE(?, additional_notes), updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ? AND deleted_at IS NULL`
      ).bind(
        finalStatus,
        processed_by || null,
        submitter_name || null,
        branch_id || null,
        branch_name || null,
        tItemsStr, tools_quantity || null,
        cItemsStr, chemical_quantity || null,
        additional_notes || null,
        id
      );
      const newData = {
        ...existing,
        status: finalStatus,
        processed_by: processed_by || existing.processed_by,
        submitter_name: submitter_name || existing.submitter_name,
        branch_id: branch_id || existing.branch_id,
        branch_name: branch_name || existing.branch_name,
        tools_items: tItemsStr || existing.tools_items,
        tools_quantity: tools_quantity || existing.tools_quantity,
        chemical_items: cItemsStr || existing.chemical_items,
        chemical_quantity: chemical_quantity || existing.chemical_quantity,
        additional_notes: additional_notes || existing.additional_notes
      };
      return atomicUpdateReport(env, user, 'reports/supply', updateStmt, existing, newData, origin);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'reports', 'delete')) return forbidden(origin);
      return softDeleteReport(env, user, 'supply_requests', 'reports/supply', id, origin);
    }
  }
  return error('Not found', 404, origin);
}

async function importSupply(request, env, user, origin) {
  return executeJournaledReportImport(request, env, user, origin, {
    module: 'reports_supply_route',
    auditModule: 'reports/supply',
    selectSql: 'SELECT *, date(submitted_at) as submitted_date FROM supply_requests WHERE deleted_at IS NULL',
    getExistingKey: (s) => (s.branch_id && s.submitted_date && s.submitter_name)
      ? `${s.branch_id}_${s.submitted_date}_${s.submitter_name.toLowerCase().trim()}`
      : null,
    getItemKey: (item) => (item.branch_id && item.submitted_at && item.submitter_name)
      ? `${item.branch_id}_${item.submitted_at}_${String(item.submitter_name).toLowerCase().trim()}`
      : null,
    buildUpdate: (envDb, item, oldRow) => {
      const status = item.status !== null && item.status !== undefined && item.status !== '' ? item.status : 'Pending';
      const t_items = item.tools_items ? item.tools_items.split(',').map(s => s.trim()) : [];
      const c_items = item.chemical_items ? item.chemical_items.split(',').map(s => s.trim()) : [];
      const afterRow = {
        ...oldRow,
        tools_items: JSON.stringify(t_items),
        chemical_items: JSON.stringify(c_items),
        additional_notes: item.additional_notes || null,
        status,
        processed_by: item.processed_by || null,
      };
      return {
        afterRow,
        stmt: envDb.DB.prepare(
          `UPDATE supply_requests SET tools_items = ?, chemical_items = ?, additional_notes = ?, status = ?, processed_by = ?, updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ? AND deleted_at IS NULL`
        ).bind(afterRow.tools_items, afterRow.chemical_items, afterRow.additional_notes, status, afterRow.processed_by, oldRow.id),
      };
    },
    buildInsert: (envDb, item) => {
      const status = item.status !== null && item.status !== undefined && item.status !== '' ? item.status : 'Pending';
      const t_items = item.tools_items ? item.tools_items.split(',').map(s => s.trim()) : [];
      const c_items = item.chemical_items ? item.chemical_items.split(',').map(s => s.trim()) : [];
      const snapshot = {
        submitted_at: item.submitted_at + ' 12:00:00',
        submitter_name: item.submitter_name,
        branch_id: item.branch_id,
        tools_items: JSON.stringify(t_items),
        chemical_items: JSON.stringify(c_items),
        additional_notes: item.additional_notes || null,
        status,
        processed_by: item.processed_by || null,
      };
      return {
        snapshot,
        stmt: envDb.DB.prepare(
          `INSERT INTO supply_requests (submitted_at, submitter_name, branch_id, tools_items, chemical_items, additional_notes, status, processed_by) 
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(snapshot.submitted_at, snapshot.submitter_name, snapshot.branch_id, snapshot.tools_items, snapshot.chemical_items, snapshot.additional_notes, snapshot.status, snapshot.processed_by),
      };
    },
  });
}
