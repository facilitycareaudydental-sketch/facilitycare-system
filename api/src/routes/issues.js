import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination, resolveSortOrder } from '../utils/pagination.js';
import { requireExportAccess, getExportPagination } from '../utils/export.js';
import { buildOutboxQuery, buildConditionalOutboxQuery, buildOutboxFromLastInsertQuery } from '../utils/sync_engine.js';
import { mapDBToPayload } from '../utils/sync_mapper.js';
import { buildAuditStatement, buildCreateAuditStatement, sanitizeAuditData } from '../utils/audit.js';
import { executeJournaledImport, extractImportRequestMeta } from '../utils/import_journal.js';

const ISSUE_SORT_MAP = Object.freeze({
  date_desc: 'i.report_date DESC, i.id DESC',
  date_asc: 'i.report_date ASC, i.id ASC',
  newest: 'i.id DESC',
});

export async function handleIssues(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace('/api/issues', '');

  if (request.method === 'GET' && path === '/export') {
    const { response } = await requireExportAccess(request, env, origin, 'issues');
    if (response) return response;
    return exportIssues(request, env, origin);
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'issues', 'read')) return forbidden(origin);

  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'GET' && path === '') return listIssues(request, env, origin);
  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'issues', 'import')) return forbidden(origin);
    return importIssues(request, env, origin, user);
  }
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'issues', 'create')) return forbidden(origin);
    return createIssue(request, env, origin, user);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') return getIssue(id, env, origin);
    if (request.method === 'PUT' || request.method === 'PATCH') {
      if (!hasPermission(user, 'issues', 'update')) return forbidden(origin);
      return updateIssue(id, request, env, origin, user);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'issues', 'delete')) return forbidden(origin);
      return deleteIssue(id, env, origin, user);
    }
  }
  return error('Not found', 404, origin);
}

function buildIssuesFilter(url) {
  const branch_id = (url.searchParams.get('branch_id') || '').trim();
  const category = (url.searchParams.get('category') || '').trim();
  const status = (url.searchParams.get('status') || '').trim();
  const search = (url.searchParams.get('search') || '').trim();
  const year = (url.searchParams.get('year') || '').trim();
  const month = (url.searchParams.get('month') || '').trim();
  const sort = url.searchParams.get('sort') || 'date_desc';

  const conditions = ['i.deleted_at IS NULL'];
  const values = [];
  if (branch_id) { conditions.push('i.branch_id = ?'); values.push(branch_id); }
  if (category) { conditions.push('i.category = ?'); values.push(category); }
  if (status === 'unresolved') {
    conditions.push("i.status != 'Done'");
  } else if (status) {
    conditions.push('i.status = ?');
    values.push(status);
  }
  if (search) {
    conditions.push('(i.complaint LIKE ? OR i.employee_name LIKE ? OR i.source LIKE ? OR i.category LIKE ? OR i.solution LIKE ? OR b.full_name LIKE ?)');
    values.push(`%${search}%`, `%${search}%`, `%${search}%`, `%${search}%`, `%${search}%`, `%${search}%`);
  }
  if (year) { conditions.push('i.report_date >= ? AND i.report_date <= ?'); values.push(`${year}-01-01`, `${year}-12-31`); }
  if (month) { conditions.push('i.report_date >= ? AND i.report_date <= ?'); values.push(`${month}-01`, `${month}-31`); }

  const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';
  const countJoin = search ? 'LEFT JOIN branches b ON i.branch_id = b.id' : '';
  const orderBy = resolveSortOrder(sort, ISSUE_SORT_MAP, 'date_desc');
  return { where, values, countJoin, orderBy };
}

async function listIssues(request, env, origin) {
  const url = new URL(request.url);
  const { page, limit, offset } = getPagination(request.url);
  const { where, values, countJoin, orderBy } = buildIssuesFilter(url);
  const includeStats = url.searchParams.get('include_stats') === '1';

  const queries = [
    env.DB.prepare(`SELECT COUNT(*) as total FROM issues i ${countJoin} ${where}`).bind(...values).first(),
    env.DB.prepare(
      `SELECT i.*, b.full_name as branch_name FROM issues i
       LEFT JOIN branches b ON i.branch_id = b.id
       ${where} ORDER BY ${orderBy} LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all(),
  ];

  if (includeStats) {
    queries.push(
      env.DB.prepare(
        `SELECT
           COUNT(*) as total,
           SUM(CASE WHEN status = 'Open' THEN 1 ELSE 0 END) as open_count,
           SUM(CASE WHEN status = 'In Progress' THEN 1 ELSE 0 END) as in_progress_count,
           SUM(CASE WHEN status = 'Done' THEN 1 ELSE 0 END) as done_count,
           SUM(CASE WHEN strftime('%Y-%m', report_date) = strftime('%Y-%m', date('now', '+7 hours')) THEN 1 ELSE 0 END) as current_month
         FROM issues
         WHERE deleted_at IS NULL`
      ).first()
    );
  }

  const [countResult, rows, statsRow] = await Promise.all(queries);
  const open = Number(statsRow?.open_count || 0);
  const inProgress = Number(statsRow?.in_progress_count || 0);
  const extra = includeStats ? {
    stats: {
      total: Number(statsRow?.total || 0),
      open,
      inProgress,
      openAndProgress: open + inProgress,
      done: Number(statsRow?.done_count || 0),
      currentMonth: Number(statsRow?.current_month || 0),
    }
  } : null;

  return paginated(rows.results || [], countResult?.total || 0, page, limit, origin, env?.ENVIRONMENT, extra);
}

async function exportIssues(request, env, origin) {
  const url = new URL(request.url);
  const { page, limit, offset } = getExportPagination(request.url);
  const { where, values, countJoin, orderBy } = buildIssuesFilter(url);

  const [countResult, rows] = await Promise.all([
    env.DB.prepare(`SELECT COUNT(*) as total FROM issues i ${countJoin} ${where}`).bind(...values).first(),
    env.DB.prepare(
      `SELECT i.*, b.full_name as branch_name FROM issues i
       LEFT JOIN branches b ON i.branch_id = b.id
       ${where} ORDER BY ${orderBy} LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all()
  ]);

  return paginated(rows.results || [], countResult?.total || 0, page, limit, origin, env?.ENVIRONMENT);
}

async function getIssue(id, env, origin) {
  const row = await env.DB.prepare(
    'SELECT i.*, b.full_name as branch_name FROM issues i LEFT JOIN branches b ON i.branch_id = b.id WHERE i.id = ? AND i.deleted_at IS NULL'
  ).bind(id).first();
  if (!row) return notFound(origin);
  return ok(row, 200, origin);
}

export const OFFICIAL_ISSUE_SOURCES = Object.freeze([
  'SPV',
  'AM',
  'RCP',
  'Perawat',
  'FC',
  'Berlin',
  'Ade',
  'Pattrel',
  'Dentrel',
]);

async function validateIssueBranch(env, branch_id) {
  if (branch_id === undefined || branch_id === null || branch_id === '') return { valid: true, branchId: null };
  const parsed = Number(branch_id);
  if (!Number.isInteger(parsed) || parsed <= 0) return { valid: false, error: 'branch_id tidak valid' };
  const row = await env.DB.prepare('SELECT id FROM branches WHERE id = ? AND deleted_at IS NULL').bind(parsed).first();
  if (!row) return { valid: false, error: 'Cabang (branch_id) tidak ditemukan atau sudah dinonaktifkan' };
  return { valid: true, branchId: parsed };
}

async function createIssue(request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
  const { report_date, branch_id, category, source, complaint, employee_name, fc_specialist, solution, status, completion_date } = body;
  if (!report_date || !complaint || !category) return error('report_date, category, complaint required', 400, origin);

  const cleanSource = source !== undefined && source !== null ? String(source).trim() : '';
  if (cleanSource && !OFFICIAL_ISSUE_SOURCES.includes(cleanSource)) {
    return error(`Sumber laporan tidak valid. Pilihan resmi: ${OFFICIAL_ISSUE_SOURCES.join(', ')}`, 400, origin);
  }

  const branchCheck = await validateIssueBranch(env, branch_id);
  if (!branchCheck.valid) return error(branchCheck.error, 400, origin);

  const cleanCompletionDate = (completion_date && typeof completion_date === 'string' && completion_date.trim()) ? completion_date.trim() : null;
  let day_count = null;
  if (cleanCompletionDate && report_date) {
    if (cleanCompletionDate < String(report_date).trim()) {
      return error('completion_date tidak boleh lebih awal dari report_date', 400, origin);
    }
    const d1 = new Date(report_date);
    const d2 = new Date(cleanCompletionDate);
    day_count = Math.max(0, Math.floor((d2 - d1) / (1000 * 60 * 60 * 24)));
  }

  const normalizedStatus = status !== null && status !== undefined && status !== '' ? status : '';
  const payload = mapDBToPayload('Permasalahan', {
    id: '',
    report_date,
    branch_id: branch_id || null,
    category,
    source: cleanSource || null,
    complaint,
    employee_name: employee_name || null,
    fc_specialist: fc_specialist || null,
    solution: solution || null,
    status: normalizedStatus,
    completion_date: completion_date || null,
    day_count,
    last_sync_source: 'FCMS',
    row_version: 1
  });

  const corrId = crypto.randomUUID();
  const auditSnapshot = JSON.stringify(sanitizeAuditData({
    report_date,
    branch_id: branch_id || null,
    category,
    source: cleanSource || null,
    complaint,
    employee_name: employee_name || null,
    fc_specialist: fc_specialist || null,
    solution: solution || null,
    status: normalizedStatus,
    completion_date: completion_date || null,
    day_count
  }));
  const createAuditStmt = env.DB.prepare(
    `INSERT INTO audit_logs (user_id, user_name, user_role, action, module, target_id, old_data, new_data, old_value, new_value, details)
     SELECT ?, ?, ?, 'CREATE', 'issues', COALESCE((SELECT entity_id FROM sync_outbox WHERE correlation_id = ? LIMIT 1), 'new'), NULL, ?, NULL, ?, ?
     WHERE (SELECT changes()) > 0`
  ).bind(
    (user?.id !== undefined && user?.id !== null && user?.id !== 0) ? user.id : null,
    user?.full_name || user?.username || 'System Automation',
    user?.role || 'system',
    corrId,
    auditSnapshot,
    auditSnapshot,
    JSON.stringify({ status: 'SUCCESS', correlation_id: corrId })
  );

  const batchRes = await env.DB.batch([
    env.DB.prepare(
      "INSERT INTO issues (report_date, branch_id, category, source, complaint, employee_name, fc_specialist, solution, status, completion_date, day_count, row_version, last_sync_source) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, 'FCMS')"
    ).bind(report_date, branch_id || null, category, cleanSource || null, complaint, employee_name || null,
      fc_specialist || null, solution || null, normalizedStatus, completion_date || null, day_count),
    buildOutboxFromLastInsertQuery(env, 'Permasalahan', payload, corrId),
    createAuditStmt
  ]);

  const newId = batchRes?.[0]?.meta?.last_row_id ?? 1;

  return ok({ id: newId }, 201, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function updateIssue(id, request, env, origin, user = null) {
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

  const existing = await env.DB.prepare('SELECT * FROM issues WHERE id = ? AND deleted_at IS NULL').bind(id).first();
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

  const { report_date, branch_id, category, source, complaint, employee_name, fc_specialist, solution, status, completion_date } = body;

  const cleanUpdateSource = source !== undefined && source !== null ? String(source).trim() : null;
  if (
    cleanUpdateSource &&
    cleanUpdateSource !== String(existing.source || '').trim() &&
    !OFFICIAL_ISSUE_SOURCES.includes(cleanUpdateSource)
  ) {
    return error(`Sumber laporan tidak valid. Pilihan resmi: ${OFFICIAL_ISSUE_SOURCES.join(', ')}`, 400, origin);
  }

  const branchCheck = await validateIssueBranch(env, branch_id);
  if (!branchCheck.valid) return error(branchCheck.error, 400, origin);

  const rd = report_date || existing.report_date;
  const hasCompletionDateKey = Object.prototype.hasOwnProperty.call(body, 'completion_date');
  const cd = hasCompletionDateKey
    ? (completion_date && typeof completion_date === 'string' && completion_date.trim() ? completion_date.trim() : null)
    : existing.completion_date;
  let day_count = existing.day_count;
  if (rd && cd) {
    if (cd < String(rd).trim()) {
      return error('completion_date tidak boleh lebih awal dari report_date', 400, origin);
    }
    day_count = Math.max(0, Math.floor((new Date(cd) - new Date(rd)) / (1000 * 60 * 60 * 24)));
  } else if (hasCompletionDateKey && !cd) {
    day_count = null;
  }

  const nextSource = Object.prototype.hasOwnProperty.call(body, 'source') ? (cleanUpdateSource || null) : existing.source;
  const nextEmployeeName = Object.prototype.hasOwnProperty.call(body, 'employee_name') ? (employee_name || null) : existing.employee_name;
  const nextFcSpecialist = Object.prototype.hasOwnProperty.call(body, 'fc_specialist') ? (fc_specialist || null) : existing.fc_specialist;
  const nextSolution = Object.prototype.hasOwnProperty.call(body, 'solution') ? (solution || null) : existing.solution;
  const nextStatus = status !== null && status !== undefined && status !== '' ? status : existing.status;

  const baseVersion = isExistingNull ? clientVersion : existing.row_version;
  const nextRowVersion = baseVersion + 1;
  const updatedRecord = {
    ...existing,
    report_date: rd,
    branch_id: branch_id !== undefined ? (branchCheck.branchId) : existing.branch_id,
    category: category || existing.category,
    source: nextSource,
    complaint: complaint || existing.complaint,
    employee_name: nextEmployeeName,
    fc_specialist: nextFcSpecialist,
    solution: nextSolution,
    status: nextStatus,
    completion_date: cd,
    day_count,
    row_version: nextRowVersion
  };
  
  const updateStmt = env.DB.prepare(
    `UPDATE issues SET report_date = ?, branch_id = ?,
     category = ?, source = ?, complaint = ?,
     employee_name = ?, fc_specialist = ?,
     solution = ?, status = ?,
     completion_date = ?, day_count = ?, updated_at = datetime('now'), row_version = COALESCE(row_version, ?) + 1, last_sync_source = 'FCMS'
     WHERE id = ? AND deleted_at IS NULL AND (row_version = ? OR (row_version IS NULL AND (? = 1 OR ? = 0)))`
  ).bind(
    updatedRecord.report_date,
    updatedRecord.branch_id,
    updatedRecord.category,
    updatedRecord.source,
    updatedRecord.complaint,
    updatedRecord.employee_name,
    updatedRecord.fc_specialist,
    updatedRecord.solution,
    updatedRecord.status,
    updatedRecord.completion_date,
    day_count,
    clientVersion,
    id,
    clientVersion,
    clientVersion,
    clientVersion
  );

  const batchStmts = [
    updateStmt,
    buildConditionalOutboxQuery(env, 'Permasalahan', id, 'UPDATE', mapDBToPayload('Permasalahan', { ...updatedRecord, last_sync_source: 'FCMS' })),
    buildAuditStatement(env, user, 'UPDATE', 'issues', id, existing, updatedRecord, null, 'SUCCESS', true)
  ];
  const batchRes = await env.DB.batch(batchStmts);

  const changes = batchRes?.[0]?.meta?.changes ?? batchRes?.[0]?.changes;
  if (changes === 0) {
    return error('Conflict: Record has been modified by another user or session', 409, origin);
  }

  return ok(
    { message: 'Issue updated', row_version: nextRowVersion },
    200,
    origin,
    undefined,
    { 'X-Audit-Batched': 'true' }
  );
}

async function deleteIssue(id, env, origin, user = null) {
  const existing = await env.DB.prepare('SELECT * FROM issues WHERE id = ? AND deleted_at IS NULL').bind(id).first();
  if (!existing) return notFound(origin);
  
  const deletedRecord = { ...existing, deleted_at: new Date().toISOString() };
  
  const { buildAuditStatement } = await import('../utils/audit.js');
  const auditStmt = buildAuditStatement(env, user, 'DELETE', 'issues', id, existing, null, null, 'SUCCESS');

  await env.DB.batch([
    env.DB.prepare(`UPDATE issues SET deleted_at = datetime('now'), updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ?`).bind(id),
    buildOutboxQuery(env, 'Permasalahan', id, 'DELETE', mapDBToPayload('Permasalahan', deletedRecord)),
    auditStmt
  ]);
  return ok({ message: 'Issue deleted' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function importIssues(request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
  if (body.length === 0) return ok({ message: 'No data to import' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });

  const importMeta = extractImportRequestMeta(request);
  try {
    return await executeJournaledImport(env, user, origin, {
      module: 'issues_route',
      rows: body,
      onDuplicate: 'upsert',
      importMeta,
      buildPlan: async ({ startRowIndex, initialCounts }) => {
        const existingAll = await env.DB.prepare('SELECT * FROM issues').all();
        const existingMap = new Map();
        (existingAll.results || []).forEach(r => {
          const bKey = r.branch_id === null || r.branch_id === undefined ? 'null' : String(r.branch_id);
          const key = `${bKey}|${r.report_date || ''}|${r.category || ''}|${r.complaint || ''}`;
          existingMap.set(key, r);
        });

        const seenNewKeys = new Set();
        const insertGroups = [];
        let inserted = Number(initialCounts?.inserted || 0);
        let updated = Number(initialCounts?.updated || 0);
        let skipped = Number(initialCounts?.skipped || 0);

        for (let idx = startRowIndex; idx < body.length; idx++) {
          const item = body[idx];
          if (!item || typeof item !== 'object' || !item.report_date || !item.complaint || !item.category) {
            skipped++;
            continue;
          }

          let day_count = null;
          if (item.completion_date && item.report_date) {
            const d1 = new Date(item.report_date);
            const d2 = new Date(item.completion_date);
            day_count = Math.floor((d2 - d1) / (1000 * 60 * 60 * 24));
          }

          const bKey = item.branch_id === null || item.branch_id === undefined ? 'null' : String(item.branch_id);
          const key = `${bKey}|${item.report_date}|${item.category}|${item.complaint}`;
          const existing = existingMap.get(key);
          const statusVal = item.status !== null && item.status !== undefined && item.status !== '' ? item.status : '';
          const nextFields = {
            report_date: item.report_date,
            branch_id: item.branch_id || null,
            category: item.category,
            source: item.source || null,
            complaint: item.complaint,
            employee_name: item.employee_name || null,
            fc_specialist: item.fc_specialist || null,
            solution: item.solution || null,
            status: statusVal,
            completion_date: item.completion_date || null,
            day_count,
          };

          if (existing) {
            const afterRow = { ...existing, ...nextFields };
            updated++;
            insertGroups.push({
              rowIndex: idx,
              counts: { inserted, updated, skipped },
              stmts: [
                env.DB.prepare(
                  `UPDATE issues SET source = ?, employee_name = ?, fc_specialist = ?, solution = ?, status = ?, completion_date = ?, day_count = ?, updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ?`
                ).bind(
                  nextFields.source, nextFields.employee_name, nextFields.fc_specialist, nextFields.solution,
                  nextFields.status, nextFields.completion_date, nextFields.day_count, existing.id
                ),
                buildAuditStatement(env, user, 'UPDATE', 'issues', existing.id, existing, afterRow, { source: 'import' }, 'SUCCESS', true),
              ],
            });
            existingMap.set(key, afterRow);
          } else if (seenNewKeys.has(key)) {
            skipped++;
          } else {
            seenNewKeys.add(key);
            inserted++;
            insertGroups.push({
              rowIndex: idx,
              counts: { inserted, updated, skipped },
              stmts: [
                env.DB.prepare(
                  `INSERT INTO issues (report_date, branch_id, category, source, complaint, employee_name, fc_specialist, solution, status, completion_date, day_count) 
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
                ).bind(
                  nextFields.report_date, nextFields.branch_id, nextFields.category, nextFields.source,
                  nextFields.complaint, nextFields.employee_name, nextFields.fc_specialist, nextFields.solution,
                  nextFields.status, nextFields.completion_date, nextFields.day_count
                ),
                buildCreateAuditStatement(env, user, 'issues', 'new', nextFields, { source: 'import' }),
              ],
            });
          }
        }

        return { insertGroups, finalCounts: { inserted, updated, skipped } };
      },
      buildFinalResult: (counts) => {
        const processed = Number(counts.inserted || 0) + Number(counts.updated || 0);
        return {
          message: `Berhasil memperbarui/mengimport ${processed} data permasalahan`,
          inserted: counts.inserted,
          updated: counts.updated,
          skipped: counts.skipped,
          processed,
        };
      },
    });
  } catch (err) {
    console.error('Import error:', err); return error('Gagal mengimport data', 500, origin);
  }
}
