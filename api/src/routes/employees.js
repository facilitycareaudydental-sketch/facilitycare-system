import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination, resolveSortOrder, HARD_MAX_PAGE_LIMIT } from '../utils/pagination.js';
import { requireExportAccess, getExportPagination } from '../utils/export.js';
import {
  buildAuditStatement,
  buildCreateAuditStatement,
  buildSoftDeleteRecordAuditStatement,
  executeAtomicImportBatch
} from '../utils/audit.js';
import { getManagementBranchSubquery, getManagementBranchSqlPredicate, resolveManagementBranchIds } from '../utils/branchMaster.js';

const OFFICIAL_DIVISIONS = Object.freeze(['FACILITY CARE', 'SECURITY']);

const EMPLOYEE_SORT_MAP = Object.freeze({
  name_asc: 'e.full_name ASC, e.id ASC',
  name_desc: 'e.full_name DESC, e.id DESC',
  newest: 'e.id DESC',
});

export async function handleEmployees(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace('/api/employees', '');

  if (request.method === 'GET' && path === '/export') {
    const { response } = await requireExportAccess(request, env, origin, 'employees');
    if (response) return response;
    return exportEmployees(request, env, origin);
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);

  if (request.method === 'GET' && path === '/selector') {
    if (!hasPermission(user, 'employees', 'read')) return forbidden(origin);
    return selectorEmployees(request, env, origin);
  }

  if (!hasPermission(user, 'employees', 'read')) return forbidden(origin);

  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'GET' && path === '') return listEmployees(request, env, origin);
  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'employees', 'import')) return forbidden(origin);
    return importEmployees(request, env, origin, user);
  }
  if (request.method === 'DELETE' && path === '/bulk') {
    if (!hasPermission(user, 'employees', 'delete')) return forbidden(origin);
    return bulkDeleteEmployees(request, env, origin, user);
  }
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'employees', 'write')) return forbidden(origin);
    return createEmployee(request, env, origin, user);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') return getEmployee(id, env, origin);
    if (request.method === 'PUT' || request.method === 'PATCH') {
      if (!hasPermission(user, 'employees', 'write')) return forbidden(origin);
      return updateEmployee(id, request, env, origin, user);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'employees', 'delete')) return forbidden(origin);
      return deleteEmployee(id, env, origin, user);
    }
  }
  return error('Not found', 404, origin);
}

function buildEmployeesFilter(url) {
  const search = (url.searchParams.get('search') || '').trim();
  const branch_id = (url.searchParams.get('branch_id') || '').trim();
  const division = (url.searchParams.get('division') || '').trim();
  const status = (url.searchParams.get('status') || '').trim();
  const sort = url.searchParams.get('sort') || 'name_asc';

  // Phase 5.66.10A.1: Management Employee Scope Enforcement
  const conditions = [
    'e.deleted_at IS NULL',
    "e.division IN ('FACILITY CARE', 'SECURITY')",
    getManagementBranchSubquery('e.branch_id')
  ];
  const values = [];
  if (search) {
    conditions.push('(e.full_name LIKE ? OR e.phone LIKE ? OR e.notes LIKE ? OR e.division LIKE ? OR b.full_name LIKE ?)');
    values.push(`%${search}%`, `%${search}%`, `%${search}%`, `%${search}%`, `%${search}%`);
  }
  if (branch_id) { conditions.push('e.branch_id = ?'); values.push(branch_id); }
  if (division) { conditions.push('e.division = ?'); values.push(division); }
  if (status === 'Nonaktif') {
    conditions.push("(e.status != 'Aktif' OR e.status IS NULL)");
  } else if (status) {
    conditions.push('e.status = ?');
    values.push(status);
  }

  const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';
  const countJoin = search ? 'LEFT JOIN branches b ON e.branch_id = b.id' : '';
  const orderBy = resolveSortOrder(sort, EMPLOYEE_SORT_MAP, 'name_asc');
  return { where, values, countJoin, orderBy };
}

async function selectorEmployees(request, env, origin) {
  const { page, limit, offset } = getPagination(request.url, {
    defaultLimit: HARD_MAX_PAGE_LIMIT,
    maxLimit: HARD_MAX_PAGE_LIMIT,
  });
  const url = new URL(request.url);
  const search = (url.searchParams.get('search') || '').trim();
  const branch_id = (url.searchParams.get('branch_id') || '').trim();
  const division = (url.searchParams.get('division') || '').trim();
  const status = (url.searchParams.get('status') || '').trim();

  // Phase 5.66.10A.1: Management Employee Scope Enforcement
  const conditions = [
    'e.deleted_at IS NULL',
    "e.division IN ('FACILITY CARE', 'SECURITY')",
    getManagementBranchSubquery('e.branch_id')
  ];
  const values = [];
  if (search) {
    conditions.push('(e.full_name LIKE ? OR b.full_name LIKE ? OR e.division LIKE ?)');
    values.push(`%${search}%`, `%${search}%`, `%${search}%`);
  }
  if (branch_id) { conditions.push('e.branch_id = ?'); values.push(branch_id); }
  if (division) { conditions.push('e.division = ?'); values.push(division); }
  if (status === 'Nonaktif') {
    conditions.push("(e.status != 'Aktif' OR e.status IS NULL)");
  } else if (status) {
    conditions.push('e.status = ?');
    values.push(status);
  }

  const where = 'WHERE ' + conditions.join(' AND ');
  const countJoin = search ? 'LEFT JOIN branches b ON e.branch_id = b.id' : '';

  const [countResult, rows] = await Promise.all([
    env.DB.prepare(
      `SELECT COUNT(*) as total FROM employees e ${countJoin} ${where}`
    ).bind(...values).first(),
    env.DB.prepare(
      `SELECT e.id, e.full_name, e.branch_id, b.full_name as branch_name, e.division, e.status
       FROM employees e
       LEFT JOIN branches b ON e.branch_id = b.id
       ${where}
       ORDER BY e.full_name ASC, e.id ASC
       LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all()
  ]);

  return paginated(rows.results || [], countResult?.total || 0, page, limit, origin, env?.ENVIRONMENT);
}

async function listEmployees(request, env, origin) {
  const url = new URL(request.url);
  const { page, limit, offset } = getPagination(request.url);
  const { where, values, countJoin, orderBy } = buildEmployeesFilter(url);
  const includeStats = url.searchParams.get('include_stats') === '1';

  const queries = [
    env.DB.prepare(`SELECT COUNT(*) as total FROM employees e ${countJoin} ${where}`).bind(...values).first(),
    env.DB.prepare(
      `SELECT e.*, b.full_name as branch_name
       FROM employees e
       LEFT JOIN branches b ON e.branch_id = b.id
       ${where} ORDER BY ${orderBy} LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all(),
  ];

  if (includeStats) {
    queries.push(
      env.DB.prepare(
        `SELECT
           COUNT(*) as total,
           SUM(CASE WHEN status = 'Aktif' THEN 1 ELSE 0 END) as active,
           SUM(CASE WHEN status != 'Aktif' OR status IS NULL THEN 1 ELSE 0 END) as nonactive,
           SUM(CASE WHEN division = 'FACILITY CARE' AND status = 'Aktif' THEN 1 ELSE 0 END) as fc,
           SUM(CASE WHEN division = 'SECURITY' AND status = 'Aktif' THEN 1 ELSE 0 END) as sec
         FROM employees
         WHERE deleted_at IS NULL
           AND division IN ('FACILITY CARE', 'SECURITY')
           AND ${getManagementBranchSubquery('branch_id')}`
      ).first()
    );
  }

  const [countResult, rows, statsRow] = await Promise.all(queries);
  const extra = includeStats ? {
    stats: {
      total: Number(statsRow?.total || 0),
      active: Number(statsRow?.active || 0),
      nonactive: Number(statsRow?.nonactive || 0),
      fc: Number(statsRow?.fc || 0),
      sec: Number(statsRow?.sec || 0),
    }
  } : null;

  return paginated(rows.results || [], countResult?.total || 0, page, limit, origin, env?.ENVIRONMENT, extra);
}

async function exportEmployees(request, env, origin) {
  const url = new URL(request.url);
  const { page, limit, offset } = getExportPagination(request.url);
  const { where, values, countJoin, orderBy } = buildEmployeesFilter(url);

  const [countResult, rows] = await Promise.all([
    env.DB.prepare(`SELECT COUNT(*) as total FROM employees e ${countJoin} ${where}`).bind(...values).first(),
    env.DB.prepare(
      `SELECT e.*, b.full_name as branch_name
       FROM employees e
       LEFT JOIN branches b ON e.branch_id = b.id
       ${where} ORDER BY ${orderBy} LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all()
  ]);

  return paginated(rows.results || [], countResult?.total || 0, page, limit, origin, env?.ENVIRONMENT);
}

async function getEmployee(id, env, origin) {
  // Phase 5.66.10A.1: Direct API enforcement - only Management area employee is retrievable
  const [row, contracts] = await Promise.all([
    env.DB.prepare(
      `SELECT e.*, b.full_name as branch_name 
       FROM employees e 
       JOIN branches b ON e.branch_id = b.id 
       WHERE e.id = ? AND e.deleted_at IS NULL
         AND e.division IN ('FACILITY CARE', 'SECURITY')
         AND ${getManagementBranchSqlPredicate('b')}`
    ).bind(id).first(),
    env.DB.prepare(
      `SELECT c.* FROM contracts c
       JOIN branches b ON c.branch_id = b.id
       WHERE c.employee_id = ? AND c.deleted_at IS NULL
         AND c.division IN ('FACILITY CARE', 'SECURITY')
         AND ${getManagementBranchSqlPredicate('b')}
       ORDER BY c.start_date DESC`
    ).bind(id).all()
  ]);
  if (!row) return notFound(origin);

  return ok({ ...row, contracts: contracts.results }, 200, origin);
}

async function createEmployee(request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
  const { full_name, branch_id, division, phone, join_date, status, notes, target_pindah_os, target_selesai_os } = body;
  if (!full_name || typeof full_name !== 'string' || !full_name.trim()) return error('full_name required', 400, origin);
  const cleanName = full_name.trim();

  let cleanDivision = 'FACILITY CARE';
  if (division !== undefined && division !== null && String(division).trim() !== '') {
    const divUpper = String(division).trim().toUpperCase();
    if (!OFFICIAL_DIVISIONS.includes(divUpper)) {
      return error('Divisi hanya boleh FACILITY CARE atau SECURITY', 400, origin);
    }
    cleanDivision = divUpper;
  }

  // Phase 5.66.10A.1: Verify branch is strictly within Management scope
  if (branch_id !== undefined && branch_id !== null && branch_id !== '') {
    const parsedBranchId = Number(branch_id);
    if (!Number.isInteger(parsedBranchId) || parsedBranchId <= 0) {
      return error('Invalid branch_id', 400, origin);
    }
    const branchRow = await env.DB.prepare(
      `SELECT id FROM branches WHERE id = ? AND deleted_at IS NULL AND ${getManagementBranchSqlPredicate('')}`
    ).bind(parsedBranchId).first();
    if (!branchRow) {
      return error('Pendaftaran karyawan dibatasi hanya untuk area Management (Cabang 000. Management)', 400, origin);
    }
  } else {
    return error('Area penempatan wajib diisi dan harus berupa area Management (Cabang 000. Management)', 400, origin);
  }

  const dup = await env.DB.prepare(
    'SELECT id FROM employees WHERE LOWER(TRIM(full_name)) = LOWER(TRIM(?)) AND deleted_at IS NULL'
  ).bind(cleanName).first();
  if (dup) {
    return error('Karyawan dengan nama tersebut sudah terdaftar aktif', 409, origin);
  }

  const normalizedStatus = status !== null && status !== undefined && status !== '' ? status : '';
  const newSnapshot = {
    full_name: cleanName,
    branch_id: branch_id || null,
    division: cleanDivision,
    phone: phone || null,
    join_date: join_date || null,
    status: normalizedStatus,
    notes: notes || null,
    target_pindah_os: target_pindah_os || null,
    target_selesai_os: target_selesai_os || null
  };

  const batchRes = await env.DB.batch([
    env.DB.prepare(
      'INSERT INTO employees (full_name, branch_id, division, phone, join_date, status, notes, target_pindah_os, target_selesai_os) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)'
    ).bind(cleanName, branch_id || null, cleanDivision, phone || null, join_date || null, normalizedStatus, notes || null, target_pindah_os || null, target_selesai_os || null),
    buildCreateAuditStatement(env, user, 'employees', 'new', newSnapshot)
  ]);

  return ok({ id: batchRes?.[0]?.meta?.last_row_id }, 201, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function updateEmployee(id, request, env, origin, user = null) {
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

  // Phase 5.66.10A.1: Direct API enforcement - target employee must be within Management scope
  const existing = await env.DB.prepare(
    `SELECT e.* FROM employees e 
     JOIN branches b ON e.branch_id = b.id 
     WHERE e.id = ? AND e.deleted_at IS NULL 
       AND e.division IN ('FACILITY CARE', 'SECURITY')
       AND ${getManagementBranchSqlPredicate('b')}`
  ).bind(id).first();
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

  const { full_name, branch_id, division, phone, join_date, status, notes, target_pindah_os, target_selesai_os } = body;

  let cleanDivision = undefined;
  if (division !== undefined && division !== null && String(division).trim() !== '') {
    const divUpper = String(division).trim().toUpperCase();
    if (!OFFICIAL_DIVISIONS.includes(divUpper)) {
      return error('Divisi hanya boleh FACILITY CARE atau SECURITY', 400, origin);
    }
    cleanDivision = divUpper;
  }
  if (branch_id !== undefined && branch_id !== null && branch_id !== '') {
    const parsedBranchId = Number(branch_id);
    if (!Number.isInteger(parsedBranchId) || parsedBranchId <= 0) {
      return error('Invalid branch_id', 400, origin);
    }
    const branchRow = await env.DB.prepare(
      `SELECT id FROM branches WHERE id = ? AND deleted_at IS NULL AND ${getManagementBranchSqlPredicate('')}`
    ).bind(parsedBranchId).first();
    if (!branchRow) {
      return error('Penempatan karyawan dibatasi hanya untuk area Management (Cabang 000. Management)', 400, origin);
    }
  }

  const baseVersion = isExistingNull ? clientVersion : existing.row_version;
  const nextRowVersion = baseVersion + 1;
  const afterSnapshot = {
    ...existing,
    full_name: full_name || existing.full_name,
    branch_id: branch_id || existing.branch_id,
    division: cleanDivision !== undefined ? cleanDivision : existing.division,
    phone: phone || existing.phone,
    join_date: join_date || existing.join_date,
    status: status !== undefined ? status : existing.status,
    notes: notes || existing.notes,
    target_pindah_os: target_pindah_os || existing.target_pindah_os,
    target_selesai_os: target_selesai_os || existing.target_selesai_os,
    row_version: nextRowVersion
  };

  const updateStmt = env.DB.prepare(
    `UPDATE employees SET 
      full_name = COALESCE(?, full_name),
      branch_id = COALESCE(?, branch_id),
      division = COALESCE(?, division),
      phone = COALESCE(?, phone),
      join_date = COALESCE(?, join_date),
      status = COALESCE(?, status),
      notes = COALESCE(?, notes),
      target_pindah_os = COALESCE(?, target_pindah_os),
      target_selesai_os = COALESCE(?, target_selesai_os),
      row_version = COALESCE(row_version, ?) + 1,
      updated_at = datetime('now')
     WHERE id = ? AND (row_version = ? OR (row_version IS NULL AND (? = 1 OR ? = 0)))`
  ).bind(full_name || null, branch_id || null, cleanDivision !== undefined ? cleanDivision : null, phone || null,
    join_date || null, status !== undefined ? status : null, notes || null, target_pindah_os || null, target_selesai_os || null,
    clientVersion, id, clientVersion, clientVersion, clientVersion);

  const batchRes = await env.DB.batch([
    updateStmt,
    buildAuditStatement(env, user, 'UPDATE', 'employees', id, existing, afterSnapshot, null, 'SUCCESS', true)
  ]);
  const changes = batchRes?.[0]?.meta?.changes ?? batchRes?.[0]?.changes;

  if (changes === 0) {
    return error('Conflict: Record has been modified by another user or session', 409, origin);
  }

  return ok({ message: 'Employee updated', row_version: nextRowVersion }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function deleteEmployee(id, env, origin, user = null) {
  // Phase 5.66.10A.1: Direct API enforcement - target employee must be within Management scope
  const existing = await env.DB.prepare(
    `SELECT e.* FROM employees e 
     JOIN branches b ON e.branch_id = b.id 
     WHERE e.id = ? AND e.deleted_at IS NULL 
       AND e.division IN ('FACILITY CARE', 'SECURITY')
       AND ${getManagementBranchSqlPredicate('b')}`
  ).bind(id).first();
  if (!existing) return notFound(origin);
  const deleteStmt = env.DB.prepare("UPDATE employees SET deleted_at = datetime('now'), updated_at = datetime('now') WHERE id = ?").bind(id);
  const { executeAtomicDelete } = await import('../utils/audit.js');
  await executeAtomicDelete(env, user, deleteStmt, 'employees', id, existing);
  return ok({ message: 'Employee deleted' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function importEmployees(request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
  if (body.length === 0) return ok({ message: 'No data to import' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });

  // Phase 5.66.10A.1 / 5.66.19: Fail-closed Management branch resolution & scoped existingMap
  let mgmtBranchIds;
  try {
    mgmtBranchIds = await resolveManagementBranchIds(env.DB);
  } catch (mgmtErr) {
    return error(mgmtErr.message, 503, origin);
  }

  const existingAll = await env.DB.prepare(
    `SELECT e.* FROM employees e
     WHERE e.deleted_at IS NULL AND e.division IN ('FACILITY CARE', 'SECURITY')
       AND ${getManagementBranchSubquery('e.branch_id')}`
  ).all();
  const existingMap = new Map();
  for (const row of (existingAll.results || [])) {
    if (row.full_name) existingMap.set(String(row.full_name).trim().toLowerCase(), row);
  }

  const stmts = [];
  let processed = 0;
  try {
    for (const item of body) {
      if (!item.full_name) continue;

      // Phase 5.66.10A.1: Enforce Management Scope & Official Divisions
      const itemBranchId = Number(item.branch_id);
      if (!mgmtBranchIds.has(itemBranchId)) {
        // Skip rows outside Management area (e.g. operational branch / vendor employees)
        continue;
      }

      const rawDiv = String(item.division || '').trim().toUpperCase();
      if (!OFFICIAL_DIVISIONS.includes(rawDiv)) {
        // Skip non-official divisions
        continue;
      }
      const safeDivision = rawDiv;

      const existing = existingMap.get(String(item.full_name).trim().toLowerCase());
      
      if (existing) {
        // 🔒 PERLINDUNGAN: Jika karyawan sudah "Tidak Aktif", "Resign", atau "Cut" di database,
        // import TIDAK BOLEH mengubahnya menjadi "Aktif" secara otomatis.
        const finalStatus = (existing.status === 'Tidak Aktif' || existing.status === 'Resign' || existing.status === 'Cut') 
          ? existing.status 
          : (item.status !== null && item.status !== undefined && item.status !== '' ? item.status : 'Aktif');
        const nextFields = {
          branch_id: itemBranchId,
          division: safeDivision,
          phone: item.phone || null,
          join_date: item.join_date || null,
          status: finalStatus,
          notes: item.notes || null,
          target_pindah_os: item.target_pindah_os || existing.target_pindah_os || null,
          target_selesai_os: item.target_selesai_os || existing.target_selesai_os || null
        };
        const afterRow = { ...existing, ...nextFields };

        stmts.push(
          env.DB.prepare(
            `UPDATE employees SET branch_id = ?, division = ?, phone = ?, join_date = ?, status = ?, notes = ?, target_pindah_os = COALESCE(?, target_pindah_os), target_selesai_os = COALESCE(?, target_selesai_os), updated_at = datetime('now') WHERE id = ?`
          ).bind(
            nextFields.branch_id,
            safeDivision,
            nextFields.phone,
            nextFields.join_date,
            finalStatus,
            nextFields.notes,
            item.target_pindah_os || null,
            item.target_selesai_os || null,
            existing.id
          ),
          buildAuditStatement(env, user, 'UPDATE', 'employees', existing.id, existing, afterRow, { source: 'import' }, 'SUCCESS', true)
        );
        existingMap.set(String(item.full_name).trim().toLowerCase(), afterRow);
      } else {
        const finalStatus = item.status !== null && item.status !== undefined && item.status !== '' ? item.status : 'Aktif';
        const newSnapshot = {
          full_name: item.full_name,
          branch_id: itemBranchId,
          division: safeDivision,
          phone: item.phone || null,
          join_date: item.join_date || null,
          status: finalStatus,
          notes: item.notes || null,
          target_pindah_os: item.target_pindah_os || null,
          target_selesai_os: item.target_selesai_os || null
        };
        stmts.push(
          env.DB.prepare(
            `INSERT INTO employees (full_name, branch_id, division, phone, join_date, status, notes, target_pindah_os, target_selesai_os) 
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
          ).bind(
            newSnapshot.full_name,
            newSnapshot.branch_id,
            safeDivision,
            newSnapshot.phone,
            newSnapshot.join_date,
            finalStatus,
            newSnapshot.notes,
            newSnapshot.target_pindah_os,
            newSnapshot.target_selesai_os
          ),
          buildCreateAuditStatement(env, user, 'employees', 'new', newSnapshot, { source: 'import' })
        );
      }
      processed++;
    }

    if (stmts.length > 0) await executeAtomicImportBatch(env, stmts);
    return ok({ message: `Berhasil memperbarui/mengimport ${processed} karyawan` }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
  } catch (err) {
    console.error('importEmployees error:', err);
    return error('Gagal mengimport data karyawan', 500, origin);
  }
}

async function bulkDeleteEmployees(request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const { ids } = body;
  if (!Array.isArray(ids) || ids.length === 0) return error('ids array required', 400, origin);

  // Validate all are positive integers
  const validIds = ids.filter(id => Number.isInteger(Number(id)) && Number(id) > 0).map(Number);
  if (validIds.length === 0) return error('No valid IDs provided', 400, origin);

  // Build parameterized placeholders: ?,?,?
  const placeholders = validIds.map(() => '?').join(',');

  try {
    const allActiveRes = await env.DB.prepare(
      `SELECT id, branch_id, division FROM employees WHERE id IN (${placeholders}) AND deleted_at IS NULL`
    ).bind(...validIds).all();
    const allActive = allActiveRes.results || [];

    const existingRes = await env.DB.prepare(
      `SELECT e.* FROM employees e
       JOIN branches b ON e.branch_id = b.id
       WHERE e.id IN (${placeholders}) AND e.deleted_at IS NULL
         AND e.division IN ('FACILITY CARE', 'SECURITY')
         AND ${getManagementBranchSqlPredicate('b')}`
    ).bind(...validIds).all();
    const existingRows = existingRes.results || [];
    if (existingRows.length !== allActive.length) {
      return error('Terdapat karyawan di luar cakupan Management (000. Management) atau divisi resmi; penghapusan massal ditolak', 400, origin);
    }
    if (existingRows.length === 0) {
      return ok({ message: 'Berhasil menghapus 0 karyawan', deleted: 0 }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
    }

    const verifiedIds = existingRows.map(r => Number(r.id));
    const verifiedPh = verifiedIds.map(() => '?').join(',');
    const batchDeletedAt = new Date().toISOString();
    const deleteStmt = env.DB.prepare(
      `UPDATE employees SET deleted_at = ?, updated_at = datetime('now') WHERE id IN (${verifiedPh}) AND deleted_at IS NULL`
    ).bind(batchDeletedAt, ...verifiedIds);

    const stmts = [
      deleteStmt,
      ...existingRows.map(row =>
        buildSoftDeleteRecordAuditStatement(env, user, 'employees', 'employees', 'id', 'deleted_at', batchDeletedAt, row, { bulk: true, batch_size: validIds.length })
      )
    ];

    const batchRes = await env.DB.batch(stmts);
    const changes = batchRes?.[0]?.meta?.changes ?? batchRes?.[0]?.changes ?? existingRows.length;
    return ok({ message: `Berhasil menghapus ${changes} karyawan`, deleted: changes }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
  } catch (err) {
    console.error('bulkDeleteEmployees error:', err);
    return error('Gagal menghapus data karyawan', 500, origin);
  }
}
