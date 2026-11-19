import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination, resolveSortOrder } from '../utils/pagination.js';
import { requireExportAccess, getExportPagination } from '../utils/export.js';
import { runSync } from '../utils/calendar.js';
import { hasTableColumn } from '../utils/schemaCompat.js';
import { buildAuditStatement, buildCreateAuditStatement, executeAtomicImportBatch } from '../utils/audit.js';
import { getManagementBranchSubquery, getManagementBranchSqlPredicate, resolveManagementBranchIds } from '../utils/branchMaster.js';

const OFFICIAL_DIVISIONS = Object.freeze(['FACILITY CARE', 'SECURITY']);

const CONTRACT_SORT_MAP = Object.freeze({
  end_date_asc: "CASE WHEN c.end_date IS NULL OR c.end_date = '' THEN 1 ELSE 0 END, c.end_date ASC, c.id ASC",
  end_date_desc: "CASE WHEN c.end_date IS NULL OR c.end_date = '' THEN 1 ELSE 0 END, c.end_date DESC, c.id DESC",
  name_asc: 'e.full_name ASC, c.id ASC',
  newest: 'c.id DESC',
});

export async function handleContracts(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace('/api/contracts', '');

  if (request.method === 'GET' && path === '/export') {
    const { response } = await requireExportAccess(request, env, origin, 'contracts');
    if (response) return response;
    return exportContracts(request, env, origin);
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'contracts', 'read')) return forbidden(origin);

  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'GET' && path === '') return listContracts(request, env, origin);
  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'contracts', 'import')) return forbidden(origin);
    return importContracts(request, env, origin, user);
  }
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'contracts', 'write')) return forbidden(origin);
    return createContract(request, env, origin, user);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') return getContract(id, env, origin);
    if (request.method === 'PUT' || request.method === 'PATCH') {
      if (!hasPermission(user, 'contracts', 'write')) return forbidden(origin);
      return updateContract(id, request, env, origin, user);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'contracts', 'delete')) return forbidden(origin);
      return deleteContract(id, env, origin, user);
    }
  }
  return error('Not found', 404, origin);
}

function buildContractsFilter(url, hasEmpNameCol = true) {
  const search = (url.searchParams.get('search') || '').trim();
  const branch_id = (url.searchParams.get('branch_id') || '').trim();
  const division = (url.searchParams.get('division') || '').trim();
  const status = (url.searchParams.get('status') || '').trim();
  const expiring_days = (url.searchParams.get('expiring_days') || '').trim();
  const month_expiry = (url.searchParams.get('month_expiry') || '').trim();
  const sort = url.searchParams.get('sort') || 'end_date_asc';

  // Phase 5.66.10A.1: Management Contract Scope Enforcement
  const conditions = [
    'c.deleted_at IS NULL',
    "c.division IN ('FACILITY CARE', 'SECURITY')",
    getManagementBranchSubquery('c.branch_id')
  ];
  const values = [];

  if (search) {
    if (hasEmpNameCol) {
      conditions.push('(COALESCE(e.full_name, c.employee_name) LIKE ? OR b.full_name LIKE ? OR c.pkwt_number LIKE ? OR c.contract_type LIKE ? OR c.notes LIKE ?)');
    } else {
      conditions.push('(e.full_name LIKE ? OR b.full_name LIKE ? OR c.pkwt_number LIKE ? OR c.contract_type LIKE ? OR c.notes LIKE ?)');
    }
    values.push(`%${search}%`, `%${search}%`, `%${search}%`, `%${search}%`, `%${search}%`);
  }
  if (branch_id) { conditions.push('c.branch_id = ?'); values.push(branch_id); }
  if (division) { conditions.push('c.division = ?'); values.push(division); }

  if (status === 'Expiring30') {
    conditions.push(
      "c.status = 'Aktif' AND c.end_date IS NOT NULL AND c.end_date != '' AND c.end_date != '-' AND c.end_date NOT LIKE '2099%' AND CAST(julianday(c.end_date) - julianday(date('now', '+7 hours')) AS INTEGER) BETWEEN 0 AND 30"
    );
  } else if (status === 'Tetap') {
    conditions.push("c.status = 'Aktif' AND c.end_date LIKE '2099%'");
  } else if (status === 'Nonaktif') {
    conditions.push("(c.status != 'Aktif' OR c.status IS NULL)");
  } else if (status) {
    conditions.push('c.status = ?');
    values.push(status);
  }

  if (expiring_days && expiring_days !== 'all') {
    const parsedDays = Number.parseInt(expiring_days, 10);
    if (Number.isInteger(parsedDays) && parsedDays >= 0) {
      conditions.push(
        "c.status = 'Aktif' AND c.end_date IS NOT NULL AND c.end_date != '' AND c.end_date != '-' AND c.end_date NOT LIKE '2099%' AND CAST(julianday(c.end_date) - julianday(date('now', '+7 hours')) AS INTEGER) BETWEEN 0 AND ?"
      );
      values.push(parsedDays);
    }
  }

  if (month_expiry && /^\d{4}-(0[1-9]|1[0-2])$/.test(month_expiry)) {
    conditions.push("c.status = 'Aktif' AND c.end_date IS NOT NULL AND c.end_date NOT LIKE '2099%' AND strftime('%Y-%m', c.end_date) = ?");
    values.push(month_expiry);
  }

  const where = 'WHERE ' + conditions.join(' AND ');
  const countJoin = search
    ? 'LEFT JOIN employees e ON c.employee_id = e.id LEFT JOIN branches b ON c.branch_id = b.id'
    : '';
  const sortMap = hasEmpNameCol
    ? { ...CONTRACT_SORT_MAP, name_asc: 'COALESCE(e.full_name, c.employee_name) ASC, c.id ASC' }
    : CONTRACT_SORT_MAP;
  const orderBy = resolveSortOrder(sort, sortMap, 'end_date_asc');
  return { where, values, countJoin, orderBy };
}

const CONTRACT_SELECT = `
  SELECT c.*, e.full_name as employee_name, e.phone as employee_phone, b.full_name as branch_name,
    CASE
      WHEN c.end_date IS NOT NULL AND c.end_date != '' AND c.end_date != '-'
      THEN CAST(julianday(c.end_date) - julianday(date('now', '+7 hours')) AS INTEGER)
      ELSE NULL
    END as days_remaining
  FROM contracts c
  LEFT JOIN employees e ON c.employee_id = e.id
  LEFT JOIN branches b ON c.branch_id = b.id
`;

function getContractSelect(hasEmpNameCol = true) {
  const empNameExpr = hasEmpNameCol
    ? 'COALESCE(e.full_name, c.employee_name)'
    : 'e.full_name';
  return `
  SELECT c.*, ${empNameExpr} as employee_name, e.phone as employee_phone, b.full_name as branch_name,
    CASE
      WHEN c.end_date IS NOT NULL AND c.end_date != '' AND c.end_date != '-'
      THEN CAST(julianday(c.end_date) - julianday(date('now', '+7 hours')) AS INTEGER)
      ELSE NULL
    END as days_remaining
  FROM contracts c
  LEFT JOIN employees e ON c.employee_id = e.id
  LEFT JOIN branches b ON c.branch_id = b.id
`;
}

async function listContracts(request, env, origin) {
  const url = new URL(request.url);
  const { page, limit, offset } = getPagination(request.url);
  const hasEmpNameCol = await hasTableColumn(env.DB, 'contracts', 'employee_name');
  const selectSql = getContractSelect(hasEmpNameCol);
  const { where, values, countJoin, orderBy } = buildContractsFilter(url, hasEmpNameCol);
  const includeStats = url.searchParams.get('include_stats') === '1';

  const queries = [
    env.DB.prepare(`SELECT COUNT(*) as total FROM contracts c ${countJoin} ${where}`).bind(...values).first(),
    env.DB.prepare(`${selectSql} ${where} ORDER BY ${orderBy} LIMIT ? OFFSET ?`).bind(...values, limit, offset).all(),
  ];

  if (includeStats) {
    queries.push(
      env.DB.prepare(
        `SELECT
           COUNT(*) as total,
           SUM(CASE WHEN c.status = 'Aktif' THEN 1 ELSE 0 END) as active,
           SUM(CASE WHEN c.status = 'Aktif' AND c.end_date IS NOT NULL AND c.end_date != '' AND c.end_date != '-' AND c.end_date NOT LIKE '2099%' AND CAST(julianday(c.end_date) - julianday(date('now', '+7 hours')) AS INTEGER) BETWEEN 0 AND 30 THEN 1 ELSE 0 END) as expiring30,
           SUM(CASE WHEN c.status = 'Aktif' AND c.end_date LIKE '2099%' THEN 1 ELSE 0 END) as permanent,
           SUM(CASE WHEN c.status != 'Aktif' OR c.status IS NULL THEN 1 ELSE 0 END) as nonactive
         FROM contracts c
         WHERE c.deleted_at IS NULL AND c.division IN ('FACILITY CARE', 'SECURITY') AND ${getManagementBranchSubquery('c.branch_id')}`
      ).first()
    );
  }

  const [countResult, rows, statsRow] = await Promise.all(queries);
  const extra = includeStats ? {
    stats: {
      total: Number(statsRow?.total || 0),
      active: Number(statsRow?.active || 0),
      expiring30: Number(statsRow?.expiring30 || 0),
      permanent: Number(statsRow?.permanent || 0),
      nonactive: Number(statsRow?.nonactive || 0),
    }
  } : null;

  return paginated(rows.results || [], countResult?.total || 0, page, limit, origin, env?.ENVIRONMENT, extra);
}

async function exportContracts(request, env, origin) {
  const url = new URL(request.url);
  const { page, limit, offset } = getExportPagination(request.url);
  const hasEmpNameCol = await hasTableColumn(env.DB, 'contracts', 'employee_name');
  const selectSql = getContractSelect(hasEmpNameCol);
  const { where, values, countJoin, orderBy } = buildContractsFilter(url, hasEmpNameCol);

  const [countResult, rows] = await Promise.all([
    env.DB.prepare(`SELECT COUNT(*) as total FROM contracts c ${countJoin} ${where}`).bind(...values).first(),
    env.DB.prepare(`${selectSql} ${where} ORDER BY ${orderBy} LIMIT ? OFFSET ?`).bind(...values, limit, offset).all()
  ]);

  return paginated(rows.results || [], countResult?.total || 0, page, limit, origin, env?.ENVIRONMENT);
}

async function getContract(id, env, origin) {
  const hasEmpNameCol = await hasTableColumn(env.DB, 'contracts', 'employee_name');
  const selectSql = getContractSelect(hasEmpNameCol);
  const row = await env.DB.prepare(
    `${selectSql} WHERE c.id = ? AND c.deleted_at IS NULL AND c.division IN ('FACILITY CARE', 'SECURITY') AND ${getManagementBranchSubquery('c.branch_id')}`
  ).bind(id).first();
  if (!row) return notFound(origin);
  return ok(row, 200, origin);
}

async function createContract(request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
  const { employee_id, employee_name, branch_id, division, start_date, end_date, contract_type, pkwt_number, status, notes } = body;
  const hasEmpNameCol = await hasTableColumn(env.DB, 'contracts', 'employee_name');
  const manualName = typeof employee_name === 'string' ? employee_name.trim() : '';
  const hasEmpId = employee_id !== undefined && employee_id !== null && String(employee_id).trim() !== '';
  if (!hasEmpId && !manualName) return error('Nama karyawan wajib diisi', 400, origin);
  if (!hasEmpId && !hasEmpNameCol) return error('employee_id is required on this schema', 400, origin);

  let cleanDivision = 'FACILITY CARE';
  if (division !== undefined && division !== null && String(division).trim() !== '') {
    const divUpper = String(division).trim().toUpperCase();
    if (!OFFICIAL_DIVISIONS.includes(divUpper)) {
      return error('Divisi hanya boleh FACILITY CARE atau SECURITY', 400, origin);
    }
    cleanDivision = divUpper;
  }

  let parsedEmpId = null;
  let emp = null;
  if (hasEmpId) {
    parsedEmpId = Number(employee_id);
    if (!Number.isInteger(parsedEmpId) || parsedEmpId <= 0) {
      return error('Invalid employee_id', 400, origin);
    }
    emp = await env.DB.prepare(
      `SELECT e.id, e.full_name, e.branch_id FROM employees e
       JOIN branches b ON e.branch_id = b.id
       WHERE e.id = ? AND e.deleted_at IS NULL
         AND e.division IN ('FACILITY CARE', 'SECURITY')
         AND ${getManagementBranchSqlPredicate('b')}`
    ).bind(parsedEmpId).first();
    if (!emp) {
      return error('Karyawan tidak ditemukan atau berada di luar cakupan Management', 400, origin);
    }
  }

  let targetBranchId = branch_id;
  if ((targetBranchId === undefined || targetBranchId === null || targetBranchId === '') && emp?.branch_id) {
    targetBranchId = emp.branch_id;
  }

  // Phase 5.66.10A.1: Verify contract branch is strictly within Management scope
  if (targetBranchId !== undefined && targetBranchId !== null && targetBranchId !== '') {
    const parsedBranchId = Number(targetBranchId);
    if (!Number.isInteger(parsedBranchId) || parsedBranchId <= 0) {
      return error('Invalid branch_id', 400, origin);
    }
    const branchRow = await env.DB.prepare(
      `SELECT id FROM branches WHERE id = ? AND deleted_at IS NULL AND ${getManagementBranchSqlPredicate('')}`
    ).bind(parsedBranchId).first();
    if (!branchRow) {
      return error('Pembuatan kontrak dibatasi hanya untuk area Management (Cabang 000. Management)', 400, origin);
    }
    targetBranchId = parsedBranchId;
  } else {
    return error('Cabang kontrak wajib diisi dan harus berupa area Management (Cabang 000. Management)', 400, origin);
  }

  const cleanStartDate = (start_date && typeof start_date === 'string' && start_date.trim()) ? start_date.trim() : null;
  const cleanEndDate = (end_date && typeof end_date === 'string' && end_date.trim()) ? end_date.trim() : null;

  if (cleanStartDate && cleanEndDate && cleanEndDate < cleanStartDate) {
    return error('end_date tidak boleh lebih awal dari start_date', 400, origin);
  }

  const dupContract = parsedEmpId
    ? await env.DB.prepare(
        `SELECT id FROM contracts WHERE employee_id = ? AND COALESCE(start_date, '') = COALESCE(?, '') AND COALESCE(end_date, '') = COALESCE(?, '') AND COALESCE(contract_type, '') = COALESCE(?, '') AND deleted_at IS NULL`
      ).bind(parsedEmpId, cleanStartDate, cleanEndDate, contract_type || null).first()
    : await env.DB.prepare(
        `SELECT id FROM contracts WHERE employee_id IS NULL AND LOWER(TRIM(COALESCE(employee_name, ''))) = LOWER(?) AND COALESCE(start_date, '') = COALESCE(?, '') AND COALESCE(end_date, '') = COALESCE(?, '') AND COALESCE(contract_type, '') = COALESCE(?, '') AND deleted_at IS NULL`
      ).bind(manualName, cleanStartDate, cleanEndDate, contract_type || null).first();
  if (dupContract) {
    return error('Kontrak dengan karyawan, periode, dan tipe yang sama sudah terdaftar', 409, origin);
  }

  const storedEmpName = parsedEmpId ? (emp?.full_name || manualName || null) : manualName;
  const normalizedStatus = status !== null && status !== undefined && status !== '' ? status : '';
  const insertStmt = hasEmpNameCol
    ? env.DB.prepare(
        'INSERT INTO contracts (employee_id, employee_name, branch_id, division, start_date, end_date, contract_type, pkwt_number, status, notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'
      ).bind(
        parsedEmpId,
        storedEmpName,
        targetBranchId || null,
        cleanDivision,
        cleanStartDate,
        cleanEndDate,
        contract_type || null,
        pkwt_number || null,
        normalizedStatus,
        notes || null
      )
    : env.DB.prepare(
        'INSERT INTO contracts (employee_id, branch_id, division, start_date, end_date, contract_type, pkwt_number, status, notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)'
      ).bind(
        parsedEmpId,
        targetBranchId || null,
        cleanDivision,
        cleanStartDate,
        cleanEndDate,
        contract_type || null,
        pkwt_number || null,
        normalizedStatus,
        notes || null
      );

  const newSnapshot = {
    employee_id: parsedEmpId,
    ...(hasEmpNameCol ? { employee_name: storedEmpName } : {}),
    branch_id: branch_id || null,
    division: cleanDivision,
    start_date: cleanStartDate,
    end_date: cleanEndDate,
    contract_type: contract_type || null,
    pkwt_number: pkwt_number || null,
    status: normalizedStatus,
    notes: notes || null
  };
  const batchRes = await env.DB.batch([
    insertStmt,
    buildCreateAuditStatement(env, user, 'contracts', 'new', newSnapshot)
  ]);

  const newId = batchRes?.[0]?.meta?.last_row_id;
  const empName = emp ? emp.full_name : (manualName || 'Karyawan');
  
  try {
    await runSync(env.DB, 'contracts', newId, {
      empName,
      branchId: branch_id || null,
      endDate: cleanEndDate,
      status: normalizedStatus
    });
  } catch (syncErr) {
    console.error('Calendar sync error (non-fatal):', syncErr.message);
  }

  return ok({ id: newId }, 201, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function updateContract(id, request, env, origin, user = null) {
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

  // Phase 5.66.10A.1: Direct API enforcement - target contract must be within Management scope
  const existing = await env.DB.prepare(
    `SELECT c.* FROM contracts c 
     JOIN branches b ON c.branch_id = b.id 
     WHERE c.id = ? AND c.deleted_at IS NULL 
       AND c.division IN ('FACILITY CARE', 'SECURITY')
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

  const hasEmpNameCol = await hasTableColumn(env.DB, 'contracts', 'employee_name');
  const empIdProvided = Object.prototype.hasOwnProperty.call(body, 'employee_id');
  const empNameProvided = Object.prototype.hasOwnProperty.call(body, 'employee_name');
  const rawEmpName = typeof body.employee_name === 'string' ? body.employee_name.trim() : '';

  let nextEmployeeId = existing.employee_id;
  let matchedEmpFullName = null;
  if (body.employee_id !== undefined && body.employee_id !== null && String(body.employee_id).trim() !== '') {
    const parsedEmpId = Number(body.employee_id);
    if (!Number.isInteger(parsedEmpId) || parsedEmpId <= 0) {
      return error('Invalid employee_id', 400, origin);
    }
    const empRow = await env.DB.prepare(
      `SELECT e.id, e.full_name FROM employees e
       JOIN branches b ON e.branch_id = b.id
       WHERE e.id = ? AND e.deleted_at IS NULL
         AND e.division IN ('FACILITY CARE', 'SECURITY')
         AND ${getManagementBranchSqlPredicate('b')}`
    ).bind(parsedEmpId).first();
    if (!empRow) {
      return error('Karyawan tidak ditemukan atau berada di luar cakupan Management', 400, origin);
    }
    nextEmployeeId = parsedEmpId;
    matchedEmpFullName = empRow?.full_name || null;
  } else if (empIdProvided && (body.employee_id === null || body.employee_id === '')) {
    if (!hasEmpNameCol) {
      return error('Invalid employee_id', 400, origin);
    }
    nextEmployeeId = null;
  }
  if (body.branch_id !== undefined && body.branch_id !== null && body.branch_id !== '') {
    const parsedBranchId = Number(body.branch_id);
    if (!Number.isInteger(parsedBranchId) || parsedBranchId <= 0) {
      return error('Invalid branch_id', 400, origin);
    }
    const branchRow = await env.DB.prepare(
      `SELECT id FROM branches WHERE id = ? AND deleted_at IS NULL AND ${getManagementBranchSqlPredicate('')}`
    ).bind(parsedBranchId).first();
    if (!branchRow) {
      return error('Cabang kontrak dibatasi hanya untuk area Management (Cabang 000. Management)', 400, origin);
    }
  }

  const employee_id = nextEmployeeId;
  const employee_name = hasEmpNameCol
    ? (empNameProvided
        ? (rawEmpName || matchedEmpFullName || null)
        : (employee_id ? (matchedEmpFullName || existing.employee_name || null) : existing.employee_name))
    : undefined;
  if (!employee_id && !(hasEmpNameCol && employee_name && String(employee_name).trim())) {
    return error('Nama karyawan wajib diisi', 400, origin);
  }
  if (body.branch_id !== undefined && (body.branch_id === null || String(body.branch_id).trim() === '')) {
    return error('Cabang kontrak wajib diisi dan dibatasi hanya untuk area Management (Cabang 000. Management)', 400, origin);
  }
  const branch_id = body.branch_id !== undefined ? Number(body.branch_id) : existing.branch_id;
  if (!branch_id || !Number.isInteger(Number(branch_id)) || Number(branch_id) <= 0) {
    return error('Cabang kontrak wajib diisi (Cabang 000. Management)', 400, origin);
  }
  
  let division = existing.division;
  if (body.division !== undefined) {
    if (body.division === null || String(body.division).trim() === '') {
      return error('Divisi wajib diisi dan hanya boleh FACILITY CARE atau SECURITY', 400, origin);
    }
    const divUpper = String(body.division).trim().toUpperCase();
    if (!OFFICIAL_DIVISIONS.includes(divUpper)) {
      return error('Divisi hanya boleh FACILITY CARE atau SECURITY', 400, origin);
    }
    division = divUpper;
  }
  const start_date = body.start_date !== undefined ? (body.start_date && String(body.start_date).trim() ? String(body.start_date).trim() : null) : existing.start_date;
  const end_date = body.end_date !== undefined ? (body.end_date && String(body.end_date).trim() ? String(body.end_date).trim() : null) : existing.end_date;
  if (start_date && end_date && end_date < start_date) {
    return error('end_date tidak boleh lebih awal dari start_date', 400, origin);
  }
  const contract_type = body.contract_type !== undefined ? (body.contract_type || null) : existing.contract_type;
  const pkwt_number = body.pkwt_number !== undefined ? (body.pkwt_number || null) : existing.pkwt_number;
  const status = body.status !== undefined ? (body.status !== null && body.status !== '' ? body.status : '') : existing.status;
  const notes = body.notes !== undefined ? (body.notes || null) : existing.notes;

  const baseVersion = isExistingNull ? clientVersion : existing.row_version;
  const nextRowVersion = baseVersion + 1;
  const updatedSnapshot = {
    ...existing,
    employee_id,
    ...(hasEmpNameCol ? { employee_name } : {}),
    branch_id,
    division,
    start_date,
    end_date,
    contract_type,
    pkwt_number,
    status,
    notes,
    row_version: nextRowVersion
  };

  const updateStmt = hasEmpNameCol
    ? env.DB.prepare(
        `UPDATE contracts SET 
          employee_id = ?,
          employee_name = ?,
          branch_id = ?,
          division = ?,
          start_date = ?,
          end_date = ?,
          contract_type = ?,
          pkwt_number = ?,
          status = ?,
          notes = ?,
          row_version = COALESCE(row_version, ?) + 1,
          updated_at = datetime('now')
         WHERE id = ? AND (row_version = ? OR (row_version IS NULL AND (? = 1 OR ? = 0)))`
      ).bind(
        employee_id,
        employee_name,
        branch_id,
        division,
        start_date,
        end_date,
        contract_type,
        pkwt_number,
        status,
        notes,
        clientVersion,
        id,
        clientVersion,
        clientVersion,
        clientVersion
      )
    : env.DB.prepare(
        `UPDATE contracts SET 
          employee_id = ?,
          branch_id = ?,
          division = ?,
          start_date = ?,
          end_date = ?,
          contract_type = ?,
          pkwt_number = ?,
          status = ?,
          notes = ?,
          row_version = COALESCE(row_version, ?) + 1,
          updated_at = datetime('now')
         WHERE id = ? AND (row_version = ? OR (row_version IS NULL AND (? = 1 OR ? = 0)))`
      ).bind(
        employee_id,
        branch_id,
        division,
        start_date,
        end_date,
        contract_type,
        pkwt_number,
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
    buildAuditStatement(env, user, 'UPDATE', 'contracts', id, existing, updatedSnapshot, null, 'SUCCESS', true)
  ]);
  const changes = batchRes?.[0]?.meta?.changes ?? batchRes?.[0]?.changes;

  if (changes === 0) {
    return error('Conflict: Record has been modified by another user or session', 409, origin);
  }

  // Sync updated values
  const syncSelectCols = hasEmpNameCol
    ? 'employee_id, employee_name, branch_id, end_date, status'
    : 'employee_id, branch_id, end_date, status';
  const updatedContract = await env.DB.prepare(`SELECT ${syncSelectCols} FROM contracts WHERE id = ?`).bind(id).first();
  if (updatedContract) {
    let empName = updatedContract.employee_name || 'Karyawan';
    if (updatedContract.employee_id) {
      const emp = await env.DB.prepare('SELECT full_name FROM employees WHERE id = ?').bind(updatedContract.employee_id).first();
      if (emp?.full_name) empName = emp.full_name;
    }
    try {
      await runSync(env.DB, 'contracts', id, {
        empName,
        branchId: updatedContract.branch_id,
        endDate: updatedContract.end_date,
        status: updatedContract.status
      });
    } catch (syncErr) {
      console.error('Calendar sync error (non-fatal):', syncErr.message);
    }
  }

  return ok({ message: 'Contract updated', row_version: nextRowVersion }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function deleteContract(id, env, origin, user = null) {
  // Phase 5.66.10A.1: Direct API enforcement - target contract must be within Management scope
  const existing = await env.DB.prepare(
    `SELECT c.* FROM contracts c 
     JOIN branches b ON c.branch_id = b.id 
     WHERE c.id = ? AND c.deleted_at IS NULL 
       AND c.division IN ('FACILITY CARE', 'SECURITY')
       AND ${getManagementBranchSqlPredicate('b')}`
  ).bind(id).first();
  if (!existing) return notFound(origin);
  const deleteStmt = env.DB.prepare("UPDATE contracts SET deleted_at = datetime('now'), updated_at = datetime('now') WHERE id = ? AND deleted_at IS NULL").bind(id);
  const { executeAtomicDelete } = await import('../utils/audit.js');
  const batchRes = await executeAtomicDelete(env, user, deleteStmt, 'contracts', id, existing);
  const changes = batchRes?.[0]?.meta?.changes ?? batchRes?.[0]?.changes;
  if (changes !== undefined && changes === 0) {
    return notFound(origin);
  }
  try { await runSync(env.DB, 'contracts', id, null); } catch (e) { /* non-fatal */ }
  return ok({ message: 'Contract deleted' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function importContracts(request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
  if (body.length === 0) return ok({ message: 'No data to import' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });

  let inserted = 0;
  let updated = 0;
  let skipped = 0;
  try {
    // Phase 5.66.10A.1 / 5.66.19: Strictly fail-closed Management branch resolution
    let mgmtBranchIds;
    try {
      mgmtBranchIds = await resolveManagementBranchIds(env.DB);
    } catch (mgmtErr) {
      return error(mgmtErr.message, 503, origin);
    }

    const [countBeforeRes, existingAll, mgmtEmpRes] = await Promise.all([
      env.DB.prepare(
        `SELECT COUNT(*) as c FROM contracts c
         WHERE c.deleted_at IS NULL AND c.division IN ('FACILITY CARE', 'SECURITY') AND ${getManagementBranchSubquery('c.branch_id')}`
      ).first(),
      env.DB.prepare(
        `SELECT c.* FROM contracts c
         WHERE c.deleted_at IS NULL AND c.division IN ('FACILITY CARE', 'SECURITY') AND ${getManagementBranchSubquery('c.branch_id')}
         ORDER BY c.id ASC`
      ).all(),
      env.DB.prepare(
        `SELECT id, branch_id, division FROM employees
         WHERE deleted_at IS NULL AND division IN ('FACILITY CARE', 'SECURITY') AND ${getManagementBranchSubquery('branch_id')}`
      ).all()
    ]);
    const countBefore = countBeforeRes.c;
    const mgmtEmpMap = new Map();
    for (const emp of (mgmtEmpRes.results || [])) {
      mgmtEmpMap.set(String(emp.id), emp);
    }
    const latestByEmp = new Map();
    for (const row of (existingAll.results || [])) {
      if (row.employee_id !== null && row.employee_id !== undefined) {
        latestByEmp.set(String(row.employee_id), row);
      }
    }

    const stmts = [];
    for (const item of body) {
      if (!item || typeof item !== 'object' || !item.employee_id) { skipped++; continue; }
      const emp = mgmtEmpMap.get(String(item.employee_id));
      if (!emp) {
        skipped++;
        continue;
      }

      const rawBranchId = (item.branch_id !== undefined && item.branch_id !== null && String(item.branch_id).trim() !== '')
        ? Number(item.branch_id)
        : Number(emp.branch_id);
      if (!Number.isInteger(rawBranchId) || !mgmtBranchIds.has(rawBranchId)) {
        skipped++;
        continue;
      }

      const rawDivProvided = item.division !== undefined && item.division !== null && String(item.division).trim() !== '';
      const rawDiv = rawDivProvided
        ? String(item.division).trim().toUpperCase()
        : String(emp.division || '').trim().toUpperCase();
      if (!OFFICIAL_DIVISIONS.includes(rawDiv)) {
        skipped++;
        continue;
      }
      const safeDivision = rawDiv;
      const cleanStartDate = (item.start_date && typeof item.start_date === 'string' && item.start_date.trim()) ? item.start_date.trim() : null;
      const cleanEndDate = (item.end_date && typeof item.end_date === 'string' && item.end_date.trim()) ? item.end_date.trim() : null;

      const statusVal = item.status !== null && item.status !== undefined && item.status !== '' ? item.status : '';
      const existing = latestByEmp.get(String(item.employee_id));
      const nextFields = {
        employee_id: Number(item.employee_id),
        branch_id: rawBranchId,
        division: safeDivision,
        start_date: cleanStartDate,
        end_date: cleanEndDate,
        contract_type: item.contract_type || null,
        pkwt_number: item.pkwt_number || null,
        status: statusVal,
        notes: item.notes || null
      };

      if (existing) {
        const afterRow = { ...existing, ...nextFields };
        stmts.push(
          env.DB.prepare(
            `UPDATE contracts SET branch_id = ?, division = ?, start_date = ?, end_date = ?, contract_type = ?, pkwt_number = ?, status = ?, notes = ? WHERE id = ?`
          ).bind(
            nextFields.branch_id,
            safeDivision,
            cleanStartDate,
            cleanEndDate,
            nextFields.contract_type,
            nextFields.pkwt_number,
            statusVal,
            nextFields.notes,
            existing.id
          ),
          buildAuditStatement(env, user, 'UPDATE', 'contracts', existing.id, existing, afterRow, { source: 'import' }, 'SUCCESS', true)
        );
        latestByEmp.set(String(item.employee_id), afterRow);
        updated++;
      } else {
        stmts.push(
          env.DB.prepare(
            `INSERT INTO contracts (employee_id, branch_id, division, start_date, end_date, contract_type, pkwt_number, status, notes) 
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
          ).bind(
            item.employee_id,
            nextFields.branch_id,
            safeDivision,
            cleanStartDate,
            cleanEndDate,
            nextFields.contract_type,
            nextFields.pkwt_number,
            statusVal,
            nextFields.notes
          ),
          buildCreateAuditStatement(env, user, 'contracts', 'new', nextFields, { source: 'import' })
        );
        inserted++;
      }
    }

    if (stmts.length > 0) await executeAtomicImportBatch(env, stmts);

    const countAfterRes = await env.DB.prepare(
      `SELECT COUNT(*) as c FROM contracts c
       WHERE c.deleted_at IS NULL AND c.division IN ('FACILITY CARE', 'SECURITY') AND ${getManagementBranchSubquery('c.branch_id')}`
    ).first();
    const countAfter = countAfterRes.c;
    
    return ok({ 
      message: `Berhasil memproses ${inserted + updated} kontrak`,
      metrics: {
        inserted,
        updated,
        skipped,
        countBefore,
        countAfter
      }
    }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
  } catch (err) {
    console.error('importContracts error:', err);
    return error('Gagal mengimport data kontrak', 500, origin);
  }
}
