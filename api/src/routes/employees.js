import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination, getSearchParam } from '../utils/pagination.js';

export async function handleEmployees(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'employees', 'read')) return forbidden(origin);

  const url = new URL(request.url);
  const path = url.pathname.replace('/api/employees', '');
  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'GET' && path === '') return listEmployees(request, env, origin);
  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'employees', 'write')) return forbidden(origin);
    return importEmployees(request, env, origin);
  }
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'employees', 'write')) return forbidden(origin);
    return createEmployee(request, env, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') return getEmployee(id, env, origin);
    if (request.method === 'PUT') {
      if (!hasPermission(user, 'employees', 'write')) return forbidden(origin);
      return updateEmployee(id, request, env, origin);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'employees', 'delete')) return forbidden(origin);
      return deleteEmployee(id, env, origin);
    }
  }
  return error('Not found', 404, origin);
}

async function listEmployees(request, env, origin) {
  const { page, limit, offset } = getPagination(request.url);
  const url = new URL(request.url);
  const search = url.searchParams.get('search') || '';
  const branch_id = url.searchParams.get('branch_id') || '';
  const division = url.searchParams.get('division') || '';
  const status = url.searchParams.get('status') || '';

  let conditions = [];
  let values = [];
  if (search) { conditions.push('(e.full_name LIKE ?)'); values.push(`%${search}%`); }
  if (branch_id) { conditions.push('e.branch_id = ?'); values.push(branch_id); }
  if (division) { conditions.push('e.division = ?'); values.push(division); }
  if (status) { conditions.push('e.status = ?'); values.push(status); }

  const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';

  const countResult = await env.DB.prepare(
    `SELECT COUNT(*) as total FROM employees e ${where}`
  ).bind(...values).first();

  const rows = await env.DB.prepare(
    `SELECT e.*, b.full_name as branch_name 
     FROM employees e 
     LEFT JOIN branches b ON e.branch_id = b.id 
     ${where} ORDER BY e.full_name LIMIT ? OFFSET ?`
  ).bind(...values, limit, offset).all();

  return paginated(rows.results, countResult.total, page, limit, origin);
}

async function getEmployee(id, env, origin) {
  const row = await env.DB.prepare(
    `SELECT e.*, b.full_name as branch_name 
     FROM employees e LEFT JOIN branches b ON e.branch_id = b.id WHERE e.id = ?`
  ).bind(id).first();
  if (!row) return notFound(origin);

  // Get contracts
  const contracts = await env.DB.prepare(
    'SELECT * FROM contracts WHERE employee_id = ? ORDER BY start_date DESC'
  ).bind(id).all();

  return ok({ ...row, contracts: contracts.results }, 200, origin);
}

async function createEmployee(request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const { full_name, branch_id, division, phone, join_date, status, notes } = body;
  if (!full_name) return error('full_name required', 400, origin);

  const result = await env.DB.prepare(
    'INSERT INTO employees (full_name, branch_id, division, phone, join_date, status, notes) VALUES (?, ?, ?, ?, ?, ?, ?)'
  ).bind(full_name, branch_id || null, division || 'FACILITY CARE', phone || null, join_date || null, status || 'Aktif', notes || null).run();

  return ok({ id: result.meta.last_row_id }, 201, origin);
}

async function updateEmployee(id, request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const existing = await env.DB.prepare('SELECT id FROM employees WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);

  const { full_name, branch_id, division, phone, join_date, status, notes } = body;
  await env.DB.prepare(
    `UPDATE employees SET 
      full_name = COALESCE(?, full_name),
      branch_id = COALESCE(?, branch_id),
      division = COALESCE(?, division),
      phone = COALESCE(?, phone),
      join_date = COALESCE(?, join_date),
      status = COALESCE(?, status),
      notes = COALESCE(?, notes),
      updated_at = datetime('now')
     WHERE id = ?`
  ).bind(full_name || null, branch_id || null, division || null, phone || null,
    join_date || null, status || null, notes || null, id).run();

  return ok({ message: 'Employee updated' }, 200, origin);
}

async function deleteEmployee(id, env, origin) {
  const existing = await env.DB.prepare('SELECT id FROM employees WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);
  await env.DB.prepare("UPDATE employees SET status = 'Tidak Aktif', updated_at = datetime('now') WHERE id = ?").bind(id).run();
  return ok({ message: 'Employee deactivated' }, 200, origin);
}

async function importEmployees(request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
  if (body.length === 0) return ok({ message: 'No data to import' }, 200, origin);

  const stmts = [];
  for (const item of body) {
    if (!item.full_name) continue;
    stmts.push(
      env.DB.prepare(
        `INSERT INTO employees (full_name, branch_id, division, phone, join_date, status, notes) 
         VALUES (?, ?, ?, ?, ?, ?, ?)`
      ).bind(
        item.full_name, 
        item.branch_id || null, 
        item.division || 'FACILITY CARE', 
        item.phone || null, 
        item.join_date || null, 
        item.status || 'Aktif', 
        item.notes || null
      )
    );
  }

  try {
    if (stmts.length > 0) await env.DB.batch(stmts);
    return ok({ message: `Berhasil mengimport ${stmts.length} karyawan` }, 200, origin);
  } catch (err) {
    return error('Gagal import data: ' + err.message, 500, origin);
  }
}
