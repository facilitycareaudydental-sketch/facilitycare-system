import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination } from '../utils/pagination.js';
import { runSync } from '../utils/calendar.js';

export async function handleContracts(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'contracts', 'read')) return forbidden(origin);

  const url = new URL(request.url);
  const path = url.pathname.replace('/api/contracts', '');
  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'GET' && path === '') return listContracts(request, env, origin);
  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'contracts', 'write')) return forbidden(origin);
    return importContracts(request, env, origin);
  }
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'contracts', 'write')) return forbidden(origin);
    return createContract(request, env, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') return getContract(id, env, origin);
    if (request.method === 'PUT') {
      if (!hasPermission(user, 'contracts', 'write')) return forbidden(origin);
      return updateContract(id, request, env, origin);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'contracts', 'delete')) return forbidden(origin);
      return deleteContract(id, env, origin);
    }
  }
  return error('Not found', 404, origin);
}

async function listContracts(request, env, origin) {
  const { page, limit, offset } = getPagination(request.url);
  const url = new URL(request.url);
  const search = url.searchParams.get('search') || '';
  const branch_id = url.searchParams.get('branch_id') || '';
  const status = url.searchParams.get('status') || '';
  const expiring_days = url.searchParams.get('expiring_days') || '';

  let conditions = ['1=1'];
  let values = [];

  if (search) { conditions.push('e.full_name LIKE ?'); values.push(`%${search}%`); }
  if (branch_id) { conditions.push('c.branch_id = ?'); values.push(branch_id); }
  if (status) { conditions.push('c.status = ?'); values.push(status); }
  if (expiring_days) {
    conditions.push("c.end_date <= date('now', '+' || ? || ' days') AND c.end_date >= date('now')");
    values.push(expiring_days);
  }

  const where = 'WHERE ' + conditions.join(' AND ');

  const countResult = await env.DB.prepare(
    `SELECT COUNT(*) as total FROM contracts c LEFT JOIN employees e ON c.employee_id = e.id ${where}`
  ).bind(...values).first();

  const rows = await env.DB.prepare(
    `SELECT c.*, e.full_name as employee_name, b.full_name as branch_name,
      CAST(julianday(c.end_date) - julianday('now') AS INTEGER) as days_remaining
     FROM contracts c 
     LEFT JOIN employees e ON c.employee_id = e.id
     LEFT JOIN branches b ON c.branch_id = b.id
     ${where} ORDER BY c.end_date ASC LIMIT ? OFFSET ?`
  ).bind(...values, limit, offset).all();

  return paginated(rows.results, countResult.total, page, limit, origin);
}

async function getContract(id, env, origin) {
  const row = await env.DB.prepare(
    `SELECT c.*, e.full_name as employee_name, b.full_name as branch_name,
      CAST(julianday(c.end_date) - julianday('now') AS INTEGER) as days_remaining
     FROM contracts c 
     LEFT JOIN employees e ON c.employee_id = e.id
     LEFT JOIN branches b ON c.branch_id = b.id
     WHERE c.id = ?`
  ).bind(id).first();
  if (!row) return notFound(origin);
  return ok(row, 200, origin);
}

async function createContract(request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const { employee_id, branch_id, division, start_date, end_date, contract_type, pkwt_number, status, notes } = body;
  if (!employee_id || !start_date || !end_date) return error('employee_id, start_date, end_date required', 400, origin);

  const result = await env.DB.prepare(
    'INSERT INTO contracts (employee_id, branch_id, division, start_date, end_date, contract_type, pkwt_number, status, notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)'
  ).bind(employee_id, branch_id || null, division || 'FACILITY CARE', start_date, end_date,
    contract_type || null, pkwt_number || null, status !== null && status !== undefined && status !== '' ? status : '', notes || null).run();

  const newId = result.meta.last_row_id;
  const emp = await env.DB.prepare('SELECT full_name FROM employees WHERE id = ?').bind(employee_id).first();
  const empName = emp ? emp.full_name : 'Karyawan';
  
  try {
    await runSync(env.DB, 'contracts', newId, {
      empName,
      branchId: branch_id || null,
      endDate: end_date,
      status: status !== null && status !== undefined && status !== '' ? status : ''
    });
  } catch (syncErr) {
    console.error('Calendar sync error (non-fatal):', syncErr.message);
  }

  return ok({ id: newId }, 201, origin);
}

async function updateContract(id, request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const existing = await env.DB.prepare('SELECT id FROM contracts WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);

  const { employee_id, branch_id, division, start_date, end_date, contract_type, pkwt_number, status, notes } = body;
  await env.DB.prepare(
    `UPDATE contracts SET 
      employee_id = COALESCE(?, employee_id),
      branch_id = COALESCE(?, branch_id),
      division = COALESCE(?, division),
      start_date = COALESCE(?, start_date),
      end_date = COALESCE(?, end_date),
      contract_type = COALESCE(?, contract_type),
      pkwt_number = COALESCE(?, pkwt_number),
      status = COALESCE(?, status),
      notes = COALESCE(?, notes),
      updated_at = datetime('now')
     WHERE id = ?`
  ).bind(employee_id || null, branch_id || null, division || null, start_date || null,
    end_date || null, contract_type || null, pkwt_number || null, status !== null && status !== undefined && status !== '' ? status : '', notes || null, id).run();

  // Sync updated values
  const updatedContract = await env.DB.prepare('SELECT employee_id, branch_id, end_date, status FROM contracts WHERE id = ?').bind(id).first();
  if (updatedContract) {
    const emp = await env.DB.prepare('SELECT full_name FROM employees WHERE id = ?').bind(updatedContract.employee_id).first();
    const empName = emp ? emp.full_name : 'Karyawan';
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

  return ok({ message: 'Contract updated' }, 200, origin);
}

async function deleteContract(id, env, origin) {
  const existing = await env.DB.prepare('SELECT id FROM contracts WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);
  await env.DB.prepare('DELETE FROM contracts WHERE id = ?').bind(id).run();
  try { await runSync(env.DB, 'contracts', id, null); } catch (e) { /* non-fatal */ }
  return ok({ message: 'Contract deleted' }, 200, origin);
}

async function importContracts(request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
  if (body.length === 0) return ok({ message: 'No data to import' }, 200, origin);

  const stmts = [];
  for (const item of body) {
    if (!item.employee_id || !item.start_date || !item.end_date) continue;
    stmts.push(
      env.DB.prepare(
        `INSERT INTO contracts (employee_id, branch_id, division, start_date, end_date, contract_type, pkwt_number, status, notes) 
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(
        item.employee_id,
        item.branch_id || null,
        item.division || 'FACILITY CARE',
        item.start_date,
        item.end_date,
        item.contract_type || null,
        item.pkwt_number || null,
        item.status !== null && item.status !== undefined && item.status !== '' ? item.status : '',
        item.notes || null
      )
    );
  }

  try {
    if (stmts.length > 0) await env.DB.batch(stmts);
    return ok({ message: `Berhasil mengimport ${stmts.length} kontrak` }, 200, origin);
  } catch (err) {
    return error('Gagal import data: ' + err.message, 500, origin);
  }
}
