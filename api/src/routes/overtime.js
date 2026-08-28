import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound } from '../utils/response.js';
import { getPagination } from '../utils/pagination.js';

export async function handleOvertime(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  // Default read permission check
  if (user.role === 'viewer' && request.method !== 'GET') return forbidden(origin);

  const url = new URL(request.url);
  const path = url.pathname.replace('/api/overtime', '');
  const idMatch = path.match(/^\/(\d+)$/);

  // GET list
  if (request.method === 'GET' && (path === '' || path === '/')) {
    try {
      const { page, limit, offset } = getPagination(request.url);
      const search = url.searchParams.get('search') || '';
      const branch_id = url.searchParams.get('branch_id') || '';
      const employee_id = url.searchParams.get('employee_id') || '';

      let conditions = ['1=1'];
      let values = [];
      if (search) { 
        conditions.push('(e.full_name LIKE ? OR b.full_name LIKE ?)'); 
        values.push(`%${search}%`, `%${search}%`); 
      }
      if (branch_id) { conditions.push('o.branch_id = ?'); values.push(branch_id); }
      if (employee_id) { conditions.push('o.employee_id = ?'); values.push(employee_id); }
      
      const where = 'WHERE ' + conditions.join(' AND ');

      const countRow = await env.DB.prepare(
        `SELECT COUNT(*) as total FROM overtime_records o 
         LEFT JOIN branches b ON o.branch_id = b.id 
         LEFT JOIN employees e ON o.employee_id = e.id 
         ${where}`
      ).bind(...values).first();

      const { results } = await env.DB.prepare(`
        SELECT o.*, b.full_name as branch_name, e.full_name as employee_name
        FROM overtime_records o
        LEFT JOIN branches b ON o.branch_id = b.id
        LEFT JOIN employees e ON o.employee_id = e.id
        ${where}
        ORDER BY o.date DESC, o.id DESC
        LIMIT ? OFFSET ?
      `).bind(...values, limit, offset).all();

      const total = countRow?.total || 0;
      return new Response(JSON.stringify({
        success: true,
        data: results || [],
        pagination: { total, page, limit, pages: Math.ceil(total / limit) }
      }), { status: 200, headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': origin } });
    } catch (e) {
      console.error('Overtime GET error:', e);
      return error('Server error: ' + e.message, 500, origin);
    }
  }

  // GET single
  if (request.method === 'GET' && idMatch) {
    try {
      const { results } = await env.DB.prepare(`
        SELECT o.*, b.full_name as branch_name, e.full_name as employee_name
        FROM overtime_records o
        LEFT JOIN branches b ON o.branch_id = b.id
        LEFT JOIN employees e ON o.employee_id = e.id
        WHERE o.id = ?
      `).bind(idMatch[1]).all();
      
      if (!results || results.length === 0) return notFound(origin);
      return ok(results[0], origin);
    } catch (e) {
      return error(e.message, 500, origin);
    }
  }

  // POST (Create)
  if (request.method === 'POST' && (path === '' || path === '/')) {
    if (user.role === 'viewer') return forbidden(origin);
    try {
      const body = await request.json();
      
      const res = await env.DB.prepare(`
        INSERT INTO overtime_records (date, branch_id, employee_id, start_time, end_time, break_hours, total_hours, reason)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        RETURNING id
      `).bind(
        body.date,
        body.branch_id,
        body.employee_id,
        body.start_time,
        body.end_time,
        body.break_hours || 0,
        body.total_hours || 0,
        body.reason
      ).first();
      
      return ok({ id: res.id, message: 'Data lembur berhasil ditambahkan' }, 201, origin);
    } catch (e) {
      return error(e.message, 500, origin);
    }
  }

  // PUT (Update)
  if (request.method === 'PUT' && idMatch) {
    if (user.role === 'viewer') return forbidden(origin);
    try {
      const id = idMatch[1];
      const body = await request.json();
      await env.DB.prepare(`
        UPDATE overtime_records
        SET date=?, branch_id=?, employee_id=?, start_time=?, end_time=?, break_hours=?, total_hours=?, reason=?, updated_at=datetime('now')
        WHERE id=?
      `).bind(
        body.date,
        body.branch_id,
        body.employee_id,
        body.start_time,
        body.end_time,
        body.break_hours || 0,
        body.total_hours || 0,
        body.reason,
        id
      ).run();
      return ok({ message: 'Data lembur berhasil diupdate' }, origin);
    } catch (e) {
      return error(e.message, 500, origin);
    }
  }

  // DELETE
  if (request.method === 'DELETE' && idMatch) {
    if (user.role !== 'superadmin' && user.role !== 'admin') return forbidden(origin);
    try {
      await env.DB.prepare('DELETE FROM overtime_records WHERE id=?').bind(idMatch[1]).run();
      return ok({ message: 'Data lembur berhasil dihapus' }, origin);
    } catch (e) {
      return error(e.message, 500, origin);
    }
  }

  return notFound(origin);
}
