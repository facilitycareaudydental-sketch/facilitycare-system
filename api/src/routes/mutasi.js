import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound } from '../utils/response.js';
import { getPagination } from '../utils/pagination.js';

export async function handleMutasi(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'mutasi', 'read')) return forbidden(origin);

  const url = new URL(request.url);
  const path = url.pathname.replace('/api/mutasi', '');
  const idMatch = path.match(/^\/(\d+)$/);

  // GET list
  if (request.method === 'GET' && (path === '' || path === '/')) {
    try {
      const { page, limit, offset } = getPagination(request.url);
      const search = url.searchParams.get('search') || '';
      const from_branch_id = url.searchParams.get('from_branch_id') || '';
      const to_branch_id = url.searchParams.get('to_branch_id') || '';

      let conditions = ['1=1'];
      let values = [];
      if (search) { conditions.push('(m.employee_name LIKE ?)'); values.push(`%${search}%`); }
      if (from_branch_id) { conditions.push('m.from_branch_id = ?'); values.push(from_branch_id); }
      if (to_branch_id) { conditions.push('m.to_branch_id = ?'); values.push(to_branch_id); }
      const where = 'WHERE ' + conditions.join(' AND ');

      const countRow = await env.DB.prepare(
        `SELECT COUNT(*) as total FROM mutasi_data m ${where}`
      ).bind(...values).first();

      const { results } = await env.DB.prepare(`
        SELECT m.*, b1.full_name as from_branch_name, b2.full_name as to_branch_name
        FROM mutasi_data m
        LEFT JOIN branches b1 ON m.from_branch_id = b1.id
        LEFT JOIN branches b2 ON m.to_branch_id = b2.id
        ${where}
        ORDER BY m.tanggal DESC, m.id DESC
        LIMIT ? OFFSET ?
      `).bind(...values, limit, offset).all();

      const total = countRow?.total || 0;
      return new Response(JSON.stringify({
        success: true,
        data: results || [],
        pagination: { total, page, limit, pages: Math.ceil(total / limit) }
      }), { status: 200, headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': origin } });
    } catch (e) {
      console.error('Mutasi GET error:', e);
      return error('Server error: ' + e.message, 500, origin);
    }
  }

  // POST create
  if (request.method === 'POST' && (path === '' || path === '/')) {
    if (!hasPermission(user, 'mutasi', 'write')) return forbidden(origin);
    try {
      let body;
      try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
      const { tanggal, employee_name, from_branch_id, to_branch_id, status, document_link } = body;
      if (!employee_name) return error('employee_name required', 400, origin);

      const result = await env.DB.prepare(
        'INSERT INTO mutasi_data (tanggal, employee_name, from_branch_id, to_branch_id, status, document_link) VALUES (?, ?, ?, ?, ?, ?)'
      ).bind(tanggal || null, employee_name, from_branch_id || null, to_branch_id || null, status !== null && status !== undefined && status !== '' ? status : '', document_link || null).run();

      return ok({ id: result.meta.last_row_id, message: 'Mutasi berhasil ditambahkan' }, 201, origin);
    } catch (e) {
      console.error('Mutasi POST error:', e);
      return error('Server error: ' + e.message, 500, origin);
    }
  }

  // GET single / PUT / DELETE by id
  if (idMatch) {
    const id = idMatch[1];

    if (request.method === 'GET') {
      const row = await env.DB.prepare(
        `SELECT m.*, b1.full_name as from_branch_name, b2.full_name as to_branch_name
         FROM mutasi_data m
         LEFT JOIN branches b1 ON m.from_branch_id = b1.id
         LEFT JOIN branches b2 ON m.to_branch_id = b2.id
         WHERE m.id = ?`
      ).bind(id).first();
      if (!row) return notFound(origin);
      return ok(row, 200, origin);
    }

    if (request.method === 'PUT') {
      if (!hasPermission(user, 'mutasi', 'write')) return forbidden(origin);
      try {
        let body;
        try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
        const { tanggal, employee_name, from_branch_id, to_branch_id, status, document_link } = body;
        await env.DB.prepare(
          `UPDATE mutasi_data SET tanggal=COALESCE(?,tanggal), employee_name=COALESCE(?,employee_name),
           from_branch_id=COALESCE(?,from_branch_id), to_branch_id=COALESCE(?,to_branch_id),
           status=COALESCE(?,status), document_link=COALESCE(?,document_link) WHERE id=?`
        ).bind(tanggal || null, employee_name || null, from_branch_id || null, to_branch_id || null, status !== null && status !== undefined && status !== '' ? status : '', document_link || null, id).run();
        return ok({ message: 'Mutasi berhasil diperbarui' }, 200, origin);
      } catch (e) {
        return error('Server error: ' + e.message, 500, origin);
      }
    }

    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'mutasi', 'delete')) return forbidden(origin);
      try {
        await env.DB.prepare('DELETE FROM mutasi_data WHERE id=?').bind(id).run();
        return ok({ message: 'Mutasi berhasil dihapus' }, 200, origin);
      } catch (e) {
        return error('Server error: ' + e.message, 500, origin);
      }
    }
  }

  return notFound(origin);
}
