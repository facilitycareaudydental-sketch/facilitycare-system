import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound } from '../utils/response.js';
import { getPagination } from '../utils/pagination.js';

export async function handleSP(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'sp', 'read')) return forbidden(origin);

  const url = new URL(request.url);
  const path = url.pathname.replace('/api/sp', '');
  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'GET' && path === '/setup-db') {
    try {
      await env.DB.prepare(`
        CREATE TABLE IF NOT EXISTS sp_data (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          tanggal TEXT,
          employee_name TEXT NOT NULL,
          branch_id INTEGER,
          sp_type TEXT,
          status TEXT DEFAULT 'Aktif',
          document_link TEXT,
          created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
          FOREIGN KEY (branch_id) REFERENCES branches(id)
        );
      `).run();
      await env.DB.prepare(`
        CREATE TABLE IF NOT EXISTS mutasi_data (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          tanggal TEXT,
          employee_name TEXT NOT NULL,
          from_branch_id INTEGER,
          to_branch_id INTEGER,
          status TEXT DEFAULT 'Proses',
          document_link TEXT,
          created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
          FOREIGN KEY (from_branch_id) REFERENCES branches(id),
          FOREIGN KEY (to_branch_id) REFERENCES branches(id)
        );
      `).run();
      return ok({ message: 'Tables created successfully' }, 200, origin);
    } catch (e) {
      return error('Server error: ' + e.message, 500, origin);
    }
  }

  // GET list
  if (request.method === 'GET' && (path === '' || path === '/')) {
    try {
      const { page, limit, offset } = getPagination(request.url);
      const search = url.searchParams.get('search') || '';
      const branch_id = url.searchParams.get('branch_id') || '';

      let conditions = ['1=1'];
      let values = [];
      if (search) { conditions.push('(s.employee_name LIKE ?)'); values.push(`%${search}%`); }
      if (branch_id) { conditions.push('s.branch_id = ?'); values.push(branch_id); }
      const where = 'WHERE ' + conditions.join(' AND ');

      const countRow = await env.DB.prepare(
        `SELECT COUNT(*) as total FROM sp_data s ${where}`
      ).bind(...values).first();

      const { results } = await env.DB.prepare(`
        SELECT s.*, b.full_name as branch_name
        FROM sp_data s
        LEFT JOIN branches b ON s.branch_id = b.id
        ${where}
        ORDER BY s.tanggal DESC, s.id DESC
        LIMIT ? OFFSET ?
      `).bind(...values, limit, offset).all();

      const total = countRow?.total || 0;
      return new Response(JSON.stringify({
        success: true,
        data: results || [],
        pagination: { total, page, limit, pages: Math.ceil(total / limit) }
      }), { status: 200, headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': origin } });
    } catch (e) {
      console.error('SP GET error:', e);
      return error('Server error: ' + e.message, 500, origin);
    }
  }

  // POST create
  if (request.method === 'POST' && (path === '' || path === '/')) {
    if (!hasPermission(user, 'sp', 'write')) return forbidden(origin);
    try {
      let body;
      try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
      const { tanggal, employee_name, branch_id, sp_type, status, document_link } = body;
      if (!employee_name) return error('employee_name required', 400, origin);

      const result = await env.DB.prepare(
        'INSERT INTO sp_data (tanggal, employee_name, branch_id, sp_type, status, document_link) VALUES (?, ?, ?, ?, ?, ?)'
      ).bind(tanggal || null, employee_name, branch_id || null, sp_type || '', status || 'Aktif', document_link || null).run();

      return ok({ id: result.meta.last_row_id, message: 'SP berhasil ditambahkan' }, 201, origin);
    } catch (e) {
      console.error('SP POST error:', e);
      return error('Server error: ' + e.message, 500, origin);
    }
  }

  // GET single / PUT / DELETE by id
  if (idMatch) {
    const id = idMatch[1];

    if (request.method === 'GET') {
      const row = await env.DB.prepare(
        'SELECT s.*, b.full_name as branch_name FROM sp_data s LEFT JOIN branches b ON s.branch_id = b.id WHERE s.id = ?'
      ).bind(id).first();
      if (!row) return notFound(origin);
      return ok(row, 200, origin);
    }

    if (request.method === 'PUT') {
      if (!hasPermission(user, 'sp', 'write')) return forbidden(origin);
      try {
        let body;
        try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
        const { tanggal, employee_name, branch_id, sp_type, status, document_link } = body;
        await env.DB.prepare(
          `UPDATE sp_data SET tanggal=COALESCE(?,tanggal), employee_name=COALESCE(?,employee_name),
           branch_id=COALESCE(?,branch_id), sp_type=COALESCE(?,sp_type),
           status=COALESCE(?,status), document_link=COALESCE(?,document_link) WHERE id=?`
        ).bind(tanggal || null, employee_name || null, branch_id || null, sp_type || null, status || null, document_link || null, id).run();
        return ok({ message: 'SP berhasil diperbarui' }, 200, origin);
      } catch (e) {
        return error('Server error: ' + e.message, 500, origin);
      }
    }

    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'sp', 'delete')) return forbidden(origin);
      try {
        await env.DB.prepare('DELETE FROM sp_data WHERE id=?').bind(id).run();
        return ok({ message: 'SP berhasil dihapus' }, 200, origin);
      } catch (e) {
        return error('Server error: ' + e.message, 500, origin);
      }
    }
  }

  return notFound(origin);
}
