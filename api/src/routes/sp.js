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
      ).bind(tanggal || null, employee_name, branch_id || null, sp_type || '', status !== null && status !== undefined && status !== '' ? status : '', document_link || null).run();

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
        ).bind(tanggal || null, employee_name || null, branch_id || null, sp_type || null, status !== null && status !== undefined && status !== '' ? status : '', document_link || null, id).run();
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

  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'sp', 'write')) return forbidden(origin);
    let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
    if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
    if (body.length === 0) return ok({ message: 'No data to import' }, 200, origin);

    const existing = await env.DB.prepare('SELECT id, branch_id, tanggal, employee_name, sp_type FROM sp_data').all();
    const existingMap = new Map();
    (existing.results || []).forEach(s => {
      if (s.branch_id && s.tanggal && s.employee_name && s.sp_type) {
        existingMap.set(s.branch_id + '_' + s.tanggal + '_' + s.employee_name.toLowerCase().trim() + '_' + s.sp_type.toLowerCase().trim(), s.id);
      }
    });

    const stmts = [];
    let inserted = 0;
    let updated = 0;

    for (const item of body) {
      if (!item.branch_id || !item.tanggal || !item.employee_name || !item.sp_type) continue;
      
      const key = item.branch_id + '_' + item.tanggal + '_' + item.employee_name.toLowerCase().trim() + '_' + item.sp_type.toLowerCase().trim();
      const status = item.status || 'Aktif';

      if (existingMap.has(key)) {
        const id = existingMap.get(key);
        stmts.push(
          env.DB.prepare(
            `UPDATE sp_data SET status = ?, document_link = ? WHERE id = ?`
          ).bind(
            status,
            item.document_link || null,
            id
          )
        );
        updated++;
      } else {
        stmts.push(
          env.DB.prepare(
            `INSERT INTO sp_data (tanggal, employee_name, branch_id, sp_type, status, document_link) 
             VALUES (?, ?, ?, ?, ?, ?)`
          ).bind(
            item.tanggal,
            item.employee_name,
            item.branch_id,
            item.sp_type,
            status,
            item.document_link || null
          )
        );
        inserted++;
      }
    }

    try {
      if (stmts.length > 0) await env.DB.batch(stmts);
      return ok({ message: 'Berhasil mengimport data', inserted, updated, total: body.length }, 200, origin);
    } catch (err) {
      return error('Gagal import data: ' + err.message, 500, origin);
    }
  }

  return notFound(origin);
}
