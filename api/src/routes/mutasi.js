import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';

export async function handleMutasi(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) return new Response('Unauthorized', { status: 401, headers: { 'Access-Control-Allow-Origin': origin } });
  if (!hasPermission(user, 'mutasi', 'read')) return forbidden(origin);

  const url = new URL(request.url);
  const path = url.pathname.replace('/api/mutasi', '');
  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'GET' && (path === '' || path === '/')) {
    try {
      const { results } = await env.DB.prepare(`
        SELECT m.*, b1.name as from_branch_name, b2.name as to_branch_name
        FROM mutasi_data m
        LEFT JOIN branches b1 ON m.from_branch_id = b1.id
        LEFT JOIN branches b2 ON m.to_branch_id = b2.id
        ORDER BY m.id DESC
      `).all();
      return ok({ data: results || [] }, origin);
    } catch (e) {
      return error('Server error', 500, origin);
    }
  }

  if (request.method === 'POST' && (path === '' || path === '/')) {
    if (!hasPermission(user, 'mutasi', 'write')) return forbidden(origin);
    try {
      const body = await request.ok();
      const { tanggal, employee_name, from_branch_id, to_branch_id, status, document_link } = body;
      
      const { success } = await env.DB.prepare(
        'INSERT INTO mutasi_data (tanggal, employee_name, from_branch_id, to_branch_id, status, document_link) VALUES (?, ?, ?, ?, ?, ?)'
      ).bind(tanggal || null, employee_name || '', from_branch_id || null, to_branch_id || null, status || '', document_link || '').run();
      
      if (success) return ok({ message: 'Mutasi added successfully' }, origin);
      return error('Bad request', 400, origin);
    } catch (e) {
      return error('Server error', 500, origin);
    }
  }

  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'PUT') {
      if (!hasPermission(user, 'mutasi', 'write')) return forbidden(origin);
      try {
        const body = await request.ok();
        const { tanggal, employee_name, from_branch_id, to_branch_id, status, document_link } = body;
        const { success } = await env.DB.prepare(
          'UPDATE mutasi_data SET tanggal=?, employee_name=?, from_branch_id=?, to_branch_id=?, status=?, document_link=? WHERE id=?'
        ).bind(tanggal || null, employee_name || '', from_branch_id || null, to_branch_id || null, status || '', document_link || '', id).run();
        
        if (success) return ok({ message: 'Mutasi updated' }, origin);
        return error('Bad request', 400, origin);
      } catch (e) {
        return error('Server error', 500, origin);
      }
    }

    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'mutasi', 'write')) return forbidden(origin);
      try {
        const { success } = await env.DB.prepare('DELETE FROM mutasi_data WHERE id=?').bind(id).run();
        if (success) return ok({ message: 'Mutasi deleted' }, origin);
        return error('Bad request', 400, origin);
      } catch (e) {
        return error('Server error', 500, origin);
      }
    }
  }

  return notFound(origin);
}
