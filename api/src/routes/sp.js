import { authenticate, hasPermission, json, created, badRequest, notFound, serverError, forbidden } from '../utils/auth.js';

export async function handleSP(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) return new Response('Unauthorized', { status: 401, headers: { 'Access-Control-Allow-Origin': origin } });
  if (!hasPermission(user, 'sp', 'read')) return forbidden(origin);

  const url = new URL(request.url);
  const path = url.pathname.replace('/api/sp', '');
  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'GET' && (path === '' || path === '/')) {
    try {
      const { results } = await env.DB.prepare(`
        SELECT s.*, b.name as branch_name
        FROM sp_data s
        LEFT JOIN branches b ON s.branch_id = b.id
        ORDER BY s.id DESC
      `).all();
      return json({ data: results || [] }, origin);
    } catch (e) {
      return serverError(e, origin);
    }
  }

  if (request.method === 'POST' && (path === '' || path === '/')) {
    if (!hasPermission(user, 'sp', 'write')) return forbidden(origin);
    try {
      const body = await request.json();
      const { tanggal, employee_name, branch_id, sp_type, status, document_link } = body;
      
      const { success } = await env.DB.prepare(
        'INSERT INTO sp_data (tanggal, employee_name, branch_id, sp_type, status, document_link) VALUES (?, ?, ?, ?, ?, ?)'
      ).bind(tanggal || null, employee_name || '', branch_id || null, sp_type || '', status || '', document_link || '').run();
      
      if (success) return created({ message: 'SP added successfully' }, origin);
      return badRequest('Failed to add SP', origin);
    } catch (e) {
      return serverError(e, origin);
    }
  }

  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'PUT') {
      if (!hasPermission(user, 'sp', 'write')) return forbidden(origin);
      try {
        const body = await request.json();
        const { tanggal, employee_name, branch_id, sp_type, status, document_link } = body;
        const { success } = await env.DB.prepare(
          'UPDATE sp_data SET tanggal=?, employee_name=?, branch_id=?, sp_type=?, status=?, document_link=? WHERE id=?'
        ).bind(tanggal || null, employee_name || '', branch_id || null, sp_type || '', status || '', document_link || '', id).run();
        
        if (success) return json({ message: 'SP updated' }, origin);
        return badRequest('Failed to update SP', origin);
      } catch (e) {
        return serverError(e, origin);
      }
    }

    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'sp', 'write')) return forbidden(origin);
      try {
        const { success } = await env.DB.prepare('DELETE FROM sp_data WHERE id=?').bind(id).run();
        if (success) return json({ message: 'SP deleted' }, origin);
        return badRequest('Failed to delete SP', origin);
      } catch (e) {
        return serverError(e, origin);
      }
    }
  }

  return notFound(origin);
}
