import { error, ok, paginated, notFound, forbidden } from '../utils/response.js';
import { authenticate, hasPermission } from '../utils/auth.js';

export async function handleAuditLogs(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user || !hasPermission(user, 'audit_logs', 'admin')) {
    return forbidden(origin);
  }

  if (request.method === 'GET') {
    const url = new URL(request.url);
    const search = url.searchParams.get('search') || '';
    const moduleFilter = url.searchParams.get('module') || '';
    const actionFilter = url.searchParams.get('action') || '';
    const limit = parseInt(url.searchParams.get('limit')) || 50;
    const page = parseInt(url.searchParams.get('page')) || 1;
    const offset = (page - 1) * limit;

    let conditions = [];
    let values = [];

    if (search) {
      conditions.push('(user_name LIKE ? OR module LIKE ? OR details LIKE ? OR old_data LIKE ? OR new_data LIKE ?)');
      const s = `%${search}%`;
      values.push(s, s, s, s, s);
    }
    if (moduleFilter) {
      conditions.push('module = ?');
      values.push(moduleFilter);
    }
    if (actionFilter) {
      conditions.push('action = ?');
      values.push(actionFilter);
    }

    const where = conditions.length > 0 ? 'WHERE ' + conditions.join(' AND ') : '';

    try {
      const countResult = await env.DB.prepare(`SELECT COUNT(*) as total FROM audit_logs ${where}`).bind(...values).first();
      const rows = await env.DB.prepare(`SELECT * FROM audit_logs ${where} ORDER BY created_at DESC LIMIT ? OFFSET ?`)
        .bind(...values, limit, offset).all();

      return paginated(rows.results, countResult.total, page, limit, origin);
    } catch (e) {
      return error('Database Error: ' + e.message, 500, origin);
    }
  }

  return error('Method not allowed', 405, origin);
}
