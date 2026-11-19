import { error, ok, paginated, notFound, forbidden, unauthorized } from '../utils/response.js';
import { authenticate, hasPermission } from '../utils/auth.js';
import { sanitizeAuditData } from '../utils/audit.js';
import { getPagination, HARD_MAX_PAGE_LIMIT } from '../utils/pagination.js';
import { getExportPagination } from '../utils/export.js';

function sanitizeAuditRow(row) {
  if (!row) return null;
  const sanitized = { ...row };
  if (!sanitized.old_data && sanitized.old_value) sanitized.old_data = sanitized.old_value;
  if (!sanitized.new_data && sanitized.new_value) sanitized.new_data = sanitized.new_value;
  if (sanitized.old_data) sanitized.old_data = sanitizeAuditData(sanitized.old_data);
  if (sanitized.new_data) sanitized.new_data = sanitizeAuditData(sanitized.new_data);
  return sanitized;
}

export async function handleAuditLogs(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) {
    return unauthorized(origin);
  }

  if (!hasPermission(user, 'audit_logs', 'admin')) {
    return forbidden(origin);
  }

  const url = new URL(request.url);
  const subPath = url.pathname.replace('/api/audit-logs', '');

  if (request.method === 'GET') {
    // GET /api/audit-logs/:id — single record detail lookup for Audit Log modal
    const idMatch = subPath.match(/^\/(\d+)$/);
    if (idMatch) {
      const targetId = Number(idMatch[1]);
      try {
        const row = await env.DB.prepare('SELECT * FROM audit_logs WHERE id = ?').bind(targetId).first();
        if (!row) return notFound(origin);
        return ok(sanitizeAuditRow(row), 200, origin);
      } catch (e) {
        console.error('handleAuditLogs detail DB error:', e);
        return error('Gagal mengambil detail audit log', 500, origin);
      }
    }

    const isExport = subPath === '/export';
    if (subPath !== '' && subPath !== '/' && !isExport) {
      return notFound(origin);
    }

    const search = (url.searchParams.get('search') || '').trim();
    const moduleFilter = (url.searchParams.get('module') || '').trim();
    const actionFilter = (url.searchParams.get('action') || '').trim();
    const roleFilter = (url.searchParams.get('role') || '').trim();
    const { page, limit, offset } = isExport
      ? getExportPagination(request.url)
      : getPagination(request.url, {
          defaultLimit: 50,
          maxLimit: HARD_MAX_PAGE_LIMIT,
        });

    let conditions = [];
    let values = [];

    if (search) {
      conditions.push('(CAST(id AS TEXT) = ? OR user_name LIKE ? OR user_role LIKE ? OR module LIKE ? OR details LIKE ? OR target_id LIKE ? OR old_data LIKE ? OR new_data LIKE ?)');
      const s = `%${search}%`;
      values.push(search, s, s, s, s, s, s, s);
    }
    if (moduleFilter) {
      if (moduleFilter === 'reports') {
        conditions.push("(module = 'reports' OR module LIKE 'reports/%')");
      } else {
        conditions.push('module = ?');
        values.push(moduleFilter);
      }
    }
    if (actionFilter) {
      conditions.push('action = ?');
      values.push(actionFilter);
    }
    if (roleFilter) {
      conditions.push('user_role = ?');
      values.push(roleFilter);
    }

    const where = conditions.length > 0 ? 'WHERE ' + conditions.join(' AND ') : '';

    try {
      const [countResult, rows] = await Promise.all([
        env.DB.prepare(`SELECT COUNT(*) as total FROM audit_logs ${where}`).bind(...values).first(),
        env.DB.prepare(`SELECT * FROM audit_logs ${where} ORDER BY created_at DESC, id DESC LIMIT ? OFFSET ?`)
          .bind(...values, limit, offset).all()
      ]);

      const sanitizedRows = (rows.results || []).map(sanitizeAuditRow);

      return paginated(sanitizedRows, countResult?.total || 0, page, limit, origin);
    } catch (e) {
      console.error('handleAuditLogs DB error:', e);
      return error('Gagal mengambil data audit log', 500, origin);
    }
  }

  return error('Method not allowed', 405, origin);
}
