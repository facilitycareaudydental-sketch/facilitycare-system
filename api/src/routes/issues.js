import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination } from '../utils/pagination.js';

export async function handleIssues(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'issues', 'read')) return forbidden(origin);

  const url = new URL(request.url);
  const path = url.pathname.replace('/api/issues', '');
  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'GET' && path === '') return listIssues(request, env, origin);
  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'issues', 'write')) return forbidden(origin);
    return importIssues(request, env, origin);
  }
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'issues', 'write')) return forbidden(origin);
    return createIssue(request, env, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') return getIssue(id, env, origin);
    if (request.method === 'PUT') {
      if (!hasPermission(user, 'issues', 'write')) return forbidden(origin);
      return updateIssue(id, request, env, origin);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'issues', 'delete')) return forbidden(origin);
      return deleteIssue(id, env, origin);
    }
  }
  return error('Not found', 404, origin);
}

async function listIssues(request, env, origin) {
  const { page, limit, offset } = getPagination(request.url);
  const url = new URL(request.url);
  const branch_id = url.searchParams.get('branch_id') || '';
  const category = url.searchParams.get('category') || '';
  const status = url.searchParams.get('status') || '';
  const search = url.searchParams.get('search') || '';
  const year = url.searchParams.get('year') || '';

  let conditions = [];
  let values = [];
  if (branch_id) { conditions.push('i.branch_id = ?'); values.push(branch_id); }
  if (category) { conditions.push('i.category = ?'); values.push(category); }
  if (status) { conditions.push('i.status = ?'); values.push(status); }
  if (search) { conditions.push('(i.complaint LIKE ? OR i.employee_name LIKE ?)'); values.push(`%${search}%`, `%${search}%`); }
  if (year) { conditions.push("strftime('%Y', i.report_date) = ?"); values.push(year); }

  const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';

  const [countResult, rows] = await Promise.all([
    env.DB.prepare(`SELECT COUNT(*) as total FROM issues i ${where}`).bind(...values).first(),
    env.DB.prepare(
      `SELECT i.*, b.full_name as branch_name FROM issues i
       LEFT JOIN branches b ON i.branch_id = b.id
       ${where} ORDER BY i.report_date DESC LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all()
  ]);

  return paginated(rows.results, countResult.total, page, limit, origin);
}

async function getIssue(id, env, origin) {
  const row = await env.DB.prepare(
    'SELECT i.*, b.full_name as branch_name FROM issues i LEFT JOIN branches b ON i.branch_id = b.id WHERE i.id = ?'
  ).bind(id).first();
  if (!row) return notFound(origin);
  return ok(row, 200, origin);
}

async function createIssue(request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const { report_date, branch_id, category, source, complaint, employee_name, fc_specialist, solution, status, completion_date } = body;
  if (!report_date || !complaint || !category) return error('report_date, category, complaint required', 400, origin);

  let day_count = null;
  if (completion_date && report_date) {
    const d1 = new Date(report_date);
    const d2 = new Date(completion_date);
    day_count = Math.floor((d2 - d1) / (1000 * 60 * 60 * 24));
  }

  const result = await env.DB.prepare(
    'INSERT INTO issues (report_date, branch_id, category, source, complaint, employee_name, fc_specialist, solution, status, completion_date, day_count) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'
  ).bind(report_date, branch_id || null, category, source || null, complaint, employee_name || null,
    fc_specialist || null, solution || null, status !== null && status !== undefined && status !== '' ? status : '', completion_date || null, day_count).run();

  return ok({ id: result.meta.last_row_id }, 201, origin);
}

async function updateIssue(id, request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const existing = await env.DB.prepare('SELECT * FROM issues WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);

  const { report_date, branch_id, category, source, complaint, employee_name, fc_specialist, solution, status, completion_date } = body;

  let day_count = existing.day_count;
  const rd = report_date || existing.report_date;
  const cd = completion_date !== undefined ? completion_date : existing.completion_date;
  if (rd && cd) {
    day_count = Math.floor((new Date(cd) - new Date(rd)) / (1000 * 60 * 60 * 24));
  }

  await env.DB.prepare(
    `UPDATE issues SET report_date = COALESCE(?, report_date), branch_id = COALESCE(?, branch_id),
     category = COALESCE(?, category), source = COALESCE(?, source), complaint = COALESCE(?, complaint),
     employee_name = COALESCE(?, employee_name), fc_specialist = COALESCE(?, fc_specialist),
     solution = COALESCE(?, solution), status = COALESCE(?, status),
     completion_date = COALESCE(?, completion_date), day_count = ?, updated_at = datetime('now')
     WHERE id = ?`
  ).bind(report_date || null, branch_id || null, category || null, source || null, complaint || null,
    employee_name || null, fc_specialist || null, solution || null, status !== null && status !== undefined && status !== '' ? status : '',
    completion_date || null, day_count, id).run();

  return ok({ message: 'Issue updated' }, 200, origin);
}

async function deleteIssue(id, env, origin) {
  const existing = await env.DB.prepare('SELECT id FROM issues WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);
  await env.DB.prepare('DELETE FROM issues WHERE id = ?').bind(id).run();
  return ok({ message: 'Issue deleted' }, 200, origin);
}

async function importIssues(request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
  if (body.length === 0) return ok({ message: 'No data to import' }, 200, origin);

  const stmts = [];
  for (const item of body) {
    if (!item.report_date || !item.complaint || !item.category) continue;
    
    let day_count = null;
    if (item.completion_date && item.report_date) {
      const d1 = new Date(item.report_date);
      const d2 = new Date(item.completion_date);
      day_count = Math.floor((d2 - d1) / (1000 * 60 * 60 * 24));
    }

    stmts.push(
      env.DB.prepare(
        `INSERT INTO issues (report_date, branch_id, category, source, complaint, employee_name, fc_specialist, solution, status, completion_date, day_count) 
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(
        item.report_date,
        item.branch_id || null,
        item.category,
        item.source || null,
        item.complaint,
        item.employee_name || null,
        item.fc_specialist || null,
        item.solution || null,
        item.status !== null && status !== undefined && status !== '' ? status : '',
        item.completion_date || null,
        day_count
      )
    );
  }

  try {
    if (stmts.length > 0) await env.DB.batch(stmts);
    return ok({ message: `Berhasil mengimport ${stmts.length} data permasalahan` }, 200, origin);
  } catch (err) {
    return error('Gagal import data: ' + err.message, 500, origin);
  }
}
