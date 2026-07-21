import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination } from '../utils/pagination.js';

export async function handleOneOnOne(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'one_on_one', 'read')) return forbidden(origin);

  const url = new URL(request.url);
  const path = url.pathname.replace('/api/one-on-one', '');
  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'GET' && path === '') return list(request, env, origin);
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'one_on_one', 'write')) return forbidden(origin);
    return create(request, env, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') return getOne(id, env, origin);
    if (request.method === 'PUT') {
      if (!hasPermission(user, 'one_on_one', 'write')) return forbidden(origin);
      return update(id, request, env, origin);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'one_on_one', 'delete')) return forbidden(origin);
      return remove(id, env, origin);
    }
  }
  return error('Not found', 404, origin);
}

async function list(request, env, origin) {
  const { page, limit, offset } = getPagination(request.url);
  const url = new URL(request.url);
  const branch_id = url.searchParams.get('branch_id') || '';
  const status = url.searchParams.get('status') || '';
  const search = url.searchParams.get('search') || '';

  let conditions = [];
  let values = [];
  if (branch_id) { conditions.push('o.branch_id = ?'); values.push(branch_id); }
  if (status) { conditions.push('o.status = ?'); values.push(status); }
  if (search) { conditions.push('(o.employee_name LIKE ? OR o.problem LIKE ?)'); values.push(`%${search}%`, `%${search}%`); }

  const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';

  const [countResult, rows] = await Promise.all([
    env.DB.prepare(`SELECT COUNT(*) as total FROM one_on_one o ${where}`).bind(...values).first(),
    env.DB.prepare(
      `SELECT o.*, b.full_name as branch_name FROM one_on_one o
       LEFT JOIN branches b ON o.branch_id = b.id
       ${where} ORDER BY o.meeting_date DESC LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all()
  ]);

  return paginated(rows.results, countResult.total, page, limit, origin);
}

async function getOne(id, env, origin) {
  const row = await env.DB.prepare(
    'SELECT o.*, b.full_name as branch_name FROM one_on_one o LEFT JOIN branches b ON o.branch_id = b.id WHERE o.id = ?'
  ).bind(id).first();
  if (!row) return notFound(origin);
  return ok(row, 200, origin);
}

async function create(request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const { meeting_date, branch_id, employee_name, pic, problem, solution, status, completion_date, document_link } = body;
  if (!meeting_date || !employee_name || !problem) return error('meeting_date, employee_name, problem required', 400, origin);

  let day_count = null;
  if (completion_date && meeting_date) {
    day_count = Math.floor((new Date(completion_date) - new Date(meeting_date)) / (1000 * 60 * 60 * 24));
  }

  const result = await env.DB.prepare(
    'INSERT INTO one_on_one (meeting_date, branch_id, employee_name, pic, problem, solution, status, completion_date, day_count, document_link) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'
  ).bind(meeting_date, branch_id || null, employee_name, pic || null, problem,
    solution || null, status !== null && status !== undefined && status !== '' ? status : '', completion_date || null, day_count, document_link || null).run();

  return ok({ id: result.meta.last_row_id }, 201, origin);
}

async function update(id, request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const existing = await env.DB.prepare('SELECT * FROM one_on_one WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);

  const { meeting_date, branch_id, employee_name, pic, problem, solution, status, completion_date, document_link } = body;
  const rd = meeting_date || existing.meeting_date;
  const cd = completion_date !== undefined ? completion_date : existing.completion_date;
  let day_count = existing.day_count;
  if (rd && cd) day_count = Math.floor((new Date(cd) - new Date(rd)) / (1000 * 60 * 60 * 24));

  await env.DB.prepare(
    `UPDATE one_on_one SET meeting_date = COALESCE(?, meeting_date), branch_id = COALESCE(?, branch_id),
     employee_name = COALESCE(?, employee_name), pic = COALESCE(?, pic), problem = COALESCE(?, problem),
     solution = COALESCE(?, solution), status = COALESCE(?, status), completion_date = COALESCE(?, completion_date),
     day_count = ?, document_link = COALESCE(?, document_link), updated_at = datetime('now') WHERE id = ?`
  ).bind(meeting_date || null, branch_id || null, employee_name || null, pic || null, problem || null,
    solution || null, status !== null && status !== undefined && status !== '' ? status : '', completion_date || null, day_count, document_link || null, id).run();

  return ok({ message: 'Updated' }, 200, origin);
}

async function remove(id, env, origin) {
  const existing = await env.DB.prepare('SELECT id FROM one_on_one WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);
  await env.DB.prepare('DELETE FROM one_on_one WHERE id = ?').bind(id).run();
  return ok({ message: 'Deleted' }, 200, origin);
}
