import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination } from '../utils/pagination.js';

export async function handleTraining(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'training', 'read')) return forbidden(origin);

  const url = new URL(request.url);
  const path = url.pathname.replace('/api/training', '');
  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'GET' && path === '') return list(request, env, origin);
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'training', 'write')) return forbidden(origin);
    return create(request, env, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') return getOne(id, env, origin);
    if (request.method === 'PUT') {
      if (!hasPermission(user, 'training', 'write')) return forbidden(origin);
      return update(id, request, env, origin);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'training', 'delete')) return forbidden(origin);
      return remove(id, env, origin);
    }
  }
  return error('Not found', 404, origin);
}

async function list(request, env, origin) {
  const { page, limit, offset } = getPagination(request.url);
  const url = new URL(request.url);
  const search = url.searchParams.get('search') || '';
  const year = url.searchParams.get('year') || '';

  let conditions = [];
  let values = [];
  if (search) { conditions.push('(t.subject LIKE ? OR t.trainer LIKE ? OR t.participants LIKE ?)'); values.push(`%${search}%`, `%${search}%`, `%${search}%`); }
  if (year) { conditions.push("strftime('%Y', t.training_date) = ?"); values.push(year); }

  const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';

  const [countResult, rows] = await Promise.all([
    env.DB.prepare(`SELECT COUNT(*) as total FROM training t ${where}`).bind(...values).first(),
    env.DB.prepare(
      `SELECT t.*, b.full_name as branch_name FROM training t
       LEFT JOIN branches b ON t.branch_id = b.id
       ${where} ORDER BY t.training_date DESC LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all()
  ]);

  return paginated(rows.results, countResult.total, page, limit, origin);
}

async function getOne(id, env, origin) {
  const row = await env.DB.prepare(
    'SELECT t.*, b.full_name as branch_name FROM training t LEFT JOIN branches b ON t.branch_id = b.id WHERE t.id = ?'
  ).bind(id).first();
  if (!row) return notFound(origin);
  return ok(row, 200, origin);
}

async function create(request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const { training_date, batch, subject, participants, branch_id, trainer, score, notes, document_link } = body;
  if (!training_date || !subject) return error('training_date and subject required', 400, origin);

  const participantsStr = Array.isArray(participants) ? JSON.stringify(participants) : (participants || null);

  const result = await env.DB.prepare(
    'INSERT INTO training (training_date, batch, subject, participants, branch_id, trainer, score, notes, document_link) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)'
  ).bind(training_date, batch || null, subject, participantsStr, branch_id || null,
    trainer || null, score || null, notes || null, document_link || null).run();

  return ok({ id: result.meta.last_row_id }, 201, origin);
}

async function update(id, request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const existing = await env.DB.prepare('SELECT id FROM training WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);

  const { training_date, batch, subject, participants, branch_id, trainer, score, notes, document_link } = body;
  const participantsStr = Array.isArray(participants) ? JSON.stringify(participants) : (participants || null);

  await env.DB.prepare(
    `UPDATE training SET training_date = COALESCE(?, training_date), batch = COALESCE(?, batch),
     subject = COALESCE(?, subject), participants = COALESCE(?, participants),
     branch_id = COALESCE(?, branch_id), trainer = COALESCE(?, trainer),
     score = COALESCE(?, score), notes = COALESCE(?, notes),
     document_link = COALESCE(?, document_link), updated_at = datetime('now') WHERE id = ?`
  ).bind(training_date || null, batch || null, subject || null, participantsStr,
    branch_id || null, trainer || null, score || null, notes || null, document_link || null, id).run();

  return ok({ message: 'Updated' }, 200, origin);
}

async function remove(id, env, origin) {
  const existing = await env.DB.prepare('SELECT id FROM training WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);
  await env.DB.prepare('DELETE FROM training WHERE id = ?').bind(id).run();
  return ok({ message: 'Deleted' }, 200, origin);
}
