import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination, resolveSortOrder } from '../utils/pagination.js';
import { requireExportAccess, getExportPagination } from '../utils/export.js';
import { buildAuditStatement, buildCreateAuditStatement, executeAtomicImportBatch } from '../utils/audit.js';

const TRAINING_SORT_MAP = Object.freeze({
  date_desc: 't.training_date DESC, t.id ASC',
  date_asc: 't.training_date ASC, t.id ASC',
  score_desc: 'COALESCE(t.score, -1) DESC, t.training_date DESC, t.id ASC',
});

export async function handleTraining(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace('/api/training', '');

  if (request.method === 'GET' && path === '/export') {
    const { response } = await requireExportAccess(request, env, origin, 'training');
    if (response) return response;
    return exportTraining(request, env, origin);
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'training', 'read')) return forbidden(origin);

  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'GET' && path === '') return list(request, env, origin);
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'training', 'write')) return forbidden(origin);
    return create(request, env, origin, user);
  }
  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'training', 'import')) return forbidden(origin);
    let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
    if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
    if (body.length === 0) return ok({ message: 'No data to import' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });

    const existing = await env.DB.prepare('SELECT * FROM training').all();
    const existingMap = new Map();
    (existing.results || []).forEach(r => {
      if (r.training_date && r.subject && r.branch_id) {
        existingMap.set(r.training_date + '_' + r.subject.toLowerCase().trim() + '_' + r.branch_id, r);
      }
    });

    const stmts = [];
    let inserted = 0; let updated = 0;
    for (const item of body) {
      if (!item.training_date || !item.subject || !item.branch_id) continue;
      const key = item.training_date + '_' + item.subject.toLowerCase().trim() + '_' + item.branch_id;
      const participantsArray = item.participants ? item.participants.split(',').map(p => p.trim()).filter(Boolean) : [];
      const participantsStr = JSON.stringify(participantsArray);
      const scoreVal = item.score !== null ? item.score : null;

      if (existingMap.has(key)) {
        const oldRow = existingMap.get(key);
        const afterRow = {
          ...oldRow,
          batch: item.batch || null,
          trainer: item.trainer || null,
          participants: participantsStr,
          score: scoreVal,
          document_link: item.document_link || null
        };
        stmts.push(
          env.DB.prepare(`UPDATE training SET batch=?, trainer=?, participants=?, score=?, document_link=?, updated_at=datetime('now') WHERE id=?`)
            .bind(item.batch || null, item.trainer || null, participantsStr, scoreVal, item.document_link || null, oldRow.id),
          buildAuditStatement(env, user, 'UPDATE', 'training', oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true)
        );
        existingMap.set(key, afterRow);
        updated++;
      } else {
        const newSnapshot = {
          training_date: item.training_date,
          batch: item.batch || null,
          subject: item.subject,
          branch_id: item.branch_id,
          trainer: item.trainer || null,
          participants: participantsStr,
          score: scoreVal,
          document_link: item.document_link || null
        };
        stmts.push(
          env.DB.prepare(`INSERT INTO training (training_date, batch, subject, branch_id, trainer, participants, score, document_link) VALUES (?, ?, ?, ?, ?, ?, ?, ?)`)
            .bind(item.training_date, item.batch || null, item.subject, item.branch_id, item.trainer || null, participantsStr, scoreVal, item.document_link || null),
          buildCreateAuditStatement(env, user, 'training', 'new', newSnapshot, { source: 'import' })
        );
        inserted++;
      }
    }
    try {
      if (stmts.length > 0) await executeAtomicImportBatch(env, stmts);
      return ok({ message: 'Import sukses', inserted, updated, total: body.length }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
    } catch (e) {
      console.error('Import error:', e); return error('Gagal mengimport data', 500, origin);
    }
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') return getOne(id, env, origin);
    if (request.method === 'PUT' || request.method === 'PATCH') {
      if (!hasPermission(user, 'training', 'write')) return forbidden(origin);
      return update(id, request, env, origin, user);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'training', 'delete')) return forbidden(origin);
      return remove(id, env, origin, user);
    }
  }
  return error('Not found', 404, origin);
}

let _legacyDatesChecked = false;

function buildTrainingFilter(url) {
  const search = (url.searchParams.get('search') || '').trim();
  const year = (url.searchParams.get('year') || '').trim();
  const batch = (url.searchParams.get('batch') || '').trim();
  const branch_id = (url.searchParams.get('branch_id') || '').trim();
  const trainer = (url.searchParams.get('trainer') || '').trim();
  const sort = url.searchParams.get('sort') || 'date_desc';

  const conditions = [];
  const values = [];
  if (search) {
    conditions.push('(t.subject LIKE ? OR t.trainer LIKE ? OR t.participants LIKE ? OR t.batch LIKE ? OR b.full_name LIKE ? OR b.name LIKE ?)');
    values.push(`%${search}%`, `%${search}%`, `%${search}%`, `%${search}%`, `%${search}%`, `%${search}%`);
  }
  if (year) { conditions.push("strftime('%Y', t.training_date) = ?"); values.push(year); }
  if (batch) { conditions.push("t.batch = ?"); values.push(batch); }
  if (branch_id) { conditions.push("t.branch_id = ?"); values.push(branch_id); }
  if (trainer) { conditions.push("t.trainer = ?"); values.push(trainer); }

  const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';
  const countJoin = search ? 'LEFT JOIN branches b ON t.branch_id = b.id' : '';
  const orderBy = resolveSortOrder(sort, TRAINING_SORT_MAP, 'date_desc');
  return { where, values, countJoin, orderBy };
}

function countUniqueTrainingParticipants(rows) {
  const participantSet = new Set();
  for (const row of rows || []) {
    if (!row?.participants) continue;
    let items = [];
    if (Array.isArray(row.participants)) {
      items = row.participants;
    } else {
      const str = String(row.participants).trim();
      if (str.startsWith('[')) {
        try {
          const parsed = JSON.parse(str);
          items = Array.isArray(parsed) ? parsed : str.split(',');
        } catch {
          items = str.split(',');
        }
      } else {
        items = str.split(',');
      }
    }
    for (const rawName of items) {
      const trimmed = String(rawName || '').trim();
      const lower = trimmed.toLowerCase();
      if (trimmed && trimmed !== '-' && lower !== 'null' && lower !== 'umum' && lower !== 'semua' && lower !== 'n/a') {
        participantSet.add(trimmed.toUpperCase());
      }
    }
  }
  return participantSet.size;
}

async function list(request, env, origin) {
  // Auto-migrate legacy 2026-08-10 training dates to correct batch start dates (once per isolate)
  if (!_legacyDatesChecked) {
    try {
      const checkBadDates = await env.DB.prepare("SELECT COUNT(*) as c FROM training WHERE training_date = '2026-08-10'").first();
      if (checkBadDates && checkBadDates.c > 0) {
        await env.DB.batch([
          env.DB.prepare("UPDATE training SET training_date = '2026-06-09', updated_at = datetime('now') WHERE batch = 'Batch 1' AND training_date = '2026-08-10'"),
          env.DB.prepare("UPDATE training SET training_date = '2026-06-12', updated_at = datetime('now') WHERE (batch = 'Batch 2' OR participants LIKE '%DEDY SUPANDI%') AND training_date = '2026-08-10'"),
          env.DB.prepare("UPDATE training SET training_date = '2026-06-17', updated_at = datetime('now') WHERE batch = 'Batch 3' AND participants NOT LIKE '%DEDY SUPANDI%' AND training_date = '2026-08-10'"),
          env.DB.prepare("UPDATE training SET training_date = '2026-06-24', updated_at = datetime('now') WHERE batch = 'Batch 4' AND training_date = '2026-08-10'"),
          env.DB.prepare("UPDATE training SET training_date = '2026-06-29', updated_at = datetime('now') WHERE batch = 'Batch 5' AND training_date = '2026-08-10'")
        ]);
      }
      _legacyDatesChecked = true;
    } catch (e) {
      console.error('Training date auto-migration error:', e);
    }
  }

  const url = new URL(request.url);
  const { page, limit, offset } = getPagination(request.url);
  const { where, values, countJoin, orderBy } = buildTrainingFilter(url);
  const includeStats = url.searchParams.get('include_stats') === '1';

  const queries = [
    env.DB.prepare(`SELECT COUNT(*) as total FROM training t ${countJoin} ${where}`).bind(...values).first(),
    env.DB.prepare(
      `SELECT t.*, b.full_name as branch_name FROM training t
       LEFT JOIN branches b ON t.branch_id = b.id
       ${where} ORDER BY ${orderBy} LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all(),
  ];

  if (includeStats) {
    queries.push(
      env.DB.prepare(
        `SELECT
           COUNT(*) as total_sessions,
           COUNT(DISTINCT NULLIF(TRIM(batch), '')) as total_batches,
           SUM(CASE WHEN strftime('%Y-%m', training_date) = strftime('%Y-%m', date('now', '+7 hours')) THEN 1 ELSE 0 END) as current_month
         FROM training`
      ).first(),
      env.DB.prepare('SELECT participants FROM training WHERE participants IS NOT NULL').all()
    );
  }

  const [countResult, rows, statsRow, participantsRes] = await Promise.all(queries);
  const extra = includeStats ? {
    stats: {
      totalSessions: Number(statsRow?.total_sessions || 0),
      totalBatches: Number(statsRow?.total_batches || 0),
      currentMonth: Number(statsRow?.current_month || 0),
      totalParticipants: countUniqueTrainingParticipants(participantsRes?.results),
    }
  } : null;

  return paginated(rows.results || [], countResult?.total || 0, page, limit, origin, env?.ENVIRONMENT, extra);
}

async function exportTraining(request, env, origin) {
  const url = new URL(request.url);
  const { page, limit, offset } = getExportPagination(request.url);
  const { where, values, countJoin, orderBy } = buildTrainingFilter(url);

  const [countResult, rows] = await Promise.all([
    env.DB.prepare(`SELECT COUNT(*) as total FROM training t ${countJoin} ${where}`).bind(...values).first(),
    env.DB.prepare(
      `SELECT t.*, b.full_name as branch_name FROM training t
       LEFT JOIN branches b ON t.branch_id = b.id
       ${where} ORDER BY ${orderBy} LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all()
  ]);

  return paginated(rows.results || [], countResult?.total || 0, page, limit, origin, env?.ENVIRONMENT);
}

async function getOne(id, env, origin) {
  const row = await env.DB.prepare(
    'SELECT t.*, b.full_name as branch_name FROM training t LEFT JOIN branches b ON t.branch_id = b.id WHERE t.id = ?'
  ).bind(id).first();
  if (!row) return notFound(origin);
  return ok(row, 200, origin);
}

async function validateTrainingBranch(env, branch_id) {
  if (branch_id === undefined || branch_id === null || branch_id === '') return { valid: true, branchId: null };
  const parsed = Number(branch_id);
  if (!Number.isInteger(parsed) || parsed <= 0) return { valid: false, error: 'branch_id tidak valid' };
  const row = await env.DB.prepare('SELECT id FROM branches WHERE id = ? AND deleted_at IS NULL').bind(parsed).first();
  if (!row) return { valid: false, error: 'Cabang (branch_id) tidak ditemukan atau sudah dinonaktifkan' };
  return { valid: true, branchId: parsed };
}

export const OFFICIAL_TRAINING_BATCHES = Object.freeze([
  'Batch 1',
  'Batch 2',
  'Batch 3',
  'Batch 4',
  'Batch 5',
]);

async function create(request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
  const { training_date, batch, subject, participants, branch_id, trainer, score, notes, document_link } = body;
  if (!training_date || !subject) return error('training_date and subject required', 400, origin);

  let cleanBatch = String(batch ?? '').trim();
  if (/^\d+$/.test(cleanBatch)) {
    cleanBatch = `Batch ${cleanBatch}`;
  }
  if (!OFFICIAL_TRAINING_BATCHES.includes(cleanBatch)) {
    return error(`Batch pelatihan tidak valid. Pilihan resmi: ${OFFICIAL_TRAINING_BATCHES.join(', ')}`, 400, origin);
  }

  const branchCheck = await validateTrainingBranch(env, branch_id);
  if (!branchCheck.valid) return error(branchCheck.error, 400, origin);

  let cleanScore = null;
  if (score !== undefined && score !== null && score !== '') {
    const parsedScore = Number(score);
    if (!Number.isFinite(parsedScore) || parsedScore < 0 || parsedScore > 100) {
      return error('score harus berupa angka antara 0 dan 100', 400, origin);
    }
    cleanScore = parsedScore;
  }

  const participantsStr = Array.isArray(participants) ? JSON.stringify(participants) : (participants || null);
  const newSnapshot = {
    training_date,
    batch: cleanBatch,
    subject,
    participants: participantsStr,
    branch_id: branch_id || null,
    trainer: trainer || null,
    score: cleanScore,
    notes: notes || null,
    document_link: document_link || null
  };

  const batchRes = await env.DB.batch([
    env.DB.prepare(
      'INSERT INTO training (training_date, batch, subject, participants, branch_id, trainer, score, notes, document_link) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)'
    ).bind(training_date, cleanBatch, subject, participantsStr, branch_id || null,
      trainer || null, cleanScore, notes || null, document_link || null),
    buildCreateAuditStatement(env, user, 'training', 'new', newSnapshot)
  ]);

  return ok({ id: batchRes?.[0]?.meta?.last_row_id }, 201, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function update(id, request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
  const existing = await env.DB.prepare('SELECT * FROM training WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);

  const { training_date, batch, subject, participants, branch_id, trainer, score, notes, document_link } = body;
  let cleanBatch = null;
  if (batch !== undefined) {
    cleanBatch = String(batch ?? '').trim();
    if (/^\d+$/.test(cleanBatch)) {
      cleanBatch = `Batch ${cleanBatch}`;
    }
    if (
      !cleanBatch ||
      (!OFFICIAL_TRAINING_BATCHES.includes(cleanBatch) && cleanBatch !== String(existing.batch || '').trim())
    ) {
      return error(`Batch pelatihan tidak valid. Pilihan resmi: ${OFFICIAL_TRAINING_BATCHES.join(', ')}`, 400, origin);
    }
  }
  const branchCheck = await validateTrainingBranch(env, branch_id);
  if (!branchCheck.valid) return error(branchCheck.error, 400, origin);
  let effScore = existing.score;
  if (Object.prototype.hasOwnProperty.call(body, 'score')) {
    if (score === null || score === '') {
      effScore = null;
    } else {
      const parsedScore = Number(score);
      if (!Number.isFinite(parsedScore) || parsedScore < 0 || parsedScore > 100) {
        return error('score harus berupa angka antara 0 dan 100', 400, origin);
      }
      effScore = parsedScore;
    }
  }
  const participantsStr = Array.isArray(participants) ? JSON.stringify(participants) : (participants || null);
  const afterSnapshot = {
    ...existing,
    training_date: training_date || existing.training_date,
    batch: cleanBatch || existing.batch,
    subject: subject || existing.subject,
    participants: participantsStr || existing.participants,
    branch_id: branch_id || existing.branch_id,
    trainer: trainer || existing.trainer,
    score: effScore,
    notes: notes || existing.notes,
    document_link: document_link || existing.document_link
  };

  const updateStmt = env.DB.prepare(
    `UPDATE training SET training_date = COALESCE(?, training_date), batch = COALESCE(?, batch),
     subject = COALESCE(?, subject), participants = COALESCE(?, participants),
     branch_id = COALESCE(?, branch_id), trainer = COALESCE(?, trainer),
     score = ?, notes = COALESCE(?, notes),
     document_link = COALESCE(?, document_link), updated_at = datetime('now') WHERE id = ?`
  ).bind(training_date || null, cleanBatch || null, subject || null, participantsStr,
    branch_id || null, trainer || null, effScore, notes || null, document_link || null, id);

  await env.DB.batch([
    updateStmt,
    buildAuditStatement(env, user, 'UPDATE', 'training', id, existing, afterSnapshot, null, 'SUCCESS', true)
  ]);

  return ok({ message: 'Updated' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function remove(id, env, origin, user = null) {
  const existing = await env.DB.prepare('SELECT * FROM training WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);
  const delStmt = env.DB.prepare('DELETE FROM training WHERE id = ?').bind(id);
  const { executeAtomicDelete } = await import('../utils/audit.js');
  await executeAtomicDelete(env, user, delStmt, 'training', id, existing);
  return ok({ message: 'Deleted' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
}
