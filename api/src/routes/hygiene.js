// Master Hygiene Standards - Reference Data API
import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination, resolveSortOrder, buildTimestampOrderBy } from '../utils/pagination.js';
import { requireExportAccess, getExportPagination } from '../utils/export.js';
import { SEED_HYGIENE_STANDARDS } from './hygiene_seeds.js';
import { buildAuditStatement, buildCreateAuditStatement, executeAtomicImportBatch } from '../utils/audit.js';

const HYGIENE_SORT_MAP = Object.freeze({
  created_desc: buildTimestampOrderBy('created_at', { direction: 'DESC', idDir: 'DESC' }),
  id_asc: 'id ASC',
});

let _initialized = false;
let _initPromise = null;

export function _resetInitializationForTesting() {
  _initialized = false;
  _initPromise = null;
}

async function ensureInitialized(env) {
  if (_initialized) return;

  if (!_initPromise) {
    _initPromise = (async () => {
      await env.DB.prepare(`
        CREATE TABLE IF NOT EXISTS master_hygiene_standards (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          room_name TEXT NOT NULL,
          item_name TEXT NOT NULL,
          cleanliness_standard TEXT NOT NULL,
          created_at TEXT NOT NULL DEFAULT (datetime('now')),
          updated_at TEXT NOT NULL DEFAULT (datetime('now'))
        )
      `).run();

      await env.DB.prepare(`CREATE INDEX IF NOT EXISTS idx_hygiene_room ON master_hygiene_standards(room_name)`).run();

      const countRow = await env.DB.prepare('SELECT COUNT(*) as cnt FROM master_hygiene_standards').first();
      if (!countRow || countRow.cnt === 0) {
        const stmts = SEED_HYGIENE_STANDARDS.map(item =>
          env.DB.prepare(
            'INSERT INTO master_hygiene_standards (room_name, item_name, cleanliness_standard) VALUES (?, ?, ?)'
          ).bind(item.room_name, item.item_name, item.cleanliness_standard)
        );
        for (let i = 0; i < stmts.length; i += 50) {
          await env.DB.batch(stmts.slice(i, i + 50));
        }
      }
      _initialized = true;
    })().finally(() => {
      _initPromise = null;
    });
  }

  try {
    await _initPromise;
  } catch (err) {
    console.warn('master_hygiene_standards initialization warning:', err.message);
  }
}

export async function handleHygieneStandards(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace('/api/hygiene-standards', '');
  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'GET' && path === '/export') {
    const { response } = await requireExportAccess(request, env, origin, 'hygiene_standards');
    if (response) return response;
    return listHygieneStandards(request, env, origin, { isExport: true });
  }

  // Phase 5.48: hygiene standards list is internal reference data.
  // Authentication + module-level read RBAC is mandatory (no public bypass).
  if (request.method === 'GET' && path === '') {
    const u = await authenticate(request, env);
    if (!u) return unauthorized(origin);
    if (!hasPermission(u, 'hygiene_standards', 'read')) return forbidden(origin);
    return listHygieneStandards(request, env, origin);
  }

  // GET list of unique rooms for filter dropdown
  if (request.method === 'GET' && path === '/rooms') {
    const u = await authenticate(request, env);
    if (!u) return unauthorized(origin);
    if (!hasPermission(u, 'hygiene_standards', 'read')) return forbidden(origin);
    return listHygieneRooms(env, origin);
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (request.method === 'GET' && !hasPermission(user, 'hygiene_standards', 'read')) return forbidden(origin);

  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') {
      return getHygieneStandard(id, env, origin);
    }
    if (request.method === 'PUT' || request.method === 'PATCH') {
      if (!hasPermission(user, 'hygiene_standards', 'write')) return forbidden(origin);
      return updateHygieneStandard(id, request, env, origin, user);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'hygiene_standards', 'delete')) return forbidden(origin);
      return deleteHygieneStandard(id, env, origin, user);
    }
  }

  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'hygiene_standards', 'write')) return forbidden(origin);
    return createHygieneStandard(request, env, origin, user);
  }

  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'hygiene_standards', 'import')) return forbidden(origin);
    return importHygieneStandards(request, env, origin, user);
  }

  return error('Not found', 404, origin);
}

async function listHygieneStandards(request, env, origin, { isExport = false } = {}) {
  const { page, limit, offset } = isExport ? getExportPagination(request.url) : getPagination(request.url);
  const url = new URL(request.url);
  const search = url.searchParams.get('search') || '';
  const roomName = url.searchParams.get('room_name') || '';
  const all = url.searchParams.get('all');
  const sort = url.searchParams.get('sort') || 'created_desc';
  const orderBy = resolveSortOrder(sort, HYGIENE_SORT_MAP, 'created_desc');

  let conditions = [];
  let values = [];

  if (search) {
    conditions.push('(room_name LIKE ? OR item_name LIKE ? OR cleanliness_standard LIKE ?)');
    values.push(`%${search}%`, `%${search}%`, `%${search}%`);
  }

  if (roomName) {
    conditions.push('room_name = ?');
    values.push(roomName);
  }

  const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';

  if (!isExport && all === '1') {
    const rows = await env.DB.prepare(
      `SELECT * FROM master_hygiene_standards ${where} ORDER BY ${orderBy}`
    ).bind(...values).all();
    return ok(rows.results, 200, origin);
  }

  const [countResult, rows] = await Promise.all([
    env.DB.prepare(`SELECT COUNT(*) as total FROM master_hygiene_standards ${where}`).bind(...values).first(),
    env.DB.prepare(`SELECT * FROM master_hygiene_standards ${where} ORDER BY ${orderBy} LIMIT ? OFFSET ?`).bind(...values, limit, offset).all()
  ]);

  return paginated(rows.results || [], countResult?.total || 0, page, limit, origin, env?.ENVIRONMENT);
}

async function listHygieneRooms(env, origin) {
  const rows = await env.DB.prepare(
    'SELECT DISTINCT room_name FROM master_hygiene_standards ORDER BY room_name ASC'
  ).all();
  const rooms = (rows.results || []).map(r => r.room_name).filter(Boolean);
  return ok(rooms, 200, origin);
}

async function getHygieneStandard(id, env, origin) {
  const row = await env.DB.prepare(
    'SELECT * FROM master_hygiene_standards WHERE id = ?'
  ).bind(id).first();
  if (!row) return notFound(origin);
  return ok(row, 200, origin);
}

async function createHygieneStandard(request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }

  const { room_name, item_name, cleanliness_standard } = body;
  if (!room_name || !String(room_name).trim()) return error('Ruangan (room_name) wajib diisi', 400, origin);
  if (!item_name || !String(item_name).trim()) return error('Item/Objek (item_name) wajib diisi', 400, origin);
  if (!cleanliness_standard || !String(cleanliness_standard).trim()) return error('Standar Kebersihan (cleanliness_standard) wajib diisi', 400, origin);

  const cleanRoom = String(room_name).trim();
  const cleanItem = String(item_name).trim();
  const cleanStandard = String(cleanliness_standard).trim();
  const newSnapshot = { room_name: cleanRoom, item_name: cleanItem, cleanliness_standard: cleanStandard };

  const batchRes = await env.DB.batch([
    env.DB.prepare(
      'INSERT INTO master_hygiene_standards (room_name, item_name, cleanliness_standard) VALUES (?, ?, ?)'
    ).bind(cleanRoom, cleanItem, cleanStandard),
    buildCreateAuditStatement(env, user, 'hygiene_standards', 'new', newSnapshot)
  ]);

  return ok({ id: batchRes?.[0]?.meta?.last_row_id, message: 'Parameter kebersihan berhasil ditambahkan' }, 201, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function updateHygieneStandard(id, request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }

  const existing = await env.DB.prepare('SELECT * FROM master_hygiene_standards WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);

  const { room_name, item_name, cleanliness_standard } = body;
  const nextRoom = room_name !== undefined && String(room_name).trim() ? String(room_name).trim() : null;
  const nextItem = item_name !== undefined && String(item_name).trim() ? String(item_name).trim() : null;
  const nextStandard = cleanliness_standard !== undefined && String(cleanliness_standard).trim() ? String(cleanliness_standard).trim() : null;
  const afterSnapshot = {
    ...existing,
    room_name: nextRoom ?? existing.room_name,
    item_name: nextItem ?? existing.item_name,
    cleanliness_standard: nextStandard ?? existing.cleanliness_standard
  };

  await env.DB.batch([
    env.DB.prepare(
      `UPDATE master_hygiene_standards
       SET room_name = COALESCE(?, room_name),
           item_name = COALESCE(?, item_name),
           cleanliness_standard = COALESCE(?, cleanliness_standard),
           updated_at = datetime('now')
       WHERE id = ?`
    ).bind(nextRoom, nextItem, nextStandard, id),
    buildAuditStatement(env, user, 'UPDATE', 'hygiene_standards', id, existing, afterSnapshot, null, 'SUCCESS', true)
  ]);

  return ok({ message: 'Parameter kebersihan berhasil diperbarui' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function deleteHygieneStandard(id, env, origin, user = null) {
  const existing = await env.DB.prepare('SELECT * FROM master_hygiene_standards WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);

  const delStmt = env.DB.prepare('DELETE FROM master_hygiene_standards WHERE id = ?').bind(id);
  const { executeAtomicDelete } = await import('../utils/audit.js');
  await executeAtomicDelete(env, user, delStmt, 'hygiene_standards', id, existing);
  return ok({ message: 'Parameter kebersihan berhasil dihapus' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function importHygieneStandards(request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!Array.isArray(body)) return error('Payload harus berupa array', 400, origin);
  if (body.length === 0) return ok({ message: 'Tidak ada data untuk diimpor' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });

  const existing = await env.DB.prepare('SELECT * FROM master_hygiene_standards').all();
  const existingMap = new Map();
  (existing.results || []).forEach(r => {
    const key = `${(r.room_name || '').toLowerCase().trim()}|||${(r.item_name || '').toLowerCase().trim()}`;
    existingMap.set(key, r);
  });

  const stmts = [];
  let inserted = 0;
  let updated = 0;

  for (const item of body) {
    const room = String(item.room_name || item['Ruangan'] || '').trim();
    const obj = String(item.item_name || item['Item/Objek'] || '').trim();
    const standard = String(item.cleanliness_standard || item['Standar Kebersihan'] || '').trim();

    if (!room || !obj || !standard) continue;

    const key = `${room.toLowerCase()}|||${obj.toLowerCase()}`;
    if (existingMap.has(key)) {
      const oldRow = existingMap.get(key);
      const afterRow = { ...oldRow, cleanliness_standard: standard };
      stmts.push(
        env.DB.prepare(
          "UPDATE master_hygiene_standards SET cleanliness_standard = ?, updated_at = datetime('now') WHERE id = ?"
        ).bind(standard, oldRow.id),
        buildAuditStatement(env, user, 'UPDATE', 'hygiene_standards', oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true)
      );
      existingMap.set(key, afterRow);
      updated++;
    } else {
      const newSnapshot = { room_name: room, item_name: obj, cleanliness_standard: standard };
      stmts.push(
        env.DB.prepare(
          'INSERT INTO master_hygiene_standards (room_name, item_name, cleanliness_standard) VALUES (?, ?, ?)'
        ).bind(room, obj, standard),
        buildCreateAuditStatement(env, user, 'hygiene_standards', 'new', newSnapshot, { source: 'import' })
      );
      inserted++;
    }
  }

  if (stmts.length > 0) {
    await executeAtomicImportBatch(env, stmts, 50);
  }

  return ok({ message: 'Import berhasil', inserted, updated, total: body.length }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
}
