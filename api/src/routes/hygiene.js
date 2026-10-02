// Master Hygiene Standards - Reference Data API
import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination } from '../utils/pagination.js';

export async function handleHygieneStandards(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace('/api/hygiene-standards', '');
  const idMatch = path.match(/^\/(\d+)$/);

  // GET list is accessible to all authenticated users (and public read if needed)
  if (request.method === 'GET' && path === '') {
    return listHygieneStandards(request, env, origin);
  }

  // GET list of unique rooms for filter dropdown
  if (request.method === 'GET' && path === '/rooms') {
    return listHygieneRooms(env, origin);
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);

  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') {
      return getHygieneStandard(id, env, origin);
    }
    if (request.method === 'PUT') {
      if (!hasPermission(user, 'hygiene_standards', 'write')) return forbidden(origin);
      return updateHygieneStandard(id, request, env, origin);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'hygiene_standards', 'delete')) return forbidden(origin);
      return deleteHygieneStandard(id, env, origin);
    }
  }

  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'hygiene_standards', 'write')) return forbidden(origin);
    return createHygieneStandard(request, env, origin);
  }

  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'hygiene_standards', 'write')) return forbidden(origin);
    return importHygieneStandards(request, env, origin);
  }

  return error('Not found', 404, origin);
}

async function listHygieneStandards(request, env, origin) {
  const { page, limit, offset } = getPagination(request.url);
  const url = new URL(request.url);
  const search = url.searchParams.get('search') || '';
  const roomName = url.searchParams.get('room_name') || '';
  const all = url.searchParams.get('all');

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

  if (all === '1') {
    const rows = await env.DB.prepare(
      `SELECT * FROM master_hygiene_standards ${where} ORDER BY id ASC`
    ).bind(...values).all();
    return ok(rows.results, 200, origin);
  }

  const [countResult, rows] = await Promise.all([
    env.DB.prepare(`SELECT COUNT(*) as total FROM master_hygiene_standards ${where}`).bind(...values).first(),
    env.DB.prepare(`SELECT * FROM master_hygiene_standards ${where} ORDER BY id ASC LIMIT ? OFFSET ?`).bind(...values, limit, offset).all()
  ]);

  return paginated(rows.results || [], countResult?.total || 0, page, limit, origin);
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

async function createHygieneStandard(request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }

  const { room_name, item_name, cleanliness_standard } = body;
  if (!room_name || !String(room_name).trim()) return error('Ruangan (room_name) wajib diisi', 400, origin);
  if (!item_name || !String(item_name).trim()) return error('Item/Objek (item_name) wajib diisi', 400, origin);
  if (!cleanliness_standard || !String(cleanliness_standard).trim()) return error('Standar Kebersihan (cleanliness_standard) wajib diisi', 400, origin);

  const res = await env.DB.prepare(
    'INSERT INTO master_hygiene_standards (room_name, item_name, cleanliness_standard) VALUES (?, ?, ?)'
  ).bind(String(room_name).trim(), String(item_name).trim(), String(cleanliness_standard).trim()).run();

  return ok({ id: res.meta.last_row_id, message: 'Parameter kebersihan berhasil ditambahkan' }, 201, origin);
}

async function updateHygieneStandard(id, request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }

  const existing = await env.DB.prepare('SELECT id FROM master_hygiene_standards WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);

  const { room_name, item_name, cleanliness_standard } = body;

  await env.DB.prepare(
    `UPDATE master_hygiene_standards
     SET room_name = COALESCE(?, room_name),
         item_name = COALESCE(?, item_name),
         cleanliness_standard = COALESCE(?, cleanliness_standard),
         updated_at = datetime('now')
     WHERE id = ?`
  ).bind(
    room_name !== undefined && String(room_name).trim() ? String(room_name).trim() : null,
    item_name !== undefined && String(item_name).trim() ? String(item_name).trim() : null,
    cleanliness_standard !== undefined && String(cleanliness_standard).trim() ? String(cleanliness_standard).trim() : null,
    id
  ).run();

  return ok({ message: 'Parameter kebersihan berhasil diperbarui' }, 200, origin);
}

async function deleteHygieneStandard(id, env, origin) {
  const existing = await env.DB.prepare('SELECT id FROM master_hygiene_standards WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);

  await env.DB.prepare('DELETE FROM master_hygiene_standards WHERE id = ?').bind(id).run();
  return ok({ message: 'Parameter kebersihan berhasil dihapus' }, 200, origin);
}

async function importHygieneStandards(request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!Array.isArray(body)) return error('Payload harus berupa array', 400, origin);
  if (body.length === 0) return ok({ message: 'Tidak ada data untuk diimpor' }, 200, origin);

  const existing = await env.DB.prepare('SELECT id, room_name, item_name FROM master_hygiene_standards').all();
  const existingMap = new Map();
  (existing.results || []).forEach(r => {
    const key = `${(r.room_name || '').toLowerCase().trim()}|||${(r.item_name || '').toLowerCase().trim()}`;
    existingMap.set(key, r.id);
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
      const id = existingMap.get(key);
      stmts.push(env.DB.prepare(
        "UPDATE master_hygiene_standards SET cleanliness_standard = ?, updated_at = datetime('now') WHERE id = ?"
      ).bind(standard, id));
      updated++;
    } else {
      stmts.push(env.DB.prepare(
        'INSERT INTO master_hygiene_standards (room_name, item_name, cleanliness_standard) VALUES (?, ?, ?)'
      ).bind(room, obj, standard));
      inserted++;
    }
  }

  if (stmts.length > 0) {
    // Execute in chunks of 50 to avoid D1 batch limit
    const chunkSize = 50;
    for (let i = 0; i < stmts.length; i += chunkSize) {
      await env.DB.batch(stmts.slice(i, i + chunkSize));
    }
  }

  return ok({ message: 'Import berhasil', inserted, updated, total: body.length }, 200, origin);
}
