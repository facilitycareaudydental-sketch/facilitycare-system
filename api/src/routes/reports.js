import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination } from '../utils/pagination.js';

// Handles: /api/reports/inspection, /api/reports/cleaning, /api/reports/fogging, /api/reports/basecamp
export async function handleReports(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace('/api/reports', '');

  // Public supply requests
  if (path.startsWith('/supply')) return handleSupply(request, env, origin, path.replace('/supply', ''));

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'reports', 'read')) return forbidden(origin);

  if (path.startsWith('/inspection')) return handleInspection(request, env, user, origin, path.replace('/inspection', ''));
  if (path.startsWith('/cleaning')) return handleCleaning(request, env, user, origin, path.replace('/cleaning', ''));
  if (path.startsWith('/fogging')) return handleFogging(request, env, user, origin, path.replace('/fogging', ''));
  if (path.startsWith('/basecamp')) return handleBasecamp(request, env, user, origin, path.replace('/basecamp', ''));

  return error('Not found', 404, origin);
}

// ---- Generic CRUD helper ----
async function crudList(request, env, origin, table, joinClause = '', extraConditions = []) {
  const { page, limit, offset } = getPagination(request.url);
  const url = new URL(request.url);
  const branch_id = url.searchParams.get('branch_id') || '';
  const period = url.searchParams.get('period') || '';
  const status = url.searchParams.get('status') || '';
  const year = url.searchParams.get('year') || '';

  let conditions = [...extraConditions];
  let values = [];
  if (branch_id) { conditions.push('t.branch_id = ?'); values.push(branch_id); }
  if (period) { conditions.push('t.period = ?'); values.push(period); }
  if (status) { conditions.push('t.status = ?'); values.push(status); }
  if (year) {
    const dateField = table === 'basecamp_reports' ? 'info_date' : 
                      table === 'inspection_reports' ? 'inspection_date' : 'activity_date';
    conditions.push(`strftime('%Y', t.${dateField}) = ?`);
    values.push(year);
  }

  const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';

  const [countResult, rows] = await Promise.all([
    env.DB.prepare(`SELECT COUNT(*) as total FROM ${table} t ${joinClause} ${where}`).bind(...values).first(),
    env.DB.prepare(
      `SELECT t.*, b.full_name as branch_name FROM ${table} t
       LEFT JOIN branches b ON t.branch_id = b.id
       ${where} ORDER BY t.id DESC LIMIT ? OFFSET ?`
    ).bind(...values, limit, offset).all()
  ]);

  return paginated(rows.results, countResult.total, page, limit, origin);
}

// ---- INSPECTION ----
async function handleInspection(request, env, user, origin, path) {
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === 'GET' && path === '') return crudList(request, env, origin, 'inspection_reports');
  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'reports', 'write')) return forbidden(origin);
    return importInspection(request, env, origin);
  }
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'reports', 'write')) return forbidden(origin);
    return createInspection(request, env, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') {
      const row = await env.DB.prepare('SELECT t.*, b.full_name as branch_name FROM inspection_reports t LEFT JOIN branches b ON t.branch_id = b.id WHERE t.id = ?').bind(id).first();
      return row ? ok(row, 200, origin) : notFound(origin);
    }
    if (request.method === 'PUT') {
      if (!hasPermission(user, 'reports', 'write')) return forbidden(origin);
      return updateInspection(id, request, env, origin);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'reports', 'delete')) return forbidden(origin);
      await env.DB.prepare('DELETE FROM inspection_reports WHERE id = ?').bind(id).run();
      return ok({ message: 'Deleted' }, 200, origin);
    }
  }
  return error('Not found', 404, origin);
}

async function createInspection(request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const { branch_id, period, inspection_date, status, fc_score, spv_score, document_link, notes } = body;
  if (!period || !inspection_date) return error('period and inspection_date required', 400, origin);
  const result = await env.DB.prepare(
    'INSERT INTO inspection_reports (branch_id, period, inspection_date, status, fc_score, spv_score, document_link, notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?)'
  ).bind(branch_id || null, period, inspection_date, status !== null && status !== undefined && status !== '' ? status : '', fc_score || null, spv_score || null, document_link || null, notes || null).run();
  return ok({ id: result.meta.last_row_id }, 201, origin);
}

async function importInspection(request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
  if (body.length === 0) return ok({ message: 'No data to import' }, 200, origin);

  const existing = await env.DB.prepare('SELECT id, branch_id, period, inspection_date FROM inspection_reports').all();
  const existingMap = new Map();
  (existing.results || []).forEach(s => {
    if (s.branch_id && s.period && s.inspection_date) {
      existingMap.set(s.branch_id + '_' + s.period.toLowerCase().trim() + '_' + s.inspection_date, s.id);
    }
  });

  const stmts = [];
  let inserted = 0;
  let updated = 0;

  for (const item of body) {
    if (!item.branch_id || !item.period || !item.inspection_date) continue;
    
    const key = item.branch_id + '_' + item.period.toLowerCase().trim() + '_' + item.inspection_date;
    const status = item.status !== null && item.status !== undefined && item.status !== '' ? item.status : '';

    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      stmts.push(
        env.DB.prepare(
          `UPDATE inspection_reports SET fc_score = ?, spv_score = ?, status = ?, document_link = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(
          item.fc_score,
          item.spv_score,
          status,
          item.document_link || null,
          item.notes || null,
          id
        )
      );
      updated++;
    } else {
      stmts.push(
        env.DB.prepare(
          `INSERT INTO inspection_reports (branch_id, period, inspection_date, fc_score, spv_score, status, document_link, notes) 
           VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
        ).bind(
          item.branch_id,
          item.period,
          item.inspection_date,
          item.fc_score,
          item.spv_score,
          status,
          item.document_link || null,
          item.notes || null
        )
      );
      inserted++;
    }
  }

  try {
    if (stmts.length > 0) await env.DB.batch(stmts);
    return ok({ 
      message: `Berhasil mengimport data`, 
      inserted, 
      updated,
      failed: 0,
      total: body.length 
    }, 200, origin);
  } catch (err) {
    return error('Gagal import data: ' + err.message, 500, origin);
  }
}

async function updateInspection(id, request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const existing = await env.DB.prepare('SELECT id FROM inspection_reports WHERE id = ?').bind(id).first();
  if (!existing) return notFound(origin);
  const { branch_id, period, inspection_date, status, fc_score, spv_score, document_link, notes } = body;
  await env.DB.prepare(
    `UPDATE inspection_reports SET branch_id = COALESCE(?, branch_id), period = COALESCE(?, period),
     inspection_date = COALESCE(?, inspection_date), status = COALESCE(?, status),
     fc_score = COALESCE(?, fc_score), spv_score = COALESCE(?, spv_score),
     document_link = COALESCE(?, document_link), notes = COALESCE(?, notes),
     updated_at = datetime('now') WHERE id = ?`
  ).bind(branch_id || null, period || null, inspection_date || null, status !== null && status !== undefined && status !== '' ? status : '',
    fc_score || null, spv_score || null, document_link || null, notes || null, id).run();
  return ok({ message: 'Updated' }, 200, origin);
}

// ---- CLEANING ----
async function handleCleaning(request, env, user, origin, path) {
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === 'GET' && path === '') return crudList(request, env, origin, 'cleaning_reports');
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'reports', 'write')) return forbidden(origin);
    let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
    const { branch_id, activity_type, period, activity_date, status, document_link, notes } = body;
    if (!activity_type || !period || !activity_date) return error('activity_type, period, activity_date required', 400, origin);
    const result = await env.DB.prepare('INSERT INTO cleaning_reports (branch_id, activity_type, period, activity_date, status, document_link, notes) VALUES (?, ?, ?, ?, ?, ?, ?)').bind(branch_id || null, activity_type, period, activity_date, status !== null && status !== undefined && status !== '' ? status : '', document_link || null, notes || null).run();
    return ok({ id: result.meta.last_row_id }, 201, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') {
      const row = await env.DB.prepare('SELECT t.*, b.full_name as branch_name FROM cleaning_reports t LEFT JOIN branches b ON t.branch_id = b.id WHERE t.id = ?').bind(id).first();
      return row ? ok(row, 200, origin) : notFound(origin);
    }
    if (request.method === 'PUT') {
      if (!hasPermission(user, 'reports', 'write')) return forbidden(origin);
      let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
      const { branch_id, activity_type, period, activity_date, status, document_link, notes } = body;
      await env.DB.prepare(`UPDATE cleaning_reports SET branch_id = COALESCE(?, branch_id), activity_type = COALESCE(?, activity_type), period = COALESCE(?, period), activity_date = COALESCE(?, activity_date), status = COALESCE(?, status), document_link = COALESCE(?, document_link), notes = COALESCE(?, notes), updated_at = datetime('now') WHERE id = ?`).bind(branch_id || null, activity_type || null, period || null, activity_date || null, status !== null && status !== undefined && status !== '' ? status : '', document_link || null, notes || null, id).run();
      return ok({ message: 'Updated' }, 200, origin);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'reports', 'delete')) return forbidden(origin);
      await env.DB.prepare('DELETE FROM cleaning_reports WHERE id = ?').bind(id).run();
      return ok({ message: 'Deleted' }, 200, origin);
    }
  }
  return error('Not found', 404, origin);
}

// ---- FOGGING ----
async function handleFogging(request, env, user, origin, path) {
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === 'GET' && path === '') return crudList(request, env, origin, 'fogging_reports');
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'reports', 'write')) return forbidden(origin);
    let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
    const { branch_id, activity_type, period, activity_date, status, document_link, notes } = body;
    if (!period) return error('period required', 400, origin);
    const result = await env.DB.prepare('INSERT INTO fogging_reports (branch_id, activity_type, period, activity_date, status, document_link, notes) VALUES (?, ?, ?, ?, ?, ?, ?)').bind(branch_id || null, activity_type || 'Fogging', period, activity_date || null, status !== null && status !== undefined && status !== '' ? status : '', document_link || null, notes || null).run();
    return ok({ id: result.meta.last_row_id }, 201, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') {
      const row = await env.DB.prepare('SELECT t.*, b.full_name as branch_name FROM fogging_reports t LEFT JOIN branches b ON t.branch_id = b.id WHERE t.id = ?').bind(id).first();
      return row ? ok(row, 200, origin) : notFound(origin);
    }
    if (request.method === 'PUT') {
      if (!hasPermission(user, 'reports', 'write')) return forbidden(origin);
      let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
      const { branch_id, activity_type, period, activity_date, status, document_link, notes } = body;
      await env.DB.prepare(`UPDATE fogging_reports SET branch_id = COALESCE(?, branch_id), activity_type = COALESCE(?, activity_type), period = COALESCE(?, period), activity_date = COALESCE(?, activity_date), status = COALESCE(?, status), document_link = COALESCE(?, document_link), notes = COALESCE(?, notes), updated_at = datetime('now') WHERE id = ?`).bind(branch_id || null, activity_type || null, period || null, activity_date || null, status !== null && status !== undefined && status !== '' ? status : '', document_link || null, notes || null, id).run();
      return ok({ message: 'Updated' }, 200, origin);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'reports', 'delete')) return forbidden(origin);
      await env.DB.prepare('DELETE FROM fogging_reports WHERE id = ?').bind(id).run();
      return ok({ message: 'Deleted' }, 200, origin);
    }
  }
  return error('Not found', 404, origin);
}

// ---- BASECAMP ----
async function handleBasecamp(request, env, user, origin, path) {
  const idMatch = path.match(/^\/(\d+)$/);
  if (request.method === 'GET' && path === '') {
    const { page, limit, offset } = getPagination(request.url);
    const url = new URL(request.url);
    const branch_id = url.searchParams.get('branch_id') || '';
    const status = url.searchParams.get('status') || '';
    const search = url.searchParams.get('search') || '';
    let conditions = []; let values = [];
    if (branch_id) { conditions.push('t.branch_id = ?'); values.push(branch_id); }
    if (status) { conditions.push('t.status = ?'); values.push(status); }
    if (search) { conditions.push('(t.problem LIKE ? OR t.pic LIKE ?)'); values.push(`%${search}%`, `%${search}%`); }
    const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';
    const [countResult, rows] = await Promise.all([
      env.DB.prepare(`SELECT COUNT(*) as total FROM basecamp_reports t ${where}`).bind(...values).first(),
      env.DB.prepare(`SELECT t.*, b.full_name as branch_name FROM basecamp_reports t LEFT JOIN branches b ON t.branch_id = b.id ${where} ORDER BY t.info_date DESC LIMIT ? OFFSET ?`).bind(...values, limit, offset).all()
    ]);
    return paginated(rows.results, countResult.total, page, limit, origin);
  }
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'reports', 'write')) return forbidden(origin);
    let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
    const { branch_id, problem, pic, info_date, done_date, status, notes } = body;
    if (!problem || !info_date) return error('problem and info_date required', 400, origin);
    const result = await env.DB.prepare('INSERT INTO basecamp_reports (branch_id, problem, pic, info_date, done_date, status, notes) VALUES (?, ?, ?, ?, ?, ?, ?)').bind(branch_id || null, problem, pic || null, info_date, done_date || null, status !== null && status !== undefined && status !== '' ? status : '', notes || null).run();
    return ok({ id: result.meta.last_row_id }, 201, origin);
  }
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') {
      const row = await env.DB.prepare('SELECT t.*, b.full_name as branch_name FROM basecamp_reports t LEFT JOIN branches b ON t.branch_id = b.id WHERE t.id = ?').bind(id).first();
      return row ? ok(row, 200, origin) : notFound(origin);
    }
    if (request.method === 'PUT') {
      if (!hasPermission(user, 'reports', 'write')) return forbidden(origin);
      let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
      const { branch_id, problem, pic, info_date, done_date, status, notes } = body;
      await env.DB.prepare(`UPDATE basecamp_reports SET branch_id = COALESCE(?, branch_id), problem = COALESCE(?, problem), pic = COALESCE(?, pic), info_date = COALESCE(?, info_date), done_date = COALESCE(?, done_date), status = COALESCE(?, status), notes = COALESCE(?, notes), updated_at = datetime('now') WHERE id = ?`).bind(branch_id || null, problem || null, pic || null, info_date || null, done_date || null, status !== null && status !== undefined && status !== '' ? status : '', notes || null, id).run();
      return ok({ message: 'Updated' }, 200, origin);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'reports', 'delete')) return forbidden(origin);
      await env.DB.prepare('DELETE FROM basecamp_reports WHERE id = ?').bind(id).run();
      return ok({ message: 'Deleted' }, 200, origin);
    }
  }
  return error('Not found', 404, origin);
}

// ---- SUPPLY REQUESTS (public POST, auth required for GET/PUT/DELETE) ----
async function handleSupply(request, env, origin, path) {
  const idMatch = path.match(/^\/(\d+)$/);

  // Public: POST /api/reports/supply
  if (request.method === 'POST' && path === '') {
    let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
    const { submitter_name, branch_id, branch_name, tools_items, tools_quantity, chemical_items, chemical_quantity, additional_notes } = body;
    const result = await env.DB.prepare(
      'INSERT INTO supply_requests (submitter_name, branch_id, branch_name, tools_items, tools_quantity, chemical_items, chemical_quantity, additional_notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?)'
    ).bind(submitter_name || null, branch_id || null, branch_name || null,
      JSON.stringify(tools_items || []), tools_quantity || null,
      JSON.stringify(chemical_items || []), chemical_quantity || null,
      additional_notes || null).run();
    return ok({ id: result.meta.last_row_id, message: 'Request submitted successfully' }, 201, origin);
  }

  // Auth required for other operations
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'reports', 'read')) return forbidden(origin);

  if (request.method === 'GET' && path === '') {
    const { page, limit, offset } = getPagination(request.url);
    const url = new URL(request.url);
    const status = url.searchParams.get('status') || '';
    let conditions = []; let values = [];
    if (status) { conditions.push('r.status = ?'); values.push(status); }
    const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';
    const [countResult, rows] = await Promise.all([
      env.DB.prepare(`SELECT COUNT(*) as total FROM supply_requests r ${where}`).bind(...values).first(),
      env.DB.prepare(`SELECT r.*, b.full_name as branch_name_ref FROM supply_requests r LEFT JOIN branches b ON r.branch_id = b.id ${where} ORDER BY r.submitted_at DESC LIMIT ? OFFSET ?`).bind(...values, limit, offset).all()
    ]);
    return paginated(rows.results, countResult.total, page, limit, origin);
  }

  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') {
      const row = await env.DB.prepare('SELECT r.*, b.full_name as branch_name_ref FROM supply_requests r LEFT JOIN branches b ON r.branch_id = b.id WHERE r.id = ?').bind(id).first();
      return row ? ok(row, 200, origin) : notFound(origin);
    }
    if (request.method === 'PUT') {
      if (!hasPermission(user, 'reports', 'write')) return forbidden(origin);
      let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
      const { status, processed_by, submitter_name, branch_id, branch_name, tools_items, tools_quantity, chemical_items, chemical_quantity, additional_notes } = body;
      
      const t_items = tools_items ? (typeof tools_items === 'string' ? tools_items.split(',').map(s=>s.trim()) : tools_items) : null;
      const c_items = chemical_items ? (typeof chemical_items === 'string' ? chemical_items.split(',').map(s=>s.trim()) : chemical_items) : null;

      await env.DB.prepare(`
        UPDATE supply_requests SET 
          status = COALESCE(?, status), 
          processed_by = COALESCE(?, processed_by), 
          processed_at = CASE WHEN ? IS NOT NULL THEN datetime('now') ELSE processed_at END,
          submitter_name = COALESCE(?, submitter_name),
          branch_id = COALESCE(?, branch_id),
          branch_name = COALESCE(?, branch_name),
          tools_items = COALESCE(?, tools_items),
          tools_quantity = COALESCE(?, tools_quantity),
          chemical_items = COALESCE(?, chemical_items),
          chemical_quantity = COALESCE(?, chemical_quantity),
          additional_notes = COALESCE(?, additional_notes)
        WHERE id = ?
      `).bind(
        status !== null && status !== undefined && status !== '' ? status : '', processed_by || null, status !== null && status !== undefined && status !== '' ? status : '',
        submitter_name || null, branch_id || null, branch_name || null,
        t_items ? JSON.stringify(t_items) : null, tools_quantity || null,
        c_items ? JSON.stringify(c_items) : null, chemical_quantity || null,
        additional_notes || null,
        id
      ).run();
      return ok({ message: 'Updated' }, 200, origin);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'reports', 'delete')) return forbidden(origin);
      await env.DB.prepare('DELETE FROM supply_requests WHERE id = ?').bind(id).run();
      return ok({ message: 'Deleted' }, 200, origin);
    }
  }
  return error('Not found', 404, origin);
}
