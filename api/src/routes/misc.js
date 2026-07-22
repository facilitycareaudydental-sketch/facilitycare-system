// SOP, Master Checklist, Master Forms, PIC list
import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination } from '../utils/pagination.js';

export async function handleMisc(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname;

  if (path.startsWith('/api/sop')) return handleTable(request, env, origin, 'sop', path.replace('/api/sop', ''));
  if (path.startsWith('/api/checklist')) return handleTable(request, env, origin, 'master_checklist', path.replace('/api/checklist', ''));
  if (path.startsWith('/api/forms')) return handleForms(request, env, origin, path.replace('/api/forms', ''));
  if (path.startsWith('/api/pic')) return handlePic(request, env, origin);
  if (path.startsWith('/api/options')) return handleOptions(request, env, origin);
  
  if (path.startsWith('/api/test-emp')) {
    try {
      try { await env.DB.prepare('ALTER TABLE sp_data ADD COLUMN division TEXT').run(); } catch(e){}
      try { await env.DB.prepare('ALTER TABLE sp_data ADD COLUMN akhir_sp TEXT').run(); } catch(e){}
      return ok({ message: 'Migration applied' }, 200, origin);
    } catch (e) {
      return error(e.message, 500, origin);
    }
  }

  return error('Not found', 404, origin);
}

async function handleTable(request, env, origin, table, path) {
  // GET list is public for SOPs and checklists
  if (request.method === 'GET' && path === '') {
    const { page, limit, offset } = getPagination(request.url);
    const url = new URL(request.url);
    const search = url.searchParams.get('search') || '';
    const all = url.searchParams.get('all');
    if (all === '1') {
      const rows = await env.DB.prepare(`SELECT * FROM ${table} ORDER BY name`).all();
      return ok(rows.results, 200, origin);
    }
    let conditions = []; let values = [];
    if (search) { conditions.push('(name LIKE ? OR category LIKE ?)'); values.push(`%${search}%`, `%${search}%`); }
    const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';
    const [countResult, rows] = await Promise.all([
      env.DB.prepare(`SELECT COUNT(*) as total FROM ${table} ${where}`).bind(...values).first(),
      env.DB.prepare(`SELECT * FROM ${table} ${where} ORDER BY name LIMIT ? OFFSET ?`).bind(...values, limit, offset).all()
    ]);
    return paginated(rows.results, countResult.total, page, limit, origin);
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'sop', 'write')) return forbidden(origin);

  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'POST' && path === '') {
    let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
    const { name, category, document_link, description, version, effective_date } = body;
    if (!name) return error('name required', 400, origin);
    const result = await env.DB.prepare(`INSERT INTO ${table} (name, category, document_link, description) VALUES (?, ?, ?, ?)`)
      .bind(name, category || null, document_link || null, description || null).run();
    return ok({ id: result.meta.last_row_id }, 201, origin);
  }

  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') {
      const row = await env.DB.prepare(`SELECT * FROM ${table} WHERE id = ?`).bind(id).first();
      return row ? ok(row, 200, origin) : notFound(origin);
    }
    if (request.method === 'PUT') {
      let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
      const { name, category, document_link, description } = body;
      await env.DB.prepare(`UPDATE ${table} SET name = COALESCE(?, name), category = COALESCE(?, category), document_link = COALESCE(?, document_link), description = COALESCE(?, description), updated_at = datetime('now') WHERE id = ?`)
        .bind(name || null, category || null, document_link || null, description || null, id).run();
      return ok({ message: 'Updated' }, 200, origin);
    }
    if (request.method === 'DELETE') {
      await env.DB.prepare(`DELETE FROM ${table} WHERE id = ?`).bind(id).run();
      return ok({ message: 'Deleted' }, 200, origin);
    }
  }
  return error('Not found', 404, origin);
}

async function handleForms(request, env, origin, path) {
  // Public GET for is_public forms
  if (request.method === 'GET' && path === '/public') {
    const rows = await env.DB.prepare('SELECT * FROM master_forms WHERE is_public = 1 ORDER BY name').all();
    return ok(rows.results, 200, origin);
  }

  if (request.method === 'GET' && path === '') {
    const { page, limit, offset } = getPagination(request.url);
    const url = new URL(request.url);
    const search = url.searchParams.get('search') || '';
    let conditions = []; let values = [];
    if (search) { conditions.push('(name LIKE ? OR category LIKE ?)'); values.push(`%${search}%`, `%${search}%`); }
    const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';
    const [countResult, rows] = await Promise.all([
      env.DB.prepare(`SELECT COUNT(*) as total FROM master_forms ${where}`).bind(...values).first(),
      env.DB.prepare(`SELECT * FROM master_forms ${where} ORDER BY name LIMIT ? OFFSET ?`).bind(...values, limit, offset).all()
    ]);
    const user = await authenticate(request, env);
    if (!user) {
      // Public: only return is_public forms
      const pubRows = await env.DB.prepare('SELECT * FROM master_forms WHERE is_public = 1 ORDER BY name').all();
      return ok(pubRows.results, 200, origin);
    }
    return paginated(rows.results, countResult.total, page, limit, origin);
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'forms', 'write')) return forbidden(origin);

  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'POST' && path === '') {
    let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
    const { name, category, document_link, description, is_public } = body;
    if (!name) return error('name required', 400, origin);
    const result = await env.DB.prepare('INSERT INTO master_forms (name, category, document_link, description, is_public) VALUES (?, ?, ?, ?, ?)')
      .bind(name, category || null, document_link || null, description || null, is_public ? 1 : 0).run();
    return ok({ id: result.meta.last_row_id }, 201, origin);
  }

  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') {
      const row = await env.DB.prepare('SELECT * FROM master_forms WHERE id = ?').bind(id).first();
      return row ? ok(row, 200, origin) : notFound(origin);
    }
    if (request.method === 'PUT') {
      let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
      const { name, category, document_link, description, is_public } = body;
      await env.DB.prepare(`UPDATE master_forms SET name = COALESCE(?, name), category = COALESCE(?, category), document_link = COALESCE(?, document_link), description = COALESCE(?, description), is_public = COALESCE(?, is_public), updated_at = datetime('now') WHERE id = ?`)
        .bind(name || null, category || null, document_link || null, description || null, is_public !== undefined ? (is_public ? 1 : 0) : null, id).run();
      return ok({ message: 'Updated' }, 200, origin);
    }
    if (request.method === 'DELETE') {
      await env.DB.prepare('DELETE FROM master_forms WHERE id = ?').bind(id).run();
      return ok({ message: 'Deleted' }, 200, origin);
    }
  }
  return error('Not found', 404, origin);
}

async function handlePic(request, env, origin) {
  // Always public GET
  if (request.method === 'GET') {
    const rows = await env.DB.prepare('SELECT * FROM pic_list WHERE is_active = 1 ORDER BY name').all();
    return ok(rows.results, 200, origin);
  }
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'sop', 'admin')) return forbidden(origin);

  if (request.method === 'POST') {
    let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
    const result = await env.DB.prepare('INSERT OR IGNORE INTO pic_list (name, role) VALUES (?, ?)').bind(body.name, body.role || null).run();
    return ok({ id: result.meta.last_row_id }, 201, origin);
  }
  return error('Not found', 404, origin);
}

async function handleOptions(request, env, origin) {
  if (request.method !== 'GET') return error('Method not allowed', 405, origin);
  const rows = await env.DB.prepare('SELECT category, value FROM validation_options ORDER BY value').all();
  
  const data = {
    pic: [],
    activity: [],
    quarter: [],
    pkwt: []
  };
  
  (rows.results || []).forEach(r => {
    if (data[r.category]) {
      data[r.category].push(r.value);
    }
  });
  
  return ok(data, 200, origin);
}

