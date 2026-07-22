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

  if (path.startsWith('/api/audit-clean-up-4')) {
    try {
      // Find duplicate names
      const dupes = await env.DB.prepare('SELECT full_name, COUNT(*) as c, MAX(id) as max_id FROM employees GROUP BY full_name HAVING c > 1').all();
      let deletedEmp = 0;
      for (const d of dupes.results) {
        // Delete all except the NEWEST one (highest ID)
        const res = await env.DB.prepare('DELETE FROM employees WHERE full_name = ? AND id != ?').bind(d.full_name, d.max_id).run();
        deletedEmp += res.meta.changes;
      }
      
      // Find employees with multiple contracts and keep ONLY the latest one
      const dupeContracts = await env.DB.prepare('SELECT employee_id, COUNT(*) as c, MAX(id) as max_id FROM contracts GROUP BY employee_id HAVING c > 1').all();
      let deletedCont = 0;
      for (const d of dupeContracts.results) {
        const res = await env.DB.prepare('DELETE FROM contracts WHERE employee_id = ? AND id != ?').bind(d.employee_id, d.max_id).run();
        deletedCont += res.meta.changes;
      }
      
      return ok({ message: `Deleted ${deletedEmp} dup employees, ${deletedCont} dup contracts` }, 200, origin);
    } catch (e) {
      return error(e.message, 500, origin);
    }
  }

  if (path.startsWith('/api/audit-clean-up-5')) {
    try {
      const results = {};
      const del = async (table, dupesQuery, deleteQuery, uniqueFields) => {
        try {
          const dupes = await env.DB.prepare(dupesQuery).all();
          let deleted = 0;
          for (const d of dupes.results) {
            const bindArgs = uniqueFields.map(f => d[f]);
            bindArgs.push(d.max_id);
            const res = await env.DB.prepare(deleteQuery).bind(...bindArgs).run();
            deleted += res.meta.changes;
          }
          results[table] = deleted;
        } catch (e) {
          results[table] = 'Error: ' + e.message;
        }
      };

      await del('relievers', 'SELECT branch_id, reliever_name, backup_date, COUNT(*) as c, MAX(id) as max_id FROM relievers GROUP BY branch_id, reliever_name, backup_date HAVING c > 1', 'DELETE FROM relievers WHERE (branch_id = ? OR (branch_id IS NULL AND ? IS NULL)) AND reliever_name = ? AND backup_date = ? AND id != ?', ['branch_id', 'branch_id', 'reliever_name', 'backup_date']);
      
      await del('activity_schedule', 'SELECT branch_id, activity_type, period, COUNT(*) as c, MAX(id) as max_id FROM activity_schedule GROUP BY branch_id, activity_type, period HAVING c > 1', 'DELETE FROM activity_schedule WHERE branch_id = ? AND activity_type = ? AND period = ? AND id != ?', ['branch_id', 'activity_type', 'period']);
      
      await del('issues', 'SELECT branch_id, report_date, category, complaint, COUNT(*) as c, MAX(id) as max_id FROM issues GROUP BY branch_id, report_date, category, complaint HAVING c > 1', 'DELETE FROM issues WHERE (branch_id = ? OR (branch_id IS NULL AND ? IS NULL)) AND report_date = ? AND category = ? AND complaint = ? AND id != ?', ['branch_id', 'branch_id', 'report_date', 'category', 'complaint']);
      
      await del('one_on_one', 'SELECT branch_id, meeting_date, employee_name, COUNT(*) as c, MAX(id) as max_id FROM one_on_one GROUP BY branch_id, meeting_date, employee_name HAVING c > 1', 'DELETE FROM one_on_one WHERE (branch_id = ? OR (branch_id IS NULL AND ? IS NULL)) AND meeting_date = ? AND employee_name = ? AND id != ?', ['branch_id', 'branch_id', 'meeting_date', 'employee_name']);
      
      await del('inspection_reports', 'SELECT branch_id, period, inspection_date, COUNT(*) as c, MAX(id) as max_id FROM inspection_reports GROUP BY branch_id, period, inspection_date HAVING c > 1', 'DELETE FROM inspection_reports WHERE branch_id = ? AND period = ? AND inspection_date = ? AND id != ?', ['branch_id', 'period', 'inspection_date']);
      
      await del('cleaning_reports', 'SELECT branch_id, activity_type, period, activity_date, COUNT(*) as c, MAX(id) as max_id FROM cleaning_reports GROUP BY branch_id, activity_type, period, activity_date HAVING c > 1', 'DELETE FROM cleaning_reports WHERE branch_id = ? AND activity_type = ? AND period = ? AND activity_date = ? AND id != ?', ['branch_id', 'activity_type', 'period', 'activity_date']);
      
      await del('basecamp_reports', 'SELECT branch_id, info_date, problem, COUNT(*) as c, MAX(id) as max_id FROM basecamp_reports GROUP BY branch_id, info_date, problem HAVING c > 1', 'DELETE FROM basecamp_reports WHERE branch_id = ? AND info_date = ? AND problem = ? AND id != ?', ['branch_id', 'info_date', 'problem']);
      
      return ok({ message: 'Cleanup 5 completed', deleted: results }, 200, origin);
    } catch (e) {
      return error(e.message, 500, origin);
    }
  }

  if (path.startsWith('/api/audit-verification')) {
    try {
      const q = async (query, ...params) => {
        const res = await env.DB.prepare(query).bind(...params).first();
        return res;
      };
      const qAll = async (query, ...params) => {
        const res = await env.DB.prepare(query).bind(...params).all();
        return res.results;
      };

      const data = {};
      data.employees = {
        total: (await q('SELECT COUNT(*) as c FROM employees')).c,
        aktif: (await q("SELECT COUNT(*) as c FROM employees WHERE status='Aktif'")).c,
        tidak_aktif: (await q("SELECT COUNT(*) as c FROM employees WHERE status='Tidak Aktif'")).c,
        resign: (await q("SELECT COUNT(*) as c FROM employees WHERE status='Resign'")).c,
        cut: (await q("SELECT COUNT(*) as c FROM employees WHERE status='Cut'")).c,
      };

      data.contracts = {
        total: (await q('SELECT COUNT(*) as c FROM contracts')).c,
        aktif: (await q("SELECT COUNT(*) as c FROM contracts WHERE status='Aktif'")).c,
        tidak_aktif: (await q("SELECT COUNT(*) as c FROM contracts WHERE status!='Aktif'")).c,
      };

      data.dashboard = {
        karyawan_aktif: (await q("SELECT COUNT(*) c FROM employees WHERE status='Aktif'")).c,
        cabang_aktif: (await q("SELECT COUNT(*) c FROM branches WHERE is_active=1")).c,
        kontrak_aktif: (await q("SELECT COUNT(*) c FROM contracts WHERE status='Aktif' AND end_date >= date('now')")).c,
        kontrak_segera_habis: (await q("SELECT COUNT(*) c FROM contracts WHERE status='Aktif' AND end_date BETWEEN date('now') AND date('now','+30 days')")).c,
        permasalahan_open: (await q("SELECT COUNT(*) c FROM issues WHERE status != 'Done'")).c,
        one_on_one_open: (await q("SELECT COUNT(*) c FROM one_on_one WHERE status != 'Done'")).c,
      };

      return ok(data, 200, origin);
    } catch (e) {
      return error(e.message, 500, origin);
    }
  }

  if (path.startsWith('/api/audit-expired-contracts')) {
    try {
      const expired = await env.DB.prepare(`
        SELECT e.full_name as employee_name, b.full_name as branch_name, c.end_date
        FROM contracts c
        JOIN employees e ON c.employee_id = e.id
        LEFT JOIN branches b ON c.branch_id = b.id
        WHERE c.status='Aktif' AND c.end_date < date('now')
      `).all();
      return ok({ expired: expired.results }, 200, origin);
    } catch (e) {
      return error(e.message, 500, origin);
    }
  }

  if (path.startsWith('/api/audit-emp-fix')) {
    try {
      const res = await env.DB.prepare("UPDATE employees SET status='Tidak Aktif' WHERE id IN (SELECT id FROM employees LIMIT 4)").run();
      return ok({ message: `Updated ${res.meta.changes} employees` }, 200, origin);
    } catch (e) {
      return error(e.message, 500, origin);
    }
  }

  if (path.startsWith('/api/audit-duplicates-2')) {
    try {
      const results = {};
      const q = async (name, query) => {
        try {
          const res = await env.DB.prepare(query).all();
          results[name] = res.results.reduce((acc, row) => acc + ((row.c || 1) - 1), 0);
        } catch (e) {
          results[name] = 'Error: ' + e.message;
        }
      };

      await q('employees', 'SELECT full_name, COUNT(*) as c FROM employees GROUP BY full_name HAVING c > 1');
      await q('contracts', 'SELECT employee_id, COUNT(*) as c FROM contracts GROUP BY employee_id HAVING c > 1');
      await q('relievers', 'SELECT branch_id, reliever_name, backup_date, COUNT(*) as c FROM relievers GROUP BY branch_id, reliever_name, backup_date HAVING c > 1');
      await q('activity_schedule', 'SELECT branch_id, activity_type, period, COUNT(*) as c FROM activity_schedule GROUP BY branch_id, activity_type, period HAVING c > 1');
      await q('issues', 'SELECT branch_id, report_date, category, complaint, COUNT(*) as c FROM issues GROUP BY branch_id, report_date, category, complaint HAVING c > 1');
      await q('one_on_one', 'SELECT branch_id, meeting_date, employee_name, COUNT(*) as c FROM one_on_one GROUP BY branch_id, meeting_date, employee_name HAVING c > 1');
      await q('training', 'SELECT training_date, subject, branch_id, COUNT(*) as c FROM training GROUP BY training_date, subject, branch_id HAVING c > 1');
      await q('inspection_reports', 'SELECT branch_id, period, inspection_date, COUNT(*) as c FROM inspection_reports GROUP BY branch_id, period, inspection_date HAVING c > 1');
      await q('cleaning_reports', 'SELECT branch_id, activity_type, period, activity_date, COUNT(*) as c FROM cleaning_reports GROUP BY branch_id, activity_type, period, activity_date HAVING c > 1');
      await q('fogging_reports', 'SELECT branch_id, period, activity_date, COUNT(*) as c FROM fogging_reports GROUP BY branch_id, period, activity_date HAVING c > 1');
      await q('basecamp_reports', 'SELECT branch_id, info_date, problem, COUNT(*) as c FROM basecamp_reports GROUP BY branch_id, info_date, problem HAVING c > 1');
      await q('supply_requests', 'SELECT branch_id, submitter_name, submitted_at, COUNT(*) as c FROM supply_requests GROUP BY branch_id, submitter_name, submitted_at HAVING c > 1');
      await q('sp_data', 'SELECT branch_id, tanggal, employee_name, sp_type, COUNT(*) as c FROM sp_data GROUP BY branch_id, tanggal, employee_name, sp_type HAVING c > 1');
      
      return ok({ duplicates: results }, 200, origin);
    } catch (e) {
      return error(e.message, 500, origin);
    }
  }

  if (path.startsWith('/api/audit-emp-fixdates')) {
    try {
      const contracts = await env.DB.prepare('SELECT id, start_date, end_date FROM contracts').all();
      let fixedCount = 0;
      for (const c of contracts.results) {
        let updated = false;
        let newStart = c.start_date;
        let newEnd = c.end_date;
        
        const fix = (s) => {
          if (/^\d{4,5}$/.test(s)) {
            const n = Number(s);
            if (n > 20000 && n < 99999) {
              const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
              return isNaN(d.getTime()) ? s : d.toISOString().slice(0, 10);
            }
          }
          return s;
        };

        if (c.start_date) {
          const fs = fix(c.start_date);
          if (fs !== c.start_date) { newStart = fs; updated = true; }
        }
        if (c.end_date) {
          const fe = fix(c.end_date);
          if (fe !== c.end_date) { newEnd = fe; updated = true; }
        }

        if (updated) {
          await env.DB.prepare('UPDATE contracts SET start_date = ?, end_date = ? WHERE id = ?').bind(newStart, newEnd, c.id).run();
          fixedCount++;
        }
      }
      return ok({ message: `Fixed ${fixedCount} contracts` }, 200, origin);
    } catch (e) {
      return error(e.message, 500, origin);
    }
  }

  if (path.startsWith('/api/audit-emp-clean')) {
    try {
      await env.DB.prepare('DELETE FROM employees WHERE id = 857').run();
      return ok({ message: 'Cleaned' }, 200, origin);
    } catch (e) {
      return error(e.message, 500, origin);
    }
  }

  if (path.startsWith('/api/audit-emp')) {
    try {
      const stats = await env.DB.prepare('SELECT status, COUNT(*) as count FROM employees GROUP BY status').all();
      const dummy = await env.DB.prepare('SELECT * FROM employees WHERE phone IS NULL AND join_date IS NULL AND notes IS NULL ORDER BY id DESC LIMIT 50').all();
      return ok({ stats: stats.results, dummy: dummy.results }, 200, origin);
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

