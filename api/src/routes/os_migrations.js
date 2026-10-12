// OS Migrations - Timeline & Module API
import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination } from '../utils/pagination.js';

const SEED_BRANCHES = [
  // AEON September (9 branches)
  { vendor: 'AEON', target_month: '2026-09', code: '003', status: 'completed' },
  { vendor: 'AEON', target_month: '2026-09', code: '007', status: 'completed' },
  { vendor: 'AEON', target_month: '2026-09', code: '012', status: 'completed' },
  { vendor: 'AEON', target_month: '2026-09', code: '014', status: 'completed' },
  { vendor: 'AEON', target_month: '2026-09', code: '019', status: 'completed' },
  { vendor: 'AEON', target_month: '2026-09', code: '021', status: 'completed' },
  { vendor: 'AEON', target_month: '2026-09', code: '011', status: 'completed' },
  { vendor: 'AEON', target_month: '2026-09', code: '046', status: 'completed' },
  { vendor: 'AEON', target_month: '2026-09', code: '059', status: 'completed' },

  // AEON November (3 branches)
  { vendor: 'AEON', target_month: '2026-11', code: '002', status: 'scheduled' },
  { vendor: 'AEON', target_month: '2026-11', code: '004', status: 'scheduled' },
  { vendor: 'AEON', target_month: '2026-11', code: '016', status: 'scheduled' },

  // AEON December (8 branches)
  { vendor: 'AEON', target_month: '2026-12', code: '005', status: 'scheduled' },
  { vendor: 'AEON', target_month: '2026-12', code: '009', status: 'scheduled' },
  { vendor: 'AEON', target_month: '2026-12', code: '010', status: 'scheduled' },
  { vendor: 'AEON', target_month: '2026-12', code: '013', status: 'scheduled' },
  { vendor: 'AEON', target_month: '2026-12', code: '015', status: 'scheduled' },
  { vendor: 'AEON', target_month: '2026-12', code: '024', status: 'scheduled' },
  { vendor: 'AEON', target_month: '2026-12', code: '027', status: 'scheduled' },
  { vendor: 'AEON', target_month: '2026-12', code: '031', status: 'scheduled' },

  // Triple S September (5 branches)
  { vendor: 'Triple S', target_month: '2026-09', code: '029', status: 'completed' },
  { vendor: 'Triple S', target_month: '2026-09', code: '008', status: 'completed' },
  { vendor: 'Triple S', target_month: '2026-09', code: '050', status: 'completed' },
  { vendor: 'Triple S', target_month: '2026-09', code: '052', status: 'completed' },
  { vendor: 'Triple S', target_month: '2026-09', code: '054', status: 'completed' },

  // Triple S October (5 branches)
  { vendor: 'Triple S', target_month: '2026-10', code: '026', status: 'in_progress' },
  { vendor: 'Triple S', target_month: '2026-10', code: '028', status: 'in_progress' },
  { vendor: 'Triple S', target_month: '2026-10', code: '042', status: 'in_progress' },
  { vendor: 'Triple S', target_month: '2026-10', code: '056', status: 'in_progress' },
  { vendor: 'Triple S', target_month: '2026-10', code: '062', status: 'in_progress' },

  // Triple S November (5 branches)
  { vendor: 'Triple S', target_month: '2026-11', code: '034', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-11', code: '038', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-11', code: '039', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-11', code: '040', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-11', code: '043', status: 'scheduled' },

  // Triple S December (27 branches)
  { vendor: 'Triple S', target_month: '2026-12', code: '006', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '030', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '022', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '017', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '018', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '020', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '023', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '025', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '032', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '033', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '035', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '036', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '041', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '044', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '045', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '047', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '048', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '049', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '051', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '053', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '055', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '057', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '060', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '063', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '064', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '065', status: 'scheduled' },
  { vendor: 'Triple S', target_month: '2026-12', code: '067', status: 'scheduled' }
];

let _initialized = false;
async function ensureInitialized(env) {
  if (_initialized) return;
  try {
    await env.DB.prepare(`
      CREATE TABLE IF NOT EXISTS os_migrations (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        branch_id INTEGER NOT NULL,
        vendor TEXT NOT NULL,
        target_month TEXT NOT NULL,
        status TEXT NOT NULL DEFAULT 'scheduled',
        actual_date TEXT,
        pic_vendor TEXT,
        notes TEXT,
        created_at TEXT NOT NULL DEFAULT (datetime('now')),
        updated_at TEXT NOT NULL DEFAULT (datetime('now')),
        FOREIGN KEY (branch_id) REFERENCES branches(id)
      )
    `).run();

    await env.DB.prepare(`CREATE INDEX IF NOT EXISTS idx_os_migrations_branch ON os_migrations(branch_id)`).run();
    await env.DB.prepare(`CREATE INDEX IF NOT EXISTS idx_os_migrations_vendor_month ON os_migrations(vendor, target_month)`).run();

    const countRow = await env.DB.prepare('SELECT COUNT(*) as cnt FROM os_migrations').first();
    if (!countRow || countRow.cnt === 0) {
      const branches = await env.DB.prepare('SELECT id, code, full_name FROM branches WHERE is_active = 1').all();
      const bMap = new Map();
      (branches.results || []).forEach(b => bMap.set(b.code, b));

      for (const item of SEED_BRANCHES) {
        const b = bMap.get(item.code);
        if (b) {
          await env.DB.prepare(`
            INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes)
            VALUES (?, ?, ?, ?, ?)
          `).bind(b.id, item.vendor, item.target_month, item.status, `Migrasi OS ke ${item.vendor} (${b.full_name})`).run();
        }
      }
    }
    _initialized = true;
  } catch (err) {
    console.warn('os_migrations initialization warning:', err.message);
  }
}

export async function handleOSMigrations(request, env, origin) {
  await ensureInitialized(env);

  const url = new URL(request.url);
  const path = url.pathname.replace('/api/os-migrations', '');
  const idMatch = path.match(/^\/(\d+)$/);
  const statusMatch = path.match(/^\/(\d+)\/status$/);

  // GET summary (stats & matrix)
  if (request.method === 'GET' && path === '/summary') {
    return getMigrationSummary(request, env, origin);
  }

  // GET list
  if (request.method === 'GET' && path === '') {
    return listOSMigrations(request, env, origin);
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);

  // PATCH status quickly
  if (request.method === 'PATCH' && statusMatch) {
    if (!hasPermission(user, 'os_migrations', 'write')) return forbidden(origin);
    return updateMigrationStatus(statusMatch[1], request, env, origin);
  }

  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') {
      return getOSMigration(id, env, origin);
    }
    if (request.method === 'PUT') {
      if (!hasPermission(user, 'os_migrations', 'write')) return forbidden(origin);
      return updateOSMigration(id, request, env, origin);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'os_migrations', 'delete')) return forbidden(origin);
      return deleteOSMigration(id, env, origin);
    }
  }

  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'os_migrations', 'write')) return forbidden(origin);
    return createOSMigration(request, env, origin);
  }

  return error('Not found', 404, origin);
}

async function listOSMigrations(request, env, origin) {
  const url = new URL(request.url);
  const { page, limit, offset } = getPagination(request.url);
  const search = url.searchParams.get('search') || '';
  const vendor = url.searchParams.get('vendor') || '';
  const month = url.searchParams.get('month') || '';
  const status = url.searchParams.get('status') || '';
  const all = url.searchParams.get('all');

  let conditions = [];
  let values = [];

  if (search) {
    conditions.push('(b.name LIKE ? OR b.code LIKE ? OR b.full_name LIKE ? OR m.notes LIKE ? OR m.pic_vendor LIKE ?)');
    values.push(`%${search}%`, `%${search}%`, `%${search}%`, `%${search}%`, `%${search}%`);
  }
  if (vendor) {
    conditions.push('m.vendor = ?');
    values.push(vendor);
  }
  if (month) {
    conditions.push('m.target_month = ?');
    values.push(month);
  }
  if (status) {
    conditions.push('m.status = ?');
    values.push(status);
  }

  const whereClause = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';

  const baseSelect = `
    SELECT 
      m.*,
      b.name as branch_name,
      b.code as branch_code,
      b.full_name as branch_full_name,
      (SELECT COUNT(*) FROM employees e WHERE e.branch_id = m.branch_id AND e.status = 'Aktif' AND e.deleted_at IS NULL) as employee_count
    FROM os_migrations m
    JOIN branches b ON m.branch_id = b.id
    ${whereClause}
    ORDER BY m.target_month ASC, m.vendor ASC, b.code ASC
  `;

  if (all === '1') {
    const rows = await env.DB.prepare(baseSelect).bind(...values).all();
    return ok(rows.results || [], 200, origin);
  }

  const countQuery = `
    SELECT COUNT(*) as total
    FROM os_migrations m
    JOIN branches b ON m.branch_id = b.id
    ${whereClause}
  `;

  const [rows, count] = await Promise.all([
    env.DB.prepare(`${baseSelect} LIMIT ? OFFSET ?`).bind(...values, limit, offset).all(),
    env.DB.prepare(countQuery).bind(...values).first()
  ]);

  return paginated(rows.results || [], count?.total || 0, page, limit, origin);
}

async function getMigrationSummary(request, env, origin) {
  const query = `
    SELECT 
      m.*,
      b.name as branch_name,
      b.code as branch_code,
      b.full_name as branch_full_name,
      (SELECT COUNT(*) FROM employees e WHERE e.branch_id = m.branch_id AND e.status = 'Aktif' AND e.deleted_at IS NULL) as employee_count
    FROM os_migrations m
    JOIN branches b ON m.branch_id = b.id
    ORDER BY m.target_month ASC, m.vendor ASC, b.code ASC
  `;

  const rows = (await env.DB.prepare(query).all()).results || [];

  const summary = {
    total_branches: rows.length,
    by_status: {
      scheduled: 0,
      in_progress: 0,
      completed: 0
    },
    by_month: {
      '2026-09': 0,
      '2026-10': 0,
      '2026-11': 0,
      '2026-12': 0
    },
    vendors: {
      'AEON': {
        total: 0,
        by_month: { '2026-09': 0, '2026-10': 0, '2026-11': 0, '2026-12': 0 },
        by_status: { scheduled: 0, in_progress: 0, completed: 0 }
      },
      'Triple S': {
        total: 0,
        by_month: { '2026-09': 0, '2026-10': 0, '2026-11': 0, '2026-12': 0 },
        by_status: { scheduled: 0, in_progress: 0, completed: 0 }
      }
    },
    matrix: rows
  };

  rows.forEach(r => {
    // Status
    if (summary.by_status[r.status] !== undefined) {
      summary.by_status[r.status]++;
    }

    // Month
    if (summary.by_month[r.target_month] !== undefined) {
      summary.by_month[r.target_month]++;
    } else {
      summary.by_month[r.target_month] = 1;
    }

    // Vendor
    if (!summary.vendors[r.vendor]) {
      summary.vendors[r.vendor] = {
        total: 0,
        by_month: {},
        by_status: { scheduled: 0, in_progress: 0, completed: 0 }
      };
    }
    const vObj = summary.vendors[r.vendor];
    vObj.total++;
    vObj.by_month[r.target_month] = (vObj.by_month[r.target_month] || 0) + 1;
    if (vObj.by_status[r.status] !== undefined) {
      vObj.by_status[r.status]++;
    }
  });

  return ok(summary, 200, origin);
}

async function getOSMigration(id, env, origin) {
  const query = `
    SELECT 
      m.*,
      b.name as branch_name,
      b.code as branch_code,
      b.full_name as branch_full_name,
      (SELECT COUNT(*) FROM employees e WHERE e.branch_id = m.branch_id AND e.status = 'Aktif' AND e.deleted_at IS NULL) as employee_count
    FROM os_migrations m
    JOIN branches b ON m.branch_id = b.id
    WHERE m.id = ?
  `;
  const row = await env.DB.prepare(query).bind(id).first();
  if (!row) return notFound(origin);
  return ok(row, 200, origin);
}

async function createOSMigration(request, env, origin) {
  try {
    const body = await request.json();
    const { branch_id, vendor, target_month, status, actual_date, pic_vendor, notes } = body;

    if (!branch_id || !vendor || !target_month) {
      return error('Cabang, Vendor, dan Target Bulan wajib diisi', 400, origin);
    }

    const branch = await env.DB.prepare('SELECT id, full_name FROM branches WHERE id = ?').bind(branch_id).first();
    if (!branch) return error('Cabang tidak valid', 400, origin);

    const safeStatus = status || 'scheduled';
    const safeNotes = notes || `Migrasi OS ke ${vendor} (${branch.full_name})`;

    const res = await env.DB.prepare(`
      INSERT INTO os_migrations (branch_id, vendor, target_month, status, actual_date, pic_vendor, notes)
      VALUES (?, ?, ?, ?, ?, ?, ?)
    `).bind(branch_id, vendor, target_month, safeStatus, actual_date || null, pic_vendor || null, safeNotes).run();

    const created = await env.DB.prepare(`
      SELECT m.*, b.name as branch_name, b.code as branch_code, b.full_name as branch_full_name
      FROM os_migrations m
      JOIN branches b ON m.branch_id = b.id
      WHERE m.id = ?
    `).bind(res.meta.last_row_id).first();

    return ok(created, 201, origin);
  } catch (err) {
    return error(`Gagal menambahkan migrasi OS: ${err.message}`, 500, origin);
  }
}

async function updateOSMigration(id, request, env, origin) {
  try {
    const body = await request.json();
    const { branch_id, vendor, target_month, status, actual_date, pic_vendor, notes } = body;

    const existing = await env.DB.prepare('SELECT id FROM os_migrations WHERE id = ?').bind(id).first();
    if (!existing) return notFound(origin);

    if (!branch_id || !vendor || !target_month) {
      return error('Cabang, Vendor, dan Target Bulan wajib diisi', 400, origin);
    }

    await env.DB.prepare(`
      UPDATE os_migrations
      SET branch_id = ?, vendor = ?, target_month = ?, status = ?, actual_date = ?, pic_vendor = ?, notes = ?, updated_at = datetime('now')
      WHERE id = ?
    `).bind(branch_id, vendor, target_month, status || 'scheduled', actual_date || null, pic_vendor || null, notes || null, id).run();

    const updated = await env.DB.prepare(`
      SELECT m.*, b.name as branch_name, b.code as branch_code, b.full_name as branch_full_name
      FROM os_migrations m
      JOIN branches b ON m.branch_id = b.id
      WHERE m.id = ?
    `).bind(id).first();

    return ok(updated, 200, origin);
  } catch (err) {
    return error(`Gagal mengupdate migrasi OS: ${err.message}`, 500, origin);
  }
}

async function updateMigrationStatus(id, request, env, origin) {
  try {
    const body = await request.json();
    const { status, actual_date } = body;

    if (!status || !['scheduled', 'in_progress', 'completed'].includes(status)) {
      return error('Status tidak valid (harus: scheduled, in_progress, atau completed)', 400, origin);
    }

    const existing = await env.DB.prepare('SELECT id FROM os_migrations WHERE id = ?').bind(id).first();
    if (!existing) return notFound(origin);

    let sql = `UPDATE os_migrations SET status = ?, updated_at = datetime('now')`;
    let binds = [status];

    if (actual_date !== undefined) {
      sql += `, actual_date = ?`;
      binds.push(actual_date || null);
    } else if (status === 'completed') {
      sql += `, actual_date = COALESCE(actual_date, date('now'))`;
    }

    sql += ` WHERE id = ?`;
    binds.push(id);

    await env.DB.prepare(sql).bind(...binds).run();

    const updated = await env.DB.prepare(`
      SELECT m.*, b.name as branch_name, b.code as branch_code, b.full_name as branch_full_name
      FROM os_migrations m
      JOIN branches b ON m.branch_id = b.id
      WHERE m.id = ?
    `).bind(id).first();

    return ok(updated, 200, origin);
  } catch (err) {
    return error(`Gagal mengubah status migrasi OS: ${err.message}`, 500, origin);
  }
}

async function deleteOSMigration(id, env, origin) {
  try {
    const existing = await env.DB.prepare('SELECT id FROM os_migrations WHERE id = ?').bind(id).first();
    if (!existing) return notFound(origin);

    await env.DB.prepare('DELETE FROM os_migrations WHERE id = ?').bind(id).run();
    return ok({ message: 'Data migrasi OS berhasil dihapus' }, 200, origin);
  } catch (err) {
    return error(`Gagal menghapus migrasi OS: ${err.message}`, 500, origin);
  }
}
