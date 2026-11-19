// OS Migrations - Timeline & Module API
import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination } from '../utils/pagination.js';
import { requireExportAccess, getExportPagination } from '../utils/export.js';
import { buildAuditStatement, buildCreateAuditStatement } from '../utils/audit.js';

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

        const stmts = [];
        for (const item of SEED_BRANCHES) {
          const b = bMap.get(item.code);
          if (b) {
            stmts.push(
              env.DB.prepare(`
                INSERT INTO os_migrations (branch_id, vendor, target_month, status, notes)
                VALUES (?, ?, ?, ?, ?)
              `).bind(b.id, item.vendor, item.target_month, item.status, `Migrasi OS ke ${item.vendor} (${b.full_name})`)
            );
          }
        }
        if (stmts.length > 0) {
          for (let i = 0; i < stmts.length; i += 50) {
            await env.DB.batch(stmts.slice(i, i + 50));
          }
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
    console.warn('os_migrations initialization warning:', err.message);
  }
}

export async function handleOSMigrations(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace('/api/os-migrations', '');

  if (request.method === 'GET' && path === '/export') {
    const { response } = await requireExportAccess(request, env, origin, 'os_migrations');
    if (response) return response;
    return listOSMigrations(request, env, origin, { isExport: true });
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'os_migrations', 'read')) return forbidden(origin);

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

  // PATCH status quickly
  if (request.method === 'PATCH' && statusMatch) {
    if (!hasPermission(user, 'os_migrations', 'write')) return forbidden(origin);
    return updateMigrationStatus(statusMatch[1], request, env, origin, user);
  }

  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') {
      return getOSMigration(id, env, origin);
    }
    if (request.method === 'PUT' || request.method === 'PATCH') {
      if (!hasPermission(user, 'os_migrations', 'write')) return forbidden(origin);
      return updateOSMigration(id, request, env, origin, user);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'os_migrations', 'delete')) return forbidden(origin);
      return deleteOSMigration(id, env, origin, user);
    }
  }

  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'os_migrations', 'write')) return forbidden(origin);
    return createOSMigration(request, env, origin, user);
  }

  return error('Not found', 404, origin);
}

async function listOSMigrations(request, env, origin, { isExport = false } = {}) {
  const url = new URL(request.url);
  const { page, limit, offset } = isExport ? getExportPagination(request.url) : getPagination(request.url);
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
      COALESCE(ec.cnt, 0) as employee_count
    FROM os_migrations m
    JOIN branches b ON m.branch_id = b.id
    LEFT JOIN (
      SELECT branch_id, COUNT(*) as cnt
      FROM employees
      WHERE status = 'Aktif' AND deleted_at IS NULL
      GROUP BY branch_id
    ) ec ON ec.branch_id = m.branch_id
    ${whereClause}
    ORDER BY m.target_month ASC, m.vendor ASC, b.code ASC, m.id ASC
  `;

  if (!isExport && all === '1') {
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

  return paginated(rows.results || [], count?.total || 0, page, limit, origin, env?.ENVIRONMENT);
}

async function getMigrationSummary(request, env, origin) {
  const query = `
    SELECT 
      m.*,
      b.name as branch_name,
      b.code as branch_code,
      b.full_name as branch_full_name,
      COALESCE(ec.cnt, 0) as employee_count
    FROM os_migrations m
    JOIN branches b ON m.branch_id = b.id
    LEFT JOIN (
      SELECT branch_id, COUNT(*) as cnt
      FROM employees
      WHERE status = 'Aktif' AND deleted_at IS NULL
      GROUP BY branch_id
    ) ec ON ec.branch_id = m.branch_id
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

async function createOSMigration(request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
  try {
    const { branch_id, vendor, target_month, status, actual_date, pic_vendor, notes } = body;

    if (!branch_id || !vendor || !target_month) {
      return error('Cabang, Vendor, dan Target Bulan wajib diisi', 400, origin);
    }
    if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(String(target_month).trim())) {
      return error('Format target_month tidak valid (gunakan YYYY-MM)', 400, origin);
    }
    const safeStatus = status || 'scheduled';
    if (!['scheduled', 'in_progress', 'completed'].includes(safeStatus)) {
      return error('Status tidak valid (harus: scheduled, in_progress, atau completed)', 400, origin);
    }

    const branch = await env.DB.prepare('SELECT id, full_name FROM branches WHERE id = ? AND deleted_at IS NULL').bind(branch_id).first();
    if (!branch) return error('Cabang tidak valid', 400, origin);

    const safeNotes = notes || `Migrasi OS ke ${vendor} (${branch.full_name})`;
    const newSnapshot = {
      branch_id,
      vendor,
      target_month,
      status: safeStatus,
      actual_date: actual_date || null,
      pic_vendor: pic_vendor || null,
      notes: safeNotes
    };

    const batchRes = await env.DB.batch([
      env.DB.prepare(`
        INSERT INTO os_migrations (branch_id, vendor, target_month, status, actual_date, pic_vendor, notes)
        VALUES (?, ?, ?, ?, ?, ?, ?)
      `).bind(branch_id, vendor, target_month, safeStatus, actual_date || null, pic_vendor || null, safeNotes),
      buildCreateAuditStatement(env, user, 'os_migrations', 'new', newSnapshot)
    ]);
    const newId = batchRes?.[0]?.meta?.last_row_id;

    const created = await env.DB.prepare(`
      SELECT m.*, b.name as branch_name, b.code as branch_code, b.full_name as branch_full_name
      FROM os_migrations m
      JOIN branches b ON m.branch_id = b.id
      WHERE m.id = ?
    `).bind(newId).first();

    return ok(created, 201, origin, undefined, { 'X-Audit-Batched': 'true' });
  } catch (err) {
    console.error('createOSMigration error:', err);
    return error('Gagal menambahkan migrasi OS', 500, origin);
  }
}

async function updateOSMigration(id, request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
  try {
    const { branch_id, vendor, target_month, status, actual_date, pic_vendor, notes } = body;

    const existing = await env.DB.prepare('SELECT * FROM os_migrations WHERE id = ?').bind(id).first();
    if (!existing) return notFound(origin);

    const finalBranchId = branch_id || existing.branch_id;
    const finalVendor = vendor || existing.vendor;
    const finalTargetMonth = target_month || existing.target_month;
    const finalStatus = status || existing.status || 'scheduled';
    const finalActualDate = actual_date !== undefined ? (actual_date || null) : existing.actual_date;
    const finalPicVendor = pic_vendor !== undefined ? (pic_vendor || null) : existing.pic_vendor;
    const finalNotes = notes !== undefined ? (notes || null) : existing.notes;

    if (!finalBranchId || !finalVendor || !finalTargetMonth) {
      return error('Cabang, Vendor, dan Target Bulan wajib diisi', 400, origin);
    }
    if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(String(finalTargetMonth).trim())) {
      return error('Format target_month tidak valid (gunakan YYYY-MM)', 400, origin);
    }
    if (!['scheduled', 'in_progress', 'completed'].includes(finalStatus)) {
      return error('Status tidak valid (harus: scheduled, in_progress, atau completed)', 400, origin);
    }

    const branch = await env.DB.prepare('SELECT id FROM branches WHERE id = ? AND deleted_at IS NULL').bind(finalBranchId).first();
    if (!branch) return error('Cabang tidak valid', 400, origin);

    const afterSnapshot = {
      ...existing,
      branch_id: finalBranchId,
      vendor: finalVendor,
      target_month: finalTargetMonth,
      status: finalStatus,
      actual_date: finalActualDate,
      pic_vendor: finalPicVendor,
      notes: finalNotes
    };

    await env.DB.batch([
      env.DB.prepare(`
        UPDATE os_migrations
        SET branch_id = ?, vendor = ?, target_month = ?, status = ?, actual_date = ?, pic_vendor = ?, notes = ?, updated_at = datetime('now')
        WHERE id = ?
      `).bind(finalBranchId, finalVendor, finalTargetMonth, finalStatus, finalActualDate, finalPicVendor, finalNotes, id),
      buildAuditStatement(env, user, 'UPDATE', 'os_migrations', id, existing, afterSnapshot, null, 'SUCCESS', true)
    ]);

    const updated = await env.DB.prepare(`
      SELECT m.*, b.name as branch_name, b.code as branch_code, b.full_name as branch_full_name
      FROM os_migrations m
      JOIN branches b ON m.branch_id = b.id
      WHERE m.id = ?
    `).bind(id).first();

    return ok(updated, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
  } catch (err) {
    console.error('updateOSMigration error:', err);
    return error('Gagal mengupdate migrasi OS', 500, origin);
  }
}

async function updateMigrationStatus(id, request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
  try {
    const { status, actual_date } = body;

    if (!status || !['scheduled', 'in_progress', 'completed'].includes(status)) {
      return error('Status tidak valid (harus: scheduled, in_progress, atau completed)', 400, origin);
    }

    const existing = await env.DB.prepare('SELECT * FROM os_migrations WHERE id = ?').bind(id).first();
    if (!existing) return notFound(origin);

    let sql = `UPDATE os_migrations SET status = ?, updated_at = datetime('now')`;
    let binds = [status];
    let resolvedActualDate = existing.actual_date;

    if (actual_date !== undefined) {
      sql += `, actual_date = ?`;
      resolvedActualDate = actual_date || null;
      binds.push(resolvedActualDate);
    } else if (status === 'completed') {
      sql += `, actual_date = COALESCE(actual_date, date('now'))`;
      resolvedActualDate = existing.actual_date || new Date().toISOString().slice(0, 10);
    }

    sql += ` WHERE id = ?`;
    binds.push(id);

    const afterSnapshot = { ...existing, status, actual_date: resolvedActualDate };
    await env.DB.batch([
      env.DB.prepare(sql).bind(...binds),
      buildAuditStatement(env, user, 'UPDATE', 'os_migrations', id, existing, afterSnapshot, { status_only: true }, 'SUCCESS', true)
    ]);

    const updated = await env.DB.prepare(`
      SELECT m.*, b.name as branch_name, b.code as branch_code, b.full_name as branch_full_name
      FROM os_migrations m
      JOIN branches b ON m.branch_id = b.id
      WHERE m.id = ?
    `).bind(id).first();

    return ok(updated, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
  } catch (err) {
    console.error('updateMigrationStatus error:', err);
    return error('Gagal mengubah status migrasi OS', 500, origin);
  }
}

async function deleteOSMigration(id, env, origin, user = null) {
  try {
    const existing = await env.DB.prepare('SELECT * FROM os_migrations WHERE id = ?').bind(id).first();
    if (!existing) return notFound(origin);

    const delStmt = env.DB.prepare('DELETE FROM os_migrations WHERE id = ?').bind(id);
    const { executeAtomicDelete } = await import('../utils/audit.js');
    await executeAtomicDelete(env, user, delStmt, 'os_migrations', id, existing);
    return ok({ message: 'Data migrasi OS berhasil dihapus' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
  } catch (err) {
    console.error('deleteOSMigration error:', err);
    return error('Gagal menghapus migrasi OS', 500, origin);
  }
}
