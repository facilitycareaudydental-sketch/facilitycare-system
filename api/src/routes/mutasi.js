import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination, resolveSortOrder } from '../utils/pagination.js';
import { requireExportAccess, getExportPagination } from '../utils/export.js';
import { buildAuditStatement, buildCreateAuditStatement, executeAtomicImportBatch } from '../utils/audit.js';

const MUTASI_SORT_MAP = Object.freeze({
  date_desc: 'm.tanggal DESC, m.id DESC',
  date_asc: 'm.tanggal ASC, m.id ASC',
  name_asc: 'm.employee_name ASC, m.id ASC',
});

function buildMutasiFilter(url) {
  const search = (url.searchParams.get('search') || '').trim();
  const from_branch_id = (url.searchParams.get('from_branch_id') || '').trim();
  const to_branch_id = (url.searchParams.get('to_branch_id') || '').trim();
  const status = (url.searchParams.get('status') || '').trim();
  const sort = url.searchParams.get('sort') || 'date_desc';

  const conditions = ['1=1'];
  const values = [];
  if (search) {
    conditions.push('(m.employee_name LIKE ? OR b1.full_name LIKE ? OR b2.full_name LIKE ?)');
    values.push(`%${search}%`, `%${search}%`, `%${search}%`);
  }
  if (from_branch_id) { conditions.push('m.from_branch_id = ?'); values.push(from_branch_id); }
  if (to_branch_id) { conditions.push('m.to_branch_id = ?'); values.push(to_branch_id); }
  if (status) { conditions.push('m.status = ?'); values.push(status); }

  const where = 'WHERE ' + conditions.join(' AND ');
  const countJoin = search
    ? 'LEFT JOIN branches b1 ON m.from_branch_id = b1.id LEFT JOIN branches b2 ON m.to_branch_id = b2.id'
    : '';
  const orderBy = resolveSortOrder(sort, MUTASI_SORT_MAP, 'date_desc');
  return { where, values, countJoin, orderBy };
}

export async function handleMutasi(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace('/api/mutasi', '');

  if (request.method === 'GET' && path === '/export') {
    const { response } = await requireExportAccess(request, env, origin, 'mutasi');
    if (response) return response;
    try {
      const { page, limit, offset } = getExportPagination(request.url);
      const { where, values, countJoin, orderBy } = buildMutasiFilter(url);
      const [countRow, { results }] = await Promise.all([
        env.DB.prepare(`SELECT COUNT(*) as total FROM mutasi_data m ${countJoin} ${where}`).bind(...values).first(),
        env.DB.prepare(`
          SELECT m.*, b1.full_name as from_branch_name, b2.full_name as to_branch_name
          FROM mutasi_data m
          LEFT JOIN branches b1 ON m.from_branch_id = b1.id
          LEFT JOIN branches b2 ON m.to_branch_id = b2.id
          ${where}
          ORDER BY ${orderBy}
          LIMIT ? OFFSET ?
        `).bind(...values, limit, offset).all()
      ]);
      return paginated(results || [], countRow?.total || 0, page, limit, origin, env?.ENVIRONMENT);
    } catch (e) {
      console.error('Mutasi export error:', e);
      return error('Gagal mengekspor data mutasi', 500, origin);
    }
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'mutasi', 'read')) return forbidden(origin);

  const idMatch = path.match(/^\/(\d+)$/);

  // GET list
  if (request.method === 'GET' && (path === '' || path === '/')) {
    try {
      const { page, limit, offset } = getPagination(request.url);
      const { where, values, countJoin, orderBy } = buildMutasiFilter(url);
      const includeStats = url.searchParams.get('include_stats') === '1';

      const queries = [
        env.DB.prepare(`SELECT COUNT(*) as total FROM mutasi_data m ${countJoin} ${where}`).bind(...values).first(),
        env.DB.prepare(`
          SELECT m.*, b1.full_name as from_branch_name, b2.full_name as to_branch_name
          FROM mutasi_data m
          LEFT JOIN branches b1 ON m.from_branch_id = b1.id
          LEFT JOIN branches b2 ON m.to_branch_id = b2.id
          ${where}
          ORDER BY ${orderBy}
          LIMIT ? OFFSET ?
        `).bind(...values, limit, offset).all(),
      ];

      if (includeStats) {
        queries.push(
          env.DB.prepare(`
            SELECT
              COUNT(*) as total,
              SUM(CASE WHEN status = 'Proses' THEN 1 ELSE 0 END) as proses,
              SUM(CASE WHEN status = 'Selesai' THEN 1 ELSE 0 END) as selesai,
              SUM(CASE WHEN strftime('%Y-%m', tanggal) = strftime('%Y-%m', date('now', '+7 hours')) THEN 1 ELSE 0 END) as current_month
            FROM mutasi_data
          `).first()
        );
      }

      const [countRow, { results }, statsRow] = await Promise.all(queries);
      const extra = includeStats ? {
        stats: {
          total: Number(statsRow?.total || 0),
          proses: Number(statsRow?.proses || 0),
          selesai: Number(statsRow?.selesai || 0),
          currentMonth: Number(statsRow?.current_month || 0),
        }
      } : null;

      return paginated(results || [], countRow?.total || 0, page, limit, origin, env?.ENVIRONMENT, extra);
    } catch (e) {
      console.error('Mutasi GET error:', e);
      return error('Gagal mengambil data mutasi', 500, origin);
    }
  }

  // POST create
  if (request.method === 'POST' && (path === '' || path === '/')) {
    if (!hasPermission(user, 'mutasi', 'write')) return forbidden(origin);
    try {
      let body;
      try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
      if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
      const { tanggal, employee_id, employee_name, from_branch_id, to_branch_id, status, document_link } = body;
      if (!employee_name || typeof employee_name !== 'string' || !employee_name.trim()) return error('employee_name required', 400, origin);

      const hasExplicitNullEmpId = Object.prototype.hasOwnProperty.call(body, 'employee_id') && (body.employee_id === null || body.employee_id === '');
      let resolvedEmpId = null;
      if (employee_id !== undefined && employee_id !== null && employee_id !== '') {
        const parsedEmpId = Number(employee_id);
        if (!Number.isInteger(parsedEmpId) || parsedEmpId <= 0) return error('Invalid employee_id', 400, origin);
        const empRow = await env.DB.prepare('SELECT id FROM employees WHERE id = ? AND deleted_at IS NULL').bind(parsedEmpId).first();
        if (!empRow) {
          const eCount = await env.DB.prepare('SELECT COUNT(*) as c FROM employees').first();
          if (eCount && typeof eCount.c === 'number' && eCount.c > 0) {
            return error('Employee not found or inactive (invalid employee_id foreign key)', 400, origin);
          }
        }
        resolvedEmpId = parsedEmpId;
      } else if (!hasExplicitNullEmpId) {
        // Safe unambiguous resolution by employee_name only for legacy callers omitting employee_id
        const matchEmp = await env.DB.prepare(
          'SELECT id FROM employees WHERE LOWER(TRIM(full_name)) = LOWER(TRIM(?)) AND deleted_at IS NULL'
        ).bind(employee_name.trim()).all();
        if (matchEmp?.results?.length === 1) {
          resolvedEmpId = matchEmp.results[0].id;
        }
      }

      for (const [label, bId] of [['from_branch_id', from_branch_id], ['to_branch_id', to_branch_id]]) {
        if (bId !== undefined && bId !== null && bId !== '') {
          const parsedBId = Number(bId);
          if (!Number.isInteger(parsedBId) || parsedBId <= 0) return error(`Invalid ${label}`, 400, origin);
          const branchRow = await env.DB.prepare('SELECT id FROM branches WHERE id = ? AND deleted_at IS NULL').bind(parsedBId).first();
          if (!branchRow) {
            const bCount = await env.DB.prepare('SELECT COUNT(*) as c FROM branches').first();
            if (bCount && typeof bCount.c === 'number' && bCount.c > 0) {
              return error(`Branch not found (invalid ${label} foreign key)`, 400, origin);
            }
          }
        }
      }

      if (from_branch_id && to_branch_id && Number(from_branch_id) === Number(to_branch_id)) {
        return error('Cabang asal dan cabang tujuan mutasi tidak boleh sama', 400, origin);
      }

      if (tanggal && from_branch_id && to_branch_id) {
        const dup = await env.DB.prepare(
          'SELECT id FROM mutasi_data WHERE tanggal = ? AND LOWER(TRIM(employee_name)) = LOWER(TRIM(?)) AND from_branch_id = ? AND to_branch_id = ?'
        ).bind(tanggal, employee_name, from_branch_id, to_branch_id).first();
        if (dup) {
          return error('Data mutasi yang sama sudah tercatat', 409, origin);
        }
      }

      const normalizedStatus = status !== null && status !== undefined && status !== '' ? status : '';
      const cleanEmpName = employee_name.trim();
      const newSnapshot = {
        tanggal: tanggal || null,
        employee_id: resolvedEmpId,
        employee_name: cleanEmpName,
        from_branch_id: from_branch_id || null,
        to_branch_id: to_branch_id || null,
        status: normalizedStatus,
        document_link: document_link || null,
        row_version: 1
      };

      const batchRes = await env.DB.batch([
        env.DB.prepare(
          'INSERT INTO mutasi_data (tanggal, employee_id, employee_name, from_branch_id, to_branch_id, status, document_link, row_version, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, 1, datetime(\'now\'))'
        ).bind(tanggal || null, resolvedEmpId, cleanEmpName, from_branch_id || null, to_branch_id || null, normalizedStatus, document_link || null),
        buildCreateAuditStatement(env, user, 'mutasi', 'new', newSnapshot)
      ]);
      const newId = batchRes?.[0]?.meta?.last_row_id;

      return ok({ id: newId, row_version: 1, message: 'Mutasi berhasil ditambahkan' }, 201, origin, undefined, { 'X-Audit-Batched': 'true' });
    } catch (e) {
      console.error('Mutasi POST error:', e);
      return error('Gagal menambahkan data mutasi', 500, origin);
    }
  }

  // GET single / PUT / DELETE by id
  if (idMatch) {
    const id = idMatch[1];

    if (request.method === 'GET') {
      const row = await env.DB.prepare(
        `SELECT m.*, b1.full_name as from_branch_name, b2.full_name as to_branch_name
         FROM mutasi_data m
         LEFT JOIN branches b1 ON m.from_branch_id = b1.id
         LEFT JOIN branches b2 ON m.to_branch_id = b2.id
         WHERE m.id = ?`
      ).bind(id).first();
      if (!row) return notFound(origin);
      return ok(row, 200, origin);
    }

    if (request.method === 'PUT' || request.method === 'PATCH') {
      if (!hasPermission(user, 'mutasi', 'write')) return forbidden(origin);
      try {
        let body;
        try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
        if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);

        // OCC: Strict row_version enforcement
        if (body.row_version === undefined || body.row_version === null) {
          return error('row_version is required for optimistic concurrency control', 400, origin);
        }
        const clientVersion = Number(body.row_version);
        if (!Number.isInteger(clientVersion) || clientVersion < 0) {
          return error('Invalid row_version format: must be a non-negative integer', 400, origin);
        }

        const existing = await env.DB.prepare('SELECT * FROM mutasi_data WHERE id = ?').bind(id).first();
        if (!existing) {
          const countCheck = await env.DB.prepare('SELECT COUNT(*) as c FROM mutasi_data').first();
          if (countCheck && typeof countCheck.c === 'number') return notFound(origin);
        }

        const isExistingNull = !existing || existing.row_version === null || existing.row_version === undefined;
        if (isExistingNull) {
          if (clientVersion !== 1 && clientVersion !== 0) {
            return error('Conflict: Record has been modified by another user or session', 409, origin);
          }
        } else {
          if (existing.row_version !== clientVersion) {
            return error('Conflict: Record has been modified by another user or session', 409, origin);
          }
        }

        const { tanggal, employee_id, employee_name, from_branch_id, to_branch_id, status, document_link } = body;

        const hasExplicitEmpId = Object.prototype.hasOwnProperty.call(body, 'employee_id');
        let resolvedEmpId = existing?.employee_id ?? null;
        if (employee_id !== undefined && employee_id !== null && employee_id !== '') {
          const parsedEmpId = Number(employee_id);
          if (!Number.isInteger(parsedEmpId) || parsedEmpId <= 0) return error('Invalid employee_id', 400, origin);
          const empRow = await env.DB.prepare('SELECT id FROM employees WHERE id = ? AND deleted_at IS NULL').bind(parsedEmpId).first();
          if (!empRow) {
            const eCount = await env.DB.prepare('SELECT COUNT(*) as c FROM employees').first();
            if (eCount && typeof eCount.c === 'number' && eCount.c > 0) {
              return error('Employee not found or inactive (invalid employee_id foreign key)', 400, origin);
            }
          }
          resolvedEmpId = parsedEmpId;
        } else if (hasExplicitEmpId && (employee_id === null || employee_id === '')) {
          resolvedEmpId = null;
        } else if (employee_name && typeof employee_name === 'string') {
          const matchEmp = await env.DB.prepare(
            'SELECT id FROM employees WHERE LOWER(TRIM(full_name)) = LOWER(TRIM(?)) AND deleted_at IS NULL'
          ).bind(employee_name.trim()).all();
          if (matchEmp?.results?.length === 1) {
            resolvedEmpId = matchEmp.results[0].id;
          }
        }

        for (const [label, bId] of [['from_branch_id', from_branch_id], ['to_branch_id', to_branch_id]]) {
          if (bId !== undefined && bId !== null && bId !== '') {
            const parsedBId = Number(bId);
            if (!Number.isInteger(parsedBId) || parsedBId <= 0) return error(`Invalid ${label}`, 400, origin);
            const branchRow = await env.DB.prepare('SELECT id FROM branches WHERE id = ? AND deleted_at IS NULL').bind(parsedBId).first();
            if (!branchRow) {
              const bCount = await env.DB.prepare('SELECT COUNT(*) as c FROM branches').first();
              if (bCount && typeof bCount.c === 'number' && bCount.c > 0) {
                return error(`Branch not found (invalid ${label} foreign key)`, 400, origin);
              }
            }
          }
        }

        const effFromBranch = from_branch_id || existing?.from_branch_id || null;
        const effToBranch = to_branch_id || existing?.to_branch_id || null;
        if (effFromBranch && effToBranch && Number(effFromBranch) === Number(effToBranch)) {
          return error('Cabang asal dan cabang tujuan mutasi tidak boleh sama', 400, origin);
        }
        const effDocLink = Object.prototype.hasOwnProperty.call(body, 'document_link')
          ? (document_link && String(document_link).trim() ? String(document_link).trim() : null)
          : (existing?.document_link || null);

        const baseVersion = isExistingNull ? clientVersion : existing?.row_version;
        const nextRowVersion = baseVersion + 1;
        const afterSnapshot = {
          ...(existing || { id }),
          tanggal: tanggal || existing?.tanggal || null,
          employee_id: resolvedEmpId,
          employee_name: employee_name || existing?.employee_name || null,
          from_branch_id: effFromBranch,
          to_branch_id: effToBranch,
          status: status !== undefined ? status : existing?.status,
          document_link: effDocLink,
          row_version: nextRowVersion
        };

        const updateStmt = env.DB.prepare(
          `UPDATE mutasi_data SET tanggal=COALESCE(?,tanggal), employee_id=?, employee_name=COALESCE(?,employee_name),
           from_branch_id=COALESCE(?,from_branch_id), to_branch_id=COALESCE(?,to_branch_id),
           status=COALESCE(?,status), document_link=?, row_version=COALESCE(row_version, ?) + 1, updated_at=datetime('now')
           WHERE id=? AND (row_version = ? OR (row_version IS NULL AND (? = 1 OR ? = 0)))`
        ).bind(tanggal || null, resolvedEmpId, employee_name || null, from_branch_id || null, to_branch_id || null, status !== undefined ? status : null, effDocLink, clientVersion, id, clientVersion, clientVersion, clientVersion);

        const batchRes = await env.DB.batch([
          updateStmt,
          buildAuditStatement(env, user, 'UPDATE', 'mutasi', id, existing || null, afterSnapshot, null, 'SUCCESS', true)
        ]);
        const changes = batchRes?.[0]?.meta?.changes ?? batchRes?.[0]?.changes;

        if (changes === 0) {
          return error('Conflict: Record has been modified by another user or session', 409, origin);
        }

        return ok({ message: 'Mutasi berhasil diperbarui', row_version: nextRowVersion }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
      } catch (e) {
        console.error('Mutasi PUT error:', e);
        return error('Gagal memperbarui data mutasi', 500, origin);
      }
    }

    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'mutasi', 'delete')) return forbidden(origin);
      try {
        const existing = await env.DB.prepare('SELECT * FROM mutasi_data WHERE id = ?').bind(id).first();
        if (!existing) {
          const countCheck = await env.DB.prepare('SELECT COUNT(*) as c FROM mutasi_data').first();
          if (countCheck && typeof countCheck.c === 'number') return notFound(origin);
        }
        const delStmt = env.DB.prepare('DELETE FROM mutasi_data WHERE id=?').bind(id);
        const { executeAtomicDelete } = await import('../utils/audit.js');
        await executeAtomicDelete(env, user, delStmt, 'mutasi', id, existing);
        return ok({ message: 'Mutasi berhasil dihapus' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
      } catch (e) {
        console.error('Mutasi DELETE error:', e);
        return error('Gagal menghapus data mutasi', 500, origin);
      }
    }
  }

  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'mutasi', 'import')) return forbidden(origin);
    let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
    if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
    if (body.length === 0) return ok({ message: 'No data to import' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });

    const existing = await env.DB.prepare('SELECT * FROM mutasi_data').all();
    const existingMap = new Map();
    (existing.results || []).forEach(s => {
      if (s.tanggal && s.employee_name && s.from_branch_id && s.to_branch_id) {
        existingMap.set(s.tanggal + '_' + s.employee_name.toLowerCase().trim() + '_' + s.from_branch_id + '_' + s.to_branch_id, s);
      }
    });

    // Cache employees for unambiguous linkage
    const empRows = await env.DB.prepare('SELECT id, LOWER(TRIM(full_name)) as name_lower FROM employees WHERE deleted_at IS NULL').all();
    const empCounts = new Map();
    const empIdMap = new Map();
    (empRows.results || []).forEach(e => {
      empCounts.set(e.name_lower, (empCounts.get(e.name_lower) || 0) + 1);
      empIdMap.set(e.name_lower, e.id);
    });

    const stmts = [];
    let inserted = 0;
    let updated = 0;

    for (const item of body) {
      if (!item.tanggal || !item.employee_name || !item.from_branch_id || !item.to_branch_id) continue;
      
      const empNameClean = String(item.employee_name).toLowerCase().trim();
      let empId = null;
      if (item.employee_id) {
        empId = Number(item.employee_id) || null;
      } else if (empCounts.get(empNameClean) === 1) {
        empId = empIdMap.get(empNameClean);
      }

      const key = item.tanggal + '_' + empNameClean + '_' + item.from_branch_id + '_' + item.to_branch_id;
      const status = item.status || 'Proses';

      if (existingMap.has(key)) {
        const oldRow = existingMap.get(key);
        const afterRow = {
          ...oldRow,
          status,
          document_link: item.document_link || null,
          employee_id: empId ?? oldRow.employee_id,
          row_version: (oldRow.row_version ?? 1) + 1
        };
        stmts.push(
          env.DB.prepare(
            `UPDATE mutasi_data SET status = ?, document_link = ?, employee_id = COALESCE(?, employee_id),
             row_version = COALESCE(row_version, 1) + 1, updated_at = datetime('now') WHERE id = ?`
          ).bind(
            status,
            item.document_link || null,
            empId,
            oldRow.id
          ),
          buildAuditStatement(env, user, 'UPDATE', 'mutasi', oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true)
        );
        existingMap.set(key, afterRow);
        updated++;
      } else {
        const newSnapshot = {
          tanggal: item.tanggal,
          employee_id: empId,
          employee_name: item.employee_name,
          from_branch_id: item.from_branch_id,
          to_branch_id: item.to_branch_id,
          status,
          document_link: item.document_link || null,
          row_version: 1
        };
        stmts.push(
          env.DB.prepare(
            `INSERT INTO mutasi_data (tanggal, employee_id, employee_name, from_branch_id, to_branch_id, status, document_link, row_version, updated_at) 
             VALUES (?, ?, ?, ?, ?, ?, ?, 1, datetime('now'))`
          ).bind(
            item.tanggal,
            empId,
            item.employee_name,
            item.from_branch_id,
            item.to_branch_id,
            status,
            item.document_link || null
          ),
          buildCreateAuditStatement(env, user, 'mutasi', 'new', newSnapshot, { source: 'import' })
        );
        inserted++;
      }
    }

    try {
      if (stmts.length > 0) {
        await executeAtomicImportBatch(env, stmts);
      }
      return ok({ message: 'Berhasil mengimport data', inserted, updated, total: body.length }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
    } catch (err) {
      console.error('Mutasi import error:', err);
      return error('Gagal mengimport data mutasi', 500, origin);
    }
  }

  return notFound(origin);
}
