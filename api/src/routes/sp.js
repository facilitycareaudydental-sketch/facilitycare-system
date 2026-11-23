import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination, resolveSortOrder, buildDateOrderBy } from '../utils/pagination.js';
import { requireExportAccess, getExportPagination } from '../utils/export.js';
import { buildAuditStatement, buildCreateAuditStatement, executeAtomicImportBatch } from '../utils/audit.js';
import { hasTableColumn } from '../utils/schemaCompat.js';
import { checkEntityNotArchived } from '../utils/archive.js';

async function getSpOrderBy(db, rawSort) {
  const includeUpdatedAt = await hasTableColumn(db, 'sp_data', 'updated_at');
  const sortMap = {
    date_desc: buildDateOrderBy('s.tanggal', { direction: 'DESC', alias: 's', includeUpdatedAt, includeCreatedAt: true, idDir: 'DESC' }),
    date_asc: buildDateOrderBy('s.tanggal', { direction: 'ASC', alias: 's', includeUpdatedAt, includeCreatedAt: true, idDir: 'DESC' }),
    name_asc: 's.employee_name ASC, s.id ASC',
    expiry_asc: buildDateOrderBy('s.akhir_sp', { direction: 'ASC', alias: 's', includeUpdatedAt, includeCreatedAt: true, idDir: 'DESC' }),
  };
  return resolveSortOrder(rawSort, sortMap, 'date_desc');
}

async function buildSPFilter(db, url) {
  const search = (url.searchParams.get('search') || '').trim();
  const branch_id = (url.searchParams.get('branch_id') || '').trim();
  const division = (url.searchParams.get('division') || '').trim();
  const filter = (url.searchParams.get('filter') || '').trim();
  const sort = url.searchParams.get('sort') || 'date_desc';

  const conditions = ['1=1'];
  const values = [];
  if (search) {
    conditions.push('(s.employee_name LIKE ? OR b.full_name LIKE ? OR s.division LIKE ? OR s.sp_type LIKE ?)');
    values.push(`%${search}%`, `%${search}%`, `%${search}%`, `%${search}%`);
  }
  if (branch_id) { conditions.push('s.branch_id = ?'); values.push(branch_id); }
  if (division) { conditions.push('s.division = ?'); values.push(division); }

  if (filter === 'active') {
    conditions.push("(s.akhir_sp IS NULL OR s.akhir_sp = '' OR date(s.akhir_sp) >= date('now', '+7 hours'))");
  } else if (filter === 'expiring') {
    conditions.push("(s.akhir_sp IS NOT NULL AND s.akhir_sp != '' AND CAST(julianday(s.akhir_sp) - julianday(date('now', '+7 hours')) AS INTEGER) BETWEEN 0 AND 30)");
  } else if (filter === 'expired') {
    conditions.push("(s.akhir_sp IS NOT NULL AND s.akhir_sp != '' AND date(s.akhir_sp) < date('now', '+7 hours'))");
  } else if (filter === 'sp1') {
    conditions.push("UPPER(TRIM(s.sp_type)) = 'SP 1'");
  } else if (filter === 'sp2') {
    conditions.push("UPPER(TRIM(s.sp_type)) = 'SP 2'");
  } else if (filter === 'sp3') {
    conditions.push("UPPER(TRIM(s.sp_type)) = 'SP 3'");
  } else if (filter === 'teguran') {
    conditions.push("UPPER(COALESCE(s.sp_type, '')) LIKE '%TEGURAN%'");
  }

  const where = 'WHERE ' + conditions.join(' AND ');
  const countJoin = search ? 'LEFT JOIN branches b ON s.branch_id = b.id' : '';
  const orderBy = await getSpOrderBy(db, sort);
  return { where, values, countJoin, orderBy };
}

export async function handleSP(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace('/api/sp', '');

  if (request.method === 'GET' && path === '/export') {
    const { response } = await requireExportAccess(request, env, origin, 'sp');
    if (response) return response;
    try {
      const { page, limit, offset } = getExportPagination(request.url);
      const { where, values, countJoin, orderBy } = await buildSPFilter(env.DB, url);
      const [countRow, { results }] = await Promise.all([
        env.DB.prepare(`SELECT COUNT(*) as total FROM sp_data s ${countJoin} ${where}`).bind(...values).first(),
        env.DB.prepare(`
          SELECT s.*, b.full_name as branch_name
          FROM sp_data s
          LEFT JOIN branches b ON s.branch_id = b.id
          ${where}
          ORDER BY ${orderBy}
          LIMIT ? OFFSET ?
        `).bind(...values, limit, offset).all()
      ]);
      return paginated(results || [], countRow?.total || 0, page, limit, origin, env?.ENVIRONMENT);
    } catch (e) {
      console.error('SP export error:', e);
      return error('Gagal mengekspor data SP', 500, origin);
    }
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'sp', 'read')) return forbidden(origin);

  const idMatch = path.match(/^\/(\d+)$/);

  // GET list
  if (request.method === 'GET' && (path === '' || path === '/')) {
    try {
      const { page, limit, offset } = getPagination(request.url);
      const { where, values, countJoin, orderBy } = await buildSPFilter(env.DB, url);
      const includeStats = url.searchParams.get('include_stats') === '1';

      const queries = [
        env.DB.prepare(`SELECT COUNT(*) as total FROM sp_data s ${countJoin} ${where}`).bind(...values).first(),
        env.DB.prepare(`
          SELECT s.*, b.full_name as branch_name
          FROM sp_data s
          LEFT JOIN branches b ON s.branch_id = b.id
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
              SUM(CASE WHEN akhir_sp IS NULL OR akhir_sp = '' OR date(akhir_sp) >= date('now', '+7 hours') THEN 1 ELSE 0 END) as active,
              SUM(CASE WHEN akhir_sp IS NOT NULL AND akhir_sp != '' AND CAST(julianday(akhir_sp) - julianday(date('now', '+7 hours')) AS INTEGER) BETWEEN 0 AND 30 THEN 1 ELSE 0 END) as expiring,
              SUM(CASE WHEN akhir_sp IS NOT NULL AND akhir_sp != '' AND date(akhir_sp) < date('now', '+7 hours') THEN 1 ELSE 0 END) as expired,
              SUM(CASE WHEN UPPER(TRIM(sp_type)) = 'SP 1' THEN 1 ELSE 0 END) as sp1,
              SUM(CASE WHEN UPPER(TRIM(sp_type)) = 'SP 2' THEN 1 ELSE 0 END) as sp2,
              SUM(CASE WHEN UPPER(TRIM(sp_type)) = 'SP 3' THEN 1 ELSE 0 END) as sp3,
              SUM(CASE WHEN UPPER(COALESCE(sp_type, '')) LIKE '%TEGURAN%' THEN 1 ELSE 0 END) as teguran,
              COUNT(DISTINCT LOWER(TRIM(COALESCE(employee_name, '')))) as unique_employees
            FROM sp_data
          `).first()
        );
      }

      const [countRow, { results }, statsRow] = await Promise.all(queries);
      const extra = includeStats ? {
        stats: {
          total: Number(statsRow?.total || 0),
          active: Number(statsRow?.active || 0),
          expiring: Number(statsRow?.expiring || 0),
          expired: Number(statsRow?.expired || 0),
          sp1: Number(statsRow?.sp1 || 0),
          sp2: Number(statsRow?.sp2 || 0),
          sp3: Number(statsRow?.sp3 || 0),
          teguran: Number(statsRow?.teguran || 0),
          uniqueEmployees: Number(statsRow?.unique_employees || 0),
        }
      } : null;

      return paginated(results || [], countRow?.total || 0, page, limit, origin, env?.ENVIRONMENT, extra);
    } catch (e) {
      console.error('SP GET error:', e);
      return error('Gagal mengambil data SP', 500, origin);
    }
  }

  // POST create
  if (request.method === 'POST' && (path === '' || path === '/')) {
    if (!hasPermission(user, 'sp', 'write')) return forbidden(origin);
    try {
      let body;
      try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
      if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
      const { tanggal, employee_id, employee_name, division, branch_id, akhir_sp, sp_type, status, document_link } = body;
      if (!employee_name || typeof employee_name !== 'string' || !employee_name.trim()) return error('employee_name required', 400, origin);
      if (tanggal && akhir_sp && String(akhir_sp).trim() && String(akhir_sp).trim() < String(tanggal).trim()) {
        return error('akhir_sp tidak boleh lebih awal dari tanggal', 400, origin);
      }

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

      if (resolvedEmpId && (await checkEntityNotArchived(env.DB, 'employees', resolvedEmpId)).archived) {
        return error('Karyawan berstatus arsip dan tidak dapat diterbitkan SP baru', 400, origin);
      }

      if (branch_id !== undefined && branch_id !== null && branch_id !== '') {
        const parsedBranchId = Number(branch_id);
        if (!Number.isInteger(parsedBranchId) || parsedBranchId <= 0) return error('Invalid branch_id', 400, origin);
        const branchRow = await env.DB.prepare('SELECT id FROM branches WHERE id = ? AND deleted_at IS NULL').bind(parsedBranchId).first();
        if (!branchRow) {
          const bCount = await env.DB.prepare('SELECT COUNT(*) as c FROM branches').first();
          if (bCount && typeof bCount.c === 'number' && bCount.c > 0) {
            return error('Branch not found (invalid branch_id foreign key)', 400, origin);
          }
        }
      }

      if (tanggal && sp_type) {
        const dup = await env.DB.prepare(
          'SELECT id FROM sp_data WHERE tanggal = ? AND LOWER(TRIM(employee_name)) = LOWER(TRIM(?)) AND LOWER(TRIM(sp_type)) = LOWER(TRIM(?))'
        ).bind(tanggal, employee_name, sp_type).first();
        if (dup) {
          return error('Data SP dengan karyawan, tanggal, dan tipe yang sama sudah tercatat', 409, origin);
        }
      }

      const normalizedStatus = status !== null && status !== undefined && status !== '' ? status : '';
      const cleanEmpName = employee_name.trim();
      const newSnapshot = {
        tanggal: tanggal || null,
        employee_id: resolvedEmpId,
        employee_name: cleanEmpName,
        division: division || null,
        branch_id: branch_id || null,
        akhir_sp: akhir_sp || null,
        sp_type: sp_type || '',
        status: normalizedStatus,
        document_link: document_link || null,
        row_version: 1
      };

      const hasEmpIdCol = await hasTableColumn(env.DB, 'sp_data', 'employee_id');
      const insertStmt = hasEmpIdCol
        ? env.DB.prepare(
            "INSERT INTO sp_data (tanggal, employee_id, employee_name, division, branch_id, akhir_sp, sp_type, status, document_link, row_version, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1, datetime('now'))"
          ).bind(tanggal || null, resolvedEmpId, cleanEmpName, division || null, branch_id || null, akhir_sp || null, sp_type || '', normalizedStatus, document_link || null)
        : env.DB.prepare(
            "INSERT INTO sp_data (tanggal, employee_name, division, branch_id, akhir_sp, sp_type, status, document_link, row_version, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1, datetime('now'))"
          ).bind(tanggal || null, cleanEmpName, division || null, branch_id || null, akhir_sp || null, sp_type || '', normalizedStatus, document_link || null);

      const batchRes = await env.DB.batch([
        insertStmt,
        buildCreateAuditStatement(env, user, 'sp', 'new', newSnapshot)
      ]);
      const newId = batchRes?.[0]?.meta?.last_row_id;

      return ok({ id: newId, row_version: 1, message: 'SP berhasil ditambahkan' }, 201, origin, undefined, { 'X-Audit-Batched': 'true' });
    } catch (e) {
      console.error('SP POST error:', e);
      return error('Gagal menambahkan data SP', 500, origin);
    }
  }

  // GET single / PUT / DELETE by id
  if (idMatch) {
    const id = idMatch[1];

    if (request.method === 'GET') {
      const row = await env.DB.prepare(
        'SELECT s.*, b.full_name as branch_name FROM sp_data s LEFT JOIN branches b ON s.branch_id = b.id WHERE s.id = ?'
      ).bind(id).first();
      if (!row) return notFound(origin);
      return ok(row, 200, origin);
    }

    if (request.method === 'PUT' || request.method === 'PATCH') {
      if (!hasPermission(user, 'sp', 'write')) return forbidden(origin);
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

        const existing = await env.DB.prepare('SELECT * FROM sp_data WHERE id = ?').bind(id).first();
        if (!existing) {
          const countCheck = await env.DB.prepare('SELECT COUNT(*) as c FROM sp_data').first();
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

        const { tanggal, employee_id, employee_name, division, branch_id, akhir_sp, sp_type, status, document_link } = body;

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

        if (resolvedEmpId && (await checkEntityNotArchived(env.DB, 'employees', resolvedEmpId)).archived) {
        return error('Karyawan berstatus arsip dan tidak dapat diterbitkan SP baru', 400, origin);
      }

      if (branch_id !== undefined && branch_id !== null && branch_id !== '') {
          const parsedBranchId = Number(branch_id);
          if (!Number.isInteger(parsedBranchId) || parsedBranchId <= 0) return error('Invalid branch_id', 400, origin);
          const branchRow = await env.DB.prepare('SELECT id FROM branches WHERE id = ? AND deleted_at IS NULL').bind(parsedBranchId).first();
          if (!branchRow) {
            const bCount = await env.DB.prepare('SELECT COUNT(*) as c FROM branches').first();
            if (bCount && typeof bCount.c === 'number' && bCount.c > 0) {
              return error('Branch not found (invalid branch_id foreign key)', 400, origin);
            }
          }
        }

        const effTanggal = tanggal || existing?.tanggal || null;
        const effAkhirSp = Object.prototype.hasOwnProperty.call(body, 'akhir_sp')
          ? (akhir_sp && String(akhir_sp).trim() ? String(akhir_sp).trim() : null)
          : (existing?.akhir_sp || null);
        const effDocLink = Object.prototype.hasOwnProperty.call(body, 'document_link')
          ? (document_link && String(document_link).trim() ? String(document_link).trim() : null)
          : (existing?.document_link || null);

        if (effTanggal && effAkhirSp && effAkhirSp < String(effTanggal).trim()) {
          return error('akhir_sp tidak boleh lebih awal dari tanggal', 400, origin);
        }

        const baseVersion = isExistingNull ? clientVersion : existing?.row_version;
        const nextRowVersion = baseVersion + 1;
        const afterSnapshot = {
          ...(existing || { id }),
          tanggal: effTanggal,
          employee_id: resolvedEmpId,
          employee_name: employee_name || existing?.employee_name || null,
          division: division || existing?.division || null,
          branch_id: branch_id || existing?.branch_id || null,
          akhir_sp: effAkhirSp,
          sp_type: sp_type || existing?.sp_type || null,
          status: status !== undefined ? status : existing?.status,
          document_link: effDocLink,
          row_version: nextRowVersion
        };

        const updateStmt = env.DB.prepare(
          `UPDATE sp_data SET tanggal=COALESCE(?,tanggal), employee_id=?, employee_name=COALESCE(?,employee_name),
           division=COALESCE(?,division), branch_id=COALESCE(?,branch_id), akhir_sp=?, sp_type=COALESCE(?,sp_type),
           status=COALESCE(?,status), document_link=?, row_version=COALESCE(row_version, ?) + 1, updated_at=datetime('now')
           WHERE id=? AND (row_version = ? OR (row_version IS NULL AND (? = 1 OR ? = 0)))`
        ).bind(tanggal || null, resolvedEmpId, employee_name || null, division || null, branch_id || null, effAkhirSp, sp_type || null, status !== undefined ? status : null, effDocLink, clientVersion, id, clientVersion, clientVersion, clientVersion);

        const batchRes = await env.DB.batch([
          updateStmt,
          buildAuditStatement(env, user, 'UPDATE', 'sp', id, existing || null, afterSnapshot, null, 'SUCCESS', true)
        ]);
        const changes = batchRes?.[0]?.meta?.changes ?? batchRes?.[0]?.changes;

        if (changes === 0) {
          return error('Conflict: Record has been modified by another user or session', 409, origin);
        }

        return ok({ message: 'SP berhasil diperbarui', row_version: nextRowVersion }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
      } catch (e) {
        console.error('SP PUT error:', e);
        return error('Gagal memperbarui data SP', 500, origin);
      }
    }

    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'sp', 'delete')) return forbidden(origin);
      try {
        const existing = await env.DB.prepare('SELECT * FROM sp_data WHERE id = ?').bind(id).first();
        if (!existing) {
          const countCheck = await env.DB.prepare('SELECT COUNT(*) as c FROM sp_data').first();
          if (countCheck && typeof countCheck.c === 'number') return notFound(origin);
        }
        const delStmt = env.DB.prepare('DELETE FROM sp_data WHERE id=?').bind(id);
        const { executeAtomicDelete } = await import('../utils/audit.js');
        await executeAtomicDelete(env, user, delStmt, 'sp', id, existing);
        return ok({ message: 'SP berhasil dihapus' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
      } catch (e) {
        console.error('SP DELETE error:', e);
        return error('Gagal menghapus data SP', 500, origin);
      }
    }
  }

  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'sp', 'import')) return forbidden(origin);
    let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
    if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
    if (body.length === 0) return ok({ message: 'No data to import' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });

    const existing = await env.DB.prepare('SELECT * FROM sp_data').all();
    const existingMap = new Map();
    (existing.results || []).forEach(s => {
      if (s.branch_id && s.tanggal && s.employee_name && s.sp_type) {
        existingMap.set(s.branch_id + '_' + s.tanggal + '_' + s.employee_name.toLowerCase().trim() + '_' + s.sp_type.toLowerCase().trim(), s);
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
      if (!item.branch_id || !item.tanggal || !item.employee_name || !item.sp_type) continue;
      
      const empNameClean = String(item.employee_name).toLowerCase().trim();
      let empId = null;
      if (item.employee_id) {
        empId = Number(item.employee_id) || null;
      } else if (empCounts.get(empNameClean) === 1) {
        empId = empIdMap.get(empNameClean);
      }

      const key = item.branch_id + '_' + item.tanggal + '_' + empNameClean + '_' + String(item.sp_type).toLowerCase().trim();
      const status = item.status || 'Aktif';

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
            `UPDATE sp_data SET status = ?, document_link = ?, employee_id = COALESCE(?, employee_id),
             row_version = COALESCE(row_version, 1) + 1, updated_at = datetime('now') WHERE id = ?`
          ).bind(
            status,
            item.document_link || null,
            empId,
            oldRow.id
          ),
          buildAuditStatement(env, user, 'UPDATE', 'sp', oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true)
        );
        existingMap.set(key, afterRow);
        updated++;
      } else {
        const newSnapshot = {
          tanggal: item.tanggal,
          employee_id: empId,
          employee_name: item.employee_name,
          branch_id: item.branch_id,
          sp_type: item.sp_type,
          status,
          document_link: item.document_link || null,
          row_version: 1
        };
        stmts.push(
          env.DB.prepare(
            `INSERT INTO sp_data (tanggal, employee_id, employee_name, branch_id, sp_type, status, document_link, row_version, updated_at) 
             VALUES (?, ?, ?, ?, ?, ?, ?, 1, datetime('now'))`
          ).bind(
            item.tanggal,
            empId,
            item.employee_name,
            item.branch_id,
            item.sp_type,
            status,
            item.document_link || null
          ),
          buildCreateAuditStatement(env, user, 'sp', 'new', newSnapshot, { source: 'import' })
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
      console.error('SP import error:', err);
      return error('Gagal mengimport data SP', 500, origin);
    }
  }

  return notFound(origin);
}
