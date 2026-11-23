import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination, resolveSortOrder } from '../utils/pagination.js';
import { requireExportAccess, getExportPagination } from '../utils/export.js';
import { buildAuditStatement, buildCreateAuditStatement, executeAtomicDelete } from '../utils/audit.js';
import { hasTableColumn } from '../utils/schemaCompat.js';
import { checkEntityNotArchived } from '../utils/archive.js';

const OVERTIME_SORT_MAP = Object.freeze({
  date_desc: 'o.date DESC, o.id DESC',
  date_asc: 'o.date ASC, o.id ASC',
  hours_desc: 'COALESCE(o.total_hours, 0) DESC, o.date DESC, o.id DESC',
});

/**
 * Computes net overtime duration (in hours) from start_time, end_time (HH:MM), and break_hours.
 * Supports overnight shifts (e.g., 22:00 -> 08:00 with 1h break = 9h).
 */
export function computeOvertimeHours(startTime, endTime, breakHours = 0) {
  const m1 = String(startTime || '').trim().match(/^(\d{1,2}):(\d{2})/);
  const m2 = String(endTime || '').trim().match(/^(\d{1,2}):(\d{2})/);
  if (!m1 || !m2) return null;
  const h1 = Number(m1[1]);
  const min1 = Number(m1[2]);
  const h2 = Number(m2[1]);
  const min2 = Number(m2[2]);
  if (h1 > 23 || min1 > 59 || h2 > 23 || min2 > 59) return null;

  const startTotalMin = h1 * 60 + min1;
  let endTotalMin = h2 * 60 + min2;
  if (endTotalMin <= startTotalMin) {
    endTotalMin += 24 * 60;
  }
  const grossHours = (endTotalMin - startTotalMin) / 60;
  const breakVal = Number.isFinite(Number(breakHours)) ? Number(breakHours) : 0;
  return Math.max(0, Math.round((grossHours - breakVal) * 100) / 100);
}

function buildOvertimeFilter(url) {
  const search = (url.searchParams.get('search') || '').trim();
  const branch_id = (url.searchParams.get('branch_id') || '').trim();
  const employee_id = (url.searchParams.get('employee_id') || '').trim();
  const reason = (url.searchParams.get('reason') || '').trim();
  const sort = url.searchParams.get('sort') || 'date_desc';

  const conditions = ['1=1'];
  const values = [];
  if (search) {
    conditions.push('(e.full_name LIKE ? OR b.full_name LIKE ? OR o.reason LIKE ?)');
    values.push(`%${search}%`, `%${search}%`, `%${search}%`);
  }
  if (branch_id) { conditions.push('o.branch_id = ?'); values.push(branch_id); }
  if (employee_id) { conditions.push('o.employee_id = ?'); values.push(employee_id); }
  if (reason) { conditions.push('o.reason = ?'); values.push(reason); }

  const where = 'WHERE ' + conditions.join(' AND ');
  const countJoin = search
    ? 'LEFT JOIN branches b ON o.branch_id = b.id LEFT JOIN employees e ON o.employee_id = e.id'
    : '';
  const orderBy = resolveSortOrder(sort, OVERTIME_SORT_MAP, 'date_desc');
  return { where, values, countJoin, orderBy };
}

export async function handleOvertime(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace('/api/overtime', '');

  if (request.method === 'GET' && path === '/export') {
    const { response } = await requireExportAccess(request, env, origin, 'overtime');
    if (response) return response;
    try {
      const { page, limit, offset } = getExportPagination(request.url);
      const { where, values, countJoin, orderBy } = buildOvertimeFilter(url);
      const [countRow, { results }] = await Promise.all([
        env.DB.prepare(`SELECT COUNT(*) as total FROM overtime_records o ${countJoin} ${where}`).bind(...values).first(),
        env.DB.prepare(`
          SELECT o.*, b.full_name as branch_name, e.full_name as employee_name
          FROM overtime_records o
          LEFT JOIN branches b ON o.branch_id = b.id
          LEFT JOIN employees e ON o.employee_id = e.id
          ${where}
          ORDER BY ${orderBy}
          LIMIT ? OFFSET ?
        `).bind(...values, limit, offset).all()
      ]);
      return paginated(results || [], countRow?.total || 0, page, limit, origin, env?.ENVIRONMENT);
    } catch (e) {
      console.error('Overtime export error:', e);
      return error('Gagal mengekspor data lembur', 500, origin);
    }
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'overtime', 'read')) return forbidden(origin);

  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'overtime', 'import')) return forbidden(origin);
    return error('Not found', 404, origin);
  }

  // GET list
  if (request.method === 'GET' && (path === '' || path === '/')) {
    try {
      const { page, limit, offset } = getPagination(request.url);
      const { where, values, countJoin, orderBy } = buildOvertimeFilter(url);
      const includeStats = url.searchParams.get('include_stats') === '1';

      const queries = [
        env.DB.prepare(`SELECT COUNT(*) as total FROM overtime_records o ${countJoin} ${where}`).bind(...values).first(),
        env.DB.prepare(`
          SELECT o.*, b.full_name as branch_name, e.full_name as employee_name
          FROM overtime_records o
          LEFT JOIN branches b ON o.branch_id = b.id
          LEFT JOIN employees e ON o.employee_id = e.id
          ${where}
          ORDER BY ${orderBy}
          LIMIT ? OFFSET ?
        `).bind(...values, limit, offset).all(),
      ];

      if (includeStats) {
        queries.push(
          env.DB.prepare(`
            SELECT
              COUNT(*) as total_records,
              COALESCE(SUM(total_hours), 0) as total_hours,
              SUM(CASE WHEN strftime('%Y-%m', date) = strftime('%Y-%m', date('now', '+7 hours')) THEN 1 ELSE 0 END) as current_month
            FROM overtime_records
          `).first()
        );
      }

      const [countRow, { results }, statsRow] = await Promise.all(queries);
      const totalRecords = Number(statsRow?.total_records || 0);
      const totalHours = Number(statsRow?.total_hours || 0);
      const extra = includeStats ? {
        stats: {
          totalRecords,
          totalHours,
          avgDuration: totalRecords > 0 ? Number((totalHours / totalRecords).toFixed(1)) : 0,
          currentMonth: Number(statsRow?.current_month || 0),
        }
      } : null;

      return paginated(results || [], countRow?.total || 0, page, limit, origin, env?.ENVIRONMENT, extra);
    } catch (e) {
      console.error('Overtime GET error:', e);
      return error('Gagal mengambil data lembur', 500, origin);
    }
  }

  // GET single
  if (request.method === 'GET' && idMatch) {
    try {
      const { results } = await env.DB.prepare(`
        SELECT o.*, b.full_name as branch_name, e.full_name as employee_name
        FROM overtime_records o
        LEFT JOIN branches b ON o.branch_id = b.id
        LEFT JOIN employees e ON o.employee_id = e.id
        WHERE o.id = ?
      `).bind(idMatch[1]).all();
      
      if (!results || results.length === 0) return notFound(origin);
      return ok(results[0], 200, origin);
    } catch (e) {
      console.error('Overtime GET single error:', e);
      return error('Gagal mengambil detail lembur', 500, origin);
    }
  }

  // POST (Create)
  if (request.method === 'POST' && (path === '' || path === '/')) {
    if (!hasPermission(user, 'overtime', 'create')) return forbidden(origin);
    let body;
    try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
    if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);

    if (!body.date || !body.branch_id || !body.employee_id || !body.start_time || !body.end_time) {
      return error('date, branch_id, employee_id, start_time, and end_time are required', 400, origin);
    }

    const parsedBranchId = Number(body.branch_id);
    const parsedEmpId = Number(body.employee_id);
    if (!Number.isInteger(parsedBranchId) || parsedBranchId <= 0) return error('Invalid branch_id', 400, origin);
    if (!Number.isInteger(parsedEmpId) || parsedEmpId <= 0) return error('Invalid employee_id', 400, origin);

    const parsedBreak = body.break_hours === undefined || body.break_hours === null || body.break_hours === ''
      ? 0
      : Number(body.break_hours);
    if (!Number.isFinite(parsedBreak) || parsedBreak < 0 || parsedBreak > 12) {
      return error('break_hours harus berupa angka antara 0 dan 12', 400, origin);
    }

    let computedTotalHours;
    if (body.total_hours !== undefined && body.total_hours !== null && body.total_hours !== '') {
      const explicitTotal = Number(body.total_hours);
      if (!Number.isFinite(explicitTotal) || explicitTotal < 0 || explicitTotal > 24) {
        return error('total_hours harus berupa angka antara 0 dan 24', 400, origin);
      }
      computedTotalHours = Math.round(explicitTotal * 100) / 100;
    } else {
      computedTotalHours = computeOvertimeHours(body.start_time, body.end_time, parsedBreak);
      if (computedTotalHours === null) {
        return error('Format start_time / end_time tidak valid (gunakan HH:MM)', 400, origin);
      }
    }

    try {
      const branchRow = await env.DB.prepare('SELECT id FROM branches WHERE id = ? AND deleted_at IS NULL').bind(parsedBranchId).first();
      if (!branchRow) {
        const bCount = await env.DB.prepare('SELECT COUNT(*) as c FROM branches').first();
        if (bCount && typeof bCount.c === 'number' && bCount.c > 0) {
          return error('Branch not found (invalid branch_id foreign key)', 400, origin);
        }
      }

      const empRow = await env.DB.prepare('SELECT id FROM employees WHERE id = ? AND deleted_at IS NULL').bind(parsedEmpId).first();
      if (!empRow) {
        const eCount = await env.DB.prepare('SELECT COUNT(*) as c FROM employees').first();
        if (eCount && typeof eCount.c === 'number' && eCount.c > 0) {
          return error('Employee not found or inactive (invalid employee_id foreign key)', 400, origin);
        }
      }
      if ((await checkEntityNotArchived(env.DB, 'employees', parsedEmpId)).archived) {
        return error('Karyawan berstatus arsip dan tidak dapat ditambahkan pada data lembur baru', 400, origin);
      }

      const dup = await env.DB.prepare(
        'SELECT id FROM overtime_records WHERE employee_id = ? AND date = ? AND start_time = ? AND end_time = ?'
      ).bind(parsedEmpId, body.date, body.start_time, body.end_time).first();
      if (dup) {
        return error('Data lembur untuk karyawan, tanggal, dan jam yang sama sudah tercatat', 409, origin);
      }

      const safeReason = body.reason || body.description || '';
      const hasRowVerCol = await hasTableColumn(env.DB, 'overtime_records', 'row_version');
      const insertStmt = hasRowVerCol
        ? env.DB.prepare(`
            INSERT INTO overtime_records (date, branch_id, employee_id, start_time, end_time, break_hours, total_hours, reason, row_version, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1, datetime('now'))
          `).bind(
            body.date,
            parsedBranchId,
            parsedEmpId,
            body.start_time,
            body.end_time,
            parsedBreak,
            computedTotalHours,
            safeReason
          )
        : env.DB.prepare(`
            INSERT INTO overtime_records (date, branch_id, employee_id, start_time, end_time, break_hours, total_hours, reason, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
          `).bind(
            body.date,
            parsedBranchId,
            parsedEmpId,
            body.start_time,
            body.end_time,
            parsedBreak,
            computedTotalHours,
            safeReason
          );
      const auditStmt = buildCreateAuditStatement(env, user, 'overtime', 'new', {
        date: body.date,
        branch_id: parsedBranchId,
        employee_id: parsedEmpId,
        start_time: body.start_time,
        end_time: body.end_time,
        break_hours: parsedBreak,
        total_hours: computedTotalHours,
        reason: safeReason || null
      });
      const batchRes = await env.DB.batch([insertStmt, auditStmt]);
      const createdId = batchRes?.[0]?.meta?.last_row_id ?? 1;

      return ok(
        { id: createdId, row_version: 1, total_hours: computedTotalHours, message: 'Data lembur berhasil ditambahkan' },
        201,
        origin,
        undefined,
        { 'X-Audit-Batched': 'true' }
      );
    } catch (e) {
      console.error('Overtime POST error:', e);
      return error('Gagal menambahkan data lembur', 500, origin);
    }
  }

  // PUT / PATCH (Update)
  if ((request.method === 'PUT' || request.method === 'PATCH') && idMatch) {
    if (!hasPermission(user, 'overtime', 'update')) return forbidden(origin);
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

    try {
      const id = idMatch[1];
      const existing = await env.DB.prepare('SELECT * FROM overtime_records WHERE id = ?').bind(id).first();
      if (!existing) {
        const countCheck = await env.DB.prepare('SELECT COUNT(*) as c FROM overtime_records').first();
        if (countCheck && typeof countCheck.c === 'number') {
          return notFound(origin);
        }
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

      if (body.branch_id !== undefined && body.branch_id !== null && body.branch_id !== '') {
        const parsedBranchId = Number(body.branch_id);
        if (!Number.isInteger(parsedBranchId) || parsedBranchId <= 0) return error('Invalid branch_id', 400, origin);
        const branchRow = await env.DB.prepare('SELECT id FROM branches WHERE id = ? AND deleted_at IS NULL').bind(parsedBranchId).first();
        if (!branchRow) {
          const bCount = await env.DB.prepare('SELECT COUNT(*) as c FROM branches').first();
          if (bCount && typeof bCount.c === 'number' && bCount.c > 0) {
            return error('Branch not found (invalid branch_id foreign key)', 400, origin);
          }
        }
      }
      if (body.employee_id !== undefined && body.employee_id !== null && body.employee_id !== '') {
        const parsedEmpId = Number(body.employee_id);
        if (!Number.isInteger(parsedEmpId) || parsedEmpId <= 0) return error('Invalid employee_id', 400, origin);
        const empRow = await env.DB.prepare('SELECT id FROM employees WHERE id = ? AND deleted_at IS NULL').bind(parsedEmpId).first();
        if (!empRow) {
          const eCount = await env.DB.prepare('SELECT COUNT(*) as c FROM employees').first();
          if (eCount && typeof eCount.c === 'number' && eCount.c > 0) {
            return error('Employee not found or inactive (invalid employee_id foreign key)', 400, origin);
          }
        }
      }

      const effStart = body.start_time ?? existing.start_time;
      const effEnd = body.end_time ?? existing.end_time;
      const effBreak = body.break_hours !== undefined && body.break_hours !== null && body.break_hours !== ''
        ? Number(body.break_hours)
        : Number(existing.break_hours ?? 0);
      if (!Number.isFinite(effBreak) || effBreak < 0 || effBreak > 12) {
        return error('break_hours harus berupa angka antara 0 dan 12', 400, origin);
      }

      let effTotalHours;
      if (body.total_hours !== undefined && body.total_hours !== null && body.total_hours !== '') {
        const explicitTotal = Number(body.total_hours);
        if (!Number.isFinite(explicitTotal) || explicitTotal < 0 || explicitTotal > 24) {
          return error('total_hours harus berupa angka antara 0 dan 24', 400, origin);
        }
        effTotalHours = Math.round(explicitTotal * 100) / 100;
      } else if (body.start_time !== undefined || body.end_time !== undefined || body.break_hours !== undefined) {
        const recalculated = computeOvertimeHours(effStart, effEnd, effBreak);
        if (recalculated === null) {
          return error('Format start_time / end_time tidak valid (gunakan HH:MM)', 400, origin);
        }
        effTotalHours = recalculated;
      } else {
        effTotalHours = existing.total_hours;
      }

      const baseVersion = isExistingNull ? clientVersion : existing.row_version;
      const nextRowVersion = baseVersion + 1;
      const updatedSnapshot = {
        ...existing,
        date: body.date ?? existing.date,
        branch_id: body.branch_id !== undefined && body.branch_id !== null && body.branch_id !== '' ? Number(body.branch_id) : existing.branch_id,
        employee_id: body.employee_id !== undefined && body.employee_id !== null && body.employee_id !== '' ? Number(body.employee_id) : existing.employee_id,
        start_time: effStart,
        end_time: effEnd,
        break_hours: effBreak,
        total_hours: effTotalHours,
        reason: body.reason !== undefined ? body.reason : existing.reason,
        row_version: nextRowVersion
      };

      const updateStmt = env.DB.prepare(`
        UPDATE overtime_records
        SET date = COALESCE(?, date),
            branch_id = COALESCE(?, branch_id),
            employee_id = COALESCE(?, employee_id),
            start_time = COALESCE(?, start_time),
            end_time = COALESCE(?, end_time),
            break_hours = ?,
            total_hours = ?,
            reason = COALESCE(?, reason),
            row_version = COALESCE(row_version, ?) + 1,
            updated_at = datetime('now')
        WHERE id = ? AND (row_version = ? OR (row_version IS NULL AND (? = 1 OR ? = 0)))
      `).bind(
        body.date ?? null,
        body.branch_id ?? null,
        body.employee_id ?? null,
        body.start_time ?? null,
        body.end_time ?? null,
        effBreak,
        effTotalHours,
        body.reason ?? null,
        clientVersion,
        id,
        clientVersion,
        clientVersion,
        clientVersion
      );

      const auditStmt = buildAuditStatement(
        env,
        user,
        'UPDATE',
        'overtime',
        id,
        existing,
        updatedSnapshot,
        null,
        'SUCCESS',
        true
      );

      const batchRes = await env.DB.batch([updateStmt, auditStmt]);
      const changes = batchRes?.[0]?.meta?.changes ?? batchRes?.[0]?.changes;
      if (changes === 0) {
        return error('Conflict: Record has been modified by another user or session', 409, origin);
      }

      return ok(
        { message: 'Data lembur berhasil diupdate', row_version: nextRowVersion },
        200,
        origin,
        undefined,
        { 'X-Audit-Batched': 'true' }
      );
    } catch (e) {
      console.error('Overtime PUT error:', e);
      return error('Gagal memperbarui data lembur', 500, origin);
    }
  }

  // DELETE
  if (request.method === 'DELETE' && idMatch) {
    if (!hasPermission(user, 'overtime', 'delete')) return forbidden(origin);
    try {
      const id = idMatch[1];
      const existing = await env.DB.prepare('SELECT id FROM overtime_records WHERE id = ?').bind(id).first();
      if (!existing) {
        const countCheck = await env.DB.prepare('SELECT COUNT(*) as c FROM overtime_records').first();
        if (countCheck && typeof countCheck.c === 'number') {
          return notFound(origin);
        }
      }
      const delStmt = env.DB.prepare('DELETE FROM overtime_records WHERE id=?').bind(id);
      await executeAtomicDelete(env, user, delStmt, 'overtime', id, existing);
      return ok({ message: 'Data lembur berhasil dihapus' }, 200, origin, { 'X-Audit-Batched': 'true' });
    } catch (e) {
      console.error('Overtime DELETE error:', e);
      return error('Gagal menghapus data lembur', 500, origin);
    }
  }

  return notFound(origin);
}
