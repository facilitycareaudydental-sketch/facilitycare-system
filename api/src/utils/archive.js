/**
 * api/src/utils/archive.js
 *
 * Unified Single Archive & Restore Lifecycle Engine (Workstream F)
 *
 * Active Archival Tables (11 Active Operational & Master Entities):
 *   1. employees          (/api/employees/:id/archive & /restore)
 *   2. contracts          (/api/contracts/:id/archive & /restore)
 *   3. branches           (/api/branches/:id/archive & /restore)
 *   4. issues             (/api/issues/:id/archive & /restore)
 *   5. activity_schedule  (/api/schedule/:id/archive & /restore)
 *   6. sop                (/api/sop/:id/archive & /restore)
 *   7. master_checklist   (/api/checklist/:id/archive & /restore)
 *   8. inspection_reports (/api/reports/inspection/:id/archive & /restore)
 *   9. cleaning_reports   (/api/reports/cleaning/:id/archive & /restore)
 *   10. fogging_reports   (/api/reports/fogging/:id/archive & /restore)
 *   11. basecamp_reports  (/api/reports/basecamp/:id/archive & /restore)
 *
 * Excluded by Design (No archived_at lifecycle required):
 *   - System/Auth/Audit tables: users, roles, sessions, audit_logs, auth_cooldowns,
 *     sync_outbox, sync_logs, import_jobs, import_job_items, pic_list.
 *   - Reference standards & immutable logs: master_hygiene_standards,
 *     master_dac_standards, training, sp_data, mutasi_data, os_migrations, overtime_records.
 *   - Retired archive-only tables (Owner Addendum): relievers, one_on_one,
 *     supply_requests, master_forms.
 */

import { ok, error, forbidden, notFound } from './response.js';
import { buildAuditStatement } from './audit.js';
import { hasTableColumn } from './schemaCompat.js';
import { isManagementBranch } from './branchMaster.js';
import { VIEWER_ROLES } from './auth.js';

export const ARCHIVABLE_ROUTE_MAP = Object.freeze({
  employees:            { table: 'employees',          moduleKey: 'employees', label: 'Employee' },
  contracts:            { table: 'contracts',          moduleKey: 'contracts', label: 'Kontrak' },
  branches:             { table: 'branches',           moduleKey: 'branches',  label: 'Cabang' },
  issues:               { table: 'issues',             moduleKey: 'issues',    label: 'Issue' },
  schedule:             { table: 'activity_schedule',  moduleKey: 'schedule',  label: 'Jadwal' },
  sop:                  { table: 'sop',                moduleKey: 'sop',       label: 'SOP' },
  checklist:            { table: 'master_checklist',   moduleKey: 'checklist', label: 'Checklist' },
  'reports/inspection': { table: 'inspection_reports', moduleKey: 'reports',   label: 'Laporan Inspeksi' },
  'reports/cleaning':   { table: 'cleaning_reports',   moduleKey: 'reports',   label: 'Laporan Cleaning' },
  'reports/fogging':    { table: 'fogging_reports',    moduleKey: 'reports',   label: 'Laporan Fogging' },
  'reports/basecamp':   { table: 'basecamp_reports',   moduleKey: 'reports',   label: 'Laporan Basecamp' },
});

/**
 * Verifies whether a caller is authorized to view archived records (?archived=true).
 * Per Single Archive Governance Matrix: only superadmin and viewer roles may view archived records;
 * field operational roles (fc_specialist_leader, editor_khusus) are denied.
 */
export function canViewArchivedRecords(user) {
  if (!user || !user.role) return false;
  const role = String(user.role).trim().toLowerCase();
  return role === 'superadmin' || VIEWER_ROLES.has(role);
}

/**
 * Checks whether an employee or branch record is currently archived (Transaction Barrier).
 */
export async function checkEntityNotArchived(db, table, id) {
  if (!db || !id) return { archived: false };
  if (!(await hasTableColumn(db, table, 'archived_at'))) return { archived: false };
  const row = await db.prepare(`SELECT id, archived_at FROM ${table} WHERE id = ?`).bind(Number(id)).first();
  if (row && row.archived_at !== null && row.archived_at !== undefined && String(row.archived_at).trim() !== '') {
    return { archived: true, archived_at: row.archived_at };
  }
  return { archived: false };
}

/**
 * Handles PATCH/POST /api/<module>/:id/archive and /api/<module>/:id/restore atomically.
 */
export async function handleEntityArchiveRestore(env, origin, user, { table, moduleKey, label = 'Data', id, mode }) {
  if (!user || String(user.role || '').trim().toLowerCase() !== 'superadmin') {
    return forbidden(origin, env?.ENVIRONMENT);
  }

  const hasArchiveCol = await hasTableColumn(env.DB, table, 'archived_at');
  if (!hasArchiveCol) {
    return error(`Kolom archived_at belum tersedia pada tabel ${table}`, 400, origin, env?.ENVIRONMENT);
  }

  const hasDeletedCol = await hasTableColumn(env.DB, table, 'deleted_at');
  const hasRowVerCol = await hasTableColumn(env.DB, table, 'row_version');
  const hasUpdatedCol = await hasTableColumn(env.DB, table, 'updated_at');

  const whereDel = hasDeletedCol ? ' AND deleted_at IS NULL' : '';
  const existing = await env.DB.prepare(`SELECT * FROM ${table} WHERE id = ?${whereDel}`).bind(Number(id)).first();
  if (!existing) return notFound(origin, env?.ENVIRONMENT);

  if (mode === 'archive') {
    if (existing.archived_at) {
      return error(`${label} sudah dalam status arsip`, 400, origin, env?.ENVIRONMENT);
    }
    if (table === 'branches' && isManagementBranch(existing)) {
      return error('Cabang 000. Management tidak dapat diarsipkan', 400, origin, env?.ENVIRONMENT);
    }
  } else {
    if (!existing.archived_at) {
      return error(`${label} tidak dalam status arsip`, 400, origin, env?.ENVIRONMENT);
    }
  }

  const nowIso = new Date().toISOString();
  const prevVersion = Number(existing.row_version ?? 0);
  const newVersion = prevVersion + 1;

  const setParts = [
    mode === 'archive' ? "archived_at = datetime('now')" : 'archived_at = NULL'
  ];
  const bindParams = [];
  if (hasRowVerCol) {
    setParts.push('row_version = ?');
    bindParams.push(newVersion);
  }
  if (hasUpdatedCol) {
    setParts.push("updated_at = datetime('now')");
  }

  let whereOcc = 'WHERE id = ?';
  bindParams.push(Number(id));
  if (hasRowVerCol && existing.row_version !== null && existing.row_version !== undefined) {
    whereOcc += ' AND row_version = ?';
    bindParams.push(existing.row_version);
  }

  const updateStmt = env.DB.prepare(`UPDATE ${table} SET ${setParts.join(', ')} ${whereOcc}`).bind(...bindParams);
  const auditStmt = buildAuditStatement(
    env,
    user,
    'UPDATE',
    moduleKey,
    Number(id),
    { archived_at: existing.archived_at || null, ...(hasRowVerCol ? { row_version: existing.row_version } : {}) },
    { archived_at: mode === 'archive' ? nowIso : null, ...(hasRowVerCol ? { row_version: newVersion } : {}) },
    { lifecycle_action: mode === 'archive' ? 'ARCHIVE' : 'RESTORE', table }
  );

  const stmts = auditStmt ? [updateStmt, auditStmt] : [updateStmt];
  const batchRes = await env.DB.batch(stmts);
  const changes = batchRes?.[0]?.meta?.changes ?? batchRes?.[0]?.changes ?? 0;

  if (changes === 0) {
    return error('Konflik OCC: Data telah diubah oleh pengguna lain. Silakan muat ulang.', 409, origin, env?.ENVIRONMENT);
  }

  return ok(
    {
      message: mode === 'archive' ? `${label} berhasil diarsipkan` : `${label} berhasil dipulihkan dari arsip`,
      id: Number(id),
      archived: mode === 'archive',
      ...(hasRowVerCol ? { row_version: newVersion } : {})
    },
    200,
    origin,
    env?.ENVIRONMENT,
    { 'X-Audit-Batched': 'true' }
  );
}
