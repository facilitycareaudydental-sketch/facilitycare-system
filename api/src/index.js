/**
 * FM Operations Management System - Cloudflare Worker API
 * Routes: /api/auth, /api/users, /api/branches, /api/employees, /api/contracts,
 *         /api/schedule, /api/issues, /api/one-on-one, /api/training,
 *         /api/relievers, /api/reports/*, /api/sop, /api/checklist,
 *         /api/forms, /api/pic, /api/dashboard, /api/sp, /api/mutasi
 */

import { handleAuth } from './routes/auth.js';
import { handleUsers } from './routes/users.js';
import { handleBranches } from './routes/branches.js';
import { handleEmployees } from './routes/employees.js';
import { handleContracts } from './routes/contracts.js';
import { handleSchedule } from './routes/schedule.js';
import { handleIssues } from './routes/issues.js';
import { handleOneOnOne } from './routes/one_on_one.js';
import { handleTraining } from './routes/training.js';
import { handleRelievers } from './routes/relievers.js';
import { handleReports } from './routes/reports.js';
import { handleMisc } from './routes/misc.js';
import { handleDashboard } from './routes/dashboard.js';
import { handleImport } from './routes/import.js';
import { handleSP } from './routes/sp.js';
import { handleMutasi } from './routes/mutasi.js';
import { handleOvertime } from './routes/overtime.js';
import { handleHygieneStandards } from './routes/hygiene.js';
import { handleOSMigrations } from './routes/os_migrations.js';
import { handleDAC } from './routes/dac.js';
import { syncGoogleSheets } from './utils/google_sync.js';
import { receiveWebhook, processOutbox, buildBulkOutboxStatements } from './utils/sync_engine.js';
import { resolveSyncMode } from './utils/sync_mode.js';
import { enforceMaintenanceFreeze, resolveMaintenanceFreezeState } from './utils/maintenance.js';
import { mapDBToPayload } from './utils/sync_mapper.js';
import { handleMonitoring } from './routes/monitoring.js';
import { resolveManagementBranchIds } from './utils/branchMaster.js';
import { options, error, ok, forbidden, unauthorized, resolveAllowedOrigin, isAllowedOrigin, corsHeaders } from './utils/response.js';
import { authenticate, hasPermission, verifyToken, fetchAuthUserWithSession, isSessionRecordValid, VIEWER_ROLES, maskViewerData } from './utils/auth.js';
import { sanitizeAuditData } from './utils/audit.js';
import { hasTable, hasTableColumn } from './utils/schemaCompat.js';
import { ARCHIVABLE_ROUTE_MAP, handleEntityArchiveRestore, canViewArchivedRecords } from './utils/archive.js';
import { checkTier3GeneralApiRateLimit } from './utils/rate_limit.js';

export const MAX_BULK_IDS = 100;

const AUDIT_TABLE_ALLOWLIST = Object.freeze({
  users: 'users',
  branches: 'branches',
  employees: 'employees',
  contracts: 'contracts',
  schedule: 'activity_schedule',
  issues: 'issues',
  'one-on-one': 'one_on_one',
  one_on_one: 'one_on_one',
  training: 'training',
  relievers: 'relievers',
  sp: 'sp_data',
  mutasi: 'mutasi_data',
  overtime: 'overtime_records',
  'hygiene-standards': 'master_hygiene_standards',
  hygiene_standards: 'master_hygiene_standards',
  'os-migrations': 'os_migrations',
  os_migrations: 'os_migrations',
  dac: 'master_dac_standards',
  sop: 'sop',
  checklist: 'master_checklist',
  forms: 'master_forms',
  pic: 'pic_list',
  options: 'validation_options',
  'reports/inspection': 'inspection_reports',
  'reports/cleaning': 'cleaning_reports',
  'reports/fogging': 'fogging_reports',
  'reports/basecamp': 'basecamp_reports',
  'reports/supply': 'supply_requests'
});

function normalizeApiPathname(pathname) {
  if (!pathname || typeof pathname !== 'string') return '/';
  const collapsed = pathname.replace(/\/{2,}/g, '/');
  if (collapsed.length > 1 && collapsed.endsWith('/')) {
    return collapsed.replace(/\/+$/, '');
  }
  return collapsed;
}

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    const normalizedPath = normalizeApiPathname(url.pathname);
    if (url.pathname !== normalizedPath) {
      url.pathname = normalizedPath;
      const normalizedUrlStr = url.toString();
      request = new Proxy(request, {
        get(target, prop) {
          if (prop === 'url') return normalizedUrlStr;
          const val = Reflect.get(target, prop);
          return typeof val === 'function' ? val.bind(target) : val;
        }
      });
    }

    const origin = resolveAllowedOrigin(request.headers.get('Origin'), env.CORS_ORIGIN, env?.ENVIRONMENT);

    // Handle CORS preflight
    if (request.method === 'OPTIONS') {
      const reqOrigin = request.headers.get('Origin');
      if (reqOrigin && !isAllowedOrigin(reqOrigin, env.CORS_ORIGIN, env?.ENVIRONMENT)) {
        return error('Origin not allowed by CORS policy', 403, origin, env?.ENVIRONMENT);
      }
      return options(origin, env?.ENVIRONMENT);
    }

    // Only handle /api/* routes
    if (!url.pathname.startsWith('/api/')) {
      return new Response('FM Operations API', {
        status: 200,
        headers: { 'Content-Type': 'text/plain' }
      });
    }

    try {
      const path = url.pathname;
      request._fcmsAuthScope = true;

      // ── JWT Scope Barrier & Restricted Token Gatekeeper (Fail-Closed) ───────
      // Restricted tokens (e.g. mustChangePassword: true) or accounts requiring password change
      // MUST NOT access any operational or administrative resources.
      const authHeader = request.headers.get('Authorization');
      if (authHeader && authHeader.startsWith('Bearer ')) {
        const token = authHeader.slice(7);
        if (env?.JWT_SECRET) {
          const payload = await verifyToken(token, env.JWT_SECRET);
          request._fcmsVerifiedPayload = payload;
          const sessionId = (payload && typeof payload.sid === 'string') ? payload.sid.trim() : '';
          if (payload && sessionId) {
            const isRestrictedClaim = payload.mustChangePassword === true || payload.scope === 'restricted:change_password';
            let dbMustChange = false;
            let sessionValid = false;
            if (env.DB && payload.sub) {
              try {
                const { user: dbUser, hasAssignedModulesColumn: hasAssignedCol } =
                  await fetchAuthUserWithSession(env, payload.sub, sessionId);
                sessionValid = !!(
                  dbUser &&
                  Number(dbUser.is_active) === 1 &&
                  isSessionRecordValid(dbUser, sessionId)
                );
                request._fcmsPreloadedDbUser = sessionValid ? dbUser : null;
                request._fcmsHasAssignedModulesCol = hasAssignedCol;
                if (sessionValid) {
                  const rawMcp = dbUser.must_change_password;
                  const isExplicitlyUnrestricted =
                    rawMcp === 0 ||
                    rawMcp === '0' ||
                    rawMcp === false ||
                    (rawMcp === undefined && !('must_change_password' in dbUser));
                  if (!isExplicitlyUnrestricted) {
                    dbMustChange = true;
                  }
                }
              } catch (dbErr) {
                console.error('Scope barrier DB error:', dbErr);
                return error('Gagal memverifikasi status keamanan akun', 503, origin);
              }
            }

            if (sessionValid && (isRestrictedClaim || dbMustChange)) {
              const isAllowedRestrictedEndpoint =
                (path === '/api/auth/change-password' && request.method === 'POST') ||
                (path === '/api/auth/me' && request.method === 'GET') ||
                (path === '/api/auth/logout' && request.method === 'POST');
              if (!isAllowedRestrictedEndpoint) {
                return error('Forbidden - Password change required before accessing other resources', 403, origin);
              }
            }
          }
        }
      }

      // ── Fail-Closed Maintenance Freeze Gate (All HTTP Routes & Mutations) ──
      const maintenanceResponse = await enforceMaintenanceFreeze(request, env, origin);
      if (maintenanceResponse) {
        return maintenanceResponse;
      }

      // Generic Bulk Delete Handler (Superadmin Only)
      if (request.method === 'DELETE' && path.endsWith('/bulk')) {
        const user = await authenticate(request, env);
        if (!user) return unauthorized(origin);
        
        const moduleMap = {
          '/api/branches/bulk': {
            table: 'branches',
            primaryKey: 'id',
            strategy: 'soft',
            softDeleteCol: 'deleted_at',
            updatedAtCol: 'updated_at',
            isActiveCol: 'is_active',
            trackSyncVersion: true,
            outboxEntity: 'Data Cabang',
            perm: 'branches',
            maxLimit: MAX_BULK_IDS
          },
          '/api/employees/bulk': { table: 'employees', primaryKey: 'id', strategy: 'soft', softDeleteCol: 'deleted_at', updatedAtCol: 'updated_at', perm: 'employees', maxLimit: MAX_BULK_IDS, managementScope: true },
          '/api/issues/bulk': {
            table: 'issues',
            primaryKey: 'id',
            strategy: 'soft',
            softDeleteCol: 'deleted_at',
            updatedAtCol: 'updated_at',
            trackSyncVersion: true,
            outboxEntity: 'Permasalahan',
            perm: 'issues',
            maxLimit: MAX_BULK_IDS
          },
          '/api/schedule/bulk': {
            table: 'activity_schedule',
            primaryKey: 'id',
            strategy: 'soft',
            softDeleteCol: 'deleted_at',
            updatedAtCol: 'updated_at',
            trackSyncVersion: true,
            outboxEntity: 'Time Line',
            calModule: 'schedule',
            perm: 'schedule',
            maxLimit: MAX_BULK_IDS
          },
          '/api/training/bulk': { table: 'training', primaryKey: 'id', strategy: 'hard', calModule: 'training', perm: 'training', maxLimit: MAX_BULK_IDS },
          '/api/sp/bulk': { table: 'sp_data', primaryKey: 'id', strategy: 'hard', perm: 'sp', maxLimit: 100 },
          '/api/mutasi/bulk': { table: 'mutasi_data', primaryKey: 'id', strategy: 'hard', perm: 'mutasi', maxLimit: 100 },
          '/api/sop/bulk': { table: 'sop', primaryKey: 'id', strategy: 'soft', softDeleteCol: 'deleted_at', updatedAtCol: 'updated_at', trackSyncVersion: true, outboxEntity: 'Master SOP', perm: 'sop', maxLimit: 100 },
          '/api/checklist/bulk': { table: 'master_checklist', primaryKey: 'id', strategy: 'soft', softDeleteCol: 'deleted_at', updatedAtCol: 'updated_at', trackSyncVersion: true, outboxEntity: 'Master Checklist', perm: 'checklist', maxLimit: 100 },
          '/api/pic/bulk': { table: 'pic_list', primaryKey: 'id', strategy: 'hard', perm: 'pic', maxLimit: 100 },
          '/api/users/bulk': { table: 'users', primaryKey: 'id', strategy: 'hard', customGuard: 'no_superadmin', perm: 'users', maxLimit: 100 },
          '/api/contracts/bulk': { table: 'contracts', primaryKey: 'id', strategy: 'soft', softDeleteCol: 'deleted_at', updatedAtCol: 'updated_at', calModule: 'contracts', perm: 'contracts', maxLimit: 100, managementScope: true },
          '/api/reports/cleaning/bulk': { table: 'cleaning_reports', primaryKey: 'id', strategy: 'soft', softDeleteCol: 'deleted_at', updatedAtCol: 'updated_at', trackSyncVersion: true, outboxEntity: 'GC-DC', perm: 'reports', maxLimit: 100 },
          '/api/reports/fogging/bulk': { table: 'fogging_reports', primaryKey: 'id', strategy: 'soft', softDeleteCol: 'deleted_at', updatedAtCol: 'updated_at', trackSyncVersion: true, outboxEntity: 'Fogging', perm: 'reports', maxLimit: 100 },
          '/api/reports/basecamp/bulk': { table: 'basecamp_reports', primaryKey: 'id', strategy: 'soft', softDeleteCol: 'deleted_at', updatedAtCol: 'updated_at', trackSyncVersion: true, outboxEntity: 'Basecamp', perm: 'reports', maxLimit: 100 },
          '/api/reports/inspection/bulk': { table: 'inspection_reports', primaryKey: 'id', strategy: 'soft', softDeleteCol: 'deleted_at', updatedAtCol: 'updated_at', trackSyncVersion: true, outboxEntity: 'Hygiene', perm: 'reports', maxLimit: 100 },
          '/api/overtime/bulk': { table: 'overtime_records', primaryKey: 'id', strategy: 'hard', perm: 'overtime', maxLimit: 100 },
          '/api/os-migrations/bulk': { table: 'os_migrations', primaryKey: 'id', strategy: 'hard', perm: 'os_migrations', maxLimit: 100 },
          '/api/hygiene-standards/bulk': { table: 'master_hygiene_standards', primaryKey: 'id', strategy: 'hard', perm: 'hygiene_standards', maxLimit: 100 },
          '/api/dac/bulk': { table: 'master_dac_standards', primaryKey: 'id', strategy: 'hard', perm: 'dac', maxLimit: 100 }
        };

        const rawConfig = moduleMap[path];
        if (!rawConfig) return error('Invalid bulk delete module', 400, origin);
        const hasSoftCol = rawConfig.strategy === 'soft' && rawConfig.softDeleteCol
          ? await hasTableColumn(env.DB, rawConfig.table, rawConfig.softDeleteCol)
          : false;
        const config = (rawConfig.strategy === 'soft' && !hasSoftCol)
          ? { ...rawConfig, strategy: 'hard', outboxEntity: undefined }
          : rawConfig;
        if (!hasPermission(user, config.perm, 'delete')) return forbidden(origin);

        let body;
        try {
          body = await request.json();
        } catch {
          return error('Invalid JSON', 400, origin);
        }

        try {
          const { ids } = body || {};
          if (!Array.isArray(ids) || ids.length === 0) return error('No IDs provided', 400, origin);
          
          const validIds = ids
            .filter(id => (typeof id === 'number' || (typeof id === 'string' && id.trim() !== '')) && Number.isInteger(Number(id)) && Number(id) > 0)
            .map(Number);
          if (validIds.length === 0) return error('No valid numeric IDs provided', 400, origin);

          let targetIds = Array.from(new Set(validIds));
          if (config.customGuard === 'no_superadmin') {
            targetIds = targetIds.filter(id => id !== Number(user.id));
            if (targetIds.length === 0) return error('Tidak dapat menghapus akun sendiri', 400, origin);

            // Ensure no active Superadmin is mass-deleted via bulk delete and at least 1 active Superadmin remains (fail-closed)
            try {
              const stmt = env.DB.prepare("SELECT id, role, is_active FROM users WHERE LOWER(role) = 'superadmin' AND is_active = 1");
              if (typeof stmt.all === 'function') {
                const activeRes = await stmt.all();
                if (activeRes && Array.isArray(activeRes.results)) {
                  const targetSet = new Set(targetIds);
                  const targetedSuperadmins = activeRes.results.filter(
                    u => String(u.role || '').toLowerCase() === 'superadmin' &&
                         targetSet.has(Number(u.id))
                  );
                  if (targetedSuperadmins.length > 0) {
                    return error('Akun Superadmin tidak dapat dihapus melalui fitur Bulk Delete', 400, origin);
                  }
                }
              }
            } catch (saErr) {
              console.error('Active superadmin check error in bulk delete:', saErr);
              return error('Gagal memverifikasi proteksi akun Superadmin', 503, origin);
            }
          }
          
          const maxBulkLimit = config.maxLimit || MAX_BULK_IDS;
          if (targetIds.length > maxBulkLimit) {
            return error(`Maksimum ${maxBulkLimit} item dapat dihapus dalam satu permintaan bulk`, 400, origin);
          }

          // 1. Fetch active records before mutation (full row needed for per-record audit snapshot & outbox payload)
          const selectCols = '*';
          const SELECT_CHUNK_SIZE = 50; // Keep bind params <= 50 (well below D1 100-param limit)
          let activeRows = [];
          try {
            for (let i = 0; i < targetIds.length; i += SELECT_CHUNK_SIZE) {
              const idChunk = targetIds.slice(i, i + SELECT_CHUNK_SIZE);
              const chunkPlaceholders = idChunk.map(() => '?').join(',');
              let chunkSql;
              if (config.strategy === 'soft') {
                chunkSql = `SELECT ${selectCols} FROM ${config.table} WHERE ${config.primaryKey} IN (${chunkPlaceholders}) AND ${config.softDeleteCol} IS NULL`;
              } else if (config.customGuard === 'no_superadmin') {
                chunkSql = `SELECT ${selectCols} FROM ${config.table} WHERE ${config.primaryKey} IN (${chunkPlaceholders}) AND LOWER(role) != 'superadmin'`;
              } else {
                chunkSql = `SELECT ${selectCols} FROM ${config.table} WHERE ${config.primaryKey} IN (${chunkPlaceholders})`;
              }
              const checkRes = await env.DB.prepare(chunkSql).bind(...idChunk).all();
              if (Array.isArray(checkRes?.results)) {
                activeRows.push(...checkRes.results);
              }
            }
          } catch (checkErr) {
            console.error('Bulk check error:', checkErr);
            return error('Gagal memeriksa data untuk penghapusan massal', 500, origin);
          }

          const targetIdSet = new Set(targetIds);
          const seenActiveIds = new Set();
          activeRows = activeRows.filter(r => {
            const rid = Number(r[config.primaryKey]);
            if (!targetIdSet.has(rid) || seenActiveIds.has(rid)) return false;
            if (config.strategy === 'soft' && r[config.softDeleteCol] !== null && r[config.softDeleteCol] !== undefined) {
              return false;
            }
            seenActiveIds.add(rid);
            return true;
          });

          if (config.managementScope && activeRows.length > 0) {
            let mgmtBranchIds;
            try {
              mgmtBranchIds = await resolveManagementBranchIds(env.DB);
            } catch (mgmtErr) {
              return error(mgmtErr.message, 503, origin, env?.ENVIRONMENT);
            }
            const OFFICIAL_DIVISIONS = new Set(['FACILITY CARE', 'SECURITY']);
            for (const row of activeRows) {
              const rowDiv = String(row.division || '').trim().toUpperCase();
              const rowBranchId = Number(row.branch_id);
              if (!mgmtBranchIds.has(rowBranchId) || !OFFICIAL_DIVISIONS.has(rowDiv)) {
                return error(
                  `Data ID ${row[config.primaryKey]} berada di luar cakupan Management (Cabang 000. Management) atau divisi resmi; penghapusan massal ditolak`,
                  400,
                  origin
                );
              }
              if (config.table === 'contracts' && row.employee_id) {
                const emp = await env.DB.prepare(
                  'SELECT id, branch_id, division FROM employees WHERE id = ? AND deleted_at IS NULL'
                ).bind(row.employee_id).first();
                if (emp) {
                  const empDiv = String(emp.division || '').trim().toUpperCase();
                  if (!mgmtBranchIds.has(Number(emp.branch_id)) || !OFFICIAL_DIVISIONS.has(empDiv)) {
                    return error(
                      `Kontrak ID ${row[config.primaryKey]} terhubung dengan karyawan di luar cakupan Management (000. Management); penghapusan massal ditolak`,
                      400,
                      origin
                    );
                  }
                }
              }
            }
          }

          const activeIds = activeRows.map(r => Number(r[config.primaryKey]));

          // If no records are active/found, return early without false SUCCESS audit log or outbox events
          if (activeIds.length === 0) {
            return ok({ 
              message: 'Tidak ada data yang dihapus (data tidak ditemukan atau sudah dihapus)', 
              count: 0, 
              requested: targetIds.length 
            }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
          }

          // 2. Prepare single atomic batch for active IDs based on module deletion strategy
          // Chunk primary UPDATE/DELETE into <= 50 IDs per statement so bound params stay <= 52 (well below D1's 100-variable limit)
          const {
            buildSoftDeleteRecordAuditStatement,
            buildBulkHardDeleteAuditStatements
          } = await import('./utils/audit.js');
          const batchDeletedAt = new Date().toISOString();
          const bulkCorrelationId = crypto.randomUUID();
          const MUTATION_CHUNK_SIZE = 50;
          const batchStmts = [];
          const mutationStmtIndices = [];
          const batchAuditMeta = {
            batch_id: bulkCorrelationId,
            batch_size: activeIds.length,
            count: activeIds.length,
            ids: activeIds,
            requested_count: targetIds.length,
            strategy: config.strategy
          };

          for (let i = 0; i < activeRows.length; i += MUTATION_CHUNK_SIZE) {
            const rowChunk = activeRows.slice(i, i + MUTATION_CHUNK_SIZE);
            const idChunk = rowChunk.map(r => Number(r[config.primaryKey]));
            const chunkPlaceholders = idChunk.map(() => '?').join(',');
            let deleteSql;
            let deleteBinds;

            if (config.strategy === 'soft') {
              const setClauses = [
                `${config.softDeleteCol} = ?`,
                `${config.updatedAtCol} = ?`
              ];
              if (config.isActiveCol) {
                setClauses.push(`${config.isActiveCol} = 0`);
              }
              if (config.trackSyncVersion) {
                setClauses.push(`row_version = COALESCE(row_version, 0) + 1`);
                setClauses.push(`last_sync_source = 'FCMS'`);
              }
              deleteSql = `UPDATE ${config.table} SET ${setClauses.join(', ')} WHERE ${config.primaryKey} IN (${chunkPlaceholders}) AND ${config.softDeleteCol} IS NULL`;
              deleteBinds = [batchDeletedAt, batchDeletedAt, ...idChunk];
            } else if (config.customGuard === 'no_superadmin') {
              deleteSql = `DELETE FROM ${config.table} WHERE ${config.primaryKey} IN (${chunkPlaceholders}) AND LOWER(role) != 'superadmin'`;
              deleteBinds = [...idChunk];
            } else {
              deleteSql = `DELETE FROM ${config.table} WHERE ${config.primaryKey} IN (${chunkPlaceholders})`;
              deleteBinds = [...idChunk];
            }

            mutationStmtIndices.push(batchStmts.length);
            batchStmts.push(env.DB.prepare(deleteSql).bind(...deleteBinds));

            // Enqueue chunked multi-row outbox DELETE events immediately after each mutation chunk so (SELECT changes()) > 0 reflects mutation count
            if (config.outboxEntity) {
              const outboxEntries = rowChunk.map(row => {
                const nextRowVersion = (row.row_version !== null && row.row_version !== undefined && !isNaN(Number(row.row_version)))
                  ? Number(row.row_version) + 1
                  : 1;
                const deletedRecord = {
                  ...row,
                  deleted_at: batchDeletedAt,
                  last_sync_source: 'FCMS',
                  row_version: nextRowVersion,
                  ...(config.isActiveCol ? { [config.isActiveCol]: 0 } : {}),
                  ...(config.table === 'branches'
                    ? { branch_code: row.code, branch_name: row.name || row.full_name }
                    : {})
                };
                return {
                  entityName: config.outboxEntity,
                  entityId: row[config.primaryKey],
                  action: 'DELETE',
                  payload: mapDBToPayload(config.outboxEntity, deletedRecord),
                  correlationId: bulkCorrelationId
                };
              });
              batchStmts.push(
                ...buildBulkOutboxStatements(env, outboxEntries, { chunkSize: 15, requireChanges: true })
              );
            }

            // For hard deletes without outbox, insert chunked per-record audit statements right after each DELETE chunk
            if (config.strategy === 'hard') {
              batchStmts.push(
                ...buildBulkHardDeleteAuditStatements(env, user, config.perm, rowChunk, config.primaryKey, batchAuditMeta)
              );
            }
          }

          // Delete associated calendar events only for modules where single delete removes them,
          // chunked safely within D1's 100 bound-parameter limit and scoped to mutated records.
          if (config.calModule && (await hasTable(env.DB, 'calendar_events'))) {
            try {
              const CAL_CHUNK_SIZE = 50; // 1 + 50 = 51 bind params <= 100 D1 limit
              for (let i = 0; i < activeIds.length; i += CAL_CHUNK_SIZE) {
                const calChunk = activeIds.slice(i, i + CAL_CHUNK_SIZE);
                const calPlaceholders = calChunk.map(() => '?').join(',');
                if (config.strategy === 'soft') {
                  batchStmts.push(
                    env.DB.prepare(
                      `DELETE FROM calendar_events WHERE module = ? AND reference_id IN (${calPlaceholders}) AND EXISTS (SELECT 1 FROM ${config.table} t WHERE t.${config.primaryKey} = calendar_events.reference_id AND t.${config.softDeleteCol} IS NOT NULL)`
                    ).bind(config.calModule, ...calChunk)
                  );
                } else {
                  batchStmts.push(
                    env.DB.prepare(
                      `DELETE FROM calendar_events WHERE module = ? AND reference_id IN (${calPlaceholders}) AND (SELECT changes()) > 0`
                    ).bind(config.calModule, ...calChunk)
                  );
                }
              }
            } catch (calError) {
              console.warn('Could not delete calendar events, table might not exist:', calError.message);
            }
          }

          // Revoke any active sessions belonging to hard-deleted users atomically in the same batch
          if (config.table === 'users') {
            const SESS_CHUNK_SIZE = 50;
            for (let i = 0; i < activeIds.length; i += SESS_CHUNK_SIZE) {
              const sessChunk = activeIds.slice(i, i + SESS_CHUNK_SIZE);
              const sessPlaceholders = sessChunk.map(() => '?').join(',');
              batchStmts.push(
                env.DB.prepare(
                  `DELETE FROM sessions WHERE user_id IN (${sessPlaceholders})`
                ).bind(...sessChunk)
              );
            }
          }

          // For soft deletes, emit one per-record DELETE audit statement guarded by EXISTS (deleted_at = batchDeletedAt)
          // so each actually modified record captures its exact pre-deletion snapshot and batch metadata with zero TOCTOU phantoms.
          if (config.strategy === 'soft') {
            for (const row of activeRows) {
              batchStmts.push(
                buildSoftDeleteRecordAuditStatement(
                  env,
                  user,
                  config.perm,
                  config.table,
                  config.primaryKey,
                  config.softDeleteCol,
                  batchDeletedAt,
                  row,
                  batchAuditMeta
                )
              );
            }
          }

          // Atomic single D1 execution: mutation + audit log + outbox + calendar/session cleanup in ONE transaction
          const batchRes = await env.DB.batch(batchStmts);

          // Fail-closed verification of changes metadata across all primary mutation statement chunks
          let changedCount = 0;
          for (const stmtIdx of mutationStmtIndices) {
            let rawChanges = batchRes?.[stmtIdx]?.meta?.changes;
            if (rawChanges === undefined || rawChanges === null) {
              rawChanges = batchRes?.[stmtIdx]?.changes;
            }
            if (typeof rawChanges !== 'number' || isNaN(rawChanges) || rawChanges < 0) {
              throw new Error('DATABASE_MUTATION_UNVERIFIED: Mutation changes metadata missing or invalid from D1 execution');
            }
            changedCount += rawChanges;
          }

          if (changedCount === 0) {
            return ok({ 
              message: 'Tidak ada data yang dihapus (data tidak ditemukan atau sudah dihapus)', 
              count: 0, 
              requested: targetIds.length 
            }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
          }

          return ok({ 
            message: `Berhasil menghapus ${changedCount} data`, 
            count: changedCount, 
            requested: targetIds.length 
          }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
        } catch (e) {
          console.error('Bulk delete error:', e);
          try {
            const { logAudit } = await import('./utils/audit.js');
            await logAudit(env, user, 'DELETE', config.perm, 'bulk', null, null, { error: e?.message || 'Bulk delete failed' }, 'FAILED');
          } catch {}
          return error('Gagal menghapus data secara massal', 500, origin);
        }
      }

      let oldData = null;
      let targetId = null;
      let moduleName = 'unknown';
      let action = request.method === 'POST' ? 'CREATE' : (request.method === 'PUT' || request.method === 'PATCH') ? 'UPDATE' : 'DELETE';

      // Capture ID and Old Data before modifying using strict table allowlist
      if (['PUT', 'PATCH', 'DELETE'].includes(request.method) && !path.startsWith('/api/auth')) {
         const parts = path.split('/').filter(Boolean); 
         if (parts.length >= 3) {
            if (parts[1] === 'reports' && parts.length >= 4) {
               moduleName = `reports/${parts[2]}`;
               targetId = parts[3];
            } else if (parts[1] === 'os-migrations' && parts[3] === 'status') {
               moduleName = parts[1];
               targetId = parts[2];
            } else {
               moduleName = parts[1];
               targetId = parts[parts.length - 1];
            }
            const safeTable = AUDIT_TABLE_ALLOWLIST[moduleName];
            if (safeTable && (!isNaN(targetId) || targetId.length > 5)) { 
               try {
                  const res = await env.DB.prepare(`SELECT * FROM ${safeTable} WHERE id = ?`).bind(targetId).first();
                  if (res) oldData = sanitizeAuditData(res);
               } catch (e) { /* table might not exist in mock DB, ignore */ }
            }
         }
      } else if (request.method === 'POST') {
         const parts = path.split('/').filter(Boolean);
         if (parts.length >= 2) {
            moduleName = (parts[1] === 'reports' && parts[2]) ? `reports/${parts[2]}` : parts[1];
         }
      }

      if (path === '/api/health' && request.method === 'GET') {
        const freezeState = await resolveMaintenanceFreezeState(env);
        const syncState = resolveSyncMode(env);
        return ok(
          {
            status: 'OK',
            service: 'fm-operations-api',
            environment: env?.ENVIRONMENT || 'unknown',
            syncMode: syncState.mode,
            maintenanceFreeze: freezeState.freezeActive
          },
          200,
          origin,
          env?.ENVIRONMENT,
          { 'Cache-Control': 'no-store' }
        );
      }

      // ── Tier 3 General API Rate Limit (120 req/min/user across all roles including superadmin) ──
      if (!path.startsWith('/api/auth') && !path.startsWith('/api/sync/webhook') && env?.ENFORCE_TIER3_RATE_LIMIT === true) {
        const rlUser = await authenticate(request, env);
        if (rlUser) {
          const tier3 = await checkTier3GeneralApiRateLimit(rlUser, env, origin);
          if (!tier3.allowed) return tier3.response;
        }
      }

      // ── Global Single Archive Read Guard (?archived=true) & Unified Archive/Restore Dispatch ──
      if (request.method === 'GET' && url.searchParams.get('archived') === 'true') {
        const archUser = await authenticate(request, env);
        if (!archUser) return unauthorized(origin);
        if (!canViewArchivedRecords(archUser)) return forbidden(origin);
      }

      const archiveRouteMatch = path.match(/^\/api\/(employees|contracts|branches|issues|schedule|sop|checklist|reports\/(?:inspection|cleaning|fogging|basecamp))\/(\d+)\/(archive|restore)$/);
      if (archiveRouteMatch && (request.method === 'PATCH' || request.method === 'POST')) {
        const archUser = await authenticate(request, env);
        if (!archUser) return unauthorized(origin);
        const routeSpec = ARCHIVABLE_ROUTE_MAP[archiveRouteMatch[1]];
        if (!routeSpec) return error('Not found', 404, origin);
        return handleEntityArchiveRestore(env, origin, archUser, {
          ...routeSpec,
          id: archiveRouteMatch[2],
          mode: archiveRouteMatch[3]
        });
      }

      let response;
      if (path.startsWith('/api/auth')) response = await handleAuth(request, env, origin);
      else if (path.startsWith('/api/users')) response = await handleUsers(request, env, origin);
      else if (path.startsWith('/api/branches')) response = await handleBranches(request, env, origin);
      else if (path.startsWith('/api/employees')) response = await handleEmployees(request, env, origin);
      else if (path.startsWith('/api/contracts')) response = await handleContracts(request, env, origin);
      else if (path.startsWith('/api/schedule')) response = await handleSchedule(request, env, origin);
      else if (path.startsWith('/api/issues')) response = await handleIssues(request, env, origin);
      else if (path.startsWith('/api/training')) response = await handleTraining(request, env, origin);
      else if (path.startsWith('/api/reports')) response = await handleReports(request, env, origin);
      else if (path.startsWith('/api/dashboard')) response = await handleDashboard(request, env, origin);
      else if (path.startsWith('/api/import')) response = await handleImport(request, env, origin);
      else if (path.startsWith('/api/sp')) response = await handleSP(request, env, origin);
      else if (path.startsWith('/api/mutasi')) response = await handleMutasi(request, env, origin);
      else if (path.startsWith('/api/overtime')) response = await handleOvertime(request, env, origin);
      else if (path.startsWith('/api/hygiene-standards')) response = await handleHygieneStandards(request, env, origin);
      else if (path.startsWith('/api/os-migrations')) response = await handleOSMigrations(request, env, origin);
      else if (path.startsWith('/api/dac')) response = await handleDAC(request, env, origin);
      else if (path.startsWith('/api/audit-logs')) response = await (await import('./routes/audit_logs.js')).handleAuditLogs(request, env, origin);
      else if (path.startsWith('/api/sop') || path.startsWith('/api/checklist') || path.startsWith('/api/pic') || path.startsWith('/api/options')) response = await handleMisc(request, env, origin);
      
      // Post-processing for Audit Logging (Handles both Success and Authorized Failed Mutations)
      if (['POST', 'PUT', 'PATCH', 'DELETE'].includes(request.method) && !path.startsWith('/api/auth') && response) {
         try {
           const user = await authenticate(request, env);
           if (user) {
             const isSuccess = response.status >= 200 && response.status < 300;
             let newData = null;
             if (isSuccess && request.method !== 'DELETE') {
               try {
                  const cloned = response.clone();
                  const resBody = await cloned.json();
                  if (resBody && resBody.data) newData = resBody.data;
               } catch(e) {}
             }
             if (request.method === 'POST' && newData && newData.id) targetId = newData.id;

             let failureDetails = null;
             if (!isSuccess) {
               try {
                 const clonedErr = response.clone();
                 const errBody = await clonedErr.json();
                 failureDetails = { statusCode: response.status, error: errBody?.error || 'Operation failed' };
               } catch {
                 failureDetails = { statusCode: response.status, error: 'HTTP error response' };
               }
             }

             const isAlreadyBatched = response.headers?.get?.('X-Audit-Batched') === 'true';
             const { logAudit } = await import('./utils/audit.js');
             if (isSuccess) {
               if (!isAlreadyBatched) {
                 await logAudit(env, user, action, moduleName, targetId, oldData, newData, null, 'SUCCESS');
               }
             } else if (!isAlreadyBatched && response.status >= 400 && response.status !== 401 && response.status !== 403) {
               // Authorized business mutation failure (e.g. 400 Bad Request, 404 Not Found, 409 OCC Conflict, 500 DB Error)
               await logAudit(env, user, action, moduleName, targetId, oldData, null, failureDetails, 'FAILED');
             }
           }
         } catch(e) {
           console.error('Audit Log Error:', e);
         }
      }

      if (response) {
        // Central response-boundary Sensitive Data Masking for Viewer roles on JSON GET responses
        const activeUser = request._fcmsAuthResult;
        if (
          request.method === 'GET' &&
          activeUser &&
          VIEWER_ROLES.has(String(activeUser.role || '').trim().toLowerCase()) &&
          response.status === 200 &&
          (response.headers?.get?.('Content-Type') || '').includes('application/json')
        ) {
          try {
            const bodyObj = await response.clone().json();
            if (bodyObj && typeof bodyObj === 'object' && 'data' in bodyObj) {
              bodyObj.data = maskViewerData(bodyObj.data, activeUser.role);
              return new Response(JSON.stringify(bodyObj), {
                status: response.status,
                headers: response.headers
              });
            }
          } catch {}
        }
        return response;
      }
      
      if (path === '/api/sync/google-sheets' && request.method === 'POST') {
        const user = await authenticate(request, env);
        if (!user) return unauthorized(origin);
        if (user.role !== 'superadmin') return forbidden(origin);
        const result = await syncGoogleSheets(env, fetch, user);
        return new Response(JSON.stringify(result), {
          status: result.success ? 200 : 500,
          headers: {
            'Content-Type': 'application/json',
            ...(result.success && !result.dryRun ? { 'X-Audit-Batched': 'true' } : {}),
            ...corsHeaders(origin)
          }
        });
      }

      if (path === '/api/sync/webhook') {
        return await receiveWebhook(request, env, origin);
      }

      if (path.startsWith('/api/sync/health') || path.startsWith('/api/sync/status') || path.startsWith('/api/sync/queue') || path.startsWith('/api/sync/metrics') || path.startsWith('/api/sync/errors') || path.startsWith('/api/sync/performance') || path.startsWith('/api/sync/actions/')) {
        return handleMonitoring(request, env, origin);
      }

      return error('API endpoint not found', 404, origin);
    } catch (err) {
      console.error('Worker error:', err);
      return error('Internal server error', 500, origin);
    } finally {
      delete request._fcmsAuthScope;
      delete request._fcmsAuthResult;
      delete request._fcmsVerifiedPayload;
      delete request._fcmsPreloadedDbUser;
      delete request._fcmsHasAssignedModulesCol;
    }
  },

  async scheduled(event, env, ctx) {
    const freezeState = await resolveMaintenanceFreezeState(env);
    if (freezeState.freezeActive) {
      console.log(`[MAINTENANCE_FREEZE] Scheduled cron halted (${freezeState.source}).`);
      return;
    }
    const syncState = resolveSyncMode(env);
    if (!syncState.isLive) {
      console.log(`[${syncState.mode}] Scheduled cron skipped external sync processing (${syncState.reason || 'DRY_RUN'}).`);
      return;
    }
    ctx.waitUntil(processOutbox(env)); // New Phase 1 Bidirectional Sweeper
  }
};
