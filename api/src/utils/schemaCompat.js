/**
 * Lightweight runtime schema compatibility probe with WeakMap caching per DB handle.
 * Used to safely adapt queries when optional additive columns (e.g. archived_at, deleted_at)
 * or optional support tables (e.g. calendar_events, feature_flags, circuit_breaker, sync_metrics)
 * are not present in the 30-table Production D1 baseline, without creating unnecessary tables.
 */
const ALLOWED_TABLES = new Set([
  'activity_schedule',
  'audit_logs',
  'auth_cooldowns',
  'basecamp_reports',
  'branches',
  'calendar_events',
  'circuit_breaker',
  'cleaning_reports',
  'contracts',
  'employees',
  'feature_flags',
  'fogging_reports',
  'import_jobs',
  'inspection_reports',
  'issues',
  'master_checklist',
  'master_dac_standards',
  'master_forms',
  'master_hygiene_standards',
  'mutasi_data',
  'one_on_one',
  'os_migrations',
  'outbox_sync',
  'overtime_records',
  'pic_list',
  'relievers',
  'roles',
  'sessions',
  'sop',
  'sp_data',
  'supply_requests',
  'sync_idempotency',
  'sync_locks',
  'sync_metrics',
  'sync_outbox',
  'sync_snapshots',
  'training',
  'users',
  'validation_options'
]);

const cache = new WeakMap();

function getDbCache(db) {
  let tableCache = cache.get(db);
  if (!tableCache) {
    tableCache = new Map();
    cache.set(db, tableCache);
  }
  return tableCache;
}

export async function hasTableColumn(db, tableName, columnName) {
  if (!db || !ALLOWED_TABLES.has(tableName)) return false;
  const tableCache = getDbCache(db);
  const cacheKey = `col:${tableName}.${columnName}`;
  if (tableCache.has(cacheKey)) {
    return tableCache.get(cacheKey);
  }

  let supported = true;
  try {
    const res = await db.prepare(`PRAGMA table_info(${tableName})`).all();
    const cols = res?.results || [];
    if (cols.length > 0) {
      supported = cols.some(c => c.name === columnName);
    }
  } catch {
    // Pure-JS mock D1 implementations may not handle PRAGMA; assume canonical schema
    supported = true;
  }

  tableCache.set(cacheKey, supported);
  return supported;
}

export async function hasTable(db, tableName) {
  if (!db || !ALLOWED_TABLES.has(tableName)) return false;
  const tableCache = getDbCache(db);
  const cacheKey = `tbl:${tableName}`;
  if (tableCache.has(cacheKey)) {
    return tableCache.get(cacheKey);
  }

  let exists = true;
  try {
    let isRealPragma = tableCache.get('__real_pragma__');
    if (isRealPragma === undefined) {
      const probe = await db.prepare('PRAGMA table_info(users)').all();
      isRealPragma = Array.isArray(probe?.results) && probe.results.length > 0;
      tableCache.set('__real_pragma__', isRealPragma);
    }
    if (isRealPragma) {
      const res = await db.prepare(`PRAGMA table_info(${tableName})`).all();
      const cols = res?.results;
      exists = Array.isArray(cols) && cols.length > 0;
    }
  } catch {
    exists = false;
  }

  tableCache.set(cacheKey, exists);
  return exists;
}

export function clearSchemaCompatCache(db) {
  if (db) cache.delete(db);
}
