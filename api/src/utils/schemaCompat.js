/**
 * Lightweight runtime schema compatibility probe with WeakMap caching per DB handle.
 * Used when minimal test mocks define subsets of canonical columns, without changing
 * canonical D1 behavior or running unapproved schema migrations.
 */
const ALLOWED_TABLES = new Set(['contracts', 'one_on_one', 'relievers']);
const cache = new WeakMap();

export async function hasTableColumn(db, tableName, columnName) {
  if (!db || !ALLOWED_TABLES.has(tableName)) return false;
  let tableCache = cache.get(db);
  if (!tableCache) {
    tableCache = new Map();
    cache.set(db, tableCache);
  }
  const cacheKey = `${tableName}.${columnName}`;
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
