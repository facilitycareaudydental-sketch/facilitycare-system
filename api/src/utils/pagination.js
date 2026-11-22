// Pagination, query, and standardized sorting helpers

const _tableColumnsCache = new WeakMap();

/**
 * Checks whether a table exposes a given column (cached per DB handle).
 * Prevents query failures across older minimal schemas and live D1 schemas.
 */
export async function hasTableColumn(db, tableName, columnName) {
  if (!db || !tableName || !columnName) return false;
  let tableCache = _tableColumnsCache.get(db);
  if (!tableCache) {
    tableCache = new Map();
    _tableColumnsCache.set(db, tableCache);
  }
  let cols = tableCache.get(tableName);
  if (!cols) {
    try {
      const res = await db.prepare(`PRAGMA table_info(${tableName})`).all();
      cols = new Set((res?.results || []).map(r => String(r.name || '').toLowerCase()));
      tableCache.set(tableName, cols);
    } catch {
      return false;
    }
  }
  return cols.has(String(columnName).toLowerCase());
}

export function getPagination(url) {
  const params = new URL(url).searchParams;
  const parsedPage = Number.parseInt(params.get('page') || '1', 10);
  const parsedLimit = Number.parseInt(params.get('limit') || '20', 10);
  const page = Number.isFinite(parsedPage) ? Math.max(1, parsedPage) : 1;
  const limit = Number.isFinite(parsedLimit) ? Math.min(10000, Math.max(1, parsedLimit)) : 20;
  const offset = (page - 1) * limit;
  return { page, limit, offset };
}

export function resolveSortOrder(rawSort, sortMap, defaultKey) {
  const key = String(rawSort || '').trim();
  if (key && Object.prototype.hasOwnProperty.call(sortMap, key)) {
    return sortMap[key];
  }
  return sortMap[defaultKey];
}

/**
 * Builds a SQLite CASE expression normalizing heterogeneous date strings into
 * canonical ISO 'YYYY-MM-DD' text, returning NULL for empty/dash/invalid dates.
 */
export function buildNormalizedDateExpr(colExpr, { allowYearMonth = false, allowExcelSerial = true } = {}) {
  const t = `TRIM(${colExpr})`;
  const isoSub = `substr(${t}, 1, 10)`;
  const dmyIso = `(substr(${t}, 7, 4) || '-' || substr(${t}, 4, 2) || '-' || substr(${t}, 1, 2))`;
  const parts = [
    `WHEN ${colExpr} IS NULL OR ${t} = '' OR ${t} = '-' THEN NULL`,
    // ISO YYYY-MM-DD or YYYY-MM-DDTHH:MM:SS...
    `WHEN ${t} GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]*' AND date(${isoSub}) = ${isoSub} THEN ${isoSub}`,
    // DD-MM-YYYY or DD/MM/YYYY
    `WHEN ${t} GLOB '[0-9][0-9][-/][0-9][0-9][-/][0-9][0-9][0-9][0-9]' AND date(${dmyIso}) = ${dmyIso} THEN ${dmyIso}`,
  ];
  if (allowExcelSerial) {
    // 5-digit Excel serial date (e.g. 45369 -> 2024-03-18 using Excel epoch 1899-12-30)
    parts.push(
      `WHEN ${t} GLOB '[0-9][0-9][0-9][0-9][0-9]' AND CAST(${t} AS INTEGER) >= 10000 THEN date('1899-12-30', '+' || CAST(${t} AS INTEGER) || ' days')`
    );
  }
  if (allowYearMonth) {
    // YYYY-MM (e.g. os_migrations.target_month)
    parts.push(
      `WHEN ${t} GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]' AND date(${t} || '-01') = (${t} || '-01') THEN (${t} || '-01')`
    );
  }
  return `CASE ${parts.join(' ')} ELSE NULL END`;
}

/**
 * Builds a SQLite expression normalizing ISO/SQLite timestamps into
 * 'YYYY-MM-DD HH:MM:SS.SSS' for accurate chronological sorting across 'T' and space formats.
 */
export function buildNormalizedTimestampExpr(colExpr) {
  const t = `TRIM(${colExpr})`;
  const dateFallback = buildNormalizedDateExpr(colExpr, { allowYearMonth: false, allowExcelSerial: false });
  return `CASE WHEN ${colExpr} IS NULL OR ${t} = '' OR ${t} = '-' THEN NULL ELSE COALESCE(strftime('%Y-%m-%d %H:%M:%f', ${t}), datetime(${t}), (${dateFallback})) END`;
}

/**
 * Standard FCMS tie-breaker: updated_at DESC (when available), created_at DESC, id DESC.
 */
export function buildTieBreaker({
  alias = '',
  includeUpdatedAt = true,
  includeCreatedAt = true,
  idDir = 'DESC',
} = {}) {
  const prefix = alias ? `${alias}.` : '';
  const dir = String(idDir).toUpperCase() === 'ASC' ? 'ASC' : 'DESC';
  const clauses = [];
  if (includeUpdatedAt) {
    const updExpr = buildNormalizedTimestampExpr(`${prefix}updated_at`);
    clauses.push(`CASE WHEN (${updExpr}) IS NULL THEN 1 ELSE 0 END ASC`, `(${updExpr}) DESC`);
  }
  if (includeCreatedAt) {
    const crExpr = buildNormalizedTimestampExpr(`${prefix}created_at`);
    clauses.push(`CASE WHEN (${crExpr}) IS NULL THEN 1 ELSE 0 END ASC`, `(${crExpr}) DESC`);
  }
  clauses.push(`${prefix}id ${dir}`);
  return clauses.join(', ');
}

/**
 * Builds a canonical SQL ORDER BY clause expression for a date column.
 * Empty/NULL/dash/invalid dates always sort LAST regardless of ASC or DESC direction.
 */
export function buildDateOrderBy(colExpr, {
  direction = 'DESC',
  alias = '',
  includeUpdatedAt = true,
  includeCreatedAt = true,
  idDir = 'DESC',
  allowYearMonth = false,
  allowExcelSerial = true,
} = {}) {
  const dir = String(direction).toUpperCase() === 'ASC' ? 'ASC' : 'DESC';
  const normExpr = buildNormalizedDateExpr(colExpr, { allowYearMonth, allowExcelSerial });
  const tieBreaker = buildTieBreaker({ alias, includeUpdatedAt, includeCreatedAt, idDir });
  return `CASE WHEN (${normExpr}) IS NULL THEN 1 ELSE 0 END ASC, (${normExpr}) ${dir}, ${tieBreaker}`;
}

/**
 * Builds a canonical SQL ORDER BY clause expression for a timestamp column (e.g. created_at).
 */
export function buildTimestampOrderBy(colExpr, {
  direction = 'DESC',
  alias = '',
  includeUpdatedAt = false,
  includeCreatedAt = false,
  idDir = 'DESC',
} = {}) {
  const dir = String(direction).toUpperCase() === 'ASC' ? 'ASC' : 'DESC';
  const normExpr = buildNormalizedTimestampExpr(colExpr);
  const tieBreaker = buildTieBreaker({ alias, includeUpdatedAt, includeCreatedAt, idDir });
  return `CASE WHEN (${normExpr}) IS NULL THEN 1 ELSE 0 END ASC, (${normExpr}) ${dir}, ${tieBreaker}`;
}

export function getSearchParam(url, key) {
  return new URL(url).searchParams.get(key) || '';
}

export function buildWhereClause(filters) {
  const conditions = [];
  const values = [];
  for (const [key, val] of Object.entries(filters)) {
    if (val !== null && val !== undefined && val !== '') {
      if (val.toString().includes('%')) {
        conditions.push(`${key} LIKE ?`);
      } else {
        conditions.push(`${key} = ?`);
      }
      values.push(val);
    }
  }
  const clause = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';
  return { clause, values };
}
