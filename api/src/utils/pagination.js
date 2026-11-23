// Pagination, query, and standardized sorting helpers
import { hasTableColumn } from './schemaCompat.js';

export { hasTableColumn };

export const DEFAULT_PAGE_LIMIT = 20;
// Hard upper bound for paginated API responses enforced across all endpoints (E10).
export const DEFAULT_MAX_PAGE_LIMIT = 100;
export const HARD_MAX_PAGE_LIMIT = 100;

function parseBoundedInt(raw, fallback, min, max) {
  if (raw === null || raw === undefined) return fallback;
  const trimmed = String(raw).trim();
  if (!trimmed || !/^-?\d+$/.test(trimmed)) return fallback;
  if (trimmed.startsWith('-')) return min;
  if (trimmed.length > 12) return max;
  const parsed = Number.parseInt(trimmed, 10);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.min(max, Math.max(min, parsed));
}

export function getPagination(url, options = {}) {
  const defaultLimit = options.defaultLimit ?? DEFAULT_PAGE_LIMIT;
  const maxLimit = options.maxLimit ?? DEFAULT_MAX_PAGE_LIMIT;
  const params = new URL(url).searchParams;
  const page = parseBoundedInt(params.get('page'), 1, 1, 1000000);
  const limit = parseBoundedInt(params.get('limit'), defaultLimit, 1, maxLimit);
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
    `WHEN ${t} GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]*' AND date(${isoSub}) = ${isoSub} THEN ${isoSub}`,
    `WHEN ${t} GLOB '[0-9][0-9][-/][0-9][0-9][-/][0-9][0-9][0-9][0-9]' AND date(${dmyIso}) = ${dmyIso} THEN ${dmyIso}`,
  ];
  if (allowExcelSerial) {
    parts.push(
      `WHEN ${t} GLOB '[0-9][0-9][0-9][0-9][0-9]' AND CAST(${t} AS INTEGER) >= 10000 THEN date('1899-12-30', '+' || CAST(${t} AS INTEGER) || ' days')`
    );
  }
  if (allowYearMonth) {
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
  secondaryDateColumns = [],
} = {}) {
  const dir = String(direction).toUpperCase() === 'ASC' ? 'ASC' : 'DESC';
  const normExpr = buildNormalizedDateExpr(colExpr, { allowYearMonth, allowExcelSerial });
  const tieBreaker = buildTieBreaker({ alias, includeUpdatedAt, includeCreatedAt, idDir });
  const secList = Array.isArray(secondaryDateColumns) ? secondaryDateColumns.filter(Boolean) : [];
  if (secList.length === 0) {
    return `CASE WHEN (${normExpr}) IS NULL THEN 1 ELSE 0 END ASC, (${normExpr}) ${dir}, ${tieBreaker}`;
  }
  const secExprs = secList.map(c => buildNormalizedDateExpr(c, { allowYearMonth: false, allowExcelSerial }));
  const secCoalesce = secExprs.length === 1 ? `(${secExprs[0]})` : `COALESCE(${secExprs.join(', ')})`;
  // Primary valid date ALWAYS sorts first (tier 0) strictly by normExpr.
  // Secondary date columns ONLY order rows whose primary date is NULL/empty/invalid (tier 1),
  // leaving completely dateless rows at tier 2 (absolute bottom).
  return `CASE WHEN (${normExpr}) IS NOT NULL THEN 0 WHEN (${secCoalesce}) IS NOT NULL THEN 1 ELSE 2 END ASC, (${normExpr}) ${dir}, CASE WHEN (${normExpr}) IS NULL THEN (${secCoalesce}) ELSE NULL END ${dir}, ${tieBreaker}`;
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
