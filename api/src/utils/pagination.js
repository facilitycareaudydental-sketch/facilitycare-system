// Pagination + query helpers

// Pagination + query helpers

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
