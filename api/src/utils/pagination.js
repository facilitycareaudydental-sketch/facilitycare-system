// Pagination + query helpers

export function getPagination(url) {
  const params = new URL(url).searchParams;
  const page = Math.max(1, parseInt(params.get('page') || '1'));
  const limit = Math.min(100, Math.max(1, parseInt(params.get('limit') || '20')));
  const offset = (page - 1) * limit;
  return { page, limit, offset };
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
