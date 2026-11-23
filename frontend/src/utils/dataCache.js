import { apiFetch } from '../config.js';
import { normalizeBranchDropdownOptions } from './branchMaster.js';

export const MAX_PAGINATED_LIMIT = 100;

let _employeesCache = null;
let _branchesCache = null;

/**
 * Fetches all matching pages from a paginated endpoint using sequential `limit<=100`
 * requests, deduplicating rows by `id` while preserving server ordering.
 */
export async function fetchAllPaginated(basePath, params = {}, options = {}) {
  const limit = Math.min(MAX_PAGINATED_LIMIT, Math.max(1, Number(options.limit) || MAX_PAGINATED_LIMIT));
  const maxPages = Math.max(1, Number(options.maxPages) || 500);
  const allRows = [];
  const seenIds = new Set();

  for (let page = 1; page <= maxPages; page++) {
    const queryParams = new URLSearchParams();
    queryParams.set('page', String(page));
    queryParams.set('limit', String(limit));

    if (params && typeof params === 'object') {
      for (const [key, val] of Object.entries(params)) {
        if (val !== undefined && val !== null && val !== '' && val !== 'all') {
          queryParams.set(key, String(val));
        }
      }
    }

    const res = await apiFetch(`${basePath}?${queryParams.toString()}`);
    if (!res.ok) {
      throw new Error(res.data?.error || `Gagal memuat data (HTTP ${res.status || 500})`);
    }

    const pageRows = Array.isArray(res.data?.data)
      ? res.data.data
      : (Array.isArray(res.data) ? res.data : []);

    for (const row of pageRows) {
      if (row && row.id !== undefined && row.id !== null) {
        if (seenIds.has(row.id)) continue;
        seenIds.add(row.id);
      }
      allRows.push(row);
    }

    const pag = res.data?.pagination;
    const totalPages = pag && Number(pag.pages) >= 0 ? Number(pag.pages) : null;
    if (totalPages !== null) {
      if (page >= totalPages) break;
    } else if (pageRows.length < limit) {
      break;
    }
    if (pageRows.length === 0) break;
  }

  return allRows;
}

/**
 * Accumulates all matching rows from a dedicated backend `/export` endpoint
 * (max 100 rows per request) preserving active filters and sorting.
 */
export async function fetchExportRows(exportPath, filters = {}, options = {}) {
  return fetchAllPaginated(exportPath, filters, { ...options, limit: MAX_PAGINATED_LIMIT });
}

/**
 * Queries `/api/employees/selector` with optional search/filter params and bounded limit (`<= 100`).
 */
export async function searchEmployeeSelector({
  search = '',
  status = 'Aktif',
  branch_id = '',
  division = '',
  page = 1,
  limit = 20,
} = {}) {
  const safeLimit = Math.min(MAX_PAGINATED_LIMIT, Math.max(1, Number(limit) || 20));
  const queryParams = new URLSearchParams({
    page: String(Math.max(1, Number(page) || 1)),
    limit: String(safeLimit),
  });
  if (search) queryParams.set('search', search);
  if (status && status !== 'all') queryParams.set('status', status);
  if (branch_id && branch_id !== 'all') queryParams.set('branch_id', branch_id);
  if (division && division !== 'all') queryParams.set('division', division);

  const res = await apiFetch(`/api/employees/selector?${queryParams.toString()}`);
  if (!res.ok) {
    throw new Error(res.data?.error || `Gagal mencari karyawan (HTTP ${res.status || 500})`);
  }
  return {
    data: Array.isArray(res.data?.data) ? res.data.data : [],
    pagination: res.data?.pagination || { total: 0, page: 1, limit: safeLimit, pages: 0 },
  };
}

export async function getCachedEmployees(forceRefresh = false) {
  if (_employeesCache && !forceRefresh) {
    return _employeesCache;
  }
  const rows = await fetchAllPaginated('/api/employees/selector', { status: 'Aktif' });
  _employeesCache = rows.map(e => ({
    value: e.id,
    id: e.id,
    label: e.full_name,
    full_name: e.full_name,
    branch_id: e.branch_id ?? null,
    branch_name: e.branch_name || '',
    division: e.division || '',
    status: e.status || '',
  }));
  return _employeesCache;
}

export async function getCachedEmployeeNames(forceRefresh = false) {
  const employees = await getCachedEmployees(forceRefresh);
  return employees.map(e => ({ value: e.label, label: e.label }));
}

export async function getCachedBranches(forceRefresh = false) {
  if (_branchesCache && !forceRefresh) return _branchesCache;
  const res = await apiFetch('/api/branches?all=1');
  if (!res.ok) {
    throw new Error(res.data?.error || 'Gagal memuat data cabang');
  }
  _branchesCache = normalizeBranchDropdownOptions(res.data?.data || []);
  return _branchesCache;
}

export function clearDataCache() {
  _employeesCache = null;
  _branchesCache = null;
}
