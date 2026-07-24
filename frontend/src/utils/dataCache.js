import { apiFetch, CLIENT_SIDE_MAX_ROWS } from '../config.js';

let _employeesCache = null;
let _branchesCache = null;

export async function getCachedEmployees(forceRefresh = false) {
  if (_employeesCache && !forceRefresh) {
    console.log("Employees Raw (Cache Hit)", _employeesCache.slice(0, 5));
    return _employeesCache;
  }
  const res = await apiFetch(`/api/employees?limit=${CLIENT_SIDE_MAX_ROWS}&status=Aktif`);
  _employeesCache = (res.data?.data || []).map(e => ({ value: e.id, label: e.full_name }));
  console.log("Employees Raw", res.data?.data?.slice(0, 5));
  console.log("Employees Mapped (ID)", _employeesCache.slice(0, 5));
  return _employeesCache;
}

export async function getCachedEmployeeNames(forceRefresh = false) {
  const employees = await getCachedEmployees(forceRefresh);
  const options = employees.map(e => ({ value: e.label, label: e.label }));
  console.log("Employee Options", options.slice(0, 5));
  return options;
}

export async function getCachedBranches(forceRefresh = false) {
  if (_branchesCache && !forceRefresh) return _branchesCache;
  const res = await apiFetch('/api/branches?all=1');
  _branchesCache = (res.data?.data || []).map(b => ({ value: b.id, label: b.full_name }));
  return _branchesCache;
}

export function clearDataCache() {
  _employeesCache = null;
  _branchesCache = null;
}
