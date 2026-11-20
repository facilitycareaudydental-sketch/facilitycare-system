export const IS_DEVELOPMENT = false;

export const STAGING_WORKER_HOSTNAME = 'fm-operations-api-staging.facilitycare-audydental.workers.dev';
export const STAGING_API_URL = 'https://fm-operations-api-staging.facilitycare-audydental.workers.dev';

export const PRODUCTION_WORKER_HOSTNAME = 'fm-operations-api.facilitycare-audydental.workers.dev';
export const PRODUCTION_API_URL = 'https://fm-operations-api.facilitycare-audydental.workers.dev';

export const APPROVED_STAGING_TARGETS = Object.freeze([
  STAGING_API_URL
]);

export function isPlaceholder(val) {
  if (typeof val !== 'string') return false;
  const trimmed = val.trim();
  return trimmed.startsWith('__') && trimmed.endsWith('__') && trimmed.length >= 4;
}

/**
 * Resolves and validates the runtime API Base URL from window.__FM_CONFIG (or explicit override).
 * Strictly fail-closed:
 * - Never falls back to Production when environment is missing, unknown, or staging.
 * - In staging, permits ONLY the approved Staging Worker URL:
 *   https://fm-operations-api-staging.facilitycare-audydental.workers.dev
 * - Rejects Production Worker URL, FCVM URLs, unknown URLs, and missing/invalid configs in staging.
 */
export function resolveApiBaseUrl(configOverride = null) {
  const cfg = (configOverride !== null && configOverride !== undefined)
    ? configOverride
    : (typeof globalThis !== 'undefined' && globalThis.window && globalThis.window.__FM_CONFIG
        ? globalThis.window.__FM_CONFIG
        : null);

  if (!cfg || typeof cfg !== 'object') {
    throw new Error('CONFIG_ERROR: Missing runtime configuration (window.__FM_CONFIG). Fail-closed.');
  }

  const rawEnv = cfg.ENVIRONMENT !== undefined ? cfg.ENVIRONMENT : cfg.ENV;
  const rawUrl = cfg.API_BASE_URL !== undefined ? cfg.API_BASE_URL : cfg.API_BASE;

  // 1. Reject unreplaced build placeholders (__ENVIRONMENT__, __API_BASE_URL__)
  if (isPlaceholder(rawEnv) || isPlaceholder(rawUrl)) {
    throw new Error('CONFIG_ERROR: Unreplaced build configuration placeholders detected. Fail-closed.');
  }

  // 2. Reject missing or empty environment
  if (rawEnv === undefined || rawEnv === null || String(rawEnv).trim() === '') {
    throw new Error('CONFIG_ERROR: Environment is not configured. Explicit environment required.');
  }

  const cleanEnv = String(rawEnv).trim().toLowerCase();

  // 3. STAGING ENVIRONMENT: Strict Allowlist & Anti-Production Guard
  if (cleanEnv === 'staging') {
    if (rawUrl === undefined || rawUrl === null || String(rawUrl).trim() === '') {
      throw new Error('CONFIG_ERROR: Missing Staging API_BASE_URL configuration. Fail-closed.');
    }

    const cleanUrl = String(rawUrl).trim();
    if (cleanUrl.startsWith('//')) {
      throw new Error('CONFIG_ERROR: Protocol-relative URLs are not allowed in staging API_BASE_URL.');
    }

    let parsed;
    try {
      parsed = new URL(cleanUrl);
    } catch {
      throw new Error(`CONFIG_ERROR: Invalid or unknown Staging API_BASE_URL '${cleanUrl}'. Fail-closed.`);
    }

    if (parsed.protocol !== 'https:' && parsed.protocol !== 'http:') {
      throw new Error('CONFIG_ERROR: Staging API_BASE_URL must use https protocol.');
    }

    const hostname = parsed.hostname.toLowerCase().trim();

    if (hostname === PRODUCTION_WORKER_HOSTNAME) {
      throw new Error('SECURITY_VIOLATION: Staging environment cannot target Production Worker API.');
    }

    if (hostname.includes('fcvm')) {
      throw new Error('SECURITY_VIOLATION: Staging environment cannot target FCVM resources.');
    }

    if (parsed.username || parsed.password) {
      throw new Error('SECURITY_VIOLATION: Embedded credentials are not permitted in Staging API_BASE_URL.');
    }

    if (
      parsed.protocol !== 'https:' ||
      hostname !== STAGING_WORKER_HOSTNAME ||
      parsed.port !== '' ||
      (parsed.pathname && parsed.pathname !== '/') ||
      parsed.search !== '' ||
      parsed.hash !== ''
    ) {
      throw new Error(`SECURITY_VIOLATION: Unknown or unapproved Staging API target '${cleanUrl}'. Only '${STAGING_API_URL}' is permitted.`);
    }

    return STAGING_API_URL;
  }

  // 4. PRODUCTION ENVIRONMENT: Preserves existing Production behavior ONLY when explicitly 'production'
  if (cleanEnv === 'production') {
    if (rawUrl !== undefined && rawUrl !== null && String(rawUrl).trim() !== '') {
      const cleanUrl = String(rawUrl).trim();
      if (cleanUrl.startsWith('//')) {
        throw new Error('CONFIG_ERROR: Protocol-relative URLs are not allowed in production API_BASE_URL.');
      }
      let parsed;
      try {
        parsed = new URL(cleanUrl);
      } catch {
        throw new Error('CONFIG_ERROR: Production API_BASE_URL is a malformed URL.');
      }
      const hostname = parsed.hostname.toLowerCase().trim();
      if (hostname.includes('fcvm')) {
        throw new Error('SECURITY_VIOLATION: Production environment cannot target FCVM resources.');
      }
      if (hostname === STAGING_WORKER_HOSTNAME) {
        throw new Error('SECURITY_VIOLATION: Production environment cannot target Staging Worker API.');
      }
      if (parsed.protocol !== 'https:' || hostname !== PRODUCTION_WORKER_HOSTNAME) {
        throw new Error(`SECURITY_VIOLATION: Unknown Production API target '${cleanUrl}'.`);
      }
      return PRODUCTION_API_URL;
    }
    return PRODUCTION_API_URL;
  }

  // 5. DEVELOPMENT ENVIRONMENT (Local dev only, still blocks Production & FCVM)
  if (cleanEnv === 'development') {
    if (rawUrl === undefined || rawUrl === null || String(rawUrl).trim() === '' || String(rawUrl).trim() === '/api') {
      return '/api';
    }
    const cleanUrl = String(rawUrl).trim();
    if (cleanUrl === '/api') return '/api';
    let parsed;
    try {
      parsed = new URL(cleanUrl);
    } catch {
      throw new Error('CONFIG_ERROR: Development API_BASE_URL is a malformed URL.');
    }
    const hostname = parsed.hostname.toLowerCase().trim();
    if (hostname === PRODUCTION_WORKER_HOSTNAME) {
      throw new Error('SECURITY_VIOLATION: Development environment cannot target Production Worker API.');
    }
    if (hostname.includes('fcvm')) {
      throw new Error('SECURITY_VIOLATION: Development environment cannot target FCVM resources.');
    }
    if (hostname !== 'localhost' && hostname !== '127.0.0.1' && hostname !== STAGING_WORKER_HOSTNAME) {
      throw new Error(`SECURITY_VIOLATION: Unknown Development API target '${cleanUrl}'.`);
    }
    return cleanUrl.replace(/\/+$/, '');
  }

  // 6. Unknown / Unrecognized Environment -> Fail-Closed
  throw new Error(`CONFIG_ERROR: Unknown environment '${cleanEnv}'. Fail-closed.`);
}

export function buildApiUrl(apiBase, path) {
  const cleanBase = String(apiBase || '').trim().replace(/\/+$/, '');
  const cleanPath = String(path || '').trim();
  const normalizedPath = cleanPath.startsWith('/') ? cleanPath : `/${cleanPath}`;

  if (!cleanBase) {
    throw new Error('CONFIG_ERROR: API base URL cannot be empty when constructing request URL.');
  }
  if (cleanBase === '/api' && (normalizedPath === '/api' || normalizedPath.startsWith('/api/'))) {
    return normalizedPath;
  }
  return `${cleanBase}${normalizedPath}`;
}

export let API_BASE_URL = '';
export let API = '';

if (typeof globalThis !== 'undefined' && globalThis.window && globalThis.window.__FM_CONFIG) {
  API_BASE_URL = resolveApiBaseUrl(globalThis.window.__FM_CONFIG);
  API = API_BASE_URL;
}

let _ephemeralMemoryToken = null;
let _ephemeralCsrfToken = null;
let _csrfRecoveryPromise = null;

function readBrowserCookie(name) {
  if (typeof document === 'undefined' || !document.cookie) return null;
  const prefix = `${name}=`;
  for (const part of document.cookie.split(';')) {
    const trimmed = part.trim();
    if (trimmed.startsWith(prefix)) {
      try {
        return decodeURIComponent(trimmed.slice(prefix.length));
      } catch {
        return trimmed.slice(prefix.length);
      }
    }
  }
  return null;
}

export function getToken() {
  if (typeof localStorage !== 'undefined') {
    localStorage.removeItem('fm_token');
  }
  return _ephemeralMemoryToken;
}

export function setToken(token) {
  if (typeof localStorage !== 'undefined') {
    localStorage.removeItem('fm_token');
  }
  _ephemeralMemoryToken = token || null;
}

export function getCsrfToken() {
  return _ephemeralCsrfToken || readBrowserCookie('fcms_csrf');
}

export function setCsrfToken(token) {
  _ephemeralCsrfToken = token || null;
}

export function clearToken() {
  _ephemeralMemoryToken = null;
  _ephemeralCsrfToken = null;
  if (typeof localStorage !== 'undefined') {
    localStorage.removeItem('fm_token');
    localStorage.removeItem('fm_user');
  }
}

export function getUser() {
  if (typeof localStorage === 'undefined') return null;
  try {
    return JSON.parse(localStorage.getItem('fm_user') || 'null');
  } catch {
    return null;
  }
}

export function setUser(user) {
  if (typeof localStorage === 'undefined') return;
  localStorage.setItem('fm_user', JSON.stringify(user));
}

export const ASSIGNABLE_MODULES = [
  { id: 'employees', label: 'Data Karyawan (HR)' },
  { id: 'contracts', label: 'Data Kontrak (HR)' },
  { id: 'sp', label: 'Surat Peringatan / SP (HR)' },
  { id: 'mutasi', label: 'Data Mutasi (HR)' },
  { id: 'overtime', label: 'Data Lembur (HR)' },
  { id: 'schedule', label: 'Time Line / Jadwal (Operasional)' },
  { id: 'os_migrations', label: 'Migrasi OS (Operasional)' },
  { id: 'issues', label: 'Permasalahan / Issues (Operasional)' },
  { id: 'training', label: 'Training Karyawan (Operasional)' },
  { id: 'reports', label: 'Laporan Kebersihan / Basecamp' },
  { id: 'sop', label: 'SOP Facility Care & Security' },
  { id: 'checklist', label: 'Master Dokumen & Formulir' },
  { id: 'hygiene_standards', label: 'Master Standar Hygiene' },
  { id: 'dac', label: 'Master Standar DAC' },
  { id: 'one_on_one', label: 'Konseling One on One' },
  { id: 'relievers', label: 'Data Reliever' },
  { id: 'branches', label: 'Data Cabang' },
];

export const ADMIN_MODULES = [
  'users',
  'audit_logs',
  'audit-logs',
  'sync',
  'sync_dashboard',
  'monitoring',
  'settings',
  'system_settings',
  'import'
];

export const ASSIGNABLE_MODULES_SET = new Set(ASSIGNABLE_MODULES.map(m => m.id));

export function getAssignedModules(userOrRaw) {
  if (!userOrRaw) return [];
  let raw = userOrRaw;
  if (typeof userOrRaw === 'object' && !Array.isArray(userOrRaw)) {
    raw = userOrRaw.assigned_modules !== undefined 
      ? userOrRaw.assigned_modules 
      : (userOrRaw.assignedModules || userOrRaw.modules);
  }
  if (!raw) return [];
  const normalizeMod = (m) => String(m).trim().toLowerCase().replace(/-/g, '_');
  const filterValid = (arr) => Array.from(new Set(arr.map(normalizeMod).filter(m => m && ASSIGNABLE_MODULES_SET.has(m))));
  if (Array.isArray(raw)) {
    return filterValid(raw);
  }
  if (typeof raw === 'string') {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) {
        return filterValid(parsed);
      }
    } catch {
      // not JSON, fallback to comma-separated
    }
    return filterValid(raw.split(','));
  }
  return [];
}

export const OFFICIAL_ROLES = Object.freeze(['superadmin', 'fc_specialist_leader']);

export function hasPermission(user, moduleOrAction = 'read', action = null) {
  if (!user || !user.role) return false;
  if (user.isRestrictedToken === true || Number(user.must_change_password) === 1) return false;
  const role = String(user.role).trim().toLowerCase();

  // 1. SUPERADMIN: Full access
  if (role === 'superadmin') return true;

  let module = null;
  let act = 'read';

  if (action !== null) {
    module = moduleOrAction;
    act = action;
  } else {
    const raw = String(moduleOrAction || 'read').toLowerCase();
    if (['read', 'write', 'create', 'update', 'delete', 'export', 'import', 'admin'].includes(raw)) {
      act = raw;
      module = null;
    } else {
      module = raw;
      act = 'read';
    }
  }

  const cleanAct = String(act || 'read').trim().toLowerCase();
  const cleanModule = module ? String(module).trim().toLowerCase().replace(/-/g, '_') : '';

  // 2. Administrative action and modules strictly forbidden for non-superadmin
  if (cleanAct === 'admin') {
    return false;
  }
  if (cleanModule && (ADMIN_MODULES.includes(cleanModule) || ADMIN_MODULES.includes(String(module).trim().toLowerCase()))) {
    return false;
  }

  // 3. FC_SPECIALIST_LEADER: Only authorized non-superadmin official operational role
  // - Read-only across operational & master modules (except admin modules blocked above)
  // - CREATE and UPDATE/WRITE allowed ONLY on overtime and issues
  // - Strictly NO delete, export, import, admin, or writes on other modules
  if (role === 'fc_specialist_leader') {
    if (cleanAct === 'read') return true;
    if (cleanAct === 'create' || cleanAct === 'update' || cleanAct === 'write') {
      return cleanModule === 'overtime' || cleanModule === 'issues';
    }
    return false;
  }

  // All legacy roles (admin, supervisor, spv, editor_khusus, input_lembur, viewer, user)
  // and unknown roles receive zero fallback access (deny by default).
  return false;
}

async function ensureCsrfTokenReady(options = {}) {
  if (getToken() || getCsrfToken()) return;
  if (_csrfRecoveryPromise) return _csrfRecoveryPromise;
  _csrfRecoveryPromise = (async () => {
    try {
      const activeApiBase = resolveApiBaseUrl(options.configOverride || null);
      const fetchFn = (typeof options.fetchImpl === 'function') ? options.fetchImpl : fetch;
      const meUrl = buildApiUrl(activeApiBase, `/api/auth/me?cb=${Date.now()}`);
      const res = await fetchFn(meUrl, {
        method: 'GET',
        credentials: 'include',
        headers: { 'Content-Type': 'application/json' }
      });
      if (res.ok) {
        const json = await res.json().catch(() => null);
        const csrf = json?.data?.csrf_token;
        if (typeof csrf === 'string' && csrf.length > 0) {
          setCsrfToken(csrf);
        }
      }
    } catch {
      // Fail closed: subsequent mutating request proceeds without header and is rejected with 401
    } finally {
      _csrfRecoveryPromise = null;
    }
  })();
  return _csrfRecoveryPromise;
}

export async function apiFetch(path, options = {}) {
  // Always resolve & validate runtime API base URL fail-closed before dispatching any request
  const activeApiBase = resolveApiBaseUrl(options.configOverride || null);
  API_BASE_URL = activeApiBase;
  API = activeApiBase;

  const methodUpper = String(options.method || 'GET').toUpperCase();
  if (
    !options.fetchImpl &&
    ['POST', 'PUT', 'PATCH', 'DELETE'].includes(methodUpper) &&
    !path.startsWith('/api/auth/login') &&
    getUser() &&
    !getToken() &&
    !getCsrfToken()
  ) {
    await ensureCsrfTokenReady(options);
  }

  const token = getToken();
  const csrfToken = getCsrfToken();
  const headers = {
    'Content-Type': 'application/json',
    ...(token ? { Authorization: `Bearer ${token}` } : {}),
    ...(csrfToken && ['POST', 'PUT', 'PATCH', 'DELETE'].includes(methodUpper) ? { 'X-CSRF-Token': csrfToken } : {}),
    ...(options.headers || {}),
  };
  
  const defaultTimeoutMs = (path.includes('/import') || path.includes('/backup') || path.includes('/export')) ? 60000 : 20000;
  const timeoutMs = (typeof options.timeoutMs === 'number' && options.timeoutMs > 0) ? options.timeoutMs : defaultTimeoutMs;
  let controller = null;
  let timeoutId = null;
  let timedOut = false;

  if (!options.signal && typeof AbortController === 'function') {
    controller = new AbortController();
    timeoutId = setTimeout(() => {
      timedOut = true;
      controller.abort();
    }, timeoutMs);
  }

  try {
    const cacheBuster = `cb=${Date.now()}`;
    const separator = path.includes('?') ? '&' : '?';
    const pathWithQuery = `${path}${separator}${cacheBuster}`;
    const finalUrl = buildApiUrl(activeApiBase, pathWithQuery);
    
    const fetchFn = (typeof options.fetchImpl === 'function') ? options.fetchImpl : fetch;
    const fetchOptions = {
      credentials: 'include',
      ...options,
      headers,
      ...(controller ? { signal: controller.signal } : {})
    };
    delete fetchOptions.timeoutMs;
    delete fetchOptions.fetchImpl;
    delete fetchOptions.configOverride;

    const res = await fetchFn(finalUrl, fetchOptions);
    let data;
    try {
      const text = await res.text();
      try {
        data = JSON.parse(text);
      } catch (e) {
        data = { error: `Server Error (${res.status}): ${text.substring(0, 80)}...` };
      }
    } catch (e) {
      data = { error: 'Gagal membaca respon dari server' };
    }
    
    if (res.status === 401) {
      clearToken();
      if (typeof window !== 'undefined' && window.location) {
        window.location.hash = '#/login';
      }
    }
    return { ok: res.ok, status: res.status, data };
  } catch (err) {
    if (timedOut || err?.name === 'AbortError') {
      return { ok: false, status: 408, data: { error: 'Permintaan waktu habis (timeout). Silakan coba lagi.' } };
    }
    // Network errors (e.g. offline)
    return { ok: false, status: 0, data: { error: `Koneksi terputus. Periksa jaringan Anda. (${err.message})` } };
  } finally {
    if (timeoutId !== null) clearTimeout(timeoutId);
  }
}
