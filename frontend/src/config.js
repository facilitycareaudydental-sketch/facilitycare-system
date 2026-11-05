export const IS_DEVELOPMENT = false;
const API_BASE_URL = 'https://fm-operations-api.facilitycare-audydental.workers.dev';
export const API = API_BASE_URL;

export const CLIENT_SIDE_MAX_ROWS = 10000;

export function getToken() {
  return localStorage.getItem('fm_token');
}

export function setToken(token) {
  localStorage.setItem('fm_token', token);
}

export function clearToken() {
  localStorage.removeItem('fm_token');
  localStorage.removeItem('fm_user');
}

export function getUser() {
  try {
    return JSON.parse(localStorage.getItem('fm_user') || 'null');
  } catch {
    return null;
  }
}

export function setUser(user) {
  localStorage.setItem('fm_user', JSON.stringify(user));
}

export function hasPermission(user, moduleOrAction = 'read', action = null) {
  if (!user || !user.role) return false;
  const role = user.role;
  if (role === 'superadmin' || role === 'admin') return true;

  let module = null;
  let act = 'read';

  if (action !== null) {
    module = moduleOrAction;
    act = action;
  } else {
    act = moduleOrAction || 'read';
  }

  // Strictly no export, import, or admin privileges for any non-admin / non-superadmin
  if (act === 'export' || act === 'import' || act === 'admin') {
    return false;
  }

  // EDITOR KHUSUS (Hanya Lembur & Masalah)
  if (role === 'editor_khusus') {
    if (module === 'issues' || module === 'overtime') {
      return act === 'read' || act === 'write' || act === 'delete';
    }
    return act === 'read';
  }

  // INPUT LEMBUR (Hanya Lembur)
  if (role === 'input_lembur') {
    if (module === 'overtime') {
      return act === 'read' || act === 'write';
    }
    return act === 'read';
  }

  // "pokoknya selain akun super admin /admin cukup viewer saja"
  // All other roles (spv, manager, viewer, etc.) are strictly VIEWER ONLY across all modules
  return act === 'read';
}

export async function apiFetch(path, options = {}) {
  const token = getToken();
  const headers = {
    'Content-Type': 'application/json',
    ...(token ? { Authorization: `Bearer ${token}` } : {}),
    ...(options.headers || {}),
  };
  
  try {
    const cacheBuster = `cb=${Date.now()}`;
    const separator = path.includes('?') ? '&' : '?';
    const finalUrl = `${API}${path}${separator}${cacheBuster}`;
    
    const res = await fetch(finalUrl, { ...options, headers });
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
      window.location.hash = '#/login';
    }
    return { ok: res.ok, status: res.status, data };
  } catch (err) {
    // Network errors (e.g. offline)
    return { ok: false, status: 0, data: { error: `Koneksi terputus. Periksa jaringan Anda. (${err.message})` } };
  }
}
