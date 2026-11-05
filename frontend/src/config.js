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

    if (act === 'export') {
      return role === 'manager';
    }

    if (role === 'input_lembur') {
      if (module === 'overtime') {
        if (act === 'read' || act === 'write') return true;
        if (act === 'delete') return false;
      } else {
        if (act === 'read') return true;
      }
      return false;
    }

    if (role === 'editor_khusus') {
      if (module === 'issues' || module === 'overtime') {
        if (act === 'read' || act === 'write' || act === 'delete') return true;
      } else {
        if (act === 'read') return true;
      }
      return false;
    }
  } else {
    act = moduleOrAction || 'read';
  }

  if (role === 'editor_khusus' || role === 'input_lembur' || role === 'viewer') {
    if (act === 'export' || act === 'write' || act === 'delete' || act === 'admin') return false;
  }

  const roleLevels = {
    superadmin: 5,
    admin: 4,
    manager: 3,
    spv: 2,
    editor_khusus: 1,
    input_lembur: 1,
    viewer: 1
  };
  const level = roleLevels[role] || 0;
  if (act === 'read') return level >= 1;
  if (act === 'write') return level >= 2;
  if (act === 'delete') return level >= 3;
  if (act === 'export') return level >= 3;
  if (act === 'admin') return level >= 4;
  return false;
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
