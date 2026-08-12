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

export async function apiFetch(path, options = {}) {
  const token = getToken();
  const headers = {
    'Content-Type': 'application/json',
    ...(token ? { Authorization: `Bearer ${token}` } : {}),
    ...(options.headers || {}),
  };
  
  try {
    let finalPath = path;
    if (finalPath.startsWith('/api/') && !finalPath.startsWith('/api/v1/')) {
      finalPath = '/api/v1/' + finalPath.substring(5);
    }
    const cacheBuster = `cb=${Date.now()}`;
    const separator = finalPath.includes('?') ? '&' : '?';
    const finalUrl = `${API}${finalPath}${separator}${cacheBuster}`;
    
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
