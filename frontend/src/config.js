// Force API calls through Same-Origin Pages Proxy to bypass strict CORS/Firewalls
const API_BASE_URL = '';

export const API = API_BASE_URL;
export const IS_DEVELOPMENT = window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
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
