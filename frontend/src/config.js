// Force API calls through Same-Origin Pages Proxy to bypass strict CORS/Firewalls
const API_BASE_URL = '';

export const API = API_BASE_URL;

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
      data = await res.json();
    } catch (e) {
      // If response is not JSON (e.g. 500 HTML error from Cloudflare)
      data = { error: 'Terjadi kesalahan pada server (Network/Server Error)' };
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
