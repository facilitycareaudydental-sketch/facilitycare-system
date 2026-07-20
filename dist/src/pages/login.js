import { apiFetch, setToken, setUser } from '../config.js';
import { toastError, toastSuccess } from '../components/toast.js';

export async function renderLogin(container) {
  document.getElementById('app').innerHTML = `
    <div class="login-page">
      <div class="login-card">

        <div class="login-header">
          <div class="login-logo-wrap">
            <div class="login-logo-icon">
              <svg width="32" height="32" viewBox="0 0 48 48" fill="none">
                <rect width="48" height="48" rx="14" fill="url(#lg)"/>
                <path d="M12 20h6v16h-6V20zm10-8h6v24h-6V12zm10 6h6v18h-6V18z" fill="#fff" fill-opacity=".9"/>
                <defs>
                  <linearGradient id="lg" x1="0" y1="0" x2="48" y2="48" gradientUnits="userSpaceOnUse">
                    <stop stop-color="#2563EB"/>
                    <stop offset="1" stop-color="#0EA5E9"/>
                  </linearGradient>
                </defs>
              </svg>
            </div>
            <div>
              <h1 class="login-title">FCMS</h1>
              <p class="login-subtitle">Facility Care Management System</p>
            </div>
          </div>
          <div class="login-divider"></div>
          <p class="login-desc">Masuk untuk mengelola operasional Facility Care</p>
        </div>

        <form class="login-form" id="login-form" novalidate>
          <div class="form-group">
            <label class="form-label">Username / Email</label>
            <div class="input-with-icon">
              <svg class="input-prefix-icon" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/>
              </svg>
              <input type="text" name="username" class="form-control has-prefix-icon"
                placeholder="Masukkan username atau email"
                required autofocus autocomplete="username">
            </div>
          </div>

          <div class="form-group">
            <label class="form-label">Password</label>
            <div class="input-with-icon">
              <svg class="input-prefix-icon" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                <rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>
              </svg>
              <input type="password" name="password" class="form-control has-prefix-icon"
                placeholder="••••••••" required autocomplete="current-password" id="login-password">
              <button type="button" class="input-icon-btn" id="toggle-password" aria-label="Toggle password">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24" id="icon-eye">
                  <path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/>
                </svg>
              </button>
            </div>
          </div>

          <div id="login-error" class="alert alert-danger" style="display:none"></div>

          <button type="submit" class="btn btn-primary btn-full btn-lg" id="login-btn" style="margin-top:4px">
            <span class="btn-text">
              <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:6px">
                <path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"/><polyline points="10 17 15 12 10 7"/><line x1="15" y1="12" x2="3" y2="12"/>
              </svg>
              Masuk ke FCMS
            </span>
            <span class="btn-spinner" style="display:none">⏳ Memproses...</span>
          </button>
        </form>

        <div class="login-footer">
          <a href="#/form/chemical" class="link-subtle">
            📦 Form Permintaan Barang (Tanpa Login)
          </a>
        </div>

        <div class="login-version">FCMS v2.0 · Facility Care Indonesia</div>
      </div>

      <!-- Decorative background blobs -->
      <div class="login-blob login-blob-1"></div>
      <div class="login-blob login-blob-2"></div>
    </div>
  `;

  const form      = document.getElementById('login-form');
  const errorEl   = document.getElementById('login-error');
  const loginBtn  = document.getElementById('login-btn');
  const togglePwd = document.getElementById('toggle-password');
  const pwdInput  = document.getElementById('login-password');

  togglePwd?.addEventListener('click', () => {
    const isText = pwdInput.type === 'text';
    pwdInput.type = isText ? 'password' : 'text';
    togglePwd.style.color = isText ? '' : 'var(--primary)';
  });

  form?.addEventListener('submit', async (e) => {
    e.preventDefault();
    errorEl.style.display = 'none';

    const username = form.username.value.trim();
    const password = form.password.value;

    if (!username || !password) {
      errorEl.textContent = 'Username dan password wajib diisi.';
      errorEl.style.display = 'block';
      return;
    }

    loginBtn.querySelector('.btn-text').style.display    = 'none';
    loginBtn.querySelector('.btn-spinner').style.display = '';
    loginBtn.disabled = true;

    try {
      const res = await apiFetch('/api/auth/login', {
        method: 'POST',
        body: JSON.stringify({ username, password }),
      });

      if (res.ok && res.data.success) {
        setToken(res.data.data.token);
        setUser(res.data.data.user);
        toastSuccess('Login berhasil! Selamat datang 👋');
        window.dispatchEvent(new Event('fm:login'));
      } else {
        errorEl.textContent = res.data.error || 'Username atau password salah.';
        errorEl.style.display = 'block';
        loginBtn.classList.add('shake');
        setTimeout(() => loginBtn.classList.remove('shake'), 600);
      }
    } catch {
      errorEl.textContent = 'Gagal terhubung ke server. Periksa koneksi internet.';
      errorEl.style.display = 'block';
    } finally {
      loginBtn.querySelector('.btn-text').style.display    = '';
      loginBtn.querySelector('.btn-spinner').style.display = 'none';
      loginBtn.disabled = false;
    }
  });
}
