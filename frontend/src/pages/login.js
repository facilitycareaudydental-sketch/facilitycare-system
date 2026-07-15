import { apiFetch, setToken, setUser } from '../config.js';
import { toastError, toastSuccess } from '../components/toast.js';

export async function renderLogin(container) {
  document.getElementById('app').innerHTML = `
    <div class="login-page">
      <div class="login-card">
        <div class="login-header">
          <div class="login-logo">🏥</div>
          <h1 class="login-title">FM<strong>Ops</strong></h1>
          <p class="login-subtitle">Facility Management Operations</p>
        </div>
        <form class="login-form" id="login-form" novalidate>
          <div class="form-group">
            <label class="form-label">Username / Email</label>
            <input type="text" name="username" class="form-control" placeholder="username atau email" required autofocus autocomplete="username">
          </div>
          <div class="form-group">
            <label class="form-label">Password</label>
            <div class="input-with-icon">
              <input type="password" name="password" class="form-control" placeholder="••••••••" required autocomplete="current-password" id="login-password">
              <button type="button" class="input-icon-btn" id="toggle-password" aria-label="Toggle password">👁</button>
            </div>
          </div>
          <div id="login-error" class="alert alert-danger" style="display:none"></div>
          <button type="submit" class="btn btn-primary btn-full btn-lg" id="login-btn">
            <span class="btn-text">Masuk</span>
            <span class="btn-spinner" style="display:none">⏳ Loading...</span>
          </button>
        </form>
        <div class="login-footer">
          <a href="#/form/chemical" class="link-subtle">📦 Form Permintaan Barang (Tanpa Login)</a>
        </div>
      </div>
    </div>
  `;

  const form = document.getElementById('login-form');
  const errorEl = document.getElementById('login-error');
  const loginBtn = document.getElementById('login-btn');
  const togglePwd = document.getElementById('toggle-password');
  const pwdInput = document.getElementById('login-password');

  togglePwd?.addEventListener('click', () => {
    pwdInput.type = pwdInput.type === 'password' ? 'text' : 'password';
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

    loginBtn.querySelector('.btn-text').style.display = 'none';
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
        toastSuccess('Login berhasil! Selamat datang.');
        window.dispatchEvent(new Event('fm:login'));
      } else {
        errorEl.textContent = res.data.error || 'Username atau password salah.';
        errorEl.style.display = 'block';
      }
    } catch (err) {
      errorEl.textContent = 'Gagal terhubung ke server. Periksa koneksi internet.';
      errorEl.style.display = 'block';
    } finally {
      loginBtn.querySelector('.btn-text').style.display = '';
      loginBtn.querySelector('.btn-spinner').style.display = 'none';
      loginBtn.disabled = false;
    }
  });
}
