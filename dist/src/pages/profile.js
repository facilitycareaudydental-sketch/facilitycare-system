import { apiFetch, getUser, setUser } from '../config.js';
import { toastSuccess, toastError } from '../components/toast.js';

export async function renderProfile(container) {
  const user = getUser();
  const initial = (user?.full_name || user?.username || 'U')[0].toUpperCase();

  const roleColors = {
    superadmin: '#7C3AED', admin: '#2563EB', manager: '#0891B2',
    spv: '#059669', viewer: '#64748B',
  };
  const roleColor = roleColors[user?.role] || '#64748B';

  container.innerHTML = `
    <div class="page-header">
      <h1 class="page-title">👤 Profil Saya</h1>
    </div>

    <div class="profile-layout">

      <!-- LEFT: Info Card -->
      <div class="chart-card profile-info-card">
        <div class="profile-avatar-wrap">
          <div class="profile-avatar-xl" style="background:linear-gradient(135deg,${roleColor},${roleColor}99)">
            ${initial}
          </div>
          <div class="profile-name-block">
            <div class="profile-fullname">${user?.full_name || '—'}</div>
            <div class="profile-username">@${user?.username || '—'}</div>
            <span class="badge badge-info" style="background:${roleColor}18;color:${roleColor};margin-top:6px">
              ${(user?.role || 'viewer').toUpperCase()}
            </span>
          </div>
        </div>

        <hr class="profile-divider">

        <div class="info-list">
          <div class="info-row">
            <span class="info-key">📧 Email</span>
            <span class="info-value">${user?.email || '—'}</span>
          </div>
          <div class="info-row">
            <span class="info-key">👤 Username</span>
            <span class="info-value">${user?.username || '—'}</span>
          </div>
          <div class="info-row">
            <span class="info-key">🎯 Role</span>
            <span class="info-value" style="color:${roleColor};font-weight:700">${user?.role || '—'}</span>
          </div>
        </div>
      </div>

      <!-- RIGHT: Change Password -->
      <div class="chart-card">
        <div class="chart-card-header">
          <div>
            <div class="chart-card-title">🔑 Ganti Password</div>
            <div class="chart-card-subtitle">Gunakan password yang kuat, minimal 6 karakter</div>
          </div>
        </div>

        <form id="change-pwd-form" novalidate style="margin-top:8px">
          <div class="form-group">
            <label class="form-label">Password Lama <span class="required">*</span></label>
            <input type="password" name="current_password" class="form-control"
              required placeholder="Masukkan password saat ini" autocomplete="current-password">
          </div>
          <div class="form-group">
            <label class="form-label">Password Baru <span class="required">*</span></label>
            <input type="password" name="new_password" class="form-control"
              required placeholder="Minimal 6 karakter" autocomplete="new-password">
          </div>
          <div class="form-group">
            <label class="form-label">Konfirmasi Password Baru <span class="required">*</span></label>
            <input type="password" name="confirm_password" class="form-control"
              required placeholder="Ulangi password baru" autocomplete="new-password">
          </div>

          <div id="pwd-error" class="alert alert-danger" style="display:none"></div>
          <div id="pwd-success" class="alert alert-success" style="display:none"></div>

          <button type="submit" class="btn btn-primary" id="btn-save-pwd">
            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px">
              <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/>
              <polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/>
            </svg>
            Simpan Password
          </button>
        </form>

        <hr class="profile-divider" style="margin-top:28px">

        <div class="chart-card-title" style="margin-bottom:12px">🔐 Keamanan Akun</div>
        <div class="info-list">
          <div class="info-row">
            <span class="info-key">Token Login</span>
            <span class="info-value">
              <span class="badge badge-success">Aktif</span>
            </span>
          </div>
          <div class="info-row">
            <span class="info-key">Session</span>
            <span class="info-value" id="session-info">Memuat...</span>
          </div>
        </div>
        <button class="btn btn-danger btn-sm" id="btn-logout" style="margin-top:16px">
          <svg width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:4px">
            <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/>
            <polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/>
          </svg>
          Keluar dari Semua Sesi
        </button>
      </div>

    </div>
  `;

  // Session info
  const token = localStorage.getItem('fm_token');
  const sessEl = document.getElementById('session-info');
  if (token && sessEl) {
    try {
      const payload = JSON.parse(atob(token.split('.')[1]));
      const exp = new Date(payload.exp * 1000);
      sessEl.textContent = `Berakhir: ${exp.toLocaleString('id-ID')}`;
    } catch { sessEl.textContent = 'Tidak tersedia'; }
  }

  // Change password
  document.getElementById('change-pwd-form')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const errEl  = document.getElementById('pwd-error');
    const sucEl  = document.getElementById('pwd-success');
    const btn    = document.getElementById('btn-save-pwd');
    errEl.style.display = 'none';
    sucEl.style.display = 'none';

    const form            = e.target;
    const current_password = form.current_password.value;
    const new_password     = form.new_password.value;
    const confirm_password = form.confirm_password.value;

    if (new_password !== confirm_password) {
      errEl.textContent = '❌ Konfirmasi password tidak cocok.';
      errEl.style.display = 'block'; return;
    }
    if (new_password.length < 6) {
      errEl.textContent = '❌ Password baru minimal 6 karakter.';
      errEl.style.display = 'block'; return;
    }

    btn.disabled = true;
    btn.textContent = '⏳ Menyimpan...';

    const res = await apiFetch('/api/auth/change-password', {
      method: 'POST',
      body: JSON.stringify({ current_password, new_password }),
    });

    btn.disabled = false;
    btn.innerHTML = `<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24" style="margin-right:5px"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></svg>Simpan Password`;

    if (res.ok) {
      sucEl.textContent = '✅ Password berhasil diubah.';
      sucEl.style.display = 'block';
      form.reset();
      toastSuccess('Password berhasil diubah.');
    } else {
      errEl.textContent = res.data?.error || 'Gagal mengubah password.';
      errEl.style.display = 'block';
    }
  });

  // Logout
  document.getElementById('btn-logout')?.addEventListener('click', () => {
    if (confirm('Keluar dari semua sesi? Anda harus login ulang.')) {
      localStorage.clear();
      window.location.reload();
    }
  });
}
