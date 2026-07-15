import { apiFetch, getUser, setUser } from '../config.js';
import { toastSuccess, toastError } from '../components/toast.js';

export async function renderProfile(container) {
  const user = getUser();

  container.innerHTML = `
    <div class="page-header">
      <h1 class="page-title">👤 Profil Saya</h1>
    </div>
    <div class="profile-grid">
      <div class="card">
        <div class="card-header"><h3 class="card-title">Informasi Akun</h3></div>
        <div class="card-body">
          <div class="profile-avatar-lg">${(user?.full_name || 'U')[0].toUpperCase()}</div>
          <div class="profile-info">
            <div class="profile-row"><span class="profile-label">Nama Lengkap</span><span>${user?.full_name || '-'}</span></div>
            <div class="profile-row"><span class="profile-label">Username</span><span>${user?.username || '-'}</span></div>
            <div class="profile-row"><span class="profile-label">Email</span><span>${user?.email || '-'}</span></div>
            <div class="profile-row"><span class="profile-label">Role</span><span class="badge badge-info">${user?.role || '-'}</span></div>
          </div>
        </div>
      </div>

      <div class="card">
        <div class="card-header"><h3 class="card-title">Ganti Password</h3></div>
        <div class="card-body">
          <form id="change-pwd-form" novalidate>
            <div class="form-group">
              <label class="form-label">Password Lama <span class="required">*</span></label>
              <input type="password" name="current_password" class="form-control" required placeholder="Password saat ini">
            </div>
            <div class="form-group">
              <label class="form-label">Password Baru <span class="required">*</span></label>
              <input type="password" name="new_password" class="form-control" required placeholder="Min. 6 karakter">
            </div>
            <div class="form-group">
              <label class="form-label">Konfirmasi Password Baru <span class="required">*</span></label>
              <input type="password" name="confirm_password" class="form-control" required placeholder="Ulangi password baru">
            </div>
            <div id="pwd-error" class="alert alert-danger" style="display:none"></div>
            <button type="submit" class="btn btn-primary">Simpan Password</button>
          </form>
        </div>
      </div>
    </div>
  `;

  document.getElementById('change-pwd-form')?.addEventListener('submit', async (e) => {
    e.preventDefault();
    const errEl = document.getElementById('pwd-error');
    errEl.style.display = 'none';
    const form = e.target;
    const current_password = form.current_password.value;
    const new_password = form.new_password.value;
    const confirm_password = form.confirm_password.value;

    if (new_password !== confirm_password) {
      errEl.textContent = 'Konfirmasi password tidak cocok.';
      errEl.style.display = 'block';
      return;
    }
    if (new_password.length < 6) {
      errEl.textContent = 'Password baru minimal 6 karakter.';
      errEl.style.display = 'block';
      return;
    }

    const res = await apiFetch('/api/auth/change-password', {
      method: 'POST',
      body: JSON.stringify({ current_password, new_password }),
    });
    if (res.ok) {
      toastSuccess('Password berhasil diubah.');
      form.reset();
    } else {
      errEl.textContent = res.data?.error || 'Gagal mengubah password.';
      errEl.style.display = 'block';
    }
  });
}
