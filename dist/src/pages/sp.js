import { apiFetch, showToast, formatDate, escapeHTML } from '../config.js';
export async function renderSP(container) {
  container.innerHTML = `
    <div class="page-header">
      <div class="page-title-group">
        <h2>Data Surat Peringatan (SP)</h2>
        <p class="text-secondary">Manajemen data SP karyawan</p>
      </div>
      <div class="page-actions">
        <button class="btn btn-primary" id="btn-add-sp">
          <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
          Tambah SP
        </button>
      </div>
    </div>

    <div class="card">
      <div class="table-container">
        <table class="table" id="sp-table">
          <thead>
            <tr>
              <th>Tanggal</th>
              <th>Nama Karyawan</th>
              <th>Cabang</th>
              <th>Jenis SP</th>
              <th>Status</th>
              <th>Dokumen</th>
              <th width="120">Aksi</th>
            </tr>
          </thead>
          <tbody>
            <tr><td colspan="7" class="text-center">Memuat data...</td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <!-- SP Modal -->
    <div class="modal-overlay" id="sp-modal">
      <div class="modal">
        <div class="modal-header">
          <h3 id="sp-modal-title">Tambah SP</h3>
          <button class="modal-close" id="sp-modal-close">&times;</button>
        </div>
        <div class="modal-body">
          <form id="sp-form">
            <input type="hidden" id="sp-id" name="id">
            <div class="form-group">
              <label class="form-label">Tanggal</label>
              <input type="date" class="form-control" name="tanggal" required>
            </div>
            <div class="form-group">
              <label class="form-label">Nama Karyawan</label>
              <input type="text" class="form-control" name="employee_name" required>
            </div>
            <div class="form-group">
              <label class="form-label">Cabang</label>
              <select class="form-control" name="branch_id" id="sp-branch-select" required>
                <option value="">Pilih Cabang</option>
              </select>
            </div>
            <div class="form-group">
              <label class="form-label">Jenis Surat Peringatan</label>
              <select class="form-control" name="sp_type" required>
                <option value="">Pilih Jenis</option>
                <option value="SP 1">SP 1</option>
                <option value="SP 2">SP 2</option>
                <option value="SP 3">SP 3</option>
                <option value="Teguran Lisan">Teguran Lisan</option>
              </select>
            </div>
            <div class="form-group">
              <label class="form-label">Status</label>
              <select class="form-control" name="status" required>
                <option value="Aktif">Aktif</option>
                <option value="Selesai">Selesai</option>
              </select>
            </div>
            <div class="form-group">
              <label class="form-label">Link Dokumen</label>
              <input type="url" class="form-control" name="document_link" placeholder="https://...">
            </div>
            <div class="form-actions">
              <button type="button" class="btn btn-secondary" id="btn-cancel-sp">Batal</button>
              <button type="submit" class="btn btn-primary" id="btn-save-sp">Simpan</button>
            </div>
          </form>
        </div>
      </div>
    </div>
  `;

  const tbody = container.querySelector('#sp-table tbody');
  const modal = container.querySelector('#sp-modal');
  const form = container.querySelector('#sp-form');
  const branchSelect = container.querySelector('#sp-branch-select');
  let spData = [];
  let branches = [];

  const openModal = () => modal.classList.add('show');
  const closeModal = () => { modal.classList.remove('show'); form.reset(); };

  container.querySelector('#btn-add-sp').addEventListener('click', () => {
    container.querySelector('#sp-modal-title').textContent = 'Tambah SP';
    form.reset();
    form.elements['id'].value = '';
    openModal();
  });
  container.querySelector('#sp-modal-close').addEventListener('click', closeModal);
  container.querySelector('#btn-cancel-sp').addEventListener('click', closeModal);

  async function loadBranches() {
    try {
      const res = await apiFetch('/api/branches');
      if (res.ok) {
        branches = res.data.data || [];
        branchSelect.innerHTML = '<option value="">Pilih Cabang</option>' + 
          branches.map(b => `<option value="${b.id}">${b.name}</option>`).join('');
      }
    } catch (e) {
      console.error(e);
    }
  }

  async function loadData() {
    try {
      const res = await apiFetch('/api/sp');
      if (res.ok) {
        spData = res.data.data || [];
        renderTable();
      }
    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="7" class="text-center text-danger">Gagal memuat data</td></tr>`;
    }
  }

  function renderTable() {
    if (spData.length === 0) {
      tbody.innerHTML = `<tr><td colspan="7" class="text-center">Tidak ada data SP.</td></tr>`;
      return;
    }
    tbody.innerHTML = spData.map(item => `
      <tr>
        <td>${formatDate(item.tanggal) || '-'}</td>
        <td class="fw-medium">${escapeHTML(item.employee_name)}</td>
        <td>${escapeHTML(item.branch_name || '-')}</td>
        <td><span class="badge badge-warning">${escapeHTML(item.sp_type)}</span></td>
        <td><span class="badge ${item.status === 'Aktif' ? 'badge-danger' : 'badge-success'}">${escapeHTML(item.status)}</span></td>
        <td>
          ${item.document_link ? `<a href="${item.document_link}" target="_blank" class="text-primary hover-underline">Lihat</a>` : '-'}
        </td>
        <td>
          <div class="action-buttons">
            <button class="btn-icon btn-edit" data-id="${item.id}" title="Edit"><svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 013 3L12 15l-4 1 1-4 9.5-9.5z"/></svg></button>
            <button class="btn-icon btn-delete text-danger" data-id="${item.id}" title="Hapus"><svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a2 2 0 012-2h4a2 2 0 012 2v2"/></svg></button>
          </div>
        </td>
      </tr>
    `).join('');

    tbody.querySelectorAll('.btn-edit').forEach(btn => {
      btn.addEventListener('click', () => {
        const item = spData.find(d => d.id == btn.dataset.id);
        if (item) {
          container.querySelector('#sp-modal-title').textContent = 'Edit SP';
          form.elements['id'].value = item.id;
          form.elements['tanggal'].value = item.tanggal || '';
          form.elements['employee_name'].value = item.employee_name || '';
          form.elements['branch_id'].value = item.branch_id || '';
          form.elements['sp_type'].value = item.sp_type || '';
          form.elements['status'].value = item.status || 'Aktif';
          form.elements['document_link'].value = item.document_link || '';
          openModal();
        }
      });
    });

    tbody.querySelectorAll('.btn-delete').forEach(btn => {
      btn.addEventListener('click', async () => {
        if (confirm('Yakin ingin menghapus data SP ini?')) {
          const res = await apiFetch(`/api/sp/${btn.dataset.id}`, { method: 'DELETE' });
          if (res.ok) { showToast('Data SP berhasil dihapus'); loadData(); }
          else { showToast('Gagal menghapus data', 'danger'); }
        }
      });
    });
  }

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    const btnSave = document.getElementById('btn-save-sp');
    btnSave.disabled = true;
    btnSave.textContent = 'Menyimpan...';

    const formData = new FormData(form);
    const payload = Object.fromEntries(formData.entries());
    const isEdit = !!payload.id;

    try {
      const res = await apiFetch(`/api/sp${isEdit ? '/' + payload.id : ''}`, {
        method: isEdit ? 'PUT' : 'POST',
        body: JSON.stringify(payload)
      });
      if (res.ok) {
        showToast(`Data SP berhasil ${isEdit ? 'diperbarui' : 'ditambahkan'}`);
        closeModal();
        loadData();
      } else {
        showToast(res.data?.message || 'Terjadi kesalahan', 'danger');
      }
    } catch (e) {
      showToast('Gagal menyimpan data', 'danger');
    } finally {
      btnSave.disabled = false;
      btnSave.textContent = 'Simpan';
    }
  });

  await loadBranches();
  await loadData();
};
