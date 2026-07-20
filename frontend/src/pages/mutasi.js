import { apiFetch } from '../config.js';
import { toast, toastSuccess, toastError } from '../components/toast.js';

export async function renderMutasi(container) {
  function formatDate(d) {
    if (!d) return '';
    return new Date(d).toLocaleDateString('id-ID', { year: 'numeric', month: 'short', day: 'numeric' });
  }
  function escapeHTML(str) {
    if (!str) return '';
    return str.toString().replace(/[&<>'"]/g, 
      tag => ({
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        "'": '&#39;',
        '"': '&quot;'
      }[tag])
    );
  }

  container.innerHTML = `
    <div class="page-header">
      <div class="page-title-group">
        <h2>Data Mutasi</h2>
        <p class="text-secondary">Manajemen rotasi dan mutasi karyawan</p>
      </div>
      <div class="page-actions">
        <button class="btn btn-primary" id="btn-add-mutasi">
          <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
          Tambah Mutasi
        </button>
      </div>
    </div>

    <div class="card">
      <div class="table-container">
        <table class="table" id="mutasi-table">
          <thead>
            <tr>
              <th>Tanggal</th>
              <th>Nama Karyawan</th>
              <th>Awal Cabang</th>
              <th>Rotasi Cabang</th>
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

    <!-- Mutasi Modal -->
    <div class="modal-overlay" id="mutasi-modal">
      <div class="modal">
        <div class="modal-header">
          <h3 id="mutasi-modal-title">Tambah Mutasi</h3>
          <button class="modal-close" id="mutasi-modal-close">&times;</button>
        </div>
        <div class="modal-body">
          <form id="mutasi-form">
            <input type="hidden" id="mutasi-id" name="id">
            <div class="form-group">
              <label class="form-label">Tanggal</label>
              <input type="date" class="form-control" name="tanggal" required>
            </div>
            <div class="form-group">
              <label class="form-label">Nama Karyawan</label>
              <input type="text" class="form-control" name="employee_name" required>
            </div>
            <div class="form-group">
              <label class="form-label">Awal Cabang</label>
              <select class="form-control branch-select" name="from_branch_id" required>
                <option value="">Pilih Cabang</option>
              </select>
            </div>
            <div class="form-group">
              <label class="form-label">Rotasi Cabang</label>
              <select class="form-control branch-select" name="to_branch_id" required>
                <option value="">Pilih Cabang</option>
              </select>
            </div>
            <div class="form-group">
              <label class="form-label">Status</label>
              <select class="form-control" name="status" required>
                <option value="Selesai">Selesai</option>
                <option value="Proses">Proses</option>
              </select>
            </div>
            <div class="form-group">
              <label class="form-label">Link Dokumen</label>
              <input type="url" class="form-control" name="document_link" placeholder="https://...">
            </div>
            <div class="form-actions">
              <button type="button" class="btn btn-secondary" id="btn-cancel-mutasi">Batal</button>
              <button type="submit" class="btn btn-primary" id="btn-save-mutasi">Simpan</button>
            </div>
          </form>
        </div>
      </div>
    </div>
  `;

  const tbody = container.querySelector('#mutasi-table tbody');
  const modal = container.querySelector('#mutasi-modal');
  const form = container.querySelector('#mutasi-form');
  const branchSelects = container.querySelectorAll('.branch-select');
  let mutasiData = [];
  let branches = [];

  const openModal = () => modal.classList.add('show');
  const closeModal = () => { modal.classList.remove('show'); form.reset(); };

  container.querySelector('#btn-add-mutasi').addEventListener('click', () => {
    container.querySelector('#mutasi-modal-title').textContent = 'Tambah Mutasi';
    form.reset();
    form.elements['id'].value = '';
    openModal();
  });
  container.querySelector('#mutasi-modal-close').addEventListener('click', closeModal);
  container.querySelector('#btn-cancel-mutasi').addEventListener('click', closeModal);

  async function loadBranches() {
    try {
      const res = await apiFetch('/api/branches');
      if (res.ok) {
        branches = res.data.data || [];
        const opts = '<option value="">Pilih Cabang</option>' + branches.map(b => `<option value="${b.id}">${b.name}</option>`).join('');
        branchSelects.forEach(sel => sel.innerHTML = opts);
      }
    } catch (e) {
      console.error(e);
    }
  }

  async function loadData() {
    try {
      const res = await apiFetch('/api/mutasi');
      if (res.ok) {
        mutasiData = res.data.data || [];
        renderTable();
      }
    } catch (e) {
      tbody.innerHTML = `<tr><td colspan="7" class="text-center text-danger">Gagal memuat data</td></tr>`;
    }
  }

  function renderTable() {
    if (mutasiData.length === 0) {
      tbody.innerHTML = `<tr><td colspan="7" class="text-center">Tidak ada data Mutasi.</td></tr>`;
      return;
    }
    tbody.innerHTML = mutasiData.map(item => `
      <tr>
        <td>${formatDate(item.tanggal) || '-'}</td>
        <td class="fw-medium">${escapeHTML(item.employee_name)}</td>
        <td>${escapeHTML(item.from_branch_name || '-')}</td>
        <td>${escapeHTML(item.to_branch_name || '-')}</td>
        <td><span class="badge ${item.status === 'Selesai' ? 'badge-success' : 'badge-warning'}">${escapeHTML(item.status)}</span></td>
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
        const item = mutasiData.find(d => d.id == btn.dataset.id);
        if (item) {
          container.querySelector('#mutasi-modal-title').textContent = 'Edit Mutasi';
          form.elements['id'].value = item.id;
          form.elements['tanggal'].value = item.tanggal || '';
          form.elements['employee_name'].value = item.employee_name || '';
          form.elements['from_branch_id'].value = item.from_branch_id || '';
          form.elements['to_branch_id'].value = item.to_branch_id || '';
          form.elements['status'].value = item.status || 'Selesai';
          form.elements['document_link'].value = item.document_link || '';
          openModal();
        }
      });
    });

    tbody.querySelectorAll('.btn-delete').forEach(btn => {
      btn.addEventListener('click', async () => {
        if (confirm('Yakin ingin menghapus data mutasi ini?')) {
          const res = await apiFetch(`/api/mutasi/${btn.dataset.id}`, { method: 'DELETE' });
          if (res.ok) { toastSuccess('Data Mutasi berhasil dihapus'); loadData(); }
          else { toastError('Gagal menghapus data'); }
        }
      });
    });
  }

  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    const btnSave = document.getElementById('btn-save-mutasi');
    btnSave.disabled = true;
    btnSave.textContent = 'Menyimpan...';

    const formData = new FormData(form);
    const payload = Object.fromEntries(formData.entries());
    const isEdit = !!payload.id;

    try {
      const res = await apiFetch(`/api/mutasi${isEdit ? '/' + payload.id : ''}`, {
        method: isEdit ? 'PUT' : 'POST',
        body: JSON.stringify(payload)
      });
      if (res.ok) {
        toastSuccess(`Data Mutasi berhasil ${isEdit ? 'diperbarui' : 'ditambahkan'}`);
        closeModal();
        loadData();
      } else {
        toastError(res.data?.message || 'Terjadi kesalahan');
      }
    } catch (e) {
      toastError('Gagal menyimpan data');
    } finally {
      btnSave.disabled = false;
      btnSave.textContent = 'Simpan';
    }
  });

  await loadBranches();
  await loadData();
}
