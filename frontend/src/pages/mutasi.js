import { apiFetch, getUser } from '../config.js';
import { getCachedBranches, getCachedEmployeeNames } from '../utils/dataCache.js';
import { createModal, confirmDialog } from '../components/modal.js';
import { toastSuccess, toastError, toastInfo } from '../components/toast.js';

// Pre-seeded fallback data jika API offline / inisialisasi awal
const INITIAL_MUTASI = [
  { id: 1, tanggal: '2026-10-02', employee_name: 'TAUFIK HIDAYAT', from_branch_id: 12, from_branch_name: '012. BSD', to_branch_id: 0, to_branch_name: '000. Management', status: 'Selesai', document_link: '' },
  { id: 2, tanggal: '2026-09-30', employee_name: 'ELVAN JANGGA PRAMANA', from_branch_id: 55, from_branch_name: '055. Gajah Mada', to_branch_id: 0, to_branch_name: '000. Management', status: 'Proses', document_link: '' },
  { id: 3, tanggal: '2026-06-23', employee_name: 'Widya Astuti', from_branch_id: 1, from_branch_name: '001. Pondok Bambu', to_branch_id: 7, to_branch_name: '007. Bekasi', status: 'Selesai', document_link: '' },
  { id: 4, tanggal: '2026-06-22', employee_name: 'Gema Azan Saputra', from_branch_id: 1, from_branch_name: '001. Pondok Bambu', to_branch_id: 21, to_branch_name: '021. Kelapa Gading Boulevard', status: 'Selesai', document_link: '' },
  { id: 5, tanggal: '2026-06-22', employee_name: 'Ahmad Farhan', from_branch_id: 1, from_branch_name: '001. Pondok Bambu', to_branch_id: 62, to_branch_name: '062. Summarecon Bekasi', status: 'Selesai', document_link: '' }
];

// Helper format tanggal lokal Indonesia
function formatDateIndo(dateStr) {
  if (!dateStr) return '-';
  try {
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    return d.toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' });
  } catch {
    return dateStr;
  }
}

export async function renderMutasi(container) {
  let mutasiList = [];
  let branchOptions = [];
  let employeeOptions = [];
  let selectedStatus = 'all'; // 'all' | 'Proses' | 'Selesai'
  let selectedFromBranch = 'all';
  let selectedToBranch = 'all';
  let searchQuery = '';
  let isLoading = false;

  const currentUser = getUser();
  const canEdit = currentUser?.role !== 'viewer';

  // Load branches and employees
  try {
    [branchOptions, employeeOptions] = await Promise.all([
      getCachedBranches(),
      getCachedEmployeeNames()
    ]);
  } catch (err) {
    console.error('Error fetching cache:', err);
  }

  // Load Mutasi Data from API
  async function loadData() {
    isLoading = true;
    renderUI();
    try {
      const res = await apiFetch('/api/mutasi?limit=10000');
      if (res.ok) {
        const raw = res.data?.data || res.data || [];
        mutasiList = Array.isArray(raw) && raw.length > 0 ? raw : INITIAL_MUTASI;
      } else {
        mutasiList = INITIAL_MUTASI;
      }
    } catch (e) {
      console.warn('Gagal load mutasi, menggunakan data lokal:', e);
      mutasiList = INITIAL_MUTASI;
    } finally {
      isLoading = false;
      renderUI();
    }
  }

  // Export to Excel
  async function exportToExcel() {
    try {
      const { downloadExcel } = await import('../utils/excel.js');
      const exportData = mutasiList.map((m, idx) => ({
        'No': idx + 1,
        'Tanggal': m.tanggal || '',
        'Nama Karyawan': m.employee_name || '',
        'Cabang Asal': m.from_branch_name || '',
        'Cabang Tujuan': m.to_branch_name || '',
        'Status': m.status || '',
        'Link Dokumen': m.document_link || ''
      }));
      downloadExcel(exportData, `Data_Mutasi_${new Date().toISOString().slice(0, 10)}`);
      toastSuccess('Data mutasi berhasil diekspor ke Excel!');
    } catch (err) {
      toastError('Gagal ekspor Excel: ' + err.message);
    }
  }

  // Delete Mutasi
  async function deleteMutasi(id) {
    if (!canEdit) return;
    const item = mutasiList.find(m => m.id === id);
    confirmDialog({
      title: 'Hapus Data Mutasi',
      message: `Apakah Anda yakin ingin menghapus data mutasi untuk "${item?.employee_name || 'karyawan ini'}"?`,
      confirmText: 'Ya, Hapus',
      confirmClass: 'btn-danger',
      onConfirm: async () => {
        try {
          const res = await apiFetch(`/api/mutasi/${id}`, { method: 'DELETE' });
          if (res.ok) {
            toastSuccess('Data mutasi berhasil dihapus');
            mutasiList = mutasiList.filter(m => m.id !== id);
            renderUI();
          } else {
            toastError(res.data?.error || 'Gagal menghapus mutasi');
          }
        } catch (e) {
          toastError('Gagal menghapus data mutasi: ' + e.message);
        }
      }
    });
  }

  // Open Form Modal (Add / Edit)
  function openMutasiModal(item = null) {
    if (!canEdit) return;
    const isEdit = !!item;
    const defaultTanggal = item?.tanggal ? item.tanggal.slice(0, 10) : new Date().toISOString().slice(0, 10);

    const content = document.createElement('div');
    content.innerHTML = `
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Mutasi <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="mutasi-modal-tanggal" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${defaultTanggal}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Mutasi <span style="color:#DC2626;">*</span>
            </label>
            <select id="mutasi-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="Proses" ${item?.status === 'Proses' ? 'selected' : ''}>⏳ Proses</option>
              <option value="Selesai" ${item?.status === 'Selesai' || !item ? 'selected' : ''}>✅ Selesai</option>
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Nama Karyawan <span style="color:#DC2626;">*</span>
          </label>
          <input list="emp-list-options" id="mutasi-modal-emp" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Ketik atau pilih nama karyawan..." value="${item?.employee_name ? String(item.employee_name).replace(/"/g, '&quot;') : ''}" required />
          <datalist id="emp-list-options">
            ${employeeOptions.map(e => {
              const name = typeof e === 'string' ? e : (e.label || e.name || '');
              return `<option value="${name}">`;
            }).join('')}
          </datalist>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Asal <span style="color:#DC2626;">*</span>
            </label>
            <select id="mutasi-modal-from" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang Asal --</option>
              ${branchOptions.map(b => `
                <option value="${b.value}" ${String(item?.from_branch_id) === String(b.value) ? 'selected' : ''}>
                  ${b.label}
                </option>
              `).join('')}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Tujuan <span style="color:#DC2626;">*</span>
            </label>
            <select id="mutasi-modal-to" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang Tujuan --</option>
              ${branchOptions.map(b => `
                <option value="${b.value}" ${String(item?.to_branch_id) === String(b.value) ? 'selected' : ''}>
                  ${b.label}
                </option>
              `).join('')}
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Link Berkas SK Mutasi / Dokumen (Opsional)
          </label>
          <input type="url" id="mutasi-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="https://drive.google.com/file/d/..." value="${item?.document_link ? String(item.document_link).replace(/"/g, '&quot;') : ''}" />
        </div>
      </div>
    `;

    createModal({
      title: isEdit ? '✏️ Edit Data Mutasi Karyawan' : '➕ Tambah Data Mutasi Karyawan',
      content,
      size: 'md',
      confirmText: isEdit ? 'Simpan Perubahan' : 'Simpan Mutasi',
      confirmClass: 'btn-primary',
      onConfirm: async (_, close) => {
        const tanggal = content.querySelector('#mutasi-modal-tanggal').value;
        const employee_name = (content.querySelector('#mutasi-modal-emp').value || '').trim();
        const from_branch_id = content.querySelector('#mutasi-modal-from').value;
        const to_branch_id = content.querySelector('#mutasi-modal-to').value;
        const status = content.querySelector('#mutasi-modal-status').value;
        const document_link = (content.querySelector('#mutasi-modal-link').value || '').trim();

        if (!tanggal) {
          toastError('Tanggal mutasi wajib diisi');
          return;
        }
        if (!employee_name) {
          toastError('Nama karyawan wajib diisi');
          return;
        }
        if (!from_branch_id) {
          toastError('Cabang asal wajib dipilih');
          return;
        }
        if (!to_branch_id) {
          toastError('Cabang tujuan wajib dipilih');
          return;
        }

        const bFrom = branchOptions.find(b => String(b.value) === String(from_branch_id));
        const bTo = branchOptions.find(b => String(b.value) === String(to_branch_id));

        const payload = {
          tanggal,
          employee_name,
          from_branch_id: Number(from_branch_id),
          from_branch_name: bFrom ? bFrom.label : 'Cabang Asal',
          to_branch_id: Number(to_branch_id),
          to_branch_name: bTo ? bTo.label : 'Cabang Tujuan',
          status,
          document_link
        };

        try {
          if (isEdit) {
            const res = await apiFetch(`/api/mutasi/${item.id}`, {
              method: 'PUT',
              body: JSON.stringify(payload)
            });
            if (res.ok) {
              toastSuccess('Data mutasi berhasil diperbarui');
              const idx = mutasiList.findIndex(m => m.id === item.id);
              if (idx !== -1) mutasiList[idx] = { ...mutasiList[idx], ...payload };
              close();
              renderUI();
            } else {
              toastError(res.data?.error || 'Gagal memperbarui data mutasi');
            }
          } else {
            const res = await apiFetch('/api/mutasi', {
              method: 'POST',
              body: JSON.stringify(payload)
            });
            if (res.ok) {
              toastSuccess('Data mutasi baru berhasil ditambahkan');
              const newId = res.data?.id || (Date.now());
              mutasiList.unshift({ id: newId, ...payload });
              close();
              renderUI();
            } else {
              toastError(res.data?.error || 'Gagal menambahkan mutasi');
            }
          }
        } catch (err) {
          toastError('Terjadi kesalahan: ' + err.message);
        }
      }
    });
  }

  // Filter & Search Logic
  function getFilteredList() {
    return mutasiList.filter(m => {
      // Search
      if (searchQuery) {
        const q = searchQuery.toLowerCase();
        const matchEmp = (m.employee_name || '').toLowerCase().includes(q);
        const matchFrom = (m.from_branch_name || '').toLowerCase().includes(q);
        const matchTo = (m.to_branch_name || '').toLowerCase().includes(q);
        if (!matchEmp && !matchFrom && !matchTo) return false;
      }

      // From Branch Filter
      if (selectedFromBranch !== 'all' && String(m.from_branch_id) !== String(selectedFromBranch)) {
        return false;
      }

      // To Branch Filter
      if (selectedToBranch !== 'all' && String(m.to_branch_id) !== String(selectedToBranch)) {
        return false;
      }

      // Status Filter
      if (selectedStatus !== 'all' && m.status !== selectedStatus) {
        return false;
      }

      return true;
    });
  }

  // Main UI Renderer
  function renderUI() {
    const filtered = getFilteredList();

    // KPI Metrics calculation
    const totalCount = mutasiList.length;
    const prosesCount = mutasiList.filter(m => m.status === 'Proses').length;
    const selesaiCount = mutasiList.filter(m => m.status === 'Selesai').length;

    // Mutasi Bulan Ini
    const now = new Date();
    const curYear = now.getFullYear();
    const curMonth = now.getMonth();
    const bulanIniCount = mutasiList.filter(m => {
      if (!m.tanggal) return false;
      const d = new Date(m.tanggal);
      return !isNaN(d.getTime()) && d.getFullYear() === curYear && d.getMonth() === curMonth;
    }).length;

    container.innerHTML = `
      <style>
        .mutasi-page {
          display: flex;
          flex-direction: column;
          gap: 20px;
        }
        .mutasi-hero-card {
          background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
          border-radius: 16px;
          padding: 24px 28px;
          color: white;
          box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25);
          position: relative;
          overflow: hidden;
        }
        .mutasi-hero-card::after {
          content: '';
          position: absolute;
          top: -40px;
          right: -40px;
          width: 220px;
          height: 220px;
          background: radial-gradient(circle, rgba(59, 130, 246, 0.15) 0%, transparent 70%);
          border-radius: 50%;
          pointer-events: none;
        }
        .mutasi-stats-row {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 14px;
          margin-top: 22px;
        }
        .mutasi-stat-box {
          background: rgba(255, 255, 255, 0.06);
          border: 1px solid rgba(255, 255, 255, 0.1);
          border-radius: 12px;
          padding: 14px 16px;
          backdrop-filter: blur(8px);
          display: flex;
          align-items: center;
          gap: 14px;
        }
        .mutasi-stat-icon {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 1.3rem;
          flex-shrink: 0;
        }
        .mutasi-stat-val {
          font-size: 1.55rem;
          font-weight: 800;
          line-height: 1.1;
          color: #F8FAFC;
        }
        .mutasi-stat-lbl {
          font-size: 0.76rem;
          text-transform: uppercase;
          letter-spacing: 0.05em;
          color: #94A3B8;
          font-weight: 600;
          margin-top: 2px;
        }

        /* Controls Card */
        .mutasi-controls-card {
          background: white;
          border: 1px solid #E2E8F0;
          border-radius: 14px;
          padding: 18px 20px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.02);
          display: flex;
          flex-direction: column;
          gap: 16px;
        }
        .mutasi-search-row {
          display: flex;
          gap: 12px;
          flex-wrap: wrap;
        }
        .mutasi-search-wrap {
          position: relative;
          flex: 1;
          min-width: 240px;
        }
        .mutasi-search-icon {
          position: absolute;
          left: 14px;
          top: 50%;
          transform: translateY(-50%);
          color: #94A3B8;
          pointer-events: none;
        }
        .mutasi-search-input {
          width: 100%;
          padding: 10px 14px 10px 40px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          outline: none;
          transition: border-color 0.15s ease;
        }
        .mutasi-search-input:focus {
          border-color: #2563EB;
          box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);
        }
        .mutasi-select {
          padding: 10px 14px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          color: #1E293B;
          background: white;
          outline: none;
          min-width: 180px;
          cursor: pointer;
        }
        .mutasi-select:focus {
          border-color: #2563EB;
        }

        /* Swipeable Pills */
        .mutasi-filter-pills {
          display: flex;
          align-items: center;
          gap: 8px;
          overflow-x: auto;
          white-space: nowrap;
          scrollbar-width: none;
          padding-bottom: 2px;
        }
        .mutasi-filter-pills::-webkit-scrollbar {
          display: none;
        }
        .mutasi-pill {
          padding: 6px 14px;
          border-radius: 20px;
          font-size: 0.82rem;
          font-weight: 600;
          cursor: pointer;
          border: 1px solid #E2E8F0;
          background: #F8FAFC;
          color: #64748B;
          transition: all 0.15s ease;
          display: inline-flex;
          align-items: center;
          gap: 6px;
          user-select: none;
          flex-shrink: 0;
        }
        .mutasi-pill:hover {
          background: #F1F5F9;
          color: #1E293B;
        }
        .mutasi-pill.active {
          background: #2563EB;
          color: white;
          border-color: #2563EB;
          box-shadow: 0 2px 6px rgba(37, 99, 235, 0.25);
        }

        /* Grid Cards View */
        .mutasi-cards-grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
          gap: 16px;
        }
        .mutasi-card {
          background: white;
          border: 1px solid #E2E8F0;
          border-radius: 14px;
          padding: 18px;
          box-shadow: 0 2px 6px rgba(0,0,0,0.03);
          transition: transform 0.15s ease, box-shadow 0.15s ease;
          display: flex;
          flex-direction: column;
          justify-content: space-between;
          position: relative;
        }
        .mutasi-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 8px 18px rgba(0,0,0,0.06);
          border-color: #CBD5E1;
        }
        .mutasi-avatar-box {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
          color: white;
          font-weight: 700;
          font-size: 1rem;
          display: flex;
          align-items: center;
          justify-content: center;
          flex-shrink: 0;
          box-shadow: 0 3px 6px rgba(37, 99, 235, 0.2);
        }
        .mutasi-route-box {
          background: #F8FAFC;
          border: 1px dashed #CBD5E1;
          border-radius: 10px;
          padding: 12px 14px;
          margin: 12px 0;
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 10px;
        }
        .mutasi-route-node {
          flex: 1;
        }
        .mutasi-route-label {
          font-size: 0.7rem;
          text-transform: uppercase;
          font-weight: 700;
          letter-spacing: 0.04em;
          color: #94A3B8;
          margin-bottom: 2px;
        }
        .mutasi-route-branch {
          font-size: 0.85rem;
          font-weight: 700;
          color: #1E293B;
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
        }
        .mutasi-route-arrow {
          display: flex;
          align-items: center;
          justify-content: center;
          width: 32px;
          height: 32px;
          border-radius: 50%;
          background: #EFF6FF;
          color: #2563EB;
          font-size: 1rem;
          flex-shrink: 0;
        }
        .mutasi-status-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.74rem;
          font-weight: 700;
        }
        .mutasi-status-badge.selesai {
          background: #DCFCE7;
          color: #15803D;
          border: 1px solid #BBF7D0;
        }
        .mutasi-status-badge.proses {
          background: #FEF3C7;
          color: #B45309;
          border: 1px solid #FDE68A;
        }
        .mutasi-icon-btn {
          width: 32px;
          height: 32px;
          border-radius: 8px;
          border: 1px solid #E2E8F0;
          background: white;
          color: #64748B;
          display: inline-flex;
          align-items: center;
          justify-content: center;
          cursor: pointer;
          transition: all 0.15s ease;
        }
        .mutasi-icon-btn:hover {
          background: #F1F5F9;
          color: #0F172A;
          border-color: #CBD5E1;
        }
        .mutasi-icon-btn.delete:hover {
          background: #FEE2E2;
          color: #DC2626;
          border-color: #FCA5A5;
        }

        /* Empty state */
        .mutasi-empty-state {
          grid-column: 1 / -1;
          background: white;
          border: 2px dashed #E2E8F0;
          border-radius: 16px;
          padding: 48px 24px;
          text-align: center;
        }

        /* Mobile Responsive */
        @media (max-width: 768px) {
          .mutasi-hero-card {
            padding: 18px 16px;
          }
          .mutasi-stats-row {
            grid-template-columns: 1fr 1fr;
            gap: 10px;
          }
          .mutasi-stat-box {
            padding: 10px 12px;
          }
          .mutasi-stat-icon {
            width: 38px;
            height: 38px;
            font-size: 1.1rem;
          }
          .mutasi-stat-val {
            font-size: 1.3rem;
          }
          .mutasi-controls-card {
            padding: 14px;
            gap: 12px;
          }
          .mutasi-search-row {
            flex-direction: column;
            gap: 10px;
          }
          .mutasi-select {
            width: 100%;
          }
          .mutasi-cards-grid {
            grid-template-columns: 1fr;
          }
        }
      </style>

      <div class="mutasi-page">

        <!-- Hero Header Card -->
        <div class="mutasi-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px; margin-bottom:6px;">
                <span style="font-size:1.6rem;">🔁</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Monitoring Data Mutasi Karyawan
                </h1>
              </div>
              <p style="margin:0; font-size:0.88rem; color:#CBD5E1; max-width:650px; line-height:1.5;">
                Sistem pengelolaan rotasi penempatan tugas, arsip SK mutasi cabang, dan riwayat perpindahan personil PT Audy Mandiri Indonesia.
              </p>
            </div>

            <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
              ${canEdit ? `
                <button id="mutasi-btn-add" class="btn btn-primary" style="background:#2563EB; padding:9px 18px; border-radius:10px; font-weight:700; display:flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3);">
                  <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Mutasi</span>
                </button>
              ` : ''}
              
              <button id="mutasi-btn-excel" class="btn btn-ghost" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; padding:9px 16px; border-radius:10px; font-weight:600; display:flex; align-items:center; gap:6px;">
                <span>📥</span>
                <span>Ekspor Excel</span>
              </button>

              <button id="mutasi-btn-refresh" class="btn btn-ghost" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; padding:9px; border-radius:10px;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M23 4v6h-6M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- KPI Stats -->
          <div class="mutasi-stats-row">
            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">
                📋
              </div>
              <div>
                <div class="mutasi-stat-val">${totalCount}</div>
                <div class="mutasi-stat-lbl">Total Mutasi</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                ⏳
              </div>
              <div>
                <div class="mutasi-stat-val">${prosesCount}</div>
                <div class="mutasi-stat-lbl">Dalam Proses</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(34,197,94,0.2); color:#4ADE80;">
                ✅
              </div>
              <div>
                <div class="mutasi-stat-val">${selesaiCount}</div>
                <div class="mutasi-stat-lbl">Mutasi Selesai</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(168,85,247,0.2); color:#C084FC;">
                🗓️
              </div>
              <div>
                <div class="mutasi-stat-val">${bulanIniCount}</div>
                <div class="mutasi-stat-lbl">Bulan Berjalan</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Filter & Search Controls -->
        <div class="mutasi-controls-card">
          <div class="mutasi-search-row">
            <div class="mutasi-search-wrap">
              <span class="mutasi-search-icon">
                <svg width="17" height="17" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
              </span>
              <input type="text" id="mutasi-search-input" class="mutasi-search-input" placeholder="Cari nama karyawan, cabang asal, atau cabang tujuan..." value="${searchQuery}" />
            </div>

            <select id="mutasi-filter-from" class="mutasi-select">
              <option value="all">Semua Cabang Asal (${branchOptions.length})</option>
              ${branchOptions.map(b => `<option value="${b.value}" ${String(selectedFromBranch) === String(b.value) ? 'selected' : ''}>${b.label}</option>`).join('')}
            </select>

            <select id="mutasi-filter-to" class="mutasi-select">
              <option value="all">Semua Cabang Tujuan (${branchOptions.length})</option>
              ${branchOptions.map(b => `<option value="${b.value}" ${String(selectedToBranch) === String(b.value) ? 'selected' : ''}>${b.label}</option>`).join('')}
            </select>
          </div>

          <!-- Status Filter Pills (Swipeable on Mobile) -->
          <div class="mutasi-filter-pills">
            <div class="mutasi-pill ${selectedStatus === 'all' ? 'active' : ''}" data-status="all">
              <span>Semua Mutasi</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${totalCount}</span>
            </div>
            <div class="mutasi-pill ${selectedStatus === 'Proses' ? 'active' : ''}" data-status="Proses">
              <span style="color:#D97706;">⏳</span>
              <span>Proses</span>
              <span class="badge" style="background:rgba(217,119,6,0.15); color:#B45309; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${prosesCount}</span>
            </div>
            <div class="mutasi-pill ${selectedStatus === 'Selesai' ? 'active' : ''}" data-status="Selesai">
              <span style="color:#16A34A;">✅</span>
              <span>Selesai</span>
              <span class="badge" style="background:rgba(22,163,74,0.15); color:#15803D; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${selesaiCount}</span>
            </div>
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="mutasi-list-container">
          ${filtered.length === 0 ? `
            <div class="mutasi-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">🔍</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Data Mutasi Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="mutasi-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          ` : `
            <div class="mutasi-cards-grid">
              ${filtered.map(m => {
                const isSelesai = m.status === 'Selesai';
                const initials = (m.employee_name || 'Karyawan').split(' ').map(w => w[0]).filter(Boolean).slice(0, 2).join('').toUpperCase();
                return `
                  <div class="mutasi-card" data-id="${m.id}">
                    <div>
                      <!-- Card Header: Status & Tanggal -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="mutasi-status-badge ${isSelesai ? 'selesai' : 'proses'}">
                          <span>${isSelesai ? '✅' : '⏳'}</span>
                          <span>${m.status || 'Proses'}</span>
                        </span>
                        <span style="font-size:0.8rem; font-weight:600; color:#64748B; display:inline-flex; align-items:center; gap:4px;">
                          <span>🗓️</span>
                          <span>${formatDateIndo(m.tanggal)}</span>
                        </span>
                      </div>

                      <!-- Employee Info -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="mutasi-avatar-box">
                          ${initials}
                        </div>
                        <div style="min-width:0;">
                          <h4 style="margin:0 0 3px 0; font-size:1.02rem; font-weight:700; color:#0F172A; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                            ${m.employee_name}
                          </h4>
                          <div style="font-size:0.75rem; color:#64748B;">
                            ID Mutasi: #${m.id}
                          </div>
                        </div>
                      </div>

                      <!-- Route Visual Box -->
                      <div class="mutasi-route-box">
                        <div class="mutasi-route-node">
                          <div class="mutasi-route-label">Cabang Asal</div>
                          <div class="mutasi-route-branch" title="${m.from_branch_name || '-'}">
                            ${m.from_branch_name || '-'}
                          </div>
                        </div>

                        <div class="mutasi-route-arrow">
                          ➔
                        </div>

                        <div class="mutasi-route-node" style="text-align:right;">
                          <div class="mutasi-route-label">Cabang Tujuan</div>
                          <div class="mutasi-route-branch" style="color:#2563EB;" title="${m.to_branch_name || '-'}">
                            ${m.to_branch_name || '-'}
                          </div>
                        </div>
                      </div>
                    </div>

                    <!-- Card Footer: Dokumen & Aksi -->
                    <div style="display:flex; justify-content:space-between; align-items:center; padding-top:12px; border-top:1px solid #F1F5F9; margin-top:8px;">
                      <div>
                        ${m.document_link ? `
                          <a href="${m.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.78rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                            <span>📄 Buka SK</span>
                          </a>
                        ` : `
                          <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                        `}
                      </div>

                      ${canEdit ? `
                        <div style="display:flex; align-items:center; gap:6px;">
                          <button class="mutasi-icon-btn mutasi-edit-btn" title="Edit Data Mutasi" data-id="${m.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </button>
                          <button class="mutasi-icon-btn delete mutasi-delete-btn" title="Hapus Mutasi" data-id="${m.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                          </button>
                        </div>
                      ` : ''}
                    </div>

                  </div>
                `;
              }).join('')}
            </div>
          `}
        </div>

      </div>
    `;

    bindEvents();
  }

  // Event Listeners Binding
  function bindEvents() {
    // Add button
    const btnAdd = container.querySelector('#mutasi-btn-add');
    if (btnAdd) btnAdd.onclick = () => openMutasiModal();

    // Excel button
    const btnExcel = container.querySelector('#mutasi-btn-excel');
    if (btnExcel) btnExcel.onclick = () => exportToExcel();

    // Refresh button
    const btnRefresh = container.querySelector('#mutasi-btn-refresh');
    if (btnRefresh) btnRefresh.onclick = () => loadData();

    // Reset filter
    const btnReset = container.querySelector('#mutasi-btn-reset-filter');
    if (btnReset) {
      btnReset.onclick = () => {
        searchQuery = '';
        selectedFromBranch = 'all';
        selectedToBranch = 'all';
        selectedStatus = 'all';
        renderUI();
      };
    }

    // Search input
    const searchInput = container.querySelector('#mutasi-search-input');
    if (searchInput) {
      searchInput.oninput = (e) => {
        searchQuery = e.target.value;
        renderUI();
        const newSearch = container.querySelector('#mutasi-search-input');
        if (newSearch) {
          newSearch.focus();
          newSearch.setSelectionRange(newSearch.value.length, newSearch.value.length);
        }
      };
    }

    // From branch filter
    const fromFilter = container.querySelector('#mutasi-filter-from');
    if (fromFilter) {
      fromFilter.onchange = (e) => {
        selectedFromBranch = e.target.value;
        renderUI();
      };
    }

    // To branch filter
    const toFilter = container.querySelector('#mutasi-filter-to');
    if (toFilter) {
      toFilter.onchange = (e) => {
        selectedToBranch = e.target.value;
        renderUI();
      };
    }

    // Status Pills
    container.querySelectorAll('.mutasi-pill').forEach(pill => {
      pill.onclick = () => {
        selectedStatus = pill.getAttribute('data-status') || 'all';
        renderUI();
      };
    });

    // Edit Buttons
    container.querySelectorAll('.mutasi-edit-btn').forEach(btn => {
      btn.onclick = (e) => {
        e.stopPropagation();
        const id = Number(btn.getAttribute('data-id'));
        const item = mutasiList.find(m => m.id === id);
        if (item) openMutasiModal(item);
      };
    });

    // Delete Buttons
    container.querySelectorAll('.mutasi-delete-btn').forEach(btn => {
      btn.onclick = (e) => {
        e.stopPropagation();
        const id = Number(btn.getAttribute('data-id'));
        deleteMutasi(id);
      };
    });
  }

  // Initial load
  await loadData();
}
