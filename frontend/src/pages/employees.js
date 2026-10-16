import { apiFetch, getUser } from '../config.js';
import { getCachedBranches } from '../utils/dataCache.js';
import { createModal, confirmDialog } from '../components/modal.js';
import { toastSuccess, toastError, toastInfo } from '../components/toast.js';

// Pre-seeded fallback data jika API offline / inisialisasi awal
const INITIAL_EMPLOYEES = [
  { id: 1, full_name: 'ACHMAD NURYAINI', branch_id: 7, branch_name: '007. Bekasi', division: 'FACILITY CARE', phone: '081298765431', join_date: '2023-10-09', status: 'Aktif', notes: '', target_pindah_os: '', target_selesai_os: '' },
  { id: 2, full_name: 'ADAM JHORDY', branch_id: 32, branch_name: '032. Cawang', division: 'FACILITY CARE', phone: '085712345678', join_date: '2024-03-18', status: 'Aktif', notes: '', target_pindah_os: '', target_selesai_os: '' },
  { id: 3, full_name: 'ADE MUSTIKA REZDAFITRA', branch_id: 16, branch_name: '016. Cibubur', division: 'FACILITY CARE', phone: '081345678901', join_date: '2022-06-21', status: 'Aktif', notes: '', target_pindah_os: '', target_selesai_os: '' },
  { id: 4, full_name: 'ADE SURAHMAN', branch_id: 0, branch_name: '000. Management', division: 'FACILITY CARE', phone: '081223344556', join_date: '2021-01-15', status: 'Aktif', notes: 'Trainer & Supervisor Operasional', target_pindah_os: '', target_selesai_os: '' },
  { id: 5, full_name: 'ADITYA', branch_id: 26, branch_name: '026. Tangerang City', division: 'FACILITY CARE', phone: '087811223344', join_date: '2024-10-19', status: 'Aktif', notes: '', target_pindah_os: '', target_selesai_os: '' },
  { id: 6, full_name: 'ADITYA PRATAMA', branch_id: 54, branch_name: '054. Bogor Gunung Batu', division: 'FACILITY CARE', phone: '089699887766', join_date: '2025-09-04', status: 'Aktif', notes: 'Lulusan Training Batch 5 nilai tertinggi', target_pindah_os: '', target_selesai_os: '' },
  { id: 7, full_name: 'AGUNG SEPTIADI', branch_id: 0, branch_name: '000. Management', division: 'FC - RELIEFER', phone: '082155667788', join_date: '2026-02-12', status: 'Aktif', notes: 'Personil cover cabang mobile', target_pindah_os: '', target_selesai_os: '' }
];

const DIVISION_LIST = ['FACILITY CARE', 'SECURITY', 'FC - RELIEFER'];
const STATUS_LIST = ['Aktif', 'Tidak Aktif', 'Resign', 'Cut'];

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

// Inisial avatar
function getInitials(name) {
  if (!name || name === '-') return '👤';
  const parts = String(name).trim().split(/\s+/);
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return (parts[0][0] + parts[1][0]).toUpperCase();
}

export async function renderEmployees(container, params) {
  let employeeList = [];
  let branchOptions = [];
  let selectedStatus = 'all';
  let selectedDivision = 'all';
  let selectedBranch = 'all';
  let searchQuery = '';
  let isLoading = false;

  // Handle dashboard filter parameter
  const dashFilter = params ? params.get('dash_filter') : null;
  if (dashFilter === 'active') {
    selectedStatus = 'Aktif';
  } else if (dashFilter === 'reliefer') {
    selectedDivision = 'FC - RELIEFER';
  }

  const currentUser = getUser();
  const canEdit = currentUser?.role !== 'viewer';

  // Load branches
  try {
    branchOptions = await getCachedBranches();
  } catch (err) {
    console.error('Error fetching branches:', err);
  }

  // Load Employees Data from API
  async function loadData() {
    isLoading = true;
    renderUI();
    try {
      const res = await apiFetch('/api/employees?limit=10000');
      if (res.ok) {
        const raw = res.data?.data || res.data || [];
        employeeList = Array.isArray(raw) && raw.length > 0 ? raw : INITIAL_EMPLOYEES;
      } else {
        employeeList = INITIAL_EMPLOYEES;
      }
    } catch (e) {
      console.warn('Gagal load employees, menggunakan data lokal:', e);
      employeeList = INITIAL_EMPLOYEES;
    } finally {
      isLoading = false;
      renderUI();
    }
  }

  // Export to Excel
  async function exportToExcel() {
    try {
      const { downloadExcel } = await import('../utils/excel.js');
      const exportData = employeeList.map((emp, idx) => ({
        'No': idx + 1,
        'Nama Lengkap': emp.full_name || '',
        'Cabang': emp.branch_name || '',
        'Divisi': emp.division || '',
        'No. HP': emp.phone || '',
        'Tgl Masuk': emp.join_date || '',
        'Status': emp.status || '',
        'Target Pindah OS': emp.target_pindah_os || '',
        'Target Selesai OS': emp.target_selesai_os || '',
        'Catatan': emp.notes || ''
      }));
      downloadExcel(exportData, `Data_Karyawan_FCMS_${new Date().toISOString().split('T')[0]}`);
      toastSuccess('Berhasil mengekspor data karyawan ke Excel');
    } catch (err) {
      toastError('Gagal ekspor Excel: ' + err.message);
    }
  }

  // Modal Tambah / Edit Karyawan
  function openEmployeeModal(item = null) {
    const isEdit = !!item;
    const defaultDivision = item?.division || 'FACILITY CARE';
    const defaultStatus = item?.status || 'Aktif';

    const content = document.createElement('div');
    content.className = 'emp-modal-body';
    content.innerHTML = `
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Lengkap Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <input type="text" id="emp-modal-name" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: Achmad Nuryaini" value="${item?.full_name ? String(item.full_name).replace(/"/g, '&quot;') : ''}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              No. Handphone / WhatsApp
            </label>
            <input type="tel" id="emp-modal-phone" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="08xx-xxxx-xxxx" value="${item?.phone ? String(item.phone).replace(/"/g, '&quot;') : ''}" />
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Penempatan <span style="color:#DC2626;">*</span>
            </label>
            <select id="emp-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${branchOptions.map(b => `
                <option value="${b.value}" ${String(item?.branch_id) === String(b.value) ? 'selected' : ''}>
                  ${b.label}
                </option>
              `).join('')}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Divisi / Bagian <span style="color:#DC2626;">*</span>
            </label>
            <select id="emp-modal-division" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${DIVISION_LIST.map(div => `
                <option value="${div}" ${defaultDivision === div ? 'selected' : ''}>${div}</option>
              `).join('')}
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Masuk Kerja
            </label>
            <input type="date" id="emp-modal-joindate" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${item?.join_date || ''}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <select id="emp-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${STATUS_LIST.map(st => `
                <option value="${st}" ${defaultStatus === st ? 'selected' : ''}>${st}</option>
              `).join('')}
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Target Pindah OS (Opsional)
            </label>
            <input type="month" id="emp-modal-target-pindah" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${item?.target_pindah_os ? String(item.target_pindah_os).slice(0, 7) : ''}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Target Selesai OS (Opsional)
            </label>
            <input type="month" id="emp-modal-target-selesai" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${item?.target_selesai_os ? String(item.target_selesai_os).slice(0, 7) : ''}" />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Catatan / Keterangan Karyawan
          </label>
          <textarea id="emp-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Catatan keahlian, tugas khusus, riwayat cover...">${item?.notes || ''}</textarea>
        </div>
      </div>
    `;

    createModal({
      title: isEdit ? '✏️ Edit Data Karyawan' : '➕ Tambah Data Karyawan',
      content,
      size: 'md',
      confirmText: isEdit ? 'Simpan Perubahan' : 'Simpan Karyawan',
      confirmClass: 'btn-primary',
      onConfirm: async (_, close) => {
        const full_name = (content.querySelector('#emp-modal-name').value || '').trim();
        const phone = (content.querySelector('#emp-modal-phone').value || '').trim();
        const branch_id = content.querySelector('#emp-modal-branch').value;
        const division = content.querySelector('#emp-modal-division').value;
        const join_date = content.querySelector('#emp-modal-joindate').value || null;
        const status = content.querySelector('#emp-modal-status').value;
        const target_pindah_os = content.querySelector('#emp-modal-target-pindah').value || null;
        const target_selesai_os = content.querySelector('#emp-modal-target-selesai').value || null;
        const notes = (content.querySelector('#emp-modal-notes').value || '').trim();

        if (!full_name || !branch_id || !division || !status) {
          toastError('Harap lengkapi Nama Lengkap, Cabang, Divisi, dan Status!');
          return;
        }

        const payload = {
          full_name,
          phone,
          branch_id: Number(branch_id),
          division,
          join_date,
          status,
          target_pindah_os,
          target_selesai_os,
          notes
        };

        try {
          let res;
          if (isEdit) {
            res = await apiFetch(`/api/employees/${item.id}`, {
              method: 'PUT',
              body: JSON.stringify(payload)
            });
          } else {
            res = await apiFetch('/api/employees', {
              method: 'POST',
              body: JSON.stringify(payload)
            });
          }

          if (res.ok) {
            toastSuccess(isEdit ? 'Data karyawan berhasil diperbarui!' : 'Karyawan baru berhasil ditambahkan!');
            close();
            await loadData();
          } else {
            toastError(res.data?.error || 'Gagal menyimpan data karyawan');
          }
        } catch (err) {
          toastError('Terjadi kesalahan: ' + err.message);
        }
      }
    });
  }

  // Delete / Deactivate Employee
  async function deleteEmployee(item) {
    confirmDialog({
      title: 'Nonaktifkan Karyawan?',
      message: `Apakah Anda yakin ingin menonaktifkan personil <strong>${item.full_name}</strong> dari cabang <em>${item.branch_name || 'cabang'}</em>?`,
      confirmText: 'Ya, Nonaktifkan',
      confirmClass: 'btn-danger',
      onConfirm: async () => {
        try {
          const res = await apiFetch(`/api/employees/${item.id}`, {
            method: 'DELETE'
          });
          if (res.ok) {
            toastSuccess('Status karyawan berhasil diubah menjadi Tidak Aktif');
            await loadData();
          } else {
            toastError(res.data?.error || 'Gagal menonaktifkan data');
          }
        } catch (err) {
          toastError('Terjadi kesalahan: ' + err.message);
        }
      }
    });
  }

  // Filter Data
  function getFilteredList() {
    return employeeList.filter(item => {
      // Search
      if (searchQuery) {
        const q = searchQuery.toLowerCase();
        const matchName = (item.full_name || '').toLowerCase().includes(q);
        const matchBranch = (item.branch_name || '').toLowerCase().includes(q);
        const matchPhone = (item.phone || '').toLowerCase().includes(q);
        const matchDiv = (item.division || '').toLowerCase().includes(q);
        const matchNotes = (item.notes || '').toLowerCase().includes(q);
        if (!matchName && !matchBranch && !matchPhone && !matchDiv && !matchNotes) {
          return false;
        }
      }

      // Status
      if (selectedStatus !== 'all') {
        if (selectedStatus === 'Aktif' && item.status !== 'Aktif') return false;
        if (selectedStatus === 'Nonaktif' && item.status === 'Aktif') return false;
      }

      // Division
      if (selectedDivision !== 'all' && item.division !== selectedDivision) {
        return false;
      }

      // Branch
      if (selectedBranch !== 'all' && String(item.branch_id) !== String(selectedBranch)) {
        return false;
      }

      return true;
    });
  }

  // Main UI Renderer
  function renderUI() {
    const filtered = getFilteredList();

    // KPI Metrics calculation
    const totalEmployees = employeeList.length;
    const activeCount = employeeList.filter(e => e.status === 'Aktif').length;
    const fcCount = employeeList.filter(e => e.division === 'FACILITY CARE' && e.status === 'Aktif').length;
    const secAndRelCount = employeeList.filter(e => (e.division === 'SECURITY' || e.division === 'FC - RELIEFER') && e.status === 'Aktif').length;

    // Counts for status pills
    const countAll = employeeList.length;
    const countActive = employeeList.filter(e => e.status === 'Aktif').length;
    const countNonactive = employeeList.filter(e => e.status !== 'Aktif').length;

    container.innerHTML = `
      <style>
        .emp-page {
          display: flex;
          flex-direction: column;
          gap: 20px;
        }
        .emp-hero-card {
          background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
          border-radius: 16px;
          padding: 24px 28px;
          color: white;
          box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25);
          position: relative;
          overflow: hidden;
        }
        .emp-hero-card::after {
          content: '';
          position: absolute;
          top: -40px;
          right: -40px;
          width: 220px;
          height: 220px;
          background: radial-gradient(circle, rgba(14, 165, 233, 0.18) 0%, transparent 70%);
          border-radius: 50%;
          pointer-events: none;
        }
        .emp-stats-row {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 14px;
          margin-top: 20px;
        }
        .emp-stat-box {
          background: rgba(255, 255, 255, 0.05);
          border: 1px solid rgba(255, 255, 255, 0.1);
          border-radius: 12px;
          padding: 14px 18px;
          display: flex;
          align-items: center;
          gap: 14px;
          backdrop-filter: blur(8px);
          transition: transform 0.15s ease, background 0.15s ease;
        }
        .emp-stat-box:hover {
          background: rgba(255, 255, 255, 0.08);
          transform: translateY(-2px);
        }
        .emp-stat-icon {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 1.3rem;
          flex-shrink: 0;
        }
        .emp-stat-val {
          font-size: 1.55rem;
          font-weight: 800;
          line-height: 1.1;
          letter-spacing: -0.02em;
        }
        .emp-stat-lbl {
          font-size: 0.72rem;
          text-transform: uppercase;
          color: #94A3B8;
          font-weight: 700;
          letter-spacing: 0.05em;
          margin-top: 2px;
        }

        /* Controls / Filter Bar */
        .emp-controls-card {
          background: white;
          border-radius: 14px;
          border: 1px solid #E2E8F0;
          padding: 16px 20px;
          box-shadow: 0 2px 8px rgba(0,0,0,0.03);
          display: flex;
          flex-direction: column;
          gap: 14px;
        }
        .emp-search-row {
          display: flex;
          gap: 12px;
          flex-wrap: wrap;
        }
        .emp-search-wrap {
          position: relative;
          flex: 1;
          min-width: 240px;
        }
        .emp-search-icon {
          position: absolute;
          left: 14px;
          top: 50%;
          transform: translateY(-50%);
          color: #94A3B8;
          pointer-events: none;
        }
        .emp-search-input {
          width: 100%;
          padding: 10px 14px 10px 40px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          outline: none;
          transition: border-color 0.15s ease;
        }
        .emp-search-input:focus {
          border-color: #2563EB;
          box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);
        }
        .emp-select {
          padding: 10px 14px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          color: #1E293B;
          background: white;
          outline: none;
          min-width: 170px;
          cursor: pointer;
        }
        .emp-select:focus {
          border-color: #2563EB;
        }

        /* Swipeable Pills */
        .emp-filter-pills {
          display: flex;
          align-items: center;
          gap: 8px;
          overflow-x: auto;
          white-space: nowrap;
          scrollbar-width: none;
          padding-bottom: 2px;
        }
        .emp-filter-pills::-webkit-scrollbar {
          display: none;
        }
        .emp-pill {
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
        .emp-pill:hover {
          background: #F1F5F9;
          color: #1E293B;
        }
        .emp-pill.active {
          background: #0F172A;
          color: white;
          border-color: #0F172A;
        }
        .emp-pill-dot {
          width: 8px;
          height: 8px;
          border-radius: 50%;
        }
        .emp-pill-count {
          padding: 1px 6px;
          border-radius: 10px;
          font-size: 0.72rem;
          background: rgba(0,0,0,0.06);
        }
        .emp-pill.active .emp-pill-count {
          background: rgba(255,255,255,0.2);
          color: white;
        }

        /* Cards Grid */
        .emp-cards-grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(360px, 1fr));
          gap: 16px;
        }
        .emp-card {
          background: white;
          border-radius: 14px;
          border: 1px solid #E2E8F0;
          padding: 18px;
          box-shadow: 0 2px 8px rgba(0,0,0,0.04);
          display: flex;
          flex-direction: column;
          gap: 14px;
          transition: transform 0.15s ease, box-shadow 0.15s ease;
        }
        .emp-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 6px 16px rgba(0,0,0,0.08);
          border-color: #CBD5E1;
        }
        .emp-card-top {
          display: flex;
          justify-content: space-between;
          align-items: center;
          gap: 8px;
        }
        
        .emp-div-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.76rem;
          font-weight: 700;
        }
        .emp-div-fc {
          background: #EFF6FF;
          color: #2563EB;
          border: 1px solid #DBEAFE;
        }
        .emp-div-sec {
          background: #F1F5F9;
          color: #334155;
          border: 1px solid #E2E8F0;
        }
        .emp-div-rel {
          background: #F5F3FF;
          color: #7C3AED;
          border: 1px solid #EDE9FE;
        }

        .emp-status-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.76rem;
          font-weight: 700;
        }
        .emp-status-aktif {
          background: #DCFCE7;
          color: #16A34A;
          border: 1px solid #BBF7D0;
        }
        .emp-status-nonaktif {
          background: #FEE2E2;
          color: #DC2626;
          border: 1px solid #FCA5A5;
        }

        /* Profile Box */
        .emp-profile-box {
          display: flex;
          align-items: center;
          gap: 14px;
        }
        .emp-avatar {
          width: 52px;
          height: 52px;
          border-radius: 12px;
          background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
          color: white;
          display: flex;
          align-items: center;
          justify-content: center;
          font-weight: 800;
          font-size: 1.15rem;
          flex-shrink: 0;
          box-shadow: 0 4px 10px rgba(37, 99, 235, 0.2);
        }
        .emp-avatar.sec {
          background: linear-gradient(135deg, #334155 0%, #0F172A 100%);
          box-shadow: 0 4px 10px rgba(15, 23, 42, 0.2);
        }
        .emp-avatar.rel {
          background: linear-gradient(135deg, #7C3AED 0%, #5B21B6 100%);
          box-shadow: 0 4px 10px rgba(124, 58, 237, 0.2);
        }
        .emp-name {
          font-size: 1.05rem;
          font-weight: 800;
          color: #0F172A;
          line-height: 1.25;
        }
        .emp-subtitle {
          font-size: 0.76rem;
          color: #64748B;
          font-weight: 600;
          margin-top: 3px;
        }

        /* Info Container */
        .emp-info-card {
          background: #F8FAFC;
          border: 1px solid #F1F5F9;
          border-radius: 10px;
          padding: 12px;
          display: flex;
          flex-direction: column;
          gap: 8px;
          font-size: 0.84rem;
        }
        .emp-info-row {
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 8px;
        }
        .emp-info-label {
          color: #64748B;
          display: flex;
          align-items: center;
          gap: 6px;
        }
        .emp-info-val {
          color: #0F172A;
          font-weight: 600;
          text-align: right;
        }
        .emp-phone-link {
          color: #2563EB;
          text-decoration: none;
          font-weight: 600;
          display: inline-flex;
          align-items: center;
          gap: 4px;
        }
        .emp-phone-link:hover {
          text-decoration: underline;
        }

        /* Notes box */
        .emp-notes-box {
          font-size: 0.8rem;
          color: #475569;
          background: #FFFBEB;
          border: 1px solid #FEF3C7;
          border-radius: 8px;
          padding: 8px 10px;
          line-height: 1.4;
        }

        /* Card footer */
        .emp-card-footer {
          display: flex;
          justify-content: space-between;
          align-items: center;
          border-top: 1px solid #F1F5F9;
          padding-top: 12px;
          margin-top: auto;
          gap: 8px;
        }
        .emp-action-btns {
          display: flex;
          align-items: center;
          gap: 6px;
          margin-left: auto;
        }
        .emp-icon-btn {
          width: 34px;
          height: 34px;
          border-radius: 8px;
          border: 1px solid #E2E8F0;
          background: white;
          color: #64748B;
          display: flex;
          align-items: center;
          justify-content: center;
          cursor: pointer;
          transition: all 0.15s ease;
        }
        .emp-icon-btn:hover {
          background: #F1F5F9;
          color: #0F172A;
          border-color: #CBD5E1;
        }
        .emp-icon-btn.delete:hover {
          background: #FEE2E2;
          color: #DC2626;
          border-color: #FCA5A5;
        }

        /* Empty state */
        .emp-empty-state {
          grid-column: 1 / -1;
          background: white;
          border: 2px dashed #E2E8F0;
          border-radius: 16px;
          padding: 48px 24px;
          text-align: center;
        }

        /* Mobile Responsive */
        @media (max-width: 768px) {
          .emp-hero-card {
            padding: 18px 16px;
          }
          .emp-stats-row {
            grid-template-columns: 1fr 1fr;
            gap: 10px;
          }
          .emp-stat-box {
            padding: 10px 12px;
          }
          .emp-stat-icon {
            width: 38px;
            height: 38px;
            font-size: 1.1rem;
          }
          .emp-stat-val {
            font-size: 1.3rem;
          }
          .emp-hero-actions {
            width: 100% !important;
            display: flex !important;
            align-items: center !important;
            gap: 8px !important;
            flex-wrap: nowrap !important;
          }
          .emp-btn-add {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.82rem !important;
            white-space: nowrap !important;
          }
          .emp-btn-excel {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.82rem !important;
            white-space: nowrap !important;
          }
          .emp-btn-refresh {
            width: 38px !important;
            height: 38px !important;
            padding: 0 !important;
            flex-shrink: 0 !important;
            justify-content: center !important;
          }
          .emp-controls-card {
            padding: 14px;
            gap: 12px;
          }
          .emp-search-row {
            flex-direction: column;
            gap: 10px;
          }
          .emp-select {
            width: 100%;
          }
          .emp-cards-grid {
            grid-template-columns: 1fr;
          }
        }
      </style>

      <div class="emp-page">

        <!-- Hero Header Card -->
        <div class="emp-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px; margin-bottom:6px;">
                <span style="font-size:1.6rem;">👥</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Master Karyawan FCMS
                </h1>
              </div>
              <p style="margin:0; font-size:0.88rem; color:#CBD5E1; max-width:650px; line-height:1.5;">
                Pusat data personil Facility Care, Security, dan Reliefer. Mengatur status keaktifan, kontak, penempatan cabang, dan riwayat kerja.
              </p>
            </div>

            <div class="emp-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${canEdit ? `
                <button id="emp-btn-add" class="btn btn-primary emp-btn-add" style="background:#2563EB; padding:9px 16px; border-radius:10px; font-weight:700; display:flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Karyawan</span>
                </button>
              ` : ''}
              
              <button id="emp-btn-excel" class="btn btn-ghost emp-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; padding:9px 14px; border-radius:10px; font-weight:600; display:flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>📥</span>
                <span>Ekspor Excel</span>
              </button>

              <button id="emp-btn-refresh" class="btn btn-ghost emp-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24" class="${isLoading ? 'spin-icon' : ''}"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- 4 KPI Stat Boxes -->
          <div class="emp-stats-row">
            <div class="emp-stat-box">
              <div class="emp-stat-icon" style="background:rgba(37, 99, 235, 0.2); color:#93C5FD;">👥</div>
              <div>
                <div class="emp-stat-val">${totalEmployees}</div>
                <div class="emp-stat-lbl">Total Karyawan</div>
              </div>
            </div>

            <div class="emp-stat-box">
              <div class="emp-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#6EE7B7;">🟢</div>
              <div>
                <div class="emp-stat-val">${activeCount}</div>
                <div class="emp-stat-lbl">Personil Aktif</div>
              </div>
            </div>

            <div class="emp-stat-box">
              <div class="emp-stat-icon" style="background:rgba(14, 165, 233, 0.2); color:#7DD3FC;">🏢</div>
              <div>
                <div class="emp-stat-val">${fcCount}</div>
                <div class="emp-stat-lbl">Facility Care</div>
              </div>
            </div>

            <div class="emp-stat-box">
              <div class="emp-stat-icon" style="background:rgba(124, 58, 237, 0.2); color:#C4B5FD;">🛡️</div>
              <div>
                <div class="emp-stat-val">${secAndRelCount}</div>
                <div class="emp-stat-lbl">Security & Reliefer</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Filter Controls Card -->
        <div class="emp-controls-card">
          <div class="emp-search-row">
            <div class="emp-search-wrap">
              <svg class="emp-search-icon" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              <input type="text" id="emp-search-input" class="emp-search-input" 
                placeholder="Cari nama karyawan, no HP, cabang..." value="${searchQuery ? String(searchQuery).replace(/"/g, '&quot;') : ''}" />
            </div>

            <select id="emp-branch-filter" class="emp-select">
              <option value="all">Semua Cabang (${branchOptions.length})</option>
              ${branchOptions.map(b => `
                <option value="${b.value}" ${String(selectedBranch) === String(b.value) ? 'selected' : ''}>
                  ${b.label}
                </option>
              `).join('')}
            </select>

            <select id="emp-division-filter" class="emp-select">
              <option value="all">Semua Divisi</option>
              ${DIVISION_LIST.map(div => `
                <option value="${div}" ${selectedDivision === div ? 'selected' : ''}>
                  ${div}
                </option>
              `).join('')}
            </select>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="emp-filter-pills">
            <div class="emp-pill ${selectedStatus === 'all' && selectedDivision === 'all' ? 'active' : ''}" data-status="all" data-div="all">
              <span>Semua Karyawan</span>
              <span class="emp-pill-count">${countAll}</span>
            </div>
            <div class="emp-pill ${selectedStatus === 'Aktif' ? 'active' : ''}" data-status="Aktif" data-div="all">
              <span class="emp-pill-dot" style="background:#10B981;"></span>
              <span>Aktif</span>
              <span class="emp-pill-count">${countActive}</span>
            </div>
            <div class="emp-pill ${selectedDivision === 'FC - RELIEFER' ? 'active' : ''}" data-status="all" data-div="FC - RELIEFER">
              <span class="emp-pill-dot" style="background:#8B5CF6;"></span>
              <span>FC - Reliefer</span>
              <span class="emp-pill-count">${employeeList.filter(e => e.division === 'FC - RELIEFER').length}</span>
            </div>
            <div class="emp-pill ${selectedStatus === 'Nonaktif' ? 'active' : ''}" data-status="Nonaktif" data-div="all">
              <span class="emp-pill-dot" style="background:#EF4444;"></span>
              <span>Nonaktif / Resign</span>
              <span class="emp-pill-count">${countNonactive}</span>
            </div>
          </div>
        </div>

        <!-- Exclusive Cards Grid -->
        <div class="emp-cards-grid">
          ${filtered.length === 0 ? `
            <div class="emp-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">🔍</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak ada karyawan yang cocok</h3>
              <p style="font-size:0.85rem; color:#64748B; margin:0 0 16px 0;">Coba sesuaikan kata kunci pencarian, filter cabang, atau filter divisi.</p>
              <button id="emp-reset-filter" class="btn btn-secondary" style="font-size:0.82rem; padding:7px 16px; border-radius:8px;">Reset Semua Filter</button>
            </div>
          ` : filtered.map(item => {
            const div = item.division || 'FACILITY CARE';
            let divClass = 'emp-div-fc';
            let avatarClass = '';
            if (div === 'SECURITY') {
              divClass = 'emp-div-sec';
              avatarClass = 'sec';
            } else if (div === 'FC - RELIEFER') {
              divClass = 'emp-div-rel';
              avatarClass = 'rel';
            }

            const isAktif = item.status === 'Aktif';
            const statusClass = isAktif ? 'emp-status-aktif' : 'emp-status-nonaktif';
            const statusText = isAktif ? '🟢 Aktif' : `⚪ ${item.status || 'Nonaktif'}`;

            return `
              <div class="emp-card" data-id="${item.id}">
                <!-- Top Badges Row -->
                <div class="emp-card-top">
                  <span class="emp-div-badge ${divClass}">
                    ${div === 'SECURITY' ? '🛡️' : (div === 'FC - RELIEFER' ? '🔄' : '🧹')} ${div}
                  </span>
                  <div style="display:flex; align-items:center; gap:8px;">
                    <span class="emp-status-badge ${statusClass}">
                      ${statusText}
                    </span>
                  </div>
                </div>

                <!-- Profile Box -->
                <div class="emp-profile-box">
                  <div class="emp-avatar ${avatarClass}">
                    ${getInitials(item.full_name)}
                  </div>
                  <div style="min-width:0; flex:1;">
                    <div class="emp-name" style="white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${item.full_name || '-'}
                    </div>
                    <div class="emp-subtitle">
                      Personil ${div}
                    </div>
                  </div>
                </div>

                <!-- Information Grid -->
                <div class="emp-info-card">
                  <div class="emp-info-row">
                    <span class="emp-info-label">📍 Cabang:</span>
                    <span class="emp-info-val">${item.branch_name || 'Cabang Tidak Ditentukan'}</span>
                  </div>
                  <div class="emp-info-row">
                    <span class="emp-info-label">📞 No. HP:</span>
                    <span class="emp-info-val">
                      ${item.phone ? `
                        <a href="tel:${item.phone}" class="emp-phone-link">
                          <span>${item.phone}</span>
                        </a>
                      ` : '<span style="color:#94A3B8;">-</span>'}
                    </span>
                  </div>
                  <div class="emp-info-row">
                    <span class="emp-info-label">🗓️ Tgl Masuk:</span>
                    <span class="emp-info-val">${formatDateIndo(item.join_date)}</span>
                  </div>
                  ${item.target_pindah_os || item.target_selesai_os ? `
                    <div class="emp-info-row" style="border-top:1px dashed #E2E8F0; padding-top:6px; margin-top:2px;">
                      <span class="emp-info-label">🔄 Target OS:</span>
                      <span class="emp-info-val" style="font-size:0.78rem; color:#6366F1;">
                        ${item.target_pindah_os ? 'Pindah: ' + item.target_pindah_os : ''}
                        ${item.target_selesai_os ? ' → Selesai: ' + item.target_selesai_os : ''}
                      </span>
                    </div>
                  ` : ''}
                </div>

                <!-- Notes box (if any) -->
                ${item.notes ? `
                  <div class="emp-notes-box">
                    📝 <strong>Catatan:</strong> ${item.notes}
                  </div>
                ` : ''}

                <!-- Footer Actions -->
                <div class="emp-card-footer">
                  <div style="font-size:0.75rem; color:#94A3B8;">
                    ID #${item.id}
                  </div>

                  <div class="emp-action-btns">
                    ${item.phone ? `
                      <a href="https://wa.me/${item.phone.replace(/[^0-9]/g, '').replace(/^0/, '62')}" target="_blank" rel="noopener" class="emp-icon-btn" title="Chat WhatsApp" style="color:#16A34A;">
                        <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                      </a>
                    ` : ''}

                    ${canEdit ? `
                      <button class="emp-icon-btn edit" data-action="edit" data-id="${item.id}" title="Edit Data Karyawan">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                      </button>
                      <button class="emp-icon-btn delete" data-action="delete" data-id="${item.id}" title="Nonaktifkan Karyawan">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                      </button>
                    ` : ''}
                  </div>
                </div>
              </div>
            `;
          }).join('')}
        </div>

      </div>
    `;

    // ── Event Handlers ──
    const searchInput = container.querySelector('#emp-search-input');
    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        searchQuery = e.target.value;
        renderUI();
        const el = container.querySelector('#emp-search-input');
        if (el) {
          el.focus();
          el.setSelectionRange(el.value.length, el.value.length);
        }
      });
    }

    const branchFilter = container.querySelector('#emp-branch-filter');
    if (branchFilter) {
      branchFilter.addEventListener('change', (e) => {
        selectedBranch = e.target.value;
        renderUI();
      });
    }

    const divisionFilter = container.querySelector('#emp-division-filter');
    if (divisionFilter) {
      divisionFilter.addEventListener('change', (e) => {
        selectedDivision = e.target.value;
        renderUI();
      });
    }

    // Status Pills
    container.querySelectorAll('.emp-pill').forEach(pill => {
      pill.addEventListener('click', () => {
        selectedStatus = pill.dataset.status;
        const divTarget = pill.dataset.div;
        if (divTarget && divTarget !== 'all') {
          selectedDivision = divTarget;
        } else if (pill.dataset.status === 'all') {
          selectedDivision = 'all';
        }
        renderUI();
      });
    });

    // Reset button
    const resetBtn = container.querySelector('#emp-reset-filter');
    if (resetBtn) {
      resetBtn.addEventListener('click', () => {
        searchQuery = '';
        selectedBranch = 'all';
        selectedDivision = 'all';
        selectedStatus = 'all';
        renderUI();
      });
    }

    // Top action buttons
    const addBtn = container.querySelector('#emp-btn-add');
    if (addBtn) {
      addBtn.addEventListener('click', () => openEmployeeModal());
    }

    const excelBtn = container.querySelector('#emp-btn-excel');
    if (excelBtn) {
      excelBtn.addEventListener('click', () => exportToExcel());
    }

    const refreshBtn = container.querySelector('#emp-btn-refresh');
    if (refreshBtn) {
      refreshBtn.addEventListener('click', () => loadData());
    }

    // Card Action Buttons Delegation
    container.querySelectorAll('.emp-card').forEach(card => {
      const id = Number(card.dataset.id);
      const item = employeeList.find(x => x.id === id);
      if (!item) return;

      const editBtn = card.querySelector('[data-action="edit"]');
      if (editBtn) {
        editBtn.addEventListener('click', () => openEmployeeModal(item));
      }

      const delBtn = card.querySelector('[data-action="delete"]');
      if (delBtn) {
        delBtn.addEventListener('click', () => deleteEmployee(item));
      }
    });
  }

  // Initial load
  await loadData();
}
