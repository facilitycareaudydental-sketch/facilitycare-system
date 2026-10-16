import { apiFetch, getUser } from '../config.js';
import { getCachedBranches, getCachedEmployees } from '../utils/dataCache.js';
import { createModal, confirmDialog } from '../components/modal.js';
import { toastSuccess, toastError, toastInfo } from '../components/toast.js';

// Pre-seeded fallback data lembur jika API offline / inisialisasi awal
const INITIAL_OVERTIME = [
  { id: 1, date: '2026-10-03', branch_id: 25, branch_name: '025. Bogor Baru', employee_id: 1, employee_name: 'ADE SURAHMAN', start_time: '23:00', end_time: '08:00', break_hours: 1, total_hours: 8, reason: 'General Cleaning' },
  { id: 2, date: '2026-10-03', branch_id: 25, branch_name: '025. Bogor Baru', employee_id: 2, employee_name: 'TAUFIK HIDAYAT', start_time: '00:00', end_time: '08:00', break_hours: 1, total_hours: 7, reason: 'General Cleaning' },
  { id: 3, date: '2026-10-03', branch_id: 25, branch_name: '025. Bogor Baru', employee_id: 3, employee_name: 'YUDA WILDAN WIGUNA', start_time: '22:00', end_time: '08:00', break_hours: 1, total_hours: 9, reason: 'General Cleaning' },
  { id: 4, date: '2026-10-03', branch_id: 25, branch_name: '025. Bogor Baru', employee_id: 4, employee_name: 'REZA RIZKI', start_time: '22:00', end_time: '08:00', break_hours: 1, total_hours: 9, reason: 'General Cleaning' },
  { id: 5, date: '2026-10-03', branch_id: 25, branch_name: '025. Bogor Baru', employee_id: 5, employee_name: 'ADITYA PRATAMA', start_time: '22:00', end_time: '08:00', break_hours: 1, total_hours: 9, reason: 'General Cleaning' }
];

const OVERTIME_REASONS = [
  'General Cleaning',
  'Deep Cleaning',
  'Cover manpower',
  'Longshift',
  'Pekerjaan urgent'
];

// Helper kalkulasi total jam lembur otomatis (termasuk lintas tengah malam)
function calculateTotalHours(start, end, breakHrs) {
  if (!start || !end) return 0;
  try {
    const [sh, sm] = start.split(':').map(Number);
    const [eh, em] = end.split(':').map(Number);
    let totalMins = (eh * 60 + em) - (sh * 60 + sm);
    if (totalMins < 0) totalMins += 24 * 60; // Lintas tengah malam (misal 23:00 - 08:00)
    let totalHrs = (totalMins / 60) - (parseFloat(breakHrs) || 0);
    if (totalHrs < 0) totalHrs = 0;
    return Number.isInteger(totalHrs) ? totalHrs : parseFloat(totalHrs.toFixed(2));
  } catch {
    return 0;
  }
}

// Helper badge tema warna alasan lembur
function getReasonBadge(reason = '') {
  const r = (reason || '').toLowerCase();
  if (r.includes('general')) {
    return { bg: '#EFF6FF', color: '#1D4ED8', border: '#BFDBFE', icon: '🧹' };
  }
  if (r.includes('deep')) {
    return { bg: '#F5F3FF', color: '#6D28D9', border: '#DDD6FE', icon: '✨' };
  }
  if (r.includes('cover')) {
    return { bg: '#ECFDF5', color: '#047857', border: '#A7F3D0', icon: '👥' };
  }
  if (r.includes('longshift')) {
    return { bg: '#FFFBEB', color: '#B45309', border: '#FDE68A', icon: '⏳' };
  }
  if (r.includes('urgent')) {
    return { bg: '#FFF1F2', color: '#BE123C', border: '#FECDD3', icon: '🚨' };
  }
  return { bg: '#F8FAFC', color: '#475569', border: '#CBD5E1', icon: '⏱️' };
}

// Format tanggal lokal Indonesia
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

export async function renderOvertime(container) {
  let overtimeList = [];
  let branchOptions = [];
  let employeeOptions = [];
  let selectedReason = 'all';
  let selectedBranch = 'all';
  let searchQuery = '';
  let isLoading = false;

  const currentUser = getUser();
  const isViewer = currentUser?.role === 'viewer';
  const isInputLembur = currentUser?.role === 'input_lembur';
  const canCreate = !isViewer;
  const canEdit = !isViewer;
  const canDelete = !isViewer && !isInputLembur;

  // Load branches & employees cache
  try {
    [branchOptions, employeeOptions] = await Promise.all([
      getCachedBranches(),
      getCachedEmployees()
    ]);
  } catch (err) {
    console.error('Error fetching cache:', err);
  }

  // Load Overtime Data from API
  async function loadData() {
    isLoading = true;
    renderUI();
    try {
      const res = await apiFetch('/api/overtime?limit=10000');
      if (res.ok) {
        const raw = res.data?.data || res.data || [];
        overtimeList = Array.isArray(raw) && raw.length > 0 ? raw : INITIAL_OVERTIME;
      } else {
        overtimeList = INITIAL_OVERTIME;
      }
    } catch (e) {
      console.warn('Gagal load lembur, menggunakan data lokal:', e);
      overtimeList = INITIAL_OVERTIME;
    } finally {
      isLoading = false;
      renderUI();
    }
  }

  // Export to Excel
  async function exportToExcel() {
    try {
      const { downloadExcel } = await import('../utils/excel.js');
      const exportData = overtimeList.map((o, idx) => ({
        'No': idx + 1,
        'Tanggal': o.date || '',
        'Cabang Lembur': o.branch_name || '',
        'Nama Karyawan': o.employee_name || '',
        'Jam Mulai': o.start_time || '',
        'Jam Selesai': o.end_time || '',
        'Istirahat (Jam)': o.break_hours != null ? o.break_hours : 0,
        'Total Jam': o.total_hours != null ? o.total_hours : 0,
        'Alasan / Keterangan': o.reason || ''
      }));
      downloadExcel(exportData, `Data_Lembur_${new Date().toISOString().slice(0, 10)}`);
      toastSuccess('Data lembur berhasil diekspor ke Excel!');
    } catch (err) {
      toastError('Gagal ekspor Excel: ' + err.message);
    }
  }

  // Delete Overtime
  async function deleteOvertime(id) {
    if (!canDelete) return;
    const item = overtimeList.find(o => o.id === id);
    confirmDialog({
      title: 'Hapus Catatan Lembur',
      message: `Apakah Anda yakin ingin menghapus catatan lembur untuk "${item?.employee_name || 'karyawan ini'}" pada tanggal ${formatDateIndo(item?.date)}?`,
      confirmText: 'Ya, Hapus',
      confirmClass: 'btn-danger',
      onConfirm: async () => {
        try {
          const res = await apiFetch(`/api/overtime/${id}`, { method: 'DELETE' });
          if (res.ok) {
            toastSuccess('Catatan lembur berhasil dihapus');
            overtimeList = overtimeList.filter(o => o.id !== id);
            renderUI();
          } else {
            toastError(res.data?.error || 'Gagal menghapus lembur');
          }
        } catch (e) {
          toastError('Gagal menghapus catatan lembur: ' + e.message);
        }
      }
    });
  }

  // Open Form Modal (Add / Edit)
  function openOvertimeModal(item = null) {
    if (!canEdit) return;
    const isEdit = !!item;
    const defaultDate = item?.date ? item.date.slice(0, 10) : new Date().toISOString().slice(0, 10);
    const defaultStart = item?.start_time || '22:00';
    const defaultEnd = item?.end_time || '08:00';
    const defaultBreak = item?.break_hours != null ? item.break_hours : 1;
    const defaultTotal = item?.total_hours != null ? item.total_hours : calculateTotalHours(defaultStart, defaultEnd, defaultBreak);

    const content = document.createElement('div');
    content.innerHTML = `
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Lembur <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="ot-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${defaultDate}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Penugasan <span style="color:#DC2626;">*</span>
            </label>
            <select id="ot-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${branchOptions.map(b => `
                <option value="${b.value}" ${String(item?.branch_id) === String(b.value) ? 'selected' : ''}>
                  ${b.label}
                </option>
              `).join('')}
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Nama Karyawan <span style="color:#DC2626;">*</span>
          </label>
          <select id="ot-modal-emp" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
            <option value="">-- Pilih Karyawan --</option>
            ${employeeOptions.map(e => {
              const label = typeof e === 'string' ? e : (e.label || e.full_name || e.name || '');
              const val = typeof e === 'object' && e !== null ? (e.value != null ? e.value : (e.id != null ? e.id : label)) : e;
              const isSelected = String(item?.employee_id) === String(val) || item?.employee_name === label;
              return `<option value="${val}" ${isSelected ? 'selected' : ''}>${label}</option>`;
            }).join('')}
          </select>
        </div>

        <!-- Shift & Waktu Lembur -->
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Jam Mulai <span style="color:#DC2626;">*</span>
            </label>
            <input type="time" id="ot-modal-start" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${defaultStart}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Jam Selesai <span style="color:#DC2626;">*</span>
            </label>
            <input type="time" id="ot-modal-end" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${defaultEnd}" required />
          </div>
        </div>

        <!-- Durasi & Istirahat -->
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Istirahat (Jam) <span style="color:#DC2626;">*</span>
            </label>
            <input type="number" step="0.5" min="0" max="10" id="ot-modal-break" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${defaultBreak}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Total Jam Lembur <span style="color:#DC2626;">*</span>
            </label>
            <input type="number" step="0.1" min="0" id="ot-modal-total" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; background:#F8FAFC; font-weight:700; color:#2563EB;" 
              value="${defaultTotal}" required />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Alasan / Keterangan Penugasan Lembur <span style="color:#DC2626;">*</span>
          </label>
          <select id="ot-modal-reason" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
            ${OVERTIME_REASONS.map(r => `
              <option value="${r}" ${item?.reason === r ? 'selected' : ''}>${r}</option>
            `).join('')}
          </select>
        </div>
      </div>
    `;

    // Auto-calculate total hours in modal
    const startIn = content.querySelector('#ot-modal-start');
    const endIn = content.querySelector('#ot-modal-end');
    const breakIn = content.querySelector('#ot-modal-break');
    const totalIn = content.querySelector('#ot-modal-total');

    function updateCalculatedHours() {
      const s = startIn.value;
      const e = endIn.value;
      const b = breakIn.value;
      totalIn.value = calculateTotalHours(s, e, b);
    }

    startIn.addEventListener('input', updateCalculatedHours);
    endIn.addEventListener('input', updateCalculatedHours);
    breakIn.addEventListener('input', updateCalculatedHours);

    createModal({
      title: isEdit ? '✏️ Edit Data Lembur Karyawan' : '➕ Catat Penugasan Lembur Baru',
      content,
      size: 'md',
      confirmText: isEdit ? 'Simpan Perubahan' : 'Catat Lembur',
      confirmClass: 'btn-primary',
      onConfirm: async (_, close) => {
        const date = content.querySelector('#ot-modal-date').value;
        const branch_id = content.querySelector('#ot-modal-branch').value;
        const employee_id = content.querySelector('#ot-modal-emp').value;
        const start_time = content.querySelector('#ot-modal-start').value;
        const end_time = content.querySelector('#ot-modal-end').value;
        const break_hours = parseFloat(content.querySelector('#ot-modal-break').value) || 0;
        const total_hours = parseFloat(content.querySelector('#ot-modal-total').value) || 0;
        const reason = content.querySelector('#ot-modal-reason').value;

        if (!date) {
          toastError('Tanggal lembur wajib diisi');
          return;
        }
        if (!branch_id) {
          toastError('Cabang penugasan wajib dipilih');
          return;
        }
        if (!employee_id) {
          toastError('Nama karyawan wajib dipilih');
          return;
        }
        if (!start_time || !end_time) {
          toastError('Jam mulai dan jam selesai wajib diisi');
          return;
        }

        const bObj = branchOptions.find(b => String(b.value) === String(branch_id));
        const eObj = employeeOptions.find(e => String(e.value) === String(employee_id) || String(e.id) === String(employee_id));
        const employee_name = eObj ? (eObj.label || eObj.full_name || eObj.name || 'Karyawan') : (typeof employee_id === 'string' && isNaN(Number(employee_id)) ? employee_id : 'Karyawan');

        const payload = {
          date,
          branch_id: Number(branch_id),
          branch_name: bObj ? bObj.label : 'Cabang Lembur',
          employee_id: Number(employee_id),
          employee_name: eObj ? eObj.label : 'Karyawan',
          start_time,
          end_time,
          break_hours,
          total_hours,
          reason
        };

        try {
          if (isEdit) {
            const res = await apiFetch(`/api/overtime/${item.id}`, {
              method: 'PUT',
              body: JSON.stringify(payload)
            });
            if (res.ok) {
              toastSuccess('Data lembur berhasil diperbarui');
              const idx = overtimeList.findIndex(o => o.id === item.id);
              if (idx !== -1) overtimeList[idx] = { ...overtimeList[idx], ...payload };
              close();
              renderUI();
            } else {
              toastError(res.data?.error || 'Gagal memperbarui data lembur');
            }
          } else {
            const res = await apiFetch('/api/overtime', {
              method: 'POST',
              body: JSON.stringify(payload)
            });
            if (res.ok) {
              toastSuccess('Data lembur berhasil ditambahkan');
              const newId = res.data?.id || Date.now();
              overtimeList.unshift({ id: newId, ...payload });
              close();
              renderUI();
            } else {
              toastError(res.data?.error || 'Gagal menambahkan lembur');
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
    return overtimeList.filter(o => {
      // Search
      if (searchQuery) {
        const q = searchQuery.toLowerCase();
        const matchEmp = (o.employee_name || '').toLowerCase().includes(q);
        const matchBranch = (o.branch_name || '').toLowerCase().includes(q);
        const matchReason = (o.reason || '').toLowerCase().includes(q);
        if (!matchEmp && !matchBranch && !matchReason) return false;
      }

      // Branch Filter
      if (selectedBranch !== 'all' && String(o.branch_id) !== String(selectedBranch)) {
        return false;
      }

      // Reason Filter
      if (selectedReason !== 'all' && o.reason !== selectedReason) {
        return false;
      }

      return true;
    });
  }

  // Main UI Renderer
  function renderUI() {
    const filtered = getFilteredList();

    // KPI Metrics calculation
    const totalRecords = overtimeList.length;
    const totalHoursAcc = overtimeList.reduce((sum, o) => sum + (parseFloat(o.total_hours) || 0), 0);
    const avgDuration = totalRecords > 0 ? (totalHoursAcc / totalRecords).toFixed(1) : 0;

    // Lembur Bulan Berjalan
    const now = new Date();
    const curYear = now.getFullYear();
    const curMonth = now.getMonth();
    const bulanIniCount = overtimeList.filter(o => {
      if (!o.date) return false;
      const d = new Date(o.date);
      return !isNaN(d.getTime()) && d.getFullYear() === curYear && d.getMonth() === curMonth;
    }).length;

    container.innerHTML = `
      <style>
        .lembur-page {
          display: flex;
          flex-direction: column;
          gap: 20px;
        }
        .lembur-hero-card {
          background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
          border-radius: 16px;
          padding: 24px 28px;
          color: white;
          box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25);
          position: relative;
          overflow: hidden;
        }
        .lembur-hero-card::after {
          content: '';
          position: absolute;
          top: -40px;
          right: -40px;
          width: 220px;
          height: 220px;
          background: radial-gradient(circle, rgba(99, 102, 241, 0.18) 0%, transparent 70%);
          border-radius: 50%;
          pointer-events: none;
        }
        .lembur-stats-row {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 14px;
          margin-top: 22px;
        }
        .lembur-stat-box {
          background: rgba(255, 255, 255, 0.06);
          border: 1px solid rgba(255, 255, 255, 0.1);
          border-radius: 12px;
          padding: 14px 16px;
          backdrop-filter: blur(8px);
          display: flex;
          align-items: center;
          gap: 14px;
        }
        .lembur-stat-icon {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 1.3rem;
          flex-shrink: 0;
        }
        .lembur-stat-val {
          font-size: 1.55rem;
          font-weight: 800;
          line-height: 1.1;
          color: #F8FAFC;
        }
        .lembur-stat-lbl {
          font-size: 0.76rem;
          text-transform: uppercase;
          letter-spacing: 0.05em;
          color: #94A3B8;
          font-weight: 600;
          margin-top: 2px;
        }

        /* Controls Card */
        .lembur-controls-card {
          background: white;
          border: 1px solid #E2E8F0;
          border-radius: 14px;
          padding: 18px 20px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.02);
          display: flex;
          flex-direction: column;
          gap: 16px;
        }
        .lembur-search-row {
          display: flex;
          gap: 12px;
          flex-wrap: wrap;
        }
        .lembur-search-wrap {
          position: relative;
          flex: 1;
          min-width: 240px;
        }
        .lembur-search-icon {
          position: absolute;
          left: 14px;
          top: 50%;
          transform: translateY(-50%);
          color: #94A3B8;
          pointer-events: none;
        }
        .lembur-search-input {
          width: 100%;
          padding: 10px 14px 10px 40px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          outline: none;
          transition: border-color 0.15s ease;
        }
        .lembur-search-input:focus {
          border-color: #2563EB;
          box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);
        }
        .lembur-select {
          padding: 10px 14px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          color: #1E293B;
          background: white;
          outline: none;
          min-width: 200px;
          cursor: pointer;
        }
        .lembur-select:focus {
          border-color: #2563EB;
        }

        /* Swipeable Pills */
        .lembur-filter-pills {
          display: flex;
          align-items: center;
          gap: 8px;
          overflow-x: auto;
          white-space: nowrap;
          scrollbar-width: none;
          padding-bottom: 2px;
        }
        .lembur-filter-pills::-webkit-scrollbar {
          display: none;
        }
        .lembur-pill {
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
        .lembur-pill:hover {
          background: #F1F5F9;
          color: #1E293B;
        }
        .lembur-pill.active {
          background: #2563EB;
          color: white;
          border-color: #2563EB;
          box-shadow: 0 2px 6px rgba(37, 99, 235, 0.25);
        }

        /* Grid Cards View */
        .lembur-cards-grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
          gap: 16px;
        }
        .lembur-card {
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
        .lembur-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 8px 18px rgba(0,0,0,0.06);
          border-color: #CBD5E1;
        }
        .lembur-avatar-box {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          background: linear-gradient(135deg, #4F46E5 0%, #3730A3 100%);
          color: white;
          font-weight: 700;
          font-size: 1rem;
          display: flex;
          align-items: center;
          justify-content: center;
          flex-shrink: 0;
          box-shadow: 0 3px 6px rgba(79, 70, 229, 0.25);
        }
        .lembur-time-box {
          background: #F8FAFC;
          border: 1px solid #E2E8F0;
          border-radius: 10px;
          padding: 10px 12px;
          margin: 12px 0;
          display: flex;
          align-items: center;
          justify-content: space-between;
        }
        .lembur-hours-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 5px 12px;
          border-radius: 8px;
          font-size: 0.88rem;
          font-weight: 800;
          background: #EEF2FF;
          color: #4F46E5;
          border: 1px solid #C7D2FE;
        }
        .lembur-reason-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.76rem;
          font-weight: 700;
        }
        .lembur-icon-btn {
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
        .lembur-icon-btn:hover {
          background: #F1F5F9;
          color: #0F172A;
          border-color: #CBD5E1;
        }
        .lembur-icon-btn.delete:hover {
          background: #FEE2E2;
          color: #DC2626;
          border-color: #FCA5A5;
        }

        /* Empty state */
        .lembur-empty-state {
          grid-column: 1 / -1;
          background: white;
          border: 2px dashed #E2E8F0;
          border-radius: 16px;
          padding: 48px 24px;
          text-align: center;
        }

        /* Mobile Responsive */
        @media (max-width: 768px) {
          .lembur-hero-card {
            padding: 18px 16px;
          }
          .lembur-stats-row {
            grid-template-columns: 1fr 1fr;
            gap: 10px;
          }
          .lembur-stat-box {
            padding: 10px 12px;
          }
          .lembur-stat-icon {
            width: 38px;
            height: 38px;
            font-size: 1.1rem;
          }
          .lembur-stat-val {
            font-size: 1.3rem;
          }
          .lembur-controls-card {
            padding: 14px;
            gap: 12px;
          }
          .lembur-search-row {
            flex-direction: column;
            gap: 10px;
          }
          .lembur-select {
            width: 100%;
          }
          .lembur-cards-grid {
            grid-template-columns: 1fr;
          }
        }
      </style>

      <div class="lembur-page">

        <!-- Hero Header Card -->
        <div class="lembur-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px; margin-bottom:6px;">
                <span style="font-size:1.6rem;">⏱️</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Monitoring Data Lembur Karyawan
                </h1>
              </div>
              <p style="margin:0; font-size:0.88rem; color:#CBD5E1; max-width:650px; line-height:1.5;">
                Pencatatan jam kerja lembur, approval penugasan overtime cabang, dan rekapitulasi durasi jam kerja PT Audy Mandiri Indonesia.
              </p>
            </div>

            <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
              ${canCreate ? `
                <button id="lembur-btn-add" class="btn btn-primary" style="background:#2563EB; padding:9px 18px; border-radius:10px; font-weight:700; display:flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3);">
                  <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Lembur</span>
                </button>
              ` : ''}
              
              <button id="lembur-btn-excel" class="btn btn-ghost" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; padding:9px 16px; border-radius:10px; font-weight:600; display:flex; align-items:center; gap:6px;">
                <span>📥</span>
                <span>Ekspor Excel</span>
              </button>

              <button id="lembur-btn-refresh" class="btn btn-ghost" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; padding:9px; border-radius:10px;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M23 4v6h-6M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- KPI Stats -->
          <div class="lembur-stats-row">
            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(79,70,229,0.25); color:#818CF8;">
                ⏱️
              </div>
              <div>
                <div class="lembur-stat-val">${totalHoursAcc} Jam</div>
                <div class="lembur-stat-lbl">Total Jam Lembur</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">
                📝
              </div>
              <div>
                <div class="lembur-stat-val">${totalRecords}</div>
                <div class="lembur-stat-lbl">Catatan Penugasan</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">
                📊
              </div>
              <div>
                <div class="lembur-stat-val">${avgDuration} Jam</div>
                <div class="lembur-stat-lbl">Rata-rata Durasi</div>
              </div>
            </div>

            <div class="lembur-stat-box">
              <div class="lembur-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                🗓️
              </div>
              <div>
                <div class="lembur-stat-val">${bulanIniCount}</div>
                <div class="lembur-stat-lbl">Bulan Berjalan</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Filter & Search Controls -->
        <div class="lembur-controls-card">
          <div class="lembur-search-row">
            <div class="lembur-search-wrap">
              <span class="lembur-search-icon">
                <svg width="17" height="17" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
              </span>
              <input type="text" id="lembur-search-input" class="lembur-search-input" placeholder="Cari nama karyawan, cabang, atau alasan lembur..." value="${searchQuery}" />
            </div>

            <select id="lembur-filter-branch" class="lembur-select">
              <option value="all">Semua Cabang (${branchOptions.length})</option>
              ${branchOptions.map(b => `<option value="${b.value}" ${String(selectedBranch) === String(b.value) ? 'selected' : ''}>${b.label}</option>`).join('')}
            </select>
          </div>

          <!-- Reason Filter Pills (Swipeable on Mobile) -->
          <div class="lembur-filter-pills">
            <div class="lembur-pill ${selectedReason === 'all' ? 'active' : ''}" data-reason="all">
              <span>Semua Lembur</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${totalRecords}</span>
            </div>
            ${OVERTIME_REASONS.map(r => {
              const count = overtimeList.filter(o => o.reason === r).length;
              const meta = getReasonBadge(r);
              return `
                <div class="lembur-pill ${selectedReason === r ? 'active' : ''}" data-reason="${r}">
                  <span>${meta.icon}</span>
                  <span>${r}</span>
                  ${count > 0 ? `<span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${count}</span>` : ''}
                </div>
              `;
            }).join('')}
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="lembur-list-container">
          ${filtered.length === 0 ? `
            <div class="lembur-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">🔍</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Data Lembur Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="lembur-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          ` : `
            <div class="lembur-cards-grid">
              ${filtered.map(o => {
                const initials = (o.employee_name || 'Karyawan').split(' ').map(w => w[0]).filter(Boolean).slice(0, 2).join('').toUpperCase();
                const reasonMeta = getReasonBadge(o.reason);
                return `
                  <div class="lembur-card" data-id="${o.id}">
                    <div>
                      <!-- Card Header: Jam & Istirahat -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="lembur-hours-badge">
                          <span>⏱️</span>
                          <span>${o.total_hours != null ? o.total_hours : 0} Jam Lembur</span>
                        </span>
                        <span style="font-size:0.8rem; font-weight:600; color:#64748B;">
                          ☕ Istirahat: ${o.break_hours != null ? o.break_hours : 0} Jam
                        </span>
                      </div>

                      <!-- Employee & Branch -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="lembur-avatar-box">
                          ${initials}
                        </div>
                        <div style="min-width:0;">
                          <h4 style="margin:0 0 3px 0; font-size:1.02rem; font-weight:700; color:#0F172A; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                            ${o.employee_name}
                          </h4>
                          <div style="font-size:0.8rem; color:#64748B; display:flex; align-items:center; gap:4px;">
                            <span>📍</span>
                            <span style="font-weight:600; color:#334155;">${o.branch_name || 'Cabang'}</span>
                          </div>
                        </div>
                      </div>

                      <!-- Shift Times & Date Box -->
                      <div class="lembur-time-box">
                        <div>
                          <div style="font-size:0.7rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:2px;">Tanggal</div>
                          <div style="font-size:0.88rem; font-weight:700; color:#1E293B;">
                            🗓️ ${formatDateIndo(o.date)}
                          </div>
                        </div>

                        <div style="text-align:right;">
                          <div style="font-size:0.7rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:2px;">Waktu Shift</div>
                          <div style="font-size:0.88rem; font-weight:700; color:#2563EB;">
                            ⏰ ${o.start_time || '00:00'} - ${o.end_time || '00:00'}
                          </div>
                        </div>
                      </div>

                      <!-- Reason Badge -->
                      <div>
                        <span class="lembur-reason-badge" style="background:${reasonMeta.bg}; color:${reasonMeta.color}; border:1px solid ${reasonMeta.border};">
                          <span>${reasonMeta.icon}</span>
                          <span>${o.reason || 'Lembur Operasional'}</span>
                        </span>
                      </div>
                    </div>

                    <!-- Card Footer: Aksi -->
                    ${canEdit ? `
                      <div style="display:flex; justify-content:flex-end; align-items:center; gap:6px; padding-top:12px; border-top:1px solid #F1F5F9; margin-top:12px;">
                        <button class="lembur-icon-btn lembur-edit-btn" title="Edit Catatan Lembur" data-id="${o.id}">
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        ${canDelete ? `
                          <button class="lembur-icon-btn delete lembur-delete-btn" title="Hapus Catatan Lembur" data-id="${o.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                          </button>
                        ` : ''}
                      </div>
                    ` : ''}

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
    const btnAdd = container.querySelector('#lembur-btn-add');
    if (btnAdd) btnAdd.onclick = () => openOvertimeModal();

    // Excel button
    const btnExcel = container.querySelector('#lembur-btn-excel');
    if (btnExcel) btnExcel.onclick = () => exportToExcel();

    // Refresh button
    const btnRefresh = container.querySelector('#lembur-btn-refresh');
    if (btnRefresh) btnRefresh.onclick = () => loadData();

    // Reset filter
    const btnReset = container.querySelector('#lembur-btn-reset-filter');
    if (btnReset) {
      btnReset.onclick = () => {
        searchQuery = '';
        selectedBranch = 'all';
        selectedReason = 'all';
        renderUI();
      };
    }

    // Search input
    const searchInput = container.querySelector('#lembur-search-input');
    if (searchInput) {
      searchInput.oninput = (e) => {
        searchQuery = e.target.value;
        renderUI();
        const newSearch = container.querySelector('#lembur-search-input');
        if (newSearch) {
          newSearch.focus();
          newSearch.setSelectionRange(newSearch.value.length, newSearch.value.length);
        }
      };
    }

    // Branch filter
    const branchFilter = container.querySelector('#lembur-filter-branch');
    if (branchFilter) {
      branchFilter.onchange = (e) => {
        selectedBranch = e.target.value;
        renderUI();
      };
    }

    // Reason Pills
    container.querySelectorAll('.lembur-pill').forEach(pill => {
      pill.onclick = () => {
        selectedReason = pill.getAttribute('data-reason') || 'all';
        renderUI();
      };
    });

    // Edit Buttons
    container.querySelectorAll('.lembur-edit-btn').forEach(btn => {
      btn.onclick = (e) => {
        e.stopPropagation();
        const id = Number(btn.getAttribute('data-id'));
        const item = overtimeList.find(o => o.id === id);
        if (item) openOvertimeModal(item);
      };
    });

    // Delete Buttons
    container.querySelectorAll('.lembur-delete-btn').forEach(btn => {
      btn.onclick = (e) => {
        e.stopPropagation();
        const id = Number(btn.getAttribute('data-id'));
        deleteOvertime(id);
      };
    });
  }

  // Initial load
  await loadData();
}
