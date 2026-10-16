import { apiFetch } from '../config.js';
import { getCachedBranches, getCachedEmployeeNames } from '../utils/dataCache.js';
import { createModal, confirmDialog } from '../components/modal.js';
import { toastSuccess, toastError, toastInfo } from '../components/toast.js';

// ── INITIAL PRE-SEEDED DATA SP (Sesuai Catatan Operasional FCMS) ────────────
const INITIAL_SP = [
  {
    id: 1,
    employee_name: 'Teguh Ariyandi',
    division: 'FACILITY CARE',
    branch_id: 18,
    branch_name: '018. Harapan Indah',
    tanggal: '2025-08-19',
    akhir_sp: '2026-02-18',
    sp_type: 'SP 1',
    status: 'Aktif',
    document_link: '',
    notes: 'Keterlambatan berulang dan ketidaklengkapan checklist harian lantai dasar.'
  },
  {
    id: 2,
    employee_name: 'Andri',
    division: 'FACILITY CARE',
    branch_id: 1,
    branch_name: '001. Pondok Bambu',
    tanggal: '2025-04-10',
    akhir_sp: '2025-10-09',
    sp_type: 'SP 1',
    status: 'Selesai',
    document_link: '',
    notes: 'Pelanggaran grooming dan standar pembersihan area tindakan tanpa APD lengkap.'
  },
  {
    id: 3,
    employee_name: 'M. FIRMANSYAH',
    division: 'FACILITY CARE',
    branch_id: 25,
    branch_name: '025. Bogor Baru',
    tanggal: '2024-03-21',
    akhir_sp: '2024-09-20',
    sp_type: 'SP 2',
    status: 'Selesai',
    document_link: '',
    notes: 'Meninggalkan area tugas tanpa konfirmasi supervisor selama jam operasional klinik.'
  }
];

// Helper kalkulasi status & sisa hari masa berlaku SP
function getSPStatusMeta(akhirSpDateStr) {
  if (!akhirSpDateStr) {
    return { status: 'Tidak Ditentukan', label: 'Tidak Ada Tanggal', bg: '#F1F5F9', color: '#64748B', border: '#CBD5E1', icon: '⚪', daysLeft: 0, isExpired: false, isExpiringSoon: false };
  }
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const target = new Date(akhirSpDateStr);
  target.setHours(0, 0, 0, 0);

  const diffTime = target.getTime() - today.getTime();
  const daysLeft = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

  if (daysLeft < 0) {
    const expiredDaysAgo = Math.abs(daysLeft);
    return {
      status: 'Selesai',
      label: `Masa Sanksi Selesai (${expiredDaysAgo} hari lalu)`,
      shortLabel: 'Selesai',
      bg: '#DCFCE7',
      color: '#15803D',
      border: '#BBF7D0',
      icon: '✅',
      daysLeft,
      isExpired: true,
      isExpiringSoon: false
    };
  } else if (daysLeft <= 30) {
    return {
      status: 'Mendekati Berakhir',
      label: `Segera Berakhir (Sisa ${daysLeft} hari)`,
      shortLabel: `Sisa ${daysLeft} hari`,
      bg: '#FEF3C7',
      color: '#B45309',
      border: '#FDE68A',
      icon: '⚠️',
      daysLeft,
      isExpired: false,
      isExpiringSoon: true
    };
  } else {
    return {
      status: 'Aktif',
      label: `Aktif Berlaku (Sisa ${daysLeft} hari)`,
      shortLabel: `Aktif (${daysLeft} hari)`,
      bg: '#FEE2E2',
      color: '#B91C1C',
      border: '#FECACA',
      icon: '🔴',
      daysLeft,
      isExpired: false,
      isExpiringSoon: false
    };
  }
}

// Helper badge jenis SP
function getSPTypeBadge(spType = '') {
  const t = (spType || '').toUpperCase().trim();
  if (t === 'SP 3') {
    return { label: 'SP 3 (Terakhir)', bg: '#7F1D1D', color: '#FFFFFF', border: '#991B1B', icon: '🚨' };
  }
  if (t === 'SP 2') {
    return { label: 'SP 2', bg: '#FEE2E2', color: '#B91C1C', border: '#FECACA', icon: '🛑' };
  }
  if (t === 'SP 1') {
    return { label: 'SP 1', bg: '#FEF3C7', color: '#B45309', border: '#FDE68A', icon: '⚠️' };
  }
  if (t.includes('TEGURAN')) {
    return { label: 'Teguran Lisan', bg: '#E0F2FE', color: '#0369A1', border: '#BAE6FD', icon: '💬' };
  }
  return { label: spType || 'SP', bg: '#F1F5F9', color: '#475569', border: '#CBD5E1', icon: '📜' };
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

export async function renderSP(container) {
  let spList = [];
  let branchOptions = [];
  let employeeOptions = [];
  let selectedFilter = 'all'; // 'all' | 'active' | 'expiring' | 'expired' | 'sp1' | 'sp2' | 'sp3' | 'teguran'
  let selectedBranch = 'all';
  let selectedDivision = 'all';
  let searchQuery = '';
  let currentView = localStorage.getItem('fcms_sp_view') || 'table'; // 'table' | 'grid'

  // Inject Styles
  const styleId = 'fcms-sp-page-styles';
  if (!document.getElementById(styleId)) {
    const styleEl = document.createElement('style');
    styleEl.id = styleId;
    styleEl.textContent = `
      .sp-page-container {
        padding: 24px;
        max-width: 1400px;
        margin: 0 auto;
        font-family: inherit;
      }
      .sp-hero-card {
        background: linear-gradient(135deg, #1E293B 0%, #0F172A 100%);
        border-radius: 18px;
        padding: 28px 32px;
        color: #FFFFFF;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.2);
        margin-bottom: 24px;
        position: relative;
        overflow: hidden;
      }
      .sp-hero-card::after {
        content: '';
        position: absolute;
        right: -50px;
        top: -50px;
        width: 220px;
        height: 220px;
        background: radial-gradient(circle, rgba(239, 68, 68, 0.15) 0%, transparent 70%);
        border-radius: 50%;
        pointer-events: none;
      }
      .sp-hero-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        gap: 16px;
      }
      .sp-hero-title {
        font-size: 1.65rem;
        font-weight: 700;
        margin: 0 0 6px 0;
        display: flex;
        align-items: center;
        gap: 12px;
        letter-spacing: -0.02em;
      }
      .sp-hero-subtitle {
        color: #94A3B8;
        font-size: 0.92rem;
        margin: 0;
        max-width: 720px;
        line-height: 1.5;
      }
      .sp-stats-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 14px;
        margin-top: 24px;
      }
      .sp-stat-box {
        background: rgba(255, 255, 255, 0.06);
        border: 1px solid rgba(255, 255, 255, 0.1);
        border-radius: 12px;
        padding: 14px 18px;
        backdrop-filter: blur(8px);
        display: flex;
        align-items: center;
        gap: 14px;
      }
      .sp-stat-icon {
        width: 44px;
        height: 44px;
        border-radius: 10px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.3rem;
        flex-shrink: 0;
      }
      .sp-stat-value {
        font-size: 1.45rem;
        font-weight: 700;
        color: #FFFFFF;
        line-height: 1.2;
      }
      .sp-stat-label {
        font-size: 0.78rem;
        color: #94A3B8;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        font-weight: 600;
      }
      .sp-controls-panel {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 18px 20px;
        margin-bottom: 24px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
        display: flex;
        flex-direction: column;
        gap: 16px;
      }
      .sp-controls-top {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
        flex-wrap: wrap;
      }
      .sp-search-wrapper {
        position: relative;
        flex: 1;
        min-width: 250px;
      }
      .sp-search-icon {
        position: absolute;
        left: 14px;
        top: 50%;
        transform: translateY(-50%);
        color: #94A3B8;
        font-size: 0.95rem;
        pointer-events: none;
      }
      .sp-search-input {
        width: 100%;
        padding: 10px 14px 10px 40px;
        border: 1.5px solid #CBD5E1;
        border-radius: 10px;
        font-size: 0.9rem;
        background: #F8FAFC;
        transition: all 0.2s ease;
      }
      .sp-search-input:focus {
        outline: none;
        border-color: #2563EB;
        background: #FFFFFF;
        box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.12);
      }
      .sp-filter-select {
        padding: 9px 14px;
        border: 1.5px solid #CBD5E1;
        border-radius: 10px;
        font-size: 0.88rem;
        background: #F8FAFC;
        color: #334155;
        font-weight: 500;
        cursor: pointer;
      }
      .sp-filter-select:focus {
        outline: none;
        border-color: #2563EB;
        background: #FFFFFF;
      }
      .sp-view-toggle {
        display: flex;
        background: #F1F5F9;
        padding: 4px;
        border-radius: 10px;
        gap: 4px;
      }
      .sp-view-btn {
        border: none;
        background: transparent;
        padding: 6px 14px;
        border-radius: 8px;
        font-size: 0.82rem;
        font-weight: 600;
        color: #64748B;
        cursor: pointer;
        display: flex;
        align-items: center;
        gap: 6px;
        transition: all 0.15s ease;
      }
      .sp-view-btn.active {
        background: #FFFFFF;
        color: #1E293B;
        box-shadow: 0 1px 4px rgba(0, 0, 0, 0.08);
      }
      .sp-filter-pills {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        align-items: center;
      }
      .sp-pill {
        border: 1px solid #E2E8F0;
        background: #F8FAFC;
        color: #475569;
        font-size: 0.82rem;
        font-weight: 600;
        padding: 6px 14px;
        border-radius: 30px;
        cursor: pointer;
        transition: all 0.2s ease;
        display: flex;
        align-items: center;
        gap: 6px;
      }
      .sp-pill:hover {
        background: #EEF2F6;
        border-color: #CBD5E1;
      }
      .sp-pill.active {
        background: #2563EB;
        border-color: #2563EB;
        color: #FFFFFF;
        box-shadow: 0 2px 6px rgba(37, 99, 235, 0.25);
      }
      .sp-pill-count {
        background: rgba(0, 0, 0, 0.08);
        padding: 1px 7px;
        border-radius: 12px;
        font-size: 0.75rem;
      }
      .sp-pill.active .sp-pill-count {
        background: rgba(255, 255, 255, 0.25);
        color: #FFFFFF;
      }
      /* Modern Table View */
      .sp-table-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
        overflow: hidden;
      }
      .sp-table {
        width: 100%;
        border-collapse: collapse;
        text-align: left;
      }
      .sp-table th {
        background: #F8FAFC;
        padding: 14px 18px;
        font-size: 0.78rem;
        font-weight: 700;
        color: #475569;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        border-bottom: 1.5px solid #E2E8F0;
      }
      .sp-table td {
        padding: 16px 18px;
        font-size: 0.88rem;
        color: #1E293B;
        border-bottom: 1px solid #F1F5F9;
        vertical-align: middle;
      }
      .sp-table tr:hover td {
        background: #F8FAFC;
      }
      /* Grid View */
      .sp-cards-grid {
        display: grid;
        grid-template-columns: repeat(auto-fill, minmax(350px, 1fr));
        gap: 20px;
      }
      .sp-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 22px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.04);
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
        position: relative;
      }
      .sp-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 12px 24px -6px rgba(15, 23, 42, 0.1);
        border-color: #CBD5E1;
      }
      .sp-status-badge {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 4px 10px;
        border-radius: 20px;
        font-size: 0.74rem;
        font-weight: 700;
        letter-spacing: 0.02em;
      }
      .sp-type-badge {
        display: inline-flex;
        align-items: center;
        gap: 4px;
        padding: 4px 9px;
        border-radius: 6px;
        font-size: 0.74rem;
        font-weight: 700;
      }
      .sp-avatar-box {
        width: 42px;
        height: 42px;
        border-radius: 10px;
        background: #EFF6FF;
        color: #2563EB;
        font-weight: 700;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.95rem;
        flex-shrink: 0;
      }
      .sp-icon-btn {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        color: #64748B;
        width: 32px;
        height: 32px;
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        cursor: pointer;
        transition: all 0.15s ease;
      }
      .sp-icon-btn:hover {
        background: #EDF2F7;
        color: #1E293B;
        border-color: #CBD5E1;
      }
      .sp-icon-btn.delete:hover {
        background: #FEE2E2;
        color: #DC2626;
        border-color: #FECACA;
      }
      .sp-empty-state {
        text-align: center;
        padding: 60px 20px;
        background: #FFFFFF;
        border-radius: 16px;
        border: 1px dashed #CBD5E1;
      }
    `;
    document.head.appendChild(styleEl);
  }

  // Load Data
  async function loadData() {
    try {
      const [branchRes, empRes] = await Promise.all([
        getCachedBranches(),
        getCachedEmployeeNames()
      ]);
      branchOptions = branchRes || [];
      employeeOptions = empRes || [];

      const res = await apiFetch('/api/sp?limit=1000');
      if (res.ok && res.data?.data && Array.isArray(res.data.data) && res.data.data.length > 0) {
        spList = res.data.data;
        localStorage.setItem('fcms_sp_data_cache', JSON.stringify(spList));
      } else {
        const cached = localStorage.getItem('fcms_sp_data_cache');
        if (cached) {
          spList = JSON.parse(cached);
        } else {
          spList = [...INITIAL_SP];
          localStorage.setItem('fcms_sp_data_cache', JSON.stringify(spList));
        }
      }
    } catch (err) {
      console.warn('Fallback loading SP data from cache/seed:', err);
      const cached = localStorage.getItem('fcms_sp_data_cache');
      spList = cached ? JSON.parse(cached) : [...INITIAL_SP];
    }
    renderUI();
  }

  // Filter & Search Logic
  function getFilteredSP() {
    return spList.filter(item => {
      // Search
      const q = searchQuery.toLowerCase().trim();
      const matchSearch = !q ||
        (item.employee_name || '').toLowerCase().includes(q) ||
        (item.branch_name || '').toLowerCase().includes(q) ||
        (item.division || '').toLowerCase().includes(q) ||
        (item.sp_type || '').toLowerCase().includes(q);

      // Branch filter
      const matchBranch = selectedBranch === 'all' || String(item.branch_id) === String(selectedBranch);

      // Division filter
      const matchDiv = selectedDivision === 'all' || item.division === selectedDivision;

      // Status/Level Filter
      const meta = getSPStatusMeta(item.akhir_sp);
      let matchFilter = true;
      if (selectedFilter === 'active') matchFilter = !meta.isExpired;
      else if (selectedFilter === 'expiring') matchFilter = meta.isExpiringSoon;
      else if (selectedFilter === 'expired') matchFilter = meta.isExpired;
      else if (selectedFilter === 'sp1') matchFilter = (item.sp_type || '').toUpperCase() === 'SP 1';
      else if (selectedFilter === 'sp2') matchFilter = (item.sp_type || '').toUpperCase() === 'SP 2';
      else if (selectedFilter === 'sp3') matchFilter = (item.sp_type || '').toUpperCase() === 'SP 3';
      else if (selectedFilter === 'teguran') matchFilter = (item.sp_type || '').toUpperCase().includes('TEGURAN');

      return matchSearch && matchBranch && matchDiv && matchFilter;
    });
  }

  // Calculate Metrics
  function getMetrics() {
    let activeCount = 0;
    let expiringCount = 0;
    let expiredCount = 0;
    let sp1Count = 0;
    let sp2Count = 0;
    let sp3Count = 0;
    let teguranCount = 0;
    const employeeSet = new Set();

    spList.forEach(item => {
      if (item.employee_name) employeeSet.add(item.employee_name.toLowerCase().trim());
      const meta = getSPStatusMeta(item.akhir_sp);
      if (meta.isExpired) expiredCount++;
      else activeCount++;

      if (meta.isExpiringSoon) expiringCount++;

      const t = (item.sp_type || '').toUpperCase();
      if (t === 'SP 1') sp1Count++;
      else if (t === 'SP 2') sp2Count++;
      else if (t === 'SP 3') sp3Count++;
      else if (t.includes('TEGURAN')) teguranCount++;
    });

    return {
      total: spList.length,
      active: activeCount,
      expiring: expiringCount,
      expired: expiredCount,
      sp1: sp1Count,
      sp2: sp2Count,
      sp3: sp3Count,
      teguran: teguranCount,
      uniqueEmployees: employeeSet.size
    };
  }

  // Modal Terbitkan / Edit SP
  function openSPModal(item = null) {
    const isEdit = !!item;
    const defaultTanggal = item?.tanggal ? item.tanggal.slice(0, 10) : new Date().toISOString().slice(0, 10);
    
    // Auto calculate akhir_sp default (+6 bulan)
    let defaultAkhir = item?.akhir_sp ? item.akhir_sp.slice(0, 10) : '';
    if (!defaultAkhir) {
      const d = new Date(defaultTanggal);
      d.setMonth(d.getMonth() + 6);
      defaultAkhir = d.toISOString().slice(0, 10);
    }

    const content = document.createElement('div');
    content.innerHTML = `
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <input type="text" id="sp-modal-emp" list="sp-emp-list" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Pilih atau ketik nama..." value="${item?.employee_name ? String(item.employee_name).replace(/"/g, '&quot;') : ''}" required />
            <datalist id="sp-emp-list">
              ${employeeOptions.map(e => `<option value="${e.label}">${e.label}</option>`).join('')}
            </datalist>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Divisi <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-div" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="FACILITY CARE" ${item?.division === 'FACILITY CARE' ? 'selected' : ''}>FACILITY CARE</option>
              <option value="SECURITY" ${item?.division === 'SECURITY' ? 'selected' : ''}>SECURITY</option>
            </select>
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Penempatan <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">Pilih Cabang...</option>
              ${branchOptions.map(b => `<option value="${b.value}" ${String(item?.branch_id) === String(b.value) ? 'selected' : ''}>${b.label}</option>`).join('')}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tingkat / Jenis SP <span style="color:#DC2626;">*</span>
            </label>
            <select id="sp-modal-type" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="SP 1" ${item?.sp_type === 'SP 1' ? 'selected' : ''}>SP 1 (Masa Berlaku 6 Bulan)</option>
              <option value="SP 2" ${item?.sp_type === 'SP 2' ? 'selected' : ''}>SP 2 (Masa Berlaku 6 Bulan)</option>
              <option value="SP 3" ${item?.sp_type === 'SP 3' ? 'selected' : ''}>SP 3 (Masa Berlaku 6 Bulan / Terakhir)</option>
              <option value="Teguran Lisan" ${item?.sp_type === 'Teguran Lisan' ? 'selected' : ''}>Teguran Lisan (Pembinaan)</option>
            </select>
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Terbit SP <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="sp-modal-tanggal" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${defaultTanggal}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Akhir Masa Berlaku SP <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="sp-modal-akhir" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${defaultAkhir}" required />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Link Berkas / Foto BAP / Scan Surat SP
          </label>
          <input type="url" id="sp-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="https://drive.google.com/..." value="${item?.document_link ? String(item.document_link).replace(/"/g, '&quot;') : ''}" />
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Kronologi Pelanggaran / Catatan Pembinaan
          </label>
          <textarea id="sp-modal-notes" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tuliskan ringkasan pelanggaran disiplin kerja yang terjadi...">${item?.notes || item?.description || ''}</textarea>
        </div>
      </div>
    `;

    // Auto update akhir SP when tanggal changes (+6 months)
    const tglEl = content.querySelector('#sp-modal-tanggal');
    const akhirEl = content.querySelector('#sp-modal-akhir');
    tglEl.addEventListener('change', () => {
      if (tglEl.value) {
        const d = new Date(tglEl.value);
        d.setMonth(d.getMonth() + 6);
        akhirEl.value = d.toISOString().slice(0, 10);
      }
    });

    createModal({
      title: isEdit ? '✏️ Edit Data Surat Peringatan (SP)' : '➕ Terbitkan Surat Peringatan (SP) Baru',
      content,
      size: 'md',
      confirmText: isEdit ? 'Simpan Perubahan' : 'Terbitkan SP',
      confirmClass: 'btn-primary',
      onConfirm: async (_, close) => {
        const employee_name = (content.querySelector('#sp-modal-emp').value || '').trim();
        const division = content.querySelector('#sp-modal-div').value;
        const branch_id = content.querySelector('#sp-modal-branch').value;
        const sp_type = content.querySelector('#sp-modal-type').value;
        const tanggal = content.querySelector('#sp-modal-tanggal').value;
        const akhir_sp = content.querySelector('#sp-modal-akhir').value;
        const document_link = (content.querySelector('#sp-modal-link').value || '').trim();
        const notes = (content.querySelector('#sp-modal-notes').value || '').trim();

        if (!employee_name) {
          toastError('Nama Karyawan wajib diisi');
          return;
        }
        if (!branch_id) {
          toastError('Cabang penempatan wajib dipilih');
          return;
        }
        if (!tanggal || !akhir_sp) {
          toastError('Tanggal SP dan Akhir SP wajib diisi');
          return;
        }

        const bObj = branchOptions.find(b => String(b.value) === String(branch_id));
        const branch_name = bObj ? bObj.label : 'Cabang';

        const payload = {
          employee_name,
          division,
          branch_id: Number(branch_id),
          branch_name,
          sp_type,
          tanggal,
          akhir_sp,
          document_link,
          notes,
          status: getSPStatusMeta(akhir_sp).status
        };

        try {
          if (isEdit) {
            const res = await apiFetch(`/api/sp/${item.id}`, {
              method: 'PUT',
              body: JSON.stringify(payload)
            });
            const idx = spList.findIndex(s => s.id === item.id);
            if (idx !== -1) spList[idx] = { ...spList[idx], ...payload };
            toastSuccess('Data SP berhasil diperbarui');
          } else {
            const res = await apiFetch('/api/sp', {
              method: 'POST',
              body: JSON.stringify(payload)
            });
            const newId = res.ok && res.data?.id ? res.data.id : Date.now();
            spList.unshift({ id: newId, ...payload });
            toastSuccess('Surat Peringatan baru berhasil diterbitkan');
          }

          localStorage.setItem('fcms_sp_data_cache', JSON.stringify(spList));
          close();
          renderUI();
        } catch (err) {
          console.error('Error saving SP:', err);
          toastError('Gagal menyimpan SP: ' + err.message);
        }
      }
    });
  }

  // Delete SP
  function deleteSP(item) {
    confirmDialog(
      `Hapus data Surat Peringatan untuk karyawan <strong>${item.employee_name}</strong> (${item.sp_type})?`,
      async () => {
        try {
          await apiFetch(`/api/sp/${item.id}`, { method: 'DELETE' });
        } catch (e) {
          console.warn('API delete error fallback to local:', e);
        }
        spList = spList.filter(s => s.id !== item.id);
        localStorage.setItem('fcms_sp_data_cache', JSON.stringify(spList));
        toastSuccess(`Data SP ${item.employee_name} berhasil dihapus`);
        renderUI();
      },
      'Hapus Data SP'
    );
  }

  // View Detail SP Modal
  function openSPDetail(item) {
    const meta = getSPStatusMeta(item.akhir_sp);
    const typeMeta = getSPTypeBadge(item.sp_type);

    const content = document.createElement('div');
    content.innerHTML = `
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div style="display:flex; align-items:center; gap:14px; padding-bottom:14px; border-bottom:1px solid #E2E8F0;">
          <div class="sp-avatar-box" style="width:54px; height:54px; font-size:1.2rem;">
            ${item.employee_name.slice(0, 2).toUpperCase()}
          </div>
          <div>
            <h3 style="margin:0 0 4px 0; font-size:1.15rem; font-weight:700; color:#0F172A;">
              ${item.employee_name}
            </h3>
            <div style="display:flex; align-items:center; gap:8px;">
              <span class="badge" style="background:#EFF6FF; color:#1D4ED8; font-weight:700;">${item.division || 'FACILITY CARE'}</span>
              <span style="color:#64748B; font-size:0.85rem;">📍 ${item.branch_name || 'Cabang'}</span>
            </div>
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:4px;">Tingkat Sanksi</div>
            <span class="sp-type-badge" style="background:${typeMeta.bg}; color:${typeMeta.color}; border:1px solid ${typeMeta.border}; font-size:0.85rem;">
              <span>${typeMeta.icon}</span>
              <span>${typeMeta.label}</span>
            </span>
          </div>

          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:4px;">Status Masa Berlaku</div>
            <span class="sp-status-badge" style="background:${meta.bg}; color:${meta.color}; border:1px solid ${meta.border}; font-size:0.85rem;">
              <span>${meta.icon}</span>
              <span>${meta.label}</span>
            </span>
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:2px;">Tanggal Terbit SP</div>
            <div style="font-size:0.95rem; font-weight:700; color:#1E293B;">${formatDateIndo(item.tanggal)}</div>
          </div>

          <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:12px;">
            <div style="font-size:0.75rem; color:#64748B; text-transform:uppercase; font-weight:600; margin-bottom:2px;">Akhir Masa Berlaku</div>
            <div style="font-size:0.95rem; font-weight:700; color:#1E293B;">${formatDateIndo(item.akhir_sp)}</div>
          </div>
        </div>

        ${item.notes || item.description ? `
          <div style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:10px; padding:14px;">
            <div style="font-size:0.78rem; font-weight:700; color:#92400E; margin-bottom:4px;">📝 Catatan Pelanggaran / Pembinaan:</div>
            <div style="font-size:0.88rem; color:#78350F; line-height:1.5;">${item.notes || item.description}</div>
          </div>
        ` : ''}

        ${item.document_link ? `
          <div>
            <a href="${item.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-primary" style="width:100%; justify-content:center; padding:10px; font-weight:600; display:flex; align-items:center; gap:8px;">
              <span>📄</span> Buka Dokumen BAP / Scan SP Fisik
            </a>
          </div>
        ` : `
          <div style="text-align:center; padding:8px; font-size:0.82rem; color:#94A3B8;">
            (Belum ada scan dokumen / BAP yang diunggah)
          </div>
        `}
      </div>
    `;

    createModal({
      title: '📋 Lembar Riwayat Surat Peringatan (SP)',
      content,
      size: 'md',
      cancelText: 'Tutup'
    });
  }

  // Export to Excel
  async function exportToExcel() {
    try {
      const { downloadExcel } = await import('../utils/excel.js');
      const exportData = spList.map((s, idx) => {
        const meta = getSPStatusMeta(s.akhir_sp);
        return {
          'No': idx + 1,
          'Nama Karyawan': s.employee_name || '',
          'Divisi': s.division || '',
          'Cabang': s.branch_name || '',
          'Jenis SP': s.sp_type || '',
          'Tanggal SP': s.tanggal || '',
          'Akhir SP': s.akhir_sp || '',
          'Status Masa Berlaku': meta.status,
          'Sisa Hari': meta.daysLeft > 0 ? meta.daysLeft : 0,
          'Link BAP / Foto': s.document_link || '',
          'Catatan / Kronologi': s.notes || s.description || ''
        };
      });
      downloadExcel(exportData, `Data_Surat_Peringatan_FCMS_${new Date().toISOString().slice(0, 10)}`);
      toastSuccess('Berhasil mengunduh data SP ke format Excel.');
    } catch (err) {
      toastError('Gagal ekspor: ' + err.message);
    }
  }

  // Render UI Function
  function renderUI() {
    const filtered = getFilteredSP();
    const metrics = getMetrics();

    container.innerHTML = `
      <div class="sp-page-container">
        
        <!-- Hero Header -->
        <div class="sp-hero-card">
          <div class="sp-hero-header">
            <div>
              <h1 class="sp-hero-title">
                <span>📜</span> Monitoring Data SP (Surat Peringatan)
              </h1>
              <p class="sp-hero-subtitle">
                Sistem pengawasan kedisiplinan kerja, masa aktif sanksi pembinaan, dan arsip dokumen BAP karyawan PT Audy Mandiri Indonesia.
              </p>
            </div>
            <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
              <button id="sp-btn-add" class="btn btn-primary" style="background:#2563EB; border:none; border-radius:10px; padding:10px 18px; font-weight:600; display:flex; align-items:center; gap:8px;">
                <span>➕</span> Terbitkan SP Baru
              </button>
              <button id="sp-btn-excel" class="btn btn-ghost" style="background:rgba(255,255,255,0.1); color:#FFFFFF; border:1px solid rgba(255,255,255,0.2); border-radius:10px; padding:10px 16px; font-weight:600; display:flex; align-items:center; gap:8px;">
                <span>📥</span> Ekspor Excel
              </button>
              <button id="sp-btn-refresh" class="btn btn-ghost" style="background:rgba(255,255,255,0.1); color:#FFFFFF; border:1px solid rgba(255,255,255,0.2); border-radius:10px; padding:10px 14px; font-weight:600;">
                <span>🔄</span>
              </button>
            </div>
          </div>

          <!-- Quick Metrics -->
          <div class="sp-stats-grid">
            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(239,68,68,0.2); color:#F87171;">🔴</div>
              <div>
                <div class="sp-stat-value">${metrics.active}</div>
                <div class="sp-stat-label">SP Aktif Berlaku</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">⚠️</div>
              <div>
                <div class="sp-stat-value">${metrics.expiring}</div>
                <div class="sp-stat-label">Segera Berakhir (&lt;30 Hari)</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">✅</div>
              <div>
                <div class="sp-stat-value">${metrics.expired}</div>
                <div class="sp-stat-label">Masa Sanksi Selesai</div>
              </div>
            </div>

            <div class="sp-stat-box">
              <div class="sp-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">👥</div>
              <div>
                <div class="sp-stat-value">${metrics.uniqueEmployees} <span style="font-size:0.8rem; font-weight:400; color:#94A3B8;">karyawan</span></div>
                <div class="sp-stat-label">Total Terkena Sanksi</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Controls Panel -->
        <div class="sp-controls-panel">
          <div class="sp-controls-top">
            <div class="sp-search-wrapper">
              <span class="sp-search-icon">🔍</span>
              <input type="text" id="sp-search-input" class="sp-search-input" placeholder="Cari nama karyawan, cabang, atau divisi..." value="${searchQuery ? searchQuery.replace(/"/g, '&quot;') : ''}" />
            </div>

            <!-- Branch Dropdown -->
            <select id="sp-filter-branch" class="sp-filter-select">
              <option value="all">Semua Cabang (${branchOptions.length})</option>
              ${branchOptions.map(b => `<option value="${b.value}" ${String(selectedBranch) === String(b.value) ? 'selected' : ''}>${b.label}</option>`).join('')}
            </select>

            <!-- Division Dropdown -->
            <select id="sp-filter-division" class="sp-filter-select">
              <option value="all" ${selectedDivision === 'all' ? 'selected' : ''}>Semua Divisi</option>
              <option value="FACILITY CARE" ${selectedDivision === 'FACILITY CARE' ? 'selected' : ''}>FACILITY CARE</option>
              <option value="SECURITY" ${selectedDivision === 'SECURITY' ? 'selected' : ''}>SECURITY</option>
            </select>

            <!-- View Switch -->
            <div class="sp-view-toggle">
              <button id="sp-view-table-btn" class="sp-view-btn ${currentView === 'table' ? 'active' : ''}">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>
                <span>Tabel</span>
              </button>
              <button id="sp-view-grid-btn" class="sp-view-btn ${currentView === 'grid' ? 'active' : ''}">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></svg>
                <span>Kartu</span>
              </button>
            </div>
          </div>

          <!-- Filter Pills -->
          <div class="sp-filter-pills">
            <button class="sp-pill ${selectedFilter === 'all' ? 'active' : ''}" data-filter="all">
              <span>Semua SP</span>
              <span class="sp-pill-count">${metrics.total}</span>
            </button>
            <button class="sp-pill ${selectedFilter === 'active' ? 'active' : ''}" data-filter="active">
              <span>🔴 SP Aktif</span>
              <span class="sp-pill-count">${metrics.active}</span>
            </button>
            <button class="sp-pill ${selectedFilter === 'expiring' ? 'active' : ''}" data-filter="expiring">
              <span>⚠️ &lt;30 Hari</span>
              <span class="sp-pill-count">${metrics.expiring}</span>
            </button>
            <button class="sp-pill ${selectedFilter === 'expired' ? 'active' : ''}" data-filter="expired">
              <span>✅ Selesai</span>
              <span class="sp-pill-count">${metrics.expired}</span>
            </button>
            <button class="sp-pill ${selectedFilter === 'sp1' ? 'active' : ''}" data-filter="sp1">
              <span>SP 1</span>
              <span class="sp-pill-count">${metrics.sp1}</span>
            </button>
            <button class="sp-pill ${selectedFilter === 'sp2' ? 'active' : ''}" data-filter="sp2">
              <span>SP 2</span>
              <span class="sp-pill-count">${metrics.sp2}</span>
            </button>
            <button class="sp-pill ${selectedFilter === 'sp3' ? 'active' : ''}" data-filter="sp3">
              <span>SP 3</span>
              <span class="sp-pill-count">${metrics.sp3}</span>
            </button>
          </div>
        </div>

        <!-- Content Area -->
        <div id="sp-list-container">
          ${filtered.length === 0 ? `
            <div class="sp-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">🔍</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Data SP Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="sp-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          ` : currentView === 'table' ? `
            <!-- Modern Table View -->
            <div class="sp-table-card">
              <table class="sp-table">
                <thead>
                  <tr>
                    <th style="width:40px; text-align:center;">No</th>
                    <th>Nama Karyawan</th>
                    <th>Divisi</th>
                    <th>Cabang</th>
                    <th>Tingkat SP</th>
                    <th>Tanggal Terbit</th>
                    <th>Akhir SP</th>
                    <th>Status Masa Berlaku</th>
                    <th>Dokumen</th>
                    <th style="width:100px; text-align:center;">Aksi</th>
                  </tr>
                </thead>
                <tbody>
                  ${filtered.map((s, idx) => {
                    const meta = getSPStatusMeta(s.akhir_sp);
                    const typeMeta = getSPTypeBadge(s.sp_type);
                    return `
                      <tr>
                        <td style="text-align:center; font-weight:600; color:#94A3B8;">${idx + 1}</td>
                        <td>
                          <div style="display:flex; align-items:center; gap:10px;">
                            <div class="sp-avatar-box">
                              ${s.employee_name.slice(0, 2).toUpperCase()}
                            </div>
                            <div>
                              <div style="font-weight:700; color:#0F172A; cursor:pointer;" class="sp-view-detail-link" data-id="${s.id}">
                                ${s.employee_name}
                              </div>
                              <div style="font-size:0.75rem; color:#64748B;">ID: #${s.id}</div>
                            </div>
                          </div>
                        </td>
                        <td>
                          <span class="badge" style="background:#EFF6FF; color:#1D4ED8; font-weight:700; font-size:0.74rem;">
                            ${s.division || 'FACILITY CARE'}
                          </span>
                        </td>
                        <td>
                          <span style="font-weight:600; color:#334155;">${s.branch_name || '-'}</span>
                        </td>
                        <td>
                          <span class="sp-type-badge" style="background:${typeMeta.bg}; color:${typeMeta.color}; border:1px solid ${typeMeta.border};">
                            <span>${typeMeta.icon}</span>
                            <span>${typeMeta.label}</span>
                          </span>
                        </td>
                        <td style="font-weight:600; color:#475569;">
                          ${formatDateIndo(s.tanggal)}
                        </td>
                        <td style="font-weight:700; color:#0F172A;">
                          ${formatDateIndo(s.akhir_sp)}
                        </td>
                        <td>
                          <span class="sp-status-badge" style="background:${meta.bg}; color:${meta.color}; border:1px solid ${meta.border};">
                            <span>${meta.icon}</span>
                            <span>${meta.shortLabel}</span>
                          </span>
                        </td>
                        <td>
                          ${s.document_link ? `
                            <a href="${s.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.78rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                              <span>📄 Buka</span>
                            </a>
                          ` : '<span style="color:#94A3B8; font-size:0.8rem;">-</span>'}
                        </td>
                        <td style="text-align:center;">
                          <div style="display:flex; justify-content:center; gap:6px;">
                            <button class="sp-icon-btn sp-detail-btn" title="Lihat Detail BAP" data-id="${s.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>
                            </button>
                            <button class="sp-icon-btn sp-edit-btn" title="Edit Data SP" data-id="${s.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                            </button>
                            <button class="sp-icon-btn delete sp-delete-btn" title="Hapus" data-id="${s.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                            </button>
                          </div>
                        </td>
                      </tr>
                    `;
                  }).join('')}
                </tbody>
              </table>
            </div>
          ` : `
            <!-- Grid Card View -->
            <div class="sp-cards-grid">
              ${filtered.map(s => {
                const meta = getSPStatusMeta(s.akhir_sp);
                const typeMeta = getSPTypeBadge(s.sp_type);
                return `
                  <div class="sp-card" data-id="${s.id}">
                    <div>
                      <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:14px;">
                        <span class="sp-type-badge" style="background:${typeMeta.bg}; color:${typeMeta.color}; border:1px solid ${typeMeta.border};">
                          <span>${typeMeta.icon}</span>
                          <span>${typeMeta.label}</span>
                        </span>
                        <span class="sp-status-badge" style="background:${meta.bg}; color:${meta.color}; border:1px solid ${meta.border};">
                          <span>${meta.icon}</span>
                          <span>${meta.shortLabel}</span>
                        </span>
                      </div>

                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:12px;">
                        <div class="sp-avatar-box" style="width:48px; height:48px; font-size:1.05rem;">
                          ${s.employee_name.slice(0, 2).toUpperCase()}
                        </div>
                        <div>
                          <h3 style="margin:0 0 3px 0; font-size:1.1rem; font-weight:700; color:#0F172A;">
                            ${s.employee_name}
                          </h3>
                          <div style="font-size:0.8rem; color:#64748B;">
                            📍 ${s.branch_name || 'Cabang'} • <span style="color:#2563EB; font-weight:600;">${s.division || 'FC'}</span>
                          </div>
                        </div>
                      </div>

                      <div style="background:#F8FAFC; border-radius:10px; padding:10px 12px; margin-bottom:14px; font-size:0.82rem; display:flex; justify-content:space-between;">
                        <div>
                          <div style="color:#64748B; font-size:0.72rem; text-transform:uppercase; font-weight:600;">Tanggal Terbit</div>
                          <div style="font-weight:700; color:#1E293B;">${formatDateIndo(s.tanggal)}</div>
                        </div>
                        <div style="text-align:right;">
                          <div style="color:#64748B; font-size:0.72rem; text-transform:uppercase; font-weight:600;">Akhir SP</div>
                          <div style="font-weight:700; color:#0F172A;">${formatDateIndo(s.akhir_sp)}</div>
                        </div>
                      </div>

                      <p style="font-size:0.83rem; color:#64748B; line-height:1.45; margin:0 0 16px 0; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical; overflow:hidden;">
                        ${s.notes || s.description || 'Pencatatan sanksi peringatan disiplin operasional.'}
                      </p>
                    </div>

                    <div style="padding-top:12px; border-top:1px solid #F1F5F9; display:flex; align-items:center; justify-content:space-between;">
                      <button class="btn btn-sm btn-ghost sp-detail-btn" style="border:1px solid #CBD5E1; border-radius:8px; font-weight:600; font-size:0.82rem; padding:6px 12px;" data-id="${s.id}">
                        👁️ Rincian BAP
                      </button>

                      <div style="display:flex; align-items:center; gap:6px;">
                        ${s.document_link ? `
                          <a href="${s.document_link}" target="_blank" rel="noopener noreferrer" class="sp-icon-btn" title="Buka Dokumen">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/><polyline points="15 3 21 3 21 9"/><line x1="10" y1="14" x2="21" y2="3"/></svg>
                          </a>
                        ` : ''}
                        <button class="sp-icon-btn sp-edit-btn" title="Edit Data SP" data-id="${s.id}">
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                        <button class="sp-icon-btn delete sp-delete-btn" title="Hapus" data-id="${s.id}">
                          <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                        </button>
                      </div>
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

  // Bind Listeners
  function bindEvents() {
    // Top Buttons
    const btnAdd = container.querySelector('#sp-btn-add');
    if (btnAdd) btnAdd.onclick = () => openSPModal();

    const btnExcel = container.querySelector('#sp-btn-excel');
    if (btnExcel) btnExcel.onclick = () => exportToExcel();

    const btnRefresh = container.querySelector('#sp-btn-refresh');
    if (btnRefresh) btnRefresh.onclick = () => loadData();

    const btnReset = container.querySelector('#sp-btn-reset-filter');
    if (btnReset) {
      btnReset.onclick = () => {
        searchQuery = '';
        selectedBranch = 'all';
        selectedDivision = 'all';
        selectedFilter = 'all';
        renderUI();
      };
    }

    // Search Input
    const searchInput = container.querySelector('#sp-search-input');
    if (searchInput) {
      searchInput.oninput = (e) => {
        searchQuery = e.target.value;
        renderUI();
        const newSearch = container.querySelector('#sp-search-input');
        if (newSearch) {
          newSearch.focus();
          newSearch.setSelectionRange(newSearch.value.length, newSearch.value.length);
        }
      };
    }

    // Branch Filter
    const branchFilter = container.querySelector('#sp-filter-branch');
    if (branchFilter) {
      branchFilter.onchange = (e) => {
        selectedBranch = e.target.value;
        renderUI();
      };
    }

    // Division Filter
    const divFilter = container.querySelector('#sp-filter-division');
    if (divFilter) {
      divFilter.onchange = (e) => {
        selectedDivision = e.target.value;
        renderUI();
      };
    }

    // View Switcher
    const btnTable = container.querySelector('#sp-view-table-btn');
    const btnGrid = container.querySelector('#sp-view-grid-btn');
    if (btnTable) {
      btnTable.onclick = () => {
        currentView = 'table';
        localStorage.setItem('fcms_sp_view', 'table');
        renderUI();
      };
    }
    if (btnGrid) {
      btnGrid.onclick = () => {
        currentView = 'grid';
        localStorage.setItem('fcms_sp_view', 'grid');
        renderUI();
      };
    }

    // Filter Pills
    container.querySelectorAll('.sp-pill').forEach(pill => {
      pill.onclick = () => {
        selectedFilter = pill.getAttribute('data-filter') || 'all';
        renderUI();
      };
    });

    // Detail Buttons / Links
    container.querySelectorAll('.sp-detail-btn, .sp-view-detail-link').forEach(el => {
      el.onclick = (e) => {
        e.stopPropagation();
        const id = Number(el.getAttribute('data-id'));
        const item = spList.find(s => s.id === id);
        if (item) openSPDetail(item);
      };
    });

    // Edit Buttons
    container.querySelectorAll('.sp-edit-btn').forEach(btn => {
      btn.onclick = (e) => {
        e.stopPropagation();
        const id = Number(btn.getAttribute('data-id'));
        const item = spList.find(s => s.id === id);
        if (item) openSPModal(item);
      };
    });

    // Delete Buttons
    container.querySelectorAll('.sp-delete-btn').forEach(btn => {
      btn.onclick = (e) => {
        e.stopPropagation();
        const id = Number(btn.getAttribute('data-id'));
        const item = spList.find(s => s.id === id);
        if (item) deleteSP(item);
      };
    });
  }

  // Initial Load
  await loadData();
}

