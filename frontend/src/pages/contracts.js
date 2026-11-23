import { apiFetch, getUser, hasPermission } from '../config.js';
import { getCachedBranches, getCachedEmployees, fetchAllPaginated, fetchExportRows, clearDataCache } from '../utils/dataCache.js';
import { createModal } from '../components/modal.js';
import { toastSuccess, toastError } from '../components/toast.js';
import { downloadExcel, parseExcel } from '../utils/excel.js';
import { createPagination } from '../components/table.js';
import { parseValidRowVersion, resolveModalRowVersion, handleOccSaveResponse } from '../utils/occ.js';
import { mountEmployeeCombobox } from '../components/employeeCombobox.js';

// Helper format tanggal lokal Indonesia
function formatDateIndo(dateStr) {
  if (!dateStr || dateStr === '-') return '-';
  if (String(dateStr).startsWith('2099')) return 'Tetap / PKWTT';
  try {
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    return d.toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' });
  } catch {
    return dateStr;
  }
}

// Cek apakah karyawan tetap (PKWTT)
function isTetap(dateStr) {
  if (!dateStr) return false;
  return String(dateStr).startsWith('2099');
}

// Inisial avatar
function getInitials(name) {
  if (!name || name === '-') return '👤';
  const parts = String(name).trim().split(/\s+/);
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return (parts[0][0] + parts[1][0]).toUpperCase();
}

// Filter function for Dashboard integration
export function filterDashboardItem(s, type) {
  const status = String(s.status || '').toLowerCase();
  if (status !== 'aktif') return false;
  
  if (type === 'active') {
    return true; // All contracts with status 'Aktif' are considered active
  }
  
  if (type === 'expiring30') {
    if (!s.end_date || String(s.end_date).startsWith('2099')) return false;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const next30 = new Date(today);
    next30.setDate(today.getDate() + 30);
    const end = new Date(s.end_date);
    end.setHours(0, 0, 0, 0);
    return end >= today && end <= next30;
  }
  return false;
}

export async function renderContracts(container, params) {
  const PAGE_LIMIT = 20;
  let contractsList = [];
  let branchOptions = [];
  let employeeOptions = [];
  let employeePhoneMap = {};

  let searchQuery = '';
  let selectedBranch = 'all';
  let selectedDivision = 'all';
  let selectedStatus = 'all';
  let selectedExpiring = 'all';
  let selectedMonthExpiry = '';
  let currentPage = 1;
  let totalItems = 0;
  let totalPages = 1;
  let statsLoaded = false;
  let serverStats = { total: 0, active: 0, expiring30: 0, permanent: 0, nonactive: 0 };
  let searchDebounce = null;
  let isLoading = false;
  let viewMode = localStorage.getItem('fcms_contracts_view_mode') || 'card';

  // Handle URL dashboard parameter
  const dashFilter = params ? params.get('dash_filter') : null;
  const monthExpiryParam = params ? (params.get('month_expiry') || '').trim() : '';
  if (monthExpiryParam && /^\d{4}-(0[1-9]|1[0-2])$/.test(monthExpiryParam)) {
    selectedMonthExpiry = monthExpiryParam;
    selectedStatus = 'Aktif';
  } else if (dashFilter === 'expiring30') {
    selectedExpiring = '30';
    selectedStatus = 'Aktif';
  } else if (dashFilter === 'active') {
    selectedStatus = 'Aktif';
  }

  const currentUser = getUser();
  const canEdit = hasPermission(currentUser, 'contracts', 'write');
  const canAdd = hasPermission(currentUser, 'contracts', 'write');
  const canDelete = hasPermission(currentUser, 'contracts', 'delete');

  // Load option data
  try {
    const [branches, employees] = await Promise.all([
      getCachedBranches(),
      getCachedEmployees()
    ]);
    const allBranches = Array.isArray(branches) ? branches : [];
    branchOptions = allBranches.filter(
      b => String(b.code || '').trim() === '000' || b.branch_scope === 'hq' || /^(?:cabang\s+)?000\.\s*management$/i.test(String(b.label || b.full_name || '').trim())
    );
    employeeOptions = Array.isArray(employees) ? employees : [];

    // Map phone numbers if present in selector options (fallback; main list gets employee_phone from API)
    employeeOptions.forEach(emp => {
      const id = emp.value !== undefined ? emp.value : emp.id;
      const name = emp.label || emp.full_name;
      if (id && emp.phone) {
        employeePhoneMap[id] = emp.phone;
      }
      if (name && emp.phone) {
        employeePhoneMap[name.toLowerCase()] = emp.phone;
      }
    });
  } catch (err) {
    console.error('Error fetching cache options for contracts:', err);
  }

  function getActiveFilters() {
    const filters = { sort: 'end_date_asc' };
    if (searchQuery.trim()) filters.search = searchQuery.trim();
    if (selectedBranch && selectedBranch !== 'all') filters.branch_id = selectedBranch;
    if (selectedDivision && selectedDivision !== 'all') filters.division = selectedDivision;
    if (selectedStatus && selectedStatus !== 'all') filters.status = selectedStatus;
    if (selectedExpiring && selectedExpiring !== 'all') filters.expiring_days = selectedExpiring;
    if (selectedMonthExpiry) filters.month_expiry = selectedMonthExpiry;
    return filters;
  }

  function updateStatsUI() {
    const set = (id, val) => {
      const el = container.querySelector(id);
      if (el) el.textContent = val;
    };
    set('#con-stat-total', serverStats.total);
    set('#con-stat-active', serverStats.active);
    set('#con-stat-expiring', serverStats.expiring30);
    set('#con-stat-permanent', serverStats.permanent);
    set('#con-pill-count-all', serverStats.total);
    set('#con-pill-count-aktif', serverStats.active);
    set('#con-pill-count-expiring', serverStats.expiring30);
    set('#con-pill-count-tetap', serverStats.permanent);
    set('#con-pill-count-nonaktif', serverStats.nonactive);
  }

  // Load Contracts data from API (Server-Side Pagination)
  async function loadData({ refreshStats = true } = {}) {
    isLoading = true;
    const hasShell = Boolean(container.querySelector('#con-list-container'));
    if (!hasShell) {
      renderUI();
    } else {
      const listEl = container.querySelector('#con-list-container');
      if (listEl) {
        listEl.innerHTML = '<div class="con-empty-state"><p style="color:#64748B; margin:0;">Memuat data kontrak...</p></div>';
      }
    }
    try {
      const qs = new URLSearchParams({
        page: String(currentPage),
        limit: String(PAGE_LIMIT),
        ...getActiveFilters(),
      });
      if (refreshStats || !statsLoaded) {
        qs.set('include_stats', '1');
      }
      const res = await apiFetch(`/api/contracts?${qs.toString()}`);
      if (res.ok) {
        const raw = res.data?.data || [];
        contractsList = Array.isArray(raw) ? raw : [];
        const pag = res.data?.pagination || {};
        totalItems = Number(pag.total ?? contractsList.length);
        totalPages = Math.max(1, Number(pag.pages ?? Math.ceil(totalItems / PAGE_LIMIT)) || 1);
        currentPage = Math.max(1, Number(pag.page ?? currentPage));
        if (res.data?.stats) {
          serverStats = res.data.stats;
          statsLoaded = true;
        }
        if (contractsList.length === 0 && currentPage > totalPages) {
          currentPage = totalPages;
          isLoading = false;
          return loadData({ refreshStats });
        }
      } else {
        contractsList = [];
        totalItems = 0;
        totalPages = 1;
        toastError('Gagal memuat data kontrak dari server.');
      }
    } catch (e) {
      console.error('Error load contracts:', e);
      contractsList = [];
      totalItems = 0;
      totalPages = 1;
      toastError('Terjadi kesalahan koneksi saat memuat kontrak.');
    } finally {
      isLoading = false;
      if (container.querySelector('#con-list-container')) {
        updateStatsUI();
        renderListContent();
      } else {
        renderUI();
      }
    }
  }

  // Filtered List calculation (server already filters & paginates)
  function getFilteredList() {
    return contractsList;
  }

  // Cek Selisih Karyawan (Karyawan Aktif yang tidak punya Kontrak Aktif)
  async function checkMissingContracts() {
    createModal({
      title: '🔍 Analisis Selisih Kontrak Karyawan',
      content: `
        <div style="text-align:center; padding:24px 10px;">
          <div style="font-size:2rem; margin-bottom:10px;">⏳</div>
          <p style="color:#64748B; font-size:0.9rem;">Sedang menyinkronkan data karyawan aktif dengan kontrak aktif di database...</p>
        </div>
      `,
      cancelText: 'Tutup'
    });

    try {
      const [activeEmployees, allContracts] = await Promise.all([
        fetchAllPaginated('/api/employees/selector', { status: 'Aktif' }),
        fetchAllPaginated('/api/contracts', {})
      ]);

      const activeContractEmpIds = new Set(
        allContracts.filter(c => c.status === 'Aktif').map(c => c.employee_id)
      );

      const missing = activeEmployees.filter(e => !activeContractEmpIds.has(e.id));

      let contentHtml = '';
      if (missing.length === 0) {
        contentHtml = `
          <div style="text-align:center; padding:20px 10px;">
            <div style="font-size:2.8rem; margin-bottom:10px;">🎉</div>
            <h4 style="font-size:1.1rem; font-weight:800; color:#10B981; margin:0 0 8px 0;">Semua Karyawan Aktif Memiliki Kontrak!</h4>
            <p style="font-size:0.86rem; color:#64748B; margin:0;">
              Total <strong>${activeEmployees.length}</strong> karyawan aktif tercatat memiliki perikatan kontrak aktif yang sah di sistem.
            </p>
          </div>
        `;
      } else {
        contentHtml = `
          <div style="display:flex; flex-direction:column; gap:12px;">
            <div style="background:#FEF2F2; border:1px solid #FECACA; border-radius:10px; padding:12px; font-size:0.84rem; color:#991B1B;">
              Ditemukan <strong>${missing.length}</strong> dari <strong>${activeEmployees.length}</strong> karyawan aktif yang <strong>belum memiliki Kontrak Aktif</strong>.
            </div>
            <div style="max-height:360px; overflow-y:auto; display:flex; flex-direction:column; gap:8px; padding-right:4px;">
              ${missing.map(m => {
                const pastContracts = allContracts.filter(c => c.employee_id === m.id);
                let historyText = '<span style="color:#D97706; font-weight:600;">Belum pernah ada data kontrak</span>';
                if (pastContracts.length > 0) {
                  const last = pastContracts[0];
                  historyText = `Riwayat: Status <strong>${last.status}</strong> (Selesai: ${formatDateIndo(last.end_date)})`;
                }
                return `
                  <div style="background:#111827; border:1px solid #29364B; border-radius:8px; padding:10px 12px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px;">
                    <div>
                      <div style="font-weight:700; font-size:0.88rem; color:#0F172A;">${m.full_name}</div>
                      <div style="font-size:0.76rem; color:#64748B;">
                        Cabang: ${m.branch_name || '-'} • Divisi: ${m.division || 'FACILITY CARE'}
                      </div>
                      <div style="font-size:0.75rem; color:#475569; margin-top:2px;">
                        ${historyText}
                      </div>
                    </div>
                    ${canAdd ? `
                      <button class="btn btn-primary btn-add-contract-for-emp" data-empid="${m.id}" data-empname="${String(m.full_name).replace(/"/g, '&quot;')}" data-branchid="${m.branch_id || ''}" data-division="${m.division || 'FACILITY CARE'}" style="font-size:0.76rem; padding:5px 10px; border-radius:6px; background:#2563EB;">
                        + Buat Kontrak
                      </button>
                    ` : ''}
                  </div>
                `;
              }).join('')}
            </div>
          </div>
        `;
      }

      createModal({
        title: 'Analisis Selisih Kontrak Karyawan',
        content: contentHtml,
        cancelText: 'Tutup'
      });

      // Bind dynamic create buttons inside missing modal
      setTimeout(() => {
        document.querySelectorAll('.btn-add-contract-for-emp').forEach(btn => {
          btn.addEventListener('click', (e) => {
            const b = e.currentTarget;
            const empId = b.getAttribute('data-empid');
            const empName = b.getAttribute('data-empname');
            const branchId = b.getAttribute('data-branchid');
            const division = b.getAttribute('data-division');
            // Close active modal
            const closeBtn = document.querySelector('.modal-close');
            if (closeBtn) closeBtn.click();
            // Open form modal prefilled
            openFormModal({
              employee_id: empId,
              employee_name: empName,
              branch_id: branchId,
              division: division,
              status: 'Aktif'
            });
          });
        });
      }, 100);

    } catch (err) {
      console.error(err);
      toastError('Gagal melakukan analisis selisih data kontrak.');
    }
  }

  // Helper hitung otomatis tanggal selesai kontrak dari tanggal mulai & durasi
  function computeContractEndDate(startDateStr, contractType) {
    if (!startDateStr) return '';
    const parts = String(startDateStr).split('-').map(Number);
    if (parts.length !== 3 || !parts[0] || !parts[1] || !parts[2]) return '';
    const dt = new Date(parts[0], parts[1] - 1, parts[2]);
    if (isNaN(dt.getTime())) return '';

    if (contractType === 'KONTRAK 6 BULAN') {
      dt.setMonth(dt.getMonth() + 6);
      dt.setDate(dt.getDate() - 1);
    } else if (contractType === 'KONTRAK 1 TAHUN') {
      dt.setFullYear(dt.getFullYear() + 1);
      dt.setDate(dt.getDate() - 1);
    } else if (contractType === 'KONTRAK 2 TAHUN') {
      dt.setFullYear(dt.getFullYear() + 2);
      dt.setDate(dt.getDate() - 1);
    } else {
      return '';
    }
    const resY = dt.getFullYear();
    const resM = String(dt.getMonth() + 1).padStart(2, '0');
    const resD = String(dt.getDate()).padStart(2, '0');
    return `${resY}-${resM}-${resD}`;
  }

  // Open Add/Edit Contract Modal
  function openFormModal(item = null) {
    const isEditCheck = !!(item && item.id);
    if (isEditCheck && !canEdit) {
      toastError('Anda tidak memiliki izin untuk mengedit kontrak.');
      return;
    }
    if (!isEditCheck && !canAdd) {
      toastError('Anda tidak memiliki izin untuk menambah kontrak.');
      return;
    }
    const isEdit = !!(item && item.id);
    const editRowVersion = isEdit ? parseValidRowVersion(item?.row_version) : null;
    const modalTitle = isEdit ? `✏️ Edit Kontrak: ${item.employee_name || 'Karyawan'}` : '➕ Tambah Kontrak Karyawan Baru';

    const defaultEmpId = item?.employee_id || '';
    const defaultBranchId = item?.branch_id || '';
    const defaultDivision = item?.division || 'FACILITY CARE';
    const defaultStatus = item?.status || 'Aktif';
    const defaultContractType = item?.contract_type || 'KONTRAK 1 TAHUN';
    const defaultPkwt = item?.pkwt_number || 'PKWT 1';
    const defaultNotes = item?.notes || '';

    // Penentuan default isTetap secara cerdas:
    // Jika jenis kontrak atau nomor PKWT secara eksplisit menunjukkan PKWT (6 Bulan, 1 Tahun, 2 Tahun, PKWT 1-6),
    // maka status default adalah PKWT (bukan Tetap), meskipun di database tanggal sebelumnya tersimpan 2099-12-31 karena default import/migrasi.
    let defaultIsTetap = false;
    if (item) {
      if (item.contract_type === 'PKWTT (TETAP)' || item.pkwt_number === 'TETAP') {
        defaultIsTetap = true;
      } else if (item.contract_type && (item.contract_type.startsWith('KONTRAK') || item.contract_type.includes('BULAN') || item.contract_type.includes('TAHUN'))) {
        defaultIsTetap = false;
      } else if (item.pkwt_number && item.pkwt_number.startsWith('PKWT')) {
        defaultIsTetap = false;
      } else {
        defaultIsTetap = isTetap(item.end_date);
      }
    }

    const defaultStartDate = item?.start_date ? String(item.start_date).slice(0, 10) : '';
    let defaultEndDate = (item?.end_date && !isTetap(item.end_date)) ? String(item.end_date).slice(0, 10) : '';

    // Jika PKWT, tanggal mulai tersedia, tapi tanggal selesai kosong, hitung otomatis sebagai saran
    if (!defaultIsTetap && defaultStartDate && !defaultEndDate) {
      defaultEndDate = computeContractEndDate(defaultStartDate, defaultContractType);
    }

    const contentHtml = `
      <div style="display:flex; flex-direction:column; gap:14px;">
        ${isEdit ? `<input type="hidden" id="con-modal-row-version" value="${editRowVersion !== null ? editRowVersion : ''}">` : ''}
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <div id="con-modal-emp-mount"></div>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Penempatan <span style="color:#DC2626;">*</span>
            </label>
            <select id="con-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${branchOptions.map(b => `
                <option value="${b.value}" ${String(defaultBranchId) === String(b.value) ? 'selected' : ''}>
                  ${b.label}
                </option>
              `).join('')}
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Divisi / Bagian <span style="color:#DC2626;">*</span>
            </label>
            <select id="con-modal-division" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="FACILITY CARE" ${defaultDivision === 'FACILITY CARE' ? 'selected' : ''}>FACILITY CARE</option>
              <option value="SECURITY" ${defaultDivision === 'SECURITY' ? 'selected' : ''}>SECURITY</option>
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Kontrak <span style="color:#DC2626;">*</span>
            </label>
            <select id="con-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="Aktif" ${defaultStatus === 'Aktif' ? 'selected' : ''}>🟢 Aktif</option>
              <option value="Tidak Aktif" ${defaultStatus === 'Tidak Aktif' ? 'selected' : ''}>⚪ Tidak Aktif</option>
              <option value="Resign" ${defaultStatus === 'Resign' ? 'selected' : ''}>🔴 Resign</option>
              <option value="Cut" ${defaultStatus === 'Cut' ? 'selected' : ''}>✂️ Cut / Pemutusan</option>
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tipe / Durasi Kontrak
            </label>
            <select id="con-modal-type" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;">
              <option value="KONTRAK 6 BULAN" ${defaultContractType === 'KONTRAK 6 BULAN' ? 'selected' : ''}>KONTRAK 6 BULAN</option>
              <option value="KONTRAK 1 TAHUN" ${defaultContractType === 'KONTRAK 1 TAHUN' ? 'selected' : ''}>KONTRAK 1 TAHUN</option>
              <option value="KONTRAK 2 TAHUN" ${defaultContractType === 'KONTRAK 2 TAHUN' ? 'selected' : ''}>KONTRAK 2 TAHUN</option>
              <option value="PKWTT (TETAP)" ${defaultContractType === 'PKWTT (TETAP)' ? 'selected' : ''}>PKWTT (TETAP)</option>
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tingkat / No. PKWT
            </label>
            <select id="con-modal-pkwt" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;">
              <option value="PKWT 1" ${defaultPkwt === 'PKWT 1' ? 'selected' : ''}>PKWT 1</option>
              <option value="PKWT 2" ${defaultPkwt === 'PKWT 2' ? 'selected' : ''}>PKWT 2</option>
              <option value="PKWT 3" ${defaultPkwt === 'PKWT 3' ? 'selected' : ''}>PKWT 3</option>
              <option value="PKWT 4" ${defaultPkwt === 'PKWT 4' ? 'selected' : ''}>PKWT 4</option>
              <option value="PKWT 5" ${defaultPkwt === 'PKWT 5' ? 'selected' : ''}>PKWT 5</option>
              <option value="PKWT 6" ${defaultPkwt === 'PKWT 6' ? 'selected' : ''}>PKWT 6</option>
              <option value="TETAP" ${defaultPkwt === 'TETAP' ? 'selected' : ''}>PKWTT / TETAP</option>
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Mulai Kontrak
            </label>
            <input type="date" id="con-modal-start" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${defaultStartDate}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Selesai Kontrak
            </label>
            <input type="date" id="con-modal-end" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; ${defaultIsTetap ? 'opacity:0.5; background-color:#1B2940; cursor:not-allowed;' : 'background-color:#1B2940;'}" 
              value="${defaultEndDate}" ${defaultIsTetap ? 'disabled' : ''} />
            <div style="margin-top:6px; display:flex; align-items:center; gap:6px;">
              <input type="checkbox" id="con-modal-is-tetap" ${defaultIsTetap ? 'checked' : ''} style="cursor:pointer;" />
              <label for="con-modal-is-tetap" style="font-size:0.78rem; color:#475569; cursor:pointer; user-select:none;">
                Karyawan Tetap / PKWTT (Tanpa Batas Tanggal)
              </label>
            </div>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Catatan Tambahan
          </label>
          <textarea id="con-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; resize:vertical;" 
            placeholder="Keterangan perpanjangan, klausul khusus, atau nomor SK...">${defaultNotes}</textarea>
        </div>
      </div>
    `;

    const contentEl = document.createElement('div');
    contentEl.innerHTML = contentHtml;

    const empCombobox = mountEmployeeCombobox(contentEl.querySelector('#con-modal-emp-mount'), {
      mode: 'single',
      inputId: 'con-modal-employee',
      inputName: 'employee_name',
      placeholder: 'Cari nama karyawan...',
      required: true,
      allowManual: true,
      hasEmployeeIdSchema: true,
      manualBadgeText: 'Belum terhubung ke Master Karyawan',
      employees: employeeOptions,
      initialEmployeeId: defaultEmpId || null,
      initialEmployeeName: item?.employee_name || '',
      onChange: (sel) => {
        if (!isEdit && sel?.employee_id) {
          const empRow = employeeOptions.find(e => Number(e.value ?? e.id) === Number(sel.employee_id));
          if (empRow) {
            const branchSel = contentEl.querySelector('#con-modal-branch');
            const divSel = contentEl.querySelector('#con-modal-division');
            if (branchSel && !branchSel.value && empRow.branch_id) branchSel.value = String(empRow.branch_id);
            if (divSel && empRow.division) divSel.value = empRow.division;
          }
        }
      },
    });

    let isSubmitting = false;

    const modalObj = createModal({
      title: modalTitle,
      content: contentEl,
      confirmText: isEdit ? '💾 Simpan Perubahan' : '➕ Tambah Kontrak',
      cancelText: 'Batal',
      onConfirm: async (modalEl, close) => {
        if (isSubmitting) return;
        const empSelection = empCombobox.getValue();
        const empEl = empCombobox.inputEl;
        const branchEl = modalEl.querySelector('#con-modal-branch');
        const divEl = modalEl.querySelector('#con-modal-division');
        const statusEl = modalEl.querySelector('#con-modal-status');
        const startEl = modalEl.querySelector('#con-modal-start');
        const endEl = modalEl.querySelector('#con-modal-end');
        const isTetapEl = modalEl.querySelector('#con-modal-is-tetap');
        const typeEl = modalEl.querySelector('#con-modal-type');
        const pkwtEl = modalEl.querySelector('#con-modal-pkwt');
        const notesEl = modalEl.querySelector('#con-modal-notes');

        const employee_id = empSelection.employee_id;
        const employee_name = (empSelection.employee_name || '').trim();
        const branch_id = branchEl?.value;
        const division = divEl?.value || 'FACILITY CARE';
        const status = statusEl?.value || 'Aktif';
        const rawStartDate = startEl?.value?.trim();
        const start_date = rawStartDate || null;
        const is_tetap = isTetapEl?.checked;
        const rawEndDate = endEl?.value?.trim();
        const end_date = is_tetap ? '2099-12-31' : (rawEndDate || null);
        const contract_type = typeEl?.value;
        const pkwt_number = pkwtEl?.value;
        const notes = notesEl?.value?.trim() || null;

        if (!employee_name) {
          toastError('Harap pilih atau ketik nama karyawan!');
          empEl?.focus();
          return;
        }
        if (!branch_id) {
          toastError('Harap pilih cabang penempatan!');
          branchEl?.focus();
          return;
        }
        if (!is_tetap && start_date && end_date && end_date < start_date) {
          toastError('Tanggal selesai kontrak tidak boleh lebih awal dari tanggal mulai kontrak!');
          endEl?.focus();
          return;
        }

        let rowVersion = null;
        if (isEdit) {
          rowVersion = resolveModalRowVersion(modalEl, '#con-modal-row-version', item?.row_version);
          if (rowVersion === null) {
            toastError('Versi data kontrak (row_version) tidak tersedia. Silakan muat ulang halaman lalu coba lagi.');
            return;
          }
        }

        const payload = {
          employee_id: employee_id ? parseInt(employee_id, 10) : null,
          employee_name,
          branch_id: parseInt(branch_id, 10),
          division,
          status,
          start_date,
          end_date,
          contract_type,
          pkwt_number,
          notes,
          ...(isEdit ? { row_version: rowVersion } : {})
        };

        isSubmitting = true;
        try {
          const url = isEdit ? `/api/contracts/${item.id}` : '/api/contracts';
          const method = isEdit ? 'PUT' : 'POST';

          const res = await apiFetch(url, {
            method,
            body: JSON.stringify(payload)
          });

          await handleOccSaveResponse({
            res,
            isEdit,
            item,
            close,
            reloadData: async () => {
              clearDataCache();
              await loadData();
            },
            editSuccessMessage: 'Data kontrak berhasil diperbarui!',
            createSuccessMessage: 'Kontrak baru berhasil ditambahkan!',
            defaultErrorMessage: 'Gagal menyimpan data kontrak.',
          });
        } catch (err) {
          console.error(err);
          toastError('Terjadi kesalahan koneksi ke server.');
        } finally {
          isSubmitting = false;
        }
      }
    });

    // Sinkronisasi otomatis tanggal selesai kontrak, durasi, PKWT, dan checkbox Tetap
    setTimeout(() => {
      const modalEl = modalObj?.overlay || document.querySelector('.modal-overlay:last-child');
      if (!modalEl) return;

      const isTetapCheck = modalEl.querySelector('#con-modal-is-tetap');
      const endInput = modalEl.querySelector('#con-modal-end');
      const startInput = modalEl.querySelector('#con-modal-start');
      const typeSelect = modalEl.querySelector('#con-modal-type');
      const pkwtSelect = modalEl.querySelector('#con-modal-pkwt');

      if (!isTetapCheck || !endInput) return;

      function updateTetapState(isTetap, syncDropdowns = true) {
        if (isTetap) {
          isTetapCheck.checked = true;
          endInput.disabled = true;
          endInput.style.opacity = '0.5';
          endInput.style.backgroundColor = '#F8FAFC';
          endInput.style.cursor = 'not-allowed';
          endInput.value = '';
          if (syncDropdowns) {
            if (typeSelect) typeSelect.value = 'PKWTT (TETAP)';
            if (pkwtSelect) pkwtSelect.value = 'TETAP';
          }
        } else {
          isTetapCheck.checked = false;
          endInput.disabled = false;
          endInput.style.opacity = '1';
          endInput.style.backgroundColor = '#FFFFFF';
          endInput.style.cursor = 'default';
          if (syncDropdowns) {
            if (typeSelect && typeSelect.value === 'PKWTT (TETAP)') typeSelect.value = 'KONTRAK 1 TAHUN';
            if (pkwtSelect && pkwtSelect.value === 'TETAP') pkwtSelect.value = 'PKWT 1';
          }
        }
      }

      isTetapCheck.addEventListener('change', (e) => {
        updateTetapState(e.target.checked, true);
        if (!e.target.checked) endInput.focus();
      });

      if (typeSelect) {
        typeSelect.addEventListener('change', (e) => {
          const val = e.target.value;
          if (val === 'PKWTT (TETAP)') {
            updateTetapState(true, false);
            if (pkwtSelect) pkwtSelect.value = 'TETAP';
          } else {
            updateTetapState(false, false);
            if (pkwtSelect && pkwtSelect.value === 'TETAP') pkwtSelect.value = 'PKWT 1';
            if (startInput && startInput.value) {
              endInput.value = computeContractEndDate(startInput.value, val);
            }
          }
        });
      }

      if (pkwtSelect) {
        pkwtSelect.addEventListener('change', (e) => {
          const val = e.target.value;
          if (val === 'TETAP') {
            if (typeSelect) typeSelect.value = 'PKWTT (TETAP)';
            updateTetapState(true, false);
          } else {
            if (typeSelect && typeSelect.value === 'PKWTT (TETAP)') typeSelect.value = 'KONTRAK 1 TAHUN';
            updateTetapState(false, false);
            if (startInput && startInput.value && typeSelect) {
              endInput.value = computeContractEndDate(startInput.value, typeSelect.value);
            }
          }
        });
      }

      if (startInput) {
        startInput.addEventListener('change', (e) => {
          if (!isTetapCheck.checked && e.target.value && typeSelect) {
            endInput.value = computeContractEndDate(e.target.value, typeSelect.value);
          }
        });
      }
    }, 20);
  }

  // Delete Contract
  function confirmDelete(item) {
    createModal({
      title: '🗑️ Hapus Data Kontrak',
      content: `
        <div style="font-size:0.92rem; color:#334155; line-height:1.5;">
          Apakah Anda yakin ingin menghapus catatan kontrak untuk <strong>${item.employee_name || 'karyawan ini'}</strong> (ID #${item.id})?
          <div style="margin-top:10px; padding:10px; background:#FEF2F2; border-radius:8px; border:1px solid #FECACA; font-size:0.8rem; color:#991B1B;">
            ⚠️ Tindakan ini akan menghapus riwayat masa berlaku kontrak terkait dari sistem.
          </div>
        </div>
      `,
      confirmText: 'Ya, Hapus',
      cancelText: 'Batal',
      onConfirm: async (_, close) => {
        try {
          const res = await apiFetch(`/api/contracts/${item.id}`, { method: 'DELETE' });
          if (res.ok) {
            toastSuccess('Data kontrak berhasil dihapus.');
            close();
            await loadData();
          } else {
            toastError('Gagal menghapus kontrak: ' + (res.data?.message || 'Terjadi kesalahan'));
          }
        } catch (err) {
          console.error(err);
          toastError('Terjadi kendala koneksi.');
        }
      }
    });
  }

  // Ekspor Excel (seluruh hasil filter aktif lintas halaman via /api/contracts/export)
  async function exportToExcel() {
    if (!hasPermission(currentUser, 'export')) {
      toastError('Anda tidak memiliki izin untuk mengekspor data.');
      return;
    }
    try {
      const rows = await fetchExportRows('/api/contracts/export', getActiveFilters());
      if (!rows.length) {
        toastError('Belum ada data kontrak untuk diekspor.');
        return;
      }
      const exportData = rows.map(c => {
        const tetap = isTetap(c.end_date);
        const hasEnd = Boolean(c.end_date && String(c.end_date).trim() && c.end_date !== '-');
        const rem = c.days_remaining !== null && c.days_remaining !== undefined && c.days_remaining !== '' ? Number(c.days_remaining) : null;
        return {
          'Nama Lengkap': c.employee_name || '',
          'Cabang': c.branch_name || '',
          'Divisi': c.division || '',
          'No. PKWT': c.pkwt_number || '',
          'Tipe Kontrak': c.contract_type || '',
          'Tanggal Mulai': c.start_date || '',
          'Tanggal Selesai': tetap ? 'Tetap / PKWTT' : (hasEnd ? c.end_date : '-'),
          'Sisa Hari': tetap ? 'Tetap' : (hasEnd && rem !== null && !isNaN(rem) ? rem + ' Hari' : '-'),
          'Status': c.status || '',
          'Catatan': c.notes || ''
        };
      });
      downloadExcel(exportData, `Data_Kontrak_Karyawan_${new Date().toISOString().slice(0, 10)}`);
      toastSuccess('File Excel Data Kontrak berhasil diunduh!');
    } catch (err) {
      toastError('Gagal mengekspor data kontrak: ' + err.message);
    }
  }


  // Render only the filtered list content HTML (Card / Table / Empty State)
  function buildListHTML(filtered) {
    if (filtered.length === 0) {
      return `
            <div class="con-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">🔍</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak ada kontrak yang sesuai kriteria</h3>
              <p style="font-size:0.85rem; color:#64748B; margin:0 0 16px 0;">Coba atur ulang kata kunci pencarian, filter cabang, atau filter masa berlaku kontrak.</p>
              <button id="con-reset-filter" class="btn btn-secondary" style="font-size:0.82rem; padding:7px 16px; border-radius:8px;">Reset Semua Filter</button>
            </div>
      `;
    }

    if (viewMode === 'table') {
      return `
            <div class="con-table-card">
              <div class="con-table-responsive">
                <table class="con-table">
                  <thead>
                    <tr>
                      <th style="width:45px; text-align:center;">No</th>
                      <th>Nama Karyawan</th>
                      <th>Cabang & Divisi</th>
                      <th style="text-align:center;">Tingkat PKWT</th>
                      <th>Mulai Kontrak</th>
                      <th>Selesai Kontrak</th>
                      <th>Masa Berlaku</th>
                      <th style="text-align:center;">Status</th>
                      <th style="width:110px; text-align:center;">Aksi</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${filtered.map((item, idx) => {
                      const div = item.division || 'FACILITY CARE';
                      let divClass = 'con-div-fc';
                      let avatarClass = '';
                      if (div === 'SECURITY') {
                        divClass = 'con-div-sec';
                        avatarClass = 'sec';
                      }

                      const hasEndDate = Boolean(item.end_date && String(item.end_date).trim() && item.end_date !== '-');
                      const tetap = isTetap(item.end_date);
                      const isAktif = item.status === 'Aktif';
                      const rawRem = item.days_remaining !== null && item.days_remaining !== undefined && item.days_remaining !== '' ? Number(item.days_remaining) : null;
                      const rem = (hasEndDate && !tetap && rawRem !== null && !isNaN(rawRem)) ? rawRem : null;

                      let statusBadgeClass = 'con-status-aktif';
                      let statusBadgeText = '🟢 Aktif';

                      if (tetap) {
                        statusBadgeClass = 'con-status-tetap';
                        statusBadgeText = '🏛️ PKWTT (Tetap)';
                      } else if (isAktif) {
                        if (rem !== null) {
                          if (rem <= 0) {
                            statusBadgeClass = 'con-status-danger';
                            statusBadgeText = '⚠️ Habis Hari Ini';
                          } else if (rem <= 30) {
                            statusBadgeClass = 'con-status-warning';
                            statusBadgeText = `⚠️ Sisa ${rem} Hari`;
                          } else {
                            statusBadgeClass = 'con-status-aktif';
                            statusBadgeText = `🟢 Aktif (${rem} Hari)`;
                          }
                        } else {
                          statusBadgeClass = 'con-status-aktif';
                          statusBadgeText = '🟢 Aktif';
                        }
                      } else {
                        statusBadgeClass = 'con-status-danger';
                        statusBadgeText = `⚪ ${item.status || 'Nonaktif'}`;
                      }

                      const phone = item.employee_phone || employeePhoneMap[item.employee_id] || employeePhoneMap[String(item.employee_name || '').toLowerCase()] || '';

                      return `
                        <tr data-id="${item.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600;">${((currentPage - 1) * PAGE_LIMIT) + idx + 1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="con-avatar ${avatarClass}" style="width:34px; height:34px; font-size:0.75rem;">
                                ${getInitials(item.employee_name)}
                              </div>
                              <div style="min-width:0;">
                                <div style="font-weight:700; color:#1E293B; font-size:0.9rem;">${item.employee_name || 'Karyawan Tanpa Nama'}</div>
                                <div style="font-size:0.74rem; color:#64748B;">ID #${item.id} ${item.pkwt_number ? '• ' + item.pkwt_number : ''}</div>
                              </div>
                            </div>
                          </td>
                          <td>
                            <div style="display:flex; flex-direction:column; gap:3px;">
                              <span style="font-weight:600; color:#334155;">${item.branch_name || 'Semua Cabang'}</span>
                              <span class="con-div-badge ${divClass}" style="align-self:flex-start;">
                                ${div === 'SECURITY' ? '🛡️' : '🧹'} ${div}
                              </span>
                            </div>
                          </td>
                          <td style="text-align:center;">
                            <span class="con-pkwt-badge" style="display:inline-block;">${item.contract_type || 'PKWT'}</span>
                          </td>
                          <td style="font-size:0.84rem; color:#475569;">
                            ${formatDateIndo(item.start_date)}
                          </td>
                          <td style="font-size:0.84rem; color:#475569;">
                            ${tetap ? '<span style="color:#6366F1; font-weight:600;">PKWTT (Tetap)</span>' : formatDateIndo(item.end_date)}
                          </td>
                          <td>
                            ${tetap ? '<span style="color:#6366F1; font-weight:600; font-size:0.82rem;">🏛️ Tidak Berbatas</span>' : (rem !== null ? (rem <= 0 ? '<span style="color:#DC2626; font-weight:700; font-size:0.82rem;">⚠️ Habis Hari Ini</span>' : rem <= 30 ? `<span style="color:#D97706; font-weight:700; font-size:0.82rem;">⚠️ Sisa ${rem} Hari</span>` : `<span style="color:#16A34A; font-weight:600; font-size:0.82rem;">🟢 ${rem} Hari</span>`) : '<span style="color:#94A3B8; font-size:0.82rem;">-</span>')}
                          </td>
                          <td style="text-align:center;">
                            <span class="con-status-badge ${statusBadgeClass}">${statusBadgeText}</span>
                          </td>
                          <td style="text-align:center;">
                            <div class="con-action-btns" style="justify-content:center;">
                              ${phone ? `
                                <a href="https://wa.me/${phone.replace(/[^0-9]/g, '').replace(/^0/, '62')}" target="_blank" rel="noopener" class="con-icon-btn" title="Chat WhatsApp Karyawan" style="color:#16A34A;">
                                  <svg width="15" height="15" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                                </a>
                              ` : ''}
                              ${canEdit ? `
                                <button class="con-icon-btn edit" data-action="edit" data-id="${item.id}" title="Edit Data Kontrak">
                                  <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                </button>
                              ` : ''}
                              ${canDelete ? `
                                <button class="con-icon-btn delete" data-action="delete" data-id="${item.id}" title="Hapus Kontrak">
                                  <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                                </button>
                              ` : ''}
                            </div>
                          </td>
                        </tr>
                      `;
                    }).join('')}
                  </tbody>
                </table>
              </div>
            </div>
      `;
    }

    return `
          <div class="con-cards-grid">
            ${filtered.map(item => {
              const div = item.division || 'FACILITY CARE';
              let divClass = 'con-div-fc';
              let avatarClass = '';
              if (div === 'SECURITY') {
                divClass = 'con-div-sec';
                avatarClass = 'sec';
              }

              const hasEndDate = Boolean(item.end_date && String(item.end_date).trim() && item.end_date !== '-');
              const tetap = isTetap(item.end_date);
              const isAktif = item.status === 'Aktif';
              const rawRem = item.days_remaining !== null && item.days_remaining !== undefined && item.days_remaining !== '' ? Number(item.days_remaining) : null;
              const rem = (hasEndDate && !tetap && rawRem !== null && !isNaN(rawRem)) ? rawRem : null;

              let statusBadgeClass = 'con-status-aktif';
              let statusBadgeText = '🟢 Aktif';

              if (tetap) {
                statusBadgeClass = 'con-status-tetap';
                statusBadgeText = '🏛️ PKWTT (Tetap)';
              } else if (isAktif) {
                if (rem !== null) {
                  if (rem <= 0) {
                    statusBadgeClass = 'con-status-danger';
                    statusBadgeText = '⚠️ Habis Hari Ini';
                  } else if (rem <= 30) {
                    statusBadgeClass = 'con-status-warning';
                    statusBadgeText = `⚠️ Sisa ${rem} Hari`;
                  } else {
                    statusBadgeClass = 'con-status-aktif';
                    statusBadgeText = `🟢 Aktif (${rem} Hari)`;
                  }
                } else {
                  statusBadgeClass = 'con-status-aktif';
                  statusBadgeText = '🟢 Aktif';
                }
              } else {
                statusBadgeClass = 'con-status-danger';
                statusBadgeText = `⚪ ${item.status || 'Nonaktif'}`;
              }

              const phone = item.employee_phone || employeePhoneMap[item.employee_id] || employeePhoneMap[String(item.employee_name || '').toLowerCase()] || '';

              return `
                <div class="con-card" data-id="${item.id}">
                  <div class="con-card-top">
                    <div style="display:flex; align-items:center; gap:6px; flex-wrap:wrap;">
                      <span class="con-div-badge ${divClass}">
                        ${div === 'SECURITY' ? '🛡️' : '🧹'} ${div}
                      </span>
                      ${item.pkwt_number ? `
                        <span class="con-pkwt-badge">
                          📜 ${item.pkwt_number}
                        </span>
                      ` : ''}
                    </div>
                    <span class="con-status-badge ${statusBadgeClass}">
                      ${statusBadgeText}
                    </span>
                  </div>

                  <div class="con-profile-box">
                    <div class="con-avatar ${avatarClass}">
                      ${getInitials(item.employee_name)}
                    </div>
                    <div style="min-width:0; flex:1;">
                      <div class="con-name">
                        ${item.employee_name || 'Karyawan Tanpa Nama'}
                      </div>
                      <div class="con-subtitle">
                        Personil ${div} • ${item.contract_type || 'Kontrak Kerja'}
                      </div>
                    </div>
                  </div>

                  <div class="con-info-grid">
                    <div class="con-info-card">
                      <div class="con-info-label">Cabang Penempatan</div>
                      <div class="con-info-val" title="${item.branch_name || '-'}">
                        🏢 ${item.branch_name || 'Semua Cabang'}
                      </div>
                    </div>
                    <div class="con-info-card">
                      <div class="con-info-label">Tingkat PKWT</div>
                      <div class="con-info-val">
                        📋 ${item.contract_type || 'PKWT 1'}
                      </div>
                    </div>
                    <div class="con-info-card">
                      <div class="con-info-label">Mulai Kontrak</div>
                      <div class="con-info-val">
                        📅 ${formatDateIndo(item.start_date)}
                      </div>
                    </div>
                    <div class="con-info-card">
                      <div class="con-info-label">Selesai Kontrak</div>
                      <div class="con-info-val">
                        ${tetap ? '🏛️ PKWTT (Tetap)' : ('📅 ' + formatDateIndo(item.end_date))}
                      </div>
                    </div>
                  </div>

                  ${tetap ? `
                    <div style="background:#EEF2FF; border:1px solid #E0E7FF; border-radius:10px; padding:10px 12px; font-size:0.82rem; color:#4338CA; display:flex; align-items:center; gap:8px;">
                      <span style="font-size:1.1rem;">🏛️</span>
                      <div>
                        <strong>Karyawan Tetap (PKWTT)</strong>
                        <div style="font-size:0.75rem; color:#6366F1;">Tidak memiliki batas akhir masa berlaku kontrak.</div>
                      </div>
                    </div>
                  ` : rem !== null ? `
                    <div style="background:${rem <= 0 ? '#FEF2F2' : (rem <= 30 ? '#FFFBEB' : '#F0FDF4')}; border:1px solid ${rem <= 0 ? '#FEE2E2' : (rem <= 30 ? '#FEF3C7' : '#DCFCE7')}; border-radius:10px; padding:10px 12px; font-size:0.82rem; color:${rem <= 0 ? '#991B1B' : (rem <= 30 ? '#92400E' : '#166534')}; display:flex; align-items:center; gap:8px;">
                      <span style="font-size:1.1rem;">${rem <= 0 ? '⚠️' : (rem <= 30 ? '⏳' : '✅')}</span>
                      <div>
                        <strong>${rem <= 0 ? 'Kontrak Telah Berakhir' : `Sisa Waktu Kontrak: ${rem} Hari`}</strong>
                        <div style="font-size:0.75rem; opacity:0.85;">
                          ${rem <= 30 && rem > 0 ? 'Segera lakukan evaluasi perpanjangan kontrak.' : (rem <= 0 ? 'Masa berlaku kontrak telah habis.' : 'Masa kontrak masih aktif dan aman.')}
                        </div>
                      </div>
                    </div>
                  ` : ''}

                  <div class="con-card-footer">
                    <div style="font-size:0.75rem; color:#94A3B8;">
                      ID #${item.id} • ${item.pkwt_number || 'PKWT'}
                    </div>

                    <div class="con-action-btns">
                      ${phone ? `
                        <a href="https://wa.me/${phone.replace(/[^0-9]/g, '').replace(/^0/, '62')}" target="_blank" rel="noopener" class="con-icon-btn" title="Chat WhatsApp Karyawan" style="color:#16A34A;">
                          <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                        </a>
                      ` : ''}

                      ${canEdit ? `
                        <button class="con-icon-btn edit" data-action="edit" data-id="${item.id}" title="Edit Data Kontrak">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                      ` : ''}
                      ${canDelete ? `
                        <button class="con-icon-btn delete" data-action="delete" data-id="${item.id}" title="Hapus Kontrak">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                        </button>
                      ` : ''}
                    </div>
                  </div>
                </div>
              `;
            }).join('')}
          </div>
    `;
  }

  function renderPagination() {
    const pagEl = container.querySelector('#con-pagination-container');
    if (!pagEl) return;
    pagEl.innerHTML = '';
    const nav = createPagination({
      page: currentPage,
      pages: totalPages,
      total: totalItems,
      limit: PAGE_LIMIT,
      onPage: (targetPage) => {
        currentPage = targetPage;
        loadData({ refreshStats: false });
      }
    });
    if (nav) pagEl.appendChild(nav);
  }

  function bindListEvents() {
    const resetBtn = container.querySelector('#con-reset-filter');
    if (resetBtn) {
      resetBtn.addEventListener('click', () => {
        searchQuery = '';
        selectedBranch = 'all';
        selectedDivision = 'all';
        selectedStatus = 'all';
        selectedExpiring = 'all';
        currentPage = 1;
        const searchInput = container.querySelector('#con-search-input');
        if (searchInput) searchInput.value = '';
        const branchFilter = container.querySelector('#con-branch-filter');
        if (branchFilter) branchFilter.value = 'all';
        const divisionFilter = container.querySelector('#con-division-filter');
        if (divisionFilter) divisionFilter.value = 'all';
        const expiringFilter = container.querySelector('#con-expiring-filter');
        if (expiringFilter) expiringFilter.value = 'all';
        container.querySelectorAll('.con-pill').forEach(p => {
          p.classList.toggle('active', p.getAttribute('data-status') === 'all');
        });
        loadData({ refreshStats: false });
      });
    }

    container.querySelectorAll('.con-icon-btn.edit').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const id = parseInt(e.currentTarget.getAttribute('data-id'), 10);
        const item = contractsList.find(c => c.id === id);
        if (item) openFormModal(item);
      });
    });

    container.querySelectorAll('.con-icon-btn.delete').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const id = parseInt(e.currentTarget.getAttribute('data-id'), 10);
        const item = contractsList.find(c => c.id === id);
        if (item) confirmDelete(item);
      });
    });
  }

  function renderListContent() {
    const listContainer = container.querySelector('#con-list-container');
    if (!listContainer) return;
    const filtered = getFilteredList();
    listContainer.innerHTML = buildListHTML(filtered);
    renderPagination();
    bindListEvents();
  }
  // Main UI Renderer
  function renderUI() {
    const filtered = getFilteredList();

    // KPI Metrics from server stats
    const totalCount = serverStats.total;
    const activeCount = serverStats.active;
    const expiringCount = serverStats.expiring30;
    const permanentCount = serverStats.permanent;
    const nonActiveCount = serverStats.nonactive;

    container.innerHTML = `
      <style>
        .con-page {
          display: flex;
          flex-direction: column;
          gap: 20px;
          width: 100%;
          max-width: 100%;
          min-width: 0;
          box-sizing: border-box;
          overflow-x: hidden;
        }

        /* Hero Header Card */
        .con-hero-card {
          background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
          border-radius: 16px;
          padding: 24px 28px;
          color: white;
          box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25);
          position: relative;
          overflow: hidden;
          box-sizing: border-box;
          width: 100%;
          max-width: 100%;
          min-width: 0;
        }
        .con-hero-card::after {
          content: '';
          position: absolute;
          top: -40px;
          right: -40px;
          width: 220px;
          height: 220px;
          background: radial-gradient(circle, rgba(59, 130, 246, 0.2) 0%, transparent 70%);
          border-radius: 50%;
          pointer-events: none;
        }
        .con-hero-actions {
          display: flex;
          align-items: center;
          gap: 8px;
          flex-wrap: wrap;
        }
        .con-hero-actions .btn {
          height: 38px;
          box-sizing: border-box;
          transition: all 0.2s ease;
        }
        .con-btn-add:hover {
          background: #1D4ED8 !important;
          transform: translateY(-1px);
        }
        .con-btn-excel:hover, .con-btn-refresh:hover {
          background: rgba(255, 255, 255, 0.16) !important;
          border-color: rgba(255, 255, 255, 0.3) !important;
          transform: translateY(-1px);
        }
        .con-btn-missing:hover {
          background: rgba(239, 68, 68, 0.25) !important;
          border-color: rgba(239, 68, 68, 0.6) !important;
          transform: translateY(-1px);
        }
        .con-stats-row {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 14px;
          margin-top: 20px;
        }
        .con-stat-box {
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
        .con-stat-box:hover {
          background: rgba(255, 255, 255, 0.08);
          transform: translateY(-2px);
        }
        .con-stat-icon {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 1.3rem;
          flex-shrink: 0;
        }
        .con-stat-val {
          font-size: 1.55rem;
          font-weight: 800;
          line-height: 1.1;
          letter-spacing: -0.02em;
        }
        .con-stat-lbl {
          font-size: 0.72rem;
          text-transform: uppercase;
          color: #94A3B8;
          font-weight: 700;
          letter-spacing: 0.05em;
          margin-top: 2px;
        }

        /* Controls / Filter Bar */
        .con-controls-card {
          background: #172033;
          border-radius: 14px;
          border: 1px solid #29364B;
          padding: 16px 20px;
          box-shadow: 0 2px 8px rgba(0,0,0,0.03);
          display: flex;
          flex-direction: column;
          gap: 14px;
          box-sizing: border-box;
          width: 100%;
          max-width: 100%;
          min-width: 0;
        }
        .con-search-row {
          display: flex;
          gap: 12px;
          align-items: center;
          flex-wrap: wrap;
        }
        .con-search-wrap {
          position: relative;
          flex: 1;
          min-width: 240px;
        }
        .con-search-icon {
          position: absolute;
          left: 14px;
          top: 50%;
          transform: translateY(-50%);
          color: #94A3B8;
          pointer-events: none;
        }
        .con-search-input {
          width: 100%;
          padding: 10px 14px 10px 40px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          outline: none;
          transition: border-color 0.15s ease;
        }
        .con-search-input:focus {
          border-color: #2563EB;
          box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);
        }
        .con-select {
          padding: 10px 14px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          color: #F8FAFC;
          background: #172033;
          outline: none;
          min-width: 170px;
          cursor: pointer;
        }
        .con-select:focus {
          border-color: #2563EB;
        }

        /* Swipeable Pills */
        .con-filter-pills {
          display: flex;
          align-items: center;
          gap: 8px;
          overflow-x: auto;
          white-space: nowrap;
          scrollbar-width: none;
          padding-bottom: 2px;
        }
        .con-filter-pills::-webkit-scrollbar {
          display: none;
        }
        .con-pill {
          padding: 6px 14px;
          border-radius: 20px;
          font-size: 0.82rem;
          font-weight: 600;
          cursor: pointer;
          border: 1px solid #29364B;
          background: #1B2940;
          color: #64748B;
          transition: all 0.15s ease;
          display: inline-flex;
          align-items: center;
          gap: 6px;
          user-select: none;
          flex-shrink: 0;
        }
        .con-pill:hover {
          background: #111827;
          color: #F8FAFC;
        }
        .con-pill.active {
          background: #0F172A;
          color: white;
          border-color: #0F172A;
        }
        .con-pill-dot {
          width: 8px;
          height: 8px;
          border-radius: 50%;
        }
        .con-pill-count {
          padding: 1px 6px;
          border-radius: 10px;
          font-size: 0.72rem;
          background: rgba(0,0,0,0.06);
        }
        .con-pill.active .con-pill-count {
          background: rgba(255,255,255,0.2);
          color: white;
        }

        /* View Mode Toggle */
        .con-view-toggle {
          display: inline-flex;
          align-items: center;
          background: #111827;
          padding: 3px;
          border-radius: 10px;
          border: 1px solid #CBD5E1;
          gap: 3px;
          flex-shrink: 0;
        }
        .con-view-btn {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          padding: 6px 12px;
          border-radius: 7px;
          font-size: 0.82rem;
          font-weight: 600;
          border: none;
          background: transparent;
          color: #64748B;
          cursor: pointer;
          transition: all 0.15s ease;
          height: 34px;
          box-sizing: border-box;
          user-select: none;
        }
        .con-view-btn:hover {
          color: #0F172A;
          background: rgba(255, 255, 255, 0.5);
        }
        .con-view-btn.active {
          background: #172033;
          color: #2563EB;
          font-weight: 700;
          box-shadow: 0 1px 4px rgba(0, 0, 0, 0.08);
        }

        /* Modern Table View */
        .con-table-card {
          background: #172033;
          border: 1px solid #29364B;
          border-radius: 14px;
          box-shadow: 0 2px 6px rgba(0, 0, 0, 0.03);
          overflow: hidden;
        }
        .con-table-responsive {
          width: 100%;
          overflow-x: auto;
          -webkit-overflow-scrolling: touch;
        }
        .con-table {
          width: 100%;
          border-collapse: collapse;
          text-align: left;
          font-size: 0.88rem;
          min-width: 960px;
        }
        .con-table th {
          background: #1B2940;
          color: #475569;
          font-size: 0.76rem;
          font-weight: 700;
          text-transform: uppercase;
          letter-spacing: 0.05em;
          padding: 13px 16px;
          border-bottom: 1.5px solid #E2E8F0;
          white-space: nowrap;
        }
        .con-table td {
          padding: 13px 16px;
          border-bottom: 1px solid #F1F5F9;
          color: #F8FAFC;
          vertical-align: middle;
        }
        .con-table tbody tr:hover td {
          background: #1B2940;
        }
        .con-table tbody tr:last-child td {
          border-bottom: none;
        }

        /* Cards Grid */
        .con-cards-grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
          gap: 16px;
          width: 100%;
          max-width: 100%;
          min-width: 0;
          box-sizing: border-box;
        }
        .con-card {
          background: #172033;
          border-radius: 14px;
          border: 1px solid #29364B;
          padding: 16px 18px;
          box-shadow: 0 2px 8px rgba(0,0,0,0.04);
          display: flex;
          flex-direction: column;
          gap: 13px;
          width: 100% !important;
          max-width: 100% !important;
          min-width: 0 !important;
          box-sizing: border-box !important;
          overflow: hidden;
          transition: transform 0.15s ease, box-shadow 0.15s ease;
        }
        .con-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 6px 16px rgba(0,0,0,0.08);
          border-color: #CBD5E1;
        }

        /* Card Top Badges */
        .con-card-top {
          display: flex;
          justify-content: space-between;
          align-items: center;
          gap: 8px;
          width: 100%;
          min-width: 0;
          box-sizing: border-box;
          flex-wrap: wrap;
        }
        .con-div-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.74rem;
          font-weight: 700;
          flex-shrink: 0;
          white-space: nowrap;
        }
        .con-div-fc {
          background: #EFF6FF;
          color: #2563EB;
          border: 1px solid #DBEAFE;
        }
        .con-div-sec {
          background: #111827;
          color: #334155;
          border: 1px solid #29364B;
        }
        .con-div-rel {
          background: #F5F3FF;
          color: #7C3AED;
          border: 1px solid #EDE9FE;
        }

        .con-pkwt-badge {
          display: inline-flex;
          align-items: center;
          gap: 4px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.74rem;
          font-weight: 700;
          background: #F5F3FF;
          color: #6D28D9;
          border: 1px solid #DDD6FE;
          flex-shrink: 0;
          white-space: nowrap;
        }

        .con-status-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.74rem;
          font-weight: 700;
          flex-shrink: 0;
          white-space: nowrap;
        }
        .con-status-aktif {
          background: #DCFCE7;
          color: #16A34A;
          border: 1px solid #BBF7D0;
        }
        .con-status-warning {
          background: #FEF3C7;
          color: #B45309;
          border: 1px solid #FDE68A;
        }
        .con-status-danger {
          background: #FEE2E2;
          color: #DC2626;
          border: 1px solid #FCA5A5;
        }
        .con-status-tetap {
          background: #E0E7FF;
          color: #4338CA;
          border: 1px solid #C7D2FE;
        }

        /* Profile Box */
        .con-profile-box {
          display: flex;
          align-items: center;
          gap: 12px;
          width: 100%;
          min-width: 0;
          box-sizing: border-box;
        }
        .con-avatar {
          width: 48px;
          height: 48px;
          border-radius: 12px;
          background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
          color: white;
          display: flex;
          align-items: center;
          justify-content: center;
          font-weight: 800;
          font-size: 1.05rem;
          flex-shrink: 0;
          box-shadow: 0 4px 10px rgba(37, 99, 235, 0.2);
        }
        .con-avatar.sec {
          background: linear-gradient(135deg, #334155 0%, #0F172A 100%);
          box-shadow: 0 4px 10px rgba(15, 23, 42, 0.2);
        }
        .con-avatar.rel {
          background: linear-gradient(135deg, #7C3AED 0%, #5B21B6 100%);
          box-shadow: 0 4px 10px rgba(124, 58, 237, 0.2);
        }
        .con-name {
          font-size: 1.02rem;
          font-weight: 800;
          color: #0F172A;
          line-height: 1.25;
          word-break: break-word;
          overflow-wrap: anywhere;
        }
        .con-subtitle {
          font-size: 0.76rem;
          color: #64748B;
          font-weight: 600;
          margin-top: 3px;
        }

        /* Info Container */
        .con-info-card {
          background: #1B2940;
          border: 1px solid #F1F5F9;
          border-radius: 10px;
          padding: 11px 12px;
          display: flex;
          flex-direction: column;
          gap: 8px;
          font-size: 0.83rem;
          width: 100%;
          max-width: 100%;
          min-width: 0;
          box-sizing: border-box;
        }
        .con-info-row {
          display: flex;
          align-items: flex-start;
          justify-content: space-between;
          gap: 8px;
          width: 100%;
          min-width: 0;
          box-sizing: border-box;
        }
        .con-info-label {
          color: #64748B;
          display: inline-flex;
          align-items: center;
          gap: 5px;
          flex-shrink: 0;
          white-space: nowrap;
          font-size: 0.8rem;
        }
        .con-info-val {
          color: #0F172A;
          font-weight: 600;
          text-align: right;
          min-width: 0;
          flex: 1;
          word-break: break-word;
          overflow-wrap: anywhere;
          font-size: 0.82rem;
        }

        /* Notes box */
        .con-notes-box {
          font-size: 0.78rem;
          color: #475569;
          background: #FFFBEB;
          border: 1px solid #FEF3C7;
          border-radius: 8px;
          padding: 8px 10px;
          line-height: 1.4;
          word-break: break-word;
          overflow-wrap: anywhere;
          width: 100%;
          box-sizing: border-box;
        }

        /* Card footer */
        .con-card-footer {
          display: flex;
          justify-content: space-between;
          align-items: center;
          border-top: 1px solid #F1F5F9;
          padding-top: 11px;
          margin-top: auto;
          gap: 8px;
          width: 100%;
          min-width: 0;
          box-sizing: border-box;
          flex-wrap: wrap;
        }
        .con-action-btns {
          display: flex;
          align-items: center;
          gap: 6px;
          margin-left: auto;
          flex-shrink: 0;
        }
        .con-icon-btn {
          width: 34px;
          height: 34px;
          border-radius: 8px;
          border: 1px solid #29364B;
          background: #172033;
          color: #64748B;
          display: flex;
          align-items: center;
          justify-content: center;
          cursor: pointer;
          transition: all 0.15s ease;
        }
        .con-icon-btn:hover {
          background: #111827;
          color: #0F172A;
          border-color: #CBD5E1;
        }
        .con-icon-btn.delete:hover {
          background: #FEE2E2;
          color: #DC2626;
          border-color: #FCA5A5;
        }

        /* Empty state */
        .con-empty-state {
          grid-column: 1 / -1;
          background: #172033;
          border: 2px dashed #29364B;
          border-radius: 16px;
          padding: 48px 24px;
          text-align: center;
        }

        /* Mobile Responsive */
        @media (max-width: 768px) {
          .con-page {
            gap: 14px;
            width: 100% !important;
            max-width: 100% !important;
            overflow-x: hidden !important;
          }
          .con-hero-card {
            padding: 16px 14px !important;
            width: 100% !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
          }
          .con-stats-row {
            grid-template-columns: 1fr 1fr;
            gap: 10px;
          }
          .con-stat-box {
            padding: 10px 12px;
          }
          .con-stat-icon {
            width: 38px;
            height: 38px;
            font-size: 1.1rem;
          }
          .con-stat-val {
            font-size: 1.3rem;
          }
          .con-hero-actions {
            width: 100% !important;
            display: flex !important;
            align-items: center !important;
            gap: 8px !important;
            flex-wrap: wrap !important;
            overflow-x: visible !important;
          }
          .con-hero-actions::-webkit-scrollbar {
            display: none;
          }
          .con-btn-add {
            width: 100% !important;
            justify-content: center !important;
            padding: 9px 12px !important;
            font-size: 0.85rem !important;
            white-space: nowrap !important;
            height: 40px !important;
          }
          .con-btn-excel {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 12px !important;
            font-size: 0.85rem !important;
            white-space: nowrap !important;
            height: 38px !important;
          }
          .con-btn-refresh {
            width: 38px !important;
            height: 38px !important;
            padding: 0 !important;
            flex-shrink: 0 !important;
            justify-content: center !important;
          }
          .con-controls-card {
            padding: 14px !important;
            gap: 12px;
            width: 100% !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
          }
          .con-search-row {
            flex-direction: column;
            gap: 10px;
          }
          .con-select {
            width: 100%;
          }
          .con-view-toggle {
            width: 100%;
            display: flex;
          }
          .con-view-btn {
            flex: 1;
            justify-content: center;
          }
          .con-table-card {
            border-radius: 12px;
          }
          .con-cards-grid {
            grid-template-columns: minmax(0, 1fr) !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
          }
          .con-card {
            padding: 14px !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
          }
          .con-info-card {
            padding: 10px !important;
            width: 100% !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
          }
        }
      </style>

      <div class="con-page">

        <!-- Hero Header Card -->
        <div class="con-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px;">
                <span style="font-size:1.6rem;">📋</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Monitoring Kontrak Kerja Karyawan
                </h1>
              </div>
            </div>

            <!-- Action Buttons 1 Baris Sejajar -->
            <div class="con-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${canAdd ? `
                <button id="con-btn-add" class="btn btn-primary con-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Kontrak</span>
                </button>
              ` : ''}

              ${hasPermission(currentUser, "export") ? `
              <button id="con-btn-excel" class="btn btn-ghost con-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>📥</span>
                <span>Ekspor Excel</span>
              </button>
              ` : ''}

              

              <button id="con-btn-refresh" class="btn btn-ghost con-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:inline-flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- 4 KPI Stat Boxes -->
          <div class="con-stats-row">
            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(59, 130, 246, 0.2); color:#60A5FA;">
                📋
              </div>
              <div>
                <div class="con-stat-val" id="con-stat-total">${totalCount}</div>
                <div class="con-stat-lbl">Total Kontrak</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#34D399;">
                🟢
              </div>
              <div>
                <div class="con-stat-val" id="con-stat-active" style="color:#34D399;">${activeCount}</div>
                <div class="con-stat-lbl">Kontrak Aktif</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(245, 158, 11, 0.2); color:#FBBF24;">
                ⚠️
              </div>
              <div>
                <div class="con-stat-val" id="con-stat-expiring" style="color:#FBBF24;">${expiringCount}</div>
                <div class="con-stat-lbl">Habis &lt; 30 Hari</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(99, 102, 241, 0.2); color:#A5B4FC;">
                🏛️
              </div>
              <div>
                <div class="con-stat-val" id="con-stat-permanent" style="color:#A5B4FC;">${permanentCount}</div>
                <div class="con-stat-lbl">PKWTT (Tetap)</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Controls / Filter Card -->
        <div class="con-controls-card">
          <div class="con-search-row">
            <div class="con-search-wrap">
              <svg class="con-search-icon" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              <input type="text" id="con-search-input" class="con-search-input" 
                placeholder="Cari nama karyawan, cabang, no PKWT, catatan..." value="${searchQuery ? String(searchQuery).replace(/"/g, '&quot;') : ''}" />
            </div>

            <select id="con-branch-filter" class="con-select">
              <option value="all">Semua Cabang (${branchOptions.length})</option>
              ${branchOptions.map(b => `
                <option value="${b.value}" ${String(selectedBranch) === String(b.value) ? 'selected' : ''}>
                  ${b.label}
                </option>
              `).join('')}
            </select>

            <select id="con-division-filter" class="con-select">
              <option value="all">Semua Divisi</option>
              <option value="FACILITY CARE" ${selectedDivision === 'FACILITY CARE' ? 'selected' : ''}>FACILITY CARE</option>
              <option value="SECURITY" ${selectedDivision === 'SECURITY' ? 'selected' : ''}>SECURITY</option>
            </select>

            <select id="con-expiring-filter" class="con-select">
              <option value="all" ${selectedExpiring === 'all' ? 'selected' : ''}>Semua Durasi Masa Habis</option>
              <option value="7" ${selectedExpiring === '7' ? 'selected' : ''}>Akan Habis ≤ 7 Hari</option>
              <option value="14" ${selectedExpiring === '14' ? 'selected' : ''}>Akan Habis ≤ 14 Hari</option>
              <option value="30" ${selectedExpiring === '30' ? 'selected' : ''}>Akan Habis ≤ 30 Hari</option>
              <option value="60" ${selectedExpiring === '60' ? 'selected' : ''}>Akan Habis ≤ 60 Hari</option>
            </select>

            <div class="con-view-toggle">
              <button type="button" class="con-view-btn ${viewMode === 'card' ? 'active' : ''}" data-view="card" title="Tampilan Kartu">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="con-view-btn ${viewMode === 'table' ? 'active' : ''}" data-view="table" title="Tampilan Tabel">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9h18M3 15h18M9 3v18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="con-filter-pills">
            <div class="con-pill ${selectedStatus === 'all' && selectedExpiring === 'all' ? 'active' : ''}" data-status="all">
              <span>Semua Kontrak</span>
              <span class="con-pill-count" id="con-pill-count-all">${totalCount}</span>
            </div>
            <div class="con-pill ${selectedStatus === 'Aktif' && selectedExpiring === 'all' ? 'active' : ''}" data-status="Aktif">
              <span class="con-pill-dot" style="background:#10B981;"></span>
              <span>Aktif</span>
              <span class="con-pill-count" id="con-pill-count-aktif">${activeCount}</span>
            </div>
            <div class="con-pill ${selectedStatus === 'Expiring30' ? 'active' : ''}" data-status="Expiring30">
              <span class="con-pill-dot" style="background:#F59E0B;"></span>
              <span>Akan Habis (&lt;30 Hari)</span>
              <span class="con-pill-count" id="con-pill-count-expiring">${expiringCount}</span>
            </div>
            <div class="con-pill ${selectedStatus === 'Tetap' ? 'active' : ''}" data-status="Tetap">
              <span class="con-pill-dot" style="background:#6366F1;"></span>
              <span>PKWTT (Tetap)</span>
              <span class="con-pill-count" id="con-pill-count-tetap">${permanentCount}</span>
            </div>
            <div class="con-pill ${selectedStatus === 'Nonaktif' ? 'active' : ''}" data-status="Nonaktif">
              <span class="con-pill-dot" style="background:#EF4444;"></span>
              <span>Nonaktif / Resign</span>
              <span class="con-pill-count" id="con-pill-count-nonaktif">${nonActiveCount}</span>
            </div>
          </div>
        </div>

        <!-- Content Area: Pilihan Kartu atau Tabel -->
        <div id="con-list-container">
          ${buildListHTML(filtered)}
        </div>
        <div id="con-pagination-container"></div>

      </div>
    `;

    // Event Listeners
    const addBtn = container.querySelector('#con-btn-add');
    if (addBtn) addBtn.addEventListener('click', () => openFormModal());

    const excelBtn = container.querySelector('#con-btn-excel');
    if (excelBtn) excelBtn.addEventListener('click', () => exportToExcel());

    const refreshBtn = container.querySelector('#con-btn-refresh');
    if (refreshBtn) refreshBtn.addEventListener('click', () => loadData({ refreshStats: true }));

    // 0. View Mode Toggle (Kartu / Tabel)
    container.querySelectorAll('.con-view-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const mode = btn.getAttribute('data-view');
        if (mode && mode !== viewMode) {
          viewMode = mode;
          localStorage.setItem('fcms_contracts_view_mode', mode);
          container.querySelectorAll('.con-view-btn').forEach(b => {
            b.classList.toggle('active', b.getAttribute('data-view') === mode);
          });
          renderListContent();
        }
      });
    });

    // 1. Search Input (debounced server search)
    const searchInput = container.querySelector('#con-search-input');
    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        searchQuery = e.target.value;
        clearTimeout(searchDebounce);
        searchDebounce = setTimeout(() => {
          currentPage = 1;
          loadData({ refreshStats: false });
        }, 350);
      });
    }

    // 2. Branch Filter
    const branchFilter = container.querySelector('#con-branch-filter');
    if (branchFilter) {
      branchFilter.addEventListener('change', (e) => {
        selectedBranch = e.target.value;
        currentPage = 1;
        loadData({ refreshStats: false });
      });
    }

    // 3. Division Filter
    const divisionFilter = container.querySelector('#con-division-filter');
    if (divisionFilter) {
      divisionFilter.addEventListener('change', (e) => {
        selectedDivision = e.target.value;
        currentPage = 1;
        loadData({ refreshStats: false });
      });
    }

    // 4. Expiring Days Filter
    const expiringFilter = container.querySelector('#con-expiring-filter');
    if (expiringFilter) {
      expiringFilter.addEventListener('change', (e) => {
        selectedExpiring = e.target.value;
        if (selectedStatus === 'Expiring30' && selectedExpiring !== '30') {
          selectedStatus = selectedExpiring === 'all' ? 'all' : 'Aktif';
        }
        container.querySelectorAll('.con-pill').forEach(p => {
          const st = p.getAttribute('data-status');
          const isActive = (st === 'all' && selectedStatus === 'all' && selectedExpiring === 'all') ||
                           (st === 'Aktif' && selectedStatus === 'Aktif' && selectedExpiring === 'all') ||
                           (st === 'Expiring30' && selectedStatus === 'Expiring30') ||
                           (st === 'Tetap' && selectedStatus === 'Tetap') ||
                           (st === 'Nonaktif' && selectedStatus === 'Nonaktif');
          p.classList.toggle('active', isActive);
        });
        currentPage = 1;
        loadData({ refreshStats: false });
      });
    }

    // 5. Filter Pills
    container.querySelectorAll('.con-pill').forEach(pill => {
      pill.addEventListener('click', (e) => {
        const status = e.currentTarget.getAttribute('data-status');
        selectedStatus = status;
        selectedMonthExpiry = '';
        if (status === 'Expiring30') {
          selectedExpiring = '30';
        } else {
          selectedExpiring = 'all';
        }
        const expEl = container.querySelector('#con-expiring-filter');
        if (expEl) expEl.value = selectedExpiring;
        container.querySelectorAll('.con-pill').forEach(p => {
          p.classList.toggle('active', p === e.currentTarget);
        });
        currentPage = 1;
        loadData({ refreshStats: false });
      });
    });

    // Bind pagination & list action events (Reset button, Edit, Delete)
    renderPagination();
    bindListEvents();
  }

  // Initial load
  await loadData({ refreshStats: true });
}
