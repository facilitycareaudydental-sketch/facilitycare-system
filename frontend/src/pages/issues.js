import { apiFetch, getUser, hasPermission } from '../config.js';
import { getCachedBranches, getCachedEmployees, fetchExportRows } from '../utils/dataCache.js';
import { createModal, confirmDialog } from '../components/modal.js';
import { toastSuccess, toastError, toastInfo } from '../components/toast.js';
import { createPagination } from '../components/table.js';
import { mountEmployeeCombobox } from '../components/employeeCombobox.js';
import { formatDateIndo } from '../utils/sorting.js';

// Fallback kosong bila API offline (mencegah data keluhan/PII tertanam di bundle klien)
const INITIAL_ISSUES = [];

const CATEGORY_LIST = ['SDM', 'Cleaning', 'Aset', 'K3', 'Lainnya'];
const STATUS_LIST = ['Open', 'In Progress', 'Done'];
const SOURCE_LIST = ['SPV', 'AM', 'RCP', 'Perawat', 'FC', 'Berlin', 'Ade', 'Pattrel', 'Dentrel'];

// Standardized Indonesian date helper imported from ../utils/sorting.js

// Inisial avatar
function getInitials(name) {
  if (!name || name === '-') return '⚠️';
  const parts = String(name).trim().split(/\s+/);
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return (parts[0][0] + parts[1][0]).toUpperCase();
}

// Validasi strict row_version untuk Optimistic Concurrency Control (OCC)
function parseValidRowVersion(raw) {
  if (raw === undefined || raw === null || raw === '') return null;
  const num = Number(raw);
  return Number.isInteger(num) && num >= 0 ? num : null;
}

export async function renderIssues(container, params) {
  const PAGE_LIMIT = 20;
  let issuesList = [];
  let branchOptions = [];
  let employeeOptions = [];
  let selectedStatus = 'all';
  let selectedCategory = 'all';
  let selectedBranch = 'all';
  let searchQuery = '';
  let currentPage = 1;
  let totalItems = 0;
  let totalPages = 1;
  let statsLoaded = false;
  let serverStats = { total: 0, open: 0, inProgress: 0, openAndProgress: 0, done: 0, currentMonth: 0 };
  let searchDebounce = null;
  let isLoading = false;

  // Handle dashboard filter parameter
  const dashFilter = params ? params.get('dash_filter') : null;
  if (dashFilter === 'open') {
    selectedStatus = 'all'; // Default to all so list is never empty; user can click In Progress pill to view active issue
  }

  const currentUser = getUser();
  const canCreate = hasPermission(currentUser, 'issues', 'create');
  const canEdit = hasPermission(currentUser, 'issues', 'update');
  const canDelete = hasPermission(currentUser, 'issues', 'delete');

  // Load cache data
  try {
    [branchOptions, employeeOptions] = await Promise.all([
      getCachedBranches(),
      getCachedEmployees()
    ]);
  } catch (err) {
    console.error('Error fetching cache:', err);
  }

  function getActiveFilters() {
    const filters = { sort: 'date_desc' };
    if (searchQuery.trim()) filters.search = searchQuery.trim();
    if (selectedBranch && selectedBranch !== 'all') filters.branch_id = selectedBranch;
    if (selectedCategory && selectedCategory !== 'all') filters.category = selectedCategory;
    if (selectedStatus && selectedStatus !== 'all') filters.status = selectedStatus;
    return filters;
  }

  function updateStatsUI() {
    const set = (id, val) => {
      const el = container.querySelector(id);
      if (el) el.textContent = val;
    };
    set('#issues-stat-total', serverStats.total);
    set('#issues-stat-open-prog', serverStats.openAndProgress);
    set('#issues-stat-done', serverStats.done);
    set('#issues-stat-month', serverStats.currentMonth);
    set('#issues-pill-count-all', serverStats.total);
    set('#issues-pill-count-open', serverStats.open);
    set('#issues-pill-count-inprogress', serverStats.inProgress);
    set('#issues-pill-count-done', serverStats.done);
  }

  // Load Issues Data from API (Server-Side Pagination)
  async function loadData({ refreshStats = true } = {}) {
    isLoading = true;
    const hasShell = Boolean(container.querySelector('#issues-list-container'));
    if (!hasShell) {
      renderUI();
    } else {
      const listEl = container.querySelector('#issues-list-container');
      if (listEl) {
        listEl.innerHTML = '<div class="issues-empty-state"><p style="color:#CBD5E1; margin:0;">Memuat data permasalahan...</p></div>';
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
      const res = await apiFetch(`/api/issues?${qs.toString()}`);
      if (res.ok) {
        const raw = res.data?.data || [];
        issuesList = Array.isArray(raw) ? raw : [];
        const pag = res.data?.pagination || {};
        totalItems = Number(pag.total ?? issuesList.length);
        totalPages = Math.max(1, Number(pag.pages ?? Math.ceil(totalItems / PAGE_LIMIT)) || 1);
        currentPage = Math.max(1, Number(pag.page ?? currentPage));
        if (res.data?.stats) {
          serverStats = res.data.stats;
          statsLoaded = true;
        }
        if (issuesList.length === 0 && currentPage > totalPages) {
          currentPage = totalPages;
          isLoading = false;
          return loadData({ refreshStats });
        }
      } else {
        issuesList = [];
        totalItems = 0;
        totalPages = 1;
        toastError(res.data?.error || 'Gagal memuat data permasalahan');
      }
    } catch (e) {
      console.warn('Gagal load issues:', e);
      issuesList = [];
      totalItems = 0;
      totalPages = 1;
      toastError('Gagal memuat data permasalahan: ' + e.message);
    } finally {
      isLoading = false;
      if (container.querySelector('#issues-list-container')) {
        updateStatsUI();
        renderListContent();
      } else {
        renderUI();
      }
    }
  }

  // Export to Excel (All matching filtered rows via /api/issues/export)
  async function exportToExcel() {
    if (!hasPermission(currentUser, 'export')) {
      toastError('Anda tidak memiliki izin untuk mengekspor data.');
      return;
    }
    try {
      const [{ downloadExcel }, rows] = await Promise.all([
        import('../utils/excel.js'),
        fetchExportRows('/api/issues/export', getActiveFilters())
      ]);
      const exportData = rows.map((item, idx) => ({
        'No': idx + 1,
        'Tanggal Info': item.report_date || '',
        'Cabang': item.branch_name || '',
        'Kategori': item.category || '',
        'Sumber Laporan': item.source || '',
        'Keluhan': item.complaint || '',
        'Nama FC / Security': item.employee_name || '',
        'FC Spesialis': item.fc_specialist || '',
        'Solusi / Tindak Lanjut': item.solution || '',
        'Status': item.status || '',
        'Tanggal Selesai': item.completion_date || '',
        'Durasi Hari': item.day_count != null ? item.day_count : ''
      }));
      downloadExcel(exportData, `Data_Permasalahan_FCMS_${new Date().toISOString().split('T')[0]}`);
      toastSuccess('Berhasil mengekspor data permasalahan ke Excel');
    } catch (err) {
      toastError('Gagal ekspor Excel: ' + err.message);
    }
  }

  // Quick Status Updater
  async function quickUpdateStatus(issue, newStatus) {
    if (!canEdit) return;
    const currentRowVersion = parseValidRowVersion(issue?.row_version);
    if (currentRowVersion === null) {
      toastError('Versi data (row_version) tidak tersedia. Silakan muat ulang data sebelum memperbarui status.');
      return;
    }
    try {
      const completion_date = newStatus === 'Done' 
        ? (issue.completion_date || new Date().toISOString().split('T')[0]) 
        : (newStatus === 'Open' ? null : issue.completion_date);

      let day_count = issue.day_count;
      if (issue.report_date && completion_date) {
        const d1 = new Date(issue.report_date);
        const d2 = new Date(completion_date);
        day_count = Math.max(0, Math.floor((d2 - d1) / (1000 * 60 * 60 * 24)));
      } else if (newStatus === 'Open') {
        day_count = null;
      }

      const payload = {
        ...issue,
        row_version: currentRowVersion,
        status: newStatus,
        completion_date,
        day_count
      };

      const res = await apiFetch(`/api/issues/${issue.id}`, {
        method: 'PUT',
        body: JSON.stringify(payload)
      });

      if (res.ok) {
        const nextRowVersion = parseValidRowVersion(res.data?.data?.row_version ?? res.data?.row_version);
        if (nextRowVersion !== null) {
          issue.row_version = nextRowVersion;
        }
        toastSuccess(`Status masalah diubah ke "${newStatus}"`);
        await loadData();
      } else if (res.status === 409) {
        toastError('Konflik data (HTTP 409): Issue ini telah diubah oleh pengguna atau sesi lain. Silakan muat ulang data terbaru.');
        await loadData();
      } else {
        toastError(res.data?.error || 'Gagal memperbarui status');
      }
    } catch (err) {
      toastError('Terjadi kesalahan: ' + err.message);
    }
  }

  // Modal Tambah / Edit
  function openIssueModal(item = null) {
    const isEdit = !!item;
    if (isEdit ? !canEdit : !canCreate) return;
    const initialRowVersion = isEdit ? parseValidRowVersion(item?.row_version) : null;
    const today = new Date().toISOString().split('T')[0];
    const defaultDate = item?.report_date || today;
    const defaultCategory = item?.category || 'SDM';
    const defaultSource = item?.source ? String(item.source).trim() : '';
    const hasLegacySource = isEdit && defaultSource && !SOURCE_LIST.includes(defaultSource);
    const defaultStatus = item?.status || 'Open';

    const content = document.createElement('div');
    content.className = 'issue-modal-body';
    content.innerHTML = `
      <div style="display:flex; flex-direction:column; gap:14px;">
        ${isEdit ? `<input type="hidden" id="iss-modal-row-version" value="${initialRowVersion !== null ? initialRowVersion : ''}" />` : ''}
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Tanggal Info <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="iss-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${defaultDate}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Cabang Klinik <span style="color:#DC2626;">*</span>
            </label>
            <select id="iss-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${branchOptions.map(b => `
                <option value="${b.value}" ${String(item?.branch_id) === String(b.value) ? 'selected' : ''}>
                  ${b.label}
                </option>
              `).join('')}
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Kategori <span style="color:#DC2626;">*</span>
            </label>
            <select id="iss-modal-category" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${CATEGORY_LIST.map(cat => `
                <option value="${cat}" ${defaultCategory === cat ? 'selected' : ''}>${cat}</option>
              `).join('')}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Sumber Laporan
            </label>
            <select id="iss-modal-source" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;">
              <option value="">-- Pilih Sumber Laporan --</option>
              ${SOURCE_LIST.map(s => `<option value="${s}" ${defaultSource === s ? 'selected' : ''}>${s}</option>`).join('')}
              ${hasLegacySource ? `<option value="${defaultSource.replace(/"/g, '&quot;')}" selected>${defaultSource} (Historis)</option>` : ''}
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
            Detail Keluhan / Permasalahan <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="iss-modal-complaint" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Jelaskan kendala, kronologi, atau komplain..." required>${item?.complaint || ''}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Nama Personil FC / Security
            </label>
            <div id="iss-modal-emp-mount"></div>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              FC Spesialis / Penanggung Jawab
            </label>
            <input list="iss-spec-options" id="iss-modal-spec" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: Mizwar, Ade, Berlin..." value="${item?.fc_specialist ? String(item.fc_specialist).replace(/"/g, '&quot;') : ''}" />
            <datalist id="iss-spec-options">
              ${employeeOptions.map(e => `<option value="${e.full_name || e.label || ''}">`).join('')}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
            Solusi / Tindak Lanjut Penanganan
          </label>
          <textarea id="iss-modal-solution" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tindakan yang telah atau akan diambil untuk menyelesaikan masalah...">${item?.solution || ''}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr 1fr; gap:10px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Status Masalah <span style="color:#DC2626;">*</span>
            </label>
            <select id="iss-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${STATUS_LIST.map(st => `
                <option value="${st}" ${defaultStatus === st ? 'selected' : ''}>${st}</option>
              `).join('')}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Tanggal Selesai
            </label>
            <input type="date" id="iss-modal-completion" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${item?.completion_date || ''}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Durasi (Hari)
            </label>
            <input type="number" id="iss-modal-daycount" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #29364B; border-radius:8px; background:#1B2940; color:#F8FAFC;" 
              value="${item?.day_count != null ? item.day_count : ''}" readonly placeholder="Otomatis" />
          </div>
        </div>
      </div>
    `;

    // Dynamic auto-calculate day_count
    const dateInput = content.querySelector('#iss-modal-date');
    const compInput = content.querySelector('#iss-modal-completion');
    const dayInput = content.querySelector('#iss-modal-daycount');
    const statusSelect = content.querySelector('#iss-modal-status');

    function updateDayCount() {
      if (dateInput.value && compInput.value) {
        const d1 = new Date(dateInput.value);
        const d2 = new Date(compInput.value);
        const diff = Math.floor((d2 - d1) / (1000 * 60 * 60 * 24));
        dayInput.value = !isNaN(diff) ? Math.max(0, diff) : '';
      } else {
        dayInput.value = '';
      }
    }

    dateInput.addEventListener('change', updateDayCount);
    compInput.addEventListener('change', updateDayCount);
    statusSelect.addEventListener('change', () => {
      if (statusSelect.value === 'Done' && !compInput.value) {
        compInput.value = new Date().toISOString().split('T')[0];
        updateDayCount();
      } else if (statusSelect.value === 'Open') {
        compInput.value = '';
        dayInput.value = '';
      }
    });

    const empCombobox = mountEmployeeCombobox(content.querySelector('#iss-modal-emp-mount'), {
      mode: 'single',
      inputId: 'iss-modal-employee',
      inputName: 'employee_name',
      placeholder: 'Cari nama karyawan...',
      allowManual: true,
      hasEmployeeIdSchema: false,
      employees: employeeOptions,
      initialEmployeeName: item?.employee_name || '',
    });

    let isSubmitting = false;

    createModal({
      title: isEdit ? '✏️ Edit Data Permasalahan' : '➕ Tambah Data Permasalahan',
      content,
      size: 'md',
      confirmText: isEdit ? 'Simpan Perubahan' : 'Simpan Permasalahan',
      confirmClass: 'btn-primary',
      onConfirm: async (overlay, close) => {
        if (isSubmitting) return;

        const report_date = content.querySelector('#iss-modal-date').value;
        const branch_id = content.querySelector('#iss-modal-branch').value;
        const category = content.querySelector('#iss-modal-category').value;
        const source = (content.querySelector('#iss-modal-source').value || '').trim();
        const complaint = (content.querySelector('#iss-modal-complaint').value || '').trim();
        const empSelection = empCombobox.getValue();
        const employee_name = empSelection.employee_name || '';
        const fc_specialist = (content.querySelector('#iss-modal-spec').value || '').trim();
        const solution = (content.querySelector('#iss-modal-solution').value || '').trim();
        const status = content.querySelector('#iss-modal-status').value;
        const completion_date = content.querySelector('#iss-modal-completion').value || null;

        if (!report_date || !branch_id || !complaint || !category) {
          toastError('Harap lengkapi Tanggal, Cabang, Kategori, dan Keluhan!');
          return;
        }

        let editRowVersion = null;
        if (isEdit) {
          const rowVersionEl = content.querySelector('#iss-modal-row-version');
          const rawVersion = rowVersionEl ? rowVersionEl.value : initialRowVersion;
          editRowVersion = parseValidRowVersion(rawVersion);
          if (editRowVersion === null) {
            toastError('Versi data (row_version) tidak tersedia. Silakan muat ulang data sebelum menyimpan perubahan.');
            return;
          }
        }

        let day_count = null;
        if (report_date && completion_date) {
          const d1 = new Date(report_date);
          const d2 = new Date(completion_date);
          day_count = Math.max(0, Math.floor((d2 - d1) / (1000 * 60 * 60 * 24)));
        }

        const payload = {
          report_date,
          branch_id: Number(branch_id),
          category,
          source,
          complaint,
          employee_name,
          fc_specialist,
          solution,
          status,
          completion_date,
          day_count,
          ...(isEdit ? { row_version: editRowVersion } : {})
        };

        const confirmBtn = overlay?.querySelector?.('.modal-confirm');
        isSubmitting = true;
        if (confirmBtn) confirmBtn.disabled = true;

        try {
          let res;
          if (isEdit) {
            res = await apiFetch(`/api/issues/${item.id}`, {
              method: 'PUT',
              body: JSON.stringify(payload)
            });
          } else {
            res = await apiFetch('/api/issues', {
              method: 'POST',
              body: JSON.stringify(payload)
            });
          }

          if (res.ok) {
            if (isEdit) {
              const nextRowVersion = parseValidRowVersion(res.data?.data?.row_version ?? res.data?.row_version);
              if (nextRowVersion !== null && item) {
                item.row_version = nextRowVersion;
              }
            } else {
              currentPage = 1;
            }
            toastSuccess(isEdit ? 'Permasalahan berhasil diperbarui!' : 'Permasalahan baru berhasil ditambahkan!');
            close();
            await loadData();
          } else if (isEdit && res.status === 409) {
            toastError('Konflik data (HTTP 409): Issue ini telah diubah oleh pengguna atau sesi lain. Silakan muat ulang data sebelum mengedit kembali.');
            close();
            await loadData();
          } else {
            toastError(res.data?.error || 'Gagal menyimpan permasalahan');
          }
        } catch (err) {
          toastError('Terjadi kesalahan: ' + err.message);
        } finally {
          isSubmitting = false;
          if (confirmBtn) confirmBtn.disabled = false;
        }
      }
    });
  }

  // Delete Issue
  async function deleteIssue(item) {
    if (!canDelete) return;
    confirmDialog({
      title: 'Hapus Permasalahan?',
      message: `Apakah Anda yakin ingin menghapus tiket permasalahan dari <strong>${item.branch_name || 'cabang'}</strong>:<br><br><em>"${item.complaint}"</em>?`,
      confirmText: 'Ya, Hapus',
      confirmClass: 'btn-danger',
      onConfirm: async () => {
        try {
          const res = await apiFetch(`/api/issues/${item.id}`, {
            method: 'DELETE'
          });
          if (res.ok) {
            toastSuccess('Permasalahan berhasil dihapus');
            await loadData();
          } else {
            toastError(res.data?.error || 'Gagal menghapus data');
          }
        } catch (err) {
          toastError('Terjadi kesalahan: ' + err.message);
        }
      }
    });
  }

  // Filter Data (server already filters & paginates)
  function getFilteredList() {
    return issuesList;
  }


  // Render only the dynamic cards grid content HTML (Cards / Empty State)
  function buildIssuesListHTML(filtered) {
    if (filtered.length === 0) {
      return `
            <div class="issues-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">🔍</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#F8FAFC; margin:0 0 6px 0;">Tidak ada permasalahan yang cocok</h3>
              <p style="font-size:0.85rem; color:#94A3B8; margin:0 0 16px 0;">Coba sesuaikan kata kunci pencarian atau filter status yang dipilih.</p>
              <button id="issues-reset-filter" class="btn btn-secondary" style="font-size:0.82rem; padding:7px 16px; border-radius:8px;">Reset Semua Filter</button>
            </div>
      `;
    }

    return filtered.map(item => {
      const cat = item.category || 'Lainnya';
      const catClass = `issues-category-${cat.toLowerCase().replace(/[^a-z0-9]/g, '')}` || 'issues-category-lainnya';
      
      const st = item.status || 'Open';
      let stClass = 'issues-status-open';
      let stText = '⚠️ Open';
      if (st === 'In Progress') {
        stClass = 'issues-status-progress';
        stText = '⏳ In Progress';
      } else if (st === 'Done') {
        stClass = 'issues-status-done';
        stText = '✅ Done';
      }

      const priority = item.priority || 'Normal';
      let prClass = 'priority-normal';
      if (priority === 'High' || priority === 'Tinggi') prClass = 'priority-high';
      if (priority === 'Low' || priority === 'Rendah') prClass = 'priority-low';

      return `
        <div class="issues-card" data-id="${item.id}">
          <div class="issues-card-top">
            <span class="issues-category-badge ${catClass}">
              ${cat === 'SDM' ? '👥' : (cat === 'Cleaning' ? '🧹' : (cat === 'Aset' ? '📦' : '📋'))} ${cat}
            </span>
            <div style="display:flex; align-items:center; gap:6px;">
              <span class="issues-priority-badge ${prClass}">${priority}</span>
              <span class="issues-status-badge ${stClass}">${stText}</span>
            </div>
          </div>

          <div>
            <div class="issues-title" style="color:#F8FAFC; font-weight:700; font-size:0.96rem; line-height:1.45;">${item.complaint || 'Tidak ada uraian permasalahan'}</div>
            <div style="font-size:0.78rem; color:#CBD5E1; margin-top:4px;">
              ID #${item.id} • Sumber: <strong style="color:#E2E8F0;">${item.source || '-'}</strong>
            </div>
          </div>

          <div class="issues-details-grid">
            <div class="issues-detail-item">
              <span>🏢</span>
              <span style="font-weight:600; color:#F8FAFC; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="${item.branch_name || '-'}">
                ${item.branch_name || 'Semua Cabang'}
              </span>
            </div>
            <div class="issues-detail-item">
              <span>📅</span>
              <span style="color:#E2E8F0; font-weight:600;">${item.report_date ? formatDateIndo(item.report_date) : '-'}</span>
            </div>
            <div class="issues-detail-item">
              <span>👤</span>
              <span style="color:#F8FAFC; font-weight:600; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;" title="${item.employee_name || '-'}">
                ${item.employee_name || '-'}
              </span>
            </div>
            <div class="issues-detail-item">
              <span>🕒</span>
              <span style="color:#CBD5E1;">${item.completion_date ? formatDateIndo(item.completion_date) : (item.created_at ? formatDateIndo(item.created_at.slice(0, 10)) : '-')}</span>
            </div>
          </div>

          ${item.solution ? `
            <div class="issues-solution-box">
              <div class="issues-solution-title">
                <span>💡</span> Solusi / Tindakan:
              </div>
              <div>${item.solution}</div>
            </div>
          ` : ''}

          <div class="issues-card-footer">
            <div style="display:flex; align-items:center; gap:6px;">
              ${canEdit ? `
                ${st === 'Open' ? `
                  <button class="issues-quick-btn start" data-action="quick-progress" title="Mulai Penanganan">
                    ▶️ Proses
                  </button>
                ` : ''}
                ${st === 'In Progress' ? `
                  <button class="issues-quick-btn complete" data-action="quick-done" title="Tandai Selesai">
                    ✅ Selesai
                  </button>
                ` : ''}
                ${st === 'Done' ? `
                  <button class="issues-quick-btn reopen" data-action="quick-reopen" title="Buka Kembali Masalah">
                    🔄 Reopen
                  </button>
                ` : ''}
              ` : ''}
            </div>

            <div class="issues-action-btns">
              ${canEdit ? `
                <button class="issues-icon-btn edit" data-action="edit" data-id="${item.id}" title="Edit Masalah">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                </button>
              ` : ''}
              ${canDelete ? `
                <button class="issues-icon-btn delete" data-action="delete" data-id="${item.id}" title="Hapus Masalah">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                </button>
              ` : ''}
            </div>
          </div>
        </div>
      `;
    }).join('');
  }

  function bindListEvents() {
    const resetBtn = container.querySelector('#issues-reset-filter');
    if (resetBtn) {
      resetBtn.addEventListener('click', () => {
        searchQuery = '';
        selectedBranch = 'all';
        selectedCategory = 'all';
        selectedStatus = 'all';
        const searchInput = container.querySelector('#issues-search-input');
        if (searchInput) searchInput.value = '';
        const branchFilter = container.querySelector('#issues-branch-filter');
        if (branchFilter) branchFilter.value = 'all';
        const catFilter = container.querySelector('#issues-category-filter');
        if (catFilter) catFilter.value = 'all';
        container.querySelectorAll('.issues-pill').forEach(p => {
          p.classList.toggle('active', p.getAttribute('data-status') === 'all');
        });
        currentPage = 1;
        loadData({ refreshStats: false });
      });
    }

    container.querySelectorAll('.issues-card').forEach(card => {
      const id = Number(card.dataset.id);
      const item = issuesList.find(x => x.id === id);
      if (!item) return;

      const editBtn = card.querySelector('[data-action="edit"]');
      if (editBtn) {
        editBtn.addEventListener('click', () => openIssueModal(item));
      }

      const delBtn = card.querySelector('[data-action="delete"]');
      if (delBtn) {
        delBtn.addEventListener('click', () => deleteIssue(item));
      }

      const progBtn = card.querySelector('[data-action="quick-progress"]');
      if (progBtn) {
        progBtn.addEventListener('click', () => quickUpdateStatus(item, 'In Progress'));
      }

      const doneBtn = card.querySelector('[data-action="quick-done"]');
      if (doneBtn) {
        doneBtn.addEventListener('click', () => quickUpdateStatus(item, 'Done'));
      }

      const reopenBtn = card.querySelector('[data-action="quick-reopen"]');
      if (reopenBtn) {
        reopenBtn.addEventListener('click', () => quickUpdateStatus(item, 'Open'));
      }
    });
  }

  function renderPagination() {
    const pagEl = container.querySelector('#issues-pagination-container');
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

  function renderListContent() {
    const listContainer = container.querySelector('#issues-list-container');
    if (!listContainer) return;
    const filtered = getFilteredList();
    listContainer.innerHTML = buildIssuesListHTML(filtered);
    renderPagination();
    bindListEvents();
  }
  // Main UI Renderer
  function renderUI() {
    const filtered = getFilteredList();

    // KPI Metrics from server stats
    const totalIssues = serverStats.total;
    const openAndProgress = serverStats.openAndProgress;
    const doneCount = serverStats.done;
    const bulanIniCount = serverStats.currentMonth;

    // Counts for status pills
    const countAll = serverStats.total;
    const countOpen = serverStats.open;
    const countInProgress = serverStats.inProgress;
    const countDone = serverStats.done;

    container.innerHTML = `
      <style>
        .issues-page {
          display: flex;
          flex-direction: column;
          gap: 20px;
        }
        .issues-hero-card {
          background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
          border-radius: 16px;
          padding: 24px 28px;
          color: white;
          box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25);
          position: relative;
          overflow: hidden;
        }
        .issues-hero-card::after {
          content: '';
          position: absolute;
          top: -40px;
          right: -40px;
          width: 220px;
          height: 220px;
          background: radial-gradient(circle, rgba(245, 158, 11, 0.18) 0%, transparent 70%);
          border-radius: 50%;
          pointer-events: none;
        }
        .issues-hero-actions {
          display: flex;
          align-items: center;
          gap: 8px;
          flex-wrap: wrap;
        }
        .issues-hero-actions .btn {
          height: 38px;
          box-sizing: border-box;
          transition: all 0.2s ease;
        }
        .issues-btn-add:hover {
          background: #1D4ED8 !important;
          transform: translateY(-1px);
        }
        .issues-btn-excel:hover, .issues-btn-refresh:hover {
          background: rgba(255, 255, 255, 0.16) !important;
          border-color: rgba(255, 255, 255, 0.3) !important;
          transform: translateY(-1px);
        }
        .issues-stats-row {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 14px;
          margin-top: 20px;
        }
        .issues-stat-box {
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
        .issues-stat-box:hover {
          background: rgba(255, 255, 255, 0.08);
          transform: translateY(-2px);
        }
        .issues-stat-icon {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 1.3rem;
          flex-shrink: 0;
        }
        .issues-stat-val {
          font-size: 1.55rem;
          font-weight: 800;
          line-height: 1.1;
          letter-spacing: -0.02em;
        }
        .issues-stat-lbl {
          font-size: 0.72rem;
          text-transform: uppercase;
          color: #94A3B8;
          font-weight: 700;
          letter-spacing: 0.05em;
          margin-top: 2px;
        }

        /* Controls / Filter Bar */
        .issues-controls-card {
          background: #172033;
          border-radius: 14px;
          border: 1px solid #29364B;
          padding: 16px 20px;
          box-shadow: 0 2px 8px rgba(0,0,0,0.03);
          display: flex;
          flex-direction: column;
          gap: 14px;
        }
        .issues-search-row {
          display: flex;
          gap: 12px;
          flex-wrap: wrap;
        }
        .issues-search-wrap {
          position: relative;
          flex: 1;
          min-width: 240px;
        }
        .issues-search-icon {
          position: absolute;
          left: 14px;
          top: 50%;
          transform: translateY(-50%);
          color: #94A3B8;
          pointer-events: none;
        }
        .issues-search-input {
          width: 100%;
          padding: 10px 14px 10px 40px;
          background: #1B2940;
          color: #F8FAFC;
          border: 1px solid #29364B;
          border-radius: 10px;
          font-size: 0.9rem;
          outline: none;
          transition: border-color 0.15s ease;
        }
        .issues-search-input:focus {
          border-color: #3B82F6;
          box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.2);
        }
        .issues-select {
          padding: 10px 14px;
          border: 1px solid #29364B;
          border-radius: 10px;
          font-size: 0.9rem;
          color: #F8FAFC;
          background: #1B2940;
          outline: none;
          min-width: 160px;
          cursor: pointer;
        }
        .issues-select:focus {
          border-color: #3B82F6;
        }

        /* Swipeable Pills */
        .issues-filter-pills {
          display: flex;
          align-items: center;
          gap: 8px;
          overflow-x: auto;
          white-space: nowrap;
          scrollbar-width: none;
          padding-bottom: 2px;
        }
        .issues-filter-pills::-webkit-scrollbar {
          display: none;
        }
        .issues-pill {
          padding: 6px 14px;
          border-radius: 20px;
          font-size: 0.82rem;
          font-weight: 600;
          cursor: pointer;
          border: 1px solid #29364B;
          background: #1B2940;
          color: #CBD5E1;
          transition: all 0.15s ease;
          display: inline-flex;
          align-items: center;
          gap: 6px;
          user-select: none;
          flex-shrink: 0;
        }
        .issues-pill:hover {
          background: #22314A;
          color: #F8FAFC;
        }
        .issues-pill.active {
          background: #2563EB;
          color: #FFFFFF;
          border-color: #2563EB;
        }
        .issues-pill-dot {
          width: 8px;
          height: 8px;
          border-radius: 50%;
        }
        .issues-pill-count {
          padding: 1px 6px;
          border-radius: 10px;
          font-size: 0.72rem;
          background: rgba(255,255,255,0.1);
          color: #E2E8F0;
        }
        .issues-pill.active .issues-pill-count {
          background: rgba(255,255,255,0.25);
          color: white;
        }

        /* Cards Grid */
        .issues-cards-grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(370px, 1fr));
          gap: 16px;
        }
        .issues-card {
          background: #172033;
          border-radius: 14px;
          border: 1px solid #29364B;
          padding: 18px;
          box-shadow: 0 2px 8px rgba(0,0,0,0.2);
          display: flex;
          flex-direction: column;
          gap: 14px;
          transition: transform 0.15s ease, box-shadow 0.15s ease;
        }
        .issues-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 6px 16px rgba(0,0,0,0.32);
          border-color: #475569;
        }
        .issues-card-top {
          display: flex;
          justify-content: space-between;
          align-items: center;
          gap: 8px;
        }
        .issues-category-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.76rem;
          font-weight: 700;
        }
        .issues-category-sdm {
          background: rgba(99, 102, 241, 0.18);
          color: #A5B4FC;
          border: 1px solid rgba(129, 140, 248, 0.35);
        }
        .issues-category-cleaning {
          background: rgba(20, 184, 166, 0.16);
          color: #5EEAD4;
          border: 1px solid rgba(45, 212, 191, 0.35);
        }
        .issues-category-aset {
          background: rgba(139, 92, 246, 0.18);
          color: #C4B5FD;
          border: 1px solid rgba(167, 139, 250, 0.35);
        }
        .issues-category-k3 {
          background: rgba(245, 158, 11, 0.18);
          color: #FCD34D;
          border: 1px solid rgba(251, 191, 36, 0.35);
        }
        .issues-category-lainnya {
          background: rgba(148, 163, 184, 0.16);
          color: #E2E8F0;
          border: 1px solid rgba(148, 163, 184, 0.3);
        }

        .issues-status-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.76rem;
          font-weight: 700;
        }
        .issues-status-open {
          background: rgba(245, 158, 11, 0.18);
          color: #FCD34D;
          border: 1px solid rgba(251, 191, 36, 0.38);
        }
        .issues-status-inprogress,
        .issues-status-progress {
          background: rgba(59, 130, 246, 0.18);
          color: #93C5FD;
          border: 1px solid rgba(96, 165, 250, 0.38);
        }
        .issues-status-done {
          background: rgba(16, 185, 129, 0.18);
          color: #6EE7B7;
          border: 1px solid rgba(52, 211, 153, 0.38);
        }

        /* Complaint Quote Box */
        .issues-complaint-box {
          background: #1B2940;
          border-left: 4px solid #F59E0B;
          border-radius: 0 10px 10px 0;
          padding: 12px 14px;
          font-size: 0.92rem;
          color: #F8FAFC;
          line-height: 1.5;
          font-weight: 600;
        }
        .issues-complaint-box.status-done {
          border-left-color: #10B981;
        }
        .issues-complaint-box.status-inprogress {
          border-left-color: #3B82F6;
        }

        /* Person / Location Box */
        .issues-person-box {
          background: #1B2940;
          border: 1px solid #29364B;
          border-radius: 10px;
          padding: 12px;
          display: flex;
          align-items: center;
          gap: 12px;
        }
        .issues-avatar-box {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          background: #4F46E5;
          color: white;
          display: flex;
          align-items: center;
          justify-content: center;
          font-weight: 800;
          font-size: 1rem;
          flex-shrink: 0;
        }
        .issues-person-name {
          font-size: 0.95rem;
          font-weight: 800;
          color: #F8FAFC;
        }
        .issues-person-sub {
          font-size: 0.74rem;
          color: #94A3B8;
          font-weight: 600;
          text-transform: uppercase;
          letter-spacing: 0.04em;
        }

        /* Details list */
        .issues-details-grid {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 8px 12px;
          font-size: 0.82rem;
          color: #E2E8F0;
          padding: 2px 4px;
        }
        .issues-detail-item {
          display: flex;
          align-items: center;
          gap: 6px;
        }

        /* Solution box */
        .issues-solution-box {
          background: rgba(16, 185, 129, 0.12);
          border: 1px solid rgba(52, 211, 153, 0.28);
          border-radius: 10px;
          padding: 10px 12px;
          font-size: 0.83rem;
          color: #E2E8F0;
          line-height: 1.5;
        }
        .issues-solution-title {
          font-weight: 700;
          color: #6EE7B7;
          display: flex;
          align-items: center;
          gap: 6px;
          margin-bottom: 4px;
        }

        /* Card footer */
        .issues-card-footer {
          display: flex;
          justify-content: space-between;
          align-items: center;
          border-top: 1px solid #29364B;
          padding-top: 12px;
          margin-top: auto;
          gap: 8px;
        }
        .issues-action-btns {
          display: flex;
          align-items: center;
          gap: 6px;
        }
        .issues-icon-btn {
          width: 34px;
          height: 34px;
          border-radius: 8px;
          border: 1px solid #29364B;
          background: #1B2940;
          color: #CBD5E1;
          display: flex;
          align-items: center;
          justify-content: center;
          cursor: pointer;
          transition: all 0.15s ease;
        }
        .issues-icon-btn:hover {
          background: #22314A;
          color: #F8FAFC;
          border-color: #475569;
        }
        .issues-icon-btn.delete:hover {
          background: rgba(239, 68, 68, 0.2);
          color: #FCA5A5;
          border-color: rgba(248, 113, 113, 0.45);
        }
        .issues-quick-btn {
          font-size: 0.76rem;
          font-weight: 700;
          padding: 6px 10px;
          border-radius: 8px;
          cursor: pointer;
          border: 1px solid transparent;
          display: inline-flex;
          align-items: center;
          gap: 4px;
          transition: all 0.15s ease;
        }
        .issues-quick-btn.start {
          background: rgba(59, 130, 246, 0.16);
          color: #93C5FD;
          border-color: rgba(96, 165, 250, 0.35);
        }
        .issues-quick-btn.start:hover {
          background: rgba(59, 130, 246, 0.26);
          color: #FFFFFF;
        }
        .issues-quick-btn.complete {
          background: rgba(16, 185, 129, 0.16);
          color: #6EE7B7;
          border-color: rgba(52, 211, 153, 0.35);
        }
        .issues-quick-btn.complete:hover {
          background: rgba(16, 185, 129, 0.26);
          color: #FFFFFF;
        }
        .issues-quick-btn.reopen {
          background: rgba(245, 158, 11, 0.16);
          color: #FCD34D;
          border-color: rgba(251, 191, 36, 0.35);
        }
        .issues-quick-btn.reopen:hover {
          background: rgba(245, 158, 11, 0.26);
          color: #FFFFFF;
        }

        /* Empty state */
        .issues-empty-state {
          grid-column: 1 / -1;
          background: #172033;
          border: 2px dashed #29364B;
          border-radius: 16px;
          padding: 48px 24px;
          text-align: center;
        }

        /* Mobile Responsive */
        @media (max-width: 768px) {
          .issues-hero-card {
            padding: 18px 16px;
          }
          .issues-stats-row {
            grid-template-columns: 1fr 1fr;
            gap: 10px;
          }
          .issues-stat-box {
            padding: 10px 12px;
          }
          .issues-stat-icon {
            width: 38px;
            height: 38px;
            font-size: 1.1rem;
          }
          .issues-stat-val {
            font-size: 1.3rem;
          }
          .issues-hero-actions {
            width: 100% !important;
            display: flex !important;
            align-items: center !important;
            gap: 8px !important;
            flex-wrap: wrap !important;
            overflow-x: visible !important;
          }
          .issues-hero-actions::-webkit-scrollbar {
            display: none;
          }
          .issues-btn-add {
            width: 100% !important;
            justify-content: center !important;
            padding: 9px 12px !important;
            font-size: 0.85rem !important;
            white-space: nowrap !important;
            height: 40px !important;
          }
          .issues-btn-excel {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 12px !important;
            font-size: 0.85rem !important;
            white-space: nowrap !important;
            height: 38px !important;
          }
          .issues-btn-refresh {
            width: 38px !important;
            height: 38px !important;
            padding: 0 !important;
            flex-shrink: 0 !important;
            justify-content: center !important;
          }
          .issues-controls-card {
            padding: 14px;
            gap: 12px;
          }
          .issues-search-row {
            flex-direction: column;
            gap: 10px;
          }
          .issues-select {
            width: 100%;
          }
          .issues-cards-grid {
            grid-template-columns: minmax(0, 1fr) !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
          }
          .issues-card {
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
            overflow: hidden;
            padding: 14px !important;
          }
        }
      </style>

      <div class="issues-page">

        <!-- Hero Header Card -->
        <div class="issues-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px;">
                <span style="font-size:1.6rem;">⚠️</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Log Permasalahan Operasional
                </h1>
              </div>
            </div>

            <div class="issues-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${canCreate ? `
                <button id="issues-btn-add" class="btn btn-primary issues-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Permasalahan</span>
                </button>
              ` : ''}
              
              ${hasPermission(currentUser, "export") ? `<button id="issues-btn-excel" class="btn btn-ghost issues-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>📥</span>
                <span>Ekspor Excel</span>
              </button>` : ''}

              <button id="issues-btn-refresh" class="btn btn-ghost issues-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:inline-flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24" class="${isLoading ? 'spin-icon' : ''}"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- 4 KPI Stat Boxes -->
          <div class="issues-stats-row">
            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(239, 68, 68, 0.2); color:#FCA5A5;">⚠️</div>
              <div>
                <div class="issues-stat-val" id="issues-stat-total">${totalIssues}</div>
                <div class="issues-stat-lbl">Total Permasalahan</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(245, 158, 11, 0.2); color:#FCD34D;">⏳</div>
              <div>
                <div class="issues-stat-val" id="issues-stat-open-prog">${openAndProgress}</div>
                <div class="issues-stat-lbl">Open & In Progress</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#6EE7B7;">✅</div>
              <div>
                <div class="issues-stat-val" id="issues-stat-done">${doneCount}</div>
                <div class="issues-stat-lbl">Tuntas (Done)</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(59, 130, 246, 0.2); color:#93C5FD;">📅</div>
              <div>
                <div class="issues-stat-val" id="issues-stat-month">${bulanIniCount}</div>
                <div class="issues-stat-lbl">Bulan Berjalan</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Filter Controls Card -->
        <div class="issues-controls-card">
          <div class="issues-search-row">
            <div class="issues-search-wrap">
              <svg class="issues-search-icon" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              <input type="text" id="issues-search-input" class="issues-search-input" 
                placeholder="Cari keluhan, nama personil, sumber, cabang..." value="${searchQuery ? String(searchQuery).replace(/"/g, '&quot;') : ''}" />
            </div>

            <select id="issues-branch-filter" class="issues-select">
              <option value="all">Semua Cabang (${branchOptions.length})</option>
              ${branchOptions.map(b => `
                <option value="${b.value}" ${String(selectedBranch) === String(b.value) ? 'selected' : ''}>
                  ${b.label}
                </option>
              `).join('')}
            </select>

            <select id="issues-category-filter" class="issues-select">
              <option value="all">Semua Kategori</option>
              ${CATEGORY_LIST.map(cat => `
                <option value="${cat}" ${selectedCategory === cat ? 'selected' : ''}>
                  Kategori: ${cat}
                </option>
              `).join('')}
            </select>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="issues-filter-pills">
            <div class="issues-pill ${selectedStatus === 'all' ? 'active' : ''}" data-status="all">
              <span>Semua Status</span>
              <span class="issues-pill-count" id="issues-pill-count-all">${countAll}</span>
            </div>
            <div class="issues-pill ${selectedStatus === 'Open' ? 'active' : ''}" data-status="Open">
              <span class="issues-pill-dot" style="background:#F59E0B;"></span>
              <span>Open</span>
              <span class="issues-pill-count" id="issues-pill-count-open">${countOpen}</span>
            </div>
            <div class="issues-pill ${selectedStatus === 'In Progress' ? 'active' : ''}" data-status="In Progress">
              <span class="issues-pill-dot" style="background:#3B82F6;"></span>
              <span>In Progress</span>
              <span class="issues-pill-count" id="issues-pill-count-inprogress">${countInProgress}</span>
            </div>
            <div class="issues-pill ${selectedStatus === 'Done' ? 'active' : ''}" data-status="Done">
              <span class="issues-pill-dot" style="background:#10B981;"></span>
              <span>Done</span>
              <span class="issues-pill-count" id="issues-pill-count-done">${countDone}</span>
            </div>
          </div>
        </div>

        <!-- Exclusive Cards Grid -->
        <div id="issues-list-container" class="issues-cards-grid">
          ${buildIssuesListHTML(filtered)}
        </div>

        <!-- Pagination Container -->
        <div id="issues-pagination-container"></div>

      </div>
    `;

    // ── Event Handlers ──
    const searchInput = container.querySelector('#issues-search-input');
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

    const branchFilter = container.querySelector('#issues-branch-filter');
    if (branchFilter) {
      branchFilter.addEventListener('change', (e) => {
        selectedBranch = e.target.value;
        currentPage = 1;
        loadData({ refreshStats: false });
      });
    }

    const categoryFilter = container.querySelector('#issues-category-filter');
    if (categoryFilter) {
      categoryFilter.addEventListener('change', (e) => {
        selectedCategory = e.target.value;
        currentPage = 1;
        loadData({ refreshStats: false });
      });
    }

    // Status Pills
    container.querySelectorAll('.issues-pill').forEach(pill => {
      pill.addEventListener('click', () => {
        selectedStatus = pill.dataset.status;
        container.querySelectorAll('.issues-pill').forEach(p => {
          p.classList.toggle('active', p === pill);
        });
        currentPage = 1;
        loadData({ refreshStats: false });
      });
    });

    // Reset button
    const resetBtn = container.querySelector('#issues-reset-filter');
    if (resetBtn) {
      resetBtn.addEventListener('click', () => {
        searchQuery = '';
        selectedBranch = 'all';
        selectedCategory = 'all';
        selectedStatus = 'all';
        currentPage = 1;
        clearTimeout(searchDebounce);
        if (searchInput) searchInput.value = '';
        if (branchFilter) branchFilter.value = 'all';
        if (categoryFilter) categoryFilter.value = 'all';
        container.querySelectorAll('.issues-pill').forEach(p => {
          p.classList.toggle('active', p.dataset.status === 'all');
        });
        loadData({ refreshStats: false });
      });
    }

    // Top action buttons
    const addBtn = container.querySelector('#issues-btn-add');
    if (addBtn) {
      addBtn.addEventListener('click', () => openIssueModal());
    }

    const excelBtn = container.querySelector('#issues-btn-excel');
    if (excelBtn) {
      excelBtn.addEventListener('click', () => exportToExcel());
    }

    const refreshBtn = container.querySelector('#issues-btn-refresh');
    if (refreshBtn) {
      refreshBtn.addEventListener('click', () => loadData({ refreshStats: true }));
    }

    renderPagination();
    bindListEvents();
  }

  // Initial load
  await loadData({ refreshStats: true });
}
