import { apiFetch, getUser, hasPermission } from '../config.js';
import { getCachedBranches, getCachedEmployees, fetchExportRows } from '../utils/dataCache.js';
import { createModal, confirmDialog } from '../components/modal.js';
import { toastSuccess, toastError, toastInfo } from '../components/toast.js';
import { createPagination } from '../components/table.js';
import { mountEmployeeCombobox } from '../components/employeeCombobox.js';
import { formatDateIndo } from '../utils/sorting.js';

// Pre-seeded fallback data jika API offline / inisialisasi awal
const INITIAL_MUTASI = [
  { id: 1, tanggal: '2026-10-02', employee_name: 'TAUFIK HIDAYAT', from_branch_id: 12, from_branch_name: '012. BSD', to_branch_id: 0, to_branch_name: '000. Management', status: 'Selesai', document_link: '' },
  { id: 2, tanggal: '2026-09-30', employee_name: 'ELVAN JANGGA PRAMANA', from_branch_id: 55, from_branch_name: '055. Gajah Mada', to_branch_id: 0, to_branch_name: '000. Management', status: 'Proses', document_link: '' },
  { id: 3, tanggal: '2026-06-23', employee_name: 'Widya Astuti', from_branch_id: 1, from_branch_name: '001. Pondok Bambu', to_branch_id: 7, to_branch_name: '007. Bekasi', status: 'Selesai', document_link: '' },
  { id: 4, tanggal: '2026-06-22', employee_name: 'Gema Azan Saputra', from_branch_id: 1, from_branch_name: '001. Pondok Bambu', to_branch_id: 21, to_branch_name: '021. Kelapa Gading Boulevard', status: 'Selesai', document_link: '' },
  { id: 5, tanggal: '2026-06-22', employee_name: 'Ahmad Farhan', from_branch_id: 1, from_branch_name: '001. Pondok Bambu', to_branch_id: 62, to_branch_name: '062. Summarecon Bekasi', status: 'Selesai', document_link: '' }
];

// Standardized Indonesian date helper imported from ../utils/sorting.js

export async function renderMutasi(container) {
  const PAGE_LIMIT = 20;
  let mutasiList = [];
  let branchOptions = [];
  let employeeOptions = [];
  let selectedStatus = 'all'; // 'all' | 'Proses' | 'Selesai'
  let selectedFromBranch = 'all';
  let selectedToBranch = 'all';
  let searchQuery = '';
  let currentPage = 1;
  let totalItems = 0;
  let totalPages = 1;
  let statsLoaded = false;
  let serverStats = { total: 0, proses: 0, selesai: 0, currentMonth: 0 };
  let searchDebounce = null;
  let isLoading = false;
  let viewMode = localStorage.getItem('fcms_mutasi_view_mode') || 'card';

  const currentUser = getUser();
  const canAdd = hasPermission(currentUser, 'mutasi', 'write');
  const canEdit = hasPermission(currentUser, 'mutasi', 'write');
  const canDelete = hasPermission(currentUser, 'mutasi', 'delete');

  // Load branches and employees
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
    if (selectedFromBranch && selectedFromBranch !== 'all') filters.from_branch_id = selectedFromBranch;
    if (selectedToBranch && selectedToBranch !== 'all') filters.to_branch_id = selectedToBranch;
    if (selectedStatus && selectedStatus !== 'all') filters.status = selectedStatus;
    return filters;
  }

  function updateStatsUI() {
    const set = (id, val) => {
      const el = container.querySelector(id);
      if (el) el.textContent = val;
    };
    set('#mutasi-stat-total', serverStats.total);
    set('#mutasi-stat-proses', serverStats.proses);
    set('#mutasi-stat-selesai', serverStats.selesai);
    set('#mutasi-stat-month', serverStats.currentMonth);
    set('#mutasi-pill-count-all', serverStats.total);
    set('#mutasi-pill-count-proses', serverStats.proses);
    set('#mutasi-pill-count-selesai', serverStats.selesai);
  }

  // Load Mutasi Data from API (Server-Side Pagination)
  async function loadData({ refreshStats = true } = {}) {
    isLoading = true;
    const hasShell = Boolean(container.querySelector('#mutasi-list-container'));
    if (!hasShell) {
      renderUI();
    } else {
      const listEl = container.querySelector('#mutasi-list-container');
      if (listEl) {
        listEl.innerHTML = '<div class="mutasi-empty-state"><p style="color:var(--text-3); margin:0;">Memuat data mutasi...</p></div>';
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
      const res = await apiFetch(`/api/mutasi?${qs.toString()}`);
      if (res.ok) {
        const raw = res.data?.data || [];
        mutasiList = Array.isArray(raw) ? raw : [];
        const pag = res.data?.pagination || {};
        totalItems = Number(pag.total ?? mutasiList.length);
        totalPages = Math.max(1, Number(pag.pages ?? Math.ceil(totalItems / PAGE_LIMIT)) || 1);
        currentPage = Math.max(1, Number(pag.page ?? currentPage));
        if (res.data?.stats) {
          serverStats = res.data.stats;
          statsLoaded = true;
        }
        if (mutasiList.length === 0 && currentPage > totalPages) {
          currentPage = totalPages;
          isLoading = false;
          return loadData({ refreshStats });
        }
      } else {
        mutasiList = [];
        totalItems = 0;
        totalPages = 1;
        toastError(res.data?.error || 'Gagal memuat data mutasi');
      }
    } catch (e) {
      console.warn('Gagal load mutasi:', e);
      mutasiList = [];
      totalItems = 0;
      totalPages = 1;
      toastError('Gagal memuat data mutasi: ' + e.message);
    } finally {
      isLoading = false;
      if (container.querySelector('#mutasi-list-container')) {
        updateStatsUI();
        renderListContent();
      } else {
        renderUI();
      }
    }
  }

  // Export to Excel (All matching filtered rows via /api/mutasi/export)
  async function exportToExcel() {
    if (!hasPermission(currentUser, 'export')) {
      toastError('Anda tidak memiliki izin untuk mengekspor data.');
      return;
    }
    try {
      const [{ downloadExcel }, rows] = await Promise.all([
        import('../utils/excel.js'),
        fetchExportRows('/api/mutasi/export', getActiveFilters())
      ]);
      const exportData = rows.map((m, idx) => ({
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
    if (!canDelete) return;
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
            await loadData({ refreshStats: true });
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
    const isEditCheck = !!item;
    if (isEditCheck && !canEdit) {
      toastError('Anda tidak memiliki izin untuk mengedit data mutasi.');
      return;
    }
    if (!isEditCheck && !canAdd) {
      toastError('Anda tidak memiliki izin untuk menambah data mutasi.');
      return;
    }
    if (!canEdit) return;
    const isEdit = !!item;
    const defaultTanggal = item?.tanggal ? item.tanggal.slice(0, 10) : new Date().toISOString().slice(0, 10);

    const content = document.createElement('div');
    content.innerHTML = `
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:repeat(auto-fit, minmax(200px, 1fr)); gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Tanggal Mutasi <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="mutasi-modal-tanggal" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${defaultTanggal}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Status Mutasi <span style="color:#DC2626;">*</span>
            </label>
            <select id="mutasi-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="Proses" ${item?.status === 'Proses' ? 'selected' : ''}>⏳ Proses</option>
              <option value="Selesai" ${item?.status === 'Selesai' || !item ? 'selected' : ''}>✅ Selesai</option>
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
            Nama Karyawan <span style="color:#DC2626;">*</span>
          </label>
          <div id="mutasi-modal-emp-mount"></div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:repeat(auto-fit, minmax(200px, 1fr)); gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
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
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
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
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
            Link Berkas SK Mutasi / Dokumen (Opsional)
          </label>
          <input type="url" id="mutasi-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="https://drive.google.com/file/d/..." value="${item?.document_link ? String(item.document_link).replace(/"/g, '&quot;') : ''}" />
        </div>
      </div>
    `;

    const empCombobox = mountEmployeeCombobox(content.querySelector('#mutasi-modal-emp-mount'), {
      mode: 'single',
      inputId: 'mutasi-modal-emp',
      inputName: 'employee_name',
      placeholder: 'Cari nama karyawan...',
      required: true,
      allowManual: true,
      hasEmployeeIdSchema: true,
      manualBadgeText: 'Belum terhubung ke Master Karyawan',
      employees: employeeOptions,
      initialEmployeeId: item?.employee_id || null,
      initialEmployeeName: item?.employee_name || '',
    });

    createModal({
      title: isEdit ? '✏️ Edit Data Mutasi Karyawan' : '➕ Tambah Data Mutasi Karyawan',
      content,
      size: 'md',
      confirmText: isEdit ? 'Simpan Perubahan' : 'Simpan Mutasi',
      confirmClass: 'btn-primary',
      onConfirm: async (_, close) => {
        const tanggal = content.querySelector('#mutasi-modal-tanggal').value;
        const empSelection = empCombobox.getValue();
        const employee_name = (empSelection.employee_name || '').trim();
        const employee_id = empSelection.employee_id ?? null;
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
          empCombobox.inputEl?.focus();
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
          employee_id,
          employee_name,
          from_branch_id: Number(from_branch_id),
          from_branch_name: bFrom ? bFrom.label : 'Cabang Asal',
          to_branch_id: Number(to_branch_id),
          to_branch_name: bTo ? bTo.label : 'Cabang Tujuan',
          status,
          document_link,
          ...(isEdit ? { row_version: item?.row_version !== undefined && item?.row_version !== null ? Number(item.row_version) : 1 } : {})
        };

        try {
          if (isEdit) {
            const res = await apiFetch(`/api/mutasi/${item.id}`, {
              method: 'PUT',
              body: JSON.stringify(payload)
            });
            if (res.ok) {
              toastSuccess('Data mutasi berhasil diperbarui');
              close();
              await loadData({ refreshStats: true });
            } else {
              if (res.status === 409) {
                toastError('Konflik OCC: Data telah diubah oleh pengguna lain. Silakan muat ulang halaman.');
                close();
                await loadData({ refreshStats: true });
                return;
              }
              toastError(res.data?.error || 'Gagal memperbarui data mutasi');
            }
          } else {
            const res = await apiFetch('/api/mutasi', {
              method: 'POST',
              body: JSON.stringify(payload)
            });
            if (res.ok) {
              toastSuccess('Data mutasi baru berhasil ditambahkan');
              close();
              await loadData({ refreshStats: true });
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

  // Filter & Search Logic (server already filters & paginates)
  function getFilteredList() {
    return mutasiList;
  }


  // Render only the dynamic list content HTML (Card / Table / Empty State)
  function buildMutasiListHTML(filtered) {
    return `
          ${filtered.length === 0 ? `
            <div class="mutasi-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">🔍</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:var(--text-1); margin:0 0 6px 0;">Tidak Ada Data Mutasi Ditemukan</h3>
              <p style="color:var(--text-3); font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="mutasi-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          ` : viewMode === 'table' ? `
            <div class="mutasi-table-card">
              <div class="mutasi-table-responsive">
                <table class="mutasi-table">
                  <thead>
                    <tr>
                      <th style="width:45px; text-align:center;">No</th>
                      <th>Nama Karyawan</th>
                      <th>Tanggal Mutasi</th>
                      <th>Cabang Asal</th>
                      <th>Cabang Tujuan</th>
                      <th style="text-align:center;">Status</th>
                      <th style="text-align:center;">Dokumen SK</th>
                      ${canEdit || canDelete ? `<th style="width:90px; text-align:center;">Aksi</th>` : ''}
                    </tr>
                  </thead>
                  <tbody>
                    ${filtered.map((m, idx) => {
                      const isSelesai = m.status === 'Selesai';
                      const initials = (m.employee_name || 'Karyawan').split(' ').map(w => w[0]).filter(Boolean).slice(0, 2).join('').toUpperCase();
                      return `
                        <tr data-id="${m.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600; font-size:0.8rem;">${((currentPage - 1) * PAGE_LIMIT) + idx + 1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="mutasi-avatar-box" style="width:34px; height:34px; font-size:0.78rem; border-radius:8px;">
                                ${initials}
                              </div>
                              <div>
                                <div style="font-weight:700; color:#F8FAFC; white-space:nowrap;">${m.employee_name}</div>
                                <div style="font-size:0.72rem; color:#94A3B8;">ID Mutasi: #${m.id}</div>
                              </div>
                            </div>
                          </td>
                          <td style="white-space:nowrap; font-weight:600; color:#F8FAFC;">
                            🗓️ ${formatDateIndo(m.tanggal)}
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:600; color:#CBD5E1; font-size:0.85rem; display:inline-flex; align-items:center; gap:4px;">
                              <span>📍</span>
                              <span>${m.from_branch_name || '-'}</span>
                            </div>
                          </td>
                          <td style="white-space:nowrap;">
                            <div style="font-weight:700; color:#93C5FD; font-size:0.85rem; display:inline-flex; align-items:center; gap:5px;">
                              <span>➔</span>
                              <span>${m.to_branch_name || '-'}</span>
                            </div>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            <span class="mutasi-status-badge ${isSelesai ? 'selesai' : 'proses'}">
                              <span>${isSelesai ? '✅' : '⏳'}</span>
                              <span>${m.status || 'Proses'}</span>
                            </span>
                          </td>
                          <td style="text-align:center; white-space:nowrap;">
                            ${m.document_link ? `
                              <a href="${m.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.76rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                                <span>📄 Buka SK</span>
                              </a>
                            ` : `
                              <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                            `}
                          </td>
                          ${canEdit || canDelete ? `
                            <td style="text-align:center; white-space:nowrap;">
                              <div style="display:inline-flex; align-items:center; gap:6px;">
                                ${canEdit ? `
                                  <button class="mutasi-icon-btn mutasi-edit-btn" title="Edit Data Mutasi" data-id="${m.id}">
                                    <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                  </button>
                                ` : ''}
                                ${canDelete ? `
                                  <button class="mutasi-icon-btn delete mutasi-delete-btn" title="Hapus Mutasi" data-id="${m.id}">
                                    <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                                  </button>
                                ` : ''}
                              </div>
                            </td>
                          ` : ''}
                        </tr>
                      `;
                    }).join('')}
                  </tbody>
                </table>
              </div>
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
                        <span style="font-size:0.8rem; font-weight:600; color:#CBD5E1; display:inline-flex; align-items:center; gap:4px;">
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
                          <h4 style="margin:0 0 3px 0; font-size:1.02rem; font-weight:700; color:#F8FAFC; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                            ${m.employee_name}
                          </h4>
                          <div style="font-size:0.75rem; color:#94A3B8;">
                            ID Mutasi: #${m.id}
                          </div>
                        </div>
                      </div>

                      <!-- Route Visual Box -->
                      <div class="mutasi-route-box">
                        <div class="mutasi-route-node mutasi-from-node">
                          <div class="mutasi-route-label">Cabang Asal</div>
                          <div class="mutasi-route-branch" title="${m.from_branch_name || '-'}">
                            ${m.from_branch_name || '-'}
                          </div>
                        </div>

                        <div class="mutasi-route-arrow">
                          ➔
                        </div>

                        <div class="mutasi-route-node mutasi-to-node">
                          <div class="mutasi-route-label">Cabang Tujuan</div>
                          <div class="mutasi-route-branch" style="color:#93C5FD;" title="${m.to_branch_name || '-'}">
                            ${m.to_branch_name || '-'}
                          </div>
                        </div>
                      </div>
                    </div>

                    <!-- Card Footer: Dokumen & Aksi -->
                    <div style="display:flex; justify-content:space-between; align-items:center; padding-top:12px; border-top:1px solid #29364B; margin-top:8px;">
                      <div>
                        ${m.document_link ? `
                          <a href="${m.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.78rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                            <span>📄 Buka SK</span>
                          </a>
                        ` : `
                          <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                        `}
                      </div>

                      ${canEdit || canDelete ? `
                        <div style="display:flex; align-items:center; gap:6px;">
                          ${canEdit ? `
                            <button class="mutasi-icon-btn mutasi-edit-btn" title="Edit Data Mutasi" data-id="${m.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                            </button>
                          ` : ''}
                          ${canDelete ? `
                            <button class="mutasi-icon-btn delete mutasi-delete-btn" title="Hapus Mutasi" data-id="${m.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                            </button>
                          ` : ''}
                        </div>
                      ` : ''}
                    </div>

                  </div>
                `;
              }).join('')}
            </div>
          `}
    `;
  }

  function renderPagination() {
    const pagEl = container.querySelector('#mutasi-pagination-container');
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
    const btnReset = container.querySelector('#mutasi-btn-reset-filter');
    if (btnReset) {
      btnReset.onclick = () => {
        searchQuery = '';
        selectedFromBranch = 'all';
        selectedToBranch = 'all';
        selectedStatus = 'all';
        currentPage = 1;
        const searchInput = container.querySelector('#mutasi-search-input');
        if (searchInput) searchInput.value = '';
        const fromFilter = container.querySelector('#mutasi-filter-from');
        if (fromFilter) fromFilter.value = 'all';
        const toFilter = container.querySelector('#mutasi-filter-to');
        if (toFilter) toFilter.value = 'all';
        container.querySelectorAll('.mutasi-pill').forEach(p => {
          p.classList.toggle('active', p.getAttribute('data-status') === 'all');
        });
        loadData({ refreshStats: false });
      };
    }

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

  function renderListContent() {
    const listContainer = container.querySelector('#mutasi-list-container');
    if (!listContainer) return;
    const filtered = getFilteredList();
    listContainer.innerHTML = buildMutasiListHTML(filtered);
    renderPagination();
    bindListEvents();
  }
  // Main UI Renderer
  function renderUI() {
    const filtered = getFilteredList();

    // KPI Metrics from server stats
    const totalCount = serverStats.total;
    const prosesCount = serverStats.proses;
    const selesaiCount = serverStats.selesai;
    const bulanIniCount = serverStats.currentMonth;

    container.innerHTML = `
      <style>
        .mutasi-page {
          display: flex;
          flex-direction: column;
          gap: 20px;
        }
        .mutasi-hero-card {
          background: #131F36;
          border: 1px solid #223456;
          border-radius: 16px;
          padding: 24px 28px;
          color: #FFFFFF;
          box-shadow: 0 4px 16px rgba(0, 0, 0, 0.22);
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
          background: radial-gradient(circle, rgba(37, 99, 235, 0.06) 0%, transparent 70%);
          border-radius: 50%;
          pointer-events: none;
        }
        .mutasi-hero-actions {
          display: flex;
          align-items: center;
          gap: 8px;
          flex-wrap: wrap;
        }
        .mutasi-hero-actions .btn {
          height: 38px;
          box-sizing: border-box;
          transition: all 0.2s ease;
        }
        .mutasi-btn-add:hover {
          background: #1D4ED8 !important;
          transform: translateY(-1px);
        }
        .mutasi-btn-excel:hover, .mutasi-btn-refresh:hover {
          background: rgba(255, 255, 255, 0.16) !important;
          border-color: rgba(255, 255, 255, 0.3) !important;
          transform: translateY(-1px);
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
          background: #172033;
          border: 1px solid #29364B;
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
          align-items: center;
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
          border: 1px solid #29364B;
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
          border: 1px solid #29364B;
          border-radius: 10px;
          font-size: 0.9rem;
          color: #F8FAFC;
          background: #172033;
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
        .mutasi-pill:hover {
          background: #22314A;
          color: #F8FAFC;
        }
        .mutasi-pill.active {
          background: #2563EB;
          color: white;
          border-color: #2563EB;
          box-shadow: 0 2px 6px rgba(37, 99, 235, 0.25);
        }

        /* View Mode Toggle */
        .mutasi-view-toggle {
          display: inline-flex;
          align-items: center;
          background: #111827;
          padding: 3px;
          border-radius: 10px;
          border: 1px solid #29364B;
          gap: 3px;
          flex-shrink: 0;
        }
        .mutasi-view-btn {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          padding: 6px 12px;
          border-radius: 7px;
          font-size: 0.82rem;
          font-weight: 600;
          border: none;
          background: transparent;
          color: #CBD5E1;
          cursor: pointer;
          transition: all 0.15s ease;
          height: 34px;
          box-sizing: border-box;
          user-select: none;
        }
        .mutasi-view-btn:hover {
          color: #F8FAFC;
          background: #22314A;
        }
        .mutasi-view-btn.active {
          background: #172033;
          color: #60A5FA;
          font-weight: 700;
          box-shadow: 0 1px 4px rgba(0, 0, 0, 0.2);
        }

        /* Modern Table View */
        .mutasi-table-card {
          background: #172033;
          border: 1px solid #29364B;
          border-radius: 14px;
          box-shadow: 0 2px 6px rgba(0, 0, 0, 0.2);
          overflow: hidden;
        }
        .mutasi-table-responsive {
          width: 100%;
          overflow-x: auto;
          -webkit-overflow-scrolling: touch;
        }
        .mutasi-table {
          width: 100%;
          border-collapse: collapse;
          text-align: left;
          font-size: 0.88rem;
          min-width: 860px;
        }
        .mutasi-table th {
          background: #111827;
          color: #CBD5E1;
          font-size: 0.76rem;
          font-weight: 700;
          text-transform: uppercase;
          letter-spacing: 0.05em;
          padding: 13px 16px;
          border-bottom: 1.5px solid #29364B;
          white-space: nowrap;
        }
        .mutasi-table td {
          padding: 13px 16px;
          border-bottom: 1px solid #29364B;
          color: #F8FAFC;
          vertical-align: middle;
        }
        .mutasi-table tbody tr:hover td {
          background: #22314A;
        }
        .mutasi-table tbody tr:last-child td {
          border-bottom: none;
        }

        /* Grid Cards View */
        .mutasi-cards-grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
          gap: 16px;
        }
        .mutasi-card {
          background: #172033;
          border: 1px solid #29364B;
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
          box-shadow: 0 8px 18px rgba(0, 0, 0, 0.25);
          border-color: #3B82F6;
        }
        .mutasi-avatar-box {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          background: #2563EB;
          color: #FFFFFF;
          font-weight: 700;
          font-size: 1rem;
          display: flex;
          align-items: center;
          justify-content: center;
          flex-shrink: 0;
        }
        .mutasi-route-box {
          background: #0E182B;
          border: 1px dashed #334155;
          border-radius: 10px;
          padding: 12px 14px;
          margin: 12px 0;
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 10px;
          min-width: 0;
          box-sizing: border-box;
          width: 100%;
        }
        .mutasi-route-node {
          flex: 1;
          min-width: 0;
        }
        .mutasi-to-node {
          text-align: right;
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
          color: #F8FAFC;
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
          background: rgba(59, 130, 246, 0.18);
          color: #93C5FD;
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
          background: rgba(16, 185, 129, 0.18);
          color: #6EE7B7;
          border: 1px solid rgba(52, 211, 153, 0.35);
        }
        .mutasi-status-badge.proses {
          background: rgba(245, 158, 11, 0.18);
          color: #FCD34D;
          border: 1px solid rgba(251, 191, 36, 0.35);
        }
        .mutasi-icon-btn {
          width: 32px;
          height: 32px;
          border-radius: 8px;
          border: 1px solid #29364B;
          background: #1B2940;
          color: #CBD5E1;
          display: inline-flex;
          align-items: center;
          justify-content: center;
          cursor: pointer;
          transition: all 0.15s ease;
        }
        .mutasi-icon-btn:hover {
          background: #22314A;
          color: #F8FAFC;
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
          background: #172033;
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
          .mutasi-hero-actions {
            width: 100% !important;
            display: flex !important;
            align-items: center !important;
            gap: 6px !important;
            flex-wrap: nowrap !important;
            overflow-x: auto !important;
            scrollbar-width: none !important;
            padding-bottom: 2px !important;
          }
          .mutasi-hero-actions::-webkit-scrollbar {
            display: none;
          }
          .mutasi-btn-add {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.82rem !important;
            white-space: nowrap !important;
          }
          .mutasi-btn-excel {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.82rem !important;
            white-space: nowrap !important;
          }
          .mutasi-btn-refresh {
            width: 38px !important;
            height: 38px !important;
            padding: 0 !important;
            flex-shrink: 0 !important;
            justify-content: center !important;
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
          .mutasi-view-toggle {
            width: 100%;
            display: flex;
          }
          .mutasi-view-btn {
            flex: 1;
            justify-content: center;
          }
          .mutasi-table-card {
            border-radius: 12px;
          }
          .mutasi-cards-grid {
            grid-template-columns: minmax(0, 1fr) !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
          }
          .mutasi-card {
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
            overflow: hidden;
            padding: 14px !important;
          }
          .mutasi-route-box {
            flex-direction: column !important;
            align-items: stretch !important;
            gap: 8px !important;
            padding: 11px 12px !important;
            width: 100% !important;
            box-sizing: border-box !important;
          }
          .mutasi-route-node {
            width: 100% !important;
            min-width: 0 !important;
            text-align: left !important;
          }
          .mutasi-to-node {
            text-align: left !important;
          }
          .mutasi-route-branch {
            white-space: normal !important;
            word-break: break-word !important;
            overflow: visible !important;
            text-overflow: clip !important;
            font-size: 0.88rem !important;
            line-height: 1.35 !important;
          }
          .mutasi-route-arrow {
            align-self: center !important;
            width: 26px !important;
            height: 26px !important;
            font-size: 0.85rem !important;
            transform: rotate(90deg) !important;
            margin: -2px 0 !important;
          }
        }
      </style>

      <div class="mutasi-page">

        <!-- Hero Header Card -->
        <div class="mutasi-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px;">
                <span style="font-size:1.6rem;">🔁</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Monitoring Data Mutasi Karyawan
                </h1>
              </div>
            </div>

            <div class="mutasi-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${canAdd ? `
                <button id="mutasi-btn-add" class="btn btn-primary mutasi-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Mutasi</span>
                </button>
              ` : ''}
              
              ${hasPermission(currentUser, "export") ? `<button id="mutasi-btn-excel" class="btn btn-ghost mutasi-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>📥</span>
                <span>Ekspor Excel</span>
              </button>` : ''}

              <button id="mutasi-btn-refresh" class="btn btn-ghost mutasi-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:inline-flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
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
                <div class="mutasi-stat-val" id="mutasi-stat-total">${totalCount}</div>
                <div class="mutasi-stat-lbl">Total Mutasi</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                ⏳
              </div>
              <div>
                <div class="mutasi-stat-val" id="mutasi-stat-proses">${prosesCount}</div>
                <div class="mutasi-stat-lbl">Dalam Proses</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(34,197,94,0.2); color:#4ADE80;">
                ✅
              </div>
              <div>
                <div class="mutasi-stat-val" id="mutasi-stat-selesai">${selesaiCount}</div>
                <div class="mutasi-stat-lbl">Mutasi Selesai</div>
              </div>
            </div>

            <div class="mutasi-stat-box">
              <div class="mutasi-stat-icon" style="background:rgba(168,85,247,0.2); color:#C084FC;">
                🗓️
              </div>
              <div>
                <div class="mutasi-stat-val" id="mutasi-stat-month">${bulanIniCount}</div>
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

            <div class="mutasi-view-toggle">
              <button type="button" class="mutasi-view-btn ${viewMode === 'card' ? 'active' : ''}" data-view="card" title="Tampilan Kartu">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="mutasi-view-btn ${viewMode === 'table' ? 'active' : ''}" data-view="table" title="Tampilan Tabel">
                <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9h18M3 15h18M9 3v18"/><rect x="3" y="3" width="18" height="18" rx="2"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Status Filter Pills (Swipeable on Mobile) -->
          <div class="mutasi-filter-pills">
            <div class="mutasi-pill ${selectedStatus === 'all' ? 'active' : ''}" data-status="all">
              <span>Semua Mutasi</span>
              <span class="badge" id="mutasi-pill-count-all" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${totalCount}</span>
            </div>
            <div class="mutasi-pill ${selectedStatus === 'Proses' ? 'active' : ''}" data-status="Proses">
              <span style="color:#D97706;">⏳</span>
              <span>Proses</span>
              <span class="badge" id="mutasi-pill-count-proses" style="background:rgba(217,119,6,0.15); color:#B45309; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${prosesCount}</span>
            </div>
            <div class="mutasi-pill ${selectedStatus === 'Selesai' ? 'active' : ''}" data-status="Selesai">
              <span style="color:#16A34A;">✅</span>
              <span>Selesai</span>
              <span class="badge" id="mutasi-pill-count-selesai" style="background:rgba(22,163,74,0.15); color:#15803D; padding:2px 6px; border-radius:10px; font-size:0.75rem;">${selesaiCount}</span>
            </div>
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="mutasi-list-container">
          ${buildMutasiListHTML(filtered)}
        </div>
        <div id="mutasi-pagination-container"></div>

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
    if (btnRefresh) btnRefresh.onclick = () => loadData({ refreshStats: true });

    // View toggle
    container.querySelectorAll('.mutasi-view-btn').forEach(btn => {
      btn.onclick = () => {
        const mode = btn.getAttribute('data-view');
        if (mode && mode !== viewMode) {
          viewMode = mode;
          localStorage.setItem('fcms_mutasi_view_mode', mode);
          container.querySelectorAll('.mutasi-view-btn').forEach(b => {
            b.classList.toggle('active', b.getAttribute('data-view') === mode);
          });
          renderListContent();
        }
      };
    });

    // Debounced search input
    const searchInput = container.querySelector('#mutasi-search-input');
    if (searchInput) {
      searchInput.oninput = (e) => {
        searchQuery = e.target.value;
        clearTimeout(searchDebounce);
        searchDebounce = setTimeout(() => {
          currentPage = 1;
          loadData({ refreshStats: false });
        }, 350);
      };
    }

    // Branch filters
    const fromFilter = container.querySelector('#mutasi-filter-from');
    if (fromFilter) {
      fromFilter.onchange = (e) => {
        selectedFromBranch = e.target.value;
        currentPage = 1;
        loadData({ refreshStats: false });
      };
    }

    const toFilter = container.querySelector('#mutasi-filter-to');
    if (toFilter) {
      toFilter.onchange = (e) => {
        selectedToBranch = e.target.value;
        currentPage = 1;
        loadData({ refreshStats: false });
      };
    }

    // Status pills
    container.querySelectorAll('.mutasi-pill').forEach(pill => {
      pill.onclick = () => {
        selectedStatus = pill.getAttribute('data-status') || 'all';
        container.querySelectorAll('.mutasi-pill').forEach(p => {
          p.classList.toggle('active', p === pill);
        });
        currentPage = 1;
        loadData({ refreshStats: false });
      };
    });

    // Render pagination and bind list events
    renderPagination();
    bindListEvents();
  }

  // Initial load
  await loadData({ refreshStats: true });
}
