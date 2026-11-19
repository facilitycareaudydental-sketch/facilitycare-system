import { apiFetch, getUser, hasPermission } from '../config.js';
import { getCachedBranches, fetchExportRows, clearDataCache } from '../utils/dataCache.js';
import { createModal, confirmDialog } from '../components/modal.js';
import { toastSuccess, toastError, toastInfo } from '../components/toast.js';
import { createPagination } from '../components/table.js';
import { parseValidRowVersion, resolveModalRowVersion, handleOccSaveResponse } from '../utils/occ.js';

// Fallback kosong bila API offline (mencegah data cabang/PII tertanam di bundle klien)
const INITIAL_EMPLOYEES = [];

const DIVISION_LIST = ['FACILITY CARE', 'SECURITY'];
const STATUS_LIST = ['Aktif', 'Tidak Aktif', 'Resign', 'Cut'];

// Helper format tanggal lokal Indonesia (termasuk konversi serial date Excel seperti 45369)
function formatDateIndo(dateStr) {
  if (!dateStr) return '-';
  try {
    const s = String(dateStr).trim();
    if (/^\d{5}(\.\d+)?$/.test(s)) {
      const excelDate = new Date(Math.round((Number(s) - 25569) * 86400 * 1000));
      if (!isNaN(excelDate.getTime())) {
        return excelDate.toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' });
      }
    }
    const d = new Date(s);
    if (isNaN(d.getTime())) return s;
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
  const PAGE_LIMIT = 20;
  let employeeList = [];
  let branchOptions = [];
  let selectedStatus = 'all';
  let selectedDivision = 'all';
  let selectedBranch = 'all';
  let searchQuery = '';
  let currentPage = 1;
  let totalItems = 0;
  let totalPages = 1;
  let statsLoaded = false;
  let serverStats = { total: 0, active: 0, nonactive: 0, fc: 0, sec: 0 };
  let searchDebounce = null;
  let isLoading = false;
  let viewMode = localStorage.getItem('fcms_employees_view_mode') || 'card';

  // Handle dashboard filter parameter
  const dashFilter = params ? params.get('dash_filter') : null;
  if (dashFilter === 'active') {
    selectedStatus = 'Aktif';
  }

  const currentUser = getUser();
  const canAdd = hasPermission(currentUser, 'employees', 'write');
  const canEdit = hasPermission(currentUser, 'employees', 'write');
  const canDelete = hasPermission(currentUser, 'employees', 'delete');

  // Load branches (Phase 5.66.19: strictly filter selector to 000. Management)
  try {
    const allBranches = await getCachedBranches();
    branchOptions = (Array.isArray(allBranches) ? allBranches : []).filter(
      b => String(b.code || '').trim() === '000' || b.branch_scope === 'hq' || /^(?:cabang\s+)?000\.\s*management$/i.test(String(b.label || b.full_name || '').trim())
    );
  } catch (err) {
    console.error('Error fetching branches:', err);
  }

  function getActiveFilters() {
    const filters = { sort: 'name_asc' };
    if (searchQuery.trim()) filters.search = searchQuery.trim();
    if (selectedBranch && selectedBranch !== 'all') filters.branch_id = selectedBranch;
    if (selectedDivision && selectedDivision !== 'all') filters.division = selectedDivision;
    if (selectedStatus && selectedStatus !== 'all') filters.status = selectedStatus;
    return filters;
  }

  function updateStatsUI() {
    const elTotal = container.querySelector('#emp-stat-total');
    const elActive = container.querySelector('#emp-stat-active');
    const elFc = container.querySelector('#emp-stat-fc');
    const elSec = container.querySelector('#emp-stat-sec');
    const elSub = container.querySelector('#emp-stat-sub');
    const elPillAll = container.querySelector('#emp-pill-count-all');
    const elPillActive = container.querySelector('#emp-pill-count-active');
    const elPillNonactive = container.querySelector('#emp-pill-count-nonactive');

    if (elTotal) elTotal.textContent = serverStats.total;
    if (elActive) elActive.textContent = serverStats.active;
    if (elFc) elFc.textContent = serverStats.fc;
    if (elSec) elSec.textContent = serverStats.sec;
    if (elSub) elSub.textContent = `${serverStats.active} Aktif • ${serverStats.nonactive} Nonaktif`;
    if (elPillAll) elPillAll.textContent = serverStats.total;
    if (elPillActive) elPillActive.textContent = serverStats.active;
    if (elPillNonactive) elPillNonactive.textContent = serverStats.nonactive;
  }

  // Load Employees Data from API (Server-Side Pagination)
  async function loadData({ refreshStats = false } = {}) {
    isLoading = true;
    const hasShell = Boolean(container.querySelector('#emp-list-container'));
    if (!hasShell) {
      renderUI();
    } else {
      const listEl = container.querySelector('#emp-list-container');
      if (listEl) {
        listEl.innerHTML = '<div class="emp-empty-state"><p style="color:#64748B; margin:0;">Memuat data karyawan...</p></div>';
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

      const res = await apiFetch(`/api/employees?${qs.toString()}`);
      if (res.ok) {
        const raw = res.data?.data || [];
        employeeList = Array.isArray(raw) ? raw : [];
        const pag = res.data?.pagination || {};
        totalItems = Number(pag.total ?? employeeList.length);
        totalPages = Math.max(1, Number(pag.pages ?? Math.ceil(totalItems / PAGE_LIMIT)) || 1);
        currentPage = Math.max(1, Number(pag.page ?? currentPage));
        if (res.data?.stats) {
          serverStats = res.data.stats;
          statsLoaded = true;
        }
        if (employeeList.length === 0 && currentPage > totalPages) {
          currentPage = totalPages;
          isLoading = false;
          return loadData({ refreshStats });
        }
      } else {
        employeeList = [];
        totalItems = 0;
        totalPages = 1;
        toastError(res.data?.error || 'Gagal memuat data karyawan');
      }
    } catch (e) {
      console.warn('Gagal load employees:', e);
      employeeList = [];
      totalItems = 0;
      totalPages = 1;
      toastError('Gagal memuat data karyawan: ' + e.message);
    } finally {
      isLoading = false;
      if (container.querySelector('#emp-list-container')) {
        updateStatsUI();
        renderListContent();
      } else {
        renderUI();
      }
    }
  }

  // Export to Excel (All matching filtered rows via paginated /export endpoint)
  async function exportToExcel() {
    if (!hasPermission(currentUser, 'export')) {
      toastError('Anda tidak memiliki izin untuk mengekspor data.');
      return;
    }
    try {
      const [{ downloadExcel }, rows] = await Promise.all([
        import('../utils/excel.js'),
        fetchExportRows('/api/employees/export', getActiveFilters())
      ]);
      const exportData = rows.map((emp, idx) => ({
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
    const isEditCheck = !!(item && item.id);
    if (isEditCheck && !canEdit) {
      toastError('Anda tidak memiliki izin untuk mengedit data karyawan.');
      return;
    }
    if (!isEditCheck && !canAdd) {
      toastError('Anda tidak memiliki izin untuk menambah data karyawan.');
      return;
    }
    const isEdit = !!(item && item.id);
    const editRowVersion = isEdit ? parseValidRowVersion(item?.row_version) : null;
    const defaultDivision = item?.division || 'FACILITY CARE';
    const defaultStatus = item?.status || 'Aktif';

    const content = document.createElement('div');
    content.className = 'emp-modal-body';
    content.innerHTML = `
      <div style="display:flex; flex-direction:column; gap:14px;">
        ${isEdit ? `<input type="hidden" id="emp-modal-row-version" value="${editRowVersion !== null ? editRowVersion : ''}">` : ''}
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

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Catatan / Keterangan Karyawan
          </label>
          <textarea id="emp-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Catatan keahlian, tugas khusus, riwayat cover...">${item?.notes || ''}</textarea>
        </div>
      </div>
    `;

    let isSubmitting = false;

    createModal({
      title: isEdit ? '✏️ Edit Data Karyawan' : '➕ Tambah Data Karyawan',
      content,
      size: 'md',
      confirmText: isEdit ? 'Simpan Perubahan' : 'Simpan Karyawan',
      confirmClass: 'btn-primary',
      onConfirm: async (_, close) => {
        if (isSubmitting) return;

        const full_name = (content.querySelector('#emp-modal-name').value || '').trim();
        const phone = (content.querySelector('#emp-modal-phone').value || '').trim();
        const branch_id = content.querySelector('#emp-modal-branch').value;
        const division = content.querySelector('#emp-modal-division').value;
        const join_date = content.querySelector('#emp-modal-joindate').value || null;
        const status = content.querySelector('#emp-modal-status').value;
        const notes = (content.querySelector('#emp-modal-notes').value || '').trim();

        if (!full_name || !branch_id || !division || !status) {
          toastError('Harap lengkapi Nama Lengkap, Cabang, Divisi, dan Status!');
          return;
        }

        let rowVersion = null;
        if (isEdit) {
          rowVersion = resolveModalRowVersion(content, '#emp-modal-row-version', item?.row_version);
          if (rowVersion === null) {
            toastError('Versi data karyawan (row_version) tidak tersedia. Silakan muat ulang halaman lalu coba lagi.');
            return;
          }
        }

        const payload = {
          full_name,
          phone,
          branch_id: Number(branch_id),
          division,
          join_date,
          status,
          notes,
          ...(isEdit ? { row_version: rowVersion } : {})
        };

        isSubmitting = true;
        try {
          const res = isEdit
            ? await apiFetch(`/api/employees/${item.id}`, { method: 'PUT', body: JSON.stringify(payload) })
            : await apiFetch('/api/employees', { method: 'POST', body: JSON.stringify(payload) });

          await handleOccSaveResponse({
            res,
            isEdit,
            item,
            close,
            reloadData: async () => {
              clearDataCache();
              await loadData({ refreshStats: true });
            },
            editSuccessMessage: 'Data karyawan berhasil diperbarui!',
            createSuccessMessage: 'Karyawan baru berhasil ditambahkan!',
            defaultErrorMessage: 'Gagal menyimpan data karyawan',
          });
        } catch (err) {
          toastError('Terjadi kesalahan: ' + err.message);
        } finally {
          isSubmitting = false;
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
            clearDataCache();
            await loadData({ refreshStats: true });
          } else {
            toastError(res.data?.error || 'Gagal menonaktifkan data');
          }
        } catch (err) {
          toastError('Terjadi kesalahan: ' + err.message);
        }
      }
    });
  }

  // Filter Data (server-side filtered slice already loaded in employeeList)
  function getFilteredList() {
    return employeeList;
  }


  // Render only the dynamic list content HTML (Card / Table / Empty State)
  function buildEmployeeListHTML(filtered) {
    if (filtered.length === 0) {
      return `
            <div class="emp-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">🔍</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak ada karyawan yang cocok</h3>
              <p style="font-size:0.85rem; color:#64748B; margin:0 0 16px 0;">Coba sesuaikan kata kunci pencarian, filter cabang, atau filter divisi.</p>
              <button id="emp-reset-filter" class="btn btn-secondary" style="font-size:0.82rem; padding:7px 16px; border-radius:8px;">Reset Semua Filter</button>
            </div>
      `;
    }

    if (viewMode === 'table') {
      return `
            <div class="emp-table-card">
              <div class="emp-table-responsive">
                <table class="emp-table">
                  <thead>
                    <tr>
                      <th style="width:45px; text-align:center;">No</th>
                      <th>Karyawan</th>
                      <th>Cabang</th>
                      <th>Divisi</th>
                      <th>No. HP</th>
                      <th>Tgl Masuk</th>
                      <th>Catatan</th>
                      <th style="text-align:center;">Status</th>
                      <th style="text-align:center;">Aksi</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${filtered.map((item, idx) => {
                      const div = item.division || 'FACILITY CARE';
                      let divClass = 'emp-div-fc';
                      let avatarClass = '';
                      if (div === 'SECURITY') {
                        divClass = 'emp-div-sec';
                        avatarClass = 'sec';
                      }
                      const isAktif = item.status === 'Aktif';
                      const phone = item.phone || '';

                      return `
                        <tr data-id="${item.id}">
                          <td style="text-align:center; color:#94A3B8; font-weight:600;">${((currentPage - 1) * PAGE_LIMIT) + idx + 1}</td>
                          <td>
                            <div style="display:flex; align-items:center; gap:10px;">
                              <div class="emp-avatar ${avatarClass}" style="width:34px; height:34px; font-size:0.75rem;">
                                ${getInitials(item.full_name)}
                              </div>
                              <div style="min-width:0;">
                                <div style="font-weight:700; color:#1E293B; font-size:0.9rem;">${item.full_name || 'Karyawan'}</div>
                                <div style="font-size:0.74rem; color:#64748B;">ID #${item.id} ${item.nik ? '• ' + item.nik : ''}</div>
                              </div>
                            </div>
                          </td>
                          <td>
                            <span style="font-weight:600; color:#334155;">${item.branch_name || 'Semua Cabang'}</span>
                          </td>
                          <td>
                            <span class="emp-div-badge ${divClass}">
                              ${div === 'SECURITY' ? '🛡️' : '🧹'} ${div}
                            </span>
                          </td>
                          <td>
                            ${phone ? `
                              <a href="https://wa.me/${phone.replace(/[^0-9]/g, '').replace(/^0/, '62')}" target="_blank" rel="noopener" style="color:#16A34A; text-decoration:none; font-weight:600; display:inline-flex; align-items:center; gap:4px;">
                                <span>📱</span> ${phone}
                              </a>
                            ` : '<span style="color:#94A3B8;">-</span>'}
                          </td>
                          <td style="font-size:0.84rem; color:#475569;">
                            ${item.join_date ? formatDateIndo(item.join_date) : '-'}
                          </td>
                          <td style="font-size:0.84rem; color:#475569;">
                            ${item.notes || '-'}
                          </td>
                          <td style="text-align:center;">
                            <span class="emp-status-badge ${isAktif ? 'emp-status-aktif' : 'emp-status-nonaktif'}">
                              ${isAktif ? '🟢 Aktif' : '⚪ Nonaktif'}
                            </span>
                          </td>
                          <td style="text-align:center;">
                            <div class="emp-action-btns" style="justify-content:center;">
                              ${canEdit ? `
                                <button class="emp-icon-btn edit" data-action="edit" data-id="${item.id}" title="Edit Data Karyawan">
                                  <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                                </button>
                              ` : ''}
                              ${canDelete ? `
                                <button class="emp-icon-btn delete" data-action="delete" data-id="${item.id}" title="Nonaktifkan Karyawan">
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
          <div class="emp-cards-grid">
            ${filtered.map(item => {
              const div = item.division || 'FACILITY CARE';
              let divClass = 'emp-div-fc';
              let avatarClass = '';
              if (div === 'SECURITY') {
                divClass = 'emp-div-sec';
                avatarClass = 'sec';
              }
              const isAktif = item.status === 'Aktif';
              const phone = item.phone || '';

              return `
                <div class="emp-card" data-id="${item.id}">
                  <div class="emp-card-top">
                    <span class="emp-div-badge ${divClass}">
                      ${div === 'SECURITY' ? '🛡️' : '🧹'} ${div}
                    </span>
                    <span class="emp-status-badge ${isAktif ? 'emp-status-aktif' : 'emp-status-nonaktif'}">
                      ${isAktif ? '🟢 Aktif' : '⚪ Nonaktif'}
                    </span>
                  </div>

                  <div class="emp-profile-box">
                    <div class="emp-avatar ${avatarClass}">
                      ${getInitials(item.full_name)}
                    </div>
                    <div style="min-width:0; flex:1;">
                      <div class="emp-name">${item.full_name || 'Karyawan'}</div>
                      <div class="emp-subtitle">Personil ${div} • ID #${item.id}</div>
                    </div>
                  </div>

                  <div class="emp-info-grid">
                    <div class="emp-info-card">
                      <div class="emp-info-label">Cabang Penempatan</div>
                      <div class="emp-info-val" title="${item.branch_name || '-'}">
                        🏢 ${item.branch_name || 'Semua Cabang'}
                      </div>
                    </div>
                    <div class="emp-info-card">
                      <div class="emp-info-label">Tanggal Masuk</div>
                      <div class="emp-info-val">
                        📅 ${item.join_date ? formatDateIndo(item.join_date) : '-'}
                      </div>
                    </div>
                    <div class="emp-info-card">
                      <div class="emp-info-label">Nomor Induk Karyawan</div>
                      <div class="emp-info-val">
                        🆔 ${item.nik || '-'}
                      </div>
                    </div>
                    <div class="emp-info-card">
                      <div class="emp-info-label">Status Personil</div>
                      <div class="emp-info-val" style="color:${isAktif ? '#16A34A' : '#DC2626'};">
                        ${isAktif ? '✅ Karyawan Aktif' : '⛔ Nonaktif / Resign'}
                      </div>
                    </div>
                  </div>

                  ${item.notes ? `
                    <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:10px 12px; font-size:0.82rem; color:#475569;">
                      <strong>Catatan:</strong> ${item.notes}
                    </div>
                  ` : ''}

                  <div class="emp-card-footer">
                    <div style="font-size:0.75rem; color:#94A3B8;">
                      ID #${item.id} ${item.nik ? '• ' + item.nik : ''}
                    </div>

                    <div class="emp-action-btns">
                      ${phone ? `
                        <a href="https://wa.me/${phone.replace(/[^0-9]/g, '').replace(/^0/, '62')}" target="_blank" rel="noopener" class="emp-icon-btn" title="Chat WhatsApp Karyawan" style="color:#16A34A;">
                          <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                        </a>
                      ` : ''}

                      ${canEdit ? `
                        <button class="emp-icon-btn edit" data-action="edit" data-id="${item.id}" title="Edit Data Karyawan">
                          <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                        </button>
                      ` : ''}
                      ${canDelete ? `
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
    `;
  }

  function renderPagination() {
    const pagEl = container.querySelector('#emp-pagination-container');
    if (!pagEl) return;
    pagEl.innerHTML = '';
    const nav = createPagination({
      page: currentPage,
      pages: totalPages,
      total: totalItems,
      limit: PAGE_LIMIT,
      onPage: (targetPage) => {
        currentPage = targetPage;
        loadData();
      }
    });
    if (nav) pagEl.appendChild(nav);
  }

  function bindListEvents() {
    const resetBtn = container.querySelector('#emp-reset-filter');
    if (resetBtn) {
      resetBtn.addEventListener('click', () => {
        searchQuery = '';
        selectedBranch = 'all';
        selectedDivision = 'all';
        selectedStatus = 'all';
        currentPage = 1;
        const searchInput = container.querySelector('#emp-search-input');
        if (searchInput) searchInput.value = '';
        const branchFilter = container.querySelector('#emp-branch-filter');
        if (branchFilter) branchFilter.value = 'all';
        const divFilter = container.querySelector('#emp-division-filter');
        if (divFilter) divFilter.value = 'all';
        container.querySelectorAll('.emp-pill').forEach(p => {
          p.classList.toggle('active', p.getAttribute('data-status') === 'all');
        });
        loadData();
      });
    }

    container.querySelectorAll('.emp-icon-btn.edit').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const id = Number(e.currentTarget.getAttribute('data-id'));
        const item = employeeList.find(x => x.id === id);
        if (item) openEmployeeModal(item);
      });
    });

    container.querySelectorAll('.emp-icon-btn.delete').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const id = Number(e.currentTarget.getAttribute('data-id'));
        const item = employeeList.find(x => x.id === id);
        if (item) deleteEmployee(item);
      });
    });
  }

  function renderListContent() {
    const listContainer = container.querySelector('#emp-list-container');
    if (!listContainer) return;
    const filtered = getFilteredList();
    listContainer.innerHTML = buildEmployeeListHTML(filtered);
    renderPagination();
    bindListEvents();
  }
  // Main UI Renderer
  function renderUI() {
    const filtered = getFilteredList();

    // KPI Metrics from server stats
    const totalEmployees = serverStats.total;
    const activeCount = serverStats.active;
    const fcCount = serverStats.fc;
    const secCount = serverStats.sec;

    // Counts for status pills
    const countAll = serverStats.total;
    const countActive = serverStats.active;
    const countNonactive = serverStats.nonactive;

    container.innerHTML = `
      <style>
        .emp-page {
          display: flex;
          flex-direction: column;
          gap: 20px;
          width: 100%;
          max-width: 100%;
          min-width: 0;
          box-sizing: border-box;
          overflow-x: hidden;
        }
        .emp-hero-card {
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
          box-sizing: border-box;
          width: 100%;
          max-width: 100%;
          min-width: 0;
        }
        .emp-search-row {
          display: flex;
          gap: 12px;
          flex-wrap: wrap;
          align-items: center;
        }
        .emp-view-toggle {
          display: inline-flex;
          background: #F1F5F9;
          padding: 4px;
          border-radius: 10px;
          border: 1px solid #E2E8F0;
          gap: 4px;
          flex-shrink: 0;
        }
        .emp-view-btn {
          display: inline-flex;
          align-items: center;
          gap: 6px;
          padding: 6px 12px;
          border-radius: 7px;
          border: none;
          background: transparent;
          color: #64748B;
          font-size: 0.82rem;
          font-weight: 700;
          cursor: pointer;
          transition: all 0.15s ease;
        }
        .emp-view-btn:hover {
          color: #0F172A;
        }
        .emp-view-btn.active {
          background: white;
          color: #2563EB;
          box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        }
        .emp-table-card {
          background: white;
          border-radius: 14px;
          border: 1px solid #E2E8F0;
          box-shadow: 0 2px 8px rgba(0,0,0,0.04);
          overflow: hidden;
          width: 100%;
          box-sizing: border-box;
        }
        .emp-table-responsive {
          width: 100%;
          overflow-x: auto;
        }
        .emp-table {
          width: 100%;
          border-collapse: collapse;
          text-align: left;
          font-size: 0.85rem;
        }
        .emp-table th {
          background: #F8FAFC;
          color: #475569;
          font-weight: 700;
          font-size: 0.75rem;
          text-transform: uppercase;
          letter-spacing: 0.04em;
          padding: 13px 16px;
          border-bottom: 1px solid #E2E8F0;
          white-space: nowrap;
        }
        .emp-table td {
          padding: 13px 16px;
          border-bottom: 1px solid #F1F5F9;
          color: #1E293B;
          vertical-align: middle;
        }
        .emp-table tbody tr {
          transition: background 0.15s ease;
        }
        .emp-table tbody tr:hover {
          background: #F8FAFC;
        }
        .emp-table tbody tr:last-child td {
          border-bottom: none;
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
          grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
          gap: 16px;
          width: 100%;
          max-width: 100%;
          min-width: 0;
          box-sizing: border-box;
        }
        .emp-card {
          background: white;
          border-radius: 14px;
          border: 1px solid #E2E8F0;
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
          width: 100%;
          min-width: 0;
          box-sizing: border-box;
          flex-wrap: wrap;
        }
        
        .emp-div-badge {
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

        .emp-status-badge {
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
          gap: 12px;
          width: 100%;
          min-width: 0;
          box-sizing: border-box;
        }
        .emp-avatar {
          width: 50px;
          height: 50px;
          border-radius: 12px;
          background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
          color: white;
          display: flex;
          align-items: center;
          justify-content: center;
          font-weight: 800;
          font-size: 1.1rem;
          flex-shrink: 0;
          box-shadow: 0 4px 10px rgba(37, 99, 235, 0.2);
        }
        .emp-avatar.sec {
          background: linear-gradient(135deg, #334155 0%, #0F172A 100%);
          box-shadow: 0 4px 10px rgba(15, 23, 42, 0.2);
        }
        .emp-name {
          font-size: 1.02rem;
          font-weight: 800;
          color: #0F172A;
          line-height: 1.25;
          word-break: break-word;
          overflow-wrap: anywhere;
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
        .emp-info-row {
          display: flex;
          align-items: flex-start;
          justify-content: space-between;
          gap: 8px;
          width: 100%;
          min-width: 0;
          box-sizing: border-box;
        }
        .emp-info-label {
          color: #64748B;
          display: inline-flex;
          align-items: center;
          gap: 5px;
          flex-shrink: 0;
          white-space: nowrap;
          font-size: 0.8rem;
        }
        .emp-info-val {
          color: #0F172A;
          font-weight: 600;
          text-align: right;
          min-width: 0;
          flex: 1;
          word-break: break-word;
          overflow-wrap: anywhere;
          font-size: 0.82rem;
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
        .emp-card-footer {
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
        .emp-action-btns {
          display: flex;
          align-items: center;
          gap: 6px;
          margin-left: auto;
          flex-shrink: 0;
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
          .emp-page {
            gap: 14px;
            width: 100% !important;
            max-width: 100% !important;
            overflow-x: hidden !important;
          }
          .emp-hero-card {
            padding: 16px 14px !important;
            width: 100% !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
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
            padding: 14px !important;
            gap: 12px;
            width: 100% !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
          }
          .emp-search-row {
            flex-direction: column;
            gap: 10px;
          }
          .emp-select {
            width: 100%;
          }
          .emp-view-toggle {
            width: 100%;
            display: flex;
          }
          .emp-view-btn {
            flex: 1;
            justify-content: center;
          }
          .emp-cards-grid {
            grid-template-columns: minmax(0, 1fr) !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
          }
          .emp-card {
            padding: 14px !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
          }
          .emp-info-card {
            padding: 10px !important;
            width: 100% !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
          }
        }
      </style>

      <div class="emp-page">

        <!-- Hero Header Card -->
        <div class="emp-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px;">
                <span style="font-size:1.6rem;">👥</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Master Karyawan FCMS
                </h1>
              </div>
            </div>

            <div class="emp-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${canAdd ? `
                <button id="emp-btn-add" class="btn btn-primary emp-btn-add" style="background:#2563EB; padding:9px 16px; border-radius:10px; font-weight:700; display:flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Karyawan</span>
                </button>
              ` : ''}
              
              ${hasPermission(currentUser, "export") ? `<button id="emp-btn-excel" class="btn btn-ghost emp-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; padding:9px 14px; border-radius:10px; font-weight:600; display:flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>📥</span>
                <span>Ekspor Excel</span>
              </button>` : ''}

              <button id="emp-btn-refresh" class="btn btn-ghost emp-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24" class="${isLoading ? 'spin-icon' : ''}"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- 4 KPI Stat Boxes -->
          <div class="emp-stats-row">
            <div class="emp-stat-box" title="Data karyawan terdaftar di sistem">
              <div class="emp-stat-icon" style="background:rgba(37, 99, 235, 0.2); color:#93C5FD;">👥</div>
              <div>
                <div class="emp-stat-val" id="emp-stat-total">${totalEmployees}</div>
                <div class="emp-stat-lbl">Karyawan Terdaftar</div>
                <div id="emp-stat-sub" style="font-size:0.68rem; color:#94A3B8; margin-top:2px;">${activeCount} Aktif • ${countNonactive} Nonaktif</div>
              </div>
            </div>

            <div class="emp-stat-box" title="Total Personil Aktif">
              <div class="emp-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#6EE7B7;">🟢</div>
              <div>
                <div class="emp-stat-val" id="emp-stat-active">${activeCount}</div>
                <div class="emp-stat-lbl">Personil Aktif</div>
                <div style="font-size:0.68rem; color:#86EFAC; margin-top:2px;">Status Aktif Operasional</div>
              </div>
            </div>

            <div class="emp-stat-box" title="Divisi Resmi Facility Care">
              <div class="emp-stat-icon" style="background:rgba(14, 165, 233, 0.2); color:#7DD3FC;">🏢</div>
              <div>
                <div class="emp-stat-val" id="emp-stat-fc">${fcCount}</div>
                <div class="emp-stat-lbl">Facility Care</div>
                <div style="font-size:0.68rem; color:#BAE6FD; margin-top:2px;">Aktif Operasional</div>
              </div>
            </div>

            <div class="emp-stat-box" title="Divisi Resmi Security">
              <div class="emp-stat-icon" style="background:rgba(124, 58, 237, 0.2); color:#C4B5FD;">🛡️</div>
              <div>
                <div class="emp-stat-val" id="emp-stat-sec">${secCount}</div>
                <div class="emp-stat-lbl">Security</div>
                <div style="font-size:0.68rem; color:#DDD6FE; margin-top:2px;">Aktif Operasional</div>
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

            <div class="emp-view-toggle" title="Pilih Tampilan">
              <button type="button" class="emp-view-btn ${viewMode === 'card' ? 'active' : ''}" data-view="card">
                <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/></svg>
                <span>Kartu</span>
              </button>
              <button type="button" class="emp-view-btn ${viewMode === 'table' ? 'active' : ''}" data-view="table">
                <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>
                <span>Tabel</span>
              </button>
            </div>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="emp-filter-pills">
            <div class="emp-pill ${selectedStatus === 'all' && selectedDivision === 'all' ? 'active' : ''}" data-status="all" data-div="all">
              <span>Semua Karyawan</span>
              <span class="emp-pill-count" id="emp-pill-count-all">${countAll}</span>
            </div>
            <div class="emp-pill ${selectedStatus === 'Aktif' ? 'active' : ''}" data-status="Aktif" data-div="all">
              <span class="emp-pill-dot" style="background:#10B981;"></span>
              <span>Aktif</span>
              <span class="emp-pill-count" id="emp-pill-count-active">${countActive}</span>
            </div>
            <div class="emp-pill ${selectedStatus === 'Nonaktif' ? 'active' : ''}" data-status="Nonaktif" data-div="all">
              <span class="emp-pill-dot" style="background:#EF4444;"></span>
              <span>Nonaktif / Resign</span>
              <span class="emp-pill-count" id="emp-pill-count-nonactive">${countNonactive}</span>
            </div>
          </div>
        </div>

        <!-- Exclusive List Container (Card / Table) -->
        <div id="emp-list-container">
          ${buildEmployeeListHTML(filtered)}
        </div>
        <div id="emp-pagination-container"></div>

      </div>
    `;

    // ── Event Handlers ──
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
      refreshBtn.addEventListener('click', () => loadData({ refreshStats: true }));
    }

    container.querySelectorAll('.emp-view-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const mode = btn.getAttribute('data-view');
        if (mode && mode !== viewMode) {
          viewMode = mode;
          localStorage.setItem('fcms_employees_view_mode', mode);
          container.querySelectorAll('.emp-view-btn').forEach(b => {
            b.classList.toggle('active', b.getAttribute('data-view') === mode);
          });
          renderListContent();
        }
      });
    });

    const searchInput = container.querySelector('#emp-search-input');
    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        searchQuery = e.target.value;
        clearTimeout(searchDebounce);
        searchDebounce = setTimeout(() => {
          currentPage = 1;
          loadData();
        }, 350);
      });
    }

    const branchFilter = container.querySelector('#emp-branch-filter');
    if (branchFilter) {
      branchFilter.addEventListener('change', (e) => {
        selectedBranch = e.target.value;
        currentPage = 1;
        loadData();
      });
    }

    const divisionFilter = container.querySelector('#emp-division-filter');
    if (divisionFilter) {
      divisionFilter.addEventListener('change', (e) => {
        selectedDivision = e.target.value;
        currentPage = 1;
        loadData();
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
        const divFilter = container.querySelector('#emp-division-filter');
        if (divFilter) divFilter.value = selectedDivision;
        container.querySelectorAll('.emp-pill').forEach(p => {
          const isActive = p === pill || (selectedStatus === 'all' && selectedDivision === 'all' && p.dataset.status === 'all');
          p.classList.toggle('active', isActive);
        });
        currentPage = 1;
        loadData();
      });
    });

    // Bind pagination & list action events (Reset button, Edit, Delete)
    renderPagination();
    bindListEvents();
  }

  // Initial load
  await loadData({ refreshStats: true });
}
