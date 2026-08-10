// Generic CRUD page builder - used by all modules
import { apiFetch, CLIENT_SIDE_MAX_ROWS, IS_DEVELOPMENT } from '../config.js';
import { createTable, createPagination } from '../components/table.js';
import { createModal, confirmDialog } from '../components/modal.js';
import { buildFormHTML, getFormData, populateForm } from '../components/form.js';
import { toastSuccess, toastError } from '../components/toast.js';
import { renderExcelButtons, parseExcel } from '../utils/excel.js';
import { notifyCalendar } from '../utils/calendarBus.js';

export function buildCrudPage({
  container,
  title,
  icon,
  apiPath,
  columns,
  formFields,
  filterFields,
  defaultFilters = {},
  itemLabel = 'Data',
  canCreate = true,
  canEdit = true,
  canDelete = true,
  onBeforeSubmit,
  onAfterLoad,
  onDataLoaded,
  extraActions = [],
  initialSearch = '',
  exportOptions = null, // { moduleName, onExport, onImport, onTemplate }
  bulkDelete = false,   // true => enable checkbox bulk-delete using DELETE apiPath/bulk
  paginationMode = 'server', // 'server' or 'client'
}) {
  let page = 1;
  let filters = { ...defaultFilters };
  if (initialSearch) filters.search = initialSearch;
  let selectedIds = new Set();

  container.innerHTML = `
    ${bulkDelete ? `
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:none; align-items:center; justify-content:space-between; background:#2563EB; padding:12px 16px; border-radius:8px; margin-bottom:16px; box-shadow:0 4px 6px -1px rgba(37,99,235,0.2);">
      <button id="btn-bulk-cancel" style="background:transparent; border:none; color:white; display:flex; align-items:center; cursor:pointer; padding:4px;">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 6L6 18M6 6l12 12"></path></svg>
      </button>
      <span id="bulk-count" style="font-weight:500; font-size:0.95rem; color:white;">0 item dipilih</span>
      <button id="btn-bulk-delete" style="background:transparent; border:none; color:white; display:flex; align-items:center; cursor:pointer; padding:4px;">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>
      </button>
    </div>` : ''}

    <div class="page-header">
      <h1 class="page-title">${icon} ${title}</h1>
      <div class="page-actions" style="display:flex; gap:8px; align-items:center;">
        ${canCreate ? `<button class="btn btn-primary" id="btn-create">+ Tambah ${itemLabel}</button>` : ''}
        ${exportOptions ? `
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#fff; border:1px solid #E2E8F0; padding:8px 16px; border-radius:8px; font-weight:600; color:#334155; display:flex; align-items:center; gap:6px; cursor:pointer;" onclick="document.getElementById('aksi-menu-main').classList.toggle('show-aksi-menu')">
              ⋮ Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#fff; border:1px solid #E2E8F0; border-radius:8px; box-shadow:0 4px 12px rgba(0,0,0,0.1); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              
              <button class="dropdown-item" id="btn-export-${exportOptions.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                📥 Export Excel
              </button>
              <button class="dropdown-item" id="btn-template-${exportOptions.moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; display:flex; gap:8px; align-items:center;">
                📄 Download Template
              </button>
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#334155; margin:0; gap:8px; align-items:center;" id="label-import-${exportOptions.moduleName}">
                📤 Import Excel
                <input type="file" id="input-import-${exportOptions.moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
              </label>

            </div>
          </div>
          <style>
            .show-aksi-menu { display: flex !important; }
            .dropdown-item:hover { background-color: #F8FAFC !important; color: #2563EB !important; }
          </style>
        ` : ''}
      </div>
    </div>
    

    ${filterFields && filterFields.length > 0 ? `
    <div class="filter-bar" style="background: var(--bg-card, #fff); border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid var(--border, #E2E8F0); box-shadow: 0 1px 4px rgba(0,0,0,0.06);">
        ${filterFields.filter(f => f.type === 'search').map(f => {
          return `<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${f.placeholder || 'Cari...'}" id="filter-search" value="${filters.search || ''}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`;
        }).join('')}
        
        <div class="filter-dropdowns-desktop">
          ${filterFields.filter(f => f.type !== 'search').map(f => {
            if (f.type === 'select' || f.type === 'combobox') {
              const placeholder = (f.label || '').startsWith('Pilih') ? f.label : `Pilih ${f.label || ''}`;
              return `<select class="filter-select" name="${f.name}" id="filter-${f.name}" style="background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #475569; cursor: pointer; outline:none;"><option value="">${placeholder}</option>${(f.options || []).map(o => `<option value="${typeof o === 'object' ? o.value : o}" ${filters[f.name] === (typeof o === 'object' ? o.value : o) ? 'selected' : ''}>${typeof o === 'object' ? o.label : o}</option>`).join('')}</select>`;
            }
            return '';
          }).join('')}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">⚙ Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${filterFields.filter(f => f.type !== 'search').map(f => {
            if (f.type === 'select' || f.type === 'combobox') {
              const placeholder = (f.label || '').startsWith('Pilih') ? f.label : `Pilih ${f.label || ''}`;
              return `<select class="filter-select filter-select-sheet" name="${f.name}-sheet" id="filter-sheet-${f.name}" style="width:100%; background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #1e293b; cursor: pointer; outline:none;"><option value="">${placeholder}</option>${(f.options || []).map(o => `<option value="${typeof o === 'object' ? o.value : o}" ${filters[f.name] === (typeof o === 'object' ? o.value : o) ? 'selected' : ''}>${typeof o === 'object' ? o.label : o}</option>`).join('')}</select>`;
            }
            return '';
          }).join('')}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #3B82F6; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>` : ''}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;

  // Bulk toolbar logic
  function updateBulkToolbar() {
    const toolbar = document.getElementById('bulk-toolbar');
    if (!toolbar) return;
    const countEl = document.getElementById('bulk-count');
    const btnDelete = document.getElementById('btn-bulk-delete');
    const btnCancel = document.getElementById('btn-bulk-cancel');
    
    countEl.textContent = `${selectedIds.size} item dipilih`;
    if (selectedIds.size > 0) {
      toolbar.style.display = 'flex';
      btnDelete.disabled = false;
      btnCancel.disabled = false;
    } else {
      toolbar.style.display = 'none';
      btnDelete.disabled = true;
      btnCancel.disabled = true;
    }
    
    // Update Select All Checkbox state berdasarkan cb.checked (bukan selectedIds.has)
    // karena cb.value selalu string tapi selectedIds menyimpan number → type mismatch
    const selectAll = document.getElementById('select-all-checkbox');
    if (selectAll) {
      const rows = document.querySelectorAll('.row-checkbox');
      if (rows.length > 0) {
        const allChecked = [...rows].every(cb => cb.checked);
        const someChecked = [...rows].some(cb => cb.checked);
        selectAll.checked = allChecked;
        selectAll.indeterminate = someChecked && !allChecked;
      } else {
        selectAll.checked = false;
        selectAll.indeterminate = false;
      }
    }
  }

  document.getElementById('btn-bulk-cancel')?.addEventListener('click', () => {
    selectedIds.clear();
    // uncheck all
    document.querySelectorAll('.row-checkbox').forEach(cb => cb.checked = false);
    const selectAll = document.getElementById('select-all-checkbox');
    if (selectAll) selectAll.checked = false;
    updateBulkToolbar();
  });

  document.getElementById('btn-bulk-delete')?.addEventListener('click', () => {
    if (selectedIds.size === 0) return;
    const ids = [...selectedIds];
    const overlay = document.createElement('div');
    overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center';
    overlay.innerHTML = `
      <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:28px;max-width:420px;width:90%;box-shadow:var(--shadow-lg);animation:fadeInUp .2s ease">
        <h3 style="margin:0 0 8px;color:var(--text-1);font-size:1rem;font-weight:700">⚠️ Hapus ${ids.length} ${itemLabel}?</h3>
        <p style="margin:0 0 24px;color:var(--text-2);font-size:.875rem">Data yang dihapus tidak dapat dikembalikan.</p>
        <div style="display:flex;gap:8px;justify-content:flex-end">
          <button id="bulk-cancel-btn" class="btn btn-ghost">Batal</button>
          <button id="bulk-confirm-btn" class="btn btn-danger">Hapus ${ids.length} Data</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);
    overlay.querySelector('#bulk-cancel-btn').addEventListener('click', () => overlay.remove());
    overlay.querySelector('#bulk-confirm-btn').addEventListener('click', async () => {
      const confirmBtn = overlay.querySelector('#bulk-confirm-btn');
      confirmBtn.disabled = true;
      confirmBtn.textContent = 'Menghapus...';
      const res = await apiFetch(`${apiPath}/bulk`, { method: 'DELETE', body: JSON.stringify({ ids }) });
      overlay.remove();
      if (res.ok) {
        toastSuccess(`${ids.length} ${itemLabel} berhasil dihapus.`);
        selectedIds.clear();
        updateBulkToolbar();
        notifyCalendar(apiPath);
        load();
      } else {
        toastError(res.data?.error || 'Gagal menghapus data.');
      }
    });
  });

  // Filter events
  const searchInput = document.getElementById('filter-search');
  let searchTimer;
  searchInput?.addEventListener('input', (e) => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      filters.search = e.target.value;
      page = 1;
      selectedIds.clear();
      updateBulkToolbar();
      load();
    }, 400);
  });

  // Desktop dropdown filters
  filterFields?.forEach(f => {
    if (f.type === 'select' || f.type === 'combobox') {
      document.getElementById(`filter-${f.name}`)?.addEventListener('change', (e) => {
        filters[f.name] = e.target.value;
        // also sync sheet select
        const sheetEl = document.getElementById(`filter-sheet-${f.name}`);
        if (sheetEl) sheetEl.value = e.target.value;
        page = 1;
        selectedIds.clear();
        updateBulkToolbar();
        load();
      });
      // Mobile sheet selects — sync to filters and close sheet
      document.getElementById(`filter-sheet-${f.name}`)?.addEventListener('change', (e) => {
        filters[f.name] = e.target.value;
        const desktopEl = document.getElementById(`filter-${f.name}`);
        if (desktopEl) desktopEl.value = e.target.value;
        page = 1;
        selectedIds.clear();
        updateBulkToolbar();
        load();
        // Close sheet after selecting
        document.getElementById('filter-options-wrapper')?.classList.remove('sheet-open');
      });
    }
  });

  // Desktop reset
  document.getElementById('btn-reset-filter')?.addEventListener('click', () => {
    filters = {};
    if (searchInput) searchInput.value = '';
    filterFields?.forEach(f => {
      const el = document.getElementById(`filter-${f.name}`);
      if (el) el.value = '';
      const sheetEl = document.getElementById(`filter-sheet-${f.name}`);
      if (sheetEl) sheetEl.value = '';
    });
    page = 1;
    selectedIds.clear();
    updateBulkToolbar();
    load();
  });

  // Mobile sheet reset
  document.getElementById('btn-reset-filter-sheet')?.addEventListener('click', () => {
    filters = {};
    if (searchInput) searchInput.value = '';
    filterFields?.forEach(f => {
      const el = document.getElementById(`filter-${f.name}`);
      if (el) el.value = '';
      const sheetEl = document.getElementById(`filter-sheet-${f.name}`);
      if (sheetEl) sheetEl.value = '';
    });
    page = 1;
    selectedIds.clear();
    updateBulkToolbar();
    load();
    document.getElementById('filter-options-wrapper')?.classList.remove('sheet-open');
  });

  // Create button
  document.getElementById('btn-create')?.addEventListener('click', () => openForm(null));

  // Aksi dropdown outside click
  if (exportOptions) {
    document.addEventListener('click', function(e) {
      const dropdown = document.getElementById('aksi-menu-main');
      const btn = document.getElementById('btn-aksi-main');
      if (dropdown && btn && !btn.contains(e.target) && !dropdown.contains(e.target)) {
        dropdown.classList.remove('show-aksi-menu');
      }
    });
  }
  
  // Mobile Filter Sheet Logic
  const btnMobileFilter = document.getElementById('btn-mobile-filter');
  const filterOptionsWrapper = document.getElementById('filter-options-wrapper');
  const btnCloseFilterSheet = document.getElementById('btn-close-filter-sheet');
  
  if (btnMobileFilter && filterOptionsWrapper) {
    btnMobileFilter.addEventListener('click', (e) => {
      e.preventDefault();
      filterOptionsWrapper.classList.add('sheet-open');
    });
    if (btnCloseFilterSheet) {
      btnCloseFilterSheet.addEventListener('click', (e) => {
        e.preventDefault();
        filterOptionsWrapper.classList.remove('sheet-open');
      });
    }
  }

  // Export/Import buttons
  if (exportOptions) {
    document.getElementById(`btn-export-${exportOptions.moduleName}`)?.addEventListener('click', async (e) => {
      const btn = e.target;
      const originalText = btn.innerHTML;
      btn.innerHTML = '⏳ Loading...';
      btn.disabled = true;
      try {
        await exportOptions.onExport();
      } catch (err) {
        toastError('Gagal export data');
      } finally {
        btn.innerHTML = originalText;
        btn.disabled = false;
      }
    });

    document.getElementById(`btn-template-${exportOptions.moduleName}`)?.addEventListener('click', () => {
      exportOptions.onTemplate();
    });

    const fileInput = document.getElementById(`input-import-${exportOptions.moduleName}`);
    fileInput?.addEventListener('change', async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      const label = document.getElementById(`label-import-${exportOptions.moduleName}`);
      const span = label ? label.querySelector('.import-text') : null;
      const originalText = span ? span.innerText : '';
      if (span) span.innerText = '⌛ Memproses...';
      if (label) label.style.pointerEvents = 'none';
      fileInput.disabled = true;
      
      try {
        const json = await parseExcel(file);
        if (json.length === 0) throw new Error('File kosong atau format salah');
        await exportOptions.onImport(json);
        toastSuccess('Import berhasil!');
        notifyCalendar(apiPath);
        load();
      } catch (err) {
        toastError(err.message || 'Gagal import data');
      } finally {
        if (span) span.innerText = originalText;
        if (label) label.style.pointerEvents = 'auto';
        fileInput.disabled = false;
        fileInput.value = ''; // reset
      }
    });
  }

  async function load() {
    // JANGAN clear selectedIds di sini agar checklist tetap terjaga saat ganti halaman
    // selectedIds hanya di-clear saat: Batalkan, setelah Hapus, atau saat filter berubah
    updateBulkToolbar();
    
    const tableContainer = document.getElementById('table-container');
    if (!tableContainer) return;
    tableContainer.innerHTML = '<div class="loading-spinner"><div class="spinner"></div></div>';

    // FIX: Tentukan mode paginasi (Hybrid Architecture)
    // Mode ditentukan dari konfigurasi (default: server).
    // Jika 'client', kita butuh semua data di frontend (Client-Side).
    // Jika 'server', kita gunakan paginasi Backend yang scalable untuk data raksasa.
    const isClientSide = paginationMode === 'client';
    const apiPage = isClientSide ? 1 : page;
    const apiLimit = isClientSide ? CLIENT_SIDE_MAX_ROWS : 20;
    
    const params = new URLSearchParams({ page: apiPage, limit: apiLimit, ...Object.fromEntries(Object.entries(filters).filter(([, v]) => v)) });
    const res = await apiFetch(`${apiPath}?${params}`);

    if (!res.ok) {
      tableContainer.innerHTML = `<div class="empty-state"><p class="text-danger">Gagal memuat data: ${res.data?.error || 'Error'}</p></div>`;
      return;
    }

    let items = res.data?.data || res.data || [];
    let pagination = res.data?.pagination;
    const originalTotal = items.length;
    
    if (isClientSide) {
       // 1. Terapkan filter khusus (Client-Side)
       items = onDataLoaded(items);
       
       // 2. Hitung jumlah total data setelah difilter
       const filteredTotal = items.length;
       const limit = 20;
       const pages = Math.ceil(filteredTotal / limit);
       
       // 3. Pastikan current page tidak melebihi total pages
       if (page > pages && pages > 0) page = pages;
       
       // 4. Hitung index slicing
       const startIndex = (page - 1) * limit;
       const endIndex = page * limit;
       
       // 5. Potong array data sesuai halaman
       items = items.slice(startIndex, endIndex);
       
       // 6. Tumpuk pagination object milik backend dengan milik frontend
       pagination = { 
          page: page, 
          limit: limit, 
          total: filteredTotal, 
          pages: pages 
       };
    }

    if (IS_DEVELOPMENT) {
      console.log({
        mode: isClientSide ? 'Client-Side' : 'Server-Side',
        module: apiPath,
        totalData: originalTotal,
        filteredData: items.length,
        currentPage: page,
        pageSize: pagination ? pagination.limit : 20,
        totalPages: pagination ? pagination.pages : 1,
        startIndex: isClientSide ? (page - 1) * 20 : 0,
        endIndex: isClientSide ? page * 20 : items.length,
        rowsRendered: items.length
      });
    }

    if (onAfterLoad) onAfterLoad(items);

    const table = createTable({
      columns,
      data: items,
      onEdit: canEdit ? (row) => openForm(row) : null,
      // Individual onDelete removed
      actions: extraActions.map(a => ({ ...a, handler: (row) => a.handler(row, load) })),
      emptyText: `Tidak ada ${String(itemLabel || '').toLowerCase()}`,
      bulkSelect: bulkDelete ? { selectedIds, onToggle: updateBulkToolbar } : null,
    });

    tableContainer.innerHTML = '';
    tableContainer.appendChild(table);

    // Pagination
    const pagEl = document.getElementById('pagination-container');
    if (pagEl) {
      pagEl.innerHTML = '';
      if (pagination && pagination.pages > 1) {
        const pag = createPagination({
          page: pagination.page,
          pages: pagination.pages,
          total: pagination.total,
          limit: pagination.limit,
          onPage: (p) => { page = p; load(); },
        });
        if (pag) pagEl.appendChild(pag);
      }
    }
  }

  function buildForm(data) {
    const fields = typeof formFields === 'function' ? formFields(data) : formFields;
    return buildFormHTML(fields);
  }

  function openForm(data) {
    const isEdit = !!data;
    const formEl = document.createElement('form');
    formEl.noValidate = true;
    formEl.innerHTML = buildForm(data);

    if (isEdit) {
      const fields = typeof formFields === 'function' ? formFields(data) : formFields;
      populateForm(formEl, data);
    }

    const { close } = createModal({
      title: isEdit ? `Edit ${itemLabel}` : `Tambah ${itemLabel}`,
      content: formEl,
      size: 'lg',
      confirmText: isEdit ? 'Simpan Perubahan' : `Tambah ${itemLabel}`,
      onConfirm: async (overlay, closeModal) => {
        if (!formEl.reportValidity()) return;

        const confirmBtn = overlay.querySelector('.modal-confirm');
        confirmBtn.disabled = true;
        confirmBtn.textContent = 'Menyimpan...';

        let body = getFormData(formEl);

        const fields = typeof formFields === 'function' ? formFields(data) : formFields;
        const processComboboxes = async (fList) => {
          for (const f of fList) {
            if (f.type === 'row') await processComboboxes(f.fields);
            else if (f.type === 'combobox' && body[f.name]) {
              const valStr = body[f.name];
              const existing = (f.options || []).find(o => {
                 const v = typeof o === 'object' ? String(o.value) : String(o);
                 const l = typeof o === 'object' ? String(o.label) : String(o);
                 return v === valStr || l === valStr;
              });
              if (existing) {
                 body[f.name] = typeof existing === 'object' ? existing.value : existing;
              } else if (f.createApi) {
                 const payload = {};
                 payload[f.createApi.field] = valStr;
                 if (f.createApi.extra) Object.assign(payload, f.createApi.extra);
                 const cRes = await apiFetch(f.createApi.path, { method: 'POST', body: JSON.stringify(payload) });
                 if (cRes.ok && cRes.data?.id) body[f.name] = cRes.data.id;
                 else if (cRes.ok && !cRes.data?.id) body[f.name] = valStr;
                 else throw new Error(`Gagal membuat master data: ${cRes.data?.error || 'Unknown error'}`);
              }
            }
          }
        };

        try {
          await processComboboxes(fields);
        } catch(e) {
          toastError(e.message);
          confirmBtn.disabled = false;
          confirmBtn.textContent = isEdit ? 'Simpan Perubahan' : `Tambah ${itemLabel}`;
          return;
        }

        if (onBeforeSubmit) body = await onBeforeSubmit(body, data);

        const method = isEdit ? 'PUT' : 'POST';
        const url = isEdit ? `${apiPath}/${data.id}` : apiPath;
        const res = await apiFetch(url, { method, body: JSON.stringify(body) });

        if (res.ok) {
          toastSuccess(isEdit ? `${itemLabel} berhasil diperbarui.` : `${itemLabel} berhasil ditambahkan.`);
          closeModal();
          notifyCalendar(apiPath);
          load();
        } else {
          toastError(res.data?.error || 'Gagal menyimpan data.');
          confirmBtn.disabled = false;
          confirmBtn.textContent = isEdit ? 'Simpan Perubahan' : `Tambah ${itemLabel}`;
        }
      },
    });
  }

  function handleDelete(row) {
    confirmDialog(
      `Hapus ${itemLabel} ini? Tindakan tidak dapat dibatalkan.`,
      async () => {
        const res = await apiFetch(`${apiPath}/${row.id}`, { method: 'DELETE' });
        if (res.ok) {
          toastSuccess(`${itemLabel} berhasil dihapus.`);
          notifyCalendar(apiPath);
          load();
        } else {
          toastError(res.data?.error || 'Gagal menghapus.');
        }
      },
      `Hapus ${itemLabel}`
    );
  }

  load();
  return load; // expose reload function
}
