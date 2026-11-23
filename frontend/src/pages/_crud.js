// Generic CRUD page builder - used by all modules
import { apiFetch, IS_DEVELOPMENT, getUser, hasPermission } from '../config.js';
import { escapeHtml } from '../utils/sanitize.js';
import { createTable, createPagination, MAX_BULK_IDS, toggleBulkSelectionItem, toggleBulkSelectionPage } from '../components/table.js';
import { createModal, confirmDialog } from '../components/modal.js';
import { buildFormHTML, getFormData, populateForm } from '../components/form.js';
import { mountEmployeeCombobox } from '../components/employeeCombobox.js';
import { toastSuccess, toastError } from '../components/toast.js';
import { renderExcelButtons, parseExcel } from '../utils/excel.js';
import { notifyCalendar } from '../utils/calendarBus.js';
import {
  clearImportSession,
  getIncompleteImportSessions,
  getImportSession,
  renderImportStatusBannerHTML,
  resumeImportSession,
} from '../utils/import_client.js';

export { MAX_BULK_IDS, toggleBulkSelectionItem, toggleBulkSelectionPage };

export function buildCrudPage({
  container,
  title,
  icon,
  apiPath,
  columns,
  formFields,
  filterFields,
  defaultFilters = {},
  defaultSort = '',
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
  mobileCardRenderer = null, // optional (row, handlers) => HTMLElement for mobile card layout
}) {
  const user = getUser();
  const moduleName = (apiPath || '').replace('/api/', '').split('/')[0].replace(/-/g, '_');

  const userCanCreate = hasPermission(user, moduleName, 'create');
  const userCanEdit   = hasPermission(user, moduleName, 'update');
  const userCanDelete = hasPermission(user, moduleName, 'delete');
  const userCanExport = hasPermission(user, moduleName, 'export') || hasPermission(user, 'export');
  const userCanImport = hasPermission(user, moduleName, 'import');

  canCreate  = canCreate && userCanCreate;
  canEdit    = canEdit && userCanEdit;
  canDelete  = canDelete && userCanDelete;
  bulkDelete = bulkDelete && userCanDelete;

  if (!userCanExport && !userCanImport) {
    exportOptions = null;
  } else if (exportOptions) {
    if (!userCanImport) {
      exportOptions.import = false;
      exportOptions.template = false;
    }
    if (!userCanExport) {
      exportOptions.export = false;
    }
    const hasExportAction = exportOptions.export !== false && typeof exportOptions.onExport === 'function';
    const hasTemplateAction = exportOptions.template !== false && typeof exportOptions.onTemplate === 'function';
    const hasImportAction = exportOptions.import !== false && typeof exportOptions.onImport === 'function';
    if (!hasExportAction && !hasTemplateAction && !hasImportAction) {
      exportOptions = null;
    }
  }

  if (!canEdit && Array.isArray(extraActions)) {
    extraActions = extraActions.filter(a => a && a.readOnly === true);
  }

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
      <h1 class="page-title">${icon} ${escapeHtml(title)}</h1>
      <div class="page-actions" style="display:flex; gap:8px; align-items:center;">
        ${canCreate ? `<button class="btn btn-primary" id="btn-create">+ Tambah ${escapeHtml(itemLabel)}</button>` : ''}
        ${exportOptions ? `
          <div class="aksi-dropdown-container" style="position:relative; display:inline-block;">
            <button class="btn btn-ghost" id="btn-aksi-main" style="background:#131F36; border:1px solid #223456; padding:8px 16px; border-radius:8px; font-weight:600; color:#F8FAFC; display:flex; align-items:center; gap:6px; cursor:pointer;">
              ⋮ Aksi
            </button>
            <div id="aksi-menu-main" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#131F36; border:1px solid #223456; border-radius:8px; box-shadow:0 8px 24px rgba(0,0,0,0.5); flex-direction:column; min-width:200px; z-index:999; padding:8px 0;">
              ${exportOptions.export !== false && typeof exportOptions.onExport === 'function' ? `
              <button type="button" class="dropdown-item" id="btn-export-${exportOptions.moduleName || 'default'}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#F8FAFC; display:flex; gap:8px; align-items:center;">
                📥 Export Excel
              </button>` : ''}
              ${exportOptions.template !== false && typeof exportOptions.onTemplate === 'function' ? `
              <button type="button" class="dropdown-item" id="btn-template-${exportOptions.moduleName || 'default'}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#F8FAFC; display:flex; gap:8px; align-items:center;">
                📄 Download Template
              </button>` : ''}
              ${exportOptions.import !== false && typeof exportOptions.onImport === 'function' ? `
              <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#F8FAFC; margin:0; gap:8px; align-items:center;" id="label-import-${exportOptions.moduleName || 'default'}">
                <span class="import-text">📤 Import Excel</span>
                <input type="file" id="input-import-${exportOptions.moduleName || 'default'}" accept=".xlsx, .xls, .csv" style="display:none;">
              </label>` : ''}
            </div>
          </div>
          <style>
            .show-aksi-menu { display: flex !important; }
            .dropdown-item:hover { background-color: #1A2A48 !important; color: #60A5FA !important; }
          </style>
        ` : ''}
      </div>
    </div>
    

    ${filterFields && filterFields.length > 0 ? `
    <div class="filter-bar" style="background: #131F36; border-radius: 12px; padding: 10px 16px; display: flex; align-items: center; gap: 8px; margin-bottom: 24px; border: 1px solid #223456; box-shadow: 0 4px 12px rgba(0,0,0,0.22);">
        ${filterFields.filter(f => f.type === 'search').map(f => {
          return `<div class="filter-search-wrap" style="flex:1; min-width:0;"><input type="search" class="filter-search" placeholder="${escapeHtml(f.placeholder || 'Cari...')}" id="filter-search" value="${escapeHtml(filters.search || '')}" style="width:100%; background: #0E182B; border: 1px solid #223456; color: #F8FAFC; border-radius: 8px; padding: 8px 12px; font-size: 0.85rem; outline:none;"></div>`;
        }).join('')}
        
        <div class="filter-dropdowns-desktop">
          ${filterFields.filter(f => f.type !== 'search').map(f => {
            if (f.type === 'select' || f.type === 'combobox') {
              const placeholder = (f.label || '').startsWith('Pilih') ? f.label : `Pilih ${f.label || ''}`;
              return `<select class="filter-select" name="${escapeHtml(f.name)}" id="filter-${escapeHtml(f.name)}" style="background: #0F1A2E; border: 1px solid #223456; border-radius: 8px; padding: 7px 10px; font-size: 0.85rem; color: #F8FAFC; cursor: pointer; outline:none;"><option value="">${escapeHtml(placeholder)}</option>${(f.options || []).map(o => `<option value="${escapeHtml(typeof o === 'object' ? o.value : o)}" ${filters[f.name] === (typeof o === 'object' ? o.value : o) ? 'selected' : ''}>${escapeHtml(typeof o === 'object' ? o.label : o)}</option>`).join('')}</select>`;
            }
            return '';
          }).join('')}
          <button id="btn-reset-filter" style="background: transparent; border: none; color: #60A5FA; font-weight: 600; font-size: 0.85rem; cursor: pointer; padding: 7px 8px; white-space:nowrap;">Reset</button>
        </div>
        
        <button id="btn-mobile-filter" class="btn-mobile-filter-trigger">⚙ Filter</button>
        
        <div class="filter-options-wrapper" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem; color:#F8FAFC;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;color:#A5B4C8;">&times;</button>
          </div>
          ${filterFields.filter(f => f.type !== 'search').map(f => {
            if (f.type === 'select' || f.type === 'combobox') {
              const placeholder = (f.label || '').startsWith('Pilih') ? f.label : `Pilih ${f.label || ''}`;
              return `<select class="filter-select filter-select-sheet" name="${escapeHtml(f.name)}-sheet" id="filter-sheet-${escapeHtml(f.name)}" style="width:100%; background: #0F1A2E; border: 1px solid #223456; border-radius: 10px; padding: 12px 14px; font-size: 0.9rem; color: #F8FAFC; cursor: pointer; outline:none;"><option value="">${escapeHtml(placeholder)}</option>${(f.options || []).map(o => `<option value="${escapeHtml(typeof o === 'object' ? o.value : o)}" ${filters[f.name] === (typeof o === 'object' ? o.value : o) ? 'selected' : ''}>${escapeHtml(typeof o === 'object' ? o.label : o)}</option>`).join('')}</select>`;
            }
            return '';
          }).join('')}
          <button id="btn-reset-filter-sheet" style="background: transparent; border: none; color: #60A5FA; font-weight: 600; font-size: 0.9rem; cursor: pointer; padding: 8px;">Reset</button>
        </div>
    </div>` : ''}

    <div id="crud-import-banner" style="display:none;margin-bottom:16px;"></div>

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
  `;

  document.getElementById('btn-aksi-main')?.addEventListener('click', (e) => {
    e.stopPropagation();
    document.getElementById('aksi-menu-main')?.classList.toggle('show-aksi-menu');
  });

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
    if (ids.length > MAX_BULK_IDS) {
      toastError(`Maksimum ${MAX_BULK_IDS} item dapat dihapus dalam satu permintaan bulk.`);
      return;
    }
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
      const effectiveFilters = defaultSort && !filters.sort ? { ...filters, sort: defaultSort } : { ...filters };
      try {
        await exportOptions.onExport(effectiveFilters);
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
    const importBannerEl = document.getElementById('crud-import-banner');
    let isCrudImporting = false;
    let lastUploadJson = null;

    function setImportBusy(busy) {
      isCrudImporting = busy;
      const label = document.getElementById(`label-import-${exportOptions.moduleName}`);
      const span = label ? label.querySelector('.import-text') : null;
      if (span) span.innerText = busy ? '⌛ Memproses...' : '📤 Import Excel';
      if (label) {
        label.style.pointerEvents = busy ? 'none' : 'auto';
        label.style.opacity = busy ? '0.6' : '1';
      }
      if (fileInput) fileInput.disabled = busy;
    }

    function showCrudImportBanner(state) {
      if (!importBannerEl) return;
      if (!state) {
        importBannerEl.style.display = 'none';
        importBannerEl.innerHTML = '';
        return;
      }
      importBannerEl.style.display = '';
      importBannerEl.innerHTML = renderImportStatusBannerHTML(state);
      importBannerEl.querySelectorAll('[data-import-action]').forEach((btn) => {
        btn.addEventListener('click', async () => {
          const action = btn.getAttribute('data-import-action');
          const fp = btn.getAttribute('data-fingerprint');
          if (action === 'dismiss') {
            if (fp) clearImportSession(fp);
            showCrudImportBanner(null);
            return;
          }
          if (isCrudImporting) {
            toastError('Import masih berjalan. Mohon tunggu hingga selesai.');
            return;
          }
          await handleBannerImportAction(action, fp);
        });
      });
    }

    function displayCrudImportSuccess(result, usedKey = '') {
      const d = result?.data || result || {};
      const isReplay = Boolean(result?.idempotentReplay || d?.idempotentReplay);
      const inserted = Number(d.inserted || 0);
      const updated = Number(d.updated || 0);
      const skipped = Number(d.skipped || 0);
      const summaryParts = [];
      if (inserted > 0) summaryParts.push(`${inserted} baru`);
      if (updated > 0) summaryParts.push(`${updated} diperbarui`);
      if (skipped > 0) summaryParts.push(`${skipped} dilewati`);
      const detailText = d.message || (summaryParts.length ? summaryParts.join(', ') : 'Data berhasil diproses');
      showCrudImportBanner({
        variant: 'success',
        title: isReplay ? '♻️ Replay Idempoten Dikonfirmasi Server' : '✅ Import Berhasil Dikonfirmasi',
        message: isReplay
          ? `${detailText} (Server mengembalikan hasil tersimpan tanpa menulis data duplikat.)`
          : detailText,
        idempotencyKey: result?.idempotencyKey || usedKey || '',
        importJobId: result?.importJobId || d?.importJobId || '',
        fingerprint: result?.fingerprint || '',
        canReplay: Boolean((result?.fingerprint || lastUploadJson) && (result?.idempotencyKey || usedKey)),
        canDismiss: true,
      });
      toastSuccess(isReplay ? 'Replay idempoten berhasil (0 duplikat).' : 'Import berhasil!');
      notifyCalendar(apiPath);
      load();
    }

    function displayCrudImportError(err) {
      const isConflict = Number(err?.status) === 409;
      const isActiveLock = isConflict && err?.code === 'IMPORT_IN_PROGRESS';
      const isNetwork = Boolean(err?.isNetworkError || Number(err?.status) === 0);
      showCrudImportBanner({
        variant: isActiveLock ? 'warning' : 'error',
        title: isNetwork
          ? '📡 Koneksi Terputus Saat Import'
          : isActiveLock
            ? '⏳ Import Sedang Berjalan / Terkunci (HTTP 409)'
            : Number(err?.status) >= 500
              ? '⚠️ Import Gagal Sebagian / Server Error (HTTP 500)'
              : `❌ Import Gagal${err?.status ? ` (HTTP ${err.status})` : ''}`,
        message: err?.message || 'Gagal mengimport data',
        idempotencyKey: err?.idempotencyKey || '',
        importJobId: err?.importJobId || '',
        committedChunks: err?.committedChunks || 0,
        totalChunks: err?.totalChunks || 0,
        errorCode: err?.code || '',
        fingerprint: err?.fingerprint || '',
        canRetry: Boolean(err?.recoverable || isNetwork || Number(err?.status) >= 500 || isActiveLock),
        canRecover: isActiveLock,
        canDismiss: true,
      });
      toastError(err?.message || 'Gagal import data');
    }

    async function handleBannerImportAction(action, fp) {
      const session = fp ? getImportSession(fp) : null;
      const recover = action === 'recover';
      const usedKey = session?.idempotencyKey || '';
      setImportBusy(true);
      showCrudImportBanner({
        variant: 'progress',
        title: recover ? '🛠️ Memulihkan & Melanjutkan Import...' : '⏳ Mengirim / Melanjutkan Import...',
        message: 'Mohon tunggu konfirmasi dari server. Jangan tutup halaman.',
        idempotencyKey: usedKey,
        importJobId: session?.importJobId || '',
        committedChunks: session?.committedChunks || 0,
        totalChunks: session?.totalChunks || 0,
        fingerprint: fp || '',
        canDismiss: false,
      });
      try {
        let result;
        if (action === 'replay' && lastUploadJson && usedKey) {
          result = await exportOptions.onImport(lastUploadJson, { idempotencyKey: usedKey, recover: false });
        } else if (session?.payloadStored) {
          result = await resumeImportSession(fp, { recover });
        } else if (lastUploadJson) {
          result = await exportOptions.onImport(lastUploadJson, { idempotencyKey: usedKey || null, recover });
        } else {
          result = await resumeImportSession(fp, { recover });
        }
        displayCrudImportSuccess(result, usedKey);
      } catch (err) {
        displayCrudImportError(err);
      } finally {
        setImportBusy(false);
      }
    }

    // Check interrupted / partial sessions after page refresh
    const pendingModuleSessions = getIncompleteImportSessions(exportOptions.moduleName);
    if (pendingModuleSessions.length > 0) {
      const latest = pendingModuleSessions[0];
      const staleInProgress = latest.status === 'IN_PROGRESS' || latest.status === 'CONFLICT_IN_PROGRESS';
      showCrudImportBanner({
        variant: 'recover',
        title: '🔄 Sesi Import Belum Selesai Terdeteksi',
        message: latest.payloadStored
          ? (latest.errorMessage || 'Import sebelumnya terhenti karena refresh atau gangguan koneksi. Anda dapat melanjutkan dengan Idempotency-Key yang sama.')
          : 'Import sebelumnya terhenti. Pilih ulang file Excel yang sama untuk melanjutkan dengan Idempotency-Key yang sama.',
        idempotencyKey: latest.idempotencyKey,
        importJobId: latest.importJobId,
        committedChunks: latest.committedChunks,
        totalChunks: latest.totalChunks,
        errorCode: latest.errorCode || latest.status,
        fingerprint: latest.fingerprint,
        canRetry: Boolean(latest.payloadStored),
        canRecover: Boolean(latest.payloadStored && staleInProgress),
        canDismiss: true,
      });
    }

    fileInput?.addEventListener('change', async (e) => {
      const file = e.target.files[0];
      if (!file) return;
      if (isCrudImporting) {
        toastError('Proses import sedang berjalan. Mohon tunggu hingga selesai.');
        fileInput.value = '';
        return;
      }
      setImportBusy(true);
      showCrudImportBanner({
        variant: 'progress',
        title: '⏳ Membaca & Mengimport File Excel...',
        message: `Memproses "${file.name}". Sistem menunggu konfirmasi server sebelum menandai selesai.`,
        canDismiss: false,
      });

      try {
        const json = await parseExcel(file);
        if (json.length === 0) throw new Error('File kosong atau format salah');
        lastUploadJson = json;
        const result = await exportOptions.onImport(json, { recover: false });
        displayCrudImportSuccess(result);
      } catch (err) {
        displayCrudImportError(err);
      } finally {
        setImportBusy(false);
        if (fileInput) fileInput.value = '';
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

    const apiPage = page;
    const apiLimit = 20;
    const activeQueryFilters = {
      ...(defaultSort ? { sort: defaultSort } : {}),
      ...Object.fromEntries(Object.entries(filters).filter(([, v]) => v)),
    };
    const params = new URLSearchParams({ page: apiPage, limit: apiLimit, ...activeQueryFilters });
    const res = await apiFetch(`${apiPath}?${params}`);

    if (!res.ok) {
      tableContainer.innerHTML = `<div class="empty-state"><p class="text-danger">Gagal memuat data: ${escapeHtml(res.data?.error || 'Error')}</p></div>`;
      return;
    }

    let items = res.data?.data || res.data || [];
    const pagination = res.data?.pagination || {
      page: apiPage,
      limit: apiLimit,
      total: Array.isArray(items) ? items.length : 0,
      pages: 1,
    };
    if (onDataLoaded) {
      items = onDataLoaded(items);
    }

    if (IS_DEVELOPMENT) {
      console.log({
        mode: 'Server-Side',
        module: apiPath,
        totalData: pagination.total,
        filteredData: items.length,
        currentPage: page,
        pageSize: pagination.limit || 20,
        totalPages: pagination.pages || 1,
        rowsRendered: items.length
      });
    }

    if (onAfterLoad) onAfterLoad(items);

    const table = createTable({
      columns,
      data: items,
      onEdit: canEdit ? (row) => openForm(row) : null,
      onDelete: canDelete ? (row) => handleDelete(row) : null,
      actions: extraActions.map(a => ({ ...a, handler: (row) => a.handler(row, load) })),
      emptyText: `Tidak ada ${String(itemLabel || '').toLowerCase()}`,
      bulkSelect: bulkDelete ? {
        selectedIds,
        maxSelect: MAX_BULK_IDS,
        onLimitExceeded: (reason) => toastError(reason),
        onToggle: updateBulkToolbar
      } : null,
      mobileCardRenderer,
    });

    tableContainer.innerHTML = '';
    tableContainer.appendChild(table);
    updateBulkToolbar();

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

    const activeFields = typeof formFields === 'function' ? formFields(data) : formFields;
    if (isEdit) {
      populateForm(formEl, data);
    }

    const empCombos = [];
    const mountEmpFields = (fList) => {
      for (const f of fList) {
        if (f.type === 'row') {
          mountEmpFields(f.fields || []);
        } else if (f.type === 'employee-combobox') {
          const mountEl = formEl.querySelector(`.fc-crud-emp-mount[data-emp-mount="${f.name}"]`);
          if (!mountEl) continue;
          const initName = data ? (data[f.name] ?? f.value ?? '') : (f.value ?? '');
          const initEmpId = (data && f.employeeIdField)
            ? (data[f.employeeIdField] ?? null)
            : (f.employeeIdValue ?? null);
          const ctrl = mountEmployeeCombobox(mountEl, {
            mode: f.mode || 'single',
            inputId: f.id || `crud-${f.name}`,
            inputName: f.name,
            placeholder: f.placeholder || '',
            required: !!f.required,
            allowManual: f.allowManual !== false,
            hasEmployeeIdSchema: f.hasEmployeeIdSchema !== false,
            manualBadgeText: f.manualBadgeText || 'Belum terhubung ke Master Karyawan',
            masterBadgePrefix: f.masterBadgePrefix || 'Terhubung ke Master Karyawan',
            employees: f.options || [],
            initialEmployeeId: initEmpId,
            initialEmployeeName: initName,
          });
          empCombos.push({ field: f, ctrl });
        }
      }
    };
    mountEmpFields(activeFields);

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

        for (const { field, ctrl } of empCombos) {
          const empVal = ctrl.getValue();
          if (field.mode === 'multi') {
            body[field.name] = empVal.participantsString || null;
          } else {
            body[field.name] = empVal.employee_name || null;
            if (field.employeeIdField) {
              body[field.employeeIdField] = empVal.employee_id ?? null;
            }
          }
        }

        if (onBeforeSubmit) body = await onBeforeSubmit(body, data);

        if (isEdit && data) {
          if (data.row_version !== undefined && data.row_version !== null) {
            body.row_version = Number(data.row_version);
          } else {
            // Default to canonical baseline version 1 if DB row_version is null/undefined
            body.row_version = 1;
          }
        }

        const method = isEdit ? 'PUT' : 'POST';
        const url = isEdit ? `${apiPath}/${data.id}` : apiPath;
        const res = await apiFetch(url, { method, body: JSON.stringify(body) });

        if (res.ok) {
          toastSuccess(isEdit ? `${itemLabel} berhasil diperbarui.` : `${itemLabel} berhasil ditambahkan.`);
          closeModal();
          notifyCalendar(apiPath);
          load();
        } else {
          if (res.status === 409) {
            toastError(res.data?.error || 'Konflik OCC: Data telah diubah oleh pengguna lain. Silakan muat ulang halaman.');
          } else {
            toastError(res.data?.error || 'Gagal menyimpan data.');
          }
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
