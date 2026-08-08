// Generic CRUD page builder - used by all modules
import { apiFetch, CLIENT_SIDE_MAX_ROWS, IS_DEVELOPMENT, getUser } from '../config.js';
import { createTable, createPagination } from '../components/table.js';
import { createModal, confirmDialog } from '../components/modal.js';
import { buildFormHTML, getFormData, populateForm } from '../components/form.js';
import { toastSuccess, toastError } from '../components/toast.js';
import { renderExcelButtons, parseExcel } from '../utils/excel.js';

export function buildCrudPage({
  container,
  title,
  icon,
  apiPath,
  columns,
  formFields,
  filterFields,
  defaultFilters = {},
  enableMobileFilterSheet = false,
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
  const user = getUser();
  if (user && typeof user === 'object' && user.role === 'viewer') {
    canCreate = false;
    canEdit = false;
    canDelete = false;
    bulkDelete = false;
    exportOptions = null;
  }
  let page = 1;
  let filters = { ...defaultFilters };
  if (initialSearch) filters.search = initialSearch;
  let selectedIds = new Set();

  container.innerHTML = `
    <div class="crud-layout-wrapper ${enableMobileFilterSheet ? 'mobile-active' : ''}">
      <div class="page-header">
      <h1 class="page-title">${icon} ${title}</h1>
      <div class="page-actions">
        ${canCreate ? `<button class="btn btn-primary" id="btn-create">+ Tambah ${itemLabel}</button>` : ''}
        ${exportOptions ? `<button class="btn btn-outline" id="btn-mobile-aksi" style="display:none; align-items:center; justify-content:center; gap:0.25rem;">⋮ Aksi</button>` : ''}
      </div>
    </div>

    ${bulkDelete ? `
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:none; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>🗑️ Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>` : ''}
    
    ${exportOptions ? `
    <div class="excel-actions-wrapper" id="excel-actions-wrapper">
      <div class="bottom-sheet-header aksi-header" style="display:none;">
        <h3 style="margin:0; font-size:1rem;">Aksi</h3>
        <button class="btn-close-sheet" id="btn-close-aksi-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
      </div>
      ${renderExcelButtons(exportOptions.moduleName)}
    </div>` : ''}

    ${filterFields && filterFields.length > 0 ? `
    <div class="filter-bar card ${enableMobileFilterSheet ? 'has-mobile-sheet' : ''}" style="padding: 1rem; display: flex; gap: 0.5rem; flex-wrap: wrap;">
        ${filterFields.filter(f => f.type === 'search' || f.type === 'search-combo').map(f => {
          if (f.type === 'search') return `<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" class="form-control" autocomplete="off" placeholder="${f.placeholder || 'Cari...'}" id="filter-search" value="${filters.search || ''}"></div>`;
          if (f.type === 'search-combo') {
            const dlId = `dl-filter-search`;
            const opts = (f.options || []).map(o => `<option value="${typeof o === 'object' ? o.label : o}"></option>`).join('');
            return `<div class="filter-search" style="flex:1; min-width:120px;"><input type="text" list="${dlId}" class="form-control" autocomplete="off" placeholder="${f.placeholder || 'Cari...'}" id="filter-search" value="${filters.search || ''}"><datalist id="${dlId}">${opts}</datalist></div>`;
          }
          return '';
        }).join('')}
        
        <div class="filter-options" id="filter-options-wrapper">
          <div class="bottom-sheet-header">
            <h3 style="margin:0; font-size:1rem;">Filter Data</h3>
            <button class="btn-close-sheet" id="btn-close-filter-sheet" style="background:none;border:none;font-size:1.5rem;cursor:pointer;line-height:1;">&times;</button>
          </div>
          ${filterFields.filter(f => f.type !== 'search' && f.type !== 'search-combo').map(f => {
            if (f.type === 'select') return `<select class="form-control filter-select" style="flex:1; min-width:100px;" name="${f.name}" id="filter-${f.name}"><option value="">Pilih ${f.label}</option>${(f.options || []).map(o => `<option value="${typeof o === 'object' ? o.value : o}" ${filters[f.name] === (typeof o === 'object' ? o.value : o) ? 'selected' : ''}>${typeof o === 'object' ? o.label : o}</option>`).join('')}</select>`;
            if (f.type === 'combobox') {
              const dlId = `dl-filter-${f.name}`;
              const cbOpts = (f.options || []).map(o => {
                let lbl = typeof o === 'object' ? (o.label || o.value || '') : (o || '');
                if (lbl === 'undefined' || lbl === '[object Object]' || lbl === 'null') lbl = '';
                if (!lbl) return '';
                return `<option value="${lbl}"></option>`;
              }).join('');
              let displayVal = filters[f.name] || '';
              if (filters[f.name]) {
                  const found = (f.options || []).find(o => (typeof o === 'object' ? String(o.value) : String(o)) == String(filters[f.name]));
                  if (found) {
                     let foundLbl = typeof found === 'object' ? (found.label || found.value || '') : (found || '');
                     if (foundLbl && foundLbl !== 'undefined' && foundLbl !== '[object Object]' && foundLbl !== 'null') {
                        displayVal = foundLbl;
                     }
                  }
              }
              return `<div class="filter-combobox" style="flex:1; min-width:120px;">
                <input type="text" name="${f.name}" id="filter-${f.name}" list="${dlId}" class="form-control filter-combobox-input" value="${displayVal}" placeholder="Pilih ${f.label}..." autocomplete="off">
                <datalist id="${dlId}">${cbOpts}</datalist>
              </div>`;
            }
            return '';
          }).join('')}
          <div style="display:flex; gap:0.5rem; margin-top:0.5rem;">
            <button class="btn btn-outline" id="btn-reset-filter" style="flex:1;">Reset</button>
            <button class="btn btn-primary" id="btn-apply-filter" style="flex:1;">✓ Terapkan (OK)</button>
          </div>
        </div>
        ${enableMobileFilterSheet ? `<button id="btn-mobile-filter" class="btn btn-outline" style="display:none; align-items:center; gap:0.25rem;">⚙️ Filter</button>` : ''}
    </div>` : ''}

    <div class="card">
      <div class="card-body p-0" id="table-container">
        <div class="loading-spinner"><div class="spinner"></div></div>
      </div>
      <div class="card-footer" id="pagination-container"></div>
    </div>
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
      toolbar.classList.add('has-items');
      btnDelete.disabled = false;
      btnCancel.disabled = false;
      
      // Bypassing aggressive mobile CSS cache by forcing position inline
      if (window.innerWidth <= 768) {
        toolbar.style.setProperty('top', 'auto', 'important');
        toolbar.style.setProperty('bottom', '0', 'important');
      } else {
        toolbar.style.removeProperty('top');
        toolbar.style.removeProperty('bottom');
      }
    } else {
      toolbar.style.display = 'none';
      toolbar.classList.remove('has-items');
      btnDelete.disabled = true;
      btnCancel.disabled = true;
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
      load();
    }, 400);
  });

  filterFields?.forEach(f => {
    if (f.type === 'select') {
      document.getElementById(`filter-${f.name}`)?.addEventListener('change', (e) => {
        filters[f.name] = e.target.value;
        page = 1;
        selectedIds.clear();
        load();
      });
    }
    if (f.type === 'combobox') {
      document.getElementById(`filter-${f.name}`)?.addEventListener('change', (e) => {
        let valStr = e.target.value;
        const existing = (f.options || []).find(o => {
           const v = typeof o === 'object' ? String(o.value) : String(o);
           const l = typeof o === 'object' ? String(o.label) : String(o);
           return v === valStr || l === valStr;
        });
        if (!valStr) {
           filters[f.name] = '';
        } else {
           filters[f.name] = existing ? (typeof existing === 'object' ? existing.value : existing) : valStr;
        }
        page = 1;
        selectedIds.clear();
        load();
      });
    }
  });

  document.getElementById('btn-reset-filter')?.addEventListener('click', () => {
    filters = { ...defaultFilters };
    if (searchInput) searchInput.value = '';
    filterFields?.forEach(f => {
      const el = document.getElementById(`filter-${f.name}`);
      if (el) el.value = '';
    });
    page = 1;
    selectedIds.clear();
    load();
  });

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
    const btnApplyFilter = document.getElementById('btn-apply-filter');
    if (btnApplyFilter) {
      btnApplyFilter.addEventListener('click', (e) => {
        e.preventDefault();
        // Since filtering is auto-applied on change, we just need to close the sheet here
        filterOptionsWrapper.classList.remove('sheet-open');
      });
    }
  }

  // Mobile Aksi Sheet Logic
  const btnMobileAksi = document.getElementById('btn-mobile-aksi');
  const excelActionsWrapper = document.getElementById('excel-actions-wrapper');
  const btnCloseAksiSheet = document.getElementById('btn-close-aksi-sheet');

  if (btnMobileAksi && excelActionsWrapper) {
    btnMobileAksi.addEventListener('click', (e) => {
      e.preventDefault();
      excelActionsWrapper.classList.add('sheet-open');
    });
    if (btnCloseAksiSheet) {
      btnCloseAksiSheet.addEventListener('click', (e) => {
        e.preventDefault();
        excelActionsWrapper.classList.remove('sheet-open');
      });
    }
  }

  // Create button
  document.getElementById('btn-create')?.addEventListener('click', () => openForm(null));

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
      fileInput.disabled = true;
      
      const overlay = document.createElement('div');
      overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.7);z-index:9999;display:flex;align-items:center;justify-content:center';
      overlay.innerHTML = `
        <div style="background:var(--bg-card);border-radius:var(--radius-xl);padding:32px;width:90%;max-width:500px;box-shadow:var(--shadow-lg);text-align:center;">
          <h3 style="margin:0 0 16px;color:var(--text-1);font-size:1.2rem">🔄 Memproses Import Data</h3>
          <div style="margin-bottom:16px;color:var(--text-2);font-size:0.9rem" id="import-progress-text">Membaca file Excel...</div>
          <div style="background:var(--bg-body);border-radius:999px;height:12px;overflow:hidden;margin-bottom:24px">
            <div id="import-progress-bar" style="background:var(--primary);height:100%;width:0%;transition:width 0.3s"></div>
          </div>
          <div id="import-summary" style="display:none;text-align:left;background:var(--bg-body);padding:16px;border-radius:8px;margin-bottom:24px;font-size:0.9rem"></div>
          <button id="import-close-btn" class="btn btn-primary" style="display:none;width:100%">Selesai</button>
        </div>
      `;
      document.body.appendChild(overlay);
      const textEl = overlay.querySelector('#import-progress-text');
      const barEl = overlay.querySelector('#import-progress-bar');
      const summaryEl = overlay.querySelector('#import-summary');
      const closeBtn = overlay.querySelector('#import-close-btn');

      closeBtn.addEventListener('click', () => {
        overlay.remove();
        load();
      });

      try {
        const json = await parseExcel(file);
        if (json.length === 0) throw new Error('File kosong atau format salah');
        
        // Chunking Logic (Stress Test Ready: 100 - 10,000 rows)
        const CHUNK_SIZE = 500;
        let inserted = 0, skipped = 0, failed = 0;
        const total = json.length;
        
        textEl.textContent = `Ditemukan ${total} baris data. Memulai import...`;
        
        for (let i = 0; i < total; i += CHUNK_SIZE) {
          const chunk = json.slice(i, i + CHUNK_SIZE);
          textEl.textContent = `Mengimport baris ${i + 1} - ${Math.min(i + CHUNK_SIZE, total)} dari ${total}...`;
          barEl.style.width = `${Math.round((i / total) * 100)}%`;
          
          try {
            // onImport should return { inserted, skipped } or throw
            const result = await exportOptions.onImport(chunk);
            if (result) {
              inserted += result.inserted || result.metrics?.inserted || chunk.length;
              skipped += result.skipped || result.metrics?.updated || 0;
            } else {
              inserted += chunk.length; // fallback
            }
          } catch (err) {
            console.error('Chunk import failed:', err);
            failed += chunk.length;
          }
        }
        
        barEl.style.width = '100%';
        textEl.innerHTML = `<strong style="color:var(--success)">✅ Import Selesai!</strong>`;
        
        summaryEl.style.display = 'block';
        summaryEl.innerHTML = `
          <div style="display:flex;justify-content:space-between;margin-bottom:8px"><span>Total Data:</span> <strong>${total}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--success)"><span>Berhasil Diimport:</span> <strong>${inserted}</strong></div>
          <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:var(--warning)"><span>Dilewati (Duplikat):</span> <strong>${skipped}</strong></div>
          <div style="display:flex;justify-content:space-between;color:var(--danger)"><span>Gagal:</span> <strong>${failed}</strong></div>
        `;
        if (failed > 0) {
          summaryEl.innerHTML += `<p style="margin-top:12px;font-size:0.8rem;color:var(--danger)">Sebagian data gagal diimport. Pastikan format kolom sesuai template dan tidak ada data kosong pada kolom wajib.</p>`;
        }
        
        closeBtn.style.display = 'block';
        fileInput.value = ''; // reset
      } catch (err) {
        textEl.innerHTML = `<strong style="color:var(--danger)">❌ Gagal Memproses File</strong><br>${err.message}`;
        barEl.style.background = 'var(--danger)';
        barEl.style.width = '100%';
        closeBtn.style.display = 'block';
        fileInput.value = ''; // reset
      } finally {
        fileInput.disabled = false;
      }
    });
  }

  async function load() {
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
    let fullItems = items;
    
    if (isClientSide) {
       // 1. Terapkan filter khusus (Client-Side)
       items = onDataLoaded(items);
       fullItems = items;
       
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
      fullData: fullItems,
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
