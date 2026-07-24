// Generic CRUD page builder - used by all modules
import { apiFetch } from '../config.js';
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
}) {
  let page = 1;
  let filters = { ...defaultFilters };
  if (initialSearch) filters.search = initialSearch;
  let selectedIds = new Set();

  container.innerHTML = `
    <div class="page-header">
      <h1 class="page-title">${icon} ${title}</h1>
      <div class="page-actions">
        ${canCreate ? `<button class="btn btn-primary" id="btn-create">+ Tambah ${itemLabel}</button>` : ''}
      </div>
    </div>

    ${bulkDelete ? `
    <div class="bulk-toolbar" id="bulk-toolbar" style="display:flex; align-items:center; gap:1rem; background:var(--bg-card); padding:0.75rem 1.25rem; border-radius:var(--radius-lg); border:1px solid var(--border-color); margin-bottom:1rem;">
      <span id="bulk-count" style="font-weight:600; font-size:0.9rem;">0 item dipilih</span>
      <button class="btn btn-danger btn-sm" id="btn-bulk-delete" disabled>🗑️ Hapus Terpilih</button>
      <button class="btn btn-ghost btn-sm" id="btn-bulk-cancel" disabled>Batalkan</button>
    </div>` : ''}
    
    ${exportOptions ? renderExcelButtons(exportOptions.moduleName) : ''}

    ${filterFields && filterFields.length > 0 ? `
    <div class="filter-bar card">
      <div class="filter-bar-inner">
        ${filterFields.map(f => {
          if (f.type === 'search') return `<div class="filter-search"><input type="search" class="form-control" placeholder="${f.placeholder || 'Cari...'}" id="filter-search" value="${filters.search || ''}"></div>`;
          if (f.type === 'select') return `<select class="form-control filter-select" name="${f.name}" id="filter-${f.name}"><option value="">-- ${f.label} --</option>${(f.options || []).map(o => `<option value="${typeof o === 'object' ? o.value : o}" ${filters[f.name] === (typeof o === 'object' ? o.value : o) ? 'selected' : ''}>${typeof o === 'object' ? o.label : o}</option>`).join('')}</select>`;
          return '';
        }).join('')}
        <button class="btn btn-ghost btn-sm" id="btn-reset-filter">Reset</button>
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
      btnDelete.disabled = false;
      btnCancel.disabled = false;
    } else {
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
      load();
    }, 400);
  });

  filterFields?.forEach(f => {
    if (f.type === 'select') {
      document.getElementById(`filter-${f.name}`)?.addEventListener('change', (e) => {
        filters[f.name] = e.target.value;
        page = 1;
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
    load();
  });

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
    const tableContainer = document.getElementById('table-container');
    if (!tableContainer) return;
    tableContainer.innerHTML = '<div class="loading-spinner"><div class="spinner"></div></div>';

    // FIX: Selalu ambil dari halaman 1 (semua data) dari backend
    // karena kita menggunakan limit 10000 dan mem-paginate di frontend.
    const apiPage = 1;
    const params = new URLSearchParams({ page: apiPage, limit: 10000, ...Object.fromEntries(Object.entries(filters).filter(([, v]) => v)) });
    const res = await apiFetch(`${apiPath}?${params}`);

    if (!res.ok) {
      tableContainer.innerHTML = `<div class="empty-state"><p class="text-danger">Gagal memuat data: ${res.data?.error || 'Error'}</p></div>`;
      return;
    }

    let items = res.data?.data || res.data || [];
    const originalTotal = items.length;
    
    // 1. Terapkan filter khusus (seperti filter dari Dashboard)
    if (onDataLoaded) {
       items = onDataLoaded(items);
    }
    
    // 2. Hitung jumlah total data setelah difilter
    const filteredTotal = items.length;
    const limit = 20;
    const pages = Math.ceil(filteredTotal / limit);
    
    // 3. Pastikan current page tidak melebihi total pages
    if (page > pages && pages > 0) page = pages;
    
    // 4. Hitung index slicing
    const startIndex = (page - 1) * limit;
    const endIndex = page * limit;
    
    // 5. Potong array data sesuai halaman (Client-Side Pagination)
    items = items.slice(startIndex, endIndex);
    
    const pagination = { 
       page: page, 
       limit: limit, 
       total: filteredTotal, 
       pages: pages 
    };

    console.log({
      module: apiPath,
      totalData: originalTotal,
      filteredData: filteredTotal,
      currentPage: page,
      pageSize: limit,
      totalPages: pages,
      startIndex,
      endIndex,
      rowsRendered: items.length
    });

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
