// Reusable data table with pagination
export function createTable({ columns, data, onEdit, onDelete, onView, actions = [], emptyText = 'Tidak ada data', bulkSelect = null }) {
  const wrapper = document.createElement('div');
  wrapper.className = 'table-wrapper';

  if (!data || data.length === 0) {
    wrapper.innerHTML = `<div class="empty-state"><p>${emptyText}</p></div>`;
    return wrapper;
  }

  const table = document.createElement('table');
  table.className = 'data-table';

  // Header
  const thead = document.createElement('thead');
  const headerRow = document.createElement('tr');

  // Checkbox "select all" header cell
  if (bulkSelect) {
    const thCheck = document.createElement('th');
    thCheck.style.width = '40px';
    thCheck.style.textAlign = 'center';
    const selectAll = document.createElement('input');
    selectAll.type = 'checkbox';
    selectAll.id = 'select-all-checkbox';
    selectAll.title = 'Pilih semua';
    selectAll.addEventListener('change', () => {
      data.forEach(row => {
        if (selectAll.checked) bulkSelect.selectedIds.add(row.id);
        else bulkSelect.selectedIds.delete(row.id);
      });
      wrapper.querySelectorAll('.row-checkbox').forEach(cb => cb.checked = selectAll.checked);
      bulkSelect.onToggle();
    });
    thCheck.appendChild(selectAll);
    headerRow.appendChild(thCheck);
  }

  columns.forEach(col => {
    const th = document.createElement('th');
    th.textContent = col.label;
    if (col.width) th.style.width = col.width;
    headerRow.appendChild(th);
  });
  if (onEdit || onDelete || onView || actions.length > 0) {
    const th = document.createElement('th');
    th.textContent = 'Aksi';
    th.style.width = '120px';
    headerRow.appendChild(th);
  }
  thead.appendChild(headerRow);
  table.appendChild(thead);

  // Body
  const tbody = document.createElement('tbody');
  data.forEach(row => {
    const tr = document.createElement('tr');

    // Checkbox cell per row
    if (bulkSelect) {
      const tdCheck = document.createElement('td');
      tdCheck.style.textAlign = 'center';
      tdCheck.style.width = '40px';
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.className = 'row-checkbox';
      cb.checked = bulkSelect.selectedIds.has(row.id);
      cb.addEventListener('change', () => {
        if (cb.checked) bulkSelect.selectedIds.add(row.id);
        else {
          bulkSelect.selectedIds.delete(row.id);
          // uncheck select-all if any deselected
          const selectAll = document.getElementById('select-all-checkbox');
          if (selectAll) selectAll.checked = false;
        }
        bulkSelect.onToggle();
      });
      tdCheck.appendChild(cb);
      tr.appendChild(tdCheck);
    }

    columns.forEach(col => {
      const td = document.createElement('td');
      if (col.render) {
        const rendered = col.render(row[col.key], row);
        if (rendered instanceof HTMLElement) td.appendChild(rendered);
        else td.innerHTML = rendered || '';
      } else {
        td.textContent = row[col.key] !== null && row[col.key] !== undefined && row[col.key] !== '' ? row[col.key] : '';
      }
      if (col.nowrap) td.style.whiteSpace = 'nowrap';
      tr.appendChild(td);
    });

    // Actions column
    if (onEdit || onDelete || onView || actions.length > 0) {
      const td = document.createElement('td');
      td.className = 'actions-cell';
      const btnGroup = document.createElement('div');
      btnGroup.className = 'btn-group';

      if (onView) {
        const btn = document.createElement('button');
        btn.className = 'btn btn-xs btn-ghost';
        btn.innerHTML = '👁';
        btn.title = 'Lihat';
        btn.addEventListener('click', () => onView(row));
        btnGroup.appendChild(btn);
      }
      if (onEdit) {
        const btn = document.createElement('button');
        btn.className = 'btn btn-xs btn-secondary';
        btn.innerHTML = '✏️';
        btn.title = 'Edit';
        btn.addEventListener('click', () => onEdit(row));
        btnGroup.appendChild(btn);
      }
      actions.forEach(action => {
        const btn = document.createElement('button');
        btn.className = `btn btn-xs ${action.class || 'btn-ghost'}`;
        btn.innerHTML = action.icon || action.label;
        btn.title = action.label;
        btn.addEventListener('click', () => action.handler(row));
        btnGroup.appendChild(btn);
      });
      // Individual delete button removed globally as per user request
      td.appendChild(btnGroup);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  wrapper.appendChild(table);
  return wrapper;
}

export function createPagination({ page, pages, total, limit, onPage }) {
  if (pages <= 1) return null;
  const nav = document.createElement('div');
  nav.className = 'pagination';

  const info = document.createElement('span');
  info.className = 'pagination-info';
  info.textContent = `Total: ${total} data`;
  nav.appendChild(info);

  const btns = document.createElement('div');
  btns.className = 'pagination-btns';

  const addBtn = (label, targetPage, disabled = false, active = false) => {
    const btn = document.createElement('button');
    btn.className = `btn btn-sm ${active ? 'btn-primary' : 'btn-ghost'} pagination-btn`;
    btn.textContent = label;
    btn.disabled = disabled;
    btn.addEventListener('click', () => onPage(targetPage));
    btns.appendChild(btn);
  };

  addBtn('«', 1, page === 1);
  addBtn('‹', page - 1, page === 1);

  const start = Math.max(1, page - 2);
  const end = Math.min(pages, page + 2);
  for (let i = start; i <= end; i++) addBtn(i, i, false, i === page);

  addBtn('›', page + 1, page === pages);
  addBtn('»', pages, page === pages);

  nav.appendChild(btns);
  return nav;
}
