import { confirmDialog } from './modal.js';
import { escapeHtml, sanitizeHtmlFragment } from '../utils/sanitize.js';

export const MAX_BULK_IDS = 100;

/**
 * Attempts to toggle a single row ID in a Set without exceeding maxLimit.
 * Existing selections remain untouched when rejected.
 */
export function toggleBulkSelectionItem(selectedIds, id, checked, maxLimit = MAX_BULK_IDS) {
  if (!checked) {
    selectedIds.delete(id);
    return { accepted: true, count: selectedIds.size };
  }
  if (selectedIds.has(id)) {
    return { accepted: true, count: selectedIds.size };
  }
  if (selectedIds.size >= maxLimit) {
    return {
      accepted: false,
      count: selectedIds.size,
      reason: `Maksimum ${maxLimit} item dapat dipilih untuk penghapusan bulk`
    };
  }
  selectedIds.add(id);
  return { accepted: true, count: selectedIds.size };
}

/**
 * Attempts to select or deselect all rows on the current page.
 * Rejects without truncating or clearing existing selections if total would exceed maxLimit.
 */
export function toggleBulkSelectionPage(selectedIds, pageRows, checked, maxLimit = MAX_BULK_IDS) {
  const ids = (pageRows || [])
    .map(r => (typeof r === 'object' && r !== null ? r.id : r))
    .filter(id => id !== undefined && id !== null);
  if (!checked) {
    ids.forEach(id => selectedIds.delete(id));
    return { accepted: true, count: selectedIds.size };
  }
  const newIds = ids.filter(id => !selectedIds.has(id));
  if (selectedIds.size + newIds.length > maxLimit) {
    return {
      accepted: false,
      count: selectedIds.size,
      reason: `Maksimum ${maxLimit} item dapat dipilih untuk penghapusan bulk (saat ini ${selectedIds.size} item sudah dipilih)`
    };
  }
  newIds.forEach(id => selectedIds.add(id));
  return { accepted: true, count: selectedIds.size };
}

/**
 * Build a Clean Navy mobile vertical data card for responsive viewports (<= 767px).
 */
export function buildMobileDataCard({
  title,
  headerRight = '',
  badges = [],
  meta = [],
  notes = '',
  notesLabel = 'Catatan',
  docLink = '',
  row,
  handlers = {},
}) {
  const card = document.createElement('article');
  card.className = 'fc-mobile-card';

  const headHtml = `
    <div class="fc-mobile-card-head">
      <div class="fc-mobile-card-title">${escapeHtml(title || '-')}</div>
      ${headerRight ? `<div class="fc-mobile-card-right">${sanitizeHtmlFragment(headerRight)}</div>` : ''}
    </div>
  `;

  const validBadges = (badges || []).filter(Boolean);
  const badgesHtml = validBadges.length > 0
    ? `<div class="fc-mobile-card-badges">${validBadges.map(b => sanitizeHtmlFragment(b)).join('')}</div>`
    : '';

  const validMeta = (meta || []).filter(m => m && (m.html || (m.value !== undefined && m.value !== null && String(m.value).trim() !== '' && String(m.value).trim() !== '-')));
  const metaHtml = validMeta.length > 0
    ? `
      <div class="fc-mobile-card-meta">
        ${validMeta.map(m => `
          <div class="fc-mobile-meta-item${m.full ? ' fc-mobile-meta-item--full' : ''}">
            <span class="fc-mobile-meta-label">${escapeHtml(m.label)}</span>
            <span class="fc-mobile-meta-value">${m.html ? sanitizeHtmlFragment(m.html) : escapeHtml(String(m.value))}</span>
          </div>
        `).join('')}
      </div>
    `
    : '';

  const notesStr = notes !== undefined && notes !== null ? String(notes).trim() : '';
  const notesHtml = notesStr && notesStr !== '-'
    ? `
      <div class="fc-mobile-card-notes">
        <span class="fc-mobile-meta-label">${escapeHtml(notesLabel)}</span>
        <div class="fc-mobile-notes-text">${escapeHtml(notesStr)}</div>
      </div>
    `
    : '';

  const safeDocHref = (typeof docLink === 'string' && /^(https?:\/\/|\/)/i.test(docLink.trim())) ? docLink.trim() : '';
  card.innerHTML = headHtml + badgesHtml + metaHtml + notesHtml;

  const { onView, onEdit, onDelete, actions = [] } = handlers || {};
  const hasActions = !!(safeDocHref || onView || onEdit || onDelete || actions.length > 0);
  if (hasActions) {
    const footer = document.createElement('div');
    footer.className = 'fc-mobile-card-footer';

    if (safeDocHref) {
      const docA = document.createElement('a');
      docA.href = safeDocHref;
      docA.target = '_blank';
      docA.rel = 'noopener';
      docA.className = 'fc-mobile-doc-btn';
      docA.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg><span>Buka Dokumen</span>`;
      footer.appendChild(docA);
    }

    const actionsWrap = document.createElement('div');
    actionsWrap.className = 'fc-mobile-card-actions';

    if (onView) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'fc-mobile-action-btn fc-action-view';
      btn.setAttribute('aria-label', 'Lihat detail');
      btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg><span>Detail</span>`;
      btn.addEventListener('click', () => onView(row));
      actionsWrap.appendChild(btn);
    }
    if (onEdit) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'fc-mobile-action-btn fc-action-edit';
      btn.setAttribute('aria-label', 'Edit data');
      btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg><span>Edit</span>`;
      btn.addEventListener('click', () => onEdit(row));
      actionsWrap.appendChild(btn);
    }
    actions.forEach(action => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = `fc-mobile-action-btn ${action.class || ''}`;
      btn.setAttribute('aria-label', action.label || 'Aksi');
      btn.innerHTML = sanitizeHtmlFragment(action.icon ? `${action.icon}<span>${escapeHtml(action.label || '')}</span>` : escapeHtml(action.label || 'Aksi'));
      btn.addEventListener('click', () => action.handler(row));
      actionsWrap.appendChild(btn);
    });
    if (onDelete) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'fc-mobile-action-btn fc-action-del';
      btn.setAttribute('aria-label', 'Hapus data');
      btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg><span>Hapus</span>`;
      btn.addEventListener('click', () => onDelete(row));
      actionsWrap.appendChild(btn);
    }

    if (actionsWrap.childElementCount > 0) {
      footer.appendChild(actionsWrap);
    }
    card.appendChild(footer);
  }

  return card;
}

// Reusable data table with pagination
export function createTable({ columns, data, onEdit, onDelete, onView, actions = [], emptyText = 'Tidak ada data', bulkSelect = null, mobileCardRenderer = null }) {
  const wrapper = document.createElement('div');
  wrapper.className = mobileCardRenderer ? 'table-wrapper fc-desktop-only-table' : 'table-wrapper';

  if (!data || data.length === 0) {
    wrapper.className = 'table-wrapper';
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
    const maxSelect = bulkSelect.maxSelect || Infinity;
    selectAll.addEventListener('change', () => {
      if (Number.isFinite(maxSelect)) {
        const res = toggleBulkSelectionPage(bulkSelect.selectedIds, data, selectAll.checked, maxSelect);
        if (!res.accepted) {
          selectAll.checked = false;
          if (typeof bulkSelect.onLimitExceeded === 'function') bulkSelect.onLimitExceeded(res.reason);
        }
        const cbs = wrapper.querySelectorAll('.row-checkbox');
        data.forEach((row, idx) => {
          if (cbs[idx]) cbs[idx].checked = bulkSelect.selectedIds.has(row.id);
        });
      } else {
        data.forEach(row => {
          if (selectAll.checked) bulkSelect.selectedIds.add(row.id);
          else bulkSelect.selectedIds.delete(row.id);
        });
        wrapper.querySelectorAll('.row-checkbox').forEach(cb => cb.checked = selectAll.checked);
      }
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
  const maxSelect = bulkSelect?.maxSelect || Infinity;
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
        if (Number.isFinite(maxSelect)) {
          const res = toggleBulkSelectionItem(bulkSelect.selectedIds, row.id, cb.checked, maxSelect);
          if (!res.accepted) {
            cb.checked = false;
            if (typeof bulkSelect.onLimitExceeded === 'function') bulkSelect.onLimitExceeded(res.reason);
          } else if (!cb.checked) {
            const selectAll = document.getElementById('select-all-checkbox');
            if (selectAll) selectAll.checked = false;
          }
        } else {
          if (cb.checked) bulkSelect.selectedIds.add(row.id);
          else {
            bulkSelect.selectedIds.delete(row.id);
            // uncheck select-all if any deselected
            const selectAll = document.getElementById('select-all-checkbox');
            if (selectAll) selectAll.checked = false;
          }
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
        else td.innerHTML = rendered ? sanitizeHtmlFragment(rendered) : '';
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
        btn.type = 'button';
        btn.className = 'fc-action-btn fc-action-view';
        btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>`;
        btn.title = 'Lihat';
        btn.setAttribute('aria-label', 'Lihat detail');
        btn.addEventListener('click', () => onView(row));
        btnGroup.appendChild(btn);
      }
      if (onEdit) {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'fc-action-btn fc-action-edit';
        btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg>`;
        btn.title = 'Edit';
        btn.setAttribute('aria-label', 'Edit data');
        btn.addEventListener('click', () => onEdit(row));
        btnGroup.appendChild(btn);
      }
      actions.forEach(action => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = `btn btn-xs ${action.class || 'btn-ghost'}`;
        btn.innerHTML = action.icon || action.label;
        btn.title = action.label;
        btn.setAttribute('aria-label', action.label || 'Aksi');
        btn.addEventListener('click', () => action.handler(row));
        btnGroup.appendChild(btn);
      });
      if (onDelete) {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'fc-action-btn fc-action-del';
        btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>`;
        btn.title = 'Hapus';
        btn.setAttribute('aria-label', 'Hapus data');
        btn.addEventListener('click', () => onDelete(row));
        btnGroup.appendChild(btn);
      }
      td.appendChild(btnGroup);
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  wrapper.appendChild(table);

  if (typeof mobileCardRenderer === 'function') {
    const host = document.createElement('div');
    host.className = 'fc-table-responsive-host';
    host.appendChild(wrapper);

    const cardsList = document.createElement('div');
    cardsList.className = 'fc-mobile-cards-list';
    const handlers = { onEdit, onDelete, onView, actions };
    data.forEach(row => {
      const cardEl = mobileCardRenderer(row, handlers);
      if (cardEl instanceof HTMLElement) {
        cardsList.appendChild(cardEl);
      }
    });
    host.appendChild(cardsList);
    return host;
  }

  return wrapper;
}

export function createPagination({ page, pages, total, limit, onPage }) {
  if (!pages || pages <= 1) return null;
  const safePage = Math.min(Math.max(1, Number(page) || 1), pages);
  const safeLimit = Math.max(1, Number(limit) || 20);
  const safeTotal = Math.max(0, Number(total) || 0);
  const nav = document.createElement('div');
  nav.className = 'pagination';

  const info = document.createElement('span');
  info.className = 'pagination-info';
  const fromIdx = safeTotal > 0 ? ((safePage - 1) * safeLimit) + 1 : 0;
  const toIdx = safeTotal > 0 ? Math.min(safeTotal, safePage * safeLimit) : 0;
  info.textContent = safeTotal > 0
    ? `Hal. ${safePage} / ${pages} • ${fromIdx}-${toIdx} dari ${safeTotal} data`
    : `Total: ${safeTotal} data`;
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

  addBtn('«', 1, safePage === 1);
  addBtn('‹', safePage - 1, safePage === 1);

  const start = Math.max(1, safePage - 2);
  const end = Math.min(pages, safePage + 2);
  for (let i = start; i <= end; i++) addBtn(i, i, false, i === safePage);

  addBtn('›', safePage + 1, safePage === pages);
  addBtn('»', pages, safePage === pages);

  nav.appendChild(btns);
  return nav;
}
