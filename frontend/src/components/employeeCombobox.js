import { searchEmployeeSelector, getCachedEmployees } from '../utils/dataCache.js';

function escapeHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function normalizeName(str) {
  return String(str ?? '').trim().replace(/\s+/g, ' ');
}

function toEmployeeOption(emp) {
  if (!emp) return null;
  if (typeof emp === 'string') {
    const label = normalizeName(emp);
    if (!label) return null;
    return { id: null, name: label, branch_name: '', division: '', status: '' };
  }
  const name = normalizeName(emp.full_name || emp.label || emp.name || '');
  if (!name) return null;
  const parsedId = Number(emp.id ?? emp.value);
  return {
    id: Number.isInteger(parsedId) && parsedId > 0 ? parsedId : null,
    name,
    branch_name: emp.branch_name || emp.branch || '',
    division: emp.division || '',
    status: emp.status || '',
  };
}

function mergeEmployeePools(primary = [], extra = []) {
  const map = new Map();
  for (const raw of [...primary, ...extra]) {
    const opt = toEmployeeOption(raw);
    if (!opt) continue;
    const key = opt.id !== null ? `id:${opt.id}` : `name:${opt.name.toLowerCase()}`;
    if (!map.has(key)) map.set(key, opt);
  }
  return Array.from(map.values());
}

/**
 * Injects modern scoped styles for Employee Combobox to guarantee
 * 100% visual consistency regardless of CDN/browser cache state.
 */
function ensureComboboxStyles() {
  if (typeof document === 'undefined') return;
  if (document.getElementById('fc-emp-combobox-scoped-css')) return;
  const styleEl = document.createElement('style');
  styleEl.id = 'fc-emp-combobox-scoped-css';
  styleEl.textContent = `
    .fc-emp-combobox {
      position: relative;
      width: 100%;
      box-sizing: border-box;
      font-family: inherit;
    }
    .fc-emp-input-wrap {
      position: relative;
      display: flex;
      align-items: center;
      width: 100%;
      box-sizing: border-box;
    }
    .fc-emp-search-icon {
      position: absolute;
      left: 14px;
      top: 50%;
      transform: translateY(-50%);
      display: inline-flex;
      align-items: center;
      justify-content: center;
      color: #94A3B8;
      pointer-events: none;
      z-index: 2;
      transition: color 0.15s ease;
    }
    .fc-emp-input-wrap:focus-within .fc-emp-search-icon {
      color: #2563EB;
    }
    .fc-emp-input {
      width: 100% !important;
      height: 46px !important;
      min-height: 46px !important;
      padding: 0 38px 0 42px !important;
      border: 1px solid #CBD5E1 !important;
      border-radius: 12px !important;
      font-size: 0.95rem !important;
      font-family: inherit !important;
      font-weight: 500 !important;
      color: #0F172A !important;
      background: #FFFFFF !important;
      outline: none !important;
      box-sizing: border-box !important;
      transition: border-color 0.15s ease, box-shadow 0.15s ease !important;
      line-height: normal !important;
    }
    .fc-emp-input::placeholder {
      color: #94A3B8 !important;
      font-weight: 400 !important;
      font-size: 0.92rem !important;
    }
    .fc-emp-input:focus {
      border-color: #2563EB !important;
      box-shadow: 0 0 0 3.5px rgba(37, 99, 235, 0.12) !important;
      background: #FFFFFF !important;
    }
    .fc-emp-clear-btn {
      position: absolute;
      right: 12px;
      top: 50%;
      transform: translateY(-50%);
      width: 22px;
      height: 22px;
      border-radius: 50%;
      border: none;
      background: #E2E8F0;
      color: #64748B;
      cursor: pointer;
      display: none;
      align-items: center;
      justify-content: center;
      font-size: 14px;
      line-height: 1;
      transition: all 0.15s ease;
      z-index: 2;
      padding: 0;
    }
    .fc-emp-clear-btn:hover {
      background: #CBD5E1;
      color: #0F172A;
    }
    .fc-emp-dropdown {
      position: absolute;
      top: calc(100% + 6px);
      left: 0;
      right: 0;
      width: 100%;
      z-index: 1070;
      background: #FFFFFF;
      border: 1px solid #E2E8F0;
      border-radius: 12px;
      box-shadow: 0 12px 28px -4px rgba(15, 23, 42, 0.14), 0 4px 10px -2px rgba(15, 23, 42, 0.05);
      max-height: 260px;
      overflow-y: auto;
      overscroll-behavior: contain;
      display: none;
      padding: 4px 0;
      box-sizing: border-box;
    }
    .fc-emp-dropdown.open {
      display: block;
    }
    .fc-emp-dropdown::-webkit-scrollbar {
      width: 6px;
    }
    .fc-emp-dropdown::-webkit-scrollbar-track {
      background: transparent;
    }
    .fc-emp-dropdown::-webkit-scrollbar-thumb {
      background: #CBD5E1;
      border-radius: 999px;
    }
    .fc-emp-dropdown::-webkit-scrollbar-thumb:hover {
      background: #94A3B8;
    }
    .fc-emp-option {
      padding: 10px 14px;
      cursor: pointer;
      border-bottom: 1px solid #F1F5F9;
      transition: background-color 0.12s ease;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
      box-sizing: border-box;
    }
    .fc-emp-option:last-child {
      border-bottom: none;
    }
    .fc-emp-option:hover,
    .fc-emp-option.active {
      background: #F0F7FF;
    }
    .fc-emp-option.active {
      background: #EFF6FF;
      border-left: 3px solid #2563EB;
      padding-left: 11px;
    }
    .fc-emp-option-main {
      display: flex;
      flex-direction: column;
      gap: 2px;
      flex: 1;
      min-width: 0;
    }
    .fc-emp-option-name {
      font-size: 0.92rem;
      font-weight: 600;
      color: #0F172A;
      line-height: 1.35;
      letter-spacing: -0.01em;
      word-break: break-word;
    }
    .fc-emp-option-meta {
      font-size: 0.78rem;
      color: #64748B;
      line-height: 1.3;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    .fc-emp-option-manual {
      background: #FFFDF5;
      border-top: 1px dashed #FCD34D;
    }
    .fc-emp-option-manual:hover,
    .fc-emp-option-manual.active {
      background: #FEF9C3;
    }
    .fc-emp-option-manual .fc-emp-option-name {
      color: #92400E;
      font-weight: 600;
    }
    .fc-emp-manual-row {
      display: flex;
      align-items: center;
      gap: 6px;
    }
    .fc-emp-manual-icon {
      font-size: 1.1rem;
      font-weight: 700;
      color: #D97706;
      line-height: 1;
    }
    .fc-emp-manual-badge-row {
      margin-top: 4px;
    }
    .fc-emp-empty {
      padding: 14px 16px;
      font-size: 0.85rem;
      color: #64748B;
      text-align: center;
    }
    .fc-emp-status-row {
      margin-top: 6px;
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 6px;
    }
    .fc-emp-badge {
      display: inline-flex;
      align-items: center;
      gap: 5px;
      padding: 4px 10px;
      border-radius: 999px;
      font-size: 0.75rem;
      font-weight: 600;
      line-height: 1.3;
    }
    .fc-emp-badge-master {
      background: #ECFDF5;
      color: #065F46;
      border: 1px solid #A7F3D0;
    }
    .fc-emp-badge-manual {
      background: #FFFBEB;
      color: #92400E;
      border: 1px solid #FCD34D;
    }
    .fc-emp-chips {
      margin-top: 8px;
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
    }
    .fc-emp-chip {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      padding: 4px 10px;
      background: #F1F5F9;
      color: #1E293B;
      border: 1px solid #CBD5E1;
      border-radius: 999px;
      font-size: 0.78rem;
      font-weight: 500;
    }
    .fc-emp-chip.manual {
      background: #FEF3C7;
      color: #92400E;
      border-color: #FCD34D;
    }
    .fc-emp-chip-tag {
      font-size: 0.68rem;
      padding: 1px 5px;
      background: #F59E0B;
      color: #FFFFFF;
      border-radius: 4px;
      font-weight: 700;
      text-transform: uppercase;
    }
    .fc-emp-chip-remove {
      background: none;
      border: none;
      padding: 0;
      font-size: 14px;
      line-height: 1;
      cursor: pointer;
      color: #64748B;
      display: inline-flex;
      align-items: center;
      justify-content: center;
    }
    .fc-emp-chip-remove:hover {
      color: #EF4444;
    }
    @media (max-width: 640px) {
      .fc-emp-input {
        font-size: 16px !important;
        height: 48px !important;
        min-height: 48px !important;
      }
    }
  `;
  document.head.appendChild(styleEl);
}

/**
 * Parses Training participants field (comma-separated or JSON array string) into clean names.
 */
export function parseParticipantList(raw) {
  if (!raw) return [];
  if (Array.isArray(raw)) return raw.map(normalizeName).filter(Boolean);
  const str = String(raw).trim();
  if (!str) return [];
  if (str.startsWith('[')) {
    try {
      const parsed = JSON.parse(str);
      if (Array.isArray(parsed)) return parsed.map(normalizeName).filter(Boolean);
    } catch {
      // Fallback to comma split
    }
  }
  return str.split(',').map(normalizeName).filter(Boolean);
}

/**
 * Mounts the unified FCMS Employee Search Combobox into `mountEl`.
 * Modern search input with icon, clean dropdown, partial matching,
 * keyboard navigation, and explicit manual-name fallback.
 */
export function mountEmployeeCombobox(mountEl, options = {}) {
  if (!mountEl) {
    throw new Error('mountEmployeeCombobox requires a valid mount element');
  }

  ensureComboboxStyles();

  const {
    mode = 'single',
    inputId = '',
    inputName = 'employee_search',
    placeholder = 'Cari nama karyawan...',
    required = false,
    allowManual = true,
    hasEmployeeIdSchema = true,
    manualBadgeText = 'Belum terhubung ke Master Karyawan',
    masterBadgePrefix = 'Terhubung ke Master Karyawan',
    employees: initialPool = [],
    initialEmployeeId = null,
    initialEmployeeName = '',
    initialParticipants = [],
    maxResults = 25,
    onChange = null,
  } = options;

  let localPool = mergeEmployeePools(initialPool);
  let remotePool = [];
  let isOpen = false;
  let activeIndex = -1;
  let searchTimer = null;
  let remoteSeq = 0;

  const parsedInitId = Number(initialEmployeeId);
  const cleanInitId = Number.isInteger(parsedInitId) && parsedInitId > 0 ? parsedInitId : null;
  const cleanInitName = normalizeName(initialEmployeeName);

  const findExactMaster = (name) => {
    const target = normalizeName(name).toLowerCase();
    if (!target) return null;
    return mergeEmployeePools(localPool, remotePool).find(e => e.name.toLowerCase() === target) || null;
  };

  let state = {
    employee_id: cleanInitId,
    employee_name: cleanInitName,
    is_manual: Boolean(cleanInitName && cleanInitId === null && (hasEmployeeIdSchema || !findExactMaster(cleanInitName))),
    selected_employee: cleanInitId !== null
      ? (mergeEmployeePools(localPool).find(e => e.id === cleanInitId) || { id: cleanInitId, name: cleanInitName })
      : null,
    confirmed: Boolean(cleanInitName),
  };

  let multiItems = mode === 'multi'
    ? parseParticipantList(initialParticipants.length ? initialParticipants : cleanInitName).map((name) => {
        const match = findExactMaster(name);
        return match
          ? { employee_id: match.id, employee_name: match.name, is_manual: false }
          : { employee_id: null, employee_name: name, is_manual: true };
      })
    : [];

  mountEl.innerHTML = `
    <div class="fc-emp-combobox" role="combobox" aria-expanded="false" aria-haspopup="listbox">
      <div class="fc-emp-input-wrap">
        <span class="fc-emp-search-icon" aria-hidden="true">
          <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
            <circle cx="11" cy="11" r="8"></circle>
            <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
          </svg>
        </span>
        <input
          type="text"
          ${inputId ? `id="${escapeHtml(inputId)}"` : ''}
          name="${escapeHtml(inputName)}"
          class="fc-emp-input"
          autocomplete="off"
          placeholder="${escapeHtml(placeholder)}"
          ${required && mode === 'single' ? 'required' : ''}
          value="${mode === 'single' ? escapeHtml(cleanInitName) : ''}"
          aria-autocomplete="list"
        />
        <button type="button" class="fc-emp-clear-btn" title="Kosongkan" aria-label="Kosongkan pencarian" style="display:none;">&times;</button>
        <div class="fc-emp-dropdown" role="listbox"></div>
      </div>
      <div class="fc-emp-status-row" aria-live="polite"></div>
      ${mode === 'multi' ? '<div class="fc-emp-chips"></div>' : ''}
    </div>
  `;

  const rootEl = mountEl.querySelector('.fc-emp-combobox');
  const inputEl = mountEl.querySelector('.fc-emp-input');
  const clearBtn = mountEl.querySelector('.fc-emp-clear-btn');
  const dropdownEl = mountEl.querySelector('.fc-emp-dropdown');
  const statusEl = mountEl.querySelector('.fc-emp-status-row');
  const chipsEl = mountEl.querySelector('.fc-emp-chips');

  let renderedItems = [];

  function syncClearBtn() {
    clearBtn.style.display = inputEl.value.trim() ? 'inline-flex' : 'none';
  }

  function emitChange() {
    if (typeof onChange === 'function') {
      onChange(api.getValue());
    }
  }

  function renderStatus() {
    if (mode === 'multi') {
      const manualCount = multiItems.filter(p => p.is_manual).length;
      const typedPending = normalizeName(inputEl.value);
      const pendingIsManual = typedPending && !findExactMaster(typedPending);
      if (multiItems.length === 0 && !typedPending) {
        statusEl.innerHTML = '';
        return;
      }
      if (manualCount > 0 || pendingIsManual) {
        statusEl.innerHTML = `<span class="fc-emp-badge fc-emp-badge-manual">⚠️ Ada peserta nama manual (${escapeHtml(manualBadgeText)})</span>`;
      } else {
        statusEl.innerHTML = `<span class="fc-emp-badge fc-emp-badge-master">✓ Semua peserta terhubung ke Master Karyawan (${multiItems.length})</span>`;
      }
      return;
    }

    if (isOpen || !state.confirmed) {
      statusEl.innerHTML = '';
      return;
    }

    const typed = normalizeName(inputEl.value);
    if (!typed) {
      statusEl.innerHTML = '';
      return;
    }

    if (state.employee_id !== null) {
      const metaBranch = state.selected_employee?.branch_name ? ` • ${escapeHtml(state.selected_employee.branch_name)}` : '';
      statusEl.innerHTML = `<span class="fc-emp-badge fc-emp-badge-master">✓ ${escapeHtml(masterBadgePrefix)} (#${state.employee_id})${metaBranch}</span>`;
      return;
    }

    if (!hasEmployeeIdSchema) {
      const matched = findExactMaster(typed);
      if (matched && !state.is_manual) {
        statusEl.innerHTML = `<span class="fc-emp-badge fc-emp-badge-master">✓ Cocok dengan Master Karyawan${matched.id ? ` (#${matched.id})` : ''}</span>`;
        return;
      }
    }

    statusEl.innerHTML = `<span class="fc-emp-badge fc-emp-badge-manual">⚠️ ${escapeHtml(manualBadgeText)}</span>`;
  }

  function renderChips() {
    if (!chipsEl) return;
    if (multiItems.length === 0) {
      chipsEl.innerHTML = '';
      return;
    }
    chipsEl.innerHTML = multiItems.map((p, idx) => `
      <span class="fc-emp-chip ${p.is_manual ? 'manual' : ''}" data-index="${idx}">
        <span>${escapeHtml(p.employee_name)}</span>
        ${p.is_manual ? '<span class="fc-emp-chip-tag">Manual</span>' : ''}
        <button type="button" class="fc-emp-chip-remove" data-remove-index="${idx}" title="Hapus peserta">&times;</button>
      </span>
    `).join('');
  }

  function buildCandidateList(query) {
    const trimmed = normalizeName(query);
    // Standard Behavior: Minimum 1 character to start search, do not display candidate pool before user types
    if (!trimmed) {
      return { items: [], totalMatched: 0 };
    }

    const q = trimmed.toLowerCase();
    const pool = mergeEmployeePools(localPool, remotePool);
    const matches = pool.filter(e =>
      e.name.toLowerCase().includes(q) ||
      (e.branch_name && e.branch_name.toLowerCase().includes(q)) ||
      (e.division && e.division.toLowerCase().includes(q))
    );

    const sliced = matches.slice(0, maxResults).map(emp => ({
      type: 'master',
      emp,
      label: emp.name,
      meta: [emp.branch_name, emp.division, emp.status].filter(Boolean).join(' · '),
    }));

    const items = [...sliced];
    const hasExactMaster = pool.some(e => e.name.toLowerCase() === q);
    if (allowManual && !hasExactMaster) {
      items.push({
        type: 'manual',
        name: trimmed,
        label: trimmed,
        meta: manualBadgeText,
      });
    }
    return { items, totalMatched: matches.length };
  }

  function positionDropdownWithinModal() {
    dropdownEl.style.top = '';
    dropdownEl.style.bottom = '';
    const rect = inputEl.getBoundingClientRect();
    const spaceBelow = window.innerHeight - rect.bottom;
    const spaceAbove = rect.top;
    if (spaceBelow < 220 && spaceAbove > spaceBelow) {
      dropdownEl.style.top = 'auto';
      dropdownEl.style.bottom = 'calc(100% + 6px)';
    } else {
      dropdownEl.style.top = 'calc(100% + 6px)';
      dropdownEl.style.bottom = 'auto';
    }
  }

  function renderDropdown() {
    const query = inputEl.value;
    const trimmed = normalizeName(query);
    if (!trimmed) {
      dropdownEl.innerHTML = '';
      closeDropdown();
      return;
    }

    const { items, totalMatched } = buildCandidateList(query);
    renderedItems = items;

    if (items.length === 0) {
      dropdownEl.innerHTML = `<div class="fc-emp-empty">Tidak ada karyawan yang cocok dengan "${escapeHtml(trimmed)}".</div>`;
    } else {
      const hintHtml = (totalMatched > maxResults)
        ? `<div class="fc-emp-empty" style="border-top:1px solid #F1F5F9; font-size:0.75rem; color:#94A3B8;">Menampilkan ${maxResults} dari ${totalMatched} karyawan. Ketik lebih spesifik untuk menyaring.</div>`
        : '';
      dropdownEl.innerHTML = items.map((opt, idx) => {
        const isActive = idx === activeIndex ? 'active' : '';
        if (opt.type === 'manual') {
          return `
            <div class="fc-emp-option fc-emp-option-manual ${isActive}" role="option" data-idx="${idx}" data-manual="1" aria-selected="${idx === activeIndex}">
              <div class="fc-emp-option-main">
                <div class="fc-emp-manual-row">
                  <span class="fc-emp-manual-icon">+</span>
                  <div class="fc-emp-option-name">Gunakan nama manual: <strong>${escapeHtml(opt.name)}</strong></div>
                </div>
                <div class="fc-emp-manual-badge-row">
                  <span class="fc-emp-badge fc-emp-badge-manual">⚠️ ${escapeHtml(opt.meta)}</span>
                </div>
              </div>
            </div>
          `;
        }
        return `
          <div class="fc-emp-option ${isActive}" role="option" data-idx="${idx}" data-emp-id="${opt.emp.id ?? ''}" aria-selected="${idx === activeIndex}">
            <div class="fc-emp-option-main">
              <div class="fc-emp-option-name">${escapeHtml(opt.label)}</div>
              ${opt.meta ? `<div class="fc-emp-option-meta">${escapeHtml(opt.meta)}</div>` : ''}
            </div>
          </div>
        `;
      }).join('') + hintHtml;
    }

    positionDropdownWithinModal();
  }

  function openDropdown() {
    const trimmed = normalizeName(inputEl.value);
    if (!trimmed) return; // Do not open if empty
    if (!isOpen) {
      isOpen = true;
      dropdownEl.classList.add('open');
      rootEl.setAttribute('aria-expanded', 'true');
    }
    renderDropdown();
  }

  function closeDropdown() {
    isOpen = false;
    activeIndex = -1;
    dropdownEl.classList.remove('open');
    rootEl.setAttribute('aria-expanded', 'false');
    const trimmed = normalizeName(inputEl.value);
    if (trimmed && mode === 'single' && !state.selected_employee) {
      state.confirmed = true;
    }
    renderStatus();
  }

  function selectOption(opt) {
    if (!opt) return;
    if (mode === 'multi') {
      const entry = opt.type === 'manual'
        ? { employee_id: null, employee_name: opt.name, is_manual: true }
        : { employee_id: opt.emp.id, employee_name: opt.emp.name, is_manual: false };
      if (!multiItems.some(p => p.employee_name.toLowerCase() === entry.employee_name.toLowerCase())) {
        multiItems.push(entry);
      }
      inputEl.value = '';
      syncClearBtn();
      renderChips();
      renderStatus();
      closeDropdown();
      emitChange();
      return;
    }

    if (opt.type === 'manual') {
      state = {
        employee_id: null,
        employee_name: opt.name,
        is_manual: true,
        selected_employee: null,
        confirmed: true,
      };
      inputEl.value = opt.name;
    } else {
      state = {
        employee_id: opt.emp.id,
        employee_name: opt.emp.name,
        is_manual: false,
        selected_employee: opt.emp,
        confirmed: true,
      };
      inputEl.value = opt.emp.name;
    }

    syncClearBtn();
    renderStatus();
    closeDropdown();
    emitChange();
  }

  function triggerRemoteSearch(query) {
    clearTimeout(searchTimer);
    const trimmed = normalizeName(query);
    if (!trimmed) return;
    const seq = ++remoteSeq;
    searchTimer = setTimeout(async () => {
      try {
        const fetched = await searchEmployeeSelector({ search: trimmed, limit: maxResults });
        if (seq !== remoteSeq) return;
        if (Array.isArray(fetched) && fetched.length > 0) {
          remotePool = mergeEmployeePools(remotePool, fetched);
          if (isOpen) renderDropdown();
        }
      } catch {
        // Fallback smoothly to cached employee pool when offline
      }
    }, 180);
  }

  // Hydrate pool lazily if caller passed an empty pool
  if (localPool.length === 0) {
    getCachedEmployees()
      .then((list) => {
        if (Array.isArray(list) && list.length > 0) {
          localPool = mergeEmployeePools(list);
          if (state.employee_id !== null && !state.selected_employee) {
            state.selected_employee = localPool.find(e => e.id === state.employee_id) || state.selected_employee;
          }
          renderStatus();
          if (isOpen) renderDropdown();
        }
      })
      .catch(() => {});
  }

  inputEl.addEventListener('focus', () => {
    const trimmed = normalizeName(inputEl.value);
    if (trimmed.length >= 1) {
      activeIndex = -1;
      openDropdown();
    }
  });

  inputEl.addEventListener('input', () => {
    const val = inputEl.value;
    const trimmed = normalizeName(val);
    if (mode === 'single') {
      // Editing text after selecting a Master row MUST immediately detach stale employee_id
      if (!trimmed) {
        state = { employee_id: null, employee_name: '', is_manual: false, selected_employee: null };
      } else {
        state = {
          employee_id: null,
          employee_name: trimmed,
          is_manual: allowManual,
          selected_employee: null,
        };
      }
    }
    activeIndex = -1;
    syncClearBtn();
    renderStatus();
    if (trimmed.length < 1) {
      closeDropdown();
      emitChange();
      return;
    }
    openDropdown();
    triggerRemoteSearch(trimmed);
    emitChange();
  });

  inputEl.addEventListener('keydown', (e) => {
    const trimmed = normalizeName(inputEl.value);
    if (!isOpen && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) {
      if (trimmed.length >= 1) openDropdown();
    }
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      if (renderedItems.length === 0) return;
      activeIndex = (activeIndex + 1) % renderedItems.length;
      renderDropdown();
      dropdownEl.querySelector('.fc-emp-option.active')?.scrollIntoView({ block: 'nearest' });
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      if (renderedItems.length === 0) return;
      activeIndex = activeIndex <= 0 ? renderedItems.length - 1 : activeIndex - 1;
      renderDropdown();
      dropdownEl.querySelector('.fc-emp-option.active')?.scrollIntoView({ block: 'nearest' });
    } else if (e.key === 'Enter') {
      if (isOpen && renderedItems.length > 0) {
        e.preventDefault();
        const picked = activeIndex >= 0 ? renderedItems[activeIndex] : renderedItems[0];
        if (picked) selectOption(picked);
      } else if (mode === 'multi' && normalizeName(inputEl.value)) {
        e.preventDefault();
        const raw = normalizeName(inputEl.value);
        const exact = findExactMaster(raw);
        selectOption(exact ? { type: 'master', emp: exact } : { type: 'manual', name: raw });
      }
    } else if (e.key === 'Escape') {
      if (isOpen) {
        e.preventDefault();
        e.stopPropagation();
        closeDropdown();
      }
    }
  });

  dropdownEl.addEventListener('mousedown', (e) => {
    // Prevent input blur before click registers
    e.preventDefault();
  });

  dropdownEl.addEventListener('click', (e) => {
    const optEl = e.target.closest('.fc-emp-option');
    if (!optEl) return;
    const idx = Number(optEl.dataset.idx);
    if (Number.isInteger(idx) && renderedItems[idx]) {
      selectOption(renderedItems[idx]);
    }
  });

  clearBtn.addEventListener('click', () => {
    inputEl.value = '';
    if (mode === 'single') {
      state = { employee_id: null, employee_name: '', is_manual: false, selected_employee: null, confirmed: false };
    }
    syncClearBtn();
    renderStatus();
    closeDropdown();
    inputEl.focus();
    emitChange();
  });

  if (chipsEl) {
    chipsEl.addEventListener('click', (e) => {
      const btn = e.target.closest('[data-remove-index]');
      if (!btn) return;
      const idx = Number(btn.dataset.removeIndex);
      if (Number.isInteger(idx) && idx >= 0 && idx < multiItems.length) {
        multiItems.splice(idx, 1);
        renderChips();
        renderStatus();
        emitChange();
      }
    });
  }

  const handleOutsideMouseDown = (e) => {
    if (!mountEl.isConnected) {
      api.destroy();
      return;
    }
    if (!mountEl.contains(e.target)) {
      closeDropdown();
    }
  };
  document.addEventListener('mousedown', handleOutsideMouseDown);

  syncClearBtn();
  renderChips();
  renderStatus();

  const api = {
    inputEl,
    getValue() {
      if (mode === 'multi') {
        const pending = normalizeName(inputEl.value);
        const list = [...multiItems];
        if (pending && !list.some(p => p.employee_name.toLowerCase() === pending.toLowerCase())) {
          const match = findExactMaster(pending);
          list.push(match
            ? { employee_id: match.id, employee_name: match.name, is_manual: false }
            : { employee_id: null, employee_name: pending, is_manual: true });
        }
        return {
          participants: list.map(p => p.employee_name),
          participantsString: list.map(p => p.employee_name).join(', '),
          items: list,
          hasManual: list.some(p => p.is_manual),
        };
      }

      const currentName = normalizeName(inputEl.value);
      if (!currentName) {
        return { employee_id: null, employee_name: '', is_manual: false, selected_employee: null };
      }
      if (state.employee_id !== null && state.employee_name.toLowerCase() === currentName.toLowerCase()) {
        return {
          employee_id: state.employee_id,
          employee_name: state.employee_name,
          is_manual: false,
          selected_employee: state.selected_employee,
        };
      }
      if (!hasEmployeeIdSchema) {
        const exact = findExactMaster(currentName);
        if (exact && !state.is_manual) {
          return { employee_id: exact.id, employee_name: exact.name, is_manual: false, selected_employee: exact };
        }
      }
      return {
        employee_id: null,
        employee_name: currentName,
        is_manual: true,
        selected_employee: null,
      };
    },
    setValue({ employee_id = null, employee_name = '' } = {}) {
      const cleanName = normalizeName(employee_name);
      const parsedId = Number(employee_id);
      const validId = Number.isInteger(parsedId) && parsedId > 0 ? parsedId : null;
      const pool = mergeEmployeePools(localPool, remotePool);
      state = {
        employee_id: validId,
        employee_name: cleanName,
        is_manual: Boolean(cleanName && validId === null),
        selected_employee: validId !== null ? (pool.find(e => e.id === validId) || { id: validId, name: cleanName }) : null,
        confirmed: Boolean(cleanName),
      };
      inputEl.value = cleanName;
      syncClearBtn();
      renderStatus();
    },
    selectOptionByIdOrName(idOrName) {
      const pool = mergeEmployeePools(localPool, remotePool);
      const match = pool.find(e =>
        (idOrName !== null && idOrName !== undefined && String(e.id) === String(idOrName)) ||
        e.name.toLowerCase() === normalizeName(idOrName).toLowerCase()
      );
      if (match) {
        selectOption({ type: 'master', emp: match });
        return true;
      }
      return false;
    },
    destroy() {
      clearTimeout(searchTimer);
      document.removeEventListener('mousedown', handleOutsideMouseDown);
    },
  };

  return api;
}
