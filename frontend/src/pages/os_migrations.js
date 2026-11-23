import { apiFetch, getUser, hasPermission } from '../config.js';
import { escapeHtml } from '../utils/sanitize.js';
import { createModal, confirmDialog } from '../components/modal.js';
import { toastSuccess, toastError } from '../components/toast.js';
import { downloadExcel } from '../utils/excel.js';
import { fetchExportRows } from '../utils/dataCache.js';
import { sortYearMonthDesc } from '../utils/sorting.js';

const MONTH_NAMES = {
  '2026-09': 'September 2026',
  '2026-10': 'Oktober 2026',
  '2026-11': 'November 2026',
  '2026-12': 'Desember 2026'
};

const STATUS_CONFIG = {
  scheduled:   { label: 'Terjadwal',     badgeClass: 'badge-warning', icon: '●', dot: '#FBBF24', bg: 'rgba(245, 158, 11, 0.14)', color: '#FCD34D', border: 'rgba(251, 191, 36, 0.36)' },
  in_progress: { label: 'Sedang Proses', badgeClass: 'badge-info',    icon: '●', dot: '#60A5FA', bg: 'rgba(37, 99, 235, 0.18)',  color: '#93C5FD', border: 'rgba(59, 130, 246, 0.42)' },
  completed:   { label: 'Selesai',       badgeClass: 'badge-success', icon: '●', dot: '#34D399', bg: 'rgba(16, 185, 129, 0.16)', color: '#6EE7B7', border: 'rgba(52, 211, 153, 0.38)' }
};

function renderMigrationActionButtons(r, canEdit, canDelete) {
  if (!canEdit && !canDelete) return '';
  const id = escapeHtml(String(r.id ?? ''));
  const name = escapeHtml(r.branch_name || 'Cabang');
  return `
    <div class="os-mig-actions" style="display:inline-flex; gap:5px; align-items:center;">
      ${canEdit ? `
        <button type="button" class="os-mig-icon-btn edit-card-btn" data-id="${id}" title="Edit ${name}" aria-label="Edit jadwal migrasi ${name}">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <path d="M12 20h9"></path>
            <path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path>
          </svg>
        </button>
      ` : ''}
      ${canDelete ? `
        <button type="button" class="os-mig-icon-btn os-mig-icon-del delete-card-btn" data-id="${id}" data-name="${name}" title="Hapus ${name}" aria-label="Hapus jadwal migrasi ${name}">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <polyline points="3 6 5 6 21 6"></polyline>
            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
            <line x1="10" y1="11" x2="10" y2="17"></line>
            <line x1="14" y1="11" x2="14" y2="17"></line>
          </svg>
        </button>
      ` : ''}
    </div>
  `;
}

export async function renderOSMigrations(container) {
  const user = getUser();
  const canEdit = hasPermission(user, 'os_migrations', 'write');
  const canDelete = hasPermission(user, 'os_migrations', 'delete');
  const canExport = hasPermission(user, 'export');

  container.innerHTML = `
    <div class="page-header" style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px; margin-bottom:20px;">
      <div>
        <h1 class="page-title" style="display:flex; align-items:center; gap:8px; margin:0; font-size:1.45rem; font-weight:800; color:var(--text-1); letter-spacing:-0.02em;">
          <span>🔄</span> Time Line Migrasi OS
        </h1>
      </div>
      <div class="os-header-actions" style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
        <div class="btn-group" style="background:#0E182B; padding:3px; border-radius:10px; display:inline-flex; align-items:center; border:1px solid #223456; height:38px; box-sizing:border-box;">
          <button id="view-board-btn" class="btn btn-sm" style="background:#2563EB; color:#fff; font-weight:600; border-radius:7px; padding:0 12px; height:100%; display:inline-flex; align-items:center; cursor:pointer;">
            📋 Daftar per Bulan
          </button>
          <button id="view-table-btn" class="btn btn-sm btn-ghost" style="font-weight:600; border-radius:7px; padding:0 12px; height:100%; display:inline-flex; align-items:center; cursor:pointer; color:#CBD5E1;">
            📊 Tabel Data
          </button>
        </div>
        ${canExport ? `
        <button id="btn-export-migrations" class="btn btn-outline" style="background:#131F36; border:1px solid #223456; color:#F8FAFC; font-weight:600; height:38px; padding:0 14px; border-radius:10px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; box-sizing:border-box;">
          📥 Export Excel
        </button>
        ` : ''}
        ${canEdit ? `
        <button id="btn-add-migration" class="btn btn-primary" style="font-weight:700; height:38px; padding:0 16px; border-radius:10px; display:inline-flex; align-items:center; gap:6px; cursor:pointer; box-sizing:border-box; background:#2563EB;">
          + Tambah Cabang
        </button>
        ` : ''}
      </div>
    </div>

    <!-- Summary KPI Cards -->
    <div id="migration-summary-cards" style="display:grid; grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); gap:16px; margin-bottom:20px;">
      <div class="kpi-card skeleton" style="height:100px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:100px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:100px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:100px; border-radius:12px;"></div>
    </div>

    <!-- Filter Toolbar -->
    <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#131F36; border:1px solid #223456; border-radius:12px;">
      <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
        <div style="display:flex; flex-wrap:wrap; gap:10px; align-items:center; flex:1; min-width:260px;">
          <div style="position:relative; flex:1; min-width:200px; max-width:320px;">
            <input type="text" id="migration-search" class="form-control" placeholder="Cari cabang, vendor, catatan..." style="padding-left:34px; height:38px; border-radius:8px; font-size:0.88rem; width:100%; background:#0E182B; border-color:#223456;">
            <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:#94A3B8; font-size:0.9rem;">🔍</span>
          </div>
          <select id="filter-vendor" class="form-control" style="width:auto; height:38px; border-radius:8px; font-size:0.875rem; min-width:136px; background:#0F1A2E; border-color:#223456;">
            <option value="">Semua Vendor</option>
            <option value="AEON">AEON</option>
            <option value="Triple S">Triple S</option>
          </select>
          <select id="filter-month" class="form-control" style="width:auto; height:38px; border-radius:8px; font-size:0.875rem; min-width:146px; background:#0F1A2E; border-color:#223456;">
            <option value="">Semua Bulan</option>
            <option value="2026-09">September 2026</option>
            <option value="2026-10">Oktober 2026</option>
            <option value="2026-11">November 2026</option>
            <option value="2026-12">Desember 2026</option>
          </select>
          <select id="filter-status" class="form-control" style="width:auto; height:38px; border-radius:8px; font-size:0.875rem; min-width:138px; background:#0F1A2E; border-color:#223456;">
            <option value="">Semua Status</option>
            <option value="scheduled">Terjadwal</option>
            <option value="in_progress">Sedang Proses</option>
            <option value="completed">Selesai</option>
          </select>
        </div>
        <div id="filter-stats-text" style="font-size:0.85rem; color:#CBD5E1; font-weight:600;">
          Memuat data...
        </div>
      </div>
    </div>

    <style>
      .os-mig-row {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
        padding: 12px 18px;
        border-bottom: 1px solid #1E2E4A;
        transition: background-color 0.14s ease;
      }
      .os-mig-row:last-child { border-bottom: none; }
      .os-mig-row:hover { background: #172642; }
      .os-mig-left {
        display: flex;
        align-items: center;
        gap: 10px;
        min-width: 0;
        flex: 1 1 260px;
      }
      .os-mig-code {
        font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
        font-size: 0.78rem;
        font-weight: 700;
        color: #93C5FD;
        background: rgba(37, 99, 235, 0.14);
        border: 1px solid rgba(59, 130, 246, 0.28);
        padding: 3px 8px;
        border-radius: 6px;
        flex-shrink: 0;
      }
      .os-mig-name {
        font-size: 0.96rem;
        font-weight: 700;
        color: #FFFFFF;
        line-height: 1.35;
      }
      .os-mig-sub {
        font-size: 0.78rem;
        color: #A5B4C8;
        margin-top: 2px;
      }
      .os-mig-right {
        display: flex;
        align-items: center;
        gap: 10px;
        flex-wrap: wrap;
        justify-content: flex-end;
      }
      .os-mig-fc-pill {
        display: inline-flex;
        align-items: center;
        gap: 5px;
        background: #0E182B;
        border: 1px solid #223456;
        color: #E2E8F0;
        font-size: 0.8rem;
        font-weight: 700;
        padding: 4px 10px;
        border-radius: 7px;
        white-space: nowrap;
      }
      .os-mig-icon-btn {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 28px;
        height: 28px;
        padding: 0;
        border-radius: 6px;
        background: transparent;
        border: 1px solid rgba(51, 65, 85, 0.75);
        color: #94A3B8;
        cursor: pointer;
        transition: all 0.14s ease;
      }
      .os-mig-icon-btn:hover {
        background: rgba(37, 99, 235, 0.18);
        border-color: rgba(96, 165, 250, 0.5);
        color: #93C5FD;
      }
      .os-mig-icon-del:hover {
        background: rgba(239, 68, 68, 0.16);
        border-color: rgba(248, 113, 113, 0.5);
        color: #FCA5A5;
      }
      .os-mig-icon-btn:focus-visible {
        outline: 2px solid #60A5FA;
        outline-offset: 2px;
      }
      @media (max-width: 768px) {
        .os-mig-row {
          flex-direction: column;
          align-items: stretch;
          gap: 10px;
          padding: 12px 14px;
        }
        .os-mig-left { flex: none; }
        .os-mig-right {
          justify-content: space-between;
        }
      }
    </style>

    <!-- Main Content Container -->
    <div id="migration-view-container"></div>
  `;

  let currentView = 'board'; // 'board' or 'table'
  let summaryData = null;
  let allBranches = [];
  // canEdit defined at top

  // Load branches for modal dropdown
  apiFetch('/api/branches?all=1').then(r => {
    if (r.ok && r.data?.data) allBranches = r.data.data;
  }).catch(() => {});

  // Fetch summary & data
  async function loadData() {
    try {
      const res = await apiFetch('/api/os-migrations/summary');
      if (!res.ok) throw new Error('Gagal memuat data migrasi');
      summaryData = res.data?.data || res.data;
      renderKPIs(summaryData);
      renderView();
    } catch (err) {
      const mvc = document.getElementById('migration-view-container');
      mvc.innerHTML = `
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">⚠️</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data: ${escapeHtml(err.message)}</p>
          <button type="button" id="os-mig-reload-btn" class="btn btn-outline btn-sm">Muat Ulang</button>
        </div>
      `;
      mvc.querySelector('#os-mig-reload-btn')?.addEventListener('click', () => window.location.reload());
    }
  }

  function renderKPIs(data) {
    const kpiEl = document.getElementById('migration-summary-cards');
    if (!kpiEl || !data) return;

    const total = data.total_branches || 0;
    const aeonTotal = data.vendors?.AEON?.total || 0;
    const tripleSTotal = data.vendors?.['Triple S']?.total || 0;
    const completed = data.by_status?.completed || 0;
    const inProgress = data.by_status?.in_progress || 0;
    const scheduled = data.by_status?.scheduled || 0;
    const pct = total > 0 ? Math.round((completed / total) * 100) : 0;

    kpiEl.innerHTML = `
      <div class="card" style="padding:16px 20px; background:#131F36; border:1px solid #223456; border-radius:12px;">
        <div style="font-size:0.8rem; font-weight:600; color:#A5B4C8; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Cabang Migrasi</div>
        <div style="font-size:1.75rem; font-weight:800; color:#FFFFFF; line-height:1.2;">${total} <span style="font-size:0.9rem; font-weight:500; color:#A5B4C8;">Cabang</span></div>
        <div style="font-size:0.76rem; color:#CBD5E1; margin-top:4px;">Timeline Periode Sep - Des 2026</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#131F36; border:1px solid #223456; border-left:4px solid #2563EB; border-radius:12px;">
        <div style="font-size:0.8rem; font-weight:600; color:#93C5FD; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor AEON</div>
        <div style="font-size:1.75rem; font-weight:800; color:#60A5FA; line-height:1.2;">${aeonTotal} <span style="font-size:0.9rem; font-weight:500; color:#A5B4C8;">Cabang</span></div>
        <div style="font-size:0.76rem; color:#CBD5E1; margin-top:4px;">Sep: 9 • Nov: 3 • Des: 8</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#131F36; border:1px solid #223456; border-left:4px solid #38BDF8; border-radius:12px;">
        <div style="font-size:0.8rem; font-weight:600; color:#7DD3FC; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor Triple S</div>
        <div style="font-size:1.75rem; font-weight:800; color:#38BDF8; line-height:1.2;">${tripleSTotal} <span style="font-size:0.9rem; font-weight:500; color:#A5B4C8;">Cabang</span></div>
        <div style="font-size:0.76rem; color:#CBD5E1; margin-top:4px;">Sep: 5 • Okt: 5 • Nov: 5 • Des: 27</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#131F36; border:1px solid #223456; border-left:4px solid #34D399; border-radius:12px;">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
          <span style="font-size:0.8rem; font-weight:600; color:#6EE7B7; text-transform:uppercase; letter-spacing:0.5px;">Progres Realisasi</span>
          <span style="font-size:0.88rem; font-weight:800; color:#34D399;">${pct}%</span>
        </div>
        <div style="width:100%; height:8px; background:#0E182B; border-radius:4px; overflow:hidden; margin-bottom:8px;">
          <div style="width:${pct}%; height:100%; background:#34D399; border-radius:4px;"></div>
        </div>
        <div style="display:flex; justify-content:space-between; font-size:0.76rem; color:#CBD5E1; font-weight:600;">
          <span>Selesai: <strong style="color:#6EE7B7;">${completed}</strong></span>
          <span>Proses: <strong style="color:#93C5FD;">${inProgress}</strong></span>
          <span>Terjadwal: <strong style="color:#FCD34D;">${scheduled}</strong></span>
        </div>
      </div>
    `;
  }

  function getFilteredRows() {
    if (!summaryData || !summaryData.matrix) return [];
    const search = (document.getElementById('migration-search')?.value || '').toLowerCase().trim();
    const vendor = document.getElementById('filter-vendor')?.value || '';
    const month = document.getElementById('filter-month')?.value || '';
    const status = document.getElementById('filter-status')?.value || '';

    return summaryData.matrix.filter(r => {
      if (vendor && r.vendor !== vendor) return false;
      if (month && r.target_month !== month) return false;
      if (status && r.status !== status) return false;
      if (search) {
        const text = `${r.branch_code} ${r.branch_name} ${r.branch_full_name} ${r.vendor} ${r.notes || ''} ${r.pic_vendor || ''}`.toLowerCase();
        if (!text.includes(search)) return false;
      }
      return true;
    });
  }

  function renderView() {
    const rows = getFilteredRows();
    const statsEl = document.getElementById('filter-stats-text');
    if (statsEl) {
      statsEl.innerHTML = `Menampilkan <strong>${rows.length}</strong> dari <strong>${summaryData?.total_branches || 0}</strong> cabang`;
    }

    const containerEl = document.getElementById('migration-view-container');
    if (currentView === 'board') {
      renderBoardView(containerEl, rows);
    } else {
      renderTableView(containerEl, rows);
    }
  }

  function renderBoardView(targetEl, rows) {
    if (!rows.length) {
      targetEl.innerHTML = `
        <div class="empty-state" style="padding:50px; text-align:center; background:#111827; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">🔍</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci pencarian atau pilihan filter di atas.</p>
        </div>
      `;
      return;
    }

    // Group rows by Vendor, then by Target Month (newest first)
    const vendors = ['AEON', 'Triple S'];
    const months = sortYearMonthDesc(Array.from(new Set(rows.map(r => r.target_month).filter(Boolean))));

    let html = '';

    vendors.forEach(v => {
      const vRows = rows.filter(r => r.vendor === v);
      if (document.getElementById('filter-vendor')?.value && document.getElementById('filter-vendor').value !== v) return;
      if (vRows.length === 0) return;

      const vTotal = vRows.length;
      const vThemeColor = v === 'AEON' ? '#93C5FD' : '#7DD3FC';
      const vBadgeBorder = v === 'AEON' ? 'rgba(59, 130, 246, 0.35)' : 'rgba(56, 189, 248, 0.35)';

      const activeMonths = months.filter(m => vRows.some(r => r.target_month === m));

      html += `
        <div class="vendor-board-section" style="margin-bottom:24px; background:#131F36; border:1px solid #223456; border-radius:14px; overflow:hidden;">
          <!-- Vendor Section Header -->
          <div style="padding:14px 18px; background:#0E182B; border-bottom:1px solid #223456; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px;">
            <div style="display:flex; align-items:center; gap:10px;">
              <span style="font-size:1.05rem; font-weight:800; color:${vThemeColor};">Vendor ${escapeHtml(v)}</span>
              <span class="badge" style="background:rgba(37,99,235,0.16); color:${vThemeColor}; border:1px solid ${vBadgeBorder}; font-weight:700; font-size:0.76rem; padding:3px 9px; border-radius:6px;">
                ${vTotal} Cabang
              </span>
            </div>
            <div style="font-size:0.8rem; color:#A5B4C8; font-weight:500;">
              Dikelompokkan per Bulan Target
            </div>
          </div>

          <!-- Single Unified Grouped List Container (Zero nested boxes or inner scrollbars) -->
          <div>
            ${activeMonths.map(m => {
              const mRows = vRows.filter(r => r.target_month === m);
              const mName = MONTH_NAMES[m] || m;

              return `
                <div class="os-mig-month-group">
                  <div style="padding:9px 18px; background:#101A2F; border-top:1px solid #1E2E4A; border-bottom:1px solid #1E2E4A; display:flex; justify-content:space-between; align-items:center;">
                    <span style="font-weight:700; font-size:0.88rem; color:#E2E8F0; letter-spacing:0.01em;">
                      📅 ${escapeHtml(mName)}
                    </span>
                    <span style="font-size:0.76rem; font-weight:700; color:#93C5FD; background:rgba(37,99,235,0.14); border:1px solid rgba(59,130,246,0.28); padding:2px 8px; border-radius:10px;">
                      ${mRows.length} Cabang
                    </span>
                  </div>

                  <div>
                    ${mRows.map(r => {
                      const st = STATUS_CONFIG[r.status] || STATUS_CONFIG.scheduled;
                      return `
                        <div class="os-mig-row">
                          <div class="os-mig-left">
                            <span class="os-mig-code">${escapeHtml(r.branch_code || '-')}</span>
                            <div style="min-width:0;">
                              <div class="os-mig-name">${escapeHtml(r.branch_name || '-')}</div>
                              ${(r.actual_date || r.notes) ? `
                                <div class="os-mig-sub">
                                  ${r.actual_date ? `Realisasi: <strong style="color:#CBD5E1;">${escapeHtml(r.actual_date)}</strong>` : ''}
                                  ${(r.actual_date && r.notes) ? ' • ' : ''}
                                  ${r.notes ? `${escapeHtml(r.notes)}` : ''}
                                </div>
                              ` : ''}
                            </div>
                          </div>

                          <div class="os-mig-right">
                            <span class="os-mig-fc-pill" title="${r.employee_count || 0} personil Facility Care aktif di cabang ini">
                              👤 ${r.employee_count || 0} FC
                            </span>
                            <button type="button" class="status-toggle-btn" data-id="${escapeHtml(String(r.id))}" data-status="${escapeHtml(r.status)}" style="background:${st.bg}; color:${st.color}; border:1px solid ${st.border}; padding:4px 10px; border-radius:7px; font-size:0.78rem; font-weight:700; display:inline-flex; align-items:center; gap:6px; cursor:${canEdit ? 'pointer' : 'default'}; white-space:nowrap;" title="${canEdit ? 'Klik untuk ubah status' : 'Status'}">
                              <span style="width:7px; height:7px; border-radius:50%; background:${st.dot}; display:inline-block;"></span>
                              <span>${st.label}${canEdit ? ' ▾' : ''}</span>
                            </button>
                            ${renderMigrationActionButtons(r, canEdit, canDelete)}
                          </div>
                        </div>
                      `;
                    }).join('')}
                  </div>
                </div>
              `;
            }).join('')}
          </div>
        </div>
      `;
    });

    targetEl.innerHTML = html;
    bindBoardEvents(targetEl);
  }

  function renderTableView(targetEl, rows) {
    if (!rows.length) {
      targetEl.innerHTML = `
        <div class="empty-state" style="padding:50px; text-align:center; background:#111827; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">🔍</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
        </div>
      `;
      return;
    }

    targetEl.innerHTML = `
      <div class="card" style="background:#131F36; border:1px solid #223456; border-radius:12px; overflow:hidden;">
        <div class="table-responsive">
          <table class="table" style="width:100%; border-collapse:collapse;">
            <thead>
              <tr style="background:#0E182B; border-bottom:1px solid #223456; text-align:left;">
                <th style="padding:12px 16px; font-size:0.78rem; font-weight:700; color:#A5B4C8; width:50px;">NO</th>
                <th style="padding:12px 16px; font-size:0.78rem; font-weight:700; color:#A5B4C8;">KODE & CABANG</th>
                <th style="padding:12px 16px; font-size:0.78rem; font-weight:700; color:#A5B4C8;">VENDOR</th>
                <th style="padding:12px 16px; font-size:0.78rem; font-weight:700; color:#A5B4C8;">TARGET BULAN</th>
                <th style="padding:12px 16px; font-size:0.78rem; font-weight:700; color:#A5B4C8;">STATUS</th>
                <th style="padding:12px 16px; font-size:0.78rem; font-weight:700; color:#A5B4C8;">REALISASI</th>
                <th style="padding:12px 16px; font-size:0.78rem; font-weight:700; color:#A5B4C8;">SDM FC</th>
                <th style="padding:12px 16px; font-size:0.78rem; font-weight:700; color:#A5B4C8;">CATATAN</th>
                ${canEdit || canDelete ? `<th style="padding:12px 16px; font-size:0.78rem; font-weight:700; color:#A5B4C8; text-align:right;">AKSI</th>` : ''}
              </tr>
            </thead>
            <tbody>
              ${rows.map((r, idx) => {
                const st = STATUS_CONFIG[r.status] || STATUS_CONFIG.scheduled;
                return `
                  <tr style="border-bottom:1px solid #1E2E4A;">
                    <td style="padding:12px 16px; font-size:0.85rem; color:#A5B4C8;">${idx + 1}</td>
                    <td style="padding:12px 16px;">
                      <div style="font-weight:700; font-size:0.9rem; color:#FFFFFF;">${escapeHtml(r.branch_full_name)}</div>
                    </td>
                    <td style="padding:12px 16px;">
                      <span class="badge" style="background:${r.vendor === 'AEON' ? 'rgba(37,99,235,0.16)' : 'rgba(56,189,248,0.16)'}; color:${r.vendor === 'AEON' ? '#93C5FD' : '#7DD3FC'}; border:1px solid ${r.vendor === 'AEON' ? 'rgba(59,130,246,0.35)' : 'rgba(56,189,248,0.35)'}; font-weight:700; border-radius:6px; padding:4px 8px; font-size:0.75rem;">
                        ${escapeHtml(r.vendor)}
                      </span>
                    </td>
                    <td style="padding:12px 16px; font-size:0.875rem; font-weight:600; color:#CBD5E1;">
                      ${escapeHtml(MONTH_NAMES[r.target_month] || r.target_month)}
                    </td>
                    <td style="padding:12px 16px;">
                      <button type="button" class="status-toggle-btn" data-id="${escapeHtml(String(r.id))}" data-status="${escapeHtml(r.status)}" style="background:${st.bg}; color:${st.color}; border:1px solid ${st.border}; padding:4px 9px; border-radius:6px; font-size:0.76rem; font-weight:700; display:inline-flex; align-items:center; gap:5px; cursor:${canEdit ? 'pointer' : 'default'};" title="${canEdit ? 'Klik untuk ubah status' : 'Status'}">
                        <span style="width:6px; height:6px; border-radius:50%; background:${st.dot}; display:inline-block;"></span> ${st.label}${canEdit ? ' ▾' : ''}
                      </button>
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:#CBD5E1;">
                      ${escapeHtml(r.actual_date || '-')}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; font-weight:700; color:#E2E8F0;">
                      ${r.employee_count > 0 ? `👤 ${r.employee_count} FC` : '-'}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:#A5B4C8; max-width:200px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${escapeHtml(r.notes || '-')}
                    </td>
                    ${canEdit || canDelete ? `
                      <td style="padding:12px 16px; text-align:right;">
                        ${renderMigrationActionButtons(r, canEdit, canDelete)}
                      </td>
                    ` : ''}
                  </tr>
                `;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>
    `;
    bindBoardEvents(targetEl);
  }

  function bindBoardEvents(targetEl) {
    // Status quick changer popup / modal (only for users with write access)
    targetEl.querySelectorAll('.status-toggle-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        if (!canEdit) return;
        const id = btn.dataset.id;
        const currentSt = btn.dataset.status;
        showStatusMenu(btn, id, currentSt);
      });
    });

    // Edit button
    targetEl.querySelectorAll('.edit-card-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.dataset.id;
        const item = summaryData?.matrix?.find(r => String(r.id) === String(id));
        if (item) showFormModal(item);
      });
    });

    // Delete button
    targetEl.querySelectorAll('.delete-card-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        if (!canDelete) {
          toastError('Hanya Superadmin yang dapat menghapus jadwal migrasi.');
          return;
        }
        const id = btn.dataset.id;
        const name = btn.dataset.name || 'Cabang';
        confirmDialog(`Apakah Anda yakin ingin menghapus jadwal migrasi untuk <strong>${name}</strong>?`, async () => {
          try {
            const res = await apiFetch(`/api/os-migrations/${id}`, { method: 'DELETE' });
            if (!res.ok) throw new Error('Gagal menghapus');
            toastSuccess(`Jadwal migrasi ${name} berhasil dihapus`);
            loadData();
          } catch (err) {
            toastError(err.message);
          }
        });
      });
    });
  }

  function showStatusMenu(anchorBtn, id, currentSt) {
    if (!canEdit) return;
    // Remove any open popup
    document.querySelectorAll('.status-popup-menu').forEach(el => el.remove());

    const rect = anchorBtn.getBoundingClientRect();
    const popup = document.createElement('div');
    popup.className = 'status-popup-menu';
    popup.style.cssText = `
      position: fixed;
      top: ${rect.bottom + window.scrollY + 4}px;
      left: ${rect.left + window.scrollX}px;
      background: #172033;
      border: 1px solid var(--border);
      border-radius: 8px;
      box-shadow: 0 4px 16px rgba(0,0,0,0.12);
      z-index: 9999;
      padding: 6px 0;
      min-width: 160px;
    `;

    const statuses = [
      { key: 'scheduled', label: '🟡 Terjadwal' },
      { key: 'in_progress', label: '🔵 Sedang Proses' },
      { key: 'completed', label: '🟢 Selesai' }
    ];

    popup.innerHTML = statuses.map(s => `
      <button class="status-option-item" data-key="${s.key}" style="
        width: 100%;
        text-align: left;
        padding: 8px 14px;
        background: none;
        border: none;
        font-size: 0.85rem;
        cursor: pointer;
        display: flex;
        align-items: center;
        gap: 8px;
        color: ${s.key === currentSt ? '#60A5FA' : 'var(--text-1)'};
        font-weight: ${s.key === currentSt ? '700' : '500'};
      ">
        ${s.label} ${s.key === currentSt ? '✓' : ''}
      </button>
    `).join('');

    const closePopup = (e) => {
      if (!popup.contains(e.target) && e.target !== anchorBtn) {
        popup.remove();
        document.removeEventListener('click', closePopup);
      }
    };
    setTimeout(() => document.addEventListener('click', closePopup), 50);

    popup.querySelectorAll('.status-option-item').forEach(opt => {
      opt.addEventListener('click', async () => {
        const nextStatus = opt.dataset.key;
        popup.remove();
        document.removeEventListener('click', closePopup);

        if (nextStatus === currentSt) return;

        try {
          const res = await apiFetch(`/api/os-migrations/${id}/status`, {
            method: 'PATCH',
            body: JSON.stringify({ status: nextStatus })
          });
          if (!res.ok) throw new Error('Gagal memperbarui status');
          toastSuccess('Status berhasil diubah');
          loadData();
        } catch (err) {
          toastError(err.message);
        }
      });
    });

    document.body.appendChild(popup);
  }

  function showFormModal(item = null) {
    if (!canEdit) {
      toastError('Anda tidak memiliki izin untuk menambah atau mengedit migrasi OS.');
      return;
    }
    const isEdit = !!item;
    const title = isEdit ? 'Edit Jadwal Migrasi OS' : 'Tambah Jadwal Migrasi OS';

    const branchOptionsHtml = allBranches.map(b => `
      <option value="${b.id}" ${item && item.branch_id === b.id ? 'selected' : ''}>
        ${b.full_name}
      </option>
    `).join('');

    const contentHtml = `
      <form id="migration-form" style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Cabang *</label>
          <select name="branch_id" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
            <option value="">-- Pilih Cabang --</option>
            ${branchOptionsHtml}
          </select>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Vendor Outsourcing *</label>
            <select name="vendor" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="AEON" ${item?.vendor === 'AEON' ? 'selected' : ''}>AEON</option>
              <option value="Triple S" ${item?.vendor === 'Triple S' ? 'selected' : ''}>Triple S</option>
              <option value="Lainnya" ${item && !['AEON', 'Triple S'].includes(item.vendor) ? 'selected' : ''}>Lainnya</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Target Bulan *</label>
            <select name="target_month" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="2026-09" ${item?.target_month === '2026-09' ? 'selected' : ''}>September 2026</option>
              <option value="2026-10" ${item?.target_month === '2026-10' ? 'selected' : ''}>Oktober 2026</option>
              <option value="2026-11" ${item?.target_month === '2026-11' ? 'selected' : ''}>November 2026</option>
              <option value="2026-12" ${item?.target_month === '2026-12' ? 'selected' : ''}>Desember 2026</option>
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Status Pelaksanaan</label>
            <select name="status" class="form-control" style="width:100%; height:38px; border-radius:8px;">
              <option value="scheduled" ${item?.status === 'scheduled' ? 'selected' : ''}>🟡 Terjadwal</option>
              <option value="in_progress" ${item?.status === 'in_progress' ? 'selected' : ''}>🔵 Sedang Proses</option>
              <option value="completed" ${item?.status === 'completed' ? 'selected' : ''}>🟢 Selesai</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Tanggal Realisasi (Opsional)</label>
            <input type="date" name="actual_date" class="form-control" value="${item?.actual_date || ''}" style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">PIC / Koordinator Vendor (Opsional)</label>
          <input type="text" name="pic_vendor" class="form-control" value="${item?.pic_vendor || ''}" style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Catatan / Keterangan</label>
          <textarea name="notes" class="form-control" rows="2" style="width:100%; border-radius:8px;">${item?.notes || ''}</textarea>
        </div>
      </form>
    `;

    createModal({
      title,
      content: contentHtml,
      confirmText: isEdit ? 'Simpan Perubahan' : 'Tambah Cabang',
      onConfirm: async (overlay, close) => {
        const form = overlay.querySelector('#migration-form');
        if (!form.checkValidity()) {
          form.reportValidity();
          return;
        }

        const formData = new FormData(form);
        const payload = {
          branch_id: parseInt(formData.get('branch_id'), 10),
          vendor: formData.get('vendor'),
          target_month: formData.get('target_month'),
          status: formData.get('status'),
          actual_date: formData.get('actual_date') || null,
          pic_vendor: formData.get('pic_vendor') || null,
          notes: formData.get('notes') || null
        };

        try {
          const url = isEdit ? `/api/os-migrations/${item.id}` : '/api/os-migrations';
          const method = isEdit ? 'PUT' : 'POST';
          const res = await apiFetch(url, { method, body: JSON.stringify(payload) });

          if (!res.ok) {
            const errMsg = res.data?.error || res.data?.message || 'Gagal menyimpan data';
            throw new Error(errMsg);
          }

          toastSuccess(isEdit ? 'Jadwal migrasi berhasil diupdate' : 'Jadwal migrasi berhasil ditambahkan');
          close();
          loadData();
        } catch (err) {
          toastError(err.message);
        }
      }
    });
  }

  // View switchers
  const boardBtn = document.getElementById('view-board-btn');
  const tableBtn = document.getElementById('view-table-btn');

  boardBtn.addEventListener('click', () => {
    currentView = 'board';
    boardBtn.style.background = '#fff';
    boardBtn.classList.remove('btn-ghost');
    tableBtn.style.background = 'transparent';
    tableBtn.classList.add('btn-ghost');
    renderView();
  });

  tableBtn.addEventListener('click', () => {
    currentView = 'table';
    tableBtn.style.background = '#fff';
    tableBtn.classList.remove('btn-ghost');
    boardBtn.style.background = 'transparent';
    boardBtn.classList.add('btn-ghost');
    renderView();
  });

  // Filter events
  ['migration-search', 'filter-vendor', 'filter-month', 'filter-status'].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('input', renderView);
      el.addEventListener('change', renderView);
    }
  });

  // Add migration button
  document.getElementById('btn-add-migration')?.addEventListener('click', () => {
    showFormModal();
  });

  // Export Excel button
  document.getElementById('btn-export-migrations')?.addEventListener('click', async () => {
    try {
      const search = (document.getElementById('migration-search')?.value || '').trim();
      const vendor = document.getElementById('filter-vendor')?.value || '';
      const month = document.getElementById('filter-month')?.value || '';
      const status = document.getElementById('filter-status')?.value || '';

      const params = {};
      if (vendor) params.vendor = vendor;
      if (month) params.month = month;
      if (status) params.status = status;
      if (search) params.search = search;

      const rows = await fetchExportRows('/api/os-migrations/export', params);
      if (!rows.length) {
        toastError('Tidak ada data untuk diexport');
        return;
      }
      const exportData = rows.map((r, idx) => ({
        'No': idx + 1,
        'Kode Cabang': r.branch_code,
        'Nama Cabang': r.branch_name,
        'Cabang Lengkap': r.branch_full_name,
        'Vendor OS': r.vendor,
        'Target Bulan': MONTH_NAMES[r.target_month] || r.target_month,
        'Status': STATUS_CONFIG[r.status]?.label || r.status,
        'Tanggal Realisasi': r.actual_date || '',
        'Jml Personil FC': r.employee_count || 0,
        'PIC Vendor': r.pic_vendor || '',
        'Catatan': r.notes || ''
      }));
      downloadExcel(exportData, `Timeline_Migrasi_OS_${new Date().toISOString().slice(0, 10)}`);
      toastSuccess('File Excel berhasil diunduh');
    } catch (err) {
      toastError('Gagal mengekspor data migrasi: ' + err.message);
    }
  });

  // Initial load
  loadData();
}
