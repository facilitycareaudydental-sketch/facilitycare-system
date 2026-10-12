import { apiFetch, getUser } from '../config.js';
import { createModal, confirmDialog } from '../components/modal.js';
import { toastSuccess, toastError } from '../components/toast.js';
import { downloadExcel } from '../utils/excel.js';

const MONTH_NAMES = {
  '2026-09': 'September 2026',
  '2026-10': 'Oktober 2026',
  '2026-11': 'November 2026',
  '2026-12': 'Desember 2026'
};

const STATUS_CONFIG = {
  scheduled: { label: 'Terjadwal', badgeClass: 'badge-warning', icon: '🟡', bg: '#FEF9C3', color: '#854D0E', border: '#FDE047' },
  in_progress: { label: 'Sedang Proses', badgeClass: 'badge-info', icon: '🔵', bg: '#DBEAFE', color: '#1E40AF', border: '#93C5FD' },
  completed: { label: 'Selesai', badgeClass: 'badge-success', icon: '🟢', bg: '#DCFCE7', color: '#166534', border: '#86EFAC' }
};

export async function renderOSMigrations(container) {
  container.innerHTML = `
    <div class="page-header" style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px; margin-bottom:20px;">
      <div>
        <h1 class="page-title" style="display:flex; align-items:center; gap:8px; margin:0 0 6px 0; font-size:1.6rem; font-weight:700; color:var(--text-1);">
          <span>🔄</span> Time Line Migrasi OS
        </h1>
        <p class="page-subtitle" style="margin:0; font-size:0.9rem; color:var(--text-3);">
          Jadwal & Status Perpindahan Vendor Outsourcing (AEON & Triple S) ke Seluruh Cabang Audy Dental
        </p>
      </div>
      <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
        <div class="btn-group" style="background:#F1F5F9; padding:3px; border-radius:8px; display:inline-flex; border:1px solid var(--border);">
          <button id="view-board-btn" class="btn btn-sm" style="background:#fff; font-weight:600; border-radius:6px; box-shadow:0 1px 2px rgba(0,0,0,0.05); padding:6px 12px; cursor:pointer;">
            🗂️ Matrix Board
          </button>
          <button id="view-table-btn" class="btn btn-sm btn-ghost" style="font-weight:600; border-radius:6px; padding:6px 12px; cursor:pointer;">
            📋 Tabel Data
          </button>
        </div>
        <button id="btn-export-migrations" class="btn btn-outline" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;">
          📥 Export Excel
        </button>
        <button id="btn-add-migration" class="btn btn-primary" style="font-weight:600; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;">
          + Tambah Cabang
        </button>
      </div>
    </div>

    <!-- Summary KPI Cards -->
    <div id="migration-summary-cards" style="display:grid; grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); gap:16px; margin-bottom:24px;">
      <div class="kpi-card skeleton" style="height:100px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:100px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:100px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:100px; border-radius:12px;"></div>
    </div>

    <!-- Filter Toolbar -->
    <div class="card" style="padding:16px 20px; margin-bottom:24px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
      <div style="display:flex; flex-wrap:wrap; gap:14px; align-items:center; justify-content:space-between;">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
          <div style="position:relative; flex:1; min-width:200px; max-width:320px;">
            <input type="text" id="migration-search" class="form-control" placeholder="Cari cabang, vendor, catatan..." style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
            <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">🔍</span>
          </div>
          <select id="filter-vendor" class="form-control" style="width:auto; height:38px; border-radius:8px; font-size:0.875rem; min-width:140px;">
            <option value="">Semua Vendor</option>
            <option value="AEON">AEON</option>
            <option value="Triple S">Triple S</option>
          </select>
          <select id="filter-month" class="form-control" style="width:auto; height:38px; border-radius:8px; font-size:0.875rem; min-width:150px;">
            <option value="">Semua Bulan</option>
            <option value="2026-09">September 2026</option>
            <option value="2026-10">Oktober 2026</option>
            <option value="2026-11">November 2026</option>
            <option value="2026-12">Desember 2026</option>
          </select>
          <select id="filter-status" class="form-control" style="width:auto; height:38px; border-radius:8px; font-size:0.875rem; min-width:140px;">
            <option value="">Semua Status</option>
            <option value="scheduled">🟡 Terjadwal</option>
            <option value="in_progress">🔵 Sedang Proses</option>
            <option value="completed">🟢 Selesai</option>
          </select>
        </div>
        <div id="filter-stats-text" style="font-size:0.85rem; color:var(--text-3); font-weight:500;">
          Memuat data...
        </div>
      </div>
    </div>

    <!-- Main Content Container -->
    <div id="migration-view-container"></div>
  `;

  let currentView = 'board'; // 'board' or 'table'
  let summaryData = null;
  let allBranches = [];
  let user = getUser();
  const canEdit = user && ['superadmin', 'admin', 'manager', 'spv'].includes(user.role);

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
      document.getElementById('migration-view-container').innerHTML = `
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">⚠️</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data: ${err.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `;
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
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Cabang Migrasi</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${total} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Timeline Periode Sep - Des 2026</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor AEON</div>
        <div style="font-size:1.8rem; font-weight:800; color:#1E40AF; line-height:1.2;">${aeonTotal} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Sep: 9 • Nov: 3 • Des: 8</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #0D9488; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#0F766E; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Vendor Triple S</div>
        <div style="font-size:1.8rem; font-weight:800; color:#0F766E; line-height:1.2;">${tripleSTotal} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Cabang</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Sep: 5 • Okt: 5 • Nov: 5 • Des: 27</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #16A34A; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
          <span style="font-size:0.8rem; font-weight:600; color:#15803D; text-transform:uppercase; letter-spacing:0.5px;">Progres Realisasi</span>
          <span style="font-size:0.85rem; font-weight:700; color:#15803D;">${pct}%</span>
        </div>
        <div style="width:100%; height:8px; background:#E2E8F0; border-radius:4px; overflow:hidden; margin-bottom:8px;">
          <div style="width:${pct}%; height:100%; background:#16A34A; border-radius:4px;"></div>
        </div>
        <div style="display:flex; justify-content:space-between; font-size:0.75rem; color:var(--text-2); font-weight:500;">
          <span>🟢 Selesai: <strong>${completed}</strong></span>
          <span>🔵 Proses: <strong>${inProgress}</strong></span>
          <span>🟡 Terjadwal: <strong>${scheduled}</strong></span>
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
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">🔍</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci pencarian atau pilihan filter di atas.</p>
        </div>
      `;
      return;
    }

    // Group rows by Vendor, then by Target Month
    const vendors = ['AEON', 'Triple S'];
    const months = ['2026-09', '2026-10', '2026-11', '2026-12'];

    let html = '';

    vendors.forEach(v => {
      const vRows = rows.filter(r => r.vendor === v);
      if (document.getElementById('filter-vendor')?.value && document.getElementById('filter-vendor').value !== v) return;

      const vTotal = vRows.length;
      const vThemeColor = v === 'AEON' ? '#1E40AF' : '#0F766E';
      const vBgColor = v === 'AEON' ? '#EFF6FF' : '#F0FDFA';
      const vBorderColor = v === 'AEON' ? '#BFDBFE' : '#99F6E4';

      html += `
        <div class="vendor-board-section" style="margin-bottom:32px; background:#fff; border:1px solid var(--border); border-radius:14px; overflow:hidden; box-shadow:0 1px 4px rgba(0,0,0,0.04);">
          <!-- Vendor Section Header -->
          <div style="padding:14px 20px; background:${vBgColor}; border-bottom:1px solid ${vBorderColor}; display:flex; justify-content:space-between; align-items:center;">
            <div style="display:flex; align-items:center; gap:10px;">
              <span style="font-size:1.1rem; font-weight:800; color:${vThemeColor};">${v === 'AEON' ? '🏢 AEON' : '🛡️ Triple S'}</span>
              <span class="badge" style="background:#fff; color:${vThemeColor}; border:1px solid ${vBorderColor}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px;">
                ${vTotal} Cabang
              </span>
            </div>
            <div style="font-size:0.8rem; color:var(--text-3); font-weight:500;">
              Format Timeline Berdasarkan Dokumen Migrasi
            </div>
          </div>

          <!-- Months 4-Column Grid -->
          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(240px, 1fr)); gap:16px; padding:18px;">
            ${months.map(m => {
              const mRows = vRows.filter(r => r.target_month === m);
              const mName = MONTH_NAMES[m] || m;
              const shortMonth = mName.split(' ')[0];

              return `
                <div class="month-column" style="background:#F8FAFC; border:1px solid var(--border); border-radius:10px; display:flex; flex-direction:column; min-height:220px;">
                  <!-- Month Header -->
                  <div style="padding:12px 14px; background:#F1F5F9; border-bottom:1px solid var(--border); border-radius:9px 9px 0 0; display:flex; justify-content:space-between; align-items:center;">
                    <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">
                      📅 ${shortMonth}
                    </div>
                    <span class="badge" style="background:#E2E8F0; color:var(--text-2); font-weight:700; font-size:0.75rem; padding:2px 7px; border-radius:4px;">
                      ${mRows.length}
                    </span>
                  </div>

                  <!-- Branch Cards Container -->
                  <div style="padding:10px; display:flex; flex-direction:column; gap:8px; flex:1; max-height:480px; overflow-y:auto;">
                    ${mRows.length === 0 ? `
                      <div style="text-align:center; padding:30px 10px; color:var(--text-3); font-size:0.8rem; font-style:italic;">
                        Tidak ada jadwal
                      </div>
                    ` : mRows.map(r => {
                      const st = STATUS_CONFIG[r.status] || STATUS_CONFIG.scheduled;
                      return `
                        <div class="branch-card" style="background:#fff; border:1px solid var(--border); border-radius:8px; padding:10px 12px; box-shadow:0 1px 2px rgba(0,0,0,0.03); transition:transform 0.15s, box-shadow 0.15s;">
                          <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:6px; margin-bottom:6px;">
                            <div style="font-weight:700; font-size:0.9rem; color:var(--text-1); line-height:1.3;">
                              <span style="color:#2563EB; font-family:monospace; margin-right:4px;">${r.branch_code}.</span>${r.branch_name}
                            </div>
                            ${canEdit ? `
                              <div style="display:flex; gap:2px;">
                                <button class="btn-icon-mini edit-card-btn" data-id="${r.id}" title="Edit" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--text-3);">✏️</button>
                                <button class="btn-icon-mini delete-card-btn" data-id="${r.id}" data-name="${r.branch_name}" title="Hapus" style="background:none; border:none; cursor:pointer; padding:2px 4px; font-size:0.8rem; color:var(--danger);">🗑️</button>
                              </div>
                            ` : ''}
                          </div>

                          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:6px; margin-top:8px;">
                            <!-- Status Selector Dropdown -->
                            <div style="position:relative;">
                              <button class="status-toggle-btn" data-id="${r.id}" data-status="${r.status}" style="background:${st.bg}; color:${st.color}; border:1px solid ${st.border}; padding:3px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; display:inline-flex; align-items:center; gap:4px; cursor:pointer;" title="Klik untuk ubah status">
                                ${st.icon} ${st.label} ▾
                              </button>
                            </div>

                            ${r.employee_count > 0 ? `
                              <span style="font-size:0.75rem; color:var(--text-3); font-weight:600; display:inline-flex; align-items:center; gap:3px;" title="${r.employee_count} personil Facility Care aktif di cabang ini">
                                👤 ${r.employee_count} FC
                              </span>
                            ` : ''}
                          </div>

                          ${r.actual_date ? `
                            <div style="font-size:0.7rem; color:var(--text-3); margin-top:6px;">
                              Realisasi: <strong>${r.actual_date}</strong>
                            </div>
                          ` : ''}
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
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">🔍</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada cabang yang cocok dengan filter</p>
        </div>
      `;
      return;
    }

    targetEl.innerHTML = `
      <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div class="table-responsive">
          <table class="table" style="width:100%; border-collapse:collapse;">
            <thead>
              <tr style="background:#F8FAFC; border-bottom:1px solid var(--border); text-align:left;">
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:50px;">NO</th>
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3);">KODE & CABANG</th>
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3);">VENDOR</th>
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3);">TARGET BULAN</th>
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3);">STATUS</th>
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3);">REALISASI</th>
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3);">SDM FC</th>
                <th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3);">CATATAN</th>
                ${canEdit ? `<th style="padding:12px 16px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right;">AKSI</th>` : ''}
              </tr>
            </thead>
            <tbody>
              ${rows.map((r, idx) => {
                const st = STATUS_CONFIG[r.status] || STATUS_CONFIG.scheduled;
                return `
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3);">${idx + 1}</td>
                    <td style="padding:12px 16px;">
                      <div style="font-weight:700; font-size:0.9rem; color:var(--text-1);">${r.branch_full_name}</div>
                    </td>
                    <td style="padding:12px 16px;">
                      <span class="badge" style="background:${r.vendor === 'AEON' ? '#DBEAFE' : '#CCFBF1'}; color:${r.vendor === 'AEON' ? '#1E40AF' : '#0F766E'}; font-weight:700; border-radius:6px; padding:4px 8px; font-size:0.75rem;">
                        ${r.vendor}
                      </span>
                    </td>
                    <td style="padding:12px 16px; font-size:0.875rem; font-weight:600; color:var(--text-2);">
                      ${MONTH_NAMES[r.target_month] || r.target_month}
                    </td>
                    <td style="padding:12px 16px;">
                      <button class="status-toggle-btn" data-id="${r.id}" data-status="${r.status}" style="background:${st.bg}; color:${st.color}; border:1px solid ${st.border}; padding:4px 8px; border-radius:6px; font-size:0.75rem; font-weight:700; cursor:pointer;" title="Klik untuk ubah status">
                        ${st.icon} ${st.label} ▾
                      </button>
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-2);">
                      ${r.actual_date || '-'}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; font-weight:600; color:var(--text-2);">
                      ${r.employee_count > 0 ? `👤 ${r.employee_count}` : '-'}
                    </td>
                    <td style="padding:12px 16px; font-size:0.85rem; color:var(--text-3); max-width:200px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${r.notes || '-'}
                    </td>
                    ${canEdit ? `
                      <td style="padding:12px 16px; text-align:right;">
                        <button class="btn btn-outline btn-sm edit-card-btn" data-id="${r.id}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px; margin-right:4px;">Edit</button>
                        <button class="btn btn-danger btn-sm delete-card-btn" data-id="${r.id}" data-name="${r.branch_name}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px;">Hapus</button>
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
    // Status quick changer popup / modal
    targetEl.querySelectorAll('.status-toggle-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
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
    // Remove any open popup
    document.querySelectorAll('.status-popup-menu').forEach(el => el.remove());

    const rect = anchorBtn.getBoundingClientRect();
    const popup = document.createElement('div');
    popup.className = 'status-popup-menu';
    popup.style.cssText = `
      position: fixed;
      top: ${rect.bottom + window.scrollY + 4}px;
      left: ${rect.left + window.scrollX}px;
      background: #fff;
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
        color: ${s.key === currentSt ? '#2563EB' : 'var(--text-1)'};
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

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
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

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
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
          <input type="text" name="pic_vendor" class="form-control" placeholder="Nama PIC vendor / no HP" value="${item?.pic_vendor || ''}" style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Catatan / Keterangan</label>
          <textarea name="notes" class="form-control" rows="2" placeholder="Catatan serah terima / info tambahan..." style="width:100%; border-radius:8px;">${item?.notes || ''}</textarea>
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
            const errData = await res.json().catch(() => ({}));
            throw new Error(errData.error || 'Gagal menyimpan data');
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
  document.getElementById('btn-add-migration').addEventListener('click', () => {
    showFormModal();
  });

  // Export Excel button
  document.getElementById('btn-export-migrations').addEventListener('click', () => {
    const rows = getFilteredRows();
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
  });

  // Initial load
  loadData();
}
