import { apiFetch, getUser } from '../config.js';
import { createModal, confirmDialog } from '../components/modal.js';
import { toastSuccess, toastError } from '../components/toast.js';
import { downloadExcel } from '../utils/excel.js';

const CATEGORY_COLORS = {
  'Persiapan': { bg: '#F3E8FF', color: '#6B21A8', border: '#D8B4FE' },
  'Area Umum': { bg: '#E0F2FE', color: '#0369A1', border: '#BAE6FD' },
  'Ruang Tindakan': { bg: '#DBEAFE', color: '#1E40AF', border: '#93C5FD' },
  'Ruang Penunjang': { bg: '#F1F5F9', color: '#334155', border: '#CBD5E1' },
  'Toilet': { bg: '#CCFBF1', color: '#0F766E', border: '#99F6E4' },
  'Area Luar': { bg: '#FEF3C7', color: '#92400E', border: '#FDE68A' },
  'Kontrol & Check': { bg: '#FFEDD5', color: '#C2410C', border: '#FED7AA' },
  'Istirahat': { bg: '#E2E8F0', color: '#475569', border: '#CBD5E1' },
  'Layanan Pasien': { bg: '#DCFCE7', color: '#15803D', border: '#86EFAC' },
  'Sampah Umum': { bg: '#FEF9C3', color: '#854D0E', border: '#FDE047' },
  'Sampah Medis': { bg: '#FEE2E2', color: '#991B1B', border: '#FCA5A5' },
  'Penutup': { bg: '#E0E7FF', color: '#3730A3', border: '#C7D2FE' }
};

export async function renderDAC(container) {
  container.innerHTML = `
    <div class="page-header" style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px; margin-bottom:20px;">
      <div>
        <h1 class="page-title" style="display:flex; align-items:center; gap:8px; margin:0 0 6px 0; font-size:1.6rem; font-weight:700; color:var(--text-1);">
          <span>⏱️</span> Master DAC (Daily Activity Cleaning)
        </h1>
        <p class="page-subtitle" style="margin:0; font-size:0.9rem; color:var(--text-3);">
          Standar Acuan Alur Rutinitas Pembersihan Harian Facility Care (Shift Pagi & Shift Siang)
        </p>
      </div>
      <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
        <div class="btn-group" style="background:#F1F5F9; padding:3px; border-radius:8px; display:inline-flex; border:1px solid var(--border);">
          <button id="view-timeline-btn" class="btn btn-sm" style="background:#fff; font-weight:600; border-radius:6px; box-shadow:0 1px 2px rgba(0,0,0,0.05); padding:6px 12px; cursor:pointer;">
            ⏱️ Alur Timeline
          </button>
          <button id="view-table-btn" class="btn btn-sm btn-ghost" style="font-weight:600; border-radius:6px; padding:6px 12px; cursor:pointer;">
            📋 Tabel Standar
          </button>
        </div>
        <button id="btn-print-dac" class="btn btn-outline" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;" title="Cetak Format SOP untuk Ruang Janitor">
          🖨️ Cetak SOP
        </button>
        <button id="btn-export-dac" class="btn btn-outline" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;">
          📥 Export Excel
        </button>
        <button id="btn-add-dac" class="btn btn-primary" style="font-weight:600; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;">
          + Tambah Aktivitas
        </button>
      </div>
    </div>

    <!-- Shift Switcher Tabs -->
    <div style="display:flex; gap:8px; margin-bottom:20px; border-bottom:2px solid #E2E8F0; padding-bottom:8px;">
      <button class="shift-tab-btn active" data-shift="pagi" style="border:none; background:#2563EB; color:#fff; font-weight:700; padding:8px 18px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem;">
        🌅 Shift Pagi (08:00 - 17:00) <span class="badge" style="background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 6px; border-radius:10px;" id="badge-count-pagi">31</span>
      </button>
      <button class="shift-tab-btn" data-shift="siang" style="border:none; background:#F1F5F9; color:var(--text-2); font-weight:700; padding:8px 18px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem;">
        🌇 Shift Siang (13:00 - 22:00) <span class="badge" style="background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 6px; border-radius:10px;" id="badge-count-siang">33</span>
      </button>
      <button class="shift-tab-btn" data-shift="all" style="border:none; background:#F1F5F9; color:var(--text-2); font-weight:700; padding:8px 18px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.9rem;">
        📑 Semua Shift <span class="badge" style="background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 6px; border-radius:10px;" id="badge-count-all">64</span>
      </button>
    </div>

    <!-- Summary KPI Cards -->
    <div id="dac-summary-cards" style="display:grid; grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); gap:16px; margin-bottom:24px;">
      <div class="kpi-card skeleton" style="height:90px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:90px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:90px; border-radius:12px;"></div>
      <div class="kpi-card skeleton" style="height:90px; border-radius:12px;"></div>
    </div>

    <!-- Filter Toolbar -->
    <div class="card" style="padding:14px 18px; margin-bottom:24px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
      <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
          <div style="position:relative; flex:1; min-width:200px; max-width:320px;">
            <input type="text" id="dac-search" class="form-control" placeholder="Cari lokasi, aktivitas, jam..." style="padding-left:34px; height:38px; border-radius:8px; font-size:0.9rem; width:100%;">
            <span style="position:absolute; left:11px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.9rem;">🔍</span>
          </div>
          <select id="dac-category-filter" class="form-control" style="width:auto; height:38px; border-radius:8px; font-size:0.875rem; min-width:160px;">
            <option value="">Semua Kategori</option>
          </select>
          <select id="dac-period-filter" class="form-control" style="width:auto; height:38px; border-radius:8px; font-size:0.875rem; min-width:150px;">
            <option value="">Semua Waktu</option>
            <option value="morning">Sesi Awal (08:00 - 12:00 / 13:00 - 17:00)</option>
            <option value="afternoon">Sesi Lanjutan (12:00 - 17:00 / 17:00 - 22:00)</option>
          </select>
        </div>
        <div id="dac-stats-text" style="font-size:0.85rem; color:var(--text-3); font-weight:500;">
          Memuat data...
        </div>
      </div>
    </div>

    <!-- Main Content Container -->
    <div id="dac-view-container"></div>
  `;

  let currentShift = 'pagi'; // 'pagi', 'siang', or 'all'
  let currentView = 'timeline'; // 'timeline' or 'table'
  let allItems = [];
  const user = getUser();
  const canEdit = user && ['superadmin', 'admin', 'manager', 'spv'].includes(user.role);

  async function loadData() {
    try {
      const res = await apiFetch('/api/dac?all=1');
      if (!res.ok) throw new Error('Gagal memuat standar DAC');
      allItems = res.data?.data || res.data || [];

      // Update badge counts
      const pagiCount = allItems.filter(r => r.shift === 'pagi').length;
      const siangCount = allItems.filter(r => r.shift === 'siang').length;
      document.getElementById('badge-count-pagi').textContent = pagiCount;
      document.getElementById('badge-count-siang').textContent = siangCount;
      document.getElementById('badge-count-all').textContent = allItems.length;

      // Populate category filter
      const catSelect = document.getElementById('dac-category-filter');
      const uniqueCats = Array.from(new Set(allItems.map(r => r.category))).sort();
      catSelect.innerHTML = '<option value="">Semua Kategori</option>' + uniqueCats.map(c => `<option value="${c}">${c}</option>`).join('');

      renderKPIs();
      renderView();
    } catch (err) {
      document.getElementById('dac-view-container').innerHTML = `
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">⚠️</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data DAC: ${err.message}</p>
          <button class="btn btn-outline btn-sm" onclick="window.location.reload()">Muat Ulang</button>
        </div>
      `;
    }
  }

  function getFilteredItems() {
    const search = (document.getElementById('dac-search')?.value || '').toLowerCase().trim();
    const category = document.getElementById('dac-category-filter')?.value || '';
    const period = document.getElementById('dac-period-filter')?.value || '';

    return allItems.filter(r => {
      if (currentShift !== 'all' && r.shift !== currentShift) return false;
      if (category && r.category !== category) return false;

      if (period === 'morning') {
        const startH = parseInt(r.start_time.split(':')[0], 10);
        if (r.shift === 'pagi' && startH >= 12) return false;
        if (r.shift === 'siang' && startH >= 17) return false;
      } else if (period === 'afternoon') {
        const startH = parseInt(r.start_time.split(':')[0], 10);
        if (r.shift === 'pagi' && startH < 12) return false;
        if (r.shift === 'siang' && startH < 17) return false;
      }

      if (search) {
        const text = `${r.start_time} ${r.end_time} ${r.category} ${r.location} ${r.activity}`.toLowerCase();
        if (!text.includes(search)) return false;
      }
      return true;
    });
  }

  function renderKPIs() {
    const kpiEl = document.getElementById('dac-summary-cards');
    if (!kpiEl) return;

    const pagiItems = allItems.filter(r => r.shift === 'pagi');
    const siangItems = allItems.filter(r => r.shift === 'siang');

    const activeList = currentShift === 'pagi' ? pagiItems : currentShift === 'siang' ? siangItems : allItems;
    const shiftHours = currentShift === 'pagi' ? '08:00 – 17:00' : currentShift === 'siang' ? '13:00 – 22:00' : '08:00 – 22:00';
    const shiftLabel = currentShift === 'pagi' ? 'Shift Pagi' : currentShift === 'siang' ? 'Shift Siang' : 'Semua Shift';

    kpiEl.innerHTML = `
      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:var(--text-3); text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Total Aktivitas (${shiftLabel})</div>
        <div style="font-size:1.8rem; font-weight:800; color:var(--text-1); line-height:1.2;">${activeList.length} <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Aktivitas</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Urutan Terjadwal dari Awal s/d Penutup</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #2563EB; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#1E40AF; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Jam Kerja Standar</div>
        <div style="font-size:1.5rem; font-weight:800; color:#1E40AF; line-height:1.2;">${shiftHours}</div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">9 Jam Operasional (Break 1 Jam)</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #0D9488; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#0F766E; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Cakupan Operasional</div>
        <div style="font-size:1.8rem; font-weight:800; color:#0F766E; line-height:1.2;">12 <span style="font-size:0.9rem; font-weight:500; color:var(--text-3);">Kategori</span></div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Tindakan, Toilet, Area Umum, Medis, dll.</div>
      </div>

      <div class="card" style="padding:16px 20px; background:#fff; border:1px solid var(--border); border-radius:12px; border-left:4px solid #F59E0B; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="font-size:0.8rem; font-weight:600; color:#B45309; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:6px;">Poin Wajib Harian</div>
        <div style="font-size:1.1rem; font-weight:700; color:#B45309; line-height:1.4;">Check Sheet & Sampah Medis</div>
        <div style="font-size:0.75rem; color:var(--text-3); margin-top:4px;">Kontrol mobilitas bertahap & sterilisasi</div>
      </div>
    `;
  }

  function renderView() {
    const items = getFilteredItems();
    const statsEl = document.getElementById('dac-stats-text');
    if (statsEl) {
      statsEl.innerHTML = `Menampilkan <strong>${items.length}</strong> aktivitas`;
    }

    const containerEl = document.getElementById('dac-view-container');
    if (currentView === 'timeline') {
      renderTimelineView(containerEl, items);
    } else {
      renderTableView(containerEl, items);
    }
  }

  function renderTimelineView(targetEl, items) {
    if (!items.length) {
      targetEl.innerHTML = `
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">🔍</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
          <p style="color:var(--text-3); font-size:0.875rem;">Coba reset kata kunci atau pilihan kategori di atas.</p>
        </div>
      `;
      return;
    }

    targetEl.innerHTML = `
      <div class="timeline-container" style="display:flex; flex-direction:column; gap:12px; max-width:960px; margin:0 auto;">
        ${items.map(item => {
          const cfg = CATEGORY_COLORS[item.category] || { bg: '#F1F5F9', color: '#334155', border: '#CBD5E1' };
          const shiftBadge = item.shift === 'pagi' ? '🌅 Pagi' : '🌇 Siang';
          const shiftBadgeBg = item.shift === 'pagi' ? '#FEF3C7' : '#E0E7FF';
          const shiftBadgeColor = item.shift === 'pagi' ? '#92400E' : '#3730A3';

          return `
            <div class="dac-timeline-card" style="background:#fff; border:1px solid var(--border); border-radius:12px; padding:16px 20px; box-shadow:0 1px 3px rgba(0,0,0,0.03); display:flex; gap:16px; align-items:flex-start; position:relative;">
              <!-- Time Column -->
              <div style="min-width:130px; text-align:center; background:#F8FAFC; border:1px solid var(--border); border-radius:8px; padding:10px 8px;">
                <div style="font-size:1.1rem; font-weight:800; color:var(--text-1); font-family:monospace;">
                  ${item.start_time} - ${item.end_time}
                </div>
                <div style="font-size:0.75rem; color:var(--text-3); font-weight:600; margin-top:4px;">
                  Item #${item.item_no}
                </div>
                ${currentShift === 'all' ? `
                  <div style="margin-top:6px;">
                    <span class="badge" style="background:${shiftBadgeBg}; color:${shiftBadgeColor}; font-size:0.7rem; font-weight:700; padding:2px 6px; border-radius:4px;">
                      ${shiftBadge}
                    </span>
                  </div>
                ` : ''}
              </div>

              <!-- Content Column -->
              <div style="flex:1;">
                <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:8px; margin-bottom:6px;">
                  <div style="display:flex; align-items:center; gap:8px; flex-wrap:wrap;">
                    <span class="badge" style="background:${cfg.bg}; color:${cfg.color}; border:1px solid ${cfg.border}; font-weight:700; font-size:0.75rem; padding:3px 10px; border-radius:6px;">
                      ${item.category}
                    </span>
                    <span style="font-weight:700; font-size:0.95rem; color:var(--text-1);">
                      📍 ${item.location}
                    </span>
                  </div>
                  ${canEdit ? `
                    <div style="display:flex; gap:4px;">
                      <button class="btn btn-ghost btn-sm edit-dac-btn" data-id="${item.id}" title="Edit" style="padding:4px 8px; font-size:0.8rem; cursor:pointer;">✏️</button>
                      <button class="btn btn-ghost btn-sm delete-dac-btn" data-id="${item.id}" data-name="${item.location}" title="Hapus" style="padding:4px 8px; font-size:0.8rem; color:var(--danger); cursor:pointer;">🗑️</button>
                    </div>
                  ` : ''}
                </div>

                <div style="font-size:0.9rem; color:var(--text-2); line-height:1.5; background:#FAFBFD; padding:10px 14px; border-radius:8px; border-left:3px solid ${cfg.color}; margin-top:8px;">
                  ${item.activity}
                </div>
              </div>
            </div>
          `;
        }).join('')}
      </div>
    `;
    bindActionEvents(targetEl);
  }

  function renderTableView(targetEl, items) {
    if (!items.length) {
      targetEl.innerHTML = `
        <div class="empty-state" style="padding:50px; text-align:center; background:#fff; border-radius:12px; border:1px solid var(--border);">
          <div style="font-size:2.5rem; margin-bottom:12px;">🔍</div>
          <p style="font-size:1.05rem; font-weight:600; color:var(--text-2); margin-bottom:6px;">Tidak ada aktivitas yang cocok dengan filter</p>
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
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:50px;">NO</th>
                ${currentShift === 'all' ? `<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3);">SHIFT</th>` : ''}
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:110px;">JAM</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:140px;">KATEGORI</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); width:200px;">LOKASI / AREA</th>
                <th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3);">AKTIVITAS</th>
                ${canEdit ? `<th style="padding:12px 14px; font-size:0.8rem; font-weight:700; color:var(--text-3); text-align:right; width:100px;">AKSI</th>` : ''}
              </tr>
            </thead>
            <tbody>
              ${items.map(item => {
                const cfg = CATEGORY_COLORS[item.category] || { bg: '#F1F5F9', color: '#334155', border: '#CBD5E1' };
                return `
                  <tr style="border-bottom:1px solid var(--border);">
                    <td style="padding:12px 14px; font-size:0.85rem; color:var(--text-3); font-weight:600;">${item.item_no}</td>
                    ${currentShift === 'all' ? `
                      <td style="padding:12px 14px;">
                        <span class="badge" style="background:${item.shift === 'pagi' ? '#FEF3C7' : '#E0E7FF'}; color:${item.shift === 'pagi' ? '#92400E' : '#3730A3'}; font-size:0.75rem; font-weight:700; padding:3px 8px; border-radius:4px;">
                          ${item.shift === 'pagi' ? 'Pagi' : 'Siang'}
                        </span>
                      </td>
                    ` : ''}
                    <td style="padding:12px 14px; font-family:monospace; font-weight:700; font-size:0.85rem; color:var(--text-1); white-space:nowrap;">
                      ${item.start_time} - ${item.end_time}
                    </td>
                    <td style="padding:12px 14px;">
                      <span class="badge" style="background:${cfg.bg}; color:${cfg.color}; border:1px solid ${cfg.border}; font-weight:700; font-size:0.75rem; padding:3px 8px; border-radius:6px; white-space:nowrap;">
                        ${item.category}
                      </span>
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; font-weight:600; color:var(--text-1);">
                      ${item.location}
                    </td>
                    <td style="padding:12px 14px; font-size:0.875rem; color:var(--text-2); line-height:1.4;">
                      ${item.activity}
                    </td>
                    ${canEdit ? `
                      <td style="padding:12px 14px; text-align:right; white-space:nowrap;">
                        <button class="btn btn-outline btn-sm edit-dac-btn" data-id="${item.id}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px; margin-right:4px;">Edit</button>
                        <button class="btn btn-danger btn-sm delete-dac-btn" data-id="${item.id}" data-name="${item.location}" style="padding:3px 8px; font-size:0.75rem; border-radius:6px;">Hapus</button>
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
    bindActionEvents(targetEl);
  }

  function bindActionEvents(targetEl) {
    // Edit item
    targetEl.querySelectorAll('.edit-dac-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.dataset.id;
        const item = allItems.find(r => String(r.id) === String(id));
        if (item) showFormModal(item);
      });
    });

    // Delete item
    targetEl.querySelectorAll('.delete-dac-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.dataset.id;
        const name = btn.dataset.name || 'item';
        confirmDialog(`Apakah Anda yakin ingin menghapus aktivitas pada lokasi <strong>${name}</strong>?`, async () => {
          try {
            const res = await apiFetch(`/api/dac/${id}`, { method: 'DELETE' });
            if (!res.ok) throw new Error('Gagal menghapus item');
            toastSuccess('Aktivitas DAC berhasil dihapus');
            loadData();
          } catch (err) {
            toastError(err.message);
          }
        });
      });
    });
  }

  function showFormModal(item = null) {
    const isEdit = !!item;
    const title = isEdit ? 'Edit Aktivitas DAC' : 'Tambah Aktivitas DAC';

    const defaultCategories = [
      'Persiapan', 'Area Umum', 'Ruang Tindakan', 'Ruang Penunjang',
      'Toilet', 'Area Luar', 'Kontrol & Check', 'Istirahat',
      'Layanan Pasien', 'Sampah Umum', 'Sampah Medis', 'Penutup'
    ];

    const contentHtml = `
      <form id="dac-form" style="display:flex; flex-direction:column; gap:14px;">
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Shift *</label>
            <select name="shift" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="pagi" ${item?.shift === 'pagi' || (!item && currentShift === 'pagi') ? 'selected' : ''}>🌅 Shift Pagi (08:00 - 17:00)</option>
              <option value="siang" ${item?.shift === 'siang' || (!item && currentShift === 'siang') ? 'selected' : ''}>🌇 Shift Siang (13:00 - 22:00)</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Nomor Urut *</label>
            <input type="number" name="item_no" class="form-control" value="${item?.item_no || ''}" placeholder="Otomatis jika kosong" style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Jam Mulai (HH:MM) *</label>
            <input type="time" name="start_time" class="form-control" value="${item?.start_time || ''}" required style="width:100%; height:38px; border-radius:8px;">
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Jam Selesai (HH:MM) *</label>
            <input type="time" name="end_time" class="form-control" value="${item?.end_time || ''}" required style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Kategori *</label>
          <input list="cat-suggestions" name="category" class="form-control" value="${item?.category || ''}" required placeholder="Pilih atau ketik kategori..." style="width:100%; height:38px; border-radius:8px;">
          <datalist id="cat-suggestions">
            ${defaultCategories.map(c => `<option value="${c}">`).join('')}
          </datalist>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Lokasi / Area *</label>
          <input type="text" name="location" class="form-control" value="${item?.location || ''}" required placeholder="Contoh: Ruang tindakan Du 101, Lobby, Toilet GF" style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Aktivitas Detail *</label>
          <textarea name="activity" class="form-control" rows="3" required placeholder="Deskripsi aktivitas pembersihan..." style="width:100%; border-radius:8px;">${item?.activity || ''}</textarea>
        </div>
      </form>
    `;

    createModal({
      title,
      content: contentHtml,
      confirmText: isEdit ? 'Simpan Perubahan' : 'Tambah Aktivitas',
      onConfirm: async (overlay, close) => {
        const form = overlay.querySelector('#dac-form');
        if (!form.checkValidity()) {
          form.reportValidity();
          return;
        }

        const formData = new FormData(form);
        const payload = {
          shift: formData.get('shift'),
          item_no: formData.get('item_no') ? parseInt(formData.get('item_no'), 10) : undefined,
          start_time: formData.get('start_time'),
          end_time: formData.get('end_time'),
          category: formData.get('category'),
          location: formData.get('location'),
          activity: formData.get('activity')
        };

        try {
          const url = isEdit ? `/api/dac/${item.id}` : '/api/dac';
          const method = isEdit ? 'PUT' : 'POST';
          const res = await apiFetch(url, { method, body: JSON.stringify(payload) });

          if (!res.ok) {
            const errData = await res.json().catch(() => ({}));
            throw new Error(errData.error || 'Gagal menyimpan aktivitas');
          }

          toastSuccess(isEdit ? 'Aktivitas berhasil diperbarui' : 'Aktivitas berhasil ditambahkan');
          close();
          loadData();
        } catch (err) {
          toastError(err.message);
        }
      }
    });
  }

  function printSOP() {
    const items = getFilteredItems();
    if (!items.length) {
      toastError('Tidak ada data untuk dicetak');
      return;
    }

    const shiftTitle = currentShift === 'pagi' ? 'SHIFT PAGI (08:00 - 17:00)' : currentShift === 'siang' ? 'SHIFT SIANG (13:00 - 22:00)' : 'SELURUH SHIFT (PAGI & SIANG)';

    const printWin = window.open('', '_blank');
    if (!printWin) {
      toastError('Popup diblokir browser, mohon izinkan popup untuk mencetak');
      return;
    }

    printWin.document.write(`
      <!DOCTYPE html>
      <html>
      <head>
        <title>STANDAR ACUAN DAC - ${shiftTitle}</title>
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; margin: 20px; color: #1e293b; font-size: 11pt; }
          .header { text-align: center; border-bottom: 2px solid #0f172a; padding-bottom: 12px; margin-bottom: 20px; }
          .title { font-size: 16pt; font-weight: bold; margin: 0; color: #0f172a; }
          .subtitle { font-size: 11pt; color: #475569; margin: 4px 0 0 0; }
          table { width: 100%; border-collapse: collapse; margin-top: 10px; }
          th, td { border: 1px solid #cbd5e1; padding: 7px 10px; text-align: left; vertical-align: top; }
          th { background-color: #f1f5f9; font-weight: bold; font-size: 10pt; text-transform: uppercase; }
          .jam { font-family: monospace; font-weight: bold; white-space: nowrap; }
          .badge { display: inline-block; padding: 2px 6px; border-radius: 4px; font-size: 8pt; font-weight: bold; }
          @media print {
            body { margin: 10mm; }
            button { display: none; }
          }
        </style>
      </head>
      <body>
        <div class="header">
          <div class="title">FACILITY CARE MANAGEMENT SYSTEM (FCMS)</div>
          <div class="subtitle">STANDAR RUTINITAS DAILY ACTIVITY CLEANING (DAC) — ${shiftTitle}</div>
        </div>
        <table>
          <thead>
            <tr>
              <th style="width:35px; text-align:center;">NO</th>
              ${currentShift === 'all' ? '<th style="width:55px;">SHIFT</th>' : ''}
              <th style="width:105px;">JAM</th>
              <th style="width:130px;">KATEGORI</th>
              <th style="width:180px;">LOKASI / AREA</th>
              <th>RINCIAN AKTIVITAS</th>
            </tr>
          </thead>
          <tbody>
            ${items.map(r => `
              <tr>
                <td style="text-align:center; font-weight:bold;">${r.item_no}</td>
                ${currentShift === 'all' ? `<td>${r.shift === 'pagi' ? 'Pagi' : 'Siang'}</td>` : ''}
                <td class="jam">${r.start_time} - ${r.end_time}</td>
                <td><strong>${r.category}</strong></td>
                <td><strong>${r.location}</strong></td>
                <td>${r.activity}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
        <div style="margin-top:20px; text-align:right; font-size:9pt; color:#64748b;">
          Dicetak pada: ${new Date().toLocaleString('id-ID')} | Standar Resmi Operasional Facility Care Audy Dental
        </div>
      </body>
      </html>
    `);
    printWin.document.close();
    printWin.focus();
    setTimeout(() => {
      printWin.print();
    }, 400);
  }

  // Shift tab switchers
  document.querySelectorAll('.shift-tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.shift-tab-btn').forEach(b => {
        b.style.background = '#F1F5F9';
        b.style.color = 'var(--text-2)';
      });
      btn.style.background = '#2563EB';
      btn.style.color = '#fff';

      currentShift = btn.dataset.shift;
      renderKPIs();
      renderView();
    });
  });

  // View switchers
  const timelineBtn = document.getElementById('view-timeline-btn');
  const tableBtn = document.getElementById('view-table-btn');

  timelineBtn.addEventListener('click', () => {
    currentView = 'timeline';
    timelineBtn.style.background = '#fff';
    timelineBtn.classList.remove('btn-ghost');
    tableBtn.style.background = 'transparent';
    tableBtn.classList.add('btn-ghost');
    renderView();
  });

  tableBtn.addEventListener('click', () => {
    currentView = 'table';
    tableBtn.style.background = '#fff';
    tableBtn.classList.remove('btn-ghost');
    timelineBtn.style.background = 'transparent';
    timelineBtn.classList.add('btn-ghost');
    renderView();
  });

  // Filter events
  ['dac-search', 'dac-category-filter', 'dac-period-filter'].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('input', renderView);
      el.addEventListener('change', renderView);
    }
  });

  // Add button
  document.getElementById('btn-add-dac').addEventListener('click', () => {
    showFormModal();
  });

  // Print button
  document.getElementById('btn-print-dac').addEventListener('click', () => {
    printSOP();
  });

  // Export Excel
  document.getElementById('btn-export-dac').addEventListener('click', () => {
    const items = getFilteredItems();
    if (!items.length) {
      toastError('Tidak ada data untuk diexport');
      return;
    }
    const exportData = items.map(r => ({
      'No': r.item_no,
      'Shift': r.shift === 'pagi' ? 'Shift Pagi' : 'Shift Siang',
      'Jam Mulai': r.start_time,
      'Jam Selesai': r.end_time,
      'Kategori': r.category,
      'Lokasi / Area': r.location,
      'Rincian Aktivitas': r.activity
    }));
    downloadExcel(exportData, `DAC_Daily_Activity_Cleaning_${currentShift}_${new Date().toISOString().slice(0, 10)}`);
    toastSuccess('File Excel berhasil diunduh');
  });

  // Load initial data
  loadData();
}
