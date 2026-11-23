import { apiFetch, getUser, hasPermission } from '../config.js';
import { createModal, confirmDialog } from '../components/modal.js';
import { toastSuccess, toastError } from '../components/toast.js';
import { downloadExcel } from '../utils/excel.js';
import { fetchExportRows } from '../utils/dataCache.js';
import { escapeHtml } from '../utils/sanitize.js';

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
  const user = getUser();
  const canEdit = hasPermission(user, 'dac', 'write');
  const canDelete = hasPermission(user, 'dac', 'delete');
  const canExport = hasPermission(user, 'export');

  container.innerHTML = `
    <div class="dac-page-shell">
      <!-- Unified Hero Card (Consistent with FCMS Clean Navy Modules) -->
      <div class="dac-hero-card">
        <div class="dac-hero-top">
          <div class="dac-hero-title-group">
            <h1 class="dac-hero-title">Master DAC (Daily Activity Cleaning)</h1>
            <p class="dac-hero-subtitle">Standar acuan alur rutinitas pembersihan harian Facility Care (Shift Pagi &amp; Shift Siang)</p>
          </div>
          <div class="dac-hero-actions">
            <div class="dac-view-switch" role="group" aria-label="Mode tampilan Master DAC">
              <button type="button" id="view-timeline-btn" class="dac-view-btn active">Alur Timeline</button>
              <button type="button" id="view-table-btn" class="dac-view-btn">Tabel Standar</button>
            </div>
            <button type="button" id="btn-print-dac" class="btn btn-outline btn-sm" title="Cetak Format SOP untuk Ruang Janitor">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg>
              Cetak SOP
            </button>
            ${canExport ? `
            <button type="button" id="btn-export-dac" class="btn btn-outline btn-sm">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
              Export Excel
            </button>
            ` : ''}
            ${canEdit ? `
            <button type="button" id="btn-add-dac" class="btn btn-primary btn-sm">
              + Tambah Aktivitas
            </button>
            ` : ''}
          </div>
        </div>

        <!-- Unified Summary Stats Grid -->
        <div id="dac-summary-cards" class="dac-stats-grid">
          <div class="dac-stat-box skeleton" style="height:78px;"></div>
          <div class="dac-stat-box skeleton" style="height:78px;"></div>
          <div class="dac-stat-box skeleton" style="height:78px;"></div>
          <div class="dac-stat-box skeleton" style="height:78px;"></div>
        </div>
      </div>

      <!-- Unified Filter Toolbar & Shift Switcher Tabs -->
      <div class="dac-controls-card">
        <div class="dac-controls-row">
          <div class="dac-search-wrap">
            <svg class="dac-search-icon" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
            <input type="text" id="dac-search" class="form-control dac-search-input" placeholder="Cari lokasi, aktivitas, jam..." aria-label="Cari aktivitas DAC">
          </div>
          <select id="dac-category-filter" class="form-control dac-filter-select" aria-label="Filter kategori DAC">
            <option value="">Semua Kategori</option>
          </select>
          <select id="dac-period-filter" class="form-control dac-filter-select" aria-label="Filter periode waktu DAC">
            <option value="">Semua Waktu</option>
            <option value="morning">Sesi Awal (08:00 - 12:00 / 13:00 - 17:00)</option>
            <option value="afternoon">Sesi Lanjutan (12:00 - 17:00 / 17:00 - 22:00)</option>
          </select>
          <div id="dac-stats-text" class="dac-stats-text">
            Memuat data...
          </div>
        </div>

        <div id="dac-shift-tabs" class="dac-shift-tabs" data-fc-scroll-tabs role="tablist" aria-label="Filter shift DAC">
          <button type="button" class="shift-tab-btn active" data-shift="pagi" role="tab" aria-selected="true">
            <span>Shift Pagi (08:00 - 17:00)</span>
            <span class="badge dac-tab-count" id="badge-count-pagi">31</span>
          </button>
          <button type="button" class="shift-tab-btn" data-shift="siang" role="tab" aria-selected="false">
            <span>Shift Siang (13:00 - 22:00)</span>
            <span class="badge dac-tab-count" id="badge-count-siang">33</span>
          </button>
          <button type="button" class="shift-tab-btn" data-shift="all" role="tab" aria-selected="false">
            <span>Semua Shift</span>
            <span class="badge dac-tab-count" id="badge-count-all">64</span>
          </button>
        </div>
      </div>

      <!-- Main Content Container (Aligned 100% Width with Shell) -->
      <div id="dac-view-container" class="dac-view-container"></div>
    </div>
  `;

  let currentShift = 'pagi'; // 'pagi', 'siang', or 'all'
  let currentView = 'timeline'; // 'timeline' or 'table'
  let allItems = [];
  // canEdit defined at top

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
      catSelect.innerHTML = '<option value="">Semua Kategori</option>' + uniqueCats.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');

      renderKPIs();
      renderView();
    } catch (err) {
      const vc = document.getElementById('dac-view-container');
      vc.innerHTML = `
        <div class="empty-state" style="padding:40px; text-align:center;">
          <div style="font-size:2rem; margin-bottom:8px;">⚠️</div>
          <p style="color:var(--danger); font-weight:600;">Gagal memuat data DAC: ${escapeHtml(err.message)}</p>
          <button type="button" id="dac-reload-btn" class="btn btn-outline btn-sm">Muat Ulang</button>
        </div>
      `;
      vc.querySelector('#dac-reload-btn')?.addEventListener('click', () => window.location.reload());
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
      <div class="dac-stat-box">
        <div class="dac-stat-icon" aria-hidden="true">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg>
        </div>
        <div class="dac-stat-body">
          <div class="dac-stat-val">${activeList.length}</div>
          <div class="dac-stat-lbl">Total Aktivitas (${shiftLabel})</div>
        </div>
      </div>

      <div class="dac-stat-box">
        <div class="dac-stat-icon" aria-hidden="true">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
        </div>
        <div class="dac-stat-body">
          <div class="dac-stat-val dac-stat-val--time">${shiftHours}</div>
          <div class="dac-stat-lbl">Jam Kerja Standar</div>
        </div>
      </div>

      <div class="dac-stat-box">
        <div class="dac-stat-icon" aria-hidden="true">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>
        </div>
        <div class="dac-stat-body">
          <div class="dac-stat-val">12</div>
          <div class="dac-stat-lbl">Cakupan Kategori</div>
        </div>
      </div>

      <div class="dac-stat-box">
        <div class="dac-stat-icon" aria-hidden="true">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>
        </div>
        <div class="dac-stat-body">
          <div class="dac-stat-val dac-stat-val--focus">Check Sheet &amp; Medis</div>
          <div class="dac-stat-lbl">Poin Wajib Harian</div>
        </div>
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
        <div class="dac-empty-card">
          <p class="dac-empty-title">Tidak ada aktivitas yang cocok dengan filter</p>
          <p class="dac-empty-sub">Coba reset kata kunci atau pilihan kategori di atas.</p>
        </div>
      `;
      return;
    }

    targetEl.innerHTML = `
      <div class="timeline-container dac-timeline-list">
        ${items.map(item => {
          const cfg = CATEGORY_COLORS[item.category] || { bg: 'rgba(148,163,184,0.14)', color: '#E2E8F0', border: 'rgba(148,163,184,0.3)' };
          const shiftBadge = item.shift === 'pagi' ? 'Pagi' : 'Siang';

          return `
            <div class="dac-timeline-card">
              <!-- Time Column -->
              <div class="dac-time-box">
                <div class="dac-time-range">
                  ${escapeHtml(item.start_time)} - ${escapeHtml(item.end_time)}
                </div>
                <div class="dac-time-itemno">
                  Item #${escapeHtml(String(item.item_no))}
                </div>
                ${currentShift === 'all' ? `
                  <div class="dac-time-shift">
                    <span class="badge ${item.shift === 'pagi' ? 'badge-warning' : 'badge-primary'}">${shiftBadge}</span>
                  </div>
                ` : ''}
              </div>

              <!-- Content Column -->
              <div class="dac-card-content">
                <div class="dac-card-head">
                  <div class="dac-card-meta">
                    <span class="badge" style="background:${cfg.bg}; color:${cfg.color}; border:1px solid ${cfg.border};">
                      ${escapeHtml(item.category)}
                    </span>
                    <span class="dac-card-location">
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/></svg>
                      <span>${escapeHtml(item.location)}</span>
                    </span>
                  </div>
                  ${canEdit || canDelete ? `
                    <div class="dac-card-actions">
                      ${canEdit ? `<button type="button" class="fc-action-btn fc-action-edit edit-dac-btn" data-id="${escapeHtml(String(item.id))}" title="Edit ${escapeHtml(item.location)}" aria-label="Edit aktivitas ${escapeHtml(item.location)}"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg></button>` : ''}
                      ${canDelete ? `<button type="button" class="fc-action-btn fc-action-del delete-dac-btn" data-id="${escapeHtml(String(item.id))}" data-name="${escapeHtml(item.location)}" title="Hapus ${escapeHtml(item.location)}" aria-label="Hapus aktivitas ${escapeHtml(item.location)}"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg></button>` : ''}
                    </div>
                  ` : ''}
                </div>

                <div class="dac-activity-desc">
                  ${escapeHtml(item.activity)}
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
        <div class="dac-empty-card">
          <p class="dac-empty-title">Tidak ada aktivitas yang cocok dengan filter</p>
          <p class="dac-empty-sub">Coba reset kata kunci atau pilihan kategori di atas.</p>
        </div>
      `;
      return;
    }

    targetEl.innerHTML = `
      <div class="dac-table-card">
        <div class="table-wrapper">
          <table class="dac-table">
            <thead>
              <tr>
                <th style="width:52px; text-align:center;">No</th>
                ${currentShift === 'all' ? '<th style="width:82px;">Shift</th>' : ''}
                <th style="width:128px;">Jam</th>
                <th style="width:150px;">Kategori</th>
                <th style="width:210px;">Lokasi / Area</th>
                <th>Aktivitas</th>
                ${canEdit || canDelete ? '<th style="text-align:center; width:96px;">Aksi</th>' : ''}
              </tr>
            </thead>
            <tbody>
              ${items.map(item => {
                const cfg = CATEGORY_COLORS[item.category] || { bg: 'rgba(148,163,184,0.14)', color: '#E2E8F0', border: 'rgba(148,163,184,0.3)' };
                return `
                  <tr class="dac-table-row">
                    <td style="text-align:center; color:#A5B4C8; font-weight:600; font-variant-numeric:tabular-nums;">${escapeHtml(String(item.item_no))}</td>
                    ${currentShift === 'all' ? `
                      <td>
                        <span class="badge ${item.shift === 'pagi' ? 'badge-warning' : 'badge-primary'}">
                          ${item.shift === 'pagi' ? 'Pagi' : 'Siang'}
                        </span>
                      </td>
                    ` : ''}
                    <td style="font-family:ui-monospace,SFMono-Regular,Menlo,monospace; font-weight:700; color:#FFFFFF; white-space:nowrap;">
                      ${escapeHtml(item.start_time)} - ${escapeHtml(item.end_time)}
                    </td>
                    <td>
                      <span class="badge" style="background:${cfg.bg}; color:${cfg.color}; border:1px solid ${cfg.border}; white-space:nowrap;">
                        ${escapeHtml(item.category)}
                      </span>
                    </td>
                    <td style="font-weight:700; color:#F8FAFC;">
                      ${escapeHtml(item.location)}
                    </td>
                    <td style="color:#E2E8F0; line-height:1.5; white-space:normal; min-width:240px;">
                      ${escapeHtml(item.activity)}
                    </td>
                    ${canEdit || canDelete ? `
                      <td style="text-align:center; white-space:nowrap;">
                        <div class="dac-card-actions" style="justify-content:center;">
                          ${canEdit ? `<button type="button" class="fc-action-btn fc-action-edit edit-dac-btn" data-id="${escapeHtml(String(item.id))}" title="Edit ${escapeHtml(item.location)}" aria-label="Edit aktivitas ${escapeHtml(item.location)}"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg></button>` : ''}
                          ${canDelete ? `<button type="button" class="fc-action-btn fc-action-del delete-dac-btn" data-id="${escapeHtml(String(item.id))}" data-name="${escapeHtml(item.location)}" title="Hapus ${escapeHtml(item.location)}" aria-label="Hapus aktivitas ${escapeHtml(item.location)}"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg></button>` : ''}
                        </div>
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
    if (!canEdit) {
      toastError('Anda tidak memiliki izin untuk menambah atau mengedit aktivitas.');
      return;
    }
    const isEdit = !!item;
    const title = isEdit ? 'Edit Aktivitas DAC' : 'Tambah Aktivitas DAC';

    const defaultCategories = [
      'Persiapan', 'Area Umum', 'Ruang Tindakan', 'Ruang Penunjang',
      'Toilet', 'Area Luar', 'Kontrol & Check', 'Istirahat',
      'Layanan Pasien', 'Sampah Umum', 'Sampah Medis', 'Penutup'
    ];

    const contentHtml = `
      <form id="dac-form" style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Pilih Shift <span class="required">*</span></label>
            <select name="shift" class="form-control" required style="width:100%; height:38px; border-radius:8px;">
              <option value="pagi" ${item?.shift === 'pagi' || (!item && currentShift === 'pagi') ? 'selected' : ''}>🌅 Shift Pagi (08:00 - 17:00)</option>
              <option value="siang" ${item?.shift === 'siang' || (!item && currentShift === 'siang') ? 'selected' : ''}>🌇 Shift Siang (13:00 - 22:00)</option>
            </select>
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Nomor Urut <span class="required">*</span></label>
            <input type="number" name="item_no" class="form-control" value="${item?.item_no || ''}" style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Jam Mulai (HH:MM) <span class="required">*</span></label>
            <input type="time" name="start_time" class="form-control" value="${item?.start_time || ''}" required style="width:100%; height:38px; border-radius:8px;">
          </div>

          <div class="form-group">
            <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Jam Selesai (HH:MM) <span class="required">*</span></label>
            <input type="time" name="end_time" class="form-control" value="${item?.end_time || ''}" required style="width:100%; height:38px; border-radius:8px;">
          </div>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Kategori <span class="required">*</span></label>
          <input list="cat-suggestions" name="category" class="form-control" value="${item?.category || ''}" required style="width:100%; height:38px; border-radius:8px;">
          <datalist id="cat-suggestions">
            ${defaultCategories.map(c => `<option value="${c}">`).join('')}
          </datalist>
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Lokasi / Area <span class="required">*</span></label>
          <input type="text" name="location" class="form-control" value="${item?.location || ''}" required style="width:100%; height:38px; border-radius:8px;">
        </div>

        <div class="form-group">
          <label style="font-weight:600; font-size:0.875rem; margin-bottom:4px; display:block;">Aktivitas Detail <span class="required">*</span></label>
          <textarea name="activity" class="form-control" rows="3" required style="width:100%; border-radius:8px;">${item?.activity || ''}</textarea>
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
            const errMsg = res.data?.error || res.data?.message || 'Gagal menyimpan aktivitas';
            throw new Error(errMsg);
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

  // Shift tab switchers (Clean Navy .active state)
  document.querySelectorAll('.shift-tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      currentShift = btn.dataset.shift;
      document.querySelectorAll('.shift-tab-btn').forEach(b => {
        const isActive = b.dataset.shift === currentShift;
        b.classList.toggle('active', isActive);
        b.setAttribute('aria-selected', isActive ? 'true' : 'false');
      });
      renderKPIs();
      renderView();
    });
  });

  // View switchers (Clean Navy .active state)
  const timelineBtn = document.getElementById('view-timeline-btn');
  const tableBtn = document.getElementById('view-table-btn');

  timelineBtn?.addEventListener('click', () => {
    currentView = 'timeline';
    timelineBtn.classList.add('active');
    tableBtn?.classList.remove('active');
    renderView();
  });

  tableBtn?.addEventListener('click', () => {
    currentView = 'table';
    tableBtn.classList.add('active');
    timelineBtn?.classList.remove('active');
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
  document.getElementById('btn-add-dac')?.addEventListener('click', () => {
    showFormModal();
  });

  // Print button
  document.getElementById('btn-print-dac').addEventListener('click', () => {
    printSOP();
  });

  // Export Excel
  document.getElementById('btn-export-dac')?.addEventListener('click', async () => {
    try {
      const search = (document.getElementById('dac-search')?.value || '').trim();
      const category = document.getElementById('dac-category-filter')?.value || '';
      const period = document.getElementById('dac-period-filter')?.value || '';

      const params = {};
      if (currentShift !== 'all') params.shift = currentShift;
      if (category) params.category = category;
      if (search) params.search = search;

      let items = await fetchExportRows('/api/dac/export', params);
      if (period === 'morning' || period === 'afternoon') {
        items = items.filter(r => {
          const startH = parseInt(String(r.start_time || '0').split(':')[0], 10);
          if (period === 'morning') {
            if (r.shift === 'pagi' && startH >= 12) return false;
            if (r.shift === 'siang' && startH >= 17) return false;
          } else {
            if (r.shift === 'pagi' && startH < 12) return false;
            if (r.shift === 'siang' && startH < 17) return false;
          }
          return true;
        });
      }

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
    } catch (err) {
      toastError('Gagal mengekspor data DAC: ' + err.message);
    }
  });

  // Load initial data
  loadData();
}
