import { apiFetch, getUser, hasPermission } from '../config.js';
import { createModal, confirmDialog } from '../components/modal.js';
import { toastSuccess, toastError } from '../components/toast.js';
import { downloadExcel } from '../utils/excel.js';
import { fetchExportRows } from '../utils/dataCache.js';
import { escapeHtml } from '../utils/sanitize.js';

function renderHygieneActionButtons(item, canEdit, canDelete, compact = false, mobileCard = false) {
  if (!canEdit && !canDelete) return '';
  const id = escapeHtml(String(item.id ?? ''));
  const name = escapeHtml(item.item_name || 'Parameter Standar');
  const size = compact ? 14 : 15;
  const clsSuffix = mobileCard
    ? ' hyg-mobile-action-btn'
    : (compact ? ' hygiene-action-btn--compact' : '');
  const btnWrapClass = mobileCard ? 'hyg-mobile-actions' : '';
  return `
    <div class="${btnWrapClass}" style="${mobileCard ? '' : `display:inline-flex; gap:${compact ? 4 : 6}px; align-items:center;`}">
      ${canEdit ? `
        <button type="button" class="hygiene-action-btn hygiene-action-edit${clsSuffix} btn-edit-item" data-id="${id}" title="Edit ${name}" aria-label="Edit parameter ${name}">
          <svg width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <path d="M12 20h9"></path>
            <path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path>
          </svg>
          ${mobileCard ? '<span>Edit</span>' : ''}
        </button>
      ` : ''}
      ${canDelete ? `
        <button type="button" class="hygiene-action-btn hygiene-action-delete${clsSuffix} btn-delete-item" data-id="${id}" data-name="${name}" title="Hapus ${name}" aria-label="Hapus parameter ${name}">
          <svg width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
            <polyline points="3 6 5 6 21 6"></polyline>
            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
            <line x1="10" y1="11" x2="10" y2="17"></line>
            <line x1="14" y1="11" x2="14" y2="17"></line>
          </svg>
          ${mobileCard ? '<span>Hapus</span>' : ''}
        </button>
      ` : ''}
    </div>
  `;
}

function renderHygieneMobileCards(items, { includeRoom = true, canEdit = false, canDelete = false } = {}) {
  return `
    <div class="hyg-mobile-cards${includeRoom ? '' : ' hyg-mobile-cards--compact'}">
      ${items.map((item, idx) => {
        const zInfo = getZoneForRoom(item.room_name);
        const actionsHtml = renderHygieneActionButtons(item, canEdit, canDelete, false, true);
        return `
          <article class="hyg-mobile-card">
            <div class="hyg-mobile-card-head">
              <span class="hyg-mobile-num">#${idx + 1}</span>
              ${includeRoom ? `
                <div class="hyg-mobile-room-group">
                  <span class="hyg-room-badge">${escapeHtml(item.room_name)}</span>
                  <span class="hyg-zone-meta">${escapeHtml(zInfo.zoneName)}</span>
                </div>
              ` : ''}
            </div>
            <div class="hyg-mobile-item-title">${escapeHtml(item.item_name)}</div>
            <div class="hyg-std-box hyg-mobile-std-box">
              <span class="hyg-std-check" aria-hidden="true">✓</span>
              <span class="hyg-std-text">${escapeHtml(item.cleanliness_standard)}</span>
            </div>
            ${actionsHtml ? `<div class="hyg-mobile-card-foot">${actionsHtml}</div>` : ''}
          </article>
        `;
      }).join('')}
    </div>
  `;
}

// 8 Kategori Zona Standar Kebersihan Fisik FCMS (Audy Dental)
export const ZONES = {
  'Standar Umum': {
    icon: '📌',
    bg: '#F3E8FF',
    color: '#6B21A8',
    border: '#D8B4FE',
    rooms: [
      'Lantai & Sudut Ruangan',
      'Dinding, Partisi & Plint',
      'Kaca, Cermin & Partisi Kaca',
      'Sanitair & Area Basah Toilet',
      'Plafon, Lampu & Kisi Ventilasi',
      'Mebel, Meja & Kursi Kerja/Pasien',
      'Peralatan Khusus & Medis Non-Steril',
      'Wadah Sampah & Utilitas'
    ]
  },
  'Area Depan & Pelayanan': {
    icon: '🛋️',
    bg: '#E0F2FE',
    color: '#0369A1',
    border: '#BAE6FD',
    rooms: ['Teras', 'Lobby 1', 'Receptionist', 'Snack Corner']
  },
  'Area Klinis & Medis': {
    icon: '🩺',
    bg: '#DBEAFE',
    color: '#1E40AF',
    border: '#93C5FD',
    rooms: ['Ruang Tindakan', 'Ronsen', 'Ruang Dokter', 'Ruang Steril']
  },
  'Toilet & Sanitasi': {
    icon: '🚿',
    bg: '#CCFBF1',
    color: '#0F766E',
    border: '#99F6E4',
    rooms: ['Toilet 1', 'Janitor', 'Ruang Limbah']
  },
  'Sirkulasi': {
    icon: '🚶',
    bg: '#F1F5F9',
    color: '#334155',
    border: '#CBD5E1',
    rooms: ['Area Tangga', 'Lift', 'Area Tangga Exit']
  },
  'Area Anak & Fasilitas': {
    icon: '🧸',
    bg: '#FEF3C7',
    color: '#92400E',
    border: '#FDE68A',
    rooms: ['Lobby 2', 'Playground', 'Musholla', 'Tempat Wudhu', 'Ruang Team']
  },
  'Gudang & Teknis': {
    icon: '⚙️',
    bg: '#FFEDD5',
    color: '#C2410C',
    border: '#FED7AA',
    rooms: ['Gudang Farmasi & Bahan', 'RoofTop', 'Ruang Kompressor', 'Area Genset']
  },
  'Eksterior': {
    icon: '🌳',
    bg: '#DCFCE7',
    color: '#15803D',
    border: '#86EFAC',
    rooms: ['Taman', 'Pos Security', 'Toilet Security', 'Halaman Parkir']
  }
};

export function getZoneForRoom(roomName) {
  for (const [zoneName, zoneData] of Object.entries(ZONES)) {
    if (zoneData.rooms.includes(roomName)) {
      return { zoneName, ...zoneData };
    }
  }
  return { zoneName: 'Lainnya', icon: '📍', bg: '#131F36', color: '#94A3B8', border: '#223456' };
}

export async function renderHygieneStandards(container) {
  const user = getUser();
  const canEdit = hasPermission(user, 'hygiene_standards', 'write');
  const canDelete = hasPermission(user, 'hygiene_standards', 'delete');
  const canExport = hasPermission(user, 'export');

  container.innerHTML = `
    <div class="hyg-page-shell">
      <!-- Unified Hero Card (Consistent with Master Karyawan, Kontrak, Master Document) -->
      <div class="hyg-hero-card">
        <div class="hyg-hero-top">
          <div class="hyg-hero-title-group">
            <h1 class="hyg-hero-title">Master Standar Kebersihan Fisik</h1>
            <p class="hyg-hero-subtitle">Acuan standar kebersihan fisik seluruh area dan ruangan klinik operasional</p>
          </div>
          <div class="hyg-hero-actions">
            <div class="hyg-view-switch" role="group" aria-label="Mode tampilan standar kebersihan">
              <button type="button" id="view-table-btn" class="hyg-view-btn active">Tabel Standar</button>
              <button type="button" id="view-zone-btn" class="hyg-view-btn">Per Zona Area</button>
            </div>
            <button type="button" id="btn-print-hygiene" class="btn btn-outline btn-sm" title="Cetak Format SOP Standar Kebersihan">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg>
              Cetak SOP
            </button>
            ${canExport ? `
            <button type="button" id="btn-export-hygiene" class="btn btn-outline btn-sm">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
              Export Excel
            </button>
            ` : ''}
            ${canEdit ? `
            <button type="button" id="btn-add-hygiene" class="btn btn-primary btn-sm">
              + Tambah Parameter
            </button>
            ` : ''}
          </div>
        </div>

        <!-- Compact Clean Navy Summary Stats -->
        <div id="hygiene-summary-cards" class="hyg-stats-grid">
          <div class="hyg-stat-box">
            <div class="hyg-stat-icon" aria-hidden="true">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M9 11l3 3L22 4"/><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11"/></svg>
            </div>
            <div class="hyg-stat-body">
              <div class="hyg-stat-val" id="card-total-params">104</div>
              <div class="hyg-stat-lbl">Total Parameter</div>
            </div>
          </div>
          <div class="hyg-stat-box">
            <div class="hyg-stat-icon" aria-hidden="true">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>
            </div>
            <div class="hyg-stat-body">
              <div class="hyg-stat-val" id="card-total-zones">${Object.keys(ZONES).length}</div>
              <div class="hyg-stat-lbl">Kategori Zona</div>
            </div>
          </div>
          <div class="hyg-stat-box">
            <div class="hyg-stat-icon" aria-hidden="true">
              <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>
            </div>
            <div class="hyg-stat-body">
              <div class="hyg-stat-val" id="card-total-rooms">35</div>
              <div class="hyg-stat-lbl">Ruangan / Lokasi</div>
            </div>
          </div>
        </div>
      </div>

      <!-- Unified Filter & Scrollable Zone Tabs Card -->
      <div class="hyg-controls-card">
        <div class="hyg-controls-row">
          <div class="hyg-search-wrap">
            <svg class="hyg-search-icon" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
            <input type="text" id="hygiene-search" class="form-control hyg-search-input" placeholder="Cari ruangan, item objek, atau standar kebersihan..." aria-label="Cari parameter kebersihan">
          </div>
          <select id="hygiene-room-filter" class="form-control hyg-room-select" aria-label="Filter ruangan">
            <option value="">Semua Ruangan</option>
          </select>
          <button type="button" id="btn-reset-filters" class="btn btn-outline btn-sm hyg-reset-btn">
            Reset
          </button>
          <div id="hygiene-stats-text" class="hyg-stats-text">
            Memuat data...
          </div>
        </div>

        <!-- Horizontal Scrollable Zone Tabs (Enhanced with Global Left/Right Arrows) -->
        <div id="hygiene-zone-tabs" class="hyg-zone-tabs" data-fc-scroll-tabs role="tablist" aria-label="Filter kategori zona kebersihan">
          <button type="button" class="zone-tab-btn active" data-zone="all" role="tab" aria-selected="true">
            <span>Semua Zona</span>
            <span class="badge hyg-tab-count" id="badge-count-all">104</span>
          </button>
          ${Object.keys(ZONES).map((name) => `
            <button type="button" class="zone-tab-btn" data-zone="${escapeHtml(name)}" role="tab" aria-selected="false">
              <span>${escapeHtml(name)}</span>
              <span class="badge hyg-tab-count" id="badge-count-${name.replace(/[\s&]+/g, '-')}">-</span>
            </button>
          `).join('')}
        </div>
      </div>

      <!-- Main Content Container -->
      <div id="hygiene-view-container">
        <div class="skeleton" style="height:320px; border-radius:12px;"></div>
      </div>
    </div>
  `;

  let currentZone = 'all';
  let currentView = 'table'; // 'table' or 'zone'
  let allItems = [];
  // canEdit defined at top

  async function loadData() {
    try {
      const res = await apiFetch('/api/hygiene-standards?all=1');
      if (!res.ok) throw new Error('Gagal memuat standar hygiene');
      allItems = Array.isArray(res.data) ? res.data : (res.data?.data || []);

      // Update room filter options
      populateRoomFilter();
      updateBadgeCounts();
      renderContent();
    } catch (err) {
      console.error(err);
      toastError(err.message || 'Terjadi kesalahan memuat data');
      const viewContainer = container.querySelector('#hygiene-view-container');
      if (viewContainer) {
        viewContainer.innerHTML = `
          <div class="empty-state" style="text-align:center; padding:40px; background:#131F36; border-radius:12px; border:1px solid #223456;">
            <div style="font-size:2.5rem; margin-bottom:12px;">⚠️</div>
            <h3 style="margin:0 0 8px 0; color:var(--text-1);">Gagal Memuat Data</h3>
            <p style="color:var(--text-3); margin-bottom:16px;">${err.message}</p>
            <button id="btn-retry-hygiene" class="btn btn-primary" style="padding:8px 16px; border-radius:8px;">Coba Lagi</button>
          </div>
        `;
        viewContainer.querySelector('#btn-retry-hygiene')?.addEventListener('click', loadData);
      }
    }
  }

  function updateBadgeCounts() {
    // Total count
    const totalEl = container.querySelector('#badge-count-all');
    if (totalEl) totalEl.textContent = allItems.length;
    const totalKpi = container.querySelector('#card-total-params');
    if (totalKpi) totalKpi.textContent = allItems.length;

    // Per zone count
    Object.entries(ZONES).forEach(([zoneName, zoneData]) => {
      const slug = zoneName.replace(/[\s&]+/g, '-');
      const badgeEl = container.querySelector(`#badge-count-${slug}`);
      if (badgeEl) {
        const cnt = allItems.filter(item => zoneData.rooms.includes(item.room_name)).length;
        badgeEl.textContent = cnt;
      }
    });

    // KPI total rooms
    const uniqueRooms = new Set(allItems.map(i => i.room_name));
    const roomsKpi = container.querySelector('#card-total-rooms');
    if (roomsKpi) roomsKpi.textContent = uniqueRooms.size;
  }

  function populateRoomFilter() {
    const roomSelect = container.querySelector('#hygiene-room-filter');
    if (!roomSelect) return;
    const currentVal = roomSelect.value;

    let availableRooms = [];
    if (currentZone === 'all') {
      availableRooms = Array.from(new Set(allItems.map(i => i.room_name))).filter(Boolean).sort();
    } else {
      const zData = ZONES[currentZone];
      if (zData) {
        availableRooms = zData.rooms.filter(r => allItems.some(i => i.room_name === r));
      }
    }

    roomSelect.innerHTML = `<option value="">Semua Ruangan (${availableRooms.length})</option>` +
      availableRooms.map(r => `<option value="${r}" ${r === currentVal ? 'selected' : ''}>${r}</option>`).join('');
  }

  function getFilteredItems() {
    const searchVal = (container.querySelector('#hygiene-search')?.value || '').toLowerCase().trim();
    const roomVal = container.querySelector('#hygiene-room-filter')?.value || '';

    return allItems.filter(item => {
      // Zone filter
      if (currentZone !== 'all') {
        const zData = ZONES[currentZone];
        if (!zData || !zData.rooms.includes(item.room_name)) return false;
      }

      // Room filter
      if (roomVal && item.room_name !== roomVal) return false;

      // Search filter
      if (searchVal) {
        const rName = (item.room_name || '').toLowerCase();
        const iName = (item.item_name || '').toLowerCase();
        const sStd = (item.cleanliness_standard || '').toLowerCase();
        if (!rName.includes(searchVal) && !iName.includes(searchVal) && !sStd.includes(searchVal)) {
          return false;
        }
      }

      return true;
    });
  }

  function renderContent() {
    const filtered = getFilteredItems();
    const statsEl = container.querySelector('#hygiene-stats-text');
    if (statsEl) {
      statsEl.innerHTML = `Menampilkan <strong>${filtered.length}</strong> dari <strong>${allItems.length}</strong> parameter`;
    }

    const viewContainer = container.querySelector('#hygiene-view-container');
    if (!viewContainer) return;

    if (filtered.length === 0) {
      viewContainer.innerHTML = `
        <div class="empty-state" style="text-align:center; padding:50px 20px; background:#131F36; border-radius:12px; border:1px solid #223456;">
          <div style="font-size:3rem; margin-bottom:12px;">🔍</div>
          <h3 style="margin:0 0 8px 0; color:var(--text-1); font-weight:700;">Tidak Ada Parameter yang Cocok</h3>
          <p style="color:var(--text-3); max-width:400px; margin:0 auto 16px auto; font-size:0.9rem;">
            Tidak ditemukan parameter kebersihan dengan kata kunci atau filter yang Anda pilih.
          </p>
          <button id="btn-clear-search-empty" class="btn btn-outline" style="border-radius:8px; padding:8px 16px; font-weight:600;">
            Reset Filter
          </button>
        </div>
      `;
      viewContainer.querySelector('#btn-clear-search-empty')?.addEventListener('click', () => {
        const searchInput = container.querySelector('#hygiene-search');
        const roomSelect = container.querySelector('#hygiene-room-filter');
        if (searchInput) searchInput.value = '';
        if (roomSelect) roomSelect.value = '';
        currentZone = 'all';
        container.querySelectorAll('.zone-tab-btn').forEach(b => {
          const isAll = b.dataset.zone === 'all';
          b.classList.toggle('active', isAll);
          b.setAttribute('aria-selected', isAll ? 'true' : 'false');
        });
        populateRoomFilter();
        renderContent();
      });
      return;
    }

    if (currentView === 'table') {
      renderTableView(viewContainer, filtered);
    } else {
      renderZoneView(viewContainer, filtered);
    }
  }

  function renderTableView(viewContainer, items) {
    viewContainer.innerHTML = `
      <div class="hyg-table-card hyg-table-card--desktop">
        <div class="table-wrapper">
          <table class="hyg-table">
            <thead>
              <tr>
                <th style="width:52px; text-align:center;">No</th>
                <th style="width:220px;">Ruangan &amp; Zona</th>
                <th style="width:280px;">Item / Objek Inspeksi</th>
                <th>Standar Kebersihan Fisik</th>
                ${canEdit || canDelete ? '<th style="width:96px; text-align:center;">Aksi</th>' : ''}
              </tr>
            </thead>
            <tbody>
              ${items.map((item, idx) => {
                const zInfo = getZoneForRoom(item.room_name);
                return `
                  <tr class="hygiene-std-row">
                    <td class="hyg-col-no">${idx + 1}</td>
                    <td class="hyg-col-room">
                      <span class="hyg-room-badge">${escapeHtml(item.room_name)}</span>
                      <div class="hyg-zone-meta">${escapeHtml(zInfo.zoneName)}</div>
                    </td>
                    <td class="hyg-col-item">
                      <div class="hyg-item-title">${escapeHtml(item.item_name)}</div>
                    </td>
                    <td class="hyg-col-std">
                      <div class="hyg-std-box">
                        <span class="hyg-std-check" aria-hidden="true">✓</span>
                        <span class="hyg-std-text">${escapeHtml(item.cleanliness_standard)}</span>
                      </div>
                    </td>
                    ${canEdit || canDelete ? `
                      <td class="hyg-col-actions">
                        ${renderHygieneActionButtons(item, canEdit, canDelete, false)}
                      </td>
                    ` : ''}
                  </tr>
                `;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>
      ${renderHygieneMobileCards(items, { includeRoom: true, canEdit, canDelete })}
    `;

    attachActionButtons(viewContainer);
  }

  function renderZoneView(viewContainer, items) {
    // Group items by zone and room
    const grouped = {};
    items.forEach(item => {
      const zInfo = getZoneForRoom(item.room_name);
      if (!grouped[zInfo.zoneName]) {
        grouped[zInfo.zoneName] = { info: zInfo, rooms: {} };
      }
      if (!grouped[zInfo.zoneName].rooms[item.room_name]) {
        grouped[zInfo.zoneName].rooms[item.room_name] = [];
      }
      grouped[zInfo.zoneName].rooms[item.room_name].push(item);
    });

    viewContainer.innerHTML = `
      <div class="hyg-zone-list">
        ${Object.entries(grouped).map(([zoneName, zGroup]) => {
          const totalInZone = Object.values(zGroup.rooms).reduce((acc, arr) => acc + arr.length, 0);
          return `
            <div class="hyg-zone-card">
              <div class="hyg-zone-card-header">
                <div>
                  <h3 class="hyg-zone-card-title">${escapeHtml(zoneName)}</h3>
                  <div class="hyg-zone-card-sub">
                    ${Object.keys(zGroup.rooms).length} Ruangan · ${totalInZone} Parameter Standar
                  </div>
                </div>
                <span class="badge badge-primary">${totalInZone} Parameter</span>
              </div>

              <div class="hyg-zone-card-body">
                ${Object.entries(zGroup.rooms).map(([roomName, rItems]) => `
                  <div class="hyg-room-block">
                    <div class="hyg-room-block-head">
                      <h4 class="hyg-room-block-title">${escapeHtml(roomName)}</h4>
                      <span class="badge badge-secondary">${rItems.length} item</span>
                    </div>

                    <div class="hyg-zone-table-wrap hyg-zone-table-wrap--desktop">
                      <table class="hyg-table hyg-table--compact">
                        <thead>
                          <tr>
                            <th style="width:44px; text-align:center;">No</th>
                            <th style="width:260px;">Item / Objek Inspeksi</th>
                            <th>Standar Kebersihan Fisik</th>
                            ${canEdit || canDelete ? '<th style="width:86px; text-align:center;">Aksi</th>' : ''}
                          </tr>
                        </thead>
                        <tbody>
                          ${rItems.map((item, idx) => `
                            <tr class="hygiene-std-row">
                              <td class="hyg-col-no">${idx + 1}</td>
                              <td class="hyg-col-item">
                                <div class="hyg-item-title">${escapeHtml(item.item_name)}</div>
                              </td>
                              <td class="hyg-col-std">
                                <div class="hyg-std-inline">
                                  <span class="hyg-std-check" aria-hidden="true">✓</span>
                                  <span class="hyg-std-text">${escapeHtml(item.cleanliness_standard)}</span>
                                </div>
                              </td>
                              ${canEdit || canDelete ? `
                                <td class="hyg-col-actions">
                                  ${renderHygieneActionButtons(item, canEdit, canDelete, true)}
                                </td>
                              ` : ''}
                            </tr>
                          `).join('')}
                        </tbody>
                      </table>
                    </div>
                    ${renderHygieneMobileCards(rItems, { includeRoom: false, canEdit, canDelete })}
                  </div>
                `).join('')}
              </div>
            </div>
          `;
        }).join('')}
      </div>
    `;

    attachActionButtons(viewContainer);
  }

  function attachActionButtons(viewContainer) {
    // Edit item
    viewContainer.querySelectorAll('.btn-edit-item').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.dataset.id;
        const item = allItems.find(i => String(i.id) === String(id));
        if (item) openItemModal(item);
      });
    });

    // Delete item
    viewContainer.querySelectorAll('.btn-delete-item').forEach(btn => {
      btn.addEventListener('click', () => {
        if (!canDelete) {
          toastError('Hanya Superadmin yang dapat menghapus parameter standar.');
          return;
        }
        const id = btn.dataset.id;
        const name = btn.dataset.name;
        confirmDialog(`Yakin ingin menghapus parameter standar <strong>"${name}"</strong>?`, async () => {
          try {
            const res = await apiFetch(`/api/hygiene-standards/${id}`, { method: 'DELETE' });
            if (!res.ok) throw new Error(res.data?.error || 'Gagal menghapus parameter');
            toastSuccess('Parameter kebersihan berhasil dihapus');
            loadData();
          } catch (err) {
            toastError(err.message);
          }
        }, 'Hapus Parameter Standar');
      });
    });
  }

  function openItemModal(item = null) {
    if (!canEdit) {
      toastError('Anda tidak memiliki izin untuk menambah atau mengedit parameter kebersihan.');
      return;
    }
    const isEdit = !!item;
    const allRooms = Array.from(new Set(allItems.map(i => i.room_name))).filter(Boolean).sort();

    const content = document.createElement('div');
    content.innerHTML = `
      <form id="form-hygiene-item" style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Ruangan / Lokasi <span class="required">*</span>
          </label>
          <input type="text" id="modal-room-name" list="room-suggestions" class="form-control" value="${item?.room_name || ''}" required style="width:100%; height:40px; border-radius:8px;">
          <datalist id="room-suggestions">
            ${allRooms.map(r => `<option value="${r}">`).join('')}
          </datalist>
          <span style="font-size:0.75rem; color:var(--text-3); margin-top:4px; display:block;">Pilih dari daftar atau ketik nama ruangan baru</span>
        </div>
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Item / Objek Inspeksi <span class="required">*</span>
          </label>
          <input type="text" id="modal-item-name" class="form-control" value="${item?.item_name || ''}" required style="width:100%; height:40px; border-radius:8px;">
        </div>
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Standar Kebersihan Fisik <span class="required">*</span>
          </label>
          <textarea id="modal-standard" class="form-control" rows="4" required style="width:100%; border-radius:8px; line-height:1.5;">${item?.cleanliness_standard || ''}</textarea>
          <span style="font-size:0.75rem; color:var(--text-3); margin-top:4px; display:block;">Jelaskan kriteria fisik: bebas debu kasat mata, steril, tidak licin, tidak berbau, dll.</span>
        </div>
      </form>
    `;

    createModal({
      title: isEdit ? '✏️ Edit Parameter Kebersihan' : '➕ Tambah Parameter Kebersihan',
      content,
      confirmText: isEdit ? 'Simpan Perubahan' : 'Tambah Parameter',
      size: 'md',
      onConfirm: async (_, close) => {
        const room_name = content.querySelector('#modal-room-name').value.trim();
        const item_name = content.querySelector('#modal-item-name').value.trim();
        const cleanliness_standard = content.querySelector('#modal-standard').value.trim();

        if (!room_name || !item_name || !cleanliness_standard) {
          toastError('Semua kolom wajib diisi');
          return;
        }

        try {
          const payload = { room_name, item_name, cleanliness_standard };
          const url = isEdit ? `/api/hygiene-standards/${item.id}` : '/api/hygiene-standards';
          const method = isEdit ? 'PUT' : 'POST';

          const res = await apiFetch(url, { method, body: JSON.stringify(payload) });
          if (!res.ok) throw new Error(res.data?.error || 'Gagal menyimpan parameter');

          toastSuccess(isEdit ? 'Parameter berhasil diperbarui' : 'Parameter berhasil ditambahkan');
          close();
          loadData();
        } catch (err) {
          toastError(err.message);
        }
      }
    });
  }

  function printSOP() {
    // Generate clean printable document
    const printWindow = window.open('', '_blank');
    if (!printWindow) {
      toastError('Pop-up printer diblokir oleh browser. Izinkan pop-up untuk mencetak SOP.');
      return;
    }

    const todayStr = new Date().toLocaleDateString('id-ID', {
      day: 'numeric',
      month: 'long',
      year: 'numeric'
    });

    const printHtml = `
      <!DOCTYPE html>
      <html lang="id">
      <head>
        <meta charset="UTF-8">
        <title>SOP Master Standar Kebersihan Fisik - Audy Dental</title>
        <style>
          @page { size: A4 portrait; margin: 12mm 15mm 15mm 15mm; }
          body { font-family: 'Segoe UI', Arial, sans-serif; font-size: 10pt; color: #1e293b; line-height: 1.4; margin: 0; padding: 0; }
          .header { border-bottom: 2px solid #1e3a8a; padding-bottom: 8px; margin-bottom: 14px; display: flex; justify-content: space-between; align-items: flex-end; }
          .logo { font-size: 16pt; font-weight: 800; color: #1e3a8a; letter-spacing: -0.5px; }
          .sublogo { font-size: 8.5pt; color: #64748b; font-weight: 600; text-transform: uppercase; }
          .doc-info { text-align: right; font-size: 8pt; color: #475569; }
          .title { text-align: center; margin: 12px 0 16px 0; }
          .title h1 { font-size: 13pt; margin: 0 0 4px 0; color: #0f172a; text-transform: uppercase; }
          .title p { margin: 0; font-size: 8.5pt; color: #64748b; }
          .zone-title { background: #f1f5f9; border-left: 4px solid #2563eb; padding: 6px 10px; font-size: 9.5pt; font-weight: 700; color: #1e3a8a; margin: 14px 0 8px 0; page-break-after: avoid; }
          table { width: 100%; border-collapse: collapse; margin-bottom: 12px; font-size: 8.5pt; page-break-inside: auto; }
          tr { page-break-inside: avoid; page-break-after: auto; }
          th { background: #f8fafc; border: 1px solid #cbd5e1; padding: 6px 8px; text-align: left; font-weight: 700; color: #334155; font-size: 8pt; text-transform: uppercase; }
          td { border: 1px solid #cbd5e1; padding: 6px 8px; vertical-align: top; }
          .td-no { text-align: center; width: 28px; }
          .td-room { width: 120px; font-weight: 600; }
          .td-obj { width: 150px; font-weight: 600; color: #0f172a; }
          .footer-sign { margin-top: 24px; display: flex; justify-content: space-between; page-break-inside: avoid; }
          .sign-box { text-align: center; width: 180px; font-size: 8.5pt; }
          .sign-line { margin-top: 45px; border-top: 1px solid #0f172a; font-weight: 600; padding-top: 4px; }
          @media print {
            .no-print { display: none !important; }
          }
        </style>
      </head>
      <body>
        <div class="no-print" style="background:#fef3c7; border:1px solid #f59e0b; padding:10px 14px; border-radius:6px; margin-bottom:14px; display:flex; justify-content:space-between; align-items:center;">
          <span style="font-size:9pt; color:#92400e;">💡 <strong>Format Cetak SOP Standar Kebersihan Fisik FCMS</strong>. Siap dicetak untuk display janitor atau arsip operasional.</span>
          <button id="hygiene-print-now-btn" style="background:#2563eb; color:#fff; border:none; padding:6px 14px; border-radius:4px; font-weight:bold; cursor:pointer;">Cetak Sekarang</button>
          <script>document.getElementById('hygiene-print-now-btn')?.addEventListener('click', function(){ window.print(); });</script>
        </div>

        <div class="header">
          <div>
            <div class="logo">AUDY DENTAL</div>
            <div class="sublogo">Facility Care Management System (FCMS)</div>
          </div>
          <div class="doc-info">
            <div>Dokumen: <strong>FCMS-SOP-HYG-01</strong></div>
            <div>Tanggal Cetak: ${todayStr}</div>
            <div>Total: <strong>${allItems.length} Parameter Fisik</strong></div>
          </div>
        </div>

        <div class="title">
          <h1>STANDAR ACUAN KEBERSIHAN FISIK SELURUH CABANG</h1>
          <p>Buku Panduan Kriteria Kebersihan Fisik 8 Zona Area (Cleaning Service & Facility Care)</p>
        </div>

        ${Object.entries(ZONES).map(([zName, zData]) => {
          const zItems = allItems.filter(i => zData.rooms.includes(i.room_name));
          if (zItems.length === 0) return '';
          return `
            <div class="zone-title">${zData.icon} ${zName.toUpperCase()} (${zItems.length} PARAMETER)</div>
            <table>
              <thead>
                <tr>
                  <th class="td-no">No</th>
                  <th class="td-room">Ruangan / Lokasi</th>
                  <th class="td-obj">Item / Objek Inspeksi</th>
                  <th>Standar Kebersihan Fisik</th>
                </tr>
              </thead>
              <tbody>
                ${zItems.map((item, idx) => `
                  <tr>
                    <td class="td-no">${idx + 1}</td>
                    <td class="td-room">${item.room_name}</td>
                    <td class="td-obj">${item.item_name}</td>
                    <td>${item.cleanliness_standard}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          `;
        }).join('')}

        <div class="footer-sign">
          <div class="sign-box">
            <div>Dibuat Oleh,</div>
            <div class="sign-line">Facility Care Specialist</div>
          </div>
          <div class="sign-box">
            <div>Ditinjau Oleh,</div>
            <div class="sign-line">Supervisor Operasional</div>
          </div>
          <div class="sign-box">
            <div>Disetujui Oleh,</div>
            <div class="sign-line">Head of Operations</div>
          </div>
        </div>

        <script>
          window.onload = function() {
            setTimeout(function() { window.print(); }, 400);
          };
        </script>
      </body>
      </html>
    `;

    printWindow.document.open();
    printWindow.document.write(printHtml);
    printWindow.document.close();
  }

  async function exportExcel() {
    try {
      const rows = await fetchExportRows('/api/hygiene-standards/export', {});
      const data = rows.map((item, idx) => {
        const zInfo = getZoneForRoom(item.room_name);
        return {
          'No': idx + 1,
          'Kategori Zona': zInfo.zoneName,
          'Ruangan / Lokasi': item.room_name,
          'Item / Objek Inspeksi': item.item_name,
          'Standar Kebersihan Fisik': item.cleanliness_standard
        };
      });

      const dateStr = new Date().toISOString().slice(0, 10);
      downloadExcel(data, `Master_Standar_Kebersihan_FCMS_${dateStr}`);
      toastSuccess('Data master kebersihan berhasil diekspor ke Excel');
    } catch (err) {
      toastError('Gagal mengekspor data: ' + err.message);
    }
  }

  // Event Listeners
  container.querySelector('#hygiene-search')?.addEventListener('input', () => renderContent());
  container.querySelector('#hygiene-room-filter')?.addEventListener('change', () => renderContent());
  container.querySelector('#btn-reset-filters')?.addEventListener('click', () => {
    const searchInput = container.querySelector('#hygiene-search');
    const roomSelect = container.querySelector('#hygiene-room-filter');
    if (searchInput) searchInput.value = '';
    if (roomSelect) roomSelect.value = '';
    renderContent();
  });

  // Zone Tabs Switcher
  container.querySelectorAll('.zone-tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      currentZone = btn.dataset.zone;
      container.querySelectorAll('.zone-tab-btn').forEach(b => {
        const isActive = b.dataset.zone === currentZone;
        b.classList.toggle('active', isActive);
        b.setAttribute('aria-selected', isActive ? 'true' : 'false');
      });
      populateRoomFilter();
      renderContent();
    });
  });

  // View Switcher (Table vs Zone)
  const viewTableBtn = container.querySelector('#view-table-btn');
  const viewZoneBtn = container.querySelector('#view-zone-btn');

  viewTableBtn?.addEventListener('click', () => {
    currentView = 'table';
    viewTableBtn.classList.add('active');
    viewZoneBtn?.classList.remove('active');
    renderContent();
  });

  viewZoneBtn?.addEventListener('click', () => {
    currentView = 'zone';
    viewZoneBtn.classList.add('active');
    viewTableBtn?.classList.remove('active');
    renderContent();
  });

  // Action Buttons
  container.querySelector('#btn-print-hygiene')?.addEventListener('click', printSOP);
  container.querySelector('#btn-export-hygiene')?.addEventListener('click', exportExcel);
  container.querySelector('#btn-add-hygiene')?.addEventListener('click', () => openItemModal(null));

  // Initial Load
  loadData();
}
