import { apiFetch, getUser } from '../config.js';
import { createModal, confirmDialog } from '../components/modal.js';
import { toastSuccess, toastError } from '../components/toast.js';
import { downloadExcel } from '../utils/excel.js';

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
  return { zoneName: 'Lainnya', icon: '📍', bg: '#F1F5F9', color: '#475569', border: '#E2E8F0' };
}

export async function renderHygieneStandards(container) {
  container.innerHTML = `
    <!-- Header Section -->
    <div class="page-header" style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px; margin-bottom:20px;">
      <div>
        <h1 class="page-title" style="display:flex; align-items:center; gap:8px; margin:0 0 6px 0; font-size:1.6rem; font-weight:700; color:var(--text-1);">
          <span>✨</span> Master Standar Kebersihan Fisik
        </h1>
        <p class="page-subtitle" style="margin:0; font-size:0.9rem; color:var(--text-3);">
          Acuan Standar Kebersihan Fisik Seluruh Area & Ruangan Klinik (Audy Dental · Facility Care)
        </p>
      </div>
      <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap;">
        <div class="btn-group" style="background:#F1F5F9; padding:3px; border-radius:8px; display:inline-flex; border:1px solid var(--border);">
          <button id="view-table-btn" class="btn btn-sm" style="background:#fff; font-weight:600; border-radius:6px; box-shadow:0 1px 2px rgba(0,0,0,0.05); padding:6px 12px; cursor:pointer;">
            📋 Tabel Standar
          </button>
          <button id="view-zone-btn" class="btn btn-sm btn-ghost" style="font-weight:600; border-radius:6px; padding:6px 12px; cursor:pointer;">
            🗂️ Per Zona Area
          </button>
        </div>
        <button id="btn-print-hygiene" class="btn btn-outline" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;" title="Cetak Format SOP Standar Kebersihan">
          🖨️ Cetak SOP
        </button>
        <button id="btn-export-hygiene" class="btn btn-outline" style="background:#fff; border:1px solid var(--border); font-weight:600; padding:8px 14px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;">
          📥 Export Excel
        </button>
        <button id="btn-add-hygiene" class="btn btn-primary" style="font-weight:600; padding:8px 16px; border-radius:8px; display:inline-flex; align-items:center; gap:6px; cursor:pointer;">
          + Tambah Parameter
        </button>
      </div>
    </div>

    <!-- KPI Summary Cards -->
    <div id="hygiene-summary-cards" style="display:grid; grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); gap:16px; margin-bottom:24px;">
      <div class="card" style="padding:16px 20px; border-radius:12px; border:1px solid var(--border); background:#fff; display:flex; align-items:center; gap:16px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="width:48px; height:48px; border-radius:10px; background:#EFF6FF; color:#2563EB; display:flex; align-items:center; justify-content:center; font-size:1.5rem; flex-shrink:0;">
          📋
        </div>
        <div>
          <div style="font-size:0.8rem; color:var(--text-3); font-weight:600; text-transform:uppercase; letter-spacing:0.5px;">Total Parameter</div>
          <div style="font-size:1.6rem; font-weight:800; color:var(--text-1);" id="card-total-params">104</div>
          <div style="font-size:0.75rem; color:#16A34A; font-weight:500;">✓ Terverifikasi Standar SOP</div>
        </div>
      </div>
      <div class="card" style="padding:16px 20px; border-radius:12px; border:1px solid var(--border); background:#fff; display:flex; align-items:center; gap:16px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="width:48px; height:48px; border-radius:10px; background:#F5F3FF; color:#7C3AED; display:flex; align-items:center; justify-content:center; font-size:1.5rem; flex-shrink:0;">
          🏢
        </div>
        <div>
          <div style="font-size:0.8rem; color:var(--text-3); font-weight:600; text-transform:uppercase; letter-spacing:0.5px;">Kategori Zona</div>
          <div style="font-size:1.6rem; font-weight:800; color:var(--text-1);" id="card-total-zones">8</div>
          <div style="font-size:0.75rem; color:var(--text-3); font-weight:500;">Klinis, Publik, Sanitair, dll.</div>
        </div>
      </div>
      <div class="card" style="padding:16px 20px; border-radius:12px; border:1px solid var(--border); background:#fff; display:flex; align-items:center; gap:16px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="width:48px; height:48px; border-radius:10px; background:#ECFDF5; color:#059669; display:flex; align-items:center; justify-content:center; font-size:1.5rem; flex-shrink:0;">
          🚪
        </div>
        <div>
          <div style="font-size:0.8rem; color:var(--text-3); font-weight:600; text-transform:uppercase; letter-spacing:0.5px;">Ruangan / Lokasi</div>
          <div style="font-size:1.6rem; font-weight:800; color:var(--text-1);" id="card-total-rooms">35</div>
          <div style="font-size:0.75rem; color:#059669; font-weight:500;">Termasuk Ruang Tindakan</div>
        </div>
      </div>
      <div class="card" style="padding:16px 20px; border-radius:12px; border:1px solid var(--border); background:#fff; display:flex; align-items:center; gap:16px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="width:48px; height:48px; border-radius:10px; background:#FFFBEB; color:#D97706; display:flex; align-items:center; justify-content:center; font-size:1.5rem; flex-shrink:0;">
          🛡️
        </div>
        <div>
          <div style="font-size:0.8rem; color:var(--text-3); font-weight:600; text-transform:uppercase; letter-spacing:0.5px;">Kondisi Fisik</div>
          <div style="font-size:1.6rem; font-weight:800; color:#16A34A;">100%</div>
          <div style="font-size:0.75rem; color:var(--text-3); font-weight:500;">Bebas Debu, Kerak & Bau</div>
        </div>
      </div>
    </div>

    <!-- Zona Switcher Tabs -->
    <div style="overflow-x:auto; padding-bottom:8px; margin-bottom:18px;">
      <div id="hygiene-zone-tabs" style="display:flex; gap:8px; min-width:max-content; border-bottom:2px solid #E2E8F0; padding-bottom:8px;">
        <button class="zone-tab-btn active" data-zone="all" style="border:none; background:#2563EB; color:#fff; font-weight:700; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.88rem;">
          🌟 Semua Zona <span class="badge" style="background:rgba(255,255,255,0.25); color:#fff; font-size:0.75rem; padding:2px 6px; border-radius:10px;" id="badge-count-all">104</span>
        </button>
        ${Object.entries(ZONES).map(([name, data]) => `
          <button class="zone-tab-btn" data-zone="${name}" style="border:none; background:#F1F5F9; color:var(--text-2); font-weight:600; padding:8px 16px; border-radius:8px; cursor:pointer; display:inline-flex; align-items:center; gap:6px; font-size:0.88rem;">
            <span>${data.icon}</span> ${name}
            <span class="badge" style="background:#CBD5E1; color:#334155; font-size:0.75rem; padding:2px 6px; border-radius:10px;" id="badge-count-${name.replace(/[\s&]+/g, '-')}">-</span>
          </button>
        `).join('')}
      </div>
    </div>

    <!-- Filter Toolbar -->
    <div class="card" style="padding:14px 18px; margin-bottom:20px; background:#fff; border:1px solid var(--border); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
      <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; justify-content:space-between;">
        <div style="display:flex; flex-wrap:wrap; gap:12px; align-items:center; flex:1; min-width:280px;">
          <div style="position:relative; flex:1; min-width:220px; max-width:360px;">
            <input type="text" id="hygiene-search" class="form-control" placeholder="Cari ruangan, item objek, atau standar kebersihan..." style="padding-left:36px; height:40px; border-radius:8px; font-size:0.9rem; width:100%;">
            <span style="position:absolute; left:12px; top:50%; transform:translateY(-50%); color:var(--text-3); font-size:0.95rem;">🔍</span>
          </div>
          <select id="hygiene-room-filter" class="form-control" style="width:auto; height:40px; border-radius:8px; font-size:0.875rem; min-width:180px;">
            <option value="">Semua Ruangan</option>
          </select>
          <button id="btn-reset-filters" class="btn btn-sm btn-ghost" style="color:var(--text-3); font-weight:600; cursor:pointer; padding:6px 10px;">
            ✕ Reset
          </button>
        </div>
        <div id="hygiene-stats-text" style="font-size:0.85rem; color:var(--text-3); font-weight:500;">
          Memuat data...
        </div>
      </div>
    </div>

    <!-- Main Content Container -->
    <div id="hygiene-view-container">
      <div class="skeleton" style="height:320px; border-radius:12px;"></div>
    </div>
  `;

  let currentZone = 'all';
  let currentView = 'table'; // 'table' or 'zone'
  let allItems = [];
  const user = getUser();
  const canEdit = user && ['superadmin', 'admin', 'manager', 'spv'].includes(user.role);

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
          <div class="empty-state" style="text-align:center; padding:40px; background:#fff; border-radius:12px; border:1px solid var(--border);">
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
        <div class="empty-state" style="text-align:center; padding:50px 20px; background:#fff; border-radius:12px; border:1px solid var(--border);">
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
        container.querySelectorAll('.zone-tab-btn').forEach(b => b.classList.toggle('active', b.dataset.zone === 'all'));
        container.querySelectorAll('.zone-tab-btn').forEach(b => {
          if (b.dataset.zone === 'all') {
            b.style.background = '#2563EB';
            b.style.color = '#fff';
          } else {
            b.style.background = '#F1F5F9';
            b.style.color = 'var(--text-2)';
          }
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
      <div class="card" style="background:#fff; border:1px solid var(--border); border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
        <div style="overflow-x:auto;">
          <table style="width:100%; border-collapse:collapse; text-align:left; font-size:0.9rem;">
            <thead>
              <tr style="background:#F8FAFC; border-bottom:2px solid #E2E8F0; color:#475569; font-size:0.78rem; text-transform:uppercase; letter-spacing:0.6px;">
                <th style="padding:14px 16px; width:50px; text-align:center;">No</th>
                <th style="padding:14px 16px; width:220px;">Ruangan & Zona</th>
                <th style="padding:14px 16px; width:280px;">Item / Objek Inspeksi</th>
                <th style="padding:14px 16px;">Standar Kebersihan Fisik</th>
                ${canEdit ? '<th style="padding:14px 16px; width:100px; text-align:center;">Aksi</th>' : ''}
              </tr>
            </thead>
            <tbody>
              ${items.map((item, idx) => {
                const zInfo = getZoneForRoom(item.room_name);
                return `
                  <tr style="border-bottom:1px solid #F1F5F9; transition:background-color 0.15s ease;" onmouseover="this.style.backgroundColor='#F8FAFC'" onmouseout="this.style.backgroundColor='transparent'">
                    <td style="padding:14px 16px; text-align:center; color:var(--text-3); font-weight:600; font-size:0.85rem;">
                      ${idx + 1}
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="margin-bottom:4px;">
                        <span style="display:inline-flex; align-items:center; gap:6px; background:${zInfo.bg}; color:${zInfo.color}; border:1px solid ${zInfo.border}; padding:4px 10px; border-radius:6px; font-weight:700; font-size:0.83rem;">
                          <span>${zInfo.icon}</span> ${item.room_name}
                        </span>
                      </div>
                      <div style="font-size:0.75rem; color:var(--text-3); font-weight:500;">
                        ${zInfo.zoneName}
                      </div>
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="font-weight:700; color:var(--text-1); line-height:1.45; font-size:0.9rem;">
                        ${item.item_name}
                      </div>
                    </td>
                    <td style="padding:14px 16px; vertical-align:top;">
                      <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-left:3px solid #10B981; border-radius:6px; padding:10px 12px; line-height:1.55; color:#334155; font-size:0.875rem;">
                        <span style="color:#059669; font-weight:700; margin-right:4px;">✓</span> ${item.cleanliness_standard}
                      </div>
                    </td>
                    ${canEdit ? `
                      <td style="padding:14px 16px; text-align:center; vertical-align:middle;">
                        <div style="display:inline-flex; gap:6px;">
                          <button class="btn btn-sm btn-ghost btn-edit-item" data-id="${item.id}" title="Edit Parameter" style="padding:6px 8px; border-radius:6px; color:#2563EB; cursor:pointer;">
                            ✏️
                          </button>
                          <button class="btn btn-sm btn-ghost btn-delete-item" data-id="${item.id}" data-name="${item.item_name}" title="Hapus Parameter" style="padding:6px 8px; border-radius:6px; color:#DC2626; cursor:pointer;">
                            🗑️
                          </button>
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
      <div style="display:flex; flex-direction:column; gap:24px;">
        ${Object.entries(grouped).map(([zoneName, zGroup]) => {
          const totalInZone = Object.values(zGroup.rooms).reduce((acc, arr) => acc + arr.length, 0);
          return `
            <div class="card" style="background:#fff; border:1px solid ${zGroup.info.border}; border-radius:14px; overflow:hidden; box-shadow:0 2px 6px rgba(0,0,0,0.03);">
              <!-- Zone Card Header -->
              <div style="background:${zGroup.info.bg}; border-bottom:1px solid ${zGroup.info.border}; padding:14px 20px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
                <div style="display:flex; align-items:center; gap:10px;">
                  <span style="font-size:1.4rem;">${zGroup.info.icon}</span>
                  <div>
                    <h3 style="margin:0; font-size:1.1rem; font-weight:700; color:${zGroup.info.color};">
                      ${zoneName}
                    </h3>
                    <div style="font-size:0.78rem; color:${zGroup.info.color}; opacity:0.85;">
                      ${Object.keys(zGroup.rooms).length} Ruangan · ${totalInZone} Parameter Standar
                    </div>
                  </div>
                </div>
                <span class="badge" style="background:#fff; color:${zGroup.info.color}; border:1px solid ${zGroup.info.border}; font-weight:700; padding:4px 10px; border-radius:12px; font-size:0.8rem;">
                  ${totalInZone} Parameter
                </span>
              </div>

              <!-- Zone Rooms List -->
              <div style="padding:16px 20px;">
                ${Object.entries(zGroup.rooms).map(([roomName, rItems]) => `
                  <div style="margin-bottom:20px; last-child:margin-bottom:0;">
                    <div style="display:flex; align-items:center; gap:8px; margin-bottom:10px;">
                      <h4 style="margin:0; font-size:0.95rem; font-weight:700; color:var(--text-1);">
                        🚪 ${roomName}
                      </h4>
                      <span class="badge badge-sm" style="background:#F1F5F9; color:var(--text-3); font-size:0.75rem;">
                        ${rItems.length} item
                      </span>
                    </div>

                    <div style="border:1px solid #E2E8F0; border-radius:8px; overflow:hidden;">
                      <table style="width:100%; border-collapse:collapse; font-size:0.875rem;">
                        <thead>
                          <tr style="background:#F8FAFC; border-bottom:1px solid #E2E8F0; color:#475569; font-size:0.75rem; text-transform:uppercase;">
                            <th style="padding:8px 12px; width:40px; text-align:center;">No</th>
                            <th style="padding:8px 12px; width:260px;">Item / Objek Inspeksi</th>
                            <th style="padding:8px 12px;">Standar Kebersihan Fisik</th>
                            ${canEdit ? '<th style="padding:8px 12px; width:80px; text-align:center;">Aksi</th>' : ''}
                          </tr>
                        </thead>
                        <tbody>
                          ${rItems.map((item, idx) => `
                            <tr style="border-bottom:1px solid #F1F5F9;">
                              <td style="padding:10px 12px; text-align:center; color:var(--text-3); font-size:0.8rem;">
                                ${idx + 1}
                              </td>
                              <td style="padding:10px 12px; font-weight:600; color:var(--text-1); vertical-align:top;">
                                ${item.item_name}
                              </td>
                              <td style="padding:10px 12px; vertical-align:top; color:#334155;">
                                <div style="display:flex; align-items:flex-start; gap:6px;">
                                  <span style="color:#10B981; font-weight:700;">✓</span>
                                  <span>${item.cleanliness_standard}</span>
                                </div>
                              </td>
                              ${canEdit ? `
                                <td style="padding:10px 12px; text-align:center; vertical-align:middle;">
                                  <div style="display:inline-flex; gap:4px;">
                                    <button class="btn btn-sm btn-ghost btn-edit-item" data-id="${item.id}" title="Edit" style="padding:4px 6px; cursor:pointer;">✏️</button>
                                    <button class="btn btn-sm btn-ghost btn-delete-item" data-id="${item.id}" data-name="${item.item_name}" title="Hapus" style="padding:4px 6px; color:#DC2626; cursor:pointer;">🗑️</button>
                                  </div>
                                </td>
                              ` : ''}
                            </tr>
                          `).join('')}
                        </tbody>
                      </table>
                    </div>
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
    const isEdit = !!item;
    const allRooms = Array.from(new Set(allItems.map(i => i.room_name))).filter(Boolean).sort();

    const content = document.createElement('div');
    content.innerHTML = `
      <form id="form-hygiene-item" style="display:flex; flex-direction:column; gap:16px;">
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Ruangan / Lokasi <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="modal-room-name" list="room-suggestions" class="form-control" placeholder="Contoh: Ruang Tindakan, Lobby 1, Toilet 1..." value="${item?.room_name || ''}" required style="width:100%; height:40px; border-radius:8px;">
          <datalist id="room-suggestions">
            ${allRooms.map(r => `<option value="${r}">`).join('')}
          </datalist>
          <span style="font-size:0.75rem; color:var(--text-3); margin-top:4px; display:block;">Pilih dari daftar atau ketik nama ruangan baru</span>
        </div>
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Item / Objek Inspeksi <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="modal-item-name" class="form-control" placeholder="Contoh: Eksterior Dental Unit & Kursi, Lantai & Sudut..." value="${item?.item_name || ''}" required style="width:100%; height:40px; border-radius:8px;">
        </div>
        <div>
          <label style="display:block; font-size:0.85rem; font-weight:600; color:var(--text-1); margin-bottom:6px;">
            Standar Kebersihan Fisik <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="modal-standard" class="form-control" rows="4" placeholder="Deskripsi kondisi fisik kebersihan yang dipersyaratkan..." required style="width:100%; border-radius:8px; line-height:1.5;">${item?.cleanliness_standard || ''}</textarea>
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
          <button onclick="window.print()" style="background:#2563eb; color:#fff; border:none; padding:6px 14px; border-radius:4px; font-weight:bold; cursor:pointer;">Cetak Sekarang</button>
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

  function exportExcel() {
    const data = allItems.map((item, idx) => {
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
        if (isActive) {
          b.style.background = '#2563EB';
          b.style.color = '#fff';
          const badge = b.querySelector('.badge');
          if (badge) {
            badge.style.background = 'rgba(255,255,255,0.25)';
            badge.style.color = '#fff';
          }
        } else {
          b.style.background = '#F1F5F9';
          b.style.color = 'var(--text-2)';
          const badge = b.querySelector('.badge');
          if (badge) {
            badge.style.background = '#CBD5E1';
            badge.style.color = '#334155';
          }
        }
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
    viewTableBtn.style.background = '#fff';
    viewTableBtn.style.boxShadow = '0 1px 2px rgba(0,0,0,0.05)';
    viewZoneBtn.style.background = 'transparent';
    viewZoneBtn.style.boxShadow = 'none';
    renderContent();
  });

  viewZoneBtn?.addEventListener('click', () => {
    currentView = 'zone';
    viewZoneBtn.style.background = '#fff';
    viewZoneBtn.style.boxShadow = '0 1px 2px rgba(0,0,0,0.05)';
    viewTableBtn.style.background = 'transparent';
    viewTableBtn.style.boxShadow = 'none';
    renderContent();
  });

  // Action Buttons
  container.querySelector('#btn-print-hygiene')?.addEventListener('click', printSOP);
  container.querySelector('#btn-export-hygiene')?.addEventListener('click', exportExcel);
  container.querySelector('#btn-add-hygiene')?.addEventListener('click', () => openItemModal(null));

  // Initial Load
  loadData();
}
