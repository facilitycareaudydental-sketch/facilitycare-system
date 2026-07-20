import { apiFetch } from '../config.js';

export async function renderCalendar(container) {
  let currentDate = new Date();
  let events = [];

  container.innerHTML = `
    <div class="page-header">
      <h1 class="page-title">📅 Kalender</h1>
    </div>
    <div class="card">
      <div class="card-header calendar-nav">
        <button class="btn btn-ghost btn-sm" id="cal-prev">‹ Prev</button>
        <span class="calendar-month-label" id="cal-month-label"></span>
        <button class="btn btn-ghost btn-sm" id="cal-next">Next ›</button>
        <div class="calendar-filters">
          <label class="filter-check"><input type="checkbox" value="schedule"        checked class="cal-filter"> Jadwal</label>
          <label class="filter-check"><input type="checkbox" value="issue"           checked class="cal-filter"> Permasalahan</label>
          <label class="filter-check"><input type="checkbox" value="reliever"        checked class="cal-filter"> Reliefer</label>
          <label class="filter-check"><input type="checkbox" value="training"        checked class="cal-filter"> Training</label>
          <label class="filter-check"><input type="checkbox" value="contract_expiry" checked class="cal-filter"> Kontrak Habis</label>
        </div>
      </div>
      <div class="card-body p-0">
        <div id="calendar-grid" style="min-height:400px"></div>
      </div>
    </div>
    <!-- Event detail popup -->
    <div id="cal-event-list" class="cal-event-sidebar" style="display:none">
      <div class="cal-event-header">
        <span id="cal-event-date"></span>
        <button class="btn btn-ghost btn-sm" id="cal-event-close">&times;</button>
      </div>
      <div id="cal-event-items"></div>
    </div>
  `;

  document.getElementById('cal-prev').addEventListener('click', () => {
    currentDate.setMonth(currentDate.getMonth() - 1);
    renderMonth();
  });
  document.getElementById('cal-next').addEventListener('click', () => {
    currentDate.setMonth(currentDate.getMonth() + 1);
    renderMonth();
  });
  document.getElementById('cal-event-close').addEventListener('click', () => {
    document.getElementById('cal-event-list').style.display = 'none';
  });
  document.querySelectorAll('.cal-filter').forEach(cb =>
    cb.addEventListener('change', renderMonth)
  );

  // ── Load events — always resolves, never throws ────────────────
  async function loadEvents() {
    try {
      const month = `${currentDate.getFullYear()}-${String(currentDate.getMonth() + 1).padStart(2, '0')}`;
      const res = await apiFetch(`/api/dashboard/calendar?month=${month}`);
      events = res.data?.data || [];
    } catch (err) {
      console.warn('[Calendar] Failed to load events, rendering empty grid:', err);
      events = []; // Render empty grid — no white screen
    }
  }

  // ── Build & inject grid HTML ────────────────────────────────────
  async function renderMonth() {
    const gridEl = document.getElementById('calendar-grid');
    if (!gridEl) return;

    // Show mini skeleton while loading
    gridEl.innerHTML = `<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:1px;background:var(--border);">
      ${Array(35).fill('<div style="background:#f8fafc;min-height:70px;"></div>').join('')}
    </div>`;

    await loadEvents();

    try {
      const year  = currentDate.getFullYear();
      const month = currentDate.getMonth();
      const monthLabel = currentDate.toLocaleDateString('id-ID', { month: 'long', year: 'numeric' });
      const labelEl = document.getElementById('cal-month-label');
      if (labelEl) labelEl.textContent = monthLabel;

      const activeFilters = new Set(
        Array.from(document.querySelectorAll('.cal-filter:checked')).map(cb => cb.value)
      );
      const filteredEvents = events.filter(e => activeFilters.has(e.type));

      // Group events by date
      const byDate = {};
      filteredEvents.forEach(e => {
        const d = (e.event_date || '').slice(0, 10);
        if (!byDate[d]) byDate[d] = [];
        byDate[d].push(e);
      });

      const firstDay    = new Date(year, month, 1).getDay();
      const daysInMonth = new Date(year, month + 1, 0).getDate();
      const dayNames    = ['Min', 'Sen', 'Sel', 'Rab', 'Kam', 'Jum', 'Sab'];
      const todayStr    = new Date().toISOString().slice(0, 10);

      let html = `<div class="calendar-grid">`;

      // Day name headers
      dayNames.forEach(d => {
        html += `<div class="cal-day-header">${d}</div>`;
      });

      // Empty cells before first day of month
      for (let i = 0; i < firstDay; i++) {
        html += `<div class="cal-cell cal-cell-empty"></div>`;
      }

      // Day cells
      for (let d = 1; d <= daysInMonth; d++) {
        const dateStr  = `${year}-${String(month + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
        const dayEvts  = byDate[dateStr] || [];
        const isToday  = dateStr === todayStr;

        html += `
          <div class="cal-cell ${isToday ? 'cal-today' : ''} ${dayEvts.length ? 'cal-has-events' : ''}"
               data-date="${dateStr}" tabindex="0" role="button" aria-label="${dateStr}">
            <div class="cal-day-num ${isToday ? 'today-num' : ''}">${d}</div>
            <div class="cal-events-preview">
              ${dayEvts.slice(0, 3).map(e => `
                <div class="cal-event-dot cal-color-${e.color || 'gray'}" title="${escHtml(e.title || e.type)}">
                  <span class="cal-event-dot-label">${truncate(e.title || e.branch_name || e.type, 18)}</span>
                </div>
              `).join('')}
              ${dayEvts.length > 3 ? `<div class="cal-more">+${dayEvts.length - 3} lagi</div>` : ''}
            </div>
          </div>`;
      }

      // Fill remaining cells to complete the last row (7-col grid)
      const totalCells = firstDay + daysInMonth;
      const remainder  = totalCells % 7;
      if (remainder !== 0) {
        for (let i = 0; i < (7 - remainder); i++) {
          html += `<div class="cal-cell cal-cell-empty"></div>`;
        }
      }

      html += `</div>`;
      gridEl.innerHTML = html;

      // Click handler for event detail
      gridEl.querySelectorAll('.cal-cell[data-date]').forEach(cell => {
        cell.addEventListener('click', () => {
          const date    = cell.dataset.date;
          const dayEvts = byDate[date] || [];
          if (!dayEvts.length) return;

          const sidebar   = document.getElementById('cal-event-list');
          const dateLabel = new Date(date + 'T00:00:00').toLocaleDateString('id-ID', {
            weekday: 'long', day: 'numeric', month: 'long', year: 'numeric'
          });
          document.getElementById('cal-event-date').textContent = dateLabel;
          document.getElementById('cal-event-items').innerHTML = dayEvts.map(e => `
            <div class="cal-event-item cal-color-border-${e.color || 'gray'}">
              <div class="cal-event-type">${typeLabel(e.type)}</div>
              <div class="cal-event-title">${escHtml(e.title || '-')}</div>
              <div class="cal-event-branch">${escHtml(e.branch_name || '')}</div>
              ${e.status ? `<div class="cal-event-status">${escHtml(e.status)}</div>` : ''}
              ${e.days_remaining !== undefined
                ? `<div class="cal-event-extra">Sisa: ${e.days_remaining} hari</div>`
                : ''}
            </div>
          `).join('');
          sidebar.style.display = 'block';
        });
      });

    } catch (err) {
      console.error('[Calendar] Render error:', err);
      if (gridEl) {
        gridEl.innerHTML = `
          <div style="padding:40px;text-align:center;color:var(--text-3)">
            <div style="font-size:2rem;margin-bottom:8px">📅</div>
            <div>Gagal memuat kalender. Silakan refresh.</div>
          </div>`;
      }
    }
  }

  // ── Initial render ─────────────────────────────────────────────
  renderMonth();
}

// ── Helpers ────────────────────────────────────────────────────────
function truncate(str, len) {
  if (!str) return '';
  return str.length > len ? str.slice(0, len) + '…' : str;
}

function escHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function typeLabel(type) {
  const map = {
    schedule:        '🗓 Jadwal',
    issue:           '⚠️ Permasalahan',
    reliever:        '🔄 Reliefer',
    training:        '🎓 Training',
    contract_expiry: '📋 Kontrak Habis',
  };
  return map[type] || type;
}
