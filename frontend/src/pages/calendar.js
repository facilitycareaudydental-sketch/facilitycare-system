import { apiFetch } from '../config.js';

export async function renderCalendar(container) {
  let currentDate = new Date();
  let events = [];

  container.innerHTML = `
    <div class="page-header">
      <h1 class="page-title">Kalender</h1>
    </div>
    <div class="card">
      <div class="card-header calendar-nav">
        <button class="btn btn-ghost btn-sm" id="cal-prev">‹ Prev</button>
        <span class="calendar-month-label" id="cal-month-label"></span>
        <button class="btn btn-ghost btn-sm" id="cal-next">Next ›</button>
        <div class="calendar-filters">
          <label class="filter-check"><input type="checkbox" value="schedule" checked class="cal-filter"> Jadwal</label>
          <label class="filter-check"><input type="checkbox" value="issue" checked class="cal-filter"> Permasalahan</label>
          <label class="filter-check"><input type="checkbox" value="reliever" checked class="cal-filter"> Reliefer</label>
          <label class="filter-check"><input type="checkbox" value="training" checked class="cal-filter"> Training</label>
          <label class="filter-check"><input type="checkbox" value="contract_expiry" checked class="cal-filter"> Kontrak Habis</label>
        </div>
      </div>
      <div class="card-body p-0">
        <div id="calendar-grid"></div>
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
  document.querySelectorAll('.cal-filter').forEach(cb => cb.addEventListener('change', renderMonth));

  async function loadEvents() {
    const month = `${currentDate.getFullYear()}-${String(currentDate.getMonth() + 1).padStart(2, '0')}`;
    const res = await apiFetch(`/api/dashboard/calendar?month=${month}`);
    events = res.data?.data || [];
  }

  async function renderMonth() {
    await loadEvents();

    const year = currentDate.getFullYear();
    const month = currentDate.getMonth();
    const monthLabel = currentDate.toLocaleDateString('id-ID', { month: 'long', year: 'numeric' });
    document.getElementById('cal-month-label').textContent = monthLabel;

    const activeFilters = new Set(
      Array.from(document.querySelectorAll('.cal-filter:checked')).map(cb => cb.value)
    );

    const filteredEvents = events.filter(e => activeFilters.has(e.type));

    // Group events by date
    const eventsByDate = {};
    filteredEvents.forEach(e => {
      const d = (e.event_date || '').slice(0, 10);
      if (!eventsByDate[d]) eventsByDate[d] = [];
      eventsByDate[d].push(e);
    });

    const firstDay = new Date(year, month, 1).getDay();
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const days = ['Min', 'Sen', 'Sel', 'Rab', 'Kam', 'Jum', 'Sab'];

    let html = `<div class="calendar-grid">`;
    // Header
    days.forEach(d => { html += `<div class="cal-day-header">${d}</div>`; });
    // Empty cells before first day
    for (let i = 0; i < firstDay; i++) html += `<div class="cal-cell cal-cell-empty"></div>`;

    const todayStr = new Date().toISOString().slice(0, 10);

    for (let d = 1; d <= daysInMonth; d++) {
      const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
      const dayEvents = eventsByDate[dateStr] || [];
      const isToday = dateStr === todayStr;
      html += `
        <div class="cal-cell ${isToday ? 'cal-today' : ''} ${dayEvents.length ? 'cal-has-events' : ''}" 
             data-date="${dateStr}" tabindex="0" role="button" aria-label="${dateStr}">
          <div class="cal-day-num ${isToday ? 'today-num' : ''}">${d}</div>
          <div class="cal-events-preview">
            ${dayEvents.slice(0, 3).map(e => `
              <div class="cal-event-dot cal-color-${e.color || 'gray'}" title="${e.title || e.type}">
                <span class="cal-event-dot-label">${truncate(e.title || e.branch_name || e.type, 18)}</span>
              </div>
            `).join('')}
            ${dayEvents.length > 3 ? `<div class="cal-more">+${dayEvents.length - 3} lagi</div>` : ''}
          </div>
        </div>`;
    }
    html += `</div>`;

    document.getElementById('calendar-grid').innerHTML = html;

    // Event detail on click
    document.querySelectorAll('.cal-cell[data-date]').forEach(cell => {
      cell.addEventListener('click', () => {
        const date = cell.dataset.date;
        const dayEvts = eventsByDate[date] || [];
        if (!dayEvts.length) return;

        const sidebar = document.getElementById('cal-event-list');
        document.getElementById('cal-event-date').textContent = new Date(date + 'T00:00:00').toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
        document.getElementById('cal-event-items').innerHTML = dayEvts.map(e => `
          <div class="cal-event-item cal-color-border-${e.color || 'gray'}">
            <div class="cal-event-type">${typeLabel(e.type)}</div>
            <div class="cal-event-title">${e.title || '-'}</div>
            <div class="cal-event-branch">${e.branch_name || ''}</div>
            ${e.status ? `<div class="cal-event-status">${e.status}</div>` : ''}
            ${e.days_remaining !== undefined ? `<div class="cal-event-extra">Sisa: ${e.days_remaining} hari</div>` : ''}
          </div>
        `).join('');
        sidebar.style.display = 'block';
      });
    });
  }

  renderMonth();
}

function truncate(str, len) {
  return str && str.length > len ? str.slice(0, len) + '…' : (str || '');
}

function typeLabel(type) {
  const map = {
    schedule: '🗓 Jadwal',
    issue: '⚠️ Permasalahan',
    reliever: '🔄 Reliefer',
    training: '🎓 Training',
    contract_expiry: '📋 Kontrak Habis',
  };
  return map[type] || type;
}
