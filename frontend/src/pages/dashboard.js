/**
 * dashboard.js — FCMS Minimalist Premium Navy Enterprise Dashboard (Frontend)
 * ─────────────────────────────────────────────────────────────────────────────
 * Modernized UI/UX based on official visual references:
 * - Acuan Utama: Minimalist Premium Navy Enterprise Layout (media_1795290926639_b09818b0.jpg)
 * - Acuan Navigasi: Deep Navy Sidebar (Deep Navy background, Royal Blue active pill)
 * - Acuan KPI: Gradient multi-color accents (Purple, Teal, Amber, Blue, Red)
 *
 * Guarantees:
 * - 100% Vanilla JavaScript SPA architecture
 * - Real data from existing FCMS API endpoints (Zero dummy numbers)
 * - Safe parallel data fetching (Promise.allSettled / safeFetch with timeouts)
 * - Zero breaking changes to backend, database, or RBAC
 * ─────────────────────────────────────────────────────────────────────────────
 */
import { apiFetch, getUser } from '../config.js';
import { escapeHtml } from '../utils/sanitize.js';

// ── Chart registry ─────────────────────────────────────────────────────────
const _charts = {};
function destroyChart(key) {
  if (_charts[key]) {
    try { _charts[key].destroy(); } catch {}
    delete _charts[key];
  }
}
function destroyAllCharts() {
  Object.keys(_charts).forEach(destroyChart);
}

// ── Safe value helpers ─────────────────────────────────────────────────────
const safeNum = (v, fb = 0) => {
  const n = Number(v);
  return isNaN(n) || v === null || v === undefined ? fb : n;
};
const safeStr = (v, fb = '—') => {
  if (v === null || v === undefined || v === '') return fb;
  const s = String(v).trim();
  return (s === '' || s === '[object Object]') ? fb : escapeHtml(s);
};

// ── Counter animation ──────────────────────────────────────────────────────
function animateCount(el, target, ms = 800) {
  if (!el) return;
  if (el._animFrame) cancelAnimationFrame(el._animFrame);
  const t = Math.max(0, Math.round(safeNum(target)));
  const curText = el.textContent ? el.textContent.replace(/[^\d]/g, '') : '';
  const current = curText !== '' ? parseInt(curText) || 0 : 0;
  if (t === current) {
    el.textContent = t.toLocaleString('id-ID');
    el._animFrame = null;
    return;
  }
  const start = Date.now();
  const tick = () => {
    const p    = Math.min((Date.now() - start) / ms, 1);
    const ease = 1 - Math.pow(1 - p, 3);
    el.textContent = Math.round(current + ease * (t - current)).toLocaleString('id-ID');
    if (p < 1) {
      el._animFrame = requestAnimationFrame(tick);
    } else {
      el.textContent = t.toLocaleString('id-ID');
      el._animFrame = null;
    }
  };
  el._animFrame = requestAnimationFrame(tick);
}

// ── Chart.js Clean Navy Blue Defaults ────────────────────────────────────────
const FONT  = { family: 'Inter', size: 10, weight: '500' };
const TICK  = '#CBD5E1';
const GRID  = 'rgba(148, 163, 184, 0.10)';
const DONUT_COLORS = ['#2563EB', '#38BDF8', '#60A5FA', '#93C5FD', '#1D4ED8', '#0EA5E9'];

const isMobile = () => window.innerWidth < 768;

function chartOpts(extra = {}) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    animation: { duration: 600, easing: 'easeOutQuart' },
    plugins: {
      legend: {
        position: isMobile() ? 'bottom' : 'top',
        labels: {
          font: FONT,
          color: TICK,
          usePointStyle: true,
          padding: 10,
          boxWidth: 8,
          boxHeight: 8,
        },
      },
      tooltip: {
        mode: 'index',
        intersect: false,
        backgroundColor: '#1E293B',
        titleColor: '#F8FAFC',
        bodyColor: '#94A3B8',
        borderColor: 'rgba(255, 255, 255, 0.1)',
        borderWidth: 1,
        padding: 10,
        boxPadding: 4,
        cornerRadius: 8,
        bodyFont: FONT,
        titleFont: { ...FONT, weight: '700' },
      },
    },
    scales: {
      x: { grid: { color: GRID }, ticks: { font: FONT, color: TICK, maxRotation: 0 } },
      y: { grid: { color: GRID }, ticks: { font: FONT, color: TICK }, beginAtZero: true },
    },
    ...extra,
  };
}

// ── Skeleton HTML ──────────────────────────────────────────────────────────
const skelKPI = () => Array(5).fill(0).map(() => `
  <div class="kpi-card" style="pointer-events:none;">
    <div class="skeleton" style="width:46px;height:46px;border-radius:12px;flex-shrink:0;"></div>
    <div style="flex:1;">
      <div class="skeleton skeleton-text" style="width:50px;height:24px;margin-bottom:6px;"></div>
      <div class="skeleton skeleton-text" style="width:85px;height:12px;"></div>
    </div>
  </div>`).join('');

function skelTable(rows = 4) {
  return Array(rows).fill(0).map((_, i) =>
    `<div class="skeleton skeleton-text" style="height:36px;margin-bottom:${i < rows - 1 ? '8px' : '0'};border-radius:8px;"></div>`
  ).join('');
}

// ── Safe API fetch with timeout + fallback ─────────────────────────────────
async function safeFetch(path, fallback, timeoutMs = 8000) {
  try {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    const res = await apiFetch(path, { signal: controller.signal }).catch(() => null);
    clearTimeout(timer);
    if (!res || !res.ok) return fallback;
    const d = res.data;
    if (!d) return fallback;
    return d.data !== undefined ? (d.data ?? fallback) : d;
  } catch {
    return fallback;
  }
}

// ── Force remove all skeletons (safety net) ────────────────────────────────
function forceRemoveSkeletons(container) {
  if (!container || !container.isConnected) return;
  ['skel-donut', 'skel-trend', 'skel-jadwal'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.display = 'none';
  });
  ['chart-donut', 'chart-trend', 'chart-jadwal'].forEach(id => {
    const el = document.getElementById(id);
    if (el && el.style.display === 'none') {
      el.style.display = 'block';
    }
  });
  const kpiRow = document.getElementById('kpi-row');
  if (kpiRow && kpiRow.querySelector('.skeleton')) {
    renderKPI({});
  }
  const act = document.getElementById('table-recent-activity');
  if (act && act.querySelector('.skeleton')) {
    act.innerHTML = `<div class="chart-empty">Belum ada aktivitas terbaru</div>`;
  }
  const sched = document.getElementById('list-upcoming-schedule');
  if (sched && sched.querySelector('.skeleton')) {
    sched.innerHTML = `<div class="chart-empty">Belum ada jadwal kegiatan</div>`;
  }
}

// ── Main Render ────────────────────────────────────────────────────────────
export async function renderDashboard(container) {
  if (!container) container = document.getElementById('main-content');
  if (!container) return;
  destroyAllCharts();
  if (container._dashRefresh) clearInterval(container._dashRefresh);
  if (container._skelTimeout) clearTimeout(container._skelTimeout);

  let disposed = false;
  const isAlive = () => !disposed && container && container.isConnected && Boolean(container.querySelector('#dash-root'));
  const cleanup = () => {
    disposed = true;
    if (container._skelTimeout) { clearTimeout(container._skelTimeout); container._skelTimeout = null; }
    if (container._dashRefresh) { clearInterval(container._dashRefresh); container._dashRefresh = null; }
    destroyAllCharts();
  };
  window.addEventListener('hashchange', cleanup, { once: true });

  const curUser = getUser();
  const isSuperadmin = String(curUser?.role || '').toLowerCase() === 'superadmin';
  const now = new Date();
  const curDay = String(now.getDate()).padStart(2, '0');
  const curMonth = now.toLocaleDateString('id-ID', { month: 'short' });
  const curYear = now.getFullYear();
  const dateRangeStr = `01 ${curMonth} – ${curDay} ${curMonth} ${curYear}`;

  container.innerHTML = `
    <div class="dashboard-wrap" id="dash-root">
      
      <!-- Dashboard Operasional Header -->
      <div class="dash-operasional-header">
        <div>
          <h1 class="dash-title">Dashboard Operasional</h1>
          <p class="dash-subtitle">Monitoring data personel, kegiatan, dan operasional Facility Care.</p>
        </div>
        <div class="dash-header-controls">
          <div class="dash-filter-pill">
            <select id="filter-jadwal-year" class="dash-header-select" aria-label="Pilih Tahun">
              <option value="${curYear}" selected>${curYear}</option>
              <option value="${curYear + 1}">${curYear + 1}</option>
            </select>
          </div>
          <div class="dash-filter-pill dash-date-display">
            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
            <span id="dash-date-range">${dateRangeStr}</span>
          </div>
          <button id="btn-dash-refresh" class="dash-refresh-btn" title="Muat ulang data">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/></svg>
            Refresh Data
          </button>
        </div>
      </div>

      <!-- Top KPI Row (5 Cards) -->
      <div class="kpi-row" id="kpi-row">${skelKPI()}</div>

      <!-- Middle Charts Row (3 Columns) -->
      <div class="charts-row">
        <!-- Chart 1: Tren Kegiatan Operasional -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">Tren Kegiatan Operasional</div>
            </div>
            <select id="filter-jadwal-period" class="dash-header-select" style="background:#0E182B; border:1px solid #223456; border-radius:6px; padding:4px 8px; font-size:0.78rem; color:#F8FAFC;" aria-label="Filter Periode">
              <option value="month">Per Bulan</option>
            </select>
          </div>
          <div class="chart-canvas-wrap" style="height:150px;position:relative;margin-top:4px">
            <div id="skel-jadwal" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-jadwal" style="display:none"></canvas>
          </div>
          <div id="jadwal-legend" style="display:flex;justify-content:center;gap:12px;margin-top:6px;font-size:0.76rem;font-weight:600;color:#CBD5E1;flex-wrap:wrap;">
            <div style="display:flex;align-items:center;gap:5px"><div style="width:8px;height:8px;border-radius:50%;background:#2563EB"></div> Inspeksi</div>
            <div style="display:flex;align-items:center;gap:5px"><div style="width:8px;height:8px;border-radius:50%;background:#3B82F6"></div> General Cleaning</div>
            <div style="display:flex;align-items:center;gap:5px"><div style="width:8px;height:8px;border-radius:50%;background:#38BDF8"></div> Deep Cleaning</div>
            <div style="display:flex;align-items:center;gap:5px"><div style="width:8px;height:8px;border-radius:50%;background:#93C5FD"></div> Fogging</div>
          </div>
        </div>

        <!-- Chart 2: Permasalahan per Kategori -->
        <div class="chart-card" style="display:flex; flex-direction:column; justify-content:space-between;">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan per Kategori</div>
          </div>
          <div style="display:flex; align-items:center; justify-content:center; gap:16px; margin: 4px 0;">
            <div class="chart-canvas-wrap" style="width:120px;height:120px;position:relative">
              <div id="skel-donut" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
              <canvas id="chart-donut" style="display:none"></canvas>
            </div>
            <div id="donut-legend" class="donut-legend" style="min-width:110px;"></div>
          </div>
          <div style="text-align:center; font-size:0.78rem; color:#CBD5E1; display:flex; align-items:center; justify-content:center; gap:6px;">
            <svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/></svg>
            <span id="donut-period-text">Distribusi Kategori Permasalahan</span>
          </div>
        </div>

        <!-- Chart 3: Tren Permasalahan 12 Bulan -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Tren Permasalahan 12 Bulan</div>
            <div style="display:flex;align-items:center;gap:12px;font-size:0.76rem;font-weight:600;color:#CBD5E1">
               <div style="display:flex;align-items:center;gap:5px"><div style="width:8px;height:8px;border-radius:50%;background:#38BDF8"></div> Open</div>
               <div style="display:flex;align-items:center;gap:5px"><div style="width:8px;height:8px;border-radius:50%;background:#2563EB"></div> Closed</div>
            </div>
          </div>
          <div class="chart-canvas-wrap" style="height:150px;position:relative;margin-top:4px">
            <div id="skel-trend" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-trend" style="display:none"></canvas>
          </div>
        </div>
      </div>

      <!-- Bottom Row (2 Columns: Aktivitas Terbaru & Jadwal Kegiatan Mendatang) -->
      <div class="bottom-row">
        <!-- Aktivitas Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Aktivitas Terbaru</div>
            ${isSuperadmin ? '<a href="#/audit-logs" class="chart-link">Lihat Semua &rarr;</a>' : ''}
          </div>
          <div id="table-recent-activity" class="dash-table-wrap dash-panel-scroll">
            ${skelTable(4)}
          </div>
        </div>

        <!-- Jadwal Kegiatan Mendatang -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Jadwal Kegiatan Mendatang</div>
            <a href="#/calendar" class="chart-link">Lihat Kalender &rarr;</a>
          </div>
          <div id="list-upcoming-schedule" class="dash-panel-scroll" style="display:flex;flex-direction:column;gap:4px">
            ${skelTable(4)}
          </div>
        </div>
      </div>

    </div>
  `;

  document.getElementById('btn-dash-refresh')
    ?.addEventListener('click', () => { if (isAlive()) fetchAll(container, isAlive); });

  document.getElementById('filter-jadwal-year')?.addEventListener('change', async (e) => {
    const year = e.target.value;
    const skel = document.getElementById('skel-jadwal');
    const cvs = document.getElementById('chart-jadwal');
    if (skel) { skel.style.display = 'block'; }
    if (cvs) cvs.style.display = 'none';

    const jadwalData = await safeFetch(`/api/dashboard/schedule-chart?year=${year}`, {}, 8000);
    if (!isAlive()) return;
    try {
      renderScheduleChart(jadwalData);
    } catch(err) {
      console.warn('ScheduleChart render error:', err);
      hideSkel('skel-jadwal', 'chart-jadwal');
    }
  });

  // Safety net: force remove skeletons after 5 seconds
  container._skelTimeout = setTimeout(() => forceRemoveSkeletons(container), 5000);

  // Load live data
  await fetchAll(container, isAlive);
  return cleanup;
}

// ── Parallel Data Fetching ─────────────────────────────────────────────────
async function fetchAll(container, isAlive = () => container && container.isConnected) {
  if (container && container._skelTimeout) {
    clearTimeout(container._skelTimeout);
    container._skelTimeout = null;
  }

  const selectedYear = document.getElementById('filter-jadwal-year')?.value || new Date().getFullYear();

  // Fire requests concurrently in parallel wave
  const [kpi, trend, issuesSum, recentIssues, calendarData, scheduleChartData, aggregatesData, activityLogData] =
    await Promise.all([
      safeFetch('/api/dashboard/kpi',               {}, 8000),
      safeFetch('/api/dashboard/issues-trend',       {}, 8000),
      safeFetch('/api/dashboard/issues-summary',     {}, 8000),
      safeFetch('/api/dashboard/stats',              {}, 8000),
      safeFetch('/api/dashboard/calendar',           [], 8000),
      safeFetch(`/api/dashboard/schedule-chart?year=${selectedYear}`, {}, 8000),
      safeFetch('/api/dashboard/dropdown-aggregates', {}, 8000),
      safeFetch('/api/dashboard/activity-log',       [], 8000),
    ]);

  if (!isAlive()) return;
  window.dashboardAggregates = aggregatesData || {};

  // Render each section independently
  try { renderKPI(kpi); } catch(e) { console.warn('KPI render:', e); }
  try { renderScheduleChart(scheduleChartData); } catch(e) { console.warn('ScheduleChart render:', e); hideSkel('skel-jadwal','chart-jadwal'); }
  try { renderDonut(Array.isArray(issuesSum?.by_category) ? issuesSum.by_category : []); } catch(e) { console.warn('Donut render:', e); hideSkel('skel-donut','chart-donut'); }
  try { renderTrend(trend); } catch(e) { console.warn('Trend render:', e); hideSkel('skel-trend','chart-trend'); }

  try {
    renderRecentActivity(activityLogData);
  } catch(e) { console.warn('RecentActivity render:', e); }

  try {
    const upcoming = recentIssues?.upcoming_schedule || [];
    renderUpcomingSchedule(upcoming, calendarData);
  } catch(e) { console.warn('UpcomingSchedule render:', e); }
}

// ── KPI Cards (5 Uniform Clean Navy Cards) ─────────────────────────────────
function renderKPI(kpi) {
  const row = document.getElementById('kpi-row');
  if (!row) return;
  kpi = kpi || {};

  const cards = [
    {
      iconSvg: '<svg width="21" height="21" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>',
      label: 'Karyawan Aktif',
      val: kpi.employees?.current ?? 0,
      href: '#/employees?dash_filter=active',
    },
    {
      iconSvg: '<svg width="21" height="21" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>',
      label: 'Kontrak Aktif',
      val: kpi.contracts?.current ?? 0,
      href: '#/contracts?dash_filter=active',
    },
    {
      iconSvg: '<svg width="21" height="21" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M2 9h20L19 3H5L2 9z"/><path d="M4 9v11a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9"/><path d="M10 22V12h4v10"/></svg>',
      label: 'Cabang Operasional',
      val: kpi.branches?.current ?? 0,
      href: '#/branches',
    },
    {
      iconSvg: '<svg width="21" height="21" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>',
      label: 'Jadwal Kegiatan',
      val: kpi.schedule?.current ?? 0,
      href: '#/timeline',
    },
    {
      iconSvg: '<svg width="21" height="21" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>',
      label: 'Permasalahan',
      val: kpi.issues?.current ?? 0,
      href: '#/issues?dash_filter=open',
    },
  ];

  row.innerHTML = cards.map(c => `
    <a href="${c.href}" class="kpi-card">
      <div class="kpi-icon-wrap">
        ${c.iconSvg}
      </div>
      <div class="kpi-body">
        <div class="kpi-value" data-target="${c.val}">${c.val}</div>
        <div class="kpi-label">${c.label}</div>
      </div>
    </a>
  `).join('');

  row.querySelectorAll('.kpi-value').forEach(el => {
    animateCount(el, parseInt(el.dataset.target) || 0);
  });
}

// ── Chart 1: Schedule Stacked Bar Chart ─────────────────────────────────────
function renderScheduleChart(data) {
  hideSkel('skel-jadwal', 'chart-jadwal');
  const canvas = document.getElementById('chart-jadwal');
  if (!canvas) return;
  destroyChart('jadwal');
  data = data || {};
  const labels = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  const isEmpty = !Object.values(data).some(arr => Array.isArray(arr) && arr.some(v => v > 0));
  if (isEmpty) {
    showEmpty(canvas, 'Belum ada data jadwal kegiatan');
    return;
  }

  const insp = data['Inspeksi Hygiene'] || Array(12).fill(0);
  const gcl  = data['General Cleaning'] || Array(12).fill(0);
  const dcl  = data['Deep Cleaning']    || Array(12).fill(0);
  const fog  = data['Fogging']          || Array(12).fill(0);

  _charts.jadwal = new Chart(canvas, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Inspeksi',         data: insp, backgroundColor: '#2563EB', borderRadius: 3 },
        { label: 'General Cleaning', data: gcl,  backgroundColor: '#3B82F6', borderRadius: 3 },
        { label: 'Deep Cleaning',    data: dcl,  backgroundColor: '#38BDF8', borderRadius: 3 },
        { label: 'Fogging',          data: fog,  backgroundColor: '#93C5FD', borderRadius: 3 },
      ]
    },
    options: chartOpts({
      plugins: { legend: { display: false } },
      datasets: { bar: { barPercentage: 0.75, categoryPercentage: 0.85 } },
      scales: {
        x: { stacked: true, grid: { display: false }, ticks: { font: { family: 'Inter', size: 9 }, color: TICK, maxRotation: 0 } },
        y: { stacked: true, grid: { color: GRID }, ticks: { font: { family: 'Inter', size: 9 }, color: TICK }, min: 0 }
      }
    }),
  });
}

// ── Chart 2: Category Donut Chart ──────────────────────────────────────────
function renderDonut(categories) {
  hideSkel('skel-donut', 'chart-donut');
  const canvas = document.getElementById('chart-donut');
  const legendWrap = document.getElementById('donut-legend');
  if (!canvas || !legendWrap) return;
  destroyChart('donut');

  const data = (categories || []).filter(c => safeNum(c.count) > 0);
  if (!data.length) {
    showEmpty(canvas, 'Belum ada data kategori');
    return;
  }

  const labels = data.map(c => safeStr(c.category, 'Lainnya'));
  const values = data.map(c => safeNum(c.count));
  const total  = values.reduce((a, b) => a + b, 0);

  // Render Custom Legend
  legendWrap.innerHTML = data.map((c, i) => {
    const color = DONUT_COLORS[i % DONUT_COLORS.length];
    const pct = total > 0 ? Math.round((c.count / total) * 100) : 0;
    return `
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
        <div style="width:8px;height:8px;border-radius:50%;background:${color};flex-shrink:0;"></div>
        <div style="display:flex;align-items:baseline;gap:4px;font-size:0.75rem;">
          <span style="font-weight:600;color:var(--text-1);">${labels[i]}</span>
          <span style="color:var(--text-2);font-weight:700;">${c.count}</span>
          <span style="font-size:0.68rem;color:var(--text-3);">(${pct}%)</span>
        </div>
      </div>
    `;
  }).join('');

  // Center Text Plugin
  const centerTextPlugin = {
    id: 'centerText',
    beforeDraw: function(chart) {
      const width = chart.width, height = chart.height, ctx = chart.ctx;
      ctx.restore();
      const fontSize = (height / 85).toFixed(2);
      ctx.font = "bold " + fontSize + "em Inter";
      ctx.textBaseline = "middle";
      ctx.fillStyle = "#F8FAFC";
      const text = total.toString();
      const textX = Math.round((width - ctx.measureText(text).width) / 2);
      const textY = height / 2;
      ctx.fillText(text, textX, textY - 6);

      ctx.font = "600 " + (fontSize * 0.38).toFixed(2) + "em Inter";
      ctx.fillStyle = "#94A3B8";
      const labelText = "Total";
      const labelX = Math.round((width - ctx.measureText(labelText).width) / 2);
      ctx.fillText(labelText, labelX, textY + 10);
      ctx.save();
    }
  };

  _charts.donut = new Chart(canvas, {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{
        data: values,
        backgroundColor: DONUT_COLORS,
        borderWidth: 2,
        borderColor: '#131F36',
        hoverBorderColor: '#131F36'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 600 },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: '#1E293B',
          titleColor: '#F8FAFC',
          bodyColor: '#94A3B8',
          borderColor: 'rgba(255, 255, 255, 0.1)',
          borderWidth: 1,
          cornerRadius: 8,
          bodyFont: FONT,
          callbacks: { label: ctx => ` ${ctx.label}: ${ctx.parsed} kasus` }
        },
      },
      cutout: '75%',
    },
    plugins: [centerTextPlugin]
  });
}

// ── Chart 3: Trend Line Chart (12 Months Open vs Closed) ───────────────────
function renderTrend(trend) {
  hideSkel('skel-trend', 'chart-trend');
  const canvas = document.getElementById('chart-trend');
  if (!canvas) return;
  destroyChart('trend');
  trend = trend || {};

  const MONTH_ID = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  const labels = (trend.labels || []).map(ym => {
    if (!ym || typeof ym !== 'string') return '';
    try {
      const [, m] = ym.split('-');
      return MONTH_ID[Number(m) - 1] || m;
    } catch { return ym; }
  });
  const open   = (trend.open   || []).map(v => safeNum(v));
  const closed = (trend.closed || []).map(v => safeNum(v));
  if (!labels.length) {
    showEmpty(canvas, 'Belum ada data trend permasalahan');
    return;
  }

  _charts.trend = new Chart(canvas, {
    type: 'line',
    data: {
      labels,
      datasets: [
        {
          label: 'Open',
          data: open,
          borderColor: '#38BDF8',
          backgroundColor: 'rgba(56,189,248,.12)',
          fill: true,
          tension: 0.4,
          pointRadius: 3.5,
          pointHoverRadius: 6,
          pointBackgroundColor: '#38BDF8',
          borderWidth: 2
        },
        {
          label: 'Closed',
          data: closed,
          borderColor: '#2563EB',
          backgroundColor: 'rgba(37,99,235,.12)',
          fill: true,
          tension: 0.4,
          pointRadius: 3.5,
          pointHoverRadius: 6,
          pointBackgroundColor: '#2563EB',
          borderWidth: 2
        },
      ]
    },
    options: chartOpts({
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { display: false }, ticks: { font: { family: 'Inter', size: 9 }, color: TICK, maxRotation: 0 } },
        y: { grid: { color: GRID }, ticks: { font: { family: 'Inter', size: 9 }, color: TICK }, beginAtZero: true }
      }
    }),
  });
}

// ── Bottom Panel 1: Aktivitas Terbaru ──────────────────────────────────────
function renderRecentActivity(rows) {
  const wrap = document.getElementById('table-recent-activity');
  if (!wrap) return;
  const items = (rows || []).slice(0, 6);
  if (!items.length) {
    wrap.innerHTML = `<div class="chart-empty">Belum ada aktivitas terbaru</div>`;
    return;
  }

  const typeConfig = {
    issue:      { icon: '<svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>', bg: '#2563EB', text: '#FFFFFF', module: 'Permasalahan', modBg: 'rgba(56,189,248,0.16)', modCol: '#7DD3FC' },
    contract:   { icon: '<svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>', bg: '#2563EB', text: '#FFFFFF', module: 'Data Kontrak', modBg: 'rgba(37,99,235,0.18)', modCol: '#93C5FD' },
    employee:   { icon: '<svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>', bg: '#2563EB', text: '#FFFFFF', module: 'Master Karyawan', modBg: 'rgba(37,99,235,0.18)', modCol: '#93C5FD' },
    training:   { icon: '<svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>', bg: '#2563EB', text: '#FFFFFF', module: 'Training', modBg: 'rgba(56,189,248,0.16)', modCol: '#7DD3FC' },
    inspection: { icon: '<svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>', bg: '#2563EB', text: '#FFFFFF', module: 'Inspeksi', modBg: 'rgba(37,99,235,0.18)', modCol: '#93C5FD' },
    schedule:   { icon: '<svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>', bg: '#2563EB', text: '#FFFFFF', module: 'Jadwal Kegiatan', modBg: 'rgba(56,189,248,0.16)', modCol: '#7DD3FC' },
  };

  wrap.innerHTML = `
    <table class="act-table">
      <thead>
        <tr>
          <th>Waktu</th>
          <th>Aktivitas</th>
          <th>Modul</th>
          <th>Pengguna</th>
        </tr>
      </thead>
      <tbody>
        ${items.map(r => {
          const parsedIso = r.created_at
            ? (String(r.created_at).includes(' ') && !String(r.created_at).includes('T')
                ? `${String(r.created_at).replace(' ', 'T')}Z`
                : String(r.created_at))
            : null;
          const dtObj = parsedIso ? new Date(parsedIso) : null;
          const timeStr = dtObj && !isNaN(dtObj.getTime())
            ? dtObj.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', timeZone: 'Asia/Jakarta' }).replace(/\./g, ':')
            : '—';
          const cfg = typeConfig[r.type] || { icon: '📌', bg: '#3B82F6', text: '#fff', module: 'Operasional', modBg: 'rgba(59,130,246,0.18)', modCol: '#60A5FA' };
          const userName = r.user_name || (r.branch ? `Area ${r.branch}` : 'Sistem FCMS');
          const userInitial = (userName || 'F')[0].toUpperCase();
          return `
            <tr>
              <td class="act-time">${timeStr}</td>
              <td>
                <div class="act-desc">
                  <span class="act-icon" style="background:${cfg.bg};color:${cfg.text}">${cfg.icon}</span>
                  <span title="${escapeHtml(r.label)}">${escapeHtml(r.label)}</span>
                </div>
              </td>
              <td>
                <span class="act-modul-pill" style="background:${cfg.modBg};color:${cfg.modCol}">
                  ${cfg.module}
                </span>
              </td>
              <td>
                <div class="act-user">
                  <span class="act-user-avatar" style="background:${cfg.bg}">
                    ${userInitial}
                  </span>
                  <span>${escapeHtml(userName)}</span>
                </div>
              </td>
            </tr>
          `;
        }).join('')}
      </tbody>
    </table>
  `;
}

// ── Bottom Panel 2: Jadwal Kegiatan Mendatang ──────────────────────────────
function renderUpcomingSchedule(upcoming, calendar) {
  const wrap = document.getElementById('list-upcoming-schedule');
  if (!wrap) return;

  let items = Array.isArray(upcoming) && upcoming.length > 0 ? upcoming : (calendar || []);
  items = items.slice(0, 5);

  if (!items.length) {
    wrap.innerHTML = `<div class="chart-empty">Belum ada jadwal kegiatan mendatang</div>`;
    return;
  }

  const DAY_NAMES = ['Min', 'Sen', 'Sel', 'Rab', 'Kam', 'Jum', 'Sab'];
  const MONTH_SHORT = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];

  wrap.innerHTML = items.map(ev => {
    const dtStr = ev.target_date || ev.event_date || '';
    let dayName = '—';
    let dateNum = '—';
    if (dtStr) {
      try {
        const d = new Date(dtStr);
        if (!isNaN(d.getTime())) {
          dayName = DAY_NAMES[d.getDay()] || 'Sen';
          dateNum = `${d.getDate()} ${MONTH_SHORT[d.getMonth()] || ''}`;
        }
      } catch {}
    }

    const title = ev.title || ev.activity_type || 'Kegiatan Operasional';
    const branch = ev.branch_name || 'Cabang Operasional';

    let pillCls = 'event-pill-insp';
    let pillText = 'Inspeksi';
    const tl = (title + ' ' + (ev.activity_type || '')).toLowerCase();
    if (tl.includes('general cleaning') || tl.includes('gc')) {
      pillCls = 'event-pill-gc';
      pillText = 'GC';
    } else if (tl.includes('deep cleaning') || tl.includes('dc')) {
      pillCls = 'event-pill-dc';
      pillText = 'Deep Cleaning';
    } else if (tl.includes('fogging')) {
      pillCls = 'event-pill-fog';
      pillText = 'Fogging';
    } else if (tl.includes('inspeksi')) {
      pillCls = 'event-pill-insp';
      pillText = 'Inspeksi';
    }

    return `
      <div class="event-item">
        <div class="event-date-box">
          <span class="event-date-day">${dayName}</span>
          <span class="event-date-num">${dateNum}</span>
        </div>
        <div class="event-info">
          <div class="event-title">${escapeHtml(title)}</div>
          <div class="event-branch">
            <svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0118 0z"/><circle cx="12" cy="10" r="3"/></svg>
            <span>${escapeHtml(branch)}</span>
          </div>
        </div>
        <div class="event-pill ${pillCls}">${pillText}</div>
      </div>
    `;
  }).join('');
}

// ── Helpers ────────────────────────────────────────────────────────────────
function hideSkel(skelId, canvasId) {
  const skel = document.getElementById(skelId);
  const cvs  = document.getElementById(canvasId);
  if (skel) { skel.style.display = 'none'; }
  if (cvs) {
    cvs.style.display = 'block';
    const wrap = cvs.parentElement;
    if (wrap) {
      const emptyMsg = wrap.querySelector('.chart-empty');
      if (emptyMsg) emptyMsg.remove();
    }
  }
}

function showEmpty(canvas, msg = 'Belum ada data') {
  if (!canvas) return;
  canvas.style.display = 'none';
  const wrap = canvas.parentElement;
  if (!wrap) return;
  const existing = wrap.querySelector('.chart-empty');
  if (!existing) {
    const div = document.createElement('div');
    div.className = 'chart-empty';
    div.textContent = msg;
    wrap.appendChild(div);
  }
}
