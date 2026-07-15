/**
 * dashboard.js — FCMS Enterprise Dashboard
 * Seluruh data dari API. Tidak ada dummy data.
 * Auto-refresh setiap 60 detik. Chart.js untuk semua grafik.
 */
import { apiFetch } from '../config.js';

// ── Chart instances registry (untuk destroy saat refresh) ────────────────────
const CHARTS = {};

function destroyCharts() {
  Object.values(CHARTS).forEach(c => { try { c.destroy(); } catch { } });
  Object.keys(CHARTS).forEach(k => delete CHARTS[k]);
}

// ── Utility ──────────────────────────────────────────────────────────────────
const safe = (v, fallback = 0) => (v !== null && v !== undefined && v !== '' && !isNaN(Number(v))) ? v : fallback;
const safeStr = (v, fallback = '—') => (v && String(v).trim() && String(v) !== '[object Object]') ? String(v).trim() : fallback;
const fmtDate = (s) => {
  if (!s) return '—';
  try { return new Date(s).toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' }); }
  catch { return s; }
};
const relTime = (s) => {
  if (!s) return '';
  const diff = Date.now() - new Date(s).getTime();
  const m = Math.floor(diff / 60000);
  if (m < 1) return 'Baru saja';
  if (m < 60) return `${m} menit lalu`;
  const h = Math.floor(m / 60);
  if (h < 24) return `${h} jam lalu`;
  return `${Math.floor(h/24)} hari lalu`;
};
const monthLabel = (ym) => {
  if (!ym) return '';
  const [y, m] = ym.split('-');
  return new Date(y, m - 1).toLocaleDateString('id-ID', { month: 'short', year: '2-digit' });
};

// Counter animation
function animateCounter(el, target, duration = 1000) {
  const start = Date.now();
  const tick = () => {
    const p = Math.min((Date.now() - start) / duration, 1);
    const ease = 1 - Math.pow(1 - p, 3);
    el.textContent = Math.round(ease * target).toLocaleString('id-ID');
    if (p < 1) requestAnimationFrame(tick);
  };
  requestAnimationFrame(tick);
}

// Trend arrow
function trendBadge(current, prev) {
  if (prev == null || prev === 0) return `<span class="kpi-trend neutral">— Baru</span>`;
  const diff = current - prev;
  const pct  = Math.abs(Math.round(diff / prev * 100));
  if (diff > 0) return `<span class="kpi-trend up">▲ ${pct}%</span>`;
  if (diff < 0) return `<span class="kpi-trend down">▼ ${pct}%</span>`;
  return `<span class="kpi-trend neutral">= Sama</span>`;
}

// Status pill colors
function statusPill(status) {
  const s = safeStr(status, 'Unknown');
  const map = {
    'Done': 'pill-success', 'Aktif': 'pill-success', 'Selesai': 'pill-success',
    'Open': 'pill-danger', 'Pending': 'pill-warning', 'In Progress': 'pill-info',
    'Tidak Aktif': 'pill-neutral', 'Resign': 'pill-neutral', 'Cut': 'pill-neutral',
  };
  const cls = map[s] || 'pill-neutral';
  return `<span class="status-pill ${cls}">${s}</span>`;
}

// Days remaining badge
function daysBadge(days) {
  const d = parseInt(days) || 0;
  if (d <= 7)  return `<span class="days-badge days-critical">${d} hari</span>`;
  if (d <= 14) return `<span class="days-badge days-warning">${d} hari</span>`;
  if (d <= 30) return `<span class="days-badge days-soon">${d} hari</span>`;
  return `<span class="days-badge days-ok">${d} hari</span>`;
}

// Activity icon per type
function activityIcon(type) {
  const icons = {
    'issue':      { icon: '⚠️', cls: 'dot-danger',  label: 'Permasalahan' },
    'contract':   { icon: '📄', cls: 'dot-info',    label: 'Kontrak' },
    'employee':   { icon: '👤', cls: 'dot-success', label: 'Karyawan' },
    'one_on_one': { icon: '🤝', cls: 'dot-purple',  label: 'One on One' },
    'training':   { icon: '🎓', cls: 'dot-primary', label: 'Training' },
    'supply':     { icon: '📦', cls: 'dot-warning', label: 'Permintaan Barang' },
  };
  return icons[type] || { icon: '📌', cls: 'dot-neutral', label: type };
}

// ── Chart defaults ───────────────────────────────────────────────────────────
const CHART_DEFAULTS = {
  animation: { duration: 800, easing: 'easeOutQuart' },
  plugins: { legend: { labels: { font: { family: 'Inter', size: 11 }, color: '#64748B' } } },
  scales: {
    x: { grid: { color: '#F1F5F9' }, ticks: { font: { family: 'Inter', size: 10 }, color: '#94A3B8' } },
    y: { grid: { color: '#F1F5F9' }, ticks: { font: { family: 'Inter', size: 10 }, color: '#94A3B8' }, beginAtZero: true },
  },
};

// ── Skeleton HTML ────────────────────────────────────────────────────────────
function skeletonKPI() {
  return Array(5).fill(0).map(() => `
    <div class="kpi-card">
      <div class="skeleton skeleton-circle" style="width:48px;height:48px;border-radius:12px;margin-bottom:12px"></div>
      <div class="skeleton skeleton-text" style="width:60%;height:28px;margin-bottom:8px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:14px"></div>
    </div>
  `).join('');
}
function skeletonMini() {
  return Array(7).fill(0).map(() => `
    <div class="mini-stat">
      <div class="skeleton skeleton-circle" style="width:36px;height:36px"></div>
      <div style="flex:1">
        <div class="skeleton skeleton-text" style="width:50%;height:20px;margin-bottom:4px"></div>
        <div class="skeleton skeleton-text" style="width:80%;height:12px"></div>
      </div>
    </div>
  `).join('');
}

// ── Main Render ───────────────────────────────────────────────────────────────
export async function renderDashboard(container) {
  destroyCharts();

  container.innerHTML = `
    <div class="dashboard-wrap" id="dash-wrap">

      <!-- KPI Row skeleton -->
      <div class="section-header">
        <h2 class="section-title">📊 Ringkasan Operasional</h2>
        <div class="dash-refresh-info">
          <span id="dash-last-updated" class="dash-last-updated"></span>
          <button class="btn btn-ghost btn-sm" id="btn-refresh-dash">
            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0114.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0020.49 15"/></svg>
            Refresh
          </button>
        </div>
      </div>

      <!-- KPI Cards -->
      <div class="kpi-row" id="kpi-row">${skeletonKPI()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${skeletonMini()}</div>

      <!-- Charts Row 1 -->
      <div class="charts-row">
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">Permasalahan per Kategori</div>
              <div class="chart-card-subtitle">Distribusi semua permasalahan</div>
            </div>
          </div>
          <div class="chart-canvas-wrap" style="height:260px;display:flex;align-items:center;justify-content:center">
            <canvas id="chart-donut"></canvas>
          </div>
        </div>
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">Trend Permasalahan 12 Bulan</div>
              <div class="chart-card-subtitle">Open vs Closed per bulan</div>
            </div>
          </div>
          <div class="chart-canvas-wrap" style="height:260px">
            <canvas id="chart-trend-line"></canvas>
          </div>
        </div>
      </div>

      <!-- Charts Row 2 -->
      <div class="charts-row">
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">Rata-rata Skor Inspeksi per Cabang</div>
              <div class="chart-card-subtitle">6 bulan terakhir (FC vs SPV)</div>
            </div>
          </div>
          <div class="chart-canvas-wrap" style="height:280px">
            <canvas id="chart-inspection-bar"></canvas>
          </div>
        </div>
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">Kontrak Berakhir 6 Bulan ke Depan</div>
              <div class="chart-card-subtitle">Jumlah kontrak per bulan</div>
            </div>
          </div>
          <div class="chart-canvas-wrap" style="height:280px">
            <canvas id="chart-contract-bar"></canvas>
          </div>
        </div>
      </div>

      <!-- Tables Row -->
      <div class="tables-row">
        <!-- Kontrak Akan Habis -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">⏰ Kontrak Akan Habis</div>
              <div class="chart-card-subtitle">30 hari ke depan</div>
            </div>
            <a href="#/contracts" class="btn btn-ghost btn-sm">Lihat Semua →</a>
          </div>
          <div id="table-contracts" class="dash-table-wrap">
            <div class="skeleton skeleton-text" style="height:40px;margin-bottom:8px"></div>
            <div class="skeleton skeleton-text" style="height:40px;margin-bottom:8px"></div>
            <div class="skeleton skeleton-text" style="height:40px"></div>
          </div>
        </div>

        <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">⚠️ Permasalahan Terbaru</div>
              <div class="chart-card-subtitle">Open dan In Progress</div>
            </div>
            <a href="#/issues" class="btn btn-ghost btn-sm">Lihat Semua →</a>
          </div>
          <div id="table-issues" class="dash-table-wrap">
            <div class="skeleton skeleton-text" style="height:40px;margin-bottom:8px"></div>
            <div class="skeleton skeleton-text" style="height:40px;margin-bottom:8px"></div>
            <div class="skeleton skeleton-text" style="height:40px"></div>
          </div>
        </div>
      </div>

      <!-- Activity Log -->
      <div class="chart-card" style="margin-top:0">
        <div class="chart-card-header">
          <div>
            <div class="chart-card-title">🕐 Aktivitas Terbaru</div>
            <div class="chart-card-subtitle">Update data terkini dari semua modul</div>
          </div>
        </div>
        <div id="activity-log" class="activity-log">
          ${Array(6).fill(0).map(() => `
            <div class="activity-item">
              <div class="skeleton skeleton-circle" style="width:32px;height:32px;flex-shrink:0"></div>
              <div style="flex:1">
                <div class="skeleton skeleton-text" style="width:60%;height:14px;margin-bottom:4px"></div>
                <div class="skeleton skeleton-text" style="width:40%;height:11px"></div>
              </div>
            </div>
          `).join('')}
        </div>
      </div>

    </div>
  `;

  // Load all data
  await loadDashboardData(container);

  // Auto-refresh setiap 60 detik
  container._dashRefresh = setInterval(() => {
    if (document.getElementById('dash-wrap')) loadDashboardData(container);
    else clearInterval(container._dashRefresh);
  }, 60000);

  // Manual refresh button
  document.getElementById('btn-refresh-dash')?.addEventListener('click', () => loadDashboardData(container));
}

// ── Load All Data ─────────────────────────────────────────────────────────────
async function loadDashboardData(container) {
  try {
    const [kpiRes, trendRes, contractChartRes, issuesSumRes, inspBarRes, expRes, issuesRes, actLogRes] = await Promise.all([
      apiFetch('/api/dashboard/kpi'),
      apiFetch('/api/dashboard/issues-trend'),
      apiFetch('/api/dashboard/contracts-chart'),
      apiFetch('/api/dashboard/issues-summary'),
      apiFetch('/api/dashboard/inspection-bar'),
      apiFetch('/api/dashboard/contracts-expiring'),
      apiFetch('/api/dashboard/stats'),
      apiFetch('/api/dashboard/activity-log'),
    ]);

    const kpi         = kpiRes.ok        ? (kpiRes.data?.data || kpiRes.data || {}) : {};
    const trend       = trendRes.ok      ? (trendRes.data?.data || {}) : {};
    const contractChart = contractChartRes.ok ? (contractChartRes.data?.data || {}) : {};
    const issuesSum   = issuesSumRes.ok  ? (issuesSumRes.data?.data || {}) : {};
    const inspBar     = inspBarRes.ok    ? (inspBarRes.data?.data || {}) : {};
    const expContracts= expRes.ok        ? (Array.isArray(expRes.data?.data) ? expRes.data.data : (expRes.data || [])) : [];
    const stats       = issuesRes.ok     ? (issuesRes.data?.data?.recent_issues || []) : [];
    const actLog      = actLogRes.ok     ? (Array.isArray(actLogRes.data?.data) ? actLogRes.data.data : []) : [];

    renderKPI(kpi);
    renderMiniStats(kpi);
    renderDonutChart(issuesSum.by_category || []);
    renderTrendChart(trend);
    renderInspectionBar(inspBar);
    renderContractBar(contractChart);
    renderContractsTable(expContracts);
    renderIssuesTable(stats);
    renderActivityLog(actLog);

    // Update last-updated timestamp
    const el = document.getElementById('dash-last-updated');
    if (el) el.textContent = `Update: ${new Date().toLocaleTimeString('id-ID')}`;

  } catch (err) {
    console.error('Dashboard load error:', err);
  }
}

// ── KPI Cards ─────────────────────────────────────────────────────────────────
function renderKPI(kpi) {
  const cards = [
    {
      icon: '👥', label: 'Karyawan Aktif',
      value: safe(kpi.employees?.current, 0),
      prev:  kpi.employees?.prev,
      color: 'kpi-blue', subtitle: 'Total karyawan aktif',
      href: '#/employees',
    },
    {
      icon: '📄', label: 'Kontrak Aktif',
      value: safe(kpi.contracts?.current, 0),
      prev:  kpi.contracts?.prev,
      color: 'kpi-green', subtitle: 'Kontrak yang masih berjalan',
      href: '#/contracts',
    },
    {
      icon: '⏰', label: 'Kontrak Habis 30 Hari',
      value: safe(kpi.expiring30?.current, 0),
      prev:  null,
      color: safe(kpi.expiring30?.current, 0) > 0 ? 'kpi-amber' : 'kpi-green',
      subtitle: safe(kpi.expiring30?.current, 0) > 0 ? '⚠️ Segera diperbarui' : '✅ Aman',
      href: '#/contracts', warn: safe(kpi.expiring30?.current, 0) > 0,
    },
    {
      icon: '⚠️', label: 'Permasalahan Open',
      value: safe(kpi.issues?.current, 0),
      prev:  kpi.issues?.prev,
      color: safe(kpi.issues?.current, 0) > 0 ? 'kpi-red' : 'kpi-green',
      subtitle: 'Belum diselesaikan',
      href: '#/issues',
    },
    {
      icon: '🤝', label: 'One on One Pending',
      value: safe(kpi.one_on_one?.current, 0),
      prev:  kpi.one_on_one?.prev,
      color: 'kpi-purple', subtitle: 'Menunggu tindak lanjut',
      href: '#/one-on-one',
    },
  ];

  const row = document.getElementById('kpi-row');
  if (!row) return;
  row.innerHTML = cards.map(c => `
    <a href="${c.href}" class="kpi-card ${c.color} ${c.warn ? 'kpi-warn' : ''}" style="text-decoration:none">
      <div class="kpi-card-top">
        <div class="kpi-icon-wrap"><span class="kpi-icon-emoji">${c.icon}</span></div>
        ${c.prev != null ? trendBadge(c.value, c.prev) : ''}
      </div>
      <div class="kpi-value" data-target="${c.value}">0</div>
      <div class="kpi-label">${c.label}</div>
      <div class="kpi-subtitle">${c.subtitle}</div>
    </a>
  `).join('');

  // Animate counters
  row.querySelectorAll('.kpi-value').forEach(el => {
    const target = parseInt(el.dataset.target) || 0;
    animateCounter(el, target);
  });
}

// ── Mini Stats ─────────────────────────────────────────────────────────────────
function renderMiniStats(kpi) {
  const stats = [
    { icon: '📅', label: 'Jadwal Pending',    value: safe(kpi.schedule?.current, 0),       href: '#/schedule',          color: 'mini-blue' },
    { icon: '🎓', label: 'Training Bulan Ini', value: safe(kpi.training_month?.current, 0),  href: '#/training',          color: 'mini-indigo' },
    { icon: '📦', label: 'Permintaan Barang',  value: safe(kpi.supply?.current, 0),          href: '#/reports/supply',    color: 'mini-orange' },
    { icon: '🔍', label: 'Inspeksi Bulan Ini', value: safe(kpi.inspection_month?.current, 0),href: '#/reports/inspection', color: 'mini-teal' },
    { icon: '🧹', label: 'GC/DC Bulan Ini',   value: safe(kpi.cleaning_month?.current, 0),  href: '#/reports/cleaning',  color: 'mini-green' },
    { icon: '🦟', label: 'Fogging Bulan Ini',  value: safe(kpi.fogging_month?.current, 0),   href: '#/reports/fogging',   color: 'mini-purple' },
    { icon: '🏢', label: 'Total Cabang',       value: safe(kpi.branches?.current, 0),        href: '#/branches',          color: 'mini-gray' },
  ];

  const row = document.getElementById('mini-stats-row');
  if (!row) return;
  row.innerHTML = stats.map(s => `
    <a href="${s.href}" class="mini-stat ${s.color}" style="text-decoration:none">
      <div class="mini-stat-icon">${s.icon}</div>
      <div class="mini-stat-body">
        <div class="mini-stat-value" data-target="${s.value}">0</div>
        <div class="mini-stat-label">${s.label}</div>
      </div>
    </a>
  `).join('');

  row.querySelectorAll('.mini-stat-value').forEach(el => {
    animateCounter(el, parseInt(el.dataset.target) || 0, 800);
  });
}

// ── Donut Chart: Issues by Category ───────────────────────────────────────────
function renderDonutChart(categories) {
  const canvas = document.getElementById('chart-donut');
  if (!canvas) return;
  if (CHARTS.donut) { CHARTS.donut.destroy(); }

  const labels = categories.map(c => safeStr(c.category, 'Lainnya'));
  const data   = categories.map(c => safe(c.count, 0));

  if (data.length === 0) {
    canvas.parentElement.innerHTML = '<div class="chart-empty">Tidak ada data</div>';
    return;
  }

  const COLORS = ['#2563EB','#10B981','#F59E0B','#EF4444','#8B5CF6','#0EA5E9','#F97316','#14B8A6'];
  CHARTS.donut = new Chart(canvas, {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{ data, backgroundColor: COLORS.slice(0, labels.length), borderWidth: 2, borderColor: '#fff', hoverOffset: 6 }],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      cutout: '65%',
      plugins: {
        legend: { position: 'right', labels: { font: { family: 'Inter', size: 11 }, padding: 12, color: '#475569', usePointStyle: true } },
        tooltip: { callbacks: { label: ctx => ` ${ctx.label}: ${ctx.parsed} kasus` } },
      },
      animation: { duration: 900 },
    },
  });
}

// ── Line Chart: Issues Trend 12 Months ────────────────────────────────────────
function renderTrendChart(trend) {
  const canvas = document.getElementById('chart-trend-line');
  if (!canvas) return;
  if (CHARTS.trend) { CHARTS.trend.destroy(); }

  const labels = (trend.labels || []).map(monthLabel);
  const open   = trend.open   || Array(12).fill(0);
  const closed = trend.closed || Array(12).fill(0);

  CHARTS.trend = new Chart(canvas, {
    type: 'line',
    data: {
      labels,
      datasets: [
        {
          label: 'Open', data: open, borderColor: '#EF4444', backgroundColor: 'rgba(239,68,68,.1)',
          fill: true, tension: 0.4, pointRadius: 4, pointBackgroundColor: '#EF4444', borderWidth: 2,
        },
        {
          label: 'Closed', data: closed, borderColor: '#10B981', backgroundColor: 'rgba(16,185,129,.1)',
          fill: true, tension: 0.4, pointRadius: 4, pointBackgroundColor: '#10B981', borderWidth: 2,
        },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'top', labels: { font: { family: 'Inter', size: 11 }, color: '#475569', usePointStyle: true } } },
      scales: {
        x: { grid: { color: '#F1F5F9' }, ticks: { font: { family: 'Inter', size: 10 }, color: '#94A3B8' } },
        y: { grid: { color: '#F1F5F9' }, ticks: { font: { family: 'Inter', size: 10 }, color: '#94A3B8' }, beginAtZero: true },
      },
      interaction: { mode: 'index', intersect: false },
      animation: { duration: 900 },
    },
  });
}

// ── Bar Chart: Inspection by Branch ──────────────────────────────────────────
function renderInspectionBar(inspBar) {
  const canvas = document.getElementById('chart-inspection-bar');
  if (!canvas) return;
  if (CHARTS.inspBar) { CHARTS.inspBar.destroy(); }

  const labels = inspBar.labels || [];
  const fc     = inspBar.fc    || [];
  const spv    = inspBar.spv   || [];

  if (labels.length === 0) {
    canvas.parentElement.innerHTML = '<div class="chart-empty">Tidak ada data inspeksi</div>';
    return;
  }

  CHARTS.inspBar = new Chart(canvas, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        { label: 'Skor FC', data: fc,  backgroundColor: 'rgba(37,99,235,.75)',  borderRadius: 4, borderSkipped: false },
        { label: 'Skor SPV', data: spv, backgroundColor: 'rgba(16,185,129,.75)', borderRadius: 4, borderSkipped: false },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'top', labels: { font: { family: 'Inter', size: 11 }, color: '#475569', usePointStyle: true } } },
      scales: {
        x: { grid: { display: false }, ticks: { font: { family: 'Inter', size: 9 }, color: '#94A3B8', maxRotation: 45 } },
        y: { grid: { color: '#F1F5F9' }, ticks: { font: { family: 'Inter', size: 10 }, color: '#94A3B8' }, min: 0, max: 100 },
      },
      animation: { duration: 900 },
    },
  });
}

// ── Bar Chart: Contracts Expiring ─────────────────────────────────────────────
function renderContractBar(contractChart) {
  const canvas = document.getElementById('chart-contract-bar');
  if (!canvas) return;
  if (CHARTS.contractBar) { CHARTS.contractBar.destroy(); }

  const labels = (contractChart.labels || []).map(monthLabel);
  const counts = contractChart.counts || [];

  CHARTS.contractBar = new Chart(canvas, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Kontrak Berakhir',
        data: counts,
        backgroundColor: counts.map(v => v > 5 ? 'rgba(239,68,68,.75)' : v > 2 ? 'rgba(245,158,11,.75)' : 'rgba(37,99,235,.6)'),
        borderRadius: 6,
        borderSkipped: false,
      }],
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { grid: { display: false }, ticks: { font: { family: 'Inter', size: 11 }, color: '#64748B' } },
        y: { grid: { color: '#F1F5F9' }, ticks: { font: { family: 'Inter', size: 10 }, color: '#94A3B8' }, beginAtZero: true, precision: 0 },
      },
      animation: { duration: 900 },
    },
  });
}

// ── Contracts Table ────────────────────────────────────────────────────────────
function renderContractsTable(rows) {
  const wrap = document.getElementById('table-contracts');
  if (!wrap) return;

  const expiring = (rows || []).filter(r => {
    const days = parseInt(r.days_remaining) || 999;
    return days <= 30;
  }).slice(0, 10);

  if (expiring.length === 0) {
    wrap.innerHTML = '<div class="chart-empty">✅ Tidak ada kontrak yang habis dalam 30 hari</div>';
    return;
  }

  wrap.innerHTML = `
    <table class="dash-table">
      <thead>
        <tr>
          <th>No</th>
          <th>Nama Karyawan</th>
          <th>Cabang</th>
          <th>Berakhir</th>
          <th>Sisa</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody>
        ${expiring.map((r, i) => `
          <tr>
            <td class="td-num">${i + 1}</td>
            <td><strong>${safeStr(r.emp_name || r.employee_name)}</strong></td>
            <td class="td-branch">${safeStr(r.branch_name)}</td>
            <td>${fmtDate(r.end_date)}</td>
            <td>${daysBadge(r.days_remaining)}</td>
            <td>${statusPill(r.status)}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;
}

// ── Issues Table ──────────────────────────────────────────────────────────────
function renderIssuesTable(rows) {
  const wrap = document.getElementById('table-issues');
  if (!wrap) return;

  const issues = (rows || []).slice(0, 8);

  if (issues.length === 0) {
    wrap.innerHTML = '<div class="chart-empty">✅ Tidak ada permasalahan terbuka</div>';
    return;
  }

  wrap.innerHTML = `
    <table class="dash-table">
      <thead>
        <tr>
          <th>Tanggal</th>
          <th>Keluhan</th>
          <th>Cabang</th>
          <th>Kategori</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody>
        ${issues.map(r => `
          <tr>
            <td style="white-space:nowrap">${fmtDate(r.report_date)}</td>
            <td class="td-complaint">${safeStr(r.complaint)}</td>
            <td class="td-branch">${safeStr(r.branch_name)}</td>
            <td><span class="category-tag">${safeStr(r.category)}</span></td>
            <td>${statusPill(r.status)}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;
}

// ── Activity Log ──────────────────────────────────────────────────────────────
function renderActivityLog(rows) {
  const wrap = document.getElementById('activity-log');
  if (!wrap) return;

  if (!rows || rows.length === 0) {
    wrap.innerHTML = '<div class="chart-empty">Belum ada aktivitas</div>';
    return;
  }

  wrap.innerHTML = `<div class="activity-list">
    ${rows.slice(0, 15).map(r => {
      const info = activityIcon(r.type);
      const label = safeStr(r.label, '(tanpa nama)');
      const branch = r.branch ? ` • ${r.branch}` : '';
      return `
        <div class="activity-item">
          <div class="activity-dot ${info.cls}">${info.icon}</div>
          <div class="activity-body">
            <div class="activity-text">
              <strong>${info.label}</strong> — ${label}${branch}
            </div>
            <div class="activity-time">${relTime(r.created_at)}</div>
          </div>
        </div>
      `;
    }).join('')}
  </div>`;
}
