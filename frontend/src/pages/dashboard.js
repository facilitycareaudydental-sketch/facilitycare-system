import { apiFetch } from '../config.js';
import { statusBadge, daysRemainingBadge, activityTypeBadge } from '../components/badges.js';

export async function renderDashboard(container) {
  container.innerHTML = `
    <div class="page-header">
      <h1 class="page-title">Dashboard</h1>
      <div class="page-actions">
        <span class="text-muted" id="dash-date"></span>
      </div>
    </div>
    <div id="dash-content">
      <div class="loading-spinner"><div class="spinner"></div></div>
    </div>
  `;

  document.getElementById('dash-date').textContent = new Date().toLocaleDateString('id-ID', {
    weekday: 'long', year: 'numeric', month: 'long', day: 'numeric'
  });

  const [statsRes, issuesSummaryRes] = await Promise.all([
    apiFetch('/api/dashboard/stats'),
    apiFetch('/api/dashboard/issues-summary'),
  ]);

  const stats = statsRes.data?.data || {};
  const issuesSummary = issuesSummaryRes.data?.data || {};

  document.getElementById('dash-content').innerHTML = `
    <!-- Stat cards -->
    <div class="stats-grid">
      <div class="stat-card stat-card-blue">
        <div class="stat-icon">👥</div>
        <div class="stat-body">
          <div class="stat-value">${stats.total_employees || 0}</div>
          <div class="stat-label">Karyawan Aktif</div>
        </div>
      </div>
      <div class="stat-card stat-card-green">
        <div class="stat-icon">📋</div>
        <div class="stat-body">
          <div class="stat-value">${stats.active_contracts || 0}</div>
          <div class="stat-label">Kontrak Aktif</div>
        </div>
      </div>
      <div class="stat-card stat-card-yellow">
        <div class="stat-icon">⏰</div>
        <div class="stat-body">
          <div class="stat-value">${stats.expiring_contracts || 0}</div>
          <div class="stat-label">Kontrak Habis 30 Hari</div>
        </div>
      </div>
      <div class="stat-card stat-card-red">
        <div class="stat-icon">⚠️</div>
        <div class="stat-body">
          <div class="stat-value">${stats.open_issues || 0}</div>
          <div class="stat-label">Permasalahan Open</div>
        </div>
      </div>
      <div class="stat-card stat-card-purple">
        <div class="stat-icon">🤝</div>
        <div class="stat-body">
          <div class="stat-value">${stats.open_one_on_one || 0}</div>
          <div class="stat-label">One on One Pending</div>
        </div>
      </div>
      <div class="stat-card stat-card-teal">
        <div class="stat-icon">🗓️</div>
        <div class="stat-body">
          <div class="stat-value">${stats.pending_schedule || 0}</div>
          <div class="stat-label">Jadwal Pending</div>
        </div>
      </div>
      <div class="stat-card stat-card-orange">
        <div class="stat-icon">📦</div>
        <div class="stat-body">
          <div class="stat-value">${stats.pending_supply_requests || 0}</div>
          <div class="stat-label">Permintaan Barang</div>
        </div>
      </div>
      <div class="stat-card stat-card-gray">
        <div class="stat-icon">🏢</div>
        <div class="stat-body">
          <div class="stat-value">${stats.total_branches || 0}</div>
          <div class="stat-label">Total Cabang</div>
        </div>
      </div>
    </div>

    <!-- Charts row -->
    <div class="dashboard-grid">
      <div class="card">
        <div class="card-header">
          <h3 class="card-title">Permasalahan per Kategori</h3>
        </div>
        <div class="card-body">
          <div id="chart-issues-category"></div>
        </div>
      </div>
      <div class="card">
        <div class="card-header">
          <h3 class="card-title">Tren Permasalahan (12 Bulan)</h3>
        </div>
        <div class="card-body">
          <div id="chart-issues-trend"></div>
        </div>
      </div>
    </div>

    <!-- Tables row -->
    <div class="dashboard-grid">
      <div class="card">
        <div class="card-header">
          <h3 class="card-title">⏰ Kontrak Akan Habis</h3>
          <a href="#/contracts" class="btn btn-sm btn-ghost">Lihat Semua</a>
        </div>
        <div class="card-body p-0">
          <div class="table-wrapper">
            <table class="data-table">
              <thead><tr><th>Nama</th><th>Cabang</th><th>Sisa</th></tr></thead>
              <tbody id="expiring-tbody"></tbody>
            </table>
          </div>
        </div>
      </div>
      <div class="card">
        <div class="card-header">
          <h3 class="card-title">⚠️ Permasalahan Terbaru</h3>
          <a href="#/issues" class="btn btn-sm btn-ghost">Lihat Semua</a>
        </div>
        <div class="card-body p-0">
          <div class="table-wrapper">
            <table class="data-table">
              <thead><tr><th>Tanggal</th><th>Cabang</th><th>Status</th></tr></thead>
              <tbody id="issues-tbody"></tbody>
            </table>
          </div>
        </div>
      </div>
    </div>

    <!-- Upcoming schedule -->
    <div class="card">
      <div class="card-header">
        <h3 class="card-title">🗓️ Jadwal Mendatang</h3>
        <a href="#/schedule" class="btn btn-sm btn-ghost">Lihat Semua</a>
      </div>
      <div class="card-body p-0">
        <div class="table-wrapper">
          <table class="data-table">
            <thead><tr><th>Cabang</th><th>Kegiatan</th><th>Periode</th><th>Target</th><th>PIC</th><th>Status</th></tr></thead>
            <tbody id="schedule-tbody"></tbody>
          </table>
        </div>
      </div>
    </div>
  `;

  // Populate expiring contracts
  const expiringTbody = document.getElementById('expiring-tbody');
  if (stats.expiring_contracts === 0) {
    expiringTbody.innerHTML = '<tr><td colspan="3" class="text-center text-muted">Tidak ada kontrak yang akan habis</td></tr>';
  }
  (statsRes.data?.data?.expiring_contracts || []).forEach(c => {
    expiringTbody.innerHTML += `
      <tr>
        <td>${c.employee_name}</td>
        <td>${c.branch_name || '-'}</td>
        <td>${daysRemainingBadge(c.days_remaining)}</td>
      </tr>`;
  });

  // Populate recent issues
  const issuesTbody = document.getElementById('issues-tbody');
  const recentIssues = statsRes.data?.data?.recent_issues || [];
  if (recentIssues.length === 0) {
    issuesTbody.innerHTML = '<tr><td colspan="3" class="text-center text-muted">Tidak ada permasalahan</td></tr>';
  }
  recentIssues.forEach(i => {
    issuesTbody.innerHTML += `
      <tr>
        <td>${i.report_date || '-'}</td>
        <td>${i.branch_name || '-'}</td>
        <td>${statusBadge(i.status)}</td>
      </tr>`;
  });

  // Populate upcoming schedule
  const schedTbody = document.getElementById('schedule-tbody');
  const upcoming = statsRes.data?.data?.upcoming_schedule || [];
  if (upcoming.length === 0) {
    schedTbody.innerHTML = '<tr><td colspan="6" class="text-center text-muted">Tidak ada jadwal mendatang</td></tr>';
  }
  upcoming.forEach(s => {
    schedTbody.innerHTML += `
      <tr>
        <td>${s.branch_name || '-'}</td>
        <td>${activityTypeBadge(s.activity_type)}</td>
        <td><span class="badge badge-info">${s.period}</span></td>
        <td>${s.target_date || '-'}</td>
        <td>${s.pic || '-'}</td>
        <td>${statusBadge(s.status)}</td>
      </tr>`;
  });

  // Simple bar chart using HTML/CSS
  renderBarChart('chart-issues-category', issuesSummary.by_category || [], 'category', 'count', 'Permasalahan');
  renderLineChart('chart-issues-trend', issuesSummary.by_month || [], 'month', 'count');
}

function renderBarChart(containerId, data, labelKey, valueKey, title) {
  const el = document.getElementById(containerId);
  if (!el || !data.length) { el && (el.innerHTML = '<p class="text-center text-muted">Tidak ada data</p>'); return; }

  const max = Math.max(...data.map(d => d[valueKey]));
  el.innerHTML = `
    <div class="bar-chart">
      ${data.slice(0, 8).map(d => `
        <div class="bar-item">
          <div class="bar-label">${d[labelKey]}</div>
          <div class="bar-track">
            <div class="bar-fill" style="width:${max > 0 ? (d[valueKey] / max * 100) : 0}%">
              <span class="bar-value">${d[valueKey]}</span>
            </div>
          </div>
        </div>
      `).join('')}
    </div>
  `;
}

function renderLineChart(containerId, data, labelKey, valueKey) {
  const el = document.getElementById(containerId);
  if (!el || !data.length) { el && (el.innerHTML = '<p class="text-center text-muted">Tidak ada data</p>'); return; }

  const max = Math.max(...data.map(d => d[valueKey]), 1);
  const width = 400, height = 120, pad = 30;
  const points = data.map((d, i) => {
    const x = pad + (i / (data.length - 1 || 1)) * (width - pad * 2);
    const y = height - pad - (d[valueKey] / max) * (height - pad * 2);
    return `${x},${y}`;
  }).join(' ');

  el.innerHTML = `
    <div class="line-chart-wrap">
      <svg viewBox="0 0 ${width} ${height}" class="line-chart-svg" role="img" aria-label="Tren permasalahan">
        <polyline points="${points}" class="line-path" fill="none"/>
        ${data.map((d, i) => {
          const x = pad + (i / (data.length - 1 || 1)) * (width - pad * 2);
          const y = height - pad - (d[valueKey] / max) * (height - pad * 2);
          return `<circle cx="${x}" cy="${y}" r="4" class="line-dot"/>
                  <text x="${x}" y="${height - 5}" text-anchor="middle" class="chart-x-label">${(d[labelKey] || '').slice(5)}</text>`;
        }).join('')}
      </svg>
    </div>
  `;
}
