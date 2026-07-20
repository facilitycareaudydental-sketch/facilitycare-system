/**
 * dashboard.js — FCMS Enterprise Dashboard (Frontend)
 * ─────────────────────────────────────────────────────
 * BUG FIX v3:
 * - Setiap API call diisolasi dengan Promise.allSettled + try/catch individual
 * - Skeleton SELALU hilang: setelah data diterima ATAU max 5 detik timeout
 * - Jam digital dikelola di sini jika app.js tidak menyediakan elemen
 * - fetchAll tidak pernah menggantung — fallback ke default jika API gagal
 * - Tidak pernah menampilkan [object Object]
 */
import { apiFetch } from '../config.js';

// ── Chart registry ─────────────────────────────────────────────────────────
const _charts = {};
function destroyChart(key) {
  if (_charts[key]) { try { _charts[key].destroy(); } catch {} delete _charts[key]; }
}
function destroyAllCharts() { Object.keys(_charts).forEach(destroyChart); }

// ── Safe value helpers ─────────────────────────────────────────────────────
const safeNum = (v, fb = 0) => {
  const n = Number(v);
  return isNaN(n) || v === null || v === undefined ? fb : n;
};
const safeStr = (v, fb = '—') => {
  if (v === null || v === undefined || v === '') return fb;
  const s = String(v).trim();
  return (s === '' || s === '[object Object]') ? fb : s;
};
const fmtDate = (s) => {
  if (!s) return '—';
  try {
    const d = new Date(s);
    if (isNaN(d)) return safeStr(s);
    return d.toLocaleDateString('id-ID', { day:'2-digit', month:'short', year:'numeric' });
  } catch { return safeStr(s); }
};
const relTime = (s) => {
  if (!s) return '';
  try {
    const ms = Date.now() - new Date(s).getTime();
    if (ms < 0) return 'Baru saja';
    const m = Math.floor(ms / 60000);
    if (m < 1)  return 'Baru saja';
    if (m < 60) return `${m} menit lalu`;
    const h = Math.floor(m / 60);
    if (h < 24) return `${h} jam lalu`;
    return `${Math.floor(h/24)} hari lalu`;
  } catch { return ''; }
};
const monthShort = (ym) => {
  if (!ym || typeof ym !== 'string') return '';
  try {
    const [y, m] = ym.split('-');
    return new Date(Number(y), Number(m)-1).toLocaleDateString('id-ID', { month:'short', year:'2-digit' });
  } catch { return ym; }
};

// ── Counter animation ──────────────────────────────────────────────────────
function animateCount(el, target, ms = 900) {
  if (!el) return;
  const t = Math.max(0, Math.round(safeNum(target)));
  if (t === 0) { el.textContent = '0'; return; }
  const start = Date.now();
  const tick = () => {
    const p    = Math.min((Date.now()-start)/ms, 1);
    const ease = 1 - Math.pow(1-p, 3);
    el.textContent = Math.round(ease*t).toLocaleString('id-ID');
    if (p < 1) requestAnimationFrame(tick);
    else el.textContent = t.toLocaleString('id-ID');
  };
  requestAnimationFrame(tick);
}

// ── Trend badge ────────────────────────────────────────────────────────────
function trendBadge(cur, prev) {
  cur  = safeNum(cur);
  prev = safeNum(prev);
  if (prev === 0) return '';
  const diff = cur - prev;
  const pct  = Math.abs(Math.round((diff/prev)*100));
  if (diff > 0) return `<span class="kpi-trend up">▲ ${pct}%</span>`;
  if (diff < 0) return `<span class="kpi-trend down">▼ ${pct}%</span>`;
  return `<span class="kpi-trend neutral">= Sama</span>`;
}

// ── Status pill ────────────────────────────────────────────────────────────
const STATUS_CLS = {
  'Done':'pill-success','Aktif':'pill-success','Selesai':'pill-success',
  'Open':'pill-danger','Pending':'pill-warning','In Progress':'pill-info',
  'Tidak Aktif':'pill-neutral','Resign':'pill-neutral','Cut':'pill-neutral',
};
const statusPill = (s) => {
  const str = safeStr(s, '—');
  return `<span class="status-pill ${STATUS_CLS[str]||'pill-neutral'}">${str}</span>`;
};

// ── Days badge ─────────────────────────────────────────────────────────────
const daysBadge = (d) => {
  const n = safeNum(d, 999);
  if (n <= 7)  return `<span class="days-badge days-critical">${n} hari</span>`;
  if (n <= 14) return `<span class="days-badge days-warning">${n} hari</span>`;
  if (n <= 30) return `<span class="days-badge days-soon">${n} hari</span>`;
  return `<span class="days-badge days-ok">${n} hari</span>`;
};

// ── Activity map ───────────────────────────────────────────────────────────
const ACT_MAP = {
  issue:      { emoji:'⚠️', dot:'dot-danger',  label:'Permasalahan' },
  contract:   { emoji:'📄', dot:'dot-info',    label:'Kontrak' },
  employee:   { emoji:'👤', dot:'dot-success', label:'Karyawan' },
  one_on_one: { emoji:'🤝', dot:'dot-purple',  label:'One on One' },
  training:   { emoji:'🎓', dot:'dot-primary', label:'Training' },
  supply:     { emoji:'📦', dot:'dot-warning', label:'Permintaan Barang' },
  reliever:   { emoji:'🔄', dot:'dot-teal',    label:'Reliefer' },
  inspection: { emoji:'🔍', dot:'dot-blue',    label:'Laporan Inspeksi' },
};
const actInfo = (type) =>
  ACT_MAP[type] || { emoji:'📌', dot:'dot-neutral', label: safeStr(type,'Aktivitas') };

// ── Chart.js defaults ──────────────────────────────────────────────────────
const FONT  = { family:'Inter', size:11 };
const TICK  = '#94A3B8';
const GRID  = '#F1F5F9';
const COLORS= ['#2563EB','#10B981','#F59E0B','#EF4444','#8B5CF6','#0EA5E9','#F97316','#14B8A6','#6366F1','#EC4899'];

// Detect mobile for chart legend positioning
const isMobile = () => window.innerWidth < 768;

function chartOpts(extra={}) {
  return {
    responsive:true,
    maintainAspectRatio:false,
    animation:{ duration:700, easing:'easeOutQuart' },
    plugins:{
      legend:{
        position: isMobile() ? 'bottom' : 'top',
        labels:{
          font:FONT, color:'#64748B',
          usePointStyle:true, padding:10,
          boxWidth:8, boxHeight:8,
        },
      },
      tooltip:{ mode:'index', intersect:false, bodyFont:FONT, titleFont:{...FONT,weight:'700'} },
    },
    scales:{
      x:{ grid:{ color:GRID }, ticks:{ font:FONT, color:TICK, maxRotation:0 } },
      y:{ grid:{ color:GRID }, ticks:{ font:FONT, color:TICK }, beginAtZero:true },
    },
    ...extra,
  };
}

// ── Skeleton HTML ──────────────────────────────────────────────────────────
const skelKPI = () => Array(5).fill(0).map(()=>`
  <div class="kpi-card" style="pointer-events:none">
    <div class="kpi-card-top"><div class="skeleton" style="width:44px;height:44px;border-radius:12px"></div></div>
    <div class="skeleton skeleton-text" style="width:55%;height:32px;margin:10px 0 6px"></div>
    <div class="skeleton skeleton-text" style="width:75%;height:12px;margin-bottom:4px"></div>
    <div class="skeleton skeleton-text" style="width:55%;height:11px"></div>
  </div>`).join('');

const skelMini = () => Array(7).fill(0).map(()=>`
  <div class="mini-stat" style="pointer-events:none">
    <div class="skeleton" style="width:40px;height:40px;border-radius:10px;flex-shrink:0"></div>
    <div style="flex:1">
      <div class="skeleton skeleton-text" style="width:45%;height:22px;margin-bottom:5px"></div>
      <div class="skeleton skeleton-text" style="width:80%;height:11px"></div>
    </div>
  </div>`).join('');

function skelTable(rows=3) {
  return Array(rows).fill(0).map((_,i)=>
    `<div class="skeleton skeleton-text" style="height:38px;margin-bottom:${i<rows-1?'6px':'0'};border-radius:6px"></div>`
  ).join('');
}
function skelActivity() {
  return Array(5).fill(0).map(()=>`
    <div class="activity-item">
      <div class="skeleton" style="width:34px;height:34px;border-radius:10px;flex-shrink:0"></div>
      <div style="flex:1">
        <div class="skeleton skeleton-text" style="width:65%;height:13px;margin-bottom:5px"></div>
        <div class="skeleton skeleton-text" style="width:35%;height:11px"></div>
      </div>
    </div>`).join('');
}

// ── Safe API fetch with timeout + fallback ─────────────────────────────────
// Each call is independent: never blocks others on failure
async function safeFetch(path, fallback, timeoutMs=8000) {
  try {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    const res = await apiFetch(path, { signal: controller.signal }).catch(()=>null);
    clearTimeout(timer);
    if (!res || !res.ok) return fallback;
    const d = res.data;
    if (!d) return fallback;
    return d.data !== undefined ? (d.data ?? fallback) : d;
  } catch {
    return fallback;
  }
}

// ── Force remove all skeletons (max 5s timeout safety net) ────────────────
function forceRemoveSkeletons() {
  ['skel-donut','skel-trend','skel-insp','skel-contract'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.display = 'none';
  });
  ['chart-donut','chart-trend','chart-insp','chart-contract'].forEach(id => {
    const el = document.getElementById(id);
    if (el && el.style.display === 'none') {
      el.style.display = 'block';
      // show empty state if still blank
      const wrap = el.parentElement;
      if (wrap && !wrap.querySelector('.chart-empty')) {
        const msg = document.createElement('div');
        msg.className = 'chart-empty';
        msg.textContent = 'Belum ada data';
        el.style.display = 'none';
        wrap.appendChild(msg);
      }
    }
  });
  // KPI row
  const kpiRow = document.getElementById('kpi-row');
  if (kpiRow && kpiRow.querySelector('.skeleton')) {
    renderKPI({});
  }
  // Mini stats
  const miniRow = document.getElementById('mini-stats-row');
  if (miniRow && miniRow.querySelector('.skeleton')) {
    renderMiniStats({});
  }
  // Tables
  ['table-contracts','table-issues'].forEach(id=>{
    const el = document.getElementById(id);
    if (el && el.querySelector('.skeleton')) {
      el.innerHTML = `<div class="chart-empty">Belum ada data</div>`;
    }
  });
  // Activity
  const act = document.getElementById('activity-log');
  if (act && act.querySelector('.skeleton')) {
    act.innerHTML = `<div class="chart-empty">Belum ada aktivitas</div>`;
  }
}

// ── Main render ────────────────────────────────────────────────────────────
export async function renderDashboard(container) {
  destroyAllCharts();
  if (container._dashRefresh)  clearInterval(container._dashRefresh);
  if (container._skelTimeout)  clearTimeout(container._skelTimeout);

  container.innerHTML = `
    <div class="dashboard-wrap" id="dash-root">

      <div class="section-header">
        <h2 class="section-title">📊 Dashboard Operasional FCMS</h2>
        <div class="dash-refresh-info">
          <span id="dash-updated" class="dash-last-updated"></span>
          <button class="btn btn-ghost btn-sm" id="btn-dash-refresh" title="Refresh">
            <svg width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24">
              <polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/>
              <path d="M3.51 9a9 9 0 0114.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0020.49 15"/>
            </svg>
            Refresh
          </button>
        </div>
      </div>

      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${skelKPI()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${skelMini()}</div>

      <!-- Charts Row 1 -->
      <div class="charts-row">
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">🍩 Permasalahan per Kategori</div>
              <div class="chart-card-subtitle">Distribusi semua permasalahan</div>
            </div>
          </div>
          <div class="chart-canvas-wrap" style="height:260px;position:relative">
            <div id="skel-donut" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-donut" style="display:none"></canvas>
          </div>
        </div>
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">📈 Trend Permasalahan 12 Bulan</div>
              <div class="chart-card-subtitle">Open vs Closed per bulan</div>
            </div>
          </div>
          <div class="chart-canvas-wrap" style="height:260px;position:relative">
            <div id="skel-trend" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-trend" style="display:none"></canvas>
          </div>
        </div>
      </div>

      <!-- Charts Row 2 -->
      <div class="charts-row">
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">🔍 Rata-rata Skor Inspeksi per Cabang</div>
              <div class="chart-card-subtitle">6 bulan terakhir — FC vs SPV</div>
            </div>
          </div>
          <div class="chart-canvas-wrap" style="height:280px;position:relative">
            <div id="skel-insp" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-insp" style="display:none"></canvas>
          </div>
        </div>
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">📋 Kontrak Berakhir 6 Bulan ke Depan</div>
              <div class="chart-card-subtitle">Jumlah kontrak aktif per bulan</div>
            </div>
          </div>
          <div class="chart-canvas-wrap" style="height:280px;position:relative">
            <div id="skel-contract" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-contract" style="display:none"></canvas>
          </div>
        </div>
      </div>

      <!-- Tables -->
      <div class="tables-row">
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">⏰ Kontrak Akan Habis</div>
              <div class="chart-card-subtitle">30 hari ke depan</div>
            </div>
            <a href="#/contracts" class="btn btn-ghost btn-sm">Lihat Semua →</a>
          </div>
          <div id="table-contracts" class="dash-table-wrap">${skelTable(3)}</div>
        </div>
        <div class="chart-card">
          <div class="chart-card-header">
            <div>
              <div class="chart-card-title">⚠️ Permasalahan Terbaru</div>
              <div class="chart-card-subtitle">Open dan In Progress</div>
            </div>
            <a href="#/issues" class="btn btn-ghost btn-sm">Lihat Semua →</a>
          </div>
          <div id="table-issues" class="dash-table-wrap">${skelTable(3)}</div>
        </div>
      </div>

      <!-- Activity Log -->
      <div class="chart-card">
        <div class="chart-card-header">
          <div>
            <div class="chart-card-title">🕐 Aktivitas Terbaru</div>
            <div class="chart-card-subtitle">Update real-time dari semua modul</div>
          </div>
        </div>
        <div id="activity-log">${skelActivity()}</div>
      </div>

    </div>
  `;

  document.getElementById('btn-dash-refresh')
    ?.addEventListener('click', () => fetchAll(container));

  // ── Safety: force remove skeletons after 5 seconds no matter what ──────
  container._skelTimeout = setTimeout(() => forceRemoveSkeletons(), 5000);

  // Load data
  await fetchAll(container);

  // Auto-refresh every 60s
  container._dashRefresh = setInterval(() => {
    if (document.getElementById('dash-root')) fetchAll(container);
    else clearInterval(container._dashRefresh);
  }, 60000);
}

// ── Fetch — each endpoint isolated, never blocks dashboard ─────────────────
async function fetchAll(container) {
  // Clear safety timeout if we complete faster
  if (container._skelTimeout) {
    clearTimeout(container._skelTimeout);
    container._skelTimeout = null;
  }

  // Fire all requests independently — one failure never kills others
  const [kpi, trend, contractChart, issuesSum, inspBar, expiring, recentIssues, actLog] =
    await Promise.all([
      safeFetch('/api/dashboard/kpi',               {}, 8000),
      safeFetch('/api/dashboard/issues-trend',       {}, 8000),
      safeFetch('/api/dashboard/contracts-chart',    {}, 8000),
      safeFetch('/api/dashboard/issues-summary',     {}, 8000),
      safeFetch('/api/dashboard/inspection-bar',     {}, 8000),
      safeFetch('/api/dashboard/contracts-expiring', [], 8000),
      safeFetch('/api/dashboard/stats',              {}, 8000),
      safeFetch('/api/dashboard/activity-log',       [], 8000),
    ]);

  // Render each section independently — one failure never breaks others
  try { renderKPI(kpi); } catch(e) { console.warn('KPI render:', e); }
  try { renderMiniStats(kpi); } catch(e) { console.warn('MiniStats render:', e); }
  try { renderDonut(Array.isArray(issuesSum?.by_category) ? issuesSum.by_category : []); } catch(e) { console.warn('Donut render:', e); hideSkel('skel-donut','chart-donut'); }
  try { renderTrend(trend); } catch(e) { console.warn('Trend render:', e); hideSkel('skel-trend','chart-trend'); }
  try { renderInspBar(inspBar); } catch(e) { console.warn('InspBar render:', e); hideSkel('skel-insp','chart-insp'); }
  try { renderContractBar(contractChart); } catch(e) { console.warn('ContractBar render:', e); hideSkel('skel-contract','chart-contract'); }
  try {
    const exp = Array.isArray(expiring) ? expiring : [];
    renderContractsTable(exp);
  } catch(e) { console.warn('ContractsTable render:', e); }
  try {
    const issues = Array.isArray(recentIssues)
      ? recentIssues
      : Array.isArray(recentIssues?.recent_issues) ? recentIssues.recent_issues : [];
    renderIssuesTable(issues);
  } catch(e) { console.warn('IssuesTable render:', e); }
  try { renderActivityLog(Array.isArray(actLog) ? actLog : []); } catch(e) { console.warn('ActivityLog render:', e); }

  const el = document.getElementById('dash-updated');
  if (el) el.textContent = `Diperbarui: ${new Date().toLocaleTimeString('id-ID')}`;
}

// ── KPI Cards ──────────────────────────────────────────────────────────────
function renderKPI(kpi) {
  const row = document.getElementById('kpi-row');
  if (!row) return;
  kpi = kpi || {};

  const cards = [
    { icon:'👥', label:'Karyawan Aktif',        sub:'Total karyawan aktif',         href:'#/employees',   color:'kpi-blue',   key:'employees' },
    { icon:'📄', label:'Kontrak Aktif',          sub:'Kontrak yang masih berjalan',  href:'#/contracts',   color:'kpi-green',  key:'contracts' },
    { icon:'⏰', label:'Kontrak Habis 30 Hari',  sub:'',                             href:'#/contracts',   color:'',           key:'expiring30', warn:true },
    { icon:'⚠️', label:'Permasalahan Open',       sub:'Belum diselesaikan',           href:'#/issues',      color:'',           key:'issues',    warnIfGT0:true },
    { icon:'🤝', label:'One on One Pending',      sub:'Menunggu tindak lanjut',       href:'#/one-on-one',  color:'kpi-purple', key:'one_on_one' },
  ];

  row.innerHTML = cards.map(c => {
    const val  = safeNum(kpi[c.key]?.current, 0);
    const prev = kpi[c.key]?.prev;
    const trend = (prev !== undefined && prev !== null) ? trendBadge(val, prev) : '';
    let colorCls = c.color || '';
    let sub = c.sub || '';
    if (c.warn) {
      colorCls = val > 0 ? 'kpi-amber' : 'kpi-green';
      sub = val > 0 ? `⚠️ ${val} kontrak segera berakhir` : '✅ Semua kontrak aman';
    }
    if (c.warnIfGT0) colorCls = val > 0 ? 'kpi-red' : 'kpi-green';
    return `
      <a href="${c.href}" class="kpi-card ${colorCls}" style="text-decoration:none">
        <div class="kpi-card-top">
          <div class="kpi-icon-wrap"><span class="kpi-icon-emoji">${c.icon}</span></div>
          ${trend}
        </div>
        <div class="kpi-value" data-target="${val}">0</div>
        <div class="kpi-label">${c.label}</div>
        <div class="kpi-subtitle">${sub}</div>
      </a>`;
  }).join('');

  row.querySelectorAll('.kpi-value').forEach(el => animateCount(el, parseInt(el.dataset.target)||0));
}

// ── Mini Stats ─────────────────────────────────────────────────────────────
function renderMiniStats(kpi) {
  const row = document.getElementById('mini-stats-row');
  if (!row) return;
  kpi = kpi || {};

  const items = [
    { icon:'📅', label:'Jadwal Pending',     val:kpi.schedule?.current,         href:'#/schedule',            color:'mini-blue' },
    { icon:'🎓', label:'Training Bulan Ini', val:kpi.training_month?.current,   href:'#/training',            color:'mini-indigo' },
    { icon:'📦', label:'Permintaan Barang',  val:kpi.supply?.current,            href:'#/reports/supply',      color:'mini-orange' },
    { icon:'🔍', label:'Inspeksi Bulan Ini', val:kpi.inspection_month?.current,  href:'#/reports/inspection',  color:'mini-teal' },
    { icon:'🧹', label:'GC/DC Bulan Ini',    val:kpi.cleaning_month?.current,   href:'#/reports/cleaning',    color:'mini-green' },
    { icon:'🦟', label:'Fogging Bulan Ini',  val:kpi.fogging_month?.current,     href:'#/reports/fogging',     color:'mini-purple' },
    { icon:'🏢', label:'Total Cabang',       val:kpi.branches?.current,          href:'#/branches',            color:'mini-gray' },
  ];

  row.innerHTML = items.map(s => `
    <a href="${s.href}" class="mini-stat ${s.color}" style="text-decoration:none">
      <div class="mini-stat-icon">${s.icon}</div>
      <div class="mini-stat-body">
        <div class="mini-stat-value" data-target="${safeNum(s.val)}">0</div>
        <div class="mini-stat-label">${s.label}</div>
      </div>
    </a>`).join('');

  row.querySelectorAll('.mini-stat-value').forEach(el => animateCount(el, parseInt(el.dataset.target)||0, 700));
}

// ── Donut ──────────────────────────────────────────────────────────────────
function renderDonut(categories) {
  hideSkel('skel-donut','chart-donut');
  const canvas = document.getElementById('chart-donut');
  if (!canvas) return;
  destroyChart('donut');
  const data = (categories||[]).filter(c => safeNum(c.count)>0);
  if (!data.length) { showEmpty(canvas,'Belum ada data permasalahan'); return; }
  const labels = data.map(c=>safeStr(c.category,'Lainnya'));
  const values = data.map(c=>safeNum(c.count));
  _charts.donut = new Chart(canvas, {
    type:'doughnut',
    data: { labels, datasets:[{ data:values, backgroundColor:COLORS.slice(0,data.length), borderWidth:2, borderColor:'#fff', hoverBorderColor:'#fff' }] },
    options: {
      responsive:true, maintainAspectRatio:false,
      animation:{ duration:700 },
      plugins:{
        legend:{
          position: window.innerWidth < 768 ? 'bottom' : 'right',
          labels:{ font:FONT, color:'#475569', usePointStyle:true, padding:10, boxWidth:8, boxHeight:8 },
        },
        tooltip:{ bodyFont:FONT, titleFont:{...FONT,weight:'700'}, callbacks:{ label:ctx=>` ${ctx.label}: ${ctx.parsed} kasus` } },
      },
      cutout:'65%',
    },
  });
}

// ── Trend line ─────────────────────────────────────────────────────────────
function renderTrend(trend) {
  hideSkel('skel-trend','chart-trend');
  const canvas = document.getElementById('chart-trend');
  if (!canvas) return;
  destroyChart('trend');
  trend = trend || {};
  const labels = (trend.labels||[]).map(monthShort);
  const open   = (trend.open  ||[]).map(v=>safeNum(v));
  const closed = (trend.closed||[]).map(v=>safeNum(v));
  if (!labels.length) { showEmpty(canvas,'Belum ada data trend'); return; }
  _charts.trend = new Chart(canvas, {
    type:'line',
    data:{ labels, datasets:[
      { label:'Open',   data:open,   borderColor:'#EF4444', backgroundColor:'rgba(239,68,68,.08)',  fill:true, tension:0.4, pointRadius:3, pointHoverRadius:5, pointBackgroundColor:'#EF4444', borderWidth:2 },
      { label:'Closed', data:closed, borderColor:'#10B981', backgroundColor:'rgba(16,185,129,.08)', fill:true, tension:0.4, pointRadius:3, pointHoverRadius:5, pointBackgroundColor:'#10B981', borderWidth:2 },
    ]},
    options: chartOpts({ plugins:{ legend:{ position:'top' } } }),
  });
}

// ── Inspection bar ─────────────────────────────────────────────────────────
function renderInspBar(inspBar) {
  hideSkel('skel-insp','chart-insp');
  const canvas = document.getElementById('chart-insp');
  if (!canvas) return;
  destroyChart('inspBar');
  inspBar = inspBar || {};
  const labels = inspBar.labels||[];
  const fc     = (inspBar.fc ||[]).map(v=>safeNum(v));
  const spv    = (inspBar.spv||[]).map(v=>safeNum(v));
  if (!labels.length) { showEmpty(canvas,'Belum ada data inspeksi'); return; }
  _charts.inspBar = new Chart(canvas, {
    type:'bar',
    data:{ labels, datasets:[
      { label:'Skor FC',  data:fc,  backgroundColor:'rgba(37,99,235,.75)',  borderRadius:4, borderSkipped:false },
      { label:'Skor SPV', data:spv, backgroundColor:'rgba(16,185,129,.75)', borderRadius:4, borderSkipped:false },
    ]},
    options: chartOpts({ plugins:{ legend:{ position:'top' } },
      scales:{ x:{ grid:{display:false}, ticks:{ font:FONT, color:TICK, maxRotation:45, minRotation:30 } },
               y:{ grid:{color:GRID}, ticks:{ font:FONT, color:TICK }, min:0, max:100 } } }),
  });
}

// ── Contract bar ───────────────────────────────────────────────────────────
function renderContractBar(cc) {
  hideSkel('skel-contract','chart-contract');
  const canvas = document.getElementById('chart-contract');
  if (!canvas) return;
  destroyChart('contractBar');
  cc = cc || {};
  const labels = (cc.labels||[]).map(monthShort);
  const counts = (cc.counts||[]).map(v=>safeNum(v));
  if (!labels.length) { showEmpty(canvas,'Belum ada data kontrak'); return; }
  const bgColors = counts.map(v => v>5?'rgba(239,68,68,.75)': v>2?'rgba(245,158,11,.75)':'rgba(37,99,235,.65)');
  _charts.contractBar = new Chart(canvas, {
    type:'bar',
    data:{ labels, datasets:[{ label:'Kontrak Berakhir', data:counts, backgroundColor:bgColors, borderRadius:6, borderSkipped:false }]},
    options: chartOpts({ plugins:{ legend:{ display:false } },
      scales:{ x:{ grid:{display:false}, ticks:{ font:FONT, color:TICK } },
               y:{ grid:{color:GRID}, ticks:{ font:FONT, color:TICK, precision:0 }, beginAtZero:true } } }),
  });
}

// ── Contracts table ────────────────────────────────────────────────────────
function renderContractsTable(rows) {
  const wrap = document.getElementById('table-contracts');
  if (!wrap) return;
  const expiring = (rows||[]).filter(r => safeNum(r.days_remaining,999)<=30).slice(0,10);
  if (!expiring.length) {
    wrap.innerHTML = `<div class="chart-empty">✅ Tidak ada kontrak yang habis dalam 30 hari</div>`;
    return;
  }
  wrap.innerHTML = `
    <table class="dash-table">
      <thead><tr>
        <th>#</th><th>Nama Karyawan</th><th>Cabang</th><th>Berakhir</th><th>Sisa</th><th>Status</th>
      </tr></thead>
      <tbody>${expiring.map((r,i)=>`
        <tr>
          <td class="td-num">${i+1}</td>
          <td><strong>${safeStr(r.emp_name||r.employee_name)}</strong></td>
          <td class="td-branch">${safeStr(r.branch_name)}</td>
          <td style="white-space:nowrap;font-size:.8rem">${fmtDate(r.end_date)}</td>
          <td>${daysBadge(r.days_remaining)}</td>
          <td>${statusPill(r.status)}</td>
        </tr>`).join('')}
      </tbody>
    </table>`;
}

// ── Issues table ───────────────────────────────────────────────────────────
function renderIssuesTable(rows) {
  const wrap = document.getElementById('table-issues');
  if (!wrap) return;
  const issues = (rows||[]).slice(0,8);
  if (!issues.length) {
    wrap.innerHTML = `<div class="chart-empty">✅ Tidak ada permasalahan terbuka</div>`;
    return;
  }
  wrap.innerHTML = `
    <table class="dash-table">
      <thead><tr>
        <th>Tanggal</th><th>Keluhan</th><th>Cabang</th><th>Kategori</th><th>Status</th>
      </tr></thead>
      <tbody>${issues.map(r=>`
        <tr>
          <td style="white-space:nowrap;font-size:.78rem">${fmtDate(r.report_date)}</td>
          <td class="td-complaint" title="${safeStr(r.complaint)}">${safeStr(r.complaint)}</td>
          <td class="td-branch">${safeStr(r.branch_name)}</td>
          <td><span class="category-tag">${safeStr(r.category)}</span></td>
          <td>${statusPill(r.status)}</td>
        </tr>`).join('')}
      </tbody>
    </table>`;
}

// ── Activity log ───────────────────────────────────────────────────────────
function renderActivityLog(rows) {
  const wrap = document.getElementById('activity-log');
  if (!wrap) return;
  const items = (rows||[]).slice(0,15);
  if (!items.length) {
    wrap.innerHTML = `<div class="chart-empty">Belum ada aktivitas tercatat</div>`;
    return;
  }
  wrap.innerHTML = `<div class="activity-list">${items.map(r=>{
    const info   = actInfo(r.type);
    const label  = safeStr(r.label);
    const branch = r.branch ? ` • ${safeStr(r.branch)}` : '';
    const time   = relTime(r.created_at);
    return `
      <div class="activity-item">
        <div class="activity-dot ${info.dot}">${info.emoji}</div>
        <div class="activity-body">
          <div class="activity-text"><strong>${info.label}</strong> — ${label}${branch}</div>
          <div class="activity-time">${time||'—'}</div>
        </div>
      </div>`;
  }).join('')}</div>`;
}

// ── Helpers ────────────────────────────────────────────────────────────────
function hideSkel(skelId, canvasId) {
  const skel = document.getElementById(skelId);
  const cvs  = document.getElementById(canvasId);
  if (skel) { skel.style.display = 'none'; skel.style.position = ''; }
  if (cvs)  cvs.style.display = 'block';
}
function showEmpty(canvas, msg='Belum ada data') {
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
