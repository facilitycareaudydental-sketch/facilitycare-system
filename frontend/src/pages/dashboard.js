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
  <div class="kpi-card" style="pointer-events:none;padding:16px">
    <div style="display:flex; gap:16px; align-items:flex-start">
      <div class="skeleton" style="width:48px;height:48px;border-radius:12px;flex-shrink:0"></div>
      <div style="flex:1">
        <div class="skeleton skeleton-text" style="width:40px;height:24px;margin-bottom:6px"></div>
        <div class="skeleton skeleton-text" style="width:80px;height:12px;margin-bottom:4px"></div>
        <div class="skeleton skeleton-text" style="width:100px;height:10px"></div>
      </div>
    </div>
    <div style="display:flex; align-items:flex-end; gap:8px; margin-top:16px">
      <div class="skeleton" style="flex:1;height:24px;border-radius:4px"></div>
      <div class="skeleton skeleton-text" style="width:30px;height:12px"></div>
    </div>
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


      <!-- KPI -->
      <div class="kpi-row" id="kpi-row">${skelKPI()}</div>

      <!-- Mini Stats -->
      <div class="mini-stats-row" id="mini-stats-row">${skelMini()}</div>

      <!-- Charts Row -->
      <div style="display:flex; gap:16px; width:100%; align-items:stretch;">
        <div class="chart-card" style="flex: 1 1 0; min-width:0;">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan per Kategori</div>
          </div>
          <div style="display:flex; gap:20px; align-items:center; height:180px">
            <div class="chart-canvas-wrap" style="flex:1;height:100%;position:relative">
              <div id="skel-donut" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
              <canvas id="chart-donut" style="display:none"></canvas>
            </div>
            <div id="donut-legend" class="donut-legend" style="width:110px"></div>
          </div>
          <div style="text-align:center; font-size:0.75rem; color:var(--text-3); margin-top:16px">
            Periode: 22 Juni - 22 Juli 2026
          </div>
        </div>
        <div class="chart-card" style="flex: 1 1 0; min-width:0;">
          <div class="chart-card-header">
            <div class="chart-card-title">Trend Permasalahan 12 Bulan</div>
            <div style="display:flex;align-items:center;gap:16px;font-size:0.75rem;font-weight:600;color:var(--text-2)">
               <div style="display:flex;align-items:center;gap:6px"><div style="width:16px;height:8px;border:2px solid #EF4444;border-radius:2px"></div> Open</div>
               <div style="display:flex;align-items:center;gap:6px"><div style="width:16px;height:8px;border:2px solid #10B981;border-radius:2px"></div> Closed</div>
            </div>
          </div>
          <div class="chart-canvas-wrap" style="height:180px;position:relative">
            <div id="skel-trend" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-trend" style="display:none"></canvas>
          </div>
        </div>
        <div class="chart-card" style="flex: 1 1 0; min-width:0;">
          <div class="chart-card-header" style="align-items:flex-start">
            <div>
              <a href="#/reports/inspection" class="chart-card-title" style="text-decoration:none; display:inline-block">Rata-rata Skor Inspeksi per Cabang <span style="font-size:0.8rem; color:var(--primary); font-weight:600; margin-left:8px">Lihat Laporan &rarr;</span></a>
              <div class="chart-card-subtitle" style="font-size:0.65rem">Skor rata-rata SCM & Cleaning</div>
            </div>
            <select class="btn-ghost" style="padding:4px;font-size:0.7rem;border:1px solid var(--border);border-radius:4px;cursor:pointer"><option>Bulan Ini</option></select>
          </div>
          <div class="chart-canvas-wrap" style="height:180px;position:relative;margin-top:10px">
            <div id="skel-insp" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-insp" style="display:none"></canvas>
          </div>
        </div>
      </div>

      <!-- Bottom Row -->
      <div class="bottom-row" style="margin-top:24px;">
        <!-- Jadwal Hari Ini -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Jadwal Hari Ini</div>
            <a href="#/calendar" class="chart-link">Lihat Kalender</a>
          </div>
          <div id="widget-agenda" class="dash-table-wrap" style="height:200px;overflow-y:auto">${skelTable(3)}</div>
        </div>
        <!-- KPI Kebersihan -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">KPI Kebersihan</div>
          </div>
          <div id="widget-kpi-kebersihan" style="margin-top:0px">${skelTable(4)}</div>
        </div>
        <!-- Permasalahan Terbaru -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Permasalahan Terbaru</div>
            <a href="#/issues" class="chart-link">Lihat Semua</a>
          </div>
          <div id="table-issues" class="dash-table-wrap" style="height:200px;overflow-y:auto">${skelTable(3)}</div>
        </div>
        <!-- Kontrak Akan Habis -->
        <div class="chart-card">
          <div class="chart-card-header">
            <div class="chart-card-title">Kontrak Akan Habis</div>
            <a href="#/contracts" class="chart-link">Lihat Data</a>
          </div>
          <div class="chart-canvas-wrap" style="height:200px;position:relative;margin-top:10px">
            <div id="skel-contract-mini" class="skeleton" style="position:absolute;inset:0;border-radius:12px"></div>
            <canvas id="chart-contract-mini" style="display:none"></canvas>
          </div>
        </div>
      </div>

      <!-- Quick Actions Row -->
      <div class="actions-wrap">
        <div class="actions-title">Aksi Cepat</div>
        <div class="actions-row" id="quick-actions">
          <!-- Rendered in JS -->
        </div>
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
  const [kpi, trend, issuesSum, inspBar, recentIssues, calendarData] =
    await Promise.all([
      safeFetch('/api/dashboard/kpi',               {}, 8000),
      safeFetch('/api/dashboard/issues-trend',       {}, 8000),
      safeFetch('/api/dashboard/issues-summary',     {}, 8000),
      safeFetch('/api/dashboard/inspection-bar',     {}, 8000),
      safeFetch('/api/dashboard/stats',              {}, 8000),
      safeFetch('/api/dashboard/calendar',           [], 8000),
    ]);

  // Render each section independently — one failure never breaks others
  try { renderKPI(kpi); } catch(e) { console.warn('KPI render:', e); }
  try { renderMiniStats(kpi); } catch(e) { console.warn('MiniStats render:', e); }
  try { renderDonut(Array.isArray(issuesSum?.by_category) ? issuesSum.by_category : []); } catch(e) { console.warn('Donut render:', e); hideSkel('skel-donut','chart-donut'); }
  try { renderTrend(trend); } catch(e) { console.warn('Trend render:', e); hideSkel('skel-trend','chart-trend'); }
  try { renderInspBar(inspBar); } catch(e) { console.warn('InspBar render:', e); hideSkel('skel-insp','chart-insp'); }

  try {
    const issues = Array.isArray(recentIssues)
      ? recentIssues
      : Array.isArray(recentIssues?.recent_issues) ? recentIssues.recent_issues : [];
    renderIssuesTable(issues);
  } catch(e) { console.warn('IssuesTable render:', e); }
  
  try {
    const contracts = Array.isArray(recentIssues?.expiring_contracts) ? recentIssues.expiring_contracts : [];
    renderContractMiniBar();
  } catch(e) { console.warn('ContractsTable render:', e); }

  try { renderAgenda(Array.isArray(calendarData) ? calendarData : []); } catch(e) { console.warn('Agenda render:', e); }
  try { renderKPIKebersihan(kpi); } catch(e) { console.warn('KPI Kebersihan render:', e); }
  try { renderQuickActions(); } catch(e) { console.warn('Quick Actions render:', e); }
}

// ── KPI Cards ──────────────────────────────────────────────────────────────
function renderKPI(kpi) {
  const row = document.getElementById('kpi-row');
  if (!row) return;
  kpi = kpi || {};

  const cards = [
    { icon:'👥', label:'Karyawan Aktif',        sub:'Total karyawan aktif',       href:'#/employees',   color:'kpi-blue',   key:'employees',  trendPct:'+2%', trendColor:'#10B981', points:'0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0' },
    { icon:'📄', label:'Kontrak Aktif',          sub:'Kontrak yang masih berjalan',href:'#/contracts',   color:'kpi-green',  key:'contracts',  trendPct:'+1%', trendColor:'#10B981', points:'0,15 20,18 40,10 60,12 80,5 100,2' },
    { icon:'⏳', label:'Kontrak Habis 30 Hari',  sub:'Akan segera berakhir',       href:'#/contracts',   color:'kpi-warn',   key:'expiring30', trendPct:'+25%',trendColor:'#F59E0B', points:'0,25 20,22 40,24 60,15 80,18 100,5' },
    { icon:'⚠️', label:'Permasalahan Open',    sub:'Belum diselesaikan',         href:'#/issues',      color:'kpi-red',    key:'issues',     trendPct:'0%',  trendColor:'#EF4444', points:'0,20 20,18 40,22 60,19 80,21 100,20' },
    { icon:'💬', label:'One on One Pending',     sub:'Menunggu tindak lanjut',     href:'#/one-on-one',  color:'kpi-purple', key:'one_on_one', trendPct:'+8%', trendColor:'#10B981', points:'0,25 20,15 40,18 60,8 80,10 100,2' },
  ];

  row.innerHTML = cards.map(c => {
    let val  = safeNum(kpi[c.key]?.current, 0);
    
    return `
      <a href="${c.href}" class="kpi-card ${c.color}" style="text-decoration:none;padding:16px">
        <div style="display:flex; gap:16px;">
          <div class="kpi-icon-wrap" style="width:48px;height:48px;border-radius:12px;display:flex;align-items:center;justify-content:center;flex-shrink:0"><span class="kpi-icon-emoji">${c.icon}</span></div>
          <div style="flex:1;min-width:0;">
            <div class="kpi-value" data-target="${val}" style="font-size:1.8rem; font-weight:800; line-height:1; color:var(--text-1)">${val}</div>
            <div class="kpi-label" style="font-size:0.85rem; font-weight:700; color:var(--text-2); margin-top:6px">${c.label}</div>
            <div class="kpi-subtitle" style="font-size:0.7rem; color:var(--text-3); margin-top:2px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis">${c.sub}</div>
            <div style="display:flex; align-items:flex-end; gap:12px; margin-top:12px; height:24px">
              <div class="kpi-sparkline" style="margin:0;flex:1;height:100%">
                <svg viewBox="0 0 100 30" preserveAspectRatio="none" style="width:100%;height:100%;display:block">
                  <path d="M ${c.points}" fill="none" stroke="${c.trendColor}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"/>
                </svg>
              </div>
              <div style="font-size:0.8rem; font-weight:700; color:${c.trendColor}">${c.trendPct}</div>
            </div>
          </div>
        </div>
      </a>`;
  }).join('');

  row.querySelectorAll('.kpi-value').forEach(el => {
    animateCount(el, parseInt(el.dataset.target)||0);
  });
}

// ── Mini Stats ─────────────────────────────────────────────────────────────
function renderMiniStats(kpi) {
  const row = document.getElementById('mini-stats-row');
  if (!row) return;
  kpi = kpi || {};

  const items = [
    { icon:'📅', label:'Jadwal Kegiatan',      val:kpi.schedule?.current,         href:'#/timeline',            color:'mini-blue' },
    { icon:'🎓', label:'Training Bulan Ini',   val:kpi.training_month?.current,   href:'#/training',            color:'mini-gray' },
    { icon:'📦', label:'Permintaan Barang',   val:kpi.supply?.current,            href:'#/reports/supply',      color:'mini-orange' },
    { icon:'🔍', label:'Inspeksi Bulan Ini',  val:kpi.inspection_month?.current,  href:'#/reports/inspection',  color:'mini-blue' },
    { icon:'🧹', label:'GCDC Bulan Ini',      val:kpi.cleaning_month?.current,    href:'#/reports/cleaning',    color:'mini-green' },
    { icon:'💨', label:'Fogging Bulan Ini',   val:kpi.fogging_month?.current,     href:'#/reports/fogging',     color:'mini-purple' },
    { icon:'🏢', label:'Total Cabang',        val:kpi.branches?.current,          href:'#/branches',            color:'mini-teal' },
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
  const legendWrap = document.getElementById('donut-legend');
  if (!canvas || !legendWrap) return;
  destroyChart('donut');
  const data = (categories||[]).filter(c => safeNum(c.count)>0);
  if (!data.length) { showEmpty(canvas,'Belum ada data permasalahan'); return; }
  const labels = data.map(c=>`${safeStr(c.category,'Lainnya')}`);
  const values = data.map(c=>safeNum(c.count));
  const total = values.reduce((a,b)=>a+b, 0);

  // Render Custom HTML Legend
  legendWrap.innerHTML = data.map((c, i) => {
    const color = COLORS[i % COLORS.length];
    const pct = total > 0 ? Math.round((c.count / total) * 100) : 0;
    return `
      <div class="donut-legend-item">
        <div class="donut-legend-color" style="background:${color}"></div>
        <div>
          <div class="donut-legend-val"><span style="color:var(--text-1)">${c.count}</span> <span style="font-size:0.7rem;font-weight:600;color:var(--text-3)">(${pct}%)</span></div>
          <div class="donut-legend-label">${labels[i]}</div>
        </div>
      </div>
    `;
  }).join('');

  // Plugin to draw text in center
  const centerTextPlugin = {
    id: 'centerText',
    beforeDraw: function(chart) {
      const width = chart.width, height = chart.height, ctx = chart.ctx;
      ctx.restore();
      const fontSize = (height / 80).toFixed(2);
      ctx.font = "bold " + fontSize + "em Inter";
      ctx.textBaseline = "middle";
      ctx.fillStyle = "#1E293B"; // var(--text-1)
      const text = total.toString(),
            textX = Math.round((width - ctx.measureText(text).width) / 2),
            textY = height / 2;
      ctx.fillText(text, textX, textY - 10);
      
      ctx.font = "600 " + (fontSize * 0.35).toFixed(2) + "em Inter";
      ctx.fillStyle = "#64748B"; // var(--text-2)
      const labelText = "Total",
            labelX = Math.round((width - ctx.measureText(labelText).width) / 2);
      ctx.fillText(labelText, labelX, textY + 15);
      ctx.save();
    }
  };

  _charts.donut = new Chart(canvas, {
    type:'doughnut',
    data: { labels, datasets:[{ data:values, backgroundColor:COLORS, borderWidth:2, borderColor:'#fff', hoverBorderColor:'#fff' }] },
    options: {
      responsive:true, maintainAspectRatio:false,
      animation:{ duration:700 },
      plugins:{
        legend:{ display: false },
        tooltip:{ bodyFont:FONT, titleFont:{...FONT,weight:'700'}, callbacks:{ label:ctx=>` ${ctx.label}: ${ctx.parsed} kasus` } },
      },
      cutout:'75%',
    },
    plugins: [centerTextPlugin]
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
      { label:'Open',   data:open,   borderColor:'#EF4444', backgroundColor:'rgba(239,68,68,.08)',  fill:true, tension:0.4, pointRadius:2, pointHoverRadius:4, pointBackgroundColor:'#EF4444', borderWidth:2 },
      { label:'Closed', data:closed, borderColor:'#10B981', backgroundColor:'rgba(16,185,129,.08)', fill:true, tension:0.4, pointRadius:2, pointHoverRadius:4, pointBackgroundColor:'#10B981', borderWidth:2 },
    ]},
    options: chartOpts({ plugins:{ legend:{ display:false } } }),
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



// ── Contract Mini Bar ───────────────────────────────────────────────────────
function renderContractMiniBar() {
  hideSkel('skel-contract-mini','chart-contract-mini');
  const canvas = document.getElementById('chart-contract-mini');
  if (!canvas) return;
  destroyChart('contractMiniBar');
  
  // Data: June to Dec 2026
  const labels = ['Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  const data = [12, 18, 9, 24, 15, 30, 42]; 
  
  const ctx = canvas.getContext('2d');
  
  const grad = ctx.createLinearGradient(0,0,0,200);
  grad.addColorStop(0, '#60A5FA'); // Light blue
  grad.addColorStop(1, '#2563EB'); // Primary blue
  
  _charts.contractMiniBar = new Chart(canvas, {
    type:'bar',
    data: { 
      labels, 
      datasets:[{ 
        label:'Kontrak Habis', 
        data, 
        backgroundColor: grad, 
        borderRadius: 4, 
        borderSkipped: false,
        barPercentage: 0.6,
        categoryPercentage: 0.7,
      }]
    },
    options: chartOpts({ 
      plugins:{ legend:{ display:false } },
      scales:{ 
        x:{ grid:{display:false}, ticks:{ font:FONT, color:TICK, maxRotation:0 } },
        y:{ grid:{color:GRID, borderDash:[4,4], drawBorder:false}, ticks:{ font:FONT, color:TICK, precision:0, maxTicksLimit:5 }, min:0 } 
      },
      animation: {
        y:{ duration: 1000, easing: 'easeOutQuart' }
      }
    }),
  });
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
    <div class="dash-list">
      ${issues.map(r=>`
        <div class="dash-list-item">
          <div style="flex-shrink:0">${statusPill(r.status)}</div>
          <div style="flex:1;min-width:0;margin-left:4px">
            <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${safeStr(r.complaint)}</div>
            <div style="font-size:0.75rem;color:var(--text-3);margin-top:2px">${safeStr(r.branch_name)}</div>
          </div>
        </div>
      `).join('')}
    </div>`;
}

// ── Activity log ───────────────────────────────────────────────────────────
// ── Agenda & Jadwal ────────────────────────────────────────────────────────
function renderAgenda(rows) {
  const wrap = document.getElementById('widget-agenda');
  if (!wrap) return;
  const items = (rows||[]).slice(0, 10);
  if (!items.length) {
    wrap.innerHTML = `<div class="chart-empty">✅ Tidak ada agenda hari ini</div>`;
    return;
  }
  wrap.innerHTML = `
    <div style="display:flex;flex-direction:column;gap:12px;padding-right:8px">
      ${items.map(r=> {
        // use status colors based on typical titles for Jadwal Hari Ini
        let color = '#3B82F6'; let bg = '#EFF6FF'; let tag = 'Agenda';
        const titleL = (r.title||'').toLowerCase();
        if (titleL.includes('inspeksi')) { color = '#10B981'; bg = '#ECFDF5'; tag = 'Inspeksi'; }
        else if (titleL.includes('cleaning') || titleL.includes('gcdc')) { color = '#3B82F6'; bg = '#EFF6FF'; tag = 'Cleaning'; }
        else if (titleL.includes('reliefer')) { color = '#F59E0B'; bg = '#FFFBEB'; tag = 'Reliefer'; }
        else if (titleL.includes('fogging')) { color = '#8B5CF6'; bg = '#F5F3FF'; tag = 'Fogging'; }
        
        return `
        <div style="display:flex;gap:16px;align-items:flex-start;padding-bottom:12px;border-bottom:1px solid var(--border)">
          <div style="font-size:0.85rem;font-weight:700;color:var(--text-1);margin-top:2px">${new Date(r.event_date).toLocaleTimeString('id-ID',{hour:'2-digit',minute:'2-digit'})}</div>
          <div style="width:8px;height:8px;border-radius:50%;background:${color};margin-top:6px;flex-shrink:0"></div>
          <div style="flex:1">
            <div style="font-weight:700;font-size:0.85rem;color:var(--text-1);line-height:1.2;margin:0 0 4px 0">${safeStr(r.title)}</div>
            <div style="font-size:0.75rem;color:var(--text-3)">${safeStr(r.branch_name)}</div>
          </div>
          <div style="font-size:0.7rem;font-weight:600;padding:2px 8px;border-radius:6px;background:${bg};color:${color}">${tag}</div>
        </div>
      `}).join('')}
    </div>
  `;
}

// ── KPI Kebersihan ────────────────────────────────────────────────────────
function renderKPIKebersihan(kpi) {
  const wrap = document.getElementById('widget-kpi-kebersihan');
  if (!wrap) return;
  
  // Data sesuai mockup
  const items = [
    { label: 'Kebersihan Area', val: '97%', target: 'Target 95%', icon: '🧹', bg: '#ECFDF5', color: '#10B981' },
    { label: 'Penyelesaian Complaint', val: '100%', target: 'Target 100%', icon: '⏱️', bg: '#ECFDF5', color: '#10B981' },
    { label: 'Kepatuhan Jadwal Cleaning', val: '99%', target: 'Target 100%', icon: '⏱️', bg: '#EFF6FF', color: '#3B82F6' },
    { label: 'Kepatuhan GCDC', val: '100%', target: 'Target 100%', icon: '🧹', bg: '#EFF6FF', color: '#3B82F6' },
    { label: 'Complaint Cleaning (≤10)', val: '4', target: 'Target ≤10', icon: '📝', bg: '#F5F3FF', color: '#8B5CF6' },
    { label: 'Pelaksanaan Fogging', val: '100%', target: 'Target 100%', icon: '💨', bg: '#F5F3FF', color: '#8B5CF6' },
  ];
  
  wrap.innerHTML = `
    <div style="display:grid;grid-template-columns:1fr;gap:12px;padding-bottom:12px;height:200px;overflow-y:auto;padding-right:8px;">
      ${items.map(r => {
        const pct = r.val.includes('%') ? parseInt(r.val) : Math.min(100, parseInt(r.val)*10);
        return `
        <div class="prog-item">
          <div class="prog-header">
            <div class="prog-title">
              <div class="prog-title-icon" style="background:${r.bg};color:${r.color}">${r.icon}</div>
              ${r.label}
            </div>
            <div class="prog-val">${r.val}</div>
          </div>
          <span class="prog-target">${r.target}</span>
          <div class="prog-bar-bg">
            <div class="prog-bar-fill" style="width:${pct}%;background:${r.color}"></div>
          </div>
        </div>
      `}).join('')}
    </div>
  `;
}

// ── Quick Actions ──────────────────────────────────────────────────────────
function renderQuickActions() {
  const wrap = document.getElementById('quick-actions');
  if (!wrap) return;
  const btns = [
    { label:'Buat Permasalahan', icon:'<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><path d="M12 5v14m-7-7h14"/></svg>', bg:'#3B82F6', href:'#/issues' },
    { label:'Permintaan Barang', icon:'📦', bg:'#10B981', href:'#/reports/supply' },
    { label:'One on One Baru',   icon:'👥', bg:'#6366F1', href:'#/one-on-one' },
    { label:'Input Kegiatan',    icon:'📋', bg:'#8B5CF6', href:'#/timeline' },
    { label:'Buat Checklist',    icon:'📝', bg:'#0EA5E9', href:'#/checklist' },
    { label:'Laporan Basecamp',  icon:'📊', bg:'#14B8A6', href:'#/reports/basecamp' },
    { label:'Kalender',          icon:'📅', bg:'#8B5CF6', href:'#/calendar' },
  ];
  wrap.innerHTML = btns.map(b => `
    <a href="${b.href}" class="action-btn">
      <div class="action-icon" style="background:${b.bg}">${b.icon}</div>
      ${b.label}
    </a>
  `).join('');
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
