import { getToken, getUser, clearToken, apiFetch, setUser, setToken } from './config.js';
import { initRouter, registerRoute, navigate } from './router.js';
import { toastError } from './components/toast.js';
import { createModal } from './components/modal.js';

window.formatDate = (d) => {
  if (!d || d === '-') return '';
  const p = d.split('-');
  if (p.length === 3 && p[0].length === 4) return `${p[2]}-${p[1]}-${p[0]}`;
  return d;
};

// Page imports
import { renderDashboard } from './pages/dashboard.js?v=force14';
import { renderLogin } from './pages/login.js';
import { renderEmployees } from './pages/employees.js';
import { renderContracts } from './pages/contracts.js';
import { renderSchedule } from './pages/schedule.js';
import { renderIssues } from './pages/issues.js';
import { renderOneOnOne } from './pages/one_on_one.js';
import { renderTraining } from './pages/training.js';
import { renderRelievers } from './pages/relievers.js';
import { renderInspectionReports } from './pages/inspection_reports.js';
import { renderCleaningReports } from './pages/cleaning_reports.js';
import { renderFoggingReports } from './pages/fogging_reports.js';
import { renderBasecampReports } from './pages/basecamp_reports.js';
import { renderSOP } from './pages/sop.js';
import { renderChecklist } from './pages/checklist.js';
import { renderForms } from './pages/forms.js';
import { renderUsers } from './pages/users.js';
import { renderBranches } from './pages/branches.js';
import { renderCalendar } from './pages/calendar.js';
import { renderProfile } from './pages/profile.js';
import { renderImportPage } from './pages/import.js';
import { renderSP } from './pages/sp.js';
import { renderMutasi } from './pages/mutasi.js';

function requireAuth(handler) {
  return async (ctx) => {
    if (!getToken()) { navigate('/login'); return; }
    return handler(ctx);
  };
}

// ── Live clock ───────────────────────────────────────────────────────────────
let clockInterval = null;
function startClock() {
  if (clockInterval) clearInterval(clockInterval);
  const updateClock = () => {
    const now  = new Date();
    const hms  = now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    const date = now.toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
    const timeEl = document.getElementById('header-clock-time');
    const dateEl = document.getElementById('header-clock-date');
    if (timeEl) timeEl.textContent = hms;
    if (dateEl) dateEl.textContent = date;
  };
  updateClock();
  clockInterval = setInterval(updateClock, 1000);
}

// ── Sidebar nav badges (load from API) ───────────────────────────────────────
async function loadSidebarBadges() {
  try {
    const res = await apiFetch('/api/dashboard/kpi');
    if (!res.ok) return;
    const d = res.data?.data || res.data || {};
    const setBadge = (id, val) => {
      const el = document.getElementById(id);
      if (el) {
        el.textContent = val > 0 ? val : '';
        el.style.display = val > 0 ? 'inline-flex' : 'none';
      }
    };
    setBadge('badge-issues',      d.issues?.current      || 0);
    setBadge('badge-contracts',   d.expiring30?.current  || 0);
    setBadge('badge-oo1',         d.one_on_one?.current  || 0);
    setBadge('badge-schedule',    d.schedule?.current    || 0);
    setBadge('badge-supply',      d.supply?.current      || 0);
  } catch { /* silent */ }
}

// ── Notifications System (load from API) ─────────────────────────────────────
let _notifications = [];
async function loadNotifications() {
  try {
    const res = await apiFetch('/api/dashboard/notifications');
    if (!res.ok) return;
    _notifications = res.data?.data || res.data || [];
    const dot = document.getElementById('notif-dot');
    if (dot) {
      dot.style.display = _notifications.length > 0 ? 'block' : 'none';
      dot.textContent = _notifications.length;
    }
  } catch { /* silent */ }
}

function showNotificationsModal() {
  if (!_notifications.length) {
    createModal({
      title: 'Notifikasi',
      content: '<div class="empty-state"><p>Tidak ada notifikasi baru.</p></div>',
      confirmText: 'Tutup',
      onConfirm: (_, close) => close()
    });
    return;
  }

  const contentHtml = `
    <div class="notif-list" style="max-height: 400px; overflow-y: auto;">
      ${_notifications.map(n => `
        <div class="notif-item notif-severity-${n.severity || 'info'}" style="padding: 12px; border-bottom: 1px solid var(--border); border-left: 4px solid var(--${n.severity === 'danger' ? 'danger' : n.severity === 'warning' ? 'warning' : 'primary'}); margin-bottom: 8px; border-radius: 4px; background: #fff;">
          <div style="font-weight: 600; font-size: 0.9rem; color: var(--text-1);">${n.title}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>📅 ${n.date}</span>
            <span class="badge badge-${n.severity === 'danger' ? 'danger' : n.severity === 'warning' ? 'warning' : 'info'}">${n.type.toUpperCase()}</span>
          </div>
        </div>
      `).join('')}
    </div>
  `;

  createModal({
    title: `Notifikasi (${_notifications.length})`,
    content: contentHtml,
    confirmText: 'Tutup',
    onConfirm: (_, close) => close()
  });
}

// ── Render Layout ────────────────────────────────────────────────────────────
function renderLayout() {
  const user = getUser();
  const initial = (user?.full_name || 'U')[0].toUpperCase();

  document.getElementById('app').innerHTML = `
    <div class="app-layout">
      <!-- Sidebar dark -->
      <aside class="sidebar" id="sidebar">
        <div class="sidebar-header">
          <div class="sidebar-logo">
            <span class="logo-icon-wrap">🏥</span>
            <span class="logo-text">FC<strong>MS</strong></span>
          </div>
          <button class="sidebar-close" id="sidebar-close" aria-label="Close">✕</button>
        </div>

        <nav class="sidebar-nav" id="sidebar-nav">

          <!-- Utama -->
          <div class="nav-section">
            <span class="nav-section-label">UTAMA</span>
            <a href="#/dashboard" class="nav-item" data-route="/dashboard">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>
              </span>
              <span class="nav-label">Dashboard</span>
            </a>
            <a href="#/calendar" class="nav-item" data-route="/calendar">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
              </span>
              <span class="nav-label">Kalender</span>
            </a>
          </div>

          <!-- SDM -->
          <div class="nav-section">
            <span class="nav-section-label">SDM</span>
            <a href="#/employees" class="nav-item" data-route="/employees">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="9" cy="7" r="4"/><path d="M3 21v-2a4 4 0 014-4h4a4 4 0 014 4v2"/><circle cx="19" cy="7" r="2"/><path d="M23 21v-1a3 3 0 00-3-3"/></svg>
              </span>
              <span class="nav-label">Master Karyawan</span>
            </a>
            <a href="#/contracts" class="nav-item" data-route="/contracts">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>
              </span>
              <span class="nav-label">Data Kontrak</span>
              <span class="nav-badge" id="badge-contracts"></span>
            </a>
            <a href="#/sp" class="nav-item" data-route="/sp">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>
              </span>
              <span class="nav-label">Data Sp</span>
            </a>
            <a href="#/mutasi" class="nav-item" data-route="/mutasi">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="16 3 21 3 21 8"/><line x1="4" y1="20" x2="21" y2="3"/><polyline points="21 16 21 21 16 21"/><line x1="15" y1="15" x2="21" y2="21"/><line x1="4" y1="4" x2="9" y2="9"/></svg>
              </span>
              <span class="nav-label">Data Mutasi</span>
            </a>
            <a href="#/relievers" class="nav-item" data-route="/relievers">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="17 1 21 5 17 9"/><path d="M3 11V9a4 4 0 014-4h14"/><polyline points="7 23 3 19 7 15"/><path d="M21 13v2a4 4 0 01-4 4H3"/></svg>
              </span>
              <span class="nav-label">Jadwal Reliefer</span>
            </a>
          </div>

          <!-- Operasional -->
          <div class="nav-section">
            <span class="nav-section-label">OPERASIONAL</span>
            <a href="#/timeline" class="nav-item" data-route="/timeline">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
              </span>
              <span class="nav-label">Time Line</span>
              <span class="nav-badge" id="badge-schedule"></span>
            </a>
            <a href="#/issues" class="nav-item" data-route="/issues">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
              </span>
              <span class="nav-label">Permasalahan</span>
              <span class="nav-badge badge-danger" id="badge-issues"></span>
            </a>
            <a href="#/one-on-one" class="nav-item" data-route="/one-on-one">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M21 15a2 2 0 01-2 2H7l-4 4V5a2 2 0 012-2h14a2 2 0 012 2z"/></svg>
              </span>
              <span class="nav-label">One on One</span>
              <span class="nav-badge badge-warning" id="badge-oo1"></span>
            </a>
            <a href="#/training" class="nav-item" data-route="/training">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M22 10v6M2 10l10-5 10 5-10 5z"/><path d="M6 12v5c3 3 9 3 12 0v-5"/></svg>
              </span>
              <span class="nav-label">Training</span>
            </a>
          </div>

          <!-- Laporan -->
          <div class="nav-section">
            <span class="nav-section-label">LAPORAN</span>
            <a href="#/reports/inspection" class="nav-item" data-route="/reports/inspection">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              </span>
              <span class="nav-label">Report Inspeksi Hygiene 2026</span>
            </a>
            <a href="#/reports/cleaning" class="nav-item" data-route="/reports/cleaning">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg>
              </span>
              <span class="nav-label">Report GCDC 2026</span>
            </a>
            <a href="#/reports/fogging" class="nav-item" data-route="/reports/fogging">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>
              </span>
              <span class="nav-label">Report Fogging 2026</span>
            </a>
            <a href="#/reports/basecamp" class="nav-item" data-route="/reports/basecamp">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/></svg>
              </span>
              <span class="nav-label">Rekap Laporan Basecamp</span>
            </a>
            <a href="#/reports/supply" class="nav-item" data-route="/reports/supply">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M6 2L3 6v14a2 2 0 002 2h14a2 2 0 002-2V6l-3-4z"/><line x1="3" y1="6" x2="21" y2="6"/><path d="M16 10a4 4 0 01-8 0"/></svg>
              </span>
              <span class="nav-label">Permintaan Chemical</span>
              <span class="nav-badge" id="badge-supply"></span>
            </a>
          </div>

          <!-- Referensi -->
          <div class="nav-section">
            <span class="nav-section-label">REFERENSI</span>
            <a href="#/sop" class="nav-item" data-route="/sop">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M4 19.5A2.5 2.5 0 016.5 17H20"/><path d="M6.5 2H20v20H6.5A2.5 2.5 0 014 19.5v-15A2.5 2.5 0 016.5 2z"/></svg>
              </span>
              <span class="nav-label">SOP</span>
            </a>
            <a href="#/checklist" class="nav-item" data-route="/checklist">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="9 11 12 14 22 4"/><path d="M21 12v7a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h11"/></svg>
              </span>
              <span class="nav-label">Master Checklist</span>
            </a>
            <a href="#/forms" class="nav-item" data-route="/forms">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
              </span>
              <span class="nav-label">Master Form</span>
            </a>
          </div>

          <!-- Admin -->
          ${user && (user.role === 'superadmin' || user.role === 'admin') ? `
          <div class="nav-section">
            <span class="nav-section-label">ADMIN</span>
            <a href="#/users" class="nav-item" data-route="/users">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>
              </span>
              <span class="nav-label">Manajemen User</span>
            </a>
            <a href="#/branches" class="nav-item" data-route="/branches">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>
              </span>
              <span class="nav-label">Cabang</span>
            </a>
            <a href="#/settings/import" class="nav-item" data-route="/settings/import">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
              </span>
              <span class="nav-label">Import Data Awal</span>
            </a>
          </div>` : ''}
        </nav>

        <div class="sidebar-footer">
          <div class="sidebar-user">
            <div class="sidebar-avatar">BA</div>
            <div class="sidebar-user-info">
              <div class="sidebar-user-name">Berlin Ariansyah</div>
              <div class="sidebar-user-role">Administrator</div>
            </div>
          </div>
          <button class="sidebar-logout" id="logout-btn">
            <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
            Keluar
          </button>
        </div>
      </aside>

      <!-- Mobile overlay -->
      <div class="sidebar-overlay" id="sidebar-overlay"></div>

      <!-- Main wrapper -->
      <div class="main-wrapper">
        <!-- Topbar -->
        <header class="topbar" id="topbar">
          <div class="topbar-left">
            <button class="topbar-menu-btn" id="topbar-menu-btn" aria-label="Menu">
              <svg width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="18" x2="21" y2="18"/></svg>
            </button>
            <div class="topbar-welcome" style="display:flex; align-items:center; gap:6px; font-size:1.15rem; font-weight:500; color:var(--text-2); margin-left:16px;">
              Welcome, <span style="font-weight:800; background:linear-gradient(90deg, #2563EB, #8B5CF6); -webkit-background-clip:text; -webkit-text-fill-color:transparent; letter-spacing:-0.5px;">Berlin. Ariansyah</span> 👋
            </div>
          </div>

          <div class="topbar-center" id="topbar-clock">
            <div class="header-clock">
              <div class="header-clock-time" id="header-clock-time">00:00:00</div>
              <div class="header-clock-date" id="header-clock-date">Memuat...</div>
            </div>
          </div>

          <div class="topbar-right">
            <button class="topbar-icon-btn" id="btn-fullscreen" title="Fullscreen" aria-label="Fullscreen">
              <svg width="17" height="17" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/><line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/></svg>
            </button>
            <button class="topbar-icon-btn notif-btn" id="btn-notif" title="Notifikasi" aria-label="Notifikasi">
              <svg width="17" height="17" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 8A6 6 0 006 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 01-3.46 0"/></svg>
              <span class="notif-dot" id="notif-dot" style="display:none"></span>
            </button>
            <a href="#/profile" class="topbar-user-btn" title="Profil">
              <img src="https://ui-avatars.com/api/?name=Berlin+Ariansyah&background=2563EB&color=fff&bold=true" class="topbar-avatar" alt="Avatar" />
              <div class="topbar-user-text">
                <span class="topbar-user-name">Berlin Ariansyah</span>
                <span class="topbar-user-role-mini">Administrator</span>
              </div>
              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24" style="margin-left:4px;color:var(--gray-400)"><polyline points="6 9 12 15 18 9"/></svg>
            </a>
          </div>
        </header>

        <main id="main-content" class="main-content"></main>
      </div>
    </div>
  `;

  // ── Sidebar toggle ────────────────────────────────────────────────────────
  const sidebar  = document.getElementById('sidebar');
  const overlay  = document.getElementById('sidebar-overlay');
  const menuBtn  = document.getElementById('topbar-menu-btn');
  const closeBtn = document.getElementById('sidebar-close');
  const openSidebar  = () => { sidebar.classList.add('open'); overlay.classList.add('show'); };
  const closeSidebar = () => { sidebar.classList.remove('open'); overlay.classList.remove('show'); };
  menuBtn?.addEventListener('click', openSidebar);
  closeBtn?.addEventListener('click', closeSidebar);
  overlay?.addEventListener('click', closeSidebar);
  document.querySelectorAll('.nav-item').forEach(el => el.addEventListener('click', closeSidebar));

  // ── Active nav ───────────────────────────────────────────────────────────
  function updateActiveNav() {
    const hash = window.location.hash.replace('#', '') || '/dashboard';
    document.querySelectorAll('.nav-item').forEach(el => {
      const route = el.dataset.route;
      el.classList.toggle('active', hash === route || (route !== '/dashboard' && hash.startsWith(route)));
    });
    const titleEl = document.getElementById('topbar-title');
    const active  = document.querySelector('.nav-item.active .nav-label');
    if (titleEl && active) titleEl.textContent = active.textContent;
  }
  window.addEventListener('hashchange', updateActiveNav);
  updateActiveNav();

  // ── Live clock ───────────────────────────────────────────────────────────
  startClock();

  // ── Fullscreen ───────────────────────────────────────────────────────────
  document.getElementById('btn-fullscreen')?.addEventListener('click', () => {
    if (!document.fullscreenElement) document.documentElement.requestFullscreen?.();
    else document.exitFullscreen?.();
  });

  // ── Logout ───────────────────────────────────────────────────────────────
  document.getElementById('logout-btn')?.addEventListener('click', async () => {
    await apiFetch('/api/auth/logout', { method: 'POST' });
    clearToken();
    if (clockInterval) clearInterval(clockInterval);
    navigate('/login');
  });

  // ── Sidebar badges ───────────────────────────────────────────────────────
  loadSidebarBadges();
  loadNotifications();

  // ── Click notif button ───────────────────────────────────────────────────
  document.getElementById('btn-notif')?.addEventListener('click', (e) => {
    e.preventDefault();
    showNotificationsModal();
  });
}

// ── Routes ───────────────────────────────────────────────────────────────────
async function init() {
  registerRoute('/login',              ({ main }) => renderLogin(main));
  registerRoute('/dashboard',          requireAuth(({ main }) => renderDashboard(main)));
  registerRoute('/calendar',           requireAuth(({ main }) => renderCalendar(main)));
  registerRoute('/employees',          requireAuth(({ main }) => renderEmployees(main)));
  registerRoute('/contracts',          requireAuth(({ main }) => renderContracts(main)));
  registerRoute('/sp',                 requireAuth(({ main }) => renderSP(main)));
  registerRoute('/mutasi',             requireAuth(({ main }) => renderMutasi(main)));
  registerRoute('/timeline',           requireAuth(({ main }) => renderSchedule(main)));
  registerRoute('/issues',             requireAuth(({ main }) => renderIssues(main)));
  registerRoute('/one-on-one',         requireAuth(({ main }) => renderOneOnOne(main)));
  registerRoute('/training',           requireAuth(({ main }) => renderTraining(main)));
  registerRoute('/relievers',          requireAuth(({ main }) => renderRelievers(main)));
  registerRoute('/reports/inspection', requireAuth(({ main }) => renderInspectionReports(main)));
  registerRoute('/reports/cleaning',   requireAuth(({ main }) => renderCleaningReports(main)));
  registerRoute('/reports/fogging',    requireAuth(({ main }) => renderFoggingReports(main)));
  registerRoute('/reports/basecamp',   requireAuth(({ main }) => renderBasecampReports(main)));
  registerRoute('/reports/supply',     requireAuth(({ main }) => renderForms(main, 'supply')));
  registerRoute('/sop',                requireAuth(({ main }) => renderSOP(main)));
  registerRoute('/checklist',          requireAuth(({ main }) => renderChecklist(main)));
  registerRoute('/forms',              requireAuth(({ main }) => renderForms(main)));
  registerRoute('/users',              requireAuth(({ main }) => renderUsers(main)));
  registerRoute('/branches',           requireAuth(({ main }) => renderBranches(main)));
  registerRoute('/profile',            requireAuth(({ main }) => renderProfile(main)));
  registerRoute('/settings/import',    requireAuth(({ main }) => renderImportPage(main)));

  const token = getToken();
  if (!token && window.location.hash !== '#/login') { navigate('/login'); }

  if (token) {
    const res = await apiFetch('/api/auth/me');
    if (res.ok) { setUser(res.data.data); renderLayout(); }
    else { clearToken(); navigate('/login'); }
  }

  window.addEventListener('fm:login', () => { renderLayout(); navigate('/dashboard'); });
  initRouter();
}

init();
