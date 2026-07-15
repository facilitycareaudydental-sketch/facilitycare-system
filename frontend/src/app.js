import { getToken, getUser, clearToken, apiFetch, setUser, setToken } from './config.js';
import { initRouter, registerRoute, navigate } from './router.js';
import { toastError } from './components/toast.js';

// Page imports
import { renderDashboard } from './pages/dashboard.js';
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

function requireAuth(handler) {
  return async (ctx) => {
    if (!getToken()) { navigate('/login'); return; }
    return handler(ctx);
  };
}

function renderLayout() {
  const user = getUser();
  document.getElementById('app').innerHTML = `
    <div class="app-layout">
      <!-- Sidebar -->
      <aside class="sidebar" id="sidebar">
        <div class="sidebar-header">
          <div class="sidebar-logo">
            <span class="logo-icon">🏥</span>
            <span class="logo-text">FM<strong>Ops</strong></span>
          </div>
          <button class="sidebar-close" id="sidebar-close" aria-label="Close sidebar">&times;</button>
        </div>

        <nav class="sidebar-nav">
          <div class="nav-section">
            <span class="nav-section-label">Utama</span>
            <a href="#/dashboard" class="nav-item" data-route="/dashboard">
              <span class="nav-icon">📊</span><span class="nav-label">Dashboard</span>
            </a>
            <a href="#/calendar" class="nav-item" data-route="/calendar">
              <span class="nav-icon">📅</span><span class="nav-label">Kalender</span>
            </a>
          </div>

          <div class="nav-section">
            <span class="nav-section-label">SDM</span>
            <a href="#/employees" class="nav-item" data-route="/employees">
              <span class="nav-icon">👥</span><span class="nav-label">Karyawan</span>
            </a>
            <a href="#/contracts" class="nav-item" data-route="/contracts">
              <span class="nav-icon">📋</span><span class="nav-label">Kontrak</span>
            </a>
            <a href="#/relievers" class="nav-item" data-route="/relievers">
              <span class="nav-icon">🔄</span><span class="nav-label">Reliefer</span>
            </a>
          </div>

          <div class="nav-section">
            <span class="nav-section-label">Operasional</span>
            <a href="#/schedule" class="nav-item" data-route="/schedule">
              <span class="nav-icon">🗓️</span><span class="nav-label">Jadwal Kegiatan</span>
            </a>
            <a href="#/issues" class="nav-item" data-route="/issues">
              <span class="nav-icon">⚠️</span><span class="nav-label">Permasalahan</span>
            </a>
            <a href="#/one-on-one" class="nav-item" data-route="/one-on-one">
              <span class="nav-icon">🤝</span><span class="nav-label">One on One</span>
            </a>
            <a href="#/training" class="nav-item" data-route="/training">
              <span class="nav-icon">🎓</span><span class="nav-label">Training</span>
            </a>
          </div>

          <div class="nav-section">
            <span class="nav-section-label">Laporan</span>
            <a href="#/reports/inspection" class="nav-item" data-route="/reports/inspection">
              <span class="nav-icon">🔍</span><span class="nav-label">Laporan Inspeksi</span>
            </a>
            <a href="#/reports/cleaning" class="nav-item" data-route="/reports/cleaning">
              <span class="nav-icon">🧹</span><span class="nav-label">Laporan GC/DC</span>
            </a>
            <a href="#/reports/fogging" class="nav-item" data-route="/reports/fogging">
              <span class="nav-icon">💨</span><span class="nav-label">Rekap Fogging</span>
            </a>
            <a href="#/reports/basecamp" class="nav-item" data-route="/reports/basecamp">
              <span class="nav-icon">📝</span><span class="nav-label">Laporan Basecamp</span>
            </a>
            <a href="#/reports/supply" class="nav-item" data-route="/reports/supply">
              <span class="nav-icon">📦</span><span class="nav-label">Permintaan Barang</span>
            </a>
          </div>

          <div class="nav-section">
            <span class="nav-section-label">Referensi</span>
            <a href="#/sop" class="nav-item" data-route="/sop">
              <span class="nav-icon">📖</span><span class="nav-label">SOP</span>
            </a>
            <a href="#/checklist" class="nav-item" data-route="/checklist">
              <span class="nav-icon">✅</span><span class="nav-label">Master Checklist</span>
            </a>
            <a href="#/forms" class="nav-item" data-route="/forms">
              <span class="nav-icon">📄</span><span class="nav-label">Master Form</span>
            </a>
          </div>

          ${user && (user.role === 'superadmin' || user.role === 'admin') ? `
          <div class="nav-section">
            <span class="nav-section-label">Admin</span>
            <a href="#/users" class="nav-item" data-route="/users">
              <span class="nav-icon">🔐</span><span class="nav-label">Manajemen User</span>
            </a>
            <a href="#/branches" class="nav-item" data-route="/branches">
              <span class="nav-icon">🏢</span><span class="nav-label">Cabang</span>
            </a>
          </div>
          <div class="nav-section">
            <span class="nav-section-label">Settings</span>
            <a href="#/settings/import" class="nav-item" data-route="/settings/import">
              <span class="nav-icon">📥</span><span class="nav-label">Import Data Awal</span>
            </a>
          </div>` : ''}
        </nav>

        <div class="sidebar-footer">
          <div class="user-info">
            <div class="user-avatar">${(user?.full_name || 'U')[0].toUpperCase()}</div>
            <div class="user-details">
              <div class="user-name">${user?.full_name || 'User'}</div>
              <div class="user-role">${user?.role || ''}</div>
            </div>
          </div>
          <button class="btn btn-ghost btn-sm logout-btn" id="logout-btn">↩ Keluar</button>
        </div>
      </aside>

      <!-- Overlay for mobile -->
      <div class="sidebar-overlay" id="sidebar-overlay"></div>

      <!-- Main content -->
      <div class="main-wrapper">
        <header class="topbar">
          <button class="topbar-menu-btn" id="topbar-menu-btn" aria-label="Toggle menu">☰</button>
          <div class="topbar-title" id="topbar-title">Dashboard</div>
          <div class="topbar-actions">
            <a href="#/profile" class="btn btn-ghost btn-sm topbar-profile">
              <span class="user-avatar-sm">${(user?.full_name || 'U')[0].toUpperCase()}</span>
            </a>
          </div>
        </header>
        <main id="main-content" class="main-content"></main>
      </div>
    </div>
  `;

  // Sidebar toggle for mobile
  const sidebar = document.getElementById('sidebar');
  const overlay = document.getElementById('sidebar-overlay');
  const menuBtn = document.getElementById('topbar-menu-btn');
  const closeBtn = document.getElementById('sidebar-close');

  const openSidebar = () => { sidebar.classList.add('open'); overlay.classList.add('show'); };
  const closeSidebar = () => { sidebar.classList.remove('open'); overlay.classList.remove('show'); };

  menuBtn?.addEventListener('click', openSidebar);
  closeBtn?.addEventListener('click', closeSidebar);
  overlay?.addEventListener('click', closeSidebar);

  // Active nav highlighting
  function updateActiveNav() {
    const hash = window.location.hash.replace('#', '') || '/dashboard';
    document.querySelectorAll('.nav-item').forEach(el => {
      const route = el.dataset.route;
      el.classList.toggle('active', hash === route || (route !== '/dashboard' && hash.startsWith(route)));
    });
    // Update topbar title
    const titleEl = document.getElementById('topbar-title');
    const active = document.querySelector('.nav-item.active .nav-label');
    if (titleEl && active) titleEl.textContent = active.textContent;
  }
  window.addEventListener('hashchange', updateActiveNav);
  updateActiveNav();

  // Logout
  document.getElementById('logout-btn')?.addEventListener('click', async () => {
    await apiFetch('/api/auth/logout', { method: 'POST' });
    clearToken();
    navigate('/login');
  });

  // Close sidebar on nav click (mobile)
  document.querySelectorAll('.nav-item').forEach(el => {
    el.addEventListener('click', closeSidebar);
  });
}

async function init() {
  // Register routes
  registerRoute('/login', ({ main }) => renderLogin(main));
  registerRoute('/dashboard', requireAuth(({ main }) => renderDashboard(main)));
  registerRoute('/calendar', requireAuth(({ main }) => renderCalendar(main)));
  registerRoute('/employees', requireAuth(({ main }) => renderEmployees(main)));
  registerRoute('/contracts', requireAuth(({ main }) => renderContracts(main)));
  registerRoute('/schedule', requireAuth(({ main }) => renderSchedule(main)));
  registerRoute('/issues', requireAuth(({ main }) => renderIssues(main)));
  registerRoute('/one-on-one', requireAuth(({ main }) => renderOneOnOne(main)));
  registerRoute('/training', requireAuth(({ main }) => renderTraining(main)));
  registerRoute('/relievers', requireAuth(({ main }) => renderRelievers(main)));
  registerRoute('/reports/inspection', requireAuth(({ main }) => renderInspectionReports(main)));
  registerRoute('/reports/cleaning', requireAuth(({ main }) => renderCleaningReports(main)));
  registerRoute('/reports/fogging', requireAuth(({ main }) => renderFoggingReports(main)));
  registerRoute('/reports/basecamp', requireAuth(({ main }) => renderBasecampReports(main)));
  registerRoute('/reports/supply', requireAuth(({ main }) => renderForms(main, 'supply')));
  registerRoute('/sop', requireAuth(({ main }) => renderSOP(main)));
  registerRoute('/checklist', requireAuth(({ main }) => renderChecklist(main)));
  registerRoute('/forms', requireAuth(({ main }) => renderForms(main)));
  registerRoute('/users', requireAuth(({ main }) => renderUsers(main)));
  registerRoute('/branches', requireAuth(({ main }) => renderBranches(main)));
  registerRoute('/profile', requireAuth(({ main }) => renderProfile(main)));
  registerRoute('/settings/import', requireAuth(({ main }) => renderImportPage(main)));

  // Redirect to public form page (no login needed)
  registerRoute('/form/chemical', () => { window.location.href = '/form.html'; });

  // Check auth state
  const token = getToken();
  if (!token && window.location.hash !== '#/login') {
    navigate('/login');
  }

  // Render layout if logged in
  if (token) {
    // Verify token and get user
    const res = await apiFetch('/api/auth/me');
    if (res.ok) {
      setUser(res.data.data);
      renderLayout();
    } else {
      clearToken();
      navigate('/login');
    }
  }

  // Watch for login events to re-render layout
  window.addEventListener('fm:login', () => {
    renderLayout();
    navigate('/dashboard');
  });

  initRouter();
}

init();
