import { getToken, getUser, clearToken, apiFetch, setUser, setToken, setCsrfToken, hasPermission } from './config.js';
import { initRouter, registerRoute, navigate } from './router.js';
import { toastError } from './components/toast.js';
import { createModal } from './components/modal.js';
import { initGlobalScrollTabsObserver } from './components/scrollTabs.js';
import { escapeHtml } from './utils/sanitize.js';

window.parseFlexibleDate = (d) => {
  if (!d || d === '-') return '';
  d = String(d).trim();
  // Handle Excel serial numbers (e.g. 45208)
  if (/^\d{5}$/.test(d)) {
    const utc_days = Math.floor(Number(d) - 25569);
    const date_info = new Date(utc_days * 86400 * 1000);
    return date_info.toISOString().split('T')[0];
  }
  // Handle DD/MM/YYYY or DD-MM-YYYY
  if (d.match(/^\d{2}[\/\-]\d{2}[\/\-]\d{4}$/)) {
    const p = d.split(/[\/\-]/);
    return `${p[2]}-${p[1]}-${p[0]}`;
  }
  
  // Handle YYYY-DD-MM (where DD > 12)
  const isoMatch = d.match(/^(\d{4})-(\d{2})-(\d{2})(?:T.*)?$/);
  if (isoMatch) {
    const yyyy = isoMatch[1];
    const p1 = parseInt(isoMatch[2], 10);
    const p2 = parseInt(isoMatch[3], 10);
    if (p1 > 12 && p2 <= 12) {
      return `${yyyy}-${isoMatch[3]}-${isoMatch[2]}`;
    }
  }

  return d.split('T')[0];
};

window.formatDate = (d) => {
  const iso = window.parseFlexibleDate(d);
  if (!iso) return '';
  const p = iso.split('-');
  if (p.length === 3 && p[0].length === 4) {
    const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    const day = parseInt(p[2], 10);
    const month = months[parseInt(p[1], 10) - 1];
    return `${day} ${month} ${p[0]}`;
  }
  return iso;
};

// Page imports
import { renderDashboard } from './pages/dashboard.js';
import { renderLogin } from './pages/login.js';
import { renderEmployees } from './pages/employees.js';
import { renderContracts } from './pages/contracts.js';
import { renderSchedule } from './pages/schedule.js';
import { renderIssues } from './pages/issues.js';
import { renderTraining } from './pages/training.js';
import { renderInspectionReports } from './pages/inspection_reports.js';
import { renderCleaningReports } from './pages/cleaning_reports.js';
import { renderFoggingReports } from './pages/fogging_reports.js';
import { renderBasecampReports } from './pages/basecamp_reports.js';
import { renderSOPFacilityCare } from './pages/sop_facility_care.js';
import { renderSOPSecurity } from './pages/sop_security.js';
import { renderChecklist } from './pages/checklist.js';
import { renderHygieneStandards } from './pages/hygiene_standards.js';
import { renderUsers } from './pages/users.js';
import { renderBranches } from './pages/branches.js';
import { renderCalendar } from './pages/calendar.js';
import { renderProfile } from './pages/profile.js';
import { renderImportPage } from './pages/import.js';
import { renderSP } from './pages/sp.js';
import { renderMutasi } from './pages/mutasi.js';
import { renderOvertime } from './pages/overtime.js';
import { renderAuditLogs } from './pages/audit_logs.js';
import { renderOSMigrations } from './pages/os_migrations.js';
import { renderDAC } from './pages/dac.js';
import { renderSyncDashboard } from './pages/sync_dashboard.js';

function requireAuth(handler) {
  return async (ctx) => {
    if (!getToken() && !getUser()) { navigate('/login'); return; }
    return handler(ctx);
  };
}

// ── Live clock ───────────────────────────────────────────────────────────────
let clockInterval = null;
function startClock() {
  if (clockInterval) clearInterval(clockInterval);
  const updateClock = () => {
    const now  = new Date();
    const hms  = now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', second: '2-digit' }).replace(/\./g, ':');
    const date = now.toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
    const timeEl = document.getElementById('header-clock-time');
    const dateEl = document.getElementById('header-clock-date');
    if (timeEl) timeEl.textContent = hms;
    if (dateEl) dateEl.textContent = date;
  };
  updateClock();
  clockInterval = setInterval(updateClock, 1000);
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
      const count = _notifications.length > 0 ? _notifications.length : 51;
      dot.style.display = 'inline-flex';
      dot.textContent = count;
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
        <div class="notif-item notif-severity-${n.severity || 'info'}" style="padding: 12px; border: 1px solid #223456; border-left: 4px solid var(--${n.severity === 'danger' ? 'danger' : n.severity === 'warning' ? 'warning' : 'primary'}); margin-bottom: 8px; border-radius: 8px; background: #0E182B; color: #F8FAFC;">
          <div style="font-weight: 600; font-size: 0.9rem; color: #F8FAFC;">${escapeHtml(n.title)}</div>
          <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 4px; font-size: 0.75rem; color: var(--text-3);">
            <span>📅 ${escapeHtml(n.date)}</span>
            <span class="badge badge-${n.severity === 'danger' ? 'danger' : n.severity === 'warning' ? 'warning' : 'info'}">${escapeHtml(String(n.type || "").toUpperCase())}</span>
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

  const isCollapsedSaved = (() => {
    try { return localStorage.getItem('fcms_sidebar_collapsed') === '1'; } catch { return false; }
  })();

  document.getElementById('app').innerHTML = `
    <div class="app-layout${isCollapsedSaved ? ' sidebar-collapsed' : ''}" id="app-layout">
      <!-- Sidebar dark -->
      <aside class="sidebar" id="sidebar" aria-label="Navigasi Utama">
        <div class="sidebar-header">
          <div class="sidebar-logo">
            <div class="logo-icon-wrap" style="width:34px;height:34px;border-radius:10px;background:#2563EB;display:flex;align-items:center;justify-content:center;color:#fff;flex-shrink:0;">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M18 20V10M12 20V4M6 20v-6"/></svg>
            </div>
            <div class="logo-text-group">
              <span class="logo-text">FCMS</span>
              <span class="logo-subtext">Facilitycare Management System</span>
            </div>
          </div>
          <button type="button" class="sidebar-collapse-btn" id="sidebar-collapse-btn" aria-label="${isCollapsedSaved ? 'Perluas Sidebar' : 'Perkecil Sidebar'}" title="${isCollapsedSaved ? 'Perluas Sidebar' : 'Perkecil Sidebar'}" aria-expanded="${isCollapsedSaved ? 'false' : 'true'}">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="15 18 9 12 15 6"/></svg>
          </button>
          <button type="button" class="sidebar-close" id="sidebar-close" aria-label="Tutup Menu">✕</button>
        </div>

        <nav class="sidebar-nav" id="sidebar-nav">

          <!-- Utama -->
          <div class="nav-section">
            <span class="nav-section-label">UTAMA</span>
            <a href="#/dashboard" class="nav-item" data-route="/dashboard" data-tooltip="Dashboard" title="Dashboard" aria-label="Dashboard">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1"/><rect x="14" y="3" width="7" height="7" rx="1"/><rect x="3" y="14" width="7" height="7" rx="1"/><rect x="14" y="14" width="7" height="7" rx="1"/></svg>
              </span>
              <span class="nav-label">Dashboard</span>
            </a>
            <a href="#/calendar" class="nav-item" data-route="/calendar" data-tooltip="Kalender" title="Kalender" aria-label="Kalender">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
              </span>
              <span class="nav-label">Kalender</span>
            </a>
          </div>

          <!-- SDM -->
          <div class="nav-section">
            <span class="nav-section-label">SDM</span>
            <a href="#/employees" class="nav-item" data-route="/employees" data-tooltip="Master Karyawan" title="Master Karyawan" aria-label="Master Karyawan">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="9" cy="7" r="4"/><path d="M3 21v-2a4 4 0 014-4h4a4 4 0 014 4v2"/><circle cx="19" cy="7" r="2"/><path d="M23 21v-1a3 3 0 00-3-3"/></svg>
              </span>
              <span class="nav-label">Master Karyawan</span>
            </a>
            <a href="#/contracts" class="nav-item" data-route="/contracts" data-tooltip="Data Kontrak" title="Data Kontrak" aria-label="Data Kontrak">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>
              </span>
              <span class="nav-label">Data Kontrak</span>
            </a>
            <a href="#/sp" class="nav-item" data-route="/sp" data-tooltip="Data SP" title="Data SP" aria-label="Data SP">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>
              </span>
              <span class="nav-label">Data Sp</span>
            </a>
            <a href="#/mutasi" class="nav-item" data-route="/mutasi" data-tooltip="Data Mutasi" title="Data Mutasi" aria-label="Data Mutasi">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="16 3 21 3 21 8"/><line x1="4" y1="20" x2="21" y2="3"/><polyline points="21 16 21 21 16 21"/><line x1="15" y1="15" x2="21" y2="21"/><line x1="4" y1="4" x2="9" y2="9"/></svg>
              </span>
              <span class="nav-label">Data Mutasi</span>
            </a>
            <a href="#/overtime" class="nav-item" data-route="/overtime" data-tooltip="Data Lembur" title="Data Lembur" aria-label="Data Lembur">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
              </span>
              <span class="nav-label">Data Lembur</span>
            </a>
          </div>

          <!-- Operasional -->
          <div class="nav-section">
            <span class="nav-section-label">OPERASIONAL</span>
            <a href="#/timeline" class="nav-item" data-route="/timeline" data-tooltip="Time Line" title="Time Line" aria-label="Time Line">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
              </span>
              <span class="nav-label">Time Line</span>
            </a>
            <a href="#/os-migrations" class="nav-item" data-route="/os-migrations" data-tooltip="Migrasi OS" title="Migrasi OS" aria-label="Migrasi OS">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="16 3 21 3 21 8"/><line x1="4" y1="20" x2="21" y2="3"/><polyline points="21 16 21 21 16 21"/><line x1="15" y1="15" x2="21" y2="21"/><line x1="4" y1="4" x2="9" y2="9"/></svg>
              </span>
              <span class="nav-label">Migrasi OS</span>
            </a>
            <a href="#/issues" class="nav-item" data-route="/issues" data-tooltip="Permasalahan" title="Permasalahan" aria-label="Permasalahan">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
              </span>
              <span class="nav-label">Permasalahan</span>
            </a>
            <a href="#/training" class="nav-item" data-route="/training" data-tooltip="Training" title="Training" aria-label="Training">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M22 10v6M2 10l10-5 10 5-10 5z"/><path d="M6 12v5c3 3 9 3 12 0v-5"/></svg>
              </span>
              <span class="nav-label">Training</span>
            </a>
          </div>

          <!-- Laporan -->
          <div class="nav-section">
            <span class="nav-section-label">LAPORAN</span>
            <a href="#/reports/inspection" class="nav-item" data-route="/reports/inspection" data-tooltip="Report Inspeksi Hygiene 2026" title="Report Inspeksi Hygiene 2026" aria-label="Report Inspeksi Hygiene 2026">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              </span>
              <span class="nav-label">Report Inspeksi Hygiene 2026</span>
            </a>
            <a href="#/reports/cleaning" class="nav-item" data-route="/reports/cleaning" data-tooltip="Report GCDC 2026" title="Report GCDC 2026" aria-label="Report GCDC 2026">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="23 6 13.5 15.5 8.5 10.5 1 18"/><polyline points="17 6 23 6 23 12"/></svg>
              </span>
              <span class="nav-label">Report GCDC 2026</span>
            </a>
            <a href="#/reports/fogging" class="nav-item" data-route="/reports/fogging" data-tooltip="Report Fogging 2026" title="Report Fogging 2026" aria-label="Report Fogging 2026">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/></svg>
              </span>
              <span class="nav-label">Report Fogging 2026</span>
            </a>
            <a href="#/reports/basecamp" class="nav-item" data-route="/reports/basecamp" data-tooltip="Rekap Laporan Basecamp" title="Rekap Laporan Basecamp" aria-label="Rekap Laporan Basecamp">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/></svg>
              </span>
              <span class="nav-label">Rekap Laporan Basecamp</span>
            </a>
          </div>

          <!-- Referensi -->
          <div class="nav-section">
            <span class="nav-section-label">REFERENSI</span>
            <a href="#/sop-fc" class="nav-item" data-route="/sop-fc" data-tooltip="SOP Facility Care" title="SOP Facility Care" aria-label="SOP Facility Care">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>
              </span>
              <span class="nav-label">SOP Facility Care</span>
            </a>
            <a href="#/sop-security" class="nav-item" data-route="/sop-security" data-tooltip="SOP Security" title="SOP Security" aria-label="SOP Security">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>
              </span>
              <span class="nav-label">SOP Security</span>
            </a>
            <a href="#/checklist" class="nav-item" data-route="/checklist" data-tooltip="Master Document" title="Master Document" aria-label="Master Document">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="9 11 12 14 22 4"/><path d="M21 12v7a2 2 0 01-2 2H5a2 2 0 01-2-2V5a2 2 0 012-2h11"/></svg>
              </span>
              <span class="nav-label">Master Document</span>
            </a>
            <a href="#/hygiene-standards" class="nav-item" data-route="/hygiene-standards" data-tooltip="Master Hygiene" title="Master Hygiene" aria-label="Master Hygiene">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01L12 2z"/></svg>
              </span>
              <span class="nav-label">Master Hygiene</span>
            </a>
            <a href="#/dac" class="nav-item" data-route="/dac" data-tooltip="Master DAC" title="Master DAC" aria-label="Master DAC">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
              </span>
              <span class="nav-label">Master DAC</span>
            </a>
          </div>

          <!-- Admin -->
          ${user && user.role === 'superadmin' ? `
          <div class="nav-section">
            <span class="nav-section-label">ADMIN</span>
            <a href="#/users" class="nav-item" data-route="/users" data-tooltip="Manajemen User" title="Manajemen User" aria-label="Manajemen User">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M20 21v-2a4 4 0 00-4-4H8a4 4 0 00-4 4v2"/><circle cx="12" cy="7" r="4"/></svg>
              </span>
              <span class="nav-label">Manajemen User</span>
            </a>
            <a href="#/branches" class="nav-item" data-route="/branches" data-tooltip="Cabang" title="Cabang" aria-label="Cabang">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>
              </span>
              <span class="nav-label">Cabang</span>
            </a>
            <a href="#/audit-logs" class="nav-item" data-route="/audit-logs" data-tooltip="Audit Log" title="Audit Log" aria-label="Audit Log">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>
              </span>
              <span class="nav-label">Audit Log</span>
            </a>
            <a href="#/settings/import" class="nav-item" data-route="/settings/import" data-tooltip="Import Data Awal" title="Import Data Awal" aria-label="Import Data Awal">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 01-2 2H5a2 2 0 01-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>
              </span>
              <span class="nav-label">Import Data Awal</span>
            </a>
          </div>` : ''}
          ${user && user.role !== 'superadmin' && hasPermission(user, 'branches', 'read') ? `
          <div class="nav-section">
            <span class="nav-section-label">MASTER DATA</span>
            <a href="#/branches" class="nav-item" data-route="/branches" data-tooltip="Cabang" title="Cabang" aria-label="Cabang">
              <span class="nav-icon">
                <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>
              </span>
              <span class="nav-label">Cabang</span>
            </a>
          </div>` : ''}
        </nav>

        <div class="sidebar-footer">
          <div class="sidebar-user">
            <div class="sidebar-avatar" data-tooltip="${escapeHtml(user?.full_name || user?.username || 'Profil')}" title="${escapeHtml(user?.full_name || user?.username || 'Profil')}">${(user?.full_name || user?.username || 'U')[0].toUpperCase()}</div>
            <div class="sidebar-user-info">
              <div class="sidebar-user-name">${escapeHtml(user?.full_name || user?.username || 'Guest')}</div>
              <div class="sidebar-user-role">${escapeHtml(user?.role || '-')}</div>
            </div>
          </div>
          <div class="sidebar-footer-actions">
            <a href="#/profile" class="sidebar-logout nav-item-profile" data-route="/profile" data-tooltip="Pengaturan" title="Pengaturan" aria-label="Pengaturan" style="flex:1; text-decoration:none;">
              <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-2 2 2 2 0 01-2-2v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83 0 2 2 0 010-2.83l.06-.06a1.65 1.65 0 00.33-1.82 1.65 1.65 0 00-1.51-1H3a2 2 0 01-2-2 2 2 0 012-2h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 010-2.83 2 2 0 012.83 0l.06.06a1.65 1.65 0 001.82.33H9a1.65 1.65 0 001-1.51V3a2 2 0 012-2 2 2 0 012 2v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 0 2 2 0 010 2.83l-.06.06a1.65 1.65 0 00-.33 1.82V9a1.65 1.65 0 001.51 1H21a2 2 0 012 2 2 2 0 01-2 2h-.09a1.65 1.65 0 00-1.51 1z"/></svg>
              <span class="sidebar-btn-label">Pengaturan</span>
            </a>
            <button type="button" class="sidebar-logout" id="logout-btn" data-tooltip="Keluar" title="Keluar" aria-label="Keluar" style="flex:1;">
              <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><polyline points="16 17 21 12 16 7"/><line x1="21" y1="12" x2="9" y2="12"/></svg>
              <span class="sidebar-btn-label">Keluar</span>
            </button>
          </div>
        </div>
      </aside>
      <div class="fc-sidebar-tooltip" id="fc-sidebar-tooltip" role="tooltip" aria-hidden="true"></div>

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
            <div class="topbar-mobile-logo" style="display:none;align-items:center;gap:8px;">
              <div style="width:28px;height:28px;border-radius:8px;background:linear-gradient(135deg,#2563EB,#0EA5E9);display:flex;align-items:center;justify-content:center;box-shadow:0 2px 8px rgba(37,99,235,.4);color:#fff;flex-shrink:0;">
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M18 20V10M12 20V4M6 20v-6"/></svg>
              </div>
              <span class="logo-text" style="font-size:1.15rem;font-weight:800;letter-spacing:-0.3px;color:#fff;">FCMS</span>
            </div>
            <div class="topbar-welcome">
              <div class="topbar-greeting">
                <span class="topbar-greeting-time">${(() => {
                  const h = new Date().getHours();
                  if (h >= 4 && h < 11) return 'Selamat Pagi';
                  if (h >= 11 && h < 15) return 'Selamat Siang';
                  if (h >= 15 && h < 18) return 'Selamat Sore';
                  return 'Selamat Malam';
                })()}, </span><span class="topbar-greeting-name">${escapeHtml(user?.full_name || user?.username || 'Guest')}</span> 👋
              </div>
              <div class="topbar-subtitle">
                Ringkasan Operasional FCMS Hari Ini
              </div>
            </div>
          </div>

          <div class="topbar-center" id="topbar-clock">
            <div class="header-clock">
              <div class="header-clock-time" id="header-clock-time">00:00:00</div>
              <div class="header-clock-date" id="header-clock-date">Memuat...</div>
            </div>
          </div>

          <div class="topbar-right">
            <button class="topbar-icon-btn notif-btn" id="btn-notif" title="Notifikasi" aria-label="Notifikasi">
              <svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M18 8A6 6 0 006 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 01-3.46 0"/></svg>
              <span class="notif-badge" id="notif-dot" style="display:none">0</span>
            </button>
            <a href="#/profile" class="topbar-user-btn" title="Profil">
              <div class="topbar-avatar" style="display:flex;align-items:center;justify-content:center;background:#2563EB;color:#fff;font-weight:700;font-size:0.85rem;border-radius:50%;">${escapeHtml(String(user?.full_name || user?.username || 'G').trim().charAt(0).toUpperCase())}</div>
              <div class="topbar-user-text">
                <span class="topbar-user-name">${escapeHtml(user?.full_name || user?.username || 'Guest')}</span>
                <span class="topbar-user-role-mini">${escapeHtml(user?.role || '-')}</span>
              </div>
              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24" style="margin-left:4px;color:var(--text-3)"><polyline points="6 9 12 15 18 9"/></svg>
            </a>
          </div>
        </header>

        <main id="main-content" class="main-content"></main>
      </div>
    </div>
  `;

  // ── Sidebar toggle (mobile drawer + desktop/tablet collapse) ──────────────
  const appLayout = document.getElementById('app-layout');
  const sidebar   = document.getElementById('sidebar');
  const overlay   = document.getElementById('sidebar-overlay');
  const menuBtn   = document.getElementById('topbar-menu-btn');
  const closeBtn  = document.getElementById('sidebar-close');
  const collapseBtn = document.getElementById('sidebar-collapse-btn');
  const tooltipEl = document.getElementById('fc-sidebar-tooltip');

  const openSidebar  = () => { sidebar.classList.add('open'); overlay.classList.add('show'); };
  const closeSidebar = () => { sidebar.classList.remove('open'); overlay.classList.remove('show'); };
  const hideTooltip  = () => {
    if (!tooltipEl) return;
    tooltipEl.classList.remove('visible');
    tooltipEl.setAttribute('aria-hidden', 'true');
  };

  const setSidebarCollapsed = (collapsed) => {
    if (!appLayout) return;
    appLayout.classList.toggle('sidebar-collapsed', collapsed);
    if (collapseBtn) {
      const label = collapsed ? 'Perluas Sidebar' : 'Perkecil Sidebar';
      collapseBtn.setAttribute('aria-label', label);
      collapseBtn.setAttribute('title', label);
      collapseBtn.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
    }
    hideTooltip();
    try { localStorage.setItem('fcms_sidebar_collapsed', collapsed ? '1' : '0'); } catch {}
    setTimeout(() => window.dispatchEvent(new Event('resize')), 230);
  };

  collapseBtn?.addEventListener('click', () => {
    const next = !appLayout?.classList.contains('sidebar-collapsed');
    setSidebarCollapsed(next);
  });

  menuBtn?.addEventListener('click', openSidebar);
  closeBtn?.addEventListener('click', closeSidebar);
  overlay?.addEventListener('click', closeSidebar);
  document.querySelectorAll('.nav-item').forEach(el => el.addEventListener('click', closeSidebar));

  // Floating tooltip for collapsed desktop/tablet sidebar items
  const bindTooltipTarget = (el) => {
    const show = () => {
      if (!tooltipEl || !appLayout?.classList.contains('sidebar-collapsed') || window.innerWidth < 768) return;
      const text = el.getAttribute('data-tooltip');
      if (!text) return;
      const rect = el.getBoundingClientRect();
      tooltipEl.textContent = text;
      tooltipEl.style.top = `${Math.round(rect.top + rect.height / 2)}px`;
      tooltipEl.style.left = `${Math.round(rect.right + 10)}px`;
      tooltipEl.classList.add('visible');
      tooltipEl.setAttribute('aria-hidden', 'false');
    };
    el.addEventListener('mouseenter', show);
    el.addEventListener('pointerenter', show);
    el.addEventListener('focus', show);
    el.addEventListener('mouseleave', hideTooltip);
    el.addEventListener('pointerleave', hideTooltip);
    el.addEventListener('blur', hideTooltip);
  };
  document.querySelectorAll('#sidebar [data-tooltip]').forEach(bindTooltipTarget);
  const sidebarNav = document.getElementById('sidebar-nav');
  sidebar?.addEventListener('scroll', hideTooltip, { passive: true });
  sidebarNav?.addEventListener('scroll', hideTooltip, { passive: true });

  // Global horizontal scrollable tabs & pills observer across all pages
  initGlobalScrollTabsObserver(document.getElementById('main-content'));

  // Enforce menu visibility for canonical and legacy operational roles
  const adminRoutes = new Set(['/users', '/settings/import', '/audit-logs', '/sync']);
  const userRole = String(user?.role || '').trim().toLowerCase();
  const canRead = hasPermission(user, 'read');
  document.querySelectorAll('.nav-item').forEach(el => {
    const route = el.dataset.route;
    if (!canRead || (userRole !== 'superadmin' && adminRoutes.has(route))) {
      el.style.display = 'none';
    }
  });
  document.querySelectorAll('.nav-section').forEach(el => {
    const hasVisible = Array.from(el.querySelectorAll('.nav-item')).some(item => item.style.display !== 'none');
    if (!hasVisible) el.style.display = 'none';
  });

  // ── Active nav ───────────────────────────────────────────────────────────
  function updateActiveNav() {
    const rawHash = window.location.hash.replace('#', '') || '/dashboard';
    const [baseRoute] = rawHash.split('?');
    document.querySelectorAll('.nav-item').forEach(el => {
      const route = el.dataset.route;
      const isActive = baseRoute === route || (route !== '/dashboard' && (baseRoute + '/').startsWith(route + '/'));
      el.classList.toggle('active', isActive);
    });
    const profileBtn = document.querySelector('.nav-item-profile');
    if (profileBtn) {
      profileBtn.classList.toggle('active', baseRoute === '/profile');
    }
    const titleEl = document.getElementById('topbar-title');
    const activeItem = document.querySelector('.nav-item.active');
    const active = activeItem?.querySelector('.nav-label');
    if (titleEl && active) titleEl.textContent = active.textContent;
    if (activeItem && typeof activeItem.scrollIntoView === 'function') {
      requestAnimationFrame(() => activeItem.scrollIntoView({ block: 'nearest', inline: 'nearest' }));
    }
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

  // ── Notifications ────────────────────────────────────────────────────────
  loadNotifications();

  // ── Click notif button ───────────────────────────────────────────────────
  document.getElementById('btn-notif')?.addEventListener('click', (e) => {
    e.preventDefault();
    showNotificationsModal();
  });
}

// ── Routes ───────────────────────────────────────────────────────────────────
async function init() {
  registerRoute('/',                   () => navigate('/dashboard'));
  registerRoute('/login',              ({ main }) => renderLogin(main));
  registerRoute('/dashboard',          requireAuth(({ main }) => renderDashboard(main)));
  registerRoute('/calendar',           requireAuth(({ main }) => renderCalendar(main)));
  registerRoute('/employees',          requireAuth(({ main, params }) => renderEmployees(main, params)));
  registerRoute('/contracts',          requireAuth(({ main, params }) => renderContracts(main, params)));
  registerRoute('/sp',                 requireAuth(({ main }) => renderSP(main)));
  registerRoute('/mutasi',             requireAuth(({ main }) => renderMutasi(main)));
  registerRoute('/overtime',           requireAuth(({ main }) => renderOvertime(main)));
  registerRoute('/timeline',           requireAuth(({ main, params }) => renderSchedule(main, params)));
  registerRoute('/os-migrations',      requireAuth(({ main }) => renderOSMigrations(main)));
  registerRoute('/issues',             requireAuth(({ main, params }) => renderIssues(main, params)));
  registerRoute('/training',           requireAuth(({ main }) => renderTraining(main)));
  registerRoute('/reports/inspection', requireAuth(({ main }) => renderInspectionReports(main)));
  registerRoute('/reports/cleaning',   requireAuth(({ main }) => renderCleaningReports(main)));
  registerRoute('/reports/fogging',    requireAuth(({ main, params }) => renderFoggingReports(main, params)));
  registerRoute('/reports/basecamp',   requireAuth(({ main }) => renderBasecampReports(main)));
  registerRoute('/sop',                () => { window.location.hash = '#/sop-fc'; });
  registerRoute('/sop-fc',             requireAuth(({ main }) => renderSOPFacilityCare(main)));
  registerRoute('/sop-security',       requireAuth(({ main }) => renderSOPSecurity(main)));
  registerRoute('/checklist',          requireAuth(({ main }) => renderChecklist(main)));
  registerRoute('/hygiene-standards', requireAuth(({ main }) => renderHygieneStandards(main)));
  registerRoute('/dac',                requireAuth(({ main }) => renderDAC(main)));
  registerRoute('/users',              requireAuth(({ main }) => {
    const user = getUser();
    if (!user || user.role !== 'superadmin') {
      navigate('/dashboard');
      return;
    }
    renderUsers(main);
  }));
  registerRoute('/branches',           requireAuth(({ main }) => renderBranches(main)));
  registerRoute('/profile',            requireAuth(({ main }) => renderProfile(main)));
  registerRoute('/settings/import',    requireAuth(({ main }) => {
    const user = getUser();
    if (!user || user.role !== 'superadmin') {
      navigate('/dashboard');
      return;
    }
    renderImportPage(main);
  }));
  registerRoute('/audit-logs',         requireAuth(({ main }) => {
    const user = getUser();
    if (!user || user.role !== 'superadmin') {
      navigate('/dashboard');
      return;
    }
    renderAuditLogs(main);
  }));
  registerRoute('/sync',               requireAuth(({ main }) => {
    const user = getUser();
    if (!user || user.role !== 'superadmin') {
      navigate('/dashboard');
      return;
    }
    renderSyncDashboard(main);
  }));

  if (window.location.hash !== '#/login') {
    const res = await apiFetch('/api/auth/me');
    if (res.ok && res.data?.data) {
      const meData = { ...res.data.data };
      const recoveredCsrf = typeof meData.csrf_token === 'string' ? meData.csrf_token : null;
      delete meData.csrf_token;
      if (recoveredCsrf) setCsrfToken(recoveredCsrf);
      setUser(meData);
      renderLayout();
    } else {
      clearToken();
      navigate('/login');
    }
  }

  window.addEventListener('fm:login', () => {
    renderLayout();
    navigate('/dashboard');
  });
  initRouter();
}

// ── Global Runtime Stability Safety Nets ─────────────────────────────────────
if (typeof window !== 'undefined') {
  window.addEventListener('unhandledrejection', (event) => {
    const reasonMsg = event?.reason?.message || String(event?.reason || 'Unknown promise rejection');
    console.error('[FCMS Stability] Unhandled promise rejection captured:', reasonMsg);
    event.preventDefault();
  });
  window.addEventListener('error', (event) => {
    const errMsg = event?.message || event?.error?.message || 'Unknown runtime error';
    console.error('[FCMS Stability] Runtime error captured:', errMsg);
  });
}

init().catch((err) => {
  console.error('[FCMS Stability] App initialization error:', err?.message || err);
  toastError('Terjadi kendala saat memulai aplikasi. Silakan muat ulang halaman.');
});


// ── PWA Service Worker & Install Support ─────────────────────────────────────
if ('serviceWorker' in navigator && window.location.protocol.startsWith('http')) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('/sw.js').then((reg) => {
      console.log('[PWA] Service Worker registered with scope:', reg.scope);
      reg.addEventListener('updatefound', () => {
        const newWorker = reg.installing;
        newWorker.addEventListener('statechange', () => {
          if (newWorker.state === 'installed' && navigator.serviceWorker.controller) {
            console.log('[PWA] New version available, reloading on user confirmation or refresh.');
          }
        });
      });
    }).catch((err) => {
      console.warn('[PWA] Service Worker registration failed:', err);
    });
  });
}

// PWA Install Prompt Listener
let deferredInstallPrompt = null;
window.addEventListener('beforeinstallprompt', (e) => {
  e.preventDefault();
  deferredInstallPrompt = e;
  window.__showInstallPrompt = () => {
    if (deferredInstallPrompt) {
      deferredInstallPrompt.prompt();
      deferredInstallPrompt.userChoice.then((choiceResult) => {
        if (choiceResult.outcome === 'accepted') {
          console.log('[PWA] User accepted the install prompt');
        }
        deferredInstallPrompt = null;
      });
    }
  };
});
