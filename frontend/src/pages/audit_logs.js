import { buildCrudPage } from './_crud.js';
import { getUser, apiFetch } from '../config.js';
import { escapeHtml } from '../utils/sanitize.js';
import { fetchExportRows } from '../utils/dataCache.js';
import { downloadExcel } from '../utils/excel.js';

export async function renderAuditLogs(container) {
  const currentUser = getUser();
  if (!currentUser || currentUser.role !== 'superadmin') {
    container.innerHTML = '<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';
    return;
  }

  if (!document.getElementById('fc-audit-log-styles')) {
    const style = document.createElement('style');
    style.id = 'fc-audit-log-styles';
    style.textContent = `
      .audit-detail-btn {
        background: rgba(59, 130, 246, 0.14);
        color: #60A5FA;
        border: 1px solid rgba(96, 165, 250, 0.42);
        border-radius: 7px;
        padding: 5px 12px;
        font-size: 0.8rem;
        font-weight: 700;
        cursor: pointer;
        transition: all 0.15s ease;
        white-space: nowrap;
      }
      .audit-detail-btn:hover,
      .audit-detail-btn:focus-visible {
        background: rgba(59, 130, 246, 0.25);
        color: #93C5FD;
        border-color: #93C5FD;
        text-decoration: underline;
        outline: none;
      }
      .audit-json-pre {
        background: #0F172A;
        color: #F8FAFC;
        border: 1px solid #334155;
        border-radius: 10px;
        padding: 14px;
        font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
        font-size: 0.84rem;
        line-height: 1.65;
        max-height: 360px;
        overflow: auto;
        white-space: pre-wrap;
        word-break: break-word;
        margin: 0;
      }
    `;
    document.head.appendChild(style);
  }

  const ACTION_BADGE_MAP = {
    CREATE: 'badge-success',
    UPDATE: 'badge-info',
    DELETE: 'badge-danger',
    LOGIN: 'badge-primary',
    LOGOUT: 'badge-neutral'
  };

  const SENSITIVE_FIELD_RE = /(password|secret|token|jwt|signature|hmac|api_key|private_key|credential|authorization)/i;
  function redactSensitiveFields(val) {
    if (val === null || val === undefined) return val;
    if (Array.isArray(val)) return val.map(redactSensitiveFields);
    if (typeof val !== 'object') return val;
    const out = {};
    for (const [k, v] of Object.entries(val)) {
      out[k] = SENSITIVE_FIELD_RE.test(k) ? '[REDACTED]' : redactSensitiveFields(v);
    }
    return out;
  }

  function formatAuditPayloadHtml(raw, emptyLabel) {
    if (raw === null || raw === undefined || (typeof raw === 'string' && raw.trim() === '')) {
      return `<span style="color:#CBD5E1; font-style:italic;">${escapeHtml(emptyLabel)}</span>`;
    }
    if (typeof raw === 'object') {
      return escapeHtml(JSON.stringify(redactSensitiveFields(raw), null, 2));
    }
    const str = String(raw);
    try {
      const parsed = JSON.parse(str);
      if (parsed !== null && typeof parsed === 'object') {
        return escapeHtml(JSON.stringify(redactSensitiveFields(parsed), null, 2));
      }
      return escapeHtml(String(parsed));
    } catch {
      return escapeHtml(str);
    }
  }

  function formatAuditTime(v) {
    if (!v) return '-';
    const d = new Date(v);
    return Number.isNaN(d.getTime()) ? escapeHtml(String(v)) : d.toLocaleString('id-ID', { dateStyle: 'short', timeStyle: 'medium' });
  }

  buildCrudPage({
    container,
    title: 'Riwayat Aktivitas',
    icon: '🛡️',
    apiPath: '/api/audit-logs',
    enableMobileFilterSheet: true,
    itemLabel: 'Log',
    canCreate: false,
    canEdit: false,
    canDelete: false,
    bulkDelete: false,
    exportOptions: {
      moduleName: 'audit_logs',
      filename: 'FCMS_Audit_Logs',
      template: false,
      import: false,
      onExport: async (filters) => {
        const rows = await fetchExportRows('/api/audit-logs/export', filters);
        const mapped = (rows || []).map((r) => ({
          'ID': r.id,
          'Waktu': r.created_at || '-',
          'Pengguna': r.user_name || 'Sistem',
          'Role': r.user_role || '-',
          'Aksi': r.action || '-',
          'Modul': r.module || '-',
          'ID Target': r.target_id || '-'
        }));
        downloadExcel(mapped, 'FCMS_Audit_Logs', 'Audit Logs');
      }
    },
    columns: [
      { key: 'created_at', label: 'Waktu', nowrap: true, render: v => `<span style="color:#F8FAFC; font-variant-numeric:tabular-nums;">${formatAuditTime(v)}</span>` },
      { key: 'user_name', label: 'Pengguna', render: (v, row) => `<strong style="color:#F8FAFC; font-weight:700;">${escapeHtml(v || 'Sistem')}</strong><br><small style="color:#CBD5E1; text-transform:capitalize; font-size:0.78rem;">${escapeHtml(row.user_role || '-')}</small>` },
      { key: 'action', label: 'Aksi', render: v => `<span class="badge ${ACTION_BADGE_MAP[v] || 'badge-neutral'}">${escapeHtml(v || '-')}</span>` },
      { key: 'module', label: 'Modul', render: v => `<span style="color:#F8FAFC; text-transform:capitalize; font-weight:500;">${escapeHtml((v || '-').replace(/_/g, ' '))}</span>` },
      { key: 'target_id', label: 'ID Target', render: v => `<span style="color:#CBD5E1; font-family:ui-monospace, monospace;">${escapeHtml(v !== null && v !== undefined && v !== '' ? String(v) : '-')}</span>` },
      { key: 'id', label: 'Detail', render: (v) => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'audit-detail-btn';
        btn.textContent = 'Lihat Detail';
        btn.addEventListener('click', () => window.viewAuditDetail?.(v));
        return btn;
      }}
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari ID, pengguna, modul...' },
      { type: 'select', name: 'action', options: [
          { value: '', label: 'Semua Aksi' },
          { value: 'CREATE', label: 'Tambah (CREATE)' },
          { value: 'UPDATE', label: 'Ubah (UPDATE)' },
          { value: 'DELETE', label: 'Hapus (DELETE)' },
          { value: 'LOGIN', label: 'Masuk (LOGIN)' },
          { value: 'LOGOUT', label: 'Keluar (LOGOUT)' }
      ]},
      { type: 'select', name: 'module', options: [
          { value: '', label: 'Semua Modul' },
          { value: 'employees', label: 'Karyawan' },
          { value: 'contracts', label: 'Kontrak' },
          { value: 'sp', label: 'Data SP' },
          { value: 'mutasi', label: 'Data Mutasi' },
          { value: 'overtime', label: 'Lembur' },
          { value: 'schedule', label: 'Time Line / Jadwal' },
          { value: 'os_migrations', label: 'Migrasi OS' },
          { value: 'issues', label: 'Permasalahan' },
          { value: 'training', label: 'Training' },
          { value: 'reports', label: 'Laporan Operasional' },
          { value: 'sop', label: 'SOP Facility Care & Security' },
          { value: 'checklist', label: 'Master Document' },
          { value: 'hygiene_standards', label: 'Master Hygiene' },
          { value: 'dac', label: 'Master DAC' },
          { value: 'users', label: 'Manajemen User' },
          { value: 'branches', label: 'Manajemen Cabang' },
          { value: 'auth', label: 'Autentikasi' }
      ]}
    ]
  });

  const pageActions = container.querySelector('.page-actions');
  if (pageActions && currentUser && currentUser.role === 'superadmin') {
    const syncBtn = document.createElement('a');
    syncBtn.href = '#/sync';
    syncBtn.id = 'btn-audit-sync-health';
    syncBtn.className = 'btn btn-primary';
    syncBtn.style.cssText = 'text-decoration:none;display:inline-flex;align-items:center;gap:6px';
    syncBtn.innerHTML = '🩺 Sync & Health Dashboard';
    pageActions.appendChild(syncBtn);
  }

  // Attach global detail viewer function using direct GET /api/audit-logs/:id
  window.viewAuditDetail = async (id) => {
    try {
      const result = await apiFetch(`/api/audit-logs/${encodeURIComponent(id)}`);
      const row = result?.data?.data || null;
      if (!result?.ok || !row) return alert('Data tidak ditemukan');

      const actionCls = ACTION_BADGE_MAP[row.action] || 'badge-neutral';
      const content = `
        <div style="display:flex; flex-direction:column; gap:14px;">
          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(150px, 1fr)); gap:10px; background:#0F172A; border:1px solid #334155; border-radius:10px; padding:12px 14px;">
            <div>
              <div style="font-size:0.72rem; color:#CBD5E1; text-transform:uppercase; font-weight:700;">ID Audit Log</div>
              <div style="font-size:0.9rem; color:#F8FAFC; font-weight:700; margin-top:2px;">#${escapeHtml(String(row.id ?? id))}</div>
            </div>
            <div>
              <div style="font-size:0.72rem; color:#CBD5E1; text-transform:uppercase; font-weight:700;">Waktu</div>
              <div style="font-size:0.86rem; color:#F8FAFC; font-weight:600; margin-top:2px;">${formatAuditTime(row.created_at)}</div>
            </div>
            <div>
              <div style="font-size:0.72rem; color:#CBD5E1; text-transform:uppercase; font-weight:700;">Pengguna</div>
              <div style="font-size:0.88rem; color:#F8FAFC; font-weight:700; margin-top:2px;">${escapeHtml(row.user_name || 'Sistem')} <span style="color:#CBD5E1; font-weight:500;">(${escapeHtml(row.user_role || '-')})</span></div>
            </div>
            <div>
              <div style="font-size:0.72rem; color:#CBD5E1; text-transform:uppercase; font-weight:700;">Aksi &amp; Modul</div>
              <div style="display:flex; align-items:center; gap:6px; flex-wrap:wrap; margin-top:3px;">
                <span class="badge ${actionCls}">${escapeHtml(row.action || '-')}</span>
                <span style="font-size:0.84rem; color:#F8FAFC; text-transform:capitalize;">${escapeHtml((row.module || '-').replace(/_/g, ' '))}${row.target_id ? ` (#${escapeHtml(String(row.target_id))})` : ''}</span>
              </div>
            </div>
          </div>

          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(280px, 1fr)); gap:16px;">
            <div>
              <h4 style="margin:0 0 8px 0; font-size:0.92rem; font-weight:700; color:#F8FAFC;">Data Lama</h4>
              <pre class="audit-json-pre">${formatAuditPayloadHtml(row.old_data, 'Tidak ada data lama')}</pre>
            </div>
            <div>
              <h4 style="margin:0 0 8px 0; font-size:0.92rem; font-weight:700; color:#F8FAFC;">Data Baru</h4>
              <pre class="audit-json-pre">${formatAuditPayloadHtml(row.new_data, 'Tidak ada data baru')}</pre>
            </div>
          </div>
        </div>
      `;

      const { createModal } = await import('../components/modal.js');
      createModal({
        title: `Detail Audit Log #${id}`,
        content,
        size: 'lg',
        cancelText: 'Tutup'
      });
    } catch {
      alert('Gagal mengambil detail');
    }
  };
}
