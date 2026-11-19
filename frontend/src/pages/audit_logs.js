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
      { key: 'created_at', label: 'Waktu', nowrap: true, render: v => {
          const d = new Date(v);
          return Number.isNaN(d.getTime()) ? escapeHtml(v || '-') : d.toLocaleString('id-ID', { dateStyle: 'short', timeStyle: 'medium' });
      }},
      { key: 'user_name', label: 'Pengguna', render: (v, row) => `<strong>${escapeHtml(v || 'Sistem')}</strong><br><small class="text-muted" style="text-transform:capitalize">${escapeHtml(row.user_role || '')}</small>` },
      { key: 'action', label: 'Aksi', render: v => {
        const colors = {
          CREATE: 'badge-success',
          UPDATE: 'badge-info',
          DELETE: 'badge-danger',
          LOGIN: 'badge-primary',
          LOGOUT: 'badge-neutral'
        };
        return `<span class="badge ${colors[v] || 'badge-neutral'}">${escapeHtml(v)}</span>`;
      }},
      { key: 'module', label: 'Modul', render: v => `<span style="text-transform:capitalize">${escapeHtml((v || '').replace(/_/g, ' '))}</span>` },
      { key: 'target_id', label: 'ID Target' },
      { key: 'id', label: 'Detail', render: (v) => {
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'btn btn-xs btn-outline';
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

       const escHtml = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
       const formatJSON = (str) => {
         if (!str) return 'Tidak ada data';
         try {
           const val = typeof str === 'object' ? str : JSON.parse(str);
           return escHtml(JSON.stringify(val, null, 2));
         } catch {
           return escHtml(String(str));
         }
       };

       const content = `
         <div style="display:flex; gap:1rem; flex-wrap:wrap">
           <div style="flex:1; min-width:300px">
              <h4>Data Lama</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${formatJSON(row.old_data)}</pre>
           </div>
           <div style="flex:1; min-width:300px">
              <h4>Data Baru</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">${formatJSON(row.new_data)}</pre>
           </div>
         </div>
       `;

       const { createModal } = await import('../components/modal.js');
       createModal({
         title: `Detail Audit Log #${id}`,
         content,
         width: '800px',
         hideFooter: true
       });
    } catch {
       alert('Gagal mengambil detail');
    }
  };
}
