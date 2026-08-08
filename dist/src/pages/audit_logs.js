import { buildCrudPage } from './_crud.js';
import { getUser } from '../config.js';

export async function renderAuditLogs(container) {
  const currentUser = getUser();
  if (!currentUser || !['superadmin', 'admin'].includes(currentUser.role)) {
    container.innerHTML = '<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';
    return;
  }

  buildCrudPage({
    container,
    title: 'Riwayat Aktivitas',
    icon: '🕵️‍♂️',
    apiPath: '/api/audit-logs',
    itemLabel: 'Log',
    canCreate: false,
    canEdit: false,
    canDelete: false,
    bulkDelete: false,
    exportOptions: null,
    columns: [
      { key: 'created_at', label: 'Waktu', nowrap: true, render: v => {
          const d = new Date(v);
          return d.toLocaleString('id-ID', { dateStyle: 'short', timeStyle: 'medium' });
      }},
      { key: 'user_name', label: 'Pengguna', render: (v, row) => `<strong>${v || 'Sistem'}</strong><br><small class="text-muted" style="text-transform:capitalize">${row.user_role || ''}</small>` },
      { key: 'action', label: 'Aksi', render: v => {
        const colors = { CREATE: 'badge-success', UPDATE: 'badge-info', DELETE: 'badge-danger' };
        return `<span class="badge ${colors[v] || 'badge-neutral'}">${v}</span>`;
      }},
      { key: 'module', label: 'Modul', render: v => `<span style="text-transform:capitalize">${(v||'').replace('_', ' ')}</span>` },
      { key: 'target_id', label: 'ID Target' },
      { key: 'id', label: 'Detail', render: (v, row) => `<button class="btn btn-xs btn-outline" onclick="window.viewAuditDetail('${v}')">Lihat Detail</button>` }
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari pengguna, modul...' },
      { type: 'select', name: 'action', options: [
          { value: '', label: 'Semua Aksi' },
          { value: 'CREATE', label: 'Tambah (CREATE)' },
          { value: 'UPDATE', label: 'Ubah (UPDATE)' },
          { value: 'DELETE', label: 'Hapus (DELETE)' }
      ]},
      { type: 'select', name: 'module', options: [
          { value: '', label: 'Semua Modul' },
          { value: 'employees', label: 'Karyawan' },
          { value: 'schedule', label: 'Jadwal' },
          { value: 'issues', label: 'Permasalahan' },
          { value: 'relievers', label: 'Reliefer' },
          { value: 'contracts', label: 'Kontrak' }
      ]}
    ]
  });

  // Attach global detail viewer function
  window.viewAuditDetail = async (id) => {
    // Wait for the modal component to be ready or just use standard fetch since we already have the row in memory.
    // Actually, since buildCrudPage handles data loading, we might not have direct access to the row here easily.
    // Instead, we can fetch the exact row from the API. Or we can just use the cached table data.
    // But since the API returns old_data and new_data, we can show it in an alert or modal.
    try {
       const res = await fetch(\`/api/audit-logs?search=\${id}\`, {
           headers: { 'Authorization': \`Bearer \${localStorage.getItem('fm_token')}\` }
       });
       const data = await res.json();
       const row = (data.data || []).find(r => String(r.id) === String(id));
       if (!row) return alert('Data tidak ditemukan');

       const formatJSON = (str) => {
         if (!str) return 'Tidak ada data';
         try { return JSON.stringify(JSON.parse(str), null, 2); } catch { return str; }
       };

       const content = \`
         <div style="display:flex; gap:1rem; flex-wrap:wrap">
           <div style="flex:1; min-width:300px">
              <h4>Data Lama</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">\${formatJSON(row.old_data)}</pre>
           </div>
           <div style="flex:1; min-width:300px">
              <h4>Data Baru</h4>
              <pre style="background:#f8f9fa; padding:10px; border-radius:5px; font-size:12px; overflow-x:auto; border:1px solid #ddd; max-height:400px; overflow-y:auto;">\${formatJSON(row.new_data)}</pre>
           </div>
         </div>
       \`;

       // Import modal dynamically
       const { createModal } = await import('../components/modal.js');
       createModal({
         title: \`Detail Audit Log #\${id}\`,
         content,
         width: '800px',
         hideFooter: true
       });
    } catch(e) {
       alert('Gagal mengambil detail');
    }
  };
}
