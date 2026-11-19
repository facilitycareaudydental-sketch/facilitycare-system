import { buildCrudPage } from './_crud.js';
import { statusBadge } from '../components/badges.js';
import { getUser, apiFetch, ASSIGNABLE_MODULES, getAssignedModules } from '../config.js';
import { downloadExcel } from '../utils/excel.js';
import { fetchExportRows } from '../utils/dataCache.js';
import { toastInfo } from '../components/toast.js';
import { submitJournaledImport } from '../utils/import_client.js';
import { escapeHtml } from '../utils/sanitize.js';

export async function renderUsers(container) {
  const currentUser = getUser();
  if (!currentUser || currentUser.role !== 'superadmin') {
    container.innerHTML = '<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';
    return;
  }

  buildCrudPage({
    container,
    title: 'Manajemen User',
    icon: '🔐',
    apiPath: '/api/users',
    bulkDelete: true,
    enableMobileFilterSheet: true,
    itemLabel: 'User',
    columns: [
      { key: 'full_name', label: 'Nama Lengkap' },
      { key: 'username', label: 'Username' },
      { key: 'email', label: 'Email' },
      { key: 'role', label: 'Role', render: v => {
        const colors = { superadmin: 'badge-danger', fc_specialist_leader: 'badge-info' };
        return `<span class="badge ${colors[v] || 'badge-neutral'}">${escapeHtml(v)}</span>`;
      }},
      { key: 'is_active', label: 'Status', render: v => v ? '<span class="badge badge-success">Aktif</span>' : '<span class="badge badge-neutral">Nonaktif</span>' },
      { key: 'created_at', label: 'Dibuat', nowrap: true, render: v => v ? new Date(v).toLocaleDateString('id-ID') : '-' },
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari nama / username...' },
    ],
    formFields: (data) => {
      const isEdit = !!data;
      const currentRole = String(data?.role || '').trim().toLowerCase();
      const SUPERADMIN_OPTION = {
        value: 'superadmin',
        label: 'Superadmin (Akses Penuh - Maks. 2 Aktif)'
      };
      const FC_LEADER_OPTION = {
        value: 'fc_specialist_leader',
        label: 'FC Specialist / Leader (View All + Input & Edit Lembur & Issue - Maks. 2 Aktif)'
      };
      const roleOptions = isEdit && currentRole === 'superadmin'
        ? [SUPERADMIN_OPTION, FC_LEADER_OPTION]
        : [FC_LEADER_OPTION, SUPERADMIN_OPTION];
      const defaultRoleValue = isEdit && (currentRole === 'superadmin' || currentRole === 'fc_specialist_leader')
        ? currentRole
        : 'fc_specialist_leader';
      return [
        {
          type: 'row', fields: [
            { name: 'full_name', label: 'Nama Lengkap', required: true, placeholder: 'Nama lengkap', value: data?.full_name },
            { name: 'username', label: 'Username', required: true, placeholder: 'username', value: data?.username, autocomplete: 'off' },
          ]
        },
        {
          type: 'row', fields: [
            { name: 'email', label: 'Email', type: 'email', required: true, placeholder: 'email@contoh.com', value: data?.email, autocomplete: 'off' },
            { name: 'role', label: 'Role Resmi', type: 'select', required: true, options: roleOptions, value: defaultRoleValue },
          ]
        },
        {
          type: 'row', fields: [
            { name: 'password', label: isEdit ? 'Password Baru (kosongkan jika tidak diubah)' : 'Password', type: 'password', required: !isEdit, placeholder: 'Min. 12 karakter', autocomplete: 'new-password' },
            { name: 'is_active', label: 'Status Aktif', type: 'checkbox', checkLabel: 'User aktif', value: isEdit ? data?.is_active : 1 },
          ]
        },
        {
          type: 'html',
          html: `
            <div style="margin-top:8px; padding:10px 12px; background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px; font-size:0.8rem; color:#1E3A8A;">
              <strong>Kebijakan RBAC &amp; Keamanan Password:</strong> Role resmi FCMS hanya <code>superadmin</code> (maks. 2 aktif) dan <code>fc_specialist_leader</code> (maks. 2 aktif). Password wajib minimal 12 karakter (kombinasi huruf besar, huruf kecil, angka, dan simbol).
            </div>
          `
        },
      ];
    },
    onBeforeSubmit: (body, data) => {
      if (body.full_name) body.full_name = String(body.full_name).trim();
      if (body.username) body.username = String(body.username).trim();
      if (body.email) body.email = String(body.email).trim().toLowerCase();
      // On edit, if password is blank, remove it so password isn't overwritten
      if (data && (!body.password || String(body.password).trim() === '')) {
        delete body.password;
      }
      body.assigned_modules = null;
      return body;
    },
    exportOptions: {
      moduleName: 'users',
      onExport: async (filters) => {
        const rows = await fetchExportRows('/api/users/export', filters);
        const data = rows.map(d => ({
          'Nama Lengkap': d.full_name || '',
          'Username': d.username || '',
          'Email': d.email || '',
          'Role': d.role || '',
          'Status': d.is_active ? 'Aktif' : 'Nonaktif'
        }));
        downloadExcel(data, 'Data_Users');
      },
      onTemplate: () => {
        const template = [
          {
            'Nama Lengkap': 'FC Leader 01',
            'Username': 'fcleader01',
            'Email': 'fcleader@contoh.com',
            'Role': 'fc_specialist_leader',
            'Password': ''
          }
        ];
        downloadExcel(template, 'Template_Import_Users');
        if (typeof toastInfo === 'function') {
          toastInfo(
            'Kolom Password pada template dikosongkan. Wajib isi password unik dan kuat (min. 12 karakter: huruf besar, kecil, angka, simbol) untuk setiap akun baru, atau kosongkan jika memperbarui akun existing tanpa mengubah password.'
          );
        }
      },
      onImport: async (json, importCtx = {}) => {
        const payload = json.map(row => ({
          full_name: String(row['Nama Lengkap'] || '').trim(),
          username: String(row['Username'] || '').trim(),
          email: String(row['Email'] || '').trim().toLowerCase(),
          role: String(row['Role'] || '').trim() || 'fc_specialist_leader',
          password: String(row['Password'] || '').trim(),
        })).filter(row => row.username || row.email || row.full_name || row.password);
        
        return await submitJournaledImport({
          module: 'users',
          label: 'Manajemen Pengguna',
          endpoint: '/api/users/import',
          body: payload,
          idempotencyKey: importCtx?.idempotencyKey,
          recover: Boolean(importCtx?.recover),
        });
      }
    }
  });
}
