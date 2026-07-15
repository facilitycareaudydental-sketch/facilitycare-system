import { buildCrudPage } from './_crud.js';
import { statusBadge } from '../components/badges.js';
import { getUser } from '../config.js';

export async function renderUsers(container) {
  const currentUser = getUser();
  if (!currentUser || !['superadmin', 'admin'].includes(currentUser.role)) {
    container.innerHTML = '<div class="empty-state"><p class="text-danger">Akses ditolak.</p></div>';
    return;
  }

  buildCrudPage({
    container,
    title: 'Manajemen User',
    icon: '🔐',
    apiPath: '/api/users',
    itemLabel: 'User',
    columns: [
      { key: 'full_name', label: 'Nama Lengkap' },
      { key: 'username', label: 'Username' },
      { key: 'email', label: 'Email' },
      { key: 'role', label: 'Role', render: v => {
        const colors = { superadmin: 'badge-danger', admin: 'badge-purple', manager: 'badge-info', spv: 'badge-secondary', viewer: 'badge-neutral' };
        return `<span class="badge ${colors[v] || 'badge-neutral'}">${v}</span>`;
      }},
      { key: 'is_active', label: 'Status', render: v => v ? '<span class="badge badge-success">Aktif</span>' : '<span class="badge badge-neutral">Nonaktif</span>' },
      { key: 'created_at', label: 'Dibuat', nowrap: true, render: v => v ? new Date(v).toLocaleDateString('id-ID') : '-' },
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari nama / username...' },
    ],
    formFields: (data) => {
      const isEdit = !!data;
      return [
        {
          type: 'row', fields: [
            { name: 'full_name', label: 'Nama Lengkap', required: true, placeholder: 'Nama lengkap', value: data?.full_name },
            { name: 'username', label: 'Username', required: !isEdit, placeholder: 'username', value: data?.username },
          ]
        },
        {
          type: 'row', fields: [
            { name: 'email', label: 'Email', type: 'email', required: !isEdit, placeholder: 'email@contoh.com', value: data?.email },
            { name: 'role', label: 'Role', type: 'select', required: true, options: [
              { value: 'superadmin', label: 'Super Admin' },
              { value: 'admin', label: 'Admin' },
              { value: 'manager', label: 'Manager' },
              { value: 'spv', label: 'Supervisor' },
              { value: 'viewer', label: 'Viewer' },
            ], value: data?.role || 'viewer' },
          ]
        },
        {
          type: 'row', fields: [
            { name: 'password', label: isEdit ? 'Password Baru (kosongkan jika tidak diubah)' : 'Password', type: 'password', required: !isEdit, placeholder: 'Min. 6 karakter' },
            { name: 'is_active', label: 'Status Aktif', type: 'checkbox', checkLabel: 'User aktif', value: isEdit ? data?.is_active : 1 },
          ]
        },
      ];
    },
  });
}
