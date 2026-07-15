import { buildCrudPage } from './_crud.js';

export async function renderBranches(container) {
  buildCrudPage({
    container,
    title: 'Manajemen Cabang',
    icon: '🏢',
    apiPath: '/api/branches',
    itemLabel: 'Cabang',
    columns: [
      { key: 'code', label: 'Kode', width: '60px' },
      { key: 'full_name', label: 'Nama Cabang' },
      { key: 'city', label: 'Kota' },
      { key: 'is_active', label: 'Status', render: v => v ? '<span class="badge badge-success">Aktif</span>' : '<span class="badge badge-neutral">Nonaktif</span>' },
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari nama / kode cabang...' },
    ],
    formFields: (data) => [
      {
        type: 'row', fields: [
          { name: 'code', label: 'Kode Cabang', required: true, placeholder: '001, A01, ...', value: data?.code },
          { name: 'name', label: 'Nama Pendek', required: true, placeholder: 'Pondok Bambu', value: data?.name },
        ]
      },
      { name: 'full_name', label: 'Nama Lengkap', required: true, placeholder: '001. Pondok Bambu', value: data?.full_name },
      {
        type: 'row', fields: [
          { name: 'city', label: 'Kota', placeholder: 'Jakarta', value: data?.city },
          { name: 'is_active', label: 'Status', type: 'checkbox', checkLabel: 'Cabang aktif', value: data?.is_active !== undefined ? data.is_active : 1 },
        ]
      },
    ],
  });
}
