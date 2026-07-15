import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { downloadExcel } from '../utils/excel.js';

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
    exportOptions: {
      moduleName: 'branches',
      onExport: async () => {
        const res = await apiFetch('/api/branches?limit=10000');
        if (res.ok) downloadExcel(res.data.data, 'Data_Cabang');
        else throw new Error('Gagal mengambil data');
      },
      onTemplate: () => {
        const template = [
          { 'Kode Cabang': '001', 'Nama Pendek': 'Pondok Bambu', 'Nama Lengkap': '001. Pondok Bambu', 'Kota': 'Jakarta Timur' },
          { 'Kode Cabang': '002', 'Nama Pendek': 'Bintaro', 'Nama Lengkap': '002. Bintaro', 'Kota': 'Tangerang Selatan' }
        ];
        downloadExcel(template, 'Template_Import_Cabang');
      },
      onImport: async (json) => {
        const payload = json.map(row => ({
          code: String(row['Kode Cabang'] || '').trim(),
          name: String(row['Nama Pendek'] || '').trim(),
          full_name: String(row['Nama Lengkap'] || '').trim(),
          city: String(row['Kota'] || '').trim(),
        })).filter(row => row.code && row.name);
        
        const res = await apiFetch('/api/branches/import', {
          method: 'POST',
          body: JSON.stringify(payload)
        });
        if (!res.ok) throw new Error(res.data?.error || 'Import gagal');
      }
    }
  });
}
