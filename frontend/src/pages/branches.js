import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { downloadExcel } from '../utils/excel.js';
import { fetchExportRows } from '../utils/dataCache.js';
import { submitJournaledImport } from '../utils/import_client.js';

export async function renderBranches(container) {
  buildCrudPage({
    container,
    title: 'Manajemen Cabang',
    icon: '🏢',
    apiPath: '/api/branches',
    enableMobileFilterSheet: true,
    itemLabel: 'Cabang',
    bulkDelete: false, // Disabled to prevent D1 vs Google Sheets outbox sync mismatch
    columns: [
      { key: 'code', label: 'Kode', width: '60px' },
      { key: 'full_name', label: 'Nama Cabang' },
      { key: 'city', label: 'Kota' },
      {
        key: 'branch_scope',
        label: 'Lingkup',
        render: (_v, row) => {
          const c = String(row?.code || '').trim();
          if (c === '000') return '<span class="badge badge-info" title="Management / HQ (non-operasional, tersedia di dropdown modul)">HQ (Non-Ops)</span>';
          if (c === 'A01') return '<span class="badge badge-warning" title="Gabung area 002. Kemang (tidak dihitung terpisah di Dashboard)">Gabung 002. Kemang</span>';
          if (/^(00[1-9]|0[1-6][0-9]|070)$/.test(c)) return '<span class="badge badge-success">Operasional</span>';
          return '<span class="badge badge-neutral">Non-Master</span>';
        }
      },
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
      onExport: async (filters) => {
        const rows = await fetchExportRows('/api/branches/export', filters);
        downloadExcel(rows, 'Data_Cabang');
      },
      onTemplate: () => {
        const template = [
          { 'Kode Cabang': '001', 'Nama Pendek': 'Pondok Bambu', 'Nama Lengkap': '001. Pondok Bambu', 'Kota': 'Jakarta Timur' },
          { 'Kode Cabang': '002', 'Nama Pendek': 'Bintaro', 'Nama Lengkap': '002. Bintaro', 'Kota': 'Tangerang Selatan' }
        ];
        downloadExcel(template, 'Template_Import_Cabang');
      },
      onImport: async (json, importCtx = {}) => {
        const payload = json.map(row => ({
          code: String(row['Kode Cabang'] || '').trim(),
          name: String(row['Nama Pendek'] || '').trim(),
          full_name: String(row['Nama Lengkap'] || '').trim(),
          city: String(row['Kota'] || '').trim(),
        })).filter(row => row.code && row.name);
        
        return await submitJournaledImport({
          module: 'branches',
          label: 'Data Cabang',
          endpoint: '/api/branches/import',
          body: payload,
          idempotencyKey: importCtx?.idempotencyKey,
          recover: Boolean(importCtx?.recover),
        });
      }
    }
  });
}
