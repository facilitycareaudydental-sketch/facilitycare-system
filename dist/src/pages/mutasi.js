import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';

let branchOptions = [];

export async function renderMutasi(container) {
  const res = await apiFetch('/api/branches?all=1');
  branchOptions = (res.data?.data || []).map(b => ({ value: b.id, label: b.full_name }));

  buildCrudPage({
    container,
    title: 'Data Mutasi',
    icon: '🔄',
    apiPath: '/api/mutasi',
    itemLabel: 'Mutasi',
    bulkDelete: true,
    columns: [
      { key: 'tanggal', label: 'Tanggal', render: v => v ? new Date(v).toLocaleDateString('id-ID', { year: 'numeric', month: 'short', day: 'numeric' }) : '-' },
      { key: 'employee_name', label: 'Nama Karyawan' },
      { key: 'from_branch_name', label: 'Cabang Asal' },
      { key: 'to_branch_name', label: 'Cabang Tujuan' },
      { key: 'status', label: 'Status', render: v => `<span class="badge ${v === 'Selesai' ? 'badge-success' : 'badge-warning'}">${v || '-'}</span>` },
      { key: 'document_link', label: 'Dokumen', render: v => v ? `<a href="${v}" target="_blank" class="text-primary hover-underline">Lihat</a>` : '-' }
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari nama karyawan...' },
      { type: 'select', name: 'from_branch_id', label: 'Cabang Asal', options: branchOptions },
      { type: 'select', name: 'to_branch_id', label: 'Cabang Tujuan', options: branchOptions },
    ],
    formFields: [
      { type: 'date', name: 'tanggal', label: 'Tanggal', required: true },
      { type: 'text', name: 'employee_name', label: 'Nama Karyawan', required: true },
      { type: 'select', name: 'from_branch_id', label: 'Cabang Asal', required: true, options: branchOptions },
      { type: 'select', name: 'to_branch_id', label: 'Cabang Tujuan', required: true, options: branchOptions },
      { type: 'select', name: 'status', label: 'Status', required: true, options: ['Proses', 'Selesai'] },
      { type: 'url', name: 'document_link', label: 'Link Dokumen (Opsional)' }
    ]
  });
}
