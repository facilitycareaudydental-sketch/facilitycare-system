import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';

let branchOptions = [];

export async function renderMutasi(container) {
  const res = await apiFetch('/api/branches?all=1');
  branchOptions = (res.data?.data || []).map(b => ({ value: b.id, label: b.name }));

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
      { key: 'from_branch_name', label: 'Awal Cabang' },
      { key: 'to_branch_name', label: 'Rotasi Cabang' },
      { key: 'status', label: 'Status', render: v => `<span class="badge ${v === 'Selesai' ? 'badge-success' : 'badge-warning'}">${v}</span>` },
      { key: 'document_link', label: 'Dokumen', render: v => v ? `<a href="${v}" target="_blank" class="text-primary hover-underline">Lihat</a>` : '-' }
    ],
    formFields: [
      { type: 'date', name: 'tanggal', label: 'Tanggal', required: true },
      { type: 'text', name: 'employee_name', label: 'Nama Karyawan', required: true },
      { type: 'select', name: 'from_branch_id', label: 'Awal Cabang', required: true, options: branchOptions },
      { type: 'select', name: 'to_branch_id', label: 'Rotasi Cabang', required: true, options: branchOptions },
      { type: 'select', name: 'status', label: 'Status', required: true, options: ['Proses', 'Selesai'] },
      { type: 'url', name: 'document_link', label: 'Link Dokumen' }
    ]
  });
}
