import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';

let branchOptions = [];

export async function renderSP(container) {
  const res = await apiFetch('/api/branches?all=1');
  branchOptions = (res.data?.data || []).map(b => ({ value: b.id, label: b.name }));

  buildCrudPage({
    container,
    title: 'Surat Peringatan',
    icon: '✉️',
    apiPath: '/api/sp',
    itemLabel: 'SP',
    bulkDelete: true,
    columns: [
      { key: 'tanggal', label: 'Tanggal', render: v => v ? new Date(v).toLocaleDateString('id-ID', { year: 'numeric', month: 'short', day: 'numeric' }) : '-' },
      { key: 'employee_name', label: 'Nama Karyawan' },
      { key: 'branch_name', label: 'Cabang' },
      { key: 'sp_type', label: 'Jenis SP', render: v => `<span class="badge badge-warning">${v}</span>` },
      { key: 'status', label: 'Status', render: v => `<span class="badge ${v === 'Aktif' ? 'badge-danger' : 'badge-success'}">${v}</span>` },
      { key: 'document_link', label: 'Dokumen', render: v => v ? `<a href="${v}" target="_blank" class="text-primary hover-underline">Lihat</a>` : '-' }
    ],
    formFields: [
      { type: 'date', name: 'tanggal', label: 'Tanggal', required: true },
      { type: 'text', name: 'employee_name', label: 'Nama Karyawan', required: true },
      { type: 'select', name: 'branch_id', label: 'Cabang', required: true, options: branchOptions },
      { type: 'select', name: 'sp_type', label: 'Jenis Surat Peringatan', required: true, options: ['SP 1', 'SP 2', 'SP 3', 'Teguran Lisan'] },
      { type: 'select', name: 'status', label: 'Status', required: true, options: ['Aktif', 'Selesai'] },
      { type: 'url', name: 'document_link', label: 'Link Dokumen' }
    ]
  });
}
