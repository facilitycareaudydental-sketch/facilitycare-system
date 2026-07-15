import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { statusBadge, divisionBadge } from '../components/badges.js';

let branchOptions = [];

async function loadBranches() {
  const res = await apiFetch('/api/branches?all=1');
  branchOptions = (res.data?.data || []).map(b => ({ value: b.id, label: b.full_name }));
}

export async function renderEmployees(container) {
  await loadBranches();

  buildCrudPage({
    container,
    title: 'Karyawan',
    icon: '👥',
    apiPath: '/api/employees',
    itemLabel: 'Karyawan',
    columns: [
      { key: 'full_name', label: 'Nama Lengkap' },
      { key: 'branch_name', label: 'Cabang' },
      { key: 'division', label: 'Divisi', render: (v) => divisionBadge(v) },
      { key: 'phone', label: 'No. HP', render: v => v ? `<a href="tel:${v}">${v}</a>` : '-' },
      { key: 'join_date', label: 'Tgl Masuk' },
      { key: 'status', label: 'Status', render: v => statusBadge(v) },
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari nama karyawan...' },
      { type: 'select', name: 'branch_id', label: 'Cabang', options: branchOptions },
      { type: 'select', name: 'division', label: 'Divisi', options: ['FACILITY CARE', 'SECURITY'] },
      { type: 'select', name: 'status', label: 'Status', options: ['Aktif', 'Tidak Aktif', 'Resign', 'Cut'] },
    ],
    formFields: (data) => [
      {
        type: 'row', fields: [
          { name: 'full_name', label: 'Nama Lengkap', required: true, placeholder: 'Nama lengkap karyawan', value: data?.full_name },
          { name: 'phone', label: 'No. HP', placeholder: '08xx-xxxx-xxxx', value: data?.phone },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'branch_id', label: 'Cabang', type: 'select', options: branchOptions, value: data?.branch_id },
          { name: 'division', label: 'Divisi', type: 'select', required: true, options: ['FACILITY CARE', 'SECURITY'], value: data?.division || 'FACILITY CARE' },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'join_date', label: 'Tanggal Masuk', type: 'date', value: data?.join_date },
          { name: 'status', label: 'Status', type: 'select', required: true, options: ['Aktif', 'Tidak Aktif', 'Resign', 'Cut'], value: data?.status || 'Aktif' },
        ]
      },
      { name: 'notes', label: 'Catatan', type: 'textarea', rows: 2, value: data?.notes },
    ],
  });
}
