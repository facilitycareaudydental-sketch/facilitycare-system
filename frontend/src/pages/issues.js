import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { statusBadge } from '../components/badges.js';

let branchOptions = [];

export async function renderIssues(container) {
  const [bRes] = await Promise.all([apiFetch('/api/branches?all=1')]);
  branchOptions = (bRes.data?.data || []).map(b => ({ value: b.id, label: b.full_name }));

  const currentYear = new Date().getFullYear();
  const years = Array.from({ length: 5 }, (_, i) => String(currentYear - i));

  buildCrudPage({
    container,
    title: 'Permasalahan',
    icon: '⚠️',
    apiPath: '/api/issues',
    itemLabel: 'Permasalahan',
    columns: [
      { key: 'report_date', label: 'Tanggal', nowrap: true },
      { key: 'branch_name', label: 'Cabang' },
      { key: 'category', label: 'Kategori', render: v => `<span class="badge badge-secondary">${v}</span>` },
      { key: 'source', label: 'Sumber' },
      { key: 'complaint', label: 'Keluhan', render: v => `<span title="${v}">${v?.length > 50 ? v.slice(0, 50) + '…' : v}</span>` },
      { key: 'employee_name', label: 'Nama FC' },
      { key: 'fc_specialist', label: 'FC Spesialis' },
      { key: 'solution', label: 'Solusi', render: v => `<span title="${v || ''}">${v?.length > 40 ? v.slice(0, 40) + '…' : (v || '-')}</span>` },
      { key: 'status', label: 'Status', render: v => statusBadge(v) },
      { key: 'completion_date', label: 'Tgl Selesai', nowrap: true },
      { key: 'day_count', label: 'Hari', render: v => v !== null && v !== undefined ? v : '-' },
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari keluhan / nama FC...' },
      { type: 'select', name: 'branch_id', label: 'Cabang', options: branchOptions },
      { type: 'select', name: 'category', label: 'Kategori', options: ['SDM', 'Cleaning', 'Aset', 'K3', 'Lainnya'] },
      { type: 'select', name: 'status', label: 'Status', options: ['Open', 'In Progress', 'Done'] },
      { type: 'select', name: 'year', label: 'Tahun', options: years },
    ],
    formFields: (data) => [
      {
        type: 'row', fields: [
          { name: 'report_date', label: 'Tanggal Info', type: 'date', required: true, value: data?.report_date },
          { name: 'branch_id', label: 'Cabang', type: 'select', required: true, options: branchOptions, value: data?.branch_id },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'category', label: 'Kategori', type: 'select', required: true, options: ['SDM', 'Cleaning', 'Aset', 'K3', 'Lainnya'], value: data?.category },
          { name: 'source', label: 'Sumber Laporan', type: 'select', options: ['SPV', 'AM', 'Berlin', 'Ade', 'Miswar', 'Pattrel', 'Perawat', 'FC', 'RCP', 'Lainnya'], value: data?.source },
        ]
      },
      { name: 'complaint', label: 'Keluhan', type: 'textarea', required: true, rows: 3, value: data?.complaint },
      {
        type: 'row', fields: [
          { name: 'employee_name', label: 'Nama FC / Security', placeholder: 'Nama yang bermasalah', value: data?.employee_name },
          { name: 'fc_specialist', label: 'FC Spesialis', type: 'select', options: ['Fajar', 'Miswar', 'Ade', 'Berlin', 'Pattrel', 'Lainnya'], value: data?.fc_specialist },
        ]
      },
      { name: 'solution', label: 'Solusi / Tindakan', type: 'textarea', rows: 3, value: data?.solution },
      {
        type: 'row', fields: [
          { name: 'status', label: 'Status', type: 'select', required: true, options: ['Open', 'In Progress', 'Done'], value: data?.status || 'Open' },
          { name: 'completion_date', label: 'Tanggal Selesai', type: 'date', value: data?.completion_date },
        ]
      },
    ],
  });
}
