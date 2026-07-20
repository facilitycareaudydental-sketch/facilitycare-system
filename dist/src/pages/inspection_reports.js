import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { statusBadge, periodBadge } from '../components/badges.js';

export async function renderInspectionReports(container) {
  const bRes = await apiFetch('/api/branches?all=1');
  const branchOptions = (bRes.data?.data || []).map(b => ({ value: b.id, label: b.full_name }));
  const years = Array.from({ length: 4 }, (_, i) => String(new Date().getFullYear() - i));

  buildCrudPage({
    container,
    title: 'Laporan Inspeksi Hygiene',
    icon: '🔍',
    apiPath: '/api/reports/inspection',
    itemLabel: 'Laporan Inspeksi',
    columns: [
      { key: 'branch_name', label: 'Cabang' },
      { key: 'period', label: 'Periode', render: v => periodBadge(v) },
      { key: 'inspection_date', label: 'Tanggal', nowrap: true },
      { key: 'fc_score', label: 'Point FC', render: v => v !== null && v !== undefined ? `<strong class="${v >= 80 ? 'text-success' : v >= 60 ? 'text-warning' : 'text-danger'}">${v}</strong>` : '-' },
      { key: 'spv_score', label: 'Point SPV', render: v => v !== null && v !== undefined ? `<strong>${v}</strong>` : '-' },
      { key: 'status', label: 'Status', render: v => statusBadge(v) },
      { key: 'document_link', label: 'Dokumen', render: v => v ? `<a href="${v}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">📄 Buka</a>` : '-' },
    ],
    filterFields: [
      { type: 'select', name: 'branch_id', label: 'Cabang', options: branchOptions },
      { type: 'select', name: 'period', label: 'Periode', options: ['Q1', 'Q2', 'Q3', 'Q4'] },
      { type: 'select', name: 'status', label: 'Status', options: ['Pending', 'Done'] },
      { type: 'select', name: 'year', label: 'Tahun', options: years },
    ],
    formFields: (data) => [
      {
        type: 'row', fields: [
          { name: 'branch_id', label: 'Cabang', type: 'select', required: true, options: branchOptions, value: data?.branch_id },
          { name: 'period', label: 'Periode', type: 'select', required: true, options: ['Q1', 'Q2', 'Q3', 'Q4'], value: data?.period },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'inspection_date', label: 'Tanggal Inspeksi', type: 'date', required: true, value: data?.inspection_date },
          { name: 'status', label: 'Status', type: 'select', required: true, options: ['Pending', 'Done'], value: data?.status || 'Pending' },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'fc_score', label: 'Point FC', type: 'number', step: '0.1', min: '0', max: '100', value: data?.fc_score },
          { name: 'spv_score', label: 'Point SPV', type: 'number', step: '0.1', min: '0', max: '100', value: data?.spv_score },
        ]
      },
      { name: 'document_link', label: 'Link Dokumen', type: 'url', placeholder: 'https://drive.google.com/...', value: data?.document_link },
      { name: 'notes', label: 'Catatan', type: 'textarea', rows: 2, value: data?.notes },
    ],
  });
}
