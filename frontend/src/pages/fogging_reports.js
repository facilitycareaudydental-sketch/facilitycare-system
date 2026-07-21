import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { statusBadge, periodBadge } from '../components/badges.js';

export async function renderFoggingReports(container) {
  const bRes = await apiFetch('/api/branches?all=1');
  const branchOptions = (bRes.data?.data || []).map(b => ({ value: b.id, label: b.full_name }));
  const years = Array.from({ length: 4 }, (_, i) => String(new Date().getFullYear() - i));

  buildCrudPage({
    container,
    title: 'Rekap Fogging',
    icon: '💨',
    apiPath: '/api/reports/fogging',
    itemLabel: 'Fogging',
    columns: [
      { key: 'branch_name', label: 'Cabang' },
      { key: 'activity_type', label: 'Jenis', render: v => `<span class="badge badge-warning">${v}</span>` },
      { key: 'period', label: 'Periode', render: v => periodBadge(v) },
      { key: 'activity_date', label: 'Tanggal', nowrap: true , render: v => window.formatDate(v) },
      { key: 'status', label: 'Status', render: v => statusBadge(v) },
      { key: 'document_link', label: 'Dokumen', render: v => v ? `<a href="${v}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">📄 Buka</a>` : '-' },
      { key: 'notes', label: 'Catatan', render: v => v || '-' },
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
          { name: 'activity_date', label: 'Tanggal', type: 'date', value: data?.activity_date },
          { name: 'status', label: 'Status', type: 'select', required: true, options: ['Pending', 'Done'], value: data?.status || '' },
        ]
      },
      { name: 'document_link', label: 'Link Dokumen', type: 'url', placeholder: 'https://...', value: data?.document_link },
      { name: 'notes', label: 'Catatan', type: 'textarea', rows: 2, value: data?.notes },
    ],
  });
}
