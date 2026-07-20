import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { statusBadge, periodBadge } from '../components/badges.js';

export async function renderCleaningReports(container) {
  const bRes = await apiFetch('/api/branches?all=1');
  const branchOptions = (bRes.data?.data || []).map(b => ({ value: b.id, label: b.full_name }));
  const years = Array.from({ length: 4 }, (_, i) => String(new Date().getFullYear() - i));

  buildCrudPage({
    container,
    title: 'Laporan General Cleaning & Deep Cleaning',
    icon: '🧹',
    apiPath: '/api/reports/cleaning',
    itemLabel: 'Laporan GC/DC',
    columns: [
      { key: 'branch_name', label: 'Cabang' },
      { key: 'activity_type', label: 'Jenis', render: v => `<span class="badge ${v === 'Deep Cleaning' ? 'badge-purple' : 'badge-success'}">${v}</span>` },
      { key: 'period', label: 'Periode', render: v => periodBadge(v) },
      { key: 'activity_date', label: 'Tanggal', nowrap: true },
      { key: 'status', label: 'Status', render: v => statusBadge(v) },
      { key: 'document_link', label: 'Dokumen', render: v => v ? `<a href="${v}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">📄 Buka</a>` : '-' },
    ],
    filterFields: [
      { type: 'select', name: 'branch_id', label: 'Cabang', options: branchOptions },
      { type: 'select', name: 'activity_type', label: 'Jenis', options: ['General Cleaning', 'Deep Cleaning'] },
      { type: 'select', name: 'period', label: 'Periode', options: ['Q1', 'Q2', 'Q3', 'Q4'] },
      { type: 'select', name: 'status', label: 'Status', options: ['Pending', 'Done'] },
      { type: 'select', name: 'year', label: 'Tahun', options: years },
    ],
    formFields: (data) => [
      {
        type: 'row', fields: [
          { name: 'branch_id', label: 'Cabang', type: 'select', required: true, options: branchOptions, value: data?.branch_id },
          { name: 'activity_type', label: 'Jenis Kegiatan', type: 'select', required: true, options: ['General Cleaning', 'Deep Cleaning'], value: data?.activity_type },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'period', label: 'Periode', type: 'select', required: true, options: ['Q1', 'Q2', 'Q3', 'Q4'], value: data?.period },
          { name: 'activity_date', label: 'Tanggal', type: 'date', required: true, value: data?.activity_date },
        ]
      },
      { name: 'status', label: 'Status', type: 'select', required: true, options: ['Pending', 'Done'], value: data?.status || 'Pending' },
      { name: 'document_link', label: 'Link Dokumen', type: 'url', placeholder: 'https://drive.google.com/...', value: data?.document_link },
      { name: 'notes', label: 'Catatan', type: 'textarea', rows: 2, value: data?.notes },
    ],
  });
}
