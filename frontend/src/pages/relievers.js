import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { statusBadge, periodBadge } from '../components/badges.js';

export async function renderRelievers(container) {
  const bRes = await apiFetch('/api/branches?all=1');
  const branchOptions = (bRes.data?.data || []).map(b => ({ value: b.id, label: b.full_name }));

  buildCrudPage({
    container,
    title: 'Jadwal Reliefer',
    icon: '🔄',
    apiPath: '/api/relievers',
    itemLabel: 'Reliefer',
    columns: [
      { key: 'backup_date', label: 'Tanggal Backup', nowrap: true },
      { key: 'branch_name', label: 'Cabang' },
      { key: 'original_fc_name', label: 'FC Digantikan' },
      { key: 'period', label: 'Periode', render: v => periodBadge(v) },
      { key: 'reliever_name', label: 'Reliefer' },
      { key: 'reason', label: 'Keterangan' },
      { key: 'shift', label: 'Shift', render: v => v ? `<span class="badge badge-info">${v}</span>` : '-' },
      { key: 'status', label: 'Status', render: v => statusBadge(v) },
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari reliefer / FC...' },
      { type: 'select', name: 'branch_id', label: 'Cabang', options: branchOptions },
      { type: 'select', name: 'period', label: 'Periode', options: ['Q1', 'Q2', 'Q3', 'Q4'] },
      { type: 'select', name: 'status', label: 'Status', options: ['Pending', 'Done', 'Tidak Datang'] },
    ],
    formFields: (data) => [
      {
        type: 'row', fields: [
          { name: 'branch_id', label: 'Cabang', type: 'select', required: true, options: branchOptions, value: data?.branch_id },
          { name: 'period', label: 'Periode', type: 'select', options: ['Q1', 'Q2', 'Q3', 'Q4'], value: data?.period },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'original_fc_name', label: 'FC yang Digantikan', placeholder: 'Nama FC / BELUM ADA FC', value: data?.original_fc_name },
          { name: 'reliever_name', label: 'Nama Reliefer', required: true, placeholder: 'Nama yang menggantikan', value: data?.reliever_name },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'backup_date', label: 'Tanggal Backup', type: 'date', required: true, value: data?.backup_date },
          { name: 'completion_date', label: 'Tanggal Selesai', type: 'date', value: data?.completion_date },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'reason', label: 'Keterangan', type: 'select', options: ['Cuti', 'Mengisi Kekosongan', 'Back Up Training', 'Deep Cleaning', 'Training Praktek Skill', 'Sakit', 'Lainnya'], value: data?.reason },
          { name: 'shift', label: 'Shift', type: 'select', options: ['Pagi', 'Siang', 'Full Shift', 'Middle'], value: data?.shift },
        ]
      },
      { name: 'status', label: 'Status', type: 'select', required: true, options: ['Pending', 'Done', 'Tidak Datang'], value: data?.status || 'Pending' },
    ],
  });
}
