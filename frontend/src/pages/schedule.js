import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { statusBadge, activityTypeBadge, periodBadge } from '../components/badges.js';

let branchOptions = [];
let picOptions = [];

export async function renderSchedule(container) {
  const [bRes, pRes] = await Promise.all([
    apiFetch('/api/branches?all=1'),
    apiFetch('/api/pic'),
  ]);
  branchOptions = (bRes.data?.data || []).map(b => ({ value: b.id, label: b.full_name }));
  picOptions = (pRes.data?.data || []).map(p => p.name);

  buildCrudPage({
    container,
    title: 'Jadwal Kegiatan',
    icon: '🗓️',
    apiPath: '/api/schedule',
    itemLabel: 'Jadwal',
    columns: [
      { key: 'branch_name', label: 'Cabang' },
      { key: 'activity_type', label: 'Kegiatan', render: v => activityTypeBadge(v) },
      { key: 'period', label: 'Periode', render: v => periodBadge(v) },
      { key: 'pic', label: 'PIC' },
      { key: 'opening_date', label: 'Tgl Opening', nowrap: true },
      { key: 'target_date', label: 'Tgl Target', nowrap: true },
      { key: 'completion_date', label: 'Tgl Selesai', nowrap: true },
      { key: 'status', label: 'Status', render: v => statusBadge(v) },
    ],
    filterFields: [
      { type: 'select', name: 'branch_id', label: 'Cabang', options: branchOptions },
      { type: 'select', name: 'activity_type', label: 'Kegiatan', options: [
        'Inspeksi Hygiene & Aset Bangunan', 'General Cleaning', 'Deep Cleaning', 'Fogging'
      ]},
      { type: 'select', name: 'period', label: 'Periode', options: ['Q1', 'Q2', 'Q3', 'Q4'] },
      { type: 'select', name: 'status', label: 'Status', options: ['Pending', 'In Progress', 'Done'] },
      { type: 'select', name: 'pic', label: 'PIC', options: picOptions },
    ],
    formFields: (data) => [
      {
        type: 'row', fields: [
          { name: 'branch_id', label: 'Cabang', type: 'select', required: true, options: branchOptions, value: data?.branch_id },
          { name: 'activity_type', label: 'Jenis Kegiatan', type: 'select', required: true, options: [
            'Inspeksi Hygiene & Aset Bangunan', 'General Cleaning', 'Deep Cleaning', 'Fogging'
          ], value: data?.activity_type },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'period', label: 'Periode', type: 'select', required: true, options: ['Q1', 'Q2', 'Q3', 'Q4'], value: data?.period },
          { name: 'pic', label: 'PIC', type: 'select', options: picOptions, value: data?.pic },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'opening_date', label: 'Tanggal Opening', type: 'date', value: data?.opening_date },
          { name: 'target_date', label: 'Tanggal Target', type: 'date', value: data?.target_date },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'completion_date', label: 'Tanggal Selesai', type: 'date', value: data?.completion_date },
          { name: 'status', label: 'Status', type: 'select', required: true, options: ['Pending', 'In Progress', 'Done'], value: data?.status || 'Pending' },
        ]
      },
      { name: 'notes', label: 'Catatan', type: 'textarea', rows: 2, value: data?.notes },
    ],
  });
}
