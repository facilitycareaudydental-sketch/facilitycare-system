import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { statusBadge } from '../components/badges.js';

export async function renderBasecampReports(container) {
  const [bRes, eRes, pRes] = await Promise.all([
    apiFetch('/api/branches?all=1'),
    apiFetch('/api/employees?limit=10000'),
    apiFetch('/api/pic')
  ]);
  const branchOptions = (bRes.data?.data || []).map(b => ({ value: b.id, label: b.full_name }));
  
  const employeeOptions = (eRes.data?.data || []).map(e => ({ value: e.full_name, label: e.full_name }));
  const rawPicOptions = (pRes.data?.data || []).filter(p => p.role === 'FC Spesialis').map(p => ({ value: p.name, label: p.name }));

  const getEmpOptions = (val) => {
    if (val && !employeeOptions.find(o => o.value === val)) {
      return [...employeeOptions, { value: val, label: val }];
    }
    return employeeOptions;
  };
  
  const getPicOptions = (val) => {
    if (val && !rawPicOptions.find(o => o.value === val)) {
      return [...rawPicOptions, { value: val, label: val }];
    }
    return rawPicOptions;
  };

  buildCrudPage({
    container,
    title: 'Rekap Laporan Basecamp',
    icon: '📝',
    apiPath: '/api/reports/basecamp',
    itemLabel: 'Laporan Basecamp',
    columns: [
      { key: 'info_date', label: 'Tgl Info', nowrap: true , render: v => window.formatDate(v) },
      { key: 'branch_name', label: 'Cabang' },
      { key: 'problem', label: 'Permasalahan', render: v => `<span title="${v || ''}">${v?.length > 60 ? v.slice(0, 60) + '…' : (v || '-')}</span>` },
      { key: 'pic', label: 'PIC' },
      { key: 'done_date', label: 'Tgl Done', nowrap: true , render: v => window.formatDate(v) },
      { key: 'status', label: 'Status', render: v => statusBadge(v) },
      { key: 'notes', label: 'Keterangan', render: v => v?.length > 40 ? v.slice(0, 40) + '…' : (v || '-') },
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari permasalahan / PIC...' },
      { type: 'select', name: 'branch_id', label: 'Cabang', options: branchOptions },
      { type: 'select', name: 'status', label: 'Status', options: ['Open', 'In Progress', 'Done'] },
    ],
    formFields: (data) => [
      {
        type: 'row', fields: [
          { name: 'branch_id', label: 'Cabang', type: 'select', required: true, options: branchOptions, value: data?.branch_id },
          { name: 'pic', label: 'PIC', type: 'select', options: getPicOptions(data?.pic), value: data?.pic },
        ]
      },
      { name: 'problem', label: 'Permasalahan', type: 'textarea', required: true, rows: 3, value: data?.problem },
      {
        type: 'row', fields: [
          { name: 'info_date', label: 'Tanggal Info', type: 'date', required: true, value: data?.info_date },
          { name: 'done_date', label: 'Tanggal Done', type: 'date', value: data?.done_date },
        ]
      },
      { name: 'status', label: 'Status', type: 'select', required: true, options: ['Open', 'In Progress', 'Done'], value: data?.status || '' },
      { name: 'notes', label: 'Keterangan', type: 'textarea', rows: 2, value: data?.notes },
    ],
  });
}
