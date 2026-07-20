import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { statusBadge } from '../components/badges.js';

export async function renderOneOnOne(container) {
  const [bRes, eRes, pRes] = await Promise.all([
    apiFetch('/api/branches?all=1'),
    apiFetch('/api/employees?limit=10000'),
    apiFetch('/api/pic')
  ]);
  const branchOptions = (bRes.data?.data || []).map(b => ({ value: b.id, label: b.full_name }));
  
  const employeeOptions = (eRes.data?.data || []).map(e => ({ value: e.full_name, label: e.full_name }));
  const rawPicOptions = (pRes.data?.data || []).map(p => ({ value: p.name, label: p.name }));

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
    title: 'One on One',
    icon: '🤝',
    apiPath: '/api/one-on-one',
    itemLabel: 'One on One',
    columns: [
      { key: 'meeting_date', label: 'Tanggal', nowrap: true },
      { key: 'branch_name', label: 'Cabang' },
      { key: 'employee_name', label: 'Nama Karyawan' },
      { key: 'pic', label: 'PIC' },
      { key: 'problem', label: 'Masalah', render: v => `<span title="${v || ''}">${v?.length > 50 ? v.slice(0, 50) + '…' : (v || '-')}</span>` },
      { key: 'solution', label: 'Solusi', render: v => `<span title="${v || ''}">${v?.length > 40 ? v.slice(0, 40) + '…' : (v || '-')}</span>` },
      { key: 'status', label: 'Status', render: v => statusBadge(v) },
      { key: 'completion_date', label: 'Tgl Selesai', nowrap: true },
      { key: 'day_count', label: 'Hari' },
      { key: 'document_link', label: 'Dokumen', render: v => v ? `<a href="${v}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">📄 Buka</a>` : '-' },
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari nama / masalah...' },
      { type: 'select', name: 'branch_id', label: 'Cabang', options: branchOptions },
      { type: 'select', name: 'status', label: 'Status', options: ['Open', 'Done'] },
    ],
    formFields: (data) => [
      {
        type: 'row', fields: [
          { name: 'meeting_date', label: 'Tanggal', type: 'date', required: true, value: data?.meeting_date },
          { name: 'branch_id', label: 'Cabang', type: 'select', options: branchOptions, value: data?.branch_id },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'employee_name', label: 'Nama Karyawan', type: 'select', required: true, options: getEmpOptions(data?.employee_name), value: data?.employee_name },
          { name: 'pic', label: 'PIC', type: 'select', options: ['SPV', 'AM', ...getPicOptions(data?.pic)], value: data?.pic },
        ]
      },
      { name: 'problem', label: 'Masalah', type: 'textarea', required: true, rows: 3, value: data?.problem },
      { name: 'solution', label: 'Solusi', type: 'textarea', rows: 3, value: data?.solution },
      {
        type: 'row', fields: [
          { name: 'status', label: 'Status', type: 'select', required: true, options: ['Open', 'Done'], value: data?.status || 'Open' },
          { name: 'completion_date', label: 'Tanggal Selesai', type: 'date', value: data?.completion_date },
        ]
      },
      { name: 'document_link', label: 'Link Dokumen', type: 'url', placeholder: 'https://drive.google.com/...', value: data?.document_link },
    ],
  });
}
