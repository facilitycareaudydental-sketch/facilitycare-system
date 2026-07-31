import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { statusBadge, periodBadge } from '../components/badges.js';
import { downloadExcel } from '../utils/excel.js';

export async function renderRelievers(container) {
  const [bRes, eRes] = await Promise.all([
    apiFetch('/api/branches?all=1'),
    apiFetch('/api/employees?limit=10000')
  ]);
  const branchOptions = (bRes.data?.data || []).map(b => ({ value: b.id, label: b.full_name }));
  const employeeOptions = (eRes.data?.data || []).map(e => ({ value: e.full_name, label: e.full_name }));

  const getEmpOptions = (val) => {
    if (val && !employeeOptions.find(o => o.value === val)) {
      return [...employeeOptions, { value: val, label: val }];
    }
    return employeeOptions;
  };

  const relieverOptions = [
    'Krishna Aryaan Permana',
    'Agung Septiadi',
    'Indra Saputro',
    'Wariskin',
    'Iqbal'
  ];

  const getRelieverOptions = (val) => {
    const opts = relieverOptions.map(name => ({ value: name, label: name }));
    if (val && !opts.find(o => o.value === val)) {
      return [...opts, { value: val, label: val }];
    }
    return opts;
  };

  buildCrudPage({
    container,
    title: 'Jadwal Reliefer',
    icon: '🔄',
    apiPath: '/api/relievers',
    bulkDelete: true,
    itemLabel: 'Reliefer',
    columns: [
      { key: 'backup_date', label: 'Tanggal Backup', nowrap: true , render: v => window.formatDate(v) },
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
          { name: 'original_fc_name', label: 'FC yang Digantikan', type: 'select', options: [{value:'', label:'BELUM ADA FC'}, ...getEmpOptions(data?.original_fc_name)], value: data?.original_fc_name },
          { name: 'reliever_name', label: 'Nama Reliefer', type: 'select', required: true, options: getRelieverOptions(data?.reliever_name), value: data?.reliever_name },
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
      { name: 'status', label: 'Status', type: 'select', required: true, options: ['Pending', 'Done', 'Tidak Datang'], value: data?.status || '' },
    ],
    exportOptions: {
      moduleName: 'relievers',
      onExport: async () => {
        const res = await apiFetch('/api/relievers?limit=10000');
        if (res.ok) {
          const data = res.data.data.map(d => ({
            'Tanggal Backup': d.backup_date || '',
            'Cabang': d.branch_name || '',
            'FC Digantikan': d.original_fc_name || '',
            'Periode': d.period || '',
            'Reliefer': d.reliever_name || '',
            'Keterangan': d.reason || '',
            'Shift': d.shift || '',
            'Tanggal Selesai': d.completion_date || '',
            'Status': d.status || ''
          }));
          downloadExcel(data, 'Data_Reliefer');
        } else throw new Error('Gagal mengambil data');
      },
      onTemplate: () => {
        const template = [
          { 'Tanggal Backup': '2024-03-10', 'Cabang': '001. Pondok Bambu', 'FC Digantikan': 'Budi Santoso', 'Periode': 'Q1', 'Reliefer': 'Andi', 'Keterangan': 'Sakit', 'Shift': 'Pagi', 'Tanggal Selesai': '2024-03-10', 'Status': 'Done' }
        ];
        downloadExcel(template, 'Template_Import_Reliefer');
      },
      onImport: async (json) => {
        const bRes = await apiFetch('/api/branches?all=1');
        const rawBranches = bRes.data?.data || [];
        
        const matchBranch = (str) => {
          if (!str) return null;
          const s = str.toLowerCase();
          const b = rawBranches.find(r => r.full_name.toLowerCase() === s || r.code.toLowerCase() === s || r.name.toLowerCase() === s);
          return b ? b.id : null;
        };

        const payload = json.map(row => ({
          branch_id: matchBranch(String(row['Cabang'] || '').trim()),
          backup_date: String(row['Tanggal Backup'] || '').trim(),
          original_fc_name: String(row['FC Digantikan'] || '').trim(),
          reliever_name: String(row['Reliefer'] || '').trim(),
          period: String(row['Periode'] || '').trim(),
          reason: String(row['Keterangan'] || '').trim(),
          shift: String(row['Shift'] || '').trim(),
          completion_date: String(row['Tanggal Selesai'] || '').trim(),
          status: String(row['Status'] || '').trim(),
        })).filter(row => row.reliever_name && row.backup_date);
        
        const res = await apiFetch('/api/relievers/import', {
          method: 'POST',
          body: JSON.stringify(payload)
        });
        if (!res.ok) throw new Error(res.data?.error || 'Import gagal');
      }
    }
  });
}
