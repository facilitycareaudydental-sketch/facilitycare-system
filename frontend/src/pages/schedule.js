import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { statusBadge, activityTypeBadge, periodBadge } from '../components/badges.js';
import { downloadExcel } from '../utils/excel.js';

let branchOptions = [];
let picOptions = [];

export async function renderSchedule(container) {
  const [bRes, eRes, pRes] = await Promise.all([
    apiFetch('/api/branches?all=1'),
    apiFetch('/api/employees?limit=10000'),
    apiFetch('/api/pic')
  ]);
  branchOptions = (bRes.data?.data || []).map(b => ({ value: b.id, label: b.full_name }));
  
  const employeeOptions = (eRes.data?.data || []).map(e => ({ value: e.full_name, label: e.full_name }));
  const rawPicOptions = (pRes.data?.data || []).filter(p => p.role === 'FC Spesialis').map(p => ({ value: p.name, label: p.name }));
  
  picOptions = [...rawPicOptions];

  const getEmpOptions = (val) => {
    if (val && !employeeOptions.find(o => o.value === val)) {
      return [...employeeOptions, { value: val, label: val }];
    }
    return employeeOptions;
  };

  buildCrudPage({
    container,
    title: 'Jadwal Kegiatan',
    icon: '🗓️',
    apiPath:,
    bulkDelete: true,
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
    exportOptions: {
      moduleName: 'schedule',
      onExport: async () => {
        const res = await apiFetch('/api/schedule?limit=10000');
        if (res.ok) {
          const data = res.data.data.map(d => ({
            'Cabang': d.branch_name || '',
            'Kegiatan': d.activity_type || '',
            'Periode': d.period || '',
            'PIC': d.pic || '',
            'Tgl Opening': d.opening_date || '',
            'Tgl Target': d.target_date || '',
            'Tgl Selesai': d.completion_date || '',
            'Status': d.status || '',
            'Catatan': d.notes || ''
          }));
          downloadExcel(data, 'Data_Jadwal_Kegiatan');
        } else throw new Error('Gagal mengambil data');
      },
      onTemplate: () => {
        const template = [
          { 'Cabang': '001. Pondok Bambu', 'Kegiatan': 'General Cleaning', 'Periode': 'Q1', 'PIC': 'Fajar', 'Tgl Opening': '2024-02-01', 'Tgl Target': '2024-02-15', 'Tgl Selesai': '2024-02-14', 'Status': 'Done', 'Catatan': '' }
        ];
        downloadExcel(template, 'Template_Import_Jadwal');
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
          activity_type: String(row['Kegiatan'] || '').trim(),
          period: String(row['Periode'] || '').trim(),
          pic: String(row['PIC'] || '').trim(),
          opening_date: String(row['Tgl Opening'] || '').trim(),
          target_date: String(row['Tgl Target'] || '').trim(),
          completion_date: String(row['Tgl Selesai'] || '').trim(),
          status: String(row['Status'] || '').trim() || 'Pending',
          notes: String(row['Catatan'] || '').trim(),
        })).filter(row => row.activity_type && row.period);
        
        const res = await apiFetch('/api/schedule/import', {
          method: 'POST',
          body: JSON.stringify(payload)
        });
        if (!res.ok) throw new Error(res.data?.error || 'Import gagal');
      }
    }
  });
}
