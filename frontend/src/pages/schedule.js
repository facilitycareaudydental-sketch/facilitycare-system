import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { getCachedBranches, getCachedEmployeeNames } from '../utils/dataCache.js';
import { statusBadge, activityTypeBadge, periodBadge } from '../components/badges.js';
import { downloadExcel } from '../utils/excel.js';

let branchOptions = [];
let picOptions = [];

export function getActivePeriod(data) {
  if (!Array.isArray(data)) return 'Q3';
  const periods = ['Q4', 'Q3', 'Q2', 'Q1'];
  for (const p of periods) {
    if (data.some(d => d.period === p)) return p;
  }
  return 'Q3'; // Default fallback
}

export function filterDashboardItem(s, type) {
  if (s.period !== 'Q3') return false;
  const status = String(s.status || '').toLowerCase();
  if (status !== 'selesai' && status !== 'completed' && status !== 'done') return false;
  
  const title = String(s.activity_type || '').toLowerCase();
  if (type === 'inspeksi') return title.includes('inspeksi');
  if (type === 'gcdc') return title.includes('general cleaning') || title.includes('deep cleaning');
  return false;
}

export async function renderSchedule(container, params) {
  branchOptions = await getCachedBranches();
  const allEmployees = await getCachedEmployeeNames();
  picOptions = ['Ade', 'Berlin'];

  const getPicOptions = (val) => {
    if (val && !picOptions.find(o => (typeof o === 'object' ? o.value : o) === val)) {
      return [...picOptions, val];
    }
    return picOptions;
  };
  
  const schedRes = await apiFetch(`/api/schedule${window.location.search ? window.location.search + '&' : '?'}limit=10000`);

  const formatDate = (d) => {
    if (!d || d === '-' || String(d).trim() === '') return '';
    const p = String(d).split('-');
    if (p.length === 3 && p[0].length === 4) return `${p[2]}-${p[1]}-${p[0]}`;
    return d;
  };

  const scheduleData = schedRes.data?.data || [];
  const activePeriod = getActivePeriod(scheduleData);
  
  const dashFilter = params ? params.get('dash_filter') : null;

  buildCrudPage({
    container,
    title: 'Jadwal Kegiatan',
    icon: '📅',
    apiPath: '/api/schedule',
    bulkDelete: true,
    itemLabel: 'Jadwal',
    paginationMode: 'client',
    defaultFilters: dashFilter ? { period: 'Q3' } : { period: activePeriod },
    onDataLoaded: (items) => {
      // 1. Filter dataset directly if dash_filter is set
      if (dashFilter) {
        items = items.filter(s => filterDashboardItem(s, dashFilter));
      }

      // 2. Sort descending by opening_date, so newest year/date is first
      return items.sort((a, b) => {
        const da = a.opening_date ? new Date(a.opening_date).getTime() : 0;
        const db = b.opening_date ? new Date(b.opening_date).getTime() : 0;
        return db - da;
      });
    },
    columns: [
      { key: 'branch_name', label: 'Cabang' },
      { key: 'activity_type', label: 'Kegiatan', render: v => activityTypeBadge(v) },
      { key: 'period', label: 'Periode', render: v => periodBadge(v) },
      { key: 'pic', label: 'PIC' },
      { key: 'opening_date', label: 'Tgl Opening', nowrap: true, render: v => formatDate(v) },
      { key: 'target_date', label: 'Tgl Target', nowrap: true, render: v => formatDate(v) },
      { key: 'completion_date', label: 'Tgl Selesai', nowrap: true, render: v => formatDate(v) },
      { key: 'status', label: 'Status', render: v => statusBadge(v) },
    ],
    filterFields: [
      { type: 'combobox', name: 'branch_id', label: 'Cabang', options: branchOptions },
      { type: 'select', name: 'activity_type', label: 'Kegiatan', options: [
        'Inspeksi Hygiene & Aset Bangunan', 'General Cleaning', 'Deep Cleaning', 'Fogging'
      ]},
      { type: 'select', name: 'period', label: 'Periode', options: ['Q1', 'Q2', 'Q3', 'Q4'] },
      { type: 'select', name: 'status', label: 'Status', options: ['Pending', 'In Progress', 'Done'] },
      { type: 'combobox', name: 'pic', label: 'PIC', options: picOptions },
    ],
    formFields: (data) => [
      {
        type: 'row', fields: [
          { name: 'branch_id', label: 'Cabang', type: 'combobox', required: true, options: branchOptions, value: data?.branch_id },
          { name: 'activity_type', label: 'Jenis Kegiatan', type: 'select', required: true, options: [
            'Inspeksi Hygiene & Aset Bangunan', 'General Cleaning', 'Deep Cleaning', 'Fogging'
          ], value: data?.activity_type },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'period', label: 'Periode', type: 'select', required: true, options: ['Q1', 'Q2', 'Q3', 'Q4'], value: data?.period },
          { name: 'pic', label: 'PIC', type: 'combobox', options: getPicOptions(data?.pic), value: data?.pic },
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
          { name: 'status', label: 'Status', type: 'select', required: true, options: ['Pending', 'In Progress', 'Done'], value: data?.status || '' },
        ]
      },
      { name: 'notes', label: 'Catatan', type: 'textarea', rows: 2, value: data?.notes },
    ],
    exportOptions: {
      moduleName: 'schedule',
      onExport: async () => {
        const res = await apiFetch(`/api/schedule${window.location.search ? window.location.search + '&' : '?'}limit=10000`);
        if (res.ok) {
          const data = res.data.data.map(d => ({
            'Cabang': d.branch_name || '',
            'Kegiatan': d.activity_type || '',
            'Periode': d.period || '',
            'PIC': d.pic || '',
            'Tgl Opening': d.opening_date || '',
            'Tgl Target': d.target_date || '',
            'Tgl Selesai': d.completion_date || '',
            'Status': d.status || ''
          }));
          downloadExcel(data, 'Data_Jadwal_Kegiatan');
        } else throw new Error('Gagal mengambil data');
      },
      onTemplate: () => {
        const template = [
          { 'Cabang': '001. Pondok Bambu', 'Kegiatan': 'General Cleaning', 'Periode': 'Q1', 'PIC': 'Fajar', 'Tgl Opening': '2024-02-01', 'Tgl Target': '2024-02-15', 'Tgl Selesai': '2024-02-14', 'Status': 'Done' }
        ];
        downloadExcel(template, 'Template_Import_Jadwal');
      },
      onImport: async (json) => {
        const bRes = await apiFetch('/api/branches?all=1');
        const rawBranches = bRes.data?.data || [];
        
        const matchBranch = (str) => {
          if (!str) return null;
          const s = String(str || '').toLowerCase();
          const b = rawBranches.find(r => String(r.full_name || '').toLowerCase() === s || String(r.code || '').toLowerCase() === s || String(r.name || '').toLowerCase() === s);
          return b ? b.id : null;
        };
        
        const parseDate = (v) => {
          if (v === undefined || v === null || v === '') return '';
          if (v instanceof Date && !isNaN(v.getTime())) return v.toISOString().slice(0, 10);
          const s = String(v).trim();
          if (s === '' || s === '0') return '';
          if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
          if (/^\d{4,5}$/.test(s)) {
            const n = Number(s);
            if (n > 20000 && n < 99999) {
              const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
              return isNaN(d.getTime()) ? '' : d.toISOString().slice(0, 10);
            }
          }
          const parts = s.split(/[\/\-\.]/);
          if (parts.length === 3) {
            const [a, b, c] = parts.map(p => p.trim());
            if (a.length === 4 && b.length <= 2 && c.length <= 2) return `${a}-${b.padStart(2, '0')}-${c.padStart(2, '0')}`;
            if (c.length === 4 && b.length <= 2 && a.length <= 2) return `${c}-${b.padStart(2, '0')}-${a.padStart(2, '0')}`;
          }
          return s; // Fallback
        };

        const payload = json.map(row => ({
          branch_id: matchBranch(String(row['Cabang'] || '').trim()),
          activity_type: String(row['Kegiatan'] || '').trim(),
          period: String(row['Periode'] || '').trim(),
          pic: String(row['PIC'] || row['Pic'] || '').trim(),
          opening_date: parseDate(row['Tgl Opening'] || row['Tanggal Opening'] || row['Tgl Openir']),
          target_date: parseDate(row['Tgl Target'] || row['Tanggal Target']),
          completion_date: parseDate(row['Tgl Selesai'] || row['Tanggal Selesai']),
          status: String(row['Status'] || '').trim(),
          notes: String(row['Catatan'] || row['Keterangan'] || '').trim(),
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
