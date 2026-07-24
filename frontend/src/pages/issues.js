import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { getCachedBranches, getCachedEmployeeNames } from '../utils/dataCache.js';
import { statusBadge } from '../components/badges.js';
import { downloadExcel } from '../utils/excel.js';

let branchOptions = [];
let employeeOptions = [];

export function filterDashboardItem(s, type) {
  const status = String(s.status || '').toLowerCase();
  if (type === 'open') return status === 'open';
  return false;
}

export async function renderIssues(container, params) {
  const dashFilter = params ? params.get('dash_filter') : null;
  branchOptions = await getCachedBranches();
  employeeOptions = await getCachedEmployeeNames();

  // Helper to ensure existing value is in options (prevents blank selects on old data)
  const getEmpOptions = (val) => {
    if (val && !employeeOptions.find(o => o.value === val)) {
      return [...employeeOptions, { value: val, label: val }];
    }
    return employeeOptions;
  };

  const currentYear = new Date().getFullYear();
  const years = Array.from({ length: 5 }, (_, i) => String(currentYear - i));

  buildCrudPage({
    container,
    title: 'Permasalahan',
    icon: '⚠️',
    apiPath: '/api/issues',
    bulkDelete: true,
    itemLabel: 'Permasalahan',
    paginationMode: 'client',
    onDataLoaded: (items) => {
      if (dashFilter) {
        return items.filter(s => filterDashboardItem(s, dashFilter));
      }
      return items;
    },
    columns: [
      { key: 'report_date', label: 'Tanggal', nowrap: true , render: v => window.formatDate(v) },
      { key: 'branch_name', label: 'Cabang' },
      { key: 'category', label: 'Kategori', render: v => `<span class="badge badge-secondary">${v}</span>` },
      { key: 'source', label: 'Sumber' },
      { key: 'complaint', label: 'Keluhan', render: v => `<span title="${v}">${v?.length > 50 ? v.slice(0, 50) + '…' : v}</span>` },
      { key: 'employee_name', label: 'Nama FC' },
      { key: 'fc_specialist', label: 'FC Spesialis' },
      { key: 'solution', label: 'Solusi', render: v => `<span title="${v || ''}">${v?.length > 40 ? v.slice(0, 40) + '…' : (v || '-')}</span>` },
      { key: 'status', label: 'Status', render: v => statusBadge(v) },
      { key: 'completion_date', label: 'Tgl Selesai', nowrap: true , render: v => window.formatDate(v) },
      { key: 'day_count', label: 'Hari', render: v => v !== null && v !== undefined ? v : '-' },
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari keluhan / nama FC...' },
      { type: 'combobox', name: 'branch_id', label: 'Cabang', options: branchOptions },
      { type: 'select', name: 'category', label: 'Kategori', options: ['SDM', 'Cleaning', 'Aset', 'K3', 'Lainnya'] },
      { type: 'select', name: 'status', label: 'Status', options: ['Open', 'In Progress', 'Done'] },
      { type: 'select', name: 'year', label: 'Tahun', options: years },
    ],
    formFields: (data) => [
      {
        type: 'row', fields: [
          { name: 'report_date', label: 'Tanggal Info', type: 'date', required: true, value: data?.report_date },
          { name: 'branch_id', label: 'Cabang', type: 'combobox', required: true, options: branchOptions, value: data?.branch_id },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'category', label: 'Kategori', type: 'select', required: true, options: ['SDM', 'Cleaning', 'Aset', 'K3', 'Lainnya'], value: data?.category },
          { name: 'source', label: 'Sumber Laporan', type: 'combobox', options: ['SPV', 'AM', 'RCP', 'Perawat', 'FC', 'Berlin', 'Ade', 'Pattrel', 'Dentrel'], value: data?.source },
        ]
      },
      { name: 'complaint', label: 'Keluhan', type: 'textarea', required: true, rows: 3, value: data?.complaint },
      {
        type: 'row', fields: [
          { name: 'employee_name', label: 'Nama FC / Security', type: 'combobox', options: getEmpOptions(data?.employee_name), value: data?.employee_name },
          { name: 'fc_specialist', label: 'FC Spesialis', type: 'combobox', options: getEmpOptions(data?.fc_specialist), value: data?.fc_specialist },
        ]
      },
      { name: 'solution', label: 'Solusi / Tindakan', type: 'textarea', rows: 3, value: data?.solution },
      {
        type: 'row', fields: [
          { name: 'status', label: 'Status', type: 'select', required: true, options: ['Open', 'In Progress', 'Done'], value: data?.status || '' },
          { name: 'completion_date', label: 'Tanggal Selesai', type: 'date', value: data?.completion_date },
        ]
      },
    ],
    exportOptions: {
      moduleName: 'issues',
      onExport: async () => {
        const res = await apiFetch(`/api/issues${window.location.search ? window.location.search + '&' : '?'}limit=10000`);
        if (res.ok) {
          const data = res.data.data.map(d => ({
            'Tanggal': d.report_date || '',
            'Cabang': d.branch_name || '',
            'Kategori': d.category || '',
            'Sumber': d.source || '',
            'Keluhan': d.complaint || '',
            'Nama FC': d.employee_name || '',
            'FC Spesialis': d.fc_specialist || '',
            'Solusi': d.solution || '',
            'Tgl Selesai': d.completion_date || '',
            'Status': d.status || ''
          }));
          downloadExcel(data, 'Data_Permasalahan');
        } else throw new Error('Gagal mengambil data');
      },
      onTemplate: () => {
        const template = [
          { 'Tanggal': '2024-03-01', 'Cabang': '001. Pondok Bambu', 'Kategori': 'Cleaning', 'Sumber': 'SPV', 'Keluhan': 'Lantai kotor', 'Nama FC': 'Budi Santoso', 'FC Spesialis': 'Fajar', 'Solusi': 'Teguran lisan', 'Tgl Selesai': '2024-03-02', 'Status': 'Done' }
        ];
        downloadExcel(template, 'Template_Import_Permasalahan');
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

        const payload = json.map(row => ({
          branch_id: matchBranch(String(row['Cabang'] || '').trim()),
          report_date: String(row['Tanggal'] || '').trim(),
          category: String(row['Kategori'] || '').trim(),
          source: String(row['Sumber'] || '').trim(),
          complaint: String(row['Keluhan'] || '').trim(),
          employee_name: String(row['Nama FC'] || '').trim(),
          fc_specialist: String(row['FC Spesialis'] || '').trim(),
          solution: String(row['Solusi'] || '').trim(),
          completion_date: String(row['Tgl Selesai'] || '').trim(),
          status: String(row['Status'] || '').trim(),
        })).filter(row => row.report_date && row.complaint && row.category);
        
        const res = await apiFetch('/api/issues/import', {
          method: 'POST',
          body: JSON.stringify(payload)
        });
        if (!res.ok) throw new Error(res.data?.error || 'Import gagal');
      }
    }
  });
}
