import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { getCachedBranches, getCachedEmployeeNames } from '../utils/dataCache.js';
import { statusBadge } from '../components/badges.js';
import { downloadExcel } from '../utils/excel.js';

let branchOptions = [];
let employeeOptions = [];

function calculateIssuesDay(e) {
  if (e.target.name === 'report_date' || e.target.name === 'completion_date') {
    const reportDateInput = document.querySelector('input[name="report_date"]');
    const compDateInput = document.querySelector('input[name="completion_date"]');
    const dayInput = document.querySelector('input[name="day_count"]');
    
    if (reportDateInput && compDateInput && dayInput) {
      if (reportDateInput.value && compDateInput.value) {
        const rd = new Date(reportDateInput.value);
        const cd = new Date(compDateInput.value);
        const diff = Math.floor((cd - rd) / 86400000);
        dayInput.value = !isNaN(diff) ? diff : '';
      } else {
        dayInput.value = '';
      }
    }
  }
}
document.body.removeEventListener('input', calculateIssuesDay);
document.body.addEventListener('input', calculateIssuesDay);
document.body.removeEventListener('change', calculateIssuesDay);
document.body.addEventListener('change', calculateIssuesDay);

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
  const years = ['2025', '2026', '2027', '2028', '2029', '2030'];
  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  const monthOptions = monthNames.map((name, i) => {
      const monthNum = String(i + 1).padStart(2, '0');
      return { value: `${currentYear}-${monthNum}`, label: `${name} ${currentYear}` };
  });

  buildCrudPage({
    container,
    title: 'Permasalahan',
    icon: '⚠️',
    apiPath: '/api/issues',
    bulkDelete: true,
    enableMobileFilterSheet: true,
    itemLabel: 'Permasalahan',
    paginationMode: 'client',
    onDataLoaded: (items) => {
      if (dashFilter) {
        return items.filter(s => filterDashboardItem(s, dashFilter));
      }
      return items;
    },
    columns: [
      { key: 'report_date', label: 'Tanggal Info', nowrap: true , render: v => window.formatDate(v) },
      { key: 'branch_name', label: 'Cabang' },
      { key: 'category', label: 'Kategori', render: v => `<span class="badge badge-secondary">${v}</span>` },
      { key: 'source', label: 'Sumber Laporan' },
      { key: 'complaint', label: 'Keluhan', render: v => `<span title="${v}">${v?.length > 50 ? v.slice(0, 50) + '...' : v}</span>` },
      { key: 'employee_name', label: 'Nama FC' },
      { key: 'fc_specialist', label: 'FC Spesialis' },
      { key: 'solution', label: 'Solusi', render: v => `<span title="${v || ''}">${v?.length > 40 ? v.slice(0, 40) + '...' : (v || '-')}</span>` },
      { key: 'status', label: 'Status', render: v => statusBadge(v) },
      { key: 'completion_date', label: 'Tanggal Selesai', nowrap: true , render: v => window.formatDate(v) },
      { key: 'day_count', label: 'Day', render: v => v !== null && v !== undefined ? v : '-' },
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari keluhan / nama FC...' },
      { type: 'select', name: 'branch_id', label: 'Cabang', options: branchOptions },
      { type: 'select', name: 'month', label: 'Bulan', options: monthOptions },
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
          { name: 'day_count', label: 'Day', type: 'number', value: data?.day_count, readonly: true },
        ]
      },
    ],
    exportOptions: {
      moduleName: 'issues',
      onExport: async () => {
        const res = await apiFetch(`/api/issues${window.location.search ? window.location.search + '&' : '?'}limit=10000`);
        if (res.ok) {
          const data = res.data.data.map(d => ({
            'Tanggal Info': d.report_date || '',
            'Cabang': d.branch_name || '',
            'Kategori': d.category || '',
            'Sumber Laporan': d.source || '',
            'Keluhan': d.complaint || '',
            'Nama FC': d.employee_name || '',
            'FC Spesialis': d.fc_specialist || '',
            'Solusi': d.solution || '',
            'Status': d.status || '',
            'Tanggal Selesai': d.completion_date || '',
            'Day': d.day_count !== null ? d.day_count : ''
          }));
          downloadExcel(data, 'Data_Permasalahan');
        } else throw new Error('Gagal mengambil data');
      },
      onTemplate: () => {
        const template = [
          { 'Tanggal Info': '2024-03-01', 'Cabang': '001. Pondok Bambu', 'Kategori': 'Cleaning', 'Sumber Laporan': 'SPV', 'Keluhan': 'Lantai kotor', 'Nama FC': 'Budi Santoso', 'FC Spesialis': 'Fajar', 'Solusi': 'Teguran lisan', 'Status': 'Done', 'Tanggal Selesai': '2024-03-02', 'Day': 1 }
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
          report_date: String(row['Tanggal Info'] || row['Tanggal'] || '').trim(),
          category: String(row['Kategori'] || '').trim(),
          source: String(row['Sumber Laporan'] || row['Sumber'] || '').trim(),
          complaint: String(row['Keluhan'] || '').trim(),
          employee_name: String(row['Nama FC'] || '').trim(),
          fc_specialist: String(row['FC Spesialis'] || '').trim(),
          solution: String(row['Solusi'] || '').trim(),
          completion_date: String(row['Tanggal Selesai'] || row['Tgl Selesai'] || '').trim(),
          status: String(row['Status'] || '').trim(),
          day_count: row['Day'] || row['Hari'] || null
        })).filter(row => row.report_date && row.complaint && row.category);
        
        const res = await apiFetch('/api/import/issues', {
          method: 'POST',
          body: JSON.stringify({ rows: payload, onDuplicate: 'update' })
        });
        if (!res.ok) throw new Error(res.data?.error || 'Import gagal');
        return res.data;
      }
    }
  });
}
