import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { statusBadge } from '../components/badges.js';
import { downloadExcel } from '../utils/excel.js';

export async function renderBasecampReports(container) {
  const [bRes, eRes, pRes] = await Promise.all([
    apiFetch('/api/branches?all=1'),
    apiFetch('/api/employees?limit=10000'),
    apiFetch('/api/pic?limit=10000')
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
    bulkDelete: true,
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
          { name: 'branch_id', label: 'Cabang', type: 'combobox', required: true, options: branchOptions, createApi: { path: '/api/branches', field: 'full_name' }, value: data?.branch_id },
          { name: 'pic', label: 'PIC', type: 'combobox', options: getPicOptions(data?.pic), createApi: { path: '/api/pic', field: 'name' }, value: data?.pic },
        ]
      },
      { name: 'problem', label: 'Permasalahan', type: 'textarea', required: true, rows: 3, value: data?.problem },
      {
        type: 'row', fields: [
          { name: 'info_date', label: 'Tanggal Info', type: 'date', required: true, value: data?.info_date },
          { name: 'done_date', label: 'Tanggal Done', type: 'date', value: data?.done_date },
        ]
      },
      { name: 'status', label: 'Status', type: 'select', options: ['Open', 'In Progress', 'Done'], value: data?.status || 'Open' },
      { name: 'notes', label: 'Keterangan / Tindak Lanjut', type: 'textarea', rows: 2, value: data?.notes },
    ],
    exportOptions: {
      moduleName: 'basecamp_reports',
      onExport: async (filters) => {
        const qs = new URLSearchParams(filters || {}).toString();
        const res = await apiFetch(`/api/reports/basecamp?limit=10000&${qs}`);
        if (res.ok) {
          const data = res.data.data.map(d => ({
            'Tgl Info': d.info_date || '',
            'Cabang': d.branch_name || '',
            'Permasalahan': d.problem || '',
            'PIC': d.pic || '',
            'Tgl Done': d.done_date || '',
            'Status': d.status || '',
            'Keterangan': d.notes || ''
          }));
          downloadExcel(data, `Rekap_Laporan_Basecamp_${new Date().toISOString().slice(0,10)}`);
        } else throw new Error('Gagal mengambil data');
      },
      onTemplate: () => {
        const template = [
          { 'Tgl Info': '2026-01-08', 'Cabang': '001. Pondok Bambu', 'Permasalahan': 'Request fogging karena banyak nyamuk', 'PIC': 'Fajar', 'Tgl Done': '2026-01-10', 'Status': 'Done', 'Keterangan': 'Sudah difogging' }
        ];
        downloadExcel(template, 'Template_Import_Basecamp');
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
          info_date: parseDate(row['Tgl Info'] || row['Tanggal Info']),
          branch_id: matchBranch(String(row['Cabang'] || '').trim()),
          problem: String(row['Permasalahan'] || '').trim(),
          pic: String(row['PIC'] || '').trim(),
          done_date: parseDate(row['Tgl Done'] || row['Tanggal Done']),
          status: String(row['Status'] || '').trim(),
          notes: String(row['Keterangan'] || row['Catatan'] || '').trim(),
        })).filter(row => row.info_date && row.branch_id && row.problem);
        
        const res = await apiFetch('/api/reports/basecamp/import', {
          method: 'POST',
          body: JSON.stringify(payload)
        });
        if (!res.ok) throw new Error(res.data?.error || 'Import gagal');
      }
    }
  });
}
