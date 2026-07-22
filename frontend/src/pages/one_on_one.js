import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { statusBadge } from '../components/badges.js';

export async function renderOneOnOne(container) {
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
    title: 'One on One',
    icon: '🤝',
    apiPath: '/api/one-on-one',
    bulkDelete: true,
    itemLabel: 'One on One',
    columns: [
      { key: 'meeting_date', label: 'Tanggal', nowrap: true , render: v => window.formatDate(v) },
      { key: 'branch_name', label: 'Cabang' },
      { key: 'employee_name', label: 'Nama Karyawan' },
      { key: 'pic', label: 'PIC' },
      { key: 'problem', label: 'Masalah', render: v => `<span title="${v || ''}">${v?.length > 50 ? v.slice(0, 50) + '…' : (v || '-')}</span>` },
      { key: 'solution', label: 'Solusi', render: v => `<span title="${v || ''}">${v?.length > 40 ? v.slice(0, 40) + '…' : (v || '-')}</span>` },
      { key: 'status', label: 'Status', render: v => statusBadge(v) },
      { key: 'completion_date', label: 'Tgl Selesai', nowrap: true , render: v => window.formatDate(v) },
      { key: 'day_count', label: 'Hari' },
      { key: 'document_link', label: 'Dokumen', render: v => v ? `<a href="${v}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">📄 Buka</a>` : '-' },
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari nama / masalah...' },
      { type: 'select', name: 'branch_id', label: 'Cabang', options: branchOptions },
      { type: 'select', name: 'status', label: 'Status', options: ['Open', 'Done'] },
    ],
    exportOptions: {
      moduleName: 'one_on_one',
      onExport: async (filters) => {
        const qs = new URLSearchParams(filters || {}).toString();
        const res = await apiFetch(`/api/one-on-one?limit=10000&${qs}`);
        if (res.ok) {
          const data = res.data.data.map(d => ({
            'Tanggal': d.meeting_date || '',
            'Cabang': d.branch_name || '',
            'Nama Karyawan': d.employee_name || '',
            'PIC': d.pic || '',
            'Masalah': d.problem || '',
            'Solusi': d.solution || '',
            'Status': d.status || '',
            'Tgl Selesai': d.completion_date || '',
            'Dokumen': d.document_link || ''
          }));
          const { downloadExcel } = await import('../utils/excel.js');
          downloadExcel(data, `Data_One_on_One_${new Date().toISOString().slice(0,10)}`);
        } else throw new Error('Gagal mengambil data');
      },
      onTemplate: async () => {
        const template = [
          { 'Tanggal': '2026-01-08', 'Cabang': '001. Pondok Bambu', 'Nama Karyawan': 'Widya Astuti', 'PIC': 'Rina', 'Masalah': 'Terlambat terus', 'Solusi': 'Teguran', 'Status': 'Open', 'Tgl Selesai': '', 'Dokumen': 'https://link.doc' }
        ];
        const { downloadExcel } = await import('../utils/excel.js');
        downloadExcel(template, 'Template_Import_OneOnOne');
      },
      onImport: async (json) => {
        const matchBranch = (str) => {
          if (!str) return null;
          const s = str.toLowerCase();
          const b = bRes.data?.data.find(r => r.full_name.toLowerCase() === s || r.code.toLowerCase() === s || r.name.toLowerCase() === s);
          return b ? b.id : null;
        };
        const parseDate = (v) => {
          if (!v) return '';
          if (v instanceof Date && !isNaN(v.getTime())) return v.toISOString().slice(0, 10);
          const s = String(v).trim();
          if (/^\d{4,5}$/.test(s)) {
            const n = Number(s);
            if (n > 20000 && n < 99999) {
              const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
              return isNaN(d.getTime()) ? '' : d.toISOString().slice(0, 10);
            }
          }
          if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
          const parts = s.split(/[\/\-\.]/);
          if (parts.length === 3) {
            const [a, b, c] = parts.map(p => p.trim());
            if (a.length === 4 && b.length <= 2 && c.length <= 2) return `${a}-${b.padStart(2, '0')}-${c.padStart(2, '0')}`;
            if (c.length === 4 && b.length <= 2 && a.length <= 2) return `${c}-${b.padStart(2, '0')}-${a.padStart(2, '0')}`;
          }
          return s;
        };
        const payload = json.map(row => ({
          meeting_date: parseDate(row['Tanggal']),
          employee_name: String(row['Nama Karyawan'] || '').trim(),
          branch_id: matchBranch(String(row['Cabang'] || '').trim()),
          pic: String(row['PIC'] || '').trim(),
          problem: String(row['Masalah'] || '').trim(),
          solution: String(row['Solusi'] || '').trim(),
          status: String(row['Status'] || '').trim(),
          completion_date: parseDate(row['Tgl Selesai']),
          document_link: String(row['Dokumen'] || '').trim(),
        })).filter(r => r.meeting_date && r.employee_name && r.branch_id);
        
        const res = await apiFetch('/api/one-on-one/import', { method: 'POST', body: JSON.stringify(payload) });
        if (!res.ok) throw new Error(res.data?.error || 'Import gagal');
      }
    },
    formFields: (data) => [
      {
        type: 'row', fields: [
          { name: 'meeting_date', label: 'Tanggal', type: 'date', required: true, value: data?.meeting_date },
          { name: 'branch_id', label: 'Cabang', type: 'combobox', options: branchOptions, createApi: { path: '/api/branches', field: 'full_name' }, value: data?.branch_id },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'employee_name', label: 'Nama Karyawan', type: 'select', required: true, options: getEmpOptions(data?.employee_name), value: data?.employee_name },
          { name: 'pic', label: 'PIC', type: 'combobox', options: getPicOptions(data?.pic), createApi: { path: '/api/pic', field: 'name' }, value: data?.pic },
        ]
      },
      { name: 'problem', label: 'Masalah', type: 'textarea', required: true, rows: 3, value: data?.problem },
      { name: 'solution', label: 'Solusi', type: 'textarea', rows: 3, value: data?.solution },
      {
        type: 'row', fields: [
          { name: 'status', label: 'Status', type: 'select', required: true, options: ['Open', 'Done'], value: data?.status || '' },
          { name: 'completion_date', label: 'Tanggal Selesai', type: 'date', value: data?.completion_date },
        ]
      },
      { name: 'document_link', label: 'Link Dokumen', type: 'url', placeholder: 'https://drive.google.com/...', value: data?.document_link },
    ],
  });
}
