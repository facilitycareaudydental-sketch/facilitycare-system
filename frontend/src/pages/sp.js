import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { getCachedBranches, getCachedEmployeeNames } from '../utils/dataCache.js';

let branchOptions = [];
let employeeOptions = [];

export async function renderSP(container) {
  branchOptions = await getCachedBranches();
  employeeOptions = await getCachedEmployeeNames();

  buildCrudPage({
    container,
    title: 'Data SP (Surat Peringatan)',
    icon: '✉️',
    apiPath: '/api/sp',
    itemLabel: 'SP',
    bulkDelete: true,
    columns: [
      { key: 'employee_name', label: 'Nama Karyawan' },
      { key: 'division', label: 'Divisi', render: (v) => v ? `<span class="badge badge-info">${v}</span>` : '-' },
      { key: 'branch_name', label: 'Cabang' },
      { key: 'tanggal', label: 'Tanggal Sp', render: v => v ? new Date(v).toLocaleDateString('id-ID', { year: 'numeric', month: 'short', day: 'numeric' }) : '-' },
      { key: 'akhir_sp', label: 'Akhir Sp', render: v => v ? new Date(v).toLocaleDateString('id-ID', { year: 'numeric', month: 'short', day: 'numeric' }) : '-' },
      { key: 'sp_type', label: 'Jenis Sp', render: v => `<span class="badge badge-warning">${v || '-'}</span>` },
      { key: 'document_link', label: 'Link Document / Foto', render: v => v ? `<a href="${v}" target="_blank" class="text-primary hover-underline">Lihat</a>` : '-' }
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari nama karyawan...' },
      { type: 'combobox', name: 'branch_id', label: 'Cabang', options: branchOptions },
    ],
    exportOptions: {
      moduleName: 'sp_data',
      onExport: async (filters) => {
        const qs = new URLSearchParams(filters || {}).toString();
        const res = await apiFetch(`/api/sp?limit=10000&${qs}`);
        if (res.ok) {
          const data = res.data.data.map(d => ({
            'Nama Karyawan': d.employee_name || '',
            'Divisi': d.division || '',
            'Cabang': d.branch_name || '',
            'Tanggal Sp': d.tanggal || '',
            'Akhir Sp': d.akhir_sp || '',
            'Jenis Sp': d.sp_type || '',
            'Link Document / Foto': d.document_link || ''
          }));
          const { downloadExcel } = await import('../utils/excel.js');
          downloadExcel(data, `Data_SP_${new Date().toISOString().slice(0,10)}`);
        } else throw new Error('Gagal mengambil data');
      },
      onTemplate: async () => {
        const template = [
          { 'Nama Karyawan': 'Budi Santoso', 'Divisi': 'FACILITY CARE', 'Cabang': '001. Pondok Bambu', 'Tanggal Sp': '2026-01-08', 'Akhir Sp': '2026-07-08', 'Jenis Sp': 'SP 1', 'Link Document / Foto': 'https://link.doc' }
        ];
        const { downloadExcel } = await import('../utils/excel.js');
        downloadExcel(template, 'Template_Import_SP');
      },
      onImport: async (json) => {
        const matchBranch = (str) => {
          if (!str) return null;
          const s = String(str || '').toLowerCase();
          const b = branchOptions.find(r => String(r.label || '').toLowerCase() === s);
          return b ? b.value : null;
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
          employee_name: String(row['Nama Karyawan'] || '').trim(),
          division: String(row['Divisi'] || '').trim(),
          branch_id: matchBranch(String(row['Cabang'] || '').trim()),
          tanggal: parseDate(row['Tanggal Sp']),
          akhir_sp: parseDate(row['Akhir Sp']),
          sp_type: String(row['Jenis Sp'] || '').trim(),
          document_link: String(row['Link Document / Foto'] || '').trim(),
        })).filter(r => r.employee_name && r.branch_id);
        
        const res = await apiFetch('/api/sp/import', { method: 'POST', body: JSON.stringify(payload) });
        if (!res.ok) throw new Error(res.data?.error || 'Import gagal');
      }
    },
    formFields: [
      { type: 'combobox', name: 'employee_name', label: 'Nama Karyawan', required: true, options: employeeOptions },
      { type: 'select', name: 'division', label: 'Divisi', options: ['FACILITY CARE', 'SECURITY'], required: true },
      { type: 'combobox', name: 'branch_id', label: 'Cabang', required: true, options: branchOptions, createApi: { path: '/api/branches', field: 'full_name' } },
      { type: 'date', name: 'tanggal', label: 'Tanggal Sp', required: true },
      { type: 'date', name: 'akhir_sp', label: 'Akhir Sp', required: true },
      { type: 'select', name: 'sp_type', label: 'Jenis Sp', required: true, options: ['SP 1', 'SP 2', 'SP 3', 'Teguran Lisan'] },
      { type: 'url', name: 'document_link', label: 'Link Document / Foto' }
    ]
  });
}
