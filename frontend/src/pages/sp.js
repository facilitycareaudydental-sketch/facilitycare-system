import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';

let branchOptions = [];
let employeeOptions = [];

export async function renderSP(container) {
  const [bRes, eRes] = await Promise.all([
    apiFetch('/api/branches?all=1'),
    apiFetch('/api/employees?limit=10000')
  ]);
  branchOptions = (bRes.data?.data || []).map(b => ({ value: b.id, label: b.full_name }));
  employeeOptions = (eRes.data?.data || []).map(e => ({ value: e.full_name, label: e.full_name }));

  buildCrudPage({
    container,
    title: 'Data SP (Surat Peringatan)',
    icon: '✉️',
    apiPath: '/api/sp',
    itemLabel: 'SP',
    bulkDelete: true,
    columns: [
      { key: 'tanggal', label: 'Tanggal', render: v => v ? new Date(v).toLocaleDateString('id-ID', { year: 'numeric', month: 'short', day: 'numeric' }) : '-' },
      { key: 'employee_name', label: 'Nama Karyawan' },
      { key: 'branch_name', label: 'Cabang' },
      { key: 'sp_type', label: 'Jenis SP', render: v => `<span class="badge badge-warning">${v || '-'}</span>` },
      { key: 'status', label: 'Status', render: v => `<span class="badge ${v === 'Aktif' ? 'badge-danger' : 'badge-success'}">${v || '-'}</span>` },
      { key: 'document_link', label: 'Dokumen', render: v => v ? `<a href="${v}" target="_blank" class="text-primary hover-underline">Lihat</a>` : '-' }
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari nama karyawan...' },
      { type: 'select', name: 'branch_id', label: 'Cabang', options: branchOptions },
    ],
    exportOptions: {
      moduleName: 'sp_data',
      onExport: async (filters) => {
        const qs = new URLSearchParams(filters || {}).toString();
        const res = await apiFetch(`/api/sp?limit=10000&${qs}`);
        if (res.ok) {
          const data = res.data.data.map(d => ({
            'Tanggal': d.tanggal || '',
            'Nama Karyawan': d.employee_name || '',
            'Cabang': d.branch_name || '',
            'Jenis SP': d.sp_type || '',
            'Status': d.status || '',
            'Dokumen': d.document_link || ''
          }));
          const { downloadExcel } = await import('../utils/excel.js');
          downloadExcel(data, `Data_SP_${new Date().toISOString().slice(0,10)}`);
        } else throw new Error('Gagal mengambil data');
      },
      onTemplate: async () => {
        const template = [
          { 'Tanggal': '2026-01-08', 'Nama Karyawan': 'Budi Santoso', 'Cabang': '001. Pondok Bambu', 'Jenis SP': 'SP 1', 'Status': 'Aktif', 'Dokumen': 'https://link.doc' }
        ];
        const { downloadExcel } = await import('../utils/excel.js');
        downloadExcel(template, 'Template_Import_SP');
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
          tanggal: parseDate(row['Tanggal']),
          employee_name: String(row['Nama Karyawan'] || '').trim(),
          branch_id: matchBranch(String(row['Cabang'] || '').trim()),
          sp_type: String(row['Jenis SP'] || '').trim(),
          status: String(row['Status'] || '').trim(),
          document_link: String(row['Dokumen'] || '').trim(),
        })).filter(r => r.tanggal && r.employee_name && r.branch_id);
        
        const res = await apiFetch('/api/sp/import', { method: 'POST', body: JSON.stringify(payload) });
        if (!res.ok) throw new Error(res.data?.error || 'Import gagal');
      }
    },
    formFields: [
      { type: 'date', name: 'tanggal', label: 'Tanggal', required: true },
      { type: 'combobox', name: 'employee_name', label: 'Nama Karyawan', required: true, options: employeeOptions, createApi: { path: '/api/employees', field: 'full_name', extra: { status: 'Aktif' } } },
      { type: 'select', name: 'branch_id', label: 'Cabang', required: true, options: branchOptions },
      { type: 'select', name: 'sp_type', label: 'Jenis Surat Peringatan', required: true, options: ['SP 1', 'SP 2', 'SP 3', 'Teguran Lisan'] },
      { type: 'select', name: 'status', label: 'Status', required: true, options: ['Aktif', 'Selesai'] },
      { type: 'url', name: 'document_link', label: 'Link Dokumen (Opsional)' }
    ]
  });
}
