import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { statusBadge, periodBadge } from '../components/badges.js';
import { downloadExcel } from '../utils/excel.js';

export async function renderFoggingReports(container) {
  const bRes = await apiFetch('/api/branches?all=1');
  const branchOptions = (bRes.data?.data || []).map(b => ({ value: b.id, label: b.full_name }));
  const years = Array.from({ length: 4 }, (_, i) => String(new Date().getFullYear() - i));

  buildCrudPage({
    container,
    title: 'Rekap Fogging',
    icon: '💨',
    apiPath: '/api/reports/fogging',
    itemLabel: 'Fogging',
    bulkDelete: true,
    columns: [
      { key: 'branch_name', label: 'Cabang' },
      { key: 'activity_type', label: 'Jenis', render: v => `<span class="badge badge-warning">${v}</span>` },
      { key: 'period', label: 'Periode', render: v => periodBadge(v) },
      { key: 'activity_date', label: 'Tanggal', nowrap: true , render: v => window.formatDate(v) },
      { key: 'status', label: 'Status', render: v => statusBadge(v) },
      { key: 'document_link', label: 'Dokumen', render: v => v ? `<a href="${v}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">📄 Buka</a>` : '-' },
      { key: 'notes', label: 'Catatan', render: v => v || '-' },
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari nama cabang/lokasi...' },
      { type: 'select', name: 'branch_id', label: 'Cabang', options: branchOptions },
      { type: 'select', name: 'period', label: 'Periode', options: ['Q1', 'Q2', 'Q3', 'Q4'] },
      { type: 'select', name: 'status', label: 'Status', options: ['Pending', 'Done'] },
      { type: 'select', name: 'year', label: 'Tahun', options: years },
    ],
    formFields: (data) => [
      {
        type: 'row', fields: [
          { name: 'branch_id', label: 'Cabang', type: 'combobox', required: true, options: (data?.branch_id && !branchOptions.find(o => o.value == data.branch_id)) ? [...branchOptions, { value: data.branch_id, label: data.branch_name || data.branch_id }] : branchOptions, createApi: { path: '/api/branches', field: 'full_name' }, value: data?.branch_id },
          { name: 'period', label: 'Periode', type: 'select', required: true, options: ['Q1', 'Q2', 'Q3', 'Q4'], value: data?.period },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'activity_date', label: 'Tanggal', type: 'date', value: data?.activity_date },
          { name: 'status', label: 'Status', type: 'select', required: true, options: ['Pending', 'Done'], value: data?.status || '' },
        ]
      },
      { name: 'document_link', label: 'Link Dokumen', type: 'url', placeholder: 'https://...', value: data?.document_link },
      { name: 'notes', label: 'Catatan', type: 'textarea', rows: 2, value: data?.notes },
    ],
    exportOptions: {
      moduleName: 'fogging_reports',
      onExport: async (filters) => {
        const qs = new URLSearchParams(filters || {}).toString();
        const res = await apiFetch(`/api/reports/fogging?limit=10000&${qs}`);
        if (res.ok) {
          const data = res.data.data.map(d => ({
            'Cabang': d.branch_name || '',
            'Jenis': d.activity_type || 'Fogging',
            'Periode': d.period || '',
            'Tanggal': d.activity_date || '',
            'Status': d.status || '',
            'Link Dokumen': d.document_link || '',
            'Catatan': d.notes || ''
          }));
          downloadExcel(data, `Laporan_Fogging_${new Date().toISOString().slice(0,10)}`);
        } else throw new Error('Gagal mengambil data');
      },
      onTemplate: () => {
        const template = [
          { 'Cabang': '001. Pondok Bambu', 'Jenis': 'Fogging', 'Periode': 'Q1', 'Tanggal': '2026-01-08', 'Status': 'Done', 'Link Dokumen': 'https://drive.google.com/...', 'Catatan': 'Tuntas' }
        ];
        downloadExcel(template, 'Template_Import_Fogging');
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
          activity_type: String(row['Jenis'] || row['Kegiatan'] || 'Fogging').trim(),
          period: String(row['Periode'] || '').trim(),
          activity_date: parseDate(row['Tanggal']),
          status: String(row['Status'] || '').trim(),
          document_link: String(row['Link Dokumen'] || '').trim(),
          notes: String(row['Catatan'] || row['Keterangan'] || '').trim(),
        })).filter(row => row.branch_id && row.period && row.activity_date);
        
        const res = await apiFetch('/api/reports/fogging/import', {
          method: 'POST',
          body: JSON.stringify(payload)
        });
        if (!res.ok) throw new Error(res.data?.error || 'Import gagal');
      }
    }
  });
}
