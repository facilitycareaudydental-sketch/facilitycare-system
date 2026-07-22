import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { statusBadge, periodBadge } from '../components/badges.js';
import { downloadExcel } from '../utils/excel.js';

export async function renderInspectionReports(container) {
  const bRes = await apiFetch('/api/branches?all=1');
  const branchOptions = (bRes.data?.data || []).map(b => ({ value: b.id, label: b.full_name }));
  const years = Array.from({ length: 4 }, (_, i) => String(new Date().getFullYear() - i));

  buildCrudPage({
    container,
    title: 'Laporan Inspeksi Hygiene',
    icon: '🔍',
    apiPath: '/api/reports/inspection',
    itemLabel: 'Laporan Inspeksi',
    bulkDelete: true,
    columns: [
      { key: 'branch_name', label: 'Cabang' },
      { key: 'period', label: 'Periode', render: v => periodBadge(v) },
      { key: 'inspection_date', label: 'Tanggal', nowrap: true , render: v => window.formatDate(v) },
      { key: 'fc_score', label: 'Point FC', render: v => v !== null && v !== undefined ? `<strong class="${v >= 80 ? 'text-success' : v >= 60 ? 'text-warning' : 'text-danger'}">${v}</strong>` : '-' },
      { key: 'spv_score', label: 'Point SPV', render: v => v !== null && v !== undefined ? `<strong>${v}</strong>` : '-' },
      { key: 'status', label: 'Status', render: v => statusBadge(v) },
      { key: 'document_link', label: 'Dokumen', render: v => v ? `<a href="${v}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">📄 Buka</a>` : '-' },
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari cabang / PIC...' },
      { type: 'select', name: 'branch_id', label: 'Cabang', options: branchOptions },
      { type: 'select', name: 'period', label: 'Periode', options: ['Q1', 'Q2', 'Q3', 'Q4'] },
      { type: 'select', name: 'status', label: 'Status', options: ['Pending', 'Done'] },
      { type: 'select', name: 'year', label: 'Tahun', options: years },
    ],
    formFields: (data) => [
      {
        type: 'row', fields: [
          { name: 'branch_id', label: 'Cabang', type: 'select', required: true, options: branchOptions, value: data?.branch_id },
          { name: 'period', label: 'Periode', type: 'select', required: true, options: ['Q1', 'Q2', 'Q3', 'Q4'], value: data?.period },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'inspection_date', label: 'Tanggal Inspeksi', type: 'date', required: true, value: data?.inspection_date },
          { name: 'status', label: 'Status', type: 'select', required: true, options: ['Pending', 'Done'], value: data?.status || '' },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'fc_score', label: 'Point FC', type: 'number', step: '0.1', min: '0', max: '100', value: data?.fc_score },
          { name: 'spv_score', label: 'Point SPV', type: 'number', step: '0.1', min: '0', max: '100', value: data?.spv_score },
        ]
      },
      { name: 'document_link', label: 'Link Dokumen', type: 'url', placeholder: 'https://drive.google.com/...', value: data?.document_link },
      { name: 'notes', label: 'Catatan', type: 'textarea', rows: 2, value: data?.notes },
    ],
    exportOptions: {
      moduleName: 'inspection_reports',
      onExport: async (filters) => {
        const qs = new URLSearchParams(filters || {}).toString();
        const res = await apiFetch(`/api/reports/inspection?limit=10000&${qs}`);
        if (res.ok) {
          const data = res.data.data.map(d => ({
            'Cabang': d.branch_name || '',
            'Periode': d.period || '',
            'Tanggal': d.inspection_date || '',
            'Point FC': d.fc_score !== null && d.fc_score !== undefined ? d.fc_score : '',
            'Point SPV': d.spv_score !== null && d.spv_score !== undefined ? d.spv_score : '',
            'Status': d.status || '',
            'Link Dokumen': d.document_link || '',
            'Catatan': d.notes || ''
          }));
          downloadExcel(data, `Laporan_Inspeksi_Hygiene_${new Date().toISOString().slice(0,10)}`);
        } else throw new Error('Gagal mengambil data');
      },
      onTemplate: () => {
        const template = [
          { 'Cabang': '001. Pondok Bambu', 'Periode': 'Q1', 'Tanggal': '2026-01-08', 'Point FC': 85, 'Point SPV': 90, 'Status': 'Done', 'Link Dokumen': 'https://drive.google.com/...', 'Catatan': 'Semua area bersih' }
        ];
        downloadExcel(template, 'Template_Import_Inspeksi');
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
          branch_id: matchBranch(String(row['Cabang'] || '').trim()),
          period: String(row['Periode'] || '').trim(),
          inspection_date: parseDate(row['Tanggal']),
          fc_score: row['Point FC'] !== undefined && row['Point FC'] !== '' ? Number(row['Point FC']) : null,
          spv_score: row['Point SPV'] !== undefined && row['Point SPV'] !== '' ? Number(row['Point SPV']) : null,
          status: String(row['Status'] || '').trim(),
          document_link: String(row['Link Dokumen'] || '').trim(),
          notes: String(row['Catatan'] || row['Keterangan'] || '').trim(),
        })).filter(row => row.branch_id && row.period && row.inspection_date);
        
        const res = await apiFetch('/api/reports/inspection/import', {
          method: 'POST',
          body: JSON.stringify(payload)
        });
        if (!res.ok) throw new Error(res.data?.error || 'Import gagal');
      }
    }
  });
}
