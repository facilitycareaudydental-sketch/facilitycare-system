import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';

export async function renderTraining(container) {
  const [bRes, eRes, pRes] = await Promise.all([
    apiFetch('/api/branches?all=1'),
    apiFetch(`/api/training${window.location.search ? window.location.search + '&' : '?'}limit=10000`),
    apiFetch(`/api/training${window.location.search ? window.location.search + '&' : '?'}limit=10000`)
  ]);
  const branchOptions = (bRes.data?.data || []).map(b => ({ value: b.id, label: b.full_name }));
  const employeeOptions = (eRes.data?.data || []).map(e => ({ value: e.full_name, label: e.full_name }));
  const picOptions = (pRes.data?.data || []).filter(p => p.role === 'FC Spesialis').map(p => ({ value: p.name, label: p.name }));

  const getEmpOptions = (val) => {
    if (val && !employeeOptions.find(o => o.value === val)) {
      return [...employeeOptions, { value: val, label: val }];
    }
    return employeeOptions;
  };
  
  const getPicOptions = (val) => {
    if (val && !picOptions.find(o => o.value === val)) {
      return [...picOptions, { value: val, label: val }];
    }
    return picOptions;
  };
  const years = Array.from({ length: 5 }, (_, i) => String(new Date().getFullYear() - i));

  buildCrudPage({
    container,
    title: 'Training',
    icon: '🎓',
    apiPath: '/api/training',
    bulkDelete: true,
    itemLabel: 'Training',
    columns: [
      { key: 'training_date', label: 'Tanggal', nowrap: true , render: v => window.formatDate(v) },
      { key: 'batch', label: 'Batch' },
      { key: 'subject', label: 'Materi' },
      { key: 'branch_name', label: 'Cabang' },
      { key: 'trainer', label: 'Trainer' },
      { key: 'participants', label: 'Peserta', render: v => {
        try { const arr = JSON.parse(v); return Array.isArray(arr) ? arr.join(', ') : (v || '-'); } catch { return v || '-'; }
      }},
      { key: 'score', label: 'Nilai', render: v => v !== null && v !== undefined ? `<strong>${v}</strong>` : '-' },
      { key: 'document_link', label: 'Dokumen', render: v => v ? `<a href="${v}" target="_blank" rel="noopener" class="btn btn-xs btn-ghost">📄 Buka</a>` : '-' },
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari materi / trainer...' },
      { type: 'select', name: 'year', label: 'Tahun', options: years },
    ],
    exportOptions: {
      moduleName: 'training',
      onExport: async (filters) => {
        const qs = new URLSearchParams(filters || {}).toString();
        const res = await apiFetch(`/api/training?limit=10000&${qs}`);
        if (res.ok) {
          const data = res.data.data.map(d => {
            let participants = d.participants || '';
            try { const arr = JSON.parse(participants); participants = Array.isArray(arr) ? arr.join(', ') : participants; } catch {}
            return {
              'Tanggal': d.training_date || '',
              'Batch': d.batch || '',
              'Materi': d.subject || '',
              'Cabang': d.branch_name || '',
              'Trainer': d.trainer || '',
              'Peserta': participants,
              'Nilai': d.score !== null && d.score !== undefined ? d.score : '',
              'Dokumen': d.document_link || ''
            };
          });
          const { downloadExcel } = await import('../utils/excel.js');
          downloadExcel(data, `Data_Training_${new Date().toISOString().slice(0,10)}`);
        } else throw new Error('Gagal mengambil data');
      },
      onTemplate: async () => {
        const template = [
          { 'Tanggal': '2026-01-08', 'Batch': 'Batch 1', 'Materi': 'Standar Kebersihan', 'Cabang': '001. Pondok Bambu', 'Trainer': 'Budi', 'Peserta': 'Rina, Agus', 'Nilai': '85', 'Dokumen': 'https://link.doc' }
        ];
        const { downloadExcel } = await import('../utils/excel.js');
        downloadExcel(template, 'Template_Import_Training');
      },
      onImport: async (json) => {
        const matchBranch = (str) => {
          if (!str) return null;
          const s = String(str || '').toLowerCase();
          const b = bRes.data?.data.find(r => String(r.full_name || '').toLowerCase() === s || String(r.code || '').toLowerCase() === s || String(r.name || '').toLowerCase() === s);
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
          training_date: parseDate(row['Tanggal']),
          batch: String(row['Batch'] || '').trim(),
          subject: String(row['Materi'] || '').trim(),
          branch_id: matchBranch(String(row['Cabang'] || '').trim()),
          trainer: String(row['Trainer'] || '').trim(),
          participants: String(row['Peserta'] || '').trim(),
          score: row['Nilai'] ? Number(row['Nilai']) : null,
          document_link: String(row['Dokumen'] || '').trim(),
        })).filter(r => r.training_date && r.subject && r.branch_id);
        
        const res = await apiFetch('/api/training/import', { method: 'POST', body: JSON.stringify(payload) });
        if (!res.ok) throw new Error(res.data?.error || 'Import gagal');
      }
    },
    formFields: (data) => [
      {
        type: 'row', fields: [
          { name: 'training_date', label: 'Tanggal Training', type: 'date', required: true, value: data?.training_date },
          { name: 'batch', label: 'Batch', placeholder: 'Batch 1, Batch 2, ...', value: data?.batch },
        ]
      },
      { name: 'subject', label: 'Materi / Topik Training', required: true, placeholder: 'Judul materi training', value: data?.subject },
      {
        type: 'row', fields: [
          { name: 'branch_id', label: 'Cabang', type: 'combobox', options: (data?.branch_id && !branchOptions.find(o => o.value == data.branch_id)) ? [...branchOptions, { value: data.branch_id, label: data.branch_name || data.branch_id }] : branchOptions, createApi: { path: '/api/branches', field: 'full_name' }, value: data?.branch_id },
          { name: 'trainer', label: 'Trainer', type: 'combobox', options: getPicOptions(data?.trainer), createApi: { path: '/api/pic', field: 'name' }, value: data?.trainer },
        ]
      },
      { name: 'participants', label: 'Peserta (pisahkan dengan koma)', type: 'textarea', rows: 3, placeholder: 'Nama Peserta 1, Nama Peserta 2, ...', value: (() => { try { const arr = JSON.parse(data?.participants); return Array.isArray(arr) ? arr.join(', ') : (data?.participants || ''); } catch { return data?.participants || ''; } })() },
      {
        type: 'row', fields: [
          { name: 'score', label: 'Nilai / Score', type: 'number', step: '0.1', min: '0', max: '100', value: data?.score },
          { name: 'document_link', label: 'Link Dokumen', type: 'url', placeholder: 'https://...', value: data?.document_link },
        ]
      },
      { name: 'notes', label: 'Catatan', type: 'textarea', rows: 2, value: data?.notes },
    ],
    onBeforeSubmit: async (body) => {
      // Convert comma-separated participants to JSON array
      if (body.participants) {
        body.participants = JSON.stringify(body.participants.split(',').map(s => s.trim()).filter(Boolean));
      }
      return body;
    },
  });
}
