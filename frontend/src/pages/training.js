import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';

export async function renderTraining(container) {
  const [bRes, eRes, pRes] = await Promise.all([
    apiFetch('/api/branches?all=1'),
    apiFetch('/api/employees?limit=10000'),
    apiFetch('/api/pic')
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
    apiPath:,
    bulkDelete: true,
    itemLabel: 'Training',
    columns: [
      { key: 'training_date', label: 'Tanggal', nowrap: true },
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
          { name: 'branch_id', label: 'Cabang', type: 'select', options: branchOptions, value: data?.branch_id },
          { name: 'trainer', label: 'Trainer', type: 'select', options: getPicOptions(data?.trainer), value: data?.trainer },
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
