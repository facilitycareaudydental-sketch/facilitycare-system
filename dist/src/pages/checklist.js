import { buildCrudPage } from './_crud.js';

export async function renderChecklist(container) {
  buildCrudPage({
    container,
    title: 'Master Checklist',
    icon: '✅',
    apiPath: '/api/checklist',
    bulkDelete: true,
    itemLabel: 'Checklist',
    columns: [
      { key: 'name', label: 'Nama Checklist' },
      { key: 'category', label: 'Kategori' },
      { key: 'document_link', label: 'Dokumen', render: v => v ? `<a href="${v}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">📄 Buka Dokumen</a>` : '-' },
      { key: 'description', label: 'Deskripsi' },
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari checklist...' },
    ],
    exportOptions: {
      moduleName: 'checklist',
      onExport: async (filters) => {
        const qs = new URLSearchParams(filters || {}).toString();
        const { apiFetch } = await import('../config.js');
        const res = await apiFetch(`/api/checklist?limit=10000&${qs}`);
        if (res.ok) {
          const data = res.data.data.map(d => ({
            'Nama Checklist': d.name || '',
            'Kategori': d.category || '',
            'Dokumen': d.document_link || '',
            'Deskripsi': d.description || ''
          }));
          const { downloadExcel } = await import('../utils/excel.js');
          downloadExcel(data, `Master_Checklist_${new Date().toISOString().slice(0,10)}`);
        } else throw new Error('Gagal mengambil data');
      },
      onTemplate: async () => {
        const template = [
          { 'Nama Checklist': 'Checklist Kebersihan Mingguan', 'Kategori': 'Master Cleaning Program', 'Dokumen': 'https://link.com', 'Deskripsi': 'Deskripsi singkat' }
        ];
        const { downloadExcel } = await import('../utils/excel.js');
        downloadExcel(template, 'Template_Import_Checklist');
      },
      onImport: async (json) => {
        const payload = json.map(row => ({
          name: String(row['Nama Checklist'] || '').trim(),
          category: String(row['Kategori'] || '').trim(),
          document_link: String(row['Dokumen'] || '').trim(),
          description: String(row['Deskripsi'] || '').trim(),
        })).filter(r => r.name);
        
        const { apiFetch } = await import('../config.js');
        const res = await apiFetch('/api/checklist/import', { method: 'POST', body: JSON.stringify(payload) });
        if (!res.ok) throw new Error(res.data?.error || 'Import gagal');
      }
    },
    formFields: (data) => [
      { name: 'name', label: 'Nama Checklist', required: true, placeholder: 'Nama checklist', value: data?.name },
      { name: 'category', label: 'Kategori', placeholder: 'Master Cleaning Program, dll.', value: data?.category },
      { name: 'document_link', label: 'Link Dokumen', type: 'url', placeholder: 'https://docs.google.com/...', value: data?.document_link },
      { name: 'description', label: 'Deskripsi', type: 'textarea', rows: 3, value: data?.description },
    ],
  });
}
