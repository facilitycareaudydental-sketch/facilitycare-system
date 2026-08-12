import { buildCrudPage } from './_crud.js';

export async function renderSOP(container) {
  buildCrudPage({
    container,
    title: 'SOP',
    icon: '📖',
    apiPath: '/api/sop',
    bulkDelete: true,
    itemLabel: 'SOP',
    columns: [
      { key: 'name', label: 'Nama SOP' },
      { key: 'category', label: 'Kategori' },
      { key: 'document_link', label: 'Dokumen', render: v => v ? `<a href="${v}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">📄 Buka Dokumen</a>` : '-' },
      { key: 'notes', label: 'Catatan' },
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari nama SOP...' },
    ],
    exportOptions: {
      moduleName: 'sop',
      onExport: async (filters) => {
        const qs = new URLSearchParams(filters || {}).toString();
        const { apiFetch } = await import('../config.js');
        const res = await apiFetch(`/api/sop?limit=10000&${qs}`);
        if (res.ok) {
          const data = res.data.data.map(d => ({
            'Nama SOP': d.name || '',
            'Kategori': d.category || '',
            'Dokumen': d.document_link || '',
            'Catatan': d.notes || d.description || ''
          }));
          const { downloadExcel } = await import('../utils/excel.js');
          downloadExcel(data, `Master_SOP_${new Date().toISOString().slice(0,10)}`);
        } else throw new Error('Gagal mengambil data');
      },
      onTemplate: async () => {
        const template = [
          { 'Nama SOP': 'SOP Cuci Tangan', 'Kategori': 'Ketentuan & Basic', 'Dokumen': 'https://link.com', 'Catatan': 'Catatan singkat' }
        ];
        const { downloadExcel } = await import('../utils/excel.js');
        downloadExcel(template, 'Template_Import_SOP');
      },
      onImport: async (json) => {
        const payload = json.map(row => ({
          name: String(row['Nama SOP'] || '').trim(),
          category: String(row['Kategori'] || '').trim(),
          document_link: String(row['Dokumen'] || '').trim(),
          description: String(row['Catatan'] || '').trim(),
        })).filter(r => r.name);
        
        const { apiFetch } = await import('../config.js');
        const res = await apiFetch('/api/sop/import', { method: 'POST', body: JSON.stringify(payload) });
        if (!res.ok) throw new Error(res.data?.error || 'Import gagal');
      }
    },
    formFields: (data) => [
      { name: 'name', label: 'Nama SOP', required: true, placeholder: 'Nama SOP', value: data?.name },
      { name: 'category', label: 'Kategori', placeholder: 'Ketentuan & Basic, Kualitas & Grooming, dst.', value: data?.category },
      { name: 'document_link', label: 'Link Dokumen', type: 'url', placeholder: 'https://docs.google.com/...', value: data?.document_link },
      { name: 'description', label: 'Deskripsi / Catatan', type: 'textarea', rows: 3, value: data?.description },
    ],
  });
}
