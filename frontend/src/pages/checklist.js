import { buildCrudPage } from './_crud.js';

export async function renderChecklist(container) {
  buildCrudPage({
    container,
    title: 'Master Checklist',
    icon: '✅',
    apiPath:,
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
    formFields: (data) => [
      { name: 'name', label: 'Nama Checklist', required: true, placeholder: 'Nama checklist', value: data?.name },
      { name: 'category', label: 'Kategori', placeholder: 'Master Cleaning Program, dll.', value: data?.category },
      { name: 'document_link', label: 'Link Dokumen', type: 'url', placeholder: 'https://docs.google.com/...', value: data?.document_link },
      { name: 'description', label: 'Deskripsi', type: 'textarea', rows: 3, value: data?.description },
    ],
  });
}
