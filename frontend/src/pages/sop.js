import { buildCrudPage } from './_crud.js';

export async function renderSOP(container) {
  buildCrudPage({
    container,
    title: 'SOP',
    icon: '📖',
    apiPath:,
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
    formFields: (data) => [
      { name: 'name', label: 'Nama SOP', required: true, placeholder: 'Nama SOP', value: data?.name },
      { name: 'category', label: 'Kategori', placeholder: 'Ketentuan & Basic, Kualitas & Grooming, dst.', value: data?.category },
      { name: 'document_link', label: 'Link Dokumen', type: 'url', placeholder: 'https://docs.google.com/...', value: data?.document_link },
      { name: 'description', label: 'Deskripsi / Catatan', type: 'textarea', rows: 3, value: data?.description },
    ],
  });
}
