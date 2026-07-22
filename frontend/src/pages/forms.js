import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { statusBadge } from '../components/badges.js';
import { createModal } from '../components/modal.js';
import { toastSuccess, toastError } from '../components/toast.js';
import { downloadExcel } from '../utils/excel.js';

export async function renderForms(container, mode = 'forms') {
  if (mode === 'supply') {
    return renderSupplyRequests(container);
  }
  renderMasterForms(container);
}

function renderMasterForms(container) {
  buildCrudPage({
    container,
    title: 'Master Form',
    icon: '📄',
    apiPath: '/api/forms',
    bulkDelete: true,
    itemLabel: 'Form',
    columns: [
      { key: 'name', label: 'Nama Form' },
      { key: 'category', label: 'Kategori' },
      { key: 'document_link', label: 'Dokumen', render: v => v ? `<a href="${v}" target="_blank" rel="noopener" class="btn btn-sm btn-primary">📄 Buka</a>` : '-' },
      { key: 'is_public', label: 'Publik', render: v => v ? '<span class="badge badge-success">Ya</span>' : '<span class="badge badge-neutral">Tidak</span>' },
      { key: 'description', label: 'Deskripsi' },
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari form...' },
    ],
    formFields: (data) => [
      { name: 'name', label: 'Nama Form', required: true, placeholder: 'Nama form', value: data?.name },
      { name: 'category', label: 'Kategori', placeholder: 'Permintaan Barang, Penilaian, dll.', value: data?.category },
      { name: 'document_link', label: 'Link Dokumen', type: 'url', placeholder: 'https://...', value: data?.document_link },
      { name: 'description', label: 'Deskripsi', type: 'textarea', rows: 2, value: data?.description },
      { name: 'is_public', label: 'Akses Publik', type: 'checkbox', checkLabel: 'Form dapat diakses tanpa login', value: data?.is_public },
    ],
  });
}

async function renderSupplyRequests(container) {
  const bRes = await apiFetch('/api/branches?all=1');
  const branchOptions = (bRes.data?.data || []).map(b => ({ value: b.id, label: b.full_name }));

  buildCrudPage({
    container,
    title: 'Permintaan Barang & Chemical',
    icon: '📦',
    apiPath: '/api/reports/supply',
    bulkDelete: true,
    itemLabel: 'Permintaan',
    canCreate: true, // Now fully editable
    columns: [
      { key: 'submitted_at', label: 'Waktu', nowrap: true, render: v => v ? new Date(v).toLocaleString('id-ID') : '-' },
      { key: 'submitter_name', label: 'Pengirim' },
      { key: 'branch_name', label: 'Cabang', render: (v, row) => row.branch_name_ref || row.branch_name || '-' },
      { key: 'tools_items', label: 'Alat/Barang', render: v => { try { const a = JSON.parse(v); return Array.isArray(a) ? a.join(', ') : v; } catch { return v || '-'; } } },
      { key: 'chemical_items', label: 'Chemical', render: v => { try { const a = JSON.parse(v); return Array.isArray(a) ? a.join(', ') : v; } catch { return v || '-'; } } },
      { key: 'additional_notes', label: 'Catatan', render: v => v?.length > 40 ? v.slice(0, 40) + '…' : (v || '-') },
      { key: 'status', label: 'Status', render: v => statusBadge(v) },
      { key: 'processed_by', label: 'Diproses Oleh' },
    ],
    filterFields: [
      { type: 'select', name: 'status', label: 'Status', options: ['Pending', 'Diproses', 'Selesai'] },
    ],
    formFields: (data) => {
      let tools = data?.tools_items;
      try { tools = Array.isArray(JSON.parse(tools)) ? JSON.parse(tools).join(', ') : tools; } catch {}
      let chems = data?.chemical_items;
      try { chems = Array.isArray(JSON.parse(chems)) ? JSON.parse(chems).join(', ') : chems; } catch {}
      return [
        {
          type: 'row', fields: [
            { name: 'submitter_name', label: 'Nama Pengirim', required: true, value: data?.submitter_name },
            { name: 'branch_id', label: 'Cabang', type: 'select', options: branchOptions, value: data?.branch_id },
          ]
        },
        {
          type: 'row', fields: [
            { name: 'tools_items', label: 'Alat / Barang', placeholder: 'Pisahkan dengan koma (Sapu, Mop)', value: tools },
            { name: 'tools_quantity', label: 'Jumlah Alat', type: 'number', value: data?.tools_quantity },
          ]
        },
        {
          type: 'row', fields: [
            { name: 'chemical_items', label: 'Chemical', placeholder: 'Pisahkan dengan koma', value: chems },
            { name: 'chemical_quantity', label: 'Jumlah Chemical', type: 'number', value: data?.chemical_quantity },
          ]
        },
        { name: 'additional_notes', label: 'Catatan', type: 'textarea', rows: 2, value: data?.additional_notes },
        { name: 'status', label: 'Status', type: 'select', options: ['Pending', 'Diproses', 'Selesai'], value: data?.status || '' },
        { name: 'processed_by', label: 'Diproses Oleh', value: data?.processed_by },
      ];
    },
    exportOptions: {
      moduleName: 'supply_requests',
      onExport: async (filters) => {
        const qs = new URLSearchParams(filters || {}).toString();
        const res = await apiFetch(`/api/reports/supply?limit=10000&${qs}`);
        if (res.ok) {
          const data = res.data.data.map(d => {
            let tools = d.tools_items;
            try { tools = Array.isArray(JSON.parse(tools)) ? JSON.parse(tools).join(', ') : tools; } catch {}
            let chems = d.chemical_items;
            try { chems = Array.isArray(JSON.parse(chems)) ? JSON.parse(chems).join(', ') : chems; } catch {}
            return {
              'Waktu': d.submitted_at || '',
              'Pengirim': d.submitter_name || '',
              'Cabang': d.branch_name_ref || d.branch_name || '',
              'Alat/Barang': tools || '',
              'Chemical': chems || '',
              'Catatan': d.additional_notes || '',
              'Status': d.status || '',
              'Diproses Oleh': d.processed_by || ''
            };
          });
          downloadExcel(data, `Permintaan_Barang_${new Date().toISOString().slice(0,10)}`);
        } else throw new Error('Gagal mengambil data');
      },
      onTemplate: () => {
        const template = [
          { 'Waktu': '2026-01-08', 'Pengirim': 'Fajar', 'Cabang': '001. Pondok Bambu', 'Alat/Barang': 'Sapu, Mop', 'Chemical': 'Karbol', 'Catatan': 'Mendesak', 'Status': 'Pending', 'Diproses Oleh': '' }
        ];
        downloadExcel(template, 'Template_Import_Permintaan');
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
          submitted_at: parseDate(row['Waktu'] || row['Tanggal']),
          submitter_name: String(row['Pengirim'] || '').trim(),
          branch_id: matchBranch(String(row['Cabang'] || '').trim()),
          tools_items: String(row['Alat/Barang'] || row['Alat'] || '').trim(),
          chemical_items: String(row['Chemical'] || '').trim(),
          additional_notes: String(row['Catatan'] || row['Keterangan'] || '').trim(),
          status: String(row['Status'] || '').trim(),
          processed_by: String(row['Diproses Oleh'] || row['PIC'] || '').trim(),
        })).filter(row => row.submitted_at && row.submitter_name && row.branch_id);
        
        const res = await apiFetch('/api/reports/supply/import', {
          method: 'POST',
          body: JSON.stringify(payload)
        });
        if (!res.ok) throw new Error(res.data?.error || 'Import gagal');
      }
    },
    extraActions: [
      {
        label: 'Update Status',
        icon: '🔄',
        class: 'btn-secondary',
        handler: (row, reload) => {
          const modal = createModal({
            title: 'Update Status Permintaan',
            content: `
              <div class="form-group">
                <label class="form-label">Status</label>
                <select class="form-control" id="supply-status">
                  <option value="Pending" ${row.status === 'Pending' ? 'selected' : ''}>Pending</option>
                  <option value="Diproses" ${row.status === 'Diproses' ? 'selected' : ''}>Diproses</option>
                  <option value="Selesai" ${row.status === 'Selesai' ? 'selected' : ''}>Selesai</option>
                </select>
              </div>
              <div class="form-group">
                <label class="form-label">Diproses Oleh</label>
                <input type="text" class="form-control" id="supply-processed-by" value="${row.processed_by || ''}" placeholder="Nama">
              </div>
            `,
            onConfirm: async (overlay, close) => {
              const status = overlay.querySelector('#supply-status').value;
              const processed_by = overlay.querySelector('#supply-processed-by').value;
              const res = await apiFetch(`/api/reports/supply/${row.id}`, {
                method: 'PUT',
                body: JSON.stringify({ status, processed_by }),
              });
              if (res.ok) { toastSuccess('Status diperbarui.'); close(); reload(); }
              else toastError('Gagal update status.');
            },
          });
        },
      },
    ],
  });
}
