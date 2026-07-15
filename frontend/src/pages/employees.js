import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { statusBadge, divisionBadge } from '../components/badges.js';
import { downloadExcel } from '../utils/excel.js';

let branchOptions = [];
let rawBranches = [];

async function loadBranches() {
  const res = await apiFetch('/api/branches?all=1');
  rawBranches = res.data?.data || [];
  branchOptions = rawBranches.map(b => ({ value: b.id, label: b.full_name }));
}

export async function renderEmployees(container) {
  await loadBranches();

  buildCrudPage({
    container,
    title: 'Karyawan',
    icon: '👥',
    apiPath: '/api/employees',
    itemLabel: 'Karyawan',
    columns: [
      { key: 'full_name', label: 'Nama Lengkap' },
      { key: 'branch_name', label: 'Cabang' },
      { key: 'division', label: 'Divisi', render: (v) => divisionBadge(v) },
      { key: 'phone', label: 'No. HP', render: v => v ? `<a href="tel:${v}">${v}</a>` : '-' },
      { key: 'join_date', label: 'Tgl Masuk' },
      { key: 'status', label: 'Status', render: v => statusBadge(v) },
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari nama karyawan...' },
      { type: 'select', name: 'branch_id', label: 'Cabang', options: branchOptions },
      { type: 'select', name: 'division', label: 'Divisi', options: ['FACILITY CARE', 'SECURITY'] },
      { type: 'select', name: 'status', label: 'Status', options: ['Aktif', 'Tidak Aktif', 'Resign', 'Cut'] },
    ],
    formFields: (data) => [
      {
        type: 'row', fields: [
          { name: 'full_name', label: 'Nama Lengkap', required: true, placeholder: 'Nama lengkap karyawan', value: data?.full_name },
          { name: 'phone', label: 'No. HP', placeholder: '08xx-xxxx-xxxx', value: data?.phone },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'branch_id', label: 'Cabang', type: 'select', options: branchOptions, value: data?.branch_id },
          { name: 'division', label: 'Divisi', type: 'select', required: true, options: ['FACILITY CARE', 'SECURITY'], value: data?.division || 'FACILITY CARE' },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'join_date', label: 'Tanggal Masuk', type: 'date', value: data?.join_date },
          { name: 'status', label: 'Status', type: 'select', required: true, options: ['Aktif', 'Tidak Aktif', 'Resign', 'Cut'], value: data?.status || 'Aktif' },
        ]
      },
      { name: 'notes', label: 'Catatan', type: 'textarea', rows: 2, value: data?.notes },
    ],
    exportOptions: {
      moduleName: 'employees',
      onExport: async () => {
        const res = await apiFetch('/api/employees?limit=10000');
        if (res.ok) {
          const data = res.data.data.map(d => ({
            'Nama Lengkap': d.full_name,
            'Cabang': d.branch_name || '',
            'Divisi': d.division || '',
            'No. HP': d.phone || '',
            'Tgl Masuk': d.join_date || '',
            'Status': d.status || '',
            'Catatan': d.notes || ''
          }));
          downloadExcel(data, 'Data_Karyawan');
        } else throw new Error('Gagal mengambil data');
      },
      onTemplate: () => {
        const template = [
          { 'Nama Lengkap': 'Budi Santoso', 'Cabang': '001. Pondok Bambu', 'Divisi': 'FACILITY CARE', 'No. HP': '08123456789', 'Tgl Masuk': '2024-01-15', 'Status': 'Aktif', 'Catatan': '' },
          { 'Nama Lengkap': 'Andi Saputra', 'Cabang': '002. Bintaro', 'Divisi': 'SECURITY', 'No. HP': '08987654321', 'Tgl Masuk': '2023-11-01', 'Status': 'Aktif', 'Catatan': '' }
        ];
        downloadExcel(template, 'Template_Import_Karyawan');
      },
      onImport: async (json) => {
        // Prepare mapping: match by full_name, code, or short name
        const matchBranch = (str) => {
          if (!str) return null;
          const s = str.toLowerCase();
          const b = rawBranches.find(r => r.full_name.toLowerCase() === s || r.code.toLowerCase() === s || r.name.toLowerCase() === s);
          return b ? b.id : null;
        };
        
        const payload = json.map(row => ({
          full_name: String(row['Nama Lengkap'] || '').trim(),
          branch_id: matchBranch(String(row['Cabang'] || '').trim()),
          division: String(row['Divisi'] || '').trim() || 'FACILITY CARE',
          phone: String(row['No. HP'] || '').trim(),
          join_date: String(row['Tgl Masuk'] || '').trim(),
          status: String(row['Status'] || '').trim() || 'Aktif',
          notes: String(row['Catatan'] || '').trim(),
        })).filter(row => row.full_name);
        
        const res = await apiFetch('/api/employees/import', {
          method: 'POST',
          body: JSON.stringify(payload)
        });
        if (!res.ok) throw new Error(res.data?.error || 'Import gagal');
      }
    }
  });
}
