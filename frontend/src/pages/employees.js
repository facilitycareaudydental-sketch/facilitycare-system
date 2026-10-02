import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { getCachedBranches } from '../utils/dataCache.js';
import { statusBadge, divisionBadge } from '../components/badges.js';
import { downloadExcel } from '../utils/excel.js';

let branchOptions = [];
let rawBranches = [];

async function loadBranches() {
  const branchOptions = await getCachedBranches();
  return branchOptions;
}

export function formatMonth(d) {
  if (!d || d === '-' || String(d).trim() === '') return '-';
  const s = String(d).trim();
  const indoMonths = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];

  // Check ISO format YYYY-MM or YYYY-MM-DD
  const mMatch = s.match(/^(\d{4})-(\d{1,2})/);
  if (mMatch) {
    const year = mMatch[1];
    const monthIdx = parseInt(mMatch[2], 10) - 1;
    if (monthIdx >= 0 && monthIdx < 12) {
      return `${indoMonths[monthIdx]} ${year}`;
    }
  }

  // Check parts MM/YYYY or DD/MM/YYYY
  const parts = s.split(/[\/\-\.]/);
  if (parts.length === 2 && parts[1].length === 4) {
    const monthIdx = parseInt(parts[0], 10) - 1;
    if (monthIdx >= 0 && monthIdx < 12) return `${indoMonths[monthIdx]} ${parts[1]}`;
  }
  if (parts.length === 3 && parts[2].length === 4) {
    const monthIdx = parseInt(parts[1], 10) - 1;
    if (monthIdx >= 0 && monthIdx < 12) return `${indoMonths[monthIdx]} ${parts[2]}`;
  }

  // If already text containing month name
  for (let i = 0; i < 12; i++) {
    if (s.toLowerCase().includes(indoMonths[i].toLowerCase())) return s;
  }
  return s;
}

export function filterDashboardItem(s, type) {
  const status = String(s.status || '').toLowerCase();
  if (type === 'active') return status === 'aktif';
  if (type === 'reliefer') return s.division === 'FC - RELIEFER' && status === 'aktif';
  return false;
}

export async function renderEmployees(container, params) {
  const branchOptions = await loadBranches();
  
  const dashFilter = params ? params.get('dash_filter') : null;

  buildCrudPage({
    container,
    title: 'Karyawan',
    icon: '👥',
    apiPath: '/api/employees',
    enableMobileFilterSheet: true,
    itemLabel: 'Karyawan',
    bulkDelete: true,
    paginationMode: 'client',
    onDataLoaded: (items) => {
      if (dashFilter) {
        return items.filter(s => filterDashboardItem(s, dashFilter));
      }
      return items;
    },
    columns: [
      { key: 'full_name', label: 'Nama Lengkap' },
      { key: 'branch_name', label: 'Cabang' },
      { key: 'division', label: 'Divisi', render: (v) => divisionBadge(v) },
      { key: 'phone', label: 'No. HP', render: v => v ? `<a href="tel:${v}">${v}</a>` : '-' },
      { key: 'join_date', label: 'Tgl Masuk' , render: v => window.formatDate(v) },
      { key: 'target_pindah_os', label: 'Target Pindah OS', render: v => formatMonth(v) },
      { key: 'target_selesai_os', label: 'Target Selesai OS', render: v => formatMonth(v) },
      { key: 'status', label: 'Status', render: v => statusBadge(v) },
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari nama karyawan...' },
      { type: 'select', name: 'branch_id', label: 'Cabang', options: branchOptions },
      { type: 'select', name: 'division', label: 'Divisi', options: ['FACILITY CARE', 'SECURITY', 'FC - RELIEFER'] },
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
          { name: 'division', label: 'Divisi', type: 'select', required: true, options: ['FACILITY CARE', 'SECURITY', 'FC - RELIEFER'], value: data?.division || 'FACILITY CARE' },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'join_date', label: 'Tanggal Masuk', type: 'date', value: data?.join_date },
          { name: 'status', label: 'Status', type: 'select', required: true, options: ['Aktif', 'Tidak Aktif', 'Resign', 'Cut'], value: data?.status || '' },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'target_pindah_os', label: 'Target Pindah OS (Bulan)', type: 'month', value: data?.target_pindah_os ? String(data.target_pindah_os).slice(0, 7) : '' },
          { name: 'target_selesai_os', label: 'Target Selesai OS (Bulan)', type: 'month', value: data?.target_selesai_os ? String(data.target_selesai_os).slice(0, 7) : '' },
        ]
      },
      { name: 'notes', label: 'Catatan', type: 'textarea', rows: 2, value: data?.notes },
    ],
    exportOptions: {
      moduleName: 'employees',
      onExport: async () => {
        const res = await apiFetch(`/api/employees${window.location.search ? window.location.search + '&' : '?'}limit=10000`);
        if (res.ok) {
          const data = res.data.data.map(d => ({
            'Nama Lengkap': d.full_name,
            'Cabang': d.branch_name || '',
            'Divisi': d.division || '',
            'No. HP': d.phone || '',
            'Tgl Masuk': d.join_date || '',
            'Status': d.status || '',
            'Target Pindah OS': formatMonth(d.target_pindah_os) === '-' ? '' : formatMonth(d.target_pindah_os),
            'Target Selesai OS': formatMonth(d.target_selesai_os) === '-' ? '' : formatMonth(d.target_selesai_os),
          }));
          downloadExcel(data, 'Data_Karyawan');
        } else throw new Error('Gagal mengambil data');
      },
      onTemplate: () => {
        const template = [
          { 'Nama Lengkap': 'Budi Santoso', 'Cabang': '001. Pondok Bambu', 'Divisi': 'FACILITY CARE', 'No. HP': '08123456789', 'Tgl Masuk': '2024-01-15', 'Status': 'Aktif', 'Target Pindah OS': 'September 2026', 'Target Selesai OS': 'Oktober 2026' },
          { 'Nama Lengkap': 'Andi Saputra', 'Cabang': '002. Bintaro', 'Divisi': 'SECURITY', 'No. HP': '08987654321', 'Tgl Masuk': '2023-11-01', 'Status': 'Aktif', 'Target Pindah OS': 'November 2026', 'Target Selesai OS': 'Desember 2026' }
        ];
        downloadExcel(template, 'Template_Import_Karyawan');
      },
      onImport: async (json) => {
        // Prepare mapping: match by full_name, code, or short name
        const matchBranch = (str) => {
          if (!str) return null;
          const s = String(str || '').toLowerCase();
          const b = branchOptions.find(r => String(r.label || '').toLowerCase() === s);
          return b ? b.value : null;
        };
        
        const payload = json.map(row => ({
          full_name: String(row['Nama Lengkap'] || '').trim(),
          branch_id: matchBranch(String(row['Cabang'] || '').trim()),
          division: String(row['Divisi'] || row['Div / Bagian'] || '').trim() || 'FACILITY CARE',
          phone: String(row['No. HP'] || row['No. Hp'] || '').trim(),
          join_date: String(row['Tgl Masuk'] || row['Tanggal Masuk'] || '').trim(),
          status: String(row['Status'] || '').trim(),
          target_pindah_os: String(row['Target Pindah OS'] || '').trim(),
          target_selesai_os: String(row['Target Selesai OS'] || '').trim(),
          notes: String(row['Catatan'] || '').trim(),
        })).filter(row => row.full_name);
        
        const res = await apiFetch('/api/import/employees', {
          method: 'POST',
          body: JSON.stringify({ rows: payload, onDuplicate: 'update' })
        });
        if (!res.ok) throw new Error(res.data?.error || 'Import gagal');
        return res.data;
      }
    }
  });
}
