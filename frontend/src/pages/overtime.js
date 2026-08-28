import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { getCachedBranches, getCachedEmployees } from '../utils/dataCache.js';

let branchOptions = [];
let employeeOptions = [];

export async function renderOvertime(container) {
  branchOptions = await getCachedBranches();
  employeeOptions = await getCachedEmployees();

  buildCrudPage({
    container,
    title: 'Data Lembur',
    icon: 'fa-clock',
    apiPath: '/api/overtime',
    enableMobileFilterSheet: true,
    itemLabel: 'Lembur',
    bulkDelete: true,
    columns: [
      { key: 'date', label: 'Tanggal', render: v => v ? new Date(v).toLocaleDateString('id-ID', { year: 'numeric', month: 'short', day: 'numeric' }) : '-' },
      { key: 'branch_name', label: 'Cabang Lembur' },
      { key: 'employee_name', label: 'Nama Karyawan' },
      { key: 'start_time', label: 'Jam Mulai' },
      { key: 'end_time', label: 'Jam Selesai' },
      { key: 'break_hours', label: 'Istirahat', render: v => v != null ? `${v} Jam` : '0 Jam' },
      { key: 'total_hours', label: 'Total Jam', render: v => v != null ? `${v} Jam` : '0 Jam' },
      { key: 'reason', label: 'Alasan / Ket Lembur' }
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari nama karyawan...' },
      { type: 'select', name: 'branch_id', label: 'Cabang', options: branchOptions },
    ],
    exportOptions: {
      moduleName: 'overtime_records',
      onExport: async (filters) => {
        const qs = new URLSearchParams(filters || {}).toString();
        const res = await apiFetch(`/api/overtime?limit=10000&${qs}`);
        if (res.ok) {
          const data = res.data.data.map(d => ({
            'Tanggal': d.date || '',
            'Cabang Lembur': d.branch_name || '',
            'Nama Karyawan': d.employee_name || '',
            'Jam Mulai': d.start_time || '',
            'Jam Selesai': d.end_time || '',
            'Istirahat (Jam)': d.break_hours || 0,
            'Total Jam': d.total_hours || 0,
            'Alasan / Ket Lembur': d.reason || ''
          }));
          const { downloadExcel } = await import('../utils/excel.js');
          downloadExcel(data, `Data_Lembur_${new Date().toISOString().slice(0,10)}`);
        } else throw new Error('Gagal mengambil data');
      }
    },
    formFields: [
      { type: 'date', name: 'date', label: 'Tanggal', required: true },
      { type: 'select', name: 'branch_id', label: 'Cabang Lembur', required: true, options: branchOptions },
      { type: 'select', name: 'employee_id', label: 'Nama Karyawan', required: true, options: employeeOptions },
      { type: 'time', name: 'start_time', label: 'Jam Mulai', required: true },
      { type: 'time', name: 'end_time', label: 'Jam Selesai', required: true },
      { type: 'number', name: 'break_hours', label: 'Istirahat (Jam)', required: true },
      { type: 'number', name: 'total_hours', label: 'Total Jam (Jam)', required: true },
      { type: 'select', name: 'reason', label: 'Alasan / Ket Lembur', required: true, options: [
        'Cover manpower', 
        'Longshift', 
        'Pekerjaan urgent', 
        'General Cleaning', 
        'Deep Cleaning'
      ] }
    ]
  });
}
