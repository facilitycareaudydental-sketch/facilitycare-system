import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { statusBadge, daysRemainingBadge, divisionBadge } from '../components/badges.js';

let branchOptions = [];
let employeeOptions = [];

async function loadOptions() {
  const [bRes, eRes] = await Promise.all([
    apiFetch('/api/branches?all=1'),
    apiFetch('/api/employees?limit=500&status=Aktif'),
  ]);
  branchOptions = (bRes.data?.data || []).map(b => ({ value: b.id, label: b.full_name }));
  employeeOptions = (eRes.data?.data || []).map(e => ({ value: e.id, label: e.full_name }));
}

export async function renderContracts(container) {
  await loadOptions();

  buildCrudPage({
    container,
    title: 'Data Kontrak',
    icon: '📋',
    apiPath: '/api/contracts',
    itemLabel: 'Kontrak',
    columns: [
      { key: 'employee_name', label: 'Nama Karyawan' },
      { key: 'branch_name', label: 'Cabang' },
      { key: 'division', label: 'Divisi', render: v => divisionBadge(v) },
      { key: 'start_date', label: 'Tgl Mulai', nowrap: true },
      { key: 'end_date', label: 'Tgl Selesai', nowrap: true },
      { key: 'days_remaining', label: 'Sisa', render: v => daysRemainingBadge(v) },
      { key: 'contract_type', label: 'Tipe Kontrak' },
      { key: 'pkwt_number', label: 'PKWT' },
      { key: 'status', label: 'Status', render: v => statusBadge(v) },
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari nama karyawan...' },
      { type: 'select', name: 'branch_id', label: 'Cabang', options: branchOptions },
      { type: 'select', name: 'status', label: 'Status', options: ['Aktif', 'Tidak Aktif'] },
      { type: 'select', name: 'expiring_days', label: 'Akan Habis', options: [
        { value: '7', label: '7 Hari' },
        { value: '14', label: '14 Hari' },
        { value: '30', label: '30 Hari' },
        { value: '60', label: '60 Hari' },
      ]},
    ],
    formFields: (data) => [
      {
        type: 'row', fields: [
          { name: 'employee_id', label: 'Karyawan', type: 'select', required: true, options: employeeOptions, value: data?.employee_id },
          { name: 'branch_id', label: 'Cabang', type: 'select', options: branchOptions, value: data?.branch_id },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'division', label: 'Divisi', type: 'select', required: true, options: ['FACILITY CARE', 'SECURITY'], value: data?.division || 'FACILITY CARE' },
          { name: 'status', label: 'Status', type: 'select', required: true, options: ['Aktif', 'Tidak Aktif'], value: data?.status || 'Aktif' },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'start_date', label: 'Tanggal Mulai', type: 'date', required: true, value: data?.start_date },
          { name: 'end_date', label: 'Tanggal Selesai', type: 'date', required: true, value: data?.end_date },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'contract_type', label: 'Tipe Kontrak', type: 'select', options: ['KONTRAK 6 BULAN', 'KONTRAK 1 TAHUN', 'KONTRAK 2 TAHUN'], value: data?.contract_type },
          { name: 'pkwt_number', label: 'No. PKWT', type: 'select', options: ['PKWT 1', 'PKWT 2', 'PKWT 3', 'PKWT 4', 'PKWT 5', 'PKWT 6'], value: data?.pkwt_number },
        ]
      },
      { name: 'notes', label: 'Catatan', type: 'textarea', rows: 2, value: data?.notes },
    ],
  });
}
