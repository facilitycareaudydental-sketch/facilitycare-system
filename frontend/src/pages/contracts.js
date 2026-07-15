import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { statusBadge, daysRemainingBadge, divisionBadge } from '../components/badges.js';
import { downloadExcel } from '../utils/excel.js';

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
    exportOptions: {
      moduleName: 'contracts',
      onExport: async () => {
        const res = await apiFetch('/api/contracts?limit=10000');
        if (res.ok) {
          const data = res.data.data.map(d => ({
            'Nama Karyawan': d.employee_name,
            'Cabang': d.branch_name || '',
            'Divisi': d.division || '',
            'Tgl Mulai': d.start_date || '',
            'Tgl Selesai': d.end_date || '',
            'Tipe Kontrak': d.contract_type || '',
            'PKWT': d.pkwt_number || '',
            'Status': d.status || '',
            'Catatan': d.notes || ''
          }));
          downloadExcel(data, 'Data_Kontrak');
        } else throw new Error('Gagal mengambil data');
      },
      onTemplate: () => {
        const template = [
          { 'Nama Karyawan': 'Budi Santoso', 'Cabang': '001. Pondok Bambu', 'Divisi': 'FACILITY CARE', 'Tgl Mulai': '2024-01-01', 'Tgl Selesai': '2024-12-31', 'Tipe Kontrak': 'KONTRAK 1 TAHUN', 'PKWT': 'PKWT 1', 'Status': 'Aktif', 'Catatan': '' }
        ];
        downloadExcel(template, 'Template_Import_Kontrak');
      },
      onImport: async (json) => {
        const [bRes, eRes] = await Promise.all([
          apiFetch('/api/branches?all=1'),
          apiFetch('/api/employees?limit=10000')
        ]);
        const rawBranches = bRes.data?.data || [];
        const rawEmployees = eRes.data?.data || [];
        
        const matchBranch = (str) => {
          if (!str) return null;
          const s = str.toLowerCase();
          const b = rawBranches.find(r => r.full_name.toLowerCase() === s || r.code.toLowerCase() === s || r.name.toLowerCase() === s);
          return b ? b.id : null;
        };
        const matchEmployee = (str) => {
          if (!str) return null;
          const s = str.toLowerCase();
          const e = rawEmployees.find(r => r.full_name.toLowerCase() === s);
          return e ? e.id : null;
        };

        const payload = json.map(row => ({
          employee_id: matchEmployee(String(row['Nama Karyawan'] || '').trim()),
          branch_id: matchBranch(String(row['Cabang'] || '').trim()),
          division: String(row['Divisi'] || '').trim() || 'FACILITY CARE',
          start_date: String(row['Tgl Mulai'] || '').trim(),
          end_date: String(row['Tgl Selesai'] || '').trim(),
          contract_type: String(row['Tipe Kontrak'] || '').trim(),
          pkwt_number: String(row['PKWT'] || '').trim(),
          status: String(row['Status'] || '').trim() || 'Aktif',
          notes: String(row['Catatan'] || '').trim(),
        })).filter(row => row.employee_id && row.start_date && row.end_date); // require these fields
        
        const res = await apiFetch('/api/contracts/import', {
          method: 'POST',
          body: JSON.stringify(payload)
        });
        if (!res.ok) throw new Error(res.data?.error || 'Import gagal');
      }
    }
  });
}
