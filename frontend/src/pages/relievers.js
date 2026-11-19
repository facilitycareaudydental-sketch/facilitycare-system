import { buildCrudPage } from './_crud.js?v=v3';
import { apiFetch } from '../config.js';
import { getCachedBranches, getCachedEmployees, fetchExportRows } from '../utils/dataCache.js';
import { statusBadge, periodBadge } from '../components/badges.js';
import { downloadExcel } from '../utils/excel.js';
import { submitJournaledImport } from '../utils/import_client.js';

export async function renderRelievers(container, params) {
  window.__RELIEVER_BUILD__ = "V3";
  console.log("RELIEVER PAGE LOADED");
  const branchOptions = await getCachedBranches();
  const employeeOptions = await getCachedEmployees();
  
  const dashFilter = params ? params.get('dash_filter') : null;
  
  let defFilters = {};
  if (dashFilter === 'reliever') {
    // If arriving from dashboard, default to Done status and selected month
    const now = new Date();
    const curM = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    const targetMonth = params && params.get('month') ? params.get('month') : curM;
    defFilters = { status: 'Done', month: targetMonth };
  }

  const getEmpOptions = (val) => {
    if (val && !employeeOptions.find(o => o.value === val)) {
      return [...employeeOptions, { value: val, label: val }];
    }
    return employeeOptions;
  };

  const relieverOptions = ['Agung Septiadi', 'Wasrikin', 'Iqbal Al Banna', 'Muhammad Tri Ismandanu'];
  
  const getRelieverOptions = (val) => {
    if (val && !relieverOptions.includes(val)) {
      return [...relieverOptions, val];
    }
    return relieverOptions;
  };

  const currentYear = new Date().getFullYear();
  const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  const monthOptions = monthNames.map((name, i) => {
      const monthNum = String(i + 1).padStart(2, '0');
      return { value: `${currentYear}-${monthNum}`, label: `${name} ${currentYear}` };
  });

  buildCrudPage({
    container,
    title: 'Jadwal Reliefer',
    icon: '🔄',
    apiPath: '/api/relievers',
    bulkDelete: false, // Disabled to prevent D1 vs Google Sheets outbox sync mismatch
    itemLabel: 'Reliefer',
    enableMobileFilterSheet: true,
    defaultFilters: defFilters,
    defaultSort: 'backup_date_desc',
    columns: [
      { key: 'branch_name', label: 'Cabang' },
      { key: 'original_fc_name', label: 'Nama Facility care' },
      { key: 'period', label: 'Periode', render: v => periodBadge(v) },
      { key: 'reliever_name', label: 'Relifer' },
      { key: 'backup_date', label: 'Tanggal Back Up', nowrap: true , render: v => window.formatDate(v) },
      { key: 'completion_date', label: 'Tanggal Selesai', nowrap: true , render: v => window.formatDate(v) },
      { key: 'reason', label: 'Keterangan' },
      { key: 'shift', label: 'Shift', render: v => v ? `<span class="badge badge-info">${v}</span>` : '-' },
      { key: 'status', label: 'Status', render: v => statusBadge(v) },
    ],
    filterFields: [
      { type: 'select', name: 'reliever_name', label: 'Cari reliefer / FC...', options: relieverOptions },
      { type: 'select', name: 'branch_id', label: 'Cabang', options: branchOptions },
      { type: 'select', name: 'month', label: 'Bulan', options: monthOptions },
      { type: 'select', name: 'period', label: 'Periode', options: ['Q1', 'Q2', 'Q3', 'Q4'] },
      { type: 'select', name: 'status', label: 'Status', options: ['Pending', 'Done', 'Tidak Datang'] },
    ],
    formFields: (data) => [
      {
        type: 'row', fields: [
          { name: 'branch_id', label: 'Cabang', type: 'combobox', required: true, options: branchOptions, value: data?.branch_id },
          { name: 'period', label: 'Periode', type: 'combobox', options: ['Q1', 'Q2', 'Q3', 'Q4'], value: data?.period },
        ]
      },
      {
        type: 'row', fields: [
          {
            name: 'original_fc_name',
            label: 'Nama Facility care',
            type: 'employee-combobox',
            allowManual: true,
            hasEmployeeIdSchema: false,
            manualBadgeText: 'Belum terhubung ke Master Karyawan (nama manual)',
            options: employeeOptions,
            value: data?.original_fc_name || '',
          },
          {
            name: 'reliever_name',
            label: 'Relifer',
            type: 'employee-combobox',
            required: true,
            allowManual: true,
            hasEmployeeIdSchema: false,
            manualBadgeText: 'Belum terhubung ke Master Karyawan (nama manual)',
            options: [...relieverOptions.map(r => ({ full_name: r, branch_name: 'Tenaga Pengganti (Reliefer)', division: '' })), ...employeeOptions],
            value: data?.reliever_name || '',
          },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'backup_date', label: 'Tanggal Back Up', type: 'date', required: true, value: data?.backup_date },
          { name: 'completion_date', label: 'Tanggal Selesai', type: 'date', value: data?.completion_date },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'reason', label: 'Keterangan', type: 'combobox', options: ['Cuti', 'Mengisi Kekosongan', 'Back Up Training', 'Deep Cleaning', 'Training Praktek Skill', 'Sakit', 'Lainnya'], value: data?.reason },
          { name: 'shift', label: 'Shift', type: 'combobox', options: ['Pagi', 'Siang', 'Full Shift', 'Middle'], value: data?.shift },
        ]
      },
      { name: 'status', label: 'Status', type: 'combobox', required: true, options: ['Pending', 'Done', 'Tidak Datang'], value: data?.status || '' },
    ],
    exportOptions: {
      moduleName: 'relievers',
      onExport: async (filters) => {
        const rows = await fetchExportRows('/api/relievers/export', filters);
        const data = rows.map(d => ({
          'Cabang': d.branch_name || '',
          'Nama Facility care': d.original_fc_name || '',
          'Periode': d.period || '',
          'Relifer': d.reliever_name || '',
          'Tanggal Back Up': d.backup_date || '',
          'Tanggal Selesai': d.completion_date || '',
          'Keterangan': d.reason || '',
          'Shift': d.shift || '',
          'Status': d.status || ''
        }));

        if (data.length === 0) {
          // Include headers even if empty
          data.push({
            'Cabang': '', 'Nama Facility care': '', 'Periode': '', 'Relifer': '',
            'Tanggal Back Up': '', 'Tanggal Selesai': '', 'Keterangan': '', 'Shift': '', 'Status': ''
          });
        }
        downloadExcel(data, 'Data_Reliefer');
      },
      onTemplate: () => {
        const template = [
          { 'Cabang': '001. Pondok Bambu', 'Nama Facility care': 'Budi Santoso', 'Periode': 'Q1', 'Relifer': 'Andi', 'Tanggal Back Up': '2024-03-10', 'Tanggal Selesai': '2024-03-10', 'Keterangan': 'Sakit', 'Shift': 'Pagi', 'Status': 'Done' }
        ];
        downloadExcel(template, 'Template_Import_Reliefer');
      },
      onImport: async (json, importCtx = {}) => {
        const bRes = await apiFetch('/api/branches?all=1');
        const rawBranches = bRes.data?.data || [];
        
        const matchBranch = (str) => {
          if (!str) return null;
          const s = String(str || '').toLowerCase();
          const b = rawBranches.find(r => String(r.full_name || '').toLowerCase() === s || String(r.code || '').toLowerCase() === s || String(r.name || '').toLowerCase() === s);
          return b ? b.id : null;
        };

        const payload = json.map(row => ({
          branch_name: String(row['Cabang'] || '').trim(),
          backup_date: String(row['Tanggal Back Up'] || row['Tanggal Backup'] || '').trim(),
          original_fc_name: String(row['Nama Facility care'] || row['FC Digantikan'] || '').trim(),
          reliever_name: String(row['Relifer'] || row['Reliefer'] || '').trim(),
          period: String(row['Periode'] || '').trim(),
          reason: String(row['Keterangan'] || '').trim(),
          shift: String(row['Shift'] || '').trim(),
          completion_date: String(row['Tanggal Selesai'] || '').trim(),
          status: String(row['Status'] || '').trim(),
        })).filter(row => row.reliever_name && row.backup_date);
        
        return await submitJournaledImport({
          module: 'relievers',
          label: 'Jadwal Reliefer',
          endpoint: '/api/import/relievers',
          body: { rows: payload, onDuplicate: 'update' },
          idempotencyKey: importCtx?.idempotencyKey,
          recover: Boolean(importCtx?.recover),
        });
      }
    }
  });
}
