import { buildCrudPage } from './_crud.js';
import { apiFetch } from '../config.js';
import { statusBadge, daysRemainingBadge, divisionBadge } from '../components/badges.js';
import { downloadExcel } from '../utils/excel.js';

let branchOptions = [];
let employeeOptions = [];

async function loadOptions() {
  const [bRes, eRes] = await Promise.all([
    apiFetch('/api/branches?all=1'),
    apiFetch('/api/employees?limit=10000&status=Aktif'),
  ]);
  branchOptions = (bRes.data?.data || []).map(b => ({ value: b.id, label: b.full_name }));
  employeeOptions = (eRes.data?.data || []).map(e => ({ value: e.id, label: e.full_name }));
}

const fetchAll = async (path) => {
  let all = [];
  let page = 1;
  while (true) {
    const m = await import('../config.js');
    const res = await m.apiFetch(`${path}${path.includes('?') ? '&' : '?'}limit=100&page=${page}`);
    if (!res.ok) break;
    const data = res.data?.data || res.data || [];
    const dataArr = Array.isArray(data) ? data : [];
    all = all.concat(dataArr);
    
    // If we got less than the limit, we're on the last page.
    if (dataArr.length < 100) break;
    
    // If the backend provides pagination metadata, respect it.
    if (res.data?.pagination && page >= res.data.pagination.pages) break;
    
    // Otherwise, if we got exactly 100 and no pagination metadata, 
    // we must assume there MIGHT be another page and fetch it.
    page++;
  }
  return all;
};

export async function renderContracts(container) {
  await loadOptions();


  buildCrudPage({
    container,
    title: 'Data Kontrak',
    icon: '📋',
    apiPath: '/api/contracts',
    bulkDelete: true,
    itemLabel: 'Kontrak',
    columns: [
      { key: 'employee_name', label: 'Nama Lengkap' },
      { key: 'branch_name', label: 'Cabang' },
      { key: 'division', label: 'Div / Bagian', render: v => divisionBadge(v) },
      { key: 'start_date', label: 'Tanggal Mulai', nowrap: true , render: v => window.formatDate(v) },
      { key: 'end_date', label: 'Tanggal Selesai', nowrap: true , render: v => (!v || String(v).startsWith('2099')) ? 'Tetap / PKWTT' : window.formatDate(v) },
      { key: 'days_remaining', label: 'Sisa Kontrak', render: (v, row) => (row.end_date && String(row.end_date).startsWith('2099')) ? '<span class="badge badge-success" style="background:#10B981;color:white;padding:4px 8px;border-radius:6px;font-size:0.75rem;font-weight:600">Tetap</span>' : daysRemainingBadge(v) },
      { key: 'status', label: 'Status', render: v => statusBadge(v) },
    ],
    filterFields: [
      { type: 'search', placeholder: 'Cari nama karyawan...' },
      { type: 'select', name: 'branch_id', label: 'Cabang', options: branchOptions },
      { type: 'select', name: 'status', label: 'Status', options: ['Aktif', 'Tidak Aktif', 'Resign', 'Cut'] },
      { type: 'select', name: 'expiring_days', label: 'Akan Habis', options: [
        { value: '7', label: '7 Hari' },
        { value: '14', label: '14 Hari' },
        { value: '30', label: '30 Hari' },
        { value: '60', label: '60 Hari' },
      ]},
    ],
    onBeforeSubmit: (body) => {
      if (!body.end_date) body.end_date = '2099-12-31';
      return body;
    },
    onAfterLoad: () => {
      if (!document.getElementById('btn-find-missing')) {
        const btn = document.createElement('button');
        btn.id = 'btn-find-missing';
        btn.className = 'btn btn-ghost';
        btn.innerHTML = '🔍 Cek Selisih Karyawan';
        btn.style.marginLeft = '8px';
        btn.style.color = '#EF4444';
        btn.style.border = '1px solid currentColor';
        btn.onclick = async () => {
          btn.innerHTML = '⌛ Mencari...';
          btn.disabled = true;
          try {
            const [activeEmps, allContracts] = await Promise.all([
              fetchAll('/api/employees?status=Aktif'),
              fetchAll('/api/contracts')
            ]);
            
            if (activeEmps.length > 0) {
              const activeContracts = allContracts.filter(c => c.status === 'Aktif');
              const activeConEmpIds = new Set(activeContracts.map(c => c.employee_id));
              
              const missing = activeEmps.filter(e => !activeConEmpIds.has(e.id));
              let html = `<p style="margin-bottom:12px">Data yang terbaca: <b>${activeEmps.length}</b> Karyawan Aktif, dan <b>${activeContracts.length}</b> Kontrak Aktif.</p>
              <p style="margin-bottom:12px">Terdapat <b>${missing.length}</b> karyawan aktif yang tidak memiliki "Kontrak Aktif". Berikut daftarnya:</p><ul style="padding-left:20px; max-height:400px; overflow-y:auto">`;
              missing.forEach(m => {
                 const theirContracts = allContracts.filter(c => c.employee_id === m.id);
                 let conInfo = '<span style="color:#F59E0B">Belum pernah di-input kontrak</span>';
                 if (theirContracts.length > 0) {
                    const last = theirContracts[0];
                    conInfo = `Pernah ada kontrak (Status: <b style="color:#EF4444">${last.status}</b>, Selesai: ${window.formatDate(last.end_date)})`;
                 }
                 html += `<li style="margin-bottom:8px"><b>${m.full_name}</b> <br><span style="font-size:0.85em;color:var(--text-2)">Cabang: ${m.branch_name || '-'} | ${conInfo}</span></li>`;
              });
              html += '</ul>';
              
              import('../components/modal.js').then(m => m.createModal({ title: 'Karyawan Tanpa Kontrak Aktif', content: html, cancelText: 'Tutup' }));
            }
          } catch(e) {
             console.error(e);
          }
          btn.innerHTML = '🔍 Cek Selisih Karyawan';
          btn.disabled = false;
        };
        const actionsEl = document.querySelector('.page-actions');
        if (actionsEl) actionsEl.appendChild(btn);
      }
    },
    formFields: (data) => [
      {
        type: 'row', fields: [
          { name: 'employee_id', label: 'Nama Lengkap', type: 'combobox', required: true, options: (data?.employee_id && !employeeOptions.find(o => o.value == data.employee_id)) ? [...employeeOptions, { value: data.employee_id, label: data.employee_name || data.employee_id }] : employeeOptions, createApi: { path: '/api/employees', field: 'full_name' }, value: data?.employee_id },
          { name: 'branch_id', label: 'Cabang', type: 'combobox', options: (data?.branch_id && !branchOptions.find(o => o.value == data.branch_id)) ? [...branchOptions, { value: data.branch_id, label: data.branch_name || data.branch_id }] : branchOptions, createApi: { path: '/api/branches', field: 'full_name' }, value: data?.branch_id },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'division', label: 'Div / Bagian', type: 'select', required: true, options: ['FACILITY CARE', 'SECURITY'], value: data?.division || 'FACILITY CARE' },
          { name: 'status', label: 'Status', type: 'select', required: true, options: ['Aktif', 'Tidak Aktif', 'Resign', 'Cut'], value: data?.status || '' },
        ]
      },
      {
        type: 'row', fields: [
          { name: 'start_date', label: 'Tanggal Mulai', type: 'date', required: true, value: data?.start_date },
          { name: 'end_date', label: 'Tanggal Selesai', type: 'date', value: data?.end_date && !String(data.end_date).startsWith('2099') ? data.end_date : '' },
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
        const res = await apiFetch(`/api/contracts${window.location.search ? window.location.search + '&' : '?'}limit=10000`);
        if (res.ok) {
          const data = res.data.data.map(d => ({
            'Nama Lengkap': d.employee_name,
            'Cabang': d.branch_name || '',
            'Div / Bagian': d.division || '',
            'Tanggal Mulai': d.start_date || '',
            'Tanggal Selesai': (d.end_date && String(d.end_date).startsWith('2099')) ? '' : (d.end_date || ''),
            'Sisa Kontrak': (d.end_date && String(d.end_date).startsWith('2099')) ? 'Tetap' : (d.days_remaining !== null && d.days_remaining !== undefined ? `${d.days_remaining} Hari` : ''),
            'Status': d.status || ''
          }));
          downloadExcel(data, 'Data_Kontrak');
        } else throw new Error('Gagal mengambil data');
      },
      onTemplate: () => {
        const template = [
          { 'Nama Lengkap': 'Budi Santoso', 'Cabang': '001. Pondok Bambu', 'Div / Bagian': 'FACILITY CARE', 'Tanggal Mulai': '2024-01-01', 'Tanggal Selesai': '2024-12-31', 'Sisa Kontrak': '365 Hari', 'Status': 'Aktif' }
        ];
        downloadExcel(template, 'Template_Import_Kontrak');
      },
      onImport: async (json) => {
        const [bRes, eData] = await Promise.all([
          apiFetch('/api/branches?limit=10000'),
          fetchAll(`/api/employees`)
        ]);
        const rawBranches = bRes.data?.data || [];
        const rawEmployees = eData || [];
        
        const matchBranch = (str) => {
          if (!str) return null;
          const s = String(str || '').replace(/\s+/g, ' ').toLowerCase().trim();
          const b = rawBranches.find(r => String(r.full_name || '').replace(/\s+/g, ' ').toLowerCase().trim() === s || String(r.code || '').replace(/\s+/g, ' ').toLowerCase().trim() === s || String(r.name || '').replace(/\s+/g, ' ').toLowerCase().trim() === s);
          return b ? b.id : null;
        };
        const matchEmployee = (str) => {
          if (!str) return null;
          const s = String(str || '').replace(/\s+/g, ' ').toLowerCase().trim();
          const e = rawEmployees.find(r => String(r.full_name || '').replace(/\s+/g, ' ').toLowerCase().trim() === s);
          return e ? e.id : null;
        };
        const parseDate = (v) => {
          if (!v) return '';
          if (v instanceof Date && !isNaN(v.getTime())) return v.toISOString().slice(0, 10);
          const s = String(v).trim();
          if (/^\d{4,5}(\.\d+)?$/.test(s)) {
            const n = Math.floor(Number(s));
            if (n > 20000 && n < 99999) {
              const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
              return isNaN(d.getTime()) ? '' : d.toISOString().slice(0, 10);
            }
          }
          if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
          const parts = s.split(/[\/\-\.]/);
          if (parts.length === 3) {
            const [a, b, c] = parts.map(p => p.trim());
            if (a.length === 4 && b.length <= 2 && c.length <= 2) return `${a}-${b.padStart(2, '0')}-${c.padStart(2, '0')}`;
            if (c.length === 4 && b.length <= 2 && a.length <= 2) return `${c}-${b.padStart(2, '0')}-${a.padStart(2, '0')}`;
          }
          return s;
        };

        const payload = json.map(row => ({
          employee_id: matchEmployee(String(row['Nama Lengkap'] || '').trim()),
          branch_id: matchBranch(String(row['Cabang'] || '').trim()),
          division: String(row['Div / Bagian'] || '').trim() || 'FACILITY CARE',
          start_date: parseDate(row['Tanggal Mulai']),
          end_date: parseDate(row['Tanggal Selesai']) || '2099-12-31',
          status: String(row['Status'] || '').trim(),
          _rawName: String(row['Nama Lengkap'] || '').trim()
        }));
        
        const missing = payload.filter(r => !r.employee_id || !r.start_date);
        if (missing.length > 0) {
           const names = missing.map(m => m._rawName).join(', ');
           const missingBecauseEmptyEmployee = payload.filter(r => !r.employee_id).length;
           const missingBecauseEmptyDate = payload.filter(r => !r.start_date).length;
           throw new Error(`[DB: ${rawEmployees.length} Emps (ex: ${rawEmployees[0]?.full_name || 'N/A'}), ${rawBranches.length} Branches] Terdapat ${missing.length} baris gagal (EmpID kosong: ${missingBecauseEmptyEmployee}, Date kosong: ${missingBecauseEmptyDate}). Cek karyawan: ${names}`);
        }
        
        const res = await apiFetch('/api/contracts/import', {
          method: 'POST',
          body: JSON.stringify(payload)
        });
        if (!res.ok) throw new Error(res.data?.error || 'Import gagal');
      }
    }
  });
}
