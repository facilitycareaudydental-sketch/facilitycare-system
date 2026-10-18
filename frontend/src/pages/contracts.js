import { apiFetch, getUser } from '../config.js';
import { getCachedBranches, getCachedEmployees } from '../utils/dataCache.js';
import { createModal } from '../components/modal.js';
import { toastSuccess, toastError } from '../components/toast.js';
import { downloadExcel, parseExcel } from '../utils/excel.js';

// Helper format tanggal lokal Indonesia
function formatDateIndo(dateStr) {
  if (!dateStr || dateStr === '-') return '-';
  if (String(dateStr).startsWith('2099')) return 'Tetap / PKWTT';
  try {
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    return d.toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' });
  } catch {
    return dateStr;
  }
}

// Cek apakah karyawan tetap (PKWTT)
function isTetap(dateStr) {
  if (!dateStr) return false;
  return String(dateStr).startsWith('2099');
}

// Inisial avatar
function getInitials(name) {
  if (!name || name === '-') return '👤';
  const parts = String(name).trim().split(/\s+/);
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return (parts[0][0] + parts[1][0]).toUpperCase();
}

// Filter function for Dashboard integration
export function filterDashboardItem(s, type) {
  const status = String(s.status || '').toLowerCase();
  if (status !== 'aktif') return false;
  
  if (type === 'active') {
    return true; // All contracts with status 'Aktif' are considered active
  }
  
  if (type === 'expiring30') {
    if (!s.end_date || String(s.end_date).startsWith('2099')) return false;
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const next30 = new Date(today);
    next30.setDate(today.getDate() + 30);
    const end = new Date(s.end_date);
    end.setHours(0, 0, 0, 0);
    return end >= today && end <= next30;
  }
  return false;
}

export async function renderContracts(container, params) {
  let contractsList = [];
  let branchOptions = [];
  let employeeOptions = [];
  let employeePhoneMap = {};

  let searchQuery = '';
  let selectedBranch = 'all';
  let selectedDivision = 'all';
  let selectedStatus = 'all';
  let selectedExpiring = 'all';
  let isLoading = false;

  // Handle URL dashboard parameter
  const dashFilter = params ? params.get('dash_filter') : null;
  if (dashFilter === 'expiring30') {
    selectedExpiring = '30';
    selectedStatus = 'Aktif';
  } else if (dashFilter === 'active') {
    selectedStatus = 'Aktif';
  }

  const currentUser = getUser();
  const canEdit = currentUser && currentUser.role !== 'viewer';

  // Load option data
  try {
    const [branches, employees] = await Promise.all([
      getCachedBranches(),
      getCachedEmployees()
    ]);
    branchOptions = Array.isArray(branches) ? branches : [];
    employeeOptions = Array.isArray(employees) ? employees : [];

    // Map phone numbers for direct WhatsApp links
    employeeOptions.forEach(emp => {
      if (emp.id && emp.phone) {
        employeePhoneMap[emp.id] = emp.phone;
      }
      if (emp.full_name && emp.phone) {
        employeePhoneMap[emp.full_name.toLowerCase()] = emp.phone;
      }
    });
  } catch (err) {
    console.error('Error fetching cache options for contracts:', err);
  }

  // Load Contracts data from API
  async function loadData() {
    isLoading = true;
    renderUI();
    try {
      const res = await apiFetch('/api/contracts?limit=10000');
      if (res.ok) {
        const raw = res.data?.data || res.data || [];
        contractsList = Array.isArray(raw) ? raw : [];
      } else {
        contractsList = [];
        toastError('Gagal memuat data kontrak dari server.');
      }
    } catch (e) {
      console.error('Error load contracts:', e);
      contractsList = [];
      toastError('Terjadi kesalahan koneksi saat memuat kontrak.');
    } finally {
      isLoading = false;
      renderUI();
    }
  }

  // Filtered List calculation
  function getFilteredList() {
    return contractsList.filter(item => {
      // 1. Text Search (nama karyawan, no pkwt, cabang, catatan, tipe kontrak)
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const empName = String(item.employee_name || '').toLowerCase();
        const brName = String(item.branch_name || '').toLowerCase();
        const pkwt = String(item.pkwt_number || '').toLowerCase();
        const ctype = String(item.contract_type || '').toLowerCase();
        const notes = String(item.notes || '').toLowerCase();
        const matches = empName.includes(q) || brName.includes(q) || pkwt.includes(q) || ctype.includes(q) || notes.includes(q);
        if (!matches) return false;
      }

      // 2. Branch Filter
      if (selectedBranch !== 'all' && String(item.branch_id) !== String(selectedBranch)) {
        return false;
      }

      // 3. Division Filter
      if (selectedDivision !== 'all' && item.division !== selectedDivision) {
        return false;
      }

      // 4. Status Filter Pills
      if (selectedStatus === 'Aktif' && item.status !== 'Aktif') return false;
      if (selectedStatus === 'Expiring30') {
        const tetap = isTetap(item.end_date);
        const rem = Number(item.days_remaining);
        if (item.status !== 'Aktif' || tetap || isNaN(rem) || rem < 0 || rem > 30) {
          return false;
        }
      }
      if (selectedStatus === 'Tetap') {
        if (!isTetap(item.end_date) || item.status !== 'Aktif') return false;
      }
      if (selectedStatus === 'Nonaktif' && item.status === 'Aktif') return false;

      // 5. Expiring Days Select Filter
      if (selectedExpiring !== 'all') {
        const limitDays = parseInt(selectedExpiring, 10);
        const tetap = isTetap(item.end_date);
        const rem = Number(item.days_remaining);
        if (tetap || item.status !== 'Aktif' || isNaN(rem) || rem < 0 || rem > limitDays) {
          return false;
        }
      }

      return true;
    });
  }

  // Cek Selisih Karyawan (Karyawan Aktif yang tidak punya Kontrak Aktif)
  async function checkMissingContracts() {
    createModal({
      title: '🔍 Analisis Selisih Kontrak Karyawan',
      content: `
        <div style="text-align:center; padding:24px 10px;">
          <div style="font-size:2rem; margin-bottom:10px;">⏳</div>
          <p style="color:#64748B; font-size:0.9rem;">Sedang menyinkronkan data karyawan aktif dengan kontrak aktif di database...</p>
        </div>
      `,
      cancelText: 'Tutup'
    });

    try {
      const [empRes, conRes] = await Promise.all([
        apiFetch('/api/employees?limit=10000&status=Aktif'),
        apiFetch('/api/contracts?limit=10000')
      ]);

      const activeEmployees = empRes.ok ? (empRes.data?.data || empRes.data || []) : [];
      const allContracts = conRes.ok ? (conRes.data?.data || conRes.data || []) : [];

      const activeContractEmpIds = new Set(
        allContracts.filter(c => c.status === 'Aktif').map(c => c.employee_id)
      );

      const missing = activeEmployees.filter(e => !activeContractEmpIds.has(e.id));

      let contentHtml = '';
      if (missing.length === 0) {
        contentHtml = `
          <div style="text-align:center; padding:20px 10px;">
            <div style="font-size:2.8rem; margin-bottom:10px;">🎉</div>
            <h4 style="font-size:1.1rem; font-weight:800; color:#10B981; margin:0 0 8px 0;">Semua Karyawan Aktif Memiliki Kontrak!</h4>
            <p style="font-size:0.86rem; color:#64748B; margin:0;">
              Total <strong>${activeEmployees.length}</strong> karyawan aktif tercatat memiliki perikatan kontrak aktif yang sah di sistem.
            </p>
          </div>
        `;
      } else {
        contentHtml = `
          <div style="display:flex; flex-direction:column; gap:12px;">
            <div style="background:#FEF2F2; border:1px solid #FECACA; border-radius:10px; padding:12px; font-size:0.84rem; color:#991B1B;">
              Ditemukan <strong>${missing.length}</strong> dari <strong>${activeEmployees.length}</strong> karyawan aktif yang <strong>belum memiliki Kontrak Aktif</strong>.
            </div>
            <div style="max-height:360px; overflow-y:auto; display:flex; flex-direction:column; gap:8px; padding-right:4px;">
              ${missing.map(m => {
                const pastContracts = allContracts.filter(c => c.employee_id === m.id);
                let historyText = '<span style="color:#D97706; font-weight:600;">Belum pernah ada data kontrak</span>';
                if (pastContracts.length > 0) {
                  const last = pastContracts[0];
                  historyText = `Riwayat: Status <strong>${last.status}</strong> (Selesai: ${formatDateIndo(last.end_date)})`;
                }
                return `
                  <div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:8px; padding:10px 12px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px;">
                    <div>
                      <div style="font-weight:700; font-size:0.88rem; color:#0F172A;">${m.full_name}</div>
                      <div style="font-size:0.76rem; color:#64748B;">
                        Cabang: ${m.branch_name || '-'} • Divisi: ${m.division || 'FACILITY CARE'}
                      </div>
                      <div style="font-size:0.75rem; color:#475569; margin-top:2px;">
                        ${historyText}
                      </div>
                    </div>
                    ${canEdit ? `
                      <button class="btn btn-primary btn-add-contract-for-emp" data-empid="${m.id}" data-empname="${String(m.full_name).replace(/"/g, '&quot;')}" data-branchid="${m.branch_id || ''}" data-division="${m.division || 'FACILITY CARE'}" style="font-size:0.76rem; padding:5px 10px; border-radius:6px; background:#2563EB;">
                        + Buat Kontrak
                      </button>
                    ` : ''}
                  </div>
                `;
              }).join('')}
            </div>
          </div>
        `;
      }

      createModal({
        title: 'Analisis Selisih Kontrak Karyawan',
        content: contentHtml,
        cancelText: 'Tutup'
      });

      // Bind dynamic create buttons inside missing modal
      setTimeout(() => {
        document.querySelectorAll('.btn-add-contract-for-emp').forEach(btn => {
          btn.addEventListener('click', (e) => {
            const b = e.currentTarget;
            const empId = b.getAttribute('data-empid');
            const empName = b.getAttribute('data-empname');
            const branchId = b.getAttribute('data-branchid');
            const division = b.getAttribute('data-division');
            // Close active modal
            const closeBtn = document.querySelector('.modal-close');
            if (closeBtn) closeBtn.click();
            // Open form modal prefilled
            openFormModal({
              employee_id: empId,
              employee_name: empName,
              branch_id: branchId,
              division: division,
              status: 'Aktif'
            });
          });
        });
      }, 100);

    } catch (err) {
      console.error(err);
      toastError('Gagal melakukan analisis selisih data kontrak.');
    }
  }

  // Open Add/Edit Contract Modal
  function openFormModal(item = null) {
    const isEdit = !!(item && item.id);
    const modalTitle = isEdit ? `✏️ Edit Kontrak: ${item.employee_name || 'Karyawan'}` : '➕ Tambah Kontrak Karyawan Baru';

    const defaultEmpId = item?.employee_id || '';
    const defaultBranchId = item?.branch_id || '';
    const defaultDivision = item?.division || 'FACILITY CARE';
    const defaultStatus = item?.status || 'Aktif';
    const defaultStartDate = item?.start_date ? String(item.start_date).slice(0, 10) : '';
    const defaultEndDate = (item?.end_date && !isTetap(item.end_date)) ? String(item.end_date).slice(0, 10) : '';
    const defaultIsTetap = item ? isTetap(item.end_date) : false;
    const defaultContractType = item?.contract_type || 'KONTRAK 1 TAHUN';
    const defaultPkwt = item?.pkwt_number || 'PKWT 1';
    const defaultNotes = item?.notes || '';

    const contentHtml = `
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Karyawan <span style="color:#DC2626;">*</span>
            </label>
            <select id="con-modal-employee" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Karyawan --</option>
              ${employeeOptions.map(e => `
                <option value="${e.id}" ${String(defaultEmpId) === String(e.id) ? 'selected' : ''}>
                  ${e.full_name} (${e.division || 'FC'} - ${e.branch_name || 'Cabang -'})
                </option>
              `).join('')}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Penempatan <span style="color:#DC2626;">*</span>
            </label>
            <select id="con-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${branchOptions.map(b => `
                <option value="${b.value}" ${String(defaultBranchId) === String(b.value) ? 'selected' : ''}>
                  ${b.label}
                </option>
              `).join('')}
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Divisi / Bagian <span style="color:#DC2626;">*</span>
            </label>
            <select id="con-modal-division" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="FACILITY CARE" ${defaultDivision === 'FACILITY CARE' ? 'selected' : ''}>FACILITY CARE</option>
              <option value="SECURITY" ${defaultDivision === 'SECURITY' ? 'selected' : ''}>SECURITY</option>
              <option value="FC - RELIEFER" ${defaultDivision === 'FC - RELIEFER' ? 'selected' : ''}>FC - RELIEFER</option>
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Kontrak <span style="color:#DC2626;">*</span>
            </label>
            <select id="con-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="Aktif" ${defaultStatus === 'Aktif' ? 'selected' : ''}>🟢 Aktif</option>
              <option value="Tidak Aktif" ${defaultStatus === 'Tidak Aktif' ? 'selected' : ''}>⚪ Tidak Aktif</option>
              <option value="Resign" ${defaultStatus === 'Resign' ? 'selected' : ''}>🔴 Resign</option>
              <option value="Cut" ${defaultStatus === 'Cut' ? 'selected' : ''}>✂️ Cut / Pemutusan</option>
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tipe / Durasi Kontrak
            </label>
            <select id="con-modal-type" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;">
              <option value="KONTRAK 6 BULAN" ${defaultContractType === 'KONTRAK 6 BULAN' ? 'selected' : ''}>KONTRAK 6 BULAN</option>
              <option value="KONTRAK 1 TAHUN" ${defaultContractType === 'KONTRAK 1 TAHUN' ? 'selected' : ''}>KONTRAK 1 TAHUN</option>
              <option value="KONTRAK 2 TAHUN" ${defaultContractType === 'KONTRAK 2 TAHUN' ? 'selected' : ''}>KONTRAK 2 TAHUN</option>
              <option value="PKWTT (TETAP)" ${defaultContractType === 'PKWTT (TETAP)' ? 'selected' : ''}>PKWTT (TETAP)</option>
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tingkat / No. PKWT
            </label>
            <select id="con-modal-pkwt" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;">
              <option value="PKWT 1" ${defaultPkwt === 'PKWT 1' ? 'selected' : ''}>PKWT 1</option>
              <option value="PKWT 2" ${defaultPkwt === 'PKWT 2' ? 'selected' : ''}>PKWT 2</option>
              <option value="PKWT 3" ${defaultPkwt === 'PKWT 3' ? 'selected' : ''}>PKWT 3</option>
              <option value="PKWT 4" ${defaultPkwt === 'PKWT 4' ? 'selected' : ''}>PKWT 4</option>
              <option value="PKWT 5" ${defaultPkwt === 'PKWT 5' ? 'selected' : ''}>PKWT 5</option>
              <option value="PKWT 6" ${defaultPkwt === 'PKWT 6' ? 'selected' : ''}>PKWT 6</option>
              <option value="TETAP" ${defaultPkwt === 'TETAP' ? 'selected' : ''}>PKWTT / TETAP</option>
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Mulai Kontrak <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="con-modal-start" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${defaultStartDate}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Selesai Kontrak
            </label>
            <input type="date" id="con-modal-end" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; ${defaultIsTetap ? 'opacity:0.5;' : ''}" 
              value="${defaultEndDate}" ${defaultIsTetap ? 'disabled' : ''} />
            <div style="margin-top:6px; display:flex; align-items:center; gap:6px;">
              <input type="checkbox" id="con-modal-is-tetap" ${defaultIsTetap ? 'checked' : ''} style="cursor:pointer;" />
              <label for="con-modal-is-tetap" style="font-size:0.78rem; color:#475569; cursor:pointer; user-select:none;">
                Karyawan Tetap / PKWTT (Tanpa Batas Tanggal)
              </label>
            </div>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Catatan Tambahan
          </label>
          <textarea id="con-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; resize:vertical;" 
            placeholder="Keterangan perpanjangan, klausul khusus, atau nomor SK...">${defaultNotes}</textarea>
        </div>
      </div>
    `;

    createModal({
      title: modalTitle,
      content: contentHtml,
      confirmText: isEdit ? '💾 Simpan Perubahan' : '➕ Tambah Kontrak',
      cancelText: 'Batal',
      onConfirm: async (modalEl, close) => {
        const empEl = modalEl.querySelector('#con-modal-employee');
        const branchEl = modalEl.querySelector('#con-modal-branch');
        const divEl = modalEl.querySelector('#con-modal-division');
        const statusEl = modalEl.querySelector('#con-modal-status');
        const startEl = modalEl.querySelector('#con-modal-start');
        const endEl = modalEl.querySelector('#con-modal-end');
        const isTetapEl = modalEl.querySelector('#con-modal-is-tetap');
        const typeEl = modalEl.querySelector('#con-modal-type');
        const pkwtEl = modalEl.querySelector('#con-modal-pkwt');
        const notesEl = modalEl.querySelector('#con-modal-notes');

        const employee_id = empEl?.value;
        const branch_id = branchEl?.value;
        const division = divEl?.value || 'FACILITY CARE';
        const status = statusEl?.value || 'Aktif';
        const start_date = startEl?.value;
        const is_tetap = isTetapEl?.checked;
        const end_date = is_tetap ? '2099-12-31' : (endEl?.value || '2099-12-31');
        const contract_type = typeEl?.value;
        const pkwt_number = pkwtEl?.value;
        const notes = notesEl?.value?.trim() || null;

        if (!employee_id) {
          toastError('Harap pilih nama karyawan!');
          empEl?.focus();
          return;
        }
        if (!branch_id) {
          toastError('Harap pilih cabang penempatan!');
          branchEl?.focus();
          return;
        }
        if (!start_date) {
          toastError('Harap tentukan tanggal mulai kontrak!');
          startEl?.focus();
          return;
        }

        const payload = {
          employee_id: parseInt(employee_id, 10),
          branch_id: parseInt(branch_id, 10),
          division,
          status,
          start_date,
          end_date,
          contract_type,
          pkwt_number,
          notes
        };

        try {
          const url = isEdit ? `/api/contracts/${item.id}` : '/api/contracts';
          const method = isEdit ? 'PUT' : 'POST';

          const res = await apiFetch(url, {
            method,
            body: JSON.stringify(payload)
          });

          if (res.ok) {
            toastSuccess(isEdit ? 'Data kontrak berhasil diperbarui!' : 'Kontrak baru berhasil ditambahkan!');
            close();
            await loadData();
          } else {
            const errData = res.data?.error || res.data?.message || 'Gagal menyimpan data kontrak.';
            toastError(errData);
          }
        } catch (err) {
          console.error(err);
          toastError('Terjadi kesalahan koneksi ke server.');
        }
      }
    });

    // Auto-disable end_date input if is_tetap checked
    setTimeout(() => {
      const isTetapCheck = document.getElementById('con-modal-is-tetap');
      const endInput = document.getElementById('con-modal-end');
      if (isTetapCheck && endInput) {
        isTetapCheck.addEventListener('change', (e) => {
          if (e.target.checked) {
            endInput.disabled = true;
            endInput.style.opacity = '0.5';
          } else {
            endInput.disabled = false;
            endInput.style.opacity = '1';
          }
        });
      }
    }, 100);
  }

  // Delete Contract
  function confirmDelete(item) {
    createModal({
      title: '🗑️ Hapus Data Kontrak',
      content: `
        <div style="font-size:0.92rem; color:#334155; line-height:1.5;">
          Apakah Anda yakin ingin menghapus catatan kontrak untuk <strong>${item.employee_name || 'karyawan ini'}</strong> (ID #${item.id})?
          <div style="margin-top:10px; padding:10px; background:#FEF2F2; border-radius:8px; border:1px solid #FECACA; font-size:0.8rem; color:#991B1B;">
            ⚠️ Tindakan ini akan menghapus riwayat masa berlaku kontrak terkait dari sistem.
          </div>
        </div>
      `,
      confirmText: 'Ya, Hapus',
      cancelText: 'Batal',
      onConfirm: async (_, close) => {
        try {
          const res = await apiFetch(`/api/contracts/${item.id}`, { method: 'DELETE' });
          if (res.ok) {
            toastSuccess('Data kontrak berhasil dihapus.');
            close();
            await loadData();
          } else {
            toastError('Gagal menghapus kontrak: ' + (res.data?.message || 'Terjadi kesalahan'));
          }
        } catch (err) {
          console.error(err);
          toastError('Terjadi kendala koneksi.');
        }
      }
    });
  }

  // Ekspor Excel
  function exportToExcel() {
    if (!contractsList.length) {
      toastError('Belum ada data kontrak untuk diekspor.');
      return;
    }
    const exportData = contractsList.map(c => {
      const tetap = isTetap(c.end_date);
      return {
        'Nama Lengkap': c.employee_name || '',
        'Cabang': c.branch_name || '',
        'Divisi': c.division || '',
        'No. PKWT': c.pkwt_number || '',
        'Tipe Kontrak': c.contract_type || '',
        'Tanggal Mulai': c.start_date || '',
        'Tanggal Selesai': tetap ? 'Tetap / PKWTT' : (c.end_date || ''),
        'Sisa Hari': tetap ? 'Tetap' : (c.days_remaining !== null && c.days_remaining !== undefined ? `${c.days_remaining} Hari` : ''),
        'Status': c.status || '',
        'Catatan': c.notes || ''
      };
    });
    downloadExcel(exportData, `Data_Kontrak_Karyawan_${new Date().toISOString().slice(0, 10)}`);
    toastSuccess('File Excel Data Kontrak berhasil diunduh!');
  }

  // Main UI Renderer
  function renderUI() {
    const filtered = getFilteredList();

    // KPI Metrics calculation
    const totalCount = contractsList.length;
    const activeCount = contractsList.filter(c => c.status === 'Aktif').length;
    const expiringCount = contractsList.filter(c => {
      const tetap = isTetap(c.end_date);
      const rem = Number(c.days_remaining);
      return c.status === 'Aktif' && !tetap && !isNaN(rem) && rem >= 0 && rem <= 30;
    }).length;
    const permanentCount = contractsList.filter(c => isTetap(c.end_date) && c.status === 'Aktif').length;
    const nonActiveCount = contractsList.filter(c => c.status !== 'Aktif').length;

    container.innerHTML = `
      <style>
        .con-page {
          display: flex;
          flex-direction: column;
          gap: 20px;
          width: 100%;
          max-width: 100%;
          min-width: 0;
          box-sizing: border-box;
          overflow-x: hidden;
        }

        /* Hero Header Card */
        .con-hero-card {
          background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
          border-radius: 16px;
          padding: 24px 28px;
          color: white;
          box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25);
          position: relative;
          overflow: hidden;
          box-sizing: border-box;
          width: 100%;
          max-width: 100%;
          min-width: 0;
        }
        .con-hero-card::after {
          content: '';
          position: absolute;
          top: -40px;
          right: -40px;
          width: 220px;
          height: 220px;
          background: radial-gradient(circle, rgba(59, 130, 246, 0.2) 0%, transparent 70%);
          border-radius: 50%;
          pointer-events: none;
        }
        .con-stats-row {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 14px;
          margin-top: 20px;
        }
        .con-stat-box {
          background: rgba(255, 255, 255, 0.05);
          border: 1px solid rgba(255, 255, 255, 0.1);
          border-radius: 12px;
          padding: 14px 18px;
          display: flex;
          align-items: center;
          gap: 14px;
          backdrop-filter: blur(8px);
          transition: transform 0.15s ease, background 0.15s ease;
        }
        .con-stat-box:hover {
          background: rgba(255, 255, 255, 0.08);
          transform: translateY(-2px);
        }
        .con-stat-icon {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 1.3rem;
          flex-shrink: 0;
        }
        .con-stat-val {
          font-size: 1.55rem;
          font-weight: 800;
          line-height: 1.1;
          letter-spacing: -0.02em;
        }
        .con-stat-lbl {
          font-size: 0.72rem;
          text-transform: uppercase;
          color: #94A3B8;
          font-weight: 700;
          letter-spacing: 0.05em;
          margin-top: 2px;
        }

        /* Controls / Filter Bar */
        .con-controls-card {
          background: white;
          border-radius: 14px;
          border: 1px solid #E2E8F0;
          padding: 16px 20px;
          box-shadow: 0 2px 8px rgba(0,0,0,0.03);
          display: flex;
          flex-direction: column;
          gap: 14px;
          box-sizing: border-box;
          width: 100%;
          max-width: 100%;
          min-width: 0;
        }
        .con-search-row {
          display: flex;
          gap: 12px;
          flex-wrap: wrap;
        }
        .con-search-wrap {
          position: relative;
          flex: 1;
          min-width: 240px;
        }
        .con-search-icon {
          position: absolute;
          left: 14px;
          top: 50%;
          transform: translateY(-50%);
          color: #94A3B8;
          pointer-events: none;
        }
        .con-search-input {
          width: 100%;
          padding: 10px 14px 10px 40px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          outline: none;
          transition: border-color 0.15s ease;
        }
        .con-search-input:focus {
          border-color: #2563EB;
          box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);
        }
        .con-select {
          padding: 10px 14px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          color: #1E293B;
          background: white;
          outline: none;
          min-width: 170px;
          cursor: pointer;
        }
        .con-select:focus {
          border-color: #2563EB;
        }

        /* Swipeable Pills */
        .con-filter-pills {
          display: flex;
          align-items: center;
          gap: 8px;
          overflow-x: auto;
          white-space: nowrap;
          scrollbar-width: none;
          padding-bottom: 2px;
        }
        .con-filter-pills::-webkit-scrollbar {
          display: none;
        }
        .con-pill {
          padding: 6px 14px;
          border-radius: 20px;
          font-size: 0.82rem;
          font-weight: 600;
          cursor: pointer;
          border: 1px solid #E2E8F0;
          background: #F8FAFC;
          color: #64748B;
          transition: all 0.15s ease;
          display: inline-flex;
          align-items: center;
          gap: 6px;
          user-select: none;
          flex-shrink: 0;
        }
        .con-pill:hover {
          background: #F1F5F9;
          color: #1E293B;
        }
        .con-pill.active {
          background: #0F172A;
          color: white;
          border-color: #0F172A;
        }
        .con-pill-dot {
          width: 8px;
          height: 8px;
          border-radius: 50%;
        }
        .con-pill-count {
          padding: 1px 6px;
          border-radius: 10px;
          font-size: 0.72rem;
          background: rgba(0,0,0,0.06);
        }
        .con-pill.active .con-pill-count {
          background: rgba(255,255,255,0.2);
          color: white;
        }

        /* Cards Grid */
        .con-cards-grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
          gap: 16px;
          width: 100%;
          max-width: 100%;
          min-width: 0;
          box-sizing: border-box;
        }
        .con-card {
          background: white;
          border-radius: 14px;
          border: 1px solid #E2E8F0;
          padding: 16px 18px;
          box-shadow: 0 2px 8px rgba(0,0,0,0.04);
          display: flex;
          flex-direction: column;
          gap: 13px;
          width: 100% !important;
          max-width: 100% !important;
          min-width: 0 !important;
          box-sizing: border-box !important;
          overflow: hidden;
          transition: transform 0.15s ease, box-shadow 0.15s ease;
        }
        .con-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 6px 16px rgba(0,0,0,0.08);
          border-color: #CBD5E1;
        }

        /* Card Top Badges */
        .con-card-top {
          display: flex;
          justify-content: space-between;
          align-items: center;
          gap: 8px;
          width: 100%;
          min-width: 0;
          box-sizing: border-box;
          flex-wrap: wrap;
        }
        .con-div-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.74rem;
          font-weight: 700;
          flex-shrink: 0;
          white-space: nowrap;
        }
        .con-div-fc {
          background: #EFF6FF;
          color: #2563EB;
          border: 1px solid #DBEAFE;
        }
        .con-div-sec {
          background: #F1F5F9;
          color: #334155;
          border: 1px solid #E2E8F0;
        }
        .con-div-rel {
          background: #F5F3FF;
          color: #7C3AED;
          border: 1px solid #EDE9FE;
        }

        .con-pkwt-badge {
          display: inline-flex;
          align-items: center;
          gap: 4px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.74rem;
          font-weight: 700;
          background: #F5F3FF;
          color: #6D28D9;
          border: 1px solid #DDD6FE;
          flex-shrink: 0;
          white-space: nowrap;
        }

        .con-status-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.74rem;
          font-weight: 700;
          flex-shrink: 0;
          white-space: nowrap;
        }
        .con-status-aktif {
          background: #DCFCE7;
          color: #16A34A;
          border: 1px solid #BBF7D0;
        }
        .con-status-warning {
          background: #FEF3C7;
          color: #B45309;
          border: 1px solid #FDE68A;
        }
        .con-status-danger {
          background: #FEE2E2;
          color: #DC2626;
          border: 1px solid #FCA5A5;
        }
        .con-status-tetap {
          background: #E0E7FF;
          color: #4338CA;
          border: 1px solid #C7D2FE;
        }

        /* Profile Box */
        .con-profile-box {
          display: flex;
          align-items: center;
          gap: 12px;
          width: 100%;
          min-width: 0;
          box-sizing: border-box;
        }
        .con-avatar {
          width: 48px;
          height: 48px;
          border-radius: 12px;
          background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
          color: white;
          display: flex;
          align-items: center;
          justify-content: center;
          font-weight: 800;
          font-size: 1.05rem;
          flex-shrink: 0;
          box-shadow: 0 4px 10px rgba(37, 99, 235, 0.2);
        }
        .con-avatar.sec {
          background: linear-gradient(135deg, #334155 0%, #0F172A 100%);
          box-shadow: 0 4px 10px rgba(15, 23, 42, 0.2);
        }
        .con-avatar.rel {
          background: linear-gradient(135deg, #7C3AED 0%, #5B21B6 100%);
          box-shadow: 0 4px 10px rgba(124, 58, 237, 0.2);
        }
        .con-name {
          font-size: 1.02rem;
          font-weight: 800;
          color: #0F172A;
          line-height: 1.25;
          word-break: break-word;
          overflow-wrap: anywhere;
        }
        .con-subtitle {
          font-size: 0.76rem;
          color: #64748B;
          font-weight: 600;
          margin-top: 3px;
        }

        /* Info Container */
        .con-info-card {
          background: #F8FAFC;
          border: 1px solid #F1F5F9;
          border-radius: 10px;
          padding: 11px 12px;
          display: flex;
          flex-direction: column;
          gap: 8px;
          font-size: 0.83rem;
          width: 100%;
          max-width: 100%;
          min-width: 0;
          box-sizing: border-box;
        }
        .con-info-row {
          display: flex;
          align-items: flex-start;
          justify-content: space-between;
          gap: 8px;
          width: 100%;
          min-width: 0;
          box-sizing: border-box;
        }
        .con-info-label {
          color: #64748B;
          display: inline-flex;
          align-items: center;
          gap: 5px;
          flex-shrink: 0;
          white-space: nowrap;
          font-size: 0.8rem;
        }
        .con-info-val {
          color: #0F172A;
          font-weight: 600;
          text-align: right;
          min-width: 0;
          flex: 1;
          word-break: break-word;
          overflow-wrap: anywhere;
          font-size: 0.82rem;
        }

        /* Notes box */
        .con-notes-box {
          font-size: 0.78rem;
          color: #475569;
          background: #FFFBEB;
          border: 1px solid #FEF3C7;
          border-radius: 8px;
          padding: 8px 10px;
          line-height: 1.4;
          word-break: break-word;
          overflow-wrap: anywhere;
          width: 100%;
          box-sizing: border-box;
        }

        /* Card footer */
        .con-card-footer {
          display: flex;
          justify-content: space-between;
          align-items: center;
          border-top: 1px solid #F1F5F9;
          padding-top: 11px;
          margin-top: auto;
          gap: 8px;
          width: 100%;
          min-width: 0;
          box-sizing: border-box;
          flex-wrap: wrap;
        }
        .con-action-btns {
          display: flex;
          align-items: center;
          gap: 6px;
          margin-left: auto;
          flex-shrink: 0;
        }
        .con-icon-btn {
          width: 34px;
          height: 34px;
          border-radius: 8px;
          border: 1px solid #E2E8F0;
          background: white;
          color: #64748B;
          display: flex;
          align-items: center;
          justify-content: center;
          cursor: pointer;
          transition: all 0.15s ease;
        }
        .con-icon-btn:hover {
          background: #F1F5F9;
          color: #0F172A;
          border-color: #CBD5E1;
        }
        .con-icon-btn.delete:hover {
          background: #FEE2E2;
          color: #DC2626;
          border-color: #FCA5A5;
        }

        /* Empty state */
        .con-empty-state {
          grid-column: 1 / -1;
          background: white;
          border: 2px dashed #E2E8F0;
          border-radius: 16px;
          padding: 48px 24px;
          text-align: center;
        }

        /* Mobile Responsive */
        @media (max-width: 768px) {
          .con-page {
            gap: 14px;
            width: 100% !important;
            max-width: 100% !important;
            overflow-x: hidden !important;
          }
          .con-hero-card {
            padding: 16px 14px !important;
            width: 100% !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
          }
          .con-stats-row {
            grid-template-columns: 1fr 1fr;
            gap: 10px;
          }
          .con-stat-box {
            padding: 10px 12px;
          }
          .con-stat-icon {
            width: 38px;
            height: 38px;
            font-size: 1.1rem;
          }
          .con-stat-val {
            font-size: 1.3rem;
          }
          .con-hero-actions {
            width: 100% !important;
            display: flex !important;
            align-items: center !important;
            gap: 6px !important;
            flex-wrap: nowrap !important;
            overflow-x: auto !important;
            scrollbar-width: none !important;
            padding-bottom: 2px !important;
          }
          .con-hero-actions::-webkit-scrollbar {
            display: none;
          }
          .con-btn-add {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.8rem !important;
            white-space: nowrap !important;
          }
          .con-btn-excel {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.8rem !important;
            white-space: nowrap !important;
          }
          .con-btn-missing {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.8rem !important;
            white-space: nowrap !important;
          }
          .con-btn-refresh {
            width: 36px !important;
            height: 36px !important;
            padding: 0 !important;
            flex-shrink: 0 !important;
            justify-content: center !important;
          }
          .con-controls-card {
            padding: 14px !important;
            gap: 12px;
            width: 100% !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
          }
          .con-search-row {
            flex-direction: column;
            gap: 10px;
          }
          .con-select {
            width: 100%;
          }
          .con-cards-grid {
            grid-template-columns: minmax(0, 1fr) !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
          }
          .con-card {
            padding: 14px !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
          }
          .con-info-card {
            padding: 10px !important;
            width: 100% !important;
            max-width: 100% !important;
            box-sizing: border-box !important;
          }
        }
      </style>

      <div class="con-page">

        <!-- Hero Header Card -->
        <div class="con-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px;">
                <span style="font-size:1.6rem;">📋</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Monitoring Kontrak Kerja Karyawan
                </h1>
              </div>
            </div>

            <!-- Action Buttons 1 Baris Sejajar -->
            <div class="con-hero-actions">
              ${canEdit ? `
                <button id="con-btn-add" class="btn btn-primary con-btn-add" style="background:#2563EB; padding:9px 15px; border-radius:10px; font-weight:700; display:flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Kontrak</span>
                </button>
              ` : ''}

              <button id="con-btn-excel" class="btn btn-ghost con-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; padding:9px 14px; border-radius:10px; font-weight:600; display:flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>📥</span>
                <span>Ekspor Excel</span>
              </button>

              <button id="con-btn-missing" class="btn btn-ghost con-btn-missing" style="background:rgba(239,68,68,0.15); border:1px solid rgba(239,68,68,0.4); color:#FCA5A5; padding:9px 13px; border-radius:10px; font-weight:600; display:flex; align-items:center; gap:6px; white-space:nowrap;" title="Cek Karyawan Aktif yang belum memiliki Kontrak Aktif">
                <span>🔍</span>
                <span>Cek Selisih</span>
              </button>

              <button id="con-btn-refresh" class="btn btn-ghost con-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- 4 KPI Stat Boxes -->
          <div class="con-stats-row">
            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(59, 130, 246, 0.2); color:#60A5FA;">
                📋
              </div>
              <div>
                <div class="con-stat-val">${totalCount}</div>
                <div class="con-stat-lbl">Total Kontrak</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#34D399;">
                🟢
              </div>
              <div>
                <div class="con-stat-val" style="color:#34D399;">${activeCount}</div>
                <div class="con-stat-lbl">Kontrak Aktif</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(245, 158, 11, 0.2); color:#FBBF24;">
                ⚠️
              </div>
              <div>
                <div class="con-stat-val" style="color:#FBBF24;">${expiringCount}</div>
                <div class="con-stat-lbl">Habis &lt; 30 Hari</div>
              </div>
            </div>

            <div class="con-stat-box">
              <div class="con-stat-icon" style="background:rgba(99, 102, 241, 0.2); color:#A5B4FC;">
                🏛️
              </div>
              <div>
                <div class="con-stat-val" style="color:#A5B4FC;">${permanentCount}</div>
                <div class="con-stat-lbl">PKWTT (Tetap)</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Controls / Filter Card -->
        <div class="con-controls-card">
          <div class="con-search-row">
            <div class="con-search-wrap">
              <svg class="con-search-icon" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              <input type="text" id="con-search-input" class="con-search-input" 
                placeholder="Cari nama karyawan, cabang, no PKWT, catatan..." value="${searchQuery ? String(searchQuery).replace(/"/g, '&quot;') : ''}" />
            </div>

            <select id="con-branch-filter" class="con-select">
              <option value="all">Semua Cabang (${branchOptions.length})</option>
              ${branchOptions.map(b => `
                <option value="${b.value}" ${String(selectedBranch) === String(b.value) ? 'selected' : ''}>
                  ${b.label}
                </option>
              `).join('')}
            </select>

            <select id="con-division-filter" class="con-select">
              <option value="all">Semua Divisi</option>
              <option value="FACILITY CARE" ${selectedDivision === 'FACILITY CARE' ? 'selected' : ''}>FACILITY CARE</option>
              <option value="SECURITY" ${selectedDivision === 'SECURITY' ? 'selected' : ''}>SECURITY</option>
              <option value="FC - RELIEFER" ${selectedDivision === 'FC - RELIEFER' ? 'selected' : ''}>FC - RELIEFER</option>
            </select>

            <select id="con-expiring-filter" class="con-select">
              <option value="all" ${selectedExpiring === 'all' ? 'selected' : ''}>Semua Durasi Masa Habis</option>
              <option value="7" ${selectedExpiring === '7' ? 'selected' : ''}>Akan Habis ≤ 7 Hari</option>
              <option value="14" ${selectedExpiring === '14' ? 'selected' : ''}>Akan Habis ≤ 14 Hari</option>
              <option value="30" ${selectedExpiring === '30' ? 'selected' : ''}>Akan Habis ≤ 30 Hari</option>
              <option value="60" ${selectedExpiring === '60' ? 'selected' : ''}>Akan Habis ≤ 60 Hari</option>
            </select>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="con-filter-pills">
            <div class="con-pill ${selectedStatus === 'all' && selectedExpiring === 'all' ? 'active' : ''}" data-status="all">
              <span>Semua Kontrak</span>
              <span class="con-pill-count">${totalCount}</span>
            </div>
            <div class="con-pill ${selectedStatus === 'Aktif' && selectedExpiring === 'all' ? 'active' : ''}" data-status="Aktif">
              <span class="con-pill-dot" style="background:#10B981;"></span>
              <span>Aktif</span>
              <span class="con-pill-count">${activeCount}</span>
            </div>
            <div class="con-pill ${selectedStatus === 'Expiring30' ? 'active' : ''}" data-status="Expiring30">
              <span class="con-pill-dot" style="background:#F59E0B;"></span>
              <span>Akan Habis (&lt;30 Hari)</span>
              <span class="con-pill-count">${expiringCount}</span>
            </div>
            <div class="con-pill ${selectedStatus === 'Tetap' ? 'active' : ''}" data-status="Tetap">
              <span class="con-pill-dot" style="background:#6366F1;"></span>
              <span>PKWTT (Tetap)</span>
              <span class="con-pill-count">${permanentCount}</span>
            </div>
            <div class="con-pill ${selectedStatus === 'Nonaktif' ? 'active' : ''}" data-status="Nonaktif">
              <span class="con-pill-dot" style="background:#EF4444;"></span>
              <span>Nonaktif / Resign</span>
              <span class="con-pill-count">${nonActiveCount}</span>
            </div>
          </div>
        </div>

        <!-- Exclusive Cards Grid -->
        <div class="con-cards-grid">
          ${filtered.length === 0 ? `
            <div class="con-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">🔍</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak ada kontrak yang sesuai kriteria</h3>
              <p style="font-size:0.85rem; color:#64748B; margin:0 0 16px 0;">Coba atur ulang kata kunci pencarian, filter cabang, atau filter masa berlaku kontrak.</p>
              <button id="con-reset-filter" class="btn btn-secondary" style="font-size:0.82rem; padding:7px 16px; border-radius:8px;">Reset Semua Filter</button>
            </div>
          ` : filtered.map(item => {
            const div = item.division || 'FACILITY CARE';
            let divClass = 'con-div-fc';
            let avatarClass = '';
            if (div === 'SECURITY') {
              divClass = 'con-div-sec';
              avatarClass = 'sec';
            } else if (div === 'FC - RELIEFER') {
              divClass = 'con-div-rel';
              avatarClass = 'rel';
            }

            const tetap = isTetap(item.end_date);
            const isAktif = item.status === 'Aktif';
            const rem = Number(item.days_remaining);

            // Determine status badge class and label
            let statusBadgeClass = 'con-status-aktif';
            let statusBadgeText = '🟢 Aktif';

            if (tetap) {
              statusBadgeClass = 'con-status-tetap';
              statusBadgeText = '🏛️ PKWTT (Tetap)';
            } else if (isAktif) {
              if (rem <= 0) {
                statusBadgeClass = 'con-status-danger';
                statusBadgeText = '⚠️ Habis Hari Ini';
              } else if (rem <= 30) {
                statusBadgeClass = 'con-status-warning';
                statusBadgeText = `⚠️ Sisa ${rem} Hari`;
              } else {
                statusBadgeClass = 'con-status-aktif';
                statusBadgeText = `🟢 Aktif (${rem} Hari)`;
              }
            } else {
              statusBadgeClass = 'con-status-danger';
              statusBadgeText = `⚪ ${item.status || 'Nonaktif'}`;
            }

            const phone = employeePhoneMap[item.employee_id] || employeePhoneMap[String(item.employee_name || '').toLowerCase()] || '';

            return `
              <div class="con-card" data-id="${item.id}">
                <!-- Top Badges Row -->
                <div class="con-card-top">
                  <div style="display:flex; align-items:center; gap:6px; flex-wrap:wrap;">
                    <span class="con-div-badge ${divClass}">
                      ${div === 'SECURITY' ? '🛡️' : (div === 'FC - RELIEFER' ? '🔄' : '🧹')} ${div}
                    </span>
                    ${item.pkwt_number ? `
                      <span class="con-pkwt-badge">
                        📜 ${item.pkwt_number}
                      </span>
                    ` : ''}
                  </div>
                  <span class="con-status-badge ${statusBadgeClass}">
                    ${statusBadgeText}
                  </span>
                </div>

                <!-- Profile Box -->
                <div class="con-profile-box">
                  <div class="con-avatar ${avatarClass}">
                    ${getInitials(item.employee_name)}
                  </div>
                  <div style="min-width:0; flex:1;">
                    <div class="con-name">
                      ${item.employee_name || 'Karyawan Tanpa Nama'}
                    </div>
                    <div class="con-subtitle">
                      Personil ${div} • ${item.contract_type || 'Kontrak Kerja'}
                    </div>
                  </div>
                </div>

                <!-- Information Grid -->
                <div class="con-info-card">
                  <div class="con-info-row">
                    <span class="con-info-label">📍 Cabang:</span>
                    <span class="con-info-val">${item.branch_name || 'Cabang Tidak Ditentukan'}</span>
                  </div>
                  <div class="con-info-row">
                    <span class="con-info-label">🗓️ Mulai Kontrak:</span>
                    <span class="con-info-val">${formatDateIndo(item.start_date)}</span>
                  </div>
                  <div class="con-info-row">
                    <span class="con-info-label">🏁 Selesai Kontrak:</span>
                    <span class="con-info-val" style="${tetap ? 'color:#4338CA; font-weight:700;' : ''}">
                      ${tetap ? 'Tetap / PKWTT' : formatDateIndo(item.end_date)}
                    </span>
                  </div>
                  <div class="con-info-row">
                    <span class="con-info-label">⏳ Masa Berlaku:</span>
                    <span class="con-info-val" style="${tetap ? 'color:#15803D;' : (rem <= 30 && isAktif ? 'color:#D97706; font-weight:700;' : '')}">
                      ${tetap ? 'Karyawan Tetap (Permanen)' : (rem <= 0 ? 'Masa Kontrak Telah Habis' : `${rem} Hari Tersisa`)}
                    </span>
                  </div>
                </div>

                <!-- Notes box (if any) -->
                ${item.notes ? `
                  <div class="con-notes-box">
                    📝 <strong>Catatan:</strong> ${item.notes}
                  </div>
                ` : ''}

                <!-- Footer Actions -->
                <div class="con-card-footer">
                  <div style="font-size:0.75rem; color:#94A3B8;">
                    ID #${item.id} • ${item.pkwt_number || 'PKWT'}
                  </div>

                  <div class="con-action-btns">
                    ${phone ? `
                      <a href="https://wa.me/${phone.replace(/[^0-9]/g, '').replace(/^0/, '62')}" target="_blank" rel="noopener" class="con-icon-btn" title="Chat WhatsApp Karyawan" style="color:#16A34A;">
                        <svg width="16" height="16" fill="currentColor" viewBox="0 0 24 24"><path d="M12.031 6.172c-3.181 0-5.767 2.586-5.768 5.766-.001 1.298.38 2.27 1.019 3.287l-.582 2.128 2.182-.573c.978.58 1.911.928 3.145.929 3.178 0 5.767-2.587 5.768-5.766.001-3.187-2.575-5.771-5.764-5.771zm3.392 8.244c-.144.405-.837.774-1.17.824-.299.045-.677.063-1.092-.069-.252-.08-.575-.187-.988-.365-1.739-.751-2.874-2.502-2.961-2.617-.087-.116-.708-.94-.708-1.793s.448-1.273.607-1.446c.159-.173.346-.217.462-.217l.332.006c.106.005.249-.04.39.298.144.347.491 1.2.534 1.287.043.087.072.188.014.304-.058.116-.087.188-.173.289l-.26.304c-.087.086-.177.18-.076.354.101.174.449.741.964 1.201.662.591 1.221.774 1.394.86s.275.072.376-.043c.101-.116.433-.506.549-.68.116-.173.231-.145.39-.086s1.011.477 1.184.564.289.13.332.202c.045.072.045.419-.099.824zm-3.423-14.416c-6.627 0-12 5.373-12 12 0 2.159.57 4.185 1.564 5.939l-1.564 5.717 5.864-1.538c1.691.924 3.63 1.458 5.698 1.458 6.627 0 12-5.373 12-12s-5.373-12-12-12z"/></svg>
                      </a>
                    ` : ''}

                    ${canEdit ? `
                      <button class="con-icon-btn edit" data-action="edit" data-id="${item.id}" title="Edit Data Kontrak">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                      </button>
                      <button class="con-icon-btn delete" data-action="delete" data-id="${item.id}" title="Hapus Kontrak">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
                      </button>
                    ` : ''}
                  </div>
                </div>
              </div>
            `;
          }).join('')}
        </div>

      </div>
    `;

    // Event Listeners
    // 1. Search Input
    const searchInput = container.querySelector('#con-search-input');
    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        searchQuery = e.target.value;
        renderUI();
      });
    }

    // 2. Branch Filter
    const branchFilter = container.querySelector('#con-branch-filter');
    if (branchFilter) {
      branchFilter.addEventListener('change', (e) => {
        selectedBranch = e.target.value;
        renderUI();
      });
    }

    // 3. Division Filter
    const divisionFilter = container.querySelector('#con-division-filter');
    if (divisionFilter) {
      divisionFilter.addEventListener('change', (e) => {
        selectedDivision = e.target.value;
        renderUI();
      });
    }

    // 4. Expiring Days Filter
    const expiringFilter = container.querySelector('#con-expiring-filter');
    if (expiringFilter) {
      expiringFilter.addEventListener('change', (e) => {
        selectedExpiring = e.target.value;
        renderUI();
      });
    }

    // 5. Filter Pills
    container.querySelectorAll('.con-pill').forEach(pill => {
      pill.addEventListener('click', (e) => {
        const status = e.currentTarget.getAttribute('data-status');
        selectedStatus = status;
        if (status === 'Expiring30') {
          selectedExpiring = '30';
        } else {
          selectedExpiring = 'all';
        }
        renderUI();
      });
    });

    // 6. Reset Filters Button
    const resetBtn = container.querySelector('#con-reset-filter');
    if (resetBtn) {
      resetBtn.addEventListener('click', () => {
        searchQuery = '';
        selectedBranch = 'all';
        selectedDivision = 'all';
        selectedStatus = 'all';
        selectedExpiring = 'all';
        renderUI();
      });
    }

    // 7. Hero Actions
    const addBtn = container.querySelector('#con-btn-add');
    if (addBtn) addBtn.addEventListener('click', () => openFormModal());

    const excelBtn = container.querySelector('#con-btn-excel');
    if (excelBtn) excelBtn.addEventListener('click', exportToExcel);

    const missingBtn = container.querySelector('#con-btn-missing');
    if (missingBtn) missingBtn.addEventListener('click', checkMissingContracts);

    const refreshBtn = container.querySelector('#con-btn-refresh');
    if (refreshBtn) refreshBtn.addEventListener('click', () => loadData());

    // 8. Card Action Buttons (Edit / Delete)
    container.querySelectorAll('.con-icon-btn.edit').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const id = parseInt(e.currentTarget.getAttribute('data-id'), 10);
        const item = contractsList.find(c => c.id === id);
        if (item) openFormModal(item);
      });
    });

    container.querySelectorAll('.con-icon-btn.delete').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const id = parseInt(e.currentTarget.getAttribute('data-id'), 10);
        const item = contractsList.find(c => c.id === id);
        if (item) confirmDelete(item);
      });
    });
  }

  // Initial load
  await loadData();
}
