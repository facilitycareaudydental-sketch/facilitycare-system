import { apiFetch, getUser } from '../config.js';
import { getCachedBranches, getCachedEmployeeNames } from '../utils/dataCache.js';
import { createModal, confirmDialog } from '../components/modal.js';
import { toastSuccess, toastError, toastInfo } from '../components/toast.js';

// Pre-seeded fallback data jika API offline / inisialisasi awal
const INITIAL_ISSUES = [
  {
    id: 1,
    report_date: '2026-08-27',
    branch_id: 19,
    branch_name: '019. Cinere',
    category: 'SDM',
    source: 'AM',
    complaint: 'Keluhan AM fc atas nama Malik bau badan, dan sudah dibilangin berulang kali terkait kebersihan seragam dan deodoran.',
    employee_name: 'MALIK',
    fc_specialist: 'Mizwar',
    solution: 'Diberikan teguran lisan, edukasi kebersihan diri, dan pendampingan seragam cadangan di klinik.',
    status: 'Done',
    completion_date: '2026-08-28',
    day_count: 1
  },
  {
    id: 2,
    report_date: '2026-08-26',
    branch_id: 0,
    branch_name: '000. Management',
    category: 'SDM',
    source: 'Berlin',
    complaint: 'Security atas nama Reza suhelmi meninggalkan area tugas jaga tanpa personil pengganti saat jam operasional klinik.',
    employee_name: 'REZA SUHELMI',
    fc_specialist: 'Ade Surahman',
    solution: 'Pemanggilan investigasi dan penerbitan Surat Peringatan (SP) sesuai SOP security.',
    status: 'Done',
    completion_date: '2026-08-27',
    day_count: 1
  },
  {
    id: 3,
    report_date: '2026-07-27',
    branch_id: 30,
    branch_name: '030. Ciputat',
    category: 'SDM',
    source: 'AM',
    complaint: 'Terinfo kembali permasalahan FC Ciputat, terkait standar pembersihan area dental unit yang belum memenuhi checklist hygiene.',
    employee_name: 'HENDRA',
    fc_specialist: 'Berlin Ariansyah',
    solution: 'Re-training modul pembersihan dental unit dan monitoring harian oleh supervisor.',
    status: 'In Progress',
    completion_date: null,
    day_count: null
  },
  {
    id: 4,
    report_date: '2026-06-24',
    branch_id: 30,
    branch_name: '030. Ciputat',
    category: 'SDM',
    source: 'AM',
    complaint: 'FC atas nama Hendra mengintimidasi rcp pada saat briefing pagi mengenai pembagian area kerja.',
    employee_name: 'HENDRA',
    fc_specialist: 'Berlin Ariansyah',
    solution: 'Mediasi bersama AM dan HR, diberikan pembinaan konseling perilaku kerja.',
    status: 'Done',
    completion_date: '2026-06-25',
    day_count: 1
  },
  {
    id: 5,
    report_date: '2026-06-05',
    branch_id: 0,
    branch_name: '000. Management',
    category: 'SDM',
    source: 'Berlin',
    complaint: 'Kinerja kurang dan ada komplain dari cabang mengenai ketidakhadiran berkala personil pengganti.',
    employee_name: 'Indra',
    fc_specialist: 'Ade',
    solution: 'Evaluasi penempatan dan pembinaan kedisiplinan kerja.',
    status: 'Done',
    completion_date: '2026-06-08',
    day_count: 3
  },
  {
    id: 6,
    report_date: '2026-06-05',
    branch_id: 0,
    branch_name: '000. Management',
    category: 'SDM',
    source: 'Ade',
    complaint: 'Sering tidak masuk kerja dan memalsukan surat dokter saat izin sakit.',
    employee_name: 'Kristian',
    fc_specialist: 'Ade Surahman',
    solution: 'Investigasi ke klinik dokter terkait dan eskalasi ke HRD untuk tindakan tegas.',
    status: 'Done',
    completion_date: '2026-06-07',
    day_count: 2
  },
  {
    id: 7,
    report_date: '2026-05-25',
    branch_id: 19,
    branch_name: '019. Cinere',
    category: 'SDM',
    source: 'SPV',
    complaint: 'Tidak masuk kerja tanpa kabar (alpa), dan sudah di-one on one oleh SPV terkait absensi.',
    employee_name: 'Malik',
    fc_specialist: 'Mizwar',
    solution: 'Pemberian SP 1 dan penandatanganan komitmen kehadiran kerja.',
    status: 'Done',
    completion_date: '2026-05-26',
    day_count: 1
  }
];

const CATEGORY_LIST = ['SDM', 'Cleaning', 'Aset', 'K3', 'Lainnya'];
const STATUS_LIST = ['Open', 'In Progress', 'Done'];
const SOURCE_LIST = ['SPV', 'AM', 'RCP', 'Perawat', 'FC', 'Security', 'Berlin', 'Ade', 'Pattrel', 'Dentrel'];

// Format tanggal lokal Indonesia
function formatDateIndo(dateStr) {
  if (!dateStr) return '-';
  try {
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return dateStr;
    return d.toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' });
  } catch {
    return dateStr;
  }
}

// Inisial avatar
function getInitials(name) {
  if (!name || name === '-') return '⚠️';
  const parts = String(name).trim().split(/\s+/);
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return (parts[0][0] + parts[1][0]).toUpperCase();
}

export async function renderIssues(container, params) {
  let issuesList = [];
  let branchOptions = [];
  let employeeOptions = [];
  let selectedStatus = 'all';
  let selectedCategory = 'all';
  let selectedBranch = 'all';
  let searchQuery = '';
  let isLoading = false;

  // Handle dashboard filter parameter
  const dashFilter = params ? params.get('dash_filter') : null;
  if (dashFilter === 'open') {
    selectedStatus = 'Open';
  }

  const currentUser = getUser();
  const canEdit = currentUser?.role !== 'viewer';

  // Load cache data
  try {
    [branchOptions, employeeOptions] = await Promise.all([
      getCachedBranches(),
      getCachedEmployeeNames()
    ]);
  } catch (err) {
    console.error('Error fetching cache:', err);
  }

  // Load Issues Data from API
  async function loadData() {
    isLoading = true;
    renderUI();
    try {
      const res = await apiFetch('/api/issues?limit=10000');
      if (res.ok) {
        const raw = res.data?.data || res.data || [];
        issuesList = Array.isArray(raw) && raw.length > 0 ? raw : INITIAL_ISSUES;
      } else {
        issuesList = INITIAL_ISSUES;
      }
    } catch (e) {
      console.warn('Gagal load issues, menggunakan data lokal:', e);
      issuesList = INITIAL_ISSUES;
    } finally {
      isLoading = false;
      renderUI();
    }
  }

  // Export to Excel
  async function exportToExcel() {
    try {
      const { downloadExcel } = await import('../utils/excel.js');
      const exportData = issuesList.map((item, idx) => ({
        'No': idx + 1,
        'Tanggal Info': item.report_date || '',
        'Cabang': item.branch_name || '',
        'Kategori': item.category || '',
        'Sumber Laporan': item.source || '',
        'Keluhan': item.complaint || '',
        'Nama FC / Security': item.employee_name || '',
        'FC Spesialis': item.fc_specialist || '',
        'Solusi / Tindak Lanjut': item.solution || '',
        'Status': item.status || '',
        'Tanggal Selesai': item.completion_date || '',
        'Durasi Hari': item.day_count != null ? item.day_count : ''
      }));
      downloadExcel(exportData, `Data_Permasalahan_FCMS_${new Date().toISOString().split('T')[0]}`);
      toastSuccess('Berhasil mengekspor data permasalahan ke Excel');
    } catch (err) {
      toastError('Gagal ekspor Excel: ' + err.message);
    }
  }

  // Quick Status Updater
  async function quickUpdateStatus(issue, newStatus) {
    try {
      const completion_date = newStatus === 'Done' 
        ? (issue.completion_date || new Date().toISOString().split('T')[0]) 
        : (newStatus === 'Open' ? null : issue.completion_date);

      let day_count = issue.day_count;
      if (issue.report_date && completion_date) {
        const d1 = new Date(issue.report_date);
        const d2 = new Date(completion_date);
        day_count = Math.max(0, Math.floor((d2 - d1) / (1000 * 60 * 60 * 24)));
      } else if (newStatus === 'Open') {
        day_count = null;
      }

      const payload = {
        ...issue,
        status: newStatus,
        completion_date,
        day_count
      };

      const res = await apiFetch(`/api/issues/${issue.id}`, {
        method: 'PUT',
        body: JSON.stringify(payload)
      });

      if (res.ok) {
        toastSuccess(`Status masalah diubah ke "${newStatus}"`);
        await loadData();
      } else {
        toastError(res.data?.error || 'Gagal memperbarui status');
      }
    } catch (err) {
      toastError('Terjadi kesalahan: ' + err.message);
    }
  }

  // Modal Tambah / Edit
  function openIssueModal(item = null) {
    const isEdit = !!item;
    const today = new Date().toISOString().split('T')[0];
    const defaultDate = item?.report_date || today;
    const defaultCategory = item?.category || 'SDM';
    const defaultSource = item?.source || 'AM';
    const defaultStatus = item?.status || 'Open';

    const content = document.createElement('div');
    content.className = 'issue-modal-body';
    content.innerHTML = `
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Info <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="iss-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${defaultDate}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Klinik <span style="color:#DC2626;">*</span>
            </label>
            <select id="iss-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${branchOptions.map(b => `
                <option value="${b.value}" ${String(item?.branch_id) === String(b.value) ? 'selected' : ''}>
                  ${b.label}
                </option>
              `).join('')}
            </select>
          </div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Kategori <span style="color:#DC2626;">*</span>
            </label>
            <select id="iss-modal-category" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${CATEGORY_LIST.map(cat => `
                <option value="${cat}" ${defaultCategory === cat ? 'selected' : ''}>${cat}</option>
              `).join('')}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Sumber Laporan
            </label>
            <input list="iss-source-options" id="iss-modal-source" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: AM, SPV, Berlin..." value="${defaultSource}" />
            <datalist id="iss-source-options">
              ${SOURCE_LIST.map(s => `<option value="${s}">`).join('')}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Detail Keluhan / Permasalahan <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="iss-modal-complaint" class="form-control" rows="3" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Jelaskan kendala, kronologi, atau komplain..." required>${item?.complaint || ''}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nama Personil FC / Security
            </label>
            <input list="iss-emp-options" id="iss-modal-employee" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Ketik nama karyawan terkait..." value="${item?.employee_name ? String(item.employee_name).replace(/"/g, '&quot;') : ''}" />
            <datalist id="iss-emp-options">
              ${employeeOptions.map(e => `<option value="${e.label}">`).join('')}
            </datalist>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              FC Spesialis / Penanggung Jawab
            </label>
            <input list="iss-spec-options" id="iss-modal-spec" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: Mizwar, Ade, Berlin..." value="${item?.fc_specialist ? String(item.fc_specialist).replace(/"/g, '&quot;') : ''}" />
            <datalist id="iss-spec-options">
              ${employeeOptions.map(e => `<option value="${e.label}">`).join('')}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Solusi / Tindak Lanjut Penanganan
          </label>
          <textarea id="iss-modal-solution" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Tindakan yang telah atau akan diambil untuk menyelesaikan masalah...">${item?.solution || ''}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr 1fr; gap:10px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Status Masalah <span style="color:#DC2626;">*</span>
            </label>
            <select id="iss-modal-status" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              ${STATUS_LIST.map(st => `
                <option value="${st}" ${defaultStatus === st ? 'selected' : ''}>${st}</option>
              `).join('')}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Selesai
            </label>
            <input type="date" id="iss-modal-completion" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${item?.completion_date || ''}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Durasi (Hari)
            </label>
            <input type="number" id="iss-modal-daycount" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px; background:#F8FAFC;" 
              value="${item?.day_count != null ? item.day_count : ''}" readonly placeholder="Otomatis" />
          </div>
        </div>
      </div>
    `;

    // Dynamic auto-calculate day_count
    const dateInput = content.querySelector('#iss-modal-date');
    const compInput = content.querySelector('#iss-modal-completion');
    const dayInput = content.querySelector('#iss-modal-daycount');
    const statusSelect = content.querySelector('#iss-modal-status');

    function updateDayCount() {
      if (dateInput.value && compInput.value) {
        const d1 = new Date(dateInput.value);
        const d2 = new Date(compInput.value);
        const diff = Math.floor((d2 - d1) / (1000 * 60 * 60 * 24));
        dayInput.value = !isNaN(diff) ? Math.max(0, diff) : '';
      } else {
        dayInput.value = '';
      }
    }

    dateInput.addEventListener('change', updateDayCount);
    compInput.addEventListener('change', updateDayCount);
    statusSelect.addEventListener('change', () => {
      if (statusSelect.value === 'Done' && !compInput.value) {
        compInput.value = new Date().toISOString().split('T')[0];
        updateDayCount();
      } else if (statusSelect.value === 'Open') {
        compInput.value = '';
        dayInput.value = '';
      }
    });

    createModal({
      title: isEdit ? '✏️ Edit Data Permasalahan' : '➕ Tambah Data Permasalahan',
      content,
      size: 'md',
      confirmText: isEdit ? 'Simpan Perubahan' : 'Simpan Permasalahan',
      confirmClass: 'btn-primary',
      onConfirm: async (_, close) => {
        const report_date = content.querySelector('#iss-modal-date').value;
        const branch_id = content.querySelector('#iss-modal-branch').value;
        const category = content.querySelector('#iss-modal-category').value;
        const source = (content.querySelector('#iss-modal-source').value || '').trim();
        const complaint = (content.querySelector('#iss-modal-complaint').value || '').trim();
        const employee_name = (content.querySelector('#iss-modal-employee').value || '').trim();
        const fc_specialist = (content.querySelector('#iss-modal-spec').value || '').trim();
        const solution = (content.querySelector('#iss-modal-solution').value || '').trim();
        const status = content.querySelector('#iss-modal-status').value;
        const completion_date = content.querySelector('#iss-modal-completion').value || null;

        if (!report_date || !branch_id || !complaint || !category) {
          toastError('Harap lengkapi Tanggal, Cabang, Kategori, dan Keluhan!');
          return;
        }

        let day_count = null;
        if (report_date && completion_date) {
          const d1 = new Date(report_date);
          const d2 = new Date(completion_date);
          day_count = Math.max(0, Math.floor((d2 - d1) / (1000 * 60 * 60 * 24)));
        }

        const payload = {
          report_date,
          branch_id: Number(branch_id),
          category,
          source,
          complaint,
          employee_name,
          fc_specialist,
          solution,
          status,
          completion_date,
          day_count
        };

        try {
          let res;
          if (isEdit) {
            res = await apiFetch(`/api/issues/${item.id}`, {
              method: 'PUT',
              body: JSON.stringify(payload)
            });
          } else {
            res = await apiFetch('/api/issues', {
              method: 'POST',
              body: JSON.stringify(payload)
            });
          }

          if (res.ok) {
            toastSuccess(isEdit ? 'Permasalahan berhasil diperbarui!' : 'Permasalahan baru berhasil ditambahkan!');
            close();
            await loadData();
          } else {
            toastError(res.data?.error || 'Gagal menyimpan permasalahan');
          }
        } catch (err) {
          toastError('Terjadi kesalahan: ' + err.message);
        }
      }
    });
  }

  // Delete Issue
  async function deleteIssue(item) {
    confirmDialog({
      title: 'Hapus Permasalahan?',
      message: `Apakah Anda yakin ingin menghapus tiket permasalahan dari <strong>${item.branch_name || 'cabang'}</strong>:<br><br><em>"${item.complaint}"</em>?`,
      confirmText: 'Ya, Hapus',
      confirmClass: 'btn-danger',
      onConfirm: async () => {
        try {
          const res = await apiFetch(`/api/issues/${item.id}`, {
            method: 'DELETE'
          });
          if (res.ok) {
            toastSuccess('Permasalahan berhasil dihapus');
            await loadData();
          } else {
            toastError(res.data?.error || 'Gagal menghapus data');
          }
        } catch (err) {
          toastError('Terjadi kesalahan: ' + err.message);
        }
      }
    });
  }

  // Filter Data
  function getFilteredList() {
    return issuesList.filter(item => {
      // Search
      if (searchQuery) {
        const q = searchQuery.toLowerCase();
        const matchComplaint = (item.complaint || '').toLowerCase().includes(q);
        const matchEmp = (item.employee_name || '').toLowerCase().includes(q);
        const matchBranch = (item.branch_name || '').toLowerCase().includes(q);
        const matchSource = (item.source || '').toLowerCase().includes(q);
        const matchCategory = (item.category || '').toLowerCase().includes(q);
        const matchSol = (item.solution || '').toLowerCase().includes(q);
        if (!matchComplaint && !matchEmp && !matchBranch && !matchSource && !matchCategory && !matchSol) {
          return false;
        }
      }

      // Status
      if (selectedStatus !== 'all' && item.status !== selectedStatus) {
        return false;
      }

      // Category
      if (selectedCategory !== 'all' && item.category !== selectedCategory) {
        return false;
      }

      // Branch
      if (selectedBranch !== 'all' && String(item.branch_id) !== String(selectedBranch)) {
        return false;
      }

      return true;
    });
  }

  // Main UI Renderer
  function renderUI() {
    const filtered = getFilteredList();

    // KPI Metrics calculation
    const totalIssues = issuesList.length;
    const openAndProgress = issuesList.filter(i => i.status === 'Open' || i.status === 'In Progress').length;
    const doneCount = issuesList.filter(i => i.status === 'Done').length;

    // Masalah Bulan Ini
    const now = new Date();
    const curYear = now.getFullYear();
    const curMonth = now.getMonth();
    const bulanIniCount = issuesList.filter(i => {
      if (!i.report_date) return false;
      const d = new Date(i.report_date);
      return !isNaN(d.getTime()) && d.getFullYear() === curYear && d.getMonth() === curMonth;
    }).length;

    // Counts for status pills
    const countAll = issuesList.length;
    const countOpen = issuesList.filter(i => i.status === 'Open').length;
    const countInProgress = issuesList.filter(i => i.status === 'In Progress').length;
    const countDone = issuesList.filter(i => i.status === 'Done').length;

    container.innerHTML = `
      <style>
        .issues-page {
          display: flex;
          flex-direction: column;
          gap: 20px;
        }
        .issues-hero-card {
          background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
          border-radius: 16px;
          padding: 24px 28px;
          color: white;
          box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25);
          position: relative;
          overflow: hidden;
        }
        .issues-hero-card::after {
          content: '';
          position: absolute;
          top: -40px;
          right: -40px;
          width: 220px;
          height: 220px;
          background: radial-gradient(circle, rgba(245, 158, 11, 0.18) 0%, transparent 70%);
          border-radius: 50%;
          pointer-events: none;
        }
        .issues-stats-row {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 14px;
          margin-top: 20px;
        }
        .issues-stat-box {
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
        .issues-stat-box:hover {
          background: rgba(255, 255, 255, 0.08);
          transform: translateY(-2px);
        }
        .issues-stat-icon {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 1.3rem;
          flex-shrink: 0;
        }
        .issues-stat-val {
          font-size: 1.55rem;
          font-weight: 800;
          line-height: 1.1;
          letter-spacing: -0.02em;
        }
        .issues-stat-lbl {
          font-size: 0.72rem;
          text-transform: uppercase;
          color: #94A3B8;
          font-weight: 700;
          letter-spacing: 0.05em;
          margin-top: 2px;
        }

        /* Controls / Filter Bar */
        .issues-controls-card {
          background: white;
          border-radius: 14px;
          border: 1px solid #E2E8F0;
          padding: 16px 20px;
          box-shadow: 0 2px 8px rgba(0,0,0,0.03);
          display: flex;
          flex-direction: column;
          gap: 14px;
        }
        .issues-search-row {
          display: flex;
          gap: 12px;
          flex-wrap: wrap;
        }
        .issues-search-wrap {
          position: relative;
          flex: 1;
          min-width: 240px;
        }
        .issues-search-icon {
          position: absolute;
          left: 14px;
          top: 50%;
          transform: translateY(-50%);
          color: #94A3B8;
          pointer-events: none;
        }
        .issues-search-input {
          width: 100%;
          padding: 10px 14px 10px 40px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          outline: none;
          transition: border-color 0.15s ease;
        }
        .issues-search-input:focus {
          border-color: #2563EB;
          box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);
        }
        .issues-select {
          padding: 10px 14px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          color: #1E293B;
          background: white;
          outline: none;
          min-width: 160px;
          cursor: pointer;
        }
        .issues-select:focus {
          border-color: #2563EB;
        }

        /* Swipeable Pills */
        .issues-filter-pills {
          display: flex;
          align-items: center;
          gap: 8px;
          overflow-x: auto;
          white-space: nowrap;
          scrollbar-width: none;
          padding-bottom: 2px;
        }
        .issues-filter-pills::-webkit-scrollbar {
          display: none;
        }
        .issues-pill {
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
        .issues-pill:hover {
          background: #F1F5F9;
          color: #1E293B;
        }
        .issues-pill.active {
          background: #0F172A;
          color: white;
          border-color: #0F172A;
        }
        .issues-pill-dot {
          width: 8px;
          height: 8px;
          border-radius: 50%;
        }
        .issues-pill-count {
          padding: 1px 6px;
          border-radius: 10px;
          font-size: 0.72rem;
          background: rgba(0,0,0,0.06);
        }
        .issues-pill.active .issues-pill-count {
          background: rgba(255,255,255,0.2);
          color: white;
        }

        /* Cards Grid */
        .issues-cards-grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(370px, 1fr));
          gap: 16px;
        }
        .issues-card {
          background: white;
          border-radius: 14px;
          border: 1px solid #E2E8F0;
          padding: 18px;
          box-shadow: 0 2px 8px rgba(0,0,0,0.04);
          display: flex;
          flex-direction: column;
          gap: 14px;
          transition: transform 0.15s ease, box-shadow 0.15s ease;
        }
        .issues-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 6px 16px rgba(0,0,0,0.08);
          border-color: #CBD5E1;
        }
        .issues-card-top {
          display: flex;
          justify-content: space-between;
          align-items: center;
          gap: 8px;
        }
        .issues-category-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.76rem;
          font-weight: 700;
        }
        .issues-category-sdm {
          background: #EEF2FF;
          color: #4F46E5;
          border: 1px solid #E0E7FF;
        }
        .issues-category-cleaning {
          background: #ECFEFF;
          color: #0891B2;
          border: 1px solid #CFFAFE;
        }
        .issues-category-aset {
          background: #F5F3FF;
          color: #7C3AED;
          border: 1px solid #EDE9FE;
        }
        .issues-category-k3 {
          background: #FEF3C7;
          color: #D97706;
          border: 1px solid #FDE68A;
        }
        .issues-category-lainnya {
          background: #F1F5F9;
          color: #475569;
          border: 1px solid #E2E8F0;
        }

        .issues-status-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.76rem;
          font-weight: 700;
        }
        .issues-status-open {
          background: #FEF3C7;
          color: #B45309;
          border: 1px solid #FDE68A;
        }
        .issues-status-inprogress {
          background: #DBEAFE;
          color: #1D4ED8;
          border: 1px solid #BFDBFE;
        }
        .issues-status-done {
          background: #DCFCE7;
          color: #15803D;
          border: 1px solid #BBF7D0;
        }

        /* Complaint Quote Box */
        .issues-complaint-box {
          background: #F8FAFC;
          border-left: 4px solid #F59E0B;
          border-radius: 0 10px 10px 0;
          padding: 12px 14px;
          font-size: 0.92rem;
          color: #0F172A;
          line-height: 1.5;
          font-weight: 600;
        }
        .issues-complaint-box.status-done {
          border-left-color: #10B981;
        }
        .issues-complaint-box.status-inprogress {
          border-left-color: #3B82F6;
        }

        /* Person / Location Box */
        .issues-person-box {
          background: #F8FAFC;
          border: 1px solid #F1F5F9;
          border-radius: 10px;
          padding: 12px;
          display: flex;
          align-items: center;
          gap: 12px;
        }
        .issues-avatar-box {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          background: #4F46E5;
          color: white;
          display: flex;
          align-items: center;
          justify-content: center;
          font-weight: 800;
          font-size: 1rem;
          flex-shrink: 0;
        }
        .issues-person-name {
          font-size: 0.95rem;
          font-weight: 800;
          color: #0F172A;
        }
        .issues-person-sub {
          font-size: 0.74rem;
          color: #64748B;
          font-weight: 600;
          text-transform: uppercase;
          letter-spacing: 0.04em;
        }

        /* Details list */
        .issues-details-grid {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 8px 12px;
          font-size: 0.82rem;
          color: #334155;
          padding: 2px 4px;
        }
        .issues-detail-item {
          display: flex;
          align-items: center;
          gap: 6px;
        }

        /* Solution box */
        .issues-solution-box {
          background: #F0FDF4;
          border: 1px solid #DCFCE7;
          border-radius: 10px;
          padding: 10px 12px;
          font-size: 0.82rem;
          color: #166534;
          line-height: 1.45;
        }
        .issues-solution-title {
          font-weight: 700;
          display: flex;
          align-items: center;
          gap: 6px;
          margin-bottom: 3px;
        }

        /* Card footer */
        .issues-card-footer {
          display: flex;
          justify-content: space-between;
          align-items: center;
          border-top: 1px solid #F1F5F9;
          padding-top: 12px;
          margin-top: auto;
          gap: 8px;
        }
        .issues-action-btns {
          display: flex;
          align-items: center;
          gap: 6px;
        }
        .issues-icon-btn {
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
        .issues-icon-btn:hover {
          background: #F1F5F9;
          color: #0F172A;
          border-color: #CBD5E1;
        }
        .issues-icon-btn.delete:hover {
          background: #FEE2E2;
          color: #DC2626;
          border-color: #FCA5A5;
        }
        .issues-quick-btn {
          font-size: 0.76rem;
          font-weight: 700;
          padding: 6px 10px;
          border-radius: 8px;
          cursor: pointer;
          border: 1px solid transparent;
          display: inline-flex;
          align-items: center;
          gap: 4px;
          transition: all 0.15s ease;
        }
        .issues-quick-btn.start {
          background: #EFF6FF;
          color: #2563EB;
          border-color: #BFDBFE;
        }
        .issues-quick-btn.start:hover {
          background: #DBEAFE;
        }
        .issues-quick-btn.complete {
          background: #F0FDF4;
          color: #16A34A;
          border-color: #BBF7D0;
        }
        .issues-quick-btn.complete:hover {
          background: #DCFCE7;
        }
        .issues-quick-btn.reopen {
          background: #FFFBEB;
          color: #D97706;
          border-color: #FDE68A;
        }
        .issues-quick-btn.reopen:hover {
          background: #FEF3C7;
        }

        /* Empty state */
        .issues-empty-state {
          grid-column: 1 / -1;
          background: white;
          border: 2px dashed #E2E8F0;
          border-radius: 16px;
          padding: 48px 24px;
          text-align: center;
        }

        /* Mobile Responsive */
        @media (max-width: 768px) {
          .issues-hero-card {
            padding: 18px 16px;
          }
          .issues-stats-row {
            grid-template-columns: 1fr 1fr;
            gap: 10px;
          }
          .issues-stat-box {
            padding: 10px 12px;
          }
          .issues-stat-icon {
            width: 38px;
            height: 38px;
            font-size: 1.1rem;
          }
          .issues-stat-val {
            font-size: 1.3rem;
          }
          .issues-hero-actions {
            width: 100% !important;
            display: flex !important;
            align-items: center !important;
            gap: 8px !important;
            flex-wrap: nowrap !important;
          }
          .issues-btn-add {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.82rem !important;
            white-space: nowrap !important;
          }
          .issues-btn-excel {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.82rem !important;
            white-space: nowrap !important;
          }
          .issues-btn-refresh {
            width: 38px !important;
            height: 38px !important;
            padding: 0 !important;
            flex-shrink: 0 !important;
            justify-content: center !important;
          }
          .issues-controls-card {
            padding: 14px;
            gap: 12px;
          }
          .issues-search-row {
            flex-direction: column;
            gap: 10px;
          }
          .issues-select {
            width: 100%;
          }
          .issues-cards-grid {
            grid-template-columns: minmax(0, 1fr) !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
          }
          .issues-card {
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
            overflow: hidden;
            padding: 14px !important;
          }
        }
      </style>

      <div class="issues-page">

        <!-- Hero Header Card -->
        <div class="issues-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px; margin-bottom:6px;">
                <span style="font-size:1.6rem;">⚠️</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Log Permasalahan Operasional
                </h1>
              </div>
              <p style="margin:0; font-size:0.88rem; color:#CBD5E1; max-width:650px; line-height:1.5;">
                Pusat monitoring keluhan operasional klinik, kendala personil FC & Security, serta rekaman investigasi dan progres penyelesaian masalah.
              </p>
            </div>

            <div class="issues-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${canEdit ? `
                <button id="issues-btn-add" class="btn btn-primary issues-btn-add" style="background:#2563EB; padding:9px 16px; border-radius:10px; font-weight:700; display:flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Permasalahan</span>
                </button>
              ` : ''}
              
              <button id="issues-btn-excel" class="btn btn-ghost issues-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; padding:9px 14px; border-radius:10px; font-weight:600; display:flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>📥</span>
                <span>Ekspor Excel</span>
              </button>

              <button id="issues-btn-refresh" class="btn btn-ghost issues-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2.2" viewBox="0 0 24 24" class="${isLoading ? 'spin-icon' : ''}"><path d="M23 4v6h-6"/><path d="M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- 4 KPI Stat Boxes -->
          <div class="issues-stats-row">
            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(239, 68, 68, 0.2); color:#FCA5A5;">⚠️</div>
              <div>
                <div class="issues-stat-val">${totalIssues}</div>
                <div class="issues-stat-lbl">Total Permasalahan</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(245, 158, 11, 0.2); color:#FCD34D;">⏳</div>
              <div>
                <div class="issues-stat-val">${openAndProgress}</div>
                <div class="issues-stat-lbl">Open & In Progress</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(16, 185, 129, 0.2); color:#6EE7B7;">✅</div>
              <div>
                <div class="issues-stat-val">${doneCount}</div>
                <div class="issues-stat-lbl">Tuntas (Done)</div>
              </div>
            </div>

            <div class="issues-stat-box">
              <div class="issues-stat-icon" style="background:rgba(59, 130, 246, 0.2); color:#93C5FD;">📅</div>
              <div>
                <div class="issues-stat-val">${bulanIniCount}</div>
                <div class="issues-stat-lbl">Bulan Berjalan</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Filter Controls Card -->
        <div class="issues-controls-card">
          <div class="issues-search-row">
            <div class="issues-search-wrap">
              <svg class="issues-search-icon" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
              <input type="text" id="issues-search-input" class="issues-search-input" 
                placeholder="Cari keluhan, nama personil, sumber, cabang..." value="${searchQuery ? String(searchQuery).replace(/"/g, '&quot;') : ''}" />
            </div>

            <select id="issues-branch-filter" class="issues-select">
              <option value="all">Semua Cabang (${branchOptions.length})</option>
              ${branchOptions.map(b => `
                <option value="${b.value}" ${String(selectedBranch) === String(b.value) ? 'selected' : ''}>
                  ${b.label}
                </option>
              `).join('')}
            </select>

            <select id="issues-category-filter" class="issues-select">
              <option value="all">Semua Kategori</option>
              ${CATEGORY_LIST.map(cat => `
                <option value="${cat}" ${selectedCategory === cat ? 'selected' : ''}>
                  Kategori: ${cat}
                </option>
              `).join('')}
            </select>
          </div>

          <!-- Status Horizontal Swipeable Pills -->
          <div class="issues-filter-pills">
            <div class="issues-pill ${selectedStatus === 'all' ? 'active' : ''}" data-status="all">
              <span>Semua Status</span>
              <span class="issues-pill-count">${countAll}</span>
            </div>
            <div class="issues-pill ${selectedStatus === 'Open' ? 'active' : ''}" data-status="Open">
              <span class="issues-pill-dot" style="background:#F59E0B;"></span>
              <span>Open</span>
              <span class="issues-pill-count">${countOpen}</span>
            </div>
            <div class="issues-pill ${selectedStatus === 'In Progress' ? 'active' : ''}" data-status="In Progress">
              <span class="issues-pill-dot" style="background:#3B82F6;"></span>
              <span>In Progress</span>
              <span class="issues-pill-count">${countInProgress}</span>
            </div>
            <div class="issues-pill ${selectedStatus === 'Done' ? 'active' : ''}" data-status="Done">
              <span class="issues-pill-dot" style="background:#10B981;"></span>
              <span>Done</span>
              <span class="issues-pill-count">${countDone}</span>
            </div>
          </div>
        </div>

        <!-- Exclusive Cards Grid -->
        <div class="issues-cards-grid">
          ${filtered.length === 0 ? `
            <div class="issues-empty-state">
              <div style="font-size:2.8rem; margin-bottom:10px;">🔍</div>
              <h3 style="font-size:1.15rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak ada permasalahan yang cocok</h3>
              <p style="font-size:0.85rem; color:#64748B; margin:0 0 16px 0;">Coba sesuaikan kata kunci pencarian atau filter status yang dipilih.</p>
              <button id="issues-reset-filter" class="btn btn-secondary" style="font-size:0.82rem; padding:7px 16px; border-radius:8px;">Reset Semua Filter</button>
            </div>
          ` : filtered.map(item => {
            const cat = item.category || 'Lainnya';
            const catClass = `issues-category-${cat.toLowerCase().replace(/[^a-z0-9]/g, '')}` || 'issues-category-lainnya';
            
            const st = item.status || 'Open';
            let stClass = 'issues-status-open';
            let stText = '⚠️ Open';
            if (st === 'In Progress') {
              stClass = 'issues-status-inprogress';
              stText = '🔄 In Progress';
            } else if (st === 'Done') {
              stClass = 'issues-status-done';
              stText = '✅ Done';
            }

            const compBoxStatus = st === 'Done' ? 'status-done' : (st === 'In Progress' ? 'status-inprogress' : '');

            return `
              <div class="issues-card" data-id="${item.id}">
                <!-- Top Badges Row -->
                <div class="issues-card-top">
                  <span class="issues-category-badge ${catClass}">
                    🏷️ ${cat}
                  </span>
                  <div style="display:flex; align-items:center; gap:8px;">
                    <span class="issues-status-badge ${stClass}">
                      ${stText}
                    </span>
                    <span style="font-size:0.8rem; color:#64748B; display:flex; align-items:center; gap:4px;">
                      🗓️ ${formatDateIndo(item.report_date)}
                    </span>
                  </div>
                </div>

                <!-- Complaint Box -->
                <div class="issues-complaint-box ${compBoxStatus}">
                  "${item.complaint || '-'}"
                </div>

                <!-- Person & Location Box -->
                <div class="issues-person-box">
                  <div class="issues-avatar-box">
                    ${getInitials(item.employee_name)}
                  </div>
                  <div style="min-width:0; flex:1;">
                    <div class="issues-person-sub">PERSONIL FC / SECURITY</div>
                    <div class="issues-person-name" style="white-space:nowrap; overflow:hidden; text-overflow:ellipsis;">
                      ${item.employee_name || 'Tidak ada personil spesifik'}
                    </div>
                  </div>
                </div>

                <!-- Details Grid -->
                <div class="issues-details-grid">
                  <div class="issues-detail-item" style="grid-column: 1 / -1;">
                    <span style="color:#64748B;">📍 Cabang:</span>
                    <strong style="color:#0F172A;">${item.branch_name || 'Cabang Tidak Diketahui'}</strong>
                  </div>
                  <div class="issues-detail-item">
                    <span style="color:#64748B;">🗣️ Sumber:</span>
                    <strong>${item.source || '-'}</strong>
                  </div>
                  <div class="issues-detail-item">
                    <span style="color:#64748B;">🛡️ Spesialis:</span>
                    <strong>${item.fc_specialist || '-'}</strong>
                  </div>
                </div>

                <!-- Solution & Completion Box (if available) -->
                ${item.solution || item.completion_date ? `
                  <div class="issues-solution-box">
                    ${item.solution ? `
                      <div class="issues-solution-title">
                        <span>💡 Tindak Lanjut / Solusi:</span>
                      </div>
                      <div style="margin-bottom:${item.completion_date ? '6px' : '0'};">
                        ${item.solution}
                      </div>
                    ` : ''}
                    ${item.completion_date ? `
                      <div style="font-size:0.76rem; color:#15803D; font-weight:600; display:flex; align-items:center; gap:4px;">
                        <span>⏱️ Selesai: ${formatDateIndo(item.completion_date)}</span>
                        ${item.day_count != null ? `<span>(Durasi: ${item.day_count} Hari)</span>` : ''}
                      </div>
                    ` : ''}
                  </div>
                ` : ''}

                <!-- Footer Actions -->
                <div class="issues-card-footer">
                  <div>
                    ${canEdit ? (
                      st === 'Open' ? `
                        <button class="issues-quick-btn start" data-action="quick-progress" data-id="${item.id}" title="Ubah status ke In Progress">
                          <span>🔄 Mulai Tangani</span>
                        </button>
                      ` : st === 'In Progress' ? `
                        <button class="issues-quick-btn complete" data-action="quick-done" data-id="${item.id}" title="Tandai selesai">
                          <span>✅ Selesaikan</span>
                        </button>
                      ` : `
                        <button class="issues-quick-btn reopen" data-action="quick-reopen" data-id="${item.id}" title="Buka kembali tiket masalah">
                          <span>↩️ Buka Kembali</span>
                        </button>
                      `
                    ) : ''}
                  </div>

                  <div class="issues-action-btns">
                    ${canEdit ? `
                      <button class="issues-icon-btn edit" data-action="edit" data-id="${item.id}" title="Edit Masalah">
                        <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                      </button>
                      <button class="issues-icon-btn delete" data-action="delete" data-id="${item.id}" title="Hapus Masalah">
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

    // ── Event Handlers ──
    const searchInput = container.querySelector('#issues-search-input');
    if (searchInput) {
      searchInput.addEventListener('input', (e) => {
        searchQuery = e.target.value;
        renderUI();
        // Keep focus and cursor
        const el = container.querySelector('#issues-search-input');
        if (el) {
          el.focus();
          el.setSelectionRange(el.value.length, el.value.length);
        }
      });
    }

    const branchFilter = container.querySelector('#issues-branch-filter');
    if (branchFilter) {
      branchFilter.addEventListener('change', (e) => {
        selectedBranch = e.target.value;
        renderUI();
      });
    }

    const categoryFilter = container.querySelector('#issues-category-filter');
    if (categoryFilter) {
      categoryFilter.addEventListener('change', (e) => {
        selectedCategory = e.target.value;
        renderUI();
      });
    }

    // Status Pills
    container.querySelectorAll('.issues-pill').forEach(pill => {
      pill.addEventListener('click', () => {
        selectedStatus = pill.dataset.status;
        renderUI();
      });
    });

    // Reset button
    const resetBtn = container.querySelector('#issues-reset-filter');
    if (resetBtn) {
      resetBtn.addEventListener('click', () => {
        searchQuery = '';
        selectedBranch = 'all';
        selectedCategory = 'all';
        selectedStatus = 'all';
        renderUI();
      });
    }

    // Top action buttons
    const addBtn = container.querySelector('#issues-btn-add');
    if (addBtn) {
      addBtn.addEventListener('click', () => openIssueModal());
    }

    const excelBtn = container.querySelector('#issues-btn-excel');
    if (excelBtn) {
      excelBtn.addEventListener('click', () => exportToExcel());
    }

    const refreshBtn = container.querySelector('#issues-btn-refresh');
    if (refreshBtn) {
      refreshBtn.addEventListener('click', () => loadData());
    }

    // Card Action Buttons Delegation
    container.querySelectorAll('.issues-card').forEach(card => {
      const id = Number(card.dataset.id);
      const item = issuesList.find(x => x.id === id);
      if (!item) return;

      const editBtn = card.querySelector('[data-action="edit"]');
      if (editBtn) {
        editBtn.addEventListener('click', () => openIssueModal(item));
      }

      const delBtn = card.querySelector('[data-action="delete"]');
      if (delBtn) {
        delBtn.addEventListener('click', () => deleteIssue(item));
      }

      const progBtn = card.querySelector('[data-action="quick-progress"]');
      if (progBtn) {
        progBtn.addEventListener('click', () => quickUpdateStatus(item, 'In Progress'));
      }

      const doneBtn = card.querySelector('[data-action="quick-done"]');
      if (doneBtn) {
        doneBtn.addEventListener('click', () => quickUpdateStatus(item, 'Done'));
      }

      const reopenBtn = card.querySelector('[data-action="quick-reopen"]');
      if (reopenBtn) {
        reopenBtn.addEventListener('click', () => quickUpdateStatus(item, 'Open'));
      }
    });
  }

  // Initial load
  await loadData();
}
