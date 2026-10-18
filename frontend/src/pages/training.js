import { apiFetch, getUser } from '../config.js';
import { getCachedBranches, getCachedEmployeeNames } from '../utils/dataCache.js';
import { createModal, confirmDialog } from '../components/modal.js';
import { toastSuccess, toastError, toastInfo } from '../components/toast.js';

// Pre-seeded fallback data training jika API offline / inisialisasi awal
const INITIAL_TRAINING = [
  { id: 1, training_date: '2026-06-29', batch: 'Batch 5', subject: 'Basic Knowledge & Practical Skills', branch_id: 6, branch_name: '006. Kelapa Gading MOI', trainer: 'Berlin Ariansyah', participants: 'MUHAMMAD PRAKAS AGAZI', score: 90, document_link: '' },
  { id: 2, training_date: '2026-06-29', batch: 'Batch 5', subject: 'Basic Knowledge & Practical Skills', branch_id: 11, branch_name: '011. Kuningan', trainer: 'Berlin Ariansyah', participants: 'KAMALLUDIN', score: 88, document_link: '' },
  { id: 3, training_date: '2026-06-29', batch: 'Batch 5', subject: 'Basic Knowledge & Practical Skills', branch_id: 42, branch_name: '042. Muara Karang', trainer: 'Berlin Ariansyah', participants: 'SAHRUL ICHSANI', score: 85, document_link: '' },
  { id: 4, training_date: '2026-06-29', batch: 'Batch 5', subject: 'Basic Knowledge & Practical Skills', branch_id: 51, branch_name: '051. Bintaro Sektor 9', trainer: 'Berlin Ariansyah', participants: 'MUHAMAD VICKRI FATHURROHMAN', score: 92, document_link: '' },
  { id: 5, training_date: '2026-06-29', batch: 'Batch 5', subject: 'Basic Knowledge & Practical Skills', branch_id: 45, branch_name: '045. FX Mall', trainer: 'Berlin Ariansyah', participants: 'DEDE AHDORI', score: 86, document_link: '' },
  { id: 6, training_date: '2026-06-29', batch: 'Batch 5', subject: 'Basic Knowledge & Practical Skills', branch_id: 54, branch_name: '054. Bogor Gunung Batu', trainer: 'Berlin Ariansyah', participants: 'ADITYA PRATAMA', score: 95, document_link: '' },
  { id: 7, training_date: '2026-06-29', batch: 'Batch 5', subject: 'Basic Knowledge & Practical Skills', branch_id: 55, branch_name: '055. Gajah Mada', trainer: 'Berlin Ariansyah', participants: 'ELVAN JANGGA PRAMANA', score: 89, document_link: '' }
];

const BATCH_LIST = ['Batch 5', 'Batch 4', 'Batch 3', 'Batch 2', 'Batch 1'];
const TRAINER_LIST = ['Berlin Ariansyah', 'Ade Surahman', 'Mizwar', 'Fajar', 'Ade', 'Berlin'];

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

// Helper ekstraksi nama peserta
function formatParticipants(raw) {
  if (!raw) return '-';
  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) return parsed.join(', ');
  } catch {}
  return String(raw);
}

export async function renderTraining(container) {
  let trainingList = [];
  let branchOptions = [];
  let employeeOptions = [];
  let selectedBatch = 'all';
  let selectedBranch = 'all';
  let selectedTrainer = 'all';
  let searchQuery = '';
  let isLoading = false;

  const currentUser = getUser();
  const canEdit = currentUser?.role !== 'viewer';

  // Load branches & employees cache
  try {
    [branchOptions, employeeOptions] = await Promise.all([
      getCachedBranches(),
      getCachedEmployeeNames()
    ]);
  } catch (err) {
    console.error('Error fetching cache:', err);
  }

  // Load Training Data from API
  async function loadData() {
    isLoading = true;
    renderUI();
    try {
      const res = await apiFetch('/api/training?limit=10000');
      if (res.ok) {
        const raw = res.data?.data || res.data || [];
        trainingList = Array.isArray(raw) && raw.length > 0 ? raw : INITIAL_TRAINING;
      } else {
        trainingList = INITIAL_TRAINING;
      }
    } catch (e) {
      console.warn('Gagal load training, menggunakan data lokal:', e);
      trainingList = INITIAL_TRAINING;
    } finally {
      isLoading = false;
      renderUI();
    }
  }

  // Export to Excel
  async function exportToExcel() {
    try {
      const { downloadExcel } = await import('../utils/excel.js');
      const exportData = trainingList.map((t, idx) => ({
        'No': idx + 1,
        'Tanggal': t.training_date || '',
        'Batch': t.batch || '',
        'Materi': t.subject || '',
        'Cabang': t.branch_name || '',
        'Trainer': t.trainer || '',
        'Peserta': formatParticipants(t.participants),
        'Nilai': t.score != null ? t.score : '',
        'Dokumen': t.document_link || '',
        'Catatan': t.notes || ''
      }));
      downloadExcel(exportData, `Data_Training_${new Date().toISOString().slice(0, 10)}`);
      toastSuccess('Data training berhasil diekspor ke Excel!');
    } catch (err) {
      toastError('Gagal ekspor Excel: ' + err.message);
    }
  }

  // Delete Training
  async function deleteTraining(id) {
    if (!canEdit) return;
    const item = trainingList.find(t => t.id === id);
    confirmDialog({
      title: 'Hapus Data Training',
      message: `Apakah Anda yakin ingin menghapus data sesi training "${item?.subject || 'ini'}" untuk ${formatParticipants(item?.participants)}?`,
      confirmText: 'Ya, Hapus',
      confirmClass: 'btn-danger',
      onConfirm: async () => {
        try {
          const res = await apiFetch(`/api/training/${id}`, { method: 'DELETE' });
          if (res.ok) {
            toastSuccess('Data training berhasil dihapus');
            trainingList = trainingList.filter(t => t.id !== id);
            renderUI();
          } else {
            toastError(res.data?.error || 'Gagal menghapus data training');
          }
        } catch (e) {
          toastError('Gagal menghapus training: ' + e.message);
        }
      }
    });
  }

  // Open Form Modal (Add / Edit)
  function openTrainingModal(item = null) {
    if (!canEdit) return;
    const isEdit = !!item;
    const defaultDate = item?.training_date ? item.training_date.slice(0, 10) : new Date().toISOString().slice(0, 10);
    const defaultBatch = item?.batch || 'Batch 5';
    const defaultTrainer = item?.trainer || 'Berlin Ariansyah';
    const participantsText = formatParticipants(item?.participants) !== '-' ? formatParticipants(item?.participants) : '';

    const content = document.createElement('div');
    content.innerHTML = `
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Tanggal Training <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="tr-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${defaultDate}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Batch Pelatihan <span style="color:#DC2626;">*</span>
            </label>
            <input list="tr-batch-options" id="tr-modal-batch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Pilih atau ketik batch..." value="${defaultBatch}" required />
            <datalist id="tr-batch-options">
              ${BATCH_LIST.map(b => `<option value="${b}">`).join('')}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Materi / Topik Pelatihan <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="tr-modal-subject" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Contoh: Basic Knowledge & Practical Skills" value="${item?.subject ? String(item.subject).replace(/"/g, '&quot;') : ''}" required />
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Cabang Klinik <span style="color:#DC2626;">*</span>
            </label>
            <select id="tr-modal-branch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Cabang --</option>
              ${branchOptions.map(b => `
                <option value="${b.value}" ${String(item?.branch_id) === String(b.value) ? 'selected' : ''}>
                  ${b.label}
                </option>
              `).join('')}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Trainer / Instruktur <span style="color:#DC2626;">*</span>
            </label>
            <input list="tr-trainer-options" id="tr-modal-trainer" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Pilih trainer..." value="${defaultTrainer}" required />
            <datalist id="tr-trainer-options">
              ${TRAINER_LIST.map(t => `<option value="${t}">`).join('')}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Nama Peserta Pelatihan <span style="color:#DC2626;">*</span>
          </label>
          <textarea id="tr-modal-participants" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Ketik nama karyawan peserta training..." required>${participantsText}</textarea>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Nilai / Score Evaluasi (0 - 100)
            </label>
            <input type="number" min="0" max="100" step="1" id="tr-modal-score" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="Contoh: 90" value="${item?.score != null ? item.score : ''}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
              Link Dokumen / Sertifikat (Opsional)
            </label>
            <input type="url" id="tr-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              placeholder="https://drive.google.com/..." value="${item?.document_link ? String(item.document_link).replace(/"/g, '&quot;') : ''}" />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:#475569; margin-bottom:5px;">
            Catatan Pelatihan
          </label>
          <textarea id="tr-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            placeholder="Catatan hasil pelatihan, rekomendasi materi lanjutan...">${item?.notes || ''}</textarea>
        </div>
      </div>
    `;

    createModal({
      title: isEdit ? '✏️ Edit Data Sesi Training' : '➕ Tambah Data Sesi Training',
      content,
      size: 'md',
      confirmText: isEdit ? 'Simpan Perubahan' : 'Simpan Training',
      confirmClass: 'btn-primary',
      onConfirm: async (_, close) => {
        const training_date = content.querySelector('#tr-modal-date').value;
        const batch = (content.querySelector('#tr-modal-batch').value || '').trim();
        const subject = (content.querySelector('#tr-modal-subject').value || '').trim();
        const branch_id = content.querySelector('#tr-modal-branch').value;
        const trainer = (content.querySelector('#tr-modal-trainer').value || '').trim();
        const participants = (content.querySelector('#tr-modal-participants').value || '').trim();
        const scoreVal = content.querySelector('#tr-modal-score').value;
        const score = scoreVal !== '' ? Number(scoreVal) : null;
        const document_link = (content.querySelector('#tr-modal-link').value || '').trim();
        const notes = (content.querySelector('#tr-modal-notes').value || '').trim();

        if (!training_date) {
          toastError('Tanggal training wajib diisi');
          return;
        }
        if (!subject) {
          toastError('Materi / topik pelatihan wajib diisi');
          return;
        }
        if (!branch_id) {
          toastError('Cabang klinik wajib dipilih');
          return;
        }
        if (!trainer) {
          toastError('Trainer wajib diisi');
          return;
        }
        if (!participants) {
          toastError('Nama peserta training wajib diisi');
          return;
        }

        const bObj = branchOptions.find(b => String(b.value) === String(branch_id));

        const payload = {
          training_date,
          batch,
          subject,
          branch_id: Number(branch_id),
          branch_name: bObj ? bObj.label : 'Cabang',
          trainer,
          participants,
          score,
          document_link,
          notes
        };

        try {
          if (isEdit) {
            const res = await apiFetch(`/api/training/${item.id}`, {
              method: 'PUT',
              body: JSON.stringify(payload)
            });
            if (res.ok) {
              toastSuccess('Data training berhasil diperbarui');
              const idx = trainingList.findIndex(t => t.id === item.id);
              if (idx !== -1) trainingList[idx] = { ...trainingList[idx], ...payload };
              close();
              renderUI();
            } else {
              toastError(res.data?.error || 'Gagal memperbarui data training');
            }
          } else {
            const res = await apiFetch('/api/training', {
              method: 'POST',
              body: JSON.stringify(payload)
            });
            if (res.ok) {
              toastSuccess('Sesi training baru berhasil ditambahkan');
              const newId = res.data?.id || Date.now();
              trainingList.unshift({ id: newId, ...payload });
              close();
              renderUI();
            } else {
              toastError(res.data?.error || 'Gagal menambahkan training');
            }
          }
        } catch (err) {
          toastError('Terjadi kesalahan: ' + err.message);
        }
      }
    });
  }

  // Filter & Search Logic
  function getFilteredList() {
    return trainingList.filter(t => {
      // Search
      if (searchQuery) {
        const q = searchQuery.toLowerCase();
        const matchSubj = (t.subject || '').toLowerCase().includes(q);
        const matchTrainer = (t.trainer || '').toLowerCase().includes(q);
        const matchPart = formatParticipants(t.participants).toLowerCase().includes(q);
        const matchBranch = (t.branch_name || '').toLowerCase().includes(q);
        const matchBatch = (t.batch || '').toLowerCase().includes(q);
        if (!matchSubj && !matchTrainer && !matchPart && !matchBranch && !matchBatch) return false;
      }

      // Batch Filter
      if (selectedBatch !== 'all' && t.batch !== selectedBatch) {
        return false;
      }

      // Branch Filter
      if (selectedBranch !== 'all' && String(t.branch_id) !== String(selectedBranch)) {
        return false;
      }

      // Trainer Filter
      if (selectedTrainer !== 'all' && t.trainer !== selectedTrainer) {
        return false;
      }

      return true;
    });
  }

  // Main UI Renderer
  function renderUI() {
    const filtered = getFilteredList();

    // KPI Metrics calculation
    const totalSessions = trainingList.length;

    // Unique participants
    const participantSet = new Set();
    trainingList.forEach(t => {
      const p = formatParticipants(t.participants);
      p.split(',').forEach(name => {
        const trimmed = name.trim();
        if (trimmed && trimmed !== '-') participantSet.add(trimmed);
      });
    });
    const totalParticipants = participantSet.size;

    // Unique Batches
    const batchSet = new Set(trainingList.map(t => t.batch).filter(Boolean));
    const totalBatches = batchSet.size;

    // Sesi Bulan Berjalan
    const now = new Date();
    const curYear = now.getFullYear();
    const curMonth = now.getMonth();
    const bulanIniCount = trainingList.filter(t => {
      if (!t.training_date) return false;
      const d = new Date(t.training_date);
      return !isNaN(d.getTime()) && d.getFullYear() === curYear && d.getMonth() === curMonth;
    }).length;

    container.innerHTML = `
      <style>
        .training-page {
          display: flex;
          flex-direction: column;
          gap: 20px;
        }
        .training-hero-card {
          background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%);
          border-radius: 16px;
          padding: 24px 28px;
          color: white;
          box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25);
          position: relative;
          overflow: hidden;
        }
        .training-hero-card::after {
          content: '';
          position: absolute;
          top: -40px;
          right: -40px;
          width: 220px;
          height: 220px;
          background: radial-gradient(circle, rgba(147, 51, 234, 0.18) 0%, transparent 70%);
          border-radius: 50%;
          pointer-events: none;
        }
        .training-hero-actions {
          display: flex;
          align-items: center;
          gap: 8px;
          flex-wrap: wrap;
        }
        .training-hero-actions .btn {
          height: 38px;
          box-sizing: border-box;
          transition: all 0.2s ease;
        }
        .training-btn-add:hover {
          background: #1D4ED8 !important;
          transform: translateY(-1px);
        }
        .training-btn-excel:hover, .training-btn-refresh:hover {
          background: rgba(255, 255, 255, 0.16) !important;
          border-color: rgba(255, 255, 255, 0.3) !important;
          transform: translateY(-1px);
        }
        .training-stats-row {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 14px;
          margin-top: 22px;
        }
        .training-stat-box {
          background: rgba(255, 255, 255, 0.06);
          border: 1px solid rgba(255, 255, 255, 0.1);
          border-radius: 12px;
          padding: 14px 16px;
          backdrop-filter: blur(8px);
          display: flex;
          align-items: center;
          gap: 14px;
        }
        .training-stat-icon {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          display: flex;
          align-items: center;
          justify-content: center;
          font-size: 1.3rem;
          flex-shrink: 0;
        }
        .training-stat-val {
          font-size: 1.55rem;
          font-weight: 800;
          line-height: 1.1;
          color: #F8FAFC;
        }
        .training-stat-lbl {
          font-size: 0.76rem;
          text-transform: uppercase;
          letter-spacing: 0.05em;
          color: #94A3B8;
          font-weight: 600;
          margin-top: 2px;
        }

        /* Controls Card */
        .training-controls-card {
          background: white;
          border: 1px solid #E2E8F0;
          border-radius: 14px;
          padding: 18px 20px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.02);
          display: flex;
          flex-direction: column;
          gap: 16px;
        }
        .training-search-row {
          display: flex;
          gap: 12px;
          flex-wrap: wrap;
        }
        .training-search-wrap {
          position: relative;
          flex: 1;
          min-width: 240px;
        }
        .training-search-icon {
          position: absolute;
          left: 14px;
          top: 50%;
          transform: translateY(-50%);
          color: #94A3B8;
          pointer-events: none;
        }
        .training-search-input {
          width: 100%;
          padding: 10px 14px 10px 40px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          outline: none;
          transition: border-color 0.15s ease;
        }
        .training-search-input:focus {
          border-color: #2563EB;
          box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1);
        }
        .training-select {
          padding: 10px 14px;
          border: 1px solid #CBD5E1;
          border-radius: 10px;
          font-size: 0.9rem;
          color: #1E293B;
          background: white;
          outline: none;
          min-width: 180px;
          cursor: pointer;
        }
        .training-select:focus {
          border-color: #2563EB;
        }

        /* Swipeable Pills */
        .training-filter-pills {
          display: flex;
          align-items: center;
          gap: 8px;
          overflow-x: auto;
          white-space: nowrap;
          scrollbar-width: none;
          padding-bottom: 2px;
        }
        .training-filter-pills::-webkit-scrollbar {
          display: none;
        }
        .training-pill {
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
        .training-pill:hover {
          background: #F1F5F9;
          color: #1E293B;
        }
        .training-pill.active {
          background: #2563EB;
          color: white;
          border-color: #2563EB;
          box-shadow: 0 2px 6px rgba(37, 99, 235, 0.25);
        }

        /* Grid Cards View */
        .training-cards-grid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
          gap: 16px;
        }
        .training-card {
          background: white;
          border: 1px solid #E2E8F0;
          border-radius: 14px;
          padding: 18px;
          box-shadow: 0 2px 6px rgba(0,0,0,0.03);
          transition: transform 0.15s ease, box-shadow 0.15s ease;
          display: flex;
          flex-direction: column;
          justify-content: space-between;
          position: relative;
        }
        .training-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 8px 18px rgba(0,0,0,0.06);
          border-color: #CBD5E1;
        }
        .training-avatar-box {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          background: linear-gradient(135deg, #7C3AED 0%, #6D28D9 100%);
          color: white;
          font-weight: 700;
          font-size: 1rem;
          display: flex;
          align-items: center;
          justify-content: center;
          flex-shrink: 0;
          box-shadow: 0 3px 6px rgba(124, 58, 237, 0.25);
        }
        .training-batch-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.76rem;
          font-weight: 700;
          background: #F3E8FF;
          color: #7E22CE;
          border: 1px solid #E9D5FF;
        }
        .training-score-badge {
          display: inline-flex;
          align-items: center;
          gap: 4px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.76rem;
          font-weight: 700;
          background: #DCFCE7;
          color: #15803D;
          border: 1px solid #BBF7D0;
        }
        .training-info-box {
          background: #F8FAFC;
          border: 1px solid #E2E8F0;
          border-radius: 10px;
          padding: 10px 12px;
          margin: 12px 0;
          display: flex;
          flex-direction: column;
          gap: 6px;
        }
        .training-icon-btn {
          width: 32px;
          height: 32px;
          border-radius: 8px;
          border: 1px solid #E2E8F0;
          background: white;
          color: #64748B;
          display: inline-flex;
          align-items: center;
          justify-content: center;
          cursor: pointer;
          transition: all 0.15s ease;
        }
        .training-icon-btn:hover {
          background: #F1F5F9;
          color: #0F172A;
          border-color: #CBD5E1;
        }
        .training-icon-btn.delete:hover {
          background: #FEE2E2;
          color: #DC2626;
          border-color: #FCA5A5;
        }

        /* Empty state */
        .training-empty-state {
          grid-column: 1 / -1;
          background: white;
          border: 2px dashed #E2E8F0;
          border-radius: 16px;
          padding: 48px 24px;
          text-align: center;
        }

        /* Mobile Responsive */
        @media (max-width: 768px) {
          .training-hero-card {
            padding: 18px 16px;
          }
          .training-stats-row {
            grid-template-columns: 1fr 1fr;
            gap: 10px;
          }
          .training-stat-box {
            padding: 10px 12px;
          }
          .training-stat-icon {
            width: 38px;
            height: 38px;
            font-size: 1.1rem;
          }
          .training-stat-val {
            font-size: 1.3rem;
          }
          .training-hero-actions {
            width: 100% !important;
            display: flex !important;
            align-items: center !important;
            gap: 6px !important;
            flex-wrap: nowrap !important;
            overflow-x: auto !important;
            scrollbar-width: none !important;
            padding-bottom: 2px !important;
          }
          .training-hero-actions::-webkit-scrollbar {
            display: none;
          }
          .training-btn-add {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.82rem !important;
            white-space: nowrap !important;
          }
          .training-btn-excel {
            flex: 1 1 auto !important;
            justify-content: center !important;
            padding: 8px 10px !important;
            font-size: 0.82rem !important;
            white-space: nowrap !important;
          }
          .training-btn-refresh {
            width: 38px !important;
            height: 38px !important;
            padding: 0 !important;
            flex-shrink: 0 !important;
            justify-content: center !important;
          }
          .training-controls-card {
            padding: 14px;
            gap: 12px;
          }
          .training-search-row {
            flex-direction: column;
            gap: 10px;
          }
          .training-select {
            width: 100%;
          }
          .training-cards-grid {
            grid-template-columns: minmax(0, 1fr) !important;
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
          }
          .training-card {
            width: 100% !important;
            max-width: 100% !important;
            min-width: 0 !important;
            box-sizing: border-box !important;
            overflow: hidden;
            padding: 14px !important;
          }
        }
      </style>

      <div class="training-page">

        <!-- Hero Header Card -->
        <div class="training-hero-card">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px;">
            <div>
              <div style="display:flex; align-items:center; gap:10px;">
                <span style="font-size:1.6rem;">🎓</span>
                <h1 style="font-size:1.45rem; font-weight:800; margin:0; letter-spacing:-0.02em;">
                  Monitoring Training Karyawan
                </h1>
              </div>
            </div>

            <div class="training-hero-actions" style="display:flex; align-items:center; gap:8px;">
              ${canEdit ? `
                <button id="tr-btn-add" class="btn btn-primary training-btn-add" style="background:#2563EB; height:38px; padding:0 16px; border-radius:10px; font-weight:700; display:inline-flex; align-items:center; gap:6px; box-shadow:0 4px 12px rgba(37,99,235,0.3); white-space:nowrap;">
                  <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2.5" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>
                  <span>Tambah Training</span>
                </button>
              ` : ''}
              
              <button id="tr-btn-excel" class="btn btn-ghost training-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>📥</span>
                <span>Ekspor Excel</span>
              </button>

              <button id="tr-btn-refresh" class="btn btn-ghost training-btn-refresh" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; width:38px; height:38px; padding:0; border-radius:10px; display:inline-flex; align-items:center; justify-content:center; flex-shrink:0;" title="Muat Ulang Data">
                <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M23 4v6h-6M1 20v-6h6"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></svg>
              </button>
            </div>
          </div>

          <!-- KPI Stats -->
          <div class="training-stats-row">
            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(124,58,237,0.25); color:#A78BFA;">
                🎓
              </div>
              <div>
                <div class="training-stat-val">${totalSessions}</div>
                <div class="training-stat-lbl">Sesi Training</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">
                👥
              </div>
              <div>
                <div class="training-stat-val">${totalParticipants}</div>
                <div class="training-stat-lbl">Peserta Terlatih</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">
                📚
              </div>
              <div>
                <div class="training-stat-val">${totalBatches} Batch</div>
                <div class="training-stat-lbl">Total Angkatan</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                🗓️
              </div>
              <div>
                <div class="training-stat-val">${bulanIniCount}</div>
                <div class="training-stat-lbl">Bulan Berjalan</div>
              </div>
            </div>
          </div>
        </div>

        <!-- Filter & Search Controls -->
        <div class="training-controls-card">
          <div class="training-search-row">
            <div class="training-search-wrap">
              <span class="training-search-icon">
                <svg width="17" height="17" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
              </span>
              <input type="text" id="tr-search-input" class="training-search-input" placeholder="Cari materi, trainer, peserta, atau cabang..." value="${searchQuery}" />
            </div>

            <select id="tr-filter-branch" class="training-select">
              <option value="all">Semua Cabang (${branchOptions.length})</option>
              ${branchOptions.map(b => `<option value="${b.value}" ${String(selectedBranch) === String(b.value) ? 'selected' : ''}>${b.label}</option>`).join('')}
            </select>

            <select id="tr-filter-trainer" class="training-select">
              <option value="all">Semua Trainer</option>
              ${TRAINER_LIST.map(t => `<option value="${t}" ${selectedTrainer === t ? 'selected' : ''}>${t}</option>`).join('')}
            </select>
          </div>

          <!-- Batch Filter Pills (Swipeable on Mobile) -->
          <div class="training-filter-pills">
            <div class="training-pill ${selectedBatch === 'all' ? 'active' : ''}" data-batch="all">
              <span>Semua Batch</span>
              <span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${totalSessions}</span>
            </div>
            ${BATCH_LIST.map(b => {
              const count = trainingList.filter(t => t.batch === b).length;
              return `
                <div class="training-pill ${selectedBatch === b ? 'active' : ''}" data-batch="${b}">
                  <span>🎓</span>
                  <span>${b}</span>
                  ${count > 0 ? `<span class="badge" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${count}</span>` : ''}
                </div>
              `;
            }).join('')}
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="training-list-container">
          ${filtered.length === 0 ? `
            <div class="training-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">🔍</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#1E293B; margin:0 0 6px 0;">Tidak Ada Data Training Ditemukan</h3>
              <p style="color:#64748B; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="tr-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #CBD5E1; border-radius:8px; padding:8px 16px;">
                Reset Filter
              </button>
            </div>
          ` : `
            <div class="training-cards-grid">
              ${filtered.map(t => {
                const participantsText = formatParticipants(t.participants);
                const firstParticipant = participantsText.split(',')[0].trim();
                const initials = firstParticipant.split(' ').map(w => w[0]).filter(Boolean).slice(0, 2).join('').toUpperCase() || 'TR';
                return `
                  <div class="training-card" data-id="${t.id}">
                    <div>
                      <!-- Card Header: Batch & Nilai/Tanggal -->
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                        <span class="training-batch-badge">
                          <span>🎓</span>
                          <span>${t.batch || 'Reguler'}</span>
                        </span>
                        <div style="display:flex; align-items:center; gap:6px;">
                          ${t.score != null ? `
                            <span class="training-score-badge">
                              <span>⭐</span>
                              <span>${t.score}</span>
                            </span>
                          ` : ''}
                          <span style="font-size:0.8rem; font-weight:600; color:#64748B;">
                            🗓️ ${formatDateIndo(t.training_date)}
                          </span>
                        </div>
                      </div>

                      <!-- Subject / Materi Title -->
                      <h4 style="margin:0 0 10px 0; font-size:1.05rem; font-weight:700; color:#0F172A; line-height:1.4;">
                        ${t.subject || 'Materi Pelatihan'}
                      </h4>

                      <!-- Participant & Branch Info -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:10px;">
                        <div class="training-avatar-box">
                          ${initials}
                        </div>
                        <div style="min-width:0;">
                          <div style="font-size:0.72rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:1px;">Peserta Training</div>
                          <div style="font-size:0.95rem; font-weight:700; color:#1E293B; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${participantsText}">
                            ${participantsText}
                          </div>
                        </div>
                      </div>

                      <!-- Detail Box: Cabang & Trainer -->
                      <div class="training-info-box">
                        <div style="display:flex; align-items:center; gap:6px; font-size:0.82rem; color:#334155;">
                          <span>📍</span>
                          <span style="font-weight:600;">${t.branch_name || 'Cabang Belum Ditentukan'}</span>
                        </div>
                        <div style="display:flex; align-items:center; gap:6px; font-size:0.82rem; color:#64748B;">
                          <span>👨‍🏫</span>
                          <span>Trainer: <strong style="color:#0F172A;">${t.trainer || '-'}</strong></span>
                        </div>
                      </div>

                      ${t.notes ? `
                        <div style="font-size:0.78rem; color:#64748B; font-style:italic; line-height:1.4; margin-bottom:8px;">
                          📝 ${t.notes}
                        </div>
                      ` : ''}
                    </div>

                    <!-- Card Footer: Dokumen & Aksi -->
                    <div style="display:flex; justify-content:space-between; align-items:center; padding-top:12px; border-top:1px solid #F1F5F9; margin-top:8px;">
                      <div>
                        ${t.document_link ? `
                          <a href="${t.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #CBD5E1; border-radius:6px; font-size:0.78rem; padding:4px 8px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                            <span>📄 Materi / Sertifikat</span>
                          </a>
                        ` : `
                          <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                        `}
                      </div>

                      ${canEdit ? `
                        <div style="display:flex; align-items:center; gap:6px;">
                          <button class="training-icon-btn training-edit-btn" title="Edit Data Training" data-id="${t.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                          </button>
                          <button class="training-icon-btn delete training-delete-btn" title="Hapus Training" data-id="${t.id}">
                            <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                          </button>
                        </div>
                      ` : ''}
                    </div>

                  </div>
                `;
              }).join('')}
            </div>
          `}
        </div>

      </div>
    `;

    bindEvents();
  }

  // Event Listeners Binding
  function bindEvents() {
    // Add button
    const btnAdd = container.querySelector('#tr-btn-add');
    if (btnAdd) btnAdd.onclick = () => openTrainingModal();

    // Excel button
    const btnExcel = container.querySelector('#tr-btn-excel');
    if (btnExcel) btnExcel.onclick = () => exportToExcel();

    // Refresh button
    const btnRefresh = container.querySelector('#tr-btn-refresh');
    if (btnRefresh) btnRefresh.onclick = () => loadData();

    // Reset filter
    const btnReset = container.querySelector('#tr-btn-reset-filter');
    if (btnReset) {
      btnReset.onclick = () => {
        searchQuery = '';
        selectedBranch = 'all';
        selectedBatch = 'all';
        selectedTrainer = 'all';
        renderUI();
      };
    }

    // Search input
    const searchInput = container.querySelector('#tr-search-input');
    if (searchInput) {
      searchInput.oninput = (e) => {
        searchQuery = e.target.value;
        renderUI();
        const newSearch = container.querySelector('#tr-search-input');
        if (newSearch) {
          newSearch.focus();
          newSearch.setSelectionRange(newSearch.value.length, newSearch.value.length);
        }
      };
    }

    // Branch filter
    const branchFilter = container.querySelector('#tr-filter-branch');
    if (branchFilter) {
      branchFilter.onchange = (e) => {
        selectedBranch = e.target.value;
        renderUI();
      };
    }

    // Trainer filter
    const trainerFilter = container.querySelector('#tr-filter-trainer');
    if (trainerFilter) {
      trainerFilter.onchange = (e) => {
        selectedTrainer = e.target.value;
        renderUI();
      };
    }

    // Batch Pills
    container.querySelectorAll('.training-pill').forEach(pill => {
      pill.onclick = () => {
        selectedBatch = pill.getAttribute('data-batch') || 'all';
        renderUI();
      };
    });

    // Edit Buttons
    container.querySelectorAll('.training-edit-btn').forEach(btn => {
      btn.onclick = (e) => {
        e.stopPropagation();
        const id = Number(btn.getAttribute('data-id'));
        const item = trainingList.find(t => t.id === id);
        if (item) openTrainingModal(item);
      };
    });

    // Delete Buttons
    container.querySelectorAll('.training-delete-btn').forEach(btn => {
      btn.onclick = (e) => {
        e.stopPropagation();
        const id = Number(btn.getAttribute('data-id'));
        deleteTraining(id);
      };
    });
  }

  // Initial load
  await loadData();
}
