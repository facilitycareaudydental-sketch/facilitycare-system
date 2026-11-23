import { apiFetch, getUser, hasPermission } from '../config.js';
import { getCachedBranches, getCachedEmployees, fetchExportRows } from '../utils/dataCache.js';
import { createPagination } from '../components/table.js';
import { createModal, confirmDialog } from '../components/modal.js';
import { toastSuccess, toastError, toastInfo } from '../components/toast.js';
import { mountEmployeeCombobox, parseParticipantList } from '../components/employeeCombobox.js';
import { formatDateIndo } from '../utils/sorting.js';

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

const BATCH_LIST = ['Batch 1', 'Batch 2', 'Batch 3', 'Batch 4', 'Batch 5'];
const TRAINER_LIST = ['Berlin Ariansyah', 'Ade Surahman', 'Mizwar', 'Fajar', 'Ade', 'Berlin'];

// Standardized Indonesian date helper imported from ../utils/sorting.js (supports ISO, DD/MM/YYYY, Excel serial)

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
  let currentPage = 1;
  const PAGE_LIMIT = 20;
  let totalItems = 0;
  let totalPages = 1;
  let serverStats = {
    totalSessions: 0,
    totalBatches: 0,
    currentMonth: 0,
    totalParticipants: 0
  };
  let searchDebounce = null;
  let hasRendered = false;

  const currentUser = getUser();
  const canEdit = hasPermission(currentUser, 'training', 'write');
  const canDelete = hasPermission(currentUser, 'training', 'delete');

  // Load branches & employees cache
  try {
    [branchOptions, employeeOptions] = await Promise.all([
      getCachedBranches(),
      getCachedEmployees()
    ]);
  } catch (err) {
    console.error('Error fetching cache:', err);
  }

  function getActiveFilters() {
    const params = { sort: 'date_desc' };
    if (searchQuery && searchQuery.trim()) params.search = searchQuery.trim();
    if (selectedBranch && selectedBranch !== 'all') params.branch_id = selectedBranch;
    if (selectedBatch && selectedBatch !== 'all') params.batch = selectedBatch;
    if (selectedTrainer && selectedTrainer !== 'all') params.trainer = selectedTrainer;
    return params;
  }

  function updateStatsUI() {
    const set = (sel, val) => {
      const el = container.querySelector(sel);
      if (el) el.textContent = String(val ?? 0);
    };
    set('#tr-stat-sessions', serverStats.totalSessions);
    set('#tr-stat-participants', serverStats.totalParticipants);
    set('#tr-stat-batches', serverStats.totalBatches);
    set('#tr-stat-month', serverStats.currentMonth);
    set('#tr-pill-count-all', serverStats.totalSessions);
  }

  // Load Training Data from API
  async function loadData({ refreshStats = true } = {}) {
    isLoading = true;
    try {
      const filters = getActiveFilters();
      const params = new URLSearchParams({
        page: String(currentPage),
        limit: String(PAGE_LIMIT),
        ...filters
      });
      if (refreshStats) params.set('include_stats', '1');

      const res = await apiFetch('/api/training?' + params.toString());
      if (res.ok && res.data && Array.isArray(res.data.data)) {
        const pg = res.data.pagination || {};
        totalItems = Number(pg.total ?? res.data.data.length);
        totalPages = Math.max(1, Number(pg.pages || pg.totalPages || Math.ceil(totalItems / PAGE_LIMIT) || 1));
        if (res.data.data.length === 0 && currentPage > totalPages) {
          currentPage = totalPages;
          isLoading = false;
          return loadData({ refreshStats });
        }
        trainingList = res.data.data;
        if (res.data.stats) {
          serverStats = { ...serverStats, ...res.data.stats };
        }
      } else {
        trainingList = [];
        totalItems = 0;
        totalPages = 1;
        toastError(res.data?.error || 'Gagal memuat data training');
      }
    } catch (e) {
      console.warn('Gagal load training:', e);
      trainingList = [];
      totalItems = 0;
      totalPages = 1;
      toastError('Gagal memuat data training: ' + e.message);
    } finally {
      isLoading = false;
      if (!hasRendered) {
        hasRendered = true;
        renderUI();
      } else {
        if (refreshStats) updateStatsUI();
        renderListContent();
      }
    }
  }

  // Export to Excel
  async function exportToExcel() {
    if (!hasPermission(currentUser, 'export')) {
      toastError('Anda tidak memiliki izin untuk mengekspor data.');
      return;
    }
    try {
      const { downloadExcel } = await import('../utils/excel.js');
      const rows = await fetchExportRows('/api/training/export', getActiveFilters());
      const exportData = rows.map((t, idx) => ({
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
      toastSuccess(`Berhasil mengunduh ${rows.length} data training ke Excel!`);
    } catch (err) {
      toastError('Gagal ekspor Excel: ' + err.message);
    }
  }

  // Delete Training
  async function deleteTraining(id) {
    if (!canDelete) return;
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
            await loadData({ refreshStats: true });
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
    const defaultBatch = isEdit ? (item?.batch ? String(item.batch).trim() : '') : '';
    const hasLegacyBatch = isEdit && defaultBatch && !BATCH_LIST.includes(defaultBatch);
    const defaultTrainer = item?.trainer || 'Berlin Ariansyah';

    const content = document.createElement('div');
    content.innerHTML = `
      <div style="display:flex; flex-direction:column; gap:16px;">
        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Tanggal Training <span style="color:#DC2626;">*</span>
            </label>
            <input type="date" id="tr-modal-date" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${defaultDate}" required />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Batch Pelatihan <span style="color:#DC2626;">*</span>
            </label>
            <select id="tr-modal-batch" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" required>
              <option value="">-- Pilih Batch Pelatihan --</option>
              ${BATCH_LIST.map(b => `<option value="${b}" ${defaultBatch === b ? 'selected' : ''}>${b}</option>`).join('')}
              ${hasLegacyBatch ? `<option value="${defaultBatch.replace(/"/g, '&quot;')}" selected>${defaultBatch} (Historis)</option>` : ''}
            </select>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
            Materi / Topik Pelatihan <span style="color:#DC2626;">*</span>
          </label>
          <input type="text" id="tr-modal-subject" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
            value="${item?.subject ? String(item.subject).replace(/"/g, '&quot;') : ''}" required />
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
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
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Trainer / Instruktur <span style="color:#DC2626;">*</span>
            </label>
            <input list="tr-trainer-options" id="tr-modal-trainer" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${defaultTrainer}" required />
            <datalist id="tr-trainer-options">
              ${TRAINER_LIST.map(t => `<option value="${t}">`).join('')}
            </datalist>
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
            Nama Peserta Pelatihan <span style="color:#DC2626;">*</span>
          </label>
          <div id="tr-modal-participants-mount"></div>
        </div>

        <div class="form-row" style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Nilai / Score Evaluasi (0 - 100)
            </label>
            <input type="number" min="0" max="100" step="1" id="tr-modal-score" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${item?.score != null ? item.score : ''}" />
          </div>

          <div>
            <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
              Link Dokumen / Sertifikat (Opsional)
            </label>
            <input type="url" id="tr-modal-link" class="form-control" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;" 
              value="${item?.document_link ? String(item.document_link).replace(/"/g, '&quot;') : ''}" />
          </div>
        </div>

        <div>
          <label style="display:block; font-size:0.82rem; font-weight:600; color:var(--text-2); margin-bottom:5px;">
            Catatan Pelatihan
          </label>
          <textarea id="tr-modal-notes" class="form-control" rows="2" style="width:100%; padding:9px 12px; border:1px solid #CBD5E1; border-radius:8px;">${item?.notes || ''}</textarea>
        </div>
      </div>
    `;

    const participantsCombo = mountEmployeeCombobox(content.querySelector('#tr-modal-participants-mount'), {
      mode: 'multi',
      inputId: 'tr-modal-participants',
      inputName: 'participants',
      placeholder: '',
      allowManual: true,
      hasEmployeeIdSchema: false,
      employees: employeeOptions,
      initialParticipants: parseParticipantList(item?.participants),
    });

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
        const participants = (participantsCombo.getValue().participantsString || '').trim();
        const scoreVal = content.querySelector('#tr-modal-score').value;
        const score = scoreVal !== '' ? Number(scoreVal) : null;
        const document_link = (content.querySelector('#tr-modal-link').value || '').trim();
        const notes = (content.querySelector('#tr-modal-notes').value || '').trim();

        if (!training_date) {
          toastError('Tanggal training wajib diisi');
          return;
        }
        if (!batch || (!(isEdit && batch === defaultBatch) && !BATCH_LIST.includes(batch))) {
          toastError('Batch pelatihan wajib dipilih (Batch 1 - Batch 5)');
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
              close();
              await loadData({ refreshStats: true });
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
              currentPage = 1;
              close();
              await loadData({ refreshStats: true });
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

  // Server-filtered page data
  function getFilteredList() {
    return trainingList;
  }


  // Render only the dynamic list content HTML (Cards Grid / Empty State)
  function buildTrainingListHTML(filtered) {
    return `
    ${filtered.length === 0 ? `
            <div class="training-empty-state">
              <div style="font-size:3rem; margin-bottom:12px;">🔍</div>
              <h3 style="font-size:1.2rem; font-weight:700; color:#F8FAFC; margin:0 0 6px 0;">Tidak Ada Data Training Ditemukan</h3>
              <p style="color:#94A3B8; font-size:0.9rem; margin:0 0 20px 0;">Silakan sesuaikan kata kunci pencarian atau reset filter aktif.</p>
              <button id="tr-btn-reset-filter" class="btn btn-ghost" style="border:1px solid #223456; color:#CBD5E1; border-radius:8px; padding:8px 16px;">
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
                      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px; gap:8px; flex-wrap:wrap;">
                        <span class="training-batch-badge">
                          <span>🎓</span>
                          <span>${t.batch || 'Reguler'}</span>
                        </span>
                        <div style="display:flex; align-items:center; gap:8px;">
                          ${t.score != null ? `
                            <span class="training-score-badge">
                              <span>⭐</span>
                              <span>${t.score}</span>
                            </span>
                          ` : ''}
                          <span style="font-size:0.8rem; font-weight:600; color:#CBD5E1;">
                            🗓️ ${formatDateIndo(t.training_date)}
                          </span>
                        </div>
                      </div>

                      <!-- Subject / Materi Title -->
                      <h4 style="margin:0 0 12px 0; font-size:1.05rem; font-weight:700; color:#F8FAFC; line-height:1.4;">
                        ${t.subject || 'Materi Pelatihan'}
                      </h4>

                      <!-- Participant & Branch Info -->
                      <div style="display:flex; align-items:center; gap:12px; margin-bottom:10px;">
                        <div class="training-avatar-box">
                          ${initials}
                        </div>
                        <div style="min-width:0;">
                          <div style="font-size: 0.76rem; text-transform:uppercase; font-weight:700; color:#94A3B8; margin-bottom:2px; letter-spacing:0.04em;">Peserta Training</div>
                          <div style="font-size:0.95rem; font-weight:700; color:#F8FAFC; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${participantsText}">
                            ${participantsText}
                          </div>
                        </div>
                      </div>

                      <!-- Detail Box: Cabang & Trainer -->
                      <div class="training-info-box">
                        <div style="display:flex; align-items:center; gap:6px; font-size:0.83rem; color:#E2E8F0;">
                          <span>📍</span>
                          <span style="font-weight:600; color:#F8FAFC;">${t.branch_name || 'Cabang Belum Ditentukan'}</span>
                        </div>
                        <div style="display:flex; align-items:center; gap:6px; font-size:0.83rem; color:#CBD5E1;">
                          <span>👨‍🏫</span>
                          <span>Trainer: <strong style="color:#FFFFFF;">${t.trainer || '-'}</strong></span>
                        </div>
                      </div>

                      ${t.notes ? `
                        <div style="font-size:0.8rem; color:#CBD5E1; font-style:italic; line-height:1.4; margin-bottom:8px;">
                          📝 ${t.notes}
                        </div>
                      ` : ''}
                    </div>

                    <!-- Card Footer: Dokumen & Aksi -->
                    <div style="display:flex; justify-content:space-between; align-items:center; padding-top:12px; border-top:1px solid #223456; margin-top:8px;">
                      <div>
                        ${t.document_link ? `
                          <a href="${t.document_link}" target="_blank" rel="noopener noreferrer" class="btn btn-sm btn-ghost" style="border:1px solid #223456; color:#93C5FD; background:rgba(59,130,246,0.12); border-radius:6px; font-size:0.78rem; padding:4px 10px; text-decoration:none; display:inline-flex; align-items:center; gap:4px;">
                            <span>📄 Materi / Sertifikat</span>
                          </a>
                        ` : `
                          <span style="font-size:0.75rem; color:#94A3B8;">(Tanpa Dokumen)</span>
                        `}
                      </div>

                      ${canEdit || canDelete ? `
                        <div style="display:flex; align-items:center; gap:6px;">
                          ${canEdit ? `
                            <button class="training-icon-btn training-edit-btn" title="Edit Data Training" data-id="${t.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>
                            </button>
                          ` : ''}
                          ${canDelete ? `
                            <button class="training-icon-btn delete training-delete-btn" title="Hapus Training" data-id="${t.id}">
                              <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>
                            </button>
                          ` : ''}
                        </div>
                      ` : ''}
                    </div>

                  </div>
                `;
              }).join('')}
            </div>
          `}
    `;
  }

  function bindListEvents() {
    // Reset Filter Button inside list
    const btnReset = container.querySelector('#tr-btn-reset-filter');
    if (btnReset) {
      btnReset.onclick = () => {
        searchQuery = '';
        selectedBranch = 'all';
        selectedBatch = 'all';
        selectedTrainer = 'all';
        currentPage = 1;
        clearTimeout(searchDebounce);
        const searchInput = container.querySelector('#tr-search-input');
        if (searchInput) searchInput.value = '';
        const branchFilter = container.querySelector('#tr-filter-branch');
        if (branchFilter) branchFilter.value = 'all';
        const trainerFilter = container.querySelector('#tr-filter-trainer');
        if (trainerFilter) trainerFilter.value = 'all';
        container.querySelectorAll('.training-pill').forEach(p => {
          p.classList.toggle('active', p.getAttribute('data-batch') === 'all');
        });
        loadData({ refreshStats: false });
      };
    }

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

  function renderPagination() {
    const pgContainer = container.querySelector('#training-pagination-container');
    if (!pgContainer) return;
    pgContainer.innerHTML = '';
    const pgEl = createPagination({
      page: currentPage,
      pages: totalPages,
      total: totalItems,
      limit: PAGE_LIMIT,
      onPage: (newPage) => {
        currentPage = newPage;
        loadData({ refreshStats: false });
      }
    });
    if (pgEl) pgContainer.appendChild(pgEl);
  }

  function renderListContent() {
    const listContainer = container.querySelector('#training-list-container');
    if (!listContainer) return;
    const filtered = getFilteredList();
    listContainer.innerHTML = buildTrainingListHTML(filtered);
    renderPagination();
    bindListEvents();
  }

  // Main UI Renderer
  function renderUI() {
    const filtered = getFilteredList();

    // Global KPI Metrics from server stats
    const totalSessions = Number(serverStats.totalSessions ?? 0);
    const totalParticipants = Number(serverStats.totalParticipants ?? 0);
    const totalBatches = Number(serverStats.totalBatches ?? 0);
    const bulanIniCount = Number(serverStats.currentMonth ?? 0);

    container.innerHTML = `
      <style>
        .training-page {
          display: flex;
          flex-direction: column;
          gap: 20px;
        }
        .training-hero-card {
          background: #131F36;
          border: 1px solid #223456;
          border-radius: 16px;
          padding: 24px 28px;
          color: #FFFFFF;
          box-shadow: 0 4px 16px rgba(0, 0, 0, 0.22);
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
          background: radial-gradient(circle, rgba(37, 99, 235, 0.06) 0%, transparent 70%);
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
          background: #131F36;
          border: 1px solid #223456;
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
          background: #0E182B;
          color: #F8FAFC;
          border: 1px solid #223456;
          border-radius: 10px;
          font-size: 0.9rem;
          outline: none;
          transition: border-color 0.15s ease;
        }
        .training-search-input:focus {
          border-color: #3B82F6;
          box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.2);
        }
        .training-select {
          padding: 10px 14px;
          border: 1px solid #223456;
          border-radius: 10px;
          font-size: 0.9rem;
          color: #F8FAFC;
          background: #0E182B;
          outline: none;
          min-width: 180px;
          cursor: pointer;
        }
        .training-select:focus {
          border-color: #3B82F6;
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
          border: 1px solid #223456;
          background: #0E182B;
          color: #CBD5E1;
          transition: all 0.15s ease;
          display: inline-flex;
          align-items: center;
          gap: 6px;
          user-select: none;
          flex-shrink: 0;
        }
        .training-pill:hover {
          background: #22314A;
          color: #F8FAFC;
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
          background: #131F36;
          border: 1px solid #223456;
          border-radius: 14px;
          padding: 18px;
          box-shadow: 0 2px 6px rgba(0,0,0,0.2);
          transition: transform 0.15s ease, box-shadow 0.15s ease;
          display: flex;
          flex-direction: column;
          justify-content: space-between;
          position: relative;
        }
        .training-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 8px 18px rgba(0,0,0,0.32);
          border-color: #475569;
        }
        .training-avatar-box {
          width: 44px;
          height: 44px;
          border-radius: 10px;
          background: #2563EB;
          color: #FFFFFF;
          font-weight: 700;
          font-size: 1rem;
          display: flex;
          align-items: center;
          justify-content: center;
          flex-shrink: 0;
        }
        .training-batch-badge {
          display: inline-flex;
          align-items: center;
          gap: 5px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.76rem;
          font-weight: 700;
          background: rgba(37, 99, 235, 0.18);
          color: #93C5FD;
          border: 1px solid rgba(96, 165, 250, 0.35);
        }
        .training-score-badge {
          display: inline-flex;
          align-items: center;
          gap: 4px;
          padding: 4px 10px;
          border-radius: 20px;
          font-size: 0.76rem;
          font-weight: 700;
          background: rgba(16, 185, 129, 0.16);
          color: #6EE7B7;
          border: 1px solid rgba(52, 211, 153, 0.35);
        }
        .training-info-box {
          background: #0E182B;
          border: 1px solid #223456;
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
          border: 1px solid #223456;
          background: #0E182B;
          color: #CBD5E1;
          display: inline-flex;
          align-items: center;
          justify-content: center;
          cursor: pointer;
          transition: all 0.15s ease;
        }
        .training-icon-btn:hover {
          background: #22314A;
          color: #F8FAFC;
          border-color: #475569;
        }
        .training-icon-btn.delete:hover {
          background: rgba(239, 68, 68, 0.2);
          color: #FCA5A5;
          border-color: rgba(248, 113, 113, 0.45);
        }

        /* Empty state */
        .training-empty-state {
          grid-column: 1 / -1;
          background: #131F36;
          border: 2px dashed #223456;
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
            flex-wrap: wrap !important;
            overflow-x: visible !important;
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
              
              ${hasPermission(currentUser, "export") ? `<button id="tr-btn-excel" class="btn btn-ghost training-btn-excel" style="background:rgba(255,255,255,0.08); border:1px solid rgba(255,255,255,0.15); color:white; height:38px; padding:0 14px; border-radius:10px; font-weight:600; display:inline-flex; align-items:center; gap:6px; white-space:nowrap;">
                <span>📥</span>
                <span>Ekspor Excel</span>
              </button>` : ''}

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
                <div class="training-stat-val" id="tr-stat-sessions">${totalSessions}</div>
                <div class="training-stat-lbl">Sesi Training</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(59,130,246,0.2); color:#60A5FA;">
                👥
              </div>
              <div>
                <div class="training-stat-val" id="tr-stat-participants">${totalParticipants}</div>
                <div class="training-stat-lbl">Peserta Terlatih</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(16,185,129,0.2); color:#34D399;">
                📚
              </div>
              <div>
                <div class="training-stat-val"><span id="tr-stat-batches">${totalBatches}</span> Batch</div>
                <div class="training-stat-lbl">Total Angkatan</div>
              </div>
            </div>

            <div class="training-stat-box">
              <div class="training-stat-icon" style="background:rgba(245,158,11,0.2); color:#FBBF24;">
                🗓️
              </div>
              <div>
                <div class="training-stat-val" id="tr-stat-month">${bulanIniCount}</div>
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
              <input type="text" id="tr-search-input" class="training-search-input" placeholder="Cari materi, trainer, peserta, atau cabang..." value="${searchQuery ? String(searchQuery).replace(/"/g, '&quot;') : ''}" />
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
              <span class="badge" id="tr-pill-count-all" style="background:rgba(0,0,0,0.06); padding:2px 6px; border-radius:10px; font-size:0.75rem;">${totalSessions}</span>
            </div>
            ${BATCH_LIST.map(b => `
              <div class="training-pill ${selectedBatch === b ? 'active' : ''}" data-batch="${b}">
                <span>🎓</span>
                <span>${b}</span>
              </div>
            `).join('')}
          </div>
        </div>

        <!-- Content Area: Murni Grid Kartu -->
        <div id="training-list-container">
          ${buildTrainingListHTML(filtered)}
        </div>

        <!-- Pagination Container -->
        <div id="training-pagination-container"></div>

      </div>
    `;

    renderPagination();
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
    if (btnRefresh) btnRefresh.onclick = () => loadData({ refreshStats: true });

    // Search input (debounced)
    const searchInput = container.querySelector('#tr-search-input');
    if (searchInput) {
      searchInput.oninput = (e) => {
        searchQuery = e.target.value;
        clearTimeout(searchDebounce);
        searchDebounce = setTimeout(() => {
          currentPage = 1;
          loadData({ refreshStats: false });
        }, 350);
      };
    }

    // Branch filter
    const branchFilter = container.querySelector('#tr-filter-branch');
    if (branchFilter) {
      branchFilter.onchange = (e) => {
        selectedBranch = e.target.value;
        currentPage = 1;
        loadData({ refreshStats: false });
      };
    }

    // Trainer filter
    const trainerFilter = container.querySelector('#tr-filter-trainer');
    if (trainerFilter) {
      trainerFilter.onchange = (e) => {
        selectedTrainer = e.target.value;
        currentPage = 1;
        loadData({ refreshStats: false });
      };
    }

    // Batch Pills
    container.querySelectorAll('.training-pill').forEach(pill => {
      pill.onclick = () => {
        selectedBatch = pill.getAttribute('data-batch') || 'all';
        container.querySelectorAll('.training-pill').forEach(p => {
          p.classList.toggle('active', p.getAttribute('data-batch') === selectedBatch);
        });
        currentPage = 1;
        loadData({ refreshStats: false });
      };
    });

    // Initial bind of list events
    bindListEvents();
  }

  // Initial load
  await loadData({ refreshStats: true });
}
