/**
 * import.js — Halaman Settings → Import Data Awal
 * Alur: Upload → Validasi lokal → Preview → Pilih duplikat → Import → Ringkasan
 */
import { apiFetch, getUser } from '../config.js';
import { toastSuccess, toastError } from '../components/toast.js';
import { validateWorkbook, generateTemplate, downloadErrorLog, SHEET_MAP } from '../utils/importer.js';
import {
  clearBatchRecoveryState,
  clearImportSession,
  computeImportFingerprint,
  escapeHtml,
  getBatchRecoveryState,
  getIncompleteImportSessions,
  renderImportStatusBannerHTML,
  resumeImportSession,
  saveBatchRecoveryState,
  submitJournaledImport,
} from '../utils/import_client.js';

// Import order: validation referensi first, then employees, etc.
// Reliever and Permintaan Chemical (supply) are excluded from FCMS Baru migration.
const IMPORT_ORDER = [
  'validation', 'employees', 'contracts', 'schedule', 'issues',
  'training', 'checklist', 'sop',
  'inspection', 'cleaning', 'fogging', 'basecamp',
  'sp', 'mutasi', 'forms',
];

export function renderImportPage(container) {
  const currentUser = getUser();
  if (!currentUser || String(currentUser.role).toLowerCase() !== 'superadmin' || currentUser.isRestrictedToken === true || Number(currentUser.must_change_password) === 1) {
    container.innerHTML = '<div class="empty-state"><p class="text-danger">Akses ditolak. Halaman ini hanya untuk Superadmin.</p></div>';
    return;
  }
  container.innerHTML = `
    <div class="page-header">
      <div class="header-content">
        <h1 class="page-title"><span class="title-icon">📥</span> Import Data Awal</h1>
        <p class="page-subtitle">Unggah file Excel untuk mengisi data aplikasi, atau sinkronkan langsung dari Google Sheets.</p>
      </div>
      <div class="page-actions" style="display:flex;gap:8px">
        <a href="#/sync" id="btn-sync-health" class="btn btn-primary" style="text-decoration:none;display:inline-flex;align-items:center;gap:6px"><span>🩺 Sync & Health Dashboard</span></a>
        <button id="btn-sync-google" class="btn btn-secondary">
          <span>🔄 Tarik Data dari Google Sheets</span>
        </button>
        <button class="btn btn-warning" id="btn-backup-db">📦 Backup Database</button>
        <button class="btn btn-secondary" id="btn-download-template">⬇️ Download Template</button>
      </div>
    </div>

    <!-- Recovery Banner (interrupted / partial / locked session) -->
    <div id="import-recovery-banner" style="display:none;margin-bottom:16px;"></div>

    <!-- STEP 1: Upload -->
    <div id="step-upload" class="import-step">
      <div class="card">
        <div class="card-body">
          <div class="import-info-box">
            <h3>📋 Petunjuk Import Data Awal</h3>
            <p>Upload file Excel (.xlsx) yang sudah diisi sesuai template. Sistem akan membaca seluruh sheet secara otomatis dan memvalidasi sebelum data disimpan.</p>
            <div class="import-sheet-list">
              ${Object.entries(SHEET_MAP).map(([sheet, { label }]) =>
                `<span class="import-sheet-tag">📄 ${sheet} → ${label}</span>`
              ).join('')}
            </div>
          </div>

          <div class="import-upload-zone" id="upload-zone">
            <div class="upload-icon">📂</div>
            <div class="upload-text">
              <strong>Drag & Drop file Excel di sini</strong>
              <span>atau klik untuk memilih file</span>
            </div>
            <input type="file" id="file-input" accept=".xlsx,.xls" style="display:none">
            <button class="btn btn-primary" id="btn-browse">Pilih File Excel</button>
            <div class="upload-hint">Format: .xlsx | Ukuran maks: 20MB</div>
          </div>
          
          <div id="file-info" style="display:none" class="file-info-bar">
            <span id="file-name-display"></span>
            <button class="btn btn-ghost btn-sm" id="btn-clear-file">✕ Ganti</button>
          </div>
        </div>
      </div>
    </div>

    <!-- STEP 2: Validating (progress) -->
    <div id="step-validating" class="import-step" style="display:none">
      <div class="card">
        <div class="card-body text-center">
          <div class="import-progress-wrap">
            <div class="spinner" style="margin:0 auto 16px"></div>
            <div id="validation-status" class="import-status-text">Membaca file Excel...</div>
            <div class="import-progress-bar"><div class="import-progress-fill" id="validation-bar" style="width:0%"></div></div>
          </div>
        </div>
      </div>
    </div>

    <!-- STEP 3: Preview -->
    <div id="step-preview" class="import-step" style="display:none">
      <!-- Duplicate Strategy -->
      <div class="card mb-12">
        <div class="card-body">
          <h3 style="margin-bottom:12px">⚙️ Pengaturan Duplikat</h3>
          <div class="dup-options">
            <label class="dup-option">
              <input type="radio" name="dup-strategy" value="skip" checked>
              <div class="dup-option-text">
                <strong>Lewati Data Duplikat</strong>
                <span>Data yang sudah ada di database tidak akan diubah</span>
              </div>
            </label>
            <label class="dup-option">
              <input type="radio" name="dup-strategy" value="update">
              <div class="dup-option-text">
                <strong>Perbarui Data yang Sudah Ada</strong>
                <span>Data lama akan ditimpa dengan data dari Excel</span>
              </div>
            </label>
          </div>
        </div>
      </div>

      <!-- Preview Table -->
      <div class="card mb-12">
        <div class="card-body p-0">
          <div class="preview-header">
            <h3>📊 Preview Validasi per Sheet</h3>
            <div id="preview-summary-badges"></div>
          </div>
          <div id="preview-table-container"></div>
        </div>
      </div>

      <!-- Error Detail -->
      <div id="error-detail-section" style="display:none" class="card mb-12">
        <div class="card-body p-0">
          <div class="preview-header">
            <h3>❌ Detail Error</h3>
            <button class="btn btn-secondary btn-sm" id="btn-download-log">⬇️ Download Log Error</button>
          </div>
          <div id="error-detail-container"></div>
        </div>
      </div>

      <!-- Actions -->
      <div class="import-action-bar">
        <button class="btn btn-ghost" id="btn-back-to-upload">← Upload Ulang</button>
        <button class="btn btn-primary" id="btn-start-import" disabled>
          🚀 Mulai Import
        </button>
      </div>
    </div>

    <!-- STEP 4: Importing -->
    <div id="step-importing" class="import-step" style="display:none">
      <div class="card">
        <div class="card-body">
          <h3 style="margin-bottom:20px;text-align:center">⏳ Sedang Mengimport Data...</h3>
          <div id="import-steps-list" class="import-steps-list"></div>
          <div class="import-progress-bar" style="margin-top:20px">
            <div class="import-progress-fill" id="import-bar" style="width:0%"></div>
          </div>
          <div id="import-current-status" class="import-status-text" style="margin-top:8px;text-align:center"></div>
        </div>
      </div>
    </div>

    <!-- STEP 5: Summary -->
    <div id="step-summary" class="import-step" style="display:none">
      <div class="card">
        <div class="card-body">
          <div class="import-summary-header" id="summary-status-icon"></div>
          <div id="summary-action-banner" style="margin-bottom:16px;"></div>
          <div class="import-summary-stats" id="summary-stats"></div>
          <div id="summary-module-results"></div>
          <div class="import-action-bar" style="margin-top:24px">
            <button class="btn btn-secondary" id="btn-import-again">🔄 Import Lagi</button>
            <button class="btn btn-primary" id="btn-go-to-dashboard">📊 Ke Dashboard</button>
          </div>
        </div>
      </div>
    </div>
  `;

  // ── State ──────────────────────────────────────────────────────────────────
  let currentFile    = null;
  let validationResults = null; // array of per-sheet results
  let startTime         = 0;
  let isImporting       = false;
  let activeBatchState  = null;
  let lastBatchResults  = null;

  // ── Elements ───────────────────────────────────────────────────────────────
  const steps = {
    upload:     document.getElementById('step-upload'),
    validating: document.getElementById('step-validating'),
    preview:    document.getElementById('step-preview'),
    importing:  document.getElementById('step-importing'),
    summary:    document.getElementById('step-summary'),
  };

  function showStep(name) {
    Object.entries(steps).forEach(([k, el]) => {
      el.style.display = k === name ? '' : 'none';
    });
  }

  // ── Template Download ──────────────────────────────────────────────────────
  document.getElementById('btn-backup-db')?.addEventListener('click', async () => {
    const btn = document.getElementById('btn-backup-db');
    btn.disabled = true;
    btn.textContent = '⏳ Memproses Backup...';
    try {
      const res = await apiFetch('/api/import/backup');
      if (res.ok) {
        if (!window.XLSX) {
          toastError('Library SheetJS belum termuat. Refresh halaman dan coba lagi.');
          return;
        }
        const XLSX = window.XLSX;
        const wb = XLSX.utils.book_new();
        Object.entries(res.data.database).forEach(([tableName, rows]) => {
          const sheetData = rows.length > 0 ? rows : [{}];
          const sheet = XLSX.utils.json_to_sheet(sheetData);
          XLSX.utils.book_append_sheet(wb, sheet, tableName.substring(0, 31));
        });
        XLSX.writeFile(wb, `FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.xlsx`);
        toastSuccess('Backup berhasil diunduh!');
      } else {
        toastError('Gagal memproses backup: ' + (res.data?.error || 'Unknown error'));
      }
    } catch (err) {
      toastError('Gagal memproses backup: ' + err.message);
    } finally {
      btn.disabled = false;
      btn.textContent = '📦 Backup Database';
    }
  });

  const btnSync = document.getElementById('btn-sync-google');
  if (btnSync) {
    btnSync.addEventListener('click', async () => {
      if (!confirm('Peringatan: Mensinkronkan data dengan Google Sheets akan memperbarui dan menambahkan data baru dari Google Sheets ke dalam FCMS. Data yang sudah Anda buat di FCMS TIDAK akan terhapus. Lanjutkan?')) return;
      
      const originalText = btnSync.innerHTML;
      btnSync.innerHTML = '<span class="spinner"></span> Menyinkronkan...';
      btnSync.disabled = true;
      
      try {
        const res = await apiFetch('/api/sync/google-sheets', { method: 'POST' });
        if (res.ok) {
          alert('Sinkronisasi Berhasil: ' + (res.data?.message || 'Data Karyawan & PIC telah diperbarui.'));
        } else {
          alert('Gagal Sinkronisasi: ' + (res.data?.error || 'Unknown error'));
        }
      } catch (err) {
        alert('Terjadi kesalahan koneksi.');
      } finally {
        btnSync.innerHTML = originalText;
        btnSync.disabled = false;
      }
    });
  }

  document.getElementById('btn-download-template').addEventListener('click', () => {
    generateTemplate();
    toastSuccess('Template Excel berhasil didownload!');
  });

  // ── File Upload ────────────────────────────────────────────────────────────
  const fileInput  = document.getElementById('file-input');
  const uploadZone = document.getElementById('upload-zone');

  document.getElementById('btn-browse').addEventListener('click', (e) => {
    e.stopPropagation(); // prevent bubbling to upload-zone
    fileInput.click();
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files[0]) handleFileSelected(e.target.files[0]);
  });

  // Drag and drop
  uploadZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadZone.classList.add('drag-over');
  });
  uploadZone.addEventListener('dragleave', () => uploadZone.classList.remove('drag-over'));
  uploadZone.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadZone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    if (file && file.name.match(/\.xlsx?$/i)) handleFileSelected(file);
    else toastError('Hanya file .xlsx atau .xls yang didukung.');
  });

  document.getElementById('btn-clear-file').addEventListener('click', () => {
    currentFile = null;
    fileInput.value = '';
    document.getElementById('file-info').style.display = 'none';
    uploadZone.style.display = '';
    showStep('upload');
  });

  // ── Handle File Selected ───────────────────────────────────────────────────
  async function handleFileSelected(file) {
    currentFile = file;
    document.getElementById('file-name-display').textContent = `📄 ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
    document.getElementById('file-info').style.display = 'flex';
    uploadZone.style.display = 'none';

    // Auto start validation
    await runValidation(file);
  }

  // ── Validation ─────────────────────────────────────────────────────────────
  async function runValidation(file) {
    showStep('validating');
    const statusEl = document.getElementById('validation-status');
    const barEl    = document.getElementById('validation-bar');

    try {
      // Ensure SheetJS is loaded
      if (!window.XLSX) throw new Error('Library SheetJS belum termuat. Refresh halaman dan coba lagi.');

      statusEl.textContent = 'Membaca file Excel...';
      barEl.style.width = '20%';
      await sleep(200);

      const arrayBuffer = await file.arrayBuffer();
      const workbook = window.XLSX.read(arrayBuffer, { type: 'array', cellDates: true });

      statusEl.textContent = `Memvalidasi ${workbook.SheetNames.length} sheet...`;
      barEl.style.width = '50%';
      await sleep(100);

      validationResults = validateWorkbook(workbook);

      barEl.style.width = '100%';
      statusEl.textContent = 'Validasi selesai!';
      await sleep(300);

      showPreview();
    } catch (err) {
      showStep('upload');
      toastError('Gagal memproses file: ' + err.message);
      document.getElementById('file-info').style.display = 'flex';
      uploadZone.style.display = 'none';
    }
  }

  // ── Show Preview ───────────────────────────────────────────────────────────
  function showPreview() {
    showStep('preview');

    const totalSheets = validationResults.filter(r => !r.skipped).length;
    const totalRows   = validationResults.reduce((s, r) => s + r.total, 0);
    const totalValid  = validationResults.reduce((s, r) => s + r.valid, 0);
    const totalErrors = validationResults.reduce((s, r) => s + r.errorCount, 0);
    const pct         = totalRows > 0 ? Math.round(totalValid / totalRows * 100) : 0;

    // Summary badges
    document.getElementById('preview-summary-badges').innerHTML = `
      <span class="badge badge-info">${totalSheets} sheet</span>
      <span class="badge badge-secondary">${totalRows} baris</span>
      <span class="badge badge-success">${totalValid} valid (${pct}%)</span>
      ${totalErrors > 0 ? `<span class="badge badge-danger">${totalErrors} error</span>` : ''}
    `;

    // Main preview table — dengan tombol Detail Error per baris
    const tableContainer = document.getElementById('preview-table-container');
    tableContainer.innerHTML = `
      <table class="data-table">
        <thead>
          <tr>
            <th>Sheet (Excel)</th>
            <th>Modul</th>
            <th style="text-align:center">Total</th>
            <th style="text-align:center">Valid</th>
            <th style="text-align:center">Error</th>
            <th style="text-align:center">Status</th>
            <th style="text-align:center">Detail</th>
          </tr>
        </thead>
        <tbody>
          ${validationResults.map((r, idx) => `
            <tr class="${r.errorCount > 0 ? 'row-error' : r.skipped ? 'row-skipped' : 'row-ok'}">
              <td><strong>${r.sheetName}</strong></td>
              <td>${r.label}</td>
              <td style="text-align:center">${r.total}</td>
              <td style="text-align:center"><span class="badge badge-success">${r.valid}</span></td>
              <td style="text-align:center">${r.errorCount > 0 ? `<span class="badge badge-danger">${r.errorCount}</span>` : '<span class="text-muted">–</span>'}</td>
              <td style="text-align:center">
                ${r.skipped
                  ? '<span class="badge badge-neutral">Dilewati</span>'
                  : r.errorCount > 0 && r.valid === 0
                    ? '<span class="badge badge-danger">❌ 0 Valid</span>'
                    : r.errorCount > 0
                      ? '<span class="badge badge-warning">⚠️ Sebagian</span>'
                      : r.valid === 0
                        ? '<span class="badge badge-neutral">Kosong</span>'
                        : '<span class="badge badge-success">✅ Siap</span>'}
              </td>
              <td style="text-align:center">
                ${r.errorCount > 0
                  ? `<button class="btn btn-ghost btn-sm btn-detail-error" data-idx="${idx}">🔍 ${r.errorCount} Error</button>`
                  : '<span class="text-muted">–</span>'}
              </td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    `;

    // Bind detail error buttons
    tableContainer.querySelectorAll('.btn-detail-error').forEach(btn => {
      btn.addEventListener('click', () => {
        const r = validationResults[Number(btn.dataset.idx)];
        showErrorDetail(r);
      });
    });

    // Error detail section (collapsible per sheet)
    const errorSection   = document.getElementById('error-detail-section');
    const errorContainer = document.getElementById('error-detail-container');
    errorContainer.innerHTML = '';
    errorSection.style.display = 'none';

    // Enable Import button (even with errors — import valid rows)
    const importBtn = document.getElementById('btn-start-import');
    if (totalValid === 0) {
      importBtn.disabled = true;
      importBtn.innerHTML = '⚠️ Tidak Ada Data Valid';
    } else {
      importBtn.disabled = false;
      if (totalErrors > 0) {
        importBtn.innerHTML = `🚀 Import ${totalValid} Data Valid (${totalErrors} dilewati)`;
        importBtn.title = 'Baris error akan dilewati, baris valid tetap diimport';
      } else {
        importBtn.innerHTML = `🚀 Mulai Import ${totalValid} Data`;
      }
    }
  }

  // ── Detail Error Modal ─────────────────────────────────────────────────────
  function showErrorDetail(r) {
    const errorSection   = document.getElementById('error-detail-section');
    const errorContainer = document.getElementById('error-detail-container');
    errorSection.style.display = '';

    // Build detailed error table
    const detailRows = r.errors.slice(0, 100).map(e => {
      const errList = Array.isArray(e.errors) ? e.errors : [];
      return errList.map(err => {
        const isObj = typeof err === 'object';
        return `
          <tr>
            <td style="text-align:center"><span class="badge badge-danger">Baris ${e.row}</span></td>
            <td><strong>${isObj ? err.column : '—'}</strong></td>
            <td><code style="font-size:.78rem;color:var(--text-secondary)">${isObj && err.originalValue !== undefined ? (err.originalValue || '(kosong)') : '—'}</code></td>
            <td class="error-msg">${isObj ? err.reason : err}</td>
            <td style="font-size:.78rem;color:var(--success)">
              ${isObj && err.aliases ? `Gunakan salah satu nama kolom:<br><em>${err.aliases}</em>` : isObj && err.hint ? err.hint : ''}
            </td>
          </tr>
        `;
      }).join('');
    }).join('');

    errorContainer.innerHTML = `
      <div class="error-sheet-block">
        <div class="error-sheet-title">
          📄 ${r.sheetName} — ${r.errorCount} baris error dari ${r.total} total
          ${r.errors.length > 100 ? `<span style="font-weight:400">(menampilkan 100 pertama)</span>` : ''}
        </div>
        <div style="overflow-x:auto">
          <table class="data-table error-table" style="min-width:700px">
            <thead>
              <tr>
                <th style="width:80px">Baris</th>
                <th style="width:140px">Kolom Gagal</th>
                <th style="width:140px">Nilai di Excel</th>
                <th>Alasan Error</th>
                <th style="width:220px">💡 Cara Memperbaiki</th>
              </tr>
            </thead>
            <tbody>${detailRows || '<tr><td colspan="5" class="text-muted" style="text-align:center">Tidak ada detail error</td></tr>'}</tbody>
          </table>
        </div>
        ${r.errors.length > 100 ? `
          <div style="padding:10px 20px;font-size:.8rem;color:var(--text-muted)">
            Hanya menampilkan 100 error pertama. Download Log Error untuk melihat semua.
          </div>` : ''}
      </div>
    `;

    // Scroll ke error section
    errorSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }


  // ── Back to upload ──────────────────────────────────────────────────────────
  document.getElementById('btn-back-to-upload').addEventListener('click', () => {
    showStep('upload');
    document.getElementById('file-info').style.display = 'none';
    uploadZone.style.display = '';
    currentFile = null;
    fileInput.value = '';
  });

  // ── Download Error Log ──────────────────────────────────────────────────────
  document.getElementById('btn-download-log').addEventListener('click', () => {
    if (!validationResults) return;
    const ok = downloadErrorLog(validationResults);
    if (!ok) toastSuccess('Tidak ada error untuk didownload.');
    else toastSuccess('Log error berhasil didownload.');
  });

  // ── Interrupted Session Recovery Banner (on mount) ─────────────────────────
  const recoveryBannerEl = document.getElementById('import-recovery-banner');

  function renderRecoveryBanner() {
    if (!recoveryBannerEl) return;
    const savedBatch = getBatchRecoveryState();
    const unfinishedBatchItems = Array.isArray(savedBatch?.items)
      ? savedBatch.items.filter(it => it.status !== 'ok')
      : [];

    if (savedBatch && (savedBatch.status === 'IN_PROGRESS' || savedBatch.status === 'INCOMPLETE' || unfinishedBatchItems.length > 0)) {
      const pendingItems = unfinishedBatchItems.length > 0 ? unfinishedBatchItems : (savedBatch.items || []);
      const firstErr = pendingItems.find(it => it.error || it.errorCode) || pendingItems[0] || {};
      const hasLockConflict = pendingItems.some(it => it.httpStatus === 409 && it.errorCode === 'IMPORT_IN_PROGRESS')
        || savedBatch.status === 'IN_PROGRESS';
      const hasStoredRows = pendingItems.every(it => Array.isArray(it.rows) && it.rows.length > 0);

      recoveryBannerEl.style.display = '';
      recoveryBannerEl.innerHTML = renderImportStatusBannerHTML({
        variant: 'recover',
        title: `🔄 Sesi Import Sebelumnya Belum Selesai (${savedBatch.fileName || 'Excel'})`,
        message: hasStoredRows
          ? `${pendingItems.length} modul belum selesai (${firstErr.error || 'Terhenti akibat refresh atau koneksi terputus'}). Klik Coba Lagi / Lanjutkan untuk meneruskan dengan Idempotency-Key yang sama.`
          : `${pendingItems.length} modul belum selesai. Upload ulang file Excel yang sama agar Idempotency-Key yang tersimpan otomatis digunakan kembali.`,
        idempotencyKey: firstErr.idempotencyKey || '',
        importJobId: firstErr.importJobId || '',
        committedChunks: Number(firstErr.committedChunks || 0),
        totalChunks: Number(firstErr.totalChunks || 0),
        errorCode: firstErr.errorCode || savedBatch.status || '',
        fingerprint: firstErr.fingerprint || 'batch',
        canRetry: hasStoredRows,
        canRecover: Boolean(hasStoredRows && hasLockConflict),
        canDismiss: true,
      });

      recoveryBannerEl.querySelector('[data-import-action="retry"]')?.addEventListener('click', () => {
        if (isImporting) return;
        runImport(savedBatch.onDuplicate || 'skip', savedBatch, { recoverLocked: false });
      });
      recoveryBannerEl.querySelector('[data-import-action="recover"]')?.addEventListener('click', () => {
        if (isImporting) return;
        runImport(savedBatch.onDuplicate || 'skip', savedBatch, { recoverLocked: true });
      });
      recoveryBannerEl.querySelector('[data-import-action="dismiss"]')?.addEventListener('click', () => {
        clearBatchRecoveryState();
        (savedBatch.items || []).forEach(it => it.fingerprint && clearImportSession(it.fingerprint));
        recoveryBannerEl.style.display = 'none';
        recoveryBannerEl.innerHTML = '';
      });
      return;
    }

    const orphanSessions = getIncompleteImportSessions().filter(s => String(s.endpoint || '').startsWith('/api/import/'));
    if (orphanSessions.length > 0) {
      const s = orphanSessions[0];
      recoveryBannerEl.style.display = '';
      recoveryBannerEl.innerHTML = renderImportStatusBannerHTML({
        variant: 'recover',
        title: `🔄 Sesi Import Modul ${s.label || s.module} Belum Selesai`,
        message: s.payloadStored
          ? (s.errorMessage || 'Import terhenti sebelum selesai. Anda dapat melanjutkan dengan Idempotency-Key yang sama.')
          : 'Upload ulang file Excel yang sama untuk melanjutkan dengan Idempotency-Key yang sama.',
        idempotencyKey: s.idempotencyKey,
        importJobId: s.importJobId,
        committedChunks: s.committedChunks,
        totalChunks: s.totalChunks,
        errorCode: s.errorCode || s.status,
        fingerprint: s.fingerprint,
        canRetry: Boolean(s.payloadStored),
        canRecover: Boolean(s.payloadStored && (s.status === 'IN_PROGRESS' || s.status === 'CONFLICT_IN_PROGRESS')),
        canDismiss: true,
      });

      const handleSingleResume = async (recover) => {
        if (isImporting) return;
        isImporting = true;
        try {
          const res = await resumeImportSession(s.fingerprint, { recover });
          clearImportSession(s.fingerprint);
          recoveryBannerEl.style.display = 'none';
          toastSuccess(res.idempotentReplay ? 'Replay idempoten berhasil (0 duplikat).' : 'Sesi import berhasil diselesaikan!');
        } catch (err) {
          toastError(err.message || 'Gagal melanjutkan import');
          renderRecoveryBanner();
        } finally {
          isImporting = false;
        }
      };

      recoveryBannerEl.querySelector('[data-import-action="retry"]')?.addEventListener('click', () => handleSingleResume(false));
      recoveryBannerEl.querySelector('[data-import-action="recover"]')?.addEventListener('click', () => handleSingleResume(true));
      recoveryBannerEl.querySelector('[data-import-action="dismiss"]')?.addEventListener('click', () => {
        clearImportSession(s.fingerprint);
        renderRecoveryBanner();
      });
      return;
    }

    recoveryBannerEl.style.display = 'none';
    recoveryBannerEl.innerHTML = '';
  }

  renderRecoveryBanner();

  // ── Start Import ───────────────────────────────────────────────────────────
  const btnStartImport = document.getElementById('btn-start-import');
  btnStartImport.addEventListener('click', () => {
    if (isImporting) {
      toastError('Proses import sedang berjalan. Mohon tunggu hingga selesai.');
      return;
    }
    const strategy = document.querySelector('input[name="dup-strategy"]:checked')?.value || 'skip';
    runImport(strategy);
  });

  // ── Import Runner ──────────────────────────────────────────────────────────
  async function runImport(onDuplicate, resumeBatch = null, { recoverLocked = false, forceReplayCompleted = false } = {}) {
    if (isImporting) return;
    isImporting = true;
    btnStartImport.disabled = true;

    showStep('importing');
    if (recoveryBannerEl) recoveryBannerEl.style.display = 'none';
    startTime = Date.now();

    let batchItems;
    if (resumeBatch && Array.isArray(resumeBatch.items)) {
      batchItems = resumeBatch.items.map(it => ({ ...it }));
    } else {
      const toImport = [];
      IMPORT_ORDER.forEach(module => {
        const r = validationResults?.find(res => res.module === module && res.mapped?.length > 0);
        if (r) toImport.push(r);
      });
      batchItems = toImport.map(r => ({
        module: r.module,
        label: r.label,
        endpoint: `/api/import/${r.module}`,
        rows: r.mapped,
        rowCount: r.mapped.length,
        status: 'pending',
        inserted: 0,
        updated: 0,
        skipped: 0,
        committedChunks: 0,
        totalChunks: 0,
        idempotencyKey: null,
        fingerprint: null,
        importJobId: null,
        idempotentReplay: false,
        httpStatus: null,
        errorCode: null,
        error: null,
        recoverable: false,
      }));
    }

    for (const item of batchItems) {
      const body = { rows: item.rows, onDuplicate };
      if (!item.fingerprint) {
        item.fingerprint = await computeImportFingerprint(item.endpoint, body);
      }
      if (!item.idempotencyKey) {
        const prefix = String(item.module || 'imp').slice(0, 20);
        item.idempotencyKey = (globalThis?.crypto?.randomUUID)
          ? `imp_${prefix}_${globalThis.crypto.randomUUID()}`
          : `imp_${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;
      }
    }

    const batchState = {
      batchId: resumeBatch?.batchId || (globalThis?.crypto?.randomUUID ? globalThis.crypto.randomUUID() : `batch_${Date.now()}`),
      fileName: resumeBatch?.fileName || currentFile?.name || 'Excel Import',
      onDuplicate,
      status: 'IN_PROGRESS',
      items: batchItems,
    };
    activeBatchState = batchState;
    saveBatchRecoveryState(batchState);

    const stepsList = document.getElementById('import-steps-list');
    stepsList.innerHTML = batchItems.map(item => `
      <div class="import-step-item" id="step-item-${item.module}">
        <span class="step-item-icon" id="step-icon-${item.module}">${item.status === 'ok' && !forceReplayCompleted ? '✅' : '⏸️'}</span>
        <span class="step-item-label">${escapeHtml(item.label)} <span class="step-item-count">(${item.rowCount || item.rows?.length || 0} data)</span></span>
        <span class="step-item-status" id="step-status-${item.module}">${item.status === 'ok' && !forceReplayCompleted ? `<span class="badge badge-success">${item.inserted || 0} berhasil</span>` : ''}</span>
      </div>
    `).join('');

    const barEl    = document.getElementById('import-bar');
    const statusEl = document.getElementById('import-current-status');

    try {
      for (let i = 0; i < batchItems.length; i++) {
        const item = batchItems[i];
        const iconEl   = document.getElementById(`step-icon-${item.module}`);
        const statusSm = document.getElementById(`step-status-${item.module}`);

        if (item.status === 'ok' && !forceReplayCompleted) {
          barEl.style.width = `${Math.round(((i + 1) / batchItems.length) * 100)}%`;
          continue;
        }

        iconEl.textContent   = '🔄';
        statusSm.textContent = recoverLocked && item.httpStatus === 409 ? 'Memulihkan lock...' : 'Mengimport...';
        statusEl.textContent = `Mengimport ${item.label}... (${i + 1}/${batchItems.length})`;
        barEl.style.width    = `${Math.round((i / Math.max(1, batchItems.length)) * 100)}%`;

        saveBatchRecoveryState(batchState);

        try {
          const shouldRecover = Boolean(
            recoverLocked && (item.httpStatus === 409 || item.errorCode === 'IMPORT_IN_PROGRESS' || item.status === 'pending' || item.status === 'error')
          );
          const res = await submitJournaledImport({
            module: item.module,
            label: item.label,
            endpoint: item.endpoint,
            body: { rows: item.rows, onDuplicate },
            idempotencyKey: item.idempotencyKey,
            recover: shouldRecover,
          });

          const d = res.data || {};
          item.status = 'ok';
          item.inserted = Number(d.inserted || 0);
          item.updated = Number(d.updated || 0);
          item.skipped = Number(d.skipped || 0);
          item.idempotentReplay = Boolean(res.idempotentReplay || d.idempotentReplay);
          item.importJobId = res.importJobId || d.importJobId || item.importJobId || null;
          item.idempotencyKey = res.idempotencyKey || item.idempotencyKey;
          item.httpStatus = res.status;
          item.error = null;
          item.errorCode = null;
          item.recoverable = false;

          iconEl.textContent = item.idempotentReplay ? '♻️' : '✅';
          const replayBadge = item.idempotentReplay ? ' <span class="badge badge-info">Replay Idempoten</span>' : '';
          const skipBadge = item.skipped > 0 ? ` <span class="badge badge-neutral">${item.skipped} skip</span>` : '';
          const updBadge = item.updated > 0 ? ` <span class="badge badge-info">${item.updated} update</span>` : '';
          statusSm.innerHTML = `<span class="badge badge-success">${item.inserted} berhasil</span>${updBadge}${skipBadge}${replayBadge}`;
        } catch (err) {
          item.status = 'error';
          item.error = err.message || 'Gagal mengimport data';
          item.errorCode = err.code || 'IMPORT_ERROR';
          item.httpStatus = Number(err.status ?? 0);
          item.committedChunks = Number(err.committedChunks || 0);
          item.totalChunks = Number(err.totalChunks || 0);
          item.importJobId = err.importJobId || item.importJobId || null;
          item.idempotencyKey = err.idempotencyKey || item.idempotencyKey || null;
          item.recoverable = Boolean(err.recoverable);

          iconEl.textContent = item.httpStatus === 409 ? '⏳' : '❌';
          const chunkBadge = item.totalChunks > 0 ? ` <span class="badge badge-warning">Chunk ${item.committedChunks}/${item.totalChunks}</span>` : '';
          statusSm.innerHTML = `<span class="badge badge-danger">${escapeHtml(item.errorCode || 'Gagal')}</span>${chunkBadge}`;
        }

        batchState.status = batchItems.every(it => it.status === 'ok') ? 'COMPLETED' : 'INCOMPLETE';
        saveBatchRecoveryState(batchState);
        await sleep(80);
      }

      const allSucceeded = batchItems.every(it => it.status === 'ok');
      if (allSucceeded) {
        clearBatchRecoveryState();
      } else {
        batchState.status = 'INCOMPLETE';
        saveBatchRecoveryState(batchState);
      }

      const summary = {
        totalSheets: batchItems.length,
        totalRows: batchItems.reduce((s, it) => s + (it.rowCount || it.rows?.length || 0), 0),
        inserted: batchItems.reduce((s, it) => s + (it.inserted || 0), 0),
        updated: batchItems.reduce((s, it) => s + (it.updated || 0), 0),
        skipped: batchItems.reduce((s, it) => s + (it.skipped || 0), 0),
        failed: batchItems.filter(it => it.status !== 'ok').length,
        hasReplays: batchItems.some(it => it.idempotentReplay),
        moduleResults: batchItems,
        batchState,
      };

      lastBatchResults = summary;
      barEl.style.width = '100%';
      statusEl.textContent = allSucceeded ? 'Selesai!' : 'Selesai dengan modul gagal/terhenti (dapat dilanjutkan)';
      await sleep(250);

      showSummary(summary);
    } finally {
      isImporting = false;
      btnStartImport.disabled = false;
    }
  }

  // ── Summary ────────────────────────────────────────────────────────────────
  function showSummary(summary) {
    showStep('summary');
    const duration = ((Date.now() - startTime) / 1000).toFixed(1);
    const allOk    = summary.failed === 0 && summary.moduleResults.every(m => m.status === 'ok');
    const hasConflict409 = summary.moduleResults.some(m => m.status !== 'ok' && m.httpStatus === 409 && m.errorCode === 'IMPORT_IN_PROGRESS');
    const firstFailed = summary.moduleResults.find(m => m.status !== 'ok');

    document.getElementById('summary-status-icon').innerHTML = `
      <div class="summary-icon">${allOk ? (summary.hasReplays ? '♻️' : '🎉') : '⚠️'}</div>
      <h2 class="summary-title">${allOk ? (summary.hasReplays ? 'Import & Replay Idempoten Berhasil!' : 'Import Berhasil!') : 'Import Belum Selesai (Dapat Dilanjutkan)'}</h2>
    `;

    const actionBannerEl = document.getElementById('summary-action-banner');
    if (actionBannerEl) {
      if (!allOk && firstFailed) {
        actionBannerEl.innerHTML = renderImportStatusBannerHTML({
          variant: hasConflict409 ? 'warning' : 'error',
          title: hasConflict409
            ? '⏳ Ada Modul Terkunci / Sedang Berjalan (HTTP 409)'
            : '⚠️ Beberapa Modul Gagal atau Terhenti di Tengah Proses',
          message: firstFailed.error || 'Klik Coba Lagi untuk melanjutkan dengan Idempotency-Key yang sama tanpa menduplikasi data.',
          idempotencyKey: firstFailed.idempotencyKey || '',
          importJobId: firstFailed.importJobId || '',
          committedChunks: firstFailed.committedChunks || 0,
          totalChunks: firstFailed.totalChunks || 0,
          errorCode: firstFailed.errorCode || '',
          fingerprint: firstFailed.fingerprint || '',
          canRetry: true,
          canRecover: hasConflict409,
          canDismiss: false,
        });
        actionBannerEl.querySelector('[data-import-action="retry"]')?.addEventListener('click', () => {
          runImport(summary.batchState?.onDuplicate || 'skip', summary.batchState, { recoverLocked: false });
        });
        actionBannerEl.querySelector('[data-import-action="recover"]')?.addEventListener('click', () => {
          runImport(summary.batchState?.onDuplicate || 'skip', summary.batchState, { recoverLocked: true });
        });
      } else if (allOk && summary.moduleResults.length > 0) {
        const firstOk = summary.moduleResults[0];
        actionBannerEl.innerHTML = renderImportStatusBannerHTML({
          variant: 'success',
          title: summary.hasReplays ? '♻️ Replay Idempoten Dikonfirmasi (0 Duplikat)' : '✅ Semua Modul Terkonfirmasi Server',
          message: summary.hasReplays
            ? 'Server mengembalikan hasil tersimpan menggunakan Idempotency-Key yang sama tanpa menulis baris ganda.'
            : 'Semua chunk telah committed di server beserta Audit Log terkait.',
          idempotencyKey: firstOk.idempotencyKey || '',
          importJobId: firstOk.importJobId || '',
          canReplay: true,
          canDismiss: false,
        });
        actionBannerEl.querySelector('[data-import-action="replay"]')?.addEventListener('click', () => {
          runImport(summary.batchState?.onDuplicate || 'skip', summary.batchState, { recoverLocked: false, forceReplayCompleted: true });
        });
      } else {
        actionBannerEl.innerHTML = '';
      }
    }

    document.getElementById('summary-stats').innerHTML = `
      <div class="summary-stat-card">
        <div class="stat-value">${summary.totalSheets}</div>
        <div class="stat-label">Total Sheet</div>
      </div>
      <div class="summary-stat-card">
        <div class="stat-value">${summary.totalRows}</div>
        <div class="stat-label">Total Data</div>
      </div>
      <div class="summary-stat-card success">
        <div class="stat-value">${summary.inserted}</div>
        <div class="stat-label">Berhasil Diimport</div>
      </div>
      <div class="summary-stat-card neutral">
        <div class="stat-value">${summary.skipped}</div>
        <div class="stat-label">Dilewati (Duplikat)</div>
      </div>
      ${summary.updated > 0 ? `<div class="summary-stat-card info"><div class="stat-value">${summary.updated}</div><div class="stat-label">Diperbarui</div></div>` : ''}
      ${summary.failed > 0 ? `<div class="summary-stat-card danger"><div class="stat-value">${summary.failed}</div><div class="stat-label">Modul Gagal</div></div>` : ''}
      <div class="summary-stat-card info">
        <div class="stat-value">${duration}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `;

    document.getElementById('summary-module-results').innerHTML = `
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status & Journal</th></tr>
        </thead>
        <tbody>
          ${summary.moduleResults.map((m, idx) => `
            <tr>
              <td>
                <div><strong>${escapeHtml(m.label)}</strong></div>
                ${m.idempotencyKey ? `<div style="font-size:.72rem;color:var(--text-muted)">Key: <code>${escapeHtml(m.idempotencyKey.slice(0, 26))}…</code>${m.importJobId ? ` | Job: <code>${escapeHtml(m.importJobId.slice(0, 10))}…</code>` : ''}</div>` : ''}
              </td>
              <td style="text-align:center"><span class="badge badge-success">${m.inserted || 0}</span>${m.updated > 0 ? ` <span class="badge badge-info">+${m.updated} upd</span>` : ''}</td>
              <td style="text-align:center"><span class="badge badge-neutral">${m.skipped || 0}</span></td>
              <td style="text-align:center">
                ${m.status === 'ok'
                  ? `<span class="badge badge-success">${m.idempotentReplay ? '♻️ Replay Idempoten' : '✅ Sukses'}</span>`
                  : `<div>
                      <span class="badge badge-danger" title="${escapeHtml(m.error || '')}">❌ ${escapeHtml(m.errorCode || 'Gagal')}${m.httpStatus ? ` (${m.httpStatus})` : ''}</span>
                      ${m.totalChunks > 0 ? `<div style="font-size:.72rem;margin-top:4px;">Chunk ${m.committedChunks}/${m.totalChunks}</div>` : ''}
                      <div style="font-size:.75rem;color:var(--danger);margin-top:2px;">${escapeHtml(m.error || '')}</div>
                      <button type="button" class="btn btn-primary btn-sm" style="margin-top:6px" data-row-retry="${idx}">🔄 Coba Lagi</button>
                      ${m.httpStatus === 409 && m.errorCode === 'IMPORT_IN_PROGRESS' ? `<button type="button" class="btn btn-warning btn-sm" style="margin-top:6px" data-row-recover="${idx}">🛠️ Pulihkan</button>` : ''}
                    </div>`}
              </td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    `;

    document.querySelectorAll('[data-row-retry]').forEach(btn => {
      btn.addEventListener('click', () => {
        runImport(summary.batchState?.onDuplicate || 'skip', summary.batchState, { recoverLocked: false });
      });
    });
    document.querySelectorAll('[data-row-recover]').forEach(btn => {
      btn.addEventListener('click', () => {
        runImport(summary.batchState?.onDuplicate || 'skip', summary.batchState, { recoverLocked: true });
      });
    });
  }

  // ── Button actions in summary ──────────────────────────────────────────────
  document.getElementById('btn-import-again').addEventListener('click', () => {
    currentFile = null; validationResults = null; activeBatchState = null; lastBatchResults = null;
    fileInput.value = '';
    document.getElementById('file-info').style.display = 'none';
    uploadZone.style.display = '';
    renderRecoveryBanner();
    showStep('upload');
  });

  document.getElementById('btn-go-to-dashboard').addEventListener('click', () => {
    window.location.hash = '/dashboard';
  });
}

// ── Utility ──────────────────────────────────────────────────────────────────
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
