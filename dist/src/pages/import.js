/**
 * import.js — Halaman Settings → Import Data Awal
 * Alur: Upload → Validasi lokal → Preview → Pilih duplikat → Import → Ringkasan
 */
import { apiFetch } from '../config.js';
import { toastSuccess, toastError } from '../components/toast.js';
import { validateWorkbook, generateTemplate, downloadErrorLog, SHEET_MAP } from '../utils/importer.js';

// Import order: validation referensi first, then employees, etc.
const IMPORT_ORDER = [
  'validation', 'employees', 'contracts', 'relievers', 'schedule', 'issues',
  'one_on_one', 'training', 'checklist', 'forms', 'sop',
  'inspection', 'cleaning', 'fogging', 'basecamp', 'supply',
];

export function renderImportPage(container) {
  container.innerHTML = `
    <div class="page-header">
      <div class="header-content">
        <h1 class="page-title"><span class="title-icon">📥</span> Import Data Awal</h1>
        <p class="page-subtitle">Unggah file Excel untuk mengisi data aplikasi, atau sinkronkan langsung dari Google Sheets.</p>
      </div>
      <div class="page-actions" style="display:flex;gap:8px">
        <button id="btn-sync-google" class="btn btn-secondary">
          <span>🔄 Tarik Data dari Google Sheets</span>
        </button>
        <button class="btn btn-warning" id="btn-backup-db">📦 Backup Database</button>
        <button class="btn btn-secondary" id="btn-download-template">⬇️ Download Template</button>
      </div>
    </div>

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
  let startTime      = 0;

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
        const blob = new Blob([JSON.stringify(res.data, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `FCMS_Database_Backup_${new Date().toISOString().slice(0,10)}.json`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
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

  // ── Start Import ───────────────────────────────────────────────────────────
  document.getElementById('btn-start-import').addEventListener('click', () => {
    const strategy = document.querySelector('input[name="dup-strategy"]:checked')?.value || 'skip';
    runImport(strategy);
  });

  // ── Import Runner ──────────────────────────────────────────────────────────
  async function runImport(onDuplicate) {
    showStep('importing');
    startTime = Date.now();

    // Build ordered list of modules to import
    const toImport = [];
    IMPORT_ORDER.forEach(module => {
      const r = validationResults?.find(r => r.module === module && r.mapped?.length > 0);
      if (r) toImport.push(r);
    });

    // Render steps list
    const stepsList = document.getElementById('import-steps-list');
    stepsList.innerHTML = toImport.map(r => `
      <div class="import-step-item" id="step-item-${r.module}">
        <span class="step-item-icon" id="step-icon-${r.module}">⏸️</span>
        <span class="step-item-label">${r.label} <span class="step-item-count">(${r.mapped.length} data)</span></span>
        <span class="step-item-status" id="step-status-${r.module}"></span>
      </div>
    `).join('');

    const barEl    = document.getElementById('import-bar');
    const statusEl = document.getElementById('import-current-status');

    const summary = {
      totalSheets: toImport.length,
      totalRows:   toImport.reduce((s, r) => s + r.mapped.length, 0),
      inserted:    0,
      skipped:     0,
      failed:      0,
      moduleResults: [],
    };

    // Import each module sequentially
    for (let i = 0; i < toImport.length; i++) {
      const r = toImport[i];
      const iconEl   = document.getElementById(`step-icon-${r.module}`);
      const statusSm = document.getElementById(`step-status-${r.module}`);

      iconEl.textContent   = '🔄';
      statusSm.textContent = 'Mengimport...';
      statusEl.textContent = `Mengimport ${r.label}...`;
      barEl.style.width    = `${Math.round((i / toImport.length) * 100)}%`;

      try {
        const res = await apiFetch(`/api/import/${r.module}`, {
          method: 'POST',
          body: JSON.stringify({ rows: r.mapped, onDuplicate }),
        });

        if (res.ok) {
          const d = res.data;
          summary.inserted += d.inserted || 0;
          summary.skipped  += d.skipped  || 0;
          summary.moduleResults.push({ label: r.label, inserted: d.inserted || 0, skipped: d.skipped || 0, status: 'ok' });
          iconEl.textContent   = '✅';
          statusSm.innerHTML   = `<span class="badge badge-success">${d.inserted || 0} berhasil</span>${d.skipped > 0 ? ` <span class="badge badge-neutral">${d.skipped} skip</span>` : ''}`;
        } else {
          summary.failed++;
          summary.moduleResults.push({ label: r.label, inserted: 0, skipped: 0, status: 'error', error: res.data?.error });
          iconEl.textContent   = '❌';
          statusSm.innerHTML   = `<span class="badge badge-danger">Gagal</span>`;
        }
      } catch (err) {
        summary.failed++;
        summary.moduleResults.push({ label: r.label, inserted: 0, skipped: 0, status: 'error', error: err.message });
        iconEl.textContent   = '❌';
        statusSm.innerHTML   = `<span class="badge badge-danger">Gagal</span>`;
      }

      await sleep(150); // small delay so user can see progress
    }

    barEl.style.width    = '100%';
    statusEl.textContent = 'Selesai!';
    await sleep(400);

    showSummary(summary);
  }

  // ── Summary ────────────────────────────────────────────────────────────────
  function showSummary(summary) {
    showStep('summary');
    const duration = ((Date.now() - startTime) / 1000).toFixed(1);
    const allOk    = summary.failed === 0;

    document.getElementById('summary-status-icon').innerHTML = `
      <div class="summary-icon">${allOk ? '🎉' : '⚠️'}</div>
      <h2 class="summary-title">${allOk ? 'Import Berhasil!' : 'Import Selesai dengan Beberapa Error'}</h2>
    `;

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
      ${summary.failed > 0 ? `<div class="summary-stat-card danger"><div class="stat-value">${summary.failed}</div><div class="stat-label">Modul Gagal</div></div>` : ''}
      <div class="summary-stat-card info">
        <div class="stat-value">${duration}s</div>
        <div class="stat-label">Durasi Proses</div>
      </div>
    `;

    document.getElementById('summary-module-results').innerHTML = `
      <table class="data-table" style="margin-top:16px">
        <thead>
          <tr><th>Modul</th><th style="text-align:center">Berhasil</th><th style="text-align:center">Dilewati</th><th style="text-align:center">Status</th></tr>
        </thead>
        <tbody>
          ${summary.moduleResults.map(m => `
            <tr>
              <td>${m.label}</td>
              <td style="text-align:center"><span class="badge badge-success">${m.inserted}</span></td>
              <td style="text-align:center"><span class="badge badge-neutral">${m.skipped}</span></td>
              <td style="text-align:center">
                ${m.status === 'ok'
                  ? '<span class="badge badge-success">✅ Sukses</span>'
                  : `<span class="badge badge-danger" title="${m.error || ''}">❌ Gagal</span>`}
              </td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    `;
  }

  // ── Button actions in summary ──────────────────────────────────────────────
  document.getElementById('btn-import-again').addEventListener('click', () => {
    currentFile = null; validationResults = null;
    fileInput.value = '';
    document.getElementById('file-info').style.display = 'none';
    uploadZone.style.display = '';
    showStep('upload');
  });

  document.getElementById('btn-go-to-dashboard').addEventListener('click', () => {
    window.location.hash = '/dashboard';
  });
}

// ── Utility ──────────────────────────────────────────────────────────────────
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
