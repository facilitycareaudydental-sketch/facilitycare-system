/**
 * import.js — Halaman Settings → Import Data Awal
 * Alur: Upload → Validasi lokal → Preview → Pilih duplikat → Import → Ringkasan
 */
import { apiFetch } from '../config.js';
import { toastSuccess, toastError } from '../components/toast.js';
import { validateWorkbook, generateTemplate, downloadErrorLog, SHEET_MAP } from '../utils/importer.js';

// Import order: employees first (contracts depend on it)
const IMPORT_ORDER = [
  'employees', 'contracts', 'relievers', 'schedule', 'issues',
  'one_on_one', 'training', 'checklist', 'forms', 'sop',
  'inspection', 'cleaning', 'fogging', 'basecamp', 'supply',
];

export function renderImportPage(container) {
  container.innerHTML = `
    <div class="page-header">
      <h1 class="page-title">📥 Import Data Awal</h1>
      <div class="page-actions">
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
      const workbook = window.XLSX.read(arrayBuffer, { type: 'array', cellDates: false });

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
    const hasErrors   = totalErrors > 0;

    // Summary badges
    document.getElementById('preview-summary-badges').innerHTML = `
      <span class="badge badge-info">${totalSheets} sheet dikenali</span>
      <span class="badge badge-secondary">${totalRows} baris</span>
      <span class="badge badge-success">${totalValid} valid</span>
      ${hasErrors ? `<span class="badge badge-danger">${totalErrors} error</span>` : ''}
    `;

    // Main preview table
    const tableContainer = document.getElementById('preview-table-container');
    tableContainer.innerHTML = `
      <table class="data-table">
        <thead>
          <tr>
            <th>Sheet</th>
            <th>Modul Tujuan</th>
            <th style="text-align:center">Total</th>
            <th style="text-align:center">Valid</th>
            <th style="text-align:center">Error</th>
            <th style="text-align:center">Status</th>
          </tr>
        </thead>
        <tbody>
          ${validationResults.map(r => `
            <tr class="${r.errorCount > 0 ? 'row-error' : r.skipped ? 'row-skipped' : 'row-ok'}">
              <td><strong>${r.sheetName}</strong></td>
              <td>${r.label}</td>
              <td style="text-align:center">${r.total}</td>
              <td style="text-align:center"><span class="badge badge-success">${r.valid}</span></td>
              <td style="text-align:center">${r.errorCount > 0 ? `<span class="badge badge-danger">${r.errorCount}</span>` : '<span class="text-muted">–</span>'}</td>
              <td style="text-align:center">
                ${r.skipped ? '<span class="badge badge-neutral">Dilewati</span>'
                  : r.errorCount > 0 ? '<span class="badge badge-danger">❌ Ada Error</span>'
                  : r.valid === 0 ? '<span class="badge badge-neutral">Kosong</span>'
                  : '<span class="badge badge-success">✅ Siap</span>'}
              </td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    `;

    // Error detail
    const errorSection = document.getElementById('error-detail-section');
    const errorContainer = document.getElementById('error-detail-container');
    const sheetsWithErrors = validationResults.filter(r => r.errorCount > 0);

    if (sheetsWithErrors.length > 0) {
      errorSection.style.display = '';
      errorContainer.innerHTML = sheetsWithErrors.map(r => `
        <div class="error-sheet-block">
          <div class="error-sheet-title">📄 ${r.sheetName} — ${r.errorCount} baris error</div>
          <table class="data-table error-table">
            <thead><tr><th>Baris</th><th>Penyebab Error</th><th>Data</th></tr></thead>
            <tbody>
              ${r.errors.slice(0, 20).map(e => `
                <tr>
                  <td><span class="badge badge-danger">Baris ${e.row}</span></td>
                  <td>${e.errors.map(msg => `<div class="error-msg">• ${msg}</div>`).join('')}</td>
                  <td><small>${Object.entries(e.data || {}).filter(([,v]) => v != null).slice(0,4).map(([k,v]) => `<b>${k}:</b> ${v}`).join(' | ')}</small></td>
                </tr>
              `).join('')}
              ${r.errors.length > 20 ? `<tr><td colspan="3" class="text-muted" style="text-align:center">... dan ${r.errors.length - 20} error lainnya (lihat Log Error)</td></tr>` : ''}
            </tbody>
          </table>
        </div>
      `).join('');
    } else {
      errorSection.style.display = 'none';
    }

    // Enable/disable Import button
    const importBtn = document.getElementById('btn-start-import');
    if (hasErrors) {
      importBtn.disabled = true;
      importBtn.title = 'Perbaiki semua error terlebih dahulu sebelum import';
      importBtn.innerHTML = '⚠️ Ada Error — Perbaiki Dulu';
    } else if (totalValid === 0) {
      importBtn.disabled = true;
      importBtn.innerHTML = '⚠️ Tidak Ada Data Valid';
    } else {
      importBtn.disabled = false;
      importBtn.innerHTML = `🚀 Mulai Import ${totalValid} Data`;
    }
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
