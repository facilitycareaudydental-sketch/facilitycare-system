/**
 * Helper utilities for Excel Export & Import using SheetJS (XLSX)
 */

/**
 * Parses an uploaded Excel file and returns an array of objects
 * @param {File} file - The uploaded Excel file
 * @returns {Promise<Array<Object>>} - JSON array of the sheet data
 */
export function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        // Assume data is in the first sheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Log pre-parsing stats
        console.log('--- START EXCEL PARSING ---');
        console.log(`File Name: ${file.name}`);
        console.log(`File Size: ${(file.size / 1024).toFixed(2)} KB`);
        console.log(`File Type: ${file.type || 'unknown'}`);
        console.log(`Sheets Found: ${workbook.SheetNames.join(', ')}`);
        console.log(`Sheet Used: ${firstSheetName}`);
        
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
        const rowCount = range.e.r - range.s.r + 1;
        const colCount = range.e.c - range.s.c + 1;
        console.log(`Total Rows (including empty): ${rowCount}`);
        console.log(`Total Columns: ${colCount}`);
        
        const headers = [];
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const cell = worksheet[XLSX.utils.encode_cell({c: C, r: range.s.r})];
          if (cell && cell.v) headers.push(cell.v);
        }
        console.log(`Headers Found: ${headers.join(', ')}`);
        console.log('---------------------------');

        const json = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
        Object.defineProperty(json, '__worksheet', { value: worksheet, enumerable: false });
        Object.defineProperty(json, '__headers', { value: headers, enumerable: false });
        resolve(json);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Generates and downloads an Excel file from JSON data
 * @param {Array<Object>} data - JSON array
 * @param {String} filename - Name of the downloaded file (without extension)
 */
export function downloadExcel(data, filename) {
  try {
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Data');
    XLSX.writeFile(workbook, `${filename}.xlsx`);
  } catch (err) {
    console.error('Error generating Excel file:', err);
    throw err;
  }
}

let _excelDropdownListenerRegistered = false;

function ensureExcelDropdownDelegation() {
  if (_excelDropdownListenerRegistered) return;
  if (typeof document === 'undefined' || typeof document.addEventListener !== 'function') return;
  _excelDropdownListenerRegistered = true;

  document.addEventListener('click', (event) => {
    const toggleBtn = event.target?.closest?.('[data-excel-toggle-menu]');
    if (toggleBtn) {
      const menuId = toggleBtn.getAttribute('data-excel-toggle-menu');
      const targetMenu = menuId ? document.getElementById(menuId) : null;
      document.querySelectorAll('.aksi-menu-content.show-aksi-menu').forEach((menuEl) => {
        if (menuEl !== targetMenu) menuEl.classList.remove('show-aksi-menu');
      });
      if (targetMenu) {
        targetMenu.classList.toggle('show-aksi-menu');
      }
      return;
    }

    if (!event.target?.closest?.('.excel-actions-dropdown')) {
      document.querySelectorAll('.aksi-menu-content.show-aksi-menu').forEach((menuEl) => {
        menuEl.classList.remove('show-aksi-menu');
      });
    }
  });
}

ensureExcelDropdownDelegation();

export function bindExcelDropdown(moduleName, root = typeof document !== 'undefined' ? document : null) {
  ensureExcelDropdownDelegation();
  if (!root || typeof root.querySelector !== 'function') return;
  const btn = root.querySelector(`#btn-aksi-${moduleName}`);
  const menu = root.querySelector(`#excel-menu-${moduleName}`);
  if (btn && menu && !btn.dataset.excelBound) {
    btn.dataset.excelBound = '1';
    if (!btn.getAttribute('data-excel-toggle-menu')) {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        menu.classList.toggle('show-aksi-menu');
      });
    }
  }
}

/**
 * Renders HTML for Export/Import action buttons (CSP-safe without inline onclick)
 * @param {String} moduleName - Name used for ID generation
 * @returns {String} - HTML string
 */
export function renderExcelButtons(moduleName) {
  ensureExcelDropdownDelegation();
  return `
    <div class="excel-actions-dropdown" style="position:relative; display:inline-block;">
      <button type="button" class="btn" id="btn-aksi-${moduleName}" data-excel-toggle-menu="excel-menu-${moduleName}" style="background:#172033; border:1px solid #334155; padding:8px 16px; border-radius:8px; font-weight:600; color:#F8FAFC; display:flex; align-items:center; gap:6px; cursor:pointer;">
        ⋮ Aksi
      </button>
      <div id="excel-menu-${moduleName}" class="aksi-menu-content" style="display:none; position:absolute; top:calc(100% + 4px); right:0; background:#172033; border:1px solid #334155; border-radius:8px; box-shadow:0 10px 25px rgba(0,0,0,0.45); flex-direction:column; min-width:180px; z-index:999; padding:8px 0;">
        <button type="button" class="dropdown-item" id="btn-export-${moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#E2E8F0; display:flex; gap:8px; align-items:center;">
          📥 Export Excel
        </button>
        <button type="button" class="dropdown-item" id="btn-template-${moduleName}" style="width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#E2E8F0; display:flex; gap:8px; align-items:center;">
          📄 Download Template
        </button>
        <label class="dropdown-item" style="display:flex; width:100%; text-align:left; padding:10px 16px; background:none; border:none; cursor:pointer; font-size:0.9rem; color:#E2E8F0; margin:0; gap:8px; align-items:center;" id="label-import-${moduleName}">
          📤 Import Excel
          <input type="file" id="input-import-${moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
        </label>
      </div>
    </div>
    <style>
      .show-aksi-menu { display: flex !important; }
      .dropdown-item:hover { background-color: #22314A !important; color: #93C5FD !important; }
    </style>
  `;
}
