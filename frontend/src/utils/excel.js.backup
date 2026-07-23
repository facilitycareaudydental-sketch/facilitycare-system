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
        // Convert to array of objects
        const json = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
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

/**
 * Renders HTML for Export/Import action buttons
 * @param {String} moduleName - Name used for ID generation
 * @returns {String} - HTML string
 */
export function renderExcelButtons(moduleName) {
  return `
    <div class="excel-actions" style="display:flex;gap:0.5rem;margin-bottom:1rem;">
      <button class="btn btn-secondary btn-sm" id="btn-export-${moduleName}">
        📥 Export Excel
      </button>
      <button class="btn btn-secondary btn-sm" id="btn-template-${moduleName}">
        📄 Download Template
      </button>
      <label class="btn btn-primary btn-sm" style="cursor:pointer;margin:0;" id="label-import-${moduleName}">
        <span class="import-text">📤 Import Excel</span>
        <input type="file" id="input-import-${moduleName}" accept=".xlsx, .xls, .csv" style="display:none;">
      </label>
    </div>
  `;
}
