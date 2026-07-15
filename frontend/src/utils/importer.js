/**
 * importer.js
 * Client-side validation & data mapping engine for Import Data Awal.
 * Runs entirely in the browser — no data is sent to the server until all validations pass.
 */

// ─── Sheet → Module Mapping ─────────────────────────────────────────────────
export const SHEET_MAP = {
  'Master Karyawan':             { module: 'employees',   label: 'Karyawan' },
  'Data Kontrak':                { module: 'contracts',   label: 'Kontrak' },
  'Jadwal Reliefer':             { module: 'relievers',   label: 'Reliefer' },
  'Time Line':                   { module: 'schedule',    label: 'Jadwal Kegiatan' },
  'Permasalahan':                { module: 'issues',      label: 'Permasalahan' },
  'One on One':                  { module: 'one_on_one',  label: 'One on One' },
  'Training':                    { module: 'training',    label: 'Training' },
  'Master Checklist':            { module: 'checklist',   label: 'Master Checklist' },
  'Master Form':                 { module: 'forms',       label: 'Master Form' },
  'SOP':                         { module: 'sop',         label: 'SOP' },
  'Report Inspeksi Hygiene 2026':{ module: 'inspection',  label: 'Laporan Inspeksi' },
  'Report GCDC 2026':            { module: 'cleaning',    label: 'Laporan GC/DC' },
  'Report Fogging 2026':         { module: 'fogging',     label: 'Laporan Fogging' },
  'Rekap Laporan Basecamp':      { module: 'basecamp',    label: 'Rekap Basecamp' },
  'Permintaan Chemical':         { module: 'supply',      label: 'Inventory Chemical' },
};

// ─── Module Schemas: required fields & field mappings ──────────────────────
const SCHEMAS = {
  employees: {
    required: ['Nama Lengkap'],
    map: (row) => ({
      full_name:    str(row['Nama Lengkap']),
      branch_name:  str(row['Cabang']),
      division:     str(row['Divisi']) || 'FACILITY CARE',
      phone:        str(row['No. HP']),
      join_date:    date(row['Tgl Masuk']),
      status:       str(row['Status']) || 'Aktif',
      notes:        str(row['Catatan']),
    }),
  },
  contracts: {
    required: ['Nama Karyawan', 'Tgl Mulai', 'Tgl Selesai'],
    map: (row) => ({
      employee_name: str(row['Nama Karyawan']),
      branch_name:   str(row['Cabang']),
      division:      str(row['Divisi']) || 'FACILITY CARE',
      start_date:    date(row['Tgl Mulai']),
      end_date:      date(row['Tgl Selesai']),
      contract_type: str(row['Tipe Kontrak']),
      pkwt_number:   str(row['PKWT']),
      status:        str(row['Status']) || 'Aktif',
      notes:         str(row['Catatan']),
    }),
  },
  relievers: {
    required: ['Reliefer', 'Tanggal Backup'],
    map: (row) => ({
      branch_name:      str(row['Cabang']),
      original_fc_name: str(row['FC Digantikan']),
      period:           str(row['Periode']),
      reliever_name:    str(row['Reliefer']),
      backup_date:      date(row['Tanggal Backup']),
      completion_date:  date(row['Tanggal Selesai']),
      reason:           str(row['Keterangan']),
      shift:            str(row['Shift']),
      status:           str(row['Status']) || 'Pending',
    }),
  },
  schedule: {
    required: ['Kegiatan', 'Periode'],
    map: (row) => ({
      branch_name:     str(row['Cabang']),
      activity_type:   str(row['Kegiatan']),
      period:          str(row['Periode']),
      pic:             str(row['PIC']),
      opening_date:    date(row['Tgl Opening']),
      target_date:     date(row['Tgl Target']),
      completion_date: date(row['Tgl Selesai']),
      status:          str(row['Status']) || 'Pending',
      notes:           str(row['Catatan']),
    }),
  },
  issues: {
    required: ['Tanggal', 'Keluhan', 'Kategori'],
    map: (row) => ({
      report_date:     date(row['Tanggal']),
      branch_name:     str(row['Cabang']),
      category:        str(row['Kategori']),
      source:          str(row['Sumber']),
      complaint:       str(row['Keluhan']),
      employee_name:   str(row['Nama FC']),
      fc_specialist:   str(row['FC Spesialis']),
      solution:        str(row['Solusi']),
      completion_date: date(row['Tgl Selesai']),
      status:          str(row['Status']) || 'Open',
    }),
  },
  one_on_one: {
    required: ['Tanggal', 'Nama Karyawan', 'Permasalahan'],
    map: (row) => ({
      meeting_date:    date(row['Tanggal']),
      branch_name:     str(row['Cabang']),
      employee_name:   str(row['Nama Karyawan']),
      pic:             str(row['PIC']),
      problem:         str(row['Permasalahan']),
      solution:        str(row['Solusi']),
      completion_date: date(row['Tgl Selesai']),
      status:          str(row['Status']) || 'Open',
    }),
  },
  training: {
    required: ['Tanggal', 'Materi'],
    map: (row) => ({
      training_date: date(row['Tanggal']),
      batch:         str(row['Batch']),
      subject:       str(row['Materi']),
      participants:  str(row['Peserta']),
      branch_name:   str(row['Cabang']),
      trainer:       str(row['Trainer']),
      score:         num(row['Nilai']),
      notes:         str(row['Catatan']),
    }),
  },
  checklist: {
    required: ['Nama'],
    map: (row) => ({
      name:          str(row['Nama']),
      category:      str(row['Kategori']),
      document_link: str(row['Link Dokumen']),
      description:   str(row['Deskripsi']),
    }),
  },
  forms: {
    required: ['Nama'],
    map: (row) => ({
      name:          str(row['Nama']),
      category:      str(row['Kategori']),
      document_link: str(row['Link Dokumen']),
      description:   str(row['Deskripsi']),
    }),
  },
  sop: {
    required: ['Nama', 'Kategori'],
    map: (row) => ({
      name:           str(row['Nama']),
      category:       str(row['Kategori']),
      document_link:  str(row['Link Dokumen']),
      version:        str(row['Versi']),
      effective_date: date(row['Tgl Berlaku']),
      notes:          str(row['Catatan']),
    }),
  },
  inspection: {
    required: ['Tanggal', 'Periode'],
    map: (row) => ({
      inspection_date: date(row['Tanggal']),
      branch_name:     str(row['Cabang']),
      period:          str(row['Periode']),
      status:          str(row['Status']) || 'Pending',
      fc_score:        num(row['Nilai FC']),
      spv_score:       num(row['Nilai SPV']),
      document_link:   str(row['Link Dokumen']),
      notes:           str(row['Catatan']),
    }),
  },
  cleaning: {
    required: ['Tanggal', 'Periode'],
    map: (row) => ({
      activity_date: date(row['Tanggal']),
      branch_name:   str(row['Cabang']),
      activity_type: str(row['Jenis Kegiatan']) || 'General Cleaning',
      period:        str(row['Periode']),
      status:        str(row['Status']) || 'Pending',
      document_link: str(row['Link Dokumen']),
      notes:         str(row['Catatan']),
    }),
  },
  fogging: {
    required: ['Periode'],
    map: (row) => ({
      activity_date: date(row['Tanggal']),
      branch_name:   str(row['Cabang']),
      period:        str(row['Periode']),
      status:        str(row['Status']) || 'Pending',
      document_link: str(row['Link Dokumen']),
      notes:         str(row['Catatan']),
    }),
  },
  basecamp: {
    required: ['Tanggal Info', 'Permasalahan'],
    map: (row) => ({
      info_date:   date(row['Tanggal Info']),
      branch_name: str(row['Cabang']),
      problem:     str(row['Permasalahan']),
      pic:         str(row['PIC']),
      done_date:   date(row['Tgl Selesai']),
      status:      str(row['Status']) || 'Open',
      notes:       str(row['Catatan']),
    }),
  },
  supply: {
    required: ['Tanggal', 'Nama Pemohon'],
    map: (row) => ({
      submitted_at:      date(row['Tanggal']),
      submitter_name:    str(row['Nama Pemohon']),
      branch_name:       str(row['Cabang']),
      tools_items:       str(row['Item Tools']),
      tools_quantity:    str(row['Qty Tools']),
      chemical_items:    str(row['Item Chemical']),
      chemical_quantity: str(row['Qty Chemical']),
      additional_notes:  str(row['Catatan']),
      status:            str(row['Status']) || 'Pending',
    }),
  },
};

// ─── Helpers ────────────────────────────────────────────────────────────────
function str(v) {
  if (v === undefined || v === null || v === '') return null;
  return String(v).trim() || null;
}

function date(v) {
  if (v === undefined || v === null || v === '') return null;
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  // Excel serial number
  if (/^\d{4,5}$/.test(s)) {
    const d = new Date(Date.UTC(1899, 11, 30) + Number(s) * 86400000);
    return d.toISOString().slice(0, 10);
  }
  // dd/mm/yyyy
  const parts = s.split(/[\/\-\.]/);
  if (parts.length === 3) {
    const [a, b, c] = parts;
    if (c.length === 4) return `${c}-${b.padStart(2, '0')}-${a.padStart(2, '0')}`;
  }
  return s;
}

function num(v) {
  if (v === undefined || v === null || v === '') return null;
  const n = parseFloat(v);
  return isNaN(n) ? null : n;
}

function isValidDate(s) {
  if (!s) return true; // optional
  return /^\d{4}-\d{2}-\d{2}$/.test(s) && !isNaN(new Date(s));
}

// ─── Main Validator ─────────────────────────────────────────────────────────
/**
 * Validates a single parsed sheet.
 * @param {string} sheetName
 * @param {Array}  rawRows  — array of plain objects (header row as keys)
 * @param {Object} context  — { employeeNames: Set } for cross-sheet validation
 * @returns {{ valid: Array, errors: Array, mapped: Array }}
 */
export function validateSheet(sheetName, rawRows, context = {}) {
  const mapping = SHEET_MAP[sheetName];
  if (!mapping) return { valid: [], errors: [], mapped: [], skipped: true };

  const schema = SCHEMAS[mapping.module];
  if (!schema) return { valid: [], errors: [], mapped: [], skipped: true };

  const validRows = [];
  const errorRows = [];
  const mappedRows = [];

  rawRows.forEach((raw, idx) => {
    const rowNum = idx + 2; // +2 because row 1 is header
    const errs = [];

    // Check required fields
    schema.required.forEach(field => {
      if (!raw[field] && raw[field] !== 0) {
        errs.push(`Kolom "${field}" wajib diisi`);
      }
    });

    // Map the row
    const mapped = schema.map(raw);

    // Date validation
    Object.entries(mapped).forEach(([k, v]) => {
      if (k.endsWith('_date') && v && !isValidDate(v)) {
        errs.push(`Format tanggal tidak valid pada kolom "${k}": "${v}"`);
      }
    });

    // Cross-sheet validation: contracts need valid employee name
    if (mapping.module === 'contracts' && mapped.employee_name) {
      if (context.employeeNames && context.employeeNames.size > 0) {
        if (!context.employeeNames.has(mapped.employee_name.toLowerCase())) {
          errs.push(`Nama karyawan "${mapped.employee_name}" tidak ditemukan di sheet Master Karyawan`);
        }
      }
    }

    if (errs.length > 0) {
      errorRows.push({ row: rowNum, data: mapped, errors: errs });
    } else {
      validRows.push(raw);
      mappedRows.push(mapped);
    }
  });

  return { valid: validRows, errors: errorRows, mapped: mappedRows };
}

/**
 * Validates all sheets from the workbook.
 * @param {Object} workbook  — SheetJS workbook object
 * @returns {Array} results  — per-sheet validation result
 */
export function validateWorkbook(workbook) {
  const results = [];

  // First pass: parse employees to build cross-ref map
  const employeeNames = new Set();
  const empSheetName = Object.keys(SHEET_MAP).find(n => SHEET_MAP[n].module === 'employees' && workbook.SheetNames.includes(n));
  if (empSheetName) {
    const ws = workbook.Sheets[empSheetName];
    const rows = window.XLSX.utils.sheet_to_json(ws, { defval: '' });
    rows.forEach(r => {
      const n = str(r['Nama Lengkap']);
      if (n) employeeNames.add(n.toLowerCase());
    });
  }

  const context = { employeeNames };

  // Process each recognised sheet in the defined order
  const IMPORT_ORDER = Object.keys(SHEET_MAP);
  const processed = new Set();

  IMPORT_ORDER.forEach(sheetName => {
    if (!workbook.SheetNames.includes(sheetName)) return;
    processed.add(sheetName);

    const ws = workbook.Sheets[sheetName];
    const rawRows = window.XLSX.utils.sheet_to_json(ws, { defval: '' });

    const result = validateSheet(sheetName, rawRows, context);
    results.push({
      sheetName,
      module:   SHEET_MAP[sheetName].module,
      label:    SHEET_MAP[sheetName].label,
      total:    rawRows.length,
      valid:    result.mapped.length,
      errorCount: result.errors.length,
      errors:   result.errors,
      mapped:   result.mapped,
      skipped:  result.skipped || false,
    });
  });

  // Warn about unrecognised sheets
  workbook.SheetNames.forEach(name => {
    if (!processed.has(name) && name !== 'Validasi' && !name.startsWith('_')) {
      results.push({
        sheetName: name,
        module: null,
        label: '(Tidak dikenali)',
        total: 0, valid: 0, errorCount: 0, errors: [], mapped: [],
        skipped: true,
      });
    }
  });

  return results;
}

// ─── Template Generator ─────────────────────────────────────────────────────
export function generateTemplate() {
  const XLSX = window.XLSX;
  const wb = XLSX.utils.book_new();

  const TEMPLATES = {
    'Master Karyawan': [{ 'Nama Lengkap': 'Budi Santoso', 'Cabang': '001. Pondok Bambu', 'Divisi': 'FACILITY CARE', 'No. HP': '081234567890', 'Tgl Masuk': '2024-01-15', 'Status': 'Aktif', 'Catatan': '' }],
    'Data Kontrak': [{ 'Nama Karyawan': 'Budi Santoso', 'Cabang': '001. Pondok Bambu', 'Divisi': 'FACILITY CARE', 'Tgl Mulai': '2024-01-01', 'Tgl Selesai': '2024-12-31', 'Tipe Kontrak': 'KONTRAK 1 TAHUN', 'PKWT': 'PKWT 1', 'Status': 'Aktif', 'Catatan': '' }],
    'Jadwal Reliefer': [{ 'Tanggal Backup': '2024-03-10', 'Cabang': '001. Pondok Bambu', 'FC Digantikan': 'Budi Santoso', 'Periode': 'Q1', 'Reliefer': 'Andi', 'Keterangan': 'Sakit', 'Shift': 'Pagi', 'Tanggal Selesai': '2024-03-10', 'Status': 'Done' }],
    'Time Line': [{ 'Cabang': '001. Pondok Bambu', 'Kegiatan': 'General Cleaning', 'Periode': 'Q1', 'PIC': 'Fajar', 'Tgl Opening': '2024-02-01', 'Tgl Target': '2024-02-15', 'Tgl Selesai': '2024-02-14', 'Status': 'Done', 'Catatan': '' }],
    'Permasalahan': [{ 'Tanggal': '2024-03-01', 'Cabang': '001. Pondok Bambu', 'Kategori': 'Cleaning', 'Sumber': 'SPV', 'Keluhan': 'Lantai kotor', 'Nama FC': 'Budi', 'FC Spesialis': 'Fajar', 'Solusi': 'Teguran', 'Tgl Selesai': '2024-03-02', 'Status': 'Done' }],
    'One on One': [{ 'Tanggal': '2024-03-05', 'Cabang': '001. Pondok Bambu', 'Nama Karyawan': 'Budi Santoso', 'PIC': 'Berlin', 'Permasalahan': 'Keterlambatan', 'Solusi': 'Coaching', 'Tgl Selesai': '2024-03-06', 'Status': 'Done' }],
    'Training': [{ 'Tanggal': '2024-02-20', 'Batch': 'Batch 1', 'Materi': 'Hygiene & Sanitasi', 'Peserta': 'Budi, Andi', 'Cabang': '001. Pondok Bambu', 'Trainer': 'Fajar', 'Nilai': '85', 'Catatan': '' }],
    'Master Checklist': [{ 'Nama': 'Checklist Harian Toilet', 'Kategori': 'Harian', 'Link Dokumen': 'https://docs.google.com/...', 'Deskripsi': '' }],
    'Master Form': [{ 'Nama': 'Form Permintaan Chemical', 'Kategori': 'Chemical', 'Link Dokumen': 'https://docs.google.com/...', 'Deskripsi': '' }],
    'SOP': [{ 'Nama': 'SOP Pembersihan Toilet', 'Kategori': 'Cleaning', 'Link Dokumen': 'https://docs.google.com/...', 'Versi': '1.0', 'Tgl Berlaku': '2024-01-01', 'Catatan': '' }],
    'Report Inspeksi Hygiene 2026': [{ 'Tanggal': '2024-01-15', 'Cabang': '001. Pondok Bambu', 'Periode': 'Q1', 'Status': 'Done', 'Nilai FC': '85', 'Nilai SPV': '90', 'Link Dokumen': '', 'Catatan': '' }],
    'Report GCDC 2026': [{ 'Tanggal': '2024-02-10', 'Cabang': '001. Pondok Bambu', 'Jenis Kegiatan': 'General Cleaning', 'Periode': 'Q1', 'Status': 'Done', 'Link Dokumen': '', 'Catatan': '' }],
    'Report Fogging 2026': [{ 'Tanggal': '2024-03-20', 'Cabang': '001. Pondok Bambu', 'Periode': 'Q1', 'Status': 'Done', 'Link Dokumen': '', 'Catatan': '' }],
    'Rekap Laporan Basecamp': [{ 'Tanggal Info': '2024-01-10', 'Cabang': '001. Pondok Bambu', 'Permasalahan': 'AC Rusak', 'PIC': 'Berlin', 'Tgl Selesai': '2024-01-15', 'Status': 'Done', 'Catatan': '' }],
    'Permintaan Chemical': [{ 'Tanggal': '2024-02-01', 'Nama Pemohon': 'Budi Santoso', 'Cabang': '001. Pondok Bambu', 'Item Tools': 'Sapu', 'Qty Tools': '2', 'Item Chemical': 'Sabun Lantai', 'Qty Chemical': '5 liter', 'Catatan': '', 'Status': 'Pending' }],
  };

  Object.entries(TEMPLATES).forEach(([sheetName, data]) => {
    const ws = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  });

  XLSX.writeFile(wb, 'Template_Import_Data_Awal_FCMS.xlsx');
}

// ─── Error Log Generator ────────────────────────────────────────────────────
export function downloadErrorLog(results) {
  const XLSX = window.XLSX;
  const wb = XLSX.utils.book_new();
  let hasErrors = false;

  results.forEach(r => {
    if (!r.errors || r.errors.length === 0) return;
    hasErrors = true;
    const rows = r.errors.map(e => ({
      'No. Baris': e.row,
      'Penyebab Error': e.errors.join('; '),
      ...Object.fromEntries(Object.entries(e.data || {}).map(([k, v]) => [k, v ?? ''])),
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const safeSheet = r.sheetName.replace(/[\\\/\[\]*?:]/g, '_').slice(0, 31);
    XLSX.utils.book_append_sheet(wb, ws, safeSheet);
  });

  if (!hasErrors) return false;
  XLSX.writeFile(wb, `Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`);
  return true;
}
