/**
 * importer.js  v4 — Exact Matching Validation Engine
 * 
 * Filosofi v4:
 * - JANGAN blokir data karena kolom opsional kosong.
 * - HANYA required fields yang diperiksa.
 * - Tidak ada Alias, Rename, Auto-mapping, atau Fuzzy matching.
 * - Kolom di-map secara literal sesuai header Excel.
 */

// ─── Sheet → Module Mapping ──────────────────────────────────────────────────
export const SHEET_MAP = {
  'Validasi':                     { module: 'validation', label: 'Master Referensi' },
  'SOP':                          { module: 'sop',        label: 'SOP' },
  'Master Karyawan':              { module: 'employees',  label: 'Karyawan' },
  'Data Kontrak':                 { module: 'contracts',  label: 'Kontrak' },
  'Permasalahan':                 { module: 'issues',     label: 'Permasalahan' },
  'One on One':                   { module: 'one_on_one', label: 'One on One' },
  'Time Line':                    { module: 'schedule',   label: 'Jadwal Kegiatan' },
  'Report Inspeksi Hygiene 2026': { module: 'inspection', label: 'Laporan Inspeksi' },
  'Report GC-DC 2026':            { module: 'cleaning',   label: 'Laporan GC/DC' },
  'Report Fogging 2026':          { module: 'fogging',    label: 'Laporan Fogging' },
  'Rekap Laporan Basecamp':       { module: 'basecamp',   label: 'Rekap Basecamp' },
  'Jadwal Reliefer':              { module: 'relievers',  label: 'Reliefer' },
  'Training':                     { module: 'training',   label: 'Training' },
  'Master Checklist':             { module: 'checklist',  label: 'Checklist' },
  'Master Form':                  { module: 'forms',      label: 'Master Form' },
  'Permintaan Chemical':          { module: 'supply',     label: 'Inventory Chemical' },
};

// ─── Date Parsing Helper ─────────────────────────────────────────────────────
function dateStr(v) {
  if (v === undefined || v === null || v === '') return null;

  // JS Date object (from cellDates: true)
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return null;
    // Use UTC to avoid timezone shifting
    return v.toISOString().slice(0, 10);
  }

  const s = String(v).trim();
  if (s === '' || s === '0') return null;

  // ISO: YYYY-MM-DD (exact match preferred)
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);

  // Excel serial number (plain integer 20000–99999)
  if (/^\d{4,5}$/.test(s)) {
    const n = Number(s);
    if (n > 20000 && n < 99999) {
      const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
      return isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10);
    }
  }

  // Parts-based parsing: DD/MM/YYYY, MM/DD/YYYY, DD-MM-YYYY, etc.
  const parts = s.split(/[\/\-\.]/);
  if (parts.length === 3) {
    const [a, b, c] = parts.map(p => p.trim());
    const na = Number(a), nb = Number(b), nc = Number(c);

    // YYYY-MM-DD or YYYY/MM/DD
    if (a.length === 4 && na > 1900) {
      return `${a}-${b.padStart(2,'0')}-${c.padStart(2,'0')}`;
    }

    // DD/MM/YYYY or MM/DD/YYYY — for Indonesia we default to DD/MM/YYYY
    if (c.length === 4 && nc > 1900) {
      // If 'a' > 12, it must be a day (DD/MM/YYYY)
      if (na > 12) return `${c}-${b.padStart(2,'0')}-${a.padStart(2,'0')}`;
      // If 'b' > 12, it must be a day (MM/DD/YYYY)
      if (nb > 12) return `${c}-${a.padStart(2,'0')}-${b.padStart(2,'0')}`;
      // Default: assume DD/MM/YYYY (Indonesian convention)
      return `${c}-${b.padStart(2,'0')}-${a.padStart(2,'0')}`;
    }

    // 2-digit year: DD/MM/YY
    if (c.length === 2 && !isNaN(nc)) {
      const year = nc >= 50 ? `19${c}` : `20${c}`;
      if (na > 12) return `${year}-${b.padStart(2,'0')}-${a.padStart(2,'0')}`;
      return `${year}-${b.padStart(2,'0')}-${a.padStart(2,'0')}`;
    }
  }

  // Fallback: try native Date parse (handles "Jul 1, 2026" etc.)
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  return null;
}

function isEmptyRow(row) {
  return Object.values(row).every(v =>
    v === undefined || v === null || String(v).trim() === ''
  );
}

// ─── Module Schemas ──────────────────────────────────────────────────────────
const SCHEMAS = {
  validation: {
    required: [],
    map: (row) => ({
      cabang:        row['CABANG'],
      pic:           row['PIC'],
      kegiatan:      row['KEGIATAN'],
      quartal:       row['QUARTAL'],
      masa_pkwt:     row['MASA PKWT'],
      pic_pelapor:   row['PIC PELAPOR'],
      kontrak:       row['KONTRAK']
    })
  },
  sop: {
    required: [{ key: 'Nama SOP', label: 'Nama SOP' }],
    map: (row) => ({
      name:           row['Nama SOP'],
      category:       row['Kategori'] || 'Umum',
      document_link:  row['Link Document'],
      version:        '1.0',
      effective_date: null,
      notes:          '',
    })
  },
  employees: {
    required: [{ key: 'Nama Lengkap', label: 'Nama Lengkap' }],
    map: (row) => ({
      full_name:   row['Nama Lengkap'],
      branch_name: row['Cabang'],
      division:    row['Div / Bagian'] || 'FACILITY CARE',
      phone:       row['No. Hp'],
      join_date:   dateStr(row['Tanggal Masuk']),
      status:      row['Status'] || '',
      notes:       '',
    }),
  },
  contracts: {
    required: [{ key: 'Nama Lengkap', label: 'Nama Lengkap' }],
    map: (row) => ({
      employee_name: row['Nama Lengkap'],
      branch_name:   row['Cabang'],
      division:      row['Div / Bagian'] || 'FACILITY CARE',
      start_date:    dateStr(row['Tanggal Mulai']),
      end_date:      dateStr(row['Tanggal Selesai']),
      contract_type: row['Tipe Kontrak'] || '',
      pkwt_number:   row['PKWT'] || '',
      status:        row['Status'] || '',
      notes:         row['keterangan'],
    }),
  },
  issues: {
    required: [{ key: 'Keluhan', label: 'Keluhan' }],
    map: (row) => ({
      report_date:     dateStr(row['Tanggal Info']),
      branch_name:     row['Cabang'],
      category:        row['Kategori'],
      source:          row['Sumber Laporan'],
      complaint:       row['Keluhan'],
      employee_name:   row['Nama FC'],
      fc_specialist:   row['FC Spesialis'],
      solution:        row['Solusi'],
      status:          row['Status'] || '',
      completion_date: dateStr(row['Tanggal Selesai']),
    }),
  },
  one_on_one: {
    required: [],
    map: (row) => ({
      meeting_date:    dateStr(row['Tanggal']),
      branch_name:     row['Cabang'],
      employee_name:   row['Nama Karyawan'],
      pic:             row['Pic'],
      problem:         row['Masalah'],
      solution:        row['Solusi'],
      status:          row['Status'] || '',
      completion_date: dateStr(row['Tanggal Selesai']),
      document_link:   row['Link Document'],
    }),
  },
  schedule: {
    required: [{ key: 'Kegiatan', label: 'Kegiatan' }],
    map: (row) => ({
      branch_name:     row['Cabang'],
      activity_type:   row['Kegiatan'],
      period:          row['Periode'],
      pic:             row['Pic'] || row['PIC'],
      opening_date:    dateStr(row['Tanggal Opening'] || row['Tgl Opening']),
      target_date:     dateStr(row['Tanggal Target'] || row['Tgl Target']),
      completion_date: dateStr(row['Tanggal Selesai'] || row['Tgl Selesai']),
      status:          row['Status'] || '',
      notes:           row['Keterangan'] || row['Catatan'],
    }),
  },
  inspection: {
    required: [],
    map: (row) => ({
      inspection_date: dateStr(row['Tanggal']),
      branch_name:     row['Cabang'],
      period:          row['Periode'],
      status:        row['Status'] || '',
      fc_score:        row['Point FC SP'] !== undefined && row['Point FC SP'] !== null ? parseFloat(String(row['Point FC SP']).replace(',', '.')) : null,
      spv_score:       row['Point SPV'] !== undefined && row['Point SPV'] !== null ? parseFloat(String(row['Point SPV']).replace(',', '.')) : null,
      document_link:   row['Link'],
      notes:           '',
    }),
  },
  cleaning: {
    required: [],
    map: (row) => ({
      activity_date: dateStr(row['Tanggal']),
      branch_name:   row['Cabang'],
      activity_type: row['Jenis Kegiatan'] || 'General Cleaning',
      period:        row['Periode'],
      status:        row['Status'] || '',
      document_link: row['Link'],
      notes:         '',
    }),
  },
  fogging: {
    required: [],
    map: (row) => ({
      activity_date: dateStr(row['Tanggal']),
      branch_name:   row['Cabang'],
      period:        row['Periode'],
      status:        row['Status'] || '',
      document_link: row['Link'],
      notes:         '',
    }),
  },
  basecamp: {
    required: [{ key: 'Permasalahan', label: 'Permasalahan' }],
    map: (row) => ({
      info_date:   dateStr(row['Tgl Info']),
      branch_name: row['Cabang'],
      problem:     row['Permasalahan'],
      pic:         row['PIC'],
      done_date:   dateStr(row['Tgl Done']),
      status:      row['Status'] || '',
      notes:       row['Ket'],
    }),
  },
  relievers: {
    required: [],
    map: (row) => ({
      branch_name:      row['Cabang'],
      original_fc_name: row['Nama Facility care'],
      period:           row['Periode'],
      reliever_name:    row['Relifer'],
      backup_date:      dateStr(row['Tanggal Back Up']),
      completion_date:  dateStr(row['Tanggal Selesai']),
      reason:           row['Keterangan'],
      shift:            row['Shift'],
      status:           row['Status'] || '',
    }),
  },
  training: {
    required: [{ key: 'Materi', label: 'Materi' }],
    map: (row) => ({
      training_date: dateStr(row['Tanggal']),
      batch:         row['Batch'],
      subject:       row['Materi'],
      participants:  row['Peserta'],
      branch_name:   row['Cabang'],
      trainer:       row['Trainer'],
      score:         row['Nilai'] !== undefined && row['Nilai'] !== null ? parseFloat(String(row['Nilai']).replace(',', '.')) : null,
      notes:         '',
    }),
  },
  checklist: {
    required: [],
    map: (row) => ({
      name:          row['Master Checklist'],
      category:      'Umum',
      document_link: row['Link Document'],
      description:   '',
    }),
  },
  forms: {
    required: [{ key: 'Master Form', label: 'Master Form' }],
    map: (row) => ({
      name:          row['Master Form'],
      category:      'Umum',
      document_link: row['Link Document'],
      description:   '',
    }),
  },
  supply: {
    required: [],
    map: (row) => ({
      submitted_at:      dateStr(row['Timestamp']),
      submitter_name:    row['Nama Lengkap'],
      branch_name:       row['Kebutuhan Untuk Cabang'],
      tools_items:       row['Alat - Alat / Barang'],
      tools_quantity:    row['Jumlah Permintaan Alat / Barang'],
      chemical_items:    row['Chemical'],
      chemical_quantity: row['Jumlah Permintaan Chemical'],
      additional_notes:  row['Tambahan  Alat / Chemical Jika Ada Permintaan Diluar List.'],
      status:            row['Status'] || '',
    }),
  },
};

// ─── Main Validator ───────────────────────────────────────────────────────────
export function validateSheet(sheetName, rawRows) {
  const mapping = SHEET_MAP[sheetName];
  if (!mapping) return { valid: [], errors: [], mapped: [], skipped: true };

  const schema = SCHEMAS[mapping.module];
  if (!schema) return { valid: [], errors: [], mapped: [], skipped: true };

  const validRows  = [];
  const errorRows  = [];
  const mappedRows = [];

  const nonEmptyRows = rawRows.filter(r => !isEmptyRow(r));

  nonEmptyRows.forEach((raw, idx) => {
    const rowNum = rawRows.indexOf(raw) + 2;
    const errs   = [];

    schema.required.forEach(({ key, label }) => {
      const val = raw[key];
      if (val === undefined || val === null || String(val).trim() === '') {
        const existingCols = Object.keys(raw).filter(k => k.trim()).join(', ');
        errs.push({
          column: label,
          originalValue: val || '',
          reason: `Kolom "${label}" wajib diisi dan tidak ditemukan`,
          hint: `Kolom yang tersedia: ${existingCols.slice(0, 120)}`
        });
      }
    });

    const mapped = schema.map(raw);

    if (errs.length > 0) {
      errorRows.push({ row: rowNum, data: mapped, raw, errors: errs });
    } else {
      validRows.push(raw);
      mappedRows.push(mapped);
    }
  });

  return { valid: validRows, errors: errorRows, mapped: mappedRows };
}

export function validateWorkbook(workbook) {
  const results = [];

  workbook.SheetNames.forEach(sheetName => {
    const mapping = SHEET_MAP[sheetName];
    if (!mapping) return; // Ignore unmapped sheets silently

    const ws      = workbook.Sheets[sheetName];
    const rawRows = window.XLSX.utils.sheet_to_json(ws, { defval: '', raw: false, dateNF: 'yyyy-mm-dd' });
    const result  = validateSheet(sheetName, rawRows);

    const nonEmpty = rawRows.filter(r => !isEmptyRow(r));

    results.push({
      sheetName,
      module:     mapping.module,
      label:      mapping.label,
      total:      nonEmpty.length,
      valid:      result.mapped.length,
      errorCount: result.errors.length,
      errors:     result.errors,
      mapped:     result.mapped,
      skipped:    false,
    });
  });

  return results;
}

export function generateTemplate() {
  const XLSX = window.XLSX;
  const wb   = XLSX.utils.book_new();
  const TEMPLATES = {
    'Validasi': [{ 'CABANG': '001. Pondok Bambu', 'NAMA KARYAWAN': 'Budi Santoso', 'PIC': 'Berlin', 'KEGIATAN': 'General Cleaning', 'QUARTAL': 'Q1', 'PIC PELAPOR': 'Berlin', 'KONTRAK': 'PKWT 1', 'MASA PKWT': '1 Tahun' }],
    'SOP': [{ 'Nama SOP': 'SOP Pembersihan Toilet', 'Kategori': 'Cleaning', 'Link Document': 'https://...' }],
    'Master Karyawan': [{ 'Nama Lengkap': 'Budi Santoso', 'Cabang': '001. Pondok Bambu', 'Div / Bagian': 'FACILITY CARE', 'No. Hp': '081234567890', 'Tanggal Masuk': '2024-01-15', 'Status': 'Aktif' }],
    'Data Kontrak': [{ 'Nama Lengkap': 'Budi Santoso', 'Cabang': '001. Pondok Bambu', 'Div / Bagian': 'FACILITY CARE', 'Tanggal Mulai': '2024-01-01', 'Tanggal Selesai': '2024-12-31', 'Sisa Kontrak': '365', 'Status': 'Aktif', 'keterangan': '' }],
    'Permasalahan': [{ 'Tanggal Info': '2024-03-01', 'Cabang': '001. Pondok Bambu', 'Kategori': 'Cleaning', 'Sumber Laporan': 'SPV', 'Keluhan': 'Lantai kotor', 'Nama FC': 'Budi', 'FC Spesialis': 'Fajar', 'Solusi': 'Teguran', 'Status': 'Done', 'Tanggal Selesai': '2024-03-02', 'Day': '1' }],
    'One on One': [{ 'Tanggal': '2024-03-05', 'Cabang': '001. Pondok Bambu', 'Nama Karyawan': 'Budi Santoso', 'Pic': 'Berlin', 'Masalah': 'Keterlambatan', 'Solusi': 'Coaching', 'Status': 'Done', 'Tanggal Selesai': '2024-03-06', 'Day': '1', 'Link Document': '' }],
  };
  Object.entries(TEMPLATES).forEach(([sn, data]) => {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), sn);
  });
  XLSX.writeFile(wb, 'Template_Import_Data_Awal_FCMS.xlsx');
}

export function downloadErrorLog(results) {
  const XLSX = window.XLSX;
  const wb   = XLSX.utils.book_new();
  let hasErrors = false;

  results.forEach(r => {
    if (!r.errors || r.errors.length === 0) return;
    hasErrors = true;
    const rows = r.errors.map(e => ({
      'No. Baris':      e.row,
      'Kolom Gagal':    (e.errors || []).map(err => err.column || err).join('; '),
      'Alasan Error':   (e.errors || []).map(err => err.reason || err).join('; '),
      ...Object.fromEntries(Object.entries(e.data || {}).map(([k,v]) => [k, v ?? ''])),
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, r.sheetName.replace(/[\\\/\[\]*?:]/g,'_').slice(0,31));
  });

  if (!hasErrors) return false;
  XLSX.writeFile(wb, `Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`);
  return true;
}
