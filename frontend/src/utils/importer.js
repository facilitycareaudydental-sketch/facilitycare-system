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
  'Time Line':                    { module: 'schedule',   label: 'Jadwal Kegiatan' },
  'Report Inspeksi Hygiene 2026': { module: 'inspection', label: 'Laporan Inspeksi' },
  'Report GC-DC 2026':            { module: 'cleaning',   label: 'Laporan GC/DC' },
  'Report Fogging 2026':          { module: 'fogging',    label: 'Laporan Fogging' },
  'Rekap Laporan Basecamp':       { module: 'basecamp',   label: 'Rekap Basecamp' },
  'Training':                     { module: 'training',   label: 'Training' },
  'Master Checklist':             { module: 'checklist',  label: 'Checklist' },
  'Data SP':                      { module: 'sp',         label: 'Data SP' },
  'Data Mutasi':                  { module: 'mutasi',     label: 'Data Mutasi' },
  'Master Form':                  { module: 'forms',      label: 'Master Form' },
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
      full_name:         row['Nama Lengkap'],
      branch_name:       row['Cabang'],
      division:          row['Div / Bagian'] || row['Divisi'] || 'FACILITY CARE',
      phone:             row['No. Hp'] || row['No. HP'],
      join_date:         dateStr(row['Tanggal Masuk'] || row['Tgl Masuk']),
      status:            row['Status'] || '',
      target_pindah_os:  dateStr(row['Target Pindah OS']),
      target_selesai_os: dateStr(row['Target Selesai OS']),
      notes:             row['Catatan'] || '',
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
    map: (row) => {
      const rawFc = row['Point FC SP'] !== undefined && row['Point FC SP'] !== null && row['Point FC SP'] !== '' ? row['Point FC SP'] : row['Point FC'];
      return {
        inspection_date: dateStr(row['Tanggal']),
        branch_name:     row['Cabang'],
        period:          row['Periode'],
        status:          row['Status'] || '',
        fc_score:        rawFc !== undefined && rawFc !== null && rawFc !== '' ? parseFloat(String(rawFc).replace(',', '.')) : null,
        spv_score:       row['Point SPV'] !== undefined && row['Point SPV'] !== null && row['Point SPV'] !== '' ? parseFloat(String(row['Point SPV']).replace(',', '.')) : null,
        document_link:   row['Link'] || row['Link Dokumen'] || '',
        notes:           row['Keterangan'] || '',
      };
    },
  },
  cleaning: {
    required: [],
    map: (row) => ({
      activity_date: dateStr(row['Tanggal']),
      branch_name:   row['Cabang'],
      activity_type: row['Jenis Kegiatan'] || 'General Cleaning',
      period:        row['Periode'],
      status:        row['Status'] || '',
      document_link: row['Link'] || row['Link Dokumen'] || '',
      notes:         row['Keterangan'] || '',
    }),
  },
  fogging: {
    required: [],
    map: (row) => ({
      activity_date: dateStr(row['Tanggal']),
      branch_name:   row['Cabang'],
      period:        row['Periode'],
      status:        row['Status'] || '',
      document_link: row['Link'] || row['Link Dokumen'] || '',
      notes:         row['Keterangan'] || '',
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
  sp: {
    required: [{ key: 'Nama Karyawan', label: 'Nama Karyawan' }],
    map: (row) => ({
      tanggal: dateStr(row['Tanggal'] || row['Tanggal Terbit']),
      employee_name: row['Nama Karyawan'],
      branch_name: row['Cabang'],
      sp_type: row['Jenis SP'] || 'SP 1',
      status: row['Status'] || 'Aktif',
      document_link: row['Link Dokumen'] || row['Link Document'] || null,
      notes: row['Keterangan'] || row['Catatan'] || null,
    }),
  },
  mutasi: {
    required: [{ key: 'Nama Karyawan', label: 'Nama Karyawan' }],
    map: (row) => ({
      tanggal: dateStr(row['Tanggal'] || row['Tanggal Efektif']),
      employee_name: row['Nama Karyawan'],
      from_branch_name: row['Cabang Asal'] || row['Dari Cabang'],
      to_branch_name: row['Cabang Tujuan'] || row['Ke Cabang'],
      status: row['Status'] || 'Selesai',
      document_link: row['Link Dokumen'] || row['Link SK'] || null,
      notes: row['Keterangan'] || row['Alasan'] || null,
    }),
  },
  forms: {
    required: [{ key: 'Nama Form', label: 'Nama Form' }],
    map: (row) => ({
      name: row['Nama Form'],
      category: row['Kategori'] || 'Umum',
      document_link: row['Link Dokumen'] || row['Link Document'] || null,
      description: row['Deskripsi'] || row['Keterangan'] || null,
      is_public: row['Publik'] === 'Ya' || row['Publik'] === 1 || row['Akses Publik'] === 1 ? 1 : 0,
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
    'Validasi': [{ 'CABANG': '001. Pondok Bambu', 'PIC': 'Berlin', 'KEGIATAN': 'General Cleaning', 'QUARTAL': 'Q1', 'PIC PELAPOR': 'Berlin', 'KONTRAK': 'PKWT 1', 'MASA PKWT': '1 Tahun' }],
    'SOP': [{ 'Nama SOP': 'SOP Pembersihan Toilet', 'Kategori': 'Cleaning', 'Link Document': 'https://...' }],
    'Master Karyawan': [{ 'Nama Lengkap': 'Budi Santoso', 'Cabang': '000. Management', 'Div / Bagian': 'FACILITY CARE', 'No. Hp': '081234567890', 'Tanggal Masuk': '2024-01-15', 'Status': 'Aktif' }],
    'Data Kontrak': [{ 'Nama Lengkap': 'Budi Santoso', 'Cabang': '000. Management', 'Div / Bagian': 'FACILITY CARE', 'Tanggal Mulai': '2024-01-01', 'Tanggal Selesai': '2024-12-31', 'Tipe Kontrak': 'PKWT 1', 'PKWT': '001/PKWT/2024', 'Status': 'Aktif', 'keterangan': '' }],
    'Permasalahan': [{ 'Tanggal Info': '2024-03-01', 'Cabang': '001. Pondok Bambu', 'Kategori': 'Cleaning', 'Sumber Laporan': 'SPV', 'Keluhan': 'Lantai kotor', 'Nama FC': 'Budi', 'FC Spesialis': 'Fajar', 'Solusi': 'Teguran', 'Status': 'Done', 'Tanggal Selesai': '2024-03-02' }],
    'Time Line': [{ 'Cabang': '001. Pondok Bambu', 'Kegiatan': 'General Cleaning', 'Periode': 'Januari', 'Pic': 'Berlin', 'Tanggal Opening': '2024-01-01', 'Tanggal Target': '2024-01-10', 'Tanggal Selesai': '2024-01-09', 'Status': 'Done', 'Keterangan': '' }],
    'Report Inspeksi Hygiene 2026': [{ 'Tanggal': '2026-01-15', 'Cabang': '001. Pondok Bambu', 'Periode': 'Q1', 'Status': 'Done', 'Point FC SP': '85.5', 'Point SPV': '90.0', 'Link': 'https://...' }],
    'Report GC-DC 2026': [{ 'Tanggal': '2026-01-20', 'Cabang': '001. Pondok Bambu', 'Jenis Kegiatan': 'General Cleaning', 'Periode': 'Q1', 'Status': 'Done', 'Link': 'https://...' }],
    'Report Fogging 2026': [{ 'Tanggal': '2026-01-25', 'Cabang': '001. Pondok Bambu', 'Periode': 'Q1', 'Status': 'Done', 'Link': 'https://...' }],
    'Rekap Laporan Basecamp': [{ 'Tgl Info': '2024-02-01', 'Cabang': '001. Pondok Bambu', 'Permasalahan': 'Lampu mati', 'PIC': 'Berlin', 'Tgl Done': '2024-02-02', 'Status': 'Done', 'Ket': '' }],
    'Training': [{ 'Tanggal': '2024-04-10', 'Batch': 'Batch 1', 'Materi': 'Basic Cleaning', 'Peserta': '5', 'Cabang': '001. Pondok Bambu', 'Trainer': 'Fajar', 'Nilai': '85', 'Keterangan': '' }],
    'Master Checklist': [{ 'Master Checklist': 'Checklist Kebersihan Toilet', 'Link Document': 'https://...' }],
    'Data SP': [{ 'Nama Karyawan': 'Budi Santoso', 'Cabang': '001. Pondok Bambu', 'Jenis SP': 'SP 1', 'Tanggal Terbit': '2026-01-10', 'Tanggal Berakhir': '2026-07-10', 'Alasan / Pelanggaran': 'Keterlambatan', 'Status': 'Aktif', 'Link Dokumen': '' }],
    'Data Mutasi': [{ 'Nama Karyawan': 'Budi Santoso', 'Cabang Asal': '001. Pondok Bambu', 'Cabang Tujuan': '002. Bintaro', 'Tanggal Efektif': '2026-02-01', 'Alasan Mutasi': 'Rotasi Reguler', 'Status': 'Selesai', 'No. SK / Dokumen': 'SK/001/2026' }],
    'Master Form': [{ 'Nama Form': 'Form Pengajuan Alat', 'Kategori': 'Operasional', 'Link Form': 'https://...', 'Deskripsi': 'Formulir pengajuan alat dan perlengkapan', 'Publik': 'Ya' }],
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
