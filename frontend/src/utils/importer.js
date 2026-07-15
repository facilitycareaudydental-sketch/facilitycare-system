/**
 * importer.js  v2 — Flexible Header Mapping Engine
 *
 * Strategi: Setiap field database dipetakan ke BANYAK kemungkinan nama header Excel.
 * readField(row, aliases) mencoba setiap alias (case-insensitive, trim) sampai menemukan nilai.
 * Tidak perlu mengubah file Excel asli.
 */

// ─── Sheet → Module Mapping ──────────────────────────────────────────────────
// Juga mendukung nama-nama sheet alternatif yang umum dipakai
export const SHEET_MAP = {
  // Nama sheet utama (exact match)
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
  // Alias sheet alternatif yang umum
  'Karyawan':                    { module: 'employees',   label: 'Karyawan' },
  'Kontrak':                     { module: 'contracts',   label: 'Kontrak' },
  'Reliefer':                    { module: 'relievers',   label: 'Reliefer' },
  'Timeline':                    { module: 'schedule',    label: 'Jadwal Kegiatan' },
  'Jadwal Kegiatan':             { module: 'schedule',    label: 'Jadwal Kegiatan' },
  'Issue':                       { module: 'issues',      label: 'Permasalahan' },
  'Issues':                      { module: 'issues',      label: 'Permasalahan' },
  'Checklist':                   { module: 'checklist',   label: 'Master Checklist' },
  'Form':                        { module: 'forms',       label: 'Master Form' },
  'Inspeksi':                    { module: 'inspection',  label: 'Laporan Inspeksi' },
  'Laporan Inspeksi':            { module: 'inspection',  label: 'Laporan Inspeksi' },
  'GCDC':                        { module: 'cleaning',    label: 'Laporan GC/DC' },
  'GC DC':                       { module: 'cleaning',    label: 'Laporan GC/DC' },
  'Fogging':                     { module: 'fogging',     label: 'Laporan Fogging' },
  'Basecamp':                    { module: 'basecamp',    label: 'Rekap Basecamp' },
  'Chemical':                    { module: 'supply',      label: 'Inventory Chemical' },
  'Supply':                      { module: 'supply',      label: 'Inventory Chemical' },
};

// ─── Global Column Aliases ───────────────────────────────────────────────────
// Setiap field dapat dibaca dari BANYAK nama kolom yang berbeda di Excel.
// Ditambahkan semua variasi umum dari file Excel Facility Care.
const ALIASES = {
  // ── Nama Karyawan / FC ──
  full_name: [
    'Nama Lengkap', 'Nama Karyawan', 'Nama FC', 'Nama', 'Name',
    'Nama Lengkap Karyawan', 'Fullname', 'Full Name', 'NAMA',
    'Nama Lengkap FC', 'Nama Pegawai', 'Nama Pekerja',
  ],
  employee_name: [
    'Nama Karyawan', 'Nama Lengkap', 'Nama FC', 'Nama', 'Name',
    'Nama Pegawai', 'Karyawan', 'FC', 'Nama Pekerja', 'Full Name',
  ],
  reliever_name: [
    'Reliefer', 'Nama Reliefer', 'Pengganti', 'Nama Pengganti',
    'Back Up', 'Backup', 'FC Reliefer', 'Nama Back Up',
  ],
  original_fc_name: [
    'FC Digantikan', 'FC Yang Digantikan', 'Nama FC', 'FC Asli',
    'Yang Digantikan', 'FC Asal', 'Posisi', 'FC Posisi',
  ],

  // ── Cabang ──
  branch_name: [
    'Cabang', 'Nama Cabang', 'Branch', 'Klinik', 'Nama Klinik',
    'Lokasi', 'Cabang / Klinik', 'Cabang/Klinik', 'Site',
    'Cabang Klinik', 'CABANG', 'Unit', 'Nama Unit',
  ],

  // ── Divisi ──
  division: [
    'Divisi', 'Division', 'Bagian', 'Departemen', 'Department',
    'Jabatan Divisi', 'Posisi Divisi',
  ],

  // ── Tanggal (generic) ──
  tanggal: [
    'Tanggal', 'Tgl', 'Tgl.', 'Date', 'Tgl Laporan',
    'Tanggal Laporan', 'Tanggal Kejadian', 'Tanggal Masuk',
    'Tgl Masuk', 'Tanggal Input', 'Input Date',
  ],

  // ── Tanggal Mulai Kontrak ──
  start_date: [
    'Tgl Mulai', 'Tanggal Mulai', 'Mulai', 'Start Date', 'Start',
    'Tgl. Mulai', 'Tgl Awal', 'Tanggal Awal', 'Tgl Kontrak Mulai',
    'Tanggal Kontrak Mulai', 'Tgl Berlaku',
  ],

  // ── Tanggal Selesai / Akhir ──
  end_date: [
    'Tgl Selesai', 'Tanggal Selesai', 'Selesai', 'End Date', 'End',
    'Tgl. Selesai', 'Tgl Akhir', 'Tanggal Akhir', 'Tgl Kontrak Selesai',
    'Tanggal Berakhir', 'Tgl Berakhir', 'Expiry', 'Expired',
  ],

  // ── Tanggal Selesai (aktual/realisasi) ──
  completion_date: [
    'Tgl Selesai', 'Tanggal Selesai', 'Tgl Done', 'Done Date',
    'Tgl Realisasi', 'Tanggal Realisasi', 'Tgl Penyelesaian',
    'Tanggal Penyelesaian', 'Selesai', 'End Date',
  ],

  // ── Tanggal Backup Reliefer ──
  backup_date: [
    'Tanggal Backup', 'Tgl Backup', 'Tanggal Penggantian',
    'Tgl Penggantian', 'Backup Date', 'Tanggal Ganti',
    'Tgl', 'Tanggal', 'Date',
  ],

  // ── Periode (Quarter) ──
  period: [
    'Periode', 'Period', 'Quarter', 'Q', 'Kuartal',
    'Semester', 'Bulan', 'Month', 'Triwulan',
  ],

  // ── Tipe Kontrak ──
  contract_type: [
    'Tipe Kontrak', 'Jenis Kontrak', 'Type Kontrak', 'Kontrak',
    'Contract Type', 'Tipe', 'Jenis', 'Status Kontrak',
  ],

  // ── PKWT ──
  pkwt_number: [
    'PKWT', 'No PKWT', 'Nomor PKWT', 'PKWT Number',
    'No. PKWT', 'PKWT Ke', 'PKWT-',
  ],

  // ── Kegiatan / Aktivitas ──
  activity_type: [
    'Kegiatan', 'Jenis Kegiatan', 'Aktivitas', 'Activity',
    'Type', 'Jenis', 'Tipe Kegiatan', 'Activity Type',
    'Nama Kegiatan', 'Program',
  ],

  // ── PIC ──
  pic: [
    'PIC', 'Penanggung Jawab', 'In Charge', 'Person In Charge',
    'Pelaksana', 'Nama PIC', 'PJ', 'Koordinator',
  ],

  // ── Tanggal Opening ──
  opening_date: [
    'Tgl Opening', 'Tanggal Opening', 'Opening Date', 'Open Date',
    'Tgl Buka', 'Tanggal Buka', 'Tgl Pembukaan',
  ],

  // ── Tanggal Target ──
  target_date: [
    'Tgl Target', 'Tanggal Target', 'Target Date', 'Deadline',
    'Due Date', 'Tgl. Target', 'Target',
  ],

  // ── Status ──
  status: [
    'Status', 'STATUS', 'Kondisi', 'State', 'Keterangan Status',
  ],

  // ── Catatan / Keterangan ──
  notes: [
    'Catatan', 'Keterangan', 'Notes', 'Note', 'Ket', 'Ket.',
    'Komentar', 'Remark', 'Remarks', 'Deskripsi',
    'Catatan Tambahan', 'Info Tambahan',
  ],

  // ── Permasalahan / Keluhan ──
  problem: [
    'Permasalahan', 'Masalah', 'Problem', 'Keluhan', 'Complaint',
    'Issue', 'Kendala', 'Temuan', 'Catatan Masalah',
  ],
  complaint: [
    'Keluhan', 'Permasalahan', 'Masalah', 'Complaint', 'Problem',
    'Issue', 'Kendala', 'Temuan', 'Deskripsi Masalah',
  ],

  // ── Kategori ──
  category: [
    'Kategori', 'Category', 'Jenis', 'Type', 'Klasifikasi',
    'Tipe', 'Golongan',
  ],

  // ── Sumber (Issues) ──
  source: [
    'Sumber', 'Source', 'Sumber Laporan', 'Dilaporkan Oleh',
    'Pelapor', 'Reporter', 'Dari',
  ],

  // ── FC / Nama FC (Issues) ──
  fc_employee: [
    'Nama FC', 'FC', 'Nama Karyawan', 'Karyawan', 'Nama Pekerja',
    'Nama Pegawai', 'FC Bermasalah', 'Nama',
  ],

  // ── FC Specialist ──
  fc_specialist: [
    'FC Spesialis', 'FC Specialist', 'Spesialis', 'Spv FC',
    'SPV', 'Supervisor', 'FC PIC', 'Penangani',
  ],

  // ── Solusi ──
  solution: [
    'Solusi', 'Solution', 'Penyelesaian', 'Tindakan',
    'Action', 'Follow Up', 'Tindak Lanjut',
  ],

  // ── Tgl Masuk (join) ──
  join_date: [
    'Tgl Masuk', 'Tanggal Masuk', 'Join Date', 'Bergabung',
    'Tgl. Masuk', 'Mulai Kerja', 'Tgl Mulai Kerja',
    'Tanggal Bergabung', 'Tanggal Mulai Kerja',
  ],

  // ── No HP ──
  phone: [
    'No. HP', 'No HP', 'HP', 'Phone', 'Telepon', 'Telp',
    'Nomor HP', 'No. Telp', 'Handphone', 'Mobile',
    'No. Handphone', 'Nomor Telepon',
  ],

  // ── Training ──
  batch: [
    'Batch', 'Angkatan', 'Gelombang', 'Batch Training',
  ],
  subject: [
    'Materi', 'Subject', 'Topik', 'Topic', 'Judul Training',
    'Nama Training', 'Training', 'Materi Training',
  ],
  participants: [
    'Peserta', 'Participants', 'Nama Peserta', 'Participant',
    'Daftar Peserta', 'Anggota', 'Karyawan',
  ],
  trainer: [
    'Trainer', 'Pelatih', 'Instruktur', 'Fasilitator',
    'PIC', 'Pemateri', 'Narasumber',
  ],
  score: [
    'Nilai', 'Score', 'Skor', 'Nilai Akhir', 'Hasil',
    'Grade', 'Point', 'Poin',
  ],

  // ── Dokumen / Link ──
  document_link: [
    'Link Dokumen', 'Link', 'URL', 'Dokumen', 'File',
    'Link File', 'Google Drive', 'Drive Link', 'Attachment',
    'Lampiran', 'Document', 'Hyperlink',
  ],

  // ── SOP ──
  sop_name: [
    'Nama', 'Nama SOP', 'SOP', 'Judul', 'Judul SOP',
    'Title', 'Nama Dokumen',
  ],
  sop_category: [
    'Kategori', 'Category', 'Jenis SOP', 'Bidang', 'Tipe',
  ],
  version: [
    'Versi', 'Version', 'Ver', 'Ver.', 'Edisi',
  ],
  effective_date: [
    'Tgl Berlaku', 'Tanggal Berlaku', 'Effective Date', 'Berlaku Mulai',
    'Tgl Efektif', 'Mulai Berlaku',
  ],

  // ── Checklist ──
  checklist_name: [
    'Nama', 'Nama Checklist', 'Checklist', 'Item', 'Judul',
    'Nama Item', 'Title',
  ],

  // ── Form ──
  form_name: [
    'Nama', 'Nama Form', 'Form', 'Judul', 'Nama Formulir',
    'Formulir', 'Title',
  ],

  // ── Inspection ──
  inspection_date: [
    'Tanggal', 'Tgl', 'Tgl Inspeksi', 'Tanggal Inspeksi',
    'Tgl Kunjungan', 'Date', 'Tgl Pelaksanaan',
  ],
  fc_score: [
    'Nilai FC', 'FC Score', 'Score FC', 'Nilai', 'Nilai Hygiene',
    'FC', 'Skor FC', 'Penilaian FC',
  ],
  spv_score: [
    'Nilai SPV', 'SPV Score', 'Score SPV', 'Nilai SPV',
    'SPV', 'Skor SPV', 'Penilaian SPV',
  ],

  // ── Cleaning / GC DC ──
  activity_date: [
    'Tanggal', 'Tgl', 'Tgl Pelaksanaan', 'Tanggal Pelaksanaan',
    'Tgl Kegiatan', 'Date', 'Tgl Realisasi',
  ],

  // ── Basecamp ──
  info_date: [
    'Tanggal Info', 'Tgl Info', 'Tanggal Laporan', 'Tgl Laporan',
    'Tanggal', 'Tgl', 'Tgl Masuk', 'Date', 'Tgl Diterima',
  ],
  done_date: [
    'Tgl Selesai', 'Tanggal Selesai', 'Done Date', 'Tgl Done',
    'Tgl Penyelesaian', 'Selesai',
  ],

  // ── Supply / Chemical ──
  submitted_at: [
    'Tanggal', 'Tgl', 'Date', 'Tgl Permintaan', 'Tanggal Permintaan',
    'Tgl Pengajuan', 'Tanggal Pengajuan', 'Tgl Input',
  ],
  submitter_name: [
    'Nama Pemohon', 'Pemohon', 'Diajukan Oleh', 'Nama Pengaju',
    'Requester', 'Nama', 'Nama Karyawan', 'Pengguna',
    'Nama Pengguna', 'User', 'Yang Meminta',
  ],
  tools_items: [
    'Item Tools', 'Tools', 'Alat', 'Peralatan', 'Item Alat',
    'Nama Tools', 'Tools Yang Diminta',
  ],
  tools_quantity: [
    'Qty Tools', 'Jumlah Tools', 'Qty Alat', 'Jumlah Alat',
    'Jml Tools', 'Banyak Tools',
  ],
  chemical_items: [
    'Item Chemical', 'Chemical', 'Bahan', 'Bahan Kimia',
    'Item Bahan', 'Nama Chemical', 'Cairan',
  ],
  chemical_quantity: [
    'Qty Chemical', 'Jumlah Chemical', 'Qty Bahan', 'Jumlah Bahan',
    'Jml Chemical', 'Volume',
  ],
  additional_notes: [
    'Catatan', 'Keterangan', 'Notes', 'Tambahan', 'Catatan Tambahan',
    'Info', 'Informasi Tambahan',
  ],

  // ── Shift ──
  shift: [
    'Shift', 'Waktu', 'Jam Kerja', 'Jadwal',
  ],

  // ── Alasan / Keterangan ──
  reason: [
    'Keterangan', 'Alasan', 'Reason', 'Sebab', 'Why',
    'Keterangan Penggantian', 'Kenapa',
  ],

  // ── Tanggal Meeting One on One ──
  meeting_date: [
    'Tanggal', 'Tgl', 'Tgl Meeting', 'Tanggal Meeting',
    'Tgl Pertemuan', 'Date', 'Tgl One on One',
  ],
};

// ─── Core Helper: readField ───────────────────────────────────────────────────
/**
 * Membaca nilai dari row berdasarkan daftar alias (case-insensitive, trim).
 * Mengembalikan nilai pertama yang ditemukan, atau null.
 */
function readField(row, aliasKey) {
  const aliases = ALIASES[aliasKey] || [aliasKey];

  // Build lowercase key index dari row (sekali saja per call, cepat)
  for (const alias of aliases) {
    const aliasLower = alias.toLowerCase().trim();
    for (const [key, val] of Object.entries(row)) {
      if (key.toLowerCase().trim() === aliasLower) {
        const v = val !== undefined && val !== null && val !== '' ? String(val).trim() : null;
        if (v) return v;
      }
    }
  }
  return null;
}

/**
 * Mencari nilai dari row menggunakan daftar alias kustom (tidak dari ALIASES global).
 */
function readDirect(row, aliases) {
  for (const alias of aliases) {
    const al = alias.toLowerCase().trim();
    for (const [k, v] of Object.entries(row)) {
      if (k.toLowerCase().trim() === al) {
        const s = v !== undefined && v !== null && v !== '' ? String(v).trim() : null;
        if (s) return s;
      }
    }
  }
  return null;
}

// ─── Helpers ────────────────────────────────────────────────────────────────
function str(v) {
  if (v === undefined || v === null || v === '') return null;
  return String(v).trim() || null;
}

function dateStr(v) {
  if (v === undefined || v === null || v === '') return null;
  const s = String(v).trim();
  // ISO format already
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  // Excel serial number (numeric string 4-5 digits)
  if (/^\d{4,5}$/.test(s)) {
    const d = new Date(Date.UTC(1899, 11, 30) + Number(s) * 86400000);
    return d.toISOString().slice(0, 10);
  }
  // dd/mm/yyyy or dd-mm-yyyy or dd.mm.yyyy
  const parts = s.split(/[\/\-\.]/);
  if (parts.length === 3) {
    const [a, b, c] = parts;
    if (c.length === 4) {
      // Assume dd/mm/yyyy (Indonesian convention)
      return `${c}-${b.padStart(2, '0')}-${a.padStart(2, '0')}`;
    }
    if (a.length === 4) {
      // yyyy/mm/dd
      return `${a}-${b.padStart(2, '0')}-${c.padStart(2, '0')}`;
    }
  }
  // Try native Date parse as last resort
  const d = new Date(s);
  if (!isNaN(d)) return d.toISOString().slice(0, 10);
  return s;
}

function readDate(row, aliasKey) {
  return dateStr(readField(row, aliasKey));
}

function readNum(row, aliasKey) {
  const v = readField(row, aliasKey);
  if (!v) return null;
  const n = parseFloat(String(v).replace(',', '.'));
  return isNaN(n) ? null : n;
}

function isValidDate(s) {
  if (!s) return true; // optional field
  return /^\d{4}-\d{2}-\d{2}$/.test(s) && !isNaN(new Date(s));
}

// ─── Module Schemas ───────────────────────────────────────────────────────────
// required: array of aliasKey — harus ada setidaknya salah satu nilainya
// map: function(row) → mapped object
const SCHEMAS = {
  employees: {
    required: [{ key: 'full_name', label: 'Nama Karyawan' }],
    map: (row) => ({
      full_name:   readField(row, 'full_name'),
      branch_name: readField(row, 'branch_name'),
      division:    readField(row, 'division') || 'FACILITY CARE',
      phone:       readField(row, 'phone'),
      join_date:   readDate(row, 'join_date'),
      status:      readField(row, 'status') || 'Aktif',
      notes:       readField(row, 'notes'),
    }),
  },

  contracts: {
    required: [
      { key: 'employee_name', label: 'Nama Karyawan' },
      { key: 'start_date',    label: 'Tanggal Mulai' },
      { key: 'end_date',      label: 'Tanggal Selesai' },
    ],
    map: (row) => ({
      employee_name: readField(row, 'employee_name'),
      branch_name:   readField(row, 'branch_name'),
      division:      readField(row, 'division') || 'FACILITY CARE',
      start_date:    readDate(row, 'start_date'),
      end_date:      readDate(row, 'end_date'),
      contract_type: readField(row, 'contract_type'),
      pkwt_number:   readField(row, 'pkwt_number'),
      status:        readField(row, 'status') || 'Aktif',
      notes:         readField(row, 'notes'),
    }),
  },

  relievers: {
    required: [
      { key: 'reliever_name', label: 'Nama Reliefer' },
      { key: 'backup_date',   label: 'Tanggal Backup' },
    ],
    map: (row) => ({
      branch_name:      readField(row, 'branch_name'),
      original_fc_name: readField(row, 'original_fc_name'),
      period:           readField(row, 'period'),
      reliever_name:    readField(row, 'reliever_name'),
      backup_date:      readDate(row, 'backup_date'),
      completion_date:  readDate(row, 'completion_date'),
      reason:           readField(row, 'reason'),
      shift:            readField(row, 'shift'),
      status:           readField(row, 'status') || 'Pending',
    }),
  },

  schedule: {
    required: [
      { key: 'activity_type', label: 'Kegiatan' },
      { key: 'period',        label: 'Periode' },
    ],
    map: (row) => ({
      branch_name:     readField(row, 'branch_name'),
      activity_type:   readField(row, 'activity_type'),
      period:          readField(row, 'period'),
      pic:             readField(row, 'pic'),
      opening_date:    readDate(row, 'opening_date'),
      target_date:     readDate(row, 'target_date'),
      completion_date: readDate(row, 'completion_date'),
      status:          readField(row, 'status') || 'Pending',
      notes:           readField(row, 'notes'),
    }),
  },

  issues: {
    required: [
      { key: 'tanggal',   label: 'Tanggal' },
      { key: 'complaint', label: 'Keluhan / Permasalahan' },
      { key: 'category',  label: 'Kategori' },
    ],
    map: (row) => ({
      report_date:     readDate(row, 'tanggal'),
      branch_name:     readField(row, 'branch_name'),
      category:        readField(row, 'category'),
      source:          readField(row, 'source'),
      complaint:       readField(row, 'complaint'),
      employee_name:   readField(row, 'fc_employee'),
      fc_specialist:   readField(row, 'fc_specialist'),
      solution:        readField(row, 'solution'),
      completion_date: readDate(row, 'completion_date'),
      status:          readField(row, 'status') || 'Open',
    }),
  },

  one_on_one: {
    required: [
      { key: 'meeting_date',   label: 'Tanggal' },
      { key: 'employee_name',  label: 'Nama Karyawan' },
      { key: 'problem',        label: 'Permasalahan' },
    ],
    map: (row) => ({
      meeting_date:    readDate(row, 'meeting_date'),
      branch_name:     readField(row, 'branch_name'),
      employee_name:   readField(row, 'employee_name'),
      pic:             readField(row, 'pic'),
      problem:         readField(row, 'problem'),
      solution:        readField(row, 'solution'),
      completion_date: readDate(row, 'completion_date'),
      status:          readField(row, 'status') || 'Open',
    }),
  },

  training: {
    required: [
      { key: 'tanggal', label: 'Tanggal' },
      { key: 'subject', label: 'Materi' },
    ],
    map: (row) => ({
      training_date: readDate(row, 'tanggal'),
      batch:         readField(row, 'batch'),
      subject:       readField(row, 'subject'),
      participants:  readField(row, 'participants'),
      branch_name:   readField(row, 'branch_name'),
      trainer:       readField(row, 'trainer'),
      score:         readNum(row, 'score'),
      notes:         readField(row, 'notes'),
    }),
  },

  checklist: {
    required: [{ key: 'checklist_name', label: 'Nama Checklist' }],
    map: (row) => ({
      name:          readField(row, 'checklist_name'),
      category:      readField(row, 'category'),
      document_link: readField(row, 'document_link'),
      description:   readField(row, 'notes'),
    }),
  },

  forms: {
    required: [{ key: 'form_name', label: 'Nama Form' }],
    map: (row) => ({
      name:          readField(row, 'form_name'),
      category:      readField(row, 'category'),
      document_link: readField(row, 'document_link'),
      description:   readField(row, 'notes'),
    }),
  },

  sop: {
    required: [
      { key: 'sop_name',     label: 'Nama SOP' },
      { key: 'sop_category', label: 'Kategori' },
    ],
    map: (row) => ({
      name:           readField(row, 'sop_name'),
      category:       readField(row, 'sop_category'),
      document_link:  readField(row, 'document_link'),
      version:        readField(row, 'version'),
      effective_date: readDate(row, 'effective_date'),
      notes:          readField(row, 'notes'),
    }),
  },

  inspection: {
    required: [
      { key: 'inspection_date', label: 'Tanggal' },
      { key: 'period',          label: 'Periode' },
    ],
    map: (row) => ({
      inspection_date: readDate(row, 'inspection_date'),
      branch_name:     readField(row, 'branch_name'),
      period:          readField(row, 'period'),
      status:          readField(row, 'status') || 'Pending',
      fc_score:        readNum(row, 'fc_score'),
      spv_score:       readNum(row, 'spv_score'),
      document_link:   readField(row, 'document_link'),
      notes:           readField(row, 'notes'),
    }),
  },

  cleaning: {
    required: [
      { key: 'activity_date', label: 'Tanggal' },
      { key: 'period',        label: 'Periode' },
    ],
    map: (row) => ({
      activity_date: readDate(row, 'activity_date'),
      branch_name:   readField(row, 'branch_name'),
      activity_type: readField(row, 'activity_type') || 'General Cleaning',
      period:        readField(row, 'period'),
      status:        readField(row, 'status') || 'Pending',
      document_link: readField(row, 'document_link'),
      notes:         readField(row, 'notes'),
    }),
  },

  fogging: {
    required: [{ key: 'period', label: 'Periode' }],
    map: (row) => ({
      activity_date: readDate(row, 'activity_date'),
      branch_name:   readField(row, 'branch_name'),
      period:        readField(row, 'period'),
      status:        readField(row, 'status') || 'Pending',
      document_link: readField(row, 'document_link'),
      notes:         readField(row, 'notes'),
    }),
  },

  basecamp: {
    required: [
      { key: 'info_date', label: 'Tanggal Info' },
      { key: 'problem',   label: 'Permasalahan' },
    ],
    map: (row) => ({
      info_date:   readDate(row, 'info_date'),
      branch_name: readField(row, 'branch_name'),
      problem:     readField(row, 'problem'),
      pic:         readField(row, 'pic'),
      done_date:   readDate(row, 'done_date'),
      status:      readField(row, 'status') || 'Open',
      notes:       readField(row, 'notes'),
    }),
  },

  supply: {
    required: [
      { key: 'submitted_at',   label: 'Tanggal' },
      { key: 'submitter_name', label: 'Nama Pemohon' },
    ],
    map: (row) => ({
      submitted_at:      readDate(row, 'submitted_at'),
      submitter_name:    readField(row, 'submitter_name'),
      branch_name:       readField(row, 'branch_name'),
      tools_items:       readField(row, 'tools_items'),
      tools_quantity:    readField(row, 'tools_quantity'),
      chemical_items:    readField(row, 'chemical_items'),
      chemical_quantity: readField(row, 'chemical_quantity'),
      additional_notes:  readField(row, 'additional_notes'),
      status:            readField(row, 'status') || 'Pending',
    }),
  },
};

// ─── Sheet Name Fuzzy Matching ───────────────────────────────────────────────
/**
 * Cari modul yang cocok dengan nama sheet, termasuk fuzzy match partial.
 * Contoh: "Report Inspeksi Hygiene 2025" → cocok ke "Report Inspeksi Hygiene 2026"
 */
function findSheetMapping(sheetName) {
  const name = sheetName.trim();

  // Exact match dulu
  if (SHEET_MAP[name]) return SHEET_MAP[name];

  // Case-insensitive exact
  const nameLower = name.toLowerCase();
  for (const [key, val] of Object.entries(SHEET_MAP)) {
    if (key.toLowerCase() === nameLower) return val;
  }

  // Partial / contains match (tahun bisa berbeda)
  for (const [key, val] of Object.entries(SHEET_MAP)) {
    const kl = key.toLowerCase();
    // Remove year numbers for comparison
    const nameNoYear = nameLower.replace(/\d{4}/g, '').trim();
    const keyNoYear  = kl.replace(/\d{4}/g, '').trim();
    if (nameNoYear === keyNoYear) return val;
    if (nameNoYear.includes(keyNoYear) || keyNoYear.includes(nameNoYear)) return val;
  }

  // Keyword-based match
  if (nameLower.includes('karyawan') || nameLower.includes('pegawai')) return SHEET_MAP['Master Karyawan'];
  if (nameLower.includes('kontrak')) return SHEET_MAP['Data Kontrak'];
  if (nameLower.includes('reliefer') || nameLower.includes('reliever') || nameLower.includes('backup')) return SHEET_MAP['Jadwal Reliefer'];
  if (nameLower.includes('timeline') || nameLower.includes('time line') || nameLower.includes('jadwal kegiatan')) return SHEET_MAP['Time Line'];
  if (nameLower.includes('masalah') || nameLower.includes('permasalahan') || nameLower.includes('issue')) return SHEET_MAP['Permasalahan'];
  if (nameLower.includes('one on one') || nameLower.includes('1 on 1') || nameLower.includes('oon')) return SHEET_MAP['One on One'];
  if (nameLower.includes('training') || nameLower.includes('pelatihan')) return SHEET_MAP['Training'];
  if (nameLower.includes('checklist')) return SHEET_MAP['Master Checklist'];
  if (nameLower.includes('form') && !nameLower.includes('fogging')) return SHEET_MAP['Master Form'];
  if (nameLower.includes('sop')) return SHEET_MAP['SOP'];
  if (nameLower.includes('inspeksi') || nameLower.includes('hygiene')) return SHEET_MAP['Report Inspeksi Hygiene 2026'];
  if (nameLower.includes('gcdc') || nameLower.includes('gc dc') || nameLower.includes('gc/dc') || nameLower.includes('general cleaning') || nameLower.includes('deep clean')) return SHEET_MAP['Report GCDC 2026'];
  if (nameLower.includes('fogging')) return SHEET_MAP['Report Fogging 2026'];
  if (nameLower.includes('basecamp') || nameLower.includes('base camp')) return SHEET_MAP['Rekap Laporan Basecamp'];
  if (nameLower.includes('chemical') || nameLower.includes('supply') || nameLower.includes('permintaan')) return SHEET_MAP['Permintaan Chemical'];

  return null;
}

// ─── Main Validator ──────────────────────────────────────────────────────────
export function validateSheet(sheetName, rawRows, context = {}) {
  const mapping = findSheetMapping(sheetName);
  if (!mapping) return { valid: [], errors: [], mapped: [], skipped: true };

  const schema = SCHEMAS[mapping.module];
  if (!schema) return { valid: [], errors: [], mapped: [], skipped: true };

  const validRows   = [];
  const errorRows   = [];
  const mappedRows  = [];

  // Detect empty rows (all cells empty or just numbers/spaces)
  const nonEmptyRows = rawRows.filter(raw => {
    const vals = Object.values(raw);
    return vals.some(v => v !== undefined && v !== null && String(v).trim() !== '');
  });

  nonEmptyRows.forEach((raw, idx) => {
    const rowNum = rawRows.indexOf(raw) + 2; // +2 for header
    const errs   = [];

    // Validate required fields
    schema.required.forEach(({ key, label }) => {
      const val = readField(raw, key);
      if (!val) errs.push(`Kolom "${label}" wajib diisi (kolom tidak ditemukan atau kosong)`);
    });

    // Map the row
    const mapped = schema.map(raw);

    // Validate date formats
    Object.entries(mapped).forEach(([k, v]) => {
      if (k.endsWith('_date') && v && !isValidDate(v)) {
        errs.push(`Format tanggal tidak valid pada "${k}": "${v}" (gunakan YYYY-MM-DD atau DD/MM/YYYY)`);
      }
    });

    // Cross-sheet: Kontrak butuh nama karyawan yang valid
    if (mapping.module === 'contracts' && mapped.employee_name) {
      if (context.employeeNames && context.employeeNames.size > 0) {
        if (!context.employeeNames.has(mapped.employee_name.toLowerCase())) {
          errs.push(`Nama karyawan "${mapped.employee_name}" tidak ditemukan di sheet Master Karyawan`);
        }
      }
      // If no employee names context (importing contracts alone), skip this check
    }

    if (errs.length > 0) {
      errorRows.push({ row: rowNum, data: mapped, errors: errs, raw });
    } else {
      validRows.push(raw);
      mappedRows.push(mapped);
    }
  });

  return { valid: validRows, errors: errorRows, mapped: mappedRows };
}

export function validateWorkbook(workbook) {
  const results = [];

  // First pass: extract employee names for cross-ref
  const employeeNames = new Set();
  workbook.SheetNames.forEach(sn => {
    const m = findSheetMapping(sn);
    if (m && m.module === 'employees') {
      const ws   = workbook.Sheets[sn];
      const rows = window.XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
      rows.forEach(r => {
        const n = readField(r, 'full_name');
        if (n) employeeNames.add(n.toLowerCase());
      });
    }
  });

  const context   = { employeeNames };
  const processed = new Set();

  // Process in defined order then any remaining
  const orderedSheets = [
    ...Object.keys(SHEET_MAP).filter(s => workbook.SheetNames.includes(s)),
    ...workbook.SheetNames.filter(s => !Object.keys(SHEET_MAP).includes(s)),
  ];
  // Deduplicate
  const seen = new Set();
  const finalOrder = [];
  for (const s of [...workbook.SheetNames]) {
    if (!seen.has(s)) { seen.add(s); finalOrder.push(s); }
  }

  finalOrder.forEach(sheetName => {
    if (['Validasi', 'Dropdown', 'Ref', 'Helper', 'Config'].includes(sheetName)) return;
    if (sheetName.startsWith('_')) return;

    const mapping = findSheetMapping(sheetName);
    if (!mapping) {
      results.push({ sheetName, module: null, label: '(Tidak dikenali — dilewati)', total: 0, valid: 0, errorCount: 0, errors: [], mapped: [], skipped: true });
      return;
    }
    if (processed.has(mapping.module)) return; // skip duplicate sheet pointing to same module
    processed.add(mapping.module);

    const ws      = workbook.Sheets[sheetName];
    const rawRows = window.XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });

    const result  = validateSheet(sheetName, rawRows, context);
    results.push({
      sheetName,
      module:     mapping.module,
      label:      mapping.label,
      total:      rawRows.filter(r => Object.values(r).some(v => String(v).trim())).length,
      valid:      result.mapped.length,
      errorCount: result.errors.length,
      errors:     result.errors,
      mapped:     result.mapped,
      skipped:    false,
    });
  });

  return results;
}

// ─── Template Generator ──────────────────────────────────────────────────────
export function generateTemplate() {
  const XLSX = window.XLSX;
  const wb   = XLSX.utils.book_new();

  const TEMPLATES = {
    'Master Karyawan': [{
      'Nama Lengkap': 'Budi Santoso', 'Cabang': '001. Pondok Bambu',
      'Divisi': 'FACILITY CARE', 'No. HP': '081234567890',
      'Tgl Masuk': '2024-01-15', 'Status': 'Aktif', 'Catatan': '',
    }],
    'Data Kontrak': [{
      'Nama Karyawan': 'Budi Santoso', 'Cabang': '001. Pondok Bambu',
      'Divisi': 'FACILITY CARE', 'Tgl Mulai': '2024-01-01',
      'Tgl Selesai': '2024-12-31', 'Tipe Kontrak': 'KONTRAK 1 TAHUN',
      'PKWT': 'PKWT 1', 'Status': 'Aktif', 'Catatan': '',
    }],
    'Jadwal Reliefer': [{
      'Tanggal Backup': '2024-03-10', 'Cabang': '001. Pondok Bambu',
      'FC Digantikan': 'Budi Santoso', 'Periode': 'Q1', 'Reliefer': 'Andi',
      'Keterangan': 'Sakit', 'Shift': 'Pagi', 'Tanggal Selesai': '2024-03-10', 'Status': 'Done',
    }],
    'Time Line': [{
      'Cabang': '001. Pondok Bambu', 'Kegiatan': 'General Cleaning', 'Periode': 'Q1',
      'PIC': 'Fajar', 'Tgl Opening': '2024-02-01', 'Tgl Target': '2024-02-15',
      'Tgl Selesai': '2024-02-14', 'Status': 'Done', 'Catatan': '',
    }],
    'Permasalahan': [{
      'Tanggal': '2024-03-01', 'Cabang': '001. Pondok Bambu', 'Kategori': 'Cleaning',
      'Sumber': 'SPV', 'Keluhan': 'Lantai kotor', 'Nama FC': 'Budi',
      'FC Spesialis': 'Fajar', 'Solusi': 'Teguran', 'Tgl Selesai': '2024-03-02', 'Status': 'Done',
    }],
    'One on One': [{
      'Tanggal': '2024-03-05', 'Cabang': '001. Pondok Bambu', 'Nama Karyawan': 'Budi Santoso',
      'PIC': 'Berlin', 'Permasalahan': 'Keterlambatan', 'Solusi': 'Coaching',
      'Tgl Selesai': '2024-03-06', 'Status': 'Done',
    }],
    'Training': [{
      'Tanggal': '2024-02-20', 'Batch': 'Batch 1', 'Materi': 'Hygiene & Sanitasi',
      'Peserta': 'Budi, Andi', 'Cabang': '001. Pondok Bambu',
      'Trainer': 'Fajar', 'Nilai': '85', 'Catatan': '',
    }],
    'Master Checklist': [{
      'Nama': 'Checklist Harian Toilet', 'Kategori': 'Harian',
      'Link Dokumen': 'https://docs.google.com/...', 'Deskripsi': '',
    }],
    'Master Form': [{
      'Nama': 'Form Permintaan Chemical', 'Kategori': 'Chemical',
      'Link Dokumen': 'https://docs.google.com/...', 'Deskripsi': '',
    }],
    'SOP': [{
      'Nama': 'SOP Pembersihan Toilet', 'Kategori': 'Cleaning',
      'Link Dokumen': 'https://docs.google.com/...', 'Versi': '1.0',
      'Tgl Berlaku': '2024-01-01', 'Catatan': '',
    }],
    'Report Inspeksi Hygiene 2026': [{
      'Tanggal': '2024-01-15', 'Cabang': '001. Pondok Bambu', 'Periode': 'Q1',
      'Status': 'Done', 'Nilai FC': '85', 'Nilai SPV': '90', 'Link Dokumen': '', 'Catatan': '',
    }],
    'Report GCDC 2026': [{
      'Tanggal': '2024-02-10', 'Cabang': '001. Pondok Bambu',
      'Jenis Kegiatan': 'General Cleaning', 'Periode': 'Q1',
      'Status': 'Done', 'Link Dokumen': '', 'Catatan': '',
    }],
    'Report Fogging 2026': [{
      'Tanggal': '2024-03-20', 'Cabang': '001. Pondok Bambu',
      'Periode': 'Q1', 'Status': 'Done', 'Link Dokumen': '', 'Catatan': '',
    }],
    'Rekap Laporan Basecamp': [{
      'Tanggal Info': '2024-01-10', 'Cabang': '001. Pondok Bambu',
      'Permasalahan': 'AC Rusak', 'PIC': 'Berlin',
      'Tgl Selesai': '2024-01-15', 'Status': 'Done', 'Catatan': '',
    }],
    'Permintaan Chemical': [{
      'Tanggal': '2024-02-01', 'Nama Pemohon': 'Budi Santoso', 'Cabang': '001. Pondok Bambu',
      'Item Tools': 'Sapu', 'Qty Tools': '2', 'Item Chemical': 'Sabun Lantai',
      'Qty Chemical': '5 liter', 'Catatan': '', 'Status': 'Pending',
    }],
  };

  Object.entries(TEMPLATES).forEach(([sheetName, data]) => {
    const ws = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  });

  XLSX.writeFile(wb, 'Template_Import_Data_Awal_FCMS.xlsx');
}

// ─── Error Log Generator ─────────────────────────────────────────────────────
export function downloadErrorLog(results) {
  const XLSX = window.XLSX;
  const wb   = XLSX.utils.book_new();
  let hasErrors = false;

  results.forEach(r => {
    if (!r.errors || r.errors.length === 0) return;
    hasErrors = true;
    const rows = r.errors.map(e => ({
      'No. Baris':      e.row,
      'Penyebab Error': e.errors.join('; '),
      ...Object.fromEntries(Object.entries(e.data || {}).map(([k, v]) => [k, v ?? ''])),
    }));
    const ws        = XLSX.utils.json_to_sheet(rows);
    const safeSheet = r.sheetName.replace(/[\\\/\[\]*?:]/g, '_').slice(0, 31);
    XLSX.utils.book_append_sheet(wb, ws, safeSheet);
  });

  if (!hasErrors) return false;
  XLSX.writeFile(wb, `Log_Error_Import_${new Date().toISOString().slice(0, 10)}.xlsx`);
  return true;
}
