/**
 * importer.js  v3 — Ultra-Permissive Validation Engine
 *
 * Filosofi v3:
 * - JANGAN blokir data karena kolom opsional kosong
 * - HANYA 1 field wajib per modul (identifier utama)
 * - Semua enum di-normalize otomatis (case-insensitive)
 * - Semua format tanggal diterima
 * - Cross-sheet validation DINONAKTIFKAN (tidak blokir kontrak)
 * - Error message actionable: nama kolom + nilai asli + cara fix
 */

// ─── Sheet → Module Mapping ──────────────────────────────────────────────────
export const SHEET_MAP = {
  'Master Karyawan':              { module: 'employees',  label: 'Karyawan' },
  'Data Kontrak':                 { module: 'contracts',  label: 'Kontrak' },
  'Jadwal Reliefer':              { module: 'relievers',  label: 'Reliefer' },
  'Time Line':                    { module: 'schedule',   label: 'Jadwal Kegiatan' },
  'Permasalahan':                 { module: 'issues',     label: 'Permasalahan' },
  'One on One':                   { module: 'one_on_one', label: 'One on One' },
  'Training':                     { module: 'training',   label: 'Training' },
  'Master Checklist':             { module: 'checklist',  label: 'Checklist' },
  'Master Form':                  { module: 'forms',      label: 'Master Form' },
  'SOP':                          { module: 'sop',        label: 'SOP' },
  'Report Inspeksi Hygiene 2026': { module: 'inspection', label: 'Laporan Inspeksi' },
  'Report GCDC 2026':             { module: 'cleaning',   label: 'Laporan GC/DC' },
  'Report Fogging 2026':          { module: 'fogging',    label: 'Laporan Fogging' },
  'Rekap Laporan Basecamp':       { module: 'basecamp',   label: 'Rekap Basecamp' },
  'Permintaan Chemical':          { module: 'supply',     label: 'Inventory Chemical' },
  // Alias sheet
  'Karyawan':          { module: 'employees',  label: 'Karyawan' },
  'Kontrak':           { module: 'contracts',  label: 'Kontrak' },
  'Reliefer':          { module: 'relievers',  label: 'Reliefer' },
  'Timeline':          { module: 'schedule',   label: 'Jadwal Kegiatan' },
  'Jadwal Kegiatan':   { module: 'schedule',   label: 'Jadwal Kegiatan' },
  'Issue':             { module: 'issues',     label: 'Permasalahan' },
  'Issues':            { module: 'issues',     label: 'Permasalahan' },
  'Checklist':         { module: 'checklist',  label: 'Checklist' },
  'Form':              { module: 'forms',      label: 'Master Form' },
  'Inspeksi':          { module: 'inspection', label: 'Laporan Inspeksi' },
  'Laporan Inspeksi':  { module: 'inspection', label: 'Laporan Inspeksi' },
  'GCDC':              { module: 'cleaning',   label: 'Laporan GC/DC' },
  'GC DC':             { module: 'cleaning',   label: 'Laporan GC/DC' },
  'Fogging':           { module: 'fogging',    label: 'Laporan Fogging' },
  'Basecamp':          { module: 'basecamp',   label: 'Rekap Basecamp' },
  'Chemical':          { module: 'supply',     label: 'Inventory Chemical' },
  'Supply':            { module: 'supply',     label: 'Inventory Chemical' },
  'One On One':        { module: 'one_on_one', label: 'One on One' },
  '1 on 1':            { module: 'one_on_one', label: 'One on One' },
};

// ─── Enum Normalizers ────────────────────────────────────────────────────────
const ENUM_STATUS = {
  'aktif': 'Aktif', 'active': 'Aktif', 'act': 'Aktif',
  'tidak aktif': 'Tidak Aktif', 'nonaktif': 'Tidak Aktif', 'inactive': 'Tidak Aktif',
  'resign': 'Resign', 'keluar': 'Resign',
  'done': 'Done', 'selesai': 'Done', 'completed': 'Done', 'complete': 'Done',
  'finish': 'Done', 'finished': 'Done', 'closed': 'Done', 'tutup': 'Done',
  'in progress': 'In Progress', 'proses': 'In Progress', 'sedang': 'In Progress',
  'on going': 'In Progress', 'ongoing': 'In Progress', 'berjalan': 'In Progress',
  'pending': 'Pending', 'menunggu': 'Pending', 'belum': 'Pending',
  'open': 'Open', 'buka': 'Open', 'terbuka': 'Open',
  'cancel': 'Batal', 'batal': 'Batal', 'cancelled': 'Batal',
};

const ENUM_SHIFT = {
  'pagi': 'Pagi', 'morning': 'Pagi', 'am': 'Pagi',
  'siang': 'Siang', 'afternoon': 'Siang',
  'sore': 'Sore', 'evening': 'Sore',
  'malam': 'Malam', 'night': 'Malam', 'pm': 'Malam',
  'full': 'Full Day', 'fullday': 'Full Day', 'full day': 'Full Day',
  'split': 'Split Shift',
};

function normalizeEnum(v, enumMap, defaultVal = null) {
  if (!v) return defaultVal;
  const key = String(v).trim().toLowerCase();
  return enumMap[key] || (defaultVal !== null ? defaultVal : String(v).trim());
}

// ─── Column Aliases ──────────────────────────────────────────────────────────
const ALIASES = {
  // Nama Karyawan / FC
  full_name: [
    'Nama Lengkap','Nama Karyawan','Nama FC','Nama','Name','NAMA',
    'Fullname','Full Name','Nama Lengkap Karyawan','Nama Pegawai',
    'Nama Pekerja','Employee Name','Nama Personil',
  ],
  employee_name: [
    'Nama Karyawan','Nama Lengkap','Nama FC','Nama','Name','NAMA',
    'Karyawan','FC','Nama Pegawai','Employee Name','Nama Personil',
    'Nama Pekerja','Nama Lengkap Karyawan',
  ],
  reliever_name: [
    'Reliefer','Nama Reliefer','Pengganti','Nama Pengganti',
    'Back Up','Backup','FC Reliefer','Nama Back Up','Nama Backup',
    'Nama','Karyawan Pengganti',
  ],
  original_fc_name: [
    'FC Digantikan','FC Yang Digantikan','Yang Digantikan','FC Asli',
    'FC Asal','Posisi','FC Posisi','Nama FC Asal','Nama FC Digantikan',
  ],
  // Cabang
  branch_name: [
    'Cabang','Nama Cabang','Branch','Klinik','Nama Klinik','Lokasi',
    'Cabang / Klinik','Cabang/Klinik','Site','Unit','Nama Unit',
    'CABANG','Cabang Klinik','Nama Lokasi',
  ],
  // Divisi
  division: [
    'Divisi','Division','Bagian','Departemen','Department',
    'Jabatan Divisi','Posisi Divisi','Dept',
  ],
  // No HP
  phone: [
    'No. HP','No HP','HP','Phone','Telepon','Telp','Nomor HP',
    'No. Telp','Handphone','Mobile','No. Handphone','Nomor Telepon',
    'Contact','Kontak','No Telp',
  ],
  // Tgl Masuk
  join_date: [
    'Tgl Masuk','Tanggal Masuk','Join Date','Bergabung','Tgl. Masuk',
    'Mulai Kerja','Tgl Mulai Kerja','Tanggal Bergabung',
    'Tanggal Mulai Kerja','Tgl Gabung','Masuk Kerja',
  ],
  // Status
  status: [
    'Status','STATUS','Kondisi','State','Keterangan Status','Sts',
  ],
  // Catatan
  notes: [
    'Catatan','Keterangan','Notes','Note','Ket','Ket.','Komentar',
    'Remark','Remarks','Deskripsi','Catatan Tambahan','Info Tambahan',
    'Description','Desc','Informasi',
  ],
  // Tanggal generic
  tanggal: [
    'Tanggal','Tgl','Tgl.','Date','Tgl Laporan','Tanggal Laporan',
    'Tanggal Kejadian','Tgl Masuk','Tanggal Input','Input Date',
    'Tgl Kejadian','Tgl Terima','Tanggal Diterima',
  ],
  // Tgl Mulai
  start_date: [
    'Tgl Mulai','Tanggal Mulai','Mulai','Start Date','Start',
    'Tgl. Mulai','Tgl Awal','Tanggal Awal','Tgl Kontrak Mulai',
    'Tanggal Kontrak Mulai','Tgl Berlaku','Tgl Efektif',
  ],
  // Tgl Selesai
  end_date: [
    'Tgl Selesai','Tanggal Selesai','Selesai','End Date','End',
    'Tgl. Selesai','Tgl Akhir','Tanggal Akhir','Tgl Kontrak Selesai',
    'Tanggal Berakhir','Tgl Berakhir','Expiry','Expired','Tgl Habis',
  ],
  // Tgl selesai aktual
  completion_date: [
    'Tgl Selesai','Tanggal Selesai','Tgl Done','Done Date',
    'Tgl Realisasi','Tanggal Realisasi','Tgl Penyelesaian',
    'Tanggal Penyelesaian','Selesai','End Date','Tgl Tutup',
  ],
  // Backup date
  backup_date: [
    'Tanggal Backup','Tgl Backup','Tanggal Penggantian',
    'Tgl Penggantian','Backup Date','Tanggal Ganti','Tgl Ganti',
    'Tgl','Tanggal','Date','Tgl Mulai Backup',
  ],
  // Periode
  period: [
    'Periode','Period','Quarter','Q','Kuartal','Semester',
    'Bulan','Month','Triwulan','Qtr',
  ],
  // Tipe Kontrak
  contract_type: [
    'Tipe Kontrak','Jenis Kontrak','Type Kontrak','Kontrak',
    'Contract Type','Tipe','Jenis','Status Kontrak','Jenis Perjanjian',
  ],
  // PKWT
  pkwt_number: [
    'PKWT','No PKWT','Nomor PKWT','PKWT Number','No. PKWT',
    'PKWT Ke','PKWT-','Nomor Kontrak','No Kontrak',
  ],
  // Kegiatan
  activity_type: [
    'Kegiatan','Jenis Kegiatan','Aktivitas','Activity','Type',
    'Jenis','Tipe Kegiatan','Activity Type','Nama Kegiatan','Program',
    'Agenda',
  ],
  // PIC
  pic: [
    'PIC','Penanggung Jawab','In Charge','Person In Charge',
    'Pelaksana','Nama PIC','PJ','Koordinator','Pengawas','SPV',
  ],
  // Tanggal Opening
  opening_date: [
    'Tgl Opening','Tanggal Opening','Opening Date','Open Date',
    'Tgl Buka','Tanggal Buka','Tgl Pembukaan','Tgl Mulai',
  ],
  // Target date
  target_date: [
    'Tgl Target','Tanggal Target','Target Date','Deadline',
    'Due Date','Tgl. Target','Target','Tgl Deadline',
  ],
  // Keluhan
  complaint: [
    'Keluhan','Permasalahan','Masalah','Complaint','Problem',
    'Issue','Kendala','Temuan','Deskripsi Masalah','Uraian Masalah',
    'Isi Laporan','Laporan','Kronologi',
  ],
  // Kategori
  category: [
    'Kategori','Category','Jenis','Type','Klasifikasi',
    'Tipe','Golongan','Jenis Masalah','Tipe Masalah',
  ],
  // Sumber
  source: [
    'Sumber','Source','Sumber Laporan','Dilaporkan Oleh',
    'Pelapor','Reporter','Dari','Oleh',
  ],
  // Nama FC (Issues)
  fc_employee: [
    'Nama FC','FC','Nama Karyawan','Karyawan','Nama Pekerja',
    'Nama Pegawai','FC Bermasalah','Nama','FC Terkait',
  ],
  // FC Specialist
  fc_specialist: [
    'FC Spesialis','FC Specialist','Spesialis','Spv FC',
    'SPV','Supervisor','FC PIC','Penangani','FC Handler',
  ],
  // Solusi
  solution: [
    'Solusi','Solution','Penyelesaian','Tindakan','Action',
    'Follow Up','Tindak Lanjut','Penanganan','Cara Penyelesaian',
  ],
  // Problem (One on One)
  problem: [
    'Permasalahan','Masalah','Problem','Keluhan','Complaint',
    'Issue','Kendala','Topik','Isi','Materi Diskusi',
  ],
  // Meeting date
  meeting_date: [
    'Tanggal','Tgl','Tgl Meeting','Tanggal Meeting','Tgl Pertemuan',
    'Date','Tgl One on One','Tgl 1on1',
  ],
  // Training
  batch: ['Batch','Angkatan','Gelombang','Batch Training','No Batch'],
  subject: [
    'Materi','Subject','Topik','Topic','Judul Training',
    'Nama Training','Training','Materi Training','Judul Materi',
  ],
  participants: [
    'Peserta','Participants','Nama Peserta','Participant',
    'Daftar Peserta','Anggota','Karyawan','Nama Peserta Training',
  ],
  trainer: [
    'Trainer','Pelatih','Instruktur','Fasilitator',
    'PIC','Pemateri','Narasumber','Pengajar',
  ],
  score: [
    'Nilai','Score','Skor','Nilai Akhir','Hasil','Grade','Point','Poin',
  ],
  // Dokumen
  document_link: [
    'Link Dokumen','Link','URL','Dokumen','File','Link File',
    'Google Drive','Drive Link','Attachment','Lampiran',
    'Document','Hyperlink','Tautan','Tautan Dokumen',
  ],
  // SOP / Checklist / Form name
  sop_name: [
    'Nama','Nama SOP','SOP','Judul','Judul SOP','Title',
    'Nama Dokumen','Nama Prosedur',
  ],
  sop_category: [
    'Kategori','Category','Jenis SOP','Bidang','Tipe','Jenis',
  ],
  version: ['Versi','Version','Ver','Ver.','Edisi','Rev'],
  effective_date: [
    'Tgl Berlaku','Tanggal Berlaku','Effective Date','Berlaku Mulai',
    'Tgl Efektif','Mulai Berlaku','Tgl Mulai Berlaku',
  ],
  checklist_name: [
    'Nama','Nama Checklist','Checklist','Item','Judul',
    'Nama Item','Title','Nama Aktivitas',
  ],
  form_name: [
    'Nama','Nama Form','Form','Judul','Nama Formulir',
    'Formulir','Title','Nama Dokumen',
  ],
  // Inspection
  inspection_date: [
    'Tanggal','Tgl','Tgl Inspeksi','Tanggal Inspeksi',
    'Tgl Kunjungan','Date','Tgl Pelaksanaan','Tgl Audit',
  ],
  fc_score: [
    'Nilai FC','FC Score','Score FC','Nilai','Nilai Hygiene',
    'FC','Skor FC','Penilaian FC','Nilai Audit FC',
  ],
  spv_score: [
    'Nilai SPV','SPV Score','Score SPV','Nilai SPV',
    'SPV','Skor SPV','Penilaian SPV','Nilai Audit SPV',
  ],
  // Cleaning
  activity_date: [
    'Tanggal','Tgl','Tgl Pelaksanaan','Tanggal Pelaksanaan',
    'Tgl Kegiatan','Date','Tgl Realisasi','Tgl Pelaksana',
  ],
  // Basecamp
  info_date: [
    'Tanggal Info','Tgl Info','Tanggal Laporan','Tgl Laporan',
    'Tanggal','Tgl','Tgl Masuk','Date','Tgl Diterima',
    'Tanggal Kejadian',
  ],
  done_date: [
    'Tgl Selesai','Tanggal Selesai','Done Date','Tgl Done',
    'Tgl Penyelesaian','Selesai','Completion Date',
  ],
  // Supply
  submitted_at: [
    'Tanggal','Tgl','Date','Tgl Permintaan','Tanggal Permintaan',
    'Tgl Pengajuan','Tanggal Pengajuan','Tgl Input','Tgl Order',
  ],
  submitter_name: [
    'Nama Pemohon','Pemohon','Diajukan Oleh','Nama Pengaju',
    'Requester','Nama','Nama Karyawan','Pengguna','Yang Meminta',
    'Nama Pengguna','User','Nama Pemesan','Nama Requestor',
  ],
  tools_items: [
    'Item Tools','Tools','Alat','Peralatan','Item Alat',
    'Nama Tools','Tools Yang Diminta','Item Peralatan',
  ],
  tools_quantity: [
    'Qty Tools','Jumlah Tools','Qty Alat','Jumlah Alat',
    'Jml Tools','Banyak Tools','Qty Peralatan',
  ],
  chemical_items: [
    'Item Chemical','Chemical','Bahan','Bahan Kimia','Item Bahan',
    'Nama Chemical','Cairan','Nama Bahan','Item Cairan',
  ],
  chemical_quantity: [
    'Qty Chemical','Jumlah Chemical','Qty Bahan','Jumlah Bahan',
    'Jml Chemical','Volume','Qty Cairan',
  ],
  additional_notes: [
    'Catatan','Keterangan','Notes','Tambahan','Catatan Tambahan',
    'Info','Informasi Tambahan','Remarks',
  ],
  shift: ['Shift','Waktu','Jam Kerja','Jadwal','Sif'],
  reason: [
    'Keterangan','Alasan','Reason','Sebab','Why',
    'Keterangan Penggantian','Kenapa','Ket',
  ],
};

// ─── Core Helpers ────────────────────────────────────────────────────────────

/** Baca field dari row dengan multi-alias, case-insensitive, trim. */
function readField(row, aliasKey) {
  const aliases = ALIASES[aliasKey] || [aliasKey];
  for (const alias of aliases) {
    const al = alias.toLowerCase().trim();
    for (const [key, val] of Object.entries(row)) {
      if (key.toLowerCase().trim() === al) {
        const v = (val !== undefined && val !== null && String(val).trim() !== '')
          ? String(val).trim() : null;
        if (v) return v;
      }
    }
  }
  return null;
}

/** Parse tanggal dari berbagai format. */
function dateStr(v) {
  if (v === undefined || v === null || v === '') return null;
  const s = String(v).trim();

  // ISO: 2024-03-15 atau dengan waktu
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);

  // Excel serial (integer 5 digit)
  if (/^\d{4,5}$/.test(s)) {
    const d = new Date(Date.UTC(1899, 11, 30) + Number(s) * 86400000);
    return d.toISOString().slice(0, 10);
  }

  const parts = s.split(/[\/\-\.]/);
  if (parts.length === 3) {
    const [a, b, c] = parts.map(p => p.trim());
    const na = Number(a), nb = Number(b), nc = Number(c);

    // yyyy/mm/dd (year first 4-digit)
    if (a.length === 4 && na > 1900) {
      return `${a}-${b.padStart(2,'0')}-${c.padStart(2,'0')}`;
    }
    // dd/mm/yyyy Indonesian (year last 4-digit)
    if (c.length === 4 && nc > 1900) {
      return `${c}-${b.padStart(2,'0')}-${a.padStart(2,'0')}`;
    }
    // dd/mm/yy (2-digit year)
    if (c.length === 2 && !isNaN(nc)) {
      const year = nc >= 50 ? `19${c}` : `20${c}`;
      return `${year}-${b.padStart(2,'0')}-${a.padStart(2,'0')}`;
    }
  }

  // Natural / locale string
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);

  return null; // return null instead of raw string — don't trigger date error
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

function readStatus(row, defaultVal = 'Pending') {
  const v = readField(row, 'status');
  return normalizeEnum(v, ENUM_STATUS, defaultVal);
}

function readShift(row) {
  const v = readField(row, 'shift');
  return normalizeEnum(v, ENUM_SHIFT, v);
}

/** Deteksi apakah baris kosong (semua kolom kosong/whitespace). */
function isEmptyRow(row) {
  return Object.values(row).every(v =>
    v === undefined || v === null || String(v).trim() === ''
  );
}

// ─── Module Schemas (v3 — sangat permisif) ───────────────────────────────────
// required: hanya 1 identifier mutlak per modul
// Semua field lain OPSIONAL
const SCHEMAS = {
  employees: {
    required: [{ key: 'full_name', label: 'Nama Karyawan', aliases: ALIASES.full_name }],
    map: (row) => ({
      full_name:   readField(row, 'full_name'),
      branch_name: readField(row, 'branch_name'),
      division:    readField(row, 'division') || 'FACILITY CARE',
      phone:       readField(row, 'phone'),
      join_date:   readDate(row, 'join_date'),
      status:      readStatus(row, 'Aktif'),
      notes:       readField(row, 'notes'),
    }),
  },
  contracts: {
    required: [{ key: 'employee_name', label: 'Nama Karyawan', aliases: ALIASES.employee_name }],
    map: (row) => ({
      employee_name: readField(row, 'employee_name'),
      branch_name:   readField(row, 'branch_name'),
      division:      readField(row, 'division') || 'FACILITY CARE',
      start_date:    readDate(row, 'start_date'),
      end_date:      readDate(row, 'end_date'),
      contract_type: readField(row, 'contract_type'),
      pkwt_number:   readField(row, 'pkwt_number'),
      status:        readStatus(row, 'Aktif'),
      notes:         readField(row, 'notes'),
    }),
  },
  relievers: {
    required: [{ key: 'reliever_name', label: 'Nama Reliefer', aliases: ALIASES.reliever_name }],
    map: (row) => ({
      branch_name:      readField(row, 'branch_name'),
      original_fc_name: readField(row, 'original_fc_name'),
      period:           readField(row, 'period'),
      reliever_name:    readField(row, 'reliever_name'),
      backup_date:      readDate(row, 'backup_date'),
      completion_date:  readDate(row, 'completion_date'),
      reason:           readField(row, 'reason'),
      shift:            readShift(row),
      status:           readStatus(row, 'Pending'),
    }),
  },
  schedule: {
    required: [{ key: 'activity_type', label: 'Kegiatan / Aktivitas', aliases: ALIASES.activity_type }],
    map: (row) => ({
      branch_name:     readField(row, 'branch_name'),
      activity_type:   readField(row, 'activity_type'),
      period:          readField(row, 'period'),
      pic:             readField(row, 'pic'),
      opening_date:    readDate(row, 'opening_date'),
      target_date:     readDate(row, 'target_date'),
      completion_date: readDate(row, 'completion_date'),
      status:          readStatus(row, 'Pending'),
      notes:           readField(row, 'notes'),
    }),
  },
  issues: {
    required: [{ key: 'complaint', label: 'Keluhan / Permasalahan', aliases: ALIASES.complaint }],
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
      status:          readStatus(row, 'Open'),
    }),
  },
  one_on_one: {
    required: [{ key: 'employee_name', label: 'Nama Karyawan', aliases: ALIASES.employee_name }],
    map: (row) => ({
      meeting_date:    readDate(row, 'meeting_date'),
      branch_name:     readField(row, 'branch_name'),
      employee_name:   readField(row, 'employee_name'),
      pic:             readField(row, 'pic'),
      problem:         readField(row, 'problem'),
      solution:        readField(row, 'solution'),
      completion_date: readDate(row, 'completion_date'),
      status:          readStatus(row, 'Open'),
    }),
  },
  training: {
    required: [{ key: 'subject', label: 'Materi Training', aliases: ALIASES.subject }],
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
    required: [{ key: 'checklist_name', label: 'Nama Checklist', aliases: ALIASES.checklist_name }],
    map: (row) => ({
      name:          readField(row, 'checklist_name'),
      category:      readField(row, 'category'),
      document_link: readField(row, 'document_link'),
      description:   readField(row, 'notes'),
    }),
  },
  forms: {
    required: [{ key: 'form_name', label: 'Nama Form', aliases: ALIASES.form_name }],
    map: (row) => ({
      name:          readField(row, 'form_name'),
      category:      readField(row, 'category'),
      document_link: readField(row, 'document_link'),
      description:   readField(row, 'notes'),
    }),
  },
  sop: {
    required: [{ key: 'sop_name', label: 'Nama SOP', aliases: ALIASES.sop_name }],
    map: (row) => ({
      name:           readField(row, 'sop_name'),
      category:       readField(row, 'sop_category') || readField(row, 'category'),
      document_link:  readField(row, 'document_link'),
      version:        readField(row, 'version'),
      effective_date: readDate(row, 'effective_date'),
      notes:          readField(row, 'notes'),
    }),
  },
  inspection: {
    required: [{ key: 'branch_name', label: 'Cabang', aliases: ALIASES.branch_name }],
    map: (row) => ({
      inspection_date: readDate(row, 'inspection_date'),
      branch_name:     readField(row, 'branch_name'),
      period:          readField(row, 'period'),
      status:          readStatus(row, 'Pending'),
      fc_score:        readNum(row, 'fc_score'),
      spv_score:       readNum(row, 'spv_score'),
      document_link:   readField(row, 'document_link'),
      notes:           readField(row, 'notes'),
    }),
  },
  cleaning: {
    required: [{ key: 'branch_name', label: 'Cabang', aliases: ALIASES.branch_name }],
    map: (row) => ({
      activity_date: readDate(row, 'activity_date'),
      branch_name:   readField(row, 'branch_name'),
      activity_type: readField(row, 'activity_type') || 'General Cleaning',
      period:        readField(row, 'period'),
      status:        readStatus(row, 'Pending'),
      document_link: readField(row, 'document_link'),
      notes:         readField(row, 'notes'),
    }),
  },
  fogging: {
    // Fogging: tidak ada field wajib — semua diterima jika baris tidak kosong
    required: [],
    map: (row) => ({
      activity_date: readDate(row, 'activity_date'),
      branch_name:   readField(row, 'branch_name'),
      period:        readField(row, 'period'),
      status:        readStatus(row, 'Pending'),
      document_link: readField(row, 'document_link'),
      notes:         readField(row, 'notes'),
    }),
  },
  basecamp: {
    required: [{ key: 'problem', label: 'Permasalahan', aliases: ALIASES.problem }],
    map: (row) => ({
      info_date:   readDate(row, 'info_date'),
      branch_name: readField(row, 'branch_name'),
      problem:     readField(row, 'problem'),
      pic:         readField(row, 'pic'),
      done_date:   readDate(row, 'done_date'),
      status:      readStatus(row, 'Open'),
      notes:       readField(row, 'notes'),
    }),
  },
  supply: {
    required: [{ key: 'submitter_name', label: 'Nama Pemohon', aliases: ALIASES.submitter_name }],
    map: (row) => ({
      submitted_at:      readDate(row, 'submitted_at'),
      submitter_name:    readField(row, 'submitter_name'),
      branch_name:       readField(row, 'branch_name'),
      tools_items:       readField(row, 'tools_items'),
      tools_quantity:    readField(row, 'tools_quantity'),
      chemical_items:    readField(row, 'chemical_items'),
      chemical_quantity: readField(row, 'chemical_quantity'),
      additional_notes:  readField(row, 'additional_notes'),
      status:            readStatus(row, 'Pending'),
    }),
  },
};

// ─── Fuzzy Sheet Name Matching ────────────────────────────────────────────────
function findSheetMapping(sheetName) {
  const name = sheetName.trim();
  if (SHEET_MAP[name]) return SHEET_MAP[name];

  const nl = name.toLowerCase();
  for (const [key, val] of Object.entries(SHEET_MAP)) {
    if (key.toLowerCase() === nl) return val;
  }

  // Remove years and compare
  const nameNoYear = nl.replace(/\d{4}/g, '').trim().replace(/\s+/g, ' ');
  for (const [key, val] of Object.entries(SHEET_MAP)) {
    const kl = key.toLowerCase().replace(/\d{4}/g, '').trim().replace(/\s+/g, ' ');
    if (nameNoYear === kl || nameNoYear.includes(kl) || kl.includes(nameNoYear)) return val;
  }

  // Keyword
  if (/karyawan|pegawai|fc master|employee/.test(nl)) return SHEET_MAP['Master Karyawan'];
  if (/kontrak|contract/.test(nl)) return SHEET_MAP['Data Kontrak'];
  if (/reliefer|reliever|backup|pengganti/.test(nl)) return SHEET_MAP['Jadwal Reliefer'];
  if (/timeline|time line|jadwal kegiatan/.test(nl)) return SHEET_MAP['Time Line'];
  if (/masalah|permasalahan|issue|keluhan/.test(nl)) return SHEET_MAP['Permasalahan'];
  if (/one.on.one|1.on.1|oon/.test(nl)) return SHEET_MAP['One on One'];
  if (/training|pelatihan/.test(nl)) return SHEET_MAP['Training'];
  if (/checklist/.test(nl)) return SHEET_MAP['Master Checklist'];
  if (/form/.test(nl) && !/fogging/.test(nl)) return SHEET_MAP['Master Form'];
  if (/^sop$|sop /.test(nl) || / sop/.test(nl)) return SHEET_MAP['SOP'];
  if (/inspeksi|hygiene/.test(nl)) return SHEET_MAP['Report Inspeksi Hygiene 2026'];
  if (/gcdc|gc.dc|general clean|deep clean/.test(nl)) return SHEET_MAP['Report GCDC 2026'];
  if (/fogging/.test(nl)) return SHEET_MAP['Report Fogging 2026'];
  if (/basecamp|base.camp/.test(nl)) return SHEET_MAP['Rekap Laporan Basecamp'];
  if (/chemical|supply|permintaan/.test(nl)) return SHEET_MAP['Permintaan Chemical'];

  return null;
}

// ─── Main Validator ───────────────────────────────────────────────────────────
export function validateSheet(sheetName, rawRows, context = {}) {
  const mapping = findSheetMapping(sheetName);
  if (!mapping) return { valid: [], errors: [], mapped: [], skipped: true };

  const schema = SCHEMAS[mapping.module];
  if (!schema) return { valid: [], errors: [], mapped: [], skipped: true };

  const validRows  = [];
  const errorRows  = [];
  const mappedRows = [];

  // Filter baris kosong
  const nonEmptyRows = rawRows.filter(r => !isEmptyRow(r));

  nonEmptyRows.forEach((raw, idx) => {
    const rowNum = rawRows.indexOf(raw) + 2; // +2 utk header
    const errs   = [];

    // Validasi hanya required field
    schema.required.forEach(({ key, label, aliases }) => {
      const val = readField(raw, key);
      if (!val) {
        // Cari kolom yang ada di baris ini utk hint
        const existingCols = Object.keys(raw).filter(k => k.trim()).join(', ');
        errs.push({
          column: label,
          originalValue: '',
          reason: `Kolom "${label}" wajib diisi dan tidak ditemukan`,
          hint: `Kolom yang tersedia: ${existingCols.slice(0, 120)}`,
          aliases: (aliases || []).slice(0, 5).join(', '),
        });
      }
    });

    // Map row
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
  const results   = [];
  const processed = new Set();

  // Baca urutan sheet asli dari workbook
  workbook.SheetNames.forEach(sheetName => {
    // Skip sheet helper
    if (['Validasi','Dropdown','Ref','Helper','Config','Sheet1','Sheet2','Sheet3'].includes(sheetName)) return;
    if (sheetName.startsWith('_')) return;

    const mapping = findSheetMapping(sheetName);
    if (!mapping) {
      results.push({
        sheetName, module: null, label: '(Tidak dikenali — dilewati)',
        total: 0, valid: 0, errorCount: 0, errors: [], mapped: [], skipped: true,
      });
      return;
    }

    // Jika modul ini sudah diproses dari sheet lain, gabungkan (tidak skip)
    const ws      = workbook.Sheets[sheetName];
    const rawRows = window.XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
    const result  = validateSheet(sheetName, rawRows, {});

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

    processed.add(mapping.module);
  });

  return results;
}

// ─── Template Generator ───────────────────────────────────────────────────────
export function generateTemplate() {
  const XLSX = window.XLSX;
  const wb   = XLSX.utils.book_new();
  const TEMPLATES = {
    'Master Karyawan':              [{ 'Nama Lengkap':'Budi Santoso','Cabang':'001. Pondok Bambu','Divisi':'FACILITY CARE','No. HP':'081234567890','Tgl Masuk':'2024-01-15','Status':'Aktif','Catatan':'' }],
    'Data Kontrak':                 [{ 'Nama Karyawan':'Budi Santoso','Cabang':'001. Pondok Bambu','Divisi':'FACILITY CARE','Tgl Mulai':'2024-01-01','Tgl Selesai':'2024-12-31','Tipe Kontrak':'KONTRAK 1 TAHUN','PKWT':'PKWT 1','Status':'Aktif','Catatan':'' }],
    'Jadwal Reliefer':              [{ 'Tanggal Backup':'2024-03-10','Cabang':'001. Pondok Bambu','FC Digantikan':'Budi Santoso','Periode':'Q1','Reliefer':'Andi','Keterangan':'Sakit','Shift':'Pagi','Tanggal Selesai':'2024-03-10','Status':'Done' }],
    'Time Line':                    [{ 'Cabang':'001. Pondok Bambu','Kegiatan':'General Cleaning','Periode':'Q1','PIC':'Fajar','Tgl Opening':'2024-02-01','Tgl Target':'2024-02-15','Tgl Selesai':'2024-02-14','Status':'Done','Catatan':'' }],
    'Permasalahan':                 [{ 'Tanggal':'2024-03-01','Cabang':'001. Pondok Bambu','Kategori':'Cleaning','Sumber':'SPV','Keluhan':'Lantai kotor','Nama FC':'Budi','FC Spesialis':'Fajar','Solusi':'Teguran','Tgl Selesai':'2024-03-02','Status':'Done' }],
    'One on One':                   [{ 'Tanggal':'2024-03-05','Cabang':'001. Pondok Bambu','Nama Karyawan':'Budi Santoso','PIC':'Berlin','Permasalahan':'Keterlambatan','Solusi':'Coaching','Tgl Selesai':'2024-03-06','Status':'Done' }],
    'Training':                     [{ 'Tanggal':'2024-02-20','Batch':'Batch 1','Materi':'Hygiene & Sanitasi','Peserta':'Budi, Andi','Cabang':'001. Pondok Bambu','Trainer':'Fajar','Nilai':'85','Catatan':'' }],
    'Master Checklist':             [{ 'Nama':'Checklist Harian Toilet','Kategori':'Harian','Link Dokumen':'https://docs.google.com/...','Deskripsi':'' }],
    'Master Form':                  [{ 'Nama':'Form Permintaan Chemical','Kategori':'Chemical','Link Dokumen':'https://docs.google.com/...','Deskripsi':'' }],
    'SOP':                          [{ 'Nama':'SOP Pembersihan Toilet','Kategori':'Cleaning','Link Dokumen':'https://docs.google.com/...','Versi':'1.0','Tgl Berlaku':'2024-01-01','Catatan':'' }],
    'Report Inspeksi Hygiene 2026': [{ 'Tanggal':'2024-01-15','Cabang':'001. Pondok Bambu','Periode':'Q1','Status':'Done','Nilai FC':'85','Nilai SPV':'90','Link Dokumen':'','Catatan':'' }],
    'Report GCDC 2026':             [{ 'Tanggal':'2024-02-10','Cabang':'001. Pondok Bambu','Jenis Kegiatan':'General Cleaning','Periode':'Q1','Status':'Done','Link Dokumen':'','Catatan':'' }],
    'Report Fogging 2026':          [{ 'Tanggal':'2024-03-20','Cabang':'001. Pondok Bambu','Periode':'Q1','Status':'Done','Link Dokumen':'','Catatan':'' }],
    'Rekap Laporan Basecamp':       [{ 'Tanggal Info':'2024-01-10','Cabang':'001. Pondok Bambu','Permasalahan':'AC Rusak','PIC':'Berlin','Tgl Selesai':'2024-01-15','Status':'Done','Catatan':'' }],
    'Permintaan Chemical':          [{ 'Tanggal':'2024-02-01','Nama Pemohon':'Budi Santoso','Cabang':'001. Pondok Bambu','Item Tools':'Sapu','Qty Tools':'2','Item Chemical':'Sabun Lantai','Qty Chemical':'5 liter','Catatan':'','Status':'Pending' }],
  };
  Object.entries(TEMPLATES).forEach(([sn, data]) => {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), sn);
  });
  XLSX.writeFile(wb, 'Template_Import_Data_Awal_FCMS.xlsx');
}

// ─── Error Log Generator ──────────────────────────────────────────────────────
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
      'Cara Fix':       (e.errors || []).map(err => err.aliases ? `Gunakan salah satu: ${err.aliases}` : (err.hint || '')).join('; '),
      ...Object.fromEntries(Object.entries(e.data || {}).map(([k,v]) => [k, v ?? ''])),
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, r.sheetName.replace(/[\\\/\[\]*?:]/g,'_').slice(0,31));
  });

  if (!hasErrors) return false;
  XLSX.writeFile(wb, `Log_Error_Import_${new Date().toISOString().slice(0,10)}.xlsx`);
  return true;
}
