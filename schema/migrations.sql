-- FM Operations Management System
-- Cloudflare D1 (SQLite) Schema
-- Run: wrangler d1 execute fm-operations-db --file=schema/migrations.sql

PRAGMA foreign_keys = ON;

-- ============================================================
-- USERS & ACCESS CONTROL
-- ============================================================
CREATE TABLE IF NOT EXISTS users (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  username TEXT NOT NULL UNIQUE,
  email TEXT NOT NULL UNIQUE,
  password_hash TEXT NOT NULL,
  full_name TEXT NOT NULL,
  role TEXT NOT NULL DEFAULT 'viewer',
  -- roles: superadmin, admin, manager, spv, viewer
  is_active INTEGER NOT NULL DEFAULT 1,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS roles (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL UNIQUE,
  description TEXT,
  permissions TEXT NOT NULL DEFAULT '{}',
  -- JSON: {"module": ["read","write","delete"]}
  created_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS sessions (
  id TEXT PRIMARY KEY,
  user_id INTEGER NOT NULL,
  expires_at TEXT NOT NULL,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
);

-- ============================================================
-- BRANCHES (CABANG)
-- ============================================================
CREATE TABLE IF NOT EXISTS branches (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  code TEXT NOT NULL UNIQUE,
  -- e.g. "001", "002", "A01"
  name TEXT NOT NULL,
  -- e.g. "Pondok Bambu"
  full_name TEXT NOT NULL,
  -- e.g. "001. Pondok Bambu"
  city TEXT,
  is_active INTEGER NOT NULL DEFAULT 1,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- ============================================================
-- EMPLOYEES (MASTER KARYAWAN)
-- ============================================================
CREATE TABLE IF NOT EXISTS employees (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  full_name TEXT NOT NULL,
  branch_id INTEGER,
  division TEXT NOT NULL DEFAULT 'FACILITY CARE',
  -- FACILITY CARE, SECURITY
  phone TEXT,
  join_date TEXT,
  status TEXT NOT NULL DEFAULT 'Aktif',
  -- Aktif, Tidak Aktif, Resign, Cut
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);

-- ============================================================
-- CONTRACTS (DATA KONTRAK)
-- ============================================================
CREATE TABLE IF NOT EXISTS contracts (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  employee_id INTEGER NOT NULL,
  branch_id INTEGER,
  division TEXT NOT NULL DEFAULT 'FACILITY CARE',
  start_date TEXT NOT NULL,
  end_date TEXT NOT NULL,
  contract_type TEXT,
  -- KONTRAK 6 BULAN, KONTRAK 1 TAHUN, etc.
  pkwt_number TEXT,
  -- PKWT 1, PKWT 2, etc.
  status TEXT NOT NULL DEFAULT 'Aktif',
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE,
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);

-- ============================================================
-- TIMELINE / ACTIVITY SCHEDULE
-- ============================================================
CREATE TABLE IF NOT EXISTS activity_schedule (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  branch_id INTEGER,
  activity_type TEXT NOT NULL,
  -- Inspeksi Hygiene & Aset Bangunan, General Cleaning, Deep Cleaning, Fogging
  period TEXT NOT NULL,
  -- Q1, Q2, Q3, Q4
  pic TEXT,
  opening_date TEXT,
  target_date TEXT,
  completion_date TEXT,
  status TEXT NOT NULL DEFAULT 'Pending',
  -- Pending, In Progress, Done
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);

-- ============================================================
-- ISSUES / PERMASALAHAN
-- ============================================================
CREATE TABLE IF NOT EXISTS issues (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  report_date TEXT NOT NULL,
  branch_id INTEGER,
  category TEXT NOT NULL,
  -- SDM, Cleaning, Aset, dll
  source TEXT,
  -- SPV, AM, Perawat, FC, Berlin, Ade, Pattrel, dst
  complaint TEXT NOT NULL,
  employee_name TEXT,
  -- Nama FC/Security yang bermasalah
  fc_specialist TEXT,
  -- Fajar, Miswar, Ade, dst
  solution TEXT,
  status TEXT NOT NULL DEFAULT 'Open',
  -- Open, In Progress, Done
  completion_date TEXT,
  day_count INTEGER,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);

-- ============================================================
-- ONE ON ONE
-- ============================================================
CREATE TABLE IF NOT EXISTS one_on_one (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  meeting_date TEXT NOT NULL,
  branch_id INTEGER,
  employee_name TEXT NOT NULL,
  pic TEXT,
  problem TEXT NOT NULL,
  solution TEXT,
  status TEXT NOT NULL DEFAULT 'Open',
  -- Open, Done
  completion_date TEXT,
  day_count INTEGER,
  document_link TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);

-- ============================================================
-- TRAINING
-- ============================================================
CREATE TABLE IF NOT EXISTS training (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  training_date TEXT NOT NULL,
  batch TEXT,
  subject TEXT NOT NULL,
  participants TEXT,
  -- JSON array of participant names
  branch_id INTEGER,
  trainer TEXT,
  score REAL,
  notes TEXT,
  document_link TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);

-- ============================================================
-- RELIEVERS / JADWAL RELIEFER
-- ============================================================
CREATE TABLE IF NOT EXISTS relievers (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  branch_id INTEGER,
  original_fc_name TEXT,
  -- Name of FC being replaced (or "BELUM ADA FC")
  period TEXT,
  -- Q1, Q2, Q3, Q4
  reliever_name TEXT NOT NULL,
  backup_date TEXT NOT NULL,
  completion_date TEXT,
  reason TEXT,
  -- Cuti, Mengisi Kekosongan, Back Up Training, Deep Cleaning, dll
  shift TEXT,
  -- Pagi, Siang, Full Shift, Middle
  status TEXT NOT NULL DEFAULT 'Pending',
  -- Pending, Done, Tidak Datang
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);

-- ============================================================
-- INSPECTION REPORTS
-- ============================================================
CREATE TABLE IF NOT EXISTS inspection_reports (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  branch_id INTEGER,
  period TEXT NOT NULL,
  -- Q1, Q2, Q3, Q4
  inspection_date TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'Pending',
  fc_score REAL,
  spv_score REAL,
  document_link TEXT,
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);

-- ============================================================
-- GC / DC REPORTS
-- ============================================================
CREATE TABLE IF NOT EXISTS cleaning_reports (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  branch_id INTEGER,
  activity_type TEXT NOT NULL,
  -- General Cleaning, Deep Cleaning
  period TEXT NOT NULL,
  activity_date TEXT NOT NULL,
  status TEXT NOT NULL DEFAULT 'Pending',
  document_link TEXT,
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);

-- ============================================================
-- FOGGING REPORTS
-- ============================================================
CREATE TABLE IF NOT EXISTS fogging_reports (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  branch_id INTEGER,
  activity_type TEXT NOT NULL DEFAULT 'Fogging',
  period TEXT NOT NULL,
  activity_date TEXT,
  status TEXT NOT NULL DEFAULT 'Pending',
  document_link TEXT,
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);

-- ============================================================
-- BASECAMP REPORTS (REKAP LAPORAN BASECAMP)
-- ============================================================
CREATE TABLE IF NOT EXISTS basecamp_reports (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  branch_id INTEGER,
  problem TEXT NOT NULL,
  pic TEXT,
  info_date TEXT NOT NULL,
  done_date TEXT,
  status TEXT NOT NULL DEFAULT 'Open',
  -- Open, In Progress, Done
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);

-- ============================================================
-- SOP
-- ============================================================
CREATE TABLE IF NOT EXISTS sop (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL,
  category TEXT NOT NULL,
  document_link TEXT,
  version TEXT,
  effective_date TEXT,
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- ============================================================
-- MASTER CHECKLIST
-- ============================================================
CREATE TABLE IF NOT EXISTS master_checklist (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL,
  category TEXT,
  document_link TEXT,
  description TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- ============================================================
-- MASTER FORMS
-- ============================================================
CREATE TABLE IF NOT EXISTS master_forms (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL,
  category TEXT,
  document_link TEXT,
  description TEXT,
  is_public INTEGER NOT NULL DEFAULT 0,
  -- 1 = accessible without login
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

-- ============================================================
-- CHEMICAL / SUPPLY REQUESTS (PUBLIC FORM)
-- ============================================================
CREATE TABLE IF NOT EXISTS supply_requests (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  submitted_at TEXT NOT NULL DEFAULT (datetime('now')),
  submitter_name TEXT,
  branch_id INTEGER,
  branch_name TEXT,
  tools_items TEXT,
  -- JSON array
  tools_quantity TEXT,
  chemical_items TEXT,
  -- JSON array
  chemical_quantity TEXT,
  additional_notes TEXT,
  status TEXT NOT NULL DEFAULT 'Pending',
  -- Pending, Diproses, Selesai
  processed_by TEXT,
  processed_at TEXT,
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);

-- ============================================================
-- VALIDASI (from Validasi sheet - dropdown references)
-- ============================================================
CREATE TABLE IF NOT EXISTS pic_list (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL UNIQUE,
  role TEXT,
  is_active INTEGER NOT NULL DEFAULT 1
);

-- ============================================================
-- INITIAL DATA SEEDS
-- ============================================================

-- Default roles
INSERT OR IGNORE INTO roles (name, description, permissions) VALUES 
  ('superadmin', 'Full access to all modules', '{"*":["read","write","delete","admin"]}'),
  ('admin', 'Admin access - all modules CRUD', '{"employees":["read","write","delete"],"contracts":["read","write","delete"],"schedule":["read","write","delete"],"issues":["read","write","delete"],"one_on_one":["read","write","delete"],"training":["read","write","delete"],"relievers":["read","write","delete"],"reports":["read","write","delete"],"sop":["read","write","delete"],"checklist":["read","write","delete"],"forms":["read","write","delete"],"users":["read","write","delete"],"branches":["read","write","delete"]}'),
  ('manager', 'Manager - read all, write most', '{"employees":["read","write"],"contracts":["read","write"],"schedule":["read","write"],"issues":["read","write"],"one_on_one":["read","write"],"training":["read","write"],"relievers":["read","write"],"reports":["read","write"],"sop":["read"],"checklist":["read"]}'),
  ('spv', 'Supervisor - read and limited write', '{"employees":["read"],"contracts":["read"],"schedule":["read"],"issues":["read","write"],"one_on_one":["read","write"],"relievers":["read"],"reports":["read"]}'),
  ('viewer', 'View only', '{"employees":["read"],"contracts":["read"],"schedule":["read"],"reports":["read"],"sop":["read"]}');

-- Default superadmin user (password: Admin@123 - change immediately)
-- bcrypt hash of "Admin@123" - we use a simple SHA256 approach in the worker
INSERT OR IGNORE INTO users (username, email, password_hash, full_name, role) VALUES 
  ('superadmin', 'admin@fm-operations.com', '$2a$10$placeholder_change_this', 'Super Admin', 'superadmin');

-- PIC List
INSERT OR IGNORE INTO pic_list (name, role) VALUES 
  ('Berlin', 'Manager'),
  ('Ade', 'FC Specialist'),
  ('Miswar', 'FC Specialist'),
  ('Fajar', 'FC Specialist'),
  ('Pattrel', 'AM'),
  ('Dentrel', 'AM'),
  ('SPV', 'Supervisor'),
  ('AM', 'Area Manager'),
  ('RCP', 'RCP'),
  ('Perawat', 'Nurse'),
  ('FC', 'Facility Care');

-- Branches seed data
INSERT OR IGNORE INTO branches (code, name, full_name) VALUES 
  ('000', 'Management', '000. Management'),
  ('001', 'Pondok Bambu', '001. Pondok Bambu'),
  ('002', 'Kemang', '002. Kemang'),
  ('003', 'Depok', '003. Depok'),
  ('004', 'Cipete', '004. Cipete'),
  ('005', 'Greenville', '005. Greenville'),
  ('006', 'Kelapa Gading MOI', '006. Kelapa Gading MOI'),
  ('007', 'Bekasi', '007. Bekasi'),
  ('008', 'PIK', '008. PIK'),
  ('009', 'Bintaro', '009. Bintaro'),
  ('010', 'Bogor', '010. Bogor'),
  ('011', 'Kuningan', '011. Kuningan'),
  ('012', 'BSD', '012. BSD'),
  ('013', 'Bandung Gatsu', '013. Bandung Gatsu'),
  ('014', 'Sunter', '014. Sunter'),
  ('015', 'Gading Serpong', '015. Gading Serpong'),
  ('016', 'Cibubur', '016. Cibubur'),
  ('017', 'Bandung Setrasari', '017. Bandung Setrasari'),
  ('018', 'Harapan Indah', '018. Harapan Indah'),
  ('019', 'Cinere', '019. Cinere'),
  ('020', 'Surabaya', '020. Surabaya'),
  ('021', 'Kelapa Gading Boulevard', '021. Kelapa Gading Boulevard'),
  ('022', 'Cikarang', '022. Cikarang'),
  ('023', 'Semarang', '023. Semarang'),
  ('024', 'Grand Wisata', '024. Grand Wisata'),
  ('025', 'Bogor Baru', '025. Bogor Baru'),
  ('026', 'Tangerang City', '026. Tangerang City'),
  ('027', 'Rawamangun', '027. Rawamangun'),
  ('028', 'Taman Palem', '028. Taman Palem'),
  ('029', 'Karawang', '029. Karawang'),
  ('030', 'Ciputat', '030. Ciputat'),
  ('031', 'Sawangan', '031. Sawangan'),
  ('032', 'Cawang', '032. Cawang'),
  ('033', 'Taman Mini', '033. Taman Mini'),
  ('034', 'Puri Indah', '034. Puri Indah'),
  ('035', 'Ciledug', '035. Ciledug'),
  ('036', 'Jagakarsa', '036. Jagakarsa'),
  ('037', 'Bali', '037. Bali'),
  ('038', 'Greenlake', '038. Greenlake'),
  ('039', 'Bandung Mekarwangi', '039. Bandung Mekarwangi'),
  ('040', 'Bekasi Timur', '040. Bekasi Timur'),
  ('041', 'Karawaci', '041. Karawaci'),
  ('042', 'Muara Karang', '042. Muara Karang'),
  ('043', 'Tebet', '043. Tebet'),
  ('044', 'Depok GDC', '044. Depok GDC'),
  ('045', 'FX Mall', '045. FX Mall'),
  ('046', 'Kaliurang Jogja', '046. Kaliurang Jogja'),
  ('047', 'Bogor Tajur', '047. Bogor Tajur'),
  ('048', 'PIK 2', '048. PIK 2'),
  ('049', 'Surabaya Merr', '049. Surabaya Merr'),
  ('050', 'Cibinong', '050. Cibinong'),
  ('051', 'Bintaro Sektor 9', '051. Bintaro Sektor 9'),
  ('052', 'BSD Rawa Buntu', '052. BSD Rawa Buntu'),
  ('053', 'Alam Sutera', '053. Alam Sutera'),
  ('054', 'Bogor Gunung Batu', '054. Bogor Gunung Batu'),
  ('055', 'Gajah Mada', '055. Gajah Mada'),
  ('056', 'Citra Garden 2', '056. Citra Garden 2'),
  ('057', 'Cikupa', '057. Cikupa'),
  ('058', 'Solo Baru', '058. Solo Baru'),
  ('059', 'Malang Suhat', '059. Malang Suhat'),
  ('060', 'Palmerah', '060. Palmerah'),
  ('061', 'Cirebon', '061. Cirebon'),
  ('062', 'Summarecon Bekasi', '062. Summarecon Bekasi'),
  ('063', 'Pasar Minggu', '063. Pasar Minggu'),
  ('064', 'Pondok Indah', '064. Pondok Indah'),
  ('065', 'Kebon Sirih', '065. Kebon Sirih'),
  ('066', 'Bandung KBP', '066. Bandung KBP'),
  ('067', 'Agora Mall', '067. Agora Mall'),
  ('068', 'Kota Wisata', '068. Kota Wisata'),
  ('069', 'Bekasi Kota Bintang', '069. Bekasi Kota Bintang'),
  ('070', 'J7', '070. J7'),
  ('A01', 'Aesthica', 'A01. Aesthica');

-- SOP seed data
INSERT OR IGNORE INTO sop (name, category, document_link) VALUES 
  ('SOP Keseluruhan', 'Ketentuan & Basic', 'https://docs.google.com/document/d/1sxCDfkFkTtZAnnE5_oy7uIymOggsA38y-Vp5TnkC-NY/edit?tab=t.0'),
  ('Standart Kualitas & Grooming', 'Kualitas & Grooming', 'https://docs.google.com/spreadsheets/d/17opZzg39V3gJ8z6fQsz2rD0kZt1IXFEy_g9CPol2fx4/edit?gid=1576198174#gid=1576198174');

-- Master Checklist seed
INSERT OR IGNORE INTO master_checklist (name, category, document_link) VALUES 
  ('Master Cleaning Program - Toilet', 'Master Cleaning Program', 'https://docs.google.com/spreadsheets/d/1euLGRgPPJMueZhnx_gDKCJQ2LPPVyDtb/edit?gid=1966419973#gid=1966419973');
