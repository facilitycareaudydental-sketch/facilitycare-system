-- Migration v5: Allow NULL/empty status on all tables
-- Alasan: Google Sheet menggunakan status kosong yang valid, sistem tidak boleh memaksakan default

-- SQLite tidak mendukung ALTER COLUMN secara langsung.
-- Kita rebuild setiap tabel yang bermasalah dengan RENAME → CREATE → INSERT → DROP

-- ── activity_schedule ────────────────────────────────────────────────────────
ALTER TABLE activity_schedule RENAME TO _activity_schedule_old;
CREATE TABLE activity_schedule (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  branch_id INTEGER,
  activity_type TEXT NOT NULL,
  period TEXT,
  pic TEXT,
  opening_date TEXT,
  target_date TEXT,
  completion_date TEXT,
  status TEXT DEFAULT '',
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);
INSERT INTO activity_schedule SELECT * FROM _activity_schedule_old;
DROP TABLE _activity_schedule_old;

-- ── employees ────────────────────────────────────────────────────────────────
ALTER TABLE employees RENAME TO _employees_old;
CREATE TABLE employees (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  full_name TEXT NOT NULL,
  branch_id INTEGER,
  division TEXT NOT NULL DEFAULT 'FACILITY CARE',
  phone TEXT,
  join_date TEXT,
  status TEXT DEFAULT '',
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);
INSERT INTO employees SELECT * FROM _employees_old;
DROP TABLE _employees_old;

-- ── contracts ────────────────────────────────────────────────────────────────
ALTER TABLE contracts RENAME TO _contracts_old;
CREATE TABLE contracts (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  employee_id INTEGER,
  employee_name TEXT,
  branch_id INTEGER,
  division TEXT NOT NULL DEFAULT 'FACILITY CARE',
  start_date TEXT,
  end_date TEXT,
  contract_type TEXT,
  pkwt_number TEXT,
  status TEXT DEFAULT '',
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE SET NULL,
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);
INSERT INTO contracts SELECT * FROM _contracts_old;
DROP TABLE _contracts_old;

-- ── issues ───────────────────────────────────────────────────────────────────
ALTER TABLE issues RENAME TO _issues_old;
CREATE TABLE issues (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  report_date TEXT NOT NULL,
  branch_id INTEGER,
  category TEXT NOT NULL,
  source TEXT,
  complaint TEXT NOT NULL,
  employee_name TEXT,
  fc_specialist TEXT,
  solution TEXT,
  status TEXT DEFAULT '',
  completion_date TEXT,
  day_count INTEGER,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);
INSERT INTO issues SELECT * FROM _issues_old;
DROP TABLE _issues_old;

-- ── one_on_one ───────────────────────────────────────────────────────────────
ALTER TABLE one_on_one RENAME TO _one_on_one_old;
CREATE TABLE one_on_one (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  meeting_date TEXT NOT NULL,
  branch_id INTEGER,
  employee_name TEXT NOT NULL,
  pic TEXT,
  problem TEXT NOT NULL,
  solution TEXT,
  status TEXT DEFAULT '',
  completion_date TEXT,
  day_count INTEGER,
  document_link TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);
INSERT INTO one_on_one SELECT * FROM _one_on_one_old;
DROP TABLE _one_on_one_old;

-- ── relievers ────────────────────────────────────────────────────────────────
ALTER TABLE relievers RENAME TO _relievers_old;
CREATE TABLE relievers (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  branch_id INTEGER,
  original_fc_name TEXT,
  period TEXT,
  reliever_name TEXT NOT NULL,
  backup_date TEXT NOT NULL,
  completion_date TEXT,
  reason TEXT,
  shift TEXT,
  status TEXT DEFAULT '',
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);
INSERT INTO relievers SELECT * FROM _relievers_old;
DROP TABLE _relievers_old;

-- ── inspection_reports ───────────────────────────────────────────────────────
ALTER TABLE inspection_reports RENAME TO _inspection_reports_old;
CREATE TABLE inspection_reports (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  branch_id INTEGER,
  period TEXT NOT NULL,
  inspection_date TEXT NOT NULL,
  status TEXT DEFAULT '',
  fc_score REAL,
  spv_score REAL,
  document_link TEXT,
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);
INSERT INTO inspection_reports SELECT * FROM _inspection_reports_old;
DROP TABLE _inspection_reports_old;

-- ── cleaning_reports ─────────────────────────────────────────────────────────
ALTER TABLE cleaning_reports RENAME TO _cleaning_reports_old;
CREATE TABLE cleaning_reports (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  branch_id INTEGER,
  activity_type TEXT NOT NULL,
  period TEXT NOT NULL,
  activity_date TEXT NOT NULL,
  status TEXT DEFAULT '',
  document_link TEXT,
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);
INSERT INTO cleaning_reports SELECT * FROM _cleaning_reports_old;
DROP TABLE _cleaning_reports_old;

-- ── fogging_reports ──────────────────────────────────────────────────────────
ALTER TABLE fogging_reports RENAME TO _fogging_reports_old;
CREATE TABLE fogging_reports (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  branch_id INTEGER,
  activity_type TEXT NOT NULL DEFAULT 'Fogging',
  period TEXT NOT NULL,
  activity_date TEXT,
  status TEXT DEFAULT '',
  document_link TEXT,
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);
INSERT INTO fogging_reports SELECT * FROM _fogging_reports_old;
DROP TABLE _fogging_reports_old;

-- ── basecamp_reports ─────────────────────────────────────────────────────────
ALTER TABLE basecamp_reports RENAME TO _basecamp_reports_old;
CREATE TABLE basecamp_reports (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  branch_id INTEGER,
  problem TEXT NOT NULL,
  pic TEXT,
  info_date TEXT NOT NULL,
  done_date TEXT,
  status TEXT DEFAULT '',
  notes TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);
INSERT INTO basecamp_reports SELECT * FROM _basecamp_reports_old;
DROP TABLE _basecamp_reports_old;

-- ── supply_requests ──────────────────────────────────────────────────────────
ALTER TABLE supply_requests RENAME TO _supply_requests_old;
CREATE TABLE supply_requests (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  submitted_at TEXT NOT NULL DEFAULT (datetime('now')),
  submitter_name TEXT,
  branch_id INTEGER,
  branch_name TEXT,
  tools_items TEXT,
  tools_quantity TEXT,
  chemical_items TEXT,
  chemical_quantity TEXT,
  additional_notes TEXT,
  status TEXT DEFAULT '',
  processed_by TEXT,
  processed_at TEXT,
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);
INSERT INTO supply_requests SELECT * FROM _supply_requests_old;
DROP TABLE _supply_requests_old;
