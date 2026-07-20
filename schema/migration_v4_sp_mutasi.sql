-- Migration v4: Create sp_data and mutasi_data tables (safe with IF NOT EXISTS)

CREATE TABLE IF NOT EXISTS sp_data (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  tanggal TEXT,
  employee_name TEXT NOT NULL,
  branch_id INTEGER,
  sp_type TEXT,
  status TEXT DEFAULT 'Aktif',
  document_link TEXT,
  created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
  FOREIGN KEY (branch_id) REFERENCES branches(id)
);

CREATE TABLE IF NOT EXISTS mutasi_data (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  tanggal TEXT,
  employee_name TEXT NOT NULL,
  from_branch_id INTEGER,
  to_branch_id INTEGER,
  status TEXT DEFAULT 'Proses',
  document_link TEXT,
  created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
  FOREIGN KEY (from_branch_id) REFERENCES branches(id),
  FOREIGN KEY (to_branch_id) REFERENCES branches(id)
);
