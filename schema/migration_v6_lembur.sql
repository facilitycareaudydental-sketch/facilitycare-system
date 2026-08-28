-- Migration v6 - Overtime Module

CREATE TABLE IF NOT EXISTS overtime_records (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  date TEXT NOT NULL,
  branch_id INTEGER NOT NULL,
  employee_id INTEGER NOT NULL,
  start_time TEXT NOT NULL,
  end_time TEXT NOT NULL,
  break_hours INTEGER DEFAULT 0,
  total_hours INTEGER NOT NULL,
  reason TEXT NOT NULL,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (branch_id) REFERENCES branches(id),
  FOREIGN KEY (employee_id) REFERENCES employees(id)
);
