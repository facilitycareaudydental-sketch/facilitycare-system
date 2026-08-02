-- Phase 2: Karyawan & Branches Migration

-- Employees
ALTER TABLE employees ADD COLUMN row_version INTEGER DEFAULT 1;
ALTER TABLE employees ADD COLUMN last_sync_source TEXT DEFAULT 'FCMS';

CREATE INDEX IF NOT EXISTS idx_employees_version ON employees(id, row_version);

-- Branches
ALTER TABLE branches ADD COLUMN row_version INTEGER DEFAULT 1;
ALTER TABLE branches ADD COLUMN last_sync_source TEXT DEFAULT 'FCMS';

CREATE INDEX IF NOT EXISTS idx_branches_version ON branches(id, row_version);

-- Data Migration (Initialize metadata)
UPDATE employees SET row_version = 1, last_sync_source = 'FCMS';
UPDATE branches SET row_version = 1, last_sync_source = 'FCMS';
