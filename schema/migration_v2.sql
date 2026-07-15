-- ============================================================
-- Migration v2: Schema relaxation for Import Data Awal
-- Safe to run multiple times (errors ignored)
-- ============================================================

-- contracts: add employee_name column (best-effort, ignore if exists)
ALTER TABLE contracts ADD COLUMN employee_name TEXT;

-- contracts: we can't remove NOT NULL in SQLite without recreating the table.
-- The INSERT in import.js now provides defaults, so existing NOT NULL constraints
-- on start_date / end_date won't block inserts that provide values.
-- For rows where dates are NULL, import.js now defaults to today().

-- activity_schedule: period can be nullable
-- SQLite cannot ALTER COLUMN constraints, handled at application layer.

-- All changes are backward-compatible.
SELECT 'Migration v2 complete' AS status;
