-- Phase 3 Metadata Injections

-- 1. Contracts
ALTER TABLE contracts ADD COLUMN row_version INTEGER DEFAULT 1;
ALTER TABLE contracts ADD COLUMN last_sync_source TEXT DEFAULT 'FCMS';
ALTER TABLE contracts ADD COLUMN deleted_at DATETIME DEFAULT NULL;
CREATE INDEX IF NOT EXISTS idx_contracts_version ON contracts(id, row_version);

-- 2. SOP
ALTER TABLE sop ADD COLUMN row_version INTEGER DEFAULT 1;
ALTER TABLE sop ADD COLUMN last_sync_source TEXT DEFAULT 'FCMS';
ALTER TABLE sop ADD COLUMN deleted_at DATETIME DEFAULT NULL;
CREATE INDEX IF NOT EXISTS idx_sop_version ON sop(id, row_version);

-- 3. Master Checklist
ALTER TABLE master_checklist ADD COLUMN row_version INTEGER DEFAULT 1;
ALTER TABLE master_checklist ADD COLUMN last_sync_source TEXT DEFAULT 'FCMS';
ALTER TABLE master_checklist ADD COLUMN deleted_at DATETIME DEFAULT NULL;
CREATE INDEX IF NOT EXISTS idx_master_checklist_version ON master_checklist(id, row_version);

-- 4. Master Forms
ALTER TABLE master_forms ADD COLUMN row_version INTEGER DEFAULT 1;
ALTER TABLE master_forms ADD COLUMN last_sync_source TEXT DEFAULT 'FCMS';
ALTER TABLE master_forms ADD COLUMN deleted_at DATETIME DEFAULT NULL;
CREATE INDEX IF NOT EXISTS idx_master_forms_version ON master_forms(id, row_version);

-- 5. Penyeragaman Soft Delete untuk modul Phase 2
ALTER TABLE branches ADD COLUMN deleted_at DATETIME DEFAULT NULL;
ALTER TABLE employees ADD COLUMN deleted_at DATETIME DEFAULT NULL;
