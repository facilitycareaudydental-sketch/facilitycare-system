-- Phase 4 & Observability Metadata Injections

-- ==========================================
-- 1. Observability Foundation
-- ==========================================
CREATE TABLE IF NOT EXISTS sync_metrics (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  event_id TEXT,
  action TEXT,
  module TEXT,
  webhook_duration_ms INTEGER DEFAULT 0,
  d1_execution_ms INTEGER DEFAULT 0,
  google_api_execution_ms INTEGER DEFAULT 0,
  queue_wait_time_ms INTEGER DEFAULT 0,
  retry_count INTEGER DEFAULT 0,
  created_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS feature_flags (
  flag_key TEXT PRIMARY KEY,
  flag_value TEXT,
  description TEXT,
  updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

INSERT OR IGNORE INTO feature_flags (flag_key, flag_value, description) VALUES 
('audit_enabled', 'true', 'Enable audit logs'),
('retry_enabled', 'true', 'Enable dead letter queue retry'),
('conflict_merge_enabled', 'true', 'Enable conflict OCC check'),
('bulk_import_enabled', 'true', 'Enable bulk import batching');

-- ==========================================
-- 2. Phase 4 Operational Modules Metadata
-- ==========================================

-- Timeline
ALTER TABLE activity_schedule ADD COLUMN row_version INTEGER DEFAULT 1;
ALTER TABLE activity_schedule ADD COLUMN last_sync_source TEXT DEFAULT 'FCMS';
ALTER TABLE activity_schedule ADD COLUMN deleted_at DATETIME DEFAULT NULL;
CREATE INDEX IF NOT EXISTS idx_activity_schedule_version ON activity_schedule(id, row_version);

-- Reliefer
ALTER TABLE relievers ADD COLUMN row_version INTEGER DEFAULT 1;
ALTER TABLE relievers ADD COLUMN last_sync_source TEXT DEFAULT 'FCMS';
ALTER TABLE relievers ADD COLUMN deleted_at DATETIME DEFAULT NULL;
CREATE INDEX IF NOT EXISTS idx_relievers_version ON relievers(id, row_version);

-- Training
ALTER TABLE training ADD COLUMN row_version INTEGER DEFAULT 1;
ALTER TABLE training ADD COLUMN last_sync_source TEXT DEFAULT 'FCMS';
ALTER TABLE training ADD COLUMN deleted_at DATETIME DEFAULT NULL;
CREATE INDEX IF NOT EXISTS idx_training_version ON training(id, row_version);

-- One on One
ALTER TABLE one_on_one ADD COLUMN row_version INTEGER DEFAULT 1;
ALTER TABLE one_on_one ADD COLUMN last_sync_source TEXT DEFAULT 'FCMS';
ALTER TABLE one_on_one ADD COLUMN deleted_at DATETIME DEFAULT NULL;
CREATE INDEX IF NOT EXISTS idx_one_on_one_version ON one_on_one(id, row_version);

-- Permasalahan
ALTER TABLE issues ADD COLUMN row_version INTEGER DEFAULT 1;
ALTER TABLE issues ADD COLUMN last_sync_source TEXT DEFAULT 'FCMS';
ALTER TABLE issues ADD COLUMN deleted_at DATETIME DEFAULT NULL;
CREATE INDEX IF NOT EXISTS idx_issues_version ON issues(id, row_version);

-- Hygiene
ALTER TABLE cleaning_reports ADD COLUMN row_version INTEGER DEFAULT 1;
ALTER TABLE cleaning_reports ADD COLUMN last_sync_source TEXT DEFAULT 'FCMS';
ALTER TABLE cleaning_reports ADD COLUMN deleted_at DATETIME DEFAULT NULL;
CREATE INDEX IF NOT EXISTS idx_cleaning_reports_version ON cleaning_reports(id, row_version);

-- GC-DC
ALTER TABLE inspection_reports ADD COLUMN row_version INTEGER DEFAULT 1;
ALTER TABLE inspection_reports ADD COLUMN last_sync_source TEXT DEFAULT 'FCMS';
ALTER TABLE inspection_reports ADD COLUMN deleted_at DATETIME DEFAULT NULL;
CREATE INDEX IF NOT EXISTS idx_inspection_reports_version ON inspection_reports(id, row_version);

-- Fogging
ALTER TABLE fogging_reports ADD COLUMN row_version INTEGER DEFAULT 1;
ALTER TABLE fogging_reports ADD COLUMN last_sync_source TEXT DEFAULT 'FCMS';
ALTER TABLE fogging_reports ADD COLUMN deleted_at DATETIME DEFAULT NULL;
CREATE INDEX IF NOT EXISTS idx_fogging_reports_version ON fogging_reports(id, row_version);

-- Basecamp
ALTER TABLE basecamp_reports ADD COLUMN row_version INTEGER DEFAULT 1;
ALTER TABLE basecamp_reports ADD COLUMN last_sync_source TEXT DEFAULT 'FCMS';
ALTER TABLE basecamp_reports ADD COLUMN deleted_at DATETIME DEFAULT NULL;
CREATE INDEX IF NOT EXISTS idx_basecamp_reports_version ON basecamp_reports(id, row_version);

-- Chemical
ALTER TABLE supply_requests ADD COLUMN row_version INTEGER DEFAULT 1;
ALTER TABLE supply_requests ADD COLUMN last_sync_source TEXT DEFAULT 'FCMS';
ALTER TABLE supply_requests ADD COLUMN deleted_at DATETIME DEFAULT NULL;
CREATE INDEX IF NOT EXISTS idx_supply_requests_version ON supply_requests(id, row_version);
