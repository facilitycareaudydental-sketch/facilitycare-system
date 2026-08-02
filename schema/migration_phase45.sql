
-- Phase 4.5: Production Hardening

-- 1. Distributed Lock Hardening (Lease Lock)
DROP TABLE IF EXISTS sync_locks;
CREATE TABLE IF NOT EXISTS sync_locks (
  lock_id TEXT PRIMARY KEY,
  worker_id TEXT NOT NULL,
  lease_token TEXT NOT NULL,
  lease_until DATETIME NOT NULL,
  heartbeat_at DATETIME NOT NULL,
  created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
);

-- 2. Correlation ID
ALTER TABLE sync_outbox ADD COLUMN correlation_id TEXT;
ALTER TABLE sync_metrics ADD COLUMN correlation_id TEXT;

-- 8. Audit Trail Upgrade
ALTER TABLE audit_logs ADD COLUMN correlation_id TEXT;
ALTER TABLE audit_logs ADD COLUMN source TEXT;
ALTER TABLE audit_logs ADD COLUMN old_value TEXT;
ALTER TABLE audit_logs ADD COLUMN new_value TEXT;
ALTER TABLE audit_logs ADD COLUMN latency INTEGER;

-- 3. Nightly Snapshot
CREATE TABLE IF NOT EXISTS sync_snapshots (
  id TEXT PRIMARY KEY,
  snapshot_date DATE NOT NULL,
  created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
  total_modules INTEGER,
  overall_status TEXT
);

CREATE TABLE IF NOT EXISTS snapshot_details (
  id TEXT PRIMARY KEY,
  snapshot_id TEXT NOT NULL,
  module TEXT NOT NULL,
  active_row_count INTEGER,
  deleted_row_count INTEGER,
  highest_row_version INTEGER,
  checksum TEXT,
  created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
  FOREIGN KEY(snapshot_id) REFERENCES sync_snapshots(id)
);

