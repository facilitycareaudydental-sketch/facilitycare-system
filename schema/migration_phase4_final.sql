
DROP TABLE IF EXISTS circuit_breaker;
CREATE TABLE IF NOT EXISTS circuit_breaker (
  id INTEGER PRIMARY KEY,
  state TEXT DEFAULT 'CLOSED',
  failure_count INTEGER DEFAULT 0,
  opened_at DATETIME,
  last_failure_at DATETIME,
  next_attempt_at DATETIME,
  updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
);
INSERT OR IGNORE INTO circuit_breaker (id, state, failure_count) VALUES (1, 'CLOSED', 0);

CREATE TABLE IF NOT EXISTS sync_locks (
  lock_id TEXT PRIMARY KEY,
  worker_id TEXT NOT NULL,
  acquired_at DATETIME DEFAULT CURRENT_TIMESTAMP,
  expires_at DATETIME NOT NULL
);

ALTER TABLE sync_outbox ADD COLUMN updated_at DATETIME DEFAULT CURRENT_TIMESTAMP;

