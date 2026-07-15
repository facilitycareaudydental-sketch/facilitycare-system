-- ============================================================
-- Migration v3: Performance Indexes untuk Dashboard
-- Dioptimalkan untuk 100.000+ record
-- Safe to run multiple times (IF NOT EXISTS)
-- ============================================================

-- ── employees ────────────────────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_employees_status
  ON employees(status);

CREATE INDEX IF NOT EXISTS idx_employees_branch_status
  ON employees(branch_id, status);

CREATE INDEX IF NOT EXISTS idx_employees_created
  ON employees(created_at DESC);

-- ── contracts ────────────────────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_contracts_status
  ON contracts(status);

CREATE INDEX IF NOT EXISTS idx_contracts_end_date
  ON contracts(end_date);

CREATE INDEX IF NOT EXISTS idx_contracts_status_end
  ON contracts(status, end_date);

CREATE INDEX IF NOT EXISTS idx_contracts_created
  ON contracts(created_at DESC);

-- ── issues ───────────────────────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_issues_status
  ON issues(status);

CREATE INDEX IF NOT EXISTS idx_issues_report_date
  ON issues(report_date DESC);

CREATE INDEX IF NOT EXISTS idx_issues_category
  ON issues(category);

CREATE INDEX IF NOT EXISTS idx_issues_branch
  ON issues(branch_id);

CREATE INDEX IF NOT EXISTS idx_issues_created
  ON issues(created_at DESC);

-- Covering index untuk trend query (12 bulan)
CREATE INDEX IF NOT EXISTS idx_issues_date_status
  ON issues(report_date, status);

CREATE INDEX IF NOT EXISTS idx_issues_completion_status
  ON issues(completion_date, status);

-- ── one_on_one ───────────────────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_oo1_status
  ON one_on_one(status);

CREATE INDEX IF NOT EXISTS idx_oo1_meeting_date
  ON one_on_one(meeting_date DESC);

CREATE INDEX IF NOT EXISTS idx_oo1_created
  ON one_on_one(created_at DESC);

-- ── activity_schedule ────────────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_schedule_status
  ON activity_schedule(status);

CREATE INDEX IF NOT EXISTS idx_schedule_target_date
  ON activity_schedule(target_date);

CREATE INDEX IF NOT EXISTS idx_schedule_branch
  ON activity_schedule(branch_id);

-- ── training ─────────────────────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_training_date
  ON training(training_date DESC);

CREATE INDEX IF NOT EXISTS idx_training_created
  ON training(created_at DESC);

-- ── inspection_reports ───────────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_inspection_date
  ON inspection_reports(inspection_date DESC);

CREATE INDEX IF NOT EXISTS idx_inspection_branch
  ON inspection_reports(branch_id);

CREATE INDEX IF NOT EXISTS idx_inspection_branch_date
  ON inspection_reports(branch_id, inspection_date DESC);

-- ── cleaning_reports ─────────────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_cleaning_date
  ON cleaning_reports(activity_date DESC);

CREATE INDEX IF NOT EXISTS idx_cleaning_created
  ON cleaning_reports(created_at DESC);

-- ── fogging_reports ──────────────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_fogging_date
  ON fogging_reports(activity_date DESC);

CREATE INDEX IF NOT EXISTS idx_fogging_created
  ON fogging_reports(created_at DESC);

-- ── supply_requests ──────────────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_supply_status
  ON supply_requests(status);

CREATE INDEX IF NOT EXISTS idx_supply_created
  ON supply_requests(created_at DESC);

-- ── relievers ────────────────────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_relievers_backup_date
  ON relievers(backup_date DESC);

CREATE INDEX IF NOT EXISTS idx_relievers_created
  ON relievers(created_at DESC);

-- ── sessions (auth performance) ──────────────────────────────
CREATE INDEX IF NOT EXISTS idx_sessions_expires
  ON sessions(expires_at);

CREATE INDEX IF NOT EXISTS idx_sessions_user
  ON sessions(user_id);

SELECT 'Migration v3: Performance indexes created successfully' AS status;
