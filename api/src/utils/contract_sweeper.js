/**
/**
 * contract_sweeper.js — FCMS Idempotent Contract Expiry Sweeper
 * 
 * Safely marks expired contracts (end_date < date('now')) as 'Tidak Aktif'
 * and records standard Audit Log (action: 'UPDATE', module: 'contracts')
 * with System Automation identity.
 */
import { buildAuditStatement, executeAtomicImportBatch } from './audit.js';
import { resolveMaintenanceFreezeState } from './maintenance.js';
import { getManagementBranchSubquery } from './branchMaster.js';

let lastSweepTimestamp = 0;
const SWEEP_INTERVAL_MS = 60 * 60 * 1000; // Throttle: at most once every 60 minutes

export async function sweepExpiredContracts(env, force = false) {
  const freezeState = await resolveMaintenanceFreezeState(env);
  if (freezeState.freezeActive) {
    return { skipped: true, reason: 'MAINTENANCE_FREEZE_ACTIVE', source: freezeState.source };
  }

  const now = Date.now();
  if (!force && (now - lastSweepTimestamp < SWEEP_INTERVAL_MS)) {
    return { skipped: true, reason: 'Throttled' };
  }
  lastSweepTimestamp = now;

  try {
    // 1. Find expired contracts strictly inside Management (000. Management) + official divisions
    const expired = await env.DB.prepare(`
      SELECT c.id, c.employee_id, COALESCE(e.full_name, 'Unknown Employee') as emp_name,
             c.status, c.end_date, c.pkwt_number, c.notes
      FROM contracts c
      LEFT JOIN employees e ON c.employee_id = e.id AND e.deleted_at IS NULL
      WHERE c.status = 'Aktif' 
        AND c.deleted_at IS NULL 
        AND c.division IN ('FACILITY CARE', 'SECURITY')
        AND ${getManagementBranchSubquery('c.branch_id')}
        AND (c.employee_id IS NULL OR (e.id IS NOT NULL AND e.division IN ('FACILITY CARE', 'SECURITY') AND ${getManagementBranchSubquery('e.branch_id')}))
        AND c.end_date IS NOT NULL 
        AND c.end_date != '' 
        AND c.end_date < date('now')
      LIMIT 100
    `).all();

    const items = expired.results || [];
    if (items.length === 0) {
      return { skipped: false, count: 0 };
    }

    const systemUser = { id: 0, full_name: 'System Automation', username: 'system_auto_expiry', role: 'system' };
    const stmts = [];

    for (const c of items) {
      const noteReason = `[AUTO-EXPIRED Masa berlaku berakhir ${c.end_date}]`;
      const updateStmt = env.DB.prepare(`
        UPDATE contracts 
        SET status = 'Tidak Aktif', 
            notes = COALESCE(notes || ' | ', '') || ?, 
            updated_at = CURRENT_TIMESTAMP 
        WHERE id = ? AND status = 'Aktif'
      `).bind(noteReason, c.id);

      const oldData = { id: c.id, employee_id: c.employee_id, employee_name: c.emp_name, status: c.status, end_date: c.end_date, pkwt_number: c.pkwt_number };
      const newData = { id: c.id, employee_id: c.employee_id, employee_name: c.emp_name, status: 'Tidak Aktif', end_date: c.end_date, pkwt_number: c.pkwt_number, reason: noteReason, auto_reason: 'CONTRACT_END_DATE_PASSED' };
      const details = { trigger: 'AUTO_EXPIRY_SWEEP', note_reason: noteReason };

      stmts.push(
        updateStmt,
        buildAuditStatement(env, systemUser, 'UPDATE', 'contracts', c.id, oldData, newData, details, 'SUCCESS', true)
      );
    }

    if (stmts.length > 0) {
      await executeAtomicImportBatch(env, stmts);
    }

    return { skipped: false, count: items.length };
  } catch (err) {
    console.error('sweepExpiredContracts error:', err?.message || err);
    return { error: err?.message || err };
  }
}
