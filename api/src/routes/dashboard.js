/**
 * dashboard.js — FCMS Dashboard API
 * Optimized for 100,000+ records using:
 * - Parallel queries via Promise.all
 * - Indexed columns only in WHERE/ORDER BY
 * - LIMIT on all list queries
 * - COUNT(*) instead of COUNT(id)
 * - Avoiding SELECT * on large tables
 * - strftime for date grouping (uses index on date columns)
 */
import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, unauthorized, error, forbidden } from '../utils/response.js';
import { sweepExpiredContracts } from '../utils/contract_sweeper.js';
import { isOperationalBranchCode, extractBranchCodeFromLabel, getManagementBranchSubquery, getManagementBranchSqlPredicate } from '../utils/branchMaster.js';
import { hasTableColumn } from '../utils/schemaCompat.js';

function isExcludedDashboardBranch(label) {
  const s = String(label || '').trim();
  if (!s) return false;
  if (/^(TEST_|\[UAT-|T\d+B?-)/i.test(s)) return true;
  const code = extractBranchCodeFromLabel(s);
  if (!code) return false;
  return !isOperationalBranchCode(code);
}

// Phase 5.48: Dashboard is internal authenticated data, not a public surface.
// Read authorization is enforced per request; visibility is uniform across the five
// official roles (all hold module-level read), while input_lembur (overtime-only
// legacy role) and restricted tokens are denied, matching their permission sets.
export async function handleDashboard(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'dashboard', 'read')) return forbidden(origin);

  const url  = new URL(request.url);
  const path = url.pathname.replace('/api/dashboard', '');

  if (path === '/dropdown-aggregates') return getDropdownAggregates(env, origin);
  if (path === '/kpi')               return getKPI(env, origin);
  if (path === '/stats' || path === '') return getStats(env, origin);
  if (path === '/issues-trend')      return getIssuesTrend(env, origin);
  if (path === '/contracts-chart')   return getContractsChart(env, origin);
  if (path === '/issues-summary')    return getIssuesSummary(env, origin);
  if (path === '/inspection-bar')    return guarded(getInspectionBar, request, env, origin, [
    ['month', INSPECTION_MONTH_PATTERN, 'Invalid month format. Expected two-digit month (01-12).']]);
  if (path === '/inspection-months') return getInspectionMonths(env, origin);
  if (path === '/contracts-expiring') return getContractsExpiring(env, origin);
  if (path === '/activity-log')      return getActivityLog(env, origin);
  if (path === '/calendar')          return guarded(getCalendarEvents, request, env, origin, [
    ['month', CALENDAR_MONTH_PATTERN, 'Invalid month format. Expected YYYY-MM.']]);
  if (path === '/schedule-chart')    return guarded(getScheduleChart, request, env, origin, [
    ['year', YEAR_PATTERN, 'Invalid year format. Expected YYYY.']]);
  if (path === '/activity-by-period') return getActivityByPeriod(env, origin);
  if (path === '/inspection-scores') return guarded(getInspectionScores, request, env, origin, [
    ['period', /^.{0,32}$/, 'Invalid period filter. Maximum 32 characters.']]);
  if (path === '/notifications')      return getNotifications(request, env, origin);
  if (path === '/reliefer-summary')   return guarded(getRelieferSummary, request, env, origin, [
    ['month', MONTH_PATTERN, 'Invalid month format. Expected YYYY-MM.']]);

  return error('Not found', 404, origin);
}

// Wraps an endpoint with Phase 5.48 parameter validation; malformed input -> 400.
function guarded(fn, request, env, origin, specs) {
  try {
    const url = new URL(request.url);
    for (const [name, pattern, message] of specs) {
      getValidatedParam(url, name, pattern, message);
    }
    return fn(request, env, origin);
  } catch (e) {
    return error(e.message, 400, origin);
  }
}

// ─── Helpers ─────────────────────────────────────────────────────────────────
function curMonthStr() {
  const n = new Date();
  return `${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,'0')}`;
}
function prevMonthStr() {
  const n = new Date();
  n.setMonth(n.getMonth() - 1);
  return `${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,'0')}`;
}
function parseFlexibleDate(d) {
  if (!d || d === '-') return '';
  d = String(d).trim();
  if (/^\d{5}$/.test(d)) {
    const utc_days = Math.floor(Number(d) - 25569);
    const date_info = new Date(utc_days * 86400 * 1000);
    return date_info.toISOString().split('T')[0];
  }
  if (d.match(/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}$/)) {
    const p = d.split(/[\/\-]/);
    return `${p[2]}-${p[1].padStart(2, '0')}-${p[0].padStart(2, '0')}`; // YYYY-MM-DD
  }
  return d.split('T')[0];
};
function monthLabels(count = 12) {
  const labels = [];
  const now = new Date();
  for (let i = count - 1; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    labels.push(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`);
  }
  return labels;
}
function futureMonthLabels(count = 6) {
  const labels = [];
  const now = new Date();
  for (let i = 0; i < count; i++) {
    const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
    labels.push(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`);
  }
  return labels;
}

// ─── Phase 5.48: query parameter validation ──────────────────────────────────
// All dashboard parameters are bound (no injection surface), but malformed values
// must be rejected consistently instead of silently degrading into empty charts.
const MONTH_PATTERN = /^\d{4}-(0[1-9]|1[0-2])$/;      // YYYY-MM
const CALENDAR_MONTH_PATTERN = MONTH_PATTERN;      // YYYY-MM (01-12)
const INSPECTION_MONTH_PATTERN = /^(0[1-9]|1[0-2])$/; // two-digit month (strftime('%m') contract)
const YEAR_PATTERN = /^\d{4}$/;                       // YYYY

function getValidatedParam(url, name, pattern, message) {
  const raw = url.searchParams.get(name);
  if (raw === null || raw === '') return null;
  if (!pattern.test(raw)) throw new Error(message);
  return raw;
}

// ─── /kpi — 12 counters + prev-month trend ───────────────────────────────────
// Consolidated single-pass per table to minimize D1 round-trips and table scans.
async function getKPI(env, origin) {
  // Opportunistic background sweep for expired contracts (throttled)
  sweepExpiredContracts(env).catch(() => {});

  const curM  = curMonthStr();
  const prevM = prevMonthStr();

  const [
    empAgg,
    contractAgg,
    issuesAgg,
    oo1Agg,
    schedAgg,
    supplyPending,
    totalBranches,
    trainingCur,
    fogCur,
    relieverDoneCur,
    relieverTotalCur
  ] = await Promise.all([
    // Single pass on active employees (replaces 2 separate queries) — Phase 5.66.10A.1: Management scope
    env.DB.prepare(
      `SELECT
         COUNT(*) as active_cur,
         SUM(CASE WHEN strftime('%Y-%m', e.created_at) = ? THEN 1 ELSE 0 END) as active_prev
       FROM employees e
       WHERE e.status = 'Aktif' AND e.deleted_at IS NULL
         AND e.division IN ('FACILITY CARE', 'SECURITY')
         AND ${getManagementBranchSubquery('e.branch_id')}`
    ).bind(prevM).first(),

    // Single pass on active contracts (aligned with /api/contracts & /notifications for manual-name and linked contracts)
    env.DB.prepare(
      `SELECT
         COUNT(*) as active_cur,
         SUM(CASE WHEN strftime('%Y-%m', c.created_at) = ? THEN 1 ELSE 0 END) as active_prev,
         SUM(CASE WHEN c.end_date IS NOT NULL AND c.end_date != '' AND c.end_date != '-' AND c.end_date NOT LIKE '2099%' AND CAST(julianday(c.end_date) - julianday(date('now', '+7 hours')) AS INTEGER) BETWEEN 0 AND 30 THEN 1 ELSE 0 END) as exp30
       FROM contracts c
       LEFT JOIN employees e ON c.employee_id = e.id
       WHERE c.status = 'Aktif' 
         AND c.deleted_at IS NULL
         AND c.division IN ('FACILITY CARE', 'SECURITY')
         AND ${getManagementBranchSubquery('c.branch_id')}
         AND (c.employee_id IS NULL OR (e.status = 'Aktif' AND e.deleted_at IS NULL AND e.division IN ('FACILITY CARE', 'SECURITY') AND ${getManagementBranchSubquery('e.branch_id')}))`
    ).bind(prevM).first(),

    // Single pass on open issues (replaces 2 separate queries)
    env.DB.prepare(
      `SELECT
         COUNT(*) as open_cur,
         SUM(CASE WHEN strftime('%Y-%m', report_date) = ? THEN 1 ELSE 0 END) as open_prev
       FROM issues
       WHERE status != 'Done' AND deleted_at IS NULL`
    ).bind(prevM).first(),

    // Single pass on open one_on_one (replaces 2 separate queries)
    env.DB.prepare(
      `SELECT
         COUNT(*) as open_cur,
         SUM(CASE WHEN strftime('%Y-%m', meeting_date) = ? THEN 1 ELSE 0 END) as open_prev
       FROM one_on_one
       WHERE status != 'Done' AND deleted_at IS NULL`
    ).bind(prevM).first(),

    // Single pass on activity_schedule (replaces 3 separate queries)
    env.DB.prepare(
      `SELECT
         COUNT(*) as total_sched,
         SUM(CASE WHEN activity_type = 'Inspeksi Hygiene' AND LOWER(TRIM(status)) IN ('done', 'selesai', 'completed') AND strftime('%Y-%m', COALESCE(completion_date, target_date)) = ? THEN 1 ELSE 0 END) as insp_cur,
         SUM(CASE WHEN activity_type IN ('General Cleaning', 'Deep Cleaning') AND LOWER(TRIM(status)) IN ('done', 'selesai', 'completed') AND strftime('%Y-%m', COALESCE(completion_date, target_date)) = ? THEN 1 ELSE 0 END) as clean_cur
       FROM activity_schedule
       WHERE deleted_at IS NULL`
    ).bind(curM, curM).first(),

    env.DB.prepare("SELECT COUNT(*) c FROM supply_requests WHERE status='Pending' AND deleted_at IS NULL").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM branches WHERE is_active=1 AND deleted_at IS NULL AND TRIM(code) BETWEEN '001' AND '070' AND LENGTH(TRIM(code)) = 3").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM training WHERE deleted_at IS NULL").first(),
    env.DB.prepare("SELECT COUNT(*) as c FROM fogging_reports WHERE deleted_at IS NULL AND LOWER(TRIM(status)) IN ('done', 'selesai', 'completed') AND strftime('%Y-%m',activity_date)=?").bind(curM).first(),
    env.DB.prepare("SELECT COUNT(*) as c FROM relievers WHERE deleted_at IS NULL AND LOWER(TRIM(status)) IN ('done', 'selesai', 'completed') AND strftime('%Y-%m',backup_date)=?").bind(curM).first(),
    env.DB.prepare("SELECT COUNT(DISTINCT reliever_name) as c FROM relievers WHERE deleted_at IS NULL AND reliever_name IS NOT NULL AND TRIM(reliever_name) != ''").first(),
  ]);

  return ok({
    employees:          { current: empAgg?.active_cur || 0,       prev: empAgg?.active_prev || 0 },
    contracts:          { current: contractAgg?.active_cur || 0,  prev: contractAgg?.active_prev || 0 },
    expiring30:         { current: contractAgg?.exp30 || 0 },
    issues:             { current: issuesAgg?.open_cur || 0,      prev: issuesAgg?.open_prev || 0 },
    one_on_one:         { current: oo1Agg?.open_cur || 0,         prev: oo1Agg?.open_prev || 0 },
    schedule:           { current: schedAgg?.total_sched || 0 },
    supply:             { current: supplyPending?.c || 0 },
    branches:           { current: totalBranches?.c || 0 },
    training_month:     { current: trainingCur?.c || 0 },
    inspection_month:   { current: schedAgg?.insp_cur || 0 },
    cleaning_month:     { current: schedAgg?.clean_cur || 0 },
    fogging_month:      { current: fogCur?.c || 0 },
    reliever_completed: { current: relieverDoneCur?.c || 0 },
    reliever_total:     { current: relieverTotalCur?.c || 0 },
  }, 200, origin);
}

// ─── /stats — backward-compat + expiring list + recent issues ────────────────
async function getStats(env, origin) {
  const [
    totalEmployees,
    totalBranches,
    contractAgg,
    issuesAgg,
    openOneOnOne,
    pendingSchedule,
    pendingSupply,
    totalTraining,
    recentIssues,
    expiringList,
    upcomingSchedule
  ] = await Promise.all([
    env.DB.prepare(
      `SELECT COUNT(*) c FROM employees e 
       WHERE e.status='Aktif' AND e.deleted_at IS NULL
         AND e.division IN ('FACILITY CARE', 'SECURITY')
         AND ${getManagementBranchSubquery('e.branch_id')}`
    ).first(),
    env.DB.prepare("SELECT COUNT(*) c FROM branches WHERE is_active=1 AND deleted_at IS NULL AND TRIM(code) BETWEEN '001' AND '070' AND LENGTH(TRIM(code)) = 3").first(),
    env.DB.prepare(
      `SELECT
         COUNT(*) as active_c,
         SUM(CASE WHEN c.end_date IS NOT NULL AND c.end_date != '' AND c.end_date != '-' AND c.end_date NOT LIKE '2099%' AND CAST(julianday(c.end_date) - julianday(date('now', '+7 hours')) AS INTEGER) BETWEEN 0 AND 30 THEN 1 ELSE 0 END) as expiring_c
       FROM contracts c
       LEFT JOIN employees e ON c.employee_id = e.id
       WHERE c.status = 'Aktif' 
         AND c.deleted_at IS NULL
         AND c.division IN ('FACILITY CARE', 'SECURITY')
         AND ${getManagementBranchSubquery('c.branch_id')}
         AND (c.employee_id IS NULL OR (e.status = 'Aktif' AND e.deleted_at IS NULL AND e.division IN ('FACILITY CARE', 'SECURITY') AND ${getManagementBranchSubquery('e.branch_id')}))`
    ).first(),
    env.DB.prepare(
      `SELECT
         SUM(CASE WHEN status != 'Done' THEN 1 ELSE 0 END) as open_c,
         SUM(CASE WHEN status = 'Done' THEN 1 ELSE 0 END) as done_c
       FROM issues
       WHERE deleted_at IS NULL`
    ).first(),
    env.DB.prepare("SELECT COUNT(*) c FROM one_on_one WHERE status != 'Done' AND deleted_at IS NULL").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM activity_schedule WHERE status = 'Pending' AND deleted_at IS NULL").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM supply_requests WHERE status = 'Pending' AND deleted_at IS NULL").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM training WHERE deleted_at IS NULL").first(),
    env.DB.prepare(
      `SELECT i.id, i.report_date, i.category, i.complaint, i.status,
       b.full_name as branch_name
       FROM issues i LEFT JOIN branches b ON i.branch_id = b.id
       WHERE i.deleted_at IS NULL
       ORDER BY i.created_at DESC LIMIT 5`
    ).all(),
    env.DB.prepare(
      `SELECT c.id, c.end_date, c.status, COALESCE(e.full_name, c.employee_name, '—') as emp_name,
       b.full_name as branch_name,
       CAST(julianday(c.end_date) - julianday(date('now', '+7 hours')) AS INTEGER) as days_remaining
       FROM contracts c 
       LEFT JOIN employees e ON c.employee_id = e.id
       LEFT JOIN branches b ON c.branch_id = b.id
       WHERE c.status='Aktif' AND c.deleted_at IS NULL
         AND c.end_date IS NOT NULL AND c.end_date != '' AND c.end_date != '-' AND c.end_date NOT LIKE '2099%'
         AND CAST(julianday(c.end_date) - julianday(date('now', '+7 hours')) AS INTEGER) BETWEEN 0 AND 30
         AND c.division IN ('FACILITY CARE', 'SECURITY')
         AND ${getManagementBranchSubquery('c.branch_id')}
         AND (c.employee_id IS NULL OR (e.status = 'Aktif' AND e.deleted_at IS NULL))
       ORDER BY c.end_date ASC LIMIT 10`
    ).all(),
    env.DB.prepare(
      `SELECT s.id, s.activity_type, s.target_date, s.status,
       b.full_name as branch_name
       FROM activity_schedule s LEFT JOIN branches b ON s.branch_id = b.id
       WHERE s.status != 'Done' AND s.deleted_at IS NULL AND s.target_date >= date('now')
       ORDER BY s.target_date ASC LIMIT 7`
    ).all()
  ]);

  return ok({
    stats: {
      total_employees:         totalEmployees?.c || 0,
      active_contracts:        contractAgg?.active_c || 0,
      expiring_contracts:      contractAgg?.expiring_c || 0,
      open_issues:             issuesAgg?.open_c || 0,
      open_one_on_one:         openOneOnOne?.c || 0,
      pending_schedule:        pendingSchedule?.c || 0,
      pending_supply_requests: pendingSupply?.c || 0,
      total_branches:          totalBranches?.c || 0,
      resolved_issues:         issuesAgg?.done_c || 0,
      total_training:          totalTraining?.c || 0,
    },
    recent_issues:      recentIssues.results    || [],
    expiring_contracts: expiringList.results    || [],
    upcoming_schedule:  upcomingSchedule.results || [],
  }, 200, origin);
}

// ─── /issues-trend — 12-month open vs closed ─────────────────────────────────
// Uses idx_issues_date_status and idx_issues_completion_status
async function getIssuesTrend(env, origin) {
  const since12m = new Date();
  since12m.setMonth(since12m.getMonth() - 11);
  const since = `${since12m.getFullYear()}-${String(since12m.getMonth()+1).padStart(2,'0')}-01`;

  const [openRows, closedRows] = await Promise.all([
    env.DB.prepare(
      `SELECT strftime('%Y-%m', report_date) m, COUNT(*) c
       FROM issues
       WHERE report_date >= ? AND status != 'Done' AND deleted_at IS NULL
       GROUP BY m ORDER BY m`
    ).bind(since).all(),
    env.DB.prepare(
      `SELECT strftime('%Y-%m', completion_date) m, COUNT(*) c
       FROM issues
       WHERE completion_date >= ? AND status = 'Done' AND deleted_at IS NULL
       GROUP BY m ORDER BY m`
    ).bind(since).all(),
  ]);

  const labels    = monthLabels(12);
  const openMap   = Object.fromEntries((openRows.results||[]).map(r=>[r.m, r.c]));
  const closedMap = Object.fromEntries((closedRows.results||[]).map(r=>[r.m, r.c]));

  return ok({
    labels,
    open:   labels.map(l => openMap[l]   || 0),
    closed: labels.map(l => closedMap[l] || 0),
  }, 200, origin);
}

// ─── /contracts-chart — contracts expiring next 6 months ─────────────────────
// Uses idx_contracts_status_end
async function getContractsChart(env, origin) {
  const rows = await env.DB.prepare(
    `SELECT strftime('%Y-%m', c.end_date) m, COUNT(*) c
     FROM contracts c
     LEFT JOIN employees e ON c.employee_id = e.id
     WHERE c.status='Aktif' AND c.deleted_at IS NULL
       AND c.end_date IS NOT NULL AND c.end_date != '' AND c.end_date != '-' AND c.end_date NOT LIKE '2099%'
       AND c.end_date >= date('now', '+7 hours') AND c.end_date <= date('now', '+7 hours', '+6 months')
       AND c.division IN ('FACILITY CARE', 'SECURITY')
       AND ${getManagementBranchSubquery('c.branch_id')}
       AND (c.employee_id IS NULL OR (e.status = 'Aktif' AND e.deleted_at IS NULL))
     GROUP BY m ORDER BY m`
  ).all();

  const labels = futureMonthLabels(6);
  const map    = Object.fromEntries((rows.results||[]).map(r=>[r.m, r.c]));

  return ok({ labels, data: labels.map(l => map[l] || 0) }, 200, origin);
}

// ─── /issues-summary — donut + by_status + by_branch ─────────────────────────
// Uses idx_issues_category, idx_issues_status, idx_issues_branch
async function getIssuesSummary(env, origin) {
  const [byCategory, byStatus, byBranch] = await Promise.all([
    env.DB.prepare(
      `SELECT COALESCE(category,'Lainnya') category, COUNT(*) c
       FROM issues WHERE deleted_at IS NULL GROUP BY category ORDER BY c DESC LIMIT 10`
    ).all(),
    env.DB.prepare(
      "SELECT COALESCE(status,'Tidak Diketahui') status, COUNT(*) c FROM issues WHERE deleted_at IS NULL GROUP BY status"
    ).all(),
    env.DB.prepare(
      `SELECT b.full_name branch_name, COUNT(*) c
       FROM issues i LEFT JOIN branches b ON i.branch_id=b.id
       WHERE i.deleted_at IS NULL
       GROUP BY i.branch_id ORDER BY c DESC LIMIT 10`
    ).all(),
  ]);

  return ok({
    by_category: (byCategory.results||[]).map(r=>({ category: r.category||'Lainnya', count: r.c||0 })),
    by_status:   (byStatus.results||[]).map(r=>({ status: r.status||'—', count: r.c||0 })),
    by_branch:   (byBranch.results||[])
      .filter(r => !isExcludedDashboardBranch(r.branch_name))
      .map(r=>({ branch_name: r.branch_name||'Tidak diketahui', count: r.c||0 })),
  }, 200, origin);
}

// ─── /inspection-bar — avg score per branch ──────────────────────────────────
// Uses idx_inspection_branch_date
async function getInspectionMonths(env, origin) {
  const rows = await env.DB.prepare(
    `SELECT DISTINCT strftime('%Y-%m', inspection_date) mo
     FROM inspection_reports
     WHERE deleted_at IS NULL
     ORDER BY mo`
  ).all();
  return ok((rows.results||[]).map(r => r.mo), 200, origin);
}

async function getInspectionBar(request, env, origin) {

  const url = new URL(request.url);
  const month = url.searchParams.get('month');
  
  let whereClause, bindVals;
  if (month) {
    const currentYear = new Date().getFullYear().toString();
    whereClause = "WHERE r.deleted_at IS NULL AND strftime('%m', r.inspection_date) = ? AND strftime('%Y', r.inspection_date) = ?";
    bindVals = [month, currentYear];
  } else {
    const since = (() => {
      const d = new Date(); d.setMonth(d.getMonth()-11);
      return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-01`;
    })();
    whereClause = "WHERE r.deleted_at IS NULL AND r.inspection_date >= ?";
    bindVals = [since];
  }

  const rows = await env.DB.prepare(
    `SELECT b.full_name branch_name,
     ROUND(AVG(CAST(r.fc_score AS REAL)),1) avg_fc,
     ROUND(AVG(CAST(r.spv_score AS REAL)),1) avg_spv,
     COUNT(*) total
     FROM inspection_reports r
     LEFT JOIN branches b ON r.branch_id=b.id
     ${whereClause}
     GROUP BY r.branch_id
     ORDER BY avg_fc DESC LIMIT 15`
  ).bind(...bindVals).all();

  const results = (rows.results || []).filter(r => !isExcludedDashboardBranch(r.branch_name));
  return ok({
    labels:  results.map(r => r.branch_name || 'N/A'),
    fc:      results.map(r => r.avg_fc  || 0),
    spv:     results.map(r => r.avg_spv || 0),
    totals:  results.map(r => r.total   || 0),
  }, 200, origin);
}

// ─── /contracts-expiring — full list sorted by end_date ──────────────────────
// Uses idx_contracts_status_end
async function getContractsExpiring(env, origin) {
  const hasConEmpName = await hasTableColumn(env.DB, 'contracts', 'employee_name');
  const empNameExpr = hasConEmpName ? "COALESCE(e.full_name, c.employee_name, '—')" : "COALESCE(e.full_name, '—')";
  const rows = await env.DB.prepare(
    `SELECT c.id, c.end_date, c.status, c.contract_type, c.pkwt_number,
     ${empNameExpr} emp_name,
     b.full_name branch_name,
     CAST(julianday(c.end_date)-julianday(date('now', '+7 hours')) AS INTEGER) days_remaining
     FROM contracts c
     LEFT JOIN employees e ON c.employee_id = e.id
     LEFT JOIN branches b ON c.branch_id=b.id
     WHERE c.status='Aktif' AND c.deleted_at IS NULL
       AND c.end_date IS NOT NULL AND c.end_date != '' AND c.end_date != '-' AND c.end_date NOT LIKE '2099%'
       AND c.end_date >= date('now', '+7 hours')
       AND c.division IN ('FACILITY CARE', 'SECURITY')
       AND ${getManagementBranchSubquery('c.branch_id')}
       AND (c.employee_id IS NULL OR (e.status = 'Aktif' AND e.deleted_at IS NULL))
     ORDER BY c.end_date ASC LIMIT 50`
  ).all();
  return ok(rows.results || [], 200, origin);
}

// ─── /activity-log — recent activity across all modules ──────────────────────
// All sub-queries + branches lookup run in a single parallel batch
async function getActivityLog(env, origin) {
  const hasConEmpName = await hasTableColumn(env.DB, 'contracts', 'employee_name');
  const contractLabelExpr = hasConEmpName
    ? "COALESCE(e.full_name, c.employee_name, 'Karyawan')"
    : "COALESCE(e.full_name, 'Karyawan')";
  const [issues, contracts, employees, oo1, training, supply, relievers, inspection, bRows] =
    await Promise.all([
      env.DB.prepare(
        `SELECT 'issue' type, COALESCE(complaint,'—') label, created_at,
         branch_id, category info
         FROM issues WHERE deleted_at IS NULL ORDER BY created_at DESC LIMIT 5`
      ).all(),
      env.DB.prepare(
        `SELECT 'contract' type, ${contractLabelExpr} label, c.created_at,
         c.branch_id, c.contract_type info
         FROM contracts c
         LEFT JOIN employees e ON e.id = c.employee_id AND e.deleted_at IS NULL
         WHERE c.deleted_at IS NULL
           AND c.division IN ('FACILITY CARE', 'SECURITY')
           AND ${getManagementBranchSubquery('c.branch_id')}
         ORDER BY c.created_at DESC LIMIT 5`
      ).all(),
      env.DB.prepare(
        `SELECT 'employee' type, e.full_name label, e.created_at,
         e.branch_id, e.status info
         FROM employees e 
         WHERE e.deleted_at IS NULL
           AND e.division IN ('FACILITY CARE', 'SECURITY')
           AND ${getManagementBranchSubquery('e.branch_id')}
         ORDER BY e.created_at DESC LIMIT 5`
      ).all(),
      env.DB.prepare(
        `SELECT 'one_on_one' type, COALESCE(employee_name,'—') label, created_at,
         branch_id, status info
         FROM one_on_one WHERE deleted_at IS NULL ORDER BY created_at DESC LIMIT 4`
      ).all(),
      env.DB.prepare(
        `SELECT 'training' type, COALESCE(subject,'Training') label, created_at,
         NULL branch_id, trainer info
         FROM training WHERE deleted_at IS NULL ORDER BY created_at DESC LIMIT 4`
      ).all(),
      env.DB.prepare(
        `SELECT 'supply' type, COALESCE(submitter_name,'Pemohon') label,
         submitted_at created_at, NULL branch_id, status info
         FROM supply_requests WHERE deleted_at IS NULL ORDER BY submitted_at DESC LIMIT 4`
      ).all(),
      env.DB.prepare(
        `SELECT 'reliever' type, COALESCE(reliever_name,'Reliefer') label, created_at,
         branch_id, reason info
         FROM relievers 
         WHERE deleted_at IS NULL AND LOWER(TRIM(status)) = 'done'
         ORDER BY created_at DESC LIMIT 3`
      ).all(),
      env.DB.prepare(
        `SELECT 'inspection' type, period label, created_at,
         branch_id, status info
         FROM inspection_reports WHERE deleted_at IS NULL ORDER BY created_at DESC LIMIT 3`
      ).all(),
      env.DB.prepare('SELECT id, full_name FROM branches WHERE deleted_at IS NULL').all()
    ]);

  const branchMap = Object.fromEntries((bRows.results||[]).map(b=>[b.id, b.full_name]));

  const all = [
    ...(issues.results||[]),
    ...(contracts.results||[]),
    ...(employees.results||[]),
    ...(oo1.results||[]),
    ...(training.results||[]),
    ...(supply.results||[]),
    ...(relievers.results||[]),
    ...(inspection.results||[]),
  ]
  .map(r => ({
    ...r,
    branch: r.branch_id ? (branchMap[r.branch_id] || null) : null,
    label:  (r.label && r.label !== '[object Object]') ? String(r.label).slice(0, 80) : '—',
  }))
  .sort((a,b) => (b.created_at||'').localeCompare(a.created_at||''))
  .slice(0, 20);

  return ok(all, 200, origin);
}

// ─── /calendar ───────────────────────────────────────────────────────────────
async function getCalendarEvents(request, env, origin) {
  const url   = new URL(request.url);
  const curM  = url.searchParams.get('month') || curMonthStr();
  const bind  = [curM];
  const hasConEmpName = await hasTableColumn(env.DB, 'contracts', 'employee_name');
  const contractEmpExpr = hasConEmpName ? "COALESCE(e.full_name, c.employee_name, '?')" : "COALESCE(e.full_name, '?')";

  const dateFilter = (col) => `strftime('%Y-%m',${col})=?`;

  const [
    sched, issR, relR, trainR, oneR, cleanR, inspR, fogR, baseR, supplyR, contrList
  ] = await Promise.all([
    env.DB.prepare(
      `SELECT s.id, s.target_date, s.opening_date, s.completion_date,
       s.activity_type, s.period, s.pic, s.status, s.notes,
       b.full_name branch_name
       FROM activity_schedule s LEFT JOIN branches b ON s.branch_id=b.id
       WHERE s.deleted_at IS NULL`
    ).all(),
    env.DB.prepare(
      `SELECT i.id,'issue' type,i.report_date event_date,i.category title,
       i.status,b.full_name branch_name,'red' color
       FROM issues i LEFT JOIN branches b ON i.branch_id=b.id
       WHERE i.deleted_at IS NULL AND ${dateFilter('i.report_date')}`
    ).bind(...bind).all(),
    env.DB.prepare(
      `SELECT r.id,'reliever' type,r.backup_date event_date,r.reliever_name title,
       r.status,b.full_name branch_name,'teal' color
       FROM relievers r LEFT JOIN branches b ON r.branch_id=b.id
       WHERE r.deleted_at IS NULL AND ${dateFilter('r.backup_date')} AND LOWER(TRIM(r.status)) = 'done'`
    ).bind(...bind).all(),
    env.DB.prepare(
      `SELECT t.id,'training' type,t.training_date event_date,t.subject title,
       'Done' status,'' branch_name,'indigo' color
       FROM training t WHERE t.deleted_at IS NULL AND ${dateFilter('t.training_date')}`
    ).bind(...bind).all(),
    env.DB.prepare(
      `SELECT o.id,'one_on_one' type,o.meeting_date event_date,'One on One: ' || o.employee_name title,
       o.status,b.full_name branch_name,'pink' color
       FROM one_on_one o LEFT JOIN branches b ON o.branch_id=b.id
       WHERE o.deleted_at IS NULL AND ${dateFilter('o.meeting_date')}`
    ).bind(...bind).all(),
    env.DB.prepare(
      `SELECT c.id,'cleaning' type,c.activity_date event_date,c.activity_type || ': ' || b.full_name title,
       c.status,b.full_name branch_name,'green' color
       FROM cleaning_reports c LEFT JOIN branches b ON c.branch_id=b.id
       WHERE c.deleted_at IS NULL AND ${dateFilter('c.activity_date')}`
    ).bind(...bind).all(),
    env.DB.prepare(
      `SELECT r.id,'inspection' type,r.inspection_date event_date,'Inspeksi: ' || b.full_name title,
       r.status,b.full_name branch_name,'blue' color
       FROM inspection_reports r LEFT JOIN branches b ON r.branch_id=b.id
       WHERE r.deleted_at IS NULL AND ${dateFilter('r.inspection_date')}`
    ).bind(...bind).all(),
    env.DB.prepare(
      `SELECT f.id,'fogging' type,f.activity_date event_date,'Fogging: ' || b.full_name title,
       f.status,b.full_name branch_name,'orange' color
       FROM fogging_reports f LEFT JOIN branches b ON f.branch_id=b.id
       WHERE f.deleted_at IS NULL AND ${dateFilter('f.activity_date')}`
    ).bind(...bind).all(),
    env.DB.prepare(
      `SELECT bp.id,'basecamp' type,bp.info_date event_date,'Basecamp: ' || bp.problem title,
       bp.status,b.full_name branch_name,'purple' color
       FROM basecamp_reports bp LEFT JOIN branches b ON bp.branch_id=b.id
       WHERE bp.deleted_at IS NULL AND ${dateFilter('bp.info_date')}`
    ).bind(...bind).all(),
    env.DB.prepare(
      `SELECT s.id,'supply' type,strftime('%Y-%m-%d',s.submitted_at) event_date,'Permintaan: ' || s.submitter_name title,
       s.status,COALESCE(b.full_name,s.branch_name) branch_name,'brown' color
       FROM supply_requests s LEFT JOIN branches b ON s.branch_id=b.id
       WHERE s.deleted_at IS NULL AND strftime('%Y-%m',s.submitted_at)=?`
    ).bind(...bind).all(),
    env.DB.prepare(
      `SELECT c.id, c.end_date, ${contractEmpExpr} emp_name
       FROM contracts c LEFT JOIN employees e ON c.employee_id=e.id
       WHERE c.status='Aktif' AND c.deleted_at IS NULL
         AND c.division IN ('FACILITY CARE', 'SECURITY')
         AND ${getManagementBranchSubquery('c.branch_id')}`
    ).all()
  ]);

  // Generate virtual contract reminder events dynamically (H-90, H-60, H-30, H-14, H-7, H-1, H-0)
  const contractEvents = [];
  const intervals = [90, 60, 30, 14, 7, 1, 0];
  (contrList.results || []).forEach(c => {
    if (!c.end_date) return;
    const end = new Date(c.end_date + 'T00:00:00');
    if (isNaN(end.getTime())) return;
    
    intervals.forEach(days => {
      const remDate = new Date(end.getTime());
      remDate.setDate(end.getDate() - days);
      const remStr = remDate.toISOString().slice(0, 10);
      
      if (remStr.startsWith(curM)) {
        const daysRemaining = Math.ceil((end.getTime() - Date.now()) / 86400000);
        const title = days === 0 
          ? `Hari H: Kontrak ${c.emp_name} Berakhir`
          : `Reminder H-${days}: Kontrak ${c.emp_name} Berakhir`;
        
        contractEvents.push({
          id: c.id,
          type: 'contract_expiry',
          event_date: remStr,
          title: title,
          status: 'Aktif',
          branch_name: '',
          days_remaining: daysRemaining,
          color: days <= 7 ? 'red' : days <= 30 ? 'orange' : 'yellow'
        });
      }
    });
  });

  // Process Schedule Events across target, opening, and completion dates
  const schedEvents = [];
  (sched.results || []).forEach(s => {
    if (s.deleted_at) return;
    const isoTarget = parseFlexibleDate(s.target_date);
    const isoOpening = parseFlexibleDate(s.opening_date);
    const isoCompletion = parseFlexibleDate(s.completion_date);

    const actType = s.activity_type || 'Kegiatan';
    const color = (actType === 'Inspeksi Hygiene' || actType === 'Inspeksi Hygiene & Aset Bangunan') ? 'blue'
                : actType === 'General Cleaning' ? 'green'
                : actType === 'Deep Cleaning' ? 'purple'
                : actType === 'Fogging' ? 'orange' : 'gray';

    const branchLabel = s.branch_name ? s.branch_name : 'Cabang';

    const baseInfo = {
      schedule_id: s.id,
      type: 'schedule',
      activity_type: actType,
      branch_name: s.branch_name || '',
      pic: s.pic || '',
      period: s.period || '',
      status: s.status || '',
      opening_date: s.opening_date || '',
      target_date: s.target_date || '',
      completion_date: s.completion_date || '',
      notes: s.notes || '',
      color
    };

    // 1. Target Date (Pelaksanaan Utama)
    if (isoTarget && isoTarget.startsWith(curM)) {
      schedEvents.push({
        ...baseInfo,
        id: `${s.id}_target`,
        event_date: isoTarget,
        title: `${actType} - ${branchLabel}`,
        date_type: 'Target Pelaksanaan'
      });
    }

    // 2. Opening Date (Opening Cabang/Klinik)
    if (isoOpening && isoOpening.startsWith(curM) && isoOpening !== isoTarget) {
      schedEvents.push({
        ...baseInfo,
        id: `${s.id}_opening`,
        event_date: isoOpening,
        title: `Opening: ${branchLabel}`,
        date_type: 'Opening Cabang',
        color: actType === 'Deep Cleaning' ? 'purple' : 'blue'
      });
    }

    // 3. Completion Date (Selesai Pelaksanaan)
    if (isoCompletion && isoCompletion.startsWith(curM) && isoCompletion !== isoTarget && isoCompletion !== isoOpening) {
      schedEvents.push({
        ...baseInfo,
        id: `${s.id}_done`,
        event_date: isoCompletion,
        title: `Selesai ${actType} - ${branchLabel}`,
        date_type: 'Selesai Pelaksanaan',
        color: 'green'
      });
    }
  });

  const events = [
    ...schedEvents,
    ...(relR.results||[]).map(r => ({ ...r, event_date: parseFlexibleDate(r.event_date) })),
    ...(cleanR.results||[]).map(c => ({ ...c, event_date: parseFlexibleDate(c.event_date) })),
    ...(inspR.results||[]).map(i => ({ ...i, event_date: parseFlexibleDate(i.event_date) })),
    ...(fogR.results||[]).map(f => ({ ...f, event_date: parseFlexibleDate(f.event_date) })),
    ...(baseR.results||[]).map(bp => ({ ...bp, event_date: parseFlexibleDate(bp.event_date) })),
    ...(issR.results||[]).map(i => ({ ...i, event_date: parseFlexibleDate(i.event_date) })),
    ...(trainR.results||[]).map(t => ({ ...t, event_date: parseFlexibleDate(t.event_date) })),
    ...(oneR.results||[]).map(o => ({ ...o, event_date: parseFlexibleDate(o.event_date) })),
    ...contractEvents
  ].filter(e => e.event_date && e.event_date.startsWith(curM))
   .sort((a,b)=>(a.event_date||'').localeCompare(b.event_date||''));

  return ok(events, 200, origin);
}

async function getActivityByPeriod(env, origin) {
  const rows = await env.DB.prepare(
    `SELECT activity_type,period,COUNT(*) total,
     SUM(CASE WHEN status='Done' THEN 1 ELSE 0 END) done
     FROM activity_schedule
     WHERE deleted_at IS NULL
     GROUP BY activity_type,period`
  ).all();
  return ok(rows.results||[], 200, origin);
}

async function getScheduleChart(request, env, origin) {
  const url = new URL(request.url);
  const year = url.searchParams.get('year') || new Date().getFullYear().toString();
  
  // Use target_date to match activities scheduled in the given year
  const rows = await env.DB.prepare(
    `SELECT activity_type, strftime('%m', target_date) as month, COUNT(*) as count 
     FROM activity_schedule 
     WHERE deleted_at IS NULL AND strftime('%Y', target_date) = ?
     GROUP BY activity_type, month`
  ).bind(year).all();

  const data = {
    'Inspeksi Hygiene': Array(12).fill(0),
    'General Cleaning': Array(12).fill(0),
    'Deep Cleaning': Array(12).fill(0),
    'Fogging': Array(12).fill(0)
  };

  if (rows && rows.results) {
    for (const r of rows.results) {
      let type = r.activity_type;
      // Normalizing names to match the frontend expectations
      if (type === 'Inspeksi') type = 'Inspeksi Hygiene';
      
      const m = parseInt(r.month, 10) - 1; // 0-11
      if (data[type] !== undefined && m >= 0 && m < 12) {
        data[type][m] = r.count;
      }
    }
  }
  
  return ok(data, 200, origin);
}

async function getInspectionScores(request, env, origin) {
  const url  = new URL(request.url);
  const period = url.searchParams.get('period')||'';
  const where  = period ? 'WHERE r.deleted_at IS NULL AND r.period=?' : 'WHERE r.deleted_at IS NULL';
  const bind   = period ? [period] : [];
  const rows   = await env.DB.prepare(
    `SELECT r.id,r.inspection_date,r.period,r.fc_score,r.spv_score,r.status,
     b.full_name branch_name
     FROM inspection_reports r LEFT JOIN branches b ON r.branch_id=b.id
     ${where} ORDER BY r.inspection_date ASC LIMIT 500`
  ).bind(...bind).all();
  return ok(rows.results||[], 200, origin);
}

async function getNotifications(request, env, origin) {
  const hasConEmpName = await hasTableColumn(env.DB, 'contracts', 'employee_name');
  const contractEmpExpr = hasConEmpName ? "COALESCE(e.full_name, c.employee_name, '?')" : "COALESCE(e.full_name, '?')";
  const [contracts, pendingSched, openIssues, pendingSupply] = await Promise.all([
    env.DB.prepare(
      `SELECT c.id, c.end_date, ${contractEmpExpr} as emp_name
       FROM contracts c LEFT JOIN employees e ON c.employee_id = e.id
       WHERE c.status = 'Aktif'
         AND c.deleted_at IS NULL
         AND c.division IN ('FACILITY CARE', 'SECURITY')
         AND ${getManagementBranchSubquery('c.branch_id')}
         AND (c.employee_id IS NULL OR (e.deleted_at IS NULL AND e.division IN ('FACILITY CARE', 'SECURITY') AND ${getManagementBranchSubquery('e.branch_id')}))`
    ).all(),
    env.DB.prepare(
      `SELECT id, activity_type, target_date FROM activity_schedule WHERE status = 'Pending' AND deleted_at IS NULL LIMIT 50`
    ).all(),
    env.DB.prepare(
      `SELECT id, category, report_date FROM issues WHERE status != 'Done' AND deleted_at IS NULL LIMIT 50`
    ).all(),
    env.DB.prepare(
      `SELECT id, submitter_name, submitted_at FROM supply_requests WHERE status = 'Pending' AND deleted_at IS NULL LIMIT 50`
    ).all()
  ]);

  const list = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // 1. Contract reminders
  const intervals = [90, 60, 30, 14, 7, 1, 0];
  (contracts.results || []).forEach(c => {
    if (!c.end_date) return;
    const end = new Date(c.end_date + 'T00:00:00');
    if (isNaN(end.getTime())) return;

    intervals.forEach(days => {
      const remDate = new Date(end.getTime());
      remDate.setDate(end.getDate() - days);
      remDate.setHours(0, 0, 0, 0);

      // Trigger reminder if today is within [remDate, remDate + 7 days]
      const diffTime = today.getTime() - remDate.getTime();
      const diffDays = diffTime / 86400000;
      if (diffDays >= 0 && diffDays <= 7) {
        const severity = days <= 7 ? 'danger' : days <= 30 ? 'warning' : 'info';
        const title = days === 0
          ? `Hari H: Kontrak ${c.emp_name} Berakhir!`
          : `Reminder H-${days}: Kontrak ${c.emp_name} Berakhir`;
        list.push({
          id: `contract-${c.id}-${days}`,
          title,
          type: 'contract',
          date: remDate.toISOString().slice(0, 10),
          severity
        });
      }
    });
  });

  // 2. Pending Schedule
  (pendingSched.results || []).forEach(s => {
    list.push({
      id: `schedule-${s.id}`,
      title: `Jadwal Pending: ${s.activity_type}`,
      type: 'schedule',
      date: s.target_date,
      severity: 'warning'
    });
  });

  // 3. Open Issues
  (openIssues.results || []).forEach(i => {
    list.push({
      id: `issue-${i.id}`,
      title: `Permasalahan Baru: ${i.category}`,
      type: 'issue',
      date: i.report_date,
      severity: 'danger'
    });
  });

  // 4. Pending Supply Requests
  (pendingSupply.results || []).forEach(s => {
    list.push({
      id: `supply-${s.id}`,
      title: `Permintaan Barang pending: ${s.submitter_name}`,
      type: 'supply',
      date: String(s.submitted_at).slice(0, 10),
      severity: 'info'
    });
  });

  // Sort by date descending
  list.sort((a, b) => b.date.localeCompare(a.date));

  return ok(list, 200, origin);
}

// ─── /reliefer-summary ───────────────────────────────────────────────────────
async function getRelieferSummary(request, env, origin) {
  const url = new URL(request.url);
  const curM = url.searchParams.get('month') || curMonthStr();
  
  const mParts = curM.split('-');
  const y = mParts[0];
  const mIdx = parseInt(mParts[1], 10) - 1;
  const indoMonth = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'][mIdx];
  const engMonth = ['January','February','March','April','May','June','July','August','September','October','November','December'][mIdx];

  const bind = [
    curM,
    `%${curM}%`,
    `%${indoMonth}%${y}%`,
    `%${engMonth}%${y}%`,
    `%${mParts[1]}%${y}%`
  ];

  const aggRow = await env.DB.prepare(`
    SELECT
      COUNT(*) as total,
      SUM(CASE WHEN LOWER(TRIM(status))='done' THEN 1 ELSE 0 END) as done
    FROM relievers 
    WHERE deleted_at IS NULL AND (
      strftime('%Y-%m', backup_date) = ? OR 
      backup_date LIKE ? OR 
      backup_date LIKE ? OR 
      backup_date LIKE ? OR 
      backup_date LIKE ?
    )`).bind(...bind).first();

  const done = aggRow?.done || 0;
  const total = aggRow?.total || 0;
  const percentage = total > 0 ? Math.round((done / total) * 100) : 0;

  return ok({
    month: curM,
    done,
    total,
    percentage
  }, 200, origin);
}

async function getDropdownAggregates(env, origin) {
  const [relieverRows, schedMonthRows, foggingRows, scheduleQRows] = await Promise.all([
    env.DB.prepare("SELECT strftime('%Y-%m', backup_date) as month, COUNT(*) as c FROM relievers WHERE deleted_at IS NULL AND LOWER(TRIM(status)) IN ('done', 'selesai', 'completed') AND backup_date IS NOT NULL GROUP BY month").all(),
    env.DB.prepare(
      `SELECT
         strftime('%Y-%m', COALESCE(completion_date, target_date)) as month,
         SUM(CASE WHEN activity_type = 'Inspeksi Hygiene' THEN 1 ELSE 0 END) as insp_c,
         SUM(CASE WHEN activity_type IN ('General Cleaning', 'Deep Cleaning') THEN 1 ELSE 0 END) as gcdc_c
       FROM activity_schedule
       WHERE deleted_at IS NULL
         AND activity_type IN ('Inspeksi Hygiene', 'General Cleaning', 'Deep Cleaning')
         AND LOWER(TRIM(status)) IN ('done', 'selesai', 'completed')
         AND COALESCE(completion_date, target_date) IS NOT NULL
       GROUP BY month`
    ).all(),
    env.DB.prepare("SELECT strftime('%Y-%m', activity_date) as month, COUNT(*) as c FROM fogging_reports WHERE deleted_at IS NULL AND LOWER(TRIM(status)) IN ('done', 'selesai', 'completed') AND activity_date IS NOT NULL GROUP BY month").all(),
    env.DB.prepare(`
      SELECT 
        period,
        COUNT(*) as c,
        SUM(CASE WHEN activity_type = 'Inspeksi Hygiene' THEN 1 ELSE 0 END) as insp,
        SUM(CASE WHEN activity_type = 'General Cleaning' THEN 1 ELSE 0 END) as gc,
        SUM(CASE WHEN activity_type = 'Deep Cleaning' THEN 1 ELSE 0 END) as dc,
        SUM(CASE WHEN activity_type = 'Fogging' THEN 1 ELSE 0 END) as fog
      FROM (
        SELECT period, activity_type FROM activity_schedule WHERE deleted_at IS NULL AND period IS NOT NULL AND period LIKE 'Q%'
        UNION ALL
        SELECT period, 'Fogging' as activity_type FROM fogging_reports WHERE deleted_at IS NULL AND period IS NOT NULL AND period LIKE 'Q%'
      )
      GROUP BY period
    `).all()
  ]);

  const toMap = (rows) => Object.fromEntries((rows.results || []).map(r => [r.month || r.period, r.c || 0]));
  const inspeksiByMonth = {};
  const gcdcByMonth = {};
  for (const r of (schedMonthRows.results || [])) {
    if (!r.month) continue;
    if (r.insp_c > 0) inspeksiByMonth[r.month] = r.insp_c;
    if (r.gcdc_c > 0) gcdcByMonth[r.month] = r.gcdc_c;
  }
  const scheduleDetails = Object.fromEntries((scheduleQRows.results || []).map(r => [
    r.period, 
    { total: r.c || 0, insp: r.insp || 0, gc: r.gc || 0, dc: r.dc || 0, fog: r.fog || 0 }
  ]));

  return ok({
    relievers_by_month: toMap(relieverRows),
    inspeksi_by_month: inspeksiByMonth,
    gcdc_by_month: gcdcByMonth,
    fogging_by_month: toMap(foggingRows),
    schedule_by_quarter: toMap(scheduleQRows),
    schedule_quarter_details: scheduleDetails
  }, 200, origin);
}
