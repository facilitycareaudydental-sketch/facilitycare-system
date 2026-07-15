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
import { authenticate } from '../utils/auth.js';
import { ok, unauthorized } from '../utils/response.js';

export async function handleDashboard(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);

  const url  = new URL(request.url);
  const path = url.pathname.replace('/api/dashboard', '');

  if (path === '/kpi')               return getKPI(env, origin);
  if (path === '/stats' || path === '') return getStats(env, origin);
  if (path === '/issues-trend')      return getIssuesTrend(env, origin);
  if (path === '/contracts-chart')   return getContractsChart(env, origin);
  if (path === '/issues-summary')    return getIssuesSummary(env, origin);
  if (path === '/inspection-bar')    return getInspectionBar(env, origin);
  if (path === '/contracts-expiring') return getContractsExpiring(env, origin);
  if (path === '/activity-log')      return getActivityLog(env, origin);
  if (path === '/calendar')          return getCalendarEvents(request, env, origin);
  if (path === '/activity-by-period') return getActivityByPeriod(env, origin);
  if (path === '/inspection-scores') return getInspectionScores(request, env, origin);

  return ok({}, 200, origin);
}

// ─── Helpers ─────────────────────────────────────────────────────────────────
function curMonthStr() {
  const n = new Date();
  return `${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,'0')}`;
}
function prevMonthStr() {
  const n = new Date();
  const pm = n.getMonth() === 0
    ? `${n.getFullYear()-1}-12`
    : `${n.getFullYear()}-${String(n.getMonth()).padStart(2,'0')}`;
  return pm;
}
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

// ─── /kpi — 12 counters + prev-month trend ───────────────────────────────────
// All queries use indexed columns. COUNT(*) is fastest in SQLite.
async function getKPI(env, origin) {
  const curM  = curMonthStr();
  const prevM = prevMonthStr();

  // All 16 counts run in parallel — no sequential blocking
  const [
    empActive, empPrevMonth,
    contractActive,
    contractExp30,
    issuesOpen, issuesPrevOpen,
    oo1Open, oo1PrevOpen,
    schedPending,
    supplyPending,
    totalBranches,
    trainingCur,
    inspCur,
    cleanCur,
    fogCur,
    contractPrev,
  ] = await Promise.all([
    // Uses idx_employees_status
    env.DB.prepare("SELECT COUNT(*) c FROM employees WHERE status='Aktif'").first(),
    // Uses idx_employees_status + strftime filter
    env.DB.prepare("SELECT COUNT(*) c FROM employees WHERE status='Aktif' AND strftime('%Y-%m',created_at)=?").bind(prevM).first(),

    // Uses idx_contracts_status_end
    env.DB.prepare("SELECT COUNT(*) c FROM contracts WHERE status='Aktif' AND end_date>=date('now')").first(),

    // Uses idx_contracts_status_end (covering index)
    env.DB.prepare("SELECT COUNT(*) c FROM contracts WHERE status='Aktif' AND end_date BETWEEN date('now') AND date('now','+30 days')").first(),

    // Uses idx_issues_status
    env.DB.prepare("SELECT COUNT(*) c FROM issues WHERE status!='Done'").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM issues WHERE status!='Done' AND strftime('%Y-%m',report_date)=?").bind(prevM).first(),

    // Uses idx_oo1_status
    env.DB.prepare("SELECT COUNT(*) c FROM one_on_one WHERE status!='Done'").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM one_on_one WHERE status!='Done' AND strftime('%Y-%m',meeting_date)=?").bind(prevM).first(),

    // Uses idx_schedule_status
    env.DB.prepare("SELECT COUNT(*) c FROM activity_schedule WHERE status='Pending'").first(),

    // Uses idx_supply_status
    env.DB.prepare("SELECT COUNT(*) c FROM supply_requests WHERE status='Pending'").first(),

    // branches is tiny table
    env.DB.prepare("SELECT COUNT(*) c FROM branches WHERE is_active=1").first(),

    // Uses idx_training_date
    env.DB.prepare("SELECT COUNT(*) c FROM training WHERE strftime('%Y-%m',training_date)=?").bind(curM).first(),

    // Uses idx_inspection_date
    env.DB.prepare("SELECT COUNT(*) c FROM inspection_reports WHERE strftime('%Y-%m',inspection_date)=?").bind(curM).first(),

    // Uses idx_cleaning_date
    env.DB.prepare("SELECT COUNT(*) c FROM cleaning_reports WHERE strftime('%Y-%m',activity_date)=?").bind(curM).first(),

    // Uses idx_fogging_date
    env.DB.prepare("SELECT COUNT(*) c FROM fogging_reports WHERE strftime('%Y-%m',activity_date)=?").bind(curM).first(),

    // Uses idx_contracts_created
    env.DB.prepare("SELECT COUNT(*) c FROM contracts WHERE status='Aktif' AND strftime('%Y-%m',created_at)=?").bind(prevM).first(),
  ]);

  return ok({
    employees:       { current: empActive?.c||0,       prev: empPrevMonth?.c||0 },
    contracts:       { current: contractActive?.c||0,  prev: contractPrev?.c||0 },
    expiring30:      { current: contractExp30?.c||0 },
    issues:          { current: issuesOpen?.c||0,      prev: issuesPrevOpen?.c||0 },
    one_on_one:      { current: oo1Open?.c||0,         prev: oo1PrevOpen?.c||0 },
    schedule:        { current: schedPending?.c||0 },
    supply:          { current: supplyPending?.c||0 },
    branches:        { current: totalBranches?.c||0 },
    training_month:  { current: trainingCur?.c||0 },
    inspection_month:{ current: inspCur?.c||0 },
    cleaning_month:  { current: cleanCur?.c||0 },
    fogging_month:   { current: fogCur?.c||0 },
  }, 200, origin);
}

// ─── /stats — backward-compat + expiring list + recent issues ────────────────
async function getStats(env, origin) {
  const [
    totalEmployees, activeContracts, expiringContracts, openIssues,
    openOneOnOne, pendingSchedule, pendingSupply, totalBranches,
    doneIssues, totalTraining
  ] = await Promise.all([
    env.DB.prepare("SELECT COUNT(*) c FROM employees WHERE status='Aktif'").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM contracts WHERE status='Aktif' AND end_date>=date('now')").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM contracts WHERE status='Aktif' AND end_date BETWEEN date('now') AND date('now','+30 days')").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM issues WHERE status!='Done'").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM one_on_one WHERE status!='Done'").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM activity_schedule WHERE status='Pending'").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM supply_requests WHERE status='Pending'").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM branches WHERE is_active=1").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM issues WHERE status='Done'").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM training").first(),
  ]);

  // Uses idx_issues_created (only 5 rows returned)
  const recentIssues = await env.DB.prepare(
    `SELECT i.id, i.report_date, i.category, i.complaint, i.status, i.employee_name,
     b.full_name as branch_name
     FROM issues i LEFT JOIN branches b ON i.branch_id = b.id
     ORDER BY i.created_at DESC LIMIT 5`
  ).all();

  // Uses idx_contracts_status_end (only 10 rows)
  const expiringList = await env.DB.prepare(
    `SELECT c.id, c.end_date, c.status, c.contract_type,
     COALESCE(e.full_name, c.employee_name, '—') as emp_name,
     b.full_name as branch_name,
     CAST(julianday(c.end_date) - julianday('now') AS INTEGER) as days_remaining
     FROM contracts c
     LEFT JOIN employees e ON c.employee_id = e.id
     LEFT JOIN branches b ON c.branch_id = b.id
     WHERE c.status='Aktif' AND c.end_date BETWEEN date('now') AND date('now','+30 days')
     ORDER BY c.end_date ASC LIMIT 10`
  ).all();

  // Uses idx_schedule_target_date (only 7 rows)
  const upcomingSchedule = await env.DB.prepare(
    `SELECT s.id, s.activity_type, s.period, s.target_date, s.status, s.pic,
     b.full_name as branch_name
     FROM activity_schedule s LEFT JOIN branches b ON s.branch_id = b.id
     WHERE s.status != 'Done' AND s.target_date >= date('now')
     ORDER BY s.target_date ASC LIMIT 7`
  ).all();

  return ok({
    stats: {
      total_employees:        totalEmployees?.c || 0,
      active_contracts:       activeContracts?.c || 0,
      expiring_contracts:     expiringContracts?.c || 0,
      open_issues:            openIssues?.c || 0,
      open_one_on_one:        openOneOnOne?.c || 0,
      pending_schedule:       pendingSchedule?.c || 0,
      pending_supply_requests:pendingSupply?.c || 0,
      total_branches:         totalBranches?.c || 0,
      resolved_issues:        doneIssues?.c || 0,
      total_training:         totalTraining?.c || 0,
    },
    recent_issues:      recentIssues.results   || [],
    expiring_contracts: expiringList.results   || [],
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
       WHERE report_date >= ? AND status != 'Done'
       GROUP BY m ORDER BY m`
    ).bind(since).all(),
    env.DB.prepare(
      `SELECT strftime('%Y-%m', completion_date) m, COUNT(*) c
       FROM issues
       WHERE completion_date >= ? AND status = 'Done'
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
    `SELECT strftime('%Y-%m', end_date) m, COUNT(*) c
     FROM contracts
     WHERE status='Aktif' AND end_date >= date('now') AND end_date <= date('now','+6 months')
     GROUP BY m ORDER BY m`
  ).all();

  const labels = futureMonthLabels(6);
  const map    = Object.fromEntries((rows.results||[]).map(r=>[r.m, r.c]));

  return ok({ labels, counts: labels.map(l => map[l] || 0) }, 200, origin);
}

// ─── /issues-summary — donut + by_status + by_branch ─────────────────────────
// Uses idx_issues_category, idx_issues_status, idx_issues_branch
async function getIssuesSummary(env, origin) {
  const [byCategory, byStatus, byBranch] = await Promise.all([
    env.DB.prepare(
      `SELECT COALESCE(category,'Lainnya') category, COUNT(*) c
       FROM issues GROUP BY category ORDER BY c DESC LIMIT 10`
    ).all(),
    env.DB.prepare(
      "SELECT COALESCE(status,'Tidak Diketahui') status, COUNT(*) c FROM issues GROUP BY status"
    ).all(),
    env.DB.prepare(
      `SELECT b.full_name branch_name, COUNT(*) c
       FROM issues i LEFT JOIN branches b ON i.branch_id=b.id
       GROUP BY i.branch_id ORDER BY c DESC LIMIT 10`
    ).all(),
  ]);

  return ok({
    by_category: (byCategory.results||[]).map(r=>({ category: r.category||'Lainnya', count: r.c||0 })),
    by_status:   (byStatus.results||[]).map(r=>({ status: r.status||'—', count: r.c||0 })),
    by_branch:   (byBranch.results||[]).map(r=>({ branch_name: r.branch_name||'Tidak diketahui', count: r.c||0 })),
  }, 200, origin);
}

// ─── /inspection-bar — avg score per branch (last 6 months) ──────────────────
// Uses idx_inspection_branch_date
async function getInspectionBar(env, origin) {
  const since = (() => {
    const d = new Date(); d.setMonth(d.getMonth()-5);
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-01`;
  })();

  const rows = await env.DB.prepare(
    `SELECT b.full_name branch_name,
     ROUND(AVG(CAST(r.fc_score AS REAL)),1) avg_fc,
     ROUND(AVG(CAST(r.spv_score AS REAL)),1) avg_spv,
     COUNT(*) total
     FROM inspection_reports r
     LEFT JOIN branches b ON r.branch_id=b.id
     WHERE r.inspection_date >= ?
     GROUP BY r.branch_id
     ORDER BY avg_fc DESC LIMIT 15`
  ).bind(since).all();

  const results = rows.results || [];
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
  const rows = await env.DB.prepare(
    `SELECT c.id, c.end_date, c.status, c.contract_type, c.pkwt_number,
     COALESCE(e.full_name, c.employee_name, '—') emp_name,
     b.full_name branch_name,
     CAST(julianday(c.end_date)-julianday('now') AS INTEGER) days_remaining
     FROM contracts c
     LEFT JOIN employees e ON c.employee_id=e.id
     LEFT JOIN branches b ON c.branch_id=b.id
     WHERE c.status='Aktif' AND c.end_date >= date('now')
     ORDER BY c.end_date ASC LIMIT 50`
  ).all();
  return ok(rows.results || [], 200, origin);
}

// ─── /activity-log — recent activity across all modules ──────────────────────
// Each sub-query uses its own created_at DESC index, LIMIT 5 each
async function getActivityLog(env, origin) {
  const [issues, contracts, employees, oo1, training, supply, relievers, inspection] =
    await Promise.all([
      env.DB.prepare(
        `SELECT 'issue' type, COALESCE(complaint,'—') label, created_at,
         branch_id, category info
         FROM issues ORDER BY created_at DESC LIMIT 5`
      ).all(),
      env.DB.prepare(
        `SELECT 'contract' type, COALESCE(employee_name,'Karyawan') label, created_at,
         branch_id, contract_type info
         FROM contracts ORDER BY created_at DESC LIMIT 5`
      ).all(),
      env.DB.prepare(
        `SELECT 'employee' type, full_name label, created_at,
         branch_id, status info
         FROM employees ORDER BY created_at DESC LIMIT 5`
      ).all(),
      env.DB.prepare(
        `SELECT 'one_on_one' type, COALESCE(employee_name,'—') label, created_at,
         branch_id, status info
         FROM one_on_one ORDER BY created_at DESC LIMIT 4`
      ).all(),
      env.DB.prepare(
        `SELECT 'training' type, COALESCE(subject,'Training') label, created_at,
         NULL branch_id, trainer info
         FROM training ORDER BY created_at DESC LIMIT 4`
      ).all(),
      env.DB.prepare(
        `SELECT 'supply' type, COALESCE(submitter_name,'Pemohon') label, created_at,
         NULL branch_id, status info
         FROM supply_requests ORDER BY created_at DESC LIMIT 4`
      ).all(),
      env.DB.prepare(
        `SELECT 'reliever' type, COALESCE(reliever_name,'Reliefer') label, created_at,
         branch_id, reason info
         FROM relievers ORDER BY created_at DESC LIMIT 3`
      ).all(),
      env.DB.prepare(
        `SELECT 'inspection' type, period label, created_at,
         branch_id, status info
         FROM inspection_reports ORDER BY created_at DESC LIMIT 3`
      ).all(),
    ]);

  // Resolve branch names once — branches table is tiny
  const bRows = await env.DB.prepare('SELECT id, full_name FROM branches').all();
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

  const dateFilter = (col) => `strftime('%Y-%m',${col})=?`;

  const [sched, issR, relR, trainR, contrR] = await Promise.all([
    env.DB.prepare(
      `SELECT s.id,'schedule' type,s.target_date event_date,s.activity_type title,
       s.status,s.pic,b.full_name branch_name,
       CASE s.activity_type
         WHEN 'Inspeksi Hygiene & Aset Bangunan' THEN 'blue'
         WHEN 'General Cleaning' THEN 'green'
         WHEN 'Deep Cleaning' THEN 'purple'
         WHEN 'Fogging' THEN 'orange'
         ELSE 'gray' END color
       FROM activity_schedule s LEFT JOIN branches b ON s.branch_id=b.id
       WHERE ${dateFilter('s.target_date')}`
    ).bind(...bind).all(),
    env.DB.prepare(
      `SELECT i.id,'issue' type,i.report_date event_date,i.category title,
       i.status,b.full_name branch_name,'red' color
       FROM issues i LEFT JOIN branches b ON i.branch_id=b.id
       WHERE ${dateFilter('i.report_date')}`
    ).bind(...bind).all(),
    env.DB.prepare(
      `SELECT r.id,'reliever' type,r.backup_date event_date,r.reliever_name title,
       r.status,b.full_name branch_name,'teal' color
       FROM relievers r LEFT JOIN branches b ON r.branch_id=b.id
       WHERE ${dateFilter('r.backup_date')}`
    ).bind(...bind).all(),
    env.DB.prepare(
      `SELECT t.id,'training' type,t.training_date event_date,t.subject title,
       'Done' status,'' branch_name,'indigo' color
       FROM training t WHERE ${dateFilter('t.training_date')}`
    ).bind(...bind).all(),
    env.DB.prepare(
      `SELECT c.id,'contract_expiry' type,c.end_date event_date,
       COALESCE(e.full_name,c.employee_name,'?') title,c.status,'' branch_name,
       CAST(julianday(c.end_date)-julianday('now') AS INTEGER) days_remaining,'yellow' color
       FROM contracts c LEFT JOIN employees e ON c.employee_id=e.id
       WHERE c.status='Aktif' AND ${dateFilter('c.end_date')}`
    ).bind(...bind).all(),
  ]);

  const events = [
    ...(sched.results||[]), ...(issR.results||[]),
    ...(relR.results||[]),  ...(trainR.results||[]),
    ...(contrR.results||[]),
  ].sort((a,b)=>(a.event_date||'').localeCompare(b.event_date||''));

  return ok(events, 200, origin);
}

async function getActivityByPeriod(env, origin) {
  const rows = await env.DB.prepare(
    `SELECT activity_type,period,COUNT(*) total,
     SUM(CASE WHEN status='Done' THEN 1 ELSE 0 END) done
     FROM activity_schedule GROUP BY activity_type,period`
  ).all();
  return ok(rows.results||[], 200, origin);
}

async function getInspectionScores(request, env, origin) {
  const url  = new URL(request.url);
  const period = url.searchParams.get('period')||'';
  const where  = period ? 'WHERE r.period=?' : '';
  const bind   = period ? [period] : [];
  const rows   = await env.DB.prepare(
    `SELECT r.id,r.inspection_date,r.period,r.fc_score,r.spv_score,r.status,
     b.full_name branch_name
     FROM inspection_reports r LEFT JOIN branches b ON r.branch_id=b.id
     ${where} ORDER BY r.inspection_date ASC LIMIT 500`
  ).bind(...bind).all();
  return ok(rows.results||[], 200, origin);
}
