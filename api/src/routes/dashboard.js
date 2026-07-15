import { authenticate } from '../utils/auth.js';
import { ok, unauthorized } from '../utils/response.js';

export async function handleDashboard(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);

  const url  = new URL(request.url);
  const path = url.pathname.replace('/api/dashboard', '');

  if (path === '/stats' || path === '')  return getStats(env, origin);
  if (path === '/kpi')                   return getKPI(env, origin);
  if (path === '/calendar')              return getCalendarEvents(request, env, origin);
  if (path === '/contracts-expiring')    return getContractsExpiring(env, origin);
  if (path === '/issues-summary')        return getIssuesSummary(env, origin);
  if (path === '/issues-trend')          return getIssuesTrend(env, origin);
  if (path === '/contracts-chart')       return getContractsChart(env, origin);
  if (path === '/activity-by-period')    return getActivityByPeriod(env, origin);
  if (path === '/inspection-scores')     return getInspectionScores(request, env, origin);
  if (path === '/inspection-bar')        return getInspectionBar(env, origin);
  if (path === '/activity-log')          return getActivityLog(env, origin);

  return ok({}, 200, origin);
}

// ─── Helpers ────────────────────────────────────────────────────────────────
const todayStr = () => new Date().toISOString().split('T')[0];

// ─── KPI — Stats + Previous Month for Trend ─────────────────────────────────
async function getKPI(env, origin) {
  const now        = new Date();
  const y          = now.getFullYear();
  const m          = String(now.getMonth() + 1).padStart(2,'0');
  const prevM      = now.getMonth() === 0 ? '12' : String(now.getMonth()).padStart(2,'0');
  const prevY      = now.getMonth() === 0 ? y - 1 : y;
  const curMonth   = `${y}-${m}`;
  const prevMonth  = `${prevY}-${prevM}`;

  const [
    empActive, empPrev,
    contractActive, contractPrev,
    contractExp30,
    issuesOpen, issuesPrev,
    oo1Open, oo1Prev,
    schedPending,
    supplyPending,
    totalBranches,
    trainingThisMonth,
    inspCount,
    cleanCount,
    fogCount,
  ] = await Promise.all([
    env.DB.prepare("SELECT COUNT(*) c FROM employees WHERE status='Aktif'").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM employees WHERE status='Aktif' AND strftime('%Y-%m',created_at)=?").bind(prevMonth).first(),
    env.DB.prepare("SELECT COUNT(*) c FROM contracts WHERE status='Aktif' AND end_date>=date('now')").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM contracts WHERE status='Aktif' AND strftime('%Y-%m',created_at)=?").bind(prevMonth).first(),
    env.DB.prepare("SELECT COUNT(*) c FROM contracts WHERE status='Aktif' AND end_date BETWEEN date('now') AND date('now','+30 days')").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM issues WHERE status!='Done'").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM issues WHERE status!='Done' AND strftime('%Y-%m',report_date)=?").bind(prevMonth).first(),
    env.DB.prepare("SELECT COUNT(*) c FROM one_on_one WHERE status!='Done'").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM one_on_one WHERE status!='Done' AND strftime('%Y-%m',meeting_date)=?").bind(prevMonth).first(),
    env.DB.prepare("SELECT COUNT(*) c FROM activity_schedule WHERE status='Pending'").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM supply_requests WHERE status='Pending'").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM branches WHERE is_active=1").first(),
    env.DB.prepare("SELECT COUNT(*) c FROM training WHERE strftime('%Y-%m',training_date)=?").bind(curMonth).first(),
    env.DB.prepare("SELECT COUNT(*) c FROM inspection_reports WHERE strftime('%Y-%m',inspection_date)=?").bind(curMonth).first(),
    env.DB.prepare("SELECT COUNT(*) c FROM cleaning_reports WHERE strftime('%Y-%m',activity_date)=?").bind(curMonth).first(),
    env.DB.prepare("SELECT COUNT(*) c FROM fogging_reports WHERE strftime('%Y-%m',activity_date)=?").bind(curMonth).first(),
  ]);

  return ok({
    employees:      { current: empActive?.c||0,        prev: empPrev?.c||0 },
    contracts:      { current: contractActive?.c||0,   prev: contractPrev?.c||0 },
    expiring30:     { current: contractExp30?.c||0 },
    issues:         { current: issuesOpen?.c||0,       prev: issuesPrev?.c||0 },
    one_on_one:     { current: oo1Open?.c||0,          prev: oo1Prev?.c||0 },
    schedule:       { current: schedPending?.c||0 },
    supply:         { current: supplyPending?.c||0 },
    branches:       { current: totalBranches?.c||0 },
    training_month: { current: trainingThisMonth?.c||0 },
    inspection_month:{ current: inspCount?.c||0 },
    cleaning_month: { current: cleanCount?.c||0 },
    fogging_month:  { current: fogCount?.c||0 },
  }, 200, origin);
}

// ─── Legacy Stats (kept for backward compat) ────────────────────────────────
async function getStats(env, origin) {
  const [
    totalEmployees, activeContracts, expiringContracts, openIssues,
    openOneOnOne, pendingSchedule, pendingSupply, totalBranches,
    doneIssues, totalTraining
  ] = await Promise.all([
    env.DB.prepare("SELECT COUNT(*) as c FROM employees WHERE status = 'Aktif'").first(),
    env.DB.prepare("SELECT COUNT(*) as c FROM contracts WHERE status = 'Aktif' AND end_date >= date('now')").first(),
    env.DB.prepare("SELECT COUNT(*) as c FROM contracts WHERE status = 'Aktif' AND end_date BETWEEN date('now') AND date('now', '+30 days')").first(),
    env.DB.prepare("SELECT COUNT(*) as c FROM issues WHERE status != 'Done'").first(),
    env.DB.prepare("SELECT COUNT(*) as c FROM one_on_one WHERE status != 'Done'").first(),
    env.DB.prepare("SELECT COUNT(*) as c FROM activity_schedule WHERE status = 'Pending'").first(),
    env.DB.prepare("SELECT COUNT(*) as c FROM supply_requests WHERE status = 'Pending'").first(),
    env.DB.prepare("SELECT COUNT(*) as c FROM branches WHERE is_active = 1").first(),
    env.DB.prepare("SELECT COUNT(*) as c FROM issues WHERE status = 'Done'").first(),
    env.DB.prepare("SELECT COUNT(*) as c FROM training").first(),
  ]);

  const recentIssues = await env.DB.prepare(
    `SELECT i.*, b.full_name as branch_name FROM issues i
     LEFT JOIN branches b ON i.branch_id = b.id
     ORDER BY i.report_date DESC LIMIT 5`
  ).all();

  const expiringList = await env.DB.prepare(
    `SELECT c.*, COALESCE(e.full_name, c.employee_name) as emp_name, b.full_name as branch_name,
     CAST(julianday(c.end_date) - julianday('now') AS INTEGER) as days_remaining
     FROM contracts c LEFT JOIN employees e ON c.employee_id = e.id
     LEFT JOIN branches b ON c.branch_id = b.id
     WHERE c.status = 'Aktif' AND c.end_date BETWEEN date('now') AND date('now', '+30 days')
     ORDER BY c.end_date ASC LIMIT 10`
  ).all();

  const upcomingSchedule = await env.DB.prepare(
    `SELECT s.*, b.full_name as branch_name FROM activity_schedule s
     LEFT JOIN branches b ON s.branch_id = b.id
     WHERE s.status != 'Done' AND s.target_date >= date('now')
     ORDER BY s.target_date ASC LIMIT 7`
  ).all();

  return ok({
    stats: {
      total_employees:       totalEmployees?.c || 0,
      active_contracts:      activeContracts?.c || 0,
      expiring_contracts:    expiringContracts?.c || 0,
      open_issues:           openIssues?.c || 0,
      open_one_on_one:       openOneOnOne?.c || 0,
      pending_schedule:      pendingSchedule?.c || 0,
      pending_supply_requests: pendingSupply?.c || 0,
      total_branches:        totalBranches?.c || 0,
      resolved_issues:       doneIssues?.c || 0,
      total_training:        totalTraining?.c || 0,
    },
    recent_issues:      recentIssues.results,
    expiring_contracts: expiringList.results,
    upcoming_schedule:  upcomingSchedule.results,
  }, 200, origin);
}

// ─── Issues Trend 12 Months ──────────────────────────────────────────────────
async function getIssuesTrend(env, origin) {
  const [open, closed] = await Promise.all([
    env.DB.prepare(
      `SELECT strftime('%Y-%m', report_date) as month, COUNT(*) as count
       FROM issues WHERE report_date >= date('now','-12 months') AND status!='Done'
       GROUP BY month ORDER BY month`
    ).all(),
    env.DB.prepare(
      `SELECT strftime('%Y-%m', completion_date) as month, COUNT(*) as count
       FROM issues WHERE completion_date >= date('now','-12 months') AND status='Done'
       GROUP BY month ORDER BY month`
    ).all(),
  ]);

  // Build 12-month labels
  const labels = [];
  const now = new Date();
  for (let i = 11; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    labels.push(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`);
  }

  const toMap = (rows) => Object.fromEntries((rows||[]).map(r => [r.month, r.count]));
  const openMap   = toMap(open.results);
  const closedMap = toMap(closed.results);

  return ok({
    labels,
    open:   labels.map(l => openMap[l]   || 0),
    closed: labels.map(l => closedMap[l] || 0),
  }, 200, origin);
}

// ─── Contracts Expiring — Chart by Month ────────────────────────────────────
async function getContractsChart(env, origin) {
  const rows = await env.DB.prepare(
    `SELECT strftime('%Y-%m', end_date) as month, COUNT(*) as count
     FROM contracts WHERE status='Aktif' AND end_date >= date('now')
     AND end_date <= date('now','+6 months')
     GROUP BY month ORDER BY month`
  ).all();

  const labels = [];
  const now = new Date();
  for (let i = 0; i < 6; i++) {
    const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
    labels.push(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`);
  }
  const map = Object.fromEntries((rows.results||[]).map(r => [r.month, r.count]));

  return ok({
    labels,
    counts: labels.map(l => map[l] || 0),
  }, 200, origin);
}

// ─── Contracts Expiring List ─────────────────────────────────────────────────
async function getContractsExpiring(env, origin) {
  const rows = await env.DB.prepare(
    `SELECT c.*, COALESCE(e.full_name, c.employee_name) as emp_name,
     b.full_name as branch_name,
     CAST(julianday(c.end_date) - julianday('now') AS INTEGER) as days_remaining
     FROM contracts c LEFT JOIN employees e ON c.employee_id = e.id
     LEFT JOIN branches b ON c.branch_id = b.id
     WHERE c.status = 'Aktif' AND c.end_date >= date('now')
     ORDER BY c.end_date ASC LIMIT 50`
  ).all();
  return ok(rows.results, 200, origin);
}

// ─── Issues Summary ──────────────────────────────────────────────────────────
async function getIssuesSummary(env, origin) {
  const [byCategory, byStatus, byMonth, byBranch] = await Promise.all([
    env.DB.prepare("SELECT category, COUNT(*) as count FROM issues GROUP BY category ORDER BY count DESC").all(),
    env.DB.prepare("SELECT status, COUNT(*) as count FROM issues GROUP BY status").all(),
    env.DB.prepare(`SELECT strftime('%Y-%m', report_date) as month, COUNT(*) as count
       FROM issues WHERE report_date >= date('now', '-12 months')
       GROUP BY month ORDER BY month`).all(),
    env.DB.prepare(`SELECT b.full_name as branch_name, COUNT(*) as count FROM issues i
       LEFT JOIN branches b ON i.branch_id = b.id
       GROUP BY i.branch_id ORDER BY count DESC LIMIT 10`).all(),
  ]);
  return ok({
    by_category: byCategory.results,
    by_status:   byStatus.results,
    by_month:    byMonth.results,
    by_branch:   byBranch.results,
  }, 200, origin);
}

// ─── Inspection Bar Chart ────────────────────────────────────────────────────
async function getInspectionBar(env, origin) {
  const rows = await env.DB.prepare(
    `SELECT b.full_name as branch_name,
     AVG(r.fc_score) as avg_fc, AVG(r.spv_score) as avg_spv, COUNT(*) as total
     FROM inspection_reports r
     LEFT JOIN branches b ON r.branch_id = b.id
     WHERE r.inspection_date >= date('now','-6 months')
     GROUP BY r.branch_id ORDER BY avg_fc DESC LIMIT 15`
  ).all();

  const labels   = (rows.results||[]).map(r => r.branch_name || 'N/A');
  const fcScores = (rows.results||[]).map(r => r.avg_fc  ? Math.round(r.avg_fc*10)/10  : 0);
  const spvScores= (rows.results||[]).map(r => r.avg_spv ? Math.round(r.avg_spv*10)/10 : 0);

  return ok({ labels, fc: fcScores, spv: spvScores }, 200, origin);
}

// ─── Activity Log ────────────────────────────────────────────────────────────
async function getActivityLog(env, origin) {
  // Collect recent items from all modules — union of recent records
  const [issues, contracts, employees, one_on_one, training, supply] = await Promise.all([
    env.DB.prepare(`SELECT 'issue' as type, complaint as label, created_at, b.full_name as branch
       FROM issues i LEFT JOIN branches b ON i.branch_id = b.id
       ORDER BY i.created_at DESC LIMIT 5`).all(),
    env.DB.prepare(`SELECT 'contract' as type, COALESCE(e.full_name,c.employee_name,'?') as label, c.created_at, b.full_name as branch
       FROM contracts c LEFT JOIN employees e ON c.employee_id=e.id
       LEFT JOIN branches b ON c.branch_id=b.id
       ORDER BY c.created_at DESC LIMIT 5`).all(),
    env.DB.prepare(`SELECT 'employee' as type, full_name as label, created_at, '' as branch
       FROM employees ORDER BY created_at DESC LIMIT 5`).all(),
    env.DB.prepare(`SELECT 'one_on_one' as type, employee_name as label, created_at, b.full_name as branch
       FROM one_on_one o LEFT JOIN branches b ON o.branch_id=b.id
       ORDER BY o.created_at DESC LIMIT 3`).all(),
    env.DB.prepare(`SELECT 'training' as type, subject as label, created_at, '' as branch
       FROM training ORDER BY created_at DESC LIMIT 3`).all(),
    env.DB.prepare(`SELECT 'supply' as type, submitter_name as label, created_at, branch_name as branch
       FROM supply_requests ORDER BY created_at DESC LIMIT 3`).all(),
  ]);

  const all = [
    ...(issues.results||[]),
    ...(contracts.results||[]),
    ...(employees.results||[]),
    ...(one_on_one.results||[]),
    ...(training.results||[]),
    ...(supply.results||[]),
  ].sort((a,b) => (b.created_at||'').localeCompare(a.created_at||'')).slice(0, 20);

  return ok(all, 200, origin);
}

// ─── Calendar Events ─────────────────────────────────────────────────────────
async function getCalendarEvents(request, env, origin) {
  const url  = new URL(request.url);
  const month = url.searchParams.get('month');
  const year  = url.searchParams.get('year');
  const now   = new Date();
  const curMonth = month || `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
  const dateFilter = "strftime('%Y-%m', date_field) = ?";
  const vals = [curMonth];

  const [scheduleRows, issueRows, relieverRows, trainingRows, contractRows] = await Promise.all([
    env.DB.prepare(
      `SELECT s.id, 'schedule' as type, s.target_date as event_date, s.activity_type as title,
       s.status, s.pic, b.full_name as branch_name,
       CASE s.activity_type
         WHEN 'Inspeksi Hygiene & Aset Bangunan' THEN 'blue'
         WHEN 'General Cleaning' THEN 'green'
         WHEN 'Deep Cleaning' THEN 'purple'
         WHEN 'Fogging' THEN 'orange'
         ELSE 'gray'
       END as color
       FROM activity_schedule s LEFT JOIN branches b ON s.branch_id = b.id
       WHERE ${dateFilter.replace('date_field', 's.target_date')}`
    ).bind(...vals).all(),
    env.DB.prepare(
      `SELECT i.id, 'issue' as type, i.report_date as event_date, i.category as title,
       i.status, i.employee_name, b.full_name as branch_name, 'red' as color
       FROM issues i LEFT JOIN branches b ON i.branch_id = b.id
       WHERE ${dateFilter.replace('date_field', 'i.report_date')}`
    ).bind(...vals).all(),
    env.DB.prepare(
      `SELECT r.id, 'reliever' as type, r.backup_date as event_date, r.reliever_name as title,
       r.status, r.reason, b.full_name as branch_name, 'teal' as color
       FROM relievers r LEFT JOIN branches b ON r.branch_id = b.id
       WHERE ${dateFilter.replace('date_field', 'r.backup_date')}`
    ).bind(...vals).all(),
    env.DB.prepare(
      `SELECT t.id, 'training' as type, t.training_date as event_date, t.subject as title,
       'Done' as status, t.trainer, '' as branch_name, 'indigo' as color
       FROM training t
       WHERE ${dateFilter.replace('date_field', 't.training_date')}`
    ).bind(...vals).all(),
    env.DB.prepare(
      `SELECT c.id, 'contract_expiry' as type, c.end_date as event_date,
       COALESCE(e.full_name, c.employee_name) as title, c.status, '' as branch_name,
       CAST(julianday(c.end_date) - julianday('now') AS INTEGER) as days_remaining,
       'yellow' as color
       FROM contracts c LEFT JOIN employees e ON c.employee_id = e.id
       WHERE c.status = 'Aktif' AND ${dateFilter.replace('date_field', 'c.end_date')}`
    ).bind(...vals).all(),
  ]);

  const events = [
    ...scheduleRows.results, ...issueRows.results, ...relieverRows.results,
    ...trainingRows.results, ...contractRows.results,
  ].sort((a, b) => (a.event_date||'').localeCompare(b.event_date||''));

  return ok(events, 200, origin);
}

async function getActivityByPeriod(env, origin) {
  const byType = await env.DB.prepare(
    "SELECT activity_type, period, COUNT(*) as total, SUM(CASE WHEN status='Done' THEN 1 ELSE 0 END) as done FROM activity_schedule GROUP BY activity_type, period"
  ).all();
  return ok(byType.results, 200, origin);
}

async function getInspectionScores(request, env, origin) {
  const url  = new URL(request.url);
  const period = url.searchParams.get('period') || '';
  let conditions = []; let values = [];
  if (period) { conditions.push('r.period = ?'); values.push(period); }
  const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';
  const rows = await env.DB.prepare(
    `SELECT r.*, b.full_name as branch_name FROM inspection_reports r
     LEFT JOIN branches b ON r.branch_id = b.id
     ${where} ORDER BY r.inspection_date ASC`
  ).bind(...values).all();
  return ok(rows.results, 200, origin);
}
