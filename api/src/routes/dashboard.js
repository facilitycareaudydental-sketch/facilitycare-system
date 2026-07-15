import { authenticate } from '../utils/auth.js';
import { ok, unauthorized } from '../utils/response.js';

export async function handleDashboard(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);

  const url = new URL(request.url);
  const path = url.pathname.replace('/api/dashboard', '');

  if (path === '/stats' || path === '') return getStats(env, origin);
  if (path === '/calendar') return getCalendarEvents(request, env, origin);
  if (path === '/contracts-expiring') return getContractsExpiring(env, origin);
  if (path === '/issues-summary') return getIssuesSummary(env, origin);
  if (path === '/activity-by-period') return getActivityByPeriod(env, origin);
  if (path === '/inspection-scores') return getInspectionScores(request, env, origin);

  return ok({}, 200, origin);
}

async function getStats(env, origin) {
  const today = new Date().toISOString().split('T')[0];

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

  // Recent issues
  const recentIssues = await env.DB.prepare(
    `SELECT i.*, b.full_name as branch_name FROM issues i
     LEFT JOIN branches b ON i.branch_id = b.id
     ORDER BY i.report_date DESC LIMIT 5`
  ).all();

  // Contracts expiring in 30 days
  const expiringList = await env.DB.prepare(
    `SELECT c.*, e.full_name as employee_name, b.full_name as branch_name,
     CAST(julianday(c.end_date) - julianday('now') AS INTEGER) as days_remaining
     FROM contracts c LEFT JOIN employees e ON c.employee_id = e.id
     LEFT JOIN branches b ON c.branch_id = b.id
     WHERE c.status = 'Aktif' AND c.end_date BETWEEN date('now') AND date('now', '+30 days')
     ORDER BY c.end_date ASC LIMIT 10`
  ).all();

  // Upcoming schedule
  const upcomingSchedule = await env.DB.prepare(
    `SELECT s.*, b.full_name as branch_name FROM activity_schedule s
     LEFT JOIN branches b ON s.branch_id = b.id
     WHERE s.status != 'Done' AND s.target_date >= date('now')
     ORDER BY s.target_date ASC LIMIT 7`
  ).all();

  return ok({
    stats: {
      total_employees: totalEmployees?.c || 0,
      active_contracts: activeContracts?.c || 0,
      expiring_contracts: expiringContracts?.c || 0,
      open_issues: openIssues?.c || 0,
      open_one_on_one: openOneOnOne?.c || 0,
      pending_schedule: pendingSchedule?.c || 0,
      pending_supply_requests: pendingSupply?.c || 0,
      total_branches: totalBranches?.c || 0,
      resolved_issues: doneIssues?.c || 0,
      total_training: totalTraining?.c || 0,
    },
    recent_issues: recentIssues.results,
    expiring_contracts: expiringList.results,
    upcoming_schedule: upcomingSchedule.results,
  }, 200, origin);
}

async function getCalendarEvents(request, env, origin) {
  const url = new URL(request.url);
  const month = url.searchParams.get('month'); // YYYY-MM
  const year = url.searchParams.get('year');

  let dateFilter = '';
  let values = [];

  if (month) {
    dateFilter = "strftime('%Y-%m', date_field) = ?";
    values = [month];
  } else if (year) {
    dateFilter = "strftime('%Y', date_field) = ?";
    values = [year];
  } else {
    // Default: current month
    const now = new Date();
    const currentMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    dateFilter = "strftime('%Y-%m', date_field) = ?";
    values = [currentMonth];
  }

  const buildQuery = (table, dateField, type, colorClass) => {
    const filter = dateFilter.replace('date_field', dateField);
    return env.DB.prepare(
      `SELECT id, ? as type, ${dateField} as event_date, ? as color_class,
       COALESCE(b.full_name, '') as branch_name
       FROM ${table} t LEFT JOIN branches b ON t.branch_id = b.id
       WHERE ${filter}`
    ).bind(type, colorClass, ...values).all();
  };

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
    ).bind(...values).all(),
    env.DB.prepare(
      `SELECT i.id, 'issue' as type, i.report_date as event_date, i.category as title,
       i.status, i.employee_name, b.full_name as branch_name, 'red' as color
       FROM issues i LEFT JOIN branches b ON i.branch_id = b.id
       WHERE ${dateFilter.replace('date_field', 'i.report_date')}`
    ).bind(...values).all(),
    env.DB.prepare(
      `SELECT r.id, 'reliever' as type, r.backup_date as event_date, r.reliever_name as title,
       r.status, r.reason, b.full_name as branch_name, 'teal' as color
       FROM relievers r LEFT JOIN branches b ON r.branch_id = b.id
       WHERE ${dateFilter.replace('date_field', 'r.backup_date')}`
    ).bind(...values).all(),
    env.DB.prepare(
      `SELECT t.id, 'training' as type, t.training_date as event_date, t.subject as title,
       'Done' as status, t.trainer, '' as branch_name, 'indigo' as color
       FROM training t
       WHERE ${dateFilter.replace('date_field', 't.training_date')}`
    ).bind(...values).all(),
    env.DB.prepare(
      `SELECT c.id, 'contract_expiry' as type, c.end_date as event_date,
       e.full_name as title, c.status, '' as branch_name,
       CAST(julianday(c.end_date) - julianday('now') AS INTEGER) as days_remaining,
       'yellow' as color
       FROM contracts c LEFT JOIN employees e ON c.employee_id = e.id
       WHERE c.status = 'Aktif' AND ${dateFilter.replace('date_field', 'c.end_date')}`
    ).bind(...values).all(),
  ]);

  const events = [
    ...scheduleRows.results,
    ...issueRows.results,
    ...relieverRows.results,
    ...trainingRows.results,
    ...contractRows.results,
  ].sort((a, b) => (a.event_date || '').localeCompare(b.event_date || ''));

  return ok(events, 200, origin);
}

async function getContractsExpiring(env, origin) {
  const rows = await env.DB.prepare(
    `SELECT c.*, e.full_name as employee_name, b.full_name as branch_name,
     CAST(julianday(c.end_date) - julianday('now') AS INTEGER) as days_remaining
     FROM contracts c LEFT JOIN employees e ON c.employee_id = e.id
     LEFT JOIN branches b ON c.branch_id = b.id
     WHERE c.status = 'Aktif' AND c.end_date >= date('now')
     ORDER BY c.end_date ASC LIMIT 50`
  ).all();
  return ok(rows.results, 200, origin);
}

async function getIssuesSummary(env, origin) {
  const byCategory = await env.DB.prepare(
    "SELECT category, COUNT(*) as count FROM issues GROUP BY category ORDER BY count DESC"
  ).all();
  const byStatus = await env.DB.prepare(
    "SELECT status, COUNT(*) as count FROM issues GROUP BY status"
  ).all();
  const byMonth = await env.DB.prepare(
    `SELECT strftime('%Y-%m', report_date) as month, COUNT(*) as count 
     FROM issues WHERE report_date >= date('now', '-12 months')
     GROUP BY month ORDER BY month`
  ).all();
  const byBranch = await env.DB.prepare(
    `SELECT b.full_name as branch_name, COUNT(*) as count FROM issues i
     LEFT JOIN branches b ON i.branch_id = b.id
     GROUP BY i.branch_id ORDER BY count DESC LIMIT 10`
  ).all();

  return ok({
    by_category: byCategory.results,
    by_status: byStatus.results,
    by_month: byMonth.results,
    by_branch: byBranch.results,
  }, 200, origin);
}

async function getActivityByPeriod(env, origin) {
  const byType = await env.DB.prepare(
    "SELECT activity_type, period, COUNT(*) as total, SUM(CASE WHEN status='Done' THEN 1 ELSE 0 END) as done FROM activity_schedule GROUP BY activity_type, period"
  ).all();
  return ok(byType.results, 200, origin);
}

async function getInspectionScores(request, env, origin) {
  const url = new URL(request.url);
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
