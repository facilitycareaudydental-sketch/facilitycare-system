import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden } from '../utils/response.js';

export async function handleImport(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'employees', 'write')) return forbidden(origin);

  const url = new URL(request.url);
  const path = url.pathname.replace('/api/import', '');

  if (path === '/backup' && request.method === 'GET') {
    return exportBackup(env, origin);
  }

  if (path === '/sync-calendar' && request.method === 'POST') {
    return triggerCalendarSync(env, origin);
  }

  // GET /api/import/counts — return row counts for all tables
  if (path === '/counts' && request.method === 'GET') {
    const tables = [
      'validation_options','employees','contracts','issues','one_on_one',
      'training','relievers','activity_schedule','inspection_reports',
      'cleaning_reports','fogging_reports','basecamp_reports','sop',
      'master_checklist','master_forms','supply_requests'
    ];
    const counts = {};
    for (const t of tables) {
      try {
        const r = await env.DB.prepare(`SELECT COUNT(*) as cnt FROM ${t}`).first();
        counts[t] = r.cnt;
      } catch (e) {
        counts[t] = -1;
      }
    }
    return ok(counts, 200, origin);
  }

  if (request.method !== 'POST') return error('Method not allowed', 405, origin);


  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }

  const { rows = [], onDuplicate = 'skip' } = body;
  if (!Array.isArray(rows)) return error('rows must be an array', 400, origin);

  try {
    switch (path) {
      case '/employees':    return importEmployees(rows, onDuplicate, env, origin);
      case '/contracts':    return importContracts(rows, onDuplicate, env, origin);
      case '/relievers':    return importRelievers(rows, onDuplicate, env, origin);
      case '/schedule':     return importSchedule(rows, onDuplicate, env, origin);
      case '/issues':       return importIssues(rows, onDuplicate, env, origin);
      case '/one_on_one':   return importOneOnOne(rows, onDuplicate, env, origin);
      case '/training':     return importTraining(rows, onDuplicate, env, origin);
      case '/checklist':    return importChecklist(rows, onDuplicate, env, origin);
      case '/forms':        return importForms(rows, onDuplicate, env, origin);
      case '/sop':          return importSop(rows, onDuplicate, env, origin);
      case '/inspection':   return importInspection(rows, onDuplicate, env, origin);
      case '/cleaning':     return importCleaning(rows, onDuplicate, env, origin);
      case '/fogging':      return importFogging(rows, onDuplicate, env, origin);
      case '/basecamp':     return importBasecamp(rows, onDuplicate, env, origin);
      case '/supply':       return importSupply(rows, onDuplicate, env, origin);
      case '/validation':   return importValidation(rows, onDuplicate, env, origin);
      default:              return error('Unknown import module', 404, origin);
    }
  } catch (err) {
    return error('Import error: ' + err.message, 500, origin);
  }
}

// ─── Helpers ─────────────────────────────────────────────────────────────────
function safeStr(v)  { return v !== undefined && v !== null && String(v).trim() !== '' ? String(v).trim() : null; }
function safeDate(v) {
  if (!v) return null;
  if (typeof v === 'string' && v.match(/^\d{4}-\d{2}-\d{2}/)) return v.slice(0,10);
  const excelDays = Number(v);
  if (!isNaN(excelDays) && excelDays > 20000 && excelDays < 60000) {
    const d = new Date((excelDays - 25569) * 86400 * 1000);
    if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  }
  return safeStr(v);
}
function today() { return new Date().toISOString().slice(0, 10); }

function makeBranchMatcher(branches) {
  return (str) => {
    if (!str) return null;
    const s = String(str).toLowerCase().trim();
    const exact = branches.find(r =>
      r.full_name?.toLowerCase() === s ||
      r.code?.toLowerCase() === s ||
      r.name?.toLowerCase() === s
    );
    if (exact) return exact.id;
    const partial = branches.find(r =>
      r.full_name?.toLowerCase().includes(s) ||
      s.includes(r.code?.toLowerCase() || '~~~')
    );
    return partial ? partial.id : null;
  };
}

async function batchInsert(db, stmts) {
  if (stmts.length === 0) return 0;
  const chunkSize = 100;
  let total = 0;
  for (let i = 0; i < stmts.length; i += chunkSize) {
    await db.batch(stmts.slice(i, i + chunkSize));
    total += stmts.slice(i, i + chunkSize).length;
  }
  return total;
}

// ─── Validation Options ──────────────────────────────────────────────────────
async function importValidation(rows, onDuplicate, env, origin) {
  const branchSet = new Set();
  const picSet = new Set();
  const kegiatanSet = new Set();
  const quartalSet = new Set();
  const pkwtSet = new Set();

  for (const row of rows) {
    if (row.cabang) {
      const val = String(row.cabang).trim();
      if (val) branchSet.add(val);
    }
    if (row.pic) {
      const val = String(row.pic).trim();
      if (val) picSet.add(val);
    }
    if (row.kegiatan) {
      const val = String(row.kegiatan).trim();
      if (val) kegiatanSet.add(val);
    }
    if (row.quartal) {
      const val = String(row.quartal).trim();
      if (val) quartalSet.add(val);
    }
    if (row.masa_pkwt) {
      const val = String(row.masa_pkwt).trim();
      if (val) pkwtSet.add(val);
    }
  }

  const db = env.DB;
  const stmts = [];

  // 1. Branches Sync
  const activeBranchCodes = [];
  for (const full_name of branchSet) {
    let code = '';
    let name = '';
    const dotIdx = full_name.indexOf('.');
    if (dotIdx > 0) {
      code = full_name.slice(0, dotIdx).trim();
      name = full_name.slice(dotIdx + 1).trim();
    } else {
      code = full_name;
      name = full_name;
    }
    activeBranchCodes.push(code);

    stmts.push(db.prepare(
      `INSERT INTO branches (code, name, full_name, is_active) VALUES (?, ?, ?, 1)
       ON CONFLICT(code) DO UPDATE SET name=excluded.name, full_name=excluded.full_name, is_active=1`
    ).bind(code, name, full_name));
  }

  if (activeBranchCodes.length > 0) {
    const placeholders = activeBranchCodes.map(() => '?').join(',');
    stmts.push(db.prepare(
      `UPDATE branches SET is_active = 0 WHERE code NOT IN (${placeholders})`
    ).bind(...activeBranchCodes));
  }

  // 2. PIC list sync
  const activePics = Array.from(picSet);
  for (const pic of activePics) {
    stmts.push(db.prepare(
      `INSERT INTO pic_list (name, is_active) VALUES (?, 1)
       ON CONFLICT(name) DO UPDATE SET is_active=1`
    ).bind(pic));
  }
  if (activePics.length > 0) {
    const placeholders = activePics.map(() => '?').join(',');
    stmts.push(db.prepare(
      `UPDATE pic_list SET is_active = 0 WHERE name NOT IN (${placeholders})`
    ).bind(...activePics));
  }

  // 3. Validation Options sync
  stmts.push(db.prepare(`DELETE FROM validation_options`));
  for (const pic of activePics) {
    stmts.push(db.prepare(`INSERT OR IGNORE INTO validation_options (category, value) VALUES ('pic', ?)`).bind(pic));
  }
  for (const keg of kegiatanSet) {
    stmts.push(db.prepare(`INSERT OR IGNORE INTO validation_options (category, value) VALUES ('activity', ?)`).bind(keg));
  }
  for (const q of quartalSet) {
    stmts.push(db.prepare(`INSERT OR IGNORE INTO validation_options (category, value) VALUES ('quarter', ?)`).bind(q));
  }
  for (const pkwt of pkwtSet) {
    stmts.push(db.prepare(`INSERT OR IGNORE INTO validation_options (category, value) VALUES ('pkwt', ?)`).bind(pkwt));
  }

  await batchInsert(db, stmts);
  return ok({ inserted: branchSet.size + picSet.size + kegiatanSet.size + quartalSet.size + pkwtSet.size, skipped: 0 }, 200, origin);
}

// ─── Employees ───────────────────────────────────────────────────────────────
async function importEmployees(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const existing = await env.DB.prepare('SELECT id, full_name FROM employees').all();
  const existingMap = new Map();
  (existing.results || []).forEach(e => {
    existingMap.set(e.full_name.toLowerCase().trim(), e.id);
  });

  const stmts = [];
  const importedNames = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;

  for (const row of rows) {
    const full_name = safeStr(row.full_name);
    if (!full_name) { skipped++; continue; }
    importedNames.push(full_name);

    const key = full_name.toLowerCase().trim();
    const branch_id = matchBranch(row.branch_name);
    const division = safeStr(row.division) || 'FACILITY CARE';
    const phone = safeStr(row.phone);
    const join_date = safeDate(row.join_date);
    const status = safeStr(row.status) || 'Aktif';
    const notes = safeStr(row.notes);

    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === 'update') {
        stmts.push(env.DB.prepare(
          `UPDATE employees SET branch_id = ?, division = ?, phone = ?, join_date = ?, status = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(branch_id, division, phone, join_date, status, notes, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env.DB.prepare(
        `INSERT INTO employees (full_name, branch_id, division, phone, join_date, status, notes) VALUES (?, ?, ?, ?, ?, ?, ?)`
      ).bind(full_name, branch_id, division, phone, join_date, status, notes));
      inserted++;
    }
  }

  await batchInsert(env.DB, stmts);

  return ok({ inserted, skipped, updated }, 200, origin);
}

// ─── Contracts ────────────────────────────────────────────────────────────────
async function importContracts(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const eRows = await env.DB.prepare('SELECT id, full_name FROM employees').all();
  const matchEmployee = (str) => {
    if (!str) return null;
    const s = str.toLowerCase().trim();
    const emp = eRows.results.find(r => r.full_name.toLowerCase().trim() === s);
    return emp ? emp.id : null;
  };

  const existing = await env.DB.prepare('SELECT id, employee_name, start_date FROM contracts').all();
  const existingMap = new Map();
  (existing.results || []).forEach(c => {
    if (c.employee_name && c.start_date) {
      existingMap.set(c.employee_name.toLowerCase().trim() + '_' + c.start_date, c.id);
    }
  });

  const stmts = [];
  const importedKeys = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;

  for (const row of rows) {
    const employee_name = safeStr(row.employee_name);
    if (!employee_name) { skipped++; continue; }

    const start_date = safeDate(row.start_date);
    const end_date = safeDate(row.end_date);
    const key = employee_name.toLowerCase().trim() + '_' + start_date;
    importedKeys.push(key);

    const employee_id = matchEmployee(employee_name);
    const branch_id = matchBranch(row.branch_name);
    const division = safeStr(row.division) || 'FACILITY CARE';
    const contract_type = safeStr(row.contract_type);
    const pkwt_number = safeStr(row.pkwt_number);
    const status = safeStr(row.status) || 'Aktif';
    const notes = safeStr(row.notes);

    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === 'update') {
        stmts.push(env.DB.prepare(
          `UPDATE contracts SET employee_id = ?, branch_id = ?, division = ?, end_date = ?, contract_type = ?, pkwt_number = ?, status = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(employee_id, branch_id, division, end_date, contract_type, pkwt_number, status, notes, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env.DB.prepare(
        `INSERT INTO contracts (employee_id, employee_name, branch_id, division, start_date, end_date, contract_type, pkwt_number, status, notes)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(employee_id, employee_name, branch_id, division, start_date, end_date, contract_type, pkwt_number, status, notes));
      inserted++;
    }
  }

  await batchInsert(env.DB, stmts);

  return ok({ inserted, skipped, updated }, 200, origin);
}

// ─── Relievers ────────────────────────────────────────────────────────────────
async function importRelievers(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const existing = await env.DB.prepare('SELECT id, reliever_name, backup_date FROM relievers').all();
  const existingMap = new Map();
  (existing.results || []).forEach(r => {
    if (r.reliever_name && r.backup_date) {
      existingMap.set(r.reliever_name.toLowerCase().trim() + '_' + r.backup_date, r.id);
    }
  });

  const stmts = [];
  const importedKeys = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;

  for (const row of rows) {
    const reliever_name = safeStr(row.reliever_name) || '-';

    const backup_date = safeDate(row.backup_date) || today();
    const key = reliever_name.toLowerCase().trim() + '_' + backup_date;
    importedKeys.push(key);

    const branch_id = matchBranch(row.branch_name);
    const original_fc_name = safeStr(row.original_fc_name);
    const period = safeStr(row.period);
    const completion_date = safeDate(row.completion_date);
    const reason = safeStr(row.reason);
    const shift = safeStr(row.shift);
    const status = safeStr(row.status) || 'Pending';

    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === 'update') {
        stmts.push(env.DB.prepare(
          `UPDATE relievers SET branch_id = ?, original_fc_name = ?, period = ?, completion_date = ?, reason = ?, shift = ?, status = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(branch_id, original_fc_name, period, completion_date, reason, shift, status, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env.DB.prepare(
        `INSERT INTO relievers (branch_id, original_fc_name, period, reliever_name, backup_date, completion_date, reason, shift, status)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(branch_id, original_fc_name, period, reliever_name, backup_date, completion_date, reason, shift, status));
      inserted++;
    }
  }

  await batchInsert(env.DB, stmts);

  return ok({ inserted, skipped, updated }, 200, origin);
}

// ─── Schedule ─────────────────────────────────────────────────────────────────
async function importSchedule(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const existing = await env.DB.prepare('SELECT id, activity_type, target_date, branch_id FROM activity_schedule').all();
  const existingMap = new Map();
  (existing.results || []).forEach(s => {
    if (s.activity_type && s.target_date) {
      existingMap.set(s.activity_type.toLowerCase().trim() + '_' + s.target_date + '_' + s.branch_id, s.id);
    }
  });

  const stmts = [];
  const importedKeys = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;

  for (const row of rows) {
    const activity_type = safeStr(row.activity_type);
    if (!activity_type) { skipped++; continue; }

    const target_date = safeDate(row.target_date);
    const branch_id = matchBranch(row.branch_name);
    const key = activity_type.toLowerCase().trim() + '_' + target_date + '_' + branch_id;
    importedKeys.push(key);

    const period = safeStr(row.period);
    const pic = safeStr(row.pic);
    const opening_date = safeDate(row.opening_date);
    const completion_date = safeDate(row.completion_date);
    const status = safeStr(row.status) || 'Pending';
    const notes = safeStr(row.notes);

    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === 'update') {
        stmts.push(env.DB.prepare(
          `UPDATE activity_schedule SET period = ?, pic = ?, opening_date = ?, completion_date = ?, status = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(period, pic, opening_date, completion_date, status, notes, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env.DB.prepare(
        `INSERT INTO activity_schedule (branch_id, activity_type, period, pic, opening_date, target_date, completion_date, status, notes)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(branch_id, activity_type, period, pic, opening_date, target_date, completion_date, status, notes));
      inserted++;
    }
  }

  await batchInsert(env.DB, stmts);

  return ok({ inserted, skipped, updated }, 200, origin);
}

// ─── Issues ───────────────────────────────────────────────────────────────────
async function importIssues(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const existing = await env.DB.prepare('SELECT id, complaint, report_date, branch_id FROM issues').all();
  const existingMap = new Map();
  (existing.results || []).forEach(i => {
    if (i.complaint && i.report_date) {
      existingMap.set(i.complaint.toLowerCase().trim() + '_' + i.report_date + '_' + i.branch_id, i.id);
    }
  });

  const stmts = [];
  const importedKeys = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;

  for (const row of rows) {
    const complaint = safeStr(row.complaint);
    if (!complaint) { skipped++; continue; }

    const report_date = safeDate(row.report_date) || today();
    const branch_id = matchBranch(row.branch_name);
    const key = complaint.toLowerCase().trim() + '_' + report_date + '_' + branch_id;
    importedKeys.push(key);

    const category = safeStr(row.category) || 'Lainnya';
    const source = safeStr(row.source);
    const employee_name = safeStr(row.employee_name);
    const fc_specialist = safeStr(row.fc_specialist);
    const solution = safeStr(row.solution);
    const status = safeStr(row.status) || 'Open';
    const completion_date = safeDate(row.completion_date);

    let day_count = null;
    if (completion_date && report_date) {
      const d1 = new Date(report_date), d2 = new Date(completion_date);
      if (!isNaN(d1) && !isNaN(d2)) day_count = Math.floor((d2 - d1) / 86400000);
    }

    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === 'update') {
        stmts.push(env.DB.prepare(
          `UPDATE issues SET category = ?, source = ?, employee_name = ?, fc_specialist = ?, solution = ?, status = ?, completion_date = ?, day_count = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(category, source, employee_name, fc_specialist, solution, status, completion_date, day_count, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env.DB.prepare(
        `INSERT INTO issues (report_date, branch_id, category, source, complaint, employee_name, fc_specialist, solution, status, completion_date, day_count)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(report_date, branch_id, category, source, complaint, employee_name, fc_specialist, solution, status, completion_date, day_count));
      inserted++;
    }
  }

  await batchInsert(env.DB, stmts);

  return ok({ inserted, skipped, updated }, 200, origin);
}

// ─── One on One ───────────────────────────────────────────────────────────────
async function importOneOnOne(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const existing = await env.DB.prepare('SELECT id, employee_name, meeting_date FROM one_on_one').all();
  const existingMap = new Map();
  (existing.results || []).forEach(o => {
    if (o.employee_name && o.meeting_date) {
      existingMap.set(o.employee_name.toLowerCase().trim() + '_' + o.meeting_date, o.id);
    }
  });

  const stmts = [];
  const importedKeys = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;

  for (const row of rows) {
    const employee_name = safeStr(row.employee_name) || '-';

    const meeting_date = safeDate(row.meeting_date) || today();
    const key = employee_name.toLowerCase().trim() + '_' + meeting_date;
    importedKeys.push(key);

    const branch_id = matchBranch(row.branch_name);
    const pic = safeStr(row.pic);
    const problem = safeStr(row.problem) || '-';
    const solution = safeStr(row.solution);
    const status = safeStr(row.status) || 'Open';
    const completion_date = safeDate(row.completion_date);
    const document_link = safeStr(row.document_link);

    let day_count = null;
    if (completion_date && meeting_date) {
      const d1 = new Date(meeting_date), d2 = new Date(completion_date);
      if (!isNaN(d1) && !isNaN(d2)) day_count = Math.floor((d2 - d1) / 86400000);
    }

    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === 'update') {
        stmts.push(env.DB.prepare(
          `UPDATE one_on_one SET branch_id = ?, pic = ?, problem = ?, solution = ?, status = ?, completion_date = ?, day_count = ?, document_link = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(branch_id, pic, problem, solution, status, completion_date, day_count, document_link, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env.DB.prepare(
        `INSERT INTO one_on_one (meeting_date, branch_id, employee_name, pic, problem, solution, status, completion_date, day_count, document_link)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(meeting_date, branch_id, employee_name, pic, problem, solution, status, completion_date, day_count, document_link));
      inserted++;
    }
  }

  await batchInsert(env.DB, stmts);

  return ok({ inserted, skipped, updated }, 200, origin);
}

// ─── Training ─────────────────────────────────────────────────────────────────
async function importTraining(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const existing = await env.DB.prepare('SELECT id, subject, training_date, batch FROM training').all();
  const existingMap = new Map();
  (existing.results || []).forEach(t => {
    if (t.subject && t.training_date) {
      existingMap.set(t.subject.toLowerCase().trim() + '_' + t.training_date + '_' + (t.batch || '').toLowerCase().trim(), t.id);
    }
  });

  const stmts = [];
  const importedKeys = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;

  for (const row of rows) {
    const subject = safeStr(row.subject);
    if (!subject) { skipped++; continue; }

    const training_date = safeDate(row.training_date) || today();
    const batch = safeStr(row.batch) || '';
    const key = subject.toLowerCase().trim() + '_' + training_date + '_' + batch.toLowerCase().trim();
    importedKeys.push(key);

    const branch_id = matchBranch(row.branch_name);
    const participants = safeStr(row.participants);
    const trainer = safeStr(row.trainer);
    const score = row.score != null ? parseFloat(row.score) : null;
    const notes = safeStr(row.notes);

    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === 'update') {
        stmts.push(env.DB.prepare(
          `UPDATE training SET branch_id = ?, participants = ?, trainer = ?, score = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(branch_id, participants, trainer, score, notes, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env.DB.prepare(
        `INSERT INTO training (training_date, batch, subject, participants, branch_id, trainer, score, notes)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(training_date, batch, subject, participants, branch_id, trainer, score, notes));
      inserted++;
    }
  }

  await batchInsert(env.DB, stmts);

  return ok({ inserted, skipped, updated }, 200, origin);
}

// ─── Checklist ────────────────────────────────────────────────────────────────
async function importChecklist(rows, onDuplicate, env, origin) {
  const existing = await env.DB.prepare('SELECT id, name FROM master_checklist').all();
  const existingMap = new Map();
  (existing.results || []).forEach(c => {
    existingMap.set(c.name.toLowerCase().trim(), c.id);
  });

  const stmts = [];
  const importedNames = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;

  for (const row of rows) {
    const name = safeStr(row.name) || '-';
    importedNames.push(name);

    const key = name.toLowerCase().trim();
    const category = safeStr(row.category) || 'Umum';
    const document_link = safeStr(row.document_link);
    const description = safeStr(row.description);

    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === 'update') {
        stmts.push(env.DB.prepare(
          `UPDATE master_checklist SET category = ?, document_link = ?, description = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(category, document_link, description, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env.DB.prepare(
        `INSERT INTO master_checklist (name, category, document_link, description) VALUES (?, ?, ?, ?)`
      ).bind(name, category, document_link, description));
      inserted++;
    }
  }

  await batchInsert(env.DB, stmts);

  return ok({ inserted, skipped, updated }, 200, origin);
}

// ─── Forms ────────────────────────────────────────────────────────────────────
async function importForms(rows, onDuplicate, env, origin) {
  const existing = await env.DB.prepare('SELECT id, name FROM master_forms').all();
  const existingMap = new Map();
  (existing.results || []).forEach(f => {
    existingMap.set(f.name.toLowerCase().trim(), f.id);
  });

  const stmts = [];
  const importedNames = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;

  for (const row of rows) {
    const name = safeStr(row.name);
    if (!name) { skipped++; continue; }
    importedNames.push(name);

    const key = name.toLowerCase().trim();
    const category = safeStr(row.category) || 'Umum';
    const document_link = safeStr(row.document_link);
    const description = safeStr(row.description);

    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === 'update') {
        stmts.push(env.DB.prepare(
          `UPDATE master_forms SET category = ?, document_link = ?, description = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(category, document_link, description, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env.DB.prepare(
        `INSERT INTO master_forms (name, category, document_link, description) VALUES (?, ?, ?, ?)`
      ).bind(name, category, document_link, description));
      inserted++;
    }
  }

  await batchInsert(env.DB, stmts);

  return ok({ inserted, skipped, updated }, 200, origin);
}

// ─── SOP ──────────────────────────────────────────────────────────────────────
async function importSop(rows, onDuplicate, env, origin) {
  const existing = await env.DB.prepare('SELECT id, name FROM sop').all();
  const existingMap = new Map();
  (existing.results || []).forEach(s => {
    existingMap.set(s.name.toLowerCase().trim(), s.id);
  });

  const stmts = [];
  const importedNames = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;

  for (const row of rows) {
    const name = safeStr(row.name);
    if (!name) { skipped++; continue; }
    importedNames.push(name);

    const key = name.toLowerCase().trim();
    const category = safeStr(row.category) || 'Umum';
    const document_link = safeStr(row.document_link);
    const version = safeStr(row.version) || '1.0';
    const effective_date = safeDate(row.effective_date);
    const notes = safeStr(row.notes);

    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === 'update') {
        stmts.push(env.DB.prepare(
          `UPDATE sop SET category = ?, document_link = ?, version = ?, effective_date = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(category, document_link, version, effective_date, notes, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env.DB.prepare(
        `INSERT INTO sop (name, category, document_link, version, effective_date, notes) VALUES (?, ?, ?, ?, ?, ?)`
      ).bind(name, category, document_link, version, effective_date, notes));
      inserted++;
    }
  }

  await batchInsert(env.DB, stmts);

  return ok({ inserted, skipped, updated }, 200, origin);
}

// ─── Inspection ───────────────────────────────────────────────────────────────
async function importInspection(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const existing = await env.DB.prepare('SELECT id, branch_id, period, inspection_date FROM inspection_reports').all();
  const existingMap = new Map();
  (existing.results || []).forEach(i => {
    if (i.inspection_date) {
      existingMap.set(i.branch_id + '_' + (i.period || '').toLowerCase().trim() + '_' + i.inspection_date, i.id);
    }
  });

  const stmts = [];
  const importedKeys = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;

  for (const row of rows) {
    const branch_name = safeStr(row.branch_name);
    if (!branch_name) { skipped++; continue; }

    const branch_id = matchBranch(branch_name);
    const period = safeStr(row.period) || '-';
    const inspection_date = safeDate(row.inspection_date) || today();
    const key = branch_id + '_' + period.toLowerCase().trim() + '_' + inspection_date;
    importedKeys.push(key);

    const status = safeStr(row.status) || 'Pending';
    const fc_score = row.fc_score != null ? parseFloat(row.fc_score) : null;
    const spv_score = row.spv_score != null ? parseFloat(row.spv_score) : null;
    const document_link = safeStr(row.document_link);
    const notes = safeStr(row.notes);

    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === 'update') {
        stmts.push(env.DB.prepare(
          `UPDATE inspection_reports SET status = ?, fc_score = ?, spv_score = ?, document_link = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(status, fc_score, spv_score, document_link, notes, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env.DB.prepare(
        `INSERT INTO inspection_reports (branch_id, period, inspection_date, status, fc_score, spv_score, document_link, notes)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(branch_id, period, inspection_date, status, fc_score, spv_score, document_link, notes));
      inserted++;
    }
  }

  await batchInsert(env.DB, stmts);

  return ok({ inserted, skipped, updated }, 200, origin);
}

// ─── Cleaning (GC/DC) ─────────────────────────────────────────────────────────
async function importCleaning(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const existing = await env.DB.prepare('SELECT id, branch_id, activity_type, period, activity_date FROM cleaning_reports').all();
  const existingMap = new Map();
  (existing.results || []).forEach(c => {
    if (c.activity_date) {
      existingMap.set(c.branch_id + '_' + c.activity_type.toLowerCase().trim() + '_' + (c.period || '').toLowerCase().trim() + '_' + c.activity_date, c.id);
    }
  });

  const stmts = [];
  const importedKeys = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;

  for (const row of rows) {
    const branch_name = safeStr(row.branch_name);
    if (!branch_name) { skipped++; continue; }

    const branch_id = matchBranch(branch_name);
    const activity_type = safeStr(row.activity_type) || 'General Cleaning';
    const period = safeStr(row.period) || '-';
    const activity_date = safeDate(row.activity_date) || today();
    const key = branch_id + '_' + activity_type.toLowerCase().trim() + '_' + period.toLowerCase().trim() + '_' + activity_date;
    importedKeys.push(key);

    const status = safeStr(row.status) || 'Pending';
    const document_link = safeStr(row.document_link);
    const notes = safeStr(row.notes);

    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === 'update') {
        stmts.push(env.DB.prepare(
          `UPDATE cleaning_reports SET status = ?, document_link = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(status, document_link, notes, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env.DB.prepare(
        `INSERT INTO cleaning_reports (branch_id, activity_type, period, activity_date, status, document_link, notes)
         VALUES (?, ?, ?, ?, ?, ?, ?)`
      ).bind(branch_id, activity_type, period, activity_date, status, document_link, notes));
      inserted++;
    }
  }

  await batchInsert(env.DB, stmts);

  return ok({ inserted, skipped, updated }, 200, origin);
}

// ─── Fogging ──────────────────────────────────────────────────────────────────
async function importFogging(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const existing = await env.DB.prepare('SELECT id, branch_id, period, activity_date FROM fogging_reports').all();
  const existingMap = new Map();
  (existing.results || []).forEach(f => {
    if (f.activity_date) {
      existingMap.set(f.branch_id + '_' + (f.period || '').toLowerCase().trim() + '_' + f.activity_date, f.id);
    }
  });

  const stmts = [];
  const importedKeys = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;

  for (const row of rows) {
    const branch_name = safeStr(row.branch_name);
    if (!branch_name) { skipped++; continue; }

    const branch_id = matchBranch(branch_name);
    const period = safeStr(row.period) || '-';
    const activity_date = safeDate(row.activity_date) || today();
    const key = branch_id + '_' + period.toLowerCase().trim() + '_' + activity_date;
    importedKeys.push(key);

    const status = safeStr(row.status) || 'Pending';
    const document_link = safeStr(row.document_link);
    const notes = safeStr(row.notes);

    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === 'update') {
        stmts.push(env.DB.prepare(
          `UPDATE fogging_reports SET status = ?, document_link = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(status, document_link, notes, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env.DB.prepare(
        `INSERT INTO fogging_reports (branch_id, activity_type, period, activity_date, status, document_link, notes)
         VALUES (?, ?, ?, ?, ?, ?, ?)`
      ).bind(branch_id, 'Fogging', period, activity_date, status, document_link, notes));
      inserted++;
    }
  }

  await batchInsert(env.DB, stmts);

  return ok({ inserted, skipped, updated }, 200, origin);
}

// ─── Basecamp ─────────────────────────────────────────────────────────────────
async function importBasecamp(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const existing = await env.DB.prepare('SELECT id, branch_id, problem, info_date FROM basecamp_reports').all();
  const existingMap = new Map();
  (existing.results || []).forEach(b => {
    if (b.info_date) {
      existingMap.set(b.branch_id + '_' + b.problem.toLowerCase().trim() + '_' + b.info_date, b.id);
    }
  });

  const stmts = [];
  const importedKeys = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;

  for (const row of rows) {
    const problem = safeStr(row.problem);
    if (!problem) { skipped++; continue; }

    const branch_id = matchBranch(row.branch_name);
    const info_date = safeDate(row.info_date) || today();
    const key = branch_id + '_' + problem.toLowerCase().trim() + '_' + info_date;
    importedKeys.push(key);

    const pic = safeStr(row.pic);
    const done_date = safeDate(row.done_date);
    const status = safeStr(row.status) || 'Open';
    const notes = safeStr(row.notes);

    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === 'update') {
        stmts.push(env.DB.prepare(
          `UPDATE basecamp_reports SET pic = ?, done_date = ?, status = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
        ).bind(pic, done_date, status, notes, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env.DB.prepare(
        `INSERT INTO basecamp_reports (branch_id, problem, pic, info_date, done_date, status, notes)
         VALUES (?, ?, ?, ?, ?, ?, ?)`
      ).bind(branch_id, problem, pic, info_date, done_date, status, notes));
      inserted++;
    }
  }

  await batchInsert(env.DB, stmts);

  return ok({ inserted, skipped, updated }, 200, origin);
}

// ─── Supply / Chemical ────────────────────────────────────────────────────────
async function importSupply(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const existing = await env.DB.prepare('SELECT id, submitter_name, submitted_at FROM supply_requests').all();
  const existingMap = new Map();
  (existing.results || []).forEach(s => {
    if (s.submitter_name && s.submitted_at) {
      existingMap.set(s.submitter_name.toLowerCase().trim() + '_' + s.submitted_at, s.id);
    }
  });

  const stmts = [];
  const importedKeys = [];
  let inserted = 0;
  let updated = 0;
  let skipped = 0;

  for (const row of rows) {
    const submitter_name = safeStr(row.submitter_name) || '-';

    const submitted_at = safeDate(row.submitted_at) || today();
    const key = submitter_name.toLowerCase().trim() + '_' + submitted_at;
    importedKeys.push(key);

    const branch_id = matchBranch(row.branch_name);
    const branchObj = branch_id ? bRows.results.find(b => b.id === branch_id) : null;
    const branch_name = branchObj ? branchObj.full_name : safeStr(row.branch_name);
    const tools_items = safeStr(row.tools_items);
    const tools_quantity = safeStr(row.tools_quantity);
    const chemical_items = safeStr(row.chemical_items);
    const chemical_quantity = safeStr(row.chemical_quantity);
    const additional_notes = safeStr(row.additional_notes);
    const status = safeStr(row.status) || 'Pending';

    if (existingMap.has(key)) {
      const id = existingMap.get(key);
      if (onDuplicate === 'update') {
        stmts.push(env.DB.prepare(
          `UPDATE supply_requests SET branch_id = ?, branch_name = ?, tools_items = ?, tools_quantity = ?, chemical_items = ?, chemical_quantity = ?, additional_notes = ?, status = ?, processed_at = datetime('now') WHERE id = ?`
        ).bind(branch_id, branch_name, tools_items, tools_quantity, chemical_items, chemical_quantity, additional_notes, status, id));
        updated++;
      } else {
        skipped++;
      }
    } else {
      stmts.push(env.DB.prepare(
        `INSERT INTO supply_requests (submitted_at, submitter_name, branch_id, branch_name, tools_items, tools_quantity, chemical_items, chemical_quantity, additional_notes, status)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(submitted_at, submitter_name, branch_id, branch_name, tools_items, tools_quantity, chemical_items, chemical_quantity, additional_notes, status));
      inserted++;
    }
  }

  await batchInsert(env.DB, stmts);

  return ok({ inserted, skipped, updated }, 200, origin);
}

async function exportBackup(env, origin) {
  const tables = [
    'users', 'branches', 'employees', 'contracts', 'activity_schedule',
    'issues', 'one_on_one', 'training', 'relievers', 'inspection_reports',
    'cleaning_reports', 'fogging_reports', 'basecamp_reports', 'supply_requests',
    'master_checklist', 'master_forms', 'sop'
  ];
  
  const promises = tables.map(async (table) => {
    try {
      const res = await env.DB.prepare(`SELECT * FROM ${table}`).all();
      return { table, rows: res.results || [] };
    } catch (e) {
      return { table, error: e.message };
    }
  });

  const results = await Promise.all(promises);
  const backup = {};
  results.forEach(item => {
    backup[item.table] = item.rows;
  });

  return ok({
    timestamp: new Date().toISOString(),
    database: backup
  }, 200, origin);
}

async function triggerCalendarSync(env, origin) {
  const { getSyncStatements } = await import('../utils/calendar.js');
  const allStmts = [];

  // Clear existing calendar_events table
  allStmts.push(env.DB.prepare('DELETE FROM calendar_events'));

  // 1. Contracts
  const contr = await env.DB.prepare(`SELECT c.id, c.end_date, c.status, c.branch_id, COALESCE(e.full_name, c.employee_name, '?') as empName FROM contracts c LEFT JOIN employees e ON c.employee_id = e.id`).all();
  (contr.results || []).forEach(c => {
    allStmts.push(...getSyncStatements(env.DB, 'contracts', c.id, c));
  });

  // 2. Schedules
  const sched = await env.DB.prepare(`SELECT * FROM activity_schedule`).all();
  (sched.results || []).forEach(s => {
    allStmts.push(...getSyncStatements(env.DB, 'schedule', s.id, s));
  });

  // 3. Relievers
  const rel = await env.DB.prepare(`SELECT * FROM relievers`).all();
  (rel.results || []).forEach(r => {
    allStmts.push(...getSyncStatements(env.DB, 'relievers', r.id, r));
  });

  // 4. Issues
  const iss = await env.DB.prepare(`SELECT * FROM issues`).all();
  (iss.results || []).forEach(i => {
    allStmts.push(...getSyncStatements(env.DB, 'issues', i.id, i));
  });

  // 5. One on One
  const ooo = await env.DB.prepare(`SELECT * FROM one_on_one`).all();
  (ooo.results || []).forEach(o => {
    allStmts.push(...getSyncStatements(env.DB, 'one_on_one', o.id, o));
  });

  // 6. Training
  const tr = await env.DB.prepare(`SELECT * FROM training`).all();
  (tr.results || []).forEach(t => {
    allStmts.push(...getSyncStatements(env.DB, 'training', t.id, t));
  });

  // 7. Cleaning Reports
  const clean = await env.DB.prepare(`SELECT * FROM cleaning_reports`).all();
  (clean.results || []).forEach(c => {
    allStmts.push(...getSyncStatements(env.DB, 'cleaning', c.id, c));
  });

  // 8. Inspection Reports
  const insp = await env.DB.prepare(`SELECT * FROM inspection_reports`).all();
  (insp.results || []).forEach(i => {
    allStmts.push(...getSyncStatements(env.DB, 'inspection', i.id, i));
  });

  // 9. Fogging Reports
  const fog = await env.DB.prepare(`SELECT * FROM fogging_reports`).all();
  (fog.results || []).forEach(f => {
    allStmts.push(...getSyncStatements(env.DB, 'fogging', f.id, f));
  });

  // 10. Basecamp Reports
  const base = await env.DB.prepare(`SELECT * FROM basecamp_reports`).all();
  (base.results || []).forEach(b => {
    allStmts.push(...getSyncStatements(env.DB, 'basecamp', b.id, b));
  });

  // 11. Supply Requests
  const supply = await env.DB.prepare(`SELECT * FROM supply_requests`).all();
  (supply.results || []).forEach(s => {
    allStmts.push(...getSyncStatements(env.DB, 'supply', s.id, s));
  });

  // Batch execute
  if (allStmts.length > 0) {
    const chunkSize = 100;
    for (let i = 0; i < allStmts.length; i += chunkSize) {
      await env.DB.batch(allStmts.slice(i, i + chunkSize));
    }
  }

  return ok({ message: `Berhasil sinkronisasi ${allStmts.length} event kalender dari data lama.` }, 200, origin);
}
