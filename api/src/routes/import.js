import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden } from '../utils/response.js';

export async function handleImport(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'employees', 'write')) return forbidden(origin);

  const url = new URL(request.url);
  const path = url.pathname.replace('/api/import', '');

  if (request.method !== 'POST') return error('Method not allowed', 405, origin);

  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }

  const { rows = [], onDuplicate = 'skip' } = body;
  if (!Array.isArray(rows)) return error('rows must be an array', 400, origin);
  if (rows.length === 0) return ok({ inserted: 0, skipped: 0, message: 'Tidak ada data' }, 200, origin);

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
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  // Excel serial number
  if (/^\d{4,5}$/.test(s)) {
    const d = new Date(Date.UTC(1899, 11, 30) + Number(s) * 86400000);
    return d.toISOString().slice(0, 10);
  }
  // dd/mm/yyyy or yyyy/mm/dd
  const parts = s.split(/[\/\-\.]/);
  if (parts.length === 3) {
    const [a, b, c] = parts;
    if (a.length === 4 && Number(a) > 1900) return `${a}-${b.padStart(2,'0')}-${c.padStart(2,'0')}`;
    if (c.length === 4 && Number(c) > 1900) return `${c}-${b.padStart(2,'0')}-${a.padStart(2,'0')}`;
    if (c.length === 2) {
      const yr = Number(c) >= 50 ? `19${c}` : `20${c}`;
      return `${yr}-${b.padStart(2,'0')}-${a.padStart(2,'0')}`;
    }
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  return null;
}
function today() { return new Date().toISOString().slice(0, 10); }

// Fuzzy branch match: by full_name, code, or partial name
function makeBranchMatcher(branches) {
  return (str) => {
    if (!str) return null;
    const s = String(str).toLowerCase().trim();
    // Exact match first
    const exact = branches.find(r =>
      r.full_name?.toLowerCase() === s ||
      r.code?.toLowerCase() === s ||
      r.name?.toLowerCase() === s
    );
    if (exact) return exact.id;
    // Partial: branch name contains the search string OR vice versa
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

// ─── Employees ───────────────────────────────────────────────────────────────
async function importEmployees(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const stmts = [];
  let skipped = 0;
  for (const row of rows) {
    const full_name = safeStr(row.full_name);
    if (!full_name) { skipped++; continue; }

    const sql = onDuplicate === 'skip'
      ? `INSERT OR IGNORE INTO employees (full_name, branch_id, division, phone, join_date, status, notes)
         VALUES (?, ?, ?, ?, ?, ?, ?)`
      : `INSERT INTO employees (full_name, branch_id, division, phone, join_date, status, notes)
         VALUES (?, ?, ?, ?, ?, ?, ?)
         ON CONFLICT(id) DO UPDATE SET
           branch_id=excluded.branch_id, division=excluded.division,
           phone=excluded.phone, join_date=excluded.join_date,
           status=excluded.status, notes=excluded.notes, updated_at=datetime('now')`;

    stmts.push(env.DB.prepare(sql).bind(
      full_name, matchBranch(row.branch_name),
      safeStr(row.division) || 'FACILITY CARE',
      safeStr(row.phone), safeDate(row.join_date),
      safeStr(row.status) || 'Aktif', safeStr(row.notes)
    ));
  }
  const inserted = await batchInsert(env.DB, stmts);
  return ok({ inserted, skipped, message: `Karyawan: ${inserted} berhasil, ${skipped} dilewati` }, 200, origin);
}

// ─── Contracts ────────────────────────────────────────────────────────────────
async function importContracts(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const eRows = await env.DB.prepare('SELECT id, full_name FROM employees').all();
  const matchEmployee = (str) => {
    if (!str) return null;
    const s = str.toLowerCase().trim();
    const emp = eRows.results.find(r =>
      r.full_name.toLowerCase() === s ||
      r.full_name.toLowerCase().includes(s.split(' ')[0])
    );
    return emp ? emp.id : null;
  };

  // Detect if employee_name column has been migrated
  let hasEmployeeNameCol = false;
  try {
    await env.DB.prepare('SELECT employee_name FROM contracts LIMIT 1').first();
    hasEmployeeNameCol = true;
  } catch { hasEmployeeNameCol = false; }

  const stmts = [];
  let skipped = 0;
  for (const row of rows) {
    const employee_name = safeStr(row.employee_name);
    if (!employee_name) { skipped++; continue; }

    const employee_id = matchEmployee(employee_name);
    const start_date  = safeDate(row.start_date) || today();
    const end_date    = safeDate(row.end_date)   || today();

    if (hasEmployeeNameCol) {
      stmts.push(env.DB.prepare(
        `INSERT INTO contracts (employee_id, employee_name, branch_id, start_date, end_date, contract_type, pkwt_number, status, notes)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(
        employee_id, employee_name, matchBranch(row.branch_name),
        start_date, end_date,
        safeStr(row.contract_type), safeStr(row.pkwt_number),
        safeStr(row.status) || 'Aktif', safeStr(row.notes)
      ));
    } else {
      // Old schema without employee_name column — need employee_id
      if (!employee_id) { skipped++; continue; }
      stmts.push(env.DB.prepare(
        `INSERT INTO contracts (employee_id, branch_id, start_date, end_date, contract_type, pkwt_number, status, notes)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
      ).bind(
        employee_id, matchBranch(row.branch_name),
        start_date, end_date,
        safeStr(row.contract_type), safeStr(row.pkwt_number),
        safeStr(row.status) || 'Aktif', safeStr(row.notes)
      ));
    }
  }
  const inserted = await batchInsert(env.DB, stmts);
  return ok({ inserted, skipped }, 200, origin);
}


// ─── Relievers ────────────────────────────────────────────────────────────────
async function importRelievers(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const stmts = [];
  let skipped = 0;
  for (const row of rows) {
    const reliever_name = safeStr(row.reliever_name);
    if (!reliever_name) { skipped++; continue; }
    stmts.push(env.DB.prepare(
      `INSERT INTO relievers (branch_id, original_fc_name, period, reliever_name, backup_date, completion_date, reason, shift, status)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
    ).bind(
      matchBranch(row.branch_name), safeStr(row.original_fc_name), safeStr(row.period),
      reliever_name, safeDate(row.backup_date), safeDate(row.completion_date),
      safeStr(row.reason), safeStr(row.shift), safeStr(row.status) || 'Pending'
    ));
  }
  const inserted = await batchInsert(env.DB, stmts);
  return ok({ inserted, skipped }, 200, origin);
}

// ─── Schedule ─────────────────────────────────────────────────────────────────
async function importSchedule(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const stmts = [];
  let skipped = 0;
  for (const row of rows) {
    const activity_type = safeStr(row.activity_type);
    if (!activity_type) { skipped++; continue; }
    stmts.push(env.DB.prepare(
      `INSERT INTO activity_schedule (branch_id, activity_type, period, pic, opening_date, target_date, completion_date, status, notes)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
    ).bind(
      matchBranch(row.branch_name), activity_type, safeStr(row.period),
      safeStr(row.pic), safeDate(row.opening_date), safeDate(row.target_date),
      safeDate(row.completion_date), safeStr(row.status) || 'Pending', safeStr(row.notes)
    ));
  }
  const inserted = await batchInsert(env.DB, stmts);
  return ok({ inserted, skipped }, 200, origin);
}

// ─── Issues ───────────────────────────────────────────────────────────────────
async function importIssues(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const stmts = [];
  let skipped = 0;
  for (const row of rows) {
    const complaint = safeStr(row.complaint);
    if (!complaint) { skipped++; continue; }

    const report_date = safeDate(row.report_date) || today();
    let day_count = null;
    if (row.completion_date && report_date) {
      const d1 = new Date(report_date), d2 = new Date(safeDate(row.completion_date));
      if (!isNaN(d1) && !isNaN(d2)) day_count = Math.floor((d2 - d1) / 86400000);
    }
    stmts.push(env.DB.prepare(
      `INSERT INTO issues (report_date, branch_id, category, source, complaint, employee_name, fc_specialist, solution, status, completion_date, day_count)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
    ).bind(
      report_date, matchBranch(row.branch_name),
      safeStr(row.category) || 'Lainnya', safeStr(row.source),
      complaint, safeStr(row.employee_name), safeStr(row.fc_specialist),
      safeStr(row.solution), safeStr(row.status) || 'Open',
      safeDate(row.completion_date), day_count
    ));
  }
  const inserted = await batchInsert(env.DB, stmts);
  return ok({ inserted, skipped }, 200, origin);
}

// ─── One on One ───────────────────────────────────────────────────────────────
async function importOneOnOne(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const stmts = [];
  let skipped = 0;
  for (const row of rows) {
    const employee_name = safeStr(row.employee_name);
    if (!employee_name) { skipped++; continue; }

    const meeting_date = safeDate(row.meeting_date) || today();
    let day_count = null;
    if (row.completion_date && meeting_date) {
      const d1 = new Date(meeting_date), d2 = new Date(safeDate(row.completion_date));
      if (!isNaN(d1) && !isNaN(d2)) day_count = Math.floor((d2 - d1) / 86400000);
    }
    stmts.push(env.DB.prepare(
      `INSERT INTO one_on_one (meeting_date, branch_id, employee_name, pic, problem, solution, status, completion_date, day_count)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
    ).bind(
      meeting_date, matchBranch(row.branch_name), employee_name,
      safeStr(row.pic), safeStr(row.problem) || '-', safeStr(row.solution),
      safeStr(row.status) || 'Open', safeDate(row.completion_date), day_count
    ));
  }
  const inserted = await batchInsert(env.DB, stmts);
  return ok({ inserted, skipped }, 200, origin);
}

// ─── Training ─────────────────────────────────────────────────────────────────
async function importTraining(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const stmts = [];
  let skipped = 0;
  for (const row of rows) {
    const subject = safeStr(row.subject);
    if (!subject) { skipped++; continue; }
    stmts.push(env.DB.prepare(
      `INSERT INTO training (training_date, batch, subject, participants, branch_id, trainer, score, notes)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
    ).bind(
      safeDate(row.training_date) || today(),
      safeStr(row.batch), subject, safeStr(row.participants),
      matchBranch(row.branch_name), safeStr(row.trainer),
      row.score != null ? parseFloat(row.score) : null,
      safeStr(row.notes)
    ));
  }
  const inserted = await batchInsert(env.DB, stmts);
  return ok({ inserted, skipped }, 200, origin);
}

// ─── Checklist ────────────────────────────────────────────────────────────────
async function importChecklist(rows, onDuplicate, env, origin) {
  const stmts = [];
  let skipped = 0;
  for (const row of rows) {
    const name = safeStr(row.name);
    if (!name) { skipped++; continue; }
    const sql = onDuplicate === 'skip'
      ? `INSERT OR IGNORE INTO master_checklist (name, category, document_link, description) VALUES (?, ?, ?, ?)`
      : `INSERT INTO master_checklist (name, category, document_link, description) VALUES (?, ?, ?, ?)
         ON CONFLICT(id) DO UPDATE SET name=excluded.name, category=excluded.category, document_link=excluded.document_link`;
    stmts.push(env.DB.prepare(sql).bind(name, safeStr(row.category), safeStr(row.document_link), safeStr(row.description)));
  }
  const inserted = await batchInsert(env.DB, stmts);
  return ok({ inserted, skipped }, 200, origin);
}

// ─── Forms ────────────────────────────────────────────────────────────────────
async function importForms(rows, onDuplicate, env, origin) {
  const stmts = [];
  let skipped = 0;
  for (const row of rows) {
    const name = safeStr(row.name);
    if (!name) { skipped++; continue; }
    stmts.push(env.DB.prepare(
      `INSERT OR IGNORE INTO master_forms (name, category, document_link, description) VALUES (?, ?, ?, ?)`
    ).bind(name, safeStr(row.category), safeStr(row.document_link), safeStr(row.description)));
  }
  const inserted = await batchInsert(env.DB, stmts);
  return ok({ inserted, skipped }, 200, origin);
}

// ─── SOP ──────────────────────────────────────────────────────────────────────
async function importSop(rows, onDuplicate, env, origin) {
  const stmts = [];
  let skipped = 0;
  for (const row of rows) {
    const name = safeStr(row.name);
    if (!name) { skipped++; continue; } // only name is required
    const sql = onDuplicate === 'skip'
      ? `INSERT OR IGNORE INTO sop (name, category, document_link, version, effective_date, notes) VALUES (?, ?, ?, ?, ?, ?)`
      : `INSERT INTO sop (name, category, document_link, version, effective_date, notes) VALUES (?, ?, ?, ?, ?, ?)
         ON CONFLICT(id) DO UPDATE SET name=excluded.name, category=excluded.category,
           document_link=excluded.document_link, version=excluded.version`;
    stmts.push(env.DB.prepare(sql).bind(
      name, safeStr(row.category) || 'Umum',
      safeStr(row.document_link), safeStr(row.version),
      safeDate(row.effective_date), safeStr(row.notes)
    ));
  }
  const inserted = await batchInsert(env.DB, stmts);
  return ok({ inserted, skipped }, 200, origin);
}

// ─── Inspection ───────────────────────────────────────────────────────────────
async function importInspection(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const stmts = [];
  let skipped = 0;
  for (const row of rows) {
    const branch_name = safeStr(row.branch_name);
    if (!branch_name) { skipped++; continue; }
    stmts.push(env.DB.prepare(
      `INSERT INTO inspection_reports (branch_id, period, inspection_date, status, fc_score, spv_score, document_link, notes)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
    ).bind(
      matchBranch(branch_name), safeStr(row.period) || '-',
      safeDate(row.inspection_date) || today(),
      safeStr(row.status) || 'Pending',
      row.fc_score != null ? parseFloat(row.fc_score) : null,
      row.spv_score != null ? parseFloat(row.spv_score) : null,
      safeStr(row.document_link), safeStr(row.notes)
    ));
  }
  const inserted = await batchInsert(env.DB, stmts);
  return ok({ inserted, skipped }, 200, origin);
}

// ─── Cleaning (GC/DC) ─────────────────────────────────────────────────────────
async function importCleaning(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const stmts = [];
  let skipped = 0;
  for (const row of rows) {
    const branch_name = safeStr(row.branch_name);
    if (!branch_name) { skipped++; continue; }
    stmts.push(env.DB.prepare(
      `INSERT INTO cleaning_reports (branch_id, activity_type, period, activity_date, status, document_link, notes)
       VALUES (?, ?, ?, ?, ?, ?, ?)`
    ).bind(
      matchBranch(branch_name),
      safeStr(row.activity_type) || 'General Cleaning',
      safeStr(row.period) || '-',
      safeDate(row.activity_date) || today(),
      safeStr(row.status) || 'Pending',
      safeStr(row.document_link), safeStr(row.notes)
    ));
  }
  const inserted = await batchInsert(env.DB, stmts);
  return ok({ inserted, skipped }, 200, origin);
}

// ─── Fogging ──────────────────────────────────────────────────────────────────
async function importFogging(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const stmts = [];
  let skipped = 0;
  for (const row of rows) {
    // Accept any non-empty row — at least branch OR date should be present
    const branch_name  = safeStr(row.branch_name);
    const activity_date = safeDate(row.activity_date);
    if (!branch_name && !activity_date) { skipped++; continue; }

    stmts.push(env.DB.prepare(
      `INSERT INTO fogging_reports (branch_id, activity_type, period, activity_date, status, document_link, notes)
       VALUES (?, ?, ?, ?, ?, ?, ?)`
    ).bind(
      matchBranch(branch_name), 'Fogging',
      safeStr(row.period) || '-',
      activity_date || today(),
      safeStr(row.status) || 'Pending',
      safeStr(row.document_link), safeStr(row.notes)
    ));
  }
  const inserted = await batchInsert(env.DB, stmts);
  return ok({ inserted, skipped }, 200, origin);
}

// ─── Basecamp ─────────────────────────────────────────────────────────────────
async function importBasecamp(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const stmts = [];
  let skipped = 0;
  for (const row of rows) {
    const problem = safeStr(row.problem);
    if (!problem) { skipped++; continue; }
    stmts.push(env.DB.prepare(
      `INSERT INTO basecamp_reports (branch_id, problem, pic, info_date, done_date, status, notes)
       VALUES (?, ?, ?, ?, ?, ?, ?)`
    ).bind(
      matchBranch(row.branch_name), problem, safeStr(row.pic),
      safeDate(row.info_date) || today(),
      safeDate(row.done_date),
      safeStr(row.status) || 'Open', safeStr(row.notes)
    ));
  }
  const inserted = await batchInsert(env.DB, stmts);
  return ok({ inserted, skipped }, 200, origin);
}

// ─── Supply / Chemical ────────────────────────────────────────────────────────
async function importSupply(rows, onDuplicate, env, origin) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const stmts = [];
  let skipped = 0;
  for (const row of rows) {
    const submitter_name = safeStr(row.submitter_name);
    if (!submitter_name) { skipped++; continue; }
    const branch_id = matchBranch(row.branch_name);
    const branchObj = branch_id ? bRows.results.find(b => b.id === branch_id) : null;
    stmts.push(env.DB.prepare(
      `INSERT INTO supply_requests (submitted_at, submitter_name, branch_id, branch_name, tools_items, tools_quantity, chemical_items, chemical_quantity, additional_notes, status)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
    ).bind(
      safeDate(row.submitted_at) || today(),
      submitter_name, branch_id,
      branchObj ? branchObj.full_name : safeStr(row.branch_name),
      safeStr(row.tools_items), safeStr(row.tools_quantity),
      safeStr(row.chemical_items), safeStr(row.chemical_quantity),
      safeStr(row.additional_notes), safeStr(row.status) || 'Pending'
    ));
  }
  const inserted = await batchInsert(env.DB, stmts);
  return ok({ inserted, skipped }, 200, origin);
}
