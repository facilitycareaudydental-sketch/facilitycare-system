import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden } from '../utils/response.js';
import {
  buildAuditStatement,
  buildCreateAuditStatement,
  buildBulkHardDeleteAuditStatements
} from '../utils/audit.js';
import {
  executeJournaledImport,
  extractImportRequestMeta,
  buildStagedWipeGroups
} from '../utils/import_journal.js';
import { isManagementBranch, getManagementBranchSubquery, resolveManagementBranchIds } from '../utils/branchMaster.js';

const SOFT_DELETE_TABLES = new Set([
  'branches', 'employees', 'contracts', 'issues', 'activity_schedule',
  'relievers', 'one_on_one', 'master_forms', 'inspection_reports',
  'cleaning_reports', 'fogging_reports', 'basecamp_reports', 'sop',
  'master_checklist', 'supply_requests'
]);

const EXCLUDED_MIGRATION_IMPORT_MODULES = new Set(['/relievers', '/supply']);

export async function handleImport(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (user.role !== 'superadmin') return forbidden(origin);

  const url = new URL(request.url);
  const path = url.pathname.replace('/api/import', '');

  if (path === '/backup' && request.method === 'GET') {
    if (!hasPermission(user, 'users', 'admin')) return forbidden(origin);
    return exportBackup(env, origin);
  }

  if (path === '/sync-calendar' && request.method === 'POST') {
    return triggerCalendarSync(env, user, origin);
  }

  // GET /api/import/counts — return row counts for active records only
  if (path === '/counts' && request.method === 'GET') {
    const tables = [
      'validation_options', 'branches', 'employees', 'contracts', 'sp_data',
      'mutasi_data', 'overtime_records', 'os_migrations', 'issues', 'one_on_one',
      'training', 'activity_schedule', 'inspection_reports', 'cleaning_reports',
      'fogging_reports', 'basecamp_reports', 'sop', 'master_checklist',
      'master_forms', 'master_hygiene_standards', 'master_dac_standards'
    ];
    const counts = {};
    for (const t of tables) {
      try {
        let whereClause = SOFT_DELETE_TABLES.has(t) ? ' WHERE deleted_at IS NULL' : '';
        if (t === 'employees' || t === 'contracts') {
          whereClause = ` WHERE deleted_at IS NULL AND division IN ('FACILITY CARE', 'SECURITY') AND ${getManagementBranchSubquery('branch_id')}`;
        } else if (t === 'branches') {
          whereClause = ' WHERE deleted_at IS NULL AND is_active = 1';
        } else if (t === 'sop') {
          whereClause = " WHERE deleted_at IS NULL AND (category != 'sop_bundle' OR category IS NULL)";
        }
        const r = await env.DB.prepare(`SELECT COUNT(*) as cnt FROM ${t}${whereClause}`).first();
        counts[t] = r.cnt;
      } catch (e) {
        counts[t] = -1;
      }
    }
    return ok(counts, 200, origin);
  }

  if (request.method !== 'POST') return error('Method not allowed', 405, origin);

  if (EXCLUDED_MIGRATION_IMPORT_MODULES.has(path)) {
    return error('Modul Reliever dan Permintaan Chemical dikecualikan dari migrasi FCMS Baru.', 400, origin);
  }


  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }

  const { rows = [], onDuplicate = 'skip' } = body;
  if (!Array.isArray(rows)) return error('rows must be an array', 400, origin);

  const importMeta = extractImportRequestMeta(request);

  try {
    switch (path) {
      case '/employees':    return importEmployees(rows, onDuplicate, env, user, origin, importMeta);
      case '/contracts':    return importContracts(rows, onDuplicate, env, user, origin, importMeta);
      case '/relievers':    return importRelievers(rows, onDuplicate, env, user, origin, importMeta);
      case '/schedule':     return importSchedule(rows, onDuplicate, env, user, origin, importMeta);
      case '/issues':       return importIssues(rows, onDuplicate, env, user, origin, importMeta);
      case '/one_on_one':   return importOneOnOne(rows, onDuplicate, env, user, origin, importMeta);
      case '/training':     return importTraining(rows, onDuplicate, env, user, origin, importMeta);
      case '/checklist':    return importChecklist(rows, onDuplicate, env, user, origin, importMeta);
      case '/forms':        return importForms(rows, onDuplicate, env, user, origin, importMeta);
      case '/sop':          return importSop(rows, onDuplicate, env, user, origin, importMeta);
      case '/inspection':   return importInspection(rows, onDuplicate, env, user, origin, importMeta);
      case '/cleaning':     return importCleaning(rows, onDuplicate, env, user, origin, importMeta);
      case '/fogging':      return importFogging(rows, onDuplicate, env, user, origin, importMeta);
      case '/basecamp':     return importBasecamp(rows, onDuplicate, env, user, origin, importMeta);
      case '/supply':       return importSupply(rows, onDuplicate, env, user, origin, importMeta);
      case '/validation':   return importValidation(rows, onDuplicate, env, user, origin, importMeta);
      case '/sp':           return importSP(rows, onDuplicate, env, user, origin, importMeta);
      case '/mutasi':       return importMutasi(rows, onDuplicate, env, user, origin, importMeta);
      default:              return error('Unknown import module', 404, origin);
    }
  } catch (err) {
    console.error('Import error:', err); return error('Gagal memproses import data', 500, origin);
  }
}

// ─── Helpers ─────────────────────────────────────────────────────────────────
function safeStr(v)  { return v !== undefined && v !== null && String(v).trim() !== '' ? String(v).trim() : null; }
function safeDate(v) {
  if (!v || v === '0') return null;

  // JS Date string from ISO format
  if (typeof v === 'string' && v.match(/^\d{4}-\d{2}-\d{2}/)) return v.slice(0, 10);

  // Excel serial number
  const excelDays = Number(v);
  if (!isNaN(excelDays) && excelDays > 20000 && excelDays < 99999) {
    // Use Date.UTC(1899,11,30) formula which matches SheetJS cellDates:true output
    const d = new Date(Date.UTC(1899, 11, 30) + excelDays * 86400000);
    if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  }

  // Parts-based parsing: DD/MM/YYYY, MM/DD/YYYY, D/M/YY etc.
  if (typeof v === 'string') {
    const s = v.trim();
    const parts = s.split(/[\/\-\.\s]+/);
    if (parts.length === 3) {
      let [a, b, c] = parts.map(p => p.trim());
      
      // Handle Indonesian month names
      const indoMonths = {
        'jan': '01', 'januari': '01',
        'feb': '02', 'februari': '02',
        'mar': '03', 'maret': '03',
        'apr': '04', 'april': '04',
        'mei': '05',
        'jun': '06', 'juni': '06',
        'jul': '07', 'juli': '07',
        'agu': '08', 'agustus': '08',
        'sep': '09', 'september': '09',
        'okt': '10', 'oktober': '10',
        'nov': '11', 'november': '11',
        'des': '12', 'desember': '12'
      };
      const bLower = b.toLowerCase();
      if (indoMonths[bLower]) {
        b = indoMonths[bLower];
      }
      
      const na = Number(a), nb = Number(b), nc = Number(c);
      if (a.length === 4 && na > 1900) {
        return `${a}-${b.padStart(2,'0')}-${c.padStart(2,'0')}`;
      }
      if (c.length === 4 && nc > 1900) {
        if (na > 12) return `${c}-${b.padStart(2,'0')}-${a.padStart(2,'0')}`;
        if (nb > 12) return `${c}-${a.padStart(2,'0')}-${b.padStart(2,'0')}`;
        return `${c}-${b.padStart(2,'0')}-${a.padStart(2,'0')}`;
      }
      if (c.length === 2 && !isNaN(nc)) {
        const year = nc >= 50 ? `19${c}` : `20${c}`;
        return `${year}-${b.padStart(2,'0')}-${a.padStart(2,'0')}`;
      }
    }
    const d = new Date(s);
    if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  }
  return null;
}
function today() { return new Date().toISOString().slice(0, 10); }

function makeBranchMatcher(branches) {
  return (str) => {
    if (!str) return null;
    const s = String(str).toLowerCase().trim();
    
    // Tahap 2: Exact prefix match
    const codeMatch = s.match(/^(\d{3})/);
    if (codeMatch) {
      const code = codeMatch[1];
      const byCode = branches.find(r => r.code === code);
      if (byCode) return byCode.id;
    }

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

// ─── Shared Journaled Import Helper ──────────────────────────────────────────
async function runStandardJournaledImport(env, user, origin, { module, rows, onDuplicate, importMeta, setup, processRow, formatResult }) {
  if (rows.length === 0) {
    return ok(
      formatResult ? formatResult({ inserted: 0, updated: 0, skipped: 0 }) : { inserted: 0, skipped: 0, updated: 0 },
      200,
      origin,
      undefined,
      { 'X-Audit-Batched': 'true' }
    );
  }
  return await executeJournaledImport(env, user, origin, {
    module,
    rows,
    onDuplicate,
    importMeta,
    buildPlan: async ({ startRowIndex, initialCounts }) => {
      const ctx = setup ? await setup() : {};
      const seenNewKeys = new Set();
      const insertGroups = [];
      let inserted = Number(initialCounts?.inserted || 0);
      let updated = Number(initialCounts?.updated || 0);
      let skipped = Number(initialCounts?.skipped || 0);

      for (let idx = startRowIndex; idx < rows.length; idx++) {
        const res = processRow(rows[idx], ctx, { onDuplicate, seenNewKeys });
        if (!res || res.action === 'skip') {
          skipped++;
          continue;
        }
        if (res.action === 'update') updated++;
        else if (res.action === 'insert') inserted++;
        insertGroups.push({
          rowIndex: idx,
          counts: { inserted, updated, skipped },
          stmts: res.stmts,
        });
      }

      return { insertGroups, finalCounts: { inserted, updated, skipped } };
    },
    buildFinalResult: (counts) => (
      formatResult ? formatResult(counts) : { inserted: counts.inserted, skipped: counts.skipped, updated: counts.updated }
    ),
  });
}

// ─── Validation Options (Differential & Stage-First Safe) ─────────────────────
async function importValidation(rows, onDuplicate, env, user, origin, importMeta = {}) {
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

  const totalItems = branchSet.size + picSet.size + kegiatanSet.size + quartalSet.size + pkwtSet.size;
  if (totalItems === 0) {
    return ok({ inserted: 0, skipped: 0 }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
  }

  return await executeJournaledImport(env, user, origin, {
    module: 'validation',
    rows,
    onDuplicate,
    importMeta,
    buildPlan: async () => {
      const insertGroups = [];
      const wipeGroups = [];

      // 1. Branches Differential Sync
      const existingBranchesRes = await env.DB.prepare('SELECT * FROM branches').all();
      const existingBranchesByCode = new Map();
      (existingBranchesRes.results || []).forEach(b => existingBranchesByCode.set(b.code, b));

      const activeBranchCodeSet = new Set();
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
        activeBranchCodeSet.add(code);
        const existingB = existingBranchesByCode.get(code);
        if (existingB) {
          if (existingB.name !== name || existingB.full_name !== full_name || Number(existingB.is_active) !== 1) {
            const afterB = { ...existingB, name, full_name, is_active: 1 };
            insertGroups.push({
              stmts: [
                env.DB.prepare(
                  `UPDATE branches SET name = ?, full_name = ?, is_active = 1 WHERE code = ?`
                ).bind(name, full_name, code),
                buildAuditStatement(env, user, 'UPDATE', 'branches', existingB.id ?? code, existingB, afterB, { source: 'import_validation' }, 'SUCCESS', true),
              ],
            });
          }
        } else {
          const snap = { code, name, full_name, is_active: 1 };
          insertGroups.push({
            stmts: [
              env.DB.prepare(
                `INSERT INTO branches (code, name, full_name, is_active) VALUES (?, ?, ?, 1)`
              ).bind(code, name, full_name),
              buildCreateAuditStatement(env, user, 'branches', code, snap, { source: 'import_validation' }),
            ],
          });
        }
      }

      if (activeBranchCodeSet.size > 0) {
        const toDeactivate = (existingBranchesRes.results || []).filter(b => !activeBranchCodeSet.has(b.code) && Number(b.is_active) !== 0);
        for (const b of toDeactivate) {
          const afterB = { ...b, is_active: 0 };
          wipeGroups.push({
            stmts: [
              env.DB.prepare(`UPDATE branches SET is_active = 0 WHERE id = ?`).bind(b.id),
              buildAuditStatement(env, user, 'UPDATE', 'branches', b.id, b, afterB, { source: 'import_validation_deactivate' }, 'SUCCESS', true),
            ],
          });
        }
      }

      // 2. PIC list Differential Sync (using string PK pic name as target_id, never rowid)
      const existingPicsRes = await env.DB.prepare('SELECT * FROM pic_list').all();
      const existingPicsMap = new Map();
      (existingPicsRes.results || []).forEach(p => existingPicsMap.set(p.name, p));

      const activePics = Array.from(picSet);
      const activePicSet = new Set(activePics);
      for (const pic of activePics) {
        const existingP = existingPicsMap.get(pic);
        if (existingP) {
          if (Number(existingP.is_active) !== 1) {
            const afterP = { ...existingP, is_active: 1 };
            insertGroups.push({
              stmts: [
                env.DB.prepare(`UPDATE pic_list SET is_active = 1 WHERE name = ?`).bind(pic),
                buildAuditStatement(env, user, 'UPDATE', 'pic', pic, existingP, afterP, { source: 'import_validation' }, 'SUCCESS', true),
              ],
            });
          }
        } else {
          const snap = { name: pic, is_active: 1 };
          insertGroups.push({
            stmts: [
              env.DB.prepare(`INSERT INTO pic_list (name, is_active) VALUES (?, 1)`).bind(pic),
              buildAuditStatement(env, user, 'CREATE', 'pic', pic, null, snap, { source: 'import_validation' }, 'SUCCESS', true),
            ],
          });
        }
      }

      if (activePics.length > 0) {
        const picsToDeactivate = (existingPicsRes.results || []).filter(p => !activePicSet.has(p.name) && Number(p.is_active) !== 0);
        for (const p of picsToDeactivate) {
          const afterP = { ...p, is_active: 0 };
          wipeGroups.push({
            stmts: [
              env.DB.prepare(`UPDATE pic_list SET is_active = 0 WHERE name = ?`).bind(p.name),
              buildAuditStatement(env, user, 'UPDATE', 'pic', p.name, p, afterP, { source: 'import_validation_deactivate' }, 'SUCCESS', true),
            ],
          });
        }
      }

      // 3. Validation Options Differential Sync (Stage new options first, remove stale in final wipe chunk)
      const desiredOptions = [];
      const desiredKeySet = new Set();
      const addDesired = (cat, val) => {
        const k = `${cat}\u0000${val}`;
        if (!desiredKeySet.has(k)) {
          desiredKeySet.add(k);
          desiredOptions.push({ category: cat, value: val, key: k });
        }
      };
      for (const pic of activePics) addDesired('pic', pic);
      for (const keg of kegiatanSet) addDesired('activity', keg);
      for (const q of quartalSet) addDesired('quarter', q);
      for (const pkwt of pkwtSet) addDesired('pkwt', pkwt);

      const existingOptsRes = await env.DB.prepare('SELECT * FROM validation_options').all();
      const existingOptsMap = new Map();
      (existingOptsRes.results || []).forEach(o => existingOptsMap.set(`${o.category}\u0000${o.value}`, o));

      for (const opt of desiredOptions) {
        if (!existingOptsMap.has(opt.key)) {
          const snap = { category: opt.category, value: opt.value };
          insertGroups.push({
            stmts: [
              env.DB.prepare(`INSERT OR IGNORE INTO validation_options (category, value) VALUES (?, ?)`).bind(opt.category, opt.value),
              buildCreateAuditStatement(env, user, 'validation', `${opt.category}:${opt.value}`, snap, { source: 'import_validation' }),
            ],
          });
        }
      }

      const staleOpts = (existingOptsRes.results || []).filter(o => !desiredKeySet.has(`${o.category}\u0000${o.value}`));
      if (staleOpts.length > 0) {
        wipeGroups.push(...buildStagedWipeGroups(env, user, 'validation_options', 'validation', staleOpts, 'id', { source: 'import_validation_prune' }));
      }

      return {
        insertGroups,
        wipeGroups,
        liveWipePlan: true,
        finalCounts: { inserted: totalItems, updated: 0, skipped: 0 },
      };
    },
    buildFinalResult: () => ({ inserted: totalItems, skipped: 0 }),
  });
}

// ─── Employees ───────────────────────────────────────────────────────────────
async function importEmployees(rows, onDuplicate, env, user, origin, importMeta = {}) {
  return runStandardJournaledImport(env, user, origin, {
    module: 'employees',
    rows,
    onDuplicate,
    importMeta,
    setup: async () => {
      const mgmtBranchIds = await resolveManagementBranchIds(env.DB);
      const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1 AND deleted_at IS NULL').all();
      const matchBranch = makeBranchMatcher(bRows.results);
      const existing = await env.DB.prepare(
        `SELECT *, LOWER(TRIM(full_name)) as hash_key FROM employees
         WHERE full_name IS NOT NULL AND deleted_at IS NULL
           AND division IN ('FACILITY CARE', 'SECURITY')
           AND ${getManagementBranchSubquery('branch_id')}`
      ).all();
      const existingMap = new Map();
      (existing.results || []).forEach(e => {
        if (e.hash_key) existingMap.set(e.hash_key, e);
      });
      return { matchBranch, mgmtBranchIds, existingMap };
    },
    processRow: (row, { matchBranch, mgmtBranchIds, existingMap }, { onDuplicate: dupMode, seenNewKeys }) => {
      const full_name = safeStr(row.full_name);
      if (!full_name) return { action: 'skip' };

      const branch_id = Number(row.branch_id || matchBranch(row.branch_name));
      // Phase 5.66.10A.1 / 5.66.19: Employee import strictly restricted to Management area (fail-closed)
      if (!Number.isInteger(branch_id) || !mgmtBranchIds.has(branch_id)) {
        return { action: 'skip' };
      }
      const rawDiv = safeStr(row.division) || 'FACILITY CARE';
      const divUpper = rawDiv.toUpperCase();
      if (divUpper !== 'FACILITY CARE' && divUpper !== 'SECURITY') {
        return { action: 'skip' };
      }
      const division = divUpper;

      const key = full_name.toLowerCase().trim();
      const phone = safeStr(row.phone);
      const join_date = safeDate(row.join_date);
      const target_pindah_os = safeDate(row.target_pindah_os || row['Target Pindah OS']);
      const target_selesai_os = safeDate(row.target_selesai_os || row['Target Selesai OS']);
      const status = safeStr(row.status) || '';
      const notes = safeStr(row.notes);

      if (existingMap.has(key)) {
        const oldRow = existingMap.get(key);
        if (dupMode === 'update') {
          const afterRow = {
            ...oldRow,
            branch_id,
            division,
            phone,
            join_date,
            status,
            notes,
            target_pindah_os: target_pindah_os || oldRow.target_pindah_os,
            target_selesai_os: target_selesai_os || oldRow.target_selesai_os,
            row_version: (oldRow.row_version ?? 1) + 1
          };
          existingMap.set(key, afterRow);
          return {
            action: 'update',
            stmts: [
              env.DB.prepare(
                `UPDATE employees SET branch_id = ?, division = ?, phone = ?, join_date = ?, status = ?, notes = ?, target_pindah_os = COALESCE(?, target_pindah_os), target_selesai_os = COALESCE(?, target_selesai_os), row_version = COALESCE(row_version, 1) + 1, updated_at = datetime('now') WHERE id = ?`
              ).bind(branch_id, division, phone, join_date, status, notes, target_pindah_os, target_selesai_os, oldRow.id),
              buildAuditStatement(env, user, 'UPDATE', 'employees', oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true)
            ]
          };
        }
        return { action: 'skip' };
      }

      if (seenNewKeys.has(key)) return { action: 'skip' };
      seenNewKeys.add(key);

      const newSnapshot = { full_name, branch_id, division, phone, join_date, status, notes, target_pindah_os, target_selesai_os, row_version: 1 };
      return {
        action: 'insert',
        stmts: [
          env.DB.prepare(
            `INSERT INTO employees (full_name, branch_id, division, phone, join_date, status, notes, target_pindah_os, target_selesai_os, row_version, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1, datetime('now'))`
          ).bind(full_name, branch_id, division, phone, join_date, status, notes, target_pindah_os, target_selesai_os),
          buildCreateAuditStatement(env, user, 'employees', 'new', newSnapshot, { source: 'import' })
        ]
      };
    },
  });
}

// ─── Contracts ────────────────────────────────────────────────────────────────
async function importContracts(rows, onDuplicate, env, user, origin, importMeta = {}) {
  return runStandardJournaledImport(env, user, origin, {
    module: 'contracts',
    rows,
    onDuplicate,
    importMeta,
    setup: async () => {
      const mgmtBranchIds = await resolveManagementBranchIds(env.DB);
      const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1 AND deleted_at IS NULL').all();
      const matchBranch = makeBranchMatcher(bRows.results);
      const eRows = await env.DB.prepare('SELECT id, full_name, branch_id, division FROM employees WHERE deleted_at IS NULL').all();
      const findEmployeeByName = (str) => {
        if (!str) return null;
        const s = str.toLowerCase().trim();
        return (eRows.results || []).find(r => String(r.full_name || '').toLowerCase().trim() === s) || null;
      };
      const existing = await env.DB.prepare(
        `SELECT *, LOWER(TRIM(employee_name)) as hash_key FROM contracts
         WHERE employee_name IS NOT NULL AND deleted_at IS NULL
           AND division IN ('FACILITY CARE', 'SECURITY')
           AND ${getManagementBranchSubquery('branch_id')}`
      ).all();
      const existingMap = new Map();
      (existing.results || []).forEach(c => {
        if (c.hash_key) existingMap.set(c.hash_key, c);
      });
      return { matchBranch, mgmtBranchIds, findEmployeeByName, existingMap };
    },
    processRow: (row, { matchBranch, mgmtBranchIds, findEmployeeByName, existingMap }, { onDuplicate: dupMode, seenNewKeys }) => {
      const employee_name = safeStr(row.employee_name);
      if (!employee_name) return { action: 'skip' };

      const empRecord = findEmployeeByName(employee_name);
      if (empRecord) {
        const empDiv = String(empRecord.division || '').trim().toUpperCase();
        if (!mgmtBranchIds.has(Number(empRecord.branch_id)) || (empDiv !== 'FACILITY CARE' && empDiv !== 'SECURITY')) {
          return { action: 'skip' };
        }
      }

      const branch_id = Number(row.branch_id || matchBranch(row.branch_name));
      // Phase 5.66.10A.1 / 5.66.19: Contract import strictly restricted to Management area (fail-closed)
      if (!Number.isInteger(branch_id) || !mgmtBranchIds.has(branch_id)) {
        return { action: 'skip' };
      }
      const employee_id = empRecord ? empRecord.id : null;
      const rawDiv = safeStr(row.division) || (empRecord?.division || 'FACILITY CARE');
      const divUpper = rawDiv.toUpperCase();
      if (divUpper !== 'FACILITY CARE' && divUpper !== 'SECURITY') {
        return { action: 'skip' };
      }
      const division = divUpper;

      const start_date = safeDate(row.start_date);
      const end_date = safeDate(row.end_date);
      const key = employee_name.toLowerCase().trim();

      const contract_type = safeStr(row.contract_type);
      const pkwt_number = safeStr(row.pkwt_number);
      const status = safeStr(row.status) || '';
      const notes = safeStr(row.notes);

      if (existingMap.has(key)) {
        const oldRow = existingMap.get(key);
        if (dupMode === 'update') {
          const afterRow = { ...oldRow, employee_id, branch_id, start_date, division, end_date, contract_type, pkwt_number, status, notes };
          existingMap.set(key, afterRow);
          return {
            action: 'update',
            stmts: [
              env.DB.prepare(
                `UPDATE contracts SET employee_id = ?, branch_id = ?, start_date = ?, division = ?, end_date = ?, contract_type = ?, pkwt_number = ?, status = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
              ).bind(employee_id, branch_id, start_date, division, end_date, contract_type, pkwt_number, status, notes, oldRow.id),
              buildAuditStatement(env, user, 'UPDATE', 'contracts', oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true)
            ]
          };
        }
        return { action: 'skip' };
      }

      if (seenNewKeys.has(key)) return { action: 'skip' };
      seenNewKeys.add(key);

      const newSnapshot = { employee_id, employee_name, branch_id, division, start_date, end_date, contract_type, pkwt_number, status, notes };
      return {
        action: 'insert',
        stmts: [
          env.DB.prepare(
            `INSERT INTO contracts (employee_id, employee_name, branch_id, division, start_date, end_date, contract_type, pkwt_number, status, notes)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
          ).bind(employee_id, employee_name, branch_id, division, start_date, end_date, contract_type, pkwt_number, status, notes),
          buildCreateAuditStatement(env, user, 'contracts', 'new', newSnapshot, { source: 'import' })
        ]
      };
    },
  });
}

// ─── Relievers (Stage-First Insert Before Wipe) ───────────────────────────────
async function importRelievers(rows, onDuplicate, env, user, origin, importMeta = {}) {
  const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
  const matchBranch = makeBranchMatcher(bRows.results);

  const monthsToSync = new Set();
  const validRows = [];

  if (!rows || rows.length === 0) {
    return error('IMPORT FAILED: File Excel kosong atau tidak ada baris data.', 400, origin);
  }

  for (const row of rows) {
    const reliever_name = safeStr(row.reliever_name) || '-';
    const rawDate = safeDate(row.backup_date);

    if (!rawDate) {
      return error('IMPORT FAILED: Terdapat baris dengan backup_date kosong atau tidak valid. Proses dihentikan sebelum menyentuh database.', 400, origin);
    }

    let m = '';
    if (typeof rawDate === 'string' && rawDate.match(/^\d{4}-\d{2}-\d{2}/)) {
      m = rawDate.substring(0, 7);
    } else {
      const excelDays = Number(rawDate);
      if (!isNaN(excelDays)) {
        const d = new Date(Date.UTC(1899, 11, 30) + excelDays * 86400000);
        if (!isNaN(d.getTime())) m = d.toISOString().slice(0, 7);
      } else {
        const parts = rawDate.split(/[\/\-]/);
        if (parts.length === 3) m = `${parts[2]}-${parts[1].padStart(2, '0')}`;
      }
    }

    if (!m) {
      return error(`IMPORT FAILED: Gagal mengekstrak bulan dari tanggal: ${rawDate}. Proses dihentikan.`, 400, origin);
    }

    monthsToSync.add(m);
    validRows.push({ ...row, backup_date: rawDate, reliever_name });
  }

  if (monthsToSync.size === 0) {
    return error('IMPORT FAILED: Tidak ada bulan valid yang terdeteksi untuk di-sync.', 400, origin);
  }

  const monthsArray = Array.from(monthsToSync);
  const placeholders = monthsArray.map(() => '?').join(',');
  const monthExpr = `CASE
      WHEN length(backup_date) = 5 AND CAST(backup_date AS INTEGER) > 0 THEN
        strftime('%Y-%m', date('1970-01-01', '+' || (CAST(backup_date AS INTEGER) - 25569) || ' days'))
      WHEN length(backup_date) = 10 AND substr(backup_date, 3, 1) IN ('/', '-') AND substr(backup_date, 6, 1) IN ('/', '-') THEN
        substr(backup_date, 7, 4) || '-' || substr(backup_date, 4, 2)
      ELSE
        substr(backup_date, 1, 7)
    END`;

  return await executeJournaledImport(env, user, origin, {
    module: 'relievers',
    rows: validRows,
    onDuplicate: 'replace',
    extraScope: { months: [...monthsArray].sort() },
    importMeta,
    buildPlan: async ({ startRowIndex, initialCounts, initialMeta }) => {
      // Preserve exact original IDs to wipe BEFORE any new rows are inserted; never wipe newly inserted rows on recovery!
      let wipeIds = Array.isArray(initialMeta?.wipeIds) ? initialMeta.wipeIds : null;
      let existingWipeRows = [];
      if (wipeIds === null) {
        existingWipeRows = (await env.DB.prepare(
          `SELECT * FROM relievers WHERE (${monthExpr}) IN (${placeholders})`
        ).bind(...monthsArray).all()).results || [];
        wipeIds = existingWipeRows.map(r => r.id);
      } else if (wipeIds.length > 0) {
        const idPlaceholders = wipeIds.map(() => '?').join(',');
        existingWipeRows = (await env.DB.prepare(
          `SELECT * FROM relievers WHERE id IN (${idPlaceholders})`
        ).bind(...wipeIds).all()).results || [];
      }

      const insertGroups = [];
      let inserted = Number(initialCounts?.inserted || 0);

      for (let idx = startRowIndex; idx < validRows.length; idx++) {
        const row = validRows[idx];
        const branch_id = row.branch_id || matchBranch(row.branch_name);
        const original_fc_name = safeStr(row.original_fc_name);
        const period = safeStr(row.period);
        const completion_date = safeDate(row.completion_date);
        const reason = safeStr(row.reason);
        const shift = safeStr(row.shift);
        const status = safeStr(row.status) || 'Pending';
        const newSnapshot = {
          branch_id,
          original_fc_name,
          period,
          reliever_name: row.reliever_name,
          backup_date: row.backup_date,
          completion_date,
          reason,
          shift,
          status
        };
        inserted++;
        insertGroups.push({
          rowIndex: idx,
          counts: { inserted, skipped: 0, updated: 0 },
          stmts: [
            env.DB.prepare(
              `INSERT INTO relievers (branch_id, original_fc_name, period, reliever_name, backup_date, completion_date, reason, shift, status)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
            ).bind(branch_id, original_fc_name, period, row.reliever_name, row.backup_date, completion_date, reason, shift, status),
            buildCreateAuditStatement(env, user, 'relievers', 'new', newSnapshot, { source: 'import' })
          ]
        });
      }

      const wipeGroups = existingWipeRows.length > 0
        ? buildStagedWipeGroups(env, user, 'relievers', 'relievers', existingWipeRows, 'id', { source: 'import_wipe', months: monthsArray })
        : [];

      return {
        insertGroups,
        wipeGroups,
        liveWipePlan: true,
        meta: { wipeIds },
        finalCounts: { inserted, skipped: 0, updated: 0 },
      };
    },
    buildFinalResult: (counts) => ({
      inserted: counts.inserted,
      skipped: counts.skipped || 0,
      updated: counts.updated || 0,
    }),
  });
}

// ─── Schedule ─────────────────────────────────────────────────────────────────
async function importSchedule(rows, onDuplicate, env, user, origin, importMeta = {}) {
  return runStandardJournaledImport(env, user, origin, {
    module: 'schedule',
    rows,
    onDuplicate,
    importMeta,
    setup: async () => {
      const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
      const matchBranch = makeBranchMatcher(bRows.results);
      const existing = await env.DB.prepare("SELECT *, LOWER(TRIM(activity_type)) || '_' || LOWER(TRIM(period)) || '_' || COALESCE(branch_id, 'null') as hash_key FROM activity_schedule WHERE activity_type IS NOT NULL AND period IS NOT NULL").all();
      const existingMap = new Map();
      (existing.results || []).forEach(s => {
        if (s.hash_key) existingMap.set(s.hash_key, s);
      });
      return { matchBranch, existingMap };
    },
    processRow: (row, { matchBranch, existingMap }, { onDuplicate: dupMode, seenNewKeys }) => {
      const activity_type = safeStr(row.activity_type);
      if (!activity_type) return { action: 'skip' };

      const target_date = safeDate(row.target_date);
      const branch_id = row.branch_id || matchBranch(row.branch_name);
      const period = safeStr(row.period);
      const key = activity_type.toLowerCase().trim() + '_' + (period || '').toLowerCase().trim() + '_' + (branch_id || 'null');

      const pic = safeStr(row.pic);
      const opening_date = safeDate(row.opening_date);
      const completion_date = safeDate(row.completion_date);
      const status = safeStr(row.status) || '';
      const notes = safeStr(row.notes);

      if (existingMap.has(key)) {
        const oldRow = existingMap.get(key);
        if (dupMode === 'update') {
          const afterRow = { ...oldRow, period, pic, opening_date, target_date, completion_date, status, notes };
          existingMap.set(key, afterRow);
          return {
            action: 'update',
            stmts: [
              env.DB.prepare(
                `UPDATE activity_schedule SET period = ?, pic = ?, opening_date = ?, target_date = ?, completion_date = ?, status = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
              ).bind(period, pic, opening_date, target_date, completion_date, status, notes, oldRow.id),
              buildAuditStatement(env, user, 'UPDATE', 'schedule', oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true)
            ]
          };
        }
        return { action: 'skip' };
      }

      if (seenNewKeys.has(key)) return { action: 'skip' };
      seenNewKeys.add(key);

      const newSnapshot = { branch_id, activity_type, period, pic, opening_date, target_date, completion_date, status, notes };
      return {
        action: 'insert',
        stmts: [
          env.DB.prepare(
            `INSERT INTO activity_schedule (branch_id, activity_type, period, pic, opening_date, target_date, completion_date, status, notes)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`
          ).bind(branch_id, activity_type, period, pic, opening_date, target_date, completion_date, status, notes),
          buildCreateAuditStatement(env, user, 'schedule', 'new', newSnapshot, { source: 'import' })
        ]
      };
    },
  });
}

// ─── Issues ───────────────────────────────────────────────────────────────────
async function importIssues(rows, onDuplicate, env, user, origin, importMeta = {}) {
  return runStandardJournaledImport(env, user, origin, {
    module: 'issues',
    rows,
    onDuplicate,
    importMeta,
    setup: async () => {
      const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
      const matchBranch = makeBranchMatcher(bRows.results);
      const existing = await env.DB.prepare("SELECT *, LOWER(TRIM(complaint)) || '_' || report_date || '_' || COALESCE(branch_id, 'null') as hash_key FROM issues WHERE complaint IS NOT NULL AND report_date IS NOT NULL").all();
      const existingMap = new Map();
      (existing.results || []).forEach(i => {
        if (i.hash_key) existingMap.set(i.hash_key, i);
      });
      return { matchBranch, existingMap };
    },
    processRow: (row, { matchBranch, existingMap }, { onDuplicate: dupMode, seenNewKeys }) => {
      const complaint = safeStr(row.complaint);
      if (!complaint) return { action: 'skip' };

      const report_date = safeDate(row.report_date) || today();
      const branch_id = row.branch_id || matchBranch(row.branch_name);
      const key = complaint.toLowerCase().trim() + '_' + report_date + '_' + branch_id;

      const category = safeStr(row.category) || 'Lainnya';
      const source = safeStr(row.source);
      const employee_name = safeStr(row.employee_name);
      const fc_specialist = safeStr(row.fc_specialist);
      const solution = safeStr(row.solution);
      const status = safeStr(row.status) || '';
      const completion_date = safeDate(row.completion_date);

      let day_count = null;
      if (completion_date && report_date) {
        const d1 = new Date(report_date), d2 = new Date(completion_date);
        if (!isNaN(d1) && !isNaN(d2)) day_count = Math.floor((d2 - d1) / 86400000);
      }

      if (existingMap.has(key)) {
        const oldRow = existingMap.get(key);
        if (dupMode === 'update') {
          const afterRow = { ...oldRow, category, source, employee_name, fc_specialist, solution, status, completion_date, day_count };
          existingMap.set(key, afterRow);
          return {
            action: 'update',
            stmts: [
              env.DB.prepare(
                `UPDATE issues SET category = ?, source = ?, employee_name = ?, fc_specialist = ?, solution = ?, status = ?, completion_date = ?, day_count = ?, updated_at = datetime('now') WHERE id = ?`
              ).bind(category, source, employee_name, fc_specialist, solution, status, completion_date, day_count, oldRow.id),
              buildAuditStatement(env, user, 'UPDATE', 'issues', oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true)
            ]
          };
        }
        return { action: 'skip' };
      }

      if (seenNewKeys.has(key)) return { action: 'skip' };
      seenNewKeys.add(key);

      const newSnapshot = { report_date, branch_id, category, source, complaint, employee_name, fc_specialist, solution, status, completion_date, day_count };
      return {
        action: 'insert',
        stmts: [
          env.DB.prepare(
            `INSERT INTO issues (report_date, branch_id, category, source, complaint, employee_name, fc_specialist, solution, status, completion_date, day_count)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
          ).bind(report_date, branch_id, category, source, complaint, employee_name, fc_specialist, solution, status, completion_date, day_count),
          buildCreateAuditStatement(env, user, 'issues', 'new', newSnapshot, { source: 'import' })
        ]
      };
    },
  });
}

// ─── One on One ───────────────────────────────────────────────────────────────
async function importOneOnOne(rows, onDuplicate, env, user, origin, importMeta = {}) {
  return runStandardJournaledImport(env, user, origin, {
    module: 'one_on_one',
    rows,
    onDuplicate,
    importMeta,
    setup: async () => {
      const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
      const matchBranch = makeBranchMatcher(bRows.results);
      const existing = await env.DB.prepare("SELECT *, LOWER(TRIM(employee_name)) || '_' || meeting_date as hash_key FROM one_on_one WHERE employee_name IS NOT NULL AND meeting_date IS NOT NULL").all();
      const existingMap = new Map();
      (existing.results || []).forEach(o => {
        if (o.hash_key) existingMap.set(o.hash_key, o);
      });
      return { matchBranch, existingMap };
    },
    processRow: (row, { matchBranch, existingMap }, { onDuplicate: dupMode, seenNewKeys }) => {
      const employee_name = safeStr(row.employee_name) || '-';
      const meeting_date = safeDate(row.meeting_date) || today();
      const key = employee_name.toLowerCase().trim() + '_' + meeting_date;

      const branch_id = row.branch_id || matchBranch(row.branch_name);
      const pic = safeStr(row.pic);
      const problem = safeStr(row.problem) || '-';
      const solution = safeStr(row.solution);
      const status = safeStr(row.status) || '';
      const completion_date = safeDate(row.completion_date);
      const document_link = safeStr(row.document_link);

      let day_count = null;
      if (completion_date && meeting_date) {
        const d1 = new Date(meeting_date), d2 = new Date(completion_date);
        if (!isNaN(d1) && !isNaN(d2)) day_count = Math.floor((d2 - d1) / 86400000);
      }

      if (existingMap.has(key)) {
        const oldRow = existingMap.get(key);
        if (dupMode === 'update') {
          const afterRow = { ...oldRow, branch_id, pic, problem, solution, status, completion_date, day_count, document_link };
          existingMap.set(key, afterRow);
          return {
            action: 'update',
            stmts: [
              env.DB.prepare(
                `UPDATE one_on_one SET branch_id = ?, pic = ?, problem = ?, solution = ?, status = ?, completion_date = ?, day_count = ?, document_link = ?, updated_at = datetime('now') WHERE id = ?`
              ).bind(branch_id, pic, problem, solution, status, completion_date, day_count, document_link, oldRow.id),
              buildAuditStatement(env, user, 'UPDATE', 'one_on_one', oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true)
            ]
          };
        }
        return { action: 'skip' };
      }

      if (seenNewKeys.has(key)) return { action: 'skip' };
      seenNewKeys.add(key);

      const newSnapshot = { meeting_date, branch_id, employee_name, pic, problem, solution, status, completion_date, day_count, document_link };
      return {
        action: 'insert',
        stmts: [
          env.DB.prepare(
            `INSERT INTO one_on_one (meeting_date, branch_id, employee_name, pic, problem, solution, status, completion_date, day_count, document_link)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
          ).bind(meeting_date, branch_id, employee_name, pic, problem, solution, status, completion_date, day_count, document_link),
          buildCreateAuditStatement(env, user, 'one_on_one', 'new', newSnapshot, { source: 'import' })
        ]
      };
    },
  });
}

// ─── Training ─────────────────────────────────────────────────────────────────
async function importTraining(rows, onDuplicate, env, user, origin, importMeta = {}) {
  return runStandardJournaledImport(env, user, origin, {
    module: 'training',
    rows,
    onDuplicate,
    importMeta,
    setup: async () => {
      const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
      const matchBranch = makeBranchMatcher(bRows.results);
      const existing = await env.DB.prepare('SELECT * FROM training').all();
      const existingMap = new Map();
      (existing.results || []).forEach(t => {
        if (t.subject && t.training_date) {
          existingMap.set(t.subject.toLowerCase().trim() + '_' + t.training_date + '_' + (t.batch || '').toLowerCase().trim(), t);
        }
      });
      return { matchBranch, existingMap };
    },
    processRow: (row, { matchBranch, existingMap }, { onDuplicate: dupMode, seenNewKeys }) => {
      const subject = safeStr(row.subject);
      if (!subject) return { action: 'skip' };

      const training_date = safeDate(row.training_date) || today();
      const batch = safeStr(row.batch) || '';
      const key = subject.toLowerCase().trim() + '_' + training_date + '_' + batch.toLowerCase().trim();

      const branch_id = row.branch_id || matchBranch(row.branch_name);
      const participants = safeStr(row.participants);
      const trainer = safeStr(row.trainer);
      const score = row.score != null ? parseFloat(row.score) : null;
      const notes = safeStr(row.notes);

      if (existingMap.has(key)) {
        const oldRow = existingMap.get(key);
        if (dupMode === 'update') {
          const afterRow = { ...oldRow, branch_id, participants, trainer, score, notes };
          existingMap.set(key, afterRow);
          return {
            action: 'update',
            stmts: [
              env.DB.prepare(
                `UPDATE training SET branch_id = ?, participants = ?, trainer = ?, score = ?, notes = ?, updated_at = datetime('now') WHERE id = ?`
              ).bind(branch_id, participants, trainer, score, notes, oldRow.id),
              buildAuditStatement(env, user, 'UPDATE', 'training', oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true)
            ]
          };
        }
        return { action: 'skip' };
      }

      if (seenNewKeys.has(key)) return { action: 'skip' };
      seenNewKeys.add(key);

      const newSnapshot = { training_date, batch, subject, participants, branch_id, trainer, score, notes };
      return {
        action: 'insert',
        stmts: [
          env.DB.prepare(
            `INSERT INTO training (training_date, batch, subject, participants, branch_id, trainer, score, notes)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
          ).bind(training_date, batch, subject, participants, branch_id, trainer, score, notes),
          buildCreateAuditStatement(env, user, 'training', 'new', newSnapshot, { source: 'import' })
        ]
      };
    },
  });
}

// ─── Checklist ────────────────────────────────────────────────────────────────
async function importChecklist(rows, onDuplicate, env, user, origin, importMeta = {}) {
  return runStandardJournaledImport(env, user, origin, {
    module: 'checklist',
    rows,
    onDuplicate,
    importMeta,
    setup: async () => {
      const existing = await env.DB.prepare('SELECT * FROM master_checklist WHERE deleted_at IS NULL').all();
      const existingMap = new Map();
      (existing.results || []).forEach(c => {
        existingMap.set(c.name.toLowerCase().trim(), c);
      });
      return { existingMap };
    },
    processRow: (row, { existingMap }, { onDuplicate: dupMode, seenNewKeys }) => {
      const name = safeStr(row.name) || '-';
      const key = name.toLowerCase().trim();
      const category = safeStr(row.category) || 'Umum';
      const document_link = safeStr(row.document_link);
      const description = safeStr(row.description);

      if (existingMap.has(key)) {
        const oldRow = existingMap.get(key);
        if (dupMode === 'update') {
          const afterRow = { ...oldRow, category, document_link, description };
          existingMap.set(key, afterRow);
          return {
            action: 'update',
            stmts: [
              env.DB.prepare(
                `UPDATE master_checklist SET category = ?, document_link = ?, description = ?, updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ? AND deleted_at IS NULL`
              ).bind(category, document_link, description, oldRow.id),
              buildAuditStatement(env, user, 'UPDATE', 'checklist', oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true)
            ]
          };
        }
        return { action: 'skip' };
      }

      if (seenNewKeys.has(key)) return { action: 'skip' };
      seenNewKeys.add(key);

      const newSnapshot = { name, category, document_link, description };
      return {
        action: 'insert',
        stmts: [
          env.DB.prepare(
            `INSERT INTO master_checklist (name, category, document_link, description) VALUES (?, ?, ?, ?)`
          ).bind(name, category, document_link, description),
          buildCreateAuditStatement(env, user, 'checklist', 'new', newSnapshot, { source: 'import' })
        ]
      };
    },
  });
}

// ─── Forms ────────────────────────────────────────────────────────────────────
async function importForms(rows, onDuplicate, env, user, origin, importMeta = {}) {
  return runStandardJournaledImport(env, user, origin, {
    module: 'forms',
    rows,
    onDuplicate,
    importMeta,
    setup: async () => {
      const existing = await env.DB.prepare('SELECT * FROM master_forms').all();
      const existingMap = new Map();
      (existing.results || []).forEach(f => {
        existingMap.set(f.name.toLowerCase().trim(), f);
      });
      return { existingMap };
    },
    processRow: (row, { existingMap }, { onDuplicate: dupMode, seenNewKeys }) => {
      const name = safeStr(row.name);
      if (!name) return { action: 'skip' };

      const key = name.toLowerCase().trim();
      const category = safeStr(row.category) || 'Umum';
      const document_link = safeStr(row.document_link);
      const description = safeStr(row.description);

      if (existingMap.has(key)) {
        const oldRow = existingMap.get(key);
        if (dupMode === 'update') {
          const afterRow = { ...oldRow, category, document_link, description };
          existingMap.set(key, afterRow);
          return {
            action: 'update',
            stmts: [
              env.DB.prepare(
                `UPDATE master_forms SET category = ?, document_link = ?, description = ?, updated_at = datetime('now') WHERE id = ?`
              ).bind(category, document_link, description, oldRow.id),
              buildAuditStatement(env, user, 'UPDATE', 'forms', oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true)
            ]
          };
        }
        return { action: 'skip' };
      }

      if (seenNewKeys.has(key)) return { action: 'skip' };
      seenNewKeys.add(key);

      const newSnapshot = { name, category, document_link, description };
      return {
        action: 'insert',
        stmts: [
          env.DB.prepare(
            `INSERT INTO master_forms (name, category, document_link, description) VALUES (?, ?, ?, ?)`
          ).bind(name, category, document_link, description),
          buildCreateAuditStatement(env, user, 'forms', 'new', newSnapshot, { source: 'import' })
        ]
      };
    },
  });
}

// ─── SOP ──────────────────────────────────────────────────────────────────────
async function importSop(rows, onDuplicate, env, user, origin, importMeta = {}) {
  return runStandardJournaledImport(env, user, origin, {
    module: 'sop',
    rows,
    onDuplicate,
    importMeta,
    setup: async () => {
      const existing = await env.DB.prepare("SELECT * FROM sop WHERE deleted_at IS NULL AND (category != 'sop_bundle' OR category IS NULL)").all();
      const existingMap = new Map();
      (existing.results || []).forEach(s => {
        existingMap.set(s.name.toLowerCase().trim(), s);
      });
      return { existingMap };
    },
    processRow: (row, { existingMap }, { onDuplicate: dupMode, seenNewKeys }) => {
      const name = safeStr(row.name);
      if (!name) return { action: 'skip' };

      const key = name.toLowerCase().trim();
      const category = safeStr(row.category) || 'Umum';
      const document_link = safeStr(row.document_link);
      const version = safeStr(row.version) || '1.0';
      const effective_date = safeDate(row.effective_date);
      const notes = safeStr(row.notes);

      if (existingMap.has(key)) {
        const oldRow = existingMap.get(key);
        if (dupMode === 'update') {
          const afterRow = { ...oldRow, category, document_link, version, effective_date, notes };
          existingMap.set(key, afterRow);
          return {
            action: 'update',
            stmts: [
              env.DB.prepare(
                `UPDATE sop SET category = ?, document_link = ?, version = ?, effective_date = ?, notes = ?, updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ? AND deleted_at IS NULL`
              ).bind(category, document_link, version, effective_date, notes, oldRow.id),
              buildAuditStatement(env, user, 'UPDATE', 'sop', oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true)
            ]
          };
        }
        return { action: 'skip' };
      }

      if (seenNewKeys.has(key)) return { action: 'skip' };
      seenNewKeys.add(key);

      const newSnapshot = { name, category, document_link, version, effective_date, notes };
      return {
        action: 'insert',
        stmts: [
          env.DB.prepare(
            `INSERT INTO sop (name, category, document_link, version, effective_date, notes) VALUES (?, ?, ?, ?, ?, ?)`
          ).bind(name, category, document_link, version, effective_date, notes),
          buildCreateAuditStatement(env, user, 'sop', 'new', newSnapshot, { source: 'import' })
        ]
      };
    },
  });
}

// ─── Inspection ───────────────────────────────────────────────────────────────
async function importInspection(rows, onDuplicate, env, user, origin, importMeta = {}) {
  return runStandardJournaledImport(env, user, origin, {
    module: 'reports/inspection',
    rows,
    onDuplicate,
    importMeta,
    setup: async () => {
      const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
      const matchBranch = makeBranchMatcher(bRows.results);
      const existing = await env.DB.prepare('SELECT * FROM inspection_reports WHERE deleted_at IS NULL').all();
      const existingMap = new Map();
      (existing.results || []).forEach(i => {
        existingMap.set(i.branch_id + '_' + (i.period || '').toLowerCase().trim(), i);
      });
      return { matchBranch, existingMap };
    },
    processRow: (row, { matchBranch, existingMap }, { onDuplicate: dupMode, seenNewKeys }) => {
      let branch_id = row.branch_id || null;
      if (!branch_id) {
        const branch_name = safeStr(row.branch_name);
        if (!branch_name) return { action: 'skip' };
        branch_id = matchBranch(branch_name);
      }
      if (!branch_id) return { action: 'skip' };
      const period = safeStr(row.period) || '-';
      const inspection_date = safeDate(row.inspection_date) || today();
      const key = branch_id + '_' + period.toLowerCase().trim();

      const status = safeStr(row.status) || '';
      const fc_score = row.fc_score != null ? parseFloat(row.fc_score) : null;
      const spv_score = row.spv_score != null ? parseFloat(row.spv_score) : null;
      const document_link = safeStr(row.document_link);
      const notes = safeStr(row.notes);

      if (existingMap.has(key)) {
        const oldRow = existingMap.get(key);
        if (dupMode === 'update') {
          const afterRow = { ...oldRow, inspection_date, status, fc_score, spv_score, document_link, notes };
          existingMap.set(key, afterRow);
          return {
            action: 'update',
            stmts: [
              env.DB.prepare(
                `UPDATE inspection_reports SET inspection_date = ?, status = ?, fc_score = ?, spv_score = ?, document_link = ?, notes = ?, updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ? AND deleted_at IS NULL`
              ).bind(inspection_date, status, fc_score, spv_score, document_link, notes, oldRow.id),
              buildAuditStatement(env, user, 'UPDATE', 'reports/inspection', oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true)
            ]
          };
        }
        return { action: 'skip' };
      }

      if (seenNewKeys.has(key)) return { action: 'skip' };
      seenNewKeys.add(key);

      const newSnapshot = { branch_id, period, inspection_date, status, fc_score, spv_score, document_link, notes };
      return {
        action: 'insert',
        stmts: [
          env.DB.prepare(
            `INSERT INTO inspection_reports (branch_id, period, inspection_date, status, fc_score, spv_score, document_link, notes)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
          ).bind(branch_id, period, inspection_date, status, fc_score, spv_score, document_link, notes),
          buildCreateAuditStatement(env, user, 'reports/inspection', 'new', newSnapshot, { source: 'import' })
        ]
      };
    },
  });
}

// ─── Cleaning (GC/DC) ─────────────────────────────────────────────────────────
async function importCleaning(rows, onDuplicate, env, user, origin, importMeta = {}) {
  return runStandardJournaledImport(env, user, origin, {
    module: 'reports/cleaning',
    rows,
    onDuplicate,
    importMeta,
    setup: async () => {
      const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
      const matchBranch = makeBranchMatcher(bRows.results);
      const existing = await env.DB.prepare('SELECT * FROM cleaning_reports WHERE deleted_at IS NULL').all();
      const existingMap = new Map();
      (existing.results || []).forEach(c => {
        existingMap.set(c.branch_id + '_' + c.activity_type.toLowerCase().trim() + '_' + (c.period || '').toLowerCase().trim(), c);
      });
      return { matchBranch, existingMap };
    },
    processRow: (row, { matchBranch, existingMap }, { onDuplicate: dupMode, seenNewKeys }) => {
      let branch_id = row.branch_id || null;
      if (!branch_id) {
        const branch_name = safeStr(row.branch_name);
        if (!branch_name) return { action: 'skip' };
        branch_id = matchBranch(branch_name);
      }
      if (!branch_id) return { action: 'skip' };
      const activity_type = safeStr(row.activity_type) || 'General Cleaning';
      const period = safeStr(row.period) || '-';
      const activity_date = safeDate(row.activity_date) || today();
      const key = branch_id + '_' + activity_type.toLowerCase().trim() + '_' + period.toLowerCase().trim();

      const status = safeStr(row.status) || '';
      const document_link = safeStr(row.document_link);
      const notes = safeStr(row.notes);

      if (existingMap.has(key)) {
        const oldRow = existingMap.get(key);
        if (dupMode === 'update') {
          const afterRow = { ...oldRow, activity_date, status, document_link, notes };
          existingMap.set(key, afterRow);
          return {
            action: 'update',
            stmts: [
              env.DB.prepare(
                `UPDATE cleaning_reports SET activity_date = ?, status = ?, document_link = ?, notes = ?, updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ? AND deleted_at IS NULL`
              ).bind(activity_date, status, document_link, notes, oldRow.id),
              buildAuditStatement(env, user, 'UPDATE', 'reports/cleaning', oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true)
            ]
          };
        }
        return { action: 'skip' };
      }

      if (seenNewKeys.has(key)) return { action: 'skip' };
      seenNewKeys.add(key);

      const newSnapshot = { branch_id, activity_type, period, activity_date, status, document_link, notes };
      return {
        action: 'insert',
        stmts: [
          env.DB.prepare(
            `INSERT INTO cleaning_reports (branch_id, activity_type, period, activity_date, status, document_link, notes)
             VALUES (?, ?, ?, ?, ?, ?, ?)`
          ).bind(branch_id, activity_type, period, activity_date, status, document_link, notes),
          buildCreateAuditStatement(env, user, 'reports/cleaning', 'new', newSnapshot, { source: 'import' })
        ]
      };
    },
  });
}

// ─── Fogging ──────────────────────────────────────────────────────────────────
async function importFogging(rows, onDuplicate, env, user, origin, importMeta = {}) {
  return runStandardJournaledImport(env, user, origin, {
    module: 'reports/fogging',
    rows,
    onDuplicate,
    importMeta,
    setup: async () => {
      const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
      const matchBranch = makeBranchMatcher(bRows.results);
      const existing = await env.DB.prepare('SELECT * FROM fogging_reports WHERE deleted_at IS NULL').all();
      const existingMap = new Map();
      (existing.results || []).forEach(f => {
        existingMap.set(f.branch_id + '_' + (f.period || '').toLowerCase().trim(), f);
      });
      return { matchBranch, existingMap };
    },
    processRow: (row, { matchBranch, existingMap }, { onDuplicate: dupMode, seenNewKeys }) => {
      const branch_name = safeStr(row.branch_name);
      if (!branch_name && !row.branch_id) return { action: 'skip' };

      const branch_id = row.branch_id || matchBranch(branch_name);
      const period = safeStr(row.period) || '-';
      const activity_date = safeDate(row.activity_date) || today();
      const key = branch_id + '_' + period.toLowerCase().trim();

      const status = safeStr(row.status) || '';
      const document_link = safeStr(row.document_link);
      const notes = safeStr(row.notes);

      if (existingMap.has(key)) {
        const oldRow = existingMap.get(key);
        if (dupMode === 'update') {
          const afterRow = { ...oldRow, activity_date, status, document_link, notes };
          existingMap.set(key, afterRow);
          return {
            action: 'update',
            stmts: [
              env.DB.prepare(
                `UPDATE fogging_reports SET activity_date = ?, status = ?, document_link = ?, notes = ?, updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ? AND deleted_at IS NULL`
              ).bind(activity_date, status, document_link, notes, oldRow.id),
              buildAuditStatement(env, user, 'UPDATE', 'reports/fogging', oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true)
            ]
          };
        }
        return { action: 'skip' };
      }

      if (seenNewKeys.has(key)) return { action: 'skip' };
      seenNewKeys.add(key);

      const newSnapshot = { branch_id, activity_type: 'Fogging', period, activity_date, status, document_link, notes };
      return {
        action: 'insert',
        stmts: [
          env.DB.prepare(
            `INSERT INTO fogging_reports (branch_id, activity_type, period, activity_date, status, document_link, notes)
             VALUES (?, ?, ?, ?, ?, ?, ?)`
          ).bind(branch_id, 'Fogging', period, activity_date, status, document_link, notes),
          buildCreateAuditStatement(env, user, 'reports/fogging', 'new', newSnapshot, { source: 'import' })
        ]
      };
    },
  });
}

// ─── Basecamp ─────────────────────────────────────────────────────────────────
async function importBasecamp(rows, onDuplicate, env, user, origin, importMeta = {}) {
  return runStandardJournaledImport(env, user, origin, {
    module: 'reports/basecamp',
    rows,
    onDuplicate,
    importMeta,
    setup: async () => {
      const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
      const matchBranch = makeBranchMatcher(bRows.results);
      const existing = await env.DB.prepare('SELECT * FROM basecamp_reports WHERE deleted_at IS NULL').all();
      const existingMap = new Map();
      (existing.results || []).forEach(b => {
        if (b.info_date) {
          existingMap.set(b.branch_id + '_' + b.problem.toLowerCase().trim() + '_' + b.info_date, b);
        }
      });
      return { matchBranch, existingMap };
    },
    processRow: (row, { matchBranch, existingMap }, { onDuplicate: dupMode, seenNewKeys }) => {
      const problem = safeStr(row.problem);
      if (!problem) return { action: 'skip' };

      const branch_id = row.branch_id || matchBranch(row.branch_name);
      const info_date = safeDate(row.info_date) || today();
      const key = branch_id + '_' + problem.toLowerCase().trim() + '_' + info_date;

      const pic = safeStr(row.pic);
      const done_date = safeDate(row.done_date);
      const status = safeStr(row.status) || '';
      const notes = safeStr(row.notes);

      if (existingMap.has(key)) {
        const oldRow = existingMap.get(key);
        if (dupMode === 'update') {
          const afterRow = { ...oldRow, pic, done_date, status, notes };
          existingMap.set(key, afterRow);
          return {
            action: 'update',
            stmts: [
              env.DB.prepare(
                `UPDATE basecamp_reports SET pic = ?, done_date = ?, status = ?, notes = ?, updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ? AND deleted_at IS NULL`
              ).bind(pic, done_date, status, notes, oldRow.id),
              buildAuditStatement(env, user, 'UPDATE', 'reports/basecamp', oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true)
            ]
          };
        }
        return { action: 'skip' };
      }

      if (seenNewKeys.has(key)) return { action: 'skip' };
      seenNewKeys.add(key);

      const newSnapshot = { branch_id, problem, pic, info_date, done_date, status, notes };
      return {
        action: 'insert',
        stmts: [
          env.DB.prepare(
            `INSERT INTO basecamp_reports (branch_id, problem, pic, info_date, done_date, status, notes)
             VALUES (?, ?, ?, ?, ?, ?, ?)`
          ).bind(branch_id, problem, pic, info_date, done_date, status, notes),
          buildCreateAuditStatement(env, user, 'reports/basecamp', 'new', newSnapshot, { source: 'import' })
        ]
      };
    },
  });
}

// ─── Supply / Chemical ────────────────────────────────────────────────────────
async function importSupply(rows, onDuplicate, env, user, origin, importMeta = {}) {
  return runStandardJournaledImport(env, user, origin, {
    module: 'reports/supply',
    rows,
    onDuplicate,
    importMeta,
    setup: async () => {
      const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
      const matchBranch = makeBranchMatcher(bRows.results);
      const existing = await env.DB.prepare('SELECT * FROM supply_requests WHERE deleted_at IS NULL').all();
      const existingMap = new Map();
      (existing.results || []).forEach(s => {
        if (s.submitter_name && s.submitted_at) {
          existingMap.set(s.submitter_name.toLowerCase().trim() + '_' + s.submitted_at, s);
        }
      });
      return { branches: bRows.results || [], matchBranch, existingMap };
    },
    processRow: (row, { branches, matchBranch, existingMap }, { onDuplicate: dupMode, seenNewKeys }) => {
      const submitter_name = safeStr(row.submitter_name) || '-';
      const submitted_at = safeDate(row.submitted_at) || today();
      const key = submitter_name.toLowerCase().trim() + '_' + submitted_at;

      const branch_id = row.branch_id || matchBranch(row.branch_name);
      const branchObj = branch_id ? branches.find(b => b.id === branch_id) : null;
      const branch_name = branchObj ? branchObj.full_name : safeStr(row.branch_name);
      const tools_items = safeStr(row.tools_items);
      const tools_quantity = safeStr(row.tools_quantity);
      const chemical_items = safeStr(row.chemical_items);
      const chemical_quantity = safeStr(row.chemical_quantity);
      const additional_notes = safeStr(row.additional_notes);
      const status = safeStr(row.status) || '';

      if (existingMap.has(key)) {
        const oldRow = existingMap.get(key);
        if (dupMode === 'update') {
          const afterRow = { ...oldRow, branch_id, branch_name, tools_items, tools_quantity, chemical_items, chemical_quantity, additional_notes, status };
          existingMap.set(key, afterRow);
          return {
            action: 'update',
            stmts: [
              env.DB.prepare(
                `UPDATE supply_requests SET branch_id = ?, branch_name = ?, tools_items = ?, tools_quantity = ?, chemical_items = ?, chemical_quantity = ?, additional_notes = ?, status = ?, processed_at = datetime('now'), updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ? AND deleted_at IS NULL`
              ).bind(branch_id, branch_name, tools_items, tools_quantity, chemical_items, chemical_quantity, additional_notes, status, oldRow.id),
              buildAuditStatement(env, user, 'UPDATE', 'reports/supply', oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true)
            ]
          };
        }
        return { action: 'skip' };
      }

      if (seenNewKeys.has(key)) return { action: 'skip' };
      seenNewKeys.add(key);

      const newSnapshot = { submitted_at, submitter_name, branch_id, branch_name, tools_items, tools_quantity, chemical_items, chemical_quantity, additional_notes, status };
      return {
        action: 'insert',
        stmts: [
          env.DB.prepare(
            `INSERT INTO supply_requests (submitted_at, submitter_name, branch_id, branch_name, tools_items, tools_quantity, chemical_items, chemical_quantity, additional_notes, status)
             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
          ).bind(submitted_at, submitter_name, branch_id, branch_name, tools_items, tools_quantity, chemical_items, chemical_quantity, additional_notes, status),
          buildCreateAuditStatement(env, user, 'reports/supply', 'new', newSnapshot, { source: 'import' })
        ]
      };
    },
  });
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
      let rows = res.results || [];
      if (table === 'users') {
        rows = rows.map(({ password_hash, ...rest }) => rest);
      }
      return { table, rows };
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

async function triggerCalendarSync(env, user, origin) {
  const { getSyncStatements } = await import('../utils/calendar.js');
  const allStmts = [];

  // Clear existing calendar_events table
  allStmts.push(env.DB.prepare('DELETE FROM calendar_events'));

  // 1. Contracts (Phase 5.66.19: strictly Management 000 + official divisions)
  const contr = await env.DB.prepare(
    `SELECT c.id, c.end_date, c.status, c.branch_id, COALESCE(e.full_name, c.employee_name, '?') as empName
     FROM contracts c LEFT JOIN employees e ON c.employee_id = e.id
     WHERE c.deleted_at IS NULL
       AND c.division IN ('FACILITY CARE', 'SECURITY')
       AND ${getManagementBranchSubquery('c.branch_id')}`
  ).all();
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
  const clean = await env.DB.prepare(`SELECT * FROM cleaning_reports WHERE deleted_at IS NULL`).all();
  (clean.results || []).forEach(c => {
    allStmts.push(...getSyncStatements(env.DB, 'cleaning', c.id, c));
  });

  // 8. Inspection Reports
  const insp = await env.DB.prepare(`SELECT * FROM inspection_reports WHERE deleted_at IS NULL`).all();
  (insp.results || []).forEach(i => {
    allStmts.push(...getSyncStatements(env.DB, 'inspection', i.id, i));
  });

  // 9. Fogging Reports
  const fog = await env.DB.prepare(`SELECT * FROM fogging_reports WHERE deleted_at IS NULL`).all();
  (fog.results || []).forEach(f => {
    allStmts.push(...getSyncStatements(env.DB, 'fogging', f.id, f));
  });

  // 10. Basecamp Reports
  const base = await env.DB.prepare(`SELECT * FROM basecamp_reports WHERE deleted_at IS NULL`).all();
  (base.results || []).forEach(b => {
    allStmts.push(...getSyncStatements(env.DB, 'basecamp', b.id, b));
  });

  // 11. Supply Requests
  const supply = await env.DB.prepare(`SELECT * FROM supply_requests WHERE deleted_at IS NULL`).all();
  (supply.results || []).forEach(s => {
    allStmts.push(...getSyncStatements(env.DB, 'supply', s.id, s));
  });

  const syncedCount = Math.max(0, allStmts.length - 1);
  allStmts.push(buildAuditStatement(env, user, 'UPDATE', 'import', 'sync-calendar', null, null, { synced_events: syncedCount }, 'SUCCESS', false));

  // Batch execute
  if (allStmts.length > 0) {
    const chunkSize = 80;
    for (let i = 0; i < allStmts.length; i += chunkSize) {
      await env.DB.batch(allStmts.slice(i, i + chunkSize));
    }
  }

  return ok({ message: `Berhasil sinkronisasi ${syncedCount} event kalender dari data lama.` }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function importSP(rows, onDuplicate, env, user, origin, importMeta = {}) {
  return runStandardJournaledImport(env, user, origin, {
    module: 'sp',
    rows,
    onDuplicate,
    importMeta,
    setup: async () => {
      const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
      const matchBranch = makeBranchMatcher(bRows.results);
      const empRows = await env.DB.prepare('SELECT id, LOWER(TRIM(full_name)) as name_lower FROM employees WHERE deleted_at IS NULL').all();
      const empCounts = new Map();
      const empIdMap = new Map();
      (empRows.results || []).forEach(e => {
        empCounts.set(e.name_lower, (empCounts.get(e.name_lower) || 0) + 1);
        empIdMap.set(e.name_lower, e.id);
      });
      const existingAll = await env.DB.prepare('SELECT * FROM sp_data').all();
      const existingMap = new Map();
      for (const row of (existingAll.results || [])) {
        if (row.branch_id && row.tanggal && row.employee_name && row.sp_type) {
          const k = `${row.branch_id}_${row.tanggal}_${String(row.employee_name).toLowerCase().trim()}_${String(row.sp_type).toLowerCase().trim()}`;
          if (!existingMap.has(k)) existingMap.set(k, row);
        }
      }
      return { matchBranch, empCounts, empIdMap, existingMap };
    },
    processRow: (item, { matchBranch, empCounts, empIdMap, existingMap }, { onDuplicate: dupMode, seenNewKeys }) => {
      const branch_id = item.branch_id || matchBranch(item.branch_name);
      if (!item.tanggal || !item.employee_name || !branch_id || !item.sp_type) return { action: 'skip' };
      const status = item.status || 'Aktif';

      const empNameClean = String(item.employee_name).toLowerCase().trim();
      let empId = null;
      if (item.employee_id) {
        empId = Number(item.employee_id) || null;
      } else if (empCounts.get(empNameClean) === 1) {
        empId = empIdMap.get(empNameClean);
      }

      const spKey = `${branch_id}_${item.tanggal}_${empNameClean}_${String(item.sp_type).toLowerCase().trim()}`;
      if (dupMode === 'update' && existingMap.has(spKey)) {
        const oldRow = existingMap.get(spKey);
        const afterRow = {
          ...oldRow,
          status,
          document_link: item.document_link || null,
          employee_id: empId ?? oldRow.employee_id,
          row_version: (oldRow.row_version ?? 1) + 1
        };
        existingMap.set(spKey, afterRow);
        return {
          action: 'update',
          stmts: [
            env.DB.prepare(
              `UPDATE sp_data SET status = ?, document_link = ?, employee_id = COALESCE(?, employee_id),
               row_version = COALESCE(row_version, 1) + 1, updated_at = datetime('now') WHERE id = ?`
            ).bind(status, item.document_link || null, empId, oldRow.id),
            buildAuditStatement(env, user, 'UPDATE', 'sp', oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true)
          ]
        };
      }

      if (existingMap.has(spKey) && dupMode === 'skip') return { action: 'skip' };
      if (seenNewKeys.has(spKey)) return { action: 'skip' };
      seenNewKeys.add(spKey);

      const newSnapshot = {
        tanggal: item.tanggal,
        employee_id: empId,
        employee_name: item.employee_name,
        branch_id,
        sp_type: item.sp_type,
        status,
        document_link: item.document_link || null,
        row_version: 1
      };
      return {
        action: 'insert',
        stmts: [
          env.DB.prepare(
            `INSERT INTO sp_data (tanggal, employee_id, employee_name, branch_id, sp_type, status, document_link, row_version, updated_at) 
             VALUES (?, ?, ?, ?, ?, ?, ?, 1, datetime('now'))`
          ).bind(
            item.tanggal, empId, item.employee_name, branch_id, item.sp_type, status, item.document_link || null
          ),
          buildCreateAuditStatement(env, user, 'sp', 'new', newSnapshot, { source: 'import' })
        ]
      };
    },
    formatResult: (counts) => ({
      message: 'Berhasil import SP',
      inserted: counts.inserted,
      updated: counts.updated,
    }),
  });
}

async function importMutasi(rows, onDuplicate, env, user, origin, importMeta = {}) {
  return runStandardJournaledImport(env, user, origin, {
    module: 'mutasi',
    rows,
    onDuplicate,
    importMeta,
    setup: async () => {
      const bRows = await env.DB.prepare('SELECT id, code, name, full_name FROM branches WHERE is_active = 1').all();
      const matchBranch = makeBranchMatcher(bRows.results);
      const empRows = await env.DB.prepare('SELECT id, LOWER(TRIM(full_name)) as name_lower FROM employees WHERE deleted_at IS NULL').all();
      const empCounts = new Map();
      const empIdMap = new Map();
      (empRows.results || []).forEach(e => {
        empCounts.set(e.name_lower, (empCounts.get(e.name_lower) || 0) + 1);
        empIdMap.set(e.name_lower, e.id);
      });
      const existingAll = await env.DB.prepare('SELECT * FROM mutasi_data').all();
      const existingMap = new Map();
      for (const row of (existingAll.results || [])) {
        if (row.tanggal && row.employee_name && row.from_branch_id && row.to_branch_id) {
          const k = `${row.tanggal}_${String(row.employee_name).toLowerCase().trim()}_${row.from_branch_id}_${row.to_branch_id}`;
          if (!existingMap.has(k)) existingMap.set(k, row);
        }
      }
      return { matchBranch, empCounts, empIdMap, existingMap };
    },
    processRow: (item, { matchBranch, empCounts, empIdMap, existingMap }, { onDuplicate: dupMode, seenNewKeys }) => {
      const from_branch_id = item.from_branch_id || matchBranch(item.from_branch_name);
      const to_branch_id = item.to_branch_id || matchBranch(item.to_branch_name);
      if (!item.tanggal || !item.employee_name || !from_branch_id || !to_branch_id) return { action: 'skip' };
      const status = item.status || 'Proses';

      const empNameClean = String(item.employee_name).toLowerCase().trim();
      let empId = null;
      if (item.employee_id) {
        empId = Number(item.employee_id) || null;
      } else if (empCounts.get(empNameClean) === 1) {
        empId = empIdMap.get(empNameClean);
      }

      const mutKey = `${item.tanggal}_${empNameClean}_${from_branch_id}_${to_branch_id}`;
      if (dupMode === 'update' && existingMap.has(mutKey)) {
        const oldRow = existingMap.get(mutKey);
        const afterRow = {
          ...oldRow,
          status,
          document_link: item.document_link || null,
          employee_id: empId ?? oldRow.employee_id,
          row_version: (oldRow.row_version ?? 1) + 1
        };
        existingMap.set(mutKey, afterRow);
        return {
          action: 'update',
          stmts: [
            env.DB.prepare(
              `UPDATE mutasi_data SET status = ?, document_link = ?, employee_id = COALESCE(?, employee_id),
               row_version = COALESCE(row_version, 1) + 1, updated_at = datetime('now') WHERE id = ?`
            ).bind(status, item.document_link || null, empId, oldRow.id),
            buildAuditStatement(env, user, 'UPDATE', 'mutasi', oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true)
          ]
        };
      }

      if (existingMap.has(mutKey) && dupMode === 'skip') return { action: 'skip' };
      if (seenNewKeys.has(mutKey)) return { action: 'skip' };
      seenNewKeys.add(mutKey);

      const newSnapshot = {
        tanggal: item.tanggal,
        employee_id: empId,
        employee_name: item.employee_name,
        from_branch_id,
        to_branch_id,
        status,
        document_link: item.document_link || null,
        row_version: 1
      };
      return {
        action: 'insert',
        stmts: [
          env.DB.prepare(
            `INSERT INTO mutasi_data (tanggal, employee_id, employee_name, from_branch_id, to_branch_id, status, document_link, row_version, updated_at) 
             VALUES (?, ?, ?, ?, ?, ?, ?, 1, datetime('now'))`
          ).bind(
            item.tanggal, empId, item.employee_name, from_branch_id, to_branch_id, status, item.document_link || null
          ),
          buildCreateAuditStatement(env, user, 'mutasi', 'new', newSnapshot, { source: 'import' })
        ]
      };
    },
    formatResult: (counts) => ({
      message: 'Berhasil import Mutasi',
      inserted: counts.inserted,
      updated: counts.updated,
    }),
  });
}

