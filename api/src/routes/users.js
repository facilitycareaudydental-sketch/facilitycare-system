import {
  authenticate,
  hashPassword,
  hasPermission,
  getAssignedModules,
  validatePasswordPolicy,
  validatePbkdf2Iterations,
  buildRevokeAllUserSessionsStatement
} from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated, corsHeaders } from '../utils/response.js';
import { getPagination, getSearchParam, resolveSortOrder, buildTimestampOrderBy } from '../utils/pagination.js';
import { requireExportAccess, getExportPagination } from '../utils/export.js';
import { buildAuditStatement, sanitizeAuditData } from '../utils/audit.js';
import { packStatementGroups } from '../utils/import_journal.js';

const USER_SORT_MAP = Object.freeze({
  created_desc: buildTimestampOrderBy('created_at', { direction: 'DESC', idDir: 'DESC' }),
  name_asc: 'full_name ASC, id ASC',
});

const MAX_BULK_USER_IDS = 100;
export const MAX_ACTIVE_SUPERADMINS = 2;
export const MAX_ACTIVE_FC_SPECIALISTS = 2;
export const MAX_ACTIVE_VIEWERS = 5;

export const OFFICIAL_ROLES = ['superadmin', 'fc_specialist_leader', 'viewer'];
export const CREATABLE_USER_ROLES = ['superadmin', 'fc_specialist_leader', 'viewer'];
export const LEGACY_COMPAT_ROLES = [];
export const VALID_ROLES = [...OFFICIAL_ROLES];

export const ASSIGNABLE_MODULES = [
  'employees',
  'contracts',
  'sp',
  'mutasi',
  'overtime',
  'schedule',
  'os_migrations',
  'issues',
  'training',
  'reports',
  'sop',
  'checklist',
  'hygiene_standards',
  'dac',
  'branches'
];

export const ADMIN_MODULES = [
  'users',
  'audit_logs',
  'audit-logs',
  'sync',
  'sync_dashboard',
  'monitoring',
  'settings',
  'system_settings',
  'import'
];

export function validateAndNormalizeAssignedModules(rawModules, role) {
  if (role !== 'editor_khusus') {
    return { valid: true, modules: null };
  }

  if (rawModules === null || rawModules === undefined || rawModules === '' || (Array.isArray(rawModules) && rawModules.length === 0)) {
    return { valid: true, modules: '[]' };
  }

  let list = [];
  if (Array.isArray(rawModules)) {
    list = rawModules;
  } else if (typeof rawModules === 'string') {
    try {
      const parsed = JSON.parse(rawModules);
      if (Array.isArray(parsed)) list = parsed;
      else list = rawModules.split(',');
    } catch {
      list = rawModules.split(',');
    }
  } else {
    return { valid: false, error: 'Format assigned_modules tidak valid' };
  }

  const cleaned = list.map(m => String(m).trim().toLowerCase().replace(/-/g, '_')).filter(Boolean);
  const deduplicated = Array.from(new Set(cleaned));

  for (const mod of deduplicated) {
    if (ADMIN_MODULES.includes(mod)) {
      return { 
        valid: false, 
        error: `Modul administratif '${mod}' tidak boleh ditugaskan kepada Editor Khusus` 
      };
    }
    if (!ASSIGNABLE_MODULES.includes(mod)) {
      return { 
        valid: false, 
        error: `Modul '${mod}' tidak valid. Modul harus salah satu dari: ${ASSIGNABLE_MODULES.join(', ')}` 
      };
    }
  }

  return { valid: true, modules: JSON.stringify(deduplicated) };
}

export async function handleUsers(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace('/api/users', '');

  if (request.method === 'GET' && path === '/export') {
    const { response } = await requireExportAccess(request, env, origin, 'users');
    if (response) return response;
    return listUsers(request, env, origin, { isExport: true });
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (user.role !== 'superadmin') return forbidden(origin);

  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'GET' && path === '') return listUsers(request, env, origin);
  if (request.method === 'POST' && path === '/import') {
    return importUsers(request, env, user, origin);
  }
  if (request.method === 'DELETE' && path === '/bulk') {
    return bulkDeleteUsers(request, env, user, origin);
  }
  if (request.method === 'POST' && path === '') return createUser(request, env, user, origin);
  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') return getUser(id, env, origin);
    if (request.method === 'PUT') return updateUser(id, request, env, user, origin);
    if (request.method === 'DELETE') return deleteUser(id, env, user, origin);
  }
  return error('Not found', 404, origin);
}

async function getActiveRoleCount(env, targetRole, excludeIds = []) {
  const excludedSet = new Set(excludeIds.map(id => Number(id)));
  const normalizedTarget = String(targetRole || '').trim().toLowerCase();
  try {
    const stmt = env.DB.prepare(
      'SELECT id, role, is_active FROM users WHERE LOWER(TRIM(COALESCE(role, \'\'))) = ? AND COALESCE(is_active, 1) = 1'
    ).bind(normalizedTarget);
    if (typeof stmt.all === 'function') {
      const res = await stmt.all();
      if (res && Array.isArray(res.results)) {
        return res.results.filter(
          u => String(u.role || '').trim().toLowerCase() === normalizedTarget &&
               Number(u.is_active ?? 1) === 1 &&
               !excludedSet.has(Number(u.id))
        ).length;
      }
    }
    if (typeof stmt.first === 'function') {
      const countRes = await env.DB.prepare(
        'SELECT COUNT(*) as total FROM users WHERE LOWER(TRIM(COALESCE(role, \'\'))) = ? AND COALESCE(is_active, 1) = 1'
      ).bind(normalizedTarget).first();
      if (countRes && typeof countRes.total === 'number') {
        return Math.max(0, countRes.total - excludedSet.size);
      }
    }
  } catch (err) {
    console.error(`Error checking active role count (${normalizedTarget}):`, err);
    return -1;
  }
  return -1;
}

async function getRemainingActiveSuperadminCount(env, excludeIds = []) {
  return getActiveRoleCount(env, 'superadmin', excludeIds);
}

async function listUsers(request, env, origin, { isExport = false } = {}) {
  const { page, limit, offset } = isExport ? getExportPagination(request.url) : getPagination(request.url);
  const search = getSearchParam(request.url, 'search').trim();
  const sort = new URL(request.url).searchParams.get('sort') || 'created_desc';
  const orderBy = resolveSortOrder(sort, USER_SORT_MAP, 'created_desc');

  let rows, count;
  try {
    if (search) {
      const query = `SELECT id, username, email, full_name, role, is_active, assigned_modules, created_at FROM users 
                     WHERE full_name LIKE ? OR username LIKE ? OR email LIKE ? 
                     ORDER BY ${orderBy} LIMIT ? OFFSET ?`;
      const countQuery = `SELECT COUNT(*) as total FROM users WHERE full_name LIKE ? OR username LIKE ? OR email LIKE ?`;
      const s = `%${search}%`;
      [rows, count] = await Promise.all([
        env.DB.prepare(query).bind(s, s, s, limit, offset).all(),
        env.DB.prepare(countQuery).bind(s, s, s).first()
      ]);
    } else {
      const query = `SELECT id, username, email, full_name, role, is_active, assigned_modules, created_at FROM users ORDER BY ${orderBy} LIMIT ? OFFSET ?`;
      const countQuery = `SELECT COUNT(*) as total FROM users`;
      [rows, count] = await Promise.all([
        env.DB.prepare(query).bind(limit, offset).all(),
        env.DB.prepare(countQuery).first()
      ]);
    }
  } catch (err) {
    const msg = String(err?.message || err || '');
    if (!msg.includes('no such column: assigned_modules')) {
      console.error('listUsers DB error:', err);
      return error('Gagal mengambil daftar user', 500, origin);
    }
    // Fallback ONLY if assigned_modules column is not present pre-migration
    try {
      if (search) {
        const query = `SELECT id, username, email, full_name, role, is_active, created_at FROM users 
                       WHERE full_name LIKE ? OR username LIKE ? OR email LIKE ? 
                       ORDER BY ${orderBy} LIMIT ? OFFSET ?`;
        const countQuery = `SELECT COUNT(*) as total FROM users WHERE full_name LIKE ? OR username LIKE ? OR email LIKE ?`;
        const s = `%${search}%`;
        [rows, count] = await Promise.all([
          env.DB.prepare(query).bind(s, s, s, limit, offset).all(),
          env.DB.prepare(countQuery).bind(s, s, s).first()
        ]);
      } else {
        const query = `SELECT id, username, email, full_name, role, is_active, created_at FROM users ORDER BY ${orderBy} LIMIT ? OFFSET ?`;
        const countQuery = `SELECT COUNT(*) as total FROM users`;
        [rows, count] = await Promise.all([
          env.DB.prepare(query).bind(limit, offset).all(),
          env.DB.prepare(countQuery).first()
        ]);
      }
    } catch (fallbackErr) {
      console.error('listUsers fallback DB error:', fallbackErr);
      return error('Gagal mengambil daftar user', 500, origin);
    }
  }

  const results = (rows?.results || []).map(u => ({
    ...u,
    assigned_modules: getAssignedModules(u)
  }));
  return paginated(results, count.total, page, limit, origin);
}

async function getUser(id, env, origin) {
  let row;
  try {
    row = await env.DB.prepare(
      'SELECT id, username, email, full_name, role, is_active, assigned_modules, created_at, updated_at FROM users WHERE id = ?'
    ).bind(id).first();
  } catch (err) {
    const msg = String(err?.message || err || '');
    if (!msg.includes('no such column: assigned_modules')) {
      console.error('getUser DB error:', err);
      return error('Gagal mengambil data user', 500, origin);
    }
    try {
      row = await env.DB.prepare(
        'SELECT id, username, email, full_name, role, is_active, created_at, updated_at FROM users WHERE id = ?'
      ).bind(id).first();
    } catch (fallbackErr) {
      console.error('getUser fallback DB error:', fallbackErr);
      return error('Gagal mengambil data user', 500, origin);
    }
  }
  if (!row) return notFound(origin);
  row.assigned_modules = getAssignedModules(row);
  return ok(row, 200, origin);
}

async function createUser(request, env, currentUser, origin) {
  if (currentUser.role !== 'superadmin') return forbidden(origin);
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!body || typeof body !== 'object' || Array.isArray(body)) {
    return error('Invalid JSON payload', 400, origin);
  }

  const { username, email, full_name, password, role, assigned_modules, is_active } = body;
  const cleanUsername = username ? String(username).trim() : null;
  const cleanEmail = email ? String(email).trim() : null;
  const cleanFullName = full_name ? String(full_name).trim() : null;
  const normUsername = cleanUsername ? cleanUsername.toLowerCase() : null;
  const normEmail = cleanEmail ? cleanEmail.toLowerCase() : null;

  let targetIsActive = 1;
  if (is_active !== undefined && is_active !== null) {
    if (is_active === 1 || is_active === true || is_active === '1') {
      targetIsActive = 1;
    } else if (is_active === 0 || is_active === false || is_active === '0') {
      targetIsActive = 0;
    } else {
      return error('Invalid is_active: harus 0, 1, atau boolean', 400, origin);
    }
  }

  if (!cleanUsername || !cleanEmail || !cleanFullName || password === undefined || password === null || password === '') {
    return error('username, email, full_name, password required', 400, origin);
  }
  if (typeof password !== 'string') {
    return error('Password harus berupa string', 400, origin);
  }

  const pwValidation = validatePasswordPolicy(password);
  if (!pwValidation.valid) {
    return error(pwValidation.error, 400, origin);
  }

  let iterations;
  try {
    iterations = validatePbkdf2Iterations(env?.PBKDF2_ITERATIONS);
  } catch (iterErr) {
    console.error('AUTH_CONFIG_ERROR:', iterErr?.message || iterErr);
    return error('Authentication configuration error: invalid PBKDF2_ITERATIONS', 500, origin);
  }

  const targetRole = role !== undefined && role !== null && String(role).trim() !== ''
    ? String(role).trim().toLowerCase()
    : '';
  if (!CREATABLE_USER_ROLES.includes(targetRole)) {
    return error(
      `Invalid role '${role || ''}': tidak diizinkan untuk pembuatan akun baru. Akun baru hanya dapat menggunakan role resmi: ${CREATABLE_USER_ROLES.join(', ')}`,
      400,
      origin
    );
  }
  if (targetIsActive === 1 && targetRole === 'superadmin') {
    const activeSuperadmins = await getActiveRoleCount(env, 'superadmin', []);
    if (activeSuperadmins < 0) {
      return error('Gagal memverifikasi jumlah Superadmin aktif', 503, origin);
    }
    if (activeSuperadmins >= MAX_ACTIVE_SUPERADMINS) {
      return error(
        `Konflik Kuota Role: Sistem hanya mengizinkan maksimal ${MAX_ACTIVE_SUPERADMINS} Superadmin aktif. Pembuatan akun Superadmin ketiga ditolak.`,
        409,
        origin
      );
    }
  }
  if (targetIsActive === 1 && targetRole === 'fc_specialist_leader') {
    const activeSpecialists = await getActiveRoleCount(env, 'fc_specialist_leader', []);
    if (activeSpecialists < 0) {
      return error('Gagal memverifikasi jumlah Facility Care Specialist aktif', 503, origin);
    }
    if (activeSpecialists >= MAX_ACTIVE_FC_SPECIALISTS) {
      return error(
        `Konflik Kuota Role: Sistem hanya mengizinkan maksimal ${MAX_ACTIVE_FC_SPECIALISTS} Facility Care Specialist aktif. Pembuatan akun ketiga ditolak.`,
        409,
        origin
      );
    }
  }
  if (targetIsActive === 1 && targetRole === 'viewer') {
    const activeViewers = await getActiveRoleCount(env, 'viewer', []);
    if (activeViewers < 0) {
      return error('Gagal memverifikasi jumlah Viewer aktif', 503, origin);
    }
    if (activeViewers >= MAX_ACTIVE_VIEWERS) {
      return error(
        `Konflik Kuota Role: Sistem hanya mengizinkan maksimal ${MAX_ACTIVE_VIEWERS} Viewer aktif. Pembuatan akun keenam ditolak.`,
        409,
        origin
      );
    }
  }

  const moduleValidation = validateAndNormalizeAssignedModules(assigned_modules, targetRole);
  if (!moduleValidation.valid) {
    return error(moduleValidation.error, 400, origin);
  }

  const existing = await env.DB.prepare(
    'SELECT id, username, email FROM users WHERE LOWER(TRIM(username)) = ? OR LOWER(TRIM(email)) = ?'
  ).bind(normUsername, normEmail).first();
  if (existing) {
    if (existing.email && existing.email.trim().toLowerCase() === normEmail) {
      return error(`Email '${cleanEmail}' sudah digunakan oleh user lain.`, 409, origin);
    }
    return error(`Username '${cleanUsername}' sudah digunakan oleh user lain.`, 409, origin);
  }

  const hash = await hashPassword(password, iterations);
  const sanitizedNew = {
    username: cleanUsername,
    email: cleanEmail,
    full_name: cleanFullName,
    role: targetRole,
    is_active: targetIsActive,
    assigned_modules: moduleValidation.modules
  };
  const newStr = JSON.stringify(sanitizedNew);
  const detailsStr = JSON.stringify({ status: 'SUCCESS' });
  const createAuditStmt = env.DB.prepare(
    `INSERT INTO audit_logs (user_id, user_name, user_role, action, module, target_id, old_data, new_data, old_value, new_value, details)
     SELECT ?, ?, ?, 'CREATE', 'users', COALESCE(CAST(last_insert_rowid() AS TEXT), ?), NULL, ?, NULL, ?, ?
     WHERE (SELECT changes()) > 0`
  ).bind(
    currentUser?.id || 0,
    currentUser?.full_name || currentUser?.username || 'System Automation',
    currentUser?.role || 'system',
    cleanUsername,
    newStr,
    newStr,
    detailsStr
  );

  try {
    const insertStmt = targetIsActive === 0
      ? env.DB.prepare(
          `INSERT INTO users (username, email, password_hash, full_name, role, is_active, assigned_modules)
           VALUES (?, ?, ?, ?, ?, 0, ?)`
        ).bind(
          cleanUsername,
          cleanEmail,
          hash,
          cleanFullName,
          targetRole,
          moduleValidation.modules
        )
      : env.DB.prepare(
          `INSERT INTO users (username, email, password_hash, full_name, role, is_active, assigned_modules)
           SELECT ?, ?, ?, ?, ?, 1, ?
           WHERE (
             (LOWER(?) = 'superadmin' AND (SELECT COUNT(*) FROM users WHERE LOWER(TRIM(COALESCE(role, ''))) = 'superadmin' AND COALESCE(is_active, 1) = 1) < ?)
             OR
             (LOWER(?) = 'fc_specialist_leader' AND (SELECT COUNT(*) FROM users WHERE LOWER(TRIM(COALESCE(role, ''))) = 'fc_specialist_leader' AND COALESCE(is_active, 1) = 1) < ?)
             OR
             (LOWER(?) = 'viewer' AND (SELECT COUNT(*) FROM users WHERE LOWER(TRIM(COALESCE(role, ''))) = 'viewer' AND COALESCE(is_active, 1) = 1) < ?)
           )`
        ).bind(
          cleanUsername,
          cleanEmail,
          hash,
          cleanFullName,
          targetRole,
          moduleValidation.modules,
          targetRole,
          MAX_ACTIVE_SUPERADMINS,
          targetRole,
          MAX_ACTIVE_FC_SPECIALISTS,
          targetRole,
          MAX_ACTIVE_VIEWERS
        );
    const batchRes = await env.DB.batch([insertStmt, createAuditStmt]);
    if (!batchRes?.[0]?.meta || Number(batchRes[0].meta.changes) < 1) {
      return error('Konflik Kuota Role: Pembuatan akun ditolak oleh pengaman kuota maksimal database.', 409, origin);
    }

    return ok(
      { id: batchRes?.[0]?.meta?.last_row_id, is_active: targetIsActive, message: 'User created' },
      201,
      origin,
      undefined,
      { 'X-Audit-Batched': 'true' }
    );
  } catch (err) {
    const msg = err?.message || String(err);
    if (msg.includes('no such column: assigned_modules')) {
      try {
        const fallbackStmt = targetIsActive === 0
          ? env.DB.prepare(
              `INSERT INTO users (username, email, password_hash, full_name, role, is_active)
               VALUES (?, ?, ?, ?, ?, 0)`
            ).bind(cleanUsername, cleanEmail, hash, cleanFullName, targetRole)
          : env.DB.prepare(
              `INSERT INTO users (username, email, password_hash, full_name, role, is_active)
               SELECT ?, ?, ?, ?, ?, 1
               WHERE (
                 (LOWER(?) = 'superadmin' AND (SELECT COUNT(*) FROM users WHERE LOWER(TRIM(COALESCE(role, ''))) = 'superadmin' AND COALESCE(is_active, 1) = 1) < ?)
                 OR
                 (LOWER(?) = 'fc_specialist_leader' AND (SELECT COUNT(*) FROM users WHERE LOWER(TRIM(COALESCE(role, ''))) = 'fc_specialist_leader' AND COALESCE(is_active, 1) = 1) < ?)
               )`
            ).bind(
              cleanUsername,
              cleanEmail,
              hash,
              cleanFullName,
              targetRole,
              targetRole,
              MAX_ACTIVE_SUPERADMINS,
              targetRole,
              MAX_ACTIVE_FC_SPECIALISTS
            );
        const fallbackBatch = await env.DB.batch([fallbackStmt, createAuditStmt]);
        if (!fallbackBatch?.[0]?.meta || Number(fallbackBatch[0].meta.changes) < 1) {
          return error('Konflik Kuota Role: Pembuatan akun ditolak oleh pengaman kuota maksimal database.', 409, origin);
        }
        return ok(
          { id: fallbackBatch?.[0]?.meta?.last_row_id, is_active: targetIsActive, message: 'User created' },
          201,
          origin,
          undefined,
          { 'X-Audit-Batched': 'true' }
        );
      } catch (fallbackErr) {
        const fbMsg = fallbackErr?.message || String(fallbackErr);
        if (fbMsg.includes('UNIQUE constraint failed: users.email')) {
          return error('Email sudah digunakan oleh user lain.', 409, origin);
        }
        if (fbMsg.includes('UNIQUE constraint failed: users.username')) {
          return error('Username sudah digunakan oleh user lain.', 409, origin);
        }
        if (
          fbMsg.includes('MAX_ACTIVE_SUPERADMINS_EXCEEDED') ||
          fbMsg.includes('MAX_ACTIVE_FC_SPECIALIST_EXCEEDED') ||
          fbMsg.includes('idx_users_single_active_superadmin') ||
          fbMsg.includes('UNIQUE constraint failed')
        ) {
          return error('Konflik Kuota Role: Batas maksimal 2 akun aktif untuk role ini telah tercapai.', 409, origin);
        }
        console.error('createUser fallback DB error:', fallbackErr);
        return error('Gagal membuat user', 500, origin);
      }
    }
    if (msg.includes('UNIQUE constraint failed: users.email')) {
      return error('Email sudah digunakan oleh user lain.', 409, origin);
    }
    if (msg.includes('UNIQUE constraint failed: users.username')) {
      return error('Username sudah digunakan oleh user lain.', 409, origin);
    }
    if (
      msg.includes('MAX_ACTIVE_SUPERADMINS_EXCEEDED') ||
      msg.includes('MAX_ACTIVE_FC_SPECIALIST_EXCEEDED') ||
      msg.includes('idx_users_single_active_superadmin') ||
      msg.includes('UNIQUE constraint failed')
    ) {
      return error('Konflik Kuota Role: Batas maksimal 2 akun aktif untuk role ini telah tercapai.', 409, origin);
    }
    console.error('createUser DB error:', err);
    return error('Gagal membuat user', 500, origin);
  }
}

async function updateUser(id, request, env, currentUser, origin) {
  if (currentUser.role !== 'superadmin') return forbidden(origin);
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }

  let existing;
  try {
    existing = await env.DB.prepare(
      'SELECT id, username, email, full_name, role, is_active, assigned_modules FROM users WHERE id = ?'
    ).bind(id).first();
  } catch (err) {
    const msg = String(err?.message || err || '');
    if (!msg.includes('no such column: assigned_modules')) {
      console.error('updateUser fetch DB error:', err);
      return error('Gagal mengambil data user', 500, origin);
    }
    try {
      existing = await env.DB.prepare(
        'SELECT id, username, email, full_name, role, is_active FROM users WHERE id = ?'
      ).bind(id).first();
    } catch (fallbackErr) {
      console.error('updateUser fallback fetch DB error:', fallbackErr);
      return error('Gagal mengambil data user', 500, origin);
    }
  }
  if (!existing) return notFound(origin);

  const { username, email, full_name, role, is_active, password, assigned_modules } = body;
  let updates = [];
  let values = [];
  const isExistingSuperadminRole = String(existing.role || '').trim().toLowerCase() === 'superadmin';
  const isExistingActiveSuperadmin = isExistingSuperadminRole && Number(existing.is_active ?? 1) === 1;
  const isExistingFcSpecialistRole = String(existing.role || '').trim().toLowerCase() === 'fc_specialist_leader';
  const isExistingActiveFcSpecialist = isExistingFcSpecialistRole && Number(existing.is_active ?? 1) === 1;
  const isExistingViewerRole = String(existing.role || '').trim().toLowerCase() === 'viewer';
  const isExistingActiveViewer = isExistingViewerRole && Number(existing.is_active ?? 1) === 1;
  let requiresRemainingSuperadminGuard = false;
  let requiresSuperadminCeilingGuard = false;
  let requiresFcSpecialistCeilingGuard = false;
  let roleChanged = false;
  let assignedModulesChanged = false;
  let deactivated = false;
  let passwordChanged = false;
  const fieldsUpdated = [];
  const updatedSnapshot = {
    username: existing.username,
    email: existing.email,
    full_name: existing.full_name,
    role: existing.role,
    is_active: existing.is_active,
    assigned_modules: existing.assigned_modules ?? null
  };

  const cleanEmail = (email !== undefined && email !== null) ? String(email).trim() : null;
  const cleanUsername = (username !== undefined && username !== null) ? String(username).trim() : null;
  const cleanFullName = (full_name !== undefined && full_name !== null) ? String(full_name).trim() : null;
  const normEmail = cleanEmail ? cleanEmail.toLowerCase() : null;
  const normUsername = cleanUsername ? cleanUsername.toLowerCase() : null;

  // 1. Email: Only check uniqueness and update if email actually changed!
  if (cleanEmail) {
    const existingNormEmail = (existing.email || '').trim().toLowerCase();
    if (normEmail !== existingNormEmail) {
      const emailConflict = await env.DB.prepare(
        'SELECT id, username, full_name FROM users WHERE LOWER(TRIM(email)) = ? AND id != ?'
      ).bind(normEmail, id).first();
      if (emailConflict) {
        return error(`Email '${cleanEmail}' sudah digunakan oleh user lain (${emailConflict.username} - ${emailConflict.full_name}). Gunakan email yang berbeda.`, 409, origin);
      }
      updates.push('email = ?');
      values.push(cleanEmail);
      updatedSnapshot.email = cleanEmail;
      fieldsUpdated.push('email');
    } else if (cleanEmail !== existing.email) {
      updates.push('email = ?');
      values.push(cleanEmail);
      updatedSnapshot.email = cleanEmail;
      fieldsUpdated.push('email');
    }
  }

  // 2. Username: Only check uniqueness and update if username actually changed!
  if (cleanUsername) {
    const existingNormUsername = (existing.username || '').trim().toLowerCase();
    if (normUsername !== existingNormUsername) {
      const usernameConflict = await env.DB.prepare(
        'SELECT id, username, full_name FROM users WHERE LOWER(TRIM(username)) = ? AND id != ?'
      ).bind(normUsername, id).first();
      if (usernameConflict) {
        return error(`Username '${cleanUsername}' sudah digunakan oleh user lain (${usernameConflict.full_name}). Gunakan username yang berbeda.`, 409, origin);
      }
      updates.push('username = ?');
      values.push(cleanUsername);
      updatedSnapshot.username = cleanUsername;
      fieldsUpdated.push('username');
    } else if (cleanUsername !== existing.username) {
      updates.push('username = ?');
      values.push(cleanUsername);
      updatedSnapshot.username = cleanUsername;
      fieldsUpdated.push('username');
    }
  }

  // 3. Full Name: Only update if changed
  if (cleanFullName && cleanFullName !== existing.full_name) {
    updates.push('full_name = ?');
    values.push(cleanFullName);
    updatedSnapshot.full_name = cleanFullName;
    fieldsUpdated.push('full_name');
  }

  // 4. Role: Only update if changed, valid official role, and user is superadmin
  let targetRole = existing.role;
  const intendedFinalActive = (is_active !== undefined && is_active !== null)
    ? ((is_active === 1 || is_active === true || is_active === '1') ? 1 : 0)
    : Number(existing.is_active ?? 1);
  if (role !== undefined && role !== null) {
    const cleanRole = String(role).trim().toLowerCase();
    if (!OFFICIAL_ROLES.includes(cleanRole)) {
      return error(`Invalid role '${role}'. Role resmi FCMS hanya: ${OFFICIAL_ROLES.join(', ')}`, 400, origin);
    }
    if (intendedFinalActive === 1 && cleanRole === 'superadmin' && !isExistingSuperadminRole) {
      const activeSuperadmins = await getActiveRoleCount(env, 'superadmin', [id]);
      if (activeSuperadmins < 0) {
        return error('Gagal memverifikasi jumlah Superadmin aktif', 503, origin);
      }
      if (activeSuperadmins >= MAX_ACTIVE_SUPERADMINS) {
        return error(
          `Konflik Kuota Role: Tidak dapat mempromosikan user menjadi Superadmin karena batas maksimal ${MAX_ACTIVE_SUPERADMINS} Superadmin aktif telah tercapai.`,
          409,
          origin
        );
      }
    }
    if (intendedFinalActive === 1 && cleanRole === 'fc_specialist_leader' && !isExistingFcSpecialistRole) {
      const activeSpecialists = await getActiveRoleCount(env, 'fc_specialist_leader', [id]);
      if (activeSpecialists < 0) {
        return error('Gagal memverifikasi jumlah Facility Care Specialist aktif', 503, origin);
      }
      if (activeSpecialists >= MAX_ACTIVE_FC_SPECIALISTS) {
        return error(
          `Konflik Kuota Role: Tidak dapat mempromosikan user menjadi Facility Care Specialist karena batas maksimal ${MAX_ACTIVE_FC_SPECIALISTS} akun aktif telah tercapai.`,
          409,
          origin
        );
      }
    }
    if (currentUser.id == id && cleanRole !== 'superadmin') {
      return error('Tidak dapat menurunkan role akun sendiri dari superadmin', 400, origin);
    }
    if (isExistingActiveSuperadmin && cleanRole !== 'superadmin') {
      const remainingSuperadmins = await getRemainingActiveSuperadminCount(env, [id]);
      if (remainingSuperadmins < 0) {
        return error('Gagal memverifikasi jumlah Superadmin aktif', 503, origin);
      }
      if (remainingSuperadmins < 1) {
        return error('Tidak dapat menurunkan role Superadmin terakhir yang aktif', 400, origin);
      }
      requiresRemainingSuperadminGuard = true;
    }
    targetRole = cleanRole;
    if (cleanRole !== existing.role) {
      updates.push('role = ?');
      values.push(cleanRole);
      roleChanged = true;
      updatedSnapshot.role = cleanRole;
      fieldsUpdated.push('role');
    }
  }

  // 4.b Assigned Modules: Only Superadmin can assign or change assigned_modules
  if (assigned_modules !== undefined) {
    const moduleValidation = validateAndNormalizeAssignedModules(assigned_modules, targetRole);
    if (!moduleValidation.valid) {
      return error(moduleValidation.error, 400, origin);
    }
    if (moduleValidation.modules !== existing.assigned_modules) {
      updates.push('assigned_modules = ?');
      values.push(moduleValidation.modules);
      assignedModulesChanged = true;
      updatedSnapshot.assigned_modules = moduleValidation.modules;
      fieldsUpdated.push('assigned_modules');
    }
  } else if (role !== undefined && targetRole !== 'editor_khusus' && existing.assigned_modules !== null && existing.assigned_modules !== undefined) {
    // When role is changed away from editor_khusus without explicit assigned_modules, clear assignment
    updates.push('assigned_modules = ?');
    values.push(null);
    assignedModulesChanged = true;
    updatedSnapshot.assigned_modules = null;
    fieldsUpdated.push('assigned_modules');
  }

  // 5. Active Status: Only update if changed and user is superadmin
  if (is_active !== undefined && is_active !== null) {
    const targetActive = (is_active === 1 || is_active === true || is_active === '1') ? 1 : 0;
    if (currentUser.id == id && targetActive === 0) {
      return error('Tidak dapat menonaktifkan akun sendiri', 400, origin);
    }
    if (isExistingActiveSuperadmin && targetActive === 0) {
      const remainingSuperadmins = await getRemainingActiveSuperadminCount(env, [id]);
      if (remainingSuperadmins < 0) {
        return error('Gagal memverifikasi jumlah Superadmin aktif', 503, origin);
      }
      if (remainingSuperadmins < 1) {
        return error('Tidak dapat menonaktifkan Superadmin terakhir yang aktif', 400, origin);
      }
      requiresRemainingSuperadminGuard = true;
    }
    if (targetActive !== existing.is_active) {
      updates.push('is_active = ?');
      values.push(targetActive);
      if (targetActive === 0) {
        deactivated = true;
      }
      updatedSnapshot.is_active = targetActive;
      fieldsUpdated.push('is_active');
    }
  }

  // 5.b Role Ceiling Guards (Max 2 Active Superadmins, Max 2 Active FC Specialists):
  // Prevent reactivating an inactive account or promoting any account into superadmin / fc_specialist_leader
  // when 2 active accounts with that role already exist.
  const finalRole = String(targetRole || existing.role || '').trim().toLowerCase();
  const finalActive = (is_active !== undefined && is_active !== null)
    ? ((is_active === 1 || is_active === true || is_active === '1') ? 1 : 0)
    : Number(existing.is_active ?? 1);
  if (finalRole === 'superadmin' && finalActive === 1 && !isExistingActiveSuperadmin) {
    const otherActiveSuperadmins = await getActiveRoleCount(env, 'superadmin', [id]);
    if (otherActiveSuperadmins < 0) {
      return error('Gagal memverifikasi jumlah Superadmin aktif', 503, origin);
    }
    if (otherActiveSuperadmins >= MAX_ACTIVE_SUPERADMINS) {
      return error(
        `Konflik Kuota Role: Tidak dapat mengaktifkan Superadmin baru karena maksimal ${MAX_ACTIVE_SUPERADMINS} Superadmin aktif telah tercapai.`,
        409,
        origin
      );
    }
    requiresSuperadminCeilingGuard = true;
  }
  if (finalRole === 'fc_specialist_leader' && finalActive === 1 && !isExistingActiveFcSpecialist) {
    const otherActiveSpecialists = await getActiveRoleCount(env, 'fc_specialist_leader', [id]);
    if (otherActiveSpecialists < 0) {
      return error('Gagal memverifikasi jumlah Facility Care Specialist aktif', 503, origin);
    }
    if (otherActiveSpecialists >= MAX_ACTIVE_FC_SPECIALISTS) {
      return error(
        `Konflik Kuota Role: Tidak dapat mengaktifkan Facility Care Specialist baru karena maksimal ${MAX_ACTIVE_FC_SPECIALISTS} akun aktif telah tercapai.`,
        409,
        origin
      );
    }
    requiresFcSpecialistCeilingGuard = true;
  }
  if (finalRole === 'viewer' && finalActive === 1 && !isExistingActiveViewer) {
    const otherActiveViewers = await getActiveRoleCount(env, 'viewer', [id]);
    if (otherActiveViewers < 0) {
      return error('Gagal memverifikasi jumlah Viewer aktif', 503, origin);
    }
    if (otherActiveViewers >= MAX_ACTIVE_VIEWERS) {
      return error(
        `Konflik Kuota Role: Tidak dapat mengaktifkan Viewer baru karena maksimal ${MAX_ACTIVE_VIEWERS} akun aktif telah tercapai.`,
        409,
        origin
      );
    }
  }

  // 6. Password update (admin password reset): Enforce centralized password policy before hashing
  if (password !== undefined && password !== null && password !== '') {
    if (typeof password !== 'string') {
      return error('Password harus berupa string', 400, origin);
    }
    const pwValidation = validatePasswordPolicy(password);
    if (!pwValidation.valid) {
      return error(pwValidation.error, 400, origin);
    }
    let iterations;
    try {
      iterations = validatePbkdf2Iterations(env?.PBKDF2_ITERATIONS);
    } catch (iterErr) {
      console.error('AUTH_CONFIG_ERROR:', iterErr?.message || iterErr);
      return error('Authentication configuration error: invalid PBKDF2_ITERATIONS', 500, origin);
    }
    const hash = await hashPassword(password, iterations);
    updates.push('password_hash = ?');
    values.push(hash);
    passwordChanged = true;
    updatedSnapshot.password_changed = true;
    fieldsUpdated.push('password');
  }

  // If nothing changed, return success directly without unnecessary DB write
  if (updates.length === 0) {
    return ok({ message: 'Tidak ada perubahan data' }, 200, origin);
  }

  const shouldRevokeSessions = passwordChanged || roleChanged || assignedModulesChanged || deactivated;

  updates.push("updated_at = datetime('now')");
  values.push(id);

  let whereConditions = ['id = ?'];
  let guardBindValues = [];
  if (requiresRemainingSuperadminGuard) {
    whereConditions.push("(SELECT COUNT(*) FROM users WHERE LOWER(TRIM(COALESCE(role, ''))) = 'superadmin' AND COALESCE(is_active, 1) = 1 AND id != ?) >= 1");
    guardBindValues.push(id);
  }
  if (requiresSuperadminCeilingGuard) {
    whereConditions.push("(SELECT COUNT(*) FROM users WHERE LOWER(TRIM(COALESCE(role, ''))) = 'superadmin' AND COALESCE(is_active, 1) = 1 AND id != ?) < ?");
    guardBindValues.push(id, MAX_ACTIVE_SUPERADMINS);
  }
  if (requiresFcSpecialistCeilingGuard) {
    whereConditions.push("(SELECT COUNT(*) FROM users WHERE LOWER(TRIM(COALESCE(role, ''))) = 'fc_specialist_leader' AND COALESCE(is_active, 1) = 1 AND id != ?) < ?");
    guardBindValues.push(id, MAX_ACTIVE_FC_SPECIALISTS);
  }
  const whereClause = `WHERE ${whereConditions.join(' AND ')}`;
  const bindValues = [...values, ...guardBindValues];

  const auditStmt = buildAuditStatement(
    env,
    currentUser,
    'UPDATE',
    'users',
    id,
    existing,
    updatedSnapshot,
    { fields_updated: fieldsUpdated, sessions_revoked: shouldRevokeSessions },
    'SUCCESS',
    true
  );

  const buildUpdateBatch = (stmt) => {
    const stmts = [stmt, auditStmt];
    if (shouldRevokeSessions) {
      stmts.push(buildRevokeAllUserSessionsStatement(env, id));
    }
    return stmts;
  };

  try {
    const updateStmt = env.DB.prepare(`UPDATE users SET ${updates.join(', ')} ${whereClause}`).bind(...bindValues);
    const batchRes = await env.DB.batch(buildUpdateBatch(updateStmt));
    if (requiresRemainingSuperadminGuard && (!batchRes?.[0]?.meta || Number(batchRes[0].meta.changes) < 1)) {
      return error('Tidak dapat menurunkan role atau menonaktifkan Superadmin terakhir yang aktif', 400, origin);
    }
    if ((requiresSuperadminCeilingGuard || requiresFcSpecialistCeilingGuard) && (!batchRes?.[0]?.meta || Number(batchRes[0].meta.changes) < 1)) {
      return error('Konflik Kuota Role: Batas maksimal 2 akun aktif untuk role ini telah tercapai.', 409, origin);
    }
    return ok(
      { message: 'User updated', sessions_revoked: shouldRevokeSessions },
      200,
      origin,
      undefined,
      { 'X-Audit-Batched': 'true' }
    );
  } catch (err) {
    const msg = err?.message || String(err);
    if (msg.includes('no such column: assigned_modules')) {
      try {
        const filteredUpdates = [];
        const filteredValues = [];
        for (let i = 0; i < updates.length; i++) {
          if (!updates[i].startsWith('assigned_modules')) {
            filteredUpdates.push(updates[i]);
            filteredValues.push(values[i]);
          }
        }
        if (filteredUpdates.length > 1) { // at least something other than updated_at
          const fallbackBindValues = [...filteredValues, ...guardBindValues];
          const fallbackStmt = env.DB.prepare(`UPDATE users SET ${filteredUpdates.join(', ')} ${whereClause}`).bind(...fallbackBindValues);
          const fallbackRes = await env.DB.batch(buildUpdateBatch(fallbackStmt));
          if (requiresRemainingSuperadminGuard && (!fallbackRes?.[0]?.meta || Number(fallbackRes[0].meta.changes) < 1)) {
            return error('Tidak dapat menurunkan role atau menonaktifkan Superadmin terakhir yang aktif', 400, origin);
          }
          if ((requiresSuperadminCeilingGuard || requiresFcSpecialistCeilingGuard) && (!fallbackRes?.[0]?.meta || Number(fallbackRes[0].meta.changes) < 1)) {
            return error('Konflik Kuota Role: Batas maksimal 2 akun aktif untuk role ini telah tercapai.', 409, origin);
          }
        }
        return ok(
          { message: 'User updated (column assigned_modules pending migration)', sessions_revoked: shouldRevokeSessions },
          200,
          origin,
          undefined,
          { 'X-Audit-Batched': 'true' }
        );
      } catch (fallbackErr) {
        const fbMsg = fallbackErr?.message || String(fallbackErr);
        if (fbMsg.includes('UNIQUE constraint failed: users.email')) {
          return error('Email sudah digunakan oleh user lain. Gunakan email yang berbeda.', 409, origin);
        }
        if (fbMsg.includes('UNIQUE constraint failed: users.username')) {
          return error('Username sudah digunakan oleh user lain. Gunakan username yang berbeda.', 409, origin);
        }
        if (
          fbMsg.includes('MAX_ACTIVE_SUPERADMINS_EXCEEDED') ||
          fbMsg.includes('MAX_ACTIVE_FC_SPECIALIST_EXCEEDED') ||
          fbMsg.includes('idx_users_single_active_superadmin') ||
          fbMsg.includes('UNIQUE constraint failed')
        ) {
          return error('Konflik Kuota Role: Batas maksimal 2 akun aktif untuk role ini telah tercapai.', 409, origin);
        }
        console.error('updateUser fallback DB error:', fallbackErr);
        return error('Gagal memperbarui user', 500, origin);
      }
    }
    if (msg.includes('UNIQUE constraint failed: users.email')) {
      return error('Email sudah digunakan oleh user lain. Gunakan email yang berbeda.', 409, origin);
    }
    if (msg.includes('UNIQUE constraint failed: users.username')) {
      return error('Username sudah digunakan oleh user lain. Gunakan username yang berbeda.', 409, origin);
    }
    if (
      msg.includes('MAX_ACTIVE_SUPERADMINS_EXCEEDED') ||
      msg.includes('MAX_ACTIVE_FC_SPECIALIST_EXCEEDED') ||
      msg.includes('idx_users_single_active_superadmin') ||
      msg.includes('UNIQUE constraint failed')
    ) {
      return error('Konflik Kuota Role: Batas maksimal 2 akun aktif untuk role ini telah tercapai.', 409, origin);
    }
    console.error('updateUser DB error:', err);
    return error('Gagal memperbarui user', 500, origin);
  }
}

async function deleteUser(id, env, currentUser, origin) {
  if (currentUser.role !== 'superadmin') return forbidden(origin);
  if (currentUser.id == id) return error('Cannot delete yourself', 400, origin);
  try {
    const existing = await env.DB.prepare('SELECT id, username, role, is_active FROM users WHERE id = ?').bind(id).first();
    if (!existing) return notFound(origin);
    const isActiveSuperadmin = String(existing.role || '').toLowerCase() === 'superadmin' && Number(existing.is_active ?? 1) === 1;
    const auditStmt = buildAuditStatement(env, currentUser, 'DELETE', 'users', id, existing, null, null, 'SUCCESS', true);
    const revokeStmt = buildRevokeAllUserSessionsStatement(env, id);

    if (isActiveSuperadmin) {
      const remainingSuperadmins = await getRemainingActiveSuperadminCount(env, [id]);
      if (remainingSuperadmins < 0) {
        return error('Gagal memverifikasi jumlah Superadmin aktif', 503, origin);
      }
      if (remainingSuperadmins < 1) {
        return error('Tidak dapat menghapus Superadmin terakhir yang aktif', 400, origin);
      }
      const delStmt = env.DB.prepare(
        "DELETE FROM users WHERE id = ? AND (SELECT COUNT(*) FROM users WHERE LOWER(role) = 'superadmin' AND is_active = 1 AND id != ?) >= 1"
      ).bind(id, id);
      const batchRes = await env.DB.batch([delStmt, auditStmt, revokeStmt]);
      if (!batchRes[0]?.meta || Number(batchRes[0].meta.changes) < 1) {
        return error('Tidak dapat menghapus Superadmin terakhir yang aktif', 400, origin);
      }
      return ok({ message: 'User deleted' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
    }
    const delStmt = env.DB.prepare('DELETE FROM users WHERE id = ?').bind(id);
    await env.DB.batch([delStmt, auditStmt, revokeStmt]);
    return ok({ message: 'User deleted' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
  } catch (err) {
    console.error('deleteUser DB error:', err);
    return error('Gagal menghapus user', 500, origin);
  }
}

async function bulkDeleteUsers(request, env, currentUser, origin) {
  if (currentUser.role !== 'superadmin') return forbidden(origin);
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const { ids } = body || {};
  if (!Array.isArray(ids) || ids.length === 0) return error('ids array required', 400, origin);

  const uniquePositiveIds = Array.from(new Set(
    ids
      .map(id => Number(id))
      .filter(id => Number.isInteger(id) && id > 0)
  ));

  if (uniquePositiveIds.length === 0) {
    return error('Tidak ada ID user valid yang dipilih', 400, origin);
  }
  if (uniquePositiveIds.length > MAX_BULK_USER_IDS) {
    return error(`Maksimum ${MAX_BULK_USER_IDS} item diperbolehkan per permintaan bulk delete.`, 400, origin);
  }

  // Exclude currently logged-in superadmin
  const validIds = uniquePositiveIds.filter(id => id !== Number(currentUser.id));
  if (validIds.length === 0) {
    return error('Tidak ada user valid untuk dihapus (akun sendiri tidak dapat dihapus)', 400, origin);
  }

  const remainingSuperadmins = await getRemainingActiveSuperadminCount(env, validIds);
  if (remainingSuperadmins < 0) {
    return error('Gagal memverifikasi jumlah Superadmin aktif', 503, origin);
  }
  if (remainingSuperadmins < 1) {
    return error('Tidak dapat menghapus seluruh Superadmin aktif (minimal 1 Superadmin harus tersisa)', 400, origin);
  }

  try {
    const existingRows = [];
    const CHUNK = 50;
    for (let i = 0; i < validIds.length; i += CHUNK) {
      const chunk = validIds.slice(i, i + CHUNK);
      const ph = chunk.map(() => '?').join(',');
      const res = await env.DB.prepare(
        `SELECT id, username, role, is_active FROM users WHERE id IN (${ph}) AND LOWER(role) != 'superadmin'`
      ).bind(...chunk).all();
      if (Array.isArray(res?.results)) existingRows.push(...res.results);
    }

    const activeIds = existingRows.map(r => Number(r.id));
    if (activeIds.length === 0) {
      return ok({ message: 'Berhasil menghapus 0 user', deleted: 0 }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
    }

    const activePh = activeIds.map(() => '?').join(',');
    const delStmt = env.DB.prepare(
      `DELETE FROM users WHERE id IN (${activePh}) AND LOWER(role) != 'superadmin'`
    ).bind(...activeIds);
    const auditStmt = buildAuditStatement(
      env,
      currentUser,
      'DELETE',
      'users',
      activeIds.slice(0, 50).join(','),
      { deleted_ids: activeIds },
      null,
      { count: activeIds.length, ids: activeIds, requested_count: validIds.length },
      'SUCCESS',
      true
    );

    const batchStmts = [delStmt, auditStmt];
    for (let i = 0; i < activeIds.length; i += CHUNK) {
      const sessChunk = activeIds.slice(i, i + CHUNK);
      const sessPh = sessChunk.map(() => '?').join(',');
      batchStmts.push(env.DB.prepare(`DELETE FROM sessions WHERE user_id IN (${sessPh})`).bind(...sessChunk));
    }

    const batchRes = await env.DB.batch(batchStmts);
    const deletedCount = Number(batchRes?.[0]?.meta?.changes ?? batchRes?.[0]?.changes ?? 0);
    return ok(
      { message: `Berhasil menghapus ${deletedCount} user`, deleted: deletedCount },
      200,
      origin,
      undefined,
      { 'X-Audit-Batched': 'true' }
    );
  } catch (err) {
    console.error('bulkDeleteUsers DB error:', err);
    return error('Gagal menghapus user', 500, origin);
  }
}

async function importUsers(request, env, currentUser, origin) {
  if (!currentUser || currentUser.role !== 'superadmin') return forbidden(origin);
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
  if (body.length === 0) return ok({ message: 'No data to import', imported: 0, created: 0, updated: 0, failed: 0, errors: [] }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });

  let iterations;
  try {
    iterations = validatePbkdf2Iterations(env?.PBKDF2_ITERATIONS);
  } catch (iterErr) {
    console.error('AUTH_CONFIG_ERROR:', iterErr?.message || iterErr);
    return error('Authentication configuration error: invalid PBKDF2_ITERATIONS', 500, origin);
  }

  const principalId = (currentUser?.id !== undefined && currentUser?.id !== null && currentUser?.id !== 0) ? currentUser.id : null;
  const principalName = currentUser?.full_name || currentUser?.username || 'System Automation';
  const principalRole = currentUser?.role || 'superadmin';
  const principalMeta = {
    id: principalId,
    username: currentUser?.username || null,
    full_name: principalName,
    role: principalRole
  };

  const batchStmts = [];
  const revokeStmts = [];
  const errors = [];
  let createdCount = 0;
  let updatedCount = 0;
  const seenUsernames = new Set();
  const seenEmails = new Set();
  let hasSuperadminOverwriteAttempt = false;
  let hasQuotaConflict = false;
  const IMPORTABLE_NEW_ROLES = ['fc_specialist_leader'];
  let projectedActiveSpecialists = await getActiveRoleCount(env, 'fc_specialist_leader', []);
  if (projectedActiveSpecialists < 0) {
    return error('Gagal memverifikasi kuota Facility Care Specialist aktif', 503, origin);
  }

  for (let i = 0; i < body.length; i++) {
    const item = body[i];
    const rowNum = i + 1;
    if (!item || typeof item !== 'object') {
      errors.push({ row: rowNum, username: null, reason: 'Invalid row object' });
      continue;
    }

    const username = item.username ? String(item.username).trim() : null;
    const email = item.email ? String(item.email).trim() : null;
    const full_name = item.full_name ? String(item.full_name).trim() : null;
    const password = item.password;
    const hasPassword =
      password !== undefined &&
      password !== null &&
      (typeof password === 'string' ? password.trim() !== '' : true);

    if (!username || !email || !full_name) {
      errors.push({ row: rowNum, username: username || null, reason: 'username, email, dan full_name wajib diisi' });
      continue;
    }
    if (hasPassword && typeof password !== 'string') {
      errors.push({ row: rowNum, username, reason: `Password untuk user '${username}' harus berupa string` });
      continue;
    }
    if (hasPassword) {
      const pwValidation = validatePasswordPolicy(password);
      if (!pwValidation.valid) {
        errors.push({ row: rowNum, username, reason: `Password untuk user '${username}': ${pwValidation.error}` });
        continue;
      }
    }

    const lowerUsername = username.toLowerCase();
    const lowerEmail = email.toLowerCase();
    if (seenUsernames.has(lowerUsername)) {
      errors.push({ row: rowNum, username, reason: `Duplicate username '${username}' di dalam payload import yang sama` });
      continue;
    }
    if (seenEmails.has(lowerEmail)) {
      errors.push({ row: rowNum, username, reason: `Duplicate email '${email}' di dalam payload import yang sama` });
      continue;
    }
    seenUsernames.add(lowerUsername);
    seenEmails.add(lowerEmail);

    const hasExplicitRole = item.role !== undefined && item.role !== null && String(item.role).trim() !== '';
    const rawRole = hasExplicitRole ? String(item.role).trim().toLowerCase() : '';
    if (rawRole === 'superadmin') {
      errors.push({ row: rowNum, username, reason: `Role 'superadmin' tidak diizinkan melalui fitur import (user: '${username}')` });
      continue;
    }

    // Lookup existing account first to protect superadmin, capture before-state, and preserve legacy roles
    let existing;
    try {
      existing = await env.DB.prepare(
        'SELECT id, username, email, full_name, role, is_active, assigned_modules FROM users WHERE LOWER(TRIM(username)) = ? OR LOWER(TRIM(email)) = ?'
      ).bind(lowerUsername, lowerEmail).first();
    } catch (lookupErr) {
      if (lookupErr.message && lookupErr.message.includes('no such column: assigned_modules')) {
        existing = await env.DB.prepare(
          'SELECT id, username, email, full_name, role, is_active FROM users WHERE LOWER(TRIM(username)) = ? OR LOWER(TRIM(email)) = ?'
        ).bind(lowerUsername, lowerEmail).first();
      } else {
        throw lookupErr;
      }
    }

    if (existing && String(existing.role || '').trim().toLowerCase() === 'superadmin') {
      hasSuperadminOverwriteAttempt = true;
      errors.push({ row: rowNum, username, reason: `Tidak dapat menimpa atau mengubah akun Superadmin ('${existing.username}') melalui fitur import` });
      continue;
    }

    if (existing && String(existing.username || '').trim().toLowerCase() !== lowerUsername) {
      errors.push({ row: rowNum, username, reason: `Email '${email}' sudah digunakan oleh akun lain ('${existing.username}')` });
      continue;
    }

    if (existing && existing.id !== undefined && existing.id !== null) {
      const emailOwner = await env.DB.prepare(
        'SELECT id, username, role FROM users WHERE LOWER(TRIM(email)) = ? AND id != ?'
      ).bind(lowerEmail, existing.id).first();
      if (emailOwner) {
        if (String(emailOwner.role || '').trim().toLowerCase() === 'superadmin') {
          hasSuperadminOverwriteAttempt = true;
          errors.push({ row: rowNum, username, reason: `Tidak dapat menggunakan email milik akun Superadmin ('${emailOwner.username}') melalui fitur import` });
        } else {
          errors.push({ row: rowNum, username, reason: `Email '${email}' sudah digunakan oleh akun lain ('${emailOwner.username}')` });
        }
        continue;
      }
    }

    let targetRole;
    if (!existing) {
      if (!hasExplicitRole || !IMPORTABLE_NEW_ROLES.includes(rawRole)) {
        errors.push({
          row: rowNum,
          username,
          reason: `Invalid role '${item.role ?? ''}' pada user baru '${username}'. Pembuatan akun baru melalui import hanya mengizinkan role: ${IMPORTABLE_NEW_ROLES.join(', ')}`
        });
        continue;
      }
      if (!hasPassword) {
        errors.push({
          row: rowNum,
          username,
          reason: `Password wajib diisi untuk pembuatan akun baru '${username}'. Gunakan password unik dan kuat (minimal 12 karakter, huruf besar, huruf kecil, angka, dan simbol)`
        });
        continue;
      }
      if (projectedActiveSpecialists + 1 > MAX_ACTIVE_FC_SPECIALISTS) {
        hasQuotaConflict = true;
        errors.push({
          row: rowNum,
          username,
          reason: `Konflik Kuota Role: Maksimal ${MAX_ACTIVE_FC_SPECIALISTS} akun aktif Facility Care Specialist telah tercapai (baris '${username}')`
        });
        continue;
      }
      targetRole = rawRole;
      projectedActiveSpecialists += 1;
    } else {
      const existingRole = String(existing.role || '').trim().toLowerCase();
      if (!hasExplicitRole) {
        if (!OFFICIAL_ROLES.includes(existingRole)) {
          errors.push({
            row: rowNum,
            username,
            reason: `User existing '${username}' memiliki role non-resmi '${existingRole}'. Wajib tetapkan role resmi (${OFFICIAL_ROLES.join(', ')}) atau gunakan import untuk migrasi ke fc_specialist_leader.`
          });
          continue;
        }
        targetRole = existingRole;
      } else if (OFFICIAL_ROLES.includes(rawRole) && (rawRole === existingRole || IMPORTABLE_NEW_ROLES.includes(rawRole))) {
        targetRole = rawRole;
      } else {
        errors.push({
          row: rowNum,
          username,
          reason: `Invalid role '${item.role}' pada update user existing '${username}'. Role resmi yang diizinkan: ${IMPORTABLE_NEW_ROLES.join(', ')}`
        });
        continue;
      }
      const existingIsActive = Number(existing.is_active ?? 1) === 1;
      if (targetRole === 'fc_specialist_leader' && existingRole !== 'fc_specialist_leader' && existingIsActive) {
        if (projectedActiveSpecialists + 1 > MAX_ACTIVE_FC_SPECIALISTS) {
          hasQuotaConflict = true;
          errors.push({
            row: rowNum,
            username,
            reason: `Konflik Kuota Role: Tidak dapat mengubah '${username}' menjadi fc_specialist_leader karena maksimal ${MAX_ACTIVE_FC_SPECIALISTS} akun aktif telah tercapai`
          });
          continue;
        }
        projectedActiveSpecialists += 1;
      } else if (existingRole === 'fc_specialist_leader' && targetRole !== 'fc_specialist_leader' && existingIsActive) {
        projectedActiveSpecialists = Math.max(0, projectedActiveSpecialists - 1);
      }
    }

    const rawModulesForValidation =
      existing && targetRole === 'editor_khusus' && item.assigned_modules === undefined
        ? (existing.assigned_modules ?? '[]')
        : item.assigned_modules;
    const moduleValidation = validateAndNormalizeAssignedModules(rawModulesForValidation, targetRole);
    if (!moduleValidation.valid) {
      errors.push({ row: rowNum, username, reason: moduleValidation.error });
      continue;
    }

    const passwordChanged = !existing || Boolean(hasPassword);
    const hash = passwordChanged ? await hashPassword(password, iterations) : null;
    const targetUsername = existing && existing.username ? existing.username : username;

    seenUsernames.add(lowerUsername);
    seenEmails.add(lowerEmail);

    if (!existing) {
      // Kasus A: New account creation via Bulk Import -> CREATE
      const insertUserStmt = env.DB.prepare(
        `INSERT INTO users (username, email, password_hash, full_name, role, assigned_modules)
         VALUES (?, ?, ?, ?, ?, ?)`
      ).bind(
        targetUsername,
        email,
        hash,
        full_name,
        targetRole,
        moduleValidation.modules
      );

      const newSnapshot = sanitizeAuditData({
        username: targetUsername,
        email,
        full_name,
        role: targetRole,
        is_active: 1,
        assigned_modules: moduleValidation.modules,
        source: 'Bulk Import'
      });
      const newStr = JSON.stringify(newSnapshot);
      const createDetailsStr = JSON.stringify({
        source: 'Bulk Import',
        import_mode: 'BULK_IMPORT',
        operation: 'CREATE',
        row: rowNum,
        username: targetUsername,
        email,
        role: targetRole,
        principal: principalMeta,
        status: 'SUCCESS'
      });
      const createAuditStmt = env.DB.prepare(
        `INSERT INTO audit_logs (user_id, user_name, user_role, action, module, target_id, old_data, new_data, old_value, new_value, details, source)
         SELECT ?, ?, ?, 'CREATE', 'users', CASE WHEN last_insert_rowid() > 0 THEN CAST(last_insert_rowid() AS TEXT) ELSE ? END, NULL, ?, NULL, ?, ?, 'Bulk Import'
         WHERE (SELECT changes()) > 0`
      ).bind(
        principalId,
        principalName,
        principalRole,
        targetUsername,
        newStr,
        newStr,
        createDetailsStr
      );

      batchStmts.push([insertUserStmt, createAuditStmt]);
      createdCount++;
    } else {
      // Kasus B: Updating existing user via Bulk Import -> UPDATE with before/after snapshot
      const existingRole = String(existing.role || '').trim().toLowerCase();
      const fieldsUpdated = [];
      if (email !== String(existing.email || '').trim().toLowerCase()) fieldsUpdated.push('email');
      if (full_name !== String(existing.full_name || '').trim()) fieldsUpdated.push('full_name');
      const roleChanged = targetRole !== existingRole;
      if (roleChanged) fieldsUpdated.push('role');
      const assignedModulesChanged = (moduleValidation.modules ?? null) !== (existing.assigned_modules ?? null);
      if (assignedModulesChanged) {
        fieldsUpdated.push('assigned_modules');
      }
      if (passwordChanged) {
        fieldsUpdated.push('password');
      }

      const updateUserStmt = passwordChanged
        ? env.DB.prepare(
            `UPDATE users
             SET email = ?,
                 full_name = ?,
                 role = ?,
                 assigned_modules = ?,
                 password_hash = ?,
                 updated_at = datetime('now')
             WHERE id = ? AND LOWER(role) != 'superadmin'`
          ).bind(
            email,
            full_name,
            targetRole,
            moduleValidation.modules,
            hash,
            existing.id
          )
        : env.DB.prepare(
            `UPDATE users
             SET email = ?,
                 full_name = ?,
                 role = ?,
                 assigned_modules = ?,
                 updated_at = datetime('now')
             WHERE id = ? AND LOWER(role) != 'superadmin'`
          ).bind(
            email,
            full_name,
            targetRole,
            moduleValidation.modules,
            existing.id
          );

      const shouldRevokeSessions = passwordChanged || roleChanged || assignedModulesChanged;

      const beforeSnapshot = sanitizeAuditData({
        id: existing.id,
        username: existing.username,
        email: existing.email,
        full_name: existing.full_name,
        role: existingRole,
        is_active: existing.is_active ?? 1,
        assigned_modules: existing.assigned_modules ?? null
      });
      const afterSnapshot = sanitizeAuditData({
        id: existing.id,
        username: targetUsername,
        email,
        full_name,
        role: targetRole,
        is_active: existing.is_active ?? 1,
        assigned_modules: moduleValidation.modules,
        source: 'Bulk Import'
      });
      const oldStr = JSON.stringify(beforeSnapshot);
      const newStr = JSON.stringify(afterSnapshot);
      const updateDetailsStr = JSON.stringify({
        source: 'Bulk Import',
        import_mode: 'BULK_IMPORT',
        operation: 'UPDATE',
        row: rowNum,
        user_id: existing.id,
        username: targetUsername,
        fields_updated: fieldsUpdated,
        password_changed: passwordChanged,
        role_changed: roleChanged,
        role_before: existingRole,
        role_after: targetRole,
        before: {
          full_name: existing.full_name,
          email: existing.email,
          role: existingRole,
          assigned_modules: existing.assigned_modules ?? null
        },
        after: {
          full_name,
          email,
          role: targetRole,
          assigned_modules: moduleValidation.modules,
          password_changed: passwordChanged
        },
        sessions_revoked: shouldRevokeSessions,
        principal: principalMeta,
        status: 'SUCCESS'
      });
      const updateAuditStmt = env.DB.prepare(
        `INSERT INTO audit_logs (user_id, user_name, user_role, action, module, target_id, old_data, new_data, old_value, new_value, details, source)
         SELECT ?, ?, ?, 'UPDATE', 'users', ?, ?, ?, ?, ?, ?, 'Bulk Import'
         WHERE (SELECT changes()) > 0`
      ).bind(
        principalId,
        principalName,
        principalRole,
        String(existing.id),
        oldStr,
        newStr,
        oldStr,
        newStr,
        updateDetailsStr
      );

      const userGroup = [updateUserStmt, updateAuditStmt];
      // Revoke existing sessions if import updated an existing user's credentials, role, or module scope
      if (shouldRevokeSessions && existing.id) {
        userGroup.push(buildRevokeAllUserSessionsStatement(env, existing.id));
      }
      batchStmts.push(userGroup);
      updatedCount++;
    }
  }

  const importedCount = createdCount + updatedCount;
  if (importedCount === 0 && errors.length > 0) {
    const status = hasSuperadminOverwriteAttempt ? 403 : (hasQuotaConflict ? 409 : 400);
    return new Response(
      JSON.stringify({ success: false, error: errors[0].reason, failed: errors.length, errors }),
      {
        status,
        headers: { 'Content-Type': 'application/json', 'X-Audit-Batched': 'true', ...corsHeaders(origin) }
      }
    );
  }

  try {
    if (importedCount > 0) {
      const userChunks = packStatementGroups(batchStmts.map((stmts) => ({ stmts })), 76);
      for (const chunk of userChunks) {
        await env.DB.batch(chunk.stmts);
      }
    }
    return ok(
      {
        message: `Berhasil mengimport ${importedCount} user`,
        imported: importedCount,
        created: createdCount,
        updated: updatedCount,
        failed: errors.length,
        errors
      },
      200,
      origin,
      undefined,
      { 'X-Audit-Batched': 'true' }
    );
  } catch (err) {
    const msg = err?.message || String(err);
    if (
      msg.includes('MAX_ACTIVE_SUPERADMINS_EXCEEDED') ||
      msg.includes('MAX_ACTIVE_FC_SPECIALIST_EXCEEDED')
    ) {
      return error('Konflik Kuota Role: Batas maksimal 2 akun aktif untuk role ini telah tercapai.', 409, origin);
    }
    console.error('importUsers DB error:', err);
    return error('Gagal mengimport data user', 500, origin);
  }
}
