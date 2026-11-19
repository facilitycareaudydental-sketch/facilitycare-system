export const SENSITIVE_AUDIT_KEYS = new Set([
  'password',
  'password_hash',
  'token',
  'jwt',
  'secret',
  'authorization',
  'auth_token',
  'access_token',
  'refresh_token',
  'api_key',
  'private_key',
  'credentials',
  'signature',
  'x-fcms-signature',
  'hmac'
]);

/**
 * Recursively sanitizes data to prevent credential/secret leakage into audit logs.
 * Handles nested objects, arrays, and JSON strings.
 */
export function sanitizeAuditData(data) {
  if (data === null || data === undefined) return null;

  if (typeof data === 'string') {
    const trimmed = data.trim();
    if ((trimmed.startsWith('{') && trimmed.endsWith('}')) || (trimmed.startsWith('[') && trimmed.endsWith(']'))) {
      try {
        const parsed = JSON.parse(trimmed);
        if (typeof parsed === 'object' && parsed !== null) {
          return JSON.stringify(sanitizeAuditData(parsed));
        }
      } catch {
        // Not valid JSON string, continue below
      }
    }
    return data;
  }

  if (typeof data !== 'object') return data;

  if (Array.isArray(data)) {
    return data.map(item => sanitizeAuditData(item));
  }

  const sanitized = {};
  for (const [key, value] of Object.entries(data)) {
    const lowerKey = key.toLowerCase();
    if (
      SENSITIVE_AUDIT_KEYS.has(lowerKey) ||
      lowerKey.includes('password') ||
      lowerKey.includes('secret') ||
      lowerKey.includes('token') ||
      lowerKey.includes('jwt') ||
      lowerKey.includes('signature') ||
      lowerKey.includes('hmac') ||
      lowerKey.includes('api_key') ||
      lowerKey.includes('private_key') ||
      lowerKey.includes('credential') ||
      lowerKey.includes('authorization')
    ) {
      sanitized[key] = '[REDACTED]';
    } else if ((typeof value === 'object' && value !== null) || typeof value === 'string') {
      sanitized[key] = sanitizeAuditData(value);
    } else {
      sanitized[key] = value;
    }
  }
  return sanitized;
}

export const OFFICIAL_ACTIONS = Object.freeze(new Set(['CREATE', 'UPDATE', 'DELETE', 'LOGIN', 'LOGOUT']));

export function normalizeAuditAction(action) {
  const act = String(action || 'UPDATE').toUpperCase().trim();
  if (OFFICIAL_ACTIONS.has(act)) return act;
  if (act.includes('CREATE') || act.includes('INSERT') || act.includes('BOOTSTRAP')) return 'CREATE';
  if (act.includes('DELETE') || act.includes('REMOVE')) return 'DELETE';
  if (act.includes('LOGIN')) return 'LOGIN';
  if (act.includes('LOGOUT')) return 'LOGOUT';
  return 'UPDATE';
}

/**
 * Prepares an INSERT statement for audit_logs without executing it,
 * enabling atomic bundling into env.DB.batch([...]) transactions.
 */
export function buildAuditStatement(env, user, action, moduleName, targetId, oldData = null, newData = null, details = null, status = 'SUCCESS', requireChanges = false) {
  const userId = (user?.id !== undefined && user?.id !== null && user?.id !== 0) ? user.id : null;
  const userName = user?.full_name || user?.username || 'System Automation';
  const userRole = user?.role || 'system';
  const normalizedAction = normalizeAuditAction(action);
  const auditStatus = status ? String(status).toUpperCase().trim() : 'SUCCESS';

  const sanitizedOld = sanitizeAuditData(oldData);
  const sanitizedNew = sanitizeAuditData(newData);

  const safeStringify = (data) => {
    if (!data) return null;
    if (typeof data === 'string') return data;
    try { return JSON.stringify(data); } catch { return String(data); }
  };

  const oldStr = safeStringify(sanitizedOld);
  const newStr = safeStringify(sanitizedNew);

  // Embed status and structured metadata cleanly into details (always sanitized)
  let combinedDetails = details;
  if (typeof details === 'object' && details !== null) {
    combinedDetails = sanitizeAuditData({ ...details, status: auditStatus });
  } else if (details) {
    combinedDetails = sanitizeAuditData({ info: String(details), status: auditStatus });
  } else {
    combinedDetails = { status: auditStatus };
  }
  const detailsStr = safeStringify(combinedDetails);

  const sql = requireChanges
    ? `INSERT INTO audit_logs (user_id, user_name, user_role, action, module, target_id, old_data, new_data, old_value, new_value, details)
       SELECT ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
       WHERE (SELECT changes()) > 0`
    : `INSERT INTO audit_logs (user_id, user_name, user_role, action, module, target_id, old_data, new_data, old_value, new_value, details)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`;

  return env.DB.prepare(sql).bind(
    userId, userName, userRole, normalizedAction, moduleName, String(targetId), oldStr, newStr, oldStr, newStr, detailsStr
  );
}

/**
 * Prepares an atomic CREATE audit log statement intended to immediately follow
 * a single INSERT statement inside env.DB.batch([...]). Captures last_insert_rowid()
 * atomically in SQLite/D1 and guards with WHERE (SELECT changes()) > 0.
 */
export function buildCreateAuditStatement(env, user, moduleName, fallbackTargetId, newData = null, details = null) {
  const userId = (user?.id !== undefined && user?.id !== null && user?.id !== 0) ? user.id : null;
  const userName = user?.full_name || user?.username || 'System Automation';
  const userRole = user?.role || 'system';
  const newStr = newData ? JSON.stringify(sanitizeAuditData(newData)) : null;
  const combinedDetails = sanitizeAuditData(
    typeof details === 'object' && details !== null
      ? { ...details, status: 'SUCCESS' }
      : details
        ? { info: String(details), status: 'SUCCESS' }
        : { status: 'SUCCESS' }
  );
  const detailsStr = JSON.stringify(combinedDetails);

  return env.DB.prepare(
    `INSERT INTO audit_logs (user_id, user_name, user_role, action, module, target_id, old_data, new_data, old_value, new_value, details)
     SELECT ?, ?, ?, 'CREATE', ?, COALESCE(CAST(last_insert_rowid() AS TEXT), ?), NULL, ?, NULL, ?, ?
     WHERE (SELECT changes()) > 0`
  ).bind(
    userId,
    userName,
    userRole,
    moduleName,
    String(fallbackTargetId ?? 'new'),
    newStr,
    newStr,
    detailsStr
  );
}

/**
 * Prepares an atomic CREATE audit log statement when an INSERT into the entity table
 * is immediately followed by buildOutboxFromLastInsertQuery(..., correlationId), which
 * overwrites SQLite's last_insert_rowid(). Resolves the exact entity_id from sync_outbox
 * inside the same D1 batch transaction.
 */
export function buildCreateAuditFromOutboxStatement(env, user, moduleName, correlationId, newData = null, details = null) {
  const userId = (user?.id !== undefined && user?.id !== null && user?.id !== 0) ? user.id : null;
  const userName = user?.full_name || user?.username || 'System Automation';
  const userRole = user?.role || 'system';
  const newStr = newData ? JSON.stringify(sanitizeAuditData(newData)) : null;
  const combinedDetails = sanitizeAuditData(
    typeof details === 'object' && details !== null
      ? { ...details, status: 'SUCCESS' }
      : details
        ? { info: String(details), status: 'SUCCESS' }
        : { status: 'SUCCESS' }
  );
  const detailsStr = JSON.stringify(combinedDetails);

  return env.DB.prepare(
    `INSERT INTO audit_logs (user_id, user_name, user_role, action, module, target_id, old_data, new_data, old_value, new_value, details)
     SELECT ?, ?, ?, 'CREATE', ?, entity_id, NULL, ?, NULL, ?, ?
     FROM sync_outbox
     WHERE correlation_id = ?
     LIMIT 1`
  ).bind(
    userId,
    userName,
    userRole,
    moduleName,
    newStr,
    newStr,
    detailsStr,
    correlationId
  );
}

/**
 * Executes interleaved [mutationStmt, auditStmt, ...] import batches atomically.
 * Keeps every (mutation, audit) pair intact inside the same env.DB.batch() call.
 */
export async function executeAtomicImportBatch(env, stmts, maxStmtsPerBatch = 80) {
  if (!Array.isArray(stmts) || stmts.length === 0) return [];
  const step = Math.max(2, maxStmtsPerBatch - (maxStmtsPerBatch % 2));
  if (stmts.length <= step) {
    return await env.DB.batch(stmts);
  }
  const results = [];
  for (let i = 0; i < stmts.length; i += step) {
    const chunkRes = await env.DB.batch(stmts.slice(i, i + step));
    if (Array.isArray(chunkRes)) results.push(...chunkRes);
  }
  return results;
}

/**
 * Prepares a per-record DELETE audit statement for soft-delete operations (single or bulk).
 * Uses EXISTS on the exact batchDeletedAt timestamp so TOCTOU-skipped rows never receive
 * phantom audit entries while every actually modified row gets its own snapshot and batch metadata.
 */
export function buildSoftDeleteRecordAuditStatement(env, user, moduleName, tableName, pkCol, softDeleteCol, batchDeletedAt, row, details = null) {
  const userId = (user?.id !== undefined && user?.id !== null && user?.id !== 0) ? user.id : null;
  const userName = user?.full_name || user?.username || 'System Automation';
  const userRole = user?.role || 'system';
  const targetId = String(row?.[pkCol] ?? '');
  const sanitizedOld = sanitizeAuditData(row);
  const oldStr = sanitizedOld ? JSON.stringify(sanitizedOld) : null;
  const combinedDetails = sanitizeAuditData(
    typeof details === 'object' && details !== null
      ? { ...details, soft_deleted: true, deleted_at: batchDeletedAt, status: 'SUCCESS' }
      : { soft_deleted: true, deleted_at: batchDeletedAt, status: 'SUCCESS' }
  );
  const detailsStr = JSON.stringify(combinedDetails);

  return env.DB.prepare(
    `INSERT INTO audit_logs (user_id, user_name, user_role, action, module, target_id, old_data, new_data, old_value, new_value, details)
     SELECT ?, ?, ?, 'DELETE', ?, ?, ?, NULL, ?, NULL, ?
     WHERE EXISTS (SELECT 1 FROM ${tableName} WHERE ${pkCol} = ? AND ${softDeleteCol} = ?)`
  ).bind(
    userId,
    userName,
    userRole,
    moduleName,
    targetId,
    oldStr,
    oldStr,
    detailsStr,
    row?.[pkCol],
    batchDeletedAt
  );
}

/**
 * Prepares chunked per-record DELETE audit statements for hard-delete bulk operations
 * within D1's 100 bind-parameter limit (8 rows * 11 params = 88 params per chunk).
 * Guarded with WHERE (SELECT changes()) > 0 right after the primary DELETE statement.
 */
export function buildBulkHardDeleteAuditStatements(env, user, moduleName, rows, pkCol, details = null) {
  if (!Array.isArray(rows) || rows.length === 0) return [];
  const userId = (user?.id !== undefined && user?.id !== null && user?.id !== 0) ? user.id : null;
  const userName = user?.full_name || user?.username || 'System Automation';
  const userRole = user?.role || 'system';
  const combinedDetails = sanitizeAuditData(
    typeof details === 'object' && details !== null
      ? { ...details, status: 'SUCCESS' }
      : { status: 'SUCCESS' }
  );
  const detailsStr = JSON.stringify(combinedDetails);

  const CHUNK_SIZE = 8; // 8 * 11 = 88 <= 100 D1 bind limit
  const stmts = [];
  for (let i = 0; i < rows.length; i += CHUNK_SIZE) {
    const chunk = rows.slice(i, i + CHUNK_SIZE);
    const params = [];
    for (const row of chunk) {
      const sanitizedOld = sanitizeAuditData(row);
      const oldStr = sanitizedOld ? JSON.stringify(sanitizedOld) : null;
      params.push(
        userId,
        userName,
        userRole,
        'DELETE',
        moduleName,
        String(row?.[pkCol] ?? ''),
        oldStr,
        null,
        oldStr,
        null,
        detailsStr
      );
    }
    const selectArms = chunk.map((_, idx) =>
      idx === 0
        ? 'SELECT ? AS user_id, ? AS user_name, ? AS user_role, ? AS action, ? AS module, ? AS target_id, ? AS old_data, ? AS new_data, ? AS old_value, ? AS new_value, ? AS details'
        : 'SELECT ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?'
    ).join('\n      UNION ALL\n      ');

    stmts.push(
      env.DB.prepare(
        `INSERT INTO audit_logs (user_id, user_name, user_role, action, module, target_id, old_data, new_data, old_value, new_value, details)
         SELECT v.user_id, v.user_name, v.user_role, v.action, v.module, v.target_id, v.old_data, v.new_data, v.old_value, v.new_value, v.details
         FROM (
           ${selectArms}
         ) AS v
         WHERE (SELECT changes()) > 0`
      ).bind(...params)
    );
  }
  return stmts;
}

/**
 * Executes a mutation statement (or array of statements) atomically bundled with
 * its corresponding audit log statement inside a single Cloudflare D1 batch transaction.
 * If either the mutation or the audit log insertion fails, the entire transaction is rolled back.
 * requireChanges (default true): ensures the audit log is only inserted if preceding mutation modified > 0 rows.
 */
export async function executeAtomicDelete(env, user, mutationStmts, moduleName, targetId, oldData = null, details = null, requireChanges = true) {
  const stmts = Array.isArray(mutationStmts) ? [...mutationStmts] : [mutationStmts];
  const auditStmt = buildAuditStatement(env, user, 'DELETE', moduleName, targetId, oldData, null, details, 'SUCCESS', requireChanges);
  stmts.push(auditStmt);
  return await env.DB.batch(stmts);
}

export async function logAudit(env, user, action, moduleName, targetId, oldData = null, newData = null, details = null, status = 'SUCCESS') {
  try {
    const stmt = buildAuditStatement(env, user, action, moduleName, targetId, oldData, newData, details, status);
    await stmt.run();
  } catch (error) {
    console.error('Failed to write audit log:', error?.message || error);
  }
}
