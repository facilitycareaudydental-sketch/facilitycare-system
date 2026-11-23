import { error, ok } from './response.js';
import { mapPayloadToDB } from './sync_mapper.js';
import { constantTimeEqual } from './auth.js';
import { buildAuditStatement } from './audit.js';
import { processOutboxEvents } from './outbox_processor.js';
import { resolveSyncMode } from './sync_mode.js';
import { resolveMaintenanceFreezeState } from './maintenance.js';
import { getManagementBranchSqlPredicate } from './branchMaster.js';

export const MAX_WEBHOOK_PAYLOAD_BYTES = 1024 * 1024; // 1 MB limit
export const MAX_TIMESTAMP_AGE_MS = 5 * 60 * 1000; // 5 minutes max age
export const MAX_FUTURE_SKEW_MS = 60 * 1000; // 60 seconds future skew tolerance

const WEBHOOK_RATE_LIMIT_WINDOW_MS = 60 * 1000; // 1 minute
const MAX_WEBHOOK_REQUESTS_PER_WINDOW = 120; // 120 requests/min per client
const webhookRateLimitMap = new Map();

export class WebhookValidationError extends Error {
  constructor(message, status = 400) {
    super(message);
    this.name = 'WebhookValidationError';
    this.status = status;
  }
}

export class WebhookConflictError extends Error {
  constructor(message, status = 409, details = null) {
    super(message);
    this.name = 'WebhookConflictError';
    this.status = status;
    this.details = details;
  }
}

export function checkWebhookRateLimit(clientId, now = Date.now()) {
  if (webhookRateLimitMap.size > 500) {
    for (const [key, val] of webhookRateLimitMap.entries()) {
      if (now - val.windowStart > WEBHOOK_RATE_LIMIT_WINDOW_MS) {
        webhookRateLimitMap.delete(key);
      }
    }
  }

  const record = webhookRateLimitMap.get(clientId);
  if (!record || now - record.windowStart > WEBHOOK_RATE_LIMIT_WINDOW_MS) {
    webhookRateLimitMap.set(clientId, { windowStart: now, count: 1 });
    return { allowed: true };
  }

  if (record.count >= MAX_WEBHOOK_REQUESTS_PER_WINDOW) {
    const retryAfter = Math.ceil((record.windowStart + WEBHOOK_RATE_LIMIT_WINDOW_MS - now) / 1000);
    return { allowed: false, retryAfter: Math.max(1, retryAfter) };
  }

  record.count += 1;
  return { allowed: true };
}

export function resetWebhookRateLimit() {
  webhookRateLimitMap.clear();
}

export function hexToBytes(hex) {
  if (typeof hex !== 'string' || !/^[0-9a-f]{64}$/i.test(hex)) return null;
  const bytes = new Uint8Array(32);
  for (let i = 0; i < 32; i++) {
    bytes[i] = Number.parseInt(hex.slice(i * 2, i * 2 + 2), 16);
  }
  return bytes;
}

/**
 * Handle incoming Webhook from Google Apps Script (Sheets -> FCMS)
 */
export async function receiveWebhook(request, env, origin) {
  const startTime = Date.now();
  try {
    // 1. Method verification
    if (request.method !== 'POST') {
      return error('Method not allowed', 405, origin);
    }

    // 2. Request size verification (via Header)
    const contentLengthHeader = request.headers.get('content-length');
    if (contentLengthHeader) {
      const contentLength = Number(contentLengthHeader);
      if (!Number.isNaN(contentLength) && contentLength > MAX_WEBHOOK_PAYLOAD_BYTES) {
        return error('Payload too large', 413, origin);
      }
    }

    // 3. Rate limiting (In-Isolate sliding window; cross-isolate enforcement relies on Cloudflare WAF edge rules)
    const isDeployedEnv = env?.ENVIRONMENT === 'production' || env?.ENVIRONMENT === 'staging';
    const cfIp = request.headers.get('cf-connecting-ip');
    const clientIp = cfIp || (isDeployedEnv ? 'untrusted-edge' : (request.headers.get('x-forwarded-for')?.split(',')[0]?.trim() || 'local-client'));
    const rateLimit = checkWebhookRateLimit(clientIp);
    if (!rateLimit.allowed) {
      return error(`Too many webhook requests. Please try again in ${rateLimit.retryAfter} seconds.`, 429, origin);
    }

    // 4. Bearer authentication via constant-time comparison
    const authHeader = request.headers.get('Authorization');
    if (!env?.SYNC_WEBHOOK_SECRET || typeof env.SYNC_WEBHOOK_SECRET !== 'string' || env.SYNC_WEBHOOK_SECRET.trim() === '') {
      return error('Unauthorized Webhook', 401, origin);
    }
    if (!authHeader || typeof authHeader !== 'string' || !authHeader.startsWith('Bearer ')) {
      return error('Unauthorized Webhook', 401, origin);
    }
    const token = authHeader.slice(7).trim();
    if (!constantTimeEqual(token, env.SYNC_WEBHOOK_SECRET)) {
      return error('Unauthorized Webhook', 401, origin);
    }

    // 5. Read raw body once and check body byte length
    const rawBody = await request.text();
    if (rawBody.length > MAX_WEBHOOK_PAYLOAD_BYTES) {
      return error('Payload too large', 413, origin);
    }

    // 6. Validate signature format & HMAC verification over raw body
    const signatureHeader = request.headers.get('X-FCMS-SIGNATURE');
    if (!signatureHeader || typeof signatureHeader !== 'string') {
      return error('Missing or malformed webhook signature', 401, origin);
    }
    const signatureBytes = hexToBytes(signatureHeader.trim());
    if (!signatureBytes) {
      return error('Missing or malformed webhook signature', 401, origin);
    }

    const encoder = new TextEncoder();
    const key = await crypto.subtle.importKey(
      'raw',
      encoder.encode(env.SYNC_WEBHOOK_SECRET),
      { name: 'HMAC', hash: 'SHA-256' },
      false,
      ['verify']
    );
    const signatureValid = await crypto.subtle.verify(
      'HMAC',
      key,
      signatureBytes,
      encoder.encode(rawBody)
    );
    if (!signatureValid) {
      return error('Invalid Signature', 403, origin);
    }

    // 6b. Fast env-level Maintenance Freeze check before DB lookup (zero extra DB calls in DRY_RUN)
    const fastFreezeState = await resolveMaintenanceFreezeState(env, { skipDbProbe: true });
    if (fastFreezeState.freezeActive) {
      return error('Service temporarily unavailable due to Maintenance Freeze', 503, origin);
    }

    // 7. Parse JSON strictly AFTER HMAC verification succeeds
    let payload;
    try {
      payload = JSON.parse(rawBody);
    } catch {
      return error('Invalid JSON', 400, origin);
    }
    if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
      return error('Invalid JSON payload structure', 400, origin);
    }

    const { event_id, sheet_name, action, data, payload_version, correlation_id, timestamp, source } = payload;

    // 8. Payload Version Validation
    if (payload_version !== '1.0') {
      return error('Unsupported payload version', 400, origin);
    }

    // Sync Loop Prevention: Never re-ingest an event originated by FCMS outbound sync
    if (source === 'FCMS') {
      const loopEventId = typeof event_id === 'string' ? event_id.trim() : '';
      return ok({ message: 'Ignored FCMS-originated event (sync loop prevented)', event_id: loopEventId, ignored: true, loopPrevented: true }, 200, origin);
    }

    // 9. Timestamp Freshness Validation
    if (!timestamp || typeof timestamp !== 'string') {
      return error('Missing or invalid timestamp in payload', 400, origin);
    }
    const eventTime = new Date(timestamp).getTime();
    if (Number.isNaN(eventTime)) {
      return error('Invalid timestamp format', 400, origin);
    }
    const now = Date.now();
    if (now - eventTime > MAX_TIMESTAMP_AGE_MS) {
      return error('Webhook timestamp has expired (stale event)', 400, origin);
    }
    if (eventTime - now > MAX_FUTURE_SKEW_MS) {
      return error('Webhook timestamp is in the future beyond acceptable clock skew', 400, origin);
    }

    // 10. Event ID and Payload Structure Validation
    if (!event_id || typeof event_id !== 'string') {
      return error('Invalid event_id: must be a non-empty string', 400, origin);
    }
    const cleanEventId = event_id.trim();
    if (!/^[a-zA-Z0-9_-]{8,64}$/.test(cleanEventId)) {
      return error('Invalid event_id format: must be 8-64 alphanumeric characters', 400, origin);
    }
    if (!sheet_name || typeof sheet_name !== 'string') {
      return error('Invalid sheet_name', 400, origin);
    }
    if (!action || !['INSERT', 'UPDATE', 'DELETE'].includes(action)) {
      return error('Invalid action', 400, origin);
    }
    if (!data || typeof data !== 'object' || Array.isArray(data)) {
      return error('Invalid data payload', 400, origin);
    }

    // 11. Fast Idempotency Read Check (Replay Prevention)
    const isProcessed = await env.DB.prepare('SELECT 1 FROM sync_idempotency WHERE event_id = ?').bind(cleanEventId).first();
    if (isProcessed) {
      return ok({ message: 'Event already processed', event_id: cleanEventId, duplicate: true }, 200, origin);
    }

    // 12. Fail-Closed SYNC_MODE validation (only explicit SYNC_MODE === 'LIVE' in non-staging may mutate DB)
    const syncState = resolveSyncMode(env);
    if (!syncState.isLive) {
      await handleGenericWebhook(env, sheet_name, action, data, payload_version, correlation_id, cleanEventId, startTime, { dryRun: true });
      return ok({
        message: 'Webhook validated; no database writes performed',
        event_id: cleanEventId,
        dryRun: true,
        sync_mode: syncState.mode,
        ...(syncState.valid ? {} : { failClosed: true, reason: syncState.reason })
      }, 200, origin);
    }

    // 12b. Full DB-backed Maintenance Freeze validation prior to any LIVE mutation
    const dbFreezeState = await resolveMaintenanceFreezeState(env);
    if (dbFreezeState.freezeActive) {
      return error('Service temporarily unavailable due to Maintenance Freeze', 503, origin);
    }

    // 13. Atomic Idempotency & Database Mutation & Audit Execution
    const result = await handleGenericWebhook(env, sheet_name, action, data, payload_version, correlation_id, cleanEventId, startTime, { dryRun: false });
    return ok({ message: 'Webhook received successfully', event_id: cleanEventId, ...result }, 200, origin);

  } catch (err) {
    if (err instanceof WebhookValidationError) {
      return error(err.message, err.status, origin);
    }
    if (err instanceof WebhookConflictError) {
      return error(err.message, err.status, origin, err.details);
    }
    const errMsg = String(err?.message || err);
    if (errMsg.includes('UNIQUE constraint failed: sync_idempotency.event_id') || errMsg.includes('sync_idempotency')) {
      return ok({ message: 'Event already processed', duplicate: true }, 200, origin);
    }
    console.error('Webhook processing error:', err);
    return error('Internal webhook error', 500, origin);
  }
}

/**
 * Process the Transactional Outbox (FCMS -> Sheets)
 * This runs via Cloudflare Cron every minute.
 *
 * Without an explicit local adapter/simulation option, LIVE remains strictly
 * fail-closed (OUTBOX_DELIVERY_NOT_IMPLEMENTED) and DRY_RUN performs zero DB writes.
 */
export async function processOutbox(env, options = {}) {
  const fastFreezeState = await resolveMaintenanceFreezeState(env, { skipDbProbe: true });
  if (fastFreezeState.freezeActive) {
    return { blocked: true, processed: 0, reason: 'MAINTENANCE_FREEZE_ACTIVE', source: fastFreezeState.source };
  }

  if (options && (options.adapter || options.simulateDryRun === true)) {
    return processOutboxEvents(env, options);
  }

  const syncState = resolveSyncMode(env);
  if (!syncState.isLive) {
    console.log(`[${syncState.mode}] processOutbox dry-run/fail-closed mode active (${syncState.reason || 'DRY_RUN'}). No external push or outbox modification.`);
    return { dryRun: true, processed: 0 };
  }

  // No remote outbound Google Sheets delivery adapter is enabled in this module.
  // Fail closed before any DB operation; never delete or mark an event successful
  // until a real remote delivery has been confirmed.
  console.warn('[SYNC_BLOCKED] Outbox delivery adapter is not implemented; queue left unchanged.');
  return { blocked: true, processed: 0, reason: 'OUTBOX_DELIVERY_NOT_IMPLEMENTED' };
}

/**
 * Helper to queue an event to the outbox from other modules.
 */
export function buildOutboxQuery(env, entityName, entityId, action, payload, correlationId = null) {
  const id = crypto.randomUUID();
  return env.DB.prepare(`
    INSERT INTO sync_outbox (id, entity_name, entity_id, action, payload, correlation_id) 
    VALUES (?, ?, ?, ?, ?, ?)
  `).bind(id, entityName, entityId, action, JSON.stringify(payload), correlationId);
}

/**
 * Conditional outbox enqueue for atomic UPDATE batches: only inserts into
 * sync_outbox when the preceding UPDATE modified at least 1 row.
 */
export function buildConditionalOutboxQuery(env, entityName, entityId, action, payload, correlationId = null) {
  const id = crypto.randomUUID();
  return env.DB.prepare(`
    INSERT INTO sync_outbox (id, entity_name, entity_id, action, payload, correlation_id)
    SELECT ?, ?, ?, ?, ?, ?
    WHERE (SELECT changes()) > 0
  `).bind(id, entityName, String(entityId), action, JSON.stringify(payload), correlationId);
}

/**
 * Atomic outbox enqueue for INSERT batches: captures last_insert_rowid() in SQLite/D1
 * and injects FCMS_ID into the JSON payload in the same D1 batch transaction.
 */
export function buildOutboxFromLastInsertQuery(env, entityName, payload, correlationId = null) {
  const id = crypto.randomUUID();
  return env.DB.prepare(`
    INSERT INTO sync_outbox (id, entity_name, entity_id, action, payload, correlation_id)
    SELECT ?, ?, CAST(last_insert_rowid() AS TEXT), 'INSERT', json_set(?, '$.FCMS_ID', last_insert_rowid()), ?
    WHERE last_insert_rowid() > 0
  `).bind(id, entityName, JSON.stringify(payload), correlationId);
}

export const OUTBOX_CHUNK_SIZE = 15; // 15 rows * 6 columns = 90 bind params (within D1 100-param limit)

/**
 * Builds chunked multi-row INSERT statements for sync_outbox suitable for env.DB.batch().
 * Honors Cloudflare D1 limits (<= 100 bound parameters per statement) and optionally
 * guards execution with WHERE (SELECT changes()) > 0 so no phantom events are queued
 * when the preceding mutation modified 0 rows.
 */
export function buildBulkOutboxStatements(env, entries, options = {}) {
  if (!Array.isArray(entries) || entries.length === 0) return [];
  const rawChunk = Number(options.chunkSize || OUTBOX_CHUNK_SIZE);
  const chunkSize = Math.max(1, Math.min(16, Number.isInteger(rawChunk) ? rawChunk : OUTBOX_CHUNK_SIZE));
  const requireChanges = options.requireChanges !== false;

  const stmts = [];
  for (let i = 0; i < entries.length; i += chunkSize) {
    const chunk = entries.slice(i, i + chunkSize);
    const params = [];
    for (const entry of chunk) {
      params.push(
        crypto.randomUUID(),
        String(entry.entityName),
        String(entry.entityId),
        String(entry.action),
        JSON.stringify(entry.payload ?? {}),
        entry.correlationId || null
      );
    }

    if (requireChanges) {
      const selectArms = chunk.map((_, idx) =>
        idx === 0
          ? 'SELECT ? AS id, ? AS entity_name, ? AS entity_id, ? AS action, ? AS payload, ? AS correlation_id'
          : 'SELECT ?, ?, ?, ?, ?, ?'
      ).join('\n      UNION ALL\n      ');

      stmts.push(
        env.DB.prepare(`
          INSERT INTO sync_outbox (id, entity_name, entity_id, action, payload, correlation_id)
          SELECT v.id, v.entity_name, v.entity_id, v.action, v.payload, v.correlation_id
          FROM (
            ${selectArms}
          ) AS v
          WHERE (SELECT changes()) > 0
        `).bind(...params)
      );
    } else {
      const valuesPlaceholders = chunk.map(() => '(?, ?, ?, ?, ?, ?)').join(', ');
      stmts.push(
        env.DB.prepare(`
          INSERT INTO sync_outbox (id, entity_name, entity_id, action, payload, correlation_id)
          VALUES ${valuesPlaceholders}
        `).bind(...params)
      );
    }
  }
  return stmts;
}

/**
 * Generic handler for incoming webhook data
 */
export async function handleGenericWebhook(env, sheetName, action, excelData, payloadVersion, correlationId, eventId, startTime = Date.now(), options = {}) {
  const syncState = resolveSyncMode(env);
  const isDryRun = options.dryRun === true || !syncState.isLive;
  let mapping;
  try {
    mapping = mapPayloadToDB(sheetName, excelData, payloadVersion);
  } catch (err) {
    throw new WebhookValidationError(err.message, 400);
  }

  const { table, soft_delete_col, data } = mapping;
  const webhookPrincipal = {
    id: null,
    username: 'google_sheets_webhook',
    full_name: 'Google Sheets Integration',
    role: 'system'
  };

  const duration = Date.now() - startTime;
  const metricsStmt = env.DB.prepare(
    `INSERT INTO sync_metrics (event_id, action, module, webhook_duration_ms, correlation_id) VALUES (?, ?, ?, ?, ?)`
  ).bind(eventId, action, sheetName, duration, correlationId || null);

  // ==========================
  // ACTION: INSERT
  // ==========================
  if (action === 'INSERT') {
    const insertData = { ...data };
    if (insertData.id === '' || insertData.id === null || insertData.id === undefined) {
      delete insertData.id;
    }
    if (insertData.row_version === undefined || insertData.row_version === null || insertData.row_version === '') {
      insertData.row_version = 1;
    } else {
      insertData.row_version = Number(insertData.row_version) || 1;
    }
    insertData.last_sync_source = 'SHEET';

    // Phase 5.66.10A.1 / 5.66.19: Employee & Contract Management scope check
    if (table === 'employees' || table === 'contracts') {
      if (table === 'contracts' && insertData.employee_id) {
        const emp = await env.DB.prepare(
          `SELECT e.branch_id, e.division FROM employees e
           JOIN branches b ON e.branch_id = b.id
           WHERE e.id = ? AND e.deleted_at IS NULL
             AND e.division IN ('FACILITY CARE', 'SECURITY')
             AND ${getManagementBranchSqlPredicate('b')}`
        ).bind(insertData.employee_id).first();
        if (!emp) {
          return { skipped: true, reason: 'OUTSIDE_MANAGEMENT_SCOPE' };
        }
        if (!insertData.branch_id) insertData.branch_id = emp.branch_id;
        if (!insertData.division) insertData.division = emp.division;
      }
      if (insertData.branch_id) {
        const mgmtBranch = await env.DB.prepare(
          `SELECT id FROM branches WHERE id = ? AND deleted_at IS NULL AND ${getManagementBranchSqlPredicate('')}`
        ).bind(insertData.branch_id).first();
        if (!mgmtBranch) {
          return { skipped: true, reason: 'OUTSIDE_MANAGEMENT_SCOPE' };
        }
      } else {
        return { skipped: true, reason: 'OUTSIDE_MANAGEMENT_SCOPE' };
      }
      if (insertData.division) {
        const divUpper = String(insertData.division).trim().toUpperCase();
        if (divUpper !== 'FACILITY CARE' && divUpper !== 'SECURITY') {
          return { skipped: true, reason: 'OUTSIDE_MANAGEMENT_SCOPE' };
        }
        insertData.division = divUpper;
      } else if (table === 'contracts') {
        return { skipped: true, reason: 'OUTSIDE_MANAGEMENT_SCOPE' };
      }
    }

    // Supply safe table-specific defaults for required NOT NULL columns if omitted from sheet mapping
    if (table === 'employees') {
      if (!insertData.full_name || String(insertData.full_name).trim() === '') {
        throw new WebhookValidationError("Missing required field: 'Nama Lengkap' (full_name)", 400);
      }
      if (!insertData.division) insertData.division = 'FACILITY CARE';
      if (!insertData.status) insertData.status = 'Aktif';
    } else if (table === 'branches') {
      if (!insertData.full_name) insertData.full_name = insertData.name || insertData.code || 'Cabang';
    } else if (table === 'sop') {
      if (!insertData.name) insertData.name = insertData.title || 'SOP';
    } else if (table === 'master_checklist') {
      if (!insertData.name) insertData.name = 'Checklist';
    } else if (table === 'master_forms') {
      if (!insertData.name) insertData.name = 'Form';
    } else if (table === 'activity_schedule') {
      if (!insertData.activity_type) insertData.activity_type = 'General';
    } else if (table === 'relievers') {
      if (!insertData.reliever_name) insertData.reliever_name = 'Reliever';
      if (!insertData.backup_date) insertData.backup_date = new Date().toISOString().slice(0, 10);
    } else if (table === 'training') {
      if (!insertData.training_date) insertData.training_date = new Date().toISOString().slice(0, 10);
    } else if (table === 'one_on_one') {
      if (!insertData.meeting_date) insertData.meeting_date = new Date().toISOString().slice(0, 10);
    } else if (table === 'issues') {
      if (!insertData.category) insertData.category = 'General';
      if (!insertData.report_date) insertData.report_date = new Date().toISOString().slice(0, 10);
    } else if (table === 'cleaning_reports') {
      if (!insertData.activity_date) insertData.activity_date = new Date().toISOString().slice(0, 10);
    } else if (table === 'inspection_reports') {
      if (!insertData.inspection_date) insertData.inspection_date = new Date().toISOString().slice(0, 10);
    }

    // Whitelist column projection safeguard: prune any key not in table's canonical columns
    if (table === 'employees') {
      const allowedCols = new Set([
        'id', 'full_name', 'branch_id', 'division', 'phone', 'join_date',
        'status', 'notes', 'target_pindah_os', 'target_selesai_os',
        'row_version', 'last_sync_source', 'deleted_at', 'created_at', 'updated_at'
      ]);
      for (const k of Object.keys(insertData)) {
        if (!allowedCols.has(k)) delete insertData[k];
      }
    }

    const keys = Object.keys(insertData).join(', ');
    const placeholders = Object.keys(insertData).map(() => '?').join(', ');
    const values = Object.values(insertData);

    if (isDryRun) {
      console.log(`[DRY_RUN] INSERT INTO ${table} (${keys}) VALUES (${values})`);
      return { dryRun: true };
    }

    const mutationStmt = env.DB.prepare(
      `INSERT INTO ${table} (${keys}) VALUES (${placeholders})`
    ).bind(...values);

    const idempotencyStmt = env.DB.prepare(
      `INSERT INTO sync_idempotency (event_id) SELECT ? WHERE (SELECT changes()) > 0`
    ).bind(eventId);

    const auditStmt = buildAuditStatement(
      env,
      webhookPrincipal,
      'CREATE',
      sheetName,
      String(insertData.id || eventId),
      null,
      insertData,
      { status: 'SUCCESS', source: 'GOOGLE_SHEETS_WEBHOOK', event_id: eventId, correlation_id: correlationId || null },
      'SUCCESS',
      true
    );

    // Atomically execute: Mutation -> Idempotency (gated on changes()) -> Audit Log (gated on changes())
    await env.DB.batch([mutationStmt, idempotencyStmt, auditStmt]);

    // Observability metrics (non-fatal)
    try {
      await metricsStmt.run();
    } catch {
      // Non-fatal if sync_metrics write fails
    }

    return { inserted: true };
  }

  // ==========================
  // ACTION: UPDATE
  // ==========================
  if (action === 'UPDATE') {
    const id = data['id'];
    if (id === undefined || id === null || id === '') {
      throw new WebhookValidationError('Missing id for UPDATE action', 400);
    }
    delete data['id'];

    // OCC Check: row_version MUST be provided and valid
    const incomingVersionRaw = data['row_version'];
    if (incomingVersionRaw === undefined || incomingVersionRaw === null || incomingVersionRaw === '') {
      throw new WebhookValidationError('Missing row_version for UPDATE: OCC requires expected row_version', 400);
    }
    const incomingVersion = Number(incomingVersionRaw);
    if (!Number.isInteger(incomingVersion) || incomingVersion < 0) {
      throw new WebhookValidationError('Invalid row_version: must be a non-negative integer', 400);
    }

    // Verify existing record in DB
    const existing = await env.DB.prepare(`SELECT * FROM ${table} WHERE id = ?`).bind(id).first();
    if (!existing) {
      throw new WebhookValidationError(`Record not found in ${table} with id ${id}`, 404);
    }

    // Phase 5.66.10A.1: Employee & Contract Management scope check
    if (table === 'employees' || table === 'contracts') {
      let isMgmtExisting;
      if (table === 'employees') {
        isMgmtExisting = await env.DB.prepare(
          `SELECT c.id FROM employees c
           JOIN branches b ON c.branch_id = b.id
           WHERE c.id = ? AND c.deleted_at IS NULL
             AND c.division IN ('FACILITY CARE', 'SECURITY')
             AND ${getManagementBranchSqlPredicate('b')}`
        ).bind(id).first();
      } else {
        isMgmtExisting = await env.DB.prepare(
          `SELECT c.id FROM contracts c
           JOIN branches b ON c.branch_id = b.id
           LEFT JOIN employees e ON c.employee_id = e.id
           LEFT JOIN branches eb ON e.branch_id = eb.id
           WHERE c.id = ? AND c.deleted_at IS NULL
             AND c.division IN ('FACILITY CARE', 'SECURITY')
             AND ${getManagementBranchSqlPredicate('b')}
             AND (c.employee_id IS NULL OR (e.deleted_at IS NULL AND eb.id IS NOT NULL AND ${getManagementBranchSqlPredicate('eb')} AND e.division IN ('FACILITY CARE', 'SECURITY')))`
        ).bind(id).first();
      }
      if (!isMgmtExisting) {
        return { skipped: true, reason: 'OUTSIDE_MANAGEMENT_SCOPE' };
      }
      if (data.branch_id) {
        const mgmtBranch = await env.DB.prepare(
          `SELECT id FROM branches WHERE id = ? AND deleted_at IS NULL AND ${getManagementBranchSqlPredicate('')}`
        ).bind(data.branch_id).first();
        if (!mgmtBranch) {
          return { skipped: true, reason: 'OUTSIDE_MANAGEMENT_SCOPE' };
        }
      }
      if (data.division) {
        const divUpper = String(data.division).trim().toUpperCase();
        if (divUpper !== 'FACILITY CARE' && divUpper !== 'SECURITY') {
          return { skipped: true, reason: 'OUTSIDE_MANAGEMENT_SCOPE' };
        }
      }
    }

    const currentVersion = Number(existing.row_version ?? 0);
    if (incomingVersion !== currentVersion) {
      const conflictMsg = `OCC Conflict: Stale row_version (incoming: ${incomingVersion}, current in DB: ${currentVersion})`;
      if (isDryRun) {
        console.log(`[DRY_RUN] CONFLICT for ${sheetName} ID ${id}: ${conflictMsg}`);
        return { conflict: true, dryRun: true };
      }
      throw new WebhookConflictError(conflictMsg, 409, {
        current_version: currentVersion,
        incoming_version: incomingVersion
      });
    }

    const updateData = { ...data };
    updateData.last_sync_source = 'SHEET';
    delete updateData.row_version; // Managed via SQL increment

    // Validation & whitelist column projection safeguard for UPDATE
    if (table === 'employees') {
      if (updateData.full_name !== undefined && (updateData.full_name === null || String(updateData.full_name).trim() === '')) {
        throw new WebhookValidationError("Field 'Nama Lengkap' (full_name) cannot be empty", 400);
      }
      const allowedCols = new Set([
        'full_name', 'branch_id', 'division', 'phone', 'join_date',
        'status', 'notes', 'target_pindah_os', 'target_selesai_os',
        'last_sync_source', 'deleted_at', 'updated_at'
      ]);
      for (const k of Object.keys(updateData)) {
        if (!allowedCols.has(k)) delete updateData[k];
      }
    }

    const sets = Object.keys(updateData).map(k => `${k} = ?`).join(', ');
    const values = Object.values(updateData);
    values.push(id, incomingVersion);

    if (isDryRun) {
      console.log(`[DRY_RUN] UPDATE ${table} SET ${sets}, row_version = row_version + 1 WHERE id = ${id} AND COALESCE(row_version, 0) = ${incomingVersion}`);
      return { dryRun: true };
    }

    const mutationStmt = env.DB.prepare(
      `UPDATE ${table} SET ${sets}, row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'SHEET', updated_at = datetime('now') WHERE id = ? AND COALESCE(row_version, 0) = ?`
    ).bind(...values);

    const idempotencyStmt = env.DB.prepare(
      `INSERT INTO sync_idempotency (event_id) SELECT ? WHERE (SELECT changes()) > 0`
    ).bind(eventId);

    const auditStmt = buildAuditStatement(
      env,
      webhookPrincipal,
      'UPDATE',
      sheetName,
      String(id),
      existing,
      { ...updateData, row_version: currentVersion + 1 },
      { status: 'SUCCESS', source: 'GOOGLE_SHEETS_WEBHOOK', event_id: eventId, correlation_id: correlationId || null, previous_version: currentVersion, new_version: currentVersion + 1 },
      'SUCCESS',
      true
    );

    const batchRes = await env.DB.batch([mutationStmt, idempotencyStmt, auditStmt]);
    const changes = Number(batchRes?.[0]?.meta?.changes ?? batchRes?.[0]?.changes ?? 0);
    if (changes === 0) {
      const alreadyCommitted = await env.DB.prepare('SELECT 1 FROM sync_idempotency WHERE event_id = ?').bind(eventId).first();
      if (alreadyCommitted) {
        return { duplicate: true, message: 'Event already processed' };
      }
      throw new WebhookConflictError('OCC Conflict: Concurrent update modified row_version before commit', 409, {
        current_version: currentVersion,
        incoming_version: incomingVersion
      });
    }

    try {
      await metricsStmt.run();
    } catch {
      // Non-fatal
    }

    return { updated: true, new_row_version: currentVersion + 1 };
  }

  // ==========================
  // ACTION: DELETE (SOFT DELETE ONLY)
  // ==========================
  if (action === 'DELETE') {
    const id = data['id'];
    if (id === undefined || id === null || id === '') {
      throw new WebhookValidationError('Missing id for DELETE action', 400);
    }

    // OCC Check: row_version MUST be provided and valid on SOFT DELETE
    const incomingVersionRaw = data['row_version'];
    if (incomingVersionRaw === undefined || incomingVersionRaw === null || incomingVersionRaw === '') {
      throw new WebhookValidationError('Missing row_version for DELETE: OCC requires expected row_version', 400);
    }
    const incomingVersion = Number(incomingVersionRaw);
    if (!Number.isInteger(incomingVersion) || incomingVersion < 0) {
      throw new WebhookValidationError('Invalid row_version: must be a non-negative integer', 400);
    }

    const existing = await env.DB.prepare(`SELECT * FROM ${table} WHERE id = ?`).bind(id).first();
    if (!existing) {
      if (isDryRun) {
        console.log(`[DRY_RUN] Record ${id} not found in ${table}`);
        return { dryRun: true, notFound: true };
      }
      await env.DB.prepare(`INSERT OR IGNORE INTO sync_idempotency (event_id) VALUES (?)`).bind(eventId).run();
      return { notFound: true, message: 'Record not found; marked idempotency' };
    }

    const delCol = soft_delete_col || 'deleted_at';
    if (existing[delCol]) {
      if (isDryRun) {
        console.log(`[DRY_RUN] Record ${id} already soft-deleted in ${table}`);
        return { dryRun: true, alreadyDeleted: true };
      }
      await env.DB.prepare(`INSERT OR IGNORE INTO sync_idempotency (event_id) VALUES (?)`).bind(eventId).run();
      return { alreadyDeleted: true, message: 'Record already soft-deleted; marked idempotency' };
    }

    // Phase 5.66.19: Enforce Management scope guard on DELETE for employees & contracts
    if (table === 'employees' || table === 'contracts') {
      let isMgmtExisting;
      if (table === 'employees') {
        isMgmtExisting = await env.DB.prepare(
          `SELECT c.id FROM employees c
           JOIN branches b ON c.branch_id = b.id
           WHERE c.id = ? AND c.deleted_at IS NULL
             AND c.division IN ('FACILITY CARE', 'SECURITY')
             AND ${getManagementBranchSqlPredicate('b')}`
        ).bind(id).first();
      } else {
        isMgmtExisting = await env.DB.prepare(
          `SELECT c.id FROM contracts c
           JOIN branches b ON c.branch_id = b.id
           LEFT JOIN employees e ON c.employee_id = e.id
           LEFT JOIN branches eb ON e.branch_id = eb.id
           WHERE c.id = ? AND c.deleted_at IS NULL
             AND c.division IN ('FACILITY CARE', 'SECURITY')
             AND ${getManagementBranchSqlPredicate('b')}
             AND (c.employee_id IS NULL OR (e.deleted_at IS NULL AND eb.id IS NOT NULL AND ${getManagementBranchSqlPredicate('eb')} AND e.division IN ('FACILITY CARE', 'SECURITY')))`
        ).bind(id).first();
      }
      if (!isMgmtExisting) {
        if (!isDryRun) {
          await env.DB.prepare(`INSERT OR IGNORE INTO sync_idempotency (event_id) VALUES (?)`).bind(eventId).run();
        }
        return { skipped: true, reason: 'OUTSIDE_MANAGEMENT_SCOPE' };
      }
    }

    const currentVersion = Number(existing.row_version ?? 0);
    if (incomingVersion !== currentVersion) {
      const conflictMsg = `OCC Conflict on DELETE: Stale row_version (incoming: ${incomingVersion}, current: ${currentVersion})`;
      if (isDryRun) {
        console.log(`[DRY_RUN] CONFLICT on DELETE for ${sheetName} ID ${id}: ${conflictMsg}`);
        return { conflict: true, dryRun: true };
      }
      throw new WebhookConflictError(conflictMsg, 409, {
        current_version: currentVersion,
        incoming_version: incomingVersion
      });
    }

    if (isDryRun) {
      console.log(`[DRY_RUN] SOFT DELETE ${table} ID ${id} (${delCol})`);
      return { dryRun: true };
    }

    const mutationStmt = env.DB.prepare(
      `UPDATE ${table} SET ${delCol} = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'SHEET', updated_at = datetime('now') WHERE id = ? AND ${delCol} IS NULL AND COALESCE(row_version, 0) = ?`
    ).bind(id, incomingVersion);

    const idempotencyStmt = env.DB.prepare(
      `INSERT INTO sync_idempotency (event_id) SELECT ? WHERE (SELECT changes()) > 0`
    ).bind(eventId);

    const auditStmt = buildAuditStatement(
      env,
      webhookPrincipal,
      'DELETE',
      sheetName,
      String(id),
      existing,
      { [delCol]: 'datetime(\'now\')', row_version: currentVersion + 1, last_sync_source: 'SHEET' },
      { status: 'SUCCESS', source: 'GOOGLE_SHEETS_WEBHOOK', strategy: 'SOFT', event_id: eventId, correlation_id: correlationId || null, previous_version: currentVersion, new_version: currentVersion + 1 },
      'SUCCESS',
      true
    );

    const batchRes = await env.DB.batch([mutationStmt, idempotencyStmt, auditStmt]);
    const changes = Number(batchRes?.[0]?.meta?.changes ?? batchRes?.[0]?.changes ?? 0);
    if (changes === 0) {
      const alreadyCommitted = await env.DB.prepare('SELECT 1 FROM sync_idempotency WHERE event_id = ?').bind(eventId).first();
      if (alreadyCommitted) {
        return { duplicate: true, message: 'Event already processed' };
      }
      throw new WebhookConflictError('OCC Conflict on DELETE: Concurrent modification changed row_version or deleted state before commit', 409, {
        current_version: currentVersion,
        incoming_version: incomingVersion
      });
    }

    try {
      await metricsStmt.run();
    } catch {
      // Non-fatal
    }

    return { deleted: true, new_row_version: currentVersion + 1 };
  }

  throw new WebhookValidationError(`Unsupported action: ${action}`, 400);
}
