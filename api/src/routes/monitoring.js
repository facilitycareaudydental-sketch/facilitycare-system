import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden } from '../utils/response.js';
import { SYNC_MAPPER } from '../utils/sync_mapper.js';
import { runReconciliation, createNightlySnapshot } from '../utils/reconciliation.js';
import { sanitizeAuditData, buildAuditStatement } from '../utils/audit.js';
import { getOutboxStats } from '../utils/outbox_processor.js';
import { resolveSyncMode } from '../utils/sync_mode.js';

export async function handleMonitoring(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname;
  
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (user.role !== 'superadmin' || !hasPermission(user, 'monitoring', 'admin')) return forbidden(origin);

  if (request.method === 'GET' && path === '/api/sync/health') return getHealth(env, origin);
  if (request.method === 'GET' && path === '/api/sync/status') return getStatus(env, origin);
  if (request.method === 'GET' && path === '/api/sync/queue') return getQueue(request, env, origin);
  if (request.method === 'GET' && path === '/api/sync/metrics') return getMetrics(env, origin);
  if (request.method === 'GET' && path === '/api/sync/performance') return getPerformance(env, origin);
  if (request.method === 'GET' && path === '/api/sync/errors') return getErrors(request, env, origin);

  // Admin Actions
  if (request.method === 'POST' && path.startsWith('/api/sync/actions/')) {
    return handleAdminActions(request, env, origin, user);
  }

  return error('Not found', 404, origin);
}

async function handleAdminActions(request, env, origin, user) {
  const url = new URL(request.url);
  const action = url.pathname.split('/').pop();
  try {
    const payload = await request.json().catch(() => ({}));

    if (action === 'retry') {
      const { ids, allFailed } = payload || {};
      if (allFailed === true) {
        const retryAllStmt = env.DB.prepare(`UPDATE sync_outbox SET status = 'PENDING', retry_count = 0, next_retry_at = NULL, claim_token = NULL, lease_expires_at = NULL, updated_at = datetime('now') WHERE status IN ('RETRY', 'FAILED', 'DEAD_LETTER')`);
        const auditStmt = buildAuditStatement(
          env,
          user,
          'UPDATE',
          'sync_outbox',
          'ALL_FAILED',
          null,
          { status: 'PENDING', retry_count: 0 },
          { operation: 'MANUAL_OUTBOX_RETRY', scope: 'allFailed' },
          'SUCCESS',
          true
        );
        await env.DB.batch([retryAllStmt, auditStmt]);
        return ok({ message: 'All failed/dead-letter events queued for retry' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
      }
      if (Array.isArray(ids) && ids.length > 0) {
        if (ids.length > 100) {
          return error('Maksimal 100 ID dalam satu permintaan retry', 400, origin);
        }
        const validIds = ids
          .map(id => {
            if (typeof id === 'number' && Number.isInteger(id) && id > 0) return String(id);
            if (typeof id === 'string') {
              const trimmed = id.trim();
              if (/^[1-9]\d*$/.test(trimmed)) return trimmed;
              if (/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(trimmed)) return trimmed;
            }
            return null;
          })
          .filter(Boolean);
        if (validIds.length === 0 || validIds.length !== ids.length) {
          return error('Invalid payload: ids harus berupa bilangan bulat positif atau UUID valid', 400, origin);
        }
        const placeholders = validIds.map(() => '?').join(',');
        const retryStmt = env.DB.prepare(`UPDATE sync_outbox SET status = 'PENDING', retry_count = 0, next_retry_at = NULL, claim_token = NULL, lease_expires_at = NULL, updated_at = datetime('now') WHERE id IN (${placeholders}) AND status IN ('RETRY', 'FAILED', 'DEAD_LETTER')`).bind(...validIds);
        const auditStmt = buildAuditStatement(
          env,
          user,
          'UPDATE',
          'sync_outbox',
          validIds.slice(0, 10).join(','),
          null,
          { status: 'PENDING', retry_count: 0 },
          { operation: 'MANUAL_OUTBOX_RETRY', count: validIds.length },
          'SUCCESS',
          true
        );
        await env.DB.batch([retryStmt, auditStmt]);
        return ok({ message: `Queued ${validIds.length} events for retry` }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
      }
      return error('Invalid payload', 400, origin);
    }

    if (action === 'pause') {
      const prev = await env.DB.prepare(`SELECT * FROM circuit_breaker WHERE id = 1`).first();
      await env.DB.batch([
        env.DB.prepare(`UPDATE circuit_breaker SET state = 'OPEN', opened_at = datetime('now'), next_attempt_at = datetime('now', '+100 years') WHERE id = 1`),
        buildAuditStatement(env, user, 'UPDATE', 'circuit_breaker', '1', prev || null, { state: 'OPEN' }, { operation: 'SYNC_PAUSE' }, 'SUCCESS', true)
      ]);
      return ok({ message: 'Sync Paused (Circuit Breaker Force OPEN)' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
    }

    if (action === 'resume') {
      const prev = await env.DB.prepare(`SELECT * FROM circuit_breaker WHERE id = 1`).first();
      await env.DB.batch([
        env.DB.prepare(`UPDATE circuit_breaker SET state = 'CLOSED', failure_count = 0 WHERE id = 1`),
        buildAuditStatement(env, user, 'UPDATE', 'circuit_breaker', '1', prev || null, { state: 'CLOSED', failure_count: 0 }, { operation: 'SYNC_RESUME' }, 'SUCCESS', true)
      ]);
      return ok({ message: 'Sync Resumed (Circuit Breaker CLOSED)' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
    }

    if (action === 'reset-stuck') {
      const rawTimeout = payload?.stuckTimeoutMinutes ?? env.SYNC_STUCK_TIMEOUT ?? 5;
      const stuckTimeoutMinutes = Math.min(Math.max(parseInt(String(rawTimeout), 10) || 5, 1), 1440);
      const intervalParam = `-${stuckTimeoutMinutes} minutes`;
      const batchRes = await env.DB.batch([
        env.DB.prepare(
          `UPDATE sync_outbox SET status = 'PENDING', claim_token = NULL, lease_expires_at = NULL, updated_at = datetime('now') WHERE status = 'PROCESSING' AND updated_at < datetime('now', ?) AND (lease_expires_at IS NULL OR lease_expires_at <= datetime('now'))`
        ).bind(intervalParam),
        buildAuditStatement(env, user, 'UPDATE', 'sync_outbox', 'stuck', null, { status: 'PENDING', stuck_timeout_minutes: stuckTimeoutMinutes }, { operation: 'RESET_STUCK_OUTBOX' }, 'SUCCESS', true)
      ]);
      const changes = batchRes?.[0]?.meta?.changes ?? batchRes?.[0]?.changes ?? 0;
      return ok({ message: `Reset stuck events.`, changes }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
    }

    if (action === 'reconcile') {
      const rawModule = payload?.module ? String(payload.module).trim() : 'ALL';
      const validModules = new Set(['ALL', ...Object.keys(SYNC_MAPPER || {})]);
      if (!validModules.has(rawModule)) {
        return error(`Invalid reconciliation module '${rawModule}'`, 400, origin);
      }
      const isRepairMode = payload?.repairMode === true;
      const res = await runReconciliation(env, rawModule, isRepairMode);
      return ok(res, 200, origin);
    }
    
    if (action === 'snapshot') {
      const res = await createNightlySnapshot(env);
      return ok(res, 200, origin);
    }

    return error('Invalid action', 400, origin);
  } catch (err) {
    console.error('handleAdminActions error:', err);
    return error('Action failed', 500, origin);
  }
}

async function getHealth(env, origin) {
  try {
    const queueCount = await env.DB.prepare("SELECT COUNT(*) as count FROM sync_outbox WHERE status = 'PENDING'").first();
    const retryCount = await env.DB.prepare("SELECT COUNT(*) as count FROM sync_outbox WHERE status = 'RETRY'").first();
    const deadLetters = await env.DB.prepare("SELECT COUNT(*) as count FROM sync_outbox WHERE status = 'DEAD_LETTER'").first();
    const failures = await env.DB.prepare("SELECT COUNT(*) as count FROM sync_outbox WHERE status = 'FAILED'").first();
    const processing = await env.DB.prepare("SELECT COUNT(*) as count FROM sync_outbox WHERE status = 'PROCESSING'").first();
    
    // Overall Health Formula
    // If dead letters > 10, health is critical.
    let status = 'HEALTHY';
    if (deadLetters.count > 0 || failures.count > 10) status = 'WARNING';
    if (deadLetters.count > 50 || processing.count > 100) status = 'CRITICAL';
    
    let cbState = 'CLOSED';
    try {
      const cb = await env.DB.prepare('SELECT state FROM circuit_breaker WHERE id = 1').first();
      if (cb && cb.state) cbState = cb.state;
    } catch (cbErr) {
      if (!/no such table:\s*circuit_breaker/i.test(String(cbErr?.message || cbErr || ''))) {
        throw cbErr;
      }
    }
    const syncState = resolveSyncMode(env);
    
    return ok({
      status,
      timestamp: new Date().toISOString(),
      queue: {
        pending: queueCount.count,
        retry: retryCount?.count ?? 0,
        processing: processing.count,
        failed: failures.count,
        dead_letter: deadLetters.count
      },
      circuit_breaker: cbState,
      env_mode: syncState.mode,
      sync_mode_valid: syncState.valid,
      ...(syncState.reason ? { sync_mode_reason: syncState.reason } : {})
    }, 200, origin);
  } catch (err) {
    console.error('getHealth error:', err);
    return error('Health check failed', 500, origin);
  }
}

async function getStatus(env, origin) {
  // Returns the current configuration from SYNC_MAPPER and feature flags
  try {
    const flagMap = {};
    try {
      const flags = await env.DB.prepare("SELECT * FROM feature_flags").all();
      for (const f of (flags?.results || [])) {
        flagMap[f.flag_key] = f.flag_value === 'true';
      }
    } catch (flagErr) {
      if (!/no such table:\s*feature_flags/i.test(String(flagErr?.message || flagErr || ''))) {
        throw flagErr;
      }
    }
    
    const modules = {};
    for (const [key, val] of Object.entries(SYNC_MAPPER)) {
      modules[key] = {
        enabled: val.enabled,
        direction: val.direction,
        table: val.table,
        delete_strategy: val.delete_strategy,
        schema_version: val.schema_version
      };
    }
    
    return ok({
      feature_flags: flagMap,
      modules
    }, 200, origin);
  } catch (err) {
    console.error('getStatus error:', err);
    return error('Failed to fetch status', 500, origin);
  }
}

async function getQueue(request, env, origin) {
  try {
    const url = new URL(request.url);
    const limit = Math.min(Math.max(parseInt(url.searchParams.get('limit') || '50', 10) || 50, 1), 100);
    const offset = Math.max(parseInt(url.searchParams.get('offset') || '0', 10) || 0, 0);
    const rawStatus = (url.searchParams.get('status') || '').trim().toUpperCase();
    const VALID_STATUSES = ['PENDING', 'PROCESSING', 'DELIVERED', 'COMPLETED', 'RETRY', 'FAILED', 'DEAD_LETTER', 'QUARANTINED_TEST', 'DISCARDED_STAGING_TEST'];
    const status = VALID_STATUSES.includes(rawStatus) ? rawStatus : null;
    
    let query = "SELECT id, entity_name, entity_id, action, status, retry_count, last_error, created_at, next_retry_at FROM sync_outbox";
    let countQuery = "SELECT COUNT(*) as total FROM sync_outbox";
    
    let rows;
    let count;

    if (status) {
      query += " WHERE status = ? ORDER BY created_at DESC, rowid DESC LIMIT ? OFFSET ?";
      countQuery += " WHERE status = ?";
      rows = await env.DB.prepare(query).bind(status, limit, offset).all();
      count = await env.DB.prepare(countQuery).bind(status).first();
    } else {
      query += " ORDER BY created_at DESC, rowid DESC LIMIT ? OFFSET ?";
      rows = await env.DB.prepare(query).bind(limit, offset).all();
      count = await env.DB.prepare(countQuery).first();
    }
    
    return ok({
      data: rows ? rows.results : [],
      meta: {
        total: count ? count.total : 0,
        limit,
        offset
      }
    }, 200, origin);
  } catch (err) {
    console.error('getQueue error:', err);
    return error('Failed to fetch queue', 500, origin);
  }
}

async function getMetrics(env, origin) {
  try {
    // Get aggregate metrics from the last 24 hours
    let results = [];
    try {
      const stats = await env.DB.prepare(`
        SELECT 
          module,
          COUNT(*) as total_events,
          AVG(webhook_duration_ms) as avg_webhook_ms,
          AVG(d1_execution_ms) as avg_d1_ms,
          AVG(queue_wait_time_ms) as avg_queue_wait_ms
        FROM sync_metrics
        WHERE created_at >= datetime('now', '-1 day')
        GROUP BY module
      `).all();
      results = stats?.results || [];
    } catch (mErr) {
      if (!/no such table:\s*sync_metrics/i.test(String(mErr?.message || mErr || ''))) {
        throw mErr;
      }
    }
    
    return ok(results, 200, origin);
  } catch (err) {
    console.error('getMetrics error:', err);
    return error('Failed to fetch metrics', 500, origin);
  }
}

async function getErrors(request, env, origin) {
  const req = (request && request.url) ? request : null;
  const actualEnv = req ? env : request;
  const actualOrigin = req ? origin : env;

  try {
    const url = req ? new URL(req.url) : null;
    const limit = Math.min(Math.max(parseInt(url?.searchParams.get('limit') || '50', 10) || 50, 1), 500);
    
    let errors;
    try {
      errors = await actualEnv.DB.prepare(`
        SELECT id, user_id, user_name, user_role, action, module, target_id, old_data, new_data, details, ip_address, created_at
        FROM audit_logs 
        WHERE details LIKE '%"status":"FAILED"%' OR details LIKE '%DEAD_LETTER%' OR details LIKE '%CONFLICT%' OR details LIKE '%SYNC_BLOCKED%' 
        ORDER BY created_at DESC 
        LIMIT ?
      `).bind(limit).all();
    } catch (colErr) {
      const colErrMsg = String(colErr?.message || colErr || '');
      if (!colErrMsg.includes('no such column')) {
        throw colErr;
      }
      errors = await actualEnv.DB.prepare(`
        SELECT id, user_id, user_name, user_role, action, module, target_id, old_value as old_data, new_value as new_data, details, ip_address, created_at
        FROM audit_logs 
        WHERE details LIKE '%"status":"FAILED"%' OR details LIKE '%DEAD_LETTER%' OR details LIKE '%CONFLICT%' OR details LIKE '%SYNC_BLOCKED%' 
        ORDER BY created_at DESC 
        LIMIT ?
      `).bind(limit).all();
    }

    const sanitizedRows = (errors?.results || []).map(row => ({
      id: row.id,
      user_id: row.user_id,
      user_name: row.user_name,
      user_role: row.user_role,
      action: row.action,
      module: row.module,
      target_id: row.target_id,
      old_data: sanitizeAuditData(row.old_data ?? row.old_value ?? null),
      new_data: sanitizeAuditData(row.new_data ?? row.new_value ?? null),
      details: sanitizeAuditData(row.details ?? null),
      ip_address: row.ip_address ?? null,
      created_at: row.created_at
    }));
    
    return ok(sanitizedRows, 200, actualOrigin);
  } catch (err) {
    console.error('getErrors error:', err);
    return error('Failed to fetch errors', 500, actualOrigin);
  }
}

async function getPerformance(env, origin) {
  try {
    // Pull recent metrics and compute percentiles in memory
    let metrics = [];
    try {
      const res = await env.DB.prepare(`
        SELECT webhook_duration_ms, d1_execution_ms, google_api_execution_ms, queue_wait_time_ms
        FROM sync_metrics
        WHERE created_at > datetime('now', '-24 hours')
        ORDER BY created_at DESC
        LIMIT 1000
      `).all();
      metrics = res?.results || [];
    } catch (pErr) {
      if (!/no such table:\s*sync_metrics/i.test(String(pErr?.message || pErr || ''))) {
        throw pErr;
      }
    }
    
    if (metrics.length === 0) {
       return ok({ message: "No data", status: "DATA_UNAVAILABLE" }, 200, origin);
    }

    const calcHist = (arr) => {
      arr.sort((a,b) => a-b);
      const l = arr.length;
      return {
        P50: arr[Math.floor(l * 0.50)] || 0,
        P90: arr[Math.floor(l * 0.90)] || 0,
        P95: arr[Math.floor(l * 0.95)] || 0,
        P99: arr[Math.floor(l * 0.99)] || 0,
        Max: arr[l - 1] || 0,
        Min: arr[0] || 0
      };
    };

    const webhookHist = calcHist(metrics.map(m => m.webhook_duration_ms || 0));
    const d1Hist = calcHist(metrics.map(m => m.d1_execution_ms || 0));
    const googleHist = calcHist(metrics.map(m => m.google_api_execution_ms || 0));
    const queueHist = calcHist(metrics.map(m => m.queue_wait_time_ms || 0));

    // Derive real rates from sync_outbox if available; never fabricate hardcoded percentages
    let aggregation = {
      success_rate: null,
      retry_rate: null,
      conflict_rate: null,
      status: 'DATA_UNAVAILABLE'
    };

    try {
      const deliveredRow = await env.DB.prepare("SELECT COUNT(*) as count FROM sync_outbox WHERE status = 'DELIVERED'").first();
      const completedRow = await env.DB.prepare("SELECT COUNT(*) as count FROM sync_outbox WHERE status = 'COMPLETED'").first();
      const failedRow = await env.DB.prepare("SELECT COUNT(*) as count FROM sync_outbox WHERE status IN ('FAILED', 'DEAD_LETTER')").first();
      const retriedRow = await env.DB.prepare("SELECT COUNT(*) as count FROM sync_outbox WHERE retry_count > 0").first();

      const completed = Number(deliveredRow?.count || 0) + Number(completedRow?.count || 0);
      const failed = Number(failedRow?.count || 0);
      const retried = Number(retriedRow?.count || 0);
      const totalProcessed = completed + failed;

      if (totalProcessed > 0) {
        aggregation = {
          success_rate: Number(((completed / totalProcessed) * 100).toFixed(2)),
          retry_rate: Number(((retried / totalProcessed) * 100).toFixed(2)),
          conflict_rate: null,
          status: 'CALCULATED_FROM_OUTBOX'
        };
      }
    } catch {
      // Leave aggregation as DATA_UNAVAILABLE if outbox counters cannot be queried
    }

    return ok({
      webhook: webhookHist,
      d1: d1Hist,
      google_api: googleHist,
      queue: queueHist,
      throughput: {
        events_per_sec: (metrics.length / 86400).toFixed(2),
        events_per_min: (metrics.length / 1440).toFixed(2)
      },
      aggregation
    }, 200, origin);
  } catch (err) {
    console.error('getPerformance error:', err);
    return error('Failed to fetch performance metrics', 500, origin);
  }
}
