/**
 * Transactional Outbox Processor (Concurrency Control, Delivery, Dead-Letter, Observability)
 *
 * Uses SQLite/D1 conditional atomic updates with unique claim_token + monotonically
 * increasing claim_version (CAS lease fencing) and dedicated lease_expires_at
 * separated from next_retry_at.
 */

import { evaluateRetryDecision, sanitizeOutboxError, toSqliteDatetime } from './outbox_retry.js';
import { createDryRunSimulationAdapter } from './outbox_adapter.js';
import { resolveSyncMode } from './sync_mode.js';
import { resolveMaintenanceFreezeState } from './maintenance.js';

export const DEFAULT_BATCH_LIMIT = 20;
export const MAX_BATCH_LIMIT = 50;
export const DEFAULT_LEASE_SECONDS = 60;

const BLOCKING_PREREQUISITE_STATUSES_SQL = "('PENDING', 'RETRY', 'PROCESSING', 'FAILED', 'DEAD_LETTER')";

function safeRowVersionSqlExpr(alias) {
  return `CASE WHEN json_valid(${alias}.payload) THEN CAST(COALESCE(json_extract(${alias}.payload, '$.ROW_VERSION'), 0) AS INTEGER) ELSE 0 END`;
}

function actionRankSqlExpr(alias) {
  return `CASE ${alias}.action WHEN 'INSERT' THEN 1 WHEN 'UPDATE' THEN 2 WHEN 'DELETE' THEN 3 ELSE 4 END`;
}

function buildOrderBySql(alias = 'sync_outbox') {
  return `${alias}.created_at ASC, ${actionRankSqlExpr(alias)} ASC, ${safeRowVersionSqlExpr(alias)} ASC, ${alias}.rowid ASC`;
}

function buildNoBlockingPrerequisiteSql(outerAlias = 'sync_outbox') {
  const prevAct = actionRankSqlExpr('prev');
  const curAct = actionRankSqlExpr(outerAlias);
  const prevVer = safeRowVersionSqlExpr('prev');
  const curVer = safeRowVersionSqlExpr(outerAlias);
  return `NOT EXISTS (
    SELECT 1 FROM sync_outbox prev
    WHERE prev.entity_name = ${outerAlias}.entity_name
      AND prev.entity_id = ${outerAlias}.entity_id
      AND prev.id != ${outerAlias}.id
      AND prev.status IN ${BLOCKING_PREREQUISITE_STATUSES_SQL}
      AND (
        prev.created_at < ${outerAlias}.created_at
        OR (
          prev.created_at = ${outerAlias}.created_at
          AND (
            ${prevAct} < ${curAct}
            OR (
              ${prevAct} = ${curAct}
              AND (
                ${prevVer} < ${curVer}
                OR (${prevVer} = ${curVer} AND prev.rowid < ${outerAlias}.rowid)
              )
            )
          )
        )
      )
  )`;
}

/**
 * Recovers stale PROCESSING events whose lease_expires_at has elapsed
 * (or legacy PROCESSING rows without lease_expires_at older than 5 minutes).
 */
export async function recoverStaleProcessingEvents(db, nowMs = Date.now()) {
  const nowSql = toSqliteDatetime(new Date(nowMs));
  const legacyCutoffSql = toSqliteDatetime(new Date(nowMs - 5 * 60 * 1000));
  const res = await db.prepare(`
    UPDATE sync_outbox
    SET status = 'RETRY',
        claim_token = NULL,
        lease_expires_at = NULL,
        next_retry_at = ?,
        last_error = COALESCE(last_error, '[LEASE_EXPIRED] Recovered stale PROCESSING event'),
        updated_at = ?
    WHERE status = 'PROCESSING'
      AND (
        (lease_expires_at IS NOT NULL AND lease_expires_at <= ?)
        OR (lease_expires_at IS NULL AND updated_at < ?)
      )
  `).bind(nowSql, nowSql, nowSql, legacyCutoffSql).run();

  return Number(res?.meta?.changes ?? res?.changes ?? 0);
}

async function isCircuitBreakerOpen(db, nowSql) {
  try {
    const cb = await db.prepare('SELECT state, next_attempt_at FROM circuit_breaker WHERE id = 1').first();
    if (!cb) return false;
    if (cb.state === 'OPEN') {
      if (!cb.next_attempt_at || cb.next_attempt_at > nowSql) {
        return true;
      }
    }
    return false;
  } catch {
    return false;
  }
}

async function openCircuitBreakerOnAuthFailure(db, nowMs) {
  try {
    const nowSql = toSqliteDatetime(new Date(nowMs));
    const nextAttemptSql = toSqliteDatetime(new Date(nowMs + 5 * 60 * 1000));
    await db.prepare(`
      UPDATE circuit_breaker
      SET state = 'OPEN',
          failure_count = COALESCE(failure_count, 0) + 1,
          opened_at = ?,
          next_attempt_at = ?
      WHERE id = 1
    `).bind(nowSql, nextAttemptSql).run();
  } catch {
    // Non-fatal if circuit_breaker table is not provisioned in minimal test DB
  }
}

async function closeCircuitBreakerOnSuccess(db) {
  try {
    await db.prepare(`
      UPDATE circuit_breaker
      SET state = 'CLOSED',
          failure_count = 0,
          next_attempt_at = NULL
      WHERE id = 1 AND state != 'CLOSED'
    `).run();
  } catch {
    // Non-fatal if circuit_breaker table is not provisioned in minimal test DB
  }
}

/**
 * Processes eligible outbox events using an explicit local/mock delivery adapter
 * or DRY_RUN simulation adapter.
 */
export async function processOutboxEvents(env, options = {}) {
  if (!env?.DB) {
    return { blocked: true, processed: 0, reason: 'DB_UNAVAILABLE' };
  }

  const freezeState = await resolveMaintenanceFreezeState(env);
  if (freezeState.freezeActive) {
    return { blocked: true, paused: true, processed: 0, reason: 'MAINTENANCE_FREEZE_ACTIVE', source: freezeState.source };
  }

  const syncState = resolveSyncMode(env);
  const isDryRunSim = !syncState.isLive && options.simulateDryRun === true;
  const adapter = options.adapter || (isDryRunSim ? createDryRunSimulationAdapter() : null);

  if (!adapter || typeof adapter.deliver !== 'function') {
    return { blocked: true, processed: 0, reason: 'OUTBOX_DELIVERY_NOT_IMPLEMENTED' };
  }
  if (adapter.isLocalOnly !== true) {
    return { blocked: true, processed: 0, reason: 'REMOTE_ADAPTER_BLOCKED_IN_LOCAL_MODE' };
  }

  const nowMs = Number(options.nowMs ?? Date.now());
  const nowSql = toSqliteDatetime(new Date(nowMs));
  const batchLimit = Math.min(MAX_BATCH_LIMIT, Math.max(1, Number(options.batchLimit ?? DEFAULT_BATCH_LIMIT)));
  const leaseSeconds = Math.max(5, Number(options.leaseSeconds ?? DEFAULT_LEASE_SECONDS));

  if (await isCircuitBreakerOpen(env.DB, nowSql)) {
    return { paused: true, processed: 0, reason: 'CIRCUIT_BREAKER_OPEN' };
  }

  const recovered = await recoverStaleProcessingEvents(env.DB, nowMs);

  const candidatesRes = await env.DB.prepare(`
    SELECT sync_outbox.id, sync_outbox.entity_name, sync_outbox.entity_id, sync_outbox.action, sync_outbox.payload,
           sync_outbox.status, sync_outbox.retry_count, sync_outbox.next_retry_at,
           sync_outbox.claim_token, sync_outbox.claim_version, sync_outbox.lease_expires_at,
           sync_outbox.correlation_id, sync_outbox.created_at, sync_outbox.updated_at
    FROM sync_outbox
    WHERE (
        sync_outbox.status = 'PENDING'
        OR (sync_outbox.status = 'RETRY' AND (sync_outbox.next_retry_at IS NULL OR sync_outbox.next_retry_at <= ?))
      )
      AND ${buildNoBlockingPrerequisiteSql('sync_outbox')}
    ORDER BY ${buildOrderBySql('sync_outbox')}
    LIMIT ?
  `).bind(nowSql, batchLimit).all();

  const candidates = candidatesRes?.results || [];
  const summary = {
    dryRun: isDryRunSim,
    recovered,
    selected: candidates.length,
    claimed: 0,
    processed: 0,
    delivered: 0,
    completed: 0,
    retried: 0,
    deadLettered: 0,
    leaseConflicts: 0,
    leaseLost: 0,
    dependencyBlocked: 0,
    paused: false,
    results: []
  };

  const blockedEntities = new Set();

  for (const row of candidates) {
    const entityKey = `${row.entity_name}::${row.entity_id}`;
    if (blockedEntities.has(entityKey)) {
      summary.dependencyBlocked += 1;
      continue;
    }

    const claimToken = crypto.randomUUID();
    const expectedPrevVersion = Number(row.claim_version || 0);
    const nextClaimVersion = expectedPrevVersion + 1;
    const leaseExpiresAt = toSqliteDatetime(new Date(nowMs + leaseSeconds * 1000));

    // Atomic conditional claim: only 1 processor can claim PENDING/RETRY AND only if no earlier prerequisite blocks it
    const claimRes = await env.DB.prepare(`
      UPDATE sync_outbox
      SET status = 'PROCESSING',
          claim_token = ?,
          claim_version = ?,
          lease_expires_at = ?,
          updated_at = ?
      WHERE id = ?
        AND COALESCE(claim_version, 0) = ?
        AND (
          status = 'PENDING'
          OR (status = 'RETRY' AND (next_retry_at IS NULL OR next_retry_at <= ?))
        )
        AND ${buildNoBlockingPrerequisiteSql('sync_outbox')}
    `).bind(claimToken, nextClaimVersion, leaseExpiresAt, nowSql, row.id, expectedPrevVersion, nowSql).run();

    const claimedChanges = Number(claimRes?.meta?.changes ?? claimRes?.changes ?? 0);
    if (claimedChanges !== 1) {
      summary.leaseConflicts += 1;
      blockedEntities.add(entityKey);
      continue;
    }
    summary.claimed += 1;

    // Optional test hook to simulate processor crash or slow stall after claim
    if (typeof options.onClaimedBeforeDeliver === 'function') {
      const shouldAbort = await options.onClaimedBeforeDeliver(row, { claimToken, claimVersion: nextClaimVersion, leaseExpiresAt });
      if (shouldAbort === 'CRASH') {
        summary.crashed = true;
        break;
      }
    }

    let deliveryResult;
    try {
      deliveryResult = await adapter.deliver(row);
    } catch (err) {
      deliveryResult = err instanceof Error ? err : new Error(sanitizeOutboxError(err));
    }

    const finishSql = toSqliteDatetime(new Date(Number(options.nowMs ?? Date.now())));

    if (isDryRunSim) {
      await env.DB.prepare(`
        UPDATE sync_outbox
        SET status = ?,
            claim_token = NULL,
            lease_expires_at = NULL,
            updated_at = ?
        WHERE id = ? AND status = 'PROCESSING' AND claim_token = ? AND claim_version = ?
      `).bind(row.status, finishSql, row.id, claimToken, nextClaimVersion).run();
      summary.processed += 1;
      summary.results.push({ id: row.id, simulated: true, valid: Boolean(deliveryResult?.ok), error: deliveryResult?.errorMessage || null });
      continue;
    }

    if (deliveryResult && deliveryResult.ok === true) {
      const updRes = await env.DB.prepare(`
        UPDATE sync_outbox
        SET status = 'DELIVERED',
            claim_token = NULL,
            lease_expires_at = NULL,
            next_retry_at = NULL,
            last_error = NULL,
            updated_at = ?
        WHERE id = ? AND status = 'PROCESSING' AND claim_token = ? AND claim_version = ?
      `).bind(finishSql, row.id, claimToken, nextClaimVersion).run();

      const updChanges = Number(updRes?.meta?.changes ?? updRes?.changes ?? 0);
      if (updChanges === 1) {
        await closeCircuitBreakerOnSuccess(env.DB);
        summary.delivered += 1;
        summary.completed += 1;
        summary.processed += 1;
        summary.results.push({ id: row.id, status: 'DELIVERED', duplicate: Boolean(deliveryResult.duplicate) });
      } else {
        summary.leaseLost += 1;
      }
      continue;
    }

    const decision = evaluateRetryDecision(row.retry_count, deliveryResult, {
      ...options.retryOptions,
      nowMs
    });

    const failRes = await env.DB.prepare(`
      UPDATE sync_outbox
      SET status = ?,
          retry_count = ?,
          next_retry_at = ?,
          last_error = ?,
          claim_token = NULL,
          lease_expires_at = NULL,
          updated_at = ?
      WHERE id = ? AND status = 'PROCESSING' AND claim_token = ? AND claim_version = ?
    `).bind(
      decision.nextStatus,
      decision.nextRetryCount,
      decision.nextRetryAt,
      decision.sanitizedError,
      finishSql,
      row.id,
      claimToken,
      nextClaimVersion
    ).run();

    const failChanges = Number(failRes?.meta?.changes ?? failRes?.changes ?? 0);
    blockedEntities.add(entityKey);
    if (failChanges !== 1) {
      summary.leaseLost += 1;
      continue;
    }

    summary.processed += 1;
    if (decision.nextStatus === 'DEAD_LETTER') {
      summary.deadLettered += 1;
    } else {
      summary.retried += 1;
    }
    summary.results.push({
      id: row.id,
      status: decision.nextStatus,
      category: decision.category,
      retry_count: decision.nextRetryCount
    });

    if (decision.pauseDelivery) {
      await openCircuitBreakerOnAuthFailure(env.DB, nowMs);
      summary.paused = true;
      summary.reason = 'AUTH_ERROR_PAUSED';
      break;
    }
  }

  return summary;
}

/**
 * Internal observability statistics for sync_outbox (Section 14).
 * Counts both canonical 'DELIVERED' and legacy 'COMPLETED' as delivered.
 */
export async function getOutboxStats(env) {
  const countsRows = await env.DB.prepare(`
    SELECT status, COUNT(*) as count
    FROM sync_outbox
    GROUP BY status
  `).all();

  const statusCounts = {
    PENDING: 0,
    PROCESSING: 0,
    DELIVERED: 0,
    COMPLETED: 0,
    RETRY: 0,
    FAILED: 0,
    DEAD_LETTER: 0,
    QUARANTINED_TEST: 0,
    DISCARDED_STAGING_TEST: 0
  };
  for (const r of countsRows?.results || []) {
    if (r.status in statusCounts) {
      statusCounts[r.status] = Number(r.count || 0);
    }
  }

  const oldestPendingRow = await env.DB.prepare(`
    SELECT created_at,
           CAST((julianday('now') - julianday(created_at)) * 86400 AS INTEGER) AS age_seconds
    FROM sync_outbox
    WHERE status IN ('PENDING', 'RETRY')
    ORDER BY created_at ASC
    LIMIT 1
  `).first();

  const lastSuccessRow = await env.DB.prepare(`
    SELECT MAX(updated_at) AS last_success_at
    FROM sync_outbox
    WHERE status IN ('DELIVERED', 'COMPLETED')
  `).first();

  const failureSumRow = await env.DB.prepare(`
    SELECT COALESCE(SUM(retry_count), 0) AS total_retry_attempts
    FROM sync_outbox
  `).first();

  const distRows = await env.DB.prepare(`
    SELECT retry_count, COUNT(*) AS count
    FROM sync_outbox
    GROUP BY retry_count
    ORDER BY retry_count ASC
  `).all();

  const retryAttemptsDistribution = {};
  for (const r of distRows?.results || []) {
    retryAttemptsDistribution[String(r.retry_count)] = Number(r.count || 0);
  }

  return {
    pending_count: statusCounts.PENDING,
    processing_count: statusCounts.PROCESSING,
    delivered_count: statusCounts.DELIVERED + statusCounts.COMPLETED,
    retry_count: statusCounts.RETRY,
    failed_count: statusCounts.FAILED,
    dead_letter_count: statusCounts.DEAD_LETTER,
    quarantined_test_count: statusCounts.QUARANTINED_TEST,
    discarded_staging_test_count: statusCounts.DISCARDED_STAGING_TEST,
    oldest_pending_age_seconds: oldestPendingRow ? Math.max(0, Number(oldestPendingRow.age_seconds || 0)) : null,
    last_successful_processing_time: lastSuccessRow?.last_success_at || null,
    delivery_failure_count: Number(failureSumRow?.total_retry_attempts || 0),
    retry_attempts_distribution: retryAttemptsDistribution
  };
}
