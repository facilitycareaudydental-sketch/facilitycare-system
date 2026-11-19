/**
 * Fail-Closed API Maintenance Freeze & Recovery Gate (Integrated & Hardened)
 *
 * Separates two fundamentally different operational controls:
 * 1. SYNC_MODE (`DRY_RUN` | `LIVE` via sync_mode.js):
 *    Controls ONLY external Google Sheets webhook/outbox sync.
 *    Note: `SYNC_MODE=PAUSED` resolves fail-closed to `DRY_RUN` and NEVER freezes
 *    inbound HTTP API mutations or background sweepers.
 * 2. Maintenance Freeze (`env.MAINTENANCE_MODE` / `feature_flags.maintenance_freeze`):
 *    Halts ALL inbound HTTP mutations (`POST`, `PUT`, `PATCH`, `DELETE`),
 *    webhooks, imports, logins, scheduled crons, and contract/outbox sweepers
 *    during incident triage, migration recovery, or transaction reconciliation.
 *
 * Recovery Exception Security Invariants:
 * - Never trusts a request header (`X-Recovery-Token`) or URL parameter alone.
 * - Requires backend-verified active `superadmin` principal with valid DB session.
 * - Enforces finite expiration (`expiresAt`) and immediate revocation (`revoked` / deleted session).
 * - Requires explicit non-empty `X-Recovery-Reason` (>= 8 chars) recorded sanitized to `audit_logs`.
 * - Strictly scoped to `RECOVERY_EXEMPT_MUTATION_PATHS` (no general API mutation access).
 * - Zero tokens, passwords, or secrets ever written to logs or audit payloads.
 */

import { authenticate, constantTimeEqual } from './auth.js';
import { buildAuditStatement, sanitizeAuditData } from './audit.js';
import { corsHeaders } from './response.js';
import { resolveSyncMode } from './sync_mode.js';

export const MUTATION_METHODS = Object.freeze(new Set(['POST', 'PUT', 'PATCH', 'DELETE']));
export const READ_ONLY_METHODS = Object.freeze(new Set(['GET', 'HEAD']));

export const ALWAYS_OPEN_OBSERVABILITY_PATHS = Object.freeze(new Set([
  '/api/health',
  '/api/sync/status',
]));

export const RECOVERY_EXEMPT_MUTATION_PATHS = Object.freeze(new Set([
  '/api/auth/logout',
  '/api/sync/actions/snapshot',
  '/api/sync/actions/reconcile',
  '/api/recovery/bundle',
  '/api/recovery/replay',
]));

const MIN_RECOVERY_REASON_LENGTH = 8;
const TRUTHY_VALUES = new Set(['1', 'true', 'yes', 'on', 'freeze', 'active', 'maintenance']);

function isTruthyFlag(val) {
  if (val === true || val === 1) return true;
  if (typeof val !== 'string') return false;
  return TRUTHY_VALUES.has(val.trim().toLowerCase());
}

/**
 * Strips any accidental credential / JWT / hex token patterns from free-text reasons.
 */
export function sanitizeRecoveryReason(rawReason) {
  if (typeof rawReason !== 'string') return '';
  return rawReason
    .trim()
    .replace(/Bearer\s+[A-Za-z0-9._-]+/gi, 'Bearer [REDACTED]')
    .replace(/eyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+/g, '[REDACTED_JWT]')
    .replace(/\b[a-f0-9]{32,}\b/gi, '[REDACTED_HEX]')
    .replace(/\brec_tok_[A-Za-z0-9_-]+\b/gi, '[REDACTED_TOKEN]')
    .slice(0, 300);
}

/**
 * Resolves whether Maintenance Freeze is active from environment and D1 feature_flags.
 * Fails closed on DB probe error unless MAINTENANCE_FAIL_CLOSED_ON_DB_ERROR === 'false'.
 */
export async function resolveMaintenanceFreezeState(env = {}, options = {}) {
  const safeEnv = env || {};
  const syncState = resolveSyncMode(safeEnv);
  const envFreeze = isTruthyFlag(safeEnv.MAINTENANCE_MODE);
  const allowReadOnly = safeEnv.MAINTENANCE_ALLOW_READONLY !== 'false';
  const failClosedOnDbError = safeEnv.MAINTENANCE_FAIL_CLOSED_ON_DB_ERROR !== 'false';
  const skipDbProbe = options.skipDbProbe === true;

  let dbFreeze = false;
  let dbProbeError = null;

  if (!skipDbProbe && safeEnv.DB && typeof safeEnv.DB.prepare === 'function') {
    try {
      const stmt = safeEnv.DB.prepare(
        "SELECT flag_value FROM feature_flags WHERE flag_key = 'maintenance_freeze' LIMIT 1"
      );
      if (stmt && typeof stmt.first === 'function') {
        const row = await stmt.first();
        if (row && isTruthyFlag(row.flag_value)) {
          dbFreeze = true;
        }
      }
    } catch (err) {
      dbProbeError = err?.message || String(err);
      const missingTable = /no such table:\s*feature_flags/i.test(dbProbeError);
      if (!missingTable && failClosedOnDbError) {
        dbFreeze = true;
      }
    }
  }

  const freezeActive = envFreeze || dbFreeze;
  const source = envFreeze && dbFreeze
    ? 'ENV+DB_FEATURE_FLAG'
    : envFreeze
      ? 'ENV_MAINTENANCE_MODE'
      : dbFreeze
        ? (dbProbeError ? 'DB_PROBE_FAIL_CLOSED' : 'DB_FEATURE_FLAG')
        : 'NONE';

  return {
    freezeActive,
    source,
    allowReadOnly,
    dbProbeError,
    syncModeState: syncState,
    rawSyncModeWasPaused: typeof safeEnv.SYNC_MODE === 'string' && safeEnv.SYNC_MODE.trim().toUpperCase() === 'PAUSED',
    syncModeProvidesApiFreeze: false,
  };
}

/**
 * Verifies backend-authenticated recovery principal, time-bounded non-revoked grant,
 * optional cryptographic recovery token second factor, and mandatory audit reason.
 */
export async function hasValidRecoveryAuthorization(request, env = {}, authenticatedUser = null) {
  // 1. Verify backend principal via existing active session + active superadmin DB record
  let user = authenticatedUser;
  if (!user && request && env?.DB) {
    try {
      user = await authenticate(request, env);
    } catch {
      user = null;
    }
  }

  if (
    !user ||
    String(user.role || '').toLowerCase() !== 'superadmin' ||
    Number(user.is_active ?? 1) !== 1 ||
    !user.sessionId ||
    user.isRestrictedToken === true
  ) {
    return {
      authorized: false,
      via: 'NONE',
      reason: 'RECOVERY_PRINCIPAL_REQUIRED',
      actor: null,
    };
  }

  // 2. Second-factor recovery token verification if configured (never sufficient alone)
  const configuredToken = typeof env?.RECOVERY_FREEZE_TOKEN === 'string'
    ? env.RECOVERY_FREEZE_TOKEN.trim()
    : '';
  const presentedToken = request?.headers?.get?.('X-Recovery-Token') || '';
  if (configuredToken.length >= 16) {
    if (!presentedToken || !constantTimeEqual(presentedToken.trim(), configuredToken)) {
      return {
        authorized: false,
        via: 'NONE',
        reason: 'RECOVERY_SECOND_FACTOR_INVALID',
        actor: { id: user.id, username: user.username, role: user.role },
      };
    }
  }

  // 3. Time-bounded & revocable recovery access grant check (DB grant preferred, else bounded env window)
  const nowMs = Date.now();
  let grantVerified = false;
  let via = 'DB_RECOVERY_GRANT';
  let grantExpiresAt = null;

  if (env?.DB && typeof env.DB.prepare === 'function') {
    try {
      const stmt = env.DB.prepare(
        "SELECT flag_value FROM feature_flags WHERE flag_key = 'recovery_access_grant' LIMIT 1"
      );
      if (stmt && typeof stmt.first === 'function') {
        const grantRow = await stmt.first();
        if (grantRow && grantRow.flag_value) {
          let grant;
          try {
            grant = JSON.parse(grantRow.flag_value);
          } catch {
            return { authorized: false, via: 'NONE', reason: 'RECOVERY_GRANT_MALFORMED', actor: { id: user.id, username: user.username, role: user.role } };
          }
          if (!grant || typeof grant !== 'object' || grant.revoked === true || grant.revokedAt) {
            return { authorized: false, via: 'NONE', reason: 'RECOVERY_GRANT_REVOKED', actor: { id: user.id, username: user.username, role: user.role } };
          }
          if (Array.isArray(grant.allowedUserIds) && !grant.allowedUserIds.includes(user.id)) {
            return { authorized: false, via: 'NONE', reason: 'RECOVERY_ACTOR_NOT_GRANTED', actor: { id: user.id, username: user.username, role: user.role } };
          }
          const expMs = typeof grant.expiresAt === 'string' ? Date.parse(grant.expiresAt) : NaN;
          if (!Number.isFinite(expMs) || expMs <= nowMs) {
            return { authorized: false, via: 'NONE', reason: 'RECOVERY_GRANT_EXPIRED', actor: { id: user.id, username: user.username, role: user.role } };
          }
          grantVerified = true;
          grantExpiresAt = new Date(expMs).toISOString();
        }
      }
    } catch {
      return { authorized: false, via: 'NONE', reason: 'RECOVERY_GRANT_PROBE_FAILED', actor: { id: user.id, username: user.username, role: user.role } };
    }
  }

  if (!grantVerified) {
    // Bounded env-level recovery grant requires BOTH finite future expiry AND recovery:operator claim or valid second factor
    const rawExp = env?.RECOVERY_ACCESS_EXPIRES_AT;
    const expMs = typeof rawExp === 'string' && rawExp.trim() ? Date.parse(rawExp.trim()) : NaN;
    if (!Number.isFinite(expMs) || expMs <= nowMs) {
      return { authorized: false, via: 'NONE', reason: 'RECOVERY_GRANT_MISSING_OR_EXPIRED', actor: { id: user.id, username: user.username, role: user.role } };
    }
    const hasOperatorScope = user.scope === 'recovery:operator' || user.tokenPayload?.scope === 'recovery:operator';
    const hasSecondFactor = configuredToken.length >= 16 && presentedToken && constantTimeEqual(presentedToken.trim(), configuredToken);
    if (!hasOperatorScope && !hasSecondFactor) {
      return { authorized: false, via: 'NONE', reason: 'RECOVERY_SCOPE_OR_TOKEN_REQUIRED', actor: { id: user.id, username: user.username, role: user.role } };
    }
    grantVerified = true;
    grantExpiresAt = new Date(expMs).toISOString();
    via = hasOperatorScope && hasSecondFactor ? 'SUPERADMIN_SESSION+TOKEN+TTL' : hasOperatorScope ? 'SUPERADMIN_RECOVERY_SCOPE+TTL' : 'SUPERADMIN_SESSION+TOKEN+TTL';
  }

  // 4. Mandatory non-empty sanitized audit reason header
  const rawReason = request?.headers?.get?.('X-Recovery-Reason') || '';
  const sanitizedReason = sanitizeRecoveryReason(rawReason);
  if (sanitizedReason.length < MIN_RECOVERY_REASON_LENGTH) {
    return {
      authorized: false,
      via: 'NONE',
      reason: 'RECOVERY_REASON_REQUIRED',
      actor: { id: user.id, username: user.username, role: user.role },
    };
  }

  return {
    authorized: true,
    via,
    reason: sanitizedReason,
    expiresAt: grantExpiresAt,
    actor: {
      id: user.id,
      username: user.username,
      role: user.role,
      sessionId: user.sessionId,
    },
  };
}

/**
 * Records sanitized recovery exception usage to audit_logs without exposing any credentials.
 */
async function logSanitizedRecoveryException(env, authResult, request, path) {
  if (!env?.DB || !authResult?.actor) return;
  try {
    const safePayload = sanitizeAuditData({
      event: 'RECOVERY_EXCEPTION_USED',
      method: String(request?.method || '').toUpperCase(),
      path,
      via: authResult.via,
      expiresAt: authResult.expiresAt,
      recovery_reason: authResult.reason,
    });
    const stmt = buildAuditStatement(
      env,
      authResult.actor,
      'UPDATE',
      'maintenance_freeze',
      path,
      null,
      safePayload,
      { scope: 'recovery_exception', granted: true },
      'SUCCESS',
      false
    );
    if (stmt && typeof stmt.run === 'function') {
      await stmt.run();
    }
  } catch {
    // Non-fatal: never emit secrets or block already-authorized recovery triage on audit insert glitch
  }
}

/**
 * Enforces Maintenance Freeze on an incoming HTTP request.
 * Returns null if allowed to proceed, or a fail-closed HTTP 503 Response if blocked.
 */
export async function enforceMaintenanceFreeze(request, env = {}, origin = null, authenticatedUser = null) {
  const method = String(request?.method || 'GET').toUpperCase();
  if (method === 'OPTIONS') return null;

  const state = await resolveMaintenanceFreezeState(env);
  if (!state.freezeActive) return null;

  const url = new URL(request.url);
  const path = url.pathname;

  if (READ_ONLY_METHODS.has(method)) {
    if (ALWAYS_OPEN_OBSERVABILITY_PATHS.has(path)) {
      return null;
    }
    if (state.allowReadOnly) {
      return null;
    }
    return buildFreezeResponse(state, origin, env?.ENVIRONMENT, 'READ_ONLY_DISABLED_DURING_FREEZE');
  }

  if (MUTATION_METHODS.has(method) && RECOVERY_EXEMPT_MUTATION_PATHS.has(path)) {
    const auth = await hasValidRecoveryAuthorization(request, env, authenticatedUser);
    if (auth.authorized) {
      await logSanitizedRecoveryException(env, auth, request, path);
      return null;
    }
    return buildFreezeResponse(state, origin, env?.ENVIRONMENT, auth.reason || 'RECOVERY_AUTHORIZATION_REQUIRED');
  }

  return buildFreezeResponse(state, origin, env?.ENVIRONMENT, 'MUTATIONS_FROZEN_DURING_RECOVERY');
}

function buildFreezeResponse(state, origin, environment, reason) {
  return new Response(
    JSON.stringify({
      success: false,
      error: 'Service temporarily unavailable due to Maintenance Freeze for data recovery',
      code: 'MAINTENANCE_FREEZE_ACTIVE',
      reason,
      freeze_source: state.source,
      sync_mode: state.syncModeState.mode,
      allow_read_only: state.allowReadOnly,
    }),
    {
      status: 503,
      headers: {
        'Content-Type': 'application/json',
        'Retry-After': '300',
        'X-Maintenance-Freeze': 'ACTIVE',
        'X-Freeze-Reason': reason,
        ...corsHeaders(origin, environment),
      },
    }
  );
}
