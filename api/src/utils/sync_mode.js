/**
 * Fail-Closed SYNC_MODE Policy Resolver
 * ──────────────────────────────────────────────────────────────────────────────
 * Strictly guarantees that sync NEVER runs in LIVE mode by default, fallback,
 * omission, whitespace, or unrecognized configuration value.
 *
 * Rules:
 * - Only exact `env.SYNC_MODE === 'LIVE'` (in a non-staging environment) enables LIVE.
 * - `DRY_RUN`, undefined, null, empty, whitespace, or unknown values resolve to `DRY_RUN`
 *   with zero external/DB write side-effects (fail-closed).
 * - Staging environments (`ENVIRONMENT === 'staging'` or `ENV === 'staging'`) are
 *   permanently locked to `DRY_RUN` even if `SYNC_MODE === 'LIVE'` is set.
 */

export const SYNC_MODES = Object.freeze({
  DRY_RUN: 'DRY_RUN',
  LIVE: 'LIVE'
});

export const ALLOWED_SYNC_MODES = Object.freeze([SYNC_MODES.DRY_RUN, SYNC_MODES.LIVE]);

/**
 * Resolves the effective sync mode fail-closed.
 * @param {object} [env={}] Worker or runtime environment bindings
 * @returns {{ mode: 'DRY_RUN'|'LIVE', raw: string|null, isLive: boolean, isDryRun: boolean, isExplicit: boolean, valid: boolean, reason: string|null }}
 */
export function resolveSyncMode(env = {}) {
  const safeEnv = (env && typeof env === 'object') ? env : {};
  const environmentName = String(safeEnv.ENVIRONMENT ?? safeEnv.ENV ?? '').trim().toLowerCase();
  const rawValue = safeEnv.SYNC_MODE;

  if (rawValue === undefined || rawValue === null) {
    return {
      mode: SYNC_MODES.DRY_RUN,
      raw: null,
      isLive: false,
      isDryRun: true,
      isExplicit: false,
      valid: false,
      reason: 'SYNC_MODE_MISSING_FAIL_CLOSED'
    };
  }

  if (typeof rawValue !== 'string') {
    return {
      mode: SYNC_MODES.DRY_RUN,
      raw: String(rawValue),
      isLive: false,
      isDryRun: true,
      isExplicit: false,
      valid: false,
      reason: 'SYNC_MODE_INVALID_TYPE_FAIL_CLOSED'
    };
  }

  const trimmed = rawValue.trim();
  if (!trimmed) {
    return {
      mode: SYNC_MODES.DRY_RUN,
      raw: null,
      isLive: false,
      isDryRun: true,
      isExplicit: false,
      valid: false,
      reason: 'SYNC_MODE_EMPTY_FAIL_CLOSED'
    };
  }

  if (trimmed === SYNC_MODES.DRY_RUN) {
    return {
      mode: SYNC_MODES.DRY_RUN,
      raw: SYNC_MODES.DRY_RUN,
      isLive: false,
      isDryRun: true,
      isExplicit: true,
      valid: true,
      reason: null
    };
  }

  if (trimmed === SYNC_MODES.LIVE) {
    if (environmentName === 'staging') {
      return {
        mode: SYNC_MODES.DRY_RUN,
        raw: SYNC_MODES.LIVE,
        isLive: false,
        isDryRun: true,
        isExplicit: false,
        valid: false,
        reason: 'STAGING_LIVE_SYNC_FORBIDDEN_FAIL_CLOSED'
      };
    }

    return {
      mode: SYNC_MODES.LIVE,
      raw: SYNC_MODES.LIVE,
      isLive: true,
      isDryRun: false,
      isExplicit: true,
      valid: true,
      reason: null
    };
  }

  // Unrecognized value -> Fail-Closed to DRY_RUN
  return {
    mode: SYNC_MODES.DRY_RUN,
    raw: trimmed,
    isLive: false,
    isDryRun: true,
    isExplicit: false,
    valid: false,
    reason: 'SYNC_MODE_UNKNOWN_FAIL_CLOSED'
  };
}
