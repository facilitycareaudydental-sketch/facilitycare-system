/**
 * api/src/utils/rate_limit.js
 *
 * Workstream E — E11 5-Tier Distributed Rate Limiting Engine (Cloudflare Free-Tier D1 + Isolate Hybrid)
 *
 * Tiers:
 *   - Tier 1: Login Failure Protection (5 failed attempts / 15 minutes per account & account+IP)
 *   - Tier 2: Authentication Cooldown (15-minute D1 persistent lockout in `auth_cooldowns` + Retry-After header)
 *   - Tier 3: General API Rate Limiting (120 requests / minute / user across all roles including superadmin)
 *   - Tier 4: Export Rate Limiting (10 requests / minute / user across all roles including superadmin, D1-backed)
 *   - Tier 5: Upload / Import Rate Limiting (10 requests / minute / user across all roles including superadmin, D1-backed)
 */

import { error } from './response.js';
import { hasTable } from './schemaCompat.js';

export const TIER1_LOGIN_MAX_ATTEMPTS = 5;
export const TIER2_LOGIN_COOLDOWN_SECONDS = 900; // 15 minutes
export const TIER3_GENERAL_API_LIMIT = 120;      // 120 req / min / user
export const TIER4_EXPORT_LIMIT = 10;            // 10 req / min / user
export const TIER5_UPLOAD_LIMIT = 10;            // 10 req / min / user
export const RATE_WINDOW_MS = 60_000;            // 60 seconds

const generalApiStore = new Map();
const exportStore = new Map();
const uploadStore = new Map();

function checkSlidingWindow(store, key, maxRequests, windowMs = RATE_WINDOW_MS, now = Date.now()) {
  if (store.size > 2000) {
    const cutoff = now - windowMs;
    for (const [k, list] of store.entries()) {
      const filtered = list.filter(t => t > cutoff);
      if (filtered.length === 0) store.delete(k);
      else store.set(k, filtered);
    }
  }
  const cutoff = now - windowMs;
  const timestamps = (store.get(key) || []).filter(t => t > cutoff);
  if (timestamps.length >= maxRequests) {
    store.set(key, timestamps);
    const oldest = timestamps[0] || now;
    const retryAfter = Math.max(1, Math.ceil((oldest + windowMs - now) / 1000));
    return { allowed: false, count: timestamps.length, retryAfter };
  }
  timestamps.push(now);
  store.set(key, timestamps);
  return { allowed: true, count: timestamps.length, retryAfter: 0 };
}

/**
 * Checks and updates a distributed D1 rate-limit bucket inside `auth_cooldowns`
 * for low-frequency / high-cost operations (Tier 4 Export & Tier 5 Upload) or
 * Tier 3 cooldown checks.
 */
async function checkAndIncrementD1Bucket(db, identifier, maxRequests, windowSeconds = 60) {
  if (!db) return { allowed: true, retryAfter: 0 };
  try {
    if (!(await hasTable(db, 'auth_cooldowns'))) {
      return { allowed: true, retryAfter: 0 };
    }
    const existing = await db.prepare(
      `SELECT failed_count, cooldown_until,
              CASE WHEN cooldown_until IS NOT NULL AND cooldown_until > datetime('now') THEN 1 ELSE 0 END AS is_locked,
              CASE WHEN last_failed_at > datetime('now', '-${Number(windowSeconds)} seconds') THEN 1 ELSE 0 END AS in_window
       FROM auth_cooldowns WHERE identifier = ?`
    ).bind(identifier).first();

    if (existing && Number(existing.is_locked) === 1) {
      return { allowed: false, retryAfter: windowSeconds };
    }

    const nextCount = (existing && Number(existing.in_window) === 1)
      ? Number(existing.failed_count || 0) + 1
      : 1;

    if (nextCount > maxRequests) {
      await db.prepare(
        `UPDATE auth_cooldowns
         SET failed_count = ?, cooldown_until = datetime('now', '+${Number(windowSeconds)} seconds'), updated_at = datetime('now')
         WHERE identifier = ?`
      ).bind(nextCount, identifier).run();
      return { allowed: false, retryAfter: windowSeconds };
    }

    const lockClause = nextCount >= maxRequests ? `datetime('now', '+${Number(windowSeconds)} seconds')` : 'NULL';
    await db.prepare(
      `INSERT INTO auth_cooldowns (identifier, failed_count, cooldown_until, last_failed_at, ip_address, created_at, updated_at)
       VALUES (?, ?, ${lockClause}, datetime('now'), 'rate-limit', datetime('now'), datetime('now'))
       ON CONFLICT(identifier) DO UPDATE SET
         failed_count = ?,
         cooldown_until = ${lockClause},
         last_failed_at = CASE WHEN auth_cooldowns.last_failed_at <= datetime('now', '-${Number(windowSeconds)} seconds') THEN datetime('now') ELSE auth_cooldowns.last_failed_at END,
         updated_at = datetime('now')`
    ).bind(identifier, nextCount, nextCount).run();

    return { allowed: true, retryAfter: 0 };
  } catch {
    return { allowed: true, retryAfter: 0 };
  }
}

/**
 * Tier 3: General API Rate Limiting (120 requests / minute / user across all roles including superadmin).
 * Uses fast in-isolate sliding window + writes/checks D1 `auth_cooldowns` on threshold breach so
 * normal requests incur zero D1 write overhead while lockouts persist across isolates.
 */
export async function checkTier3GeneralApiRateLimit(user, env, origin) {
  if (!user?.id) return { allowed: true, response: null };
  const key = `api_rl:${user.id}`;
  const mem = checkSlidingWindow(generalApiStore, key, TIER3_GENERAL_API_LIMIT, RATE_WINDOW_MS);

  if (!mem.allowed) {
    // Persist 60s cooldown to D1 auth_cooldowns so other isolates also enforce the cooldown
    if (env?.DB) {
      try {
        if (await hasTable(env.DB, 'auth_cooldowns')) {
          await env.DB.prepare(
            `INSERT INTO auth_cooldowns (identifier, failed_count, cooldown_until, last_failed_at, ip_address, created_at, updated_at)
             VALUES (?, ?, datetime('now', '+60 seconds'), datetime('now'), 'tier3-api', datetime('now'), datetime('now'))
             ON CONFLICT(identifier) DO UPDATE SET
               failed_count = excluded.failed_count,
               cooldown_until = datetime('now', '+60 seconds'),
               last_failed_at = datetime('now'),
               updated_at = datetime('now')`
          ).bind(key, mem.count).run();
        }
      } catch {}
    }
    return {
      allowed: false,
      response: error(
        `Batas permintaan API tercapai (Maksimal ${TIER3_GENERAL_API_LIMIT} request/menit). Silakan coba lagi dalam ${mem.retryAfter} detik.`,
        429,
        origin,
        env?.ENVIRONMENT,
        { 'Retry-After': String(mem.retryAfter) }
      )
    };
  }

  // If this isolate has count === 1 (fresh isolate or new window), check if D1 has an active Tier 3 cooldown from another isolate
  if (mem.count === 1 && env?.DB) {
    try {
      if (await hasTable(env.DB, 'auth_cooldowns')) {
        const locked = await env.DB.prepare(
          "SELECT cooldown_until FROM auth_cooldowns WHERE identifier = ? AND cooldown_until > datetime('now')"
        ).bind(key).first();
        if (locked) {
          return {
            allowed: false,
            response: error(
              `Batas permintaan API tercapai (Maksimal ${TIER3_GENERAL_API_LIMIT} request/menit). Silakan coba lagi dalam 60 detik.`,
              429,
              origin,
              env?.ENVIRONMENT,
              { 'Retry-After': '60' }
            )
          };
        }
      }
    } catch {}
  }

  return { allowed: true, response: null };
}

/**
 * Tier 4: Export Rate Limiting (10 requests / minute / user across all roles including superadmin).
 * Enforced via both in-isolate sliding window and D1 `auth_cooldowns`.
 */
export async function checkTier4ExportRateLimit(user, env, origin) {
  if (!user?.id) return { allowed: true, response: null };
  const key = `export_rl:${user.id}`;
  const mem = checkSlidingWindow(exportStore, key, TIER4_EXPORT_LIMIT, RATE_WINDOW_MS);
  const d1 = await checkAndIncrementD1Bucket(env?.DB, key, TIER4_EXPORT_LIMIT, 60);

  if (!mem.allowed || !d1.allowed) {
    const retryAfter = Math.max(mem.retryAfter || 60, d1.retryAfter || 60);
    return {
      allowed: false,
      response: error(
        'Batas permintaan export tercapai. Maksimal 10 export per menit. Coba lagi setelah 60 detik.',
        429,
        origin,
        env?.ENVIRONMENT,
        { 'Retry-After': String(retryAfter) }
      )
    };
  }
  return { allowed: true, response: null };
}

/**
 * Tier 5: Upload / Import Rate Limiting (10 requests / minute / user across all roles including superadmin).
 * Enforced via both in-isolate sliding window and D1 `auth_cooldowns`.
 */
export async function checkTier5UploadRateLimit(user, env, origin) {
  if (!user?.id) return { allowed: true, response: null };
  const key = `upload_rl:${user.id}`;
  const mem = checkSlidingWindow(uploadStore, key, TIER5_UPLOAD_LIMIT, RATE_WINDOW_MS);
  const d1 = await checkAndIncrementD1Bucket(env?.DB, key, TIER5_UPLOAD_LIMIT, 60);

  if (!mem.allowed || !d1.allowed) {
    const retryAfter = Math.max(mem.retryAfter || 60, d1.retryAfter || 60);
    return {
      allowed: false,
      response: error(
        'Batas permintaan upload/import tercapai. Maksimal 10 upload per menit. Coba lagi setelah 60 detik.',
        429,
        origin,
        env?.ENVIRONMENT,
        { 'Retry-After': String(retryAfter) }
      )
    };
  }
  return { allowed: true, response: null };
}

export function resetAllE11RateLimits() {
  generalApiStore.clear();
  exportStore.clear();
  uploadStore.clear();
}
