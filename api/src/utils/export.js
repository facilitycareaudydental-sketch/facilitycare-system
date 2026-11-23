import { authenticate, hasPermission } from './auth.js';
import { unauthorized, forbidden } from './response.js';
import { getPagination, HARD_MAX_PAGE_LIMIT } from './pagination.js';
import { checkTier4ExportRateLimit, checkTier5UploadRateLimit } from './rate_limit.js';

/**
 * Authenticates and verifies `'export'` RBAC permission for `/api/<module>/export`.
 * Also enforces Tier 4 export rate limit: 10 requests/minute/user (D1 + isolate hybrid).
 * Returns `{ user }` on success or `{ response }` (401/403/429) on denial.
 */
export async function requireExportAccess(request, env, origin, moduleKey) {
  const user = await authenticate(request, env);
  if (!user) {
    return { user: null, response: unauthorized(origin, env?.ENVIRONMENT) };
  }
  if (!hasPermission(user, moduleKey, 'export')) {
    return { user: null, response: forbidden(origin, env?.ENVIRONMENT) };
  }
  // Tier 4: Export rate limit (10 req/min/user, distributed D1 + isolate)
  const rl = await checkTier4ExportRateLimit(user, env, origin);
  if (!rl.allowed) {
    return {
      user: null,
      response: rl.response,
    };
  }
  return { user, response: null };
}

/**
 * Checks Tier 5 upload rate limit for the authenticated user: 10 req/min/user (D1 + isolate).
 * Returns `{ allowed: true }` or `{ allowed: false, response }`.
 */
export async function checkUploadRateLimit(user, envOrOrigin, maybeOrigin) {
  if (!user?.id) return { allowed: true };
  const isEnvObj = envOrOrigin && typeof envOrOrigin === 'object';
  const env = isEnvObj ? envOrOrigin : null;
  const origin = isEnvObj ? maybeOrigin : envOrOrigin;
  return checkTier5UploadRateLimit(user, env, origin);
}

/**
 * Parses pagination for `/export` endpoints with strict ceiling `<= HARD_MAX_PAGE_LIMIT` (100).
 */
export function getExportPagination(url) {
  return getPagination(url, {
    defaultLimit: HARD_MAX_PAGE_LIMIT,
    maxLimit: HARD_MAX_PAGE_LIMIT,
  });
}
