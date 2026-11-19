import { authenticate, hasPermission } from './auth.js';
import { unauthorized, forbidden } from './response.js';
import { getPagination, HARD_MAX_PAGE_LIMIT } from './pagination.js';

/**
 * Authenticates and verifies `'export'` RBAC permission for `/api/<module>/export`.
 * Returns `{ user }` on success or `{ response }` (401/403) on denial.
 */
export async function requireExportAccess(request, env, origin, moduleKey) {
  const user = await authenticate(request, env);
  if (!user) {
    return { user: null, response: unauthorized(origin, env?.ENVIRONMENT) };
  }
  if (!hasPermission(user, moduleKey, 'export')) {
    return { user: null, response: forbidden(origin, env?.ENVIRONMENT) };
  }
  return { user, response: null };
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
