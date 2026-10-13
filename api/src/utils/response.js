export const DEFAULT_ALLOWED_ORIGIN = 'https://facilitycare-system.pages.dev';

const ALLOWED_ORIGIN_REGEX = /^https:\/\/([a-z0-9-]+\.)?facilitycare-system\.pages\.dev$/;
const LOCAL_ALLOWED_ORIGINS = new Set([
  'http://localhost:8787',
  'http://127.0.0.1:8787',
  'http://localhost:3000',
  'http://127.0.0.1:3000',
  'http://localhost:5173',
  'http://127.0.0.1:5173',
]);

/**
 * Validates incoming Origin against strict whitelist:
 * - Production: https://facilitycare-system.pages.dev
 * - Preview: https://*.facilitycare-system.pages.dev
 * - Local Dev: localhost / 127.0.0.1 on ports 8787, 3000, 5173
 * - Optional configured env.CORS_ORIGIN
 * Safe fallback: DEFAULT_ALLOWED_ORIGIN (never reflects untrusted origin).
 */
export function resolveAllowedOrigin(requestOrigin, configuredOrigin) {
  if (!requestOrigin) {
    return configuredOrigin || DEFAULT_ALLOWED_ORIGIN;
  }
  const norm = requestOrigin.trim();
  if (LOCAL_ALLOWED_ORIGINS.has(norm)) {
    return norm;
  }
  if (ALLOWED_ORIGIN_REGEX.test(norm)) {
    return norm;
  }
  if (configuredOrigin && norm === configuredOrigin.trim()) {
    return norm;
  }
  return DEFAULT_ALLOWED_ORIGIN;
}

export function corsHeaders(origin) {
  return {
    'Access-Control-Allow-Origin': origin || DEFAULT_ALLOWED_ORIGIN,
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization',
    'Access-Control-Max-Age': '86400',
  };
}

export function ok(data, status = 200, origin) {
  return new Response(JSON.stringify({ success: true, data }), {
    status,
    headers: { 'Content-Type': 'application/json', ...corsHeaders(origin) },
  });
}

export function paginated(data, total, page, limit, origin) {
  return new Response(JSON.stringify({
    success: true,
    data,
    pagination: { total, page, limit, pages: Math.ceil(total / limit) },
  }), {
    status: 200,
    headers: { 'Content-Type': 'application/json', ...corsHeaders(origin) },
  });
}

export function error(message, status = 400, origin) {
  return new Response(JSON.stringify({ success: false, error: message }), {
    status,
    headers: { 'Content-Type': 'application/json', ...corsHeaders(origin) },
  });
}

export function unauthorized(origin) {
  return error('Unauthorized', 401, origin);
}

export function forbidden(origin) {
  return error('Forbidden - insufficient permissions', 403, origin);
}

export function notFound(origin) {
  return error('Not found', 404, origin);
}

export function options(origin) {
  return new Response(null, {
    status: 204,
    headers: corsHeaders(origin),
  });
}
