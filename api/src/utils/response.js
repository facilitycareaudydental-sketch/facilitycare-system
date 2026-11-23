export const DEFAULT_ALLOWED_ORIGIN = 'https://facilitycare-system.pages.dev';
export const STAGING_DEFAULT_ALLOWED_ORIGIN = 'https://staging.facilitycare-system-staging.pages.dev';

export const PROD_ORIGIN_REGEX = /^https:\/\/([a-z0-9-]+\.)?facilitycare-system\.pages\.dev$/;
export const STAGING_ORIGIN_REGEX = /^https:\/\/([a-z0-9-]+\.)?facilitycare-system-staging\.pages\.dev$/;

export const LOCAL_ALLOWED_ORIGINS = new Set([
  'http://localhost:8787',
  'http://127.0.0.1:8787',
  'http://localhost:8788',
  'http://127.0.0.1:8788',
  'http://localhost:3000',
  'http://127.0.0.1:3000',
  'http://localhost:5173',
  'http://127.0.0.1:5173',
]);

/**
 * Validates incoming Origin strictly according to the active environment:
 * - ENVIRONMENT === 'staging':
 *     Allows: *.facilitycare-system-staging.pages.dev and explicit staging CORS_ORIGIN, plus local dev.
 *     Strictly REJECTS: production domains (*.facilitycare-system.pages.dev) and any foreign/spoofed domains.
 * - ENVIRONMENT === 'production' (or default):
 *     Allows: *.facilitycare-system.pages.dev, explicit production CORS_ORIGIN, plus local dev.
 */
export function isAllowedOrigin(requestOrigin, configuredOrigin, environment) {
  if (!requestOrigin) return false;
  const norm = requestOrigin.trim();
  if (LOCAL_ALLOWED_ORIGINS.has(norm)) return true;

  const isStaging = (environment === 'staging');

  if (isStaging) {
    // 1. Explicitly reject ANY production domain from reaching Staging Worker
    if (PROD_ORIGIN_REGEX.test(norm)) {
      return false;
    }
    // 2. Allow Staging root, branch, and preview domains
    if (STAGING_ORIGIN_REGEX.test(norm)) {
      return true;
    }
    // 3. Allow explicit configured origins for staging (never allowing production domains)
    if (configuredOrigin) {
      const list = configuredOrigin.split(',').map(s => s.trim()).filter(Boolean);
      if (list.includes(norm) && !PROD_ORIGIN_REGEX.test(norm)) {
        return true;
      }
    }
    return false;
  } else {
    // Production policy: strictly allow production domains & configured origins
    if (PROD_ORIGIN_REGEX.test(norm)) {
      return true;
    }
    if (configuredOrigin) {
      const list = configuredOrigin.split(',').map(s => s.trim()).filter(Boolean);
      if (list.includes(norm)) {
        return true;
      }
    }
    return false;
  }
}

/**
 * Resolves the allowed origin for CORS headers:
 * - If valid and whitelisted, returns the exact normalized requestOrigin.
 * - If requestOrigin is absent (e.g. server-to-server, curl, non-browser clients):
 *     - Staging: returns null (ensuring NO Access-Control-Allow-Origin header is emitted; never substitutes default Staging or Production).
 *     - Production: falls back to DEFAULT_ALLOWED_ORIGIN (preventing reflection of attacker origin while preserving legacy production contract).
 * - If requestOrigin is present but unauthorized:
 *     - Staging: returns null (ensuring NO valid Access-Control-Allow-Origin is ever echoed).
 *     - Production: falls back to DEFAULT_ALLOWED_ORIGIN (preventing reflection of attacker origin).
 */
export function resolveAllowedOrigin(requestOrigin, configuredOrigin, environment) {
  const isStaging = (environment === 'staging');

  if (!requestOrigin || !requestOrigin.trim()) {
    if (isStaging) {
      return null; // In Staging, missing/null Origin NEVER receives an allow-origin header and is not substituted
    }
    const prodFallback = configuredOrigin ? configuredOrigin.split(',')[0].trim() : DEFAULT_ALLOWED_ORIGIN;
    return prodFallback;
  }

  const norm = requestOrigin.trim();
  if (isAllowedOrigin(norm, configuredOrigin, environment)) {
    return norm;
  }

  // Unauthorized origin provided:
  if (isStaging) {
    return null; // In Staging, never return an allowed origin for untrusted or cross-environment callers
  }
  const prodFallback = configuredOrigin ? configuredOrigin.split(',')[0].trim() : DEFAULT_ALLOWED_ORIGIN;
  return prodFallback;
}

export function corsHeaders(origin) {
  const headers = {
    'Access-Control-Allow-Methods': 'GET, POST, PUT, PATCH, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, Authorization, X-CSRF-Token, X-FCMS-SIGNATURE, Idempotency-Key, X-Idempotency-Key, X-Import-Recover',
    'Access-Control-Max-Age': '86400',
    'Vary': 'Origin',
    'X-Content-Type-Options': 'nosniff',
    'X-Frame-Options': 'DENY',
    'Referrer-Policy': 'strict-origin-when-cross-origin',
    'Strict-Transport-Security': 'max-age=31536000; includeSubDomains; preload',
    'Content-Security-Policy': "default-src 'none'; frame-ancestors 'none'",
  };

  if (origin) {
    headers['Access-Control-Allow-Origin'] = origin;
    headers['Access-Control-Allow-Credentials'] = 'true';
  }

  return headers;
}

export function ok(data, status = 200, origin, environment, extraHeaders = {}) {
  let extra = extraHeaders;
  let env = environment;
  if (environment && typeof environment === 'object' && !Array.isArray(environment)) {
    extra = environment;
    env = undefined;
  }
  return new Response(JSON.stringify({ success: true, data }), {
    status,
    headers: { 'Content-Type': 'application/json', ...corsHeaders(origin, env), ...extra },
  });
}

export function paginated(data, total, page, limit, origin, environment, extra = null) {
  const safeLimit = Math.max(1, Number(limit) || 20);
  const safeTotal = Math.max(0, Number(total) || 0);
  const payload = {
    success: true,
    data,
    pagination: {
      total: safeTotal,
      page,
      limit: safeLimit,
      pages: Math.ceil(safeTotal / safeLimit),
    },
  };
  if (extra && typeof extra === 'object' && !Array.isArray(extra)) {
    Object.assign(payload, extra);
  }
  return new Response(JSON.stringify(payload), {
    status: 200,
    headers: { 'Content-Type': 'application/json', ...corsHeaders(origin, environment) },
  });
}

export function error(message, status = 400, origin, environment, extraHeaders = {}) {
  let extra = extraHeaders;
  let env = environment;
  if (environment && typeof environment === 'object' && !Array.isArray(environment)) {
    extra = environment;
    env = undefined;
  }
  const headers = { 'Content-Type': 'application/json', ...corsHeaders(origin, env), ...(extra || {}) };
  if (status === 429 && !headers['Retry-After']) {
    const secMatch = String(message || '').match(/(\d+)\s*(?:seconds?|detik)/i);
    const minMatch = String(message || '').match(/(\d+)\s*(?:minutes?|menit)/i);
    if (secMatch) {
      headers['Retry-After'] = String(Math.max(1, Number(secMatch[1])));
    } else if (minMatch) {
      headers['Retry-After'] = String(Math.max(1, Number(minMatch[1]) * 60));
    } else {
      headers['Retry-After'] = '60';
    }
  }
  return new Response(JSON.stringify({ success: false, error: message }), {
    status,
    headers,
  });
}

export function unauthorized(origin, environment) {
  return error('Unauthorized', 401, origin, environment);
}

export function forbidden(origin, environment) {
  return error('Forbidden - insufficient permissions', 403, origin, environment);
}

export function notFound(origin, environment) {
  return error('Not found', 404, origin, environment);
}

export function options(origin, environment) {
  return new Response(null, {
    status: 204,
    headers: corsHeaders(origin, environment),
  });
}
