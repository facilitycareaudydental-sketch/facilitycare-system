/**
 * Cloudflare Pages Functions API Proxy: functions/api/[[path]].js
 * ──────────────────────────────────────────────────────────────
 * Proxies /api/* requests from Cloudflare Pages to the backend Worker API.
 * Hardened in Stage N7B.4.1-B, N7B.4.1-D & N7B.4.1-F:
 * - Robust multi-source staging detection with case and whitespace normalization.
 * - Staging detection includes ENVIRONMENT, CF_PAGES_BRANCH, CF_PAGES_URL, and requestUrl.
 * - Conflicting signals fail safe to Staging (never falls back to Production).
 * - Ambiguous / unverified environments fail closed with HTTP 502 (never defaults to Production).
 * - Production default is selected ONLY when explicitly verified as Production.
 * - Exact Production hostname and FCVM denylists strictly enforced.
 */

export const PRODUCTION_WORKER_URL = 'https://fm-operations-api.facilitycare-audydental.workers.dev';
export const PRODUCTION_WORKER_HOSTNAME = 'fm-operations-api.facilitycare-audydental.workers.dev';

export function resolveProxyTarget(env = {}, requestUrl = null) {
  const envVal = String(env.ENVIRONMENT || env.ENV || '').toLowerCase().trim();
  const branchVal = String(env.CF_PAGES_BRANCH || '').toLowerCase().trim();
  const pagesUrl = String(env.CF_PAGES_URL || '').toLowerCase().trim();
  const hostVal = (requestUrl && typeof requestUrl.hostname === 'string')
    ? requestUrl.hostname.toLowerCase().trim()
    : '';

  // 1. Conservative Staging Signal Detection (Case and whitespace normalized)
  const hasStagingSignal = 
    envVal === 'staging' ||
    branchVal === 'staging' ||
    pagesUrl.includes('staging') ||
    hostVal.includes('staging');

  // Staging signals take absolute precedence over any conflicting production signal
  const isStaging = hasStagingSignal;
  const configuredBase = String(env.API_BASE_URL || env.API_BASE || '').trim();

  if (isStaging) {
    if (!configuredBase) {
      const err = new Error('STAGING_BACKEND_UNCONFIGURED: Staging Pages proxy requires an explicit API_BASE_URL environment variable.');
      err.status = 502;
      err.code = 'STAGING_BACKEND_UNCONFIGURED';
      throw err;
    }

    let parsed;
    try {
      parsed = new URL(configuredBase);
    } catch {
      const err = new Error('STAGING_TARGET_SECURITY_VIOLATION: Staging Pages proxy API_BASE_URL must be a valid absolute URL.');
      err.status = 502;
      err.code = 'STAGING_TARGET_SECURITY_VIOLATION';
      throw err;
    }

    if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') {
      const err = new Error('STAGING_TARGET_SECURITY_VIOLATION: Staging Pages proxy API_BASE_URL must use http or https protocol.');
      err.status = 502;
      err.code = 'STAGING_TARGET_SECURITY_VIOLATION';
      throw err;
    }

    const hostname = parsed.hostname.toLowerCase().trim();

    if (hostname === PRODUCTION_WORKER_HOSTNAME) {
      const err = new Error('STAGING_TARGET_SECURITY_VIOLATION: Staging Pages proxy is strictly forbidden from targeting Production Worker API.');
      err.status = 502;
      err.code = 'STAGING_TARGET_SECURITY_VIOLATION';
      throw err;
    }

    if (hostname.includes('fcvm')) {
      const err = new Error('STAGING_TARGET_SECURITY_VIOLATION: Staging Pages proxy is strictly forbidden from targeting FCVM resources.');
      err.status = 502;
      err.code = 'STAGING_TARGET_SECURITY_VIOLATION';
      throw err;
    }

    return configuredBase.replace(/\/+$/, '');
  }

  // 2. Production Path: ONLY if explicitly verified as production and ZERO staging signals present!
  const isExplicitProduction = 
    envVal === 'production' &&
    !hasStagingSignal;

  if (isExplicitProduction) {
    if (configuredBase) {
      let parsed;
      try {
        parsed = new URL(configuredBase);
      } catch {
        const err = new Error('PRODUCTION_TARGET_CONFIG_ERROR: Configured API_BASE_URL is malformed.');
        err.status = 502;
        err.code = 'PRODUCTION_TARGET_CONFIG_ERROR';
        throw err;
      }
      if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') {
        const err = new Error('PRODUCTION_TARGET_CONFIG_ERROR: Configured API_BASE_URL must use http or https.');
        err.status = 502;
        err.code = 'PRODUCTION_TARGET_CONFIG_ERROR';
        throw err;
      }
      const hostname = parsed.hostname.toLowerCase().trim();
      if (hostname.includes('fcvm')) {
        const err = new Error('PRODUCTION_TARGET_SECURITY_VIOLATION: Targeting FCVM is strictly forbidden.');
        err.status = 502;
        err.code = 'PRODUCTION_TARGET_SECURITY_VIOLATION';
        throw err;
      }
      return configuredBase.replace(/\/+$/, '');
    }
    // Verified production default
    return PRODUCTION_WORKER_URL;
  }

  // 3. Ambiguous / Unknown Environment: FAIL-CLOSED (Never default to Production!)
  const err = new Error(`AMBIGUOUS_ENVIRONMENT_ERROR: Environment could not be securely verified (ENVIRONMENT='${env.ENVIRONMENT}', CF_PAGES_BRANCH='${env.CF_PAGES_BRANCH}'). Explicit environment required.`);
  err.status = 502;
  err.code = 'AMBIGUOUS_ENVIRONMENT_ERROR';
  throw err;
}

export async function onRequest(context) {
  const { request, env = {} } = context;
  const url = new URL(request.url);

  let apiBase;
  try {
    apiBase = resolveProxyTarget(env, url);
  } catch (configErr) {
    return new Response(JSON.stringify({
      error: configErr.message,
      code: configErr.code || 'PROXY_CONFIG_ERROR'
    }), {
      status: configErr.status || 502,
      headers: {
        'Content-Type': 'application/json',
        'X-FCMS-Proxy-Error': '1'
      }
    });
  }

  // Rewrite the URL to point to the target worker
  const targetUrl = new URL(url.pathname + url.search, apiBase);

  // Headers sanitization
  const headers = new Headers(request.headers);
  headers.delete('Host');
  headers.delete('Origin');
  headers.delete('Referer');
  headers.delete('Content-Length');

  const fetcher = (typeof context.fetch === 'function') ? context.fetch : fetch;

  return fetcher(targetUrl, {
    method: request.method,
    headers: headers,
    body: request.method !== 'GET' && request.method !== 'HEAD' ? request.body : null,
    redirect: 'manual',
    duplex: 'half'
  });
}
