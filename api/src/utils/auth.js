// Auth utilities - JWT + password hashing + RBAC helpers using Web Crypto API
import { isAllowedOrigin } from './response.js';

export const OFFICIAL_ROLES = Object.freeze(['superadmin', 'fc_specialist_leader']);
export const SESSION_COOKIE_NAME = 'fcms_session';
export const CSRF_COOKIE_NAME = 'fcms_csrf';

export const PBKDF2_MIN_ITERATIONS = 10000;
export const PBKDF2_MAX_ITERATIONS = 100000;
export const PBKDF2_DEFAULT_ITERATIONS = 100000;
export const PASSWORD_MIN_LENGTH = 12;
export const PASSWORD_MAX_LENGTH = 256;
export const NORMAL_TOKEN_TTL_MS = 8 * 60 * 60 * 1000;
export const RESTRICTED_TOKEN_TTL_MS = 1 * 60 * 60 * 1000;
export const MAX_IAT_FUTURE_SKEW_MS = 60 * 1000;

/**
 * Official assignable operational modules for Editor Khusus (single source of truth
 * shared with api/src/routes/users.js).
 */
export const ASSIGNABLE_MODULES = Object.freeze([
  'employees',
  'contracts',
  'sp',
  'mutasi',
  'overtime',
  'schedule',
  'os_migrations',
  'issues',
  'training',
  'reports',
  'sop',
  'checklist',
  'hygiene_standards',
  'dac',
  'one_on_one',
  'relievers',
  'branches'
]);

const ASSIGNABLE_MODULES_SET = new Set(ASSIGNABLE_MODULES);

const RESTRICTED_ADMIN_MODULES = new Set([
  'users',
  'audit_logs',
  'audit-logs',
  'sync',
  'sync_dashboard',
  'monitoring',
  'settings',
  'system_settings',
  'import'
]);

/**
 * Constant-time string comparison to prevent timing attacks.
 * Pure byte-level XOR loop compatible across Cloudflare Workers and Node.js runtimes.
 */
export function constantTimeEqual(a, b) {
  if (typeof a !== 'string' || typeof b !== 'string') return false;
  const encoder = new TextEncoder();
  const aBuf = encoder.encode(a);
  const bBuf = encoder.encode(b);
  let mismatch = aBuf.byteLength ^ bBuf.byteLength;
  const maxLen = Math.max(aBuf.byteLength, bBuf.byteLength);
  for (let i = 0; i < maxLen; i++) {
    const byteA = i < aBuf.byteLength ? aBuf[i] : 0;
    const byteB = i < bBuf.byteLength ? bBuf[i] : 0;
    mismatch |= (byteA ^ byteB);
  }
  return mismatch === 0;
}

/**
 * Legacy SHA-256 hash utility (64 hex characters, unsalted).
 * Preserved for verifying pre-existing user passwords before transparent upgrade.
 */
export async function hashPasswordLegacySHA256(password) {
  const encoder = new TextEncoder();
  const data = encoder.encode(password);
  const hashBuffer = await crypto.subtle.digest('SHA-256', data);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
}

/**
 * Centralized validator for PBKDF2_ITERATIONS.
 * - Defaults to PBKDF2_DEFAULT_ITERATIONS (100,000) when undefined, null, or empty string.
 * - Strictly rejects non-positive integers, floats, NaN, booleans, objects, or non-numeric strings.
 * - Enforces bounds [PBKDF2_MIN_ITERATIONS (10,000), PBKDF2_MAX_ITERATIONS (100,000)]
 *   to prevent weak hashing on the low end and Worker CPU exhaustion on the high end.
 */
export function validatePbkdf2Iterations(rawIterations) {
  if (rawIterations === undefined || rawIterations === null || rawIterations === '') {
    return PBKDF2_DEFAULT_ITERATIONS;
  }

  let parsed;
  if (typeof rawIterations === 'number') {
    if (!Number.isSafeInteger(rawIterations) || rawIterations <= 0) {
      throw new Error('Invalid PBKDF2_ITERATIONS: must be a positive integer');
    }
    parsed = rawIterations;
  } else if (typeof rawIterations === 'string') {
    const trimmed = rawIterations.trim();
    if (!/^\d+$/.test(trimmed)) {
      throw new Error('Invalid PBKDF2_ITERATIONS: must be a positive integer');
    }
    parsed = Number(trimmed);
    if (!Number.isSafeInteger(parsed) || parsed <= 0) {
      throw new Error('Invalid PBKDF2_ITERATIONS: must be a positive integer');
    }
  } else {
    throw new Error('Invalid PBKDF2_ITERATIONS: must be a positive integer');
  }

  if (parsed < PBKDF2_MIN_ITERATIONS || parsed > PBKDF2_MAX_ITERATIONS) {
    throw new Error(
      `Invalid PBKDF2_ITERATIONS: must be between ${PBKDF2_MIN_ITERATIONS} and ${PBKDF2_MAX_ITERATIONS}`
    );
  }

  return parsed;
}

/**
 * Centralized password complexity validator for authentication endpoints
 * (Bootstrap & Change Password).
 * Enforces:
 * - String type
 * - Minimum 12 characters (max 256 characters)
 * - At least one uppercase letter, one lowercase letter, one digit, and one special character
 */
export function validatePasswordPolicy(password) {
  if (typeof password !== 'string') {
    return { valid: false, error: 'Password must be a valid string' };
  }
  if (password.length < PASSWORD_MIN_LENGTH) {
    return { valid: false, error: `Password must be at least ${PASSWORD_MIN_LENGTH} characters long` };
  }
  if (password.length > PASSWORD_MAX_LENGTH) {
    return { valid: false, error: `Password must not exceed ${PASSWORD_MAX_LENGTH} characters` };
  }
  const hasUpper = /[A-Z]/.test(password);
  const hasLower = /[a-z]/.test(password);
  const hasNumber = /[0-9]/.test(password);
  const hasSpecial = /[^A-Za-z0-9]/.test(password);
  if (!hasUpper || !hasLower || !hasNumber || !hasSpecial) {
    return {
      valid: false,
      error: 'Password must contain uppercase, lowercase, numbers, and special characters'
    };
  }
  return { valid: true };
}

/**
 * PBKDF2-HMAC-SHA256 password hashing using Web Crypto API.
 * Uses a cryptographically secure 16-byte random salt and validated iterations.
 * Output format: pbkdf2_sha256$<iterations>$<saltHex>$<hashHex>
 */
export async function hashPasswordPBKDF2(password, iterations = PBKDF2_DEFAULT_ITERATIONS) {
  if (!password || typeof password !== 'string') {
    throw new Error('Password must be a non-empty string');
  }
  const validIterations = validatePbkdf2Iterations(iterations);

  const saltBytes = new Uint8Array(16);
  crypto.getRandomValues(saltBytes);
  const saltHex = Array.from(saltBytes).map(b => b.toString(16).padStart(2, '0')).join('');

  const enc = new TextEncoder();
  const passwordKey = await crypto.subtle.importKey(
    'raw',
    enc.encode(password),
    { name: 'PBKDF2' },
    false,
    ['deriveBits']
  );
  const derivedBits = await crypto.subtle.deriveBits(
    {
      name: 'PBKDF2',
      salt: saltBytes,
      iterations: validIterations,
      hash: 'SHA-256'
    },
    passwordKey,
    256 // 32 bytes
  );
  const hashHex = Array.from(new Uint8Array(derivedBits)).map(b => b.toString(16).padStart(2, '0')).join('');
  return `pbkdf2_sha256$${validIterations}$${saltHex}$${hashHex}`;
}

/**
 * Primary hash function: default to PBKDF2.
 */
export async function hashPassword(password, iterations = PBKDF2_DEFAULT_ITERATIONS) {
  return hashPasswordPBKDF2(password, iterations);
}

/**
 * Detailed password verification against multiple historical and modern formats:
 * 1. Sentinel strings ('$2a$10$placeholder_change_this', 'DISABLED_PENDING_BOOTSTRAP'): Rejected immediately.
 * 2. Modern PBKDF2 ('pbkdf2_sha256$<iter>$<salt>$<hash>'): Timing-safe PBKDF2 verification with strict [MIN, MAX] iteration bounds.
 * 3. Legacy SHA-256 (64 hex characters): Timing-safe SHA-256 verification + flags needsUpgrade = true.
 * 4. Malformed / unknown strings: Rejected immediately without throwing.
 */
export async function verifyPasswordDetailed(password, storedHash) {
  if (!password || typeof password !== 'string' || !storedHash || typeof storedHash !== 'string') {
    return { valid: false, needsUpgrade: false, reason: 'missing_input' };
  }

  // 1. Explicitly neutralize and reject sentinel strings
  if (storedHash === '$2a$10$placeholder_change_this' || storedHash === 'DISABLED_PENDING_BOOTSTRAP') {
    return { valid: false, needsUpgrade: false, reason: 'unbootstrapped' };
  }

  // 2. Modern PBKDF2-HMAC-SHA256
  if (storedHash.startsWith('pbkdf2_sha256$')) {
    const parts = storedHash.split('$');
    if (parts.length !== 4) {
      return { valid: false, needsUpgrade: false, reason: 'malformed_pbkdf2' };
    }
    const iterStr = parts[1];
    const saltHex = parts[2];
    const expectedHashHex = parts[3];

    if (!/^\d+$/.test(iterStr)) {
      return { valid: false, needsUpgrade: false, reason: 'malformed_pbkdf2_params' };
    }
    const iterations = Number(iterStr);

    if (
      !Number.isSafeInteger(iterations) ||
      iterations < PBKDF2_MIN_ITERATIONS ||
      iterations > PBKDF2_MAX_ITERATIONS ||
      !/^[a-f0-9]{32}$/i.test(saltHex) ||
      !/^[a-f0-9]{64}$/i.test(expectedHashHex)
    ) {
      return { valid: false, needsUpgrade: false, reason: 'malformed_pbkdf2_params' };
    }

    try {
      const saltBytes = new Uint8Array(saltHex.match(/.{1,2}/g).map(byte => parseInt(byte, 16)));
      const enc = new TextEncoder();
      const passwordKey = await crypto.subtle.importKey(
        'raw',
        enc.encode(password),
        { name: 'PBKDF2' },
        false,
        ['deriveBits']
      );
      const derivedBits = await crypto.subtle.deriveBits(
        {
          name: 'PBKDF2',
          salt: saltBytes,
          iterations: iterations,
          hash: 'SHA-256'
        },
        passwordKey,
        256
      );
      const computedHashHex = Array.from(new Uint8Array(derivedBits))
        .map(b => b.toString(16).padStart(2, '0'))
        .join('');

      const valid = constantTimeEqual(computedHashHex.toLowerCase(), expectedHashHex.toLowerCase());
      return { valid, needsUpgrade: false };
    } catch {
      return { valid: false, needsUpgrade: false, reason: 'crypto_error' };
    }
  }

  // 3. Legacy unsalted SHA-256 (64 hex characters)
  if (/^[a-f0-9]{64}$/i.test(storedHash)) {
    try {
      const legacyComputed = await hashPasswordLegacySHA256(password);
      const valid = constantTimeEqual(legacyComputed.toLowerCase(), storedHash.toLowerCase());
      return { valid, needsUpgrade: valid };
    } catch {
      return { valid: false, needsUpgrade: false, reason: 'crypto_error' };
    }
  }

  // 4. Any other format is invalid / corrupt
  return { valid: false, needsUpgrade: false, reason: 'unknown_format' };
}

/**
 * Standard password verification.
 * Returns boolean for direct compatibility with existing callers,
 * or full details object when options.detailed is true.
 */
export async function verifyPassword(password, storedHash, options = {}) {
  const details = await verifyPasswordDetailed(password, storedHash);
  if (options && options.detailed) {
    return details;
  }
  return details.valid;
}

export function generateSessionId() {
  const array = new Uint8Array(32);
  crypto.getRandomValues(array);
  return Array.from(array).map(b => b.toString(16).padStart(2, '0')).join('');
}

export function buildSessionExpiryIso(expiresInMs = NORMAL_TOKEN_TTL_MS) {
  const ttl = (typeof expiresInMs === 'number' && Number.isFinite(expiresInMs) && expiresInMs > 0)
    ? expiresInMs
    : NORMAL_TOKEN_TTL_MS;
  return new Date(Date.now() + ttl).toISOString();
}

function decodeBase64Segment(segment) {
  if (!segment || typeof segment !== 'string') return null;
  if (!/^[A-Za-z0-9+/_-]+={0,2}$/.test(segment)) return null;
  const normalized = segment.replace(/-/g, '+').replace(/_/g, '/');
  const rem = normalized.length % 4;
  if (rem === 1) return null;
  const padded = rem === 0 ? normalized : normalized + '='.repeat(4 - rem);
  try {
    return atob(padded);
  } catch {
    return null;
  }
}

const _hmacKeyCache = new Map();

async function getHmacKey(secret) {
  let cached = _hmacKeyCache.get(secret);
  if (cached) return cached;
  const encoder = new TextEncoder();
  const keyData = encoder.encode(secret);
  const key = await crypto.subtle.importKey(
    'raw', keyData, { name: 'HMAC', hash: 'SHA-256' }, false, ['sign', 'verify']
  );
  if (_hmacKeyCache.size > 16) _hmacKeyCache.clear();
  _hmacKeyCache.set(secret, key);
  return key;
}

export async function createToken(payload, secret, expiresInMs = null) {
  if (!secret || typeof secret !== 'string' || secret.trim() === '') {
    throw new Error('JWT_SECRET_REQUIRED: Valid JWT_SECRET is required to sign tokens.');
  }
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
    throw new Error('JWT_PAYLOAD_REQUIRED: Token payload must be a plain object.');
  }
  const now = Date.now();
  const tokenPayload = { ...payload, iat: now };
  const isRestricted = tokenPayload.mustChangePassword === true || tokenPayload.scope === 'restricted:change_password';
  const defaultTtl = isRestricted ? RESTRICTED_TOKEN_TTL_MS : NORMAL_TOKEN_TTL_MS;

  if (expiresInMs !== null && expiresInMs !== undefined) {
    if (typeof expiresInMs !== 'number' || !Number.isFinite(expiresInMs) || expiresInMs <= 0) {
      throw new Error('JWT_EXP_INVALID: expiresInMs must be a positive finite number.');
    }
    tokenPayload.exp = now + expiresInMs;
  } else if (typeof payload.exp === 'number' && Number.isFinite(payload.exp)) {
    tokenPayload.exp = payload.exp;
  } else {
    tokenPayload.exp = now + defaultTtl;
  }

  const header = btoa(JSON.stringify({ alg: 'HS256', typ: 'JWT' }));
  const body = btoa(JSON.stringify(tokenPayload));
  const data = `${header}.${body}`;
  const encoder = new TextEncoder();
  const messageData = encoder.encode(data);
  const key = await getHmacKey(secret);
  const signature = await crypto.subtle.sign('HMAC', key, messageData);
  const sigArray = Array.from(new Uint8Array(signature));
  const sig = btoa(String.fromCharCode(...sigArray));
  return `${data}.${sig}`;
}

export async function verifyToken(token, secret) {
  if (!token || typeof token !== 'string') return null;
  if (!secret || typeof secret !== 'string' || secret.trim() === '') {
    return null;
  }
  try {
    const parts = token.split('.');
    if (parts.length !== 3) return null;
    const [headerB64, bodyB64, sigB64] = parts;
    if (!headerB64 || !bodyB64 || !sigB64) return null;

    // Step 1: Parse and strictly validate JWT header algorithm and type
    const headerRaw = decodeBase64Segment(headerB64);
    if (!headerRaw) return null;
    const headerObj = JSON.parse(headerRaw);
    if (!headerObj || typeof headerObj !== 'object' || Array.isArray(headerObj)) return null;
    if (headerObj.alg !== 'HS256' || headerObj.typ !== 'JWT') return null;

    // Step 2: Verify HMAC-SHA256 signature BEFORE parsing/trusting payload claims
    const sigRaw = decodeBase64Segment(sigB64);
    if (!sigRaw) return null;
    const sigBytes = Uint8Array.from(sigRaw, c => c.charCodeAt(0));
    const data = `${headerB64}.${bodyB64}`;
    const encoder = new TextEncoder();
    const messageData = encoder.encode(data);
    const key = await getHmacKey(secret);
    const valid = await crypto.subtle.verify('HMAC', key, sigBytes, messageData);
    if (!valid) return null;

    // Step 3: Parse verified payload and enforce mandatory temporal claims (iat, exp, TTL)
    const bodyRaw = decodeBase64Segment(bodyB64);
    if (!bodyRaw) return null;
    const payload = JSON.parse(bodyRaw);
    if (!payload || typeof payload !== 'object' || Array.isArray(payload)) return null;

    if (typeof payload.iat !== 'number' || !Number.isFinite(payload.iat) || payload.iat <= 0) return null;
    if (typeof payload.exp !== 'number' || !Number.isFinite(payload.exp) || payload.exp <= payload.iat) return null;

    const now = Date.now();
    if (payload.iat > now + MAX_IAT_FUTURE_SKEW_MS) return null;
    if (now >= payload.exp) return null;

    const isRestricted = payload.mustChangePassword === true || payload.scope === 'restricted:change_password';
    const maxAllowedTtl = isRestricted ? RESTRICTED_TOKEN_TTL_MS : NORMAL_TOKEN_TTL_MS;
    if ((payload.exp - payload.iat) > maxAllowedTtl) return null;

    return payload;
  } catch {
    return null;
  }
}

/**
 * Strictly checks whether a database error specifically indicates that the
 * optional pre-Migration-0004 `assigned_modules` column does not exist yet.
 * No other error or column name matches.
 */
function isMissingAssignedModulesColumnError(err) {
  const msg = String(err?.message || err || '');
  return /\bno such column:\s*(?:[\w"`[\]]+\.)?assigned_modules\b/i.test(msg);
}

/**
 * Authenticates an incoming request against JWT signature + authoritative D1 database record.
 *
 * FAIL-CLOSED GUARANTEES:
 * 1. `must_change_password` is a mandatory security column. If the database schema does not
 *    have `must_change_password` (`no such column: must_change_password`), authentication
 *    is rejected (`return null`) rather than assuming `must_change_password = 0`.
 * 2. If `must_change_password` is present on the DB record, only an explicit `0` (or `'0'` / `false`)
 *    is treated as unrestricted (`0`). Any `null`, `NaN`, or non-zero value is treated as `1` (restricted).
 * 3. Only a recognized missing `assigned_modules` column error triggers the fallback query, and in that
 *    fallback `user.assigned_modules` is forced to `[]` (never falling back to JWT claims).
 * 4. Any other database error fails closed (`return null`).
 */
export function parseSessionExpiryMs(value) {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value !== 'string' || !value.trim()) return NaN;
  const trimmed = value.trim();
  const normalized = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}(\.\d+)?$/.test(trimmed)
    ? `${trimmed.replace(' ', 'T')}Z`
    : trimmed;
  return Date.parse(normalized);
}

export function isSessionRecordValid(userRow, expectedSessionId) {
  if (!userRow || !expectedSessionId || typeof expectedSessionId !== 'string') return false;
  if (!userRow._session_id || userRow._session_id !== expectedSessionId) return false;
  const expiresMs = parseSessionExpiryMs(userRow._session_expires_at);
  if (!Number.isFinite(expiresMs) || Date.now() >= expiresMs) return false;
  return true;
}

export async function fetchAuthUserWithSession(env, userId, sessionId) {
  let hasAssignedModulesColumn = true;
  let user = null;
  try {
    user = await env.DB.prepare(
      `SELECT u.id, u.username, u.email, u.full_name, u.role, u.is_active, u.must_change_password, u.assigned_modules,
              s.id AS _session_id, s.expires_at AS _session_expires_at
       FROM users u
       LEFT JOIN sessions s ON s.id = ? AND s.user_id = u.id
       WHERE u.id = ?`
    ).bind(sessionId, userId).first();
  } catch (err1) {
    if (isMissingAssignedModulesColumnError(err1)) {
      hasAssignedModulesColumn = false;
      user = await env.DB.prepare(
        `SELECT u.id, u.username, u.email, u.full_name, u.role, u.is_active, u.must_change_password,
                s.id AS _session_id, s.expires_at AS _session_expires_at
         FROM users u
         LEFT JOIN sessions s ON s.id = ? AND s.user_id = u.id
         WHERE u.id = ?`
      ).bind(sessionId, userId).first();
    } else {
      throw err1;
    }
  }
  return { user, hasAssignedModulesColumn };
}

export function buildRevokeSessionStatement(env, sessionId, userId) {
  return env.DB.prepare(
    'DELETE FROM sessions WHERE id = ? AND user_id = ?'
  ).bind(sessionId, userId);
}

export function buildRevokeAllUserSessionsStatement(env, userId) {
  return env.DB.prepare(
    'DELETE FROM sessions WHERE user_id = ?'
  ).bind(userId);
}

export function parseCookies(request) {
  const cookieHeader = request?.headers?.get?.('Cookie');
  const map = {};
  if (!cookieHeader || typeof cookieHeader !== 'string') return map;
  for (const part of cookieHeader.split(';')) {
    const eqIdx = part.indexOf('=');
    if (eqIdx <= 0) continue;
    const key = part.slice(0, eqIdx).trim();
    const rawVal = part.slice(eqIdx + 1).trim();
    if (!key) continue;
    try {
      map[key] = decodeURIComponent(rawVal);
    } catch {
      map[key] = rawVal;
    }
  }
  return map;
}

export async function authenticate(request, env) {
  if (request && typeof request === 'object' && '_fcmsAuthResult' in request) {
    return request._fcmsAuthResult;
  }
  const authHeader = request?.headers?.get?.('Authorization');
  let token = null;
  let authSource = null;

  if (authHeader && authHeader.startsWith('Bearer ')) {
    token = authHeader.slice(7);
    authSource = 'bearer';
  } else {
    const cookies = parseCookies(request);
    if (cookies[SESSION_COOKIE_NAME]) {
      token = cookies[SESSION_COOKIE_NAME];
      authSource = 'cookie';
    }
  }

  if (!token) return null;

  if (authSource === 'cookie') {
    const method = String(request?.method || 'GET').toUpperCase();
    if (['POST', 'PUT', 'PATCH', 'DELETE'].includes(method)) {
      const cookies = parseCookies(request);
      const cookieCsrf = cookies[CSRF_COOKIE_NAME] || '';
      const headerCsrf = request?.headers?.get?.('X-CSRF-Token') || '';
      if (!cookieCsrf || !headerCsrf || !constantTimeEqual(cookieCsrf, headerCsrf)) {
        return null;
      }
      const reqOrigin = request?.headers?.get?.('Origin');
      if (reqOrigin && !isAllowedOrigin(reqOrigin, env?.CORS_ORIGIN, env?.ENVIRONMENT)) {
        return null;
      }
    }
  }

  if (!env?.JWT_SECRET || typeof env.JWT_SECRET !== 'string' || env.JWT_SECRET.trim() === '') {
    console.error('AUTH_FATAL: JWT_SECRET environment variable is not configured.');
    return null;
  }
  const payload = (request && request._fcmsVerifiedPayload !== undefined)
    ? request._fcmsVerifiedPayload
    : await verifyToken(token, env.JWT_SECRET);
  const sessionId = (payload && typeof payload.sid === 'string') ? payload.sid.trim() : '';
  if (!payload || payload.sub === undefined || payload.sub === null || !sessionId) {
    if (request && typeof request === 'object' && request._fcmsAuthScope) request._fcmsAuthResult = null;
    return null;
  }
  if (!env?.DB) return null;

  let user = null;
  let hasAssignedModulesColumn = true;
  if (request && request._fcmsPreloadedDbUser !== undefined) {
    user = request._fcmsPreloadedDbUser ? { ...request._fcmsPreloadedDbUser } : null;
    hasAssignedModulesColumn = request._fcmsHasAssignedModulesCol !== false;
  } else {
    try {
      const fetched = await fetchAuthUserWithSession(env, payload.sub, sessionId);
      user = fetched.user;
      hasAssignedModulesColumn = fetched.hasAssignedModulesColumn;
    } catch (err) {
      console.error('AUTH_DB_ERROR: User and session lookup failed:', err?.message || err);
      return null;
    }
  }

  if (!user || Number(user.is_active) !== 1 || !isSessionRecordValid(user, sessionId)) {
    if (request && typeof request === 'object' && request._fcmsAuthScope) request._fcmsAuthResult = null;
    return null;
  }

  delete user._session_id;
  delete user._session_expires_at;
  user.sessionId = sessionId;

  // Fail-closed evaluation of must_change_password:
  // Never assume 0 if the DB column returned null, NaN, or a non-zero value.
  if (Object.prototype.hasOwnProperty.call(user, 'must_change_password')) {
    const isExplicitlyUnrestricted =
      user.must_change_password === 0 ||
      user.must_change_password === '0' ||
      user.must_change_password === false;
    user.must_change_password = isExplicitlyUnrestricted ? 0 : 1;
  } else {
    user.must_change_password = 0;
  }

  user.isRestrictedToken = !!(
    user.must_change_password === 1 || 
    payload.mustChangePassword === true || 
    payload.scope === 'restricted:change_password'
  );
  user.tokenPayload = payload;

  // Authoritative Database Sync (Fail-Closed):
  // Only trust assigned_modules from the database record, filtered against ASSIGNABLE_MODULES.
  // Never fall back to JWT payload.assigned_modules if the column is missing or unverified.
  if (hasAssignedModulesColumn && user.assigned_modules !== undefined) {
    user.assigned_modules = getAssignedModules(user.assigned_modules);
  } else {
    user.assigned_modules = [];
  }
  if (request && typeof request === 'object' && request._fcmsAuthScope) {
    request._fcmsAuthResult = user;
  }
  return user;
}

/**
 * Extracts, normalizes (hyphen -> underscore), deduplicates, and validates assigned modules
 * against the official ASSIGNABLE_MODULES allowlist.
 * Unrecognized or administrative module names are strictly ignored/filtered out.
 * Strictly ignores JWT tokenPayload claims.
 */
export function getAssignedModules(userOrRaw) {
  if (!userOrRaw) return [];
  let raw = userOrRaw;
  if (typeof userOrRaw === 'object' && !Array.isArray(userOrRaw)) {
    raw = userOrRaw.assigned_modules !== undefined 
      ? userOrRaw.assigned_modules 
      : (userOrRaw.assignedModules || userOrRaw.modules);
  }
  if (!raw) return [];

  const normalizeMod = (m) => String(m).trim().toLowerCase().replace(/-/g, '_');
  let list = [];

  if (Array.isArray(raw)) {
    list = raw.map(normalizeMod);
  } else if (typeof raw === 'string') {
    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) {
        list = parsed.map(normalizeMod);
      } else {
        list = raw.split(',').map(normalizeMod);
      }
    } catch {
      list = raw.split(',').map(normalizeMod);
    }
  }

  return Array.from(
    new Set(list.filter(mod => Boolean(mod) && ASSIGNABLE_MODULES_SET.has(mod)))
  );
}

/**
 * FCMS Official RBAC Permission Guard (Phase 5.66.19)
 * Enforces strictly the 2 official roles authorized by Owner policy:
 * 1. SUPERADMIN: Full access across entire system and modules (max 2 active).
 * 2. FC_SPECIALIST_LEADER: Read-only across operational & master modules (except admin modules);
 *    CREATE/UPDATE/WRITE allowed ONLY on 'overtime' and 'issues'.
 * Legacy roles (admin, supervisor, spv, editor_khusus, input_lembur, viewer, user) and
 * unknown roles receive ZERO permissions (fail-closed deny by default).
 */
export function hasPermission(user, moduleOrAction = 'read', action = null) {
  if (!user || typeof user !== 'object' || !user.role) return false;

  // Restricted tokens or accounts requiring password change have zero operational permissions
  if (user.isRestrictedToken === true || Number(user.must_change_password) === 1) {
    return false;
  }

  const role = String(user.role).trim().toLowerCase();

  // 1. SUPERADMIN: Full access across all modules and actions
  if (role === 'superadmin') {
    return true;
  }

  let module = null;
  let act = 'read';

  if (action !== null) {
    module = moduleOrAction;
    act = action;
  } else {
    // If only one argument was passed, determine if it's an action or module
    const raw = String(moduleOrAction || 'read').toLowerCase();
    if (['read', 'write', 'create', 'update', 'delete', 'export', 'import', 'admin'].includes(raw)) {
      act = raw;
      module = null;
    } else {
      module = raw;
      act = 'read';
    }
  }

  const cleanAct = String(act || 'read').trim().toLowerCase();
  const cleanModule = module ? String(module).trim().toLowerCase().replace(/-/g, '_') : '';

  // 2. FORBIDDEN FOR ALL NON-SUPERADMINS:
  // Administrative action and administrative system modules are strictly forbidden.
  if (cleanAct === 'admin') {
    return false;
  }
  if (cleanModule && (RESTRICTED_ADMIN_MODULES.has(cleanModule) || RESTRICTED_ADMIN_MODULES.has(String(module).trim().toLowerCase()))) {
    return false;
  }

  // 3. FC_SPECIALIST_LEADER: Only authorized non-superadmin official operational role
  // - Read-only across all operational & master modules (except restricted admin modules above)
  // - CREATE and UPDATE/WRITE allowed ONLY on overtime and issues
  // - Strictly NO delete, export, import, admin, or writes on other modules
  if (role === 'fc_specialist_leader') {
    if (cleanAct === 'read') return true;
    if (cleanAct === 'create' || cleanAct === 'update' || cleanAct === 'write') {
      return cleanModule === 'overtime' || cleanModule === 'issues';
    }
    return false;
  }

  // All legacy roles (admin, supervisor, spv, editor_khusus, input_lembur, viewer, user)
  // and unknown roles: deny by default (zero fallback access)
  return false;
}
