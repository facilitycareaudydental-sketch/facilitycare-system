import {
  hashPassword,
  verifyPassword,
  verifyPasswordDetailed,
  constantTimeEqual,
  generateSessionId,
  buildSessionExpiryIso,
  NORMAL_TOKEN_TTL_MS,
  RESTRICTED_TOKEN_TTL_MS,
  SESSION_COOKIE_NAME,
  CSRF_COOKIE_NAME,
  parseCookies,
  buildRevokeSessionStatement,
  buildRevokeAllUserSessionsStatement,
  createToken,
  authenticate,
  getAssignedModules,
  validatePbkdf2Iterations,
  validatePasswordPolicy
} from '../utils/auth.js';
import { ok, error, unauthorized, isAllowedOrigin, LOCAL_ALLOWED_ORIGINS } from '../utils/response.js';
import { logAudit, buildAuditStatement } from '../utils/audit.js';

// ── Rate Limiting (Best-Effort Per-Isolate Only — NOT Distributed) ──────────
// ARCHITECTURAL SECURITY & RISK NOTE:
// `rateLimitMap` and `bootstrapRateLimitMap` are stored in Cloudflare Worker
// isolate memory. Because Worker isolates are ephemeral and horizontally
// distributed across edge locations without shared memory (and without paid
// Durable Objects, KV writes, or paid WAF Rate Limiting), this mechanism only
// throttles repeated requests hitting the SAME warm isolate. It MUST NOT be
// claimed or relied upon as distributed brute-force protection.
//
// EVICTION & ACTIVE-COOLDOWN INTEGRITY POLICY:
// - Bounded at `MAX_RATE_LIMIT_ENTRIES` (1,000) to prevent isolate heap exhaustion.
// - When the Map reaches capacity:
//   1. Expired cooldown entries (`now >= cooldownUntil`) and stale non-cooldown
//      entries (`now - lastAttempt > cooldownMs`) are purged first.
//   2. If still at capacity, ONLY the oldest entry that is NOT in an active
//      cooldown (`!(cooldownUntil && now < cooldownUntil)`) is evicted.
//   3. Entries with an active cooldown (`cooldownUntil > now`) are NEVER evicted
//      early. If all 1,000 slots are occupied by active cooldowns, new unknown
//      requests fail closed with HTTP 429 until an active cooldown expires.
const rateLimitMap = new Map();
const bootstrapRateLimitMap = new Map();
const RATE_LIMIT_ATTEMPTS = 5;
const RATE_LIMIT_COOLDOWN_MS = 15 * 60 * 1000; // 15 minutes
const BOOTSTRAP_RATE_LIMIT_ATTEMPTS = 5;
const BOOTSTRAP_RATE_LIMIT_COOLDOWN_MS = 15 * 60 * 1000;
const MAX_RATE_LIMIT_ENTRIES = 1000;
const MIN_BOOTSTRAP_SECRET_BYTES = 32;
const MIN_BOOTSTRAP_SECRET_UNIQUE_CHARS = 10;

function normalizeRateLimitKey(rawKey) {
  return String(rawKey || '').trim().toLowerCase() || 'unknown';
}

function isEntryInActiveCooldown(entry, now) {
  return Boolean(entry && typeof entry.cooldownUntil === 'number' && now < entry.cooldownUntil);
}

/**
 * Prunes expired/stale entries when `map` reaches `MAX_RATE_LIMIT_ENTRIES`.
 * Never evicts an entry that is currently in an active cooldown (`now < cooldownUntil`).
 * Returns `true` if space is available (`map.size < MAX_RATE_LIMIT_ENTRIES`), or `false`
 * if all slots are currently occupied by active cooldowns.
 */
function pruneRateLimitMap(map, cooldownMs, now = Date.now()) {
  if (map.size < MAX_RATE_LIMIT_ENTRIES) return true;

  // Pass 1: Remove expired cooldowns and stale non-cooldown attempt records
  for (const [k, v] of map.entries()) {
    const expiredCooldown = typeof v?.cooldownUntil === 'number' && now >= v.cooldownUntil;
    const staleAttempt = !isEntryInActiveCooldown(v, now) && typeof v?.lastAttempt === 'number' && (now - v.lastAttempt > cooldownMs);
    if (expiredCooldown || staleAttempt) {
      map.delete(k);
    }
  }

  if (map.size < MAX_RATE_LIMIT_ENTRIES) return true;

  // Pass 2: Evict the oldest entry that is NOT in an active cooldown
  for (const [k, v] of map.entries()) {
    if (!isEntryInActiveCooldown(v, now)) {
      map.delete(k);
      break;
    }
  }

  return map.size < MAX_RATE_LIMIT_ENTRIES;
}

function checkRateLimit(username) {
  const key = normalizeRateLimitKey(username);
  const now = Date.now();
  const record = rateLimitMap.get(key);
  
  if (!record) {
    if (rateLimitMap.size >= MAX_RATE_LIMIT_ENTRIES) {
      const hasCapacity = pruneRateLimitMap(rateLimitMap, RATE_LIMIT_COOLDOWN_MS, now);
      if (!hasCapacity) {
        return { allowed: false, remaining: 0, retryAfter: 60 };
      }
    }
    return { allowed: true, remaining: RATE_LIMIT_ATTEMPTS };
  }
  
  // Cooldown expired, reset
  if (record.cooldownUntil && now >= record.cooldownUntil) {
    rateLimitMap.delete(key);
    return { allowed: true, remaining: RATE_LIMIT_ATTEMPTS };
  }
  
  // Still in cooldown period
  if (record.cooldownUntil && now < record.cooldownUntil) {
    return { allowed: false, remaining: 0, retryAfter: Math.ceil((record.cooldownUntil - now) / 1000) };
  }
  
  // Reset attempts if 15 minutes passed since last attempt
  if (now - record.lastAttempt > RATE_LIMIT_COOLDOWN_MS) {
    record.attempts = 0;
  }
  
  if (record.attempts >= RATE_LIMIT_ATTEMPTS) {
    return { allowed: false, remaining: 0, retryAfter: Math.ceil((record.lastAttempt + RATE_LIMIT_COOLDOWN_MS - now) / 1000) };
  }
  
  return { allowed: true, remaining: RATE_LIMIT_ATTEMPTS - record.attempts - 1 };
}

function recordFailedLogin(username) {
  const key = normalizeRateLimitKey(username);
  const now = Date.now();
  const existing = rateLimitMap.get(key);

  if (!existing && rateLimitMap.size >= MAX_RATE_LIMIT_ENTRIES) {
    const hasCapacity = pruneRateLimitMap(rateLimitMap, RATE_LIMIT_COOLDOWN_MS, now);
    if (!hasCapacity) {
      return;
    }
  }

  const record = existing || { attempts: 0, lastAttempt: now };
  record.attempts = (record.attempts || 0) + 1;
  record.lastAttempt = now;
  
  if (record.attempts >= RATE_LIMIT_ATTEMPTS) {
    record.cooldownUntil = now + RATE_LIMIT_COOLDOWN_MS;
  }
  
  rateLimitMap.set(key, record);
}

function recordSuccessfulLogin(username) {
  const key = normalizeRateLimitKey(username);
  rateLimitMap.delete(key);
}

function checkBootstrapRateLimit(identifier) {
  const key = normalizeRateLimitKey(identifier);
  const now = Date.now();
  const record = bootstrapRateLimitMap.get(key);
  if (!record) {
    if (bootstrapRateLimitMap.size >= MAX_RATE_LIMIT_ENTRIES) {
      const hasCapacity = pruneRateLimitMap(bootstrapRateLimitMap, BOOTSTRAP_RATE_LIMIT_COOLDOWN_MS, now);
      if (!hasCapacity) {
        return { allowed: false, remaining: 0, retryAfter: 60 };
      }
    }
    return { allowed: true, remaining: BOOTSTRAP_RATE_LIMIT_ATTEMPTS };
  }

  if (record.cooldownUntil && now >= record.cooldownUntil) {
    bootstrapRateLimitMap.delete(key);
    return { allowed: true, remaining: BOOTSTRAP_RATE_LIMIT_ATTEMPTS };
  }

  if (record.cooldownUntil && now < record.cooldownUntil) {
    return { allowed: false, remaining: 0, retryAfter: Math.ceil((record.cooldownUntil - now) / 1000) };
  }

  if (now - record.lastAttempt > BOOTSTRAP_RATE_LIMIT_COOLDOWN_MS) {
    record.attempts = 0;
  }

  if (record.attempts >= BOOTSTRAP_RATE_LIMIT_ATTEMPTS) {
    return { allowed: false, remaining: 0, retryAfter: Math.ceil((record.lastAttempt + BOOTSTRAP_RATE_LIMIT_COOLDOWN_MS - now) / 1000) };
  }

  return { allowed: true, remaining: BOOTSTRAP_RATE_LIMIT_ATTEMPTS - record.attempts - 1 };
}

function recordFailedBootstrap(identifier) {
  const key = normalizeRateLimitKey(identifier);
  const now = Date.now();
  const existing = bootstrapRateLimitMap.get(key);

  if (!existing && bootstrapRateLimitMap.size >= MAX_RATE_LIMIT_ENTRIES) {
    const hasCapacity = pruneRateLimitMap(bootstrapRateLimitMap, BOOTSTRAP_RATE_LIMIT_COOLDOWN_MS, now);
    if (!hasCapacity) {
      return;
    }
  }

  const record = existing || { attempts: 0, lastAttempt: now };
  record.attempts = (record.attempts || 0) + 1;
  record.lastAttempt = now;

  if (record.attempts >= BOOTSTRAP_RATE_LIMIT_ATTEMPTS) {
    record.cooldownUntil = now + BOOTSTRAP_RATE_LIMIT_COOLDOWN_MS;
  }

  bootstrapRateLimitMap.set(key, record);
}

/**
 * Validates that a D1/SQLite UPDATE result contains a genuine positive integer
 * row-change count (`meta.changes >= 1`). Prevents the JavaScript `Number(undefined) < 1 === false`
 * (`NaN < 1 === false`) pitfall.
 */
function hasValidUpdatedRows(updateResult) {
  const rawChanges = updateResult?.meta?.changes;
  return typeof rawChanges === 'number' && Number.isInteger(rawChanges) && rawChanges >= 1;
}

/**
 * Validates `SUPERADMIN_BOOTSTRAP_TOKEN` strength.
 *
 * CRYPTOGRAPHIC GENERATION REQUIREMENT:
 * `SUPERADMIN_BOOTSTRAP_TOKEN` MUST be generated using a Cryptographically Secure
 * Pseudo-Random Number Generator (CSPRNG) with at least 256 bits of entropy
 * (e.g., `openssl rand -hex 32` or `crypto.getRandomValues(new Uint8Array(32))`),
 * NOT a human-typed phrase or repeated padding string that merely meets minimum length.
 * Never log or print the secret or client-supplied token.
 */
function isStrongBootstrapSecret(secret) {
  if (!secret || typeof secret !== 'string') return false;
  const trimmed = secret.trim();
  const byteLength = new TextEncoder().encode(trimmed).byteLength;
  if (byteLength < MIN_BOOTSTRAP_SECRET_BYTES) return false;
  const uniqueChars = new Set(trimmed).size;
  if (uniqueChars < MIN_BOOTSTRAP_SECRET_UNIQUE_CHARS) return false;
  return true;
}

// Exported strictly for local deterministic unit testing of rate-limit Map boundaries
export const __rateLimitTestUtils = {
  rateLimitMap,
  bootstrapRateLimitMap,
  MAX_RATE_LIMIT_ENTRIES,
  RATE_LIMIT_ATTEMPTS,
  RATE_LIMIT_COOLDOWN_MS,
  BOOTSTRAP_RATE_LIMIT_ATTEMPTS,
  BOOTSTRAP_RATE_LIMIT_COOLDOWN_MS,
  checkRateLimit,
  recordFailedLogin,
  recordSuccessfulLogin,
  checkBootstrapRateLimit,
  recordFailedBootstrap,
  pruneRateLimitMap
};

export async function handleAuth(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace('/api/auth', '');

  // POST /api/auth/bootstrap
  if (path === '/bootstrap' && request.method === 'POST') {
    return handleBootstrap(request, env, origin);
  }

  // POST /api/auth/login
  if (path === '/login' && request.method === 'POST') {
    return handleLogin(request, env, origin);
  }

  // POST /api/auth/logout
  if (path === '/logout' && request.method === 'POST') {
    return handleLogout(request, env, origin);
  }

  // GET /api/auth/me
  if (path === '/me' && request.method === 'GET') {
    return handleMe(request, env, origin);
  }

  // POST /api/auth/change-password
  if (path === '/change-password' && request.method === 'POST') {
    return handleChangePassword(request, env, origin);
  }

  return error('Not found', 404, origin);
}

async function handleBootstrap(request, env, origin) {
  // 1. Fail-closed check: SUPERADMIN_BOOTSTRAP_TOKEN must be CSPRNG-generated (>= 32 bytes & sufficient entropy).
  // Never log the secret value or client token.
  const bootstrapSecret = env?.SUPERADMIN_BOOTSTRAP_TOKEN;
  if (!isStrongBootstrapSecret(bootstrapSecret)) {
    console.error('AUTH_FATAL: SUPERADMIN_BOOTSTRAP_TOKEN is missing or does not meet CSPRNG length/entropy requirements.');
    return error('Bootstrap service unavailable', 503, origin);
  }

  // 2. Best-effort in-isolate rate limiting check (global multi-isolate enforcement handled at Cloudflare WAF edge)
  const isDeployedEnv = env?.ENVIRONMENT === 'production' || env?.ENVIRONMENT === 'staging';
  const cfIp = request.headers?.get?.('CF-Connecting-IP');
  const clientIp = cfIp || (isDeployedEnv ? 'untrusted-edge' : (request.headers?.get?.('X-Forwarded-For')?.split(',')[0]?.trim() || 'client'));
  const rateKey = `bootstrap_${clientIp}`;
  const rateLimit = checkBootstrapRateLimit(rateKey);
  if (!rateLimit.allowed) {
    return error(`Too many failed bootstrap attempts. Please try again in ${rateLimit.retryAfter} seconds.`, 429, origin);
  }

  // 3. Request parsing & input validation
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!body || typeof body !== 'object') {
    return error('Invalid JSON payload', 400, origin);
  }
  const { username, bootstrapToken, newPassword } = body;
  if (
    typeof username !== 'string' || !username ||
    typeof bootstrapToken !== 'string' || !bootstrapToken ||
    newPassword === undefined || newPassword === null || newPassword === ''
  ) {
    return error('username, bootstrapToken, and newPassword are required', 400, origin);
  }

  const normUsername = String(username).trim().toLowerCase();
  if (normUsername !== 'superadmin') {
    recordFailedBootstrap(rateKey);
    return error('Bootstrap is only permissible for the superadmin account', 403, origin);
  }

  // 4. Timing-safe verification of bootstrap token (never log bootstrapToken or bootstrapSecret)
  const tokenValid = constantTimeEqual(String(bootstrapToken), bootstrapSecret.trim());
  if (!tokenValid) {
    recordFailedBootstrap(rateKey);
    return error('Invalid bootstrap token', 401, origin);
  }

  // 5. Centralized password complexity validation (type, min 12 chars, upper, lower, number, special char)
  const pwValidation = validatePasswordPolicy(newPassword);
  if (!pwValidation.valid) {
    return error(pwValidation.error, 400, origin);
  }

  // 6. Validate PBKDF2_ITERATIONS centrally before hashing to prevent CPU overload
  let iterations;
  try {
    iterations = validatePbkdf2Iterations(env?.PBKDF2_ITERATIONS);
  } catch (iterErr) {
    console.error('AUTH_CONFIG_ERROR:', iterErr.message);
    return error('Authentication configuration error: invalid PBKDF2_ITERATIONS', 500, origin);
  }

  const newHash = await hashPassword(newPassword, iterations);

  // 7. Atomic UPDATE: Only update if superadmin is inactive and in sentinel pending state
  let updateResult;
  try {
    updateResult = await env.DB.prepare(`
      UPDATE users 
      SET 
        password_hash = ?, 
        is_active = 1, 
        must_change_password = 0, 
        activated_at = datetime('now'),
        updated_at = datetime('now')
      WHERE 
        LOWER(TRIM(username)) = ? 
        AND LOWER(role) = 'superadmin'
        AND password_hash = 'DISABLED_PENDING_BOOTSTRAP'
        AND is_active = 0
        AND (SELECT COUNT(*) FROM users WHERE LOWER(role) = 'superadmin' AND is_active = 1) = 0
    `).bind(newHash, normUsername).run();
  } catch (err) {
    const msg = String(err?.message || err || '');
    if (msg.includes('idx_users_single_active_superadmin') || msg.includes('UNIQUE constraint failed')) {
      return error('Account already activated or ineligible for bootstrap', 409, origin);
    }
    console.error('AUTH_DB_ERROR: Bootstrap database update failed.');
    return error('Bootstrap service temporarily unavailable', 503, origin);
  }

  if (!hasValidUpdatedRows(updateResult)) {
    // Account was already bootstrapped or is ineligible
    return error('Account already activated or ineligible for bootstrap', 409, origin);
  }

  // 8. Successful activation: reset rate limit and log audit without leaking secrets
  bootstrapRateLimitMap.delete(normalizeRateLimitKey(rateKey));
  try {
    await logAudit(env, { id: 'bootstrap', username: normUsername, role: 'superadmin' }, 'CREATE', 'auth', normUsername, null, { action: 'Superadmin account initialized via one-time bootstrap' });
  } catch (auditErr) {
    console.error('Non-critical bootstrap audit log failed:', auditErr?.message || auditErr);
  }

  // 9. Return success confirmation only (No normal JWT issued!)
  return ok({
    message: 'Superadmin account successfully activated. Please log in with your credentials.'
  }, 200, origin);
}

async function recordD1FailedLogin(env, identifier, ip, acctIdentifier = null) {
  if (!env?.DB) return;
  const upsertSql = `
    INSERT INTO auth_cooldowns (identifier, failed_count, cooldown_until, last_failed_at, ip_address, created_at, updated_at)
    VALUES (?, 1, NULL, datetime('now'), ?, datetime('now'), datetime('now'))
    ON CONFLICT(identifier) DO UPDATE SET
      failed_count = auth_cooldowns.failed_count + 1,
      cooldown_until = CASE WHEN auth_cooldowns.failed_count + 1 >= 5 THEN datetime('now', '+15 minutes') ELSE NULL END,
      last_failed_at = datetime('now'),
      updated_at = datetime('now')
  `;
  try {
    await env.DB.prepare(upsertSql).bind(identifier, ip).run();
    if (acctIdentifier && acctIdentifier !== identifier) {
      await env.DB.prepare(upsertSql).bind(acctIdentifier, ip).run();
    }
  } catch (e) {
    console.warn('Non-critical D1 failed login recording skipped:', e?.message || e);
  }
}

async function recordD1SuccessfulLogin(env, identifier, acctIdentifier = null) {
  if (!env?.DB) return;
  try {
    await env.DB.prepare('DELETE FROM auth_cooldowns WHERE identifier = ?').bind(identifier).run();
    if (acctIdentifier && acctIdentifier !== identifier) {
      await env.DB.prepare('DELETE FROM auth_cooldowns WHERE identifier = ?').bind(acctIdentifier).run();
    }
  } catch (e) {
    console.warn('Non-critical D1 successful login reset skipped:', e?.message || e);
  }
}

async function handleLogin(request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!body || typeof body !== 'object') {
    return error('Invalid JSON payload', 400, origin);
  }

  const { username, password } = body;
  if (typeof username !== 'string' || typeof password !== 'string' || !username.trim() || !password) {
    return error('Username and password required', 400, origin);
  }

  // Consistent normalization for login identity: trim and lowercase
  // Password remains strictly case-sensitive and unnormalized
  const normalizedIdentity = username.trim().toLowerCase();
  const clientIp = request?.headers?.get?.('cf-connecting-ip') || '127.0.0.1';
  const d1Identifier = `${normalizedIdentity}:${clientIp}`;
  const d1AcctIdentifier = `acct:${normalizedIdentity}`;

  // ── Tier 2: D1 Persistent Lockout Check (Global Persistence across Isolate Restarts & IP Rotation) ──
  if (env?.DB) {
    try {
      const lockout = await env.DB.prepare(
        "SELECT failed_count, cooldown_until FROM auth_cooldowns WHERE (identifier = ? OR identifier = ?) AND cooldown_until > datetime('now') LIMIT 1"
      ).bind(d1Identifier, d1AcctIdentifier).first();
      if (lockout) {
        return error('Terlalu banyak percobaan login gagal. Akun dikunci sementara selama 15 menit.', 429, origin, env?.ENVIRONMENT, { 'Retry-After': '900' });
      }
    } catch (e) {
      console.warn('Non-critical auth_cooldowns check skipped:', e?.message || e);
    }
  }

  // ── Tier 1: In-Isolate Memory Rate Limit Check ────────────────────────────
  const rateLimit = checkRateLimit(normalizedIdentity);
  if (!rateLimit.allowed) {
    return error(`Too many failed attempts. Please try again in ${rateLimit.retryAfter} seconds.`, 429, origin, env?.ENVIRONMENT, { 'Retry-After': String(rateLimit.retryAfter) });
  }

  let user;
  try {
    user = await env.DB.prepare(
      'SELECT * FROM users WHERE (LOWER(TRIM(username)) = ? OR LOWER(TRIM(email)) = ?) AND is_active = 1'
    ).bind(normalizedIdentity, normalizedIdentity).first();
  } catch {
    console.error('AUTH_DB_ERROR: Login user lookup failed.');
    return error('Authentication service temporarily unavailable', 503, origin);
  }

  if (!user) {
    recordFailedLogin(normalizedIdentity);
    await recordD1FailedLogin(env, d1Identifier, clientIp, d1AcctIdentifier);
    try {
      await logAudit(
        env,
        { id: null, username: normalizedIdentity, role: null },
        'LOGIN',
        'auth',
        normalizedIdentity,
        null,
        null,
        { reason: 'invalid_credentials_or_inactive' },
        'FAILED'
      );
    } catch (auditErr) {
      console.warn('Non-critical failed login audit skipped:', auditErr?.message || auditErr);
    }
    return unauthorized(origin);
  }

  // ── Password Verification (PBKDF2 / Legacy SHA-256 / Sentinel check) ───────
  const verifyDetails = await verifyPasswordDetailed(password, user.password_hash);
  if (!verifyDetails.valid) {
    recordFailedLogin(normalizedIdentity);
    await recordD1FailedLogin(env, d1Identifier, clientIp, d1AcctIdentifier);
    try {
      await logAudit(
        env,
        { id: user.id, username: user.username, full_name: user.full_name, role: user.role },
        'LOGIN',
        'auth',
        user.id,
        null,
        null,
        { reason: 'invalid_password' },
        'FAILED'
      );
    } catch (auditErr) {
      console.warn('Non-critical failed login audit skipped:', auditErr?.message || auditErr);
    }
    return unauthorized(origin);
  }

  // ── Transparent Hash Upgrade for Legacy SHA-256 ───────────────────────────
  if (verifyDetails.needsUpgrade) {
    try {
      const iterations = validatePbkdf2Iterations(env?.PBKDF2_ITERATIONS);
      const upgradedHash = await hashPassword(password, iterations);
      // Atomic conditional update to prevent race conditions during concurrent logins
      await env.DB.prepare(
        "UPDATE users SET password_hash = ?, updated_at = datetime('now') WHERE id = ? AND password_hash = ?"
      ).bind(upgradedHash, user.id, user.password_hash).run();
    } catch (upgradeErr) {
      console.error('Non-critical password hash upgrade skipped:', upgradeErr?.message || upgradeErr);
    }
  }

  // ── Successful login - reset rate limit ────────────────────────────────────
  recordSuccessfulLogin(normalizedIdentity);
  await recordD1SuccessfulLogin(env, d1Identifier, d1AcctIdentifier);

  if (!env?.JWT_SECRET || typeof env.JWT_SECRET !== 'string' || env.JWT_SECRET.trim() === '') {
    console.error('AUTH_FATAL: JWT_SECRET environment variable is not configured.');
    return error('Authentication service misconfigured', 500, origin);
  }

  const isMustChangePassword = Number(user.must_change_password || 0) === 1;
  const assignedModules = getAssignedModules(user);
  const ttlMs = isMustChangePassword ? RESTRICTED_TOKEN_TTL_MS : NORMAL_TOKEN_TTL_MS;
  const sessionId = generateSessionId();
  const expiresAtIso = buildSessionExpiryIso(ttlMs);

  // Persist stateful session BEFORE minting JWT (fail-closed if session storage fails)
  try {
    await env.DB.prepare(
      'INSERT INTO sessions (id, user_id, expires_at) VALUES (?, ?, ?)'
    ).bind(sessionId, user.id, expiresAtIso).run();
  } catch (sessErr) {
    console.error('AUTH_DB_ERROR: Failed to persist login session:', sessErr?.message || sessErr);
    return error('Authentication service temporarily unavailable', 503, origin);
  }

  // Non-critical cleanup of already-expired sessions for this user
  try {
    await env.DB.prepare(
      "DELETE FROM sessions WHERE user_id = ? AND datetime(expires_at) <= datetime('now')"
    ).bind(user.id).run();
  } catch (cleanupErr) {
    console.warn('Non-critical expired session cleanup failed:', cleanupErr?.message || cleanupErr);
  }

  const tokenPayload = {
    sub: user.id,
    username: user.username,
    role: user.role,
    assigned_modules: assignedModules,
    sid: sessionId
  };

  if (isMustChangePassword) {
    tokenPayload.scope = 'restricted:change_password';
    tokenPayload.mustChangePassword = true;
  }

  const secret = env.JWT_SECRET;
  const token = await createToken(tokenPayload, secret, ttlMs);

  const userObj = {
    id: user.id,
    username: user.username,
    email: user.email,
    full_name: user.full_name,
    role: user.role,
    assigned_modules: assignedModules,
    must_change_password: isMustChangePassword ? 1 : 0
  };

  // AVAILABILITY & AUDIT CONSISTENCY POLICY:
  // `updated_at` and `logAudit` are executed in separate try/catch blocks so a
  // failure in updating `users.updated_at` does not skip `audit_logs` insertion.
  // Login is NOT rejected if D1 write quota is exhausted on non-critical post-login
  // writes, preventing total system lockout of Superadmin during D1 write incidents.
  // Zero passwords, hashes, tokens, or secrets are passed to logs or audit payloads.
  try {
    await env.DB.prepare("UPDATE users SET updated_at = datetime('now') WHERE id = ?").bind(user.id).run();
  } catch (err) {
    console.error('Non-critical login updated_at write failed:', err?.message || err);
  }

  try {
    await logAudit(env, userObj, 'LOGIN', 'auth', user.id, null, { action: 'User logged in' });
  } catch (auditErr) {
    console.error('AUDIT_WARNING: Login audit log write failed:', auditErr?.message || auditErr);
  }

  const csrfToken = generateSessionId();
  const res = ok({
    token,
    user: userObj,
    must_change_password: isMustChangePassword,
    csrf_token: csrfToken
  }, 200, origin);

  const isHttps = /^https:/i.test(String(request?.url || '')) || /^https:/i.test(String(origin || ''));
  const sameSiteAttr = isHttps ? 'Secure; SameSite=None' : 'SameSite=Lax';
  const maxAgeSec = Math.max(60, Math.floor(ttlMs / 1000));
  res.headers.append(
    'Set-Cookie',
    `fcms_session=${encodeURIComponent(token)}; HttpOnly; Path=/; Max-Age=${maxAgeSec}; ${sameSiteAttr}`
  );
  res.headers.append(
    'Set-Cookie',
    `fcms_csrf=${encodeURIComponent(csrfToken)}; Path=/; Max-Age=${maxAgeSec}; ${sameSiteAttr}`
  );
  return res;
}

/**
 * POST /api/auth/logout
 * Revokes the caller's active session (or all sessions for that user when `{ all_sessions: true }`
 * or `?all=true` is supplied) atomically with the LOGOUT audit event. Fails closed on DB error.
 */
async function handleLogout(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user || !user.sessionId) {
    return unauthorized(origin);
  }

  let revokeAll = false;
  try {
    const reqUrl = new URL(request.url);
    if (reqUrl.searchParams.get('all') === 'true' || reqUrl.searchParams.get('all_sessions') === 'true') {
      revokeAll = true;
    }
  } catch {}
  if (!revokeAll) {
    try {
      const contentType = request.headers.get('Content-Type') || '';
      if (contentType.includes('application/json')) {
        const body = await request.clone().json();
        if (body && typeof body === 'object' && body.all_sessions === true) {
          revokeAll = true;
        }
      }
    } catch {}
  }

  const revokeStmt = revokeAll
    ? buildRevokeAllUserSessionsStatement(env, user.id)
    : buildRevokeSessionStatement(env, user.sessionId, user.id);
  const auditStmt = buildAuditStatement(
    env,
    user,
    'LOGOUT',
    'auth',
    user.id,
    null,
    { action: revokeAll ? 'User logged out (all sessions)' : 'User logged out' },
    { scope: revokeAll ? 'all_sessions' : 'current_session' },
    'SUCCESS',
    true
  );

  try {
    await env.DB.batch([revokeStmt, auditStmt]);
  } catch (err) {
    console.error('AUTH_DB_ERROR: Session revocation on logout failed:', err?.message || err);
    return error('Failed to revoke active session', 503, origin);
  }

  const res = ok({
    message: 'Logged out',
    revoked: true,
    scope: revokeAll ? 'all_sessions' : 'current_session'
  }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });

  const isHttps = /^https:/i.test(String(request?.url || '')) || /^https:/i.test(String(origin || ''));
  const sameSiteAttr = isHttps ? 'Secure; SameSite=None' : 'SameSite=Lax';
  res.headers.append('Set-Cookie', `fcms_session=; HttpOnly; Path=/; Max-Age=0; Expires=Thu, 01 Jan 1970 00:00:00 GMT; ${sameSiteAttr}`);
  res.headers.append('Set-Cookie', `fcms_csrf=; Path=/; Max-Age=0; Expires=Thu, 01 Jan 1970 00:00:00 GMT; ${sameSiteAttr}`);
  return res;
}

async function deriveSessionCsrfToken(jwtSecret, sessionId) {
  const encoder = new TextEncoder();
  const key = await crypto.subtle.importKey(
    'raw',
    encoder.encode(jwtSecret),
    { name: 'HMAC', hash: 'SHA-256' },
    false,
    ['sign']
  );
  const sig = await crypto.subtle.sign('HMAC', key, encoder.encode(`fcms-csrf:${sessionId}`));
  return Array.from(new Uint8Array(sig)).map((b) => b.toString(16).padStart(2, '0')).join('');
}

async function handleMe(request, env, origin) {
  const reqOrigin = request?.headers?.get?.('Origin');
  if (reqOrigin && !isAllowedOrigin(reqOrigin, env?.CORS_ORIGIN, env?.ENVIRONMENT)) {
    return error('Origin not allowed by CORS policy', 403, null, env?.ENVIRONMENT);
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);

  let isLocalSameOrigin = false;
  try {
    const reqHostOrigin = new URL(request.url).origin;
    isLocalSameOrigin = LOCAL_ALLOWED_ORIGINS.has(reqHostOrigin);
  } catch {}

  const cookies = parseCookies(request);
  const hasSessionCookie = Boolean(cookies[SESSION_COOKIE_NAME]);
  const isOriginPermittedForCsrf = Boolean(
    (reqOrigin && isAllowedOrigin(reqOrigin, env?.CORS_ORIGIN, env?.ENVIRONMENT)) ||
    (!reqOrigin && isLocalSameOrigin)
  );

  let csrfToken = null;
  let shouldSetCsrfCookie = false;

  if (isOriginPermittedForCsrf && (hasSessionCookie || reqOrigin)) {
    const existingCsrf = typeof cookies[CSRF_COOKIE_NAME] === 'string' ? cookies[CSRF_COOKIE_NAME].trim() : '';
    if (existingCsrf.length >= 16) {
      // Preserve active CSRF cookie so concurrent reloads and multiple tabs stay in sync
      csrfToken = existingCsrf;
    } else if (env?.JWT_SECRET && user.sessionId) {
      csrfToken = await deriveSessionCsrfToken(env.JWT_SECRET, user.sessionId);
      shouldSetCsrfCookie = hasSessionCookie;
    }
  }

  const responseData = csrfToken ? { ...user, csrf_token: csrfToken } : user;
  const res = ok(responseData, 200, origin, env?.ENVIRONMENT, { 'Cache-Control': 'no-store' });

  if (shouldSetCsrfCookie && csrfToken) {
    const isHttps = /^https:/i.test(String(request?.url || '')) || /^https:/i.test(String(origin || ''));
    const sameSiteAttr = isHttps ? 'Secure; SameSite=None' : 'SameSite=Lax';
    const remainingSec = user?.tokenPayload?.exp
      ? Math.max(60, Math.floor((user.tokenPayload.exp - Date.now()) / 1000))
      : Math.floor(NORMAL_TOKEN_TTL_MS / 1000);
    res.headers.append(
      'Set-Cookie',
      `${CSRF_COOKIE_NAME}=${encodeURIComponent(csrfToken)}; Path=/; Max-Age=${remainingSec}; ${sameSiteAttr}`
    );
  }

  return res;
}

async function handleChangePassword(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);

  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!body || typeof body !== 'object') {
    return error('Invalid JSON payload', 400, origin);
  }

  const { current_password, new_password } = body;
  if (current_password === undefined || current_password === null || current_password === '' ||
      new_password === undefined || new_password === null || new_password === '') {
    return error('Both passwords required', 400, origin);
  }
  if (typeof current_password !== 'string' || typeof new_password !== 'string') {
    return error('current_password and new_password must be strings', 400, origin);
  }

  // Enforce centralized password policy (min 12 chars + upper + lower + number + special char)
  const pwValidation = validatePasswordPolicy(new_password);
  if (!pwValidation.valid) {
    return error(pwValidation.error, 400, origin);
  }

  // Validate PBKDF2_ITERATIONS centrally before touching DB or running crypto
  let iterations;
  try {
    iterations = validatePbkdf2Iterations(env?.PBKDF2_ITERATIONS);
  } catch (iterErr) {
    console.error('AUTH_CONFIG_ERROR:', iterErr.message);
    return error('Authentication configuration error: invalid PBKDF2_ITERATIONS', 500, origin);
  }

  // Ensure user record and password_hash exist in DB before reading password_hash.
  // Handle DB SELECT failure in a controlled manner without exposing internal DB errors.
  let dbUser;
  try {
    dbUser = await env.DB.prepare('SELECT password_hash FROM users WHERE id = ?').bind(user.id).first();
  } catch {
    console.error('AUTH_DB_ERROR: Change-password user lookup failed.');
    return error('Authentication service temporarily unavailable', 503, origin);
  }

  if (!dbUser || typeof dbUser.password_hash !== 'string' || !dbUser.password_hash) {
    return error('User not found', 404, origin);
  }

  const valid = await verifyPassword(current_password, dbUser.password_hash);
  if (!valid) return error('Current password incorrect', 400, origin);

  const newHash = await hashPassword(new_password, iterations);

  // Conditional atomic update + audit log + revoke ALL sessions for this user in one D1 batch.
  // Order matters: auditStmt (requireChanges = true) MUST immediately follow updateStmt so
  // changes() reflects the password update rather than a 0-row session cleanup.
  const updateStmt = env.DB.prepare(
    "UPDATE users SET password_hash = ?, must_change_password = 0, updated_at = datetime('now') WHERE id = ? AND password_hash = ?"
  ).bind(newHash, user.id, dbUser.password_hash);
  const auditStmt = buildAuditStatement(
    env,
    { ...user, must_change_password: 0, isRestrictedToken: false },
    'UPDATE',
    'auth',
    user.id,
    null,
    { action: 'User changed password' },
    { sessions_revoked: true },
    'SUCCESS',
    true
  );
  const revokeAllStmt = buildRevokeAllUserSessionsStatement(env, user.id);

  let batchRes;
  try {
    batchRes = await env.DB.batch([updateStmt, auditStmt, revokeAllStmt]);
  } catch {
    console.error('AUTH_DB_ERROR: Change-password atomic batch update failed.');
    return error('Failed to update password', 500, origin);
  }

  if (!hasValidUpdatedRows(batchRes?.[0])) {
    return error('Failed to update password. Account state may have changed.', 409, origin);
  }

  // Only clear in-memory restriction flags after DB confirms password update succeeded with >= 1 row changed
  user.must_change_password = 0;
  user.isRestrictedToken = false;

  const res = ok(
    { message: 'Password changed successfully', sessions_revoked: true },
    200,
    origin,
    undefined,
    { 'X-Audit-Batched': 'true' }
  );
  const isHttps = /^https:/i.test(String(request?.url || '')) || /^https:/i.test(String(origin || ''));
  const sameSiteAttr = isHttps ? 'Secure; SameSite=None' : 'SameSite=Lax';
  res.headers.append('Set-Cookie', `fcms_session=; HttpOnly; Path=/; Max-Age=0; Expires=Thu, 01 Jan 1970 00:00:00 GMT; ${sameSiteAttr}`);
  res.headers.append('Set-Cookie', `fcms_csrf=; Path=/; Max-Age=0; Expires=Thu, 01 Jan 1970 00:00:00 GMT; ${sameSiteAttr}`);
  return res;
}
