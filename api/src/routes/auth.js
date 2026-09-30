import { hashPassword, verifyPassword, createToken, authenticate } from '../utils/auth.js';
import { ok, error, unauthorized } from '../utils/response.js';
import { logAudit } from '../utils/audit.js';

// ── Rate Limiting Map (isolate-scoped, not distributed) ──────────────────────
// Key: username, Value: { attempts: number, lastAttempt: timestamp, cooldownUntil: timestamp }
const rateLimitMap = new Map();
const RATE_LIMIT_ATTEMPTS = 5;
const RATE_LIMIT_COOLDOWN_MS = 15 * 60 * 1000; // 15 minutes

function checkRateLimit(username) {
  const now = Date.now();
  const record = rateLimitMap.get(username);
  
  if (!record) return { allowed: true, remaining: RATE_LIMIT_ATTEMPTS };
  
  // Cooldown expired, reset
  if (record.cooldownUntil && now >= record.cooldownUntil) {
    rateLimitMap.delete(username);
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
  const now = Date.now();
  const record = rateLimitMap.get(username) || { attempts: 0, lastAttempt: now };
  record.attempts = (record.attempts || 0) + 1;
  record.lastAttempt = now;
  
  if (record.attempts >= RATE_LIMIT_ATTEMPTS) {
    record.cooldownUntil = now + RATE_LIMIT_COOLDOWN_MS;
  }
  
  rateLimitMap.set(username, record);
}

function recordSuccessfulLogin(username) {
  rateLimitMap.delete(username);
}

export async function handleAuth(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace('/api/auth', '');

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

async function handleLogin(request, env, origin) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }

  const { username, password } = body;
  if (!username || !password) return error('Username and password required', 400, origin);

  // ── Rate Limit Check ──────────────────────────────────────────────────────
  const rateLimit = checkRateLimit(username);
  if (!rateLimit.allowed) {
    return error(`Too many failed attempts. Please try again in ${rateLimit.retryAfter} seconds.`, 429, origin);
  }

  const user = await env.DB.prepare(
    'SELECT * FROM users WHERE (username = ? OR email = ?) AND is_active = 1'
  ).bind(username, username).first();

  if (!user) {
    recordFailedLogin(username);
    return unauthorized(origin);
  }

  // First login: if hash is placeholder, set password
  let valid = false;
  if (user.password_hash === '$2a$10$placeholder_change_this') {
    if (password === 'Admin@123') {
      // Force set real hash on first login
      const hash = await hashPassword(password);
      await env.DB.prepare('UPDATE users SET password_hash = ? WHERE id = ?').bind(hash, user.id).run();
      valid = true;
    }
  } else {
    valid = await verifyPassword(password, user.password_hash);
  }

  if (!valid) {
    recordFailedLogin(username);
    return unauthorized(origin);
  }

  // ── Successful login - reset rate limit ────────────────────────────────────
  recordSuccessfulLogin(username);

  const secret = env.JWT_SECRET || 'dev-secret-change-me';
  const token = await createToken(
    { sub: user.id, username: user.username, role: user.role, exp: Date.now() + 8 * 60 * 60 * 1000 },
    secret
  );

  const userObj = { id: user.id, username: user.username, email: user.email, full_name: user.full_name, role: user.role };

  try {
    await env.DB.prepare('UPDATE users SET updated_at = datetime(\'now\') WHERE id = ?').bind(user.id).run();
    await logAudit(env, userObj, 'LOGIN', 'auth', user.id, null, { action: 'User logged in' });
  } catch (err) {
    console.error('Non-critical login DB write failed:', err);
  }

  return ok({
    token,
    user: userObj
  }, 200, origin);
}

async function handleLogout(request, env, origin) {
  // JWT is stateless; client just discards token
  const user = await authenticate(request, env);
  if (user) {
    await logAudit(env, user, 'LOGOUT', 'auth', user.id, null, { action: 'User logged out' });
  }
  return ok({ message: 'Logged out' }, 200, origin);
}

async function handleMe(request, env, origin) {
  const { authenticate: auth } = await import('../utils/auth.js');
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  return ok(user, 200, origin);
}

async function handleChangePassword(request, env, origin) {
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);

  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  const { current_password, new_password } = body;
  if (!current_password || !new_password) return error('Both passwords required', 400, origin);
  if (new_password.length < 12) return error('New password must be at least 12 characters', 400, origin);

  const dbUser = await env.DB.prepare('SELECT password_hash FROM users WHERE id = ?').bind(user.id).first();
  const valid = await verifyPassword(current_password, dbUser.password_hash);
  if (!valid) return error('Current password incorrect', 400, origin);

  const newHash = await hashPassword(new_password);
  await env.DB.prepare('UPDATE users SET password_hash = ?, updated_at = datetime(\'now\') WHERE id = ?')
    .bind(newHash, user.id).run();

  await logAudit(env, user, 'CHANGE_PASSWORD', 'auth', user.id, null, { action: 'User changed password' });

  return ok({ message: 'Password changed successfully' }, 200, origin);
}
