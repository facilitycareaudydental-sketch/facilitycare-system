import { hashPassword, verifyPassword, createToken, authenticate } from '../utils/auth.js';
import { ok, error, unauthorized } from '../utils/response.js';

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

  const user = await env.DB.prepare(
    'SELECT * FROM users WHERE (username = ? OR email = ?) AND is_active = 1'
  ).bind(username, username).first();

  if (!user) return unauthorized(origin);

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

  if (!valid) return unauthorized(origin);

  const secret = env.JWT_SECRET || 'dev-secret-change-me';
  const token = await createToken(
    { sub: user.id, username: user.username, role: user.role, exp: Date.now() + 8 * 60 * 60 * 1000 },
    secret
  );

  await env.DB.prepare('UPDATE users SET updated_at = datetime(\'now\') WHERE id = ?').bind(user.id).run();

  return ok({
    token,
    user: { id: user.id, username: user.username, email: user.email, full_name: user.full_name, role: user.role }
  }, 200, origin);
}

async function handleLogout(request, env, origin) {
  // JWT is stateless; client just discards token
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
  if (new_password.length < 6) return error('New password must be at least 6 characters', 400, origin);

  const dbUser = await env.DB.prepare('SELECT password_hash FROM users WHERE id = ?').bind(user.id).first();
  const valid = await verifyPassword(current_password, dbUser.password_hash);
  if (!valid) return error('Current password incorrect', 400, origin);

  const newHash = await hashPassword(new_password);
  await env.DB.prepare('UPDATE users SET password_hash = ?, updated_at = datetime(\'now\') WHERE id = ?')
    .bind(newHash, user.id).run();

  return ok({ message: 'Password changed successfully' }, 200, origin);
}
