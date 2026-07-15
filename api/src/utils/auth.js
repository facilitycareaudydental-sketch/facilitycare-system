// Auth utilities - JWT + password hashing using Web Crypto API

export async function hashPassword(password) {
  const encoder = new TextEncoder();
  const data = encoder.encode(password);
  const hashBuffer = await crypto.subtle.digest('SHA-256', data);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
}

export async function verifyPassword(password, hash) {
  const computed = await hashPassword(password);
  return computed === hash;
}

export function generateSessionId() {
  const array = new Uint8Array(32);
  crypto.getRandomValues(array);
  return Array.from(array).map(b => b.toString(16).padStart(2, '0')).join('');
}

export async function createToken(payload, secret) {
  const header = btoa(JSON.stringify({ alg: 'HS256', typ: 'JWT' }));
  const body = btoa(JSON.stringify({ ...payload, iat: Date.now() }));
  const data = `${header}.${body}`;
  const encoder = new TextEncoder();
  const keyData = encoder.encode(secret);
  const messageData = encoder.encode(data);
  const key = await crypto.subtle.importKey(
    'raw', keyData, { name: 'HMAC', hash: 'SHA-256' }, false, ['sign']
  );
  const signature = await crypto.subtle.sign('HMAC', key, messageData);
  const sigArray = Array.from(new Uint8Array(signature));
  const sig = btoa(String.fromCharCode(...sigArray));
  return `${data}.${sig}`;
}

export async function verifyToken(token, secret) {
  try {
    const parts = token.split('.');
    if (parts.length !== 3) return null;
    const [header, body, sig] = parts;
    const data = `${header}.${body}`;
    const encoder = new TextEncoder();
    const keyData = encoder.encode(secret);
    const messageData = encoder.encode(data);
    const key = await crypto.subtle.importKey(
      'raw', keyData, { name: 'HMAC', hash: 'SHA-256' }, false, ['verify']
    );
    const sigBytes = Uint8Array.from(atob(sig), c => c.charCodeAt(0));
    const valid = await crypto.subtle.verify('HMAC', key, sigBytes, messageData);
    if (!valid) return null;
    const payload = JSON.parse(atob(body));
    // Check expiry
    if (payload.exp && Date.now() > payload.exp) return null;
    return payload;
  } catch {
    return null;
  }
}

export async function authenticate(request, env) {
  const authHeader = request.headers.get('Authorization');
  if (!authHeader || !authHeader.startsWith('Bearer ')) return null;
  const token = authHeader.slice(7);
  const payload = await verifyToken(token, env.JWT_SECRET || 'dev-secret-change-me');
  if (!payload) return null;
  // Look up user
  const user = await env.DB.prepare(
    'SELECT id, username, email, full_name, role, is_active FROM users WHERE id = ?'
  ).bind(payload.sub).first();
  if (!user || !user.is_active) return null;
  return user;
}

export function hasPermission(user, module, action) {
  if (user.role === 'superadmin') return true;
  // For simplicity, role-based checks
  const rolePermissions = {
    admin: { level: 4 },
    manager: { level: 3 },
    spv: { level: 2 },
    viewer: { level: 1 },
  };
  const writeModules = ['employees', 'contracts', 'schedule', 'issues', 'one_on_one',
    'training', 'relievers', 'reports', 'sop', 'checklist', 'forms', 'supply_requests'];
  const userLevel = rolePermissions[user.role]?.level || 0;
  if (action === 'read') return userLevel >= 1;
  if (action === 'write') return userLevel >= 2;
  if (action === 'delete') return userLevel >= 3;
  if (action === 'admin') return userLevel >= 4;
  return false;
}
