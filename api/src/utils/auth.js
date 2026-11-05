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
  if (!secret || typeof secret !== 'string' || secret.trim() === '') {
    throw new Error('JWT_SECRET_REQUIRED: Valid JWT_SECRET is required to sign tokens.');
  }
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
  if (!secret || typeof secret !== 'string' || secret.trim() === '') {
    return null;
  }
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
  if (!env?.JWT_SECRET || typeof env.JWT_SECRET !== 'string' || env.JWT_SECRET.trim() === '') {
    console.error('AUTH_FATAL: JWT_SECRET environment variable is not configured.');
    return null;
  }
  const payload = await verifyToken(token, env.JWT_SECRET);
  if (!payload) return null;
  // Look up user
  const user = await env.DB.prepare(
    'SELECT id, username, email, full_name, role, is_active FROM users WHERE id = ?'
  ).bind(payload.sub).first();
  if (!user || !user.is_active) return null;
  return user;
}

export function hasPermission(user, module, action) {
  if (user.role === 'superadmin' || user.role === 'admin') return true;
  
  if (action === 'export' || action === 'import' || action === 'admin') {
    return false;
  }

  if (user.role === 'input_lembur') {
    if (module === 'overtime') {
      if (action === 'read' || action === 'write') return true;
      if (action === 'delete') return false;
    } else {
      if (action === 'read') return true;
    }
    return false;
  }

  if (user.role === 'editor_khusus') {
    if (module === 'issues' || module === 'overtime') {
      if (action === 'read' || action === 'write' || action === 'delete') return true;
    } else {
      if (action === 'read') return true;
    }
    return false;
  }
  
  // Selain superadmin / admin: cukup viewer saja!
  return action === 'read';
}
