export async function logAudit(env, user, action, moduleName, targetId, oldData = null, newData = null) {
  try {
    const userId = user?.id || 0;
    const userName = user?.full_name || 'System';
    const userRole = user?.role || 'system';
    
    const safeStringify = (data) => {
      if (!data) return null;
      if (typeof data === 'string') return data;
      try { return JSON.stringify(data); } catch { return String(data); }
    };

    await env.DB.prepare(
      `INSERT INTO audit_logs (user_id, user_name, user_role, action, module, target_id, old_data, new_data)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?)`
    ).bind(
      userId, userName, userRole, action, moduleName, String(targetId), safeStringify(oldData), safeStringify(newData)
    ).run();
  } catch (error) {
    console.error('Failed to write audit log:', error);
  }
}
