import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination, getSearchParam } from '../utils/pagination.js';
import { requireExportAccess, getExportPagination } from '../utils/export.js';
import { buildOutboxQuery, buildConditionalOutboxQuery, buildOutboxFromLastInsertQuery } from '../utils/sync_engine.js';
import { mapDBToPayload } from '../utils/sync_mapper.js';
import { buildAuditStatement, buildCreateAuditStatement, buildCreateAuditFromOutboxStatement, executeAtomicImportBatch } from '../utils/audit.js';
import {
  isOfficialMasterBranchCode,
  isSyntheticTestBranch,
  enrichBranchMasterRow,
  compareBranchesMasterOrder,
} from '../utils/branchMaster.js';

export async function handleBranches(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname.replace('/api/branches', '');
  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'GET' && path === '/export') {
    const { response } = await requireExportAccess(request, env, origin, 'branches');
    if (response) return response;
    return listBranches(request, env, origin, { isExport: true });
  }

  // Authentication is mandatory; strictly enforces official 2-role RBAC without legacy role exceptions.
  if (request.method === 'GET' && path === '') {
    const u = await authenticate(request, env);
    if (!u) return unauthorized(origin);
    if (!hasPermission(u, 'branches', 'read')) return forbidden(origin);
    return listBranches(request, env, origin);
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (request.method === 'GET' && !hasPermission(user, 'branches', 'read')) return forbidden(origin);

  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') return getBranch(id, env, origin);
    if (request.method === 'PUT' || request.method === 'PATCH') {
      if (!hasPermission(user, 'branches', 'write')) return forbidden(origin);
      return updateBranch(id, request, env, origin, user);
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'branches', 'delete')) return forbidden(origin);
      return deleteBranch(id, env, origin, user);
    }
  }
  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'branches', 'import')) return forbidden(origin);
    return importBranches(request, env, origin, user);
  }
  if (request.method === 'POST' && path === '') {
    if (!hasPermission(user, 'branches', 'write')) return forbidden(origin);
    return createBranch(request, env, origin, user);
  }
  return error('Not found', 404, origin);
}

function filterAndSortMasterBranches(rawRows, { search = '', officialOnlyWhenPopulated = true } = {}) {
  const enriched = (rawRows || []).map(r => enrichBranchMasterRow(r));
  const withoutSynthetic = enriched.filter(r => !isSyntheticTestBranch(r));
  const officialCount = withoutSynthetic.filter(r => isOfficialMasterBranchCode(r.code)).length;
  let result = (officialOnlyWhenPopulated && officialCount >= 65
    ? withoutSynthetic.filter(r => isOfficialMasterBranchCode(r.code))
    : withoutSynthetic
  );
  if (search) {
    const q = search.toLowerCase();
    result = result.filter(r =>
      String(r.name || '').toLowerCase().includes(q) ||
      String(r.code || '').toLowerCase().includes(q) ||
      String(r.full_name || '').toLowerCase().includes(q) ||
      String(r.city || '').toLowerCase().includes(q)
    );
  }
  return result.sort(compareBranchesMasterOrder);
}

async function listBranches(request, env, origin, { isExport = false } = {}) {
  const { page, limit, offset } = isExport ? getExportPagination(request.url) : getPagination(request.url);
  const search = getSearchParam(request.url, 'search').trim();
  const all = new URL(request.url).searchParams.get('all');

  if (!isExport && all === '1') {
    const rows = await env.DB.prepare('SELECT * FROM branches WHERE is_active = 1 AND deleted_at IS NULL ORDER BY code ASC, id ASC').all();
    const finalList = filterAndSortMasterBranches(rows.results || []);
    return ok(finalList, 200, origin);
  }

  const allActiveRows = await env.DB.prepare('SELECT * FROM branches WHERE deleted_at IS NULL ORDER BY code ASC, id ASC').all();
  const filteredSorted = filterAndSortMasterBranches(allActiveRows.results || [], { search });
  const total = filteredSorted.length;
  const paged = filteredSorted.slice(offset, offset + limit);
  return paginated(paged, total, page, limit, origin, env?.ENVIRONMENT);
}

async function getBranch(id, env, origin) {
  const row = await env.DB.prepare('SELECT * FROM branches WHERE id = ? AND deleted_at IS NULL').bind(id).first();
  if (!row) return notFound(origin);
  return ok(enrichBranchMasterRow(row), 200, origin);
}

async function createBranch(request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
  const { city, full_name } = body;
  let { code, name } = body;
  if (!full_name || !String(full_name).trim()) return error('full_name required', 400, origin);
  if (!code || !String(code).trim()) code = `AUTO-${Math.floor(Math.random() * 100000)}`;
  code = String(code).trim();
  if (!name || !String(name).trim()) name = String(full_name).trim();
  else name = String(name).trim();

  const dupBranch = await env.DB.prepare(
    'SELECT id FROM branches WHERE LOWER(TRIM(code)) = LOWER(TRIM(?))'
  ).bind(code).first();
  if (dupBranch) {
    return error(`Kode cabang '${code}' sudah terdaftar`, 409, origin);
  }

  const newSnapshot = {
    code,
    name,
    full_name,
    city: city || null,
    is_active: 1,
    row_version: 1,
    last_sync_source: 'FCMS'
  };
  const payload = mapDBToPayload('Data Cabang', {
    id: '',
    ...newSnapshot
  });
  const corrId = crypto.randomUUID();

  const batchRes = await env.DB.batch([
    env.DB.prepare(
      "INSERT INTO branches (code, name, full_name, city, row_version, last_sync_source) VALUES (?, ?, ?, ?, 1, 'FCMS')"
    ).bind(code, name, full_name, city || null),
    buildOutboxFromLastInsertQuery(env, 'Data Cabang', payload, corrId),
    buildCreateAuditFromOutboxStatement(env, user, 'branches', corrId, newSnapshot)
  ]);

  const newId = batchRes?.[0]?.meta?.last_row_id ?? 1;

  return ok({ id: newId }, 201, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function updateBranch(id, request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);

  // OCC: Strict row_version enforcement
  if (body.row_version === undefined || body.row_version === null) {
    return error('row_version is required for optimistic concurrency control', 400, origin);
  }
  const clientVersion = Number(body.row_version);
  if (!Number.isInteger(clientVersion) || clientVersion < 0) {
    return error('Invalid row_version format: must be a non-negative integer', 400, origin);
  }

  const existing = await env.DB.prepare('SELECT * FROM branches WHERE id = ? AND deleted_at IS NULL').bind(id).first();
  if (!existing) return notFound(origin);

  const isExistingNull = existing.row_version === null || existing.row_version === undefined;
  if (isExistingNull) {
    if (clientVersion !== 1 && clientVersion !== 0) {
      return error('Conflict: Record has been modified by another user or session', 409, origin);
    }
  } else {
    if (existing.row_version !== clientVersion) {
      return error('Conflict: Record has been modified by another user or session', 409, origin);
    }
  }

  const { code, name, full_name, city, is_active } = body;
  
  const finalCode = (code !== undefined && code !== null && String(code).trim() !== '') ? String(code).trim() : existing.code;
  const finalName = (name !== undefined && name !== null && String(name).trim() !== '') ? String(name).trim() : existing.name;
  const finalFullName = (full_name !== undefined && full_name !== null && String(full_name).trim() !== '') ? String(full_name).trim() : existing.full_name;
  const finalCity = (city !== undefined) ? (city !== null && String(city).trim() !== '' ? String(city).trim() : null) : existing.city;
  const finalIsActive = (is_active !== undefined) ? (is_active ? 1 : 0) : existing.is_active;

  const baseVersion = isExistingNull ? clientVersion : existing.row_version;
  const nextRowVersion = baseVersion + 1;

  const updatedRecord = {
    ...existing,
    code: finalCode,
    name: finalName,
    full_name: finalFullName,
    city: finalCity,
    is_active: finalIsActive,
    row_version: nextRowVersion,
    last_sync_source: 'FCMS',
    branch_code: finalCode,
    branch_name: finalName || finalFullName
  };
  const payload = mapDBToPayload('Data Cabang', updatedRecord);

  const updateStmt = env.DB.prepare(
    "UPDATE branches SET code = COALESCE(?, code), name = COALESCE(?, name), full_name = COALESCE(?, full_name), city = COALESCE(?, city), is_active = COALESCE(?, is_active), updated_at = datetime('now'), row_version = COALESCE(row_version, ?) + 1, last_sync_source = 'FCMS' WHERE id = ? AND (row_version = ? OR (row_version IS NULL AND (? = 1 OR ? = 0)))"
  ).bind(
    (code !== undefined && code !== null && String(code).trim() !== '') ? String(code).trim() : null,
    (name !== undefined && name !== null && String(name).trim() !== '') ? String(name).trim() : null,
    (full_name !== undefined && full_name !== null && String(full_name).trim() !== '') ? String(full_name).trim() : null,
    (city !== undefined) ? (city !== null && String(city).trim() !== '' ? String(city).trim() : null) : null,
    (is_active !== undefined) ? (is_active ? 1 : 0) : null,
    clientVersion,
    id,
    clientVersion,
    clientVersion,
    clientVersion
  );

  const batchRes = await env.DB.batch([
    updateStmt,
    buildAuditStatement(env, user, 'UPDATE', 'branches', id, existing, updatedRecord, null, 'SUCCESS', true),
    buildConditionalOutboxQuery(env, 'Data Cabang', id, 'UPDATE', payload)
  ]);

  const changes = batchRes?.[0]?.meta?.changes ?? batchRes?.[0]?.changes;
  if (changes === 0) {
    return error('Conflict: Record has been modified by another user or session', 409, origin);
  }
  
  return ok({ message: 'Branch updated', row_version: nextRowVersion }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function deleteBranch(id, env, origin, user = null) {
  const existing = await env.DB.prepare('SELECT * FROM branches WHERE id = ? AND deleted_at IS NULL').bind(id).first();
  if (!existing) return notFound(origin);
  
  const deletedRecord = {
    ...existing,
    is_active: 0,
    deleted_at: new Date().toISOString(),
    branch_code: existing.code,
    branch_name: existing.name || existing.full_name
  };
  const payload = mapDBToPayload('Data Cabang', deletedRecord);

  const { buildAuditStatement } = await import('../utils/audit.js');
  const auditStmt = buildAuditStatement(env, user, 'DELETE', 'branches', id, existing, null, null, 'SUCCESS');

  await env.DB.batch([
    env.DB.prepare("UPDATE branches SET is_active = 0, deleted_at = datetime('now'), updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ?").bind(id),
    buildOutboxQuery(env, 'Data Cabang', id, 'DELETE', payload),
    auditStmt
  ]);
  
  return ok({ message: 'Branch deactivated' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
}

async function importBranches(request, env, origin, user = null) {
  let body;
  try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
  if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
  if (body.length === 0) return ok({ message: 'No data to import' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });

  const existingRes = await env.DB.prepare('SELECT * FROM branches').all();
  const existingMap = new Map();
  for (const row of (existingRes.results || [])) {
    if (row.code) existingMap.set(row.code, row);
  }

  const stmts = [];
  let processed = 0;
  for (const item of body) {
    if (!item.code || !item.name) continue;
    const fullName = item.full_name || item.name;
    const cityVal = item.city || null;
    const oldRow = existingMap.get(item.code);
    if (oldRow) {
      const updatedSnapshot = { ...oldRow, name: item.name, full_name: fullName, city: cityVal, is_active: 1 };
      stmts.push(
        env.DB.prepare(
          `UPDATE branches SET name = ?, full_name = ?, city = ?, is_active = 1, updated_at = datetime('now') WHERE id = ?`
        ).bind(item.name, fullName, cityVal, oldRow.id),
        buildAuditStatement(env, user, 'UPDATE', 'branches', oldRow.id, oldRow, updatedSnapshot, { source: 'import' }, 'SUCCESS', true)
      );
      existingMap.set(item.code, updatedSnapshot);
    } else {
      const createdSnapshot = { code: item.code, name: item.name, full_name: fullName, city: cityVal, is_active: 1 };
      stmts.push(
        env.DB.prepare(
          `INSERT INTO branches (code, name, full_name, city, is_active) VALUES (?, ?, ?, ?, 1)`
        ).bind(item.code, item.name, fullName, cityVal),
        buildCreateAuditStatement(env, user, 'branches', item.code, createdSnapshot, { source: 'import' })
      );
    }
    processed++;
  }

  try {
    if (stmts.length > 0) await executeAtomicImportBatch(env, stmts);
    return ok({ message: `Berhasil mengimport ${processed} cabang` }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
  } catch (err) {
    console.error('Import error:', err);
    return error('Gagal mengimport data', 500, origin);
  }
}
