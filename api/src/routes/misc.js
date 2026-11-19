// SOP, Master Checklist, Master Forms, PIC list
// Phase 5.48: all internal reference endpoints require authentication + backend
// RBAC (module read). SOP bundle writes are restricted to an explicit key
// allowlist with size/shape validation and credential-field scanning.
import { authenticate, hasPermission } from '../utils/auth.js';
import { ok, error, unauthorized, forbidden, notFound, paginated } from '../utils/response.js';
import { getPagination } from '../utils/pagination.js';
import { requireExportAccess, getExportPagination } from '../utils/export.js';
import { FORBIDDEN_PAYLOAD_PATTERNS } from '../utils/sync_mapper.js';
import {
  buildAuditStatement,
  buildCreateAuditStatement,
  buildSoftDeleteRecordAuditStatement,
  executeAtomicImportBatch
} from '../utils/audit.js';
import { executeJournaledImport, extractImportRequestMeta } from '../utils/import_journal.js';

// ─── Phase 5.48: input validation constants & helpers ────────────────────────
const MAX_SEARCH_LENGTH = 200;
const MAX_NAME_LENGTH = 255;
const MAX_CATEGORY_LENGTH = 100;
const MAX_LINK_LENGTH = 1000;
const MAX_DESCRIPTION_LENGTH = 50000;
// Bundle keys used by the SOP editors (frontend/src/pages/sop_facility_care.js,
// sop_security.js). Writes outside this allowlist are rejected.
const SOP_BUNDLE_KEYS = new Set([
  'fcms_sop_fc_ik_v1',
  'fcms_sop_fc_tools_v1',
  'fcms_sop_fc_chem_v1',
  'fcms_sop_sec_ik_v1',
  'fcms_sop_sec_assets_v1'
]);
// Phase 5.66.50: Isolated OAT fixture namespace allowed ONLY in staging environment.
// Production and all non-staging environments strictly enforce SOP_BUNDLE_KEYS only.
const SOP_OAT_STAGING_KEY_RE = /^fcms_sop_(fc|sec)_oat_[a-z0-9_]{1,40}$/;
function isStagingOatSopBundleKey(key, env) {
  return String(env?.ENVIRONMENT || '').trim().toLowerCase() === 'staging' && SOP_OAT_STAGING_KEY_RE.test(key);
}
const SOP_BUNDLE_MAX_ITEMS = 500;
const SOP_BUNDLE_MAX_CHARS = 500000;

function fieldError(value, field, { max = MAX_NAME_LENGTH, required = false } = {}) {
  if (value === undefined || value === null || value === '') {
    return required ? `${field} is required` : null;
  }
  if (typeof value !== 'string') return `${field} must be a string`;
  if (value.length > max) return `${field} exceeds maximum length of ${max}`;
  return null;
}

function firstFieldError(body, specs) {
  for (const [field, options] of specs) {
    const err = fieldError(body[field], field, options || {});
    if (err) return err;
  }
  return null;
}

function findForbiddenKey(value, keyName = '') {
  if (keyName && FORBIDDEN_PAYLOAD_PATTERNS.some((re) => re.test(keyName))) return keyName;
  if (Array.isArray(value)) {
    for (const item of value) {
      const found = findForbiddenKey(item);
      if (found) return found;
    }
    return null;
  }
  if (value && typeof value === 'object') {
    for (const [k, v] of Object.entries(value)) {
      const found = findForbiddenKey(v, k);
      if (found) return found;
    }
  }
  return null;
}

export async function handleMisc(request, env, origin) {
  const url = new URL(request.url);
  const path = url.pathname;

  if (path.startsWith('/api/sop/bundle')) return handleSopBundle(request, env, origin, path.replace('/api/sop/bundle', ''));
  if (path.startsWith('/api/sop')) return handleTable(request, env, origin, 'sop', path.replace('/api/sop', ''));
  if (path.startsWith('/api/checklist')) return handleTable(request, env, origin, 'master_checklist', path.replace('/api/checklist', ''));
  if (path.startsWith('/api/forms')) return handleForms(request, env, origin, path.replace('/api/forms', ''));
  if (path.startsWith('/api/pic')) return handlePic(request, env, origin);
  if (path.startsWith('/api/options')) return handleOptions(request, env, origin);

  return error('Not found', 404, origin);
}

async function listTableRows(request, env, origin, table, { isExport = false } = {}) {
  const { page, limit, offset } = isExport ? getExportPagination(request.url) : getPagination(request.url);
  const url = new URL(request.url);
  const search = (url.searchParams.get('search') || '').trim();
  const category = (url.searchParams.get('category') || '').trim();
  if (search.length > MAX_SEARCH_LENGTH || category.length > MAX_SEARCH_LENGTH) {
    return error(`Search/filter parameters exceed maximum length of ${MAX_SEARCH_LENGTH}`, 400, origin);
  }
  const all = url.searchParams.get('all');
  const includeStats = !isExport && url.searchParams.get('include_stats') === '1';
  const baseConditions = ['deleted_at IS NULL'];
  if (table === 'sop') baseConditions.push("(category != 'sop_bundle' OR category IS NULL)");
  const baseCondition = baseConditions.join(' AND ');
  const descField = table === 'sop' ? 'notes' : 'description';

  if (!isExport && all === '1') {
    const sql = `SELECT * FROM ${table} WHERE ${baseCondition} ORDER BY name ASC, id ASC`;
    const rows = await env.DB.prepare(sql).all();
    return ok(rows.results, 200, origin);
  }

  const conditions = [baseCondition];
  const values = [];
  if (category) {
    if (category === 'Lainnya') {
      conditions.push("(category IS NULL OR TRIM(category) = '' OR category = 'Lainnya')");
    } else {
      conditions.push('category = ?');
      values.push(category);
    }
  }
  if (search) {
    conditions.push(`(name LIKE ? OR category LIKE ? OR ${descField} LIKE ?)`);
    values.push(`%${search}%`, `%${search}%`, `%${search}%`);
  }
  const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';

  const queries = [
    env.DB.prepare(`SELECT COUNT(*) as total FROM ${table} ${where}`).bind(...values).first(),
    env.DB.prepare(`SELECT * FROM ${table} ${where} ORDER BY name ASC, id ASC LIMIT ? OFFSET ?`).bind(...values, limit, offset).all(),
  ];

  if (includeStats) {
    const statsWhere = baseCondition ? `WHERE ${baseCondition}` : '';
    queries.push(
      env.DB.prepare(
        `SELECT COALESCE(NULLIF(TRIM(category), ''), 'Lainnya') as cat, COUNT(*) as cnt
         FROM ${table} ${statsWhere}
         GROUP BY cat
         ORDER BY cnt DESC, cat ASC`
      ).all(),
      env.DB.prepare(
        `SELECT SUM(CASE WHEN document_link IS NOT NULL AND TRIM(document_link) != '' THEN 1 ELSE 0 END) as linked_count
         FROM ${table} ${statsWhere}`
      ).first()
    );
  }

  const [countResult, rows, catStatsRes, linkStatsRes] = await Promise.all(queries);
  let extra = null;
  if (includeStats) {
    const categoryCounts = { all: 0 };
    for (const r of catStatsRes?.results || []) {
      const cnt = Number(r.cnt || 0);
      categoryCounts.all += cnt;
      categoryCounts[r.cat] = cnt;
    }
    extra = {
      stats: {
        categoryCounts,
        linkedCount: Number(linkStatsRes?.linked_count || 0)
      }
    };
  }

  return paginated(rows.results || [], countResult?.total || 0, page, limit, origin, env?.ENVIRONMENT, extra);
}

async function handleTable(request, env, origin, table, path) {
  const moduleKey = table === 'sop' ? 'sop' : 'checklist';

  if (request.method === 'GET' && path === '/export') {
    const { response } = await requireExportAccess(request, env, origin, moduleKey);
    if (response) return response;
    return listTableRows(request, env, origin, table, { isExport: true });
  }

  // Phase 5.48: SOP/checklist lists are internal reference data.
  // Authentication + module-level read RBAC is mandatory (no public bypass).
  if (request.method === 'GET' && path === '') {
    const user = await authenticate(request, env);
    if (!user) return unauthorized(origin);
    if (!hasPermission(user, moduleKey, 'read')) return forbidden(origin);
    return listTableRows(request, env, origin, table);
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (request.method === 'GET') {
    if (!hasPermission(user, moduleKey, 'read')) return forbidden(origin);
  } else if (request.method === 'DELETE') {
    if (!hasPermission(user, moduleKey, 'delete')) return forbidden(origin);
  } else {
    if (!hasPermission(user, moduleKey, 'write')) return forbidden(origin);
  }

  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'POST' && path === '') {
    let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
    if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
    const { name, category, document_link, description } = body;
    const fieldErr = firstFieldError(body, [
      ['name', { required: true }],
      ['category', { max: MAX_CATEGORY_LENGTH }],
      ['document_link', { max: MAX_LINK_LENGTH }],
      ['description', { max: MAX_DESCRIPTION_LENGTH }]
    ]);
    if (fieldErr) return error(fieldErr, 400, origin);
    if (table === 'sop' && (category || '').trim() === 'sop_bundle') {
      return error("category 'sop_bundle' is reserved for internal SOP bundles (use /api/sop/bundle)", 400, origin);
    }
    const descField = table === 'sop' ? 'notes' : 'description';
    const insertStmt = env.DB.prepare(
      `INSERT INTO ${table} (name, category, document_link, ${descField}) VALUES (?, ?, ?, ?)`
    ).bind(name, category || null, document_link || null, description || null);
    const auditStmt = buildCreateAuditStatement(env, user, moduleKey, 'new', {
      name,
      category: category || null,
      document_link: document_link || null,
      [descField]: description || null
    });
    const batchRes = await env.DB.batch([insertStmt, auditStmt]);
    return ok({ id: batchRes?.[0]?.meta?.last_row_id }, 201, origin, undefined, { 'X-Audit-Batched': 'true' });
  }

  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, moduleKey, 'import')) return forbidden(origin);
    let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
    if (!Array.isArray(body)) return error('Payload must be an array', 400, origin);
    if (body.length === 0) return ok({ message: 'No data to import' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });

    const importMeta = extractImportRequestMeta(request);
    try {
      return await executeJournaledImport(env, user, origin, {
        module: `${moduleKey}_route`,
        rows: body,
        onDuplicate: 'upsert',
        importMeta,
        buildPlan: async ({ startRowIndex, initialCounts }) => {
          const baseWhere = table === 'sop' ? " WHERE deleted_at IS NULL AND (category != 'sop_bundle' OR category IS NULL)" : " WHERE deleted_at IS NULL";
          const existing = await env.DB.prepare(`SELECT * FROM ${table}${baseWhere}`).all();
          const existingMap = new Map();
          (existing.results || []).forEach(r => {
            const key = (r.name || '').toLowerCase().trim() + '_' + (r.category || '').toLowerCase().trim();
            existingMap.set(key, r);
          });

          const seenNewKeys = new Set();
          const insertGroups = [];
          let inserted = Number(initialCounts?.inserted || 0);
          let updated = Number(initialCounts?.updated || 0);
          let skipped = Number(initialCounts?.skipped || 0);
          const descField = table === 'sop' ? 'notes' : 'description';

          for (let idx = startRowIndex; idx < body.length; idx++) {
            const item = body[idx];
            if (!item || typeof item !== 'object' || Array.isArray(item)) { skipped++; continue; }
            if (!item.name) { skipped++; continue; }
            if (table === 'sop' && typeof item.category === 'string' && item.category.trim() === 'sop_bundle') { skipped++; continue; }
            const nameErr = fieldError(item.name, 'name');
            const catErr = fieldError(item.category, 'category', { max: MAX_CATEGORY_LENGTH });
            const linkErr = fieldError(item.document_link, 'document_link', { max: MAX_LINK_LENGTH });
            const descErr = fieldError(item.description, 'description', { max: MAX_DESCRIPTION_LENGTH });
            if (nameErr || catErr || linkErr || descErr || findForbiddenKey(item)) { skipped++; continue; }

            const key = (item.name || '').toLowerCase().trim() + '_' + (item.category || '').toLowerCase().trim();
            if (existingMap.has(key)) {
              const oldRow = existingMap.get(key);
              const nextCategory = item.category || oldRow.category || null;
              const nextLink = item.document_link || oldRow.document_link || null;
              const nextDesc = item.description || oldRow[descField] || null;
              const afterRow = { ...oldRow, category: nextCategory, document_link: nextLink, [descField]: nextDesc };
              updated++;
              insertGroups.push({
                rowIndex: idx,
                counts: { inserted, updated, skipped },
                stmts: [
                  env.DB.prepare(
                    `UPDATE ${table} SET category = COALESCE(?, category), document_link = COALESCE(?, document_link), ${descField} = COALESCE(?, ${descField}), updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ? AND deleted_at IS NULL`
                  ).bind(item.category || null, item.document_link || null, item.description || null, oldRow.id),
                  buildAuditStatement(env, user, 'UPDATE', moduleKey, oldRow.id, oldRow, afterRow, { source: 'import' }, 'SUCCESS', true),
                ],
              });
              existingMap.set(key, afterRow);
            } else if (seenNewKeys.has(key)) {
              skipped++;
            } else {
              seenNewKeys.add(key);
              const newSnapshot = {
                name: item.name,
                category: item.category || null,
                document_link: item.document_link || null,
                [descField]: item.description || null,
              };
              inserted++;
              insertGroups.push({
                rowIndex: idx,
                counts: { inserted, updated, skipped },
                stmts: [
                  env.DB.prepare(
                    `INSERT INTO ${table} (name, category, document_link, ${descField}) VALUES (?, ?, ?, ?)`
                  ).bind(item.name, item.category || null, item.document_link || null, item.description || null),
                  buildCreateAuditStatement(env, user, moduleKey, 'new', newSnapshot, { source: 'import' }),
                ],
              });
            }
          }

          return { insertGroups, finalCounts: { inserted, updated, skipped } };
        },
        buildFinalResult: (counts) => ({
          message: 'Import berhasil',
          inserted: counts.inserted,
          updated: counts.updated,
          skipped: counts.skipped,
          total: body.length,
        }),
      });
    } catch (e) {
      console.error('Misc import error:', e);
      return error('Gagal mengimport data', 500, origin);
    }
  }

  if (idMatch) {
    const id = idMatch[1];
    const extraWhere = table === 'sop' ? " AND (category != 'sop_bundle' OR category IS NULL)" : "";
    if (request.method === 'GET') {
      const row = await env.DB.prepare(`SELECT * FROM ${table} WHERE id = ? AND deleted_at IS NULL${extraWhere}`).bind(id).first();
      return row ? ok(row, 200, origin) : notFound(origin);
    }
    if (request.method === 'PUT') {
      let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
      if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
      const existing = await env.DB.prepare(`SELECT * FROM ${table} WHERE id = ? AND deleted_at IS NULL${extraWhere}`).bind(id).first();
      if (!existing) return notFound(origin);
      const { name, category, document_link, description } = body;
      const fieldErr = firstFieldError(body, [
        ['name', {}],
        ['category', { max: MAX_CATEGORY_LENGTH }],
        ['document_link', { max: MAX_LINK_LENGTH }],
        ['description', { max: MAX_DESCRIPTION_LENGTH }]
      ]);
      if (fieldErr) return error(fieldErr, 400, origin);
      if (table === 'sop' && typeof category === 'string' && category.trim() === 'sop_bundle') {
        return error("category 'sop_bundle' is reserved for internal SOP bundles (use /api/sop/bundle)", 400, origin);
      }
      const descField = table === 'sop' ? 'notes' : 'description';
      const updateStmt = env.DB.prepare(
        `UPDATE ${table} SET name = COALESCE(?, name), category = COALESCE(?, category), document_link = COALESCE(?, document_link), ${descField} = COALESCE(?, ${descField}), updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ? AND deleted_at IS NULL${extraWhere}`
      ).bind(name || null, category || null, document_link || null, description || null, id);
      const newData = {
        ...existing,
        name: name || existing.name,
        category: category || existing.category,
        document_link: document_link || existing.document_link,
        [descField]: description || existing[descField]
      };
      const auditStmt = buildAuditStatement(env, user, 'UPDATE', moduleKey, id, existing, newData, null, 'SUCCESS', true);
      const batchRes = await env.DB.batch([updateStmt, auditStmt]);
      const changes = Number(batchRes?.[0]?.meta?.changes ?? batchRes?.[0]?.changes ?? 0);
      if (changes === 0) return notFound(origin);
      return ok({ message: 'Updated' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, moduleKey, 'delete')) return forbidden(origin);
      const existing = await env.DB.prepare(`SELECT * FROM ${table} WHERE id = ? AND deleted_at IS NULL${extraWhere}`).bind(id).first();
      if (!existing) return notFound(origin);
      const deletedAtIso = new Date().toISOString();
      const delStmt = env.DB.prepare(
        `UPDATE ${table} SET deleted_at = ?, updated_at = ?, row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ? AND deleted_at IS NULL${extraWhere}`
      ).bind(deletedAtIso, deletedAtIso, id);
      const auditStmt = buildSoftDeleteRecordAuditStatement(env, user, moduleKey, table, 'id', 'deleted_at', deletedAtIso, existing);
      const batchRes = await env.DB.batch([delStmt, auditStmt]);
      const changes = Number(batchRes?.[0]?.meta?.changes ?? batchRes?.[0]?.changes ?? 0);
      if (changes === 0) return notFound(origin);
      return ok({ message: 'Deleted' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
    }
  }
  return error('Not found', 404, origin);
}

async function handleForms(request, env, origin, path) {
  if (request.method === 'GET' && path === '/export') {
    const { response } = await requireExportAccess(request, env, origin, 'forms');
    if (response) return response;
    const { page, limit, offset } = getExportPagination(request.url);
    const url = new URL(request.url);
    const search = (url.searchParams.get('search') || '').trim();
    const conditions = [];
    const values = [];
    if (search) { conditions.push('(name LIKE ? OR category LIKE ?)'); values.push(`%${search}%`, `%${search}%`); }
    const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';
    const [countResult, rows] = await Promise.all([
      env.DB.prepare(`SELECT COUNT(*) as total FROM master_forms ${where}`).bind(...values).first(),
      env.DB.prepare(`SELECT * FROM master_forms ${where} ORDER BY name ASC, id ASC LIMIT ? OFFSET ?`).bind(...values, limit, offset).all()
    ]);
    return paginated(rows.results || [], countResult?.total || 0, page, limit, origin, env?.ENVIRONMENT);
  }

  // Phase 5.48: authenticated forms reference (is_public filter kept for response shape)
  if (request.method === 'GET' && path === '/public') {
    const user = await authenticate(request, env);
    if (!user) return unauthorized(origin);
    if (!hasPermission(user, 'forms', 'read')) return forbidden(origin);
    const rows = await env.DB.prepare('SELECT * FROM master_forms WHERE is_public = 1 ORDER BY name ASC, id ASC').all();
    return ok(rows.results, 200, origin);
  }

  if (request.method === 'GET' && path === '') {
    const user = await authenticate(request, env);
    if (!user) return unauthorized(origin);
    if (!hasPermission(user, 'forms', 'read')) return forbidden(origin);
    const { page, limit, offset } = getPagination(request.url);
    const url = new URL(request.url);
    const search = (url.searchParams.get('search') || '').trim();
    if (search.length > MAX_SEARCH_LENGTH) return error(`Search parameter exceeds maximum length of ${MAX_SEARCH_LENGTH}`, 400, origin);
    let conditions = []; let values = [];
    if (search) { conditions.push('(name LIKE ? OR category LIKE ?)'); values.push(`%${search}%`, `%${search}%`); }
    const where = conditions.length ? 'WHERE ' + conditions.join(' AND ') : '';
    const [countResult, rows] = await Promise.all([
      env.DB.prepare(`SELECT COUNT(*) as total FROM master_forms ${where}`).bind(...values).first(),
      env.DB.prepare(`SELECT * FROM master_forms ${where} ORDER BY name ASC, id ASC LIMIT ? OFFSET ?`).bind(...values, limit, offset).all()
    ]);
    return paginated(rows.results || [], countResult?.total || 0, page, limit, origin, env?.ENVIRONMENT);
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);

  if (request.method === 'POST' && path === '/import') {
    if (!hasPermission(user, 'forms', 'import')) return forbidden(origin);
    let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
    const items = Array.isArray(body) ? body : (Array.isArray(body?.data) ? body.data : (Array.isArray(body?.rows) ? body.rows : null));
    if (!items) return error('Payload must be an array or contain data/rows array', 400, origin);
    if (items.length === 0) return ok({ message: 'No data to import', inserted: 0, updated: 0 }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });

    let inserted = 0;
    let updated = 0;
    try {
      const existingAll = await env.DB.prepare('SELECT * FROM master_forms').all();
      const existingMap = new Map();
      for (const row of (existingAll.results || [])) {
        if (row.name) existingMap.set(String(row.name).toLowerCase(), row);
      }

      const stmts = [];
      for (const item of items) {
        const name = String(item.name || item['Nama Form'] || '').trim();
        if (!name) continue;
        if (name.length > MAX_NAME_LENGTH) continue;
        const category = item.category || item['Kategori'] || null;
        if (typeof category === 'string' && category.length > MAX_CATEGORY_LENGTH) continue;
        const document_link = item.document_link || item['Link Form'] || item['Link Dokumen'] || null;
        if (typeof document_link === 'string' && document_link.length > MAX_LINK_LENGTH) continue;
        const description = item.description || item['Deskripsi'] || null;
        if (typeof description === 'string' && description.length > MAX_DESCRIPTION_LENGTH) continue;
        if (findForbiddenKey(item)) continue;
        const is_public = item.is_public !== undefined
          ? (item.is_public === true || item.is_public === 1 || String(item.is_public).toLowerCase() === 'ya' ? 1 : 0)
          : (String(item['Publik'] || '').toLowerCase() === 'tidak' ? 0 : 1);

        const existing = existingMap.get(name.toLowerCase());
        if (existing) {
          const afterRow = {
            ...existing,
            category: category ?? existing.category,
            document_link: document_link ?? existing.document_link,
            description: description ?? existing.description,
            is_public
          };
          stmts.push(
            env.DB.prepare(
              `UPDATE master_forms SET category = COALESCE(?, category), document_link = COALESCE(?, document_link), description = COALESCE(?, description), is_public = ?, updated_at = datetime('now') WHERE id = ?`
            ).bind(category, document_link, description, is_public, existing.id),
            buildAuditStatement(env, user, 'UPDATE', 'forms', existing.id, existing, afterRow, { source: 'import' }, 'SUCCESS', true)
          );
          existingMap.set(name.toLowerCase(), afterRow);
          updated++;
        } else {
          const newSnapshot = { name, category, document_link, description, is_public };
          stmts.push(
            env.DB.prepare(
              'INSERT INTO master_forms (name, category, document_link, description, is_public) VALUES (?, ?, ?, ?, ?)'
            ).bind(name, category, document_link, description, is_public),
            buildCreateAuditStatement(env, user, 'forms', 'new', newSnapshot, { source: 'import' })
          );
          inserted++;
        }
      }
      if (stmts.length > 0) await executeAtomicImportBatch(env, stmts);
      return ok({ message: `Berhasil memproses ${inserted + updated} data Master Form`, inserted, updated }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
    } catch (e) {
      console.error('Forms import error:', e);
      return error('Gagal mengimport data Master Form', 500, origin);
    }
  }

  if (request.method === 'GET') {
    if (!hasPermission(user, 'forms', 'read')) return forbidden(origin);
  } else if (request.method === 'DELETE') {
    if (!hasPermission(user, 'forms', 'delete')) return forbidden(origin);
  } else {
    if (!hasPermission(user, 'forms', 'write')) return forbidden(origin);
  }

  const idMatch = path.match(/^\/(\d+)$/);

  if (request.method === 'POST' && path === '') {
    let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
    if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
    const { name, category, document_link, description, is_public } = body;
    const fieldErr = firstFieldError(body, [
      ['name', { required: true }],
      ['category', { max: MAX_CATEGORY_LENGTH }],
      ['document_link', { max: MAX_LINK_LENGTH }],
      ['description', { max: MAX_DESCRIPTION_LENGTH }]
    ]);
    if (fieldErr) return error(fieldErr, 400, origin);
    const isPublicFlag = is_public ? 1 : 0;
    const newSnapshot = {
      name,
      category: category || null,
      document_link: document_link || null,
      description: description || null,
      is_public: isPublicFlag
    };
    const batchRes = await env.DB.batch([
      env.DB.prepare('INSERT INTO master_forms (name, category, document_link, description, is_public) VALUES (?, ?, ?, ?, ?)')
        .bind(name, newSnapshot.category, newSnapshot.document_link, newSnapshot.description, isPublicFlag),
      buildCreateAuditStatement(env, user, 'forms', 'new', newSnapshot)
    ]);
    return ok({ id: batchRes?.[0]?.meta?.last_row_id }, 201, origin, undefined, { 'X-Audit-Batched': 'true' });
  }

  if (idMatch) {
    const id = idMatch[1];
    if (request.method === 'GET') {
      const row = await env.DB.prepare('SELECT * FROM master_forms WHERE id = ?').bind(id).first();
      return row ? ok(row, 200, origin) : notFound(origin);
    }
    if (request.method === 'PUT') {
      let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
      if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
      const { name, category, document_link, description, is_public } = body;
      const fieldErr = firstFieldError(body, [
        ['name', {}],
        ['category', { max: MAX_CATEGORY_LENGTH }],
        ['document_link', { max: MAX_LINK_LENGTH }],
        ['description', { max: MAX_DESCRIPTION_LENGTH }]
      ]);
      if (fieldErr) return error(fieldErr, 400, origin);
      const existing = await env.DB.prepare('SELECT * FROM master_forms WHERE id = ?').bind(id).first();
      if (!existing) return notFound(origin);
      const nextIsPublic = is_public !== undefined ? (is_public ? 1 : 0) : null;
      const afterSnapshot = {
        ...existing,
        name: name || existing.name,
        category: category || existing.category,
        document_link: document_link || existing.document_link,
        description: description || existing.description,
        is_public: nextIsPublic !== null ? nextIsPublic : existing.is_public
      };
      await env.DB.batch([
        env.DB.prepare(`UPDATE master_forms SET name = COALESCE(?, name), category = COALESCE(?, category), document_link = COALESCE(?, document_link), description = COALESCE(?, description), is_public = COALESCE(?, is_public), updated_at = datetime('now') WHERE id = ?`)
          .bind(name || null, category || null, document_link || null, description || null, nextIsPublic, id),
        buildAuditStatement(env, user, 'UPDATE', 'forms', id, existing, afterSnapshot, null, 'SUCCESS', true)
      ]);
      return ok({ message: 'Updated' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
    }
    if (request.method === 'DELETE') {
      if (!hasPermission(user, 'forms', 'delete')) return forbidden(origin);
      const existing = await env.DB.prepare('SELECT * FROM master_forms WHERE id = ?').bind(id).first();
      if (!existing) return notFound(origin);
      const delStmt = env.DB.prepare('DELETE FROM master_forms WHERE id = ?').bind(id);
      const { executeAtomicDelete } = await import('../utils/audit.js');
      await executeAtomicDelete(env, user, delStmt, 'forms', id, existing);
      return ok({ message: 'Deleted' }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
    }
  }
  return error('Not found', 404, origin);
}

async function handlePic(request, env, origin) {
  if (request.method === 'GET') {
    // Phase 5.48: PIC reference list is internal data; authentication + read required.
    const user = await authenticate(request, env);
    if (!user) return unauthorized(origin);
    if (!hasPermission(user, 'pic', 'read')) return forbidden(origin);
    const rows = await env.DB.prepare('SELECT * FROM pic_list WHERE is_active = 1 ORDER BY name').all();
    return ok(rows.results, 200, origin);
  }
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'sop', 'admin')) return forbidden(origin);

  if (request.method === 'POST') {
    let body; try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
    if (!body || typeof body !== 'object' || Array.isArray(body)) return error('Payload must be a JSON object', 400, origin);
    const nameErr = fieldError(body.name, 'name', { required: true });
    if (nameErr) return error(nameErr, 400, origin);
    const roleErr = fieldError(body.role, 'role', { max: MAX_CATEGORY_LENGTH });
    if (roleErr) return error(roleErr, 400, origin);
    const newSnapshot = { name: body.name, role: body.role || null };
    const batchRes = await env.DB.batch([
      env.DB.prepare('INSERT OR IGNORE INTO pic_list (name, role) VALUES (?, ?)').bind(body.name, newSnapshot.role),
      buildAuditStatement(env, user, 'CREATE', 'pic', body.name, null, newSnapshot, null, 'SUCCESS', true)
    ]);
    return ok({ id: batchRes?.[0]?.meta?.last_row_id, name: body.name }, 201, origin, undefined, { 'X-Audit-Batched': 'true' });
  }
  return error('Not found', 404, origin);
}

async function handleOptions(request, env, origin) {
  if (request.method !== 'GET') return error('Method not allowed', 405, origin);
  // Phase 5.48: validation enumerations are internal; authentication + read required.
  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);
  if (!hasPermission(user, 'options', 'read')) return forbidden(origin);
  const rows = await env.DB.prepare('SELECT category, value FROM validation_options ORDER BY value').all();
  
  const data = {
    pic: [],
    activity: [],
    quarter: [],
    pkwt: []
  };
  
  (rows.results || []).forEach(r => {
    if (data[r.category]) {
      data[r.category].push(r.value);
    }
  });
  
  return ok(data, 200, origin);
}

async function handleSopBundle(request, env, origin, path) {
  const key = path.replace(/^\//, '');
  if (!key) return error('Key required', 400, origin);

  // Phase 5.48 + 5.66.50: operational bundle keys are a fixed allowlist; staging-only OAT keys
  // (fcms_sop_(fc|sec)_oat_*) are isolated fixtures and strictly rejected outside staging.
  const isOatFixtureKey = isStagingOatSopBundleKey(key, env);
  if (!SOP_BUNDLE_KEYS.has(key) && !isOatFixtureKey) return error('Unknown SOP bundle key', 400, origin);

  if (request.method === 'GET') {
    // Phase 5.48: bundle content is internal reference data; authentication + read required.
    const user = await authenticate(request, env);
    if (!user) return unauthorized(origin);
    if (!hasPermission(user, 'sop', 'read')) return forbidden(origin);
    const row = await env.DB.prepare("SELECT description FROM sop WHERE name = ? AND category = 'sop_bundle' AND deleted_at IS NULL LIMIT 1")
      .bind(key).first();
    if (!row || !row.description) {
      return ok(null, 200, origin);
    }
    try {
      return ok(JSON.parse(row.description), 200, origin);
    } catch {
      return ok(row.description, 200, origin);
    }
  }

  const user = await authenticate(request, env);
  if (!user) return unauthorized(origin);

  if (request.method === 'PUT' || request.method === 'POST') {
    if (!hasPermission(user, 'sop', 'write')) return forbidden(origin);
    let body;
    try { body = await request.json(); } catch { return error('Invalid JSON', 400, origin); }
    // Phase 5.48: SOP bundles are arrays of content items (contract of the SOP editors).
    if (!Array.isArray(body)) return error('SOP bundle payload must be a JSON array', 400, origin);
    if (body.length > SOP_BUNDLE_MAX_ITEMS) {
      return error(`SOP bundle exceeds maximum of ${SOP_BUNDLE_MAX_ITEMS} items`, 400, origin);
    }
    for (const item of body) {
      if (!item || typeof item !== 'object' || Array.isArray(item)) {
        return error('Each SOP bundle item must be a JSON object', 400, origin);
      }
    }
    const serialized = JSON.stringify(body);
    if (serialized.length > SOP_BUNDLE_MAX_CHARS) {
      return error('SOP bundle payload too large', 400, origin);
    }
    const forbiddenKey = findForbiddenKey(body);
    if (forbiddenKey) {
      return error(`SOP bundle contains forbidden field: ${forbiddenKey}`, 400, origin);
    }

    const existing = await env.DB.prepare("SELECT * FROM sop WHERE name = ? AND category = 'sop_bundle' LIMIT 1")
      .bind(key).first();
    if (existing) {
      const wasDeleted = existing.deleted_at !== null && existing.deleted_at !== undefined;
      const mutStmt = env.DB.prepare(
        "UPDATE sop SET description = ?, deleted_at = NULL, updated_at = datetime('now'), row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ?"
      ).bind(serialized, existing.id);
      const auditStmt = buildAuditStatement(
        env,
        user,
        wasDeleted ? 'CREATE' : 'UPDATE',
        'sop',
        key,
        wasDeleted ? null : { id: existing.id, name: existing.name, category: existing.category, description: existing.description },
        { id: existing.id, name: key, category: 'sop_bundle', description: serialized },
        { bundle_key: key, restored: wasDeleted },
        'SUCCESS',
        true
      );
      await env.DB.batch([mutStmt, auditStmt]);
    } else {
      const insertStmt = env.DB.prepare(
        "INSERT INTO sop (name, category, description, created_at, updated_at) VALUES (?, 'sop_bundle', ?, datetime('now'), datetime('now'))"
      ).bind(key, serialized);
      const auditStmt = buildCreateAuditStatement(env, user, 'sop', key, { name: key, category: 'sop_bundle', description: serialized }, { bundle_key: key });
      await env.DB.batch([insertStmt, auditStmt]);
    }
    return ok({ success: true, key }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
  }

  if (request.method === 'DELETE') {
    if (!hasPermission(user, 'sop', 'delete')) return forbidden(origin);
    const existing = await env.DB.prepare("SELECT * FROM sop WHERE name = ? AND category = 'sop_bundle' AND deleted_at IS NULL LIMIT 1").bind(key).first();
    if (!existing) return notFound(origin);
    if (isOatFixtureKey) {
      // Staging-only OAT fixture cleanup: hard-delete so zero residue rows remain in sop table.
      const delStmt = env.DB.prepare(
        "DELETE FROM sop WHERE id = ? AND name = ? AND category = 'sop_bundle'"
      ).bind(existing.id, key);
      const { executeAtomicDelete } = await import('../utils/audit.js');
      await executeAtomicDelete(env, user, delStmt, 'sop', key, existing, { bundle_key: key, oat_fixture_cleanup: true });
      return ok({ success: true, key, hardDeleted: true }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
    }
    const deletedAtIso = new Date().toISOString();
    const delStmt = env.DB.prepare(
      "UPDATE sop SET deleted_at = ?, updated_at = ?, row_version = COALESCE(row_version, 0) + 1, last_sync_source = 'FCMS' WHERE id = ? AND deleted_at IS NULL"
    ).bind(deletedAtIso, deletedAtIso, existing.id);
    const auditStmt = buildSoftDeleteRecordAuditStatement(env, user, 'sop', 'sop', 'id', 'deleted_at', deletedAtIso, existing, { bundle_key: key });
    const batchRes = await env.DB.batch([delStmt, auditStmt]);
    const changes = Number(batchRes?.[0]?.meta?.changes ?? batchRes?.[0]?.changes ?? 0);
    if (changes === 0) return notFound(origin);
    return ok({ success: true, key }, 200, origin, undefined, { 'X-Audit-Batched': 'true' });
  }

  return error('Method not allowed', 405, origin);
}
