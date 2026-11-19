import { buildBulkHardDeleteAuditStatements } from './audit.js';
import { corsHeaders, ok } from './response.js';

const ACTIVE_IMPORT_LOCKS = new Set();

function canonicalizeValue(val) {
  if (val === null || val === undefined) return null;
  if (Array.isArray(val)) return val.map(canonicalizeValue);
  if (typeof val === 'object') {
    const out = {};
    for (const k of Object.keys(val).sort()) {
      out[k] = canonicalizeValue(val[k]);
    }
    return out;
  }
  return val;
}

export async function computeImportPayloadHash(moduleName, onDuplicate, rows, extraScope = null) {
  const canonical = JSON.stringify({
    module: String(moduleName || ''),
    onDuplicate: String(onDuplicate || 'skip'),
    extraScope: canonicalizeValue(extraScope),
    rows: canonicalizeValue(rows),
  });
  const encoded = new TextEncoder().encode(canonical);
  const digest = await crypto.subtle.digest('SHA-256', encoded);
  return Array.from(new Uint8Array(digest))
    .map((b) => b.toString(16).padStart(2, '0'))
    .join('');
}

export async function ensureImportJournalSchema(env) {
  await env.DB.batch([
    env.DB.prepare(
      `CREATE TABLE IF NOT EXISTS import_jobs (
        id TEXT PRIMARY KEY,
        module TEXT NOT NULL,
        payload_hash TEXT NOT NULL,
        idempotency_key TEXT,
        status TEXT NOT NULL DEFAULT 'IN_PROGRESS',
        total_chunks INTEGER NOT NULL DEFAULT 0,
        committed_chunks INTEGER NOT NULL DEFAULT 0,
        committed_row_index INTEGER NOT NULL DEFAULT -1,
        committed_wipe_chunks INTEGER NOT NULL DEFAULT 0,
        saved_counts_json TEXT,
        meta_json TEXT,
        result_json TEXT,
        error_message TEXT,
        user_id INTEGER,
        created_at TEXT DEFAULT (datetime('now')),
        updated_at TEXT DEFAULT (datetime('now'))
      )`
    ),
    env.DB.prepare(
      `CREATE UNIQUE INDEX IF NOT EXISTS idx_import_jobs_active_lock
       ON import_jobs(module, payload_hash)
       WHERE status = 'IN_PROGRESS'`
    ),
    env.DB.prepare(
      `CREATE UNIQUE INDEX IF NOT EXISTS idx_import_jobs_idempotency
       ON import_jobs(module, idempotency_key)
       WHERE idempotency_key IS NOT NULL`
    ),
  ]);
}

export function extractImportRequestMeta(request) {
  if (!request?.headers) {
    return { idempotencyKey: null, recover: false };
  }
  const rawKey = request.headers.get('Idempotency-Key') || request.headers.get('X-Idempotency-Key');
  const idempotencyKey = rawKey ? rawKey.trim().slice(0, 128) : null;
  const recover = (request.headers.get('X-Import-Recover') || '').trim().toLowerCase() === 'true';
  return { idempotencyKey: idempotencyKey || null, recover };
}

function conflictResponse(message, code, job, origin, environment) {
  return new Response(
    JSON.stringify({
      success: false,
      error: message,
      code,
      importJobId: job?.id || null,
      status: job?.status || 'IN_PROGRESS',
      committedChunks: Number(job?.committed_chunks || 0),
      totalChunks: Number(job?.total_chunks || 0),
    }),
    {
      status: 409,
      headers: {
        'Content-Type': 'application/json',
        ...corsHeaders(origin),
      },
    }
  );
}

function partialFailureResponse(err, jobId, committedChunks, totalChunks, committedRowIndex, origin) {
  return new Response(
    JSON.stringify({
      success: false,
      error: err?.message || 'Import partially failed and recorded in import journal',
      code: 'IMPORT_PARTIAL_FAILED',
      importJobId: jobId,
      committedChunks,
      totalChunks,
      committedRowIndex,
      recoverable: true,
    }),
    {
      status: 500,
      headers: {
        'Content-Type': 'application/json',
        ...corsHeaders(origin),
      },
    }
  );
}

/**
 * Packs grouped statements into chunks where NO group spans across two batches
 * and each chunk preserves space for 1 trailing journal checkpoint statement (maxStmtsPerChunk <= 76).
 */
export function packStatementGroups(groups, maxStmtsPerChunk = 76) {
  if (!Array.isArray(groups) || groups.length === 0) return [];
  const limit = Math.max(2, Math.min(76, maxStmtsPerChunk));
  const chunks = [];
  let currentStmts = [];
  let maxRowIndex = -1;
  let lastCounts = null;

  for (const grp of groups) {
    const stmts = Array.isArray(grp?.stmts) ? grp.stmts : [];
    if (stmts.length === 0) continue;
    if (stmts.length > limit) {
      throw new Error(`Import group exceeds chunk limit (${stmts.length} > ${limit})`);
    }
    if (currentStmts.length + stmts.length > limit && currentStmts.length > 0) {
      chunks.push({
        stmts: currentStmts,
        endRowIndex: maxRowIndex,
        counts: lastCounts,
      });
      currentStmts = [];
    }
    currentStmts.push(...stmts);
    if (typeof grp.rowIndex === 'number' && grp.rowIndex > maxRowIndex) {
      maxRowIndex = grp.rowIndex;
    }
    if (grp.counts) {
      lastCounts = { ...grp.counts };
    }
  }

  if (currentStmts.length > 0) {
    chunks.push({
      stmts: currentStmts,
      endRowIndex: maxRowIndex,
      counts: lastCounts,
    });
  }
  return chunks;
}

/**
 * Builds staged DELETE + per-record DELETE audit chunks for wipe-and-replace flows
 * after all replacement insert chunks have committed.
 */
export function buildStagedWipeGroups(env, user, tableName, moduleName, oldRows, pkCol = 'id', details = null, idsPerChunk = 48) {
  if (!Array.isArray(oldRows) || oldRows.length === 0) return [];
  const safeIdsPerChunk = Math.max(1, Math.min(64, idsPerChunk));
  const groups = [];
  for (let i = 0; i < oldRows.length; i += safeIdsPerChunk) {
    const slice = oldRows.slice(i, i + safeIdsPerChunk);
    const ids = slice.map((r) => r[pkCol]);
    const placeholders = ids.map(() => '?').join(', ');
    const delStmt = env.DB.prepare(`DELETE FROM ${tableName} WHERE ${pkCol} IN (${placeholders})`).bind(...ids);
    const auditStmts = buildBulkHardDeleteAuditStatements(env, user, moduleName, slice, pkCol, details);
    groups.push({ stmts: [delStmt, ...auditStmts] });
  }
  return groups;
}

/**
 * Executes an import through the durable import_jobs journal:
 * - Enforces concurrency lock (409 on active identical import)
 * - Supports instant idempotent replay when Idempotency-Key already COMPLETED
 * - Resumes PARTIAL_FAILED or recovered IN_PROGRESS jobs strictly from committed_row_index / committed_wipe_chunks
 * - Bundles journal checkpoint inside the exact same env.DB.batch() as each data+audit chunk
 * - Runs staged wipeChunks strictly AFTER all insertChunks have committed
 */
export async function executeJournaledImport(env, user, origin, options) {
  const {
    module: moduleName,
    rows,
    onDuplicate = 'skip',
    extraScope = null,
    importMeta = {},
    buildPlan,
    buildFinalResult,
    maxStmtsPerBatch = 76,
    extraHeaders = {},
  } = options;

  const { idempotencyKey = null, recover = false } = importMeta || {};
  await ensureImportJournalSchema(env);

  const payloadHash = await computeImportPayloadHash(moduleName, onDuplicate, rows, extraScope);
  const lockKey = `${moduleName}:${payloadHash}`;

  // 1. Check in-process active lock immediately (catches microtask-level concurrency)
  if (ACTIVE_IMPORT_LOCKS.has(lockKey)) {
    return conflictResponse(
      `Import untuk modul ${moduleName} sedang berjalan`,
      'IMPORT_IN_PROGRESS',
      null,
      origin
    );
  }

  // 2. Look up idempotency key first (if supplied)
  let existingJob = null;
  if (idempotencyKey) {
    const byKey = await env.DB.prepare(
      `SELECT * FROM import_jobs WHERE module = ? AND idempotency_key = ? LIMIT 1`
    ).bind(moduleName, idempotencyKey).first();

    if (byKey) {
      if (byKey.payload_hash !== payloadHash) {
        return conflictResponse(
          'Idempotency-Key telah digunakan untuk payload yang berbeda',
          'IDEMPOTENCY_KEY_MISMATCH',
          byKey,
          origin
        );
      }
      if (byKey.status === 'COMPLETED' && byKey.result_json) {
        const replayData = JSON.parse(byKey.result_json);
        return ok({ ...replayData, idempotentReplay: true, importJobId: byKey.id }, 200, origin, {
          'X-Audit-Batched': 'true',
          'X-Import-Idempotent-Replay': 'true',
          ...extraHeaders,
        });
      }
      existingJob = byKey;
    }
  }

  // 3. Otherwise look up active / partial-failed job for same module + payload hash
  if (!existingJob) {
    existingJob = await env.DB.prepare(
      `SELECT * FROM import_jobs
       WHERE module = ? AND payload_hash = ? AND status IN ('IN_PROGRESS', 'PARTIAL_FAILED')
       ORDER BY created_at DESC LIMIT 1`
    ).bind(moduleName, payloadHash).first();
  }

  if (existingJob?.status === 'IN_PROGRESS' && !recover) {
    return conflictResponse(
      `Import untuk modul ${moduleName} sedang berjalan`,
      'IMPORT_IN_PROGRESS',
      existingJob,
      origin
    );
  }

  ACTIVE_IMPORT_LOCKS.add(lockKey);
  let jobId = existingJob?.id || crypto.randomUUID();
  let committedChunks = Number(existingJob?.committed_chunks || 0);
  let committedRowIndex = existingJob ? Number(existingJob.committed_row_index ?? -1) : -1;
  let committedWipeChunks = existingJob ? Number(existingJob.committed_wipe_chunks || 0) : 0;
  let totalChunksPlanned = Number(existingJob?.total_chunks || 0);

  try {
    const initialCounts = existingJob?.saved_counts_json
      ? JSON.parse(existingJob.saved_counts_json)
      : { inserted: 0, updated: 0, skipped: 0 };
    const initialMeta = existingJob?.meta_json ? JSON.parse(existingJob.meta_json) : null;

    // 4. Build remaining operations strictly after committedRowIndex
    const plan = await buildPlan({
      startRowIndex: committedRowIndex + 1,
      initialCounts,
      initialMeta,
      isResuming: Boolean(existingJob),
    });

    const insertGroups = Array.isArray(plan?.insertGroups) ? plan.insertGroups : [];
    const wipeGroups = Array.isArray(plan?.wipeGroups) ? plan.wipeGroups : [];
    const metaToPersist = plan?.meta ?? initialMeta ?? null;

    const insertChunks = packStatementGroups(insertGroups, maxStmtsPerBatch);
    const allWipeChunks = packStatementGroups(wipeGroups, maxStmtsPerBatch);
    const remainingWipeChunks = plan?.liveWipePlan ? allWipeChunks : allWipeChunks.slice(committedWipeChunks);

    if (!existingJob) {
      totalChunksPlanned = insertChunks.length + allWipeChunks.length;
      try {
        await env.DB.prepare(
          `INSERT INTO import_jobs (id, module, payload_hash, idempotency_key, status, total_chunks, committed_chunks, committed_row_index, committed_wipe_chunks, saved_counts_json, meta_json, user_id)
           VALUES (?, ?, ?, ?, 'IN_PROGRESS', ?, 0, -1, 0, ?, ?, ?)`
        ).bind(
          jobId,
          moduleName,
          payloadHash,
          idempotencyKey,
          totalChunksPlanned,
          JSON.stringify(initialCounts),
          metaToPersist ? JSON.stringify(metaToPersist) : null,
          user?.id ?? null
        ).run();
      } catch (insertErr) {
        const msg = String(insertErr?.message || insertErr);
        if (msg.includes('UNIQUE constraint failed')) {
          return conflictResponse(
            `Import untuk modul ${moduleName} sedang berjalan`,
            'IMPORT_IN_PROGRESS',
            null,
            origin
          );
        }
        throw insertErr;
      }
    } else {
      // On resume, recompute total chunks based on already committed + remaining chunks
      totalChunksPlanned = committedChunks + insertChunks.length + remainingWipeChunks.length;
      await env.DB.prepare(
        `UPDATE import_jobs
         SET status = 'IN_PROGRESS',
             total_chunks = ?,
             error_message = NULL,
             updated_at = datetime('now')
         WHERE id = ?`
      ).bind(totalChunksPlanned, jobId).run();
    }

    const totalRemaining = insertChunks.length + remainingWipeChunks.length;
    const finalCounts = plan?.finalCounts || initialCounts;
    const finalResult = buildFinalResult(finalCounts, plan);

    if (totalRemaining === 0) {
      await env.DB.prepare(
        `UPDATE import_jobs
         SET status = 'COMPLETED',
             total_chunks = ?,
             committed_chunks = ?,
             committed_row_index = ?,
             saved_counts_json = ?,
             result_json = ?,
             error_message = NULL,
             updated_at = datetime('now')
         WHERE id = ?`
      ).bind(
        totalChunksPlanned,
        committedChunks,
        Math.max(committedRowIndex, rows.length - 1),
        JSON.stringify(finalCounts),
        JSON.stringify(finalResult),
        jobId
      ).run();

      return ok({ ...finalResult, importJobId: jobId }, 200, origin, {
        'X-Audit-Batched': 'true',
        ...extraHeaders,
      });
    }

    // 5. Execute remaining INSERT/UPDATE chunks first, then WIPE chunks.
    // Every chunk appends the journal checkpoint inside the EXACT same env.DB.batch().
    let execIdx = 0;

    for (const chunk of insertChunks) {
      execIdx++;
      committedChunks++;
      const nextRowIdx = chunk.endRowIndex >= 0 ? chunk.endRowIndex : committedRowIndex;
      const chunkCounts = chunk.counts || finalCounts;
      const isLastOverall = execIdx === totalRemaining;
      const progressStmt = isLastOverall
        ? env.DB.prepare(
            `UPDATE import_jobs
             SET status = 'COMPLETED',
                 committed_chunks = ?,
                 committed_row_index = ?,
                 saved_counts_json = ?,
                 result_json = ?,
                 error_message = NULL,
                 updated_at = datetime('now')
             WHERE id = ?`
          ).bind(
            committedChunks,
            Math.max(nextRowIdx, rows.length - 1),
            JSON.stringify(finalCounts),
            JSON.stringify(finalResult),
            jobId
          )
        : env.DB.prepare(
            `UPDATE import_jobs
             SET committed_chunks = ?,
                 committed_row_index = ?,
                 saved_counts_json = ?,
                 updated_at = datetime('now')
             WHERE id = ?`
          ).bind(committedChunks, nextRowIdx, JSON.stringify(chunkCounts), jobId);

      try {
        await env.DB.batch([...chunk.stmts, progressStmt]);
        committedRowIndex = nextRowIdx;
      } catch (chunkErr) {
        committedChunks--;
        // Do NOT overwrite committed_chunks / committed_row_index / completed status:
        // - If D1 rolled back the batch, import_jobs already holds the previous committed checkpoint.
        // - If D1 committed before a post-commit transport error/crash, progressStmt inside the batch
        //   already recorded the new checkpoint (or COMPLETED status).
        await env.DB.prepare(
          `UPDATE import_jobs
           SET status = CASE WHEN status = 'COMPLETED' THEN 'COMPLETED' ELSE 'PARTIAL_FAILED' END,
               error_message = ?,
               updated_at = datetime('now')
           WHERE id = ?`
        ).bind(
          String(chunkErr?.message || chunkErr).slice(0, 500),
          jobId
        ).run();
        const latestJob = await env.DB.prepare(
          `SELECT committed_chunks, committed_row_index FROM import_jobs WHERE id = ?`
        ).bind(jobId).first();
        return partialFailureResponse(
          chunkErr,
          jobId,
          Number(latestJob?.committed_chunks ?? committedChunks),
          totalChunksPlanned,
          Number(latestJob?.committed_row_index ?? committedRowIndex),
          origin
        );
      }
    }

    // Mark all rows processed before entering staged wipe phase
    if (committedRowIndex < rows.length - 1 && remainingWipeChunks.length > 0) {
      committedRowIndex = rows.length - 1;
    }

    for (const wipeChunk of remainingWipeChunks) {
      execIdx++;
      committedChunks++;
      committedWipeChunks++;
      const isLastOverall = execIdx === totalRemaining;
      const progressStmt = isLastOverall
        ? env.DB.prepare(
            `UPDATE import_jobs
             SET status = 'COMPLETED',
                 committed_chunks = ?,
                 committed_row_index = ?,
                 committed_wipe_chunks = ?,
                 saved_counts_json = ?,
                 result_json = ?,
                 error_message = NULL,
                 updated_at = datetime('now')
             WHERE id = ?`
          ).bind(
            committedChunks,
            committedRowIndex,
            committedWipeChunks,
            JSON.stringify(finalCounts),
            JSON.stringify(finalResult),
            jobId
          )
        : env.DB.prepare(
            `UPDATE import_jobs
             SET committed_chunks = ?,
                 committed_row_index = ?,
                 committed_wipe_chunks = ?,
                 updated_at = datetime('now')
             WHERE id = ?`
          ).bind(committedChunks, committedRowIndex, committedWipeChunks, jobId);

      try {
        await env.DB.batch([...wipeChunk.stmts, progressStmt]);
      } catch (wipeErr) {
        committedChunks--;
        committedWipeChunks--;
        await env.DB.prepare(
          `UPDATE import_jobs
           SET status = CASE WHEN status = 'COMPLETED' THEN 'COMPLETED' ELSE 'PARTIAL_FAILED' END,
               error_message = ?,
               updated_at = datetime('now')
           WHERE id = ?`
        ).bind(
          String(wipeErr?.message || wipeErr).slice(0, 500),
          jobId
        ).run();
        const latestJob = await env.DB.prepare(
          `SELECT committed_chunks, committed_row_index FROM import_jobs WHERE id = ?`
        ).bind(jobId).first();
        return partialFailureResponse(
          wipeErr,
          jobId,
          Number(latestJob?.committed_chunks ?? committedChunks),
          totalChunksPlanned,
          Number(latestJob?.committed_row_index ?? committedRowIndex),
          origin
        );
      }
    }

    return ok({ ...finalResult, importJobId: jobId }, 200, origin, {
      'X-Audit-Batched': 'true',
      ...extraHeaders,
    });
  } catch (err) {
    // Failure during planning or non-chunk execution step
    try {
      await env.DB.prepare(
        `UPDATE import_jobs
         SET status = 'PARTIAL_FAILED',
             error_message = ?,
             updated_at = datetime('now')
         WHERE id = ? AND status = 'IN_PROGRESS'`
      ).bind(String(err?.message || err).slice(0, 500), jobId).run();
    } catch { /* ignore secondary failure */ }
    throw err;
  } finally {
    ACTIVE_IMPORT_LOCKS.delete(lockKey);
  }
}
