import { error, ok } from './response.js';
import { mapPayloadToDB } from './sync_mapper.js';

/**
 * Handle incoming Webhook from Google Apps Script (Sheets -> FCMS)
 */
export async function receiveWebhook(request, env) {
  const startTime = Date.now();
  try {
    const authHeader = request.headers.get('Authorization');
    const signatureHeader = request.headers.get('X-FCMS-SIGNATURE');
    
    if (env.SYNC_WEBHOOK_SECRET && authHeader !== `Bearer ${env.SYNC_WEBHOOK_SECRET}`) {
      return error('Unauthorized Webhook', 401);
    }

    const payloadText = await request.text();
    let payload;
    try {
      payload = JSON.parse(payloadText);
    } catch (e) {
      return error('Invalid JSON', 400);
    }

    // HMAC Signature Verification
    if (env.SYNC_WEBHOOK_SECRET && signatureHeader) {
      const encoder = new TextEncoder();
      const key = await crypto.subtle.importKey(
        'raw', encoder.encode(env.SYNC_WEBHOOK_SECRET),
        { name: 'HMAC', hash: 'SHA-256' }, false, ['verify']
      );
      const computedBuffer = await crypto.subtle.sign('HMAC', key, encoder.encode(payloadText));
      const computedHex = Array.from(new Uint8Array(computedBuffer)).map(b => b.toString(16).padStart(2, '0')).join('');
      if (computedHex !== signatureHeader) {
        return error('Invalid Signature', 403);
      }
    }

    const { event_id, sheet_name, action, data, payload_version, correlation_id } = payload;

    // Validation Layer
    if (payload_version !== '1.0') {
      return error('Unsupported payload version', 400);
    }
    if (!event_id || typeof event_id !== 'string') return error('Invalid event_id', 400);
    if (!sheet_name || typeof sheet_name !== 'string') return error('Invalid sheet_name', 400);
    if (!action || !['INSERT', 'UPDATE', 'DELETE'].includes(action)) return error('Invalid action', 400);
    if (!data || typeof data !== 'object') return error('Invalid data payload', 400);

    // 1. Idempotency Check
    const isProcessed = await env.DB.prepare('SELECT 1 FROM sync_idempotency WHERE event_id = ?').bind(event_id).first();
    if (isProcessed) {
      return ok({ message: 'Event already processed', event_id }, 200);
    }

    // 2. Route webhook to generic handler
    await handleGenericWebhook(env, sheet_name, action, data, payload_version, correlation_id);

    // Mark as processed
    await env.DB.prepare('INSERT INTO sync_idempotency (event_id) VALUES (?)').bind(event_id).run();

    const duration = Date.now() - startTime;
    await env.DB.prepare('INSERT INTO sync_metrics (event_id, action, module, webhook_duration_ms, correlation_id) VALUES (?, ?, ?, ?, ?)').bind(event_id, action, sheet_name, duration, correlation_id || null).run();

    return ok({ message: 'Webhook received successfully', event_id }, 200);
  } catch (err) {
    console.error('Webhook processing error:', err);
    return error('Internal webhook error', 500);
  }
}

/**
 * Process the Transactional Outbox (FCMS -> Sheets)
 * This runs via Cloudflare Cron every minute.
 */
export async function processOutbox(env) {
  console.log('Starting Outbox Sweeper...');
  const workerId = crypto.randomUUID();
  const leaseToken = crypto.randomUUID();
  
  try {
    // 0a. Distributed Lock Hardening (Lease Lock)
    await env.DB.prepare("DELETE FROM sync_locks WHERE lease_until < datetime('now')").run();
    
    const lock = await env.DB.prepare(`
      INSERT INTO sync_locks (lock_id, worker_id, lease_token, lease_until, heartbeat_at) 
      VALUES ('outbox_sweeper', ?, ?, datetime('now', '+2 minutes'), datetime('now'))
      ON CONFLICT(lock_id) DO NOTHING
    `).bind(workerId, leaseToken).run();
    
    if (lock.meta.changes === 0) {
      console.log('Another worker holds the outbox lease lock. Exiting.');
      return;
    }

    const extendLease = async () => {
      const res = await env.DB.prepare(`
        UPDATE sync_locks 
        SET lease_until = datetime('now', '+2 minutes'), heartbeat_at = datetime('now'), updated_at = datetime('now')
        WHERE lock_id = 'outbox_sweeper' AND lease_token = ?
      `).bind(leaseToken).run();
      if (res.meta.changes === 0) throw new Error('FENCING_TOKEN_EXPIRED');
    };

    // 0b. Circuit Breaker Check
    const cb = await env.DB.prepare('SELECT * FROM circuit_breaker WHERE id = 1').first();
    if (cb && cb.state === 'OPEN') {
      const now = new Date();
      const nextAttempt = cb.next_attempt_at ? new Date(cb.next_attempt_at) : new Date(0);
      if (now < nextAttempt) {
        console.log('Circuit Breaker is OPEN. Skipping outbox processing.');
        return;
      } else {
        console.log('Circuit Breaker transitioning to HALF_OPEN.');
        await env.DB.prepare(`UPDATE circuit_breaker SET state = 'HALF_OPEN', updated_at = datetime('now') WHERE id = 1`).run();
        await env.DB.prepare(`INSERT INTO audit_logs (action, module, details) VALUES ('CIRCUIT_BREAKER', 'Outbox', 'State changed to HALF_OPEN')`).run();
      }
    }

    // 0c. PROCESSING Recovery (Stuck Events)
    const stuckTimeoutMinutes = env.SYNC_STUCK_TIMEOUT || 5;
    const stuckEvents = await env.DB.prepare(`
      SELECT id FROM sync_outbox 
      WHERE status = 'PROCESSING' AND updated_at < datetime('now', '-${stuckTimeoutMinutes} minutes')
    `).all();
    
    if (stuckEvents.results && stuckEvents.results.length > 0) {
      const ids = stuckEvents.results.map(e => e.id);
      const placeholders = ids.map(() => '?').join(',');
      await env.DB.batch([
        env.DB.prepare(`
          UPDATE sync_outbox 
          SET status = 'PENDING', retry_count = retry_count + 1, updated_at = datetime('now'), last_error = 'Recovered from stuck processing (Timeout/Crash)'
          WHERE id IN (${placeholders})
        `).bind(...ids),
        env.DB.prepare(`INSERT INTO audit_logs (action, module, details) VALUES ('PROCESSING_RECOVERY', 'Outbox', 'Recovered ${ids.length} stuck events')`)
      ]);
      console.log(`Recovered ${ids.length} stuck events to PENDING.`);
    }

    // 1. Fetch pending or failed (but ready for retry) events
    const batchSize = env.SYNC_BATCH_SIZE || 50;
    const query = `
      SELECT * FROM sync_outbox 
      WHERE status = 'PENDING' 
         OR (status = 'FAILED' AND next_retry_at <= datetime('now'))
      ORDER BY created_at ASC
      LIMIT ${Number(batchSize)}
    `;
    const pendingEvents = (await env.DB.prepare(query).all()).results;

    if (!pendingEvents || pendingEvents.length === 0) {
      console.log('No pending sync events.');
      return;
    }

    // Dependency Graph Ranking (Lowest runs first)
    const orderGraph = {
      'Data Cabang': 1,
      'Master Karyawan': 2,
      'Master Kontrak': 3,
      'PIC': 4,
      'Master SOP': 5,
      'Master Checklist': 6,
      'Master Form': 7,
      'Time Line': 8,
      'Jadwal Reliefer': 9
    };
    
    pendingEvents.sort((a, b) => {
      const rankA = orderGraph[a.entity_name] || 99;
      const rankB = orderGraph[b.entity_name] || 99;
      if (rankA !== rankB) return rankA - rankB;
      return new Date(a.created_at) - new Date(b.created_at);
    });

    console.log(`Found ${pendingEvents.length} events to process.`);

    const idsToProcess = pendingEvents.map(e => e.id);
    const placeholders = idsToProcess.map(() => '?').join(',');
    
    await env.DB.prepare(`
      UPDATE sync_outbox 
      SET status = 'PROCESSING' 
      WHERE id IN (${placeholders})
    `).bind(...idsToProcess).run();

    const maxRetries = 5;

    for (const event of pendingEvents) {
      try {
        await extendLease();

        if (env.MOCK_API_ERROR) {
          throw new Error(env.MOCK_API_ERROR);
        }
        
        const queueWaitTime = Date.now() - new Date(event.created_at).getTime();
        await env.DB.batch([
          env.DB.prepare(`DELETE FROM sync_outbox WHERE id = ?`).bind(event.id),
          env.DB.prepare(`INSERT INTO audit_logs (action, module, target_id, details, correlation_id) VALUES ('SYNC_SUCCESS', 'Outbox', ?, ?, ?)`).bind(event.entity_name, `Successfully pushed event ${event.id}`, event.correlation_id),
          env.DB.prepare(`INSERT INTO sync_metrics (event_id, action, module, queue_wait_time_ms, retry_count, correlation_id) VALUES (?, ?, ?, ?, ?, ?)`).bind(event.id, event.action, event.entity_name, queueWaitTime, event.retry_count, event.correlation_id),
          env.DB.prepare(`UPDATE circuit_breaker SET state = 'CLOSED', failure_count = 0 WHERE id = 1 AND (state = 'HALF_OPEN' OR failure_count > 0)`)
        ]);


      } catch (err) {
        console.error(`Failed to process event ${event.id}:`, err);
        
        const nextRetryCount = event.retry_count + 1;
        
        if (nextRetryCount > maxRetries) {
          await env.DB.prepare(`
            UPDATE sync_outbox 
            SET status = 'DEAD_LETTER', last_error = ?, next_retry_at = NULL
            WHERE id = ?
          `).bind(err.message, event.id).run();
          
          await env.DB.prepare(`
            INSERT INTO audit_logs (action, module, target_id, details, correlation_id)
            VALUES ('DEAD_LETTER', ?, ?, ?, ?)
          `).bind('sync_engine', event.entity_name, `Event ${event.id} failed permanently: ${err.message}`, event.correlation_id).run();
        } else {
          const backoffMinutes = [0.5, 2, 10, 60, 120][event.retry_count] || 5;
          
          await env.DB.batch([
            env.DB.prepare(`
              UPDATE sync_outbox 
              SET status = 'FAILED', retry_count = ?, last_error = ?, next_retry_at = datetime('now', '+${backoffMinutes} minutes')
              WHERE id = ?
            `).bind(nextRetryCount, err.message, event.id),
            env.DB.prepare(`INSERT INTO audit_logs (action, module, target_id, details, correlation_id) VALUES ('SYNC_RETRY', 'Outbox', ?, ?, ?)`).bind(event.entity_name, `Event ${event.id} failed, retry ${nextRetryCount}/${maxRetries}. Error: ${err.message}`, event.correlation_id)
          ]);
        }
        
        if (err.message.includes('429') || err.message.includes('Timeout')) {
          await env.DB.prepare(`UPDATE circuit_breaker SET failure_count = failure_count + 1 WHERE id = 1`).run();
          const currentCb = await env.DB.prepare(`SELECT failure_count FROM circuit_breaker WHERE id = 1`).first();
          if (currentCb && currentCb.failure_count >= 3) {
            console.log('Circuit Breaker transitioning to OPEN due to consecutive failures.');
            await env.DB.prepare(`UPDATE circuit_breaker SET state = 'OPEN', last_failure_at = datetime('now'), next_attempt_at = datetime('now', '+5 minutes') WHERE id = 1`).run();
            await env.DB.prepare(`INSERT INTO audit_logs (action, module, details) VALUES ('CIRCUIT_BREAKER', 'Outbox', 'State changed to OPEN after 3 failures')`).run();
            break;
          }
        }
      }

    }

    console.log('Outbox Sweeper finished.');
  } catch (err) {
    console.error('Outbox Sweeper Error:', err);
  } finally {
    try {
      await env.DB.prepare("DELETE FROM sync_locks WHERE lock_id = 'outbox_sweeper' AND lease_token = ?").bind(leaseToken).run();
    } catch (e) {
      console.error('Failed to release lease lock:', e);
    }
  }
}

/**
 * Helper to queue an event to the outbox from other modules.
 */
export function buildOutboxQuery(env, entityName, entityId, action, payload, correlationId = null) {
  const id = crypto.randomUUID();
  return env.DB.prepare(`
    INSERT INTO sync_outbox (id, entity_name, entity_id, action, payload, correlation_id) 
    VALUES (?, ?, ?, ?, ?, ?)
  `).bind(id, entityName, entityId, action, JSON.stringify(payload), correlationId);
}

/**
 * Generic handler for incoming webhook data
 */
async function handleGenericWebhook(env, sheetName, action, excelData, payloadVersion, correlationId) {
  const { table, delete_strategy, soft_delete_col, data } = mapPayloadToDB(sheetName, excelData, payloadVersion);
  const isDryRun = env.SYNC_MODE === 'DRY_RUN';

  if (action === 'INSERT') {
    const keys = Object.keys(data).join(', ');
    const placeholders = Object.keys(data).map(() => '?').join(', ');
    const values = Object.values(data);
    
    if (isDryRun) {
      console.log(`[DRY_RUN] INSERT INTO ${table} (${keys}) VALUES (${values})`);
      return;
    }

    await env.DB.batch([
      env.DB.prepare(`INSERT INTO ${table} (${keys}) VALUES (${placeholders})`).bind(...values),
      env.DB.prepare(`INSERT INTO audit_logs (action, module, details) VALUES ('WEBHOOK_INSERT', ?, 'Inserted from sheets')`).bind(sheetName)
    ]);
    
  } else if (action === 'UPDATE') {
    const id = data['id'];
    if (!id) return; // Need ID for update
    delete data['id'];
    
    const sets = Object.keys(data).map(k => `${k} = ?`).join(', ');
    const values = Object.values(data);
    values.push(id);
    
    // Check for conflict (OCC)
    const existing = await env.DB.prepare(`SELECT row_version FROM ${table} WHERE id = ?`).bind(id).first();
    const incomingVersion = data['row_version'];
    
    if (existing && incomingVersion && incomingVersion < existing.row_version) {
      if (isDryRun) {
        console.log(`[DRY_RUN] CONFLICT for ${sheetName} ID ${id}`);
        return;
      }
      await env.DB.prepare(`INSERT INTO audit_logs (action, module, target_id, details) VALUES ('CONFLICT', ?, ?, 'Ignored webhook due to outdated row_version')`).bind(sheetName, String(id)).run();
      return;
    }
    
    if (isDryRun) {
      console.log(`[DRY_RUN] UPDATE ${table} SET ${sets} WHERE id = ${id}`);
      return;
    }

    await env.DB.batch([
      env.DB.prepare(`UPDATE ${table} SET ${sets}, updated_at = datetime('now') WHERE id = ?`).bind(...values),
      env.DB.prepare(`INSERT INTO audit_logs (action, module, target_id, details) VALUES ('WEBHOOK_UPDATE', ?, ?, 'Updated from sheets')`).bind(sheetName, String(id))
    ]);
    
  } else if (action === 'DELETE') {
    const id = data['id'];
    if (!id) return;
    
    if (delete_strategy === 'SOFT') {
      const delCol = soft_delete_col || 'deleted_at';
      if (isDryRun) {
        console.log(`[DRY_RUN] SOFT DELETE ${table} ID ${id} (${delCol})`);
        return;
      }
      await env.DB.batch([
        env.DB.prepare(`UPDATE ${table} SET ${delCol} = datetime('now'), updated_at = datetime('now') WHERE id = ?`).bind(id),
        env.DB.prepare(`INSERT INTO audit_logs (action, module, target_id, details) VALUES ('WEBHOOK_DELETE', ?, ?, 'Soft deleted from sheets')`).bind(sheetName, String(id))
      ]);
    } else if (delete_strategy === 'HARD') {
      if (isDryRun) {
        console.log(`[DRY_RUN] HARD DELETE ${table} ID ${id}`);
        return;
      }
      await env.DB.batch([
        env.DB.prepare(`DELETE FROM ${table} WHERE id = ?`).bind(id),
        env.DB.prepare(`INSERT INTO audit_logs (action, module, target_id, details) VALUES ('WEBHOOK_DELETE', ?, ?, 'Hard deleted from sheets')`).bind(sheetName, String(id))
      ]);
    } else if (delete_strategy === 'IGNORE') {
      if (isDryRun) {
        console.log(`[DRY_RUN] IGNORE DELETE for ${table} ID ${id}`);
      }
    } else if (delete_strategy === 'ARCHIVE') {
      // Custom archive logic can be placed here. Fallback to soft delete for now.
      const delCol = soft_delete_col || 'deleted_at';
      if (isDryRun) {
        console.log(`[DRY_RUN] ARCHIVE ${table} ID ${id}`);
        return;
      }
      await env.DB.batch([
        env.DB.prepare(`UPDATE ${table} SET ${delCol} = datetime('now'), updated_at = datetime('now') WHERE id = ?`).bind(id),
        env.DB.prepare(`INSERT INTO audit_logs (action, module, target_id, details) VALUES ('WEBHOOK_ARCHIVE', ?, ?, 'Archived from sheets')`).bind(sheetName, String(id))
      ]);
    }
  }
}
