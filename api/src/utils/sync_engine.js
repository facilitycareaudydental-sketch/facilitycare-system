import { error, ok } from './response.js';
import { mapPayloadToDB } from './sync_mapper.js';

/**
 * Handle incoming Webhook from Google Apps Script (Sheets -> FCMS)
 */
export async function receiveWebhook(request, env) {
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
      // Expected signature is hex, we need to convert it to array buffer or convert crypto result to hex
      // Alternatively, compute our own HMAC and compare hex strings
      const computedBuffer = await crypto.subtle.sign('HMAC', key, encoder.encode(payloadText));
      const computedHex = Array.from(new Uint8Array(computedBuffer)).map(b => b.toString(16).padStart(2, '0')).join('');
      if (computedHex !== signatureHeader) {
        return error('Invalid Signature', 403);
      }
    }

    const { event_id, sheet_name, action, data, payload_version } = payload;

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
    await handleGenericWebhook(env, sheet_name, action, data, payload_version);

    // Mark as processed
    await env.DB.prepare('INSERT INTO sync_idempotency (event_id) VALUES (?)').bind(event_id).run();

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
  try {
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
    
    // Sort pending events based on dependency ranking, then by created_at
    pendingEvents.sort((a, b) => {
      const rankA = orderGraph[a.entity_name] || 99;
      const rankB = orderGraph[b.entity_name] || 99;
      if (rankA !== rankB) return rankA - rankB;
      return new Date(a.created_at) - new Date(b.created_at);
    });

    console.log(`Found ${pendingEvents.length} events to process.`);

    // 2. Pessimistic Lock: Claim these events
    const idsToProcess = pendingEvents.map(e => e.id);
    const placeholders = idsToProcess.map(() => '?').join(',');
    
    await env.DB.prepare(`
      UPDATE sync_outbox 
      SET status = 'PROCESSING' 
      WHERE id IN (${placeholders})
    `).bind(...idsToProcess).run();

    // 3. Process each event (In a real system, we batch by entity_name)
    // For Phase 1, we simulate processing. Phase 2 will implement actual Google Sheets API push.
    
    const maxRetries = 5;

    for (const event of pendingEvents) {
      try {
        // --- SIMULATED GOOGLE SHEETS API PUSH ---
        // In Phase 2: await pushToGoogleSheets(event.entity_name, event.action, JSON.parse(event.payload));
        
        // Success
        await env.DB.batch([
          env.DB.prepare(`DELETE FROM sync_outbox WHERE id = ?`).bind(event.id),
          env.DB.prepare(`INSERT INTO audit_logs (action, module, target_id, details) VALUES ('SYNC_SUCCESS', 'Outbox', ?, ?)`).bind(event.entity_name, `Successfully pushed event ${event.id}`)
        ]);

      } catch (err) {
        console.error(`Failed to process event ${event.id}:`, err);
        
        // Exponential Backoff calculation
        const nextRetryCount = event.retry_count + 1;
        
        if (nextRetryCount > maxRetries) {
          // Dead Letter
          await env.DB.prepare(`
            UPDATE sync_outbox 
            SET status = 'DEAD_LETTER', last_error = ?, next_retry_at = NULL
            WHERE id = ?
          `).bind(err.message, event.id).run();
          
          // Log to audit_logs
          await env.DB.prepare(`
            INSERT INTO audit_logs (action, module, target_id, details)
            VALUES ('DEAD_LETTER', ?, ?, ?)
          `).bind('sync_engine', event.entity_name, `Event ${event.id} failed permanently: ${err.message}`).run();
        } else {
          // Backoff: 30s, 2m, 10m, 60m... (approximated)
          const backoffMinutes = [0.5, 2, 10, 60, 120][event.retry_count] || 5;
          
          await env.DB.batch([
            env.DB.prepare(`
              UPDATE sync_outbox 
              SET status = 'FAILED', retry_count = ?, last_error = ?, next_retry_at = datetime('now', '+${backoffMinutes} minutes')
              WHERE id = ?
            `).bind(nextRetryCount, err.message, event.id),
            env.DB.prepare(`INSERT INTO audit_logs (action, module, target_id, details) VALUES ('SYNC_RETRY', 'Outbox', ?, ?)`).bind(event.entity_name, `Event ${event.id} failed, retry ${nextRetryCount}/${maxRetries}. Error: ${err.message}`)
          ]);
        }
      }
    }

    console.log('Outbox Sweeper finished.');
  } catch (err) {
    console.error('Outbox Sweeper Error:', err);
  }
}

/**
 * Helper to queue an event to the outbox from other modules.
 * This should ideally be run within the same D1 transaction as the main entity update.
 */
export function buildOutboxQuery(env, entityName, entityId, action, payload) {
  const id = crypto.randomUUID();
  return env.DB.prepare(`
    INSERT INTO sync_outbox (id, entity_name, entity_id, action, payload) 
    VALUES (?, ?, ?, ?, ?)
  `).bind(id, entityName, entityId, action, JSON.stringify(payload));
}

/**
 * Generic handler for incoming webhook data
 */
async function handleGenericWebhook(env, sheetName, action, excelData, payloadVersion) {
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
