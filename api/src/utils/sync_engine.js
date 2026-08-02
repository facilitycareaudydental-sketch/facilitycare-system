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

    // Phase 2: Route webhook to handler
    if (sheet_name === 'Master Karyawan') {
      await handleKaryawanWebhook(env, action, data);
    }

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
        await env.DB.prepare(`
          DELETE FROM sync_outbox WHERE id = ?
        `).bind(event.id).run();

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
          
          await env.DB.prepare(`
            UPDATE sync_outbox 
            SET status = 'FAILED', 
                retry_count = ?, 
                last_error = ?, 
                next_retry_at = datetime('now', '+${backoffMinutes} minutes')
            WHERE id = ?
          `).bind(nextRetryCount, err.message, event.id).run();
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
 * Handle incoming webhook data for Master Karyawan
 */
async function handleKaryawanWebhook(env, action, excelData) {
  const { table, data } = mapPayloadToDB('Master Karyawan', excelData);
  
  if (action === 'INSERT') {
    const keys = Object.keys(data).join(', ');
    const placeholders = Object.keys(data).map(() => '?').join(', ');
    const values = Object.values(data);
    
    await env.DB.prepare(
      `INSERT INTO ${table} (${keys}) VALUES (${placeholders})`
    ).bind(...values).run();
    
  } else if (action === 'UPDATE') {
    const id = data['id'];
    delete data['id'];
    const sets = Object.keys(data).map(k => `${k} = ?`).join(', ');
    const values = Object.values(data);
    values.push(id);
    
    await env.DB.prepare(
      `UPDATE ${table} SET ${sets}, updated_at = datetime('now') WHERE id = ?`
    ).bind(...values).run();
    
  } else if (action === 'DELETE') {
    const id = data['id'];
    await env.DB.prepare(`UPDATE ${table} SET status = 'Tidak Aktif', updated_at = datetime('now') WHERE id = ?`).bind(id).run();
  }
}
