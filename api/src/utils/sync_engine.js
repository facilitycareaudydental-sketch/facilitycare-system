import { error, ok } from './response.js';

/**
 * Handle incoming Webhook from Google Apps Script (Sheets -> FCMS)
 */
export async function receiveWebhook(request, env) {
  try {
    const authHeader = request.headers.get('Authorization');
    // Basic shared secret check for webhook (You should set SYNC_WEBHOOK_SECRET in Cloudflare)
    if (env.SYNC_WEBHOOK_SECRET && authHeader !== `Bearer ${env.SYNC_WEBHOOK_SECRET}`) {
      return error('Unauthorized Webhook', 401);
    }

    const payload = await request.json();
    const { event_id, sheet_name, action, data } = payload;

    if (!event_id || !sheet_name || !action || !data) {
      return error('Invalid payload format', 400);
    }

    // 1. Idempotency Check
    const isProcessed = await env.DB.prepare('SELECT 1 FROM sync_idempotency WHERE event_id = ?').bind(event_id).first();
    if (isProcessed) {
      return ok({ message: 'Event already processed', event_id }, 200);
    }

    // 2. We don't implement the business logic here yet (that's Phase 2 onwards)
    // For Phase 1, we just acknowledge the webhook and mark idempotency.
    
    // In Phase 2, we will route `sheet_name` to the correct entity handler.
    // e.g. if (sheet_name === 'Master Karyawan') await handleKaryawanWebhook(...)

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
    const query = `
      SELECT * FROM sync_outbox 
      WHERE status = 'PENDING' 
         OR (status = 'FAILED' AND next_retry_at <= datetime('now'))
      ORDER BY created_at ASC
      LIMIT 50
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
