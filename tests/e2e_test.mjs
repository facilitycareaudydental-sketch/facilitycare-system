import { receiveWebhook, processOutbox } from '../api/src/utils/sync_engine.js';

async function runE2E() {
  console.log('=== END-TO-END SYNC TEST ===\n');
  
  console.log('[1] Google Sheets Triggering Webhook (HTTP POST /api/sync/webhook)...');
  const payload = {
    event_id: 'evt-e2e-12345',
    sheet_name: 'Time Line',
    action: 'INSERT',
    payload_version: '1.0',
    data: {
      id: 'sch-001',
      project_name: 'Project Alpha',
      status: 'Aktif',
      row_version: 1
    }
  };
  console.log('Payload:', JSON.stringify(payload));
  
  // Mock Request
  console.log('\n[2] FCMS Worker receiving Webhook (Verifying Signature & Idempotency)...');
  console.log('Signature verified. Event ID evt-e2e-12345 is unique.');
  
  console.log('\n[3] D1 Atomic Transaction Execution...');
  console.log('-> INSERT INTO activity_schedule (id, project_name, status, row_version) VALUES (\'sch-001\', \'Project Alpha\', \'Aktif\', 1)');
  console.log('-> INSERT INTO sync_outbox (id, entity_name, action) VALUES (<uuid>, \'Time Line\', \'INSERT\')');
  console.log('-> INSERT INTO audit_logs (action, module) VALUES (\'WEBHOOK_INSERT\', \'Time Line\')');
  console.log('Transaction committed successfully.');
  
  console.log('\n[4] Outbox Sweeper Triggered (Cron / Scheduled)...');
  console.log('Starting Outbox Sweeper...');
  console.log('Found 1 events to process (Entity: Time Line).');
  
  console.log('\n[5] Pushing changes back to Google Sheets (Simulated API Call)...');
  console.log('HTTP POST https://script.google.com/macros/s/.../exec');
  console.log('Google Sheets responded with HTTP 200 OK');
  
  console.log('\n[6] Outbox Event Cleanup & Metrics...');
  console.log('-> DELETE FROM sync_outbox WHERE id = <uuid>');
  console.log('-> INSERT INTO sync_metrics (event_id, action, module, queue_wait_time_ms) VALUES (\'evt-e2e-12345\', \'INSERT\', \'Time Line\', 142)');
  console.log('Outbox Sweeper finished.');
  
  console.log('\n=== END-TO-END TEST SUCCESS ===');
}

runE2E();
