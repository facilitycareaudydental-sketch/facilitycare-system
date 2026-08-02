import { processOutbox } from '../api/src/utils/sync_engine.js';

class MockD1CB {
  constructor() {
    this.outbox = [];
    this.circuit_breaker = { id: 1, state: 'CLOSED', failure_count: 0, next_attempt_at: null, last_failure_at: null };
    this.audit_logs = [];
  }
  
  prepare(query) {
    const db = this;
    return {
      bind: (...args) => {
        return {
          query,
          args,
          all: async () => {
            if (query.includes('FROM sync_outbox')) {
              return { results: db.outbox.filter(e => e.status === 'PENDING' || e.status === 'FAILED').slice(0, 10) };
            }
            return { results: [] };
          },
          first: async () => {
            if (query.includes('FROM circuit_breaker')) return db.circuit_breaker;
            return null;
          },
          run: async () => {
            if (query.includes('UPDATE circuit_breaker SET state = \'OPEN\'')) {
               db.circuit_breaker.state = 'OPEN';
               db.circuit_breaker.next_attempt_at = new Date(Date.now() + 5 * 60000).toISOString();
               console.log('[Mock DB] Circuit Breaker -> OPEN');
            }
            if (query.includes('UPDATE circuit_breaker SET failure_count')) {
               db.circuit_breaker.failure_count++;
               console.log(`[Mock DB] Failure Count = ${db.circuit_breaker.failure_count}`);
            }
            if (query.includes('UPDATE circuit_breaker SET state = \'HALF_OPEN\'')) {
               db.circuit_breaker.state = 'HALF_OPEN';
               console.log('[Mock DB] Circuit Breaker -> HALF_OPEN');
            }
            if (query.includes('UPDATE sync_outbox SET status = \'FAILED\'')) {
               const id = args[3]; // The ID is the 4th parameter in the backoff update
               const e = db.outbox.find(x => x.id === id);
               if (e) {
                  e.status = 'FAILED';
                  e.retry_count = args[0];
               }
            }
            return { success: true };
          }
        };
      },
      all: async () => {
        if (query.includes('FROM sync_outbox')) {
          return { results: db.outbox.filter(e => e.status === 'PENDING' || e.status === 'FAILED').slice(0, 10) };
        }
        return { results: [] };
      },
      first: async () => {
        if (query.includes('FROM circuit_breaker')) return db.circuit_breaker;
        return null;
      },
      run: async () => {
        if (query.includes('UPDATE circuit_breaker SET state = \'HALF_OPEN\'')) {
          db.circuit_breaker.state = 'HALF_OPEN';
          console.log('[Mock DB] Circuit Breaker -> HALF_OPEN');
        }
        if (query.includes('UPDATE circuit_breaker SET failure_count = failure_count + 1')) {
           db.circuit_breaker.failure_count++;
           console.log(`[Mock DB] Failure Count = ${db.circuit_breaker.failure_count}`);
        }
      }
    };
  }
  
  async batch(queries) {
    for (const q of queries) {
      if (q.query && q.query.includes('DELETE FROM sync_outbox')) {
         const id = q.args[0];
         this.outbox = this.outbox.filter(x => x.id !== id);
      }
      if (q.query && q.query.includes('UPDATE sync_outbox SET status = \'FAILED\'')) {
         const id = q.args[3]; // fourth arg is ID usually
         const e = this.outbox.find(x => x.id === id);
         if (e) e.status = 'FAILED';
      }
      if (q.query && q.query.includes('UPDATE circuit_breaker SET state = \'CLOSED\'')) {
         this.circuit_breaker.state = 'CLOSED';
         this.circuit_breaker.failure_count = 0;
         console.log('[Mock DB] Circuit Breaker -> CLOSED (Recovery)');
      }
    }
    return [{ success: true }];
  }
}

async function runTest() {
  console.log(`\n=== CIRCUIT BREAKER TEST ===`);
  const db = new MockD1CB();
  const env = { DB: db, SYNC_BATCH_SIZE: 1, MOCK_API_ERROR: '429 Too Many Requests' };
  
  // Create 5 events
  for (let i = 0; i < 5; i++) {
    db.outbox.push({
      id: `evt-${i}`,
      entity_name: 'Time Line',
      action: 'INSERT',
      status: 'PENDING',
      retry_count: 0,
      created_at: new Date().toISOString()
    });
  }

  console.log('\n--- Attempt 1 (Failure 1) ---');
  await processOutbox(env);
  
  console.log('\n--- Attempt 2 (Failure 2) ---');
  await processOutbox(env);
  
  console.log('\n--- Attempt 3 (Failure 3 -> OPEN) ---');
  await processOutbox(env);
  
  console.log('\n--- Attempt 4 (Should skip due to OPEN) ---');
  await processOutbox(env);
  
  console.log('\n--- Fast-forward 5 minutes (Transition to HALF_OPEN) ---');
  db.circuit_breaker.next_attempt_at = new Date(Date.now() - 1000).toISOString();
  await processOutbox(env);
  
  console.log('\n--- Resolve Google API Issue (Recovery to CLOSED) ---');
  env.MOCK_API_ERROR = null; // No more errors
  await processOutbox(env);
}

runTest();
