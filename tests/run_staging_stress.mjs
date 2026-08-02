import fs from 'fs';
import path from 'path';

const API_URL = process.env.FCMS_API_URL || 'https://staging.yourdomain.workers.dev/api';
const BEARER_TOKEN = process.env.FCMS_BEARER_TOKEN || 'YOUR_TOKEN_HERE';
const CONCURRENCY = parseInt(process.env.CONCURRENCY || '50', 10);
const EVENT_COUNT = parseInt(process.env.EVENT_COUNT || '100', 10);

const reportsDir = path.join(process.cwd(), 'reports');
if (!fs.existsSync(reportsDir)) fs.mkdirSync(reportsDir, { recursive: true });

async function injectEvent(id) {
  const start = Date.now();
  try {
    const payload = {
      event_id: `stress-${Date.now()}-${id}`,
      sheet_name: 'Time Line',
      action: 'INSERT',
      payload_version: '1.0',
      data: {
        id: `sch-stress-${id}`,
        project_name: 'Stress Test Project',
        status: 'Aktif',
        row_version: 1
      }
    };
    
    // In staging, we would hit the webhook endpoint to trigger the outbox creation
    const res = await fetch(`${API_URL}/sync/webhook`, {
      method: 'POST',
      headers: { 
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${BEARER_TOKEN}` 
      },
      body: JSON.stringify(payload)
    });
    
    return { id, success: res.ok, status: res.status, latency: Date.now() - start };
  } catch (err) {
    return { id, success: false, error: err.message, latency: Date.now() - start };
  }
}

async function runStressTest() {
  console.log(`Starting Staging Stress Test...`);
  console.log(`API URL: ${API_URL}`);
  console.log(`Events: ${EVENT_COUNT} | Concurrency: ${CONCURRENCY}`);
  
  let completed = 0;
  const results = [];
  const startTotal = Date.now();
  
  // Async queue worker
  const queue = Array.from({ length: EVENT_COUNT }, (_, i) => i);
  const workers = Array.from({ length: CONCURRENCY }, async () => {
    while (queue.length > 0) {
      const id = queue.shift();
      const res = await injectEvent(id);
      results.push(res);
      completed++;
      if (completed % 50 === 0) console.log(`Progress: ${completed} / ${EVENT_COUNT}`);
    }
  });
  
  await Promise.all(workers);
  
  const totalDuration = Date.now() - startTotal;
  const successes = results.filter(r => r.success);
  const avgLatency = results.reduce((acc, r) => acc + r.latency, 0) / EVENT_COUNT;
  
  const report = {
    timestamp: new Date().toISOString(),
    eventCount: EVENT_COUNT,
    concurrency: CONCURRENCY,
    totalDurationMs: totalDuration,
    successRate: (successes.length / EVENT_COUNT) * 100,
    avgLatencyMs: avgLatency,
    results
  };
  
  fs.writeFileSync(path.join(reportsDir, 'stress-test.json'), JSON.stringify(report, null, 2));
  
  const csv = `Timestamp,EventCount,Concurrency,DurationMs,SuccessRate,AvgLatency\n${report.timestamp},${EVENT_COUNT},${CONCURRENCY},${totalDuration},${report.successRate},${avgLatency}`;
  fs.writeFileSync(path.join(reportsDir, 'stress-summary.csv'), csv);
  
  console.log(`\nTest Completed in ${totalDuration}ms`);
  console.log(`Success Rate: ${report.successRate}%`);
  console.log(`Avg Latency: ${avgLatency.toFixed(2)}ms`);
  console.log(`Reports saved to reports/ directory.`);
}

runStressTest().catch(console.error);
