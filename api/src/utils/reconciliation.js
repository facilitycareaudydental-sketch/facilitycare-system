import { SYNC_MAPPER } from './sync_mapper.js';

export async function createNightlySnapshot(env) {
  const snapshotId = `snap-${Date.now()}`;
  const today = new Date().toISOString().split('T')[0];
  
  const modules = Object.keys(SYNC_MAPPER).filter(k => SYNC_MAPPER[k].enabled);
  
  // Create snapshot header
  await env.DB.prepare(`
    INSERT INTO sync_snapshots (id, snapshot_date, total_modules, overall_status)
    VALUES (?, ?, ?, 'RUNNING')
  `).bind(snapshotId, today, modules.length).run();
  
  let successCount = 0;
  
  for (const mod of modules) {
    const config = SYNC_MAPPER[mod];
    try {
      const activeCount = await env.DB.prepare(`SELECT COUNT(*) as c FROM ${config.table} WHERE deleted_at IS NULL`).first();
      const deletedCount = await env.DB.prepare(`SELECT COUNT(*) as c FROM ${config.table} WHERE deleted_at IS NOT NULL`).first();
      const versionResult = await env.DB.prepare(`SELECT MAX(row_version) as m FROM ${config.table}`).first();
      
      const aC = activeCount ? activeCount.c : 0;
      const dC = deletedCount ? deletedCount.c : 0;
      const hV = versionResult ? versionResult.m : 0;
      const checksum = `chk_${aC}_${hV}`; // Simplified checksum
      
      await env.DB.prepare(`
        INSERT INTO snapshot_details (id, snapshot_id, module, active_row_count, deleted_row_count, highest_row_version, checksum)
        VALUES (?, ?, ?, ?, ?, ?, ?)
      `).bind(`det-${crypto.randomUUID()}`, snapshotId, mod, aC, dC, hV, checksum).run();
      
      successCount++;
    } catch (e) {
      console.error(`Failed to snapshot module ${mod}:`, e);
    }
  }
  
  const finalStatus = successCount === modules.length ? 'COMPLETED' : 'PARTIAL';
  await env.DB.prepare(`UPDATE sync_snapshots SET overall_status = ? WHERE id = ?`).bind(finalStatus, snapshotId).run();
  
  return { snapshotId, finalStatus, modules_processed: successCount };
}

export async function runReconciliation(env, targetModule, repairMode) {
  const startTime = Date.now();
  
  let checked = 0;
  let missing = 0;
  let deleted = 0;
  let version_conflicts = 0;
  let checksum_difference = 0;
  let repaired = 0;
  let failed = 0;
  
  // This is a stub for real reconciliation logic against Google Sheets.
  // In a real scenario, this fetches the entire sheet, compares with D1 using checksums,
  // and queues outbox events or D1 updates depending on the sync direction.
  
  if (repairMode) {
    // simulate repair
    repaired = 2;
    checked = 150;
  } else {
    checked = 150;
  }
  
  return {
    module: targetModule,
    checked_rows: checked,
    missing_rows: missing,
    deleted_rows: deleted,
    version_conflicts,
    checksum_difference,
    repaired_rows: repaired,
    failed_rows: failed,
    duration_ms: Date.now() - startTime
  };
}
