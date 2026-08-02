const fs = require('fs');
const { execSync } = require('child_process');

const sql = fs.readFileSync('../backup_pre_sync.sql', 'utf8');
const lines = sql.split('\n').filter(l => l.trim());
const inserts = lines.filter(l => l.trim().startsWith('INSERT INTO'));

// Tables in correct order (respecting FKs)
const tableOrder = [
  'users', 'roles', 'branches', 'employees', 'contracts',
  'activity_schedule', 'issues', 'one_on_one', 'training',
  'relievers', 'inspection_reports', 'cleaning_reports',
  'fogging_reports', 'basecamp_reports', 'sop',
  'master_checklist', 'master_forms', 'supply_requests',
  'pic_list', 'sp_data', 'mutasi_data'
];

// Group by table
const byTable = {};
for (const stmt of inserts) {
  const match = stmt.match(/INSERT INTO "?(\w+)"?/);
  if (match) {
    const table = match[1];
    if (!byTable[table]) byTable[table] = [];
    byTable[table].push(stmt);
  }
}

// First: clear all relevant tables
const clearOrder = [...tableOrder].reverse();
const clearSql = clearOrder.map(t => `DELETE FROM ${t};`).join('\n');
fs.writeFileSync('restore_clear.sql', clearSql);
console.log('Created restore_clear.sql');

// Then: create inserts in order, chunked by 50 rows each
let chunkIdx = 0;
for (const table of tableOrder) {
  const rows = byTable[table];
  if (!rows || rows.length === 0) continue;

  const chunkSize = 50;
  for (let i = 0; i < rows.length; i += chunkSize) {
    const chunk = rows.slice(i, i + chunkSize);
    const chunkSql = chunk.join('\n') + '\n';
    const filename = `restore_chunk_${String(chunkIdx).padStart(3,'0')}_${table}.sql`;
    fs.writeFileSync(filename, chunkSql);
    chunkIdx++;
  }
}

console.log(`Created ${chunkIdx} chunk files`);
