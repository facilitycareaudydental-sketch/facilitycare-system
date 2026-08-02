const fs = require('fs');
const { execSync } = require('child_process');

// Read the backup SQL
const sql = fs.readFileSync('../backup_pre_sync.sql', 'utf8');
const lines = sql.split('\n').filter(l => l.trim());

console.log(`Total lines: ${lines.length}`);

// Only do INSERT statements, skip CREATE TABLE (schema already exists)
const inserts = lines.filter(l => l.trim().startsWith('INSERT INTO'));
console.log(`Total INSERT statements: ${inserts.length}`);

// Group by table name
const byTable = {};
for (const stmt of inserts) {
  const match = stmt.match(/INSERT INTO "?(\w+)"?/);
  if (match) {
    const table = match[1];
    if (!byTable[table]) byTable[table] = [];
    byTable[table].push(stmt);
  }
}

console.log('\nData per table:');
for (const [table, rows] of Object.entries(byTable)) {
  console.log(`  ${table}: ${rows.length} rows`);
}
