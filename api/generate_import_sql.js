const fs = require('fs');

const raw = fs.readFileSync('data.json', 'utf8');
const parsed = JSON.parse(raw);
const db = parsed.data.database;

let sql = '';

// Helper to escape strings
function escapeSql(val) {
  if (val === null || val === undefined) return 'NULL';
  if (typeof val === 'number') return val;
  if (typeof val === 'boolean') return val ? 1 : 0;
  return "'" + String(val).replace(/'/g, "''") + "'";
}

// Ensure proper insertion order for FKs
const tables = [
  'branches',
  'users',
  'employees',
  'contracts',
  'sp_data',
  'mutasi_data',
  'activity_schedule',
  'issues',
  'one_on_one',
  'training',
  'relievers',
  'inspection_reports',
  'cleaning_reports',
  'fogging_reports',
  'basecamp_reports',
  'supply_requests',
  'master_checklist',
  'master_forms',
  'sop'
];

for (const table of tables) {
  if (!db[table] || db[table].length === 0) continue;
  
  // Clear existing data? Yes, we want to replace it.
  sql += `DELETE FROM ${table};\n`;
  
  const rows = db[table];
  const columns = Object.keys(rows[0]);
  
  for (const row of rows) {
    const values = columns.map(col => escapeSql(row[col])).join(', ');
    const cols = columns.join(', ');
    sql += `INSERT INTO ${table} (${cols}) VALUES (${values});\n`;
  }
}

fs.writeFileSync('import_all.sql', sql);
console.log('Successfully generated import_all.sql with length: ' + sql.length);
