const Database = require('better-sqlite3');
try {
  const db = new Database('database.sqlite');
  console.log(db.prepare(`SELECT COUNT(*) as total FROM inspection_reports t LEFT JOIN branches b ON t.branch_id = b.id WHERE (b.full_name LIKE '%BINTAR%')`).get());
} catch (e) {
  console.error(e);
}
