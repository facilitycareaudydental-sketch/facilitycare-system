const { execSync } = require('child_process');
const fs = require('fs');

// Failed files - retry with OR IGNORE
const failedFiles = [
  'restore_chunk_004_employees.sql',
  'restore_chunk_005_employees.sql',
  'restore_chunk_008_contracts.sql',
  'restore_chunk_009_contracts.sql',
  'restore_chunk_035_pic_list.sql'
];

let success = 0;
let failed = 0;

for (const file of failedFiles) {
  try {
    // Read original SQL and convert INSERT INTO to INSERT OR IGNORE INTO
    const original = fs.readFileSync(file, 'utf8');
    const fixed = original.replace(/INSERT INTO/g, 'INSERT OR IGNORE INTO');
    const fixedFile = file.replace('.sql', '_fixed.sql');
    fs.writeFileSync(fixedFile, fixed);

    process.stdout.write(`  Retrying ${file}... `);
    execSync(`npx wrangler d1 execute fcms-production-db --remote --file="${fixedFile}"`, {
      stdio: 'pipe',
      cwd: process.cwd()
    });
    console.log('OK');
    success++;
  } catch (e) {
    console.log('FAILED: ' + e.stderr?.toString().substring(0, 200));
    failed++;
  }
}

console.log(`\nRetry done! ${success} success, ${failed} still failed`);
