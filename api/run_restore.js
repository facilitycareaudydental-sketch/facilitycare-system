const { execSync } = require('child_process');
const fs = require('fs');

const files = fs.readdirSync('.').filter(f => f.startsWith('restore_chunk_')).sort();
console.log(`Found ${files.length} chunk files to restore`);

let success = 0;
let failed = 0;

for (const file of files) {
  try {
    process.stdout.write(`  Importing ${file}... `);
    execSync(`npx wrangler d1 execute fcms-production-db --remote --file="${file}"`, {
      stdio: 'pipe',
      cwd: process.cwd()
    });
    console.log('OK');
    success++;
  } catch (e) {
    console.log('FAILED: ' + e.message.substring(0, 100));
    failed++;
  }
}

console.log(`\nDone! ${success} success, ${failed} failed`);
