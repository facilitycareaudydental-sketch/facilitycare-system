const fs = require('fs');
const path = require('path');

const srcDir = path.join(__dirname, 'src');
const componentsDir = path.join(srcDir, 'components');
const pagesDir = path.join(srcDir, 'pages');

let issues = [];

function scanDir(dir) {
  if (!fs.existsSync(dir)) return;
  const files = fs.readdirSync(dir);
  for (const f of files) {
    const full = path.join(dir, f);
    if (fs.statSync(full).isDirectory()) {
      scanDir(full);
    } else if (full.endsWith('.js') || full.endsWith('.jsx')) {
      const code = fs.readFileSync(full, 'utf8');
      
      // 1. Audit Buttons
      const btnMatches = code.match(/<button[^>]*>/g);
      if (btnMatches) {
        btnMatches.forEach((b, idx) => {
          if (!b.includes('onClick') && !b.includes('type="submit"') && !b.includes("type='submit'") && !b.includes('disabled')) {
            issues.push(`[${path.basename(full)}] Button without action found (Index ${idx})`);
          }
        });
      }
      
      // 2. Audit Import Excel (if component handles import)
      if (code.includes('import') && code.toLowerCase().includes('excel')) {
         if (!code.includes('progress') && !code.includes('loading')) {
            issues.push(`[${path.basename(full)}] Import function missing Progress Bar / Loading Indicator`);
         }
         if (!code.includes('rollback') && !code.includes('DELETE FROM')) {
            // Very rough check for rollback logic
            // issues.push(`[${path.basename(full)}] Import function might be missing rollback logic`);
         }
      }
      
      // 3. Audit Export Excel
      if (code.includes('export') && code.toLowerCase().includes('excel')) {
         if (!code.includes('xlsx')) {
            issues.push(`[${path.basename(full)}] Export logic found but not using 'xlsx' library properly`);
         }
      }
      
      // 4. Hardcoded Data check
      if (code.match(/dummyData/i) || code.match(/placeholder/i)) {
         issues.push(`[${path.basename(full)}] Potential hardcoded dummy data found`);
      }
    }
  }
}

scanDir(componentsDir);
scanDir(pagesDir);

console.log('--- UI AUDIT RESULTS ---');
if (issues.length > 0) {
  console.log(issues.join('\n'));
} else {
  console.log('All checks passed.');
}
