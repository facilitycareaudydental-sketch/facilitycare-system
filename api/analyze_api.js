const fs = require('fs');
const path = require('path');
const dir = 'C:\\Users\\Facility Care\\.gemini\\antigravity\\scratch\\facilitycare-system\\api\\src\\routes';
const files = fs.readdirSync(dir).filter(f => f.endsWith('.js'));
let report = [];

for (const file of files) {
  const content = fs.readFileSync(path.join(dir, file), 'utf8');
  let issues = [];
  
  if (!content.includes('authenticate(')) issues.push('Missing authentication');
  
  if (content.includes("request.method === 'POST'") && !content.includes("try { body = await request.json()")) {
    issues.push('Missing try/catch on POST body');
  }
  
  if (content.match(/bind\(\)\.run/)) issues.push('Empty bind().run()');
  
  report.push(file + ': ' + (issues.length ? issues.join(', ') : 'OK'));
}
console.log(report.join('\n'));
