const fs = require('fs');
const path = './frontend/src/pages';
const files = fs.readdirSync(path).filter(f => f.endsWith('.js'));
let fixed = 0;

for (const file of files) {
  const p = path + '/' + file;
  let content = fs.readFileSync(p, 'utf8');
  let original = content;

  content = content.replace(/apiFetch\(\/api\/([a-zA-Z0-9_-]+)limit=10000\)/g, "apiFetch(\/api/$1\\limit=10000\)");

  if (content !== original) {
    fs.writeFileSync(p, content, 'utf8');
    fixed++;
    console.log('Fixed', file);
  }
}
console.log('Total fixed:', fixed);
