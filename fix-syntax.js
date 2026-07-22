const fs = require('fs');
const path = './frontend/src/pages';
const files = fs.readdirSync(path).filter(f => f.endsWith('.js'));
for (const file of files) {
  const p = path + '/' + file;
  let c = fs.readFileSync(p, 'utf8');
  let o = c;
  
  const modMatch = c.match(/moduleName:\s*'([^']+)'/);
  if (modMatch) {
    const mod = modMatch[1];
    c = c.replace(/apiFetch\(\/api\/[a-zA-Z0-9_\\\-]*limit=10000\)/g, "apiFetch('/api/" + mod + "?limit=10000')");
    c = c.replace(/apiFetch\('(\/api\/[a-zA-Z0-9_-]+)\?limit=10000'\)/g, (match, url) => {
      return "apiFetch(`" + url + "${window.location.search ? window.location.search + '&' : '?'}limit=10000`)";
    });
  } else {
    c = c.replace(/apiFetch\(\/api\/([a-zA-Z0-9_\\\-]*)limit=10000\)/g, (match, p1) => {
      return "apiFetch('/api/" + p1.replace(/\\/g, '') + "?limit=10000')";
    });
  }
  
  if (c !== o) {
    fs.writeFileSync(p, c);
    console.log('Fixed', file);
  }
}
