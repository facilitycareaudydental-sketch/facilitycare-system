const fs = require('fs');
const html = fs.readFileSync('htmlview_full.html', 'utf8');

const regex = /items\.push\(\{name:\s*"([^"]+)",\s*pageUrl:\s*"([^"]+)"/g;
let match;
const sheets = {};
while ((match = regex.exec(html)) !== null) {
  const url = match[2].replace(/\\\//g, '/');
  const gidMatch = url.match(/gid=(\d+)/);
  if (gidMatch) {
    sheets[match[1]] = gidMatch[1];
  }
}
console.log(JSON.stringify(sheets, null, 2));
