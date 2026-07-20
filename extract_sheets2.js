const fs = require('fs');
const html = fs.readFileSync('htmlview.html', 'utf8');
const regex = /<li[^>]*><a[^>]*>(.*?)<\/a><\/li>/g;
let match;
const sheets = new Set();
while ((match = regex.exec(html)) !== null) {
  sheets.add(match[1]);
}
console.log("SHEETS:", Array.from(sheets));
