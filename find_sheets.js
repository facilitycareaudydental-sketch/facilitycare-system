const fs = require('fs');
const html = fs.readFileSync('htmlview_full.html', 'utf8');
const regex = /<li id="sheet-button-([^"]+)"><a href="#" onclick="switchToSheet\('[^']+'\);return false;">(.*?)<\/a><\/li>/g;
let match;
const sheets = [];
while ((match = regex.exec(html)) !== null) {
  sheets.push({ gid: match[1], name: match[2] });
}
console.log(sheets);
if (sheets.length === 0) {
    const listRegex = /<li id="sheet-button(.*?)<\/li>/g;
    const items = html.match(listRegex);
    console.log(items);
}
