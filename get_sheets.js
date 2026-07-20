const https = require('https');
https.get('https://docs.google.com/spreadsheets/d/1kdORjAnJ4UB-j_eDMHyiK2TLBgfTIiBMnJOXJcaep0o/htmlview', res => {
  let html = '';
  res.on('data', chunk => html += chunk);
  res.on('end', () => {
    const regex = /\{"name":"([^"]+)"/g;
    let match;
    const sheets = new Set();
    while ((match = regex.exec(html)) !== null) {
      sheets.add(match[1]);
    }
    console.log("SHEETS:", Array.from(sheets));
  });
});
