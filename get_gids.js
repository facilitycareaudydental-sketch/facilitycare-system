const fs = require('fs');
const https = require('https');

https.get('https://docs.google.com/spreadsheets/d/1kdORjAnJ4UB-j_eDMHyiK2TLBgfTIiBMnJOXJcaep0o/htmlview', res => {
  let html = '';
  res.on('data', chunk => html += chunk);
  res.on('end', () => {
    // Look for grid names and gids in the initial HTML or JS payload
    const matches = html.match(/"name":"([^"]+)".*?"gid":"?(\d+)"?/g);
    if (matches) {
      console.log("MATCHES:", matches);
    } else {
      console.log("No matches found. Trying another regex.");
      const matches2 = html.match(/\{[^{]*"name":"[^"]+"[^{]*"gid":"?\d+"?[^}]*\}/g);
      console.log("MATCHES 2:", matches2 ? matches2.length : 0);
    }
    fs.writeFileSync('htmlview_full.html', html);
  });
});
