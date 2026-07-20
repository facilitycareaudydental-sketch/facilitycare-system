const https = require('https');
https.get('https://docs.google.com/spreadsheets/d/1kdORjAnJ4UB-j_eDMHyiK2TLBgfTIiBMnJOXJcaep0o/gviz/tq?tqx=out:csv&sheet=Jadwal%20Reliefer', res => {
  let data = '';
  res.on('data', c => data+=c);
  res.on('end', () => {
    const rows = data.split('\n').map(r => r.split('\",\"').map(s=>s.replace(/\"/g,'')));
    const rel = new Set(rows.slice(1).map(r => r[3]).filter(p => p && p.trim() !== ''));
    console.log(Array.from(rel));
  });
});
