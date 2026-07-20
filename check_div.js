const https = require('https');
https.get('https://docs.google.com/spreadsheets/d/1kdORjAnJ4UB-j_eDMHyiK2TLBgfTIiBMnJOXJcaep0o/gviz/tq?tqx=out:csv&sheet=Master%20Karyawan', res => {
  let data = '';
  res.on('data', c => data+=c);
  res.on('end', () => {
    const rows = data.split('\n').map(r => r.split('\",\"').map(s=>s.replace(/\"/g,'')));
    const divs = new Set(rows.slice(1).map(r => r[2]).filter(d => d && d.trim() !== ''));
    console.log(Array.from(divs));
  });
});
