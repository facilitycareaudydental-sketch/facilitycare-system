const https = require('https');
https.get('https://docs.google.com/spreadsheets/d/1kdORjAnJ4UB-j_eDMHyiK2TLBgfTIiBMnJOXJcaep0o/gviz/tq?tqx=out:csv&sheet=Master%20Karyawan', res => {
  let data = '';
  res.on('data', c => data+=c);
  res.on('end', () => {
    const rows = data.split('\n').map(r => r.split('\",\"').map(s=>s.replace(/\"/g,'')));
    const rels = rows.filter(r => ['Krishna Aryaan Permana', 'Agung Septiadi', 'Indra Saputro', 'Wariskin'].some(n => r[2] && r[2].toLowerCase().includes(n.toLowerCase())));
    console.log(rels);
  });
});
