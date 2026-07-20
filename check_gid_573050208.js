const https = require('https');
https.get('https://docs.google.com/spreadsheets/d/1kdORjAnJ4UB-j_eDMHyiK2TLBgfTIiBMnJOXJcaep0o/export?format=csv&gid=573050208', res => {
  let data = '';
  res.on('data', c => data+=c);
  res.on('end', () => console.log(data.split('\n').slice(0, 5).join('\n')));
});
