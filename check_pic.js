const https = require('https');
https.get('https://facilitycare-system.pages.dev/api/pic', res => {
  let data = '';
  res.on('data', d => data += d);
  res.on('end', () => console.log('PIC:', data));
});
