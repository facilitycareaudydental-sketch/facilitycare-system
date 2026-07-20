const https = require('https');
const options = {
  hostname: 'facilitycare-system.pages.dev',
  port: 443,
  path: '/api/google-sync',
  method: 'POST',
  headers: {
    'Content-Type': 'application/json'
  }
};
const req = https.request(options, res => {
  let data = '';
  res.on('data', d => data += d);
  res.on('end', () => console.log('Sync Result:', data));
});
req.on('error', e => console.error(e));
req.end();
