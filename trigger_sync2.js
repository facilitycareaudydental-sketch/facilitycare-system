const https = require('https');
const options = {
  hostname: 'fm-operations-api.facilitycare-audydental.workers.dev',
  port: 443,
  path: '/api/sync/google-sheets',
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
