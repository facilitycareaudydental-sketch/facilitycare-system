const https = require('https');

const req = https.request({
  hostname: 'facilitycare-system.pages.dev',
  port: 443,
  path: '/api/audit-clean-up-reports',
  method: 'POST'
}, (res) => {
  let data = '';
  res.on('data', d => data += d);
  res.on('end', () => {
    console.log(`Status: ${res.statusCode}`);
    console.log(`Response: ${data}`);
  });
});

req.on('error', (e) => {
  console.error(`Error: ${e.message}`);
});

req.end();
