const fs = require('fs');
const https = require('https');

const raw = fs.readFileSync('./raw-data.txt', 'utf8');
const lines = raw.split('\n').filter(x => x.trim().length > 0);
const payload = [];

for (const line of lines) {
  const parts = line.split('\t');
  if (parts.length >= 2) {
     payload.push({ emp: parts[0].trim(), branch: parts[1].trim() });
  }
}

console.log('Sending payload size:', payload.length);

const data = JSON.stringify(payload);
const options = {
  hostname: 'facilitycare-system.pages.dev',
  port: 443,
  path: '/api/audit-force-branches',
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'Content-Length': data.length
  }
};

const req = https.request(options, (res) => {
  let resData = '';
  res.on('data', d => resData += d);
  res.on('end', () => {
    console.log('Response:', res.statusCode, resData);
  });
});

req.on('error', (e) => {
  console.error(e);
});
req.write(data);
req.end();
