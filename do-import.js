const fs = require('fs');
const https = require('https');

const parseDate = (v) => {
  if (!v) return '';
  const s = String(v).trim();
  if (s === '' || s === '0') return '';
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  if (/^\d{4,5}$/.test(s)) {
    const n = Number(s);
    if (n > 20000 && n < 99999) {
      const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
      return isNaN(d.getTime()) ? '' : d.toISOString().slice(0, 10);
    }
  }
  const parts = s.split(/[\/\-\.]/);
  if (parts.length === 3) {
    if (parts[2].length === 4) {
      const dd = parts[0].padStart(2, '0');
      const mm = parts[1].padStart(2, '0');
      return `${parts[2]}-${mm}-${dd}`;
    }
  }
  return '';
};

const csv = fs.readFileSync('schedule.csv', 'utf8');
const lines = csv.split('\n');

const payload = [];

for (let i = 1; i < lines.length; i++) {
  const line = lines[i].trim();
  if (!line) continue;
  
  const cols = line.split(',');
  
  payload.push({
    branch_name: cols[0],
    activity_type: cols[1],
    period: cols[2],
    pic: cols[3],
    opening_date: parseDate(cols[4]),
    target_date: parseDate(cols[5]),
    completion_date: parseDate(cols[6]),
    status: cols[7],
    notes: cols[8]
  });
}

const req = https.request({
  hostname: 'facilitycare-system.pages.dev',
  port: 443,
  path: '/api/audit-import-schedule',
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
    'Content-Length': Buffer.byteLength(JSON.stringify(payload))
  }
}, res => {
  let data = '';
  res.on('data', d => data += d);
  res.on('end', () => console.log(data));
});

req.on('error', e => console.error(e));
req.write(JSON.stringify(payload));
req.end();
