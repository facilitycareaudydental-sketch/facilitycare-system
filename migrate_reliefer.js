const fs = require('fs');
const Papa = require('papaparse');
const https = require('https');

async function migrateReliefer() {
  const fileData = fs.readFileSync('reliefer.csv', 'utf8');
  const parsed = Papa.parse(fileData, { header: true, skipEmptyLines: true });
  
  const branchesReq = await new Promise(resolve => {
      https.get('https://fm-operations-api.facilitycare-audydental.workers.dev/api/branches?all=1', res => {
          let bdata = '';
          res.on('data', c => bdata += c);
          res.on('end', () => resolve(JSON.parse(bdata)));
      });
  });
  
  const branches = branchesReq.data || [];
  const matchBranch = (str) => {
    if (!str) return null;
    const s = str.toLowerCase();
    const b = branches.find(r => r.full_name.toLowerCase() === s || r.code.toLowerCase() === s || r.name.toLowerCase() === s);
    return b ? b.id : null;
  };

  const payload = parsed.data.map(row => ({
    branch_id: matchBranch(row['Cabang']),
    backup_date: (row['Tanggal Back Up'] || '').trim(),
    original_fc_name: (row['Nama Facility care'] || '').trim(),
    reliever_name: (row['Relifer'] || '').trim(),
    period: (row['Periode'] || '').trim(),
    reason: (row['Keterangan'] || '').trim(),
    shift: (row['Shift'] || '').trim(),
    completion_date: (row['Tanggal Selesai'] || '').trim(),
    status: (row['Status'] || '').trim() || 'Pending'
  })).filter(row => row.reliever_name && row.backup_date);

  console.log(`Found ${payload.length} rows to migrate to Relievers...`);
  if (payload.length === 0) return;

  const options = {
    hostname: 'fm-operations-api.facilitycare-audydental.workers.dev',
    port: 443,
    path: '/api/relievers/import',
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    }
  };
  
  const req = https.request(options, res => {
    let rdata = '';
    res.on('data', d => rdata += d);
    res.on('end', () => console.log('Import Reliever Result:', rdata));
  });
  req.write(JSON.stringify(payload));
  req.end();
}

migrateReliefer();
