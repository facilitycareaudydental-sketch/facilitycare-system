const Papa = require('papaparse');
const https = require('https');

const SPREADSHEET_ID = '1kdORjAnJ4UB-j_eDMHyiK2TLBgfTIiBMnJOXJcaep0o';

function fetchCSV(sheetName) {
  return new Promise((resolve, reject) => {
    const url = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
    https.get(url, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve(data));
    }).on('error', reject);
  });
}

async function migrateValidasi() {
  try {
    const data = await fetchCSV('Validasi');
    const parsed = Papa.parse(data, { header: true, skipEmptyLines: true });
    
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
      branch_id: matchBranch((row['CABANG'] || '').trim()),
      activity_type: (row['KEGIATAN'] || '').trim(),
      period: (row['QUARTAL'] || '').trim(),
      pic: (row['PIC'] || '').trim(),
      status: 'Pending'
    })).filter(r => r.activity_type && r.period);

    console.log(`Found ${payload.length} rows to migrate to Schedule...`);
    
    const options = {
      hostname: 'fm-operations-api.facilitycare-audydental.workers.dev',
      port: 443,
      path: '/api/schedule/import',
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      }
    };
    
    const req = https.request(options, res => {
      let rdata = '';
      res.on('data', d => rdata += d);
      res.on('end', () => console.log('Import Schedule Result:', rdata));
    });
    req.write(JSON.stringify(payload));
    req.end();
  } catch(e) {
    console.error(e);
  }
}

migrateValidasi();
