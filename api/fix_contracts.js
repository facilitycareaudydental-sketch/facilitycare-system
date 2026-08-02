const names = [
  'AHMAD SARIP HIDAYATULOH',
  'CHANDRA AJI KURNIAWAN',
  'HENDRA DWI SAPUTRA',
  'IQBAL AL BANNA',
  'MUHAMAD FATHAN MUBIN',
  'MUHAMAD HUSNI',
  'MUHAMAD KEMAL RIZQULLAH ALGHIFARI',
  'MUHAMMAD TRI ISMANDANU',
  'TANTO SUTRISNO'
];

const branches = {
  'AHMAD SARIP HIDAYATULOH': '004. Cipete',
  'CHANDRA AJI KURNIAWAN': '016. Cibubur',
  'HENDRA DWI SAPUTRA': '014. Sunter',
  'IQBAL AL BANNA': '000. Management',
  'MUHAMAD FATHAN MUBIN': '067. Agora Mall',
  'MUHAMAD HUSNI': '015. Gading Serpong',
  'MUHAMAD KEMAL RIZQULLAH ALGHIFARI': '036. Jagakarsa',
  'MUHAMMAD TRI ISMANDANU': '000. Management',
  'TANTO SUTRISNO': '065. Kebon Sirih'
};

const divisions = {
  'AHMAD SARIP HIDAYATULOH': 'SECURITY',
  'CHANDRA AJI KURNIAWAN': 'SECURITY',
  'HENDRA DWI SAPUTRA': 'SECURITY',
  'IQBAL AL BANNA': 'FACILITY CARE',
  'MUHAMAD FATHAN MUBIN': 'FACILITY CARE',
  'MUHAMAD HUSNI': 'FACILITY CARE',
  'MUHAMAD KEMAL RIZQULLAH ALGHIFARI': 'FACILITY CARE',
  'MUHAMMAD TRI ISMANDANU': 'FACILITY CARE',
  'TANTO SUTRISNO': 'FACILITY CARE'
};

const start_dates = {
  'AHMAD SARIP HIDAYATULOH': '2026-08-07',
  'IQBAL AL BANNA': '2026-04-07'
};

const { execSync } = require('child_process');

async function fix() {
  console.log('Fetching employees...');
  const eStr = execSync('npx wrangler d1 execute fcms-production-db --remote --json --command "SELECT id, full_name FROM employees"').toString();
  const emps = JSON.parse(eStr)[0].results;
  
  const bStr = execSync('npx wrangler d1 execute fcms-production-db --remote --json --command "SELECT id, name, code, full_name FROM branches"').toString();
  const branchList = JSON.parse(bStr)[0].results;
  
  const matchBranch = (str) => {
      const s = String(str || '').replace(/\s+/g, ' ').toLowerCase().trim();
      const b = branchList.find(r => String(r.full_name || '').replace(/\s+/g, ' ').toLowerCase().trim() === s || String(r.code || '').replace(/\s+/g, ' ').toLowerCase().trim() === s || String(r.name || '').replace(/\s+/g, ' ').toLowerCase().trim() === s);
      return b ? b.id : null;
  };
  
  let queries = [];
  
  for (const name of names) {
    const e = emps.find(x => x.full_name.toLowerCase() === name.toLowerCase());
    if (e) {
      console.log('Found employee:', e.full_name, e.id);
      
      const bName = branches[name];
      const bId = matchBranch(bName);
      const div = divisions[name];
      const start = start_dates[name] || '';
      const end = '2099-12-31';
      
      queries.push(`INSERT INTO contracts (employee_id, branch_id, division, start_date, end_date, status) VALUES ('${e.id}', ${bId ? "'" + bId + "'" : 'NULL'}, '${div}', '${start}', '${end}', 'Aktif');`);
    } else {
      console.log('NOT FOUND:', name);
    }
  }
  
  fs.writeFileSync('insert_missing.sql', queries.join('\n'));
  console.log('Saved to insert_missing.sql');
}

const fs = require('fs');
fix();
