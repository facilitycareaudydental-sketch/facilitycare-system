const fs = require('fs');

const csv = fs.readFileSync('schedule.csv', 'utf8');
const lines = csv.split('\n');

const counts = {};
const duplicates = [];

for (let i = 1; i < lines.length; i++) {
  const line = lines[i].trim();
  if (!line) continue;
  
  const cols = line.split(',');
  const branch = cols[0];
  const activity = cols[1];
  const period = cols[2];
  
  // parse target date to get year
  let year = 'noyear';
  const targetDateStr = cols[5] || '';
  // date format could be dd/mm/yyyy or yyyy-mm-dd or Excel serial
  if (targetDateStr) {
    if (targetDateStr.includes('-')) {
       // simple check for yyyy-mm-dd or dd-mm-yyyy
       const parts = targetDateStr.split('-');
       if (parts[0].length === 4) year = parts[0];
       else if (parts[2] && parts[2].length === 4) year = parts[2];
    } else if (targetDateStr.includes('/')) {
       const parts = targetDateStr.split('/');
       if (parts[2] && parts[2].length === 4) year = parts[2];
    }
  }
  
  const key = `${branch}_${activity}_${period}_${year}`;
  
  if (counts[key]) {
    duplicates.push(`Baris ${i + 1}: ${key} (Asli di Baris ${counts[key]})`);
  } else {
    counts[key] = i + 1;
  }
}

console.log(`Total baris (termasuk header): ${lines.length}`);
console.log(`Total data valid: ${lines.length - 1}`);
console.log(`Total duplikat: ${duplicates.length}`);
if (duplicates.length > 0) {
  console.log('Duplikat ditemukan pada:');
  console.log(duplicates.join('\n'));
} else {
  console.log('TIDAK ADA DUPLIKAT!');
}
