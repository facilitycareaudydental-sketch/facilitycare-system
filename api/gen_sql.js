const fs = require('fs'); 
const lines = fs.readFileSync('data.txt', 'utf8').trim().split('\n'); 
let sql = ''; 
lines.forEach(line => { 
  const parts = line.split('\t'); 
  if(parts.length < 5) return; 
  const [branch, period, dateStr, status, fc, spv] = parts; 
  const branch_name = branch.trim(); 
  const p = period.trim(); 
  const [d,m,y] = dateStr.trim().split('/'); 
  const date = y+'-'+m.padStart(2,'0')+'-'+d.padStart(2,'0'); 
  const f = fc && fc.trim() ? fc.trim() : 'NULL'; 
  const s = spv && spv.trim() ? spv.trim() : 'NULL'; 
  sql += `UPDATE inspection_reports SET fc_score = ${f}, spv_score = ${s} WHERE period = '${p}' AND inspection_date = '${date}' AND branch_id = (SELECT id FROM branches WHERE full_name = '${branch_name}' OR name = '${branch_name}' LIMIT 1);\n`; 
}); 
fs.writeFileSync('update_scores.sql', sql);
