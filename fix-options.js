const fs = require('fs');
const files = fs.readdirSync('./frontend/src/pages').filter(f => f.endsWith('.js'));
for (const file of files) {
  const p = './frontend/src/pages/' + file;
  let c = fs.readFileSync(p, 'utf8');
  let o = c;

  // Replace options: branchOptions -> options: data?.branch_id && !branchOptions.find(o => o.value == data.branch_id) ? [...branchOptions, {value: data.branch_id, label: data.branch_name || data.branch_id}] : branchOptions
  // But wait, what if it's in the filterFields? Filter fields don't have data.
  // So we only replace it if it's on a line that also has alue: data?.branch_id
  
  c = c.replace(/options:\s*branchOptions(.*?value:\s*data\?\.branch_id)/g, 
    "options: (data?.branch_id && !branchOptions.find(o => o.value == data.branch_id)) ? [...branchOptions, { value: data.branch_id, label: data.branch_name || data.branch_id }] : branchOptions");

  c = c.replace(/options:\s*employeeOptions(.*?value:\s*data\?\.employee_id)/g, 
    "options: (data?.employee_id && !employeeOptions.find(o => o.value == data.employee_id)) ? [...employeeOptions, { value: data.employee_id, label: data.employee_name || data.employee_id }] : employeeOptions");

  c = c.replace(/options:\s*picOptions(.*?value:\s*data\?\.pic_id)/g, 
    "options: (data?.pic_id && !picOptions.find(o => o.value == data.pic_id)) ? [...picOptions, { value: data.pic_id, label: data.pic_name || data.pic_id }] : picOptions");

  if (c !== o) {
    fs.writeFileSync(p, c);
    console.log('Fixed', file);
  }
}
