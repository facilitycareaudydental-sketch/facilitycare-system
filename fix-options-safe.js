const fs = require('fs');
const files = fs.readdirSync('./frontend/src/pages').filter(f => f.endsWith('.js'));
let fixedFiles = 0;

for (const file of files) {
  const p = './frontend/src/pages/' + file;
  let c = fs.readFileSync(p, 'utf8');
  let original = c;

  let lines = c.split('\n');
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].includes('options: branchOptions') && lines[i].includes('value: data?.branch_id')) {
      lines[i] = lines[i].replace('options: branchOptions', "options: (data?.branch_id && !branchOptions.find(o => o.value == data.branch_id)) ? [...branchOptions, { value: data.branch_id, label: data.branch_name || data.branch_id }] : branchOptions");
    }
    if (lines[i].includes('options: employeeOptions') && lines[i].includes('value: data?.employee_id')) {
      lines[i] = lines[i].replace('options: employeeOptions', "options: (data?.employee_id && !employeeOptions.find(o => o.value == data.employee_id)) ? [...employeeOptions, { value: data.employee_id, label: data.employee_name || data.employee_id }] : employeeOptions");
    }
    if (lines[i].includes('options: picOptions') && lines[i].includes('value: data?.pic_id')) {
      lines[i] = lines[i].replace('options: picOptions', "options: (data?.pic_id && !picOptions.find(o => o.value == data.pic_id)) ? [...picOptions, { value: data.pic_id, label: data.pic_name || data.pic_id }] : picOptions");
    }
  }

  c = lines.join('\n');
  if (c !== original) {
    fs.writeFileSync(p, c, 'utf8');
    fixedFiles++;
    console.log('Fixed:', file);
  }
}

console.log('Total fixed:', fixedFiles);
