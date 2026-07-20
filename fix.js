const fs = require('fs');
const files = {
  'checklist.js': '/api/checklist',
  'contracts.js': '/api/contracts',
  'forms.js': '/api/forms',
  'issues.js': '/api/issues',
  'relievers.js': '/api/relievers',
  'schedule.js': '/api/schedule',
  'sop.js': '/api/sop',
  'training.js': '/api/training',
  'users.js': '/api/users',
  'one_on_one.js': '/api/one-on-one'
};

for (const [f, p] of Object.entries(files)) {
  const path = 'frontend/src/pages/' + f;
  if (fs.existsSync(path)) {
    let c = fs.readFileSync(path, 'utf8');
    c = c.replace(/apiPath:,\s*bulkDelete: true,/g, "apiPath: '" + p + "',\n    bulkDelete: true,");
    fs.writeFileSync(path, c);
    console.log('Fixed ' + f);
  }
}
