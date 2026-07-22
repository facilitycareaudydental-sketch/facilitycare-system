const fs = require('fs');
const path = require('path');
const dir = 'src/pages';
const files = fs.readdirSync(dir).filter(f => f.endsWith('.js'));

files.forEach(f => {
  let p = path.join(dir, f);
  let c = fs.readFileSync(p, 'utf8');
  let o = c;

  c = c.replace(
    /{ name: 'branch_id',([^}]*)type: 'select',([^}]*)options: branchOptions, value: data\?\.branch_id }/g,
    "{ name: 'branch_id',$1type: 'combobox',$2options: branchOptions, createApi: { path: '/api/branches', field: 'full_name' }, value: data?.branch_id }"
  );

  c = c.replace(
    /{ type: 'select', name: 'from_branch_id',([^}]*)options: branchOptions }/g,
    "{ type: 'combobox', name: 'from_branch_id',$1options: branchOptions, createApi: { path: '/api/branches', field: 'full_name' } }"
  );

  c = c.replace(
    /{ type: 'select', name: 'to_branch_id',([^}]*)options: branchOptions }/g,
    "{ type: 'combobox', name: 'to_branch_id',$1options: branchOptions, createApi: { path: '/api/branches', field: 'full_name' } }"
  );

  c = c.replace(
    /{ name: 'employee_name',([^}]*)type: 'select',([^}]*)options: ([a-zA-Z_()?\.]+), value: data\?\.employee_name }/g,
    "{ name: 'employee_name',$1type: 'combobox',$2options: $3, createApi: { path: '/api/employees', field: 'full_name', extra: { status: 'Aktif' } }, value: data?.employee_name }"
  );

  c = c.replace(
    /{ name: 'pic',([^}]*)type: 'select',([^}]*)options: ([a-zA-Z_()?\.]+), value: data\?\.pic }/g,
    "{ name: 'pic',$1type: 'combobox',$2options: $3, createApi: { path: '/api/pic', field: 'name' }, value: data?.pic }"
  );

  c = c.replace(
    /{ name: 'trainer',([^}]*)type: 'select',([^}]*)options: ([a-zA-Z_()?\.]+), value: data\?\.trainer }/g,
    "{ name: 'trainer',$1type: 'combobox',$2options: $3, createApi: { path: '/api/pic', field: 'name' }, value: data?.trainer }"
  );

  if (o !== c) {
    fs.writeFileSync(p, c);
    console.log('Updated ' + f);
  }
});
