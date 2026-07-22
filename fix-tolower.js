const fs = require('fs');
const path = require('path');

function walkDir(dir) {
  let results = [];
  const list = fs.readdirSync(dir);
  list.forEach(file => {
    file = path.join(dir, file);
    const stat = fs.statSync(file);
    if (stat && stat.isDirectory()) {
      results = results.concat(walkDir(file));
    } else if (file.endsWith('.js') || file.endsWith('.jsx')) {
      results.push(file);
    }
  });
  return results;
}

const files = walkDir('./frontend/src');
let bad = 0;
for (const file of files) {
  let c = fs.readFileSync(file, 'utf8');
  let o = c;

  c = c.replace(/const s = str\.toLowerCase\(\);/g, "const s = String(str || '').toLowerCase();");
  c = c.replace(/r\.full_name\.toLowerCase\(\)/g, "String(r.full_name || '').toLowerCase()");
  c = c.replace(/r\.code\.toLowerCase\(\)/g, "String(r.code || '').toLowerCase()");
  c = c.replace(/r\.name\.toLowerCase\(\)/g, "String(r.name || '').toLowerCase()");
  c = c.replace(/itemLabel\.toLowerCase\(\)/g, "String(itemLabel || '').toLowerCase()");

  if (c !== o) {
    fs.writeFileSync(file, c);
    console.log('Fixed', file);
    bad++;
  }
}
console.log('Total fixed files:', bad);
