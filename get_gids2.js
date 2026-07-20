const fs = require('fs');
const html = fs.readFileSync('htmlview_full.html', 'utf8');

// Find things that look like ["SheetName",gid] or similar in the JS payload
const regex = /\["([^"]+)",(\d{4,12})\]/g;
let match;
let count = 0;
while ((match = regex.exec(html)) !== null) {
  console.log(`Sheet: ${match[1]} - GID: ${match[2]}`);
  count++;
}
if(count===0){
  console.log("Still no match. Let's dump the first 2000 chars of JS payload.");
  const jsMatch = html.match(/<script.*?>(.*?)<\/script>/gs);
  if(jsMatch) {
     jsMatch.forEach(m => {
       if (m.includes('gid') || m.includes('grid')) {
         console.log(m.substring(0, 500));
       }
     });
  }
}
