const fs = require('fs');
let c = fs.readFileSync('src/pages/dashboard.js', 'utf8');
const searchString = `key:'employees',  trendPct:'+2%', trendColor:'#10B981', points:'0,20 10,18 20,22 30,12 40,15 50,8 60,10 70,5 80,6 90,2 100,0' },`;
const replaceString = searchString + `\n      { icon:'🔄', label:'Reliefer Aktif',        sub:'Karyawan reliefer',       href:'#/employees?dash_filter=reliefer',   color:'kpi-purple',   key:'reliever_total',  trendPct:'0%', trendColor:'#10B981', points:'0,15 20,18 40,10 60,12 80,5 100,2' },`;
c = c.replace(searchString, replaceString);
fs.writeFileSync('src/pages/dashboard.js', c);
console.log("Done");
