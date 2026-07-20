const Papa = require('papaparse');
const https = require('https');

const SPREADSHEET_ID = '1kdORjAnJ4UB-j_eDMHyiK2TLBgfTIiBMnJOXJcaep0o';

function fetchCSV(sheetName) {
  return new Promise((resolve, reject) => {
    const url = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`;
    https.get(url, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve(data));
    }).on('error', reject);
  });
}

fetchCSV('Validasi').then(d => {
  console.log("Validasi RAW size:", d.length);
  console.log("Lines:", d.split('\\n').length);
  const parsed = Papa.parse(d, {header: true});
  console.log("Parsed rows:", parsed.data.length);
  console.log("First row:", parsed.data[0]);
});
