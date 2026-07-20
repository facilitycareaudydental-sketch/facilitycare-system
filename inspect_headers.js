const xlsx = require('xlsx');

const workbook = xlsx.readFile('C:\\Users\\Facility Care\\.gemini\\antigravity\\scratch\\facilitycare-system\\New FM Operations.xlsx');

const targetSheets = ['One on One', 'Jadwal Reliefer', 'Master Checklist', 'Permintaan Chemical'];

for (const sheetName of targetSheets) {
    if (workbook.SheetNames.includes(sheetName)) {
        const sheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
        if (data.length > 0) {
            console.log(`\n--- ${sheetName} ---`);
            console.log(JSON.stringify(data[0], null, 2));
        } else {
            console.log(`\n--- ${sheetName} (EMPTY) ---`);
        }
    } else {
        console.log(`\n--- ${sheetName} (NOT FOUND) ---`);
    }
}
