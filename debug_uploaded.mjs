import XLSX from 'xlsx';
const filename = 'NFI_431031518958.xlsx';
const wb = XLSX.readFile(filename);

wb.SheetNames.forEach(name => {
    console.log(`\n--- Sheet: ${name} ---`);
    const ws = wb.Sheets[name];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
    data.slice(0, 50).forEach(row => {
        const rowStr = row.join(' | ');
        if (rowStr.includes('27603960') || rowStr.includes('431031518958')) {
            console.log(rowStr);
        }
    });
});
