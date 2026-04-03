import XLSX from 'xlsx';

const filename = 'public/2021자료.xlsx';
try {
    const wb = XLSX.readFile(filename);
    wb.SheetNames.forEach(sheetName => {
        console.log(`Sheet: ${sheetName}`);
        const ws = wb.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
        if (data.length > 0) {
            console.log('Columns:', data[0]);
        }
    });
} catch (e) {
    console.error('Error:', e);
}
