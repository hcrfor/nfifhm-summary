import XLSX from 'xlsx';

const filename = 'NFI_110964619515_입력중.xlsm';
const wb = XLSX.readFile(filename);

console.log('Sheet Names:', wb.SheetNames);

const standSheetName = wb.SheetNames.find(n => n.includes('임분조사'));
if (standSheetName) {
    console.log(`\n--- Sheet: ${standSheetName} ---`);
    const ws = wb.Sheets[standSheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    
    console.log('Sample rows (top 15):');
    data.slice(0, 15).forEach((row, i) => {
        console.log(`Row ${i}:`, row.join(' | '));
    });
} else {
    console.log('임분조사 sheet not found');
}

const generalSheetName = wb.SheetNames.find(n => n.includes('일반정보'));
if (generalSheetName) {
    console.log(`\n--- Sheet: ${generalSheetName} ---`);
    const ws = wb.Sheets[generalSheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    
    console.log('Sample rows (top 5):');
    data.slice(0, 5).forEach((row, i) => {
        console.log(`Row ${i}:`, row.join(' | '));
    });
}
