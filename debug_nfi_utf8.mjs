import XLSX from 'xlsx';
import fs from 'fs';

const filename = 'NFI_110964619515_입력중.xlsm';
const wb = XLSX.readFile(filename);

let output = '';
output += 'Sheet Names: ' + JSON.stringify(wb.SheetNames) + '\n';

const standSheetName = wb.SheetNames.find(n => n.includes('임분조사'));
if (standSheetName) {
    output += `\n--- Sheet: ${standSheetName} ---\n`;
    const ws = wb.Sheets[standSheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    
    output += 'Sample rows (top 15):\n';
    data.slice(0, 15).forEach((row, i) => {
        output += `Row ${i}: ${row.join(' | ')}\n`;
    });
}

const generalSheetName = wb.SheetNames.find(n => n.includes('일반정보'));
if (generalSheetName) {
    output += `\n--- Sheet: ${generalSheetName} ---\n`;
    const ws = wb.Sheets[generalSheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    
    output += 'Sample rows (top 5):\n';
    data.slice(0, 5).forEach((row, i) => {
        output += `Row ${i}: ${row.join(' | ')}\n`;
    });
}

fs.writeFileSync('debug_utf8.txt', output, 'utf8');
console.log('Done writing debug_utf8.txt');
