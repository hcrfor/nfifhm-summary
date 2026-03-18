import XLSX from 'xlsx';
import fs from 'fs';

const filename = 'NFI_110964619515_입력중.xlsm';
const wb = XLSX.readFile(filename);

let output = '';

const standSheetName = wb.SheetNames.find(n => n.includes('임분조사'));
if (standSheetName) {
    output += `\n--- Sheet: ${standSheetName} ---\n`;
    const ws = wb.Sheets[standSheetName];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    
    output += 'Top 50 rows:\n';
    data.slice(0, 50).forEach((row, i) => {
        output += `Row ${i}: ${row.join(' | ')}\n`;
    });
}

fs.writeFileSync('debug_stand_50.txt', output, 'utf8');
