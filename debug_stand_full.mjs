import XLSX from 'xlsx';
import fs from 'fs';

const filename = 'NFI_110964619515_입력중.xlsm';
const wb = XLSX.readFile(filename);

const standSheetName = wb.SheetNames.find(n => n.includes('임분조사'));
const ws = wb.Sheets[standSheetName];
const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

let output = `Sheet: ${standSheetName}\nTotal Rows: ${data.length}\n\n`;
data.forEach((row, i) => {
    output += `Row ${i} (len ${row.length}): ${row.join(' | ')}\n`;
});

fs.writeFileSync('debug_stand_full.txt', output, 'utf8');
