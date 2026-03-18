import XLSX from 'xlsx';
import fs from 'fs';

const filename = 'NFI_110964619515_입력중.xlsm';
const wb = XLSX.readFile(filename);

let output = '';
wb.SheetNames.forEach(name => {
    const ws = wb.Sheets[name];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    output += `Sheet: ${name}, Rows: ${data.length}\n`;
    if (data.length > 0) {
        output += `  Row 0: ${data[0].slice(0, 5).join(' | ')} ...\n`;
    }
    if (data.length > 1) {
        output += `  Row 1: ${data[1].slice(0, 5).join(' | ')} ...\n`;
    }
});

fs.writeFileSync('debug_all_sheets.txt', output, 'utf8');
