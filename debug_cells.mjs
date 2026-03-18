import XLSX from 'xlsx';

const filename = 'NFI_110964619515_입력중.xlsm';
const wb = XLSX.readFile(filename);

const name = '임분조사표';
const ws = wb.Sheets[name];
console.log(`Sheet: ${name}`);
console.log(`!ref: ${ws['!ref']}`);

// Check specific cells that should have data based on screenshot
// Row 2 Col A should be '1109646195151'
// row index is 1, col index is 0 => cell A2
console.log(`A1: ${ws['A1'] ? ws['A1'].v : 'empty'}`);
console.log(`A2: ${ws['A2'] ? ws['A2'].v : 'empty'}`);
console.log(`A3: ${ws['A3'] ? ws['A3'].v : 'empty'}`);
console.log(`A4: ${ws['A4'] ? ws['A4'].v : 'empty'}`);
console.log(`A5: ${ws['A5'] ? ws['A5'].v : 'empty'}`);

// Also check row 10 to see if it's there
console.log(`A10: ${ws['A10'] ? ws['A10'].v : 'empty'}`);
