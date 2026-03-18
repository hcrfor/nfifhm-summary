import XLSX from 'xlsx';

const filename = 'NFI_110964619515_입력중.xlsm';
const wb = XLSX.readFile(filename);

const name = '일반정보';
const ws = wb.Sheets[name];
console.log(`Sheet: ${name}`);
console.log(`!ref: ${ws['!ref']}`);

for(let i=1; i<=10; i++) {
    console.log(`A${i}: ${ws['A'+i] ? ws['A'+i].v : 'empty'}`);
}
