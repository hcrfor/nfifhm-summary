import XLSX from 'xlsx';
const wb = XLSX.readFile('2021자료.xlsx');
const ws = wb.Sheets[wb.SheetNames.find(n => n.includes('2021')) || wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(ws);

const p2 = '4310315189582';
const p3 = '4310315189583';

const match2 = data.find(r => String(r['표본점번호'] || '').trim() === p2);
const match3 = data.find(r => String(r['표본점번호'] || '').trim() === p3);

console.log(`Point 2: ${match2 ? match2['구표본점번호'] : 'NOT FOUND'}`);
console.log(`Point 3: ${match3 ? match3['구표본점번호'] : 'NOT FOUND'}`);
