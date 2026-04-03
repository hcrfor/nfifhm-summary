import XLSX from 'xlsx';
const wb = XLSX.readFile('2021자료.xlsx');
const ws = wb.Sheets[wb.SheetNames.find(n => n.includes('2021')) || wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(ws);

const found = data.filter(r => String(r['표본점번호'] || '').includes('43103151895'));

console.log(`Found: ${found.length}`);
found.forEach(f => {
    console.log(`P: ${f['표본점번호']}, OP: ${f['구표본점번호']}, C: ${f['집락번호']}`);
});
