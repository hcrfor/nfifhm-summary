import XLSX from 'xlsx';
const wb = XLSX.readFile('2021자료.xlsx');
const ws = wb.Sheets[wb.SheetNames.find(n => n.includes('2021')) || wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(ws);

const cluster = '431031518958';
const points = data.filter(r => String(r['집락번호'] || '').trim() === cluster);

console.log(`Cluster: ${cluster}`);
points.forEach(p => {
    console.log(`P: ${p['표본점번호']}, OP: ${p['구표본점번호']}`);
});
